#requires -version 5.1
<#
.SYNOPSIS
    Show-DiskUsage.ps1 - Visualizador de uso de disco estilo TreeSize em PowerShell + WinForms.
    Solucao propria, sem instalar software de terceiros (compliance-safe).

.DESCRIPTION
    Pergunta ao proprio Windows o tamanho FISICO alocado de cada arquivo (GetCompressedFileSizeW).
    Com isso o numero bate com a coluna "Allocated" do TreeSize/WizTree:
      - compressao NTFS, arquivos sparse e hardlinks sao tratados pelo SO (sem chute);
      - arquivos so-na-nuvem (OneDrive Files On-Demand, desidratados) alocam ~0 e saem da conta;
      - junctions/symlinks (reparse points de diretorio) nao sao percorridos (evita dupla contagem).

    - Maior sempre no topo, com percentual relativo ao pai
    - Mostra pastas E arquivos; arvore navegavel (expansao instantanea via cache)
    - Menu de contexto (botao direito): Abrir no Explorer / Copiar caminho / Deletar (com confirmacao)

.NOTES
    Requer Administrador e PowerShell 5.1+. A varredura percorre o filesystem (mais lenta que ler a
    MFT crua: 1-3 min para C: inteiro), em troca de um numero fisico confiavel. A GUI fica ocupada
    durante a varredura; o status mostra qual pasta de topo esta sendo processada.

.EXAMPLE
    PowerShell -ExecutionPolicy Bypass -File .\tools\Show-DiskUsage.ps1
#>

# ---- Requer Administrador (rode num PowerShell aberto como administrador) ----
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
$pr = New-Object Security.Principal.WindowsPrincipal($id)
if (-not $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host ''
    Write-Host '  ERRO: este script precisa de Administrador.' -ForegroundColor Red
    Write-Host '  Abra o PowerShell como administrador (botao direito -> Executar como administrador)' -ForegroundColor Yellow
    Write-Host '  e rode de novo:  powershell -ExecutionPolicy Bypass -File .\tools\Show-DiskUsage.ps1' -ForegroundColor Yellow
    Write-Host ''
    Read-Host '  Pressione Enter para sair'
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Blindagem: qualquer erro fatal vira MessageBox + log (senao a janela fecha sem mostrar nada)
try {

# ============================================================================
#  Motor: varre o filesystem e pega o tamanho FISICO alocado de cada arquivo
#  diretamente do Windows (GetCompressedFileSizeW). Sem parsear NTFS na mao.
# ============================================================================
$cs = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace CascoDigital {
  public class DiskWalker {
    [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern uint GetCompressedFileSizeW(string path, out uint high);

    const uint INVALID_SIZE = 0xFFFFFFFF;
    const long SANITY_CAP   = 0x4000000000000L; // 1 PB: tamanho maior que isso num arquivo = lixo

    public Dictionary<string, long> DirSize =
        new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
    public long FileCount;

    // Prefixo \\?\ para vencer o limite de MAX_PATH (260) na chamada de tamanho
    static string Long(string p) {
      if (p.StartsWith("\\\\?\\")) return p;
      if (p.StartsWith("\\\\"))    return "\\\\?\\UNC\\" + p.Substring(2);
      return "\\\\?\\" + p;
    }

    // Normaliza a chave do cache. "C:\" e "C:" viram "C:"; demais perdem a barra final.
    static string Key(string p) {
      p = p.TrimEnd('\\');
      return p;
    }

    // Tamanho FISICO alocado de um arquivo, direto do Windows. 0 se falhar ou vier lixo.
    public long FileAlloc(string path) {
      uint high;
      uint low = GetCompressedFileSizeW(Long(path), out high);
      if (low == INVALID_SIZE && Marshal.GetLastWin32Error() != 0) return 0;
      long v = ((long)high << 32) | (long)low;
      if (v < 0 || v > SANITY_CAP) return 0;
      return v;
    }

    public long DirAlloc(string path) {
      long v;
      return DirSize.TryGetValue(Key(path), out v) ? v : 0;
    }

    // Varre uma subarvore ITERATIVAMENTE (pilha explicita -> sem StackOverflow), usando a
    // enumeracao nativa do .NET (DirectoryInfo) para percorrer e GetCompressedFileSize para medir.
    public long ScanTree(string root) {
      string r = Key(root);
      Dictionary<string, long>   direct = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
      Dictionary<string, string> parent = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
      List<string> order = new List<string>();
      Stack<string> stack = new Stack<string>();
      direct[r] = 0; order.Add(r); stack.Push(r);

      while (stack.Count > 0) {
        string dir = stack.Pop();
        long files = 0;
        try {
          // "C:" precisa virar "C:\" para o DirectoryInfo apontar para a raiz, nao o dir atual
          string dpath = (dir.Length == 2 && dir[1] == ':') ? dir + "\\" : dir;
          DirectoryInfo di = new DirectoryInfo(dpath);
          foreach (FileSystemInfo fsi in di.EnumerateFileSystemInfos()) {
            FileAttributes a;
            try { a = fsi.Attributes; } catch { continue; }
            bool isDir = (a & FileAttributes.Directory)     != 0;
            bool isRep = (a & FileAttributes.ReparsePoint)  != 0;
            if (isDir) {
              if (isRep) continue;                 // junction/symlink: nao percorre
              string full = Key(fsi.FullName);
              if (!direct.ContainsKey(full)) {     // guarda contra qualquer ciclo
                direct[full] = 0; parent[full] = dir; order.Add(full); stack.Push(full);
              }
            } else {
              files += FileAlloc(fsi.FullName);
              FileCount++;
            }
          }
        } catch { }                                // acesso negado / caminho longo: ignora a pasta
        direct[dir] = files;                       // bytes dos arquivos diretos desta pasta
      }

      // Rollup do mais fundo pro mais raso: cada pasta soma seu total ao pai.
      for (int i = 0; i < order.Count; i++) DirSize[order[i]] = direct[order[i]];
      for (int i = order.Count - 1; i >= 0; i--) {
        string d = order[i], p;
        if (parent.TryGetValue(d, out p)) DirSize[p] += DirSize[d];
      }
      long t; return DirSize.TryGetValue(r, out t) ? t : 0;
    }
  }
}
'@
Add-Type -TypeDefinition $cs -Language CSharp -ErrorAction Stop

# ============================================================================
#  Estado e helpers
# ============================================================================
$script:Walker = $null
$MAX_CHILDREN  = 5000

function Format-Size([long]$bytes) {
    if ($bytes -ge 1TB) { return ('{0:N2} TB' -f ($bytes / 1TB)) }
    if ($bytes -ge 1GB) { return ('{0:N2} GB' -f ($bytes / 1GB)) }
    if ($bytes -ge 1MB) { return ('{0:N2} MB' -f ($bytes / 1MB)) }
    if ($bytes -ge 1KB) { return ('{0:N2} KB' -f ($bytes / 1KB)) }
    return "$bytes B"
}

function Test-Reparse($item) {
    return (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
}

# Popula os filhos de um node (uma camada). Pastas usam o cache; arquivos consultam o alocado na hora.
function Build-Children($node) {
    $node.Nodes.Clear()
    $path = $node.Name
    $parentSize = $script:Walker.DirAlloc($path)
    $items = @()
    try {
        foreach ($d in (Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction SilentlyContinue)) {
            if (Test-Reparse $d) { continue }   # junction/symlink: nao contamos (igual a varredura)
            $items += [pscustomobject]@{ Name=$d.Name; Path=$d.FullName; Size=($script:Walker.DirAlloc($d.FullName)); IsDir=$true }
        }
        foreach ($f in (Get-ChildItem -LiteralPath $path -File -Force -ErrorAction SilentlyContinue)) {
            $items += [pscustomobject]@{ Name=$f.Name; Path=$f.FullName; Size=($script:Walker.FileAlloc($f.FullName)); IsDir=$false }
        }
    } catch {}

    $shown = 0
    foreach ($c in ($items | Sort-Object Size -Descending)) {
        if ($shown -ge $MAX_CHILDREN) {
            [void]$node.Nodes.Add((New-Object System.Windows.Forms.TreeNode("... (+$($items.Count - $shown) itens nao exibidos)")))
            break
        }
        $pct = if ($parentSize -gt 0) { 100.0 * $c.Size / $parentSize } else { 0 }
        $icon = if ($c.IsDir) { '[DIR]' } else { '     ' }
        $tn = New-Object System.Windows.Forms.TreeNode
        $tn.Text = ('{0} {1}  {2}  ({3:N1}%)' -f $icon, $c.Name, (Format-Size $c.Size), $pct)
        $tn.Name = $c.Path     # caminho completo (para Explorer/Copiar/Deletar e re-expansao)
        if ($c.IsDir) { [void]$tn.Nodes.Add((New-Object System.Windows.Forms.TreeNode('...'))) }
        [void]$node.Nodes.Add($tn)
        $shown++
    }
}

# ============================================================================
#  GUI
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Show-DiskUsage (Casco Digital) - Admin'
$form.Size = New-Object System.Drawing.Size(940,700)
$form.StartPosition = 'CenterScreen'

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = 'Top'; $panel.Height = 40
$txt = New-Object System.Windows.Forms.TextBox
$txt.Text = 'C:\'; $txt.Location = '8,9'; $txt.Width = 720
$btn = New-Object System.Windows.Forms.Button
$btn.Text = 'Scan'; $btn.Location = '740,8'; $btn.Width = 80; $btn.Height = 24
$panel.Controls.AddRange(@($txt,$btn))

$status = New-Object System.Windows.Forms.StatusStrip
$lbl = New-Object System.Windows.Forms.ToolStripStatusLabel
$lbl.Text = 'Pronto. Digite um caminho (ex: C:\) e clique Scan. A varredura pode levar 1-3 min.'
[void]$status.Items.Add($lbl)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Dock = 'Fill'
$tree.Font = New-Object System.Drawing.Font('Consolas',10)
$tree.HideSelection = $false

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = $menu.Items.Add('Abrir no Explorer')
$miCopy = $menu.Items.Add('Copiar caminho')
$miDel  = $menu.Items.Add('Deletar...')
$tree.ContextMenuStrip = $menu

$tree.Add_NodeMouseClick({
    param($s,$e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { $tree.SelectedNode = $e.Node }
})
$miOpen.Add_Click({
    $n = $tree.SelectedNode; if (-not $n -or -not $n.Name) { return }
    $p = $n.Name
    if (Test-Path -LiteralPath $p -PathType Container) { Start-Process explorer.exe $p }
    elseif (Test-Path -LiteralPath $p) { Start-Process explorer.exe "/select,`"$p`"" }
})
$miCopy.Add_Click({
    $n = $tree.SelectedNode; if ($n -and $n.Name) { [System.Windows.Forms.Clipboard]::SetText($n.Name) }
})
$miDel.Add_Click({
    $n = $tree.SelectedNode; if (-not $n -or -not $n.Name) { return }
    $p = $n.Name
    $r = [System.Windows.Forms.MessageBox]::Show("Deletar PERMANENTEMENTE (sem Lixeira)?`n`n$p","Confirmar",'YesNo','Warning')
    if ($r -eq 'Yes') {
        try {
            Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop
            $n.Remove(); $lbl.Text = "Deletado: $p"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Falhou: $($_.Exception.Message)","Erro") | Out-Null
        }
    }
})

$tree.Add_BeforeExpand({
    param($s,$e)
    $n = $e.Node
    if ($n.Nodes.Count -eq 1 -and $n.Nodes[0].Text -eq '...') { Build-Children $n }
})

$btn.Add_Click({
    $root = $txt.Text
    if (-not (Test-Path -LiteralPath $root)) {
        [System.Windows.Forms.MessageBox]::Show("Caminho invalido.","Erro") | Out-Null; return
    }
    $tree.Nodes.Clear()
    $form.Cursor = 'WaitCursor'
    $script:Walker = New-Object CascoDigital.DiskWalker
    $rootKey = $root.TrimEnd('\')
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $total = 0L

    # Percorre as pastas de topo uma a uma, so para dar progresso visivel no status
    $topDirs = @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)
    foreach ($d in $topDirs) {
        if (Test-Reparse $d) { continue }
        $lbl.Text = "Varrendo $($d.FullName) ...  ($($script:Walker.FileCount) arquivos ate agora)"
        [System.Windows.Forms.Application]::DoEvents()
        $total += $script:Walker.ScanTree($d.FullName)
    }
    # Arquivos soltos na raiz
    foreach ($f in (Get-ChildItem -LiteralPath $root -File -Force -ErrorAction SilentlyContinue)) {
        $total += $script:Walker.FileAlloc($f.FullName)
    }
    $script:Walker.DirSize[$rootKey] = $total
    $sw.Stop()

    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Name = $root   # caminho completo COM barra ("C:\"); "C:" sem barra = dir atual da unidade
    $rootNode.Text = ('[DIR] {0}  {1}  (100%)' -f $root, (Format-Size $total))
    [void]$rootNode.Nodes.Add((New-Object System.Windows.Forms.TreeNode('...')))
    [void]$tree.Nodes.Add($rootNode)
    Build-Children $rootNode
    $rootNode.Expand()
    $form.Cursor = 'Default'
    $lbl.Text = "OK: $($script:Walker.FileCount) arquivos em $([math]::Round($sw.Elapsed.TotalSeconds,1))s  |  $root (fisico) = $(Format-Size $total)"
})

$form.Controls.Add($tree)
$form.Controls.Add($status)
$form.Controls.Add($panel)
[void]$form.ShowDialog()

}
catch {
    $msg = ($_ | Out-String) + "`n`n" + ($_.ScriptStackTrace | Out-String)
    try { Set-Content -Path "$env:USERPROFILE\Desktop\Show-DiskUsage-error.log" -Value $msg -Encoding UTF8 } catch {}
    try {
        [System.Windows.Forms.MessageBox]::Show(
            $msg.Substring(0, [Math]::Min(2000, $msg.Length)),
            'Show-DiskUsage - ERRO FATAL', 'OK', 'Error') | Out-Null
    } catch {
        Write-Host $msg -ForegroundColor Red
        Read-Host 'Pressione Enter para sair'
    }
}
