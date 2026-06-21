#requires -version 5.1
<#
.SYNOPSIS
    Show-DiskUsage.ps1 - Visualizador de uso de disco estilo TreeSize/WizTree em PowerShell + WinForms.
    Solucao propria, sem instalar software de terceiros (compliance-safe).

.DESCRIPTION
    Le a Master File Table ($MFT) crua do volume NTFS (mesma tecnica do WizTree) -> scan do disco
    inteiro em segundos, nao minutos. Apos o scan, a navegacao e instantanea.

    - Auto-eleva para Administrador (obrigatorio para abrir o handle do volume)
    - Modo MFT: scan ultra-rapido de volumes NTFS locais (C:, D:, ...)
    - Fallback automatico para enumeracao classica (Get-ChildItem) se a MFT nao puder ser lida
      (volumes de rede, ReFS/FAT, falha de acesso)
    - Maior sempre no topo, com percentual relativo ao pai
    - Mostra pastas E arquivos
    - Menu de contexto (botao direito): Abrir no Explorer / Copiar caminho / Deletar (com confirmacao)

.NOTES
    Requer Administrador e PowerShell 5.1+. O modo MFT so funciona em volume NTFS local fixo.
    A MFT entrega o tamanho logico (real) dos dados.

.EXAMPLE
    PowerShell -ExecutionPolicy Bypass -File .\tools\Show-DiskUsage.ps1
#>

# ---- Auto-elevacao ----
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
$pr = New-Object Security.Principal.WindowsPrincipal($id)
if (-not $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList @(
        '-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`""
    )
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
#  Parser nativo da $MFT (NTFS) em C#. Le o volume cru e reconstroi a arvore.
# ============================================================================
$cs = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace CascoDigital {
  public class MftNode {
    public long Index;
    public long Parent;
    public string Name;
    public long Size;           // tamanho proprio (arquivo); pasta = 0
    public long RecursiveSize;  // soma recursiva (pasta); arquivo = Size
    public bool IsDir;
    public long Display { get { return IsDir ? RecursiveSize : Size; } }
  }

  public class MftScanner {
    public const long ROOT = 5;
    public char Drive;
    public Dictionary<long, MftNode> Nodes = new Dictionary<long, MftNode>();
    public Dictionary<long, List<long>> Kids = new Dictionary<long, List<long>>();
    public int FileCount;

    [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Auto)]
    static extern SafeFileHandle CreateFile(string name, uint access, uint share,
        IntPtr sec, uint disp, uint flags, IntPtr templ);

    static ushort U16(byte[] b, int o) { return BitConverter.ToUInt16(b, o); }
    static uint   U32(byte[] b, int o) { return BitConverter.ToUInt32(b, o); }
    static long   I64(byte[] b, int o) { return BitConverter.ToInt64(b, o); }

    public static MftScanner Scan(char drive) {
      MftScanner s = new MftScanner();
      s.Drive = char.ToUpper(drive);
      string vol = "\\\\.\\" + s.Drive + ":";
      SafeFileHandle h = CreateFile(vol, 0x80000000, 0x00000003, IntPtr.Zero, 3, 0, IntPtr.Zero);
      if (h.IsInvalid) throw new IOException("CreateFile falhou no volume " + vol + " (erro " + Marshal.GetLastWin32Error() + ")");

      using (FileStream fs = new FileStream(h, FileAccess.Read)) {
        // --- Boot sector / BPB ---
        byte[] boot = new byte[512];
        fs.Seek(0, SeekOrigin.Begin);
        ReadFull(fs, boot, 0, 512);
        if (boot[3] != (byte)'N' || boot[4] != (byte)'T' || boot[5] != (byte)'F' || boot[6] != (byte)'S')
          throw new IOException("Volume nao e NTFS.");

        int bytesPerSector = U16(boot, 0x0B);
        int secsPerCluster = boot[0x0D];
        long bytesPerCluster = (long)bytesPerSector * secsPerCluster;
        long mftLcn = I64(boot, 0x30);
        sbyte cpr = (sbyte)boot[0x40];
        int recSize = cpr > 0 ? (int)(cpr * bytesPerCluster) : (1 << (-cpr));

        // --- Record 0 ($MFT) para descobrir os extents da propria MFT ---
        byte[] rec0 = new byte[recSize];
        fs.Seek(mftLcn * bytesPerCluster, SeekOrigin.Begin);
        ReadFull(fs, rec0, 0, recSize);
        ApplyFixup(rec0, bytesPerSector);
        List<long[]> extents = ParseMftDataRuns(rec0, bytesPerCluster); // [diskOffset, byteLength]

        // --- Itera todos os registros da MFT lendo os extents em blocos ---
        byte[] rec = new byte[recSize];
        long index = 0;
        const int CHUNK = 8 * 1024 * 1024;
        byte[] buf = new byte[CHUNK];

        foreach (long[] ex in extents) {
          long off = ex[0];
          long remaining = ex[1];
          fs.Seek(off, SeekOrigin.Begin);
          while (remaining > 0) {
            int want = (int)Math.Min((long)CHUNK, remaining);
            want -= want % recSize;
            if (want <= 0) break;
            ReadFull(fs, buf, 0, want);
            remaining -= want;
            for (int p = 0; p + recSize <= want; p += recSize, index++) {
              Buffer.BlockCopy(buf, p, rec, 0, recSize);
              s.ParseRecord(rec, index, bytesPerSector);
            }
          }
        }
      }

      s.Rollup();
      s.BuildKids();
      return s;
    }

    static void ReadFull(FileStream fs, byte[] b, int off, int count) {
      int got = 0;
      while (got < count) {
        int n = fs.Read(b, off + got, count - got);
        if (n <= 0) throw new IOException("Leitura do volume terminou cedo.");
        got += n;
      }
    }

    // Aplica o Update Sequence Array (fixup) ao registro
    static void ApplyFixup(byte[] rec, int bytesPerSector) {
      int usaOff = U16(rec, 0x04);
      int usaCnt = U16(rec, 0x06);
      if (usaCnt == 0) return;
      // primeira entrada = valor de verificacao; demais substituem os ultimos 2 bytes de cada setor
      for (int i = 1; i < usaCnt; i++) {
        int sectorEnd = i * bytesPerSector - 2;
        if (sectorEnd + 1 >= rec.Length) break;
        rec[sectorEnd]     = rec[usaOff + i * 2];
        rec[sectorEnd + 1] = rec[usaOff + i * 2 + 1];
      }
    }

    // Le os data runs do $DATA nao-residente do registro 0 -> extents absolutos da MFT
    static List<long[]> ParseMftDataRuns(byte[] rec, long bytesPerCluster) {
      List<long[]> ext = new List<long[]>();
      int attrOff = U16(rec, 0x14);
      int pos = attrOff;
      while (pos + 4 <= rec.Length) {
        uint type = U32(rec, pos);
        if (type == 0xFFFFFFFF) break;
        uint len = U32(rec, pos + 4);
        if (len == 0) break;
        if (type == 0x80 && rec[pos + 8] == 1) { // $DATA nao-residente
          int runOff = U16(rec, pos + 0x20);
          int rp = pos + runOff;
          long lcn = 0;
          while (rp < rec.Length && rec[rp] != 0) {
            int header = rec[rp++];
            int lenBytes = header & 0x0F;
            int offBytes = (header >> 4) & 0x0F;
            long runLen = 0;
            for (int i = 0; i < lenBytes; i++) runLen |= (long)rec[rp++] << (8 * i);
            long runOffVal = 0;
            for (int i = 0; i < offBytes; i++) runOffVal |= (long)rec[rp++] << (8 * i);
            if (offBytes > 0 && (rec[rp - 1] & 0x80) != 0) // sinal (relativo)
              runOffVal |= (-1L) << (8 * offBytes);
            lcn += runOffVal;
            ext.Add(new long[] { lcn * bytesPerCluster, runLen * bytesPerCluster });
          }
          break;
        }
        pos += (int)len;
      }
      return ext;
    }

    void ParseRecord(byte[] rec, long index, int bytesPerSector) {
      if (rec[0] != (byte)'F' || rec[1] != (byte)'I' || rec[2] != (byte)'L' || rec[3] != (byte)'E') return;
      ApplyFixup(rec, bytesPerSector);
      ushort flags = U16(rec, 0x16);
      if ((flags & 0x01) == 0) return;        // registro nao em uso
      bool isDir = (flags & 0x02) != 0;

      int attrOff = U16(rec, 0x14);
      int pos = attrOff;
      long parent = -1;
      string name = null;
      int nameNs = -1;
      long size = 0;
      bool gotData = false;

      while (pos + 4 <= rec.Length) {
        uint type = U32(rec, pos);
        if (type == 0xFFFFFFFF) break;
        uint len = U32(rec, pos + 4);
        if (len == 0 || pos + (int)len > rec.Length) break;
        byte nonResident = rec[pos + 8];

        if (type == 0x30) { // FILE_NAME (residente)
          int vOff = U16(rec, pos + 0x14);
          int v = pos + vOff;
          long pref = I64(rec, v + 0x00) & 0x0000FFFFFFFFFFFF;
          int nLen = rec[v + 0x40];
          int ns = rec[v + 0x41];
          // prefere Win32 (1) ou Win32&DOS (3); ignora DOS puro (2) se ja temos algo melhor
          if (ns != 2 || name == null) {
            if (name == null || ns != 2) {
              parent = pref;
              name = Encoding.Unicode.GetString(rec, v + 0x42, nLen * 2);
              nameNs = ns;
            }
          }
        } else if (type == 0x80 && rec[pos + 9] == 0 && !gotData) { // $DATA sem nome
          if (nonResident == 0) {
            size = U32(rec, pos + 0x10);            // tamanho residente
          } else {
            size = I64(rec, pos + 0x30);            // real size nao-residente
          }
          gotData = true;
        }
        pos += (int)len;
      }

      if (name == null) return;
      MftNode n = new MftNode {
        Index = index, Parent = parent, Name = name,
        Size = isDir ? 0 : size, IsDir = isDir
      };
      Nodes[index] = n;
      if (!isDir) FileCount++;
    }

    void Rollup() {
      foreach (MftNode n in Nodes.Values) {
        if (n.IsDir) { n.RecursiveSize = n.RecursiveSize; } // pastas somam via filhos abaixo
      }
      foreach (MftNode n in Nodes.Values) {
        long s = n.Size;
        if (s <= 0) continue;
        long p = n.Parent;
        int guard = 0;
        while (guard++ < 512) {
          MftNode pn;
          if (!Nodes.TryGetValue(p, out pn)) break;
          pn.RecursiveSize += s;
          if (p == ROOT) break;
          if (pn.Parent == p) break;
          p = pn.Parent;
        }
      }
      // arquivo: display usa Size; pasta: RecursiveSize (ja calculado)
    }

    void BuildKids() {
      foreach (MftNode n in Nodes.Values) {
        if (n.Index == ROOT) continue;
        List<long> lst;
        if (!Kids.TryGetValue(n.Parent, out lst)) { lst = new List<long>(); Kids[n.Parent] = lst; }
        lst.Add(n.Index);
      }
      foreach (List<long> lst in Kids.Values) {
        lst.Sort(delegate (long a, long b) {
          long da = Nodes.ContainsKey(a) ? Nodes[a].Display : 0;
          long db = Nodes.ContainsKey(b) ? Nodes[b].Display : 0;
          return db.CompareTo(da);
        });
      }
    }

    public long[] Children(long idx) {
      List<long> lst;
      if (Kids.TryGetValue(idx, out lst)) return lst.ToArray();
      return new long[0];
    }

    public MftNode Get(long idx) {
      MftNode n; return Nodes.TryGetValue(idx, out n) ? n : null;
    }

    public long FindIndex(string path) {
      string p = path.Trim();
      int colon = p.IndexOf(':');
      if (colon >= 0) p = p.Substring(colon + 1);
      p = p.Trim('\\');
      if (p.Length == 0) return ROOT;
      long cur = ROOT;
      foreach (string part in p.Split('\\')) {
        if (part.Length == 0) continue;
        long found = -1;
        foreach (long c in Children(cur)) {
          MftNode n = Get(c);
          if (n != null && n.IsDir && string.Equals(n.Name, part, StringComparison.OrdinalIgnoreCase)) { found = c; break; }
        }
        if (found < 0) return -1;
        cur = found;
      }
      return cur;
    }

    public string FullPath(long idx) {
      if (idx == ROOT) return Drive + ":\\";
      Stack<string> parts = new Stack<string>();
      long cur = idx;
      int guard = 0;
      while (cur != ROOT && guard++ < 512) {
        MftNode n = Get(cur);
        if (n == null) break;
        parts.Push(n.Name);
        if (n.Parent == cur) break;
        cur = n.Parent;
      }
      return Drive + ":\\" + string.Join("\\", parts.ToArray());
    }
  }
}
'@
try { Add-Type -TypeDefinition $cs -Language CSharp -ErrorAction Stop } catch { throw "Falha ao compilar o parser MFT: $($_.Exception.Message)" }

# ============================================================================
#  Estado e helpers
# ============================================================================
$script:Mode    = 'WALK'   # 'MFT' ou 'WALK'
$script:Scanner = $null
$script:DirSize = @{}       # usado no modo WALK
$MAX_CHILDREN   = 5000

function Format-Size([long]$bytes) {
    if ($bytes -ge 1TB) { return ('{0:N2} TB' -f ($bytes / 1TB)) }
    if ($bytes -ge 1GB) { return ('{0:N2} GB' -f ($bytes / 1GB)) }
    if ($bytes -ge 1MB) { return ('{0:N2} MB' -f ($bytes / 1MB)) }
    if ($bytes -ge 1KB) { return ('{0:N2} KB' -f ($bytes / 1KB)) }
    return "$bytes B"
}

# Retorna filhos como objetos uniformes: Name, Path, Size, IsDir, HasChildren
function Get-Children($node) {
    $items = New-Object System.Collections.ArrayList
    if ($script:Mode -eq 'MFT') {
        $idx = [long]$node.Tag
        foreach ($cid in $script:Scanner.Children($idx)) {
            $n = $script:Scanner.Get($cid)
            if ($null -eq $n) { continue }
            [void]$items.Add([pscustomobject]@{
                Name=$n.Name; Path=$null; Size=$n.Display; IsDir=$n.IsDir;
                HasChildren=($n.IsDir -and $script:Scanner.Children($cid).Length -gt 0); Key=$cid
            })
        }
        return $items   # ja vem ordenado desc do C#
    }
    # WALK
    $path = $node.Name
    $parentSize = Get-CachedSize $path
    $tmp = @()
    try {
        foreach ($d in (Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction SilentlyContinue)) {
            $tmp += [pscustomobject]@{ Name=$d.Name; Path=$d.FullName; Size=(Get-CachedSize $d.FullName); IsDir=$true; HasChildren=$true; Key=$d.FullName }
        }
        foreach ($f in (Get-ChildItem -LiteralPath $path -File -Force -ErrorAction SilentlyContinue)) {
            $tmp += [pscustomobject]@{ Name=$f.Name; Path=$f.FullName; Size=[long]$f.Length; IsDir=$false; HasChildren=$false; Key=$f.FullName }
        }
    } catch {}
    foreach ($c in ($tmp | Sort-Object Size -Descending)) { [void]$items.Add($c) }
    return $items
}

# ---- Modo WALK (fallback): scan unico recursivo, cache em $script:DirSize ----
function Get-CachedSize([string]$path) {
    $k = $path.TrimEnd('\')
    if ($script:DirSize.ContainsKey($k)) { return [long]$script:DirSize[$k] }
    return 0L
}
function Invoke-WalkScan([string]$root) {
    $script:DirSize = @{}
    $rootKey = (Resolve-Path -LiteralPath $root).Path.TrimEnd('\')
    $count = 0
    Get-ChildItem -LiteralPath $root -File -Force -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $len = $_.Length; $d = $_.DirectoryName
        while ($d -and $d.Length -ge $rootKey.Length -and
               $d.StartsWith($rootKey, [System.StringComparison]::OrdinalIgnoreCase)) {
            $script:DirSize[$d] = [long]($script:DirSize[$d]) + $len
            if ($d.Length -le $rootKey.Length) { break }
            $d = [System.IO.Path]::GetDirectoryName($d)
        }
        $count++
        if (($count % 2000) -eq 0) {
            $form.Text = "Show-DiskUsage [WALK] - escaneando... $count arquivos"
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    return $rootKey
}

# Popula os filhos de um node (instantaneo nos dois modos)
function Build-Children($node) {
    $node.Nodes.Clear()
    $children = Get-Children $node
    $shown = 0
    foreach ($c in $children) {
        if ($shown -ge $MAX_CHILDREN) {
            $more = New-Object System.Windows.Forms.TreeNode("... (+$($children.Count - $shown) itens nao exibidos)")
            [void]$node.Nodes.Add($more); break
        }
        $pct = 0.0
        $parentDisp = if ($script:Mode -eq 'MFT') { $script:Scanner.Get([long]$node.Tag).Display } else { Get-CachedSize $node.Name }
        if ($parentDisp -gt 0) { $pct = 100.0 * $c.Size / $parentDisp }
        $icon = if ($c.IsDir) { '[DIR]' } else { '     ' }
        $tn = New-Object System.Windows.Forms.TreeNode
        $tn.Text = ('{0} {1}  {2}  ({3:N1}%)' -f $icon, $c.Name, (Format-Size $c.Size), $pct)
        $tn.Tag  = $c.Key
        # .Name guarda o caminho completo (para Explorer/Copiar/Deletar)
        $tn.Name = if ($script:Mode -eq 'MFT') { $script:Scanner.FullPath([long]$c.Key) } else { $c.Path }
        if ($c.HasChildren) { [void]$tn.Nodes.Add((New-Object System.Windows.Forms.TreeNode('...'))) }
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
$lbl.Text = 'Pronto. Digite um caminho (ex: C:\) e clique Scan.'
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
            $n.Remove()
            $lbl.Text = "Deletado: $p"
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
    $rootNode = New-Object System.Windows.Forms.TreeNode

    # Tenta modo MFT se for raiz/caminho de um drive NTFS local
    $drive = $null
    if ($root -match '^[A-Za-z]:') { $drive = $root.Substring(0,1) }
    $useMft = $false
    if ($drive) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $lbl.Text = "Lendo a MFT de $drive`: ..."; [System.Windows.Forms.Application]::DoEvents()
            $script:Scanner = [CascoDigital.MftScanner]::Scan([char]$drive)
            $script:Mode = 'MFT'
            $useMft = $true
            $sw.Stop()
            $startIdx = $script:Scanner.FindIndex($root)
            if ($startIdx -lt 0) { $startIdx = [CascoDigital.MftScanner]::ROOT }
            $sn = $script:Scanner.Get($startIdx)
            $disp = if ($sn) { $sn.Display } else { 0 }
            $rootNode.Tag  = [long]$startIdx
            $rootNode.Name = $script:Scanner.FullPath($startIdx)
            $rootNode.Text = ('[DIR] {0}  {1}  (100%)' -f $rootNode.Name, (Format-Size $disp))
            $lbl.Text = "MFT: $($script:Scanner.FileCount) arquivos em $([math]::Round($sw.Elapsed.TotalSeconds,1))s  |  $($rootNode.Name) = $(Format-Size $disp)"
        } catch {
            $useMft = $false
            $lbl.Text = "MFT indisponivel ($($_.Exception.Message)). Usando enumeracao classica..."
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    if (-not $useMft) {
        $script:Mode = 'WALK'
        $rootKey = Invoke-WalkScan $root
        $rootNode.Tag  = $rootKey
        $rootNode.Name = $rootKey
        $rootNode.Text = ('[DIR] {0}  {1}  (100%)' -f $root, (Format-Size (Get-CachedSize $rootKey)))
        $lbl.Text = "WALK: $root = $(Format-Size (Get-CachedSize $rootKey))"
    }

    [void]$tree.Nodes.Add($rootNode)
    Build-Children $rootNode
    $rootNode.Expand()
    $form.Cursor = 'Default'
})

$form.Controls.Add($tree)
$form.Controls.Add($status)
$form.Controls.Add($panel)
[void]$form.ShowDialog()
