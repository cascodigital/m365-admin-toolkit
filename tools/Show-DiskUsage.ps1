#requires -version 5.1
<#
.SYNOPSIS
    Show-DiskUsage.ps1 - Visualizador de uso de disco estilo TreeSize/WizTree em PowerShell + WinForms.
    Solucao propria, sem instalar software de terceiros (compliance-safe).

.DESCRIPTION
    - Auto-eleva para Administrador (varre pastas de sistema sem Access Denied)
    - Faz UM unico passe de scan e cacheia os tamanhos -> navegacao instantanea depois
    - Maior sempre no topo, com percentual relativo ao pai
    - Mostra pastas E arquivos
    - Menu de contexto (botao direito): Abrir no Explorer / Copiar caminho / Deletar (com confirmacao)

.NOTES
    Limitacao: o scan inicial enumera arquivo a arquivo (mais lento que o WizTree, que le a MFT do NTFS
    em codigo nativo). Apos o primeiro scan, a navegacao e instantanea. Para discos inteiros muito grandes,
    prefira apontar para uma subpasta especifica.

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

# ---- Estado global ----
$script:DirSize = @{}   # caminho-da-pasta -> tamanho recursivo (bytes)

function Format-Size([long]$bytes) {
    if ($bytes -ge 1TB) { return ('{0:N2} TB' -f ($bytes / 1TB)) }
    if ($bytes -ge 1GB) { return ('{0:N2} GB' -f ($bytes / 1GB)) }
    if ($bytes -ge 1MB) { return ('{0:N2} MB' -f ($bytes / 1MB)) }
    if ($bytes -ge 1KB) { return ('{0:N2} KB' -f ($bytes / 1KB)) }
    return "$bytes B"
}

# Scan unico: soma cada arquivo em todos os seus diretorios-pai ate o root.
function Invoke-Scan([string]$root) {
    $script:DirSize = @{}
    $rootNorm = (Resolve-Path -LiteralPath $root).Path.TrimEnd('\')
    if ($rootNorm.Length -eq 2) { $rootNorm += '\' }  # "C:" -> "C:\"
    $rootKey = $rootNorm.TrimEnd('\')

    $count = 0
    Get-ChildItem -LiteralPath $root -File -Force -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $len = $_.Length
        $d   = $_.DirectoryName
        while ($d -and $d.Length -ge $rootKey.Length -and
               $d.StartsWith($rootKey, [System.StringComparison]::OrdinalIgnoreCase)) {
            $script:DirSize[$d] = [long]($script:DirSize[$d]) + $len
            if ($d.Length -le $rootKey.Length) { break }
            $d = [System.IO.Path]::GetDirectoryName($d)
        }
        $count++
        if (($count % 2000) -eq 0) {
            $form.Text = "TreeSize Lite - escaneando... $count arquivos ($(Format-Size $script:DirSize[$rootKey]))"
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    return $rootKey
}

# Tamanho recursivo de uma pasta a partir do cache (0 se vazia/sem cache)
function Get-CachedSize([string]$path) {
    $k = $path.TrimEnd('\')
    if ($script:DirSize.ContainsKey($k)) { return [long]$script:DirSize[$k] }
    return 0L
}

# Monta os filhos diretos de um node (pastas + arquivos), ordenado desc. Instantaneo (usa cache).
function Build-Children($node) {
    $node.Nodes.Clear()
    $path = $node.Tag
    $parentSize = Get-CachedSize $path
    $items = @()

    try {
        foreach ($d in (Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction SilentlyContinue)) {
            $items += [pscustomobject]@{ Name=$d.Name; Path=$d.FullName; Size=(Get-CachedSize $d.FullName); IsDir=$true }
        }
        foreach ($f in (Get-ChildItem -LiteralPath $path -File -Force -ErrorAction SilentlyContinue)) {
            $items += [pscustomobject]@{ Name=$f.Name; Path=$f.FullName; Size=[long]$f.Length; IsDir=$false }
        }
    } catch {}

    foreach ($c in ($items | Sort-Object Size -Descending)) {
        $pct = if ($parentSize -gt 0) { 100.0 * $c.Size / $parentSize } else { 0 }
        $icon = if ($c.IsDir) { '[DIR]' } else { '     ' }
        $node2 = New-Object System.Windows.Forms.TreeNode
        $node2.Text = ('{0} {1}  {2}  ({3:N1}%)' -f $icon, $c.Name, (Format-Size $c.Size), $pct)
        $node2.Tag  = $c.Path
        if ($c.IsDir) {
            [void]$node2.Nodes.Add((New-Object System.Windows.Forms.TreeNode('...')))  # placeholder lazy
        }
        [void]$node.Nodes.Add($node2)
    }
}

# ---- GUI ----
$form = New-Object System.Windows.Forms.Form
$form.Text = 'TreeSize Lite (Skippy Edition) - Admin'
$form.Size = New-Object System.Drawing.Size(900,680)
$form.StartPosition = 'CenterScreen'

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = 'Top'; $panel.Height = 38
$txt = New-Object System.Windows.Forms.TextBox
$txt.Text = 'C:\'; $txt.Location = '8,8'; $txt.Width = 700
$btn = New-Object System.Windows.Forms.Button
$btn.Text = 'Scan'; $btn.Location = '720,7'; $btn.Width = 80
$panel.Controls.AddRange(@($txt,$btn))

$tree = New-Object System.Windows.Forms.TreeView
$tree.Dock = 'Fill'
$tree.Font = New-Object System.Drawing.Font('Consolas',10)
$tree.HideSelection = $false

# Menu de contexto (botao direito)
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
    $n = $tree.SelectedNode; if (-not $n) { return }
    $p = $n.Tag
    if (Test-Path -LiteralPath $p -PathType Container) { Start-Process explorer.exe $p }
    else { Start-Process explorer.exe "/select,`"$p`"" }
})
$miCopy.Add_Click({
    $n = $tree.SelectedNode; if ($n) { [System.Windows.Forms.Clipboard]::SetText($n.Tag) }
})
$miDel.Add_Click({
    $n = $tree.SelectedNode; if (-not $n) { return }
    $p = $n.Tag
    $r = [System.Windows.Forms.MessageBox]::Show("Deletar PERMANENTEMENTE?`n`n$p","Confirmar",'YesNo','Warning')
    if ($r -eq 'Yes') {
        try {
            Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop
            $n.Remove()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Falhou: $($_.Exception.Message)","Erro") | Out-Null
        }
    }
})

# Lazy: ao expandir, troca o placeholder pelos filhos reais (do cache)
$tree.Add_BeforeExpand({
    param($s,$e)
    $n = $e.Node
    if ($n.Nodes.Count -eq 1 -and $n.Nodes[0].Text -eq '...') { Build-Children $n }
})

$btn.Add_Click({
    $root = $txt.Text
    if (-not (Test-Path -LiteralPath $root)) {
        [System.Windows.Forms.MessageBox]::Show("Caminho invalido, macaco.","Erro") | Out-Null
        return
    }
    $tree.Nodes.Clear()
    $form.Cursor = 'WaitCursor'
    $rootKey = Invoke-Scan $root
    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Text = ('[DIR] {0}  {1}  (100%)' -f $root, (Format-Size (Get-CachedSize $rootKey)))
    $rootNode.Tag  = $rootKey
    [void]$tree.Nodes.Add($rootNode)
    Build-Children $rootNode
    $rootNode.Expand()
    $form.Cursor = 'Default'
    $form.Text = "TreeSize Lite - $root  =  $(Format-Size (Get-CachedSize $rootKey))"
})

$form.Controls.Add($tree)
$form.Controls.Add($panel)
[void]$form.ShowDialog()
