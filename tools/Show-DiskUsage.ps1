#requires -version 5.1
<#
.SYNOPSIS
    Show-DiskUsage.ps1 - TreeSize-style disk usage viewer built with PowerShell and WinForms.
    Self-contained; no third-party software to install.

.DESCRIPTION
    Asks Windows for the PHYSICAL allocated size of each file (GetCompressedFileSizeW),
    so the numbers are close to the "Allocated" column in TreeSize/WizTree:
      - NTFS compression and sparse files are accounted for by the OS;
      - cloud-only files (OneDrive Files On-Demand, dehydrated) allocate ~0 and are not counted;
      - directory junctions/symlinks (reparse points) are not traversed (avoids double counting);
      - hard links are NOT de-duplicated: a file with several hard links is counted once per
        link, so totals can be higher than TreeSize/WizTree (which de-duplicate) on folders
        such as C:\Windows\WinSxS.

    - Largest items always on top, with percentage of the parent
    - Shows folders AND files; navigable tree (instant expansion from cache)
    - Context menu (right-click): Open in Explorer / Copy path

.NOTES
    Requires Administrator and PowerShell 5.1+ (Windows PowerShell or PowerShell 7 on Windows).
    The scan walks the file system (slower than reading the raw MFT: 1-3 min for a full C:),
    in exchange for a reliable physical number. The GUI is busy during the scan; the status
    bar shows which top-level folder is being processed.

    If the script is not elevated, or not running in an STA thread (required by WinForms),
    it relaunches itself with -STA (and UAC elevation if needed) using the same PowerShell
    edition: pwsh when started from pwsh, otherwise powershell.exe.

.EXAMPLE
    PowerShell -ExecutionPolicy Bypass -File .\tools\Show-DiskUsage.ps1

.EXAMPLE
    pwsh -ExecutionPolicy Bypass -File .\tools\Show-DiskUsage.ps1
#>

# ---- Requires Administrator and an STA thread; relaunch if needed (bypass only for that process) ----
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
$pr = New-Object Security.Principal.WindowsPrincipal($id)
$isAdmin = $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$isSta = [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA

if (-not $isAdmin -or -not $isSta) {
    Write-Host ''
    if (-not $isAdmin) {
        Write-Host '  This script requires Administrator. Requesting elevation via UAC...' -ForegroundColor Yellow
    } else {
        Write-Host '  WinForms requires an STA thread. Relaunching with -STA...' -ForegroundColor Yellow
    }

    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    if ($scriptPath) {
        # Prefer the same PowerShell edition; fall back to Windows PowerShell 5.1
        $hostExe = $null
        if ($PSVersionTable.PSEdition -eq 'Core') {
            $candidate = Join-Path $PSHOME 'pwsh.exe'
            if (Test-Path -LiteralPath $candidate) { $hostExe = $candidate }
        }
        if (-not $hostExe) {
            $hostExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        }

        $relaunchArgs = @(
            '-NoProfile',
            '-STA',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            "`"$scriptPath`""
        )

        if (-not $isAdmin) {
            Start-Process -FilePath $hostExe -Verb RunAs -ArgumentList $relaunchArgs
        } else {
            Start-Process -FilePath $hostExe -ArgumentList $relaunchArgs
        }
        exit 0
    }

    Write-Host '  Could not determine the path of the current script.' -ForegroundColor Red
    Write-Host '  Run: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tools\Show-DiskUsage.ps1' -ForegroundColor Yellow
    Write-Host ''
    Read-Host '  Press Enter to exit'
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Safety net: any fatal error becomes a MessageBox + log (otherwise the window closes silently)
try {

# ============================================================================
#  Engine: walks the file system and gets the PHYSICAL allocated size of each file
#  directly from Windows (GetCompressedFileSizeW). No manual NTFS parsing.
# ============================================================================
$cs = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace DiskUsageTool {
  public class DiskWalker {
    [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern uint GetCompressedFileSizeW(string path, out uint high);

    const uint INVALID_SIZE = 0xFFFFFFFF;
    const long SANITY_CAP   = 0x4000000000000L; // 1 PB: a single file larger than this is treated as garbage

    public Dictionary<string, long> DirSize =
        new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
    public long FileCount;

    // \\?\ prefix to get past the MAX_PATH (260) limit in the size call
    static string Long(string p) {
      if (p.StartsWith("\\\\?\\")) return p;
      if (p.StartsWith("\\\\"))    return "\\\\?\\UNC\\" + p.Substring(2);
      return "\\\\?\\" + p;
    }

    // Normalizes the cache key. "C:\" and "C:" become "C:"; other paths lose the trailing backslash.
    static string Key(string p) {
      p = p.TrimEnd('\\');
      return p;
    }

    // PHYSICAL allocated size of a file, from Windows. 0 on failure or garbage values.
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

    // Walks a subtree ITERATIVELY (explicit stack -> no StackOverflow), using .NET
    // enumeration (DirectoryInfo) to traverse and GetCompressedFileSize to measure.
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
          // "C:" must become "C:\" so DirectoryInfo points to the root, not the current directory
          string dpath = (dir.Length == 2 && dir[1] == ':') ? dir + "\\" : dir;
          DirectoryInfo di = new DirectoryInfo(dpath);
          foreach (FileSystemInfo fsi in di.EnumerateFileSystemInfos()) {
            FileAttributes a;
            try { a = fsi.Attributes; } catch { continue; }
            bool isDir = (a & FileAttributes.Directory)     != 0;
            bool isRep = (a & FileAttributes.ReparsePoint)  != 0;
            if (isDir) {
              if (isRep) continue;                 // junction/symlink: not traversed
              string full = Key(fsi.FullName);
              if (!direct.ContainsKey(full)) {     // guard against any cycle
                direct[full] = 0; parent[full] = dir; order.Add(full); stack.Push(full);
              }
            } else {
              files += FileAlloc(fsi.FullName);
              FileCount++;
            }
          }
        } catch { }                                // access denied / long path: skip the folder
        direct[dir] = files;                       // bytes of the files directly in this folder
      }

      // Roll up from deepest to shallowest: each folder adds its total to its parent.
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
# Namespace versioned by a hash of the C# source: each code version becomes a uniquely named
# class. Running a NEW version in the same session does not reuse the old loaded class, and
# running the identical version again does not recompile (no TYPE_ALREADY_EXISTS).
$verHash = ([BitConverter]::ToString(
    [System.Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($cs))
)).Replace('-','').Substring(0,12)
$ns = "DiskUsageTool_$verHash"
$cs = $cs.Replace('namespace DiskUsageTool {', "namespace $ns {")
$script:TypeName = "$ns.DiskWalker"
if (-not ([System.Management.Automation.PSTypeName]$script:TypeName).Type) {
    Add-Type -TypeDefinition $cs -Language CSharp -ErrorAction Stop
}

# ============================================================================
#  State and helpers
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

# Populates the children of a node (one level). Folders use the cache; files are measured on demand.
function Build-Children($node) {
    $node.Nodes.Clear()
    $path = $node.Name
    $parentSize = $script:Walker.DirAlloc($path)
    $items = @()
    try {
        foreach ($d in (Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction SilentlyContinue)) {
            if (Test-Reparse $d) { continue }   # junction/symlink: not counted (same as the scan)
            $items += [pscustomobject]@{ Name=$d.Name; Path=$d.FullName; Size=($script:Walker.DirAlloc($d.FullName)); IsDir=$true }
        }
        foreach ($f in (Get-ChildItem -LiteralPath $path -File -Force -ErrorAction SilentlyContinue)) {
            $items += [pscustomobject]@{ Name=$f.Name; Path=$f.FullName; Size=($script:Walker.FileAlloc($f.FullName)); IsDir=$false }
        }
    } catch {}

    $sorted  = @($items | Sort-Object Size -Descending)
    $maxSize = if ($sorted.Count -gt 0) { [long]$sorted[0].Size } else { 0 }
    $shown = 0
    foreach ($c in $sorted) {
        if ($shown -ge $MAX_CHILDREN) {
            [void]$node.Nodes.Add((New-Object System.Windows.Forms.TreeNode("... (+$($sorted.Count - $shown) items not shown)")))
            break
        }
        $pct = if ($parentSize -gt 0) { 100.0 * $c.Size / $parentSize } else { 0 }
        $icon = if ($c.IsDir) { '[DIR]' } else { '     ' }
        $tn = New-Object System.Windows.Forms.TreeNode
        $tn.Text = ('{0} {1}  {2}  ({3:N1}%)' -f $icon, $c.Name, (Format-Size $c.Size), $pct)
        $tn.Name = $c.Path     # full path (for Explorer/Copy and re-expansion)
        # Pastel gradient relative to the largest sibling: red = largest, green = smallest.
        # Minimum channel 200 keeps the background light so black text stays readable.
        $frac = if ($maxSize -gt 0) { [double]$c.Size / $maxSize } else { 0 }
        $r = [int][math]::Min(255, 200 + 110 * $frac)
        $g = [int][math]::Min(255, 200 + 110 * (1 - $frac))
        $tn.BackColor = [System.Drawing.Color]::FromArgb($r, $g, 200)
        if ($c.IsDir) { [void]$tn.Nodes.Add((New-Object System.Windows.Forms.TreeNode('...'))) }
        [void]$node.Nodes.Add($tn)
        $shown++
    }
}

# ============================================================================
#  GUI
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Show-DiskUsage - Admin'
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
$lbl.Text = 'Ready. Enter a path (e.g. C:\) and click Scan. The scan can take 1-3 min.'
[void]$status.Items.Add($lbl)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Dock = 'Fill'
$tree.Font = New-Object System.Drawing.Font('Consolas',10)
$tree.HideSelection = $false

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = $menu.Items.Add('Open in Explorer')
$miCopy = $menu.Items.Add('Copy path')
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

$tree.Add_BeforeExpand({
    param($s,$e)
    $n = $e.Node
    if ($n.Nodes.Count -eq 1 -and $n.Nodes[0].Text -eq '...') { Build-Children $n }
})

$btn.Add_Click({
    $root = $txt.Text
    if (-not (Test-Path -LiteralPath $root)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid path.","Error") | Out-Null; return
    }
    $tree.Nodes.Clear()
    $form.Cursor = 'WaitCursor'
    $script:Walker = New-Object $script:TypeName
    $rootKey = $root.TrimEnd('\')
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $total = 0L

    # Walk the top-level folders one at a time, only to show progress in the status bar
    $topDirs = @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)
    foreach ($d in $topDirs) {
        if (Test-Reparse $d) { continue }
        $lbl.Text = "Scanning $($d.FullName) ...  ($($script:Walker.FileCount) files so far)"
        [System.Windows.Forms.Application]::DoEvents()
        $total += $script:Walker.ScanTree($d.FullName)
    }
    # Files directly in the root
    foreach ($f in (Get-ChildItem -LiteralPath $root -File -Force -ErrorAction SilentlyContinue)) {
        $total += $script:Walker.FileAlloc($f.FullName)
    }
    $script:Walker.DirSize[$rootKey] = $total
    $sw.Stop()

    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Name = $root   # full path WITH backslash ("C:\"); "C:" without it = current directory of the drive
    $rootNode.Text = ('[DIR] {0}  {1}  (100%)' -f $root, (Format-Size $total))
    [void]$rootNode.Nodes.Add((New-Object System.Windows.Forms.TreeNode('...')))
    [void]$tree.Nodes.Add($rootNode)
    Build-Children $rootNode
    $rootNode.Expand()
    $form.Cursor = 'Default'
    $lbl.Text = "OK: $($script:Walker.FileCount) files in $([math]::Round($sw.Elapsed.TotalSeconds,1))s  |  $root (physical) = $(Format-Size $total)"
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
            'Show-DiskUsage - FATAL ERROR', 'OK', 'Error') | Out-Null
    } catch {
        Write-Host $msg -ForegroundColor Red
        Read-Host 'Press Enter to exit'
    }
}
