<#
.SYNOPSIS
  Turns [[Page Title]] markers in OneNote pages into links to the page with that title.

.DESCRIPTION
  Uses OneNote desktop's local COM interface (OneNote for Microsoft 365 on Windows).
  No Azure app registration, tokens, or network calls are involved.

  Supported markers:
    [[Linux]]                  -> link to the page titled "Linux"
    [[Linux|the Linux kernel]] -> same link, displayed as "the Linux kernel"

  Dry-run by default: nothing is written unless -Apply is given.

.PARAMETER Apply       Actually write changes to OneNote. Without it, only reports.
.PARAMETER Notebook    Only process pages in notebooks whose name matches (wildcards ok).
.PARAMETER Section     Only process pages in sections whose name matches (wildcards ok).
                       Link targets are always looked up across ALL notebooks.
.PARAMETER BackupDir   Folder for pre-change page XML copies. Default: .\onenote-backup
.PARAMETER NoBackup    Skip local backups (OneNote's own page history still applies).

.EXAMPLE
  .\Update-WikiLinks.ps1 -Section "Scratch"            # dry run on one section
  .\Update-WikiLinks.ps1 -Section "Scratch" -Apply     # write changes
#>
[CmdletBinding()]
param(
    [switch]$Apply,
    [string]$Notebook = '*',
    [string]$Section = '*',
    [string]$BackupDir = (Join-Path (Get-Location) 'onenote-backup'),
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
$ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
$hsPages = 4   # HierarchyScope.hsPages
$piBasic = 0   # PageInfo.piBasic

try { $on = New-Object -ComObject OneNote.Application }
catch { throw "Could not start OneNote's COM interface. Install/open the desktop OneNote (Microsoft 365), not the 'OneNote for Windows 10' app." }

# --- Build page index -------------------------------------------------------
$h = ''
$on.GetHierarchy('', $hsPages, [ref]$h)
$hier = New-Object System.Xml.XmlDocument
$hier.LoadXml($h)
$nsm = New-Object System.Xml.XmlNamespaceManager($hier.NameTable)
$nsm.AddNamespace('one', $ns)

$pages = foreach ($p in $hier.SelectNodes('//one:Page', $nsm)) {
    [pscustomobject]@{
        Id       = $p.GetAttribute('ID')
        Title    = $p.GetAttribute('name')
        Section  = $p.ParentNode.GetAttribute('name')
        Notebook = $p.ParentNode.ParentNode.GetAttribute('name')
    }
}
# Sections nested in section groups: walk up to the notebook element.
foreach ($pg in $pages) {
    $n = $hier.SelectSingleNode("//one:Page[@ID='$($pg.Id)']", $nsm)
    while ($n -and $n.LocalName -ne 'Notebook') { $n = $n.ParentNode }
    if ($n) { $pg.Notebook = $n.GetAttribute('name') }
}

$byTitle = @{}
foreach ($pg in $pages) {
    $k = $pg.Title.Trim().ToLowerInvariant()
    if (-not $byTitle.ContainsKey($k)) { $byTitle[$k] = New-Object System.Collections.ArrayList }
    [void]$byTitle[$k].Add($pg)
}
Write-Host ("Indexed {0} pages." -f @($pages).Count)

$linkCache = @{}
function Get-PageLink($id) {
    if (-not $linkCache.ContainsKey($id)) {
        $l = ''
        $on.GetHyperlinkToObject($id, '', [ref]$l)
        $linkCache[$id] = $l
    }
    $linkCache[$id]
}

# Matches an existing <a>...</a> (left alone) OR a [[marker]].
$rx = [regex]'(?s)(<a\b[^>]*>.*?</a>)|\[\[([^\[\]|<>]+?)(?:\|([^\[\]<>]+?))?\]\]'

$unresolved = New-Object System.Collections.ArrayList
$totalLinks = 0
$changedPages = 0

$targets = $pages | Where-Object { $_.Notebook -like $Notebook -and $_.Section -like $Section }
foreach ($pg in $targets) {
    $x = ''
    $on.GetPageContent($pg.Id, [ref]$x, $piBasic)
    if ($x -notmatch '\[\[') { continue }

    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($x)
    $m = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $m.AddNamespace('one', $ns)

    $pageLinks = 0
    foreach ($t in $doc.SelectNodes('//one:T', $m)) {
        $html = $t.InnerText
        if ($html -notmatch '\[\[') { continue }

        $new = $rx.Replace($html, {
            param($mt)
            if ($mt.Groups[1].Success) { return $mt.Value }   # existing link: untouched
            $name = [System.Net.WebUtility]::HtmlDecode($mt.Groups[2].Value).Trim()
            $shown = if ($mt.Groups[3].Success) { $mt.Groups[3].Value.Trim() } else { $mt.Groups[2].Value.Trim() }
            $hits = $byTitle[$name.ToLowerInvariant()]
            if (-not $hits -or $hits.Count -eq 0) {
                [void]$unresolved.Add("'$($pg.Title)': [[${name}]] - no page with that title")
                return $mt.Value
            }
            if ($hits.Count -gt 1) {
                $where = ($hits | ForEach-Object { "$($_.Notebook)/$($_.Section)" }) -join '; '
                [void]$unresolved.Add("'$($pg.Title)': [[${name}]] - ambiguous ($where)")
                return $mt.Value
            }
            if ($hits[0].Id -eq $pg.Id) { return $mt.Value }  # no self-links
            $href = [System.Net.WebUtility]::HtmlEncode((Get-PageLink $hits[0].Id))
            $script:pageLinks++
            "<a href=`"$href`">$shown</a>"
        })

        if ($new -ne $html) {
            $t.RemoveAll()
            [void]$t.AppendChild($doc.CreateCDataSection($new))
        }
    }

    if ($pageLinks -eq 0) { continue }
    $changedPages++
    $totalLinks += $pageLinks
    Write-Host ("{0}  [{1} / {2}]  {3} link(s)" -f $pg.Title, $pg.Notebook, $pg.Section, $pageLinks)

    if ($Apply) {
        if (-not $NoBackup) {
            New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
            $safe = ($pg.Title -replace '[^\w\- ]', '_')
            $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
            [System.IO.File]::WriteAllText((Join-Path $BackupDir "$stamp-$safe.xml"), $x)
        }
        $on.UpdatePageContent($doc.OuterXml)
    }
}

Write-Host ''
Write-Host ("{0} {1} link(s) across {2} page(s)." -f ($(if ($Apply) { 'Wrote' } else { 'Would write (dry run)' })), $totalLinks, $changedPages)
if ($unresolved.Count) {
    Write-Host "`nUnresolved markers (left as-is):"
    $unresolved | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" }
}
if (-not $Apply -and $totalLinks) { Write-Host "`nRe-run with -Apply to write these changes." }
