#requires -Version 5.1
<#
    Chicago Daily Ops Report - daily update  (TEST VERSION: start times from the raw data + trackers)

    Run "Daily Ops.bat" (double-click, or drag today's files onto it).
    Type 'help' at the prompt for every command.

    Does what the workbook buttons did, in one go, through Excel:
      report date + day columns, Raw Machine Data -> YEARLY DATA, start times (raw data, trackers fill the gaps, + typed),
      scheduled headcount, absent per shop -> present, Regular/OT/DT hours (files or .zip),
      cubic-inch pivot filter, latest-day yellow cells, formula check,
      then saves the workbook and/or the 2 pictures (Summary + All shops).

    Files next to this script:
      DailyOpsConfig.xlsx   shops (sheet names, departments, asset areas) + settings (folders, shift starts)

    Start times: for each machine of a shop (Assets sheet area, or the shop's one Machine #) and shift,
      1. the Raw Machine Data: first operation logged in the shift = its start
         (already running across the shift start = 0; up to StartEarlyMin early is looked at)
      2. no raw value, a raw start over StartLateCap, or a NEGATIVE raw start -> the start-time
         tracker's time for that machine, if it has one
      3. a negative raw start with no tracker time = 0; a negative tracker time = 0
    Shift = average of its machines; shop's day value = average of its shifts (one number).
    Everything is also written to Output\Start times from raw data.csv for comparing.
#>
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$CmdArgs
)

$ErrorActionPreference = 'Stop'

$Script:Root       = $PSScriptRoot
$Script:ConfigPath = Join-Path $Script:Root 'DailyOpsConfig.xlsx'
$Script:TypedPath  = Join-Path (Join-Path $Script:Root 'Saved') 'typed.json'
$Script:Inv        = [Globalization.CultureInfo]::InvariantCulture
$Script:US         = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$Script:DayAbbr    = @('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun')
$Script:ShiftNames = @('1st', '2nd', '3rd')
$Script:XlExt      = @('.xlsx', '.xlsm', '.xls', '.xlsb')
$Script:Warn       = New-Object System.Collections.ArrayList

$Script:SUMMARY = 'Summary'
$Script:CUBIC   = 'Machining_query_cubic in'
$Script:YEARLY  = 'Machining_YEARLY DATA'
$Script:ShotSummary = 'A3:U59'
$Script:ShotCubic   = 'K1:U22'
$Script:Origin     = @{}

$Script:DefaultSettings = [ordered]@{
    ReportFolder     = @('J:\Departments\Machine\4M\KACPER', 'Where the Daily Ops Report workbook is (newest .xlsm there is offered first).')
    RawFolder        = @('',              'Where Raw Machine Data downloads land. Blank = your Downloads folder.')
    TrackerFolder    = @('J:\Departments\Machine\4M', 'Where the machine start-time trackers are (they fill in machines the raw data has nothing for).')
    ScheduleFolder   = @('J:\Departments\Machine\SOUTH\23 Attendance\Schedules', 'Where the Hours Required schedules are.')
    HoursFolder      = @('',              'Where Shift Hours and OT Reports (or .zip of them) land. Blank = your Downloads folder.')
    OutputFolder     = @('',              'Where pictures (and copies) are saved. Blank = "Output" folder next to the script. Can be a full path.')
    DaysToLook       = @('14',            'How many days back to look for downloaded files.')
    Shift1Start      = @('6:00',          'Start of 1st shift (start-time minutes are counted from this).')
    Shift2Start      = @('14:00',         'Start of 2nd shift.')
    Shift3Start      = @('22:00',         'Start of 3rd shift.')
    SaveOverOriginal = @('No',            'Yes = save the report workbook in place (a backup is kept). No = save a copy in the output folder (default here, so this test version doesn''t touch the real report).')
    KeepBackups      = @('5',             'How many backups of the report workbook to keep in Backups\ (0 = none).')
    FixFormulas      = @('Yes',           'Put back formulas that were typed over / broken (#REF!) in the Summary rows and the Core rows they use (Yes/No).')
    LatestDayCell    = @('Yes',           'Yellow cell right of each Performance % row uses the latest report day instead of always Friday (Yes/No).')
    CallInScan       = @('No',            'Yes = the absent step asks whether to scan Outlook for call-in voicemail MP3s. No = never asks.')
    OpenWhenDone     = @('Yes',           'Open the workbook + the pictures folder when done (Yes/No).')
    StartShifts      = @('1,2,3',         'Shifts that count for the start-time average (raw data).')
    StartEarlyMin    = @('30',            'Raw starts up to this many minutes BEFORE the shift start are looked at; negative = tracker time for that machine, else 0.')
    StartLateCap     = @('120',           'Raw first start later than this many minutes is not used (late logging / idle): tracker time if it has the machine, else left out.')
    OutlookSubject   = @('Shift Hours and OT Report', 'Hours step also looks in Outlook (Inbox + its subfolders) for emails with this in the subject and takes their attachments. Blank = off.')
    PageFolder       = @('',              'Folder on J: for the Daily Ops web page (index.html + the 2 pictures). Blank = off.')
}

$Script:DefaultShops = @(
    [pscustomobject]@{ Code = 'HMS';    Label = 'HMS';     Sheet = 'Core Machining_HMS';     TrackerTab = 'HMS';     Dept = 153; Machine = 0;     SchedOnly = '';    SchedExcept = '';    FileWords = 'HMS';         Areas = 'HMS' },
    [pscustomobject]@{ Code = 'MMS';    Label = 'MMS';     Sheet = 'Core Machining_MMS';     TrackerTab = 'MMS';     Dept = 154; Machine = 0;     SchedOnly = '';    SchedExcept = '';    FileWords = 'MMS';         Areas = 'MMS' },
    [pscustomobject]@{ Code = 'LMSSMS'; Label = 'LMS-SMS'; Sheet = 'Core Machining_LMS-SMS'; TrackerTab = 'SMS_FEL'; Dept = 145; Machine = 0;     SchedOnly = '';    SchedExcept = '175'; FileWords = 'SMS,LMS,FEL'; Areas = 'LMS,SMS' },
    [pscustomobject]@{ Code = 'FEL';    Label = 'FEL';     Sheet = 'Core Machining_FEL';     TrackerTab = 'SMS_FEL'; Dept = 175; Machine = 10177; SchedOnly = '175'; SchedExcept = '';    FileWords = 'SMS,LMS,FEL'; Areas = 'FEL' }
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

#region ---------- console helpers ----------



function Say([string]$Text = '', [string]$Color = 'Gray') { Write-Host $Text -ForegroundColor $Color }
function Head([string]$Text) { Write-Host ''; Write-Host $Text -ForegroundColor Cyan }
function Warn([string]$Text) { [void]$Script:Warn.Add($Text); Say "  ! $Text" Yellow }

function Start-Dots([string]$Text) {
    Write-Host "  $Text " -ForegroundColor DarkGray -NoNewline
    $flag = [hashtable]::Synchronized(@{ Run = $true })
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({
        param($f)
        $n = 0
        while ($f.Run) {
            Start-Sleep -Milliseconds 100
            if ((++$n % 10) -eq 0 -and $f.Run) { $c = [Console]::ForegroundColor; [Console]::ForegroundColor = 'DarkGray'; [Console]::Write('.'); [Console]::ForegroundColor = $c }
        }
    }).AddArgument($flag)
    return [pscustomobject]@{ Ps = $ps; H = $ps.BeginInvoke(); Flag = $flag; Clock = [Diagnostics.Stopwatch]::StartNew() }
}

function Stop-Dots($Dots, [string]$Done = 'done') {
    if ($null -eq $Dots) { return }
    $Dots.Flag.Run = $false
    try { [void]$Dots.Ps.EndInvoke($Dots.H) } catch { }
    $Dots.Ps.Dispose()
    Write-Host (' {0} ({1:0} s)' -f $Done, $Dots.Clock.Elapsed.TotalSeconds) -ForegroundColor DarkGray
}

function Ask([string]$Prompt, [string]$Default = '') {
    $p = $Prompt
    if ($Default -ne '') { $p += " [$Default]" }
    $a = Read-Host $p
    if ($null -eq $a) { $a = '' }
    $a = $a.Trim()
    if ($a -ieq 'cancel') { throw (New-Object OperationCanceledException) }
    if ($a -eq '') { return $Default }
    return $a
}

function AskYesNo([string]$Prompt, [bool]$Default = $false) {
    $d = 'y/N'
    if ($Default) { $d = 'Y/n' }
    while ($true) {
        $a = (Ask "$Prompt ($d)").ToLower()
        if ($a -eq '') { return $Default }
        if ($a -eq 'y' -or $a -eq 'yes') { return $true }
        if ($a -eq 'n' -or $a -eq 'no') { return $false }
    }
}

function AskInt([string]$Prompt, [int]$Min, [int]$Max, [int]$Default) {
    while ($true) {
        $a = Ask $Prompt "$Default"
        $n = 0
        if ([int]::TryParse($a, [ref]$n) -and $n -ge $Min -and $n -le $Max) { return $n }
        Say "  Enter a number from $Min to $Max." Yellow
    }
}

function Clean-PathText([string]$s) {
    $s = "$s".Trim()
    if ($s.StartsWith('& ')) { $s = $s.Substring(2).Trim() }
    return $s.Trim('"').Trim("'").Trim()
}

function Split-PathList([string]$s) {
    $s = "$s".Trim()
    if ($s.StartsWith('& ')) { $s = $s.Substring(2).Trim() }
    $whole = Clean-PathText $s
    if ($whole -ne '' -and (Test-Path -LiteralPath $whole)) { return @($whole) }
    $out = @()
    foreach ($m in [regex]::Matches($s, '"([^"]+)"|''([^'']+)''|(\S+)')) {
        $v = $m.Groups[1].Value + $m.Groups[2].Value + $m.Groups[3].Value
        if ($v -ne '') { $out += $v }
    }
    return $out
}

function Get-SafeFileName([string]$name) {
    $bad = [IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object Text.StringBuilder
    foreach ($ch in $name.ToCharArray()) { if ($bad -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) } }
    $r = $sb.ToString().Trim().TrimEnd('.')
    if ($r -eq '') { $r = 'file' }
    return $r
}

function Get-UniquePath([string]$Dir, [string]$Name) {
    $p = Join-Path $Dir $Name
    $base = [IO.Path]::GetFileNameWithoutExtension($Name)
    $ext = [IO.Path]::GetExtension($Name)
    $i = 2
    while (Test-Path -LiteralPath $p) { $p = Join-Path $Dir ("{0} ({1}){2}" -f $base, $i, $ext); $i++ }
    return $p
}

function Remove-Com($o) {
    if ($null -ne $o) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch { } }
}

function Get-NumberList([string]$Text, [int]$Max) {
    $idx = @()
    foreach ($tok in ($Text -split '[,\s]+')) {
        if ($tok -match '^(\d+)-(\d+)$') { $idx += ([int]$Matches[1])..([int]$Matches[2]) }
        elseif ($tok -match '^\d+$') { $idx += [int]$tok }
        elseif ($tok -ne '') { return $null }
    }
    $idx = @($idx | Where-Object { $_ -ge 1 -and $_ -le $Max } | Select-Object -Unique)
    if ($idx.Count -eq 0) { return $null }
    return $idx
}

function Format-Stamp([datetime]$d) { return $d.ToString('ddd MM/dd h:mm tt', $Script:US) }
function Format-Day([datetime]$d) { return $d.ToString('ddd MM/dd', $Script:US) }

#endregion

#region ---------- minimal .xlsx reader / writer for the config ----------

function Get-ZipText($Zip, [string]$Name) {
    $e = $Zip.GetEntry($Name)
    if ($null -eq $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open())
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Get-ColIndex([string]$Ref) {
    $n = 0
    foreach ($ch in (($Ref -replace '\d', '').ToUpper()).ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n - 1
}

function Get-ColLetter([int]$Index) {
    $s = ''; $i = $Index + 1
    while ($i -gt 0) { $m = ($i - 1) % 26; $s = [string][char](65 + $m) + $s; $i = [math]::Floor(($i - 1) / 26) }
    return $s
}

function Read-Xlsx([string]$Path) {
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $zip = $null
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Read)
        $shared = New-Object 'System.Collections.Generic.List[string]'
        $sst = Get-ZipText $zip 'xl/sharedStrings.xml'
        if ($sst) {
            $x = [xml]$sst
            foreach ($si in $x.DocumentElement.SelectNodes("*[local-name()='si']")) {
                $sb = New-Object Text.StringBuilder
                foreach ($t in $si.SelectNodes(".//*[local-name()='t' and not(ancestor::*[local-name()='rPh'])]")) { [void]$sb.Append($t.InnerText) }
                $shared.Add($sb.ToString())
            }
        }
        $wbx = [xml](Get-ZipText $zip 'xl/workbook.xml')
        $relx = [xml](Get-ZipText $zip 'xl/_rels/workbook.xml.rels')
        $rel = @{}
        foreach ($r in $relx.DocumentElement.SelectNodes("*[local-name()='Relationship']")) { $rel[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }
        $result = [ordered]@{}
        foreach ($s in $wbx.SelectNodes("//*[local-name()='sheet']")) {
            $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = $rel[$rid]
            if ($target.StartsWith('/')) { $target = $target.TrimStart('/') } else { $target = 'xl/' + $target }
            $sx = [xml](Get-ZipText $zip $target)
            $rows = New-Object System.Collections.ArrayList
            foreach ($row in $sx.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
                $cells = @{}; $max = -1; $auto = 0
                foreach ($c in $row.SelectNodes("*[local-name()='c']")) {
                    $ref = $c.GetAttribute('r')
                    if ($ref) { $ci = Get-ColIndex $ref } else { $ci = $auto }
                    $auto = $ci + 1
                    $t = $c.GetAttribute('t')
                    $v = $c.SelectSingleNode("*[local-name()='v']")
                    $val = ''
                    if ($t -eq 's') { if ($null -ne $v) { $val = $shared[[int]$v.InnerText] } }
                    elseif ($t -eq 'inlineStr') { $is = $c.SelectSingleNode("*[local-name()='is']"); if ($null -ne $is) { $val = $is.InnerText } }
                    elseif ($t -eq 'b') { if ($null -ne $v) { if ($v.InnerText -eq '1') { $val = 'TRUE' } else { $val = 'FALSE' } } }
                    elseif ($null -ne $v) { $val = $v.InnerText }
                    $cells[$ci] = $val
                    if ($ci -gt $max) { $max = $ci }
                }
                $arr = New-Object string[] ($max + 1)
                foreach ($k in $cells.Keys) { $arr[$k] = $cells[$k] }
                [void]$rows.Add($arr)
            }
            $result[$s.GetAttribute('name')] = $rows
        }
        return $result
    }
    finally {
        if ($null -ne $zip) { $zip.Dispose() }
        $fs.Dispose()
    }
}

function Get-SheetTable($Book, [string]$Sheet) {
    $out = New-Object System.Collections.ArrayList
    $key = $null
    foreach ($k in $Book.Keys) { if ($k -ieq $Sheet) { $key = $k } }
    if ($null -eq $key) { return , $out }
    $rows = $Book[$key]
    if ($rows.Count -lt 1) { return , $out }
    $hdr = $rows[0]
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $o = @{}; $any = $false
        for ($c = 0; $c -lt $hdr.Count; $c++) {
            $h = "$($hdr[$c])".Trim()
            if ($h -eq '') { continue }
            $v = ''
            if ($c -lt $rows[$i].Count) { $v = "$($rows[$i][$c])".Trim() }
            if ($v -ne '') { $any = $true }
            $o[$h] = $v
        }
        if ($any) { [void]$out.Add($o) }
    }
    return , $out
}

function Esc([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [Security.SecurityElement]::Escape($s)
}

function Write-Xlsx([string]$Path, $Sheets) {
    $nsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $nsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $parts = [ordered]@{}
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $wbSheets = ''
    $wbRels = ''
    for ($si = 0; $si -lt $Sheets.Count; $si++) {
        $sh = $Sheets[$si]
        $n = $si + 1
        $ct += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $wbSheets += "<sheet name=`"$(Esc $sh.Name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
        $wbRels += "<Relationship Id=`"rId$n`" Type=`"$nsRel/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        $cols = $sh.Columns
        $sb = New-Object Text.StringBuilder
        [void]$sb.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"$nsMain`" xmlns:r=`"$nsRel`">")
        $sel = ''
        if ($si -eq 0) { $sel = ' tabSelected="1"' }
        [void]$sb.Append("<sheetViews><sheetView workbookViewId=`"0`"$sel><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/><selection pane=`"bottomLeft`" activeCell=`"A2`" sqref=`"A2`"/></sheetView></sheetViews>")
        [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><cols>')
        for ($c = 0; $c -lt $cols.Count; $c++) { [void]$sb.Append("<col min=`"$($c+1)`" max=`"$($c+1)`" width=`"$($cols[$c].Width)`" customWidth=`"1`"/>") }
        [void]$sb.Append('</cols><sheetData><row r="1">')
        for ($c = 0; $c -lt $cols.Count; $c++) { [void]$sb.Append("<c r=`"$(Get-ColLetter $c)1`" t=`"inlineStr`" s=`"1`"><is><t>$(Esc $cols[$c].Header)</t></is></c>") }
        [void]$sb.Append('</row>')
        $r = 2
        foreach ($row in $sh.Rows) {
            [void]$sb.Append("<row r=`"$r`">")
            for ($c = 0; $c -lt $cols.Count; $c++) {
                $v = $null
                if ($c -lt $row.Count) { $v = $row[$c] }
                if ($null -eq $v -or "$v" -eq '') { continue }
                $ref = "$(Get-ColLetter $c)$r"
                $style = ''
                if ($cols[$c].Wrap) { $style = ' s="2"' }
                $num = 0.0
                if ($cols[$c].Numeric -and [double]::TryParse("$v", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$num)) {
                    [void]$sb.Append("<c r=`"$ref`"$style><v>$($num.ToString($Script:Inv))</v></c>")
                }
                else {
                    [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"$style><is><t xml:space=`"preserve`">$(Esc "$v")</t></is></c>")
                }
            }
            [void]$sb.Append('</row>')
            $r++
        }
        [void]$sb.Append('</sheetData>')
        [void]$sb.Append('<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
        $parts["xl/worksheets/sheet$n.xml"] = $sb.ToString()
    }
    $ct += '</Types>'
    $parts['[Content_Types].xml'] = $ct
    $parts['_rels/.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $stylesRid = $Sheets.Count + 1
    $parts['xl/_rels/workbook.xml.rels'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$wbRels<Relationship Id=`"rId$stylesRid`" Type=`"$nsRel/styles`" Target=`"styles.xml`"/></Relationships>"
    $parts['xl/workbook.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"$nsMain`" xmlns:r=`"$nsRel`"><bookViews><workbookView/></bookViews><sheets>$wbSheets</sheets></workbook>"
    $parts['xl/styles.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><styleSheet xmlns=`"$nsMain`"><fonts count=`"2`"><font><sz val=`"11`"/><name val=`"Calibri`"/><family val=`"2`"/></font><font><b/><sz val=`"11`"/><color rgb=`"FFFFFFFF`"/><name val=`"Calibri`"/><family val=`"2`"/></font></fonts><fills count=`"3`"><fill><patternFill patternType=`"none`"/></fill><fill><patternFill patternType=`"gray125`"/></fill><fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF1F4E78`"/><bgColor indexed=`"64`"/></patternFill></fill></fills><borders count=`"1`"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`"/></cellStyleXfs><cellXfs count=`"3`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`"/><xf numFmtId=`"0`" fontId=`"1`" fillId=`"2`" borderId=`"0`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`"/><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyAlignment=`"1`"><alignment vertical=`"top`" wrapText=`"1`"/></xf></cellXfs><cellStyles count=`"1`"><cellStyle name=`"Normal`" xfId=`"0`" builtinId=`"0`"/></cellStyles></styleSheet>"

    $dir = Split-Path -Parent $Path
    $tmp = Join-Path $dir ('~' + [IO.Path]::GetFileNameWithoutExtension($Path) + '.tmp')
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
    $utf8 = New-Object Text.UTF8Encoding($false)
    $zip = [IO.Compression.ZipFile]::Open($tmp, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($k in $parts.Keys) {
            $e = $zip.CreateEntry($k)
            $w = New-Object IO.StreamWriter($e.Open(), $utf8)
            try { $w.Write($parts[$k]) } finally { $w.Dispose() }
        }
    }
    finally { $zip.Dispose() }
    try { [IO.File]::Copy($tmp, $Path, $true) }
    catch { throw "Can't save $([IO.Path]::GetFileName($Path)) - it's probably open in Excel. Close it and try again." }
    finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}

#endregion

#region ---------- config ----------

function ConvertTo-IntSafe([string]$s, [int]$Fallback = 0) {
    $n = 0.0
    if ([double]::TryParse("$s", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$n)) { return [int]$n }
    return $Fallback
}

function Get-Config {
    if (-not (Test-Path -LiteralPath $Script:ConfigPath)) {
        $cfg = [pscustomobject]@{ Shops = New-Object System.Collections.ArrayList; Settings = [ordered]@{} }
        foreach ($s in $Script:DefaultShops) { [void]$cfg.Shops.Add($s) }
        foreach ($k in $Script:DefaultSettings.Keys) { $cfg.Settings[$k] = $Script:DefaultSettings[$k][0] }
        try { Save-Config $cfg; Say "  Made $([IO.Path]::GetFileName($Script:ConfigPath)) with the default shops and folders." DarkGray } catch { }
        return $cfg
    }
    $book = Read-Xlsx $Script:ConfigPath
    $cfg = [pscustomobject]@{ Shops = New-Object System.Collections.ArrayList; Settings = [ordered]@{} }
    foreach ($r in (Get-SheetTable $book 'Shops')) {
        $code = ("$($r['Code'])".Trim().ToUpper()) -replace '\s', ''
        if ($code -eq '') { continue }
        $label = "$($r['Label'])".Trim(); if ($label -eq '') { $label = $code }
        $sheet = "$($r['Sheet'])".Trim(); if ($sheet -eq '') { $sheet = "Core Machining_$label" }
        [void]$cfg.Shops.Add([pscustomobject]@{
                Code = $code; Label = $label; Sheet = $sheet
                TrackerTab = "$($r['TrackerTab'])".Trim()
                Dept = (ConvertTo-IntSafe $r['Dept']); Machine = (ConvertTo-IntSafe $r['Machine'])
                SchedOnly = "$($r['SchedOnly'])".Trim(); SchedExcept = "$($r['SchedExcept'])".Trim()
                FileWords = "$($r['FileWords'])".Trim()
                Areas = "$($r['Areas'])".Trim()
            })
    }
    if ($cfg.Shops.Count -eq 0) { foreach ($s in $Script:DefaultShops) { [void]$cfg.Shops.Add($s) } }
    foreach ($s in $cfg.Shops) {                                                 # older config without the Areas column
        if (-not $s.Areas) { $d = @($Script:DefaultShops | Where-Object { $_.Code -eq $s.Code }); if ($d.Count -gt 0) { $s.Areas = $d[0].Areas } else { $s.Areas = $s.Label } }
    }
    foreach ($k in $Script:DefaultSettings.Keys) { $cfg.Settings[$k] = $Script:DefaultSettings[$k][0] }
    $inFile = @{}
    foreach ($r in (Get-SheetTable $book 'Settings')) {
        $k = "$($r['Setting'])".Trim()
        if ($k -eq '') { continue }
        $match = $null
        foreach ($dk in @($cfg.Settings.Keys)) { if ($dk -ieq $k) { $match = $dk } }
        if ($null -eq $match) { $match = $k }
        $cfg.Settings[$match] = "$($r['Value'])".Trim()
        $inFile[$match] = $true
    }
    $new = @($Script:DefaultSettings.Keys | Where-Object { -not $inFile.ContainsKey($_) })
    if ($new.Count -gt 0) {                                   # settings added in a newer script -> put them in the config file
        try { Save-Config $cfg; Say "  Added to $([IO.Path]::GetFileName($Script:ConfigPath)): $($new -join ', ')" DarkGray } catch { }
    }
    return $cfg
}

function Save-Config($Cfg) {
    $shopRows = New-Object System.Collections.ArrayList
    foreach ($s in $Cfg.Shops) {
        $m = ''; if ($s.Machine -gt 0) { $m = $s.Machine }
        [void]$shopRows.Add(@($s.Code, $s.Label, $s.Sheet, $s.TrackerTab, $s.Dept, $m, $s.SchedOnly, $s.SchedExcept, $s.FileWords, $s.Areas))
    }
    $setRows = New-Object System.Collections.ArrayList
    foreach ($k in $Cfg.Settings.Keys) {
        $note = ''
        if ($Script:DefaultSettings.Contains($k)) { $note = $Script:DefaultSettings[$k][1] }
        [void]$setRows.Add(@($k, $Cfg.Settings[$k], $note))
    }
    $sheets = @(
        @{ Name = 'Shops'; Rows = $shopRows
            Columns = @(@{ Header = 'Code'; Width = 9 }, @{ Header = 'Label'; Width = 10 }, @{ Header = 'Sheet'; Width = 26 }, @{ Header = 'TrackerTab'; Width = 12 },
                @{ Header = 'Dept'; Width = 7; Numeric = $true }, @{ Header = 'Machine'; Width = 9; Numeric = $true }, @{ Header = 'SchedOnly'; Width = 11 },
                @{ Header = 'SchedExcept'; Width = 12 }, @{ Header = 'FileWords'; Width = 14 }, @{ Header = 'Areas'; Width = 12 })
        },
        @{ Name = 'Settings'; Rows = $setRows
            Columns = @(@{ Header = 'Setting'; Width = 18 }, @{ Header = 'Value'; Width = 52 }, @{ Header = 'Notes'; Width = 90; Wrap = $true })
        }
    )
    Write-Xlsx $Script:ConfigPath $sheets
}

function Get-Setting($Cfg, [string]$Name) {
    if ($Cfg.Settings.Contains($Name)) { return "$($Cfg.Settings[$Name])" }
    if ($Script:DefaultSettings.Contains($Name)) { return $Script:DefaultSettings[$Name][0] }
    return ''
}

function Get-SettingInt($Cfg, [string]$Name, [int]$Fallback) { return (ConvertTo-IntSafe (Get-Setting $Cfg $Name) $Fallback) }
function Test-SettingYes($Cfg, [string]$Name) { return ((Get-Setting $Cfg $Name) -match '^(y|yes|true|1)$') }

function Get-DownloadsFolder {
    try {
        $p = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    catch { }
    if ($env:USERPROFILE) { return (Join-Path $env:USERPROFILE 'Downloads') }
    return (Join-Path $HOME 'Downloads')
}

function Get-Folder($Cfg, [string]$Name) {
    $v = Get-Setting $Cfg $Name
    if ($v -eq '' -or $v -ieq 'downloads') { return (Get-DownloadsFolder) }
    return $v
}

function Get-OutputRoot($Cfg) {
    $o = Get-Setting $Cfg 'OutputFolder'
    if ($o -eq '') { return (Join-Path $Script:Root 'Output') }
    if (-not [IO.Path]::IsPathRooted($o)) { return (Join-Path $Script:Root $o) }
    return $o
}

function Get-ShiftStarts($Cfg) {
    $out = @{}
    foreach ($s in 1..3) {
        $t = Get-Setting $Cfg "Shift${s}Start"
        $m = [regex]::Match($t, '^(\d{1,2})(?::(\d{2}))?')
        if ($m.Success) { $h = [int]$m.Groups[1].Value; $mi = 0; if ($m.Groups[2].Success) { $mi = [int]$m.Groups[2].Value }; $out[$s] = ($h * 60 + $mi) / 1440.0 }
        else { $out[$s] = @(0, 6, 14, 22)[$s] / 24.0 }
    }
    return $out
}

function Find-Shop($Shops, [string]$Text) {
    $t = ("$Text".Trim().ToLower()) -replace '[^a-z0-9]', ''
    if ($t -eq '') { return $null }
    $n = 0
    if ([int]::TryParse($t, [ref]$n) -and $n -ge 1 -and $n -le $Shops.Count) { return $Shops[$n - 1] }
    foreach ($s in $Shops) { if ($s.Code.ToLower() -eq $t -or (($s.Label.ToLower()) -replace '[^a-z0-9]', '') -eq $t) { return $s } }
    $hits = @($Shops | Where-Object { $_.Code.ToLower().StartsWith($t) -or (($_.Label.ToLower()) -replace '[^a-z0-9]', '').StartsWith($t) -or (@($_.Label.ToLower() -split '[^a-z0-9]+') -contains $t) })
    if ($hits.Count -eq 1) { return $hits[0] }
    return $null
}

#endregion

#region ---------- dates ----------

function Get-Serial([datetime]$d) { return [double][math]::Floor($d.ToOADate()) }
function Get-DateOf([double]$Serial) { return [datetime]::FromOADate($Serial) }
function Get-SundayOf([datetime]$d) { $w = (([int]$d.DayOfWeek + 6) % 7) + 1; return $d.Date.AddDays(7 - $w) }

function Get-WeekNum1([datetime]$d) {
    $jan1 = New-Object DateTime($d.Year, 1, 1)
    return [int][math]::Floor(($d.DayOfYear - 1 + [int]$jan1.DayOfWeek) / 7) + 1
}

function Get-ReportWeek([datetime]$ReportDate) {
    $R = $ReportDate.Date
    $last = $R.AddDays(-1)
    $sun = Get-SundayOf $last
    $mon = $sun.AddDays(-6)
    $days = 0
    for ($i = 0; $i -lt 7; $i++) { if ($mon.AddDays($i) -le $last) { $days = $i + 1 } }
    return [pscustomobject]@{ Report = $R; LastDay = $last; WeekEnd = $sun; Monday = $mon; Days = $days }
}


function ConvertTo-DateNear([string]$Text) {
    $t = $Text.Trim() -replace '\.', '/'
    $today = (Get-Date).Date
    if ($t -ieq 't' -or $t -ieq 'today') { return $today }
    if ($t -ieq 'y' -or $t -ieq 'yesterday') { return $today.AddDays(-1) }
    $m = [regex]::Match($t, '^(\d{1,2})[/-](\d{1,2})$')
    if ($m.Success) {
        $mo = [int]$m.Groups[1].Value; $da = [int]$m.Groups[2].Value
        $best = $null
        foreach ($y in @(($today.Year - 1), $today.Year, ($today.Year + 1))) {
            try { $d = New-Object DateTime($y, $mo, $da) } catch { continue }
            if ($null -eq $best -or [math]::Abs(($d - $today).TotalDays) -lt [math]::Abs(($best - $today).TotalDays)) { $best = $d }
        }
        return $best
    }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($t, $Script:US, [Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.Date }
    return $null
}
function Get-OutlookCallInAudio($Cfg) {
    $today = (Get-Date).Date
    $defaultRange = "{0} - {1}" -f $today.AddDays(-1).ToString('MM/dd', $Script:US), $today.ToString('MM/dd', $Script:US)
    
    $inputDays = Ask "Enter days to scan for call-ins (e.g. 10/16 - 10/19 or 10/16, 10/17)" $defaultRange
    
    $targetDates = New-Object System.Collections.Generic.List[datetime]
    if ($inputDays -match '^\s*(\d{1,2}/\d{1,2})\s*-\s*(\d{1,2}/\d{1,2})\s*$') {
        $startD = ConvertTo-DateNear $Matches[1]
        $endD = ConvertTo-DateNear $Matches[2]
        if ($null -ne $startD -and $null -ne $endD) {
            for ($d = $startD; $d -le $endD; $d = $d.AddDays(1)) { [void]$targetDates.Add($d) }
        }
    } else {
        foreach ($tok in ($inputDays -split '[,\s]+')) {
            if ($tok -ne '') {
                $dt = ConvertTo-DateNear $tok
                if ($null -ne $dt) { [void]$targetDates.Add($dt) }
            }
        }
    }
    if ($targetDates.Count -eq 0) {
        Say "  No valid dates parsed. Scanning default (today + yesterday)..." Yellow
        [void]$targetDates.Add($today.AddDays(-1))
        [void]$targetDates.Add($today)
    }

    # Load employee directory if it exists
    $empMapPath = Join-Path (Join-Path $Script:Root 'Saved') 'employees.json'
    $empMap = @{}
    if (Test-Path -LiteralPath $empMapPath) {
        try {
            $json = Get-Content -LiteralPath $empMapPath -Raw | ConvertFrom-Json
            foreach ($p in $json.PSObject.Properties) {
                $cleanNum = ("$($p.Name)" -replace '[^\d]', '')
                $empMap[$cleanNum] = "$($p.Value)"
            }
        } catch { Say "  Couldn't read employees.json directory." Yellow }
    }

    Say "  Scanning Outlook for Chicago Call Off voicemails across $($targetDates.Count) day(s)..." DarkGray
    
    $ol = $null
    try {
        $ol = New-Object -ComObject Outlook.Application
    } catch {
        Warn "Outlook is not running or COM automation is unavailable. Open Outlook and try again."
        return
    }

    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    
    $minDate = @($targetDates | Sort-Object)[0]
    $maxDate = @($targetDates | Sort-Object)[-1]
    # Outlook compares these in UTC, so ask for a day extra on each side; the exact days are checked below
    $startDate = $minDate.AddDays(-1).ToString('MM/dd/yyyy', $Script:US) + ' 00:00'
    $endDate = $maxDate.AddDays(1).ToString('MM/dd/yyyy', $Script:US) + ' 23:59'
    $filter = "@SQL=""urn:schemas:httpmail:datereceived"" >= '$startDate' AND ""urn:schemas:httpmail:datereceived"" <= '$endDate'"
    
    $items = $inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)

    $mainCallInRoot = Join-Path $Script:Root "CallIns"
    New-Item -ItemType Directory -Force -Path $mainCallInRoot | Out-Null

    $count = 0
    foreach ($item in $items) {
        $subj = "$($item.Subject)"
        if ($subj -match "Shared Voicemail" -or $subj -match "Chicago Call Off" -or $subj -match "voice mail") {
            
            $itemDate = $item.ReceivedTime.Date
            if (-not ($targetDates | Where-Object { $_.Date -eq $itemDate })) { continue }

            $dateFolder = Join-Path $mainCallInRoot $itemDate.ToString("yyyy-MM-dd")
            New-Item -ItemType Directory -Force -Path $dateFolder | Out-Null

            $shopTag = "CallOff"
            if ($subj -match "HMS" -or $item.To -match "DG-DLVM-HMS") { $shopTag = "HMS" }
            elseif ($subj -match "MMS" -or $item.To -match "DG-DLVM-MMS") { $shopTag = "MMS" }
            elseif ($subj -match "SMS" -or $item.To -match "DG-DLVM-SMS") { $shopTag = "SMS" }

            $recTime = $item.ReceivedTime.ToString("yyyy-MM-dd_HHmm")

            # Extract phone number from sender name or subject/body
            $rawSender = "$($item.SenderName)"
            $phoneMatch = [regex]::Match($rawSender, '(\+?1?\d[\d\-\(\)\s]{8,}\d)')
            $phoneNumber = ''
            if ($phoneMatch.Success) {
                $phoneNumber = ($phoneMatch.Value -replace '[^\d]', '')
                if ($phoneNumber.Length -eq 11 -and $phoneNumber.StartsWith('1')) { $phoneNumber = $phoneNumber.Substring(1) }
            }

            # Determine employee prefix name
            $empName = ''
            if ($phoneNumber -ne '' -and $empMap.ContainsKey($phoneNumber)) {
                $empName = Get-SafeFileName $empMap[$phoneNumber]
            } else {
                $empName = Get-SafeFileName $rawSender
                if ($empName -eq 'file' -or $empName -eq '') { $empName = 'UnknownCaller' }
            }

            foreach ($att in $item.Attachments) {
                if ($att.FileName -like "*.mp3" -or $att.FileName -like "*.wav") {
                    $saveName = "${recTime}_${empName}_${shopTag}_$($att.FileName)"
                    $savePath = Get-UniquePath $dateFolder $saveName
                    
                    $att.SaveAsFile($savePath)
                    Say "  + Saved to $($itemDate.ToString('yyyy-MM-dd'))\: $([IO.Path]::GetFileName($savePath))" Green
                    $count++
                }
            }
        }
        Remove-Com $item
    }
    Remove-Com $items
    Remove-Com $inbox
    Remove-Com $ns
    Remove-Com $ol

    if ($count -gt 0) {
        Say "  Successfully extracted $count voicemail MP3(s) organized under: $mainCallInRoot" Green
    } else {
        Say "  No matching call-off voicemail emails found for the selected dates." DarkGray
    }
}
function Get-DateFromName([string]$Name) {
    $m = [regex]::Match([IO.Path]::GetFileName($Name), '(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})')
    if (-not $m.Success) { return $null }
    $y = [int]$m.Groups[3].Value; if ($y -lt 100) { $y += 2000 }
    try { return (New-Object DateTime($y, [int]$m.Groups[1].Value, [int]$m.Groups[2].Value)) } catch { return $null }
}

function Get-DayIndex([string]$Tok) {
    if ($Tok.Length -lt 2) { return -1 }
    $t = $Tok.Substring(0, [math]::Min(3, $Tok.Length)).ToLower()
    for ($i = 0; $i -lt 7; $i++) { if ($Script:DayAbbr[$i].ToLower().StartsWith($t)) { return $i } }
    return -1
}

#endregion
#region ---------- start times from the raw machine data ----------

$Script:RawTimeFormats = [string[]]@('M/d/yyyy h:mm:ss tt', 'M/d/yyyy h:mm tt', 'M/d/yyyy H:mm:ss', 'M/d/yyyy H:mm', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd H:mm')

# Start / End Time cell -> [datetime] (text like '9/8/2026 9:20:00 AM' or an Excel date number)
function ConvertTo-RawTime($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [psobject]) { $v = $v.psobject.BaseObject }
    if ($v -is [datetime]) { return $v }
    if ($v -is [double]) { if ($v -gt 1) { return [datetime]::FromOADate($v) }; return $null }
    $t = "$v".Trim()
    if ($t -eq '') { return $null }
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($t, $Script:RawTimeFormats, $Script:US, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$d)) { return $d }
    if ([datetime]::TryParse($t, $Script:US, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

# machine number as text: 10307, '10307', 10307.0 -> '10307'
function ConvertTo-MachineKey($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [psobject]) { $v = $v.psobject.BaseObject }
    if ($v -is [double] -or $v -is [int] -or $v -is [long]) { $x = [double]$v; if ($x -eq [math]::Floor($x)) { return ([long]$x).ToString($Script:Inv) }; return $x.ToString($Script:Inv) }
    $t = "$v".Trim(); $n = 0.0
    if ([double]::TryParse($t, [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$n) -and $n -eq [math]::Floor($n)) { return ([long]$n).ToString($Script:Inv) }
    return $t
}

function ConvertTo-Array1($v) {
    if ($v -is [array]) { return , $v }
    $a = [Array]::CreateInstance([object], @(1, 1), @(1, 1)); $a[1, 1] = $v
    return , $a
}

# columns D (Tool), K (Start Time), L (End Time) of a raw-data sheet from FirstRow down
function Read-RawColumns($Ws, [int]$FirstRow) {
    $last = $Ws.Cells.Item($Ws.Rows.Count, 11).End(-4162).Row                   # last Start Time (xlUp)
    $n = $last - $FirstRow + 1
    if ($n -lt 1) { return [pscustomobject]@{ N = 0; Tool = $null; Start = $null; End = $null } }
    return [pscustomobject]@{
        N     = $n
        Tool  = (ConvertTo-Array1 $Ws.Range("D${FirstRow}:D$last").Value2)
        Start = (ConvertTo-Array1 $Ws.Range("K${FirstRow}:K$last").Value2)
        End   = (ConvertTo-Array1 $Ws.Range("L${FirstRow}:L$last").Value2)
    }
}

# -> hashtable machine -> list of @(start, end) for operations touching [From, To)
function Get-RawOps($Cols, [datetime]$From, [datetime]$To) {
    $ops = @{}
    if ($null -eq $Cols -or $Cols.N -lt 1) { return $ops }
    $t0 = $Cols.Tool.GetLowerBound(0); $s0 = $Cols.Start.GetLowerBound(0); $e0 = $Cols.End.GetLowerBound(0)
    $tc = $Cols.Tool.GetLowerBound(1); $sc = $Cols.Start.GetLowerBound(1); $ec = $Cols.End.GetLowerBound(1)
    for ($i = 0; $i -lt $Cols.N; $i++) {
        $st = ConvertTo-RawTime $Cols.Start[($s0 + $i), $sc]
        if ($null -eq $st -or $st -ge $To) { continue }
        $en = ConvertTo-RawTime $Cols.End[($e0 + $i), $ec]
        if ($null -eq $en) { if ($st -lt $From) { continue } }
        elseif ($en -lt $From) { continue }
        $k = ConvertTo-MachineKey $Cols.Tool[($t0 + $i), $tc]
        if ($k -eq '') { continue }
        if (-not $ops.ContainsKey($k)) { $ops[$k] = New-Object System.Collections.ArrayList }
        [void]$ops[$k].Add(@($st, $en))
    }
    return $ops
}

# Assets sheet of the report workbook: A = machine #, C = Area/Department
function Get-AssetAreas($G) {
    $areas = @{}
    for ($r = $G.R1; $r -le (Get-LastRow $G); $r++) {
        $k = ConvertTo-MachineKey (GridV $G $r 1); $a = "$(GridV $G $r 3)".Trim()
        if ($k -ne '' -and $a -ne '' -and -not $areas.ContainsKey($k)) { $areas[$k] = $a }
    }
    return $areas
}

# the shop's machines: its one Machine # if set (FEL = 10177), else every asset whose area is in Areas
function Get-ShopMachines($Shop, $Areas) {
    if ($Shop.Machine -gt 0) { return @("$($Shop.Machine)") }
    $want = @("$($Shop.Areas)" -split '[,;]+' | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ -ne '' })
    return @($Areas.Keys | Where-Object { $want -contains $Areas[$_].ToUpper() } | Sort-Object)
}

# Does a machine on a tracker tab count for this shop? Its own Machine # if it has one; on a tab two shops
# share (SMS_FEL) the Assets area decides; on a shop's own tab every machine listed counts.
function Test-TrackerMachine($Shop, [string]$Machine, $Shops, $Areas) {
    if ($Shop.Machine -gt 0) { return ($Machine -eq "$($Shop.Machine)") }
    $sharing = @($Shops | Where-Object { $_.TrackerTab -ieq $Shop.TrackerTab -and $_.Code -ne $Shop.Code })
    if (@($sharing | Where-Object { $_.Machine -gt 0 -and "$($_.Machine)" -eq $Machine }).Count -gt 0) { return $false }
    if ($sharing.Count -eq 0) { return $true }
    if (-not $Areas.ContainsKey($Machine)) { return $true }
    $want = @("$($Shop.Areas)" -split '[,;]+' | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ -ne '' })
    return ($want -contains $Areas[$Machine].ToUpper())
}

# Per-machine times from the start-time trackers -> @{ "day|shift|machine" = minutes } for one shop
# (minutes after the shift start; a 3rd-shift time after midnight counts for the evening before)
function Get-TrackerMachineTimes($Trackers, $Shop, $Shops, $Areas, [datetime]$Monday, $ShiftStarts) {
    $sum = @{}; $cnt = @{}; $mon = Get-Serial $Monday
    foreach ($f in $Trackers) {
        if ($f.Shops -notcontains $Shop.Code) { continue }
        $sh = Get-BookSheet $f.Book $Shop.TrackerTab
        if ($null -eq $sh) { continue }
        $g = Get-SheetGrid $sh
        $avgRow = Find-InCol $g 1 1 (Get-LastRow $g) 'AVG Start time'
        if ($avgRow -eq 0) { continue }
        $dateRow = 0
        for ($r = $avgRow + 1; $r -le $avgRow + 10; $r++) { if (Test-DateValue (GridV $g $r 2)) { $dateRow = $r; break } }
        if ($dateRow -eq 0) { continue }
        $startT = $ShiftStarts[$f.Shift]
        $cols = @(); for ($d = 0; $d -lt 7; $d++) { $cols += (Match-InRow $g $dateRow ($mon + $d)) }
        for ($r = $g.R1; $r -lt $avgRow; $r++) {
            $m = ConvertTo-MachineKey (GridV $g $r 1)
            if ($m -eq '' -or $m -notmatch '^\d+$') { continue }
            if (-not (Test-TrackerMachine $Shop $m $Shops $Areas)) { continue }
            for ($d = 0; $d -lt 7; $d++) {
                if ($cols[$d] -eq 0) { continue }
                $v = GridV $g $r $cols[$d]
                if ($v -isnot [double]) { continue }
                $mins = ($v - [math]::Floor($v) - $startT) * 1440
                if ($mins -lt -720) { $mins += 1440 }
                $k = "$d|$($f.Shift)|$m"
                if ($sum.ContainsKey($k)) { $sum[$k] += $mins; $cnt[$k]++ } else { $sum[$k] = $mins; $cnt[$k] = 1 }
            }
        }
    }
    $out = @{}; foreach ($k in $sum.Keys) { $out[$k] = $sum[$k] / $cnt[$k] }
    return $out
}

# Raw data first, the tracker fills the holes, negatives -> 0. One number per day.
#   per machine & shift:  raw >= 0 (or already running = 0)          -> raw
#                         raw negative                                -> tracker if it has the machine, else 0
#                         no raw / raw over the cap                   -> tracker if it has the machine (else nothing / left out)
#                         tracker negative                            -> 0
#   shift = average of its machines; day = average of the shifts
# -> @{ Values = 7 x (minutes | $null); Shifts = per day/shift summary; Machines = per machine rows }
function Get-RawStartTimes($Ops, $Machines, $Week, $ShiftStarts, [int[]]$Shifts, [int]$Early, [int]$Cap, $Trk) {
    if ($null -eq $Trk) { $Trk = @{} }
    $all = @($Machines)
    foreach ($k in $Trk.Keys) { $m = ($k -split '\|')[2]; if ($all -notcontains $m) { $all += $m } }
    $shiftList = @($Shifts)
    foreach ($k in $Trk.Keys) { $x = [int](($k -split '\|')[1]); if ($shiftList -notcontains $x) { $shiftList += $x } }
    $shiftList = @($shiftList | Sort-Object)
    $values = @(); $sum = New-Object System.Collections.ArrayList; $rows = New-Object System.Collections.ArrayList
    for ($d = 0; $d -lt 7; $d++) {
        if ($d -ge $Week.Days) { $values += $null; continue }
        $day = $Week.Monday.AddDays($d)
        $avgs = @()
        foreach ($sh in $shiftList) {
            $ws = $day.AddMinutes([math]::Round($ShiftStarts[$sh] * 1440))      # 3rd shift = the evening of this day
            $we = $ws.AddHours(8); $lo = $ws.AddMinutes(-$Early)
            $mins = @(); $late = 0; $running = 0; $nRaw = 0; $nTrk = 0
            foreach ($m in $all) {
                $rawV = $null; $rawNote = ''
                if ($Shifts -contains $sh -and $Ops.ContainsKey($m)) {
                    $isRun = $false; $first = $null
                    foreach ($o in $Ops[$m]) {
                        $st = $o[0]; $en = $o[1]
                        if ($st -lt $ws -and $null -ne $en -and $en -gt $ws) { $isRun = $true; break }
                        if ($st -ge $lo -and $st -lt $we -and ($null -eq $first -or $st -lt $first)) { $first = $st }
                    }
                    if ($isRun) { $rawV = 0.0; $rawNote = 'running' }
                    elseif ($null -ne $first) { $rawV = ($first - $ws).TotalMinutes }
                }
                $tk = "$d|$sh|$m"; $trkV = $null; if ($Trk.ContainsKey($tk)) { $trkV = [double]$Trk[$tk] }
                $use = $null; $src = ''; $note = ''
                if ($null -ne $rawV -and $rawV -ge 0 -and $rawV -le $Cap) { $use = $rawV; $src = 'raw'; if ($rawNote) { $note = 'already running' } }
                elseif ($null -ne $rawV -and $rawV -lt 0) {
                    if ($null -ne $trkV) { $use = $trkV; $src = 'tracker'; $note = "raw was $([int][math]::Round($rawV)) (negative)" }
                    else { $use = 0.0; $src = 'raw'; $note = "negative ($([int][math]::Round($rawV))) -> 0" }
                }
                elseif ($null -ne $trkV) { $use = $trkV; $src = 'tracker'; if ($null -ne $rawV) { $note = "raw was $([int][math]::Round($rawV)) (over $Cap)" } }
                elseif ($null -ne $rawV) { $late++; [void]$rows.Add([pscustomobject]@{ Date = $day; Shift = $sh; Machine = $m; Minutes = [int][math]::Round($rawV); Source = 'raw'; Note = "left out (over $Cap min)" }); continue }
                else { continue }
                if ($use -lt 0) { $n2 = "tracker $([int][math]::Round($use)) (negative) -> 0"; if ($note) { $note = "$note; $n2" } else { $note = $n2 }; $use = 0.0 }
                if ($src -eq 'raw') { $nRaw++ } else { $nTrk++ }
                if ($note -eq 'already running') { $running++ }
                $mins += $use
                [void]$rows.Add([pscustomobject]@{ Date = $day; Shift = $sh; Machine = $m; Minutes = [int][math]::Round($use); Source = $src; Note = $note })
            }
            $avg = $null
            if ($mins.Count -gt 0) { $avg = ($mins | Measure-Object -Sum).Sum / $mins.Count; $avgs += $avg }
            [void]$sum.Add([pscustomobject]@{ Date = $day; Shift = $sh; Machines = $mins.Count; Raw = $nRaw; Tracker = $nTrk; Avg = $avg; LeftOut = $late; Running = $running })
        }
        if ($avgs.Count -gt 0) { $values += [int][math]::Round((($avgs | Measure-Object -Sum).Sum / $avgs.Count), 0) } else { $values += $null }
    }
    return [pscustomobject]@{ Values = $values; Shifts = $sum; Machines = $rows }
}

function Show-RawShiftTable($Shops, $Raw, $Week) {
    $line = '  {0,-9}{1,-8}' -f 'by shift', ''
    for ($d = 0; $d -lt $Week.Days; $d++) { $line += '{0,-13}' -f (Format-Day $Week.Monday.AddDays($d)) }
    Say $line DarkGray
    foreach ($s in $Shops) {
        $r = $Raw[$s.Code]; if ($null -eq $r) { continue }
        $first = $true
        foreach ($sh in @($r.Shifts | ForEach-Object { $_.Shift } | Select-Object -Unique)) {
            $lbl = ''; if ($first) { $lbl = $s.Label }; $first = $false
            $line = '  {0,-9}{1,-8}' -f $lbl, $Script:ShiftNames[$sh - 1]
            for ($d = 0; $d -lt $Week.Days; $d++) {
                $x = @($r.Shifts | Where-Object { $_.Shift -eq $sh -and $_.Date -eq $Week.Monday.AddDays($d) })[0]
                $t = '.'
                if ($x -and ($null -ne $x.Avg -or $x.LeftOut -gt 0)) {
                    $avg = '.'; if ($null -ne $x.Avg) { $avg = [int][math]::Round($x.Avg) }
                    $n = "$($x.Raw)"; if ($x.Tracker -gt 0) { $n += "+$($x.Tracker)t" }
                    if ($x.LeftOut -gt 0) { $n += "|$($x.LeftOut)" }
                    $t = "$avg ($n)"
                }
                $line += '{0,-13}' -f $t
            }
            Say $line DarkGray
        }
    }
    Say "  minutes late (raw machines + Nt from the tracker | left out: raw over the StartLateCap, no tracker)" DarkGray
}

# cumulative comparison log + per-machine detail of this run
function Write-RawStartLog($Cfg, $Shops, $Raw, $Week, $Typed) {
    $root = Get-OutputRoot $Cfg
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $log = Join-Path $root 'Start times from raw data.csv'
    $new = @()
    foreach ($s in $Shops) {
        $r = $Raw[$s.Code]; if ($null -eq $r) { continue }
        foreach ($x in $r.Shifts) {
            $d = [int]($x.Date - $Week.Monday).TotalDays
            $typedV = ''; $k = Get-TypedKey $s $x.Date; if ($Typed.Start.ContainsKey($k)) { $typedV = "$($Typed.Start[$k])" }
            $avg = ''; if ($null -ne $x.Avg) { $avg = [math]::Round($x.Avg, 1).ToString($Script:Inv) }
            $dv = ''; if ($null -ne $r.Values[$d]) { $dv = "$($r.Values[$d])" }
            $new += [pscustomobject]@{ Date = $x.Date.ToString('yyyy-MM-dd'); Day = $Script:DayAbbr[$d]; Shop = $s.Label; Shift = $Script:ShiftNames[$x.Shift - 1]
                Machines = $x.Machines; FromRaw = $x.Raw; FromTracker = $x.Tracker; AvgMinutes = $avg; LeftOut = $x.LeftOut; AlreadyRunning = $x.Running; ShopDayValue = $dv; Typed = $typedV; Updated = (Get-Date).ToString('yyyy-MM-dd HH:mm') }
        }
    }
    try {
        $keys = @{}; foreach ($n in $new) { $keys["$($n.Date)|$($n.Shop)|$($n.Shift)"] = $true }
        $old = @(); if (Test-Path -LiteralPath $log) { $old = @(Import-Csv -LiteralPath $log | Where-Object { -not $keys.ContainsKey("$($_.Date)|$($_.Shop)|$($_.Shift)") }) }
        @($old + $new) | Sort-Object Date, Shop, Shift | Export-Csv -LiteralPath $log -NoTypeInformation -Encoding UTF8
    }
    catch { Warn "Couldn't update $log (open in Excel?): $($_.Exception.Message)" }
    try {
        $dir = Join-Path $root $Week.Report.ToString('yyyy-MM-dd', $Script:Inv)
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $det = Join-Path $dir ("Start times by machine {0}.csv" -f $Week.Report.ToString('yyyy-MM-dd', $Script:Inv))
        $rows = @()
        foreach ($s in $Shops) { $r = $Raw[$s.Code]; if ($null -eq $r) { continue }; foreach ($m in $r.Machines) { $rows += [pscustomobject]@{ Date = $m.Date.ToString('yyyy-MM-dd'); Shop = $s.Label; Shift = $Script:ShiftNames[$m.Shift - 1]; Machine = $m.Machine; Minutes = $m.Minutes; Source = $m.Source; Note = $m.Note } } }
        $rows | Sort-Object Date, Shop, Shift, Machine | Export-Csv -LiteralPath $det -NoTypeInformation -Encoding UTF8
    }
    catch { Warn "Couldn't write the per-machine start times: $($_.Exception.Message)" }
    Say "  Start times logged to $log" DarkGray
}

#endregion

#region ---------- Outlook ----------

# Runs in its own PowerShell (a job) so a stuck Outlook can't hang the script. Saves the Excel attachments
# of matching emails (only ones whose file-name date is in the report week) and stamps each file with the
# time the email arrived.
$Script:OutlookScan = {
    param([string]$Subject, [datetime]$Since, [datetime]$Monday, [datetime]$WeekEnd, [string]$OutDir, $Factory)
    $ErrorActionPreference = 'Stop'
    if ($Factory) { $ol = & $Factory } else { $ol = New-Object -ComObject Outlook.Application }
    $ns = $ol.GetNamespace('MAPI')
    $inbox = $ns.GetDefaultFolder(6)                                          # olFolderInbox
    $folders = New-Object System.Collections.ArrayList
    $walk = { param($f, $depth) [void]$folders.Add($f); if ($depth -gt 0) { foreach ($sub in $f.Folders) { & $walk $sub ($depth - 1) } } }
    & $walk $inbox 2
    $out = @()
    foreach ($f in $folders) {
        $items = $f.Items
        $items.Sort('[ReceivedTime]', $true)                                   # newest first
        $it = $items.GetFirst(); $n = 0
        while ($null -ne $it -and $n -lt 3000) {
            $n++
            $rt = $null; try { $rt = [datetime]$it.ReceivedTime } catch { }
            if ($null -ne $rt -and $rt -lt $Since) { break }
            if ($null -ne $rt -and $it.Class -eq 43 -and "$($it.Subject)" -like "*$Subject*") {      # 43 = mail
                $saved = @(); $skipped = 0; $cnt = $it.Attachments.Count
                for ($i = 1; $i -le $cnt; $i++) {
                    $a = $it.Attachments.Item($i); $fn = "$($a.FileName)"
                    if ($fn -notmatch '(?i)\.xls[xmb]?$') { continue }
                    $m = [regex]::Match($fn, '(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})'); $d = $null
                    if ($m.Success) { $y = [int]$m.Groups[3].Value; if ($y -lt 100) { $y += 2000 }; try { $d = New-Object DateTime($y, [int]$m.Groups[1].Value, [int]$m.Groups[2].Value) } catch { } }
                    if ($null -eq $d -or $d -lt $Monday -or $d -gt $WeekEnd) { $skipped++; continue }
                    $dir = Join-Path $OutDir ($rt.ToString('yyyyMMdd_HHmmss') + "_$i")
                    New-Item -ItemType Directory -Force -Path $dir | Out-Null
                    $p = Join-Path $dir $fn
                    $a.SaveAsFile($p)
                    (Get-Item -LiteralPath $p).LastWriteTime = $rt
                    $saved += $p
                }
                $out += [pscustomobject]@{ Subject = "$($it.Subject)"; Received = $rt; Folder = "$($f.Name)"; Count = $cnt; Files = $saved; Skipped = $skipped }
            }
            $it = $items.GetNext()
        }
    }
    return $out
}

# -> list of saved attachment paths (empty if Outlook can't be reached); tells the user what it found
function Get-OutlookHours($Cfg, $Week, [string]$Work, [int]$Days, [int]$TimeoutSec = 120) {
    $subject = Get-Setting $Cfg 'OutlookSubject'
    if ($subject -eq '') { return @() }
    $dir = Join-Path $Work 'outlook'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Say "  Looking in Outlook for '$subject' emails from the last $Days days..." DarkGray
    $since = (Get-Date).Date.AddDays(-$Days)
    $mails = @(); $err = $null
    $job = Start-Job -ScriptBlock $Script:OutlookScan -ArgumentList $subject, $since, $Week.Monday, $Week.WeekEnd, $dir, $null
    try {
        if (Wait-Job $job -Timeout $TimeoutSec) {
            try { $mails = @(Receive-Job $job -ErrorAction Stop) } catch { $err = $_.Exception.Message }
            if ($job.State -eq 'Failed' -and -not $err) { $err = "$($job.ChildJobs[0].JobStateInfo.Reason.Message)" }
        }
        else { $err = "Outlook didn't answer in $TimeoutSec seconds" }
    }
    finally { Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue }
    if ($err) {
        Say "  Outlook: $err" Yellow
        Say '  (The new Outlook can''t be read by scripts - only classic Outlook. In the new Outlook use' DarkGray
        Say '   Save All Attachments on the email and drop the .zip / files here. Turn this off: set OutlookSubject "")' DarkGray
        return @()
    }
    if ($mails.Count -eq 0) { Say '  No matching emails in Outlook.' DarkGray; return @() }
    $files = @()
    foreach ($m in ($mails | Sort-Object Received)) {
        $extra = ''; if ($m.Skipped -gt 0) { $extra = ", $($m.Skipped) not this week" }
        Say ('  Outlook  {0}  {1}  ({2} attached{3})' -f (Format-Stamp ([datetime]$m.Received)), $m.Subject, @($m.Files).Count, $extra) DarkGray
        foreach ($f in @($m.Files)) { $files += "$f"; $Script:Origin["$f"] = 'Outlook ' + (Format-Stamp ([datetime]$m.Received)) }
    }
    return $files
}

#endregion
#region ---------- sheet grids & automation logic ----------

function New-Grid($R1, $C1, $V) {
    if ($V -isnot [array]) { $a = [Array]::CreateInstance([object], @(1, 1), @(1, 1)); $a[1, 1] = $V; $V = $a }
    return [pscustomobject]@{ R1 = [int]$R1; C1 = [int]$C1; Rows = $V.GetLength(0); Cols = $V.GetLength(1); V = $V }
}

function Get-ComGrid($Ws) {
    $ur = $Ws.UsedRange
    try { return (New-Grid $ur.Row $ur.Column $ur.Value2) } finally { Remove-Com $ur }
}

function GridV($G, [int]$r, [int]$c) {
    $i = $r - $G.R1 + 1; $j = $c - $G.C1 + 1
    if ($i -lt 1 -or $j -lt 1 -or $i -gt $G.Rows -or $j -gt $G.Cols) { return $null }
    return $G.V[$i, $j]
}

function Get-LastRow($G) { return $G.R1 + $G.Rows - 1 }
function Get-LastCol($G) { return $G.C1 + $G.Cols - 1 }
function Test-Num($v) { return ($v -is [double]) }

function Get-CellText($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) { return $v.ToString($Script:Inv) }
    return "$v"
}

function Find-InCol($G, [int]$Col, [int]$R1, [int]$R2, [string]$Label) {
    $want = $Label.ToLower()
    $last = [math]::Min($R2, (Get-LastRow $G))
    for ($r = $R1 + 1; $r -le $last; $r++) {
        $v = GridV $G $r $Col
        if ($null -ne $v -and (Get-CellText $v).ToLower().Contains($want)) { return $r }
    }
    $v = GridV $G $R1 $Col
    if ($null -ne $v -and (Get-CellText $v).ToLower().Contains($want)) { return $R1 }
    return 0
}

function Match-InRow($G, [int]$Row, [double]$Value) {
    for ($c = 1; $c -le (Get-LastCol $G); $c++) { $v = GridV $G $Row $c; if ($v -is [double] -and $v -eq $Value) { return $c } }
    return 0
}

function Match-InCol($G, [int]$Col, [double]$Value, [bool]$Lenient) {
    for ($r = 1; $r -le (Get-LastRow $G); $r++) {
        $v = GridV $G $r $Col
        if ($v -is [double] -and $v -eq $Value) { return $r }
        if ($Lenient -and $v -is [string]) { $n = 0.0; if ($v.Trim() -ne '' -and [double]::TryParse($v.Trim(), [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$n) -and $n -eq $Value) { return $r } }
    }
    return 0
}

function Test-DateValue($v) {
    if ($v -is [double]) { return ($v -ge 36526 -and $v -lt 73051) }
    if ($v -is [string]) { $d = [datetime]::MinValue; return [datetime]::TryParse($v, $Script:US, [Globalization.DateTimeStyles]::None, [ref]$d) }
    return $false
}

function Get-Val($v) {
    if ($v -is [double]) { return $v }
    if ($v -is [string]) { $m = [regex]::Match($v.Trim(), '^[-+]?\d*\.?\d+'); if ($m.Success) { return [double]::Parse($m.Value, $Script:Inv) } }
    return 0.0
}

function Test-HasWord([string]$Name, [string]$Word) { return [regex]::IsMatch($Name, '(^|[^A-Za-z])' + [regex]::Escape($Word) + '([^A-Za-z]|$)', 'IgnoreCase') }
function Get-FileWords($Shop) { return @("$($Shop.FileWords)" -split '[,;\s]+' | Where-Object { $_ -ne '' }) }

function Get-BookSheet($Book, [string]$Name) {
    foreach ($s in $Book.Sheets) { if ($s.Name -ieq $Name) { return $s } }
    return $null
}

function Get-SheetGrid($Sheet) {
    if ($null -eq $Sheet.Grid) { $Sheet.Grid = Get-ComGrid $Sheet.Ws }
    return $Sheet.Grid
}

function Get-TrackerShops($Book, $Shops) {
    $name = [IO.Path]::GetFileName($Book.Name)
    $tabOf = { param($s, $visibleOnly) foreach ($t in $Book.Sheets) { if ($s.TrackerTab -ne '' -and $t.Name -ieq $s.TrackerTab -and (-not $visibleOnly -or -not $t.Hidden)) { return $true } }; return $false }
    $ids = @($Shops | Where-Object { & $tabOf $_ $true } | ForEach-Object { $_.Code })
    $any = @($Shops | Where-Object { & $tabOf $_ $false } | ForEach-Object { $_.Code })
    $byName = @($Shops | Where-Object { $s = $_; @(Get-FileWords $s | Where-Object { Test-HasWord $name $_ }).Count -gt 0 } | ForEach-Object { $_.Code })
    if ($byName.Count -gt 0) {
        $both = @($ids | Where-Object { $byName -contains $_ })
        if ($both.Count -gt 0) { $ids = $both }
        elseif ($ids.Count -eq 0) { $ids = @($byName | Where-Object { $any -contains $_ }) }
    }
    if ($ids.Count -eq 0) { $ids = $any }
    return $ids
}

function Get-ShiftOfName([string]$Name) {
    $n = [IO.Path]::GetFileName($Name).ToLower()
    if ($n.Contains('2nd')) { return 2 }
    if ($n.Contains('3rd')) { return 3 }
    return 1
}

function Get-StartTimes($Trackers, $Shop, [datetime]$Monday, $ShiftStarts) {
    $total = New-Object double[] 7; $cnt = New-Object int[] 7
    $notes = @(); $used = @()
    $mon = Get-Serial $Monday
    foreach ($f in $Trackers) {
        if ($f.Shops -notcontains $Shop.Code) { continue }
        $fn = [IO.Path]::GetFileName($f.Book.Name)
        $sh = Get-BookSheet $f.Book $Shop.TrackerTab
        if ($null -eq $sh) { $notes += "${fn}: no '$($Shop.TrackerTab)' tab"; continue }
        $g = Get-SheetGrid $sh
        $avgRow = Find-InCol $g 1 1 (Get-LastRow $g) 'AVG Start time'
        $dateRow = 0
        if ($avgRow -gt 0) { for ($r = $avgRow + 1; $r -le $avgRow + 10; $r++) { if (Test-DateValue (GridV $g $r 2)) { $dateRow = $r; break } } }
        $readRow = $avgRow
        if ($Shop.Machine -gt 0) { $readRow = Match-InCol $g 1 $Shop.Machine $true }
        if ($readRow -eq 0) { if ($Shop.Machine -gt 0) { $notes += "${fn}: machine $($Shop.Machine) not found" } else { $notes += "${fn}: no 'AVG Start time' row" }; continue }
        if ($dateRow -eq 0) { $notes += "${fn}: no date row under 'AVG Start time'"; continue }
        $startT = $ShiftStarts[$f.Shift]
        $used += $fn
        for ($d = 0; $d -lt 7; $d++) {
            $col = Match-InRow $g $dateRow ($mon + $d)
            if ($col -eq 0) { continue }
            $v = GridV $g $readRow $col
            if ($v -is [double]) {
                $mins = ($v - [math]::Floor($v) - $startT) * 1440
                if ($mins -lt -720) { $mins += 1440 }
                $total[$d] += $mins; $cnt[$d]++
            }
        }
    }
    $values = @()
    for ($d = 0; $d -lt 7; $d++) { if ($cnt[$d] -gt 0) { $values += [math]::Round($total[$d] / $cnt[$d], 0) } else { $values += $null } }
    return [pscustomobject]@{ Values = $values; Counts = $cnt; Notes = $notes; Used = $used }
}

function Get-ScheduleShops($Book, $Shops) {
    $name = [IO.Path]::GetFileName($Book.Name)
    $ids = @($Shops | Where-Object { $s = $_; @(Get-FileWords $s | Where-Object { Test-HasWord $name $_ }).Count -gt 0 } | ForEach-Object { $_.Code })
    if ($ids.Count -eq 0 -and $Book.Sheets.Count -gt 1) {
        $t = $Book.Sheets[1].Name.ToUpper()
        foreach ($s in $Shops) { if (@(Get-FileWords $s | Where-Object { $t.Contains($_.ToUpper()) }).Count -gt 0) { $ids += $s.Code } }
        if ($ids.Count -gt 0) { $first = Find-Shop $Shops $ids[0]; $ids = @($Shops | Where-Object { $_.FileWords -eq $first.FileWords } | ForEach-Object { $_.Code }) }
    }
    return $ids
}

function Test-DeptOK([string]$Dept, $Shop) {
    $k = ",$Dept,"
    if ($Shop.SchedOnly -ne '') { return (("," + ($Shop.SchedOnly -replace '\s', '') + ",").Contains($k)) }
    return (-not (("," + ($Shop.SchedExcept -replace '\s', '') + ",").Contains($k)))
}

function Get-ScheduledHeadcount($Book, $Shop, [datetime]$Monday) {
    $values = @(); $tabUsed = ''
    $mon = Get-Serial $Monday
    for ($d = 0; $d -lt 7; $d++) {
        $n = -1
        foreach ($sh in $Book.Sheets) {
            if ($sh.Name -eq 'Format Template') { continue }
            $g = Get-SheetGrid $sh
            $hdr = Find-InCol $g 1 1 20 'DEPT'
            if ($hdr -eq 0) { continue }
            $col = Match-InRow $g $hdr ($mon + $d)
            if ($col -eq 0) { continue }
            $n = 0
            $end = Get-LastRow $g
            for ($r = $hdr + 1; $r -le $end; $r++) {
                $a = GridV $g $r 1
                $dept = $null
                if ($a -is [double]) { $dept = $a.ToString($Script:Inv) }
                elseif ($a -is [string] -and $a.Trim() -ne '') { $x = 0.0; if ([double]::TryParse($a.Trim(), [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$x)) { $dept = $a.Trim() } }
                if ($null -eq $dept) { continue }
                if (-not (Test-DeptOK $dept $Shop)) { continue }
                $v = GridV $g $r $col
                if ($v -is [string]) { $u = $v.ToUpper(); if ($u.Contains('AM') -or $u.Contains('PM')) { $n++ } }
                elseif ($v -is [double] -and ($v - [math]::Floor($v)) -ne 0 -and $v -lt 1) { $n++ }
            }
            if ($d -eq 0) { $tabUsed = $sh.Name }
            break
        }
        if ($n -ge 0) { $values += $n } else { $values += 'N/A' }
    }
    return [pscustomobject]@{ Values = $values; Tab = $tabUsed }
}

function Read-HoursBook($Book, [string]$Name, $Shops) {
    $out = [pscustomobject]@{ Name = $Name; Date = (Get-DateFromName $Name); Ok = $false; Note = ''; Depts = @{} }
    if ($null -eq $out.Date) { $out.Note = 'no date in the file name'; return $out }
    $sh = Get-BookSheet $Book 'Pivot Table'
    if ($null -eq $sh) { $out.Note = "no 'Pivot Table' tab"; return $out }
    $g = Get-SheetGrid $sh
    $hdr = Find-InCol $g 1 1 20 'Row Labels'
    if ($hdr -eq 0) { $out.Note = "no 'Row Labels' header"; return $out }
    $cReg = 0; $cOT = 0; $cDT = 0
    for ($c = 2; $c -le 10; $c++) {
        $h = (Get-CellText (GridV $g $hdr $c)).ToLower()
        if ($h.Contains('regular')) { $cReg = $c }
        elseif ($h.Contains('overtime')) { $cOT = $c }
        elseif ($h.Contains('double')) { $cDT = $c }
    }
    if ($cReg -eq 0 -or $cOT -eq 0 -or $cDT -eq 0) { $out.Note = 'Regular / Overtime / Double columns not found'; return $out }
    foreach ($s in $Shops) {
        $r = Match-InCol $g 1 $s.Dept $false
        if ($r -gt 0) { $out.Depts[$s.Code] = @{ Reg = (Get-Val (GridV $g $r $cReg)); OT = (Get-Val (GridV $g $r $cOT)); DT = (Get-Val (GridV $g $r $cDT)); Found = $true } }
        else { $out.Depts[$s.Code] = @{ Reg = 0.0; OT = 0.0; DT = 0.0; Found = $false } }
    }
    $out.Ok = $true
    return $out
}

function Get-FileKind([string]$Path) {
    $n = [IO.Path]::GetFileName($Path).ToLower()
    if ($n.EndsWith('.zip')) { return 'hours' }
    if ($Script:XlExt -notcontains [IO.Path]::GetExtension($n)) { return 'unknown' }
    if ($n -match 'raw.?machine.?data') { return 'raw' }
    if ($n -match 'start.?time' -or $n -match 'tracker') { return 'tracker' }
    if ($n -match 'hours.?required' -or $n -match 'schedule') { return 'schedule' }
    if ($n -match 'shift.?hours' -or $n -match 'ot.?report' -or $n -match 'overtime') { return 'hours' }
    if ($n -match 'daily.?ops' -or $n -match 'ops.?report' -or $n.EndsWith('.xlsm')) { return 'main' }
    return 'unknown'
}

function ConvertFrom-StartText([string]$Text, $ShiftStarts) {
    $t = "$Text".Trim()
    if ($t -eq '') { return '' }
    if ($t -match '^(?i)n/?a$') { return 'N/A' }
    $m = [regex]::Match($t, '^(\d{1,2})[:.](\d{2})(?::(\d{2}))?\s*([aApP]\.?[mM]\.?)?$')
    if ($m.Success) {
        $h = [int]$m.Groups[1].Value; $mi = [double]$m.Groups[2].Value
        if ($m.Groups[3].Success) { $mi += [double]$m.Groups[3].Value / 60 }
        if ($m.Groups[4].Success) { $pm = $m.Groups[4].Value.ToLower().StartsWith('p'); if ($h -eq 12) { $h = 0 }; if ($pm) { $h += 12 } }
        if ($h -gt 23 -or $mi -ge 60) { return $null }
        if (-not $m.Groups[4].Success -and $h -lt 3) { return [int][math]::Round($h * 60 + $mi, 0) }
        $tod = $h * 60 + $mi
        $s1 = $ShiftStarts[1] * 1440; $s2 = $ShiftStarts[2] * 1440; $s3 = $ShiftStarts[3] * 1440
        if ($tod -ge $s3 -or $tod -lt $s1 - 120) { $start = $s3 } elseif ($tod -ge $s2 - 120) { $start = $s2 } else { $start = $s1 }
        $mins = $tod - $start
        if ($mins -lt -720) { $mins += 1440 }
        return [int][math]::Round($mins, 0)
    }
    $n = 0.0
    if ($t -match '^-?\d+(\.\d+)?$' -and [double]::TryParse("$t", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$n)) { return [int][math]::Round($n, 0) }
    return $null
}

function Get-LaneFixes($Cells) {
    $forms = [ordered]@{}
    $nonBlank = 0
    foreach ($x in $Cells) {
        $f = "$($x.F)"
        if ($f -eq '') { continue }
        $nonBlank++
        if ($f.StartsWith('=') -and -not $f.Contains('#REF!')) {
            if (-not $forms.Contains($f)) { $forms[$f] = New-Object System.Collections.ArrayList }
            [void]$forms[$f].Add([int]$x.C)
        }
    }
    $best = $null; $bestN = 0
    foreach ($k in $forms.Keys) { if ($forms[$k].Count -gt $bestN) { $best = $k; $bestN = $forms[$k].Count } }
    $out = @()
    if ($null -eq $best -or $bestN -lt 3 -or ($bestN * 2) -le $nonBlank) { return $out }
    $lo = ($forms[$best] | Measure-Object -Minimum).Minimum
    $hi = ($forms[$best] | Measure-Object -Maximum).Maximum
    foreach ($x in $Cells) {
        $f = "$($x.F)"
        if ($f -eq $best) { continue }
        if ($f -eq '' -and ($x.C -lt $lo -or $x.C -gt $hi)) { continue }
        $kind = 'typed'
        if ($f -eq '') { $kind = 'blank' } elseif ($f.Contains('#REF!')) { $kind = 'ref' } elseif ($f.StartsWith('=')) { $kind = 'formula' }
        $out += [pscustomobject]@{ C = [int]$x.C; Kind = $kind; Was = $f; Fix = $best }
    }
    return $out
}

function Get-RowFixes($FormulasR1C1, [int]$FirstCol) {
    $a = @(); $b = @()
    for ($i = 0; $i -lt $FormulasR1C1.Count; $i++) {
        $cell = @{ C = $FirstCol + $i; F = $FormulasR1C1[$i] }
        if ($i % 2 -eq 0) { $a += $cell } else { $b += $cell }
    }
    return @(Get-LaneFixes $a) + @(Get-LaneFixes $b)
}

function Get-LatestDayFormula([string]$Formula, [int]$Row) {
    $m = [regex]::Match($Formula.Trim(), '^=\$?[A-S]\$?(\d+)\s*\*\s*\$?C\$?(\d+)$', 'IgnoreCase')
    if (-not $m.Success -or [int]$m.Groups[1].Value -ne $Row) { return $null }
    $goal = '$C$' + $m.Groups[2].Value
    $day = 'IF($C$5="",7,MIN(7,MAX(1,$C$5-$C$4+6)))'
    $expr = '0'
    for ($d = 1; $d -le 7; $d++) {
        $a = (Get-ColLetter (2 + 2 * $d)) + $Row
        $expr = "IF(AND($day>=$d,N($a)<>0),$a,$expr)"
    }
    return "=$expr*$goal"
}

function Get-Typed {
    $t = [pscustomobject]@{ Absent = @{}; Start = @{} }
    if (-not (Test-Path -LiteralPath $Script:TypedPath)) { return $t }
    try {
        $j = Get-Content -LiteralPath $Script:TypedPath -Raw | ConvertFrom-Json
        if ($j.Absent) { foreach ($p in $j.Absent.PSObject.Properties) { $t.Absent[$p.Name] = @($p.Value) } }
        if ($j.Start) { foreach ($p in $j.Start.PSObject.Properties) { $t.Start[$p.Name] = $p.Value } }
    }
    catch { Say "  Couldn't read $Script:TypedPath - starting fresh." Yellow }
    return $t
}

function Save-Typed($Typed) {
    $cut = (Get-Date).Date.AddDays(-60).ToString('yyyy-MM-dd')
    $keep = { param($h) $o = [ordered]@{}; foreach ($k in ($h.Keys | Sort-Object)) { if (($k -split '\|')[1] -ge $cut) { $o[$k] = $h[$k] } }; return $o }
    $obj = [ordered]@{ Absent = (& $keep $Typed.Absent); Start = (& $keep $Typed.Start) }
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Script:TypedPath) | Out-Null
    $obj | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Script:TypedPath -Encoding UTF8
}

function Get-TypedKey($Shop, [datetime]$Day) { return ($Shop.Code + '|' + $Day.ToString('yyyy-MM-dd', $Script:Inv)) }

function Start-Excel {
    Say '  Starting Excel...' DarkGray
    $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $xl = New-Object -ComObject Excel.Application
    $Script:XlPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    $xl.EnableEvents = $false
    try { $xl.AutomationSecurity = 3 } catch { }
    try { $xl.AskToUpdateLinks = $false } catch { }
    return $xl
}

function Stop-Excel($Xl) {
    if ($null -eq $Xl) { return }
    try { foreach ($w in @($Xl.Workbooks)) { try { $w.Close($false) } catch { } } } catch { }
    try { $Xl.Quit() } catch { }
    Remove-Com $Xl
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
} 

function Copy-ToWork([string]$Path, [string]$Work) {
    $dir = Join-Path $Work 'in'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $dest = Get-UniquePath $dir ([IO.Path]::GetFileName($Path))
    try { Copy-Item -LiteralPath $Path -Destination $dest -Force }
    catch { throw "Can't copy $([IO.Path]::GetFileName($Path)) from $(Split-Path -Parent $Path): $($_.Exception.Message)" }
    try { Unblock-File -LiteralPath $dest -ErrorAction SilentlyContinue } catch { }
    try { (Get-Item -LiteralPath $dest).IsReadOnly = $false } catch { }
    return $dest
}

function Open-ExcelBook($Xl, [string]$Path, [bool]$ReadOnly) {
    $name = [IO.Path]::GetFileName($Path)
    $wb = $null
    try { $wb = $Xl.Workbooks.Open($Path, 0, $ReadOnly) } catch { throw "Excel couldn't open ${name}: $($_.Exception.Message)" }
    if ($null -eq $wb) { try { $wb = $Xl.ProtectedViewWindows.Open($Path).Edit() } catch { } }
    if ($null -eq $wb) { throw "Excel opened $name in Protected View and won't hand it over. Open it once in Excel, click Enable Editing, close it, and run again." }
    return $wb
}

function Test-FileLocked([string]$Path) {
    try { $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None); $fs.Close(); return $false }
    catch { return $true }
}

function Show-Error($Err) {
    Say ('  Error: ' + $Err.Exception.Message) Red
    $ii = $Err.InvocationInfo
    if ($null -ne $ii -and $ii.ScriptLineNumber -gt 0) { Say ("  (DailyOps.ps1 line {0}: {1})" -f $ii.ScriptLineNumber, "$($ii.Line)".Trim()) DarkGray }
}

function Open-InputBook($Xl, [string]$Path, [string]$Work) {
    $wb = Open-ExcelBook $Xl (Copy-ToWork $Path $Work) $true
    $sheets = @()
    foreach ($ws in $wb.Worksheets) { $sheets += [pscustomobject]@{ Name = $ws.Name; Hidden = ($ws.Visible -ne -1); Ws = $ws; Grid = $null } }
    return [pscustomobject]@{ Name = [IO.Path]::GetFileName($Path); Path = $Path; Wb = $wb; Sheets = $sheets }
}

function Close-InputBook($Book) {
    if ($null -eq $Book -or $null -eq $Book.Wb) { return }
    foreach ($s in $Book.Sheets) { Remove-Com $s.Ws; $s.Ws = $null }
    try { $Book.Wb.Close($false) } catch { }
    Remove-Com $Book.Wb
    $Book.Wb = $null
}

function Get-WbSheet($Wb, [string]$Name) {
    foreach ($ws in $Wb.Worksheets) { if ($ws.Name -ieq $Name) { return $ws } }
    return $null
}

function ConvertTo-ComValue($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [psobject]) { $v = $v.psobject.BaseObject }
    if ($v -is [double] -or $v -is [string] -or $v -is [bool]) { return $v }
    if ($v -is [int] -or $v -is [long] -or $v -is [decimal] -or $v -is [single] -or $v -is [int16] -or $v -is [byte]) { return [double]$v }
    if ($v -is [datetime]) { return $v.ToOADate() }
    return [string]$v
}

function Convert-TextToNumbers($Ws, [int]$Col, [int]$FirstRow, [int]$Count) {
    $err = @{ -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'; -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!' }
    $chunk = 20000
    for ($start = 0; $start -lt $Count; $start += $chunk) {
        $len = [math]::Min($chunk, $Count - $start)
        $rng = $Ws.Cells.Item($FirstRow + $start, $Col).Resize($len, 1)
        try {
            $rng.NumberFormat = 'General'
            $v = $rng.Value2
            $out = [Array]::CreateInstance([object], $len, 1)
            for ($i = 0; $i -lt $len; $i++) {
                if ($v -is [array]) { $x = $v[($v.GetLowerBound(0) + $i), $v.GetLowerBound(1)] } else { $x = $v }
                if ($x -is [int] -and $err.ContainsKey($x)) { $x = $err[$x] }
                $out[$i, 0] = (ConvertTo-ComValue $x)
            }
            try { Set-RangeValue $rng $out }
            catch { [void]$rng.TextToColumns($rng, 1, 1, $false, $false, $false, $false, $false, $false) }
        }
        catch { Warn "Couldn't turn column D text into numbers (rows $($FirstRow + $start)-$($FirstRow + $start + $len - 1)): $($_.Exception.Message)" }
        finally { Remove-Com $rng }
    }
}

function Set-RangeValue($Rng, $x) {
    try { $Rng.Value2 = $x; return } catch { $first = $_.Exception }
    $com = $Rng; if ($com -is [psobject]) { $com = $com.psobject.BaseObject }
    $a = [Array]::CreateInstance([object], 1); $a[0] = $x
    try { [void][System.__ComObject].InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $com, $a); return } catch { }
    if ($x -isnot [array]) {
        $t = $null
        if ($x -is [double]) { $t = $x.ToString('R', $Script:Inv) } elseif ($x -is [bool]) { $t = $x.ToString().ToUpper() } elseif ($x -is [string] -and -not $x.StartsWith('=')) { $t = $x }
        if ($null -ne $t) { try { $Rng.Formula = $t; return } catch { } }
    }
    throw $first
}

function Set-CellValue($Ws, [int]$r, [int]$c, $v) {
    $addr = (Get-ColLetter ($c - 1)) + $r
    $cell = $null
    try {
        $cell = $Ws.Range($addr)
        if ($null -eq $v -or "$v" -eq '') { [void]$cell.ClearContents() } else { Set-RangeValue $cell (ConvertTo-ComValue $v) }
    }
    catch {
        $x = ConvertTo-ComValue $v; $tn = 'empty'; if ($null -ne $x) { $tn = $x.GetType().Name }
        $sn = '?'; try { $sn = $Ws.Name } catch { }
        Warn ("Couldn't write '{0}'!{1} = '{2}' ({3}): {4}  - left as it was, type it in by hand" -f $sn, $addr, $v, $tn, $_.Exception.Message)
    }
}

function Save-RangePicture($Xl, $Rng, [string]$Path) {
    try { Save-RangePictureOnce $Rng $Path }
    catch {
        if ($Xl.Visible) { throw }
        $Xl.Visible = $true
        try { $Xl.WindowState = -4140 } catch { }
        Start-Sleep -Milliseconds 500
        Save-RangePictureOnce $Rng $Path
    }
}

function Save-RangePictureOnce($Rng, [string]$Path) {
    $ok = $false
    if ([Threading.Thread]::CurrentThread.ApartmentState -eq 'STA') {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing
            for ($try = 0; $try -lt 3 -and -not $ok; $try++) {
                [void]$Rng.CopyPicture(1, 2)
                Start-Sleep -Milliseconds (200 + 300 * $try)
                $img = [Windows.Forms.Clipboard]::GetImage()
                if ($null -ne $img) { $img.Save($Path, [Drawing.Imaging.ImageFormat]::Png); $img.Dispose(); $ok = $true }
            }
        }
        catch { }
    }
    if (-not $ok) {
        $ws = $Rng.Worksheet
        [void]$Rng.CopyPicture(1, 2)
        $co = $ws.ChartObjects().Add($Rng.Left, $Rng.Top, $Rng.Width, $Rng.Height)
        try {
            $co.Chart.ChartArea.Format.Line.Visible = 0
            [void]$co.Activate()
            Start-Sleep -Milliseconds 300
            [void]$co.Chart.Paste()
            [void]$co.Chart.Export($Path)
            $ok = (Test-Path -LiteralPath $Path)
        }
        finally { try { [void]$co.Delete() } catch { } ; Remove-Com $co }
    }
    if (-not $ok) { throw "Couldn't make the picture $([IO.Path]::GetFileName($Path))." }
}

function Find-Files([string]$Dir, [string[]]$Patterns, [int]$Days, [int]$Depth = 0) {
    if (-not $Dir -or -not (Test-Path -LiteralPath $Dir)) { return @() }
    $cut = [datetime]::MinValue
    if ($Days -gt 0) { $cut = (Get-Date).AddDays(-$Days) }
    $files = @()
    try {
        if ($Depth -gt 0) { $all = @(Get-ChildItem -LiteralPath $Dir -File -Recurse -Depth $Depth -ErrorAction SilentlyContinue) }
        else { $all = @(Get-ChildItem -LiteralPath $Dir -File -ErrorAction SilentlyContinue) }
    }
    catch { $all = @(Get-ChildItem -LiteralPath $Dir -File -ErrorAction SilentlyContinue) }
    foreach ($f in $all) {
        if ($f.Name.StartsWith('~$') -or $f.LastWriteTime -lt $cut) { continue }
        foreach ($p in $Patterns) { if ($f.Name -like $p) { $files += $f; break } }
    }
    return @($files | Sort-Object LastWriteTime -Descending)
}

$Script:KindNames = @{ main = 'report workbook'; raw = 'Raw Machine Data file'; tracker = 'start-time tracker'; schedule = 'schedule'; hours = 'Shift Hours / OT report' }
$Script:DoneKinds = @()

function Show-FileDialog([string]$Title, [string]$Dir, [bool]$Multi) {
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { Say '  Browse isn''t available in this window - drag the files in instead.' Yellow; return @() }
    try { Add-Type -AssemblyName System.Windows.Forms } catch { Say '  Browse isn''t available here - drag the files in instead.' Yellow; return @() }
    $d = New-Object Windows.Forms.OpenFileDialog
    $d.Title = $Title
    $d.Multiselect = $Multi
    $d.Filter = 'Excel files and zips (*.xls;*.xlsx;*.xlsm;*.xlsb;*.zip)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.zip|All files (*.*)|*.*'
    if ($Dir -and (Test-Path -LiteralPath $Dir)) { $d.InitialDirectory = $Dir }
    $owner = New-Object Windows.Forms.Form
    $owner.TopMost = $true; $owner.ShowInTaskbar = $false
    try { if ($d.ShowDialog($owner) -eq [Windows.Forms.DialogResult]::OK) { return @($d.FileNames) } }
    finally { $owner.Dispose(); $d.Dispose() }
    return @()
}

function Expand-InputPaths($Paths) {
    $out = @()
    foreach ($p in @($Paths)) {
        $p = Clean-PathText $p
        if ($p -eq '') { continue }
        if (-not (Test-Path -LiteralPath $p)) { Say "  Not found: $p" Yellow; continue }
        if (Test-Path -LiteralPath $p -PathType Container) {
            $inDir = @(Get-ChildItem -LiteralPath $p -File | Where-Object { ($Script:XlExt + '.zip') -contains $_.Extension.ToLower() -and -not $_.Name.StartsWith('~$') } | Sort-Object Name | ForEach-Object { $_.FullName })
            if ($inDir.Count -eq 0) { Say "  No Excel files or zips in $p" Yellow }
            $out += $inDir
            continue
        }
        $ext = [IO.Path]::GetExtension($p).ToLower()
        if (($Script:XlExt + '.zip') -notcontains $ext) { Say "  Not an Excel file or zip: $([IO.Path]::GetFileName($p))" Yellow; continue }
        $out += (Resolve-Path -LiteralPath $p).Path
    }
    return $out
}

function Select-Files {
    param([string]$What, [string]$Kind, $Found, $Drop, [string]$Folder, [switch]$Multi, [switch]$AllByDefault, [string]$SkipText = 'Skip')
    $Found = @($Found)
    $Script:DoneKinds += $Kind
    $chosen = New-Object System.Collections.ArrayList
    $add = {
        param($list, [string]$how)
        foreach ($f in @($list)) {
            if ($chosen -contains $f) { continue }
            $k = Get-FileKind $f
            if ($how -ne 'pick' -and $k -ne $Kind -and $k -ne 'unknown' -and $Drop.ContainsKey($k)) {
                if ($Script:DoneKinds -contains $k) { Say "  $([IO.Path]::GetFileName($f)) looks like a $($Script:KindNames[$k]) - that step is done, so it's ignored." Yellow }
                else { if ($Drop[$k] -notcontains $f) { $Drop[$k] += $f }; Say "  $([IO.Path]::GetFileName($f)) looks like a $($Script:KindNames[$k]) - kept for that step." DarkGray }
                continue
            }
            if (-not $Multi -and $chosen.Count -ge 1) { Say "  Only one file here - ignored $([IO.Path]::GetFileName($f))." Yellow; continue }
            [void]$chosen.Add($f)
            Say "  + $([IO.Path]::GetFileName($f))" Green
        }
    }
    & $add @($Drop[$Kind]) 'dropped'
    $Drop[$Kind] = @()

    if ($Multi) { Say "  Drop the $What here - several at once is fine, and you can drop again to add more." }
    else { Say "  Drop the $What here." }
    $browseDir = $Folder; if (-not $browseDir -or -not (Test-Path -LiteralPath $browseDir)) { $browseDir = '' }
    if ($browseDir) { Say "  [B] Browse...  (opens $browseDir)" } else { Say '  [B] Browse...' }
    if ($Found.Count -gt 0) {
        if ($Folder) { Say "  Found in $($Folder):" DarkGray }
        for ($i = 0; $i -lt $Found.Count; $i++) { Say ('  [{0}] {1}   ({2})' -f ($i + 1), $Found[$i].Name, (Format-Stamp $Found[$i].LastWriteTime)) }
        if ($Multi -and $Found.Count -gt 1) { Say '  [A] All of those' }
    }
    elseif ($Folder) { Say "  (nothing found in $Folder)" DarkGray }
    Say "  [S] $SkipText"

    $default = 'S'
    if ($Found.Count -gt 0) { $default = '1'; if ($Multi -and $AllByDefault -and $Found.Count -gt 1) { $default = 'A' } }
    while ($true) {
        if (-not $Multi -and $chosen.Count -ge 1) { return @($chosen) }
        if ($chosen.Count -gt 0) {
            $a = Ask ("  {0} file{1} - drop more, or Enter = done  (undo / clear)" -f $chosen.Count, $(if ($chosen.Count -ne 1) { 's' } else { '' }))
            if ($a -eq '') { return @($chosen) }
        }
        else {
            $hint = 'drop file(s), B, a number'; if ($Multi) { $hint = 'drop file(s), B, numbers (1,3 / 1-3)' }
            $a = Ask "  $($Script:KindNames[$Kind]) - $hint [$default]"
            if ($a -eq '') {
                if ($default -eq 'S') { return @() }
                if ($default -eq 'A') { & $add @($Found | ForEach-Object { $_.FullName }) 'pick' } else { & $add @($Found[0].FullName) 'pick' }
                return @($chosen)
            }
        }
        if ($a -ieq 'S') { return @() }
        if ($a -ieq 'undo') { if ($chosen.Count -gt 0) { Say "  - $([IO.Path]::GetFileName($chosen[$chosen.Count - 1]))" Yellow; $chosen.RemoveAt($chosen.Count - 1) }; continue }
        if ($a -ieq 'clear') { $chosen.Clear(); Say '  Cleared.' Yellow; continue }
        if ($a -ieq 'B') { & $add (Show-FileDialog "Daily Ops - pick the $What" $browseDir ([bool]$Multi)) 'browse'; continue }
        if ($a -ieq 'A' -and $Found.Count -gt 0) { & $add @($Found | ForEach-Object { $_.FullName }) 'pick'; continue }
        $idx = Get-NumberList $a $Found.Count
        if ($null -ne $idx -and $Found.Count -gt 0) { & $add @($idx | ForEach-Object { $Found[$_ - 1].FullName }) 'pick'; continue }
        $paths = @(Expand-InputPaths (Split-PathList $a))
        if ($paths.Count -eq 0) { Say "  Didn't get a file from that - drag the file(s) into this window, or type B to browse." Yellow; continue }
        & $add $paths 'drop'
    }
}
# when the report was saved by whoever made it (docProps/core.xml) - the best "which copy is newest"
function Get-XlsxMade($Stream) {
    $z = $null
    try {
        $z = New-Object IO.Compression.ZipArchive($Stream, [IO.Compression.ZipArchiveMode]::Read, $true)
        $x = Get-ZipText $z 'docProps/core.xml'
        if (-not $x) { return $null }
        $m = [regex]::Match($x, '<dcterms:modified[^>]*>([^<]+)<')
        if (-not $m.Success) { return $null }
        $d = [datetime]::MinValue
        if ([datetime]::TryParse($m.Groups[1].Value, $Script:Inv, [Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$d)) { return $d.ToLocalTime() }
    }
    catch { }
    finally { if ($null -ne $z) { $z.Dispose() } }
    return $null
}

# Hours: plain files + entries of .zip files -> one per day (newest copy wins), only days in the report week
function Get-HoursCandidates($Paths, [datetime]$Monday, [datetime]$WeekEnd) {
    $list = @()
    foreach ($p in $Paths) {
        if ($p.ToLower().EndsWith('.zip')) {
            $z = $null
            try {
                $z = [IO.Compression.ZipFile]::OpenRead($p)
                foreach ($e in $z.Entries) {
                    $ext = [IO.Path]::GetExtension($e.Name).ToLower()
                    if ($e.Name -eq '' -or $e.Name.StartsWith('~$') -or $Script:XlExt -notcontains $ext) { continue }
                    $made = $null
                    if ($ext -ne '.xls') { $ms = New-Object IO.MemoryStream; $es = $e.Open(); try { $es.CopyTo($ms) } finally { $es.Dispose() }; $ms.Position = 0; $made = Get-XlsxMade $ms; $ms.Dispose() }
                    $list += [pscustomobject]@{ Name = $e.Name; Zip = $p; Entry = $e.FullName; Path = $null; Date = (Get-DateFromName $e.Name); Time = $e.LastWriteTime.DateTime; Made = $made; Use = $false; Note = '' }
                }
            }
            catch { Warn "$([IO.Path]::GetFileName($p)): can't open the zip ($($_.Exception.Message))" }
            finally { if ($null -ne $z) { $z.Dispose() } }
        }
        else {
            $fi = Get-Item -LiteralPath $p
            $made = $null
            if ($fi.Extension -ine '.xls') { try { $fs = [IO.File]::OpenRead($fi.FullName); try { $made = Get-XlsxMade $fs } finally { $fs.Dispose() } } catch { } }
            $list += [pscustomobject]@{ Name = $fi.Name; Zip = $null; Entry = $null; Path = $fi.FullName; Date = (Get-DateFromName $fi.Name); Time = $fi.LastWriteTime; Made = $made; Use = $false; Note = '' }
        }
    }
    foreach ($h in $list) {
        if ($null -eq $h.Date) { $h.Note = 'no date in the name' }
        elseif ($h.Date -lt $Monday -or $h.Date -gt $WeekEnd) { $h.Note = 'not in this report week' }
    }
    foreach ($grp in ($list | Where-Object { $_.Note -eq '' } | Group-Object { $_.Date.ToString('yyyy-MM-dd') })) {
        $sorted = @($grp.Group | Sort-Object @{ Expression = { if ($_.Made) { $_.Made } else { $_.Time } }; Descending = $true }, @{ Expression = { $_.Time }; Descending = $true })
        $sorted[0].Use = $true
        for ($i = 1; $i -lt $sorted.Count; $i++) { $sorted[$i].Note = 'older copy, skipped' }
    }
    return @($list | Sort-Object @{ Expression = { if ($_.Date) { $_.Date } else { [datetime]::MaxValue } } }, @{ Expression = { $_.Time }; Descending = $true })
}

function Show-StartGrid($Shops, $Grid, $Week) {
    $line = '  {0,-9}{1,-17}' -f 'min late', ''
    for ($d = 0; $d -lt 7; $d++) { $line += '{0,-10}' -f (Format-Day $Week.Monday.AddDays($d)) }
    Say $line White
    foreach ($s in $Shops) {
        $g = $Grid[$s.Code]
        $line = '  {0,-9}{1,-17}' -f $s.Label, $g.Src
        for ($d = 0; $d -lt 7; $d++) {
            if ($d -ge $Week.Days) { $txt = '-' }
            else { $v = $g.Values[$d]; if ($null -eq $v -or "$v" -eq '') { $txt = '.' } else { $txt = "$v" }; if ($g.Typed -contains $d) { $txt += '*' } }
            $line += '{0,-10}' -f $txt
        }
        Say $line
        foreach ($n in $g.Notes) { Say "           $n" DarkGray }
    }
}

# total absent for a shop/day (older saved entries were per shift - they're added up)
function Get-AbsentTotal($A) {
    $v = @($A | Where-Object { $null -ne $_ })
    if ($v.Count -eq 0) { return $null }
    return [int](($v | Measure-Object -Sum).Sum)
}

function Show-AbsentGrid($Shops, $Absent, $Sched, $Week) {
    $line = '  {0,-17}' -f 'absent'
    for ($d = 0; $d -lt $Week.Days; $d++) { $line += '{0,-10}' -f (Format-Day $Week.Monday.AddDays($d)) }
    Say $line White
    foreach ($s in $Shops) {
        $line = '  {0,-17}' -f $s.Label
        $line2 = '  {0,-17}' -f '  present/sched'
        for ($d = 0; $d -lt $Week.Days; $d++) {
            $t = Get-AbsentTotal $Absent[(Get-TypedKey $s $Week.Monday.AddDays($d))]
            $sch = $Sched[$s.Code][$d]
            $line += '{0,-10}' -f $(if ($null -ne $t) { "$t" } else { '.' })
            $sTxt = '-'; if ($sch -is [double] -or $sch -is [int]) { $sTxt = "$sch" }
            $pTxt = '-'; if ($null -ne $t -and ($sch -is [double] -or $sch -is [int])) { $pTxt = "$($sch - $t)" }
            $line2 += '{0,-10}' -f "$pTxt/$sTxt"
        }
        Say $line
        Say $line2 DarkGray
    }
}

# After a run: the pictures folder in Explorer + the saved workbook in Excel (once the hidden Excel has closed,
# so the workbook doesn't open inside it).
function Open-WhenDone([string]$Workbook, [string]$Folder) {
    if ($Folder -and (Test-Path -LiteralPath $Folder)) {
        try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $Folder + '"') } catch { try { Invoke-Item -LiteralPath $Folder } catch { } }
    }
    if (-not $Workbook -or -not (Test-Path -LiteralPath $Workbook)) { return }
    $ids = @($Script:XlPids | Where-Object { $_ })
    if ($ids.Count -gt 0) {
        try { Wait-Process -Id $ids -Timeout 15 -ErrorAction SilentlyContinue } catch { }
        foreach ($id in $ids) { try { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue } catch { } }   # ours, already closed + saved
    }
    Say "  Opening $([IO.Path]::GetFileName($Workbook))..." DarkGray
    try { Invoke-Item -LiteralPath $Workbook } catch { Say "  Couldn't open the workbook: $($_.Exception.Message)" Yellow }
}

function Invoke-Report([string[]]$InputPaths, [switch]$PicturesOnly) {
    $cfg = Get-Config
    $shops = @($cfg.Shops)
    $Script:Warn.Clear()
    $Script:DoneKinds = @()
    $days = Get-SettingInt $cfg 'DaysToLook' 14
    $starts = Get-ShiftStarts $cfg
    $typed = Get-Typed

    $drop = @{ main = @(); raw = @(); tracker = @(); schedule = @(); hours = @() }
    foreach ($p in @($InputPaths)) {
        if (-not $p) { continue }
        if (Test-Path -LiteralPath $p -PathType Container) {
            foreach ($f in (Get-ChildItem -LiteralPath $p -File -Recurse -ErrorAction SilentlyContinue)) { $k = Get-FileKind $f.FullName; if ($drop.ContainsKey($k)) { $drop[$k] += $f.FullName } }
            continue
        }
        $k = Get-FileKind $p
        if ($drop.ContainsKey($k)) { $drop[$k] += (Resolve-Path -LiteralPath $p).Path } else { Say "  Not sure what $([IO.Path]::GetFileName($p)) is - skipping it." Yellow }
    }

    $nSteps = 8; if ($PicturesOnly) { $nSteps = 1 }
    $xl = $null; $wb = $null; $openWb = $null; $openDir = $null
    $work = Join-Path ([IO.Path]::GetTempPath()) ('DailyOps_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    try {
        Head "STEP 1 of $nSteps - Report workbook"
        $repDir = Get-Folder $cfg 'ReportFolder'
        $found = @(Find-Files $repDir @('*.xlsm', '*.xlsx') 0 | Where-Object { $_.Name -match '(?i)ops|report' } | Select-Object -First 6)
        $sel = @(Select-Files -What 'Daily Ops Report workbook (.xlsm)' -Kind 'main' -Found $found -Drop $drop -Folder $repDir -SkipText 'Cancel')
        if ($sel.Count -eq 0) { throw (New-Object OperationCanceledException) }
        $mainPath = $sel[0]

        $xl = Start-Excel
        $readOnly = Test-FileLocked $mainPath
        if ($readOnly) { Warn 'The workbook is open somewhere else (you or someone on the J: drive) - a copy will be saved in the output folder instead.' }
        $dots = Start-Dots "Opening $([IO.Path]::GetFileName($mainPath))"
        try {
            $localMain = Copy-ToWork $mainPath $work
            $wb = Open-ExcelBook $xl $localMain $false
        }
        catch { Stop-Dots $dots 'failed'; $dots = $null; throw }
        Stop-Dots $dots
        $xl.Calculation = -4135
        $need = @($Script:SUMMARY, $Script:CUBIC, $Script:YEARLY) + @($shops | ForEach-Object { $_.Sheet })
        $missing = @($need | Where-Object { $null -eq (Get-WbSheet $wb $_) })
        if ($missing.Count -gt 0) { throw "This workbook is missing: $($missing -join ', '). Is it the Daily Ops Report?" }
        $sum = Get-WbSheet $wb $Script:SUMMARY
        $core = @{}
        foreach ($s in $shops) { $ws = Get-WbSheet $wb $s.Sheet; $core[$s.Code] = [pscustomobject]@{ Ws = $ws; Grid = (Get-ComGrid $ws) } }

        if ($PicturesOnly) {
            $week = $null
            $c5 = $sum.Range('C5').Value2
            if ($c5 -is [double]) { $week = Get-ReportWeek ([datetime]::FromOADate($c5)) }
            $mode = 2
        }
        else {
            Head "STEP 2 of $nSteps - Report date"
            $c5 = $sum.Range('C5').Value2
            if ($c5 -is [double]) { Say "  In the workbook now: $(Format-Day ([datetime]::FromOADate($c5)))" DarkGray }
            while ($true) {
                $a = Ask 'Report date (Enter = today, or e.g. 10/19, y = yesterday)' ((Get-Date).ToString('MM/dd', $Script:Inv))
                $rd = ConvertTo-DateNear $a
                if ($null -ne $rd) { break }
                Say "  Didn't understand '$a'." Yellow
            }
            $week = Get-ReportWeek $rd
            if ($week.Days -eq 7) { $span = 'the whole week' } else { $span = "Mon - $($Script:DayAbbr[$week.Days - 1])" }
            Say ("  -> {0}: week ending Sun {1}, shows {2} ({3} - {4})" -f (Format-Day $week.Report), $week.WeekEnd.ToString('MM/dd', $Script:Inv), $span, (Format-Day $week.Monday), (Format-Day $week.LastDay)) Green

            Head "STEP 3 of $nSteps - Raw Machine Data"
            $rawDir = Get-Folder $cfg 'RawFolder'
            $found = @(Find-Files $rawDir @('*raw*machine*data*.xls*') $days | Select-Object -First 5)
            $rawPath = @(Select-Files -What 'Raw Machine Data file' -Kind 'raw' -Found $found -Drop $drop -Folder $rawDir -SkipText 'Skip - keep the data already in YEARLY DATA') | Select-Object -First 1

            Head "STEP 4 of $nSteps - Machine start times (raw data first, trackers fill the gaps)"
            $trDir = Get-Folder $cfg 'TrackerFolder'
            $all = @(Find-Files $trDir @('*start*time*.xls*') 0 1)
            $found = @(); $seen = @{}
            foreach ($f in $all) {
                $grp = $f.Name
                foreach ($s in $shops) { if (@(Get-FileWords $s | Where-Object { Test-HasWord $f.Name $_ }).Count -gt 0) { $grp = $s.FileWords; break } }
                $key = $grp + '|' + (Get-ShiftOfName $f.Name)
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true; $found += $f
            }
            $found = @($found | Select-Object -First 12)
            $trackPaths = @(Select-Files -What 'start-time trackers (they only fill in machines with no raw data)' -Kind 'tracker' -Found $found -Drop $drop -Folder $trDir -Multi -AllByDefault -SkipText 'None - raw data only')
            $trackers = @()
            foreach ($p in $trackPaths) {
                Say "  Reading $([IO.Path]::GetFileName($p))" DarkGray
                try { $b = Open-InputBook $xl $p $work } catch { Warn "$([IO.Path]::GetFileName($p)): couldn't open it ($($_.Exception.Message))"; continue }
                $ids = @(Get-TrackerShops $b $shops)
                if ($ids.Count -eq 0) { Warn "$($b.Name): no shop tab - skipped"; Close-InputBook $b; continue }
                $trackers += [pscustomobject]@{ Book = $b; Shift = (Get-ShiftOfName $p); Shops = $ids }
            }
            $rawCols = $null
            if ($rawPath) {
                Say "  Reading start times from $([IO.Path]::GetFileName($rawPath))..." DarkGray
                try {
                    $rb = Open-InputBook $xl $rawPath $work
                    try { $rawCols = Read-RawColumns $rb.Sheets[0].Ws 2 } finally { Close-InputBook $rb }
                }
                catch { Warn "Couldn't read the raw data file for start times: $($_.Exception.Message)" }
            }
            else {
                Say '  No raw file picked - using the data already in YEARLY DATA...' DarkGray
                try { $ydWs = Get-WbSheet $wb $Script:YEARLY; $rawCols = Read-RawColumns $ydWs 4 } catch { Warn "Couldn't read YEARLY DATA for start times: $($_.Exception.Message)" }
            }
            $ops = Get-RawOps $rawCols $week.Monday.AddDays(-1) $week.Monday.AddDays(8)
            $assetWs = Get-WbSheet $wb 'Assets'
            $areas = @{}; if ($assetWs) { $areas = Get-AssetAreas (Get-ComGrid $assetWs) } else { Warn "No 'Assets' sheet in the workbook - can't tell which machines belong to which shop." }
            $rShifts = @("$(Get-Setting $cfg 'StartShifts')" -split '[,;\s]+' | Where-Object { $_ -match '^[123]$' } | ForEach-Object { [int]$_ })
            if ($rShifts.Count -eq 0) { $rShifts = @(1, 2, 3) }
            $rawStart = @{}
            foreach ($s in $shops) {
                $mach = @(Get-ShopMachines $s $areas)
                $trk = Get-TrackerMachineTimes $trackers $s $shops $areas $week.Monday $starts
                $rawStart[$s.Code] = Get-RawStartTimes $ops $mach $week $starts $rShifts (Get-SettingInt $cfg 'StartEarlyMin' 30) (Get-SettingInt $cfg 'StartLateCap' 120) $trk
                $rawStart[$s.Code] | Add-Member -NotePropertyName MachineCount -NotePropertyValue ($mach.Count + $trk.Count) -Force
                $rawStart[$s.Code] | Add-Member -NotePropertyName TrackerCount -NotePropertyValue $trk.Count -Force
            }
            foreach ($t in $trackers) { Close-InputBook $t.Book }
            Say ("  {0} operations this week on {1} machines" -f (@($ops.Values | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum, $ops.Count) DarkGray
            $startGrid = @{}
            foreach ($s in $shops) {
                $g = $core[$s.Code].Grid
                $row = Find-InCol $g 3 1 (Get-LastRow $g) 'Average time to start tool'
                $r = $rawStart[$s.Code]
                $haveData = ($r.TrackerCount -gt 0 -or ($r.MachineCount -gt 0 -and $ops.Count -gt 0))
                $base = @($r.Values | ForEach-Object { if ($null -eq $_) { 'N/A' } else { $_ } })
                $src = 'raw'; if ($r.TrackerCount -gt 0) { $src = 'raw + tracker' }; $notes = @()
                if ($s.Machine -gt 0) { $src += " #$($s.Machine)" }
                if (-not $haveData) { $notes += "no raw data or tracker times for $($s.Label) - left as in the workbook"; $base = @($null, $null, $null, $null, $null, $null, $null) }
                if ($row -eq 0) { $notes += "can't find 'Average time to start tool' in column C of $($s.Sheet)" }
                $vals = @($base); $typedDays = @()
                for ($d = 0; $d -lt $week.Days; $d++) {
                    $k = Get-TypedKey $s $week.Monday.AddDays($d)
                    if ($typed.Start.ContainsKey($k)) { $vals[$d] = $typed.Start[$k]; $typedDays += $d }
                }
                $startGrid[$s.Code] = [pscustomobject]@{ Values = $vals; Base = $base; Src = $src; Notes = $notes; Typed = $typedDays; Row = $row; FromData = $haveData }
            }
            Write-Host ''
            Show-RawShiftTable $shops $rawStart $week
            while ($true) {
                Write-Host ''
                Show-StartGrid $shops $startGrid $week
                Say "  Enter = OK.  Type over:  'hms tue 12' minutes | 'hms tue 6:25' clock time | 'fel wed na' | 'hms tue -' back to computed | 'undo hms'" DarkGray
                $a = Ask 'Start times OK?'
                if ($a -eq '') { break }
                $m = [regex]::Match($a, '^(?i)undo\s+(\S+)$')
                if ($m.Success) {
                    $s = Find-Shop $shops $m.Groups[1].Value
                    if ($null -eq $s) { Say '  Unknown shop.' Yellow; continue }
                    $g = $startGrid[$s.Code]
                    foreach ($d in $g.Typed) { $g.Values[$d] = $g.Base[$d]; [void]$typed.Start.Remove((Get-TypedKey $s $week.Monday.AddDays($d))) }
                    $g.Typed = @()
                    continue
                }
                $m = [regex]::Match($a, '^(\S+)\s+(\S+)\s+(.+)$')
                if (-not $m.Success) { Say '  Didn''t understand that.' Yellow; continue }
                $s = Find-Shop $shops $m.Groups[1].Value; $d = Get-DayIndex $m.Groups[2].Value
                if ($null -eq $s) { Say '  Unknown shop.' Yellow; continue }
                if ($d -lt 0) { Say '  Unknown day.' Yellow; continue }
                if ($d -ge $week.Days) { Say '  That day is after the report date - it stays blank.' Yellow; continue }
                $g = $startGrid[$s.Code]; $k = Get-TypedKey $s $week.Monday.AddDays($d)
                $txt = $m.Groups[3].Value.Trim()
                if ($txt -eq '-') { $g.Values[$d] = $g.Base[$d]; $g.Typed = @($g.Typed | Where-Object { $_ -ne $d }); [void]$typed.Start.Remove($k); continue }
                $v = ConvertFrom-StartText $txt $starts
                if ($null -eq $v) { Say "  Couldn't read '$txt' - type minutes, a time like 6:25, or NA." Yellow; continue }
                if ($v -is [string] -and $v -eq '') { $v = 'N/A' }
                $g.Values[$d] = $v; $typed.Start[$k] = $v
                if ($g.Typed -notcontains $d) { $g.Typed += $d }
            }
            Write-RawStartLog $cfg $shops $rawStart $week $typed

            Head "STEP 5 of $nSteps - Schedules (scheduled headcount)"
            $scDir = Get-Folder $cfg 'ScheduleFolder'
            $all = @(Find-Files $scDir @('*hours*required*.xls*', '*schedule*.xls*') 0 1)
            $found = @(); $seenWords = @{}
            foreach ($f in $all) {
                $grp = ''
                foreach ($s in $shops) { if (@(Get-FileWords $s | Where-Object { Test-HasWord $f.Name $_ }).Count -gt 0) { $grp = $s.FileWords; break } }
                if ($grp -eq '' -or $seenWords.ContainsKey($grp)) { continue }
                $seenWords[$grp] = $true; $found += $f
            }
            $schedPaths = @(Select-Files -What 'schedules (Hours Required)' -Kind 'schedule' -Found $found -Drop $drop -Folder $scDir -Multi -AllByDefault -SkipText 'None - keep the scheduled numbers in the workbook')
            $sched = @{}; $schedFrom = @{}
            foreach ($p in $schedPaths) {
                Say "  Reading $([IO.Path]::GetFileName($p))" DarkGray
                try { $b = Open-InputBook $xl $p $work } catch { Warn "$([IO.Path]::GetFileName($p)): couldn't open it ($($_.Exception.Message))"; continue }
                try {
                    $ids = @(Get-ScheduleShops $b $shops)
                    if ($ids.Count -eq 0) { Warn "$($b.Name): can't tell which shop this is"; continue }
                    foreach ($id in $ids) {
                        $s = Find-Shop $shops $id
                        $res = Get-ScheduledHeadcount $b $s $week.Monday
                        $sched[$s.Code] = $res.Values; $schedFrom[$s.Code] = "$($b.Name) (tab $($res.Tab))"
                        Say ('  {0,-8} {1}   tab {2}' -f $s.Label, (($res.Values | Select-Object -First $week.Days) -join ' '), $res.Tab) DarkGray
                    }
                }
                finally { Close-InputBook $b }
            }
            $schedNow = @{}
            foreach ($s in $shops) {
                $g = $core[$s.Code].Grid
                $row = Find-InCol $g 3 1 (Get-LastRow $g) '# of employee present'
                $core[$s.Code] | Add-Member -NotePropertyName PresentRow -NotePropertyValue $row -Force
                $vals = @()
                for ($d = 0; $d -lt 7; $d++) {
                    if ($sched.ContainsKey($s.Code)) { $vals += $sched[$s.Code][$d] }
                    elseif ($row -gt 0) { $vals += (GridV $g $row (7 + 2 * $d)) }
                    else { $vals += $null }
                }
                $schedNow[$s.Code] = $vals
            }

            Head "STEP 6 of $nSteps - Absent per shop"

    if ((Test-SettingYes $cfg 'CallInScan') -and (AskYesNo "Scan Outlook for call-in voicemail MP3s?" $false)) {
        try { Get-OutlookCallInAudio $cfg } catch { Warn "Call-in voicemail scan stopped: $($_.Exception.Message)" }
    }

    # one number per shop per day: first the latest report day, shop by shop
    $dLast = $week.Days - 1
    $kDay = Format-Day $week.Monday.AddDays($dLast)
    Say "  How many were absent on ${kDay}?   (number, Enter = keep [shown] / skip, - = clear)" DarkGray
    foreach ($s in $shops) {
        $k = Get-TypedKey $s $week.Monday.AddDays($dLast)
        $cur = Get-AbsentTotal $typed.Absent[$k]
        while ($true) {
            $a = Ask ("  {0} absent {1}" -f $s.Label, $kDay) $(if ($null -ne $cur) { "$cur" } else { '' })
            if ($a -eq '') { break }
            if ($a -eq '-') { [void]$typed.Absent.Remove($k); break }
            if ($a -match '^\d{1,3}$') { $typed.Absent[$k] = @([int]$a, $null, $null); break }
            Say '  Just a number (e.g. 2), Enter, or -' Yellow
        }
    }
    try { Save-Typed $typed } catch { }

    # then the table, with edits like  hms mon 3
    while ($true) {
        Write-Host ''
        Show-AbsentGrid $shops $typed.Absent $schedNow $week
        Say "  Enter = OK.  'hms mon 3' = HMS Monday 3 absent | 'hms 3' = latest day | 'hms mon -' = clear | 'clear hms'" DarkGray
        $a = Ask 'Absences OK?'
        if ($a -eq '') { break }

        $m = [regex]::Match($a, '^(?i)clear\s+(\S+)$')
        if ($m.Success) {
            $s = Find-Shop $shops $m.Groups[1].Value
            if ($null -eq $s) { Say '  Unknown shop.' Yellow; continue }
            for ($d = 0; $d -lt 7; $d++) { [void]$typed.Absent.Remove((Get-TypedKey $s $week.Monday.AddDays($d))) }
            continue
        }

        $tok = @($a -split '\s+' | Where-Object { $_ -ne '' })
        $s = Find-Shop $shops $tok[0]
        if ($null -eq $s -or $tok.Count -lt 2) { Say '  Start with the shop, e.g. hms mon 3' Yellow; continue }
        $d = $dLast; $i = 1
        if ((Get-DayIndex $tok[1]) -ge 0) { $d = Get-DayIndex $tok[1]; $i = 2 }
        if ($d -ge $week.Days) { Say '  That day is after the report date.' Yellow; continue }
        $rest = @($tok | Select-Object -Skip $i)
        if ($rest.Count -ne 1 -or $rest[0] -notmatch '^(\d{1,3}|-)$') { Say '  One number per day, e.g. hms mon 3  (or - to clear)' Yellow; continue }
        $k = Get-TypedKey $s $week.Monday.AddDays($d)
        if ($rest[0] -eq '-') { [void]$typed.Absent.Remove($k) } else { $typed.Absent[$k] = @([int]$rest[0], $null, $null) }
        try { Save-Typed $typed } catch { }
    }
            try { Save-Typed $typed } catch { Say "  Couldn't save typed values: $($_.Exception.Message)" Yellow }

            Head "STEP 7 of $nSteps - Hours (Shift Hours and OT Reports)"
            $hDir = Get-Folder $cfg 'HoursFolder'
            $found = @(Find-Files $hDir @('*shift*hours*.xls*', '*ot*report*.xls*', '*.zip') $days | Where-Object { -not $_.Name.ToLower().EndsWith('.zip') -or $_.Name -match '(?i)hour|ot.?report|overtime|shift' } | Select-Object -First 15)
            foreach ($f in @(Get-OutlookHours $cfg $week $work $days)) { if ($drop.hours -notcontains $f) { $drop.hours += $f } }
            $hPaths = @(Select-Files -What 'Shift Hours and OT Reports' -Kind 'hours' -Found $found -Drop $drop -Folder $hDir -Multi -AllByDefault -SkipText 'None - keep the hours in the workbook')
            $hoursList = @(Get-HoursCandidates $hPaths $week.Monday $week.WeekEnd)
            foreach ($h in $hoursList) {
                $from = ''; if ($h.Zip) { $from = "  (in $([IO.Path]::GetFileName($h.Zip)))" } elseif ($h.Path -and $Script:Origin.ContainsKey($h.Path)) { $from = "  ($($Script:Origin[$h.Path]))" }
                if ($h.Use) { Say ('  {0}  {1}{2}' -f (Format-Day $h.Date), $h.Name, $from) Green }
                else { Say ('  --         {0}{1}  - {2}' -f $h.Name, $from, $h.Note) DarkGray }
            }
            $hours = @()
            foreach ($h in @($hoursList | Where-Object { $_.Use })) {
                $p = $h.Path
                if ($h.Zip) {
                    $z = [IO.Compression.ZipFile]::OpenRead($h.Zip)
                    try { $p = Get-UniquePath $work (Get-SafeFileName $h.Name); [IO.Compression.ZipFileExtensions]::ExtractToFile($z.GetEntry($h.Entry), $p) } finally { $z.Dispose() }
                }
                try { $b = Open-InputBook $xl $p $work } catch { Warn "$($h.Name): couldn't open it ($($_.Exception.Message))"; continue }
                try {
                    $r = Read-HoursBook $b $h.Name $shops
                    if ($r.Ok) { $hours += $r } else { Warn "$($h.Name): $($r.Note)" }
                }
                finally { Close-InputBook $b }
            }

            Head "STEP 8 of $nSteps - Save"
            Say '  [1] Workbook + the 2 pictures   [2] Pictures only (workbook not saved)   [3] Workbook only'
            $mode = AskInt 'What to save?' 1 3 1
        }

        Head 'Building...'
        if (-not $PicturesOnly) {
            $sum.Range('C5').Value2 = (ConvertTo-ComValue (Get-Serial $week.Report))
            $sum.Range('C4').Value2 = (ConvertTo-ComValue (Get-Serial $week.WeekEnd))
            for ($i = 0; $i -lt 7; $i++) {
                $hide = ($week.Monday.AddDays($i) -gt $week.LastDay)
                $c1 = Get-ColLetter (3 + 2 * $i); $c2 = Get-ColLetter (4 + 2 * $i)
                $sum.Range("${c1}:${c2}").EntireColumn.Hidden = $hide
            }
            $sum.Range('R:S').EntireColumn.Hidden = ($week.WeekEnd -gt $week.LastDay)
            Say '  Report date + day columns set' DarkGray

            if ($rawPath) {
                Say "  Importing $([IO.Path]::GetFileName($rawPath))..." DarkGray
                $src = Open-ExcelBook $xl (Copy-ToWork $rawPath $work) $true
                try {
                    $s1 = $src.Worksheets.Item(1)
                    $dst = Get-WbSheet $wb $Script:YEARLY
                    $lastSrc = $s1.Cells.Item($s1.Rows.Count, 1).End(-4162).Row
                    $lastDst = $dst.Cells.Item($dst.Rows.Count, 1).End(-4162).Row
                    $n = $lastSrc - 1
                    if ($lastDst -ge 4) { [void]$dst.Range("A4:U$lastDst").ClearContents() }
                    if ($n -gt 0) {
                        [void]$s1.Range('A2').Resize($n, 21).Copy()
                        [void]$dst.Range('A4').PasteSpecial(-4163)
                        try { $xl.CutCopyMode = 0 } catch { }
                        Convert-TextToNumbers $dst 4 4 $n
                        $lastF = $dst.Cells.Item($dst.Rows.Count, 22).End(-4162).Row
                        if ($lastF -ge 4 -and $lastF -lt $n + 3) {
                            [void]$dst.Range("V${lastF}:AD${lastF}").Copy($dst.Range("V$($lastF + 1):AD$($n + 3)"))
                            Say "  Filled the V:AD formulas down to row $($n + 3)" DarkGray
                        }
                    }
                    Say "  $n rows imported into YEARLY DATA" Green
                }
                finally { try { $src.Close($false) } catch { } ; Remove-Com $src }
            }

            foreach ($s in $shops) {
                $ws = $core[$s.Code].Ws; $g = $core[$s.Code].Grid
                $sg = $startGrid[$s.Code]
                if ($sg.Row -gt 0 -and ($sg.FromData -or $sg.Typed.Count -gt 0)) {
                    for ($d = 0; $d -lt 7; $d++) {
                        if ($d -ge $week.Days) { Set-CellValue $ws $sg.Row (8 + 2 * $d) $null; continue }
                        $v = $sg.Values[$d]; if ($null -eq $v -or "$v" -eq '') { $v = 'N/A' }
                        Set-CellValue $ws $sg.Row (8 + 2 * $d) $v
                    }
                }
                $prow = $core[$s.Code].PresentRow
                if ($sched.ContainsKey($s.Code)) {
                    if ($prow -gt 0) { for ($d = 0; $d -lt 7; $d++) { Set-CellValue $ws $prow (7 + 2 * $d) $sched[$s.Code][$d] } }
                }
                for ($d = 0; $d -lt $week.Days; $d++) {
                    $a = $typed.Absent[(Get-TypedKey $s $week.Monday.AddDays($d))]
                    if (-not $a -or @($a | Where-Object { $null -ne $_ }).Count -eq 0) { continue }
                    if ($prow -eq 0) { break }
                    $sch = $schedNow[$s.Code][$d]
                    if (-not ($sch -is [double] -or $sch -is [int])) { continue }
                    $present = $sch - (($a | Where-Object { $null -ne $_ } | Measure-Object -Sum).Sum)
                    Set-CellValue $ws $prow (8 + 2 * $d) $present
                }
                if ($hours.Count -gt 0) {
                    $rReg = Find-InCol $g 3 1 (Get-LastRow $g) 'Regular Hours'
                    if ($rReg -gt 0) {
                        foreach ($h in $hours) {
                            $col = 7 + 2 * [int]($h.Date - $week.Monday).TotalDays
                            $dd = $h.Depts[$s.Code]
                            $vals = @([math]::Round($dd.Reg, 2), [math]::Round($dd.OT, 2), [math]::Round($dd.DT, 2))
                            for ($k = 0; $k -lt 3; $k++) { Set-CellValue $ws ($rReg + $k) $col $vals[$k]; Set-CellValue $ws ($rReg + $k) ($col + 1) $vals[$k] }
                        }
                    }
                }
            }

            Invoke-FormulaCheck $wb $sum (Test-SettingYes $cfg 'FixFormulas')

            if (Test-SettingYes $cfg 'LatestDayCell') {
                $fixed = @()
                $cv = $sum.Range('C9:C58').Value2; $tf = $sum.Range('T9:T58').Formula
                for ($r = 9; $r -le 58; $r++) {
                    $lab = $cv[($r - 8), 1]
                    if ($lab -isnot [string] -or $lab -notmatch '(?i)performance') { continue }
                    $new = Get-LatestDayFormula "$($tf[($r - 8), 1])" $r
                    if ($new) { $sum.Range("T$r").Formula = [string]$new; $fixed += "T$r" }
                }
            }
        }

        try { Set-AbsenteeismCheck $shops $core } catch { Warn "Absenteeism check: $($_.Exception.Message)" }

        Say '  Recalculating...' DarkGray
        $xl.Calculation = -4105
        $xl.Calculate()

        if ($null -ne $week -and $week.Days -gt 0) {
            Say '  Updating the cubic-inch pivots...' DarkGray
            Set-CubicPivots $wb $week
            $xl.Calculate()
        }

        $outDir = Join-Path (Get-OutputRoot $cfg) ($(if ($week) { $week.Report.ToString('yyyy-MM-dd', $Script:Inv) } else { (Get-Date).ToString('yyyy-MM-dd') }))
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
        $tag = ''; if ($week) { $tag = $week.Report.ToString('yyyy-MM-dd', $Script:Inv) } else { $tag = (Get-Date).ToString('yyyy-MM-dd') }
        $done = @()
        if ($mode -eq 1 -or $mode -eq 2) {
            Say '  Making the pictures...' DarkGray
            $xl.ScreenUpdating = $true
            $cub = Get-WbSheet $wb $Script:CUBIC
            $p1 = Join-Path $outDir "Daily Ops All shops $tag.png"
            $p2 = Join-Path $outDir "Daily Ops Summary $tag.png"
            Save-RangePicture $xl $cub.Range($Script:ShotCubic) $p1
            Save-RangePicture $xl $sum.Range($Script:ShotSummary) $p2
            $done += $p2, $p1
        }
        if ($mode -eq 1 -or $mode -eq 3) {
            Say '  Saving the workbook...' DarkGray
            $saved = $null
            if (-not $readOnly -and (Test-SettingYes $cfg 'SaveOverOriginal')) {
                $keep = Get-SettingInt $cfg 'KeepBackups' 5
                if ($keep -gt 0) {
                    try {
                        $bdir = Join-Path $Script:Root 'Backups'
                        New-Item -ItemType Directory -Force -Path $bdir | Out-Null
                        $bn = [IO.Path]::GetFileNameWithoutExtension($mainPath) + ' ' + (Get-Date).ToString('yyyy-MM-dd HHmm') + [IO.Path]::GetExtension($mainPath)
                        Copy-Item -LiteralPath $mainPath -Destination (Get-UniquePath $bdir $bn)
                        $old = @(Get-ChildItem -LiteralPath $bdir -File | Where-Object { $_.Name.StartsWith([IO.Path]::GetFileNameWithoutExtension($mainPath) + ' ') } | Sort-Object LastWriteTime -Descending | Select-Object -Skip $keep)
                        foreach ($o in $old) { Remove-Item -LiteralPath $o.FullName -Force -ErrorAction SilentlyContinue }
                    }
                    catch { Warn "Couldn't make a backup copy: $($_.Exception.Message)" }
                }
                $wb.Save()
                $wb.Close($false); Remove-Com $wb; $wb = $null
                Say "  Copying it back to $(Split-Path -Parent $mainPath)..." DarkGray
                try { Copy-Item -LiteralPath $localMain -Destination $mainPath -Force; $saved = $mainPath }
                catch {
                    $saved = Get-UniquePath $outDir ([IO.Path]::GetFileNameWithoutExtension($mainPath) + " $tag" + [IO.Path]::GetExtension($mainPath))
                    Copy-Item -LiteralPath $localMain -Destination $saved -Force
                    Warn "Couldn't write over $mainPath - saved to output folder instead."
                }
            }
            else {
                $saved = Get-UniquePath $outDir ([IO.Path]::GetFileNameWithoutExtension($mainPath) + " $tag" + [IO.Path]::GetExtension($mainPath))
                $wb.SaveCopyAs($saved)
            }
            $done = @($saved) + $done
        }
        if ($mode -eq 1 -or $mode -eq 2) {
            $pageWb = $null; if ($mode -eq 1) { $pageWb = $saved }
            try { Publish-PageFolder $cfg $p2 $p1 $week $pageWb } catch { Warn "Updating the page failed: $($_.Exception.Message)" }
        }

        Head 'Done'
        foreach ($p in $done) { Say "  $p" Green }
        if ($mode -ne 3) { Say '  (the Summary picture is also on the clipboard - Ctrl+V into an email)' DarkGray }
        if ($Script:Warn.Count -gt 0) {
            Say ''
            Say "  $($Script:Warn.Count) thing(s) to check:" Yellow
            foreach ($w in $Script:Warn) { Say "   - $w" Yellow }
        }
        if (Test-SettingYes $cfg 'OpenWhenDone') {
            if ($mode -ne 3) { $openDir = $outDir }
            if ($mode -ne 2) { $openWb = $saved }
        }
    }
    finally {
        if ($null -ne $wb) { try { $wb.Close($false) } catch { } ; Remove-Com $wb }
        Stop-Excel $xl
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
    Open-WhenDone $openWb $openDir
}

# ---------- Daily Ops page on the J: drive ----------
# The page that goes in the folder (rewritten on every publish, so it stays current)
$Script:PageHtml = @'
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="refresh" content="180">
<title>MS Daily Report</title>
</head>
<body>
<p>MS Daily Report<br>
updated __UPDATED__</p>
__WORKBOOK__<p><a href="images/summary.png?v=__V__"><img src="images/summary.png?v=__V__" alt="Summary" style="max-width:100%"></a></p>
<p><a href="images/all-shops.png?v=__V__"><img src="images/all-shops.png?v=__V__" alt="All shops" style="max-width:100%"></a></p>
</body>
</html>
'@

# write a file next to its final name first, then swap it in - people looking at the page never get half a picture.
# Shared drives often allow writing but not deleting / renaming: then it just writes straight over the file.
function Write-FileSafe([string]$Path, [byte[]]$Bytes) {
    $tmp = $Path + '.tmp'
    try {
        [IO.File]::WriteAllBytes($tmp, $Bytes)
        if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($tmp, $Path, $null) } else { [IO.File]::Move($tmp, $Path) }
        return
    }
    catch { if (Test-Path -LiteralPath $tmp -PathType Leaf) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } }
    try {
        if (Test-Path -LiteralPath $Path) { try { (Get-Item -LiteralPath $Path).IsReadOnly = $false } catch { } }
        [IO.File]::WriteAllBytes($Path, $Bytes)
    }
    catch { throw "can't write $Path - $($_.Exception.Message)" }
}

# copies the 2 pictures + the page (with the time in it) into PageFolder (blank = off)
# the workbook the page links to: the one just saved, else the last one linked (Saved\page-workbook.txt).
# On the same drive as the page -> link straight to it; somewhere else (e.g. this PC) -> a copy goes in the page folder.
function Get-PageWorkbookLink([string]$Dir, [string]$Workbook) {
    $mem = Join-Path (Join-Path $Script:Root 'Saved') 'page-workbook.txt'
    if ($Workbook -and (Test-Path -LiteralPath $Workbook)) {
        try { New-Item -ItemType Directory -Force -Path (Split-Path -Parent $mem) | Out-Null; Set-Content -LiteralPath $mem -Value $Workbook } catch { }
    }
    elseif (Test-Path -LiteralPath $mem) { $Workbook = (Get-Content -LiteralPath $mem -Raw).Trim() }
    if (-not $Workbook -or -not (Test-Path -LiteralPath $Workbook)) { return '' }
    $target = $Workbook
    if ([IO.Path]::GetPathRoot($Workbook) -ine [IO.Path]::GetPathRoot($Dir)) {
        $target = Join-Path $Dir ('MS Daily Report' + [IO.Path]::GetExtension($Workbook))
        Say "  Copying the workbook next to the page..." DarkGray
        Write-FileSafe $target ([IO.File]::ReadAllBytes($Workbook))
    }
    $base = New-Object Uri ($Dir.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar)
    $rel = $base.MakeRelativeUri((New-Object Uri $target)).ToString()                    # relative, so a different drive letter still works
    $folderRel = $base.MakeRelativeUri((New-Object Uri ((Split-Path -Parent $target).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar))).ToString()
    if ($folderRel -eq '') { $folderRel = './' }
    $fileUri = (New-Object Uri $target).AbsoluteUri
    $js = { param($u) [Net.WebUtility]::HtmlEncode($u.Replace("'", '%27').Replace('\\', '/')) }   # safe inside onclick='...'
    $b = New-Object Text.StringBuilder
    [void]$b.Append('<p>')
    [void]$b.Append(('<button type="button" onclick="var a=document.createElement(''a'');a.href=''{0}'';a.download='''';document.body.appendChild(a);a.click();a.remove()">Download report</button> ' -f (& $js $rel)))
    [void]$b.Append(('<button type="button" onclick="location.href=''ms-excel:ofv|u|{0}''">Open in Excel</button> ' -f (& $js $fileUri)))
    [void]$b.Append(('<button type="button" onclick="location.href=''{0}''">Show folder</button>' -f (& $js $folderRel)))
    [void]$b.Append(('<br>{0}</p>' -f [Net.WebUtility]::HtmlEncode($target)))
    return $b.ToString() + "`n"
}

function Publish-PageFolder($Cfg, [string]$SummaryPng, [string]$ShopsPng, $Week, [string]$Workbook = '') {
    $dir = (Get-Setting $Cfg 'PageFolder').Trim()
    if ($dir -eq '') { return }
    try {
        $img = Join-Path $dir 'images'
        try { New-Item -ItemType Directory -Force -Path $img | Out-Null } catch { throw "can't make the folder $img - $($_.Exception.Message)" }
        Write-FileSafe (Join-Path $img 'summary.png') ([IO.File]::ReadAllBytes($SummaryPng))
        Write-FileSafe (Join-Path $img 'all-shops.png') ([IO.File]::ReadAllBytes($ShopsPng))
        $now = Get-Date
        # the time goes inside index.html itself (shared drives often block .js / .json files); written last,
        # so the time only changes once the pictures are there. ?v= makes browsers load the new pictures.
        $wbLink = ''
        try { $wbLink = Get-PageWorkbookLink $dir $Workbook } catch { Warn "Couldn't put the workbook link on the page: $($_.Exception.Message)" }
        $html = $Script:PageHtml.Replace('__UPDATED__', $now.ToString('MM/dd/yyyy, h:mm tt', $Script:US)).Replace('__V__', $now.ToString('yyyyMMddHHmmss', $Script:Inv)).Replace('__WORKBOOK__', $wbLink)
        Write-FileSafe (Join-Path $dir 'index.html') ([Text.Encoding]::UTF8.GetBytes($html))
        Say "  Page updated: $(Join-Path $dir 'index.html')" Green
    }
    catch { Warn "Couldn't update the page: $($_.Exception.Message)  (check you can make and change files in that folder, or use another folder: set PageFolder ...)" }
}

# 'page' command: puts the newest pictures from the output folder on the page without a run
function Invoke-PublishPage {
    $cfg = Get-Config
    if ((Get-Setting $cfg 'PageFolder') -eq '') { Say '  Set the folder first, e.g.  set PageFolder J:\Departments\Machine\4M\KACPER\DAILY OPS REPORTING\Daily Ops Page' Yellow; return }
    $root = Get-OutputRoot $cfg
    $sumP = @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'Daily Ops Summary *.png' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)[0]
    $shopP = @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'Daily Ops All shops *.png' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)[0]
    if ($null -eq $sumP -or $null -eq $shopP) { Say "  No pictures in $root yet - run the report first." Yellow; return }
    $week = $null; $m = [regex]::Match($sumP.Name, '(\d{4}-\d{2}-\d{2})'); if ($m.Success) { $week = Get-ReportWeek ([datetime]::ParseExact($m.Groups[1].Value, 'yyyy-MM-dd', $Script:Inv)) }
    Say "  $($sumP.Name) + $($shopP.Name)" DarkGray
    $Script:Warn.Clear()
    Publish-PageFolder $cfg $sumP.FullName $shopP.FullName $week
}

# Absenteeism % row on each Core sheet: negative -> 0%, 100% (no one entered as present) or no one scheduled -> N/A.
# Goes into the row's formulas (Actual days + week column), so it also holds when numbers are typed in later.
# Only the usual formulas are swapped:  =(Sched-Actual)/Sched  and  =IFERROR((Sched-Actual)/Sched,0);  week =AVERAGE(...) gets IFERROR(...,"N/A").
function Set-AbsenteeismCheck($Shops, $Core) {
    $changed = @()
    foreach ($s in $Shops) {
        $c = $Core[$s.Code]; if ($null -eq $c -or $null -eq $c.Ws) { continue }
        $g = $c.Grid; $last = Get-LastRow $g
        $r = Find-InCol $g 3 1 $last 'absente'                    # 'Absenteeism %' (also matches 'Absenteism')
        $p = Find-InCol $g 3 1 $last '# of employee present'
        if ($r -le 0 -or $p -le 0 -or $p -eq $r) { continue }
        $o = $p - $r
        $x = "(R[$o]C[-1]-R[$o]C)/R[$o]C[-1]"
        $new = "=IFERROR(IF($x>=1,`"N/A`",MAX(0,$x)),`"N/A`")"
        $old = @("=$x", "=IFERROR($x,0)") | ForEach-Object { ($_ -replace '\s', '').ToUpperInvariant() }
        $n = 0
        foreach ($col in @(8, 10, 12, 14, 16, 18, 20, 22)) {           # H J L N P R T (Actual) + V (week)
            $cell = $c.Ws.Cells.Item($r, $col)
            try {
                $cur = "$($cell.FormulaR1C1)"
                $k = ($cur -replace '\s', '').ToUpperInvariant()
                $put = $null
                if ($old -contains $k) { $put = $new }
                elseif ($col -eq 22 -and $k.StartsWith('=AVERAGE(')) { $put = '=IFERROR(' + $cur.Trim().Substring(1) + ',"N/A")' }
                if ($put) { $cell.FormulaR1C1 = [string]$put; $n++ }
            }
            finally { Remove-Com $cell }
        }
        if ($n -gt 0) { $changed += $s.Label }
    }
    if ($changed.Count -gt 0) { Say "  Absenteeism check added ($($changed -join ', ')): below 0 -> 0%, 100% -> N/A" DarkGray }
}

function Invoke-FormulaCheck($Wb, $Sum, [bool]$Fix) {
    $found = @()
    $fr = $Sum.Range('D9:S58').FormulaR1C1
    $fa = $Sum.Range('D9:S58').Formula
    $coreRows = [ordered]@{}
    for ($r = 9; $r -le 58; $r++) {
        if ($Sum.Rows.Item($r).Hidden) { continue }
        $row = @(); for ($j = 1; $j -le 16; $j++) { $row += "$($fr[($r - 8), $j])" }
        foreach ($x in (Get-RowFixes $row 4)) { $found += [pscustomobject]@{ Ws = $Sum; Sheet = $Script:SUMMARY; R = $r; C = $x.C; Kind = $x.Kind; Was = $x.Was; Fix = $x.Fix } }
        for ($j = 1; $j -le 16; $j++) {
            foreach ($m in [regex]::Matches("$($fa[($r - 8), $j])", "'?(Core Machining_[^'!]+)'?!\`$?[A-Z]{1,3}\`$?(\d+)")) {
                $k = $m.Groups[1].Value + '|' + $m.Groups[2].Value
                if (-not $coreRows.Contains($k)) { $coreRows[$k] = @($m.Groups[1].Value, [int]$m.Groups[2].Value) }
            }
        }
    }
    $sheets = @{}
    foreach ($k in $coreRows.Keys) {
        $name = $coreRows[$k][0]; $r = $coreRows[$k][1]
        if (-not $sheets.ContainsKey($name)) { $sheets[$name] = Get-WbSheet $Wb $name }
        $ws = $sheets[$name]; if ($null -eq $ws) { continue }
        $f = $ws.Range("G${r}:T${r}").FormulaR1C1
        $row = @(); for ($j = 1; $j -le 14; $j++) { $row += "$($f[1, $j])" }
        foreach ($x in (Get-RowFixes $row 7)) { $found += [pscustomobject]@{ Ws = $ws; Sheet = $name; R = $r; C = $x.C; Kind = $x.Kind; Was = $x.Was; Fix = $x.Fix } }
    }
    if ($found.Count -eq 0) { Say '  Formulas OK' DarkGray; return }
    $bySheet = [ordered]@{}
    foreach ($x in $found) {
        $addr = (Get-ColLetter ($x.C - 1)) + $x.R
        $was = $x.Was; if ($x.Kind -eq 'blank') { $was = 'blank' } elseif ($x.Kind -eq 'ref') { $was = '#REF!' } elseif ($x.Kind -eq 'formula') { $was = 'other formula' }
        if ($Fix) { $cell = $x.Ws.Cells.Item($x.R, $x.C); $cell.FormulaR1C1 = [string]$x.Fix; Remove-Com $cell }
        if (-not $bySheet.Contains($x.Sheet)) { $bySheet[$x.Sheet] = @() }
        $bySheet[$x.Sheet] += "$addr ($was)"
    }
    if ($Fix) { Say "  Formulas put back where they were typed over / broken ($($found.Count) cells):" Yellow }
    else { Warn "$($found.Count) cells break their row's formula:" }
    foreach ($k in $bySheet.Keys) { Say ("    {0}: {1}" -f $k, ($bySheet[$k] -join ', ')) DarkYellow }
}

function Set-CubicPivots($Wb, $Week) {
    $keep = @{ Year = @{}; Month = @{}; Week = @{}; Day = @{} }
    $last = $Week.LastDay; if ($last -gt $Week.WeekEnd) { $last = $Week.WeekEnd }
    for ($d = $Week.Monday; $d -le $last; $d = $d.AddDays(1)) {
        $keep.Year["$($d.Year)"] = $true; $keep.Month["$($d.Month)"] = $true
        $keep.Week["$(Get-WeekNum1 $d)"] = $true; $keep.Day["$($d.Day)"] = $true
    }
    $ws = Get-WbSheet $Wb $Script:CUBIC
    $pts = $ws.PivotTables()
    if ($pts.Count -lt 1) { Warn 'No pivot tables on the cubic-inch sheet.'; return }
    [void]$pts.Item(1).PivotCache().Refresh()
    foreach ($pt in $pts) {
        try {
            $pt.ManualUpdate = $true
            $wk = $pt.PivotFields('Week')
            if ($wk.Orientation -eq 0) { $wk.Orientation = 3 }
            $have = $true
            foreach ($fn in @('Year', 'Month', 'Week', 'Day')) { if (-not (Set-PivotItems $pt.PivotFields($fn) $keep[$fn])) { $have = $false } }
            if (-not $have) {
                $dayF = $pt.PivotFields('Day')
                [void]$dayF.ClearAllFilters()
                [void]$dayF.PivotFilters.Add2(15, [Type]::Missing, '#none#')
            }
            $pt.ManualUpdate = $false
        }
        catch { try { $pt.ManualUpdate = $false } catch { } ; Warn "Couldn't filter pivot '$($pt.Name)': $($_.Exception.Message)" }
    }
}

function Set-PivotItems($Pf, $Keep) {
    [void]$Pf.ClearAllFilters()
    $found = 0
    foreach ($pi in $Pf.PivotItems()) { if ($Keep.ContainsKey("$($pi.Name)")) { $found++ } }
    if ($found -eq 0) { return $false }
    if ($Pf.Orientation -eq 3) { $Pf.EnableMultiplePageItems = $true }
    foreach ($pi in $Pf.PivotItems()) { if (-not $Keep.ContainsKey("$($pi.Name)")) { $pi.Visible = $false } }
    return $true
}

#endregion

#region ---------- commands & REPL ----------

function Show-Shops($Cfg) {
    Head 'Shops'
    Say ('  {0,-8} {1,-26} {2,-10} {3,-6} {4,-8} {5}' -f 'Shop', 'Sheet', 'Tracker', 'Dept', 'Machine', 'Schedule depts') White
    foreach ($s in $Cfg.Shops) {$sd = 'all'; if ($s.SchedOnly) {$sd = "only $($s.SchedOnly)" } elseif ($s.SchedExcept) {$sd = "all but $($s.SchedExcept)" }
        $m = ''; if ($s.Machine -gt 0) { $m =$s.Machine }
        Say ('  {0,-8} {1,-26} {2,-10} {3,-6} {4,-8} {5}' -f $s.Label, $s.Sheet, $s.TrackerTab, $s.Dept, $m,$sd)
    }
}
function Show-Settings($Cfg) {
    Head 'Settings'
    foreach ($k in $Cfg.Settings.Keys) {
        $v = $Cfg.Settings[$k]
        if ($v -eq '') { $v = '(blank)' }
        Say ('  {0,-17} {1}' -f $k, $v) White
        if ($Script:DefaultSettings.Contains($k)) { 
            Say ('  {0,-17} {1}' -f '', $Script:DefaultSettings[$k][1]) DarkGray 
        }
    }
    Say ('  Downloads now: ' + (Get-DownloadsFolder)) DarkGray
    Say ('  Output now:    ' + (Get-OutputRoot $Cfg)) DarkGray
}

function Invoke-Set([string[]]$ArgList) {
    $cfg = Get-Config
    if ($ArgList.Count -lt 1) { 
        Show-Settings $cfg
        return 
    }
    $key = $null
    foreach ($k in $cfg.Settings.Keys) { 
        if ($k -ieq $ArgList[0]) { $key = $k } 
    }
    if ($null -eq $key) { 
        Say "  Unknown setting '$($ArgList[0])'." Yellow
        return 
    }
    if ($ArgList.Count -ge 2) { 
        $val = ($ArgList[1..($ArgList.Count - 1)] -join ' ') 
    } else { 
        $val = Ask "New value for $key" $cfg.Settings[$key] 
    }
    if ($val -eq '""' -or $val -eq "''") { $val = '' }
    $cfg.Settings[$key] = (Clean-PathText $val)
    Save-Config $cfg
    Say "  $key = $($cfg.Settings[$key])" Green
}

function Show-Typed {
    $t = Get-Typed
    Head 'Typed values kept for re-runs'
    if ($t.Start.Count -eq 0 -and $t.Absent.Count -eq 0) { 
        Say '  (none)' DarkGray
        return 
    }
    foreach ($k in ($t.Start.Keys | Sort-Object)) { 
        Say ('  start   {0,-22} {1}' -f $k, $t.Start[$k]) 
    }
    foreach ($k in ($t.Absent.Keys | Sort-Object)) { 
        Say ('  absent  {0,-22} {1}' -f $k, ((@($t.Absent[$k]) | ForEach-Object { if ($null -eq $_) { '.' } else { $_ } }) -join ' ')) 
    }
}
function Show-Help {
    Head 'Commands'
    $rows = @(
        @('run [files]', 'Daily update', 'r'),
        @('pictures', 'Just make the 2 pictures', 'pics'),
        @('shops', 'Show shops', ''),
        @('settings', 'Show settings', ''),
        @('set <name> <value>', 'Change a setting', ''),
        @('typed', 'Show typed start times/absences', ''),
        @('forget', 'Clear typed start times/absences', ''),
        @('open config|output|report|backups', 'Open folder', ''),
        @('outlook [date]', 'Try the Outlook scan for the hours emails (report date, default today)', ''),
        @('page', 'Put the newest 2 pictures on the J: page now', ''),
        @('help', 'This list', '?'),
        @('exit', 'Quit', 'q')
    )
    foreach ($r in $rows) {$alias = ''; if ($r[2]) {$alias = "  [$($r[2])]" }
        Say ('  {0,-34} {1}{2}' -f $r[0], $r[1],$alias)
    }
}
function Invoke-OpsCommand([string]$Line) {
    $tokens = @(Split-PathList $Line)
    if ($tokens.Count -eq 0) { return }
    $cmd = $tokens[0].ToLower()
    $rest = @()
    if ($tokens.Count -gt 1) { $rest = @($tokens[1..($tokens.Count - 1)]) }
    $restText = ($rest -join ' ')
    $alias = @{ 
        r = 'run'; e = 'run'; export = 'run'; go = 'run'; 
        pics = 'pictures'; pic = 'pictures'; images = 'pictures'; 
        '?' = 'help'; h = 'help'; q = 'exit'; quit = 'exit'; 
        setting = 'settings'; config = 'settings'; clear = 'forget' 
    }
    if ($alias.ContainsKey($cmd)) { $cmd = $alias[$cmd] }

    if ($cmd -ne 'run' -and [IO.Path]::IsPathRooted((Clean-PathText $tokens[0])) -and (Test-Path -LiteralPath (Clean-PathText $tokens[0]) -ErrorAction SilentlyContinue)) { 
        Invoke-Report @($tokens | ForEach-Object { Clean-PathText $_ })
        return 
    }

    switch ($cmd) {
        'run' { Invoke-Report @($rest | ForEach-Object { Clean-PathText $_ } | Where-Object { Test-Path -LiteralPath $_ }) }
        'pictures' { Invoke-Report @($rest | ForEach-Object { Clean-PathText $_ } | Where-Object { Test-Path -LiteralPath $_ }) -PicturesOnly }
        'page' { Invoke-PublishPage }
        'outlook' {
            $cfg = Get-Config
            $rd = (Get-Date).Date; if ($restText) { $x = ConvertTo-DateNear $restText; if ($x) { $rd = $x } }
            $wk = Get-ReportWeek $rd
            Head "Outlook scan - report week $(Format-Day $wk.Monday) - $(Format-Day $wk.WeekEnd)"
            $tmp = Join-Path ([IO.Path]::GetTempPath()) ('DailyOps_ol_' + [guid]::NewGuid().ToString('N').Substring(0, 6))
            try {
                $files = @(Get-OutlookHours $cfg $wk $tmp (Get-SettingInt $cfg 'DaysToLook' 14))
                foreach ($h in (Get-HoursCandidates $files $wk.Monday $wk.WeekEnd)) {
                    if ($h.Use) { Say ('  {0}  {1}  ({2})' -f (Format-Day $h.Date), $h.Name, $Script:Origin[$h.Path]) Green }
                    else { Say ('  --         {0}  ({1})  - {2}' -f $h.Name, $Script:Origin[$h.Path], $h.Note) DarkGray }
                }
            }
            finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
        }
        'shops' { Show-Shops (Get-Config) }
        'settings' { Show-Settings (Get-Config) }
        'set' { Invoke-Set $rest }
        'typed' { Show-Typed }
        'forget' { 
            if (AskYesNo 'Clear all typed start times and absences?' $false) { 
                Remove-Item -LiteralPath $Script:TypedPath -Force -ErrorAction SilentlyContinue
                Say '  Cleared.' Green 
            } 
        }
        'open' {
            $what = "$restText".ToLower()
            $cfg = Get-Config
            if ($what -like 'out*' -or $what -like 'pic*') { 
                $o = Get-OutputRoot $cfg
                New-Item -ItemType Directory -Force -Path $o | Out-Null
                Invoke-Item -LiteralPath $o 
            }
            elseif ($what -like 'rep*' -or $what -like 'work*') { 
                Invoke-Item -LiteralPath (Get-Folder $cfg 'ReportFolder') 
            }
            elseif ($what -like 'back*') { 
                $o = Join-Path $Script:Root 'Backups'
                New-Item -ItemType Directory -Force -Path $o | Out-Null
                Invoke-Item -LiteralPath $o 
            }
            else { 
                Invoke-Item -LiteralPath $Script:ConfigPath 
            }
        }
        'help' { Show-Help }
        'exit' { return 'exit' }
        default { Say "  Unknown command '$cmd'. Type 'help'." Yellow }
    }
}

function Start-Repl {
    Say ''
    Say '  CHICAGO DAILY OPS REPORT  -  test version: start times from raw data + trackers' Cyan
    Say '  Press Enter to run the daily update, or type a command.' DarkGray
    while ($true) {
        Write-Host ''
        $line = Read-Host 'daily ops'
        if ($null -eq $line) { break }
        $line = $line.Trim()
        if ($line -eq '') { $line = 'run' }
        try { if ((Invoke-OpsCommand $line) -eq 'exit') { break } }
        catch [OperationCanceledException] { Say '  Cancelled.' Yellow }
        catch { Show-Error $_ }
    }
}

# ---------- entry point ----------
if ($MyInvocation.InvocationName -ne '.') {
    if ($CmdArgs -and $CmdArgs.Count -gt 0) {
        $paths = @($CmdArgs | Where-Object { Test-Path -LiteralPath $_ -ErrorAction SilentlyContinue })
        try {
            if ($paths.Count -eq $CmdArgs.Count) { Invoke-Report $paths }
            else { [void](Invoke-OpsCommand ($CmdArgs -join ' ')) }
        }
        catch [OperationCanceledException] { Say '  Cancelled.' Yellow }
        catch { Show-Error $_ }
        if ($paths.Count -eq $CmdArgs.Count) { Start-Repl }
    }
    else { Start-Repl }
}