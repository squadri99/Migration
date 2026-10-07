#Requires -Version 5.1

trap {
    $m = "pnfp Migration Tool failed to start.`n`n" +
         "Error : $($_.Exception.Message)`n" +
         "Type  : $($_.Exception.GetType().FullName)`n" +
         "Line  : $($_.InvocationInfo.ScriptLineNumber)"
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show($m,"Startup Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } catch { }
    Write-Host $m -ForegroundColor Red
    break
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------
# SCRIPT DIRECTORY AND NOTES FOLDER
# ---------------------------------------------------------------

# Single output folder on the user's Desktop, created once if missing.
# All outputs (notes, scripts, connection captures) are saved directly in this folder.
$script:AutoMigDir = ""
try {
    $desktop = [Environment]::GetFolderPath('Desktop')
    if($desktop){ $script:AutoMigDir = Join-Path $desktop "PNFP-AUTO_MIGRATION" }
} catch { }
if(-not $script:AutoMigDir){ $script:AutoMigDir = Join-Path $env:USERPROFILE "Desktop\PNFP-AUTO_MIGRATION" }
if (-not (Test-Path $script:AutoMigDir)) {
    try { New-Item -ItemType Directory -Path $script:AutoMigDir -Force | Out-Null } catch { }
}

$script:ScriptDir   = $script:AutoMigDir
$script:NotesFolder = $script:AutoMigDir

# ---------------------------------------------------------------
# THEME
# ---------------------------------------------------------------
$Clr = @{
    BG     = [System.Drawing.Color]::FromArgb(18,  20,  28)
    Panel  = [System.Drawing.Color]::FromArgb(26,  30,  42)
    Card   = [System.Drawing.Color]::FromArgb(32,  37,  52)
    Border = [System.Drawing.Color]::FromArgb(45,  52,  72)
    Input  = [System.Drawing.Color]::FromArgb(24,  28,  40)
    Text   = [System.Drawing.Color]::FromArgb(220, 225, 235)
    Dim    = [System.Drawing.Color]::FromArgb(120, 130, 155)
    Green  = [System.Drawing.Color]::FromArgb(60,  210, 120)
    Blue   = [System.Drawing.Color]::FromArgb(64,  156, 255)
    Red    = [System.Drawing.Color]::FromArgb(255,  80,  80)
    Yellow = [System.Drawing.Color]::FromArgb(255, 200,  60)
    Orange = [System.Drawing.Color]::FromArgb(255, 150,  50)
    Purple = [System.Drawing.Color]::FromArgb(180, 120, 255)
}
$FontTitle = New-Object System.Drawing.Font("Segoe UI",13,[System.Drawing.FontStyle]::Bold)
$FontUI    = New-Object System.Drawing.Font("Segoe UI", 9)
$FontUIB   = New-Object System.Drawing.Font("Segoe UI", 9,[System.Drawing.FontStyle]::Bold)
$FontMono  = New-Object System.Drawing.Font("Consolas", 9)
$FontSm    = New-Object System.Drawing.Font("Segoe UI", 8)
$FontSec   = New-Object System.Drawing.Font("Segoe UI", 9,[System.Drawing.FontStyle]::Bold)
$FontLg    = New-Object System.Drawing.Font("Segoe UI",11,[System.Drawing.FontStyle]::Bold)

# ---------------------------------------------------------------
# STATE
# ---------------------------------------------------------------
$script:Mode          = ""          # "PRE" or "POST"
$script:NotesFile     = ""          # path to auto-saved notes file
$script:SrcConn       = $null       # SqlConnection - source
$script:DstConn       = $null       # SqlConnection - destination
$script:SrcCS         = ""
$script:DstCS         = ""
$script:Running       = $false
$script:Cancel        = $false

# ---------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------
function Save-Notes {
    try {
        if(-not $script:NotesFile){ return }
        $TxtNotes.Text | Out-File -FilePath $script:NotesFile -Encoding UTF8 -Force
    } catch { }
}

function Load-Notes {
    try {
        if($script:NotesFile -and (Test-Path $script:NotesFile)){
            $TxtNotes.Text = [System.IO.File]::ReadAllText($script:NotesFile, [System.Text.Encoding]::UTF8)
        }
    } catch { }
}
function Write-Log {
    param([string]$Msg, [object]$Color = $null)
    $col = if($Color){ $Color } else { $Clr.Text }
    $RtbLog.SelectionStart  = $RtbLog.TextLength
    $RtbLog.SelectionColor  = $col
    $ts = Get-Date -Format "HH:mm:ss"
    $RtbLog.AppendText("[$ts]  $Msg`r`n")
    $RtbLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Write-Sep {
    param([string]$Title)
    Write-Log ("=" * 70) $Clr.Border
    Write-Log "  $Title" $Clr.Blue
    Write-Log ("=" * 70) $Clr.Border
}

function Set-Status {
    param([string]$T, [object]$C = $null)
    $SLbl.Text      = $T
    $SLbl.ForeColor = if($C){ $C } else { $Clr.Dim }
    [System.Windows.Forms.Application]::DoEvents()
}

function Build-CS {
    param([string]$Server,[bool]$SqlAuth,[string]$User,[string]$Pass,[string]$DB="master")
    $cb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $cb["Data Source"]    = $Server
    $cb["Initial Catalog"]= $DB
    $cb["Connect Timeout"]= 30
    if($SqlAuth){ $cb["User ID"]=$User; $cb["Password"]=$Pass }
    else { $cb["Integrated Security"]=$true }
    return $cb.ConnectionString
}

function Open-Conn {
    param([string]$CS)
    $cn = New-Object System.Data.SqlClient.SqlConnection $CS
    $cn.Open()
    return $cn
}

function Exec-Sql {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$Sql,[int]$Timeout=3600)
    $cmd = New-Object System.Data.SqlClient.SqlCommand $Sql,$Conn
    $cmd.CommandTimeout = $Timeout
    $cmd.ExecuteNonQuery() | Out-Null
}

function Query-Scalar {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$Sql)
    $cmd = New-Object System.Data.SqlClient.SqlCommand $Sql,$Conn
    $cmd.CommandTimeout = 60
    return $cmd.ExecuteScalar()
}

function Query-Table {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$Sql)
    $cmd = New-Object System.Data.SqlClient.SqlCommand $Sql,$Conn
    $cmd.CommandTimeout = 120
    $r   = $cmd.ExecuteReader()
    $dt  = New-Object System.Data.DataTable
    $dt.Load($r)
    if(-not $r.IsClosed){ $r.Close() }
    # The leading comma stops PowerShell from unrolling the DataTable into loose DataRow objects
    return ,$dt
}

function Get-SqlVersion {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $v = Query-Scalar $Conn "SELECT SERVERPROPERTY('ProductMajorVersion')"
    return [int]$v
}

# Compat level for the latest SQL version detected on destination
function Get-LatestCompatLevel {
    param([int]$MajorVersion)
    switch ($MajorVersion) {
        17 { return 170 }   # SQL 2025
        16 { return 160 }   # SQL 2022
        15 { return 150 }   # SQL 2019
        14 { return 140 }   # SQL 2017
        13 { return 130 }   # SQL 2016
        12 { return 120 }   # SQL 2014
        11 { return 110 }   # SQL 2012
        default { return 150 }
    }
}

# ---------------------------------------------------------------
# CONNECTION CAPTURE (AIDO - Extended Events)
# ---------------------------------------------------------------
function Select-CaptureTarget {
    param($SrcConn, $DstConn)

    $f = New-Object System.Windows.Forms.Form
    $f.Text            = "Run Connection Capture"
    $f.StartPosition   = 'CenterScreen'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox     = $false
    $f.MinimizeBox     = $false
    $f.ShowInTaskbar   = $false
    $f.ClientSize      = New-Object System.Drawing.Size(420,130)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = "Which server should the connection capture query run against?"
    $lbl.Location = New-Object System.Drawing.Point(15,15)
    $lbl.Size     = New-Object System.Drawing.Size(390,35)
    $f.Controls.Add($lbl)

    $btnSrc = New-Object System.Windows.Forms.Button
    $btnSrc.Location = New-Object System.Drawing.Point(15,60)
    $btnSrc.Size     = New-Object System.Drawing.Size(125,50)
    $btnSrc.Text     = if($SrcConn){ "Source`n$($SrcConn.DataSource)" } else { "Source`n(not connected)" }
    $btnSrc.Enabled  = [bool]$SrcConn
    $btnSrc.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $f.Controls.Add($btnSrc)

    $btnDst = New-Object System.Windows.Forms.Button
    $btnDst.Location = New-Object System.Drawing.Point(150,60)
    $btnDst.Size     = New-Object System.Drawing.Size(125,50)
    $btnDst.Text     = if($DstConn){ "Destination`n$($DstConn.DataSource)" } else { "Destination`n(not connected)" }
    $btnDst.Enabled  = [bool]$DstConn
    $btnDst.DialogResult = [System.Windows.Forms.DialogResult]::No
    $f.Controls.Add($btnDst)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(285,60)
    $btnCancel.Size     = New-Object System.Drawing.Size(120,50)
    $btnCancel.Text     = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($btnCancel)
    $f.CancelButton = $btnCancel

    $res = $f.ShowDialog()
    $f.Dispose()
    switch ($res) {
        'Yes' { return 'Source' }
        'No'  { return 'Destination' }
        default { return $null }
    }
}

function Invoke-ConnectionCapture {
    $srcOpen = ($script:SrcConn -and $script:SrcConn.State -eq 'Open')
    $dstOpen = ($script:DstConn -and $script:DstConn.State -eq 'Open')

    if(-not $srcOpen -and -not $dstOpen){
        [System.Windows.Forms.MessageBox]::Show(
            "Connect to the source or destination server first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $target = Select-CaptureTarget -SrcConn $(if($srcOpen){ $script:SrcConn }) -DstConn $(if($dstOpen){ $script:DstConn })
    if(-not $target){ return }
    $conn = if($target -eq 'Source'){ $script:SrcConn } else { $script:DstConn }

    Write-Sep "CONNECTION CAPTURE - AIDO ($target)"
    Set-Status "Running connection capture from AIDO ($target)..." $Clr.Yellow

    # Output goes straight into the single Desktop folder created at startup
    $connFolder = $script:ScriptDir
    if(-not $connFolder -or -not (Test-Path $connFolder)){
        Write-Log "Output folder not found: $connFolder" $Clr.Red
        return
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"

    # ---- Run the date range query ----
    $rangeFrom = ""; $rangeTo = ""
    try {
        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.Connection     = $conn
        $cmd.CommandTimeout = 120
        $cmd.CommandText    = @"
USE AIDO;
SELECT MIN(timestamp) date_range_from, MAX(timestamp) date_range_to
FROM AIDO.dbo.Extended_Events_Capture_Logins;
"@
        $rdr = $cmd.ExecuteReader()
        if($rdr.Read()){
            $rangeFrom = if($rdr["date_range_from"] -ne [DBNull]::Value){ "$($rdr['date_range_from'])" } else { "unknown" }
            $rangeTo   = if($rdr["date_range_to"]   -ne [DBNull]::Value){ "$($rdr['date_range_to'])" }   else { "unknown" }
        }
        if(-not $rdr.IsClosed){ $rdr.Close() }
        Write-Log "  Date range: $rangeFrom  to  $rangeTo" $Clr.Dim
        if($rangeFrom -eq "unknown" -and $rangeTo -eq "unknown"){ Write-Log "  Table is empty on $($conn.DataSource). Check that you picked the server running the XE capture." $Clr.Yellow }
    } catch {
        Write-Log "  Warning: could not read date range - $($_.Exception.Message)" $Clr.Yellow
    }

    # ---- Run the main connections query ----
    $dt = New-Object System.Data.DataTable
    try {
        $cmd2 = New-Object System.Data.SqlClient.SqlCommand
        $cmd2.Connection     = $conn
        $cmd2.CommandTimeout = 300
        $cmd2.CommandText    = @"
USE AIDO;
SET NOCOUNT ON;
SELECT @@SERVERNAME                AS server_name,
       client_hostname,
       server_principal_name,
       database_name,
       COUNT(*)                    AS connection_count
FROM AIDO.dbo.Extended_Events_Capture_Logins
WHERE 1=1
  AND server_principal_name <> 'sareports'
  AND server_principal_name <> 'SNV\13017'
  AND server_principal_name <> 'SNV\68604'
  AND server_principal_name <> 'SNV\svc_Redgate_service'
  AND server_principal_name <> 'SNV\13055'
  AND server_principal_name <> 'SNV\02837'
  AND server_principal_name <> 'SNV\19776'
GROUP BY client_hostname,
         server_principal_name,
         database_name
ORDER BY database_name;
"@
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd2
        $null = $adapter.Fill($dt)
        Write-Log "  $($dt.Rows.Count) connection row(s) returned." $Clr.Green
    } catch {
        Write-Log "Connection capture query failed: $($_.Exception.Message)" $Clr.Red
        return
    }

    if($dt.Rows.Count -eq 0){
        Write-Log "No connection data found in AIDO.dbo.Extended_Events_Capture_Logins." $Clr.Yellow
        return
    }

    # ---- Write CSV ----
    $csvPath = Join-Path $connFolder "Connections_$stamp.csv"
    try {
        $lines = @()
        $headers = ($dt.Columns | ForEach-Object { $_.ColumnName }) -join ","
        $lines += $headers
        foreach($row in $dt.Rows){
            $vals = @()
            foreach($col in $dt.Columns){
                $v = "$($row[$col.ColumnName])" -replace '"','""'
                $vals += if($v -match '[,"]'){ "`"$v`"" } else { $v }
            }
            $lines += $vals -join ","
        }
        $lines | Out-File -FilePath $csvPath -Encoding UTF8 -Force
        Write-Log "  CSV saved : $csvPath" $Clr.Green
    } catch {
        Write-Log "  CSV save failed: $($_.Exception.Message)" $Clr.Red
    }

    # ---- Write XLSX (no external library - build Open XML manually) ----
    $xlsxPath = Join-Path $connFolder "Connections_$stamp.xlsx"
    try {
        Add-Type -AssemblyName WindowsBase -ErrorAction Stop

        $pkg = New-Object System.IO.Packaging.Package
        # Use reflection to call Open with write access
        $xlsxStream = [System.IO.File]::Open($xlsxPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite)
        $pkg = [System.IO.Packaging.Package]::Open($xlsxStream, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite)

        function Add-Part {
            param($Pkg,[string]$Uri,[string]$ContentType,[string]$Xml)
            $partUri  = New-Object System.Uri($Uri, [System.UriKind]::Relative)
            $part     = $Pkg.CreatePart($partUri, $ContentType, [System.IO.Packaging.CompressionOption]::Normal)
            $stream   = $part.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
            $bytes    = [System.Text.Encoding]::UTF8.GetBytes($Xml)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Close()
            return $partUri
        }

        function Escape-Xml { param([string]$s)
            $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'","&apos;"
        }

        # Shared strings
        $strings = [System.Collections.Generic.List[string]]::new()
        $strIdx  = @{}
        function Get-StrIdx { param([string]$s)
            if(-not $strIdx.ContainsKey($s)){ $strIdx[$s]=$strings.Count; $strings.Add($s) }
            return $strIdx[$s]
        }

        # Build sheet rows
        $sheetRows = New-Object System.Text.StringBuilder
        $rowNum = 1

        # Header row (bold via style 1)
        [void]$sheetRows.Append("<row r=`"$rowNum`">")
        $colIdx = 0
        foreach($col in $dt.Columns){
            $ci = Get-StrIdx (Escape-Xml $col.ColumnName)
            $addr = "$(([char](65+$colIdx)))$rowNum"
            [void]$sheetRows.Append("<c r=`"$addr`" t=`"s`" s=`"1`"><v>$ci</v></c>")
            $colIdx++
        }
        [void]$sheetRows.Append("</row>")
        $rowNum++

        # Data rows
        foreach($row in $dt.Rows){
            [void]$sheetRows.Append("<row r=`"$rowNum`">")
            $colIdx = 0
            foreach($col in $dt.Columns){
                $v = "$($row[$col.ColumnName])"
                $addr = "$(([char](65+$colIdx)))$rowNum"
                $isNum = $v -match '^-?\d+(\.\d+)?$'
                if($isNum){
                    [void]$sheetRows.Append("<c r=`"$addr`"><v>$v</v></c>")
                } else {
                    $ci = Get-StrIdx (Escape-Xml $v)
                    [void]$sheetRows.Append("<c r=`"$addr`" t=`"s`"><v>$ci</v></c>")
                }
                $colIdx++
            }
            [void]$sheetRows.Append("</row>")
            $rowNum++
        }

        # Meta row: date range info
        $rowNum++
        [void]$sheetRows.Append("<row r=`"$rowNum`">")
        $ci = Get-StrIdx "Date range: $rangeFrom to $rangeTo"
        [void]$sheetRows.Append("<c r=`"A$rowNum`" t=`"s`"><v>$ci</v></c>")
        [void]$sheetRows.Append("</row>")

        # Shared strings XML
        $ssXml  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        $ssXml += "<sst xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" count=`"$($strings.Count)`" uniqueCount=`"$($strings.Count)`">"
        foreach($s in $strings){ $ssXml += "<si><t xml:space=`"preserve`">$s</t></si>" }
        $ssXml += "</sst>"

        # Styles (bold for header)
        $stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        $stylesXml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        $stylesXml += '<fonts><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        $stylesXml += '<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        $stylesXml += '<borders><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        $stylesXml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        $stylesXml += '<cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        $stylesXml += '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        $stylesXml += '</styleSheet>'

        # Sheet XML
        $sheetXml  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        $sheetXml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        $sheetXml += '<sheetData>' + $sheetRows.ToString() + '</sheetData></worksheet>'

        # Workbook XML
        $wbXml  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        $wbXml += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        $wbXml += '<sheets><sheet name="Connections" sheetId="1" r:id="rId1"/></sheets></workbook>'

        # Relationships
        $wbRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        $wbRelsXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        $wbRelsXml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        $wbRelsXml += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        $wbRelsXml += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        $wbRelsXml += '</Relationships>'

        $pkgRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        $pkgRelsXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        $pkgRelsXml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        $pkgRelsXml += '</Relationships>'

        $ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

        Add-Part $pkg "/xl/worksheets/sheet1.xml"   "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" $sheetXml  | Out-Null
        Add-Part $pkg "/xl/sharedStrings.xml"        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" $ssXml | Out-Null
        Add-Part $pkg "/xl/styles.xml"               "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" $stylesXml     | Out-Null
        Add-Part $pkg "/xl/workbook.xml"             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" $wbXml      | Out-Null
        Add-Part $pkg "/xl/_rels/workbook.xml.rels"  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" $wbRelsXml  | Out-Null
        Add-Part $pkg "/_rels/.rels"                 "application/vnd.openxmlformats-package.relationships+xml" $pkgRelsXml                   | Out-Null

        $pkg.Close()
        $xlsxStream.Close()
        Write-Log "  XLSX saved: $xlsxPath" $Clr.Green
    } catch {
        Write-Log "  XLSX save failed: $($_.Exception.Message)" $Clr.Red
        Write-Log "  CSV was saved successfully and contains the same data." $Clr.Dim
    }

    Write-Log "Connection capture complete. Files in: $connFolder" $Clr.Green
    Set-Status "Connection capture complete." $Clr.Green

    $open = [System.Windows.Forms.MessageBox]::Show(
        "Connection data saved to:`n$connFolder`n`nOpen folder?",
        "Connection Capture Complete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Information)
    if($open -eq [System.Windows.Forms.DialogResult]::Yes){
        Start-Process explorer.exe -ArgumentList "`"$connFolder`""
    }
}

# ---------------------------------------------------------------
# SCRIPT OUT
# Asks the same questions as Run, then writes the T-SQL for the real pipeline without executing it:
#   - Source script       : copy-only backups (run on the SOURCE server)
#   - Destination script  : restore, post-restore steps (run on the DESTINATION server)
#   - AG scripts          : primary / secondary / rollback (only when the AG option is chosen)
#   - Login script        : same file Run creates (only when the source is connected)
# ---------------------------------------------------------------
# T-SQL for the post-restore steps of ONE database (statistics, compatibility level, orphaned users),
# in the same order the pipeline uses for the mode ("PRE": stats, compat, orphans / "POST": compat, stats, orphans).
function Get-PostRestoreSql {
    param([string]$Target,[string]$Mode,$CompatTarget)
    $tq = ConvertTo-SqlBracket $Target
    $head = "-- ============================================================`n-- [$Target]  post-restore steps`n-- ============================================================`n"
    $compatSql = ""
    if($null -ne $CompatTarget){
        $compatSql = "-- Set compatibility level`nALTER DATABASE [$tq] SET COMPATIBILITY_LEVEL = $CompatTarget;`nGO`n`n"
    }
    $statsSql = "-- Update statistics`nUSE [$tq];`nGO`nEXEC sp_updatestats;`nGO`nUSE master;`nGO`n`n"
    $orphanSql = @"
-- Find and fix orphaned users (relinks each user to the login with the same name)
USE [$tq];
GO
DECLARE @fix nvarchar(max) = N'';
SELECT @fix = @fix + N'ALTER USER ' + QUOTENAME(dp.name) + N' WITH LOGIN = ' + QUOTENAME(dp.name) + N';' + CHAR(13) + CHAR(10)
FROM sys.database_principals dp
WHERE dp.type IN ('S','U','G')
  AND dp.name NOT IN ('dbo','guest','INFORMATION_SCHEMA','sys')
  AND dp.sid IS NOT NULL AND dp.sid <> 0x00
  AND NOT EXISTS (SELECT 1 FROM sys.server_principals sp WHERE sp.sid = dp.sid)
  AND EXISTS (SELECT 1 FROM sys.server_principals sp2 WHERE sp2.name = dp.name AND sp2.type IN ('S','U','G'));
PRINT @fix;
IF LEN(@fix) > 0 EXEC sys.sp_executesql @fix;
GO
-- Orphans that still have no login on this server (create the login, then run the block above again):
SELECT dp.name AS orphaned_user, dp.type_desc
FROM sys.database_principals dp
WHERE dp.type IN ('S','U','G')
  AND dp.name NOT IN ('dbo','guest','INFORMATION_SCHEMA','sys')
  AND dp.sid IS NOT NULL AND dp.sid <> 0x00
  AND NOT EXISTS (SELECT 1 FROM sys.server_principals sp WHERE sp.sid = dp.sid);
GO
USE master;
GO


"@
    if($Mode -eq "PRE"){ return ($head + $statsSql + $compatSql + $orphanSql) }
    return ($head + $compatSql + $statsSql + $orphanSql)
}

# T-SQL for the migration-day checks that follow the restores: PAGE_VERIFY CHECKSUM and sa as database owner.
function Get-MigrationDayFinalSql {
    param([string[]]$Targets)
    $sb = New-Object System.Text.StringBuilder
    $txt = @"
-- ============================================================
-- PAGE_VERIFY CHECKSUM on all user databases that are not already using it
-- ============================================================
DECLARE @pv nvarchar(max) = N'';
SELECT @pv = @pv + N'ALTER DATABASE ' + QUOTENAME(name) + N' SET PAGE_VERIFY CHECKSUM;' + CHAR(13) + CHAR(10)
FROM sys.databases
WHERE database_id > 4 AND source_database_id IS NULL AND is_distributor = 0
  AND page_verify_option_desc <> 'CHECKSUM' AND state_desc = 'ONLINE' AND is_read_only = 0;
PRINT @pv;
IF LEN(@pv) > 0 EXEC sys.sp_executesql @pv;
GO

-- ============================================================
-- Make sa the owner of the restored databases
-- ============================================================


"@
    [void]$sb.Append($txt)
    foreach($t in $Targets){
        $txt = @"
DECLARE @own nvarchar(max) = N'ALTER AUTHORIZATION ON DATABASE::' + QUOTENAME(N'$(ConvertTo-SqlLiteral $t)') + N' TO ' + QUOTENAME((SELECT name FROM sys.server_principals WHERE principal_id = 1)) + N';';
EXEC sys.sp_executesql @own;
GO


"@
        [void]$sb.Append($txt)
    }
    return $sb.ToString()
}

# Run-time: the post-restore steps for every restored database, run AFTER all databases have been restored.
function Invoke-PostRestoreSteps {
    param([string[]]$Databases,[string]$Mode,$CompatTarget)
    foreach($name in @($Databases | Where-Object { $_ })){
        if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }
        Write-Log "[$name]" $Clr.Blue
        $stepStats = { 
            Write-Log "  Update statistics on [$name]" $Clr.Blue
            try { Update-Statistics $script:DstConn $name }
            catch { Write-Log "  UPDATE STATS FAILED: $($_.Exception.Message)" $Clr.Red }
        }
        $stepCompat = {
            if($null -ne $CompatTarget){
                Write-Log "  Setting compatibility level to $CompatTarget on [$name]" $Clr.Blue
                try { Set-CompatibilityLevel $script:DstConn $name $CompatTarget }
                catch { Write-Log "  COMPAT LEVEL FAILED: $($_.Exception.Message)" $Clr.Red }
            }
        }
        if($Mode -eq "PRE"){ & $stepStats; if($script:Cancel){ break }; & $stepCompat }
        else               { & $stepCompat; if($script:Cancel){ break }; & $stepStats }
        if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

        Write-Log "  Find and fix orphaned users in [$name]" $Clr.Blue
        try { Fix-OrphanedUsers $script:DstConn $name }
        catch { Write-Log "  ORPHAN FIX FAILED: $($_.Exception.Message)" $Clr.Red }
        Write-Log "  [$name] post-restore steps complete." $Clr.Green
    }
}

# Run-time, restore left NORECOVERY / STANDBY: the databases cannot take these steps yet, so write them
# to a script file (full set, same SQL as Script Out) to run once the databases have been recovered.
function Save-DeferredPostRestoreScript {
    param([string[]]$Databases,[string]$Mode,$CompatTarget,[string]$RecoveryMode)
    $Databases = @($Databases | Where-Object { $_ })
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $path  = Join-Path (Get-JobFileFolder) ("pnfp_PostRestoreSteps_{0}_{1}.sql" -f (Get-SafeFileName (Get-DstFolderName)), $stamp)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((New-ScriptHeader -Title "post-restore steps (deferred)" -RunOn "the DESTINATION server, after the databases are recovered" -Databases $Databases -ModeText $(if($Mode -eq "PRE"){ "PRE-MIGRATION" } else { "MIGRATION-DAY" })))
    [void]$sb.AppendLine("-- The restore was $RecoveryMode, so these steps could not run. Recover the databases first, then run this file:")
    foreach($d in $Databases){ [void]$sb.AppendLine("-- RESTORE DATABASE [$(ConvertTo-SqlBracket $d)] WITH RECOVERY;") }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("USE master;")
    [void]$sb.AppendLine("GO")
    [void]$sb.AppendLine("")
    foreach($d in $Databases){ [void]$sb.Append((Get-PostRestoreSql -Target $d -Mode $Mode -CompatTarget $CompatTarget)) }
    if($Mode -eq "POST"){ [void]$sb.Append((Get-MigrationDayFinalSql -Targets $Databases)) }
    try {
        $sb.ToString() | Out-File -FilePath $path -Encoding UTF8 -Force
        Write-Log "  Post-restore steps saved for after recovery: $path" $Clr.Green
    } catch {
        Write-Log "  Could not save the post-restore script: $($_.Exception.Message)" $Clr.Red
    }
}

function New-ScriptHeader {
    param([string]$Title,[string]$RunOn,[string[]]$Databases,[string]$ModeText)
    $srcSrv = $TxtSrcSrv.Text.Trim()
    $dstSrv = $TxtDstSrv.Text.Trim()
    $notes  = $TxtNotes.Text.Trim()
    $h = New-Object System.Text.StringBuilder
    [void]$h.AppendLine("/*")
    [void]$h.AppendLine("  pnfp Migration - $Title")
    [void]$h.AppendLine("  Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$h.AppendLine("  Mode       : $ModeText")
    [void]$h.AppendLine("  Run on     : $RunOn")
    [void]$h.AppendLine("  Source     : $srcSrv")
    [void]$h.AppendLine("  Destination: $dstSrv")
    [void]$h.AppendLine("  Databases  : $(if($Databases.Count -gt 0){ $Databases -join ', ' } else { '(none selected)' })")
    if($notes){
        [void]$h.AppendLine("")
        [void]$h.AppendLine("  MIGRATION NOTES:")
        foreach($line in ($notes -split "`r?`n")){ [void]$h.AppendLine("  $line") }
    }
    [void]$h.AppendLine("*/")
    [void]$h.AppendLine("")
    return $h.ToString()
}

function Invoke-ScriptOut {
    if($script:Running){ return }
    if(-not $script:Mode){
        [System.Windows.Forms.MessageBox]::Show("Select a mode first.", "No Mode",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    $dbs = @(Get-SelectedDatabases)
    if($dbs.Count -eq 0){
        [System.Windows.Forms.MessageBox]::Show("No databases selected.", "Select Databases",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $isPre     = ($script:Mode -eq "PRE")
    $modeLbl   = if($isPre){ "PreMigration" } else { "PostMigration" }
    $modeText  = if($isPre){ "PRE-MIGRATION" } else { "MIGRATION-DAY" }
    $subFolder = if($isPre){ $script:BackupSubFolder } else { $script:PostBackupSubFolder }
    $srcOpen   = [bool]($script:SrcConn -and $script:SrcConn.State -eq 'Open')
    $dstOpen   = [bool]($script:DstConn -and $script:DstConn.State -eq 'Open')

    # ---- Same prompts as Run (the AG question comes first) ----
    $agPlan = Get-AgPlan -ForScriptOut $true
    if($null -eq $agPlan){ Write-Log "Script Out cancelled." $Clr.Dim; return }

    $renameMap = Show-RenameDialog -Databases $dbs
    if($null -eq $renameMap){ Write-Log "Script Out cancelled." $Clr.Dim; return }

    if($isPre){
        # Pre-migration restores fresh: no REPLACE, WITH RECOVERY
        $opts = [PSCustomObject]@{ Replace = $false; KeepReplication = $false; RestrictedUser = $false; RecoveryMode = "RECOVERY"; CloseConnections = $false }
    } else {
        $opts = Show-RestoreOptionsDialog
        if(-not $opts){ Write-Log "Script Out cancelled - no restore options chosen." $Clr.Dim; return }
    }

    $dstVersion = 15
    if($dstOpen){ try { $dstVersion = Get-SqlVersion $script:DstConn } catch { } }
    $compatTarget = Show-CompatLevelDialog -DstMajorVersion $dstVersion

    $script:StripePlan = @{}
    if($srcOpen){
        if(-not (Get-StripePlan $dbs)){ Write-Log "Script Out cancelled." $Clr.Dim; return }
    } else {
        Write-Log "  Source not connected - backups are scripted as single files." $Clr.Dim
    }


    Write-Sep "SCRIPT OUT ($modeText)"
    Set-Status "Generating scripts..." $Clr.Yellow

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outDir = $script:NotesFolder
    if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $script:ScriptDir }
    if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $env:TEMP }

    $dstFolder = Get-DstFolderName

    # ---- Logical file names from the source (the backup has the same names) ----
    $logical = @{}
    if($srcOpen){
        foreach($db in $dbs){
            try {
                $dt = Query-Table $script:SrcConn ("SELECT name, type FROM sys.master_files WHERE database_id = DB_ID(N'{0}') ORDER BY file_id" -f (ConvertTo-SqlLiteral $db))
                $list = @()
                foreach($row in $dt.Rows){
                    $t = "D"
                    if([int]$row["type"] -eq 1){ $t = "L" } elseif([int]$row["type"] -eq 2){ $t = "S" }
                    $list += [PSCustomObject]@{ LogicalName = [string]$row["name"]; Type = $t }
                }
                if($list.Count -gt 0){ $logical[$db] = $list }
            } catch {
                Write-Log "  Could not read file names for [$db]: $($_.Exception.Message)" $Clr.Yellow
            }
        }
    } else {
        Write-Log "  Source not connected - MOVE logical file names are left as placeholders." $Clr.Yellow
    }

    # ---- Destination default data / log directories ----
    $dataDir = "<data_dir>"; $logDir = "<log_dir>"
    if($dstOpen){
        try {
            $v = Query-Scalar $script:DstConn "SELECT CONVERT(NVARCHAR(512), SERVERPROPERTY('InstanceDefaultDataPath'))"
            if($v -and $v -isnot [System.DBNull]){ $dataDir = ([string]$v).TrimEnd('\') }
        } catch { }
        try {
            $v = Query-Scalar $script:DstConn "SELECT CONVERT(NVARCHAR(512), SERVERPROPERTY('InstanceDefaultLogPath'))"
            if($v -and $v -isnot [System.DBNull]){ $logDir = ([string]$v).TrimEnd('\') } else { $logDir = $dataDir }
        } catch { $logDir = $dataDir }
    } else {
        Write-Log "  Destination not connected - data/log directories are left as placeholders." $Clr.Yellow
    }

    # =============== SOURCE SCRIPT ===============
    $src = New-Object System.Text.StringBuilder
    [void]$src.Append((New-ScriptHeader -Title "$modeText source script" -RunOn "the SOURCE server" -Databases $dbs -ModeText $modeText))
    $backupSets = @{}
    foreach($db in $dbs){
        $dbq = ConvertTo-SqlBracket $db
        $dbl = ConvertTo-SqlLiteral $db
        $folder = ($db -replace '[\\/:*?"<>|]', '_').Trim()
        $dir    = "$($script:BackupShareRoot)\$dstFolder\$subFolder\$folder"
        $base   = "$dir\${db}_COPYONLY_$stamp.bak"
        $bkFiles = @(Get-BackupStripeFiles $null $db $base)
        $backupSets[$db] = $bkFiles
        $toClause = ($bkFiles | ForEach-Object { "DISK = N'" + (ConvertTo-SqlLiteral $_) + "'" }) -join ",`n   "
        $txt = @"
-- ============================================================
-- [$db]  STEP 1 - Copy-only backup
-- Backup folder (must exist, writable by the SQL Server service account): $dir
-- ============================================================
BACKUP DATABASE [$dbq]
TO $toClause
WITH COPY_ONLY, COMPRESSION, MAXTRANSFERSIZE = 4194304, INIT,
     NAME = N'$dbl - pnfp Copy-Only Backup';
GO


"@
        [void]$src.Append($txt)
    }

    if(-not $isPre){
        $txt = @"
-- ============================================================
-- OPTIONAL: SET SOURCE DATABASES READ_ONLY
-- Run after confirming the destination is good (prevents accidental writes during validation).
-- ============================================================


"@
        [void]$src.Append($txt)
        foreach($db in $dbs){
            $dbq = ConvertTo-SqlBracket $db
            $dbl = ConvertTo-SqlLiteral $db
            [void]$src.AppendLine("ALTER DATABASE [$dbq] SET READ_ONLY WITH ROLLBACK IMMEDIATE;")
            [void]$src.AppendLine("GO")
            [void]$src.AppendLine("SELECT name, is_read_only FROM sys.databases WHERE name = N'$dbl';")
            [void]$src.AppendLine("GO")
            [void]$src.AppendLine("")
        }
        [void]$src.AppendLine("-- To undo READ_ONLY (rollback plan):")
        foreach($db in $dbs){
            [void]$src.AppendLine("-- ALTER DATABASE [$(ConvertTo-SqlBracket $db)] SET READ_WRITE WITH ROLLBACK IMMEDIATE;")
        }
        [void]$src.AppendLine("")
    }

    # =============== DESTINATION SCRIPT ===============
    $dst = New-Object System.Text.StringBuilder
    [void]$dst.Append((New-ScriptHeader -Title "$modeText destination script" -RunOn "the DESTINATION server" -Databases $dbs -ModeText $modeText))
    [void]$dst.AppendLine("USE master;")
    [void]$dst.AppendLine("GO")
    [void]$dst.AppendLine("")

    $restoredTargets = @()
    foreach($db in $dbs){
        $target = $db
        if($renameMap.ContainsKey($db)){ $target = $renameMap[$db] }
        $tq = ConvertTo-SqlBracket $target
        $tl = ConvertTo-SqlLiteral $target
        $restoredTargets += $target

        $fromClause = (@($backupSets[$db]) | ForEach-Object { "DISK = N'" + (ConvertTo-SqlLiteral $_) + "'" }) -join ", "

        # MOVE clauses: real logical names when the source is connected, placeholders otherwise
        $filelist = ""
        if($logical.ContainsKey($db)){
            $mv = @(Get-RestoreMoveClauses -Files $logical[$db] -RestoreName $target -DataDir $dataDir -LogDir $logDir)
        } else {
            $mv = @("MOVE N'<logical_data_name>' TO N'$(ConvertTo-SqlLiteral "$dataDir\${target}_data0.mdf")'", "MOVE N'<logical_log_name>' TO N'$(ConvertTo-SqlLiteral "$logDir\${target}_log1.ldf")'")
            $filelist = "-- Logical file names unknown - list them, then edit the MOVE clauses below (one per file):`nRESTORE FILELISTONLY FROM $fromClause;`nGO`n"
        }

        $with = @() + $mv
        if($opts.Replace){ $with += "REPLACE" }
        if($opts.KeepReplication){ $with += "KEEP_REPLICATION" }
        if($opts.RestrictedUser){ $with += "RESTRICTED_USER" }
        if($opts.RecoveryMode -eq "STANDBY"){
            $with += "STANDBY = N'$(ConvertTo-SqlLiteral "$dataDir\${target}_undo.tuf")'"
        } else {
            $with += $opts.RecoveryMode
        }
        $with += "STATS = 10"
        $withStr = $with -join ",`n     "

        [void]$dst.AppendLine("-- ============================================================")
        if($target -ne $db){
            [void]$dst.AppendLine("-- [$db] restored AS [$target]  STEP - Restore")
        } else {
            [void]$dst.AppendLine("-- [$db]  STEP - Restore")
        }
        if($isPre){
            [void]$dst.AppendLine("-- Pre-migration is a fresh restore (no REPLACE): [$target] should not exist on the destination yet.")
        }
        [void]$dst.AppendLine("-- ============================================================")
        if($filelist){ [void]$dst.Append($filelist) }
        if($opts.CloseConnections){
            [void]$dst.AppendLine("IF DB_ID(N'$tl') IS NOT NULL ALTER DATABASE [$tq] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;")
            [void]$dst.AppendLine("GO")
        }
        $txt = @"
RESTORE DATABASE [$tq]
FROM $fromClause
WITH $withStr;
GO


"@
        [void]$dst.Append($txt)
    }

    # ---- Post-restore steps: after ALL databases above have been restored ----
    [void]$dst.AppendLine("-- ############################################################")
    [void]$dst.AppendLine("-- POST-RESTORE STEPS - run after ALL databases above have been restored")
    [void]$dst.AppendLine("-- ############################################################")
    if($opts.RecoveryMode -ne "RECOVERY"){
        [void]$dst.AppendLine("-- The restores above are $($opts.RecoveryMode). Run this section only after every database has been recovered:")
        foreach($t in $restoredTargets){ [void]$dst.AppendLine("-- RESTORE DATABASE [$(ConvertTo-SqlBracket $t)] WITH RECOVERY;") }
    }
    [void]$dst.AppendLine("")
    $postMode = if($isPre){ "PRE" } else { "POST" }
    foreach($t in $restoredTargets){
        [void]$dst.Append((Get-PostRestoreSql -Target $t -Mode $postMode -CompatTarget $compatTarget))
    }

    if(-not $isPre){
        [void]$dst.Append((Get-MigrationDayFinalSql -Targets $restoredTargets))
    }

    # =============== WRITE FILES ===============
    $written = @()
    $srcPath = Join-Path $outDir "pnfp_${modeLbl}_Source_Script_${stamp}.sql"
    $dstPath = Join-Path $outDir "pnfp_${modeLbl}_Destination_Script_${stamp}.sql"
    try {
        [void]$src.AppendLine("/* End of pnfp Migration source script */")
        $src.ToString() | Out-File -FilePath $srcPath -Encoding UTF8 -Force
        Write-Log "  Source script saved: $srcPath" $Clr.Green
        $written += $srcPath
    } catch {
        Write-Log "  Source script save failed: $($_.Exception.Message)" $Clr.Red
    }
    try {
        [void]$dst.AppendLine("/* End of pnfp Migration destination script */")
        $dst.ToString() | Out-File -FilePath $dstPath -Encoding UTF8 -Force
        Write-Log "  Destination script saved: $dstPath" $Clr.Green
        $written += $dstPath
    } catch {
        Write-Log "  Destination script save failed: $($_.Exception.Message)" $Clr.Red
    }

    if($agPlan.Enabled){
        try {
            $lf = @{}
            foreach($db in $dbs){
                $t = $db
                if($renameMap.ContainsKey($db)){ $t = $renameMap[$db] }
                if($logical.ContainsKey($db)){ $lf[$t] = $logical[$db] }
            }
            $ags = New-AgScripts -Databases $restoredTargets -AgName $agPlan.AgName -Seeding $agPlan.Seeding `
                -Replicas @($agPlan.Replicas) -SubFolder $subFolder -Stamp $stamp -LogicalFiles $lf
            $written += @(Save-AgScripts $ags $stamp)
        } catch {
            Write-Log "  AG script generation failed: $($_.Exception.Message)" $Clr.Red
        }
    }

    if($srcOpen){
        Write-Log "Scripting source logins and permissions..." $Clr.Blue
        try {
            Export-LoginScript $script:SrcConn
        } catch {
            Write-Log "  Login script failed: $($_.Exception.Message)" $Clr.Red
        }
    } else {
        Write-Log "  Source not connected - login script not generated." $Clr.Dim
    }

    if($written.Count -eq 0){ Set-Status "Script Out failed - see log." $Clr.Red; return }
    Set-Status "Script Out complete: $($written.Count) file(s) written." $Clr.Green
    Write-Log "Script Out complete. Files in: $outDir" $Clr.Green

    $open = [System.Windows.Forms.MessageBox]::Show(
        "Scripts saved to:`n$outDir`n`n$($written.Count) file(s) written.`n`nOpen folder?",
        "Script Out Complete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Information)
    if($open -eq [System.Windows.Forms.DialogResult]::Yes){
        Start-Process explorer.exe -ArgumentList "`"$outDir`""
    }
}

# ---------------------------------------------------------------
# MODAL: SET SOURCE DATABASES READ-ONLY
# ---------------------------------------------------------------
function Show-ReadOnlyDialog {
    param([string[]]$Databases)

    if($Databases.Count -eq 0){ return $null }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Set Source Databases READ_ONLY"
    $dlg.Size            = [System.Drawing.Size]::new(500, 448)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    # Header
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock      = [System.Windows.Forms.DockStyle]::Top
    $hdr.Height    = 46
    $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Text      = "  Set Source Databases to READ_ONLY"
    $ht.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $ht.ForeColor = $Clr.Orange
    $ht.Font      = $FontUIB
    $ht.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($ht)

    # Info
    $info = New-Object System.Windows.Forms.Label
    $info.Text      = "After restoring to the destination, you can set the source databases`nto READ_ONLY to prevent accidental writes during validation.`nSelect which databases to set READ_ONLY on the source:"
    $info.Location  = [System.Drawing.Point]::new(16, 58)
    $info.Size      = [System.Drawing.Size]::new(462, 54)
    $info.ForeColor = $Clr.Dim
    $info.Font      = $FontSm
    $dlg.Controls.Add($info)

    # Action selector: set READ_ONLY, or roll back to READ_WRITE
    $rbReadOnly = New-Object System.Windows.Forms.RadioButton
    $rbReadOnly.Text      = "Set READ_ONLY"
    $rbReadOnly.Location  = [System.Drawing.Point]::new(16, 114)
    $rbReadOnly.Size      = [System.Drawing.Size]::new(200, 24)
    $rbReadOnly.Checked   = $true
    $rbReadOnly.ForeColor = $Clr.Text; $rbReadOnly.Font = $FontUI
    $dlg.Controls.Add($rbReadOnly)

    $rbReadWrite = New-Object System.Windows.Forms.RadioButton
    $rbReadWrite.Text      = "Rollback: set READ_WRITE"
    $rbReadWrite.Location  = [System.Drawing.Point]::new(230, 114)
    $rbReadWrite.Size      = [System.Drawing.Size]::new(240, 24)
    $rbReadWrite.Checked   = $false
    $rbReadWrite.ForeColor = $Clr.Text; $rbReadWrite.Font = $FontUI
    $dlg.Controls.Add($rbReadWrite)

    $UpdateMode = {
        if($rbReadWrite.Checked){
            $dlg.Text  = "Rollback Source Databases to READ_WRITE"
            $ht.Text   = "  Rollback Source Databases to READ_WRITE"
            $info.Text = "Use this to undo a previous READ_ONLY on the source databases`nand allow writes again.`nSelect which databases to set READ_WRITE on the source:"
            $warn.Text = "Note: READ_WRITE WITH ROLLBACK IMMEDIATE disconnects active sessions."
        } else {
            $dlg.Text  = "Set Source Databases READ_ONLY"
            $ht.Text   = "  Set Source Databases to READ_ONLY"
            $info.Text = "After restoring to the destination, you can set the source databases`nto READ_ONLY to prevent accidental writes during validation.`nSelect which databases to set READ_ONLY on the source:"
            $warn.Text = "Note: READ_ONLY can be reversed with ALTER DATABASE [name] SET READ_WRITE."
        }
    }

    # Checklist
    $ChkList = New-Object System.Windows.Forms.CheckedListBox
    $ChkList.Location     = [System.Drawing.Point]::new(16, 146)
    $ChkList.Size         = [System.Drawing.Size]::new(462, 170)
    $ChkList.BackColor    = $Clr.Input
    $ChkList.ForeColor    = $Clr.Text
    $ChkList.Font         = $FontMono
    $ChkList.BorderStyle  = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ChkList.CheckOnClick = $true
    foreach($db in $Databases){ $ChkList.Items.Add($db, $true) | Out-Null }
    $dlg.Controls.Add($ChkList)

    # All / None buttons
    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text="All"; $btnAll.Location=[System.Drawing.Point]::new(16,324)
    $btnAll.Size=[System.Drawing.Size]::new(60,24); $btnAll.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnAll.BackColor=$Clr.Card; $btnAll.ForeColor=$Clr.Text; $btnAll.Font=$FontSm
    $btnAll.FlatAppearance.BorderColor=$Clr.Border
    $btnAll.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$true) } })
    $dlg.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text="None"; $btnNone.Location=[System.Drawing.Point]::new(82,324)
    $btnNone.Size=[System.Drawing.Size]::new(60,24); $btnNone.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnNone.BackColor=$Clr.Card; $btnNone.ForeColor=$Clr.Text; $btnNone.Font=$FontSm
    $btnNone.FlatAppearance.BorderColor=$Clr.Border
    $btnNone.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$false) } })
    $dlg.Controls.Add($btnNone)

    # Warning note
    $warn = New-Object System.Windows.Forms.Label
    $warn.Text      = "Note: READ_ONLY can be reversed with ALTER DATABASE [name] SET READ_WRITE."
    $warn.Location  = [System.Drawing.Point]::new(16, 354)
    $warn.Size      = [System.Drawing.Size]::new(462, 18)
    $warn.ForeColor = $Clr.Yellow
    $warn.Font      = $FontSm
    $dlg.Controls.Add($warn)

    # Buttons
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text      = "Apply + Script Out"
    $btnApply.Location  = [System.Drawing.Point]::new(196, 378)
    $btnApply.Size      = [System.Drawing.Size]::new(130, 28)
    $btnApply.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnApply.BackColor = $Clr.Orange
    $btnApply.ForeColor = [System.Drawing.Color]::FromArgb(10,10,10)
    $btnApply.Font      = $FontUIB
    $btnApply.FlatAppearance.BorderColor = $Clr.Orange
    $btnApply.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnApply)

    $btnScriptOnly = New-Object System.Windows.Forms.Button
    $btnScriptOnly.Text      = "Script Only"
    $btnScriptOnly.Location  = [System.Drawing.Point]::new(334, 378)
    $btnScriptOnly.Size      = [System.Drawing.Size]::new(90, 28)
    $btnScriptOnly.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnScriptOnly.BackColor = $Clr.Card
    $btnScriptOnly.ForeColor = $Clr.Purple
    $btnScriptOnly.Font      = $FontSm
    $btnScriptOnly.FlatAppearance.BorderColor = $Clr.Purple
    $btnScriptOnly.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $dlg.Controls.Add($btnScriptOnly)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Text      = "Skip"
    $btnSkip.Location  = [System.Drawing.Point]::new(432, 378)
    $btnSkip.Size      = [System.Drawing.Size]::new(52, 28)
    $btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkip.BackColor = $Clr.Card
    $btnSkip.ForeColor = $Clr.Dim
    $btnSkip.Font      = $FontSm
    $btnSkip.FlatAppearance.BorderColor = $Clr.Border
    $btnSkip.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnSkip)

    $rbReadWrite.Add_CheckedChanged($UpdateMode)
    $dlg.Add_Shown({ $dlg.Activate() })
    $r = $dlg.ShowDialog()

    $chosenMode = if($rbReadWrite.Checked){ "READ_WRITE" } else { "READ_ONLY" }
    $chosen = @()
    foreach($i in $ChkList.CheckedIndices){ $chosen += $ChkList.Items[$i].ToString() }
    $dlg.Dispose()

    return [PSCustomObject]@{
        Action    = $r          # OK=apply+script, Retry=script only, Cancel=skip
        Databases = $chosen
        Mode      = $chosenMode  # "READ_ONLY" | "READ_WRITE"
    }
}

# Helper: set databases READ_ONLY (or roll back to READ_WRITE) on source and/or script it out
function Apply-ReadOnly {
    param(
        [string[]]$Databases,
        [bool]$Execute,
        [bool]$ScriptOut,
        [ValidateSet("READ_ONLY","READ_WRITE")][string]$Mode = "READ_ONLY"
    )

    $other   = if($Mode -eq "READ_ONLY"){ "READ_WRITE" } else { "READ_ONLY" }
    $label   = if($Mode -eq "READ_ONLY"){ "Set Source Databases READ_ONLY" } else { "Rollback Source Databases to READ_WRITE" }
    $fileTag = if($Mode -eq "READ_ONLY"){ "SetReadOnly" } else { "SetReadWrite" }

    if($Databases.Count -eq 0){
        Write-Log "  No databases selected for $Mode." $Clr.Dim; return
    }

    if($ScriptOut){
        $stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
        $outDir  = $script:NotesFolder
        if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $script:ScriptDir }
        if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $env:TEMP }
        $outPath = Join-Path $outDir "pnfp_${fileTag}_$stamp.sql"

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("/*")
        [void]$sb.AppendLine("  pnfp Migration - $label")
        [void]$sb.AppendLine("  Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$sb.AppendLine("  Source     : $($TxtSrcSrv.Text.Trim())")
        [void]$sb.AppendLine("  Databases  : $($Databases -join ', ')")
        [void]$sb.AppendLine("  Run this on the SOURCE server.")
        [void]$sb.AppendLine("  To reverse: ALTER DATABASE [name] SET $other WITH ROLLBACK IMMEDIATE;")
        [void]$sb.AppendLine("*/")
        [void]$sb.AppendLine("")
        foreach($db in $Databases){
            [void]$sb.AppendLine("-- Set [$db] to $Mode (kicks existing connections with rollback)")
            [void]$sb.AppendLine("ALTER DATABASE [$db] SET $Mode WITH ROLLBACK IMMEDIATE;")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("-- Verify:")
            [void]$sb.AppendLine("SELECT name, is_read_only FROM sys.databases WHERE name = N'$db';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
        }
        [void]$sb.AppendLine("/*")
        [void]$sb.AppendLine("  To undo ($other) on all databases above:")
        foreach($db in $Databases){
            [void]$sb.AppendLine("  ALTER DATABASE [$db] SET $other WITH ROLLBACK IMMEDIATE;")
        }
        [void]$sb.AppendLine("*/")

        try {
            $sb.ToString() | Out-File -FilePath $outPath -Encoding UTF8 -Force
            Write-Log "  $Mode script saved: $outPath" $Clr.Green
        } catch {
            Write-Log "  Script save failed: $($_.Exception.Message)" $Clr.Red
        }
    }

    if($Execute){
        if(-not $script:SrcConn -or $script:SrcConn.State -ne 'Open'){
            Write-Log "  Source not connected - cannot apply $Mode." $Clr.Red; return
        }
        foreach($db in $Databases){
            try {
                Write-Log "  Setting [$db] $Mode on source..." $Clr.Dim
                Exec-Sql $script:SrcConn "ALTER DATABASE [$db] SET $Mode WITH ROLLBACK IMMEDIATE"
                Write-Log "  [$db] is now $Mode" $Clr.Green
            } catch {
                Write-Log "  FAILED to set [$db] $Mode`: $($_.Exception.Message)" $Clr.Red
            }
        }
    }
}

# ---------------------------------------------------------------
# DESTINATION: make sure all user databases use PAGE_VERIFY CHECKSUM
# ---------------------------------------------------------------
function Set-UserDbPageVerifyChecksum {
    param([System.Data.SqlClient.SqlConnection]$Conn)

    $dt = Query-Table $Conn @"
SELECT name, page_verify_option_desc, state_desc, CONVERT(INT, is_read_only) AS is_read_only
FROM sys.databases
WHERE database_id > 4
  AND source_database_id IS NULL
  AND is_distributor = 0
ORDER BY name;
"@
    $total = 0; $already = 0; $changed = 0; $skipped = 0; $failed = 0
    foreach($row in $dt.Rows){
        $total++
        $name  = [string]$row["name"]
        $pv    = [string]$row["page_verify_option_desc"]
        $state = [string]$row["state_desc"]
        $ro    = ([int]$row["is_read_only"] -eq 1)

        if($pv -eq "CHECKSUM"){ $already++; continue }
        if($state -ne "ONLINE"){
            Write-Log "  [$name] is $state (PAGE_VERIFY = $pv) - skipped, cannot be changed while not ONLINE." $Clr.Yellow
            $skipped++; continue
        }
        if($ro){
            Write-Log "  [$name] is READ_ONLY (PAGE_VERIFY = $pv) - skipped. Set READ_WRITE, change it, then set READ_ONLY again." $Clr.Yellow
            $skipped++; continue
        }
        try {
            $q = $name.Replace(']',']]')
            Exec-Sql $Conn "ALTER DATABASE [$q] SET PAGE_VERIFY CHECKSUM"
            Write-Log "  [$name] PAGE_VERIFY changed from $pv to CHECKSUM." $Clr.Green
            $changed++
        } catch {
            Write-Log "  [$name] could not set PAGE_VERIFY CHECKSUM: $($_.Exception.Message)" $Clr.Red
            $failed++
        }
    }
    Write-Log "  User databases checked: $total  |  already CHECKSUM: $already  |  changed: $changed  |  skipped: $skipped  |  failed: $failed" $Clr.Dim
    if($changed -gt 0){
        Write-Log "  Note: existing pages get a checksum the next time they are written, not immediately." $Clr.Dim
    }
}

# ---------------------------------------------------------------
# DESTINATION: make sa the owner of the databases restored in this run.
# Other user databases are never changed; a warning is logged if their owner is not sa.
# ---------------------------------------------------------------
function Set-UserDbOwnerSa {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [string[]]$RestoredDatabases = @()
    )

    # sa is always principal_id 1, even if it has been renamed
    $saName = [string](Query-Scalar $Conn "SELECT name FROM sys.server_principals WHERE principal_id = 1")
    if(-not $saName){ throw "Could not find the sa login (principal_id 1) on the destination." }
    $saQ = $saName.Replace(']',']]')

    $dt = Query-Table $Conn @"
SELECT d.name,
       d.state_desc,
       CONVERT(INT, d.is_read_only)       AS is_read_only,
       CONVERT(INT, d.is_trustworthy_on)  AS is_trustworthy_on,
       SUSER_SNAME(d.owner_sid)           AS owner_name,
       CASE WHEN d.owner_sid = 0x01 THEN 1 ELSE 0 END AS is_sa
FROM sys.databases d
WHERE d.database_id > 4
  AND d.source_database_id IS NULL
  AND d.is_distributor = 0
ORDER BY d.name;
"@
    $total = 0; $already = 0; $changed = 0; $skipped = 0; $failed = 0; $othersWarned = 0
    foreach($row in $dt.Rows){
        $name  = [string]$row["name"]
        $state = [string]$row["state_desc"]
        $ro    = ([int]$row["is_read_only"] -eq 1)
        $trust = ([int]$row["is_trustworthy_on"] -eq 1)
        $owner = if($row["owner_name"] -is [System.DBNull]){ "(no matching login)" } else { [string]$row["owner_name"] }

        # Databases that were not restored in this run: warn only, never change
        if(-not ($RestoredDatabases -contains $name)){
            if([int]$row["is_sa"] -ne 1){
                Write-Log "  WARNING: [$name] was not part of this restore and its owner is $owner, not $saName. Not changed." $Clr.Orange
                $othersWarned++
            }
            continue
        }

        $total++
        if([int]$row["is_sa"] -eq 1){ $already++; continue }
        if($state -ne "ONLINE"){
            Write-Log "  [$name] is $state (owner: $owner) - skipped, cannot be changed while not ONLINE." $Clr.Yellow
            $skipped++; continue
        }
        if($ro){
            Write-Log "  [$name] is READ_ONLY (owner: $owner) - skipped. Set READ_WRITE, change the owner, then set READ_ONLY again." $Clr.Yellow
            $skipped++; continue
        }
        try {
            $q = $name.Replace(']',']]')
            Exec-Sql $Conn "ALTER AUTHORIZATION ON DATABASE::[$q] TO [$saQ]"
            Write-Log "  [$name] owner changed from $owner to $saName." $Clr.Green
            $changed++
            if($trust){
                Write-Log "    Note: [$name] has TRUSTWORTHY ON and is now owned by sa. Review whether it needs TRUSTWORTHY." $Clr.Orange
            }
        } catch {
            Write-Log "  [$name] could not change owner to ${saName}: $($_.Exception.Message)" $Clr.Red
            $failed++
        }
    }
    Write-Log "  Restored databases checked: $total  |  already owned by $saName`: $already  |  changed: $changed  |  skipped: $skipped  |  failed: $failed" $Clr.Dim
    if($othersWarned -gt 0){
        Write-Log "  Other user databases on the destination not owned by ${saName}: $othersWarned (warned above, not changed)." $Clr.Orange
    }
}

# ---------------------------------------------------------------
# SOURCE LOGINS: script out logins (with password hashes and SIDs), server roles and
# server-level permissions to a .sql file.
# The logins section comes from the instance login script (PART 2 of the SP_REV logins scripts),
# which needs sp_hexadecimal in master or AIDO on the source (installed by PART 1). If it is missing, a
# warning is logged and the run continues without a login file. The roles and permissions sections
# are generated here.
# ---------------------------------------------------------------
# Optional: set to a folder to keep the login script out of the default output folder.
$script:LoginScriptFolder = ""

# Instance login script (logins part of "Script Instance SP_REV Logins script PART2").
# Needs sp_hexadecimal on the source, in master or AIDO (installed by PART 1). It only reads the source and
# PRINTs a guarded CREATE LOGIN script (SQL logins with password hash and SID, Windows logins,
# deny / disabled state). Server roles and permissions are generated separately below.
$script:InstanceLoginSql = @'
SET NOCOUNT ON;
USE [master];

DECLARE @Logins TABLE(LoginName sysname);
DECLARE @LoginName sysname;
SET @LoginName = '';

DECLARE @name sysname;
DECLARE @type varchar(1);
DECLARE @hasaccess int;
DECLARE @denylogin int;
DECLARE @is_disabled int;
DECLARE @PWD_varbinary varbinary(256);
DECLARE @PWD_string varchar(514);
DECLARE @SID_varbinary varbinary(85);
DECLARE @SID_string varchar(514);
DECLARE @tmpstr nvarchar(4000);
DECLARE @is_policy_checked varchar(3);
DECLARE @is_expiration_checked varchar(3);
DECLARE @defaultdb sysname;
DECLARE @nameq nvarchar(600);
DECLARE @qnameq nvarchar(600);

INSERT INTO @Logins
SELECT p.name
FROM sys.server_principals p
  LEFT JOIN sys.syslogins l ON (l.name = p.name)
WHERE (@LoginName = '' OR l.loginname = @LoginName)
  AND p.type IN ('S', 'G', 'U')
  AND l.loginname NOT IN
(
'sa',
'##MS_SQLResourceSigningCertificate##',
'##MS_SQLReplicationSigningCertificate##',
'##MS_SQLAuthenticatorCertificate##',
'##MS_PolicySigningCertificate##',
'##MS_SmoExtendedSigningCertificate##',
'##MS_PolicyTsqlExecutionLogin##',
'NT AUTHORITY\SYSTEM',
'NT SERVICE\MSSQLSERVER',
'NT AUTHORITY\NETWORK SERVICE',
'NT AUTHORITY\ANONYMOUS LOGON',
'##MS_PolicyEventProcessingLogin##',
'##MS_AgentSigningCertificate##'
)
  AND l.loginname NOT LIKE ('NT SERVICE\%') AND l.loginname NOT LIKE ('%$');

SET @tmpstr = 'USE [master]' + CHAR(13);
PRINT @tmpstr;

WHILE ((SELECT COUNT(*) FROM @Logins) > 0)
BEGIN
    SET @LoginName = (SELECT TOP 1 LoginName FROM @Logins ORDER BY 1);
    SET @name = NULL;
    SET @is_policy_checked = NULL;
    SET @is_expiration_checked = NULL;

    SELECT
        @SID_varbinary = p.sid
      , @name = p.name
      , @type = p.type
      , @is_disabled = p.is_disabled
      , @defaultdb = p.default_database_name
      , @hasaccess = l.hasaccess
      , @denylogin = l.denylogin
    FROM sys.server_principals p
         LEFT JOIN sys.syslogins l ON (l.name = p.name)
    WHERE p.type IN ('S', 'G', 'U') AND p.name = @LoginName;

    IF (@name IS NOT NULL)
    BEGIN
        SET @nameq  = REPLACE(@name, '''', '''''');
        SET @qnameq = REPLACE(QUOTENAME(@name), '''', '''''');

        PRINT '';
        SET @tmpstr = 'Print ''-- Login: ' + @nameq + '''';
        PRINT @tmpstr;

        IF (@type IN ('G', 'U'))
        BEGIN -- NT authenticated account/group
            SET @tmpstr =
                'IF not exists(select sp.name from master.sys.server_principals sp where sp.name=''' + @nameq + ''')' + CHAR(13) +
                '  begin' + CHAR(13) +
                '    CREATE LOGIN ' + QUOTENAME(@name) + ' FROM WINDOWS WITH DEFAULT_DATABASE = ' + QUOTENAME(@defaultdb);
        END
        ELSE
        BEGIN -- SQL Server authentication
            SET @PWD_varbinary = CAST(LOGINPROPERTY(@name, 'PasswordHash') AS varbinary(256));
            EXEC [__HEXDB__]..sp_hexadecimal @PWD_varbinary, @PWD_string OUT;
            EXEC [__HEXDB__]..sp_hexadecimal @SID_varbinary, @SID_string OUT;

            SELECT @is_policy_checked = CASE is_policy_checked WHEN 1 THEN 'ON' WHEN 0 THEN 'OFF' ELSE NULL END
            FROM sys.sql_logins WHERE name = @name;

            SELECT @is_expiration_checked = CASE is_expiration_checked WHEN 1 THEN 'ON' WHEN 0 THEN 'OFF' ELSE NULL END
            FROM sys.sql_logins WHERE name = @name;

            SET @tmpstr =
                'IF not exists(select sp.name from master.sys.server_principals sp where sp.name=''' + @nameq + ''')' + CHAR(13) +
                '  begin' + CHAR(13) +
                '    CREATE LOGIN ' + QUOTENAME(@name) + ' WITH PASSWORD = ' + @PWD_string +
                ' HASHED, SID = ' + @SID_string + ', DEFAULT_DATABASE = ' + QUOTENAME(@defaultdb);

            IF (@is_policy_checked IS NOT NULL)
                SET @tmpstr = @tmpstr + ', CHECK_POLICY = ' + @is_policy_checked;
            IF (@is_expiration_checked IS NOT NULL)
                SET @tmpstr = @tmpstr + ', CHECK_EXPIRATION = ' + @is_expiration_checked;
        END

        IF (@denylogin = 1)
        BEGIN -- login is denied access
            SET @tmpstr = @tmpstr + '; DENY CONNECT SQL TO ' + QUOTENAME(@name);
        END
        ELSE IF (@hasaccess = 0)
        BEGIN -- login exists but does not have access
            SET @tmpstr = @tmpstr + '; REVOKE CONNECT SQL FROM ' + QUOTENAME(@name);
        END
        IF (@is_disabled = 1)
        BEGIN -- login is disabled
            SET @tmpstr = @tmpstr + '; ALTER LOGIN ' + QUOTENAME(@name) + ' DISABLE';
        END

        SET @tmpstr = @tmpstr + CHAR(13) + '    PRINT ''CREATED USER: ' + @qnameq + '''';
        SET @tmpstr = @tmpstr + CHAR(13) + '  end' + CHAR(13);

        PRINT @tmpstr;
    END

    DELETE FROM @Logins WHERE LoginName = @LoginName;
END
'@

# Returns the database that holds sp_hexadecimal (master first, then AIDO), or $null if not found.
function Get-HexProcDb {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    foreach($db in @("master","AIDO")){
        try {
            $n = Query-Scalar $Conn "SELECT COUNT(*) FROM [$db].sys.procedures WHERE name = N'sp_hexadecimal'"
            if([int]$n -gt 0){ return $db }
        } catch { }
    }
    return $null
}

# Run the instance login script and return what it PRINTs (the CREATE LOGIN script) as text.
function Invoke-InstanceLoginScript {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$HexDb = "master")
    $lines = New-Object System.Collections.Generic.List[string]
    $h = { param($sender,$e) foreach($err in $e.Errors){ [void]$lines.Add([string]$err.Message) } }.GetNewClosure()
    $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler]$h
    $prevDb = $Conn.Database
    $Conn.add_InfoMessage($handler)
    try {
        $cmd = New-Object System.Data.SqlClient.SqlCommand ($script:InstanceLoginSql.Replace("__HEXDB__", $HexDb.Replace("]","]]"))), $Conn
        $cmd.CommandTimeout = 300
        [void]$cmd.ExecuteNonQuery()
    } finally {
        $Conn.remove_InfoMessage($handler)
        try { if($prevDb -and $Conn.Database -ne $prevDb){ $Conn.ChangeDatabase($prevDb) } } catch { }
    }
    $text = ($lines -join "`r`n")
    return ($text -replace "`r(?!`n)", "`r`n")
}

function Export-LoginScript {
    param([System.Data.SqlClient.SqlConnection]$Conn)

    $sqlName = { param($s) $s.Replace("'","''") }       # for N'...' literals
    $brk     = { param($s) $s.Replace(']',']]') }       # for [...] identifiers

    # sp_hexadecimal (installed by PART 1) must already exist in master or AIDO on the source.
    # Nothing is installed by this tool.
    $hexDb = Get-HexProcDb $Conn
    if(-not $hexDb){
        Write-Log "  WARNING: sp_hexadecimal was not found in master or AIDO on the source. Run the SP_REV logins scripts (PART 1, then PART 2) on the source server to copy the logins to the destination. Continuing without a login file." $Clr.Orange
        return
    }
    $revText = $null
    try {
        $revText = Invoke-InstanceLoginScript $Conn $hexDb
    } catch {
        Write-Log "  WARNING: the instance login script failed: $($_.Exception.Message). Run the SP_REV logins scripts manually on the source server to copy the logins. Continuing without a login file." $Clr.Orange
        return
    }
    if(-not $revText -or $revText -notmatch 'CREATE LOGIN'){
        Write-Log "  WARNING: the instance login script returned no logins. Run the SP_REV logins scripts manually on the source server if logins are expected. Continuing without a login file." $Clr.Orange
        return
    }
    Write-Log "  Logins scripted with the instance login script (sp_hexadecimal in [$hexDb])." $Clr.Dim

    $srcName = $TxtSrcSrv.Text.Trim()
    $stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
    $folder  = $script:LoginScriptFolder
    if(-not $folder -or -not (Test-Path $folder)){ $folder = Get-JobFileFolder }
    $outPath = Join-Path $folder ("pnfp_Logins_{0}_{1}.sql" -f (Get-SafeFileName $srcName), $stamp)

    $notSystem = @"
sp.principal_id > 1
  AND sp.name NOT LIKE '##%'
  AND sp.name NOT LIKE 'NT SERVICE\%'
  AND sp.name NOT LIKE 'NT AUTHORITY\%'
  AND sp.name NOT LIKE '%[$]'
"@

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("/*")
    [void]$sb.AppendLine("  pnfp Migration - source logins, server roles and server permissions")
    [void]$sb.AppendLine("  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("  Source    : $srcName")
    [void]$sb.AppendLine("  Run this on the DESTINATION server.")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("  WARNING: this file contains SQL login password hashes. Treat it like a credential file.")
    [void]$sb.AppendLine("  Restrict access to it and delete it when the migration is finished.")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("  Logins come from the instance login script and keep their original default database, so CREATE LOGIN")
    [void]$sb.AppendLine("  fails for a database that does not exist on the destination yet. Run the file after the restores,")
    [void]$sb.AppendLine("  or expect those errors. Roles and permissions below are generated by this tool and exclude sa,")
    [void]$sb.AppendLine("  NT SERVICE\* and NT AUTHORITY\* logins, endpoint/login/availability group permissions, credentials")
    [void]$sb.AppendLine("  and proxies. Database users and their permissions travel inside the database backups.")
    [void]$sb.AppendLine("*/")
    [void]$sb.AppendLine("")

    # ---- Logins (from the instance login script) ----
    [void]$sb.AppendLine("-- ====================== LOGINS (instance login script) ======================")
    [void]$sb.AppendLine($revText)
    [void]$sb.AppendLine("GO")
    [void]$sb.AppendLine("")

    # ---- User-defined server roles, memberships and server permissions (SQL Server 2012+) ----
    $nRoles = 0; $nMembers = 0; $nPerms = 0
    try {
        $roles = Query-Table $Conn @"
SELECT r.name, ISNULL(o.name, N'sa') AS owner_name
FROM sys.server_principals r
LEFT JOIN sys.server_principals o ON o.principal_id = r.owning_principal_id
WHERE r.type = 'R' AND r.is_fixed_role = 0 AND r.name <> N'public'
ORDER BY r.name;
"@
        [void]$sb.AppendLine("-- ===================== USER-DEFINED SERVER ROLES =====================")
        foreach($row in $roles.Rows){
            $n = [string]$row["name"]; $o = [string]$row["owner_name"]
            [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'$(& $sqlName $n)' AND type = 'R')")
            [void]$sb.AppendLine("    CREATE SERVER ROLE [$(& $brk $n)] AUTHORIZATION [$(& $brk $o)];")
            $nRoles++
        }
        [void]$sb.AppendLine("GO")
        [void]$sb.AppendLine("")

        $members = Query-Table $Conn @"
SELECT r.name AS role_name, sp.name AS member_name
FROM sys.server_role_members rm
JOIN sys.server_principals r  ON r.principal_id  = rm.role_principal_id
JOIN sys.server_principals sp ON sp.principal_id = rm.member_principal_id
WHERE sp.type IN ('S','U','G')
  AND $notSystem
ORDER BY r.name, sp.name;
"@
        [void]$sb.AppendLine("-- ======================= SERVER ROLE MEMBERSHIP =======================")
        foreach($row in $members.Rows){
            $rn = [string]$row["role_name"]; $mn = [string]$row["member_name"]
            [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.server_role_members rm JOIN sys.server_principals r ON r.principal_id = rm.role_principal_id JOIN sys.server_principals m ON m.principal_id = rm.member_principal_id WHERE r.name = N'$(& $sqlName $rn)' AND m.name = N'$(& $sqlName $mn)')")
            [void]$sb.AppendLine("    ALTER SERVER ROLE [$(& $brk $rn)] ADD MEMBER [$(& $brk $mn)];")
            $nMembers++
        }
        [void]$sb.AppendLine("GO")
        [void]$sb.AppendLine("")

        $perms = Query-Table $Conn @"
SELECT sp.name AS grantee, p.state_desc, p.permission_name
FROM sys.server_permissions p
JOIN sys.server_principals sp ON sp.principal_id = p.grantee_principal_id
WHERE p.class = 100
  AND ( (sp.type IN ('S','U','G') AND $notSystem)
        OR (sp.type = 'R' AND sp.is_fixed_role = 0 AND sp.name <> N'public') )
  AND NOT (p.permission_name = 'CONNECT SQL' AND p.state_desc = 'GRANT')
ORDER BY sp.name, p.permission_name;
"@
        [void]$sb.AppendLine("-- ===================== SERVER-LEVEL PERMISSIONS =====================")
        foreach($row in $perms.Rows){
            $g = & $brk ([string]$row["grantee"]); $perm = [string]$row["permission_name"]; $st = [string]$row["state_desc"]
            switch($st){
                "GRANT"                  { [void]$sb.AppendLine("GRANT $perm TO [$g];") }
                "GRANT_WITH_GRANT_OPTION"{ [void]$sb.AppendLine("GRANT $perm TO [$g] WITH GRANT OPTION;") }
                "DENY"                   { [void]$sb.AppendLine("DENY $perm TO [$g];") }
            }
            $nPerms++
        }
        [void]$sb.AppendLine("GO")
    } catch {
        [void]$sb.AppendLine("-- Server roles / permissions could not be scripted: $($_.Exception.Message)")
        Write-Log "  Server roles / permissions could not be scripted: $($_.Exception.Message)" $Clr.Yellow
    }

    $sb.ToString() | Out-File -FilePath $outPath -Encoding UTF8 -Force
    Write-Log "  Roles: $nRoles  |  role memberships: $nMembers  |  server permissions: $nPerms  (logins from the instance login script)" $Clr.Dim
    Write-Log "  Login script saved: $outPath" $Clr.Green
    Write-Log "  This file contains password hashes. Restrict access and delete it after the migration." $Clr.Orange
}

# ---------------------------------------------------------------
# SOURCE SQL AGENT JOBS: disable (pre-migration) / enable (rollback)
# ---------------------------------------------------------------
function Get-AgentJobs {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $dt = Query-Table $Conn @"
SELECT CONVERT(VARCHAR(36), j.job_id) AS job_id,
       j.name,
       CONVERT(INT, j.enabled) AS enabled,
       ISNULL(c.name, N'') AS category
FROM msdb.dbo.sysjobs j
LEFT JOIN msdb.dbo.syscategories c ON c.category_id = j.category_id
ORDER BY j.name;
"@
    $list = @()
    foreach($row in $dt.Rows){
        $list += [PSCustomObject]@{
            Id       = [string]$row["job_id"]
            Name     = [string]$row["name"]
            Enabled  = ([int]$row["enabled"] -eq 1)
            Category = [string]$row["category"]
        }
    }
    return ,$list
}

function Get-JobFileFolder {
    $d = $script:NotesFolder
    if(-not $d -or -not (Test-Path $d)){ $d = $script:ScriptDir }
    if(-not $d -or -not (Test-Path $d)){ $d = $env:TEMP }
    return $d
}

function Get-SafeFileName {
    param([string]$Text)
    return ($Text -replace '[^A-Za-z0-9_.-]', '_')
}

# Most recent list of jobs this tool disabled on the source (written by Apply-AgentJobs)
function Get-SavedDisabledJobList {
    $script:LastJobListPath = ""
    try {
        $srcTag = Get-SafeFileName ($TxtSrcSrv.Text.Trim())
        $f = Get-ChildItem -Path (Get-JobFileFolder) -Filter "pnfp_DisabledJobs_${srcTag}_*.csv" -ErrorAction Stop |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if($f){
            $script:LastJobListPath = $f.FullName
            return @(Import-Csv -Path $f.FullName)
        }
    } catch { }
    return @()
}

function Show-AgentJobsDialog {
    param(
        $AllJobs,
        [ValidateSet("DISABLE","ENABLE")][string]$DefaultMode = "DISABLE",
        $SavedJobs = @()
    )

    $enabledJobs  = @($AllJobs | Where-Object { $_.Enabled })
    $disabledJobs = @($AllJobs | Where-Object { -not $_.Enabled })
    $savedNames   = @($SavedJobs | ForEach-Object { $_.Name })
    $state = @{ Items = @() }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Source SQL Agent Jobs"
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock = [System.Windows.Forms.DockStyle]::Top; $hdr.Height = 46; $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Dock = [System.Windows.Forms.DockStyle]::Fill; $ht.ForeColor = $Clr.Orange; $ht.Font = $FontUIB
    $ht.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($ht)

    $info = New-Object System.Windows.Forms.Label
    $info.Location = [System.Drawing.Point]::new(16, 56); $info.Size = [System.Drawing.Size]::new(488, 44)
    $info.ForeColor = $Clr.Dim; $info.Font = $FontSm
    $dlg.Controls.Add($info)

    $rbDisable = New-Object System.Windows.Forms.RadioButton
    $rbDisable.Text = "Disable jobs"; $rbDisable.Location = [System.Drawing.Point]::new(16, 102); $rbDisable.Size = [System.Drawing.Size]::new(200, 24)
    $rbDisable.ForeColor = $Clr.Text; $rbDisable.Font = $FontUI
    $dlg.Controls.Add($rbDisable)

    $rbEnable = New-Object System.Windows.Forms.RadioButton
    $rbEnable.Text = "Rollback: enable jobs"; $rbEnable.Location = [System.Drawing.Point]::new(230, 102); $rbEnable.Size = [System.Drawing.Size]::new(250, 24)
    $rbEnable.ForeColor = $Clr.Text; $rbEnable.Font = $FontUI
    $dlg.Controls.Add($rbEnable)

    $ChkList = New-Object System.Windows.Forms.CheckedListBox
    $ChkList.Location = [System.Drawing.Point]::new(16, 132); $ChkList.Size = [System.Drawing.Size]::new(488, 206)
    $ChkList.BackColor = $Clr.Input; $ChkList.ForeColor = $Clr.Text; $ChkList.Font = $FontMono
    $ChkList.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ChkList.CheckOnClick = $true; $ChkList.HorizontalScrollbar = $true
    $dlg.Controls.Add($ChkList)

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text="All"; $btnAll.Location=[System.Drawing.Point]::new(16,346); $btnAll.Size=[System.Drawing.Size]::new(60,24)
    $btnAll.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnAll.BackColor=$Clr.Card; $btnAll.ForeColor=$Clr.Text; $btnAll.Font=$FontSm
    $btnAll.FlatAppearance.BorderColor=$Clr.Border
    $btnAll.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$true) } })
    $dlg.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text="None"; $btnNone.Location=[System.Drawing.Point]::new(82,346); $btnNone.Size=[System.Drawing.Size]::new(60,24)
    $btnNone.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnNone.BackColor=$Clr.Card; $btnNone.ForeColor=$Clr.Text; $btnNone.Font=$FontSm
    $btnNone.FlatAppearance.BorderColor=$Clr.Border
    $btnNone.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$false) } })
    $dlg.Controls.Add($btnNone)

    $count = New-Object System.Windows.Forms.Label
    $count.Location = [System.Drawing.Point]::new(152, 349); $count.Size = [System.Drawing.Size]::new(352, 20)
    $count.ForeColor = $Clr.Dim; $count.Font = $FontSm
    $dlg.Controls.Add($count)

    $warn = New-Object System.Windows.Forms.Label
    $warn.Location = [System.Drawing.Point]::new(16, 378); $warn.Size = [System.Drawing.Size]::new(488, 36)
    $warn.ForeColor = $Clr.Yellow; $warn.Font = $FontSm
    $dlg.Controls.Add($warn)

    $UpdateList = {
        $ChkList.Items.Clear()
        if($rbEnable.Checked){
            $dlg.Text  = "Rollback: Enable Source SQL Agent Jobs"
            $ht.Text   = "  Rollback: Enable Source SQL Agent Jobs"
            $state.Items = @($disabledJobs)
            if($savedNames.Count -gt 0){
                $info.Text = "Lists the jobs that are currently disabled on the source. Jobs disabled by the earlier`npre-migration run are pre-checked, so jobs that were already off before it stay off."
            } else {
                $info.Text = "Lists the jobs that are currently disabled on the source. No saved list from a`npre-migration run was found, so all of them are pre-checked."
            }
            $warn.Text = "Tick only the jobs that should run again. Use All to enable every disabled job."
        } else {
            $dlg.Text  = "Disable Source SQL Agent Jobs"
            $ht.Text   = "  Disable Source SQL Agent Jobs"
            $state.Items = @($enabledJobs)
            $info.Text = "Lists the jobs that are currently enabled on the source. Disabling them stops new runs`nso nothing changes data while the migration is in progress."
            $warn.Text = "Disabling does not stop a job that is already running. Uncheck any job that must keep running (replication, CDC, log shipping)."
        }
        foreach($j in $state.Items){
            $chk = $true
            if($rbEnable.Checked -and $savedNames.Count -gt 0){ $chk = ($savedNames -contains $j.Name) }
            $label = $j.Name
            if($j.Category){ $label = "$($j.Name)   [$($j.Category)]" }
            [void]$ChkList.Items.Add($label, $chk)
        }
        $count.Text = "$($state.Items.Count) job(s) listed"
    }

    # Buttons
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text="Apply + Script Out"; $btnApply.Location=[System.Drawing.Point]::new(212,424); $btnApply.Size=[System.Drawing.Size]::new(130,28)
    $btnApply.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnApply.BackColor=$Clr.Orange
    $btnApply.ForeColor=[System.Drawing.Color]::FromArgb(10,10,10); $btnApply.Font=$FontUIB
    $btnApply.FlatAppearance.BorderColor=$Clr.Orange
    $btnApply.DialogResult=[System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnApply)

    $btnScriptOnly = New-Object System.Windows.Forms.Button
    $btnScriptOnly.Text="Script Only"; $btnScriptOnly.Location=[System.Drawing.Point]::new(350,424); $btnScriptOnly.Size=[System.Drawing.Size]::new(90,28)
    $btnScriptOnly.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnScriptOnly.BackColor=$Clr.Card
    $btnScriptOnly.ForeColor=$Clr.Purple; $btnScriptOnly.Font=$FontSm
    $btnScriptOnly.FlatAppearance.BorderColor=$Clr.Purple
    $btnScriptOnly.DialogResult=[System.Windows.Forms.DialogResult]::Retry
    $dlg.Controls.Add($btnScriptOnly)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Text="Skip"; $btnSkip.Location=[System.Drawing.Point]::new(448,424); $btnSkip.Size=[System.Drawing.Size]::new(56,28)
    $btnSkip.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnSkip.BackColor=$Clr.Card
    $btnSkip.ForeColor=$Clr.Dim; $btnSkip.Font=$FontSm
    $btnSkip.FlatAppearance.BorderColor=$Clr.Border
    $btnSkip.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnSkip)

    $dlg.ClientSize = [System.Drawing.Size]::new(520, 468)

    if($DefaultMode -eq "ENABLE"){ $rbEnable.Checked = $true } else { $rbDisable.Checked = $true }
    & $UpdateList
    $rbEnable.Add_CheckedChanged($UpdateList)
    $dlg.Add_Shown({ $dlg.Activate() })
    $r = $dlg.ShowDialog()

    $mode = if($rbEnable.Checked){ "ENABLE" } else { "DISABLE" }
    $chosen = @()
    foreach($i in $ChkList.CheckedIndices){ $chosen += $state.Items[$i] }
    $dlg.Dispose()

    return [PSCustomObject]@{
        Action = $r        # OK = apply + script, Retry = script only, Cancel = skip
        Mode   = $mode     # "DISABLE" | "ENABLE"
        Jobs   = $chosen
    }
}

function Apply-AgentJobs {
    param(
        $Jobs,
        [bool]$Execute,
        [bool]$ScriptOut,
        [ValidateSet("DISABLE","ENABLE")][string]$Mode
    )
    $Jobs = @($Jobs)
    $val  = if($Mode -eq "ENABLE"){ 1 } else { 0 }
    $verb = if($Mode -eq "ENABLE"){ "Enable" } else { "Disable" }
    $undo = if($Mode -eq "ENABLE"){ "disable" } else { "enable" }
    $undoVal = 1 - $val
    $srcName = $TxtSrcSrv.Text.Trim()
    $stamp   = Get-Date -Format "yyyyMMdd_HHmmss"

    if($Jobs.Count -eq 0){
        Write-Log "  No jobs selected." $Clr.Dim; return
    }

    if($ScriptOut){
        $outPath = Join-Path (Get-JobFileFolder) ("pnfp_{0}AgentJobs_{1}.sql" -f $verb, $stamp)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("/*")
        [void]$sb.AppendLine("  pnfp Migration - $verb source SQL Agent jobs")
        [void]$sb.AppendLine("  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$sb.AppendLine("  Source    : $srcName")
        [void]$sb.AppendLine("  Jobs      : $($Jobs.Count)")
        [void]$sb.AppendLine("  Run this on the SOURCE server.")
        [void]$sb.AppendLine("*/")
        [void]$sb.AppendLine("")
        foreach($j in $Jobs){
            $n = $j.Name.Replace("'","''")
            [void]$sb.AppendLine("EXEC msdb.dbo.sp_update_job @job_name = N'$n', @enabled = $val;")
        }
        [void]$sb.AppendLine("GO")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("/*  To $undo these jobs again:")
        foreach($j in $Jobs){
            $n = $j.Name.Replace("'","''")
            [void]$sb.AppendLine("  EXEC msdb.dbo.sp_update_job @job_name = N'$n', @enabled = $undoVal;")
        }
        [void]$sb.AppendLine("*/")
        try {
            $sb.ToString() | Out-File -FilePath $outPath -Encoding UTF8 -Force
            Write-Log "  $verb script saved: $outPath" $Clr.Green
        } catch {
            Write-Log "  Script save failed: $($_.Exception.Message)" $Clr.Red
        }
    }

    if($Execute){
        if(-not $script:SrcConn -or $script:SrcConn.State -ne 'Open'){
            Write-Log "  Source not connected - cannot $($verb.ToLower()) jobs." $Clr.Red
        } else {
            $ok = 0; $bad = 0
            foreach($j in $Jobs){
                try {
                    Exec-Sql $script:SrcConn "EXEC msdb.dbo.sp_update_job @job_id = N'$($j.Id)', @enabled = $val"
                    Write-Log "  $verb`d: $($j.Name)" $Clr.Green
                    $ok++
                } catch {
                    Write-Log "  FAILED to $($verb.ToLower()) [$($j.Name)]: $($_.Exception.Message)" $Clr.Red
                    $bad++
                }
            }
            Write-Log "  Agent jobs - $($verb.ToLower())d: $ok  |  failed: $bad" $Clr.Dim
        }
    }

    # Remember which jobs were disabled so the rollback can re-enable only those
    if($Mode -eq "DISABLE" -and ($Execute -or $ScriptOut)){
        try {
            $srcTag  = Get-SafeFileName $srcName
            $listPath = Join-Path (Get-JobFileFolder) ("pnfp_DisabledJobs_{0}_{1}.csv" -f $srcTag, $stamp)
            $Jobs | Select-Object Id, Name | Export-Csv -Path $listPath -NoTypeInformation -Encoding UTF8
            Write-Log "  Disabled-job list saved for rollback: $listPath" $Clr.Dim
        } catch {
            Write-Log "  Could not save the disabled-job list: $($_.Exception.Message)" $Clr.Yellow
        }
    }
}

# Shows the jobs dialog, then applies / scripts the choice. DefaultMode picks the radio button.
function Invoke-AgentJobPrompt {
    param([ValidateSet("DISABLE","ENABLE")][string]$DefaultMode)

    if(-not $script:SrcConn -or $script:SrcConn.State -ne 'Open'){
        Write-Log "Source not connected - skipping SQL Agent job step." $Clr.Yellow
        return
    }
    try {
        $all = @(Get-AgentJobs $script:SrcConn)
    } catch {
        Write-Log "Could not read SQL Agent jobs on the source: $($_.Exception.Message)" $Clr.Red
        return
    }
    if($all.Count -eq 0){
        Write-Log "No SQL Agent jobs found on the source." $Clr.Dim
        return
    }

    $saved = @(Get-SavedDisabledJobList)
    $res = Show-AgentJobsDialog -AllJobs $all -DefaultMode $DefaultMode -SavedJobs $saved
    if($res -and $res.Action -ne [System.Windows.Forms.DialogResult]::Cancel){
        Write-Sep $(if($res.Mode -eq "ENABLE"){ "ROLLBACK: ENABLE SOURCE SQL AGENT JOBS" } else { "DISABLE SOURCE SQL AGENT JOBS" })
        $execute = ($res.Action -eq [System.Windows.Forms.DialogResult]::OK)
        Apply-AgentJobs -Jobs $res.Jobs -Execute $execute -ScriptOut $true -Mode $res.Mode
    } else {
        Write-Log "Source SQL Agent jobs: skipped." $Clr.Dim
    }
}

# ---------------------------------------------------------------
# RE-CHECK ORPHANED USERS ON THE DESTINATION (e.g. after the login script has been run)
# ---------------------------------------------------------------
function Show-DbPickDialog {
    param(
        [string]$Title,
        [string]$Info,
        [string[]]$Items,
        [string[]]$Checked = @(),
        [string]$OkText = "Run"
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = $Title
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock = [System.Windows.Forms.DockStyle]::Top; $hdr.Height = 46; $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Text = "  $Title"; $ht.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ht.ForeColor = $Clr.Orange; $ht.Font = $FontUIB
    $ht.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($ht)

    $info = New-Object System.Windows.Forms.Label
    $info.Text = $Info
    $info.Location = [System.Drawing.Point]::new(16, 56); $info.Size = [System.Drawing.Size]::new(442, 54)
    $info.ForeColor = $Clr.Dim; $info.Font = $FontSm
    $dlg.Controls.Add($info)

    $ChkList = New-Object System.Windows.Forms.CheckedListBox
    $ChkList.Location = [System.Drawing.Point]::new(16, 114); $ChkList.Size = [System.Drawing.Size]::new(442, 190)
    $ChkList.BackColor = $Clr.Input; $ChkList.ForeColor = $Clr.Text; $ChkList.Font = $FontMono
    $ChkList.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ChkList.CheckOnClick = $true
    foreach($i in $Items){ [void]$ChkList.Items.Add($i, ($Checked -contains $i)) }
    $dlg.Controls.Add($ChkList)

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text="All"; $btnAll.Location=[System.Drawing.Point]::new(16,312); $btnAll.Size=[System.Drawing.Size]::new(60,24)
    $btnAll.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnAll.BackColor=$Clr.Card; $btnAll.ForeColor=$Clr.Text; $btnAll.Font=$FontSm
    $btnAll.FlatAppearance.BorderColor=$Clr.Border
    $btnAll.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$true) } })
    $dlg.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text="None"; $btnNone.Location=[System.Drawing.Point]::new(82,312); $btnNone.Size=[System.Drawing.Size]::new(60,24)
    $btnNone.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnNone.BackColor=$Clr.Card; $btnNone.ForeColor=$Clr.Text; $btnNone.Font=$FontSm
    $btnNone.FlatAppearance.BorderColor=$Clr.Border
    $btnNone.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$false) } })
    $dlg.Controls.Add($btnNone)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text=$OkText; $btnOK.Location=[System.Drawing.Point]::new(276,348); $btnOK.Size=[System.Drawing.Size]::new(100,28)
    $btnOK.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnOK.BackColor=$Clr.Orange
    $btnOK.ForeColor=[System.Drawing.Color]::FromArgb(10,10,10); $btnOK.Font=$FontUIB
    $btnOK.FlatAppearance.BorderColor=$Clr.Orange
    $btnOK.DialogResult=[System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text="Cancel"; $btnCancel.Location=[System.Drawing.Point]::new(384,348); $btnCancel.Size=[System.Drawing.Size]::new(74,28)
    $btnCancel.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat; $btnCancel.BackColor=$Clr.Card
    $btnCancel.ForeColor=$Clr.Dim; $btnCancel.Font=$FontSm
    $btnCancel.FlatAppearance.BorderColor=$Clr.Border
    $btnCancel.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    $dlg.ClientSize = [System.Drawing.Size]::new(474, 392)
    $r = $dlg.ShowDialog()
    $chosen = @()
    foreach($i in $ChkList.CheckedIndices){ $chosen += $ChkList.Items[$i].ToString() }
    $dlg.Dispose()

    return [PSCustomObject]@{ Action = $r; Items = $chosen }
}

function Invoke-OrphanRecheck {
    if($script:Running){ return }
    if(-not $script:DstConn -or $script:DstConn.State -ne 'Open'){
        [System.Windows.Forms.MessageBox]::Show(
            "Connect to the destination server first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    try {
        $dt = Query-Table $script:DstConn @"
SELECT name FROM sys.databases
WHERE database_id > 4 AND state_desc = 'ONLINE' AND is_read_only = 0 AND source_database_id IS NULL
ORDER BY name;
"@
    } catch {
        Write-Log "Could not list destination databases: $($_.Exception.Message)" $Clr.Red
        return
    }
    $items = @(); foreach($row in $dt.Rows){ $items += [string]$row["name"] }
    if($items.Count -eq 0){ Write-Log "No online, writable user databases found on the destination." $Clr.Yellow; return }

    $pre = @()
    try { $pre = @(Get-SelectedDatabases | Where-Object { $items -contains $_ }) } catch { }

    $res = Show-DbPickDialog -Title "Re-check Orphaned Users" `
        -Info "Finds users whose login is missing on the destination and relinks them to the login with the same name. Run this after the login script has been run on the destination. Databases you selected for migration are pre-checked. READ_ONLY and offline databases are not listed." `
        -Items $items -Checked $pre -OkText "Re-check"
    if($res.Action -ne [System.Windows.Forms.DialogResult]::OK -or $res.Items.Count -eq 0){
        Write-Log "Orphaned user re-check: skipped." $Clr.Dim
        return
    }

    $script:Running = $true
    $BtnRun.Enabled = $false
    try {
        Write-Sep "RE-CHECK ORPHANED USERS (destination)"
        foreach($name in $res.Items){
            try {
                Fix-OrphanedUsers $script:DstConn $name
            } catch {
                Write-Log "  ORPHAN FIX FAILED for [$name]: $($_.Exception.Message)" $Clr.Red
            }
        }
        Write-Log "Orphaned user re-check complete." $Clr.Green
    } finally {
        $script:Running = $false
        $BtnRun.Enabled = $true
    }
}

# ---------------------------------------------------------------
# MODAL: DATABASE RENAME (restore as new name)
# ---------------------------------------------------------------
function Show-RenameDialog {
    param([string[]]$Databases)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Restore As - Database Names"
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock=$([System.Windows.Forms.DockStyle]::Top); $hdr.Height=46; $hdr.BackColor=$Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Text="  Restore As New Name (optional)"; $ht.Dock=$([System.Windows.Forms.DockStyle]::Fill)
    $ht.ForeColor=$Clr.Blue; $ht.Font=$FontUIB; $ht.TextAlign=$([System.Drawing.ContentAlignment]::MiddleLeft)
    $hdr.Controls.Add($ht)

    $info = New-Object System.Windows.Forms.Label
    $info.Text = "Leave the Target Name blank to restore with the original name.`nEnter a new name to restore the database under a different name (useful when restoring to the same instance)."
    $info.Location=[System.Drawing.Point]::new(14,56); $info.Size=[System.Drawing.Size]::new(520,36)
    $info.ForeColor=$Clr.Dim; $info.Font=$FontSm; $info.BackColor=[System.Drawing.Color]::Transparent
    $dlg.Controls.Add($info)

    # Header row
    $hdrSource = New-Object System.Windows.Forms.Label
    $hdrSource.Text="Source Database"; $hdrSource.Location=[System.Drawing.Point]::new(14,100)
    $hdrSource.Size=[System.Drawing.Size]::new(240,20); $hdrSource.ForeColor=$Clr.Blue
    $hdrSource.Font=$FontUIB; $hdrSource.BackColor=[System.Drawing.Color]::Transparent
    $dlg.Controls.Add($hdrSource)
    $hdrTarget = New-Object System.Windows.Forms.Label
    $hdrTarget.Text="Target Name (blank = same)"; $hdrTarget.Location=[System.Drawing.Point]::new(264,100)
    $hdrTarget.Size=[System.Drawing.Size]::new(270,20); $hdrTarget.ForeColor=$Clr.Blue
    $hdrTarget.Font=$FontUIB; $hdrTarget.BackColor=[System.Drawing.Color]::Transparent
    $dlg.Controls.Add($hdrTarget)

    $Y = 124
    $nameBoxes = @{}

    foreach($db in $Databases){
        $srcLbl = New-Object System.Windows.Forms.Label
        $srcLbl.Text=$db; $srcLbl.Location=[System.Drawing.Point]::new(14,$Y+3)
        $srcLbl.Size=[System.Drawing.Size]::new(240,22); $srcLbl.ForeColor=$Clr.Text
        $srcLbl.Font=$FontMono; $srcLbl.BackColor=[System.Drawing.Color]::Transparent
        $dlg.Controls.Add($srcLbl)

        $tgtBox = New-Object System.Windows.Forms.TextBox
        $tgtBox.Text=""; $tgtBox.Location=[System.Drawing.Point]::new(264,$Y)
        $tgtBox.Size=[System.Drawing.Size]::new(268,24); $tgtBox.BackColor=$Clr.Input
        $tgtBox.ForeColor=$Clr.Text; $tgtBox.Font=$FontMono
        $tgtBox.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
        $dlg.Controls.Add($tgtBox)

        $nameBoxes[$db] = $tgtBox
        $Y += 32
    }

    $Y += 8
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text="Proceed"; $btnOK.Location=[System.Drawing.Point]::new(340,$Y)
    $btnOK.Size=[System.Drawing.Size]::new(90,28); $btnOK.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnOK.BackColor=$Clr.Green; $btnOK.ForeColor=[System.Drawing.Color]::FromArgb(10,10,10)
    $btnOK.Font=$FontUIB; $btnOK.FlatAppearance.BorderColor=$Clr.Green
    $btnOK.DialogResult=[System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text="Cancel"; $btnCancel.Location=[System.Drawing.Point]::new(438,$Y)
    $btnCancel.Size=[System.Drawing.Size]::new(80,28); $btnCancel.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.BackColor=$Clr.Card; $btnCancel.ForeColor=$Clr.Dim; $btnCancel.Font=$FontSm
    $btnCancel.FlatAppearance.BorderColor=$Clr.Border
    $btnCancel.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel); $dlg.CancelButton=$btnCancel
    $dlg.ClientSize=[System.Drawing.Size]::new(548,$Y+44)
    $dlg.Add_Shown({ $dlg.Activate() })

    $r = $dlg.ShowDialog()
    $map = @{}
    if($r -eq [System.Windows.Forms.DialogResult]::OK){
        foreach($db in $Databases){
            $tgt = $nameBoxes[$db].Text.Trim()
            $map[$db] = if($tgt){ $tgt } else { $db }
        }
    }
    $dlg.Dispose()
    if($r -eq [System.Windows.Forms.DialogResult]::OK){ return $map }
    return $null
}

# ---------------------------------------------------------------
# MODAL: COMPATIBILITY LEVEL PICKER
# ---------------------------------------------------------------
function Show-CompatLevelDialog {
    param([int]$DstMajorVersion)

    # Map of every supported compat level to a friendly label
    $levels = @(
        [PSCustomObject]@{ Level=170; Label="170  -  SQL Server 2025" }
        [PSCustomObject]@{ Level=160; Label="160  -  SQL Server 2022" }
        [PSCustomObject]@{ Level=150; Label="150  -  SQL Server 2019" }
        [PSCustomObject]@{ Level=140; Label="140  -  SQL Server 2017" }
        [PSCustomObject]@{ Level=130; Label="130  -  SQL Server 2016" }
        [PSCustomObject]@{ Level=120; Label="120  -  SQL Server 2014" }
        [PSCustomObject]@{ Level=110; Label="110  -  SQL Server 2012" }
        [PSCustomObject]@{ Level=100; Label="100  -  SQL Server 2008 / 2008 R2" }
    )

    # Which level matches the detected destination version
    $detectedLevel = Get-LatestCompatLevel $DstMajorVersion

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Step 4 - Choose Compatibility Level"
    $dlg.Size            = [System.Drawing.Size]::new(460, 370)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock      = [System.Windows.Forms.DockStyle]::Top
    $hdr.Height    = 46
    $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Text      = "  Set Database Compatibility Level"
    $ht.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $ht.ForeColor = $Clr.Blue
    $ht.Font      = $FontUIB
    $ht.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($ht)

    $info = New-Object System.Windows.Forms.Label
    $info.Text      = "Select the compatibility level to apply to each restored database.`nDetected destination SQL version: $DstMajorVersion  (suggested: $detectedLevel)"
    $info.Location  = [System.Drawing.Point]::new(16, 58)
    $info.Size      = [System.Drawing.Size]::new(420, 36)
    $info.ForeColor = $Clr.Dim
    $info.Font      = $FontSm
    $dlg.Controls.Add($info)

    $Combo = New-Object System.Windows.Forms.ComboBox
    $Combo.Location      = [System.Drawing.Point]::new(16, 104)
    $Combo.Size          = [System.Drawing.Size]::new(414, 26)
    $Combo.BackColor     = $Clr.Input
    $Combo.ForeColor     = $Clr.Text
    $Combo.Font          = $FontUI
    $Combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Combo.FlatStyle     = [System.Windows.Forms.FlatStyle]::Flat
    foreach($lv in $levels){ $Combo.Items.Add($lv.Label) | Out-Null }
    # Pre-select the suggested level
    $suggestedIdx = [Array]::FindIndex($levels, [Predicate[object]]{ param($x) $x.Level -eq $detectedLevel })
    if($suggestedIdx -ge 0){ $Combo.SelectedIndex = $suggestedIdx } else { $Combo.SelectedIndex = 0 }
    $dlg.Controls.Add($Combo)

    # Explanation panel that updates when the selection changes
    $ExplainPanel = New-Object System.Windows.Forms.Panel
    $ExplainPanel.Location  = [System.Drawing.Point]::new(16, 142)
    $ExplainPanel.Size      = [System.Drawing.Size]::new(414, 138)
    $ExplainPanel.BackColor = $Clr.Card
    $dlg.Controls.Add($ExplainPanel)

    $ExplainLbl = New-Object System.Windows.Forms.Label
    $ExplainLbl.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $ExplainLbl.ForeColor = $Clr.Text
    $ExplainLbl.Font      = $FontSm
    $ExplainLbl.Padding   = [System.Windows.Forms.Padding]::new(10,8,10,8)
    $ExplainPanel.Controls.Add($ExplainLbl)

    $explanations = @{
        170 = "SQL Server 2025 level.`nEnables all 2025 features including the latest intelligent query processing improvements and new T-SQL enhancements.`nUse this when the destination is SQL 2025 and you want the full feature set."
        160 = "SQL Server 2022 level.`nEnables all 2022 features: intelligent query processing enhancements, parameter-sensitive plan optimisation, resumable operations.`nUse this when the destination is SQL 2022 and you want the full feature set."
        150 = "SQL Server 2019 level.`nEnables batch mode on rowstore, table variable deferred compilation, scalar UDF inlining.`nSafe choice for most migrations to SQL 2019 or later."
        140 = "SQL Server 2017 level.`nEnables adaptive query processing (adaptive joins, interleaved execution, batch mode memory grant feedback).`nUse when the destination is SQL 2017 and the application has been tested on 2017 behaviour."
        130 = "SQL Server 2016 level.`nEnables parallel DML on columnstore, batch mode for sorted data, new cardinality estimator.`nCommon compatibility target when migrating away from SQL 2012/2014."
        120 = "SQL Server 2014 level.`nUses the 2014 cardinality estimator. Safe middle ground if queries are known to regress on the 2016+ CE.`nConsider running Query Store after cutover and reviewing plan regressions."
        110 = "SQL Server 2012 level.`nOlder cardinality estimator. Use only when applications depend on specific 2012 query plan behaviour.`nPlan to test and upgrade to 130+ after stabilisation."
        100 = "SQL Server 2008/2008 R2 level.`nLegacy cardinality estimator. Use only as a temporary measure for problematic applications.`nThis level is unsupported in SQL 2022 - check destination version before applying."
    }

    $updateExplain = {
        $sel = $Combo.SelectedIndex
        if($sel -ge 0 -and $sel -lt $levels.Count){
            $lv = $levels[$sel].Level
            $txt = if($explanations.ContainsKey($lv)){ $explanations[$lv] } else { "Select a level to see details." }
            $ExplainLbl.Text = $txt
            $ExplainLbl.ForeColor = if($lv -eq $detectedLevel){ $Clr.Green } else { $Clr.Yellow }
        }
    }
    & $updateExplain
    $Combo.Add_SelectedIndexChanged($updateExplain)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text      = "Apply This Level"
    $btnOK.Location  = [System.Drawing.Point]::new(200, 298)
    $btnOK.Size      = [System.Drawing.Size]::new(130, 28)
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.BackColor = $Clr.Blue
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.Font      = $FontUIB
    $btnOK.FlatAppearance.BorderColor = $Clr.Blue
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Text      = "Skip This Step"
    $btnSkip.Location  = [System.Drawing.Point]::new(340, 298)
    $btnSkip.Size      = [System.Drawing.Size]::new(100, 28)
    $btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkip.BackColor = $Clr.Card
    $btnSkip.ForeColor = $Clr.Dim
    $btnSkip.Font      = $FontSm
    $btnSkip.FlatAppearance.BorderColor = $Clr.Border
    $btnSkip.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnSkip)

    $dlg.Add_Shown({ $dlg.Activate() })
    $r = $dlg.ShowDialog()

    $chosen = $null
    if($r -eq [System.Windows.Forms.DialogResult]::OK){
        $sel = $Combo.SelectedIndex
        if($sel -ge 0){ $chosen = $levels[$sel].Level }
    }
    $dlg.Dispose()
    return $chosen    # $null = skip
}

# ---------------------------------------------------------------
# MODAL 1: CHOOSE MODE (PRE / POST)
# ---------------------------------------------------------------
function Show-ModeDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "pnfp Migration - Select Mode"
    $dlg.Size            = [System.Drawing.Size]::new(540, 340)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock      = [System.Windows.Forms.DockStyle]::Top
    $hdr.Height    = 60
    $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)

    $htl = New-Object System.Windows.Forms.Label
    $htl.Text      = "  pnfp Migration Tool"
    $htl.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $htl.ForeColor = $Clr.Purple
    $htl.Font      = $FontTitle
    $htl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($htl)

    $ql = New-Object System.Windows.Forms.Label
    $ql.Text      = "What type of migration run is this?"
    $ql.Location  = [System.Drawing.Point]::new(30, 76)
    $ql.Size      = [System.Drawing.Size]::new(480, 22)
    $ql.ForeColor = $Clr.Text
    $ql.Font      = $FontUIB
    $dlg.Controls.Add($ql)

    # PRE card
    $cardPre = New-Object System.Windows.Forms.Panel
    $cardPre.Location  = [System.Drawing.Point]::new(30, 108)
    $cardPre.Size      = [System.Drawing.Size]::new(220, 130)
    $cardPre.BackColor = $Clr.Card
    $cardPre.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $dlg.Controls.Add($cardPre)

    $accPre = New-Object System.Windows.Forms.Panel
    $accPre.Dock      = [System.Windows.Forms.DockStyle]::Top
    $accPre.Height    = 4
    $accPre.BackColor = $Clr.Green
    $cardPre.Controls.Add($accPre)

    $lPre1 = New-Object System.Windows.Forms.Label
    $lPre1.Text      = "PRE-MIGRATION"
    $lPre1.Location  = [System.Drawing.Point]::new(10, 14)
    $lPre1.Size      = [System.Drawing.Size]::new(200, 22)
    $lPre1.ForeColor = $Clr.Green
    $lPre1.Font      = $FontUIB
    $lPre1.BackColor = [System.Drawing.Color]::Transparent
    $cardPre.Controls.Add($lPre1)

    $lPre2 = New-Object System.Windows.Forms.Label
    $lPre2.Text      = "- Copy-only backup on source`n- Restore to destination`n- Update statistics`n- Set compatibility level`n- Fix orphaned users`n- Source Agent jobs (optional)"
    $lPre2.Location  = [System.Drawing.Point]::new(10, 38)
    $lPre2.Size      = [System.Drawing.Size]::new(200, 90)
    $lPre2.ForeColor = $Clr.Dim
    $lPre2.Font      = $FontSm
    $lPre2.BackColor = [System.Drawing.Color]::Transparent
    $cardPre.Controls.Add($lPre2)

    # POST card
    $cardPost = New-Object System.Windows.Forms.Panel
    $cardPost.Location  = [System.Drawing.Point]::new(280, 108)
    $cardPost.Size      = [System.Drawing.Size]::new(220, 130)
    $cardPost.BackColor = $Clr.Card
    $cardPost.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $dlg.Controls.Add($cardPost)

    $accPost = New-Object System.Windows.Forms.Panel
    $accPost.Dock      = [System.Windows.Forms.DockStyle]::Top
    $accPost.Height    = 4
    $accPost.BackColor = $Clr.Orange
    $cardPost.Controls.Add($accPost)

    $lPost1 = New-Object System.Windows.Forms.Label
    $lPost1.Text      = "MIGRATION-DAY"
    $lPost1.Location  = [System.Drawing.Point]::new(10, 14)
    $lPost1.Size      = [System.Drawing.Size]::new(200, 22)
    $lPost1.ForeColor = $Clr.Orange
    $lPost1.Font      = $FontUIB
    $lPost1.BackColor = [System.Drawing.Color]::Transparent
    $cardPost.Controls.Add($lPost1)

    $lPost2 = New-Object System.Windows.Forms.Label
    $lPost2.Text      = "- Backup source + restore`n  (REPLACE / RECOVERY options)`n- Compat level, stats, orphans`n- CHECKSUM check on user DBs`n- Source READ_ONLY / job options"
    $lPost2.Location  = [System.Drawing.Point]::new(10, 38)
    $lPost2.Size      = [System.Drawing.Size]::new(200, 90)
    $lPost2.ForeColor = $Clr.Dim
    $lPost2.Font      = $FontSm
    $lPost2.BackColor = [System.Drawing.Color]::Transparent
    $cardPost.Controls.Add($lPost2)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text      = "Cancel"
    $cancelBtn.Location  = [System.Drawing.Point]::new(210, 260)
    $cancelBtn.Size      = [System.Drawing.Size]::new(90, 28)
    $cancelBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelBtn.BackColor = $Clr.Card
    $cancelBtn.ForeColor = $Clr.Dim
    $cancelBtn.Font      = $FontSm
    $cancelBtn.FlatAppearance.BorderColor = $Clr.Border
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($cancelBtn)
    $dlg.CancelButton = $cancelBtn

    $script:ModeResult = ""

    $LaunchPre = {
        $script:ModeResult = "PRE"
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    }
    $LaunchPost = {
        $script:ModeResult = "POST"
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    }
    $hoverOn  = [System.Drawing.Color]::FromArgb(42,48,66)
    $hoverOff = $Clr.Card

    foreach($ctrl in @($cardPre, $lPre1, $lPre2, $accPre)){
        $ctrl.Add_Click($LaunchPre)
        $ctrl.Add_MouseEnter({ $cardPre.BackColor=$hoverOn }.GetNewClosure())
        $ctrl.Add_MouseLeave({ $cardPre.BackColor=$hoverOff }.GetNewClosure())
    }
    foreach($ctrl in @($cardPost, $lPost1, $lPost2, $accPost)){
        $ctrl.Add_Click($LaunchPost)
        $ctrl.Add_MouseEnter({ $cardPost.BackColor=$hoverOn }.GetNewClosure())
        $ctrl.Add_MouseLeave({ $cardPost.BackColor=$hoverOff }.GetNewClosure())
    }

    $r = $dlg.ShowDialog()
    $dlg.Dispose()
    if($r -eq [System.Windows.Forms.DialogResult]::OK){ return $script:ModeResult }
    return $null
}

# ---------------------------------------------------------------
# MODAL 2: MIGRATION-DAY RESTORE OPTIONS
# ---------------------------------------------------------------
function Show-RestoreOptionsDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Migration Day - Restore Options"
    $dlg.Size            = [System.Drawing.Size]::new(560, 620)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    # Header
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock      = [System.Windows.Forms.DockStyle]::Top
    $hdr.Height    = 46
    $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Text      = "  Restore Options"
    $ht.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $ht.ForeColor = $Clr.Orange
    $ht.Font      = $FontUIB
    $ht.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($ht)

    function Add-Section {
        param($Parent, [string]$Title, [int]$Y)
        $p = New-Object System.Windows.Forms.Panel
        $p.Location  = [System.Drawing.Point]::new(14,$Y)
        $p.Size      = [System.Drawing.Size]::new(520,22)
        $p.BackColor = [System.Drawing.Color]::FromArgb(20,50,100)
        $l = New-Object System.Windows.Forms.Label
        $l.Text=$Title; $l.Dock=[System.Windows.Forms.DockStyle]::Fill
        $l.ForeColor=$Clr.Blue; $l.Font=$FontSec
        $l.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft
        $p.Controls.Add($l); $Parent.Controls.Add($p)
        return $Y+28
    }
    function Add-Opt {
        param($Parent,[string]$Text,[bool]$Checked,[int]$X,[int]$Y,[int]$W=490)
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text=$Text; $chk.Checked=$Checked
        $chk.Location=[System.Drawing.Point]::new($X,$Y); $chk.Size=[System.Drawing.Size]::new($W,22)
        $chk.ForeColor=$Clr.Text; $chk.Font=$FontUI; $chk.BackColor=[System.Drawing.Color]::Transparent
        $Parent.Controls.Add($chk); return $chk
    }
    function Add-Desc {
        param($Parent,[string]$Text,[int]$Y,[int]$Lines=2)
        $l = New-Object System.Windows.Forms.Label
        $l.Text=$Text; $l.Location=[System.Drawing.Point]::new(36,$Y)
        $l.Size=[System.Drawing.Size]::new(492,($Lines*16)); $l.ForeColor=$Clr.Dim
        $l.Font=$FontSm; $l.BackColor=[System.Drawing.Color]::Transparent
        $Parent.Controls.Add($l)
        return $Y+($Lines*16)+4
    }

    $Y = 56

    # ---- RESTORE OPTIONS ----
    $Y = Add-Section $dlg "  RESTORE OPTIONS" $Y
    $Y += 4

    $chkReplace = Add-Opt $dlg "Overwrite the existing database  (WITH REPLACE)" $true 18 $Y
    $Y += 24
    $Y = Add-Desc $dlg "Use when the database already exists at the destination and you want to completely replace it." $Y 2

    $chkKeepReplication = Add-Opt $dlg "Preserve replication settings  (WITH KEEP_REPLICATION)" $false 18 $Y
    $Y += 24
    $Y = Add-Desc $dlg "Keeps replication settings when restoring a published database. Only needed if the destination participates in replication." $Y 2

    $chkRestrictedUser = Add-Opt $dlg "Restrict access to the restored database  (WITH RESTRICTED_USER)" $false 18 $Y
    $Y += 24
    $Y = Add-Desc $dlg "Limits access to db_owner, dbcreator, or sysadmin after restore. Useful for validation before opening to users." $Y 2

    $Y += 8

    # ---- RECOVERY STATE ----
    $Y = Add-Section $dlg "  RECOVERY STATE" $Y
    $Y += 8

    $RadioPanel = New-Object System.Windows.Forms.Panel
    $RadioPanel.Location  = [System.Drawing.Point]::new(14,$Y)
    $RadioPanel.Size      = [System.Drawing.Size]::new(524,148)
    $RadioPanel.BackColor = [System.Drawing.Color]::Transparent
    $dlg.Controls.Add($RadioPanel)

    $rbRecovery = New-Object System.Windows.Forms.RadioButton
    $rbRecovery.Text      = "RESTORE WITH RECOVERY  (default)"
    $rbRecovery.Location  = [System.Drawing.Point]::new(4, 4)
    $rbRecovery.Size      = [System.Drawing.Size]::new(510, 20)
    $rbRecovery.Checked   = $true
    $rbRecovery.ForeColor = $Clr.Text; $rbRecovery.Font=$FontUI
    $rbRecovery.BackColor = [System.Drawing.Color]::Transparent
    $RadioPanel.Controls.Add($rbRecovery)

    $lbRecovery = New-Object System.Windows.Forms.Label
    $lbRecovery.Text      = "Leaves the database ready to use by rolling back uncommitted transactions.`n  Additional transaction logs cannot be restored after this."
    $lbRecovery.Location  = [System.Drawing.Point]::new(22,24); $lbRecovery.Size=[System.Drawing.Size]::new(494,32)
    $lbRecovery.ForeColor = $Clr.Dim; $lbRecovery.Font=$FontSm; $lbRecovery.BackColor=[System.Drawing.Color]::Transparent
    $RadioPanel.Controls.Add($lbRecovery)

    $rbNoRecovery = New-Object System.Windows.Forms.RadioButton
    $rbNoRecovery.Text      = "RESTORE WITH NORECOVERY"
    $rbNoRecovery.Location  = [System.Drawing.Point]::new(4, 62)
    $rbNoRecovery.Size      = [System.Drawing.Size]::new(510, 20)
    $rbNoRecovery.Checked   = $false
    $rbNoRecovery.ForeColor = $Clr.Text; $rbNoRecovery.Font=$FontUI
    $rbNoRecovery.BackColor = [System.Drawing.Color]::Transparent
    $RadioPanel.Controls.Add($rbNoRecovery)

    $lbNoRecovery = New-Object System.Windows.Forms.Label
    $lbNoRecovery.Text      = "Leaves the database in a restoring state so additional transaction log backups can be applied."
    $lbNoRecovery.Location  = [System.Drawing.Point]::new(22,82); $lbNoRecovery.Size=[System.Drawing.Size]::new(494,16)
    $lbNoRecovery.ForeColor = $Clr.Dim; $lbNoRecovery.Font=$FontSm; $lbNoRecovery.BackColor=[System.Drawing.Color]::Transparent
    $RadioPanel.Controls.Add($lbNoRecovery)

    $rbStandby = New-Object System.Windows.Forms.RadioButton
    $rbStandby.Text      = "RESTORE WITH STANDBY"
    $rbStandby.Location  = [System.Drawing.Point]::new(4, 104)
    $rbStandby.Size      = [System.Drawing.Size]::new(510, 20)
    $rbStandby.Checked   = $false
    $rbStandby.ForeColor = $Clr.Text; $rbStandby.Font=$FontUI
    $rbStandby.BackColor = [System.Drawing.Color]::Transparent
    $RadioPanel.Controls.Add($rbStandby)

    $lbStandby = New-Object System.Windows.Forms.Label
    $lbStandby.Text      = "Leaves the database in read-only mode between log restores (useful for log shipping)."
    $lbStandby.Location  = [System.Drawing.Point]::new(22,124); $lbStandby.Size=[System.Drawing.Size]::new(494,16)
    $lbStandby.ForeColor = $Clr.Dim; $lbStandby.Font=$FontSm; $lbStandby.BackColor=[System.Drawing.Color]::Transparent
    $RadioPanel.Controls.Add($lbStandby)

    $Y += 156

    # ---- SERVER CONNECTIONS ----
    $Y = Add-Section $dlg "  SERVER CONNECTIONS" $Y
    $Y += 4

    $chkCloseConnections = Add-Opt $dlg "Close existing connections to the destination database" $true 18 $Y
    $Y += 24
    $Y = Add-Desc $dlg "Kills any active sessions connected to the database before restore begins, preventing 'database in use' errors." $Y 2

    $Y += 4

    # Buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text      = "Proceed"
    $btnOK.Location  = [System.Drawing.Point]::new(348, $Y)
    $btnOK.Size      = [System.Drawing.Size]::new(96, 28)
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.BackColor = $Clr.Orange
    $btnOK.ForeColor = [System.Drawing.Color]::FromArgb(10,10,10)
    $btnOK.Font      = $FontUIB
    $btnOK.FlatAppearance.BorderColor = $Clr.Orange
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text      = "Cancel"
    $btnCancel.Location  = [System.Drawing.Point]::new(452, $Y)
    $btnCancel.Size      = [System.Drawing.Size]::new(82, 28)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.BackColor = $Clr.Card
    $btnCancel.ForeColor = $Clr.Dim
    $btnCancel.Font      = $FontSm
    $btnCancel.FlatAppearance.BorderColor = $Clr.Border
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $dlg.Add_Shown({ $dlg.Activate() })

    # Resize to fit content
    $dlg.ClientSize = [System.Drawing.Size]::new(548, $Y + 40)

    $r = $dlg.ShowDialog()
    $res = $null
    if($r -eq [System.Windows.Forms.DialogResult]::OK){
        $recoveryMode = if($rbRecovery.Checked){"RECOVERY"}elseif($rbNoRecovery.Checked){"NORECOVERY"}else{"STANDBY"}
        $res = [PSCustomObject]@{
            Replace          = $chkReplace.Checked
            KeepReplication  = $chkKeepReplication.Checked
            RestrictedUser   = $chkRestrictedUser.Checked
            RecoveryMode     = $recoveryMode       # "RECOVERY" | "NORECOVERY" | "STANDBY"
            CloseConnections = $chkCloseConnections.Checked
        }
    }
    $dlg.Dispose()
    return $res
}

# ---------------------------------------------------------------
# CORE OPERATIONS
# ---------------------------------------------------------------

# Backup location is fixed: <share root>\<destination server>\PRE_MIGRATION (or MIGRATION_DAY)\<database>
$script:BackupShareRoot = '\\pstces002.isilon.snv.net\SQLBKUP'
$script:BackupSubFolder = 'PRE_MIGRATION'
$script:PostBackupSubFolder = 'MIGRATION_DAY'   # used in MIGRATION-DAY mode

# Timeout for the long operations (backup, restore, update statistics), in seconds. 0 = no limit.
$script:LongOpTimeoutSec = 0
# Databases with more data than this (GB) trigger the "how many backup files" prompt at Run.
$script:StripeThresholdGB = 300
# Chosen number of backup files per database for this run. Databases not listed use a single file.
$script:StripePlan = @{}
# Remembers which files make up each backup set, so the restore reads all stripes.
$script:BackupFileSets = @{}

function Format-Elapsed {
    param([TimeSpan]$Ts)
    return ('{0:00}:{1:00}:{2:00}' -f [int][math]::Floor($Ts.TotalHours), $Ts.Minutes, $Ts.Seconds)
}

# Runs a long SQL statement in the background so the window stays responsive.
# Logs percent complete (from sys.dm_exec_requests) and honours the Cancel button.
function Invoke-LongSql {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [string]$Sql,
        [string]$Label = "operation"
    )

    # Optional second connection to read progress. Only used with Windows authentication,
    # because an open SQL-auth connection string does not keep the password.
    $spid = $null
    $mon  = $null
    try { $spid = [int](Query-Scalar $Conn "SELECT @@SPID") } catch { }
    if($spid -and $Conn.ConnectionString -match 'Integrated Security\s*=\s*(True|SSPI)'){
        try { $mon = Open-Conn $Conn.ConnectionString } catch { $mon = $null }
    }

    $cmd = New-Object System.Data.SqlClient.SqlCommand $Sql,$Conn
    $cmd.CommandTimeout = $script:LongOpTimeoutSec
    $sw   = [System.Diagnostics.Stopwatch]::StartNew()
    $task = $cmd.ExecuteNonQueryAsync()

    $lastPoll   = Get-Date
    $lastLogAt  = Get-Date
    $lastLogPct = -100.0
    $cancelSent = $false
    try {
        while(-not $task.IsCompleted){
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 250

            if($script:Cancel -and -not $cancelSent){
                $cancelSent = $true
                Write-Log "  Cancel requested - stopping $Label..." $Clr.Orange
                try { $cmd.Cancel() } catch { }
            }

            if(((Get-Date) - $lastPoll).TotalSeconds -ge 10){
                $lastPoll = Get-Date
                $pct = 0.0; $etaMs = 0
                if($mon){
                    try {
                        $dt = Query-Table $mon "SELECT percent_complete, estimated_completion_time FROM sys.dm_exec_requests WHERE session_id = $spid"
                        if($dt.Rows.Count -gt 0){
                            $pct   = [double]$dt.Rows[0]["percent_complete"]
                            $etaMs = [int64]$dt.Rows[0]["estimated_completion_time"]
                        }
                    } catch { }
                }
                $elapsed = Format-Elapsed $sw.Elapsed
                if($pct -gt 0){
                    $eta = Format-Elapsed ([TimeSpan]::FromMilliseconds($etaMs))
                    Set-Status ("{0}: {1:N0}% - elapsed {2}, about {3} left" -f $Label,$pct,$elapsed,$eta) $Clr.Yellow
                    if(($pct - $lastLogPct) -ge 5 -or ((Get-Date) - $lastLogAt).TotalSeconds -ge 120){
                        Write-Log ("  {0}: {1:N0}% complete, elapsed {2}, about {3} left" -f $Label,$pct,$elapsed,$eta) $Clr.Dim
                        $lastLogPct = $pct; $lastLogAt = Get-Date
                    }
                } else {
                    Set-Status ("{0}: running - elapsed {1}" -f $Label,$elapsed) $Clr.Yellow
                    if(((Get-Date) - $lastLogAt).TotalSeconds -ge 120){
                        Write-Log "  $Label`: still running, elapsed $elapsed" $Clr.Dim
                        $lastLogAt = Get-Date
                    }
                }
            }
        }
    } finally {
        if($mon){ try { $mon.Close(); $mon.Dispose() } catch { } }
    }

    if($task.IsFaulted){
        if($cancelSent){ throw "Cancelled by user" }
        throw $task.Exception.InnerException
    }
    if($task.IsCanceled){ throw "Cancelled by user" }
    Write-Log "  $Label finished in $(Format-Elapsed $sw.Elapsed)" $Clr.Dim
}

# Return the list of backup files for a database, based on the choice made at Run time.
function Get-BackupStripeFiles {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$DbName,[string]$BasePath)
    $n = 1
    if($script:StripePlan.ContainsKey($DbName)){ $n = [int]$script:StripePlan[$DbName] }
    if($n -lt 1){ $n = 1 }
    if($n -gt 8){ $n = 8 }
    Write-Log "  Backup files: $n" $Clr.Dim
    if($n -eq 1){ return @($BasePath) }
    $dir  = Split-Path $BasePath -Parent
    $name = [System.IO.Path]::GetFileNameWithoutExtension($BasePath)
    $ext  = [System.IO.Path]::GetExtension($BasePath)
    $list = @()
    for($i = 1; $i -le $n; $i++){ $list += (Join-Path $dir ("{0}_{1}of{2}{3}" -f $name,$i,$n,$ext)) }
    return $list
}

# Prompt for the number of backup files per large database.
function Show-StripeDialog {
    param($LargeDbs, [int]$Suggested, [int]$Threshold)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Large Databases - Backup Files"
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock = [System.Windows.Forms.DockStyle]::Top; $hdr.Height = 46; $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Text = "  Large Databases Detected"; $ht.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ht.ForeColor = $Clr.Orange; $ht.Font = $FontUIB
    $ht.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($ht)

    $info = New-Object System.Windows.Forms.Label
    $info.Text      = "These databases have more than $Threshold GB of data. Choose how many files each one is backed up to. More files can speed up big backups, but every file is needed to restore. Databases under $Threshold GB use a single file."
    $info.Location  = [System.Drawing.Point]::new(16, 58)
    $info.Size      = [System.Drawing.Size]::new(442, 66)
    $info.ForeColor = $Clr.Dim; $info.Font = $FontSm
    $dlg.Controls.Add($info)

    $rows = [Math]::Min(8, [Math]::Max(2, @($LargeDbs).Count))
    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Location    = [System.Drawing.Point]::new(16, 130)
    $lb.Size        = [System.Drawing.Size]::new(442, ($rows * 18 + 8))
    $lb.BackColor   = $Clr.Input; $lb.ForeColor = $Clr.Text; $lb.Font = $FontMono
    $lb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    foreach($d in $LargeDbs){ [void]$lb.Items.Add(("{0,-38} {1,10:N1} GB" -f $d.Name, $d.GB)) }
    $dlg.Controls.Add($lb)

    $y = 130 + $lb.Height + 14
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Backup files per large database:"
    $lbl.Location = [System.Drawing.Point]::new(16, ($y + 3)); $lbl.Size = [System.Drawing.Size]::new(230, 20)
    $lbl.ForeColor = $Clr.Text; $lbl.Font = $FontUI
    $dlg.Controls.Add($lbl)

    $num = New-Object System.Windows.Forms.NumericUpDown
    $num.Minimum = 1; $num.Maximum = 8; $num.Value = [Math]::Min(8, [Math]::Max(1, $Suggested))
    $num.Location = [System.Drawing.Point]::new(252, $y); $num.Size = [System.Drawing.Size]::new(60, 24)
    $num.BackColor = $Clr.Input; $num.ForeColor = $Clr.Text; $num.Font = $FontUI
    $dlg.Controls.Add($num)

    $y += 32
    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Suggested: $Suggested (about one file per 200 GB of the largest database, max 8)."
    $hint.Location = [System.Drawing.Point]::new(16, $y); $hint.Size = [System.Drawing.Size]::new(442, 18)
    $hint.ForeColor = $Clr.Dim; $hint.Font = $FontSm
    $dlg.Controls.Add($hint)

    $y += 34
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Proceed"; $btnOK.Location = [System.Drawing.Point]::new(268, $y); $btnOK.Size = [System.Drawing.Size]::new(96, 28)
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.BackColor = $Clr.Orange; $btnOK.ForeColor = [System.Drawing.Color]::FromArgb(10,10,10); $btnOK.Font = $FontUIB
    $btnOK.FlatAppearance.BorderColor = $Clr.Orange
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)
    $dlg.AcceptButton = $btnOK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel run"; $btnCancel.Location = [System.Drawing.Point]::new(372, $y); $btnCancel.Size = [System.Drawing.Size]::new(86, 28)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.BackColor = $Clr.Card; $btnCancel.ForeColor = $Clr.Dim; $btnCancel.Font = $FontSm
    $btnCancel.FlatAppearance.BorderColor = $Clr.Border
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    $dlg.ClientSize = [System.Drawing.Size]::new(474, ($y + 46))
    $r = $dlg.ShowDialog()
    $val = [int]$num.Value
    $dlg.Dispose()
    if($r -eq [System.Windows.Forms.DialogResult]::OK){ return $val }
    return $null
}

# Called at Run. Finds selected databases over the threshold and asks how many backup files to use.
# Returns $false if the user cancels the run.
function Get-StripePlan {
    param([string[]]$Databases)
    $script:StripePlan = @{}
    $threshold = [double]$script:StripeThresholdGB
    $large = @()
    try {
        $dt = Query-Table $script:SrcConn "SELECT DB_NAME(database_id) AS name, SUM(CONVERT(BIGINT, size)) * 8 / 1024.0 / 1024.0 AS gb FROM sys.master_files WHERE type = 0 GROUP BY database_id"
        foreach($row in $dt.Rows){
            $nm = [string]$row["name"]; $gb = [double]$row["gb"]
            if(($Databases -contains $nm) -and $gb -gt $threshold){
                $large += [PSCustomObject]@{ Name = $nm; GB = $gb }
            }
        }
    } catch {
        Write-Log "  Could not check database sizes ($($_.Exception.Message)). Using single-file backups." $Clr.Yellow
        return $true
    }

    if($large.Count -eq 0){
        Write-Log "No selected database is over $($script:StripeThresholdGB) GB - single-file backups." $Clr.Dim
        return $true
    }

    $largest   = ($large | Measure-Object -Property GB -Maximum).Maximum
    $suggested = [int][Math]::Min(8, [Math]::Max(2, [Math]::Ceiling($largest / 200)))
    $n = Show-StripeDialog -LargeDbs ($large | Sort-Object GB -Descending) -Suggested $suggested -Threshold $script:StripeThresholdGB
    if($null -eq $n){ return $false }

    foreach($d in $large){ $script:StripePlan[$d.Name] = $n }
    Write-Log "Backup files per large database: $n  ($(($large | ForEach-Object { $_.Name }) -join ', '))" $Clr.Blue
    return $true
}

# Note if the database is TDE-encrypted and whether the destination has the certificate.
# Informational only - never stops the run.
function Test-TdeCertificate {
    param([System.Data.SqlClient.SqlConnection]$SrcConn,[System.Data.SqlClient.SqlConnection]$DstConn,[string]$DbName)
    try {
        $thumb = Query-Scalar $SrcConn "SELECT CONVERT(VARCHAR(100), encryptor_thumbprint, 1) FROM sys.dm_database_encryption_keys WHERE database_id = DB_ID(N'$DbName') AND encryptor_type = 'CERTIFICATE'"
        if(-not $thumb -or $thumb -is [System.DBNull]){ return }
        Write-Log "  [$DbName] is TDE-encrypted. The certificate and private key must exist on the destination before the restore." $Clr.Yellow
        $found = Query-Scalar $DstConn "SELECT COUNT(*) FROM master.sys.certificates WHERE thumbprint = CONVERT(VARBINARY(32), '$thumb', 1)"
        if([int]$found -gt 0){
            Write-Log "  Matching TDE certificate found on destination." $Clr.Green
        } else {
            Write-Log "  WARNING: no matching TDE certificate found on the destination. The restore will fail until it is restored there." $Clr.Red
        }
    } catch { }
}

function Get-BackupPath {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$DbName,[string]$SubFolder = $script:BackupSubFolder)

    if(-not ($script:DstConn -and $script:DstConn.State -eq 'Open')){
        Write-Log "  Connect to the destination server first. The backup folder is named after it." $Clr.Red
        throw "Destination server not connected"
    }

    $dstName = [string](Query-Scalar $script:DstConn "SELECT @@SERVERNAME")
    # Named instances (HOST\INST) cannot be used as a single folder name
    $dstName = $dstName.Replace('\','_').Trim()
    if(-not $dstName){ throw "Could not determine destination server name" }

    # One folder per database inside the PRE_MIGRATION / MIGRATION_DAY folder
    $dbFolder = ($DbName -replace '[\\/:*?"<>|]', '_').Trim()
    if(-not $dbFolder){ throw "Invalid database name for folder: $DbName" }
    $dir = Join-Path (Join-Path (Join-Path $script:BackupShareRoot $dstName) $SubFolder) $dbFolder
    if(-not (Test-Path $dir)){
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        catch {
            Write-Log "  Could not create backup folder $dir : $($_.Exception.Message)" $Clr.Red
            throw "Backup folder not available: $dir"
        }
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    return "$dir\${DbName}_COPYONLY_$stamp.bak"
}

# Copy-only backup. Returns the file path used.
function Backup-DatabaseCopyOnly {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [string]$DbName,
        [string]$BackupPath
    )
    Write-Log "  Backing up [$DbName] to: $BackupPath" $Clr.Dim
    # Verify the target directory exists before attempting backup
    $backupDir = Split-Path $BackupPath -Parent
    if(-not (Test-Path $backupDir)){
        Write-Log "  Backup directory does not exist: $backupDir" $Clr.Red
        Write-Log "  Check access to $($script:BackupShareRoot) from this machine." $Clr.Red
        throw "Backup directory not found: $backupDir"
    }

    $bkFiles = @(Get-BackupStripeFiles $Conn $DbName $BackupPath)
    $script:BackupFileSets[$BackupPath] = $bkFiles
    $toClause = ($bkFiles | ForEach-Object { "DISK = N'$_'" }) -join ",`n   "

    # MAXTRANSFERSIZE above 64 KB also lets compression work on TDE-encrypted databases
    $sql = @"
BACKUP DATABASE [$DbName]
TO $toClause
WITH COPY_ONLY, COMPRESSION, MAXTRANSFERSIZE = 4194304, INIT,
     NAME = N'${DbName} - pnfp Copy-Only Backup';
"@
    try {
        Invoke-LongSql $Conn $sql "Backup of [$DbName]"
    } catch {
        # Do not leave partial multi-GB files on the share
        foreach($f in $bkFiles){ try { if(Test-Path $f){ Remove-Item $f -Force } } catch { } }
        Write-Log "  Removed partial backup file(s)." $Clr.Dim
        throw
    }
    Write-Log "  Backup complete ($($bkFiles.Count) file(s)): $BackupPath" $Clr.Green
    return $BackupPath
}

# Restore. Options control REPLACE, RECOVERY and optional rename.
function Restore-Database {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [string]$DbName,              # source/original database name (used for backup reading)
        [string]$BackupPath,
        [string]$TargetName      = "", # if non-empty, restore AS this name instead
        [bool]$WithReplace       = $true,
        [bool]$WithRecovery      = $true,
        [string]$RecoveryMode    = "RECOVERY",
        [bool]$KeepReplication   = $false,
        [bool]$RestrictedUser    = $false,
        [bool]$CloseConnections  = $true
    )

    # Use TargetName for the restored database, fall back to DbName
    $restoreName = if($TargetName -and $TargetName.Trim()){ $TargetName.Trim() } else { $DbName }
    if($restoreName -ne $DbName){
        Write-Log "  Restoring as new name: [$restoreName]" $Clr.Blue
    }

    # Kill existing connections on the TARGET database
    $setSingleUser = $false
    if($CloseConnections){
        Write-Log "  Closing existing connections on [$restoreName]..." $Clr.Dim
        try { $Conn.ChangeDatabase("master") } catch { }
        try {
            $dbExists = Query-Scalar $Conn "SELECT COUNT(1) FROM sys.databases WHERE name = N'$restoreName'"
            if([int]$dbExists -gt 0){
                Exec-Sql $Conn "ALTER DATABASE [$restoreName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
                $setSingleUser = $true
            } else {
                Write-Log "  [$restoreName] does not yet exist - no connections to close." $Clr.Dim
            }
        } catch {
            Write-Log "  Warning: could not set SINGLE_USER on [$restoreName]: $($_.Exception.Message)" $Clr.Yellow
        }
    }

    # All stripes of the backup set (single file unless the backup was striped)
    $bkFiles    = if($script:BackupFileSets.ContainsKey($BackupPath)){ @($script:BackupFileSets[$BackupPath]) } else { @($BackupPath) }
    $fromClause = ($bkFiles | ForEach-Object { "DISK = N'$_'" }) -join ", "

    # Read logical file names using SqlDataAdapter (more reliable than DataTable.Load)
    Write-Log "  Reading backup header..." $Clr.Dim
    $files = @()
    try {
        $cmd = New-Object System.Data.SqlClient.SqlCommand ("RESTORE FILELISTONLY FROM $fromClause"), $Conn
        $cmd.CommandTimeout = 120
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dt = New-Object System.Data.DataTable
        $null = $adapter.Fill($dt)
        foreach($row in $dt.Rows){
            $files += [PSCustomObject]@{
                LogicalName = "$($row['LogicalName'])"
                Type        = "$($row['Type'])"
            }
        }
    } catch {
        throw "RESTORE FILELISTONLY failed: $($_.Exception.Message)"
    }

    if($files.Count -eq 0){
        throw "RESTORE FILELISTONLY returned no rows. Verify the backup file is accessible from the destination SQL Server: $BackupPath"
    }
    Write-Log "  Found $($files.Count) file(s) in backup." $Clr.Dim

    # Get destination default data and log directories
    $dataDir = ""; $logDir = ""
    try { $dataDir = Query-Scalar $Conn "SELECT CONVERT(NVARCHAR(512), SERVERPROPERTY('InstanceDefaultDataPath'))" } catch {}
    try { $logDir  = Query-Scalar $Conn "SELECT CONVERT(NVARCHAR(512), SERVERPROPERTY('InstanceDefaultLogPath'))"  } catch {}
    if(-not $dataDir){ $dataDir = "C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA" }
    if(-not $logDir) { $logDir  = $dataDir }
    $dataDir = $dataDir.TrimEnd('')
    $logDir  = $logDir.TrimEnd('')

    # Build MOVE clauses - use restoreName for physical filenames to avoid collisions
    $moveClauses = @()
    $fileIdx = 0
    foreach($f in $files){
        $lname = $f.LogicalName
        if($f.Type -eq "L"){
            $moveClauses += "  MOVE N'$lname' TO N'$logDir\${restoreName}_log${fileIdx}.ldf'"
        } else {
            $moveClauses += "  MOVE N'$lname' TO N'$dataDir\${restoreName}_data${fileIdx}.mdf'"
        }
        $fileIdx++
    }
    $moveStr = $moveClauses -join ",`n"

    # Build WITH options
    $opts = @()
    if($WithReplace)      { $opts += "REPLACE" }
    if($KeepReplication)  { $opts += "KEEP_REPLICATION" }
    if($RestrictedUser)   { $opts += "RESTRICTED_USER" }
    $opts += $RecoveryMode
    $opts += "STATS = 10"
    $optStr = $opts -join ", "

    Write-Log "  Restoring [$restoreName] WITH: $optStr" $Clr.Dim

    # RESTORE must run from master - switch context before executing
    try { $Conn.ChangeDatabase("master") } catch { }

    $sql = @"
USE master;
RESTORE DATABASE [$restoreName]
FROM $fromClause
WITH
$moveStr,
$optStr;
"@
    try {
        Invoke-LongSql $Conn $sql "Restore of [$restoreName]"
        Write-Log "  Restore complete: [$restoreName]" $Clr.Green
    } catch {
        if("$($_.Exception.Message)" -like "Cancelled*"){
            Write-Log "  Restore cancelled. [$restoreName] may be left in RESTORING state - re-run the restore or use RESTORE DATABASE [$restoreName] WITH RECOVERY." $Clr.Orange
        }
        # Restore failed - if we set SINGLE_USER, put the database back to MULTI_USER
        # so it isn't left inaccessible
        if($setSingleUser){
            try {
                Write-Log "  Restore failed - resetting [$restoreName] to MULTI_USER..." $Clr.Yellow
                try { $Conn.ChangeDatabase("master") } catch { }
                Exec-Sql $Conn "IF EXISTS (SELECT 1 FROM sys.databases WHERE name = N'$restoreName') ALTER DATABASE [$restoreName] SET MULTI_USER WITH ROLLBACK IMMEDIATE"
                Write-Log "  [$restoreName] reset to MULTI_USER." $Clr.Green
            } catch {
                Write-Log "  Could not reset to MULTI_USER: $($_.Exception.Message)" $Clr.Red
                Write-Log "  Run manually: ALTER DATABASE [$restoreName] SET MULTI_USER WITH ROLLBACK IMMEDIATE" $Clr.Red
            }
        }
        throw
    }

    # Return the actual name used so callers can run stats/compat/orphan fix on the right DB
    return $restoreName
}

# Update statistics on a restored database
function Update-Statistics {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$DbName)
    Write-Log "  Updating statistics on [$DbName]..." $Clr.Dim
    $sql = @"
USE [$DbName];
EXEC sp_updatestats;
"@
    Invoke-LongSql $Conn $sql "Update statistics on [$DbName]"
    Write-Log "  Statistics updated: [$DbName]" $Clr.Green
}

# Set compatibility level to the destination server's own level
function Set-CompatibilityLevel {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$DbName,[int]$Level)
    $current = Query-Scalar $Conn "SELECT compatibility_level FROM sys.databases WHERE name = N'$DbName'"
    if([int]$current -eq $Level){
        Write-Log "  [$DbName] already at compat level $Level - no change needed" $Clr.Dim
        return
    }
    Write-Log "  Setting [$DbName] compatibility level: $current -> $Level" $Clr.Dim
    Exec-Sql $Conn "ALTER DATABASE [$DbName] SET COMPATIBILITY_LEVEL = $Level"
    Write-Log "  Compatibility level set to $Level on [$DbName]" $Clr.Green
}

# Detect and fix orphaned users
function Fix-OrphanedUsers {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$DbName)
    Write-Log "  Checking for orphaned users in [$DbName]..." $Clr.Dim

    # Find orphans: users with no matching login SID in master
    $sql = @"
USE [$DbName];
SELECT dp.name AS UserName, dp.type_desc
FROM sys.database_principals dp
WHERE dp.type IN ('S','U','G')
  AND dp.name NOT IN ('dbo','guest','INFORMATION_SCHEMA','sys')
  AND dp.sid IS NOT NULL
  AND dp.sid <> 0x00
  AND NOT EXISTS (
        SELECT 1 FROM sys.server_principals sp
        WHERE sp.sid = dp.sid )
ORDER BY dp.name;
"@
    $orphans = Query-Table $Conn $sql

    if($orphans.Rows.Count -eq 0){
        Write-Log "  No orphaned users found in [$DbName]" $Clr.Green
        return
    }

    Write-Log "  Found $($orphans.Rows.Count) orphaned user(s) in [$DbName]:" $Clr.Yellow
    foreach($row in $orphans.Rows){
        $uname = $row["UserName"]
        $utype = $row["type_desc"]
        Write-Log "    $uname  ($utype)" $Clr.Yellow

        if($utype -eq "SQL_USER"){
            # Check if a login with the same name exists - if so, re-link by SID
            $loginExists = Query-Scalar $Conn "SELECT COUNT(1) FROM sys.server_principals WHERE name = N'$uname' AND type = 'S'"
            if([int]$loginExists -gt 0){
                try {
                    Exec-Sql $Conn "USE [$DbName]; ALTER USER [$uname] WITH LOGIN = [$uname]"
                    Write-Log "    -> Relinked [$uname] to SQL login" $Clr.Green
                } catch {
                    Write-Log "    -> Relink failed for [$uname]: $($_.Exception.Message)" $Clr.Red
                    # Fall back to sp_change_users_login
                    try {
                        Exec-Sql $Conn "USE [$DbName]; EXEC sp_change_users_login 'UPDATE_ONE', '$uname', '$uname'"
                        Write-Log "    -> Fixed via sp_change_users_login" $Clr.Green
                    } catch {
                        Write-Log "    -> Could not fix [$uname] automatically. Manual action needed." $Clr.Red
                    }
                }
            } else {
                Write-Log "    -> No matching login for [$uname]. Create the login on destination then relink." $Clr.Orange
            }
        } elseif($utype -in @("WINDOWS_USER","WINDOWS_GROUP")){
            # For Windows users just try ALTER USER WITH LOGIN
            try {
                Exec-Sql $Conn "USE [$DbName]; ALTER USER [$uname] WITH LOGIN = [$uname]"
                Write-Log "    -> Relinked Windows user/group [$uname]" $Clr.Green
            } catch {
                Write-Log "    -> Cannot relink [$uname]: $($_.Exception.Message). The Windows account may not exist on this domain." $Clr.Red
            }
        }
    }
}

# ---------------------------------------------------------------
# ALWAYS ON AVAILABILITY GROUP (AG) SUPPORT
# Off unless the user answers Yes to the AG question at Run / Script Out.
# The tool only ever connects to the destination (primary). Scripts are generated for the
# primary replica, the secondary replicas and a rollback; secondaries are never executed against.
# ---------------------------------------------------------------
function ConvertTo-SqlBracket { param([string]$s) return $s.Replace(']',']]') }
function ConvertTo-SqlLiteral { param([string]$s) return $s.Replace("'","''") }

# Destination server name as used for the backup folder (same rule as Get-BackupPath)
function Get-DstFolderName {
    $n = ""
    if($script:DstConn -and $script:DstConn.State -eq 'Open'){
        try { $n = [string](Query-Scalar $script:DstConn "SELECT @@SERVERNAME") } catch { }
    }
    if(-not $n){ $n = $TxtDstSrv.Text.Trim() }
    return $n.Replace('\','_').Trim()
}

# MOVE clauses using the same file naming as Restore-Database (<name>_data<N>.mdf / <name>_log<N>.ldf).
# $Files items need LogicalName and Type (L = log, S = FILESTREAM, anything else = data).
function Get-RestoreMoveClauses {
    param($Files,[string]$RestoreName,[string]$DataDir,[string]$LogDir)
    $d = $DataDir.TrimEnd('\')
    $l = $LogDir.TrimEnd('\')
    $out = @()
    $idx = 0
    foreach($f in @($Files)){
        $ln = ConvertTo-SqlLiteral $f.LogicalName
        if($f.Type -eq "L"){
            $path = "$l\${RestoreName}_log$idx.ldf"
        } elseif($f.Type -eq "S"){
            $path = "$d\${RestoreName}_fs$idx"
        } else {
            $path = "$d\${RestoreName}_data$idx.mdf"
        }
        $out += "MOVE N'$ln' TO N'$(ConvertTo-SqlLiteral $path)'"
        $idx++
    }
    return $out
}

# Lists the availability groups on a server with their replicas (empty if Always On is not enabled).
function Get-AgInventory {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $hadr = 0
    try { $hadr = Query-Scalar $Conn "SELECT CONVERT(INT, ISNULL(SERVERPROPERTY('IsHadrEnabled'), 0))" } catch { $hadr = 0 }
    if([int]$hadr -ne 1){ return }

    $sqlNew = @"
SELECT ag.name AS ag_name, ar.replica_server_name, ar.availability_mode_desc, ar.failover_mode_desc,
       ar.seeding_mode_desc, ISNULL(rs.role_desc, N'UNKNOWN') AS role_desc
FROM sys.availability_groups ag
JOIN sys.availability_replicas ar ON ar.group_id = ag.group_id
LEFT JOIN sys.dm_hadr_availability_replica_states rs ON rs.replica_id = ar.replica_id
ORDER BY ag.name, ar.replica_server_name;
"@
    $dt = $null
    try {
        $dt = Query-Table $Conn $sqlNew
    } catch {
        # seeding_mode_desc does not exist before SQL Server 2016
        $dt = Query-Table $Conn ($sqlNew.Replace("ar.seeding_mode_desc,", "N'N/A' AS seeding_mode_desc,"))
    }
    $list = @()
    foreach($row in $dt.Rows){
        $list += [PSCustomObject]@{
            AgName       = [string]$row["ag_name"]
            Replica      = [string]$row["replica_server_name"]
            Availability = [string]$row["availability_mode_desc"]
            Failover     = [string]$row["failover_mode_desc"]
            Seeding      = [string]$row["seeding_mode_desc"]
            Role         = [string]$row["role_desc"]
        }
    }
    return $list
}

# Dialog: pick (dropdown) or type an AG name, choose the seeding method.
function Show-AgSettingsDialog {
    param([string[]]$AgNames,[bool]$ShowExecute = $true)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Always On Availability Group"
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor       = $Clr.BG
    $dlg.ForeColor       = $Clr.Text
    $dlg.Font            = $FontUI
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock = [System.Windows.Forms.DockStyle]::Top; $hdr.Height = 46; $hdr.BackColor = $Clr.Panel
    $dlg.Controls.Add($hdr)
    $ht = New-Object System.Windows.Forms.Label
    $ht.Text = "  Add Restored Databases to an Availability Group"; $ht.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ht.ForeColor = $Clr.Orange; $ht.Font = $FontUIB
    $ht.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($ht)

    $info = New-Object System.Windows.Forms.Label
    $info.Text = "After each database is restored and recovered on the destination, it is added to the AG. Scripts are written for the primary replica, the secondary replicas and a rollback."
    $info.Location = [System.Drawing.Point]::new(16, 56); $info.Size = [System.Drawing.Size]::new(442, 44)
    $info.ForeColor = $Clr.Dim; $info.Font = $FontSm
    $dlg.Controls.Add($info)

    $lblAg = New-Object System.Windows.Forms.Label
    $lblAg.Text = "Availability group (pick one or type a name):"
    $lblAg.Location = [System.Drawing.Point]::new(16, 106); $lblAg.Size = [System.Drawing.Size]::new(442, 18)
    $lblAg.ForeColor = $Clr.Text; $lblAg.Font = $FontUI
    $dlg.Controls.Add($lblAg)

    $Combo = New-Object System.Windows.Forms.ComboBox
    $Combo.Location = [System.Drawing.Point]::new(16, 128); $Combo.Size = [System.Drawing.Size]::new(442, 26)
    $Combo.BackColor = $Clr.Input; $Combo.ForeColor = $Clr.Text; $Combo.Font = $FontUI
    $Combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    foreach($n in $AgNames){ [void]$Combo.Items.Add($n) }
    if($Combo.Items.Count -gt 0){ $Combo.SelectedIndex = 0 }
    $dlg.Controls.Add($Combo)

    $found = New-Object System.Windows.Forms.Label
    if($AgNames.Count -gt 0){
        $found.Text = "$($AgNames.Count) availability group(s) found on the destination."
        $found.ForeColor = $Clr.Green
    } else {
        $found.Text = "No availability groups were read from the destination - type the AG name. Replica details will not be scripted."
        $found.ForeColor = $Clr.Yellow
    }
    $found.Location = [System.Drawing.Point]::new(16, 158); $found.Size = [System.Drawing.Size]::new(442, 30); $found.Font = $FontSm
    $dlg.Controls.Add($found)

    $lblSeed = New-Object System.Windows.Forms.Label
    $lblSeed.Text = "Seeding method:"
    $lblSeed.Location = [System.Drawing.Point]::new(16, 194); $lblSeed.Size = [System.Drawing.Size]::new(442, 18)
    $lblSeed.ForeColor = $Clr.Text; $lblSeed.Font = $FontUIB
    $dlg.Controls.Add($lblSeed)

    $rbAuto = New-Object System.Windows.Forms.RadioButton
    $rbAuto.Text = "Automatic seeding"; $rbAuto.Checked = $true
    $rbAuto.Location = [System.Drawing.Point]::new(16, 216); $rbAuto.Size = [System.Drawing.Size]::new(442, 20)
    $rbAuto.ForeColor = $Clr.Text; $rbAuto.Font = $FontUI
    $dlg.Controls.Add($rbAuto)
    $dAuto = New-Object System.Windows.Forms.Label
    $dAuto.Text = "SQL Server copies the database to each secondary over the AG endpoint (SQL Server 2016+). Needs SEEDING_MODE = AUTOMATIC on the replicas and GRANT CREATE ANY DATABASE on each secondary."
    $dAuto.Location = [System.Drawing.Point]::new(34, 238); $dAuto.Size = [System.Drawing.Size]::new(424, 44)
    $dAuto.ForeColor = $Clr.Dim; $dAuto.Font = $FontSm
    $dlg.Controls.Add($dAuto)

    $rbManual = New-Object System.Windows.Forms.RadioButton
    $rbManual.Text = "Manual (backup, then restore WITH NORECOVERY on each secondary)"
    $rbManual.Location = [System.Drawing.Point]::new(16, 288); $rbManual.Size = [System.Drawing.Size]::new(442, 20)
    $rbManual.ForeColor = $Clr.Text; $rbManual.Font = $FontUI
    $dlg.Controls.Add($rbManual)
    $dManual = New-Object System.Windows.Forms.Label
    $dManual.Text = "A full and log backup are taken on the primary. The secondary script restores them WITH NORECOVERY and joins the database to the AG."
    $dManual.Location = [System.Drawing.Point]::new(34, 310); $dManual.Size = [System.Drawing.Size]::new(424, 32)
    $dManual.ForeColor = $Clr.Dim; $dManual.Font = $FontSm
    $dlg.Controls.Add($dManual)

    $chkExec = New-Object System.Windows.Forms.CheckBox
    $chkExec.Text = "Run the primary-replica steps now (otherwise they are only scripted)"
    $chkExec.Checked = $true
    $chkExec.Location = [System.Drawing.Point]::new(16, 350); $chkExec.Size = [System.Drawing.Size]::new(442, 22)
    $chkExec.ForeColor = $Clr.Text; $chkExec.Font = $FontUI; $chkExec.BackColor = [System.Drawing.Color]::Transparent
    $chkExec.Visible = $ShowExecute
    $dlg.Controls.Add($chkExec)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Proceed"; $btnOK.Location = [System.Drawing.Point]::new(268, 384); $btnOK.Size = [System.Drawing.Size]::new(96, 28)
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.BackColor = $Clr.Orange; $btnOK.ForeColor = [System.Drawing.Color]::FromArgb(10,10,10); $btnOK.Font = $FontUIB
    $btnOK.FlatAppearance.BorderColor = $Clr.Orange
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOK)
    $dlg.AcceptButton = $btnOK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel run"; $btnCancel.Location = [System.Drawing.Point]::new(372, 384); $btnCancel.Size = [System.Drawing.Size]::new(86, 28)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.BackColor = $Clr.Card; $btnCancel.ForeColor = $Clr.Dim; $btnCancel.Font = $FontSm
    $btnCancel.FlatAppearance.BorderColor = $Clr.Border
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    $dlg.ClientSize = [System.Drawing.Size]::new(474, 428)
    $dlg.Add_Shown({ $dlg.Activate() })
    $r = $dlg.ShowDialog()

    $agName = $Combo.Text.Trim()
    $seed   = if($rbManual.Checked){ "MANUAL" } else { "AUTOMATIC" }
    $exec   = ($ShowExecute -and $chkExec.Checked)
    $dlg.Dispose()

    if($r -ne [System.Windows.Forms.DialogResult]::OK -or -not $agName){ return $null }
    return [PSCustomObject]@{ AgName = $agName; Seeding = $seed; Execute = $exec }
}

# Asks the AG question and returns the plan.
#   $null              = the user cancelled
#   Enabled = $false   = standalone migration (nothing AG-related happens)
#   Enabled = $true    = AgName, Seeding (AUTOMATIC|MANUAL), Execute, Replicas
function Get-AgPlan {
    param([bool]$ForScriptOut = $false)

    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Is this migration for an Always On Availability Group?`n`nYes - add the restored databases to an AG on the destination.`nNo - standalone migration.",
        "Always On Availability Group",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if($ans -ne [System.Windows.Forms.DialogResult]::Yes){
        return [PSCustomObject]@{ Enabled = $false }
    }

    $inv = @()
    if($script:DstConn -and $script:DstConn.State -eq 'Open'){
        try { $inv = @(Get-AgInventory $script:DstConn) }
        catch { Write-Log "  Could not read availability groups from the destination: $($_.Exception.Message)" $Clr.Yellow }
    } else {
        Write-Log "  Destination not connected - type the AG name; replica details will not be scripted." $Clr.Yellow
    }
    $names = @($inv | ForEach-Object { $_.AgName } | Sort-Object -Unique)

    $sel = Show-AgSettingsDialog -AgNames $names -ShowExecute (-not $ForScriptOut)
    if($null -eq $sel){ return $null }

    $reps = @($inv | Where-Object { $_.AgName -eq $sel.AgName })
    return [PSCustomObject]@{
        Enabled  = $true
        AgName   = $sel.AgName
        Seeding  = $sel.Seeding
        Execute  = $sel.Execute
        Replicas = $reps
    }
}

# Read-only capture of the source AG topology (AGs, replicas, listeners, member databases).
function Export-AgTopology {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string[]]$Databases)

    $inv = @(Get-AgInventory $Conn)
    if($inv.Count -eq 0){
        Write-Log "  No availability groups found on the source (or Always On is not enabled). Nothing to capture." $Clr.Dim
        return
    }

    $sqlMembers = @"
SELECT ag.name AS ag_name, adc.database_name
FROM sys.availability_databases_cluster adc
JOIN sys.availability_groups ag ON ag.group_id = adc.group_id
ORDER BY ag.name, adc.database_name;
"@
    $sqlListeners = @"
SELECT ag.name AS ag_name, l.dns_name, l.port
FROM sys.availability_group_listeners l
JOIN sys.availability_groups ag ON ag.group_id = l.group_id
ORDER BY ag.name;
"@
    $members   = Query-Table $Conn $sqlMembers
    $listeners = $null
    try { $listeners = Query-Table $Conn $sqlListeners } catch { }

    $out = New-Object System.Collections.Generic.List[string]
    $out.Add("pnfp Migration - source availability group topology (read-only capture)")
    $out.Add("Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $out.Add("Source    : $($TxtSrcSrv.Text.Trim())")
    $out.Add("")

    foreach($agName in @($inv | ForEach-Object { $_.AgName } | Sort-Object -Unique)){
        $out.Add("Availability group: $agName")
        foreach($r in @($inv | Where-Object { $_.AgName -eq $agName })){
            $out.Add(("  Replica {0}  role={1}  availability={2}  failover={3}  seeding={4}" -f $r.Replica,$r.Role,$r.Availability,$r.Failover,$r.Seeding))
        }
        if($listeners){
            foreach($row in $listeners.Rows){
                if([string]$row["ag_name"] -eq $agName){ $out.Add(("  Listener {0}:{1}" -f [string]$row["dns_name"],[string]$row["port"])) }
            }
        }
        foreach($row in $members.Rows){
            if([string]$row["ag_name"] -eq $agName){
                $dn = [string]$row["database_name"]
                $mark = ""
                if($Databases -contains $dn){ $mark = "   <-- selected for migration" }
                $out.Add("  Database $dn$mark")
                if($mark){ Write-Log "  [$dn] is a member of availability group [$agName] on the source." $Clr.Yellow }
            }
        }
        $out.Add("")
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $path  = Join-Path (Get-JobFileFolder) ("pnfp_AG_Topology_{0}_{1}.txt" -f (Get-SafeFileName $TxtSrcSrv.Text.Trim()), $stamp)
    try {
        $out | Out-File -FilePath $path -Encoding UTF8 -Force
        Write-Log "  Source AG topology saved: $path" $Clr.Green
    } catch {
        Write-Log "  Could not save AG topology: $($_.Exception.Message)" $Clr.Red
    }
}

# Builds the primary, secondary and rollback scripts. Nothing is executed here.
function New-AgScripts {
    param(
        [string[]]$Databases,        # database names as restored on the destination
        [string]$AgName,
        [string]$Seeding,            # AUTOMATIC | MANUAL
        $Replicas = @(),             # rows from Get-AgInventory for this AG (may be empty)
        [string]$SubFolder,
        [string]$Stamp,
        [hashtable]$LogicalFiles = @{}   # restored db name -> files (LogicalName, Type)
    )

    $reps    = @($Replicas)
    $dstName = Get-DstFolderName
    $dstSrv  = $TxtDstSrv.Text.Trim()
    $agQ     = ConvertTo-SqlBracket $AgName
    $agL     = ConvertTo-SqlLiteral $AgName
    $dbList  = $Databases -join ', '
    $inList  = ($Databases | ForEach-Object { "N'" + (ConvertTo-SqlLiteral $_) + "'" }) -join ', '
    $gen     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $files = @{}
    foreach($db in $Databases){
        $folder = ($db -replace '[\\/:*?"<>|]', '_').Trim()
        $dir    = "$($script:BackupShareRoot)\$dstName\$SubFolder\$folder"
        $files[$db] = [PSCustomObject]@{
            Dir  = $dir
            Full = "$dir\${db}_AG_FULL_$Stamp.bak"
            Log  = "$dir\${db}_AG_LOG_$Stamp.trn"
        }
    }

    # ---------------- PRIMARY ----------------
    $p = New-Object System.Text.StringBuilder
    $txt = @"
/*
  pnfp Migration - Always On AG: PRIMARY replica steps
  Generated : $gen
  Run on    : the DESTINATION PRIMARY replica ($dstSrv)
  AG        : $AgName
  Seeding   : $Seeding
  Databases : $dbList

  Prerequisites: the databases are already restored WITH RECOVERY on this server, and the
  logins exist on every replica (use the pnfp_Logins_*.sql file generated by this tool).
*/


"@
    [void]$p.Append($txt)

    if($Seeding -eq "AUTOMATIC"){
        [void]$p.AppendLine("-- Automatic seeding needs SEEDING_MODE = AUTOMATIC on every replica (run on the primary).")
        $needs = @($reps | Where-Object { $_.Seeding -ne 'AUTOMATIC' })
        if($reps.Count -eq 0){
            [void]$p.AppendLine("-- Replica details were not read from the server. Make sure every replica uses SEEDING_MODE = AUTOMATIC:")
            [void]$p.AppendLine("-- ALTER AVAILABILITY GROUP [$agQ] MODIFY REPLICA ON N'<replica server name>' WITH (SEEDING_MODE = AUTOMATIC);")
        } elseif($needs.Count -eq 0){
            [void]$p.AppendLine("-- All replicas already use SEEDING_MODE = AUTOMATIC.")
        } else {
            if(@($needs | Where-Object { $_.Seeding -eq 'N/A' }).Count -gt 0){
                [void]$p.AppendLine("-- WARNING: automatic seeding needs SQL Server 2016 or later on every replica. Use the Manual method on older versions.")
            }
            foreach($r in $needs){
                [void]$p.AppendLine("ALTER AVAILABILITY GROUP [$agQ] MODIFY REPLICA ON N'$(ConvertTo-SqlLiteral $r.Replica)' WITH (SEEDING_MODE = AUTOMATIC);")
            }
            [void]$p.AppendLine("GO")
        }
        [void]$p.AppendLine("")
    }

    foreach($db in $Databases){
        $dbq = ConvertTo-SqlBracket $db
        $dbl = ConvertTo-SqlLiteral $db
        $f   = $files[$db]
        # A log backup is only needed for manual seeding (the secondary restores it)
        $logBk = ""
        if($Seeding -eq "MANUAL"){
            $logBk = "BACKUP LOG [$dbq]`nTO DISK = N'$(ConvertTo-SqlLiteral $f.Log)'`nWITH COMPRESSION, CHECKSUM, INIT, NAME = N'$dbl - AG log backup';`nGO`n"
        }
        $txt = @"
-- ============================================================
-- [$db]
-- Backup folder (must exist, writable by the SQL Server service account): $($f.Dir)
-- ============================================================
ALTER DATABASE [$dbq] SET RECOVERY FULL;
GO
BACKUP DATABASE [$dbq]
TO DISK = N'$(ConvertTo-SqlLiteral $f.Full)'
WITH COMPRESSION, CHECKSUM, INIT, NAME = N'$dbl - AG full backup';
GO
${logBk}ALTER AVAILABILITY GROUP [$agQ] ADD DATABASE [$dbq];
GO


"@
        [void]$p.Append($txt)
    }

    $txt = @"
-- Verify (run on the primary once the secondaries have joined):
SELECT DB_NAME(drs.database_id) AS database_name, ar.replica_server_name,
       drs.synchronization_state_desc, drs.synchronization_health_desc
FROM sys.dm_hadr_database_replica_states drs
JOIN sys.availability_replicas ar ON ar.replica_id = drs.replica_id
WHERE DB_NAME(drs.database_id) IN ($inList)
ORDER BY database_name, ar.replica_server_name;
GO
"@
    [void]$p.AppendLine($txt)

    # ---------------- SECONDARY ----------------
    $s = New-Object System.Text.StringBuilder
    $repNames = "(replicas not read from the server)"
    $secondaries = @($reps | Where-Object { $_.Role -ne 'PRIMARY' })
    if($secondaries.Count -gt 0){ $repNames = ($secondaries | ForEach-Object { $_.Replica }) -join ', ' }
    $txt = @"
/*
  pnfp Migration - Always On AG: SECONDARY replica steps
  Generated : $gen
  Run on    : EACH SECONDARY replica of [$AgName]
  Secondaries: $repNames
  Seeding   : $Seeding
  Databases : $dbList

  Run the pnfp_Logins_*.sql file on every secondary first so the logins exist with the same SIDs.
*/


"@
    [void]$s.Append($txt)

    if($Seeding -eq "AUTOMATIC"){
        $txt = @"
-- Let the AG create the seeded databases on this secondary.
ALTER AVAILABILITY GROUP [$agQ] GRANT CREATE ANY DATABASE;
GO
-- Seeding starts when ADD DATABASE runs on the primary. If it was already run before this grant,
-- seeding retries automatically. Monitor on the primary: SELECT * FROM sys.dm_hadr_automatic_seeding;


"@
        [void]$s.Append($txt)
    } else {
        foreach($db in $Databases){
            $dbq = ConvertTo-SqlBracket $db
            $f   = $files[$db]
            $mv  = @()
            if($LogicalFiles.ContainsKey($db)){
                $mv = @(Get-RestoreMoveClauses -Files $LogicalFiles[$db] -RestoreName $db -DataDir "<secondary_data_dir>" -LogDir "<secondary_log_dir>")
            }
            $withParts = @("NORECOVERY", "STATS = 10") + $mv
            $withStr   = $withParts -join ",`n     "
            $moveNote  = "-- Edit the MOVE paths for this secondary (remove the MOVE clauses if the file paths are identical)."
            if($mv.Count -eq 0){ $moveNote = "-- Add MOVE clauses if the data/log paths on this secondary differ (RESTORE FILELISTONLY lists the logical names)." }
            $txt = @"
-- ============================================================
-- [$db]
$moveNote
-- ============================================================
RESTORE DATABASE [$dbq]
FROM DISK = N'$(ConvertTo-SqlLiteral $f.Full)'
WITH $withStr;
GO
RESTORE LOG [$dbq]
FROM DISK = N'$(ConvertTo-SqlLiteral $f.Log)'
WITH NORECOVERY, STATS = 10;
GO
ALTER DATABASE [$dbq] SET HADR AVAILABILITY GROUP = [$agQ];
GO


"@
            [void]$s.Append($txt)
        }
    }

    $txt = @"
-- ============================================================
-- LOGIN / ORPHANED USER CHECK
-- Orphan fixes made on the primary replicate to the secondaries through the log. What each secondary
-- needs is the login with the same SID. Run this in each database on a readable secondary, or on the
-- new primary after a failover. Any row returned is a user without a matching login on this server.
-- ============================================================
SELECT dp.name AS orphaned_user, dp.type_desc
FROM sys.database_principals dp
WHERE dp.type IN ('S','U','G')
  AND dp.name NOT IN ('dbo','guest','INFORMATION_SCHEMA','sys')
  AND dp.sid IS NOT NULL AND dp.sid <> 0x00
  AND NOT EXISTS (SELECT 1 FROM sys.server_principals sp WHERE sp.sid = dp.sid);
GO
"@
    [void]$s.AppendLine($txt)

    # ---------------- ROLLBACK ----------------
    $rb = New-Object System.Text.StringBuilder
    $txt = @"
/*
  pnfp Migration - Always On AG: ROLLBACK
  Generated : $gen
  Part 1 runs on the PRIMARY replica. Part 2 runs on EACH SECONDARY replica.
  AG        : $AgName
  Databases : $dbList
*/

-- PART 1 (primary): take the databases out of the availability group. They stay online on the primary.

"@
    [void]$rb.Append($txt)
    foreach($db in $Databases){
        [void]$rb.AppendLine("ALTER AVAILABILITY GROUP [$agQ] REMOVE DATABASE [$(ConvertTo-SqlBracket $db)];")
        [void]$rb.AppendLine("GO")
    }
    [void]$rb.AppendLine("")
    [void]$rb.AppendLine("-- PART 2 (each secondary): the copies are left in the RESTORING state. Drop them to clean up.")
    foreach($db in $Databases){
        [void]$rb.AppendLine("-- DROP DATABASE [$(ConvertTo-SqlBracket $db)];")
    }
    [void]$rb.AppendLine("")
    [void]$rb.AppendLine("-- Automatic seeding only: undo the grant on each secondary if it is no longer wanted.")
    [void]$rb.AppendLine("-- ALTER AVAILABILITY GROUP [$agQ] DENY CREATE ANY DATABASE;")

    return [PSCustomObject]@{
        Primary   = $p.ToString()
        Secondary = $s.ToString()
        Rollback  = $rb.ToString()
        Files     = $files
    }
}

# Writes the three AG scripts to the output folder and returns the paths.
function Save-AgScripts {
    param($Scripts,[string]$Stamp)
    $folder = Get-JobFileFolder
    $tag    = Get-SafeFileName (Get-DstFolderName)
    $paths  = @()
    $set = @(
        @("Primary",   $Scripts.Primary),
        @("Secondary", $Scripts.Secondary),
        @("Rollback",  $Scripts.Rollback)
    )
    foreach($item in $set){
        $path = Join-Path $folder ("pnfp_AG_{0}_{1}_{2}.sql" -f $item[0], $tag, $Stamp)
        try {
            $item[1] | Out-File -FilePath $path -Encoding UTF8 -Force
            Write-Log "  AG $($item[0]) script saved: $path" $Clr.Green
            $paths += $path
        } catch {
            Write-Log "  AG $($item[0]) script save failed: $($_.Exception.Message)" $Clr.Red
        }
    }
    return $paths
}

# Run-time AG step on the destination primary: writes the scripts, then (if chosen) runs the
# primary-side steps for each restored database. Secondaries are only scripted.
function Invoke-AgAddOnPrimary {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [string[]]$Databases,
        $AgPlan,
        [string]$SubFolder
    )

    $Databases = @($Databases | Where-Object { $_ })
    if($Databases.Count -eq 0){
        Write-Log "  No restored databases to add to the availability group." $Clr.Yellow
        return
    }
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"

    # Logical file names of the restored databases (used for the secondary restore script)
    $logical = @{}
    foreach($db in $Databases){
        try {
            $dt = Query-Table $Conn ("SELECT name, type FROM sys.master_files WHERE database_id = DB_ID(N'{0}') ORDER BY file_id" -f (ConvertTo-SqlLiteral $db))
            $list = @()
            foreach($row in $dt.Rows){
                $t = "D"
                if([int]$row["type"] -eq 1){ $t = "L" } elseif([int]$row["type"] -eq 2){ $t = "S" }
                $list += [PSCustomObject]@{ LogicalName = [string]$row["name"]; Type = $t }
            }
            if($list.Count -gt 0){ $logical[$db] = $list }
        } catch { }
    }

    $scripts = New-AgScripts -Databases $Databases -AgName $AgPlan.AgName -Seeding $AgPlan.Seeding `
        -Replicas @($AgPlan.Replicas) -SubFolder $SubFolder -Stamp $stamp -LogicalFiles $logical
    $paths = @(Save-AgScripts $scripts $stamp)

    if(-not $AgPlan.Execute){
        Write-Log "  AG steps scripted only (not executed). Run the primary script on the primary replica." $Clr.Dim
        return
    }

    $agL = ConvertTo-SqlLiteral $AgPlan.AgName
    $agQ = ConvertTo-SqlBracket $AgPlan.AgName

    try { $Conn.ChangeDatabase("master") } catch { }
    $role = $null
    try {
        $role = Query-Scalar $Conn ("SELECT rs.role_desc FROM sys.availability_groups ag JOIN sys.dm_hadr_availability_replica_states rs ON rs.group_id = ag.group_id AND rs.is_local = 1 WHERE ag.name = N'{0}'" -f $agL)
    } catch { }
    if(-not $role -or $role -is [System.DBNull]){
        Write-Log "  Availability group [$($AgPlan.AgName)] was not found on the destination. Steps were scripted only." $Clr.Orange
        return
    }
    if([string]$role -ne "PRIMARY"){
        Write-Log "  The destination is not the PRIMARY replica of [$($AgPlan.AgName)] (role: $role). Steps were scripted only - run the primary script on the primary replica." $Clr.Orange
        return
    }

    if($AgPlan.Seeding -eq "AUTOMATIC"){
        foreach($r in @($AgPlan.Replicas | Where-Object { $_.Seeding -ne 'AUTOMATIC' })){
            Write-Log "  WARNING: replica $($r.Replica) is not set to automatic seeding (current: $($r.Seeding)). The AG settings are not changed by this tool - see the primary script." $Clr.Orange
        }
    }

    foreach($db in $Databases){
        if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }
        $dbq = ConvertTo-SqlBracket $db
        $dbl = ConvertTo-SqlLiteral $db
        $f   = $scripts.Files[$db]
        try {
            try { $Conn.ChangeDatabase("master") } catch { }
            $state = Query-Scalar $Conn "SELECT state_desc FROM sys.databases WHERE name = N'$dbl'"
            if([string]$state -ne "ONLINE"){
                Write-Log "  [$db] is not ONLINE on the destination (state: $state) - not added to the AG." $Clr.Yellow
                continue
            }
            $inAg = Query-Scalar $Conn "SELECT COUNT(1) FROM sys.availability_databases_cluster WHERE database_name = N'$dbl'"
            if([int]$inAg -gt 0){
                Write-Log "  [$db] is already in an availability group - skipped." $Clr.Yellow
                continue
            }
            if(-not (Test-Path $f.Dir)){ New-Item -ItemType Directory -Path $f.Dir -Force | Out-Null }

            Write-Log "  [$db] setting recovery model FULL..." $Clr.Dim
            Exec-Sql $Conn "ALTER DATABASE [$dbq] SET RECOVERY FULL"

            $sqlFull = "BACKUP DATABASE [$dbq] TO DISK = N'$(ConvertTo-SqlLiteral $f.Full)' WITH COMPRESSION, CHECKSUM, INIT, NAME = N'$dbl - AG full backup'"
            Invoke-LongSql $Conn $sqlFull "AG full backup of [$db]"
            if($AgPlan.Seeding -eq "MANUAL"){
                $sqlLog = "BACKUP LOG [$dbq] TO DISK = N'$(ConvertTo-SqlLiteral $f.Log)' WITH COMPRESSION, CHECKSUM, INIT, NAME = N'$dbl - AG log backup'"
                Invoke-LongSql $Conn $sqlLog "AG log backup of [$db]"
            }

            try { $Conn.ChangeDatabase("master") } catch { }
            Exec-Sql $Conn "ALTER AVAILABILITY GROUP [$agQ] ADD DATABASE [$dbq]"
            Write-Log "  [$db] added to availability group [$($AgPlan.AgName)]." $Clr.Green
        } catch {
            Write-Log "  AG ADD FAILED for [$db]: $($_.Exception.Message)" $Clr.Red
        }
    }

    $secPath = $paths | Where-Object { $_ -like "*pnfp_AG_Secondary_*" } | Select-Object -First 1
    if($secPath){ Write-Log "  Next: run the SECONDARY script on each secondary replica: $secPath" $Clr.Blue }
    Write-Log "  Make sure the login script has been run on every replica." $Clr.Dim
}

# ---------------------------------------------------------------
# MAIN PIPELINE: PRE-MIGRATION
# ---------------------------------------------------------------
function Run-PreMigration {
    $dbs = Get-SelectedDatabases
    if($dbs.Count -eq 0){
        [System.Windows.Forms.MessageBox]::Show("No databases selected.","Select Databases",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    if(-not $script:SrcConn -or $script:SrcConn.State -ne 'Open'){
        Write-Log "Source is not connected." $Clr.Red; return }
    if(-not $script:DstConn -or $script:DstConn.State -ne 'Open'){
        Write-Log "Destination is not connected." $Clr.Red; return }

    $script:Running = $true
    $script:Cancel  = $false
    $BtnRun.Enabled = $false
    $BtnCancel.Enabled = $true

    try {
        # Always On Availability Group option (asked once for the whole run)
        $agPlan = Get-AgPlan
        if($null -eq $agPlan){ Write-Log "Cancelled." $Clr.Dim; return }
        if($agPlan.Enabled){
            Write-Log "Always On AG: [$($agPlan.AgName)], $($agPlan.Seeding) seeding" $Clr.Blue
            Write-Log "Capturing source availability group topology..." $Clr.Blue
            try { Export-AgTopology $script:SrcConn $dbs } catch { Write-Log "  AG topology capture failed: $($_.Exception.Message)" $Clr.Red }
        } else {
            Write-Log "Always On AG: not used (standalone)." $Clr.Dim
        }

        $dstVersion = Get-SqlVersion $script:DstConn
        Write-Log "Destination SQL version: $dstVersion" $Clr.Dim

        # Ask for target names (once for all dbs - blank = keep original name)
        $renameMap = Show-RenameDialog -Databases $dbs
        if($null -eq $renameMap){ Write-Log "Cancelled." $Clr.Dim; return }

        # Ask the user which compat level to apply (once for the whole run)
        $compatTarget = Show-CompatLevelDialog -DstMajorVersion $dstVersion
        if($null -eq $compatTarget){
            Write-Log "STEP 4 - Compatibility level: skipped by user." $Clr.Dim
        } else {
            Write-Log "STEP 4 - Compatibility level chosen: $compatTarget" $Clr.Blue
        }

        # Ask how many backup files to use if any selected database is large
        if(-not (Get-StripePlan $dbs)){ Write-Log "Cancelled." $Clr.Dim; return }
        $restoredDbs = @()

        # Script out source logins, server roles and server permissions for the destination
        Write-Log "Scripting source logins and permissions..." $Clr.Blue
        try {
            Export-LoginScript $script:SrcConn
        } catch {
            Write-Log "  Login script failed: $($_.Exception.Message)" $Clr.Red
        }

        foreach($db in $dbs){
            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            Write-Sep "PRE-MIGRATION: [$db]"
            Set-Status "Processing [$db]..." $Clr.Yellow

            # 1. Copy-only backup on source
            Write-Log "STEP 1 - Copy-only backup on source" $Clr.Blue
            Test-TdeCertificate $script:SrcConn $script:DstConn $db
            try {
                $backupPath = Get-BackupPath $script:SrcConn $db
                Backup-DatabaseCopyOnly $script:SrcConn $db $backupPath
            } catch {
                Write-Log "  BACKUP FAILED: $($_.Exception.Message)" $Clr.Red
                continue
            }

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            # Resolve target name for this database
            $targetDb = if($renameMap.ContainsKey($db)){ $renameMap[$db] } else { $db }

            # 2. Restore to destination - pre-migration means a FRESH restore.
            # The database should NOT exist on the destination yet.
            # Check first and warn if it does, then restore WITHOUT REPLACE so SQL Server
            # protects against accidentally overwriting an existing production database.
            $restoredName = $db
            $dstDbExists = $false
            try {
                $chk = Query-Scalar $script:DstConn "SELECT COUNT(1) FROM sys.databases WHERE name = N'$targetDb'"
                $dstDbExists = ([int]$chk -gt 0)
            } catch { }

            if($dstDbExists){
                Write-Log "  WARNING: [$targetDb] already exists on the destination." $Clr.Yellow
                Write-Log "  Pre-migration expects a fresh destination. If you want to overwrite, use MIGRATION-DAY mode." $Clr.Yellow
                Write-Log "  Aborting restore of [$targetDb] to protect the existing database." $Clr.Red
                continue
            }

            Write-Log "STEP 2 - Fresh restore to destination as [$targetDb] (WITH RECOVERY, no REPLACE)" $Clr.Blue
            try {
                $restoredName = Restore-Database $script:DstConn $db $backupPath -TargetName $targetDb -WithReplace $false -RecoveryMode "RECOVERY" -CloseConnections $false
            } catch {
                Write-Log "  RESTORE FAILED: $($_.Exception.Message)" $Clr.Red
                continue
            }
            $restoredDbs += [string]$restoredName

            Write-Log "  [$db] restored. Post-restore steps run once all databases are restored." $Clr.Green
        }

        if(-not $script:Cancel){
            # Statistics, compatibility level and orphaned users - after ALL databases are restored
            Write-Sep "POST-RESTORE STEPS (all databases)"
            Invoke-PostRestoreSteps -Databases $restoredDbs -Mode "PRE" -CompatTarget $compatTarget

            Write-Sep "PRE-MIGRATION COMPLETE"
            Write-Log "All selected databases processed." $Clr.Green
            Set-Status "Pre-migration complete." $Clr.Green

            # Add the restored databases to the availability group on the destination
            if($agPlan.Enabled){
                Write-Sep "ALWAYS ON AVAILABILITY GROUP (destination)"
                try {
                    Invoke-AgAddOnPrimary -Conn $script:DstConn -Databases $restoredDbs -AgPlan $agPlan -SubFolder $script:BackupSubFolder
                } catch {
                    Write-Log "  AG step failed: $($_.Exception.Message)" $Clr.Red
                }
            }

            # Offer to disable SQL Agent jobs on the source
            Invoke-AgentJobPrompt -DefaultMode "DISABLE"
        }

    } finally {
        $script:Running    = $false
        $BtnRun.Enabled    = $true
        $BtnCancel.Enabled = $false
    }
}

# ---------------------------------------------------------------
# MAIN PIPELINE: MIGRATION-DAY
# ---------------------------------------------------------------
function Run-PostMigration {
    $dbs = Get-SelectedDatabases
    if($dbs.Count -eq 0){
        [System.Windows.Forms.MessageBox]::Show("No databases selected.","Select Databases",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    if(-not $script:SrcConn -or $script:SrcConn.State -ne 'Open'){
        Write-Log "Source is not connected." $Clr.Red; return }
    if(-not $script:DstConn -or $script:DstConn.State -ne 'Open'){
        Write-Log "Destination is not connected." $Clr.Red; return }

    # Always On Availability Group option (asked once for the whole run)
    $agPlan = Get-AgPlan
    if($null -eq $agPlan){ Write-Log "Cancelled." $Clr.Dim; return }
    if($agPlan.Enabled){
        Write-Log "Always On AG: [$($agPlan.AgName)], $($agPlan.Seeding) seeding" $Clr.Blue
        Write-Log "Capturing source availability group topology..." $Clr.Blue
        try { Export-AgTopology $script:SrcConn $dbs } catch { Write-Log "  AG topology capture failed: $($_.Exception.Message)" $Clr.Red }
    } else {
        Write-Log "Always On AG: not used (standalone)." $Clr.Dim
    }

    # Ask for target names first
    $renameMap = Show-RenameDialog -Databases $dbs
    if($null -eq $renameMap){ Write-Log "Cancelled." $Clr.Dim; return }

    # Get restore options via modal
    $opts = Show-RestoreOptionsDialog
    if(-not $opts){ Write-Log "Cancelled - no restore options chosen." $Clr.Dim; return }

    $script:Running    = $true
    $script:Cancel     = $false
    $BtnRun.Enabled    = $false
    $BtnCancel.Enabled = $true

    try {
        # Ask which compat level to apply - should match source or destination version
        $dstVersion    = Get-SqlVersion $script:DstConn
        Write-Log "Destination SQL version: $dstVersion" $Clr.Dim
        $compatTarget = Show-CompatLevelDialog -DstMajorVersion $dstVersion
        if($null -eq $compatTarget){
            Write-Log "STEP 3 - Compatibility level: skipped by user." $Clr.Dim
        } else {
            Write-Log "STEP 3 - Compatibility level chosen: $compatTarget" $Clr.Blue
        }

        # Ask how many backup files to use if any selected database is large
        if(-not (Get-StripePlan $dbs)){ Write-Log "Cancelled." $Clr.Dim; return }

        # Script out source logins, server roles and server permissions for the destination
        Write-Log "Scripting source logins and permissions..." $Clr.Blue
        try {
            Export-LoginScript $script:SrcConn
        } catch {
            Write-Log "  Login script failed: $($_.Exception.Message)" $Clr.Red
        }

        $restoredDbs = @()

        foreach($db in $dbs){
            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            Write-Sep "MIGRATION-DAY: [$db]"
            Set-Status "Processing [$db]..." $Clr.Yellow

            # 1. Copy-only backup on source
            Write-Log "STEP 1 - Copy-only backup on source" $Clr.Blue
            Test-TdeCertificate $script:SrcConn $script:DstConn $db
            try {
                $backupPath = Get-BackupPath $script:SrcConn $db -SubFolder $script:PostBackupSubFolder
                Backup-DatabaseCopyOnly $script:SrcConn $db $backupPath
            } catch {
                Write-Log "  BACKUP FAILED: $($_.Exception.Message)" $Clr.Red
                continue
            }

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            $targetDb = if($renameMap.ContainsKey($db)){ $renameMap[$db] } else { $db }

            # 2. Restore with chosen options (possibly under a new name)
            Write-Log "STEP 2 - Restore to destination as [$targetDb] (WITH: $($opts.RecoveryMode)$(if($opts.Replace){', REPLACE'})$(if($opts.KeepReplication){', KEEP_REPLICATION'})$(if($opts.RestrictedUser){', RESTRICTED_USER'}))" $Clr.Blue
            $restoredName = $targetDb
            try {
                $restoredName = Restore-Database $script:DstConn $db $backupPath `
                    -TargetName        $targetDb `
                    -WithReplace       $opts.Replace `
                    -RecoveryMode      $opts.RecoveryMode `
                    -KeepReplication   $opts.KeepReplication `
                    -RestrictedUser    $opts.RestrictedUser `
                    -CloseConnections  $opts.CloseConnections
            } catch {
                Write-Log "  RESTORE FAILED: $($_.Exception.Message)" $Clr.Red
                continue
            }
            $restoredDbs += [string]$restoredName

            Write-Log "  [$db] restored. Post-restore steps run once all databases are restored." $Clr.Green
        }

        if(-not $script:Cancel){
            # Compatibility level, statistics and orphaned users - after ALL databases are restored
            Write-Sep "POST-RESTORE STEPS (all databases)"
            if($opts.RecoveryMode -eq "RECOVERY"){
                Invoke-PostRestoreSteps -Databases $restoredDbs -Mode "POST" -CompatTarget $compatTarget
            } else {
                Write-Log "  Restore was $($opts.RecoveryMode): the databases are not recovered yet, so these steps cannot run now." $Clr.Yellow
                Save-DeferredPostRestoreScript -Databases $restoredDbs -Mode "POST" -CompatTarget $compatTarget -RecoveryMode $opts.RecoveryMode
            }

            Write-Sep "PAGE_VERIFY CHECK (destination user databases)"
            try {
                Set-UserDbPageVerifyChecksum $script:DstConn
            } catch {
                Write-Log "  PAGE_VERIFY check failed: $($_.Exception.Message)" $Clr.Red
            }

            Write-Sep "DATABASE OWNER CHECK (restored databases)"
            try {
                Set-UserDbOwnerSa $script:DstConn -RestoredDatabases $restoredDbs
            } catch {
                Write-Log "  Owner check failed: $($_.Exception.Message)" $Clr.Red
            }

            if($agPlan.Enabled){
                Write-Sep "ALWAYS ON AVAILABILITY GROUP (destination)"
                if($opts.RecoveryMode -eq "RECOVERY"){
                    try {
                        Invoke-AgAddOnPrimary -Conn $script:DstConn -Databases $restoredDbs -AgPlan $agPlan -SubFolder $script:PostBackupSubFolder
                    } catch {
                        Write-Log "  AG step failed: $($_.Exception.Message)" $Clr.Red
                    }
                } else {
                    Write-Log "  Restore was $($opts.RecoveryMode) - databases cannot be added to an availability group until they are recovered. AG step skipped." $Clr.Yellow
                }
            }

            Write-Sep "MIGRATION-DAY COMPLETE"
            Write-Log "All selected databases processed." $Clr.Green
            Set-Status "Migration day complete." $Clr.Green

            # Offer to set source databases READ_ONLY
            $roResult = Show-ReadOnlyDialog -Databases $dbs
            if($roResult -and $roResult.Action -ne [System.Windows.Forms.DialogResult]::Cancel){
                Write-Sep $(if($roResult.Mode -eq "READ_WRITE"){ "ROLLBACK SOURCE TO READ_WRITE" } else { "SET SOURCE READ_ONLY" })
                $execute  = ($roResult.Action -eq [System.Windows.Forms.DialogResult]::OK)
                $scriptIt = $true
                Apply-ReadOnly -Databases $roResult.Databases -Execute $execute -ScriptOut $scriptIt -Mode $roResult.Mode
            } else {
                Write-Log "Set source READ_ONLY / READ_WRITE: skipped." $Clr.Dim
            }

            # Offer rollback / disable of SQL Agent jobs on the source
            Invoke-AgentJobPrompt -DefaultMode "ENABLE"

            Write-Log "If the login script was run on the destination after the restores, click Re-check Orphaned Users to relink users." $Clr.Dim
        }

    } finally {
        $script:Running    = $false
        $BtnRun.Enabled    = $true
        $BtnCancel.Enabled = $false
    }
}

function Get-SelectedDatabases {
    $selected = @()
    foreach($item in $DbList.CheckedItems){
        # Strip state annotation like "  [OFFLINE]" that may be appended for non-ONLINE databases
        $name = $item.ToString() -replace '\s+\[.*\]$',''
        $selected += $name.Trim()
    }
    return $selected
}

# ---------------------------------------------------------------
# FORM
# ---------------------------------------------------------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text            = "pnfp Migration Tool"
$Form.Size            = [System.Drawing.Size]::new(1200, 820)
$Form.MinimumSize     = [System.Drawing.Size]::new(900, 600)
$Form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form.BackColor       = $Clr.BG
$Form.ForeColor       = $Clr.Text
$Form.Font            = $FontUI
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

# ---- Header ----
$Hdr = New-Object System.Windows.Forms.Panel
$Hdr.Dock      = [System.Windows.Forms.DockStyle]::Top
$Hdr.Height    = 60
$Hdr.BackColor = $Clr.Panel

$ModeBar = New-Object System.Windows.Forms.Panel
$ModeBar.Location  = [System.Drawing.Point]::new(0,0)
$ModeBar.Size      = [System.Drawing.Size]::new(6,60)
$ModeBar.BackColor = $Clr.Purple
$Hdr.Controls.Add($ModeBar)

$HTitle = New-Object System.Windows.Forms.Label
$HTitle.Text      = "  pnfp Migration Tool"
$HTitle.Location  = [System.Drawing.Point]::new(8,2)
$HTitle.Size      = [System.Drawing.Size]::new(600,36)
$HTitle.ForeColor = $Clr.Purple
$HTitle.Font      = $FontTitle
$HTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$Hdr.Controls.Add($HTitle)

$HModeLbl = New-Object System.Windows.Forms.Label
$HModeLbl.Text      = "  Mode: not selected"
$HModeLbl.Location  = [System.Drawing.Point]::new(8,38)
$HModeLbl.Size      = [System.Drawing.Size]::new(800,20)
$HModeLbl.ForeColor = $Clr.Dim
$HModeLbl.Font      = $FontSm
$Hdr.Controls.Add($HModeLbl)

$HLine = New-Object System.Windows.Forms.Panel
$HLine.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$HLine.Height    = 1
$HLine.BackColor = $Clr.Border
$Hdr.Controls.Add($HLine)

# ---- Status bar ----
$SBar = New-Object System.Windows.Forms.Panel
$SBar.Dock      = [System.Windows.Forms.DockStyle]::Bottom
$SBar.Height    = 28
$SBar.BackColor = $Clr.Panel

$SLbl = New-Object System.Windows.Forms.Label
$SLbl.Dock      = [System.Windows.Forms.DockStyle]::Fill
$SLbl.ForeColor = $Clr.Dim
$SLbl.Font      = $FontSm
$SLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$SLbl.Padding   = [System.Windows.Forms.Padding]::new(10,0,0,0)
$SLbl.Text      = "Select mode to begin."
$SBar.Controls.Add($SLbl)

# ---- Main split ----
$Split = New-Object System.Windows.Forms.SplitContainer
$Split.Dock             = [System.Windows.Forms.DockStyle]::Fill
$Split.Orientation      = [System.Windows.Forms.Orientation]::Vertical
$Split.BackColor        = $Clr.Border
$Split.Panel1.BackColor = $Clr.BG
$Split.Panel2.BackColor = $Clr.BG

# ---- LEFT PANEL ----
$LeftScroll = New-Object System.Windows.Forms.Panel
$LeftScroll.Dock       = [System.Windows.Forms.DockStyle]::Fill
$LeftScroll.AutoScroll = $true
$LeftScroll.BackColor  = $Clr.BG
$Split.Panel1.Controls.Add($LeftScroll)

$Y = 12

# -- Migration Notes --
$NotesSec = New-Object System.Windows.Forms.Panel
$NotesSec.Location  = [System.Drawing.Point]::new(10,$Y)
$NotesSec.Size      = [System.Drawing.Size]::new(430,28)
$NotesSec.BackColor = [System.Drawing.Color]::FromArgb(20,50,100)
$nl = New-Object System.Windows.Forms.Label
$nl.Text      = "  MIGRATION NOTES"
$nl.Dock      = [System.Windows.Forms.DockStyle]::Fill
$nl.ForeColor = $Clr.Blue
$nl.Font      = $FontSec
$nl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$NotesSec.Controls.Add($nl)
$LeftScroll.Controls.Add($NotesSec)
$Y += 34

$NotesHint = New-Object System.Windows.Forms.Label
$NotesHint.Text      = "Record server names, change ticket, cutover window, contacts, constraints..."
$NotesHint.Location  = [System.Drawing.Point]::new(10,$Y)
$NotesHint.Size      = [System.Drawing.Size]::new(430,16)
$NotesHint.ForeColor = $Clr.Dim
$NotesHint.Font      = $FontSm
$NotesHint.BackColor = [System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($NotesHint)
$Y += 18

$TxtNotes = New-Object System.Windows.Forms.TextBox
$TxtNotes.Multiline    = $true
$TxtNotes.Location     = [System.Drawing.Point]::new(10,$Y)
$TxtNotes.Size         = [System.Drawing.Size]::new(420,110)
$TxtNotes.BackColor    = $Clr.Input
$TxtNotes.ForeColor    = $Clr.Text
$TxtNotes.Font         = $FontSm
$TxtNotes.BorderStyle  = [System.Windows.Forms.BorderStyle]::FixedSingle
$TxtNotes.ScrollBars   = [System.Windows.Forms.ScrollBars]::Vertical
$TxtNotes.AcceptsReturn= $true
$LeftScroll.Controls.Add($TxtNotes)
$Y += 116

$BtnSaveNotes = New-Object System.Windows.Forms.Button
$BtnSaveNotes.Text      = "Save Notes"
$BtnSaveNotes.Location  = [System.Drawing.Point]::new(10,$Y)
$BtnSaveNotes.Size      = [System.Drawing.Size]::new(90,24)
$BtnSaveNotes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$BtnSaveNotes.BackColor = $Clr.Card
$BtnSaveNotes.ForeColor = $Clr.Green
$BtnSaveNotes.Font      = $FontSm
$BtnSaveNotes.FlatAppearance.BorderColor = $Clr.Green
$LeftScroll.Controls.Add($BtnSaveNotes)

$NotesStatusLbl = New-Object System.Windows.Forms.Label
$NotesStatusLbl.Location  = [System.Drawing.Point]::new(108,$Y+4)
$NotesStatusLbl.Size      = [System.Drawing.Size]::new(322,16)
$NotesStatusLbl.ForeColor = $Clr.Dim
$NotesStatusLbl.Font      = $FontSm
$NotesStatusLbl.BackColor = [System.Drawing.Color]::Transparent
$NotesStatusLbl.Text      = "Notes auto-saved alongside this script."
$LeftScroll.Controls.Add($NotesStatusLbl)
$Y += 32

# -- Connection Capture --
$ConnCapSec = New-Object System.Windows.Forms.Panel
$ConnCapSec.Location  = [System.Drawing.Point]::new(10,$Y)
$ConnCapSec.Size      = [System.Drawing.Size]::new(430,28)
$ConnCapSec.BackColor = [System.Drawing.Color]::FromArgb(20,50,100)
$ccl = New-Object System.Windows.Forms.Label
$ccl.Text="  0  CONNECTION CAPTURE  (AIDO - Setup Phase)"
$ccl.Dock=[System.Windows.Forms.DockStyle]::Fill; $ccl.ForeColor=$Clr.Blue; $ccl.Font=$FontSec
$ccl.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft; $ConnCapSec.Controls.Add($ccl)
$LeftScroll.Controls.Add($ConnCapSec)
$Y += 34

$BtnConnCapture = New-Object System.Windows.Forms.Button
$BtnConnCapture.Text      = "Run Connection Capture"
$BtnConnCapture.Location  = [System.Drawing.Point]::new(10,$Y)
$BtnConnCapture.Size      = [System.Drawing.Size]::new(180,26)
$BtnConnCapture.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$BtnConnCapture.BackColor = $Clr.Purple
$BtnConnCapture.ForeColor = [System.Drawing.Color]::White
$BtnConnCapture.Font      = $FontUIB
$BtnConnCapture.FlatAppearance.BorderColor = $Clr.Purple
$LeftScroll.Controls.Add($BtnConnCapture)

$ConnCapHint = New-Object System.Windows.Forms.Label
$ConnCapHint.Text      = "Requires AIDO database on source. Saves CSV + XLSX to the PNFP-AUTO_MIGRATION folder on your Desktop."
$ConnCapHint.Location  = [System.Drawing.Point]::new(198,$Y+4)
$ConnCapHint.Size      = [System.Drawing.Size]::new(242,18)
$ConnCapHint.ForeColor = $Clr.Dim
$ConnCapHint.Font      = $FontSm
$ConnCapHint.BackColor = [System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($ConnCapHint)
$Y += 36

$BtnOrphan = New-Object System.Windows.Forms.Button
$BtnOrphan.Text      = "Re-check Orphaned Users"
$BtnOrphan.Location  = [System.Drawing.Point]::new(10,$Y)
$BtnOrphan.Size      = [System.Drawing.Size]::new(180,26)
$BtnOrphan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$BtnOrphan.BackColor = $Clr.Card
$BtnOrphan.ForeColor = $Clr.Purple
$BtnOrphan.Font      = $FontUIB
$BtnOrphan.FlatAppearance.BorderColor = $Clr.Purple
$LeftScroll.Controls.Add($BtnOrphan)

$OrphanHint = New-Object System.Windows.Forms.Label
$OrphanHint.Text      = "Run on the destination after the login script has been run there."
$OrphanHint.Location  = [System.Drawing.Point]::new(198,$Y+4)
$OrphanHint.Size      = [System.Drawing.Size]::new(242,30)
$OrphanHint.ForeColor = $Clr.Dim
$OrphanHint.Font      = $FontSm
$OrphanHint.BackColor = [System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($OrphanHint)
$Y += 38

# -- Mode selector --
$ModeSec = New-Object System.Windows.Forms.Panel
$ModeSec.Location  = [System.Drawing.Point]::new(10,$Y)
$ModeSec.Size      = [System.Drawing.Size]::new(430,28)
$ModeSec.BackColor = [System.Drawing.Color]::FromArgb(20,50,100)
$ml = New-Object System.Windows.Forms.Label
$ml.Text      = "  1  MODE"
$ml.Dock      = [System.Windows.Forms.DockStyle]::Fill
$ml.ForeColor = $Clr.Blue
$ml.Font      = $FontSec
$ml.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$ModeSec.Controls.Add($ml)
$LeftScroll.Controls.Add($ModeSec)
$Y += 34

$BtnChangeMode = New-Object System.Windows.Forms.Button
$BtnChangeMode.Text      = "Select / Change Mode"
$BtnChangeMode.Location  = [System.Drawing.Point]::new(10,$Y)
$BtnChangeMode.Size      = [System.Drawing.Size]::new(180,26)
$BtnChangeMode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$BtnChangeMode.BackColor = $Clr.Card
$BtnChangeMode.ForeColor = $Clr.Blue
$BtnChangeMode.Font      = $FontUIB
$BtnChangeMode.FlatAppearance.BorderColor = $Clr.Blue
$LeftScroll.Controls.Add($BtnChangeMode)
$Y += 36

# -- Source connection --
$SrcSec = New-Object System.Windows.Forms.Panel
$SrcSec.Location  = [System.Drawing.Point]::new(10,$Y)
$SrcSec.Size      = [System.Drawing.Size]::new(430,28)
$SrcSec.BackColor = [System.Drawing.Color]::FromArgb(20,50,100)
$sl = New-Object System.Windows.Forms.Label
$sl.Text      = "  2  SOURCE SERVER"
$sl.Dock      = [System.Windows.Forms.DockStyle]::Fill
$sl.ForeColor = $Clr.Blue
$sl.Font      = $FontSec
$sl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$SrcSec.Controls.Add($sl)
$LeftScroll.Controls.Add($SrcSec)
$Y += 34

$LblSrcSrv = New-Object System.Windows.Forms.Label
$LblSrcSrv.Text      = "Server:"
$LblSrcSrv.Location  = [System.Drawing.Point]::new(10,$Y+3)
$LblSrcSrv.Size      = [System.Drawing.Size]::new(50,18)
$LblSrcSrv.ForeColor = $Clr.Dim; $LblSrcSrv.Font=$FontSm; $LblSrcSrv.BackColor=[System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($LblSrcSrv)

$TxtSrcSrv = New-Object System.Windows.Forms.TextBox
$TxtSrcSrv.Text        = "localhost"
$TxtSrcSrv.Location    = [System.Drawing.Point]::new(62,$Y)
$TxtSrcSrv.Size        = [System.Drawing.Size]::new(200,24)
$TxtSrcSrv.BackColor   = $Clr.Input; $TxtSrcSrv.ForeColor=$Clr.Text; $TxtSrcSrv.Font=$FontUI
$TxtSrcSrv.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$LeftScroll.Controls.Add($TxtSrcSrv)

$ChkSrcSQL = New-Object System.Windows.Forms.CheckBox
$ChkSrcSQL.Text      = "SQL Auth"
$ChkSrcSQL.Location  = [System.Drawing.Point]::new(272,$Y+2)
$ChkSrcSQL.Size      = [System.Drawing.Size]::new(80,20)
$ChkSrcSQL.ForeColor = $Clr.Text; $ChkSrcSQL.Font=$FontSm; $ChkSrcSQL.BackColor=[System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($ChkSrcSQL)
$Y += 28

$LblSrcUser = New-Object System.Windows.Forms.Label
$LblSrcUser.Text="Username:"; $LblSrcUser.Location=[System.Drawing.Point]::new(10,$Y+3)
$LblSrcUser.Size=[System.Drawing.Size]::new(50,18); $LblSrcUser.ForeColor=$Clr.Dim
$LblSrcUser.Font=$FontSm; $LblSrcUser.BackColor=[System.Drawing.Color]::Transparent; $LblSrcUser.Visible=$false
$LeftScroll.Controls.Add($LblSrcUser)
$TxtSrcUser = New-Object System.Windows.Forms.TextBox
$TxtSrcUser.Location=[System.Drawing.Point]::new(62,$Y); $TxtSrcUser.Size=[System.Drawing.Size]::new(140,24)
$TxtSrcUser.BackColor=$Clr.Input; $TxtSrcUser.ForeColor=$Clr.Text; $TxtSrcUser.Font=$FontUI
$TxtSrcUser.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle; $TxtSrcUser.Visible=$false
$LeftScroll.Controls.Add($TxtSrcUser)
$LblSrcPass = New-Object System.Windows.Forms.Label
$LblSrcPass.Text="Password:"; $LblSrcPass.Location=[System.Drawing.Point]::new(210,$Y+3)
$LblSrcPass.Size=[System.Drawing.Size]::new(55,18); $LblSrcPass.ForeColor=$Clr.Dim
$LblSrcPass.Font=$FontSm; $LblSrcPass.BackColor=[System.Drawing.Color]::Transparent; $LblSrcPass.Visible=$false
$LeftScroll.Controls.Add($LblSrcPass)
$TxtSrcPass = New-Object System.Windows.Forms.TextBox
$TxtSrcPass.Location=[System.Drawing.Point]::new(267,$Y); $TxtSrcPass.Size=[System.Drawing.Size]::new(140,24)
$TxtSrcPass.BackColor=$Clr.Input; $TxtSrcPass.ForeColor=$Clr.Text; $TxtSrcPass.Font=$FontUI
$TxtSrcPass.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$TxtSrcPass.UseSystemPasswordChar=$true; $TxtSrcPass.Visible=$false
$LeftScroll.Controls.Add($TxtSrcPass)
$Y += 28

$BtnConnSrc = New-Object System.Windows.Forms.Button
$BtnConnSrc.Text="Connect Source"; $BtnConnSrc.Location=[System.Drawing.Point]::new(10,$Y)
$BtnConnSrc.Size=[System.Drawing.Size]::new(130,26); $BtnConnSrc.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnConnSrc.BackColor=$Clr.Green; $BtnConnSrc.ForeColor=[System.Drawing.Color]::FromArgb(10,10,10)
$BtnConnSrc.Font=$FontUIB; $BtnConnSrc.FlatAppearance.BorderColor=$Clr.Green
$LeftScroll.Controls.Add($BtnConnSrc)

$LblSrcStatus = New-Object System.Windows.Forms.Label
$LblSrcStatus.Text="Not connected"; $LblSrcStatus.Location=[System.Drawing.Point]::new(148,$Y+4)
$LblSrcStatus.Size=[System.Drawing.Size]::new(280,18); $LblSrcStatus.ForeColor=$Clr.Dim; $LblSrcStatus.Font=$FontSm
$LblSrcStatus.BackColor=[System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($LblSrcStatus)
$Y += 36

# -- Destination connection --
$DstSec = New-Object System.Windows.Forms.Panel
$DstSec.Location=[System.Drawing.Point]::new(10,$Y); $DstSec.Size=[System.Drawing.Size]::new(430,28)
$DstSec.BackColor=[System.Drawing.Color]::FromArgb(20,50,100)
$dl = New-Object System.Windows.Forms.Label; $dl.Text="  3  DESTINATION SERVER"
$dl.Dock=[System.Windows.Forms.DockStyle]::Fill; $dl.ForeColor=$Clr.Blue; $dl.Font=$FontSec
$dl.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft; $DstSec.Controls.Add($dl)
$LeftScroll.Controls.Add($DstSec)
$Y += 34

$LblDstSrv = New-Object System.Windows.Forms.Label
$LblDstSrv.Text="Server:"; $LblDstSrv.Location=[System.Drawing.Point]::new(10,$Y+3)
$LblDstSrv.Size=[System.Drawing.Size]::new(50,18); $LblDstSrv.ForeColor=$Clr.Dim
$LblDstSrv.Font=$FontSm; $LblDstSrv.BackColor=[System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($LblDstSrv)
$TxtDstSrv = New-Object System.Windows.Forms.TextBox
$TxtDstSrv.Text="localhost\DEST"; $TxtDstSrv.Location=[System.Drawing.Point]::new(62,$Y)
$TxtDstSrv.Size=[System.Drawing.Size]::new(200,24); $TxtDstSrv.BackColor=$Clr.Input
$TxtDstSrv.ForeColor=$Clr.Text; $TxtDstSrv.Font=$FontUI
$TxtDstSrv.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$LeftScroll.Controls.Add($TxtDstSrv)
$ChkDstSQL = New-Object System.Windows.Forms.CheckBox
$ChkDstSQL.Text="SQL Auth"; $ChkDstSQL.Location=[System.Drawing.Point]::new(272,$Y+2)
$ChkDstSQL.Size=[System.Drawing.Size]::new(80,20); $ChkDstSQL.ForeColor=$Clr.Text
$ChkDstSQL.Font=$FontSm; $ChkDstSQL.BackColor=[System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($ChkDstSQL)
$Y += 28

$LblDstUser = New-Object System.Windows.Forms.Label
$LblDstUser.Text="Username:"; $LblDstUser.Location=[System.Drawing.Point]::new(10,$Y+3)
$LblDstUser.Size=[System.Drawing.Size]::new(50,18); $LblDstUser.ForeColor=$Clr.Dim
$LblDstUser.Font=$FontSm; $LblDstUser.BackColor=[System.Drawing.Color]::Transparent; $LblDstUser.Visible=$false
$LeftScroll.Controls.Add($LblDstUser)
$TxtDstUser = New-Object System.Windows.Forms.TextBox
$TxtDstUser.Location=[System.Drawing.Point]::new(62,$Y); $TxtDstUser.Size=[System.Drawing.Size]::new(140,24)
$TxtDstUser.BackColor=$Clr.Input; $TxtDstUser.ForeColor=$Clr.Text; $TxtDstUser.Font=$FontUI
$TxtDstUser.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle; $TxtDstUser.Visible=$false
$LeftScroll.Controls.Add($TxtDstUser)
$LblDstPass = New-Object System.Windows.Forms.Label
$LblDstPass.Text="Password:"; $LblDstPass.Location=[System.Drawing.Point]::new(210,$Y+3)
$LblDstPass.Size=[System.Drawing.Size]::new(55,18); $LblDstPass.ForeColor=$Clr.Dim
$LblDstPass.Font=$FontSm; $LblDstPass.BackColor=[System.Drawing.Color]::Transparent; $LblDstPass.Visible=$false
$LeftScroll.Controls.Add($LblDstPass)
$TxtDstPass = New-Object System.Windows.Forms.TextBox
$TxtDstPass.Location=[System.Drawing.Point]::new(267,$Y); $TxtDstPass.Size=[System.Drawing.Size]::new(140,24)
$TxtDstPass.BackColor=$Clr.Input; $TxtDstPass.ForeColor=$Clr.Text; $TxtDstPass.Font=$FontUI
$TxtDstPass.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$TxtDstPass.UseSystemPasswordChar=$true; $TxtDstPass.Visible=$false
$LeftScroll.Controls.Add($TxtDstPass)
$Y += 28

$BtnConnDst = New-Object System.Windows.Forms.Button
$BtnConnDst.Text="Connect Destination"; $BtnConnDst.Location=[System.Drawing.Point]::new(10,$Y)
$BtnConnDst.Size=[System.Drawing.Size]::new(150,26); $BtnConnDst.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnConnDst.BackColor=$Clr.Green; $BtnConnDst.ForeColor=[System.Drawing.Color]::FromArgb(10,10,10)
$BtnConnDst.Font=$FontUIB; $BtnConnDst.FlatAppearance.BorderColor=$Clr.Green
$LeftScroll.Controls.Add($BtnConnDst)

$LblDstStatus = New-Object System.Windows.Forms.Label
$LblDstStatus.Text="Not connected"; $LblDstStatus.Location=[System.Drawing.Point]::new(168,$Y+4)
$LblDstStatus.Size=[System.Drawing.Size]::new(260,18); $LblDstStatus.ForeColor=$Clr.Dim; $LblDstStatus.Font=$FontSm
$LblDstStatus.BackColor=[System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($LblDstStatus)
$Y += 36

# -- Database selector --
$DbSec = New-Object System.Windows.Forms.Panel
$DbSec.Location=[System.Drawing.Point]::new(10,$Y); $DbSec.Size=[System.Drawing.Size]::new(430,28)
$DbSec.BackColor=[System.Drawing.Color]::FromArgb(20,50,100)
$dbl = New-Object System.Windows.Forms.Label; $dbl.Text="  4  SELECT DATABASES"
$dbl.Dock=[System.Windows.Forms.DockStyle]::Fill; $dbl.ForeColor=$Clr.Blue; $dbl.Font=$FontSec
$dbl.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft; $DbSec.Controls.Add($dbl)
$LeftScroll.Controls.Add($DbSec)
$Y += 34

$BtnLoadDBs = New-Object System.Windows.Forms.Button
$BtnLoadDBs.Text="Load Databases from Source"; $BtnLoadDBs.Location=[System.Drawing.Point]::new(10,$Y)
$BtnLoadDBs.Size=[System.Drawing.Size]::new(200,26); $BtnLoadDBs.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnLoadDBs.BackColor=$Clr.Card; $BtnLoadDBs.ForeColor=$Clr.Blue; $BtnLoadDBs.Font=$FontUIB
$BtnLoadDBs.FlatAppearance.BorderColor=$Clr.Blue
$LeftScroll.Controls.Add($BtnLoadDBs)

$BtnChkAll = New-Object System.Windows.Forms.Button
$BtnChkAll.Text="All"; $BtnChkAll.Location=[System.Drawing.Point]::new(218,$Y)
$BtnChkAll.Size=[System.Drawing.Size]::new(44,26); $BtnChkAll.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnChkAll.BackColor=$Clr.Card; $BtnChkAll.ForeColor=$Clr.Text; $BtnChkAll.Font=$FontSm
$BtnChkAll.FlatAppearance.BorderColor=$Clr.Border
$LeftScroll.Controls.Add($BtnChkAll)

$BtnChkNone = New-Object System.Windows.Forms.Button
$BtnChkNone.Text="None"; $BtnChkNone.Location=[System.Drawing.Point]::new(266,$Y)
$BtnChkNone.Size=[System.Drawing.Size]::new(50,26); $BtnChkNone.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnChkNone.BackColor=$Clr.Card; $BtnChkNone.ForeColor=$Clr.Text; $BtnChkNone.Font=$FontSm
$BtnChkNone.FlatAppearance.BorderColor=$Clr.Border
$LeftScroll.Controls.Add($BtnChkNone)
$Y += 32

$DbList = New-Object System.Windows.Forms.CheckedListBox
$DbList.Location  = [System.Drawing.Point]::new(10,$Y)
$DbList.Size      = [System.Drawing.Size]::new(420,160)
$DbList.BackColor = $Clr.Input
$DbList.ForeColor = $Clr.Text
$DbList.Font      = $FontMono
$DbList.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$DbList.CheckOnClick  = $true
$LeftScroll.Controls.Add($DbList)
$Y += 168

# -- Backup folder --
$BkSec = New-Object System.Windows.Forms.Panel
$BkSec.Location=[System.Drawing.Point]::new(10,$Y); $BkSec.Size=[System.Drawing.Size]::new(430,28)
$BkSec.BackColor=[System.Drawing.Color]::FromArgb(20,50,100)
$bkl = New-Object System.Windows.Forms.Label; $bkl.Text="  5  BACKUP FOLDER (fixed)"
$bkl.Dock=[System.Windows.Forms.DockStyle]::Fill; $bkl.ForeColor=$Clr.Blue; $bkl.Font=$FontSec
$bkl.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft; $BkSec.Controls.Add($bkl)
$LeftScroll.Controls.Add($BkSec)
$Y += 34

$TxtBackupDir = New-Object System.Windows.Forms.TextBox
$TxtBackupDir.Text="$($script:BackupShareRoot)\<destination server>\$($script:BackupSubFolder)\<database>"; $TxtBackupDir.ReadOnly=$true
$TxtBackupDir.Location=[System.Drawing.Point]::new(10,$Y); $TxtBackupDir.Size=[System.Drawing.Size]::new(310,24)
$TxtBackupDir.BackColor=$Clr.Input; $TxtBackupDir.ForeColor=$Clr.Dim; $TxtBackupDir.Font=$FontSm
$TxtBackupDir.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$LeftScroll.Controls.Add($TxtBackupDir)

$BtnBrowse = New-Object System.Windows.Forms.Button
$BtnBrowse.Text="Browse"; $BtnBrowse.Location=[System.Drawing.Point]::new(328,$Y)
$BtnBrowse.Size=[System.Drawing.Size]::new(70,24); $BtnBrowse.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnBrowse.BackColor=$Clr.Card; $BtnBrowse.ForeColor=$Clr.Text; $BtnBrowse.Font=$FontSm
$BtnBrowse.FlatAppearance.BorderColor=$Clr.Border; $BtnBrowse.Enabled=$false
$LeftScroll.Controls.Add($BtnBrowse)
$Y += 34

# Note about backup file sharing
$BkNote = New-Object System.Windows.Forms.Label
$BkNote.Text = "Note: The backup path must be accessible from BOTH source (to write) and destination (to read). Use a shared network path if source and destination are different servers."
$BkNote.Location = [System.Drawing.Point]::new(10,$Y)
$BkNote.Size     = [System.Drawing.Size]::new(420,42)
$BkNote.ForeColor= $Clr.Yellow
$BkNote.Font     = $FontSm
$BkNote.BackColor= [System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($BkNote)
$Y += 50

# -- Run buttons --
$BtnSec = New-Object System.Windows.Forms.Panel
$BtnSec.Location=[System.Drawing.Point]::new(10,$Y); $BtnSec.Size=[System.Drawing.Size]::new(430,28)
$BtnSec.BackColor=[System.Drawing.Color]::FromArgb(20,50,100)
$rl = New-Object System.Windows.Forms.Label; $rl.Text="  6  RUN"
$rl.Dock=[System.Windows.Forms.DockStyle]::Fill; $rl.ForeColor=$Clr.Blue; $rl.Font=$FontSec
$rl.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft; $BtnSec.Controls.Add($rl)
$LeftScroll.Controls.Add($BtnSec)
$Y += 34

$BtnRun = New-Object System.Windows.Forms.Button
$BtnRun.Text="Run Migration"; $BtnRun.Location=[System.Drawing.Point]::new(10,$Y)
$BtnRun.Size=[System.Drawing.Size]::new(160,36); $BtnRun.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnRun.BackColor=$Clr.Green; $BtnRun.ForeColor=[System.Drawing.Color]::FromArgb(10,10,10)
$BtnRun.Font=$FontUIB; $BtnRun.FlatAppearance.BorderColor=$Clr.Green
$LeftScroll.Controls.Add($BtnRun)

$BtnCancel = New-Object System.Windows.Forms.Button
$BtnCancel.Text="Cancel"; $BtnCancel.Location=[System.Drawing.Point]::new(178,$Y)
$BtnCancel.Size=[System.Drawing.Size]::new(80,36); $BtnCancel.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnCancel.BackColor=$Clr.Red; $BtnCancel.ForeColor=[System.Drawing.Color]::White
$BtnCancel.Font=$FontUIB; $BtnCancel.FlatAppearance.BorderColor=$Clr.Red; $BtnCancel.Enabled=$false
$LeftScroll.Controls.Add($BtnCancel)

$BtnClearLog = New-Object System.Windows.Forms.Button
$BtnClearLog.Text="Clear Log"; $BtnClearLog.Location=[System.Drawing.Point]::new(266,$Y+5)
$BtnClearLog.Size=[System.Drawing.Size]::new(80,26); $BtnClearLog.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnClearLog.BackColor=$Clr.Card; $BtnClearLog.ForeColor=$Clr.Dim; $BtnClearLog.Font=$FontSm
$BtnClearLog.FlatAppearance.BorderColor=$Clr.Border
$LeftScroll.Controls.Add($BtnClearLog)

$BtnScriptOut = New-Object System.Windows.Forms.Button
$BtnScriptOut.Text      = "Script Out"
$BtnScriptOut.Location  = [System.Drawing.Point]::new(354,$Y+5)
$BtnScriptOut.Size      = [System.Drawing.Size]::new(80,26)
$BtnScriptOut.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$BtnScriptOut.BackColor = $Clr.Card
$BtnScriptOut.ForeColor = $Clr.Purple
$BtnScriptOut.Font      = $FontSm
$BtnScriptOut.FlatAppearance.BorderColor = $Clr.Purple
$LeftScroll.Controls.Add($BtnScriptOut)

# ---- RIGHT PANEL: Log ----
$LogLbl = New-Object System.Windows.Forms.Label
$LogLbl.Text      = "  OPERATION LOG"
$LogLbl.Dock      = [System.Windows.Forms.DockStyle]::Top
$LogLbl.Height    = 24
$LogLbl.BackColor = [System.Drawing.Color]::FromArgb(20,50,100)
$LogLbl.ForeColor = $Clr.Blue
$LogLbl.Font      = $FontSec
$LogLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$RtbLog = New-Object System.Windows.Forms.RichTextBox
$RtbLog.Dock        = [System.Windows.Forms.DockStyle]::Fill
$RtbLog.BackColor   = $Clr.Input
$RtbLog.ForeColor   = $Clr.Text
$RtbLog.Font        = $FontMono
$RtbLog.ReadOnly    = $true
$RtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$RtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$RtbLog.WordWrap    = $false

# Fill first, then Top-docked label
$Split.Panel2.Controls.Add($RtbLog)
$Split.Panel2.Controls.Add($LogLbl)

# ---- Assemble ----
$Form.Controls.Add($Split)
$Form.Controls.Add($SBar)
$Form.Controls.Add($Hdr)

# ---------------------------------------------------------------
# EVENT HANDLERS
# ---------------------------------------------------------------

# SQL auth toggle - source
$ChkSrcSQL.Add_CheckedChanged({
    $v = $ChkSrcSQL.Checked
    $LblSrcUser.Visible=$v; $TxtSrcUser.Visible=$v
    $LblSrcPass.Visible=$v; $TxtSrcPass.Visible=$v
})

# SQL auth toggle - destination
$ChkDstSQL.Add_CheckedChanged({
    $v = $ChkDstSQL.Checked
    $LblDstUser.Visible=$v; $TxtDstUser.Visible=$v
    $LblDstPass.Visible=$v; $TxtDstPass.Visible=$v
})

# Connect source
$BtnConnSrc.Add_Click({
    try {
        $cs = Build-CS $TxtSrcSrv.Text.Trim() $ChkSrcSQL.Checked $TxtSrcUser.Text.Trim() $TxtSrcPass.Text
        $cn = Open-Conn $cs
        if($script:SrcConn -and $script:SrcConn.State -eq 'Open'){ try{$script:SrcConn.Close()}catch{} }
        $script:SrcConn = $cn
        $script:SrcCS   = $cs
        $LblSrcStatus.Text      = "Connected: $($TxtSrcSrv.Text.Trim())"
        $LblSrcStatus.ForeColor = $Clr.Green
        Write-Log "Source connected: $($TxtSrcSrv.Text.Trim())" $Clr.Green
    } catch {
        $LblSrcStatus.Text      = "Failed: $($_.Exception.Message.Split([char]13)[0])"
        $LblSrcStatus.ForeColor = $Clr.Red
        Write-Log "Source connection failed: $($_.Exception.Message)" $Clr.Red
    }
})

# Connect destination
$BtnConnDst.Add_Click({
    try {
        $cs = Build-CS $TxtDstSrv.Text.Trim() $ChkDstSQL.Checked $TxtDstUser.Text.Trim() $TxtDstPass.Text -DB "master"
        $cn = Open-Conn $cs
        if($script:DstConn -and $script:DstConn.State -eq 'Open'){ try{$script:DstConn.Close()}catch{} }
        $script:DstConn = $cn
        $script:DstCS   = $cs
        $LblDstStatus.Text      = "Connected: $($TxtDstSrv.Text.Trim())"
        $LblDstStatus.ForeColor = $Clr.Green
        Write-Log "Destination connected: $($TxtDstSrv.Text.Trim())" $Clr.Green
    } catch {
        $LblDstStatus.Text      = "Failed: $($_.Exception.Message.Split([char]13)[0])"
        $LblDstStatus.ForeColor = $Clr.Red
        Write-Log "Destination connection failed: $($_.Exception.Message)" $Clr.Red
    }
})

# Load databases from source
$BtnLoadDBs.Add_Click({
    if(-not $script:SrcCS){
        Write-Log "Connect to source first." $Clr.Red; return }
    try {
        # Re-open a fresh connection each time - avoids stale connection state issues
        $freshConn = New-Object System.Data.SqlClient.SqlConnection $script:SrcCS
        $freshConn.Open()
        # Include all databases that are not system databases, regardless of state
        # (user may want to see RESTORING, OFFLINE etc. to be aware of them)
        $sql = "SELECT name, state_desc FROM sys.databases WHERE database_id > 4 ORDER BY name"
        $dt  = New-Object System.Data.DataTable
        $cmd = New-Object System.Data.SqlClient.SqlCommand $sql,$freshConn
        $cmd.CommandTimeout = 30
        $r   = $cmd.ExecuteReader()
        $dt.Load($r)
        if(-not $r.IsClosed){ $r.Close() }
        $freshConn.Close()

        $DbList.Items.Clear()
        $systemNames = @('master','tempdb','model','msdb')
        foreach($row in $dt.Rows){
            $dbName = $row["name"]
            $dbState = $row["state_desc"]
            if($systemNames -contains $dbName){ continue }
            $displayName = if($dbState -ne "ONLINE"){ "$dbName  [$dbState]" } else { $dbName }
            $DbList.Items.Add($displayName) | Out-Null
        }
        Write-Log "Loaded $($DbList.Items.Count) user database(s) from source." $Clr.Green
    } catch {
        Write-Log "Failed to load databases: $($_.Exception.Message)" $Clr.Red
    }
})

# Check all / none
$BtnChkAll.Add_Click({
    for($i=0;$i -lt $DbList.Items.Count;$i++){ $DbList.SetItemChecked($i,$true) }
})
$BtnChkNone.Add_Click({
    for($i=0;$i -lt $DbList.Items.Count;$i++){ $DbList.SetItemChecked($i,$false) }
})

# Browse backup folder
$BtnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select backup folder (must be reachable from both servers)"
    if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $TxtBackupDir.Text      = $dlg.SelectedPath
        $TxtBackupDir.ForeColor = $Clr.Text
    }
})

# Mode selector
$BtnChangeMode.Add_Click({
    $m = Show-ModeDialog
    if($m){
        $script:Mode = $m
        $TxtBackupDir.Text = "$($script:BackupShareRoot)\<destination server>\$(if($m -eq 'POST'){ $script:PostBackupSubFolder } else { $script:BackupSubFolder })\<database>"
        if($m -eq "PRE"){
            $HModeLbl.Text      = "  Mode: PRE-MIGRATION  |  Backup -> Restore -> Update Stats -> Set Compat Level -> Fix Orphans"
            $ModeBar.BackColor  = $Clr.Green
            $BtnRun.BackColor   = $Clr.Green
            $BtnRun.FlatAppearance.BorderColor = $Clr.Green
            $BtnRun.Text        = "Run Pre-Migration"
        } else {
            $HModeLbl.Text      = "  Mode: MIGRATION-DAY  |  Backup -> Restore (WITH REPLACE / WITH RECOVERY options)"
            $ModeBar.BackColor  = $Clr.Orange
            $BtnRun.BackColor   = $Clr.Orange
            $BtnRun.FlatAppearance.BorderColor = $Clr.Orange
            $BtnRun.Text        = "Run Migration Day"
        }
        Write-Log "Mode set to: $m" $Clr.Blue
    }
})

# Run
$BtnRun.Add_Click({
    if($script:Running){ return }
    if(-not $script:Mode){
        [System.Windows.Forms.MessageBox]::Show("Choose a mode first (Pre or Post) using the 'Select / Change Mode' button.",
            "No Mode Selected",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)|Out-Null
        return
    }
    if($script:Mode -eq "PRE"){ Run-PreMigration }
    else { Run-PostMigration }
})

# Cancel
$BtnCancel.Add_Click({ $script:Cancel = $true; Set-Status "Cancelling - stopping the running operation..." $Clr.Orange })

# Clear log
$BtnClearLog.Add_Click({ $RtbLog.Clear() })
$BtnScriptOut.Add_Click({ Invoke-ScriptOut })
$BtnConnCapture.Add_Click({ Invoke-ConnectionCapture })
$BtnOrphan.Add_Click({ Invoke-OrphanRecheck })

$BtnSaveNotes.Add_Click({
    Save-Notes
    $NotesStatusLbl.Text      = "Saved $(Get-Date -Format 'HH:mm:ss')"
    $NotesStatusLbl.ForeColor = $Clr.Green
})
$TxtNotes.Add_TextChanged({
    $NotesStatusLbl.Text      = "Unsaved changes"
    $NotesStatusLbl.ForeColor = $Clr.Yellow
})

# ---------------------------------------------------------------
# LAUNCH
# ---------------------------------------------------------------
$script:CloseAction = "exit"

$Form.Add_Shown({
    try {
        # Show mode picker immediately on startup
        $m = Show-ModeDialog
        if($m){
            $script:Mode = $m
            $TxtBackupDir.Text = "$($script:BackupShareRoot)\<destination server>\$(if($m -eq 'POST'){ $script:PostBackupSubFolder } else { $script:BackupSubFolder })\<database>"
            if($m -eq "PRE"){
                $HModeLbl.Text      = "  Mode: PRE-MIGRATION  |  Backup -> Restore -> Update Stats -> Set Compat Level -> Fix Orphans"
                $ModeBar.BackColor  = $Clr.Green
                $BtnRun.BackColor   = $Clr.Green
                $BtnRun.FlatAppearance.BorderColor = $Clr.Green
                $BtnRun.Text        = "Run Pre-Migration"
            } else {
                $HModeLbl.Text      = "  Mode: MIGRATION-DAY  |  Backup -> Restore (WITH REPLACE / WITH RECOVERY options)"
                $ModeBar.BackColor  = $Clr.Orange
                $BtnRun.BackColor   = $Clr.Orange
                $BtnRun.FlatAppearance.BorderColor = $Clr.Orange
                $BtnRun.Text        = "Run Migration Day"
            }
        } else {
            $Form.Close()
            return
        }
    } catch { }

    try {
        $screen = [System.Windows.Forms.Screen]::FromControl($Form).WorkingArea
        $Form.SetBounds($screen.X, $screen.Y, $screen.Width, $screen.Height)
        $Form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
    } catch { }

    try { $Split.SplitterDistance = 480 } catch { }
    try { $Form.Activate(); $Form.TopMost = $false } catch { }

    # Set notes file path and load any previously saved notes
    $script:NotesFile = Join-Path $script:NotesFolder "pnfpMigration_Notes.txt"
    Load-Notes
    if($TxtNotes.Text -eq ''){
        $TxtNotes.Text = "Source     : $($TxtSrcSrv.Text)`r`nDestination: $($TxtDstSrv.Text)`r`nMode       : $($script:Mode)`r`nDate       : $(Get-Date -Format 'yyyy-MM-dd')`r`nNotes      : "
        $TxtNotes.ForeColor = $Clr.Dim
    }
    Write-Log "pnfp Migration Tool ready. Mode: $($script:Mode)" $Clr.Blue
    Write-Log "Notes file: $($script:NotesFile)" $Clr.Dim
    Write-Log "Connect source and destination, load databases, then click Run." $Clr.Dim
})

$Form.Add_FormClosing({
    Save-Notes   # auto-save notes on exit
    if($env:DBATOOL_SIG){
        try { Set-Content -Path $env:DBATOOL_SIG -Value $script:CloseAction -Force -ErrorAction SilentlyContinue } catch {}
    }
    if($script:SrcConn){ try{$script:SrcConn.Close()}catch{} }
    if($script:DstConn){ try{$script:DstConn.Close()}catch{} }
})

$Form.ShowDialog() | Out-Null
