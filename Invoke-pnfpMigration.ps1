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

# PNFP-AUTO_MIGRATION folder on the local system drive - created automatically if missing.
# This is where the script lives and where all outputs (notes, scripts, connections) are saved.
$script:AutoMigDir = Join-Path $env:SystemDrive "PNFP-AUTO_MIGRATION"
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
    return $dt
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
function Invoke-ConnectionCapture {
    if(-not $script:SrcConn -or $script:SrcConn.State -ne 'Open'){
        [System.Windows.Forms.MessageBox]::Show(
            "Connect to the source server first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    Write-Sep "CONNECTION CAPTURE - AIDO"
    Set-Status "Running connection capture from AIDO..." $Clr.Yellow

    # Output folder
    $connFolder = Join-Path $script:ScriptDir "Connections"
    if(-not (Test-Path $connFolder)){
        try { New-Item -ItemType Directory -Path $connFolder -Force | Out-Null }
        catch {
            Write-Log "Could not create Connections folder: $($_.Exception.Message)" $Clr.Red
            return
        }
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"

    # ---- Run the date range query ----
    $rangeFrom = ""; $rangeTo = ""
    try {
        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.Connection     = $script:SrcConn
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
    } catch {
        Write-Log "  Warning: could not read date range - $($_.Exception.Message)" $Clr.Yellow
    }

    # ---- Run the main connections query ----
    $dt = New-Object System.Data.DataTable
    try {
        $cmd2 = New-Object System.Data.SqlClient.SqlCommand
        $cmd2.Connection     = $script:SrcConn
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
  AND server_principal_name <> 'SNVX17'
  AND server_principal_name <> 'SNV8604'
  AND server_principal_name <> 'SNV\svc_Redgate_service'
  AND server_principal_name <> 'SNVX55'
  AND server_principal_name <> 'SNV837'
  AND server_principal_name <> 'SNV9776'
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
# ---------------------------------------------------------------
function Invoke-ScriptOut {
    $dbs = Get-SelectedDatabases
    if(-not $script:Mode){
        [System.Windows.Forms.MessageBox]::Show("Select a mode first.", "No Mode",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $stamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $outDir = $script:NotesFolder
    if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $script:ScriptDir }
    if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $env:TEMP }

    $modeLbl  = if($script:Mode -eq "PRE"){"PreMigration"}else{"PostMigration"}
    $fileName = "pnfp_${modeLbl}_Script_${stamp}.sql"
    $outPath  = Join-Path $outDir $fileName

    $srcSrv = $TxtSrcSrv.Text.Trim()
    $dstSrv = $TxtDstSrv.Text.Trim()
    $notes  = $TxtNotes.Text.Trim()

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("/*")
    [void]$sb.AppendLine("  pnfp Migration Script")
    [void]$sb.AppendLine("  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("  Mode      : $($script:Mode)")
    [void]$sb.AppendLine("  Source    : $srcSrv")
    [void]$sb.AppendLine("  Destination: $dstSrv")
    [void]$sb.AppendLine("  Databases : $(if($dbs.Count -gt 0){ $dbs -join ', ' } else { '(none selected)' })")
    if($notes){
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("  MIGRATION NOTES:")
        foreach($line in ($notes -split "`r?`n")){ [void]$sb.AppendLine("  $line") }
    }
    [void]$sb.AppendLine("*/")
    [void]$sb.AppendLine("")

    if($script:Mode -eq "PRE"){
        foreach($db in $dbs){
            [void]$sb.AppendLine("/* ============================================================ */")
            [void]$sb.AppendLine("/* DATABASE: $db                                              */")
            [void]$sb.AppendLine("/* ============================================================ */")
            [void]$sb.AppendLine("")

            [void]$sb.AppendLine("-- STEP 1: Copy-only backup on $srcSrv")
            [void]$sb.AppendLine("-- (Run this on the SOURCE server: $srcSrv)")
            [void]$sb.AppendLine("BACKUP DATABASE [$db]")
            [void]$sb.AppendLine("TO DISK = N'<your_backup_path>\${db}_COPYONLY_$stamp.bak'")
            [void]$sb.AppendLine("WITH COPY_ONLY, COMPRESSION, STATS = 10, FORMAT,")
            [void]$sb.AppendLine("     NAME = N'$db - pnfp Copy-Only Backup';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")

            [void]$sb.AppendLine("-- STEP 2: Restore to destination $dstSrv")
            [void]$sb.AppendLine("-- (Run this on the DESTINATION server: $dstSrv)")
            [void]$sb.AppendLine("-- First check the logical file names:")
            [void]$sb.AppendLine("RESTORE FILELISTONLY FROM DISK = N'<your_backup_path>\${db}_COPYONLY_$stamp.bak';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("-- Then restore (update MOVE paths to match destination data/log directories):")
            [void]$sb.AppendLine("RESTORE DATABASE [$db]")
            [void]$sb.AppendLine("FROM DISK = N'<your_backup_path>\${db}_COPYONLY_$stamp.bak'")
            [void]$sb.AppendLine("WITH REPLACE, RECOVERY, STATS = 10,")
            [void]$sb.AppendLine("     MOVE N'<logical_data_name>' TO N'<data_dir>\${db}.mdf',")
            [void]$sb.AppendLine("     MOVE N'<logical_log_name>'  TO N'<log_dir>\${db}_log.ldf';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")

            [void]$sb.AppendLine("-- STEP 3: Update statistics")
            [void]$sb.AppendLine("USE [$db];")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("EXEC sp_updatestats;")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")

            [void]$sb.AppendLine("-- STEP 4: Set compatibility level")
            [void]$sb.AppendLine("-- Replace 150 with your chosen level (100/110/120/130/140/150/160)")
            [void]$sb.AppendLine("ALTER DATABASE [$db] SET COMPATIBILITY_LEVEL = 150;")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")

            [void]$sb.AppendLine("-- STEP 5: Find orphaned users")
            [void]$sb.AppendLine("USE [$db];")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("SELECT dp.name AS OrphanedUser, dp.type_desc")
            [void]$sb.AppendLine("FROM sys.database_principals dp")
            [void]$sb.AppendLine("WHERE dp.type IN ('S','U','G')")
            [void]$sb.AppendLine("  AND dp.name NOT IN ('dbo','guest','INFORMATION_SCHEMA','sys')")
            [void]$sb.AppendLine("  AND dp.sid IS NOT NULL AND dp.sid <> 0x00")
            [void]$sb.AppendLine("  AND NOT EXISTS (SELECT 1 FROM sys.server_principals sp WHERE sp.sid = dp.sid);")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("-- Fix each orphan (repeat for each user returned above):")
            [void]$sb.AppendLine("-- ALTER USER [<username>] WITH LOGIN = [<username>];")
            [void]$sb.AppendLine("-- GO")
            [void]$sb.AppendLine("-- If the login does not exist, create it first:")
            [void]$sb.AppendLine("-- CREATE LOGIN [<username>] WITH PASSWORD = '<password>';")
            [void]$sb.AppendLine("-- GO")
            [void]$sb.AppendLine("")
        }
    } else {
        # POST-MIGRATION script
        foreach($db in $dbs){
            [void]$sb.AppendLine("/* ============================================================ */")
            [void]$sb.AppendLine("/* DATABASE: $db                                              */")
            [void]$sb.AppendLine("/* ============================================================ */")
            [void]$sb.AppendLine("")

            [void]$sb.AppendLine("-- STEP 1: Copy-only backup on $srcSrv")
            [void]$sb.AppendLine("-- (Run on SOURCE: $srcSrv)")
            [void]$sb.AppendLine("BACKUP DATABASE [$db]")
            [void]$sb.AppendLine("TO DISK = N'<your_backup_path>\${db}_POST_COPYONLY_$stamp.bak'")
            [void]$sb.AppendLine("WITH COPY_ONLY, COMPRESSION, STATS = 10, FORMAT,")
            [void]$sb.AppendLine("     NAME = N'$db - pnfp Post-Migration Copy-Only Backup';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")

            [void]$sb.AppendLine("-- STEP 2: Restore to destination $dstSrv")
            [void]$sb.AppendLine("-- (Run on DESTINATION: $dstSrv)")
            [void]$sb.AppendLine("-- Choose options: WITH REPLACE overrides existing DB, WITH RECOVERY brings it online.")
            [void]$sb.AppendLine("-- Remove WITH REPLACE if the database does not yet exist at destination.")
            [void]$sb.AppendLine("-- Use NORECOVERY instead of RECOVERY if you plan to apply more log backups.")
            [void]$sb.AppendLine("RESTORE FILELISTONLY FROM DISK = N'<your_backup_path>\${db}_POST_COPYONLY_$stamp.bak';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("-- Check logical file names first:")
            [void]$sb.AppendLine("RESTORE FILELISTONLY FROM DISK = N'<your_backup_path>\${db}_POST_COPYONLY_$stamp.bak';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("-- Close existing connections before restore:")
            [void]$sb.AppendLine("ALTER DATABASE [$db] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("RESTORE DATABASE [$db]")
            [void]$sb.AppendLine("FROM DISK = N'<your_backup_path>\${db}_POST_COPYONLY_$stamp.bak'")
            [void]$sb.AppendLine("WITH")
            [void]$sb.AppendLine("     REPLACE,                          -- Remove if database does not exist at destination")
            [void]$sb.AppendLine("     RECOVERY,                         -- Change to NORECOVERY to apply more logs after, or STANDBY for read-only between restores")
            [void]$sb.AppendLine("     -- KEEP_REPLICATION,             -- Uncomment to preserve replication settings")
            [void]$sb.AppendLine("     -- RESTRICTED_USER,              -- Uncomment to limit access to sysadmin/dbcreator/db_owner after restore")
            [void]$sb.AppendLine("     STATS = 10,")
            [void]$sb.AppendLine("     MOVE N'<logical_data_name>' TO N'<data_dir>\${db}.mdf',")
            [void]$sb.AppendLine("     MOVE N'<logical_log_name>'  TO N'<log_dir>\${db}_log.ldf';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
        }
    }

    # For post-migration, append a READ_ONLY template section
    if($script:Mode -eq "POST" -and $dbs.Count -gt 0){
        [void]$sb.AppendLine("/* ============================================================ */")
        [void]$sb.AppendLine("/* OPTIONAL: SET SOURCE DATABASES READ_ONLY                   */")
        [void]$sb.AppendLine("/* Run on SOURCE server after confirming destination is good  */")
        [void]$sb.AppendLine("/* ============================================================ */")
        [void]$sb.AppendLine("")
        foreach($db in $dbs){
            [void]$sb.AppendLine("-- Set [$db] READ_ONLY on source (prevents accidental writes during validation)")
            [void]$sb.AppendLine("ALTER DATABASE [$db] SET READ_ONLY WITH ROLLBACK IMMEDIATE;")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("SELECT name, is_read_only FROM sys.databases WHERE name = N'$db';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
        }
        [void]$sb.AppendLine("-- To undo READ_ONLY (rollback plan):")
        foreach($db in $dbs){
            [void]$sb.AppendLine("-- ALTER DATABASE [$db] SET READ_WRITE WITH ROLLBACK IMMEDIATE;")
        }
        [void]$sb.AppendLine("")
    }

    [void]$sb.AppendLine("/* End of pnfp Migration Script */")

    try {
        $sb.ToString() | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Write-Log "Script saved to: $outPath" $Clr.Green
        Set-Status "Script written: $fileName" $Clr.Green

        $open = [System.Windows.Forms.MessageBox]::Show(
            "Script saved to:`n$outPath`n`nOpen containing folder?",
            "Script Saved",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information)
        if($open -eq [System.Windows.Forms.DialogResult]::Yes){
            Start-Process explorer.exe -ArgumentList "/select,`"$outPath`""
        }
    } catch {
        Write-Log "Script save failed: $($_.Exception.Message)" $Clr.Red
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
    $dlg.Size            = [System.Drawing.Size]::new(500, 420)
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

    # Checklist
    $ChkList = New-Object System.Windows.Forms.CheckedListBox
    $ChkList.Location     = [System.Drawing.Point]::new(16, 120)
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
    $btnAll.Text="All"; $btnAll.Location=[System.Drawing.Point]::new(16,298)
    $btnAll.Size=[System.Drawing.Size]::new(60,24); $btnAll.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnAll.BackColor=$Clr.Card; $btnAll.ForeColor=$Clr.Text; $btnAll.Font=$FontSm
    $btnAll.FlatAppearance.BorderColor=$Clr.Border
    $btnAll.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$true) } })
    $dlg.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text="None"; $btnNone.Location=[System.Drawing.Point]::new(82,298)
    $btnNone.Size=[System.Drawing.Size]::new(60,24); $btnNone.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
    $btnNone.BackColor=$Clr.Card; $btnNone.ForeColor=$Clr.Text; $btnNone.Font=$FontSm
    $btnNone.FlatAppearance.BorderColor=$Clr.Border
    $btnNone.Add_Click({ for($i=0;$i -lt $ChkList.Items.Count;$i++){ $ChkList.SetItemChecked($i,$false) } })
    $dlg.Controls.Add($btnNone)

    # Warning note
    $warn = New-Object System.Windows.Forms.Label
    $warn.Text      = "Note: READ_ONLY can be reversed with ALTER DATABASE [name] SET READ_WRITE."
    $warn.Location  = [System.Drawing.Point]::new(16, 328)
    $warn.Size      = [System.Drawing.Size]::new(462, 18)
    $warn.ForeColor = $Clr.Yellow
    $warn.Font      = $FontSm
    $dlg.Controls.Add($warn)

    # Buttons
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text      = "Apply + Script Out"
    $btnApply.Location  = [System.Drawing.Point]::new(196, 352)
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
    $btnScriptOnly.Location  = [System.Drawing.Point]::new(334, 352)
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
    $btnSkip.Location  = [System.Drawing.Point]::new(432, 352)
    $btnSkip.Size      = [System.Drawing.Size]::new(52, 28)
    $btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkip.BackColor = $Clr.Card
    $btnSkip.ForeColor = $Clr.Dim
    $btnSkip.Font      = $FontSm
    $btnSkip.FlatAppearance.BorderColor = $Clr.Border
    $btnSkip.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnSkip)

    $dlg.Add_Shown({ $dlg.Activate() })
    $r = $dlg.ShowDialog()

    $chosen = @()
    foreach($i in $ChkList.CheckedIndices){ $chosen += $ChkList.Items[$i].ToString() }
    $dlg.Dispose()

    return [PSCustomObject]@{
        Action    = $r          # OK=apply+script, Retry=script only, Cancel=skip
        Databases = $chosen
    }
}

# Helper: set databases READ_ONLY on source and/or script it out
function Apply-ReadOnly {
    param(
        [string[]]$Databases,
        [bool]$Execute,
        [bool]$ScriptOut
    )

    if($Databases.Count -eq 0){
        Write-Log "  No databases selected for READ_ONLY." $Clr.Dim; return
    }

    if($ScriptOut){
        $stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
        $outDir  = $script:NotesFolder
        if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $script:ScriptDir }
        if(-not $outDir -or -not (Test-Path $outDir)){ $outDir = $env:TEMP }
        $outPath = Join-Path $outDir "pnfp_SetReadOnly_$stamp.sql"

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("/*")
        [void]$sb.AppendLine("  pnfp Migration - Set Source Databases READ_ONLY")
        [void]$sb.AppendLine("  Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$sb.AppendLine("  Source     : $($TxtSrcSrv.Text.Trim())")
        [void]$sb.AppendLine("  Databases  : $($Databases -join ', ')")
        [void]$sb.AppendLine("  Run this on the SOURCE server.")
        [void]$sb.AppendLine("  To reverse: ALTER DATABASE [name] SET READ_WRITE WITH ROLLBACK IMMEDIATE;")
        [void]$sb.AppendLine("*/")
        [void]$sb.AppendLine("")
        foreach($db in $Databases){
            [void]$sb.AppendLine("-- Set [$db] to READ_ONLY (kicks existing connections with rollback)")
            [void]$sb.AppendLine("ALTER DATABASE [$db] SET READ_ONLY WITH ROLLBACK IMMEDIATE;")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("-- Verify:")
            [void]$sb.AppendLine("SELECT name, is_read_only FROM sys.databases WHERE name = N'$db';")
            [void]$sb.AppendLine("GO")
            [void]$sb.AppendLine("")
        }
        [void]$sb.AppendLine("/*")
        [void]$sb.AppendLine("  To undo READ_ONLY on all databases above:")
        foreach($db in $Databases){
            [void]$sb.AppendLine("  ALTER DATABASE [$db] SET READ_WRITE WITH ROLLBACK IMMEDIATE;")
        }
        [void]$sb.AppendLine("*/")

        try {
            $sb.ToString() | Out-File -FilePath $outPath -Encoding UTF8 -Force
            Write-Log "  READ_ONLY script saved: $outPath" $Clr.Green
        } catch {
            Write-Log "  Script save failed: $($_.Exception.Message)" $Clr.Red
        }
    }

    if($Execute){
        if(-not $script:SrcConn -or $script:SrcConn.State -ne 'Open'){
            Write-Log "  Source not connected - cannot apply READ_ONLY." $Clr.Red; return
        }
        foreach($db in $Databases){
            try {
                Write-Log "  Setting [$db] READ_ONLY on source..." $Clr.Dim
                Exec-Sql $script:SrcConn "ALTER DATABASE [$db] SET READ_ONLY WITH ROLLBACK IMMEDIATE"
                Write-Log "  [$db] is now READ_ONLY" $Clr.Green
            } catch {
                Write-Log "  FAILED to set [$db] READ_ONLY: $($_.Exception.Message)" $Clr.Red
            }
        }
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
    $lPre2.Text      = "- Copy-only backup on source`n- Restore to destination`n- Update statistics`n- Set compatibility level`n- Fix orphaned users"
    $lPre2.Location  = [System.Drawing.Point]::new(10, 38)
    $lPre2.Size      = [System.Drawing.Size]::new(200, 85)
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
    $lPost1.Text      = "POST-MIGRATION"
    $lPost1.Location  = [System.Drawing.Point]::new(10, 14)
    $lPost1.Size      = [System.Drawing.Size]::new(200, 22)
    $lPost1.ForeColor = $Clr.Orange
    $lPost1.Font      = $FontUIB
    $lPost1.BackColor = [System.Drawing.Color]::Transparent
    $cardPost.Controls.Add($lPost1)

    $lPost2 = New-Object System.Windows.Forms.Label
    $lPost2.Text      = "- Copy-only backup on source`n- Restore to destination`n  (WITH REPLACE or WITH RECOVERY)`n- Your choice of restore options"
    $lPost2.Location  = [System.Drawing.Point]::new(10, 38)
    $lPost2.Size      = [System.Drawing.Size]::new(200, 85)
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
# MODAL 2: POST-MIGRATION RESTORE OPTIONS
# ---------------------------------------------------------------
function Show-RestoreOptionsDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Post-Migration - Restore Options"
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
        $l.Size=[System.Drawing.Size]::new(504,($Lines*16)); $l.ForeColor=$Clr.Dim
        $l.Font=$FontSm; $l.BackColor=[System.Drawing.Color]::Transparent
        $Parent.Controls.Add($l)
        return $Y+($Lines*16)+4
    }

    $Y = 56

    # ---- RESTORE OPTIONS ----
    $Y = Add-Section $dlg "  RESTORE OPTIONS" $Y
    $Y += 4

    $chkReplace = Add-Opt $dlg "Overwrite the existing database  (WITH REPLACE)" $true 18 $Y
    $Y = Add-Desc $dlg "Use when the database already exists at the destination and you want to completely replace it." $Y 2

    $chkKeepReplication = Add-Opt $dlg "Preserve replication settings  (WITH KEEP_REPLICATION)" $false 18 $Y
    $Y = Add-Desc $dlg "Keeps replication settings when restoring a published database. Only needed if the destination participates in replication." $Y 2

    $chkRestrictedUser = Add-Opt $dlg "Restrict access to the restored database  (WITH RESTRICTED_USER)" $false 18 $Y
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

# Build the backup path for a database on the source server
function Get-BackupPath {
    param([System.Data.SqlClient.SqlConnection]$Conn,[string]$DbName)

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"

    # 1. If the user specified a custom folder in the UI, use that
    $customDir = $TxtBackupDir.Text.Trim()
    $isAutoText = $customDir -like "*(auto*" -or $customDir -eq ""
    if(-not $isAutoText -and (Test-Path $customDir)){
        return "$($customDir.TrimEnd('\'))\${DbName}_COPYONLY_$stamp.bak"
    }

    # 2. Try SQL Server default backup directory via SERVERPROPERTY (SQL 2012+, no sysadmin needed)
    $dir = ""
    try {
        $dir = Query-Scalar $Conn "SELECT CONVERT(NVARCHAR(512), SERVERPROPERTY('InstanceDefaultDataPath'))"
        # SERVERPROPERTY gives the data path - back up one level to get Backup sibling
        if($dir){
            $parent = Split-Path $dir.TrimEnd('') -Parent
            $candidate = Join-Path $parent "Backup"
            if(Test-Path $candidate){ $dir = $candidate }
            else { $dir = $parent }
        }
    } catch { $dir = "" }

    # 3. Try registry via xp_instance_regread (requires sysadmin)
    if(-not $dir){
        try {
            $dir = Query-Scalar $Conn "EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'Software\Microsoft\MSSQLServer\MSSQLServer',N'BackupDirectory'"
        } catch { $dir = "" }
    }

    # 4. Last resort - C:\PNFP-MIGRATION (our known writable folder)
    if(-not $dir -or -not (Test-Path $dir)){
        $dir = $script:NotesFolder
        if(-not $dir -or -not (Test-Path $dir)){ $dir = "C:\Temp" }
        Write-Log "  Warning: could not determine SQL backup directory. Using $dir" $Clr.Yellow
    }

    return "$($dir.TrimEnd('\'))\${DbName}_COPYONLY_$stamp.bak"
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
        Write-Log "  Use the Browse button in section 5 to choose an accessible folder." $Clr.Red
        throw "Backup directory not found: $backupDir"
    }

    $sql = @"
BACKUP DATABASE [$DbName]
TO DISK = N'$BackupPath'
WITH COPY_ONLY, COMPRESSION, STATS = 10, INIT,
     NAME = N'${DbName} - pnfp Copy-Only Backup';
"@
    Exec-Sql $Conn $sql 7200
    Write-Log "  Backup complete: $BackupPath" $Clr.Green
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

    # Read logical file names using SqlDataAdapter (more reliable than DataTable.Load)
    Write-Log "  Reading backup header..." $Clr.Dim
    $files = @()
    try {
        $cmd = New-Object System.Data.SqlClient.SqlCommand ("RESTORE FILELISTONLY FROM DISK = N'$BackupPath'"), $Conn
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
FROM DISK = N'$BackupPath'
WITH
$moveStr,
$optStr;
"@
    try {
        Exec-Sql $Conn $sql 7200
        Write-Log "  Restore complete: [$restoreName]" $Clr.Green
    } catch {
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
    Exec-Sql $Conn $sql 3600
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

        foreach($db in $dbs){
            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            Write-Sep "PRE-MIGRATION: [$db]"
            Set-Status "Processing [$db]..." $Clr.Yellow

            # 1. Copy-only backup on source
            Write-Log "STEP 1 - Copy-only backup on source" $Clr.Blue
            $backupPath = Get-BackupPath $script:SrcConn $db
            try {
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
                Write-Log "  Pre-migration expects a fresh destination. If you want to overwrite, use POST-MIGRATION mode." $Clr.Yellow
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

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            # 3. Update statistics on the restored database name
            Write-Log "STEP 3 - Update statistics on [$restoredName]" $Clr.Blue
            try {
                Update-Statistics $script:DstConn $restoredName
            } catch {
                Write-Log "  UPDATE STATS FAILED: $($_.Exception.Message)" $Clr.Red
            }

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            # 4. Set compatibility level on the restored database name
            if($null -ne $compatTarget){
                Write-Log "STEP 4 - Setting compatibility level to $compatTarget on [$restoredName]" $Clr.Blue
                try {
                    Set-CompatibilityLevel $script:DstConn $restoredName $compatTarget
                } catch {
                    Write-Log "  COMPAT LEVEL FAILED: $($_.Exception.Message)" $Clr.Red
                }
            }

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            # 5. Fix orphaned users
            Write-Log "STEP 5 - Find and fix orphaned users in [$restoredName]" $Clr.Blue
            try {
                Fix-OrphanedUsers $script:DstConn $restoredName
            } catch {
                Write-Log "  ORPHAN FIX FAILED: $($_.Exception.Message)" $Clr.Red
            }

            Write-Log "  [$db] complete." $Clr.Green
        }

        if(-not $script:Cancel){
            Write-Sep "PRE-MIGRATION COMPLETE"
            Write-Log "All selected databases processed." $Clr.Green
            Set-Status "Pre-migration complete." $Clr.Green
        }

    } finally {
        $script:Running    = $false
        $BtnRun.Enabled    = $true
        $BtnCancel.Enabled = $false
    }
}

# ---------------------------------------------------------------
# MAIN PIPELINE: POST-MIGRATION
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
            Write-Log "STEP 4 - Compatibility level: skipped by user." $Clr.Dim
        } else {
            Write-Log "STEP 4 - Compatibility level chosen: $compatTarget" $Clr.Blue
        }

        foreach($db in $dbs){
            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            Write-Sep "POST-MIGRATION: [$db]"
            Set-Status "Processing [$db]..." $Clr.Yellow

            # 1. Copy-only backup on source
            Write-Log "STEP 1 - Copy-only backup on source" $Clr.Blue
            $backupPath = Get-BackupPath $script:SrcConn $db
            try {
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

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            # 3. Update statistics
            Write-Log "STEP 3 - Update statistics on [$restoredName]" $Clr.Blue
            try {
                Update-Statistics $script:DstConn $restoredName
            } catch {
                Write-Log "  UPDATE STATS FAILED: $($_.Exception.Message)" $Clr.Red
            }

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            # 4. Set compatibility level
            if($null -ne $compatTarget){
                Write-Log "STEP 4 - Setting compatibility level to $compatTarget on [$restoredName]" $Clr.Blue
                try {
                    Set-CompatibilityLevel $script:DstConn $restoredName $compatTarget
                } catch {
                    Write-Log "  COMPAT LEVEL FAILED: $($_.Exception.Message)" $Clr.Red
                }
            }

            if($script:Cancel){ Write-Log "Cancelled by user." $Clr.Orange; break }

            # 5. Fix orphaned users
            Write-Log "STEP 5 - Find and fix orphaned users in [$restoredName]" $Clr.Blue
            try {
                Fix-OrphanedUsers $script:DstConn $restoredName
            } catch {
                Write-Log "  ORPHAN FIX FAILED: $($_.Exception.Message)" $Clr.Red
            }

            Write-Log "  [$db] complete." $Clr.Green
        }

        if(-not $script:Cancel){
            Write-Sep "POST-MIGRATION COMPLETE"
            Write-Log "All selected databases processed." $Clr.Green
            Set-Status "Post-migration complete." $Clr.Green

            # Offer to set source databases READ_ONLY
            $roResult = Show-ReadOnlyDialog -Databases $dbs
            if($roResult -and $roResult.Action -ne [System.Windows.Forms.DialogResult]::Cancel){
                Write-Sep "SET SOURCE READ_ONLY"
                $execute  = ($roResult.Action -eq [System.Windows.Forms.DialogResult]::OK)
                $scriptIt = $true
                Apply-ReadOnly -Databases $roResult.Databases -Execute $execute -ScriptOut $scriptIt
            } else {
                Write-Log "Set source READ_ONLY: skipped." $Clr.Dim
            }
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
$ConnCapHint.Text      = "Requires AIDO database on source. Saves CSV + XLSX to Connections folder."
$ConnCapHint.Location  = [System.Drawing.Point]::new(198,$Y+4)
$ConnCapHint.Size      = [System.Drawing.Size]::new(242,18)
$ConnCapHint.ForeColor = $Clr.Dim
$ConnCapHint.Font      = $FontSm
$ConnCapHint.BackColor = [System.Drawing.Color]::Transparent
$LeftScroll.Controls.Add($ConnCapHint)
$Y += 36

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
$bkl = New-Object System.Windows.Forms.Label; $bkl.Text="  5  BACKUP FOLDER (on source server)"
$bkl.Dock=[System.Windows.Forms.DockStyle]::Fill; $bkl.ForeColor=$Clr.Blue; $bkl.Font=$FontSec
$bkl.TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft; $BkSec.Controls.Add($bkl)
$LeftScroll.Controls.Add($BkSec)
$Y += 34

$TxtBackupDir = New-Object System.Windows.Forms.TextBox
$TxtBackupDir.Text="(auto - uses SQL Server default backup directory)"
$TxtBackupDir.Location=[System.Drawing.Point]::new(10,$Y); $TxtBackupDir.Size=[System.Drawing.Size]::new(310,24)
$TxtBackupDir.BackColor=$Clr.Input; $TxtBackupDir.ForeColor=$Clr.Dim; $TxtBackupDir.Font=$FontSm
$TxtBackupDir.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$LeftScroll.Controls.Add($TxtBackupDir)

$BtnBrowse = New-Object System.Windows.Forms.Button
$BtnBrowse.Text="Browse"; $BtnBrowse.Location=[System.Drawing.Point]::new(328,$Y)
$BtnBrowse.Size=[System.Drawing.Size]::new(70,24); $BtnBrowse.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat
$BtnBrowse.BackColor=$Clr.Card; $BtnBrowse.ForeColor=$Clr.Text; $BtnBrowse.Font=$FontSm
$BtnBrowse.FlatAppearance.BorderColor=$Clr.Border
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
        if($m -eq "PRE"){
            $HModeLbl.Text      = "  Mode: PRE-MIGRATION  |  Backup -> Restore -> Update Stats -> Set Compat Level -> Fix Orphans"
            $ModeBar.BackColor  = $Clr.Green
            $BtnRun.BackColor   = $Clr.Green
            $BtnRun.FlatAppearance.BorderColor = $Clr.Green
            $BtnRun.Text        = "Run Pre-Migration"
        } else {
            $HModeLbl.Text      = "  Mode: POST-MIGRATION  |  Backup -> Restore (WITH REPLACE / WITH RECOVERY options)"
            $ModeBar.BackColor  = $Clr.Orange
            $BtnRun.BackColor   = $Clr.Orange
            $BtnRun.FlatAppearance.BorderColor = $Clr.Orange
            $BtnRun.Text        = "Run Post-Migration"
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
$BtnCancel.Add_Click({ $script:Cancel = $true; Set-Status "Cancelling after current step..." $Clr.Orange })

# Clear log
$BtnClearLog.Add_Click({ $RtbLog.Clear() })
$BtnScriptOut.Add_Click({ Invoke-ScriptOut })
$BtnConnCapture.Add_Click({ Invoke-ConnectionCapture })

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
            if($m -eq "PRE"){
                $HModeLbl.Text      = "  Mode: PRE-MIGRATION  |  Backup -> Restore -> Update Stats -> Set Compat Level -> Fix Orphans"
                $ModeBar.BackColor  = $Clr.Green
                $BtnRun.BackColor   = $Clr.Green
                $BtnRun.FlatAppearance.BorderColor = $Clr.Green
                $BtnRun.Text        = "Run Pre-Migration"
            } else {
                $HModeLbl.Text      = "  Mode: POST-MIGRATION  |  Backup -> Restore (WITH REPLACE / WITH RECOVERY options)"
                $ModeBar.BackColor  = $Clr.Orange
                $BtnRun.BackColor   = $Clr.Orange
                $BtnRun.FlatAppearance.BorderColor = $Clr.Orange
                $BtnRun.Text        = "Run Post-Migration"
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
