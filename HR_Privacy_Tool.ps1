# --- CARICAMENTO LIBRERIE ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Security

# Carica il driver MySQL (Deve essere nella stessa cartella)
$DllPath = Join-Path $PSScriptRoot "MySql.Data.dll"
if (Test-Path $DllPath) {
    Add-Type -Path $DllPath
} else {
    [System.Windows.Forms.MessageBox]::Show("Errore Critico: Manca il file 'MySql.Data.dll'.", "Errore Driver", 0, 16)
    exit
}

# Fix DPI per schermi alta risoluzione
try {
    $code = @"
    using System;
    using System.Runtime.InteropServices;
    public class DPI {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
"@
    Add-Type -TypeDefinition $code -Language CSharp
    [DPI]::SetProcessDPIAware() | Out-Null
} catch { }

# --- CARICAMENTO CONFIGURAZIONE SICURA ---
$ConfigFile = Join-Path $PSScriptRoot "config.json"

if (-not (Test-Path $ConfigFile)) {
    [System.Windows.Forms.MessageBox]::Show("Errore Critico: Manca il file 'config.json' con le credenziali DB!", "Sicurezza", 0, 16)
    exit
}

try {
    $Config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    $ConnString = "Server=$($Config.DbServer);Database=$($Config.DbName);Uid=$($Config.DbUser);Pwd=$($Config.DbPassword);Charset=utf8;"
} catch {
    [System.Windows.Forms.MessageBox]::Show("Errore lettura 'config.json'. Verifica il formato JSON.", "Errore Config", 0, 16)
    exit
}

# --- FUNZIONI DI UTILITY (ROBUSTEZZA) ---
function Get-SafeString ($inputObj) {
    if ($null -eq $inputObj) { return "" }
    $current = $inputObj
    # Srotola array e oggetti complessi (es. Link Excel) fino a trovare il valore
    while ($current -is [Array] -or $current -is [System.Collections.IEnumerable] -and $current -isnot [string]) {
        $current = $current | Select-Object -First 1
    }
    return "$current".Trim()
}

# --- GESTIONE CSV LOCALE ---
function Load-MappingCSV ($csvPath) {
    $map = @{}
    if (Test-Path $csvPath) {
        $lines = Get-Content $csvPath -Encoding UTF8
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $parts = $line -split ";"
            if ($parts.Count -ge 2) {
                $key = $parts[0]
                $val = $parts[1]
                if (-not $map.ContainsKey($key)) { $map[$key] = $val }
            }
        }
    }
    return $map
}

function Append-ToMappingCSV ($csvPath, $originale, $pseudonimo) {
    $line = "$originale;$pseudonimo"
    Add-Content -Path $csvPath -Value $line -Encoding UTF8
}

# --- GESTIONE LOGGING DB ---
function Log-ActivityToDB ($file, $cols, $type) {
    $fileStr = Get-SafeString $file
    $colsStr = Get-SafeString $cols
    $typeStr = Get-SafeString $type

    $conn = New-Object MySql.Data.MySqlClient.MySqlConnection
    $conn.ConnectionString = $ConnString
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "INSERT INTO log_anonimizzazioni (timestamp, utente, file_input, colonne, tipo_operazione) VALUES (NOW(), @user, @file, @cols, @type)"
        $cmd.Parameters.AddWithValue("@user", $env:USERNAME) | Out-Null
        $cmd.Parameters.AddWithValue("@file", [System.IO.Path]::GetFileName($fileStr)) | Out-Null
        $cmd.Parameters.AddWithValue("@cols", $colsStr) | Out-Null
        $cmd.Parameters.AddWithValue("@type", $typeStr) | Out-Null
        $cmd.ExecuteNonQuery()
    } catch { 
        # Non blocchiamo l'utente se il log fallisce, ma potremmo notificarlo
    } finally { $conn.Close() }
}

function Get-AnonHash ($text) {
    try {
        $safeText = Get-SafeString $text
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($safeText)
        $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
        return (-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 12).ToUpper()
    } catch { return "ERROR" }
}

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "HR Privacy Tool (ISO 27001)"
$form.Size = New-Object System.Drawing.Size(600, 650) 
$form.StartPosition = "CenterScreen"
$form.BackColor = "WhiteSmoke"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Strumento Privacy Dati"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(25, 20)
$lblTitle.ForeColor = "#333333"
$form.Controls.Add($lblTitle)

$grpFile = New-Object System.Windows.Forms.GroupBox
$grpFile.Text = "1. Selezione File"
$grpFile.Location = New-Object System.Drawing.Point(25, 70)
$grpFile.Size = New-Object System.Drawing.Size(530, 80)
$form.Controls.Add($grpFile)

$btnFile = New-Object System.Windows.Forms.Button
$btnFile.Text = "Scegli File Excel..."
$btnFile.Size = New-Object System.Drawing.Size(490, 40)
$btnFile.Location = New-Object System.Drawing.Point(20, 25)
$btnFile.BackColor = "#0078D7"
$btnFile.ForeColor = "White"
$btnFile.FlatStyle = "Flat"
$btnFile.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$grpFile.Controls.Add($btnFile)

$grpCols = New-Object System.Windows.Forms.GroupBox
$grpCols.Text = "2. Seleziona colonne da nascondere"
$grpCols.Location = New-Object System.Drawing.Point(25, 160)
$grpCols.Size = New-Object System.Drawing.Size(530, 160)
$form.Controls.Add($grpCols)

$chkList = New-Object System.Windows.Forms.CheckedListBox
$chkList.Location = New-Object System.Drawing.Point(20, 25)
$chkList.Size = New-Object System.Drawing.Size(490, 120)
$chkList.CheckOnClick = $true
$chkList.BorderStyle = "FixedSingle"
$grpCols.Controls.Add($chkList)

$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = "3. Tipo di Protezione"
$grpMode.Location = New-Object System.Drawing.Point(25, 335)
$grpMode.Size = New-Object System.Drawing.Size(530, 130)
$form.Controls.Add($grpMode)

$radPseudo = New-Object System.Windows.Forms.RadioButton
$radPseudo.Text = "Pseudonimizzazione (Reversibile)`nSalva mappatura CSV nella cartella del file."
$radPseudo.Location = New-Object System.Drawing.Point(20, 25)
$radPseudo.AutoSize = $true 
$radPseudo.Checked = $true
$grpMode.Controls.Add($radPseudo)

$radAnon = New-Object System.Windows.Forms.RadioButton
$radAnon.Text = "Anonimizzazione (Irreversibile)`nCifra i dati. Nessun salvataggio mappatura."
$radAnon.Location = New-Object System.Drawing.Point(20, 80) 
$radAnon.AutoSize = $true
$grpMode.Controls.Add($radAnon)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "AVVIA PROCESSO"
$btnRun.Size = New-Object System.Drawing.Size(530, 50)
$btnRun.Location = New-Object System.Drawing.Point(25, 485)
$btnRun.BackColor = "#28a745"
$btnRun.ForeColor = "White"
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnRun.FlatStyle = "Flat"
$form.Controls.Add($btnRun)

# --- LOGICA EVENTI ---
$global:ExcelPath = $null

$btnFile.Add_Click({
    $openDlg = New-Object System.Windows.Forms.OpenFileDialog
    $openDlg.Filter = "Excel Files|*.xlsx;*.xls"
    if ($openDlg.ShowDialog() -eq "OK") {
        $global:ExcelPath = $openDlg.FileName
        $btnFile.Text = [System.IO.Path]::GetFileName($global:ExcelPath)
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $wb = $excel.Workbooks.Open($global:ExcelPath, 0, $true) 
            $ws = $wb.Sheets.Item(1)
            $chkList.Items.Clear()
            $col = 1
            while ($col -le 50) {
                $raw = $ws.Cells.Item(1, $col).Value2
                $val = Get-SafeString $raw
                if ([string]::IsNullOrWhiteSpace($val)) { break }
                $chkList.Items.Add($val) | Out-Null
                $col++
            }
            $wb.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch { [System.Windows.Forms.MessageBox]::Show("Errore Excel: " + $_, "Errore", 0, 16) }
    }
})

$btnRun.Add_Click({
    if (-not $global:ExcelPath) { [System.Windows.Forms.MessageBox]::Show("Seleziona un file!", "Stop"); return }
    if ($chkList.CheckedItems.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Seleziona colonne!", "Stop"); return }

    $btnRun.Text = "Elaborazione in corso..."
    $btnRun.Enabled = $false
    $form.Refresh() 

    $isPseudo = $radPseudo.Checked
    
    $colsToProcess = @()
    foreach ($item in $chkList.CheckedItems) {
        $colsToProcess += Get-SafeString $item
    }
    
    # Logica CSV Dinamica: Il CSV vive accanto al file Excel
    $dir = [System.IO.Path]::GetDirectoryName($global:ExcelPath)
    $dynamicCsvPath = Join-Path $dir "mappatura_segreta.csv"

    $localMap = @{}
    if ($isPseudo) {
        $localMap = Load-MappingCSV $dynamicCsvPath
    }
    $pseudoCounter = $localMap.Count + 1

    try {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($global:ExcelPath)
        $ext = [System.IO.Path]::GetExtension($global:ExcelPath)
        $suffix = if ($isPseudo) { "_PSEUDO" } else { "_ANON" }
        $newPath = Join-Path $dir ($name + $suffix + $ext)

        Copy-Item -Path $global:ExcelPath -Destination $newPath -Force

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($newPath)
        $ws = $wb.Sheets.Item(1)

        $colIndices = @{}
        $c = 1
        while ($c -le 100) {
            $raw = $ws.Cells.Item(1, $c).Value2
            $h = Get-SafeString $raw
            if ([string]::IsNullOrWhiteSpace($h)) { break }
            if ($colsToProcess -contains $h) { $colIndices[$c] = $h }
            $c++
        }

        $usedRange = $ws.UsedRange
        $rowCount = $usedRange.Rows.Count
        
        for ($r = 2; $r -le $rowCount; $r++) {
            foreach ($cIdx in $colIndices.Keys) {
                
                $rawVal = $ws.Cells.Item($r, $cIdx).Value2
                $cellVal = Get-SafeString $rawVal
                if ([string]::IsNullOrWhiteSpace($cellVal)) { continue }

                $newVal = ""
                
                if ($isPseudo) {
                    if ($localMap.ContainsKey($cellVal)) {
                        $newVal = $localMap[$cellVal]
                    } else {
                        $newVal = "USER_{0:D5}" -f $pseudoCounter
                        $pseudoCounter++
                        $localMap[$cellVal] = $newVal
                        Append-ToMappingCSV $dynamicCsvPath $cellVal $newVal
                    }
                } else {
                    $newVal = Get-AnonHash $cellVal
                }
                
                $ws.Cells.Item($r, $cIdx).Value2 = $newVal
            }
        }

        $wb.Save() 
        $wb.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        # Logging DB
        $logType = if ($isPseudo) { "Pseudonimizzazione" } else { "Anonimizzazione" }
        $colsString = $colsToProcess -join ", "
        Log-ActivityToDB $global:ExcelPath $colsString $logType

        $btnRun.Text = "AVVIA PROCESSO"
        $btnRun.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show("Successo!`nFile creato: $newPath`nMappatura (se attiva) salvata nella cartella del file.", "ISO 27001", 0, 64)

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore: " + $_, "Errore", 0, 16)
        if ($excel) { $excel.Quit() }
        $btnRun.Text = "AVVIA PROCESSO"
        $btnRun.Enabled = $true
    }
})

$form.ShowDialog()