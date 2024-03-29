Add-Type -AssemblyName System.Windows.Forms

# Settings
$Global:RegistryKey     = 'HKCU:\SOFTWARE\JenTronic\Label Builder'
$Global:Templatesfolder = "$PSScriptRoot\Templates"
$Global:Outfolder       = "$PSScriptRoot\..\Labels"

# Files
$Global:PrinterSettings = "$PSScriptRoot\Printer Functions.ps1"
$Global:Icon            = "$PSScriptRoot\JenTronic.ico"

# Global vars
$Global:Filename        = ""

# Load required assembly
[System.Reflection.Assembly]::LoadWithPartialName("MySql.Data") | Out-Null

# Load printer settings
. $Global:PrinterSettings

# Function for GZip-compressing data
function Get-CompressedByteArray {

   [CmdletBinding()]
   Param (
      [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
      [byte[]]$byteArray = $(Throw("-byteArray is required"))
   )
	Process {
      
      [System.IO.MemoryStream] $output = New-Object System.IO.MemoryStream
      $gzipStream = New-Object System.IO.Compression.GzipStream $output, ([IO.Compression.CompressionMode]::Compress)
      $gzipStream.Write( $byteArray, 0, $byteArray.Length )
      $gzipStream.Close()
      $output.Close()
      Write-Output $output.ToArray()
   }

}

# Function for decompressing GZip-compressed data
function Get-DecompressedByteArray {

   [CmdletBinding()]
   Param (
      [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
      [byte[]] $byteArray = $(Throw("-byteArray is required"))
   )
   Process {
      
      $output = New-Object System.IO.MemoryStream
      $gzipStream = New-Object System.IO.Compression.GzipStream $(New-Object System.IO.MemoryStream(, $byteArray)), ([IO.Compression.CompressionMode]::Decompress)
      $gzipStream.CopyTo( $output )
      $gzipStream.Close()
      Write-Output $output.ToArray()
   
   }

}

# Fetch data from the remote PrestaShop database, and prepare label for final editing
Function GenerateLabel {

   $Panel_Generate.Enabled = $false
   $Panel_Print.Enabled = $false
   $Panel_Database.Enabled = $false

   [System.Windows.Forms.Application]::DoEvents()

  try {

      # Retrieve address from the database
      $MYSQL_Connection = [MySql.Data.MySqlClient.MySqlConnection]@{ConnectionString="server=$($Textbox_Server.Text);uid=$($Textbox_Username.text);pwd=$($Textbox_Password.Text);database=$($Textbox_Database.Text);port=3306"}
      $MYSQL_Connection.Open()

      $MYSQL_Command            = New-Object MySql.Data.MySqlClient.MySqlCommand
      $MYSQL_DataAdapter        = New-Object MySql.Data.MySqlClient.MySqlDataAdapter
      $MYSQL_DataSet            = New-Object System.Data.DataSet 
      $MYSQL_Command.Connection = $MYSQL_Connection

      $AddressTable = 'id_address_delivery'
      Switch ($Combobox_Address.SelectedItem) {

         'Shipping Address' { $AddressTable = 'id_address_delivery' }
         'Invoice Address'  { $AddressTable = 'id_address_invoice'  }

      }

      $OrderID = [Math]::Abs($($Textbox_ID.Text -replace '^[^\d]*(\d+)[^\d]*$', '$1'))

      $MYSQL_Command.CommandText = "SELECT invoice_number AS Invoice,
                                           Company,
                                           CONCAT(firstname, `" `", lastname) AS Name, 
                                           CONCAT(address1, `", `", address2) AS Address,
                                           Postcode AS Zip,
                                           City,
                                           $($Textbox_TablesPrefix.text)orders.reference AS Reference
                                    FROM $($Textbox_TablesPrefix.text)address
                                       INNER JOIN $($Textbox_TablesPrefix.text)orders
                                          ON $($Textbox_TablesPrefix.text)address.id_address = $($Textbox_TablesPrefix.text)orders.$($AddressTable)
                                    WHERE $($Textbox_TablesPrefix.text)orders.id_order = $($OrderID)"

      $MYSQL_DataAdapter.SelectCommand = $MYSQL_Command

      if (-not($MYSQL_DataAdapter.Fill($MYSQL_DataSet, "data") -eq 1)) {
         throw "Invalid order ID ($($OrderID))."
      }

      if ([string]$($MYSQL_DataSet.tables[0][0].Company).Length -gt 0) { $Address = "$($MYSQL_DataSet.tables[0][0].Company.trim())`rAtt: " }
      $Address += "$($MYSQL_DataSet.tables[0][0].Name.trim())`r$($MYSQL_DataSet.tables[0][0].Address.trim() -replace '\,$', '')`r$($MYSQL_DataSet.tables[0][0].Zip.trim()) $($MYSQL_DataSet.tables[0][0].City.trim())"

      $Invoice   = $MYSQL_DataSet.tables[0][0].Invoice.tostring('000000')
      $Reference = $MYSQL_DataSet.tables[0][0].Reference.trim()

      $MYSQL_Connection.Close()

      #Load template and insert address and reference
      $Global:Filename = "$($Global:OutFolder)\$($Invoice).fodt"

      if (-not(Test-Path $Global:Outfolder -PathType Container)) { New-Item -Path $Global:Outfolder -ItemType Directory }

      if (Test-Path $Global:Filename -PathType Leaf) {
         throw "Label already exists for order with ID $($OrderID).`r`n`Please remove the old label:`r`n`r`n$($Global:Filename | Split-Path -leaf)"
      }

      (Get-Content -path "$($Global:Templatesfolder)\$($Combobox_Type.SelectedItem).fodt" -Raw -Encoding utf8) -replace '\[Adresse\]', ($Address -replace "`r", '<text:line-break/>') -replace '\[Reference\]', $Reference | Out-File $($Global:Filename) -Encoding utf8
      Start-Process -FilePath $Global:Filename -WindowStyle Normal -Wait

      $Label_Document.Text = $Global:Filename | Split-Path -leaf
      $Panel_Print.Enabled = $true

   }
   catch {

      Add-Type -AssemblyName PresentationFramework
      [System.Windows.MessageBox]::Show($_, 'Label builder', 'OK', 'Error') | Out-Null

   }

   $Panel_Generate.Enabled = $true
   $Panel_Database.Enabled = $true

   [System.Windows.Forms.Application]::DoEvents()

}

# Send completed label to selected printer
Function PrintLabel {

   $Panel_Generate.Enabled = $false
   $Panel_Database.Enabled = $false

   [System.Windows.Forms.Application]::DoEvents()

   try {

      $DefaultPrinter = Get-WmiObject -Query "SELECT Name FROM Win32_Printer WHERE Default=$true" | Select-Object -ExpandProperty Name

      (New-Object -ComObject WScript.Network).SetDefaultPrinter($Combobox_Printer.SelectedItem.ToString())
      $PrinterConfig = $("BeforePrinting-" + $Combobox_Printer.SelectedItem.ToString() -replace '\s', '_')
      if (Get-Command $PrinterConfig -errorAction SilentlyContinue) { $PrinterConfig | Invoke-Expression }
      Start-Sleep -Milliseconds 500

      Start-Process -FilePath $Global:Filename -Verb Print -WindowStyle Hidden -Wait
      Wait-Process -Name swriter -Timeout 30 -ErrorAction SilentlyContinue

      Start-Sleep -Milliseconds 500
      $PrinterConfig = $("AfterPrinting-" + $Combobox_Printer.SelectedItem.ToString() -replace '\s', '_')
      if (Get-Command $PrinterConfig -errorAction SilentlyContinue) { $PrinterConfig | Invoke-Expression }
      (New-Object -ComObject WScript.Network).SetDefaultPrinter($DefaultPrinter)

   }
   catch {

      Add-Type -AssemblyName PresentationFramework
      [System.Windows.MessageBox]::Show($_, 'Label builder', 'OK', 'Error') | Out-Null

   }

   $Panel_Generate.Enabled = $true
   $Panel_Database.Enabled = $true

   [System.Windows.Forms.Application]::DoEvents()

}

# Save database settings to registry
Function SaveSettings {

   try {

      $Base = ''
      foreach ($Item in $Global:RegistryKey.Split("\")) {
         $Base += "$($item)\"
         if (-not (Test-Path -Path $Base)) { New-Item -Path $Base }
      }

      Set-ItemProperty -Path $Global:RegistryKey -Name "Server"       -Value $Textbox_Server.Text
      Set-ItemProperty -Path $Global:RegistryKey -Name "Database"     -Value $Textbox_Database.Text
      Set-ItemProperty -Path $Global:RegistryKey -Name "TablesPrefix" -Value $Textbox_TablesPrefix.Text
      Set-ItemProperty -Path $Global:RegistryKey -Name "Username"     -Value $Textbox_Username.Text
      Set-ItemProperty -Path $Global:RegistryKey -Name "Password"     -Value $([Convert]::ToBase64String($(Get-CompressedByteArray -byteArray $([System.Text.Encoding]::UTF8.GetBytes($($Textbox_Password.Text | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString))))))

   }
   catch {

      Add-Type -AssemblyName PresentationFramework
      [System.Windows.MessageBox]::Show($_, 'Label builder', 'OK', 'Error') | Out-Null

   }

}

# Create a new form
$Form                    = New-Object system.Windows.Forms.Form

# Define the size, title and background color
$Form.ClientSize       = '340, 555'
$Form.text             = "Label Builder"
$Form.MinimizeBox      = $false
$Form.MaximizeBox      = $false
$Form.FormBorderStyle  = 1
$Form.Icon             = $Global:Icon

# Main panel
$Panel_Generate             = New-Object system.Windows.Forms.Panel
$Panel_Generate.Font        = 'Segoe UI,9'
$Panel_Generate.AutoSize    = $false
$Panel_Generate.Width       = 300
$Panel_Generate.Height      = 170
$Panel_Generate.BorderStyle = 'FixedSingle'
$Panel_Generate.Location    = New-Object System.Drawing.Point(20,20)

$Label_ID = New-Object system.Windows.Forms.Label
$Label_ID.text = "Order ID"
$Label_ID.AutoSize = $true
$Label_ID.Width = 100
$Label_ID.Height = 20
$Label_ID.Location = New-Object System.Drawing.Point(20,20)

$Label_Address = New-Object system.Windows.Forms.Label
$Label_Address.text = "Address"
$Label_Address.AutoSize = $true
$Label_Address.Width = 100
$Label_Address.Height = 20
$Label_Address.Location = New-Object System.Drawing.Point(20,50)

$Label_Type = New-Object system.Windows.Forms.Label
$Label_Type.text = "Type"
$Label_Type.AutoSize = $true
$Label_Type.Width = 100
$Label_Type.Height = 20
$Label_Type.Location = New-Object System.Drawing.Point(20,80)

$Textbox_ID = New-Object system.Windows.Forms.Textbox
$Textbox_ID.AutoSize = $true
$Textbox_ID.Width = 150
$Textbox_ID.Height = 20
$Textbox_ID.Location = New-Object System.Drawing.Point(120,20)

$Combobox_Address = New-Object system.Windows.Forms.Combobox
$Combobox_Address.AutoSize = $true
$Combobox_Address.Width = 150
$Combobox_Address.Height = 20
$Combobox_Address.Location = New-Object System.Drawing.Point(120,50)

$Combobox_Type = New-Object system.Windows.Forms.Combobox
$Combobox_Type.AutoSize = $true
$Combobox_Type.Width = 150
$Combobox_Type.Height = 20
$Combobox_Type.Location = New-Object System.Drawing.Point(120,80)

$Button_GenerateLetter           = New-Object system.Windows.Forms.button
$Button_GenerateLetter.Text      = 'Generate Label'
$Button_GenerateLetter.Width     = 150
$Button_GenerateLetter.Height    = 20
$Button_GenerateLetter.Location  = New-Object System.Drawing.Point(120,130)
$Button_GenerateLetter.Add_Click({GenerateLabel})

# Printing panel
$Panel_Print             = New-Object system.Windows.Forms.Panel
$Panel_Print.Font        = 'Segoe UI,9'
$Panel_Print.AutoSize    = $false
$Panel_Print.Width       = 300
$Panel_Print.Height      = 90
$Panel_Print.BorderStyle = 'FixedSingle'
$Panel_Print.Enabled     = $false
$Panel_Print.Location    = New-Object System.Drawing.Point(20,210)

$Label_SelectPrinter = New-Object system.Windows.Forms.Label
$Label_SelectPrinter.text = "Select printer"
$Label_SelectPrinter.AutoSize = $true
$Label_SelectPrinter.Width = 100
$Label_SelectPrinter.Height = 20
$Label_SelectPrinter.Location = New-Object System.Drawing.Point(20,20)

$Label_Document = New-Object system.Windows.Forms.Label
$Label_Document.AutoSize = $true
$Label_Document.Width = 100
$Label_Document.Height = 20
$Label_Document.ForeColor = 'Gray'
$Label_Document.Location = New-Object System.Drawing.Point(20,50)

$Combobox_Printer = New-Object system.Windows.Forms.Combobox
$Combobox_Printer.AutoSize = $true
$Combobox_Printer.Width = 150
$Combobox_Printer.Height = 20
$Combobox_Printer.Location = New-Object System.Drawing.Point(120,20)

$Button_Print           = New-Object system.Windows.Forms.button
$Button_Print.Text      = 'Print label'
$Button_Print.Width     = 150
$Button_Print.Height    = 20
$Button_Print.Location  = New-Object System.Drawing.Point(120,50)
$Button_Print.Add_Click({PrintLabel})

# Database panel
$Panel_Database             = New-Object system.Windows.Forms.Panel
$Panel_Database.Font        = 'Segoe UI,9'
$Panel_Database.AutoSize    = $false
$Panel_Database.Width       = 300
$Panel_Database.Height      = 215
$Panel_Database.BorderStyle = 'FixedSingle'
$Panel_Database.Location    = New-Object System.Drawing.Point(20,320)

$Label_Server = New-Object system.Windows.Forms.Label
$Label_Server.text = "Database server:"
$Label_Server.AutoSize = $true
$Label_Server.Width = 100
$Label_Server.Height = 20
$Label_Server.Location = New-Object System.Drawing.Point(20,20)

$Label_Database = New-Object system.Windows.Forms.Label
$Label_Database.text = "Database:"
$Label_Database.AutoSize = $true
$Label_Database.Width = 100
$Label_Database.Height = 20
$Label_Database.Location = New-Object System.Drawing.Point(20,50)

$Label_TablesPrefix = New-Object system.Windows.Forms.Label
$Label_TablesPrefix.text = "Tables prefix:"
$Label_TablesPrefix.AutoSize = $true
$Label_TablesPrefix.Width = 100
$Label_TablesPrefix.Height = 20
$Label_TablesPrefix.Location = New-Object System.Drawing.Point(20,80)

$Label_Username = New-Object system.Windows.Forms.Label
$Label_Username.text = "Username:"
$Label_Username.AutoSize = $true
$Label_Username.Width = 100
$Label_Username.Height = 20
$Label_Username.Location = New-Object System.Drawing.Point(20,110)

$Label_Password = New-Object system.Windows.Forms.Label
$Label_Password.text = "Password:"
$Label_Password.AutoSize = $true
$Label_Password.Width = 100
$Label_Password.Height = 20
$Label_Password.Location = New-Object System.Drawing.Point(20,140)

$Textbox_Server = New-Object system.Windows.Forms.Textbox
$Textbox_Server.AutoSize = $true
$Textbox_Server.Width = 150
$Textbox_Server.Height = 20
$Textbox_Server.Location = New-Object System.Drawing.Point(120,20)

$Textbox_Database = New-Object system.Windows.Forms.Textbox
$Textbox_Database.AutoSize = $true
$Textbox_Database.Width = 150
$Textbox_Database.Height = 20
$Textbox_Database.Location = New-Object System.Drawing.Point(120,50)

$Textbox_TablesPrefix = New-Object system.Windows.Forms.Textbox
$Textbox_TablesPrefix.AutoSize = $true
$Textbox_TablesPrefix.Width = 150
$Textbox_TablesPrefix.Height = 20
$Textbox_TablesPrefix.Location = New-Object System.Drawing.Point(120,80)

$Textbox_Username = New-Object system.Windows.Forms.Textbox
$Textbox_Username.AutoSize = $true
$Textbox_Username.Width = 150
$Textbox_Username.Height = 20
$Textbox_Username.Location = New-Object System.Drawing.Point(120,110)

$Textbox_Password = New-Object system.Windows.Forms.Textbox
$Textbox_Password.AutoSize = $true
$Textbox_Password.Width = 150
$Textbox_Password.Height = 20
$Textbox_Password.PasswordChar = '*'
$Textbox_Password.Location = New-Object System.Drawing.Point(120,140)

$Button_SaveSettings           = New-Object system.Windows.Forms.button
$Button_SaveSettings.Text      = 'Save Settings'
$Button_SaveSettings.Width     = 150
$Button_SaveSettings.Height    = 20
$Button_SaveSettings.Location  = New-Object System.Drawing.Point(120,170)
$Button_SaveSettings.Add_Click({SaveSettings})

# Attaching form elements to form
$Form.controls.AddRange(@($Panel_Generate, $Panel_Database, $Panel_Print))

$Panel_Generate.controls.AddRange(@($Label_ID, $Label_Address, $Label_Type,
                                    $Textbox_ID, $Combobox_Address, $Combobox_Type,
                                    $Button_GenerateLetter))

$Panel_Print.controls.AddRange(@($Label_SelectPrinter, $Label_Document,
                                 $Combobox_Printer,
                                 $Button_Print))

$Panel_Database.controls.AddRange(@($Label_Server, $Label_Database, $Label_TablesPrefix, $Label_Username, $Label_Password
                                    $Textbox_Server, $Textbox_Database, $Textbox_TablesPrefix, $Textbox_Username, $Textbox_Password,
                                    $Button_SaveSettings))

# Populating comboboxes
$Combobox_Address.Items.Add('Shipping Address') | Out-Null
$Combobox_Address.Items.Add('Invoice Address') | Out-Null
$Combobox_Address.SelectedIndex = 0

Get-ChildItem -Path $Global:Templatesfolder -Filter '*.fodt' | ForEach-Object { $Combobox_Type.Items.Add($($_.BaseName -replace '\d\s\-\s(.+)', '$1')) | Out-Null }
$Combobox_Type.SelectedIndex = 0

$DefaultPrinter = Get-WmiObject -Query "SELECT Name FROM Win32_Printer WHERE Default=$true" | Select-Object -ExpandProperty Name
Get-Printer | ForEach-Object {

   $Combobox_Printer.Items.Add($_.Name) | Out-Null
   if ($_.Name -eq $DefaultPrinter) { $Combobox_Printer.SelectedIndex = $($Combobox_Printer.Items.Count - 1) }
   
}

if (Test-Path -Path $Global:RegistryKey) {

   $Settings = Get-ItemProperty -Path "$($Global:RegistryKey)"

   if ([bool]($Settings.PSobject.Properties.name -match "Server")) {
      $Textbox_Server.Text = $Settings.Server
   }

   if ([bool]($Settings.PSobject.Properties.name -match "Database")) {
      $Textbox_Database.Text = $Settings.Database
   }

   if ([bool]($Settings.PSobject.Properties.name -match "TablesPrefix")) {
      $Textbox_TablesPrefix.Text = $Settings.TablesPrefix
   }

   if ([bool]($Settings.PSobject.Properties.name -match "Username")) {
      $Textbox_Username.Text = $Settings.Username
   }

   if ([bool]($Settings.PSobject.Properties.name -match "Password")) {
      $Textbox_Password.Text = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($([System.Text.Encoding]::UTF8.GetString( $(Get-DecompressedByteArray -byteArray $([byte[]][Convert]::FromBase64String($Settings.Password)))) | ConvertTo-SecureString)))
   }

}

# Display the form
[void]$Form.ShowDialog()