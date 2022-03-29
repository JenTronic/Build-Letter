Add-Type -AssemblyName System.Windows.Forms

# Settings
$Global:RegistryKey     = 'HKCU:\SOFTWARE\JenTronic\Label Builder'
$Global:Templatesfolder = "$PSScriptRoot\Templates"
$Global:Outfolder       = "$PSScriptRoot\Labels"

# Global vars
$Global:Filename        = ""

# Load required assembly
[System.Reflection.Assembly]::LoadWithPartialName("MySql.Data") | Out-Null

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
      $tmp = $output.ToArray()
      Write-Output $tmp
   }

}

function Get-DecompressedByteArray {

   [CmdletBinding()]
   Param (
      [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
      [byte[]] $byteArray = $(Throw("-byteArray is required"))
   )
   Process {
      
      $input = New-Object System.IO.MemoryStream( , $byteArray )
      $output = New-Object System.IO.MemoryStream
      $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)
      $gzipStream.CopyTo( $output )
      $gzipStream.Close()
      $input.Close()
      [byte[]] $byteOutArray = $output.ToArray()
      Write-Output $byteOutArray
   
   }

}

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

      $MYSQL_Command.CommandText = "SELECT invoice_number AS Invoice, Company, CONCAT(firstname, `" `", lastname) AS Name, CONCAT(address1, `", `", address2) AS Address, Postcode AS Zip, City, $($Textbox_TablesPrefix.text)orders.reference AS Reference FROM $($Textbox_TablesPrefix.text)address INNER JOIN $($Textbox_TablesPrefix.text)orders ON $($Textbox_TablesPrefix.text)address.id_address = $($Textbox_TablesPrefix.text)orders.$($AddressTable) WHERE $($Textbox_TablesPrefix.text)orders.id_order = abs($($Textbox_ID.Text -replace '^[^\d]*(\d+)[^\d]*$', '$1'))"
      $MYSQL_DataAdapter.SelectCommand = $MYSQL_Command

      if (-not($MYSQL_DataAdapter.Fill($MYSQL_DataSet, "data") -eq 1)) {
         throw "Invalid order ID."
      }

      if ([string]$($MYSQL_DataSet.tables[0][0].Company).Length -gt 0) { $Address = "$($MYSQL_DataSet.tables[0][0].Company.trim())`rAtt: " }
      $Address += "$($MYSQL_DataSet.tables[0][0].Name.trim())`r$($MYSQL_DataSet.tables[0][0].Address.trim() -replace '\,$', '')`r$($MYSQL_DataSet.tables[0][0].Zip.trim()) $($MYSQL_DataSet.tables[0][0].City.trim())"

      $Invoice   = $MYSQL_DataSet.tables[0][0].Invoice.tostring('000000')
      $Reference = $MYSQL_DataSet.tables[0][0].Reference.trim()

      $MYSQL_Connection.Close()

      #Load template and insert address
      $Global:Filename = "$($Global:OutFolder)\$($Invoice).fodt"

      if (-not(Test-Path $Global:Outfolder -PathType Container)) { New-Item -Path $Global:Outfolder -ItemType Directory }

      if (Test-Path $Global:Filename -PathType Leaf) {

         $msgBody   = "Label already exists. Please remove the old label:`r`n`r`n$($Global:Filename | Split-Path -leaf)"
         $msgTitle  = "Label builder"
         $msgButton = 'OK'
         $msgImage  = 'Error'
   
         Add-Type -AssemblyName PresentationFramework
         [System.Windows.MessageBox]::Show($msgBody, $msgTitle, $msgButton, $msgImage) | Out-Null

         throw

      }

      (Get-Content -path "$($Global:Templatesfolder)\$($Combobox_Type.SelectedItem).fodt" -Raw -Encoding utf8) -replace '\[Adresse\]', ($Address -replace "`r", '<text:line-break/>') -replace '\[Reference\]', $Reference | Out-File $($Global:Filename) -Encoding utf8
      Start-Process -FilePath $Global:Filename -WindowStyle Normal -Wait

      $Label_Document.Text = $Global:Filename | Split-Path -leaf
      $Panel_Print.Enabled = $true

   }
   catch {

      write-host $_

   }

   $Panel_Generate.Enabled = $true
   $Panel_Database.Enabled = $true

   [System.Windows.Forms.Application]::DoEvents()

}

Function PrintLabel {

   $Panel_Generate.Enabled = $false
   $Panel_Database.Enabled = $false

   [System.Windows.Forms.Application]::DoEvents()

   try {

      $DefaultPrinter = Get-WmiObject -Query "SELECT Name FROM Win32_Printer WHERE Default=$true" | Select-Object -ExpandProperty Name

      (New-Object -ComObject WScript.Network).SetDefaultPrinter($Combobox_Printer.SelectedItem.ToString())
      $PrinterConfig = $("Set-" + $Combobox_Printer.SelectedItem.ToString() -replace '\s', '_')
      if (Get-Command $PrinterConfig -errorAction SilentlyContinue) { $PrinterConfig | Invoke-Expression }
      Start-Sleep -Seconds 1

      Start-Process -FilePath $Global:Filename -Verb Print -WindowStyle Hidden -Wait
      Wait-Process -Name swriter -Timeout 30 -ErrorAction SilentlyContinue

      Start-Sleep -Seconds 1
      $PrinterConfig = $("Reset-" + $Combobox_Printer.SelectedItem.ToString() -replace '\s', '_')
      if (Get-Command $PrinterConfig -errorAction SilentlyContinue) { $PrinterConfig | Invoke-Expression }
      (New-Object -ComObject WScript.Network).SetDefaultPrinter($DefaultPrinter)

   }
   catch {

      write-host $_

   }

   $Panel_Generate.Enabled = $true
   $Panel_Database.Enabled = $true

   [System.Windows.Forms.Application]::DoEvents()

}


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

      write-host $_

   }

}

function Set-Brother_QL-1110NWB {

   $Hex = @"
   42,00,72,00,6f,00,74,00,68,00,65,00,72,00,20,00,51,00,\
   4c,00,2d,00,31,00,31,00,31,00,30,00,4e,00,57,00,42,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,01,04,00,05,\
   dc,00,52,01,0f,65,01,00,02,00,81,01,6b,06,f8,03,64,00,01,00,00,00,2c,01,01,\
   00,01,00,2c,01,03,00,00,00,31,00,30,00,33,00,6d,00,6d,00,20,00,78,00,20,00,\
   31,00,36,00,34,00,6d,00,6d,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,1e,00,00,00,ef,00,81,01,00,00,3b,00,01,00,00,00,1e,00,95,07,00,00,01,\
   00,00,00,42,52,50,54,00,00,00,00,00,00,00,00,6b,81,04,00,07,e5,0b,00,50,52,\
   49,56,a0,30,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,18,00,00,00,00,00,10,27,10,27,10,27,00,00,\
   10,27,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,0d,00,18,00,00,00,00,00,03,00,00,95,07,00,00,6b,06,00,00,6b,06,00,00,30,\
   75,00,00,fe,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,01,00,00,01,3b,00,\
   00,00,00,00,01,00,01,01,01,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,6b,81,04,00,07,e5,0b,00,00,00,00,00,00,00,12,01,00,0f,01,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00
"@
   
   $Data = [byte[]]$($($Hex -replace "\\|\s",'').Split(',') | % { "0x$_"})
   
   New-ItemProperty -Path HKCU:Printers\DevModes2 -Name "Brother QL-1110NWB" -PropertyType Binary -Value $Data -Force | Out-Null
   New-ItemProperty -Path HKCU:Printers\DevModePerUser -Name "Brother QL-1110NWB" -PropertyType Binary -Value $Data -Force | Out-Null
   
}

function Reset-Brother_QL-1110NWB {

   $Hex = @"
   42,00,72,00,6f,00,74,00,68,00,65,00,72,00,20,00,51,00,\
   4c,00,2d,00,31,00,31,00,31,00,30,00,4e,00,57,00,42,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,01,04,00,05,\
   dc,00,52,01,0f,65,01,00,02,00,4d,01,68,06,e6,00,64,00,01,00,00,00,2c,01,01,\
   00,01,00,2c,01,03,00,00,00,4a,00,65,00,6e,00,54,00,72,00,6f,00,6e,00,69,00,\
   63,00,00,00,6d,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,1e,00,00,00,ef,00,04,01,00,00,23,00,01,00,00,00,1e,00,91,07,00,00,01,\
   00,00,00,42,52,50,54,00,00,00,00,00,00,00,00,6b,81,04,00,07,e5,0b,00,50,52,\
   49,56,a0,30,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,18,00,00,00,00,00,10,27,10,27,10,27,00,00,\
   10,27,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,18,00,02,00,11,00,00,00,03,00,00,91,07,00,00,68,06,00,00,68,06,00,00,30,\
   75,00,00,fe,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,01,00,00,01,23,00,\
   00,00,00,00,01,00,01,01,01,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
   00,00,00,00,00,6b,81,04,00,07,e5,0b,00,00,00,00,00,00,00,12,01,00,0f,01,00,\
   00,00,00,00,00,00,00,00,00,00,00,00,00,00,00
"@
   
   $Data = [byte[]]$($($Hex -replace "\\|\s",'').Split(',') | % { "0x$_"})
   
   New-ItemProperty -Path HKCU:Printers\DevModes2 -Name "Brother QL-1110NWB" -PropertyType Binary -Value $Data -Force | Out-Null
   New-ItemProperty -Path HKCU:Printers\DevModePerUser -Name "Brother QL-1110NWB" -PropertyType Binary -Value $Data -Force | Out-Null
   
}

# Create a new form
$Form                    = New-Object system.Windows.Forms.Form

# Define the size, title and background color
$Form.ClientSize       = '340, 555'
$Form.text             = "Label Builder"
$Form.MinimizeBox      = $false
$Form.MaximizeBox      = $false
$Form.FormBorderStyle  = 1
$Form.Icon             = "$PSScriptRoot\JenTronic.ico"

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

      $securepwd = [System.Text.Encoding]::UTF8.GetString( $(Get-DecompressedByteArray -byteArray $([byte[]][Convert]::FromBase64String($Settings.Password)))) | ConvertTo-SecureString

      $Marshal = [System.Runtime.InteropServices.Marshal]
      $BinString = $Marshal::SecureStringToBSTR($securepwd)
      $Password = $Marshal::PtrToStringAuto($BinString)

      $Textbox_Password.Text = $Password

      write-host $Password

      $Marshal::ZeroFreeBSTR($BinString)

   }

}

# Display the form
[void]$Form.ShowDialog()