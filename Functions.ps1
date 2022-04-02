<#

  In this file, you may add functions to be executed
  before and after printing on specific printers.

  The functions must be named like this:

  Before-[printer name with undercore instead of whitespace]
  After-[printer name with undercore instead of whitespace]

#>

function Before-Brother_QL-1110NWB {

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

function After-Brother_QL-1110NWB {

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
