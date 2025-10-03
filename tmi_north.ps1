# Queries all Windows users accounts

$users = Get-ChildItem -Path "C:\Users"

# Loops through users and deletes Karmak.ProfitMaster.exe.Config file from each users \AppData\Roaming\Karmak Inc\ProfitMaster\ directory

$users | ForEach-Object {

            Remove-Item -Path "C:\Users\$($_.Name)\AppData\Roaming\Karmak Inc\ProfitMaster\Karmak.ProfitMaster.exe.Config" -Force -ErrorAction SilentlyContinue

}

#Reads Fusion Client config file

$basefolder = "C:\Program Files (x86)\KarmakInc\KarmakBusinessSystem"

$newip = "App-27148.karmakxcelerate.com" #New IP address of Fusion Server.

$port = "8280" #port for application.

$branchip = "10.0.0.6" #IP of the branch server

$kusport = "65002" #port for appstart.config for Application server clients 65001. For branch servers 65002

$envname = "TruckMaintenance" #name of environment

 

$filter = @("barcodesetup.exe","test.txt")

$configs = Get-ChildItem $basefolder -Name | Where-Object { $_ -inotin $filter }

 

$configsToUpdate = @()

if ($configs.Count -gt 1) {

    Write-Host "Found multiple Karmak Client installations. Which client would you like to update?" -ForegroundColor Yellow

    $fullOutput = @()

    for ($i = 0; $i -lt $configs.Count; $i++) {

        $config = $configs[$i]

 

        $output = New-Object psobject -Property @{

            Number     = $i + 1

            FolderName = $config

        }

 

        $fullOutput += $output | Select-Object Number, FolderName

    }

 

    $fullOutput | Format-Table

    do {

        $selection = $(Write-Host "Your selection:" -ForegroundColor Cyan; Read-Host)

    } while ($selection -NotIn $fullOutput.Number)

 

    $configsToUpdate = $configs[$selection - 1]

    $configsToUpdate

}

else {

    Write-Host "Found one, continuing with it."

 

    $configsToUpdate = $configs

}

 

If (Test-Path ( "$basefolder\$configsToUpdate\Karmak.ProfitMaster.exe.config" ))

{

    $configxml = @()

    [xml]$configxml = (Select-Xml -Path "$basefolder\$configsToUpdate\Karmak.ProfitMaster.exe.config" -XPath /).Node

    $configxml.configuration.remoting_connection_manager.connections.default_connection_id = $envname

    $configxml.configuration.remoting_connection_manager.connections.connection.id = $envname

    $configxml.configuration.remoting_connection_manager.connections.connection.ip_address = $newip

    $configxml.configuration.remoting_connection_manager.connections.connection.port = $port

                [xml]$configxml1 = (Select-Xml -Path "$basefolder\$configsToUpdate\Appstart.config" -XPath /).Node

                $currentip = $configxml1.configuration.update_server.ipaddress

                $configxml1.configuration.update_server.ipaddress = $branchip

                $configxml1.configuration.update_server.port = $kusport

                #Settings object will instruct how the xml elements are written to the file

                $settings1 = New-Object System.Xml.XmlWriterSettings

                $settings1.Indent = $true

                #NewLineChars will affect all newlines

                $settings1.NewLineChars = "`r`n"

                #Set an optional encoding, UTF-8 is the most used (without BOM)

                $settings1.Encoding = New-Object System.Text.UTF8Encoding($false)

                $y = [System.Xml.XmlWriter]::Create("$basefolder\$configsToUpdate\appstart.config", $settings1)

                Try

                {

                                $configxml1.Save($y)

                }

                Finally

                {

                                Write-Output "Config file is transformed!"

                }

                #Settings object will instruct how the xml elements are written to the file

                $settings = New-Object System.Xml.XmlWriterSettings

                $settings.Indent = $true

                #NewLineChars will affect all newlines

                $settings.NewLineChars = "`r`n"

                #Set an optional encoding, UTF-8 is the most used (without BOM)

                $settings.Encoding = New-Object System.Text.UTF8Encoding($false)

                $w = [System.Xml.XmlWriter]::Create("$basefolder\$configsToUpdate\Karmak.ProfitMaster.exe.config", $settings)

                Try

                {

                                $configxml.Save($w)

                }

                Finally

                {

                                Write-Output "Config file is transformed!"

                }

}
