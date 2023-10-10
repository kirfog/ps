$SearchBase = "ou=ru,ou=comps, dc=lan, dc=local"
$result = "I:\Inventory\invent.xlsx"
if (Test-Path variable:global:cred){
    $cred = Get-Content variable:global:cred
}else {
    $cred = Get-Credential
}
$Excel = New-Object -Com Excel.Application
$ex = $false
if (Test-Path -Path $result)  {
    $Book = $Excel.Workbooks.Open($result)
    $Sheet=$Book.WorkSheets.Item(1)
    $ex = $true
}
else {
    $Book=$Excel.Workbooks.Add()
    $Book.WorkSheets.item(1).Name = "Computers"
    $Sheet=$Book.WorkSheets.Item(1)
    $cols = [ordered]@{"Name" = 20; "Updated" = 15; "Description" = 20; "Location" = 20; "Serial number" = 20;
    "Last user" = 20; "Operating system" = 25; "CPU" = 25; "Memory" = 20; "Motherboard" = 15; "Hard drive" = 25;
    "Video card" = 25; "Network" = 35; "Cisco" = 35; "Sophos" = 35; "TeamViewer" = 35; "Dell" = 35
    }
    $coln = 0
    foreach ($col in $cols.keys) {
        $coln = $coln + 1
        $Sheet.Cells.Item(1,$coln) = "$col"
        $Sheet.Columns.Item($coln).columnWidth = $cols[$col]
    }
    $Sheet.UsedRange.Interior.ColorIndex = 5
    $Sheet.UsedRange.Font.ColorIndex = 20
    $Sheet.UsedRange.Font.Bold = $True
    $sheet.Rows.Item(1).HorizontalAlignment = 3
}

$Rowcount=$Sheet.UsedRange.Rows.Count
$Row=$Rowcount
$comps = Get-ADcomputer -Credential $cred -Filter * -SearchBase $SearchBase | Where-Object { $_.DistinguishedName -notlike '*OU=lost*'} | Sort-Object name | Select-Object -ExpandProperty name
function anerror {
    $Exc = $error[0].ToString() -replace "\<[^\>]+\>",""
    $Dot = $Exc.IndexOf(".")
    $Exc = $Exc.Substring(0, $Dot)
    Write-Output $Exc
    $Sheet.Cells.Item($Row,2) = $Exc
}

ForEach ($compname in $comps ) {
    $Row = $Row + 1
    for($i=2; $i -le $Rowcount; $i++){
        if ($Sheet.cells.Item($i, 1).text -eq $compname){
            $Row = $i
    }}

    Write-Output "`n---`n"
    Write-Output "Gathering information about $compname ..."

    if (Test-WSMan -ComputerName $compname -ErrorAction SilentlyContinue) {
        $cimses = New-Cimsession -Credential $cred -computername $compname
        if ($cimses){
            #Serial number
            $sn = Get-CimInstance -CimSession $cimses win32_SystemEnclosure | select-object -ExpandProperty serialnumber
            $Sheet.Cells.Item($Row,5) = $sn
            #Last user
            $lastuser = Get-CimInstance -CimSession $cimses Win32_NetworkLoginProfile | Sort-Object LastLogon -Descending | select-object -ExpandProperty Name -first 1
            $Sheet.Cells.Item($Row,6) = $lastuser
            #System
            $sys = Get-CimInstance -CimSession $cimses Win32_OperatingSystem
            $Sheet.Cells.Item($Row,7) = $sys.caption + "`n" + $sys.csdversion
            #CPU
            $cpu = Get-CimInstance -CimSession $cimses Win32_Processor
            $Sheet.Cells.Item($Row,8) = $cpu.name+"`n" + $cpu.caption + "`n" + $cpu.SocketDesignation
            #Memory
            $ram = Get-CimInstance -CimSession $cimses Win32_Physicalmemory
            foreach ($dimm in $ram){
                $mem = $mem + $dimm.capacity
                $dimms = $dimms + 1
                $parts = $parts + ($dimm.capacity / 1Gb).tostring("F00") + "GB " + $dimm.speed +"Mhz" + "`n" + $dimm.PartNumber.ToString() + "`n"
            }
            $Sheet.Cells.Item($Row,9) = ($mem / 1Gb).tostring("F00") + "GB`n" + $dimms + " DIMMs" + "`n" + $parts
            #Motherboard
            $mb = Get-CimInstance -CimSession $cimses Win32_BaseBoard
            $Sheet.Cells.Item($Row,10) = $mb.Manufacturer + "`n" + $mb.Product
            #Disk
            foreach ($hard in Get-CimInstance -CimSession $cimses win32_diskdrive){
                if ($hard.MediaType.ToLower().StartsWith("fixed")){
                    $disk=$disk+(($hard.size)/1Gb).tostring("F00") + "GB - " + $hard.model +"`n"
                }
            }
            $Sheet.Cells.Item($Row,11) = $disk.TrimEnd("`n")
            #Video
            foreach ($card in Get-CimInstance -CimSession $cimses Win32_videoController){
                if ($card.AdapterRAM -gt 0){
                    $video = $video + $card.name + "`n" + ($card.AdapterRAM/1Mb).tostring("F00") + "MB`n"}
                }
            $Sheet.Cells.Item($Row,12) = $video.TrimEnd("`n")
            #Network
            foreach ($card in Get-CimInstance -CimSession $cimses Win32_NetworkAdapter -Filter "NetConnectionStatus = 2"){
                $net=$net+$card.name + " " + $card.macaddress + "`n"
            }
            $Sheet.Cells.Item($Row,13) = $net.TrimEnd("`n")
            #Software
            $Soft = Get-CimInstance -CimSession $cimses -Class Win32_Product
            $Soft | Where-Object vendor -like Cisco* | Select-Object Name, Version | ForEach-Object {
                $Soft_cisco = $Soft_cisco + $_.name + " " + $_.version + "`n"
            }
            $Sheet.Cells.Item($Row,14) = $Soft_cisco
            $Soft | Where-Object vendor -like Sophos* | Select-Object Name, Version | ForEach-Object {
                $Soft_sophos = $Soft_sophos + $_.name + " " + $_.version + "`n"
            }
            $Sheet.Cells.Item($Row,15) = $Soft_sophos
            $Soft | Where-Object vendor -like TeamViewer* | Select-Object Name, Version | ForEach-Object {
                $Soft_tw = $Soft_tw + $_.name + " " + $_.version + "`n"
            }
            $Sheet.Cells.Item($Row,16) = $Soft_tw
            $Soft | Where-Object vendor -like Dell* | Select-Object Name, Version | ForEach-Object {
                $Soft_dell = $Soft_dell + $_.name + " " + $_.version + "`n"
            }
            $Sheet.Cells.Item($Row,17) = $Soft_dell

            #Updated
            if ($cpu) { $Sheet.Cells.Item($Row,2) = Get-Date }
            $cimses.close()
        }
        else { anerror
        }
    }
    else { anerror
    }

    #Name
    $Sheet.Cells.Item($Row,1) = $compname
    #description
    $des = Get-ADComputer $compname -Properties "Description" | Select-Object -ExpandProperty Description
    $Sheet.Cells.Item($Row,3) = $des
    #location
    $loc = Get-ADComputer $compname -Properties "Location" | Select-Object -ExpandProperty Location
    $Sheet.Cells.Item($Row,4) = $loc

    $parts, $Soft, $Soft_dell, $Soft_tw, $Soft_sophos, $Soft_cisco, $net, $video, $disk, $mb, $ram, $cpu, $sys, $lastuser, $sn, $loc, $des, $tableline, $Updated = ""
    $mem, $dimms = 0

    Write-Output "`n...`n"
    
}

if (-Not $ex)  { $Sheet.Rows.Item(1).AutoFilter() }
$Sheet.UsedRange.WrapText = 1
$Sheet.UsedRange.EntireRow.AutoFit()
$Sheet.UsedRange.Cells.borders.TintAndShade = 1
$Sheet.UsedRange.VerticalAlignment = 2

$excel.ActiveWindow.SplitRow = 1
$excel.ActiveWindow.SplitColumn = 1
$excel.ActiveWindow.FreezePanes = $true
$excel.DisplayAlerts = $false
#$Excel.visible=$True
$Book.SaveAs($result)
$excel.Quit()