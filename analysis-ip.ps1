$xlCellTypeLastCell = 11
$startRow = 2
$excel = New-Object -Com Excel.Application
$wb = $excel.Workbooks.Open($args[0])
$sh = $wb.Sheets.Item(1)
$endRow = $sh.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$users = @()
$web_client = new-object system.net.webclient
function Create-Excel {
    [cmdletbinding()]
    param (
        [array]
        $Data
    )
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    # Creacion fichero excel
    $libro = $excel.Workbooks.add()
    # Creacion hoja excel
    $hoja = $libro.WorkSheets.Add()
    $hoja.Name = "Registros ISP"
    $hoja.Select()
    $hoja.Cells.Item(2,1) = "Time"
    $hoja.Cells.Item(2,2) = "Usuario"
    $hoja.Cells.Item(2,3) = "IP"
    $hoja.Cells.Item(2,4) = "Location"
    $hoja.Cells.Item(2,5) = "ISP"
    #Escribimos hoja
    $Fila = 3
    for ( $i=0; $i -lt $Data.Count; $i++ )
    {
        $hoja.Cells.Item($Fila,1) = $Data[$i].Time
        $hoja.Cells.Item($Fila,2) = $Data[$i].User
        $hoja.Cells.Item($Fila,3) = $Data[$i].Ip
        $hoja.Cells.Item($Fila,4) = $Data[$i].Location
        $hoja.Cells.Item($Fila,5) = $Data[$i].Isp
        $Fila++
    }
    #Grabamos y abrimos
    #$name_file = Get-Date -Format dd_MM_yyyy
    $libro.SaveAs($args[1])
    $excel.Workbooks.Close()
}


for ($i=$startRow; $i -le $endRow; $i++)
{
    $time = [DateTime]::FromOADate($sh.Cells.Item($startRow,2).Value2) | Get-Date
    $ip = $sh.Cells.Item($startRow,5).Value2
    $user = $sh.Cells.Item($startRow,4).Value2
    $location = $sh.Cells.Item($startRow,8).Value2
    $response =$web_client.DownloadString("https://api.iplocation.net/?ip="+$ip) | ConvertFrom-Json
    $object = New-Object PSObject -Property @{ Time=$time;User=$user;Ip = $ip;Location=$location;Isp=$response.isp; }
    $users +=($object)
    $startRow++
}

Create-Excel -Data $users
$excel.Workbooks.Close()