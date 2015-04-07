# add old driver name here
$oldDrvr = "KONICA MINOLTA 423SeriesPS"

# add new driver name here
$newDrvr = "KONICA MINOLTA 423SeriesPS-8"


# finds all printers
$printers = gwmi win32_printer | select Name,DriverName
 

foreach($printer in $printers){
        $name = $printer.name
        $dname = $printer.DriverName
        if($dname -eq $oldDrvr){
            & rundll32 printui.dll PrintUIEntry /Xs /n $name DriverName $newDrvr
            Write-Host $name " - Processed" 
            }

   
}
