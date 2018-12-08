Write-Host "Importamos modulo VMM" -ForegroundColor Green
Import-Module "C:\Program Files\Microsoft System Center 2012 R2\Virtual Machine Manager\bin\psModules\virtualmachinemanager\virtualmachinemanager.psd1"

$VMMBS= "VMMBSV02P.adgbs.com"
$VMMSF= "VMMISV02P.adgbs.com"
$VMMBSVDI = "VMMBSV01P.adgbs.com"
$VMMSFVDI = "VMMISV01P.adgbs.com"
$HyperV = "hypbsf150p"

Get-VMMServer $VMMBSVDI  | Get-VMHost -ComputerName $HyperV |Get-VM |Select-Object Name


hypbsf005p
hypbsf006p
hypbsf150p

hypbsf381p
hypbsf507p
hypbsf508p
hypbsf519p
hypbsf520p
hypbsf521p
hypbsf524p
hypbsf525p
hypbsf526p
hypbsf527p
hypbsf528p
hypbsf529p
hypbsf532p
hypbsf6000p
hypisf102p
hypisf104p
hypisf336p
hypisf337p
hypisf338p
hypisf374p
hypisf500p
hypisf501p
hypisf502p
hypisf503p
hypisf504p
hypisf506p
hypisf507p
hypisf510p
hypisf511p
hypisf512p
hypisf513p
hypisf514p
hypisf515p
hypisf516p
hypisf517p
hypisf518p
hypisf519p
hypisf523p
hypisf525p
hypisf526p
hypisf527p
hypisf528p
hypisf529p
hypisf531p
hypbsf540p
hypbsf6101p
hypisf541p
hypisf6101p
hypbsf541p
hypisf1003p
hypisf540p