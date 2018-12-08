$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path



###################################################################################





echo "##################################" > "$scriptDir\Final_Report.txt"
echo "############## APPS ##############" >> "$scriptDir\Final_Report.txt"
echo "########## XAVBSV0009P ###########" >> "$scriptDir\Final_Report.txt"
$Hotfixes_Real = (get-wmiobject -class win32_quickfixengineering -computer "XAVBSV0009P").HotFixID


$Hotfixes_APP = Get-Content "$scriptDir\KB_APP.txt"

foreach ($fix in $Hotfixes_APP){

if ($Hotfixes_Real -contains $fix){

        write-host "YA EXISTE $fix"
        #echo "YA EXISTE $fix" >> "$scriptDir\Final_Report.txt"
        } 

else {
write-host " NECESARIO APLICAR $fix"
echo "$fix" >> "$scriptDir\Final_Report.txt"
}

}


###################################################################################
###################################################################################

#echo "#################################" >> "$scriptDir\Final_Report.txt"
#echo "############## XA5 ##############" >> "$scriptDir\Final_Report.txt"
#echo "########## XA3ISV0010P ###########" >> "$scriptDir\Final_Report.txt"

#$Hotfixes_Real = (get-wmiobject -class win32_quickfixengineering -computer# "XA3ISV0010P").HotFixID


#$Hotfixes_APP = Get-Content "$scriptDir\KB_XA5.txt"

#foreach ($fix in $Hotfixes_APP){

#if ($Hotfixes_Real -contains $fix){

#        write-host "YA EXISTE $fix"
        #echo "YA EXISTE $fix" >> "$scriptDir\Final_Report.txt"
#        } 

#else {
#write-host "NECESARIO APLICAR $fix"
#echo "$fix" >> "$scriptDir\Final_Report.txt"
#}

#}


###################################################################################
###################################################################################

echo "##################################" >> "$scriptDir\Final_Report.txt"
echo "############## CORE ##############" >> "$scriptDir\Final_Report.txt"
echo "########## XAPBSV0010P ###########" >> "$scriptDir\Final_Report.txt"

$Hotfixes_Real = (get-wmiobject -class win32_quickfixengineering -computer "XAPBSV0010P").HotFixID


$Hotfixes_APP = Get-Content "$scriptDir\KB_CORE.txt"

foreach ($fix in $Hotfixes_APP){

if ($Hotfixes_Real -contains $fix){

        write-host "YA EXISTE $fix" >> "$scriptDir\Final_Report.txt"
        #echo "YA EXISTE $fix" >> "$scriptDir\Final_Report.txt"
        } 

else {
write-host "NECESARIO APLICAR $fix"
echo "$fix" >> "$scriptDir\Final_Report.txt"
}

}


###################################################################################
###################################################################################

#echo "##################################" >> "$scriptDir\Final_Report.txt"
#echo "############## VDI ##############" >> "$scriptDir\Final_Report.txt"
#echo "########## VDNISV0010P ###########" >> "$scriptDir\Final_Report.txt"

#$Hotfixes_Real = (get-wmiobject -class win32_quickfixengineering -computer "VDNISV0010P").HotFixID


#$Hotfixes_APP = Get-Content "$scriptDir\KB_VDI.txt"

#foreach ($fix in $Hotfixes_APP){

#if ($Hotfixes_Real -contains $fix){
#
#        write-host "YA EXISTE $fix"
#        #echo "YA EXISTE $fix" >> "$scriptDir\Final_Report.txt"
#        } 

#else {
#write-host "NECESARIO APLICAR $fix"
#echo "$fix" >> "$scriptDir\Final_Report.txt"
#}

#}  