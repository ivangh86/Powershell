$days=((Get-Date).AddDays(-5)) 

#Get-WinEvent -Computername "XACISV3044P" -FilterHashtable @{Logname="Application" ;ID=@(1511); StartTime=$days } -EA 0
#Get-WinEvent -Computername "XACISV3044P" -FilterHashtable @{Logname="System" ;ID=@(5); StartTime=$days } -EA 0


Get-WinEvent -Computername "XACISV3044P" -FilterHashtable @{Logname="Application" ;ID=@(1530); StartTime=$days } -EA 0 

