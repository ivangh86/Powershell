$objSID = New-Object System.Security.Principal.SecurityIdentifier ("S-1-5-80-3815303045-3875903309-1662709731-719598877-3593490124")
$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
$objUser.Value 