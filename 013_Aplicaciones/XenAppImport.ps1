#Adding Citrix stuff
Add-PSSnapin Citrix* 
 
# Importing applications and associating icons with applications 
# Setting Delivery Group
$dg = "SDI W2K8R2 OFI DES"
 
# Importing Application Data
# Set the path to your .xml
$apps = Import-Clixml C:\temp\Apps1.xml
 
foreach ($app in $apps)
{
   #Importing Icon 
   $IconUid = New-BrokerIcon -EncodedIconData $app.EncodedIconData
  
   #Adding Application
    Write-Host $app.Name
    Try
    {
    if ($app.CommandLineArguments.Length -lt 2) {$app.CommandLineArguments = " "} 
        #Adding Application
        write-host "BrowserName" $app.BrowserName
        write-host "ComdLineExe" $app.CommandLineExecutable
        write-host "Description" $app.Description
        write-host "ComdLineArg" $app.CommandLineArguments
        write-host "Enabled    " $app.Enabled
        write-host "Name       " $app.Name
        write-host "UserFiltEna" $app.UserFilterEnabled
        write-host "WorkingDire" $app.WorkingDirectory
        write-host "Published  " $app.PublishedName
        write-host "ApplicationName"  $app.ApplicationName
        write-host "AdminFolderName"  $app.AdminFolderName
        
 
        If($app.ClientFolder -ne $null)
        {
        
        #  | Out-Null
        # -AdminFolder $app.AdminFolderName
        New-BrokerApplication  -BrowserName $app.BrowserName -CommandLineExecutable $app.CommandLineExecutable -WorkingDirectory $app.WorkingDirectory -DesktopGroup $dg  -Name $app.ApplicationName -PublishedName $app.PublishedName -ClientFolder $app.ClientFolder -Enabled $app.Enabled -UserFilterEnabled $app.UserFilterEnabled -CommandLineArguments $app.CommandLineArguments -Description $app.Description | Out-Null
        }else {
        

        #  | Out-Null
        # -AdminFolder $app.AdminFolderName
        New-BrokerApplication  -BrowserName $app.BrowserName -CommandLineExecutable $app.CommandLineExecutable -WorkingDirectory $app.WorkingDirectory -DesktopGroup $dg  -Name $app.ApplicationName -PublishedName $app.PublishedName -ClientFolder "\" -Enabled $app.Enabled -UserFilterEnabled $app.UserFilterEnabled -CommandLineArguments $app.CommandLineArguments -Description $app.Description | Out-Null
        }
 
        #Setting applications icon  
        Set-BrokerApplication $app.Name -IconUid $IconUid.Uid
        write-host "Working on Icon" $IconUid.Uid
    }
    Catch
    {
        write-host $_.Exception.Message
        write-host $_.Exception.ItemName
        write-host "Error on "  $app.BrowserName
        write-host "Error on "  $app.CommandLineExecutable
        write-host "Error on "  $app.Description
        write-host "Error on "  $app.CommandLineArguments
        write-host "Error on "  $app.Enabled
        write-host "Error on "  $app.Name
        write-host "Error on "  $app.UserFilterEnabled
 
    }
 
  # Adding User Associations
 If($app.AssociatedUserNames -ne $null) 
 {
 Try
 {
    $users = $app.AssociatedUserNames
 
    foreach($user in $users)
    {   
        #estaba este $app.Name 
        write-host "Working on User  "  $user  $app.ApplicationName
        Write-Host $app.ApplicationName
        Add-BrokerUser -Name "$user" -Application $app.Name
              }
       }
    Catch
    {
        write-host $_.Exception.Message
        write-host $_.Exception.ItemName
        write-host "Error on User  "  $user  $app.ApplicationName
 }
}
}