$securepassword = ConvertTo-SecureString -string "<your password>" -AsPlainText -Force 

$cred = new-object System.Management.Automation.PSCredential ("<your logon>", $securepassword)