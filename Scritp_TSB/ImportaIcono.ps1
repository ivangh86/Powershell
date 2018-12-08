 # Titulo: Importar exportar iconos individuales.
 #
 # Autor: Ivan Gomez
 # Fecha: 16/03/2018
 # Descripción: En este caso ya hemos exportados los iconos a un ficheros .xml  con el script IconsExport.ps1 
 # Fichero se llama \\serverhqr1\soft_ctx\Team\Ivan\Iconos.xml 
 # En dicho fichero estan codificados los Iconos de las aplicaciones Appv por si nos falta alguno con el DR para cogerlo de alli y volver a importar.
 # Para ello cogemos la cadena del icono que nos interesa del fichero Iconos.xml y la copiamos en la variable $EncondeIconData
 # Tambien incluimos en la variable $PublishedNameApp el nombre de la aplicacion que queremos modificar el icono original, concretamente el campo PublishedName de la aplicación.
 #
 ##############################################################################################################################################
 #$app = Get-BrokerApplication -PublishedName "3270 TSB FirstData UATD-TEST"
 #$EncodedIconData = (Get-Brokericon -Uid $app.IconUid).EncodedIconData # Grabs Icon Image

 
 #Modificar por el PublishedName de la aplicacion que quremos cambiar el icono por el que habia
 $PublishedNameApp = "Hana Studio Test"


 #Modificar por la cadena que pertenezca al icono que queremos importar
  $EncodedIconData = "AAABAAEAICAQAAEABADoAgAAFgAAACgAAAAgAAAAQAAAAAEABAAAAAAAgAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC/AAC/AAAAv78AvwAAAL8AvwC/vwAAwMDAAICAgAAAAP8AAP8AAAD//wD/AAAA/wD/AP//AAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAHd3d3d3d3cAAAAAAAAAAAA
Ad3d3d3cAAAAAAAAAAAAAB3d3d3d3cAAAAAAAAAAAAHd3d3d3d3cAAAAAAIiIiIiIiIiIiIiIiIiIiACHd3d3d3d3d3d3d3d3d3gAh3d3d3d3d3d3d3d3d3d4AId4iIiIiIiIiIiIiIiHeACHeMzMzMzMzMzMzMzMh3gAh3i8zMzMzMzMzMzMzId4AId4u8zMzMzMzMzMzMqHeACHeLu8zJzM
zMzMzKqqh3gAh3i7u8mZzMzMzKqqqod4AId4u7uZmZzMqqqqqqqHeACHeLu5mZmZyqqqqqqqh3gAh3i7mZmZmZqqqqqqqod4AId4u7mZmZmZqqqqqqqHeACHeLu7mZmZmqqqqqqqh3gAh3i7u7mZmaqqqqqqqod4AId4u7u7mZqqqqqqqqqHeACHeLu7u5n/qqqqqqqvh3gAh3i7u7////qqq
qqq/4d4AId4u7//////qqqqr/+HeACHeLv///////qqqv//h3gAh3i/////////qq///4d4AId4v/////////r///+HeACHeIiIiIiIiIiIiIiIh3gAh3d3d3d3d3d3d3d3d3d4AId3d3d3d3d3d3d3d3d3eACIiIiIiIiIiIiIiIiIiIgAAAAAAAAAAAAAAAAAAAAADgAAAH8AAAD/+AAf//
AAD/gAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAQ=="

$IconUid = New-BrokerIcon -EncodedIconData $EncodedIconData

 

Get-BrokerApplication -PublishedName $PublishedNameApp |  Set-BrokerApplication  -IconUid $IconUid.Uid



  


