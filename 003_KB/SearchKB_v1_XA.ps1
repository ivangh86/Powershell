####################################################
#
#      Autor:        Iván Gómez Hernández  
#      Date:         10/05/2017
#      Script:       SearchKB_v1_XA.ps1
#      Descripción:  Consulta Parches Seguridad KB
#      
#      Resultados:   "$scriptDir\ReportKB.html" "$scriptDir\ReportKB.csv" 
#
#      Modificaciones: 
#      18/05/2017 - Filtrado para descartar maquinas apagadas
#      18/05/2017 - Añadir Estado Apagado y Falta KB
#      18/05/2017 - Formatear colores estados APAGADOS y FALTA KB
#      19/05/2017 - Modificada la consulta "wmic /node:$Servidor qfe" ya que en VDI no se aceptan. Cambiada a "get-wmiobject -class win32_quickfixengineering"
#      22/05/2017 - Se intenta adaptar para que haga las consultas en paralelo mediante "JOBS" pero se producen otro problemas como el Out of memory o fichero en uso.
#      24/05/2017 - Se añade que revise si tiene los PSSnapin de Citrix y en caso contrarios que los instale.
#
#####################################################
#Declaración de constantes
##Contenido HTML
$HTML_Report = ""
$Fecha = get-Date
####################################################

####################################################
#Modulos necesarios


if  ((Get-PSSnapin -Name "Citrix.Common.Commands" -ErrorAction SilentlyContinue) -eq $null )

{
    write-host "Añadimos Snapin Citrix Citrix.Common.Commands"
    Add-PsSnapin Citrix*
}


if  ((Get-PSSnapin -Name "Citrix.Common.GroupPolicy" -ErrorAction SilentlyContinue) -eq $null )

{
    write-host "Añadimos Snapin Citrix.Common.GroupPolicy"
    Add-PsSnapin Citrix*
}


if  ((Get-PSSnapin -Name "Citrix.XenApp.Commands" -ErrorAction SilentlyContinue) -eq $null )

{
    write-host "Añadimos Snapin Citrix.XenApp.Commands"
    Add-PsSnapin Citrix*
}



$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path



####################################################
#Generamos LOG

echo "" > "$scriptDir\searchkbv1.log"

$Log = "$scriptDir\searchkbv1.log"

####################################################
#Funciones para gestionar la salida HTML


function InsertarCabeceraHTML
{
	#Preparar HTML_report

	$HTML_Report = "<html>`n"
	$HTML_Report += "<head>`n"
	$HTML_Report += "<style type=`"text/css`"> table.tableCitrix {font-family:arial;background-color: #CDCDCD;margin:10px 0pt 15px;font-size: 8pt;width: 100%;	text-align: left;}</style>`n"
	$HTML_Report += "<style type=`"text/css`"> table.tableCitrix thead tr th, table.tableCitrix tfoot tr th {	background-color: #E6EEEE;	border: 1px solid #FFF;	font-size: 8pt;	padding: 4px;}</style>`n"
	$HTML_Report += "<style type=`"text/css`"> table.tableCitrix tbody td {	color: #3D3D3D;	padding: 4px;	background-color: #FFF;	vertical-align: top;}</style>`n"
	$HTML_Report += "</head>`n"
	$HTML_Report += "<body>`n"
	return $HTML_Report
}

function InsertarMenuHTML
{
	$HTML_Report += "`n"
	return $HTML_Report
}

function InsertarTituloHTML ($HTML_Report,$Titulo)
{
	$HTML_Report += "<h1><a name=`"$Titulo`">$Titulo</a></h1>`n"
	return $HTML_Report
}

function InsertarTextoHTML ($HTML_Report,$Texto)
{
	$HTML_Report += "<p>$Texto</p>`n"
	return $HTML_Report
}

function InsertarFinHTML ($HTML_Report)
{
	$HTML_Report += "<p>Documento creado por" +"&nbsp"+ "<b>" + "Wintel Citrix BS" + "</b>" +"&nbsp"+ "a fecha de: $Fecha </p>`n"
	$HTML_Report += "</body>`n"
	$HTML_Report += "</html>`n"
	return $HTML_Report
}

function InsertarTablaHTML ($HTML_Report, $Titulo, $TituloColumnas, $PropiedadesColumnas, $Contenido)
{
	#Contamos el número de columnas
	$ColumnasNum = $TituloColumnas.Count

	#Preparamos la tabla
	$HTML_Report += "<table id=`"$Titulo`" class=`"tableCitrix`">`n"

	#Insertamos los titulos de las columnas
	
	$HTML_Report += "<thead>`n<tr>`n"
	for ($col = 0; $col -lt $ColumnasNum; $col++) {
		$a = $TituloColumnas[$col]
    	$HTML_Report += "<th>$a</th>`n"
	}
	$HTML_Report += "</tr>`n</thead>`n"
	

	#Añadimos los datos a la tabla
	$i = 1
	$j = $Contenido.Count
	
	for($row = 0; $row -lt $Contenido.Count; $row++)
	{
		$HTML_Report += "<tr>`n"
  		for ($col = 0; $col -lt $ColumnasNum; $col++){
			$a = $Contenido[$row].($PropiedadesColumnas[$col])
			
			$b = $a.ToString().replace("\n","<br/>")
			$c = $b.replace("\t","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; - ")
			$HTML_Report += "<td>$c</td>`n"
		}
		$HTML_Report += "</tr>`n"
	}
	
	#Finalizamos la tabla
	$HTML_Report += "</table>`n"
	$HTML_Report += "<a href=`"#Menu`">[Menu]</a>`n"
	return $HTML_Report
}


function InsertarMenuHTML ($HTML_Report, $Enlaces)
{
	$HTML_Report += "<a name=`"Menu`"></a>"
	$HTML_Report += "<table id=`"Menu`" class=`"tableCitrix`">`n"
	$HTML_Report += "<tr>`n"
	
	foreach ($Enlace in $Enlaces)
	{
		$HTML_Report += "<td>`n"
		$HTML_Report += "<a href=`"#$Enlace`">$Enlace</a>"
		$HTML_Report += "</td>`n"
	}
	$HTML_Report += "</tr>`n"
	$HTML_Report += "</table>`n"
	return $HTML_Report
}


###########################################################################

#Podemos pasarle un fichero o una consulta más generica de todos los serviores de la granja
#
#$Servidores = Get-XAServer XATISV0001P
#$Servidores = Get-BrokerDesktop
$Servidores = Get-XAServer | Sort-Object -Property ServerName
#$Servidores = Get-BrokerDesktop | Where-Object {$_.DNSName -like "*vdnbpv0001d*" } | select-object DNSName
#$Servidores = "localhost","XAC1","XAC","XAC2","XAC3","XAC4","XAC5","XAC6","XAC7","XAC8"
#
###########################################################################
            


# Array con los campos y valores para la tabla

$ServidoresContenido = @()
Foreach ($Servidor in $Servidores)


{

TRY {

    $Serv = "" | select-Object Nombre, Citrix, SistemaOperativo, Carpeta, Zona, Eleccion, Hotfixes
    #Formateamos la salida
    $Serv.Nombre= $Servidor.ServerName
 
    $ServidorCitrix = $Servidor.CitrixProductName + ", " + $Servidor.CitrixVersion + ", " + $Servidor.CitrixEdition + ", " + $Servidor.CitrixServicePack
    $Serv.Citrix += "$ServidorCitrix"
 
    $SO=$Servidor.OSVersion.ToString() + ", " + $Servidor.OSServicePack
    $Serv.SistemaOperativo = $SO
 
    $Serv.Carpeta = $Servidor.FolderPath
 
    $Serv.Zona = $Servidor.ZoneName
 
    $Serv.Eleccion = $Servidor.ElectionPreference
 

        # filtramos si la máquina esta apagada para descartarla de la busqued 
        $MaquinaEncendida = Test-Connection -BufferSize 32 -Count 1 -ComputerName $Servidor -Quiet

        if ($MaquinaEncendida -eq "True") {
            $date = Get-date
            write-host "Buscando parches en $Servidor - $date"
            # Escogemos los KB que queramos consultar (KB4012212,KB4012213,KB4012214,KB4012598)
            #$Hotfixes =  wmic /node:$Servidor qfe get hotfixid | Select-String 'KB4012212','KB4012213','KB4012214','KB4012598'
            $Hotfixes = get-wmiobject -class win32_quickfixengineering -computer $Servidor | select-object HotFixID | Select-String 'KB4012212','KB4012213','KB4012214','KB4012598'
            $Serv.Hotfixes  = ""
                Write-Host $Hotfixes
                if (!$Hotfixes)

                {
                #<B><FONT COLOR="red">Texto ROJO </FONT>
                $Serv.Hotfixes = "\t" + '<B><FONT COLOR="red">' + "FALTA KB" + "</FONT>" + "\n"
                }

                else {
                    
                    Foreach($hotfix in $Hotfixes)
                    {
                    $Serv.Hotfixes += "\t" + $hotfix + "\n"
                    }
                }
        }
        else {
        $Serv.Hotfixes ="\t"+ "<b>" + "APAGADA" + "</b>" + "\n"
        }   
    
    $ServidoresContenido += $Serv

 }

Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            
            Get-Date | Add-Content "$scriptDir\searchkbv1.log"
            echo "" | Add-Content "$scriptDir\searchkbv1.log"
            echo "Se ha producido un error" | Add-Content "$scriptDir\searchkbv1.log"
            echo "" | Add-Content "$scriptDir\searchkbv1.log"
            echo "Mensaje:$ErrorMessage" | Add-Content "$scriptDir\searchkbv1.log"
            echo "Objeto:$FailedItem"    | Add-Content "$scriptDir\searchkbv1.log"
            echo "" | Add-Content "$scriptDir\searchkbv1.log"
            
               }
 
}


####Procesamos el archivo de salida

$HTML_Report = ""
$HTML_Report = InsertarCabeceraHTML

$Enlaces = @("Granja", "Registro de Configuracion", "Administradores", "Aplicaciones", "Directivas", "Directiva de Equilibrio de Carga", "Grupos de Trabajo", "Patrones de Carga", "Servidores", "Zonas")
$HTML_Report = InsertarMenuHTML $HTML_Report $Enlaces

#SERVIDORES
$TituloColumnas = @("Nombre", "Versión Citrix", "Sistema Operativo", "Carpeta", "Zona", "Preferencia Eleccion", "Hotfixes")
$PropiedadesColumnas = @("Nombre", "Citrix", "SistemaOperativo", "Carpeta", "Zona", "Eleccion", "Hotfixes")
$HTML_Report = InsertarTituloHTML $HTML_Report "Servidores Produccion"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Servidores" $TituloColumnas $PropiedadesColumnas $ServidoresContenido

$ArchivoHTML="$scriptDir\ReportKB.html"
$HTML_Report = InsertarFinHTML $HTML_Report
$HTML_Report > $ArchivoHTML


# Añadimos el informe a EXCEL
$Prueba = $ServidoresContenido | convertto-csv -delimiter ","  -notypeinformation | out-file "$scriptDir\ReportKB.csv"



Invoke-Expression $ArchivoHTML