####################################################
#
#      Autor        Iván Gómez Hernández  
#      Date:        10/05/2017
#      Script       SearchKB_v1.ps1
#      Descripción  Consulta Parches Seguridad KB
#      
#
#      Modificaciones: 
#      18/05/2017 - Filtrado para descartar maquinas apagadas
#
#####################################################
#Declaración de constantes
##Contenido HTML
$HTML_Report = ""

####################################################

####################################################
#Modulos necesarios

Add-PSSnapin citrix.*
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path



####################################################
#Generamos LOG

echo "" > "$scriptDir\WG.log"

$Log = "$scriptDir\WG.log"

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
	$HTML_Report += "<p>Documento creado por XXXX a las </p>`n"
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
#$Servidores = Get-XAServer -ServerName XARISV0600P
#$Servidores = Get-BrokerDesktop
$Servidores = Get-XAWorkerGroup


            


# Array con los campos y valores para la tabla

$ServidoresContenido = @()
Foreach ($Servidor in $Servidores)


{

TRY {

    $Serv = "" | select-Object Nombre, Descripcion, Numero
    #Formateamos la salida
    $Serv.Nombre = $Servidor.WorkerGroupName
 
    $Serv.Descripcion = $Servidor.Description
 
    $Serv.Numero = ($Servidor.ServerNames).count
 

        if ($Serv.Numero -eq "0") {
            
                $Serv.Numero = "\t" + '<B><FONT COLOR="red">' + "SIN MAQUINAS ASIGNADAS" + "</FONT>" + "\n"
                }


    $ServidoresContenido += $Serv

 }

Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            
            Get-Date | Add-Content "$scriptDir\WG.log"
            echo "" | Add-Content "$scriptDir\WG.log"
            echo "Se ha producido un error" | Add-Content "$scriptDir\WG.log"
            echo "" | Add-Content "$scriptDir\WG.log"
            echo "Mensaje:$ErrorMessage" | Add-Content "$scriptDir\WG.log"
            echo "Objeto:$FailedItem"    | Add-Content "$scriptDir\WG.log"
            echo "" | Add-Content "$scriptDir\WG.log"
            
               }
 
}

####Procesamos el archivo de salida

$HTML_Report = ""
$HTML_Report = InsertarCabeceraHTML

$Enlaces = @("Granja", "Registro de Configuracion", "Administradores", "Aplicaciones", "Directivas", "Directiva de Equilibrio de Carga", "Grupos de Trabajo", "Patrones de Carga", "Servidores", "Zonas")
$HTML_Report = InsertarMenuHTML $HTML_Report $Enlaces

#SERVIDORES
$TituloColumnas = @("Nombre", "Descripcion" , "Numero")
$PropiedadesColumnas = @("Nombre", "Descripcion", "Numero")
$HTML_Report = InsertarTituloHTML $HTML_Report "WorkerGroups"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_WorkerGroups" $TituloColumnas $PropiedadesColumnas $ServidoresContenido

$ArchivoHTML="$scriptDir\prueba.html"
$HTML_Report = InsertarFinHTML $HTML_Report
$HTML_Report > $ArchivoHTML

Invoke-Expression $ArchivoHTML