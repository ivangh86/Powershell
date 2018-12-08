########################################################################################################
#
#	Inventario de granja Citrix XenApp 6
#
#	Fecha: 2012/12/31
#
#

#Declaraci�n de constantes
##Contenido HTML
$HTML_Report = ""
#
####################################################

####################################################
#Declaracion de variables

####################################################


####################################################
#Funciones para gestionar la salida HTML



##########################
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
	#Contamos el n�mero de columnas
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
	

	#A�adimos los datos a la tabla
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


#########################
#
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

















####Servidores
$Servidores = Get-XAServer | sort-object ServerName
$ServidoresContenido = @()
Foreach ($Servidor in $Servidores)
{
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
		
	$Hotfixes = Get-XAServerHotfix -ServerName $Servidor.ServerName -Computername $Servidor_DC | sort-object Valid
	$Serv.Hotfixes = ""
	Foreach($hotfix in $hotfixes)
	{
		if ($hotfix.Valid)
		{
			$Serv.Hotfixes += "\t<b>" + $hotfix.HotfixName + "<b>\n"
		}
		else
		{
			$Serv.Hotfixes += "\t" + $hotfix.HotfixName + " (Reemplazado)\n"
		}
	}
	$ServidoresContenido += $Serv

}
####Fin Servidores











#########################################
#Informaci�n general de la granja
$Granja=Get-XAFarm 
$GranjaContenido = @()
$GranjaInformacion = "" | Select-Object Nombre, Version, NumeroServidores, NumeroAplicaciones

$GranjaInformacion.Nombre = $Granja.FarmName
$GranjaInformacion.Version = $Granja.ServerVersion
$GranjaInformacion.NumeroServidores = $Servidores.Count
$GranjaInformacion.NumeroAplicaciones = $AplicacionesReport.Count

$GranjaContenido += $GranjaInformacion


####Procesamos el archivo de salida

$HTML_Report = ""
$HTML_Report = InsertarCabeceraHTML

$Enlaces = @("Granja", "Registro de Configuracion", "Administradores", "Aplicaciones", "Directivas", "Directiva de Equilibrio de Carga", "Grupos de Trabajo", "Patrones de Carga", "Servidores", "Zonas")
$HTML_Report = InsertarMenuHTML $HTML_Report $Enlaces

#GRANJA
$TituloColumnas = @("Nombre","Version","Numero de Servidores","Numero de Aplicaciones")
$PropiedadesColumnas = @("Nombre","Version","NumeroServidores","NumeroAplicaciones")
$HTML_Report = InsertarTituloHTML $HTML_Report "Granja"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Granja" $TituloColumnas $PropiedadesColumnas $GranjaContenido

#REGISTRO DE CONFIGURACI�N
If ($RegConfiguracion.LoggingEnabled)
{
	$TituloColumnas = @("Habilitado","Permitir cambios con la bbdd desconectada","Solicitar credenciales","Tipo de Conexion", "Autenticacion", "Usuario", "Servidor", "BBDD")
	$PropiedadesColumnas = @("Habilitado", "Desconectado", "Limpiar", "Tipo", "Autenticacion", "Usuario", "Servidor", "BBDD")
	$HTML_Report = InsertarTituloHTML $HTML_Report "Registro de Configuracion"
	$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Reg" $TituloColumnas $PropiedadesColumnas $RegConfiguracionContenido
}
Else
{
	$HTML_Report = InsertarTituloHTML $HTML_Report "Registro de Configuracion"
	$HTML_Report = InsertarTextoHTML $HTML_Report "No est� configurado"
}

#ADMINISTRADORES
$TituloColumnas = @("Administrador","Tipo","Habilitado","Permisos en la granja","Permisos en las carpetas")
$PropiedadesColumnas = @("Administrador","Tipo","Habilitado","PermGranja","PermCarpetas")
$HTML_Report = InsertarTituloHTML $HTML_Report "Administradores"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Admin" $TituloColumnas $PropiedadesColumnas $AdministradoresContenido

#APLICACIONES
$TituloColumnas = @("Nombre", "Tipo", "Carpeta", "Comando", "Directorio de Trabajo", "Usuarios", "Grupos", "Servidores", "Audio", "Encriptacion", "Resolucion", "Color", "Habilitada")
$PropiedadesColumnas = @("Nombre", "Tipo", "Carpeta", "Comando", "Directorio", "Usuarios", "Grupos", "Servidores", "Audio", "Encriptacion", "Resolucion", "Color", "Habilitada")
$HTML_Report = InsertarTituloHTML $HTML_Report "Aplicaciones"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Aplica" $TituloColumnas $PropiedadesColumnas $AplicacionesContenido

#DIRECTIVAS
$TituloColumnas = @("Nombre", "Descripci�n", "Tipo", "Habilitada", "Prioridad", "Filtro", "Configuracion")
$PropiedadesColumnas = @("Nombre", "Descripcion", "Tipo", "Habilitada", "Prioridad", "Filtro", "Configuracion")
$HTML_Report = InsertarTituloHTML $HTML_Report "Directivas"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Directivas" $TituloColumnas $PropiedadesColumnas $DirectivasContenido

#DIRECTIVAS DE EQUILIBRIO DE CARGA
If ($DirectivasEquilibrioContenido)
{
	$TituloColumnas = @("Nombre", "Descripcion", "Habilitado", "Prioridad", "FiltroAC", "FiltroIP", "FiltroNombre", "FiltroUsuario", "Preferencia")
	$PropiedadesColumnas = @("Nombre", "Descripcion", "Habilitado", "Prioridad", "FiltroAC", "FiltroIP", "FiltroNombre", "FiltroUsuario", "Preferencia")
	$HTML_Report = InsertarTituloHTML $HTML_Report "Directivas de Equilibrio de Carga"
	$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Directivas_Carga" $TituloColumnas $PropiedadesColumnas $DirectivasEquilibrioContenido
}
Else
{
	$HTML_Report = InsertarTituloHTML $HTML_Report "Directivas de Equilibrio de Carga"
	$HTML_Report = InsertarTextoHTML $HTML_Report "No hay Directivas de Equilibrio de Carga"
}

#GRUPOS DE TRABAJO
$TituloColumnas = @("Nombre", "Descripcion", "Servidores","Aplicaciones")
$PropiedadesColumnas = @("Nombre", "Descripcion", "Servidores","Aplicaciones")
$HTML_Report = InsertarTituloHTML $HTML_Report "Grupos de Trabajo"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Grupos" $TituloColumnas $PropiedadesColumnas $GruposDeServidoresContenido

#PATRONES DE CARGA
$TituloColumnas = @("Nombre", "Descripcion", "Tipo", "Reglas")
$PropiedadesColumnas = @("Nombre", "Descripcion", "Tipo", "Reglas")
$HTML_Report = InsertarTituloHTML $HTML_Report "Patrones de Carga"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Patrones" $TituloColumnas $PropiedadesColumnas $PatronesContenido

#SERVIDORES
$TituloColumnas = @("Nombre", "Versi�n Citrix", "Sistema Operativo", "Carpeta", "Zona", "Preferencia Eleccion", "Hotfixes")
$PropiedadesColumnas = @("Nombre", "Citrix", "SistemaOperativo", "Carpeta", "Zona", "Eleccion", "Hotfixes")
$HTML_Report = InsertarTituloHTML $HTML_Report "Servidores"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Servidores" $TituloColumnas $PropiedadesColumnas $ServidoresContenido

#ZONAS
$TituloColumnas = @("Nombre", "DataCollector", "Servidores")
$PropiedadesColumnas = @("Nombre", "DC", "Servidores")
$HTML_Report = InsertarTituloHTML $HTML_Report "Zonas"
$HTML_Report = InsertarTablaHTML $HTML_Report "Tabla_Zonas" $TituloColumnas $PropiedadesColumnas $ZonasContenido




####Finalizando el documento
$ArchivoHTML="./$Granja.html"
$HTML_Report = InsertarFinHTML $HTML_Report
$HTML_Report > $ArchivoHTML

Invoke-Expression $ArchivoHTML