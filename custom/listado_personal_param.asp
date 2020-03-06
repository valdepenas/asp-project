<%@ Language=VBScript %>
<%' JCI 21/07/2003 : Por defecto , las consultas sólo se pueden buscar y ejecutar. Para crearlas,
  '                  modificarlas y/o borrarlas hay que pasarle el parámetro ges.
  ' JMG 26/03/2004 : Aceptación de parámetros en las consultas personalizadas.
  ' JMG 26/04/2004 : Muestra de parámetros personalizados en el listado.
  ' JMG 06/05/2004 : Permitir la gestión de las consultas personalizadas a través del sistema de gestión.
  ' JMG 19/05/2004 : Gestionar los parámetros personalizados.
  ' JMG 31/05/2004 : Incluir la exportación.
  
%>
<%response.buffer = true
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<% Server.ScriptTimeout = 1200 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="exportacion.inc" -->
<!--#include file="listado_personal_param.inc" -->
<!--#include file="../styles/ilionp.css.inc" -->
<!--#include file= "../styles/listTable.css.inc"-->
<title><%=LitTituloPag%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=Session("caracteres")%>" />
<meta http-equiv="Content-style-Type" content="text/css"/>
<link rel="stylesheet" href="../pantalla.css" media="SCREEN"/>
<link rel="stylesheet" href="../impresora.css" media="PRINT"/>
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
    //Desencadena la búsqueda del cliente cuyo numero se indica
    function TraerCliente(mode) {
        document.location.href="listado_personal_param.asp?ncliente=" + document.listado_personal_param.ncliente.value + "&mode=" + mode + "&dfecha=" + document.listado_personal_param.Dfecha.value + "&hfecha=" + document.listado_personal_param.Hfecha.value;
    }

    function Editar(fecha,usuario) {
        desdeGestion=document.listado_personal_param.desdeGestion.value;
        if (document.listado_personal_param.ges.value=="SI" || desdeGestion=="true") {
            document.getElementById("waitBoxOculto").style.visibility="visible";
            parent.botones.document.location="listado_personal_param_bt.asp?mode=browseConsulta&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&desdeGestion=" + desdeGestion + "&rutaFich=" + document.listado_personal_param.rutaFich.value;
            document.listado_personal_param.action="listado_personal_param.asp?fecha=" + fecha + "&mode=browseConsulta&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&ndoc=" + usuario + "&rutaFich=" + document.listado_personal_param.rutaFich.value;
            document.listado_personal_param.submit();
        }
        else {
            document.getElementById("waitBoxOculto").style.visibility="visible";
            parent.botones.document.location="listado_personal_param_bt.asp?mode=getParams&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&desdeGestion=" + desdeGestion + "&rutaFich=" + document.listado_personal_param.rutaFich.value;
            document.listado_personal_param.action="listado_personal_param.asp?fecha=" + fecha + "&mode=getParams&confirma=NO&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&ndoc=" + usuario + "&rutaFich=" + document.listado_personal_param.rutaFich.value;
            document.listado_personal_param.submit();
        }
    }

    // Añadida por JMG
    function validar(dato,tipo) {
        valido=true;
        switch(tipo)
        {
            case 0:  // Texto
                break;
            case 1:  // Numero
                expReg=new RegExp("[\+\-]?[0-9]*[.]?[0-9]*");
                cadena=expReg.exec(dato);
                if (cadena[0]!=dato)
                    valido=false;
                break;
            case 2:  // Fecha
                valido=chkdatetime(dato)  //Comprobamos que la fecha tenga el formato correcto
                break;
        }

        return valido;
    }

    function formatearFecha(fecha) {
        nuevaFechaEnCadena=fecha.getDate() + "-" + (fecha.getMonth()+1) + "-" + fecha.getFullYear();
        nuevaFechaEnCadena+=" " + fecha.getHours() + ":" + fecha.getMinutes() + ":" + fecha.getSeconds();

        return nuevaFechaEnCadena;
    }

    function generarNuevaConsulta(nparam,titulo,tipodato,campotabla,nuevaConsulta)
    {
        nuevaConsulta=nuevaConsulta.toLowerCase();

        if (campotabla=="")
        {
            valido=true;
            do
            {
                if (valido) valor=prompt("<%=LitIntroduzca%> " + titulo + ": ","");  //Pedimos el valor del argumento
                else valor=prompt("<%=LitVuelvaAIntroducir%> " + titulo + ": ","");  //Volvemos a pedir el valor
                if ((valor=="") || (valor==null))
                {
                    switch(tipodato)
                    {
                        case 0:  // Cadena de caracteres
                            valor="";
                            break;
                        case 1:  // Número
                            valor="0";
                            break;
                        case 2:  // Fecha
                            //valor=formatearFecha(new Date());
                            break;
                    }
                }
                if (tipodato==1) valor.replace(",",".");
                valido=validar(valor,tipodato);
                if (!valido)
                {
                    switch(tipodato)
                    {
                        case 1:
                            alert("<%=LitError%> : <%=LitNoEsUnValorNumerico%>");
                            break;
                        case 2:
                            alert("<%=LitError%> : <%=LitFormatoDeFechaErroneo%>");
                            break;
                    }
                }
            }
            while(!valido);
        }
        else valor=campotabla;

        document.listado_personal_param.valorParam.value=document.listado_personal_param.valorParam.value + " " + titulo + " = " + valor + "<br/>";
        document.getElementById(nparam).value=valor;

    
        /*20111221:
        do
         {
             nuevaConsulta=nuevaConsulta.replace(nparam,valor);
         }
         while (nuevaConsulta.indexOf(nparam)!=-1);
         */
        return nuevaConsulta;
    }

    //funcion para avanzar o retroceder páginas en los listados según la variable sentido
    function Mas(sentido,lote,campo,criterio,texto,modo) {
        if (modo!="configParam") {
            if (divObj = document.getElementById("waitBoxOculto")) document.getElementById("waitBoxOculto").style.visibility = "visible";
            document.forms[0].action=document.forms[0].name + ".asp?mode=" + modo + "&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto;
            document.forms[0].submit();
        }
        else {
            if (divObj = document.getElementById("waitBoxOculto")) document.getElementById("waitBoxOculto").style.visibility = "visible";
            fr_Tabla.document.listado_personal_param_det.action="listado_personal_param_det.asp?mode=search&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto;
            fr_Tabla.document.listado_personal_param_det.submit();
        }
    }

    function IrAPagina(dedonde,campo,criterio,texto,maximo,NomParamPag,modo) {
        if (divObj = document.getElementById("waitBoxOculto")) document.getElementById("waitBoxOculto").style.visibility = "visible";

        elemento="SaltoPagina" + dedonde;
        if (document.forms[0].name == "opciones") indiceform=1;
        else indiceform=0;
        if (isNaN(document.forms[indiceform].elements[elemento].value)) npagina=1;
        else {
            if (document.forms[indiceform].elements[elemento].value > maximo) npagina=maximo;
            else {
                if (document.forms[indiceform].elements[elemento].value <= 0) npagina=1;
                else npagina=document.forms[indiceform].elements[elemento].value;
            }
        }

        document.forms[indiceform].action=document.forms[indiceform].name + ".asp?" + NomParamPag + "=" + npagina +
        "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&mode=" + modo;
        document.forms[indiceform].submit();
    }

    function modifica(objeto) {
        usuariosAdd=document.listado_personal_param.usuariosAdd;
        usuariosDel=document.listado_personal_param.usuariosDel;
        if (objeto.checked) {
            if (usuariosDel.value.indexOf(objeto.name)>=0) usuariosDel.value=usuariosDel.value.replace(objeto.name,"");
            else usuariosAdd.value=usuariosAdd.value + "<->" + objeto.name + "<->";
        }
        else {
            if (usuariosAdd.value.indexOf(objeto.name)>=0) usuariosAdd.value=usuariosAdd.value.replace(objeto.name,"");
            else usuariosDel.value=usuariosDel.value + "<->" + objeto.name + "<->";
        }

        return true;
    }

    function ValidarCamposParam() {
        if (document.listado_personal_param.nparam.value=="") {
            window.alert("<%=LitMsgNParamVacio%>");
            document.listado_personal_param.nparam.focus();
            return false;
        }
        else {
            expReg=new RegExp("[0-9]*");
            cadena=expReg.exec(document.listado_personal_param.nparam.value);
            if (cadena[0]!=document.listado_personal_param.nparam.value) {
                window.alert("<%=LitNoEsUnValorNumerico%>");
                document.listado_personal_param.nparam.focus();
                return false;
            }
        }

        if (document.listado_personal_param.tipodato.value=="") {
            window.alert("<%=LitMsgTipoDatoVacio%>");
            document.listado_personal_param.tipodato.focus();
            return false;
        }
        return true;
    }

    function Insertar() {
        if (ValidarCamposParam()) {
            fr_Tabla.document.listado_personal_param_det.nparam.value=document.listado_personal_param.nparam.value;
            fr_Tabla.document.listado_personal_param_det.titulo.value=document.listado_personal_param.titulo.value;
            fr_Tabla.document.listado_personal_param_det.tipodato.value=document.listado_personal_param.tipodato.value;
            fr_Tabla.document.listado_personal_param_det.campotabla.value=document.listado_personal_param.campotabla.value;

            document.listado_personal_param.nparam.value="";
            document.listado_personal_param.titulo.value="";
            document.listado_personal_param.tipodato.selectedIndex=0;
            document.listado_personal_param.campotabla.value="";
            document.listado_personal_param.nparam.focus();
            fr_Tabla.document.listado_personal_param_det.action="listado_personal_param_det.asp?mode=save"
            fr_Tabla.document.listado_personal_param_det.submit();
        }
    }

</script>
<body class="BODY_ASP" onload="self.status='';">
<%function CadenaBusqueda(campo,criterio,texto,ncliente,sesion_usuario)
	if texto>"" then
		select case criterio
			case "contiene"
				CadenaBusqueda=campo + " like '%" + texto + "%' and usuario like '" & ncliente & sesion_usuario & "' order by fecha desc"
			case "empieza"
				CadenaBusqueda=campo + " like '" + texto + "%' and usuario like '" & ncliente & sesion_usuario & "' order by fecha desc"
			case "termina"
				CadenaBusqueda=campo + " like '%" + texto + "' and usuario like '" & ncliente & sesion_usuario & "' order by fecha desc"
			case "igual"
				CadenaBusqueda=campo + "='" + texto + "' and usuario like '" & ncliente & sesion_usuario & "' order by fecha desc"
		end select
	else
		CadenaBusqueda=" usuario like '" & ncliente & sesion_usuario & "' order by fecha desc"
	end if
end function

'Funciones añadidas por jmg

'Función que genera la cadena de búsqueda de los usuarios
function CadenaBusquedaUsuarios(campo,criterio,texto,ncliente)
	if texto>"" then
		select case criterio
			case "contiene"
				CadenaBusquedaUsuarios="where " & campo & " like '%" & texto & "%' and ncliente='" & ncliente & "'"
			case "empieza"
				CadenaBusquedaUsuarios="where " & campo & " like '" & texto & "%' and ncliente='" & ncliente & "'"
			case "termina"
				CadenaBusquedaUsuarios="where " & campo & " like '%" & texto & "' and ncliente='" & ncliente & "'"
			case "igual"
				CadenaBusquedaUsuarios="where " & campo & "='" & texto & "' and ncliente='" & ncliente & "'"
		end select
	else
		CadenaBusquedaUsuarios=" where ncliente='" & ncliente & "'"
	end if
end function

function enMatriz(matriz,valor)
	Dim encontrado
	Dim i

	encontrado = false

	if IsArray(matriz) then
		for i = LBound(matriz) to UBound(matriz)
			if matriz(i) = valor then
				encontrado = true
			end if
		next
	end if

	enMatriz = encontrado
end function

function obtenerParams(consulta)
	Dim parametros,listaParam()
	Dim expReg,i

	set expReg = new RegExp  'Expresión regular
	expReg.Pattern = "arg[0-9]+"
	expReg.IgnoreCase = true  'Ignoramos mayusculas o minusculas
	expReg.Global = true  'Aplicamos a toda la cadena
	set parametros=expReg.execute(consulta)

	i = 0
	for each parametro in parametros
		ReDim preserve listaParam(i)

		if enMatriz(listaParam,Right(parametro,Len(parametro)-3)) then
			ReDim preserve listaParam(i-1)
		else
			listaParam(i) = Right(parametro,Len(parametro)-3)
			i = i + 1
		end if
	next

	if parametros.count=0 then
		obtenerParams=""
	else
		obtenerParams = listaParam
	end if
end function

sub mostrarParams(cadenaParam)
	if cadenaParam<>"" then
		response.write "<div class='TDSINBORDECELDA7'>" & cadenaParam & "</div>" & "<br/>"
	end if

end sub

'Transformar fecha formato DDMMYYYY a YYYYMMDD
function convertDateToISOFormat(date)
    Dim fechaFin

    if InStr(date,"a. m.")<>0 then 
		date=replace(date,"a. m.","a.m.")
	end if
	if InStr(date,"p. m.")<>0 then 
		date=replace(date,"p. m.","p.m.")
	end if
    date = replace(replace(Reemplazar(date,".",":"), "a:m:", "am"), "p:m:", "pm")
    'Asegurar de que la fecha no tenga '+'
    date = replace(date,"+"," ")

    fechaFin = date

    'Tres o dos partes: DD/MM/YYYY, hh:mm:ss, y/o am o pm
    arrayDateCompl = split(date," ")
    'split para sacar DD(0) MM(1) YYYY(2)
    arrayDate2 = split(arrayDateCompl(0),"/")
    'Comprobamos que YYYY sea de 4 digitos
    if len(arrayDate2(2))=4 then
        'formato YYYYMMDD
        fechaFin = arrayDate2(2) & arrayDate2(1) & arrayDate2(0)
        'formato YYYYMMDD hh:mm:ss
        fechaFin = fechaFin & " " & arrayDateCompl(1)
        'comprobar si tienen am o pm, contando el tamaño del array
        sizeOfArray = uBound(arrayDateCompl) + 1
        'si existe am o pm
        if sizeOfArray=3 then
            'forma final YYYYMMDD hh:mm:ss xm
            fechaFin = fechaFin & " " & arrayDateCompl(2)
        end if
    end if

    convertDateToISOFormat = fechaFin
end function

'******************************************************************************
'Botones de navegación para las búsquedas.
sub NextPrev(lote,lotes,campo,criterio,texto,pos,modo)%>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
	<tr><td class="MAS">
        <%lote=cint(lote)
	    lotes=cint(lotes)
	    varias=false
		if lote>1 then
			%><a class="CELDAREF" href="javascript:Mas('prev',<%=lote%>,'<%=campo%>','<%=criterio%>','<%=texto%>','<%=modo%>');">
			<img src="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a><%
			varias=true
		end if
		textopag=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)
		%><font class="CELDA"><%=textopag%></font> <%

		if lote<lotes then
			%><a class="CELDAREF" href="javascript:Mas('next',<%=lote%>,'<%=campo%>','<%=criterio%>','<%=texto%>','<%=modo%>');">
			<img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a><%
			varias=true
		end if
		if varias=true then
	  	  %><font class="CELDA">&nbsp;&nbsp; <%=LitPagIrA%> <input class="CELDA" type="text" name="SaltoPagina<%=pos%>" size="2"/>&nbsp;&nbsp;<a class="CELDAREF" href="javascript:IrAPagina(<%=enc.EncodeForJavascript(null_s(pos))%>,'<%=enc.EncodeForJavascript(null_s(campo))%>','<%=enc.EncodeForJavascript(null_s(criterio))%>','<%=enc.EncodeForJavascript(null_s(texto))%>',<%=rst.PageCount%>,'lote','<%=enc.EncodeForJavascript(null_s(modo))%>');"><%=LitIr%></a></font><%
	  	end if
	%></td></tr>
</table>
<%end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0

	'Leer parámetros de la página
	mode=Request.QueryString("mode")

	campo=limpiaCadena(Request.QueryString("campo"))
	criterio=limpiaCadena(Request.QueryString("criterio"))
	texto=limpiaCadena(Request.QueryString("texto"))
	sop=limpiaCadena(Request.QueryString("ndoc"))
	if sop & "" = "" then sop = limpiaCadena(request.QueryString("sop"))
	if sop & "" = "" then sop = request.Form("sop")

	'Parámetro para saber si la consulta es modificable o no.
	ges=limpiaCadena(request("ges"))

	''ricardo 15-5-2007 se añade parametro de extension del fichero que se va a exportar
	dim ExtFichExport
	''ExtFichExport=limpiaCadena(request("ext"))
	
	''ricardo 22-12-2009 se añade parametro de ruta del fichero que se va a exportar
	dim RutaFichExport

	lote=limpiaCadena(request.QueryString("lote"))
	sentido=limpiaCadena(request.QueryString("sentido"))

	ncliente=limpiaCadena(Request("ncliente"))

	MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='079'", DSNIlion)
	MAXPDF=d_lookup("maxpdf", "limites_listados", "item='079'", DSNIlion)

	if mode="gestion" or mode="asignUser" or mode="asignSave" or mode="configParam" then
		desdeGestion="true"
		if mode="gestion" then
			mode="search"
		end if
	else
		desdeGestion=Request.form("desdeGestion")
		if desdeGestion="" then
			desdeGestion="false"
		end if
	end if

	if desdeGestion="true" then
		if mode="asignSave" or mode="asignUser" then
			PaintHeaderPopUp "ges_clientes.asp", LitAsignarConsulta
		elseif mode="configParam" then
			PaintHeaderPopUp "ges_clientes.asp", LitConfigurarParam
		else
			PaintHeaderPopUp "ges_clientes.asp", "Consulta Personalizada"
		end if
	else
     if request.querystring("acc")&""="link" then
        PintarCabeceraPopUp "Consulta Personalizada"
     else
		PintarCabecera "listado_personal_param.asp"
     end if
	end if%>

	<form name="listado_personal_param" method="post">

    <%Ayuda "listado_personal_param.asp"

	if desdeGestion="true" then
		ges="SI"
		rsocial=d_lookup("rsocial","clientes","ncliente='" & ncliente & "'",DSNIlion)%>
		<font class="ENCABEZADO"><b><%=LitCliente%>: </b></font><font class="CELDA"><%=enc.EncodeForHtmlAttribute(null_s(ncliente))%> - <%=enc.EncodeForHtmlAttribute(null_s(rsocial))%></font><br/>
    <%end if

	usuariosAdd=limpiaCadena(Request.form("usuariosAdd"))
	usuariosDel=limpiaCadena(Request.form("usuariosDel"))

    
	ndoc=Request("ndoc")

	cantidad=limpiaCadena(Request.QueryString("cantidad"))

	'A la consulta no se le pasa el limpia cadena, ya que tiene el select, la comilla simple, etc.
	consulta = Trim(Request.form("consulta"))
''ricardo 20-7-2006 ya no se quitaran los retornos de carro
''	consulta = replace(consulta,vbCrLf," ")
'''''''''''''''''''
	descripcion = limpiaCadena(Request.form("descripcion"))

	'Añadidas por jmg
	consultaConParam = Trim(Request.form("consultaConParam"))
''ricardo 20-7-2006 ya no se quitaran los retornos de carro
''''	consultaConParam = replace(consultaConParam,vbCrLf," ")
'''''''''''''''''''''''''''''''
	valorParams = Request.form("valorParam")
        
	fechadoc=limpiaCadena(Request.Form("fechadoc"))
    if fechadoc="" then
        fechadoc=limpiaCadena(Request.QueryString("fechadoc"))
    end if

	if fechadoc="" then
        fecha=limpiaCadena(Request.Form("fecha"))
        if fecha="" then
            fecha=limpiaCadena(Request.QueryString("fecha"))
        end if
	else
		fecha=fechadoc
	end if  
	
    fecha=replace(fecha,"+"," ")

	'Internacionalizacion del componente fecha
	if InStr(fecha,"a. m.")<>0 then 
		fecha=replace(fecha,"a. m.","a.m.")
	end if
	if InStr(fecha,"p. m.")<>0 then 
		fecha=replace(fecha,"p. m.","p.m.")
	end if

	confirma=limpiaCadena(Request.QueryString("confirma"))

	accede= limpiaCadena(Request.QueryString("acc"))
	if accede="" then
		accede= "otro"
	end if

	apaisado=Request.form("apaisado")

	if ((usuariosAdd & "")<>"" and mode="save") or (cantidad="mas" and mode="save") then
		ndoc=""
	end if
            
    dim hil
    obtenerParametros("ConsultaPer")
       
    %>
    <input type="hidden" name="mode_accesos_tienda" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
	<input type="hidden" name="ges" value="<%=enc.EncodeForHtmlAttribute(ges)%>"/>
	<input type="hidden" name="ext" value="<%=enc.EncodeForHtmlAttribute(ExtFichExport)%>"/>
	<input type="hidden" name="rutaFich" value="<%=enc.EncodeForHtmlAttribute(RutaFichExport)%>"/>
	<input type="hidden" name="desdeGestion" value="<%=enc.EncodeForHtmlAttribute(desdeGestion)%>"/>
	<input type="hidden" name="ndoc" value="<%=enc.EncodeForHtmlAttribute(ndoc)%>"/>
	<input type="hidden" name="maxpdf" value="<%=enc.EncodeForHtmlAttribute(MAXPDF)%>"/>
	<input type="hidden" name="sop" value="<%=enc.EncodeForHtmlAttribute(sop)%>"/>
    <input type="hidden" name="ncliente" value="<%=enc.EncodeForHtmlAttribute(ncliente)%>"/>

    <iframe id="id_exportar" name="exportar" src="listado_personal_param_exportar.asp" style="display:none;height:0px;width:0px"></iframe>

    <%set conn = Server.CreateObject("ADODB.Connection")
	set command =  Server.CreateObject("ADODB.Command")

	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")

	if desdeGestion<>"true" then
		ncliente=session("ncliente")
		dsnCliente=session("dsn_cliente")
		sesion_usuario=session("usuario")

		''ricardo 13/10/2004 se cambiara el usuario del dsncliente por el de DSNImport
		initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")

		donde=inStr(1,DSNImport,"Initial Catalog=",1)
		donde_fin=InStr(donde,DSNImport,";",1)
		if donde_fin=0 then
			donde_fin=len(DSNImport)
		end if
		cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))

		dsnCliente=cadena_dsn_final
	else
		consultaConParam=consulta
		sesion_usuario="%"
		cadenaDSN="select dsn as dsn_cliente from clientes with(nolock) where ncliente='" & ncliente & "'"
        rstAux.cursorlocation=3
		rstAux.Open cadenaDSN,DSNIlion
		if not rstAux.eof then
			dsnCliente=rstAux("dsn_cliente")
		end if
		rstAux.close

		''ricardo 13/10/2004 se cambiara el usuario del dsncliente por el de DSNImport
		initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")

		donde=inStr(1,DSNImport,"Initial Catalog=",1)
		donde_fin=InStr(donde,DSNImport,";",1)
		if donde_fin=0 then
			donde_fin=len(DSNImport)
		end if
		cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
		dsnCliente=cadena_dsn_final
	end if

	WaitBoxOculto LitEsperePorFavor
	Alarma "listado_personal_param.asp"

'****************************************************************************************************************
	if mode="save" or mode="first_save" then
		modificado=false
		if desdeGestion<>"true" then
			if mode="first_save" or mode="save" then
				if fecha>"" then
					fecha=replace(replace(Reemplazar(fecha,".",":"), "a:m:", "am"), "p:m:", "pm")
				else
					fecha=replace(replace(Reemplazar(now(),".",":"), "a:m:", "am"), "p:m:", "pm")
				end if

				seleccion="select * from queries where usuario like '" & ncliente & sesion_usuario & "' and fecha='" & convertDateToISOFormat(fecha) & "'"
				rst.open seleccion,dsnCliente,adOpenKeyset,adLockOptimistic
                nueva_querie=0
				if rst.eof then
                    nueva_querie=1
					rst.addnew
				end if
				rst("usuario")=ncliente & sesion_usuario
                ''ricardo 20-7-2006 la fecha no se modificara
                if nueva_querie=1 then
				    rst("fecha")=fecha
                end if
				rst("descripcion")=descripcion
				rst("consulta")=consultaConParam  'Modificado por jmg; cambio consulta por consultaConParam
				rst.update
				rst.close
				mode="browse"
				modificado=true
			end if
		else
			listaUsuariosAdd=split(usuariosAdd,"<->")
			listaUsuariosDel=split(usuariosDel,"<->")
			if InStr(usuariosAdd,"check")>0 or InStr(usuariosDel,"check")>0 or LBound(listaUsuariosAdd)<UBound(listaUsuariosAdd) or LBound(listaUsuariosDel)<UBound(listaUsuariosDel) then

				'Se va a modificar las consultas de los usuarios
				if fecha>"" then
					fecha=replace(replace(Reemplazar(fecha,".",":"), "a:m:", "am"), "p:m:", "pm")
				else
					fecha=replace(replace(Reemplazar(now(),".",":"), "a:m:", "am"), "p:m:", "pm")
				end if

				for each usuario in listaUsuariosAdd
					if Len(usuario)>0 then
                        
                                      
                        'convertir fecha
                        fecha_replace = Replace(fecha,"-","/")
				        fecha_replace = Replace(fecha_replace,"_"," ")
				        fecha_replace = Replace(fecha_replace,".",":")

						seleccion="select * from queries where usuario like '" & ncliente & replace(usuario,"check","") & "' and fecha='" & convertDateToISOFormat(fecha_replace) & "'"
						rst.open seleccion,dsnCliente,adOpenKeyset,adLockOptimistic

                        nueva_querie=0
						if rst.eof then
							rst.addnew
                            nueva_querie=1
						end if
						rst("usuario")=ncliente & replace(usuario,"check","")
                        ''ricardo 20-7-2006 la fecha no se modificara
                        if nueva_querie=1 then
						    rst("fecha")=fecha_replace
                        end if
						rst("descripcion")=descripcion
						rst("consulta")=consulta
						rst.update
						rst.close
						modificado=true
					end if
				next

				for each usuario in listaUsuariosDel
					if Len(usuario)>0 then
						seleccion="delete from queries where usuario like '" & ncliente & replace(usuario,"check","") & "' and fecha='" & convertDateToISOFormat(fecha) & "'"
						rst.open seleccion,dsnCliente,adOpenKeyset,adLockOptimistic
						modificado=true
					end if
				next
			else
                if fecha>"" then
					fecha=replace(replace(Reemplazar(fecha,".",":"), "a:m:", "am"), "p:m:", "pm")
				else
				    fecha=replace(replace(Reemplazar(now(),".",":"), "a:m:", "am"), "p:m:", "pm")
				end if
				if ndoc<>"" then

                    ' convertir fecha
                    fecha_replace = Replace(fecha,"-","/")
				    fecha_replace = Replace(fecha_replace,"_"," ")
				    fecha_replace = Replace(fecha_replace,".",":")

					seleccion="select * from queries where usuario like '" & ndoc & "' and fecha='" & convertDateToISOFormat(fecha_replace) & "'"
					rst.open seleccion,dsnCliente,adOpenKeyset,adLockOptimistic
                    nueva_querie=0
					if rst.eof then
						rst.addnew
						rst("fecha")=fecha_replace
                        nueva_querie=1
					end if
					rst("usuario")=ndoc
					rst("descripcion")=descripcion
					rst("consulta")=consulta
					rst.update
					rst.close
					modificado=true
				elseif cantidad="mas" then
					seleccion="select * from queries with(updlock) where usuario like '" & ncliente & "%' and fecha='" & convertDateToISOFormat(fecha) & "'"
					rst.open seleccion,dsnCliente,adOpenKeyset,adLockOptimistic
					fecha=Reemplazar(now(),".",":")
       
					while not rst.eof
						rst("usuario")=rst("usuario")
                        ''ricardo 20-7-2006 la fecha no se modificara
						''rst("fecha")=fecha
						rst("descripcion")=descripcion
						rst("consulta")=consulta
						rst.update
						rst.MoveNext
					Wend
					rst.close
					modificado=true
				end if
			end if

			if modificado=false then
				'No hay ningún usuario asignado%>
				<input type="hidden" name="consulta" value="<%=enc.EncodeForHtmlAttribute(null_s(consulta))%>"/>
				<input type="hidden" name="descripcion" value="<%=enc.EncodeForHtmlAttribute(null_s(descripcion))%>"/>
				<script language="javascript" type="text/javascript">
				    alert("<%=LitMsgNoUsuarios%>");
				  
				    <%

                    fecha_aux = Replace(fecha,"/","-")
				    fecha_aux = Replace(fecha_aux," ","_")
				    fecha_aux = Replace(fecha_aux,":",".")
				    'fecha_aux = "28-03-2014_12.55.20" 
                    %>
				    
				    parent.botones.document.location = "listado_personal_param_bt.asp?mode=asignSave&ncliente=<%=enc.EncodeForJavascript(ncliente)%>&titulo=<%=LitTituloAsignarConsulta%>&fechadoc=<%=enc.EncodeForHtmlAttribute(fecha_aux)%>"
				    document.listado_personal_param.action = "listado_personal_param.asp?mode=asignSave&ncliente=<%=enc.EncodeForJavascript(ncliente)%>&titulo=<%=LitTituloAsignarConsulta%>&fechadoc=<%=enc.EncodeForHtmlAttribute(fecha_aux)%>";

				    document.listado_personal_param.submit();
				</script>
			<%else%>
				<script language="javascript" type="text/javascript">
				    padre = top.opener;
				    // Identificamos la ventana que abrió esta ventana.
				    if ((padre != null) && ((padre.document.location.href.indexOf("listado_personal_param.asp?mode=search") >= 0) || (padre.document.location.href.indexOf("listado_personal_param.asp?mode=gestion") >= 0))) {
				        padre.document.listado_personal_param.action = "listado_personal_param.asp?mode=gestion";
				        padre.document.listado_personal_param.submit();
				        top.close();
				    }
				    else {
				        document.listado_personal_param.action = "listado_personal_param.asp?mode=gestion";
				        document.listado_personal_param.submit();
				    }
				</script>
			<%end if

			mode="search"
		end if
	end if

'****************************************************************************************************************
	if mode="delete" then
		fecha=Reemplazar(fecha,".",":")
		if cantidad<>"1" then
			rst.open "delete from queries with(rowlock) where usuario like '" & ncliente & sesion_usuario & "' and fecha='" & convertDateToISOFormat(fecha) & "'",dsnCliente,adOpenKeyset,adLockOptimistic
		else
			rst.open "delete from queries with(rowlock) where usuario like '" & ndoc & "' and fecha='" & convertDateToISOFormat(fecha) & "'",dsnCliente,adOpenKeyset,adLockOptimistic
		end if
		mode="search"
	end if
'****************************************************************************************************************
	if mode="add" then
		modif="disabled"
		if ges="SI" then modif=""

		if consulta="" then
			fecha=""
		end if%>

		<input type="hidden" name="fecha" value="<%=enc.EncodeForHtmlAttribute(fecha)%>"/>
		<hr/>
		<table width='100%' border='<%=borde%>' cellspacing="1" cellpadding="1"><%
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitDescripcion +":"
			 	DrawInputCelda "CELDA " & modif,"","",100,0,"","descripcion",descripcion
			CloseFila
			DrawFila color_blau
				%><td><br/><br/></td><%
			CloseFila
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitConsulta +":"
				DrawTextCelda "CELDA " & modif,"","",10,100,"","consulta",consultaConParam
			CloseFila%>

			<script language="javascript" type="text/javascript">document.listado_personal_param.descripcion.focus();</script>
		</table><hr/><%

'****************************************************************************************************************
	'Inicio modificacion por jmg
	'Obtener los parametros
	elseif mode="getParams" then
        ''ricardo 11-4-2006 si la consulta se esta abriendo desde otra el usuario sera el que venga por parametro
        if request.querystring("acc")="link" and ndoc & "">"" then
	        sesion_usuario=trimCodEmpresa(ndoc)
        end if

		if consulta="" then
			'Obtener la consulta a ejecutar
            rst.cursorlocation=3
			rst.Open "select usuario, consulta, descripcion from queries with(nolock) where usuario like '" & ncliente & sesion_usuario & "' and fecha='" & convertDateToISOFormat(fecha) & "'",dsnCliente
			if not rst.eof then
				if desdeGestion<>"true" then
					checkCadena rst("usuario")
				end if
                ''ricardo 20-7-2006 ya no se quitaran los retornos de carro
                '''consulta=replace(rst("consulta"),vbCrLf," ")
                consulta=rst("consulta")
                '''''''''''''''''''''
				descripcion=rst("descripcion")
			end if
			rst.close
		end if%>
		<input type="hidden" name="descripcion" value="<%=enc.EncodeForHtmlAttribute(null_s(descripcion))%>"/>
        <%if ges="SI" then%>
		<input type="hidden" name="consulta" value="<%=enc.EncodeForHtmlAttribute(null_s(consulta))%>"/>
        <input type="hidden" name="consultaConParam" value="<%=enc.EncodeForHtmlAttribute(null_s(consulta))%>"/>
        <%end if%>
		<input type="hidden" name="fecha" value="<%=enc.EncodeForHtmlAttribute(fecha)%>"/>
		<input type="hidden" name="valorParam" value="<%=enc.EncodeForHtmlAttribute(valorParams)%>"/>
        <%Dim params,listaParametros,consultaParams,i%>
		<script language="javascript" type="text/javascript">
		    <%params=obtenerParams(consulta)
		    if (isarray(params)) then  'Ignoramos mayusculas y minusculas
		    listaParametros = ""
		    for i = LBound(params) to UBound(params)
                listaParametros = listaParametros & "nparam=" & params(i)

                if i < UBound(params) then
                    listaParametros = listaParametros & " or "
                end if
            next
               
            'Obtenemos los nombres de cada parametro
            consultaParams = "select * from paramqueries where " & listaParametros
            rst.cursorlocation = 3
            rst.Open consultaParams,dsnCliente%>

		        //ricardo 20-7-2006 ya no se quitaran los retornos de carro
            nuevaConsulta="<%=replace(consulta,vbCrLf,"\r\n")%>";
            <%if not rst.eof then%>
                parametros=new Array(0);
                cadenaParam="";
                listParamhidden=""
                <%do while not rst.eof%>
                    parametros[<%=rst("nparam")%>]=new Object();
                    parametros[<%=rst("nparam")%>].nombre = "arg<%=rst("nparam")%>";
                    parametros[<%=rst("nparam")%>].titulo = "<%=rst("titulo")%>";
                    parametros[<%=rst("nparam")%>].tipodato = <%=rst("tipodato")%>;
                    <%if isNull(rst("campotabla")) then%>
                        parametros[<%=rst("nparam")%>].campotabla = "";
                    <%else%>
                        parametros[<%=rst("nparam")%>].campotabla ="<%=rst("campotabla")%>";
                    <%end if%>
                    <%'20111221:
                    listParamhidden=listParamhidden&"<input name='arg"&rst("nparam")&"' id='arg"&rst("nparam")&"'  type='hidden' value='' />"
                       
                    rst.MoveNext
                loop
                %></script>
             <%=listParamhidden%>
             <script language="javascript" type="text/javascript">
                 <%i=0
                 for each parametro in params
                     if accede="link" then%>
                         nuevaConsulta=generarNuevaConsulta(parametros[<%=parametro%>].nombre,parametros[<%=parametro%>].titulo,parametros[<%=parametro%>].tipodato,"<%=enc.EncodeForJavascript(limpiaCadena(Request.QueryString("arg"+params(i))))%>",nuevaConsulta);
                     <%else%>
                         nuevaConsulta=generarNuevaConsulta(parametros[<%=parametro%>].nombre,parametros[<%=parametro%>].titulo,parametros[<%=parametro%>].tipodato,parametros[<%=parametro%>].campotabla,nuevaConsulta);
                     <%end if
                     i= i + 1
                 next%>
                     //document.listado_personal_param.consulta.value=nuevaConsulta;
             <%end if
             rst.Close
         end if%>
         document.getElementById("waitBoxOculto").style.visibility="visible";
         parent.botones.document.location="listado_personal_param_bt.asp?mode=browse&ges=" + document.listado_personal_param.ges.value +"&ext=" + document.listado_personal_param.ext.value + "&acc=<%=enc.EncodeForJavascript(accede)%>&rutaFich=" + document.listado_personal_param.rutaFich.value;
         document.listado_personal_param.action="listado_personal_param.asp?mode=browse&confirma=NO&ges=" + document.listado_personal_param.ges.value +"&ext=" + document.listado_personal_param.ext.value + "&acc=<%=enc.EncodeForJavascript(accede)%>&rutaFich=" + document.listado_personal_param.rutaFich.value;
         document.listado_personal_param.submit();
        </script>
       
	    <%'Fin modificacion por jmg
'****************************************************************************************************************
	'Mostrar el listado.
	elseif mode="browse" then    'Modificada por jmg; Añadido consultaConParam
		Dim tieneParams
		tieneParams = false
        consulta = ""
        onlyread = 0

		if consulta="" then
       	'Obtener la consulta a ejecutar

            fecha=replace(replace(Reemplazar(fecha,".",":"), "a:m:", "am"), "p:m:", "pm")
            rst.cursorlocation=3
            ''rst.Open "select usuario, consulta, descripcion from queries with(nolock) where usuario like '" & ncliente & sesion_usuario & "' and fecha='" & fecha & "'",dsnCliente
			
			rst.Open "select usuario, consulta, descripcion, onlyread from queries with(nolock) where usuario like '" & ncliente & sesion_usuario & "' and fecha='" & convertDateToISOFormat(fecha) & "'",dsnCliente
			if not rst.eof then
				if desdeGestion<>"true" then
					checkCadena rst("usuario")
				end if
                ''ricardo 20-7-2006 ya no se quitaran los retornos de carro
                '''consulta=replace(rst("consulta"),vbCrLf," ")
                consulta=rst("consulta")
                ''''''''''''''''
				descripcion=rst("descripcion")
                
                onlyread=rst("onlyread")

                if consultaConParam="" then
			        consultaConParam=consulta
		        end if

				params=obtenerParams(consulta)
				if (isarray(params)) then  'La consulta tiene parametros, hay que reemplazarlos
					'20111221:
                    'tieneParams = true
                    parameters=""
			        for i = LBound(params) to UBound(params)
				        consulta = replace(consulta,("arg"&params(i)),(request.Form("arg"&params(i))),1,-1,1)
                        parameters = parameters&("arg"&params(i))&"="&(request.Form("arg"&params(i)))&"#"
                        '20120102:
                        response.Write "<input name='arg"&params(i)&"' id='"&("arg"&params(i))&"'  type='hidden' value='"&(request.Form("arg"&params(i)))&"' />"
			        next
                    parameters = left(parameters, len(parameters)-1)
				end if
			end if
			rst.close
        else
            if consultaConParam="" then
			    consultaConParam=consulta
		    end if
		end if%>
        <input type="hidden" name="params" value="<%=enc.EncodeForHtmlAttribute(parameters)%>"/>
        <input type="hidden" name="ncliente" value="<%=enc.EncodeForHtmlAttribute(ncliente)%>"/>
        <input type="hidden" name="sesion_usuario" value="<%=enc.EncodeForHtmlAttribute(sesion_usuario)%>"/>
		<input type="hidden" name="descripcion" value="<%=enc.EncodeForHtmlAttribute(descripcion)%>"/>
        <%if ges="SI" then%>
		<input type="hidden" name="consulta" value="<%=enc.EncodeForHtmlAttribute(null_s(consulta))%>"/>
		<input type="hidden" name="consultaConParam" value="<%=enc.EncodeForHtmlAttribute(null_s(consultaConParam))%>"/>
        <%end if%>
        <input type="hidden" name="fecha" value="<%=enc.EncodeForHtmlAttribute(fecha)%>"/>
		<input type="hidden" name="valorParam" value="<%=enc.EncodeForHtmlAttribute(valorParams)%>"/>
        <%if tieneParams then%>
			<script language="javascript" type="text/javascript">
			    document.getElementById("waitBoxOculto").style.visibility="visible";
			    parent.botones.document.location="listado_personal_param_bt.asp?mode=getParams&ges=" + document.listado_personal_param.ges.value +"&ext=" + document.listado_personal_param.ext.value + "&acc=<%=enc.EncodeForJavascript(accede)%>&rutaFich=" + document.listado_personal_param.rutaFich.value ;
			    document.listado_personal_param.action = "listado_personal_param.asp?mode=getParams&confirma=NO&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&acc=<%=enc.EncodeForJavascript(accede)%>" + "&arg16=<%=enc.EncodeForJavascript(limpiaCadena(request.querystring("arg16")))%>&rutaFich=" + document.listado_personal_param.rutaFich.value;
			    document.listado_personal_param.submit();
			</script>
        <%elseif consulta<>"" then
            ''ricardo 14-4-2009 se auditara la ejecucion de la consulta, añadiendo a la descripcion los parametros puestos
            descripcionAudit=descripcion & " ." & LitParametros & ": " & replace(valorParams,"<br/>"," , ")



            ExecuteConsultaPersonalizadaBOT=d_lookup("1", "auditoria", "login='"&session("usuario")&"' and datediff(s,dateadd(s,-2,getdate()),fecha) >=0 and accion ='Fin Consulta Perso' and fecha > '"&Date()&"'", dsnCliente)
            if(ExecuteConsultaPersonalizadaBOT="1") then
                if session("ncliente")>"00001" then auditar_ins_bor session("usuario"),descripcionAudit,"","","","","Inicio Consulta Perso"
            else
                   
                ''ricardo 16/5/2005 se cambia el metodo de ejecucion para asi poder ponerle tiempo de espera
			    'JAR 04/04/06: Auditar el inicio de la ejecución.
			    if session("ncliente")>"00001" then auditar_ins_bor session("usuario"),descripcionAudit,"","","","","Inicio Consulta Perso"
                
            if (onlyread = True) then
			    conn.open Session("backendlistados")
            else
                conn.open dsnCliente
            end if
            	''conn.open dsnCliente
			    conn.cursorlocation=3
			    command.ActiveConnection =conn
			    ''ricardo 23-7-2010 se aumenta el tiempo de 90 a 300 para el tema de reparto de genero de frutas eloy
			    command.CommandTimeout = 300''90
			    command.CommandText=consulta
			    set rst = Command.Execute        
            end if
           

			'JAR 04/04/06: Auditar la finalización de la ejecución.
			if session("ncliente")>"00001" then auditar_ins_bor session("usuario"),descripcionAudit,"","","","","Fin Consulta Perso"

			if err.number<>0 then
				textoError=err.description
				rst.close
                conn.close

				on error goto 0
				%><script language="javascript" type="text/javascript">
				      alert("<%=LitError%> : <%=textoError%>");
				      parent.botones.document.location = "listado_personal_param_bt.asp?mode=browseConsulta&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&rutaFich=" + document.listado_personal_param.rutaFich.value;
				      document.listado_personal_param.action = "listado_personal_param.asp?fecha=" + document.listado_personal_param.fecha.value
					 + "&mode=browseConsulta&consulta=" + document.listado_personal_param.consulta.value + "&descripcion=" + document.listado_personal_param.descripcion.value + "&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&ndoc=<%=enc.EncodeForJavascript(ncliente&sesion_usuario)%>&rutaFich=" + document.listado_personal_param.rutaFich.value;
				      document.listado_personal_param.submit();
				</script><%
			else
				on error goto 0

				if rst.State>0 then%>
					<input type="hidden" name="NumRegs" value="<%=rst.recordcount%>"/>
					<%if not rst.eof then
						lotes=rst.recordcount/MAXPAGINA
						if lotes>clng(lotes) then
							lotes=clng(lotes)+1
						else
							lotes=clng(lotes)
						end if

						if lote="" then lote=1
						if sentido="next" then
							lote=lote+1
						elseif sentido="prev" then
							lote=lote-1
						end if

						rst.PageSize=MAXPAGINA
						rst.AbsolutePage=lote

						%><font class=CABECERA><b><%=enc.EncodeForHtmlAttribute(null_s(descripcion))%></b></font><br/><br/><%

						if InStr(1,consultaConParam,"arg",1)<>0 then
							call mostrarParams(valorParams)
						end if%>

						<hr/>
                        <%if hil&""<>"1" then %>
						    <font class=CABECERA><%=LitImprimirApaisado%>: </font>
						    <input type="checkbox" name="apaisado" <%=iif(nz_b(apaisado)<>0,"checked","")%>/>
                        <%
                        end if
                        NextPrev enc.EncodeForJavascript(lote),lotes,enc.EncodeForJavascript(campo),enc.EncodeForJavascript(criterio),enc.EncodeForJavascript(texto),1,enc.EncodeForJavascript(mode)
                        %>
						<table width='100%' border='0' cellspacing="0" cellpadding="0">
							<thead><%
							DrawFila color_fondo
								for each campo in rst.fields
									%><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"><%=enc.EncodeForHtmlAttribute(null_s(campo.name))%></td><%
								next
							CloseFila
							%></thead>
							<tbody><%
							fila=1
							while not rst.eof and fila<=MAXPAGINA
								DrawFila ""
									for each campo in rst.fields
										' NOTA: si el campo es binario (por ejemplo una imagen) se producirá un error, no se puede mostrar%>
									    <!--<td class="tdbordeCELDA7"><%=iif(rst(campo.name)&""="","&nbsp;",rst(campo.name))%></td>-->
									    <td class="tdbordeCELDA7" style="font-weight: inherit; color: #555; font-family: 'Gill Sans', 'Gill Sans MT', 'Calibri', 'Trebuchet MS', 'sans-serif';">
									        <%strResponse = iif(rst(campo.name)&""="","&nbsp;",rst(campo.name))

									        '   GPD (03/04/2007).
									        '   Si tiene contratado el módulo de gasóleo profesional.
									        If ModuloContratado(session("ncliente"),"35") Then
									            If InStr(UCase(strResponse),"<A") = 0 And InStr(UCase(strResponse),"</a>") = 0 Then
									                strResponse = Replace(strResponse,"<","&lt;")
									                strResponse = Replace(strResponse,">","&gt;")
									            End If
									        End If

                                            Response.Write strResponse%>
									    </td>
                                    <%next
								CloseFila
                                response.Flush()
								rst.movenext

								fila=fila+1
							wend%>
							</tbody>
						</table><%
						NextPrev enc.EncodeForJavascript(lote),lotes,enc.EncodeForJavascript(campo),enc.EncodeForJavascript(criterio),enc.EncodeForJavascript(texto),2,enc.EncodeForJavascript(mode)
						rst.close
                        conn.close
					else
						rst.close
                        conn.close
						%><font class='CEROFILAS'><%=LitCeroFilas%></font><%
					end if
				else
					%><font class='CEROFILAS'><%=LitConsultaCorrecta%></font><%
				end if
			end if
		else
			%><font class='CEROFILAS'><%=LitCeroFilas%></font><%
		end if
'****************************************************************************************************************
	elseif mode="browseConsulta" then
		fecha=replace(replace(Reemplazar(fecha,".",":"), "a:m:", "am"), "p:m:", "pm")

		modif="disabled"
		if ges="SI" then	modif=""

		if desdeGestion<>"true" then
            rst.cursorlocation=3
			rst.open "select * from queries with(nolock) where usuario like '" & ndoc & "' and fecha='" & convertDateToISOFormat(fecha) & "'",dsnCliente
		else
			strselect="select * from queries with(nolock) where usuario like '" & ndoc & "' and fecha='" & convertDateToISOFormat(fecha) & "'"
            rst.cursorlocation=3
			rst.open strselect,dsnCliente
		end if%>

		<table width="100%" border="<%=borde%>" cellspacing="1" cellpadding="1">
    		<%if not rst.eof then
				%><input type="hidden" name="fecha" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("fecha")))%>"/><%
				if desdeGestion="true" then
					DrawFila color_blau
						DrawCelda2 "CELDA", "left", true, LitUsuario & ":"
						DrawCelda2 "CELDA", "left", false, right(rst("usuario"),len(rst("usuario"))-5)
					CloseFila
					DrawFila color_blau
						DrawCelda2 "CELDA colspan='2'", "left", true, "<hr/>"
					CloseFila
				end if
				DrawFila color_blau
					DrawCelda2 "CELDA", "left", true, LitFecha +":"
					DrawCelda2 "CELDA", "left", false, rst("fecha")
				CloseFila
				DrawFila color_blau
					DrawCelda2 "CELDA " & modif, "left", true, LitDescripcion +":"
					DrawInputCelda "CELDA " & modif,"","",100,0,"","descripcion",rst("descripcion")
				CloseFila
				DrawFila color_blau
					DrawCelda2 "CELDA " & modif, "left", true, LitConsulta +":"
					DrawTextCelda "CELDA " & modif,"","",10,100,"","consulta",rst("consulta")
				CloseFila
			else
				DrawFila color_blau
					DrawCelda2 "CELDA " & modif, "left", false, LitDescripcion +":"
					DrawInputCelda "CELDA " & modif,"","",100,0,"","descripcion",""
				CloseFila
				DrawFila color_blau
					DrawCelda2 "CELDA " & modif, "left", false, LitConsulta +":"
					DrawTextCelda "CELDA " & modif,"","",10,100,"","consulta",""
				CloseFila
			end if
			
			''ricardo 10-3-2009 se comenta esta linea, ya que esta duplicada mas arriba
			%>
			<!--<input type="hidden" name="ndoc" value="<%=ndoc%>"/>-->
			<%
			''fin ricardo 10-3-2009
			
			if modif="disabled" then%>
				<input type="hidden" name="descripcion" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))%>"/>
				<!--<input type="hidden" name="consulta" value="<%'=rst("consulta")%>"/>-->
			<%end if%>
		</table><%
		rst.close
'****************************************************************************************************************
	elseif mode="search" then
		%><hr/><%
		if ges<>"" and desdeGestion<>"true" then
			%><script language="javascript" type="text/javascript">
			      parent.botones.document.location = "listado_personal_param_bt.asp?mode=<%=enc.EncodeForJavascript(mode)%>&ges=" + document.listado_personal_param.ges.value + "&ext=" + document.listado_personal_param.ext.value + "&desdeGestion=<%=enc.EncodeForJavascript(desdeGestion)%>&rutaFich=" + document.listado_personal_param.rutaFich.value;
			</script><%
		end if

		if lote="" then lote=1
			strwhere=CadenaBusqueda(campo,criterio,texto,ncliente,sesion_usuario)
		if desdeGestion<>"true" then
			strselect="select usuario,fecha,descripcion,consulta from queries with(nolock) where " & strwhere
            rst.cursorlocation=3
			rst.Open strselect,dsnCliente
		else
			strselect="select * from queries with(nolock) "
			strselect=strselect & iif(strwhere<>"","where " & strwhere,"")
            rst.cursorlocation=3
			rst.Open strselect,dsnCliente
		end if

		if not rst.EOF then
			lotes=rst.RecordCount/NumReg
			if lotes>clng(lotes) then
				lotes=clng(lotes)+1
			else
				lotes=clng(lotes)
			end if

			if sentido="next" then
				lote=lote+1
			elseif sentido="prev" then
				lote=lote-1
			end if

			rst.PageSize=NumReg
			rst.AbsolutePage=lote

			NextPrev enc.EncodeForJavascript(lote),lotes,enc.EncodeForJavascript(campo),enc.EncodeForJavascript(criterio),enc.EncodeForJavascript(texto),1,enc.EncodeForJavascript(mode)

			%><table width='100%' border='0' cellspacing="1" cellpadding="1"><%
				'Fila de encabezado
				DrawFila color_fondo
					DrawCelda "ENCABEZADOL","","",0,LitFecha
					DrawCelda "ENCABEZADOL","","",0,LitDescripcion
					if desdeGestion="true" then
						DrawCelda "ENCABEZADOL","","",0,LitUsuario
						if sop <> "1" then 
						DrawCelda "ENCABEZADOL","","",0,""
						DrawCelda "ENCABEZADOL","","",0,""
						end if
					end if
				CloseFila

				fila=1
				while not rst.EOF and fila<=NumReg
					'Comprobar que la consulta es de la empresa en la que está identificado.
					if desdeGestion<>"true" then
						checkCadena rst("usuario")
					end if
					'Seleccionar el color de la fila.
					if ((fila+1) mod 2)=0 then
						color=color_blau
					else
						color=color_terra
					end if

					DrawFila color
					    if sop <> "1" then 
						DrawCeldahref "CELDAREF","left","false",rst("fecha"),"javascript:Editar('" & replace(rst("fecha")," ","+") & "','" & enc.EncodeForHtmlAttribute(null_s(rst("usuario"))) & "')"
						else
						DrawCelda "CELDA","","",0,rst("fecha")
						end if
						DrawCelda "CELDA","","",0,rst("descripcion")
						if desdeGestion="true" then
							DrawCelda "CELDA","","",0,right(rst("usuario"),len(rst("usuario"))-5)
							if sop <> "1" then 
							DrawCeldahref "CELDAREF","left","false","<img " & ParamImgUsuarios & " src='../gestion/images/" & ImgUsuarios & "' alt='" & LitAsignar & "' title='" & LitAsignar & "'/>","'javascript:AbrirVentana(""../central.asp?pag1=custom/listado_personal_param.asp&pag2=custom/listado_personal_param_bt.asp&mode=asignUser&ncliente=" & enc.EncodeForJavascript(ncliente) & "&titulo=" & LitTituloAsignarConsulta & "&fechadoc=" & enc.EncodeForJavascript(null_s(rst("fecha"))) & "&ndoc=" & enc.EncodeForJavascript(null_s(rst("usuario"))) & ""","""",410,775);'"
							DrawCeldahref "CELDAREF","left","false","<img " & ParamParamQueries & " src='../gestion/images/" & ImgParamQueries & "' alt='" & LitConfigurarParam & "' title='" & LitConfigurarParam & "'/>","'javascript:AbrirVentana(""../central.asp?pag1=custom/listado_personal_param.asp&pag2=custom/listado_personal_param_bt.asp&mode=configParam&ncliente=" & enc.EncodeForJavascript(ncliente) & "&titulo=" & LitConfigurarParam & "&fechadoc=" & enc.EncodeForJavascript(null_s(rst("fecha"))) & "&ndoc=" & enc.EncodeForJavascript(null_s(rst("usuario"))) & ""","""",410,775);'"
							end if
						end if
					CloseFila
					fila=fila+1
					rst.MoveNext
				wend
			%></table><%
			NextPrev enc.EncodeForJavascript(lote),lotes,enc.EncodeForJavascript(campo),enc.EncodeForJavascript(criterio),enc.EncodeForJavascript(texto),2,enc.EncodeForJavascript(mode)

			rst.close
		else
			%><font class="CEROFILAS"><%=LitCeroFilas%></font><%
		end if
'****************************************************************************************************************
	elseif mode="asignUser" or mode="asignSave" then  'Lista de usuarios
        fecha=replace(replace(Reemplazar(fecha,".",":"), "a:m:", "am"), "p:m:", "pm")
             
		if mode="asignUser" then

			'Obtener la consulta
			cadenaSelect="select * from queries with(nolock) where usuario='" & ndoc & "' and fecha='" & convertDateToISOFormat(fecha) & "'"
            rst.cursorlocation=3
			rst.Open cadenaSelect,dsnCliente
			if not rst.eof then
''ricardo 20-7-2006 ya no se quitaran los retornos de carro
''				consulta=replace(rst("consulta"),vbCrLf," ")
				consulta=rst("consulta")
''''''''''''''''''''''''''''''''''''
				descripcion=rst("descripcion")
			end if
			rst.close
		end if%>
		<input type="hidden" name="fecha" value="<%=enc.EncodeForHtmlAttribute(fecha)%>"/>
		<input type="hidden" name="ndoc" value="<%=enc.EncodeForHtmlAttribute(ndoc)%>"/>
        <%if ges="SI" then%>
		<input type="hidden" name="consulta" value="<%=consulta%>"/>
        <%end if%>
		<input type="hidden" name="descripcion" value="<%=descripcion%>"/>
		<input type="hidden" name="usuariosAdd" value="<%=enc.EncodeForHtmlAttribute(usuariosAdd)%>"/>
		<input type="hidden" name="usuariosDel" value="<%=enc.EncodeForHtmlAttribute(usuariosDel)%>"/>
		<font class='ENCABEZADO'><b><%=LitConsulta%>: </b></font><font class='CELDA'><%=enc.EncodeForHtmlAttribute(null_s(descripcion))%></font><br/>
		<hr/>
		<%strWhere=CadenaBusquedaUsuarios(campo,criterio,texto,ncliente) & " and administrar=1 and cu.fbaja is null"
		strSelect="select distinct i.entrada as login,i.nombre as usuario from clientes_users as cu, indice as i " & strWhere & " and i.entrada=cu.usuario order by i.nombre"
        rst.cursorlocation=3
		rst.Open strSelect,DSNIlion

		if not rst.eof then
			lotes=rst.RecordCount/NumReg
			if lotes>clng(lotes) then
				lotes=clng(lotes)+1
			else
				lotes=clng(lotes)
			end if

			if lote="" then lote=1
			if sentido="next" then
				lote=lote+1
			elseif sentido="prev" then
				lote=lote-1
			end if

			rst.PageSize=NumReg
			rst.AbsolutePage=lote

			NextPrev enc.EncodeForJavascript(lote),lotes,enc.EncodeForJavascript(campo),enc.EncodeForJavascript(criterio),enc.EncodeForJavascript(texto),1,enc.EncodeForJavascript(mode)

			%><table width='100%' border='0' cellspacing="1" cellpadding="1"><%
				'Fila de encabezado
				DrawFila color_fondo
					DrawCelda "ENCABEZADOL","","",0,LitAsignarConsulta
					DrawCelda "ENCABEZADOL","","",0,LitLogin
					DrawCelda "ENCABEZADOL","","",0,LitUsuario
				CloseFila

				fila=1
				do while not rst.EOF and fila<=NumReg

					'Seleccionar el color de la fila.
					if ((fila+1) mod 2)=0 then
						color=color_blau
					else
						color=color_terra
					end if

                    ' convertir fecha                                                        
				    fecha_replace = Replace(fecha,"-","/")
				    fecha_replace = Replace(fecha_replace,"_"," ")
				    fecha_replace = Replace(fecha_replace,".",":")
                
					strQueries="select * from queries with(nolock) where usuario like '" & ncliente & rst("login") & "' and fecha='" & convertDateToISOFormat(fecha_replace) & "'"

                    'strQueries="select * from queries with(nolock) where usuario like '" & ncliente & rst("login") & "' and fecha='" & fecha & "'"

                    rstAux.cursorlocation=3
					rstAux.Open strQueries,dsnCliente

					DrawFila color
						DrawCelda "CELDA","","",0,"<input type='checkbox' name='check" & rst("login") & "' " & iif(((not rstAux.eof) and (inStr(usuariosDel,"<->check" & enc.EncodeForHtmlAttribute(null_s(rst("login"))) & "<->")=0)) or (inStr(usuariosAdd,"<->check" & enc.EncodeForHtmlAttribute(null_s(rst("login"))) & "<->")),"checked","") & " onclick='modifica(this)'/>"
						DrawCelda "CELDA","","",0,rst("login")
						DrawCelda "CELDA","","",0,rst("usuario")
					CloseFila

					rstAux.close

					fila=fila+1
					rst.MoveNext
				loop
			%></table><%
			NextPrev enc.EncodeForJavascript(lote),lotes,enc.EncodeForJavascript(campo),enc.EncodeForJavascript(criterio),enc.EncodeForJavascript(texto),2,enc.EncodeForJavascript(mode)
		else
			%><font class="CEROFILAS"><%=LitCeroFilas%></font><%
		end if
'****************************************************************************************************************
	elseif mode="configParam" then  'Configurar parámetros%>
		<hr/>
		<table width="100%" border="<%=borde%>" cellspacing="1" cellpadding="1">
			<%'Fila de encabezado
			DrawFila color_fondo
				DrawCelda "ENCABEZADOL","","",0,LitNParam
				DrawCelda "ENCABEZADOL","","",0,LitTitulo
				DrawCelda "ENCABEZADOL","","",0,LitTipoDato
				DrawCelda "ENCABEZADOL","","",0,LitCampoTabla
				DrawCelda "ENCABEZADOL","","",0,""
			Closefila

			DrawFila color_blau
				DrawInputCelda "CELDA","","",5,0,"","nparam",""
				DrawInputCelda "CELDA maxlength='50'","","",29,0,"","titulo",""%>
				<td class="CELDA">
					<select class="IN" name="tipodato">
						<option value="0"><%=LitTexto%></option>
						<option value="1"><%=LitNumero%></option>
						<option value="2"><%=LitFecha%></option>
					</select>
				</td>
				<%DrawInputCelda "CELDA maxlength='50' disabled","","",29,0,"","campotabla",""%>
				<td width="10%" align="center">
					<a href="javascript:Insertar();" ><img src="../images/<%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>
				</td>
			<%CloseFila%>
			<script language="javascript" type="text/javascript">document.listado_personal_param.nparam.focus();</script>
		</table>

		<!-- IFrame que muestra los parametros ya insertados -->
		<iframe id="fr_Tabla" name="fr_Tabla" src='listado_personal_param_det.asp?mode=search&ncliente=<%=enc.EncodeForHtmlAttribute(ncliente)%>' width='100%' height='190' frameborder="yes" noresize="noresize">
		</iframe>
		<table width="100%">
			<tr>
				<td class="CELDA7" width="250">
					<span ID="barras" style="display:none">
					</span>
				</td>
			</tr>
		</table>
	<%end if

	set rstAux = Nothing
	set rst = Nothing
	set conn = nothing
	set command=nothing%>
</form>
<%end if%>
</body>
</html>