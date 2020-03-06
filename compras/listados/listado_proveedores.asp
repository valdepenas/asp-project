<%@ Language=VBScript %>
<% 
dim enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
<meta http-equiv="Content-style-Type" content="text/css">
<LINK REL="styleSHEET" href="../../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../../impresora.css" MEDIA="PRINT">
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<% if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../varios2.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->
<!--#include file="../proveedores.inc" -->

<!--#include file="../../tablasResponsive.inc" -->

<!--#include file="../../styles/formularios.css.inc" -->
     
<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">

    function PonerFechas(modo)
    {
	    if (modo=="alta")
	    {
		    document.listado_proveedores.fbhasta.value="";
		    document.listado_proveedores.fbdesde.value="";
	    }
	    else
	    {
		    document.listado_proveedores.fhasta.value="";
		    document.listado_proveedores.fdesde.value="";
	    }
    }
</script>
<body onload="self.status='';" class="BODY_ASP">
<%'RGU 17/11/2007 CAMBIO DSN PARA LISTADOS
sub escribir_cabecera()%>
					<td class=tdbordeCELDA7><b><%=LitNumProveedor%></b></td>
					<td class=tdbordeCELDA7><b><%=LitRazonSocial%></b></td>
					<%if opccif = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitCif%></b></td>
					<%end if%>
					<%if opccontacto = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitContacto1%></b></td>
					<%end if%>
					<%if opcdomicilio = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitDomicilio%></b></td>
					<%end if%>
					<%if opccodigopostal = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitCodPostal%></b></td>
					<%end if%>
					<%if opcpoblacion = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitPoblacion%></b></td>
					<%end if%>
					<%if opcprovincia = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitProvincia%></b></td>
					<%end if%>
					<%if opctelefono = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitTelefono%></b></td>
					<%end if%>
					<%if opcfalta = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitFechaAlta%></b></td>
					<%end if
					if opcfbaja = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitFechaBaja%></b></td>
					<%end if
					if opcnomcom= "1" then %>
						<td class=tdbordeCELDA7><b><%=LitNomComProv%></b></td>
					<%end if
					if opcfax= "1" then %>
						<td class=tdbordeCELDA7><b><%=LitFax%></b></td>
					<%end if
					if opctelmov= "1" then %>
						<td class=tdbordeCELDA7><b><%=LitTel2%></b></td>
					<%end if
					if opcweb= "1" then %>
						<td class=tdbordeCELDA7><b><%=ucase(LitWEB)%></b></td>
					<%end if
					if opcobs= "1" then %>
						<td class=tdbordeCELDA7><b><%=LitObservaciones%></b></td>
					<%end if
					if opcemail= "1" then %>
						<td class=tdbordeCELDA7><b><%=LitEmaProv%></b></td>
					<%end if
					if opctarifa = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitTarifa%></b></td>
					<%end if
					if opccuenta = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitCCont%></b></td>
					<%end if
					if opcformapago = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitFormaPago%></b></td>
					<%end if
					if opctipopago = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitTipoPago%></b></td>
					<%end if
					if opcrfinanciero = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitRFinan%></b></td>
					<%end if
					if opcrequivalencia = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitREquiv%></b></td>
					<%end if
					if opcIRPF = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitIRPF1%></b></td>
					<%end if
					if opclven1 = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitVencom1%></b></td>
					<%end if
					if opclven2 = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitVencom2%></b></td>
					<%end if
					if opcentidad = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitEntidad%></b></td>
					<%end if%>
					<%if opcnumcuenta = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitNumcuenta%></b></td>
					<%end if
					if opcactividad = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitTactividad%></b></td>
					<%end if
					if opctproveedor = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitTProveedor%></b></td>
					<%end if%>

					<%if opcportes = "1" then %>
						<td class=tdbordeCELDA7><b><%=LitPortes%></b></td>
					<%end if
					if ver_campo1="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "01' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo2="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "02' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo3="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "03' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo4="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "04' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo5="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "05' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if

					if ver_campo6="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "06' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo7="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "07' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo8="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "08' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo9="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "09' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if
					if ver_campo10="on" then
						titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "10' and tabla='PROVEEDORES'",session("backendlistados"))
						%><td class=TDBORDECELDA7><b><%=EncodeForHtml(titulo)%></b></td><%
					end if

end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************

const borde=0
 %>
<form name="listado_proveedores" method="post">

<% PintarCabecera "listado_proveedores.asp"
   WaitBoxOculto LitEsperePorFavor
'Leer parámetros de la página
	mode = EncodeForHtml(Request.QueryString("mode"))
	if ucase(mode) = "BROWSE" then mode ="imp"

	apaisado=iif(limpiaCadena(request.form("apaisado"))>"","SI","")

	fdesde		= limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde	= limpiaCadena(Request.form("fdesde"))
	end if

	fhasta		= limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta	= limpiaCadena(Request.form("fhasta"))
	end if

	fbdesde		= limpiaCadena(Request.QueryString("fbdesde"))
	if fbdesde ="" then
		fbdesde	= limpiaCadena(Request.form("fbdesde"))
	end if

	fbhasta		= limpiaCadena(Request.QueryString("fbhasta"))
	if fbhasta ="" then
		fbhasta	= limpiaCadena(Request.form("fbhasta"))
	end if

	razon_social		= limpiaCadena(Request.QueryString("razon_social"))
	if razon_social ="" then
		razon_social	= limpiaCadena(Request.form("razon_social"))
	end if

	poblacion		= limpiaCadena(Request.QueryString("poblacion"))
	if poblacion ="" then
		poblacion	= limpiaCadena(Request.form("poblacion"))
	end if

	provincia		= limpiaCadena(Request.QueryString("provincia"))
	if provincia ="" then
		provincia	= limpiaCadena(Request.form("provincia"))
	end if

	tarifa		= limpiaCadena(Request.QueryString("tarifa"))
	if tarifa ="" then
		tarifa	= limpiaCadena(Request.form("tarifa"))
	end if

	formapago		= limpiaCadena(Request.QueryString("formapago"))
	if formapago ="" then
		formapago	= limpiaCadena(Request.form("formapago"))
	end if

	actividad	= limpiaCadena(Request.QueryString("actividad"))
	if actividad ="" then
		actividad	= limpiaCadena(Request.form("actividad"))
	end if

	tipo_proveedor	= limpiaCadena(Request.QueryString("tipo_proveedor"))
	if tipo_proveedor ="" then
		tipo_proveedor	= limpiaCadena(Request.form("tipo_proveedor"))
	end if

	ordenar	= limpiaCadena(Request.QueryString("ordenar"))
	if ordenar ="" then
		ordenar	= limpiaCadena(Request.form("ordenar"))
	end if

	if request.form("opcproveedorbaja")>"" then
		opcproveedorbaja = "1"
	end if

	if request.form("solodistribuidores")>"" then
		solodistribuidores= "1"
	end if

	if request.form("mostrarcontactos")>"" then
		mostrarcontactos= "1"
	end if

	if request.form("opccif")>"" then
		opccif = "1"
	end if

	if request.form("opccontacto")>"" then
		opccontacto = "1"
	end if

	if request.form("opcdomicilio")>"" then
		opcdomicilio = "1"
	end if

	if request.form("opccodigopostal")>"" then
		opccodigopostal = "1"
	end if

	if request.form("opcpoblacion")>"" then
		opcpoblacion = "1"
	end if

	if request.form("opcprovincia")>"" then
		opcprovincia = "1"
	end if

	if request.form("opctelefono")>"" then
		opctelefono = "1"
	end if

	if request.form("opcfalta")>"" then
		opcfalta = "1"
	end if

	if request.form("opcfbaja")>"" then
		opcfbaja = "1"
	end if

	if request.form("opctarifa")>"" then
		opctarifa = "1"
	end if

	if request.form("opccuenta")>"" then
		opccuenta = "1"
	end if

	if request.form("opcformapago")>"" then
		opcformapago = "1"
	end if

	if request.form("opctipopago")>"" then
		opctipopago = "1"
	end if

	if request.form("opcrfinanciero")>"" then
		opcrfinanciero = "1"
	end if

	if request.form("opcrequivalencia")>"" then
		opcrequivalencia = "1"
	end if

	if request.form("opcIRPF")>"" then
		opcIRPF = "1"
	end if

	''if request.form("opcexentoiva")>"" then
	''	opcexentoiva = "1"
	''end if

	if request.form("opclven1")>"" then
		opclven1 = "1"
	end if

	if request.form("opclven2")>"" then
		opclven2 = "1"
	end if

	if request.form("opcentidad")>"" then
		opcentidad = "1"
	end if

	if request.form("opcnumcuenta")>"" then
		opcnumcuenta = "1"
	end if

	if request.form("opcactividad")>"" then
		opcactividad = "1"
	end if

	if request.form("opctproveedor")>"" then
		opctproveedor = "1"
	end if

	if request.form("opcportes")>"" then
		opcportes = "1"
	end if

	if request.form("opcnomcom")>"" then
		opcnomcom= "1"
	end if
	if request.form("opcfax")>"" then
		opcfax= "1"
	end if
	if request.form("opctelmov")>"" then
		opctelmov= "1"
	end if
	if request.form("opcweb")>"" then
		opcweb= "1"
	end if
	if request.form("opcobs")>"" then
		opcobs= "1"
	end if
	if request.form("opcemail")>"" then
		opcemail= "1"
	end if

		if request.form("campo1")>"" then
			campo1 = limpiaCadena(request.form("campo1"))
		else
			campo1 = limpiaCadena(request.querystring("campo1"))
		end if

		if request.form("campo2")>"" then
			campo2 = limpiaCadena(request.form("campo2"))
		else
			campo2 = limpiaCadena(request.querystring("campo2"))
		end if

		if request.form("campo3")>"" then
			campo3 = limpiaCadena(request.form("campo3"))
		else
			campo3 = limpiaCadena(request.querystring("campo3"))
		end if

		if request.form("campo4")>"" then
			campo4 = limpiaCadena(request.form("campo4"))
		else
			campo4 = limpiaCadena(request.querystring("campo4"))
		end if

		if request.form("campo5")>"" then
			campo5 = limpiaCadena(request.form("campo5"))
		else
			campo5 = limpiaCadena(request.querystring("campo5"))
		end if
		if request.form("campo6")>"" then
			campo6 = limpiaCadena(request.form("campo6"))
		else
			campo6 = limpiaCadena(request.querystring("campo6"))
		end if
		if request.form("campo7")>"" then
			campo7 = limpiaCadena(request.form("campo7"))
		else
			campo7 = limpiaCadena(request.querystring("campo7"))
		end if
		if request.form("campo8")>"" then
			campo8 = limpiaCadena(request.form("campo8"))
		else
			campo8 = limpiaCadena(request.querystring("campo8"))
		end if
		if request.form("campo9")>"" then
			campo9 = limpiaCadena(request.form("campo9"))
		else
			campo9 = limpiaCadena(request.querystring("campo9"))
		end if
		if request.form("campo10")>"" then
			campo10 = limpiaCadena(request.form("campo10"))
		else
			campo10 = limpiaCadena(request.querystring("campo10"))
		end if

		if request.form("ver_campo1")>"" then
			ver_campo1 = limpiaCadena(request.form("ver_campo1"))
		else
			ver_campo1 = limpiaCadena(request.querystring("ver_campo1"))
		end if

		if request.form("ver_campo2")>"" then
			ver_campo2 = limpiaCadena(request.form("ver_campo2"))
		else
			ver_campo2 = limpiaCadena(request.querystring("ver_campo2"))
		end if

		if request.form("ver_campo3")>"" then
			ver_campo3 = limpiaCadena(request.form("ver_campo3"))
		else
			ver_campo3 = limpiaCadena(request.querystring("ver_campo3"))
		end if

		if request.form("ver_campo4")>"" then
			ver_campo4 = limpiaCadena(request.form("ver_campo4"))
		else
			ver_campo4 = limpiaCadena(request.querystring("ver_campo4"))
		end if

		if request.form("ver_campo5")>"" then
			ver_campo5 = limpiaCadena(request.form("ver_campo5"))
		else
			ver_campo5 = limpiaCadena(request.querystring("ver_campo5"))
		end if

		if request.form("ver_campo6")>"" then
			ver_campo6 = limpiaCadena(request.form("ver_campo6"))
		else
			ver_campo6 = limpiaCadena(request.querystring("ver_campo6"))
		end if
		if request.form("ver_campo7")>"" then
			ver_campo7 = limpiaCadena(request.form("ver_campo7"))
		else
			ver_campo7 = limpiaCadena(request.querystring("ver_campo7"))
		end if
		if request.form("ver_campo8")>"" then
			ver_campo8 = limpiaCadena(request.form("ver_campo8"))
		else
			ver_campo8 = limpiaCadena(request.querystring("ver_campo8"))
		end if
		if request.form("ver_campo9")>"" then
			ver_campo9 = limpiaCadena(request.form("ver_campo9"))
		else
			ver_campo9 = limpiaCadena(request.querystring("ver_campo9"))
		end if
		if request.form("ver_campo10")>"" then
			ver_campo10 = limpiaCadena(request.form("ver_campo10"))
		else
			ver_campo10 = limpiaCadena(request.querystring("ver_campo10"))
		end if

		numregs	= limpiaCadena(Request.QueryString("numregs"))
		if numregs="" then
			numregs	= limpiaCadena(Request.form("numregs"))
		end if
		if numregs&""="" then numregs=0

	strwhere=""

	Alarma "listado_proveedores.asp"

	set rst = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstPedido = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")


	''ricardo 24-3-2004 si existen campos personalizables con titulo no nulo si saldra la pestaña de campos personalizables
	si_campo_personalizables=0
	rst.open "select ncampo from camposperso where tabla='PROVEEDORES' and titulo is not null and titulo<>'' and ncampo like '" & session("ncliente") & "%'",session("backendlistados"),adOpenKeyset,adLockOptimistic
	if not rst.eof then
		si_campo_personalizables=1
	else
		si_campo_personalizables=0
	end if
	rst.close
	%><input type="hidden" name="si_campo_personalizables" value="<%=EncodeForHtml(si_campo_personalizables)%>"><%

	if si_campo_personalizables=1 then

		tipo_campo_perso1 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "01' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso2 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "02' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso3 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "03' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso4 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "04' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso5 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "05' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso6 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "06' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso7 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "07' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso8 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "08' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso9 =d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "09' and tabla='PROVEEDORES'",session("backendlistados"))
		tipo_campo_perso10=d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "10' and tabla='PROVEEDORES'",session("backendlistados"))

	else
		tipo_campo_perso1=""
		tipo_campo_perso2=""
		tipo_campo_perso3=""
		tipo_campo_perso4=""
		tipo_campo_perso5=""
		tipo_campo_perso6=""
		tipo_campo_perso7=""
		tipo_campo_perso8=""
		tipo_campo_perso9=""
		tipo_campo_perso10=""
	end if

	if mode="select1" then 'Parametros del listado%>

		<table width=96% border='<%=EncodeForHtml(borde)%>' cellspacing="1" cellpadding="1"><%
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitDesdeFechaAlta & " : "
			 	'DrawInputCelda "CELDA","","",10,0,"","fdesde",iif(fdesde>"",fdesde,"01/01/" & year(date))
				'DrawInputCeldaAction "CELDA colspan=2","","",10,0,"","fdesde",iif(fdesde>"",fdesde,"01/01/" & year(date)),"onchange", "PonerFechas('alta')",false
                
                DrawDiv "1", "", ""
                DrawLabel "", "", LitDesdeFechaAlta
                DrawInput "", "", "fdesde", EncodeForHtml(iif(fdesde>"",fdesde,"01/01/" & year(date))), "onchange=""PonerFechas('alta')"" size='10'"
                DrawCalendar "fdesde"
                CloseDiv
                
				'DrawCelda2 "CELDA style='width:175px'", "left", false,LitHastaFechaAlta & " : "
			 	'DrawInputCelda "CELDA","","",10,0,"","fhasta",iif(fhasta>"",fhasta,day(date) & "/" & month(date) & "/" & year(date))
				'DrawInputCeldaAction "CELDA colspan=2","","",10,0,"","fhasta",iif(fhasta>"",fhasta,day(date) & "/" & month(date) & "/" & year(date)),"onchange", "PonerFechas('alta')",false
			
                DrawDiv "1", "", ""
                DrawLabel "", "", LitHastaFechaAlta
                DrawInput "", "", "fhasta", EncodeForHtml(iif(fhasta>"",fhasta,day(date) & "/" & month(date) & "/" & year(date))), "onchange=""PonerFechas('alta')"" size='10'"
                DrawCalendar "fhasta"
                CloseDiv
            
            'CloseFila
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitDesdeFechaBaja & " : "
			 	'DrawInputCelda "CELDA","","",10,0,"","fbdesde",iif(fbdesde>"",fbdesde,"")
				'DrawInputCeldaAction "CELDA colspan=2","","",10,0,"","fbdesde",iif(fbdesde>"",fbdesde,""),"onchange", "PonerFechas('baja')",false

                DrawDiv "1", "", ""
                DrawLabel "", "", LitDesdeFechaBaja
                DrawInput "", "", "fbdesde", EncodeForHtml(iif(fbdesde>"",fbdesde,"")), "onchange=""PonerFechas('baja')"" size='10'"
                DrawCalendar "fbdesde"
                CloseDiv

				'DrawCelda2 "CELDA style='width:175px'", "left", false,LitHastaFechaBaja & " : "
			 	'DrawInputCelda "CELDA","","",10,0,"","fbhasta",iif(fbhasta>"",fbhasta,"")
				'DrawInputCeldaAction "CELDA colspan=2","","",10,0,"","fbhasta",iif(fbhasta>"",fbhasta,""),"onchange", "PonerFechas('baja')",false

                DrawDiv "1", "", ""
                DrawLabel "", "", LitHastaFechaBaja
                DrawInput "", "", "fbhasta", EncodeForHtml(iif(fbhasta>"",fbhasta,"")), "onchange=""PonerFechas('baja')"" size='10'"
                DrawCalendar "fbhasta"
                CloseDiv

			'CloseFila
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:195px'", "left", false, LitRazonSocial & " : "
				'DrawInputCeldaSpan "CELDA colspan=2","","",40,0,"","razon_social","",3

                DrawDiv "1", "", ""
                DrawLabel "", "", LitRazonSocial
                DrawInput "", "", "razon_social", "", "size='40'"
                CloseDiv

			'CloseFila
                'DrawCelda2 "CELDA style='width:175px'", "left", false, LitPoblacion & " : "
			 	'DrawInputCelda "CELDA colspan=2","","",20,0,"","poblacion",""
                DrawDiv "1","",""
                DrawLabel "","",LitPoblacion
                DrawInput "","","poblacion","","size='20'"
                CloseDiv
				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitProvincia & " : "
			 	'DrawInputCelda "CELDA colspan=2","","",20,0,"","provincia",""
                DrawDiv "1","",""
                DrawLabel "","",LitProvincia
                DrawInput "","","provincia","","size='20'"
                CloseDiv
            'CloseFila
			'DrawFila color_blau
			 	'DrawCelda2 "CELDA", "left", false, LitTarifa & " : "
				'rstSelect.open "select codigo,descripcion from tarifas where codigo<>'" & session("ncliente") & "BASE' order by descripcion",session("backendlistados"),adOpenKeyset,adLockOptimistic
				'DrawSelectCelda "CELDA","","",0,"","tarifa",rstSelect,"","codigo","descripcion","",""
				'rstSelect.close

			 	'DrawCelda2 "CELDA style='width:175px'", "left", false, LitFormaPago & " : "
				rstSelect.open "select codigo,descripcion from formas_pago where codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados"),adOpenKeyset,adLockOptimistic
				DrawSelectCelda "CELDA colspan=2",200,"",0,LitFormaPago,"formapago",rstSelect,"","codigo","descripcion","",""
				rstSelect.close
			'CloseFila
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitTactividad & " : "
				rstSelect.open "select codigo,descripcion from tipo_actividad where codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados"),adOpenKeyset,adLockOptimistic
				DrawSelectCelda "CELDA colspan=2",200,"",0,LitTactividad,"actividad",rstSelect,"","codigo","descripcion","",""
				rstSelect.close

				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitTProveedor & " : "
				rstSelect.open "select codigo,descripcion from tipos_entidades where tipo='" & LitPROVEEDOR & "' and codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados"),adOpenKeyset,adLockOptimistic
				DrawSelectCelda "CELDA colspan=2",200,"",0,LitTProveedor,"tipo_proveedor",rstSelect,"","codigo","descripcion","",""
				rstSelect.close

			'CloseFila%>
		<!--</table>
		<table border='<%=borde%>' cellspacing="1" cellpadding="1">--><%
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:240px'", "left", false, LitProveedorBaja & " : "
				DrawCheckCelda "","","",0,LitProveedorBaja,"opcproveedorbaja",""
			'CloseFila
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitSoloDistribuidores & " : "
				DrawCheckCelda "","","",0,LitSoloDistribuidores,"solodistribuidores",""
			'CloseFila
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:210px'", "left", false, LitMostrarContactos & " : "
				DrawCheckCelda "","","",0,LitMostrarContactos,"mostrarcontactos",""
			'CloseFila
			%>
		<!--</table>-->
		<!--<tr><td width="100%" colspan="10"><hr/></td></tr>-->
        <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><hr/></div>
        <%
		if si_campo_personalizables=1 then
			%><!--<table border='<%=borde%>' cellspacing="1" cellpadding="1">--><%
				'DrawFila color_fondo
				'	DrawCelda2 "ENCABEZADOL", "left", false, LitCampPersoPro
				'CloseFila
			%><h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitCampPersoPro%></h6>
            <!--</table>--><%
		end if
		rst.cursorlocation=3
		rst.open "select * from camposperso where tabla='PROVEEDORES' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("backendlistados")
		if not rst.eof then
			num_campos_existen=rst.recordcount
			%><!--<table border='<%=borde%>' cellspacing="1" cellpadding="1">--><%
				'DrawFila ""
					num_campo=1
					num_campo2=1
					num_puestos=0
					while not rst.eof
						'if num_campo2>1 and ((num_campo2-1) mod 2)=0 then
						if (num_puestos mod 2)=0 then
							'CloseFila
							'DrawFila ""
						end if
						if rst("titulo") & "">"" then
							if ((num_puestos-1) mod 2)=0 then
								'DrawCelda "CELDA7 style='width:5%'","","",0,"&nbsp;"
							end if
							num_puestos=num_puestos+1
							%><input type="hidden" name="<%="si_campo" & num_campo%>" value="1"><%
							'DrawCelda "CELDA style='width:200px'","","",0,rst("titulo") & " : "
                            DrawDiv "1", "", ""
                            DrawLabel "", "", EncodeForHtml(null_s(rst("titulo")))

							valor_campo_perso=""
							if rst("tipo")=1 then
								if isNumeric(rst("tamany")) then
									tamany= EncodeForHtml(null_s(rst("tamany")))
								else
									tamany=1
								end if
								%><!--<td class="CELDA" style='width:165px'>-->
									<input type="text" class="" align="left" name="<%="campo" & num_campo%>" size="30" maxlength="<%=tamany%>" value="<%=valor_campo_perso%>">
								<!--</td>--><%
							elseif rst("tipo")=2 then
								'DrawCheckCelda "CELDALEFT","","",0,"","campo" & num_campo,iif(valor_campo_perso="on",-1,0)
                                DrawCheck "CELDALEFT", "", "campo" & num_campo, iif(valor_campo_perso="on",-1,0)
							elseif rst("tipo")=3 then
								num_campo_str=cstr(num_campo)
								if len(num_campo_str)=1 then
									num_campo_str="0" & num_campo_str
								end if
								strSelListVal="select ndetlista,valor from campospersolista where tabla='PROVEEDORES' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
								rstAux.open strSelListVal,session("backendlistados"),adOpenKeyset, adLockOptimistic
								'sera este
								'DrawSelectCelda "CELDA7","200","",0,"","campo" & num_campo,rstAux,valor_campo_perso,"ncampo","titulo","",""
								'o bien este
								%><!--<td class="CELDA7" style='width:200px'>-->
									<select class="" name="campo<%=num_campo%>" >
										<%
										encontrado=0
										while not rstAux.eof
											if valor_campo_perso & "">"" and isnumeric(valor_campo_perso) then
												valor_campo_perso_aux=cint(valor_campo_perso)
											else
												valor_campo_perso_aux=0
											end if
											if valor_campo_perso_aux=cint(rstAux("ndetlista")) then
												texto_selected="selected"
												if encontrado=0 then encontrado=1
											else
												texto_selected=""
											end if
											%>
											<option value="<%=EncodeForHtml(null_s(rstAux("ndetlista")))%>"  <%=texto_selected%> ><%=EncodeForHtml(null_s(rstAux("valor")))%></option>
											<%rstAux.movenext
										wend%>
										<option <%=iif(encontrado=1,"","selected")%> value=""></option>
									</select>
								<!--</td>--><%
								rstAux.close
							elseif rst("tipo")=4 then
								if isNumeric(rst("tamany")) then
									tamany= EncodeForHtml(null_s(rst("tamany")))
								else
									tamany=1
								end if
								'DrawInputCelda "CELDA7 style='width:180px' maxlength='" & tamany & "'","","",35,0,"","campo" & num_campo,valor_campo_perso
								%><!--<td class="CELDA7" style='width:180px' align="left">-->
									<input type="text" class="" name="<%="campo" & num_campo%>" size="35" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>">
								<!--</td>--><%
							elseif rst("tipo")=5 then
								if isNumeric(rst("tamany")) then
									tamany= EncodeForHtml(null_s(rst("tamany")))
								else
									tamany=1
								end if
								'DrawInputCelda "CELDA7 style='width:180px' maxlength='" & tamany & "'","","",35,0,"","campo" & num_campo,valor_campo_perso
								%><!--<td class="CELDA7" style='width:180px' align="left">-->
									<input type="text" class="" name="<%="campo" & num_campo%>" size="35" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>">
								<!--</td>--><%
							end if

                            CloseDiv
						else
							%><input type="hidden" name="<%="si_campo" & num_campo%>" value="0"><%
							%><input type="hidden" name="campo<%=num_campo%>" value=""><%
						end if
						%><input type="hidden" name="tipo_campo<%=num_campo%>" value="<%=EncodeForHtml(rst("tipo"))%>"><%
						%><input type="hidden" name="titulo_campo<%=num_campo%>" value="<%=EncodeForHtml(null_s(rst("titulo")))%>"><%
						rst.movenext
						num_campo=num_campo+1
						if not rst.eof then
							if rst("titulo") & "">"" then
								num_campo2=num_campo2+1
							end if
						end if
					wend
				'CloseFila
			%><!--</table>--><%
			%>
            <input type="hidden" name="num_puestos" value="<%=EncodeForHtml(num_puestos)%>"><%
			%><input type="hidden" name="num_campos" value="<%=EncodeForHtml(num_campos_existen)%>"><%
		else
			%><input type="hidden" name="num_puestos" value="0"><%
			%><input type="hidden" name="num_campos" value="0"><%
		end if
		rst.close

		if si_campo_personalizables=1 then%>
			<!--<tr><td width="100%" colspan="10"><hr/></td></tr>-->
            <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><hr/></div>
		<%end if%>
		<!--<table border='<%=borde%>' cellspacing="1" cellpadding="1">--><%
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitOrdenar & " : "
                DrawDiv "1", "", ""
                DrawLabel "", "", LitOrdenar
				%><!--<td>-->
					<select class="" name="ordenar">
						<option value="<%=LitFechaAlta%>"><%=LitFechaAlta%></option>
						<option value="<%=LitFechaBaja%>"><%=LitFechaBaja%></option>
						<option selected value="<%=LitNumProveedor%>"><%=LitNumProveedor%></option>
						<option value="<%=LitRazonSocial%>"><%=LitRazonSocial%></option>
					</select>
				<!--</td>--><%
                CloseDiv
			'CloseFila
			'DrawFila color_blau
				'DrawCelda2 "CELDA style='width:175px'", "left", false, LitApaisado & " : "
				DrawCheckCelda "","","",0,LitApaisado,"apaisado",""
			'CloseFila%>
		<!--</table>-->

		<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><hr/></div>

		<!--<table border='<%=borde%>' cellspacing="1" cellpadding="1">--><%
			'DrawFila color_fondo
			'	DrawCelda2 "ENCABEZADOL", "left", false, LitCamposOpcionales
			'CloseFila%>
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitCamposOpcionales%></h6><%
                
                DrawDiv "7", "",""
                      DrawDiv "3-sub", "background-color: #eae7e3", ""
                        %><label class="ENCABEZADOL" style="text-align:left"><%=LitDatosGenerales%></label>                    
                        <%CloseDiv
						'DrawCelda2 "CELDA", "left", false, LitCIF
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitCIF, "opccif", "", EncodeForHtml(cstr(forma_pago))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitContacto1
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitContacto1, "opccontacto", "", EncodeForHtml(cstr(tipo_pago))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitNomComProv
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitNomComProv, "opcnomcom", "", EncodeForHtml(cstr(opcnomcom))

                            EligeCelda "check-listado", "edit", "", "", "", 0, LitDomicilio, "opcdomicilio", "", EncodeForHtml(cstr(portes))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitCodPostal
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitCodPostal, "opccodigopostal", "", EncodeForHtml(cstr(responsable))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitFax
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitFax, "opcfax", "", EncodeForHtml(cstr(opcfax))
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitPoblacion
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitPoblacion, "opcpoblacion", "", EncodeForHtml(cstr(envio))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitProvincia
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitProvincia, "opcprovincia", "", EncodeForHtml(cstr(num_envio))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitTel2
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitTel2, "opctelmov", "", EncodeForHtml(cstr(opctelmov))
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitTelefono
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitTelefono, "opctelefono", "", EncodeForHtml(cstr(transportista))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitFechaAlta
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitFechaAlta, "opcfalta", "", EncodeForHtml(cstr(bultos))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, ucase(LitWEB)
                            EligeCelda "check-listado", "edit", "", "", "", 0, ucase(LitWEB), "opcweb", "", EncodeForHtml(cstr(opcweb))
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitFechaBaja
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitFechaBaja, "opcfbaja", "", ""
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitEmaProv
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitEmaProv, "opcemail", "", EncodeForHtml(cstr(opcemail))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitObservaciones
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitObservaciones, "opcobs", "", EncodeForHtml(cstr(opcobs))
					'CloseFila
                CloseDiv
                DrawDiv "7", "",""
                      DrawDiv "3-sub", "background-color: #eae7e3", ""
                        %><label class="ENCABEZADOL" style="text-align:left"><%=LitDatosComerciales%></label>                    
                        <%CloseDiv

                        EligeCelda "check-listado", "edit", "", "", "", 0, LitFormaPago, "opcformapago", "", EncodeForHtml(cstr(forma_pago))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitTipoPago
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitTipoPago, "opctipopago", "", EncodeForHtml(cstr(tipo_pago))
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitRFinan
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitRFinan, "opcrfinanciero", "", EncodeForHtml(cstr(forma_pago))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitREquiv
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitREquiv, "opcrequivalencia", "", EncodeForHtml(cstr(tipo_pago))
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitIRPF1
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitIRPF1, "opcIRPF", "", EncodeForHtml(cstr(forma_pago))
						'DrawCelda "CELDA","10%","",0," "
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitCCont
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitCCont, "opccuenta", "", EncodeForHtml(cstr(tipo_pago))
					'CloseFila
                CloseDiv
                DrawDiv "7", "",""
                      DrawDiv "3-sub", "background-color: #eae7e3", ""
                        %><label class="ENCABEZADOL" style="text-align:left"><%=LitOtrosDatos%></label>                    
                        <%CloseDiv
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitEntidad
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitEntidad, "opcentidad", "", EncodeForHtml(cstr(forma_pago))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitNumCuenta
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitNumCuenta, "opcnumcuenta", "", EncodeForHtml(cstr(tipo_pago))
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitTactividad
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitTactividad, "opcactividad", "", EncodeForHtml(cstr(forma_pago))
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitTProveedor
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitTProveedor, "opctproveedor", "", EncodeForHtml(cstr(forma_pago))
					'CloseFila
					'DrawFila color_blau
						'DrawCelda2 "CELDA", "left", false, LitPortes
                            EligeCelda "check-listado", "edit", "", "", "", 0, LitPortes, "opcportes", "", EncodeForHtml(cstr(tipo_pago))
					'CloseFila
                CloseDiv
				if si_campo_personalizables=1 then
                    DrawDiv "7", "",""
                      DrawDiv "3-sub", "background-color: #eae7e3", ""
                        %><label class="ENCABEZADOL" style="text-align:left"><%=LitCampPersoPro%></label>                    
                        <%CloseDiv
							rst.open "select * from camposperso where tabla='PROVEEDORES' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("backendlistados"),adOpenKeyset, adLockOptimistic
							if not rst.eof then
								'DrawFila ""
								num_campo=1
								num_campo2=1
								while not rst.eof
									if num_campo2>1 and ((num_campo2-1) mod 2)=0 then
										'CloseFila
										'DrawFila ""
									end if
									if rst("titulo") & "">"" then
										'DrawCelda "CELDA","","",0,rst("titulo") & " : "
										DrawCheckCelda "","","",0, EncodeForHtml(null_s(rst("titulo"))),"ver_campo" & num_campo,""
										'DrawCelda "CELDA","10%","",0," "
									else
										%><!--<input type="hidden" name="ver_campo<%=num_campo%>" value="">--><%
										'DrawCheckCelda "CELDA style='display:none'","","",0,"","ver_campo" & num_campo,""
                                        DrawCheck "", "display:none", "ver_campo" & num_campo, ""
									end if
									rst.movenext
									num_campo=num_campo+1
									if not rst.eof then
										if rst("titulo") & "">"" then
											num_campo2=num_campo2+1
										end if
									end if
								wend
							'CloseFila
						end if
						rst.close
					CloseDiv
              end if
     elseif mode="imp" then

%>
		<input type="hidden" name="fdesde" value="<%=EncodeForHtml(fdesde)%>">
		<input type="hidden" name="fhasta" value="<%=EncodeForHtml(fhasta)%>">
		<input type="hidden" name="fbdesde" value="<%=EncodeForHtml(fbdesde)%>">
		<input type="hidden" name="fbhasta" value="<%=EncodeForHtml(fbhasta)%>">
		<input type="hidden" name="razon_social" value="<%=EncodeForHtml(razon_social)%>">
		<input type="hidden" name="poblacion" value="<%=EncodeForHtml(poblacion)%>">
		<input type="hidden" name="provincia" value="<%=EncodeForHtml(provincia)%>">
		<input type="hidden" name="tarifa" value="<%=EncodeForHtml(tarifa)%>">
		<input type="hidden" name="formapago" value="<%=EncodeForHtml(formapago)%>">
		<input type="hidden" name="actividad" value="<%=EncodeForHtml(actividad)%>">
		<input type="hidden" name="tipo_proveedor" value="<%=EncodeForHtml(tipo_proveedor)%>">
		<input type="hidden" name="ordenar" value="<%=EncodeForHtml(ordenar)%>">
		<input type="hidden" name="campo1" value="<%=EncodeForHtml(campo1)%>">
		<input type="hidden" name="campo2" value="<%=EncodeForHtml(campo2)%>">
		<input type="hidden" name="campo3" value="<%=EncodeForHtml(campo3)%>">
		<input type="hidden" name="campo4" value="<%=EncodeForHtml(campo4)%>">
		<input type="hidden" name="campo5" value="<%=EncodeForHtml(campo5)%>">
		<input type="hidden" name="campo6" value="<%=EncodeForHtml(campo6)%>">
		<input type="hidden" name="campo7" value="<%=EncodeForHtml(campo7)%>">
		<input type="hidden" name="campo8" value="<%=EncodeForHtml(campo8)%>">
		<input type="hidden" name="campo9" value="<%=EncodeForHtml(campo9)%>">
		<input type="hidden" name="campo10" value="<%=EncodeForHtml(campo10)%>">

		<input type="hidden" name="opcproveedorbaja" value="<%=EncodeForHtml(opcproveedorbaja)%>">
		<input type="hidden" name="solodistribuidores" value="<%=EncodeForHtml(solodistribuidores)%>">
		<input type="hidden" name="mostrarcontactos" value="<%=EncodeForHtml(mostrarcontactos)%>">
		<input type="hidden" name="opccif" value="<%=EncodeForHtml(opccif)%>">
		<input type="hidden" name="opccontacto" value="<%=EncodeForHtml(opccontacto)%>">
		<input type="hidden" name="opcdomicilio" value="<%=EncodeForHtml(opcdomicilio)%>">
		<input type="hidden" name="opccodigopostal" value="<%=EncodeForHtml(opccodigopostal)%>">
		<input type="hidden" name="opcpoblacion" value="<%=EncodeForHtml(opcpoblacion)%>">
		<input type="hidden" name="opcprovincia" value="<%=EncodeForHtml(opcprovincia)%>">
		<input type="hidden" name="opctelefono" value="<%=EncodeForHtml(opctelefono)%>">
		<input type="hidden" name="opcfalta" value="<%=EncodeForHtml(opcfalta)%>">
		<input type="hidden" name="opcfbaja" value="<%=EncodeForHtml(opcfbaja)%>">
		<input type="hidden" name="opctarifa" value="<%=EncodeForHtml(opctarifa)%>">
		<input type="hidden" name="opccuenta" value="<%=EncodeForHtml(opccuenta)%>">
		<input type="hidden" name="opcformapago" value="<%=EncodeForHtml(opcformapago)%>">
		<input type="hidden" name="opctipopago" value="<%=EncodeForHtml(opctipopago)%>">
		<input type="hidden" name="opcrfinanciero" value="<%=EncodeForHtml(opcrfinanciero)%>">
		<input type="hidden" name="opcrequivalencia" value="<%=EncodeForHtml(opcrequivalencia)%>">
		<input type="hidden" name="opcIRPF" value="<%=EncodeForHtml(opcIRPF)%>">
<!--	<input type="hidden" name="opcexentoiva" value="<%=EncodeForHtml(opcexentoiva)%>">-->
		<input type="hidden" name="opclven1" value="<%=EncodeForHtml(opclven1)%>">
		<input type="hidden" name="opclven2" value="<%=EncodeForHtml(opclven2)%>">
		<input type="hidden" name="opcentidad" value="<%=EncodeForHtml(opcentidad)%>">
		<input type="hidden" name="opcnumcuenta" value="<%=EncodeForHtml(opcnumcuenta)%>">
		<input type="hidden" name="opcactividad" value="<%=EncodeForHtml(opcactividad)%>">
		<input type="hidden" name="opctproveedor" value="<%=EncodeForHtml(opctproveedor)%>">
		<input type="hidden" name="opcportes" value="<%=EncodeForHtml(opcportes)%>">
		<input type="hidden" name="opcnomcom" value="<%=EncodeForHtml(opcnomcom)%>">
		<input type="hidden" name="opcfax" value="<%=EncodeForHtml(opcfax)%>">
		<input type="hidden" name="opctelmov" value="<%=EncodeForHtml(opctelmov)%>">
		<input type="hidden" name="opcweb" value="<%=EncodeForHtml(opcweb)%>">
		<input type="hidden" name="opcobs" value="<%=EncodeForHtml(opcobs)%>">
		<input type="hidden" name="opcemail" value="<%=EncodeForHtml(opcemail)%>">
		<input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>">
		<input type="hidden" name="ver_campo1" value="<%=EncodeForHtml(ver_campo1)%>">
		<input type="hidden" name="ver_campo2" value="<%=EncodeForHtml(ver_campo2)%>">
		<input type="hidden" name="ver_campo3" value="<%=EncodeForHtml(ver_campo3)%>">
		<input type="hidden" name="ver_campo4" value="<%=EncodeForHtml(ver_campo4)%>">
		<input type="hidden" name="ver_campo5" value="<%=EncodeForHtml(ver_campo5)%>">
		<input type="hidden" name="ver_campo6" value="<%=EncodeForHtml(ver_campo6)%>">
		<input type="hidden" name="ver_campo7" value="<%=EncodeForHtml(ver_campo7)%>">
		<input type="hidden" name="ver_campo8" value="<%=EncodeForHtml(ver_campo8)%>">
		<input type="hidden" name="ver_campo9" value="<%=EncodeForHtml(ver_campo9)%>">
		<input type="hidden" name="ver_campo10" value="<%=EncodeForHtml(ver_campo10)%>">


<%
        
  		total_valor_general = 0
		total_pendiente_general = 0
		MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='107'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='107'", DSNIlion)
			%><input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>
			<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'>
			<input type='hidden' name='maxmb' value='<%=EncodeForHtml(MB)%>'><%

		VinculosPagina(MostrarProveedores)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

		strwhere="where"
		if fdesde > "" then
			strwhere=strwhere & " falta>='" & fdesde & "' and"
			%><font class=cab><b><%=LitDesdeFechaAlta%>&nbsp; : </b></font><font class=cab><%=EncodeForHtml(fdesde)%></font><br/><%
		end if
		if fhasta > "" then
			strwhere=strwhere & " falta<='" & fhasta & "' and"
			%><font class=cab><b><%=LitHastaFechaAlta%>&nbsp; : </b></font><font class=cab><%=EncodeForHtml(fhasta)%></font><br/><%
		end if
		if fbdesde > "" then
			strwhere=strwhere & " fbaja>='" & fbdesde & "' and"
			%><font class=cab><b><%=LitDesdeFechaBaja%>&nbsp; : </b></font><font class=cab><%=EncodeForHtml(fbdesde)%></font><br/><%
		end if
		if fbhasta > "" then
			strwhere=strwhere & " fbaja<='" & fbhasta & "' and"
			%><font class=cab><b><%=LitHastaFechaBaja%>&nbsp; : </b></font><font class=cab><%=EncodeForHtml(fbhasta)%></font><br/><%
		end if
		if razon_social > "" then
			strwhere=strwhere & " razon_social like '%" & razon_social & "%' and"
			%><font class=cab><b><%=LitRazonSocial%>&nbsp; : </b></font><font class=cab><%=EncodeForHtml(razon_social)%></font><br/><%
		end if
		if poblacion > "" then
			strwhere=strwhere & " poblacion like '%" & poblacion & "%' and"
			%><font class=cab><b><%=LitPoblacion%>&nbsp; : </b></font><font class=cab><%=EncodeForHtml(poblacion)%></font><br/><%
		end if
		if provincia > "" then
			strwhere=strwhere & " provincia like '%" & provincia & "%' and"
			%><font class=cab><b><%=LitProvincia%>&nbsp; : </b></font><font class=cab><%=EncodeForHtml(provincia)%></font><br/><%
		end if
		if tarifa > "" then
			strwhere=strwhere & " tarifa='" & tarifa & "' and"
			%><font class=cab><b><%=LitTarifa%> :&nbsp;</b></font><font class=cab><%=EncodeForHtml(d_lookup("descripcion","tarifas","codigo like '" & session("ncliente") & "%' and codigo='" & tarifa & "'",session("backendlistados")))%></font><br/><%
		end if
		if formapago > "" then
			strwhere=strwhere & " forma_pago='" & formapago & "' and"
			%><font class=cab><b><%=LitFormaPago%> :&nbsp;</b></font><font class=cab><%=EncodeForHtml(d_lookup("descripcion","formas_pago","codigo like '" & session("ncliente") & "%' and codigo='" & formapago & "'",session("backendlistados")))%></font><br/><%
		end if
		if actividad > "" then
			strwhere=strwhere & " tactividad='" & actividad & "' and"
			%><font class=cab><b><%=LitTactividad%> :&nbsp;</b></font><font class=cab><%=EncodeForHtml(d_lookup("descripcion","tipo_actividad","codigo like '" & session("ncliente") & "%' and codigo='" & actividad & "'",session("backendlistados")))%></font><br/><%
		end if
		if tipo_proveedor > "" then
			strwhere=strwhere & " tipo_proveedor='" & tipo_proveedor & "' and"
			%><font class=cab><b><%=LitTProveedor%> :&nbsp;</b></font><font class=cab><%=EncodeForHtml(d_lookup("descripcion","tipos_entidades","codigo like '" & session("ncliente") & "%' and codigo='" & tipo_proveedor & "'",session("backendlistados")))%></font><br/><%
		end if
		if opcproveedorbaja > "" then
			strwhere=strwhere & " fbaja is null and"
			%><font class=cab><b><%=LitProveedorBaja%></b></font><br/><%
		else
			'<font class=cab><b>=LitProveedorBaja</b></font><br/>
		end if
		if campo1>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "01' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso1=2 then
				valor_contiene=""
				if ucase(campo1)="ON" or cstr(null_s(campo1))="1" then
					valor_campo1="Sí"
					valor_campo1_where="=1"
				else
					valor_campo1="No"
					valor_campo1_where="=0"
				end if
			elseif tipo_campo_perso1=3 then
				if campo1 & "">"" then
					valor_campo1=d_lookup("valor","campospersolista","ndetlista=" & campo1 & " and ncampo='" & session("ncliente") & "01' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo1=""
				end if
				valor_contiene=LitCampPersoPro
				valor_campo1_where=" like '%" & campo1 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo1=campo1
				valor_campo1_where=" like '%" & campo1 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo1)%></font><br/><%
			strwhere = strwhere & " campo01 " & valor_campo1_where & " and"
		end if

		if campo2>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "02' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso2=2 then
				valor_contiene=""
				if ucase(campo2)="ON" or cstr(null_s(campo2))="1" then
					valor_campo2="Sí"
					valor_campo2_where="=1"
				else
					valor_campo2="No"
					valor_campo2_where="=0"
				end if
			elseif tipo_campo_perso2=3 then
				if campo2 & "">"" then
					valor_campo2=d_lookup("valor","campospersolista","ndetlista=" & campo2 & " and ncampo='" & session("ncliente") & "02' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo2=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo2_where=" like '%" & campo2 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo2=campo2
				valor_campo2_where=" like '%" & campo2 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo2)%></font><br/><%
			strwhere = strwhere & " campo02 " & valor_campo2_where & " and"
		end if

		if campo3>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "03' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso3=2 then
				valor_contiene=""
				if ucase(campo3)="ON" or cstr(null_s(campo3))="1" then
					valor_campo3="Sí"
					valor_campo3_where="=1"
				else
					valor_campo3="No"
					valor_campo3_where="=0"
				end if
			elseif tipo_campo_perso3=3 then
				if campo3 & "">"" then
					valor_campo3=d_lookup("valor","campospersolista","ndetlista=" & campo3 & " and ncampo='" & session("ncliente") & "03' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo3=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo3_where=" like '%" & campo3 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo3=campo3
				valor_campo3_where=" like '%" & campo3 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo3)%></font><br/><%
			strwhere = strwhere & " campo03 " & valor_campo3_where & " and"
		end if

		if campo4>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "04' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso4=2 then
				valor_contiene=""
				if ucase(campo4)="ON" or cstr(null_s(campo4))="1" then
					valor_campo4="Sí"
					valor_campo4_where="=1"
				else
					valor_campo4="No"
					valor_campo4_where="=0"
				end if
			elseif tipo_campo_perso4=3 then
				if campo4 & "">"" then
					valor_campo4=d_lookup("valor","campospersolista","ndetlista=" & campo4 & " and ncampo='" & session("ncliente") & "04' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo4=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo4_where=" like '%" & campo4 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo4=campo4
				valor_campo4_where=" like '%" & campo4 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo4)%></font><br/><%
			strwhere = strwhere & " campo04 " & valor_campo4_where & " and"
		end if

		if campo5>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "05' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso5=2 then
				valor_contiene=""
				if ucase(campo5)="ON" or cstr(null_s(campo5))="1" then
					valor_campo5="Sí"
					valor_campo5_where="=1"
				else
					valor_campo5="No"
					valor_campo5_where="=0"
				end if
			elseif tipo_campo_perso5=3 then
				if campo5 & "">"" then
					valor_campo5=d_lookup("valor","campospersolista","ndetlista=" & campo5 & " and ncampo='" & session("ncliente") & "05' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo5=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo5_where=" like '%" & campo5 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo5=campo5
				valor_campo5_where=" like '%" & campo5 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo5)%></font><br/><%
			strwhere = strwhere & " campo05 " & valor_campo5_where & " and"
		end if

		if campo6>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "06' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso6=2 then
				valor_contiene=""
				if ucase(campo6)="ON" or cstr(null_s(campo6))="1" then
					valor_campo6="Sí"
					valor_campo6_where="=1"
				else
					valor_campo6="No"
					valor_campo6_where="=0"
				end if
			elseif tipo_campo_perso6=3 then
				if campo6 & "">"" then
					valor_campo6=d_lookup("valor","campospersolista","ndetlista=" & campo6 & " and ncampo='" & session("ncliente") & "06' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo6=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo6_where=" like '%" & campo6 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo6=campo6
				valor_campo6_where=" like '%" & campo6 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo6)%></font><br/><%
			strwhere = strwhere & " campo06 " & valor_campo6_where & " and"
		end if
		if campo7>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "07' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso7=2 then
				valor_contiene=""
				if ucase(campo7)="ON" or cstr(null_s(campo7))="1" then
					valor_campo7="Sí"
					valor_campo7_where="=1"
				else
					valor_campo7="No"
					valor_campo7_where="=0"
				end if
			elseif tipo_campo_perso7=3 then
				if campo7 & "">"" then
					valor_campo7=d_lookup("valor","campospersolista","ndetlista=" & campo7 & " and ncampo='" & session("ncliente") & "07' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo7=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo7_where=" like '%" & campo7 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo7=campo7
				valor_campo7_where=" like '%" & campo7 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo7)%></font><br/><%
			strwhere = strwhere & " campo07 " & valor_campo7_where & " and"
		end if
		if campo8>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "08' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso8=2 then
				valor_contiene=""
				if ucase(campo8)="ON" or cstr(null_s(campo8))="1" then
					valor_campo8="Sí"
					valor_campo8_where="=1"
				else
					valor_campo8="No"
					valor_campo8_where="=0"
				end if
			elseif tipo_campo_perso8=3 then
				if campo8 & "">"" then
					valor_campo8=d_lookup("valor","campospersolista","ndetlista=" & campo8 & " and ncampo='" & session("ncliente") & "08' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo8=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo8_where=" like '%" & campo8 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo8=campo8
				valor_campo8_where=" like '%" & campo8 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo8)%></font><br/><%
			strwhere = strwhere & " campo08 " & valor_campo8_where & " and"
		end if
		if campo9>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "09' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso9=2 then
				valor_contiene=""
				if ucase(campo9)="ON" or cstr(null_s(campo9))="1" then
					valor_campo9="Sí"
					valor_campo9_where="=1"
				else
					valor_campo9="No"
					valor_campo9_where="=0"
				end if
			elseif tipo_campo_perso9=3 then
				if campo9 & "">"" then
					valor_campo9=d_lookup("valor","campospersolista","ndetlista=" & campo9 & " and ncampo='" & session("ncliente") & "09' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo9=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo9_where=" like '%" & campo9 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo9=campo9
				valor_campo9_where=" like '%" & campo9 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo9)%></font><br/><%
			strwhere = strwhere & " campo09 " & valor_campo9_where & " and"
		end if
		if campo10>"" then
			titulo=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & "10' and tabla='PROVEEDORES'",session("backendlistados"))
			if tipo_campo_perso10=2 then
				valor_contiene=""
				if ucase(campo10)="ON" or cstr(null_s(campo10))="1" then
					valor_campo10="Sí"
					valor_campo10_where="=1"
				else
					valor_campo10="No"
					valor_campo10_where="=0"
				end if
			elseif tipo_campo_perso10=3 then
				if campo10 & "">"" then
					valor_campo10=d_lookup("valor","campospersolista","ndetlista=" & campo10 & " and ncampo='" & session("ncliente") & "10' and tabla='PROVEEDORES'",session("backendlistados"))
				else
					valor_campo10=""
				end if
				valor_contiene=LitCampPersoParCont
				valor_campo10_where=" like '%" & campo10 & "%'"
			else
				valor_contiene=LitCampPersoParCont
				valor_campo10=campo10
				valor_campo10_where=" like '%" & campo10 & "%'"
			end if
			%><font class='CELDA'><b><%=EncodeForHtml(titulo)%><%=EncodeForHtml(valor_contiene)%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(valor_campo10)%></font><br/><%
			strwhere = strwhere & " campo10 " & valor_campo10_where & " and"
		end if

		if solodistribuidores> "" then
			lista=""
			rstAux.open "select nproveedor from distribuidores with(nolock) where nproveedor like '" & session("ncliente") & "%'",session("backendlistados"),adOpenKeyset,adLockOptimistic
			if not rstAux.eof then
				while not rstAux.eof
					lista=lista & "'" & rstAux("nproveedor") & "',"
					rstAux.movenext
				wend
				lista=mid(lista,1,len(lista)-len(","))
			else
				lista=""
			end if
			rstAux.close
			if lista>"" then
				strwhere=strwhere & " proveedores.nproveedor in (" & lista & ") and"
			else
				strwhere=strwhere & " proveedores.nproveedor in ('xxxxx') and"
			end if
			%><font class=cab><b><%=LitSoloDistribuidores%></b></font><br/><%
		else

		end if

		strwhere=strwhere & " domicilios.pertenece=proveedores.nproveedor and domicilios.tipo_domicilio='PRINCIPAL_PROV' "

''ricardo 21/8/2003 se cambia esta linea
		'''strwhere=strwhere & " and domicilios.codigo=(select max(codigo) from domicilios where tipo_domicilio='PRINCIPAL_PROV' and domicilios.pertenece=proveedores.nproveedor) and"
'''por esta otra, ya que el listado puede llegar a dar tiempo de espera agotado si se tiene unos 538 proveedores
		strwhere=strwhere & " and domicilios.codigo=proveedores.dir_principal and"
'''''''''

		strwhere=strwhere & " proveedores.nproveedor like '" & session("ncliente") & "%' and"

		%><hr/><%
		if strwhere="where" then
			strwhere=""
			condicion=""
		else
			strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
			condicion=right(strwhere,len(strwhere)-5)
		end if

		select case ucase(ordenar)
			case "NUM.PROVEEDOR"
				strwhere=strwhere & " order by proveedores.nproveedor"
			case "RAZÓN SOCIAL"
				strwhere=strwhere & " order by razon_social"
			case "F.ALTA"
				strwhere=strwhere & " order by falta,razon_social"
			case "F.BAJA"
				strwhere=strwhere & " order by fbaja,razon_social"
		end select
		seleccion="select proveedores.*,domicilios.* "
		if mostrarcontactos>"" then
			seleccion=seleccion & ",contactos_pro.ncontacto,contactos_pro.nombre as nomcontact,contactos_pro.domicilio as domiciliocontac,contactos_pro.movil "
		end if
		if opctarifa>"" then
			seleccion=seleccion & ",tarifas.descripcion as NomTarifa "
		end if
		if opcformapago>"" then
			seleccion=seleccion & ",formas_pago.descripcion as NomFpago "
		end if
		if opctipopago>"" then
			seleccion=seleccion & ",tipo_pago.descripcion as NomTpago "
		end if
		if opcactividad>"" then
			seleccion=seleccion & ",tipo_actividad.descripcion as NomActividad "
		end if
		if opctproveedor>"" then
			seleccion=seleccion & ",tipos_entidades.descripcion as NomTProveedor "
		end if
		if ver_campo1="on" then
			seleccion=seleccion & ",campo01"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo2="on" then
			seleccion=seleccion & ",campo02"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo3="on" then
			seleccion=seleccion & ",campo03"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo4="on" then
			seleccion=seleccion & ",campo04"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo5="on" then
			seleccion=seleccion & ",campo05"
		else
			seleccion=seleccion & ",NULL"
		end if

		if ver_campo6="on" then
			seleccion=seleccion & ",campo06"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo7="on" then
			seleccion=seleccion & ",campo07"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo8="on" then
			seleccion=seleccion & ",campo08"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo9="on" then
			seleccion=seleccion & ",campo09"
		else
			seleccion=seleccion & ",NULL"
		end if
		if ver_campo10="on" then
			seleccion=seleccion & ",campo10"
		else
			seleccion=seleccion & ",NULL"
		end if

		from=" from domicilios with(nolock), proveedores with(nolock) "
		if mostrarcontactos>"" then
			from =from  & " LEFT OUTER JOIN contactos_pro with(nolock) ON proveedores.nproveedor = contactos_pro.nproveedor "
		end if
		if opctarifa>"" then
			from=from & " LEFT OUTER JOIN tarifas with(nolock) ON proveedores.tarifa = tarifas.codigo "
		end if
		if opcformapago>"" then
			from=from & " LEFT OUTER JOIN formas_pago with(nolock) ON proveedores.forma_pago = formas_pago.codigo "
		end if
		if opctipopago>"" then
			from=from & " LEFT OUTER JOIN tipo_pago with(nolock) ON proveedores.tipo_pago = tipo_pago.codigo "
		end if
		if opcactividad>"" then
			from=from & " LEFT OUTER JOIN tipo_actividad with(nolock) ON proveedores.tactividad = tipo_actividad.codigo "
		end if
		if opctproveedor>"" then
			from=from & " LEFT OUTER JOIN tipos_entidades with(nolock) ON proveedores.tipo_proveedor = tipos_entidades.codigo "
		end if

		rstPedido.cursorlocation=3
		rstPedido.Open seleccion & from & strwhere, session("backendlistados")
		%><input type="hidden" name="NumRegs" value="<%=EncodeForHtml(rstPedido.Recordcount)%>"><%
		if rstPedido.EOF then
			rstPedido.Close
			%><input type="hidden" name="nRegsImp" value="0"><%
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitMsgDatosNoExiste%>");
			      parent.window.frames["botones"].document.location = "listado_proveedores_bt.asp?mode=select1";
				document.location="listado_proveedores.asp?mode=select1";
			</script><%
		else
			'Calculos de páginas--------------------------
		   	lote=limpiaCadena(Request.QueryString("lote"))
		   	if lote="" then
				lote=1
		   	end if
		   	sentido=limpiaCadena(Request.QueryString("sentido"))

		   	lotes=rstPedido.RecordCount/MAXPAGINA
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

		   	rstPedido.PageSize=MAXPAGINA
		   	rstPedido.AbsolutePage=lote
		  	'-----------------------------------------
			NavPaginas lote,lotes,campo,criterio,texto,1%>

			<table width='100%' style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
				<thead>
                <tr bgcolor="<%=color_fondo %>"><%
				'Fila de encabezado
				'DrawFila color_fondo
					escribir_cabecera

				'CloseFila%>
				</tr>
				</thead>
				<tbody><%
				SumTotAlb=0
				'DivisaAnt=rstPedido("divisa")
				nproveedorAnt="@#@#@#@zzzz"
				fila=1
				MAXPAGINAAUX=0
				estoy_escribiendo_contacto=0
				while not rstPedido.EOF and fila<=MAXPAGINA

					CheckCAdena rstPedido("nproveedor")
				'if rstPedido("divisa")=DivisaAnt then
					if rstPedido("nproveedor")<>nproveedorAnt or null_s(rstPedido("nproveedor"))="" or IsNull(null_s(rstPedido("nproveedor"))) then
						if estoy_escribiendo_contacto=1 then
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
							'CloseFila
							'DrawFila color_fondo
                            %>
                            </tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>"></tr>
                            <tr bgcolor="<%=color_fondo %>">
                            <%
								escribir_cabecera
                            %>
                            </tr>
                            <%
							'CloseFila
							'DrawFila color_fondo
							estoy_escribiendo_contacto=0
						end if
						n_decimales=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & null_s(rstPedido("divisa")) & "'",session("backendlistados"))
						'DrawFila ""%>
                        <tr>
							<td class=tdbordeCELDA7>
								<%=Hiperv(OBJProveedores,rstPedido("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(null_s(rstPedido("nproveedor")))),LitVerProveedor)%>
							</td>
							<td class=tdbordeCELDA7>
								<%=EncodeForHtml(null_s(rstPedido("razon_social")))%>
							</td>
							<%if opccif = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("cif")))%></td>
							<%end if%>
							<%if opccontacto = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("contacto")))%></td>
							<%end if%>
							<%if opcdomicilio = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("domicilio")))%></td>
							<%end if%>
							<%if opccodigopostal = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("cp")))%></td>
							<%end if%>
							<%if opcpoblacion = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("poblacion")))%></td>
							<%end if%>
							<%if opcprovincia = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("provincia")))%></td>
							<%end if%>
							<%if opctelefono = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("telefono")))%></td>
							<%end if%>
							<%if opcfalta = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("falta")))%></td>
							<%end if
							if opcfbaja = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("fbaja")))%></td>
							<%end if

							if opcnomcom= "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("nombre")))%></td>
							<%end if
							if opcfax= "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("fax")))%></td>
							<%end if
							if opctelmov= "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("telefono2")))%></td>
							<%end if
							if opcweb= "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("web")))%></td>
							<%end if
							if opcobs= "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("observaciones")))%></td>
							<%end if
							if opcemail= "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("email")))%></td>
							<%end if

							if opctarifa = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("NomTarifa")))%></td>
							<%end if%>
							<%if opccuenta = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("cuenta_contable")))%></td>
							<%end if
							if opcformapago = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("NomFPago")))%></td>
							<%end if%>
							<%if opctipopago = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("NomTPago")))%></td>
							<%end if
							if opcrfinanciero = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("recargo")))%>%</td>
							<%end if%>
							<%if opcrequivalencia = "1" then %>
								<td class=tdbordeCELDA7><%=Visualizar(EncodeForHtml(null_s(rstPedido("re"))))%></td>
							<%end if
							if opcIRPF = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("IRPF")))%>%</td>
							<%end if%>
							<%
							''if opcexentoiva = "1" then
							%>
							<!--<td class=tdbordeCELDA7><%=Visualizar(rstPedido("exento_iva"))%></td>-->
							<%
							''end if

							if opclven1 = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("vencom1")))%></td>
							<%end if%>
							<%if opclven2 = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("vencom2")))%></td>
							<%end if
							if opcentidad = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("banco")))%></td>
							<%end if%>
							<%if opcnumcuenta = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("ncuenta")))%></td>
							<%end if
							if opcactividad = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("nomactividad")))%></td>
							<%end if
							if opctproveedor = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("nomtproveedor")))%></td>
							<%end if%>

							<%if opcportes = "1" then %>
								<td class=tdbordeCELDA7><%=EncodeForHtml(null_s(rstPedido("portes")))%></td>
							<%end if
							if ver_campo1="on" then%>
								<%
								if tipo_campo_perso1=2 then
									if ucase(rstPedido("campo01"))="ON" or cstr(null_s(rstPedido("campo01")))="1" then
										valor_campo1="Sí"
									else
										valor_campo1="No"
									end if
								elseif tipo_campo_perso1=3 then
									if rstPedido("campo01") & "">"" then
										valor_campo1=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo01")) & " and ncampo='" & session("ncliente") & "01' and tabla='PROVEEDORES'",session("backendlistados"))	
									else
										valor_campo1=""
									end if
								else
									valor_campo1= null_s(rstPedido("campo01"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo1>"",EncodeForHtml(valor_campo1),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo2="on" then%>
								<%
								if tipo_campo_perso2=2 then
									if ucase(rstPedido("campo02"))="ON" or cstr(null_s(rstPedido("campo02")))="1" then
										valor_campo2="Sí"
									else
										valor_campo2="No"
									end if
								elseif tipo_campo_perso2=3 then
									if rstPedido("campo02") & "">"" then
										valor_campo2=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo02")) & " and ncampo='" & session("ncliente") & "02' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo2=""
									end if
								else
									valor_campo2= null_s(rstPedido("campo02"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo2>"",EncodeForHtml(valor_campo2),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo3="on" then%>
								<%
								if tipo_campo_perso3=2 then
									if ucase(rstPedido("campo03"))="ON" or cstr(null_s(rstPedido("campo03")))="1" then
										valor_campo3="Sí"
									else
										valor_campo3="No"
									end if
								elseif tipo_campo_perso3=3 then
									if rstPedido("campo03") & "">"" then
										valor_campo3=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo03")) & " and ncampo='" & session("ncliente") & "03' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo3=""
									end if
								else
									valor_campo3= null_s(rstPedido("campo03"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo3>"",EncodeForHtml(valor_campo3),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo4="on" then%>
								<%
								if tipo_campo_perso4=2 then
									if ucase(rstPedido("campo04"))="ON" or cstr(null_s(rstPedido("campo04")))="1" then
										valor_campo4="Sí"
									else
										valor_campo4="No"
									end if
								elseif tipo_campo_perso4=3 then
									if rstPedido("campo04") & "">"" then
										valor_campo4=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo04")) & " and ncampo='" & session("ncliente") & "04' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo4=""
									end if
								else
									valor_campo4= null_s(rstPedido("campo04"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo4>"",EncodeForHtml(valor_campo4),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo5="on" then%>
								<%
								if tipo_campo_perso5=2 then
									if ucase(rstPedido("campo05"))="ON" or cstr(null_s(rstPedido("campo05")))="1" then
										valor_campo5="Sí"
									else
										valor_campo5="No"
									end if
								elseif tipo_campo_perso5=3 then
									if rstPedido("campo05") & "">"" then
										valor_campo5=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo05")) & " and ncampo='" & session("ncliente") & "05' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo5=""
									end if
								else
									valor_campo5= null_s(rstPedido("campo05"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo5>"",EncodeForHtml(valor_campo5),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo6="on" then%>
								<%
								if tipo_campo_perso6=2 then
									if ucase(rstPedido("campo06"))="ON" or cstr(null_s(rstPedido("campo06")))="1" then
										valor_campo6="Sí"
									else
										valor_campo6="No"
									end if
								elseif tipo_campo_perso6=3 then
									if rstPedido("campo06") & "">"" then
										valor_campo6=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo06")) & " and ncampo='" & session("ncliente") & "06' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo6=""
									end if
								else
									valor_campo6= null_s(rstPedido("campo06"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo6>"",EncodeForHtml(valor_campo6),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo7="on" then%>
								<%
								if tipo_campo_perso7=2 then
									if ucase(rstPedido("campo07"))="ON" or cstr(null_s(rstPedido("campo07")))="1" then
										valor_campo7="Sí"
									else
										valor_campo7="No"
									end if
								elseif tipo_campo_perso7=3 then
									if rstPedido("campo07") & "">"" then
										valor_campo7=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo07")) & " and ncampo='" & session("ncliente") & "07' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo7=""
									end if
								else
									valor_campo7= null_s(rstPedido("campo07"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo7>"",EncodeForHtml(valor_campo7),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo8="on" then%>
								<%
								if tipo_campo_perso8=2 then
									if ucase(rstPedido("campo08"))="ON" or cstr(null_s(rstPedido("campo08")))="1" then
										valor_campo8="Sí"
									else
										valor_campo8="No"
									end if
								elseif tipo_campo_perso8=3 then
									if rstPedido("campo08") & "">"" then
										valor_campo8=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo08")) & " and ncampo='" & session("ncliente") & "08' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo8=""
									end if
								else
									valor_campo8= null_s(rstPedido("campo08"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo8>"",EncodeForHtml(valor_campo8),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo9="on" then%>
								<%
								if tipo_campo_perso9=2 then
									if ucase(rstPedido("campo09"))="ON" or cstr(null_s(rstPedido("campo09")))="1" then
										valor_campo9="Sí"
									else
										valor_campo9="No"
									end if
								elseif tipo_campo_perso9=3 then
									if rstPedido("campo09") & "">"" then
										valor_campo9=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo09")) & " and ncampo='" & session("ncliente") & "09' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo9=""
									end if
								else
									valor_campo9= null_s(rstPedido("campo09"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo9>"",EncodeForHtml(valor_campo9),"&nbsp;")%></td>
							<%end if%>
							<%if ver_campo10="on" then%>
								<%
								if tipo_campo_perso10=2 then
									if ucase(rstPedido("campo10"))="ON" or cstr(null_s(rstPedido("campo10")))="1" then
										valor_campo10="Sí"
									else
										valor_campo10="No"
									end if
								elseif tipo_campo_perso10=3 then
									if rstPedido("campo10") & "">"" then
										valor_campo10=d_lookup("valor","campospersolista","ndetlista=" & null_s(rstPedido("campo10")) & " and ncampo='" & session("ncliente") & "10' and tabla='PROVEEDORES'",session("backendlistados"))
									else
										valor_campo10=""
									end if
								else
									valor_campo10= null_s(rstPedido("campo10"))
								end if
								%>
								<td class=TDBORDECELDA7><%=iif(valor_campo10>"",EncodeForHtml(valor_campo10),"&nbsp;")%></td>
							<%end if%>
                        </tr><%
						'CloseFila
						fila=fila+1
						if mostrarcontactos>"" then
							if rstPedido("ncontacto") & "">"" then

								'DrawFila color_terra
                                %>
                                <tr bgcolor="<%=color_terra %>">
                                <%
									DrawCelda2 "tdbordeCELDA7 bgcolor='" & color_blau & "'","",false,"&nbsp;"
									DrawCelda2 "tdbordeCELDA7","",true,LitNombreContacto
									if opcdomicilio = "1" then
										DrawCelda2 "tdbordeCELDA7","",true,LitDomicilio
									end if
									if opccodigopostal = "1" then
										DrawCelda2 "tdbordeCELDA7","",true,LitCodPostal
									end if
									if opcpoblacion = "1" then
										DrawCelda2 "tdbordeCELDA7","",true,LitPoblacion
									end if
									if opcprovincia = "1" then
										DrawCelda2 "tdbordeCELDA7","",true,LitProvincia
									end if
									if opctelefono = "1" then
										DrawCelda2 "tdbordeCELDA7","",true,LitTelefono
									end if
								'CloseFila
                                %>
                                </tr>
                                <%
								'if rstPedido("nproveedor")<>nproveedorAnt then
									MAXPAGINAAUX=MAXPAGINAAUX+1
								'end if
								'DrawFila ""
                                %>
                                <tr>
                                <%
									DrawCelda2 "tdbordeCELDA7","",false,"&nbsp;" '& rstPedido("ncontacto")
									DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstPedido("nomcontact")))
									strselec2="select * from domicilios where codigo='" & null_s(rstPedido("domiciliocontac")) & "'"
									rstAux2.cursorlocation=3
									rstAux2.open strselec2,session("backendlistados")
									if not rstAux2.eof then
										if opcdomicilio = "1" then
											DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("domicilio")))
										end if
										if opccodigopostal = "1" then
											DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("cp")))
										end if
										if opcpoblacion = "1" then
											DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("poblacion")))
										end if
										if opcprovincia = "1" then
											DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("provincia")))
										end if
										if opctelefono = "1" then
											DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("telefono")))
										end if
									end if
									rstAux2.close
									'if opctelefono = "1" then
									'	DrawCelda "tdbordeCELDA7","","",0,rstPedido("movil")
									'end if
								'CloseFila
                                %>
                                </tr>
                                <%
								fila=fila+1
								estoy_escribiendo_contacto=1
							end if
						end if
					else
						'ahora imprimimos los contactos, si los tienes
						'DrawFila ""
                        %>
                        <tr>
                        <%
							DrawCelda2 "tdbordeCELDA7","",false,"&nbsp;" & EncodeForHtml(null_s(rstPedido("ncontacto")))
							DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstPedido("nomcontact")))
							strselec2="select * from domicilios where codigo='" & null_s(rstPedido("domiciliocontac")) & "'"
							rstAux2.cursorlocation=3
							rstAux2.open strselec2,session("backendlistados")
							if not rstAux2.eof then
								if opcdomicilio = "1" then
									DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("domicilio")))
								end if
								if opccodigopostal = "1" then
									DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("cp")))
								end if
								if opcpoblacion = "1" then
									DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("poblacion")))
								end if
								if opcprovincia = "1" then
									DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("provincia")))
								end if
								if opctelefono = "1" then
									DrawCelda2 "tdbordeCELDA7","",false, EncodeForHtml(null_s(rstAux2("telefono")))
								end if
							end if
							rstAux2.close
							'if opctelefono = "1" then
							'	DrawCelda "tdbordeCELDA7","","",0,rstPedido("movil")
							'end if
						'CloseFila
                        %>
                        </tr>
                        <%
						fila=fila+1
					end if
					if not rstPedido.eof then
						nproveedorAnt= null_s(rstPedido("nproveedor"))
					end if
					rstPedido.MoveNext
				wend
				%></tbody><%
			NavPaginas lote,lotes,campo,criterio,texto,2
			rstPedido.Close%>
			</table><%
			%><input type="hidden" name="nRegsImp" value="<%=EncodeForHtml(fila-1)%>"><%
		end if
	end if%>
    <iframe name="marcoExportar" style='display:none' src="listado_proveedores_exportar.asp?mode=ver" frameborder='0' width='500' height='200'></iframe>
</form>

<%end if
set rst=nothing
set rstSelect=nothing
set rstPedido=nothing
set rstAux=nothing
set rstAux2=nothing
%>
</body>
</html>
