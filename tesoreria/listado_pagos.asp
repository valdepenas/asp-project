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
function pintar_saltos_nuevo(texto)
	texto=Replace(texto,"&#10;","")
	texto=Replace(texto,"&#13;","<br>")
	pintar_saltos_nuevo=texto
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=titulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
<META HTTP-EQUIV="Content-style-Type" CONTENT="text/css">
<LINK REL="styleSHEET" href="../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../impresora.css" MEDIA="PRINT">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="pagos.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file= "../CatFamSubResponsive.inc"-->
<!--#include file= "../styles/formularios.css.inc"-->  
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<%
if request.querystring("viene")="tienda" or request.form("viene")="tienda" then
	titulo=LitTituloLis2
else
	titulo=LitTituloLis
end if
%>

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
function agruparDia(){
	if(document.listado_pagos.opcagrupardia.checked) document.listado_pagos.opcagruparcuenta.checked=true;
}

//Desencadena la búsqueda del proveedor cuya referencia se indica
function TraerProveedor(mode) {
	document.location.href="listado_pagos.asp?nproveedor=" + document.listado_pagos.nproveedor.value + "&mode=" + mode + "&nserie=" + document.listado_pagos.nserie.value;
}

function ver_documento(ndocumento,tipodocumento,nproveedor){

	if (tipodocumento=="albaranes"){
		if(nproveedor!="")
			document.location="../compras/albaranes_pro_imp.asp?nalbaran=('" + ndocumento + "')&mode=browse&empresa="+nproveedor.substr(0,5);
		else document.location="../compras/albaranes_pro_imp.asp?nalbaran=('" + ndocumento + "')&mode=browse&empresa=<%=session("ncliente")%>";
	}
	else{
	    <%'FLM:20090505:tengo encuenta los efectos para las tiendas.%>
	    if(tipodocumento=="efectos")
	    {
	        if (nproveedor!="")
			    document.location="../central.asp?pag1=netInic.asp&s=/ventas/compras/efectos_pro.aspx&pag2=&cod=" + ndocumento + "&titulo=<%=LitEfectoProveedor%>&mode=browse&empresa="+nproveedor.substr(0,5);			    
			else document.location="../central.asp?pag1=netInic.asp&s=/ventas/compras/efectos_pro.aspx&pag2=&cod=" + ndocumento + "&titulo=<%=LitEfectoProveedor%>&mode=browse&empresa=<%=session("ncliente")%>";	    	
	    }
	    else{
		    if(nproveedor!="")
			    document.location="../compras/facturas_pro_imp.asp?nfactura=('" + ndocumento + "')&mode=browse&empresa="+nproveedor.substr(0,5);
		    else document.location="../compras/facturas_pro_imp.asp?nfactura=('" + ndocumento + "')&mode=browse&empresa=<%=session("ncliente")%>";
		}
	}
	parent.parent.topFrame.document.all("regresar").style.display="";
}
</script>

<body onload="self.status='';" class="BODY_ASP">
<%
'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
'campo: Nombre del campo con el cual se realizará la búsqueda
'criterio: Tipo de búsqueda
'texto: Texto a buscar.
function CadenaBusquedaTienda(campo,criterio,texto)
	if texto > "" then
		select case criterio
			case "contiene"
				CadenaBusquedaTienda=" where facturas_pro." + campo + " like '%" + texto + "%' and"
			case "empieza"
				CadenaBusquedaTienda=" where facturas_pro." + campo + " like '" + texto + "%' and"
			case "termina"
				CadenaBusquedaTienda=" where facturas_pro." + campo + " like '%" + texto + "' and"
			case "igual"
				CadenaBusquedaTienda=" where facturas_pro." + campo + "='" + texto + "' and"
		end select
	else
		CadenaBusquedaTienda=" where "
	end if
end function

'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
	'nproveedor:
function CadenaBusqueda(nproveedor,nserie)
	CadenaBusqueda = ""

	if nproveedor > "" then
		CadenaBusqueda = " where nproveedor='" & nproveedor & "' and"
	end if
	if nserie > ""then
			CadenaBusqueda = CadenaBusqueda + " serie='" & nserie & "' and"
	end if
	CadenaBusqueda = CadenaBusqueda + " pagada = 0 order by nfactura_pro, fecha desc,nproveedor"
end function

'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
const borde=0
'FLM:200900405:variable que controla si se tiene que realizar la ejecución de la búsqueda.
noEjecutesSQL=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
%>
<form name="listado_pagos" method="post">
<% PintarCabecera "listado_pagos.asp"
'Leer parámetros de la página
    mode=EncodeForHtml(Request.QueryString("mode"))
	if mode="browse" then mode="imp"
	nproveedor	= limpiaCadena(Request.QueryString("nproveedor"))&""
	if nproveedor ="" then
		nproveedor	= limpiaCadena(Request.form("nproveedor"))&""
	end if

	if nproveedor="" then
		nproveedor= limpiaCadena(Request.QueryString("ndoc"))&""
		if nproveedor ="" then
			nproveedor	= limpiaCadena(Request.form("ndoc"))&""
		end if
	end if
	if nproveedor & "">"" and request("viene")&""<>"tienda" then
		nproveedor=session("ncliente") & completar(nproveedor,5,"0")
	end if

	actividad	= limpiaCadena(Request.QueryString("actividad"))
	if actividad ="" then
		actividad	= limpiaCadena(Request.form("actividad"))
	end if

    'FLM:20090430:capturo las series de efectos si es una factura.
    if request.form("serie_efec")>"" then
        nserieEfec=replace(limpiaCadena(request.form("serie_efec"))," ","")
    elseif nserieEfec="" and request.form("nserieEfec")>"" then
        nserieEfec=replace(limpiaCadena(request.form("nserieEfec"))," ","")
    else
        nserieEfec=replace(limpiaCadena(request.querystring("serie_efec"))," ","")
    end if		
         
    nserie	= limpiaCadena(Request.QueryString("nserie"))
    if nserie ="" then
	    nserie	= limpiaCadena(Request.form("nserie"))
    end if

	ncuentacargo	= limpiaCadena(Request.QueryString("ncuentacargo"))
	if ncuentacargo ="" then
		ncuentacargo	= limpiaCadena(Request.form("ncuentacargo"))
	end if

	if request.form("opcagruparcuenta")>"" then
		opcagruparcuenta=request.form("opcagruparcuenta")
	else
		opcagruparcuenta=limpiaCadena(request.querystring("opcagruparcuenta"))
	end if

	if request.form("opcproveedorbaja")>"" then
		opcproveedorbaja=request.form("opcproveedorbaja")
	else
		opcproveedorbaja=limpiaCadena(request.querystring("opcproveedorbaja"))
	end if

	if request.form("opcagruparcuenta")>"" then
		opcagruparcuenta = "1"
	end if

	if request.form("opcagrupardia")>"" then
		opcagrupardia = "1"
	end if

	apaisado=iif(request.form("apaisado")>"","SI","")
	%><input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>"><%                                          

	viene	= limpiaCadena(Request.QueryString("viene"))
	if viene="" then
		viene	= Request.form("viene")
	end if

	' IML 28/04/2004 : Validamos si el usuario tiene acceso
	if viene="tienda" then
		sesionNCliente=left(nproveedor,5)
		if sesionNCliente&""="" then sesionNCliente=session("ncliente")
		checkAccesoTienda sesionNCliente,"",nproveedor
	else
		sesionNCliente=session("ncliente")
	end if
	' FIN IML 28/04/2004 : Validamos si el usuario tiene acceso

	''ricardo 21-3-2005
	DFecha=limpiaCadena(Request.Form("Dfecha"))
	if DFecha & ""="" then DFecha=limpiaCadena(request.querystring("Dfecha"))
	HFecha=limpiaCadena(Request.Form("Hfecha"))
	if Hfecha & ""="" then Hfecha=limpiaCadena(request.querystring("Hfecha"))

	campo=limpiaCadena(limpiaCadena(Request.querystring("campo")))
	criterio=limpiaCadena(Request.querystring("criterio"))
	texto=limpiaCadena(Request.querystring("texto"))

	strwhere=""

	if mode="imp" then%>
		<table width='100%' cellspacing="1" cellpadding="1">
   		<tr>
			<td width="30%" align="left">
		  		<font class=CELDAL7>&nbsp;(<%=LitEmitido%>&nbsp; <%=day(date)%>/<%=month(date)%>/<%=year(date)%>)</font>
			</td>
			<td>
				<font class='CABECERA'><b></b></font>
		 	      <font class=CELDA7><b></b></font>
			</td>
	   	</tr>
	    </table>
		<hr/>
	<%end if
	WaitBoxOculto LitEsperePorFavor
	Alarma "listado_pagos.asp"

	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstProveedor = Server.CreateObject("ADODB.Recordset")
	set rstVencimientos = Server.CreateObject("ADODB.Recordset")
	set rstAgrupar = Server.CreateObject("ADODB.Recordset")
	set rstPendiente = Server.CreateObject("ADODB.Recordset")

	'moneda_base = d_lookup("codigo","divisas","moneda_base=1 and codigo like '" & sesionNCliente & "%'",session("dsn_cliente"))
    strselect="select codigo from divisas with(nolock) where moneda_base =1 and codigo like ?+'%'"
    moneda_base= DLookupP1(strselect, sesionNCliente &"", adVarChar,15, session("dsn_cliente"))

	'n_decimalesMB=d_lookup("ndecimales","divisas","moneda_base<>0 and codigo like '" & sesionNCliente & "%'",session("dsn_cliente"))
    strselect="select ndecimales from divisas with(nolock) where moneda_base<>0 and codigo like ?+'%'"
    n_decimalesMB=DLookupP1(strselect, sesionNCliente &"", adVarChar,15, session("dsn_cliente"))

	'abreviaturaMB=d_lookup("abreviatura","divisas","moneda_base<>0 and codigo like '" & sesionNCliente & "%'",session("dsn_cliente"))
    strselect = "select abreviatura from divisas with(nolock)  where moneda_base<>0 and codigo like ?+'%'"
    abreviaturaMB= DLookupP1(strselect, sesionNCliente &"", adVarChar,15, session("dsn_cliente"))

	'mostrar_equivalencia=d_lookup("imp_equiv","configuracion","nempresa='" & sesionNCliente & "'",session("dsn_cliente"))
    strselect= "select imp_equiv from configuracion with(nolock) where nempresa=?"
    mostrar_equivalencia= DLookupP1(strselect, sesionNCliente &"", adVarChar, 5, session("dsn_cliente"))

	moneda_Ptas = sesionNCliente & "01"

	'n_decimalesPtas=d_lookup("ndecimales","divisas","codigo='" & moneda_Ptas & "'",session("dsn_cliente"))
    strselect= "select ndecimales from divisas with(nolock) where codigo=?"
    n_decimalesPtas= DLookupP1(strselect, moneda_Ptas, adVarChar, 15, session("dsn_cliente"))

	'abreviaturaPtas=d_lookup("abreviatura","divisas","codigo='" & moneda_Ptas & "'",session("dsn_cliente"))
    strselect= "select abreviatura from divisas with(nolock) where codigo=?"
    abreviaturaPtas= DLookupP1(strselect, moneda_Ptas, adVarChar, 15, session("dsn_cliente"))

	if mode="select1" then
                DrawDiv "1","",""
                DrawLabel "","",LitDesdeFecha
                DrawInput "", "", "Dfecha", iif(TmpDfecha>"",EncodeForHtml(TmpDfecha),"01/01/" & year(date)), ""
                DrawCalendar "Dfecha"
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitHastaFecha                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             
                DrawInput "", "", "Hfecha", iif(TmpHfecha>"",EncodeForHtml(TmpHfecha),day(date) & "/" & month(date) & "/" & year(date)), ""
                DrawCalendar "Hfecha"
                CloseDiv			                                                                                                                                                                                                                                  
			
                DrawDiv "1","",""                
				if nproveedor >"" then                                                                                                                                                                                                 
					'nom_proveedor=d_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente"))
                    strselect="select razon_social from proveedores with(nolock) where nproveedor=?"
                    nom_proveedor= DLookupP1(strselect, nproveedor &"", adChar, 10, session("dsn_cliente"))
				else
					nom_proveedor=""                                                                                          
				end if
				'DrawCelda2 "CELDA style='width:190px'", "left", false, LitProveedor+": "
                DrawLabel "","",LitProveedor
                %><input class='width15' type="text" name="nproveedor" value="<%=EncodeForHtml(trimCodEmpresa(nproveedor))%>" size="8" maxlength=5 onchange="TraerProveedor('<%=enc.EncodeForJavascript(mode)%>','<%=enc.EncodeForJavascript(ndet)%>');">
                <a class='CELDAREFB' href="javascript:AbrirVentana('../compras/proveedores_busqueda.asp?ndoc=listado_pagos&titulo=<%=LitSelProveedor%>&mode=search&viene=listado_pagos','P',<%=altoventana%>,<%=anchoventana%>)">
                    <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/>
                </a>
                <input class="width40" disabled="disabled" type="text" name="razon_social" size="40" class="CELDA" value="<%=EncodeForHtml(nom_proveedor)%>"><%
                CloseDiv
			
				'DrawCelda2 "CELDA style='width:190px'", "left", false, LitActividad + ": "
				rstSelect.open "select codigo,descripcion from tipo_actividad with(nolock) where codigo like '" & sesionNCliente & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				DrawSelectCelda "","","",0,LitActividad,"actividad",rstSelect,"","codigo","descripcion","",""
				rstSelect.close
                DrawDiv "1", "", ""
                DrawLabel "","",LitProveedorBaja%><input type="checkbox" name="opcproveedorbaja" <%=iif(opcproveedorbaja="true","checked","")%>><%			
                CloseDiv
			'FLM:20090504:Modifico para incluir los efectos.			
                'DrawCelda2 "CELDA", "left", false,LitEfecPend&":"
                'FLM:20090428:Modifico para incluir los efectos.
				rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & sesionNCliente & "%' and tipo_documento ='EFECTO PROVEEDOR'",session("backendlistados"),adOpenKeyset,adLockOptimistic
				DrawSelectMultipleCelda "","","",0,LitEfecPend,"serie_efec",rstAux,iif(nserie>"" and tabla="facturas_pro",nserie,""),"nserie","descripcion","",""
				rstAux.close
			    rstSelect.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='FACTURA DE PROVEEDOR' and nserie like '" & sesionNCliente & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				DrawSelectCelda "","","",0,LitSerie,"nserie",rstSelect,nserie,"nserie","descripcion","",""
				rstSelect.close			
		
			'DrawCelda2 "CELDA style='width:190px'", "left", false, LitNCuentaCargo + ":"
			rstSelect.cursorlocation=3
			rstSelect.open "select ncuenta from bancos with(nolock) where nbanco like '" & sesionNCliente & "%'",session("dsn_cliente")
			DrawSelectCelda "","","",0,LitNCuentaCargo,"ncuentacargo",rstSelect,"","ncuenta","ncuenta","",""
			rstSelect.close
			'DrawCelda2 "CELDA width='40px'", "left", false, "&nbsp;"
			'DrawCelda2 "CELDA width='120px'", "left", false, LitAgruparCuenta + ":"
            DrawDiv "1", "", ""
            DrawLabel "","",LitAgruparCuenta%><input type="checkbox" name="opcagruparcuenta" onclick="javascript:agruparDia();"><%
            CloseDiv
		
            DrawDiv "1", "", ""
            DrawLabel "","",LitAgruparDia%><input type="checkbox" name="opcagrupardia" onclick="javascript:agruparDia();"><%
            CloseDiv
		    
elseif mode="imp" then                                                         
''ricardo 25-5-2006 comienzo de la select
''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"inicio_listado_pagos"%>
		<input type="hidden" name="nproveedor" value="<%=EncodeForHtml(iif(viene="tienda",nproveedor,trimCodEmpresa(nproveedor)))%>">
		<input type="hidden" name="actividad" value="<%=EncodeForHtml(actividad)%>">
		<input type="hidden" name="nserie" value="<%=EncodeForHtml(nserie)%>">
		<input type="hidden" name="opcproveedorbaja" value="<%=EncodeForHtml(opcproveedorbaja)%>">
		<input type="hidden" name="ncuentacargo" value="<%=EncodeForHtml(ncuentacargo)%>">                             
		<input type="hidden" name="opcagruparcuenta" value="<%=EncodeForHtml(opcagruparcuenta)%>">
		<input type="hidden" name="opcagrupardia" value="<%=EncodeForHtml(opcagrupardia)%>">
		<input type="hidden" name="mode" value="<%=EncodeForHtml(mode)%>">
		<% if viene="tienda" then%>
			<input type="hidden" name="viene" value="<%=EncodeForHtml(viene)%>">
			<input type="hidden" name="campo" value="<%=EncodeForHtml(campo)%>">
			<input type="hidden" name="criterio" value="<%=EncodeForHtml(criterio)%>">
			<input type="hidden" name="texto" value="<%=EncodeForHtml(texto)%>">
		<%end if%>
		<input type="hidden" name="Dfecha" value="<%=EncodeForHtml(Dfecha)%>">
		<input type="hidden" name="Hfecha" value="<%=EncodeForHtml(Hfecha)%>">
		<input type="hidden" name="nserieEfec" value="<%=EncodeForHtml(nserieEfec)%>" />

		<%
		'MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='123'", DSNIlion)   
        strselect="select maxpagina from limites_listados with(nolock) where item=?"
        MAXPAGINA= DLookupP1(strselect, "123", adVarChar,3, DSNIlion)

	    'MAXPDF=d_lookup("maxpdf", "limites_listados", "item='123'", DSNIlion)
        strselect="select maxpdf from limites_listados with(nolock) where item=?"
        MAXPDF= DLookupP1(strselect, "123", adVarChar,3, DSNIlion)
		%><input type="hidden" name="maxpdf" value="<%=EncodeForHtml(MAXPDF)%>">
		<input type="hidden" name="maxpagina" value="<%=EncodeForHtml(MAXPAGINA)%>"><%

		strwhere=CadenaBusquedaTienda(campo,criterio,texto)
		strwhere2=strwhere
		strwhereVenc=strwhere
		strwhereEfec=" where "
		strwhere=strwhere & " proveedores.nproveedor = facturas_pro.nproveedor and divisas.codigo=facturas_pro.divisa and"
		strwhereVenc=strwhereVenc & " proveedores.nproveedor = facturas_pro.nproveedor and divisas.codigo=facturas_pro.divisa and"
		'FLM:20090916:añado like de tablas proveedor y divisas.
		strwhere=strwhere & " facturas_pro.nfactura like '" & sesionNCliente & "%' and proveedores.nproveedor like '" & sesionNCliente & "%' and divisas.codigo like '" & sesionNCliente & "%' and"
		strwhereVenc=strwhereVenc & " facturas_pro.nfactura like '" & sesionNCliente & "%' and proveedores.nproveedor like '" & sesionNCliente & "%' and divisas.codigo like '" & sesionNCliente & "%' and"
		strwhere2=strwhere2 & " divisas.codigo=facturas_pro.divisa and"
		strwhere2=strwhere2 & " facturas_pro.nfactura like '" & sesionNCliente & "%' and proveedores.nproveedor like '" & sesionNCliente & "%' and divisas.codigo like '" & sesionNCliente & "%' and"
		'FLM:20090504:para los efectos.
		strwhereEfec=strwhereEfec & " proveedores.nproveedor = efectos_pro.nproveedor and divisas.codigo=efectos_pro.divisa and efectos_pro.nefecto like '" & sesionNCliente & "%' and proveedores.nproveedor like '" & sesionNCliente & "%' and divisas.codigo like '" & sesionNCliente & "%' and"

		hay_cabecera=0
		if nserieEfec&"">"" then
				hay_cabecera=1
		end if
		if nproveedor > "" then
			strwhere = strwhere & " facturas_pro.nproveedor='" & nproveedor & "' and"
			strwhereVenc = strwhereVenc & " facturas_pro.nproveedor='" & nproveedor & "' and"
			strwhere2=strwhere2 & " facturas_pro.nproveedor='" & nproveedor & "' and"
			'FLM:20090504:para los efectos.
			strwhereEfec=strwhereEfec & " efectos_pro.nproveedor='" & nproveedor & "' and"
            'd_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente"))
            strselect = "select razon_social from proveedores with(nolock) where nproveedor=?"
			if viene<>"tienda" then
				%><font class='CELDA'><b><%=LitProveedor%>: </b><%=EncodeForHtml(trimCodEmpresa(nproveedor)) & " - " & EncodeForHtml(DLookupP1(strselect, nproveedor &"", adChar,10, session("dsn_cliente"))) %></font><br/><%
				hay_cabecera=1
			end if                                                                           
		else
			if opcproveedorbaja="" then
				strbaja=" "
				''strwhere=strwhere & strbaja
			else
			      %><font class='CELDA'><b><%=LitProveedorBaja%></b></font><br/><%
				strbaja=" proveedores.fbaja is null and"
				strwhere=strwhere & strbaja
				strwhereVenc=strwhereVenc & strbaja
				'FLM:20090504:para los efectos.
		        strwhereEfec=strwhereEfec & " proveedores.fbaja is null and "

				hay_cabecera=1
			end if
		end if
		if Dfecha & "">"" then
			strwhere= strWhere & " facturas_pro.fecha>='" & Dfecha & " ' and"
			strwhereVenc= strwhereVenc & " vencimientos_entrada.fecha>='" & Dfecha & " ' and"
			strwhere2= strWhere2 & " albaranes_pro.fecha>='" & Dfecha & " 00:00:00' and"
			'FLM:20090504:para los efectos.
			strwhereEfec=strwhereEfec & " isnull(efectos_pro.fechavto,efectos_pro.fecha)>='" & Dfecha & "' and"
			%><font class='CELDA'><b><%=LitDesdeFecha%>: </b><%=EncodeForHtml(Dfecha)%></font><br/><%                                           
			hay_cabecera=1
		end if
		if Hfecha & "">"" then
			strwhere= strWhere & " facturas_pro.fecha<='" & Hfecha & "' and"
			strwhereVenc= strwhereVenc & " vencimientos_entrada.fecha<='" & Hfecha & "' and"
			''strwhereEfec= strwhereEfec & " facturas_pro.fecha<='" & Hfecha & "' and"
			strwhere2= strWhere2 & " albaranes_pro.fecha<='" & Hfecha & " 23:59:00' and"
			'FLM:20090504:para los efectos.
			strwhereEfec=strwhereEfec & " isnull(efectos_pro.fechavto,efectos_pro.fecha)<='" & Hfecha & "' and"
			%><font class='CELDA'><b><%=LitHastaFecha%>: </b><%=EncodeForHtml(Hfecha)%></font><br/><%                                 
			hay_cabecera=1
		end if
		if actividad>"" then
			strwhere = strwhere & " proveedores.tactividad='" & actividad & "' and "
			strwhereVenc = strwhereVenc & " proveedores.tactividad='" & actividad & "' and "
			'FLM:20090504:para los efectos.
			strwhereEfec=strwhereEfec & " proveedores.tactividad='" & actividad & "' and "
            'd_lookup("descripcion","tipo_actividad","codigo='" & actividad & "'",session("dsn_cliente"))
            strselect="select descripcion from tipo_actividad with(nolock) where codigo=?"
			%><font class='CELDA'><b><%=LitActividad%>: </b><%=EncodeForHtml(DLookupP1(strselect, actividad &"" , adVarChar,10,session("dsn_cliente")))%></font><br/><%
			hay_cabecera=1
		end if
		if nserie>"" then
			strwhere= strWhere & " serie in  ('" & replace(replace(nserie," ",""),",","','") & "') and"
			strwhereVenc= strwhereVenc & " serie in  ('" & replace(replace(nserie," ",""),",","','") & "') and"
			strwhere2= strWhere2 & " serie in  ('" & replace(replace(nserie," ",""),",","','") & "') and"
			strwhere7= " fce.serie in (''" & replace(replace(nserie," ",""),",","'',''") & "'') and "			
			%><font class='CELDA'><b><%=LitSerie%>: </b>
			<%'FLM:20090504:no sirve porque puede ser múltiple. %>
			<%'=d_lookup("nombre","series","nserie in ('" & replace(replace(nserie," ",""),",","','") & "')",session("dsn_cliente"))%>
			<%rstAux.open "select nombre from series with(nolock) where nserie like '" & sesionNCliente & "%' and nserie in ('" & replace(replace(nserie," ",""),",","','") & "') ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			salida=""
			while not rstAux.EOF				    
		        salida=salida& rstAux("nombre")&", "
		        rstAux.MoveNext
		    wend
			rstAux.close
			response.Write EncodeForHtml(left(salida,len(salida)-2)) %>
			</font><br/><%
			hay_cabecera=1
		end if
        'FLM:20090820:series de efectos
		if nserieEfec&"">"" then
		    strwhereEfec = strwhereEfec & " efectos_pro.serie in ('" & replace(replace(nserieEfec," ",""),",","','") & "') and"
		    rstAux.open "select substring(nserie,6,len(nserie))+' '+nombre as nombre from series with(nolock) where nserie like '" & sesionNCliente & "%' and nserie in ('" & replace(replace(nserieEfec," ",""),",","','") & "')",session("backendlistados")
		    listaSerie=""
		    while not rstAux.eof  
	            listaSerie= listaSerie&rstAux("nombre")&","
	            rstAux.moveNext		          
	        wend 
	        rstAux.Close() 
		    if listaSerie>"" then %>
		         <font class="cab"><b><%=LitSerieEfec%>:&nbsp;</b></font>
		        <font class="cab"><%=EncodeForHtml(left(listaSerie,len(listaSerie)-1))%></font><br/>
		    <%end if%>
		<%end if
		if ncuentacargo>"" then
			strwhere= strWhere & " facturas_pro.ncuenta='" & ncuentacargo & "' and"
			strwhereVenc= strwhereVenc & " facturas_pro.ncuenta='" & ncuentacargo & "' and"
			strwhere2= strWhere2 & " facturas_pro.ncuenta='" & ncuentacargo & "' and"
			strwhereEfec= strwhereEfec & " banc.ncuenta='" & ncuentacargo & "' and"
''			'FLM:20090504:para los efectos==> si son efectos y hay una cuenta de cargo marcada para la búsqueda,el resultado debe ser vacío, ya que los efectos no tienen cuenta de cargo.
''			if nserieEfec&"">"" then
''                noEjecutesSQL=1
''            end if
            'd_lookup("ncuenta","bancos","ncuenta='" & ncuentacargo & "'",session("dsn_cliente"))
            strselect="select ncuenta from bancos where ncuenta =?"
			%><font class='CELDA'><b><%=LitNCuenta%>: </b><%=EncodeForHtml(DLookupP1(strselect, ncuentacargo  &"", adVarChar, 25, session("dsn_cliente")))%></font><br/><%
			hay_cabecera=1
		end if

		if hay_cabecera=1 then
			%><hr/><%
		end if

		strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
		strwhereVenc=mid(strwhereVenc,1,len(strwhereVenc)-4) 'Quitamos el último AND
		strwhere2=mid(strwhere2,1,len(strwhere2)-4) 'Quitamos el último AND
		'FLM:20090504:para los efectos.
        strwhereEfec=mid(strwhereEfec,1,len(strwhereEfec)-4) 'Quitamos el último AND

		'' SIN AGRUPACIONES

		if opcagruparcuenta<>"1" and opcagrupardia<>"1" then
		    'FLM:20090504: si son efectos diferenciamos.
		    if nserieEfec&"">"" then
		       ' 'Inserto las facturas y los vencimientos.
		       ' seleccion=" set nocount on; "
		        'seleccion = seleccion & " select AA.* into #temp_fac_pago from "
		        seleccion=seleccion &"/*( */SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'F' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, deuda AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores with(nolock) " & strwhere & " and pagada=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and facturas_pro.deuda<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'V' AS Tipo, vencimientos_entrada.nfactura AS Ndoc, nvencimiento AS Nvto,vencimientos_entrada.fecha AS Fecha, importe AS Total,importe AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores with(nolock), vencimientos_entrada with(nolock) " & strwhereVenc & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and vencimientos_entrada.importe<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura, divisas.ndecimales, facturas_pro.nfactura_pro AS nfactura_pro,facturas_pro.divisa AS divisa,facturas_pro.nproveedor AS nproveedor, razon_social,'VI' AS Tipo, nfactura AS Ndoc, NULL AS Nvto, fecha AS Fecha,facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda,facturas_pro.ncuenta AS ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock), FACTURAS_PRO with(nolock), proveedores with(nolock) " & strwhere & " and pagada=0 and facturas_pro.total_irpf<>0"
			    seleccion=seleccion & " UNION "
			    seleccion = seleccion &" SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.ndocefecto as nfactura_pro,efectos_pro.divisa as divisa,efectos_pro.nproveedor as nproveedor,razon_social, 'E' AS Tipo, efectos_pro.nefecto AS Ndoc, NULL AS Nvto,efectos_pro.fecha AS Fecha, efectos_pro.importe AS Total, efectos_pro.importe AS Deuda , banc.ncuenta as ncuentacargo "
			    seleccion = seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
			    seleccion = seleccion & " left outer join bancos as banc with(NOLOCK) on banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "

			    seleccion = seleccion & ", proveedores with(nolock) " & strwhereEfec & " and pendiente=1 "
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and efectos_pro.importe<>0"
			    seleccion = seleccion & " ORDER BY razon_social, Ndoc, tipo, Nvto"
			    'seleccion =seleccion & ") as AA "
			    ''Borro las facturas y vencimientos que están en los efectos.
			    'seleccion = seleccion & " DELETE from #temp_fac_pago "
		        'seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock)"
		        'seleccion = seleccion & " inner join detalles_efpro de with(nolock) on de.nefecto like '"&sesionNCliente&"%' and de.nefecto=efectos_pro.nefecto "            
		        'seleccion = seleccion & " inner join facturas_pro fce with(nolock) on fce.nfactura like '"&sesionNCliente&"%' and " &strwhere7&" (fce.nfactura=isnull(de.nfactura,'')+isnull(de.nfacturavto,'') )  "            
                'seleccion = seleccion & " , proveedores with(nolock) " & strwhereEfec & " and pendiente=1 "
                'WHERE DEL DELETE
                'seleccion=seleccion& " and (( #temp_fac_pago.ndoc=de.nfactura and #temp_fac_pago.Tipo='F') or (#temp_fac_pago.ndoc=de.nfacturavto+'-'+convert(varchar,de.nvto) and #temp_fac_pago.Tipo='V')  ) "
                ''SELECT DE LA TABLA TEMPORAL Y DE LOS EFECTOS
                'seleccion = seleccion & " select * from #temp_fac_pago "
			    'seleccion = seleccion & " UNION "
                'seleccion = seleccion &" SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.ndocefecto as nfactura_pro,efectos_pro.divisa as divisa,efectos_pro.nproveedor as nproveedor,razon_social, 'E' AS Tipo, efectos_pro.nefecto AS Ndoc, NULL AS Nvto,efectos_pro.fecha AS Fecha, efectos_pro.importe AS Total, efectos_pro.importe AS Deuda , null as ncuentacargo "
			    'seleccion = seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
			    ''seleccion = seleccion & " left join detalles_efpro de with(nolock) on de.nefecto like '"&sesionNCliente&"%' and de.nefecto=efectos_pro.nefecto "
			    ''seleccion = seleccion & " left join facturas_pro fce with(nolock) on fce.nfactura like '"&sesionNCliente&"%' and " &strwhere7&" (fce.nfactura=isnull(de.nfactura,'')+isnull(de.nfacturavto,'') )  "            
			    'seleccion = seleccion & ", proveedores with(nolock) " & strwhereEfec & " and pendiente=1 "
			    'seleccion = seleccion & " ORDER BY razon_social, Ndoc, tipo, Nvto"
			    'seleccion = seleccion & " set nocount off;"
		    else
			    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'F' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, deuda AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores with(nolock) " & strwhere & " and pagada=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and facturas_pro.deuda<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'V' AS Tipo, vencimientos_entrada.nfactura AS Ndoc, nvencimiento AS Nvto,vencimientos_entrada.fecha AS Fecha, importe AS Total,importe AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores with(nolock), vencimientos_entrada with(nolock) " & strwhereVenc & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and vencimientos_entrada.importe<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura, divisas.ndecimales, facturas_pro.nfactura_pro AS nfactura_pro,facturas_pro.divisa AS divisa,facturas_pro.nproveedor AS nproveedor, razon_social,'VI' AS Tipo, nfactura AS Ndoc, NULL AS Nvto, fecha AS Fecha,facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda,facturas_pro.ncuenta AS ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock), FACTURAS_PRO with(nolock), proveedores with(nolock) " & strwhere & " and pagada=0 and facturas_pro.total_irpf<>0"
			    seleccion=seleccion & " ORDER BY razon_social, Ndoc, tipo, Nvto"
			end if
			'FLM:20090405:la ejecucion la saco fuera del if para que sólo exista una.
			'rst.cursorlocation=3
			'rst.Open seleccion,session("dsn_cliente")
			'num_registros=rst.RecordCount

		'' AGRUPANDO POR CUENTA
	 	elseif opcagruparcuenta="1" and opcagrupardia<>"1" then
	 	     'FLM:20090504: si son efectos diferenciamos.
		    if nserieEfec&"">"" then
		        ''Inserto las facturas y los vencimientos.
		        'seleccion=" set nocount on; "
		        'seleccion = seleccion & " select AA.* into #temp_fac_pago from "
		        seleccion=seleccion &"/*(*/ SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'F' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, deuda AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores  with(nolock)" & strwhere & " and pagada=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and facturas_pro.deuda<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'V' AS Tipo, vencimientos_entrada.nfactura AS Ndoc, nvencimiento AS Nvto,vencimientos_entrada.fecha AS Fecha, importe AS Total,importe AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores with(nolock), vencimientos_entrada  with(nolock)" & strwhereVenc & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and vencimientos_entrada.importe<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura, divisas.ndecimales,facturas_pro.nfactura_pro AS nfactura_pro,facturas_pro.divisa AS divisa,facturas_pro.nproveedor AS nproveedor, razon_social,'VI' AS Tipo, nfactura AS Ndoc, NULL AS Nvto, fecha AS Fecha,facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda,facturas_pro.ncuenta AS ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock), FACTURAS_PRO with(nolock), proveedores with(nolock) " & strwhere & " and pagada=0 and facturas_pro.total_irpf<>0 "
			    seleccion=seleccion & " UNION "
			    seleccion = seleccion & "SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.ndocefecto as nfactura_pro,efectos_pro.divisa as divisa,efectos_pro.nproveedor as nproveedor,razon_social, 'E' AS Tipo, efectos_pro.nefecto AS Ndoc, NULL AS Nvto,efectos_pro.fecha AS Fecha, efectos_pro.importe AS Total, efectos_pro.importe AS Deuda , banc.ncuenta as ncuentacargo "
			    seleccion = seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
			    seleccion = seleccion & " left outer join bancos as banc with(NOLOCK) on banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
			    'seleccion = seleccion & " left join detalles_efpro de with(nolock) on de.nefecto like '"&sesionNCliente&"%' and de.nefecto=efectos_pro.nefecto "
			    'seleccion = seleccion & " left join facturas_pro fce with(nolock) on fce.nfactura like '"&sesionNCliente&"%' and " &strwhere7&" (fce.nfactura=isnull(de.nfactura,'')+isnull(de.nfacturavto,'') )  "            
			    seleccion = seleccion & ", proveedores with(nolock) " & strwhereEfec & " and pendiente=1 "
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and efectos_pro.importe<>0"
			    seleccion=seleccion & " ORDER BY ncuentacargo,razon_social, Ndoc, tipo, Nvto"
			    
		        'seleccion =seleccion & ") as AA "
			    ''Borro las facturas y vencimientos que están en los efectos.
			    'seleccion = seleccion & " DELETE from #temp_fac_pago "    
		        'seleccion = seleccion & " DELETE from #temp_fac_pago "
		        'seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
		        'seleccion = seleccion & " inner join detalles_efpro de with(nolock) on de.nefecto like '"&sesionNCliente&"%' and de.nefecto=efectos_pro.nefecto "            
		        'seleccion = seleccion & " inner join facturas_pro fce with(nolock) on fce.nfactura like '"&sesionNCliente&"%' and " &strwhere7&" (fce.nfactura=isnull(de.nfactura,'')+isnull(de.nfacturavto,'') )  "            
                'seleccion = seleccion & " , proveedores with(nolock) " & strwhereEfec & " and pendiente=1 "
		        'WHERE DEL DELETE
                'seleccion=seleccion& " and (( #temp_fac_pago.ndoc=de.nfactura and #temp_fac_pago.Tipo='F') or (#temp_fac_pago.ndoc=de.nfacturavto+'-'+convert(varchar,de.nvto) and #temp_fac_pago.Tipo='V')  ) "
                'SELECT DE LA TABLA TEMPORAL Y DE LOS EFECTOS
                'seleccion = seleccion & " select * from #temp_fac_pago "
			    'seleccion = seleccion & " UNION "
                'seleccion = seleccion & "SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.ndocefecto as nfactura_pro,efectos_pro.divisa as divisa,efectos_pro.nproveedor as nproveedor,razon_social, 'E' AS Tipo, efectos_pro.nefecto AS Ndoc, NULL AS Nvto,efectos_pro.fecha AS Fecha, efectos_pro.importe AS Total, efectos_pro.importe AS Deuda , null as ncuentacargo "
			    'seleccion = seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
			    'seleccion = seleccion & " left join detalles_efpro de with(nolock) on de.nefecto like '"&sesionNCliente&"%' and de.nefecto=efectos_pro.nefecto "
			    'seleccion = seleccion & " left join facturas_pro fce with(nolock) on fce.nfactura like '"&sesionNCliente&"%' and " &strwhere7&" (fce.nfactura=isnull(de.nfactura,'')+isnull(de.nfacturavto,'') )  "            
			    'seleccion = seleccion & ", proveedores with(nolock) " & strwhereEfec & " and pendiente=1 "
			    'seleccion=seleccion & " ORDER BY ncuentacargo,razon_social, Ndoc, tipo, Nvto"
			    'seleccion = seleccion & " set nocount off;"
		    else
			    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'F' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, deuda AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores  with(nolock)" & strwhere & " and pagada=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and facturas_pro.deuda<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.nfactura_pro as nfactura_pro,facturas_pro.divisa as divisa,facturas_pro.nproveedor as nproveedor,razon_social, 'V' AS Tipo, vencimientos_entrada.nfactura AS Ndoc, nvencimiento AS Nvto,vencimientos_entrada.fecha AS Fecha, importe AS Total,importe AS Deuda , facturas_pro.ncuenta as ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock), proveedores with(nolock), vencimientos_entrada  with(nolock)" & strwhereVenc & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0"
			    ''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			    seleccion=seleccion & " and vencimientos_entrada.importe<>0"
			    seleccion=seleccion & " UNION "
			    seleccion=seleccion & " SELECT divisas.abreviatura, divisas.ndecimales,facturas_pro.nfactura_pro AS nfactura_pro,facturas_pro.divisa AS divisa,facturas_pro.nproveedor AS nproveedor, razon_social,'VI' AS Tipo, nfactura AS Ndoc, NULL AS Nvto, fecha AS Fecha,facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda,facturas_pro.ncuenta AS ncuentacargo"
			    seleccion=seleccion & " FROM divisas with(nolock), FACTURAS_PRO with(nolock), proveedores with(nolock) " & strwhere & " and pagada=0 and facturas_pro.total_irpf<>0"
			    seleccion=seleccion & " ORDER BY ncuentacargo,razon_social, Ndoc, tipo,Nvto"
			end if
			'FLM:20090405:la ejecucion la saco fuera del if para que sólo exista una.
			'rst.cursorlocation=3
			'rst.Open seleccion,session("dsn_cliente")
			'num_registros=rst.RecordCount
		'' AGRUPANDO POR CUENTA Y POR DIA
	 	elseif opcagruparcuenta="1" and opcagrupardia="1" then
			set conn = Server.CreateObject("ADODB.Connection")
			conn.open session("dsn_cliente")

			if opcproveedorbaja & ""="" then
				mostrar_opcproveedorbaja="0"
			else
				mostrar_opcproveedorbaja="1"
			end if
			'FLM:200904504:Modifico el parámetro de serie para que admita múltipes y añado un parametro más si es un listado de efectos para que liste los efectos.
			seleccion="exec ListPagosPendAgrupCuentaYDia "
			seleccion=seleccion & "@nfactura='',@nproveedor='" & nproveedor & "',@proveedor_baja=" & mostrar_opcproveedorbaja & ",@serieFac='" & replace(replace(nserie," ",""),",","'',''") & "',@actividad='" & actividad & "',@ncuenta='" & ncuentacargo & "',@usuario='" & session("usuario") & "',@nempresa='" & sesionNCliente & "'"
			if nserieEfec&"">"" then
			    seleccion=seleccion &",@serieEfec='''" & replace(replace(nserieEfec," ",""),",","'',''") & "'''"
			else    
			    seleccion=seleccion &",@serieEfec=''"
			end if
			'FLM:20090405:Añado parámetros para poder filtrar por fechas.
			fechaFin=Hfecha
			seleccion=seleccion &",@fechaDesde='"&Dfecha&"',@fechaHasta='"&fechaFin&"'"
			
            'FLM:20090405:la ejecucion la saco fuera del if para que sólo exista una.
			'set rst = conn.execute(seleccion)
			'if not rst.eof then
			'	num_registros=rst("contador")
			'else
			'	num_registros=0
			'end if
	 	end if
	 	'FLM:20090504: si no se tiene que ejecutar o si no hay registros,hayRegistros=0. 1 en otro caso.
	 	num_registros=0
 	    rst.cursorlocation=3
 	    'FLM:20090821:ejecuto sobre el backend
		rst.Open seleccion,session("backendlistados")
		if rst.EOF then
	        hayRegistros=0
		    rst.Close
		else
		    num_registros=rst.RecordCount
		    hayRegistros=1
	    end if

		if hayRegistros=1 then 'HAY REGISTROS
			%><input type="hidden" name="NumRegsTotal" value="<%=EncodeForHtml(num_registros)%>"><%                               
			''Calculos de páginas--------------------------
		      lote=limpiaCadena(Request.QueryString("lote"))
		      if lote="" then lote=1
		      sentido=limpiaCadena(Request.QueryString("sentido"))

		      lotes=num_registros/MAXPAGINA
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

			if opcagruparcuenta="1" and opcagrupardia="1" then
				num_reg_pagT=MAXPAGINA*(lote-1)
				num_reg_pag=1
				while not rst.eof and num_reg_pag<=num_reg_pagT
					rst.movenext
					num_reg_pag=num_reg_pag+1
				wend
			else
				rst.PageSize=MAXPAGINA
				rst.AbsolutePage=lote
			end if
		      ''-----------------------------------------

			NavPaginas lote,lotes,campo,criterio,texto,1
			if lotes>1 then
				%><hr/><%
			end if
			%><table width='100%' border='0'  style='BORDER: 1px solid Black;' style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
				''Fila de encabezado

				if opcagrupardia<>"1" then
					DrawFila color_fondo
						if opcagruparcuenta="1" then
							DrawCelda "ENCABEZADOL7 width='130px' style='BORDER: 1px solid Black;'","","",0,LitNCuenta
						end if
						DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitProveedor						
						'FLM:20090504:Para los efectos opngo otra cabecera.
						if efectos&"">"" then 
						    DrawCelda "ENCABEZADOL7 style='width:150px;BORDER: 1px solid Black;'","","",0,LitEfectos
						else
						    DrawCelda "ENCABEZADOL7 style='width:150px;BORDER: 1px solid Black;'","","",0,LitFacVen
						end if
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitFechaF
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitFechaV
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitValorF
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitValorV
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitPendienteF
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitPendienteV
					CloseFila
				else
					DrawFila color_fondo
						DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitNCuenta
						if rst("HayDeudaIRPF")<>"0" then
							DrawCelda "ENCABEZADOL7 colspan=2 style='BORDER: 1px solid Black;'","","",0,LitFecha
						else
							DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitFecha
						end if
						DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitImporte
						if rst("HayDeudaIRPF")<>"0" then
							DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitIRPF
						end if
					CloseFila
				end if

			if viene<>"tienda" then
				VinculosPagina(MostrarProveedores)=1:VinculosPagina(MostrarFacturasPro)=1:VinculosPagina(MostrarEfectosPro)=1
				CargarRestricciones session("usuario"),sesionNCliente,Permisos,Enlaces,VinculosPagina
			end if

			Gtotal_valor = 0
			Gtotal_pendiente = 0
			Gtotal_pendienteV = 0
			total_valor = 0
			total_pendiente = 0
			total_pendienteV = 0
			cli=1
			ProveedorAnt=""
			CuentaAnt=""
			fila = 1
			fila2=1
			primeralinea=0

			' ******************  CASO DE NO AGRUPACION

			if opcagruparcuenta<>"1" and opcagrupardia<>"1" then
				while not rst.EOF and fila<=MAXPAGINA
					if rst("nproveedor")<>ProveedorAnt then
						if ProveedorAnt<>"" then
						   ' if not nserieEfec&"">"" then 'CONDICION PARA EFECTOS/FACTURAS
							    '' Facturas
							    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
							    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0 and facturas_pro.deuda<>0"
							    rstAux.cursorlocation=3
							    rstAux.open seleccion,session("dsn_cliente")
							    total_valor=0
							    total_pendiente=0
							    while not rstAux.eof
								    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
								    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
								    rstAux.movenext
							    wend
							    registrosT=rstAux.recordcount
							    rstAux.close

							    '' Vencimientos
							    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
							    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & replace(strwhere,"facturas_pro.fecha","vencimientos_entrada.fecha") & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 and vencimientos_entrada.importe<>0 and facturas_pro.nproveedor='" & ProveedorAnt & "'"
							    rstAux.cursorlocation=3
							    rstAux.open seleccion,session("dsn_cliente")
							    if not rstAux.eof then
							        registrosT=registrosT+rstAux.recordcount
							    end if
							    rstAux.close

							    '' IRPF
							    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
							    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0 and facturas_pro.total_irpf<>0"
							    rstAux.cursorlocation=3
							    rstAux.open seleccion,session("dsn_cliente")
							    while not rstAux.eof
								    registrosT=registrosT+1
								    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
								    rstAux.movenext
							    wend
							    rstAux.close

							    ' FLM:20090504:efectos
							    if nserieEfec&"">"" then
							        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
	    					        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
		    				        if cuentaAnt<>LitSinAsignar then
			    			            seleccion = seleccion & ", bancos as banc with(NOLOCK) "
				    		        end if
					    	        seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and efectos_pro.nproveedor='" & ProveedorAnt & "' and pendiente=1 and efectos_pro.importe<>0"
						            if cuentaAnt<>LitSinAsignar then
						                seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
						            end if
						            seleccion = seleccion & strwhereEfec3

							        rstAux.cursorlocation=3
							        rstAux.open seleccion,session("dsn_cliente")
							        while not rstAux.eof
							            registrosT=registrosT+1
								        total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
						    		    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							        rstAux.movenext
							        wend							    
							        rstAux.close
							    end if
							'end if'FIN CONDICIÓN EFECTOS / FACTURAS

							DrawFila color_fondo
								DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalP & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
							CloseFila
							fila2=1
							if fila+1<=MAXPAGINA then
								'dejamos ahora dos espacios en blanco
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								'Fila de encabezado
								DrawFila color_fondo
									DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitProveedor
									DrawCelda "ENCABEZADOL7 style='width:150px;BORDER: 1px solid Black;'","","",0,LitFacVen
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitFechaF
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitFechaV
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitValorF
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitValorV
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitPendienteF
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitPendienteV
								CloseFila
							end if
						else
							DrawFila ""
						end if
						total_valor=0
						total_pendiente=0
						total_pendienteV=0
						if viene<>"tienda" then
							dat1=Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(rst("razon_social")),LitVerProveedor)
						else
							dat1=EncodeForHtml(rst("razon_social"))
						end if
						DrawCelda "ENCABEZADOL7","","",0,dat1
					else
						DrawCelda "CELDA","","",0,"&nbsp;"
					end if

					if rst("Tipo")="F" then		'' FACTURAS
						primeralinea=0
						if viene<>"tienda" then
							dat1=Hiperv(OBJFacturasPro,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(rst("nfactura_pro")),LitVerFactura)
						else
							dat1="<a class='CELDAREFB' href=javascript:ver_documento('" & enc.EncodeForJavascript(rst("ndoc")) & "','facturas','"&enc.EncodeForJavascript(nproveedor)&"') alt='" & LitVerFacturas & "'>" & EncodeForHtml(rst("nfactura_pro")) & "</a>"
						end if
						DrawCelda "tdbordeCELDA7","","",0,dat1
						DrawCelda "tdbordeCELDA7 ","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1) & " "  & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						total_valor = total_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						total_pendiente = total_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
						Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
					elseif rst("Tipo")="V" then		'' VENCIMIENTOS
                        'd_lookup("nfactura_pro","facturas_pro","nfactura='" & rst("ndoc") & "'",session("dsn_cliente"))
                        strselect="select nfactura_pro from facturas_pro where nfactura =?"
						DrawCelda "tdbordeCELDA7 align='right'","","",0, EncodeForHtml(DLookupP1(strselect, rst("ndoc") &"",adVarChar,20, session("dsn_cliente"))  & "-" & rst("Nvto"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 ","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "tdbordeCELDA7","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))
						total_pendienteV = total_pendienteV + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
						Gtotal_pendienteV = Gtotal_pendienteV + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
						'FLM:20090821:solo para los efectos.
						if nserieEfec>"" then
						    total_valorV = total_valorV + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						    Gtotal_valorV = Gtotal_valorV + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						end if
					elseif rst("Tipo")="VI" then	'' IRPF
						DrawCelda "tdbordeCELDA7 align='right'","","",0,LitIRPF & ": " & EncodeForHtml(rst("Total")) & "%"
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
					''FLM:20090504:EFECTOS	
                    elseif rst("Tipo")="E" then
						primeralinea=0

						if viene<>"tienda" then
							dat1=Hiperv(OBJEfectosPro,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(rst("nfactura_pro")) & " ("&LitEfecto&")",LitVerEfectoPro)
						else
							dat1="<a class='CELDAREFB' href=javascript:ver_documento('" & enc.EncodeForJavascript(rst("ndoc")) & "','efectos','"&enc.EncodeForJavascript(nproveedor)&"') alt='" & LitVerEfectoPro & "'>" & EncodeForHtml(rst("nfactura_pro")) & " ("&LitEfecto&")</a>"
						end if
                        'd_lookup("fechavto","efectos_pro","nefecto='" & rst("ndoc") & "'",session("dsn_cliente"))
                        strselect="select fechavto from efectos_pro where nefecto=?"
						DrawCelda "tdbordeCELDA7","","",0,dat1
						DrawCelda "tdbordeCELDA7 ","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "tdbordeCELDA7 ","","",0,""& EncodeForHtml(DLookupP1(strselect, rst("ndoc") &"", adVarChar, 20,  session("dsn_cliente")))
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1) & " "  & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						total_valor = total_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						total_pendiente = total_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
						Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
					end if

					CloseFila
					ProveedorAnt=rst("nproveedor")

					cuentaAnt=iif(nulear(rst("ncuentacargo"))>"",nulear(rst("ncuentacargo")),LitSinAsignar)

					fila=fila+1
					fila2=fila2+1
					rst.MoveNext
				wend

				if lote=lotes then
				    '' FACTURAS
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and pagada=0 and facturas_pro.deuda<>0 and facturas_pro.nproveedor='" & ProveedorAnt & "'"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    total_valor=0
				    total_pendiente=0
				    total_pendienteV=0
				    while not rstAux.eof
					    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    registrosT=rstAux.recordcount
				    rstAux.close

				    '' VENCIMIENTOS
				    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & replace(strwhere,"facturas_pro.fecha","vencimientos_entrada.fecha") & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 and vencimientos_entrada.importe<>0 and facturas_pro.nproveedor='" & ProveedorAnt & "'"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
				        registrosT=registrosT+1
				        total_pendienteV = total_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
				        rstAux.movenext
				    wend
				    rstAux.close

				    '' IRPF
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and facturas_pro.total_irpf<>0 and pagada=0 and facturas_pro.total_irpf<>0 and facturas_pro.nproveedor='" & ProveedorAnt & "'"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
					    registrosT=registrosT+1
					    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    rstAux.close
				   ' FLM:20090504:EFECTOS
				   if nserieEfec&"">"" then
				        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
				        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
				        if cuentaAnt<>LitSinAsignar then
    			            seleccion = seleccion & ", bancos as banc with(NOLOCK) "
	    		        end if
		    	        seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and efectos_pro.nproveedor='" & ProveedorAnt & "' and pendiente=1 and efectos_pro.importe<>0"
			            if cuentaAnt<>LitSinAsignar then
			                seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
			            end if
			            seleccion = seleccion & strwhereEfec3
				        rstAux.cursorlocation=3
				        rstAux.open seleccion,session("dsn_cliente")
				        while not rstAux.eof
				            registrosT=registrosT+1
				            total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					        total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					        rstAux.movenext
				        wend
				        rstAux.close
				    end if					  
                    DrawFila color_fondo
						if opcagruparcuenta="1" then
							DrawCelda "tdbordeCELDA7","","",0,""
						end if
						DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalP & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
					CloseFila
                    
                    'if not nserieEfec&"">"" then 'CONDICION PARA EFECTOS/FACTURAS
					    '' FACTURAS
					    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
					    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and pagada=0 and facturas_pro.deuda<>0"
					    rstAux.cursorlocation=3
					    rstAux.open seleccion,session("dsn_cliente")
					    Gtotal_valor=0
					    Gtotal_pendiente=0
					    Gtotal_pendienteV=0
					    while not rstAux.eof
						    Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
						    Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
						    rstAux.movenext
					    wend
					    registrosT=rstAux.recordcount
					    rstAux.close

					    '' VENCIMIENTOS
					    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
					    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & replace(strwhere,"facturas_pro.fecha","vencimientos_entrada.fecha") & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 and vencimientos_entrada.importe<>0 "
					    rstAux.cursorlocation=3
					    rstAux.open seleccion,session("dsn_cliente")
					    while not rstAux.eof
					        registrosT=registrosT+1
					        Gtotal_pendienteV = Gtotal_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					        rstAux.movenext
					    wend
					    rstAux.close

					    '' IRPF
					    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
					    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and facturas_pro.total_irpf<>0 and pagada=0"
					    rstAux.cursorlocation=3
					    rstAux.open seleccion,session("dsn_cliente")
					    while not rstAux.eof
						    registrosT=registrosT+1
						    Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
						    rstAux.movenext
					    wend
					    rstAux.close
					    ' FLM:20090504:Efectos
					    if nserieEfec&"">"" then
					        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
					        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
    				        if cuentaAnt<>LitSinAsignar then
	    			            seleccion = seleccion & ", bancos as banc with(NOLOCK) "
		    		        end if
			    	        seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and pendiente=1 and efectos_pro.importe<>0"
				            if cuentaAnt<>LitSinAsignar then
				                seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
				            end if
				            seleccion = seleccion & strwhereEfec3

					        rstAux.cursorlocation=3
					        rstAux.open seleccion,session("dsn_cliente")
					        while not rstAux.eof
					            registrosT=registrosT+1
						        Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
						        Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
						        rstAux.movenext
					        wend
					        rstAux.close	
					    end if				    
					DrawFila color_fondo
						if opcagruparcuenta="1" then
							DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,""
						end if
						DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitRegParcial & " : " & EncodeForHtml(registrosT)
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitTotales & "(" & EncodeForHtml(abreviaturaMB) & ")"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_valor),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
					CloseFila
					'EQUIVALENCIA EN PTAS
					if mostrar_equivalencia then
						DrawFila color_fondo
							if opcagruparcuenta="1" then
								DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,""
							end if
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitTotales & "(" & EncodeForHtml(abreviaturaPtas) & ")"
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(formatnumber(null_z(Gtotal_valor),n_decimalesMB,-1,0,-1)),moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1)) & "</b>"
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)),moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1)) & "</b>"
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(formatnumber(null_z(Gtotal_pendienteV),n_decimalesMB,-1,0,-1)),moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1)) & "</b>"
						CloseFila
					end if
				end if
				rst.Close
			end if    ' ********* FIN DEL CASO NO AGRUPACIONES

			' **************  CASO DE AGRUPACION POR CUENTA

			if opcagruparcuenta="1" and opcagrupardia<>"1" then
				while not rst.EOF and fila<=MAXPAGINA
					if rst("nproveedor")<>ProveedorAnt then
						if ProveedorAnt<>"" then
							if cuentaAnt=LitSinAsignar then
								strwhere3 = " and facturas_pro.ncuenta is null "	
								strwhereEfec3=" and efectos_pro.banco is null "
							else
								strwhere3 = " and facturas_pro.ncuenta='" & cuentaAnt & "' "
								strwhereEfec3=" and banc.ncuenta='" & cuentaAnt & "' "
							end if

						    '' FACTURAS
						    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    total_valor=0
						    total_pendiente=0
						    total_pendienteV=0
						    while not rstAux.eof
							    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
							    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    registrosT=rstAux.recordcount
						    rstAux.close

						    '' VENCIMIENTOS
						    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & strwhereVenc & strwhere3 & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 and facturas_pro.nproveedor='" & ProveedorAnt & "'"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    while not rstAux.eof
							    total_pendienteV = total_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    registrosT=registrosT + rstAux.recordcount
						    rstAux.close

						    '' IRPF
						    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.total_irpf<>0 and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    while not rstAux.eof
							    registrosT=registrosT+1
							    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    rstAux.close
							
						    '' FLM:20090504:EFECTOS
						    if nserieEfec&"">"" then ''and cuentaAnt=LitSinAsignar then
						        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
						        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
						        if cuentaAnt<>LitSinAsignar then
						            seleccion = seleccion & ", bancos as banc with(NOLOCK) "
						        end if
						        seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and efectos_pro.nproveedor='" & ProveedorAnt & "' and pendiente=1"
						        if cuentaAnt<>LitSinAsignar then
						            seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
						        end if
						        seleccion = seleccion & strwhereEfec3
						        rstAux.cursorlocation=3
						        rstAux.open seleccion,session("dsn_cliente")
						        while not rstAux.eof
						            registrosT=registrosT+1
							        total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
							        total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							        rstAux.movenext
						        wend							    
						        rstAux.close
						    end if

							DrawFila color_fondo
								DrawCelda "tdbordeCELDA7","","",0,""
								DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalP & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
							CloseFila

							if nulear(rst("ncuentacargo"))<>CuentaAnt then
								if cuentaAnt=LitSinAsignar then
									strwhere3 = " and facturas_pro.ncuenta is null "	
									strwhereEfec3=" and efectos_pro.banco is null "
								else
									strwhere3 = " and facturas_pro.ncuenta='" & cuentaAnt & "' "
									strwhereEfec3=" and banc.ncuenta='" & cuentaAnt & "' "
								end if
								
							    '' FACTURAS
							    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
							    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and pagada=0"
							    rstAux.cursorlocation=3
							    rstAux.open seleccion,session("dsn_cliente")
							    total_valor=0
							    total_pendiente=0
							    total_pendienteV=0
							    while not rstAux.eof
								    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
								    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
								    rstAux.movenext
							    wend
							    registrosT=rstAux.recordcount
							    rstAux.close

							    '' VENCIMIENTOS
							    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
							    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & strwhereVenc & strwhere3 & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 "
							    rstAux.cursorlocation=3
							    rstAux.open seleccion,session("dsn_cliente")
							    while not rstAux.eof
								    total_pendienteV = total_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
								    rstAux.movenext
							    wend
							    registrosT=registrosT + rstAux.recordcount
							    rstAux.close

							    '' IRPF
							    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
							    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.total_irpf<>0 and pagada=0"
							    rstAux.cursorlocation=3
							    rstAux.open seleccion,session("dsn_cliente")
							    while not rstAux.eof
								    registrosT=registrosT+1
								    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
								    rstAux.movenext
							    wend
							    rstAux.close

							    '' FLM:20090504:EFECTOS
							    if nserieEfec&"">"" then ''and cuentaAnt=LitSinAsignar then
							        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
							        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
							        if cuentaAnt<>LitSinAsignar then
							            seleccion = seleccion & ", bancos as banc with(NOLOCK) "
							        end if
							        seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and pendiente=1"
							        if cuentaAnt<>LitSinAsignar then
							            seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
							        end if
							        seleccion = seleccion & strwhereEfec3
							        rstAux.cursorlocation=3
							        rstAux.open seleccion,session("dsn_cliente")
							        while not rstAux.eof
							            registrosT=registrosT+1
								        total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
								        total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
								        rstAux.movenext
							        wend
							        rstAux.close
							    end if
								DrawFila color_fondo
									DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
									DrawCelda "tdbordeCELDA7","","",0,""
									DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalC & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
									DrawCelda "tdbordeCELDA7 ","","",0,""
									DrawCelda "tdbordeCELDA7 ","","",0,""
									DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
									DrawCelda "tdbordeCELDA7 ","","",0,""
									DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
									DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
								CloseFila
							end if

							fila2=1
							if fila+1<=MAXPAGINA then
								'dejamos ahora dos espacios en blanco
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								DrawFila ""
								CloseFila
								'Fila de encabezado
								DrawFila color_fondo
									DrawCelda "ENCABEZADOL7 width='125px' style='BORDER: 1px solid Black;'","","",0,LitNCuenta
									DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitProveedor
									DrawCelda "ENCABEZADOL7 style='width:150px;BORDER: 1px solid Black;'","","",0,LitFacVen
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitFechaF
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitFechaV
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitValorF
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitValorV
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitPendienteF
									DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitPendienteV
								CloseFila
							end if
						else
							DrawFila ""
						end if

						total_valor=0
						total_pendiente=0
                        total_pendienteV=0
						Cuenta=iif(nulear(rst("ncuentacargo"))>"",nulear(rst("ncuentacargo")),LitSinAsignar)
						if nulear(rst("ncuentacargo"))<>CuentaAnt or Cuenta=LitSinAsignar then
							DrawCelda "ENCABEZADOL7","","",0,EncodeForHtml(Cuenta)
						else
							DrawCelda "CELDA","","",0,"&nbsp;"
						end if

						if viene<>"tienda" then
							dat1=Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(rst("razon_social")),LitVerProveedor)
						else
							dat1=EncodeForHtml(rst("razon_social"))
						end if
						DrawCelda "ENCABEZADOL7","","",0,dat1
					else
						if nulear(rst("ncuentacargo"))<>CuentaAnt and CuentaAnt<>"" then
							if cuentaAnt=LitSinAsignar then
								strwhere3 = " and facturas_pro.ncuenta is null "	
								strwhereEfec3=" and efectos_pro.banco is null "
							else
								strwhere3 = " and facturas_pro.ncuenta='" & cuentaAnt & "' "
								strwhereEfec3=" and banc.ncuenta='" & cuentaAnt & "' "
							end if

                           '' FACTURAS
						    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    total_valor=0
						    total_pendiente=0
						    total_pendienteV=0
						    while not rstAux.eof
							    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
							    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    registrosT=rstAux.recordcount
						    rstAux.close

						    '' VENCIMIENTOS
						    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & strwhereVenc & strwhere3 & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 and facturas_pro.nproveedor='" & ProveedorAnt & "'"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    while not rstAux.eof
							    total_pendienteV = total_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    registrosT=registrosT + rstAux.recordcount
						    rstAux.close

						    '' IRPF
						    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.total_irpf<>0 and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    while not rstAux.eof
							    registrosT=registrosT+1
							    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    rstAux.close

						    '' FLM:20090504:EFECTOS
						    if nserieEfec&"">"" then ''and cuentaAnt=LitSinAsignar then
						        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total,importe AS Deuda"
						        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock)"
						        if cuentaAnt<>LitSinAsignar then
						            seleccion = seleccion & ", bancos as banc with(NOLOCK) "
						        end if
						        seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and efectos_pro.nproveedor='" & ProveedorAnt & "' and pendiente=1"
						        if cuentaAnt<>LitSinAsignar then
						            seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
						        end if
						        seleccion = seleccion & strwhereEfec3
							        
						        rstAux.cursorlocation=3
						        rstAux.open seleccion,session("dsn_cliente")
						        while not rstAux.eof
						            registrosT=registrosT+1
							        total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
							        total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							        rstAux.movenext
						        wend
						        rstAux.close
						    end if
							DrawFila color_fondo
								DrawCelda "tdbordeCELDA7","","",0,""
								DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalP & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
							CloseFila
							if cuentaAnt=LitSinAsignar then
								strwhere3 = " and facturas_pro.ncuenta is null "	
								strwhereEfec3=" and efectos_pro.banco is null "
							else
								strwhere3 = " and facturas_pro.ncuenta='" & cuentaAnt & "' "
								strwhereEfec3=" and banc.ncuenta='" & cuentaAnt & "' "
							end if

                           '' FACTURAS
						    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and pagada=0"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    total_valor=0
						    total_pendiente=0
						    total_pendienteV=0
						    while not rstAux.eof
							    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
							    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    registrosT=rstAux.recordcount
						    rstAux.close

						    '' VENCIMIENTOS
						    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & strwhereVenc & strwhere3 & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 "
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    while not rstAux.eof
							    total_pendienteV = total_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    registrosT=registrosT + rstAux.recordcount
						    rstAux.close

						    '' IRPF
						    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
						    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.total_irpf<>0 and pagada=0"
						    rstAux.cursorlocation=3
						    rstAux.open seleccion,session("dsn_cliente")
						    while not rstAux.eof
							    registrosT=registrosT+1
							    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							    rstAux.movenext
						    wend
						    rstAux.close

						    '' FLM:20090504:EFECTOS
						    if nserieEfec&"">"" then ''and cuentaAnt=LitSinAsignar then
						        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
						        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock)"
					            if cuentaAnt<>LitSinAsignar then
					                seleccion = seleccion & ", bancos as banc with(NOLOCK) "
					            end if
					            seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and pendiente=1"
					            if cuentaAnt<>LitSinAsignar then
					                seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
					            end if
					            seleccion = seleccion & strwhereEfec3
						        rstAux.cursorlocation=3
						        rstAux.open seleccion,session("dsn_cliente")
						        while not rstAux.eof
						            registrosT=registrosT+1
							        total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
							        total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
							        rstAux.movenext
						        wend
						        rstAux.close
						    end if

							DrawFila color_fondo
								DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
								DrawCelda "tdbordeCELDA7","","",0,""
								DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalC & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 ","","",0,""
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
							CloseFila
						end if

						Cuenta=iif(nulear(rst("ncuentacargo"))>"",nulear(rst("ncuentacargo")),LitSinAsignar)
						if nulear(rst("ncuentacargo"))<>CuentaAnt then
							DrawCelda "ENCABEZADOL7","","",0,EncodeForHtml(Cuenta)
						else
							DrawCelda "CELDA","","",0,"&nbsp;"
						end if

						DrawCelda "CELDA","","",0,"&nbsp;"
					end if

					if rst("Tipo")="F" then		'' FACTURAS
						primeralinea=0
						if viene<>"tienda" then
							dat1=Hiperv(OBJFacturasPro,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(rst("nfactura_pro")),LitVerFactura)
						else
							dat1="<a class='CELDAREFB' href=javascript:ver_documento('" & enc.EncodeForJavascript(rst("ndoc")) & "','facturas','"&enc.EncodeForJavascript(nproveedor)&"') alt='" & LitVerFacturas & "'>" & EncodeForHtml(rst("nfactura_pro")) & "</a>"
						end if
						DrawCelda "tdbordeCELDA7","","",0,dat1
						DrawCelda "tdbordeCELDA7 ","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1) & " "  & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						total_valor = total_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						total_pendiente = total_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
						Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
					elseif rst("Tipo")="V" then		'' VENCIMIENTOS
                        'd_lookup("nfactura_pro","facturas_pro","nfactura='" & rst("ndoc") & "'",session("dsn_cliente"))
                        strselect= "select factura_pro from facturas_pro where nfactura =?"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(DLookupP1(strselect,rst("ndoc") &"",adVarChar,20,session("dsn_cliente")) & "-" & rst("Nvto"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 ","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "tdbordeCELDA7","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))
					elseif rst("Tipo")="VI" then	'' VENCIMIENTOS
						DrawCelda "tdbordeCELDA7 align='right'","","",0,LitIRPF & ": " & EncodeForHtml(rst("Total")) & "%"
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1) & " "  & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7","","",0,"&nbsp;"
					''FLM:20090504:EFECTOS
					elseif rst("Tipo")="E" then
						primeralinea=0
						if viene<>"tienda" then
							dat1=Hiperv(OBJEfectosPro,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(rst("nfactura_pro"))& " ("&LitEfecto&")",LitVerEfectoPro)
						else
							dat1="<a class='CELDAREFB' href=javascript:ver_documento('" & enc.EncodeForJavascript(rst("ndoc")) & "','efectos','"&enc.EncodeForJavascript(nproveedor)&"') alt='" & LitVerEfectoPro & "'>" & EncodeForHtml(rst("nfactura_pro"))& " ("&LitEfecto&")" & "</a>"
						end if
                        'd_lookup("fechavto","efectos_pro","nefecto='" & rst("ndoc") & "'",session("dsn_cliente"))
                        strselect="select fechavto from efectos_pro where nefecto=?"
						DrawCelda "tdbordeCELDA7","","",0,dat1
						DrawCelda "tdbordeCELDA7 ","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "tdbordeCELDA7 ","","",0,""& EncodeForHtml(DLookupP1(strselect, rst("ndoc") &"", adVarChar, 20,  session("dsn_cliente")))
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1) & " "  & rst("abreviatura"))
						DrawCelda "tdbordeCELDA7 ","","",0,"&nbsp;"
						total_valor = total_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						total_pendiente = total_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
						Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rst("Total")),rst("divisa"),moneda_base)
						Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rst("deuda")),rst("divisa"),moneda_base)
					end if

					CloseFila
					ProveedorAnt=rst("nproveedor")

					cuentaAnt=iif(nulear(rst("ncuentacargo"))<>"",nulear(rst("ncuentacargo")),LitSinAsignar)

					fila=fila+1
					fila2=fila2+1
					rst.MoveNext
				wend

				if lote=lotes then
					if cuentaAnt=LitSinAsignar then
						strwhere3 = " and facturas_pro.ncuenta is null "	
						strwhereEfec3=" and efectos_pro.banco is null "
					else
						strwhere3 = " and facturas_pro.ncuenta='" & cuentaAnt & "' "
						strwhereEfec3=" and banc.ncuenta='" & cuentaAnt & "' "
					end if

				    '' FACTURAS
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    total_valor=0
				    total_pendiente=0
				    total_pendienteV=0
				    while not rstAux.eof
					    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    registrosT=rstAux.recordcount
				    rstAux.close

				    '' VENCIMIENTOS
				    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & strwhereVenc & strwhere3 & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 and facturas_pro.nproveedor='" & ProveedorAnt & "'"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
					    total_pendienteV = total_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    registrosT=registrosT + rstAux.recordcount
				    rstAux.close

				    '' IRPF
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.total_irpf<>0 and facturas_pro.nproveedor='" & ProveedorAnt & "' and pagada=0"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
					    registrosT=registrosT+1
					    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    rstAux.close

				    '' FLM:20090504:EFECTOS
				    if nserieEfec&"">"" then ''and cuentaAnt=LitSinAsignar then
				        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
				        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock)"
			            if cuentaAnt<>LitSinAsignar then
			                seleccion = seleccion & ", bancos as banc with(NOLOCK) "
			            end if
			            seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and efectos_pro.nproveedor='" & ProveedorAnt & "' and pendiente=1"
			            if cuentaAnt<>LitSinAsignar then
			                seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
			            end if
			            seleccion = seleccion & strwhereEfec3
				        rstAux.cursorlocation=3
				        rstAux.open seleccion,session("dsn_cliente")

				        while not rstAux.eof
				            registrosT=registrosT+1
					        total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					        total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					        rstAux.movenext
				        wend
				        rstAux.close
				    end if
                    
					DrawFila color_fondo
						DrawCelda "tdbordeCELDA7","","",0,""
						DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalP & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
					CloseFila

              	    '' FACTURAS
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and pagada=0"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    total_valor=0
				    total_pendiente=0
				    total_pendienteV=0
				    while not rstAux.eof
					    total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    registrosT=rstAux.recordcount
				    rstAux.close

				    '' VENCIMIENTOS
				    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & strwhereVenc & strwhere3 & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 "
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
					    total_pendienteV = total_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    registrosT=registrosT + rstAux.recordcount
				    rstAux.close

				    '' IRPF
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & strwhere3 & " and facturas_pro.total_irpf<>0 and pagada=0"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
					    registrosT=registrosT+1
					    total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    rstAux.close

				    '' FLM:20090504:EFECTOS
				    if nserieEfec&"">"" then ''and cuentaAnt=LitSinAsignar then
				        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
				        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
			            if cuentaAnt<>LitSinAsignar then
			                seleccion = seleccion & ", bancos as banc with(NOLOCK) "
			            end if
			            seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and pendiente=1"
			            if cuentaAnt<>LitSinAsignar then
			                seleccion = seleccion & " and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
			            end if
			            seleccion = seleccion & strwhereEfec3
				        rstAux.cursorlocation=3
				        rstAux.open seleccion,session("dsn_cliente")
    				    
				        while not rstAux.eof
				            registrosT=registrosT+1
					        total_valor = total_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					        total_pendiente = total_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					        rstAux.movenext
				        wend
				        rstAux.close
				    end if
					DrawFila color_fondo
						DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(registrosT) & "</b>"
						DrawCelda "tdbordeCELDA7","","",0,""
						DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalC & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_valor),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(total_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
					CloseFila

                    '' FACTURAS
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, total_factura AS Total, deuda AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and pagada=0"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    Gtotal_valor=0
				    Gtotal_pendiente=0
				    Gtotal_pendienteV=0
				    while not rstAux.eof
					    Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					    Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    registrosT=rstAux.recordcount
				    rstAux.close

				    '' VENCIMIENTOS
				    seleccion=" SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, importe AS Total,importe AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),vencimientos_entrada with(nolock),proveedores with(nolock) " & strwhereVenc & " and facturas_pro.nfactura = vencimientos_entrada.nfactura and pagado=0 "
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
					    Gtotal_pendienteV = Gtotal_pendienteV + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    registrosT=registrosT + rstAux.recordcount
				    rstAux.close

				    '' IRPF
				    seleccion="SELECT divisas.abreviatura,divisas.ndecimales,facturas_pro.divisa as divisa, facturas_pro.irpf AS Total, facturas_pro.total_irpf AS Deuda"
				    seleccion=seleccion & " FROM divisas with(nolock),FACTURAS_PRO with(nolock),proveedores with(nolock) " & strwhere & " and facturas_pro.total_irpf<>0 and pagada=0"
				    rstAux.cursorlocation=3
				    rstAux.open seleccion,session("dsn_cliente")
				    while not rstAux.eof
					    registrosT=registrosT+1
					    Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					    rstAux.movenext
				    wend
				    rstAux.close

				    '' FLM:20090504:EFECTOS
				    if nserieEfec&"">""  then
				        seleccion="SELECT divisas.abreviatura,divisas.ndecimales,efectos_pro.divisa as divisa, importe AS Total, importe AS Deuda"
				        seleccion=seleccion & " FROM divisas with(nolock),efectos_pro with(nolock) "
			            ''if cuentaAnt<>LitSinAsignar then
			            if ncuentacargo>"" then
			                seleccion = seleccion & ", bancos as banc with(NOLOCK) "
			            end if
			            seleccion = seleccion & " ,proveedores with(nolock) " & strwhereEfec &  " and pendiente=1"
			            ''if cuentaAnt<>LitSinAsignar then
			            if ncuentacargo>"" then
			                seleccion = seleccion & "  and banc.nbanco like '" & session("ncliente") & "%' and banc.nbanco=efectos_pro.banco "
			            end if
			            ''seleccion = seleccion & strwhereEfec3
				        rstAux.cursorlocation=3
				        rstAux.open seleccion,session("dsn_cliente")

				        while not rstAux.eof
				            registrosT=registrosT+1
					        Gtotal_valor = Gtotal_valor + CambioDivisa(null_z(rstAux("Total")),rstAux("divisa"),moneda_base)
					        Gtotal_pendiente = Gtotal_pendiente + CambioDivisa(null_z(rstAux("deuda")),rstAux("divisa"),moneda_base)
					        rstAux.movenext
				        wend				        	    
				        rstAux.close
				    end if
					DrawFila color_fondo
						DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitRegParcial & " : " & EncodeForHtml(registrosT)
						DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,""
						DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitTotales & "(" & EncodeForHtml(abreviaturaMB) & ")"
						'DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_valor),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,""
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteV),n_decimalesMB,-1,0,-1)) & "</b>"
					CloseFila
					'EQUIVALENCIA EN PTAS
					if mostrar_equivalencia then
						DrawFila color_fondo
							if opcagruparcuenta="1" then
								DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,""
							end if
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "ENCABEZADOR7 style='BORDER: 1px solid Black;'","","",0,LitTotales & "(" & EncodeForHtml(abreviaturaPtas) & ")"
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(formatnumber(null_z(Gtotal_valor),n_decimalesMB,-1,0,-1)),moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1)) & "</b>"
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)),moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1)) & "</b>"
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(formatnumber(null_z(Gtotal_pendienteV),n_decimalesMB,-1,0,-1)),moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1)) & "</b>"
						CloseFila
					end if
				end if
				rst.Close
			end if   ' ******* FIN DEL CASO DE AGRUPACION POR CUENTA

			' ******************  CASO DE AGRUPACION POR CUENTA Y DIA

			if opcagruparcuenta="1" and opcagrupardia="1" then
				fila=1

				if not rst.eof then
					ncuenta_old=rst("ncuenta") & ""
				end if

				while not rst.EOF and fila<=MAXPAGINA
					if rst("ncuenta")<>ncuenta_old or (rst("ncuenta")&""="" and ncuenta_old&"">"") or (rst("ncuenta")&"">"" and ncuenta_old&""="") then
						DrawFila color_fondo
							DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(ultimo_contador_ncuenta) & "</b>"
							DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalC & "</b>"
							if rst("HayDeudaIRPF")<>"0" then
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deuda_ncuenta+ultima_deudaIRPF_ncuenta),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB) & "</b>"
							end if
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deuda_ncuenta),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB) & "</b>"
							if rst("HayDeudaIRPF")<>"0" then
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deudaIRPF_ncuenta),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB) & "</b>"
							end if
						CloseFila

						if fila+1<=MAXPAGINA then
							''dejamos ahora dos espacios en blanco
							DrawFila ""
							CloseFila
							DrawFila ""
							CloseFila
							DrawFila ""
							CloseFila
							DrawFila ""
							CloseFila
							DrawFila ""
							CloseFila
							DrawFila ""
							CloseFila
							DrawFila color_fondo
								DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitNCuenta
								if rst("HayDeudaIRPF")<>"0" then
									DrawCelda "ENCABEZADOL7 colspan=2 style='BORDER: 1px solid Black;'","","",0,LitFecha
								else
									DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitFecha
								end if
								DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitImporte
								if rst("HayDeudaIRPF")<>"0" then
									DrawCelda "ENCABEZADOL7 style='BORDER: 1px solid Black;'","","",0,LitIRPF
								end if
							CloseFila
						end if

						DrawFila color_blau
						if rst("ncuenta") & "">"" then
							DrawCelda "ENCABEZADOL7","","",0,EncodeForHtml(rst("ncuenta"))
						else
							DrawCelda "ENCABEZADOL7","","",0,LitSinAsignar
						end if
					else
						DrawFila color_blau
						if fila>1 then
							DrawCelda "ENCABEZADOL7","","",0,""
						else
							if rst("ncuenta") & "">"" then
								DrawCelda "ENCABEZADOL7","","",0,EncodeForHtml(rst("ncuenta"))
							else
								DrawCelda "ENCABEZADOL7","","",0,LitSinAsignar
							end if
						end if
					end if
					if rst("HayDeudaIRPF")<>"0" then
						DrawCelda "tdbordeCELDA7 colspan=2","","",0,EncodeForHtml(rst("fecha"))
					else
						DrawCelda "tdbordeCELDA7 ","","",0,EncodeForHtml(rst("fecha"))
					end if
					DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB)
					if rst("HayDeudaIRPF")<>"0" then
						DrawCelda "tdbordeCELDA7 align='right'","","",0,EncodeForHtml(formatnumber(null_z(rst("deudaIRPF")),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB)
					end if
					CloseFila

					if not rst.eof then
						ncuenta_old=rst("ncuenta")
						ultima_deuda_ncuenta=rst("deuda_ncuenta")
						if rst("HayDeudaIRPF")<>"0" then
							ultima_deudaIRPF_ncuenta=rst("deudaIRPF_ncuenta")
						end if
						ultimo_contador_ncuenta=rst("cuenta_ncuenta")
					end if

					fila=fila+1
					rst.movenext
				wend

				if lote=lotes then
					if rst.eof then
						rst.movefirst
					end if
					DrawFila color_fondo
						DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(ultimo_contador_ncuenta) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalC & "</b>"
						if rst("HayDeudaIRPF")<>"0" then
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deuda_ncuenta+ultima_deudaIRPF_ncuenta),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB) & "</b>"
						end if
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deuda_ncuenta),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB) & "</b>"
						if rst("HayDeudaIRPF")<>"0" then
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deudaIRPF_ncuenta),n_decimalesMB,-1,0,-1) & " "  & abreviaturaMB) & "</b>"
						end if
					CloseFila
					'dejamos ahora dos espacios en blanco
					DrawFila ""
					CloseFila
					DrawFila ""
					CloseFila
					DrawFila ""
					CloseFila
					DrawFila ""
					CloseFila
					DrawFila ""
					CloseFila
					DrawFila ""
					CloseFila
					DrawFila color_fondo
						DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegTotal & " : " & EncodeForHtml(rst("contador")) & "</b>"
						DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotales & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
						if rst("HayDeudaIRPF")<>"0" then
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(cdbl(rst("deuda_total"))+cdbl(rst("deudaIRPF_total"))),n_decimalesMB,-1,0,-1) & " " & abreviaturaMB) & "</b>"
						end if
						DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(rst("deuda_total")),n_decimalesMB,-1,0,-1) & " " & abreviaturaMB) & "</b>"
						if rst("HayDeudaIRPF")<>"0" then
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(rst("deudaIRPF_total")),n_decimalesMB,-1,0,-1) & " " & abreviaturaMB) & "</b>"
						end if
					CloseFila
					if mostrar_equivalencia then
						DrawFila color_fondo
							DrawCelda "tdbordeCELDA7","","",0,""
							DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotales & "(" & EncodeForHtml(abreviaturaPtas) & ")</b>"
							if rst("deuda_total") & "">"" then
								deuda_total=cdbl(rst("deuda_total"))
								deudaIRPF_total=cdbl(rst("deudaIRPF_total"))
							else
								deuda_total=0
								deudaIRPF_total=0
							end if
							if rst("HayDeudaIRPF")<>"0" then
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(deuda_total+deudaIRPF_total,moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1) & " " & abreviaturaPtas) & "</b>"
							end if
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(deuda_total,moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1) & " " & abreviaturaPtas) & "</b>"
							if rst("HayDeudaIRPF")<>"0" then
								if rst("deudaIRPF_total") & "">"" then
									deudaIRPF_total=cdbl(rst("deudaIRPF_total"))
								else
									deudaIRPF_total=0
								end if
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(deudaIRPF_total,moneda_base,moneda_Ptas),n_decimalesPtas,-1,0,-1) & " " & abreviaturaPtas) & "</b>"
							end if
						CloseFila
					end if
				else
					if rst("ncuenta")<>ncuenta_old or (rst("ncuenta")&""="" and ncuenta_old&"">"") or (rst("ncuenta")&"">"" and ncuenta_old&""="") then
						DrawFila color_fondo
							DrawCelda "tdbordeCELDA7","","",0,"<b>" & LitRegParcial & " : " & EncodeForHtml(ultimo_contador_ncuenta) & "</b>"
							DrawCelda "tdbordeCELDA7 ","","",0,"<b>" + LitTotalC & "</b>"
							if rst("HayDeudaIRPF")<>"0" then
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deuda_ncuenta+ultima_deudaIRPF_ncuenta),n_decimalesMB,-1,0,-1) & " " & abreviaturaMB) & "</b>"
							end if
							DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deuda_ncuenta),n_decimalesMB,-1,0,-1) & " " & abreviaturaMB) & "</b>"
							if rst("HayDeudaIRPF")<>"0" then
								DrawCelda "tdbordeCELDA7 align='right'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(ultima_deudaIRPF_ncuenta),n_decimalesMB,-1,0,-1) & " " & abreviaturaMB) & "</b>"
							end if
						CloseFila
					end if
				end if

				rst.Close
				conn.close
			end if    ' ********* FIN DEL CASO AGRUPACIONES POR CUENTA Y DIA

			%></table><%
			if lotes>1 then
				%><hr/><%
			end if
			NavPaginas lote,lotes,campo,criterio,texto,2
		else
		    'FLM:20090504:no hace falta ya.
		    ' rst.Close 
			%><input type="hidden" name="NumRegsTotal" value="0">
			<script>
			    window.alert("<%=LitMsgNoDocumentos%>");
			    parent.window.frames["botones"].document.location = "listado_pagos_bt.asp?mode=select1";
			</script><%
			if viene="tienda" then
			else%>
				<script>
					document.location="listado_pagos.asp?mode=select1";
					
				</script><%
			end if
		end if
        ''ricardo 25-5-2006 comienzo de la select
        ''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
        auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"fin_listado_pagos"
	end if%>
</form>
<%set rst=nothing
set conn=nothing
end if
connRound.close
set connRound = Nothing
set rstSelect = Nothing
set rstAux = Nothing
set rstProveedor = Nothing
set rstVencimientos = Nothing
set rstAgrupar = Nothing
set rstPendiente = Nothing
%>
</body>
</html>
