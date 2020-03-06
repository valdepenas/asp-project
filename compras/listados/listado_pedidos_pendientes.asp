<%@ Language=VBScript %>
<%
  '***RGU 20/12/2005: Añadir al listado columnas de "cantidad pendiente", "Pedido Cliente", "Item"(cliente)
  'RGU 17/11/2007 CAMBIO DSN PARA LISTADOS

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=titulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
<META HTTP-EQUIV="Content-style-TypeCONTENT="text/css">
<LINK REL="styleSHEET" href="../../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../../impresora.css" MEDIA="PRINT">
</head>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<% if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../modulos.inc" -->
<!--#include file="../pedidos_pro.inc" -->
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../../styles/formularios.css.inc" -->  
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->

<%if request.querystring("viene")="tienda" or request.form("viene")="tienda" then
	titulo=LitTituloListadoPen2
else
	titulo=LitTituloListadoPen
end if%>

<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">

function keypress2(){
	tecla=window.event.keyCode;
	//keyPressed(tecla);
}

function keyPressed(tecla) {
}

function traerreferencia(){
}

//Desencadena la búsqueda del proveedor cuya referencia se indica
function TraerProveedor(mode) {
	document.listado_pedidos_pendientes.action="listado_pedidos_pendientes.asp?mode=traerproveedor"
	document.listado_pedidos_pendientes.submit();
}

function ver_pedido(npedido,nproveedor){
	if (nproveedor!=""){
		parent.document.location="../../compras/pedidos_pro_imp.asp?npedido=('" + npedido + "')&mode=browse&empresa="+nproveedor.substr(0,5);
	}else{
		parent.document.location="../../compras/pedidos_pro_imp.asp?npedido=('" + npedido + "')&mode=browse&empresa=<%=session("ncliente")%>";
	}
	parent.parent.topFrame.document.all("regresar").style.display="";
}
function ver_pedidoCli(npedido){
	parent.document.location="../../ventas/pedidos_cli_imp.asp?npedido=('" + npedido + "')&mode=browse&empresa=<%=session("ncliente")%>";
	parent.parent.topFrame.document.all("regresar").style.display="";
}

function Ver_Articulo(referencia,nproveedor){
	pagina="../../tiendas/pedido_ficha_articulo.asp?referencia=" + referencia + "&viene=listado_pedidos_pendientes&nproveedor=" + nproveedor;
	ven=AbrirVentana(pagina,'C',250,400);

}
</script>

<body onload="self.status='';" class="BODY_ASP">
<%
'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
'campo: Nombre del campo con el cual se realizará la búsqueda
'criterio: Tipo de búsqueda
'texto: Texto a buscar.
function CadenaBusqueda(campo,criterio,texto)
	if texto > "" then
		select case criterio
			case "contiene"
				CadenaBusqueda=" where pedidos_pro." + campo + " like '%" + texto + "%' and"
			case "empieza"
				CadenaBusqueda=" where pedidos_pro." + campo + " like '" + texto + "%' and"
			case "termina"
				CadenaBusqueda=" where pedidos_pro." + campo + " like '%" + texto + "' and"
			case "igual"
				CadenaBusqueda=" where pedidos_pro." + campo + "='" + texto + "' and"
		end select
	else
		CadenaBusqueda=" where "
	end if
end function
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
const borde=0

 %>
<form name="listado_pedidos_pendientes" method="post">
<% PintarCabecera "listado_pedidos_pendientes.asp"
'Leer parámetros de la página
	mode		= Request.QueryString("mode")
	if mode="browse" then mode="imp"
	nproveedor	= limpiaCadena(Request.QueryString("nproveedor"))
	if nproveedor ="" then
		nproveedor	= limpiaCadena(Request.form("nproveedor"))
	end if

	if nproveedor="" then
		nproveedor	= limpiaCadena(Request.QueryString("ndoc"))
		if nproveedor ="" then
			nproveedor	= limpiaCadena(Request.form("ndoc"))
		end if
	end if

	'if nproveedor > "" then
	'	nproveedor = session("ncliente") & completar(nproveedor,5,"0")
	'end if

	actividad	= limpiaCadena(Request.QueryString("actividad"))
	if actividad ="" then
		actividad	= limpiaCadena(Request.form("actividad"))
	end if

	fdesde		= limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde	= limpiaCadena(Request.form("fdesde"))
	end if

	fhasta		= limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta	= limpiaCadena(Request.form("fhasta"))
	end if

	fentregadesde 		= limpiaCadena(Request.QueryString("fdesde"))
	if fentregadesde ="" then
		fentregadesde	= limpiaCadena(Request.form("fentregadesde"))
	end if

	fentregahasta		= limpiaCadena(Request.QueryString("fentregahasta"))
	if fentregahasta="" then
		fentregahasta	= limpiaCadena(Request.form("fentregahasta"))
	end if

	nserie	= limpiaCadena(Request.QueryString("nserie"))
	if nserie ="" then
		nserie	= limpiaCadena(Request.form("nserie"))
	end if
	'CheckCadena nserie

	if request.form("opcproveedorbaja")>"" then
		opcproveedorbaja=limpiaCadena(request.form("opcproveedorbaja"))
	else
		opcproveedorbaja=limpiaCadena(request.querystring("opcproveedorbaja"))
	end if

	if request.form("npedido_ant")>"" then
		npedido2_ant=limpiaCadena(request.form("npedido_ant"))
	else
		npedido2_ant=limpiaCadena(request.querystring("npedido_ant"))
	end if

	if request.form("nproveedor_ant2")>"" then
		nproveedor_ant2=limpiaCadena(request.form("nproveedor_ant2"))
	else
		nproveedor_ant2=limpiaCadena(request.querystring("nproveedor_ant2"))
	end if

	cod_proyecto	= limpiaCadena(Request.QueryString("cod_proyecto"))
	if cod_proyecto="" then
		cod_proyecto	= limpiaCadena(Request.form("cod_proyecto"))
	end if
	'CheckCadena cod_proyecto

	if request.form("viene")>"" then
		viene=limpiaCadena(request.form("viene"))
	else
		viene=limpiaCadena(request.querystring("viene"))
	end if

	campo=limpiaCadena(Request.querystring("campo"))
	criterio=limpiaCadena(Request.querystring("criterio"))
	texto=limpiaCadena(Request.querystring("texto"))

	if request.form("opc_cod_proyecto")>"" then
		opc_cod_proyecto="1"
	end if

	if request.form("opcfechaentrega")>"" then
		opcfechaentrega="1"
	end if

	if request.form("referencia")>"" then
		referencia=limpiaCadena(request.form("referencia"))
	else
		referencia=limpiaCadena(request.querystring("referencia"))
	end if

	if referencia & ""="" then
		if request.form("tdocumento")>"" then
			referencia=limpiaCadena(request.form("tdocumento"))
		else
			referencia=limpiaCadena(request.querystring("tdocumento"))
		end if
	end if

	if request.form("nombreart")>"" then
		nombreart=limpiaCadena(request.form("nombreart"))
	else
		nombreart=limpiaCadena(request.querystring("nombreart"))
	end if

	if request.form("familia")>"" then
		familia=limpiaCadena(request.form("familia"))
	else
		familia=limpiaCadena(request.querystring("familia"))
	end if
	'CheckCadena familia

    ' JFT 16/12/2011 : Select for group by date
    opcagrupar	= limpiaCadena(Request.QueryString("opcagrupar"))
	if opcagrupar ="" then
		opcagrupar	= limpiaCadena(Request.form("opcagrupar"))
	end if
    'END JFT 16/12/2011 : Select for group by date

	' IML 28/04/2004 : Validamos si el usuario tiene acceso
	if viene="tienda" then
		sesionNCliente=left(nproveedor,5)
		if sesionNCliente&""="" then sesionNCliente=session("ncliente")
		checkAccesoTienda sesionNCliente,"",nproveedor
		nproveedor=trimCodEmpresa(nproveedor)
	else
		sesionNCliente=session("ncliente")
	end if
	checkCadenaTienda sesionNCliente,nserie
	checkCadenaTienda sesionNCliente,cod_proyecto
	checkCadenaTienda sesionNCliente,familia
	' FIN IML 28/04/2004 : Validamos si el usuario tiene acceso


	si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)

    WaitBoxOculto LitEsperePorFavor
	Alarma "listado_pedidos_pendientes.asp"

	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstPedido = Server.CreateObject("ADODB.Recordset")
	set rstPedido2 = Server.CreateObject("ADODB.Recordset")
	set rstPedido3 = Server.CreateObject("ADODB.Recordset")

	if mode="traerarticulo" then
		if referencia>"" then
            rstAux.cursorlocation=3
			rstAux.open "select referencia,nombre from articulos with(NOLOCK) where referencia='" & referencia & "'",session("backendlistados")
			if rstAux.EOF then
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgArticuloNoExiste%>");
					//document.listado_pedidos_pendientes.referencia.focus();
					//document.listado_pedidos_pendientes.referencia.select();
				</script><%
				referencia=""
				nombreart=""
			else
				nombreart= enc.EncodeForHtmlAttribute(null_s(rstAux("nombre")))
			end if
			rstaux.close
		else
			nombreart=""
		end if
		mode="select1"
	end if

	if mode="traerproveedor" then
		if nproveedor>"" then
			nproveedor=session("ncliente") & completar(nproveedor,5,"0")
            rstAux.cursorlocation=3
			rstAux.open "select nproveedor,razon_social from proveedores with(NOLOCK) where nproveedor='" & nproveedor & "'",session("backendlistados")
			if rstAux.EOF then
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgProveedorNoExiste%>");
					//document.listado_pedidos_pendientes.nproveedor.focus();
					//document.listado_pedidos_pendientes.nproveedor.select();
				</script><%
				nproveedor=""
				nombre=""
			else
				nombre= enc.EncodeForHtmlAttribute(null_s(rstAux("razon_social")))
			end if
			rstaux.close
		else
			nombre=""
		end if
		mode="select1"
	end if

	strwhere=""
%>
	<%if mode="imp" then%>
		<table width='100%' cellspacing="1" cellpadding="1">
   			<tr>
				<td width="30%" align="left" >
				</td>
				<td class=CELDARIGHT bgcolor="<%=color_blau%>">
					<%if fdesde>"" then
						if fhasta>"" then
							%><%=LitPeriodoFechas%> : <%=enc.EncodeForHtmlAttribute(fdesde)%> - <%=enc.EncodeForHtmlAttribute(fhasta)%><%
						else
							%><%=LitPeriodoFechas%> : <%=LitDesde%>&nbsp;<%=enc.EncodeForHtmlAttribute(fdesde)%><%
						end if
					else
						if fhasta>"" then
							%><%=LitPeriodoFechas%> : <%=LitHasta%>&nbsp;<%=enc.EncodeForHtmlAttribute(fhasta)%><%
						else
						end if
					end if%>
				</td>
	   		</tr>
		</table>
		<%if fdesde>"" or fhasta>"" then%>
			<hr/>
		<%end if
	end if

	if mode="select1" then%>
		<%
            EligeCelda "input", "add", "", "", "", 0, LitDesdeFecha, "fdesde", "", fdesde
            DrawCalendar "fdesde"
			EligeCelda "input", "add", "", "", "", 0, LitHastaFecha, "fhasta", "", fhasta
            DrawCalendar "fhasta"
			
			EligeCelda "input", "add", "", "", "", 0, LitDesdeFechaEntrega, "fentregadesde", "", fentregadesde
            DrawCalendar "fentregadesde"
			EligeCelda "input", "add", "", "", "", 0, LitHastaFechaEntrega, "fentregahasta", "", fentregahasta
            DrawCalendar "fentregahasta"

            rstSelect.cursorlocation=3
    		rstSelect.open "select codigo,descripcion from tipo_actividad with(NOLOCK) where codigo like '" & sesionNCliente & "%' order by descripcion",session("backendlistados")
			DrawSelectCelda "CELDA","175","",0,LitActividad,"actividad",rstSelect,enc.EncodeForHtmlAttribute(null_s(actividad)),"codigo","descripcion","",""
			rstSelect.close
	
            rstSelect.cursorlocation=3                                                                                                                                                                                                                                                                                                                                                                                 
			rstSelect.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series with(NOLOCK) where tipo_documento ='PEDIDO A PROVEEDOR' and nserie like '" & sesionNCliente & "%' order by nserie",session("backendlistados")
			DrawSelectCelda "CELDA","175","",0,LitSerie,"nserie",rstSelect,enc.EncodeForHtmlAttribute(null_s(nserie)),"nserie","descripcion","",""
			rstSelect.close
			strselect = "select razon_social from proveedores with(nolock) where nproveedor=?"
            
			DrawDiv "1", "", ""
               DrawLabel "", "", LitProveedor%><input class='width15' type="text" name="nproveedor" value="<%=enc.EncodeForHtmlAttribute(trimCodEmpresa(nproveedor))%>" size="10" onchange="TraerProveedor('<%=enc.EncodeForJavascript(mode)%>','<%=enc.EncodeForJavascript(ndet)%>');"/><a class='CELDAREFB' href="javascript:AbrirVentana('../proveedores_busqueda.asp?ndoc=listado_pedidos_pendientes&titulo=<%=LitSelProv%>&mode=search&viene=listado_pedidos_pendientes','P',<%=altoventana%>,<%=anchoventana%>)"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' disabled type="text" name="nombre" value="<%=iif(nproveedor>"",DLookupP1(strselect, nproveedor&"",adVarChar,10,session("backendlistados")),"")%>" size="25" /><%
			CloseDiv
			DrawDiv "1", "", ""
               DrawLabel "", "", LitProveedorBaja%><input type="checkbox" name="opcproveedorbaja" <%=iif(opcproveedorbaja>"","checked","")%>/>
			<%
			CloseDiv
            if si_tiene_modulo_proyectos<>0 then
				%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><input class="CELDA" type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(cod_proyecto)%>"/><label><%=LitProyecto%></label><%
				%><iframe class="width60 iframe-menu" id='frProyecto' src='../../mantenimiento/docproyectos_responsive.asp?viene=listado_pedidos_pendientes&mode=<%=enc.EncodeForHtmlAttribute(mode)%>&cod_proyecto=<%=enc.EncodeForHtmlAttribute(cod_proyecto)%>' frameborder="no" scrolling="no" noresize="noresize"></iframe></div><%	
			end if
			DrawDiv "1", "", ""
                DrawLabel "", "", LitPedPendPendConRef%><input type="text" maxlength='25' size="25" name="referencia" value="<%=iif(referencia>"",enc.EncodeForHtmlAttribute(trimCodEmpresa(referencia)),"")%>" size=12 onchange="traerreferencia();"/><%
		    CloseDiv
            EligeCelda "input", "add", "", "", "", 0, LitPedPendPendConNom, "nombreart", "", enc.EncodeForHtmlAttribute(null_s(nombreart))

            rstSelect.cursorlocation=3
    		rstAux.open " select codigo, nombre from familias with(NOLOCK) where codigo like '" & sesionNCliente & "%' order by nombre", session("backendlistados")
	   		DrawSelectCelda "CELDA","175","",0,LitPedPendPendSubFamilia,"familia",rstAux,enc.EncodeForHtmlAttribute(null_s(familia)),"codigo","nombre","",""
			rstAux.close
             
            DrawDiv "1", "", ""
                DrawLabel "", "", LitGroupby%><select class='width60' name="opcagrupar"><%
                
                if opcagrupar="proveedor" or opcagrupar="" then
				    %>
                        <option value="proveedor" selected="selected"><%=LitProveedor%></option>
                        <option value="fechaentrega"><%=LitFechaEntrega %></option>
                    <%
                elseif opcagrupar="fechaentrega" then
                    %>
                        <option value="proveedor" ><%=LitProveedor %></option>
                        <option value="fechaentrega" selected="selected"><%=LitFechaEntrega%></option>
                    <%
                end if
                %></select><%
                
                'END JFT 16/12/2011 : Select group by date
			CloseDiv
			%>
		<hr/>
        <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitCamposOpcionales%></h6>
        <%
			if si_tiene_modulo_proyectos<>0 then
			    EligeCelda "check", "add", "", "", "", 0, LitProyecto, "opc_cod_proyecto", "", iif(opc_cod_proyecto=1,-1,0)
			end if
            EligeCelda "check", "add", "", "", "", 0, LitFechaEntrega, "opcfechaentrega", "", iif(opcfechaentrega=1,-1,0)%>
		<hr/>
		<%
	elseif mode="imp" then
		'MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='049'", DSNIlion)
        strselect= "select maxpagina from limites_listados with(nolock) where item=?"
        MAXPAGINA = DLookupP1(strselect, "049", adVarChar, 3, DSNIlion)    
            
		'MAXPDF=d_lookup("maxpdf", "limites_listados", "item='049'", DSNIlion)    
         strselect= "select maxpdf from limites_listados with(nolock) where item=?"
        MAXPDF= DLookupP1(strselect, "049", adVarChar, 3, DSNIlion)
		%><input type='hidden' name='maxpdf' value='<%=enc.EncodeForHtmlAttribute(MAXPDF)%>'/>
		<input type='hidden' name='maxpagina' value='<%=enc.EncodeForHtmlAttribute(MAXPAGINA)%>'/>
		<input type="hidden" name="fdesde" value="<%=enc.EncodeForHtmlAttribute(fdesde)%>"/>
		<input type="hidden" name="fhasta" value="<%=enc.EncodeForHtmlAttribute(fhasta)%>"/>
		<input type="hidden" name="fentregadesde" value="<%=enc.EncodeForHtmlAttribute(fentregadesde)%>"/>
		<input type="hidden" name="fentregahasta" value="<%=enc.EncodeForHtmlAttribute(fentregahasta)%>"/>
		<input type="hidden" name="nserie" value="<%=enc.EncodeForHtmlAttribute(nserie)%>"/>
		<input type="hidden" name="nproveedor" value="<%=iif(viene="tienda",sesionNCliente&nproveedor,nproveedor)%>"/>
		<input type="hidden" name="opcproveedorbaja" value="<%=enc.EncodeForHtmlAttribute(opcproveedorbaja)%>"/>
		<input type="hidden" name="actividad" value="<%=enc.EncodeForHtmlAttribute(actividad)%>"/>
		<input type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(cod_proyecto)%>"/>
		<input type="hidden" name="opc_cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(opc_cod_proyecto)%>"/>
		<input type="hidden" name="opcfechaentrega" value="<%=enc.EncodeForHtmlAttribute(opcfechaentrega)%>"/>
		<% if viene="tienda" then%>
			<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
			<input type="hidden" name="campo" value="<%=enc.EncodeForHtmlAttribute(campo)%>"/>
			<input type="hidden" name="criterio" value="<%=enc.EncodeForHtmlAttribute(criterio)%>"/>
			<input type="hidden" name="texto" value="<%=enc.EncodeForHtmlAttribute(texto)%>"/>
		<%end if%>
		<%if viene="articulos" then%>
			<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
		<%end if%>
		<input type="hidden" name="referencia" value="<%=enc.EncodeForHtmlAttribute(referencia)%>"/>
		<input type="hidden" name="nombreart" value="<%=enc.EncodeForHtmlAttribute(nombreart)%>"/>
		<input type="hidden" name="familia" value="<%=enc.EncodeForHtmlAttribute(familia)%>"/>
        <input type="hidden" name="opcagrupar" value="<%=enc.EncodeForHtmlAttribute(opcagrupar)%>"/>
		<%
		'ndecimalesEmp=d_lookup("ndecimales", "divisas", "codigo like '" & sesionNCliente & "%' and moneda_base<>0", session("backendlistados"))
        strselect= "select ndecimales from divisas with(nolock) where codigo like ?+'%' and moneda_base<>?"
        ndecimalesEmp= DLookupP2(strselect, sesionNCliente &"",adVarchar,15, 0,adInteger,, session("backendlistados"))

		'abreviaturaEmp=d_lookup("abreviatura", "divisas", "codigo like '" & sesionNCliente & "%' and moneda_base<>0", session("backendlistados"))
        strsel= "select abreviatura from divisas with(nolock) where codigo like ?+'%' and moneda_base<>?"
        abreviaturaEmp= DLookupP2(strselect, sesionNCliente &"",adVarchar,15, 0,adInteger,, session("backendlistados"))
		if viene<>"tienda" then
			VinculosPagina(MostrarProveedores)=1:VinculosPagina(MostrarPedidosPro)=1:VinculosPagina(MostrarArticulos)=1:VinculosPagina(MostrarPedidosCli)=1
			CargarRestricciones session("usuario"),sesionNCliente,Permisos,Enlaces,VinculosPagina
		end if

		total_valor_general = 0
		total_pendiente_general = 0
		encabezado=0
		'strwhere="where"

		strwhere=CadenaBusqueda(campo,criterio,texto)

		if nproveedor > "" then
			nproveedor=sesionNCliente & completar(nproveedor,5,"0")
			strwhere=strwhere & " pedidos_pro.nproveedor='" & nproveedor & "' and"
			if viene<>"tienda" then
            'd_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("backendlistados")
            strselect="select razon_social from proveedores with(nolock) where nproveedor =?"
				%><font class='CELDA'><b><%=LitProveedor%>: </b><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(nproveedor)) & " - " & DLookupP1(strselect, nproveedor&"", adVarChar,10 ,session("backendlistados"))%></font><br/><%
				encabezado=1
			end if
		else
			if opcproveedorbaja="" then
				strbaja=" "
			else
				strbaja=" proveedores.fbaja is null and"
				strwhere=strwhere & strbaja
				%><font class='CELDA'><b><%=LitProveedorBaja%></b></font><br/><%
				encabezado=1
			end if
		end if
		if nserie > "" then
            'd_lookup("nombre","series","nserie='" & nserie & "'",session("backendlistados"))
            strselect= "select nombre from series with(nolock) where nserie=?"
			strwhere=strwhere & " serie='" & nserie & "' and"
			%><font class='CELDA'><b><%=LitSerie%>:&nbsp;</b></font><font class='CELDA'><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(nserie))%>&nbsp;<%=DLookupP1(strselect, nserie &"", adVarchar,10,session("backendlistados"))%></font><br/><%
			encabezado=1
		end if
		if fdesde > "" then strwhere=strwhere & " fecha>='" & fdesde & "' and"
		if fhasta > "" then strwhere=strwhere & " fecha<='" & fhasta & "' and"

		if fentregadesde > "" then
			strwhere=strwhere & " fecha_entrega>='" & fentregadesde & "' and"
			%><font class='CELDA'><b><%=LitDesdeFechaEntrega%>:&nbsp;</b></font><font class='CELDA'><%=enc.EncodeForHtmlAttribute(fentregadesde)%></font><br/><%
			encabezado=1
		end if
		if fentregahasta > "" then
			strwhere=strwhere & " fecha_entrega<='" & fentregahasta & "' and"
			%><font class='CELDA'><b><%=LitHastaFechaEntrega%>:&nbsp;</b></font><font class='CELDA'><%=enc.EncodeForHtmlAttribute(fentregahasta)%></font><br/><%
			encabezado=1
		end if

		if actividad > "" then
                'd_lookup("descripcion","tipo_actividad","codigo='" & actividad & "'",session("backendlistados"))
                strselect= "select descripcion from tipo_actividad with(nolock) where codigo=?"
			strwhere=strwhere & " tactividad='" & actividad & "' and"
			%><font class='CELDA'><b><%=LitActividad%>:&nbsp;</b></font><font class='CELDA'><%=DLookupP1(strselect, actividad &"", adVarChar, 10,session("backendlistados")) %></font><br/><%
			encabezado=1
		end if
		if cod_proyecto>"" then
            'd_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",session("backendlistados"))
            strselect= "select nombre from proyectos with(nolock) where codigo=?"
			strwhere=strwhere & " cod_proyecto='" & cod_proyecto & "' and"
			%><font class='CELDA'><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=DLookupP1(strselect, cod_proyecto &"", adVarChar,15,session("backendlistados")) %></font><br/><%
			encabezado=1
		end if

		'Tenemos una clausula where para los detalles y otra para los conceptos
		strwhereConceptos=strwhere

		if referencia>"" then
			if viene<>"articulos" then
				referencia=sesionNCliente & referencia
				strwhere=strwhere & " detalles_ped_pro.referencia like '%" & referencia & "%' and"
				strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"
				%><font class='CELDA'><b><%=LitPedPendPendConRef%> : </b></font><font class="CELDA"><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(referencia))%></font><br/><%
			else
				strwhere=strwhere & " detalles_ped_pro.referencia='" & referencia & "' and"
				strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"
                %><font class='CELDA'><b><%=LitReferencia%>:&nbsp;</b></font><font class='CELDA'><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(referencia))%></font><br/><%
			end if

			encabezado=1
		end if
		'**RGU 24/4/2006
		strwhere=strwhere & " detalles_ped_pro.cantidadpend <> 0 and"
		'**RGU
		if nombreart>"" and referencia & ""="" then
			strwhere=strwhere & " articulos.referencia=detalles_ped_pro.referencia and articulos.nombre like '%" & nombreart & "%' and"
			strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"
			%><font class='CELDA'><b><%=LitPedPendPendConNom%> : </b></font><%=enc.EncodeForHtmlAttribute(nombreart)%><br/><%
			encabezado=1
		end if
		if familia>"" then
            'd_lookup("nombre","familias","codigo='" & familia & "'",session("backendlistados"))
            strselect= "select nombre from familias with(nolock) where codigo =?"
			strwhere=strwhere & " articulos.referencia=detalles_ped_pro.referencia and articulos.familia='" & familia & "' and"
			strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"
			%><font class='CELDA'><b><%=LitPedPendPendSubFamilia%> : </b></font><%=DLookupP1(strselect, familia &"", adVarChar,10,session("backendlistados"))%><br/><%
			encabezado=1
		end if
        
		strwhere=strwhere + " pedidos_pro.nproveedor=proveedores.nproveedor and nfactura is null and nalbaran is null"
		strwhere=strwhere + " and pedidos_pro.npedido=detalles_ped_pro.npedido"
		strwhere=strwhere + " and pedidos_pro.npedido like '" & sesionNCliente & "%' "
		strwhere=strwhere + " and detalles_ped_pro.npedido like '" & sesionNCliente & "%' "
		strwhere=strwhere + " and divisas.codigo like '" & sesionNCliente & "%' "
		strwhere=strwhere + " and proveedores.nproveedor like '" & sesionNCliente & "%' "
		strwhere=strwhere + " and divisas.codigo=pedidos_pro.divisa "
		

		strwhereConceptos=strwhereConceptos + " pedidos_pro.nproveedor=proveedores.nproveedor and nfactura is null and nalbaran is null"
		strwhereConceptos=strwhereConceptos + " and pedidos_pro.npedido=conceptos_ped_pro.npedido"
		strwhereConceptos=strwhereConceptos + " and pedidos_pro.npedido like '" & sesionNCliente & "%' "
		strwhereConceptos=strwhereConceptos + " and conceptos_ped_pro.npedido like '" & sesionNCliente & "%' "
		strwhereConceptos=strwhereConceptos + " and conceptos_ped_pro.descripcion not like 'Anulación Cantidades%' "
		strwhereConceptos=strwhereConceptos + " and divisas.codigo like '" & sesionNCliente & "%' "
		strwhereConceptos=strwhereConceptos + " and proveedores.nproveedor like '" & sesionNCliente & "%' "
		strwhereConceptos=strwhereConceptos + " and divisas.codigo=pedidos_pro.divisa "

	strwhereAntes=strwhere
	strwhereConceptosAntes=strwhereConceptos 
	'**rgu 20/9/07 **
	strwhereConceptosAntes=strwhereConceptosAntes + " and s.nserie like '" & sesionNCliente & "%' "
	strwhereConceptosAntes=strwhereConceptosAntes + " and s.nserie =pedidos_pro.serie "
	'**rgu**
    'JFT 16/12/2011 : Select group by date
	if opcagrupar="proveedor" or opcagrupar&""="" then
        strwhereConceptosAntes=strwhereConceptosAntes + " order by pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"
        strwhere2=strwhere + " and mainitem is null GROUP BY pedidos_pro.npedido, pedidos_pro.nproveedor, fecha, pedidos_pro.fecha_entrega"
	    strwhereConceptos=strwhereConceptos + " GROUP BY pedidos_pro.npedido, pedidos_pro.nproveedor, fecha, pedidos_pro.fecha_entrega"
	    strwhereConceptos=strwhereConceptos + " ORDER BY pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"

        %><font class='CELDA'><b><%=LitGroupby + " " + LitProveedor%></b></font><br/><%
        encabezado=1
    elseif opcagrupar="fechaentrega" then
		strwhereConceptosAntes=strwhereConceptosAntes + " order by fecha_entrega asc, pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"
        strwhere2=strwhere + " and mainitem is null GROUP BY pedidos_pro.npedido, pedidos_pro.nproveedor, fecha, pedidos_pro.fecha_entrega"
	    strwhereConceptos=strwhereConceptos + " GROUP BY pedidos_pro.npedido, pedidos_pro.nproveedor, fecha, pedidos_pro.fecha_entrega"
	    strwhereConceptos=strwhereConceptos + " ORDER BY pedidos_pro.fecha_entrega, pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"
        
		%><font class='CELDA'><b><%=LitGroupby + " " + LitFechaEntrega%></b></font><br/><%
		encabezado=1
	end if
    
	''strwhere=strwhere + " order by pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"
    'END JFT 16/12/2011 : Select group by date

'Calculos de páginas--------------------------
    'JFT 16/12/2011 : Select group by date
		strselect="select count(pedidos_pro.npedido) as contador,pedidos_pro.npedido,pedidos_pro.nproveedor,fecha,'detalles',fecha_entrega "
		strselect=strselect & " from pedidos_pro with (NOLOCK),divisas with (NOLOCK),proveedores with (NOLOCK),detalles_ped_pro with (NOLOCK) "

		strselectConceptos="select count(pedidos_pro.npedido) as contador,pedidos_pro.npedido,pedidos_pro.nproveedor,fecha,'conceptos',fecha_entrega "
		strselectConceptos=strselectConceptos & " from pedidos_pro with (NOLOCK),divisas with (NOLOCK),proveedores with (NOLOCK),conceptos_ped_pro with (NOLOCK) "
    'END JFT 16/12/2011 : Select group by date
		if (nombreart>"" and referencia & ""="") or familia>"" then
			strselect=strselect & ",articulos with (NOLOCK) "
		end if
		''strselect=strselect & strwhere2
		strselect=strselect & strwhere2 & " union " & strselectConceptos & strwhereConceptos
		rstPedido2.cursorlocation=3
		rstPedido2.Open strselect, session("backendlistados")

        if not rstPedido2.eof then
			suma=0

			while not rstPedido2.eof
				suma=suma+2+ enc.EncodeForHtmlAttribute(null_s(rstPedido2("contador")))
				rstPedido2.movenext
			wend
            
			cantidad=suma/MAXPAGINA
			if cantidad>clng(cantidad) then
			   cantidad=clng(cantidad)+1
			else
			  cantidad=clng(cantidad)
			end if

			dim lista
			redim lista(cantidad+1,4)
			rstPedido2.movefirst
			i=1
			lista(1,1)=0
			lista(1,2)=0
			if not rstPedido2.eof then
				lista(1,3)= enc.EncodeForHtmlAttribute(null_s(rstPedido2("npedido")))
			else
				lista(1,3)=""
			end if
			sumat=0
			suma2=0
			j=0
            
            pedidoanterior=""
			while not rstPedido2.eof
                
                if pedidoanterior <> rstPedido2("npedido") then
				    suma2=suma2+2+ enc.EncodeForHtmlAttribute(null_s(rstPedido2("contador")))
                else
                    suma2=suma2+ enc.EncodeForHtmlAttribute(null_s(rstPedido2("contador")))
                end if
                pedidoanterior = enc.EncodeForHtmlAttribute(null_s(rstPedido2("npedido")))
				lista(i,2)=suma2
				lista(i,1)=lista(i,1)+ enc.EncodeForHtmlAttribute(null_s((rstPedido2("contador"))))

				j=j+1
                
				if suma2>=MAXPAGINA then
					lista(i,4)= enc.EncodeForHtmlAttribute(null_s(rstPedido2("npedido")))
					i=i+1
					j=0
					lista(i,1)=0
					lista(i,2)=0
					sumat=sumat+suma2
					suma2=0
					rstPedido2.movenext
					if not rstPedido2.eof then
                        if not rstPedido2.eof then
                            pedidoanteriorOld = enc.EncodeForHtmlAttribute(null_s(rstPedido2("npedido")))&""
                        end if
                        while pedidoanterior = pedidoanteriorOld and not rstPedido2.eof
                            rstPedido2.movenext
                            if not rstPedido2.eof then
                                pedidoanteriorOld = enc.EncodeForHtmlAttribute(null_s(rstPedido2("npedido")))&""
                            end if
                        wend
                        if not rstPedido2.eof then
						    lista(i,3)=enc.EncodeForHtmlAttribute(null_s(rstPedido2("npedido")))
                        else
                            lista(i,3)=pedidoanteriorOld
                        end if
					else
						rstPedido2.moveprevious
						lista(i,3)= enc.EncodeForHtmlAttribute(null_s(rstPedido2("npedido")))
						rstPedido2.movenext
					end if
				else
					rstPedido2.movenext
				end if
			wend
			rstPedido2.close
		else
			suma=0
			'lista(1)=0
		end if

		sumat=sumat + suma2
        
        mm=0

		strselect="select item,detalles_ped_pro.descripcion,fecha,pedidos_pro.nproveedor,pedidos_pro.npedido,pedidos_pro.divisa,pedidos_pro.fecha_entrega,cantidad,detalles_ped_pro.referencia,almacenes.descripcion as almacen,detalles_ped_pro.importe,cod_proyecto,'detalle' as tipo "
		strselect=strselect & ",detalles_ped_pro.cantidadpend, detalles_ped_pro.npedidocli,detalles_ped_pro.itempedidocli "
		strselect=strselect & ",pedidos_pro.total_pedido,proyectos.nombre as nom_proyecto,proveedores.razon_social,divisas.abreviatura "

        ''ricardo 25-7-2006 que salga el importe pendiente de cada detalle
        strselect=strselect & ",(round((detalles_ped_pro.cantidadpend*detalles_ped_pro.pvp),divisas.ndecimales)-"
        strselect=strselect & "round((((round((detalles_ped_pro.cantidadpend*detalles_ped_pro.pvp),divisas.ndecimales))*detalles_ped_pro.descuento)/100),divisas.ndecimales))-"
        strselect=strselect & "round(((((round((detalles_ped_pro.cantidadpend*detalles_ped_pro.pvp),divisas.ndecimales)-"
        strselect=strselect & "round((((round((detalles_ped_pro.cantidadpend*detalles_ped_pro.pvp),divisas.ndecimales))*detalles_ped_pro.descuento)/100),divisas.ndecimales)))*detalles_ped_pro.descuento2)/100),divisas.ndecimales)"

		strselect=strselect & " as Importepend "
		
		

		strselect=strselect & ",null "

		strselect=strselect & "as sumimportepend "
		
		'**rgu 20/9/07:añadimos el codigo i descripcion de la serie al listado
		strselect=strselect & ", right(s.nserie,len(s.nserie)-5)+'-'+s.nombre as NSerie "
		'**rgu**


		strselect=strselect & " from pedidos_pro with (NOLOCK) "
		strselect=strselect & " left outer join proyectos on proyectos.codigo=pedidos_pro.cod_proyecto "
		strselect=strselect & ",proveedores with (NOLOCK),divisas with (NOLOCK),detalles_ped_pro with (NOLOCK) "
		strselect=strselect & "left outer join almacenes with (NOLOCK) on almacenes.codigo=detalles_ped_pro.almacen "
		'**rgu 20/9/07:
		strselect=strselect & ", series s with(nolock) "
		'**rgu**
		

		strselectConceptos="select nconcepto,convert(varchar(8000),conceptos_ped_pro.descripcion),fecha,pedidos_pro.nproveedor,pedidos_pro.npedido,pedidos_pro.divisa,pedidos_pro.fecha_entrega,conceptos_ped_pro.cantidad,'' as referencia,'' as almacen,conceptos_ped_pro.importe,cod_proyecto,'concepto' as tipo"
		strselectConceptos=strselectConceptos & ",0 as cantidadpend,null as npedidocli,null as itempedidocli"
		strselectConceptos=strselectConceptos & ",pedidos_pro.total_pedido,proyectos.nombre as nom_proyecto,proveedores.razon_social,divisas.abreviatura "
		strselectConceptos=strselectConceptos & " ,0 as Importepend "
		strselectConceptos=strselectConceptos & ",null as sumimportepend "
		'**rgu 20/9/07
		strselectConceptos=strselectConceptos & ", right(s.nserie,len(s.nserie)-5)+'-'+s.nombre as NSerie "
		'**rgu**

		strselectConceptos=strselectConceptos & " from pedidos_pro with (NOLOCK)"
		strselectConceptos=strselectConceptos & " left outer join proyectos on proyectos.codigo=pedidos_pro.cod_proyecto "
		strselectConceptos=strselectConceptos & ",proveedores with (NOLOCK),divisas with (NOLOCK),conceptos_ped_pro with (NOLOCK) "
		'**rgu 20/9/07
		strselectConceptos=strselectConceptos & ", series s with(nolock) "
		'**rgu**
		if (nombreart>"" and referencia & ""="") or familia>"" then
			strselect=strselect & ",articulos "
		end if
		''strselect=strselect & strwhere
		'**RGU 24/4/2006
		''strwherecero=" and detalles_ped_pro.cantidadpend <> 0 "
		'**RGU
		
		'**rgu 20/9/07 **
		strwhereAntes=strwhereAntes + " and s.nserie like '" & sesionNCliente & "%' "
		strwhereAntes=strwhereAntes + " and s.nserie =pedidos_pro.serie "
		'**rgu**
		
		strselect=strselect & strwhereAntes & strwherecero & " and mainitem is null union " & strselectConceptos & strwhereConceptosAntes & ",detalles_ped_pro.referencia desc"

	DropTable session("usuario"), session("backendlistados")
	crear="CREATE TABLE [" & session("usuario") & "] (num int identity(1,1)"
	crear=crear & ",item smallint,descripcion varchar(8000),fecha smalldatetime,nproveedor char(10),npedido varchar(20),divisa varchar(10)"
	crear=crear & ",fecha_entrega smalldatetime,cantidad real"
	crear=crear & ",referencia varchar(30),almacen varchar(100),importe money,cod_proyecto varchar(30),tipo varchar(50)"
''	crear=crear & ",stock real,precibir real,salida smalldatetime"
	crear=crear & ",cantidadpend real,npedidocli varchar(20),itempedidocli smallint,total_pedido money,nom_proyecto varchar(100),razon_social varchar(100)"
	crear=crear & ",abreviatura varchar(10),Importepend money,sumimportepend money"
	'**rgu 20/9/07
	crear=crear & ", descserie varchar(55) )"
	'**rgu**
	rstPedido.open crear,session("backendlistados"),adUseClient,adLockReadOnly
	GrantUser session("usuario"), session("backendlistados")
    	
	strinsertar="insert into [" & session("usuario") & "](item,descripcion,fecha,nproveedor,npedido,divisa,fecha_entrega,cantidad,referencia,almacen,importe,cod_proyecto,tipo"
    ''strinsertar=strinsertar & ",stock,precibir,salida
    strinsertar=strinsertar & ",cantidadpend,npedidocli,itempedidocli,total_pedido,nom_proyecto,razon_social,abreviatura,Importepend,sumimportepend,descserie) " & strselect
	rstPedido.open strinsertar,session("backendlistados"),adUseClient,adLockReadOnly

	strselect= "select * from [" & session("usuario") & "] order by num"
   

		rstPedido.cursorlocation=3
		rstPedido.Open strselect,session("backendlistados")
''response.write("el strselect es-" & strselect & "-<br>")
''response.end
		if rstPedido.EOF then
			rstPedido.Close
			%><input type="hidden" name="NumRegsTotal" value="0"/>
            <div class="CEROFILAS"><%=LitMsgDatosNoExiste%></div>
            <%			
		else
			%><input type="hidden" name="NumRegsTotal" value="<%=rstPedido.RecordCount%>"/><%

			'Calculos de páginas--------------------------
		   lote=limpiaCadena(Request.QueryString("lote"))
			if lote="" then
				lote=1
			end if
			sentido=limpiaCadena(Request.QueryString("sentido"))
			lotes=sumat/MAXPAGINA
			if lotes>clng(lotes) then
				lotes=clng(lotes)+1
			else
				lotes=clng(lotes)
			end if

			if lotes>(i-1) then
				if j<MAXPAGINA and j>0 then
					lotes=i
				else
					lotes=i-1
				end if
			else
				'lotes=lotes
			end if

			if sentido="next" then
				lote=lote+1
			elseif sentido="prev" then
				lote=lote-1
			end if
			loteaux=lote
			while loteaux>0
				tamano=lista(loteaux,1)
				loteaux=loteaux-1
			wend
			if tamano>clng(tamano) then
				tamano=clng(tamano)+1
			else
				tamano=clng(tamano)
			end if
			if lote=lotes then
				tamano=tamano-1
			end if
			if not rstPedido.eof then
				continuar=1
				while continuar=1

					if rstPedido("npedido")<>lista(lote,3) then
						rstPedido.movenext
						if not rstPedido.eof then
							continuar=1
						else
							continuar=0
						end if
					else
						continuar=0
					end if
				wend
			end if
			'-----------------------------------------

			if lotes>1 and encabezado=1 then
				%><hr/><%
			end if

			NavPaginas lote,lotes,campo,criterio,texto,1
''response.write("voy 2")
''response.end

			%>
			<table width='100%' border='0' style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%

			'Fila de encabezado
			nproveedor_ant="#@#"
			npedido_ant = "#@#"
			filab=0
			filac=1

			while not rstPedido.EOF and filab<lista(lote,2) 'MAXPAGINA
				'CheckCadena rstPedido("nproveedor")
				CheckCadenaTienda sesionNCliente, enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor")))
				if rstPedido("nproveedor")<>nproveedor_ant and rstPedido("npedido")<>npedido_ant then
					DrawFila color_blau
						DrawCeldaSpan "ENCABEZADOL","","",0,"<hr/>",11
					CloseFila
					DrawFila color_blau
						if viene<>"tienda" then
							dat1=Hiperv(OBJProveedores,enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor"))),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor")))) & " - " & enc.EncodeForHtmlAttribute(null_s(rstPedido("razon_social"))),LitVerProveedor)
						else
							dat1=trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor")))) & " - " & enc.EncodeForHtmlAttribute(null_s(rstPedido("razon_social")))
						end if
						DrawCeldaSpan "ENCABEZADOL","","",0,LitProveedor & " : " & dat1,8
					CloseFila
					nproveedor_ant = rstPedido("nproveedor")
				end if
				if rstPedido("npedido")<>npedido_ant then
					DrawFila color_terra
						DrawCelda "ENCABEZADOL7","5%","",0," "
						if viene<>"tienda" then
							dat1=Hiperv(OBJPedidosPro,rstPedido("npedido"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("npedido")))),LitVerPedido)
						else
							dat1="<a class='CELDAREFB' href=javascript:ver_pedido('" & enc.EncodeForHtmlAttribute(null_s(rstPedido("npedido"))) & "','"&nproveedor&"') alt='" & LitVerPedido & "'>" & trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("npedido")))) & "</a>"
						end if
						dat2=" - "&LitSerie&": "& enc.EncodeForHtmlAttribute(null_s(rstPedido("descserie")))
						dat2=dat2&" - " + LitFecha+": "+cstr(enc.EncodeForHtmlAttribute(null_s(rstPedido("fecha"))))+"-"+LitImporteListPendSer+": "+formatnumber(enc.EncodeForHtmlAttribute(null_s(rstPedido("total_pedido"))),ndecimalesEmp,-1,0,-1)+" "+ enc.EncodeForHtmlAttribute(null_s(rstPedido("abreviatura")))
						if opc_cod_proyecto="1" then
							dat2=dat2 & " - " & LitProyecto & " : " & enc.EncodeForHtmlAttribute(null_s(rstPedido("nom_proyecto")))
						end if
						if opcfechaentrega="1" then
							DrawCeldaSpan "ENCABEZADOL","","",0,LitPedidoMin + " : " & dat1 & dat2,10
						else
							DrawCeldaSpan "ENCABEZADOL","","",0,LitPedidoMin + " : " & dat1 & dat2,9
						end if
					CloseFila
					npedido_ant = enc.EncodeForHtmlAttribute(null_s(rstPedido("npedido")))
					DrawFila color_terra
						DrawCeldaSpan "ENCABEZADOL","","",0,"",2
						'DrawCelda "ENCABEZADOL7","5%","",0," "
						'DrawCelda "ENCABEZADOL7","5%","",0," "
						DrawCelda "ENCABEZADOR7","","",0,LitCantidad
						DrawCelda "ENCABEZADOL7","","",0,LitReferencia
						DrawCelda "ENCABEZADOL7","","",0,LitDescripcionConcepto
						if opcfechaentrega="1" then
							DrawCelda "ENCABEZADOL7","","",0,LitFechaEntrega
						end if
						DrawCelda "ENCABEZADOL7","","",0,LitAlmacen
						DrawCelda "ENCABEZADOL7","","",0,LitPendiente
						DrawCelda "ENCABEZADOL7","","",0,LitPedidoCli
						DrawCelda "ENCABEZADOL7","","",0,LitItem
''ricardo 25-7-2006 ya que tenemos tabla temporal que se calcule de ella
sumimportepend=d_sum("Importepend","[" & session("usuario") & "]","npedido='" & enc.EncodeForHtmlAttribute(null_s(rstPedido("npedido"))) & "'",session("backendlistados"))
						DrawCelda "ENCABEZADOR7","","",0,LitImportePendiente & ": " & formatnumber(sumimportepend,ndecimalesEmp,-1,0,-1) & " " & enc.EncodeForHtmlAttribute(null_s(rstPedido("abreviatura")))
					CloseFila
					filab=filab+2
				end if
				DrawFila color_blau
					DrawCelda "ENCABEZADOL7","5%","",0," "
					DrawCelda "ENCABEZADOL7","5%","",0," "
					DrawCelda "CELDAR7","","",0, enc.EncodeForHtmlAttribute(null_s(rstPedido("cantidad")))
					if viene<>"tienda" then
						dat1=Hiperv(OBJArticulos,enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia"))),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia")))),LitVerArticulo)
						dat3=Hiperv(OBJPedidosCli, enc.EncodeForHtmlAttribute(null_s(rstPedido("npedidocli"))),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("npedidocli")))),LitVerPedido)
					else
						dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia"))) & "','" & enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor"))) & "') alt='" & LitVerArticulo & "' border='0'>" & trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia")))) & "</a>"
						dat3="<a class='CELDAREFB' href=javascript:ver_pedidoCli('" & enc.EncodeForHtmlAttribute(null_s(rstPedido("npedidocli"))) & "') alt='" & LitVerPedido & "'>" & trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("npedidocli")))) & "</a>"
					end if
					DrawCelda "CELDAL7","","",0,dat1
					'if viene<>"tienda" then
					'	dat1=Hiperv(OBJArticulos,rstPedido("referencia"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,rstPedido("descripcion"),LitVerArticulo)
					'else
					'	dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & rstPedido("referencia") & "','" & rstPedido("nproveedor") & "') alt='" & LitVerArticulo & "' border='0'>" & rstPedido("descripcion") & "</a>"
					'end if
					dat1= enc.EncodeForHtmlAttribute(null_s(rstPedido("descripcion")))
					DrawCelda "CELDAL7","","",0,dat1
					if opcfechaentrega="1" then
						DrawCelda "CELDAL7","","",0, enc.EncodeForHtmlAttribute(null_s(rstPedido("fecha_entrega")))
					end if
					DrawCelda "CELDAL7","","",0, enc.EncodeForHtmlAttribute(null_s(rstPedido("almacen")))
					DrawCelda "CELDAL7","","",0,iif(rstPedido("tipo")="detalle",null_z(rstPedido("cantidadpend")),"")
					'DrawCelda "CELDAL7","","",0,rstPedido("npedidocli")
					DrawCelda "CELDAL7","","",0,dat3
					DrawCelda "CELDAL7","","",0,iif(rstPedido("npedidocli")>"",rstPedido("itempedidocli"),"")
					DrawCelda "CELDAR7","","",0,formatnumber(null_z(rstPedido("Importepend")),ndecimalesEmp,-1,0,-1)+" "+rstPedido("abreviatura")
				CloseFila
				filab=filab+1
				rstPedido.MoveNext
				filac=filac+1
                response.Flush
			wend
''response.write("voy 3")
''response.end

			' si no se ha acabado el pedido, se escribira hasta que se acabe
			filab=0
			if not rstPedido.eof then
				continuar=1
				while continuar=1
					if rstPedido("npedido")=npedido_ant and rstPedido("nproveedor")=nproveedor_ant then
						DrawFila color_blau
							DrawCelda "ENCABEZADOL7","5%","",0," "
							DrawCelda "ENCABEZADOL7","5%","",0," "
							DrawCelda "CELDAR7","","",0,rstPedido("cantidad")
							if viene<>"tienda" then
								dat1=Hiperv(OBJArticulos,enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia"))),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia")))),LitVerArticulo)
							else
								dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia"))) & "','" & enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor"))) & "') alt='" & LitVerArticulo & "' border='0'>" & enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia"))) & "</a>"
							end if
							DrawCelda "CELDAL7","","",0,dat1
							'if viene<>"tienda" then
							'	dat1=Hiperv(OBJArticulos,rstPedido("referencia"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,rstPedido("descripcion"),LitVerArticulo)
							'else
							'	dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & rstPedido("referencia") & "','" & rstPedido("nproveedor") & "') alt='" & LitVerArticulo & "' border='0'>" & rstPedido("descripcion") & "</a>"
							'end if
							dat1=rstPedido("descripcion")
							DrawCelda "CELDAL7","","",0,dat1
							if opcfechaentrega="1" then
								DrawCelda "CELDAL7","","",0,enc.EncodeForHtmlAttribute(null_s(rstPedido("fecha_entrega")))
							end if
							DrawCelda "CELDAL7","","",0,enc.EncodeForHtmlAttribute(null_s(rstPedido("almacen")))
							DrawCelda "CELDAL7","","",0,iif(rstPedido("tipo")="detalle",null_z(rstPedido("cantidadpend")),"")
							DrawCelda "CELDAL7","","",0,enc.EncodeForHtmlAttribute(null_s(rstPedido("npedidocli")))
							DrawCelda "CELDAL7","","",0,iif(rstPedido("npedidocli")>"",rstPedido("itempedidocli"),"")
							DrawCelda "CELDAR7","","",0,formatnumber(enc.EncodeForHtmlAttribute(null_s(rstPedido("Importepend"))),ndecimalesEmp,-1,0,-1)+" "+ enc.EncodeForHtmlAttribute(null_s(rstPedido("abreviatura")))
						CloseFila
						filab=filab+1
						rstPedido.movenext
						if not rstPedido.eof then
							if rstPedido("npedido")=npedido_ant then
								continuar=1
							else
								continuar=0
							end if
						else
							continuar=0
						end if
					else
						continuar=0
					end if
				wend
				%><input type="hidden" name="npedido_ant" value="<%=enc.EncodeForHtmlAttribute(npedido_ant)%>"/><%
				%><input type="hidden" name="nproveedor_ant2" value="<%=enc.EncodeForHtmlAttribute(nproveedor_ant)%>"/><%
			end if

if lote=lotes then
''ricardo 25-7-2006 ya que tenemos tabla temporal que se calcule de ella
sumimportependtotal=d_sum("Importepend","[" & session("usuario") & "]","",session("backendlistados"))
		DrawFila color_fondo
            'JFT 19/12/2011 colspan last line
            if opcfechaentrega="1" then
		        DrawCelda "CELDAR7 colspan='11'","","",0,"<b>" & LitTotListPendSer & ": " & formatnumber(sumimportependtotal,ndecimalesEmp,-1,0,-1) & " " & abreviaturaEmp & "</b>"
	        else
		        DrawCelda "CELDAR7 colspan='10'","","",0,"<b>" & LitTotListPendSer & ": " & formatnumber(sumimportependtotal,ndecimalesEmp,-1,0,-1) & " " & abreviaturaEmp & "</b>"
	        end if
			'END JFT 19/12/2011 colspan last line
		CloseFila

end if

			%></table><%

			%><hr/><%

			rstPedido.Close

			NavPaginas lote,lotes,campo,criterio,texto,2
		end if
		%>

	<%
	end if%>
</form><%

end if
set rstSelect=nothing
set rstAux=nothing
set rst=nothing
set rstPedido=nothing
set rstPedido2=nothing
set rstPedido3=nothing
%>
</body>
</html>
