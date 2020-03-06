<%@ Language=VBScript %>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function%>
<%
  '***RGU 20/12/2005: Añadir al listado columnas de "cantidad pendiente", "Pedido Cliente", "Item"(cliente)
  response.buffer=true
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=titulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<META HTTP-EQUIV="Content-style-TypeCONTENT="text/css">
<LINK REL="styleSHEET" href="../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../impresora.css" MEDIA="PRINT">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<% if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="Pedido_tiendas.inc" -->

<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

<%if request.querystring("viene")="tienda" or request.form("viene")="tienda" then
	titulo=LitTituloListadoPen2
else
	titulo=LitTituloListadoPen
end if%>

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

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
function TraerCliente(mode) {
	document.anular_cantidades_ptes.action="anular_cantidades_ptesp.asp?mode=traercliente"
	document.anular_cantidades_ptes.submit();
}

function Ver_Articulo(referencia,almorigen)
{
	pagina="../tiendas/pedido_ficha_articulo.asp?referencia=" + referencia + "&viene=anular_cantidades_ptes&almorigen=" + almorigen;
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
				CadenaBusqueda=" where pedidos_tienda." + campo + " like '%" + texto + "%' and"
			case "empieza"
				CadenaBusqueda=" where pedidos_tienda." + campo + " like '" + texto + "%' and"
			case "termina"
				CadenaBusqueda=" where pedidos_tienda." + campo + " like '%" + texto + "' and"
			case "igual"
				CadenaBusqueda=" where pedidos_tienda." + campo + "='" + texto + "' and"
		end select
	else
		CadenaBusqueda=" where "
	end if
end function
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
const borde=0

 %>
<form name="anular_cantidades_ptes" method="post">
<% PintarCabecera "anular_cantidades_ptesp.asp"
'Leer parámetros de la página


	mode		= enc.EncodeForJavascript(Request.QueryString("mode"))
	if mode="browse" then mode="imp"
	almorigen	= limpiaCadena(Request.QueryString("almorigen"))
	if almorigen ="" then
		almorigen	= limpiaCadena(Request.form("almorigen"))
	end if

	if almorigen="" then
		almorigen	= limpiaCadena(Request.QueryString("ndoc"))
		if almorigen ="" then
			almorigen	= limpiaCadena(Request.form("ndoc"))
		end if
	end if

	'if ncliente > "" then
	'	ncliente = session("ncliente") & completar(ncliente,5,"0")
	'end if

	almorigen	= limpiaCadena(Request.QueryString("almorigen"))
	if almorigen ="" then
		almorigen	= limpiaCadena(Request.form("almorigen"))
	end if
	almdestino		= limpiaCadena(Request.QueryString("almdestino"))
	if almdestino="" then
		almdestino	= limpiaCadena(Request.form("almdestino"))
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

	nserie	= limpiaCadena(Request.QueryString("nserie"))
	if nserie ="" then
		nserie	= limpiaCadena(Request.form("nserie"))
	end if
	'CheckCadena nserie

	'cag
	if enc.EncodeForJavascript(request.form("opcElimPedSinDet"))>"" then
		opcElimPedSinDet=limpiaCadena(request.form("opcElimPedSinDet"))
	else
		opcElimPedSinDet=limpiaCadena(request.querystring("opcElimPedSinDet"))
	end if
	if opcElimPedSinDet>"" then
	   opcEliminar=1
	else
	    opcEliminar=0
	end if

	if enc.EncodeForJavascript(request.form("h_nregistros"))>"" then
		h_nregistros=limpiaCadena(request.form("h_nregistros"))
	else
		h_nregistros=limpiaCadena(request.querystring("h_nregistros"))
	end if
   'fin cag
   
	if enc.EncodeForJavascript(request.form("nDias"))>"" then
		nDias=limpiaCadena(request.form("nDias"))
	else
		nDias=limpiaCadena(request.querystring("nDias"))
	end if

	if enc.EncodeForJavascript(request.form("npedido_ant"))>"" then
		npedido2_ant=limpiaCadena(request.form("npedido_ant"))
	else
		npedido2_ant=limpiaCadena(request.querystring("npedido_ant"))
	end if

	if enc.EncodeForJavascript(request.form("ncliente_ant2"))>"" then
		ncliente_ant2=limpiaCadena(request.form("ncliente_ant2"))
	else
		ncliente_ant2=limpiaCadena(request.querystring("ncliente_ant2"))
	end if

	cod_proyecto	= limpiaCadena(Request.QueryString("cod_proyecto"))
	if cod_proyecto="" then
		cod_proyecto	= limpiaCadena(Request.form("cod_proyecto"))
	end if
	'CheckCadena cod_proyecto

	if enc.EncodeForJavascript(request.form("viene"))>"" then
		viene=limpiaCadena(request.form("viene"))
	else
		viene=limpiaCadena(request.querystring("viene"))
	end if

	campo=limpiaCadena(Request.querystring("campo"))
	criterio=limpiaCadena(Request.querystring("criterio"))
	texto=limpiaCadena(Request.querystring("texto"))

	if enc.EncodeForJavascript(request.form("opc_cod_proyecto"))>"" then
		opc_cod_proyecto="1"
	end if

	if enc.EncodeForJavascript(request.form("opcfechaentrega"))>"" then
		opcfechaentrega="1"
	end if

	if enc.EncodeForJavascript(request.form("referencia"))>"" then
		referencia=limpiaCadena(request.form("referencia"))
	else
		referencia=limpiaCadena(request.querystring("referencia"))
	end if

	if referencia & ""="" then
		if enc.EncodeForJavascript(request.form("tdocumento"))>"" then
			referencia=limpiaCadena(request.form("tdocumento"))
		else
			referencia=limpiaCadena(request.querystring("tdocumento"))
		end if
	end if

	if enc.EncodeForJavascript(request.form("nombreart"))>"" then
		nombreart=limpiaCadena(request.form("nombreart"))
	else
		nombreart=limpiaCadena(request.querystring("nombreart"))
	end if

	if enc.EncodeForJavascript(request.form("familia"))>"" then
		familia=limpiaCadena(request.form("familia"))
	else
		familia=limpiaCadena(request.querystring("familia"))
	end if
	if enc.EncodeForJavascript(request.form("opcInsertaConceptos"))>"" then
		opcInsertaConceptos="1"
	end if

	'CheckCadena familia

	'cag
	nDias =limpiaCadena(request.form("nDias"))
	'fin cag

	' IML 28/04/2004 : Validamos si el usuario tiene acceso
	if viene="tienda" then
		sesionNCliente=left(almorigen,5)
		if sesionNCliente&""="" then sesionNCliente=session("ncliente")
		checkAccesoTienda sesionNCliente,"",almorigen
		almorigen=trimCodEmpresa(almorigen)
	else
		sesionNCliente=session("ncliente")
	end if
	checkCadenaTienda sesionNCliente,nserie
	checkCadenaTienda sesionNCliente,cod_proyecto
	checkCadenaTienda sesionNCliente,familia
	' FIN IML 28/04/2004 : Validamos si el usuario tiene acceso

	si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)

	WaitBoxOculto LitEsperePorFavor
	Alarma "anular_cantidades_ptesp.asp"

	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstPedido = Server.CreateObject("ADODB.Recordset")
	set rstPedido2 = Server.CreateObject("ADODB.Recordset")
	set rstPedido3 = Server.CreateObject("ADODB.Recordset")

	if mode="traerarticulo" then
		if referencia>"" then
			rstAux.open "select referencia,nombre from articulos with(nolock) where referencia='" & referencia & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if rstAux.EOF then
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgArticuloNoExiste%>");
				</script><%
				referencia=""
				nombreart=""
			else
				nombreart=rstAux("nombre")
			end if
			rstaux.close
		else
			nombreart=""
		end if
		mode="select1"
	end if

	strwhere=""

	if mode="imp" then
		'cag
		'Necesito gestionar las fechas Inicio y Fin; inputs de fechas o input numero dias hacia atras
		Dim vectDias
		vectDias= Array(31,28,31,30,31,30,31,31,30,31,30,31)

		diaHoy= day(date)
		mesHoy= month(date)
		anyoHoy= year(date)

		if nDias > "" then

			  numDias=nDias

		  
			'Calculo de anyo bisiesto				
			if (anyoHoy mod 4) = 0 then
				   vectDias(1)=29
			end if

			'Calcular fecha inicio de calculo
			if diaHoy-numDias > 0 then
			      diaInicio = diahoy - numDias
				  mesInicio = mesHoy
				  anyoInicio = anyoHoy
			else
   			   conta= diaHoy
			   mesPiv= mesHoy
    	       anyoInicio = anyoHoy
				   
			   faltan= numDias-conta
				   
			   while faltan>=0
			   		mesPiv= mesPiv-1
					if mesPiv=0 then
					     mesPiv=12
						 anyoInicio = anyoInicio -1
						 if (anyoInicio mod 4) = 0 then  'es bisiesto
			  				 vectDias(1)=29
						 end if
					end if
					if faltan < vectDias(mesPiv-1) then
					    diaInicio = vectDias(mesPiv-1)-faltan
						mesInicio = mesPiv
						faltan=-1
					else	
						faltan = faltan - vectDias(mesPiv-1)
					end if
		   	  wend
		    end if				   
			if mesInicio <=9 then
			   mesInicio = "0" & mesInicio
			end if

			fechaInicio=diaInicio & "/" & mesInicio & "/" & anyoInicio
			fechaHasta = day(date) & "/" & month(date) & "/" & year(date)

			'comprobar si las fechas se han calculado bien
			if not isdate(fechaInicio) or not isdate(fechaHasta) then
				if not isdate(fechaHasta) then
					%><script language="javascript" type="text/javascript">
						window.alert("<%=LitMsgDesdeFechaFecha%>");
					</script><%
				else
					%><script language="javascript" type="text/javascript">
						window.alert("<%=LitMsgHastaFechaFecha%>");
					</script><%
				end if
			end if
		else
'		   fechaInicio=""
'		   fechaHasta=""
			fechaInicio=diaHoy & "/" & mesHoy & "/" & anyoHoy
			fechaHasta = diaHoy & "/" & mesHoy & "/" & anyoHoy
		end if
		'fin cag%>
		<table width='100%' cellspacing="1" cellpadding="1">
   			<tr>
				<td width="30%" align="left" >
				</td>
				<td class=CELDARIGHT bgcolor="<%=color_blau%>">
					<%
						if fdesde>"" then
							if fhasta>"" then
								%><%=LitPeriodoFechas%> : <%=fdesde%> - <%=fhasta%><%
							else
								%><%=LitPeriodoFechas%> : <%=LitDesde%>&nbsp;<%=fdesde%><%
							end if
						else
							if fhasta>"" then
								%><%=LitPeriodoFechas%> : <%=LitHasta%>&nbsp;<%=fhasta%><%
							end if
						end if
						if nDias>"" then
  						 %><br/> <%=LitPedsAntA%> <%=fechaInicio%> <%					 
						end if%>
				</td>
	   		</tr>
		</table>
		<%if fdesde>"" or fhasta>"" then%>
			<hr/>
		<%end if
	end if

	if mode="select1" then
    dni=d_lookup("dni","personal","login='" & session("usuario") & "' and dni like '"+Session("ncliente")+"%'",session("dsn_cliente"))%>

		<SPAN ID="CapaNoAltaPersonal" style="display:none">
		 	<%waitbox LitMsgUsuarioPersonalNoExiste%>
		</SPAN>
  		<SPAN ID="CapaParametros" style="display:none">
		<%
			EligeCelda "input", "add", "left", "", "", 0, LitDesdeFecha, "fdesde", 10, fdesde
            DrawCalendar "fdesde"
            EligeCelda "input", "add", "left", "", "", 0, LitHastaFecha, "fhasta", 10, fhasta
            DrawCalendar "fhasta"

			rstSelect.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='PEDIDOS ENTRE ALMACENES' and nserie like '" & sesionNCliente & "%' order by descripcion asc",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "CELDA","175","",0,LitSerie,"nserie",rstSelect,nserie,"nserie","descripcion","",""
			rstSelect.close

			rstSelect.open "select codigo,descripcion from almacenes with(nolock) where codigo like '" & sesionNCliente & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "CELDA","175","",0,LitAlmOrigen,"almorigen",rstSelect,almorigen,"codigo","descripcion","",""
			rstSelect.close
			
			rstSelect.open "select codigo,descripcion from almacenes with(nolock) where codigo like '" & sesionNCliente & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "CELDA","175","",0,LitAlmDestino,"almdestino",rstSelect,almdestino,"codigo","descripcion","",""
			rstSelect.close			

			DrawDiv "1", "", ""
                DrawLabel "", "", LitPedPendServirConRef
				%><input type="text" maxlength='25' size="25" name="referencia" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(referencia>"",trimCodEmpresa(referencia),"")))%>" size=12 onchange="traerreferencia();">
				<%
            CloseDiv
			
            EligeCelda "input", "add", "left", "", "", 0, LitPedPendServirConNom, "nombreart", 25, nombreart

			rstAux.open " select codigo, nombre from familias with(nolock) where codigo like '" & sesionNCliente & "%' order by nombre asc", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
   			DrawSelectCelda "CELDA","175","",0,LitPedPendServirSubFamilia,"familia",rstAux,familia,"codigo","nombre","",""
			rstAux.close%>
		<hr/><%
                DrawDiv "3-sub", "background-color: #eae7e3", ""
                %><label class="ENCABEZADOL" style="text-align:left"><%=LitCamposOpcionales%></label>                    
                <%
                CloseDiv
			DrawDiv "1", "", ""
                DrawLabel "", "", LitPedidosAnt
                DrawInput "", "margin-right: 2px;", "nDias", "", ""%><label><%DrawSpan "CELDA", "", LitPedidosAntDias, ""%></label><% 
			CloseDiv%>
		</SPAN>
        <%if dni&""="" then
			%><script language="javascript" type="text/javascript">
				parent.botones.document.location="reg_inventario_bt.asp";
				CapaNoAltaPersonal.style.display = "";
			</script><%
		else
			%><script language="javascript" type="text/javascript">
				CapaParametros.style.display = "";
			</script><%
		end if
	elseif mode="imp" then
		MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='049'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='049'", DSNIlion)
		%><input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>
		<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'>

		<input type="hidden" name="fdesde" value="<%=EncodeForHtml(fdesde)%>">
		<input type="hidden" name="fhasta" value="<%=EncodeForHtml(fhasta)%>">
		<input type="hidden" name="almorigen" value="<%=EncodeForHtml(almorigen)%>">
		<input type="hidden" name="almdestino" value="<%=EncodeForHtml(almdestino)%>">
		<input type="hidden" name="nserie" value="<%=EncodeForHtml(nserie)%>">
		<input type="hidden" name="opcalmacenbaja" value="<%=EncodeForHtml(opcalmacenbaja)%>">
		<input type="hidden" name="opcElimPedSinDet" value="<%=EncodeForHtml(opcElimPedSinDet)%>">
		<input type="hidden" name="nDias" value="<%=EncodeForHtml(nDias)%>">
		<input type="hidden" name="actividad" value="<%=EncodeForHtml(actividad)%>">
		<input type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(cod_proyecto)%>">
		<input type="hidden" name="opc_cod_proyecto" value="<%=EncodeForHtml(opc_cod_proyecto)%>">
		<input type="hidden" name="opcInsertaConceptos" value="<%=EncodeForHtml(opcInsertaConceptos)%>">
		<input type="hidden" name="opcfechaentrega" value="<%=EncodeForHtml(opcfechaentrega)%>">
		<% if viene="tienda" then%>
			<input type="hidden" name="viene" value="<%=EncodeForHtml(viene)%>">
			<input type="hidden" name="campo" value="<%=EncodeForHtml(campo)%>">
			<input type="hidden" name="criterio" value="<%=EncodeForHtml(criterio)%>">
			<input type="hidden" name="texto" value="<%=EncodeForHtml(texto)%>">
		<%end if%>
		<%if viene="articulos" then%>
			<input type="hidden" name="viene" value="<%=EncodeForHtml(viene)%>">
		<%end if%>
		<input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>">
		<input type="hidden" name="nombreart" value="<%=EncodeForHtml(nombreart)%>">
		<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>">
		<%

		if viene<>"tienda" then
			VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarPedidosPro)=1:VinculosPagina(MostrarArticulos)=1:VinculosPagina(MostrarPedidosCli)=1
			CargarRestricciones session("usuario"),sesionNCliente,Permisos,Enlaces,VinculosPagina
		end if

		total_valor_general = 0
		total_pendiente_general = 0
		encabezado=0
		'strwhere="where"

		strwhere=CadenaBusqueda(campo,criterio,texto)

		if nserie > "" then
			strwhere=strwhere & " nserie='" & nserie & "' and"
			%><font class='CELDA'><b><%=LitSerie%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nserie))%>&nbsp;<%=d_lookup("nombre","series","nserie='" & nserie & "'",session("dsn_cliente"))%></font><br/><%
			encabezado=1
		end if

		if fechaInicio>"" and fechaHasta>"" then
			strwhere=strwhere & " fecha<='" & fechaInicio & "' and"			
		end if
		if fdesde > "" then 
			strwhere=strwhere & " fecha>='" & fdesde & "' and"
		end if	
		if fhasta > "" then 
		    strwhere=strwhere & " fecha<='" & fhasta & "' and"
		end if
		if almorigen > "" then 
		    strwhere=strwhere & " pedidos_tienda.almorigen='" & almorigen & "' and"
			%><font class='CELDA'><b><%=LitAlmOrigen%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(almorigen))%>&nbsp;<%=d_lookup("descripcion","almacenes","codigo='" & almorigen & "'",session("dsn_cliente"))%></font><br/><%
		    
		end if
		if almdestino > "" then 
		    strwhere=strwhere & " pedidos_tienda.almdestino='" & almdestino & "' and"
			%><font class='CELDA'><b><%=LitAlmDestino%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(almorigen))%>&nbsp;<%=d_lookup("descripcion","almacenes","codigo='" & almdestino & "'",session("dsn_cliente"))%></font><br/><%
		end if

		'Tenemos una clausula where para los detalles y otra para los conceptos
		strwhereConceptos=strwhere

		if referencia>"" then
			if viene<>"articulos" then
				referencia=sesionNCliente & referencia
				strwhere=strwhere & " detalles_ped_tienda.referencia like '%" & referencia & "%' and"
						
				%><font class='CELDA'><b><%=LitPedPendServirConRef%> : </b></font><%=EncodeForHtml(trimCodEmpresa(referencia))%><br/><%
			else
				strwhere=strwhere & " detalles_ped_tienda.referencia='" & referencia & "' and"
				
				%><font class='CELDA'><b><%=LitPedPendServirConRef%> : </b></font><%=EncodeForHtml(trimCodEmpresa(referencia))%><br/><%
			end if

			encabezado=1
		end if
		if nombreart>"" and referencia & ""="" then
			strwhere=strwhere & " articulos.referencia=detalles_ped_tienda.referencia and articulos.nombre like '%" & nombreart & "%' and"
				
			%><font class='CELDA'><b><%=LitPedPendServirConNom%> : </b></font><%=EncodeForHtml(nombreart)%><br/><%
			encabezado=1
		end if
		if familia>"" then
			strwhere=strwhere & " articulos.referencia=detalles_ped_tienda.referencia and articulos.familia='" & familia & "' and"
					
			%><font class='CELDA'><b><%=LitPedPendServirSubFamilia%> : </b></font><%=d_lookup("nombre","familias","codigo='" & familia & "'",session("dsn_cliente"))%><br/><%
			encabezado=1
		end if

		strwhere=strwhere + " pedidos_tienda.almdestino=almacenes.codigo and nmovimiento is null and nalbaran is null"
		strwhere=strwhere + " and pedidos_tienda.npedido=detalles_ped_tienda.npedido"
		strwhere=strwhere + " and pedidos_tienda.npedido like '" & sesionNCliente & "%' "

		strwhere=strwhere + " and cantpendiente<>0 "

	strwhereAntes=strwhere

		strwhere2=strwhere + "  GROUP BY pedidos_tienda.npedido, pedidos_tienda.almorigen, fecha"
		'cag strwhere2=strwhere2 + " ORDER BY pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"
		''cag2
		strwhere=strwhere + " order by pedidos_tienda.cliente, fecha, pedidos_tienda.npedido"

'Calculos de páginas--------------------------

		strselect="select count(pedidos_tienda.npedido) as contador,pedidos_tienda.npedido,pedidos_tienda.almorigen,fecha,'detalles' "
		strselect=strselect & " from pedidos_tienda,tiendas,detalles_ped_tienda, almacenes "

		if (nombreart>"" and referencia & ""="") or familia>"" then
			strselect=strselect & ",articulos "
		end if
		'strselect=strselect & strwhere2
		strselect=strselect & strwhere2 ''& " union " & strselectConceptos & strwhereConceptos		

		rstPedido2.cursorlocation=3
		rstPedido2.Open strselect, session("dsn_cliente")
		if not rstPedido2.eof then
			suma=0

			while not rstPedido2.eof
				suma=suma+2+rstPedido2("contador")
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
				lista(1,3)=rstPedido2("npedido")
			else
				lista(1,3)=""
			end if
			sumat=0
			suma2=0
			j=0
			while not rstPedido2.eof
				suma2=suma2  + 2 + rstPedido2("contador")
				lista(i,2)=suma2
				lista(i,1)=lista(i,1)+(rstPedido2("contador"))
				j=j+1
				if suma2>=MAXPAGINA then
					lista(i,4)=rstPedido2("npedido")
					i=i+1
					j=0
					lista(i,1)=0
					lista(i,2)=0
					sumat=sumat+suma2
					suma2=0
					rstPedido2.movenext
					if not rstPedido2.eof then
						lista(i,3)=rstPedido2("npedido")
					else
						rstPedido2.moveprevious
						lista(i,3)=rstPedido2("npedido")
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

		'cag
		'strselect="select item,detalles_ped_pro.descripcion,fecha,pedidos_pro.nproveedor,pedidos_pro.npedido,pedidos_pro.divisa,pedidos_pro.fecha_entrega,cantidad,detalles_ped_pro.referencia,almacenes.descripcion as almacen,detalles_ped_pro.importe,cod_proyecto,'detalle' as tipo,detalles_ped_pro.cantidadpend, detalles_ped_pro.npedidopro,detalles_ped_pro.itempedidocli "
		strselect="select item,detalles_ped_tienda.descripcion,fecha,pedidos_tienda.almorigen,pedidos_tienda.almdestino,pedidos_tienda.npedido,cantidad,detalles_ped_tienda.referencia,almacenes.codigo as codAlmacen,almacenes.descripcion as almacen,'detalle' as tipo,detalles_ped_tienda.cantpendiente  "
		'fin cag
		strselect=strselect & " from detalles_ped_tienda with(nolock) ,pedidos_tienda with(nolock)"
		strselect=strselect & "left outer join almacenes with(nolock) on almacenes.codigo=pedidos_tienda.almdestino "

		if (nombreart>"" and referencia & ""="") or familia>"" then
			strselect=strselect & ",articulos "
		end if

'		strselect=strselect & strwhereAntes & " and mainitem is null union " & strselectConceptos & strwhereConceptosAntes & ",detalles_ped_pro.referencia desc"
		strselect=strselect & strwhereAntes & " order by pedidos_tienda.fecha,pedidos_tienda.npedido " 

		rstPedido.cursorlocation=2
		rstPedido.Open strselect,session("dsn_cliente")
		if rstPedido.EOF then
			rstPedido.Close
			%><input type="hidden" name="NumRegsTotal" value="0">
			        
			<div class="CEROFILAS"><%=LitMsgDatosNoExiste%></div><%
			
		else
			%><input type="hidden" name="NumRegsTotal" value="<%=rstPedido.RecordCount%>"><%

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
				%>
				<hr/>
				<%
			end if

'			NavPaginas lote,lotes,campo,criterio,texto,1


			%>
			<table width='100%' border='<%=borde%>' style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
			'Fila de encabezado
			ncliente_ant="#@#"
			npedido_ant = "#@#"
			filab=0
			filac=1			
		'cag	Con este while controlo que no se pagine
			nregistros=1   'Para controlar el check de una fila en concreto
		
		while lote<=lotes
		'fin cag
			while not rstPedido.EOF and filab<lista(lote,2) 'MAXPAGINA
						if rstPedido("almorigen")<>almorigen_ant and rstPedido("npedido")<>npedido_ant then
							DrawFila color_blau
								DrawCeldaSpan "ENCABEZADOL","","",0,"<hr/>",11
							CloseFila
							DrawFila color_blau

								dat1=trimCodEmpresa(rstPedido("almorigen")) & " - " & d_lookup("descripcion","almacenes","codigo='" & rstPedido("almorigen") & "'",session("dsn_cliente"))
								DrawCeldaSpan "ENCABEZADOL","","",0,LitAlmOrigen & " : " & dat1,8
							CloseFila
							almorigen_ant = rstPedido("almorigen")
						end if
						
						    ''cag 2
							if rstPedido("npedido")<>npedido_ant and rstPedido("tipo")="detalle" then
							'if rstPedido("npedido")<>npedido_ant then
								DrawFila color_terra
									DrawCelda "ENCABEZADOL7","5%","",0," "
									if viene<>"tienda" then
										dat1=Hiperv(OBJPedidosAlm,rstPedido("npedido"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("npedido")),LitVerPedido)
									else
										dat1="<a class='CELDAREFB' href=javascript:ver_pedido('" & rstPedido("npedido") & "','"&almorigen&"') alt='" & LitVerPedido & "'>" & trimCodEmpresa(rstPedido("npedido")) & "</a>"
									end if
                                        dat2=" - "&LitFecha &" "&rstPedido("fecha")
									'DrawCeldaSpan "ENCABEZADOL","","",0,LitPedidoMin + " : " & dat1 & dat2,9
									DrawCeldaSpan "ENCABEZADOL","","",0,LitPedido + " : " & dat1 & dat2,10
								CloseFila
								npedido_ant = rstPedido("npedido")
								'cag 2
								if rstPedido("tipo")="detalle" then		
								'fin cag2
								DrawFila color_terra
									DrawCelda "ENCABEZADOL7","5%","",0," "
									DrawCelda "ENCABEZADOL7","5%","",0," "
									'cag
									DrawCelda "ENCABEZADOL7","","",0,LitEliminar
									'fin cag
									DrawCelda "ENCABEZADOR7","","",0,LitCantidad
									DrawCelda "ENCABEZADOL7","","",0,LitReferencia
									DrawCelda "ENCABEZADOL7","","",0,LitDescripcion

									DrawCelda "ENCABEZADOL7","","",0,LitAlmDestino
									DrawCelda "ENCABEZADOL7","","",0,LitPendiente
								CloseFila
								'cag2
								end if 		 'fin del if para no mostrar las lineas de conceptos
								'fin cag2
								filab=filab+2
							end if
								'cag 2
								if rstPedido("tipo")="detalle" then		
								'fin cag2
								DrawFila color_blau
									DrawCelda "ENCABEZADOL7","5%","",0," "
									DrawCelda "ENCABEZADOL7","5%","",0," "%>
									<!-- cag -->
			   						 <td class="ENCABEZADOC7" width="5%" align="center"> <%'=nregistros%>
									   	<input type="checkbox" name='checkElim<%=nregistros%>' > </td>
										<input type="hidden" name="nPedido<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("npedido"))%>">
										<input type="hidden" name="ctdad<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("cantidad"))%>">
										<input type="hidden" name="ctdadPend<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("cantpendiente"))%>">
										<input type="hidden" name="almorigen<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("almdestino"))%>">
										<input type="hidden" name="item<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("item"))%>">										
										 <%
									'fin cag
									DrawCelda "CELDAR7","","",0,EncodeForHtml(rstPedido("cantidad"))
									if viene<>"tienda" then
										dat1=Hiperv(OBJArticulos,rstPedido("referencia"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("referencia")),LitVerArticulo)
										'dat3=Hiperv(OBJPedidosAlm,rstPedido("npedidopro"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("npedidopro")),LitVerPedido)
									else
										dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & enc.EncodeForJavascript(rstPedido("referencia")) & "','" & enc.EncodeForJavascript(rstPedido("almorigen")) & "') alt='" & LitVerArticulo & "' border='0'>" & EncodeForHtml(trimCodEmpresa(rstPedido("referencia"))) & "</a>"
										'dat3="<a class='CELDAREFB' href=javascript:ver_pedidoCli('" & rstPedido("npedidopro") & "') alt='" & LitVerPedido & "'>" & trimCodEmpresa(rstPedido("npedidopro")) & "</a>"
									end if
									DrawCelda "CELDAL7","","",0,dat1
									dat1=rstPedido("descripcion")
									DrawCelda "CELDAL7","","",0,dat1
									DrawCelda "CELDAL7","15","",0,EncodeForHtml(trimcodempresa(rstPedido("codalmacen")))&"-"&EncodeForHtml(rstPedido("almacen"))
									DrawCelda "CELDAL7","","",0,EncodeForHtml(rstPedido("cantpendiente"))
									'DrawCelda "CELDAL7","","",0,rstPedido("npedidopro")
									'DrawCelda "CELDAL7","","",0,dat3
									'DrawCelda "CELDAR7","","",0,cstr(null_z(rstPedido("Importe")))+" "+d_lookup("abreviatura","divisas","codigo='" & rstPedido("divisa") & "'",session("dsn_cliente"))
								CloseFila
							end if 'fin del tipo=detalle
							filab=filab+1

				rstPedido.MoveNext
				filac=filac+1
				'cag
				nregistros=nregistros+1
				'fin cag
				Response.Flush
			wend

			' si no se ha acabado el pedido, se escribira hasta que se acabe
			filab=0
			if not rstPedido.eof then
				continuar=1
				while continuar=1
					if rstPedido("npedido")=npedido_ant and rstPedido("almorigen")=ncliente_ant then
						'cag 2
						if rstPedido("tipo")="detalle" then		
						'fin cag2
						DrawFila color_blau
							DrawCelda "ENCABEZADOL7","5%","",0," "
							DrawCelda "ENCABEZADOL7","5%","",0," " %>
							<!-- cag -->
	   						<td class="ENCABEZADOL7" width="5%" align="center"> <%'=nregistros%>
							    <input type="checkbox" name='checkElim<%=nregistros%>' ></td>
								<input type="hidden" name="nPedido<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("npedido"))%>">
								
								<input type="hidden" name="ctdad<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("cantidad"))%>">
								<input type="hidden" name="ctdadPend<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("cantidadpend"))%>">
								<input type="hidden" name="almorigen<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("almorigen"))%>">
								<input type="hidden" name="item<%=nregistros%>" value="<%=EncodeForHtml(rstPedido("item"))%>">
							   <%
							'fin cag
							DrawCelda "CELDAR7","","",0,EncodeForHtml(rstPedido("cantidad"))
							if viene<>"tienda" then
								dat1=Hiperv(OBJArticulos,rstPedido("referencia"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("referencia")),LitVerArticulo)
							else
								dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & enc.EncodeForJavascript(rstPedido("referencia")) & "','" & enc.EncodeForJavascript(rstPedido("almorigen")) & "') alt='" & LitVerArticulo & "' border='0'>" & rstPedido("referencia") & "</a>"
							end if
							DrawCelda "CELDAL7","","",0,dat1
							dat1=rstPedido("descripcion")
							DrawCelda "CELDAL7","","",0,dat1
							if opcfechaentrega="1" then
								DrawCelda "CELDAL7","","",0,EncodeForHtml(rstPedido("fecha_entrega"))
							end if
							DrawCelda "CELDAL7","","",0,EncodeForHtml(rstPedido("almacen"))
							DrawCelda "CELDAL7","","",0,EncodeForHtml(rstPedido("cantidadpend"))
							DrawCelda "CELDAL7","","",0,EncodeForHtml(rstPedido("npedidopro"))
							DrawCelda "CELDAL7","","",0,iif(rstPedido("npedidopro")>"",rstPedido("itempedidopro"),"")
							DrawCelda "CELDAR7","","",0,cstr(rstPedido("Importe"))+" "+d_lookup("abreviatura","divisas","codigo='" & rstPedido("divisa") & "'",session("dsn_cliente"))
						CloseFila
						'cag2
						end if  'fin del tipo=detalle
						'fin cag2
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
						'cag
						nregistros=nregistros+1
						'fin cag						
					else
						continuar=0
					end if
					Response.Flush
				wend

				%><input type="hidden" name="npedido_ant" value="<%=EncodeForHtml(npedido_ant)%>">
				<input type="hidden" name="ncliente_ant2" value="<%=EncodeForHtml(ncliente_ant)%>"><%
			end if
		'cag	
			lote=lote+1
        wend  ' del while lote<=lotes
           %><input type="hidden" name="h_nregistros" value="<%=EncodeForHtml(nregistros-1)%>"><%
		   
		'fin cag
			%></table>
			<hr/><%

			rstPedido.Close
'			NavPaginas lote,lotes,campo,criterio,texto,2
		end if
		%>

	<%
	'cag
	elseif mode="delete" then
  		Nregistros=null_z(h_nregistros)
		lista="("
		lista2="("
		lista3=""
		'listaPedidosItems="("
		'listaPedidos="("

		' Con esto obtengo todos los registros seleccionados
  		for i=1 to Nregistros
			nombre="checkElim" & i
			'nombreCli="cliente" & i
			pedidoSeleccionado = "nPedido"&i
			itemSeleccionado = "item"&i
			if request.form(nombre) > "" then ' DOCUMENTO SELECCIONADO
			    filaSeleccionada = i
				ndocumento=trim(limpiaCadena(request.form(nombre)))
				lista=lista & "''" & ndocumento & "'',"
				filasSel=lista & "''" & filaSeleccionada & "'',"
			end if
		next
	
		lista=mid(lista,1,len(lista)-1) & ")" 'Quitamos la última coma y cerramos el paréntesis
		'listaPedidosItems = mid(listaPedidosItems,1,len(listaPedidosItems)-1) & ")"
		'listaPedidos = mid(listaPedidos,1,len(listaPedidos)-1) & ")"
'		lista3=mid(lista3,1,len(lista3)-1)    'Quitamos la última coma

		if lista=")" then
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgDocumentoNoSel%>");
				parent.botones.document.location = "anular_cantidades_ptesp_bt.asp?mode=select1";
                document.location="anular_cantidades_ptesp.asp?mode=select1";
				
			</script><%
	    end if
		
		
		'Creo temporal de usuario para realizar anulación masiva en el procedimiento
		strdrop ="if exists (select * from sysobjects where id = object_id('egesticet.[" & session("usuario") & "]') ) drop table egesticet.[" & session("usuario") & "]"		
		rst.open strdrop,session("dsn_cliente")
		
		strselect="create table [" & session("usuario") & "] (numPedido varchar(20), ctdad real, ctdadPend real,  almorigen varchar(10), item smallint)"
		rst.open strselect,session("dsn_cliente")

        'rst.close
  		for i=1 to Nregistros
			nombre="checkElim" & i
			if request.form(nombre) > "" then ' DOCUMENTO SELECCIONADO
			   referencia="refer" & i
			   numPedido="nPedido"& i
			   cantidad="ctdad" & i
			   cantidadPdte="ctdadPend" & i
			   almacen="almacen" & i
			   codAlmacen="codAlmacen" & i
			   almorigen="almorigen" & i
			   nItem="item" & i
			   strselect="insert into [" & session("usuario") & "] ( numPedido, ctdad, ctdadPend,almorigen,item) "
			   strselect= strselect & " values ('" & request.form(numPedido) & "'," & replace(request.form(cantidad),",",".") & "," & replace(request.form(cantidadPdte),",",".") & ",'" & request.form(almorigen) & "','" & request.form(nItem)& "')"
	   		   rst.open strselect,session("dsn_cliente")
			end if
		next

    'llamada al procedimiento para anulación masiva

		Resultado="0"
		nomusuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",dsnilion)

		set Command =  Server.CreateObject("ADODB.Command")
		
	    set conn = Server.CreateObject("ADODB.Connection")
		conn.ConnectionTimeout = 300
		conn.CommandTimeout = 300
        set conn2 = Server.CreateObject("ADODB.Connection")
	    conn2.ConnectionTimeout = 300
	    conn2.CommandTimeout = 300
	
		conn.open session("dsn_cliente")
		Command.ActiveConnection =conn
		Command.CommandTimeout = 0
		Command.CommandText="AnularCantidadesPendientesProductos"
		Command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
'		Command.Parameters.Append Command.CreateParameter("@opcElimPedidos",adInteger,adParamInput,2,opcEliminar)

		Command.Parameters.Append Command.CreateParameter("@opcusuario", adVarChar, adParamInput,50, session("usuario"))		
		Command.Parameters.Append Command.CreateParameter("@sesion_ncliente", adVarChar, adParamInput,5, session("ncliente"))
		Command.Parameters.Append Command.CreateParameter("@nameUsuario", adVarChar, adParamInput,25, nomusuario)
		Command.Parameters.Append Command.CreateParameter("@ipUsuario", adVarChar, adParamInput,255, Request.ServerVariables("REMOTE_ADDR"))
		Command.Parameters.Append Command.CreateParameter("@resul", adVarChar, adParamOutput,2, Resultado)
		Command.Execute,,adExecuteNoRecords
		
		Resultado = Command.Parameters("@resul").Value
		conn.close
		set command=nothing
		set conn=nothing

	     if Resultado="0" then
			%>
		 <script language="javascript" type="text/javascript">
             window.alert("<%=LitProcesoFinCorrecto%>")
			 parent.botones.document.location = "anular_cantidades_ptesp_bt.asp?mode=select1";
             document.location="anular_cantidades_ptesp.asp?mode=select1";
			 
		  </script>
		<%end if
		'fin coger excepcion y parametro salida
	'fin cag
	end if%>
</form>
<%end if
set rstSelect=nothing
set rstAux=nothing
set rst=nothing
set rstPedido=nothing
set rstPedido2=nothing
set rstPedido3=nothing
set conn2=nothing
%>
</body>
</html>