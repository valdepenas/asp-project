<%@ Language=VBScript %>
<%
  '***RGU 20/12/2005: Añadir al listado columnas de "cantidad pendiente", "Pedido Cliente", "Item"(cliente)
  response.buffer=true
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">

<head>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  


<%if request.querystring("viene")="tienda" or request.form("viene")="tienda" then
	titulo=LitTituloListadoPen2
else
	titulo=LitTituloListadoPen
end if%>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../productos/articulos.inc" -->
<!--#include file="../productos/listados/listado_articulos.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../CatFamSub.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../catFamSub.inc" -->
<!--#include file="pedidos_pro.inc" -->
<!--#include file="../styles/formularios.css.inc"-->
<!--#include file="../styles/formulariosAnularCantidades.css.inc"-->

<title><%=titulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<meta http-equiv="Content-style-Type" content="text/css"/>
<link rel="stylesheet" href="../pantalla.css" media="screen"/>
<link rel="stylesheet" href="../impresora.css" media="print"/>
</head>
<script type="text/javascript" language="javascript" src="../jfunciones.js"></script>
<script type="text/javascript" language="javascript" src="/lib/js/shortKey.js"></script>
<script type="text/javascript" language="javascript">
function traerreferencia(){
}

//Desencadena la búsqueda del proveedor cuya referencia se indica
function TraerProveedor(mode) {
	document.anular_cantidades_ptes.action="anular_cantidades_ptes.asp?mode=traerproveedor"
	document.anular_cantidades_ptes.submit();
}

function ver_pedido(npedido,nproveedor){
	if (nproveedor!="")
		parent.document.location="../compras/pedidos_pro_imp.asp?npedido=('" + npedido + "')&mode=browse&empresa="+nproveedor.substr(0,5);
	else parent.document.location="../compras/pedidos_pro_imp.asp?npedido=('" + npedido + "')&mode=browse&empresa=<%=session("ncliente")%>";
	parent.parent.topFrame.document.getElementById("regresar").style.display="";
}
function ver_pedidoCli(npedido){
	parent.document.location="../ventas/pedidos_cli_imp.asp?npedido=('" + npedido + "')&mode=browse&empresa=<%=session("ncliente")%>";
	parent.parent.topFrame.document.getElementById("regresar").style.display="";
}

function Ver_Articulo(referencia,nproveedor){
	pagina="../tiendas/pedido_ficha_articulo.asp?referencia=" + referencia + "&viene=anular_cantidades_ptes&nproveedor=" + nproveedor;
	ven=AbrirVentana(pagina,'C',250,400);

}
</script>
<body class="BODY_ASP" onload="self.status='';">
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
<form name="anular_cantidades_ptes" method="post">
<% PintarCabecera "anular_cantidades_ptes.asp"
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

    'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
    numSerieTienda	= limpiaCadena(Request.QueryString("numSerieTienda"))
	if numSerieTienda = "" then
		numSerieTienda	= limpiaCadena(Request.form("numSerieTienda"))
	end if
    '--------------------------------------------------------  

	'CheckCadena nserie

	if request.form("opcproveedorbaja")>"" then
		opcproveedorbaja=limpiaCadena(request.form("opcproveedorbaja"))
	else
		opcproveedorbaja=limpiaCadena(request.querystring("opcproveedorbaja"))
	end if
	
	'cag
	if request.form("opcElimPedSinDet")>"" then
		opcElimPedSinDet=limpiaCadena(request.form("opcElimPedSinDet"))
	else
		opcElimPedSinDet=limpiaCadena(request.querystring("opcElimPedSinDet"))
	end if
	if opcElimPedSinDet>"" then
	   opcEliminar=1
	else
	    opcEliminar=0
	end if

	if request.form("h_nregistros")>"" then
		h_nregistros=limpiaCadena(request.form("h_nregistros"))
	else
		h_nregistros=limpiaCadena(request.querystring("h_nregistros"))
	end if
   'fin cag
   
	if request.form("nDias")>"" then
		nDias=limpiaCadena(request.form("nDias"))
	else
		nDias=limpiaCadena(request.querystring("nDias"))
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

    'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
    refCategoriaCombustible	= limpiaCadena(Request.QueryString("refCategoriaCombustible"))
	if refCategoriaCombustible = "" then
		refCategoriaCombustible	= limpiaCadena(Request.form("refCategoriaCombustible"))
	end if
    '------------------------------------------------------------------------

	'CheckCadena familia
	if request.form("opcInsertaConceptos")>"" then
		opcInsertaConceptos="1"
	end if
	if request.querystring("almacen")>"" then
		TmpSoloAlmacen=limpiaCadena(request.querystring("almacen"))
	else
		TmpSoloAlmacen=limpiaCadena(request.Form("almacen"))
	end if

    'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
    numAlmacenTienda	= limpiaCadena(Request.QueryString("numAlmacenTienda"))
	if numAlmacenTienda = "" then
		numAlmacenTienda	= limpiaCadena(Request.form("numAlmacenTienda"))
	end if
    '------------------------------------------------------------------------

	CheckCadena TmpSoloAlmacen	
	'cag
	nDias =limpiaCadena(request.form("nDias"))
	'fin cag

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
	Alarma "anular_cantidades_ptes.asp"

	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstPedido = Server.CreateObject("ADODB.Recordset")
	set rstPedido2 = Server.CreateObject("ADODB.Recordset")
	set rstPedido3 = Server.CreateObject("ADODB.Recordset")

	if mode="traerarticulo" then
		if referencia>"" then
            'rstAux.cursorlocation=3
			'rstAux.open "select referencia,nombre from articulos with(nolock) where referencia like '"&session("ncliente")&"%' and referencia='" & referencia & "'",session("dsn_cliente")
            set conn=  server.CreateObject("ADODB.Connection")
            set cmd=  server.CreateObject("ADODB.Command")
            conn.open session("dsn_cliente")
            cmd.ActiveConnection=conn
            conn.cursorlocation=3
            cmd.CommandText="select referencia,nombre from articulos with(nolock) where referencia like ?+'%' and referencia=?"
	        cmd.CommandType = adCmdText 
            cmd.Parameters.Append cmd.CreateParameter("@ncliente", adVarChar,adParamInput,30,session("ncliente")&"")
            cmd.Parameters.Append cmd.CreateParameter("@referencia", adVarChar,adParamInput,30,referencia&"")
            set rstAux=cmd.execute

			if rstAux.EOF then
				%><script type="text/javascript" language="javascript">
					window.alert("<%=LitMsgArticuloNoExiste%>");
					//document.anular_cantidades_ptes.referencia.focus();
					//document.anular_cantidades_ptes.referencia.select();
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

	if mode="traerproveedor" then
		if nproveedor>"" then
			nproveedor=session("ncliente") & completar(nproveedor,5,"0")
            'rstAux.cursorlocation=3
			'rstAux.open "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like '"&session("ncliente")&"%' and nproveedor='" & nproveedor & "'",session("dsn_cliente")

            set conn=  server.CreateObject("ADODB.Connection")
            set cmd=  server.CreateObject("ADODB.Command")
            conn.open session("dsn_cliente")
            cmd.ActiveConnection=conn
            conn.cursorlocation=3
            cmd.CommandText="select nproveedor,razon_social from proveedores with(nolock) where nproveedor like ?+'%' and nproveedor=?"
	        cmd.CommandType = adCmdText 
            cmd.Parameters.Append cmd.CreateParameter("@ncliente", adVarChar,adParamInput,30,session("ncliente")&"")
            cmd.Parameters.Append cmd.CreateParameter("@nproveedor", adVarChar,adParamInput,30,nproveedor&"")
            set rstAux=cmd.execute

			if rstAux.EOF then
				%><script type="text/javascript" language="javascript">
					window.alert("<%=LitMsgProveedorNoExiste%>");
					//document.anular_cantidades_ptes.nproveedor.focus();
					//document.anular_cantidades_ptes.nproveedor.select();
				</script><%
				nproveedor=""
				nombre=""
			else
				nombre=rstAux("razon_social")
			end if
			rstaux.close
		else
			nombre=""
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
					%><script type="text/javascript" language="javascript">
						alert("<%=LitMsgDesdeFechaFecha%>");
					</script><%
				else
					%><script type="text/javascript" language="javascript">
						alert("<%=LitMsgHastaFechaFecha%>");
					</script><%
				end if
			end if
		else
			fechaInicio=diaHoy & "/" & mesHoy & "/" & anyoHoy
			fechaHasta = diaHoy & "/" & mesHoy & "/" & anyoHoy
		end if

		'fin cag%>
		<table width='100%' cellspacing="1" cellpadding="1">
   			<tr>
				<td width="30%" align="left" >
				</td>
				<td class="CELDARIGHT customBackground" bgcolor="<%=color_blau%>">
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

	if mode="select1" then%>
		<table width='100%' border='<%=borde%>' cellspacing="1" cellpadding="1"><%
			DrawFila color_blau
				DrawCelda2 "CELDA style='width:145px'", "left", false, LitDesdeFecha+" : "
			 	DrawInputCelda "CELDA","","",10,0,"","fdesde",fdesde
				DrawCelda2 "CELDA style='width:120px'", "left", false,""
				DrawCelda2 "CELDA style='width:100px'", "left", false,LitHastaFecha+" : "
			 	DrawInputCelda "CELDA","","",10,0,"","fhasta",fhasta
			CloseFila
			DrawFila color_blau
				DrawCelda2 "CELDA style='width:145px'", "left", false, LitDesdeFechaEntrega+" : "
			 	DrawInputCelda "CELDA","","",10,0,"","fentregadesde",fentregadesde
				DrawCelda2 "CELDA style='width:120px'", "left", false,""
				DrawCelda2 "CELDA style='width:145px'", "left", false,LitHastaFechaEntrega+" : "
			 	DrawInputCelda "CELDA","","",10,0,"","fentregahasta",fentregahasta
			CloseFila
			DrawFila color_blau
				DrawCelda2 "CELDA style='width:145px'", "left", false, LitActividad
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo,descripcion from tipo_actividad with(nolock) where codigo like '" & sesionNCliente & "%' order by descripcion",session("dsn_cliente")
				DrawSelectCelda "CELDA","175","",0,"","actividad",rstSelect,actividad,"codigo","descripcion","",""
				rstSelect.close

                'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
                
                'Consulta Tienda Usuario                
                set conn3 = Server.CreateObject("ADODB.Connection")
                set command3 =  Server.CreateObject("ADODB.Command")
                conn3.open session("dsn_cliente")
                command3.ActiveConnection = conn3
                command3.CommandTimeout = 0
                command3.CommandText="REPSOLPERUGetStoreFromPersonal" 'Nombre Procedimiento
                command3.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                command3.Parameters.Append command3.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))
                command3.Parameters.Append command3.CreateParameter("@personal", adVarChar, adParamInput, 50, session("ncliente") + session("usuario"))
            
                set rstTienda = command3.execute
                set rstTiendaAux = command3.execute            

                tiendaUsuario = ""              

                if not rstTiendaAux.eof then
                    tiendaUsuario = rstTiendaAux("STORE")
                end if

                rstTienda.close    
                rstTiendaAux.close 
                set rstTienda=nothing
                set rstTiendaAux=nothing

                'Consulta Identificador Categoria 'Combustible'
                set rstCombustible = Server.CreateObject("ADODB.Recordset")
                query = "select codigo from categorias where codigo like '" + session("ncliente") + "%' and nombre LIKE 'Combustible'"
                if rstCombustible.state<>0 then rstCombustible.close
                rstCombustible.open query, session("dsn_cliente")

                refCategoriaCombustible = ""  

                if not rstCombustible.eof then
                    refCategoriaCombustible = rstCombustible("CODIGO")
                end if

                rstCombustible.close    
                set rstCombustible=nothing

                %>
                <input type="hidden" name="refCategoriaCombustible" value="<%=enc.EncodeForHtmlAttribute(null_s(refCategoriaCombustible))%>"/>
                <%

                'Consulta Series
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 =  Server.CreateObject("ADODB.Command")
                conn2.open session("dsn_cliente")
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 0

                If tiendaUsuario = "" Then 'El usuario no tiene tienda asignada                                
                    command2.CommandText="getAllSeriesByTypeDocument" 'Nombre Procedimiento
                    command2.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    command2.Parameters.Append command2.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))
                    command2.Parameters.Append command2.CreateParameter("@type_document", adVarChar, adParamInput, 50, "PEDIDO A PROVEEDOR")                               
                Else 'El usuario tiene tienda asignada
                    command2.CommandText="GetAllSeriesByTypeDocumentAndStore" 'Nombre Procedimiento
                    command2.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    command2.Parameters.Append command2.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))
                    command2.Parameters.Append command2.CreateParameter("@type_document", adVarChar, adParamInput, 50, "PEDIDO A PROVEEDOR")
                    command2.Parameters.Append command2.CreateParameter("@store", adVarChar, adParamInput, 10, tiendaUsuario) 
                End If

                set rstSeries = command2.execute
                set rstSeriesAux = command2.execute 
                numSerieTienda = ""

                dim listaSeries
                listaSeries = ""

                While not rstSeriesAux.eof
		            listaSeries = listaSeries & rstSeriesAux("nserie") & ","
                    rstSeriesAux.MoveNext
                Wend

                if listaSeries > "" Then
		            numSerieTienda = left(listaSeries, len(listaSeries)-1)
                End If

				DrawCelda2 "CELDA style='width:120px'", "left", false,""
				DrawCelda2 "CELDA", "left", false, LitSerie +": "              
				DrawSelectCelda "CELDA","175","",0,"","nserie",rstSeries,nserie,"nserie","descripcion","",""

                %>
                <input type="hidden" name="numSerieTienda" value="<%=enc.EncodeForHtmlAttribute(null_s(numSerieTienda))%>"/>
                <%

                rstSeries.close  
                rstSeriesAux.close  
                set rstSeries=nothing
                set rstSeriesAux=nothing
            '------------------------------------------------------------
			CloseFila
			DrawFila color_blau
				''if nproveedor>"" then nproveedor=Completar(nproveedor,5,"0")
                nptSELECT="select razon_social from proveedores with(nolock) where nproveedor=?"
                    'd_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente"))
				DrawCelda2 "CELDA style='width:145px'", "left", false, LitProveedor+": "%>
				<td class="CELDA7" style='width:310px'>
					<input class="CELDA7" type="text" name="nproveedor" value="<%=enc.EncodeForHtmlAttribute(trimCodEmpresa(nproveedor))%>" size="10" onchange="TraerProveedor('<%=enc.EncodeForJavascript(mode)%>','<%=enc.EncodeForJavascript(ndet)%>');"/>
					<a class="CELDAREFB" href="javascript:AbrirVentana('proveedores_busqueda.asp?ndoc=anular_cantidades_ptes&titulo=<%=LitSelProv%>&mode=search&viene=anular_cantidades_ptes','P',<%=altoventana%>,<%=anchoventana%>)"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a> 
					<input class="CELDA" type="text" name="nombre" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(nproveedor>"",DLookupP1(nptSELECT,nproveedor&"",adChar,10,session("dsn_cliente")),"")))%>" size="25"/>
				</td><%
			CloseFila
			DrawFila color_blau
				%><td class="CELDA" style='width:250px' colspan='2'>
					<font class="CELDA"><%=LitProveedorBaja +":  "%></font>
					<input type="checkbox" name="opcproveedorbaja" <%=enc.EncodeForHtmlAttribute(null_s(iif(opcproveedorbaja>"","checked","")))%>/>
				</td><%
			CloseFila
			if si_tiene_modulo_proyectos<>0 then
				DrawFila color_blau
					DrawCelda "CELDA style='width:145px'","","",0,iif(mode="browse","<b>"+LitProyecto+":</b>",LitProyecto+":")
					%><td class="CELDA">
						<input class="CELDA" type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(cod_proyecto)%>"/>
						<iframe id='frProyecto' name="fr_Proyecto" src='../mantenimiento/docproyectos.asp?viene=anular_cantidades_ptes&mode=<%=enc.EncodeForHtmlAttribute(mode)%>&cod_proyecto=<%=enc.EncodeForHtmlAttribute(cod_proyecto)%>' width='250' height='35' frameborder="no" scrolling="no" noresize="noresize"></iframe>
					</td><%
				CloseFila
			end if
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitPedPendPendConRef + ": "
				 %><td class="CELDA">
					<input class="CELDA" type="text" maxlength='25' size="25" name="referencia" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(referencia>"",trimCodEmpresa(referencia),"")))%>" size="12" onchange="traerreferencia();"/>
				</td><%
				DrawCelda2 "CELDA style='width:120px'", "left", false,""
				DrawCelda2 "CELDA", "left", false, LitPedPendPendConNom + ": "
				DrawInputCelda "CELDA","","",25,0,"","nombreart",nombreart
			CloseFila%>
		</table>

        <table style="margin-top:10px;" class="customWidth"><%
        'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
		    DrawFila color_blau
			    dim ConfigDespleg (3,13)

				i=0
				ConfigDespleg(i,0)="categoria"
				ConfigDespleg(i,1)="200"
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' and nombre <> 'Combustible' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="CELDA"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitCategoria
				ConfigDespleg(i,10)=categoria
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				i=1
				ConfigDespleg(i,0)="familia_padre"
				ConfigDespleg(i,1)="200"
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="CELDA"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitFamilia
				ConfigDespleg(i,10)=familia_padre
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				i=2
				ConfigDespleg(i,0)="familia"
				ConfigDespleg(i,1)="200"
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="CELDA"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitSubFamilia2
				ConfigDespleg(i,10)=familia
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				DibujaDesplegables ConfigDespleg,session("backendlistados")
		    CloseFila
        '------------------------------------------------------------
    %></table>
		<hr/>
		<table border="<%=borde%>" cellpadding="1" cellspacing="1"><%
			DrawFila color_fondo
				DrawCelda2 "'ENCABEZADOL sub-subsection'", "left", false, LitCamposOpcionales
			CloseFila%>
		</table>
		<table border="<%=borde%>" cellpadding="1" cellspacing="1"><%
			if si_tiene_modulo_proyectos<>0 then
				DrawFila color_blau
					DrawCelda2 "CELDA", "left", false, LitProyecto
					DrawCheckCelda "CELDA","","",0,"","opc_cod_proyecto",iif(opc_cod_proyecto=1,-1,0)
				CloseFila
			end if
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitFechaEntrega
				DrawCheckCelda "CELDA","","",0,"","opcfechaentrega",iif(opcfechaentrega=1,-1,0)
			CloseFila%>
		</table>
		<hr/>
		<table border='<%=borde%>' cellspacing="1" cellpadding="1"><%
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitPedidosAntig
			 	DrawInputCelda "CELDA","","",10,0,"","nDias",30
				DrawCelda2 "CELDA", "left", false, LitDias 
			CloseFila	
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitInsertarConceptos %>
			   <td class="CELDA" style='width:150px' colspan='2'><input type="checkbox" name="opcInsertaConceptos" <%=enc.EncodeForHtmlAttribute(null_s(iif(opcInsertaConceptos="1","checked",""))) %>></td> <%
			CloseFila
		    DrawFila color_blau
             'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
			    DrawCelda2 "CELDA valign='top' style='width:400px' ", "left", false, LitSoloAlmacen + ": "
                rstAux.cursorlocation=3        
                          
                numAlmacenTienda = ""

                If tiendaUsuario = "" Then 'El usuario no tiene tienda asignada                                
                    queryAlmacenes = "select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%' and fbaja is null and tienda is null order by descripcion "          
                Else 'El usuario tiene tienda asignada
                    queryAlmacenes = "select almacenes.codigo, almacenes.descripcion from almacenes with(nolock) inner join tiendas on almacenes.codigo = tiendas.almacen where tiendas.codigo like '" & tiendaUsuario & "'"
                End If

			    rstAux.open queryAlmacenes,session("dsn_cliente")

                If tiendaUsuario > "" Then 'Si tiene tienda asignada cargamos su almacen
                    numAlmacenTienda = rstAux("codigo")
                else 'Si no tiene tienda asignada cargamos todos los encontrados
                   dim listaAlmacenes
                   listaAlmacenes = ""

                   While not rstAux.eof
                        listaAlmacenes = listaAlmacenes & rstAux("codigo") & ","
                   rstAux.MoveNext
                   Wend

                   if listaAlmacenes > "" Then
                        numAlmacenTienda = left(listaAlmacenes, len(listaAlmacenes)-1)
                        rstAux.MoveFirst
                   End If
                End If

			    DrawSelectCelda "'CELDA' style='width:270px' align='left' size=5 multiple  colspan='2' ","","",0,"","almacen",rstAux,iif(TmpSoloAlmacen>"",TmpSoloAlmacen,""),"codigo","descripcion","",""
			                                  
                %>

                <input type="hidden" name="numAlmacenTienda" value="<%=enc.EncodeForHtmlAttribute(null_s(numAlmacenTienda))%>"/>

                <%

                rstAux.close
              '------------------------------------------------------------
		    CloseFila%>
		</table>		
	<%elseif mode="imp" then
		MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='428'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='428'", DSNIlion)
		%><input type='hidden' name='maxpdf' value='<%=enc.EncodeForHtmlAttribute(MAXPDF)%>'/>
		<input type='hidden' name='maxpagina' value='<%=enc.EncodeForHtmlAttribute(MAXPAGINA)%>'/>
		<input type="hidden" name="fdesde" value="<%=enc.EncodeForHtmlAttribute(fdesde)%>"/>
		<input type="hidden" name="fhasta" value="<%=enc.EncodeForHtmlAttribute(fhasta)%>"/>
		<input type="hidden" name="fentregadesde" value="<%=enc.EncodeForHtmlAttribute(fentregadesde)%>"/>
		<input type="hidden" name="fentregahasta" value="<%=enc.EncodeForHtmlAttribute(fentregahasta)%>"/>
		<input type="hidden" name="nserie" value="<%=enc.EncodeForHtmlAttribute(nserie)%>"/>
		<input type="hidden" name="nproveedor" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(viene="tienda",sesionNCliente&nproveedor,nproveedor)))%>"/>
		<input type="hidden" name="opcproveedorbaja" value="<%=enc.EncodeForHtmlAttribute(opcproveedorbaja)%>"/>
		<input type="hidden" name="opcElimPedSinDet" value="<%=enc.EncodeForHtmlAttribute(opcElimPedSinDet)%>"/>
		<input type="hidden" name="nDias" value="<%=enc.EncodeForHtmlAttribute(nDias)%>"/>
		<input type="hidden" name="actividad" value="<%=enc.EncodeForHtmlAttribute(actividad)%>"/>
		<input type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(cod_proyecto)%>"/>
		<input type="hidden" name="opc_cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(opc_cod_proyecto)%>"/>
        <input type="hidden" name="opcInsertaConceptos" value="<%=enc.EncodeForHtmlAttribute(opcInsertaConceptos)%>"/>
		<input type="hidden" name="opcfechaentrega" value="<%=enc.EncodeForHtmlAttribute(opcfechaentrega)%>"/>
		<% if viene="tienda" then%>
			<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
			<input type="hidden" name="campo" value="<%=enc.EncodeForHtmlAttribute(campo)%>"/>
			<input type="hidden" name="criterio" value="<%=enc.EncodeForHtmlAttribute(criterio)%>"/>
			<input type="hidden" name="texto" value="<%=enc.EncodeForHtmlAttribute(texto)%>"/>
		<%end if
		if viene="articulos" then%>
			<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
		<%end if%>
		<input type="hidden" name="referencia" value="<%=enc.EncodeForHtmlAttribute(referencia)%>"/>
		<input type="hidden" name="nombreart" value="<%=enc.EncodeForHtmlAttribute(nombreart)%>"/>
		<input type="hidden" name="familia" value="<%=enc.EncodeForHtmlAttribute(familia)%>"/>
		<input type="hidden" name="almacen" value="<%=enc.EncodeForHtmlAttribute(TmpSoloAlmacen)%>"/>
		
		<%if viene<>"tienda" then
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
            fntSELECT="select razon_social from proveedores with (nolock) where nproveedor = ?"
            'd_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente"))
			if viene<>"tienda" then
				%><font class="CELDA"><b><%=LitProveedor%>: </b><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(nproveedor)) & " - " & DLookupP1(fntSELECT, nproveedor&"",adChar,10,session("dsn_cliente"))%></font><br/><%
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

        'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
		if nserie > "" or numSerieTienda > "" then
            
            numSerie = ""
           
            if nserie > "" then
                numSerie = nserie
                    nsrSELECT = "select nombre from series with (nolock) where nserie = ?"
                    'd_lookup("nombre","series","nserie='" & nserie & "'",session("dsn_cliente"))
                %><font class='CELDA'><b><%=LitSerie%>:&nbsp;</b></font><font class='CELDA'><%=enc.EncodeForHtmlAttribute(null_s(trimCodEmpresa(nserie)))%>&nbsp;<%=DLookupP1(nsrSELECT,nserie&"",adVarchar,10,session("dsn_cliente"))%></font><br/><%
			    encabezado=1
            else
                numSerie = replace(replace(numSerieTienda," ",""),",","','")
            end if
			strwhere=strwhere & " serie in ('" & numSerie & "') and"			
		end if
        '---------------------------------------------------------

		if fechaInicio>"" and fechaHasta>"" then
			strwhere=strwhere & " fecha<='" & fechaInicio & "' and"			
		end if
		if fdesde > "" then 
			strwhere=strwhere & " fecha>='" & fdesde & "' and"
		end if	
		if fhasta > "" then 
		    strwhere=strwhere & " fecha<='" & fhasta & "' and"
		end if

		if fentregadesde > "" then
			strwhere=strwhere & " fecha_entrega>='" & fentregadesde & "' and"
			%><font class='CELDA'><b><%=LitDesdeFechaEntrega%>:&nbsp;</b></font><font class='CELDA'><%=fentregadesde%></font><br/><%
			encabezado=1
		end if
		if fentregahasta > "" then
			strwhere=strwhere & " fecha_entrega<='" & fentregahasta & "' and"
			%><font class='CELDA'><b><%=LitHastaFechaEntrega%>:&nbsp;</b></font><font class='CELDA'><%=fentregahasta%></font><br/><%
			encabezado=1
		end if

		if actividad > "" then
			strwhere=strwhere & " tactividad='" & actividad & "' and"
                actvdSELECT="select descripcion from tipo_actividad with (nolock) where codigo = ?"
                'd_lookup("descripcion","tipo_actividad","codigo='" & actividad & "'",session("dsn_cliente"))
			%><font class='CELDA'><b><%=LitActividad%>:&nbsp;</b></font><font class='CELDA'><%=DLookupP1(actvdSELECT,actividad&"",advarchar,10,session("dsn_cliente"))%></font><br/><%
			encabezado=1
		end if
		if cod_proyecto>"" then
			strwhere=strwhere & " cod_proyecto='" & cod_proyecto & "' and"
            cdSELECT="select nombre from proyectos with (nolock) where codigo = ?"
                'd_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",session("dsn_cliente"))
			%><font class='CELDA'><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=DLookupP1(cdSELECT,cod_proyecto&"",adVarchar,15,session("dsn_cliente"))%></font><br/><%
			encabezado=1
		end if

		'Tenemos una clausula where para los detalles y otra para los conceptos
		strwhereConceptos=strwhere

		if referencia>"" then
			if viene<>"articulos" then
				referencia=sesionNCliente & referencia
				strwhere=strwhere & " detalles_ped_pro.referencia like '%" & referencia & "%' and"
				strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"				
				%><font class='CELDA'><b><%=LitPedPendPendConRef%> : </b></font><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(referencia))%><br/><%
			else
				strwhere=strwhere & " detalles_ped_pro.referencia='" & referencia & "' and"
				strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"
				%><font class='CELDA'><b><%=LitReferencia%> : </b></font><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(referencia))%><br/><%
			end if

			encabezado=1
		end if
		if nombreart>"" and referencia & ""="" then
			strwhere=strwhere & " articulos.referencia=detalles_ped_pro.referencia and articulos.nombre like '%" & nombreart & "%' and"
			strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"			
			%><font class='CELDA'><b><%=LitPedPendPendConNom%> : </b></font><%=enc.EncodeForHtmlAttribute(nombreart)%><br/><%
			encabezado=1
		end if

        'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
		if familia>"" or refCategoriaCombustible>"" then
            if familia > "" then
                refFamilias = replace(replace(familia," ",""),",","','")
                strlistaFamilias = replace(replace(familia,"'"," "),session("ncliente"),"")	
                strlistaFamilias = NombresEntidades(familia,"familias","codigo","nombre",session("backendListados"))	
                strwhere=strwhere & " articulos.referencia=detalles_ped_pro.referencia and articulos.familia in ('" & refFamilias & "') and articulos.categoria <> '" & refCategoriaCombustible & "' and"
                strwhereConceptos=strwhereConceptos & " pedidos_pro.nproveedor='XXXX' and"			
			    %><font class='cab'><b><%=LitPedPendPendSubFamilia%> : </b></font><font class='cab'><%=enc.EncodeForHtmlAttribute(null_s(strlistaFamilias))%></font><br/><%
			    encabezado=1
            else
                strwhere=strwhere & " articulos.referencia=detalles_ped_pro.referencia and articulos.categoria <> '" & refCategoriaCombustible & "' and"
            end if				
		end if		
        
		if TmpSoloAlmacen>"" or numAlmacenTienda>"" then

            stralmacenes = ""

            if TmpSoloAlmacen > "" then
                stralmacenes = replace(replace(TmpSoloAlmacen," ",""),",","','")
                strlistaalmacenes = replace(replace(TmpSoloAlmacen,"'"," "),session("ncliente"),"")	
			    strlistaalmacenes=NombresEntidades(TmpSoloAlmacen,"almacenes","codigo","descripcion",session("backendListados"))		
			    %><font class='cab'><b><%=LitSoloAlmacen%> : </b></font><font class='cab'><%=strlistaalmacenes%></font><br/><%
			    encabezado=1
            else
                 stralmacenes = replace(replace(numAlmacenTienda," ",""),",","','")
            end if		   
            strwhere=strwhere & " detalles_ped_pro.almacen in ('" & stralmacenes & "') and"		
		end if		
        '-----------------------------------------------------------	

		strwhere=strwhere + " pedidos_pro.nproveedor=proveedores.nproveedor and nfactura is null and nalbaran is null"
		strwhere=strwhere + " and pedidos_pro.npedido=detalles_ped_pro.npedido"
		strwhere=strwhere + " and pedidos_pro.npedido like '" & sesionNCliente & "%' "

		strwhere=strwhere + " and cantidadpend<>0 "

		strwhereConceptos=strwhereConceptos + " pedidos_pro.nproveedor=proveedores.nproveedor and nfactura is null and nalbaran is null"
		strwhereConceptos=strwhereConceptos + " and pedidos_pro.npedido=conceptos_ped_pro.npedido"
		strwhereConceptos=strwhereConceptos + " and pedidos_pro.npedido like '" & sesionNCliente & "%' "

	strwhereAntes=strwhere

	strwhereConceptosAntes=strwhereConceptos + " order by pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"	

		strwhere2=strwhere + " and mainitem is null GROUP BY pedidos_pro.npedido, pedidos_pro.nproveedor, fecha"
		'cag strwhere2=strwhere2 + " ORDER BY pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"
		''cag2
		strwhereConceptos=strwhereConceptos + " GROUP BY pedidos_pro.npedido, pedidos_pro.nproveedor, fecha"
		strwhereConceptos=strwhereConceptos + " ORDER BY pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"
		strwhere=strwhere + " order by pedidos_pro.nproveedor, fecha, pedidos_pro.npedido"          
'Calculos de páginas--------------------------

		strselect="select count(pedidos_pro.npedido) as contador,pedidos_pro.npedido,pedidos_pro.nproveedor,fecha,'detalles' "
		strselect=strselect & " from pedidos_pro with(nolock),proveedores with(nolock),detalles_ped_pro with(nolock) "

		strselectConceptos="select count(pedidos_pro.npedido) as contador,pedidos_pro.npedido,pedidos_pro.nproveedor,fecha,'conceptos' "
		strselectConceptos=strselectConceptos & " from pedidos_pro with(nolock),proveedores with(nolock),conceptos_ped_pro with(nolock) "

        'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
		if (nombreart>"" and referencia & ""="") or familia>"" or refCategoriaCombustible>"" then
			strselect=strselect & ",articulos with(nolock) "
		end if
        '--------------------------------------------------------
		strselect=strselect & strwhere2

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
		'strselect="select item,detalles_ped_pro.descripcion,fecha,pedidos_pro.nproveedor,pedidos_pro.npedido,pedidos_pro.divisa,pedidos_pro.fecha_entrega,cantidad,detalles_ped_pro.referencia,almacenes.descripcion as almacen,detalles_ped_pro.importe,cod_proyecto,'detalle' as tipo,detalles_ped_pro.cantidadpend, detalles_ped_pro.npedidocli,detalles_ped_pro.itempedidocli "
		strselect="select item,detalles_ped_pro.descripcion,fecha,pedidos_pro.nproveedor,pedidos_pro.npedido,pedidos_pro.divisa,pedidos_pro.fecha_entrega,cantidad,detalles_ped_pro.referencia,almacenes.codigo as codAlmacen,almacenes.descripcion as almacen,detalles_ped_pro.importe,cod_proyecto,'detalle' as tipo,detalles_ped_pro.cantidadpend, detalles_ped_pro.npedidocli,detalles_ped_pro.itempedidocli "
		'fin cag
		strselect=strselect & " from pedidos_pro with(nolock),proveedores with(nolock),detalles_ped_pro  with(nolock) "
		strselect=strselect & "left outer join almacenes with(nolock) on almacenes.codigo=detalles_ped_pro.almacen "

		strselectConceptos="select nconcepto,convert(varchar,conceptos_ped_pro.descripcion),fecha,pedidos_pro.nproveedor,pedidos_pro.npedido,pedidos_pro.divisa,pedidos_pro.fecha_entrega,conceptos_ped_pro.cantidad,'' as referencia,'' as almacen,'',conceptos_ped_pro.importe,cod_proyecto,'concepto' as tipo,null,null,null"
		strselectConceptos=strselectConceptos & " from pedidos_pro with(nolock),proveedores with(nolock),conceptos_ped_pro with(nolock) "	

        'EVERIS (15/03/2017) Anular cantidades pendientes (ID105)
		if (nombreart>"" and referencia & ""="") or familia>"" or refCategoriaCombustible>"" then
			strselect=strselect & ",articulos with(nolock) "
		end if
        strselect=strselect & strwhereAntes & " and mainitem is null order by pedidos_pro.npedido " 

        
        '***************** Descomentar para ver la query
        ' DrawCelda2 "CELDA valign='top' style='width:400px' ", "left", false, strselect
        '*****************
        '-----------------------------------------------------

		rstPedido.cursorlocation=3
		rstPedido.Open strselect,session("dsn_cliente")
		if rstPedido.EOF then
			rstPedido.Close
			%><input type="hidden" name="NumRegsTotal" value="0"/>
			<script type="text/javascript" language="javascript">
					alert("<%=LitMsgDatosNoExiste%>");
			</script><%
            if viene<>"tienda" and viene<>"articulos" then
				%><script type="text/javascript" language="javascript">
		              //parent.botones.document.location = "anular_cantidades_ptes_bt.asp?mode=select1";
                      window.parent.botones.location = "anular_cantidades_ptes_bt.asp?mode=select1";
				      document.location = "anular_cantidades_ptes.asp?mode=select1";
				</script><%
			else
				%><script type="text/javascript" language="javascript">
				      //parent.botones.document.location="anular_cantidades_ptes_bt.asp?mode=imp&viene=<%=viene%>&ncliente=<%=ncliente%>";
					//document.location="anular_cantidades_ptes.asp?mode=imp&viene=<%=viene%>&ncliente=<%=ncliente%>";
				</script><%
			end if
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
if esto_no_sirve=1 then
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
end if
			'-----------------------------------------

			if lotes>1 and encabezado=1 then%>
				<hr/>
			<%end if%>
			<table id="dataTable" width="100%" border="0" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
			'Fila de encabezado
			nproveedor_ant="#@#"
			npedido_ant = "#@#"
			filab=0
			filac=1			
		'cag	Con este while controlo que no se pagine
			nregistros=1   'Para controlar el check de una fila en concreto
		while lote<=lotes
		'fin cag	
			while not rstPedido.EOF and filab<lista(lote,2) 'MAXPAGINA
				CheckCadenaTienda sesionNCliente,rstPedido("nproveedor")

						if rstPedido("nproveedor")<>nproveedor_ant and rstPedido("npedido")<>npedido_ant then
							DrawFila color_blau
								DrawCeldaSpan "'ENCABEZADOL customBackground'","","",0,"<hr/>",11
							CloseFila
							DrawFila color_blau
								if viene<>"tienda" then
									dat1=Hiperv(OBJProveedores,rstPedido("nproveedor"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("nproveedor")) & " - " & d_lookup("razon_social","proveedores","nproveedor='" & rstPedido("nproveedor") & "'",session("dsn_cliente")),LitVerProveedor)
								else
									dat1=trimCodEmpresa(rstPedido("nproveedor")) & " - " & d_lookup("razon_social","proveedores","nproveedor='" & rstPedido("nproveedor") & "'",session("dsn_cliente"))
								end if
								DrawCeldaSpan "'ENCABEZADOL underOrange NO_BORDER_H padding-top20'","","",0,LitProveedor & " : " & dat1,8
							CloseFila
							nproveedor_ant = rstPedido("nproveedor")
						end if
						    ''cag 2
							if rstPedido("npedido")<>npedido_ant and rstPedido("tipo")="detalle" then
							'if rstPedido("npedido")<>npedido_ant then
								DrawFila color_terra
									DrawCelda "'ENCABEZADOL7 customBackground'","5%","",0," "
									if viene<>"tienda" then
										dat1=Hiperv(OBJPedidosPro,rstPedido("npedido"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("npedido")),LitVerPedido)
									else
										dat1="<a class='CELDAREFB' href=javascript:ver_pedido('" & rstPedido("npedido") & "','"&nproveedor&"') alt='" & LitVerPedido & "'>" & trimCodEmpresa(rstPedido("npedido")) & "</a>"
									end if
									ndecimales=d_lookup("ndecimales", "divisas", "codigo like '"&session("ncliente")&"%' and codigo ='"&d_lookup("divisa","pedidos_pro","npedido='" & rstPedido("npedido") & "'",session("dsn_cliente"))&"'", session("dsn_cliente"))
									dat2=" - " + LitFecha+": "+cstr(rstPedido("fecha"))+"-"+LitImporte+": "+cstr(formatnumber(d_lookup("total_pedido","pedidos_pro","npedido='" & rstPedido("npedido") & "'",session("dsn_cliente")), ndecimales,-1,0,-1))+" "+d_lookup("abreviatura","divisas","codigo='" & rstPedido("divisa") & "'",session("dsn_cliente"))
									if opc_cod_proyecto="1" then
										dat2=dat2 & " - " & LitProyecto & " : " & d_lookup("nombre","proyectos","codigo='" & rstPedido("cod_proyecto") & "'",session("dsn_cliente"))
									end if
									if opcfechaentrega="1" then									
										'DrawCeldaSpan "ENCABEZADOL","","",0,LitPedidoMin + " : " & dat1 & dat2,10
										DrawCeldaSpan "'ENCABEZADOL customBackground padding-top10'","","",0,LitPedidoMin + " : " & dat1 & dat2,11
									else
										'DrawCeldaSpan "ENCABEZADOL","","",0,LitPedidoMin + " : " & dat1 & dat2,9
										DrawCeldaSpan "'ENCABEZADOL customBackground padding-top10'","","",0,LitPedidoMin + " : " & dat1 & dat2,10
									end if
								CloseFila
								npedido_ant = rstPedido("npedido")
								'cag 2
								if rstPedido("tipo")="detalle" then		
								'fin cag2
								DrawFila color_terra
									'DrawCeldaSpan "ENCABEZADOL","","",0,"",2
									DrawCelda "'ENCABEZADOL7 underOrange'","5%","",0," "
									DrawCelda "'ENCABEZADOL7 underOrange'","5%","",0," "
									'cag
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitEliminar
									'fin cag
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitCantidad
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitReferencia
									'DrawCelda "ENCABEZADOL7","","",0,LitDescripcionConcepto
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitDescripcion
									if opcfechaentrega="1" then
										DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitFechaEntrega
									end if
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitAlmacen
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitPendiente
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitPedidoCli
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitItemPed
									DrawCelda "'ENCABEZADOL7 underOrange'","","",0,LitImporte
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
			   						 <td class="ENCABEZADOL7" width="5%" align="center"> <%'=nregistros%>
									   	<input type="checkbox" name='checkElim<%=nregistros%>'/> </td>
										<input type="hidden" name="nPedido<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("npedido")))%>"/>
										<input type="hidden" name="refer<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia")))%>"/>
										<input type="hidden" name="ctdad<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("cantidad")))%>"/>
										<input type="hidden" name="ctdadPend<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("cantidadpend")))%>"/>
										<input type="hidden" name="almacen<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("almacen")))%>"/>
										<input type="hidden" name="codAlmacen<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("codAlmacen")))%>"/>
										<input type="hidden" name="nProveedor<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor")))%>"/>
										<input type="hidden" name="item<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("item")))%>"/>
										 <%
									'fin cag
									DrawCelda "CELDAR7","","",0,formatnumber(rstPedido("cantidad"),DEC_CANT,-1,0,-1)
									if viene<>"tienda" then
										dat1=Hiperv(OBJArticulos,rstPedido("referencia"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("referencia")),LitVerArticulo)
										dat3=Hiperv(OBJPedidosCli,rstPedido("npedidocli"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("npedidocli")),LitVerPedido)
									else
										dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & rstPedido("referencia") & "','" & rstPedido("nproveedor") & "') alt='" & LitVerArticulo & "' border='0'>" & trimCodEmpresa(rstPedido("referencia")) & "</a>"
										dat3="<a class='CELDAREFB' href=javascript:ver_pedidoCli('" & rstPedido("npedidocli") & "') alt='" & LitVerPedido & "'>" & trimCodEmpresa(rstPedido("npedidocli")) & "</a>"
									end if
									DrawCelda "CELDAL7","","",0,dat1
									dat1=rstPedido("descripcion")
									DrawCelda "CELDAL7","","",0,dat1
									if opcfechaentrega="1" then
										DrawCelda "CELDAL7","","",0,rstPedido("fecha_entrega")
									end if
									DrawCelda "CELDAL7","","",0,rstPedido("almacen")
									DrawCelda "CELDAR7","","",0,formatnumber(rstPedido("cantidadpend"),DEC_CANT,-1,0,-1)
									'DrawCelda "CELDAL7","","",0,rstPedido("npedidocli")
									DrawCelda "CELDAL7","","",0,dat3
									DrawCelda "CELDAL7","","",0,iif(rstPedido("npedidocli")>"",rstPedido("itempedidocli"),"")
									DrawCelda "CELDAR7","","",0,cstr(formatnumber(null_z(rstPedido("Importe")),ndecimales,-1,0,-1))+" "+d_lookup("abreviatura","divisas","codigo='" & rstPedido("divisa") & "'",session("dsn_cliente"))
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
					if rstPedido("npedido")=npedido_ant and rstPedido("nproveedor")=nproveedor_ant then
						'cag 2
						if rstPedido("tipo")="detalle" then		
						'fin cag2
						DrawFila color_blau
							DrawCelda "ENCABEZADOL7","5%","",0," "
							DrawCelda "ENCABEZADOL7","5%","",0," " %>
							<!-- cag -->
	   						<td class="ENCABEZADOL7" width="5%" align="center"> <%'=nregistros%>
							    <input type="checkbox" name='checkElim<%=nregistros%>' /></td>
								<input type="hidden" name="nPedido<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("npedido")))%>"/>
								<input type="hidden" name="refer<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("referencia")))%>"/>
								<input type="hidden" name="ctdad<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("cantidad")))%>"/>
								<input type="hidden" name="ctdadPend<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("cantidadpend")))%>"/>
								<input type="hidden" name="almacen<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("almacen")))%>"/>
								<input type="hidden" name="codAlmacen<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("codAlmacen")))%>"/>
								<input type="hidden" name="nProveedor<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("nproveedor")))%>"/>
								<input type="hidden" name="item<%=nregistros%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstPedido("item")))%>"/>
							   <%
							'fin cag
							DrawCelda "CELDAR7","","",0,formatnumber(rstPedido("cantidad"),DEC_CANT,-1,0,-1)
							if viene<>"tienda" then
								dat1=Hiperv(OBJArticulos,rstPedido("referencia"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,trimCodEmpresa(rstPedido("referencia")),LitVerArticulo)
							else
								dat1="<a class='CELDAREFB' href=javascript:Ver_Articulo('" & rstPedido("referencia") & "','" & rstPedido("nproveedor") & "') alt='" & LitVerArticulo & "' border='0'>" & rstPedido("referencia") & "</a>"
							end if
							DrawCelda "CELDAL7","","",0,dat1
							dat1=rstPedido("descripcion")
							DrawCelda "CELDAL7","","",0,dat1
							if opcfechaentrega="1" then
								DrawCelda "CELDAL7","","",0,rstPedido("fecha_entrega")
							end if
							DrawCelda "CELDAL7","","",0,rstPedido("almacen")
							DrawCelda "CELDAR7","","",0,formatnumber(rstPedido("cantidadpend"),DEC_CANT,-1,0,-1)
							DrawCelda "CELDAL7","","",0,rstPedido("npedidocli")
							DrawCelda "CELDAL7","","",0,iif(rstPedido("npedidocli")>"",rstPedido("itempedidocli"),"")
							ndecimales=d_lookup("ndecimales", "divisas", "codigo like '"&session("ncliente")&"%' and codigo ='"&d_lookup("divisa","pedidos_pro","npedido='" & rstPedido("npedido") & "'",session("dsn_cliente"))&"'", session("dsn_cliente"))
							DrawCelda "CELDAR7","","",0,cstr(formatnumber(rstPedido("Importe"),ndecimales,-1,0,-1))+" "+d_lookup("abreviatura","divisas","codigo='" & rstPedido("divisa") & "'",session("dsn_cliente"))
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

				%><input type="hidden" name="npedido_ant" value="<%=enc.EncodeForHtmlAttribute(null_s(npedido_ant))%>"/>
				<input type="hidden" name="nproveedor_ant2" value="<%=enc.EncodeForHtmlAttribute(null_s(nproveedor_ant))%>"/><%
			end if
		'cag	
			lote=lote+1
        wend  ' del while lote<=lotes
           %><input type="hidden" name="h_nregistros" value="<%=nregistros-1%>"/><%
		   
		'fin cag
			%></table>
            <hr/><%
			rstPedido.Close
		end if
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
		if lista=")" then
			%><script type="text/javascript" language="javascript">
			      window.alert("<%=LitMsgDocumentoNoSel%>");
			      parent.botones.document.location = "anular_cantidades_ptes_bt.asp?mode=select1";
				document.location="anular_cantidades_ptes.asp?mode=select1";
			</script><%
	    end if
		
		
		'Creo temporal de usuario para realizar anulación masiva en el procedimiento
		strdrop ="if exists (select * from sysobjects where id = object_id('egesticet.[" & session("usuario") & "]') ) drop table egesticet.[" & session("usuario") & "]"		
		rst.open strdrop,session("dsn_cliente")
		
		strselect="create table [" & session("usuario") & "] (referencia varchar(30),numPedido varchar(20), ctdad real, ctdadPend real, almacen varchar(10), nProveedor varchar(10), item smallint)"
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
			   nProveedor="nProveedor" & i
			   nItem="item" & i
			   strselect="insert into [" & session("usuario") & "] (referencia, numPedido, ctdad, ctdadPend,almacen,nProveedor,item) "
			   strselect= strselect & " values ('" & request.form(referencia) & "','" & request.form(numPedido) & "'," & replace(request.form(cantidad), ",", ".") & "," & replace(request.form(cantidadPdte), ",", ".") & ",'" & request.form(codAlmacen) & "','" & request.form(nProveedor) & "','" & request.form(nItem)& "')"
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
		Command.CommandText="AnularCantidadesPendientes"
		Command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		Command.Parameters.Append Command.CreateParameter("@opcusuario", adVarChar, adParamInput,50, session("usuario"))		
		Command.Parameters.Append Command.CreateParameter("@sesion_ncliente", adVarChar, adParamInput,5, session("ncliente"))
		Command.Parameters.Append Command.CreateParameter("@nameUsuario", adVarChar, adParamInput,25, nomusuario)
		Command.Parameters.Append Command.CreateParameter("@ipUsuario", adVarChar, adParamInput,255, Request.ServerVariables("REMOTE_ADDR"))
        Command.Parameters.Append Command.CreateParameter("@opcInsertaConceptos", adVarChar, adParamInput,5, iif(opcInsertaConceptos="1",1,0))
        Command.Parameters.Append Command.CreateParameter("@resul", adVarChar, adParamOutput,2, Resultado)
        Command.Execute,,adExecuteNoRecords

		Resultado = Command.Parameters("@resul").Value
		conn.close
		set command=nothing
		set conn=nothing
	     if Resultado="0" then%>
		  <script type="text/javascript" language="javascript">
		      window.alert("<%=LitProcesoFinCorrecto%>")
		      parent.botones.document.location = "anular_cantidades_ptes_bt.asp?mode=select1";
			 document.location="anular_cantidades_ptes.asp?mode=select1";
		  </script>
		<%end if
		'fin coger excepcion y parametro salida
	
	'fin cag
	end if%>
</form>
<%
set rstSelect=nothing
set rstAux=nothing
set rst=nothing
set rstPedido=nothing
set rstPedido2=nothing
set rstPedido3=nothing
set conn2=nothing

end if%>
</body>
</html>