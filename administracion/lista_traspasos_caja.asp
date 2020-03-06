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
<title><%=LitTituloLTC%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>

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
<!--#include file="Ahoja_gastos.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file= "../CatFamSubResponsive.inc"-->
<!--#include file="../styles/formularios.css.inc" -->  
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" --> 

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<body onload="self.status='';" class="BODY_ASP">
<%
'*******************************************************************************'
function CalculaPagina(ndoc,tdoc,ndoc_pro)
	select case tdoc
		case "ALBARAN DE PROVEEDOR"
			CalculaPagina=Hiperv(OBJAlbaranesPro,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),ndoc_pro,LitVerAlbaran)
		case "FACTURA DE PROVEEDOR"
			CalculaPagina=Hiperv(OBJFacturasPro,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),ndoc_pro,LitVerFactura)
		case "VENCIMIENTO_ENTRADA"
			nfac1=d_lookup("nfactura","vencimientos_entrada","nfactura+'-'+cast(nvencimiento as varchar(10))='" & ndoc & "'",session("dsn_cliente"))
			nfac=d_lookup("nfactura_pro","facturas_pro","nfactura='" & nfac1 & "'",session("dsn_cliente"))
			CalculaPagina=Hiperv(OBJFacturasPro,nfac1,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),ndoc_pro,LitVerVencimiento)
		case "PEDIDO A PROVEEDOR"
			CalculaPagina=Hiperv(OBJPedidosPro,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(ndoc),LitVerPedido)
		case "ALBARAN DE SALIDA"
			CalculaPagina=Hiperv(OBJAlbaranesCli,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(ndoc),LitVerAlbaran)
		case "FACTURA A CLIENTE"
			CalculaPagina=Hiperv(OBJFacturasCli,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(ndoc),LitVerFactura)
		case "VENCIMIENTO_SALIDA"
			nfac=d_lookup("nfactura","vencimientos_salida","nrecibo='" & ndoc & "'",session("dsn_cliente"))
			CalculaPagina=Hiperv(OBJFacturasCli,nfac,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(ndoc),LitVerVencimiento)
		case "PEDIDO DE CLIENTE"
			CalculaPagina=Hiperv(OBJPedidosCli,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(ndoc),LitVerPedido)
		case "HOJA DE GASTOS"
			CalculaPagina=Hiperv(OBJHojaGastos,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(ndoc),LitVerHojaGasto)
		case "TICKET"
		    CalculaPagina = trimCodEmpresa(ndoc)
		case "ORDEN"
			CalculaPagina=Hiperv(OBJOrdenes,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(ndoc),LitVerOrden)
		case else
			CalculaPagina=""
	end select
end function

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************

const borde=0
%>
	<form name="lista_traspasos_caja" method="post">
	<%PintarCabecera "lista_traspasos_caja.asp"

	'Leer parámetros de la página
	mode = EncodeForHtml(Request.QueryString("mode"))
	if ucase(mode) = "BROWSE" then mode ="imp"

	if request.querystring("Dfecha")>"" then
		TmpDfecha=limpiaCadena(request.querystring("Dfecha"))
	else
		TmpDfecha=limpiaCadena(request.Form("Dfecha"))
	end if

	if request.querystring("Hfecha")>"" then
		TmpHfecha=limpiaCadena(request.querystring("Hfecha"))
	else
		TmpHfecha=limpiaCadena(request.Form("Hfecha"))
	end if

	if request.querystring("serie")>"" then
		TmpSerie=limpiaCadena(request.querystring("serie"))
	else
		TmpSerie=limpiaCadena(request.form("serie"))
	end if

	if request.querystring("medio")>"" then
		TmpMedio=limpiaCadena(request.querystring("medio"))
	else
		TmpMedio=limpiaCadena(request.Form("medio"))
	end if

	if request.querystring("descripcion")>"" then
		TmpDescripcion=limpiaCadena(request.querystring("descripcion"))
	else
		TmpDescripcion=limpiaCadena(request.Form("descripcion"))
	end if

	if request.querystring("responsable")>"" then
		TmpResponsable=limpiaCadena(request.querystring("responsable"))
	else
		TmpResponsable=limpiaCadena(request.Form("responsable"))
	end if
	TmpResponsable2=TmpResponsable

	if request.querystring("cajaorigen")>"" then
		TmpCajaOrigen=limpiaCadena(request.querystring("cajaorigen"))
	else
		TmpCajaOrigen=limpiaCadena(request.Form("cajaorigen"))
	end if

	if request.querystring("cajadestino")>"" then
		TmpCajaDestino=limpiaCadena(request.querystring("cajadestino"))
	else
		TmpCajaDestino=limpiaCadena(request.Form("cajadestino"))
	end if

	TmpMuestraDetalles=iif(limpiaCadena(request.form("mostrardetalles"))>"","SI","")

	if request.querystring("ordenarpor")>"" then
		TmpOrdenarPor=limpiaCadena(request.querystring("ordenarpor"))
	else
		TmpOrdenarPor=limpiaCadena(request.Form("ordenarpor"))
	end if

	strwhere=""%>

		<table width='100%' cellspacing="1" cellpadding="1">
   			<tr>
				<%if mode="imp" then%>
					<td class='CELDA7' bgcolor="">
						<%fdesde=EncodeForHtml(TmpDfecha)
						fhasta=TmpHfecha
						fhasta=EncodeForHtml(day(fhasta) & "/" & month(fhasta) & "/" & year(fhasta))
						if fdesde>"" then
							if fhasta>"" then
								%><%=LitPeriodoFechas%> : <b><%=fdesde%> - <%=fhasta%></b><%
							else
								%><%=LitPeriodoFechas%> : <b><%=LitDesde%>&nbsp;<%=fdesde%></b><%
							end if
						else
							if fhasta>"" then
								%><%=LitPeriodoFechas%> : <b><%=LitHasta%>&nbsp;<%=fhasta%></b><%
							else
							end if
						end if%>
					</td><%
				else%>
					<td></td><%
				end if%>
	   		</tr>
		</table>
		<hr/>
		<% Alarma "lista_traspasos_caja.asp"

  		set conn = Server.CreateObject("ADODB.Connection")
  		conn.open session("dsn_cliente")

		set rstAux = Server.CreateObject("ADODB.Recordset")
		set rst = Server.CreateObject("ADODB.Recordset")
		set rstSelect = Server.CreateObject("ADODB.Recordset")
		set rstAux1 = Server.CreateObject("ADODB.Recordset")
		set rstTraspaso = Server.CreateObject("ADODB.Recordset")

		if mode="param" then
			
                    DrawDiv "1","",""
                    DrawLabel "","",LitDesdeFecha
                    DrawInput "", "", "Dfecha", EncodeForHtml(iif(TmpDfecha>"",TmpDfecha,"01/01/" & year(date))), ""
                    DrawCalendar "Dfecha"
                    CloseDiv
					
                    DrawDiv "1","",""
                    DrawLabel "","",LitHastaFecha    
                    DrawInput "", "", "Hfecha",EncodeForHtml(iif(TmpHfecha>"",TmpHfecha,day(date) & "/" & month(date) & "/" & year(date))), ""
                    DrawCalendar "Hfecha"
                    CloseDiv
				
					rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='TRASPASO ENTRE CAJAS' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					DrawSelectCelda "","","",0,LitSerie,"serie",rstAux,TmpSerie,"nserie","descripcion","",""
					rstAux.close
				
					rstAux.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					DrawSelectCelda "","","",0,LitMedio,"medio",rstAux,TmpMedio,"codigo","descripcion","",""
					rstAux.close
				
					rstSelect.open "select dni, nombre from personal with(nolock) where fbaja is null and dni like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					DrawSelectCelda "","","",0,LitResponsable,"responsable",rstSelect,TmpResponsable,"dni","nombre","",""
					rstSelect.close
				
					rstSelect.open "select codigo, descripcion from cajas with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					DrawSelectCelda "","","",0,LitCajaOrg,"cajaorigen",rstSelect,TmpCajaOrigen,"codigo","descripcion","",""
					rstSelect.close
					
					rstSelect.open "select codigo, descripcion from cajas with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					DrawSelectCelda "","","",0,LitCajaDes,"cajadestino",rstSelect,TmpCajaDestino,"codigo","descripcion","",""
					rstSelect.close
				
                    EligeCelda "input","add","left","","",0,LitDescripcion2,"descripcion",0,EncodeForHtml(TmpDescripcion)
				
                DrawDiv "1", "", ""
                DrawLabel "", "", LitOrdenarPor%><select class="width60" name="ordenarpor">
							<option selected="selected" value="<%=LitFecha%>"><%=LitFecha%></option>
							<option value="<%=LitResponsable%>"><%=LitResponsable%></option>
						</select>					
				<%
                CloseDiv				
                    EligeCelda "check","add","","","",0,LitMostrarDetalles,"mostrardetalles",8,""
				%>			
			<hr/>
		<%elseif mode="imp" then%>
			<input type="hidden" name="Dfecha" value="<%=EncodeForHtml(TmpDfecha)%>">
			<input type="hidden" name="Hfecha" value="<%=EncodeForHtml(TmpHfecha)%>">
			<input type="hidden" name="medio" value="<%=EncodeForHtml(TmpMedio)%>">
			<input type="hidden" name="serie" value="<%=EncodeForHtml(TmpSerie)%>">
			<input type="hidden" name="responsable" value="<%=EncodeForHtml(TmpResponsable)%>">
			<input type="hidden" name="cajaorigen" value="<%=EncodeForHtml(TmpCajaOrigen)%>">
			<input type="hidden" name="cajadestino" value="<%=EncodeForHtml(TmpCajaDestino)%>">
			<input type="hidden" name="descripcion" value="<%=EncodeForHtml(TmpDescripcion)%>">
			<input type="hidden" name="ordenarpor" value="<%=EncodeForHtml(TmpOrdenarPor)%>">
			<input type="hidden" name="mostrardetalles" value="<%=EncodeForHtml(mostrardetalles)%>">
			<input type="hidden" name="TmpMuestraDetalles" value="<%=EncodeForHtml(TmpMuestraDetalles)%>">

			<%MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='148'", DSNIlion)
			MAXPDF=d_lookup("maxpdf", "limites_listados", "item='148'", DSNIlion)%>
			<input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>
			<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'>
			<input type='hidden' name='maxmb' value='<%=EncodeForHtml(MB)%>'>

			<%strwhere=" where"
			encabezado=0

			if TmpMedio > "" then%>
				<font class=cab><b><%=LitMedio%>:&nbsp;</b></font><font class=cab><%=EncodeForHtml(trimCodEmpresa(TmpMedio) & " - " & d_lookup("descripcion","tipo_pago","codigo='" & TmpMedio & "'",session("dsn_cliente")))%></font><br/>
			<%end if

			if TmpSerie > "" then%>
				<font class=cab><b><%=LitSerie%>:&nbsp;</b></font><font class=cab><%=EncodeForHtml(trimCodEmpresa(TmpSerie) & " - " & d_lookup("nombre","series","nserie='" & TmpSerie & "'",session("dsn_cliente")))%></font><br/>
			<%end if

			if TmpDescripcion > "" then%>
				<font class=cab><b><%=LitDescripcion2%>:&nbsp;</b></font><font class=cab><%=EncodeForHtml(TmpDescripcion)%></font><br/>
			<%end if
			if TmpResponsable > "" then%>
				<font class=cab><b><%=LitResponsable%>:&nbsp;</b></font><font class=cab><%=EncodeForHtml(trimCodEmpresa(TmpResponsable) & " - " & d_lookup("nombre","personal","dni='" & TmpResponsable & "'",session("dsn_cliente")))%></font><br/>
			<%end if
			if TmpCajaOrigen > "" then%>
				<font class=cab><b><%=LitCajaOrg%>:&nbsp;</b></font><font class=cab><%=EncodeForHtml(trimCodEmpresa(TmpCajaOrigen) & " - " & d_lookup("descripcion","cajas","codigo='" & TmpCajaOrigen & "'",session("dsn_cliente")))%></font><br/>
			<%end if
			if TmpCajaDestino > "" then%>
				<font class=cab><b><%=LitCajaDes%>:&nbsp;</b></font><font class=cab><%=EncodeForHtml(trimCodEmpresa(TmpCajaDestino) & " - " & d_lookup("descripcion","cajas","codigo='" & TmpCajaDestino & "'",session("dsn_cliente")))%></font><br/>
			<%end if
			
			set rstAux1 = conn.execute("EXEC ListaTraspasos @NomTabla='" & session("usuario") & "' ,@FechaDesde='" & TmpDfecha & "' ,@FechaHasta='" & TmpHfecha & "' ,@Serie='" & TmpSerie & "' ,@CajaOrigen='" & TmpCajaOrigen & "' ,@CajaDestino='" & TmpCajaDestino & "',@MedioPago='" & TmpMedio & "',@Responsable='" & TmpResponsable & "',@Descripcion='" & TmpDescripcion & "',@MostrarDetalles='" & TmpMuestraDetalles & "',@ordenarPor='" & TmpOrdenarPor & "',@session_ncliente='" & session("ncliente") & "'")

		rstTraspaso.cursorlocation=3
		rstTraspaso.open "select * from [" & session("usuario") & "]",session("dsn_cliente")

		%><input type="hidden" name="NumRegs" value="<%=EncodeForHtml(rstTraspaso.Recordcount)%>"><%

		if rstTraspaso.EOF then
			rstTraspaso.Close
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitMsgDatosNoExiste%>");
			      parent.window.frames["botones"].document.location = "lista_traspasos_caja_bt.asp?mode=param";
				document.location="lista_traspasos_caja.asp?mode=param";
			</script><%
		else
			'Calculos de páginas--------------------------'
		   lote=limpiaCadena(Request.QueryString("lote"))
		   if lote="" then
			lote=1
		   end if
		   sentido=limpiaCadena(Request.QueryString("sentido"))

		   lotes=rstTraspaso.RecordCount/MAXPAGINA
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

		   rstTraspaso.PageSize=MAXPAGINA
		   rstTraspaso.AbsolutePage=lote
		  '-----------------------------------------
		%><hr/><%
		NavPaginas lote,lotes,campo,criterio,texto,1


		'PERMISOS DE HIPERVINCULOS ----------------------------------
		VinculosPagina(MostrarTraspasosCaja)=1:VinculosPagina(MostrarPersonal)=1
		VinculosPagina(MostrarProveedores)=1:VinculosPagina(MostrarClientes)=1
		VinculosPagina(MostrarAlbaranespro)=1:VinculosPagina(MostrarAlbaranesCli)=1
		VinculosPagina(MostrarFacturaspro)=1:VinculosPagina(MostrarFacturasCli)=1
		VinculosPagina(MostrarPedidosPro)=1:VinculosPagina(MostrarPedidosCli)=1
		VinculosPagina(MostrarHojaGastos)=1:VinculosPagina(MostrarOrdenes)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
		'---------------------------------


ndecimales=d_lookup("ndecimales","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
abreviatura=d_lookup("abreviatura","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))

			traspasoAnt=""
			extraAnt="X"%>
				<table width='100%' style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
			    <%'Fila de encabezado'
				DrawFila color_fondo
					DrawCelda "ENCABEZADOC width=6%","","",0,LitFecha
					DrawCelda "ENCABEZADOC width=6%","","",0,LitNTraspaso
					DrawCelda "ENCABEZADOC width=16%","","",0,LitDescripcion
					DrawCelda "ENCABEZADOC width=16%","","",0,LitResponsable
					DrawCelda "ENCABEZADOC width=13%","","",0,LitCajaOrg
					DrawCelda "ENCABEZADOC width=13%","","",0,LitCajaDes
					DrawCelda "ENCABEZADOC width=10%","","",0,LitImporte
					DrawCelda "ENCABEZADOC width=20%","","",0,LitMedio
				CloseFila

			totalImporte=0
			totalTraspasos=0
			fila=1

			while not rstTraspaso.EOF and fila<=MAXPAGINA
				CheckCadena rstTraspaso("ntraspaso")

				'***** CASO 1: NO MOSTRAR DETALLES  ********'

				if TmpMuestraDetalles="" then
					'Seleccionar el color de la fila.'
					if ((fila+1) mod 2)=0 then
						color=color_blau
					else
						color=color_terra
					end if
					DrawFila color
						DrawCelda "TDBORDECELDA7 width=6%","","",0,EncodeForHtml(rstTraspaso("fecha"))
						doc=rstTraspaso("ntraspaso")
						url=Hiperv(OBJTraspasosCaja,rstTraspaso("ntraspaso"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(doc),LitVerTraspaso)
				    	 	%><td class=TDBORDECELDA7 width="6%"><%=EncodeForHtml(url)%></td><%
						DrawCelda "TDBORDECELDA7 width=16%","","",0,EncodeForHtml(rstTraspaso("descripcion"))
						doc=rstTraspaso("nomresponsable")
						url=Hiperv(OBJPersonal,rstTraspaso("responsable"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),doc,LitVerPersonal)
				    	 	%><td class=TDBORDECELDA7 width="16%"><%=EncodeForHtml(url)%></td><%
						DrawCelda "TDBORDECELDA7 width=13%","","",0,EncodeForHtml(trimCodEmpresa(rstTraspaso("cajaorg"))) & " - " & EncodeForHtml(rstTraspaso("desccajaorg"))
						DrawCelda "TDBORDECELDA7 width=13%","","",0,EncodeForHtml(trimCodEmpresa(rstTraspaso("cajadest"))) & " - " & EncodeForHtml(rstTraspaso("desccajadest"))
						DrawCelda "TDBORDECELDA7 width=10% align='right'","","",0,EncodeForHtml(cstr(formatnumber(rstTraspaso("importe"),rstTraspaso("ndecimales"),-1,0,-1))) & " " & EncodeForHtml(rstTraspaso("abreviatura"))
						DrawCelda "TDBORDECELDA7 width=20%","","",0,EncodeForHtml(rstTraspaso("medio"))
					CloseFila
					totalImporte=totalImporte+rstTraspaso("importe")
					totalTraspasos=totalTraspasos+1
					fila=fila+1
				end if

				'********* CASO 2: MOSTRAR DETALLES **********'
				if TmpMuestraDetalles="SI" then
					if rstTraspaso("ntraspaso")<>traspasoAnt then
						DrawFila color_blau
							DrawCelda "TDBORDECELDA7 width=6%","","",0,EncodeForHtml(rstTraspaso("fecha"))
							doc=rstTraspaso("ntraspaso")
							url=Hiperv(OBJTraspasosCaja,rstTraspaso("ntraspaso"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(doc),LitVerTraspaso)
					    	 	%><td class=TDBORDECELDA7 width="6%"><%=EncodeForHtml(url)%></td><%
							DrawCelda "TDBORDECELDA7 width=16%","","",0,EncodeForHtml(rstTraspaso("descripcion"))
							doc=rstTraspaso("nomresponsable")
							url=Hiperv(OBJPersonal,rstTraspaso("responsable"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),doc,LitVerPersonal)
					    	 	%><td class=TDBORDECELDA7 width="16%"><%=EncodeForHtml(url)%></td><%
							DrawCelda "TDBORDECELDA7 width=13%","","",0,EncodeForHtml(trimCodEmpresa(rstTraspaso("cajaorg"))) & " - " & EncodeForHtml(rstTraspaso("desccajaorg"))
							DrawCelda "TDBORDECELDA7 width=13%","","",0,EncodeForHtml(trimCodEmpresa(rstTraspaso("cajadest"))) & " - " & EncodeForHtml(rstTraspaso("desccajadest"))
							DrawCelda "TDBORDECELDA7 width=10% align='right'","","",0,EncodeForHtml(cstr(formatnumber(rstTraspaso("importe"),rstTraspaso("ndecimales"),-1,0,-1))) & " " & EncodeForHtml(rstTraspaso("abreviatura"))
							DrawCelda "TDBORDECELDA7 width=20%","","",0,EncodeForHtml(rstTraspaso("medio"))
						CloseFila
						totalImporte=totalImporte+rstTraspaso("importe")
						totalTraspasos=totalTraspasos+1

						if rstTraspaso("fechadet")>"" then%>
						        <%DrawFila color_blau
									DrawCelda "tdbordeCELDA7 bgcolor=" & color_terra & "  width=2%","","",0,""
									DrawCelda "tdbordeCELDA7 bgcolor=" & color_terra & "  width=10%","","",0,"<b>" & LitFecha & "</b>"
									DrawCelda "tdbordeCELDA7 colspan=2 bgcolor=" & color_terra & "  width=32%","","",0,"<b>" & LitDescripcion & "</b>"
									DrawCelda "tdbordeCELDA7 bgcolor=" & color_terra & "  width=13%","","",0,"<b>" & LitNDocumento & "</b>"
									DrawCelda "tdbordeCELDA7 bgcolor=" & color_terra & "  width=13%","","",0,"<b>" & LitTipoDocumento & "</b>"
									DrawCelda "tdbordeCELDA7 bgcolor=" & color_terra & "  width=10%","","",0,"<b>" & LitImporte & "</b>"
									DrawCelda "tdbordeCELDA7 bgcolor=" & color_terra & "  width=20%","","",0,"<b>" & LitMedio & "</b>"
								CloseFila
								fila=fila+1

								traspasoAnt=rstTraspaso("ntraspaso")
								traspasoActual=rstTraspaso("ntraspaso")

								while traspasoAnt=traspasoActual and not rstTraspaso.eof
									DrawFila color_blau
										DrawCelda "tdbordeCELDA7 width=2%","","",0,""
										DrawCelda "tdbordeCELDA7 width=10%","","",0,EncodeForHtml(rstTraspaso("fechaDet"))
										DrawCelda "tdbordeCELDA7 colspan=2 width=32%","","",0,EncodeForHtml(rstTraspaso("descripcionDet"))%>
										<td class='TDBORDECELDA7' style="width:13%">
											<%=EncodeForHtml(CalculaPagina(rstTraspaso("ndocumentodet"),rstTraspaso("tdocumentodet"),rstTraspaso("ndocumento_pro")))%>
										</td>
										<%DrawCelda "tdbordeCELDA7 width=13%","","",0,EncodeForHtml(rstTraspaso("tdocumentodet"))
										DrawCelda "tdbordeCELDA7 align=right width=10%","","",0,EncodeForHtml(cstr(formatnumber(rstTraspaso("importeDet"),rstTraspaso("ndecimales"),-1,0,-1))) & " " & EncodeForHtml(rstTraspaso("abreviatura"))
										DrawCelda "tdbordeCELDA7 width=20%","","",0, EncodeForHtml(rstTraspaso("medioDet"))
									CloseFila
									fila=fila+1
									traspasoAnt=rstTraspaso("ntraspaso")
									rstTraspaso.movenext
									if not rstTraspaso.eof then
										traspasoActual=rstTraspaso("ntraspaso")
									end if
								wend
								rstTraspaso.moveprevious
								'Fila de encabezado'
								DrawFila color_fondo
									DrawCelda "ENCABEZADOC width=6%","","",0,LitFecha
									DrawCelda "ENCABEZADOC width=6%","","",0,LitNTraspaso
									DrawCelda "ENCABEZADOC width=16%","","",0,LitDescripcion
									DrawCelda "ENCABEZADOC width=16%","","",0,LitResponsable
									DrawCelda "ENCABEZADOC width=13%","","",0,LitCajaOrg
									DrawCelda "ENCABEZADOC width=13%","","",0,LitCajaDes
									DrawCelda "ENCABEZADOC width=10%","","",0,LitImporte
									DrawCelda "ENCABEZADOC width=20%","","",0,LitMedio
								CloseFila
						end if
					end if
				end if
			  	traspasoAnt=rstTraspaso("ntraspaso")
				rstTraspaso.MoveNext
			wend
			DrawFila color_fondo
				DrawCelda "tdbordeCELDA7 colspan=4 width='50%'","","",0,"<b>" & LitTotalTraspasos & "</b>"
				DrawCelda "tdbordeCELDA7 colspan=4 align='right' width='50%'","","",0,"<b>" & totalTraspasos & "</b>"
			CloseFila
			DrawFila color_fondo
				DrawCelda "tdbordeCELDA7 colspan=4 width='50%'","","",0,"<b>" & LitTotalImporte & "</b>"
				DrawCelda "tdbordeCELDA7 colspan=4 align='right' width='50%'","","",0,"<b>" & EncodeForHtml(formatnumber(totalImporte,ndecimales,-1,0,-1)) & " " & EncodeForHtml(abreviatura) & "</b>"
			CloseFila%>
				</table>
			<%NavPaginas lote,lotes,campo,criterio,texto,2
			rstTraspaso.Close%>
        </table>
        <%end if
	end if%>
</form>
<%end if
set conn=nothing
set rstAux=nothing
set rst=nothing
set rstSelect=nothing
set rstAux1=nothing
set rstTraspaso=nothing
%>
</body>
</html>