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
<%Dim CodigoHTML%>

<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<META NAME='GENERATOR' Content='Microsoft Visual Studio 6.0'>
<META HTTP-EQUIV='Content-Type' Content="text/html; charset=<%=session("caracteres")%>">
<META HTTP-EQUIV='Content-style-Type' CONTENT='text/css'>
<LINK REL='styleSHEET' href='../pantalla.css' MEDIA='SCREEN'>
<LINK REL='styleSHEET' href='../impresora.css' MEDIA='PRINT'>
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
<!--#include file="../modulos.inc" -->
<!--#include file="cierres_caja.inc" -->

<!--#include file="../perso.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" --> 
<script language='javascript' src='../jfunciones.js'></script>

<script language="javascript" type="text/javascript">
//******************************************************************************
function IrAPagina(dedonde,maximo,NomParamPag)
{
	elemento="SaltoPagina" + dedonde;
	if (document.forms[0].name == "opciones")
	{
		indiceform=1;
	}
	else
	{
		indiceform=0;
	}
	if (isNaN(document.forms[indiceform].elements[elemento].value)) {
		npagina=1;
	}
	else {
		if (document.forms[indiceform].elements[elemento].value > maximo) {
			npagina=maximo;
		}
		else {
			if (document.forms[indiceform].elements[elemento].value <= 0) {
				npagina=1;
			}
			else {
				npagina=document.forms[indiceform].elements[elemento].value;
			}
		}
	}
	document.forms[indiceform].action=document.forms[indiceform].name + ".asp?" + NomParamPag + "=" + npagina +
	"&mode=browse";
	document.forms[indiceform].submit();
}

</script>

<body onload="self.status='';" class="BODY_ASP">

<%'Botones de navegación para las búsquedas.
sub NextPrev(lote,lotes,campo,criterio,texto,pos)%>
<table width='100%' border='0' cellspacing="1" cellpadding="1">
	<tr><td class='MAS'>
	    <%lote=cint(lote)
	    lotes=cint(lotes)
	    varias=false
        if lote>1 then%>
			<a class='CELDAREF' href="javascript:Mas('prev',<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(tdocumento)%>');">
			<img src="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a>
			<%varias=true
		end if
		if lotes>1 then
		    textopag=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)
		end if%>
		<font class='CELDA'><%=textopag%></font>

		<%if lote<lotes then%>
			<a class='CELDAREF' href="javascript:Mas('next',<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(tdocumento)%>');">
			<img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a>
			<%varias=true
		end if
		if varias=true then%>
	  	  <font class='CELDA'>&nbsp;&nbsp; Ir a Pag. <input class='CELDA' type="text" name="SaltoPagina<%=EncodeForHtml(null_s(pos))%>" size="2">&nbsp;&nbsp;<a class='CELDAREF' href="javascript:IrAPagina(<%=enc.EncodeForJavascript(pos)%>,<%=enc.EncodeForJavascript(lotes)%>,'lote');">Ir</a></font>
	  <%end if%>
	</td></tr>
</table>
<%end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion

	%><form name='listado_resumen_cierres' method='post'><%

	PintarCabecera "listado_resumen_cierres.asp"

	'Leer parámetros de la página'
  	mode=EncodeForHtml(Request.QueryString("mode"))
	lote=limpiaCadena(Request.QueryString("lote"))
	sentido=limpiaCadena(Request.QueryString("sentido"))

	dfecha = limpiaCadena(Request.form("Dfecha"))
	hfecha = limpiaCadena(Request.form("Hfecha"))
	dcierre = limpiaCadena(Request.form("Dcierre"))
	hcierre = limpiaCadena(Request.form("Hcierre"))
	tienda = limpiaCadena(Request.form("tienda"))
	caja = limpiaCadena(Request.form("caja"))

	set rst = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")

	%><input type="hidden" name="caju" value="<%=EncodeForHtml(cajau)%>"><%

	'WaitBoxOculto LitEsperePorFavor

	Alarma "listado_resumen_cierres.asp"
	%><hr/><%
'****************************************************************************************************************'
	if (mode="param") then
        
	    WaitBoxOculto LitEsperePorFavor
                DrawDiv "1","",""
                DrawLabel "","",LitDesdeFecha
                oneMonthAgo=DateAdd("m",-1,date)
                DrawInput "", "", "Dfecha", EncodeForHtml(day(oneMonthAgo) & "/" & month(oneMonthAgo) & "/" & year(oneMonthAgo)), ""
                DrawCalendar "Dfecha"
                CloseDiv
				
                DrawDiv "1","",""
                DrawLabel "","",LitHastaFecha    
                DrawInput "", "", "Hfecha",EncodeForHtml(day(date) & "/" & month(date) & "/" & year(date)), ""
                DrawCalendar "Hfecha"
                CloseDiv

			
                EligeCelda "input","add","left","","",0,LitDesdeCierre,"DCierre",0,""
                EligeCelda "input","add","left","","",0,LitHastaCierre,"HCierre",0,""
               
				    dim ConfigDespleg (1,13)
					dim ConfigDespleg2 (1,13)

					i=0
					ConfigDespleg(i,0)="tienda"
					ConfigDespleg(i,1)="200"
					ConfigDespleg(i,2)="10"
					ConfigDespleg(i,3)="select codigo, descripcion from tiendas with(nolock) where codigo like '" & session("ncliente") & "%'  order by descripcion"
					ConfigDespleg(i,4)=1
					ConfigDespleg(i,5)="width60"
					ConfigDespleg(i,6)="multiple"
					ConfigDespleg(i,7)="codigo"
					ConfigDespleg(i,8)="descripcion"
					ConfigDespleg(i,9)=LitTienda
					ConfigDespleg(i,10)=""
					ConfigDespleg(i,11)=""
					ConfigDespleg(i,12)=""


					i=0
					ConfigDespleg2(i,0)="caja"
					ConfigDespleg2(i,1)="200"
					ConfigDespleg2(i,2)="10"
					ConfigDespleg2(i,3)="select codigo, descripcion, tienda from cajas with(nolock) where codigo like '" & session("ncliente") & "%'  order by descripcion"
					ConfigDespleg2(i,4)=1
					ConfigDespleg2(i,5)="width60"
					ConfigDespleg2(i,6)="multiple"
					ConfigDespleg2(i,7)="codigo"
					ConfigDespleg2(i,8)="descripcion"
					ConfigDespleg2(i,9)=LitCaja
					ConfigDespleg2(i,10)=""
					ConfigDespleg2(i,11)=""
					ConfigDespleg2(i,12)=""

					DrawFila color_blau
		   			DibujaDesplegablesTienda3 ConfigDespleg,session("dsn_cliente")
					DibujaDesplegablesTienda3 ConfigDespleg2,session("dsn_cliente")
                    'response.Write("COMIENZO EL TEST")
					CloseFila
                    'javp fin
				'end if

            
'****************************************************************************************************************'
		'Mostrar el listado.'
	elseif mode="browse" then
	        DFecha=limpiacadena(request.Form("DFecha"))
	        Fechah=limpiacadena(request.Form("HFecha"))
	        DCierre=limpiacadena(request.Form("DCierre"))
	        Cierreh=limpiacadena(request.Form("HCierre"))
	        tienda=limpiacadena(request.Form("tienda"))
	        caja=limpiacadena(request.Form("caja"))

	        lot=limpiaCadena(request.querystring("lote"))
	        if lot="" then
	            lote=1
	        else
	            lote=cint(lot)
	        end if
	        sentido=limpiaCadena(request.QueryString("sentido"))
	        if sentido&""="next" then
	            lote=lote+1
	        end if
            if sentido&""="prev" then
	            lote=lote-1
	        end if%>

            <input type="hidden" name="dfecha" value="<%=EncodeForHtml(null_s(DFecha))%>" />
            <input type="hidden" name="hfecha" value="<%=EncodeForHtml(null_s(hfecha))%>" />
            <input type="hidden" name="dcierre" value="<%=EncodeForHtml(null_s(dcierre))%>" />
            <input type="hidden" name="hcierre" value="<%=EncodeForHtml(null_s(hcierre))%>" />
            <input type="hidden" name="tienda" value="<%=EncodeForHtml(null_s(tienda))%>" />
            <input type="hidden" name="caja" value="<%=EncodeForHtml(null_s(caja))%>" />

			<%n_decimales = null_z(d_lookup("ndecimales", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("dsn_cliente")))

			if DFecha>"" then
				%><font class=cab><b><%=LitDesdeFecha%> : </b></font><font class=cab><%=EncodeForHtml(DFecha)%></font><%
			end if
			if Fechah>"" then
				%><font class=cab><b><%=iif(DFecha>"","&nbsp;&nbsp;","") %><%=LitHastaFecha%> : </b></font><font class=cab><%=EncodeForHtml(Fechah)%></font><br/><%
			end if
			if DCierre>"" then
				%><font class=cab><b><%=LitDesdeCierre%> : </b></font><font class=cab><%=EncodeForHtml(DCierre)%></font><%
			end if
			if Cierreh>"" then
				%><font class=cab><b><%=iif(DCierre>"","&nbsp;&nbsp;","") %><%=LitHastaCierre%> : </b></font><font class=cab><%=EncodeForHtml(Cierreh)%></b></font><br/><%
			end if
			if tienda>"" then
			    strselect="select descripcion from tiendas where codigo in('"&replace(tienda,", ","','")&"') order by descripcion"
			    rstAux.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

			    strtienda=""
			    if not rstAux.EOF then
			        strtienda= rstAux("descripcion")
			        rstAux.MoveNext
			        while not rstAux.Eof
			            strtienda=strtienda&", "&rstAux("descripcion")
			            rstAux.MoveNext
			        wend
			    end if
			    rstAux.Close%>
				<!--<font class=cab><b><%=LitTienda%> : </b></font><font class=cab><%=replace(right(tienda,len(tienda)-5),", "&session("ncliente"),", ")%></b></font><br/>-->
				<font class=cab><b><%=LitTienda%> : </b></font><font class=cab><%=EncodeForHtml(strtienda)%></b></font><br/>
			<%end if
			if caja>"" then
                strselect="select descripcion from cajas where codigo in('"&replace(caja,", ","','")&"') order by descripcion"
			    rstAux.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

			    strcaja=""
			    if not rstAux.EOF then
			        strcaja= rstAux("descripcion")
			        rstAux.MoveNext
			        while not rstAux.Eof
			            strcaja=strcaja&", "&rstAux("descripcion")
			            rstAux.MoveNext
			        wend
			    end if
			    rstAux.Close%>
				<font class=cab><b><%=LitCaja%> : </b></font><font class=cab><%=EncodeForHtml(strcaja)%></b></font><br/>
			<%end if%>
			<hr/>

			<%'str=HFecha
			strexec="exec ResumenCierresCaja '"&caja&"','"&tienda&"'"
			strexec=strexec&",'"&DFecha&"' "
			strexec=strexec&",'"&Fechah&"' "
			strexec=strexec&",'"&session("ncliente")&"' "
			strexec=strexec&","&iif(DCierre="","null",DCierre)&" "
			strexec=strexec&","&iif(Cierreh="","null",Cierreh)&" "
			strexec=strexec&","&lote&" "

			rst.Open strexec,session("backendlistados"),adOpenKeyset,adLockOptimistic
			if rst.EOF then
			    %><font class='CEROFILAS'><%=LitCeroFilas%></font><%
			else
			    lotes=rst("paginas")
			    NextPrev lote,lotes,campo,criterio,texto,1

			    %><table width="100%"><%
			    TmpTienda=rst("codtienda")
			    DrawFila color_fondo
			    DrawCelda2Span "CELDA7", "left", true, LitTienda & ":" & EncodeForHtml(rst("tienda")) ,7
			    CloseFila
			    while not rst.EOF
			        if TmpTienda <> rst("codtienda") then
			            DrawFila ""
			                DrawCelda2 "CELDA7", "left", true,"&nbsp;"
			            CloseFila
                        DrawFila ""
                            DrawCelda2 "CELDA7", "left", true,"&nbsp;"
			            CloseFila
			            DrawFila color_fondo
			                DrawCelda2Span "CELDA7", "left", true, LitTienda & ":" & EncodeForHtml(rst("tienda")) , 7
			            CloseFila
			            TmpTienda=rst("codtienda")
			        end if
			        salida=1
			        if not rst.EOF then
			            if rst("orden")=1 then
		                    DrawFila color_fondo
			                    DrawCelda2Span "CELDA7", "left", true, LitMovimientosCaja,7
		                    CloseFila
		                    DrawFila color_terra
			                    DrawCelda2 "CELDA7", "left", true, LitTipoPago
			                    DrawCelda2Span "CELDAR7", "right", true, LitEntradas,2 '& " " & AbreviaturaMB
			                    DrawCelda2Span "CELDAR7", "right", true, LitSalidas,2 '& " " & AbreviaturaMB
			                    DrawCelda2Span "CELDAR7", "right", true, LitSaldoMin,2 '& " " & AbreviaturaMB
		                    CloseFila
		                    salida=0
		                 end if
			        end if
			        tot=0
			        sumEntradas=0
			        sumSalidas=0

			        while salida=0
			            tot=1
				        DrawFila color_blau
					        DrawCelda2 "CELDA7", "left", false, EncodeForHtml(rst("nomcampo"))
					        DrawCelda2Span "CELDAR7", "right", false, EncodeForHtml(formatNumber(rst("entradas"),n_decimales,-1,0,-1)),2
					        DrawCelda2Span "CELDAR7", "right", false, EncodeForHtml(formatNumber(rst("salidas"),n_decimales,-1,0,-1)),2
					        DrawCelda2Span "CELDAR7", "right", false, EncodeForHtml(formatNumber(rst("entradas")- rst("salidas"),n_decimales,-1,0,-1)),2
				        CloseFila
				        sumEntradas=sumEntradas+rst("entradas")
				        sumSalidas=sumSalidas+rst("salidas")
			            rst.MoveNext
				        if rst.EOF  then
				            salida=1
				        elseif rst("orden")<>1 then
				            salida=1
				        end if
			        wend

			        if tot=1 then
		                DrawFila color_terra
			                DrawCelda2 "CELDA7", "left", true, LitTotales
			                DrawCelda2Span "CELDAR7", "right", true, EncodeForHtml(formatNumber(SumEntradas,n_decimales,-1,0,-1)),2
			                DrawCelda2Span "CELDAR7", "right", true, EncodeForHtml(formatNumber(SumSalidas,n_decimales,-1,0,-1)),2
			                DrawCelda2Span "CELDAR7", "right", true, EncodeForHtml(formatNumber(SumEntradas-SumSalidas,n_decimales,-1,0,-1)),2
		                CloseFila
			        end if

			        '*********
			         salida=1
			        if not rst.EOF then
			            if rst("orden")=2 then
			        	    DrawFila color_fondo
			                    DrawCelda2Span "CELDA7", "left", true, LitTicketsEmitidos,7
		                    CloseFila
		                     salida=0
		                end if
		            end if
			        tot=0
			        sumIva=0
			        sumBi=0
			        sumTotal=0
			        TsumTotal=0
			        SerieAnt=""
			        while salida=0
					   SumBi=0
					   SumIva=0
					   SumTotal=0
                       if rst("iva")&"" = "" then
					       DrawFila color_blau
					           DrawCelda2Span "CELDAR7 bgcolor=" & color_terra, "left", true, " ",1
					        DrawCelda2 "CELDAR7 bgcolor=" & color_terra, "left", true, Litserie
					        DrawCelda2 "CELDAR7 bgcolor=" & color_terra, "left", true, LitTickets
					        DrawCelda2Span "CELDAR7 bgcolor=" & color_terra, "left", true, LitTicketDesde,2
					        DrawCelda2Span "CELDAR7 bgcolor=" & color_terra, "left", true, LitTicketHasta,2
					       CloseFila
                           salida=0
                           while salida =0
                               DrawFila color_blau
					               DrawCelda2Span "CELDAR7 ", "left", true, " ",1
					            DrawCelda2 "CELDAR7", "left", false, EncodeForHtml(null_s(rst("nomcampo")))
					            DrawCelda2 "CELDAR7", "left", false, EncodeForHtml(null_z(rst("tickets")))
					            DrawCelda2Span "CELDAR7", "left", false, EncodeForHtml(trimcodempresa(null_s(rst("ticketmin")))),2
					            DrawCelda2Span "CELDAR7", "left", false, EncodeForHtml(trimcodempresa(null_s(rst("ticketmax")))),2
					            rst.MoveNext
					            if rst.EOF then
					                salida=1
					            elseif rst("ticketmin")&""="" and rst("ticketmax")&""="" then
					                salida=1
					            end if
					        CloseFila
                           wend

                       end if
				       ' end if

				        if rst("tiva")&""<>"" then
				             tot=1
                                DrawFila color_terra
					                DrawCelda2Span "CELDA7 bgcolor=" & color_terra, "left", true, " ",1
						            'DrawCelda2 "CELDA7 bgcolor=" & color_terra, "left", true, Litiva
						            'DrawCelda2 "CELDA7 bgcolor=" & color_terra, "left", true, LitTickets
						            'DrawCelda2Span "CELDA7 bgcolor=" & color_terra, "left", true, LitTicketDesde,2
						            'DrawCelda2Span "CELDA7 bgcolor=" & color_terra, "left", true, LitTicketHasta,2
					            'DrawFila color_blau
						            DrawCelda2 "CELDAR7", "left", true, LitTipoIva
						            DrawCelda2 "CELDAR7", "left", true,LitBaseImponible
						            DrawCelda2Span "CELDAR7", "left", true, LitIva,2
						            DrawCelda2Span "CELDAR7", "left", true,LitTotalMin,2
					        'CloseFila
					            CloseFila
				            salida=0
				            while salida=0

				                DrawFila color_blau
				                    DrawCelda2Span "CELDA7 ", "left", true, " ",1
					                DrawCelda2 "CELDAR7", "left", false, EncodeForHtml(rst("tiva"))
					                DrawCelda2 "CELDAR7", "left", false, EncodeForHtml(formatnumber(rst("baseimponible"),n_decimales,-1,0,-1))
					                SumBi=SumBi + rst("baseimponible")
					                DrawCelda2Span "CELDAR7", "left", false, EncodeForHtml(formatnumber(rst("iva"),n_decimales,-1,0,-1)),2
					                SumIva=SumIva + rst("iva")
					                'DrawCelda2Span "CELDAR7", "left", false,formatnumber(rst("ventas"),n_decimales,-1,0,-1),2
					                'DrawCelda2Span "CELDAR7", "left", false,"&nbsp;",2
					                DrawCelda2Span "CELDAR7", "left", false, EncodeForHtml(formatnumber(rst("ventas"),n_decimales,-1,0,-1)),2
					                SumTotal=SumTotal + rst("ventas")
					                TSumTotal=TSumTotal + rst("ventas")
						            rst.MoveNext
						            if rst.EOF then
						                salida=1
						            elseif (rst("iva")&""="" and rst("iva")&""="") or(TmpTienda <> rst("codtienda")) or rst("orden")<>2  then
						                salida=1
						            end if
				                CloseFila
				             wend
				        end if
				        if rst.EOF  then
				            salida=1
				        elseif rst("orden")<>2 then
				            salida=1
				        end if
			        wend

			        if tot=1 then
			            DrawFila color_terra
				            'DrawCelda2 "CELDAR7", "left", true, LitTotalserie
				            DrawCelda2Span "CELDA7 ", "left", true, " ",1
				            DrawCelda2 "CELDAR7", "left", true,LitTotalMin
				            DrawCelda2 "CELDAR7", "left", true,EncodeForHtml(formatnumber(SumBi,n_decimales,-1,0,-1))
				            DrawCelda2Span "CELDAR7", "left", true,EncodeForHtml(formatnumber(SumIva,n_decimales,-1,0,-1)),2
				            DrawCelda2Span "CELDAR7", "left", true,EncodeForHtml(formatnumber(SumTotal,n_decimales,-1,0,-1)),2
				            'DrawCelda2Span "CELDAR7", "left", true,"&nbsp;",2
			            CloseFila
			        end if

			        '***********
			        salida=1
			        if not rst.EOF then
			            if rst("orden")=3 then
			        	    DrawFila color_fondo
			                    DrawCelda2Span "CELDA7", "left", true, LitVentasTipoPago,7
		                    CloseFila
		                    DrawFila color_terra
			                    DrawCelda2 "CELDA7", "left", true, LitTipoPago
			                    DrawCelda2Span "CELDAR7", "right", true, LitVentas,2 '& " " & AbreviaturaMB,1
			                    DrawCelda2Span "CELDAR7", "right", true, LitAnulaciones,2 '& " " & AbreviaturaMB,1
			                    DrawCelda2Span "CELDAR7", "right", true, LitTotalMin,2 '& " " & AbreviaturaMB,1
		                    CloseFila
	                        'DrawFila color_blau
			                 '   DrawCelda2 "CELDA7", "left", true, "&nbsp;"
			                  '  DrawCelda2Span "CELDAR7", "right", true, LitImporte,2
			                   ' DrawCelda2Span "CELDAR7", "right", true, LitImporte,2
			                    'DrawCelda2Span "CELDAR7", "right", true, LitImporte,2
		                    'CloseFila
		                    salida=0
		                end if
		            end if
		            SumTicketsVentas=0
		            SumVentas=0
		            SumTicketsAnul=0
		            SumAnulaciones=0
		            SumTotalTickets=0
		            SumTotalImporte=0
		            tot=0
		            while salida=0
		                tot=1
				        DrawFila color_blau
					        DrawCelda2 "CELDA7", "left", false, EncodeForHtml(rst("nomcampo"))
					        SumTicketsVentas=SumTicketsVentas + rst("tickets")
					        DrawCelda2Span "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("entradas"),n_decimales,-1,0,-1)),2
					        SumVentas=SumVentas + rst("entradas")
					        SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					        DrawCelda2Span "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("salidas"),n_decimales,-1,0,-1)),2
					        SumAnulaciones=SumAnulaciones + rst("salidas")
					        SumTotalTickets=SumTotalTickets + rst("ticketsanul")
					        DrawCelda2Span "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("saldo"),n_decimales,-1,0,-1)),2
					        SumTotalImporte=SumTotalImporte + rst("saldo")
				        CloseFila
		                rst.MoveNext
				        if rst.EOF  then
				            salida=1
				        elseif rst("orden")<>3 then
				            salida=1
				        end if
		            wend
		            if tot=1 then
                        DrawFila color_terra
				            DrawCelda2 "CELDA7", "left", true, LitTotales
				            'DrawCelda2 "CELDAR7", "right", true, SumTicketsVentas
				            DrawCelda2Span "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumVentas,n_decimales,-1,0,-1)),2
				            'DrawCelda2 "CELDAR7", "right", true, SumTicketsAnul
				            DrawCelda2Span "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumAnulaciones,n_decimales,-1,0,-1)),2
				            'DrawCelda2 "CELDAR7", "right", true, SumTotalTickets
				            DrawCelda2Span "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumTotalImporte,n_decimales,-1,0,-1)),2
			            CloseFila
		            end if

		            '******
		            salida=1
		            if not rst.EOF then
		                if rst("orden")=4 then
		                    DrawFila color_fondo
			                    DrawCelda2Span "CELDA7", "left", true, LitVentasOperadores,7
		                    CloseFila
		                    DrawFila color_terra
			                    DrawCelda2 "CELDA7", "left", true, LitOperador
			                    DrawCelda2Span "CELDAR7", "right", true, LitVentas,2 '& " " & AbreviaturaMB,2
			                    DrawCelda2Span "CELDAR7", "right", true, LitAnulaciones,2 '& " " & AbreviaturaMB,2
			                    DrawCelda2Span "CELDAR7", "right", true, LitTotalMin,2 '& " " & AbreviaturaMB,2
		                    CloseFila
		                    DrawFila color_blau
			                    DrawCelda2 "CELDA7", "left", true, "&nbsp;"
			                    DrawCelda2 "CELDAR7", "right", true, LitTickets
			                    DrawCelda2 "CELDAR7", "right", true, LitImporte
			                    DrawCelda2 "CELDAR7", "right", true, LitTickets
			                    DrawCelda2 "CELDAR7", "right", true, LitImporte
			                    DrawCelda2 "CELDAR7", "right", true, LitTickets
			                    DrawCelda2 "CELDAR7", "right", true, LitImporte
		                    CloseFila
		                    salida=0
		                 end if
		            end if
		            SumTicketsVentas=0
		            SumTicketsAnul=0
		            SumVentas=0
		            SumAnulaciones=0
		            SumTotalTickets=0
		            SumTotalImporte=0
		            tot=0
		            while salida=0
		                tot=1
				        DrawFila color_blau
					        DrawCelda2 "CELDA7", "left", false, EncodeForHtml(rst("nomcampo"))
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(rst("tickets"))
					        SumTicketsVentas=SumTicketsVentas + rst("tickets")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("entradas"),n_decimales,-1,0,-1))
					        SumVentas=SumVentas + rst("entradas")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(rst("ticketsanul"))
					        SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("salidas"),n_decimales,-1,0,-1))
					        SumAnulaciones=SumAnulaciones + rst("salidas")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(rst("tottickets"))
					        SumTotalTickets=SumTotalTickets + rst("tottickets")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("saldo"),n_decimales,-1,0,-1))
					        SumTotalImporte=SumTotalImporte + rst("saldo")
				        CloseFila
		                rst.MoveNext
				        if rst.EOF  then
				            salida=1
				        elseif rst("orden")<>4 then
				            salida=1
				        end if
		            wend
		            if tot=1 then
			            DrawFila color_terra
				            DrawCelda2 "CELDA7", "left", true, LitTotales
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(SumTicketsVentas)
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumVentas,n_decimales,-1,0,-1))
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(SumTicketsAnul)
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumAnulaciones,n_decimales,-1,0,-1))
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(SumTotalTickets)
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumTotalImporte,n_decimales,-1,0,-1))
			            CloseFila
		            end if

		            '*****
		            salida=1
		             if not rst.EOF then
		                if rst("orden")=5 then
 		                    DrawFila color_fondo
			                    DrawCelda2Span "CELDA7", "left", true, LitVentasTpv,7
		                    CloseFila
		                    DrawFila color_terra
			                    DrawCelda2 "CELDA7", "left", true, LitTpv
			                    DrawCelda2Span "CELDAR7", "right", true, LitVentas,2 '& " " & AbreviaturaMB,2
			                    DrawCelda2Span "CELDAR7", "right", true, LitAnulaciones,2 '& " " & AbreviaturaMB,2
			                    DrawCelda2Span "CELDAR7", "right", true, LitTotalMin,2 '& " " & AbreviaturaMB,2
		                    CloseFila
		                    DrawFila color_blau
			                    DrawCelda2 "CELDA7", "left", true, "&nbsp;"
			                    DrawCelda2 "CELDAR7", "right", true, LitTickets
			                    DrawCelda2 "CELDAR7", "right", true, LitImporte
			                    DrawCelda2 "CELDAR7", "right", true, LitTickets
			                    DrawCelda2 "CELDAR7", "right", true, LitImporte
			                    DrawCelda2 "CELDAR7", "right", true, LitTickets
			                    DrawCelda2 "CELDAR7", "right", true, LitImporte
		                    CloseFila
		                    salida=0
                        end if
		             end if
		             SumTicketsVentas=0
		             SumVentas=0
		             SumTicketsAnul=0
		             SumAnulaciones=0
		             SumTotalTickets=0
		             SumTotalImporte=0
		             tot=0
		             while salida=0
		                tot=1
				        DrawFila color_blau
					        DrawCelda2 "CELDA7", "left", false, EncodeForHtml(rst("nomcampo"))
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(rst("tickets"))
					        SumTicketsVentas=SumTicketsVentas + rst("tickets")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("entradas"),n_decimales,-1,0,-1))
					        SumVentas=SumVentas + rst("entradas")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(rst("ticketsanul"))
					        SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("salidas"),n_decimales,-1,0,-1))
					        SumAnulaciones=SumAnulaciones + rst("salidas")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(rst("TotTickets"))
					        SumTotalTickets=SumTotalTickets + rst("TotTickets")
					        DrawCelda2 "CELDAR7", "right", false, EncodeForHtml(formatnumber(rst("saldo"),n_decimales,-1,0,-1))
					        SumTotalImporte=SumTotalImporte + rst("saldo")
				        CloseFila
				        rst.movenext
				        if rst.EOF  then
				            salida=1
				        elseif rst("orden")<>5 then
				            salida=1
				        end if
			         wend
			         if tot=1 then
			            DrawFila color_terra
				            DrawCelda2 "CELDA7", "left", true, LitTotales
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(SumTicketsVentas)
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumVentas,n_decimales,-1,0,-1))
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(SumTicketsAnul)
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumAnulaciones,n_decimales,-1,0,-1))
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(SumTotalTickets)
				            DrawCelda2 "CELDAR7", "right", true, EncodeForHtml(formatnumber(SumTotalImporte,n_decimales,-1,0,-1))
			            CloseFila
			         end if
			        'rst.MoveNext
			    wend%>
			    </table>
			    <%rst.Close
			    NextPrev lote,lotes,campo,criterio,texto,2
			end if
	end if%>
</form>
<%connRound.close
set connRound = Nothing
set rst=nothing
set rstAux=nothing

end if%>
</body>
</html>