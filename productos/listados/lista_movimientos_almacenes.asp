<%@ Language=VBScript %>
<%
''ricardo 16-11-2007 se cambia la dsn desde dsncliente a backendlistados
%>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloListMov%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>

<LINK REL="styleSHEET" href="../../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../../impresora.css" MEDIA="PRINT">
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../movimientos_almacenes.inc" -->
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../../catFamSubResponsive.inc" -->
<!--#include file="../../styles/formularios.css.inc" -->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->

<script language="javascript" src="../../jfunciones.js"></script>

<body onload="self.status='';" class="BODY_ASP">
<%sub Totales()%>
	<td class=dato align="right" bgcolor=<%=color_fondo%>><b><%=LitTotal%>&nbsp;</b></td>
	<%DrawCelda "DATO colspan='3' bgcolor=" & color_fondo,"","",0,"&nbsp;"
	total=rst.recordcount%>
	<td class=dato align="right" bgcolor=<%=color_fondo%>><b><%=cstr(formatnumber(total,0,-1,0,-1))%></b></td>
	<%if coste="on" then%>
	    <td class=dato align="right" bgcolor=<%=color_fondo%>><b><%=cstr(formatnumber(rstAux("total_c"),dec_prec,-1,0,-1))%></b></td>
	<% end if
	if importe="on" then%> 
	    <td class=dato align="right" bgcolor=<%=color_fondo%>><b><%=cstr(formatnumber(rstAux("total_i"),dec_prec,-1,0,-1))%></b></td>
	<%end if
end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************

const borde=0
%>
	<form name="lista_movimientos_almacenes" method="post">
	<% PintarCabecera "lista_movimientos_almacenes.asp"
		'***RGU 5/1/2006***
		dim mmp
	    ObtenerParametros("movimientos_almacenes")

		if enc.EncodeForJavascript(request.QueryString("mmp"))& "">"" then
			mmp=limpiaCadena(request.QueryString("mmp"))
		elseif enc.EncodeForJavascript(request.form("mmp")) & "">"" then
			mmp=limpiaCadena(request.form("mmp"))
		end if
		%><input type="hidden" name="mmp" value="<%=enc.EncodeForHtmlAttribute(mmp)%>"><%
		'***

		'Leer parámetros de la página
		mode		= enc.EncodeForJavascript(Request.QueryString("mode"))
		if ucase(mode) = "BROWSE" then mode ="imp"

		if enc.EncodeForJavascript(request.querystring("Dfecha"))>"" then
			TmpDfecha=limpiaCadena(request.querystring("Dfecha"))
		else
			TmpDfecha=limpiaCadena(request.Form("Dfecha"))
		end if

		if enc.EncodeForJavascript(request.querystring("Hfecha"))>"" then
			TmpHfecha=limpiaCadena(request.querystring("Hfecha"))
		else
			TmpHfecha=limpiaCadena(request.Form("Hfecha"))
		end if

		if enc.EncodeForJavascript(request.querystring("nserie"))>"" then
			TmpNSerie=limpiaCadena(request.querystring("nserie"))
		else
			TmpNSerie=limpiaCadena(request.Form("nserie"))
		end if
		CheckCadena TmpNSerie

		if enc.EncodeForJavascript(request.querystring("referencia"))>"" then
			TmpReferencia=limpiaCadena(request.querystring("referencia"))
		else
			TmpReferencia=limpiaCadena(request.Form("referencia"))
		end if
		TmpReferencia2=TmpReferencia
		if enc.EncodeForJavascript(request.querystring("nombre"))>"" then
			TmpDescripcion=limpiaCadena(request.querystring("nombre"))
		else
			TmpDescripcion=limpiaCadena(request.Form("nombre"))
		end if

		if enc.EncodeForJavascript(request.querystring("responsable"))>"" then
			TmpResponsable=limpiaCadena(request.querystring("responsable"))
		else
			TmpResponsable=limpiaCadena(request.Form("responsable"))
		end if
		CheckCadena TmpResponsable
		TmpResponsable2=TmpResponsable

		if enc.EncodeForJavascript(request.querystring("almorigen"))>"" then
			TmpAlmOrigen=limpiaCadena(request.querystring("almorigen"))
		else
			TmpAlmOrigen=limpiaCadena(request.Form("almorigen"))
		end if
		CheckCadena TmpAlmOrigen

		if enc.EncodeForJavascript(request.querystring("almdestino"))>"" then
			TmpAlmDestino=limpiaCadena(request.querystring("almdestino"))
		else
			TmpAlmDestino=limpiaCadena(request.Form("almdestino"))
		end if
		if enc.EncodeForJavascript(request.form("coste"))>"" then
			coste = limpiaCadena(request.form("coste"))
		else
			coste = limpiaCadena(request.querystring("coste"))
		end if		
		if enc.EncodeForJavascript(request.form("importe"))>"" then
			importe = limpiaCadena(request.form("importe"))
		else
			importe = limpiaCadena(request.querystring("importe"))
		end if			
		CheckCadena TmpAlmDestino

		strwhere=""

		if mode="imp" then%>
			<table width='100%' cellspacing="1" cellpadding="1">
   				<tr>
					<td width="30%" align="left">
					  	<font class=CELDAC7>&nbsp;(<%=LitEmitido%>&nbsp; <%=day(date)%>/<%=month(date)%>/<%=year(date)%>)</font>
					</td>
   					<td><font class='CABECERA'><b></b></font>
 	     					<font class=CELDA><b></b></font>
					</td>
					<td class=CELDARIGHT bgcolor="">
						<%fdesde=TmpDfecha
						fhasta=TmpHfecha
						fhasta=day(fhasta) & "/" & month(fhasta) & "/" & year(fhasta)
						if fdesde>"" then
							if fhasta>"" then
								%><%=LitPeriodoFechas%> : <%=fdesde%> - <%=fhasta%><%
							else
								%><%=LitPeriodoFechas%> : <%=LitDesde%>&nbsp;<%=fdesde%><%
							end if
						else
							if fhasta>"" then
								%><%=LitPeriodoFechas%> : <%=LitHasta%>&nbsp;<%=fhasta%><%
							else
							end if
						end if%>
					</td>
		   		</tr>
			</table>
			<hr/>
		<%else%>
			<br/>
		<%end if
		Alarma "lista_movimientos_almacenes.asp"

		set rstAux = Server.CreateObject("ADODB.Recordset")
		set rst = Server.CreateObject("ADODB.Recordset")
		set rst2 = Server.CreateObject("ADODB.Recordset")
		set rstSelect = Server.CreateObject("ADODB.Recordset")

		if mode="param"then
			
                    DrawDiv "1","",""
                    DrawLabel "","",LitDesdeFecha
                    DrawInput "", "", "Dfecha",iif(TmpDfecha>"",TmpDfecha,"01/01/" & year(date)), "" 
                    DrawCalendar "Dfecha"
                    CloseDiv
					
                    DrawDiv "1","",""
                    DrawLabel "","",LitHastaFecha
                    DrawInput "", "", "Hfecha",iif(TmpHfecha>"",TmpHfecha,day(date) & "/" & month(date) & "/" & year(date)), "" 
                    DrawCalendar "Hfecha"
                    CloseDiv
				
					rstAux.cursorlocation=3
					rstAux.open "select nserie,nombre as descripcion from series with (NOLOCK) where tipo_documento ='MOVIMIENTOS ENTRE ALMACENES' and nserie like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
					DrawSelectCelda "","","",0,LitSerie,"nserie",rstAux,enc.EncodeForHtmlAttribute(null_s(TmpNSerie)),"nserie","descripcion","",""
					rstAux.close
					
                    DrawDiv "1","",""
                    DrawLabel "","",LitReferencia2 %><input type="text" name="referencia" value="<%=iif(TmpReferencia>"",enc.EncodeForHtmlAttribute(null_s(TmpReferencia)),"")%>" size=20>						
					<%
                    CloseDiv
					
                    EligeCelda "input","add","left","","",0,LitNombre2,"nombre",0,iif(TmpDescripcion>"",enc.EncodeForHtmlAttribute(null_s(TmpDescripcion)),"")
				
					rstSelect.cursorlocation=3
					rstSelect.open "select dni, nombre from personal with (NOLOCK) where fbaja is null and dni like '" & session("ncliente") & "%' order by nombre",session("backendlistados")
					DrawSelectCelda "","","",0,LitResponsable,"responsable",rstSelect,enc.EncodeForHtmlAttribute(null_s(TmpResponsable)),"dni","nombre","",""
					rstSelect.close
				
					rstSelect.cursorlocation=3
					rstSelect.open "select codigo, descripcion from almacenes with (NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
					DrawSelectCelda "","","",0,LitAlmacenDestino,"almdestino",rstSelect,enc.EncodeForHtmlAttribute(null_s(TmpAlmDestino)),"codigo","descripcion","",""
					rstSelect.close
				
					rstSelect.cursorlocation=3
					rstSelect.open "select codigo, descripcion from almacenes with (NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
					DrawSelectCelda "","","",0,LitAlmacenOrigen,"almorigen",rstSelect,enc.EncodeForHtmlAttribute(null_s(TmpAlmOrigen)),"codigo","descripcion","",""
					rstSelect.close
				
			%><hr/><h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitCamposOpcionales%></h6><%     
        
				EligeCelda "check","add","","","",0,LitCosteTotal,"coste",0,""

                EligeCelda "check","add","","","",0,LitImporteTotal,"importe",0,""
		elseif mode="imp" then%>
			<input type="hidden" name="Dfecha" value="<%=enc.EncodeForHtmlAttribute(TmpDfecha)%>">
			<input type="hidden" name="Hfecha" value="<%=enc.EncodeForHtmlAttribute(TmpHfecha)%>">
			<input type="hidden" name="nserie" value="<%=enc.EncodeForHtmlAttribute(TmpNSerie)%>">
			<input type="hidden" name="referencia" value="<%=enc.EncodeForHtmlAttribute(TmpReferencia)%>">
			<input type="hidden" name="nombre" value="<%=enc.EncodeForHtmlAttribute(TmpDescripcion)%>">
			<input type="hidden" name="responsable" value="<%=enc.EncodeForHtmlAttribute(TmpResponsable)%>">
			<input type="hidden" name="almorigen" value="<%=enc.EncodeForHtmlAttribute(TmpAlmOrigen)%>">
			<input type="hidden" name="almdestino" value="<%=enc.EncodeForHtmlAttribute(TmpAlmDestino)%>">
			<input type="hidden" name="coste" value="<%=enc.EncodeForHtmlAttribute(coste)%>">
			<input type="hidden" name="importe" value="<%=enc.EncodeForHtmlAttribute(importe)%>">			
			<%MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='144'", DSNIlion)
			MAXPDF=d_lookup("maxpdf", "limites_listados", "item='144'", DSNIlion)%>
			<input type='hidden' name='maxpdf' value='<%=enc.EncodeForHtmlAttribute(MAXPDF)%>'>
			<input type='hidden' name='maxpagina' value='<%=enc.EncodeForHtmlAttribute(MAXPAGINA)%>'>
			<input type='hidden' name='maxmb' value='<%=enc.EncodeForHtmlAttribute(MB)%>'><%

			strwhere=" where"
			encabezado=0
                                                                                                          
			if TmpNSerie > "" then                                                            
				strwhere=strwhere & " m.nserie='" & TmpNSerie & "' and"
				%><font class=cab><b><%=LitSerie%>:&nbsp;</b></font><font class=cab><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(TmpNSerie)) & " - " & enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre","series","nserie='" & TmpNSerie & "'",session("backendlistados"))))%></font><br/><%
				encabezado=1
			end if
			if TmpReferencia > "" then
				strwhere=strwhere & " d.ref like '%" & TmpReferencia & "%' and"
				%><font class=cab><b><%=LitReferencia2%>:&nbsp;</b></font><font class=cab><%=enc.EncodeForHtmlAttribute(TmpReferencia)%>&nbsp;<%=enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre","articulos","referencia='" & TmpReferencia & "'",session("backendlistados"))))%></font><br/><%
				encabezado=1
			end if
			if TmpDescripcion > "" and TmpReferencia="" then
			    rstSelect.cursorlocation=3
				rstSelect.open "select referencia from articulos with (NOLOCK) where nombre like '%" & TmpDescripcion & "%' and referencia like '" & session("ncliente") & "%' ",session("backendlistados")
				if not rstSelect.eof then
					lista="('"
					while not rstSelect.eof
						lista=lista & rstSelect("referencia") & "','"
						rstSelect.movenext
					wend
					lista=mid(lista,1,len(lista)-2) & ")"
				else
					lista="('')"
				end if
				rstSelect.close
				strwhere=strwhere & " d.ref in " & lista & " and"
				%><font class=cab><b><%=LitNombre2%>:&nbsp;</b></font><font class=cab><%=TmpDescripcion%></font><br/><%
				encabezado=1
			end if                                                                                             
			if TmpResponsable > "" then
				strwhere=strwhere & " m.responsable='" & TmpResponsable & "' and"
				%><font class=cab><b><%=LitResponsable%>:&nbsp;</b></font><font class=cab><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(TmpResponsable)) & " - " & enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre","personal","dni='" & TmpResponsable & "'",session("backendlistados"))))%></font><br/><%
				encabezado=1
			else
				strwhere=strwhere & " m.responsable like '" & session("ncliente") & "%' and"
			end if
			if TmpAlmOrigen > "" then
				strwhere=strwhere & " d.almorigen='" & TmpAlmOrigen & "' and"
				%><font class=cab><b><%=LitAlmacenOrigen%>:&nbsp;</b></font><font class=cab><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(TmpAlmOrigen)) & " - " & enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","almacenes","codigo='" & TmpAlmOrigen & "'",session("backendlistados"))))%></font><br/><%
				encabezado=1
			end if
			if TmpAlmDestino > "" then
				strwhere=strwhere & " m.almdestino='" & TmpAlmDestino & "' and"
				%><font class=cab><b><%=LitAlmacenDestino%>:&nbsp;</b></font><font class=cab><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(TmpAlmDestino)) & " - " & enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","almacenes","codigo='" & TmpAlmDestino & "'",session("backendlistados"))))%></font><br/><%
				encabezado=1
			end if
			if fdesde > "" then
				strwhere=strwhere & " m.fecha>='" & TmpDfecha & "' and"
			end if
			if fhasta > "" then
				strwhere=strwhere & " m.fecha<='" & TmpHfecha & "' and"
			end if

			strwhere=strwhere & " m.responsable=p.dni and d.nmovimiento=m.nmovimiento and m.almdestino=a.codigo"
			strwhere=strwhere & " and m.nmovimiento like '" & session("ncliente") & "%' "

			if encabezado=1 then
				%><hr/><%
			end if

			'**RGU 5/1/2006
			if mmp=1 then
				usuario=session("ncliente") & session("usuario")
				strwhere=strwhere & " and m.responsable = '"&usuario&"' "
			end if
			'***

			strselect="select distinct m.nmovimiento,m.fecha,m.responsable,p.nombre as nomresponsable,m.almdestino,a.descripcion as nomalmacen,m.mercrecibida,m.total_coste,m.total_importe from movimientos as m with (NOLOCK),detalles_movimientos as d with (NOLOCK),personal as p with (NOLOCK),almacenes as a with (NOLOCK) " & strwhere & " order by m.fecha,m.nmovimiento"
			strselect2="select sum(total_coste)as total_c,sum(total_importe)as total_i from (select distinct m.nmovimiento, m.total_coste,m.total_importe from movimientos as m with (NOLOCK),detalles_movimientos as d with (NOLOCK),personal as p with (NOLOCK),almacenes as a with (NOLOCK) " & strwhere & ")tmp "
			rst.cursorlocation=3
			rst.Open strselect, session("backendlistados")
			rstAux.cursorlocation=3
			rstAux.Open strselect2, session("backendlistados")
			%><input type="hidden" name="NumRegs" value="<%=rst.Recordcount%>"><%
			if rst.EOF then
				rst.Close
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgDatosNoExiste%>");
				      parent.window.frames["botones"].document.location = "lista_movimientos_almacenes_bt.asp?mode=param";
					document.lista_movimientos_almacenes.action="lista_movimientos_almacenes.asp?mode=param";
					document.lista_movimientos_almacenes.submit();
				</script><%
			else
				'Calculos de páginas--------------------------
				lote=limpiaCadena(Request.QueryString("lote"))
				if lote="" then
					lote=1
				end if
				sentido=limpiaCadena(Request.QueryString("sentido"))

				lotes=rst.RecordCount/MAXPAGINA
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

				rst.PageSize=MAXPAGINA
				rst.AbsolutePage=lote
				'-----------------------------------------
				if lotes>1 then
					%><hr/><%
				end if
				NavPaginas lote,lotes,campo,criterio,texto,1%>
				<table width='100%' style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
					<%
					'Fila de encabezado
					DrawFila color_fondo
						DrawCelda "ENCABEZADOC","","",0,LitFecha
						DrawCelda "ENCABEZADOC","","",0,LitMovimiento
						DrawCelda "ENCABEZADOC","","",0,LitResponsable
						DrawCelda "ENCABEZADOC","","",0,LitAlmacenDestino
						DrawCelda "ENCABEZADOC","","",0,LitMercRecMov
						if coste="on" then
						    DrawCelda "ENCABEZADOR","","",0,LitCosteTotal
						end if
						if importe="on" then 
						    DrawCelda "ENCABEZADOR","","",0,LitImporteTotal
						end if
					CloseFila

		'PERMISOS DE HIPERVINCULOS ----------------------------------
		VinculosPagina(MostrarMovimientosAlm)=1:VinculosPagina(MostrarPersonal)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
		'---------------------------------
					suma=0
					fila=1
					while not rst.EOF and fila<=MAXPAGINA
						CheckCadena rst("nmovimiento")
						'Seleccionar el color de la fila.
						if ((fila+1) mod 2)=0 then
							color=color_blau                                                                      
						else                                                                                   
							color=color_terra                                                                      
						end if                                                                                      
						DrawFila color                                                                                                         
							DrawCelda "DATO","","",0,enc.EncodeForHtmlAttribute(null_s(rst("fecha")))
							doc=rst("nmovimiento")
							url=Hiperv(OBJMovimientosAlm,rst("nmovimiento"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),doc,LitVerMovimiento)
					    	 	%><td class=DATO><%=url%></td><%
							doc=rst("nomresponsable")
							url=Hiperv(OBJPersonal,rst("responsable"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),doc,LitVerPersonal)
					    	 	%><td class=DATO><%=url%></td><%
							DrawCelda "DATO","","",0,enc.EncodeForHtmlAttribute(null_s(rst("nomalmacen")))
							DrawCelda "DATO align='center'","","",0,iif(nz_b(rst("mercrecibida"))=-1,"Sí","No")
						    if coste="on" then
						        DrawCelda "DATO align='right'","","",0,enc.EncodeForHtmlAttribute(formatnumber(rst("TOTAL_COSTE")),dec_prec,-1,0,-1)
						    end if
						    if importe="on" then 
						        DrawCelda "DATO align='right'","","",0,enc.EncodeForHtmlAttribute(formatnumber(rst("TOTAL_IMPORTE")),dec_prec,-1,0,-1)
						    end if							
						CloseFila
						fila=fila+1
						rst.MoveNext
					wend
					if lote=lotes and not rstAux.EOF then
						Totales
					end if
					 NavPaginas lote,lotes,campo,criterio,texto,2
					rst.Close
					rstAux.Close
				%></table><%
			end if
		end if%>
	</form>
<%end if
set rstAux=nothing
set rst=nothing
set rst2=nothing
set rstSelect=nothing
%>
</body>
</html>
