<%@ Language=VBScript %>
<%Server.ScriptTimeout = 1200
response.Buffer=true%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
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

<!--#include file="reg_inventario.inc" -->

<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" --> 
<!--#include file="../catfamsubResponsive.inc" -->
<%
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    function PonerHtml(sentido, lote) {
        cadena = "";
        cadena = "<table width='100%' BORDER='0' cellspacing='1' cellpadding='1'>";
        cadena = cadena + "<tr><td class='MAS'>";
        if (sentido == "next") lote = parseInt(marcoStock.document.inventario.lote.value) + 1;
        if (sentido == "prev") lote = parseInt(marcoStock.document.inventario.lote.value) - 1;
        if (sentido == "nulo") lote = parseInt(marcoStock.document.inventario.lote.value);

        lotes = parseInt(marcoStock.document.inventario.lotes.value);
        varias = false;
        if (lote > 1) {
            cadena = cadena + "<button type=\"button\" title=\"<%=LitAnterior%>\" aria-label=\"<%=LitAnterior%>\" onclick=\"javascript:Mas('prev'," + lote + ");\"><span class=\"ui-icon ui-icon-e\"></span></button>";
            varias = true;
        }
        texto = "<%=LitPagina%>" + " " + lote + " " + "<%=LitDe%>" + " " + lotes;
        cadena = cadena + "<font class='CELDA'>" + texto + "</font>";

        if (lote < lotes) {
            cadena = cadena + "<button type=\"button\" title=\"<%=LitSiguiente%>\" aria-label=\"<%=LitSiguiente%>\" onclick=\"javascript:Mas('next'," + lote + ");\"><span class=\"ui-icon ui-icon-w\"></span></button>";
            varias = true;
        }
        cadena = cadena + "</td></tr>";
        cadena = cadena + "</table>";
        document.getElementById("barras").innerHTML = cadena;
    }

    function Mas(sentido, lote) {
        document.getElementById("barras").style.display = "none";
        marcoStock.document.inventario.action = "inventario_datos.asp?mode=ver&sentido=" + sentido + "&lote=" + lote;
        marcoStock.document.inventario.submit();
    }

    function Cargar() {
        if (isNaN(document.inventario.stockmayoroigual.value)) window.alert("<%= LitNumStockMayor %>");
        else if (isNaN(document.inventario.nreg.value)) {
            alert("<%= LitNumRegPag %>");
        }
        else if (document.inventario.nreg.value > <%= MaxRegPag %>) {
            alert("<%= LitMaxRegPag %>");
        }
        else {
            marcoStock.marco1.document.getElementById("waitBoxOculto").style.visibility = "visible";
            cadena = "reg_inventario_sel.asp?mode=save&tabla=clientes&sms=0";
            document.inventario.target = document.getElementById("fr_hmarco").name;
            document.inventario.action = cadena;
            document.inventario.submit();
        }
    }

    function seleccionar(marco, formulario, check) {
        nregistros = eval(marco + ".document." + formulario + ".h_numRegs.value-1");
        if (eval("document.inventario." + check + ".checked")) {
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
            }
        }
        else {
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
            }
        }
    }

    function AplicarStock() {
        if (!isNaN(document.inventario.stockgral.value.replace(",", ".")) && document.inventario.stockgral.value != "") {
            elementos = marcoStock.document.inventario.hNRegs.value;
            for (i = 1; i <= elementos - 1; i++) {
                if (eval("marcoStock.document.inventario.check" + i + ".checked"))
                    eval("marcoStock.document.inventario.stock" + i + ".value=document.inventario.stockgral.value.replace('.',',')");
            }
        }
    }

    function AplicarSmin() {
        if (!isNaN(document.inventario.smingral.value.replace(",", ".")) && document.inventario.smingral.value != "") {
            elementos = marcoStock.document.inventario.hNRegs.value;
            for (i = 1; i <= elementos - 1; i++) {
                if (eval("marcoStock.document.inventario.check" + i + ".checked"))
                    eval("marcoStock.document.inventario.stock_minimo" + i + ".value=document.inventario.smingral.value.replace('.',',')");
            }
        }
    }

    function AplicarPrec() {
        if (!isNaN(document.inventario.precgral.value.replace(",", ".")) && document.inventario.precgral.value != "") {
            elementos = marcoStock.document.inventario.hNRegs.value;
            for (i = 1; i <= elementos - 1; i++) {
                if (eval("marcoStock.document.inventario.check" + i + ".checked"))
                    eval("marcoStock.document.inventario.p_recibir" + i + ".value=document.inventario.precgral.value.replace('.',',')");
            }
        }
    }

    function AplicarPser() {

        if (!isNaN(document.inventario.psergral.value.replace(",", ".")) && document.inventario.psergral.value != "") {
            elementos = marcoStock.document.inventario.hNRegs.value;
            for (i = 1; i <= elementos - 1; i++) {
                if (eval("marcoStock.document.inventario.check" + i + ".checked"))
                    eval("marcoStock.document.inventario.p_servir" + i + ".value=document.inventario.psergral.value.replace('.',',')");
            }
        }
    }

    function AplicarRep() {

        if (!isNaN(document.inventario.reposgral.value.replace(",", ".")) && document.inventario.reposgral.value != "") {
            elementos = marcoStock.document.inventario.hNRegs.value;
            for (i = 1; i <= elementos - 1; i++) {
                if (eval("marcoStock.document.inventario.check" + i + ".checked"))
                    eval("marcoStock.document.inventario.reposicion" + i + ".value=document.inventario.reposgral.value.replace('.',',')");
            }
        }
    }

    //----------------------------------------
    //Funcion para guardar registro
    //----------------------------------------
    function RegInventario() {
        var cadena;
        if (marcoStock.marco2.document.marcosSeleccion.h_numRegs.value <= 0) window.alert("<%= LitNoRegistros %>");
        else if (window.confirm("<%=LitMsgNInvConfirm%>")) {
            cadena = window.prompt("<%=LitObservDocumento%>", "");
            if (window.confirm("<%=LitMsgGenNInvConfirm%>")) {
                marcoStock.document.getElementById("waitBoxOculto").style.visibility = "visible";
                document.location = "reg_inventarioResultado.asp?mode=regularizacion&pend=" + document.inventario.pend.value + "&obs=" + cadena;
                parent.botones.document.location = "reg_inventario_bt.asp?mode=impresion";
            }
        }
    }

    function GuardarStock() {
        elementos = marcoStock.document.inventario.hNRegs.value;
        if (elementos == "") elementos = 0;
        if (elementos > 0) {
            error = "NO";
            msg = "";
            tiene = 0;
            for (i = 1; i <= elementos - 1; i++) {
                if (eval("marcoStock.document.inventario.check" + i + ".checked")) {
                    if (eval("isNaN(marcoStock.document.inventario.stock" + i + ".value.replace(',','.')) && marcoStock.document.inventario.stock" + i + ".value!=''")) {
                        if (tiene == 1) { msg = msg + " y "; }
                        msg = msg + "<%=LitStockReg%> " + i + " <%=LitMal%>";
                        error = "SI";
                        tiene = 1;
                    }

                    if (eval("isNaN(marcoStock.document.inventario.stock_minimo" + i + ".value.replace(',','.')) && marcoStock.document.inventario.stock_minimo" + i + ".value!=''")) {
                        if (tiene == 1) { msg = msg + " y "; }
                        msg = msg + "<%=LitStockMinReg%> " + i + " <%=LitMal%>";
                        error = "SI";
                        tiene = 1;
                    }

                    if (eval("isNaN(marcoStock.document.inventario.p_recibir" + i + ".value.replace(',','.')) && marcoStock.document.inventario.p_recibir" + i + ".value!=''")) {
                        if (tiene == 1) { msg = msg + " y "; }
                        msg = msg + "<%=LitPRecReg%> " + i + " <%=LitMal%>";
                        error = "SI";
                        tiene = 1;
                    }

                    if (eval("isNaN(marcoStock.document.inventario.p_servir" + i + ".value.replace(',','.')) && marcoStock.document.inventario.p_servir" + i + ".value!=''")) {
                        if (tiene == 1) { msg = msg + " y "; }
                        msg = msg + "<%=LitPServirReg %> " + i + " <%=LitMal%>";
                        error = "SI";
                        tiene = 1;
                    }

                    if (eval("isNaN(marcoStock.document.inventario.reposicion" + i + ".value.replace(',','.')) && marcoStock.document.inventario.reposicion" + i + ".value!=''")) {
                        if (tiene == 1) { msg = msg + " y "; }
                        msg = msg + "<%=LitReposReg%> " + i + " <%=LitMal%>";
                        error = "SI";
                        tiene = 1;
                    }
                };
            };
            if (error == "SI") window.alert(msg);
            else {
                marcoStock.document.inventario.action = "inventario_datos.asp?mode=save";
                marcoStock.document.inventario.submit();
            }
        }
        else window.alert("<%=LitNoCargaInven%>");
    }

    function MuestraFichero() {
        if (document.inventario.pistola.value != "") document.getElementById("frUpload").style.display = "";
        else document.getElementById("frUpload").style.display = "none";
    }
</script>
<body onload="self.status='';" class="BODY_ASP">
<%
'----------------------------------------------------------------------------
'Funciones
'----------------------------------------------------------------------------

'Botones de navegación para las búsquedas.
sub SpanNextPrev(lote,lotes,pos)%>
<table width='100%' border='0' cellspacing="1" cellpadding="1">
	<tr><td class='MAS'><%
	   lote=cint(lote)
	   lotes=cint(lotes)
	    varias=false
		if lote>1 then%>
            <button type="button" title="<%=LitAnterior%>" aria-label="<%=LitAnterior%>" onclick="javascript:Mas('prev',<%=enc.EncodeForJavascript(lote)%>);"><span class="ui-icon ui-icon-e"></span></button>
            <%varias=true
		end if
		texto=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)
		%><font class='CELDA'> <%=texto%> </font> <%

		if lote<lotes then%>
            <button type="button" title="<%=LitSiguiente%>" aria-label="<%=LitSiguiente%>" onclick="javascript:Mas('next',<%=enc.EncodeForJavascript(lote)%>);"><span class="ui-icon ui-icon-w"></span></button>
            <%varias=true
		end if

	%></td></tr>
</table>
<%end sub

sub SpanStock()%>
	<iframe class='width100 iframe-data md-table-responsive' name="fr_hmarco" id="fr_hmarco" src='reg_inventario_sel.asp?mode=add&sms=<%=enc.EncodeForHtmlAttribute(sms)%>&pend=<%=enc.EncodeForHtmlAttribute(CargaPendiente)%>' hidden></iframe>		
    <iframe class='col-lg-11 col-xxs-12' name="marcoStock" id="marcoStock" src='reg_MarcosSeleccion.asp?mode=browse' frameborder="no" scrolling="no" noresize="noresize" height="350"></iframe>
<%end sub

'Muestra los datos generales del documento
Sub DatosDocumento()%>
	<table width='100%' border='<%=borde%>' cellspacing="0" cellpadding="0">
		<tr bgcolor='<%=color_blau%>'><td>&nbsp;</td><td>&nbsp;</td></tr>
<%			DrawFila color_blau%>
				<td class="CABECERA" width="50%" align="left"><b><%=LitInventario%> : 
				<%=trimCodEmpresa(rstDet("ninventario"))%></b></td>
				<td class="CELDA" width="50%" align="center"><b><%=LitFecha%> : </b>
				<%=rstDet("fecha")%></td>
		</tr>
		<tr bgcolor='<%=color_blau%>'>
			<td class="CELDA" colspan="2"><b><%= LitResponsable%> : </b>
			<%= rstDet("nomresponsable")%>
			</td>
		</tr>
		<tr bgcolor='<%=color_blau%>'>
			<td class="CELDA" colspan="2"><b><%= LitObservaciones%> : </b>
			<%= rstDet("observaciones") %>
			</td>
		</tr>
	</table>

<%
end sub

sub Imprime
serieInv=session("ncliente") & "INVXX"
	
	LineasDatosAlbaranCli= 7
			LineasPaginaAlbaran= 33
			LineasPie= 10

strcadena="select n.observaciones,n.fecha,d.*,alm.descripcion as nomalmacen,art.nombre as nomarticulo,per.nombre as nomresponsable "
strcadena=strcadena & " from inventarios as n with (NOLOCK),detalles_inventario as d with (NOLOCK),almacenes as alm with (NOLOCK),articulos as art with (NOLOCK),personal as per with (NOLOCK) "
strcadena=strcadena & " where n.ninventario like '" & session("ncliente") & "%' and d.ninventario like '" & session("ncliente") & "%'"
strcadena=strcadena & " and alm.codigo like '" & session("ncliente") & "%' and art.referencia like '" & session("ncliente") & "%'"
strcadena=strcadena & " and per.dni like '" & session("ncliente") & "%' "
strcadena=strcadena & " and d.ninventario=n.ninventario and n.ninventario='" & ninventario & "' and art.referencia=d.referencia "
strcadena=strcadena & " and alm.codigo=d.almacen and per.dni=n.responsable "
''			rst.open "select * from inventarios where ninventario='"& ninventario & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
''			rstDet.open "select * from detalles_inventario where ninventario='"& ninventario & "' order by item", session("dsn_cliente"),adOpenKeyset,adLockOptimistic


rstDet.cursorlocation=3
rstDet.open strcadena,session("dsn_cliente")

			Dim Salir

			TotalesEscrito="NO"
			DesgloseEscrito="NO"
			Pagina=1
			Salir=False

			'Bucle de control del Número de páginas del documento
			while Not Salir
					LineaActual=1

					'Cabecera del albaran ***********************************************************************************
'					Cabecera Pagina,total_paginas
					total_paginas=total_paginas+1
					'LineaActual=LineaActual + 7
					'Datos del albaran **************************************************************************************
					if Pagina=1 then
						DatosDocumento
'						LineaActual=LineaActual + LineasDatosAlbaranCli - 1
					end if
					'Linea en blanco%>
					<br/>
<%
					LineaActual=LineaActual + 1

					'Cabecera de los detalles
					if not rstDet.EOF then%>
						<table width='100%' border='0' cellspacing="1" cellpadding="1">
                            <%DrawFila color_fondo
								DrawCelda "ENCABEZADOL","","",0,LitItem
								DrawCelda "ENCABEZADOL","","",0,LitRef
								DrawCelda "ENCABEZADOL","","",0,LitNombre
								DrawCelda "ENCABEZADOL","","",0,LitAlmacen
''ricardo 5-5-2006 eleccion de ver el coste por parametro
if cstr(vcri)<>"1" then
								DrawCelda "ENCABEZADOR","","",0,LitCoste
end if
								DrawCelda "ENCABEZADOR","","",0,LitStockActual
								DrawCelda "ENCABEZADOR","","",0,LitNuevoStock
								DrawCelda "ENCABEZADOR","","",0,LitStockMin
								DrawCelda "ENCABEZADOR","","",0,LitNuevoStockMin
								DrawCelda "ENCABEZADOR","","",0,LitStockRepos
								DrawCelda "ENCABEZADOR","","",0,LitNuevoStockRepos
							CloseFila
							LineaActual=LineaActual + 1
					end if
					'Detalles del albaran

					while not rstDet.EOF and LineaActual<=(LineasPaginaAlbaran-(LineasPie + 1))

						DrawFila color_blau

							DrawCelda "CELDALEFT","","",0,enc.EncodeForHtmlAttribute(null_s(rstDet("item")))
							DrawCelda "CELDALEFT","","",0,trimCodEmpresa(rstDet("referencia"))
						'	Lineasdetalle=cdbl(NumeroDeLineasObservaciones(iif(rstDet("descripcion")>"",rstDet("descripcion"),"")))
							
							DrawCelda "CELDALEFT","","",0,enc.EncodeForHtmlAttribute(null_s(rstDet("nomarticulo")))
							DrawCelda "CELDALEFT","","",0,enc.EncodeForHtmlAttribute(null_s(rstDet("nomalmacen")))
''ricardo 5-5-2006 eleccion de ver el coste por parametro
if cstr(vcri)<>"1" then
							DrawCelda "CELDARIGHT","","",0,formatnumber(rstDet("coste"),DEC_PREC,-1,0,-1)
end if
							
							DrawCelda "CELDARIGHT","","",0,formatnumber(rstDet("stock"),DEC_PREC,-1,0,-1)
							DrawCelda "CELDARIGHT","","",0,formatnumber(rstDet("nuevo_stock"),DEC_PREC,-1,0,-1)
							DrawCelda "CELDARIGHT","","",0,formatnumber(rstDet("stock_min"),DEC_PREC,-1,0,-1)
							DrawCelda "CELDARIGHT","","",0,formatnumber(rstDet("nuevo_stock_min"),DEC_PREC,-1,0,-1)
							DrawCelda "CELDARIGHT","","",0,formatnumber(rstDet("stock_repos"),DEC_PREC,-1,0,-1)
							DrawCelda "CELDARIGHT","","",0,formatnumber(rstDet("nuevo_stock_repos"),DEC_PREC,-1,0,-1)

						CloseFila
'						LineaActual=LineaActual + 1 	' + (Lineasdetalle-1)
                        response.Flush
						rstDet.MoveNext
					wend

					if rstDet.EOF  then
%>
						</table>
						<br/>
<%
						LineaActual=LineaActual + 1
					end if

					if LineaActual>=(LineasPaginaAlbaran-(LineasPie + 1)) then%>
						</table>
<%
					end if

				rstDet.movelast

				'Observaciones del albaran ******************************************************************************
				LineasObservaciones2=cdbl(NumeroDeLineasObservaciones(iif(rstDet("observaciones")>"",rstDet("observaciones"),"")))
				if (LineaActual + LineasObservaciones2)<=(LineasPaginaAlbaran-(LineasPie + 1)) then
					if rstDet("observaciones")>"" then%>
						<table width='100%' border='<%=borde%>' cellspacing="0" cellpadding="0">
							<tr bgcolor='<%=color_blau%>'><td>&nbsp;</td><td>&nbsp;</td></tr>
<%									LineaActual=LineaActual + 1%>
							<tr bgcolor='<%=color_blau%>' rowspan="2">
<%
								EligeCelda "text", mode,"CELDA valign='top'","17","",0,"<b>"+LitObservaciones+"</b>","observaciones",2,""
								EligeCelda "text", mode,"CELDA","83","",0,"","observaciones",2,enc.EncodeForHtmlAttribute(pintar_saltos_espacios(iif(rstDet("observaciones")>"",rstDet("observaciones"),"")))
%>
							</tr>
							<%LineaActual=LineaActual + 2 + LineasObservaciones2%>
							<tr bgcolor='<%=color_blau%>'><td>&nbsp;</td><td>&nbsp;</td></tr><%
							LineaActual=LineaActual + 1%>
						</table>
<%
						end if
						Salir=True
					end if


				'Completar lineas en blanco hasta el final del cuerpo%>
				<table width='100%' border='0' cellspacing="1" cellpadding="1">
<%
					For LineasEnBlanco=LineaActual to (LineasPaginaAlbaran-(LineasPie + 1))
						DrawFila color_blau
							DrawCelda "ENCABEZADOL","","",0,"&nbsp;"
						CloseFila
						LineaActual=LineaActual + 1
					next
%>
				</table>

<%				'Leyenda y datos registrales
				if 1=2 then
%>

				<table width='100%' border='<%=borde%>' cellspacing="0" cellpadding="0">
<%
					DrawFila color_blau
						DrawCelda "ENCABEZADOC","","",0,iif(Leyenda>"",leyenda,"&nbsp;")
					CloseFila
					LineaActual=LineaActual + 1
						TmpCif=d_lookup("empresa","series","nserie='" & serieInv & "'",session("dsn_cliente"))
						rstCliente.cursorlocation=3
						rstCliente.Open "select leyenda from Empresas where cif='" & TmpCif & "'",session("dsn_cliente")
						contenido="<table width='100%'><tr><td class='CELDAC6' align='center' width='100%'>" & enc.EncodeForHtmlAttribute(null_s(rstCliente("leyenda"))) & "</td></tr></table>"
						DibujarLeyenda total_paginas-1,contenido,"","","px","","998","","","vertical"
						rstCliente.close
						LineaActual=LineaActual + 1%>
				</table>
                <%end if
				'Salto a pagina nueva ***********************************************************************************
				if CopiaActual<Ncopias then%>
					<h6 class=SALTO>&nbsp;</h6>
                <%elseif Not Salir then %>
					<h6 class=SALTO>&nbsp;</h6>
                <%end if
				Pagina=Pagina+1
                response.Flush
			wend 'while not salir
			rstDet.Close
end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0%>
<form name="inventario" method="post">
	<%  Ayuda "inventario.asp"

		'Leer parámetros de la página
  		mode=Request.QueryString("mode")
		if request.querystring("lote") >"" then
		   lote = limpiaCadena(request.querystring("lote"))
		elseif request.form("lote")>"" then
		   lote = limpiaCadena(request.form("lote"))
		else
		   lote = 1
		end if
		if request.querystring("referencia") >"" then
		   referencia = limpiaCadena(request.querystring("referencia"))
		else
		   referencia = limpiaCadena(request.form("referencia"))
		end if
		if request.querystring("nombre") >"" then
		   nombre = limpiaCadena(request.querystring("nombre"))
		else
		   nombre = limpiaCadena(request.form("nombre"))
		end if
		if request.querystring("almacen") >"" then
		   almacen = limpiaCadena(request.querystring("almacen"))
		else
		   almacen = limpiaCadena(request.form("almacen"))
		end if
		if request.querystring("familia") >"" then
		   familia = limpiaCadena(request.querystring("familia"))
		else
		   familia = limpiaCadena(request.form("familia"))
		end if
		if request.querystring("ordenar") >"" then
		   ordenar = limpiaCadena(request.querystring("ordenar"))
		else
		   ordenar = limpiaCadena(request.form("ordenar"))
		end if
        dim CargaPendiente
		if request.querystring("pend") >"" then
		   CargaPendiente = limpiaCadena(request.querystring("pend"))
		else
		   CargaPendiente = limpiaCadena(request.form("pend"))
	    end if
        if request.querystring("art_type") >"" then
		   art_type = limpiaCadena(request.querystring("art_type"))
		else
		   art_type = limpiaCadena(request.form("art_type"))
		end if
        if request.querystring("artbaja") >"" then
		   artbaja = limpiaCadena(request.querystring("artbaja"))
		else
		   artbaja = limpiaCadena(request.form("artbaja"))
		end if

	    observaciones = limpiaCadena(request.querystring("obs"))

		
''ricardo 5-5-2006 se pone un parametro de usuario para que no se vea el coste en el formato de impresion
dim vcri
' cag 6-6-6 Parametro de usuario que trae lista de almacenes dependiendo del usuario
dim au  

ObtenerParametros "reg_inventario.asp"

%><input type="hidden" name="vcri" value="<%=enc.EncodeForHtmlAttribute(vcri)%>">
<input type="hidden" name="au" value="<%=enc.EncodeForHtmlAttribute(au)%>"> <%

	au=preparar_lista(au)

	if mode<>"ver" then
		PintarCabecera "reg_inventario.asp"
		%><br/>
<%
	else
%>
		<table width='100%'>
		   	<tr>
				<td width="30%" align="center" bgcolor="<%=color_fondo%>"><font class='CABECERA'><b><%=LitTitulo%></b></font></td>
				<td><font class='CABECERA'><b></b></font><font class=CELDA><b></b></font></td>
		  		<td align="right">
      		    	<font class='CABECERA'><b></b></font>
					<%pagina="inventario_imp.asp?referencia=" & referencia & "&nombre=" & nombre & _
					"&familia=" & familia & "&almacen=" & almacen & "&ordenar=" + ordenar
					response.write(pagima)%>
          			<a class='CELDAREFB' href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(pagina)%>','I',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitFormatoImpre%>'; return true;" OnMouseOut="self.status=''; return true;"><font class=CELDA><b>Formato de Impresión</b></font></a>
		  		</td>
			</tr>
		</table>
		<hr/>
<%
	end if

	Alarma "inventario.asp"

  set rstAux = Server.CreateObject("ADODB.Recordset")
  set rst = Server.CreateObject("ADODB.Recordset")
  set rstSelect = Server.CreateObject("ADODB.Recordset")
  set rstDet = Server.CreateObject("ADODB.Recordset")

	dni=d_lookup("dni","personal","login='" & session("usuario") & "' and dni like '"+Session("ncliente")+"%'",session("dsn_cliente"))

  '*********************************************************************************************
  'Se muestran parametros de seleccion
  '*********************************************************************************************
  if mode="param" then
		TmpAlmacen= ""
		TmpAlmacenDef= ""
		maxpagina= 20	'd_lookup("maxpagina", "limites_listados", "item='52'", DSNIlion)
		
		'Poner por defecto el almacen correspondiente a la tienda donde esta el tpv y caja del fichero cetel.tpv'

		linea1=session("f_tpv")
		linea2=session("f_caja")
		linea3=session("f_empr")

		if linea1<>"" and linea2<>"" and linea3<>"" then

			'Obtenemos la tienda y el almacen correspondiente.

			strSelect = "select c.almacen from tpv a with(nolock), cajas b with(nolock), tiendas c with(nolock) where a.caja=b.codigo and b.tienda=c.codigo and tpv='" & linea1 & "' and b.codigo='" & linea2 & "'"

			rstAux.cursorlocation=3
			rstAux.open strSelect,session("dsn_cliente")

			if linea3=session("ncliente") then
				if rstAux.eof then
					TmpAlmacenDef = ""
				else
					TmpAlmacenDef = rstAux("almacen")
				end if
			else
				TmpAlmacenDef = ""
			end if
			rstAux.close
		end if%>
		<SPAN ID="CapaNoAltaPersonal" style="display:none">
		 	<%waitbox LitMsgUsuarioPersonalNoExiste%>
		</SPAN>
  
		<SPAN ID="CapaParametros" style="display:none">
                    <%EligeCelda "input","add","","","",0,LitConref,"referencia",25,enc.EncodeForHtmlAttribute(referencia)
					
                    EligeCelda "input","add","","","",0,LitConNombre,"nombre",25,enc.EncodeForHtmlAttribute(nombre)
					'cag
					if au>"" then
						rstSelect.open "select codigo, descripcion from almacenes with(nolock) where codigo in " & au  & " order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						'cag
                        DrawDiv "1","",""
                        DrawLabel "","",LitAlmacen
						%>
							<select class='width60' name="almacen" style='width:165px'>
							<%do while not rstSelect.Eof %>
							<option <%=iif(rstSelect("codigo")=almacen,"selected","")%> value="<%=enc.EncodeForHtmlAttribute(null_s(rstSelect("codigo")))%>"> <%=enc.EncodeForHtmlAttribute(null_s(rstSelect("descripcion")))%></option>
							<% rstSelect.moveNext
							loop %>		
							</select>
						 <%CloseDiv
					else
					'fin cag
						if TmpAlmacenDef > "" and referencia="" then
							rstSelect.open "select codigo, descripcion from almacenes with(nolock) where codigo = '" & TmpAlmacenDef & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
							DrawSelectCelda "width60 style='width: 175px;' ","","",0,LitAlmacen,"almacen",rstSelect,enc.EncodeForHtmlAttribute(null_s(TmpAlmacenDef)),"codigo","descripcion","",""
							CargaPendiente=TmpAlmacenDef
						else
							TmpAlmacen=d_lookup("almacen","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
							rstSelect.open "select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
							DrawSelectCelda "width60 style='width: 175px;' ","","",0,LitAlmacen,"almacen",rstSelect,enc.EncodeForHtmlAttribute(null_s(TmpAlmacen)),"codigo","descripcion","",""
						end if
					'cag
					end if
					'fin cag
					rstSelect.close
					rstSelect.open "select codigo, tipo,descripcion from pistolas with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					DrawSelectCelda "CELDAL7 style='width: 175px;' ","","",0,LitPistola,"pistola",rstSelect,"","codigo","descripcion","onchange","MuestraFichero()"
					rstSelect.close
                    %><iframe name='marcoUpload' id='frUpload' class="width60" src='reg_inventario_upload.asp?mode=param' frameborder='no' scrolling='no' noresize='noresize' style='display: none;'></iframe>
                    <input type="hidden" name="pend" value="<%=enc.EncodeForHtmlAttribute(CargaPendiente)%>" /><%
					dim ConfigDespleg (3,13)
				
					i=0
					ConfigDespleg(i,0)="categoria"
					ConfigDespleg(i,1)=""
					ConfigDespleg(i,2)="5"
					ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
					ConfigDespleg(i,4)=1
					ConfigDespleg(i,5)="width60"
					ConfigDespleg(i,6)="MULTIPLE"
					ConfigDespleg(i,7)="codigo"
					ConfigDespleg(i,8)="nombre"
					ConfigDespleg(i,9)=LitCategoria
					ConfigDespleg(i,10)=categoria
					ConfigDespleg(i,11)=""
					ConfigDespleg(i,12)=""
				
					i=1
					ConfigDespleg(i,0)="familia_padre"
					ConfigDespleg(i,1)=""
					ConfigDespleg(i,2)="5"
					ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
					ConfigDespleg(i,4)=1
					ConfigDespleg(i,5)="width60"
					ConfigDespleg(i,6)="MULTIPLE"
					ConfigDespleg(i,7)="codigo"
					ConfigDespleg(i,8)="nombre"
					ConfigDespleg(i,9)=LitFamilia
					ConfigDespleg(i,10)=familia_padre
					ConfigDespleg(i,11)=""
					ConfigDespleg(i,12)=""
				
					i=2
					ConfigDespleg(i,0)="familia"
					ConfigDespleg(i,1)=""
					ConfigDespleg(i,2)="5"
					ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
					ConfigDespleg(i,4)=1
					ConfigDespleg(i,5)="width60"
					ConfigDespleg(i,6)="MULTIPLE"
					ConfigDespleg(i,7)="codigo"
					ConfigDespleg(i,8)="nombre"
					ConfigDespleg(i,9)=LitSubFamilia
					ConfigDespleg(i,10)=familia
					ConfigDespleg(i,11)=""
					ConfigDespleg(i,12)=""
				
					DibujaDesplegables ConfigDespleg,session("dsn_cliente")
                    DrawDiv "1","",""
                    DrawLabel "","",LitArtType
                    set conn = Server.CreateObject("ADODB.Connection")
                    set command =  Server.CreateObject("ADODB.Command")
                    conn.open session("backendListados")
                    command.ActiveConnection =conn
                    command.CommandTimeout = 0
                    command.CommandText="getAllEntityTypeByType"
                    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    command.Parameters.Append command.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))
                    command.Parameters.Append command.CreateParameter("@type", adVarChar, adParamInput, 20, "ARTICULO")

                    set rstArtType = command.execute
		            
		            %><select multiple="multiple" size="5" class="width60" name="art_type">
			 	        <%while not rstArtType.eof%>
		   			        <option value="<%=enc.EncodeForHtmlAttribute(rstArtType("codigo"))%>"><%=enc.EncodeForHtmlAttribute(null_s(rstArtType("descripcion")))%></option>
					        <%rstArtType.movenext%>
				        <%wend%>
			    	        <option selected="selected" value=""> </option>
		   	        </select>
			        <%rstArtType.close
                    conn.close
                    set rstArtType = nothing
                    set command = nothing
                    set conn = nothing
                    CloseDiv
                     DrawDiv "1","",""
                     DrawLabel "","",LitOrdenar
					%><select class='width60' name="ordenar"><option selected value="REFERENCIA"><%=LitReferenciaMay%></option>
							<%if ordenar="NOMBRE" then%>
				   				<option selected value="NOMBRE"><%=LitNombreMay%></option>
							<%else%>
					   			<option value="NOMBRE"><%=LitNombreMay%></option>
							<%end if %>
							<%if ordenar="ALMACEN" then%>
		   						<option selected value="ALMACEN"><%=LitAlmacenMay%></option>
							<%else%>
					   			<option value="ALMACEN"><%=LitAlmacenMay%></option>
							<%end if %>
							<%if ordenar="STOCK" then%>
				   				<option selected value="STOCK"><%=LitStockMay%></option>
							<%else%>
			   					<option value="STOCK"><%=LitStockMay%></option>
							<%end if %>
			   			</select>		
		            <%CloseDiv
				    'cag
					if au>"" then
						stockmayoroigual=""
					else
					'fin cag
						if stockmayoroigual & ""="" then stockmayoroigual="0"
					'cag
					end if
					'fin cag
                    EligeCelda "input","add","","","",0,LitStockMayorOIgual,"stockmayoroigual",25,stockmayoroigual

                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitStockMin2
					if stock_min="on" then
						%><input type='checkbox' name='stock_min' checked><%
					else
						%><input type='checkbox' name='stock_min'><%
					end if
					CloseDiv
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitPrep2
					if stock_rep="on" then
						%><input type='checkbox' name='stock_rep' checked><%
					else%><input type='checkbox' name='stock_rep'>						
                    <%end if
                    CloseDiv
                     EligeCelda "input","add","","","",0,LitMaxReg,"nreg",6,maxpagina
                     DrawDiv "1", "", ""
                     DrawLabel "", "", LitArtUnsusb
                    %><input type='checkbox' name='artbaja' <%=iif(artbaja&""<>"", "checked", "")%>>
                     
                <%CloseDiv
                     DrawDiv "1", "", ""
                     DrawLabel "", "", LitCargarArticulos
                    %><a class="CELDAREF ic-accept" href="javascript:Cargar();"><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>"></a><%
                     CloseDiv%>
			<hr/>
            <%WaitBoxOculto LitEsperePorFavor%><%
		
			'*********************************************************************************************
			' Se muestran los datos de la consulta
			'*********************************************************************************************%>
			<SPAN ID="Modificar" style="display: ">
				<%SpanStock%>
			</SPAN>
		</SPAN>
        <%if dni&""="" then%>
			<script language="javascript" type="text/javascript">
			    parent.botones.document.location="reg_inventario_bt.asp";
			    CapaNoAltaPersonal.style.display = "";
			</script>
		<%else%>
			<script language="javascript" type="text/javascript">
			    CapaParametros.style.display = "";
			</script><%
		end if
	elseif mode="regnoleidos" then
		ninventario= limpiacadena(request.querystring("ninventario"))
		
		set conn = Server.CreateObject("ADODB.Connection")
		set Command =  Server.CreateObject("ADODB.Command")
		conn.open session("dsn_cliente")	
		Command.ActiveConnection =conn
		Command.CommandTimeout = 0
		command.CommandText="RegularizaStockNoLeidos"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_ninventario",adVarChar,adParamInput,30,ninventario)
		Command.Execute , , adExecuteNoRecords
		
		%><input type="hidden" name="hninventario" value="<%=enc.EncodeForHtmlAttribute(ninventario)%>"><%
		
		Imprime
		
		conn.close
		set command=nothing
		set conn=nothing
	elseif mode="regularizacion" then
		set conn = Server.CreateObject("ADODB.Connection")
		
		''ricardo 07/09/2011 se cambiara el usuario del dsncliente por el de DSNImport
		dsnCliente=session("dsn_cliente")
		initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")

		donde=inStr(1,DSNImport,"Initial Catalog=",1)
		donde_fin=InStr(donde,DSNImport,";",1)
		if donde_fin=0 then
			donde_fin=len(DSNImport)
		end if
		cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))

		dsnCliente=cadena_dsn_final
		
		conn.open dsnCliente
		conn.CommandTimeout = 0
	
		sdni= d_lookup("dni", "personal", "login='"& session("usuario") &"' and dni like '" & session("ncliente") & "%'", session("dsn_cliente"))

		if isnull(sdni) then%>
			<script language="javascript" type="text/javascript">
			    window.alert("<%= LitUsuarioPersonal %>");
			</script>
        <%else
			ninventario= "0"    
			set rstAux = conn.execute("EXEC GeneraDocRegularizacionInventario " & _
			" @p_ncliente='"& session("ncliente") &"', @p_nusuario='"& session("usuario") &"', "& _
			" @p_dni='"& sdni &"', @p_observaciones='"& observaciones &"', @p_ninventario='"& ninventario &"'")
            
            if rstAux.state<>0 then
			    if not rstAux.Eof then
				    ninventario= rstAux("ninventario")
			    else
				    Response.Write(LitProcVacio)
				    Response.End
			    end if
			else
			    Response.Write(LitProcVacio)
			    Response.End
			end if
			rstAux.close
			%><input type="hidden" name="hninventario" value="<%=enc.EncodeForHtmlAttribute(ninventario)%>"><%
			rstAux.Open "delete from regularizaciones_pendientes with(rowlock) where nregularizacion like '"&session("ncliente")&"%' and almacen='"&CargaPendiente&"'", session("dsn_cliente")
	        strdrop2 ="delete from INVENTARIOS_TEMP with(rowlock) where usuario = '" & session("usuario") &"'  and ncliente='"&session("ncliente")&"' " 
            rstAux.open strdrop2,session("dsn_cliente")    

			Imprime
		end if
		set command=nothing
		set conn=nothing
	elseif mode="regularizacion_recarga" then
		ninventario=limpiacadena(request.querystring("ninventario"))
		Imprime
   end if%>
</form>
<%set rstAux = nothing
set rst = nothing
set rstSelect = nothing
set rstDet = nothing
end if%>
</body>
</html>