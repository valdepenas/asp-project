<%@ Language=VBScript %>
<%
'' IML : 27/11/03 : Control de Impresion (controlimpresion.inc)
%>
<% 
dim  enc
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
<title><%=LitTituloRVT%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<!--#include file="../../calculos.inc" -->
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../tickets.inc" -->
<!--#include file="../../styles/Master.css.inc" -->
<script language="javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1);

function cambiarfecha(fecha,modo){
	var fecha_ar=new Array();

	if (fecha!=""){
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length){
			if (fecha.substring(l,l+1)=='/'){
				suma++;
				fecha_ar[suma]="";
			}
			else{
				if (fecha.substring(l,l+1)!='') fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
			}
			l++;
		}
		if (suma!=2) {
			window.alert("<%=LitFechaMalIntro%> " + modo );
			return false;
		}
		else {
			nonumero=0;
			while (suma>=0 && nonumero==0){
				if (isNaN(fecha_ar[suma])) nonumero=1;
				if (fecha_ar[suma].length>2 && suma!=2) nonumero=1;
				if (fecha_ar[suma].length>4 && suma==2) nonumero=1;
				suma--;
			}

			if (nonumero==1){
				window.alert("<%=LitFechaMalIntro%> " + modo);
				return false;
			}
		}
	}
	return true;
}

function Imprimir() {
	if (pr) //NS4, IE5
		parent.pantalla.print()
	else if (da && !mac) // IE4 (Windows)
		alert("<%=LitNoImprime%>");
	else // Otros Navegadores
		alert("<%=LitNoImprime%>");
	return false;
}

//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if(parent.pantalla.document.listado_tickets.dfecha.value=="" || parent.pantalla.document.listado_tickets.hfecha.value==""){
		window.alert("<%=LitDebeExistirPeriodoFechas%>");
		return false;
	}

	if (!cambiarfecha(parent.pantalla.document.listado_tickets.dfecha.value,"DESDE FECHA")){
		return false;
	}

	if(!checkdate(parent.pantalla.document.listado_tickets.dfecha)){
		window.alert("<%=LitFechaMalIntro%>" +" "+ "<%=LitDesdeFecha %>");
		return false;
	}

	if (!cambiarfecha(parent.pantalla.document.listado_tickets.hfecha.value,"HASTA FECHA")){
		return false;
	}

	if(!checkdate(parent.pantalla.document.listado_tickets.hfecha)){
		window.alert("<%=LitFechaMalIntro%>" +" "+ "<%=LitHastaFecha %>");
		return false;
	}
	
	if ((parent.pantalla.document.listado_tickets.dhora.value != "") && !checkhora(parent.pantalla.document.listado_tickets.dhora)){
	    window.alert("<%=LitHoraMalIntro%>"+" "+"<%=LitDHora %>");
	    return false;
	}
	
	if ((parent.pantalla.document.listado_tickets.hhora.value != "") && !checkhora(parent.pantalla.document.listado_tickets.hhora)){
	    window.alert("<%=LitHoraMalIntro%>"+" "+"<%=LitHHora %>");
	    return false;
	}
	
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "param":
			switch (pulsado) {
				case "select": //Aceptar
					if (ValidarCampos())
					{
						parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						cadena="listado_tickets_comp.asp?mode=imp" +
						"&operador=" + parent.pantalla.document.listado_tickets.operador.value +
						"&dfecha=" + parent.pantalla.document.listado_tickets.dfecha.value +
						"&hfecha=" + parent.pantalla.document.listado_tickets.hfecha.value +
						"&serie=" + parent.pantalla.document.listado_tickets.serie.value +
						"&tienda=" + parent.pantalla.document.listado_tickets.tienda.value +
						"&caja=" + parent.pantalla.document.listado_tickets.caja.value +
						"&tpv=" + parent.pantalla.document.listado_tickets.tpv.value +
						"&mediopago=" + parent.pantalla.document.listado_tickets.mediopago.value +
						"&tipoperador=" + parent.pantalla.document.listado_tickets.tipoperador.value +
						"&conref=" + parent.pantalla.document.listado_tickets.conref.value +
						"&connombre=" + parent.pantalla.document.listado_tickets.connombre.value +
						"&tiparticulo=" + parent.pantalla.document.listado_tickets.tiparticulo.value +
						"&familia=" + parent.pantalla.document.listado_tickets.familia.value +
						"&mostrarcolum=" + parent.pantalla.document.listado_tickets.mostrarcolum.value;
						parent.pantalla.fr_ComprLimitColum.location=cadena;
					}
					break;
			}
			break;
		case "imp":
			switch (pulsado) {
				case "cancel": //Volver atrás
					dtofam=parent.pantalla.document.listado_ticketsResultado.dtofam.value;
					verp=parent.pantalla.document.listado_ticketsResultado.verp.value;
					vert=parent.pantalla.document.listado_ticketsResultado.vert.value;
					verc=parent.pantalla.document.listado_ticketsResultado.verc.value;
					caju=parent.pantalla.document.listado_ticketsResultado.caju.value;
					i=parent.pantalla.document.listado_ticketsResultado.i.value;

					cadena="listado_tickets.asp?mode=param"
					cadena=cadena + "&dtofam=" + dtofam
					cadena=cadena + "&verp=" + verp
					cadena=cadena + "&vert=" + vert
					cadena=cadena + "&verc=" + verc
					cadena=cadena + "&caju=" + caju
					cadena=cadena + "&i=" + i
					parent.pantalla.document.listado_ticketsResultado.action=cadena;
					parent.pantalla.document.listado_ticketsResultado.submit();
					document.location="listado_tickets_bt.asp?mode=param";
					break;
				case "imprimir": //Volver atrás
					parent.pantalla.focus();
					parent.pantalla.print();
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.listado_ticketsResultado.NumRegs.value)>=parseInt(parent.pantalla.document.listado_ticketsResultado.maxpdf.value))
						alert("<%=LitLimitePDF%>");
					else {
					    parent.pantalla.document.listado_ticketsResultado.action = "listado_tickets_pdf.asp?mode=browse&xls=0";
						parent.pantalla.document.listado_ticketsResultado.submit();

						dtofam=parent.pantalla.document.listado_ticketsResultado.dtofam.value;
						verp=parent.pantalla.document.listado_ticketsResultado.verp.value;
						vert=parent.pantalla.document.listado_ticketsResultado.vert.value;
						verc=parent.pantalla.document.listado_ticketsResultado.verc.value;
						caju=parent.pantalla.document.listado_ticketsResultado.caju.value;
						i=parent.pantalla.document.listado_ticketsResultado.i.value;

						cadena="listado_tickets_bt.asp?mode=pdf"
						cadena=cadena + "&dtofam=" + dtofam
						cadena=cadena + "&verc=" + verc;
						cadena=cadena + "&vert=" + vert;
						cadena=cadena + "&verp=" + verp;
						cadena=cadena + "&caju=" + caju;
						cadena=cadena + "&i=" + i;
						document.location=cadena;
					}
					break;
			    case "exportar":
			        cadena = "listado_tickets_pdf.asp?mode=browse&xls=1" +
					"&operador=" + parent.pantalla.document.listado_ticketsResultado.operador.value +
					"&dfecha=" + parent.pantalla.document.listado_ticketsResultado.dfecha.value +
					"&hfecha=" + parent.pantalla.document.listado_ticketsResultado.hfecha.value +
					"&serie=" + parent.pantalla.document.listado_ticketsResultado.serie.value +
					"&tienda=" + parent.pantalla.document.listado_ticketsResultado.tienda.value +
					"&caja=" + parent.pantalla.document.listado_ticketsResultado.caja.value +
					"&tpv=" + parent.pantalla.document.listado_ticketsResultado.tpv.value +
					"&mediopago=" + parent.pantalla.document.listado_ticketsResultado.mediopago.value +
					"&tipoperador=" + parent.pantalla.document.listado_ticketsResultado.tipoperador.value +
                    "&agruparpor=" + parent.pantalla.document.listado_ticketsResultado.agruparpor.value +
					"&conref=" + parent.pantalla.document.listado_ticketsResultado.conref.value +
					"&connombre=" + parent.pantalla.document.listado_ticketsResultado.connombre.value +
					"&tiparticulo=" + parent.pantalla.document.listado_ticketsResultado.tiparticulo.value +
					"&familia=" + parent.pantalla.document.listado_ticketsResultado.familia.value +
					"&mostrarcolum=" + parent.pantalla.document.listado_ticketsResultado.mostrarcolum.value +
			        "&mostrardesc=" + parent.pantalla.document.listado_ticketsResultado.mostrardesc.value +
			        "&apaisado=" + parent.pantalla.document.listado_ticketsResultado.apaisado.value +
			        "&mostrarinfo=" + parent.pantalla.document.listado_ticketsResultado.mostrarinfo.value +
			        "&dtofam=" + parent.pantalla.document.listado_ticketsResultado.dtofam.value +
			        "&verp=" + parent.pantalla.document.listado_ticketsResultado.verp.value +
			        "&vert=" + parent.pantalla.document.listado_ticketsResultado.vert.value +
                    "&verc=" + parent.pantalla.document.listado_ticketsResultado.verc.value +
                    "&caju=" + parent.pantalla.document.listado_ticketsResultado.caju.value +
                    "&i=" + parent.pantalla.document.listado_ticketsResultado.i.value +
                    "&familia_padre=" + parent.pantalla.document.listado_ticketsResultado.familia_padre.value +
                    "&categoria=" + parent.pantalla.document.listado_ticketsResultado.categoria.value +
                    "&dhora=" + parent.pantalla.document.listado_ticketsResultado.dhora.value +
                    "&hhora=" + parent.pantalla.document.listado_ticketsResultado.hhora.value +
                    "&opcprecioiva=" + parent.pantalla.document.listado_ticketsResultado.opcprecioiva.value +
			        "&detalle=" + parent.pantalla.document.listado_ticketsResultado.detalle.value;
			        parent.pantalla.frameExport.document.location = cadena;
			        break;
			}
			break;

		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
					cadena="listado_tickets.asp&mode=param"
					cadena=cadena + "&dtofam=" + document.opciones.dtofam.value
					cadena=cadena + "&verc=" + document.opciones.verc.value
					cadena=cadena + "&vert=" + document.opciones.vert.value
					cadena=cadena + "&verp=" + document.opciones.verp.value;
					cadena=cadena + "&caju=" + document.opciones.caju.value;
					cadena=cadena + "&i=" + document.opciones.i.value;
					parent.document.location="../../central.asp?pag1=ventas/listados/" + cadena + "&pag2=ventas/listados/listado_tickets_bt.asp";
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%if request.querystring("dtofam")>"" then
	dtofam=limpiaCadena(request.querystring("dtofam"))
else
	dtofam=request.Form("dtofam")
end if

if request.querystring("verp")>"" then
	verp=limpiaCadena(request.querystring("verp"))
else
	verp=request.Form("verp")
end if

if request.querystring("vert")>"" then
	vert=limpiaCadena(request.querystring("vert"))
else
	vert=request.Form("vert")
end if

if request.querystring("verc")>"" then
	verc=limpiaCadena(request.querystring("verc"))
else
	verc=request.Form("verc")
end if

if request.querystring("caju")>"" then
	caju=limpiaCadena(request.querystring("caju"))
else
	caju=request.Form("caju")
end if

if request.querystring("i")>"" then
	coniva=limpiaCadena(request.querystring("i"))
else
	coniva=request.Form("i")
end if

mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
	<div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
		            <%if mode="param" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('param','select');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
			        <%elseif mode="imp" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					        <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				        </td>
                        <td class="CELDABOT" onclick="javascript:Accion('imp','exportar');">
				            <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			            </td>
			            <td class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
				            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			            </td>
			        <%elseif mode="pdf" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				        </td>
			        <%end if%>
		        </tr>
	        </table>                                                                                 
        </div>
    </div>
	<input type="hidden" name="dtofam" value="<%=EncodeForHtml(dtofam)%>" />
	<input type="hidden" name="verp" value="<%=EncodeForHtml(verp)%>" />
	<input type="hidden" name="vert" value="<%=EncodeForHtml(vert)%>" />
	<input type="hidden" name="verc" value="<%=EncodeForHtml(verc)%>" />
	<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>" />
	<input type="hidden" name="i" value="<%=EncodeForHtml(coniva)%>" />
	<table style="width:100%;height:42px;vertical-align:bottom;" align="center">
        <tr>
            <td style="width:100%;height:42px; vertical-align:bottom; text-align:center;">
                <%ImprimirPie_bt%>
            </td>
        </tr>
    </table>
</form>
</body>
</html>