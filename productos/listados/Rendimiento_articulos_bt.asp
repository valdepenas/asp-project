<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../styles/Master.css.inc" -->
<!--#include file="rendimiento_articulos.inc" -->
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1);

function Imprimir() {
	if (pr) //NS4, IE5
		parent.pantalla.print()
		//vbImprimir()
	else if (da && !mac) // IE4 (Windows)
		//vbImprimir()
		alert("<%=LitNoImprime%>");
	else // Otros Navegadores
		alert("<%=LitNoImprime%>");
	return false;
}

function printWindowAnt() {
	factory.printing.header = "<%=LitRendimientoArticulos%>. -- Registros : " + parent.pantalla.document.rendimiento_articulos.NumRegsTotal.value + "&bPágina &p de &P"
	factory.printing.footer = "<%=PieListados%>&bFecha : &d"
	factory.printing.portrait = false
	factory.printing.leftMargin = 19.0
	factory.printing.topMargin = 19.0
	factory.printing.rightMargin = 19.0
	factory.printing.bottomMargin = 19.0
	factory.printing.Print(false, parent.pantalla)
}

//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if (!checkdate(parent.pantalla.document.rendimiento_articulos.fdesde) || (parent.pantalla.document.rendimiento_articulos.fdesde.value=="")) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.rendimiento_articulos.fhasta) || (parent.pantalla.document.rendimiento_articulos.fhasta.value=="")) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}
	if (parent.pantalla.document.rendimiento_articulos.ncliente.value!="" && parent.pantalla.document.rendimiento_articulos.nombre.value==""){
		window.alert("<%=LitMsgClienteNoExiste%>");
		return false;
	}
	if (parent.pantalla.document.rendimiento_articulos.nproveedor.value!="" && parent.pantalla.document.rendimiento_articulos.nomproveedor.value==""){
		window.alert("<%=LitMsgProveedorNoExiste%>");
		return false;
	}

    var cuantas_series_facturas = 0;
    for (var i=0; i<parent.pantalla.document.rendimiento_articulos.seriesFactura.length; i++) {
        if (parent.pantalla.document.rendimiento_articulos.seriesFactura[i].selected) {
            if (parent.pantalla.document.rendimiento_articulos.seriesFactura[i].value!="") cuantas_series_facturas += 1;
        }
    }
    var cuantas_series_albaranes = 0;
    for (var i=0; i<parent.pantalla.document.rendimiento_articulos.seriesAlbaranes.length; i++) {
        if (parent.pantalla.document.rendimiento_articulos.seriesAlbaranes[i].selected) {
            if (parent.pantalla.document.rendimiento_articulos.seriesAlbaranes[i].value!="") cuantas_series_albaranes += 1;
        }
    }
    var cuantas_series_tickets = 0;
    if (parent.pantalla.document.rendimiento_articulos.tiene_t.value==1)
    {
        for (var i=0; i<parent.pantalla.document.rendimiento_articulos.selecttickets.length; i++) {
            if (parent.pantalla.document.rendimiento_articulos.selecttickets[i].selected) {
                if (parent.pantalla.document.rendimiento_articulos.selecttickets[i].value!="") cuantas_series_tickets += 1;
            }
        }
    }
	if (cuantas_series_facturas==0 && cuantas_series_albaranes==0 && cuantas_series_tickets==0)
	{
		window.alert("<%=LitNoHaSelNinSerie%>");
		return false;
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode, pulsado) {
	switch (mode) {
		case "select1":
			switch (pulsado) {
				case "imp": //Aceptar
					if (ValidarCampos()) {
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
						if (parent.pantalla.document.rendimiento_articulos.h_rendimiento.value=="articulos") {
							parent.pantalla.document.rendimiento_articulos.action="rendimiento_articulosResultado.asp?mode=" + pulsado;
							parent.pantalla.document.rendimiento_articulos.submit();
							document.location="rendimiento_articulos_bt.asp?mode=" + pulsado;
						} else {
							parent.pantalla.document.rendimiento_articulos.action="rendimiento_documentosResultado.asp?mode=" + pulsado;
							parent.pantalla.document.rendimiento_articulos.submit();
							document.location="rendimiento_documentos_bt.asp?mode=" + pulsado;
						}

					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.rendimiento_articulos.action="rendimiento_articulos.asp?mode=" + pulsado;
					parent.pantalla.document.rendimiento_articulos.submit();
					document.location="rendimiento_articulos_bt.asp?mode=" + pulsado;
					break;
			}
			break;
		case "imp":
			switch (pulsado) {
				case "cancel": //Volver atrás
					parent.pantalla.document.location="rendimiento_articulos.asp?mode=select1";
					document.location="rendimiento_articulos_bt.asp?mode=select1";
					break;
				case "imprimir": //Volver atrás
					parent.pantalla.focus();
					//printWindow();
                    parent.pantalla.print();
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.rendimiento_articulosResultado.NumRegsTotal.value)>=parseInt(parent.pantalla.document.rendimiento_articulosResultado.maxpdf.value))
						alert("<%=LitMsgDemReg%>");
					else {
					    parent.pantalla.document.rendimiento_articulosResultado.action = "rendimiento_articulos_pdf.asp?mode=browse&xls=0";
						parent.pantalla.document.rendimiento_articulosResultado.submit();
						document.location="rendimiento_articulos_bt.asp?mode=pdf";
					}
					break;
			    case "exportar":
			        cadena = "";
			        cadena = cadena + "&ncliente=" + parent.pantalla.document.rendimiento_articulosResultado.ncliente.value;
			        cadena = cadena + "&ncenter=" + parent.pantalla.document.rendimiento_articulosResultado.ncenter.value;
			        cadena = cadena + "&actividad=" + parent.pantalla.document.rendimiento_articulosResultado.actividad.value;
			        cadena = cadena + "&nodetallado=" + parent.pantalla.document.rendimiento_articulosResultado.nodetallado.value;
			        cadena = cadena + "&nproveedor=" + parent.pantalla.document.rendimiento_articulosResultado.nproveedor.value;
			        cadena = cadena + "&fdesde=" + parent.pantalla.document.rendimiento_articulosResultado.fdesde.value;
			        cadena = cadena + "&fhasta=" + parent.pantalla.document.rendimiento_articulosResultado.fhasta.value;
			        cadena = cadena + "&nserie=" + parent.pantalla.document.rendimiento_articulosResultado.nserie.value;
			        cadena = cadena + "&tipoCliente=" + parent.pantalla.document.rendimiento_articulosResultado.tipoCliente.value;
			        cadena = cadena + "&cod_proyecto=" + parent.pantalla.document.rendimiento_articulosResultado.cod_proyecto.value;
			        cadena = cadena + "&familia=" + parent.pantalla.document.rendimiento_articulosResultado.familia.value;
			        cadena = cadena + "&familia_padre=" + parent.pantalla.document.rendimiento_articulosResultado.familia_padre.value;
			        cadena = cadena + "&categoria=" + parent.pantalla.document.rendimiento_articulosResultado.categoria.value;
			        cadena = cadena + "&verCodCFS=" + parent.pantalla.document.rendimiento_articulosResultado.verCodCFS.value;
			        cadena = cadena + "&comercial=" + parent.pantalla.document.rendimiento_articulosResultado.comercial.value;
			        cadena = cadena + "&referencia=" + parent.pantalla.document.rendimiento_articulosResultado.referencia.value;
			        cadena = cadena + "&nombreart=" + parent.pantalla.document.rendimiento_articulosResultado.nombreart.value;
			        cadena = cadena + "&coste=" + parent.pantalla.document.rendimiento_articulosResultado.coste.value;
			        cadena = cadena + "&ordenar=" + parent.pantalla.document.rendimiento_articulosResultado.ordenar.value;
			        cadena = cadena + "&ordenardoc=" + parent.pantalla.document.rendimiento_articulosResultado.ordenardoc.value;
			        cadena = cadena + "&artbaja=" + parent.pantalla.document.rendimiento_articulosResultado.artbaja.value;
			        cadena = cadena + "&pedsinf=" + parent.pantalla.document.rendimiento_articulosResultado.pedsinf.value;
			        cadena = cadena + "&calcost=" + parent.pantalla.document.rendimiento_articulosResultado.calcost.value;
			        cadena = cadena + "&agruparart=" + parent.pantalla.document.rendimiento_articulosResultado.agruparart.value;
			        cadena = cadena + "&agrupardoc=" + parent.pantalla.document.rendimiento_articulosResultado.agrupardoc.value;
			        cadena = cadena + "&tipoarticulo=" + parent.pantalla.document.rendimiento_articulosResultado.tipoarticulo.value;

			        parent.pantalla.frameExportar.document.location = "rendimiento_articulos_pdf.asp?mode=browse&xls=1" + cadena;
			        break;
			}
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
				    parent.document.location="../../central.asp?pag1=productos/listados/rendimiento_articulos.asp&mode=select1&pag2=productos/listados/rendimiento_articulos_bt.asp";
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%
mode=enc.EncodeForJavascript(Request.QueryString("mode"))%>
<form name="opciones" method="post">
<div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
	        <table id="BUTTONS_CENTER_ASP">
	
		<tr>
		    <%if mode="select1" then%>
				<td id="idaccept" class="CELDABOT" onclick="javascript:Accion('select1','imp');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('select1','select1');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				</td>
			<%elseif mode="imp" then%>
				<td id="idprint" class="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					<%PintarBotonBT LITBOTIMPRIMIRPAG,ImgImprimir_pag,ParamImgImprimir_pag,""%>
				</td>
				<td id="idprintlist" class="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					<%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,""%>
				</td>
                <td id="idexcel" class="CELDABOT" onclick="javascript:Accion('imp','exportar');">
				    <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			    </td>
			    <td id="idreturn" class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
				    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,""%>
			    </td>
			<%elseif mode="pdf" then%>
				<td id="idreturn" class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					<%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,""%>
				</td>
			<%end if%>
		</tr>
	</table>
    </div>
    </div>
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