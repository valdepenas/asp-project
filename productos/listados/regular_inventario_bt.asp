<%@ Language=VBScript %>
<%'------------------------------CODIGOS DE AÑADIDURAS/MODIFICACIONES ------------------------
'JCI-090103-01 : He vuelto a añadir lo del total, ya que al quitarlo
'                salía 0 como total general en el PDF para cualquier agrupacion
'	FECHA :09/01/03
' AUTOR :JCI
'----------------------------------------------------------------------------------------------
'JCI 03/04/2003 : Control de caché y objeto de impresión
'' IML : 27/11/03 : Control de Impresion (controlimpresion.inc)
%>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloRegInv%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="../reg_inventario.inc" -->

<!--#include file="../../styles/Master.css.inc" -->

<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
//Validación de campos numéricos y fechas.
function ValidarCampos()
{
	if (!checkdate(parent.pantalla.document.regular_inventario.fdesde)) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.regular_inventario.fhasta)) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}

	if (parent.pantalla.document.regular_inventario.fdesde.value=="" && parent.pantalla.document.regular_inventario.fhasta.value=="") {
		window.alert("<%=LitMsgFechasNulas%>");
		return false;
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
					parent.document.location="../../central.asp?pag1=productos/listados/regular_inventario.asp&mode=select1&pag2=productos/listados/regular_inventario_bt.asp";
					break;
			}
			break;

		case "browse":
			switch (pulsado) {
				case "imprimir": //Imprimir Listado
					parent.pantalla.focus();
					parent.pantalla.print();
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.regular_inventarioResultado.NumRegsTotal.value)>=parseInt(parent.pantalla.document.regular_inventarioResultado.maxpdf.value))
						alert("<%=LitMsgDemReg%>");
					else {
						parent.pantalla.document.regular_inventarioResultado.action="regular_inventario_pdf.asp?mode=browse&xls=0";
						parent.pantalla.document.regular_inventarioResultado.submit();
						document.location="regular_inventario_bt.asp?mode=pdf";
					}
					break;
				case "cancelar": //Cancelar operacion
					parent.pantalla.document.location="regular_inventario.asp?mode=select1";					
					document.location="regular_inventario_bt.asp?mode=select1";
					break;
			    case "exportar":
			        cadena = "";
			        cadena = cadena + "&fdesde=" + parent.pantalla.document.regular_inventarioResultado.fdesde.value;
			        cadena = cadena + "&fhasta=" + parent.pantalla.document.regular_inventarioResultado.fhasta.value;
			        cadena = cadena + "&almacen=" + parent.pantalla.document.regular_inventarioResultado.almacen.value;
			        cadena = cadena + "&comercial=" + parent.pantalla.document.regular_inventarioResultado.comercial.value;
			        cadena = cadena + "&numDocum=" + parent.pantalla.document.regular_inventarioResultado.numDocum.value;
			        cadena = cadena + "&referencia=" + parent.pantalla.document.regular_inventarioResultado.referencia.value;
			        cadena = cadena + "&categoria=" + parent.pantalla.document.regular_inventarioResultado.categoria.value;
			        cadena = cadena + "&familia_padre=" + parent.pantalla.document.regular_inventarioResultado.familia_padre.value;
			        cadena = cadena + "&familia=" + parent.pantalla.document.regular_inventarioResultado.familia.value;
			        cadena = cadena + "&nombre=" + parent.pantalla.document.regular_inventarioResultado.nombre.value;
			        cadena = cadena + "&agrupar=" + parent.pantalla.document.regular_inventarioResultado.agrupar.value;
			        cadena = cadena + "&detalle=" + parent.pantalla.document.regular_inventarioResultado.detalle.value;
			        cadena = cadena + "&verCostes=" + parent.pantalla.document.regular_inventarioResultado.verCostes.value;
			        cadena = cadena + "&verSolDesv=" + parent.pantalla.document.regular_inventarioResultado.verSolDesv.value;
			        cadena = cadena + "&ordenar=" + parent.pantalla.document.regular_inventarioResultado.ordenar.value;

			        parent.pantalla.frameExportar.document.location = "regular_inventario_pdf.asp?mode=browse&xls=1" + cadena;
			        break;
			}
			break;

		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos()) {
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
						parent.pantalla.document.regular_inventario.action="regular_inventarioResultado.asp?mode=browse&confirma=NO&save=true";
						parent.pantalla.document.regular_inventario.submit();
						document.location="regular_inventario_bt.asp?mode=browse";
					}
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
		            <%if mode="browse" then%>
				        <td id="idprint" class="CELDABOT" onclick="javascript:Accion('browse','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <td id="idprintlist" class="CELDABOT" onclick="javascript:Accion('browse','imprimirp');">
					        <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				        </td>
                        <td id="idexcel" class="CELDABOT" onclick="javascript:Accion('browse','exportar');">
				            <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			            </td>
			            <td id="idreturn" class="CELDABOT" onclick="javascript:Accion('browse','cancelar');">
				            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			            </td>
			        <%elseif mode="select1" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('add','save');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
			        <%elseif mode="pdf" then%>
				        <td id="idreturn" class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				        </td>
			        <%end if%>
		        </tr>
	        </table>
        </div>
    </div>
    <table style="width:100%;height:30px;vertical-align:bottom;" align="center">
        <tr>
            <td style="width:100%;height:30px; vertical-align:bottom; text-align:center;">
                <%ImprimirPie_bt%>
            </td>
        </tr>
    </table>
</form>
</body>
</html>