<%@ Language=VBScript %>
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloListMov%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="../movimientos_almacenes.inc" -->

<!--#include file="../../styles/Master.css.inc" -->

<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
//Validación de campos numéricos y fechas.
    function ValidarCampos() {
	if (!checkdate(parent.pantalla.document.lista_movimientos_almacenes.Dfecha)) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (parent.pantalla.document.lista_movimientos_almacenes.Dfecha.value=="") {
		window.alert("<%=LitDesdeFechaNoNulo%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.lista_movimientos_almacenes.Hfecha)) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}
	if (parent.pantalla.document.lista_movimientos_almacenes.Hfecha.value=="") {
		window.alert("<%=LitHastaFechaNoNulo%>");
		return false;
    }

	return true;
}

    

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "param":
			switch (pulsado) {
				case "imp": //Aceptar
					if (ValidarCampos()) {
						parent.pantalla.document.lista_movimientos_almacenes.action="lista_movimientos_almacenesResultado.asp?mode=" + pulsado;
						parent.pantalla.document.lista_movimientos_almacenes.submit();
						document.location="lista_movimientos_almacenes_bt.asp?mode=" + pulsado;
					}
					break;
			}
			break;
		case "imp":
			switch (pulsado) {
				case "cancel": //Volver atrás
					parent.pantalla.document.lista_movimientos_almacenesResultado.action="lista_movimientos_almacenes.asp?mode=param";
					parent.pantalla.document.lista_movimientos_almacenesResultado.submit();
					document.location="lista_movimientos_almacenes_bt.asp?mode=param";
					break;
				case "imprimir": //Volver atrás
					parent.pantalla.focus();
					parent.pantalla.print();
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.lista_movimientos_almacenesResultado.NumRegs.value)>=parseInt(parent.pantalla.document.lista_movimientos_almacenesResultado.maxpdf.value))
						alert("<%=LitDemReg%>");
					else {
						parent.pantalla.document.lista_movimientos_almacenesResultado.action="lista_movimientos_almacenes_pdf.asp?mode=browse";
						parent.pantalla.document.lista_movimientos_almacenesResultado.submit();
						document.location="lista_movimientos_almacenes_bt.asp?mode=pdf";
					}
					break;
			}
			break;
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
				    parent.document.location="../../central.asp?pag1=productos/listados/lista_movimientos_almacenes.asp&pag2=productos/listados/lista_movimientos_almacenes_bt.asp&mode=param";
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
		            <%if mode="param" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('param','imp');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
			        <%elseif mode="imp" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					        <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
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