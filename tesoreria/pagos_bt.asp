<%@ Language=VBScript %>
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="pagos.inc" -->

<!--#include file="../styles/Master.css.inc" -->
<!--#include file="../styles/FootButton.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if (parent.pantalla.document.pagos.fdesde.value=="") {
		window.alert("<%=LitMsgDesdeFechaNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.pagos.fhasta.value=="") {
		window.alert("<%=LitMsgHastaFechaNoNulo%>");
		return false;
	}
    if (!checkdate(parent.pantalla.document.pagos.fdesde)) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.pagos.fhasta)) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}
	return true;
}

function ValidarCampos2()
{
	if (parent.pantalla.document.pagos.ncaja.value=="")
	{
		if (!window.confirm("<%=LitMsgPagoSinCajaConfirm%>")) return false;
	}

	if ((parent.pantalla.document.pagos.ncaja.value!="") && (parent.pantalla.document.pagos.i_pago.value=="")) {
		window.alert("<%=LitMsgTipoPagoNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.pagos.fechapago.value=="") {
		window.alert("<%=LitMsgFechaPagoNoNulo%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.pagos.fechapago)) {
		window.alert("<%=LitMsgFechaMal%>");
		return false;
	}
	if (!window.confirm(parent.pantalla.document.pagos.mensaje.value)) return false;

	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "select1":
			switch (pulsado) {
				case "select2": //Aceptar
					if (ValidarCampos()) {
						parent.pantalla.document.pagos.action="pagosResultado.asp?mode=" + pulsado;
						parent.pantalla.document.pagos.submit();
						document.location="pagos_bt.asp?mode=" + pulsado;
					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.pagos.action="pagos.asp?mode=" + pulsado;
					parent.pantalla.document.pagos.submit();
					document.location="pagos_bt.asp?mode=" + pulsado;
					break;
			}
			break;
		case "select2":
			switch (pulsado) {
				case "todos": //Seleccionar todos los registros
					nregistros=parent.pantalla.document.pagos.h_nfilas.value;
					var totalImportePagar=0.00;
					for (i=1;i<=nregistros;i++) {
						nombre="check" + i;
						parent.pantalla.document.pagos.elements[nombre].checked=true;
						//FLM:20090529:Mostramos el total del importe seleccionado en la remesa.
						totalImportePagar+=parseFloat(Redondear(parseFloat(parent.pantalla.document.pagos.elements["imp"+i].value.replace(",","."))*parseFloat(parent.pantalla.document.pagos.elements["factcambio"+i].value.replace(",",".")),parent.pantalla.numDecimalesEmpresa).replace(".","").replace(",","."));
					}
					//FLM:20090529:Mostramos el total del importe seleccionado en la remesa.
					parent.pantalla.totalImportePagar=totalImportePagar;
	                parent.pantalla.document.getElementById("totalAPagar").innerHTML=totalImportePagar.toFixed(parent.pantalla.numDecimalesEmpresa);					
					break;

				case "ninguno": //No seleccionar ningun registro
					nregistros=parent.pantalla.document.pagos.h_nfilas.value;
					for (i=1;i<=nregistros;i++) {
						nombre="check" + i;
						parent.pantalla.document.pagos.elements[nombre].checked=false;
					}
					//FLM:20090529:Mostramos el total del importe seleccionado en la remesa.
					parent.pantalla.totalImportePagar=0;
	                parent.pantalla.document.getElementById("totalAPagar").innerHTML="0.00";
					break;

				case "confirm": //Aceptar
					if (ValidarCampos2())
					{
						parent.pantalla.document.pagos.action="pagos.asp?mode=confirm";
						parent.pantalla.document.pagos.submit();
						document.location="pagos_bt.asp?mode=select1";
					}
					break;

				case "select1": //Cancelar
					parent.pantalla.document.pagos.action="pagos.asp?mode=select1";
					parent.pantalla.document.pagos.submit();
					document.location="pagos_bt.asp?mode=select1";
					break;
			}
			break;
	}
}
</script>

<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
<div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_ASP" >
        <table id="BUTTONS_CENTER_ASP" >
		<tr>
		    <%if mode="select1" then
		        %>
				<td id="idSelectDocument" class="CELDABOT" onclick="javascript:Accion('select1','select2');">
					<%PintarBotonBT LITBOTSELDOCU,ImgSelecc_doc,ParamImgSelecc_doc,LITBOTSELDOCUTITLE%>
				</td>
			<%elseif mode="select2" then
			    %>
				<td id="idSelectAll" class="CELDABOT" onclick="javascript:Accion('select2','todos');">
					<%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,LITBOTSELTODOTITLE%>
				</td>
				<td id="idSelectNothing" class="CELDABOT" onclick="javascript:Accion('select2','ninguno');">
					<%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,LITBOTDSELTODOTITLE%>
				</td>
				<td id="idaccept" class="CELDABOT" id="boton_cobrar" onclick="javascript:Accion('select2','confirm');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('select2','select1');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
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
    </table></form>
</body>
</html>