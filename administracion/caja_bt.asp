<%@ Language=VBScript %>
<%'' JCI 23/06/2003 : MIGRACION A MONOBASE%>
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
<!--#include file="../adovbs.inc" -->

<!--#include file="Ahoja_gastos.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
function Buscar()
 {
	ok=1;
	if (!checkdate(parent.pantalla.document.caja.fdesde) || (parent.pantalla.document.caja.fdesde.value=="")) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		ok=0;
	}
	if (ok==1 && !checkdate(parent.pantalla.document.caja.fhasta) || (parent.pantalla.document.caja.fhasta.value=="")) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		ok=0;
	}
	if (ok==1 && parent.pantalla.document.caja.ncaja.value=="") {
		window.alert("<%=LitMsgCajaNoNulo%>");
		ok=0;
	}
	if (ok==1)
	{
		permc=parent.pantalla.document.caja.permc.value;
		if (permc=="SI") pagina="detalles_caja_mod.asp";
		else pagina="detalles_caja.asp";
		parent.pantalla.document.getElementById("frEntradas").src=pagina + "?mode=entradas&submode=search&campo=" + document.opciones.campos.value +
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&ncaja=" + parent.pantalla.document.caja.ncaja.value
		+ "&fdesde=" + parent.pantalla.document.caja.fdesde.value + "&fhasta=" + parent.pantalla.document.caja.fhasta.value + "&metalico=" + parent.pantalla.document.caja.metalico.checked + "&nometalico=" + parent.pantalla.document.caja.nometalico.checked;
		;
		parent.pantalla.document.getElementById("frSalidas").src=pagina + "?mode=salidas&submode=search&campo=" + document.opciones.campos.value +
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&ncaja=" + parent.pantalla.document.caja.ncaja.value
		+ "&fdesde=" + parent.pantalla.document.caja.fdesde.value + "&fhasta=" + parent.pantalla.document.caja.fhasta.value + "&metalico=" + parent.pantalla.document.caja.metalico.checked + "&nometalico=" + parent.pantalla.document.caja.nometalico.checked;
		;
	}
}
//****************************************************************************************************************
function comprobar_enter(e)
{
	document.opciones.criterio.focus();
	Buscar();
}
</script>

<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
<div id="PageFooter_ASP" >
    <div id="FILTERS_MASTER_ASP">
		<select class="IN_S" name="campos">
         	<option selected value="c.descripcion"><%=LitDescripcion%></option>
			<option value="ndocumento"><%=LitDocumento%></option>
        </select>
		<select class="IN_S" name="criterio">
			<option value="contiene"><%=LitContiene%></option>
			<option value="termina"><%=LitTermina%></option>
			<option value="igual"><%=LitIgual%></option>
		</select>
		<input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
		<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>"/></a>
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