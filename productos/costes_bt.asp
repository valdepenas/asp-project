<%@ Language=VBScript %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="costes.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">

</head>
<script language="JavaScript" src="../jfunciones.js"></script>
<script language="JavaScript" type="text/javascript">
function Accion(accion)
{
	switch(accion)
	{
		case "imprimir": //Imprimir Listado
			parent.pantalla.focus();
			parent.pantalla.print();
			break;

		case "volver": //Volver a la pantalla de costes
			parent.pantalla.document.location= "costesn.asp?mode=param";
			document.location= "costes_bt.asp?mode=param";
			break;
	}
}
</script>
<body class="body_master_ASP">
<%	
	if Request.QueryString("mode")="impresion" then%>
    <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
	        <table id="BUTTONS_CENTER_ASP">
		        <tr>
	
		    <td class="CELDABOT" onclick="javascript:Accion('imprimir');">
			    <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,""%>
		    </td>
		    <td class="CELDABOT" onclick="javascript:Accion('volver');">
			    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,""%>
		    </td>
		</tr>
	</table>
     </div>
    </div>
    <%end if
ImprimirPie_bt%>
</body>
</html>