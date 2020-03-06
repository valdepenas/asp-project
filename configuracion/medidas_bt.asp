<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->

<TITLE><%=LitTituloUM%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
function Buscar() {
	parent.pantalla.fr_Tabla.document.medidas_det.texto.value=document.opciones.texto.value;
	parent.pantalla.fr_Tabla.document.medidas_det.action="medidas_det.asp?mode=search&lote=1&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value;
	parent.pantalla.fr_Tabla.document.medidas_det.submit();
	document.location="medidas_bt.asp";
}

function Cancelar() {
	parent.pantalla.fr_Tabla.document.medidas_det.action="medidas_det.asp?mode=browse?lote=1";
	parent.pantalla.fr_Tabla.document.medidas_det.submit();
}

if(window.document.addEventListener)
{
    window.document.addEventListener("keydown", callkeydownhandler, false);
}
else
{
    window.document.attachEvent("onkeydown", callkeydownhandler);
}

var ev = null;

function callkeydownhandler(evnt)
{
    ev = (evnt) ? evnt : event;
    comprobar_enter();
}

//****************************************************************************************************************
function comprobar_enter()
{
    var keycode = ev.keyCode;
	//si se ha pulsado la tecla enter
	if (keycode==13){
		document.opciones.criterio.focus();
		Buscar();
	}
}
</script>
<body class="body_master_ASP">
<form name="opciones" method="post">
<%if request("mode")="edit" then
    param=1
else
    param=2
end if%>
    <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_left_ASP" >
	        <table id="BUTTONS_CENTER_ASP">
		        <tr>
			        <td CLASS="CELDABOT" onclick="javascript:Cancelar();">
				        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			        </td>
                </tr>
            </table>
        </div>
        <div id="FILTERS_MASTER_ASP">
			<!--<td CLASS=CELDABOT><%=LitBuscar & ": "%>-->
				<SELECT class="IN_S" name="campos">
					<OPTION  value="codigo"><%=LitCodigo%></OPTION>
					<OPTION selected value="descripcion"><%=LitDescripcion%></OPTION>
				</SELECT>
			<!--</td>
			<td CLASS=CELDABOT>-->
				<SELECT class="IN_S" name="criterio">
					<OPTION value="contiene"><%=LitContiene%></OPTION>
					<!--<OPTION value="empieza"><%=LitComienza%></OPTION>-->
					<OPTION value="termina"><%=LitTermina%></OPTION>
					<OPTION value="igual"><%=LitIgual%></OPTION>
				</SELECT>
			<!--</td>
			<td class="CELDABOT">-->
				<input class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" onkeypress="javascript:comprobar_enter();">
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> ALT="<%=LitBuscar%>"></a>
			<!--</td>-->
		</div>
	</div>
<%ImprimirPie_bt%>
</form>
</BODY>
</HTML>