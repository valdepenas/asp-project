<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>


<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="perso_camposFS.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../styles/Master.css.inc" -->

<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
function Buscar()
{
	parent.pantalla.fr_Tabla.document.perso_campos_det.action="perso_campos_det.asp?mode=search&lote=1&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value+"&sop="+parent.pantalla.fr_Tabla.document.perso_campos_det.sop.value;
	parent.pantalla.fr_Tabla.document.perso_campos_det.submit();
	document.location="perso_campos_bt.asp";
}

function Cancelar()
{
	parent.pantalla.fr_Tabla.document.perso_campos_det.action="perso_campos_det.asp?mode=browse?lote=1&sop="+parent.pantalla.fr_Tabla.document.perso_campos_det.sop.value;
	parent.pantalla.fr_Tabla.document.perso_campos_det.submit();
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
	if (keycode==13)
	{
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
				<SELECT class='IN_S' name="campos">
					<OPTION  value="substring(ncampo,6,10)"><%=LitNCampo%></OPTION>
					<!--<OPTION selected value="tabla"><%=LitCamptabla%></OPTION>-->
					<OPTION selected value="titulo"><%=LitCampTitulo%></OPTION>
				</SELECT>
			<!--</td>
			<td CLASS=CELDABOT>-->
				<SELECT class='IN_S' name="criterio">
					<OPTION value="contiene"><%=LitContiene%></OPTION>
					<!--<OPTION value="empieza"><%=LitComienza%></OPTION>-->
					<OPTION value="termina"><%=LitTermina%></OPTION>
					<OPTION value="igual"><%=LitIgual%></OPTION>
				</SELECT>
			<!--</td>
			<td CLASS=CELDABOT>-->
				<INPUT class='IN_S' type="text" name="texto" size=20 maxLength=20 value="" onKeyPress="javascript:comprobar_enter();">
				<A CLASS=CELDAREF href="javascript:Buscar();"><IMG src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> ALT="<%=LitBuscar%>"></A>
			<!--</td>-->
        </div>
    </div>
<span id="sisges" style="display:none">
<%ImprimirPiePopUp_bt%>
</span>
<span id="nosisges" style="display:none">
<%ImprimirPie_bt%>
</span>
<script language="javascript">
var page_loaded = 0;

function continuar()
{
	var vengo_de_sist_gestion;
	vengo_de_sist_gestion=parent.pantalla.document.perso_campos.vengo_de_sist_gestion.value;

	if (vengo_de_sist_gestion=="1")
	{
		document.all("sisges").style.display="";
		document.all("nosisges").style.display="none";
	}
	else
	{
		document.all("sisges").style.display="none";
		document.all("nosisges").style.display="";
	}
}

function esperar()
{
	if(page_loaded==0) window.setTimeout("esperar();", 500);
	else continuar();
}

if(page_loaded==0) esperar();
</script>
</form>
</BODY>
</HTML>
