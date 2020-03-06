<%@ Language=VBScript %>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../mensajes.inc" -->

<!--#include file="categorias.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
function Buscar() {
        parent.pantalla.fr_Tabla.document.categorias_det.action = "categorias_det.asp?mode=search&lote=1&campo=" + document.opciones.campos.value +
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value;
        parent.pantalla.fr_Tabla.document.categorias_det.submit();
		document.location="categorias_bt.asp";
}

function Cancelar(mode) {
	if (mode=="browse") {
		parent.pantalla.document.location="categorias.asp";
		document.location="categorias_bt.asp";
	}
	else {
	    parent.pantalla.fr_Tabla.document.categorias_det.action = "categorias_det.asp?mode=browse?lote=1";
	    parent.pantalla.fr_Tabla.document.categorias_det.submit();
	}
}

//****************************************************************************************************************
/*
if (window.document.addEventListener) {
    window.document.addEventListener("keydown", callkeydownhandler, false);
} else {
    window.document.attachEvent("onkeydown", callkeydownhandler);
}
function callkeydownhandler(evnt) {
    ev = (evnt) ? evnt : event;
    comprobar_enter(ev);
}
*/
function comprobar_enter(e){
    //si se ha pulsado la tecla enter
    //var keycode = e.keyCode;
    //if (keycode == 13) {
		document.opciones.criterio.focus();
		Buscar();
	//}
}
</script>
<body class="body_master_ASP">

<%
mode=Request.QueryString("mode")
if request("mode")="edit" then
    param=1
else
    param=2
end if%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
			<%if request("mode")<>"browse" then %>
	            <td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar('');">
			        <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
		        </td>
		    </tr>
	    </table>
        </div>
    
        <div id="FILTERS_MASTER_ASP">
				<!--<td class=CELDABOT><%=LitBuscar & ": "%>-->
					<select class="IN_S" name="campos">
						<option  value="codigo"><%=LitCodigo%></option>
						<option selected value="nombre"><%=LitDescripcion%></option>
					</select>
				<!--</td>
				<td class=CELDABOT>-->
					<select class="IN_S" name="criterio">
						<option value="contiene"><%=LitContiene%></option>
						<option value="empieza"><%=LitComienza%></option>
						<option value="termina"><%=LitTermina%></option>
						<option value="igual"><%=LitIgual%></option>
					</select>
				<!--</td>
				<td class=CELDABOT>-->
                    <input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<!--</td>
				<td class=CELDABOT>-->
				   <a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
				<!--</td>-->
            </div>
            </div>
			<%else%>
	            <td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar('browse');">
			        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
		        </td>
		    </tr>
	    </table>
        </div>
			<%end if%>

        
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