<%@ Language=VBScript %>
<%
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->

<TITLE><%=LitTituloDiv%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="JavaScript">
function Guardar(param) {
	ok=1;
	switch(param){
        case 1:
			if (parent.pantalla.document.divisas.e_codigo.value=="") {
        		window.alert ("<%=LitMsgCodigoNoNulo%>");
	            ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.e_descripcion.value=="")  {
      	      	window.alert ("<%=LitMsgDescripcionNoNulo%>");
	            ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.e_FactorCambio.value=="")  {
        		window.alert ("<%=LitMsgFactorCambioNoNulo%>");
	            ok=0;
      		}
			if (ok==1 && parent.pantalla.document.divisas.e_FechaCotizacion.value=="")  {
            	window.alert ("<%=LitMsgFechaCotizacionNoNulo%>");
	            ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.e_Abreviatura.value=="")  {
      	      	window.alert ("<%=LitMsgAbreviaturaNoNulo%>");
	      	    ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.e_Ndecimales.value=="")  {
				window.alert ("<%=LitMsgNumeroDecimalesNoNulo%>");
				ok=0;
			}
			else {
				if (isNaN(parent.pantalla.document.divisas.e_Ndecimales.value)) {
			   		window.alert("<%=LitMsgNumeroDecimalesNumerico%>");
			   		ok=0;
				}
			}
			if (ok==1 && isNaN(parent.pantalla.document.divisas.e_FactorCambio.value.replace(",","."))) {
				window.alert("<%=LitValFactCmbNum%>");
				ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.e_FactorCambio.value.replace(",",".")=="0") {
				window.alert("<%=LitValFactCmbNum%>");
				ok=0;
			}
			//'COMPROBAMOS SI LA DIVISA A EDITAR SE VA A PONER COMO MONEDA BASE Y SE PIDE CONFIRMACION'
			if(parent.pantalla.document.divisas.e_MonedaBase.checked==true){
				if (parent.pantalla.document.divisas.h_MonedaBase.value=="NO") {
               		if (window.confirm("<%=LitMsgModedaBaseConfirm%>")==true)
               		    parent.pantalla.document.divisas.action="divisas.asp?CambiarMoneda=SI";
				}
			}
			break;
        case 2:
			if (parent.pantalla.document.divisas.i_codigo.value=="") {
				window.alert ("<%=LitMsgCodigoNoNulo%>");
				ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.i_descripcion.value=="")  {
				window.alert ("<%=LitMsgDescripcionNoNulo%>");
				ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.i_FactorCambio.value=="")  {
				window.alert ("<%=LitMsgFactorCambioNoNulo%>");
				ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.i_FechaCotizacion.value=="")  {
				window.alert ("<%=LitMsgFechaCotizacionNoNulo%>");
				ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.i_Abreviatura.value=="")  {
				window.alert ("<%=LitMsgAbreviaturaNoNulo%>");
				ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.i_Ndecimales.value=="")  {
				window.alert ("<%=LitMsgNumeroDecimalesNoNulo%>");
				ok=0;
			}
			else {
				if (ok==1 && isNaN(parent.pantalla.document.divisas.i_Ndecimales.value)) {
					window.alert("<%=LitMsgNumeroDecimalesNumerico%>");
					ok=0;
				}
			}
			if (ok==1 && isNaN(parent.pantalla.document.divisas.i_FactorCambio.value.replace(",","."))) {
				window.alert("<%=LitValFactCmbNum%>");
				ok=0;
			}
			if (ok==1 && parent.pantalla.document.divisas.i_FactorCambio.value.replace(",",".")=="0") {
				window.alert("<%=LitValFactCmbNum%>");
				ok=0;
			}
			if(ok==1 && parent.pantalla.document.divisas.i_MonedaBase.checked==true){
           		if (window.confirm("<%=LitMsgModedaBaseConfirm%>")==true)
					parent.pantalla.document.divisas.action="divisas.asp?CambiarMoneda=SI";
			}
			break;
	}
	if (ok==1) {
		parent.pantalla.document.divisas.submit();
		document.location="divisas_bt.asp?mode=browse";
    }
}

function Buscar() {
	parent.pantalla.document.location="divisas.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next";
	document.location="divisas_bt.asp";
}

/*function Eliminar(param) {
   if (window.confirm("<%=LitMsgEliminarDivisaConfirm%>")==true) {

      switch (param){
         case 1:
	        if (parent.pantalla.document.divisas.e_codigo.value=="") {
               window.alert ("<%=LitMsgCodigoNoNulo%>");
            }
            else {
               parent.pantalla.document.location="divisas.asp?mode=delete&codigo=" + parent.pantalla.document.divisas.e_codigo.value;
            }
    		break;

         case 2:
	        if (parent.pantalla.document.divisas.i_codigo.value=="") {
               window.alert ("<%=LitMsgCodigoNoNulo%>");
            }
            else {
               parent.pantalla.document.location="divisas.asp?mode=delete&codigo=" + parent.pantalla.document.divisas.i_codigo.value;
            }
			break;
      }
      document.location="divisas_bt.asp";
   }

}*/

function Cancelar() {
	parent.pantalla.document.location="divisas.asp?npagina="+parent.pantalla.document.divisas.h_npagina.value;
	document.location="divisas_bt.asp";
}

//****************************************************************************************************************
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

function comprobar_enter(){
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
    		        <td CLASS="CELDABOT" onclick="javascript:Guardar(<%=param%>);">
				        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
			        </td>
			        <% if param=1 then %>
			            <td CLASS="CELDABOT" onclick="javascript:Cancelar();">
				            <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			            </td>
			        <% end if %>
                </tr>
            </table>
        </div>
        <div id="FILTERS_MASTER_ASP">
			<!--<td CLASS=CELDABOT><%=LitBuscar & ": "%>-->
				<SELECT class="IN_S" name="campos">
					<OPTION value="codigo"><%=LitCodigo%></OPTION>
					<OPTION value="descripcion" selected><%=LitDescripcion%></OPTION>
					<OPTION value="abreviatura"><%=LitAbreviatura%></OPTION>
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
			<td CLASS=CELDABOT>-->
				<INPUT class="IN_S" type="text" name="texto" size="20" maxLength="25" value="" onKeyPress="javascript:comprobar_enter();">
				<A CLASS=CELDAREF href="javascript:Buscar();"><IMG src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> ALT="<%=LitBuscar%>"></A>
			<!--</td>-->
		</div>
    </div>
<%ImprimirPie_bt%>
</form>
</BODY>
</HTML>