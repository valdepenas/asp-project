<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>

    <% dim  enc
    set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../styles/Master.css.inc" -->

<!--#include file="promociones.inc" -->
<title><%=LitTituloColor%></title>


<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
function Guardar(param) {
    ok = 1;

    mode = document.opciones.mode.value;
   if (mode == "edit") {
       param = 1;
   }
   else {
       param = 2;
   }

switch(param){
    case 1:
        if (parent.pantalla.document.promociones.e_qt_total.value < parent.pantalla.document.promociones.e_qt_dtos.value) {
            window.alert("<%=LITMSGCDTOSMAYORCTOTAL%>");
            ok = 0;
            break;
        }


        if (parent.pantalla.document.promociones.e_codigo.value == "") {
            window.alert("<%=LitMsgCodigoNoNulo%>");
            ok = 0;
            break;
        }
        if (comp_car_ext(parent.pantalla.document.promociones.e_codigo.value, 1) == 1) {
            window.alert("<%=LITMSGTARDESCARNOVAL%>");
            ok = 0;
            break;
        }
        if (parent.pantalla.document.promociones.e_descripcion.value == "") {
            window.alert("<%=LitMsgDescripcionNoNulo%>");
            ok = 0;
            break;
        }
        if (comp_car_ext(parent.pantalla.document.promociones.e_descripcion.value, 0) == 1) {
            window.alert("<%=LITMSGTARDESCARNOVAL%>");
            ok = 0;
            break;
        }
        if ((parent.pantalla.document.promociones.e_dto.value != "") && (parent.pantalla.document.promociones.e_dto2.value != "")) {
            window.alert("<%=LITSOLUNAOPC%>");
            ok = 0;
            break;
        }
        if (parent.pantalla.document.promociones.e_description_tpv.value == "") {
            window.alert("<%=LitMsgDescrNull%>");
            ok = 0;
            break;
        }
        if (parent.pantalla.document.promociones.e_v_from.value == "") {
            window.alert("<%=LITMSGDESDEFECHANONULO%>");
            ok = 0;
            break;
        }
        else {
            if (!checkdate(parent.pantalla.document.promociones.e_v_from)) {
                window.alert("<%=LITMSGDESDEFECHAFECHA%>");
                ok = 0;
                break;
            }
        }
        if (parent.pantalla.document.promociones.e_v_to.value == "") {
            window.alert("<%=LITMSGHASTAFECHANONULO%>");
            ok = 0;
            break;
        }
        else {
            if (!checkdate(parent.pantalla.document.promociones.e_v_to)) {
                window.alert("<%=LITMSGHASTAFECHAFECHA%>");
                ok = 0;
                break;
            }
        }
        if (parent.pantalla.document.promociones.e_qt_total.value == "" || parent.pantalla.document.promociones.e_qt_total.value == 0) {
            window.alert("<%=LITARTNECNONULO%>");
            ok = 0;
            break;
        }
        else {
            if (isNaN(parent.pantalla.document.promociones.e_qt_total.value.replace(",", "."))) {
                window.alert("<%=LITARTNECNUM%>");
                ok = 0;
                break;
            }
            else {
                if (parent.pantalla.document.promociones.e_qt_total.value.replace(",", ".") % 1 != 0) {
                    window.alert("<%=LITARTNECINT%>");
                    ok = 0;
                    break;
                }
            }
        }
        
        if (parent.pantalla.document.promociones.e_TypePromotion.value == 0) {
            if (parent.pantalla.document.promociones.e_qt_dtos.value == "") {
                window.alert("<%=LITCONDESNONULO%>");
                ok = 0;
                break;
            }
            if (parent.pantalla.document.promociones.impMinimo.value != 0) {
                if (isNaN(parent.pantalla.document.promociones.impMinimo.value.replace(",", "."))) {
                    window.alert("<%=LIT_MSG_LIMIT_CASH_NULO%>");
                    ok = 0;
                    break;
                }
            }
            if (parent.pantalla.document.promociones.e_dto.value != "") {
                if (isNaN(parent.pantalla.document.promociones.e_dto.value.replace(",", "."))) {
                    window.alert("<%=LitMsgImporteNumerico%>");
                    ok = 0;
                    break;
                }
            }
            if (parent.pantalla.document.promociones.e_dto2.value != "") {
                if (isNaN(parent.pantalla.document.promociones.e_dto2.value.replace(",", "."))) {
                    window.alert("<%=LitMsgDescuentoNumerico%>");
                    ok = 0;
                    break;
                }
                else {
                    if ((parent.pantalla.document.promociones.e_dto2.value > 100)) {
                        window.alert("<%=LITMSCDTOSVAL%>");
                        ok = 0;
                        break;
                    }
                }
            }
            if (!ValidarCamposE()) {
                    window.alert("<%=LITMSGPORDESCNONULO%>");
                    ok = 0;
                    break;
            }
	        if (!ValidarCondiciones()) {
	            window.alert("<%=LITERRORPROM%>");
			    ok = 0;
			    break;
		    }
        }else {
            if (parent.pantalla.document.promociones.referencia.value == "") {
                window.alert("<%=LitmsgGiftNulo%>");
                ok = 0;
                break;
            }
        }

        break;
    case 2:

        if (parent.pantalla.document.promociones.i_qt_total.value < parent.pantalla.document.promociones.i_qt_dtos.value) {
            window.alert("<%=LITMSGCDTOSMAYORCTOTAL%>");
            ok = 0;
            break;
        }
        if (parent.pantalla.document.promociones.i_codigo.value == "") {
            window.alert("<%=LitMsgCodigoNoNulo%>");
            ok = 0;
            break;
        }
        if (comp_car_ext(parent.pantalla.document.promociones.i_codigo.value, 3) == 1) {
            window.alert("<%=LITMSGTARDESCARNOVAL%>");
            ok = 0;
            break;
        }

        if (parent.pantalla.document.promociones.i_descripcion.value == "") {
            window.alert("<%=LitMsgDescripcionNoNulo%>");
            ok = 0;
            break;
        }
        if (comp_car_ext(parent.pantalla.document.promociones.i_descripcion.value, 0) == 1) {
            window.alert("<%=LITMSGTARDESCARNOVAL%>");
            ok = 0;
            break;
        }
        if ((parent.pantalla.document.promociones.i_dto.value != "") && (parent.pantalla.document.promociones.i_dto2.value != "")) {
            window.alert("<%=LITSOLUNAOPC%>");
            ok = 0;
            break;
        }
        if (parent.pantalla.document.promociones.i_descripcion_tpv.value == "") {
            window.alert("<%=LitMsgDescrNull%>");
            ok = 0;
            break;
        }
        if (parent.pantalla.document.promociones.i_v_from.value == "") {
            window.alert("<%=LITMSGDESDEFECHANONULO%>");
            ok = 0;
            break;
        }
        else {
            if (!checkdate(parent.pantalla.document.promociones.i_v_from)) {
                window.alert("<%=LITMSGDESDEFECHAFECHA%>");
                ok = 0;
                break;
            }
        }
        if (parent.pantalla.document.promociones.i_v_to.value == "") {
            window.alert("<%=LITMSGHASTAFECHANONULO%>");
            ok = 0;
            break;
        }
        else {
            if (!checkdate(parent.pantalla.document.promociones.i_v_to)) {
                window.alert("<%=LITMSGHASTAFECHAFECHA%>");
                ok = 0;
                break;
            }
        }
        if (parent.pantalla.document.promociones.i_qt_total.value == "" || parent.pantalla.document.promociones.i_qt_total.value == 0) {
            window.alert("<%=LITARTNECNONULO%>");
            ok = 0;
            break;
        }
        else {
            if (isNaN(parent.pantalla.document.promociones.i_qt_total.value.replace(",", "."))) {
                window.alert("<%=LITARTNECNUM%>");
                ok = 0;
                break;
            }
            else {
                if (parent.pantalla.document.promociones.i_qt_total.value.replace(",", ".") % 1 != 0) {
                    window.alert("<%=LITARTNECINT%>");
                    ok = 0;
                    break;
                }
            }
        }
        if (parent.pantalla.document.promociones.i_TypePromotion.value == 0) {
            if (parent.pantalla.document.promociones.i_qt_dtos.value == "") {
                window.alert("<%=LITCONDESNONULO%>");
                ok = 0;
                break;
            }
            if (parent.pantalla.document.promociones.i_dto.value != "") {
                if (isNaN(parent.pantalla.document.promociones.i_dto.value.replace(",", "."))) {
                    window.alert("<%=LitMsgImporteNumerico%>");
                    ok = 0;
                    break;
                }
            }
            if (parent.pantalla.document.promociones.i_dto2.value != "") {
                if (isNaN(parent.pantalla.document.promociones.i_dto2.value.replace(",", "."))) {
                    window.alert("<%=LitMsgDescuentoNumerico%>");
                    ok = 0;
                    break;
                }
                else {
                    if ((parent.pantalla.document.promociones.i_dto2.value > 100)) {
                        window.alert("<%=LITMSCDTOSVAL%>");
                        ok = 0;
                        break;

                    }
                }
            }
            if (!ValidarCamposI()) {
                window.alert("<%=LITMSGPORDESCNONULO%>");
                ok = 0;
                break;
            }
        } else {
            if (parent.pantalla.document.promociones.referencia.value == "") {
                window.alert("<%=LitmsgGiftNulo%>");
                ok = 0;
                break;
            }
        }

        break;
    }
	
    if (ok==1) {
		if (parent.pantalla.document.getElementById("frCondicionesAdd") !== null) {
			parent.pantalla.document.getElementById("frCondicionesAdd").contentDocument.CondicionesPromocion.action = "CondicionesPromocion.asp?mode=save&tarifa=" + parent.pantalla.document.promociones.htarifa.value;
			parent.pantalla.document.getElementById("frCondicionesAdd").contentDocument.CondicionesPromocion.submit();
		}
        parent.pantalla.document.promociones.submit();
        document.location = "promociones_bt.asp?mode=add";
    }
}

function ValidarCamposI(){
    var resultado  = false;
    if (((parent.pantalla.document.promociones.i_dto.value != "") || (parent.pantalla.document.promociones.i_dto2.value != "")))
        resultado = true;
    
    return resultado;
}

function ValidarCamposE(){
    var resultado = false;
    if (((parent.pantalla.document.promociones.e_dto.value != "") || (parent.pantalla.document.promociones.e_dto2.value != "")))
        resultado = true;

    return resultado;
}

function Buscar() {
	parent.pantalla.document.location="promociones.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next";
	document.location = "promociones_bt.asp?mode=add";
}

function ValidarCondiciones() {
	var elements = parent.pantalla.document.getElementById("frCondicionesAdd").contentDocument.CondicionesPromocion.elements;
	var pattern = /^\d+(\.\d|\.\d\d)?$/
	for (var i = 0, element; element = elements[i++];)
		if (element.type === "text" || element.type === "number")
			if (!pattern.test(element.value.replace(",", ".")))
				return false;
	return true;
}

function Eliminar(param) {

    mode = document.opciones.mode.value;
    if (mode == "edit") {
        param = 1;
    }
    else {
        param = 2;
    }

   if (window.confirm("<%=LitMsgEliminarPromocionConfirm%>")==true) {
      switch (param){
         case 1:
	        if (parent.pantalla.document.promociones.e_codigo.value=="") {
               window.alert ("<%=LitMsgCodigoNoNulo%>");
            }
            else {
               parent.pantalla.document.location="promociones.asp?mode=delete&codigo=" + parent.pantalla.document.promociones.e_codigo.value;
            }
    		break;

         case 2:
	        if (parent.pantalla.document.promociones.i_codigo.value=="") {
               window.alert ("<%=LitMsgCodigoNoNulo%>");
            }
            else {
               parent.pantalla.document.location="promociones.asp?mode=delete&codigo=" + parent.pantalla.document.promociones.i_codigo.value;
            }
			break;
      }
    document.location = "promociones_bt.asp?mode=add";
   }
}

function Cancelar() {
	parent.pantalla.document.location="promociones.asp?npagina="+parent.pantalla.document.promociones.h_npagina.value;
	document.location = "promociones_bt.asp?mode=add";
}

//****************************************************************************************************************
function comprobar_enter(){
	//si se ha pulsado la tecla enter
	//if (window.event.keyCode==13){
		document.opciones.criterio.focus();
		Buscar();
	//}
}
</script>
<body class="body_master_ASP">
<%
mode=enc.EncodeForJavascript(Request.QueryString("mode"))
%>
<form name="opciones" method="post">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
 <%if mode="edit" then
		param=1
	else
		param=2
	end if%>

<div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_left_ASP" >
        <table id="BUTTONS_CENTER_ASP">
		    <tr>
              <td id="idsave" class="CELDABOT" onclick="javascript:Guardar(<%=enc.EncodeForJavascript(param)%>);">
				        <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
			  </td>
                <%
                visible=""
                if mode<>"edit" then
                    visible=" style='display:none;' "
                end if%>
    			<td id="iddelete" <%=visible%> class="CELDABOT" onclick="javascript:Eliminar(<%=enc.EncodeForJavascript(param)%>);">
				    <%PintarBotonBTLeft LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
			    </td>
			    <td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar();">
				        <%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			        </td>
            </tr>
	    </table>
    </div>
    
        <div id="FILTERS_MASTER_ASP">
			
			<select class="IN_S" name="campos">
					<option value="code"><%=LitCodigo%></option>
					<option selected="selected" value="description"><%=LitDescripcion%></option>
				</select>
		
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<!--<option value="empieza"><%=LitComienza%></option>-->
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
				</select>
			
				<input id="KeySearch"  class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=themeIlion %><%=ImgBuscar_bt%>" <%=ParamImgBuscar_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			</div>
            </div>
    <table style="width:100%;height:42px;vertical-align:bottom;" align="center">
    <tr>
    <td style="width:100%;height:42px; vertical-align:bottom; text-align:center;">
    <%ImprimirPie_bt
    %>
    </td>
    </tr>
    </table>
</form>
</body>
</html>