<%@ Language=VBScript %>
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

<!--#include file="../calculos.inc" -->

<!--#include file="cierres_caja.inc" -->

<!--#include file="../styles/Master.css.inc" -->
<!--#include file="../styles/FootButton.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
    function Buscar() {
        parent.pantalla.document.cierres_caja.action = "cierres_cajaResultado.asp?mode=search&campo=" + document.opciones.campos.value +
        "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value;
        parent.pantalla.document.cierres_caja.submit();
        document.location = "cierres_cajaResultado_bt.asp?mode=search";
    }

    //Validación de campos numéricos y fechas.
    function ValidarCampos() {
        if (parent.pantalla.document.cierres_caja.caja.value == "") {
            window.alert("<%=LitMsgCajaNoNulo%>");
            return false;
        }
        if (!checkdate(parent.pantalla.document.cierres_caja.Dfecha) || parent.pantalla.document.cierres_caja.Dfecha.value == "") {
            window.alert("<%=LitMsgDesdeFechaFecha%>");
            return false;
        }
        if (!checkdate(parent.pantalla.document.cierres_caja.Hfecha) || parent.pantalla.document.cierres_caja.Hfecha.value == "") {
            window.alert("<%=LitMsgHastaFechaFecha%>");
            return false;
        }
        return true;
    }

    //Realizar la acción correspondiente al botón pulsado.
    function Accion(mode, pulsado) {
        switch (mode) {
            case "add":
                switch (pulsado) {
                    case "save": //Guardar registro
                        if (parent.pantalla.document.cierres_caja.dni.value != "") {
                            if (ValidarCampos()) {
                                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                                parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?mode=first_save";
                                parent.pantalla.document.cierres_caja.submit();
                                document.location = "cierres_caja_bt.asp?mode=previo";
                            }
                        }
                        else alert("<%=LITMSGUSUARIOPERSONALNOEXISTE%>");
                        break;

                    case "cancel": //Cancelar edición
                        parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?mode=add";
                        parent.pantalla.document.cierres_caja.submit();
                        document.location = "cierres_caja_bt.asp?mode=add";
                        break;

                }
                break;

            case "browse":
                switch (pulsado) {
                    case "add": //Nuevo registro
                        if (parent.pantalla.document.cierres_caja.dni.value != "") {
                            parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?mode=" + pulsado;
                            parent.pantalla.document.cierres_caja.submit();
                            document.location = "cierres_caja_bt.asp?mode=" + pulsado;
                        }
                        else {
                            window.alert("<%=LITMSGUSUARIOPERSONALNOEXISTE%>");
                        }
                        break;

                    case "delete": //Eliminar registro
                        if (parent.pantalla.document.cierres_caja.ultimocierre.value == "1") {
                            if (window.confirm("<%=LitMsgEliminarCierreConfirm%>") == true) {
                                parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?mode=" + pulsado;
                                parent.pantalla.document.cierres_caja.submit();
                                document.location = "cierres_caja_bt.asp?mode=add";
                            }
                        }
                        else alert("<%=LitMsgBorrarUltimoCierre%>");
                        break;

                    case "print": //Imprimir
                        parent.pantalla.focus();
                        parent.pantalla.print();
                        break;
                }
                break;
            case "search":
                switch (pulsado) {
                    case "add": //Nuevo registro
                        if (parent.pantalla.document.cierres_cajaResultado.dni.value != "") {
                            parent.pantalla.document.cierres_cajaResultado.action = "cierres_caja.asp?mode=" + pulsado;
                            parent.pantalla.document.cierres_cajaResultado.submit();
                            document.location = "cierres_caja_bt.asp?mode=" + pulsado;
                        }
                        else {
                            window.alert("<%=LITMSGUSUARIOPERSONALNOEXISTE%>");
                        }
                        break;
                }
                break;

            case "previo":
                switch (pulsado) {
                    case "confirm":
                        if (confirm("<%=LitMsgConfirmCierre%>")) {
                            parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?mode=confirm";
                            parent.pantalla.document.cierres_caja.submit();
                            document.location = "cierres_caja_bt.asp?mode=browse";
                        }
                        break;

                    case "cancel":
                        parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?mode=add";
                        parent.pantalla.document.cierres_caja.submit();
                        document.location = "cierres_caja_bt.asp?mode=add";
                        break;

                    case "print":
                        if (confirm("<%=LitMsgRecuerdeOcultarMostrar%>")) {
                            parent.pantalla.focus();
                            try {
                                printWindow();
                            }
                            catch (e) {
                                parent.pantalla.print();
                            }
                        }
                        break;
                }
                break;

            case "edit":
                switch (pulsado) {
                    case "save": //Guardar registro
                        if (parent.pantalla.document.cierres_caja.dni.value != "") {
                            if (ValidarCampos()) {
                                parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?nhoja=" + parent.pantalla.document.cierres_caja.hnhoja.value +
                                "&mode=save";
                                parent.pantalla.document.cierres_caja.submit();
                                document.location = "cierres_caja_bt.asp?mode=browse";
                            }
                        }
                        else {
                            window.alert("<%=LITMSGUSUARIOPERSONALNOEXISTE%>");
                        }
                        break;

                    case "cancel": //Cancelar edición
                        parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?nhoja=" + parent.pantalla.document.cierres_caja.hnhoja.value +
                        "&mode=browse";
                        parent.pantalla.document.cierres_caja.submit();
                        document.location = "cierres_caja_bt.asp?mode=browse";
                        break;
                }
                break;
            case "print":
                switch (pulsado) {
                    case "cancel": //Cancelar edición
                        parent.pantalla.document.cierres_caja.action = "cierres_caja.asp?nhoja=" + parent.pantalla.document.cierres_caja.hnhoja.value +
                        "&mode=browse";
                        parent.pantalla.document.cierres_caja.submit();
                        document.location = "cierres_caja_bt.asp?mode=browse";
                        break;
                }
                break;
        }
    }

    //****************************************************************************************************************
    function comprobar_enter() {
        document.opciones.criterio.focus();
        Buscar();
    }
</script>

<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
<div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_left_ASP" >
        <table id="BUTTONS_CENTER_ASP">
		<tr>
			<%if mode>"" then
				if mode="add" then
                    'Validar si tiene configurado el módulo comercial VerticalEESS
                    const ModVerticalEESS = "V0"
                    si_tiene_modulo_vertical=ModuloContratado(session("ncliente"),ModVerticalEESS) 
                    if (si_tiene_modulo_vertical <> 0) then    
                    %>
                        <td id="idaccept" class="CELDABOT" style="display:none">
				        </td> <%
                    else %>
                        <td id="idaccept" class="CELDABOT" onclick="javascript:Accion('add','save');">
					        <%PintarBotonBTLeft  LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
                    <%end if%>			        
				<%elseif mode="previo" then%>
			        <td id="Td1" class="CELDABOT" onclick="javascript:Accion('previo','confirm');">
					    <%PintarBotonBTLeft  LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('previo','cancel');">
					    <%PintarBotonBTLeft  LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				    <td id="idprint" class="CELDABOT" onclick="javascript:Accion('previo','print');">
					    <%PintarBotonBTLeft  LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				    </td>
				<%elseif mode="browse" then%>
		            <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft  LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeft  LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				    <td id="Td2" class="CELDABOT" onclick="javascript:Accion('browse','print');">
					    <%PintarBotonBTLeft  LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				    </td>
				<%elseif mode="search" then%>
		            <td id="Td3" class="CELDABOT" onclick="javascript:Accion('search','add');">
					    <%PintarBotonBTLeft  LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				<%end if%>
            </tr>
        </table>
    </div>

    <div id="FILTERS_MASTER_ASP">
		<select class="IN_S" name="campos">
         	<option selected value="CA.descripcion"><%=LitCaja%></option>
			<option value="P.nombre"><%=LitOperador%></option>
			<option value="CI.fecha"><%=LitFecha%></option>
			<option value="CI.codigo"><%=LitCierre%></option>
        </select>
		<select class="IN_S" name="criterio">
			<option value="contiene"><%=LitContiene%></option>
			<option value="termina"><%=LitTermina%></option>
			<option value="igual"><%=LitIgual%></option>
		</select>
		<input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
		<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>"/></a>
			<%end if%>
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