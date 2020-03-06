<%@ Language=VBScript %><% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  
<script id=DebugDirectives runat=server language=javascript>
    // Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloExt%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="Ahoja_gastos.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
    //Validación de campos numéricos y fechas.
    function ValidarCampos() {
        var dHoraisValid = /^([0-1][0-9]|2[0-3]):([0-5][0-9])$/.test(parent.pantalla.document.listado_caja_param.Dhora.value);
        var hHoraisValid = /^([0-1][0-9]|2[0-3]):([0-5][0-9])$/.test(parent.pantalla.document.listado_caja_param.Hhora.value);
        if (!dHoraisValid) {
            window.alert("<%=LITMSGHDESDE%>")
            return false;
        }
        if (!hHoraisValid) {
            window.alert("<%=LITMSGHHASTA%>")
            return false;
        }
        if (!checkdate(parent.pantalla.document.listado_caja_param.Dfecha)) {
            window.alert("<%=LitMsgDesdeFechaFecha%>");
            return false;
        }
        if (!checkdate(parent.pantalla.document.listado_caja_param.Hfecha)) {
            window.alert("<%=LitMsgHastaFechaFecha%>");
            return false;
        }
        if (parent.pantalla.document.listado_caja_param.caja.value == "") {
            window.alert("<%=LitMsgCajaNoNulo%>");
            return false;
        }
        if (parent.pantalla.document.listado_caja_param.ndocumento.value > "" && parent.pantalla.document.listado_caja_param.tdocumento.value == "") {
            window.alert("<%=LitMsgNDocNoNuloYTDocNulo%>");
            return false;
        }
        return true;
    }

    //Realizar la acción correspondiente al botón pulsado.
    function Accion(mode, pulsado) {
        switch (mode) {
            case "browse":
                switch (pulsado) {
                    case "imprimir": //Imprimir Pagina Actual
                        parent.pantalla.focus();
                        parent.pantalla.print();
                        break;
                    case "imprimirp": //Imprimir Listado en PDF
                        if (parseInt(parent.pantalla.document.listado_caja_paramResultado.NumRegs.value) >= parseInt(parent.pantalla.document.listado_caja_paramResultado.maxpdf.value))
                            alert("<%=LitMsgDemReg%>");
                        else {
                            parent.pantalla.document.listado_caja_paramResultado.action = "listado_caja_pdf.asp?mode=browse&xls=0&apaisado=" + parent.pantalla.document.listado_caja_paramResultado.apaisado.value;
                            parent.pantalla.document.listado_caja_paramResultado.submit();
                            document.location = "listado_caja_param_bt.asp?mode=pdf";
                        }
                        break;
                    case "exportar": //Exportar a formato Excel
                        if (parseInt(parent.pantalla.document.listado_caja_paramResultado.NumRegs.value) >= parseInt(parent.pantalla.document.listado_caja_paramResultado.maxpdf.value))
                            alert("<%=LitMsgDemReg%>");
                        else {
                            cadena = "";
                            cadena = cadena + "&apaisado=" + parent.pantalla.document.listado_caja_paramResultado.apaisado.value;
                            cadena = cadena + "&Dfecha=" + parent.pantalla.document.listado_caja_paramResultado.Dfecha.value;
                            cadena = cadena + "&Hfecha=" + parent.pantalla.document.listado_caja_paramResultado.Hfecha.value;
                            cadena = cadena + "&tdocumento=" + parent.pantalla.document.listado_caja_paramResultado.tdocumento.value;
                            cadena = cadena + "&ndocumento=" + parent.pantalla.document.listado_caja_paramResultado.ndocumento.value;
                            cadena = cadena + "&tanotacion=" + parent.pantalla.document.listado_caja_paramResultado.tanotacion.value;
                            cadena = cadena + "&agrAnotacion=" + parent.pantalla.document.listado_caja_paramResultado.agrAnotacion.value;
                            cadena = cadena + "&agrTipoPago=" + parent.pantalla.document.listado_caja_paramResultado.agrTipoPago.value;
                            cadena = cadena + "&descripcion=" + parent.pantalla.document.listado_caja_paramResultado.descripcion.value;
                            cadena = cadena + "&tpago=" + parent.pantalla.document.listado_caja_paramResultado.tpago.value;
                            cadena = cadena + "&tapunte=" + parent.pantalla.document.listado_caja_paramResultado.tapunte.value;
                            cadena = cadena + "&mostrarSaldo=" + parent.pantalla.document.listado_caja_paramResultado.mostrarSaldo.value;
                            cadena = cadena + "&mostrarGasto=" + parent.pantalla.document.listado_caja_paramResultado.mostrarGasto.value;
                            cadena = cadena + "&caja=" + parent.pantalla.document.listado_caja_paramResultado.caja.value;
                            parent.pantalla.marcoExportar.document.location = "listado_caja_pdf.asp?mode=browse&xls=1" + cadena;
                        }
                        break;
                    case "cancelar": //Volver a la pantalla de parametros
                        parent.pantalla.document.location = "listado_caja_param.asp?mode=add";
                        document.location = "listado_caja_param_bt.asp?mode=add";
                        break;
                }
                break;

            case "pdf":
                switch (pulsado) {
                    case "back": //Volver a la pantalla anterior
                        parent.document.location = "../central.asp?pag1=administracion/listado_caja_param.asp&mode=add&pag2=administracion/listado_caja_param_bt.asp";
                        break;
                }
                break;

            case "add":
                switch (pulsado) {
                    case "save": //Guardar registro
                        if (ValidarCampos()) {
                            parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            parent.pantalla.document.getElementById("ocultartapunte").style.visibility = "hidden";
                            parent.pantalla.document.listado_caja_param.action = "listado_caja_paramResultado.asp?mode=browse&confirma=NO";
                            parent.pantalla.document.listado_caja_param.submit();
                            document.location = "listado_caja_param_bt.asp?mode=browse";
                        }
                        break;
                }
                break;
        }
    }
</script>
<body class="body_master_ASP">
<!-- MeadCo ScriptX -->
<%mode=enc.EncodeForJavascript(Request.QueryString("mode"))%>
<form name="opciones" method="post">
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_ASP" >
    <table id="BUTTONS_CENTER_ASP" >
		<tr>
		    <%if mode="browse" then%>
				<td class="CELDABOT" onclick="javascript:Accion('browse','imprimir');">
					<%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				</td>
				<td class="CELDABOT" onclick="javascript:Accion('browse','imprimirp');">
					<%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				</td>
                <td class="CELDABOT" onclick="javascript:Accion('browse','exportar');">
                    <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
                </td>
			    <td class="CELDABOT" onclick="javascript:Accion('browse','cancelar');">
				    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			    </td>
			<%elseif mode="add" then%>
				<td class="CELDABOT" onclick="javascript:Accion('add','save');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
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