<%@ Language=VBScript %>
<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html LANG="<%=session("lenguaje")%>">
<head>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  
    <title><%=LitTituloResCompra%></title>
    <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <META HTTP-EQUIV="Content-Type" Content="text/html; charset=<%=session("caracteres")%>">
</head>

<!--#include file="../../calculos.inc" -->
<!--#include file="../../constantes.inc" -->

<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->

<!--#include file="../facturas_pro.inc" -->

<!--#include file="../../styles/Master.css.inc" -->

<script type="text/javascript" language="JavaScript" src="../../jfunciones.js"></script>
<script type="text/javascript" language="JavaScript">
//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if (!checkdate(parent.pantalla.document.resumen_compras_pro.fdesde)) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.resumen_compras_pro.fhasta)) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}

	if (parent.pantalla.document.resumen_compras_pro.agrupar.value=="MESES") {
		fechamenor=parent.pantalla.document.resumen_compras_pro.fdesde.value;
		fechamayor=parent.pantalla.document.resumen_compras_pro.fhasta.value;
		que="dias"
		diasd=DiferenciaTiempo(fechamayor,fechamenor,que)
	}
	
	if (parent.pantalla.document.resumen_compras_pro.fdesde.value=="" && parent.pantalla.document.resumen_compras_pro.fhasta.value=="") {
		window.alert("<%=LitMsgFechasNulas%>");
		return false;
	}

	if (parent.pantalla.document.resumen_compras_pro.ver_conceptos.checked==true && (parent.pantalla.document.resumen_compras_pro.familia.value>"" || parent.pantalla.document.resumen_compras_pro.referencia.value>"" || parent.pantalla.document.resumen_compras_pro.nombreart.value>"")){
		window.alert("<%=LitSoloConceptos%>");
		return false;
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
    switch (mode) {
        case "pdf":
            switch (pulsado) {
                case "back": //Volver a la pantalla anterior
                    parent.document.location="../../central.asp?pag1=compras/listados/resumen_compras_pro.asp&mode=select1&pag2=compras/listados/resumen_compras_pro_bt.asp";
	                break;
            }
            break;

        case "browse":
            switch (pulsado) {
                case "imprimir": //Imprimir Listado
                    parent.pantalla.focus();
                    parent.pantalla.print();
                    break;

                case "cancelar": //Cancelar operacion
                    parent.pantalla.document.location = "resumen_compras_pro.asp?mode=select1";
                    document.location = "resumen_compras_pro_bt.asp?mode=select1";
                    break;

                case "imprimirp": //Imprimir Listado en PDF
                    if (parseInt(parent.pantalla.document.resumen_compras_proResultado.NumRegs.value) >= parseInt(parent.pantalla.document.resumen_compras_proResultado.maxpdf.value)) {
                        alert("<%=LitMsgRegPDF%>");
                    } else {
                        parent.pantalla.document.resumen_compras_proResultado.action = "resumen_compras_pro_pdf.asp?mode=browse&elTotal=" + parent.pantalla.document.resumen_compras_proResultado.elTotal.value +
						                           "&fdesde=" + parent.pantalla.document.resumen_compras_proResultado.fdesde.value +
										           "&fhasta=" + parent.pantalla.document.resumen_compras_proResultado.fhasta.value +
										           "&nserie=" + parent.pantalla.document.resumen_compras_proResultado.nserie.value +
										           "&nproveedor=" + parent.pantalla.document.resumen_compras_proResultado.nproveedor.value +
										           "&actividad=" + parent.pantalla.document.resumen_compras_proResultado.actividad.value +
										           "&tactividad=" + parent.pantalla.document.resumen_compras_proResultado.tactividad.value +
										           "&referencia=" + parent.pantalla.document.resumen_compras_proResultado.referencia.value +
										           "&nombreart=" + parent.pantalla.document.resumen_compras_proResultado.nombreart.value +
										           "&familia=" + parent.pantalla.document.resumen_compras_proResultado.familia.value +
										           "&agrupar=" + parent.pantalla.document.resumen_compras_proResultado.agrupar.value +
										           "&conceptos=" + parent.pantalla.document.resumen_compras_proResultado.conceptos.value +
										           "&ordenar_compras=" + parent.pantalla.document.resumen_compras_proResultado.ordenar_compras.value +
										           "&ver_conceptos=" + parent.pantalla.document.resumen_compras_proResultado.ver_conceptos.value +
										           "&cod_proyecto=" + parent.pantalla.document.resumen_compras_proResultado.cod_proyecto.value +
										           "&opc_cod_proyecto=" + parent.pantalla.document.resumen_compras_proResultado.opc_cod_proyecto.value;
                        parent.pantalla.document.resumen_compras_proResultado.submit();
                        document.location = "resumen_compras_pro_bt.asp?mode=pdf";
                    }
                    break;

                case "exportar": //Exportar el fichero a CSV
                    cadena = "";
                    cadena = cadena + "&fdesde=" + parent.pantalla.document.resumen_compras_proResultado.fdesde.value +
					        "&fhasta=" + parent.pantalla.document.resumen_compras_proResultado.fhasta.value +
					        "&nserie=" + parent.pantalla.document.resumen_compras_proResultado.nserie.value +
					        "&nproveedor=" + parent.pantalla.document.resumen_compras_proResultado.nproveedor.value +
					        "&actividad=" + parent.pantalla.document.resumen_compras_proResultado.actividad.value +
					        "&tactividad=" + parent.pantalla.document.resumen_compras_proResultado.tactividad.value +
					        "&referencia=" + parent.pantalla.document.resumen_compras_proResultado.referencia.value +
					        "&nombreart=" + parent.pantalla.document.resumen_compras_proResultado.nombreart.value +
					        "&familia=" + parent.pantalla.document.resumen_compras_proResultado.familia.value +
					        "&agrupar=" + parent.pantalla.document.resumen_compras_proResultado.agrupar.value +
					        "&conceptos=" + parent.pantalla.document.resumen_compras_proResultado.conceptos.value +
                            "&ver_conceptos=" + parent.pantalla.document.resumen_compras_proResultado.ver_conceptos.value +
					        "&ordenar_compras=" + parent.pantalla.document.resumen_compras_proResultado.ordenar_compras.value +
                            "&opcproveedorbaja=" + parent.pantalla.document.resumen_compras_proResultado.opcproveedorbaja.value +
					        "&cod_proyecto=" + parent.pantalla.document.resumen_compras_proResultado.cod_proyecto.value +
					        "&opc_cod_proyecto=" + parent.pantalla.document.resumen_compras_proResultado.opc_cod_proyecto.value +
                            "&seriesapf=" + parent.pantalla.document.resumen_compras_proResultado.seriesapf.value +
                            "&prohojassep=" + parent.pantalla.document.resumen_compras_proResultado.prohojassep.value +
                            "&opc_cantidad=" + parent.pantalla.document.resumen_compras_proResultado.opc_cantidad.value +
                            "&opc_coste=" + parent.pantalla.document.resumen_compras_proResultado.opc_coste.value +
                            "&opc_comprasnetas=" + parent.pantalla.document.resumen_compras_proResultado.opc_comprasnetas.value +
                            "&apaisado=" + parent.pantalla.document.resumen_compras_proResultado.apaisado.value +
                            "&tipo_proveedor=" + parent.pantalla.document.resumen_compras_proResultado.tipo_proveedor.value +
                            "&tipo_articulo=" + parent.pantalla.document.resumen_compras_proResultado.tipo_articulo.value +
                            "&mostrarfilas=" + parent.pantalla.document.resumen_compras_proResultado.mostrarfilas.value;
                    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    parent.pantalla.frameExportar.document.location = "resumen_compras_pro_exportar.asp?mode=exportar" + cadena;
                    break;
            }
            break;
			
		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos()) {
						parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						parent.pantalla.document.resumen_compras_pro.action="resumen_compras_proResultado.asp?mode=browse&confirma=NO&save=true";
						parent.pantalla.document.resumen_compras_pro.submit();
						document.location="resumen_compras_pro_bt.asp?mode=browse";
					}
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%
mode=enc.EncodeForJavascript(Request.QueryString("mode"))%>                                                               
<form name="opciones" method="post">
<div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_ASP" >
	    <table id="BUTTONS_CENTER_ASP" >
		    <tr><%
			    if mode="browse" then
				    %>
			        <td CLASS="CELDABOT" onclick="javascript:Accion('browse','imprimir');">
				        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
			        </td>
			        <td CLASS="CELDABOT" onclick="javascript:Accion('browse','imprimirp');">
				        <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
			        </td>
            	    <td CLASS="CELDABOT" onclick="javascript:Accion('browse','exportar');">
		                <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
	                </td>
	                <td CLASS="CELDABOT" onclick="javascript:Accion('browse','cancelar');">
		                <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
	                </td>
				    <%	
			    elseif mode="select1" then
				    %>
			        <td CLASS="CELDABOT" onclick="javascript:Accion('add','save');">
				        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
			        </td>
				    <%		
			    elseif mode="pdf" then
				    %>
			        <td CLASS="CELDABOT" onclick="javascript:Accion('pdf','back');">
				        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			        </td>
				    <%
			    end if%>
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