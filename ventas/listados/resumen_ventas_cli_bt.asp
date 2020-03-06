<%@ Language=VBScript %>
<%'------------------------------CODIGOS DE AÑADIDURAS/MODIFICACIONES ------------------------
'JCI-090103-01 : He vuelto a añadir lo del total, ya que al quitarlo
'                salía 0 como total general en el PDF para cualquier agrupacion
'	FECHA :09/01/03
' AUTOR :JCI
'----------------------------------------------------------------------------------------------
'JCI 03/04/2003 : Control de caché y objeto de impresión
'' IML : 27/11/03 : Control de Impresion (controlimpresion.inc)
%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
    <head>
        <title><%=LitTituloResVent%></title>
        <% dim  enc
        set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
        <meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
        <!--#include file="../../constantes.inc" -->
        <!--#include file="../../cache.inc" -->
        <!--#include file="../../calculos.inc" -->
        <!--#include file="../../ilion.inc" -->
        <!--#include file="../../mensajes.inc" -->
        <!--#include file="../../tablas.inc" -->
        <!--#include file="../../varios_bt.inc" -->
        <!--#include file="../../ico.inc" -->
        <!--#include file="../../styles/Master.css.inc" -->
        <!--#include file="../facturas_cli.inc" -->
    </head>


    <script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
    <script language="javascript" type="text/javascript">
        //Validación de campos numéricos y fechas.
        function ValidarCampos() {
	        if (parent.pantalla.document.resumen_ventas_cli.agrupar.value=="MESES") {
		        fechamenor=parent.pantalla.document.resumen_ventas_cli.fdesde.value;
		        fechamayor=parent.pantalla.document.resumen_ventas_cli.fhasta.value;
		        que="dias"
		        diasd=DiferenciaTiempo(fechamayor,fechamenor,que)
	        }

	        if (parent.pantalla.document.resumen_ventas_cli.ncliente.value!="" && parent.pantalla.document.resumen_ventas_cli.acliente.value==""
	        || parent.pantalla.document.resumen_ventas_cli.ncliente.value=="" && parent.pantalla.document.resumen_ventas_cli.acliente.value!="") {
		        alert("<%=LitDebeHaberDesdeYHasta%>");
		        return false;
	        }

	        if (parent.pantalla.document.resumen_ventas_cli.nombre.value!="" && parent.pantalla.document.resumen_ventas_cli.anombre.value!="") {
		        if (!isNaN(parent.pantalla.document.resumen_ventas_cli.ncliente.value) && !isNaN(parent.pantalla.document.resumen_ventas_cli.acliente.value)) {
			        if (parseFloat(parent.pantalla.document.resumen_ventas_cli.ncliente.value)>parseFloat(parent.pantalla.document.resumen_ventas_cli.acliente.value)) {
				        alert("<%=LitClienteHastaMayorClienteDesde%>");
				        return false;
			        }
		        }
	        }

	        if (parent.pantalla.document.resumen_ventas_cli.ncliente.value!="" && parent.pantalla.document.resumen_ventas_cli.nombre.value=="") {
		        alert("<%=LitMsgClienteNoExiste%>");
		        parent.pantalla.document.resumen_ventas_cli.ncliente.focus();
		        parent.pantalla.document.resumen_ventas_cli.ncliente.select();
		        return false;
	        }

	        if (parent.pantalla.document.resumen_ventas_cli.acliente.value!="" && parent.pantalla.document.resumen_ventas_cli.anombre.value=="") {
		        alert("<%=LitMsgClienteNoExiste%>");
		        parent.pantalla.document.resumen_ventas_cli.acliente.focus();
		        parent.pantalla.document.resumen_ventas_cli.acliente.select();
		        return false;
	        }

	        if (parent.pantalla.document.resumen_ventas_cli.nproveedor.value!="" && parent.pantalla.document.resumen_ventas_cli.razon_social.value=="") {
		        alert("<%=LitMsgProveedorNoExiste%>");
		        parent.pantalla.document.resumen_ventas_cli.nproveedor.focus();
		        parent.pantalla.document.resumen_ventas_cli.nproveedor.select();
		        return false;
	        }

	        if (!checkdate(parent.pantalla.document.resumen_ventas_cli.fdesde)) {
		        alert("<%=LitMsgDesdeFechaFecha%>");
		        return false;
	        }
	        if (!checkdate(parent.pantalla.document.resumen_ventas_cli.fhasta)) {
		        alert("<%=LitMsgHastaFechaFecha%>");
		        return false;
	        }

            if (parent.pantalla.document.resumen_ventas_cli.hfrom.value != "") {
                if (!checkhora(parent.pantalla.document.resumen_ventas_cli.hfrom)) {
                    alert("<%=LITMSGERRORHOURFROM%>");
                    return false;
                }
            }
            if (parent.pantalla.document.resumen_ventas_cli.hto.value != "") {
                if (!checkhora(parent.pantalla.document.resumen_ventas_cli.hto)) {
                    alert("<%=LITMSGERRORHOURTO%>");
                    return false;
                }
            }

	        if (parent.pantalla.document.resumen_ventas_cli.fdesde.value=="" && parent.pantalla.document.resumen_ventas_cli.fhasta.value=="") {
		        alert("<%=LitMsgFechasNulas%>");
		        return false;
            }
            if (parent.pantalla.document.resumen_ventas_cli.hfrom.value != "") {
                var array_dateFrom = parent.pantalla.document.resumen_ventas_cli.fdesde.value.split('/');
                var array_hourFrom = parent.pantalla.document.resumen_ventas_cli.hfrom.value.split(':');
                var array_dateTo = parent.pantalla.document.resumen_ventas_cli.fhasta.value.split('/');
                var array_hourTo;
                if (parent.pantalla.document.resumen_ventas_cli.hto.value != "") {
                    array_hourTo = parent.pantalla.document.resumen_ventas_cli.hto.value.split(':');
                }
                var dateFrom = new Date(array_dateFrom[2], array_dateFrom[1] - 1, array_dateFrom[0], array_hourFrom[0], array_hourFrom[1]);
                var dateTo;
                if (parent.pantalla.document.resumen_ventas_cli.hto.value != "") {
                    dateTo = new Date(array_dateTo[2], array_dateTo[1] - 1, array_dateTo[0], array_hourTo[0], array_hourTo[1]);
                }
                else {
                    dateTo = new Date(array_dateTo[2], array_dateTo[1] - 1, array_dateTo[0], 23, 59, 59, 999);
                }
                if (dateFrom > dateTo) {
                    alert("<%=LITMSGERRORDATEHOUR%>");
                    return false;
                }
            }

	        return true;
        }

        //Realizar la acción correspondiente al botón pulsado.
        function Accion(mode,pulsado) {
	        switch (mode) {
		        case "pdf":
			        switch (pulsado) {
			            case "back": //Volver a la pantalla anterior			       
					        verdc=document.opciones.verdc.value;
					        parent.document.location="../../central.asp?pag1=ventas/listados/resumen_ventas_cli.asp&pag2=ventas/listados/resumen_ventas_cli_bt.asp&mode=select1&ndoc=" + verdc;
					        break;
			        }
			        break;
		        case "browse":
			        switch (pulsado) {
				        case "imprimir": //Imprimir Listado
					        parent.pantalla.focus();
					        parent.pantalla.print(parent.pantalla.document.resumen_ventas_cliResultado.apaisado.value,parent.pantalla.document.resumen_ventas_cliResultado.NumRegs.value);
					        break;
			            case "cancelar": //Cancelar operacion			       
					        verdc=parent.pantalla.document.resumen_ventas_cliResultado.verdc.value;
					        parent.pantalla.document.location="resumen_ventas_cli.asp?mode=select1&verdc=" + verdc;
					        document.location="resumen_ventas_cli_bt.asp?mode=select1";
					        break;
				        case "imprimirp": //Imprimir Listado en PDF
        			        if (parseInt(parent.pantalla.document.resumen_ventas_cliResultado.NumRegs.value)>=parseInt(parent.pantalla.document.resumen_ventas_cliResultado.maxpdf.value))
						        alert("<%=LitMsgRegPDF%>");
                            else {
						        verdc=parent.pantalla.document.resumen_ventas_cliResultado.verdc.value;
						        parent.pantalla.document.resumen_ventas_cliResultado.action="resumen_ventas_cli_pdf.asp?mode=browse" +
								           <%'***** COD : JCI-090103-01 *****%>
								           "&elTotal="        + parent.pantalla.document.resumen_ventas_cliResultado.elTotal.value +
								           <%'***** FIN COD : JCI-090103-01 *****%>
						                   "&fdesde="         + parent.pantalla.document.resumen_ventas_cliResultado.fdesde.value +
										           "&fhasta="         + parent.pantalla.document.resumen_ventas_cliResultado.fhasta.value +
										           "&nserie="         + parent.pantalla.document.resumen_ventas_cliResultado.nserie.value +
										           "&ncliente="       + parent.pantalla.document.resumen_ventas_cliResultado.ncliente.value +
										           "&actividad="      + parent.pantalla.document.resumen_ventas_cliResultado.actividad.value +
										           "&tactividad="     + parent.pantalla.document.resumen_ventas_cliResultado.tactividad.value +
										           "&referencia="     + parent.pantalla.document.resumen_ventas_cliResultado.referencia.value +
										           "&nombreart="      + parent.pantalla.document.resumen_ventas_cliResultado.nombreart.value +
										           "&familia="        + parent.pantalla.document.resumen_ventas_cliResultado.familia.value +
										           "&agrupar="        + parent.pantalla.document.resumen_ventas_cliResultado.agrupar.value +
										           "&conceptos="      + parent.pantalla.document.resumen_ventas_cliResultado.conceptos.value +
										           "&ordenar_ventas=" + parent.pantalla.document.resumen_ventas_cliResultado.ordenar_ventas.value +
										           "&ver_conceptos="  + parent.pantalla.document.resumen_ventas_cliResultado.ver_conceptos.value +
										           "&cod_proyecto="   + parent.pantalla.document.resumen_ventas_cliResultado.cod_proyecto.value +
										           "&apaisado="       + parent.pantalla.document.resumen_ventas_cliResultado.apaisado.value +
										           "&opc_cod_proyecto="  + parent.pantalla.document.resumen_ventas_cliResultado.opc_cod_proyecto.value +
                                                    "&provincia=" + parent.pantalla.document.resumen_ventas_cliResultado.provincia.value
                                                    //INICIO IMPORTE MEDIO IVA
                                                    "&importeMedioIva=" + parent.pantalla.document.resumen_ventas_cliResultado.importeMedioIva.value;
                                                    //FIN IMPORTE MEDIO IVA
                                                    "&hfrom=" + parent.pantalla.document.resumen_ventas_cliResultado.hfrom.value;
                                                    "&hto=" + parent.pantalla.document.resumen_ventas_cliResultado.hto.value;
						        parent.pantalla.document.resumen_ventas_cliResultado.submit();
						        document.location="resumen_ventas_cli_bt.asp?mode=pdf&verdc=" + verdc;
					        }
					        break;
                        case "exportar": //Exportar el fichero a CSV
                            cadena="";
                            cadena=cadena + "&fdesde=" + parent.pantalla.document.resumen_ventas_cliResultado.fdesde.value;
                            cadena=cadena + "&fhasta=" + parent.pantalla.document.resumen_ventas_cliResultado.fhasta.value;
                            cadena=cadena + "&mostrarfilas=" + parent.pantalla.document.resumen_ventas_cliResultado.mostrarfilas.value;
                            cadena=cadena + "&opc_cantidad=" + parent.pantalla.document.resumen_ventas_cliResultado.opc_cantidad.value;
                            cadena=cadena + "&opc_ventasnetas=" + parent.pantalla.document.resumen_ventas_cliResultado.opc_ventasnetas.value;
                            cadena=cadena + "&ordenar_ventas=" + parent.pantalla.document.resumen_ventas_cliResultado.ordenar_ventas.value;

                            parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
                            parent.pantalla.frameExportar.document.location="resumen_ventas_cli_exportar.asp?mode=exportar" + cadena;
                        break;
			        }
			        break;
		        case "add":
			        switch (pulsado) {
				        case "save": //Guardar registro
				            if (ValidarCampos()) {				        
						        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						        parent.pantalla.document.resumen_ventas_cli.action="resumen_ventas_cliResultado.asp?mode=browse&confirma=NO&save=true";
						        parent.pantalla.document.resumen_ventas_cli.submit();
						        document.location="resumen_ventas_cli_bt.asp?mode=browse&agrupar="+parent.pantalla.document.resumen_ventas_cli.agrupar.value;
					        }
					        break;
			        }
			        break;
	        }     
        }
    </script>

    <body class="body_master_ASP">
        <% 
        mode=Request.QueryString("mode")
        agrupar=limpiaCadena(Request.QueryString("agrupar"))

	        if request.querystring("verdc")>"" then
		        verdc=limpiaCadena(request.querystring("verdc"))
	        else
		        verdc=limpiaCadena(request.form("verdc"))
	        end if
         %>
        <form name="opciones" method="post">
            <input type="hidden" name="verdc" value="<%=enc.EncodeForHtmlAttribute(verdc)%>" />
            <div id="PageFooter_ASP" >
                <div id="ControlPanelFooter_ASP" >
	                <table id="BUTTONS_CENTER_ASP">
		                <tr>
		                    <%if mode="browse" then%>
			                    <td class="CELDABOT" onclick="javascript:Accion('browse','imprimir');">
				                    <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRPAGTITLE%>
			                    </td>
			                    <td class="CELDABOT" onclick="javascript:Accion('browse','imprimirp');">
				                    <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
			                    </td>
	                            <td class="CELDABOT" onclick="javascript:Accion('browse','cancelar');">
		                            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTCANCELARTITLE%>
	                            </td>
			                    <%if agrupar="MESES" then%>
	                                <td class="CELDABOT" onclick="javascript:Accion('browse','exportar');">
		                                <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
	                                </td>
			                    <%end if
			                elseif mode="select1" then%>
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
            <table style="width:100%;height:30px;vertical-align:bottom;" align="center">
                <tr>
                    <td style="width:100%;height:30px; vertical-align:bottom; text-align:center;">
                        <%ImprimirPie_bt%>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>