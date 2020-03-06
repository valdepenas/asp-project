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
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="stock_almacenes.inc" -->

<!--#include file="../../styles/Master.css.inc" -->

<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">

//Validacion de campos
function ValidarCampos() {
    if (parseInt(parent.pantalla.mainFrame.document.stock_almacenes.nRegs.value)>parseInt(parent.pantalla.mainFrame.document.stock_almacenes.maxpdf.value)) {
		window.alert("<%=LitMsgLimitePdf%>");
        return false;
    }
	return true;
}

function ValidarCampos2(){
	if (parent.pantalla.document.stock_almacenes.StockAFecha.value!="" && !checkdate(parent.pantalla.document.stock_almacenes.elements["StockAFecha"])){
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}

	if (isNaN(parent.pantalla.document.stock_almacenes.stockmayoroigual.value.replace(",","."))){
		window.alert("<%=LitMsgStockNumerico%>");
		return false
	}
	if (parent.pantalla.document.stock_almacenes.elements["sinventa"].checked==true && !checkdate(parent.pantalla.document.stock_almacenes.elements["sinventafdesde"])){
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		parent.pantalla.document.stock_almacenes.elements["sinventafdesde"].focus();
		return false
	}
	if (parent.pantalla.document.stock_almacenes.elements["sinventa"].checked==true && !checkdate(parent.pantalla.document.stock_almacenes.elements["sinventafhasta"])){
		window.alert("<%=LitMsgHastaFechaFecha%>");
		parent.pantalla.document.stock_almacenes.elements["sinventafhasta"].focus();
		return false
	}
	if (parent.pantalla.document.stock_almacenes.elements["sinventa"].checked==true && parent.pantalla.document.stock_almacenes.elements["sinventafdesde"].value==""){
		window.alert("<%=LitMsgDesdeFechaNoNulo%>");
		parent.pantalla.document.stock_almacenes.elements["sinventafdesde"].focus();
		return false
	}
	if (parent.pantalla.document.stock_almacenes.elements["sinventa"].checked==true && parent.pantalla.document.stock_almacenes.elements["sinventafhasta"].value==""){
		window.alert("<%=LitMsgHastaFechaNoNulo%>");
		parent.pantalla.document.stock_almacenes.elements["sinventafhasta"].focus();
		return false
	}
	if (parent.pantalla.document.stock_almacenes.almacen.value=="" && parent.pantalla.document.stock_almacenes.StockAFecha.value!=""){
		window.alert("<%=LitStockAFechaSinAlm%>");
		parent.pantalla.document.stock_almacenes.elements["almacen"].focus();
		return false
	}
	return true;
}


//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "add":
			switch (pulsado) {
				case "aceptar": //Aceptar
					if (ValidarCampos2()){
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
						parent.pantalla.document.stock_almacenes.action="stock_almacenes_datos.asp?mode=ver"
						parent.pantalla.document.stock_almacenes.submit();
						document.location="stock_almacenes_bt.asp?mode=ver";
					}
					break;
				case "cancelar": //Cancelar
					parent.pantalla.document.stock_almacenes.action="stock_almacenes.asp?mode=add";
					parent.pantalla.document.stock_almacenes.submit();
					document.location="stock_almacenes_bt.asp?mode=add";
					break;
			}
			break;
		case "ver":
			switch (pulsado) {
				case "volver": //Volver atrás
					parent.pantalla.document.location="stock_almacenes.asp?mode=add";
					document.location="stock_almacenes_bt.asp?mode=add";
					break;
			    case "imprimir": //Volver atrás
					parent.pantalla.mainFrame.focus();
					parent.pantalla.print();
					break;
			     case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.mainFrame.document.stock_almacenesResultado.totregs.value)>=parseInt(parent.pantalla.mainFrame.document.stock_almacenesResultado.maxpdf.value))
						alert("<%=LitMsgLimitePdf%>");
					else if (parseInt(parent.pantalla.mainFrame.document.stock_almacenesResultado.totregs.value)==0) {
						alert("<%=LitMsgNoResultado%>");
					}
					else {
						//ricardo 4-4-2006 con esta linea ya no se comprobara en la funcion comprobar()
						//si va todo bien en el fichero stock_almacenes_datos
						parent.pantalla.vcomp=1;
						parent.pantalla.mainFrame.document.stock_almacenesResultado.action = "stock_almacenes_pdf.asp?mode=browse&xls=0";
						parent.pantalla.mainFrame.document.stock_almacenesResultado.submit();
						document.location = "stock_almacenes_bt.asp?mode=pdf";
					}
					break;
	            case "exportar":
	            case "exportar2":
                    cadena="";
                    cadena=cadena + "&comocalccostcomp=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.comocalccostcomp.value;
                    cadena=cadena + "&ordenar=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ordenar.value;
                    cadena=cadena + "&ver_familia=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_familia.value;
                    cadena=cadena + "&ver_stock=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_stock.value;
                    cadena=cadena + "&ver_smin=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_smin.value;
                    cadena=cadena + "&ver_reposicion=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_reposicion.value;
                    cadena=cadena + "&ver_precibir=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_precibir.value;
                    cadena=cadena + "&ver_pservir=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_pservir.value;
                    cadena=cadena + "&ver_pmin=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_pmin.value;
                    cadena=cadena + "&ver_pvd=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_pvd.value;
                    cadena=cadena + "&ver_coste=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_coste.value;
                    cadena=cadena + "&ver_pvp=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_pvp.value;
                    cadena=cadena + "&ver_dto=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_dto.value;
                    cadena=cadena + "&ver_iva=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_iva.value;
                    cadena=cadena + "&ver_valor_mercado=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_valor_mercado.value;
                    cadena=cadena + "&ver_coste_medio=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_coste_medio.value;
                    cadena=cadena + "&ver_proveedores=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_proveedores.value;
                    cadena=cadena + "&ver_fecha_inventario=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_fecha_inventario.value;
                    cadena = cadena + "&ver_coste_articulo=" + parent.pantalla.mainFrame.document.stock_almacenesResultado.ver_coste_articulo.value;
                    parent.pantalla.topFrame.document.location = "stock_almacenes_pdf.asp?mode=browse&xls=1" + cadena;
					break;
			}
			break;
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
					parent.pantalla.document.location="stock_almacenes.asp?mode=add";
					document.location="stock_almacenes_bt.asp?mode=add";
					break;
			}
			break;

		case "edit":
			switch (pulsado) {
				case "cancel": //Volver atrás
					parent.pantalla.document.stock_almacenes.action="stock_almacenes.asp?mode=ver";
					parent.pantalla.document.stock_almacenes.submit();
					document.location="stock_almacenes_bt.asp?mode=ver";
					break;
				case "save": //Almacenar
	                if(ValidarCampos()){
					   parent.pantalla.document.stock_almacenes.action="stock_almacenes.asp?mode=save";
					   parent.pantalla.document.stock_almacenes.submit();
					   document.location="stock_almacenes_bt.asp?mode=ver";
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
	        <table id="BUTTONS_CENTER_ASP">
		        <tr>
		            <%if mode="add" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('add','aceptar');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('add','cancelar');">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
			        <%elseif mode="ver" then%>
				        <td class="CELDABOT"  style="display:none" onclick="javascript:Accion('ver','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <td class="CELDABOT"  style="display:none" onclick="javascript:Accion('ver','imprimirp');">
					        <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				        </td>
			            <td class="CELDABOT" onclick="javascript:Accion('ver','volver');">
				            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			            </td>
			            <td class="CELDABOT"  style="display:none" onclick="javascript:Accion('ver','exportar');">
			                <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
		                </td>
			        <%elseif mode="pdf" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				        </td>
			        <%elseif mode="edit" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('edit','save');">
					        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
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