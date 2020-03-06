<%@ Language=VBScript %>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
</HEAD>

<!--#include file="../../calculos.inc" -->
<!--#include file="../../constantes.inc" -->

<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->

<!--#include file="../proveedores.inc" -->

<!--#include file="../../styles/Master.css.inc" -->

<script language="JavaScript" src="../../jfunciones.js"></script>

<script language="JavaScript">
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1);

//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if (!checkdate(parent.pantalla.document.listado_proveedores.fdesde)) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.listado_proveedores.fhasta)) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.listado_proveedores.fbdesde)) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.listado_proveedores.fbhasta)) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}

	if (parent.pantalla.document.listado_proveedores.si_campo_personalizables.value==1){
		num_campos=parent.pantalla.document.listado_proveedores.num_campos.value;
		respuesta=comprobarCampPerso("parent.pantalla.",num_campos,"listado_proveedores");
		if(respuesta!=0){
			titulo="titulo_campo" + respuesta;
			tipo="tipo_campo" + respuesta;
			titulo=parent.pantalla.document.listado_proveedores.elements[titulo].value;
			tipo=parent.pantalla.document.listado_proveedores.elements[tipo].value;
			if (tipo==4) {
				nomTipo="<%=LitTipoNumericoPro%>";
			}
			else if (tipo==5) {
				nomTipo="<%=LitTipoFechaPro%>";
			}

			window.alert("<%=LitMsgCampoPro%> " + titulo + " <%=LitMsgTipoPro%> " + nomTipo);

			return false;
		}
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "select1":
			switch (pulsado) {
				case "imp": //Aceptar
					if (ValidarCampos()) {
						parent.pantalla.document.listado_proveedores.action="listado_proveedoresResultado.asp?mode=" + pulsado;
						parent.pantalla.document.listado_proveedores.submit();
						document.location="listado_proveedores_bt.asp?mode=" + pulsado;
					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.listado_proveedores.action="listado_proveedores.asp?mode=" + pulsado;
					parent.pantalla.document.listado_proveedores.submit();
					document.location="listado_proveedores_bt.asp?mode=" + pulsado;
					break;
			}
		    break;

        case "imp":
            switch (pulsado) {
                case "cancel": //Volver atrás
                    parent.pantalla.document.location = "listado_proveedores.asp?mode=select1";
                    document.location = "listado_proveedores_bt.asp?mode=select1";
                    break;
                case "imprimir": //Volver atrás
                    parent.pantalla.focus();
                    parent.pantalla.print();
                    break;
                case "imprimirp": //Imprimir Listado en PDF
                    if (parseInt(parent.pantalla.document.listado_proveedoresResultado.NumRegs.value) >= parseInt(parent.pantalla.document.listado_proveedoresResultado.maxpdf.value)) {
                        alert("<%=LitMsgRegPDF%>");
                    }
                    else {
                        parent.pantalla.document.listado_proveedoresResultado.action = "listado_proveedores_pdf.asp?mode=browse&apaisado=" + parent.pantalla.document.listado_proveedoresResultado.apaisado.value;
                        parent.pantalla.document.listado_proveedoresResultado.submit();
                        document.location = "listado_proveedores_bt.asp?mode=pdf";
                    }
                    break;
                case "exportar":
                    var cadena = ""
                    cadena = cadena + "&fdesde=" + parent.pantalla.document.listado_proveedoresResultado.fdesde.value;
                    cadena = cadena + "&fhasta=" + parent.pantalla.document.listado_proveedoresResultado.fhasta.value;
                    cadena = cadena + "&fbdesde=" + parent.pantalla.document.listado_proveedoresResultado.fbdesde.value;
                    cadena = cadena + "&fbhasta=" + parent.pantalla.document.listado_proveedoresResultado.fbhasta.value;
                    cadena = cadena + "&razon_social=" + parent.pantalla.document.listado_proveedoresResultado.razon_social.value;
                    cadena = cadena + "&poblacion=" + parent.pantalla.document.listado_proveedoresResultado.poblacion.value;
                    cadena = cadena + "&provincia=" + parent.pantalla.document.listado_proveedoresResultado.provincia.value;
                    cadena = cadena + "&tarifa=" + parent.pantalla.document.listado_proveedoresResultado.tarifa.value;
                    cadena = cadena + "&formapago=" + parent.pantalla.document.listado_proveedoresResultado.formapago.value;
                    cadena = cadena + "&actividad=" + parent.pantalla.document.listado_proveedoresResultado.actividad.value;
                    cadena = cadena + "&tipo_proveedor=" + parent.pantalla.document.listado_proveedoresResultado.fbhasta.value;
                    cadena = cadena + "&ordenar=" + parent.pantalla.document.listado_proveedoresResultado.fbhasta.value;

                    cadena = cadena + "&opcproveedorbaja=" + parent.pantalla.document.listado_proveedoresResultado.opcproveedorbaja.value;
                    cadena = cadena + "&solodistribuidores=" + parent.pantalla.document.listado_proveedoresResultado.solodistribuidores.value;
                    cadena = cadena + "&mostrarcontactos=" + parent.pantalla.document.listado_proveedoresResultado.mostrarcontactos.value;
                    cadena = cadena + "&opccif=" + parent.pantalla.document.listado_proveedoresResultado.opccif.value;
                    cadena = cadena + "&opccontacto=" + parent.pantalla.document.listado_proveedoresResultado.opccontacto.value;
                    cadena = cadena + "&opcdomicilio=" + parent.pantalla.document.listado_proveedoresResultado.opcdomicilio.value;
                    cadena = cadena + "&opccodigopostal=" + parent.pantalla.document.listado_proveedoresResultado.opccodigopostal.value;
                    cadena = cadena + "&opcpoblacion=" + parent.pantalla.document.listado_proveedoresResultado.opcpoblacion.value;
                    cadena = cadena + "&opcprovincia=" + parent.pantalla.document.listado_proveedoresResultado.opcprovincia.value;
                    cadena = cadena + "&opctelefono=" + parent.pantalla.document.listado_proveedoresResultado.opctelefono.value;
                    cadena = cadena + "&opcfalta=" + parent.pantalla.document.listado_proveedoresResultado.opcfalta.value;
                    cadena = cadena + "&opcfbaja=" + parent.pantalla.document.listado_proveedoresResultado.opcfbaja.value;
                    cadena = cadena + "&opctarifa=" + parent.pantalla.document.listado_proveedoresResultado.opctarifa.value;
                    cadena = cadena + "&opccuenta=" + parent.pantalla.document.listado_proveedoresResultado.opccuenta.value;
                    cadena = cadena + "&opcformapago=" + parent.pantalla.document.listado_proveedoresResultado.opcformapago.value;
                    cadena = cadena + "&opctipopago=" + parent.pantalla.document.listado_proveedoresResultado.opctipopago.value;
                    cadena = cadena + "&opcrfinanciero=" + parent.pantalla.document.listado_proveedoresResultado.opcrfinanciero.value;
                    cadena = cadena + "&opcIRPF=" + parent.pantalla.document.listado_proveedoresResultado.opcIRPF.value;
                    cadena = cadena + "&opclven1=" + parent.pantalla.document.listado_proveedoresResultado.opclven1.value;
                    cadena = cadena + "&opclven2=" + parent.pantalla.document.listado_proveedoresResultado.opclven2.value;
                    cadena = cadena + "&opcentidad=" + parent.pantalla.document.listado_proveedoresResultado.opcentidad.value;
                    cadena = cadena + "&opcnumcuenta=" + parent.pantalla.document.listado_proveedoresResultado.opcnumcuenta.value;
                    cadena = cadena + "&opcactividad=" + parent.pantalla.document.listado_proveedoresResultado.opcactividad.value;
                    cadena = cadena + "&opctproveedor=" + parent.pantalla.document.listado_proveedoresResultado.opctproveedor.value;
                    cadena = cadena + "&opcportes=" + parent.pantalla.document.listado_proveedoresResultado.opcportes.value;
                    cadena = cadena + "&opcnomcom=" + parent.pantalla.document.listado_proveedoresResultado.opcnomcom.value;
                    cadena = cadena + "&opcfax=" + parent.pantalla.document.listado_proveedoresResultado.opcfax.value;
                    cadena = cadena + "&opctelmov=" + parent.pantalla.document.listado_proveedoresResultado.opctelmov.value;
                    cadena = cadena + "&opcweb=" + parent.pantalla.document.listado_proveedoresResultado.opcweb.value;
                    cadena = cadena + "&opcobs=" + parent.pantalla.document.listado_proveedoresResultado.opcobs.value;
                    cadena = cadena + "&opcemail=" + parent.pantalla.document.listado_proveedoresResultado.opcemail.value;
                    cadena = cadena + "&apaisado=" + parent.pantalla.document.listado_proveedoresResultado.apaisado.value;
            
                    cadena = cadena + "&campo1=" + parent.pantalla.document.listado_proveedoresResultado.campo1.value;
                    cadena = cadena + "&campo2=" + parent.pantalla.document.listado_proveedoresResultado.campo2.value;
                    cadena = cadena + "&campo3=" + parent.pantalla.document.listado_proveedoresResultado.campo3.value;
                    cadena = cadena + "&campo4=" + parent.pantalla.document.listado_proveedoresResultado.campo4.value;
                    cadena = cadena + "&campo5=" + parent.pantalla.document.listado_proveedoresResultado.campo5.value;
                    cadena = cadena + "&campo6=" + parent.pantalla.document.listado_proveedoresResultado.campo6.value;
                    cadena = cadena + "&campo7=" + parent.pantalla.document.listado_proveedoresResultado.campo7.value;
                    cadena = cadena + "&campo8=" + parent.pantalla.document.listado_proveedoresResultado.campo8.value;
                    cadena = cadena + "&campo9=" + parent.pantalla.document.listado_proveedoresResultado.campo9.value;
                    cadena = cadena + "&campo10=" + parent.pantalla.document.listado_proveedoresResultado.campo10.value;

                    cadena = cadena + "&ver_campo1=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo1.value;
                    cadena = cadena + "&ver_campo2=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo2.value;
                    cadena = cadena + "&ver_campo3=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo3.value;
                    cadena = cadena + "&ver_campo4=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo4.value;
                    cadena = cadena + "&ver_campo5=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo5.value;
                    cadena = cadena + "&ver_campo6=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo6.value;
                    cadena = cadena + "&ver_campo7=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo7.value;
                    cadena = cadena + "&ver_campo8=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo8.value;
                    cadena = cadena + "&ver_campo9=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo9.value;
                    cadena = cadena + "&ver_campo10=" + parent.pantalla.document.listado_proveedoresResultado.ver_campo10.value;
            

                    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    parent.pantalla.marcoExportar.document.location = "listado_proveedores_exportar.asp?mode=exportar" + cadena;
                    break;
            }
            break;

		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
				    parent.document.location="../../central.asp?pag1=compras/listados/listado_proveedores.asp&mode=select1&pag2=compras/listados/listado_proveedores_bt.asp";
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP" class="BODY_ASP">
<%
mode=enc.EncodeForJavascript(Request.QueryString("mode"))%>
<form name="opciones" method="post">
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_ASP" >
	<table id="BUTTONS_CENTER_ASP" >
		<tr><%
			if mode="select1" then
				%>
				<td CLASS="CELDABOT" onclick="javascript:Accion('select1','imp');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				</td>
				<td CLASS="CELDABOT" onclick="javascript:Accion('select1','select1');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
				<%
			elseif mode="imp" then
				%>	
				<td CLASS="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					<%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				</td>
				<td CLASS="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					<%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				</td>
                <td class="CELDABOT" onclick="javascript:Accion('imp','exportar');">
				    <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			    </td>
			    <td CLASS="CELDABOT" onclick="javascript:Accion('imp','cancel');">
				    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
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
</BODY>
</HTML>
