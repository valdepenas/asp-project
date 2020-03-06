<%@ Language=VBScript %>
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloLis%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="pagos.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if (!checkdate(parent.pantalla.document.listado_pagos.Dfecha)) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.listado_pagos.Hfecha)) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}
	return true;
}

function Buscar() {
		parent.pantalla.document.location="listado_pagos.asp?mode=imp&campo=" + document.opciones.campos.value +
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=" + "&nproveedor=" + parent.pantalla.document.listado_pagos.nproveedor.value + "&viene=" + parent.pantalla.document.listado_pagos.viene.value;
		document.location="listado_pagos_bt.asp?viene=" + parent.pantalla.document.listado_pagos.viene.value;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "select1":
			switch (pulsado) {
				case "imp": //Aceptar
					if (ValidarCampos()) {
						parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						parent.pantalla.document.listado_pagos.action="listado_pagosResultado.asp?mode=" + pulsado;
						parent.pantalla.document.listado_pagos.submit();
						document.location="listado_pagos_bt.asp?mode=" + pulsado;
					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.location="listado_pagos.asp?mode=select1";
					document.location="listado_pagos_bt.asp?mode=select1";
					break;
			}
			break;

case "imp":

    switch (pulsado) {

        case "cancel": //Volver atrás
            parent.pantalla.document.location = "listado_pagos.asp?mode=select1";
            document.location = "listado_pagos_bt.asp?mode=select1";
            break;

        case "imprimir": //Volver atrás
            parent.pantalla.focus();
            parent.pantalla.print(); 
            break;

        case "imprimirp": //Imprimir Listado en PDF
            if (parseInt(parent.pantalla.document.listado_pagosResultado.NumRegsTotal.value) >= parseInt(parent.pantalla.document.listado_pagosResultado.maxpdf.value))
                alert("<%=LitMsgRegistros%>");
            else {
                parent.pantalla.document.listado_pagosResultado.action = "listado_pagos_pdf.asp?mode=browse&xls=0?ncliente=" + parent.pantalla.document.listado_pagosResultado.h_ncliente + "?nserie=" + parent.pantalla.document.listado_pagosResultado.h_nserie + "?actividad=" + parent.pantalla.document.listado_pagosResultado.h_actividad;
                parent.pantalla.document.listado_pagosResultado.submit();
                document.location = "listado_pagos_bt.asp?mode=pdf";
            }
            break;
        case "exportar":
            parent.pantalla.document.listado_pagosResultado.action = "listado_pagos_pdf.asp?mode=browse&xls=1";
            parent.pantalla.document.listado_pagosResultado.submit();
            document.location = "listado_pagos_bt.asp?mode=pdf";
            break;
    }

    break;
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
				            //FLM:20090916:arreglo botón volver.
							//parent.pantalla.document.location="listado_pagos.asp?mode=select1";
							parent.pantalla.location.href="listado_pagos.asp?mode=select1";
							document.location="listado_pagos_bt.asp?mode=select1";
							break;
			}
				break;
	}
}
</script>
<%if request.querystring("viene")>"" then
	viene=limpiaCadena(request.querystring("viene"))
elseif request.form("viene")>"" then
	viene=request.form("viene")
end if
if viene="tienda" then
	cadena="topmargin='0'"
else
	cadena=""
end if%>
<body class="body_master_ASP">
<%mode=enc.EncodeForJavascript(Request.QueryString("mode"))%>

<form name="opciones" method="post" action="javascript:if ('<%=viene%>'=='tienda') {document.opciones.criterio.focus();Buscar()}">

    <%if viene="tienda" then%>
		<table width='100%' border='0' cellspacing="1" cellpadding="1">
			<tr>
				<td class=CELDABOT>
					<select class="IN_S" name="campos">
						<option selected value="nfactura_pro"><%=LitFactura%></option>
						<option value="fecha"><%=LitFecha%></option>
					</select>
				</td>
				<td class=CELDABOT>
					<select class="IN_S" name="criterio">
						<option value="contiene"><%=LitContiene%></option>
						<option value="termina"><%=LitTermina%></option>
						<option value="igual"><%=LitIgual%></option>
					</select>
				</td>
				<td class=CELDABOT>
					<input class="IN_S" type="text" name="texto" size=20 maxlength=20 value="">
				</td>
				<td class=CELDABOT>
				   <a class='CELDAREF' href="javascript:Buscar();"><img src="../images/<%=ImgBuscar_bt%>" <%=ParamImgBuscar_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"></a>
				</td>
			</tr>
		</table>
	<%else%>
    <div id="PageFooter_ASP" >
<div id="ControlPanelFooter_ASP" >	
		<table id="BUTTONS_CENTER_ASP" >
			<tr>
			    <%if mode="select1" then%>
				    <td class="CELDABOT" onclick="javascript:Accion('select1','imp');">
					    <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				    </td>
				    <td class="CELDABOT" onclick="javascript:Accion('select1','select1');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				<%elseif mode="imp" then%>
				    <td class="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					    <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				    </td>
				    <td class="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					    <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				    </td>
                    <td class="CELDABOT" onclick="javascript:Accion('imp','exportar');">
				        <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			        </td>
				    <td class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
					    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
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
	<%end if%>
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