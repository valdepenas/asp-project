<%@ Language=VBScript %>
<script id='DebugDirectives' runat='server' language='javascript'>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../styles/Master.css.inc" -->
<!--#include file="listado_escandallo.inc" -->
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>

</head>
<script language="javascript" type="text/javascript">
//Validacion de campos
function ValidarCampos() {
	if (parseInt(parent.pantalla.document.listado_escandalloResultado.nRegs.value)>parseInt(parent.pantalla.document.listado_escandalloResultado.maxpdf.value)) {
        window.alert("<%=LitMsgLimitePdf%>");
        return false;
    }
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode, pulsado) {
	switch (mode) {
		case "add":
			switch (pulsado) {
			    case "aceptar": //Aceptar
			        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
					parent.pantalla.document.listado_escandallo.action="listado_escandalloResultado.asp?mode=ver";
					parent.pantalla.document.listado_escandallo.submit();
					document.location="listado_escandallo_bt.asp?mode=ver";
					break;

				case "cancelar": //Cancelar
					parent.pantalla.document.listado_escandallo.action="listado_escandallo.asp?mode=add";
					parent.pantalla.document.listado_escandallo.submit();
					document.location="listado_escandallo_bt.asp?mode=add";
					break;
			}
			break;
case "ver":
    switch (pulsado) {
        case "volver": //Volver atrás
            parent.pantalla.document.location = "listado_escandallo.asp?mode=add";
            document.location = "listado_escandallo_bt.asp?mode=add";
            break;
        case "imprimir": //Volver atrás
            parent.pantalla.focus();
            //printWindow();
            parent.pantalla.print();
            break;
        case "imprimirp": //Imprimir Listado en PDF
            if (ValidarCampos()) {
                if (parseInt(parent.pantalla.document.listado_escandalloResultado.maxpagina.value) >= parseInt(parent.pantalla.document.listado_escandalloResultado.maxpdf.value))
                    alert("<%=LitMsgLimitePdf%>");
                else {
                    parent.pantalla.document.listado_escandalloResultado.action = "listado_escandallo_pdf.asp?mode=add";
                    parent.pantalla.document.listado_escandalloResultado.submit();
                    document.location = "listado_escandallo_bt.asp?mode=pdf";
                }
            }
            break;
    }
    break;
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
				    parent.document.location="../../central.asp?pag1=productos/listados/listado_escandallo.asp&pag2=productos/listados/listado_escandallo_bt.asp&mode=add";
					break;
			}
			break;

		case "edit":
			switch (pulsado) {
				case "cancel": //Volver atrás
					parent.pantalla.document.listado_escandalloResultado.action="listado_escandalloResultado.asp?mode=ver";
					parent.pantalla.document.listado_escandalloResultado.submit();
					document.location="listado_escandallo_bt.asp?mode=ver";
					break;
				case "save": //Almacenar
	                if(ValidarCampos()){
					   parent.pantalla.document.listado_escandalloResultado.action="listado_escandalloResultado.asp?mode=save";
					   parent.pantalla.document.listado_escandalloResultado.submit();
					   document.location="listado_escandallo_bt.asp?mode=ver";
					}
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%
mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
	<div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
	        <table id="BUTTONS_CENTER_ASP">
		        <tr>
		    <%if mode="add" then%>
				<td id="idaccept" class="CELDABOT" onclick="javascript:Accion('add','aceptar');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancelar');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				</td>
			<%elseif mode="ver" then%>
				<td id="idprintpag" class="CELDABOT" onclick="javascript:Accion('ver','imprimir');">
					<%PintarBotonBT LITBOTIMPRIMIRPAG,ImgImprimir_pag,ParamImgImprimir_pag,""%>
				</td>
				<td id="idprintlist" class="CELDABOT" onclick="javascript:Accion('ver','imprimirp');">
					<%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,""%>
				</td>
			    <td id="idreturn" class="CELDABOT" onclick="javascript:Accion('ver','volver');">
				    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,""%>
			    </td>
			<%elseif mode="pdf" then%>
				<td id="idreturn" class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					<%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,""%>
				</td>
			<%elseif mode="edit" then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					<%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
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