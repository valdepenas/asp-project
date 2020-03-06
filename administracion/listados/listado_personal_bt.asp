<%@ Language=VBScript %>
<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<TITLE><%=LitTituloLP%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
</HEAD>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="../personal.inc" -->

<!--#include file="../../styles/Master.css.inc" -->

<script language="JavaScript">
//Validacion de campos
function ValidarCampos()
{
	return true;
}

//Realizar la acción correspondiente al botón pulsado. 
function Accion(mode,pulsado) {
	switch (mode) {
		case "param":
			switch (pulsado) {
				case "aceptar": //Aceptar				
				  if(ValidarCampos()){
					parent.pantalla.document.listado_personal.action="listado_personalResultado.asp?mode=browse";
					parent.pantalla.document.listado_personal.submit();
					document.location="listado_personal_bt.asp?mode=browse";
				  }
					
					break;
				case "cancelar": //Cancelar
					parent.pantalla.document.listado_personal.action="listado_personal.asp?mode=param";
					parent.pantalla.document.listado_personal.submit();
					document.location="listado_personal_bt.asp?mode=param";
					break;
			}
			break;
		case "browse":
			switch (pulsado) {
				case "volver": //Volver atrás
					parent.pantalla.document.location="listado_personal.asp?mode=param";
					document.location="listado_personal_bt.asp?mode=param";
					break;
			    case "imprimir": //Imprimir
				        parent.pantalla.print(parent.pantalla.mainFrame);          	    
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.listado_personalResultado.NumRegs.value)>=parseInt(parent.pantalla.document.listado_personalResultado.maxpdf.value))
						alert("<%=LitDemReg%>");
					else
					{
						parent.pantalla.document.listado_personalResultado.action="listado_personal_pdf.asp?mode=browse";
						parent.pantalla.document.listado_personalResultado.submit();
						document.location="listado_personal_bt.asp?mode=pdf";
					}
					break;		
			}
		break;
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
				    parent.document.location="../../central.asp?pag1=administracion/listados/listado_personal.asp&mode=param&pag2=administracion/listados/listado_personal_bt.asp";
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
		            <%if mode="param" then%>
				        <td CLASS="CELDABOT" onclick="javascript:Accion('param','aceptar');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
			            <td CLASS="CELDABOT" onclick="javascript:Accion('param','cancelar');">
				            <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			            </td>
			        <%elseif mode="browse" then%>
				        <td CLASS="CELDABOT" onclick="javascript:Accion('browse','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <td CLASS="CELDABOT" onclick="javascript:Accion('browse','imprimirp');">
					        <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				        </td>
			            <td CLASS="CELDABOT" onclick="javascript:Accion('browse','volver');">
				            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			            </td>
			        <%elseif mode="pdf" then%>
				        <td CLASS="CELDABOT" onclick="javascript:Accion('pdf','back');">
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
</BODY>
</HTML>