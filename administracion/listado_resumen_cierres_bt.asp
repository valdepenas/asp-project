<%@ Language=VBScript %><% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="cierres_caja.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
function ValidarCampos()
{
    if (!checkdate(parent.pantalla.document.listado_resumen_cierres.Dfecha) || parent.pantalla.document.listado_resumen_cierres.Dfecha.value=='')
    {
        alert("<%=LitMsgFdesde%>");
        return false;
    }
    if (!checkdate(parent.pantalla.document.listado_resumen_cierres.Hfecha) || parent.pantalla.document.listado_resumen_cierres.Hfecha.value=='')
    {
        alert("<%=LitMsgFhasta%>");
        return false;
    }

    if (isNaN(parent.pantalla.document.listado_resumen_cierres.DCierre.value))
    {
        alert("<%=LitMsgDCierre%>");
        return false;
    }
    if (isNaN(parent.pantalla.document.listado_resumen_cierres.HCierre.value))
    {
        alert("<%=LitMsgHCierre%>");
        return false;
    }
    if ( DiferenciaTiempo(parent.pantalla.document.listado_resumen_cierres.Hfecha.value,parent.pantalla.document.listado_resumen_cierres.Dfecha.value,"dias")>31)
    {
        alert("<%=LitMsgPeriodo%>");
        return false;
    }

    return true;
}

function Accion(mode,pulsado)
{
    switch(mode)
    {
        case "param":
            switch(pulsado){
                case "Aceptar":
                    if (ValidarCampos()){
                        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
                        parent.pantalla.document.listado_resumen_cierres.action="listado_resumen_cierresResultado.asp?mode=browse"
                        parent.pantalla.document.listado_resumen_cierres.submit();
                        document.location="listado_resumen_cierres_bt.asp?mode=browse"
                    }
                    break;
             }
             break;
        case "browse":
            switch(pulsado){
                case "Imprimir":
                    parent.pantalla.focus();
				    parent.pantalla.print();
				    break;
                case "Cancelar":
                    parent.pantalla.document.listado_resumen_cierresResultado.action="listado_resumen_cierres.asp?mode=param"
                    parent.pantalla.document.listado_resumen_cierresResultado.submit();
                    document.location="listado_resumen_cierres_bt.asp?mode=param"
                    break;
            }
            break;
    }
}
</script>

<body class="body_master_ASP">

<form name="opciones" method="post">
    <%mode=enc.EncodeForJavascript(request.QueryString("mode"))%>
	<div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >	
		    <table id="BUTTONS_CENTER_ASP" >
		        <tr>
                    <%if mode="param" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('<%=enc.EncodeForJavascript(mode)%>','Aceptar');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
                    <%elseif mode="browse" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('<%=enc.EncodeForJavascript(mode)%>','Imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('<%=enc.EncodeForJavascript(mode)%>','Cancelar');">
					        <%PintarBotonBTRed LITBOTVOLVER,ImgCancelar,ParamImgCancelar,LITBOTVOLVERTITLE%>
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