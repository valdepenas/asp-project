<%@ Language=VBScript %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../calculos.inc" -->

<!--#include file="reg_inventario.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    function Accion(mode)
    {
        switch(mode)
        {
            case "imprimir": //Imprimir Listado
                parent.pantalla.focus();
                parent.pantalla.print();
                break;

            case "cancelar": //Cancelar operacion
                parent.pantalla.document.location= "reg_inventario.asp?mode=param";
                document.location= "reg_inventario_bt.asp?mode=param";
                break;
		
            case "imprimirp": //Imprimir Listado
                AbrirVentana('reg_inventario_imp.asp?mode=browsedif&ninventario=' + parent.pantalla.document.inventario.hninventario.value,'I',<%=AltoVentana%>,<%=AnchoVentana%>);
                break;
			
            case "regnoleidos": //Imprimir Listado
                if (confirm("<%=LitEstabStockCeroConfirm%>")) {
                    parent.pantalla.document.location= "reg_inventario.asp?mode=regnoleidos&ninventario=" + parent.pantalla.document.inventario.hninventario.value;
                    document.location= "reg_inventario_bt.asp?mode=impresion";
                }
                break;
        }
    }
</script>

<body class="body_master_ASP">
    <form name="opciones" method="post" action="">

<%
	if Request.QueryString("mode")="impresion" then%>

    <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
				    <td class="CELDABOT" onclick="javascript:Accion('imprimir');">
					    <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				    </td>
				    <td class="CELDABOT" onclick="javascript:Accion('imprimirp');">
					    <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				    </td>
                  <td class="CELDABOT" onclick="javascript:Accion('regnoleidos');">
				        <%PintarBotonBT LITBOTPONERACERO,ImgPonerTodoA0,ParamPonerTodoA0,LITBOTPONERACEROTITLE%>
			        </td>
			        <td class="CELDABOT" onclick="javascript:Accion('cancelar');">
				        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			        </td>
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