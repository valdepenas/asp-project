<%@ Language=VBScript %>
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  


<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="pedpro_albpro_param.inc" -->

<!--#include file="../styles/Master.css.inc" -->

</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
//Validación de campos numéricos y fechas.
function ValidarCampos()
{
	if (!checkdate(parent.pantalla.document.albpedpro_facpro_param.fdesde)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return false;
    }

	if (!checkdate(parent.pantalla.document.albpedpro_facpro_param.fhasta)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return false;
    }

	if (parent.pantalla.document.albpedpro_facpro_param.fdesde.value=="") {
		window.alert("<%=LitMsgDesdeFechaNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.albpedpro_facpro_param.fhasta.value=="") {
		window.alert("<%=LitMsgHastaFechaNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.albpedpro_facpro_param.nproveedor.value=="") {
		window.alert("<%=LitMsgProveedorNoNulo%>");
		return false;
	}
	return true;
}

function ValidarCampos2()
{
	if (parent.pantalla.document.albpedpro_facpro_param.nserie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.albpedpro_facpro_param.ffactura.value=="") {
		window.alert("<%=LitMsgFechaNoNulo%>");
		return false;
	}

	if (!checkdate(parent.pantalla.document.albpedpro_facpro_param.ffactura)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return false;
    }

	if (parent.pantalla.document.albpedpro_facpro_param.nfactura_pro.value=="") {
		window.alert("<%=LitMsgFacturaNoNulo%>");
		return false;
	}
	if (!window.confirm(parent.pantalla.document.albpedpro_facpro_param.mensaje.value)) return false;

	return true;
}

function to_select2()
{
    if (ValidarCampos())
    {
		parent.pantalla.document.albpedpro_facpro_param.action="albpedpro_facpro_paramResultado.asp?mode=select2";
		parent.pantalla.document.albpedpro_facpro_param.submit();
		document.location="albpedpro_facpro_param_bt.asp?mode=select2";
	}
	else document.location="albpedpro_facpro_param_bt.asp?mode=select1";
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "select1":
			switch (pulsado) {
				case "select2": //Aceptar
					document.location="albpedpro_facpro_param_bt.asp?mode=recarga" ;
					break;
				case "select1": //Cancelar
					parent.pantalla.document.albpedpro_facpro_param.action="albpedpro_facpro_param.asp?mode=" + pulsado;
					parent.pantalla.document.albpedpro_facpro_param.submit();
					document.location="albpedpro_facpro_param_bt.asp?mode=" + pulsado;
					break;
			}
			break;
case "select2":
    switch (pulsado) {
        case "todos": //Seleccionar todos los registros
            nregistros = parent.pantalla.document.albpedpro_facpro_param.h_nfilas.value;
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                parent.pantalla.document.albpedpro_facpro_param.elements[nombre].checked = true;
            }
            break;

        case "ninguno": //No seleccionar ningun registro
            nregistros = parent.pantalla.document.albpedpro_facpro_param.h_nfilas.value;
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                parent.pantalla.document.albpedpro_facpro_param.elements[nombre].checked = false;
            }
            break;
        case "confirm": //Aceptar
            if (parent.pantalla.document.albpedpro_facpro_param.viene.value == "facturas_pro") {
                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                parent.pantalla.document.albpedpro_facpro_param.action = "albpedpro_facpro_paramResultado.asp?mode=confirm";
                parent.pantalla.document.albpedpro_facpro_param.submit();
                document.location = "albpedpro_facpro_param_bt.asp?mode=select1";
            }
            else {
                if (ValidarCampos2()) {
                    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    parent.pantalla.document.albpedpro_facpro_param.action = "albpedpro_facpro_paramResultado.asp?mode=confirm";
                    parent.pantalla.document.albpedpro_facpro_param.submit();
                    nregistros = parent.pantalla.document.albpedpro_facpro_param.h_nfilas.value;
                    allChecked = false;
                    cont = 0;
                    for (i = 1; i <= nregistros; i++) {
                        nombre = "check" + i;
                        if (parent.pantalla.document.albpedpro_facpro_param.elements[nombre].checked) cont++;
                    }
                    if (cont == nregistros) allChecked = true;
                    if (allChecked) document.location = "albpedpro_facpro_param_bt.asp?mode=select1";
                }
            }
            break;
        case "select1": //Cancelar
            parent.pantalla.document.albpedpro_facpro_param.action = "albpedpro_facpro_param.asp?mode=select1";
            parent.pantalla.document.albpedpro_facpro_param.submit();
            document.location = "albpedpro_facpro_param_bt.asp?mode=select1";
            break;
    }
    break;
		case "imp":
			switch (pulsado) {
				case "cancel": //Volver atrás
					parent.pantalla.document.location=history.back();
					history.back();
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
		    <%if mode="select1" then%>
				<td id="idaccept" class="CELDABOT" onclick="javascript:Accion('select1','select2');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('select1','select1');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
    		<%elseif mode="select2" then%>
				<td id="idSelectAll" class="CELDABOT" onclick="javascript:Accion('select2','todos');">
					<%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,LITBOTSELTODOTITLE%>
				</td>
				<td id="idSelectNothing" class="CELDABOT" onclick="javascript:Accion('select2','ninguno');">
					<%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,LITBOTDSELTODOTITLE%>
				</td>
				<td id="idaccept" class="CELDABOT" onclick="javascript:Accion('select2','confirm');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('select2','select1');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="imp" then%>
				<td id="idreturn" class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
					<%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				</td>
			<%elseif mode="recarga" then%>
			  <script language="javascript" type="text/javascript">to_select2();</script>
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
