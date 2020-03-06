<%@ Language=VBScript %>
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="cobros_param.inc" -->

<!--#include file="../styles/Master.css.inc" -->
<!--#include file="../styles/FootButton.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
//Validación de campos numéricos y fechas.
function ValidarCampos()
{
	if (!checkdate(parent.pantalla.document.cobros_param.Dfecha))
	{
		window.alert("<%=LitMsgFechaMal%>");
		return false;
	}
	if (parent.pantalla.document.cobros_param.Dfecha.value=="")
	{
		window.alert("<%=LitMsgDesdeFechaNoNulo%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.cobros_param.Hfecha))
	{
		window.alert("<%=LitMsgFechaMal%>");
		return false;
	}
	if (parent.pantalla.document.cobros_param.Hfecha.value=="")
	{
		window.alert("<%=LitMsgHastaFechaNoNulo%>");
		return false;
	}
	var diferencia = DiferenciaTiempo(parent.pantalla.document.cobros_param.Hfecha.value,parent.pantalla.document.cobros_param.Dfecha.value, "dias");
	if (diferencia>365)
	{
		window.alert("<%=LitDiferenciaFechas%>");
		return false;
	}
	/*
	if (parent.pantalla.document.cobros_param.serie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}*/
	return true;
}

function ValidarCampos2()
{
	if (parent.pantalla.document.cobros_param.fechafactura.value=="")
	{
		window.alert("<%=LitMsgFechaNoNulo%>");
		return false;
	}
	return true;
}

function ValidarCampos3()
{
	if (parent.pantalla.document.cobros_paramResultado.fechacobro.value=="")
	{
		window.alert('<%=LitMsgFechaCobroNoNulo%>');
		return false;
	}
	if (!checkdate(parent.pantalla.document.cobros_paramResultado.fechacobro))
	{
		window.alert("<%=LitMsgFechaMal%>");
		return false;
	}
	if (parent.pantalla.document.cobros_paramResultado.h_tabla.value=="vencimientos_salida")
	{
		nregistros=parent.pantalla.document.cobros_paramResultado.h_nfilas.value;
		i=1;
		no_continuar=0;
		while (i<=nregistros && no_continuar==0)
		{
			nombre="importecob" + i;
			if (isNaN(parent.pantalla.document.cobros_paramResultado.elements[nombre].value.replace(",","."))==true) no_continuar=1;
			i++;
		}
		if (no_continuar==1)
		{
			window.alert('<%=LitImporteCobNoNumero%>');
			return false;
		}
	}
	return true;
}
momentoActual = new Date() 
hora = momentoActual.getHours() 
minuto = momentoActual.getMinutes() 
segundo = momentoActual.getSeconds() 
horaImprimible = hora + " : " + minuto + " : " + segundo 


//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado)
{
	switch (mode)
	{
	    case "browse":
	        switch (pulsado) {
	            case "todos": //Seleccionar todos los registros
	                nregistros = parent.pantalla.document.cobros_paramResultado.h_nfilas.value;
	                for (i = 1; i <= nregistros; i++) {
	                    nombre = "check" + i;
	                    if (parent.pantalla.document.cobros_paramResultado.elements[nombre].checked == false) {
	                        parent.pantalla.document.cobros_paramResultado.elements[nombre].checked = true;
	                        parent.pantalla.cambiar_importecob(i);
	                    }
	                }
	                parent.pantalla.document.cobros_paramResultado.check.checked = true;
	                break;

	            case "ninguno": //Editar registro
	                nregistros = parent.pantalla.document.cobros_paramResultado.h_nfilas.value;
	                for (i = 1; i <= nregistros; i++) {
	                    nombre = "check" + i;
	                    if (parent.pantalla.document.cobros_paramResultado.elements[nombre].checked == true) {
	                        parent.pantalla.document.cobros_paramResultado.elements[nombre].checked = false;
	                        parent.pantalla.cambiar_importecob(i);
	                    }
	                }
	                //FLM:20090529:al desmarcar el check de todos el importe debe ser 0. Lo ponemos a 0 y actualizamos el total.
	                parent.pantalla.totalImporteCobrar = 0;
	                parent.pantalla.document.getElementById("totalACobrar").innerHTML = truncar(0.00, parent.pantalla.numDecimalesEmpresa);

	                parent.pantalla.document.cobros_paramResultado.check.value = "xxx";
	                parent.pantalla.document.cobros_paramResultado.check.checked = false;
	                break;

	            case "cobrar": //Realizar el cobro
	                //document.all("boton_cobrar").href="";
	                //document.all("boton_cobrar2").style.display="none";
                   var peru = parent.pantalla.document.cobros_paramResultado.peru.value;
                    
                   if(peru==0){
	                if(confirm("<%=LITCONFIRMCOBRO%>")) {
	                    if (ValidarCampos3()) {
	                        if (parent.pantalla.document.cobros_paramResultado.ncaja.value == "") {
	                            if (parent.pantalla.document.cobros_paramResultado.h_tabla.value == "tickets_cli") alert("<%=LitCajaVacia%>");
	                            else {
	                                if (window.confirm("<%=LitMsgCobroSinCajaConfirm%>")) {
	                                    if (window.confirm("<%=LitMsgCobrarConfirm%>") == true) {
	                                        if ((parent.pantalla.document.cobros_paramResultado.ncaja.value != "") && (parent.pantalla.document.cobros_paramResultado.i_pago.value == ""))
	                                            window.alert("<%=LitMsgTipoPagoNoNulo%>");
	                                        else {
	                                            parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
	                                            parent.pantalla.document.cobros_paramResultado.action = "cobros_param.asp?mode=save";
	                                            parent.pantalla.document.cobros_paramResultado.submit();
	                                            document.location = "cobros_param_bt.asp?mode=browse";
	                                        }
	                                    }
	                                }
	                            }
	                        }
	                        else {
	                            if ((parent.pantalla.document.cobros_paramResultado.ncaja.value != "") && (parent.pantalla.document.cobros_paramResultado.i_pago.value == ""))
	                                window.alert("<%=LitMsgTipoPagoNoNulo%>");
	                            else {
	                                if (parent.pantalla.document.cobros_paramResultado.h_tabla.value == "tickets_cli")
	                                {
	                                    momentoActual = new Date() 
	                                    hora = momentoActual.getHours() 
	                                    minuto = momentoActual.getMinutes() 
	                                    segundo = momentoActual.getSeconds() 
	                                    horaImprimible = hora + ":" + minuto + ":" + segundo 

	                                    parent.pantalla.document.cobros_paramResultado.fechacobro.value=parent.pantalla.document.cobros_paramResultado.fechacobro.value+" "+horaImprimible;
	                                    
	                                }
	                                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
	                                parent.pantalla.document.cobros_paramResultado.action = "cobros_param.asp?mode=save";
	                                parent.pantalla.document.cobros_paramResultado.submit();
	                                document.location = "cobros_param_bt.asp?mode=browse";
	                            }
	                        }
	                    }
	                    else {
	                        if (parent.pantalla.document.cobros_paramResultado.h_tabla.value != "tickets_cli")
	                            document.all("boton_cobrar2").style.display = "";
	                    }
	                }
                   }
                    else{
                            if (ValidarCampos3()) {
	                        if (parent.pantalla.document.cobros_paramResultado.ncaja.value == "") {
	                            if (parent.pantalla.document.cobros_paramResultado.h_tabla.value == "tickets_cli") alert("<%=LitCajaVacia%>");
	                            else {
	                                if (window.confirm("<%=LitMsgCobroSinCajaConfirm%>")) {
	                                    if (window.confirm("<%=LitMsgCobrarConfirm%>") == true) {
	                                        if ((parent.pantalla.document.cobros_paramResultado.ncaja.value != "") && (parent.pantalla.document.cobros_paramResultado.i_pago.value == ""))
	                                            window.alert("<%=LitMsgTipoPagoNoNulo%>");
	                                        else {
	                                            parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
	                                            parent.pantalla.document.cobros_paramResultado.action = "cobros_param.asp?mode=save";
	                                            parent.pantalla.document.cobros_paramResultado.submit();
	                                            document.location = "cobros_param_bt.asp?mode=browse";
	                                        }
	                                    }
	                                }
	                            }
	                        }
	                        else {
	                            if ((parent.pantalla.document.cobros_paramResultado.ncaja.value != "") && (parent.pantalla.document.cobros_paramResultado.i_pago.value == ""))
	                                window.alert("<%=LitMsgTipoPagoNoNulo%>");
	                            else {
	                                if (parent.pantalla.document.cobros_paramResultado.h_tabla.value == "tickets_cli")
	                                {
	                                    momentoActual = new Date() 
	                                    hora = momentoActual.getHours() 
	                                    minuto = momentoActual.getMinutes() 
	                                    segundo = momentoActual.getSeconds() 
	                                    horaImprimible = hora + ":" + minuto + ":" + segundo 

	                                    parent.pantalla.document.cobros_paramResultado.fechacobro.value=parent.pantalla.document.cobros_paramResultado.fechacobro.value+" "+horaImprimible;
	                                    alert( parent.pantalla.document.cobros_paramResultado.fechacobro.value)
	                                }
	                                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
	                                parent.pantalla.document.cobros_paramResultado.action = "cobros_param.asp?mode=save";
	                                parent.pantalla.document.cobros_paramResultado.submit();
	                                document.location = "cobros_param_bt.asp?mode=browse";
	                            }
	                        }
	                    }
	                    else {
	                        if (parent.pantalla.document.cobros_paramResultado.h_tabla.value != "tickets_cli")
	                            document.all("boton_cobrar2").style.display = "";
	                    }

            
                        }
	                break;

	            case "cancelar": //Cancelar operacion
	                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
	                caju = parent.pantalla.document.cobros_paramResultado.caju.value
	                parent.pantalla.document.location = "cobros_param.asp?mode=add&caju=" + caju;
	                document.location = "cobros_param_bt.asp?mode=add";
	                break;
	        }
	        break;

		case "edit":
			switch (pulsado) 
			{
				case "save": //Guardar registro
					if (ValidarCampos())
					{
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						parent.pantalla.document.cobros_param.action="cobros_param.asp?npedido=" + parent.pantalla.document.cobros_param.h_npedido.value +
						"&mode=save";
						parent.pantalla.document.cobros_param.submit();
						document.location="cobros_param_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.cobros_param.action="cobros_param.asp?npedido=" + parent.pantalla.document.cobros_param.h_npedido.value +
					"&mode=browse";
					parent.pantalla.document.cobros_param.submit();
					document.location="cobros_param_bt.asp?mode=browse";
					break;
			}
			break;

		case "add":
			switch (pulsado)
			{
				case "save": //Guardar registro
					if (ValidarCampos())
					{
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						parent.pantalla.document.cobros_param.action="cobros_paramResultado.asp?mode=browse";
						parent.pantalla.document.cobros_param.submit();
						document.location="cobros_param_bt.asp?mode=browse";
					}
					break;
			}
			break;
		case "imp":
			switch (pulsado)
			{
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
<div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_ASP" >
        <table id="BUTTONS_CENTER_ASP" >
		<tr>
			<%if mode="browse" then%>
				<td id="idSelectAll" class="CELDABOT" onclick="javascript:Accion('browse','todos');">
					<%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,LITBORSELTODOTITLE%>
				</td>
				<td id="idSelectNothing" class="CELDABOT" onclick="javascript:Accion('browse','ninguno');">
					<%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,LITBOTDSELTODOTITLE%>
				</td>
				<td id="idcharge" class="CELDABOT" id="boton_cobrar" onclick="javascript:Accion('browse','cobrar');">
					<%PintarBotonBT LITBOTCOBRAR,ImgCobrar,ParamImgCobrar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('browse','cancelar');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="search" then%>
				<td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					<%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				</td>
			<%elseif mode="edit" then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					<%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="add" then%>
				<td id="idSelectAll" class="CELDABOT" onclick="javascript:Accion('add','save');">
					<%PintarBotonBT LITBOTSELDOCU,ImgSelecc_doc,ParamImgSelecc_doc,LITBOTSELDOCUTITLE%>
				</td>
			<%elseif mode="imp" then%>
				<td id="idreturn" class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
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
    </table></form>
</body>
</html>