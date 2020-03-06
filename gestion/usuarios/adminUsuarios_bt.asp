<%@ Language=VBScript %>
<% Response.Expires= 0%>
<script id="DebugDirectives" runat="server" language="javascript">
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../styles/Master.css.inc" -->

<!--#include file="adminUsuarios.inc" -->
</head>
<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1);

function Imprimir() {
	if (pr) //NS4, IE5
		parent.pantalla.print()
		//vbImprimir()
	else if (da && !mac) // IE4 (Windows)
		//vbImprimir()
		alert("<%=LitNoImprime%>");
	else // Otros Navegadores
		alert("<%=LitNoImprime%>");
	return false;
}

    function Buscar()
    {
        var errorsDetected = false;

        if (parent.pantalla.document.adminUsuarios.hnusuario1 == null)
        {
            errorsDetected = true;
        }

        if (parent.pantalla.document.adminUsuarios.hnusuario2 == null)
        {
            errorsDetected = true;
        }
        
        if (parent.pantalla.document.adminUsuarios.hnusuarios == null)
        {
            errorsDetected = true;
        }
        
        if (parent.pantalla.document.adminUsuarios.hnmodulos == null)
        {
            errorsDetected = true;
        }
        
        if (parent.pantalla.document.adminUsuarios.hOMC == null)
        {
            errorsDetected = true;
        }
        
        if (parent.pantalla.document.adminUsuarios.ver == null)
        {
            errorsDetected = true;
        }

        var campos = document.opciones.campos.value;
        var criterio = document.opciones.criterio.value;
        var opciones = document.opciones.texto.value;

        if (!errorsDetected)
        {
            parent.pantalla.document.adminUsuarios.action = "adminUsuarios.asp?hnusuario1=" + parent.pantalla.document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + parent.pantalla.document.adminUsuarios.hnusuario2.value + "&ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value + "&mode=search&campo=" + document.opciones.campos.value + "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&nusuarios=" + parent.pantalla.document.adminUsuarios.hnusuarios.value + "&nmodulos=" + parent.pantalla.document.adminUsuarios.hnmodulos.value + "&OMC=" + parent.pantalla.document.adminUsuarios.hOMC.value + "&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
        }
        else
        {
            parent.pantalla.document.adminUsuarios.action = "adminUsuarios.asp?ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value + "&mode=search&campo=" + document.opciones.campos.value + "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value;
        }

        parent.pantalla.document.adminUsuarios.submit();
        document.location = "adminUsuarios_bt.asp?mode=edit&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
	    
    }

function comprobar_enter(e)
{
    //var keycode = e.keyCode;
	//si se ha pulsado la tecla enter
	//if (keycode==13) {
        document.opciones.criterio.focus();
        Buscar();
    //}
}

<% ' >>> MCA 23/12/04 : Función para validar las franjas horarias del formulario dinámico de usuarios %>
function ValidarFranjasHorarias(nusuario1,nusuario2) 
{
	var campo_desdehora1, campo_hastahora1, campo_desdehora2, campo_hastahora2;
	var desdehora1, hastahora1, desdehora2, hastahora2;
	var i, mensaje= "";

	for (i=nusuario1;i<=nusuario2;i++)
	{
		campo_desdehora1= eval("parent.pantalla.document.adminUsuarios.desdehora1_u"+ i);
		desdehora1= campo_desdehora1.value;
		campo_hastahora1= eval("parent.pantalla.document.adminUsuarios.hastahora1_u"+ i);
		hastahora1= campo_hastahora1.value;
		campo_desdehora2= eval("parent.pantalla.document.adminUsuarios.desdehora2_u"+ i);
		desdehora2= campo_desdehora2.value;		
		campo_hastahora2= eval("parent.pantalla.document.adminUsuarios.hastahora2_u"+ i);
		hastahora2= campo_hastahora2.value;				

		if (desdehora1.length>0 && hastahora1=="") {
			window.alert("<%=LitFranjaHorariaIniFin%>");
			campo_hastahora1.focus();
			return false;
		}
		
		if (desdehora1=="" && hastahora1.length>0) {
			window.alert("<%=LitFranjaHorariaIniFin%>");
			campo_desdehora1.focus();
			return false;
		}
		if (desdehora2.length>0 && hastahora2=="") {
			window.alert("<%=LitFranjaHorariaIniFin%>");
			campo_hastahora2.focus();
			return false;
		}
		if (desdehora2=="" && hastahora2.length>0) {
			window.alert("<%=LitFranjaHorariaIniFin%>");
			campo_desdehora2.focus();
			return false;
		}
	
		if (!checkFormatohoraValue(desdehora1)) {
			mensaje= " <%=LitCompruebaHora%> "+ desdehora1 +" <%=LitCompruebaHora1%>";
			window.alert(mensaje);
			return false;
		}
		else {
			if (desdehora1.length>0 && desdehora1.length<=2) {
				campo_desdehora1.value= desdehora1 + ":00";
				desdehora1= desdehora1 + ":00";
			}	
			var horaentrada1= convertir_horatxt_horams(desdehora1);			
		}
		
		if (!checkFormatohoraValue(hastahora1)) {
			mensaje= " <%=LitCompruebaHora%> "+ hastahora1 +" <%=LitCompruebaHora1%>";
			window.alert(mensaje);
			return false;
		}
		else {
			if (hastahora1.length>0 && hastahora1.length<=2) {
				campo_hastahora1.value= hastahora1 + ":00";
				hastahora1= hastahora1 + ":00";
			}
			var horasalida1= convertir_horatxt_horams(hastahora1);			
		}
			
		if (horasalida1 <= horaentrada1)
		{
			mensaje= " <%=LitHoraFinMayorHoraIni %>";
			window.alert(mensaje);
			campo_hastahora1.focus();			
			return false;
		}
	
		if (!checkFormatohoraValue(desdehora2)) {
			mensaje= " <%=LitCompruebaHora%> "+ desdehora2 +" <%=LitCompruebaHora1%>";
			window.alert(mensaje);
			return false;
		}
		else {
			if (desdehora2.length>0 && desdehora2.length<=2) {
				campo_desdehora2.value= desdehora2 + ":00";
				desdehora2= desdehora2 + ":00";
			}
			var horaentrada2= convertir_horatxt_horams(desdehora2);			
		}
			
		if (!checkFormatohoraValue(hastahora2)) {
			mensaje= " <%=LitCompruebaHora%> "+ hastahora2 +" <%=LitCompruebaHora1%>";
			window.alert(mensaje);
			return false;
		}
		else {
			if (hastahora2.length>0 && hastahora2.length<=2) {
				campo_hastahora2.value=  hastahora2 + ":00";
				hastahora2= hastahora2 + ":00";
			}
			var horasalida2= convertir_horatxt_horams(hastahora2);			
		}
	
		if (horasalida2 <= horaentrada2)		
		{
			mensaje= " <%=LitHoraFinMayorHoraIni%>";
			window.alert(mensaje);
			campo_hastahora2.focus();			
			return false;
		}
	}

	return true;
}
<% ' <<< MCA 23/12/04 : Función para validar las franjas horarias del formulario dinámico de usuarios %>

//Validación de campos numéricos y fechas.
function ValidarCampos(mode) {
	if (parent.pantalla.document.adminUsuarios.cliente.value=="") {
		window.alert("<%=LitCampoClienteNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.adminUsuarios.nuser.value=="") {
		window.alert("<%=LitCampoNomUsuNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.adminUsuarios.pwd.value=="") {
		window.alert("<%=LitCampoPassNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.adminUsuarios.copwd.value=="") {
		window.alert("<%=LitCampoConfirmPassNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.adminUsuarios.claves.value=="") {
		window.alert("<%=LitCampoClavesNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.adminUsuarios.pwd.value!=parent.pantalla.document.PagUsuario.copwd.value) {
		window.alert("<%=LitPassYConfirmDistintas%>");
		return false;
	}
	return true;
}

function ValidarCampos2(mode) {
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	var hmode,nusuario1,nusuario2,nusuarios;
	hmode= parent.pantalla.document.adminUsuarios.hmode.value;

    if (parent.pantalla.document.adminUsuarios.sinUsuarios.value!="SI") {
	    switch (mode) {
		    case "browse":
			    switch (pulsado) {
				    case "save": //Guardar
					    if (hmode=="gestionhorarios" || hmode=="guardahorarios")
					    {
						    nusuario1= parent.pantalla.document.adminUsuarios.hnusuario1.value;
						    nusuario2= parent.pantalla.document.adminUsuarios.hnusuario2.value;
						    if (ValidarFranjasHorarias(nusuario1,nusuario2))
						    {
					            parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
				                parent.pantalla.document.getElementById("waitBoxOculto").focus();

							    parent.pantalla.document.adminUsuarios.action="adminUsuarios.asp?mode=guardahorarios&hnusuario1=" + parent.pantalla.document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + parent.pantalla.document.adminUsuarios.hnusuario2.value + "&ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value +
							    "&nusuarios=" + parent.pantalla.document.adminUsuarios.hnusuarios.value + "&nmodulos=" + parent.pantalla.document.adminUsuarios.hnmodulos.value + "&lote=" + parent.pantalla.document.adminUsuarios.hlote.value +
							    "&campo=" + parent.pantalla.document.adminUsuarios.hcampo.value + "&criterio=" + parent.pantalla.document.adminUsuarios.hcriterio.value + "&texto=" + parent.pantalla.document.adminUsuarios.htexto.value +
							    "&OMC=" + parent.pantalla.document.adminUsuarios.hOMC.value + 
							    "&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
							    parent.pantalla.document.adminUsuarios.submit();
							    document.location="adminUsuarios_bt.asp?mode=save&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
						    }
						    break;
					    }
					    else
					    {
					        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
				            parent.pantalla.document.getElementById("waitBoxOculto").focus();

						    parent.pantalla.document.adminUsuarios.action="adminUsuarios.asp?hnusuario1=" + parent.pantalla.document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + parent.pantalla.document.adminUsuarios.hnusuario2.value + "&ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value +
						    "&mode=" + pulsado + "&nusuarios=" + parent.pantalla.document.adminUsuarios.hnusuarios.value + "&nmodulos=" + parent.pantalla.document.adminUsuarios.hnmodulos.value + "&lote=" + parent.pantalla.document.adminUsuarios.hlote.value +
						    "&campo=" + parent.pantalla.document.adminUsuarios.hcampo.value + "&criterio=" + parent.pantalla.document.adminUsuarios.hcriterio.value + "&texto=" + parent.pantalla.document.adminUsuarios.htexto.value +
						    "&OMC=" + parent.pantalla.document.adminUsuarios.hOMC.value + 
						    "&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
						    parent.pantalla.document.adminUsuarios.submit();
						    document.location="adminUsuarios_bt.asp?mode=save&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
						    break;
					    }

				    case "cancel": //Cancelar

					    if (hmode=="gestionhorarios" || hmode=="guardahorarios")
					    {
						    parent.pantalla.document.adminUsuarios.action="adminUsuarios.asp?mode=gestionhorarios&ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value +
							    "&nusuarios=" + parent.pantalla.document.adminUsuarios.hnusuarios.value + "&nmodulos=" + parent.pantalla.document.adminUsuarios.hnmodulos.value +
							    "&OMC=" + parent.pantalla.document.adminUsuarios.hOMC.value + 
							    "&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
						    parent.pantalla.document.adminUsuarios.submit();
						    document.location="adminUsuarios_bt.asp?mode=save&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
						    break;
					    }
					    else
					    {		
						    parent.pantalla.document.adminUsuarios.action="adminUsuarios.asp?ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value +
							    "&mode=edit" + "&nusuarios=" + parent.pantalla.document.adminUsuarios.hnusuarios.value + "&nmodulos=" + parent.pantalla.document.adminUsuarios.hnmodulos.value +
							    "&OMC=" + parent.pantalla.document.adminUsuarios.hOMC.value + 
							    "&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
						    parent.pantalla.document.adminUsuarios.submit();
						    document.location="adminUsuarios_bt.asp?mode=save&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
						    break;
					    }
    					
				    case "adminsave": //Guardar
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
				        parent.pantalla.document.getElementById("waitBoxOculto").focus();

					    parent.pantalla.document.adminUsuarios.action="adminUsuarios.asp?hnusuario1=" + parent.pantalla.document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + parent.pantalla.document.adminUsuarios.hnusuario2.value + "&ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value +
						    "&mode=adminsave&nusuarios=" + parent.pantalla.document.adminUsuarios.hnusuarios.value + "&nmodulos=" + parent.pantalla.document.adminUsuarios.hnmodulos.value + "&lote=" + parent.pantalla.document.adminUsuarios.hlote.value +
						    "&campo=" + parent.pantalla.document.adminUsuarios.hcampo.value + "&criterio=" + parent.pantalla.document.adminUsuarios.hcriterio.value + "&texto=" + parent.pantalla.document.adminUsuarios.htexto.value + 
						    "&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
					    parent.pantalla.document.adminUsuarios.submit();
					    document.location="adminUsuarios_bt.asp?mode=adminsave&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
					    break;

				    case "admincancel": //Cancelar
					    parent.pantalla.document.adminUsuarios.action="adminUsuarios.asp?ncliente=" + parent.pantalla.document.adminUsuarios.hncliente.value +
						    "&mode=adminedit" + "&nusuarios=" + parent.pantalla.document.adminUsuarios.hnusuarios.value + "&nmodulos=" + parent.pantalla.document.adminUsuarios.hnmodulos.value + 
						    "&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
					    parent.pantalla.document.adminUsuarios.submit();
					    document.location="adminUsuarios_bt.asp?mode=adminsave&ver=" + parent.pantalla.document.adminUsuarios.ver.value;
					    break;

				    case "search": //Buscar datos
					    break;
			    }
			    break;

		    case "search":
			    switch (pulsado) {
				    case "search": //Buscar datos
					    break;
			    }
			    break;
	    }
    }
}
</script>
<body class="body_master_ASP">
<%mode=	Request.QueryString("mode")
OMC=	limpiaCadena(Request.QueryString("OMC"))
hmode= 	limpiaCadena(Request.QueryString("hmode"))
ver=    limpiaCadena(Request.QueryString("ndoc"))
if ver & "" = "" then ver = limpiaCadena(request.QueryString("ver"))%>
<form name="opciones" method="post">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
<div id="PageFooter_ASP" >

<%if OMC<>"SI" or hmode="gestionhorarios" or hmode="guardahorarios" then%>
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
		    <%if mode="save" and ver <> "1"  then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('browse','save');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('browse','cancel');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				</td>
			<%end if
			if mode="edit" and ver <> "1"  then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('browse','save');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('browse','cancel');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				</td>
			<%end if
			if mode="adminsave" and ver <> "1"  then%>
			    <td id="idsave" class="CELDABOT" onclick="javascript:Accion('browse','adminsave');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('browse','admincancel');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				</td>
			<%end if
			if mode="adminedit" and ver <> "1" then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('browse','adminsave');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('browse','admincancel');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				</td>
			<%end if%>
	</table>
    </div>
<%end if%>

    
    <div id="FILTERS_MASTER_ASP">
		<!--<td class=CELDABOT><%=LitBuscar & ": "%>-->
				<select class="IN_S" name="campos">
		  			<option selected="selected" value="nombre"><%=LitUser%></option>
        		</select>
			<!--</td>
			<td class=CELDABOT>-->
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
				</select>
			<!--</td><td class=CELDABOT>-->
                <input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
			<!--</td><td class=CELDABOT>-->
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			<!--</td>
		</tr>
	</table>-->
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