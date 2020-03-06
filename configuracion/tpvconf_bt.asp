<%@ Language=VBScript %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">-->
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="tpvconf.inc" -->
<title><%=LitTitulo%></title>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../cache.inc" -->
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
	    case "browse":
	        switch (pulsado) {
	            case "edit":
	                parent.pantalla.document.tpvconf.action="tpvconf.asp?mode=edit";
	                parent.pantalla.document.tpvconf.submit();
	                document.location="tpvconf_bt.asp?mode=edit";
	                break;
	        }
	        break;
		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
	                parent.pantalla.document.tpvconf.action="tpvconf.asp?mode=save";
	                parent.pantalla.document.tpvconf.submit();
	                document.location="tpvconf_bt.asp?mode=save";
					break;
			}
			break;
		case "asistente":
			switch (pulsado) {
			    case "edit":
	                parent.pantalla.document.tpvconf.action="tpvconf.asp?mode=edit&viene=asistente";
	                parent.pantalla.document.tpvconf.submit();
	                document.location="tpvconf_bt.asp?mode=edit&viene=asistente";
	                break;
				case "save": //Guardar registro
	                parent.pantalla.document.tpvconf.action="tpvconf.asp?mode=save&viene=asistente";
	                parent.pantalla.document.tpvconf.submit();
	                document.location="tpvconf_bt.asp?mode=browse&viene=asistente";
					break;
				case "cancel":
	                parent.pantalla.document.tpvconf.action="tpvconf.asp?mode=edit&viene=asistente";
	                parent.pantalla.document.tpvconf.submit();
	                document.location="tpvconf_bt.asp?mode=browse&viene=asistente";
	                break;
			}
			break;
	}
}

function Buscar() {
		parent.pantalla.fr_Tabla.document.tpvconf_det.action="tpvconf_det.asp?mode=search&lote=1&campo=" + document.opciones.campos.value +
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value;
		parent.pantalla.fr_Tabla.document.tpvconf_det.submit();
		document.location="tpvconf_bt.asp?mode=browse";
}

function Cancelar(mode) {
        parent.pantalla.document.tpvconf.action="tpvconf.asp?mode=browse";
        parent.pantalla.document.tpvconf.submit();
        document.location="tpvconf_bt.asp?mode=browse";
}

//****************************************************************************************************************
function comprobar_enter(e)
{
    //var keycode = e.keyCode;
	//si se ha pulsado la tecla enter
	//if (keycode==13){
		document.opciones.criterio.focus();
		Buscar();
	//}
}

function MoverPagPM(ruta,rbt)
{
    if (rbt=="1") parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"1";        
  	else parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"2";	      
}

function Cerrar(){
    parent.location="../Applets/asistentePM.asp?mode=cancel";
}
</script>

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<body class="body_master_ASP">
<%viene=enc.EncodeForJavascript(Request.QueryString("viene"))                                 
 tipo="1"
 ruta=ObtenerNombreFichero(tipo) 
%>
<form name="opciones" method="post" action="">
    <%if request("mode")="edit" then                                   
	    mode="edit"
    else
        mode="browse"
    end if%>
    
    <%if viene<>"asistente" then %>
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
            <%if mode="browse" then%>
				<td id="ideedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					<%PintarBotonBTLeft LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				</td> 
			<%end if
			if mode="edit" then%>
		        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
				    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
			    </td>
			<%end if%>
		    <td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar();">
			    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
		    </td>
        </tr>
	</table>
    </div>
    
    <div id="FILTERS_MASTER_ASP">
			<!--<td class="CELDABOT">-->
				<select class="IN_S" name="campos">
					<option value="tpv"><%=LitTpv%></option>
					<option value="caja"><%=LitCaja%></option>
					<option selected value="descripcion"><%=LitDescripcion%></option>
					<option value="estado"><%=LitEstado%></option>
				</select>
			<!--</td>-->
			<!--<td class="CELDABOT">-->
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<!--<option value="empieza"><%=LitComienza%></option>-->
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
				</select>
			<!--</td>-->
			<!--<td class="CELDABOT">-->
                <input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			<!--</td>-->
		<!--</tr>
	</table>-->
    </div>
    </div>

    <script type="text/javascript" language="javascript">
    /*
	    if (document.opciones.texto != null)
	    {
	        function texto_callkeydownhandler(evnt)
            {
                ev = (evnt) ? evnt : event;
                comprobar_enter(ev);
            }

            if(window.document.opciones.texto.addEventListener)
            {
                window.document.opciones.texto.addEventListener("keydown", texto_callkeydownhandler, false);
            }
            else
            {
                window.document.opciones.texto.attachEvent("onkeydown", texto_callkeydownhandler);
            }
        }
        */
	</script>
	<%else%>
    <div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_left_ASP" >
        <table id="BUTTONS_CENTER_ASP">
		    <tr>
                <%if mode="browse" then%>   
				    <td id="ideedit" class="CELDABOT" onclick="javascript:Accion('asistente','edit');">
					    <%PintarBotonBTLeft LITBOTEDITAR,ImgEditar,ParamImgEditar,""%>
				    </td>            
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Cerrar();">
					    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				    </td>    
				    <td id="idprevious" class="CELDABOT" onclick="javascript:MoverPagPM('<%=ruta%>','2');">
					    <%PintarBotonBTLeft LITBOTANTERIOR,ImgAnterior,ParamImgAnterior,""%>
				    </td> 
				    <td id="idnext" class="CELDABOT" onclick="javascript:MoverPagPM('<%=ruta%>','1');">
					    <%PintarBotonBTLeft LITBOTSIGUIENTE,ImgSiguiente,ParamImgSiguiente,""%>
				    </td> 
			    <%end if
			    if mode="edit" then%>
				    <td id="idsave" class="CELDABOT" onclick="javascript:Accion('asistente','save');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel');">
					    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
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
</HTML>