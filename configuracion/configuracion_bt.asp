<%@ Language=VBScript %>
<!--#include file="../cache.inc" -->
<SCRIPT id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</SCRIPT> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script language="JavaScript">
var ret_validateBin;


function getHTTPObject() {
    var xmlhttp;
    if (!xmlhttp && typeof XMLHttpRequest != 'undefined') {
        try {
            xmlhttp = new XMLHttpRequest();
        }
        catch (e) { xmlhttp = false; }
    }
    return xmlhttp;
}

var enProceso = false; // lo usamos para ver si hay un proceso activo
var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest

function handleHttpResponse() {
    if (http.readyState == 4) {
        if (http.status == 200) {
            if (http.responseText.indexOf('invalid') == -1) {
                results = http.responseText;
                enProceso = false;
                ret_validateBin = unescape(results)
                document.opciones.vbin.value = ret_validateBin;
                if (navigator.userAgent.indexOf("Firefox") > 0) {
                    document.location = "configuracion_bt.asp?mode=save&ok=" + ret_validateBin;
                }

            } 
        } 
    } 
}


//Validación de campos numéricos y fechas.
function ValidarCampos() {
    si_tiene_paginaSMS=parent.pantalla.document.configuracion.si_tiene_paginaSMS.value;
    if (si_tiene_paginaSMS!=0){
	    if (parent.pantalla.document.configuracion.mensajeria_sms.checked && parent.pantalla.document.configuracion.mensajeria_smsbd.value=="0") {
   		    if (!confirm("<%=LitAceptandoCondiciones%>")) return false;
	    }
    }
    /*if (parent.pantalla.document.configuracion.preg_ped_bis.checked==false){
        if (parent.pantalla.document.configuracion.serie.value==""){
	        window.alert("<%=LitMsgSerieNoNulo%>");
		    return false;
	    }
    }*/
    if (isNaN(parent.pantalla.document.configuracion.recargo.value.replace(",",".")) || isNaN(parent.pantalla.document.configuracion.preciokm.value.replace(",",".")) || isNaN(parent.pantalla.document.configuracion.deccantidades.value.replace(",",".")) || isNaN(parent.pantalla.document.configuracion.decprecios.value.replace(",","."))) {
   	    window.alert("<%=LitDatosNumericos%>");
		return false;
    }
    if (isNaN(parent.pantalla.document.configuracion.contgencodbarras.value)) {
   		window.alert("<%=LitcontgencodbarrasMal%>");
		return false;
    }
    if ((parent.pantalla.document.configuracion.contgencodbarras.value<-2147483648)||(parent.pantalla.document.configuracion.contgencodbarras.value>2147483648)) {
   		window.alert("<%=LitcontgencodbarrasInt%>");
		return false;
    }

    si_tiene_modulo_tiendas=parent.pantalla.document.configuracion.si_tiene_modulo_tiendas.value;
    if (si_tiene_modulo_tiendas!=0){
	    if (comp_car_ext(parent.pantalla.document.configuracion.leyticketcabecera.value,0)==1){
			window.alert("<%=LitMsgConf1DesCarNoVal%>");
			return false;
		}

	   if (comp_car_ext(parent.pantalla.document.configuracion.leyticketdespedida.value,0)==1){
			window.alert("<%=LitMsgConf2DesCarNoVal%>");
			return false;
		}
    }

    
    
    if (si_tiene_modulo_tiendas!=0){
	    if (isNaN(parent.pantalla.document.configuracion.validez.value) ) {
   			window.alert("<%=LitNoValValidez%>");
			return false;
	    }
	    if (isNaN(parent.pantalla.document.configuracion.anulacion.value) ) {
   			window.alert("<%=LitNoValAnulacion%>");
			return false;
	    }
    }
    if (si_tiene_modulo_tiendas!=0){
	    if (isNaN(parent.pantalla.document.configuracion.valorTicket.value.replace(",",".")) ) {
   			window.alert("<%=LitNoValValorTicket%>");
			return false;
	    }
    }
    
    si_tiene_modulo_Asesorias=parent.pantalla.document.configuracion.si_tiene_modulo_Asesorias.value;
    if (si_tiene_modulo_Asesorias!=0){
        if ((parent.pantalla.document.configuracion.chkAlertaFinContrato.checked || parent.pantalla.document.configuracion.chkAltaBaja.checked ) && parent.pantalla.document.configuracion.asesoriamail2.value=="" && parent.pantalla.document.configuracion.asesoriasms.value==""  ) {
   			alert("<%=LitMsgSinMailSMS %>");
			return false;
	    }
    }
    
    si_tiene_modulo_mantenimiento=parent.pantalla.document.configuracion.si_tiene_modulo_mantenimiento.value;
    if (si_tiene_modulo_mantenimiento!=0)
    {
        /*if (parent.pantalla.document.configuracion.senderemailincidences.value == "")
        {
   			alert("<%=LitMsgErrorSenderEmailIncidencesNotNull%>");
			return false;
	    }*/
	    if (parent.pantalla.document.configuracion.senderemailincidences.value != "")
        {
            var RegExPattern = /[\w-\.]{3,}@([\w-]{2,}\.)*([\w-]{2,}\.)[\w-]{2,4}/;  
            if (!parent.pantalla.document.configuracion.senderemailincidences.value.match(RegExPattern))
            {
   			    alert("<%=LitMsgErrorSenderEmailIncidencesNotValid%>");
			    return false;
			}
	    }
    }

    si_tiene_modulo_TGB = parent.pantalla.document.configuracion.si_tiene_modulo_TGB.value;
    if (si_tiene_modulo_TGB != 0 ) {
        bincode = parent.pantalla.document.configuracion.TGBBIN.value;
        origbincode = parent.pantalla.document.configuracion.tgbBinOrig.value;
        if (origbincode!=bincode && bincode != "") {
            if (bincode.length == 6 && IsNumeric(bincode)) {
                ret_validateBin = "";
                if (!enProceso && http) {
                    var timestamp = Number(new Date());
                    var url = "ValidateBin.asp?bincode=" + bincode;
                    document.opciones.vbin.value = "";
                    if (navigator.userAgent.indexOf("Firefox") > 0) {
                        http.open("GET", url, true);
                        http.onreadystatechange = handleHttpResponse;
                        enProceso = false;
                        http.send(null);
                        /*while (document.opciones.vbin.value == "") {
                        }
                        ret_validateBin = document.opciones.vbin.value;*/
                        return false;
                    } else {
                        http.open("GET", url, false);
                        http.onreadystatechange = handleHttpResponse;
                        enProceso = false;
                        http.send(null);
                    }
                    
                    
                    
                }
                
                
                if (ret_validateBin != "1") {
                    alert("<%=LitMsgBinUsed %>")
                    return false;
                }

            } else {
                alert("<%=LitMsgErrBin %>")
                return false;
            }
        }

        /**/
        if (!IsNumeric(parent.pantalla.document.configuracion.TGBCEP.value)) {
            window.alert("<%=LitMsgTGBCEPNoNum%>");
            return false;
        }
        if (!IsNumeric(parent.pantalla.document.configuracion.TGBCED.value)) {
            window.alert("<%=LitMsgTGBCEDNoNum%>");
            return false;
        }
        
    }
    
    parent.pantalla.document.configuracion.deccantidades.value=parseInt(parent.pantalla.document.configuracion.deccantidades.value);
    parent.pantalla.document.configuracion.decprecios.value=parseInt(parent.pantalla.document.configuracion.decprecios.value);

    return true;
}




//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado,ruta,rbt) {
    
	switch (mode) {
		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ruta=="1" || ValidarCampos()) {
                        act_cli=""					    
					    if (parent.pantalla.document.configuracion.chkMostrarAsesoria!=null){
                            if ( (parent.pantalla.document.configuracion.chkMostrarAsesoria.checked==true && parent.pantalla.document.configuracion.h_mostrarasesoria.value==0)
                               ||(parent.pantalla.document.configuracion.chkMostrarAsesoria.checked==false && parent.pantalla.document.configuracion.h_mostrarasesoria.value==1)
                               ){
                               if (parent.pantalla.document.configuracion.chkMostrarAsesoria.checked==true){
                                   if (window.confirm("<%=LitMsgActCli1%>")){
                                        act_cli="&act_cli=1"
                                   }
                               }else{
                                    if (window.confirm("<%=LitMsgActCli2%>")){
                                        act_cli="&act_cli=1"
                                   }
                               }
                            }
                        }
                        
						parent.pantalla.document.configuracion.action="configuracion.asp?mode=save"+act_cli;
						parent.pantalla.document.configuracion.submit();
						document.location="configuracion_bt.asp?mode=edit";
					}					
					break;
				case "cancel": //Cancelar edición
					parent.pantalla.document.configuracion.action="configuracion.asp?mode=edit";
					parent.pantalla.document.configuracion.submit();
					document.location="configuracion_bt.asp?mode=edit";
					break;
			}
			break;
		case "asistente":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ruta=="1" || ValidarCampos()) {
                        act_cli=""					    
					    if (parent.pantalla.document.configuracion.chkMostrarAsesoria!=null){
                            if ( (parent.pantalla.document.configuracion.chkMostrarAsesoria.checked==true && parent.pantalla.document.configuracion.h_mostrarasesoria.value==0)
                               ||(parent.pantalla.document.configuracion.chkMostrarAsesoria.checked==false && parent.pantalla.document.configuracion.h_mostrarasesoria.value==1)
                               ){
                               if (parent.pantalla.document.configuracion.chkMostrarAsesoria.checked==true){
                                   if (window.confirm("<%=LitMsgActCli1%>")) act_cli="&act_cli=1";
                               }
                               else{
                                    if (window.confirm("<%=LitMsgActCli2%>")) act_cli="&act_cli=1";
                               }
                            }
                        }
                        
						parent.pantalla.document.configuracion.action="configuracion.asp?mode=save"+act_cli;
						parent.pantalla.document.configuracion.submit();
						document.location="configuracion_bt.asp?mode=edit&viene=asistente";
					}				
					break;
				case "anterior": 			
					parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"2";						
					break;
			    case "cerrar": 
				    parent.location="../Applets/asistentePM.asp?mode=cancel";
				    break;
			    case "sig": 
				    //parent.location="../central.asp?pag1="+ruta+"&pag2="+rbt;				   
				    parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"1";		  				
				    break;
			}
			break;
	}
}
</script>
<!--<body leftmargin="<%=LitLeftPosBT%>" topmargin="<%=LitTopPosBT%>" bgcolor="<%=color_fondo_bt%>">-->
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<%'AMP 28/07/2010 : Añadimos parametro viene para asistente puesta en marcha.
 'viene=Request.QueryString("viene") 
 'tipo="1"
 'ruta=ObtenerNombreFichero(tipo) 
 if mode="save" then
    ok=request.QueryString("ok")&""
    %>
    <script>
        if ("<%=ok%>" != "1") {
            alert("<%=LitMsgBinUsed %>");
            document.location = "configuracion_bt.asp?mode=edit";
        } else {
            Accion("edit", "save", "1")
        }
    </script>
    <%
 else
%>
<form name="opciones" method="post">
    <input type="hidden" name="vbin" value="">
    <div id="PageFooter_ASP">
	<!--<table width=100% BORDER="0" CELLSPACING="1" CELLPADDING="1">-->
    <table  id="BUTTONS_CENTER_ASP">
		<tr>
		    <%if mode="edit" and viene<>"asistente" then%>
    		    <td CLASS="CELDABOT" onclick="javascript:Accion('edit','save','','');">
				    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
			    </td>
			    <td CLASS="CELDABOT" onclick="javascript:Accion('edit','cancel','','');">
				    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			    </td>
			<%end if

			'AMP 28/07/2010 : Incorporación botones asistente puesta en marcha
			if mode="edit" and viene="asistente" then%>
    		    <td CLASS="CELDABOT" onclick="javascript:Accion('asistente','save','');">
				    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
			    </td>
			    <td CLASS="CELDABOT" onclick="javascript:Accion('asistente','cerrar','');">
				    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			    </td>
			    <td CLASS=CELDABOT>
					<A CLASS=CELDAREF href="javascript:Accion('asistente','anterior','<%=ruta%>');"><IMG SRC="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></A>
				</td>
			    <td CLASS=CELDABOT>
					<A CLASS=CELDAREF href="javascript:Accion('asistente','sig','<%=ruta%>');"><IMG SRC="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></A>
				</td>
			<%end if%>
			<td CLASS=CELDABOT>
		</tr>
	</table>
    </div>
	<%ImprimirPie_bt%>
</form>
<%end if %>
</body>
</HTML>
