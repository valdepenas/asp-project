<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
     <% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<title></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../styles/Master.css.inc" -->

<!--#include file="empresa.inc" -->
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
function ValidarCampos() {

	ok=0;
	if (parent.pantalla.document.empresa.cif.value==""){
		window.alert("<%=LitMsgCIFNoNulo%>");
		ok=1;
	}

	if (ok==0 && parent.pantalla.document.empresa.nombre.value==""){
		window.alert("<%=LitMsgNombreNoNulo%>");
		ok=1;
	}

	if (ok==0){
		var indefinido;
		if (parent.pantalla.document.empresa.logotipo.value!="" && parent.pantalla.document.empresa.logotipo.value!=indefinido)
        {
            if (navigator.appName == "Microsoft Internet Explorer")
            {
			    var fso = new ActiveXObject("Scripting.FileSystemObject");
			    if (!fso.FileExists(parent.pantalla.document.empresa.logotipo.value)){
				    window.alert ("<%=LitFicheroNoExisteLogEmp1%>");
				    return false;
			    }
			    if (parent.pantalla.document.empresa.logotipo.value!=""){
				    if (fso.GetFile(parent.pantalla.document.empresa.logotipo.value).Size><%=LitTamnyFotoEmp1%>){
					    window.alert ("<%=LitErrTamnyFoto1 & LitTamnyFotoEmp1 & LitErrTamnyFoto4%>");
					    return false;
				    }
			    }
            }
            else
            {
                if (parent.pantalla.document.empresa.logotipo.files[0].size><%=cLng(LitTamnyFotoEmp1)%>){
			        window.alert ("<%=LitErrTamnyFoto1 & LitTamnyFotoEmp1 & LitErrTamnyFoto4%>");
			        return false;
		        }
            }
		}

		if (parent.pantalla.document.empresa.logotipo2.value!="" && parent.pantalla.document.empresa.logotipo2.value!=indefinido){
            if (navigator.appName == "Microsoft Internet Explorer")
            {
			    var fso = new ActiveXObject("Scripting.FileSystemObject");
			    if (!fso.FileExists(parent.pantalla.document.empresa.logotipo2.value)){
				    window.alert ("<%=LitFicheroNoExisteLogEmp2%>");
				    return false;
			    }
			    if (parent.pantalla.document.empresa.logotipo2.value!=""){
				    if (fso.GetFile(parent.pantalla.document.empresa.logotipo2.value).Size><%=LitTamnyFotoEmp2%>){
					    window.alert ("<%=LitErrTamnyFoto2 & LitTamnyFotoEmp2 & LitErrTamnyFoto4%>");
					    return false;
				    }
			    }
            }
            else
            {
                if (parent.pantalla.document.empresa.logotipo2.files[0].size><%=cLng(LitTamnyFotoEmp2)%>){
			        window.alert ("<%=LitErrTamnyFoto2 & LitTamnyFotoEmp2 & LitErrTamnyFoto4%>");
			        return false;
		        }
            }
		}

		if (parent.pantalla.document.empresa.logotipo3.value!="" && parent.pantalla.document.empresa.logotipo3.value!=indefinido)
        {
            if (navigator.appName == "Microsoft Internet Explorer")
            {
			    var fso = new ActiveXObject("Scripting.FileSystemObject");
			    if (!fso.FileExists(parent.pantalla.document.empresa.logotipo3.value)){
				    window.alert ("<%=LitFicheroNoExisteLogEmp3%>");
				    return false;
			    }
			    if (parent.pantalla.document.empresa.logotipo3.value!=""){
				    if (fso.GetFile(parent.pantalla.document.empresa.logotipo3.value).Size><%=LitTamnyFotoEmp3%>){
					    window.alert ("<%=LitErrTamnyFoto3 & LitTamnyFotoEmp3 & LitErrTamnyFoto4%>");
					    return false;
				    }
			    }
            }
            else
            {
                if (parent.pantalla.document.empresa.logotipo3.files[0].size><%=cLng(LitTamnyFotoEmp3)%>){
			        window.alert ("<%=LitErrTamnyFoto3 & LitTamnyFotoEmp3 & LitErrTamnyFoto4%>");
			        return false;
		        }
            }
		}
		return true;
	}
	else{
		return false;
	}
}

//validaremos el cif2 en caso de querer actualizar el cif de la empresa
function ValidarCif2() {
    if (parent.pantalla.document.empresa.cif2.value==""){
		window.alert("<%=LitMsgCIFNoNulo%>");
		return false;
	}
	
	
	
	return true; 
}

// comprobamos si se ha modificado el CIF y en caso de se haya hecho mostraremos una advertencia por pantalla
    function comprobarCambioCif() {
    cif1=trimCodEmpresa(parent.pantalla.document.empresa.cif.value);
    if (parent.pantalla.document.empresa.cif2.value != cif1  ){
	    
	    if(window.confirm("<%=LitMsgConfModCIF %>")){
	        return true;
	    }
	    else{
	        return false;
	    }
    }
	
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado,ruta) {           
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Guardar registro
				    limiteEmpresasCreadas=parseInt(parent.pantalla.document.empresa.limiteEmpresasCreadas.value);
					if (parent.pantalla.document.empresa.cantidad.value<limiteEmpresasCreadas){
						parent.pantalla.document.empresa.action="empresaFormulario.asp?mode=add";
						parent.pantalla.document.empresa.submit();
						document.location="empresa_bt.asp?mode=edit2";
					}
					else
					{
						mensaje_a_salir="<%=LitSoloPuedeInsGenerico1%>\n<%=LitSoloPuedeInsGenerico2%>\n<%=LitSoloPuedeInsGenerico3%>";
						window.alert(mensaje_a_salir);
					}
					break;
				case "cancel": //Cancelar edición				   
					parent.pantalla.document.empresa.action="empresa.asp?mode=browse";
					parent.pantalla.document.empresa.submit();
					document.location="empresa_bt.asp?mode=browse";
					break;
			}
			break;
		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos() && ValidarCif2()) {
					    //añadimos un parámetro para comprobar cuando el cif ha sido cambiado
					    //cambioCIF="&cambioCIF=";
					    comprobarCambioCif();
//					    if(){
//					        cambioCIF=cambioCIF+"1";
//					    }
//					    else{
//					          cambioCIF=cambioCIF+"0";
					  //  }
						parent.pantalla.document.empresa.action="empresa.asp?mode=save";  //+cambioCIF;
						parent.pantalla.document.empresa.submit();
						document.location="empresa_bt.asp?mode=browse";
					}
					break;
				case "first_save": //Guardar registro
					if (ValidarCampos()) {
						parent.pantalla.document.empresa.action="empresa.asp?mode=first_save";
						parent.pantalla.document.empresa.submit();
						document.location="empresa_bt.asp?mode=browse";
					}
					break;
				case "delete": //Guardar registro
					if (parent.pantalla.document.empresa.cantidad.value>1){
						if (window.confirm("<%=LitMsgEliminarEmpresaConfirm%>")==true) {
							parent.pantalla.document.empresa.action="empresa.asp?mode=delete";
							parent.pantalla.document.empresa.submit();
							document.location="empresa_bt.asp?mode=browse";
						}
					}
					else window.alert("<%=LitDebeDejarUna%>");
					break;
				case "cancel": //Cancelar edición				    
					parent.pantalla.document.empresa.action="empresa.asp?mode=browse";
					parent.pantalla.document.empresa.submit();
					document.location="empresa_bt.asp?mode=browse";
					break;
			}
			break;
	    case "asistente":
			switch (pulsado) {
			    case "add": //Guardar registro
				    limiteEmpresasCreadas=parseInt(parent.pantalla.document.empresa.limiteEmpresasCreadas.value);
					if (parent.pantalla.document.empresa.cantidad.value<limiteEmpresasCreadas){
						parent.pantalla.document.empresa.action="empresa.asp?mode=add";
						parent.pantalla.document.empresa.submit();
						document.location="empresa_bt.asp?mode=edit2&viene=asistente";
					}
					else
					{
						mensaje_a_salir="<%=LitSoloPuedeInsGenerico1%>\n<%=LitSoloPuedeInsGenerico2%>\n<%=LitSoloPuedeInsGenerico3%>";
						window.alert(mensaje_a_salir);
					}
					break;
				case "save": //Guardar registro
					if (ValidarCampos()) {
					    parent.pantalla.document.empresa.action="empresa.asp?mode=save&viene=asistente";
						parent.pantalla.document.empresa.submit();
						document.location="empresa_bt.asp?mode=browse";						
					}
					break;
				case "first_save": //Guardar registro
					if (ValidarCampos()) {
					    parent.pantalla.document.empresa.action="empresa.asp?mode=first_save&viene=asistente";
						parent.pantalla.document.empresa.submit();
						document.location="empresa_bt.asp?mode=browse";									
					}					
					break;
				case "cancel": //Cancelar edición				    
					parent.pantalla.document.empresa.action="empresa.asp?mode=browse&viene=asistente";
					parent.pantalla.document.empresa.submit();
					document.location="empresa_bt.asp?mode=browse&viene=asistente";
					break;	
			    case "sig": //Saltar a siguiente pantalla	         
			        parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"1";			     
                   	//parent.location="../central.asp?pag1="+ruta+"&pag2="+rbt;			             				       	
					break;
			     case "cerrar": //Cerrar asistente.				    
					parent.location="../Applets/asistentePM.asp?mode=cancel";
					break;					
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<%'AMP 28/07/2010 : Añadimos parametro viene para asistente de puesta en marcha.
 viene=limpiaCadena(Request.QueryString("viene"))
 tipo="1"
 ruta=ObtenerNombreFichero(tipo) 
 %>
<form name="opciones" method="post">
    <input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
    <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
		    <% if mode="browse" then
		        if viene="asistente" then %>
    		        <td id="idadd" class="CELDABOT" onclick="javascript:Accion('asistente','add','','');">
				        <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
			        </td>
			        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cerrar','','');">
				        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			        </td>		         
				     <td class="CELDABOT">				   		    
					     <a class="CELDAREF" href="javascript:Accion('asistente','sig','<%=enc.EncodeForJavascript(ruta)%>');"><img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a>					
				     </td>
			    <%else %>
    		        <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add','','');">
				        <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
			        </td>
			    <%end if
			end if
			if mode="edit" then%>
			    <%if viene="asistente" then %>
    		        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('asistente','save');">
				        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
			        </td>
			        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel','','');">
				        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			        </td>	
				<%else%>
    		        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
				        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
			        </td>
			        <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('edit','delete','','');">
				        <%PintarBotonBTRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
			        </td>
			        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel','','');">
				        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			        </td>
			    <%end if %>				
				
			<%end if
			if mode="edit2" then%>
			     <%if viene="asistente" then %>
    		        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('asistente','first_save');">
				        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
			        </td>
			        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel','','');">
				        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			        </td>
				<%else%>
    		        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','first_save');">
				        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
			        </td>
			        <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('edit','delete','','');">
				        <%PintarBotonBTRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
			        </td>
			        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel','','');">
				        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			        </td>
			    <%end if %>				
				
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