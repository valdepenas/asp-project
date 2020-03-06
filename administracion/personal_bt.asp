<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../calculos.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="personal.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">

    window.addEventListener("load", function hideFilterMaster() {
        var classNameLeft = parent.parent.document.getElementById('left').className;
        if (classNameLeft == "search-panel open") {
            document.getElementById("FILTERS_MASTER_ASP").setAttribute("style", "display:none")
        }
    });

function Buscar()
{
	SearchPage("personal_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&viene=<%=enc.EncodeForJavascript(limpiaCadena(Request.QueryString("viene")))%>",1);
    document.opciones.texto.value = "";
}

//Validación de campos numéricos y fechas.
function ValidarCampos(mode)
{
	si_tiene_modulo_comercial=parent.pantalla.document.personal.h_si_tiene_modulo_comercial.value;
	si_tiene_modulo_produccion=parent.pantalla.document.personal.h_si_tiene_modulo_produccion.value;
	si_tiene_modulo_mantenimiento=parent.pantalla.document.personal.h_si_tiene_modulo_mantenimiento.value;

    /* Comprobacion de nulos */
    if (parent.pantalla.document.personal.dni.value=="")
    {
        window.alert ("<%=LitMsgDniNoNulo%>");
        return false;
    }
    else
    {
        if (parent.pantalla.document.personal.nombre.value=="")
        {
            window.alert ("<%=LitMsgNombreNoNulo%>");
            return false;
        }
	    else
	    {
	        if (parent.pantalla.document.personal.domicilio.value=="")
	        {
                window.alert ("<%=LitMsgDireccionNoNulo%>");
                return false;
            }
	    }
    }
	/*Comprobación de fechas */
	if (parent.pantalla.document.personal.antiguedad.value!="")
	{
		if (!checkdate(parent.pantalla.document.personal.antiguedad))
		{
			window.alert("<%=LitMsgPerAntNoFecha%>");
			return false;
		}
	}

	if (parent.pantalla.document.personal.fbaja.value!="")
	{
		if (!checkdate(parent.pantalla.document.personal.fbaja))
		{
			window.alert("<%=LitMsgFechaBajaFecha%>");
			return false;
		}

        fechaActual = new Date();
		if (DiferenciaTiempo(fechaActual.getDate()+"/"+(parseInt(fechaActual.getMonth())+1)+"/"+fechaActual.getFullYear(),parent.pantalla.document.personal.fbaja.value,"dias")<0)
		{
			window.alert("<%=LitMsgPerFBajaAntig%>");
			return false;
		}
	}

	/*Comprobación de horas de la Jornada*/
	if (parent.pantalla.document.personal.horaIniMa.value!="" || parent.pantalla.document.personal.horaFinMa.value!=""){
		if (parent.pantalla.document.personal.horaIniMa.value=="" || parent.pantalla.document.personal.horaFinMa.value==""){
			window.alert("<%=LitMsgJornadaMIniyFin%>");
			return false;
		}
		if ((parent.pantalla.document.personal.horaIniMa.value.length!=5) || !checkhora(parent.pantalla.document.personal.horaIniMa)) {
			window.alert("<%=LitMsgHoraIniMMal%>");
			return false;
		}
		if ((parent.pantalla.document.personal.horaFinMa.value.length!=5) || !checkhora(parent.pantalla.document.personal.horaFinMa)) {
			window.alert("<%=LitMsgHoraFinMMal%>");
			return false;
		}
		//Una mayor que otra
		if	(convertir_fechatexto_fechamilisegundos("1/1/1900 "+parent.pantalla.document.personal.horaIniMa.value)>=convertir_fechatexto_fechamilisegundos("1/1/1900 "+parent.pantalla.document.personal.horaFinMa.value)){
			window.alert("<%=LitMsgMFinPosteriorIni%>");
			return false;
		}
	}
	if (parent.pantalla.document.personal.horaIniTa.value!="" || parent.pantalla.document.personal.horaFinTa.value!=""){
		if (parent.pantalla.document.personal.horaIniTa.value=="" || parent.pantalla.document.personal.horaFinTa.value==""){
			window.alert("<%=LitMsgJornadaTIniyFin%>");
			return false;
		}
		if ((parent.pantalla.document.personal.horaIniTa.value.length!=5) || !checkhora(parent.pantalla.document.personal.horaIniTa)) {
			window.alert("<%=LitMsgHoraIniTMal%>");
			return false;
		}
		if ((parent.pantalla.document.personal.horaFinTa.value.length!=5) || !checkhora(parent.pantalla.document.personal.horaFinTa)) {
			window.alert("<%=LitMsgHoraFinTMal%>");
			return false;
		}
		//Una mayor que otra
		if	(convertir_fechatexto_fechamilisegundos("1/1/1900 "+parent.pantalla.document.personal.horaIniTa.value)>=convertir_fechatexto_fechamilisegundos("1/1/1900 "+parent.pantalla.document.personal.horaFinTa.value)){
			window.alert("<%=LitMsgTFinPosteriorIni%>");
			return false;
		}
		//Si existe mañana y tarde, mañana debe ser menor que tarde
		if (parent.pantalla.document.personal.horaIniMa.value!=""){
			if	(convertir_fechatexto_fechamilisegundos("1/1/1900 "+parent.pantalla.document.personal.horaFinMa.value)>=convertir_fechatexto_fechamilisegundos("1/1/1900 "+parent.pantalla.document.personal.horaIniTa.value)){
				window.alert("<%=LitMsgCruceJornadas%>");
				return false;
			}
		}
	}

    /*Comprobacion de valores numéricos */
   if (isNaN(parent.pantalla.document.personal.jornada.value.replace(",","."))) {
      window.alert("<%=LitMsgJornadaNumerico%>");
	  return false;
   }
   else{
      if (isNaN(parent.pantalla.document.personal.sueldo.value.replace(",","."))) {
	     window.alert("<%=LitMsgSueldoNumerico%>");
	     return false;
	  }
	  else{
	     if (isNaN(parent.pantalla.document.personal.phextra.value.replace(",","."))) {
		    window.alert("<%=LitMsgHoraExtraNumerico%>");
			return false;
		 }
	  }


   }

   if (isNaN(parent.pantalla.document.personal.maxamount.value.replace(',', '.'))) {
       alert("<%=LitMaxAmountNoNumber%>");
       return false;
   }

    if (mode != "add") {
        if (parent.pantalla.document.personal.existe_tecnico.value == "1") {
            if (si_tiene_modulo_mantenimiento != 0) {
                if (isNaN(parent.pantalla.document.personal.tphlaboral.value.replace(",", "."))) {
                    window.alert("<%=LIT_MSGPHLABORALNUMERICO%>");
                    return false;
                }
            }
        }
    }

   /*Comprobamos valores del comercial*/

   if (mode!="add")
   {
   	    if (parent.pantalla.document.personal.existe_comercial.value=="1")
   	    {
		    if (si_tiene_modulo_comercial!="0")
		    {
			    while (parent.pantalla.document.personal.cventas.value.search(" ")!=-1)
				    parent.pantalla.document.personal.cventas.value=parent.pantalla.document.personal.cventas.value.replace(" ","");
			    while (parent.pantalla.document.personal.mganancia.value.search(" ")!=-1)
				    parent.pantalla.document.personal.mganancia.value=parent.pantalla.document.personal.mganancia.value.replace(" ","");
			    while (parent.pantalla.document.personal.per_ob.value.search(" ")!=-1)
				    parent.pantalla.document.personal.per_ob.value=parent.pantalla.document.personal.per_ob.value.replace(" ","");
			    while (parent.pantalla.document.personal.objetivo.value.search(" ")!=-1)
				    parent.pantalla.document.personal.objetivo.value=parent.pantalla.document.personal.objetivo.value.replace(" ","");
			    while (parent.pantalla.document.personal.cbase.value.search(" ")!=-1)
				    parent.pantalla.document.personal.cbase.value=parent.pantalla.document.personal.cbase.value.replace(" ","");
			    while (parent.pantalla.document.personal.cconcepto.value.search(" ")!=-1)
				    parent.pantalla.document.personal.cconcepto.value=parent.pantalla.document.personal.cconcepto.value.replace(" ","");
			    while (parent.pantalla.document.personal.pena.value.search(" ")!=-1)
				    parent.pantalla.document.personal.pena.value=parent.pantalla.document.personal.pena.value.replace(" ","");

			    if (isNaN(parent.pantalla.document.personal.cventas.value.replace(",","."))) {
				    window.alert("<%=LitMsgComisionNumerico%>");
				    return false;
			    }
			    else
			    {
				    if (isNaN(parent.pantalla.document.personal.mganancia.value.replace(",",".")))
				    {
					    window.alert("<%=LitMsgMargenNumerico%>");
					    return false;
				    }
				    else{
					    if (isNaN(parent.pantalla.document.personal.per_ob.value.replace(",",".")))
					    {
						    window.alert("<%=LitMsgPeriodicidadNumerico%>");
						    return false;
					    }
					    else
					    {
						    if (isNaN(parent.pantalla.document.personal.objetivo.value.replace(",","."))) 
						    {
							    window.alert("<%=LitMsgObjetivodNumerico%>");
							    return false;
						    }
						    else
						    {
							    if (isNaN(parent.pantalla.document.personal.cbase.value.replace(",","."))) 
							    {
								    if (si_tiene_modulo_comercial!=0) window.alert("<%=LitMsgCBaseNumericoModCom%>");
								    else window.alert("<%=LitMsgCBaseNumerico%>");
								    return false;
							    }
							    else{
								    if (isNaN(parent.pantalla.document.personal.cconcepto.value.replace(",",".")))
								    {
									    if (si_tiene_modulo_comercial!=0) window.alert("<%=LitMsgCConceptoNumericoModCom%>");
									    else window.alert("<%=LitMsgCConceptoNumerico%>");
									    return false;
								    }
								    else
								    {
									    if (isNaN(parent.pantalla.document.personal.pena.value.replace(",",".")))
									    {
										    if (si_tiene_modulo_comercial!=0) window.alert("<%=LitMsgPenaNumericoModCom%>");
										    else window.alert("<%=LitMsgPenaNumerico%>");
										    return false;
									    }
								    }
							    }
						    }
					    }
				    }
			    }
		    }
        }

        /*Comprobamos valores del tecnico*/

        if (parent.pantalla.document.personal.existe_tecnico.value=="1"){
	        if (si_tiene_modulo_mantenimiento!=0){
		        while (parent.pantalla.document.personal.tcomision.value.search(" ")!=-1)
			        parent.pantalla.document.personal.tcomision.value=parent.pantalla.document.personal.tcomision.value.replace(" ","");
			    while (parent.pantalla.document.personal.tphextralab.value.search(" ")!=-1)
				    parent.pantalla.document.personal.tphextralab.value=parent.pantalla.document.personal.tphextralab.value.replace(" ","");
			    while (parent.pantalla.document.personal.tphextrafes.value.search(" ")!=-1)
				    parent.pantalla.document.personal.tphextrafes.value=parent.pantalla.document.personal.tphextrafes.value.replace(" ","");
			    while (parent.pantalla.document.personal.tincentivo1.value.search(" ")!=-1)
				    parent.pantalla.document.personal.tincentivo1.value=parent.pantalla.document.personal.tincentivo1.value.replace(" ","");
			    while (parent.pantalla.document.personal.tincentivo2.value.search(" ")!=-1)
				    parent.pantalla.document.personal.tincentivo2.value=parent.pantalla.document.personal.tincentivo2.value.replace(" ","");

		        if (isNaN(parent.pantalla.document.personal.tcomision.value.replace(",","."))) {
		  		    window.alert("<%=LitMsgComisionTecnicoNumerico%>");
				    return false;
			    }
	            if (isNaN(parent.pantalla.document.personal.tphextralab.value.replace(",","."))) {
                    window.alert("<%=LitMsgHextraLabNumerico%>");
	      	        return false;
		        }
		        if (isNaN(parent.pantalla.document.personal.tphextrafes.value.replace(",","."))) {
			        window.alert("<%=LitMsgHextraFesNumerico%>");
			        return false;
		        }
		        if (isNaN(parent.pantalla.document.personal.tincentivo1.value.replace(",","."))) {
			        window.alert("<%=LitMsgIncentivo1Numerico%>");
			        return false;
		        }
		        if (isNaN(parent.pantalla.document.personal.tincentivo2.value.replace(",","."))) {
			        window.alert("<%=LitMsgIncentivo2Numerico%>");
			        return false;
		        }
	        }
        }

        /*Comprobamos valores del operario*/

        if (parent.pantalla.document.personal.existe_operario.value=="1")
        {
	        if (si_tiene_modulo_produccion!=0)
		    {
			    while (parent.pantalla.document.personal.ocoste_hora.value.search(" ")!=-1)
				    parent.pantalla.document.personal.ocoste_hora.value=parent.pantalla.document.personal.ocoste_hora.value.replace(" ","");
    	
		        if (isNaN(parent.pantalla.document.personal.ocoste_hora.value.replace(",",".")))
		        {
		  		    window.alert("<%=LitMsgCosteHoraNumerico%>");
				    return false;
	      	    }
		    }
        }

        /*comprobacion de la mensajeria SMS */
        if (parent.pantalla.document.personal.telefono2.value.indexOf(" ")!=-1 ||
		    parent.pantalla.document.personal.telefono2.value.indexOf(".")!=-1 ||
		    parent.pantalla.document.personal.telefono2.value.indexOf("-")!=-1 ||
		    parent.pantalla.document.personal.telefono2.value.indexOf("(")!=-1 ||
		    parent.pantalla.document.personal.telefono2.value.indexOf(")")!=-1 ||
		    parent.pantalla.document.personal.telefono2.value.indexOf("/")!=-1 ||
		    parent.pantalla.document.personal.telefono2.value.indexOf("\\")!=-1 )
	    {
		    alert("<%=LitMsgCaracteresIncorrectosMovil%>");
		    return false;
	    }
	    if (parent.pantalla.document.personal.si_tiene_paginaSMS.value=="1"){
		    if (parent.pantalla.document.personal.mensajeria_sms.checked && parent.pantalla.document.personal.mensajeria_smsbd.value=="0") {
			    alert("<%=LitNoMensajeriaSMS%>");
		 	    return false;
		     }
	    }
	}//del mode!=add
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
<%viene = limpiaCadena(Request.QueryString("viene")) %>
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Nuevo registro
					parent.pantalla.document.personal.action="personal.asp?mode=" + pulsado;
					parent.pantalla.document.personal.submit();
					document.location="personal_bt.asp?mode=" + pulsado;
					break;

				case "edit": //Editar registro
					parent.pantalla.document.personal.action="personal.asp?dni=" + parent.pantalla.document.personal.hdni.value +
					"&mode=" + pulsado+ "&domicilio="+parent.pantalla.document.personal.hdomicilio.value;
					parent.pantalla.document.personal.submit();
					document.location="personal_bt.asp?mode=" + pulsado+"&viene=<%=enc.EncodeForJavascript(viene)%>";
					break;

				case "delete": //Eliminar registro
					if (window.confirm("<%=LitMsgEliminarFichaConfirm%>")==true) {
						parent.pantalla.document.personal.action="personal.asp?mode=" + pulsado + "&dni=" + parent.pantalla.document.personal.hdni.value;
						parent.pantalla.document.personal.submit();
						//ASP 27/12/2011                            
                            reloadPanelGlobal("<%=enc.EncodeForJavascript(viene)%>");
                        //FIN ASP 27/12/2011
						document.location="personal_bt.asp?mode=browse";
					}
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
				    if (parent.pantalla.document.personal.fbaja.value!="") {
      			        if (window.confirm("<%=LitMsgEliminarPersonalConfirm%>")==true) {
					        if (ValidarCampos(mode)) {
						        parent.pantalla.document.personal.action="personal.asp?mode=save"+ "&dni=" + parent.pantalla.document.personal.dni.value +
						        "&domicilio="+parent.pantalla.document.personal.domicilio.value;
						        parent.pantalla.document.personal.submit();
                                //ASP 27/12/2011
						        reloadPanelGlobal("<%=enc.EncodeForJavascript(viene)%>");
                                //FIN ASP 27/12/2011
						        document.location="personal_bt.asp?mode=browse"+"&viene=<%=enc.EncodeForJavascript(viene)%>";
					        }
				        }
				    }
				    else {
					    if (ValidarCampos(mode)) {
						    parent.pantalla.document.personal.action="personal.asp?mode=save"+ "&dni=" + parent.pantalla.document.personal.dni.value +
						    "&domicilio="+parent.pantalla.document.personal.domicilio.value;
						    parent.pantalla.document.personal.submit();
                                //ASP 27/12/2011
						    reloadPanelGlobal("<%=enc.EncodeForJavascript(viene)%>");
                                //FIN ASP 27/12/2011
						    document.location="personal_bt.asp?mode=browse";
					    }
				    }
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.personal.action="personal.asp?dni=" + parent.pantalla.document.personal.hdni.value +
					"&mode=browse";
					parent.pantalla.document.personal.submit();
					document.location="personal_bt.asp?mode=browse";
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode)) {
						parent.pantalla.document.personal.action="personal.asp?mode=first_save&dni="+parent.pantalla.document.personal.dni.value+"&domicilio="+parent.pantalla.document.personal.domicilio.value;
						parent.pantalla.document.personal.submit();
						document.location="personal_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.personal.action="personal.asp?mode=add";
					parent.pantalla.document.personal.submit();
					document.location="personal_bt.asp?mode=add";
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
	     case "asistente":
	        switch (pulsado) {				
			    case "add": //Nuevo registro			       
					parent.pantalla.document.personal.action="personal.asp?mode=add&viene=asistente";
					parent.pantalla.document.personal.submit();
					document.location="personal_bt.asp?mode=add&viene=asistente";
					break;				
				case "cancel":
				    parent.pantalla.document.personal.action="personal.asp?mode=search&viene=asistente";
					parent.pantalla.document.personal.submit();
					document.location="personal_bt.asp?mode=search&viene=asistente";
					break;
				case "first_save":
				    if (ValidarCampos("add")) {					   
					  	parent.pantalla.document.personal.action="personal.asp?mode=first_save&dni="+parent.pantalla.document.personal.dni.value+"&domicilio="+parent.pantalla.document.personal.domicilio.value;
						parent.pantalla.document.personal.submit();						
						
						parent.pantalla.document.personal.action="personal.asp?mode=search&viene=asistente";
						parent.pantalla.document.personal.submit();
						document.location="personal_bt.asp?mode=search&viene=asistente";
					}
					break;	
				case "save": //Guardar registro
					if (ValidarCampos()) {					   
					    parent.pantalla.document.personal.action="personal.asp?mode=save"+ "&dni=" + parent.pantalla.document.personal.dni.value +
				        "&domicilio="+parent.pantalla.document.personal.domicilio.value;
						parent.pantalla.document.personal.submit();						
						
						parent.pantalla.document.personal.action="personal.asp?mode=search&viene=asistente";
						parent.pantalla.document.personal.submit();
						document.location="personal_bt.asp?mode=search&viene=asistente";
					}
					break;
				case "edit": //Editar registro
					parent.pantalla.document.personal.action="personal.asp?dni=" + parent.pantalla.document.personal.hdni.value +
					"&mode=" + pulsado+ "&domicilio="+parent.pantalla.document.personal.hdomicilio.value;
					parent.pantalla.document.personal.submit();
					document.location="personal_bt.asp?mode=edit&viene=asistente";				
					break;				
			}
			break;
	}
}

// ASP 27/11/2011
 function reloadPanelGlobal(viene)
 {
    if(viene == "GlobalAgenda")
    {
        parent.window.opener.__doPostBack("reload",""); 
    }
 }
 
 //FIN ASP 27/11/2011
//****************************************************************************************************************
function comprobar_enter(){
    //var keycode = ev.keyCode;
	//si se ha pulsado la tecla enter
	//if (keycode==13){
		//document.opciones.criterio.focus();
		Buscar();
	//}
}
function MoverPagPM(ruta,rbt){
    if (rbt=="1")
  	    parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"1";        
  	else
        parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"2";	     
}

function Cerrar(){
    parent.location="../Applets/asistentePM.asp?mode=cancel";
}

//  End -->
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")
viene=limpiaCadena(Request.QueryString("viene"))
tipo="1"%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />	
	   <%if viene<>"asistente" then %> 
        <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_left_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
			    <%if mode="browse" then%>
                   <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBTLeft LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				    </td>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeftRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				<%elseif mode="search" then%>
                   <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				<%elseif mode="edit" then%>
			        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				<%elseif mode="add" then%>
			        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				<%end if%>
		        </tr>
	        </table>
            </div>
    
            <div id="FILTERS_MASTER_ASP">
				<!--<td class=CELDABOT><%=LitBuscar & ": "%>-->
					<select class="IN_S" name="campos">
						<option value="nombre"><%=Litnombre%></option>
						<option value="dni"><%=Litdni%></option>
						<option value="personal.codigo"><%=LitCodigoOp%></option>
						<option value="domicilio"><%=LitDomicilio%></option>
						<option value="cp"><%=LitCp%></option>
						<option value="provincia"><%=LitProvincia%></option>
						<option value="pais"><%=LitPais%></option>
						<option value="telefono"><%=LitTel1%></option>
						<option value="telefono2"><%=LitTel2%></option>
						<option value="fax"><%=LitFax%></option>
						<option value="observaciones"><%=LitObservaciones%></option>
					</select>
				<!--</td>
				<td class=CELDABOT>-->
					<select class="IN_S" name="criterio">
						<option value="contiene"><%=LitContiene%></option>
						<!--<option value="empieza"><%=LitComienza%></option>-->
						<option value="termina"><%=LitTermina%></option>
						<option value="igual"><%=LitIgual%></option>
					</select>
				<!--</td>
				<td class=CELDABOT>-->
                    <input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<!--</td>
				<td class=CELDABOT>-->
                    <a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
				<!--</td>
		        </tr>
	        </table>-->
            </div>
            </div>
		<%else%>
        <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
			    <%if mode="browse" then%>
                    <td id="idadd" class="CELDABOT" onclick="javascript:Accion('asistente','add');">
					    <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('asistente','edit');">
					    <%PintarBotonBT LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				    </td>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('asistente','cancel');">
					    <%PintarBotonBTRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>			   	      					
				<%elseif mode="search" then%>
                    <td id="idadd" class="CELDABOT" onclick="javascript:Accion('asistente','add');">
					    <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Cerrar();">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
                    <td id="idanterior" class="CELDABOT" onclick="javascript:MoverPagPM('<%=enc.EncodeForJavascript(ruta)%>','2');">
					    <%PintarBotonBT LITBOTANTERIOR,ImgAnterior,ParamImgAnterior,LITBOTANTERIORTITLE%>
				    </td>
                    <td id="idsiguiente" class="CELDABOT" onclick="javascript:MoverPagPM('<%=enc.EncodeForJavascript(ruta)%>','1');">
					    <%PintarBotonBT LITBOTSIGUIENTE,ImgSiguiente,ParamImgSiguiente,LITBOTSIGUIENTETITLE%>
				    </td>     	
				<%elseif mode="edit" then%>
                	<td id="idsave" class="CELDABOT" onclick="javascript:Accion('asistente','save');">
					    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				<%elseif mode="add" then%>
                    <td id="idsave" class="CELDABOT" onclick="javascript:Accion('asistente','first_save');">
					    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				<%end if%>			
		        </tr>
	        </table>
            </div>
            </div>
		<%end if
        %>
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