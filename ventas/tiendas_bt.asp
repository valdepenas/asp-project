<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head> 
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<!--#include file="tiendas.inc" -->

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
mode=Request.QueryString("mode")
    %>  

</head>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">

    window.addEventListener("load", function hideFilterMaster() {
        var classNameLeft = parent.parent.document.getElementById('left').className;
        if (classNameLeft == "search-panel open") {
            document.getElementById("FILTERS_MASTER_ASP").setAttribute("style", "display:none")
        }
    });
// Toni climent 14-01-2009 Se añade a la validacion los rangos y valores permitidos
// Los valores Máximo y Mínimo vienen definidos dentro de las cosntantes MINVALSAL y MAXVALSAL
var MINVALSAL =   0;
var MAXVALSAL = 60;
var ev = null;

si_tiene_modulo_ebesa=0;
si_tiene_modulo_OrCU = 0;
si_tiene_modulo_ModFidelizacionPremium = 0;
si_tiene_modulo_Agroclub = 0;
tieneTPV="";
es_hostelera=0;

function Buscar()
{
	SearchPage("tiendas_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value, 1);
	document.opciones.texto.value = "";
}

//Validación de campos numéricos y fechas.
function ValidarCampos()
{
    si_tiene_modulo_ebesa = parent.pantalla.document.tiendas.si_tiene_modulo_ebesa.value;
    si_tiene_modulo_OrCU = parent.pantalla.document.tiendas.si_tiene_modulo_OrCU.value;
    si_tiene_modulo_ModFidelizacionPremium = parent.pantalla.document.tiendas.si_tiene_modulo_ModFidelizacionPremium.value;
    si_tiene_modulo_Agroclub = parent.pantalla.document.tiendas.si_tiene_modulo_Agroclub.value;
    tieneTPV = parent.pantalla.document.tiendas.h_tieneTPV.value;
    es_hostelera = parent.pantalla.document.tiendas.h_es_hostelera.value;
	if (parent.pantalla.document.tiendas.codigo.value=="")
	{
		window.alert("<%=LitMsgCodigoNoNulo%>");
		return false;
	}
	else
	{
		if (comp_car_ext(parent.pantalla.document.tiendas.codigo.value,1)==1){
		    window.alert("<%=LitMsgTienDesCarNoVal%>");
		    return false;
		}
        if (parent.pantalla.document.tiendas.descripcion.value=="") {
            window.alert("<%=LitMsgTiendaNoNulo%>");
            return false;
        }

        if (parent.pantalla.document.tiendas.descripcion.value.length>50) {
            window.alert("<%=LitMsgDescrTiendaLargo%>");
            return false;
        }

        if (parent.pantalla.document.tiendas.domicilio.value=="") {
            window.alert("<%=LitMsgDireccionNoNulo%>");
            return false;
        }

        if (parent.pantalla.document.tiendas.domicilio.value.length>100) {
            window.alert("<%=LitMsgDescrDomicLargo%>");
            return false;
        }
		else{
			if (parent.pantalla.document.tiendas.almacen.value=="") {
                window.alert("<%=LitMsgAlmacenNoNulo%>");
                return false;
            }
	   		else
	   		{
				if (comp_car_ext(parent.pantalla.document.tiendas.almacen.value,0)==1)
				{
			        window.alert("<%=LitMsgAlmaDesCarNoVal%>");
			        return false;
				}
			}
        }
	}

    if (si_tiene_modulo_ebesa!=0){
	    if (isNaN(parent.pantalla.document.tiendas.pc1.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pri1.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pc2.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pri2.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pc3.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pri3.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pc4.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pri4.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pc5.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.pri5.value.replace(',','.'))) {
			    window.alert("<%=LitMsgNoNum%>");
			    return false;
	    }
    }

    if (es_hostelera!=0 && "<%=mode%>" != "add"){
        //Toni Climent 15-01-2009 Comprobacion si los valores son numericos y si estan dentro del valor
        if (isNaN(parent.pantalla.document.tiendas.sal1.value) || parseInt(parent.pantalla.document.tiendas.sal1.value).toString() != parent.pantalla.document.tiendas.sal1.value) {
		    window.alert("<%=LitMsgNoNum1%><%=LitSal1%><%=LitMsgNoNum4%>");
		    return false;
	    }
	    else{
	        if(parseInt(parent.pantalla.document.tiendas.sal1.value) <  MINVALSAL|| parseInt(parent.pantalla.document.tiendas.sal1.value) >  MAXVALSAL){
	            window.alert("<%=LitMsgOutRange1%><%=LitSal1%><%=LitMsgOutRange2%>" + MINVALSAL.toString() + "<%=LitMsgOutRange3%>" + MAXVALSAL.toString());
			    return false;
		    }
	    }
	    
	    if (isNaN(parent.pantalla.document.tiendas.sal2.value) || parseInt(parent.pantalla.document.tiendas.sal2.value).toString() != parent.pantalla.document.tiendas.sal2.value) {
		    window.alert("<%=LitMsgNoNum1%><%=LitSal2%><%=LitMsgNoNum4%>");
		    return false;
	    }
	    else{
	        if(parseInt(parent.pantalla.document.tiendas.sal2.value) <  MINVALSAL|| parseInt(parent.pantalla.document.tiendas.sal2.value) >  MAXVALSAL){
	            window.alert("<%=LitMsgOutRange1%><%=LitSal2%><%=LitMsgOutRange2%>" + MINVALSAL.toString() + "<%=LitMsgOutRange3%>" + MAXVALSAL.toString());
			    return false;
		    }
	    }
	    if (isNaN(parent.pantalla.document.tiendas.sal3.value) || parseInt(parent.pantalla.document.tiendas.sal3.value).toString() != parent.pantalla.document.tiendas.sal3.value) {
		    window.alert("<%=LitMsgNoNum1%><%=LitSal3%><%=LitMsgNoNum4%>");
		    return false;
	    }
	    else{
	        if(parseInt(parent.pantalla.document.tiendas.sal3.value) <  MINVALSAL|| parseInt(parent.pantalla.document.tiendas.sal3.value) >  MAXVALSAL){
	            window.alert("<%=LitMsgOutRange1%><%=LitSal3%><%=LitMsgOutRange2%>" + MINVALSAL.toString() + "<%=LitMsgOutRange3%>" + MAXVALSAL.toString());
			    return false;
		    }
	    }
	    if (isNaN(parent.pantalla.document.tiendas.sal4.value) || parseInt(parent.pantalla.document.tiendas.sal4.value).toString() != parent.pantalla.document.tiendas.sal4.value) {
		    window.alert("<%=LitMsgNoNum1%><%=LitSal4%><%=LitMsgNoNum4%>");
		    return false;
	    }
	    else{
	        if(parseInt(parent.pantalla.document.tiendas.sal4.value) <  MINVALSAL|| parseInt(parent.pantalla.document.tiendas.sal4.value) >  MAXVALSAL){
	            window.alert("<%=LitMsgOutRange1%><%=LitSal4%><%=LitMsgOutRange2%>" + MINVALSAL.toString() + "<%=LitMsgOutRange3%>" + MAXVALSAL.toString());
			    return false;
		    }
	    }
	    if (isNaN(parent.pantalla.document.tiendas.sal5.value) || parseInt(parent.pantalla.document.tiendas.sal5.value).toString() != parent.pantalla.document.tiendas.sal5.value) {
		    window.alert("<%=LitMsgNoNum1%><%=LitSal5%><%=LitMsgNoNum4%>");
		    return false;
	    }
	    else{
	        if(parseInt(parent.pantalla.document.tiendas.sal5.value) <  MINVALSAL|| parseInt(parent.pantalla.document.tiendas.sal5.value) >  MAXVALSAL){
	            window.alert("<%=LitMsgOutRange1%><%=LitSal5%><%=LitMsgOutRange2%>" + MINVALSAL.toString() + "<%=LitMsgOutRange3%>" + MAXVALSAL.toString());
			    return false;
		    }
	    }
    }

    if (si_tiene_modulo_ebesa!=0){
	    if (isNaN(parent.pantalla.document.tiendas.dtol.value.replace(',','.'))) {
		    window.alert("<%=LitMsgNoNum1%><%=LitDtol%><%=LitMsgNoNum2%>");
		    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.dtodia.value.replace(',','.'))) {
		    window.alert("<%=LitMsgNoNum1%><%=LitDtoDia%><%=LitMsgNoNum2%>");
		    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.dtoregalo.value.replace(',','.'))) {
		    window.alert("<%=LitMsgNoNum1%><%=LitDtoRegalo%><%=LitMsgNoNum2%>");
		    return false;
	    }
	    if (isNaN(parent.pantalla.document.tiendas.dtoencargado.value.replace(',','.'))) {
		    window.alert("<%=LitMsgNoNum1%><%=LitDtoEncargado%><%=LitMsgNoNum2%>");
		    return false;
	    }

	    if (!checkdate(parent.pantalla.document.tiendas.dtodiadesde)  )
	    {
		    window.alert("<%=LitMsgFDesde%>");
		    return false;
	    }
	    else
	    {
		    if (!checkhora(parent.pantalla.document.tiendas.dtohoradesde) && parent.pantalla.document.tiendas.dtohoradesde.value!="" ){
			    window.alert("<%=LitMsgHDesde%>");
			    return false;
		    }
	    }

	    if (!checkdate(parent.pantalla.document.tiendas.dtodiahasta)  )
	    {
		    window.alert("<%=LitMsgFHasta%>");
		    return false;
	    }
	    else
	    {
		    if (!checkhora(parent.pantalla.document.tiendas.dtohorahasta) && parent.pantalla.document.tiendas.dtohorahasta.value!="" ){
			    window.alert("<%=LitMsgHHasta%>");
			    return false;
		    }
	    }
    }

    if (tieneTPV == "SI") {
	    if ((parent.pantalla.document.tiendas.puerto.value!="") && (parent.pantalla.document.tiendas.numPuerto.value=="")) {
		    window.alert("<%=LitMsgDatosPuerto%>");
		    return false;
	    }
	    if ((parent.pantalla.document.tiendas.puerto.value=="") && (parent.pantalla.document.tiendas.numPuerto.value!="")) {
		    window.alert("<%=LitMsgDatosPuerto%>");
		    return false;
	    }
    }
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado)
{
	switch (mode) {
	    case "browse":
	        switch (pulsado) {
	            case "add": //Nuevo registro
	                parent.pantalla.document.tiendas.action = "tiendas.asp?mode=" + pulsado;
	                parent.pantalla.document.tiendas.submit();
	                document.location = "tiendas_bt.asp?mode=" + pulsado;
	                break;

	            case "edit": //Editar registro
	                parent.pantalla.document.tiendas.action = "tiendas.asp?codigo=" + parent.pantalla.document.tiendas.hcodigo.value +
					"&mode=" + pulsado + "&domicilio=" + parent.pantalla.document.tiendas.hdomicilio.value;
	                parent.pantalla.document.tiendas.submit();
	                document.location = "tiendas_bt.asp?mode=" + pulsado;
	                break;

	            case "delete": //Eliminar registro
	                if (window.confirm("<%=LitMsgEliminarAlmacenConfirm%>") == true) {
	                    per_del = 1;
	                    del_st = "";
	                    if (parseInt(parent.pantalla.document.tiendas.countusers.value) > 0) {
	                        if (!window.confirm("<%=LitMsgUsersAsigned%>")) {
	                            per_del = 0;
	                        }
                            del_st="&del_st=1"
	                    }
	                    if (parseInt(parent.pantalla.document.tiendas.countoperations.value) > 0) {
	                        window.alert("<%=LitErrOperation%>");
	                        per_del = 0;
	                    }
	                    if (per_del == 1) {
	                        parent.pantalla.document.tiendas.action = "tiendas.asp?mode=" + pulsado + "&codigo=" + parent.pantalla.document.tiendas.hcodigo.value + del_st;
	                        parent.pantalla.document.tiendas.submit();
	                        document.location = "tiendas_bt.asp?mode=browse";
	                    }
	                }
	                break;

	            case "search": //Buscar datos
	                break;
	        }
	        break;

		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos()) {
						parent.pantalla.document.tiendas.action="tiendas.asp?mode=save"+ "&codigo=" + parent.pantalla.document.tiendas.codigo.value +
						"&domicilio="+parent.pantalla.document.tiendas.domicilio.value;
						parent.pantalla.document.tiendas.submit();
						document.location="tiendas_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.tiendas.action="tiendas.asp?codigo=" + parent.pantalla.document.tiendas.hcodigo.value +
					"&mode=browse";
					parent.pantalla.document.tiendas.submit();
					document.location="tiendas_bt.asp?mode=browse";
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos()) {
						parent.pantalla.document.tiendas.action="tiendas.asp?mode=first_save&codigo="+parent.pantalla.document.tiendas.codigo.value
						+ "&domicilio="+parent.pantalla.document.tiendas.domicilio.value;
						parent.pantalla.document.tiendas.submit();
						document.location="tiendas_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.tiendas.action="tiendas.asp?mode=add";
					parent.pantalla.document.tiendas.submit();
					document.location="tiendas_bt.asp?mode=add";
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
					parent.pantalla.document.tiendas.action="tiendas.asp?mode=add&viene=asistente";
					parent.pantalla.document.tiendas.submit();
					document.location="tiendas_bt.asp?mode=add&viene=asistente";
					break;					
				case "cancel":
				    parent.pantalla.document.tiendas.action="tiendas.asp?mode=search&viene=asistente";
					parent.pantalla.document.tiendas.submit();
					document.location="tiendas_bt.asp?mode=search&viene=asistente";
					break;					
				case "save": //Guardar registro
					if (ValidarCampos()) {					   
					    parent.pantalla.document.tiendas.action="tiendas.asp?mode=save"+ "&codigo=" + parent.pantalla.document.tiendas.codigo.value +
						"&domicilio="+parent.pantalla.document.tiendas.domicilio.value;
						parent.pantalla.document.tiendas.submit();
						
						parent.pantalla.document.tiendas.action="tiendas.asp?mode=search&viene=asistente";
						parent.pantalla.document.tiendas.submit();
						document.location="tiendas_bt.asp?mode=search&viene=asistente";
					}
					break;				
					
				case "edit": //Editar registro
					parent.pantalla.document.tiendas.action="tiendas.asp?codigo=" + parent.pantalla.document.tiendas.hcodigo.value +
					"&mode=" + pulsado+ "&domicilio="+parent.pantalla.document.tiendas.hdomicilio.value;
					parent.pantalla.document.tiendas.submit();
					document.location="tiendas_bt.asp?mode=edit&viene=asistente";			
					break;
			}
			break;
	}
}

function comprobar_enter() {
	//si se ha pulsado la tecla enter
	Buscar();
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
</script>
<body class="body_master_ASP">
<%
    mode=Request.QueryString("mode")
    viene=limpiaCadena(Request.QueryString("viene"))
    tipo="1"
%> 
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(null_s(mode))%>" />
<div id="PageFooter_ASP" >
    <%if viene<>"asistente" then %>
        <div id="ControlPanelFooter_left_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
            <%
			    if mode="browse" then
				    %><td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,""%>
				    </td>
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBT LITBOTEDITAR,ImgEditar,ParamImgEditar,""%>
				    </td>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,""%>
				    </td><%
			    elseif mode="search" then
				    %><td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,""%>
				    </td><%
			    elseif mode="edit" then
				    %><td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				    </td><%
			    elseif mode="add" then
				    %><td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save');">
					    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
				    </td><%
			    end if%>
		        </tr>
	        </table>
            </div>
    
            <div id="FILTERS_MASTER_ASP">
                <select class="IN_S"name="campos">
                    <option value="codigo"><%=LitCodigo%></option>
                    <option selected value="descripcion"><%=LitDescripcion%></option>
                    <option value="domicilio"><%=LitDomicilio%></option>
                    <option value="cp"><%=LitCp%></option>
                    <option value="poblacion"><%=LitPoblacion%></option>
                    <option value="provincia"><%=LitProvincia%></option>
                    <option value="pais"><%=LitPais%></option>
                    <option value="telefono"><%=LitTel1%></option>
                    <option value="fax"><%=LitFax%></option>
                    <option value="observaciones"><%=LitObservaciones%></option>
                </select>
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
				</select>
				<input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
	  <%else%>
        <div id="ControlPanelFooter_left_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
                <%
			        if mode="browse" then
				        %>
                        <td id="idadd" class="CELDABOT" onclick="javascript:Accion('asistente','add');">
					        <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				        </td>
                        <td id="idedit" class="CELDABOT" onclick="javascript:Accion('asistente','edit');">
					        <%PintarBotonBT LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				        </td>
                        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel');">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
                        <%
			        elseif mode="search" then
				        %>
                        <td id="idadd" class="CELDABOT" onclick="javascript:Accion('asistente','add');">
					        <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				        </td>
                        <td id="idcancel" class="CELDABOT" onclick="javascript:Cerrar();">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
                        <td id="idbefore" class="CELDABOT" onclick="javascript:MoverPagPM('<%=enc.EncodeForJavascript(null_s(ruta))%>','2');">
					        <%PintarBotonBT LITBOTANTERIOR,ImgAnterior,ParamImgAnterior,LITBOTANTERIORTITLE%>
				        </td>
                        <td id="idnext" class="CELDABOT" onclick="javascript:MoverPagPM('<%=enc.EncodeForJavascript(null_s(ruta))%>','1');">
					        <%PintarBotonBT LITBOTSIGUIENTE,ImgSiguiente,ParamImgSiguiente,LITBOTSIGUIENTETITLE%>
				        </td>
                       <%
			        elseif mode="edit" then
				        %>
                        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('asistente','save');">
					        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				        </td>
                        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel');">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
                        <%
			        elseif mode="add" then
				        %>
                        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('asistente','save');">
					        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				        </td>
                        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asistente','cancel');">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
                        <%
			        end if%>			        
		        </tr>
	        </table>
            </div>
    <%end if%>
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