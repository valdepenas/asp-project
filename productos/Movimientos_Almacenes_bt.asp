<%@ Language=VBScript %>
<%' JCI 17/06/2003 : MIGRACION A MONOBASE%>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>

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
<!--#include file="../calculos.inc" -->

<!--#include file="movimientos_almacenes.inc" -->

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

var esperar_grabacion_lotes = 0;
var cuanto_espero = 5;

function GetXmlHttpObject()
{ 
   var xmlHttp=null; 
   try
   { 
        // Firefox, Opera 8.0+, Safari 
        xmlHttp=new XMLHttpRequest(); 
   }
   catch (e)
   { 
        //Internet Explorer 
       try
       { 
           xmlHttp=new ActiveXObject("Msxml2.XMLHTTP"); 
       }
       catch (e)
       { 
           xmlHttp=new ActiveXObject("Microsoft.XMLHTTP"); 
       } 
   } 
   return xmlHttp; 
 } 
 var xmlHttp, ServerResponse = null;
function ComprobarExisteLotesDestino(nmovimiento)
{
    ok=1;
    xmlHttp = GetXmlHttpObject();
    if (xmlHttp != null) 
    {
        //window.alert("los datos son-" + nmovimiento + "-");
        var url = "../fabricacion/lotes_asignar.asp?mode=consultaAJAX&nmovimiento=" + nmovimiento;
        //window.alert(url);
        xmlHttp.open("GET",url,false);
        xmlHttp.send(null);
        respuesta=xmlHttp.responseText;
        //window.alert(nmovimiento + "-" + respuesta + "-");
        if (respuesta.toUpperCase()=="OK")
        {
            ok=1;//esto todo correcto
        }
        else
        {
            ok=0;//no hay lotes de destino
        }
    }
    return ok;
}



function cambiarfecha(fecha,modo)
{
	var fecha_ar=new Array();
	if (fecha!="")
	{
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length){
			if (fecha.substring(l,l+1)=='/'){
				suma++;
				fecha_ar[suma]="";
			}
			else{
				if (fecha.substring(l,l+1)!='') fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
			}
			l++;
		}
		if (suma!=2) {
			window.alert("<%=LitFechaMal%> en el campo " + modo );
			return false;
		}
		else {
			nonumero=0;
			while (suma>=0 && nonumero==0){
				if (isNaN(fecha_ar[suma])) nonumero=1;
				if (fecha_ar[suma].length>2 && suma!=2) nonumero=1;
				if (fecha_ar[suma].length>4 && suma==2) nonumero=1;
				suma--;
			}
	
			if (nonumero==1){
				window.alert("<%=LitFechaMal%> en el campo " + modo);
				return false;
			}
		}
	}
	return true;
}

//Validación de campos numéricos y fechas.
function ValidarCampos(mode) {
	if (parent.pantalla.document.movimientos_almacenes.fecha.value=="") {
		window.alert("<%=LitMsgFechaNoNulo%>");
		return false;
	}

	if (!cambiarfecha(parent.pantalla.document.movimientos_almacenes.fecha.value,"FECHA MOVIMIENTO")){
		return false;
	}

    if (parent.pantalla.document.movimientos_almacenes.SAFTMOVEMENTSTARTTIME!=null)
    {
        if (parent.pantalla.document.movimientos_almacenes.SAFTMOVEMENTSTARTTIME.value!="")
        {
            /*
	        if (!cambiarfecha(parent.pantalla.document.movimientos_almacenes.SAFTMOVEMENTSTARTTIME.value,"FECHA SAFT")){
		        return false;
	        }
            */
            /*
	        if (!checkhora(parent.pantalla.document.movimientos_almacenes.SAFTMOVEMENTSTARTTIME)) {
		        window.alert("<%=LITFECHAMAL%>");
		        return false;
	        }
            */
	        if (!chkdatetime(parent.pantalla.document.movimientos_almacenes.SAFTMOVEMENTSTARTTIME.value)) {
		        window.alert("<%=LITFECHAMAL%>");
		        return;
	        }
        }
    }

	if (parent.pantalla.document.movimientos_almacenes.nserie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.movimientos_almacenes.responsable.value=="") {
		window.alert("<%=LitMsgResponsableNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.movimientos_almacenes.almdestino.value=="") {
		window.alert("<%=LitMsgAlmDestinoNoNulo%>");
		return false;
	}

	if(mode!="add"){
		h_mercrecibida=parent.pantalla.document.movimientos_almacenes.h_mercrecibida.value;
		mercrecibida=parent.pantalla.document.movimientos_almacenes.merric.checked;
		almdestino=parent.pantalla.document.movimientos_almacenes.almdestino.value;
		h_almdestino=parent.pantalla.document.movimientos_almacenes.h_almdestino.value;
		almdc=0;
		if (almdestino!=h_almdestino) almdc=1;
		if (mercrecibida==true && h_mercrecibida=="-1" && almdc==1){
			window.alert("<%=LitMercRecNoCambAlm%>");
			return false;
		}
	}
	return true;
}

function Buscar() {
	SearchPage("Movimientos_Almacenes_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + 
    "&mmp=" + parent.pantalla.document.movimientos_almacenes.mmp.value, 1);
	
    document.opciones.texto.value = "";
}
function GuardarLotesEnAlmDestino()
{
    esperar_grabacion_lotes=document.opciones.esperar_grabacion_lotes.value;
    //window.alert(esperar_grabacion_lotes);
    if (esperar_grabacion_lotes == 0 && cuanto_espero >= 0)
    {
        var t = setTimeout("GuardarLotesEnAlmDestino()", 500);
    }
    else {
        //window.alert("voy a grabar");
        parent.pantalla.document.movimientos_almacenes.action = "movimientos_almacenes.asp?nmovimiento=" + parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value + "&mode=save";
        parent.pantalla.document.movimientos_almacenes.submit();
        document.location = "movimientos_almacenes_bt.asp?mode=browse";
    }
    //Ricardo 08-08-2013 esperaremos indefinidamente
    //cuanto_espero--;
}
//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Nuevo registro
					//para que al dar al boton añadir, no se ponga por defecto el responsable del movimiento actual se borra el responsable
					if (parent.pantalla.document.movimientos_almacenes.mode.value=="browse")
						parent.pantalla.document.movimientos_almacenes.responsable.value="";
					parent.pantalla.document.movimientos_almacenes.action="movimientos_almacenes.asp?mode=" + pulsado + "&responsable=";
					parent.pantalla.document.movimientos_almacenes.submit();
					document.location="movimientos_almacenes_bt.asp?mode=" + pulsado;
					break;

				case "edit": //Editar registro
                    bloqueado=0;
                    try{
                        bloqueado=parent.pantalla.movimientos_almacenes.bloqueado.value;
                    }
                    catch(e)
                    {
                        bloqueado=0;
                    }
                    if (bloqueado==0)
                    {
					    parent.pantalla.document.movimientos_almacenes.action="movimientos_almacenes.asp?nmovimiento=" + parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value + "&mode=" + pulsado;
					    parent.pantalla.document.movimientos_almacenes.submit();
					    document.location="movimientos_almacenes_bt.asp?mode=" + pulsado;
                    }
                    else
                    {
                        window.alert("<%=LITNOMODIFBYSAFT%>");
                    }
					break;

				case "delete": //Eliminar registro
					h_mercrecibida=parent.pantalla.document.movimientos_almacenes.h_mercrecibida.value;
                    bloqueado=0;
                    try{
                        bloqueado=parent.pantalla.movimientos_almacenes.bloqueado.value;
                    }
                    catch(e)
                    {
                        bloqueado=0;
                    }
					if (h_mercrecibida==0)
                    {
                        if (bloqueado==0)
                        {
                            nmovimiento=parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value;
                            try{
                                existen_lotes_compra = parent.pantalla.document.movimientos_almacenes.existen_lotes_compra.value;
                            }
                            catch(e)
                            {
                                existen_lotes_compra=0;
                            }
                            parent.pantalla.comprobacionVentasEnLotes=0;
                            if (existen_lotes_compra==1)
                            {
                                parent.pantalla.ComprobarVentasConLoteCab(nmovimiento,"DELETE");
                            }
                            if (parent.pantalla.comprobacionVentasEnLotes==0)
                            {
						        if (window.confirm("<%=LitDeseaBorrarMovimientoConfirm%>")==true) {
							        parent.pantalla.document.movimientos_almacenes.action="movimientos_almacenes.asp?nmovimiento=" + parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value +
							        "&mode=" + pulsado + "&submode=" + parent.pantalla.document.movimientos_almacenes.mode.value;
							        parent.pantalla.document.movimientos_almacenes.submit();
							        document.location="movimientos_almacenes_bt.asp?mode=browse";
						        }
                            }
                            else{
                                window.alert("<%=LITMSGNOBORRARPORLOTE%>");
                            }
                        }
                        else
                        {
                            window.alert("<%=LITNOMODIFBYSAFT%>");
                        }
					}
					else window.alert("<%=LitMercRecNoBorrarMov%>");
					break;
			}
			break;

case "edit":
    switch (pulsado) {
        case "save": //Guardar registro
            if (ValidarCampos(mode)) {
                h_mercrecibida = parent.pantalla.document.movimientos_almacenes.h_mercrecibida.value;
                mercrecibida = parent.pantalla.document.movimientos_almacenes.merric.checked;
                texto_aviso_MR = "";
                no_hace_nada=0;
                /*trazas
                almdestino = parent.pantalla.document.movimientos_almacenes.almdestino.value;
                h_almdestino = parent.pantalla.document.movimientos_almacenes.h_almdestino.value;
                selitem = parent.pantalla.document.movimientos_almacenes.h_almdestino2.selectedIndex;
                existen_lotes_compra = parent.pantalla.document.movimientos_almacenes.existen_lotes_compra.value;
                window.alert("los datos son-" + h_mercrecibida + "-" + mercrecibida + "-" + almdestino + "-" + h_almdestino + "-" + selitem + "-" + existen_lotes_compra + "-");
                fin trazas*/
                almdestino_nom = "";
                if (mercrecibida == false && h_mercrecibida == "-1") {
                    if (parent.pantalla.document.movimientos_almacenes.h_almdestino2!=null)
                    {
                        try{
                            selitem = parent.pantalla.document.movimientos_almacenes.h_almdestino2.selectedIndex;
                        }
                        catch(e)
                        {
                            selitem="undefined";
                        }
                        if (selitem != "undefined") {
                            try{
                                almdestino_nom = parent.pantalla.document.movimientos_almacenes.h_almdestino2.options[selitem].text;
                            }
                            catch(e)
                            {
                                //almdestino_nom = "";
                            }
                        }
                        else {
                            almdestino_nom = "";
                        }
                    }
                    if (almdestino_nom=="")
                    {
                        if (parent.pantalla.document.movimientos_almacenes.almdestino!=null)
                        {
                            try{
                                selitem = parent.pantalla.document.movimientos_almacenes.almdestino.selectedIndex;
                            }
                            catch(e)
                            {
                                selitem="undefined";
                            }
                            if (selitem != "undefined") {
                                try{
                                    almdestino_nom = parent.pantalla.document.movimientos_almacenes.almdestino.options[selitem].text;
                                }
                                catch(e)
                                {
                                    //almdestino_nom = "";
                                }
                            }
                            else {
                                almdestino_nom = "";
                            }
                        }
                    }
                    if (almdestino_nom!=null && almdestino_nom!="")
                    {
                        texto_aviso_MR = "<%=LitEstSegMarcRec3%>" + almdestino_nom + ".<%=LitEstSegMarcRec2%>";
                    }
                    else{
                        texto_aviso_MR="";
                    }
                    /*si algun detalle tiene lote de compra y el movimiento tiene la mercancia recibida, comprobaremos que dicho lote no haya sido vendido*/
                    nmovimiento=parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value;
                    comprobacionVentasEnLotes=0;
                    try{
                        existen_lotes_compra = parent.pantalla.document.movimientos_almacenes.existen_lotes_compra.value;
                    }
                    catch(e)
                    {
                        existen_lotes_compra=0;
                    }
                    parent.pantalla.comprobacionVentasEnLotes=0;
                    if (existen_lotes_compra==1)
                    {
                        parent.pantalla.http = new XMLHttpRequest();
                        parent.pantalla.enProceso = false;
                        parent.pantalla.ComprobarVentasConLoteCab(nmovimiento,"SAVE");
                    }
                    if (parent.pantalla.comprobacionVentasEnLotes!=0)
                    {
                        texto_aviso_MR="";
                        no_hace_nada=1;
                        window.alert("<%=LITNODESBLMARCRECPORLOTE%>");
                    }
                }
                de_mercrecibida_a_mercrecibida = 0;
                if (mercrecibida == true && h_mercrecibida == "0") {
                    de_mercrecibida_a_mercrecibida = 1;
                    if (parent.pantalla.document.movimientos_almacenes.h_almdestino2!=null)
                    {
                        try{
                            selitem = parent.pantalla.document.movimientos_almacenes.h_almdestino2.selectedIndex;
                        }
                        catch(e)
                        {
                            selitem="undefined";
                        }
                        if (selitem != "undefined") {
                            almdestino_nom = parent.pantalla.document.movimientos_almacenes.h_almdestino2.options[selitem].text;
                        }
                        else {
                            almdestino_nom = "";
                        }
                    }
                    if (almdestino_nom=="")
                    {
                        if (parent.pantalla.document.movimientos_almacenes.almdestino!=null)
                        {
                            try{
                                selitem = parent.pantalla.document.movimientos_almacenes.almdestino.selectedIndex;
                            }
                            catch(e)
                            {
                                selitem="undefined";
                            }
                            if (selitem != "undefined") {
                                almdestino_nom = parent.pantalla.document.movimientos_almacenes.almdestino.options[selitem].text;
                            }
                            else {
                                almdestino_nom = "";
                            }
                        }
                    }
                    if (almdestino_nom!=null && almdestino_nom!="")
                    {
                        texto_aviso_MR = "<%=LitEstSegMarcRec1%>" + almdestino_nom + ".<%=LitEstSegMarcRec2%>";
                    }
                    else{
                        texto_aviso_MR="";
                    }

                }
                if (texto_aviso_MR != "") {
                    if (window.confirm(texto_aviso_MR) == true) {
                        //si algun detalle tiene lotes de entrada, saldra una pantalla para elegir que lote para el almacen destino
                        if (de_mercrecibida_a_mercrecibida == 1) {
                            try{
                                existen_lotes_compra = parent.pantalla.document.movimientos_almacenes.existen_lotes_compra.value;
                            }
                            catch(e)
                            {
                                existen_lotes_compra=0;
                            }
                        }
                        else {
                            existen_lotes_compra = 0;
                        }
                        if (existen_lotes_compra == 1)
                        {
                            nmovimiento=parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value;
                            //consulta AJAX para saber si van a salir datos en la pantalla de lotes_asignar, sino directamente iremos a guardar y alli crearemos los nuevos lotes
                            if (ComprobarExisteLotesDestino(nmovimiento)==1)
                            {
                                
                                AbrirVentana("../fabricacion/lotes_asignar.asp?mode=browse&nmovimiento=" + nmovimiento+ "&viene=movimientos_cabecera","P",<%=AltoVentana%>,<%=AnchoVentana%>);
                                esperar_grabacion_lotes = 0;
                                try{
                                    document.opciones.esperar_grabacion_lotes.value=0;
                                }
                                catch(e)
                                {
                                }
                                parent.pantalla.document.movimientos_almacenes.FaltaGrabarLotesCompra.value="0";
                            }
                            else{
                                esperar_grabacion_lotes = 1;
                                try{
                                    document.opciones.esperar_grabacion_lotes.value=1;
                                }
                                catch(e)
                                {
                                }
                                parent.pantalla.document.movimientos_almacenes.FaltaGrabarLotesCompra.value="1";
                            }
                        }
                        else {
                            esperar_grabacion_lotes = 1;
                            try{
                            document.opciones.esperar_grabacion_lotes.value=1;
                            }
                            catch(e)
                            {
                            }
                        }
                        GuardarLotesEnAlmDestino();
                    }
                }
                else {
                    if(no_hace_nada==0)
                    {
                        parent.pantalla.document.movimientos_almacenes.action = "movimientos_almacenes.asp?nmovimiento=" + parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value + "&mode=save";
                        parent.pantalla.document.movimientos_almacenes.submit();
                        document.location = "movimientos_almacenes_bt.asp?mode=browse";
                    }
                }
            }
            break;

        case "cancel": //Cancelar edición
            parent.pantalla.document.movimientos_almacenes.action = "movimientos_almacenes.asp?nmovimiento=" + parent.pantalla.document.movimientos_almacenes.h_nmovimiento.value + "&mode=browse";
            parent.pantalla.document.movimientos_almacenes.submit();
            document.location = "movimientos_almacenes_bt.asp?mode=browse";
            break;
    }
    break;

		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode)) {
						parent.pantalla.document.movimientos_almacenes.action="movimientos_almacenes.asp?mode=first_save";
						parent.pantalla.document.movimientos_almacenes.submit();
						document.location="movimientos_almacenes_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.movimientos_almacenes.action="movimientos_almacenes.asp?mode=add";
					parent.pantalla.document.movimientos_almacenes.submit();
					document.location="movimientos_almacenes_bt.asp?mode=add";
					break;
			}
			break;
	}
}
function comprobar_enter() {
    //si se ha pulsado la tecla enter
    //if (window.event.keyCode==13){
    //document.opciones.criterio.focus();
    Buscar();
    //}
}
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")
bma=limpiaCadena(Request.QueryString("bma"))%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
<input type="hidden" name="esperar_grabacion_lotes" value="0" />
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
				    <%if bma & "" <> "0" then%>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeftRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				    <%end if
			elseif mode="search" then%>
                   <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
			<%elseif mode="edit" then
				%>
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
				<select class="IN_S" name="campos">
          			<option value="nmovimiento"><%=LitMovimiento%></option>
					<option value="p.nombre"><%=LitResponsable%></option>
					<option value="a.descripcion"><%=LitAlmacenDestino%></option>
				</select>
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContieneB%></option>
					<option value="termina"><%=LitTerminaB%></option>
					<option value="igual"><%=LitIgualB%></option>
				</select>
                <input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
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