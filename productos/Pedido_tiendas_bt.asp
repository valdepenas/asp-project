<%@ Language=VBScript %>
<%' JCI 17/06/2003 : MIGRACION A MONOBASE%>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="Pedido_tiendas.inc" -->

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

function cambiarfecha(fecha,modo){
	var fecha_ar=new Array();

	if (fecha!=""){
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length){
			if (fecha.substring(l,l+1)=='/'){
				suma++;
				fecha_ar[suma]="";
			}
			else{
				if (fecha.substring(l,l+1)!=''){
					fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
				}
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
				if (isNaN(fecha_ar[suma])) {
					nonumero=1;
				}
				if (fecha_ar[suma].length>2 && suma!=2) {
					nonumero=1;
				}
				if (fecha_ar[suma].length>4 && suma==2) {
					nonumero=1;
				}
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
	if (parent.pantalla.document.pedido_tiendas.fecha.value=="") {
		window.alert("<%=LitMsgFechaNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.pedido_tiendas.almdestino.value==parent.pantalla.document.pedido_tiendas.almorigen.value)
	{
		window.alert("<%=LitMsgAlmOAlmDIguales%>");
		return false;
	}
	if (parent.pantalla.document.pedido_tiendas.responsable.value=="") {
		window.alert("<%=LitMsgResponsableNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.pedido_tiendas.almdestino.value=="") {
		window.alert("<%=LitMsgAlmDestinoNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.pedido_tiendas.almorigen.value=="") {
		window.alert("<%=LitMsgAlmDestinoNoNulo%>");
		return false;
    }
    if (parent.pantalla.document.pedido_tiendas.nserie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}
	return true;
}

function comprobar_enter() {
    //si se ha pulsado la tecla enter
    //if (window.event.keyCode==13){
    //document.opciones.criterio.focus();
    Buscar();
    //}
}

//DGB: change to page search
function Buscar() {
    SearchPage("pedido_tiendas_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	           "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + 
               "&mmp=" + parent.pantalla.document.pedido_tiendas.mmp.value, 1);

    document.opciones.texto.value = "";
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Nuevo registro
					//para que al dar al boton añadir, no se ponga por defecto el responsable del movimiento actual se borra el responsable
					if (parent.pantalla.document.pedido_tiendas.mode.value=="browse"){
						parent.pantalla.document.pedido_tiendas.responsable.value="";
					}
					parent.pantalla.document.pedido_tiendas.action="pedido_tiendas.asp?mode=" + pulsado + "&responsable=";
					parent.pantalla.document.pedido_tiendas.submit();
					document.location="pedido_tiendas_bt.asp?mode=" + pulsado;
					break;

				case "edit": //Editar registro
						parent.pantalla.document.pedido_tiendas.action="pedido_tiendas.asp?npedido=" + parent.pantalla.document.pedido_tiendas.h_npedido.value +
						"&mode=" + pulsado;
						parent.pantalla.document.pedido_tiendas.submit();
						document.location="pedido_tiendas_bt.asp?mode=" + pulsado;
					break;

				case "delete": //Eliminar registro
					h_nmovimiento=parent.pantalla.document.pedido_tiendas.h_nmovimiento.value;
					h_nalbaran=parent.pantalla.document.pedido_tiendas.h_nalbaran.value;
						
					if (h_nmovimiento=="" && h_nalbaran==""){
						if (window.confirm("<%=LitDeseaBorrarPedidoConfirm%>")==true) {
							parent.pantalla.document.pedido_tiendas.action="pedido_tiendas.asp?npedido=" + parent.pantalla.document.pedido_tiendas.h_npedido.value +
							"&mode=" + pulsado + "&submode=add" 
							parent.pantalla.document.pedido_tiendas.submit();
							document.location="pedido_tiendas_bt.asp?mode=add";
						}
					}
					else window.alert("<%=LitMovAlmNoBorrarPed%>");
					break;
			}
			break;

		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode)) {
						parent.pantalla.document.pedido_tiendas.action="pedido_tiendas.asp?npedido=" + parent.pantalla.document.pedido_tiendas.h_npedido.value + "&mode=save";
						parent.pantalla.document.pedido_tiendas.submit();
						document.location="pedido_tiendas_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.pedido_tiendas.action="pedido_tiendas.asp?npedido=" + parent.pantalla.document.pedido_tiendas.h_npedido.value + "&mode=browse";
					parent.pantalla.document.pedido_tiendas.submit();
					document.location="pedido_tiendas_bt.asp?mode=browse";
					break;
			}
			break;

		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode)) {
						parent.pantalla.document.pedido_tiendas.action="pedido_tiendas.asp?mode=first_save";
						parent.pantalla.document.pedido_tiendas.submit();
						document.location="pedido_tiendas_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.pedido_tiendas.action="pedido_tiendas.asp?mode=add";
					parent.pantalla.document.pedido_tiendas.submit();
					document.location="pedido_tiendas_bt.asp?mode=add";
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
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
				<select class="IN_S" name="campos">
          		    <option value="npedido"><%=LitPedido%></option>
					<option value="p.nombre"><%=LitResponsable%></option>
					<option value="a.descripcion"><%=LitAlmacenDestino%></option>
				</select>
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
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
