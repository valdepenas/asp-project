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
</head>

<!--#include file="../calculos.inc" -->
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="pedidos_pro.inc" -->

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
</script>


<script language="javascript" type="text/javascript">
function comprobar_enter(){
	//si se ha pulsado la tecla enter
	//if (window.event.keyCode==13){
		//document.opciones.criterio.focus();
		Buscar();
	//}
}
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
			window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo );
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
				window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
				return false;
			}
		}
	}
	return true;
}

function Buscar() {
    parent.pantalla.document.pedidos_pro.campo.value="";
    parent.pantalla.document.pedidos_pro.texto.value="";
    parent.pantalla.document.pedidos_pro.criterio.value="";
    parent.pantalla.document.pedidos_pro.lote.value="";
    parent.pantalla.document.pedidos_pro.total_paginas.value="";

    SearchPage("purchaseOrder_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value,1);

    document.opciones.texto.value = "";
}

//Validación de campos numéricos y fechas.
function ValidarCampos(mode) {
	if (parent.pantalla.document.pedidos_pro.fecha.value=="") {
		window.alert("<%=LitMsgFechaNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.pedidos_pro.serie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.pedidos_pro.nproveedor.value=="") {
		window.alert("<%=LitMsgProveedorNoNulo%>");
		return false;
	}

	if (!cambiarfecha(parent.pantalla.document.pedidos_pro.fecha.value,"FECHA PEDIDO")) return false;

	if (!checkdate(parent.pantalla.document.pedidos_pro.fecha)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return;
	}

	if (!cambiarfecha(parent.pantalla.document.pedidos_pro.fecha_entrega.value,"FECHA ENTREGA")) return false;

	if (!checkdate(parent.pantalla.document.pedidos_pro.fecha_entrega)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return;
	}
	
	if (!cambiarfecha(parent.pantalla.document.pedidos_pro.salida.value,"FECHA DE PAGO")) return false;

	if (!checkdate(parent.pantalla.document.pedidos_pro.salida)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return;
	}
	
	//AMP validación campo factor de cambio.
	factcambio=parent.pantalla.document.pedidos_pro.nfactcambio.value.replace(",","."); 		
    if (!/^([0-9])*[.]?[0-9]*$/.test(factcambio))
    { 
        alert("<%=LitMsgFactCambioI%>"); 
        return false;
    }
    if (parent.pantalla.document.pedidos_pro.nfactcambio.value=="")
    {
         alert("<%=LitMsgFactCambioI%>"); 
         return false;
    }

	// JMA 17/12/04. Campos personalizables.
	if ((mode=="add"||mode=="edit")&&(parent.pantalla.document.pedidos_pro.si_campo_personalizables.value==1)){
		num_campos=parent.pantalla.document.pedidos_pro.num_campos.value;

		respuesta=comprobarCampPerso("parent.pantalla.",num_campos,"pedidos_pro");
		if(respuesta!=0){
			titulo="titulo_campo" + respuesta;
			tipo="tipo_campo" + respuesta;
			titulo=parent.pantalla.document.pedidos_pro.elements[titulo].value;
			tipo=parent.pantalla.document.pedidos_pro.elements[tipo].value;
			if (tipo==4) nomTipo="<%=LitTipoNumerico%>";
			else if (tipo==5) {
				nomTipo="<%=LitTipoFecha%>";
			}

			alert("<%=LitMsgCampo%> " + titulo + " <%=LitMsgTipo%> " + nomTipo);
			return false;
		}
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Nuevo registro
					if (parent.pantalla.document.pedidos_pro.mode.value!="search"){
						if (parent.pantalla.document.pedidos_pro.mode.value!="browse" && parent.pantalla.document.pedidos_pro.mode.value!="first_save" && parent.pantalla.document.pedidos_pro.mode.value!="save" && parent.pantalla.document.pedidos_pro.mode.value!="delete")
							parent.pantalla.document.pedidos_pro.nproveedor.value="";
						parent.pantalla.document.pedidos_pro.h_nproveedor.value="";
					}
					parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?mode=" + pulsado;
					parent.pantalla.document.pedidos_pro.submit();
					document.location="pedidos_pro_bt.asp?mode=" + pulsado;
					break;

				case "edit": //Editar registro
					if (parent.pantalla.document.pedidos_pro.h_nalbaranpro.value=="NO") {
						if (parent.pantalla.document.pedidos_pro.h_nfacturapro.value=="NO") {
							parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?npedido=" + parent.pantalla.document.pedidos_pro.h_npedido.value +
							"&mode=" + pulsado;
							parent.pantalla.document.pedidos_pro.submit();
							document.location="pedidos_pro_bt.asp?mode=" + pulsado;
						}
						else alert("<%=LitMsgModifPedidoF%>" + parent.pantalla.document.pedidos_pro.h_nfacturapro.value);
					}
					else alert("<%=LitMsgModifPedidoA%>" + parent.pantalla.document.pedidos_pro.h_nalbaranpro.value);
					break;

				case "delete": //Eliminar registro
					if(parent.pantalla.document.pedidos_pro.borrarpedido.value=="NO") alert("<%=LitMsgBorradoConvertidoParcial%>");
					else
					{
						if(parent.pantalla.document.pedidos_pro.convertidoPedCli.value=="SI")
						{
							if(window.confirm("<%=LitMsgBorraPedidoConvertido%>"))
							{
								parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?npedido=" + parent.pantalla.document.pedidos_pro.h_npedido.value + "&mode=" + pulsado;
								parent.pantalla.document.pedidos_pro.submit();
								document.location="pedidos_pro_bt.asp?mode=browse";
							}

						}
						else if (window.confirm("<%=LitMsgEliminarPedidoConfirm%>")==true)
						{
							parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?npedido=" + parent.pantalla.document.pedidos_pro.h_npedido.value + "&mode=" + pulsado;
							parent.pantalla.document.pedidos_pro.submit();
							document.location="pedidos_pro_bt.asp?mode=browse";
						}
					}
					break;
			}
			break;

		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode)) {
                        //ricardo 25/4/2003 para que salga el mensaje de que se ha cambiado las propiedades del documento
                        // y que puede afectar al importe de los detalles
                        nempresa="<%=session("ncliente")%>";
                        recalcular_importes=1;
                        if (parent.pantalla.document.pedidos_pro.h_nproveedor.value!=(nempresa + parent.pantalla.document.pedidos_pro.nproveedor.value) ||
	                        parent.pantalla.document.pedidos_pro.h_fecha.value!=parent.pantalla.document.pedidos_pro.fecha.value ||
	                        parent.pantalla.document.pedidos_pro.h_divisa.value!=parent.pantalla.document.pedidos_pro.olddivisa.value){
	                        if (window.confirm("<%=LitMsgCamPropDocCamPrec%>")==false) recalcular_importes=0;
                        }
						parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?npedido=" + parent.pantalla.document.pedidos_pro.h_npedido.value +
						"&mode=save" + "&recalcular_importes=" + recalcular_importes;
						parent.pantalla.document.pedidos_pro.submit();
						document.location="pedidos_pro_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
				    parent.pantalla.document.pedidos_pro.divisafc.value="";
					parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?npedido=" + parent.pantalla.document.pedidos_pro.h_npedido.value +
					"&mode=browse";
					parent.pantalla.document.pedidos_pro.submit();
					document.location="pedidos_pro_bt.asp?mode=browse";
					break;
			}
			break;

		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode)) {
						parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?mode=first_save";
						parent.pantalla.document.pedidos_pro.submit();
						document.location="pedidos_pro_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
				    parent.pantalla.document.pedidos_pro.divisafc.value="";
					parent.pantalla.document.pedidos_pro.nproveedor.value="";
					parent.pantalla.document.pedidos_pro.serie.value="";
					parent.pantalla.document.pedidos_pro.action="pedidos_pro.asp?mode=add";
					parent.pantalla.document.pedidos_pro.submit();
					document.location="pedidos_pro_bt.asp?mode=add";
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")
    ' DGM 12/01/11 Recogida de parametros para ocultar Editar/Borrar
    dim oeditar
    oeditar = "0"
    dim oborrar
    oborrar = "0"
    obtenerparametros("pedidos_pro_det")%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=mode%>" />
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
		    <%if mode="browse" then%>
				<td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					<%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				</td>
				<%if cstr(oeditar) ="0" then %>
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBTLeft LITBOTEDITCAB,ImgEditar_Cab,ParamImgEditar_Cab,LITBOTEDITCABTITLE%>
				    </td>
			    <%end if
			    if cstr(oborrar) = "0" then %>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeftRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				<%end if
			elseif mode="search" then%>
                <td class="CELDABOT" onclick="javascript:Accion('browse','add');">
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
			<!--<td class=CELDABOT>-->
				<select class="IN_S" name="campos"><%=LitBuscar & ": "%>
					<option value="npedido"><%=LitPedido%></option>
			        <!--<option value="nombre"><%=LitProveedor%></option>-->
			        <!--<option value="razon_social"><%=LitRazonSocial%></option>-->
			        <option value="razon_social"><%=LitProveedor%></option>
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