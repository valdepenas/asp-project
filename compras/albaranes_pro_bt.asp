<%@ Language=VBScript %>
<script id="DebugDirectives" runat="server" language="javascript">
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../calculos.inc" -->
<!--#include file="../constantes.inc" -->

<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="albaranes_pro.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->


<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/iban.js"></script>
<script language="javascript" type="text/javascript">
function comprobar_enter(){
	//si se ha pulsado la tecla enter
	//if (window.event.keyCode==13){
		//document.opciones.criterio.focus();
		Buscar();
	//}
}
function cambiarfecha(fecha,modo)
{
	var fecha_ar=new Array();

	if (fecha!="")
	{
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length)
		{
			if (fecha.substring(l,l+1)=='/')
			{
				suma++;
				fecha_ar[suma]="";
			}
			else
			{
				if (fecha.substring(l,l+1)!='') fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
			}
			l++;
		}
		if (suma!=2)
		{
			window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo );
			return false;
		}
		else
		{
			nonumero=0;
			while (suma>=0 && nonumero==0)
			{
				if (isNaN(fecha_ar[suma])) nonumero=1;
				if (fecha_ar[suma].length>2 && suma!=2) nonumero=1;
				if (fecha_ar[suma].length>4 && suma==2) nonumero=1;
				suma--;
			}

			if (nonumero==1)
			{
				window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
				return false;
			}
		}
	}
	return true;
}

//Validacion de campos numericos y fechas.
    //FLM: 200309: vailodo cuenta de abono con javascript.
function ComprobarCuentaBancaria()
{
    ok=true;
    if (parent.pantalla.document.albaranes_pro.ncuenta_pro.value!="")
    {
        while (parent.pantalla.document.albaranes_pro.ncuenta_pro.value.search(" ")!=-1)
        {
            parent.pantalla.document.albaranes_pro.ncuenta_pro.value=parent.pantalla.document.albaranes_pro.ncuenta_pro.value.replace(" ","");
        }
    
        cuenta=parent.pantalla.document.albaranes_pro.ncuenta_pro.value.substring(4,parent.pantalla.document.albaranes_pro.ncuenta_pro.value.length);
        //window.alert("los datos 1 son-" + cuenta + "-");
        if (isNaN(cuenta)) {
            window.alert("<%=LitCuentaAbonoError%>");
            ok=false;
        }

        var country = parent.pantalla.document.albaranes_pro.ncuenta_pro.value.substring(0,2);
        var i_iban = parent.pantalla.document.albaranes_pro.ncuenta_pro.value.substring(2,4);
        //window.alert("los datos 2 son-" + i_iban + "-");
        if (country == "") country ="ES";
        var account=parent.pantalla.document.albaranes_pro.ncuenta_pro.value.substring(4,parent.pantalla.document.albaranes_pro.ncuenta_pro.value.length);
        var iban = CreateIBAN(country, account);
        //window.alert("los datos 3 son-" + country + "-" + account + "-" + i_iban + "-" + iban + "-");
        if (i_iban != iban)
        {
            alert("<%=LitCuentaAbonoError%>");
            ok=false;
        }
    }
    return ok;
}
function ValidarCampos(mode)
{
	if (parent.pantalla.document.albaranes_pro.fecha.value=="") {
		window.alert("<%=LitMsgFechaNoNulo%>");
		return false;
	}

	if (!cambiarfecha(parent.pantalla.document.albaranes_pro.fecha.value,"FECHA ALBARAN")){
		return;
	}

	if (!checkdate(parent.pantalla.document.albaranes_pro.fecha)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return false;
    }
	if (parent.pantalla.document.albaranes_pro.serie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.albaranes_pro.divisa.value=="") {
		window.alert("<%=LitMsgDivisaNoNulo%>");
		return false;
	}

	//ricardo 20-1-2003
	//no se comprueba, para que al dejarlo vacio,se ponga el de nalbaran
	//if (parent.pantalla.document.albaranes_pro.nalbaran_pro.value=="") {
	//	window.alert("<%=LitMsgNalbaranProNoNulo%>");
	//	return false;
	//}
	if (comp_car_ext(parent.pantalla.document.albaranes_pro.nalbaran_pro.value,1)==1){
		window.alert("<%=LitMsgAlbpDesCarNoVal%>");
		return false;
	}


	if (parent.pantalla.document.albaranes_pro.nproveedor.value=="") {
		window.alert("<%=LitMsgProveedorNoNulo%>");
		return false;
	}
	
	//AMP validacion campo factor de cambio.
	factcambio=parent.pantalla.document.albaranes_pro.nfactcambio.value.replace(",","."); 		
    if (!/^([0-9])*[.]?[0-9]*$/.test(factcambio))
    { 
        alert("<%=LitMsgFactCambioI%>"); 
        return false;
    }
    if (parent.pantalla.document.albaranes_pro.nfactcambio.value=="")
    {
         alert("<%=LitMsgFactCambioI%>"); 
         return false;
    }

	// JMA 20/12/04. Campos personalizables.
	if ((mode=="add"||mode=="edit")&&(parent.pantalla.document.albaranes_pro.si_campo_personalizables.value==1)){
		num_campos=parent.pantalla.document.albaranes_pro.num_campos.value;

		respuesta=comprobarCampPerso("parent.pantalla.",num_campos,"albaranes_pro");
		if(respuesta!=0){
			titulo="titulo_campo" + respuesta;
			tipo="tipo_campo" + respuesta;
			titulo=parent.pantalla.document.albaranes_pro.elements[titulo].value;
			tipo=parent.pantalla.document.albaranes_pro.elements[tipo].value;
			if (tipo==4) nomTipo="<%=LitTipoNumerico%>";
			else if (tipo==5) {
				nomTipo="<%=LitTipoFecha%>";
			}

			window.alert("<%=LitMsgCampo%> " + titulo + " <%=LitMsgTipo%> " + nomTipo);

			return false;
		}
	}
    //FLM:200309 cuenta de abono correcta.
	cuenta_a_validar=parent.pantalla.document.albaranes_pro.ncuenta_pro.value.substring(4,parent.pantalla.document.albaranes_pro.ncuenta_pro.value.length)
	//window.alert("los datos 3 son-" + cuenta_a_validar + "-");
	if(validarCCC(cuenta_a_validar)==false){
	    window.alert("<%=LitCuentaAbonoError%>");
	    return false;
	}

	//ricardo 15-1-2008 si editamos , no podremos borrar el nalbaran_pro
	if (mode=="edit" && parent.pantalla.document.albaranes_pro.nalbaran_pro.value=="") {
        window.alert("<%=LitMsgNalbaranProNoNulo%>");
        return false;
    }
	if (ComprobarCuentaBancaria()==false)
	{
	    return false;
	}
	return true;
}

function Buscar()
{
    parent.pantalla.document.albaranes_pro.campo.value="";
    parent.pantalla.document.albaranes_pro.texto.value="";
    parent.pantalla.document.albaranes_pro.criterio.value="";
    parent.pantalla.document.albaranes_pro.lote.value="";
    parent.pantalla.document.albaranes_pro.total_paginas.value="";
	
    SearchPage("deliveryNote_pro_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value +
    "&modd=" + parent.pantalla.document.albaranes_pro.modd.value  + "&modi=" + parent.pantalla.document.albaranes_pro.modi.value +
    "&modp=" + parent.pantalla.document.albaranes_pro.modp.value ,1);

    document.opciones.texto.value = "";
}


//Realizar la accion correspondiente al boton pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Nuevo registro
					if (parent.pantalla.document.albaranes_pro.mode.value!="search"){
						if (parent.pantalla.document.albaranes_pro.mode.value!="browse" && parent.pantalla.document.albaranes_pro.mode.value!="first_save" && parent.pantalla.document.albaranes_pro.mode.value!="save" && parent.pantalla.document.albaranes_pro.mode.value!="delete")
							parent.pantalla.document.albaranes_pro.nproveedor.value="";
						parent.pantalla.document.albaranes_pro.h_nproveedor.value="";
					}
					parent.pantalla.document.albaranes_pro.action="albaranes_pro.asp?mode=" + pulsado;
					parent.pantalla.document.albaranes_pro.submit();
					document.location="albaranes_pro_bt.asp?mode=" + pulsado;
					break;

				case "edit": //Editar registro				   
				    if (parent.pantalla.document.albaranes_pro.h_nbalance.value>"") alert("<%=LitMsgModifAlbaranBalance%>");
					else if (parent.pantalla.document.albaranes_pro.h_nfactura.value=="NO") {
						parent.pantalla.document.albaranes_pro.action="albaranes_pro.asp?nalbaran=" + parent.pantalla.document.albaranes_pro.h_nalbaran.value +
						"&mode=" + pulsado;
						parent.pantalla.document.albaranes_pro.submit();
						document.location="albaranes_pro_bt.asp?mode=" + pulsado;
					}
					else alert("<%=LitMsgModifAlbaran%>" + parent.pantalla.document.albaranes_pro.h_nfactura.value);
					break;

				case "edit2": //Editar Detalles del documento
					if (parent.pantalla.document.albaranes_pro.h_nbalance.value>"") alert("<%=LitMsgModifAlbaranBalance%>");
					else if (parent.pantalla.document.albaranes_pro.h_nfactura.value=="NO") {
						document.location.ref="albaranes_pro_bt.asp?mode=browse";
						pagina="../central.asp?pag1=compras/albaranes_prodet.asp&ndoc=" + parent.pantalla.document.albaranes_pro.h_nalbaran.value +
						"&nproveedor=" + parent.pantalla.document.albaranes_pro.h_nproveedor.value + "&mode=browse&pag2=compras/albaranes_prodet_bt.asp&titulo=<%=LitDetAlb%> " + parent.pantalla.document.albaranes_pro.h_nalbaran.value;
						ven=AbrirVentana(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
					}
					else alert("<%=LitMsgModifAlbaran%>" + parent.pantalla.document.albaranes_pro.h_nfactura.value);
					break;

				case "delete": //Eliminar registro
					if (parent.pantalla.document.albaranes_pro.h_nbalance.value>"")alert("<%=LitMsgEliminarAlbaranBalance%>");
					else if (window.confirm("<%=LitMsgEliminarAlbaranConfirm%>")==true) {
						if (parent.pantalla.document.albaranes_pro.h_nfactura.value=="NO") {
							parent.pantalla.document.albaranes_pro.action="albaranes_pro.asp?nalbaran=" + parent.pantalla.document.albaranes_pro.h_nalbaran.value +
							"&mode=" + pulsado;
							parent.pantalla.document.albaranes_pro.submit();
							document.location="albaranes_pro_bt.asp?mode=browse";
						}
						else alert("<%=LitMsgBorrarAlbaran%>" + parent.pantalla.document.albaranes_pro.h_nfactura.value);
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
                        if (parent.pantalla.document.albaranes_pro.h_nproveedor.value!=(nempresa + parent.pantalla.document.albaranes_pro.nproveedor.value) ||
	                        parent.pantalla.document.albaranes_pro.h_fecha.value!=parent.pantalla.document.albaranes_pro.fecha.value ||
	                        parent.pantalla.document.albaranes_pro.h_divisa.value!=parent.pantalla.document.albaranes_pro.olddivisa.value){
	                        if (window.confirm("<%=LitMsgCamPropDocCamPrec%>")==false) recalcular_importes=0;
                        }

						parent.pantalla.document.albaranes_pro.action="albaranes_pro.asp?nalbaran=" + parent.pantalla.document.albaranes_pro.h_nalbaran.value +
						"&mode=save" + "&recalcular_importes=" + recalcular_importes;
						parent.pantalla.document.albaranes_pro.submit();
						document.location="albaranes_pro_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edici�n
				    parent.pantalla.document.albaranes_pro.divisafc.value="";
				    parent.pantalla.document.albaranes_pro.h_divisa.value = "";
					parent.pantalla.document.albaranes_pro.action="albaranes_pro.asp?nalbaran=" + parent.pantalla.document.albaranes_pro.h_nalbaran.value +
					"&mode=browse";
					parent.pantalla.document.albaranes_pro.submit();
					document.location="albaranes_pro_bt.asp?mode=browse";
					break;
			}
			break;

		case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode)) {
						parent.pantalla.document.albaranes_pro.action="albaranes_pro.asp?mode=first_save";
						parent.pantalla.document.albaranes_pro.submit();
						document.location="albaranes_pro_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edicion
				    parent.pantalla.document.albaranes_pro.divisafc.value="";
					parent.pantalla.document.albaranes_pro.nproveedor.value="";
					parent.pantalla.document.albaranes_pro.serie.value="";
					parent.pantalla.document.albaranes_pro.action="albaranes_pro.asp?mode=add&viene=cancelar";
					parent.pantalla.document.albaranes_pro.submit();
					document.location="albaranes_pro_bt.asp?mode=add";
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%mode=limpiaCadena(Request.QueryString("mode"))

' DGM 12/01/11 Recogida de parametros para ocultar Editar/Borrar
dim oeditar
oeditar = "0"
dim oborrar
oborrar = "0"
obtenerparametros("albaranes_pro_det")%>
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
				<%if cstr(oeditar) = "0" then %>
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBTLeft LITBOTEDITCAB,ImgEditar_Cab,ParamImgEditar_Cab,LITBOTEDITCABTITLE%>
				    </td>
		        <%end if
		        if cstr(oborrar) = "0" then%>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeft LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				<%end if
			elseif mode="search" then%>
                <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					<%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				</td>
			<%elseif mode="edit" then%>
                <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="add" then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%end if%>
		</tr>
	</table>
    </div>
    
    <div id="FILTERS_MASTER_ASP">
			<!--<td CLASS=CELDABOT><%=LitBuscar & ": "%>-->
				<select class="IN_S" name="campos">
          			<option selected value="nalbaran"><%=LitAlbaran%></option>
          			<!--<option value="nombre"><%=LitProveedor%></option>-->
					<!--<option value="razon_social"><%=LitRazonSocial%></option>-->
					<option value="razon_social"><%=LitProveedor%></option>
        		</select>
        	<!--</td><td CLASS=CELDABOT>-->
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<!--<option value="empieza"><%=LitComienza%></option>-->
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
				</select>
			<!--</td><td CLASS=CELDABOT>-->
				<input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
		    <!--</td><td CLASS=CELDABOT>-->
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