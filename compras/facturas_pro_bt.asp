<%@ Language=VBScript %>
<% 
dim enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
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
<!--#include file="../varios.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="facturas_pro.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/iban.js"></script>
<script language="javascript" type="text/javascript">
    window.addEventListener("load", function hideFilterMaster() {
        var classNameLeft = parent.parent.document.getElementById('left').className;;
        if (classNameLeft == "search-panel open") {
            document.getElementById("FILTERS_MASTER_ASP").setAttribute("style", "display:none")
        }
    });

    function comprobar_enter(){
        Buscar();
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

                if (nonumero==1){
                    window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
                    return false;
                }
            }
        }
        return true;
    }

    function Buscar()
    {
        SearchPage("facturas_pro_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
        "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value,1);
        document.opciones.texto.value = "";
    }

    //Validación de campos numéricos y fechas.
    //FLM: 200309: vailodo cuenta de abono con javascript.
    function ComprobarCuentaBancaria()
    {
        ok=true;
        if (parent.pantalla.document.facturas_pro.ncuenta_pro.value!="")
        {
            while (parent.pantalla.document.facturas_pro.ncuenta_pro.value.search(" ")!=-1)
            {
                parent.pantalla.document.facturas_pro.ncuenta_pro.value=parent.pantalla.document.facturas_pro.ncuenta_pro.value.replace(" ","");
            }
    
            cuenta=parent.pantalla.document.facturas_pro.ncuenta_pro.value.substring(4,parent.pantalla.document.facturas_pro.ncuenta_pro.value.length);
            if (isNaN(cuenta)) {
                window.alert("<%=LitCuentaAbonoError%>");
                ok=false;
            }

            var country = parent.pantalla.document.facturas_pro.ncuenta_pro.value.substring(0,2);
            var i_iban = parent.pantalla.document.facturas_pro.ncuenta_pro.value.substring(2,4);
            if (country == "") country ="ES";
            var account=parent.pantalla.document.facturas_pro.ncuenta_pro.value.substring(4,parent.pantalla.document.facturas_pro.ncuenta_pro.value.length);
            var iban = CreateIBAN(country, account);
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
        if (parent.pantalla.document.facturas_pro.serie.value=="")
        {
            window.alert("<%=LitMsgSerieNoNulo%>");
            return false;
        }
        if (parent.pantalla.document.facturas_pro.nfactura_pro.value=="")
        {
            window.alert("<%=LitMsgFacturaNoNulo%>");
            return false;
        }

        if (parent.pantalla.document.facturas_pro.razon_social.value=="")
        {
            window.alert("<%=LitMsgProveedorNoExiste%>");
            return false;
        }

        if (comp_car_ext(parent.pantalla.document.facturas_pro.nfactura_pro.value,1)==1)
        {
            window.alert("<%=LitMsgFacpDesCarNoVal%>");
            return false;
        }

        if (parent.pantalla.document.facturas_pro.pagada.checked)
        {
            if (parent.pantalla.document.facturas_pro.h_pagada.value == 0)
            {
                if (!window.confirm("<%=LitMsgPagFactVencConfirm%>")) return false;
            }
            else
            {
                window.alert("<%=LitMsgFacturaNoModif%>");
                return false;
            }
        }
        else{
            if (parent.pantalla.document.facturas_pro.h_pagada.value == 1)
            {
                if (!window.confirm("<%=LitMsgAnPagFactVencConfirm%>")) return false;
            }
        }
        if (parent.pantalla.document.facturas_pro.nproveedor.value=="") {
            window.alert("<%=LitMsgProveedorNoNulo%>");
            return false;
        }

        if (parent.pantalla.document.facturas_pro.fecha.value==""){
            window.alert("<%=LitMsgFechaNoNulo%>");
            return false;
        }

        if (!cambiarfecha(parent.pantalla.document.facturas_pro.fecha.value,"FECHA FACTURA")) return false;

        if (!checkdate(parent.pantalla.document.facturas_pro.fecha))
        {
            window.alert("<%=LitMsgFechaFecha%>");
            return;
        }
	
        //AMP validación campo factor de cambio.
        factcambio=parent.pantalla.document.facturas_pro.nfactcambio.value.replace(",","."); 		
        if (!/^([0-9])*[.]?[0-9]*$/.test(factcambio))
        { 
            alert("<%=LitMsgFactCambioI%>"); 
            return false;
        }
        if (parent.pantalla.document.facturas_pro.nfactcambio.value=="")
        {
            alert("<%=LitMsgFactCambioI%>"); 
            return false;
        }


        // JMA 20/12/04. Campos personalizables.
        if ((mode=="add"||mode=="edit")&&(parent.pantalla.document.facturas_pro.si_campo_personalizables.value==1))
        {
            num_campos=parent.pantalla.document.facturas_pro.num_campos.value;

            respuesta=comprobarCampPerso("parent.pantalla.",num_campos,"facturas_pro");
            if(respuesta!=0)
            {
                titulo="titulo_campo" + respuesta;
                tipo="tipo_campo" + respuesta;
                titulo=parent.pantalla.document.facturas_pro.elements[titulo].value;
                tipo=parent.pantalla.document.facturas_pro.elements[tipo].value;
                if (tipo==4) nomTipo="<%=LitTipoNumerico%>";
                else if (tipo==5)
                {
                    nomTipo="<%=LitTipoFecha%>";
                }

                window.alert("<%=LitMsgCampo%> " + titulo + " <%=LitMsgTipo%> " + nomTipo);
                return false;
            }
        }
        //FLM:200309 cuenta de abono correcta.
        if (parent.pantalla.document.facturas_pro.ncuenta_pro.value.substring(2, 2) == "ES")
        {
            cuenta_a_validar=parent.pantalla.document.facturas_pro.ncuenta_pro.value.substring(4,parent.pantalla.document.facturas_pro.ncuenta_pro.value.length)
            if(validarCCC(cuenta_a_validar)==false)
            {
                window.alert("<%=LitCuentaAbonoError%>");
                return false;
            }
            if (ComprobarCuentaBancaria()==false)
            {
                return false;
            }
        }
        return true;
    }

    //Realizar la acción correspondiente al botón pulsado.
    function Accion(mode,pulsado)
    {
        gen_vencimiento=parent.pantalla.document.facturas_pro.gen_vencimiento.value;
        switch (mode) {
            case "browse":
                switch (pulsado) {
                    case "add": //Nuevo registro
                        if (parent.pantalla.document.facturas_pro.mode.value!="search"){
                            if (parent.pantalla.document.facturas_pro.mode.value!="browse" && parent.pantalla.document.facturas_pro.mode.value!="first_save" && parent.pantalla.document.facturas_pro.mode.value!="save" && parent.pantalla.document.facturas_pro.mode.value!="delete"){
                                parent.pantalla.document.facturas_pro.nproveedor.value="";
                            }
                            parent.pantalla.document.facturas_pro.h_nproveedor.value="";
                        }
                        parent.pantalla.document.facturas_pro.action="facturas_pro.asp?mode=" + pulsado;
                        parent.pantalla.document.facturas_pro.submit();
                        document.location="facturas_pro_bt.asp?mode=" + pulsado;
                        break;

                    case "edit": //Editar registro
                        if (parent.pantalla.document.facturas_pro.h_contabilizada.value=="1" && document.opciones.bloqContab.value=="1")
                            alert("<%=Lit_NoModContab%>");
                        else {
                            if(parent.pantalla.document.facturas_pro.h_nbalance.value!="") alert('<%=LitMsgFactTieneBalance%>');
                            else
                            {
                                parent.pantalla.document.facturas_pro.action="facturas_pro.asp?nfactura=" + parent.pantalla.document.facturas_pro.nfactura.value +
                                "&mode=" + pulsado;
                                parent.pantalla.document.facturas_pro.submit();
                                document.location="facturas_pro_bt.asp?mode=" + pulsado;
                            }
                        }
                        break;

                    case "delete": //Eliminar registro
                        if (parent.pantalla.document.facturas_pro.h_contabilizada.value=="1" && document.opciones.bloqContab.value=="1")
                            alert("<%=Lit_NoModContab%>");
                        else {
                            if (parent.pantalla.document.facturas_pro.nasiento.value!="") alert('<%=LitMsgFactTieneAsiento%>');
                            else if(parent.pantalla.document.facturas_pro.h_nbalance.value!=""){
                                window.alert('<%=LitMsgFactTieneBalance%>');
                            }
                            else
                            {
                                if (parent.pantalla.document.facturas_pro.factura_cli.value!="")
                                {
                                    if(window.confirm("<%=LitProvieneCliente%>"))
                                    {
                                        parent.pantalla.document.facturas_pro.action="facturas_pro.asp?mode=" + pulsado;
                                        parent.pantalla.document.facturas_pro.submit();
                                        document.location="facturas_pro_bt.asp?mode=browse";

                                    }
                                }
                                else if (window.confirm("<%=LitMsgEliminarFacturasConfirm%>")==true) {
                                    parent.pantalla.document.facturas_pro.action="facturas_pro.asp?mode=" + pulsado;
                                    parent.pantalla.document.facturas_pro.submit();
                                    document.location="facturas_pro_bt.asp?mode=browse";
                                }
                            }
                        }
                        break;
                }
                break;

            case "edit":
                switch (pulsado) {
                    case "save": //Guardar registro
                        cont="SI";
                        contLlekoAdmin="SI";
                       if(parent.pantalla.document.facturas_pro.h_llekoAdmin!=null && parent.pantalla.document.facturas_pro.h_llekoAdmin.value=="SI"){
                            if(parent.pantalla.document.facturas_pro.isCollectionRecipient!=null && parent.pantalla.document.facturas_pro.isCollectionRecipient.value!=""){
                                if (parent.pantalla.document.facturas_pro.h_pagada.value=="1" && !parent.pantalla.document.facturas_pro.pagada.checked){
                                    if (!confirm("<%=LitConfirmUpdateStatus%>")){ 
                                        contLlekoAdmin="NO"; 
                                    }                                   
                                }
                                if (parent.pantalla.document.facturas_pro.h_pagada.value=="0" && parent.pantalla.document.facturas_pro.pagada.checked){
                                    if (!confirm("<%=LitConfirmUpdateStatus2%>")){ 
                                        contLlekoAdmin="NO"; 
                                    }                                   
                                }
                            }
                        }  
                        if(contLlekoAdmin=="SI"){
                            if ((parent.pantalla.document.facturas_pro.h_pagada.value=="0") && (parent.pantalla.document.facturas_pro.pagada.checked)) {
                                if (!confirm("<%=LitMsgPagFacturaSinCajaConfirm%>")) cont="NO";
                            }
                            if (cont=="SI") {
                                if (ValidarCampos(mode)) {

                                    // DGM 16/05/2012
                                    if (parent.pantalla.document.facturas_pro.invoiceHasCosts != null)
                                        var hasCosts = parent.pantalla.document.facturas_pro.invoiceHasCosts.value;
                                    else
                                        var hasCosts = "0";

                                    if (parent.pantalla.document.facturas_pro.cod_proyecto != null)                                
                                        var codPro = parent.pantalla.document.facturas_pro.cod_proyecto.value;
                                    else
                                        var codPro = "";

                                    if (parent.pantalla.document.facturas_pro.cod_proyectoOLD != null)
                                        var codProOld = parent.pantalla.document.facturas_pro.cod_proyectoOLD.value;
                                    else
                                        var codProOld = "";
                        
                                    if (hasCosts!= "0" && codPro != "" && codProOld == ""){
                                        if (!window.confirm("<%=LITCONFIRMDELETECOSTS %>")){
                                            parent.pantalla.document.facturas_pro.cod_proyecto.value = "";
                                            paramCosts="";
                                        }
                                        else{
                                            paramCosts = "&delCosts=1";
                                        }
                                    }
                                    else
                                        paramCosts="";
                                    if(parent.pantalla.document.facturas_pro.h_llekoAdmin!=null && parent.pantalla.document.facturas_pro.h_llekoAdmin!="" && parent.pantalla.document.facturas_pro.h_llekoAdmin.value=="SI")
                                    {      
                                    
                                    }
                                    //ricardo 25/4/2003 para que salga el mensaje de que se ha cambiado las propiedades del documento
                                    // y que puede afectar al importe de los detalles
                                    nempresa="<%=session("ncliente")%>";
                                    recalcular_importes=1;
                                    if (parent.pantalla.document.facturas_pro.h_nproveedor.value!=(nempresa + parent.pantalla.document.facturas_pro.nproveedor.value) ||
                                    parent.pantalla.document.facturas_pro.h_fecha.value!=parent.pantalla.document.facturas_pro.fecha.value ||
                                    parent.pantalla.document.facturas_pro.divisa.value!=parent.pantalla.document.facturas_pro.olddivisa.value){
                                        if (window.confirm("<%=LitMsgCamPropDocCamPrec%>")==false) recalcular_importes=0;
                                    }
                                    continuar=1;
                                    continuarf=1;
                                    continuari=1;                                
                                    if (gen_vencimiento=='-1' && (parent.pantalla.document.facturas_pro.forma_pago.value!=parent.pantalla.document.facturas_pro.forma_pago_ant.value)){
                                        if(window.confirm("<%=LitMsgGenVenForPagConfirm%>")) continuar=0;
                                        else continuar=1;
                                    }
                                    if (gen_vencimiento=='-1' && (parent.pantalla.document.facturas_pro.fecha.value!=parent.pantalla.document.facturas_pro.fecha_ant.value)){
                                        if(window.confirm("<%=LitMsgGenVenFechaConfirm%>")) continuarf=0;
                                        else continuarf=1;
                                    }
                                    if (gen_vencimiento=='-1' && (parent.pantalla.document.facturas_pro.total_factura.value!=parent.pantalla.document.facturas_pro.importe_ant.value))
                                        continuari=0;
                                    else continuari=1;                                
                                    parent.pantalla.document.facturas_pro.action="facturas_pro.asp?mode=save&continuar=" + continuar + "&continuarf=" + continuarf +
                                    "&continuari=" + continuari + "&recalcular_importes=" + recalcular_importes;
                                    parent.pantalla.document.facturas_pro.submit();
                                    document.location="facturas_pro_bt.asp?mode=browse";
                                }
                            }
                        }
                        break;

                    case "cancel": //Cancelar edición
                        parent.pantalla.document.facturas_pro.action="facturas_pro.asp?nfactura=" + parent.pantalla.document.facturas_pro.nfactura.value +
                        "&mode=browse";
                        parent.pantalla.document.facturas_pro.submit();
                        document.location="facturas_pro_bt.asp?mode=browse";
                        break;
                }
                break;

            case "add":
                switch (pulsado) {
                    case "save": //Guardar registro
                        if (ValidarCampos(mode)) {
                            parent.pantalla.document.facturas_pro.action="facturas_pro.asp?mode=first_save";
                            parent.pantalla.document.facturas_pro.submit();
                            document.location="facturas_pro_bt.asp?mode=browse";
                        }
                        break;

                    case "cancel": //Cancelar edición				   
                        parent.pantalla.document.facturas_pro.nproveedor.value="";
                        parent.pantalla.document.facturas_pro.serie.value="";
                        parent.pantalla.document.facturas_pro.nfactura.value="";
                        parent.pantalla.document.facturas_pro.h_nfactura.value="";
                        parent.pantalla.document.facturas_pro.serie.value="";
                        parent.pantalla.document.facturas_pro.action="facturas_pro.asp?mode=add";
                        parent.pantalla.document.facturas_pro.submit();
                        document.location="facturas_pro_bt.asp?mode=add";
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
dim bloqContab
bloqContab = "0"
obtenerparametros("facturas_pro_det")%>

<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(null_s(mode))%>" />
<input type="hidden" name="bloqContab" value="<%=bloqContab%>" />
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
			    if cstr(oborrar) = "0" then %>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeftRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>				
				<%end if
			elseif mode="search" then%>
                <td id="Td1" class="CELDABOT" onclick="javascript:Accion('browse','add');">
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
				<td id="Td2" class="CELDABOT" onclick="javascript:Accion('add','save');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				</td>
				<td id="Td3" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					<%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%end if%>
		</tr>
	</table>
    </div>
    
    <div id="FILTERS_MASTER_ASP">
	    <select class="IN_S" name="campos">
			    <option value="nfactura_pro"><%=LitFactura%></option>
			    <option value="razon_social"><%=LitProveedor%></option>
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