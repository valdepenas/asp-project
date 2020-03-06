<%@ Language=VBScript %>
<%
' VGR 7/3/03 : Listar clientes cuando viene de la Barra de Opcione de Agente o Comercial
' RGU 8/10/2007: Si el parametro pagsl=1 no se pueden editar los datos del cliente
%>
<!--<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"/>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<style>
    .NEGRO {font-Family: Tahoma ;font-size: 10.0px;color: #999999;text-align: center;}
	.PIE {font-Family: Tahoma;font-size: 10.0px;color: #999999;text-align: center;}
</style>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../mensajes.inc" -->

<%
'FLM:20/01/2009:Añadir filtro para módulo ORCU.
si_tiene_modulo_OrCU=ModuloContratado(session("ncliente"),ModOrCU)
''MPC 12/02/2009 Se añade el filtro para el módulo de EBESA
si_tiene_modulo_EBESA = ModuloContratado(session("ncliente"),ModEBESA)
''MPC 12/02/2009 Se añade el filtro para el módulo de Gasoleo B
si_tiene_modulo_TGB=ModuloContratado(session("ncliente"),ModTGB)

'' MPC 08/10/2008 Se obtiene el parámetro cifrepe para controlar si se puede insertar cif repetidos
dim cifrepe
'' MPC 11/05/2009 Se obtiene el parámetro coblg para obligar o no el campo teléfono
dim coblg
ObtenerParametros("clientes")
campos_obligatorios = split(coblg, ",")
'dgb  07/04/2008  anyadimos opcion de idioma para Portugues
Dim pais_idioma, idioma
'Detectamos el pais del usuario
pais_idioma = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")

'JCI - 02/02/2010 Sólo se carga el clientes.inc que es el que gestiona el idioma
%>
<!--#include file="clientes.inc" -->
<%

''EBF Se añade para que por medio de un parametro de usuario no haya que introducir cif. El parametro es &nocif=0
dim nocif, pagsl, emailUsuario, fpUsuario
nocif=limpiaCadena(request.queryString("nocif"))
if(nocif="") then nocif=request.form("nocif")
ObtenerParametros("clientes")
if request.QueryString("viene")>"" then
	viene= limpiaCadena(request.QueryString("viene"))
else
	viene=request.form("viene")
end if%>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  
</head>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/iban.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script type="text/javascript" language="javascript">
    window.addEventListener("load", function hideFilterMaster() {
        var classNameLeft = parent.parent.document.getElementById("left").className;
        if (classNameLeft == "search-panel open") {
            document.getElementById("FILTERS_MASTER_ASP").setAttribute("style", "display:none");
        }
    })


    var da = (document.all) ? 1 : 0;
    var pr = (window.print) ? 1 : 0;
    var mac = (navigator.userAgent.indexOf("Mac") != -1);
    var myArray = new Array();

    function Imprimir() {
        if (pr) //NS4, IE5
            parent.pantalla.print()
        //vbImprimir()
        else if (da && !mac) // IE4 (Windows)
            //vbImprimir()
            alert("<%=LitNoImprime%>");
        else // Otros Navegadores
            alert("<%=LitNoImprime%>");
        return false;
    }

    //H: check exists CIF -----------------------------------------
    //---------Begin---------
    var xmlHttp, ServerResponse = null;
    function ComprobarExisteCIF() {
        var ncliente = "";
        if (parent.pantalla.document.clientes.hncliente != null) ncliente = parent.pantalla.document.clientes.hncliente.value;

        //alert("vamos allá-10");
        xmlHttp = GetXmlHttpObject();
        if (xmlHttp != null) {
            var url = "existeCIFCliente.asp?cif=" + parent.pantalla.document.clientes.cif.value + "&mode=<%=enc.EncodeForJavascript(Request.QueryString("mode"))%>&ncliente=" + ncliente;
            //xmlHttp.onreadystatechange = getData; //asynchronous method
            //xmlHttp.open("GET",url,true); //asynchronous method
            xmlHttp.open("GET", url, false);  //synchronous method
            xmlHttp.send(null);
            //if (ServerResponse!=null) return ServerResponse; //asynchronous method
            return xmlHttp.responseText; //synchronous method
        }
    }
    function getData() {
        if (xmlHttp.readyState == 4 || xmlHttp.readyState == "complete") {
            ServerResponse = xmlHttp.responseText;
        }
    }

    function GetXmlHttpObject() {
        var xmlHttp = null;
        try {
            // Firefox, Opera 8.0+, Safari 
            xmlHttp = new XMLHttpRequest();
        }
        catch (e) {
            //Internet Explorer 
            try {
                xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
            }
            catch (e) {
                xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
        }
        return xmlHttp;
    }

    //----------End----------

    /*
    function ComprobarExisteCIF()
    {
        // Use the native cross-browser nitobi Ajax object
        var myAjaxRequest = new nitobi.ajax.HttpRequest();
        var ncliente="";
    
        if (parent.pantalla.document.clientes.hncliente != null) ncliente = parent.pantalla.document.clientes.hncliente.value;
    
        // Define the url for your generatekey script
        myAjaxRequest.handler = "existeCIFCliente.asp?cif=" + parent.pantalla.document.clientes.cif.value + "&mode=<%=Request.QueryString("mode")%>&ncliente=" + ncliente;
        myAjaxRequest.async = false;
        myAjaxRequest.get();
    
        // return the result to the grid
        return myAjaxRequest.httpObj.responseText;
    }
    */

    function SelecCampoAdi(num) {
        indice = document.opciones.campos.selectedIndex
        if (indice > num) {
            ind = indice - 13;
            if (myArray[ind] == 2) {//se trata de un campo de tipo checkbox
                document.opciones.texto.value = "1";
                //document.opciones.texto.disabled=true; 
                document.opciones.criterio[2].selected = true;
            }
            else {
                document.opciones.texto.value = "";
                //document.opciones.texto.disabled=false; 
                document.opciones.criterio[0].selected = true;
            }
        }
        else {
            document.opciones.texto.value = "";
            //document.opciones.texto.disabled=false; 
            document.opciones.criterio[0].selected = true;
        }
    }

    //DGB: change to page search
    function Buscar() {
        indice = document.opciones.campos.selectedIndex;

        if (parent.pantalla.document.clientes.viene.value == "agentes" || parent.pantalla.document.clientes.viene.value == "comercial") {
            if (indice < 14) {
                parent.pantalla.document.clientes.action = "clientes.asp?mode=search&campo=" + document.opciones.campos.value +
                    "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value +
                    "&viene=" + parent.pantalla.document.clientes.viene.value;
            }
            else {
                ind = indice - 14;
                parent.pantalla.document.clientes.action = "clientes.asp?mode=search&campo=" + document.opciones.campos.value +
                    "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value +
                    "&viene=" + parent.pantalla.document.clientes.viene.value + "&nt=" + myArray[ind];
            }
            parent.pantalla.document.clientes.submit();
            document.location = "clientes_bt.asp?mode=search&viene=" + document.opciones.viene.value;

        }
        else {
            if (indice < 14) {
                SearchPage("client_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
                    "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value +
                    "&viene=" + parent.pantalla.document.clientes.viene.value, 1);
            }
            else {
                ind = indice - 14;
                SearchPage("client_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
                    "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value +
                    "&viene=" + parent.pantalla.document.clientes.viene.value + "&nt=" + myArray[ind], 1);
            }

            document.opciones.texto.value = "";
        }
    }

    //Validación de campos numéricos y fechas.
    function ValidarCampos(mode) {
        //debugger;
        if (parent.pantalla.document.clientes.rsocial.value == "") {
            window.alert("<%=LitMsgRsocialNoNulo%>");
            return false;
        }
    <%if nocif <> "0" then %>
	if (parent.pantalla.document.clientes.cif.value == "") {
            window.alert("<%=LitMsgCIFNoNulo%>");
            return false;
        }
        if (parent.pantalla.document.clientes.pais.value == "España" && (parent.pantalla.document.clientes.validarhacienda.value == "Verdadero" || parent.pantalla.document.clientes.validarhacienda.value == "True" || parent.pantalla.document.clientes.validarhacienda.value == "On")) {
            aux = RespuestaDelWS(parent.pantalla.document.clientes.cif.value, parent.pantalla.document.clientes.rsocial.value);
            if (aux == "Incorrecto") {
                window.alert("<%=LITDATOSCIFINVALIDO%>");
                return false;
            }
            else if (aux.substring(0, 8) != "Correcto") {
                if (aux == "NeedUpdate") {
                    window.alert("<%=LITCERTIFICADONEEDUPDATE%>");
                }
                else if (aux == "NoCertificado") {
                    window.alert("<%=LITCERTIFICADONOTFOUND%>");
                }
                else {
                    window.alert("ValidacionHacienda: " + aux);
                }
                return false;
            }
            else if (aux != "Correcto" + parent.pantalla.document.clientes.rsocial.value) {
                window.alert("<%=LITNOMBREDIFERENTE%>");
                parent.pantalla.document.clientes.rsocial.value = aux.substring(8, aux.length);
            }
        }
    <% end if%>
	<%
            ' Se comprueban el email y la forma de pago cuando tenga los parámetros de usuario de &email y &fp
            %>
    <%if emailUsuario & "" > "" then %>
	if (parent.pantalla.document.clientes.email.value == "") {
            window.alert("<%=LIT_MSG_EMAIL_NONULO%>");
            return false;
        }
        else if (!checkmail(parent.pantalla.document.clientes.email)) {
            return false;
        }
    <% else%>
    if (parent.pantalla.document.clientes.email.value != "") {
            if (!checkmail(parent.pantalla.document.clientes.email)) {
                return false;
            }
        }
    <% end if%>
	<%if fpUsuario & "" > "" then %>
	if (parent.pantalla.document.clientes.fpago.value == "") {
            window.alert("<%=LIT_MSG_FORMAPAGO_NONULO%>");
            return false;
        }
	<% end if%>
    
	if (parent.pantalla.document.clientes.validacionCliente.value == "True" || parent.pantalla.document.clientes.validacionCliente.value == "Verdadero" || parent.pantalla.document.clientes.validacionCliente.value == "On") {
                if (parent.pantalla.valida_nif_cif_nie() <= 0) {
                    if (window.confirm("<%=LIT_MSG_CIFNIF_NOVALIDO%>") == false)
                        return false;
                }
            }

        if (parent.pantalla.document.clientes.domicilio.value == "") {
            window.alert("<%=LitMsgDireccionNoNulo%>");
            return false;
        }
	<%for i = 0 to ubound(campos_obligatorios) %>
	    <%if instr(campos_obligatorios(i), "campo") > 0 then
        numero_campo = cint(mid(campos_obligatorios(i), len("campo") + 1, len("campo") + 2)) %>
	if (parent.pantalla.document.clientes.campo <%= numero_campo %>.value == "")
        {
            alert("<%=LitElCampo%> " + parent.pantalla.document.clientes.titulo_campo <%=numero_campo %>.value + " <%=LitNoPuedeSerNulo%>");
            return false;
        }
	    <%else%>
	if (parent.pantalla.document.clientes.<%= campos_obligatorios(i) %>.value == "")
        {
            alert("<%=LitElCampo%> <%=campos_obligatorios(i)%> <%=LitNoPuedeSerNulo%>");
            return false;
        }
	    <% end if%>
	<% next %>
	if (parent.pantalla.document.clientes.falta.value == "") {
                window.alert("<%=LitMsgFechaAltaNoNulo%>");
                return false;
            }
            else {
                if (!checkdate(parent.pantalla.document.clientes.falta)) {
                    window.alert("<%=LitMsgFechaAltaFecha%>");
                    return false;
                }
            }
        if (parent.pantalla.document.clientes.fbaja.value != "") {
            if (!checkdate(parent.pantalla.document.clientes.fbaja)) {
                window.alert("<%=LitMsgFechaBajaFecha%>");
                return false;
            }
        }
        if (parent.pantalla.document.clientes.divisa.value == "") {
            window.alert("<%=LitMsgDivisaNoNulo%>");
            return false;
        }
        if (parent.pantalla.document.clientes.de_domicilio.value != null) {
            if (parent.pantalla.document.clientes.de_domicilio.value == "") {
                if (parent.pantalla.document.clientes.de_poblacion.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_cp.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_provincia.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_pais.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_telefono.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
            }
        }

        if (parent.pantalla.document.clientes.de_domicilioF.value != null) {
            if (parent.pantalla.document.clientes.de_domicilioF.value == "") {
                if (parent.pantalla.document.clientes.de_poblacionF.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_cpF.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_provinciaF.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_paisF.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
                if (parent.pantalla.document.clientes.de_telefonoF.value > "") {
                    window.alert("<%=LitMsgDomDENoNulo%>");
                    return false;
                }
            }
        }

        // >>> MCA 22/09/05 : Admitir decimales en los descuentos.
        if (isNaN(parent.pantalla.document.clientes.descuento.value.replace(",", "."))) {
            window.alert("<%=LitMsgDescuentoNumerico%>");
            return false;
        }
        //parent.pantalla.document.clientes.descuento.value= parent.pantalla.document.clientes.descuento.value.replace(".",",");

        if (isNaN(parent.pantalla.document.clientes.descuento2.value.replace(",", "."))) {
            window.alert("<%=LitMsgDescuento2Numerico%>");
            return false;
        }
        //parent.pantalla.document.clientes.descuento2.value= parent.pantalla.document.clientes.descuento2.value.replace(".",",");

        if (isNaN(parent.pantalla.document.clientes.descuento3.value.replace(",", "."))) {
            window.alert("<%=LitMsgDescuento3Numerico%>");
            return false;
        }
        //parent.pantalla.document.clientes.descuento3.value= parent.pantalla.document.clientes.descuento3.value.replace(".",",");

        if (isNaN(parent.pantalla.document.clientes.descuentoLineal.value.replace(",", "."))) {
            window.alert("<%=LitMsgDescuentoLinealNumerico%>");
            return false;
        }
        //parent.pantalla.document.clientes.descuentoLineal.value= parent.pantalla.document.clientes.descuentoLineal.value.replace(".",",");

        if (isNaN(parent.pantalla.document.clientes.recargo.value)) {
            window.alert("<%=LitMsgRFNumerico%>");
            return false;
        }

        if (isNaN(parent.pantalla.document.clientes.e_primer_ven.value)) {
            window.alert("<%=LitPrimerVenNoNum%>");
            parent.pantalla.document.clientes.e_primer_ven.focus();
            return false;
        }
        else if (isNaN(parent.pantalla.document.clientes.e_segundo_ven.value)) {
            window.alert("<%=LitSegundoVenNoNum%>");
            parent.pantalla.document.clientes.e_segundo_ven.focus();
            return false;
        }
        else if (isNaN(parent.pantalla.document.clientes.e_tercer_ven.value)) {
            window.alert("<%=LitTercerVenNoNum%>");
            parent.pantalla.document.clientes.e_tercer_ven.focus();
            return false;
        }
        mode = "<%=enc.EncodeForJavascript(Request.QueryString("mode"))%>";
        ver = "<%=enc.EncodeForJavascript(limpiaCadena(Request.QueryString("ver")))%>";
        if (ver == "0")
            mode = "add"
        if (mode != "add") {
            if (parent.pantalla.document.clientes.ndist == null) { }
            else
                if (parent.pantalla.document.clientes.ndist.value == "") {
                }
                else {
                    if (isNaN(parent.pantalla.document.clientes.vencom1.value.replace(",", "."))) {
                        window.alert("<%=LitMsgLV1Numerico%>");
                        return false;
                    }
                    //parent.pantalla.document.clientes.vencom1.value=parent.pantalla.document.clientes.vencom1.value.replace(".",",");
                    if (isNaN(parent.pantalla.document.clientes.vencom2.value.replace(",", "."))) {
                        window.alert("<%=LitMsgLV2Numerico%>");
                        return false;
                    }
                    //parent.pantalla.document.clientes.vencom2.value=parent.pantalla.document.clientes.vencom2.value.replace(".",",");
                    if (isNaN(parent.pantalla.document.clientes.pcom1.value.replace(",", "."))) {
                        window.alert("<%=LitMsgPorComisionNumerico%>");
                        return false;
                    }
                    //parent.pantalla.document.clientes.pcom1.value=parent.pantalla.document.clientes.pcom1.value.replace(".",",");
                    if (isNaN(parent.pantalla.document.clientes.pcom2.value.replace(",", "."))) {
                        window.alert("<%=LitMsgPorComision2Numerico%>");
                        return false;
                    }
                    //parent.pantalla.document.clientes.pcom2.value=parent.pantalla.document.clientes.pcom2.value.replace(".",",");
                }
        }

        if (isNaN(parent.pantalla.document.clientes.pht.value.replace(",", "."))) {
            window.alert("<%=LitMsgPKMNumerico%>");
            return false;
        }

        if (isNaN(parent.pantalla.document.clientes.pkm.value.replace(",", "."))) {
            window.alert("<%=LitMsgPKMNumerico%>");
            return false;
        }

        if (isNaN(parent.pantalla.document.clientes.pd.value.replace(",", "."))) {
            window.alert("<%=LitMsgPDesplzaNumerico%>");
            return false;
        }

        if (parent.pantalla.document.clientes.NEntidad.value != "") {
            if (isNaN(parent.pantalla.document.clientes.NEntidad.value)) {
                window.alert("<%=LitMsgCuentaNumerico%>");
                return false;
            }
        }
        if (parent.pantalla.document.clientes.Oficina.value != "") {
            if (isNaN(parent.pantalla.document.clientes.Oficina.value)) {
                window.alert("<%=LitMsgCuentaNumerico%>");
                return false;
            }
        }
        if (parent.pantalla.document.clientes.DC.value != "") {
            if (isNaN(parent.pantalla.document.clientes.DC.value)) {
                window.alert("<%=LitMsgCuentaNumerico%>");
                return false;
            }
        }
        if (parent.pantalla.document.clientes.Cuenta.value != "") {
            if (isNaN(parent.pantalla.document.clientes.Cuenta.value)) {
                window.alert("<%=LitMsgCuentaNumerico%>");
                return false;
            }
        }

        SumaC = 0;

        while (parent.pantalla.document.clientes.NEntidad.value.search(" ") != -1) {
            parent.pantalla.document.clientes.NEntidad.value = parent.pantalla.document.clientes.NEntidad.value.replace(" ", "");
        }
        while (parent.pantalla.document.clientes.Oficina.value.search(" ") != -1) {
            parent.pantalla.document.clientes.Oficina.value = parent.pantalla.document.clientes.Oficina.value.replace(" ", "");
        }
        while (parent.pantalla.document.clientes.DC.value.search(" ") != -1) {
            parent.pantalla.document.clientes.DC.value = parent.pantalla.document.clientes.DC.value.replace(" ", "");
        }
        while (parent.pantalla.document.clientes.Cuenta.value.search(" ") != -1) {
            parent.pantalla.document.clientes.Cuenta.value = parent.pantalla.document.clientes.Cuenta.value.replace(" ", "");
        }

        if (parent.pantalla.document.clientes.NEntidad.value != "") {
            SumaC++;
        }
        if (parent.pantalla.document.clientes.Oficina.value != "") {
            SumaC++;
        }
        if (parent.pantalla.document.clientes.DC.value != "") {
            SumaC++;
        }
        if (parent.pantalla.document.clientes.Cuenta.value != "") {
            SumaC++;
        }

        if (SumaC != 0 && SumaC != 4) {
            window.alert("<%=LitCuentaBancariaIncorrecta%>");
            return false;
        }

        var country = parent.pantalla.document.clientes.country.value;
        var i_iban = parent.pantalla.document.clientes.iban.value;
        if (country == "") {
            country = "ES";
            //parent.pantalla.document.clientes.country.value = "ES";
        }
        var account = parent.pantalla.document.clientes.NEntidad.value + parent.pantalla.document.clientes.Oficina.value + parent.pantalla.document.clientes.DC.value + parent.pantalla.document.clientes.Cuenta.value;

        if (account.toString().length != 0 && account.toString().length < 11) {
            alert("<%=LitCodeIBANIncorrect%>");
            return false;
        }

        if (country == "ES" && account.toString().length != 0) {

            var iban = CreateIBAN(country, account);

            if (parent.pantalla.document.clientes.iban.value != "" && account.toString().length == 20) {
                var i_iban = parent.pantalla.document.clientes.iban.value;

                if (!ValidateIBAN(iban, i_iban)) {
                    alert("<%=LitCodeIBANIncorrect%>");
                    parent.pantalla.document.clientes.iban.select();
                    parent.pantalla.document.clientes.iban.focus();
                    return false;
                }
            }
            else {
                window.alert("<%=LitCodeIBANIncorrect%>");
                parent.pantalla.document.clientes.country.value = "ES";
                parent.pantalla.document.clientes.iban.value = iban;
                return false;
            }

        }
    <%if si_tiene_modulo_OrCU<>0 or si_tiene_modulo_TGB<>0 then%>
    if (parent.pantalla.document.clientes.ibanGB != null && account.toString().length == 20) {
            var countryGB = parent.pantalla.document.clientes.countryGB.value;
            if (countryGB == "") countryGB = "ES";
            if (countryGB == "ES") {
                var accountGB = parent.pantalla.document.clientes.NEntidadGB.value + parent.pantalla.document.clientes.OficinaGB.value + parent.pantalla.document.clientes.DCGB.value + parent.pantalla.document.clientes.CuentaGB.value;
                var ibanGB = CreateIBAN(countryGB, accountGB);
                if (parent.pantalla.document.clientes.ibanGB.value != "") {
                    var i_ibanGB = parent.pantalla.document.clientes.ibanGB.value;

                    if (!ValidateIBAN(ibanGB, i_ibanGB)) {
                        alert("<%=LitCodeIBANIncorrect%>");
                        parent.pantalla.document.clientes.ibanGB.select();
                        parent.pantalla.document.clientes.ibanGB.focus();
                        return false;
                    }
                }
                else parent.pantalla.document.clientes.ibanGB.value = ibanGB;
            }
        }
    <%end if%>

	//FLM:20/01/2009: Añadir otros campos para módulo Orcu
    <%if si_tiene_modulo_OrCU<>0 then %>
    
                /*if (parent.pantalla.document.clientes.NEntidad2.value!=""){
                    if (isNaN(parent.pantalla.document.clientes.NEntidad2.value)) {
                        window.alert("<%=LitMsgCuentaNumerico%>");
                        return false;
                    }
                }
                if (parent.pantalla.document.clientes.Oficina2.value!=""){
                    if (isNaN(parent.pantalla.document.clientes.Oficina2.value)) {
                        window.alert("<%=LitMsgCuentaNumerico%>");
                        return false;
                    }
                }
                if (parent.pantalla.document.clientes.DC2.value!=""){
                    if (isNaN(parent.pantalla.document.clientes.DC2.value)) {
                        window.alert("<%=LitMsgCuentaNumerico%>");
                        return false;
                    }
                }
                if (parent.pantalla.document.clientes.Cuenta2.value!=""){
                    if (isNaN(parent.pantalla.document.clientes.Cuenta2.value)) {
                        window.alert("<%=LitMsgCuentaNumerico%>");
                        return false;
                    }
                }
            
                SumaC=0;
            
                while (parent.pantalla.document.clientes.NEntidad2.value.search(" ")!=-1){
                    parent.pantalla.document.clientes.NEntidad2.value=parent.pantalla.document.clientes.NEntidad2.value.replace(" ","");
                }
                while (parent.pantalla.document.clientes.Oficina2.value.search(" ")!=-1){
                    parent.pantalla.document.clientes.Oficina.value=parent.pantalla.document.clientes.Oficina2.value.replace(" ","");
                }
                while (parent.pantalla.document.clientes.DC2.value.search(" ")!=-1){
                    parent.pantalla.document.clientes.DC2.value=parent.pantalla.document.clientes.DC2.value.replace(" ","");
                }
                while (parent.pantalla.document.clientes.Cuenta2.value.search(" ")!=-1){
                    parent.pantalla.document.clientes.Cuenta2.value=parent.pantalla.document.clientes.Cuenta2.value.replace(" ","");
                }
            
                if (parent.pantalla.document.clientes.NEntidad2.value!=""){
                    SumaC++;
                }
                if (parent.pantalla.document.clientes.Oficina2.value!=""){
                    SumaC++;
                }
                if (parent.pantalla.document.clientes.DC2.value!=""){
                    SumaC++;
                }
                if (parent.pantalla.document.clientes.Cuenta2.value!=""){
                    SumaC++;
                }
            
                if (SumaC!=0 && SumaC!=4){
                    window.alert("<%=LitCuentaBancariaIncorrecta%>");
                    return false;
                }*/

                mmaa="1/" + parent.pantalla.document.clientes.fcaducidad2.value.substring(0, 2) + "/" + parent.pantalla.document.clientes.fcaducidad2.value.substring(2, 4)
        if (parent.pantalla.document.clientes.fcaducidad2.value > "" && (isNaN(parent.pantalla.document.clientes.fcaducidad2.value.replace(".", ",")) || parent.pantalla.document.clientes.fcaducidad2.value.replace(" ", "").length != 4 || !chkdatetime(mmaa))) {
            window.alert("<%=LitMsgFCaducidad%>");
            parent.pantalla.document.clientes.fcaducidad2.focus();
            parent.pantalla.document.clientes.fcaducidad2.select();
            return false;
        }

        //ricardo 23-07-2013 si se modifica el campo saldomax, el usuario debera estar dado de alta en personal
        hd_saldoMax = 0;
        saldomax_act = 0;
        try {
            if (parent.pantalla.document.clientes.saldoMaxOld != null) {
                hd_saldoMax = parseFloat(parent.pantalla.document.clientes.saldoMaxOld.value.replace(",", "."));
                if (isNaN(hd_saldoMax)) {
                    hd_saldoMax = 0;
                }
            }
        }
        catch (e) {
        }
        try {
            if (parent.pantalla.document.clientes.saldomax != null) {
                saldomax_act = parseFloat(parent.pantalla.document.clientes.saldomax.value.replace(",", "."));
                if (isNaN(saldomax_act)) {
                    saldomax_act = 0;
                }
            }
        }
        catch (e) {
        }
        UsuarioEnPersonal = "";
        try {
            if (parent.pantalla.document.clientes.UsuarioEnPersonal != null) {
                UsuarioEnPersonal = parent.pantalla.document.clientes.UsuarioEnPersonal.value;
            }
        }
        catch (e) {
        }
        //window.alert("los datos son-" + hd_saldoMax + "-" + saldomax_act + "-" + UsuarioEnPersonal + "-");
        if (hd_saldoMax != saldomax_act) {
            if (mode != "add") {
                if (UsuarioEnPersonal == "") {
                    window.alert("<%=LITCHANGBALSTAFF%>");
                    return false;
                }
            }
        }
    //fin ricardo 23-07-2013
    <%end if %>

	if (parent.pantalla.document.clientes.fnacimiento != null) {
                if (parent.pantalla.document.clientes.fnacimiento != "") {
                    if (!checkdate(parent.pantalla.document.clientes.fnacimiento)) {
                        window.alert("<%=LitMsgFechaNacimientoFecha%>");
                        return false;
                    }
                }
            }
        /*
        if (isNaN(parent.pantalla.document.clientes.honorario_lab.value)) {
            window.alert("El honorario laboral tiene que ser un dato numérico.");
            return false;
        }
    
        if (Date.parse(parent.pantalla.document.clientes.alta.value)==NaN) {
            window.alert("Fecha de alta incorrecta.");
            return false;
        }
    
        */

        while (parent.pantalla.document.clientes.rgomaxaut.value.search(" ") != -1) {
            parent.pantalla.document.clientes.rgomaxaut.value = parent.pantalla.document.clientes.rgomaxaut.value.replace(" ", "");
        }
        if (isNaN(parent.pantalla.document.clientes.rgomaxaut.value.replace(",", "."))) {
            window.alert("<%=LitMsgRiesgoMaxNumerico%>");
            parent.pantalla.document.clientes.rgomaxaut.value = 0;
            return false;
        }
        if (parseFloat(parent.pantalla.document.clientes.rgomaxaut.value) < 0) {
            window.alert("<%=LitMsgRiesgoMaxNoNegativo%>");
            parent.pantalla.document.clientes.rgomaxaut.value = 0;
            return false;
        }
        //parent.pantalla.document.clientes.rgomaxaut.value=parent.pantalla.document.clientes.rgomaxaut.value.replace(".",",");

        /***RGU 27/6/06**/
        mmaa = "1/" + parent.pantalla.document.clientes.fcaducidad.value.substring(0, 2) + "/" + parent.pantalla.document.clientes.fcaducidad.value.substring(2, 4)
        if (parent.pantalla.document.clientes.fcaducidad.value > "" && (isNaN(parent.pantalla.document.clientes.fcaducidad.value.replace(".", ",")) || parent.pantalla.document.clientes.fcaducidad.value.replace(" ", "").length != 4 || !chkdatetime(mmaa))) {
            window.alert("<%=LitMsgFCaducidad%>");
            parent.pantalla.document.clientes.fcaducidad.focus();
            parent.pantalla.document.clientes.fcaducidad.select();
            return false;
        }
	/**RGU**/

    <%if si_tiene_modulo_EBESA = 0 then%>
	if (mode != "add") {
            rgomaxaut = parseFloat(parent.pantalla.document.clientes.rgomaxaut.value.replace(",", "."));
            rgomaxaut_ant = parseFloat(parent.pantalla.document.clientes.rgomaxaut_ant.value.replace(",", "."));

            if (parent.pantalla.document.clientes.rgomaxaut.value != "" && parent.pantalla.document.clientes.rgomaxaut.value.replace(",", ".") > 0 && rgomaxaut_ant == 0) {
                //if (rgomaxaut!=rgomaxaut_ant){
                if (window.confirm("<%=LitRiesgoDistCero%>") == true) parent.pantalla.document.clientes.rcalc.value = 1;
                else {
                    parent.pantalla.document.clientes.rgomaxaut.value = 0;
                    return false;
                }
            }
        }
	<%end if%>

            mode=mode.toUpperCase();
        mode = mode.toLowerCase();
        if (mode != "add") {
            if (parent.pantalla.marcoFoto != null)
                //window.alert(parent.pantalla.marcoFoto.document.upload.blob.value)
                if (parent.pantalla.marcoFoto.document.upload.blob.value != "") {
                    todocorrecto = 1;
                    if (navigator.appName == "Microsoft Internet Explorer") {
                        var fso = new ActiveXObject("Scripting.FileSystemObject");
                        if (!fso.FileExists(parent.pantalla.marcoFoto.document.upload.blob.value)) {
                            window.alert("<%=LITFICHERONOEXISTE%>");
                            todocorrecto = 0;
                        }
                        if (todocorrecto == 1) {
                            if (fso.GetFile(parent.pantalla.marcoFoto.document.upload.blob.value).Size ><%=cLng(uploadFileLimit) %>)
                            {
                                window.alert("<%=LitFicTamGrande%> <%=formatnumber(((uploadFileLimit-390)/1024),0,-1,0,-1)%><%=LitFicTamGrande2%>");
                                todocorrecto = 0;
                            }
                        }
                    }
                    else {
                        //window.alert("Entra aqui")
                        if (parent.pantalla.marcoFoto.document.upload.blob.files[0].size ><%=cLng(uploadFileLimit) %>)
                        {
                            alert("<%=LitFicTamGrande%> <%=formatnumber(((uploadFileLimit-390)/1024),0,-1,0,-1)%><%=LitFicTamGrande2%>");
                            todocorrecto = 0;
                        }
                    }
                    if (todocorrecto == 0) {
                        //window.alert("Entra aqui 2")
                        return false;
                    }

            }
            //window.alert("todocorrecto = " & todocorrecto)
        }

        if ((mode == "add" || mode == "edit") && (parent.pantalla.document.clientes.si_campo_personalizables.value == 1)) {
            num_campos = parent.pantalla.document.clientes.num_campos.value;

            respuesta = comprobarCampPerso("parent.pantalla.", num_campos, "clientes");
            if (respuesta != 0) {
                titulo = "titulo_campo" + respuesta;
                tipo = "tipo_campo" + respuesta;
                titulo = parent.pantalla.document.clientes.elements[titulo].value;
                tipo = parent.pantalla.document.clientes.elements[tipo].value;
                if (tipo == 4) nomTipo = "<%=LitTipoNumericoCli%>";
                else if (tipo == 5) {
                    nomTipo = "<%=LitTipoFechaCli%>";
                }

                alert("<%=LitMsgCampoCli%> " + titulo + " <%=LitMsgTipoCli%> " + nomTipo);
                return false;
            }
        }
        return true;
    }

    var page_compncom = 0;
    function comprobar_ncom(mode) {
        //window.alert(page_compncom + "-" + mode);
        if (page_compncom == 0) {
            window.setTimeout("comprobar_ncom('" + mode + "');", 500);
        }
        else {
            continuar = 1;
            if (parent.pantalla.document.clientes.comp_ncom.value == 1) {
                if (window.confirm("<%=LitNComerExis%>") == false) {
                    continuar = 0;
                    parent.pantalla.document.clientes.comp_ncom.value = 0;
                    page_compncom = 0;
                }
            }
            if (continuar == 1) {
                var cade = "";
                //mode=parent.pantalla.document.clientes.mode_accesos_tienda.value;
                if (mode == "edit") {
                    cade = "save&ncliente=" + parent.pantalla.document.clientes.hncliente.value;
                }
                if (mode == "add") {
                    cade = "save";
                }
                repe = 0;
            <%if cifrepe = "1" then%>
                    existe = ComprobarExisteCIF();
                if (existe != "") {
                    if (confirm("<%=LitCifExisteContinuar%>")) repe = 1;
                }
            <%end if%>
                parent.pantalla.document.clientes.action="clientes.asp?mode=" + cade + "&repe=" + repe;
			<%if cifrepe = "1" then%>
			if (repe == 1 || existe == "") {
                    parent.pantalla.document.clientes.submit();
                    document.location = "clientes_bt.asp?mode=browse";
                }
			<%else%>
                    parent.pantalla.document.clientes.submit();
            document.location = "clientes_bt.asp?mode=browse";
			<%end if%>
		}
        }
    }

    function esPais(valor) {
        if (valor.toUpperCase() == "ESPAÑA" || valor == "")
            return true;

        return false;
    }

    // ASP 09/1/2012
    function reloadPanelGlobal(viene) {
        if (viene == "GlobalAgenda") {
            parent.window.opener.__doPostBack("reload", "");
        }
    }
    //FIN ASP 09/1/2012

    //Realizar la acción correspondiente al botón pulsado.
    function Accion(mode, pulsado) {
        switch (mode) {
            case "browse":
                switch (pulsado) {
                    case "add": //Nuevo registro
                        parent.pantalla.document.clientes.action = "clientes.asp?mode=" + pulsado;
                        parent.pantalla.document.clientes.submit();
                        if (parent.pantalla.document.clientes.viene.value == "GlobalAgenda") {
                            document.location = "clientes_bt.asp?mode=" + pulsado + "&viene=GlobalAgenda";
                        }
                        else {
                            document.location = "clientes_bt.asp?mode=" + pulsado;
                        }
                        break;

                    case "edit": //Editar registro
                        parent.pantalla.document.clientes.action = "clientes.asp?ncliente=" + parent.pantalla.document.clientes.hncliente.value +
                            "&mode=" + pulsado;
                        parent.pantalla.document.clientes.submit();
                        document.location = "clientes_bt.asp?mode=" + pulsado + '&viene=<%=enc.EncodeForJavascript(viene)%>';
                        break;

                    case "delete": //Eliminar registro
                        if (window.confirm("<%=LitMsgEliminarClienteConfirm%>") == true) {
                            parent.pantalla.document.clientes.action = "clientes.asp?mode=" + pulsado + "&ncliente=" + parent.pantalla.document.clientes.hncliente.value;
                            parent.pantalla.document.clientes.submit();
                            document.location = "clientes_bt.asp?mode=browse";
                        }
                        break;
                    case "print": //Imprimir ficha
                        parent.pantalla.focus();
                        Imprimir();
                        break;

                    case "search": //Buscar datos
                        break;
                }
                break;

            case "edit":
                switch (pulsado) {
                    case "save": //Guardar registro
                        if (ValidarCampos(mode)) {
					        //ActiveDDL();
					        <%if viene= "Facturas_cli_E" then%>
                                parent.pantalla.GuardarFactura();
					        <%end if%>
                                //ricardo 19-10-2007 si la empresa es ebesa no se comprobara si el nombre comercial existe
                                empresa_ebesa="00010";
                            empresa_session = "<%=session("ncliente")%>";
                            if (empresa_session != empresa_ebesa) {
				    		    /*if((parent.pantalla.document.clientes.poblacion.value != "" || parent.pantalla.document.clientes.provincia.value != "")&& parent.pantalla.document.clientes.codPoblacion.value == "")
					           {
					             parent.pantalla.alcerrarModal("#SELECCIONAR_POBLACION2","1","0");
					           }
					           else
					           {*/
                                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                                parent.pantalla.fr_CompNcom.document.location = "clientes_ncomercial.asp?nc=" + parent.pantalla.document.clientes.hncliente.value + "&mode=comprobar" + "&ncom=" + escape(parent.pantalla.document.clientes.ncomercial.value);
                                page_compncom = 0;
                                //ASP 09/01/2012                            
                                reloadPanelGlobal("<%=enc.EncodeForJavascript(viene)%>");
                                //FIN ASP 09/01/2012
                                if (page_compncom == 0) {
                                    comprobar_ncom(mode);
                                }
                                //}
                            }
                            else {
                                repe = 0;
                                <%if cifrepe = "1" then%>
                                    existe =ComprobarExisteCIF();
                                if (existe != "") {
                                    if (confirm("<%=LitCifExisteContinuar%>")) repe = 1;
                                }
                                <%end if%>
				                
				                
                                    /*  if(esPais(parent.pantalla.document.clientes.pais.value) && (parent.pantalla.clientes.poblacion.value != "" || parent.pantalla.clientes.provincia.value != "")&& parent.pantalla.clientes.codPoblacion.value == "")
                                     {
                                       parent.pantalla.alcerrarModal("#SELECCIONAR_POBLACION2",repe,"0");
                                     }
                                     else
                                     {*/
                                    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
                                parent.pantalla.document.clientes.action = "clientes.asp?mode=save&ncliente=" + parent.pantalla.document.clientes.hncliente.value + "&repe=" + repe;
				                <%if cifrepe = "1" then%>
			                    if (repe == 1 || existe == "") {
                                    parent.pantalla.document.clientes.submit();
                                    //ASP 09/01/2012                            
                                    reloadPanelGlobal("<%=enc.EncodeForJavascript(viene)%>");
                                    //FIN ASP 09/01/2012
                                    document.location = "clientes_bt.asp?mode=browse";
                                }
			                    <%else%>
                                    parent.pantalla.document.clientes.submit();
                                //ASP 09/01/2012                            
                                reloadPanelGlobal("<%=enc.EncodeForJavascript(viene)%>");
                                //FIN ASP 09/01/2012
                                document.location = "clientes_bt.asp?mode=browse";
			                    <%end if%>
			                    //}
			                    
		    			    }
                            //			parent.pantalla.document.clientes.action="clientes.asp?mode="+ mode + "&ncliente=" + parent.pantalla.document.clientes.hncliente.value;
                            //			parent.pantalla.document.clientes.submit();
                            //			document.location="clientes_bt.asp?mode=browse";
                            if (parent.pantalla.marcoFoto.document.upload.pepe.value == "1") {
                                parent.pantalla.marcoFoto.document.upload.action = "sube_int_cli.asp";
                                parent.pantalla.marcoFoto.document.upload.submit();
                            }
                            else {
                                todocorrecto = 1;
                                if (parent.pantalla.marcoFoto.document.upload.blob.value != "") {
                                    if (navigator.appName == "Microsoft Internet Explorer") {
                                        var fso = new ActiveXObject("Scripting.FileSystemObject");
                                        if (!fso.FileExists(parent.pantalla.marcoFoto.document.upload.blob.value)) {
                                            window.alert("<%=LITFICHERONOEXISTE%>");
                                            todocorrecto = 0;
                                        }
                                        if (todocorrecto == 1) {
                                            if (fso.GetFile(parent.pantalla.marcoFoto.document.upload.blob.value).Size ><%=cLng(uploadFileLimit) %>)
                                            {
                                                window.alert("<%=LitFicTamGrande%> <%=formatnumber(((uploadFileLimit-390)/1024),0,-1,0,-1)%><%=LitFicTamGrande2%>");
                                                todocorrecto = 0;
                                            }
                                        }
                                    }
                                    else {
                                        if (parent.pantalla.marcoFoto.document.upload.blob.files[0].size ><%=cLng(uploadFileLimit) %>)
                                        {
                                            alert("<%=LitFicTamGrande%> <%=formatnumber(((uploadFileLimit-390)/1024),0,-1,0,-1)%><%=LitFicTamGrande2%>");
                                            todocorrecto = 0;
                                        }
                                    }
                                    if (todocorrecto == 1) {
                                        //window.alert("Entro aqui 3")
                                        parent.pantalla.marcoFoto.document.upload.action = "sube_int_cli.asp";
                                        parent.pantalla.marcoFoto.document.upload.submit();
                                    }
                                }
                                else {
                                }
                            }
                        }
                        break;

                    case "cancel": //Cancelar edición
                        if (document.opciones.viene.value == "") {
                            parent.pantalla.document.clientes.action = "clientes.asp?ncliente=" + parent.pantalla.document.clientes.hncliente.value +
                                "&mode=browse";
                            parent.pantalla.document.clientes.submit();
                            document.location = "clientes_bt.asp?mode=browse";
                        }
                        else parent.window.close();
                        break;

                    case "search": //Buscar datos
                        break;
                }
                break;

            case "add":
                switch (pulsado) {
                    case "save": //Guardar registro
                        if (ValidarCampos(mode)) {
                            // ActiveDDL();
                            //ricardo 19-10-2007 si la empresa es ebesa no se comprobara si el nombre comercial existe
                            empresa_ebesa = "00010";
                            empresa_session = "<%=session("ncliente")%>";
                            if (empresa_session != empresa_ebesa) {
                                /*if((parent.pantalla.document.clientes.poblacion.value != "" || parent.pantalla.document.clientes.provincia.value != "")&& parent.pantalla.document.clientes.codPoblacion.value == "")
                                {
                                  parent.pantalla.alcerrarModal("#SELECCIONAR_POBLACION2","1","1");
                                }
                                else
                                {*/
                                parent.pantalla.fr_CompNcom.document.location = "clientes_ncomercial.asp?nc=&mode=comprobar&ncom=" + escape(parent.pantalla.document.clientes.ncomercial.value);
                                page_compncom = 0;
                                if (page_compncom == 0) {
                                    comprobar_ncom(mode);
                                }
                                // }
                            }
                            else {
                                repe = 0;
                            <%if cifrepe = "1" then%>
                                    existe =ComprobarExisteCIF();
                                if (existe != "") {
                                    if (document.clientes.cif.value != document.clientes.hcif.value) {
                                        if (confirm("<%=LitCifExisteContinuar%>")) repe = 1;
                                    }
                                    else repe = 1;
                                }
                            <%end if%>
                            
                                    /*if((parent.pantalla.clientes.poblacion.value != "" || parent.pantalla.clientes.provincia.value != "")&& parent.pantalla.clientes.codPoblacion.value == "")
                                    {
                                      parent.pantalla.alcerrarModal("#SELECCIONAR_POBLACION2",repe,"1");
                                    }
                                    else
                                    {*/
                                parent.pantalla.document.clientes.action="clientes.asp?mode=save&repe=" + repe;
	    					    <%if cifrepe = "1" then%>
			                    if (repe == 1 || existe == "") {
                                    parent.pantalla.document.clientes.submit();
                                    document.location = "clientes_bt.asp?mode=browse";
                                }
			                    <%else%>
                                    parent.pantalla.document.clientes.submit();
                                document.location = "clientes_bt.asp?mode=browse";
			                    <%end if%>
                            //}
		    			}
                        }
                        break;

                    case "cancel": //Cancelar edición
                        if (document.opciones.viene.value == "") {
                            parent.pantalla.document.clientes.action = "clientes.asp?mode=add";
                            parent.pantalla.document.clientes.submit();
                            document.location = "clientes_bt.asp?mode=add";
                        }
                        else {
                            parent.window.close();
                        }
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
        }
    }

    function ActiveDDL() {

        parent.pantalla.document.clientes.pais.disabled = false;
        parent.pantalla.document.clientes.de_pais.disabled = false;
        parent.pantalla.document.clientes.de_paisF.disabled = false;
        parent.pantalla.document.clientes.provincia.disabled = false;
        parent.pantalla.document.clientes.de_provincia.disabled = false;
        parent.pantalla.document.clientes.de_provinciaF.disabled = false;
    }

    //****************************************************************************************************************
    function comprobar_enter() {
        //si se ha pulsado la tecla enter
        //if (window.event.keyCode==13){
        //document.opciones.criterio.focus();
        Buscar();
        //}
    }
    //document.onkeypress=gestionarTecla()


    function RespuestaDelWS(nif, nombre) {
        xmlHttp = GetXmlHttpObject();
        if (xmlHttp != null) {
            var url = "ValidacionHacienda.asp?nif=" + nif + "&nombre=" + escape(nombre);
            //xmlHttp.onreadystatechange = getData; //asynchronous method
            //xmlHttp.open("GET",url,true); //asynchronous method
            xmlHttp.open("GET", url, false);  //synchronous method
            xmlHttp.send(null);
            //if (ServerResponse!=null) return ServerResponse; //asynchronous method
            return xmlHttp.responseText; //synchronous method
        }
    }
</script>

<body class="body_master_ASP">


<%'mmgClin:
dim bp
dim bpca
dim ver
ObtenerParametros("clientes")
mode=Request.QueryString("mode")
if ver="0" and mode="search" then
    mode="add"
end if

if ver="0" and mode="browse" then
    mode="nbrowse"
end if

if viene="agentes" or viene="comercial" then
	mode="search"
end if

%>

<form name="opciones" method="post" action="">
<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
        <tr>
		    <%if mode="browse" then%>
                <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					<%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				</td>
				<% if pagsl&""<>"1" then%>
				    <td id="ideedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBTLeft LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				    </td>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeft LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				<%end if%>
				<td id="idprint" class="CELDABOT" onclick="javascript:Accion('browse','print');">
					<%PintarBotonBTLeft LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				</td>
			<%elseif mode="nbrowse" then%>
				<td id="idadd" align="center" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					<%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				</td>
			<%elseif mode="search" and viene<>"agentes" and viene<>"comercial"  then%>
				<td id="idadd" align="center" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					<%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				</td>
			<%elseif mode="edit" then
			     if pagsl&""<>"1" then%>
				    <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				 <%end if %>
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
	    <!--<table width="100%" border='0' cellspacing="1" cellpadding="1">-->
        <%
        if viene = "GlobalAgenda" then 
            ver=0
        end if
        if ver <> "0" then
		    if viene="" or viene="agentes" or viene="comercial" or viene="albaranes_cli_fast" or viene="facturas_cli_E" then%>
			<!--<td class=CELDABOT>-->
			    <%'a la funcion SelecCampoAdi se le pasa un 12 porque son las opciones con campos no adicionales %>
				<select class="IN_S" name="campos" onchange ="javascript:SelecCampoAdi(12);">
					<option value="rsocial"><%=LitRSocial%></option>
					<option value="ncliente"><%=LitNCliente%></option>
					<option value="ncomercial"><%=LitNComercial%></option>
					<option value="cif"><%=LitCif%></option>
					<option value="contacto"><%=LitContacto%></option>
					<option value="domicilio"><%=LitDomicilio%></option>
					<option value="cp"><%=LitCp%></option>
					<option value="provincia"><%=LitProvincia%></option>
					<option value="poblacion"><%=LitPoblacion%></option>
					<option value="pais"><%=LitPais%></option>
					<option value="telefono"><%=LitTel1%></option>
					<option value="telefono2"><%=LitTel2%></option>
					<option value="email"><%=LitEMail%></option>
					<option value="fax"><%=LitFax%></option>
					<%'mmgClin: 28/03/2008 >> Se pone por defecto el nombre comercial en la búsqueda
			        if bp="comercial" then
			            'vamos a calcular y mostrar los elementos de busqueda propios del usuario
			            vect=split(replace(replace(bpca,"(",""),")",""),",")
			            Dim myArray()
			            ReDim myArray(UBound(vect))
			            
			            set rstMM = Server.CreateObject("ADODB.Recordset")
			            for i=0 to UBound(vect)            
			                cadMM="select titulo,tipo from camposperso with(nolock) where tabla='FICHA' and ncampo like '"+session("ncliente")+"%' and ncampo='" +session("ncliente")+vect(i)+"'"
			                if rstMM.State<>0 then rstMM.Close
				            rstMM.open cadMM,session("dsn_cliente")
						
				            if not rstMM.eof then%>
				                <script type="text/javascript" language="javascript">
                                    myArray[<%=i %>]=<%=rstMM("tipo") %>
				                </script>
				                <option value="<%=vect(i) %>"><%=rstMM("titulo")%></option>
				            <%end if
                        next
                        set rstMM=nothing
			        end if%>
				</select>
			<!--</td>-->
			<%'mmgClin: 28/03/2008 >> Se pone por defecto el nombre comercial en la búsqueda
			if bp="comercial" then%>
			    <script type="text/javascript" language="javascript">
                                        document.opciones.campos[2].selected = true
		        </script>
			<%end if%>

			<!--<td class=CELDABOT>-->
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContieneB%></option>
					<!--<option value="empieza"><%=LitComienzaB%></option>-->
					<option value="termina"><%=LitTerminaB%></option>
					<option value="igual"><%=LitIgualB%></option>
				</select>
			<!--</td>
			<td class=CELDABOT>-->
			    <%'mmgClin: 28/03/2007 >> Configuramos el campo de busqueda según la opcion seleccionada%>
				<input id="KeySearch" class="IN_S" type="text" name="texto" size="11" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			<!--</td>-->
		    <%end if
		end if%>
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
