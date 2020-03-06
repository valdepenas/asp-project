<%@ Language=VBScript %>

<% Server.ScriptTimeout = 300 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>" xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>" />
<meta http-equiv="Content-style-Type" content="text/css" />   


<%
      
    'JMM 20090805 Comillas dobles, para facitlitar su uso en asp
'gfg 08/07/2011 : baja de newsletter
dim cd
cd = CHR(34)

'dgb  07/04/2008  anyadimos opcion de idioma para Portugues
Dim pais_idioma, idioma, sem
'Detectamos el pais del usuario
pais_idioma = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")

set rstOrCU = Server.CreateObject("ADODB.Recordset")
%>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../BankAccountValidation.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="clientes.inc" -->

<!--#include file="../varios2.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="riesgo.inc" --> 

<!--#include file="../js/generic.js.inc"-->   
<!--#include file="../common/modal2.inc" -->  
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../common/googlemaps.inc" -->  
<!--#include file="../styles/formularios.css.inc" -->  
    
<!--#include file="../js/dropdown.js.inc" -->
<!--#include file="clientes_linkextra.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->

<!--#include file="../common/clientesActionDrop.inc" -->


<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  
<style type="text/css">
 html .ui-autocomplete {
		height: 100px;
	}
</style>
</head>
<%si_tiene_modulo_21=ModuloContratado(session("ncliente"),"21")
si_tiene_modulo_22=ModuloContratado(session("ncliente"),"22")
si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)
si_tiene_modulo_ecomerce=ModuloContratado(session("ncliente"),ModEComerce)
si_tiene_modulo_agrario=ModuloContratado(session("ncliente"),ModAgrario)
si_tiene_modulo_fidelizacion=ModuloContratado(session("ncliente"),ModFidelizacion)
si_tiene_modulo_importaciones=ModuloContratado(session("ncliente"),ModImportaciones)
si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
si_tiene_modulo_TGP = ModuloContratado(session("ncliente"),ModTGP)
si_tiene_modulo_EBESA = ModuloContratado(session("ncliente"),ModEBESA)
si_tiene_modulo_petroleos = ModuloContratado(session("ncliente"),ModPetroleos)
si_tiene_modulo_proyectos = ModuloContratado(session("ncliente"),ModProyectos)
si_tiene_modulo_contabilidad= ModuloContratado(session("ncliente"), ModContabilidad)
si_tiene_modulo_OrCU=ModuloContratado(session("ncliente"),ModOrCU)
si_tiene_modulo_AudioCenter = ModuloContratado(session("ncliente"),ModAudioCenter)
si_tiene_modulo_CMS = ModuloContratado(session("ncliente"),ModCMS)
si_tiene_modulo_eshop = ModuloContratado(session("ncliente"),ModTiendaWeb)
si_tiene_modulo_CRMComunicacion = ModuloContratado(session("ncliente"),ModCRMComunicacion)
si_tiene_modulo_fidelizacionPremium = ModuloContratado(session("ncliente"),ModFidelizacionPremium)
si_tiene_modulo_TGB=ModuloContratado(session("ncliente"),ModTGB)
si_tiene_modulo_NETTFI=ModuloContratado(session("ncliente"),ModNettfi)
si_tiene_modulo_fidelizacion30 = ModuloContratado(session("ncliente"),ModFidelizacion30)
si_tiene_modulo_BILLIB=ModuloContratado(session("ncliente"),ModBillib)   
  
validacionCliente = d_lookup("valcliente", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente"))
SuplidosActivados = nz_b2(d_lookup("USE_SUPLIDOS", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente")))

''MPC 27/05/2009 Se añade gestión del parámetro coblg campos obligatorios
dim cifrepe
dim coblg
dim valchanges
dim hide
ObtenerParametros("clientes")
campos_obligatorios = split(coblg, ",")

if request.QueryString("viene")>"" then
	viene=enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("viene") & ""))
else
	viene=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("viene") & ""))
end if

''Ricardo 16-05-2014 comprobamos que el comercial solamente vea sus clientes
dim CADCOMSOLVERSUSCLI,comercialSolSusCli
CADCOMSOLVERSUSCLI=d_lookup("CADCOMSOLVERSUSCLI","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))&""
''response.write("el CADCOMSOLVERSUSCLI es-" & CADCOMSOLVERSUSCLI & "-" & nz_b2(CADCOMSOLVERSUSCLI) & "-<br>")
if nz_b2(CADCOMSOLVERSUSCLI)=1 then
    comercialSolSusCli=d_lookup("p.dni","comerciales as c,personal as p","p.dni like '" & session("ncliente") & "%' and c.comercial=p.dni and p.login='" & session("usuario") & "'",session("dsn_cliente"))&""
    escomercialsuperior=d_lookup("c.superior","comerciales as c","c.comercial like '" & session("ncliente") & "%' and c.superior='" & session("ncliente") & session("usuario") & "'",session("dsn_cliente"))&""
    ''Ricardo 25-07-2014 si el usuario es un comercial superior, podra ver todos los clientes
    if escomercialsuperior & "">"" then
        CADCOMSOLVERSUSCLI=0
        comercialSolSusCli=""
    end if
end if
''response.write("el comercialSolSusCli es-" & comercialSolSusCli & "-<br>")<link rel="stylesheet" href="/lib/estilos/hubble/v7/breadcrumb.css" type="text/css" />
    %>

<script language="javascript" type="text/javascript" src="/lib/js/InterfaceLoadTime.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('AddDG', 'fade=1');
    animatedcollapse.addDiv('AddDC', 'fade=1');
    animatedcollapse.addDiv('AddDB', 'fade=1');
    animatedcollapse.addDiv('AddDE', 'fade=1');
    animatedcollapse.addDiv('AddOD', 'fade=1');
    animatedcollapse.addDiv('AddDD', 'fade=1');
    animatedcollapse.addDiv('AddCP', 'fade=1');
    animatedcollapse.addDiv('BrowseDG', 'fade=1');
    animatedcollapse.addDiv('BrowseDC', 'fade=1');
    animatedcollapse.addDiv('BrowseDB', 'fade=1');
    animatedcollapse.addDiv('BrowseDE', 'fade=1');
    animatedcollapse.addDiv('BrowseOD', 'fade=1');
    animatedcollapse.addDiv('BrowseDD', 'fade=1');
    animatedcollapse.addDiv('BrowseCP', 'fade=1');
    animatedcollapse.addDiv('EditDG', 'fade=1');
    animatedcollapse.addDiv('EditDC', 'fade=1');
    animatedcollapse.addDiv('EditDB', 'fade=1');
    animatedcollapse.addDiv('EditDE', 'fade=1');
    animatedcollapse.addDiv('EditOD', 'fade=1');
    animatedcollapse.addDiv('EditDD', 'fade=1');
    animatedcollapse.addDiv('EditCP', 'fade=1');

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //jQuery: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init();

</script>

<script language="javascript" type="text/javascript">
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
    function PermitirSuministro(ncliente) {
        if (window.confirm("<%=LITDESPERMSUMI%>") == true) {
            //window.alert("paso1-" + ncliente + "-");
            xmlHttp = GetXmlHttpObject();
            if (xmlHttp != null) {
                var url = "permitirSuministro.asp?mode=save&ncliente=" + ncliente;
                //window.alert("paso2-" + url + "-");
                xmlHttp.open("GET", url, false);
                xmlHttp.send(null);
                respuesta = xmlHttp.responseText;
                //window.alert(respuesta);
                if (respuesta == "OK") {
                    window.alert("<%=LITDESPERMREALIZADO%>");
                    document.getElementById("idpermitirSuministro").style.display = "none";
                }
            }
        }
    }

    //jcg 02/02/2008: necesaria para unificar con las demas paginas de compras.
    /*
    function keyPressed(tecla){
        return;
    }
    */

    function ver_imagen(modo) {
        modo = modo.toUpperCase();
        modo = modo.toLowerCase();
        if (modo != "add") {
            ncliente = document.clientes.hncliente.value;
            hmostrar_foto = marcoFoto.document.clientes_imagen.hmostrar_foto.value;
            marcoFoto.document.location = "clientes_imagen.asp?ncliente=" + ncliente + "&mode=" + modo + "&mf=" + hmostrar_foto;
        }
    }

    //asp 31/05/2011
    function cambioPoblacion() {
        document.clientes.SELECCIONAR_POBLACION1.location = "../configuracion/poblaciones.asp?mode=buscar&viene=clientes&titulo=SELECCIONAR POBLACION&diference=" + document.clientes.poblacion.value;
    }

    function validarCampoCarta(cliente, empresa) {
        if (document.clientes.cartas.value != "") {
            //ricardo 24/4/2003 se cambia todas las carta a un fichero
            AbrirVentana('generar_carta.asp?ncliente=' + cliente + '&ncentro=&ndocumento=&mode=browse&ncarta=' + document.clientes.cartas.value + '&empresa=' + empresa + '&tdocumento=clientes', 'I',<%=AltoVentana %>,<%=AnchoVentana %>);
        }
        else alert("<%=LitNoCarta%>");
    }

    function abrir_acceso(ncliente, viene, titulo, correo) {
        if (correo == '') alert("<%=LitNoAccesoTienNoMail%>");
        else AbrirVentana("../central.asp?pag1=tiendas/accesos_tienda.asp&pag2=tiendas/accesos_tienda_bt.asp&ndoc=" + ncliente + "&viene=" + viene + "&titulo=" + titulo, 'P',<%=AltoVentana %>,<%=AnchoVentana %>);
    }

    function comprobardist(cliente) {
        if (document.clientes.distribuidor.value == cliente) {
            window.alert("<%=LitMismoClientDist%>");
            document.clientes.distribuidor.value = "";
        }
    }

    function tier1Menu(objMenu, objImage) {
        if (objMenu.style.display == "none") {
            objMenu.style.display = "";
            objImage.src = "../Images/CarpetaAbierta.gif";
        }
        else {
            objMenu.style.display = "none";
            objImage.src = "../images/CarpetaCerrada.gif";
        }
    }

    function Inicia() {
        parent.document.location = "default.htm";
    }
    var modRef = 0;

    // ASP 8/06/2011
    function alcerrarModal(referencia, modoReferencia, apunta) {
        modRef = modoReferencia;
        if (apunta == "1") {
            reloadClass(referencia, "../configuracion/poblaciones.asp?mode=buscar&viene=clientes3&titulo=Lista clientes3&apunta=1&diference=" + document.clientes.poblacion.value);
        }

        if (apunta == "0") {
            reloadClass(referencia, "../configuracion/poblaciones.asp?mode=buscar&viene=clientes3&titulo=Lista clientes3&apunta=0&diference=" + document.clientes.poblacion.value + "&diference2=" + document.clientes.provincia.value);
        }
        alPresionar(referencia);
        jQuery('.window .close').click(function (e) {
            recargaGuardar(apunta);
        });
        jQuery('#mask').click(function () {
            recargaGuardar(apunta);
        });
    }

    function recargaGuardar(apunta) {
        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
        //reloadIframe("SELECCIONAR_POBLACION2","../configuracion/poblaciones.asp?mode=buscar&viene=clientes&titulo=Lista clientes");
        if (apunta == "1") {
            document.clientes.action = "clientes.asp?mode=save&repe=" + modRef;
        }
        if (apunta == "0") {
            document.clientes.action = "clientes.asp?mode=save&ncliente=" + document.clientes.hncliente.value + "&repe=" + modRef;
        }
        parent.botones.document.opciones.action = "clientes_bt.asp?mode=browse";
        parent.botones.document.opciones.submit();
        document.clientes.submit();
    }

    function borrarCodigos(de) {
        if (de == "1") {
            parent.pantalla.document.clientes.codPoblacion.value = "";
        }

        if (de == "2") {
            parent.pantalla.document.clientes.de_codPoblacion.value = "";
        }
        if (de == "3") {
            parent.pantalla.document.clientes.codPoblacionGB.value = "";
        }
    }
    //FIN ASP
    //dgm - para la dirección de envio de factura

    function RecargarModales(frame, url) {
        reloadClass(frame, url);
        alPresionar(frame);
    }

    function borrarCodigosF(de) {
        if (de == "1") {

            parent.pantalla.document.clientes.codPoblacionF.value = "";
        }
        5600

        if (de == "2") {
            parent.pantalla.document.clientes.de_codPoblacionF.value = "";
        }
    }

    function ModalPoblacionFactura() {

    }
    //fin dgm

    function CopiarCampos() {
        document.clientes.de_domicilio.value = document.clientes.domicilio.value;
        document.clientes.de_poblacion.value = document.clientes.poblacion.value;
        document.clientes.de_cp.value = document.clientes.cp.value;
        document.clientes.de_provincia.value = document.clientes.provincia.value;
        document.clientes.de_pais.value = document.clientes.pais.value;
        document.clientes.de_telefono.value = document.clientes.telefono.value;
        document.clientes.de_codProvincia.value = document.clientes.codProvincia.value;
        document.clientes.de_codPoblacion.value = document.clientes.codPoblacion.value;
        document.clientes.de_codPais.value = document.clientes.codPais.value;
        ddlProvincia = document.clientes.provinciaDDL;
        if (ddlProvincia.options[ddlProvincia.selectedIndex].text != "") {
            de_ddlProvincia = document.clientes.de_provinciaDDL;
            de_ddlProvincia.options[de_ddlProvincia.selectedIndex].text = ddlProvincia.options[ddlProvincia.selectedIndex].text;
            document.clientes.de_provincia.readOnly = true;
            de_ddlProvincia.options[de_ddlProvincia.selectedIndex].value = ddlProvincia.options[ddlProvincia.selectedIndex].value;
        }
        if (ddlPais.options[ddlPais.selectedIndex].text != "") {
            de_ddlPais = document.clientes.de_paisDDL;
            de_ddlPais.options[de_ddlPais.selectedIndex].text = ddlPais.options[ddlPais.selectedIndex].text;
            document.clientes.de_pais.readOnly = true;
            de_ddlPais.options[de_ddlPais.selectedIndex].value = ddlPais.options[ddlPais.selectedIndex].value;
        }
    }

    function CopiarCamposF() {
        document.clientes.de_domicilioF.value = document.clientes.domicilio.value;
        document.clientes.de_poblacionF.value = document.clientes.poblacion.value;
        document.clientes.de_cpF.value = document.clientes.cp.value;
        document.clientes.de_provinciaF.value = document.clientes.provincia.value;
        document.clientes.de_paisF.value = document.clientes.pais.value;
        document.clientes.de_telefonoF.value = document.clientes.telefono.value;
        document.clientes.de_codProvinciaF.value = document.clientes.codProvincia.value;
        document.clientes.de_codPoblacionF.value = document.clientes.codPoblacion.value;
        document.clientes.de_codPaisF.value = document.clientes.codPais.value;
        ddlProvincia = document.clientes.provinciaDDL;
        if (ddlProvincia.options[ddlProvincia.selectedIndex].text != "") {
            de_ddlProvinciaF = document.clientes.de_provinciaDDLF;
            de_ddlProvinciaF.options[de_ddlProvinciaF.selectedIndex].text = ddlProvincia.options[ddlProvincia.selectedIndex].text;
            document.clientes.de_provinciaF.readOnly = true;
            de_ddlProvinciaF.options[de_ddlProvinciaF.selectedIndex].value = ddlProvincia.options[ddlProvincia.selectedIndex].value;

        }
        if (ddlPais.options[ddlPais.selectedIndex].text != "") {
            de_ddlPaisF = document.clientes.de_paisDDLF;
            de_ddlPaisF.options[de_ddlPaisF.selectedIndex].text = ddlPais.options[ddlPais.selectedIndex].text;
            document.clientes.de_paisF.readOnly = true;
            de_ddlPaisF.options[de_ddlPaisF.selectedIndex].value = ddlPais.options[ddlPais.selectedIndex].value;
        }
    }

    function EliminarDirEnvio(ncliente) {
        if (window.confirm("<%=LitMsgElimDirEnvioConfirm%>") == true) {
            document.clientes.action = "clientes.asp?mode=borrardirenvio&ncliente=" + ncliente;
            document.clientes.submit();
        }
    }

    function EliminarDirEnvioF(ncliente) {
        if (window.confirm("<%=LitMsgElimDirEnvioConfirm%>") == true) {
            document.clientes.action = "clientes.asp?mode=borrardirenvioF&ncliente=" + ncliente;
            document.clientes.submit();
        }
    }

    function actualizavalores() {
        if (isNaN(document.clientes.vencom1.value.replace(",", "."))) {
            alert("<%=LitMsgLV1Numerico%>");
            return false;
        }
        document.clientes.vencom1.value = document.clientes.vencom1.value.replace(".", ",");
        if (isNaN(document.clientes.vencom2.value.replace(",", "."))) {
            alert("<%=LitMsgLV2Numerico%>");
            return false;
        }
        document.clientes.vencom2.value = document.clientes.vencom2.value.replace(".", ",");
        if (isNaN(document.clientes.pcom1.value.replace(",", "."))) {
            alert("<%=LitMsgPorComisionNumerico%>");
            return false;
        }
        document.clientes.pcom1.value = document.clientes.pcom1.value.replace(".", ",");
        if (isNaN(document.clientes.pcom2.value.replace(",", "."))) {
            alert("<%=LitMsgPorComision2Numerico%>");
            return false;
        }
        document.clientes.pcom2.value = document.clientes.pcom2.value.replace(".", ",");
    }

    function convclidist(ncliente) {
        if (window.confirm("<%=LitConvertirCliDis%>")) document.location = "clientes.asp?mode=convertirclidist&ncliente=" + ncliente + "&salto=no";
    }

    function CrearCentro(ncliente) {
        if (window.confirm("<%=LitCrearCentro%>")) document.location = "clientes.asp?mode=crearcentro&ncliente=" + ncliente + "&salto=no";
    }

    function VerCultivos(ncliente) {
        AbrirVentana("../central.asp?pag1=ventas/cultivos_cli.asp&pag2=ventas/cultivos_cli_bt.asp&ncliente=" + ncliente + "&viene=clientes&titulo=<%=LitCutlDelCli%> : " + trimCodEmpresa(ncliente), 'P',<%=AltoVentana %>,<%=AnchoVentana %>);
    }

    function ComprobarCantRiesgo() {
        while (document.clientes.rgomaxaut.value.search(" ") != -1) document.clientes.rgomaxaut.value = document.clientes.rgomaxaut.value.replace(" ", "");
        if (isNaN(document.clientes.rgomaxaut.value.replace(",", "."))) {
            window.alert("<%=LitMsgRiesgoMaxNumerico%>");
            document.clientes.rgomaxaut.value = 0;
            return;
        }
        if (parseFloat(document.clientes.rgomaxaut.value) < 0) {
            window.alert("<%=LitMsgRiesgoMaxNoNegativo%>");
            document.clientes.rgomaxaut.value = 0;
            return;
        }
        document.clientes.rgomaxaut.value = document.clientes.rgomaxaut.value.replace(".", ",");
    }
    //-->

    function comprobar() {
        if (parseInt(document.clientes.e_primer_ven.value) > 31) document.clientes.e_primer_ven.value = 31;
        /* cag */
        if (parseInt(document.clientes.e_segundo_ven.value) > 31) document.clientes.e_segundo_ven.value = 31;
        if (parseInt(document.clientes.e_tercer_ven.value) > 31) document.clientes.e_tercer_ven.value = 31;
        if ((parseInt(document.clientes.e_segundo_ven.value) <= parseInt(document.clientes.e_primer_ven.value))
            && (parseInt(document.clientes.e_segundo_ven.value) > 0)) {
            alert("<%=LITDIAPAGO2INFERIORDIAPAGO1%>");
            document.clientes.e_segundo_ven.focus();
        }
        else
            if ((parseInt(document.clientes.e_tercer_ven.value) <= parseInt(document.clientes.e_segundo_ven.value))
                && (parseInt(document.clientes.e_tercer_ven.value) > 0)) {
                alert("<%=LITDIAPAGO3INFERIORDIAPAGO2%>");
                document.clientes.e_tercer_ven.focus();
            }
    }

    /*
    function ComprobarExisteCIF()
    {
        // Use the native cross-browser nitobi Ajax object
        var myAjaxRequest = new nitobi.ajax.HttpRequest();
        var ncliente="";
    
        if (document.clientes.hncliente != null) ncliente = document.clientes.hncliente.value;
    
        // Define the url for your generatekey script
        myAjaxRequest.handler = "existeCIFCliente.asp?cif=" + document.clientes.cif.value + "&mode=<%=Request.QueryString("mode")%>&ncliente=" + ncliente;
        myAjaxRequest.async = false;
        myAjaxRequest.get();
    	
        //window.alert("el clientes es-" + document.clientes.cif.value + "-" + ncliente + "-" + myAjaxRequest.httpObj.responseText + "-");
    
        // return the result to the grid
        return myAjaxRequest.httpObj.responseText;
    }
    */

    function GuardarFactura() {
        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
        parent.pantalla.document.clientes.action = "clientes.asp?mode=save&viene=facturas_cli_E&accion=si&ndoc=<%=enc.EncodeForJavascript(Request.QueryString("ndoc"))%>&ncliente=<%=enc.EncodeForJavascript(limpiaCadena(Request.QueryString("ncliente")))%>&nfactura="+ document.clientes.nfactura.value + "&obs=" + document.clientes.obs.value;
        parent.pantalla.document.clientes.submit();
    }

    function iva0() {
        if (document.clientes.intra.checked) {
            document.clientes.iva.value = 0;
            document.clientes.iva.disabled = true;
        }
        else document.clientes.iva.disabled = "";
    }

    /*AMP 28/06/2010 Función para realizar llamada saliente */
    function Actualizar(cliente) {
        document.clientes.action = "clientes.asp?ncliente=" + cliente + "mode=browse";
        document.clientes.submit();
        parent.botones.document.location = "clientes_bt.asp?mode=browse";
    }

    function RealizarLlamada(telf, imagen, d) {
        var myString = "";
        var driver;
        driver = 1; //GetDriver(d);
        // alert(driver);
        if (driver > 0) {
            if (imagen == "telffijo") {
                myString = String(document.images.telffijo.src);
                var mySplitResult = myString.split("images/");
                if (mySplitResult[1] == "<%=ImgTelfLlamar%>") {
                    document.images.telffijo.src = "../images/" + "<%=ImgTelfColgar%>";
                    document.images.telffijo.alt = "<%=LitColgar%>";
                    document.images.telffijo.title = "<%=LitColgar%>";
                    //PressConnect(telf);
                }
                else {
                    document.images.telffijo.src = "../images/" + "<%=ImgTelfLlamar%>";
                    document.images.telffijo.alt = "<%=LitLlamar%>";
                    document.images.telffijo.title = "<%=LitLlamar%>";
                    //DisconnectCall(0);
                }
            }
            else
                if (imagen == "telfmovil") {
                    myString = String(document.images.telfmovil.src);
                    var mySplitResult = myString.split("images/");
                    if (mySplitResult[1] == "<%=ImgTelfLlamar%>") {
                        document.images.telfmovil.src = "../images/" + "<%=ImgTelfColgar%>";
                        document.images.telfmovil.alt = "<%=LitColgar%>";
                        document.images.telfmovil.title = "<%=LitColgar%>";
                        //PressConnect(telf);         
                    }
                    else {
                        document.images.telfmovil.src = "../images/" + "<%=ImgTelfLlamar%>";
                        document.images.telfmovil.alt = "<%=LitLlamar%>";
                        document.images.telfmovil.title = "<%=LitLlamar%>";
                        //DisconnectCall(0);   
                    }
                }
        }
        else {
            alert("<%=LitDriver%>");
        }
    }

    function str_replace(search, replace, subject) {
        var f = search, r = replace, s = subject;
        var ra = r instanceof Array, sa = s instanceof Array, f = [].concat(f), r = [].concat(r), i = (s = [].concat(s)).length;

        while (j = 0, i--) {
            if (s[i]) {
                while (s[i] = s[i].split(f[j]).join(ra ? r[j] || "" : r[0]), ++j in f) { };
            }
        };

        return sa ? s : s[0];
    }

    function valida_nif_cif_nie() {
    <% if validacionCliente then %>
    var cif = document.clientes.cif.value;
        var temp = cif.toUpperCase();
        var cadenadni = "TRWAGMYFPDXBNJZSQVHLCKE";
        if (temp !== '') {

            //si no tiene un formato valido devuelve error
            if ((!/^[A-Z]{1}[0-9]{7}[A-Z0-9]{1}$/.test(temp) && !/^[T]{1}[A-Z0-9]{8}$/.test(temp)) && !/^[0-9]{8}[A-Z]{1}$/.test(temp)) {
                return 0;
            }

            //comprobacion de NIFs estandar
            if (/^[0-9]{8}[A-Z]{1}$/.test(temp)) {

                posicion = cif.substring(8, 0) % 23;
                letra = cadenadni.charAt(posicion);
                var letradni = temp.charAt(8);
                if (letra == letradni) {
                    return 1;
                }
                else {
                    return -1;
                }
            }

            //comprobacion de CIFs
            if (/^[ABCDEFGHJNPQRSUVW]{1}/.test(temp)) {
                var pares = 0;
                var impares = 0;
                var suma;
                var ultima;
                var unumero;
                var uletra = new Array("J", "A", "B", "C", "D", "E", "F", "G", "H", "I");
                var xxx;
                texto = temp;
                var regular = new RegExp(/^[ABCDEFGHKLMNPQS]\d{7}[0-9,A-J]$/g);
                if (!regular.exec(texto)) return false;

                ultima = texto.substr(8, 1);

                for (var cont = 1; cont < 7; cont++) {
                    xxx = (2 * parseInt(texto.substr(cont++, 1))).toString() + "0";
                    impares += parseInt(xxx.substr(0, 1)) + parseInt(xxx.substr(1, 1));
                    pares += parseInt(texto.substr(cont, 1));
                }

                xxx = (2 * parseInt(texto.substr(cont, 1))).toString() + "0";
                impares += parseInt(xxx.substr(0, 1)) + parseInt(xxx.substr(1, 1));

                suma = (pares + impares).toString();
                unumero = parseInt(suma.substr(suma.length - 1, 1));
                unumero = (10 - unumero).toString();
                if (unumero == 10) unumero = 0;

                if ((ultima == unumero) || (ultima == uletra[unumero]))
                    return 2;
                else
                    return -2;
            }

            //comprobacion de NIEs
            //T
            if (/^[T]{1}/.test(temp)) {
                if (temp[8] == /^[T]{1}[A-Z0-9]{8}$/.test(temp)) {
                    return 3;
                }
                else {
                    return -3;
                }
            }

            //XYZ
            if (/^[XYZ]{1}/.test(temp)) {
                var dni = temp;
                var pre = dni.substr(0, 1);
                var prev = '0';
                if (pre == 'X')
                    prev = '0';
                else if (pre == 'Y')
                    prev = '1';
                else if (pre == 'Z')
                    prev = '2';
                numero = prev + dni.substr(1, dni.length - 1);

                var tempC = numero;
                var tempN = tempC.toUpperCase();
                posicion = tempC.substring(8, 0) % 23;
                letra = cadenadni.charAt(posicion);
                var letradni = tempN.charAt(8);
                if (letra == letradni) {
                    return 1;
                }
                else {
                    return -1;
                }
            }
        }

	<% else %>
	    return 5;
    <% end if %>
 
	return 0;
    }

    function gestionCIF_NIF_NIE() {
        if (valida_nif_cif_nie() <= 0)
            alert("<%=LIT_MSG_CIFNIF_NOVALIDO %>");
    }

    //AMF:21/12/2010:Lleva el centro a la incidencia que abrio la ventana de creacion de clientes.
    function LlevarCentroIncidencia(ncentro) {
        window.top.opener.parent.pantalla.TraerNuevoCentro(ncentro, "../mantenimiento/");
        parent.window.close();
    }

    function LlevarClienteAOrdenes(ncliente, ncentro, generadoCentro) {
        if (generadoCentro == "1") {
            //window.alert(window.top.opener.parent.pantalla.document.location);
            //window.alert(window.top.opener.window.top.opener.parent.pantalla.document.location);
            window.top.opener.window.top.opener.parent.pantalla.TraerNuevoCentro(ncentro, "../mantenimiento/");
            window.top.opener.parent.window.close();
            parent.window.close();
        }
        else {
            window.top.opener.parent.pantalla.frProyecto.document.docclientes.ncliente.value = trimCodEmpresa(ncliente);
            window.top.opener.parent.pantalla.frProyecto.TraerCliente("add", "centros", "");
            parent.window.close();
        }
    }

    function saldoMaxChanged() {
        if (isNaN(document.clientes.saldomax.value.replace(",", "."))) {
            window.alert("<%= Msgerrsaldo%>");
            document.clientes.saldomax.value = document.clientes.hd_saldoMax.value;
            return false;
        }
        document.clientes.hd_saldoEnvidado.value = 1;
        document.clientes.hd_saldoMax.value = document.clientes.saldoact.value = document.clientes.saldomax.value;
    }

    function SaldoSinLimiteChanged() {
        if (document.clientes.cbSaldoSinLimite.checked == true) {
            document.clientes.saldomax.value = '9999999.99';
            document.clientes.saldomax.disabled = true;
            saldoMaxChanged();
        }
        else {
            document.clientes.saldomax.value = '';
            document.clientes.saldomax.disabled = false;
            saldoMaxChanged();
        }
    }

    function inicio() {
        if (document.getElementById('AddDG') != null) {
            //if (AddDG.style.display != "none")
            if (document.getElementById('AddDG').style.display != "none") {
                document.getElementsByName('rsocial')[0].focus();
            }
        }

        if (document.getElementById('EditDG') != null) {
            //if (EditDG.style.display != "none")
            if (document.getElementById('EditDG').style.display != "none") {
                //document.getElementById('rsocial').focus();
                document.getElementsByName('rsocial')[0].focus();
            }
        }
    }

    function TraerPais() {
        ddlPais = document.clientes.paisDDL;
        if (ddlPais.options[ddlPais.selectedIndex].text != "") {
            document.clientes.pais.value = ddlPais.options[ddlPais.selectedIndex].text;
            document.clientes.codPais.value = ddlPais.options[ddlPais.selectedIndex].value;
            document.clientes.pais.readOnly = true;
            document.clientes.pais.style.color = "#ACA899";
            document.clientes.pais.style.cursor = "default";
        }
        else {
            //document.clientes.pais.value = "";
            document.clientes.codPais.value = "";
            document.clientes.pais.readOnly = false;
            document.clientes.pais.style.color = "#000000";
            document.clientes.pais.style.cursor = "text";
        }

        FiltrarProvinciasAlCambiarPais();
    }

    // Production steps of ECMA-262, Edition 6, 22.1.2.1
    if (!Array.from) {
        Array.from = (function () {
            var toStr = Object.prototype.toString;
            var isCallable = function (fn) {
                return typeof fn === 'function' || toStr.call(fn) === '[object Function]';
            };
            var toInteger = function (value) {
                var number = Number(value);
                if (isNaN(number)) { return 0; }
                if (number === 0 || !isFinite(number)) { return number; }
                return (number > 0 ? 1 : -1) * Math.floor(Math.abs(number));
            };
            var maxSafeInteger = Math.pow(2, 53) - 1;
            var toLength = function (value) {
                var len = toInteger(value);
                return Math.min(Math.max(len, 0), maxSafeInteger);
            };

            // The length property of the from method is 1.
            return function from(arrayLike/*, mapFn, thisArg */) {
                // 1. Let C be the this value.
                var C = this;

                // 2. Let items be ToObject(arrayLike).
                var items = Object(arrayLike);

                // 3. ReturnIfAbrupt(items).
                if (arrayLike == null) {
                    throw new TypeError('Array.from requires an array-like object - not null or undefined');
                }

                // 4. If mapfn is undefined, then let mapping be false.
                var mapFn = arguments.length > 1 ? arguments[1] : void undefined;
                var T;
                if (typeof mapFn !== 'undefined') {
                    // 5. else
                    // 5. a If IsCallable(mapfn) is false, throw a TypeError exception.
                    if (!isCallable(mapFn)) {
                        throw new TypeError('Array.from: when provided, the second argument must be a function');
                    }

                    // 5. b. If thisArg was supplied, let T be thisArg; else let T be undefined.
                    if (arguments.length > 2) {
                        T = arguments[2];
                    }
                }

                // 10. Let lenValue be Get(items, "length").
                // 11. Let len be ToLength(lenValue).
                var len = toLength(items.length);

                // 13. If IsConstructor(C) is true, then
                // 13. a. Let A be the result of calling the [[Construct]] internal method 
                // of C with an argument list containing the single item len.
                // 14. a. Else, Let A be ArrayCreate(len).
                var A = isCallable(C) ? Object(new C(len)) : new Array(len);

                // 16. Let k be 0.
                var k = 0;
                // 17. Repeat, while k < len… (also steps a - h)
                var kValue;
                while (k < len) {
                    kValue = items[k];
                    if (mapFn) {
                        A[k] = typeof T === 'undefined' ? mapFn(kValue, k) : mapFn.call(T, kValue, k);
                    } else {
                        A[k] = kValue;
                    }
                    k += 1;
                }
                // 18. Let putStatus be Put(A, "length", len, true).
                A.length = len;
                // 20. Return A.
                return A;
            };
        }());
    }

    function FiltrarProvinciasAlCambiarPais() {
        var idPais = document.clientes.paisDDL.value;
        var provinciasDDL = document.clientes.provinciaDDL;
        // ie 11 not support Array.from(), added code if (!Array.from)
        var provincias = Array.from(provinciasDDL.options);
        var inputProvincia = document.clientes.provincia;

        // Mostrar las provincias con código de país igual al seleccionado
        //Script1002 Syntax error, ie 11 not support arrow functions
        //provincias
        //    .filter(p => p.hasAttribute('data-country') && p.getAttribute('data-country') == idPais)
        //    .forEach(p => p.style.display = "block");
        provincias
            .filter(function(p) { return p.hasAttribute("data-country") && p.getAttribute("data-country") == idPais; })
            .forEach(function (p) { p.style.display = "block" });
        // Esconder las provincias sin código de país igual al seleccionado
        //provincias
        //    .filter(p => p.hasAttribute('data-country') && p.getAttribute('data-country') != idPais)
        //    .forEach(p => p.style.display = "none");
        provincias
            .filter(function(p) { return p.hasAttribute("data-country") && p.getAttribute("data-country") != idPais; })
            .forEach(function (p) { p.style.display = "none" });

        provinciasDDL.selectedIndex = provincias.length - 1;
        inputProvincia.value = "";
    }

    function TraerPaisDe() {
        de_ddlPais = document.clientes.de_paisDDL;
        if (de_ddlPais.options[de_ddlPais.selectedIndex].text != "") {
            document.clientes.de_pais.value = de_ddlPais.options[de_ddlPais.selectedIndex].text;
            //document.clientes.de_pais.text =de_ddlPais.options[de_ddlPais.selectedIndex].text;
            document.clientes.de_codPais.value = de_ddlPais.options[de_ddlPais.selectedIndex].value;
            document.clientes.de_pais.readOnly = true;
            document.clientes.de_pais.style.color = "#ACA899";
            document.clientes.de_pais.style.cursor = "default";
        }
        else {
            //document.clientes.de_pais.value = "";
            document.clientes.de_codPais.value = "";
            document.clientes.de_pais.readOnly = false;
            document.clientes.de_pais.style.color = "#000000";
            document.clientes.de_pais.style.cursor = "text";
        }
    }

    function TraerPaisDeF() {
        de_ddlPaisF = document.clientes.de_paisDDLF;
        if (de_ddlPaisF.options[de_ddlPaisF.selectedIndex].text != "") {
            document.clientes.de_paisF.value = de_ddlPaisF.options[de_ddlPaisF.selectedIndex].text;
            document.clientes.de_codPaisF.value = de_ddlPaisF.options[de_ddlPaisF.selectedIndex].value;
            document.clientes.de_paisF.readOnly = true;
            document.clientes.de_paisF.style.color = "#ACA899";
            document.clientes.de_paisF.style.cursor = "default";
        }
        else {
            //document.clientes.de_paisF.value = "";
            document.clientes.de_codPaisF.value = "";
            document.clientes.de_paisF.readOnly = false;
            document.clientes.de_paisF.style.color = "#000000";
            document.clientes.de_paisF.style.cursor = "text";
        }
    }

    function TraerProvincia() {
        ddlProvincia = document.clientes.provinciaDDL;
        if (ddlProvincia.options[ddlProvincia.selectedIndex].text != "") {
            document.clientes.provincia.value = ddlProvincia.options[ddlProvincia.selectedIndex].text;
            document.clientes.codProvincia.value = ddlProvincia.options[ddlProvincia.selectedIndex].value;
            document.clientes.provincia.readOnly = true;
            document.clientes.provincia.style.color = "#ACA899";
            document.clientes.provincia.style.cursor = "default";
        }
        else {
            //document.clientes.provincia.value = "";
            document.clientes.codProvincia.value = "";
            document.clientes.provincia.readOnly = false;
            document.clientes.provincia.style.color = "#000000";
            document.clientes.provincia.style.cursor = "text";
        }
    }

    function TraerProvinciaDe() {
        de_ddlProvincia = document.clientes.de_provinciaDDL;
        if (de_ddlProvincia.options[de_ddlProvincia.selectedIndex].text != "") {
            document.clientes.de_provincia.value = de_ddlProvincia.options[de_ddlProvincia.selectedIndex].text;
            document.clientes.de_codProvincia.value = de_ddlProvincia.options[de_ddlProvincia.selectedIndex].value;
            document.clientes.de_provincia.readOnly = true;
            document.clientes.de_provincia.style.color = "#ACA899";
            document.clientes.de_provincia.style.cursor = "default";
        }
        else {
            //document.clientes.de_provincia.value = "";
            document.clientes.de_codProvincia.value = "";
            document.clientes.de_provincia.readOnly = false;
            document.clientes.de_provincia.style.color = "#000000";
            document.clientes.de_provincia.style.cursor = "text";
        }
    }

    function TraerProvinciaDeF() {
        de_ddlProvinciaF = document.clientes.de_provinciaDDLF;
        if (de_ddlProvinciaF.options[de_ddlProvinciaF.selectedIndex].text != "") {
            document.clientes.de_provinciaF.value = de_ddlProvinciaF.options[de_ddlProvinciaF.selectedIndex].text;
            document.clientes.de_codProvinciaF.value = de_ddlProvinciaF.options[de_ddlProvinciaF.selectedIndex].value;
            document.clientes.de_provinciaF.readOnly = true;
            document.clientes.de_provinciaF.style.color = "#ACA899";
            document.clientes.de_provinciaF.style.cursor = "default";
        }
        else {
            //document.clientes.de_provinciaF.value = "";
            document.clientes.de_codProvinciaF.value = "";
            document.clientes.de_provinciaF.readOnly = false;
            document.clientes.de_provinciaF.style.color = "#000000";
        }
    }

    function TraerProvinciaGB() {
        gb_ddlProvincia = document.clientes.gb_provincia;
        if (gb_ddlProvincia.options[gb_ddlProvincia.selectedIndex].text != "") {
            document.clientes.codProvinciaGB.value = gb_ddlProvincia.options[gb_ddlProvincia.selectedIndex].value;
        }
        else {
            document.clientes.codProvinciaGB.value = "";
        }
    }

    function ComprobarDDL() {
        ddlPais = document.clientes.paisDDL;
        if (ddlPais.options[ddlPais.selectedIndex].text != "") {
            document.clientes.pais.readOnly = true;
            document.clientes.pais.style.color = "#ACA899";
            document.clientes.pais.style.cursor = "default";
        }
        else if (document.clientes.pais.value != "") {
            var A = ddlPais.options, L = A.length;
            while (L > 0) {
                if (A[--L].text == document.clientes.pais.value) {
                    ddlPais.selectedIndex = L;
                    L = 0;
                    document.clientes.pais.readOnly = true;
                    document.clientes.pais.style.color = "#ACA899";
                    document.clientes.pais.style.cursor = "default";
                }
            }
        }

        de_ddlPais = document.clientes.de_paisDDL;
        if (de_ddlPais.options[de_ddlPais.selectedIndex].text != "") {
            document.clientes.de_pais.readOnly = true;
            document.clientes.de_pais.style.color = "#ACA899";
            document.clientes.de_pais.style.cursor = "default";
        }
        else if (document.clientes.de_pais.value != "") {
            var A = de_ddlPais.options, L = A.length;
            while (L > 0) {
                if (A[--L].text == document.clientes.de_pais.value) {
                    de_ddlPais.selectedIndex = L;
                    L = 0;
                    document.clientes.de_pais.readOnly = true;
                    document.clientes.de_pais.style.color = "#ACA899";
                    document.clientes.de_pais.style.cursor = "default";
                }
            }
        }

        de_ddlPaisF = document.clientes.de_paisDDLF;
        if (de_ddlPaisF.options[de_ddlPaisF.selectedIndex].text != "") {
            document.clientes.de_paisF.readOnly = true;
            document.clientes.de_paisF.style.color = "#ACA899";
            document.clientes.de_paisF.style.cursor = "default";

        }
        else if (document.clientes.de_paisF.value != "") {
            var A = de_ddlPaisF.options, L = A.length;
            while (L > 0) {
                if (A[--L].text == document.clientes.de_paisF.value) {
                    de_ddlPaisF.selectedIndex = L;
                    L = 0;
                    document.clientes.de_paisF.readOnly = true;
                    document.clientes.de_paisF.style.color = "#ACA899";
                    document.clientes.de_paisF.style.cursor = "default";
                }
            }
        }

        ddlProvincia = document.clientes.provinciaDDL;
        if (ddlProvincia.options[ddlProvincia.selectedIndex].text != "") {
            document.clientes.provincia.readOnly = true;
            document.clientes.provincia.style.color = "#ACA899";
            document.clientes.provincia.style.cursor = "default";
        }
        else if (document.clientes.provincia.value != "") {
            var A = ddlProvincia.options, L = A.length;
            while (L > 0) {
                if (A[--L].text == document.clientes.provincia.value) {
                    ddlProvincia.selectedIndex = L;
                    L = 0;
                    document.clientes.provincia.readOnly = true;
                    document.clientes.provincia.style.color = "#ACA899";
                    document.clientes.provincia.style.cursor = "default";
                }
            }
        }

        de_ddlProvincia = document.clientes.de_provinciaDDL;
        if (de_ddlProvincia.options[de_ddlProvincia.selectedIndex].text != "") {
            document.clientes.de_provincia.readOnly = true;
            document.clientes.de_provincia.style.color = "#ACA899";
            document.clientes.de_provincia.style.cursor = "default";

        }
        else if (document.clientes.de_provincia.value != "") {
            var A = de_ddlProvincia.options, L = A.length;
            while (L > 0) {
                if (A[--L].text == document.clientes.de_provincia.value) {
                    de_ddlProvincia.selectedIndex = L;
                    L = 0;
                    document.clientes.de_provincia.readOnly = true;
                    document.clientes.de_provincia.style.color = "#ACA899";
                    document.clientes.de_provincia.style.cursor = "default";
                }
            }
        }

        de_ddlProvinciaF = document.clientes.de_provinciaDDLF;
        if (de_ddlProvinciaF.options[de_ddlProvinciaF.selectedIndex].text != "") {
            document.clientes.de_provinciaF.readOnly = true;
            document.clientes.de_provinciaF.style.color = "#ACA899";
            document.clientes.de_provinciaF.style.cursor = "default";

        }
        else if (document.clientes.de_provinciaF.value != "") {
            var A = de_ddlProvinciaF.options, L = A.length;
            while (L > 0) {
                if (A[--L].text == document.clientes.de_provinciaF.value) {
                    de_ddlProvinciaF.selectedIndex = L;
                    L = 0;
                    document.clientes.de_provinciaF.readOnly = true;
                    document.clientes.de_provinciaF.style.color = "#ACA899";
                    document.clientes.de_provinciaF.style.cursor = "default";
                }
            }
        }
    }

    function ConfirmChangesCustomer() {
        var ran = Math.random();
        paginaModal = "/<%=CarpetaProduccion%>/central.asp?pag1=gestion/clientes/confirmDataCustomer.asp&ncliente=" + document.clientes.hncliente.value + "&pag2=gestion/clientes/confirmDataCustomer_bt.asp&ran=" + ran;
        cambiarTamanyo("#SELECCIONAR_POBLACION2", "400", "860");
        reloadClass("#SELECCIONAR_POBLACION2", paginaModal);
        alPresionar("#SELECCIONAR_POBLACION2");
    }
    //jQuery(document).on("ready", function () {
    //    jQuery(".checkboxes_radio").checkboxradio();
    //}); 
    

    jQuery(function () {
        jQuery("input[name=poblacion]").autocomplete({
            source: "datosJson.asp?consulta=poblaciones",
            minLength: 3,

            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=codPoblacion]").val(ui.item.codigo);
            }
        });
    });

    jQuery(function () {
        jQuery("input[name=de_poblacion]").autocomplete({
            source: "datosJson.asp?consulta=poblaciones",
            minLength: 3,

            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=de_codPoblacion]").val(ui.item.id);
            }
        });
    });

    jQuery(function () {
        jQuery("input[name=de_poblacionF]").autocomplete({
            source: "datosJson.asp?consulta=poblaciones",
            minLength: 3,

            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=de_codPoblacionF]").val(ui.item.id);
            }
        });
    });


    jQuery(function () {
        jQuery("input[name=provincia]").autocomplete({
            source: "datosJson.asp?consulta=provincias",
            minLength: 3,

            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=codProvincia]").val(ui.item.id);
            }
        });
    });

    jQuery(function () {

        jQuery("input[name=pais]").autocomplete({
            source: "datosJson.asp?consulta=paises",
            minLength: 3,

            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=codPais]").val(ui.item.id);
            }
        });
    });

    jQuery(function () {
        jQuery("input[name=de_provincia]").autocomplete({
            source: "datosJson.asp?consulta=provincias",
            minLength: 3,
            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=de_codProvincia]").val(ui.item.id);
            }
        });
    });

    jQuery(function () {
        jQuery("input[name=de_pais]").autocomplete({
            source: "datosJson.asp?consulta=paises",
            minLength: 3,
            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=de_codPais]").val(ui.item.id);
            }
        });
    });

    jQuery(function () {
        jQuery("input[name=de_provinciaF]").autocomplete({
            source: "datosJson.asp?consulta=provincias",
            minLength: 3,
            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=de_codProvinciaF]").val(ui.item.id);
            }
        });
    });

    jQuery(function () {
        jQuery("input[name=de_paisF]").autocomplete({
            source: "datosJson.asp?consulta=paises",
            minLength: 3,
            select: function (event, ui) {
                // perhaps do something with these?      
                region = ui.item.label;
                jQuery("input[name=de_codPaisF]").val(ui.item.id);
                //TraerCliente('add','1');
                //cargaCliente(ui.item.id);
            }
        });
    });

    jQuery(function () {
        jQuery("input[name=PoblacionGB]").autocomplete({
            source: "datosJson.asp?consulta=poblaciones",
            minLength: 3,

            select: function (event, ui) {
                region = ui.item.label;
                jQuery("input[name=codPoblacionGB]").val(ui.item.codigo);
            }
        });
    });

    // *************************************************************************************
    //                                  LISTAS DEPENDIENTES 
    // *************************************************************************************

    // LISTAS DEPENDIENTES 
    function AgregarElementoArray(CodPadre, CodHijo, TextoHijo) {
        this.CodPadre = CodPadre
        this.CodHijo = CodHijo
        this.TextoHijo = TextoHijo
    }

    function RelacionListas(CampoPadre, CampoHijo) {
        this.CampoPadre = CampoPadre
        this.CampoHijo = CampoHijo
    }

    function borrarListas(ElementoBorrar) {
        ElementoBorrar.length = 0
        for (ii = 0; ii < lov_relaciones.length; ii++) {
            if (lov_relaciones[ii].CampoPadre == ElementoBorrar.name) // es dependiente de alguna lista, hay que borrar tb sus hijos
            {
                nuevoBorrar = document.getElementById(lov_relaciones[ii].CampoHijo)
                borrarListas(nuevoBorrar)
            }
        }
    }

    function ControlListasDependientes() {
        var donde;
        var cual = this
        var recorreExistentes;
        for (i = 0; i < lov_relaciones.length; i++) {
            if (lov_relaciones[i].CampoPadre == cual.name) // es dependiente de alguna lista
            {
                donde = document.getElementById(lov_relaciones[i].CampoHijo)
                borrarListas(donde) // borramos la LOV hija y sus subhijas...
                array = eval("lov_" + lov_relaciones[i].CampoPadre + "_" + lov_relaciones[i].CampoHijo);
                recorreExistentes = 0
                for (m = 0; m < array.length; m++) {
                    if (cual.value == array[m].CodPadre) {
                        if (recorreExistentes == 0) {
                            var nuevaOpcion = new Option();
                            donde.options[recorreExistentes] = nuevaOpcion;
                            recorreExistentes++;
                        }
                        var nuevaOpcion = new Option(array[m].TextoHijo);
                        donde.options[recorreExistentes] = nuevaOpcion;
                        donde.options[recorreExistentes].value = array[m].CodHijo;
                        recorreExistentes++;
                    }
                }
            }
        }
    }

<%if session_ncliente & "" > "" then
    cadenaDSNCP = d_lookup("dsn", "clientes", "ncliente='" & session_ncliente & "'", dsnilion)
    nempresa = session_ncliente
else
    cadenaDSNCP = session("dsn_cliente")
    nempresa = session("ncliente")
    end if

mode= enc.EncodeForHtmlAttribute(request.querystring("mode") & "")

ver= enc.EncodeForHtmlAttribute(limpiaCadena(request.querystring("ver") & ""))

    if ver= "0" then
    mode = "add"
    end if

if mode = "add" or mode = "edit" or (mode = "delete" and request.querystring("submode") = "reapertura") Then 
    set rc = server.CreateObject("ADODB.Recordset")
    c_select = "selecT cp1.nCampo, cp1.tabla, " & _
    "'campo' + ltrim(rtrim(convert(char, convert(int, right(cp1.ncampo, 2))))) campoHijo, " & _
    "'campo' + ltrim(rtrim(convert(char, convert(int, right(cp.ncampo, 2))))) campoPadre " & _
    "from CamposPerso cp with (nolock)	" & _
    "inner join CamposPerso cp1 with (nolock) on cp1.ncampo_dep = cp.ncampo and cp1.tabla_dep = cp.tabla " & _
    "where cp1.tabla = 'CLIENTES' " & _
    "  and cp1.ncampo like '" & nempresa & "%' " & _
    "  and cp1.tipo = 3 "

    rc.Open c_select, cadenaDSNCP, adUseClient, adLockReadOnly %>
    var lov_relaciones = new Array();
    <%contadorRelacion = 0
    do while not rc.eof 
    contador = 0 %>
        lov_relaciones[<%=contadorRelacion %>] = new RelacionListas('<%=rc("campoPadre")%>', '<%=rc("campoHijo")%>')

        var lov_<%=rc("campoPadre") %>_<%=rc("campoHijo") %>=new Array()

        <%query2 = "select COALESCE(NDETLISTA_DEP, 0) valorPadre, NDETLISTA valorHijo, VALOR textoHijo " & _
    "from campospersolista " & _
    "where tabla = 'CLIENTES' " & _
    "and ncampo = '" & rc("nCampo") & "' " & _
    "order by COALESCE(NDETLISTA_DEP, 0), VALOR  "
    set rc2 = server.CreateObject("ADODB.Recordset")
    rc2.Open query2, cadenaDSNCP, adUseClient, adLockReadOnly
    do while not rc2.eof %>
        lov_<%=rc("campoPadre") %>_<%=rc("campoHijo") %>[<%=contador %>] = new AgregarElementoArray('<%=rc2("valorPadre")%>', '<%=rc2("valorHijo")%>', '<%=rc2("textoHijo")%>')
            <%contador = contador + 1
    rc2.moveNext
    Loop

    contadorRelacion = contadorRelacion + 1
    rc.moveNext
    Loop
    set rc = nothing
    End if%>

        function AsignarEventoonchangeLista() {
    <%query3 = "selecT 'campo' + ltrim(rtrim(convert(char, convert(int, right(cp.ncampo, 2))))) campoPadre " & _
            "from CamposPerso cp with (nolock)	" & _
            "inner join CamposPerso cp1 with (nolock) on cp1.ncampo_dep = cp.ncampo and cp1.tabla_dep = cp.tabla  " & _
            "where cp1.tabla = 'CLIENTES' " & _
            "  and cp1.ncampo like '" & nempresa & "%' " & _
            "  and cp1.tipo = 3 "
            set rc3 = server.CreateObject("ADODB.Recordset")
            rc3.Open query3, cadenaDSNCP, adUseClient, adLockReadOnly
            do while not rc3.eof %>
                // Asignar la función externa al elemento
                document.getElementById("<%=rc3("campoPadre") %>").onchange = ControlListasDependientes;
    <%rc3.moveNext
            Loop
            set rc3 = nothing %>
}

        <%'Everilion Interface Timing%>

window.onload = function () {
            self.status = '';    
<%if mode = "add" or mode = "edit" Then %>
                ComprobarDDL();
            AsignarEventoonchangeLista();
            inicio();
<%end if
if tracetime> 0 then %>
                    StoreTiming("<%=CarpetaProduccion%>", <%=tracetime %>, "<%=enc.EncodeForJavascript(Request.QueryString("mode"))&""%>", "<%=enc.EncodeForJavascript(session("usuario"))%>", "<%=enc.EncodeForJavascript(session("ncliente"))%>", window.location.pathname);
<%end if %>
}

// *************************************************************************************
//                                  FIN LISTAS DEPENDIENTES 
// *************************************************************************************
</script>

<%
dim ver,c01,c02,c03,c04,c05,c06,c07,c08,c09,c10,c11,c12,c13,c14,c15,c16,c17,c18,c19,c20,objTAPI
objTAPI=0
ObtenerParametros("clientes")
'mode=request.querystring("mode")
'mmg:07/04/2008 Modificación Covaldroper >> Los comerciales (ver=0) sólo pueden añadir nuevos clientes
'ver=limpiaCadena(request.querystring("ver"))

'if ver="0" then
'    mode="add"
'end if%>
<!--<body bgcolor="<%=color_blau%>" onload="load()">-->
<body class="BODY_ASP">
    <%if request.querystring("viene") = "centralita" then%>
        <script language="javascript" type="text/javascript">
            SearchPage("client_lsearch.asp?mode=search&campo=<%=enc.EncodeForJavascript(limpiaCadena(request.querystring("campo")))%>&criterio=<%=enc.EncodeForJavascript(limpiaCadena(request.querystring("criterio")))%>&texto=<%=enc.EncodeForJavascript(limpiaCadena(request.querystring("texto")))%>&viene=<%=enc.EncodeForJavascript(limpiaCadena(request.querystring("viene")))%>", 1);
        </script>  
    <%end if

'*** i AMP  CARGAMOS OBJETOS PARA REALIZAR LLAMADAS TELEFONICAS Y LIBRERIA DE LLAMADAS A TAPI.  
    pushtocall= d_lookup("pushtocall", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente"))
    modcentralita = d_lookup("usuario", "modulosc_users", "ncliente='" & session("ncliente") & "' and usuario='" & session("usuario") & "' and nmodulo='"&ModCentralitas&"'", DsnIlion)
    if pushtocall>0 and modcentralita>""  then                 
        if mode="browse" or mode="save" then
            objTAPI=1
            demo=1
            if demo=0 then
                %> 
                <!--#include file="../centralita.inc" --> 
                <div style="display:none">
                    <object classid="clsid:21D6D48E-A88B-11D0-83DD-00AA003CCABD" id="TAPIOBJ"></object>
                    <object classid="clsid:E9225296-C759-11d1-A02B-00C04FB6809F" id="MAPPER"></object> 
                </div>                           
                <script language="javascript" type="text/javascript">
                    InicializarObjetosCall();
                </script>  
            <%end if
      end if 
    end if
'*** f AMP%>
<script language="javascript" type="text/javascript">
                    //Redirecciona a la opcion pulsada en la capa de navegación entre registros
                    function Navegar(destino, origen) {
                        document.clientes.action = "clientes.asp?ncliente=" + origen + "&donde=" + destino + "&mode=search";
                        document.clientes.submit();
                    }


</script>
<%'********** FUNCIONES
'****************************************************************************************************************
sub Foto(modo,ncliente)
	response.write("<iframe scrolling='No' name='marcoFoto' id='frFoto' src='clientes_imagen.asp?mode=" & modo & "&ncliente=" & ncliente & "&mf=" & mostrar_foto &"' frameborder='0' width='100%' height='240'></iframe>")
end sub

'******************************************************************
'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
function GuardarRegistro(ncliente)

' if (request.form("e_primer_ven") >= request.form("e_segundo_ven") and request.form("e_segundo_ven")>0) or (request.form("e_segundo_ven") >= request.form("e_tercer_ven") and request.form("e_tercer_ven")>0)  then
    if (clng(request.form("e_primer_ven")) >= clng(request.form("e_segundo_ven")) and clng(request.form("e_segundo_ven"))>0) or _
    (clng(request.form("e_segundo_ven")) >= clng(request.form("e_tercer_ven")) and clng(request.form("e_tercer_ven"))>0)  then%>
        <script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgDiasMal%>");
                    history.back();
                    history.back();
                    parent.botones.document.location = "clientes_bt.asp?mode=edit";
        </script>
    <%else

    	clialta = false

	    crear_cliente=1
    	if ncliente="" then
	        clialta = true
        	'Obtener el último nº de clientes de CONFIGURACION.
'		    rstAux.Open "select ncliente from configuracion where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            
            ''DGM 27/7/11 Nueva forma de obtener el numero de cliente.
            '' El contador incluye letras
            
            set connId = Server.CreateObject("ADODB.Connection")
            set commandId =  Server.CreateObject("ADODB.Command")

            connId.open session("dsn_cliente")
            commandId.ActiveConnection =connId
            commandId.CommandTimeout = 0
            commandId.CommandText="GetAlphaNumericCustomerID"
            commandId.CommandType = adCmdStoredProc 'Procedimiento Almacenado
            commandId.Parameters.Append commandId.CreateParameter("@ncompany",adVarChar,adParamInput,5,session("ncliente"))
            commandId.Parameters.Append commandId.CreateParameter("@id",adVarChar,adParamOutput,5)
            commandId.Parameters.Append commandId.CreateParameter("@idAnt",adVarChar,adParamOutput,5)
            commandId.Parameters.Append commandId.CreateParameter("@idFin",adVarChar,adParamOutput,5)
            commandId.Parameters.Append commandId.CreateParameter("@letterFin",adVarChar,adParamOutput,5)
            commandId.Parameters.Append commandId.CreateParameter("@lastId",adVarChar,adParamOutput,5)
            commandId.Parameters.Append commandId.CreateParameter("@lastLetter",adVarChar,adParamOutput,5)

            commandId.Execute,,adExecuteNoRecords
            result_ncliente=commandId.Parameters("@id").Value
            result_nclienteAnt = commandId.Parameters("@idAnt").Value
            result_Num = commandId.Parameters("@idFin").Value
            result_Letter = commandId.Parameters("@letterFin").Value
            result_NumAnt = commandId.Parameters("@idAnt").Value
            result_LetterAnt = commandId.Parameters("@lastLetter").Value

            connId.close
            set commandId=nothing
            set connId=nothing

 '   		if not rstAux.EOF then
'    			num=rstAux("ncliente")+1
'	    		num=string(5-len(cstr(num)),"0") + cstr(num)
                num = result_ncliente
                ''ricardo 20-3-2006 si el parametro bh=1 entonces se pondra el contador al primer hueco libre
'               if cstr(bh)="1" then
	                ''se actualizara justo cuando se termine de grabar el cliente
	                ''ya que si no se ha grabado el procedimiento para calcular
	                ''el hueco libre puede volver a dar el mismo y por lo tanto
	                ''no se habria actualizado el contador
'                else
	                'Actualizar el nº de cliente de CONFIGURACION.
' ACTUALIZAR!!!!!!!!!!!!!!!!!!!!	                rstAux("ncliente")=rstAux("ncliente")+1
'                end if
'			    rstAux.Update
'			    rstAux.Close
'		    else
'			    rstAux.addnew
'			    rstAux("ncliente")=1
'			    rstAux.Update
'			    rstAux.Close
'			    num=1
'			    num=string(5-len(cstr(num)),"0") + cstr(num)
'		    end if

		    ''ricardo 10-1-2005 se comprobara que no existe el ncliente segun el contador de datos de configuracion
		    
		    'DGM 27/7/11 Actualizo el numero en configuracion una vez comprobado que no existe, evitamos actualizar al numero nuevo y después volver
		    ' al anterior
		    
		    ncliente_a_buscar=session("ncliente") & num
		    rstAux.cursorlocation=3
		    rstAux.open "select ncliente from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente_a_buscar & "'",session("dsn_cliente")
		    if not rstAux.eof then
			    rstAux.close
			    'rstAux.open "update configuracion set ncliente=" & clng(result_NumAnt) &" and customer_letter ='" & result_LetterAnt &"' where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    if rstAux.state<>0 then rstAux.close
			    crear_cliente=0%>
			    <script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgCliExistRevCont%>");
                    history.back();
                    history.back();
                    parent.botones.document.location = "clientes_bt.asp?mode=add";
			    </script>
			    <%''response.end
		    else
			    rstAux.close
			    'DGM AQUI ACTUALIZABA EL CONTADOR
			    		    ' Una vez insertado el cliente, entonces actualizo para asegurarme que no se incremente y después se produzca un error.
    	    rstAux.open "update configuracion with(updlock) set ncliente=" & result_Num &" , customer_letter ='" & result_Letter &"' where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	        if rstAux.state<>0 then rstAux.close

		    end if
		    ''rstAux.close

		    if crear_cliente=1 then
			    'Crear un nuevo registro.
			    rst.AddNew
			    rst("ncliente")=session("ncliente") & num
			    'rst("ndist")=NULL
		    end if
	    end if

	    if crear_cliente=1 then
		    'Datos del domicilio de los DATOS GENERALES y los DATOS DE ENVIO
		    if ncliente="" then 'MODO AÑADIR NUEVO

			    'Abrimos la tabla de domicilios y creamos un registro nuevo para cliente
			    rstDomi.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + rst("ncliente") +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    rstDomi.AddNew
			    rstDomi("pertenece") = rst("ncliente")
			    rstDomi("tipo_domicilio") = "PRINCIPAL_CLI"
			    rstDomi("domicilio") = Nulear(request.form("domicilio"))
			    'rstDomi("cp")        = Nulear(request.form("cp"))
			    if Request.Form("poblacion")>"" and Request.Form("cp")="" then
				    cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(request.form("poblacion"),"'","''") & "'",DsnIlion)
				    if cp_aux="00000" then cp_aux=""
				    rstDomi("cp")     		=Nulear(cp_aux)
			    else
				    rstDomi("cp")     		=Request.Form("cp")
			    end if

			    TmpCP=rstDomi("cp") 'este valor se utiliza para la creacion del contacto

			    rstDomi("poblacion") = Nulear(request.form("poblacion"))
			    rstDomi("provincia") = Nulear(request.form("provincia"))
			    rstDomi("pais")      = Nulear(request.form("pais"))
			    rstDomi("codpoblacion")      = Nulear(request.form("codPoblacion"))
			    rstDomi("codprovincia")      = Nulear(request.form("codProvincia"))
			    rstDomi("codpais")      = Nulear(request.form("codPais"))
			    rstDomi("telefono")  = Nulear(request.form("telefono"))
			    rstDomi.Update
			    rst("dir_principal") = rstDomi("codigo")
			    rstDomi.Close
			    'Guardamos la direccion de delegacion caso de que exista
			    if request.form("de_domicilio") > "" then
				    'Abrimos la tabla de domicilios y creamos un registro nuevo
				    rstDomi.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + rst("ncliente") +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    rstDomi.AddNew
				    rstDomi("pertenece") = rst("ncliente")
				    rstDomi("tipo_domicilio") = "ENVIO_CLI"
				    rstDomi("domicilio") = Nulear(request.form("de_domicilio"))
				    'rstDomi("cp")        = Nulear(request.form("de_cp"))
				    if Request.Form("de_poblacion")>"" and Request.Form("de_cp")="" then
					    cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(request.form("de_poblacion"),"'","''") & "'",DsnIlion)
					    if cp_aux="00000" then cp_aux=""
					    rstDomi("cp")     		=Nulear(cp_aux)
				    else
					    rstDomi("cp")     		=Request.Form("de_cp")
				    end if
				    rstDomi("poblacion") = Nulear(request.form("de_poblacion"))
				    rstDomi("provincia") = Nulear(request.form("de_provincia"))
				    rstDomi("pais")      = Nulear(request.form("de_pais"))				    
			        rstDomi("codpoblacion")      = Nulear(request.form("de_codPoblacion"))
			        rstDomi("codprovincia")      = Nulear(request.form("de_codProvincia"))
			        rstDomi("codpais")      = Nulear(request.form("de_codPais"))
				    rstDomi("telefono")  = Nulear(request.form("de_telefono"))
				    rstDomi.Update
				    rst("dir_envio") = rstDomi("codigo")
				    
				    rstDomi.Close
			    end if
			    ' DGM Guardamos la direccion de delegación de factura caso de que exista
			    if request.Form("de_domicilioF") > "" then
			        rstDomi.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + rst("ncliente") +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    rstDomi.AddNew
				    rstDomi("pertenece") = rst("ncliente")
				    rstDomi("tipo_domicilio") = "CUST_INVOICE_AD" ' COMPROBAR NOMENCLATURA!! DGM
				    rstDomi("domicilio") = Nulear(request.form("de_domicilioF"))
				    'rstDomi("cp")        = Nulear(request.form("de_cp"))
				    if Request.Form("de_poblacionF")>"" and Request.Form("de_cpF")="" then
					    cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(request.form("de_poblacionF"),"'","''") & "'",DsnIlion)
					    if cp_aux="00000" then cp_aux=""
					    rstDomi("cp")     		=Nulear(cp_aux)
				    else
					    rstDomi("cp")     		=Request.Form("de_cpF")           
				    end if
				    rstDomi("poblacion") = Nulear(request.form("de_poblacionF"))
				    rstDomi("provincia") = Nulear(request.form("de_provinciaF"))
				    rstDomi("pais")      = Nulear(request.form("de_paisF"))				    
			        rstDomi("codpoblacion")      = Nulear(request.form("de_codPoblacionF"))
			        rstDomi("codprovincia")      = Nulear(request.form("de_codProvinciaF"))
			        rstDomi("codpais")      = Nulear(request.form("de_codPaisF"))
				    rstDomi("telefono")  = Nulear(request.form("de_telefonoF"))
				    rstDomi.Update
				    rst("invoice_address") = rstDomi("codigo")
				    rstDomi.Close
			    end if
		    else 'MODO EDITAR

			    'cogemos el anterior DirPrincipal para poder modificar el del contacto
			    TmpDirPrincipal_ant=rst("dir_principal")
			    'Abrimos la tabla de domicilios y modificamos el registro para CLIENTE
			    Seleccion="SELECT * FROM domicilios WHERE pertenece like '" & session("ncliente") & "%' and codigo ='"+cstr(null_z(rst("dir_principal")))+"'"
			    rstDomi.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

			    if not ( rstDomi("domicilio")&"" = request.form("domicilio")&"" and rstDomi("cp")&""= request.form("cp")&"" and _
				    rstDomi("poblacion")&"" = request.form("poblacion")&"" and rstDomi("provincia")&"" = request.form("provincia")&"" and _
				    rstDomi("pais")&"" = request.form("pais")&"" and rstDomi("telefono")&"" = request.form("telefono")&"") then

				    rstDomi.AddNew
				    rstDomi("pertenece") = rst("ncliente")
				    rstDomi("tipo_domicilio") = "PRINCIPAL_CLI"
				    rstDomi("domicilio") = Nulear(request.form("domicilio"))		    
			        rstDomi("codpoblacion")      = Nulear(request.form("codPoblacion"))
			        rstDomi("codprovincia")      = Nulear(request.form("codProvincia"))
			        rstDomi("codpais")      = Nulear(request.form("codPais"))
				    'rstDomi("cp")        = Nulear(request.form("cp"))
				    if Request.Form("poblacion")>"" and Request.Form("cp")="" then
					    cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(request.form("poblacion"),"'","''") & "'",DsnIlion)
					    if cp_aux="00000" then cp_aux=""
					    rstDomi("cp")     		=Nulear(cp_aux)
				    else
					    rstDomi("cp")     		=Request.Form("cp")
				    end if

				    TmpCP_Nuevo=rstDomi("cp")

				    rstDomi("poblacion") = Nulear(request.form("poblacion"))
				    rstDomi("provincia") = Nulear(request.form("provincia"))
				    rstDomi("pais")      = Nulear(request.form("pais"))				    
			        rstDomi("codpoblacion")      = Nulear(request.form("codPoblacion"))
			        rstDomi("codprovincia")      = Nulear(request.form("codProvincia"))
			        rstDomi("codpais")      = Nulear(request.form("codPais"))
				    rstDomi("telefono")  = Nulear(request.form("telefono"))
				    rstDomi.Update
				    rst("dir_principal") = rstDomi("codigo")
			    end if
			    rstDomi.Close
			    'Modificamos la direccion de envío caso de que exista

			    if request.form("de_domicilio") > "" then
				    'Abrimos la tabla de domicilios y modificamos el registro para ENVIO
				    nuevo = "false"
				    Seleccion="SELECT * FROM domicilios WHERE pertenece like '" & session("ncliente") & "%' and codigo ='" +cstr(null_z(rst("dir_envio"))) + "'"
				    rstDomi.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if not rstDomi.EOF then
					    if not ( rstDomi("domicilio")&"" = request.form("de_domicilio")&"" and _
							    rstDomi("cp")&""        = request.form("de_cp")&"" and _
							    rstDomi("poblacion")&"" = request.form("de_poblacion")&"" and _
							    rstDomi("provincia")&"" = request.form("de_provincia")&"" and _
							    rstDomi("pais")&""      = request.form("de_pais")&"" and _
							    rstDomi("telefono")&""  = request.form("de_telefono")&"" ) then
						    nuevo = "true"
					    end if
				    else
					    nuevo = "true"
				    end if
				    if nuevo = "true" then
					    rstDomi.AddNew
					    rstDomi("pertenece") = rst("ncliente")
					    rstDomi("tipo_domicilio") = "ENVIO_CLI"
					    rstDomi("domicilio") = Nulear(request.form("de_domicilio"))
					    'rstDomi("cp")        = Nulear(request.form("de_cp"))
					    if Request.Form("de_poblacion")>"" and Request.Form("de_cp")="" then
						    cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(request.form("de_poblacion"),"'","''") & "'",DsnIlion)
						    if cp_aux="00000" then cp_aux=""
						    rstDomi("cp")     		=Nulear(cp_aux)
					    else
						    rstDomi("cp")     		=Request.Form("de_cp")
					    end if
					    rstDomi("poblacion") = Nulear(request.form("de_poblacion"))
					    rstDomi("provincia") = Nulear(request.form("de_provincia"))

					    rstDomi("pais")      = Nulear(request.form("de_pais"))					    
			            rstDomi("codpoblacion")      = Nulear(request.form("de_codPoblacion"))
			            rstDomi("codprovincia")      = Nulear(request.form("de_codProvincia"))
			            rstDomi("codpais")      = Nulear(request.form("de_codPais"))
					    rstDomi("telefono")  = Nulear(request.form("de_telefono"))
					    rstDomi.Update
					    rst("dir_envio") = rstDomi("codigo")
				    end if
				    rstDomi.Close
			    end if
			    'Modificamos la dirección de envio de factura caso de que exista

			    if request.form("de_domicilioF") > "" then
				    'Abrimos la tabla de domicilios y modificamos el registro para ENVIO
				    nuevo = "false"

				    Seleccion="SELECT * FROM domicilios WHERE pertenece like '" & session("ncliente") & "%' and codigo ='" +cstr(null_z(rst("invoice_address"))) + "'"
				    rstDomi.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if not rstDomi.EOF then
					    if not ( rstDomi("domicilio")&"" = request.form("de_domicilioF")&"" and _
							    rstDomi("cp")&""        = request.form("de_cpF")&"" and _
							    rstDomi("poblacion")&"" = request.form("de_poblacionF")&"" and _
							    rstDomi("provincia")&"" = request.form("de_provinciaF")&"" and _
							    rstDomi("pais")&""      = request.form("de_paisF")&"" and _
							    rstDomi("telefono")&""  = request.form("de_telefonoF")&"" ) then
						    nuevo = "true"
					    end if
				    else
					    nuevo = "true"
				    end if
				    if nuevo = "true" then
					    rstDomi.AddNew
					    rstDomi("pertenece") = rst("ncliente")
					    rstDomi("tipo_domicilio") = "CUST_INVOICE_AD"
					    rstDomi("domicilio") = Nulear(request.form("de_domicilioF"))
					    'rstDomi("cp")        = Nulear(request.form("de_cp"))
					    if Request.Form("de_poblacionF")>"" and Request.Form("de_cpF")="" then
						    cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(request.form("de_poblacionF"),"'","''") & "'",DsnIlion)
						    if cp_aux="00000" then cp_aux=""
						    rstDomi("cp")     		=Nulear(cp_aux)
					    else
						    rstDomi("cp")     		=Request.Form("de_cpF")
					    end if
					    rstDomi("poblacion") = Nulear(request.form("de_poblacionF"))
					    rstDomi("provincia") = Nulear(request.form("de_provinciaF"))
					    rstDomi("pais")      = Nulear(request.form("de_paisF"))					    
			            rstDomi("codpoblacion")      = Nulear(request.form("de_codPoblacionF"))
			            rstDomi("codprovincia")      = Nulear(request.form("de_codProvinciaF"))
			            rstDomi("codpais")      = Nulear(request.form("de_codPaisF"))
					    rstDomi("telefono")  = Nulear(request.form("de_telefonoF"))
					    rstDomi.Update
					    rst("invoice_address") = rstDomi("codigo")
				    end if
				    rstDomi.Close
			    end if

		    end if
		    'Asignar los nuevos valores a los campos del recordset.
		    ''EBF Se añade para que por medio de un parametro de usuario no haya que introducir cif. El parametro es &nocif=0
		    if nocif="0" and Request.Form("cif")="" then
			    if num="" then num=trimcodempresa(rst("ncliente"))
			    rst("cif")           = num
			    rst("cifedi")        = num
			    CIF=num
		    else
			    rst("cif")           = Nulear(Request.Form("cif"))
			    rst("cifedi")        = CIF
		    end if

		    'cogemos el anterior razon_social para poder modificar el del contacto
		    TmpRazon_Social_Ant=rst("rsocial")

		    rst("rsocial")       = Nulear(Request.Form("rsocial"))
		    rst("ncomercial")    = Nulear(Request.Form("ncomercial"))

		    if si_asesoria=true then
			    rst("fjuridica") = Nulear(Request.Form("fjuridica"))
			    rst("titular")   = Nulear(Request.Form("titular"))
		    end if

		    'cogemos el anterior contacto para poder modificar el del contacto
		    TmpContacto_Ant=rst("contacto")

		    rst("contacto")      = Nulear(request.form("contacto"))
		    rst("web")           = Nulear(request.form("web"))

		    'cogemos el anterior email para poder modificar el del contacto
		    TmpEmail_Ant=rst("email")

		    rst("email")         = Nulear(request.form("email"))
		    rst("observaciones") = Nulear(request.form("observaciones"))
		    rst("aviso")         = Nulear(request.form("aviso"))
		    rst("falta")         = Nulear(request.form("falta"))
		    rst("fbaja")         = Nulear(request.form("fbaja"))

		    'cogemos el anterior movil para poder modificar el del contacto
		    TmpMovil_Ant=rst("telefono2")

		    rst("telefono2")     = Nulear(request.form("telefono2"))

		    'cogemos el anterior fax para poder modificar el del contacto
		    TmpFax_Ant=rst("fax")

		    rst("fax")           = Nulear(request.form("fax"))
		    'DATOS COMERCIALES
		    rst("tarifa")		 = Nulear(request.form("tarifa"))
		    rst("divisa")		 = Nulear(request.form("divisa"))

            ' >>> MCA 22/09/05 : Admitir decimales en los descuentos
            rst("dto")=null_z(Nulear(replace(request.form("descuento"),",",".")))

            rst("dto2")=null_z(Nulear(replace(request.form("descuento2"),",",".")))

            rst("dto3")=null_z(Nulear(replace(request.form("descuento3"),",",".")))

            rst("dtoLineal")=null_z(Nulear(replace(request.form("descuentoLineal"),",",".")))

		    rst("fpago")         = Nulear(request.form("fpago"))

		    rst("tpago")         = Nulear(request.form("tpago"))

		    ''cag dias pago
		    rst("primer_ven")    = Nulear(request.form("e_primer_ven"))
		    rst("segundo_ven")    = Nulear(request.form("e_segundo_ven"))
		    rst("tercer_ven")    = Nulear(request.form("e_tercer_ven"))

		    ''fin cag
		    rst("recargo")       = Null_z(request.form("recargo"))
		    rst("re")            = Nz_b(request.form("re"))
		    rst("iva")=			nulear(request.form("iva"))

		    rst("mesnopago")=  Null_z(request.form("mesNoPago"))
            col =0

           ' rst2.open "select * from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		   ' if not rst2.eof then
		'	    rst2("vencom1")       = Null_z(request.form("vencom1"))
		'	    rst2("pcom1")         = Null_z(request.form("pcom1"))
		'	    rst2("vencom2")       = Null_z(request.form("vencom2"))
		'	    rst2("pcom2")         = Null_z(request.form("pcom2"))
		'	    rst2.update
		'	    rst3.open "select * from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & request.form("distribuidor") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		'	    if not rst3.eof then
		''		    rst("ndist")=rst3("ndist")
		'	    end if
		'	    rst3.close
		 '   else
		'	    rst3.open "select * from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & request.form("distribuidor") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		'	    if not rst3.eof then
		''		    rst("ndist")=rst3("ndist")
		'	    end if
		'	    rst3.close
		 '   end if
        '
          rst2.open "select name from distribuidores with(nolock) where ndist like '" &session("ncliente")&"%' and ndist = '"& request.Form("distribuidor") & "' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
           if not rst2.eof then
              if rst2("name")&"" > "" then
                    col = 1
               else
                    col = 0
               end if
            end if
            rst2.close
		    'modificamos los valores de la tabla distribuidor
		    if col = 0 then
		        rst2.open "select * from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		    else
                rst2.open "select * from distribuidores where ndist like '" &session("ncliente")&"%' and ndist = '"&request.Form("distribuidor")&"' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic  
		    end if
		    if not rst2.eof then
			    rst2("vencom1")       = Null_z(request.form("vencom1"))
			    rst2("pcom1")         = Null_z(request.form("pcom1"))
			    rst2("vencom2")       = Null_z(request.form("vencom2"))
			    rst2("pcom2")         = Null_z(request.form("pcom2"))
			    rst2.update
			    
			    if col = 0 then
			        rst3.open "select * from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & request.form("distribuidor") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			        if not rst3.eof then
				        rst("ndist")=rst3("ndist")
			        end if
			        rst3.close
			    else
			        if request.Form("distribuidor")&"" > "" then
			            rst("ndist") = request.Form("distribuidor")
			        else
			            rst("ndist") = null
			        end if
			    end if
		    else
		        if col = 0 then
			        rst3.open "select * from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & request.form("distribuidor") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    else
			        rst3.open "select * from distribuidores with(nolock) where ndist like '" &session("ncliente")&"%' and ndist = '"&request.Form("distribuidor")&"' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic  
			    end if
			    if not rst3.eof then
			        if rst3("ndist")&"" > "" then
				        rst("ndist")=rst3("ndist")
				    else
				        rst("ndist")=null
				    end if
			    
			   ' else ' En el caso de colaboradores, que puede tener varios clientes no podemos comprobar la restriccion
			   '     rst("ndist") = request.Form("distribuidor")
			    end if
			    rst3.close
		    end if
		    
		    rst2.close
 
		    rst("ccontable")     = Nulear(request.form("ccontable"))
		    rst("ccontable_efecto")     = Nulear(request.form("ccontable_efecto"))
            rst("CCONTABLE_SUPLIDOS")     = Nulear(request.form("CCONTABLE_SUPLIDOS"))
            rst("INTRA")     = nz_b2(request.form("intra"))
            if (nz_b2(request.form("intra"))=1) then
                rst("iva")=0
            end if
		    rst("riesgo1")=null_z(Nulear(replace(request.form("rgomaxaut"),",",".")))
		    if si_tiene_modulo_EBESA = 0 then
    		    if rst("riesgo1")=0 then
	    		    rst("riesgo2")=0
		        end if
		    end if
            changedCom = false
            if si_tiene_modulo_NETTFI <> 0 then
                if rst("comercial") <> Nulear(Request.Form("comasignado")) then
                    changedCom = true
                end if
            end if
		    rst("comercial")=Nulear(request.form("comasignado"))
		    rst("agente")=Nulear(request.form("agenteasignado"))
		    'DATOS BANCARIOS
            ''response.write("el strBanco es-" & strBanco & "-<br>")
            ncuenta_nom=""
            if strIban&strPais&"">"" then
                ncuenta_nom=strIban&strPais
            else
                ncuenta_nom=left(cuenta,4)
            end if
		    rst("banco")		 = Nulear(d_lookup("entidad","bancos","codigo='" & ncuenta_nom & "'",DsnIlion))
		    rst("bancodom")      = Nulear(request.form("DomBanco"))
            p_pais=Request.Form("country")
            p_paisGB=Request.Form("countryGB")
            ''response.write("el p_pais es-" & p_pais & "-<br>")
            if p_pais & ""=""  then
                p_pais="ES"
                existe_pais=1
            else
                existe_pais=0
                set rstPais = server.CreateObject("ADODB.Recordset")
                rstPais.cursorlocation=3
                rstPais.open "select iso2 from paises with(NOLOCK) where iso2='" & p_pais & "'",dsnilion
                if not rstPais.eof then
                    existe_pais=1
                end if
                rstPais.close
                set rstPais=nothing
            end if
            if si_tiene_modulo_OrCU<>0 or si_tiene_modulo_TGB<>0 then
                if p_paisGB & ""=""  then
                    p_paisGB="ES"
                    existe_paisGB=1
                else
                    existe_paisGB=0
                    set rstPaisGB = server.CreateObject("ADODB.Recordset")
                    rstPaisGB.cursorlocation=3
                    rstPaisGB.open "select iso2 from paises with(NOLOCK) where iso2='" & p_paisGB & "'",dsnilion
                    if not rstPaisGB.eof then
                        existe_paisGB=1
                    end if
                    rstPaisGB.close
                    set rstPaisGB=nothing
                end if
            end if

            ''response.write("el existe_pais es-" & existe_pais & "-" & cuenta & "-<br>")
            if existe_pais=1 and cuenta & "" <> "" then
                p_iban=Request.Form("IBAN")
                ''response.write("el p_iban es-" & p_iban & "-<br>")
                if p_iban & ""="" then
                    set conn=  server.CreateObject("ADODB.Connection")
                    set cmd=  server.CreateObject("ADODB.Command")
                    conn.open session("dsn_cliente")
                    cmd.ActiveConnection=conn
                    conn.cursorlocation=3
                    cmd.CommandText="CalculateIBAN"
	                cmd.CommandType = adCmdStoredProc 
                    cmd.Parameters.Append cmd.CreateParameter("@banknumber", adVarChar, ,20,cuenta)
                    cmd.Parameters.Append cmd.CreateParameter("@country", adVarChar, ,20,p_pais)
                    set rstGetIban=cmd.execute
                    if not rstGetIban.eof then
                        p_iban=rstGetIban("result")
                    end if
                    conn.close
                    if rstGetIban.state<>0 then rstGetIban.close
                    set rstGetIban=nothing
                    set cmd=nothing
                    set conn=nothing
                end if
                entidadSW=""
                if cuenta & "">"" then
                    entidadSW=left(cuenta,4)
                end if
                cuenta=p_pais & p_iban & cuenta
                swift=Request.form("bic")
                ''response.write("el swift es-" & swift & "-<br>")
                if swift & ""="" then
                    set conn=  server.CreateObject("ADODB.Connection")
                    set cmd=  server.CreateObject("ADODB.Command")
                    conn.open dsnilion
                    cmd.ActiveConnection=conn
                    conn.cursorlocation=3
                    cmd.CommandText="GetDefaultDataBank"
	                cmd.CommandType = adCmdStoredProc 
                    cmd.Parameters.Append cmd.CreateParameter("@bankCode", adVarChar, , 4,entidadSW)
                    set rstGetSwift=cmd.execute
                    if not rstGetSwift.eof then
                        swift=rstGetSwift("swift_code")
                    end if
                    conn.close
                    if rstGetSwift.state<>0 then rstGetSwift.close
                    set rstGetSwift=nothing
                    set cmd=nothing
                    set conn=nothing
                end if
            end if

            if si_tiene_modulo_OrCU<>0 or si_tiene_modulo_TGB<>0 then
                if existe_paisGB=1 and cuentaTGB & "" <> "" then
                    p_ibanGB=Request.Form("IBANGB")
                    if p_ibanGB & "" = "" then
                        set conn=  server.CreateObject("ADODB.Connection")
                        set cmd=  server.CreateObject("ADODB.Command")
                        conn.open session("dsn_cliente")
                        cmd.ActiveConnection=conn
                        conn.cursorlocation=3
                        cmd.CommandText="CalculateIBAN"
	                    cmd.CommandType = adCmdStoredProc 
                        cmd.Parameters.Append cmd.CreateParameter("@banknumber", adVarChar, ,20,cuentaTGB)
                        cmd.Parameters.Append cmd.CreateParameter("@country", adVarChar, ,20,p_paisGB)
                        set rstGetIban=cmd.execute
                        if not rstGetIban.eof then
                            p_ibanGB=rstGetIban("result")
                        end if
                        conn.close
                        if rstGetIban.state<>0 then rstGetIban.close
                        set rstGetIban=nothing
                        set cmd=nothing
                        set conn=nothing
                    end if
                    cuentaTGB=p_pais & p_ibanGB & cuentaTGB
                    swift2=Request.form("bicGB")
                    NEntidadGB=request.Form("NEntidadGB")
                    if swift2 & "" = "" then
                        set conn=  server.CreateObject("ADODB.Connection")
                        set cmd=  server.CreateObject("ADODB.Command")
                        conn.open dsnilion
                        cmd.ActiveConnection=conn
                        conn.cursorlocation=3
                        cmd.CommandText="GetDefaultDataBank"
	                    cmd.CommandType = adCmdStoredProc 
                        cmd.Parameters.Append cmd.CreateParameter("@bankCode", adVarChar, , 4,NEntidadGB)
                        set rstGetSwift=cmd.execute
                        if not rstGetSwift.eof then
                            swift2=rstGetSwift("swift_code")
                        end if
                        conn.close
                        if rstGetSwift.state<>0 then rstGetSwift.close
                        set rstGetSwift=nothing
                        set cmd=nothing
                        set conn=nothing
                    end if
                end if
            end if

            if existe_pais=1 or cuenta & "" = "" then
		        rst("ncuenta")       = Nulear(cuenta)
            else
                rst("ncuenta")       = ""
            end if
            rst("swift_code")    = swift
		    rst("ntarjeta")      = Nulear(request.form("NumTarjeta"))
		    rst("fcaducidad")    = Nulear(request.form("fcaducidad"))
		    rst("domrec")        = Nz_b(request.form("Domiciliacion"))

		    rst("formatobanco")  = Nulear(request.form("formatobanco"))
		    'FLM:20/01/2009: OTROS DATOS BANCARIOS MÓDULO ORCU.
		    rst("banco2")		 = d_lookup("entidad","bancos","codigo='" & strBanco2 & "'",DsnIlion)
		    rst("bancodom2")      = Nulear(request.form("DomBanco2"))
		    rst("ncuenta2")       = cuenta2
		    rst("ntarjeta2")      = Nulear(request.form("NumTarjeta2"))
		    rst("fcaducidad2")    = Nulear(request.form("fcaducidad2"))
		    rst("domrec2")        = Nz_b(request.form("Domiciliacion2"))
		    rst("formatobanco2")   = Nulear(request.form("formatobanco2"))
		    'OTROS DATOS
		    rst("proyecto")      = Nulear(request.form("cod_proyecto")) 'jcg 02/02/2008
		    rst("tactividad")    = Nulear(request.form("tactividad"))
		    rst("zona")          = Nulear(request.form("zona"))
		    rst("transportista") = Nulear(request.form("transportista"))   
		    rst("portes")        = Nulear(request.form("portes"))
		    rst("hmanyana")      = Nulear(request.form("hmanyana"))
		    rst("htarde")        = Nulear(request.form("htarde"))
		    rst("pht")           = reemplazar(Null_z(request.form("pht")),".",",")
		    rst("pkm")           = reemplazar(Null_z(request.form("pkm")), ".", ",")
		    rst("pd")            = reemplazar(Null_z(request.form("pd")), ".", ",")

            rst("TGBBANCO")      = Nulear(cuentaTGB)
            rst("swift_code2")   = Nulear(swift2)
            rst("TGBBANCODOM")   = Nulear(request.form("DomBancoGB"))
            rst("TGBBANCOPOB")   = Nulear(request.form("PoblacionGB"))
            rst("TGBBANCOPROV")  = Nulear(request.form("codProvinciaGB"))

		    if viene="facturas_cli_E" and ncliente="" then
			    rst("tipo_cliente")=session("ncliente")&"1003"
		    else
			    rst("tipo_cliente")  = Nulear(request.form("tipo_cliente"))
		    end if

		    'DGM Si viene de colaboraciones se fuerza a que el modo sea COLABORADOR
		    if viene="collaborations" then
		        rst("tipo_cliente") = session("ncliente") & "COLAB"
		    end if
            'FIN DMG
		    if ncliente & "">"" then
			    rst("verstock")=Nz_b(request.form("verstock"))
		    end if
		    if si_asesoria=true then
			    rst("periodicidad")=Nulear(request.form("periodicidad"))
			    rst("dsn_nominaplus")=Nulear(request.form("dsn_nominaplus"))
			    rst("segsocial")      = Nulear(request.form("segsocial"))
			    rst("tienda")=Nulear(request.form("sucursal"))
			    rst("ASESORIALIST")=Nz_b2(request.form("mostrarListPortal"))
		    end if
		    if si_tiene_modulo_agrario<>0 then
			    rst("fnacimiento")    = Nulear(request.form("fnacimiento"))
			    rst("segsocial")      = Nulear(request.form("segsocial"))
			    rst("atp")            = Nz_b(request.form("atp"))
		    end if

		    if si_tiene_modulo_EBESA <> 0 then
		        rst("atp") = Nz_b(request.form("dtoimpfact"))
            end if

            'DGM Guardamos campos de comunicaciones
            
            rst("submit_advertising") = Nz_b(request.form("submit_advertising"))
            rst("email_communication") = Nz_b(request.form("email_communication"))
            rst("sms_communication") = Nz_b(request.form("sms_communication"))
                        
            
            'Fin campos comunicaciones
            
		    'ricardo 20-5-2004 actualizamos los campos personalizables

		    num_campos= d_count("NCAMPO","CAMPOSPERSO","TABLA='CLIENTES' and NCAMPO like '"&SESSION("NCLIENTE")&"%' ",session("dsn_cliente"))

		    '**RGU 17/11/2006
		    num_campos_tabla=limpiaCadena(num_campos)

		    if num_campos_tabla&"">"" then
			    redim lista_valores(num_campos_tabla)
		    else
			    redim lista_valores(0)
		    end if

            
		    '**RGU

		    if num_campos & "">"" then
			    'redim lista_valores(num_campos+5)
			    for ki=1 to num_campos
				    nom_campo="campo" & ki
				    valor_form=Nulear(limpiaCadena(request.querystring(nom_campo)))
				    if valor_form & ""="" then
					    valor_form=Nulear(limpiaCadena(request.form(nom_campo)))
				    end if
				    tipo_campo_perso=request.form("tipo_campo" & ki)

				    if tipo_campo_perso & ""="" then tipo_campo_perso=-1
				    if tipo_campo_perso=2 then
					    if valor_form="on" then
						    lista_valores(ki)=1
					    else
						    lista_valores(ki)=0
					    end if
				    else
					    lista_valores(ki)=valor_form
				    end if
			    next
		    else
			    'redim lista_valores(10+5)
			    for ki=1 to num_campos_tabla
				    lista_valores(ki)=""
			    next
		    end if
		    
            changedPartNetit = false
            changedAseFiscNetit = false
            changedAseFiscAvzNetit = false
            valorAseFiscNetit=""
            valorAseFiscNetitOld=""
            valorAseFiscAvzNetit=""
            valorAseFiscAvzNetitOld=""
		    for ki=1 to num_campos_tabla
		        'mmg: aqui actualiza los campos personalizables
			    ncampo=iif(len(ki)=1,"0"&ki,ki)
			    rst("campo"&ncampo)=lista_valores(ki)
		    next

            'mmg: OrCU
            if si_tiene_modulo_OrCU <> 0 then
                rst("Campo01")    = Nulear(request.form("dtoCli1"))
                rst("Campo02")    = Nulear(request.form("dtoCli2"))
                rst("Campo03")    = Nulear(request.form("dtoCli3"))
                'dgb 27/10/2009  Xenteo pago gasoleo
                if modulo_Xenteo <> 0 then
                     rst("pagoenpostea")  = Nz_b(request.form("pagoA"))
                     rst("pagoenposteb")  = Nz_b(request.form("pagoB"))
                end if
                
            end if
            'FLM:20090429:campo que indica si la facturación se va a agrupar por cliente o por tarjeta
		    if si_tiene_modulo_OrCU <> 0 or si_tiene_modulo_petroleos <> 0  then
		        rst("tipoagrupacionsum")=Null_z(request.form("modFactSum"))
		    end if
   		    'cag
		    if request.form("ncopiasFacturas")&"">"" then
			    rst("numcopiasFactura")=request.form("ncopiasFacturas")
		    end if
		    'fin cag

		    if si_tiene_modulo_EBESA <> 0 then
                rst("campo11")         = Nulear(request.form("tpagonp1"))
                rst("campo12")         = Nulear(request.form("tpagonp2"))
                rst("campo13")         = Nulear(request.form("tpagonp3"))
            end if
            
            SaldoMaxAnterior=""
            SaldoMaxNuevo=""
            'FLM:20090727:nuevos campos orcu , saldo. 
            if si_tiene_modulo_OrCU <> 0 then
                'si está a 1 es que se ha modificado el saldo
                if Nulear(request.form("hd_saldoEnvidado"))="1" then 
                    rst("enviado")=Date()
                    saldoMaxForm= request.form("saldomax")
                    if request.Form("cbSaldoSinLimite") = "1" then 
                        saldoMaxForm="9999999.99"
                    end if
                   ' response.Write saldoMaxForm & " " & request.Form("cbSaldoSinLimite")
                    'response.End
                    SaldoMaxAnterior=cstr(rst("saldomax")&"")
                    if Nulear(saldoMaxForm)<>"" then rst("saldomax")= replace(Nulear(saldoMaxForm),",",".") else rst("saldomax")= null end if
                    SaldoMaxNuevo=cstr(rst("saldomax")&"")
                    if Nulear(request.form("saldoact"))<>"" then rst("saldo")= replace(Nulear(request.form("saldoact")),",",".") else rst("saldo")=null end if
                    if Nulear(request.form("saldooffline"))<>"" then rst("saldooffline")= replace(Nulear(request.form("saldooffline")),",",".") else rst("saldooffline")=0 end if
                end if
            end if
            unsubsCust = false
            if rst("fbaja") & "" <> "" then
                unsubsCust = true
            end if
		    'Actualizar el registro.
		    rst.Update

''response.write("los datos de nettit son -" & si_tiene_modulo_NETTFI & "-" & changedCom & "-" & unsubsCust & "-" & changedPartNetit & "-" & changedAseFiscNetit & "-" & changedAseFiscAvzNetit & "-" & rst("ncliente") & "-<br>")
            if si_tiene_modulo_NETTFI <> 0 then
                if changedCom then
                    set connNet = Server.CreateObject("ADODB.Connection")
                    set commandNet =  Server.CreateObject("ADODB.Command")
                    connNet.open session("dsn_cliente")
                    commandNet.ActiveConnection =connNet
                    commandNet.CommandTimeout = 0
                    commandNet.CommandText="saveLastComercialChangedDate"
                    commandNet.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    commandNet.Parameters.Append commandNet.CreateParameter("@ncliente",adChar,adParamInput,10,rst("ncliente"))
                    commandNet.Execute,,adExecuteNoRecords
                    connNet.close
                    set commandNet=nothing
                    set connNet=nothing

                    set connNet = Server.CreateObject("ADODB.Connection")
                    set commandNet =  Server.CreateObject("ADODB.Command")
                    ''response.Write "ncontacto-" & rst("ncontacto") & "-"
                    connNet.open session("dsn_cliente")
                    commandNet.ActiveConnection =connNet
                    commandNet.CommandTimeout = 0
                    commandNet.CommandText="ComercialChanged"
                    commandNet.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    commandNet.Parameters.Append commandNet.CreateParameter("@codigo",adChar,adParamInput,10,rst("ncontacto"))
                    commandNet.Execute,,adExecuteNoRecords
                    connNet.close
                    set commandNet=nothing
                    set connNet=nothing

                    set connNet = Server.CreateObject("ADODB.Connection")
                    set commandNet =  Server.CreateObject("ADODB.Command")
                    ''response.Write "ncontacto-" & rst("ncontacto") & "-"
                    connNet.open session("dsn_cliente")
                    commandNet.ActiveConnection =connNet
                    commandNet.CommandTimeout = 0
                    commandNet.CommandText="AgentChangedAlarm"
                    commandNet.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    commandNet.Parameters.Append commandNet.CreateParameter("@CODIGOCONTACTOCOMERCIAL",adChar,adParamInput,20,rst("ncontacto"))
                    commandNet.Parameters.Append commandNet.CreateParameter("@carpetaProduccion",adChar,adParamInput,100,CarpetaProduccion)
                    commandNet.Parameters.Append commandNet.CreateParameter("@empresa",adChar,adParamInput,5,session("ncliente"))
                    commandNet.Execute,,adExecuteNoRecords
                    connNet.close
                    set commandNet=nothing
                    set connNet=nothing
                end if
                if unsubsCust then
		            ''ricardo 19/09/2004 se cambiara el usuario del dsncliente por el de DSNImport
		            initial_catalogC=encontrar_datos_dsn(session("dsn_cliente"),"Initial Catalog=")

		            donde=inStr(1,DSNImport,"Initial Catalog=",1)
		            donde_fin=InStr(donde,DSNImport,";",1)
		            if donde_fin=0 then
			            donde_fin=len(DSNImport)
		            end if
		            cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))

		            dsnCliente=cadena_dsn_final

                    set connNet = Server.CreateObject("ADODB.Connection")
                    set commandNet =  Server.CreateObject("ADODB.Command")
                    connNet.open dsnCliente
                    commandNet.ActiveConnection =connNet
                    commandNet.CommandTimeout = 0
                    ''Ricardo 19-09-2014 se cambia el procedimiento
                    ''commandNet.CommandText="EndClient"
                    commandNet.CommandText="unregisterNETTIClientBKO"
                    commandNet.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    commandNet.Parameters.Append commandNet.CreateParameter("@ncliente",adChar,adParamInput,10,rst("ncliente"))
                    commandNet.Execute,,adExecuteNoRecords
                    connNet.close
                    set commandNet=nothing
                    set connNet=nothing
                end if
               
''response.write("los datos son-" & changedAseFiscNetit & "-" & changedAseFiscAvzNetit & "-" & valorAseFiscNetit & "-" & valorAseFiscNetitOld & "-" & valorAseFiscAvzNetit & "-<br>")
                ''cambio de asesor fiscal

            end if

		    'JMM 2009.05.23 Copiar/Modificar Cliente en Repositorio Covaldroper
		    SyncClientesCovaldroper rst("ncliente"), ""

            ''ricardo 23-07-2013 auditamos la modificacion del saldo max
            if si_tiene_modulo_OrCU <> 0 then
                if (SaldoMaxAnterior & "")<>(SaldoMaxNuevo&"") then
                    ''Ricardo 02-08-2013 se cambia el texto de esto
                    text_anotacion=LITCLIENTES3 & " " & replace(formatnumber(null_z(SaldoMaxAnterior),ndecimales,-1,0,-1),",",".") & " " & replace(formatnumber(null_z(SaldoMaxNuevo),ndecimales,-1,0,-1),",",".") & " " & replace(formatnumber((cdbl(null_z(SaldoMaxNuevo))-cdbl(null_z(SaldoMaxAnterior))),ndecimales,-1,0,-1),",",".")
                    ''fin texto
                    set connActSal = Server.CreateObject("ADODB.Connection")
	                set commandActSal =  Server.CreateObject("ADODB.Command")

	                connActSal.open session("dsn_cliente")
	                commandActSal.ActiveConnection =connActSal
	                commandActSal.CommandTimeout = 0
	                commandActSal.CommandText="AuditCustomerMaxBalance"
	                commandActSal.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	                commandActSal.Parameters.Append commandActSal.CreateParameter("@ncompany",adVarChar,adParamInput,5,session("ncliente"))
                    commandActSal.Parameters.Append commandActSal.CreateParameter("@user",adVarChar,adParamInput,20,session("ncliente") & session("usuario"))
	                commandActSal.Parameters.Append commandActSal.CreateParameter("@ncustomer",adVarChar,adParamInput,10,rst("ncliente"))
                    commandActSal.Parameters.Append commandActSal.CreateParameter("@category",adVarChar,adParamInput,20,"Fuel Audit")
                    commandActSal.Parameters.Append commandActSal.CreateParameter("@anotacion",adLongVarChar,adParamInput,Len(text_anotacion),text_anotacion)
	                commandActSal.Execute,,adExecuteNoRecords
	                connActSal.close
	                set commandActSal=nothing
	                set connActSal=nothing
                end if
            end if

            'mmg: insertamos el cliente en la tabla ORCU_DATOS_SINCRONIZAR para que se actualice en cada aparato
            if si_tiene_modulo_OrCU <> 0 then
                if ncliente="" then
                    tc="1"
                else
                    tc="2"
                end if
                cadena="insert into ORCU_DATOS_SINCRONIZAR(NEmpresa,objeto,Id,Instalacion,TipoCambio,fecha)"
                cadena=cadena&" select '"&session("ncliente")&"',2,'"&rst("ncliente")&"',codigo,"&tc&",getdate() from tiendas where codigo like '"&session("ncliente")&"%' and cod_controlador is not null"
		        rstOrCU.Open cadena,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            end if
		    
		    'ricardo 20-4-2004
		    ''MPC 12/02/2009 Si el cliente tiene el módulo de EBESA no se ejecuta ninguna acción en caso de cambiar el riesgo máximo autorizado
		    if si_tiene_modulo_EBESA = 0 then
    		    if rcalc=1 then
	    		    CalcularRiesgo(ncliente)
		        end if
		    end if
		    ''FIN MPC

            ''ricardo 20-3-2006 si el parametro bh=1 entonces se pondra el contador al primer hueco libre
            if ncliente="" and cstr(bh)="1" then 'MODO AÑADIR NUEVO
                'Obtener el último nº de clientes de CONFIGURACION.
                rstAux.Open "select ncliente from configuracion where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

                if not rstAux.EOF then
	                result_ncliente=""
	                set conn = Server.CreateObject("ADODB.Connection")
	                set command =  Server.CreateObject("ADODB.Command")

	                conn.open session("dsn_cliente")
	                command.ActiveConnection =conn
	                command.CommandTimeout = 0
	                command.CommandText="BuscarNclienteLibre"
	                command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	                command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))
	                command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamOutput,10)
	                command.Execute,,adExecuteNoRecords
	                result_ncliente=command.Parameters("@ncliente").Value
	                conn.close
	                set command=nothing
	                set conn=nothing

	                if result_ncliente & "">"" then
		                rstAux("ncliente")=result_ncliente
		                rstAux.update
	                end if
                end if
                rstAux.close
            end if

		    'si se tiene mantenimiento y en configuración está marcado se creará un centro para el cliente
		    '********************** mantenimiento
		    'si_centros   = d_lookup("mostrar","restricciones","item='" & OBJCentros & "' and entrada='" & session("usuario") & "' and ncliente='" & session("ncliente") & "'",DSNILion)
		    si_centros = VerObjeto(OBJCentros)
		    if clialta = true then
			    if d_lookup("autocentro", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente")) then
				    CrearCentro session("ncliente") & num, false
			    end if
		    end if
		    '********************** mantenimiento
		    if ncliente="" then 'MODO AÑADIR NUEVO
			    'creamos un contacto, en caso de que exista contacto o email
			    if request.form("contacto")>"" or request.form("email")>"" or request.form("telefono2")>"" then
				    TmpNContacto=d_max("substring(ncontacto,6,10)","contactos_cli","ncontacto like '" & session("ncliente") & "%'",session("dsn_cliente"))+1
				    TmpNContacto=session("ncliente") & completar(TmpNContacto,5,"0")
				    strselect="select * from domicilios where pertenece like '" & session("ncliente") & "%' and tipo_domicilio='CONTACTO_CLI' and pertenece='" & TmpNContacto & "'"
				    rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstAux.eof then
					    rstAux.AddNew
					    rstAux("pertenece")=TmpNContacto
					    rstAux("tipo_domicilio")="CONTACTO_CLI"
				    end if
				    rstAux("domicilio")=Nulear(request.form("domicilio"))
				    rstAux("CP")=TmpCP
				    rstAux("poblacion")=Nulear(request.form("poblacion"))
				    rstAux("provincia")=Nulear(request.form("provincia"))
				    rstAux("pais") = Nulear(request.Form("pais"))
				    rstAux("codpoblacion")=Nulear(request.form("codPoblacion"))
				    rstAux("codprovincia")=Nulear(request.form("codProvincia"))
				    rstAux("codpais") = Nulear(request.Form("codPais"))
				    rstAux("telefono")=Nulear(request.form("telefono"))				    
				    rstAux.update				        				    
                    rstAux.close
				    'ahora grabamos el contacto
				    strselect="select * from contactos_cli where ncontacto like '" & session("ncliente") & "%' and ncontacto='" & TmpNContacto & "'"
				    rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstAux.eof then
					    rstAux.AddNew
					    rstAux("ncontacto")=TmpNContacto
					    rstAux("ncliente")=rst("ncliente")
					    rstAux("domicilio")=d_lookup("codigo","domicilios","pertenece='"&TmpNContacto&"' and  tipo_domicilio='CONTACTO_CLI'",session("dsn_cliente"))
					    if request.form("contacto")>"" then
						    rstAux("nombre")=Nulear(request.form("contacto"))
					    else
						    rstAux("nombre")=Nulear(Request.Form("rsocial"))
					    end if
				    end if
				    rstAux("cargo")=""
				    rstAux("movil")=Nulear(request.form("telefono2"))
				    rstAux("fax")=Nulear(request.form("fax"))
                    texto_email=Nulear(request.form("email"))
                    if texto_email & ""<>"" then
                        texto_email=mid(texto_email,1,50)
                    end if
				    rstAux("mail")=texto_email
				    rstAux.update
				    rstAux.close

                    'JMM 2009.05.23 - Copia automatica de los nuevos clientes al respositorio de Covaldroper (ALTA CON CONTACTO_CLI)
                    SyncClientesCovaldroper rst("ncliente"), TmpNContacto
			    end if
		    else 'modo modificar direccion de los conctactos que tengan los mismos datos
			    rstAux2.open "select ncontacto,nombre,mail,movil,fax from contactos_cli with(updlock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    if not rstAux2.eof then
				    'cogemos los datos de la direccion del proveedor anterior al cambio
				    strselect="select * from domicilios with(nolock) where pertenece like '" & session("ncliente") & "%' and codigo='" & TmpDirPrincipal_ant & "'"
				    rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if not rstAux.eof then
					    TmpDomicilio_Ant=rstAux("domicilio")
					    TmpCP_Ant=rstAux("cp")
					    TmpPoblacion_Ant=rstAux("poblacion")
					    TmpProvincia_Ant=rstAux("provincia")
					    TmpTelefono_Ant=rstAux("telefono")
					    set codigoDomicilio = rstAux("codigo")
				    end if
				    rstAux.close
				    while not rstAux2.eof
					    strselect="select * from domicilios with(updlock) where pertenece like '" & session("ncliente") & "%' and tipo_domicilio='CONTACTO_CLI' and pertenece='" & rstAux2("ncontacto") & "'"
					    rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					    if not rstAux.eof then
						    if null_s(TmpDomicilio_Ant)=null_s(rstAux("domicilio")) and null_s(TmpCP_Ant)=null_s(rstAux("cp")) and null_s(TmpPoblacion_Ant)=null_s(rstAux("poblacion")) and null_s(TmpProvincia_Ant)=null_s(rstAux("provincia")) and null_s(TmpTelefono_Ant)=null_s(rstAux("telefono")) then
							    rstAux("domicilio")=Nulear(request.form("domicilio"))
							    rstAux("cp")=TmpCP_Nuevo
							    rstAux("poblacion")=Nulear(request.form("poblacion"))
							    rstAux("provincia")=Nulear(request.form("provincia"))					    
				                rstAux("pais") = Nulear(request.Form("pais"))
				                rstAux("codpoblacion")=Nulear(request.form("codPoblacion"))
				                rstAux("codprovincia")=Nulear(request.form("codProvincia"))
				                rstAux("codpais") = Nulear(request.Form("codPais"))
							    rstAux("telefono")=Nulear(request.form("telefono"))
							    rstAux.update
						    end if
					    end if
					    
                        'asp repasamos por si la provincia tiene alguna coincidencia
				        if request.form("codPoblacion") = "" then
				            set conn=server.CreateObject("ADODB.Connection")
	                        set command=server.CreateObject("ADODB.Command")
	                        dsnMixta = ObtenDSNMixta(session("dsn_cliente"),DsnIlion)
	                        conn.open dsnMixta
	                        conn.cursorlocation=3
	                        command.activeConnection=conn
	                        command.CommandType = adCmdStoredProc
                            command.CommandText= "UpdateCodesAddres"
                            command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5, session("ncliente"))
                            command.Parameters.Append command.CreateParameter("@pertenece",adVarChar,adParamInput,20, ncliente)
                            command.Parameters.Append command.CreateParameter("@codigo",adInteger,adParamInput,, 0)
                            command.Parameters.Append command.CreateParameter("@poblacion",adVarChar,adParamInput,50, Nulear(request.form("poblacion")))
                            command.Parameters.Append command.CreateParameter("@provincia",adVarChar,adParamInput,50, Nulear(request.form("provincia")))
                            command.Parameters.Append command.CreateParameter("@pais",adVarChar,adParamInput,30, Nulear(request.Form("pais")))
                            command.Parameters.Append command.CreateParameter("@codpoblacionIn",adVarChar,adParamInput,50, Nulear(request.form("codpoblacion")))
                            command.Parameters.Append command.CreateParameter("@codprovinciaIn",adVarChar,adParamInput,50, Nulear(request.form("codProvincia")))
                            command.Parameters.Append command.CreateParameter("@codpaisIn",adVarChar,adParamInput,30, Nulear(request.Form("codPais")))
                            
                            command.execute
				            conn.close
                            set command=nothing
                            set conn=nothing
                       end if
                       rstAux.close
                        'dgb: 28/04/2009  solventar error al intentar insertar NULL al campo Nombre que no lo permite
					    if null_s(TmpRazon_Social_Ant)=null_s(rstAux2("nombre")) then
					        if Nulear(Request.Form("rsocial"))<>NULL then
						        rstAux2("nombre")=Request.Form("rsocial")
						    end if
					    else
					    if null_s(TmpContacto_Ant)=null_s(rstAux2("nombre")) then
					        if Nulear(request.form("contacto"))<>NULL then
						        rstAux2("nombre")=request.form("contacto")
						    end if
					    end if
					    end if
					    rstAux2.update
					    rstAux2.movenext
				    wend
				else
					'ricardo 11-2-2010 creamos un contacto, en caso de que exista contacto o email
			        if request.form("contacto")>"" or request.form("email")>"" or request.form("telefono2")>"" then
				        TmpNContacto=d_max("substring(ncontacto,6,10)","contactos_cli","ncontacto like '" & session("ncliente") & "%'",session("dsn_cliente"))+1
				        TmpNContacto=session("ncliente") & completar(TmpNContacto,5,"0")
				        strselect="select * from domicilios where pertenece like '" & session("ncliente") & "%' and tipo_domicilio='CONTACTO_CLI' and pertenece='" & TmpNContacto & "'"
				        rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				        if rstAux.eof then
					        rstAux.AddNew
					        rstAux("pertenece")=TmpNContacto
					        rstAux("tipo_domicilio")="CONTACTO_CLI"
				        end if
				        rstAux("domicilio")=Nulear(request.form("domicilio"))
				        rstAux("CP")=TmpCP
				        rstAux("poblacion")=Nulear(request.form("poblacion"))
				        rstAux("provincia")=Nulear(request.form("provincia"))					    
                        rstAux("pais") = Nulear(request.Form("pais"))
                        rstAux("codpoblacion")=Nulear(request.form("codPoblacion"))
                        rstAux("codprovincia")=Nulear(request.form("codProvincia"))
                        rstAux("codpais") = Nulear(request.Form("codPais"))				        
				        rstAux("telefono")=Nulear(request.form("telefono"))
				        set codigoDomicilio = rstAux("codigo")
				        rstAux.update
    				   
				        rstAux.close
				        'asp repasamos por si la provincia tiene alguna coincidencia
				        if request.form("codPoblacion") = "" then
				            codigoDomicilio = d_lookup("codigo","domicilios","pertenece='"&TmpNContacto&"' and  tipo_domicilio='CONTACTO_CLI'",session("dsn_cliente"))
				            set conn=server.CreateObject("ADODB.Connection")
	                        set command=server.CreateObject("ADODB.Command")
	                        dsnMixta = ObtenDSNMixta(session("dsn_cliente"),DsnIlion)
	                        conn.open dsnMixta
	                        conn.cursorlocation=3
	                        command.activeConnection=conn
	                        command.CommandType = adCmdStoredProc
                            command.CommandText= "UpdateCodesAddres"   
                            command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5, session("ncliente"))
                            command.Parameters.Append command.CreateParameter("@pertenece",adVarChar,adParamInput,20, ncliente)
                            command.Parameters.Append command.CreateParameter("@codigo",adInteger,adParamInput,, codigoDomicilio)
                            command.Parameters.Append command.CreateParameter("@poblacion",adVarChar,adParamInput,50, Nulear(request.form("poblacion")))
                            command.Parameters.Append command.CreateParameter("@provincia",adVarChar,adParamInput,50, Nulear(request.form("provincia")))
                            command.Parameters.Append command.CreateParameter("@pais",adVarChar,adParamInput,30, Nulear(request.Form("pais")))                            
                            command.Parameters.Append command.CreateParameter("@codpoblacionIn",adVarChar,adParamInput,50, Nulear(request.form("codpoblacion")))
                            command.Parameters.Append command.CreateParameter("@codprovinciaIn",adVarChar,adParamInput,50, Nulear(request.form("codProvincia")))
                            command.Parameters.Append command.CreateParameter("@codpaisIn",adVarChar,adParamInput,30, Nulear(request.Form("codPais")))
                            command.execute
				            conn.close
                            set command=nothing
                            set conn=nothing
                       end if
				        'ahora grabamos el contacto
				        strselect="select * from contactos_cli where ncontacto like '" & session("ncliente") & "%' and ncontacto='" & TmpNContacto & "'"
				        rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				        if rstAux.eof then
					        rstAux.AddNew
					        rstAux("ncontacto")=TmpNContacto
					        rstAux("ncliente")=rst("ncliente")
					        rstAux("domicilio")=d_lookup("codigo","domicilios","pertenece='"&TmpNContacto&"' and  tipo_domicilio='CONTACTO_CLI'",session("dsn_cliente"))
					        if request.form("contacto")>"" then
						        rstAux("nombre")=Nulear(request.form("contacto"))
					        else
						        rstAux("nombre")=Nulear(Request.Form("rsocial"))
					        end if
				        end if
				        rstAux("cargo")=""
				        rstAux("movil")=Nulear(request.form("telefono2"))
				        rstAux("fax")=Nulear(request.form("fax"))
				        rstAux("mail")=Nulear(request.form("email"))
				        rstAux.update
				        rstAux.close
                        'JMM 2009.05.23 - Copia automatica de los nuevos clientes al respositorio de Covaldroper (ALTA CON CONTACTO_CLI)
                        SyncClientesCovaldroper rst("ncliente"), TmpNContacto
			        end if
			    end if
		    end if

		    ''ricardo 12-3-2003
		    'PONEMOS LOS DATOS EN DOCUMENTO_CLI
		    if ncliente & ""="" then
			    ncliente_doc_cli=rst("ncliente")
		    else
			    ncliente_doc_cli=ncliente
		    end if
		    rstAux.open "select * from documentos_cli where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente_doc_cli & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		    if rstAux.eof then
			    rstAux.addnew
		    end if
		    rstAux("ncliente")=ncliente_doc_cli
		    rstAux("valorado_pre")=nz_b(request.form("valorado_pre"))
		    rstAux("serie_pre")=iif(request.form("serie_pre")>"",request.form("serie_pre"),NULL)
		    rstAux("valorado_ped")=nz_b(request.form("valorado_ped"))
		    rstAux("serie_ped")=iif(request.form("serie_ped")>"",request.form("serie_ped"),NULL)
		    rstAux("valorado_alb")=nz_b(request.form("valorado_alb"))
		    rstAux("serie_alb")=iif(request.form("serie_alb")>"",request.form("serie_alb"),NULL)
		    rstAux("serie_fac")=iif(request.form("serie_fac")>"",request.form("serie_fac"),NULL)
		    if ModuloContratado(session("ncliente"),ModPostVenta) <> 0 then
		        rstAux("serie_incidencia")=iif(request.form("serie_incidencia")>"",request.form("serie_incidencia"),NULL)
		    end if
		    rstAux.update
		    rstAux.close
		    '''''''''
	    end if

    if viene="facturas_cli_E" and mode="save" and nfactura<>"" then

        set rstCambioCliente =Server.CreateObject("ADODB.Recordset")
        rstCambioCliente.open "select observaciones,ncliente from facturas_cli with(nolock) where nfactura='"&nfactura&"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        if not rstCambioCliente.eof then
            observ = rstCambioCliente("observaciones")
            clienteAnt=rstCambioCliente("ncliente")
        end if
        rstCambioCliente.close

        'SE ACTUALIZAN LAS OBSERVACIONES SI CAMBIAMOS EL CLIENTE SIEMPRE Y CUANDO NO SEA CONVERTIR UN TICKET EN FACTURA
        strSetUpdate=""
        if clienteAnt<>session("ncliente") & "00000" and clienteAnt<>ncliente_doc_cli then
            rstCambioCliente.open "select rsocial from clientes with(nolock) where ncliente='" & clienteAnt & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            if not rstCambioCliente.eof then
                nombre=rstCambioCliente("rsocial")
            end if
            rstCambioCliente.close
            if obs="si" then
                strSetUpdate=", observaciones='"&observ & "' + char(13) + 'CAMBIO EL CLIENTE EL " &now()&"' + char(13) + 'CLIENTE ANTERIOR "&trimCodEmpresa(clienteAnt)&" - " &nombre & "'"
            end if
        end if

        rstCambioCliente.open "update facturas_cli with(updlock) set ncliente='" & ncliente_doc_cli & "'" & strSetUpdate & " where nfactura='" & nfactura & "'", session("dsn_cliente")
        rstCambioCliente.open "select * from facturas_cli where nfactura='"&nfactura&"'",session("dsn_cliente")
        Dom = Domicilios("VENTAS","FAC_ENV_CLI",ncliente,rstCambioCliente)

        rstCambioCliente.update
        rstCambioCliente.close%>
        <script language=javascript type="text/javascript">
                    window.top.opener.parent.pantalla.document.facturas_cli_E.action = "Facturas_cli_E.asp?mode=browse&nfactura=<%=enc.EncodeForJavascript(nfactura)%>";
                    window.top.opener.parent.pantalla.document.facturas_cli_E.submit();
                    window.top.opener.parent.botones.document.opciones.action = "Facturas_cli_bt_E.asp?mode=browse";
                    window.top.opener.parent.botones.document.opciones.submit();
                    parent.close();
        </script>
        <%set rstCambioCliente = Nothing
        Response.End
    end if

    if viene="facturas_cli_E" and mode="save" then%>
        
                <script language=javascript>
                    if (window.top.opener.parent.frames[0].name == "pantalla") {
                        window.top.opener.parent.pantalla.document.facturas_cli_E.action = "Facturas_cli_E.asp?mode=savefactebesa&nfactura=<%=enc.EncodeForJavascript(resultado)%>&pestanya=obs&vendedor=<%=enc.EncodeForJavascript(vendedor)%>&nclienteebesa=<%=enc.EncodeForJavascript(ncliente_doc_cli)%>";
                        window.top.opener.parent.pantalla.document.facturas_cli_E.submit();
                        window.top.opener.parent.botones.document.opciones.action = "Facturas_cli_bt_E.asp?mode=browse";
                        window.top.opener.parent.botones.document.opciones.submit();
                        parent.close();
                    } else {
                        window.top.opener.parent.parent.pantalla.document.facturas_cli_E.action = "Facturas_cli_E.asp?mode=savefactebesa&nfactura=<%=enc.EncodeForJavascript(resultado)%>&pestanya=obs&vendedor=<%=enc.EncodeForJavascript(vendedor)%>&nclienteebesa=<%=enc.EncodeForJavascript(ncliente_doc_cli)%>";
                        window.top.opener.parent.parent.pantalla.document.facturas_cli_E.submit();
                        window.top.opener.parent.parent.botones.document.opciones.action = "Facturas_cli_bt_E.asp?mode=browse";
                        window.top.opener.parent.parent.botones.document.opciones.submit();
                        parent.close();
                    }
                </script>
            <%response.End
    end if

    GuardarRegistro=crear_cliente
     end if
end function

function CadenaBusqueda(campo,criterio,texto,vienecomercial,agente)
   if texto > "" then
	  texto=replace(texto,"'","''")
	  select case criterio
		  case "contiene"
			  CadenaBusqueda=" and " + campo + " like '%" + texto + "%' "
  		  case "empieza"
			  CadenaBusqueda=" and " + campo + " like '" + texto + "%' "
		  case "termina"
			  CadenaBusqueda=" and " + campo + " like '%" + texto + "' "
		  case "igual"
			  CadenaBusqueda=" and " + campo + "='" + texto + "' "
	  end select
   end if
end function

function CadenaBusquedaMM(criterio,texto,nt)
   if texto > "" then
	  texto=replace(texto,"'","''")
	  select case nt
	    case 1
	        var="f.valor"
	    case 2
	        var="f.valor"
	    case 3
	        var="cper.valor"	
	    case 4
	        var="f.valor" 
	    case 5
	        var="f.valor"  
	  end select
	  
	  select case criterio
		  case "contiene"
			  CadenaBusquedaMM=" and "+var+" like '%" + texto + "%' "
  		  case "empieza"
			  CadenaBusquedaMM=" and "+var+" like '" + texto + "%' "
		  case "termina"
			  CadenaBusquedaMM=" and "+var+" like '%" + texto + "' "
		  case "igual"
			  CadenaBusquedaMM=" and "+var+" ='" + texto + "' "
	  end select
	  
   end if
end function
'Botones de navegación para las búsquedas.
''sub NextPrev(lote,lotes,campo,criterio,texto,pos)
sub NextPrev(lote,firstReg,lastReg,campo,criterio,texto,sentido,firstRegAll,lastRegAll)%>
<table width='100%' border='0' cellspacing="1" cellpadding="1">
	<tr><td class='MAS'>
	<%
		if firstReg<>firstRegAll then
			if (sentido="next" or sentido="last" or (sentido="prev" and lote<>1) ) then%>
				<a class='CELDAREFB7' href="javascript:Mas('first','<%=enc.EncodeForJavascript(lote)%>','<%=enc.EncodeForJavascript(firstReg)%>','<%=enc.EncodeForJavascript(lastReg)%>','<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(firstRegAll)%>','<%=enc.EncodeForJavascript(lastRegAll)%>');">|<img src="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> border="0" alt="<%=LitPrimero%>" title="<%=LitPrimero%>"/></a>
			<%end if

			if (sentido="next" or sentido="last" or (sentido="prev" and lote<>1) ) then%>
				<a class='CELDAREFB7' href="javascript:Mas('prev','<%=enc.EncodeForJavascript(lote)%>','<%=enc.EncodeForJavascript(firstReg)%>','<%=enc.EncodeForJavascript(lastReg)%>','<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(firstRegAll)%>','<%=enc.EncodeForJavascript(lastRegAll)%>');"><img src="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a>
			<%end if
		end if
		textopag=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)
        ''ricardo 31-5-2006 ya no se puede poner en que pagina nos encontramos ya que no sabemos cuantas tenemos ni en cual estamos%>
		<font class='CELDA'> <%=LitPaginacion%> </font>

		<%''if lote<lotes then
		if lastReg<>lastRegAll then
			if sentido<>"last" then%>
				<a class='CELDAREFB7' href="javascript:Mas('next','<%=enc.EncodeForJavascript(lote)%>','<%=enc.EncodeForJavascript(firstReg)%>','<%=enc.EncodeForJavascript(lastReg)%>','<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(firstRegAll)%>','<%=enc.EncodeForJavascript(lastRegAll)%>');"><img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a>
			<%end if
			if sentido<>"last" then%>
				<a class='CELDAREFB7' href="javascript:Mas('last','<%=enc.EncodeForJavascript(lote)%>','<%=enc.EncodeForJavascript(firstReg)%>','<%=enc.EncodeForJavascript(lastReg)%>','<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(firstRegAll)%>','<%=enc.EncodeForJavascript(lastRegAll)%>');"><img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> border="0" alt="<%=LitUltimo%>" title="<%=LitUltimo%>"/>|</a>
			<%end if
		end if%>
	</td></tr>
</table>
<%end sub

'Elimina los datos del registro cuando se pulsa BORRAR.
function BorrarRegistro(ncliente)
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")

    conn.open session("dsn_cliente")
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="BorrarCliente"
    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    command.Parameters.Append command.CreateParameter("@nempresa", adVarChar, adParamInput, 5, session("ncliente"))
    command.Parameters.Append command.CreateParameter("@ncliente", adVarChar, adParamInput, 10, ncliente)
    command.Parameters.Append command.CreateParameter("@si_tiene_modulo_importaciones", adInteger,adParamInput, 10, iif(si_tiene_modulo_importaciones<>0, 1, 0))
    command.Parameters.Append command.CreateParameter("@bh", adVarChar, adParamInput, 10, cstr(bh))
    command.Parameters.Append command.CreateParameter("@error", adInteger, adParamOutput, 10)
    command.Execute,,adExecuteNoRecords
    error=command.Parameters("@error").Value
    conn.close
    set command=nothing
    set conn=nothing

    if error = 1 then
		mens_err=LitNoPueBorrClienPorDocAsocDoc1
		mens_err2=LitNoPueBorrClienPorDocAsocMot1
    elseif error = 2 then
        mens_err=LitNoPueBorrClienPorDocAsocDoc2
		mens_err2=LitNoPueBorrClienPorDocAsocMot2
    elseif error = 3 then
        mens_err=LitNoPueBorrClienPorDocAsocDoc3
		mens_err2=LitNoPueBorrClienPorDocAsocMot3
	elseif error = 4 then%>
	    <script language="javascript" type="text/javascript">
                    alert("<%=LitNoPueBorrClienPorDocAsocDoc4%>");
		</script>
	<%elseif error = 5 then
	    mens_err=LitNoPueBorrClienPorDocAsocDoc5
		mens_err2=LitNoPueBorrClienPorDocAsocMot5
	elseif error = 6 then
	    mens_err=LitNoPueBorrClienPorDocAsocDoc6
		mens_err2=LitNoPueBorrClienPorDocAsocMot6
	elseif error = 13 then
	    mens_err=LitNoPueBorrClienPorDocAsocDoc13
		mens_err2=LitNoPueBorrClienPorDocAsocMot13
	elseif error = 7 then
	    mens_err=LitNoPueBorrClienPorDocAsocDoc7
		mens_err2=LitNoPueBorrClienPorDocAsocMot7
	elseif error = 12 then
	    mens_err=LitNoPueBorrClienPorDocAsocDoc12
		mens_err2=LitNoPueBorrClienPorDocAsocMot12
	elseif error = 14 then
        mens_err=LitNoPueBorrClienPorDocAsocDoc14
		mens_err2=LitNoPueBorrClienPorDocAsocMot13
    elseif error = 16 then
        mens_err=LITSTORES
		mens_err2=LITCLI
        mens_err3=LITERRDELETESTORES
    end if

	if error = 0 then
		BorrarRegistro=""
	else
		if mens_err="" then mens_err=LitNoPueBorrClienPorDocAsocDoc8
		if mens_err2="" then mens_err2=LitNoPueBorrClienPorDocAsocMot8
        if mens_err3="" then mens_err3=LitNoPueBorrClienPorDocAsoc11
        %>
		<script language="javascript" type="text/javascript">
                    alert("<%=LitNoPueBorrClienPorDocAsoc9%><%=mens_err2%><%=LitNoPueBorrClienPorDocAsoc10%><%=mens_err%><%=mens_err3%>");
		</script>
		<%BorrarRegistro=ncliente
	end if
	'*****************
end function

'Crea la tabla que contiene la barra de grupos de datos (Generales,Comerciales,etc)
sub BarraNavegacion(modo)%>
	
		   <%if viene<>"facturas_cli_E" then
           'DC
           'DB
           %>
           <script language="javascript" type="text/javascript">
               jQuery("#<%="S_"&modo & "DC"%>").show();
               jQuery("#<%="S_"&modo & "DB"%>").show();
           </script>
           <%else %>
           <script language="javascript" type="text/javascript">

               jQuery("#<%="S_"&modo & "DC"%>").hide();
               jQuery("#<%="S_"&modo & "DB"%>").hide();

           </script>
		   <%end if
		   if viene<>"facturas_cli_E" then
           'DE  cerrada
           %>
           <script language="javascript" type="text/javascript">           
               jQuery("#<%=enc.EncodeForJavascript(modo) & "DE"%>").attr("style", "display:none");


           </script>
		   <%else
           
           'DE ABIERTA
           %>
		   <script language="javascript" type="text/javascript">
               jQuery("#<%=enc.EncodeForJavascript(modo) & "DE"%>").attr("style", "display:run-in");

           </script>
		   <%end if
		   if viene<>"facturas_cli_E" then
           'OD
           'DD
           %>
            <script language="javascript" type="text/javascript">
                jQuery("#<%="S_"&modo & "OD"%>").show();
                jQuery("#<%="S_"&modo & "DD"%>").show();
           </script>
           
		   
		    <%if si_campo_personalizables=1 then
            'CP
            %>
              <script language="javascript" type="text/javascript">
                  jQuery("#<%="S_"&modo & "CP"%>").show();

           </script>
            <%else %>
               <script language="javascript" type="text/javascript">
                   jQuery("#<%="S_"&modo & "CP"%>").hide();

           </script>
		   
		    <%end if
        else %>
         <script language="javascript" type="text/javascript">
                   jQuery("#<%="S_"&modo & "OD"%>").hide();
                   jQuery("#<%="S_"&modo & "DD"%>").hide();
           </script>
           
		<%end if%>
        
 <%end sub




'****************************************************************************************************************
'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
function GuardarProveedor(nproveedor,ncliente)
	continuar=0
	strselect="select * from clientes where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'"
	rst4.cursorlocation=3
	rst4.open strselect,session("dsn_cliente")
	if not rst4.eof then
		if nproveedor & ""="" then
			'Obtener el último nº de proveedores de CONFIGURACION.
			rstAux.Open "select nproveedor from configuracion where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if not rstAux.eof then
				num=rstAux("nproveedor")+1
				num=string(5-len(cstr(num)),"0") + cstr(num)

				'Actualizar el nº de proveedor de CONFIGURACION.
				rstAux("nproveedor")=rstAux("nproveedor")+1
				rstAux.Update
				rstAux.Close
			else
				rstAux.addnew
				rstAux("nproveedor")=1
				rstAux.Update
				rstAux.Close
				num=1
				num=string(5-len(cstr(num)),"0") + cstr(num)
			end if

            ''ricardo 10-3-2009 se pone el top 1 para que esta consulta no tarde mucho
			strselect="select top 1 * from proveedores where nproveedor like '" & session("ncliente") & "%'"
			rst3.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			'Crear un nuevo registro.
			rst3.AddNew
			rst3("nproveedor")=session("ncliente") & num
			continuar=1
		else
			strselect="select * from proveedores where nproveedor like '" & session("ncliente") & "%' and nproveedor='" & nproveedor & "'"
			rst3.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if not rst3.eof then
				continuar=1
			end if
		end if
		if continuar=1 then
			continuar=0
			'Asignar los nuevos valores a los campos del recordset.
			'DATOS GENERALES
			rst3("cif")           = Nulear(rst4("cif"))
			rst3("cifedi")        = Nulear(rst4("cifedi"))
			rst3("razon_social")  = Nulear(rst4("rsocial"))
			if nproveedor="" then
				rst3("falta")   		 = day(date) & "/" & month(date) & "/" & year(date)
				rst3("fbaja")  		 = NULL
			else
				'rst3("falta")=rst4("falta")
				'rst3("fbaja")=rst4("fbaja")
			end if
			rst3("nombre")        = Nulear(rst4("ncomercial"))
			rst3("contacto")      = Nulear(rst4("contacto"))
			rst3("web")           = Nulear(rst4("web"))
			rst3("email")         = Nulear(rst4("email"))
			rst3("observaciones") = Nulear(rst4("observaciones"))
			rst3("telefono2")     = Nulear(rst4("telefono2"))
			rst3("fax")           = Nulear(rst4("fax"))
			'DATOS COMERCIALES
			rst3("divisa") = iif(rst("divisa")>"",Nulear(rst("divisa")),d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base <>0",session("dsn_cliente")))

			'DATOS BANCARIOS
			'Datos del domicilio de los DATOS GENERALES y los DATOS DE ENVIO

			if Nulear(rst4("dir_principal"))>"" then
				'Abrimos la tabla de domicilios y modificamos el registro para CLIENTE
				Seleccion="SELECT * FROM domicilios WHERE pertenece like '" & session("ncliente") & "%' and codigo ='"+cstr(null_z(rst4("dir_principal")))+"'"
				rstDomi.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if not rstDomi.eof then
					if nproveedor & ""="" then
						'Abrimos la tabla de domicilios y modificamos el registro para PROVEEDOR
						rstDomi2.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + rst3("nproveedor") +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						rstDomi2.AddNew
						rstDomi2("pertenece") = rst3("nproveedor")
						rstDomi2("tipo_domicilio") = "PRINCIPAL_PROV"
						continuar=1
					else
						'Abrimos la tabla de domicilios y modificamos el registro para PROVEEDOR
						Seleccion="SELECT * FROM domicilios WHERE pertenece like '" & session("ncliente") & "%' and codigo ='"+cstr(null_z(rst3("dir_principal")))+"'"
						rstDomi2.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						if not rstDomi2.eof then
							continuar=1
						end if
					end if
					if continuar=1 then
						rstDomi2("domicilio") = Nulear(rstDomi("domicilio"))
						'rstDomi2("cp")        = Nulear(rstDomi("cp"))
						if Nulear(rstDomi("poblacion"))>"" and Nulear(rstDomi("cp"))="" then
							cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(Nulear(rstDomi("poblacion")),"'","''") & "'",DsnIlion)
							if cp_aux="00000" then cp_aux=""
							rstDomi2("cp")     		=Nulear(cp_aux)
						else
							rstDomi2("cp")     		=Nulear(rstDomi("cp"))
						end if
						rstDomi2("poblacion") = Nulear(rstDomi("poblacion"))
						rstDomi2("provincia") = Nulear(rstDomi("provincia"))
						rstDomi2("pais")      = Nulear(rstDomi("pais"))
						rstDomi2("telefono")  = Nulear(rstDomi("telefono"))
						rstDomi2.Update
						rst3("dir_principal")=Nulear(rstDomi2("codigo"))
					end if
					rstDomi2.Close
				end if
				rstDomi.Close
			end if
			continuar=0
			'Modificamos la direccion de envío caso de que exista
			if rst4("dir_envio")&"">"" then
				'Abrimos la tabla de domicilios para CLIENTE
				Seleccion="SELECT * FROM domicilios WHERE pertenece like '" & session("ncliente") & "%' and codigo ='"+cstr(null_z(rst4("dir_envio")))+"'"
				rstDomi.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if not rstDomi.eof then
					if rstDomi2.state<>0 then rstDomi2.close
					if nproveedor & ""="" or rst3("dir_envio") & ""="" then
						rstDomi2.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + rst3("nproveedor") +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						rstDomi2.AddNew
						rstDomi2("pertenece") = rst3("nproveedor")
						rstDomi2("tipo_domicilio") = "ENVIO_PROV"
						continuar=1
					else
						'Abrimos la tabla de domicilios y modificamos el registro para PROVEEDOR
						Seleccion="SELECT * FROM domicilios WHERE pertenece like '" & session("ncliente") & "%' and codigo ='"+cstr(null_z(rst3("dir_envio")))+"'"
						rstDomi2.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						if not rstDomi2.eof then
							continuar=1
						end if
					end if
					if continuar=1 then
						if rst3("dir_envio")>"" and nproveedor>"" then
							rstDomi2("domicilio") = Nulear(rstDomi("domicilio"))
							if Nulear(rstDomi("poblacion"))>"" and Nulear(rstDomi("cp"))="" then
								cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(Nulear(rstDomi("poblacion")),"'","''") & "'",DsnIlion)
								if cp_aux="00000" then cp_aux=""
								rstDomi2("cp")     		=Nulear(cp_aux)
							else
								rstDomi2("cp")     		=Nulear(rstDomi("cp"))
							end if
							rstDomi2("poblacion") = Nulear(rstDomi("poblacion"))
							rstDomi2("provincia") = Nulear(rstDomi("provincia"))
							rstDomi2("pais")      = Nulear(rstDomi("pais"))
							rstDomi2("telefono")  = Nulear(rstDomi("telefono"))
							rstDomi2.Update
						else
							'Abrimos la tabla de domicilios y creamos un registro nuevo
							rstDomi2("pertenece") = rst3("nproveedor")
							rstDomi2("tipo_domicilio") = "ENVIO_PROV"
							rstDomi2("domicilio") = Nulear(rstDomi("domicilio"))
							if Nulear(rstDomi("poblacion"))>"" and Nulear(rstDomi("cp"))="" then
								cp_aux=d_lookup("cod_postal","poblaciones","poblacion='" & replace(Nulear(rstDomi("poblacion")),"'","''") & "'",DsnIlion)
								if cp_aux="00000" then cp_aux=""
								rstDomi2("cp")     		=Nulear(cp_aux)
							else
								rstDomi2("cp")     		=Nulear(rstDomi("cp"))
							end if
							rstDomi2("poblacion") = Nulear(rstDomi("poblacion"))
							rstDomi2("provincia") = Nulear(rstDomi("provincia"))
							rstDomi2("pais")      = Nulear(rstDomi("pais"))
							rstDomi2("telefono")  = Nulear(rstDomi("telefono"))
							rstDomi2.Update
							rst3("dir_envio") = Nulear(rstDomi2("codigo"))
						end if
					end if
					rstDomi2.Close
				end if
				rstDomi.Close
			end if
			rst3.update
			nproveedor=rst3("nproveedor")
		end if
		rst3.close
	end if
	rst4.close

	GuardarProveedor=nproveedor
end function

'ega 15/04/2008 quita las comillas dobles, simples y los salto de linea del texto
function LimpiarTexto(texto)
    texto = Replace(texto, chr(34), "")'quitar comillas dobles
    texto = Replace(texto, chr(39), "")'quitar comillas simples
    texto = Replace(texto, chr(10), "")'quitar salto de linea
    texto = Replace(texto, chr(13), "")'quitar salto de carro
    LimpiarTexto=texto
end function

'**************************************** crea un centro a partir del cliente
sub CrearCentro (p_ncliente, p_mostrar)
	rst4.Open "select * from clientes where ncliente like '" & session("ncliente") & "%' and ncliente='" & p_ncliente & "'" ,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if not rst4.eof then
		'obtenemos el nº de centro
		rstAux.Open "select ncentro from configuracion where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

		if not rstAux.EOF then
			num=rstAux("ncentro")+1
			num=session("ncliente") & string(5-len(cstr(num)),"0") + cstr(num)
			'Actualizar el nº de centro de CONFIGURACION.
			rstAux("ncentro")=rstAux("ncentro")+1
			rstAux.Update
			rstAux.Close
		else
			rstAux.addnew
			rstAux("ncentro")=1
			rstAux.Update
			rstAux.Close
			num=1
			num=session("ncliente") & string(5-len(cstr(num)),"0") + cstr(num)
		end if

		'insertamos registro en domicilios para domicilio principal
		rstAux.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + rst4("ncliente") +"' and tipo_domicilio='PRINCIPAL_CLI' order by codigo desc",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if not rstAux.eof then
			rstDomi.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + num +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rstDomi.AddNew
			rstDomi("pertenece") = num
			rstDomi("tipo_domicilio") = "PRINCIPAL_CTRO"
			rstDomi("domicilio")      = rstAux("domicilio")
	      rstDomi("cp")             = rstAux("cp")
			rstDomi("poblacion")      = rstAux("poblacion")
			rstDomi("provincia")      = rstAux("provincia")
			rstDomi("pais")           = rstAux("pais")
			rstDomi("telefono")       = rstAux("telefono")
			rstDomi.Update
         dir_principal             = rstDomi("codigo")
         rstDomi.Close
		end if
		rstAux.close


		'insertamos registro en domicilios para direccion de envio
		rstAux.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + rst4("ncliente") +"' and tipo_domicilio='ENVIO_CLI' order by codigo desc",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		dir_envio=""
		if not rstAux.eof then
			rstDomi.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and pertenece='" + num +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rstDomi.AddNew
			rstDomi("pertenece") = num
			rstDomi("tipo_domicilio") = "ENVIO_CTRO"
			rstDomi("domicilio")      = rstAux("domicilio")
	      	rstDomi("cp")             = rstAux("cp")
			rstDomi("poblacion")      = rstAux("poblacion")
			rstDomi("provincia")      = rstAux("provincia")
			rstDomi("pais")           = rstAux("pais")
			rstDomi("telefono")       = rstAux("telefono")
			rstDomi.Update
         dir_envio                 = rstDomi("codigo")
         rstDomi.Close
		end if
		rstAux.close

		'creamos registro para centro
		rstAux.open "select * from centros where ncentro like '" & session("ncliente") & "%' and ncentro='" & num & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rstAux.addnew
		rstAux("ncentro")      = num
		rstAux("ncliente")      = rst4("ncliente")
		rstAux("cif")           = rst4("cif")
		rstAux("rsocial")       = rst4("rsocial")
		rstAux("ncomercial")    = rst4("ncomercial")
		rstAux("contacto")      = rst4("contacto")
		rstAux("hmanyana")      = rst4("hmanyana")
		rstAux("htarde")        = rst4("htarde")
		rstAux("pht")           = null_z(rst4("pht"))
		rstAux("pkm")           = null_z(rst4("pkm"))
		rstAux("pd")            = null_z(rst4("pd"))
		rstAux("zona")          = rst4("zona")
		rstAux("falta")         = day(date) & "/" & month(date) & "/" & year(date)
		rstAux("movil")     		= rst4("telefono2")
		rstAux("dir_principal") = dir_principal
		rstAux("tipo_pago")     = rst4("tpago")
		rstAux("comercial")     = rst4("comercial")
		if dir_envio>"" then
	   	rstAux("dir_envio")  = dir_envio
		end if
		rstAux("codcliente")    = rst4("ncliente")
		rstAux.update
		rstAux.close

        if p_mostrar = true then%>
			<script language="javascript" type="text/javascript">
                alert("<%=LitCentroCreado%><%=enc.EncodeForJavascript(trimCodEmpresa(num))%>")
            </script>
		<%end if
	else%>
		<script language="javascript" type="text/javascript">
                alert("<%=LitErrorConversion%>")
        </script>
	<%end if
	rst4.close
end sub

'****************************************************************************************************************
'********** CODIGO PRINCIPAL DE LA PÁGINA ***********************************************************************
'****************************************************************************************************************
const borde=0

%>    
	<form name="clientes" method="post" >


    	    <%WaitBoxOculto LitEsperePorFavor

        result_validarhacienda= d_lookup("validarhacienda", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente"))
        %>
        <input type="hidden" name="validarhacienda" value="<%=enc.EncodeForHtmlAttribute(result_validarhacienda & "")%>"/> 
        <%


	    'Leer parámetros de la página
		'mode     = Request.QueryString("mode")
		ncliente = enc.EncodeForHtmlAttribute(limpiaCadena(Request.QueryString("ncliente") & ""))
		if limpiaCadena(Request.QueryString("ncliente"))&""="" then
		    ncliente = enc.EncodeForHtmlAttribute(limpiaCadena(Request.Form("ncliente") & ""))
		end if
		campo    = enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("campo") & ""))
		criterio = enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("criterio") & ""))
		texto    = enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("texto") & ""))

		if request.QueryString("verd")>"" then
			verd=enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("verd") & ""))
		else
			verd=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("verd") & ""))
		end if

		if request.QueryString("verl")>"" then
			verl=enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("verl") & ""))
		else
			verl=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("verl") & ""))
		end if

		if viene="lista_agentes" then
			ncliente =enc.EncodeForHtmlAttribute(limpiaCadena(request.querystring("ndoc") & ""))
		end if

		obs = enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("obs") & ""))
        if obs &""="" then obs= enc.EncodeForHtmlAttribute(request.Form("obs") & "")

        '' MPC 08/10/2008 Lectura del parámetro cifrepe
        repe=enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("repe") & ""))

        '' MPC 25/02/2014 Lectura del parámetro hide
        if hide & "" = "" then hide=enc.EncodeForHtmlAttribute(limpiaCadena(request.Form("hide") & ""))

		if viene="facturas_cli_E" then
		    vendedor=enc.EncodeForHtmlAttribute(limpiaCadena(request.querystring("ndoc") & ""))
		    if vendedor&""="" then vendedor=enc.EncodeForHtmlAttribute(Request.Form("vendedor") & "")
		    nfactura=enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("nfactura") & ""))
		    if nfactura&""="" then nfactura=enc.EncodeForHtmlAttribute(request.Form("nfactura") & "")
		end if
		accion=enc.EncodeForHtmlAttribute(limpiaCadena(Request.QueryString("accion") & ""))
		'No se pasa el checkCadena cuando el parámetro ncliente contiene LitCliente, lo que significa que viene'
		'de la función hiperv y que se está intentando añadir un nuevo cliente.'
		if ncliente & "">"" and ncliente<>LitCliente then
			CheckCadena ncliente
		end if

		''cag
		p_e_segundo_ven=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("e_segundo_ven") & ""))
		p_e_tercer_ven=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("e_tercer_ven") & ""))

		''fin cag
		
		''dgb: 27/10/2009  XENTEO
		'modulo_Xenteo=d_lookup("xenteo","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")) '**RGU 6/11/2009: lo comento para obtenerlo mas tarde
		'Conexion y cursores
		set rst = Server.CreateObject("ADODB.Recordset")
		set rst2 = Server.CreateObject("ADODB.Recordset")
		set rst3 = Server.CreateObject("ADODB.Recordset")
		set rst4 = Server.CreateObject("ADODB.Recordset")
		set rstAux = Server.CreateObject("ADODB.Recordset")
		set rstAux2 = Server.CreateObject("ADODB.Recordset")
        set rstAux3 = Server.CreateObject("ADODB.Recordset")
		set rstSelect = Server.CreateObject("ADODB.Recordset")
		set rstDomi = Server.CreateObject("ADODB.Recordset")
		set rstDomi2 = Server.CreateObject("ADODB.Recordset")
		set rstComer = Server.CreateObject("ADODB.Recordset")
		set rstCom=Server.CreateObject("ADODB.Recordset")
        set rstCF = Server.CreateObject("ADODB.Recordset")
		'**RGU 17/11/2006
		set rstCP =Server.CreateObject("ADODB.Recordset")
		'**RGU

        confirmChange = false
        if mode = "browse" and valchanges=1 then
            rstCF.open "select COMPANY_FACT from ILIONTECA_CUSTOMERS with(nolock) where COMPANY_FACT = '" & session("ncliente") & "' and CLI_AFACT = '" & trimcodempresa(ncliente) & "'", DSNILION,adOpenKeyset,adLockOptimistic
            if not rstCF.eof then
                rstCF.close
                rstCF.open "select NCUSTOMER from TMP_IliontecaCustomers with(nolock) where NCUSTOMER = '" & ncliente & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                if not rstCF.eof then
                    confirmChange = true
                end if
                rstCF.close
            else
                rstCF.close
            end if
        end if

        %><input type="hidden" name="valchanges" value="<%=enc.EncodeForHtmlAttribute(valchanges & "")%>"/><%

		'**RGU 6/11/2009:
		rst.CursorLocation=3
		rst.Open "select xenteo, gestbono from configuracion with(nolock) where nempresa='" & session("ncliente") & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		gestbono=request.Form("gestbono")&""
		modulo_Xenteo=0
		if not rst.EOF then
		    modulo_Xenteo=rst("xenteo")
		    if gestbono="" then gestbono=rst("gestbono")
		end if
		rst.Close
		%><input type="hidden" name="gestbono" value="<%=enc.EncodeForHtmlAttribute(gestbono & "")%>"/><%

		''ricardo 24-3-2004 si existen campos personalizables con titulo no nulo si saldra la pestaña de campos personalizables
		si_campo_personalizables=0
		rst.cursorlocation=3
		rst.open "select ncampo from camposperso with(nolock) where tabla='CLIENTES' and titulo is not null and titulo <> '' and ncampo like '" & session("ncliente") & "%'",session("dsn_cliente")
		if not rst.eof then
			si_campo_personalizables=1
		else
			si_campo_personalizables=0
		end if
		rst.close%>
		<input type="hidden" name="si_campo_personalizables" value="<%=enc.EncodeForHtmlAttribute(si_campo_personalizables & "")%>"/>
		<input type="hidden" name="comp_ncom" value="0"/>
		<iframe style='display:none' id="frCompNcom" name="fr_CompNcom" src='clientes_ncomercial.asp?nc=<%=enc.EncodeForHtmlAttribute(ncliente)%>&ncom=&mode=' class="width60 iframe-menu" frameborder="no" noresize="noresize"></iframe>
		<%'Obtenemos la variable que nos indicará si mostramos los datos o no relacionados con la Asesoria
		si_asesoria = VerObjeto(OBJAsesoriaCli)

		ndecimales=d_lookup("ndecimales","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
		abreviatura=d_lookup("abreviatura","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
		if mode="browse" or mode="edit" or mode="save" then
			cuantos_accesos_tienda=0
			if ncliente & "">"" then
				strselect="select count(usuario) as contador from clientes_users as cu,clientes as c,indice as i"
				strselect=strselect & " where cu.ncliente='" & session("ncliente") & "' and cu.ncliente=c.ncliente"
				strselect=strselect & " and CU.cliente_int='" & ncliente & "'"
				strselect=strselect & " and cu.usuario=i.entrada and cu.fbaja is null"
				rst.cursorlocation=3
				rst.Open strselect,dsnilion
				if not rst.eof then
					cuantos_accesos_tienda=rst("contador")
				end if
				rst.close
			end if
			if cuantos_accesos_tienda>0 then
				mostrar_verstock=""
			else
				mostrar_verstock="none"
			end if
		end if

	    if mode="borrardirenvio" then
		    dir_envio=d_lookup("dir_envio","clientes","ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"))
		    if dir_envio & "">"" then
				    rstAux.open "delete from domicilios where pertenece like '" & session("ncliente") & "%' and codigo='" & dir_envio & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstAux.state<>0 then rstAux.close
				    rstAux.open "update clientes set dir_envio=NULL where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstAux.state<>0 then rstAux.close
				    rst2.open "select ncliente,nproveedor from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if not rst2.eof then
					    rstAux.open "update proveedores set dir_envio=NULL where nproveedor like '" & session("ncliente") & "%' and nproveedor='" & rst2("nproveedor") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					    if rstAux.state<>0 then rstAux.close
				    end if
				    rst2.close
			    'else%>
				    <script language="javascript" type="text/javascript">
				        //window.alert("<%=LitMsgBorrarDirEnvio%>");
				    </script>
			    <%'end if
			    'rst.close
	        end if
	        mode="edit"
	    'Incluimos el borrar de direccion de envio de factura
	    elseif mode="borrardirenvioF" then  
	        invoice_address=d_lookup("invoice_address","clientes","ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"))  
		    if invoice_address & "">"" then
				    rstAux.open "delete from domicilios where pertenece like '" & session("ncliente") & "%' and codigo='" & invoice_address & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstAux.state<>0 then rstAux.close
				    rstAux.open "update clientes set invoice_address=NULL where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstAux.state<>0 then rstAux.close
				    rst2.open "select ncliente,nproveedor from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    rst2.close
			    'else%>
				    <script language="javascript" type="text/javascript">
				        //window.alert("<%=LitMsgBorrarDirEnvio%>");
				    </script>
			    <%'end if
			    'rst.close
	        end if
	        mode="edit"
        end if

   ' **** 26/02/03 VGR
	agente=""
	if viene="agentes" then
		mode="search2"
		agente=enc.EncodeForHtmlAttribute(limpiaCadena(request.querystring("ndoc") & ""))
		if agente="" then
			agente=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("ndoc") & ""))
		end if
		if agente="" then
			agente=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("agente") & ""))
		end if
	end if
	' **** 26/02/03 VGR

	' *** VGR : 07/03/03 ***
	vienecomercial=""
	if viene="comercial" then
		mode="search2"
		vienecomercial=enc.EncodeForHtmlAttribute(limpiaCadena(request.querystring("ndoc") & ""))
		if vienecomercial="" then
			vienecomercial=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("ndoc") & ""))
		end if
		if vienecomercial="" then
			vienecomercial=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("vienecomercial") & ""))
		end if
	end if
	' *** VGR : 07/03/03 ***

		if request.QueryString("submode")>"" then
			submode=enc.EncodeForHtmlAttribute(request.QueryString("submode") & "")
		else
			submode=enc.EncodeForHtmlAttribute(request.form("submode") & "")
		end if
		salto=enc.EncodeForHtmlAttribute(limpiaCadena(request.QueryString("salto") & ""))

		NEntidad=enc.EncodeForJavascript(limpiaCadena(request.form("NEntidad") & ""))
		Oficina=limpiaCadena(request.form("Oficina"))
		DC=limpiaCadena(request.form("DC"))
		Cuenta=limpiaCadena(request.form("Cuenta"))
		Domiciliacion=limpiaCadena(request.form("Domiciliacion"))

		ccontable=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("ccontable") & ""))
		ccontable_efecto=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("ccontable_efecto") & ""))
        CCONTABLE_SUPLIDOS=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("CCONTABLE_SUPLIDOS") & ""))
		rgomaxaut=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("rgomaxaut") & ""))
		rcalc=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("rcalc") & ""))

        'FLM:20/01/2009: duplicar campos bancarios para módulo ORCU
        if si_tiene_modulo_OrCU<>0 then
            'NEntidad2=limpiaCadena(request.form("NEntidad2"))
		   ' Oficina2=limpiaCadena(request.form("Oficina2"))
		   ' DC2=limpiaCadena(request.form("DC2"))
		   ' Cuenta2=limpiaCadena(request.form("Cuenta2"))
		   ' Domiciliacion2=limpiaCadena(request.form("Domiciliacion2"))
		    'dgb 27/10/2009
		    if modulo_Xenteo<>0 then
		        pagoA=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("pagoA") & ""))
		        pagoB=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("pagoB") & ""))
		    end if
		    
		end if
		''EBF Se añade para que por medio de un parametro de usuario no haya que introducir cif. El parametro es &nocif=0

		dim nocif
		dim v_ped, v_alb,v_fra
		dim bh
		dim nmc
		dim emailUsuario
		dim fpUsuario

		nocif=enc.EncodeForHtmlAttribute(limpiaCadena(request.queryString("nocif") & ""))
		if(nocif="") then nocif=enc.EncodeForHtmlAttribute(request.form("nocif") & "")

		''CAG Se añade para indicar si se permite cambiar el numero de copias de una factura
		dim cf
		cf=enc.EncodeForHtmlAttribute(limpiaCadena(request.queryString("cf") & ""))
		if(cf="") then cf=enc.EncodeForHtmlAttribute(request.form("cf") & "")
		'fin ebf Parametro ="NO" No permite ver clientes de tienda.
        dim vint
		vint=enc.EncodeForHtmlAttribute(limpiaCadena(request.queryString("vint") & ""))
		if(vint="") then cf=enc.EncodeForHtmlAttribute(request.form("vint") & "")
		'fin ebf

		ObtenerParametros("clientes")

        if vint="NO" then

            cliente_tienda=d_lookup("codigo","tiendas","ncliente ='" & ncliente & "'",session("dsn_cliente"))
            if cliente_tienda&"">"" then
                %><script languaje="javascript">
                      alert("No tiene permiso para ver este cliente");
                </script><%    
                mode="add"
                ncliente=""
            end if
        end if
		dim lista_valores
		cif=enc.EncodeForHtmlAttribute(limpiaCadena(request.querystring("dni") & ""))
		if cif="" then	cif=enc.EncodeForHtmlAttribute(limpiaCadena(Request.form("cif") & ""))

		'**RGU 13/6/2006
		if nmc="" then nmc=enc.EncodeForHtmlAttribute(limpiaCadena(request.querystring("nmc") & ""))
		if nmc="" then nmc=enc.EncodeForHtmlAttribute(limpiaCadena(request.form("nmc") & ""))
		'**RGU
        mode_accesos_tienda=mode
        if mode_accesos_tienda & ""="delete" then
            mode_accesos_tienda="add"
        end if
        %>
		<input type="hidden" name="mode_accesos_tienda" value="<%=enc.EncodeForHtmlAttribute(mode_accesos_tienda & "")%>"/>
		<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene & "")%>"/>
		<input type="hidden" name="agente" value="<%=enc.EncodeForHtmlAttribute(agente & "")%>"/>
		<input type="hidden" name="vienecomercial" value="<%=enc.EncodeForHtmlAttribute(vienecomercial & "")%>"/>
		<input type="hidden" name="ndoc" value="<%=enc.EncodeForHtmlAttribute(vienecomercial & "")%>"/>
		<input type="hidden" name="verd" value="<%=enc.EncodeForHtmlAttribute(verd & "")%>"/>
		<input type="hidden" name="verl" value="<%=enc.EncodeForHtmlAttribute(verl & "")%>"/>
		<input type="hidden" name="nocif" value="<%=enc.EncodeForHtmlAttribute(nocif & "")%>"/>
		<input type="hidden" name="bh" value="<%=enc.EncodeForHtmlAttribute(bh & "")%>"/>
		<input type="hidden" name="cf" value="<%=enc.EncodeForHtmlAttribute(cf & "")%>"/>
		<input type="hidden" name="vendedor" value="<%=enc.EncodeForHtmlAttribute(vendedor & "")%>"/>
		<input type="hidden" name="nfactura" value="<%=enc.EncodeForHtmlAttribute(nfactura & "")%>"/>
		<input type="hidden" name="obs" value="<%=enc.EncodeForHtmlAttribute(obs & "")%>"/>
        <input type="hidden" name="validacionCliente" value="<%=enc.EncodeForHtmlAttribute(validacionCliente & "")%>"/>       
        <input type="hidden" name="vint" value="<%=enc.EncodeForHtmlAttribute(vint & "")%>"/>
        <input type="hidden" name="hide" value="<%=enc.EncodeForHtmlAttribute(hide & "")%>"/>
        <%if mode="crearcentro" then
			CrearCentro ncliente, true
			mode="browse"
		end if

	''ricardo 19-5-2004 añadir campos personalizables a clientes
	if mode="browse" or mode="edit" or mode="add" or mode="delete" or mode="convertirclidist" or mode = "pasaracontactocomercial" then
	    'JMM 20090805 Modo para relacionar un cliente a un contacto comercial
	    if mode = "pasaracontactocomercial" then
	        'response.Write "Pasar a Contacto Comercial!!!<br/>"
	        
	        contactocom = enc.EncodeForHtmlAttribute(limpiaCadena(Request.QueryString("contactocom") & ""))
	        ncliente = enc.EncodeForHtmlAttribute(limpiaCadena(Request.QueryString("ncliente") & ""))
	        
	        'Contacto Comercial nuevo
	        if contactocom&"" = "nuevo" then
            
	            'Obtenemos el ncontacto que corresponde
	            set rstC = Server.CreateObject("ADODB.Recordset")
	            strSQL = "SELECT NCONTACTO+1 AS NCONTACTO FROM CONFIGURACION with(NOLOCK) WHERE NEMPRESA = '" & session("ncliente") & "'"
	            rstC.CursorLocation=3
	            rstC.Open strSQL, session("dsn_cliente")
	            if not rstC.EOF then
	                contactocom = session("ncliente") & replace(space(5-len(rstC("ncontacto")))," ","0") & rstC("ncontacto")
	            else
	                contactocom=session("ncliente") & "00001"
	            end if
	            rstC.Close
	            'response.Write "ncontacto nuevo: " & contactocom & "<br/>"
	            
	            'Obtenemos el comercial
	            strSQL = "SELECT DNI FROM PERSONAL with(NOLOCK) WHERE DNI LIKE '" & session("ncliente") & "%' AND LOGIN = '" & session("usuario") & "'"
	            rstC.CursorLocation=3
	            rstC.Open strSQL, session("dsn_cliente")
	            if not rstC.EOF then
	                comercial = rstC("dni")
	            end if
	            rstC.Close
	            
	            'Obtenemos los datos del c y direccion cruzados, para generar el nuevo cc
	            strSQL = "SELECT DOMICILIO, CP, POBLACION, PROVINCIA, PAIS, TELEFONO, RSOCIAL, CIF, TELEFONO2, EMAIL " & _
	                     "FROM CLIENTES C WITH(NOLOCK) INNER JOIN DOMICILIOS D WITH(NOLOCK) ON C.DIR_PRINCIPAL = D.CODIGO " & _
	                     "WHERE C.NCLIENTE = '" & ncliente & "' AND D.TIPO_DOMICILIO = 'PRINCIPAL_CLI'"
	            'response.Write strSQL & "<br/>"
	            rstC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
	            
	            'Montamos la query para insertar toda la información del nuevo contacto
	            strSQL = "INSERT INTO CONTACTOSCOMERCIAL WITH(ROWLOCK) " & _
	                     "(CODIGO, NOMBRE, CIF, MOVIL, EMAIL, DIRECCION, CP, POBLACION, PROVINCIA, PAIS, TELEFONO1, COMERCIAL, USUARIO, NCLIENTE) " & _
	                     "VALUES ('" & _
	                                contactocom & "','"& _
	                                rstC("rsocial") &"','"& _
	                                rstC("cif") &"','"& _
	                                rstC("telefono2") &"','"& _
	                                rstC("email") &"','"& _
	                                rstC("domicilio") &"','"& _
	                                rstC("cp") &"','"& _
	                                rstC("poblacion") &"','"& _
	                                rstC("provincia") &"','"& _
	                                rstC("pais") &"','"& _
	                                rstC("telefono") &"','"& _
	                                comercial &"','"& _
	                                trimCodEmpresa(comercial) &"','"& _
	                                ncliente &"')"
	            'response.Write strSQL & "<br/>"
	            rstC.Close
	            
	            rstC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
	            
	            'response.Write "Asignar nuevo CC a C<br/>"
	            strSQL = "UPDATE CLIENTES WITH(UPDLOCK) SET NCONTACTO = '" & contactocom & "' WHERE NCLIENTE = '" & ncliente & "'"
	            rstC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
	            
	            'response.Write "Actualizar ncontaco en configuración<br/>"
	            strSQL = "UPDATE CONFIGURACION WITH(UPDLOCK) SET NCONTACTO = NCONTACTO+1 WHERE NEMPRESA = '" & session("ncliente") & "'"
	            rstC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
	            
	            'Regresamos a la pantalla de cliente
	            'response.Write "<script>document.location = " & cd & "clientes.asp?ncliente=" & ncliente & "&mode=browse" & cd & ";</script>"
	            response.Redirect "clientes.asp?ncliente=" & ncliente & "&mode=browse"
	        'Contacto Comercial existente
	        else
	            'Comprobamos si el CC tiene asociado ya un cliente
	            strSQL = "SELECT NCLIENTE FROM CONTACTOSCOMERCIAL WITH(NOLOCK) WHERE CODIGO = '" & contactocom & "'"
	            set rstCC = Server.CreateObject("ADODB.Recordset")
	            rstCC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
	            
	            'Asignación directa
	            if rstCC("ncliente")&"" = "" then
	                rstCC.Close
	                'Desvinculamos el cliente anterior (si lo estuviese)
	                'Vinculamos el cliente con el contacto comercial y viceversa
	                strSQL = "UPDATE CONTACTOSCOMERCIAL WITH(UPDLOCK) SET NCLIENTE = '" & ncliente & "' WHERE CODIGO = '" & contactocom & "'"
	                rstCC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
	                strSQL = "UPDATE CLIENTES WITH(UPDLOCK) SET NCONTACTO = '" & contactocom & "' WHERE NCLIENTE = '" & ncliente & "'"
	                rstCC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
	                
	                'response.Write "<script>alert('CC: " & contactocom & " asignado a C: " & ncliente & "');</script>"
	                'response.Write "<script>document.location = " & cd & "clientes.asp?ncliente=" & ncliente & "&mode=browse" & cd & ";</script>"
	                response.Redirect "clientes.asp?ncliente=" & ncliente & "&mode=browse"
	            'Reasignación, hay que pedir confirmación al usuario
	            else
	                'response.Write "El CC NO está disponible directamente, será necesario reasignarlo<br/>"
	                rstCC.Close
	                
	                reasignar = limpiaCadena(request.QueryString("reasignar"))
	                
	                if reasignar&"" = "" then
	                    'Preguntamos al usuario si desea reasignar el CC con el C
	                    %>
	                   <script language="javascript" type="text/javascript">
                           reasignar = window.confirm("<%=LitReasignarContacto%>");
                           if (reasignar == true) {
                               reasignar = window.confirm("<%=LitAsigarOK%>");
                               if (reasignar == true)
                                   document.location = "clientes.asp?mode=pasaracontactocomercial&contactocom=<%=enc.EncodeForJavascript(request.querystring("contactocom"))%>&ncliente=<%=enc.EncodeForJavascript(request.querystring("ncliente"))%>&reasignar=SI";
	                            else
                               {
                                   window.alert("<%=LitCanceladoUsuario%>");
                                   document.location = "clientes.asp?ncliente=<%=enc.EncodeForJavascript(request.querystring("ncliente"))%>&mode=browse";
                               }
                           }
                           else {
                               alert("<%=LitAccionNoRealizada%>");
                               document.location = "clientes.asp?ncliente=<%=enc.EncodeForJavascript(request.querystring("ncliente"))%>&mode=browse";
                           }
	                    </script>
	                    <%
	                else if reasignar&"" = "NO" then
	                    'Mostramos la cancelación
	                    response.Write "<script language='javascript' type='text/javascript'>alert('" & LitCanceladoUsuario & "');</script>"
	                else if reasignar&"" = "SI" then
	                    'Realizamos la operación y lo indicamos por pantalla
	                    'Desvinculamos el cliente anterior
                        strSQL = "UPDATE CLIENTES WITH(UPDLOCK) SET NCONTACTO = NULL WHERE NCONTACTO = '" & contactocom & "'"
                        rstCC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
                        'Vinculamos el cliente con el contacto comercial y viceversa
                        strSQL = "UPDATE CONTACTOSCOMERCIAL WITH(UPDLOCK) SET NCLIENTE = '" & ncliente & "' WHERE CODIGO = '" & contactocom & "'"
                        rstCC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
                        strSQL = "UPDATE CLIENTES WITH(UPDLOCK) SET NCONTACTO = '" & contactocom & "' WHERE NCLIENTE = '" & ncliente & "'"
                        rstCC.Open strSQL, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
                        
                        response.Redirect "clientes.asp?ncliente=" & ncliente & "&mode=browse"
	                end if
	                end if
	                end if
	            end if
	        end if
	        
	        response.End
	    end if
	
		'**RGU 17/11/2006
		num_campos= d_count("NCAMPO","CAMPOSPERSO","TABLA='CLIENTES' and NCAMPO like '"&SESSION("NCLIENTE")&"%' ",session("dsn_cliente"))
		%><input type="hidden" name="num_campos_tabla" value="<%=enc.EncodeForJavascript(num_campos)%>"/><%
		'**RGU
		'num_campos=0

		if mode="add" then
			'redim lista_valores(10+2)
			redim lista_valores(num_campos)
			for ki=1 to num_campos
				lista_valores(ki)=""
			next
			'num_campos=10
		else
			strcampo=""
            rstAux2.cursorlocation=3
			rstAux2.open "select NCAMPO from CAMPOSPERSO with(nolock) where NCAMPO like '" & session("ncliente") & "%' and TABLA='CLIENTES' order by NCAMPO",session("dsn_cliente")
            while not rstAux2.EOF
            strcampo=strcampo&" , c.campo"&trimcodempresa(rstAux2("ncampo"))
                
                rstAux2.MoveNext
            wend
            rstAux2.Close
			''for ki=1 to num_campos
			''	strcampo=strcampo&" , c.campo"&iif(len(ki)=1,"0"&ki,ki)
			''next
			rstAux2.cursorlocation=3
			rstAux2.open "select c.ncliente "&strcampo&" from clientes as c with(nolock) where c.ncliente like '" & session("ncliente") & "%' and c.ncliente='" & ncliente & "'",session("dsn_cliente")
			if not rstAux2.eof then
				redim lista_valores(num_campos)
                rst.cursorlocation=3
			    rst.open "select NCAMPO from CAMPOSPERSO with(nolock) where NCAMPO like '" & session("ncliente") & "%' and TABLA='CLIENTES' order by NCAMPO",session("dsn_cliente")
                cont = 1
				while not rst.EOF
                    if isnull(rstAux2("campo"&trimcodempresa(rst("ncampo")))) = false then
                        lista_valores(cont)=Nulear(rstAux2("campo"&trimcodempresa(rst("ncampo"))))
                    else
                        lista_valores(cont)=""
                    end if

                    rst.MoveNext
                    cont = cont + 1
                wend
                rst.Close
                'for ki=1 to num_campos
				'	ncampo=iif(len(ki)=1,"0"&ki,ki)
				'	lista_valores(ki)=Nulear(rstAux2("campo"&ncampo))

				'next
			else
				redim lista_valores(num_campos)
				for ki=1 to num_campos
					lista_valores(ki)=""
				next
			end if
			rstAux2.close
		end if
	end if

  if mode="convertirclidist" then
	rst.Open "select * from clientes where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if rst.eof then
		'no se puede añadir un distribuidor sin estar dato de alta como cliente
	else
		ndist=""
		error= "no"
		p_nproveedor=""
		if rst("cifedi")>"" then
			if salto="no" then
				rstAux.Open "select nproveedor from proveedores with(nolock) where nproveedor like '" & session("ncliente") & "%' and cifedi='" & rst("cifedi") & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic

				if not rstAux.EOF then
					nproveedor=rstAux("nproveedor")
					error = "si"%>
					<script language="javascript" type="text/javascript">
                           //window.alert("<%=LitNoPuedeCrearDistribuidor%>");
                           if (window.confirm("<%=LitCrearDistribuidorProveedorExistente%>") == false) {
                               document.clientes.action = "clientes.asp?ncliente=<%=enc.EncodeForJavascript(ncliente)%>&mode=browse";
                               document.clientes.submit();
                           }
                           else {
                               document.clientes.action = "clientes.asp?ncliente=<%=enc.EncodeForJavascript(ncliente)%>&mode=convertirclidist&salto=si&nproveedor=<%=enc.EncodeForJavascript(nproveedor)%>";
                               document.clientes.submit();
                           }
					</script>
				<%end if
				rstAux.Close
			else
				p_nproveedor=limpiaCadena(request.querystring("nproveedor"))
			end if
		else
			error = "si"%>
			<script language="javascript" type="text/javascript">
                            window.alert("<%=LITCIFINCORRECTO%>");
			</script>
		<%end if
		if error= "no" then
			nproveedor=GuardarProveedor(p_nproveedor,ncliente)
			rst2.open "select * from distribuidores where ndist like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			num=d_max("substring(ndist,6,10)","distribuidores","ndist like '" & session("ncliente") & "%'",session("dsn_cliente"))
			rst2.AddNew
			rst2("ndist")=session("ncliente") & completar(cstr(num + 1),5,"0")
			ndist=rst2("ndist")
			rst2("ncliente")=ncliente
			rst2("nproveedor")=nproveedor
			rst2.update
			rst2.close
		end if

		if error= "no" and ndist>"" then
			'rst("ndist")=ndist
			'rst.update
			auditar_ins_bor session("usuario"),ndist,nproveedor,"alta","proveedor","","distribuidores"
		end if
		ndist=""
	end if
	rst.close
	mode="browse"
  end if

 'Acción a realizar
  if mode="save" then
                
        ibanCom_tmp = Request.Form("country")&Request.Form("iban")&Request.Form("NEntidad")&Request.Form("Oficina")&Request.Form("DC")&Request.Form("Cuenta")
        tmp_bnc=Mid(ibanCom_tmp, 5, 4)
        cuentaOK=ComprobarCuenta(tmp_bnc, ibanCom_tmp)
        'debemos reasignar la cuenta por logica de la pagina
        cuenta=Request.Form("NEntidad")&Request.Form("Oficina")&Request.Form("DC")&Request.Form("Cuenta")
  		'cuenta=trim(NEntidad&Oficina&DC&Cuenta)
		'cuentaOK=true
		'if Nz_b(Domiciliacion)<>0 then
			'strBanco = Mid(cuenta, 1, 4)
    		'strOficina = Mid(cuenta, 5, 4)
    		'strDC1 = Mid(cuenta, 9, 1)
    		'strDC2 = Mid(cuenta, 10, 1)
    		'strCuenta = Mid(cuenta, 11, 10)
			'if cuenta="" then
				'CuentaOK=false
			'ElseIf Not Validar_cuenta(strBanco & strOficina, strDC1, False, strDC1Bueno) and len(cuenta) = 20 Then
     			'CuentaOK=false
    		'ElseIf Not Validar_cuenta(strCuenta, strDC2, True, strDC2Bueno) and len(cuenta) = 20  Then
     			'CuentaOK=false
    		'Else
      			'CuentaOK=true
    		'End If
		'end if

		'FLM:20/01/2009: añadir una segunda cuenta contable. modulo ORCU.
		'if si_tiene_modulo_OrCU<> 0 then
		  '  cuenta2=trim(NEntidad2&Oficina2&DC2&Cuenta2)
    	    cuentaOK2=true
		   ' if Nz_b(Domiciliacion2)<>0 then
			'    strBanco2 = Mid(cuenta2, 1, 4)
    		'	    strOficina2 = Mid(cuenta2, 5, 4)
    		'	    strDC12 = Mid(cuenta2, 9, 1)
    		'	    strDC2 = Mid(cuenta2, 10, 1)
    		'	    strCuenta2 = Mid(cuenta2, 11, 10)
			'    if cuenta2="" then
			'	    CuentaOK2=false
			'    ElseIf Not Validar_cuenta(strBanco2 & strOficina2, strDC12, False, strDC1Bueno2) Then
     		'		    CuentaOK2=false
    		'	    ElseIf Not Validar_cuenta(strCuenta2, strDC2, True, strDC2Bueno2) Then
     		'		    CuentaOK2=false
    		'	    Else
      		'	    CuentaOK2=true
    		'	    End If
		    'end if
		'end if
		'Fin ''''''
        if si_tiene_modulo_TGB <>0 then

            NEntidadTGB=enc.EncodeForJavascript(limpiaCadena(request.form("NEntidadGB")))
		    OficinaTGB=limpiaCadena(request.form("OficinaGB"))
		    DCTGB=limpiaCadena(request.form("DCGB"))
            cuentaTGB=limpiaCadena(request.form("cuentaGB"))
            cuentaTGB=trim(NEntidadTGB&OficinaTGB&DCTGB&CuentaTGB)
		    cuentaTGBOK=true
		    if cuentaTGB&"">"" then
			    strBancoTGB = Mid(cuentaTGB, 1, 4)
    			strOficinaTGB = Mid(cuentaTGB, 5, 4)
    			strDC1TGB = Mid(cuentaTGB, 9, 1)
    			strDC2TGB = Mid(cuentaTGB, 10, 1)
    			strCuentaTGB = Mid(cuentaTGB, 11, 10)
			    if cuentaTGB="" then
				    CuentaTGBOK=false
			    ElseIf Not Validar_cuenta(strBancoTGB & strOficinaTGB, strDC1TGB, False, strDC1BuenoTGB) Then
     				    CuentaTGBOK=false
    			ElseIf Not Validar_cuenta(strCuentaTGB, strDC2TGB, True, strDC2BuenoTGB) Then
     				    CuentaTGBOK=false
    			Else
      			    CuentaTGBOK=true
    			End If
                if CuentaTGBOK=false then
                    cuentaTGB=""
                    %><script language="javascript" type="text/javascript">alert("<%=LitErrBankTGB%>")</script><%
                end if
            end if

		end if
		
		'comprobamos que la cuenta contable no este ya en el sistema
		no_seguir=0
		if ccontable& "">"" then
			rst.Open "select * from clientes where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'", _
			session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if rst.eof then 'AÑADIR
				rst.close
			else
				rst.close
			end if
		end if

		if no_seguir=0 then
			CIF=LimpiarCIF(cif)

			if CIF>"" or nocif="0" then
		  		rst.Open "select * from clientes where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if rst.eof then 'AÑADIR
				
					rstAux.Open "select cifedi from clientes where ncliente like '" & session("ncliente") & "%' and cifedi='" & CIF & "'", _
					session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					'' MPC 08/10/2008 Si tiene el parámetro cifrepe y a la pregunta responde que si entonces inserta el cliente aunque exista el cif
					if rstAux.EOF or repe =1 then

					    'FLM:21/01/2009:Mod para que tenga encuenta la cuenta2 si tiene el mód de ORCU.
						'if   (CuentaOK and si_tiene_modulo_OrCU=0 )THEN or (CuentaOK and ( si_tiene_modulo_OrCU<>0))then
						if   (CuentaOK )then
							rstAux.close
							if GuardarRegistro(ncliente)=1 then
								'Si este cliente es tambien distribuidor entonces se guardara las modificaciones tambien en proveedores
								rst2.open "select * from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
								if not rst2.eof then
									nproveedor=GuardarProveedor(rst2("nproveedor"),ncliente)
								end if
								rst2.close
								mode="browse"
								ncliente=rst("ncliente")
								rst.Close
								auditar_ins_bor session("usuario"),"",ncliente,"alta","","","clientes"
								if viene>"" then
									if viene="albaranes_cli" then
										pagina="albaranes_cli"
									elseif viene="albaranes_cli_fast" then
										pagina="albaranes_cli_fast"
									elseif viene="facturas_cli" then
										pagina="facturas_cli"
									elseif viene="pedidos_cli" then
										pagina="pedidos_cli"
									elseif viene="presupuestos_cli" then
										pagina="presupuestos_cli"
								    elseif viene="collaborations" then
								        pagina="clientes"
                                    elseif viene = "facturar_tickets" then
                                        pagina = "facturar_tickets"
									end if
									rst.Open "select * from clientes where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente")
									if not rst.eof then

							            'DGM si viene de colaboraciones, devolvemos ciertos datos
									    ' Lo excluyo de demás comprobaciones para evitar choques de datos
							            if viene="collaborations" then
							                rst2.open "select CODIGO,domicilio,cp,poblacion, provincia,pais,telefono,codprovincia,codpoblacion,codpais from domicilios with(nolock) where pertenece like '" & session("ncliente") & "%' and tipo_domicilio = 'PRINCIPAL_CLI' and pertenece = '" & ncliente &"' ",_
							                session("dsn_cliente"),adOpenKeyset,adLockOptimistic%>
							                    <script language="javascript" type="text/javascript">
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.ncliente.value="<%=enc.EncodeForJavascript(trimCodEmpresa(ncliente))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.ncomercial.value="<%=enc.EncodeForJavascript(rst("RSOCIAL"))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.nombre.value="<%=enc.EncodeForJavascript(rst("RSOCIAL"))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.contacto.value="<%=enc.EncodeForJavascript(rst("CONTACTO"))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.web.value="<%=enc.EncodeForJavascript(rst("WEB"))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.email.value="<%=rst("EMAIL")%>";
											    </script>
							                <%if not rst2.eof then
							                    %><script language="javascript" type="text/javascript">
                                                      window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.hnaddress.value="<%=rst2("CODIGO")%>";
                                                      window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.poblacion.value="<%=rst2("POBLACION")%>";
                                                      window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.provincia.value="<%=rst2("PROVINCIA")%>";
                                                      window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.pais.value="<%=rst2("PAIS")%>";
                                                      window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.telefono.value="<%=rst2("TELEFONO")%>";
                                                      window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.domicilio.value="<%=rst2("DOMICILIO")%>";
                                                      window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.cp.value="<%=rst2("CP")%>";
											    </script><%
							                end if
							                rst2.close
							            'FIN DGM
                                        elseif viene = "facturar_tickets" then %>
                                            <script language="javascript" type="text/javascript">
                                                      window.top.opener.document.facturar_tickets.ncliente.value = "<%=enc.EncodeForJavascript(trimCodEmpresa(ncliente))%>";
                                                      window.top.opener.document.facturar_tickets.nombre.value = "<%=enc.EncodeForJavascript(rst("RSOCIAL"))%>";
                                            </script>
										<% elseif viene<>"pedidos_cli_nc" then
										    if viene<>"subcuentas" and viene<>"facturas_cli_E" then%>
											    <script language="javascript" type="text/javascript">
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.ncliente.value="<%=enc.EncodeForJavascript(trimCodEmpresa(ncliente))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.nombre.value="<%=enc.EncodeForJavascript(rst("RSOCIAL"))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.forma_pago.value="<%=rst("fpago")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.tipo_pago.value="<%=rst("tpago")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.transportista.value="<%=enc.EncodeForJavascript(rst("transportista"))%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.portes.value="<%=rst("portes")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.dto1.value="<%=rst("dto")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.dto2.value="<%=rst("dto2")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.rf.value="<%=rst("recargo")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.tarifa.value="<%=rst("tarifa")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.divisa.value="<%=rst("divisa")%>";
                                                    window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.h_divisa.value="<%=rst("divisa")%>";
											    </script>
											    <%if viene="albaranes_cli_fast" and si_tiene_modulo_comercial<>0 then%>
												    <script language="javascript" type="text/javascript">
                                                        window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.comercial.value="<%=rst("comercial")%>";
												    </script>
											    <%elseif viene<>"albaranes_cli_fast" then%>
												    <script language="javascript" type="text/javascript">
                                                        window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.comercial.value="<%=rst("comercial")%>";
												    </script>
											    <%end if
											    if viene="presupuestos_cli" then%>
												    <script language="javascript" type="text/javascript">
                                                        window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.contacto_com.value="";
                                                        window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(pagina) %>.nombre_com.value="";
												    </script>
											    <%end if
											    
											    'AMF:21/12/2010:Llevamos el centro a la incidencia si se crea automaticamente.
											    if viene="incidencias" then
											        if d_lookup("autocentro", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente")) then
				                                        ncentro = d_lookup("ncentro", "centros", "ncentro like '" & session("ncliente") & "%' and ncliente = '" & ncliente & "'", session("dsn_cliente"))
			                                            ncentro = trimCodEmpresa(ncentro)%>
			                                                <script language="javascript" type="text/javascript">
                                                                LlevarCentroIncidencia("<%=enc.EncodeForJavascript(ncentro)%>");
												            </script>
			                                        <%else%>
			                                                <script language="javascript" type="text/javascript">
                                                                parent.window.close();
												            </script>
			                                        <%end if
			                                    end if
			                                    if viene="centros" then
			                                        ncentro=""
			                                        generadoCentro="0"
											        if d_lookup("autocentro", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente")) then
				                                        ncentro = d_lookup("ncentro", "centros", "ncentro like '" & session("ncliente") & "%' and ncliente = '" & ncliente & "'", session("dsn_cliente"))
			                                            ncentro = trimCodEmpresa(ncentro)
			                                            generadoCentro="1"
			                                        end if%>
		                                                <script language="javascript" type="text/javascript">
                                                                LlevarClienteAOrdenes("<%=enc.EncodeForJavascript(ncliente)%>", "<%=enc.EncodeForJavascript(ncentro)%>", "<%=enc.EncodeForJavascript(generadoCentro)%>");
											            </script>
		                                        <%end if
    											
											    'jcg 02/02/2008
											    if si_tiene_modulo_proyectos<>0 then%>
												    <script language="javascript" type="text/javascript">
                                                                window.top.opener.parent.pantalla.document.<%=pagina %>.cod_proyecto.value="<%=rst("proyecto")%>";
												    </script>	
											    <%end if
										    else%>
											    <script language="javascript" type="text/javascript">
                                                            window.top.opener.parent.pantalla.document.subcuentas.ncliente.value = "<%=enc.EncodeForJavascript(ncliente)%>";
                                                            window.top.opener.parent.pantalla.fr_Cliente.document.docclientes.ncliente.value = "<%=enc.EncodeForJavascript(trimCodEmpresa(ncliente))%>";
                                                            window.top.opener.parent.pantalla.fr_Cliente.document.docclientes.nom_cliente.value = "<%=enc.EncodeForJavascript(rst("RSOCIAL"))%>";
											    </script>
										    <%end if
										else
                                        %>
												<script language="javascript" type="text/javascript">
                                                        window.top.opener.parent.pantalla.document.pedidos_cli.ncliente.value = "<%=enc.EncodeForJavascript(trimcodEmpresa(ncliente))%>";
                                                        window.top.opener.parent.pantalla.TraerCliente('add', '1');
												</script>
										<%end if%>
										<script language="javascript" type="text/javascript">
                                                    parent.parent.window.close();
										</script>
									<%else%>
										<script language="javascript" type="text/javascript">
                                                    window.alert("<%=LITNOCREARCLIENTE%>");
                                                    parent.window.close();
										</script>
									<%end if
									rst.close
								end if

							else
								ncliente=""
								mode="add"
								rst.Close
							end if
						else
							rstAux.Close
							rst.Close%>
							<script language="javascript" type="text/javascript">
                                            //visto
                                            alert("<%=LitCuentaBancariaIncorrecta%>");
                                            history.back();
                                            //parent.botones.history.back();
                                            parent.botones.location = "clientes_bt.asp?mode=add";
								<%if viene> "" then%>
								<%else%>
									//parent.botones.history.back();
								<%end if%>
							</script>
						<%end if

					else
						rstAux.Close
						rst.Close%>
						<script language="javascript" type="text/javascript">
                                    alert("<%=LITCIFYAEXISTE%>");
                                history.back();
                                //parent.botones.history.back();
                                parent.botones.location = "clientes_bt.asp?mode=add";
							<%if viene> "" then%>
							<%else%>
								//parent.botones.history.back();
							<%end if%>
						</script>
					<%end if
				else 'EDITAR

					if ucase(trim(CIF))=ucase(trim(rst("cifedi"))) then

					    'FLM:21/01/2009:Mod para que tenga encuenta la cuenta2 si tiene el mód de ORCU.
						if (CuentaOK) then' and (CuentaOK2 and si_tiene_modulo_OrCU<>0)) or (CuentaOK and si_tiene_modulo_OrCU=0 )then

							if GuardarRegistro(ncliente)=1 then

								'Si este cliente es tambien distribuidor entonces se guardara las modificaciones tambien en proveedores
								rst2.open "select * from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
								if not rst2.eof then
									nproveedor=GuardarProveedor(rst2("nproveedor"),ncliente)
								end if
								rst2.close
								mode="browse"
								ncliente=rst("ncliente")
								rst.Close
							else

								ncliente=""
								mode="add"
								rst.Close
							end if
						else

							rst.Close%>
							<script language="javascript" type="text/javascript">
                                    //visto
                                    alert("<%=LitCuentaBancariaIncorrecta%>");
                                history.back();
                                history.back();
                                //parent.botones.history.back();
                                parent.botones.location = "clientes_bt.asp?mode=edit";
							</script>
						<%end if

					else

					    rstAux.Open "select cifedi from clientes where ncliente like '" & session("ncliente") & "%' and cifedi='" & CIF & "' and ncliente <> '"&ncliente&"'", _
						session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						'' MPC 08/10/2008 Si tiene el parámetro cifrepe y a la pregunta responde que si entonces inserta el cliente aunque exista el cif
						if rstAux.EOF or repe = 1 then
						    'FLM:21/01/2009:Mod para que tenga encuenta la cuenta2 si tiene el mód de ORCU.
							if (CuentaOK) then ' and (CuentaOK2 and si_tiene_modulo_OrCU<>0)) or (CuentaOK and si_tiene_modulo_OrCU=0 )then
								rstAux.close
								if GuardarRegistro(ncliente)=1 then
									'Si este cliente es tambien distribuidor entonces se guardara las modificaciones tambien en proveedores
									rst2.open "select * from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
									if not rst2.eof then
										nproveedor=GuardarProveedor(rst2("nproveedor"),ncliente)
									end if
									rst2.close
									mode="browse"
									ncliente=rst("ncliente")
									rst.Close
								else
									ncliente=""
									mode="add"
									rst.Close
								end if
							else
								rstAux.Close
								rst.Close%>
								<script language="javascript" type="text/javascript">
                                    //visto
                                    alert("<%=LitCuentaBancariaIncorrecta%>");
                                    history.back();
                                    parent.botones.location = "clientes_bt.asp?mode=edit";
								</script>
							<%end if
						else
							rstAux.Close
							rst.Close%>
							<script language="javascript" type="text/javascript">
                                    alert("<%=LITCIFYAEXISTE%>");
                                    history.back();
                                    parent.botones.location = "clientes_bt.asp?mode=edit";
							</script>
						<%end if
					end if
				end if
			else%>
				<script language="javascript" type="text/javascript">
                                    alert("<%=LITCIFINCORRECTO%>");
                                    history.back();
                                    parent.botones.history.back();
				</script>
			<%end if
		end if
  elseif mode="delete" then

		he_borrado=1
	    'response.write("Cliente a borrar " + ncliente)
		nomcliaux = d_lookup("rsocial","clientes","ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente"))
		rstAux.open "select ndist,nproveedor from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente")
		if not rstAux.eof then
			ndist=rstAux("ndist")
			nproveedor=rstAux("nproveedor")
		else
			ndist=""
			nproveedor=""
		end if
		rstAux.close
		cli=BorrarRegistro(ncliente)
		if cli>"" then
			mode="browse"
			ncliente=cli
		else
			if ndist>"" then
				auditar_ins_bor session("usuario"),ndist,ncliente,"baja",nproveedor,"","distribuidores"
			else
				auditar_ins_bor session("usuario"),"",ncliente,"baja",nomcliaux,"","clientes"
			end if
			ndist=""
			nproveedor=""
            'DGB:  change to add
			mode="add"
			ncliente=""%>
			<script language="javascript" type="text/javascript">
                                //dgb: chante to add, refresh search page and open it
                                parent.botones.document.location = "clientes_bt.asp?mode=add";
                                SearchPage("client_lsearch.asp?mode=init", 0);

			</script>
		<%end if
  end if
  nomcli   = d_lookup("rsocial","clientes","ncliente like '" & session("ncliente") & "%' and ncliente='"+ncliente+"'",session("dsn_cliente"))

  if viene="agentes" then
	PintarCabeceraPopUp "Clientes"
  else
  	  if si_tiene_modulo_agrario<>0 then
	  	  PintarCabecera "socios.asp"
	  else
		  PintarCabecera "clientes.asp"
	  end if
  end if

	rst2.cursorlocation=3
	rst2.open "select ndist from distribuidores where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'",session("dsn_cliente")
	if not rst2.eof then
		titulo2=LitClientes2
	else
		titulo2=LitClientes
	end if
	rst2.close



if mode<>"search" then

	if VerObjeto(OBJContacto) then VinculosPagina(MostrarContacto)=1
	CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

	rstAux.open "select con.codigo,con.nombre from contactoscomercial con with(nolock),clientes as cli with(nolock) where con.codigo like '" & session("ncliente") & "%'  and con.ncliente=cli.ncliente and cli.ncliente='"&ncliente&"'",session("dsn_cliente")

	if viene="agentes" or viene="comercial" then
		if viene="agentes" then
			nomcom = d_lookup("nombre","agentes","codigo like '" & session("ncliente") & "%' and codigo='" + agente + "'",session("dsn_cliente"))
		end if
		if viene="comercial" then
			nomcom = d_lookup("nombre","personal","dni like '" & session("ncliente") & "%' and dni='" + vienecomercial + "'",session("dsn_cliente"))
		end if

            DrawDiv "header-agent","",""
            DrawLabel "","",iif(viene="agentes",LitAgen,LitComer) & " " & enc.EncodeForHtmlAttribute(nomcom & "")
            CloseDiv

            if mode="browse" then    
			rstAux3.cursorlocation=3
			' Añadimos el formato de carta DOMICILIACIÓN BANCARIA'
			rstAux3.Open "select codigo, nombre from cartas with(nolock) where codigo like '" & session("ncliente") & "%' union select '99999','Domiciliación Bancaria COAG' as nombre order by nombre", session("dsn_cliente")%>

            <%DrawDiv "header-presletter","",""
            %><a class="CELDAREFB" href="javascript:validarCampoCarta('<%=enc.EncodeForJavascript(ncliente)%>','<%=enc.EncodeForJavascript(session("ncliente"))%>');" onmouseover="self.status='<%=LitCartaPresentacion%>'; return true;" onmouseout="self.status=''; return true;"><b><%=LitCartaPresentacion%></b></a><%
            DrawSelectHeaderPressLetter "CELDARIGHT","60","",0,"","cartas",rstAux3,"","codigo","nombre","",""
            CloseDiv
			rstAux3.close
            %>
            
            <div class="headers-wrapper"></div>
            <%  
                PaintAction 
            %>

            <%end if %>

	<%else
        if not rstAux.eof then
            if mode="browse" then
                ancho_cliente="38%"
                ancho_segcomercial="38%"
            else
                ancho_cliente="49%"
                ancho_segcomercial="49%"
            end if
        else
            if mode="browse" then
                ancho_cliente="76%"
                ancho_segcomercial="0%"
            else
                ancho_cliente="99%"
                ancho_segcomercial="0%"
            end if
        end if
        no_mostrar=0
        ''ricardo 31/7/2003 comprobamos que existe el cliente
''response.write("los datos son-" & he_borrado & "-" & ncliente & "-" & comercialSolSusCli & "-<bR>")
        if he_borrado<>1 and ncliente & "">"" then
            set rstAuxCli=Server.CreateObject("ADODB.Recordset")
		    rstAuxCli.cursorlocation=3
            StrselCli="select ncliente from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'"
            if comercialSolSusCli & "">"" then
                StrselCli=StrselCli & " and comercial='" & comercialSolSusCli & "' "
            end if
''response.write("el StrselCli es-" & StrselCli & "-<br>")
		    rstAuxCli.open StrselCli, session("dsn_cliente")
		    if rstAuxCli.eof then
			    %>
			    <script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgDocsNoExiste%>");
                    document.clientes.action = "clientes.asp?mode=add"
                    document.clientes.submit();
                    parent.botones.document.location = "clientes_bt.asp?mode=add";
			    </script>
			    <%
                mode="add"
                ncliente=""
                nomcli=""
			    no_mostrar=1
		    end if
		    rstAuxCli.close
            set rstAuxCli=nothing
        end if
        if no_mostrar=0 then
            DrawDiv "header-client","",""
            DrawLabel "headerLabel","",iif(si_tiene_modulo_agrario<>0,LitNsocio,LitNcliente)
            DrawSpan "","",trimCodEmpresa(ncliente),""
            CloseDiv
            ''EligeCeldaResponsiveCabecera "text",mode,clase,"","",0,"",iif(si_tiene_modulo_agrario<>0,LitNsocio,LitNcliente),"",trimCodEmpresa(ncliente)
            DrawDiv "header-rsocial","",""
            DrawLabel "headerLabel","",LitRSocial
            DrawSpan "","",enc.EncodeForHtmlAttribute(nomcli & ""),""
            CloseDiv
            ''EligeCeldaResponsiveCabecera "text",mode,clase,"","",0,"",LitRSocial,"",replace(replace(nomcli,"<","&lt;"),"'","&#39;")
                    if not rstAux.eof then
                    DrawDiv "header-ncontact","",""
                    DrawLabel "headerLabel","",LitNContacto
                    DrawSpan "","",Hiperv(OBJContacto,rstAux("codigo"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rstAux("codigo")),LitVerCliente),""
                    CloseDiv
                    ''EligeCeldaResponsiveCabecera "text",mode,clase,"","",0,"",LitNContacto,"",Hiperv(OBJContacto,rstAux("codigo"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rstAux("codigo")),LitVerCliente)
            
                    tieneContacto = "SI"
                    
                    DrawDiv "header-name","",""
                    DrawLabel "headerLabel","",LitNombre
                    DrawSpan "","",rstAux("nombre"),""
                    CloseDiv
                    ''EligeCeldaResponsiveCabecera "text",mode,clase,"","",0,"",LitNombre,"",rstAux("nombre")

				    end if
                    if mode="browse" then %>
                        <%    
			            rstAux3.cursorlocation=3
			            ' Añadimos el formato de carta DOMICILIACIÓN BANCARIA'
			            rstAux3.Open "select codigo, nombre from cartas with(nolock) where codigo like '" & session("ncliente") & "%' union select '99999','Domiciliación Bancaria COAG' as nombre order by nombre", session("dsn_cliente")%>
                            
			                
                        <%DrawDiv "header-presletter","",""
                        %><a class="CELDAREFB" href="javascript:validarCampoCarta('<%=enc.EncodeForJavascript(ncliente)%>','<%=enc.EncodeForJavascript(session("ncliente"))%>');" onmouseover="self.status='<%=LitCartaPresentacion%>'; return true;" onmouseout="self.status=''; return true;"><b><%=LitCartaPresentacion%></b></a><%
			            DrawSelectHeaderPressLetter "CELDARIGHT","60","",0,"","cartas",rstAux3,"","codigo","nombre","",""
                        CloseDiv
			            rstAux3.close
                        %><div class="headers-wrapper"></div>
                        <%  
                        PaintAction 
                        %>

                    <%end if %>
	        <%
        end if
    end if
	rstAux.Close
end if


  Alarma "clientes.asp"
  
  if mode="add" then
	' Inicio Borde Span%>
    <table width="100%">
    <div id="CollapseSection">
    <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['AddDG', 'AddDC', 'AddDB','AddDE','AddOD','AddDD','AddCP']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title=""/></a> 
    <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['AddDG', 'AddDC', 'AddDB','AddDE','AddOD','AddDD','AddCP']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title=""/></a>
    </div>
            <table></table>

	<% 'DATOS GENERALES MODO AÑADIR %>
     <div   class="Section" id="S_AddDG">
    <a href="#" rel="toggle[AddDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader displayed" >
    <%=LitDatosGenerales%>
    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
     </div></a>
    <div class="SectionPanel" style="display: block;" id="AddDG">
	 
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
        <%
            DrawInputCeldaLabel "","txtMandatory",35,LitRSocial,"rsocial",""
            DrawInputCelda "","","",35,0,LitNComercial,"ncomercial",""
            
            if si_asesoria=true then
				    rstSelect.open "select codigo, descripcion from FORMAS_JURIDICAS with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                    DrawSelectCelda "","200","",0,LitFormaJuridica,"fjuridica",rstSelect,"","codigo","descripcion","",""
                    rstSelect.close
                    DrawInputCelda "","","",35,0,LitTitular,"titular",""
		    end if

            DrawInputCeldaLabel "","txtMandatory",20,LitCIF,"cif",cif
            DrawInputCelda "","","",35,0,LitContacto,"contacto",""
            DrawInputCeldaLabel "","txtMandatory",35,LitDomicilio,"domicilio",""
            DrawDiv "1", "", ""
            CloseDiv

            DrawDiv "1","",""
            'DrawLabel "", "", LitPoblacion
            
                %><label><%=LitPoblacion%></label><%
                %><input class="width50" type="text" size="25" name="poblacion" onchange="borrarCodigos('1')"/><%
                %><a class='CELDAREFB' onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes&titulo=<%=LITSSVERPOBLACIONES %>');"  href="#SELECCIONAR_POBLACION2"  onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><%
                    %><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/><%
                %></a><%
            
            CloseDiv

			%>
			<input type="hidden" name="codPoblacion" value=""/> 
			<input type="hidden" name="codProvincia" value=""/>
			<input type="hidden" name="codPais" value=""/><%
             'ASP FIN 
             EligeCelda "input",mode,"left","","",0,LitCP,"cp",5,""

             rstSelect.cursorlocation=3
			 rstSelect.open "select id, nombre from PAISES with(nolock) order by nombre",DSNIlion
			 DrawSelectCeldaInput "","200","",0,LitPais,"paisDDL",rstSelect,"","id","nombre","onchange","TraerPais()","pais",30,""
			 rstSelect.close

             rstSelect.cursorlocation=3
		 	 rstSelect.open "select idprovincia, descripcion, idpais from PROVINCIAS with(nolock) order by idpais, descripcion",DSNIlion

              	%>
             <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
                 <label><%=LitProvincia%></label>           
                 <select class="width30" style="" name="provinciaDDL" onchange="TraerProvincia()"> 
                 <% while not rstSelect.EOF %>
                        <option style="display:none;" value="<%=rstSelect("idprovincia")%>" data-country="<%=rstSelect("idpais")%>"><%=rstSelect("descripcion")%></option>
                 <% 
                        rstSelect.MoveNext
                    wend
                 %>
                    <option selected="" value=""></option>
                 </select>
                 <input class="width30" type='text' name="provincia" value="" size="25" />
             </div>            
             <%
        

             'DrawSelectCeldaInput "","200","",0,LitProvincia,"provinciaDDL",rstSelect,"","idprovincia","descripcion","onchange","TraerProvincia()","provincia",25,""

			 rstSelect.close             

			 EligeCelda "input",mode,"left","","",0,LitTel1,"telefono",25,""
             EligeCelda "input",mode,"left","","",0,LitTel2,"telefono2",20,""
             
             EligeCelda "input",mode,"left","","",0,LitFax,"fax",20,""

             EligeCelda "input",mode,"left","","",0,LitFechaAlta,"falta",10,day(date) & "/" & month(date) & "/" & year(date)
             DrawCalendar "falta"
             EligeCelda "input",mode,"left","","",0,LitFechaBaja,"fbaja",10,""
             DrawCalendar "fbaja"

             EligeCelda "input",mode,"maxlength='255'","","",35,LitEMail,"email",35,emailUsuario
             EligeCelda "input",mode,"left","","",0,LitWEB,"web",35,""

             EligeCelda "text",mode,"left","","",60,LitObservaciones,"observaciones",2,""
             EligeCelda "text",mode,"left","","",60,LitAviso,"aviso",2,""
		%>
  </table>
  
  </div>   
  </div>


  	 <% 'DATOS COMERCIALES MODO AÑADIR %>
     <div  class="Section" id="S_AddDC">
        <a href="#" rel="toggle[AddDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
        <%=LitDatosComerciales%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
             </div>  </a>
    <div class="SectionPanel" id="AddDC" style="display: none">
        	 
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
        <%
		    'mmg: OrCU
		     if si_tiene_modulo_OrCU <> 0 then
                rstSelect.cursorlocation=3
		        rstSelect.open "select codigo, descripcion from tarifas with(nolock) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is null order by descripcion", 	session("dsn_cliente")
		        DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTarifaPre,"tarifa",rstSelect,"","codigo","descripcion","",""
			    rstSelect.close
			 else
                rstSelect.cursorlocation=3
			    rstSelect.open "select codigo, descripcion from tarifas with(nolock) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", 	session("dsn_cliente")
			    DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTarifa,"tarifa",rstSelect,"","codigo","descripcion","",""
			    rstSelect.close
			 end if
			 moneda_base=d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0",session("dsn_cliente"))
             rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, abreviatura from divisas with(nolock) where codigo like '" & session("ncliente") & "%' order by abreviatura",session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px","","",0,LitDivisa,"divisa",rstSelect,moneda_base,"codigo","abreviatura","",""
			 rstSelect.close

	    if si_tiene_modulo_OrCU <> 0 then
	            rstSelect.cursorlocation=3
	            rstSelect.open "select codigo, descripcion from tarifas with(nolock) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is not null order by descripcion", 	session("dsn_cliente")
	            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitDto1,"dtoCli1",rstSelect,"","codigo","descripcion","",""
	            rstSelect.close
			 
	            rstSelect.cursorlocation=3
	            rstSelect.open "select codigo, descripcion from tarifas with(nolock) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is not null order by descripcion", 	session("dsn_cliente")
  	            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitDto2,"dtoCli2",rstSelect,"","codigo","descripcion","",""
	            rstSelect.close

	            rstSelect.cursorlocation=3
	            rstSelect.open "select codigo, descripcion from tarifas with(nolock) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is not null order by descripcion", 	session("dsn_cliente")
	            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitDto3,"dtoCli3",rstSelect,"","codigo","descripcion","",""
	            rstSelect.close
	            'FLM:20090727:SI TIENE MÓDULO ORCU TIENE QUE TENER CRÉDITO PARA LOS SUMINISTROS Y UN CRÉDITO MÁXIMO.
		        IF si_tiene_modulo_OrCU<>0 then 

                    'DrawCelda "","","","",LitCreditoSuministro
                    'DrawCeldaResponsive1 "width100","","",0,LitCreditoSuministro
                    DrawDiv "3-sub","",""
                    DrawLabel "","",LitCreditoSuministro
                    CloseDiv
                
                    DrawInputCelda "","","",8,0,LitSaldoOffline,"saldooffline",""

                    DrawInputCelda "' onchange='javascript:saldoMaxChanged()","","",10,0,LitSaldoMax,"saldomax",""

                    DrawDiv "1","",""
                    DrawLabel "", "", LitSaldoSinLimite
                    DrawCheck "'' onclick='javascript: SaldoSinLimiteChanged()' value='1'", "", "cbSaldoSinLimite", ""
                    CloseDiv
                    
                    DrawInputCelda "' readonly='readonly","","",10,0,LitSaldoActual,"saldoact",""

			        %>
                    <input type="hidden" name="hd_saldoMax" value="<%=enc.EncodeForHtmlAttribute(saldomax & "")%>" />
                    <input type="hidden" name="saldoMaxOld" value="<%=enc.EncodeForHtmlAttribute(saldomax & "") %>" />
			        <input type="hidden" name="hd_saldoEnvidado" value="0" />
			        <input type="hidden" name="hd_saldoOffline" value="<%=enc.EncodeForHtmlAttribute(saldooffline & "")%>" /><%
			    else	
                    DrawInputCelda "","","",20,0,LitCredito,"credito",""
	            end if
        end if%>
	    <!--</table>
	    <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=2>--><%

             EligeCelda "input",mode,"","","",0,LitDescuento1,"descuento",4,"" 
			 EligeCelda "input",mode,"","","",0,LitDescuento2,"descuento2",4,"" 

             EligeCelda "input",mode,"","","",0,LitDescuento3,"descuento3",4,"" 

			 if si_tiene_modulo_EBESA <> 0 then
                 EligeCelda "input",mode,"","","",0,LitDtoImpFactura,"dtoimpfact",4,"" 
    	     end if

	    if si_tiene_modulo_petroleos<>0 then
                EligeCelda "input",mode,"","","",0,LitDescuentoLineal,"descuentoLineal",4,"" 

	    else
	        %><input type="hidden" name="descuentoLineal" value="0"/><%
	    end if
	    'FLM:20090429:añado campo para saber como se agrupan los suministros en las facturas: por cliente o por tarjeta...
	    if si_tiene_modulo_petroleos<>0 or si_tiene_modulo_OrCU<>0 then 
            DrawDiv "1","",""
            'DrawLabel "","",LitModFactSum
            
                %><label><%=LitModFactSum %></label><%
                %><select class="width60" name="modFactSum">
				    <option value="0"><%=LitAgrXcli%></option>
			        <option value="1"><%=LitSepXTar%></option>
			    </select><%
            CloseDiv
		else%>
		<input type="hidden" name="modFactSum" value="null"/>
		<%end if%>
	    
	    
	    <!--</table>
	    <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=2>-->
	    <%DrawFila color_blau
	         '' MPC 20/02/2008  Modificación para que a EBESA solo salgan como tipo de pagos aquellas cuyo campo copiasticket sea 2
             rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from formas_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitFormaPago,"fpago",rstSelect,fpUsuario    ,"codigo","descripcion","",""
			 rstSelect.close

			 strselect = "select codigo, descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion"
             rstSelect.cursorlocation=3
	    	 rstSelect.open strselect,session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago,"tpago",rstSelect,"","codigo","descripcion","",""
			 rstSelect.close
		CloseFila
        if si_tiene_modulo_EBESA <> 0 then
		    DrawFila color_blau
			    if si_tiene_modulo_ebesa <>0 then
			        strselect = "select codigo, descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' and copiasticket=2 order by descripcion"
			    else
			        strselect = "select codigo, descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion"
	        	end if
                rstSelect.cursorlocation=3
    		    rstSelect.open strselect,session("dsn_cliente")
			    DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago+" (" & LitNoVender & ")-1:","tpagonp1",rstSelect,"","codigo","descripcion","",""
			    rstSelect.movefirst
			    DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago+" (" & LitNoVender & ")-2:","tpagonp2",rstSelect,"","codigo","descripcion","",""

		    CloseFila
		    DrawFila color_blau
		        rstSelect.movefirst
			    DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago+" (" & LitNoVender & ")-3:","tpagonp3",rstSelect,"","codigo","descripcion","",""
			    rstSelect.close
		    CloseFila
		end if

		''cag dias pago******************************
		DrawInputCeldaActionDiv "","","","3",0,LitPrimerVen,"e_primer_ven",0, "onchange", "comprobar()",false

	  	DrawInputCeldaActionDiv "","","","3",0,LitSegunVen,"e_segundo_ven",0, "onchange", "comprobar()",false

	  	DrawInputCeldaActionDiv "","","","3",0,LitTercerVen,"e_tercer_ven",0, "onchange", "comprobar()",false

        DrawDiv "1","",""
        'DrawLabel "","",LitMesNoPago
        %><label><%=LitMesNoPago %></label><%
        %><select class="width60" name="mesNoPago">
					<option value="0" >0</option>
					<%opcionMenu=1
					while opcionMenu<=12 %>
					 <option value="<%=opcionMenu%>"> <%=opcionMenu%> </option>
					<%opcionMenu=opcionMenu+1
					wend %>
				 </select>
		<%
        CloseDiv

             EligeCelda "input",mode,"","","",0,LitRFinanciero,"recargo",8,"" 
             EligeCelda "check",mode,"","","",0,LitREquivalencia,"re",8,"" 


			DrawInputCeldaActionDiv "''","","",25,0,LitRiesMaxAut,"rgomaxaut",0,"onchange","ComprobarCantRiesgo()",false
			%><input type="hidden" name="rgomaxaut_ant" value="0"/>
			<input type="hidden" name="rcalc" value="0"/><%
            EligeCelda "Text",mode,"","","",0,LitRiesAlc,"",8,formatnumber(0,ndecimales,-1,0,-1) & " " & abreviatura

		if nz_b2(gestbono)=1 then
            EligeCelda "Text",mode,"","","",0,LitSaldoBonoMax,"",8,formatnumber(0,ndecimales,-1,0,-1) & " " & abreviatura
            EligeCelda "Text",mode,"","","",0,LitSaldoBono,"",8,formatnumber(0,ndecimales,-1,0,-1) & " " & abreviatura
		end if

			defecto=""
            ''Ricardo 25-07-2014 si tiene modulo nettit y ademas CADCOMSOLVERSUSCLI=1
            if si_tiene_modulo_NETTFI <> 0 and nz_b2(CADCOMSOLVERSUSCLI)=1 then
                defecto=comercialSolSusCli
            end if
			rstAux.cursorlocation=3
			rstAux.open "select dni, nombre from personal with(nolock),comerciales with(nolock) where personal.dni like '" & session("ncliente") & "%' and comerciales.comercial like '" & session("ncliente") & "%' and comerciales.fbaja is null and dni like '" & session("ncliente") & "%' and dni=comercial order by nombre",session("dsn_cliente")
			DrawSelectCeldaResponsive1 "width:200px","200","",0,iif(si_tiene_modulo_comercial<>0,LitComAsignadoModCom, LitComAsignado),"comasignado",rstAux,defecto,"dni","nombre","",""
			rstAux.close
			if si_tiene_modulo_comercial<>0 then
				defecto=""
				rstAux.cursorlocation=3
				rstAux.open "select codigo, nombre from agentes with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre",session("dsn_cliente")
				DrawSelectCeldaResponsive1 "width:200px","200","",0,LitAgenteAsignado,"agenteasignado",rstAux,defecto,"codigo","nombre","",""
				rstAux.close
			end if

			'cag
			cf=limpiaCadena(request.querystring("cf"))
			if cf="" then
				cf=limpiaCadena(request.form("cf"))
			end if

			if cstr(cf)="1" then
                EligeCelda "input",mode,"","","",0,LitNumCopiasFacturas,"ncopiasFacturas",1,"" 
			end if
			'fin cag

			rstSelect.cursorlocation=3
			rstSelect.open "select tipo_iva as codigo,tipo_iva as descripcion from tipos_iva with(nolock) order by tipo_iva",session("dsn_cliente")
			DrawSelectCeldaResponsive1 "width:50px",50,"",0,LitTipIvaCli,"iva",rstSelect,defecto,"codigo","descripcion","",""
			rstSelect.close
			'**RGU**
        
            DrawDiv "1","",""
         	%>
                <input type="hidden" name="distribuidor" value="<%=enc.EncodeForHtmlAttribute(TmpDistribuidor & "")%>"/>
			    <label><%=LIT_DISTCOLLAB%></label><%
                %><iframe id="frDistribuidores" name="fr_Distribuidores" src='distribuidores_clientes.asp?distribuidor=<%=enc.EncodeForJavascript(TmpDistribuidor)%>' class="width60 iframe-menu" frameborder="no" scrolling="no" noresize="noresize"></iframe>
            <%
            CloseDiv
		%>
                  </table>
   <%'DGB %>
       
                    
                <%  DrawDiv "3-sub","",""
                    DrawLabel "","",LITCONTA
                    CloseDiv
                
                EligeCelda "input",mode,"","","",0,LitCContable,"ccontable",25,"" 
                EligeCelda "input",mode,"","","",0,LitCContable_efecto,"ccontable_efecto",25,"" 

                DrawDiv "1","",""
                DrawLabel "","",LitIntracomunitario
                DrawCheck "'' onclick='javascript: iva0()'", "", "intra", ""
                CloseDiv
                        
                if SuplidosActivados=1 then
    			    DrawInputCelda "width=200px","","",25,0,LitCContable_suplidos,"CCONTABLE_SUPLIDOS",""
                end if
                %>
                    
            <table class="DataTable"></table>
    </div>
</div>
  	 <%'DATOS BANCARIOS MODO AÑADIR %>
      <div  class="Section" id="S_AddDB">
  <a href="#" rel="toggle[AddDB]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
    <%=LitDatosBancarios%>
  <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
       </div>  </a>
    <div class="SectionPanel" id="AddDB" style="display: none">
	
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
         
         

        <%if si_tiene_modulo_importaciones<>0 and false then
            DrawDiv "1","width:100px",""
        %>
                <a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=importaciones/bancos.asp&pag2=importaciones/bancos_bt.asp&ncliente=<%=enc.EncodeForJavascript(ncliente)%>&titulo=LISTA DE BANCOS&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" onmouseover="self.status='<%=LitIrBancos%>'; return true;" onmouseout="self.status=''; return true;"><%=LitBancos&ncliente%></a>
        <%
            CloseDiv
		end if
                
        	if si_tiene_modulo_OrCU<>0 then
                'DrawCeldaResponsive1 "width100","","",0,LitORCUGasOtros ''Hacer un salto de linea aquí
                DrawDiv "3-sub","",""
                DrawLabel "","",LitORCUGasOtros
                CloseDiv
            end if

            DrawInputCelda "","","",35,0,LitEntidad,"Entidad",""
            DrawInputCelda "' maxlength='50","","",35,0,LitDomBanco,"DomBanco",""

           
			%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                %><label><%=LitNumCuenta%></label><%
                %><div class="inlineTable width100"><%
                %><div class="width10 tableCell"><input class='' type="text" name="country" maxlength="2" onkeyup="if (this.value.length==2) document.clientes.iban.focus()" onblur="this.value=this.value.toUpperCase();"/></div><%
                %><div class="width10 tableCell"><input class='' type="text" name="iban" maxlength="2" onkeyup="if (this.value.length==2) document.clientes.NEntidad.focus()"/></div><%
			 	%><div class="width20 tableCell"><input class='' type="text" name="NEntidad" maxlength="4" onkeyup="if (this.value.length==4) document.clientes.Oficina.focus()"/></div><%
			 	%><div class="width20 tableCell"><input class='' type="text" name="Oficina" maxlength="4" onkeyup="if (this.value.length==4) document.clientes.DC.focus()"/></div><%
			 	%><div class="width10 tableCell"><input class='' type="text" name="DC" maxlength="2" onkeyup="if (this.value.length==2) document.clientes.Cuenta.focus()"/></div><%
			 	%><div class="width40 tableCell"><input class='' type="text" name="Cuenta"/></div><%
                %></div><%
			 %></div><%
			 
			 DrawInputCelda "' maxlength='16","","",35,0,LitNumTarjeta,"NumTarjeta",""
			 DrawInputCelda "' maxlength='11","","",11,0,LitBICSWIFT,"bic",""
			 DrawInputCelda "' maxlength='4","","",7,0,LitFCaducidad,"fcaducidad",""
             DrawCheckCelda "","","",0,LitDomiciliacion,"Domiciliacion","" 
 
		if si_tiene_modulo_OrCU<>0 then
			 rstSelect.cursorlocation=3
			 rstSelect.open "select b.nbanco,isnull(b.entidad,'')+'-'+isnull(norma,'')+isnull('-'+case when b.tipo_gasoleo='0' then '"+LitTodosGas+"' else te.descripcion end,'') as entidad from bancos b  with(nolock) left join tipos_entidades te  with(nolock) on te.codigo=b.tipo_gasoleo and te.codigo like '" & session("ncliente") &"%' where b.nbanco like '" & session("ncliente") & "%'"&" order by b.entidad",session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px",400,"",0,LitFormBancario,"formatoBanco",rstSelect,"","nbanco","entidad","",""
			 rstSelect.close
			'dgb 27/10/2009  Xenteo, indica forma de pago en poste
			 if modulo_Xenteo<>0 then
			    DrawCheckCelda "","","",0,LitPoste,"pagoA",""
			 end if
        end if
        %></table><%
        'FLM:20/01/2009:Datos bancarios duplicados para el tipo de gasoleo B.
		if si_tiene_modulo_OrCU<>0 or si_tiene_modulo_TGB<>0 then
            %><!--<div class="subsection"><%=LitORCUGasB%></div>
                <div class="subsectionpanel">--><%
            'DrawCeldaResponsive1 "width100","","",0,LitORCUGasB
            DrawDiv "3-sub","",""
            DrawLabel "","",LitORCUGasB
            CloseDiv

            if si_tiene_modulo_TGB<>0 then
            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                %><label><%=LitNumCuenta%></label><%
                %><div class="inlineTable width100"><%
                        %><div class="width10 tableCell"><input class='' type="text" name="countryGB" maxlength="2" onkeyup="if (this.value.length==2) document.clientes.ibanGB.focus()"  onblur="this.value=this.value.toUpperCase();"/></div><%
                        %><div class="width10 tableCell"><input class='' type="text" name="ibanGB" maxlength="2" onkeyup="if (this.value.length==2) document.clientes.NEntidadGB.focus()"/></div><%
                        %><div class="width20 tableCell"><input class='' type="text" name="NEntidadGB" maxlength="4" onkeyup="if (this.value.length==4) document.clientes.OficinaGB.focus()"/></div><%
			 	        %><div class="width20 tableCell"><input class='' type="text" name="OficinaGB" maxlength="4" onkeyup="if (this.value.length==4) document.clientes.DCGB.focus()"/></div><%
			 	        %><div class="width10 tableCell"><input class='' type="text" name="DCGB" maxlength="2" onkeyup="if (this.value.length==2) document.clientes.CuentaGB.focus()"/></div><%
			 	        %><div class="width40 tableCell"><input class='' type="text" name="CuentaGB"/></div><%
			    %></div><%
	        %></div><%

                    DrawInputCelda "' maxlength='11","","",11,0,LitBICSWIFT,"bicGB",""

                    DrawInputCelda "' maxlength='50","","",35,0,LitDomBanco,"DomBancoGB",""

                    DrawDiv "1","",""
                    
                    %><label><%=LitPoblacion%></label><%
                    %><input class="width50" type="text" size="25" name="PoblacionGB" value="" onchange="borrarCodigos('3')"/><%
			        %><input type="hidden" name="codPoblacionGB" value=""/> 
			        <input type="hidden" name="codProvinciaGB" value=""/>
			        <input type="hidden" name="codPaisGB" value=""/>
			        <input type="hidden" name="paisHGB" value="" /><%
				    %><a class='CELDAREFB' class="#dialog1" href="#SELECCIONAR_POBLACION2" onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes5&titulo=<%=LITSSVERPOBLACIONES %>');"  onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><%

                    CloseDiv

		 	        rstSelect.open "select idprovincia, descripcion from PROVINCIAS with(nolock) order by descripcion",DSNIlion,adOpenKeyset,adLockOptimistic
			        DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitProvincia,"gb_provincia",rstSelect,"","idprovincia","descripcion","onchange","TraerProvinciaGB()"
			        rstSelect.close

            end if


			     rstSelect.cursorlocation=3
			     rstSelect.open "select b.nbanco,isnull(b.entidad,'')+'-'+isnull(norma,'')+isnull('-'+case when b.tipo_gasoleo='0' then '"+LitTodosGas+"' else te.descripcion end,'') as entidad from bancos b  with(nolock) left join tipos_entidades te  with(nolock) on te.codigo=b.tipo_gasoleo and te.codigo like '" & session("ncliente") &"%' where b.nbanco like '" & session("ncliente") & "%'"&" order by b.entidad",session("dsn_cliente")
			     DrawSelectCeldaResponsive1 "width:200px",400,"",0,LitFormBancario,"formatoBanco2",rstSelect,"","nbanco","entidad","",""
			     rstSelect.close
			     'dgb 27/10/2009  Xenteo, indica forma de pago en poste
			     if modulo_Xenteo<>0 then
			        DrawCheckCelda "","","",0,LitPoste,"pagoB",""
			     end if

                 DrawInputCelda "' maxlength='16","","",35,0,LitNumTarjeta,"NumTarjeta2",""

			     DrawInputCelda "' maxlength='4","","",7,0,LitFCaducidad,"fcaducidad2",""
                %><!--</div>-->
                
            <table class="DataTable"></table><%
         end if
        'FIN OTROS DATOS BANCARIOS PARA OTRO TIPO DE GASOLEO: MÓDULO ORCU.
		%>
 
  </div>
  </div>
  	 <%'DIRECCION ENVIO MODO AÑADIR%>
     <div   class="Section" id="S_AddDE">
    <a href="#" rel="toggle[AddDE]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitDireccionEnvio%>
    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
    </div></a>
    
  	 <%if viene="facturas_cli_E" then%>
     	 <div class="SectionPanel" id="AddDE" style="display: ">
         <%else%>
          <div class="SectionPanel" id="AddDE" style="display: none">
         <%end if%>
	
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=2>


			<div class="tableCell" style="width:100%">
				<%if si_tiene_modulo_agrario<>0 then
					%><a class="reflink" href="javascript:CopiarCampos()"><%=LitCopiarDirEnvioSocio%></a><%
				else
					%><a class="reflink" href="javascript:CopiarCampos()"><%=LitCopiarDirEnvio%></a><%
				end if%>
			</div>
	    <%
	
			DrawInputCelda "width:200px","","",35,0,LitDomicilio,"de_domicilio",""

            DrawDiv "1","",""
			%><label><%=LitPoblacion%></label><% 
            %><input class="width50" type="text" size="25" name="de_poblacion" onchange="borrarCodigos('2')"/><a class='CELDAREFB' onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes2&titulo=<%=LITSSVERPOBLACIONES %>');" href="#SELECCIONAR_POBLACION2"   onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			<input type="hidden" name="de_codPoblacion" value=""/> 
			<input type="hidden" name="de_codProvincia" value=""/>
			<input type="hidden" name="de_codPais" value=""/>
			<input type="hidden" name="de_paisH" value="" />
			
			<% 'AbrirModal "SELECCIONAR_POBLACION1","../configuracion/poblaciones.asp?mode=buscar&viene=clientes2&titulo=SELECCIONAR POBLACION",AnchoVentana,AltoVentana,"no","si","no","si",LitBuscar%>
			<%
            CloseDiv
                
			 DrawInputCelda "width:100","","",5,0,LitCP,"de_cp",""

             rstSelect.cursorlocation=3
			 rstSelect.open "select idprovincia, descripcion from PROVINCIAS with(nolock) order by descripcion",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitProvincia,"de_provinciaDDL",rstSelect,"","idprovincia","descripcion","onchange","TraerProvinciaDe()","de_provincia",25,""
			 rstSelect.close

             rstSelect.cursorlocation=3
		     rstSelect.open "select id, nombre from PAISES with(nolock) order by nombre",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitPais,"de_paisDDL",rstSelect,"","id","nombre","onchange","TraerPaisDe()","de_pais",30,""
			 rstSelect.close

             DrawInputCelda "width:100","","",20,0,LitTel1,"de_telefono",""%>
  </table>
  <% ' DGM 19/9/11 Ocultamos la direccion de envio de factura %>
  <div style="display:none" >
  <% 'DGM 28/7/11 Para la dirección de envío de la factura %>
  <table width="100%"  border='0' cellpadding=0 cellspacing=0>
        <%
        DrawCeldaResponsive1 "width100","","",0,LITINVOICEADDRESS
            %>
  </table>
       <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=2>

				<%DrawDiv "1","",""
                if si_tiene_modulo_agrario<>0 then
					%><a class="reflink" href="javascript:CopiarCamposF()"><%=LitCopiarDirEnvioSocio%></a><%
				else
					%><a class="reflink" href="javascript:CopiarCamposF()"><%=LitCopiarDirEnvio%></a><%
				end if
                CloseDiv        
                %>
	    <%
			 DrawInputCeldaResponsive1 "width:100","","",35,0,LitDomicilio,"de_domicilioF",""

			%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><% 
            %><label><%=LitPoblacion%></label><%
            %><input class="width50" type="text" size="25" name="de_poblacionF" onchange="borrarCodigosF('2')"/><a class='CELDAREFB' onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes4&titulo=<%=LITSSVERPOBLACIONES %>');" href="#SELECCIONAR_POBLACION2"   onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"></a>
			<input type="hidden" name="de_codPoblacionF" value=""/> 
			<input type="hidden" name="de_codProvinciaF" value=""/>
			<input type="hidden" name="de_codPaisF" value=""/>
			<input type="hidden" name="de_PaisFH" value="" />
			<% 'AbrirModal "SELECCIONAR_POBLACION1","../configuracion/poblaciones.asp?mode=buscar&viene=clientes2&titulo=SELECCIONAR POBLACION",AnchoVentana,AltoVentana,"no","si","no","si",LitBuscar%>
			</div><div class="ui-widget" ><%

			 DrawInputCelda "width:100","","",5,0,LitCP,"de_cpF",""

			 rstSelect.cursorlocation=3
			 rstSelect.open "select idprovincia, descripcion from PROVINCIAS with(nolock) order by descripcion",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitProvincia,"de_provinciaDDLF",rstSelect,"","idprovincia","descripcion","onchange","TraerProvinciaDeF()","de_provinciaF",25,""
			 rstSelect.close

             rstSelect.cursorlocation=3
			 rstSelect.open "select id, nombre from PAISES with(nolock) order by nombre",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitPais,"de_paisDDLF",rstSelect,"","id","nombre","onchange","TraerPaisDeF()","de_paisF",30,""
			 rstSelect.close

			 DrawInputCeldaResponsive2 "width:100","","",20,0,LitTel1,"de_telefonoF",""%>
  </table>
  </div>
  
  </div>
  </div>

  <% 'OTROS DATOS MODO AÑADIR %>
   <div   class="Section" id="S_AddOD">
    <a href="#" rel="toggle[AddOD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitOtrosDatos%>
      <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
    </div></a>
    <div class="SectionPanel" style="display:none " id="AddOD">
    
        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">

        <%'jcg 02/02/2008
        if si_tiene_modulo_proyectos<>0 then

            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                %><label><%=iif(mode="browse","<b>"+LitProyecto+"</b>",LitProyecto+"")%></label><%
                %><input class="width100" type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(tmpProyecto & "")%>"/><%
                %><iframe id='frProyecto' name="fr_Proyecto" src='../mantenimiento/docproyectos.asp?viene=clientes&mode=<%=enc.EncodeForHtmlAttribute(mode)%>&cod_proyecto=<%=enc.EncodeForHtmlAttribute(tmpProyecto)%>' class="width60 iframe-menu" frameborder="no" scrolling="no" noresize="noresize"></iframe><%
			%></div>
        </table>
        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5">
			<%
        end if
		
			rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from tipo_actividad with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTActividad,"tactividad",rstSelect,"","codigo","descripcion","",""
			 rstSelect.close

			rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from zonas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitZona,"zona",rstSelect,"","codigo","descripcion","",""
			 rstSelect.close


			 DrawInputCeldaResponsive2 "width:200px","","",25,0,LitTransportista,"transportista",""

			 defecto=""

            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                %><label><%=LitPortes%></label><select class="width60" name="portes">
						<%if defecto=LitDebidos then
							%><option selected="selected" value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitPagados then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option selected="selected" value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitMedios then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option selected="selected" value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitOtros then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option selected="selected" value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitPagosCargFact then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option selected="selected" value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						else
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option selected="selected" value=""></option><%
						end if%>
					    </select>
				</div><%

			 DrawInputCelda "","","",25,0,LitHmanyana,"hmanyana",""

			 DrawInputCelda "","","",20,0,LitHTarde,"htarde",""

			 DrawInputCelda "","","",6,0,LitPht,"pht",""

			 DrawInputCelda "","","",6,0,LitPkm,"pkm",""

			 DrawInputCelda "","","",6,0,LitPD,"pd",""

			rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from tipos_entidades with(NOLOCK) where codigo like '" & session("ncliente") & "%' and tipo='" & LitCLIENTE & "' order by descripcion", 	session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:100",200,"",0,iif(si_tiene_modulo_agrario<>0,LitTSocio,LitTCliente),"tipo_cliente",rstSelect,"","codigo","descripcion","",""

			 rstSelect.close

		if si_asesoria=true then

				 DrawInputCelda "' maxlength='25","","",25,0,LitCarpetaNominas,"dsn_nominaplus",""

				 rstSelect.cursorlocation=3
				 rstSelect.open "select codigo, descripcion from periodos with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
				 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitPeriodicidadFact,"periodicidad",rstSelect,"","codigo","descripcion","",""
				 rstSelect.close

                 DrawCheckCelda "","","",0,LitMostrarListPortal,"mostrarListPortal",True

		end if
		if si_tiene_modulo_agrario<>0 then
				 DrawInputCeldaResponsive2 "width:200px","","",12,0,LitFNacimiento,"fnacimiento",""
				 DrawInputCeldaResponsive2 "width:200px","","",25,0,LitSegSocial,"segsocial",""
		elseif si_asesoria=true then
				 DrawInputCeldaResponsive2 "width:200px","","",25,0,LitSegSocial,"segsocial",""
		end if
		if si_asesoria=true then
				rstSelect.cursorlocation=3
				 rstSelect.open "select codigo, descripcion from tiendas with (nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
				 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSucursal,"sucursal",rstSelect,"","codigo","descripcion","",""
				 rstSelect.close
		end if
		
		if si_tiene_modulo_CRMComunicacion <> 0 then
                DrawCheckCelda "","","",0,LITDELIVERYADV,"submit_advertising",""
                DrawCheckCelda "","","",0,LITCOMMEMAIL,"email_communication",""
                DrawCheckCelda "","","",0,LITCOMMSMS,"sms_communication",""
		 end if
		%>
  </table>
  
  </div>
  </div>

  <% 'CONFIG DOC MODO AÑADIR %>

   <div   class="Section" id="S_AddDD" >
    <a href="#" rel="toggle[AddDD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitConfDoc2%>
      <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
    </div></a>
    <div class="SectionPanel" style="display:none " id="AddDD">
    
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=2>
        <%             
        'AMF:2/11/2010:Cambiada la obtencion de las series por una llamada a procedure.
        set connSeries = Server.CreateObject("ADODB.Connection")
	    set commandSeries = Server.CreateObject("ADODB.Command")
	               
	    connSeries.open session("dsn_cliente")
	                    
        commandSeries.ActiveConnection = connSeries
	    commandSeries.CommandTimeout = 0
	    commandSeries.CommandText="ObtenerNombreSeriesParaSDefecto"
	    commandSeries.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	    
	    commandSeries.Parameters.Append commandSeries.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente") & "")

			EligeCelda "check", mode,"width:200px","","",0,LitValorPre,"valorado_pre",0,iif(valorado_pre="",-1,nz_b(valorado_pre))

			commandSeries.Parameters.Append commandSeries.CreateParameter("@tipodoc",adVarChar,adParamInput,50,"PRESUPUESTO A CLIENTE")
			set rstAux=commandSeries.Execute
			DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSeriePre,"serie_pre",rstAux,serie_pre,"nserie","nombre","",""
			rstAux.close

			EligeCelda "check", mode,"width:200px","","",0,LitValorPed,"valorado_ped",0,iif(valorado_ped="",-1,nz_b(valorado_ped))

			commandSeries.Parameters("@tipodoc")="PEDIDO DE CLIENTE"
			set rstAux=commandSeries.Execute			
	 		DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSeriePed,"serie_ped",rstAux,serie_ped,"nserie","nombre","",""
			rstAux.close

			EligeCelda "check", mode,"width:200px","","",0,LitValorAlb,"valorado_alb",0,iif(valorado_alb="",-1,nz_b(valorado_alb))

			commandSeries.Parameters("@tipodoc")="ALBARAN DE SALIDA"
			set rstAux=commandSeries.Execute
            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSerieAlb,"serie_alb",rstAux,serie_alb,"nserie","nombre","",""
			rstAux.close

			commandSeries.Parameters("@tipodoc")="FACTURA A CLIENTE"
			set rstAux=commandSeries.Execute
            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSerieFac,"serie_fac",rstAux,serie_fac,"nserie","nombre","",""
			rstAux.close

		'AMF:29/10/2010:Serie de la incidencia, lista desplegable.
		if ModuloContratado(session("ncliente"),ModPostVenta) <> 0 then
			commandSeries.Parameters("@tipodoc")="INCIDENCIA"
			set rstAux=commandSeries.Execute
            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSerieIncidencia,"serie_incidencia",rstAux,serie_incidencia,"nserie","nombre","",""
			rstAux.close
		end if
        set commandSeries=nothing
        connSeries.close
	    set connSeries = nothing%>
  </table>
  
  
  </div>
  </div>

  <% 'CAMPOS PERSONALIZABLES MODO AÑADIR %>
   <div   class="Section"  id="S_AddCP">
    <a href="#" rel="toggle[AddCP]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitCampPersoCli%>
      <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
    </div></a>
    <div class="SectionPanel" style="display: none" id="AddCP">
     
	<table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5"><%
      	
		rst.cursorlocation=3
		rst.open "select * from camposperso with(NOLOCK) where tabla='CLIENTES' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
		if not rst.eof then
			num_campos_existen=rst.recordcount
			DrawFila ""
				num_campo=1
				num_campo2=1
				num_puestos=0
				num_puestos2=0
				while not rst.eof
					if num_puestos2>0 and (num_puestos2 mod 2)=0 then
						'DrawCelda "CELDA7 style='width:125px'","","",0,"&nbsp;"
						'CloseFila
						'DrawFila ""
						num_puestos2=0
					end if

                    if rst("titulo") & "" <>"" then
                       datosTitulo=enc.EncodeForHtmlAttribute(rst("titulo") & "")
                    else
                       datosTitulo=rst("titulo")
                    end if

					if rst("titulo") & "">"" and nz_b(rst("system_reg")) <> -1 then
						if ((num_puestos-1) mod 2)=0 then
							'DrawCelda "CELDA7 style='width:155px'","","",0,"&nbsp;"
						end if
						num_puestos=num_puestos+1
						num_puestos2=num_puestos2+1

						valor_campo_perso=""
						if rst("tipo")=1 then
							if isNumeric(rst("tamany")) then
								tamany=rst("tamany")
							else
								tamany=1
							end if
                            DrawInputCelda "' maxlength='" & tamany,"","",35,0,datosTitulo,"campo" & enc.EncodeForHtmlAttribute(num_campo),enc.EncodeForHtmlAttribute(valor_campo_perso)
						elseif rst("tipo")=2 then
                            DrawCheckCelda "","","",0,datosTitulo,"campo" & enc.EncodeForHtmlAttribute(num_campo), iif(valor_campo_perso="on",True,"")
						elseif rst("tipo")=3 then
							num_campo_str=cstr(num_campo)
							if len(num_campo_str)=1 then
								num_campo_str="0" & num_campo_str
							end if
							'response.Write "num_campo" & num_campo
							'mmg 29/05/2008: comprobamos si el campo es 01,02 o 03 y hacemos la select que corresponda
                            sem=0
                            if (num_campo_str="01" and c01="c") or (num_campo_str="02" and c02="c") or (num_campo_str="03" and c03="c") or (num_campo_str="04" and c04="c") or (num_campo_str="05" and c05="c") or (num_campo_str="06" and c06="c") or (num_campo_str="07" and c07="c") or (num_campo_str="08" and c08="c") or (num_campo_str="09" and c09="c") or (num_campo_str="10" and c10="c") or (num_campo_str="11" and c11="c") or (num_campo_str="12" and c12="c") or (num_campo_str="13" and c13="c") or (num_campo_str="14" and c14="c") or (num_campo_str="15" and c15="c") or (num_campo_str="16" and c16="c") or (num_campo_str="17" and c17="c") or (num_campo_str="18" and c18="c") or (num_campo_str="19" and c19="c") or (num_campo_str="20" and c20="c") then
                                strSelListVal="select dni as ndetlista, nombre as valor from personal with(NOLOCK),comerciales with(NOLOCK) where personal.dni like '" & session("ncliente") & "%' and comerciales.comercial like '" & session("ncliente") & "%' and comerciales.fbaja is null and dni like '" & session("ncliente") & "%' and dni=comercial order by valor,ndetlista"
                                sem=1
                            else
                                strSelListVal="select ndetlista,valor from campospersolista with(NOLOCK) where tabla='CLIENTES' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
                            end if

							rstAux.cursorlocation=3
							rstAux.open strSelListVal,session("dsn_cliente")
							'sera este
							'DrawSelectCelda "CELDA7","200","",0,"","campo" & num_campo,rstAux,valor_campo_perso,"ncampo","titulo","",""
							'o bien este
							%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                %><label><%=datosTitulo%></label><select class="width200" name="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" id="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>">
									<%encontrado=0
									while not rstAux.eof
									
									if sem=0 then
			                            if valor_campo_perso & "">"" and isnumeric(valor_campo_perso) then
											valor_campo_perso_aux=clng(valor_campo_perso)
										else
											valor_campo_perso_aux=0
										end if
										if valor_campo_perso_aux=clng(rstAux("ndetlista")) then
											texto_selected="selected"
											if encontrado=0 then encontrado=1
										else
											texto_selected=""
										end if
			                        else
			                            if valor_campo_perso & "">"" then
				                            valor_campo_perso_aux=valor_campo_perso
			                            else
				                            valor_campo_perso_aux="0"
			                            end if
			                            if valor_campo_perso_aux=rstAux("ndetlista")&"" then
				                            texto_selected="selected"
				                            if encontrado=0 then encontrado=1
			                            else
				                            texto_selected=""
			                            end if
			                        end if%>
										<option value="<%=enc.EncodeForHtmlAttribute(rstAux("ndetlista") & "")%>"  <%=enc.EncodeForHtmlAttribute(texto_selected)%> ><%=enc.EncodeForHtmlAttribute(rstAux("valor"))%></option>
										<%rstAux.movenext
									wend%>
									<option <%=iif(encontrado=1,"","selected")%> value=""></option>
								</select>
							</div><%
							rstAux.close
						elseif rst("tipo")=4 then
							if isNumeric(rst("tamany")) then
								tamany=rst("tamany")
							else
								tamany=1
							end if
                            DrawDiv "1","",""
                            DrawLabel "","",datosTitulo
                            DrawInput "", "", "campo" & enc.EncodeForHtmlAttribute(num_campo), enc.EncodeForHtmlAttribute(valor_campo_perso), "size='35' maxlength='" & tamany & "'"
                            CloseDiv
						elseif rst("tipo")=5 then
							if isNumeric(rst("tamany")) then
								tamany=rst("tamany")
							else
								tamany=1
							end if                            
                            DrawDiv "1","",""
                            DrawLabel "","",datosTitulo
                            DrawInput "", "", "campo" & enc.EncodeForHtmlAttribute(num_campo), enc.EncodeForHtmlAttribute(valor_campo_perso), "size='30' maxlength='" & tamany & "'"
                            DrawCalendar "campo" & enc.EncodeForHtmlAttribute(num_campo)
                            CloseDiv
						end if
					else
						%><input type="hidden" name="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" id="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" value=""/><%
					end if
					%><input type="hidden" name="tipo_campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" value="<%=enc.EncodeForHtmlAttribute(rst("tipo") & "")%>"/>
					<input type="hidden" name="titulo_campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" value="<%=datosTitulo%>"/><%
					rst.movenext
					num_campo=num_campo+1
					if not rst.eof then
						if rst("titulo") & "">"" then
							num_campo2=num_campo2+1
						end if
					end if
				wend
			num_campos=num_puestos
		else
			num_campos=0
			num_campos_existen=0
		end if
		rst.close
	%></table>
	<input type="hidden" name="num_campos" value="<%=enc.EncodeForHtmlAttribute(num_campos_existen & "")%>"/>

</div>
</div>

    <%BarraNavegacion "Add" %>
    <%elseif mode="browse" then
    	no_mostrar=0

        'DGM 27/7/11 Comprobamos si tiene tarjetas
        rst.cursorlocation=3
        rst.Open "select count(pan) as total from tarjetas with(NOLOCK) where ncliente like '" & session("ncliente") & "%' and ncliente ='" & ncliente &"'",DsnIlion
        if not rst.eof then
            if clng(rst("total")) > 0 then
                tiene_tarjetas = 1
            else
                tiene_tarjetas = 0
            end if
        else
            tiene_tarjetas = 0
        end if
        rst.Close

        ' Comprobamos que tenga alguna anotación en el historial
        rst.cursorlocation=3
        rst.Open "select count(nanotacion) as total from historial_cliente with(nolock) where ncliente like '" &session("ncliente") & "%' and ncliente = '" & ncliente & "'", session("dsn_cliente")
        if not rst.eof then
            if clng(rst("total")) > 0 then
                tiene_historial = 1
            else
                tiene_historial = 0
            end if
        else
            tiene_historial = 0
        end if
        rst.Close
        
        
        ''ricardo 31/7/2003 comprobamos que existe el cliente
        if he_borrado<>1 then
		    rstAux.cursorlocation=3
            StrselCli="select ncliente from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'"
            if comercialSolSusCli & "">"" then
                StrselCli=StrselCli & " and comercial='" & comercialSolSusCli & "' "
            end if
''response.write("el StrselCli es-" & StrselCli & "-<br>")
		    rstAux.open StrselCli, session("dsn_cliente")
		    if rstAux.eof then
			    ncliente=""%>
			    <script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgDocsNoExiste%>");
                    document.clientes.action = "clientes.asp?mode=add"
                    document.clientes.submit();
                    parent.botones.document.location = "clientes_bt.asp?mode=add";
			    </script>
			    <%mode="add"
			    no_mostrar=1
		    end if
		    rstAux.close
        end if

        if no_mostrar=0 then
	        rstSelect.cursorlocation=3
    	    rstSelect.Open "select * from documentos_cli with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "'", session("dsn_cliente")
	        if rstSelect.eof then%>
		        <script language="javascript" type="text/javascript">
                    window.alert("<%=LitClieteNoPuedMostContDist%>");
                    document.clientes.action = "clientes.asp?mode=add"
                    document.clientes.submit();
                    parent.botones.document.location = "clientes_bt.asp?mode=add";
    		    </script>
        	    <%no_mostrar=1
	        end if
	        rstSelect.close
        end if

	if no_mostrar=0 then
		rst.cursorlocation=3
		  'i(EJM 25/09/2006) la consulta devuelve el artículo relacionado con la zona. Necesario para mostrar el precio de los portes si estos son debidos.
		  'i(EJM 25/09/2006) la consulta devuelve la descripción de la divisa
		  strSelect="select c.*,do.*,(select referencia from zonas where codigo = zona) as referencia, " & _
						" (select descripcion from divisas where codigo=divisa) as divisa,d.abreviatura,d.ndecimales " & _
						" from clientes c " & _
						" join domicilios do on c.ncliente like '" & session("ncliente") & "%' and do.pertenece like '" & session("ncliente") & "%' and ncliente='" & ncliente & "' And do.codigo=c.dir_principal " & _
						" left join divisas d on d.codigo=c.divisa"

		  rst.Open strSelect,session("dsn_cliente")
        if not rst.eof then
		    ''ricardo 3-9-2004 para ver foto en clientes
    	 	if rst("tipo_foto")>"" and not isnull(rst("tipo_foto")) then
	        	mostrar_foto=1
	 	    else
		    	mostrar_foto=0
    	 	end if
        else
	        mostrar_foto=0
        end if
		if viene<>"facturas_cli_E" then
            if mode <> "browse" then %><div class="headers-wrapper"></div><% end if
		  BarraOpciones "browse", ncliente
		end if%>
		  <input type="hidden" name="hncliente" value="<%=rst("ncliente")%>"/>
		<%
		rstAux.cursorlocation=3
		'Si el distribuidor tiene nombre el cliente no se corresponde con el actual, cambiamos la consulta
		col = 0
		rstAux.open "select name from distribuidores with(nolock) where ndist like '" & session("ncliente") &"%' and ndist = '" & rst("ndist") & "'",session("dsn_cliente")
		if not rstAux.eof then
		    if rstAux("name")&"" > "" then
		        col = 1
		    else
		        col = 0
		    end if
		end if
		rstAux.close

		if col = 0 then
		    rstAux.open "select ndist from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente")
		    if not rstAux.eof then
			    ndist_aux=rstAux("ndist")
		    else
			    ndist_aux=""
		    end if
		rstAux.close
		else
		    ndist_aux = rst("ndist")
		end if
		
		%><input type="hidden" name="ndist" value="<%=enc.EncodeForHtmlAttribute(ndist_aux & "")%>"/><%
		ndist_aux=""

		 %>
            <%

		 'DATOS GENERALES MODO BROWSE


		' Inicio Borde Span
		%>
         <table width="100%">
         <div id="CollapseSection">
    <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['BrowseDG', 'BrowseDC', 'BrowseDB','BrowseDE','BrowseOD','BrowseDD','BrowseCP']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll %>" alt="" title="" <%=ParamImgCollapseAll %> /></a> 
    <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['BrowseDG', 'BrowseDC', 'BrowseDB','BrowseDE','BrowseOD','BrowseDD','BrowseCP']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll %>" alt="" title="" <%=ParamImgCollapseAll %> /></a>
    </div>
             <table></table>
            

    <div class="Section" id="S_BrowseDG">
        <a href="#" rel="toggle[BrowseDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader displayed">
            <%=LitDatosGenerales%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" style="border:0;" />
        </div>
        </a>
        
    <div class="SectionPanel" style="" id="BrowseDG">
        <table border='0' cellpadding="1" cellspacing="1" width='100%'>
		  
		    <% 
              if rst("ncomercial") & "" <>"" then
                  datosNComercial=enc.EncodeForHtmlAttribute(rst("ncomercial") & "")
              else
                  datosNComercial=rst("ncomercial")
              end if
              clase="span-browser"
              DrawCeldaResponsiveLabel clase,"",LitNComercial,datosNComercial

            DrawDiv "1", "", ""
            CloseDiv
			if si_asesoria=true then

              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFormaJuridica,"",d_lookup("descripcion","formas_juridicas","codigo like '" & session("ncliente") & "%' and codigo='"&rst("fjuridica")&"'",session("dsn_cliente"))

              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTitular,"",enc.EncodeForHtmlAttribute(rst("titular") & "")

			end if

            
            DrawCeldaResponsiveLabel clase,"txtMandatory",LitCIF,enc.EncodeForHtmlAttribute(rst("cif") & "")

            
   		  
	   		   pushtocall2= d_lookup("pushtocall", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente"))    		   
	   		   if pushtocall2>0 and not isnull(rst("telefono")) and modcentralita>"" then	   		     
	   		     pretocall= d_lookup("pretocall", "configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente"))  	   		    
	   		     if isnull(pretocall) then
	   		        linkTelf = "<a> <img name='telffijo'  style='cursor: pointer;' src='../images/"+ImgTelfLlamar+"' "+ ParamImgTelefono +" alt='"+LitLlamar+"' title='"+LitLlamar+"' onclick="+chr(34)+"javascript:RealizarLlamada('"+ rst("telefono")+"','telffijo','"+CStr(pushtocall2)+"');"+chr(34)+"></img> </a> "
	   		     else		   	
                    linkTelf = "<a> <img name='telffijo'  style='cursor: pointer;' src='../images/"+ImgTelfLlamar+"' "+ ParamImgTelefono +" alt='"+LitLlamar+"' title='"+LitLlamar+"' onclick="+chr(34)+"javascript:RealizarLlamada('"+pretocall+rst("telefono")+"','telffijo','"+CStr(pushtocall2)+"');"+chr(34)+"></img> </a> "
                 end if
               else
                linkTelf = ""
               end if

              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTel1,"","<a>" & rst("telefono") & "</a>" + " " + linkTelf

              if rst("contacto") & "" <>"" then
                  datosContacto=enc.EncodeForHtmlAttribute(rst("contacto") & "")
              else
                  datosContacto=rst("contacto")
              end if
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitContacto,"",datosContacto

               if pushtocall2 and not isnull(rst("telefono2")) and modcentralita>"" then                
               
                 if isnull(pretocall) then               
	   		        linkMovil = "<a> <img name='telfmovil' style='cursor: pointer;' src='../images/"+ImgTelfLlamar+"' "+ ParamImgTelefono +" alt='"+LitLlamar+"' title='"+LitLlamar+"' onclick="+chr(34)+"javascript:RealizarLlamada('"+ rst("telefono2")+"','telfmovil','"+CStr(pushtocall2)+"');"+chr(34)+"></img> </a>" 
                 else
	   		        linkMovil = "<a> <img name='telfmovil' style='cursor: pointer;' src='../images/"+ImgTelfLlamar+"' "+ ParamImgTelefono +" alt='"+LitLlamar+"' title='"+LitLlamar+"' onclick="+chr(34)+"javascript:RealizarLlamada('"+pretocall + rst("telefono2")+"','telfmovil','"+CStr(pushtocall2)+"');"+chr(34)+"></img> </a>" 
                 end if	   		        
	   		   else
	   		     linkMovil = "" 
	   		   end if
               
               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTel2,"","<a href=javascript:AbrirVentana('../servicios/mensajes_sms.asp?mode=add&numero="&rst("telefono2")&"','P',270,430); style='text-decoration:underline;' title='" & LITSSENVIARSMS & "' onmouseover="&cd&"self.status='" & LITSSENVIARSMS & "'; return true;"&cd&">" & rst("telefono2") & "</a>" + "    " +linkMovil 
               
               if rst("domicilio") & "" <>"" then
                  datosDomicilio=enc.EncodeForHtmlAttribute(rst("domicilio") & "")
               else
                  datosDomicilio=rst("domicilio")
               end if

               if rst("poblacion") & "" <>"" then
                  datosPoblacion=enc.EncodeForHtmlAttribute(rst("poblacion") & "")
               else
                  datosPoblacion=rst("poblacion")
               end if

               DrawCeldaResponsiveLabel clase,"txtMandatory",LitDomicilio + " " +LinkGoogleMap(enc.EncodeForHtmlAttribute(rst("rsocial") & ""),datosDomicilio, datosPoblacion, rst("cp"), rst("provincia"), rst("pais"),0),datosDomicilio

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFAX,"",enc.EncodeForHtmlAttribute(rst("fax")& "")

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPoblacion,"",enc.EncodeForHtmlAttribute(datosPoblacion& "")

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCP,"",enc.EncodeForHtmlAttribute(rst("cp")& "")

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitProvincia,"",enc.EncodeForHtmlAttribute(rst("provincia")& "")

               if rst("web") & "" <>"" then
                  datosWeb=enc.EncodeForHtmlAttribute(rst("web") & "")
               else
                  datosWeb=rst("web")
               end if

               if mid(rst("web"),1,7) = "http://" or mid(rst("web"),1,8) = "https://" then
	   		       EligeCeldaResponsive "text",mode,clase,"","",0,"",LitWeb,"","<a href=javascript:AbrirVentana('"& datosWeb & "','A','640','950'); onmouseover="&cd&"self.status='" & LitWeb & "'; return true;"&cd&" onmouseout="&cd&"self.status=''; return true;"&cd&">" & datosWeb & "</a>"     
	   		   else
	   		       EligeCeldaResponsive "text",mode,clase,"","",0,"",LitWeb,"","<a href=javascript:AbrirVentana('http://" & datosWeb & "','A','640','950'); onmouseover="&cd&"self.status='" & LitWeb & "'; return true;"&cd&" onmouseout="&cd&"self.status=''; return true;"&cd&">" & datosWeb & "</a>"
	   		   end if

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPais,"",enc.EncodeForHtmlAttribute(rst("pais") & "")

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitEMail,"","<a href=javascript:AbrirVentana('../servicios/enviaremail.asp?mode=add&dedonde=CLIENTES&destinatario=" & ncliente & "','A','640','950'); onmouseover="&cd&"self.status='" & LITSSENVIAREMAIL & "'; return true;"&cd&" onmouseout="&cd&"self.status=''; return true;"&cd&">" & rst("email") & "</a>"

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFechaAlta,"",enc.EncodeForHtmlAttribute(rst("falta") & "")

               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFechaBaja,"",enc.EncodeForHtmlAttribute(rst("fbaja") & "")

               if rst("observaciones") & "" <>"" then
                  datosObservaciones=enc.EncodeForHtmlAttribute(rst("observaciones") & "")
               else
                  datosObservaciones=rst("observaciones")
               end if
               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitObservaciones,3,enc.EncodeForHtmlAttribute(pintar_saltos_espacios(datosObservaciones&""))

               if rst("aviso") & "" <>"" then
                  datosAviso=enc.EncodeForHtmlAttribute(rst("aviso") & "")
               else
                  datosAviso=rst("aviso")
               end if
               EligeCeldaResponsive "text",mode,clase,"","",0,"",LitAviso,3,enc.EncodeForHtmlAttribute(pintar_saltos_espacios(datosAviso&""))
		    %>
		    </table>
		  
	</div>
    </div>


		<% 'DATOS COMERCIALES MODO BROWSE %>
    <div   class="Section" id="S_BrowseDC" >
        <a href="#" rel="toggle[BrowseDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosComerciales%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
        </div>
        </a>
    <div class="SectionPanel" style="display: none" id="BrowseDC">
	
	     <table width="100% "bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=5>
      	  
	        <%
                clase="span-browser"

				 if rst("tarifa")>"" then
				 	Tarifa = d_lookup("descripcion","tarifas","codigo like '" & session("ncliente") & "%' and codigo='" + rst("tarifa") + "'",session("dsn_cliente"))
				 else
				 	Tarifa=""
				 end if
                 'DrawCelda2 "CELDA width='30%'","left", false, Tarifa
                 if si_tiene_modulo_OrCU <> 0 then
		            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTarifaPre,"",enc.EncodeForHtmlAttribute(Tarifa & "")
			     else
			        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTarifa,"",enc.EncodeForHtmlAttribute(Tarifa & "")
			     end if
                 
                 EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDivisa,"",enc.EncodeForHtmlAttribute(null_s(rst("abreviatura")))


		    if si_tiene_modulo_OrCU <> 0 then
		            rstSelect.cursorlocation=3
	                rstSelect.open "select descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo='" & rst("Campo01") & "'", session("dsn_cliente")
		            if rstSelect.EOF then
		                dd=""
		            else
		                dd=rstSelect("descripcion")
		            end if

                    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDto1,"",enc.EncodeForHtmlAttribute(dd & "")
		            rstSelect.close
		            
		            rstSelect.cursorlocation=3
	                rstSelect.open "select descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo='" & rst("Campo02") & "'", session("dsn_cliente")
		            if rstSelect.EOF then
		                dd=""
		            else
		                dd=rstSelect("descripcion")
		            end if
		            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDto2,"",enc.EncodeForHtmlAttribute(dd & "")
		            rstSelect.close

		            rstSelect.cursorlocation=3
	                rstSelect.open "select descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo='" & rst("Campo03") & "'",session("dsn_cliente")
		            if rstSelect.EOF then
		                dd=""
		            else
		                dd=rstSelect("descripcion")
		            end if
		            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDto3,"",enc.EncodeForHtmlAttribute(dd & "")
		            rstSelect.close

			        IF si_tiene_modulo_OrCU<>0  then 

                        'EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCreditoSuministro,"","&nbsp;&nbsp;"
                        'DrawCeldaResponsive1 "width100","","",0,LitCreditoSuministro
                        DrawDiv "3-sub","",""
                        DrawLabel "","",LitCreditoSuministro
                        CloseDiv

                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSaldoOffline,"",enc.EncodeForHtmlAttribute(rst("saldooffline") & "")
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSaldoMax,"",enc.EncodeForHtmlAttribute(rst("saldomax") & "")
			            %> 
			                <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
                                <label><%=LitSaldoActual %></label>
                                <%
                                DrawSpan clase, "", rst("saldo"), "" 
                                if rst("look")&""="1" then%>
			                    <a id="idpermitirSuministro" class="CELDAREF" href="javascript:PermitirSuministro('<%=rst("ncliente")%>')"><%=LITPERMITIRSUMINISTRO%></a>
                                <%end if%>
                            </div>

                            <input type="hidden" name="hd_saldoMax" value="<%=rst("saldomax")%>" />
                            <input type="hidden" name="saldoMaxOld" value="<%=rst("saldomax") %>" />
                            <%

			        else
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCredito,"",enc.EncodeForHtmlAttribute(iif(isnull(rst("campo04")),"",rst("campo04")))

			        end if	 
	        end if
			    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDescuento1,"",enc.EncodeForHtmlAttribute(rst("dto") & "")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDescuento2,"",enc.EncodeForHtmlAttribute(rst("dto2") & "")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDescuento3,"",enc.EncodeForHtmlAttribute(rst("dto3") & "")

				 if si_tiene_modulo_EBESA <> 0 then
	    			 if rst("atp") =  0 then
		    		    sino = "No"
			    	 else
				        sino = "Sí"
    				 end if
                     EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDtoImpFactura,"",enc.EncodeForHtmlAttribute(sino & "")
	             end if
            if si_tiene_modulo_petroleos<>0 then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDescuentoLineal,"",enc.EncodeForHtmlAttribute(rst("dtolineal")&"")

	        else
	            %><input type="hidden" name="descuentoLineal" value="0"/> <%
	        end if
	        'FLM:20090429:añado campo para saber como se agrupan los suministros en las facturas: por cliente o por tarjeta...
	        if si_tiene_modulo_petroleos<>0 or si_tiene_modulo_OrCU<>0 then 
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitModFactSum,"",iif(rst("tipoagrupacionsum"),LitSepXTar,LitAgrXcli)
		    else%>
		    <input type="hidden" name="modFactSum" value="0"/>
		    <%end if


				 if rst("fpago")>"" then
				 	FormaPago = d_lookup("descripcion","formas_pago","codigo like '" & session("ncliente") & "%' and codigo='" + rst("fpago") + "'",session("dsn_cliente"))
				 else
				 	FormaPago=""
				 end if
                 EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFormaPago,"",FormaPago


				 if rst("tpago")>"" then
				 	TipoPago = d_lookup("descripcion","tipo_pago","codigo like '" & session("ncliente") & "%' and codigo='" + rst("tpago") + "'",session("dsn_cliente"))
				 else
				 	TipoPago = ""
				 end if

                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTipoPago,"",TipoPago


			if si_tiene_modulo_EBESA <> 0 then
				    if rst("campo11") & "" <>"" then
				 	    TipoPago = d_lookup("descripcion","tipo_pago","codigo like '" & session("ncliente") & "%' and codigo='" & rst("campo11") & "'",session("dsn_cliente"))
				    else
				 	    TipoPago = ""
				    end if
				    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTipoPago + " (" & LitNoVender & ")-1:","",TipoPago

				    if rst("campo12") & "" <>"" then
				 	    TipoPago = d_lookup("descripcion","tipo_pago","codigo like '" & session("ncliente") & "%' and codigo='" & rst("campo12") & "'",session("dsn_cliente"))
				    else
				 	    TipoPago = ""
				    end if

                    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTipoPago + " (" & LitNoVender & ")-2:","",TipoPago

				    if rst("campo13") & "" <>"" then
				 	    TipoPago = d_lookup("descripcion","tipo_pago","codigo like '" & session("ncliente") & "%' and codigo='" & rst("campo13") & "'",session("dsn_cliente"))
				    else
				 	    TipoPago = ""
				    end if

                    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTipoPago + " (" & LitNoVender & ")-3:","",TipoPago

			end if
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPrimerVen,"",rst("primer_ven")
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSegunVen,"",rst("segundo_ven")
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTercerVen,"",rst("tercer_ven")
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitMesNoPago,"",rst("mesnopago")

			''fin cag
                 EligeCeldaResponsive "text",mode,clase,"","",0,"",LitRFinanciero,"",rst("recargo")

				 if rst("re") =  0 then
				    sino = "No"
				 else
				    sino = "Sí"
				 end if

                 EligeCeldaResponsive "text",mode,clase,"","",0,"",LitREquivalencia,"",sino
		rst2.cursorlocation=3
		if col = 0 then
		    rst2.open "select ndist from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente")
        else
		    rst2.open "select ndist from distribuidores with(nolock) where ndist like '" & session("ncliente") & "%' and ndist='" & rst("ndist") & "'",session("dsn_cliente")
        end if
		if not rst2.eof then
			no_mostrar=1
		else
			no_mostrar=0
		end if
		rst2.close
		if no_mostrar=1 then
			rst2.cursorlocation=3
			if col = 0 then
			    rst2.open "select * from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente")
			else
			    rst2.open "select * from distribuidores with(nolock) where ndist like '" & session("ncliente") & "%' and ndist='" & rst("ndist") & "'",session("dsn_cliente")
			end if
			if not rst2.eof then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitVencom1,"",Null_z(rst2("vencom1"))
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPorcom1,"",Null_z(rst2("pcom1"))
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitVencom2,"",Null_z(rst2("vencom2"))
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPorcom2,"",Null_z(rst2("pcom2"))
			end if
			rst2.close
		end if

		
		

            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitRiesMaxAut,"",formatnumber(null_z(rst("riesgo1")),ndecimales,-1,0,-1) & " " & abreviatura

			if rst("riesgo2")>rst("riesgo1") then
				clase="'CELDAREDBOLD span-browser'"
			else
				clase="span-browser"
			end if
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitRiesAlc,"",formatnumber(null_z(rst("riesgo2")),ndecimales,-1,0,-1) & " " & abreviatura
            clase="span-browser"

		
        if nz_b2(gestbono)=1 then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSaldoBonoMax,"",formatnumber(null_z(rst("saldobonomax")),ndecimales,-1,0,-1) & " " & abreviatura
			    clase="span-browser"
		        if rst("SaldoBono")<0 then
			        clase="'CELDAREDBOLD span-browser'"
		        end if			    

                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSaldoBono,"",formatnumber(null_z(rst("saldobono")),ndecimales,-1,0,-1) & " " & abreviatura
                clase="span-browser"
		end if


			if si_tiene_modulo_comercial<>0 then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitComAsignadoModCom,"",d_lookup("nombre","personal","dni like '" & session("ncliente") & "%' and dni='" & rst("comercial") & "'",session("dsn_cliente"))
			else
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitComAsignado,"",d_lookup("nombre","personal","dni like '" & session("ncliente") & "%' and dni='" & rst("comercial") & "'",session("dsn_cliente"))
			end if

			if si_tiene_modulo_comercial<>0 then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitAgenteAsignado,"",d_lookup("nombre","agentes","codigo like '" & session("ncliente") & "%' and codigo='" & rst("agente") & "'",session("dsn_cliente"))
			end if

		'cag
			cf=limpiaCadena(request.querystring("cf"))
			if cf="" then
				cf=limpiaCadena(request.form("cf"))
			end if
			if cstr(cf)="1" then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitNumCopiasFacturas,"",rst("numCopiasFactura")

			end if
		'fin cag
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTipIvaCli,"",rst("iva")
			if rst("ndist")>"" then
				rst2.cursorlocation=3
				if col = 0 then
				    rst2.open "select ncliente from distribuidores with(nolock) where ndist like '" & session("ncliente") & "%' and ndist='" & rst("ndist") & "'",session("dsn_cliente")
				    if not rst2.eof then
					descrip=d_lookup("rsocial","clientes","ncliente like '" & session("ncliente") & "%' and ncliente='" & rst2("ncliente") & "'",session("dsn_cliente"))
				    else
					    descrip=""
				    end if
				    rst2.close
			    else    
			       ' rst2.open "select name from distribuidores with(nolock) where ndist like '" & session("ncliente") &"%' and ndist = '"& rst("ndist") &"'",session("dsn_cliente")
			        descrip=d_lookup("name","distribuidores","ndist like '" & session("ncliente") & "%' and ndist='" & rst("ndist") & "'",session("dsn_cliente"))
			    end if
				
			end if

            EligeCeldaResponsive "text",mode,clase,"","",0,"",LIT_DISTCOLLAB,"",descrip



		%>
	    </table>
	   <%'DGB
                DrawDiv "3-sub","",""
                    DrawLabel "","",LITCONTA
                    CloseDiv%>
                <table class="DataTable">
                        <%
                            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCContable,"",rst("ccontable")
                            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCContable_efecto,"",rst("ccontable_efecto")
                            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitIntracomunitario,"",visualizar(rst("intra"))
                if SuplidosActivados=1 then
                    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCContable_suplidos,"",rst("CCONTABLE_SUPLIDOS")
                end if
                            %>
                </table>
      </div>
</div>
		<% 'DATOS BANCARIOS MODO BROWSE %>
          <div   class="Section" id="S_BrowseDB" >
    <a href="#" rel="toggle[BrowseDB]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitDatosBancarios%>
   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
     
    </div></a>
    <div class="SectionPanel" style="display: none" id="BrowseDB">
       <table width="100% "bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5">
		      <%if si_tiene_modulo_importaciones<>0 then%>
			   	<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
	   				<a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=importaciones/bancos.asp&pag2=importaciones/bancos_bt.asp&ncliente=<%=enc.EncodeForJavascript(ncliente)%>&titulo=LISTA DE BANCOS&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" onmouseover="self.status='<%=LitIrBancos%>'; return true;" onmouseout="self.status=''; return true;"><%=LitBancos%></a>
                </div><%
			end if
			
			'FLM:20/01/2009:Añadir nuevos datos bancarios para el módulo ORCU
			if si_tiene_modulo_OrCU<>0 then
                'DrawCeldaResponsive1 "width100","","",0,LitORCUGasOtros ''Hacer un salto de linea aquí
                DrawDiv "3-sub","",""
                DrawLabel "","",LitORCUGasOtros
                CloseDiv
            end if
            %>

	        </table><table width="100% "bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5">
            <%

            clase="span-browser"

            if rst("banco") & "" <>"" then
                datosBanco=enc.EncodeForHtmlAttribute(rst("banco"))
            else
                datosBanco=rst("banco")
            end if
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitEntidad,"",datosBanco

            if rst("bancodom") & "" <>"" then
               datosBancodom=enc.EncodeForHtmlAttribute(rst("bancodom"))
            else
               datosBancodom=rst("bancodom")
            end if
            EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDombanco,"",datosBancodom

            if rst("ncuenta")&"">"" then
                'if len(rst("ncuenta"))>=24 then
                    country=mid(rst("ncuenta"),1,2)
                    iban=mid(rst("ncuenta"),3,2)
                    strBanco = Mid(rst("ncuenta"), 5, 4)
	    		    strOficina = Mid(rst("ncuenta"), 9, 4)
    			    strDC = Mid(rst("ncuenta"), 13, 2)
    			    strCuenta = Mid(rst("ncuenta"), 15, len(rst("ncuenta")) - 14)
                    ncuenta = country & " " & iban & " " & strBanco & "-" & strOficina & "-" & strDC & "-" & strCuenta
                'else
                    'strBanco = Mid(rst("ncuenta"), 1, 4)
	    		    'strOficina = Mid(rst("ncuenta"), 5, 4)
    			    'strDC = Mid(rst("ncuenta"), 9, 2)
    			    'strCuenta = Mid(rst("ncuenta"), 11, 10)
                    'ncuenta = strBanco & "-" & strOficina & "-" & strDC & "-" & strCuenta
                'end if
            end if

             EligeCeldaResponsive "text",mode,clase,"","",0,"",LitNumCuenta,"",ncuenta

             EligeCeldaResponsive "text",mode,clase,"","",0,"",LitNumTarjeta,"",rst("ntarjeta")
             EligeCeldaResponsive "text",mode,clase,"","",0,"",LitBICSWIFT,"",rst("swift_code")
             EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFCaducidad,"",rst("fcaducidad")

			 if rst("domrec") =  0 then
			    sino = LitNo
			 else
			    sino = LitSi
			 end if

             EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDomiciliacion,"",sino

			'FLM:20/01/2009:Añadir nuevos datos bancarios para el módulo ORCU
			'FLM:20/01/2009:Formato bancario.	 
		    if si_tiene_modulo_OrCU<>0 then
                    rstSelect.cursorlocation=3
			        rstSelect.open "select b.nbanco,isnull(b.entidad,'')+'-'+isnull(norma,'')+isnull('-'+case when b.tipo_gasoleo='0' then '"+LitTodosGas+"' else te.descripcion end,'') as entidad from bancos b  with(nolock) left join tipos_entidades te  with(nolock) on te.codigo=b.tipo_gasoleo and te.codigo like '" + session("ncliente") +"%' where b.nbanco like '" & session("ncliente") & "%' and nbanco='"&rst("formatobanco")&"' order by b.entidad",session("dsn_cliente")
                    if rstSelect.eof then
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFormBancario,"",""
                    else
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFormBancario,"",enc.EncodeForHtmlAttribute(rstSelect("entidad"))
                    end if
                    rstSelect.close
                    'dgb: 27/10/2009  Xenteo pago de gasoleo
                    if modulo_Xenteo<>0 then
			            if rst("pagoenpostea") =  0 then
			                sino = LitNo
			            else
			                sino = LitSi
			           end if
			           EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPoste,"",sino
                    end if

                %>
        <!--<div class="subsection"><%=LitORCUGasB%></div>
                <div class="subsectionpanel">
                    <table class="DataTable"></table>-->
                        <%
        'DrawCeldaResponsive1 "width100","","",0,LitORCUGasB
        DrawDiv "3-sub","",""
        DrawLabel "","",LitORCUGasB
        CloseDiv

                if si_tiene_modulo_TGB<>0 then

                        if rst("TGBBANCO")&"">"" then
                            'if len(rst("TGBBANCO"))>=24 then
                                country=mid(rst("TGBBANCO"),1,2)
                                iban=mid(rst("TGBBANCO"),3,2)
                                strBanco = Mid(rst("TGBBANCO"), 5, 4)
	    		                strOficina = Mid(rst("TGBBANCO"), 9, 4)
    			                strDC = Mid(rst("TGBBANCO"), 13, 2)
    			                strCuenta = Mid(rst("TGBBANCO"), 15, len(rst("TGBBANCO")) - 14)
                                ncuentaGB = country & " " & iban & " " & strBanco & "-" & strOficina & "-" & strDC & "-" & strCuenta
                            'else
                                'strBanco = Mid(rst("TGBBANCO"), 1, 4)
	    		                'strOficina = Mid(rst("TGBBANCO"), 5, 4)
    			                'strDC = Mid(rst("TGBBANCO"), 9, 2)
    			                'strCuenta = Mid(rst("TGBBANCO"), 11, 10)
                                'ncuentaGB = strBanco & "-" & strOficina & "-" & strDC & "-" & strCuenta
                            'end if
                         end if
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitNumCuenta,"",ncuentaGB
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitBICSWIFT,"",rst("swift_code2")

                        if rst("TGBBANCODOM") & "" <>"" then
                           datosTGBBANCODOM=enc.EncodeForHtmlAttribute(rst("TGBBANCODOM"))
                        else
                           datosTGBBANCODOM=rst("TGBBANCODOM")
                        end if
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDomBanco,"",datosTGBBANCODOM

                        if rst("TGBBANCOPOB") & "" <>"" then
                           datosTGBBANCOPOB=enc.EncodeForHtmlAttribute(rst("TGBBANCOPOB"))
                        else
                           datosTGBBANCOPOB=rst("TGBBANCOPOB")
                        end if
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPoblacion,"",datosTGBBANCOPOB

                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitProvincia,"",d_lookup("descripcion", "provincias", "idprovincia='"&rst("TGBBANCOPROV")&"'",DSNIlion)&""

                end if
                    

                     rstSelect.cursorlocation=3
			        rstSelect.open "select b.nbanco,isnull(b.entidad,'')+'-'+isnull(norma,'')+isnull('-'+case when b.tipo_gasoleo='0' then '"+LitTodosGas+"' else te.descripcion end,'') as entidad from bancos b  with(nolock) left join tipos_entidades te  with(nolock) on te.codigo=b.tipo_gasoleo and te.codigo like '" + session("ncliente") +"%' where b.nbanco like '" & session("ncliente") & "%' and nbanco='"&rst("formatobanco2")&"' order by b.entidad",session("dsn_cliente")
			        if rstSelect.eof then
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFormBancario,"",""
                    else
                        EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFormBancario,"",enc.EncodeForHtmlAttribute(rstSelect("entidad"))
                    end if
                    rstSelect.close
                     'dgb: 27/10/2009  Xenteo pago de gasoleo
                    if modulo_Xenteo<>0 then
			            if rst("pagoenposteb") =  0 then
			                sino = LitNo
			            else
			                sino = LitSi
			           end if

                       EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPoste,"",sino
                    end if
                    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitNumTarjeta,"",rst("ntarjeta2")
                    EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFCaducidad,"",rst("fcaducidad2")
			end if%>
	  </table>
	  <!--</div>-->
      </div>
    </div>
		 <%'DIRECCION ENVIO MODO BROWSE
		rstSelect.cursorlocation=3
		 rstSelect.Open "select domicilio,telefono,poblacion,cp,provincia,pais from domicilios with(NOLOCK) where pertenece like '" & session("ncliente") & "%' and codigo='" & rst("dir_envio") & "'", session("dsn_cliente")
		 if not rstSelect.EOF then
			'telefono2=rstSelect("telefono2")
			'fax=rstSelect("fax")
			if isnull(rst("rsocial")) then rsocialMap = "" else rsocialMap = enc.EncodeForHtmlAttribute(rst("rsocial")) end if
			if isnull(rstSelect("domicilio")) then domicilio = "" else domicilio = enc.EncodeForHtmlAttribute(rstSelect("domicilio")) end if
			if isnull(rstSelect("telefono")) then telefono = "" else telefono = rstSelect("telefono") end if
   	        if isnull(rstSelect("poblacion")) then poblacion = "" else poblacion = enc.EncodeForHtmlAttribute(rstSelect("poblacion")) end if
  	        if isnull(rstSelect("cp")) then cp = "" else cp = rstSelect("cp") end if
  	        if isnull(rstSelect("provincia")) then provincia = "" else provincia = rstSelect("provincia") end if      	  
  	        if isnull(rstSelect("pais")) then pais = "" else pais = rstSelect("pais") end if
      	        
			rstSelect.close
		 else
		 	rstSelect.Close
		 end if
		 rstSelect.cursorlocation=3
		 rstSelect.Open "select domicilio,telefono,poblacion,cp,provincia,pais from domicilios with(NOLOCK) where pertenece like '" & session("ncliente") & "%' and codigo='" & rst("invoice_address") & "'", session("dsn_cliente")
		 if not rstSelect.EOF then
			'telefono2=rstSelect("telefono2")
			'fax=rstSelect("fax")
			if isnull(rst("rsocial")) then rsocialMapF = "" else rsocialMapF = enc.EncodeForHtmlAttribute(rst("rsocial")) end if
			if isnull(rstSelect("domicilio")) then domicilioF = "" else domicilioF = enc.EncodeForHtmlAttribute(rstSelect("domicilio")) end if
			if isnull(rstSelect("telefono")) then telefonoF = "" else telefonoF = rstSelect("telefono") end if
   	        if isnull(rstSelect("poblacion")) then poblacionF = "" else poblacionF = enc.EncodeForHtmlAttribute(rstSelect("poblacion")) end if
  	        if isnull(rstSelect("cp")) then cpF = "" else cpF = rstSelect("cp") end if
  	        if isnull(rstSelect("provincia")) then provinciaF = "" else provinciaF = rstSelect("provincia") end if      	  
  	        if isnull(rstSelect("pais")) then paisF = "" else paisF = rstSelect("pais") end if
			rstSelect.close
		 else
		 	rstSelect.Close
		 end if%>
     <div class="Section" id="S_BrowseDE">
    <a href="#" rel="toggle[BrowseDE]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitDireccionEnvio%>
    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
     
    </div></a>
    <div class="SectionPanel" style="display: none" id="BrowseDE">
        <% clase="span-browser"%>
      	<table border='0' cellpadding="1" cellspacing="1" width='100%' >
		  <%
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDomicilio + LinkGoogleMap(rsocialMap,domicilio, poblacion,cp, provincia, pais,0),"",domicilio
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTel1,"",telefono
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPoblacion,"",poblacion
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCP,"",cp
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitProvincia,"",provincia
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPais,"",pais
              %>
        </table>
            <% ' DGM 19/9/11 Ocultamos la direccion de envio de factura %>	  
		  <div style="display:none">
		<table border='0' cellpadding=1 cellspacing=1 width='100%' >
		  <%  EligeCeldaResponsive "text",mode,clase,"","",0,"",LitDomicilio + LinkGoogleMap(rsocialMapF,domicilioF, poblacionF,cpF, provinciaF, paisF,0),"",domicilioF
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTel1,"",telefonoF
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPoblacion,"",poblacionF
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCP,"",cpF
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitProvincia,"",provinciaF
              EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPais,"",paisF
          %>
		 </table>
		  </div>
		  </div>
          </div>
		  <% 'OTROS DATOS MODO BROWSE%>

    <div class="Section" id="S_BrowseOD" >
    <a href="#" rel="toggle[BrowseOD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitOtrosDatos%>
    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />

    </div></a>
    <div class="SectionPanel" style="display: none" id="BrowseOD">
		<table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="5">
      		<%'jcg 02/02/2008
            clase="span-browser"
            if si_tiene_modulo_proyectos<>0 then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitProyecto,"",d_lookup("nombre","proyectos","codigo like '" & session("ncliente") & "%' and codigo='" & rst("proyecto") & "'",session("dsn_cliente"))				  
            end if
				if rst("tactividad")>"" then
				 	Actividad = d_lookup("descripcion","tipo_actividad","codigo like '" & session("ncliente") & "%' and codigo='" + rst("tactividad") + "'",session("dsn_cliente"))
				else
				 	Actividad = ""
				end if
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTActividad,"",Actividad
				EligeCeldaResponsive "text",mode,clase,"","",0,"",LitZona,"",d_lookup("descripcion","zonas","codigo like '" & session("ncliente") & "%' and codigo='" & rst("zona") & "'",session("dsn_cliente"))&""			  
                
                if rst("transportista") & "" <>"" then
                    datosTransportista=enc.EncodeForHtmlAttribute(rst("transportista"))
                else
                    datosTransportista=rst("transportista")
                end if
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTransportista,"",datosTransportista
				if rst("portes")=LitDebidos and rtrim(rst("referencia"))&"">"" then
						precioPortesFinal= PrecioArticulo (rst("referencia"),date(),1,rst("tarifa"),precioPortes,precioPortesDto)
						precioPortesFinal= " ("& formatnumber(precioPortesFinal,rst("ndecimales"),0,0,-1) &" " & rst("abreviatura") &")"
				end if
				EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPortes,"",rst("portes") & precioPortesFinal
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitHmanyana,"",rst("hmanyana")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitHTarde,"",rst("htarde")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPht,"",rst("pht")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPkm,"",rst("pkm")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPD,"",rst("pd")

				if rst("tipo_cliente")>"" then
				 	TCliente = d_lookup("descripcion","tipos_entidades","codigo like '" & session("ncliente") & "%' and codigo='" + rst("tipo_cliente") + "'",session("dsn_cliente"))
				else
				 	TCliente = ""
				end if
                if si_tiene_modulo_agrario<>0 then
			 		EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTSocio,"",TCliente
			 	else
			 		EligeCeldaResponsive "text",mode,clase,"","",0,"",LitTCliente,"",TCliente
			 	end if

			if si_tiene_modulo_ecomerce<>0 then
                  DrawDiv "1","","mostrar_verstock1"
                  DrawLabel "","width:150px;display:" & mostrar_verstock,LitMostStockTienda
					if nz_b(rst("verstock"))<>0 then
                        DrawSpan clase, "", LitSi, ""
					else
                        DrawSpan clase, "", LitNo, ""
					end if
                  CloseDiv
			end if
			if si_asesoria=true then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitCarpetaNominas,"",rst("dsn_nominaplus")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitPeriodicidadFact,"", d_lookup("descripcion","periodos","codigo like '" & session("ncliente") & "%' and codigo='"&rst("periodicidad")&"'",session("dsn_cliente"))
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSucursal,"",d_lookup("descripcion","tiendas","codigo='"&rst("tienda")&"'",session("dsn_cliente"))
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFPublicacion,"",rst("fpubnominas")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitMostrarListPortal,"",visualizar(rst("ASESORIALIST"))
			end if
			if si_tiene_modulo_agrario<>0 then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitFNacimiento,"",rst("fnacimiento")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSegSocial,"",rst("segsocial")
			elseif si_asesoria=true then
				EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSegSocial,"",rst("segsocial")
			end if
			
		    if si_tiene_modulo_CRMComunicacion <> 0 then
                  EligeCeldaResponsive "text",mode,clase,"","",0,"",LITDELIVERYADV,"",iif(rst("submit_advertising")<>0,LitSi,LitNo)
                  EligeCeldaResponsive "text",mode,clase,"","",0,"",LITCOMMEMAIL,"", iif(rst("email_communication")<>0,LitSi,LitNo)
                  EligeCeldaResponsive "text",mode,clase,"","",0,"",LITCOMMSMS,"",iif(rst("sms_communication")<>0,LitSi,LitNo)
		    end if
            if si_tiene_modulo_fidelizacion30<>0  then
                  EligeCeldaResponsive "text",mode,clase,"","",0,"",LITROL,"",iif(null_z(rst("role"))=0,LITCLIENTEMIN, LITNETWORK)
            end if
            
             %>
		  </table>
    		<%Foto "browse",rst("ncliente")%>
		  </div>
          </div>

		  <% 'CONFIG DOC MODO BROWSE%>

    <div class="Section" id="S_BrowseDD" >
    <a href="#" rel="toggle[BrowseDD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitConfDoc2%>
    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
     
    </div></a>
    <div class="SectionPanel" style="display: none" id="BrowseDD">
 	      
	     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=5>
      	    
	        <%rstSelect.cursorlocation=3
			rstSelect.Open "select * from documentos_cli where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'", session("dsn_cliente")
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitValorPre,"",iif(rstSelect("valorado_pre")<>0,LitSi,LitNo)
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSeriePre,"",trimCodEmpresa(rstSelect("serie_pre")) & " - " & d_lookup("nombre","series","nserie like '" & session("ncliente") & "%' and nserie='" & rstSelect("serie_pre") & "'",session("dsn_cliente"))&""
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitValorPed,"",iif(rstSelect("valorado_ped")<>0,LitSi,LitNo)
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSeriePed,"",trimCodEmpresa(rstSelect("serie_ped")) & " - " & d_lookup("nombre","series","nserie like '" & session("ncliente") & "%' and nserie='" & rstSelect("serie_ped") & "'",session("dsn_cliente"))&""
			
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitValorAlb,"",iif(rstSelect("valorado_alb")<>0,LitSi,LitNo)
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSerieAlb,"",trimCodEmpresa(rstSelect("serie_alb")) & " - " & d_lookup("nombre","series","nserie like '" & session("ncliente") & "%' and nserie='" & rstSelect("serie_alb") & "'",session("dsn_cliente"))&""

                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSerieFac,"",trimCodEmpresa(rstSelect("serie_fac")) & " - " & d_lookup("nombre","series","nserie like '" & session("ncliente") & "%' and nserie='" & rstSelect("serie_fac") & "'",session("dsn_cliente"))&""

			'AMF:29/10/2010:Serie de la incidencia.
			if ModuloContratado(session("ncliente"),ModPostVenta) <> 0 then
                EligeCeldaResponsive "text",mode,clase,"","",0,"",LitSerieIncidencia,"",trimCodEmpresa(rstSelect("serie_incidencia")) & " - " & d_lookup("nombre","series","nserie like '" & session("ncliente") & "%' and nserie='" & rstSelect("serie_incidencia") & "'",session("dsn_cliente"))&""
			end if
			rstSelect.close
			rst.close%>
		  </table>
		  </div>
          </div>
	    <% 'CAMPOS PERSONALIZABLES MODO BROWSE%>
          <div class="Section"  id="S_BrowseCP">
    <a href="#" rel="toggle[BrowseCP]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
    <%=LitCampPersoCli%>
    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
     
    </div></a>
    <div class="SectionPanel" style="display: none" id="BrowseCP">
        
            
	        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding='2' cellspacing='5'><%
        clase="span-browser"
        	rst.cursorlocation=3
	    	rst.open "select * from camposperso with(NOLOCK) where tabla='CLIENTES' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
		    if not rst.eof then
			    DrawFila ""
				num_campo=1
				num_campo2=1
				num_puestos=0
				num_puestos2=0
				while not rst.eof
					if num_puestos2>0 and (num_puestos2 mod 2)=0 then
						num_puestos2=0
					end if

                    if rst("titulo") & "" <>"" then
                       datosTitulo=enc.EncodeForHtmlAttribute(rst("titulo"))
                    else
                       datosTitulo=rst("titulo")
                    end if

					if rst("titulo") & "">"" and nz_b(rst("system_reg")) <> -1 then
						if ((num_puestos-1) mod 2)=0 then
							'DrawCelda "CELDA7 style='width:155px'","","",0,"&nbsp;"
						end if
						num_puestos=num_puestos+1
						num_puestos2=num_puestos2+1
						if rst("tipo")=2 then
                            EligeCeldaResponsive "text",mode,clase,"","",0,"",datosTitulo,"",iif(lista_valores(num_campo)="1","Sí","No")
						elseif rst("tipo")=3 then
						    'mmg 06/06/2008 >> Adaptacion para más de un comercial	    
							if lista_valores(num_campo) & "">"" then
								num_campo_str=cstr(num_campo)
								if len(num_campo_str)=1 then
									num_campo_str="0" & num_campo_str
								end if
								
								if (num_campo_str="01" and c01="c") or (num_campo_str="02" and c02="c") or (num_campo_str="03" and c03="c") or (num_campo_str="04" and c04="c") or (num_campo_str="05" and c05="c") or (num_campo_str="06" and c06="c") or (num_campo_str="07" and c07="c") or (num_campo_str="08" and c08="c") or (num_campo_str="09" and c09="c") or (num_campo_str="10" and c10="c") or (num_campo_str="11" and c11="c") or (num_campo_str="11" and c11="c") then
								    'valor_ListCampPerso=d_lookup("campo"+num_campo_str,"clientes","ncliente like '" & session("ncliente") & "%' and ncliente='"+ncliente+"'",session("dsn_cliente"))
								    'obtenemos el nombre del comercial asignado
								    cc="select nombre from personal p with(NOLOCK), clientes c with(NOLOCK) where c.ncliente like '" & session("ncliente") & "%' and p.DNI like '" & session("ncliente") & "%' and c.ncliente='"+ncliente+"' and c.campo"+num_campo_str+ "= p.DNI"
								    rstCom.cursorlocation=3
								    rstCom.open cc,session("dsn_cliente")
								    if not rstCom.eof then
								        valor_ListCampPerso=rstCom("nombre")
								    end if
								    rstCom.Close
								else
								    if isNumeric(lista_valores(num_campo)) then
								        'response.Write("<br/>PPPPP:"&lista_valores(num_campo))
								        valor_ListCampPerso=d_lookup("valor","campospersolista","ncampo like '" & session("ncliente") & "%' and ncampo='" & session("ncliente") & num_campo_str & "' and tabla='CLIENTES' and ndetlista=" & lista_valores(num_campo),session("dsn_cliente"))
								        'response.Write("<br/>"&valor_ListCampPerso)
								    else
								        valor_ListCampPerso=""
								    end if
								end if
							else
								valor_ListCampPerso=""
							end if
                            EligeCeldaResponsive "text",mode,clase,"","",0,"",datosTitulo,"",valor_ListCampPerso&""
						else
                            EligeCeldaResponsive "text",mode,clase,"","",0,"",datosTitulo,"",lista_valores(num_campo)
						end if
					end if
					rst.movenext
					num_campo=num_campo+1
					if not rst.eof then
						if rst("titulo") & "">"" then
							num_campo2=num_campo2+1
						end if
					end if
				wend
			num_campos=num_puestos
		else
			num_campos=0
		end if

		rst.close
	%></table>
		  
		  </div>
          </div>

    <%end if ''de comprobar si el cliente tiene registro en documentos_cli
        BarraNavegacion "Browse"
   elseif mode="edit" then
   
        'DGM 27/7/11 Comprobamos si tiene tarjetas
        rst.cursorlocation=3
        rst.Open "select count(pan) as total from tarjetas with(NOLOCK) where ncliente like '" & session("ncliente") & "%' and ncliente ='" & ncliente &"'",DsnIlion,adOpenKeyset,adLockOptimistic
        if not rst.eof then
            if clng(rst("total")) > 0 then
                tiene_tarjetas = 1
            else
                tiene_tarjetas = 0
            end if
        else
            tiene_tarjetas = 0
        end if
        rst.Close
        rst.cursorlocation=3
    rst.Open "select * from clientes with(NOLOCK),domicilios with(NOLOCK) where ncliente like '" & session("ncliente") & "%' and ncliente='" & ncliente & "' And codigo=dir_principal", session("dsn_cliente")
    if not rst.eof then
		    ''ricardo 3-9-2004 para ver foto en clientes
	 	    if rst("tipo_foto")>"" and not isnull(rst("tipo_foto")) then
		    	    mostrar_foto=1
	 	    else
		    	    mostrar_foto=0
	 	    end if
    else
	    mostrar_foto=0
    end if

	if viene<>"facturas_cli_E" then
        if mode <> "browse" then %><div class="headers-wrapper"></div><% end if
	  BarraOpciones "edit" , ncliente
	end if



    UsuarioEnPersonal=d_lookup("DNI","personal","DNI LIKE '" & session("ncliente") & "%' AND LOGIN='" & session("usuario") & "'",session("dsn_cliente"))%>
    <input type="hidden" name="UsuarioEnPersonal" value="<%=enc.EncodeForHtmlAttribute(UsuarioEnPersonal)%>"/>
	<input type="hidden" name="hncliente" value="<%=enc.EncodeForHtmlAttribute(rst("ncliente"))%>"/>
    <input type="hidden" name="hcif" value="<%=enc.EncodeForHtmlAttribute(rst("cif"))%>"/>
	<%rstAux.cursorlocation=3
	
	rstAux.open "select ndist from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente")
	if not rstAux.eof then
		ndist_aux=rstAux("ndist")
	else
		ndist_aux=""
	end if
	rstAux.close
	  %><input type="hidden" name="ndist" value="<%=enc.EncodeForHtmlAttribute(ndist_aux)%>"/>

   <%'DATOS GENERALES MODO EDIT

	' Inicio Borde Span
	%>
         <table width="100%">
    <div id="CollapseSection">
        <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['EditDG', 'EditDC', 'EditDB','EditDE','EditOD','EditDD','EditCP']);hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll %>" alt="" title="" <%=ParamImgCollapse %> /></a> 
        <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['EditDG', 'EditDC', 'EditDB','EditDE','EditOD','EditDD','EditCP']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll %>" alt="" title="" <%=ParamImgCollapse %> /></a>
    </div>
            <table></table>

    <div class="Section" id="S_EditDG" >
        <a href="#" rel="toggle[EditDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader displayed">
                <%=LitDatosGenerales%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
            </div>
        </a>   
        <!--<table width="100%" border='0' cellpadding="2" cellspacing="2">
        </table>-->
        <div class="SectionPanel"  id="EditDG">

     <%
        DrawDiv "1","",""
        DrawLabel "txtMandatory","",LitRSocial
        DrawInput "", "", "rsocial", enc.EncodeForHtmlAttribute(rst("rsocial")) , "maxlength='50' size='35'"
        CloseDiv
        
        DrawDiv "1","",""
        DrawLabel "","",LitNComercial
		if rst("ncomercial") & "">"" then
            DrawInput "", "", "ncomercial", enc.EncodeForHtmlAttribute(rst("ncomercial")) , "maxlength='50' size='35'"
		else
            DrawInput "", "", "ncomercial", "", "maxlength='50' size='35'"
		end if
        CloseDiv

		if si_asesoria=true then
			 rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from FORMAS_JURIDICAS with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitFormaJuridica,"fjuridica",rstSelect,rst("fjuridica"),"codigo","descripcion","",""
			 rstSelect.close

            DrawDiv "1","",""
            DrawLabel "","",LitTitular
            DrawInput "", "", "titular", rst("titular") , "maxlength='50' size='35'"
            CloseDiv

		end if
            DrawInputCeldaLabel "","txtMandatory",20,LitCIF,"cif",rst("cif")

            if rst("contacto") & "" <>"" then
                datosContacto=enc.EncodeForHtmlAttribute(rst("contacto"))
            else
                datosContacto=rst("contacto")
            end if
            EligeCelda "input",mode,"","","",0,LitContacto,"contacto",35,datosContacto

            if rst("domicilio") & "" <>"" then
                datosDomicilio=enc.EncodeForHtmlAttribute(rst("domicilio"))
            else
                datosDomicilio=rst("domicilio")
            end if
            if rst("poblacion") & "" <>"" then
                datosPoblacion=enc.EncodeForHtmlAttribute(rst("poblacion"))
            else
                datosPoblacion=rst("poblacion")
            end if
            DrawInputCeldaLabel "","txtMandatory",35,LitDomicilio + " " + LinkGoogleMap(enc.EncodeForHtmlAttribute(rst("rsocial")),datosDomicilio, datosPoblacion,rst("cp"), rst("provincia"), rst("pais"),0),"domicilio",datosDomicilio
            DrawDiv "1","",""
            CloseDiv
            'EligeCelda "input",mode,"","","",0,LitTitular,"rsocial",35,rst("titular")

			%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                %><label><%=LitPoblacion%></label><%
                %><input class="width50" type="text" size="25" name="poblacion" value="<%=datosPoblacion%>" onchange="borrarCodigos('1')"/><%
                %><a class='CELDAREFB'  href="#SELECCIONAR_POBLACION2" onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes&titulo=<%=LITSSVERPOBLACIONES %>');" onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><%
                    %><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/><%
                %></a>
			<input type="hidden" name="codPoblacion" value="<%=enc.EncodeForHtmlAttribute(rst("codpoblacion")& "") %>" /> 
			<input type="hidden" name="codProvincia" value="<%=enc.EncodeForHtmlAttribute(rst("codprovincia")& "") %>"/>
			<input type="hidden" name="codPais" value="<%=enc.EncodeForHtmlAttribute(rst("codpais")& "") %>"/>
			</div><%
             EligeCelda "input",mode,"","","",0,LitCP,"cp",20,rst("cp")

             rstSelect.cursorlocation=3
			 rstSelect.open "select id, nombre from PAISES with(nolock) order by nombre",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitPais,"paisDDL",rstSelect,rst("codpais")&"","id","nombre","onchange","TraerPais()","pais",30,rst("pais")
             rstSelect.close

             rstSelect.cursorlocation=3
			 rstSelect.open "select idprovincia, descripcion, idpais from PROVINCIAS with(nolock) order by descripcion",DSNIlion

            nombreprovincia = ""
            %>
            <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
                <label><%=LitProvincia%></label>           
                <select class="width30" style="" name="provinciaDDL" onchange="TraerProvincia()"> 
                <% 
                    
                while not rstSelect.EOF
                    if rst("codpais")&"" <> rstSelect("idpais")&"" then
                        %>
                        <option value="<%=enc.EncodeForHtmlAttribute(rstSelect("idprovincia"))%>" data-country="<%=enc.EncodeForHtmlAttribute(rstSelect("idpais"))%>" style="display:none;" ><%=enc.EncodeForHtmlAttribute(rstSelect("descripcion"))%></option>
                        <% 
                    elseif rst("codprovincia")&"" <> rstSelect("idprovincia")&"" then 
                        %>  
                        <option value="<%=enc.EncodeForHtmlAttribute(rstSelect("idprovincia"))%>" data-country="<%=enc.EncodeForHtmlAttribute(rstSelect("idpais"))%>"><%=enc.EncodeForHtmlAttribute(rstSelect("descripcion"))%></option>
                        <% 
                    elseif rst("codprovincia")&"" = rstSelect("idprovincia")&"" then
                        nombreprovincia = rstSelect("descripcion")
                        %>
                        <option value="<%=enc.EncodeForHtmlAttribute(rstSelect("idprovincia"))%>" data-country="<%=enc.EncodeForHtmlAttribute(rstSelect("idpais"))%>" selected="selected" ><%=enc.EncodeForHtmlAttribute(rstSelect("descripcion"))%></option>
                        <% 
                    end if
                    
                    rstSelect.MoveNext
                wend

                if rst("codprovincia")&"" = "" then 
                    %><option value="" selected="selected"></option><% 
                    strselProv = "select provincia from domicilios with(NOLOCK) where pertenece like ?+'%' and codigo=? "
                    nombreprovincia=DLookupP2(strselProv,session("ncliente")&"",adVarchar,5,rst("dir_principal")&"",adVarchar,10,session("dsn_cliente")&"")
                else
                    %><option value="" ></option><% 
                end if
                %>
                </select>
                <input class="width30" type='text' name="provincia" value="<%=nombreprovincia & ""%>" size="25" />
            </div>
            <%			 
             rstSelect.close             

             EligeCelda "input",mode,"","","",0,LitTel1,"telefono",20,enc.EncodeForHtmlAttribute(rst("telefono")&"")
             EligeCelda "input",mode,"","","",0,LitTel2,"telefono2",20,enc.EncodeForHtmlAttribute(rst("telefono2")&"")
             EligeCelda "input",mode,"","","",0,LitFax,"fax",20,enc.EncodeForHtmlAttribute(rst("fax")&"")
             EligeCelda "input",mode,"","","",0,LitFechaAlta,"falta",10,enc.EncodeForHtmlAttribute(rst("falta")&"")
             DrawCalendar "falta"
             EligeCelda "input",mode,"","","",0,LitFechaBaja,"fbaja",10,enc.EncodeForHtmlAttribute(rst("fbaja")&"")
             DrawCalendar "fbaja"
             EligeCelda "input",mode,"","","",0,LitEMail,"email",35,enc.EncodeForHtmlAttribute(rst("email")&"")

             if rst("web") & "" <>"" then
                 datosWeb=enc.EncodeForHtmlAttribute(rst("web")&"")
             else
                 datosWeb=rst("web")
             end if
             EligeCelda "input",mode,"","","",0,LitWEB,"web",35,enc.EncodeForHtmlAttribute(datosWeb&"")

             if rst("observaciones") & "" <>"" then
                 datosObservaciones=enc.EncodeForHtmlAttribute(rst("observaciones")&"")
             else
                 datosObservaciones=rst("observaciones")
             end if
             EligeCelda "text",mode,"","","",60,LitObservaciones,"observaciones",2,enc.EncodeForHtmlAttribute(datosObservaciones&"")

             if rst("aviso") & "" <>"" then
                 datosAviso=enc.EncodeForHtmlAttribute(rst("aviso")&"")
             else
                 datosAviso=rst("aviso")
             end if
             EligeCelda "text",mode,"","","",60,LitAviso,"aviso",2,enc.EncodeForHtmlAttribute(datosAviso&"")
        %>
        </div>
    </div>


 <% 'DATOS COMERCIALES MODO EDIT' %>
        <div  class="Section" id="S_EditDC">
   <a href="#" rel="toggle[EditDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
    <%=LitDatosComerciales%>
   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
             </div>  </a>
    <div class="SectionPanel" id="EditDC" style="display: none">
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
        <%
             rstSelect.cursorlocation=3
		     if si_tiene_modulo_OrCU <> 0 then
		         rstSelect.open "select codigo, descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is null order by descripcion", 	session("dsn_cliente")
			     DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTarifaPre,"tarifa",rstSelect,rst("tarifa"),"codigo","descripcion","",""
             else
			    rstSelect.open "select codigo, descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", 	session("dsn_cliente")
			    DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTarifa,"tarifa",rstSelect,rst("tarifa"),"codigo","descripcion","",""
             end if
			 rstSelect.close

			 rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, abreviatura from divisas with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by abreviatura",session("dsn_cliente")
             DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitDivisa,"divisa",rstSelect,rst("divisa"),"codigo","abreviatura","",""
			 rstSelect.close

		if si_tiene_modulo_OrCU <> 0 then
		     rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is not null order by descripcion", 	session("dsn_cliente")
			 DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitDto1,"dtoCli1",rstSelect,rst("Campo01"),"codigo","descripcion","",""
			 rstSelect.close
			 
			 rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is not null order by descripcion", 	session("dsn_cliente")
             DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitDto2,"dtoCli2",rstSelect,rst("Campo02"),"codigo","descripcion","",""
			 rstSelect.close

		     rstSelect.cursorlocation=3
			 rstSelect.open "select codigo, descripcion from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and TarifaCliente is not null order by descripcion", 	session("dsn_cliente")
             DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitDto3,"dtoCli3",rstSelect,rst("Campo03"),"codigo","descripcion","",""
			 rstSelect.close

			 if si_tiene_modulo_OrCU<>0 then 
			    'DrawCelda "", "", "", "", LitCreditoSuministro
                'DrawCeldaResponsive1 "width100","","",0,LitCreditoSuministro
                DrawDiv "3-sub","",""
                DrawLabel "","",LitCreditoSuministro
                CloseDiv
            
			    DrawInputCelda "","","",8,0,LitSaldoOffline,"saldooffline",rst("saldooffline")

			    saldoMaxactivo=""
			    saldoMaxDisabled=""   
			    
			    if  rst("saldomax")&"">"" then 
			         if replace(rst("saldomax"),",",".")="9999999.99" then
			            saldoMaxDisabled=""     
	                    saldoMaxactivo="checked"
	                 end if
	            end if
                DrawDiv "1","",""
                DrawLabel "","",LitSaldoMax
                DrawInput "", "", "saldomax", rst("saldomax"), "onchange='javascript:saldoMaxChanged()' size='10'"
                CloseDiv

			    DrawDiv "1","",""
                DrawLabel "","",LitSaldoSinLimite
                %><input type="checkbox" name="cbSaldoSinLimite" value="1" onclick="javascript: SaldoSinLimiteChanged()" <%=saldoMaxactivo%>/><%
                CloseDiv

                %>
                <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12" >
                    <label><%=LitSaldoActual %></label>
                    <input type="text" name="saldoact" readonly="readonly" value="<%=enc.EncodeForHtmlAttribute(rst("saldo")&"")%>" size="10" maxlength="10" />
                        <%if rst("look")&""="1" then%>
			            <a id="idpermitirSuministro" class="CELDAREF" href="javascript:PermitirSuministro('<%=EncodeForJavascript(rst("ncliente"))%>')"><%=LITPERMITIRSUMINISTRO%></a>
                        <%end if %>
                </div>
                <input type="hidden" name="hd_saldoMax" value="<%=rst("saldomax") %>" />
                <input type="hidden" name="saldoMaxOld" value="<%=rst("saldomax")%>" />
			    <input type="hidden" name="hd_saldoEnvidado" value="0" />
			    <input type="hidden" name="hd_saldoOffline" value="<%=rst("saldooffline") %>" />
			    <script language="javascript" type="text/javascript">
                    //hacemos esta llamada para que en el caso de estar chequeado el saldomaximo se ponga el input de saldomax a disabled
                    //esto no podemos hacerlo en el lado del servidor porque si ie recibe el campo como disabled al principio
                    // ya no lo pone a enabled aunque lo cambiemos con javascript. Por eso lo que hacemos es que se ponga a disabled con javascript 
                    //ya que así si que nos permite luego volver a activarlo.
                    SaldoSinLimiteChanged();
                    document.clientes.saldomax.value = '<%=rst("saldomax") %>';
                    document.clientes.saldoact.value = '<%=rst("saldo") %>';
                        </script>
			    <%
			 else

                if isnull(rst("campo04")) then  credito="" else credito=rst("campo04") end if
                DrawInputCelda "","","",20,0,LitCredito,"credito",credito
			end if
  	 
	    end if
		DrawInputCelda "","","",4,0,LitDescuento1,"descuento",rst("dto")
		DrawInputCelda "","","",4,0,LitDescuento2,"descuento2",rst("dto2")
		DrawInputCelda "","","",4,0,LitDescuento3,"descuento3",rst("dto3")

		if si_tiene_modulo_EBESA <> 0 then
	        DrawCheckCelda "","","",0,LitDtoImpFactura,"dtoimpfact",rst("atp")
	    end if

		if si_tiene_modulo_petroleos<>0 then
	        'EligeCelda "input",mode,"","","",0,LitDescuentoLineal,"descuentolineal",4,rst("dtoLineal")
	        DrawInputCelda "","","",4,0,LitDescuentoLineal,"descuentoLineal",rst("dtoLineal")
	    else%>
	        <input type="hidden" name="descuentoLineal" value="0"/>
	    <%end if
	    'FLM:20090429:añado campo para saber como se agrupan los suministros en las facturas: por cliente o por tarjeta...
	    if si_tiene_modulo_petroleos<>0 or si_tiene_modulo_OrCU<>0 then 
			%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
                <label><%=LitModFactSum %></label>
                <select name="modFactSum" class="width60" >
			        <option value="0" <%=iif(rst("tipoagrupacionsum"),"","selected") %>><%=LitAgrXcli%></option>
			        <option value="1" <%=iif(rst("tipoagrupacionsum"),"selected","") %>><%=LitSepXTar%></option>
			    </select>
			</div>
			<%
		else%>
		<input type="hidden" name="modFactSum" value="0"/>
		<%end if
		rstSelect.cursorlocation=3
		rstSelect.open "select codigo, descripcion from formas_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
        DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitFormaPago,"fpago",rstSelect,rst("fpago"),"codigo","descripcion","",""
		rstSelect.close

		rstSelect.cursorlocation=3
		'' MPC 20/02/2008 SE ha modificado para que a EBESA solo salgan aquellos tipos de pago cuyo campo copiasticket sea 2
		strselect = "select codigo, descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%'  order by descripcion"
        rstSelect.open strselect,session("dsn_cliente")
        DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago,"tpago",rstSelect,rst("tpago"),"codigo","descripcion","",""
		rstSelect.close

		if si_tiene_modulo_EBESA <> 0 then
			rstSelect.cursorlocation=3
	        strselect = "select codigo, descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' and copiasticket=2 order by descripcion"
			rstSelect.open strselect,session("dsn_cliente")
            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago + " (" & LitNoVender & ")-1","tpagonp1",rstSelect,rst("campo11"),"codigo","descripcion","",""

			rstSelect.movefirst
            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago + " (" & LitNoVender & ")-2","tpagonp2",rstSelect,rst("campo12"),"codigo","descripcion","",""

		    rstSelect.movefirst
            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTipoPago + " (" & LitNoVender & ")-3","tpagonp3",rstSelect,rst("campo13"),"codigo","descripcion","",""

			rstSelect.close
		end if
		''cag dias pago
        DrawInputCeldaActionDiv "'' maxlength='2'","","","3",0,LitPrimerVen,"e_primer_ven",rst("primer_ven"), "onchange", "comprobar()",false
        DrawInputCeldaActionDiv "'' maxlength='2'","","","3",0,LitSegunVen,"e_segundo_ven",rst("segundo_ven"), "onchange", "comprobar()",false
        DrawInputCeldaActionDiv "'' maxlength='2'","","","3",0,LitTercerVen,"e_tercer_ven",rst("tercer_ven"), "onchange", "comprobar()",false
        %>
			  <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
                  <label><%=LitMesNoPago %></label>
                  <select class="width60" name="mesNoPago">
					<option value="0" <%=iif(rst("mesnopago")="","selected","")%> >0</option>
					<%opcionMenu=1
					while opcionMenu<=12 %>
					 	<option value="<%=enc.EncodeForHtmlAttribute(opcionMenu&"")%>" <%=iif(opcionMenu=rst("mesnopago"),"selected","")%>> <%=enc.EncodeForHtmlAttribute(opcionMenu&"")%> </option>
					<%opcionMenu=opcionMenu+1
					wend %>
				 </select>
			  </div> <%
			 DrawInputCelda "","","",4,0,LitRFinanciero,"recargo",enc.EncodeForHtmlAttribute(rst("recargo")&"")
			 DrawCheckCelda "","","",0,LitREquivalencia,"re",rst("re")

	    rst2.cursorlocation=3
	    rst2.open "select name from distribuidores with(nolock) where ndist like '" & session("ncliente") & "%' and ndist='" & rst("ndist")&"' ", session("dsn_cliente")
	    if not rst2.eof then
	        if rst2("name")&"" > "" then
	            col = 1
	        else
	            col = 0
	        end if
	    end if
	    rst2.close
	    if col = 0 then
	        rst2.open "select ndist from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente")
	        if not rst2.eof then
		        no_mostrar=1
	        else
		        no_mostrar=0
	        end if
	        rst2.close
	    else
	        no_mostrar=1
	    end if
	
	    if no_mostrar=1 then
		    rst2.cursorlocation=3
		    if col = 0 then
		        rst2.open "select * from distribuidores with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente")
            else
		        rst2.open "select * from distribuidores with(nolock) where ndist like '" & session("ncliente") & "%' and ndist='" & rst("ndist") & "'",session("dsn_cliente")
            end if
		    if not rst2.eof then
            
                DrawInputCeldaActionDiv "","","","10",0,LitVencom1,"vencom1",rst2("vencom1"), "onchange", "actualizavalores()",false
                DrawInputCeldaActionDiv "","","","4",0,LitPorcom1,"pcom1",rst2("pcom1"), "onchange", "actualizavalores()",false
                DrawInputCeldaActionDiv "","","","10",0,LitVencom2,"vencom2",rst2("vencom2"), "onchange", "actualizavalores()",false
                DrawInputCeldaActionDiv "","","","4",0,LitPorcom2,"pcom2",rst2("pcom2"), "onchange", "actualizavalores()",false

		    end if
		    rst2.close
	    end if
	''MPC 12/02/2009 Si el cliente tiene el módulo de EBESA no se ejecuta ninguna acción en caso de cambiar el riesgo máximo autorizado
	    if si_tiene_modulo_EBESA <>0 then
		    DrawInputCelda "","","",25,0,LitRiesMaxAut,"rgomaxaut",null_z(rst("riesgo1"))
	    else
            DrawInputCeldaActionDiv "","","","25",0,LitRiesMaxAut,"rgomaxaut",null_z(rst("riesgo1")), "onchange", "ComprobarCantRiesgo();",false
	    end if
	''FIN MPC
    %>
	    <input type="hidden" name="rgomaxaut_ant" value="<%=null_z(rst("riesgo1"))%>"/>
	    <input type="hidden" name="rcalc" value="0"/><%

	    if rst("riesgo2")>rst("riesgo1") then
		    clase="'CELDAREDBOLD span-browser'"
	    else
		    clase="span-browser"
	    end if
        DrawCeldaResponsive  clase & "'' name='riesgo2'","","",25,LitRiesAlc,formatnumber(null_z(rst("riesgo2")),ndecimales,-1,0,-1) & " " & abreviatura

	    if nz_b2(gestbono)=1 then
            DrawCeldaResponsive clase & "'' name='SaldobonoMax'","","",25,LitSaldoBonoMax,formatnumber(null_z(rst("SaldobonoMax")),ndecimales,-1,0,-1) & " " & abreviatura
		    clase="span-browser"
		    if rst("SaldoBono")<0 then
			    clase="'CELDAREDBOLD span-browser'"
		    end if
            DrawCeldaResponsive clase & "'' name='SaldoBono'","","",25,LitSaldoBono,formatnumber(null_z(rst("SaldoBono")),ndecimales,-1,0,-1) & " " & abreviatura
	    end if
	
    DrawDiv "1","",""
    if si_tiene_modulo_comercial<>0 then
		DrawLabel "","",LitComAsignadoModCom
	else
		DrawLabel "","",LitComAsignado
	end if
	defecto=rst("comercial")
	rstAux.cursorlocation=3
	rstAux.open "select dni, nombre from personal with(NOLOCK),comerciales with(NOLOCK) where personal.dni like '" & session("ncliente") & "%' and comerciales.comercial like '" & session("ncliente") & "%' and comerciales.fbaja is null and dni like '" & session("ncliente") & "%' and dni=comercial order by nombre",session("dsn_cliente")
	'***RGU 13/6/2006
	if (nmc<>"1" or isnull(defecto) or defecto="" ) then
        DrawSelect "", "width:200px", "comasignado", rstAux, defecto, "dni", "nombre", "", ""
	else
        DrawSpan "", "", d_lookup("nombre","personal","dni='" & defecto & "'",session("dsn_cliente")), ""
        %><input type="hidden" name="comasignado" value="<%=enc.EncodeForHtmlAttribute(defecto&"")%>" /><%
    end if
	rstAux.close
    CloseDiv

	'***RGU
	if si_tiene_modulo_comercial<>0 then
		defecto=rst("agente")
		rstAux.cursorlocation=3
		rstAux.open "select codigo, nombre from agentes with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre",session("dsn_cliente")
        DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitAgenteAsignado,"agenteasignado",rstAux,defecto,"codigo","nombre","",""
		rstAux.close
	end if

	'cag
	cf=limpiaCadena(request.querystring("cf"))
	if cf="" then
		cf=limpiaCadena(request.form("cf"))
	end if
	if cstr(cf)="1" then
        DrawInputCelda "","","",5,0,LitNumCopiasFacturas,"ncopiasFacturas",rst("numCopiasFactura")
	end if
	'fin cag
	defecto=iif(rst("iva")>"",rst("iva"),"")
	rstSelect.cursorlocation=3
	rstSelect.open "select tipo_iva as codigo,tipo_iva as descripcion from tipos_iva with(NOLOCK) order by tipo_iva",session("dsn_cliente")
	chd=""
	dsb=""
    DrawDiv "1","",""
    DrawLabel "","",LitTipIvaCli
    DrawSelect "'' "&dsb, "width:50px", "iva",rstSelect,defecto,"codigo","descripcion","",""
    CloseDiv
	rstSelect.close
	%>
	<div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
        <label><%=LIT_DISTCOLLAB %></label>
		<input type="hidden" name="distribuidor" value="<%=enc.EncodeForHtmlAttribute(TmpDistribuidor&"")%>"/>
		<iframe id="frDistribuidores" name="fr_Distribuidores" src='distribuidores_clientes.asp?distribuidor=<%=rst("ndist")%>' class="width60 iframe-menu" frameborder="no" scrolling="no" noresize="noresize"></iframe>
	</div>
</table>
   <%'DGB  
                DrawDiv "3-sub","",""
                    DrawLabel "","",LITCONTA
                CloseDiv%>
             <table class="DataTable">
          
             <%
		    DrawInputCelda "","","",25,0,LitCContable,"ccontable",enc.EncodeForHtmlAttribute(rst("ccontable")&"")
		    DrawInputCelda "","","",25,0,LitCContable_efecto,"ccontable_efecto",enc.EncodeForHtmlAttribute(rst("ccontable_efecto")&"")
            
            DrawDiv "1","",""
            DrawLabel "","",LitIntracomunitario
            %><input class='' type='checkbox' name='intra' <%=iif(nz_b2(rst("INTRA"))=1,"CHECKED", "")%> onclick="javascript: iva0()" /><%
            CloseDiv

            if SuplidosActivados=1 then
	            DrawInputCelda "","","",25,0,LitCContable_suplidos,"CCONTABLE_SUPLIDOS",enc.EncodeForHtmlAttribute(rst("CCONTABLE_SUPLIDOS")&"")
            end if

        if nz_b2(rst("INTRA"))=1 then
                %><script>iva0();</script><%
        end if
                %>
            </table>
     </div>
 </div>

  	<%'DATOS BANCARIOS MODO EDIT %>
      <div  class="Section" id="S_EditDB">
   <a href="#" rel="toggle[EditDB]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
    <%=LitDatosBancarios%>
   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
             </div>  </a>
    <div class="SectionPanel" id="EditDB" style="display: none">
     
     <table  width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
        <%
        	if si_tiene_modulo_importaciones<>0 then %>
			   	<div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"  style='width:100px'>
	   				<a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=importaciones/bancos.asp&pag2=importaciones/bancos_bt.asp&ncliente=<%=enc.EncodeForJavascript(ncliente)%>&titulo=LISTA DE BANCOS&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" onmouseover="self.status='<%=LitIrBancos%>'; return true;" onmouseout="self.status=''; return true;"><%=LitBancos%></a>
	   			</div>
	   		
            <%end if
            %></table><!--<table width="100% "bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5">--><%
	   		if si_tiene_modulo_OrCU<>0 then
                'DrawCeldaResponsive1 "width100","","",0,LitORCUGasOtros
                DrawDiv "3-sub","",""
                DrawLabel "","",LitORCUGasOtros
                CloseDiv
            end if

             if rst("banco") & "" <>"" then
                datosBanco=enc.EncodeForHtmlAttribute(rst("banco")&"")
             else
                datosBanco=rst("banco")
             end if
			 DrawInputCelda "width:200px","","",35,0,LitEntidad,"Entidad",datosBanco

             if rst("bancodom") & "" <>"" then
                datosBancodom=enc.EncodeForHtmlAttribute(rst("bancodom")&"")
             else
                datosBancodom=rst("bancodom")
             end if
			 DrawInputCelda "width:200px' maxlength='50","","",35,0,LitDomBanco,"DomBanco",datosBancodom

             if rst("ncuenta")&"">"" then
                'if len(rst("ncuenta"))>=24 then
                    country=mid(rst("ncuenta"),1,2)
                    iban=mid(rst("ncuenta"),3,2)
                    'ncuenta=mid(rst("ncuenta"),5,len(rst("ncuenta")))
                    strBanco = Mid(rst("ncuenta"), 5, 4)
			        strOficina = Mid(rst("ncuenta"), 9, 4)
			        strDC = Mid(rst("ncuenta"), 13, 2)
			        strCuenta = Mid(rst("ncuenta"), 15, len(rst("ncuenta")) - 14)
                'else
                    'strBanco = Mid(rst("ncuenta"), 1, 4)
			        'strOficina = Mid(rst("ncuenta"), 5, 4)
			        'strDC = Mid(rst("ncuenta"), 9, 2)
			        'strCuenta = Mid(rst("ncuenta"), 11, 10)
                'end if
            end if
	        %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
		        %><label><%=LitNumCuenta%></label><%
		        %><div class="inlineTable width100"><%
                %><div class="width10 tableCell"><input class='width:200px' type="text" name="country" value='<%=enc.EncodeForHtmlAttribute(country&"")%>' maxlength="2"  onkeyup="if (this.value.length==2) document.clientes.iban.focus()"  onblur="this.value=this.value.toUpperCase();"/></div><%
                %><div class="width10 tableCell"><input class='width:200px' type="text" name="iban" value='<%=enc.EncodeForHtmlAttribute(iban&"")%>' maxlength="2" onkeyup="if (this.value.length==2) document.clientes.NEntidad.focus()"/></div><%
			 	%><div class="width20 tableCell"><input class='width:200px' type="text" name="NEntidad" value='<%=enc.EncodeForHtmlAttribute(strBanco&"")%>' maxlength="4" onkeyup="if (this.value.length==4) document.clientes.Oficina.focus()"/></div><%
			 	%><div class="width20 tableCell"><input class='width:200px' type="text" name="Oficina" value='<%=enc.EncodeForHtmlAttribute(strOficina&"")%>' maxlength="4" onkeyup="if (this.value.length==4) document.clientes.DC.focus()"/></div><%
			 	%><div class="width10 tableCell"><input class='width:200px' type="text" name="DC" value='<%=enc.EncodeForHtmlAttribute(strDC&"")%>' maxlength="2"  onkeyup="if (this.value.length==2) document.clientes.Cuenta.focus()"/></div><%
			 	%><div class="width40 tableCell"><input class='width:200px' type="text" name="Cuenta" value='<%=enc.EncodeForHtmlAttribute(strCuenta&"")%>'/></div><%
		        %></div><%
	         %></div><%
             DrawInputCelda "width:200px' maxlength='16","","",35,0,LitNumTarjeta,"NumTarjeta",rst("ntarjeta")
			 DrawInputCelda "width:200px' maxlength='11","","",11,0,LitBICSWIFT,"bic",rst("swift_code")
			 DrawInputCelda "width:200px' maxlength='4","","",7,0,LitFCaducidad,"fcaducidad",rst("fcaducidad")
             DrawCheckCelda "","","",0,LitDomiciliacion,"Domiciliacion",rst("domrec")

		    'FLM:20/01/2009: Añadir campos datos bancarios módulo ORCU.
             
		if si_tiene_modulo_OrCU <> 0 or si_tiene_modulo_TGB <>0 then 
                rstSelect.cursorlocation=3
			    rstSelect.open "select b.nbanco,isnull(b.entidad,'')+'-'+isnull(norma,'')+isnull('-'+case when b.tipo_gasoleo='0' then '"+LitTodosGas+"' else te.descripcion end,'') as entidad from bancos b  with(nolock) left join tipos_entidades te  with(nolock) on te.codigo=b.tipo_gasoleo and te.codigo like '" & session("ncliente") &"%' where b.nbanco like '" & session("ncliente") & "%'"&" order by b.entidad",session("dsn_cliente")
                DrawSelectCeldaResponsive1 "width:200px",400,"",0,LitFormBancario,"formatoBanco",rstSelect,rst("formatoBanco"),"nbanco","entidad","",""
			    rstSelect.close
			     'dgb: 27/10/2009  Xenteo pago de gasoleo
			     if modulo_Xenteo<>0 then
                    DrawCheckCelda "","","",0,LitPoste,"pagoA",rst("pagoenpostea")
			     end if
            %><!--</table>-->
                <!--<div class="subsection"><%=LitORCUGasB%></div>
                    <div class="subsectionpanel">
                        <table class="DataTable"></table>-->
            <%
            'DrawCeldaResponsive1 "width100","","",0,LitORCUGasB
            DrawDiv "3-sub","",""
            DrawLabel "","",LitORCUGasB
            CloseDiv

            if si_tiene_modulo_TGB<>0 then
                    if rst("TGBBANCO")&"">"" then
                        'if len(rst("TGBBANCO"))>=24 then
                            ibanGB=mid(rst("TGBBANCO"),3,2)
                            paisGB=left(rst("TGBBANCO"),2)
                            strBancoGB = Mid(rst("TGBBANCO"), 5, 4)
			                strOficinaGB = Mid(rst("TGBBANCO"), 9, 4)
			                strDCGB = Mid(rst("TGBBANCO"), 13, 2)
			                strCuentaGB = Mid(rst("TGBBANCO"), 15, len(rst("TGBBANCO")) - 14)
                        'else
                            'strBancoGB = Mid(rst("TGBBANCO"), 1, 4)
			                'strOficinaGB = Mid(rst("TGBBANCO"), 5, 4)
			                'strDCGB = Mid(rst("TGBBANCO"), 9, 2)
			                'strCuentaGB = Mid(rst("TGBBANCO"), 11, 10)
                        'end if
                    end if
			        %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
		                %><label><%=LitNumCuenta%></label><%
		                %><div class="inlineTable width100"><%
                        %><div class="width10 tableCell"><input class='width:200px' type="text" name="countryGB" value='<%=enc.EncodeForHtmlAttribute(paisGB&"")%>' maxlength="2" onkeyup="if (this.value.length==2) document.clientes.ibanGB.focus()" onblur="this.value=this.value.toUpperCase();"/></div><%
                        %><div class="width10 tableCell"><input class='width:200px' type="text" name="ibanGB" value='<%=enc.EncodeForHtmlAttribute(ibanGB&"")%>' maxlength="2"" onkeyup="if (this.value.length==2) document.clientes.NEntidadGB.focus()"/></div><%
			 	        %><div class="width20 tableCell"><input class='width:200px' type="text" name="NEntidadGB" value='<%=enc.EncodeForHtmlAttribute(strBancoGB&"")%>' maxlength="4"  onkeyup="if (this.value.length==4) document.clientes.OficinaGB.focus()"/></div><%
			 	        %><div class="width20 tableCell"><input class='width:200px' type="text" name="OficinaGB" value='<%=enc.EncodeForHtmlAttribute(strOficinaGB&"")%>' maxlength="4" onkeyup="if (this.value.length==4) document.clientes.DCGB.focus()"/></div><%
			 	        %><div class="width10 tableCell"><input class='width:200px' type="text" name="DCGB" value='<%=enc.EncodeForHtmlAttribute(strDCGB&"")%>'  maxlength="2" onkeyup="if (this.value.length==2) document.clientes.CuentaGB.focus()"/></div><%
			 	        %><div class="width40 tableCell"><input class='width:200px' type="text" name="CuentaGB" value='<%=enc.EncodeForHtmlAttribute(strCuentaGB&"")%>' maxlength="14" /></div><%
			        	%></div><%
	                %></div><%
                    DrawInputCelda "width:200px' maxlength='11","","",11,0,LitBICSWIFT,"bicGB",rst("swift_code2")

                    if rst("TGBBANCODOM") & "" <>"" then
                        datosTGBBANCODOM=enc.EncodeForHtmlAttribute(rst("TGBBANCODOM")&"")
                    else
                        datosTGBBANCODOM=rst("TGBBANCODOM")&""
                    end if
			        DrawInputCelda "width:200px' maxlength='50","","",35,0,LitDomBanco,"DomBancoGB",datosTGBBANCODOM 

                    if rst("TGBBANCOPOB") & "" <>"" then
                        datosTGBBANCOPOB=enc.EncodeForHtmlAttribute(rst("TGBBANCOPOB")&"")
                    else
                        datosTGBBANCOPOB=rst("TGBBANCOPOB")&""
                    end if
                    %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                        %><label><%=LitPoblacion%></label><%
                        %><input class="width50" type="text" size="25" name="PoblacionGB" value="<%=datosTGBBANCOPOB&""%>" onchange="borrarCodigos('3')"/>
			            <input type="hidden" name="codPoblacionGB" value="<%=datosTGBBANCOPOB%>"/> 
			            <input type="hidden" name="codProvinciaGB" value="<%=rst("TGBBANCOPROV")%>"/>
			            <input type="hidden" name="codPaisGB" value=""/>
			            <input type="hidden" name="paisHGB" value="" />
				        <a class='CELDAREFB' class="#dialog1" href="#SELECCIONAR_POBLACION2" onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes5&titulo=<%=LITSSVERPOBLACIONES %>');"  onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"></a><%
                    %></div><%

                    rstSelect.cursorlocation=3
		 	        rstSelect.open "select idprovincia, descripcion from PROVINCIAS with(nolock) order by descripcion",DSNIlion
			        DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitProvincia,"gb_provincia",rstSelect,rst("TGBBANCOPROV")&"","idprovincia","descripcion","onchange","TraerProvinciaGB()"
			        rstSelect.close
            end if
                rstSelect.cursorlocation=3
			    rstSelect.open "select b.nbanco,isnull(b.entidad,'')+'-'+isnull(norma,'')+isnull('-'+case when b.tipo_gasoleo='0' then '"+LitTodosGas+"' else te.descripcion end,'') as entidad from bancos b  with(nolock) left join tipos_entidades te  with(nolock)  on te.codigo=b.tipo_gasoleo and te.codigo like '" & session("ncliente") &"%' where b.nbanco like '" & session("ncliente") & "%'"&" order by b.entidad",session("dsn_cliente")
			    DrawSelectCeldaResponsive1 "width:200px",400,"",0,LitFormBancario,"formatoBanco2",rstSelect,rst("formatoBanco2"),"nbanco","entidad","",""
			    rstSelect.close
			     'dgb: 27/10/2009  Xenteo pago de gasoleo
			     if modulo_Xenteo<>0 then
			        DrawCheckCelda "","","",0,LitPoste,"pagoB",rst("pagoenposteb")
			     end if

			     DrawInputCelda "width:200px' maxlength='16","","",35,0,LitNumTarjeta,"NumTarjeta2",rst("ntarjeta2")
                 DrawInputCelda "width:200px' maxlength='4","","",7,0,LitFCaducidad,"fcaducidad2",rst("fcaducidad2")	

        end if
        %>
        <!--</div>-->
    </div>
  </div>

  	<%'DIRECCION ENVIO MODO EDIT
	rstSelect.cursorlocation=3
	 rstSelect.Open "select * from domicilios where pertenece like '" & session("ncliente") & "%' and codigo='" & rst("dir_envio") & "'", session("dsn_cliente")
	  if not rstSelect.EOF then
	 	domicilio	= rstSelect("domicilio")
		telefono	= rstSelect("telefono")
		poblacion	= rstSelect("poblacion")
		codPoblacion = rstSelect("codpoblacion")
		codPais = rstSelect("codpais")
		codProvincia = rstSelect("codprovincia")
		cp			= rstSelect("cp")
		provincia	= rstSelect("provincia")
		pais		= rstSelect("pais")
		rstSelect.close
	 else
	 	rstSelect.Close
	 end if
	 rstSelect.cursorlocation=3
	 rstSelect.Open "select * from domicilios with(NOLOCK) where pertenece like '" & session("ncliente") & "%' and codigo='" & rst("invoice_address") & "'", session("dsn_cliente")
	 if not rstSelect.EOF then
	 	domicilioF	= rstSelect("domicilio")
		telefonoF	= rstSelect("telefono")
		poblacionF	= rstSelect("poblacion")
		codPoblacionF = rstSelect("codpoblacion")
		codPaisF = rstSelect("codpais")
		codProvinciaF = rstSelect("codprovincia")
		cpF			= rstSelect("cp")
		provinciaF	= rstSelect("provincia")
		paisF		= rstSelect("pais")
		rstSelect.close
	 else
	    rstSelect.Close
	 end if
     %>

      <div  class="Section" id="S_EditDE">
   <a href="#" rel="toggle[EditDE]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
   <%=LitDireccionEnvio%>
   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
             </div>  </a>
    
     
	<% if viene<>"facturas_cli_E" then%>
     <div class="SectionPanel" id="EditDE" style="display: none">
     <%else%>
     <div class="SectionPanel" id="EditDE" >
     <%end if%>
	 
     <table width="100%"  border='0' >
	     <div class="tableCell" style="width:100%">
			<%if si_tiene_modulo_agrario<>0 then
				%><a class="reflink" href="javascript:CopiarCampos()"><%=LitCopiarDirEnvioSocio%></a><%
			else
				%><a class="reflink" href="javascript:CopiarCampos()"><%=LitCopiarDirEnvio%></a><%
			end if%>
			<a class="CELDAREFR7" href="javascript:EliminarDirEnvio('<%=rst("ncliente")%>')" ><img src="../images/<%=ImgVaciarCampo%>" <%=ParamImgVaciarCampo%> alt="<%=LitElimiDirEnvio%>" title="<%=LitElimiDirEnvio%>"/></a>
		</div>
	</table>
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding='2' cellspacing='2'></table>
        
			
	    <%
            if domicilio & "" <>"" then
                datosDomicilio=enc.EncodeForHtmlAttribute(domicilio&"")
            else
                datosDomicilio=domicilio
            end if
            DrawInputCelda "width:200px","","",35,0,LitDomicilio,"de_domicilio",datosDomicilio

			%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
            %><label><%=LitPoblacion%></label><input class="width50" type="text" size="25" name="de_poblacion" value="<%=enc.EncodeForHtmlAttribute(poblacion & "")%>" onchange="borrarCodigos('2')"/>
			<input type="hidden" name="de_codPoblacion" value="<%=enc.EncodeForHtmlAttribute(codPoblacion & "")%>"/> 
			<input type="hidden" name="de_codProvincia" value="<%=enc.EncodeForHtmlAttribute(codProvincia & "")%>"/>
			<input type="hidden" name="de_codPais" value="<%=enc.EncodeForHtmlAttribute(codpais & "")%>"/>
			<input type="hidden" name="de_paisH" value="<%=enc.EncodeForHtmlAttribute(pais & "")%>" />
				<a class='CELDAREFB' class="#dialog1" href="#SELECCIONAR_POBLACION2" onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes2&titulo=<%=LITSSVERPOBLACIONES %>');"  onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
            <% 'AbrirModal "SELECCIONAR_POBLACION3","../configuracion/poblaciones.asp?mode=buscar&viene=clientes2&titulo=SELECCIONAR POBLACION",AnchoVentana,AltoVentana,"no","si","no","si",LitBuscar%>
            </div><%

            DrawInputCelda "width:200px","","",5,0,LitCP,"de_cp",cp

            rstSelect.cursorlocation=3
			 rstSelect.open "select idprovincia, descripcion from PROVINCIAS with(nolock) order by descripcion",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitProvincia,"de_provinciaDDL",rstSelect,codProvincia,"idprovincia","descripcion","onchange","TraerProvinciaDe()","de_provincia",25,provincia
			 rstSelect.close


             rstSelect.cursorlocation=3
			 rstSelect.open "select id, nombre from PAISES with(nolock) order by nombre",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitPais,"de_paisDDL",rstSelect,codPais,"id","nombre","onchange","TraerPaisDe()","de_pais",30,pais
			 rstSelect.close

			 DrawInputCelda "width:200px","","",20,0,LitTel1,"de_telefono",telefono%>
  
<% ' DGM 19/9/11 Ocultamos la direccion de envio de factura %>
  <div style="display:none">
  <% 'Direccion de envio para factura %>
  <table width="100%"  border='0' cellpadding='0' cellspacing='0'>
        <%
            DrawCeldaResponsive1 "width100","","",0,LitInvoiceAddress
            %>

		<div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
			<a class="CELDAREFR7"  href="javascript:EliminarDirEnvioF('<%=rst("ncliente")%>')" ><img src="../images/<%=ImgVaciarCampo%>" <%=ParamImgVaciarCampo%> alt="<%=LitElimiDirEnvio%>" title="<%=LitElimiDirEnvio%>"/></a>
		</div>
	</table>
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=2>
        <%DrawFila color_blau%>
			<div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
				<%if si_tiene_modulo_agrario<>0 then
					%><a class="reflink" href="javascript:CopiarCamposF()"><%=LitCopiarDirEnvioSocio%></a><%
				else
					%><a class="reflink" href="javascript:CopiarCamposF()"><%=LitCopiarDirEnvio%></a><%
				end if%>
			</div>
	    <%
            if domicilioF & "" <>"" then
                datosDomicilioF=enc.EncodeForHtmlAttribute(domicilioF&"")
            else
                datosDomicilioF=domicilioF
            end if
			DrawInputCelda "width:200px","","",35,0,LitDomicilio,"de_domicilioF",domicilioF

			%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
            %><label><%=LitPoblacion%></label><input class="width50" type="text" size="25" name="de_poblacionF" value="<%=enc.EncodeForHtmlAttribute(poblacionF & "")%>" onchange="borrarCodigos('2')"/>
			<input type="hidden" name="de_codPoblacionF" value="<%=enc.EncodeForHtmlAttribute(codPoblacionF & "")%>"/> 
			<input type="hidden" name="de_codProvinciaF" value="<%=enc.EncodeForHtmlAttribute(codProvinciaF & "")%>"/>
			<input type="hidden" name="de_codPaisF" value="<%=enc.EncodeForHtmlAttribute(codpaisF & "")%>"/>
			<input type="hidden" name="de_PaisFH" value="<%=enc.EncodeForHtmlAttribute(paisF & "")%>" />
			
				<a class='CELDAREFB' class="#dialog1" href="#SELECCIONAR_POBLACION2" onclick="javascript:RecargarModales('#SELECCIONAR_POBLACION2','../configuracion/poblaciones.asp?mode=buscar&viene=clientes4&titulo=<%=LITSSVERPOBLACIONES %>');"  onmouseover="self.status='<%=LITSSVERPOBLACIONES%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
            <% 'AbrirModal "SELECCIONAR_POBLACION3","../configuracion/poblaciones.asp?mode=buscar&viene=clientes2&titulo=SELECCIONAR POBLACION",AnchoVentana,AltoVentana,"no","si","no","si",LitBuscar%>
            </div><%

			 DrawInputCelda "width:200px","","",5,0,LitCP,"de_cpF",cpF

             rstSelect.cursorlocation=3
			 rstSelect.open "select idprovincia, descripcion from PROVINCIAS with(nolock) order by descripcion",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitProvincia,"de_provinciaDDLF",rstSelect,codProvinciaF,"idprovincia","descripcion","onchange","TraerProvinciaDeF()","de_provinciaF",25,provinciaF
			 rstSelect.close

             rstSelect.cursorlocation=3
			 rstSelect.open "select id, nombre from PAISES with(nolock) order by nombre",DSNIlion
			 DrawSelectCeldaInput "",200,"",0,LitPais,"de_paisDDLF",rstSelect,codPaisF,"id","nombre","onchange","TraerPaisDeF()","de_paisF",30,paisF
			 rstSelect.close

	         DrawInputCelda "width:200px","","",20,0,LitTel1,"de_telefonoF", telefonoF%>
  </table>
  </div>
  
  </div>
  </div>
	<% 'OTROS DATOS MODO EDIT %>
      <div  class="Section" id="S_EditOD">
   <a href="#" rel="toggle[EditOD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
   <%=LitOtrosDatos%>
   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
      </div>  </a>
     <div class="SectionPanel" id="EditOD" style="display: none">
     
	<table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
		<%
		

		'jcg 02/02/2008
        if si_tiene_modulo_proyectos<>0 then
            %>
				<div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
                    <label><%=iif(mode="browse","<b>"+LitProyecto+"</b>",LitProyecto+"") %></label>
                    <input class="width:200px" type="hidden" name="cod_proyecto" value="<%=rst("proyecto")%>"/>
                    <iframe id='frProyecto' name="fr_Proyecto" src='../mantenimiento/docproyectos.asp?viene=clientes&mode=<%=enc.EncodeForHtmlAttribute(mode& "")%>&cod_proyecto=<%=rst("proyecto")%>' class="width60 iframe-menu" frameborder="no" scrolling="no" noresize="noresize"></iframe>
				</div>
                </table><table width="100% "bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5">
			<%
        end if	
		rstSelect.cursorlocation=3
		rstSelect.open "select codigo, descripcion from tipo_actividad with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
		DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitTActividad,"tactividad",rstSelect,rst("tactividad"),"codigo","descripcion","",""
		rstSelect.close

		rstSelect.cursorlocation=3
		rstSelect.open "select codigo, descripcion from zonas with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
		DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitZona,"zona",rstSelect,rst("zona"),"codigo","descripcion","",""
		rstSelect.close

            if rst("transportista") & "" <>"" then
                datosTransportista=enc.EncodeForHtmlAttribute(rst("transportista") & "")
            else
                datosTransportista=rst("transportista")
            end if
            DrawInputCelda "width:200px","","",25,0,LitTransportista,"transportista",enc.EncodeForHtmlAttribute(datosTransportista & "")

			defecto=rst("portes")

            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                %><label><%=LitPortes%></label>
                <select class="width60" name="portes">
						<%if defecto=LitDebidos then
							%><option selected="selected" value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitPagados then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option selected="selected" value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitMedios then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option selected="selected" value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitOtros then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option selected="selected" value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						elseif defecto=LitPagosCargFact then
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option selected="selected" value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option value=""></option><%
						else
							%><option value="<%=LitDebidos%>"><%=LitDebidos%></option>
							<option value="<%=LitOtros%>"><%=LitOtros%></option>
							<option value="<%=LitPagados%>"><%=LitPagados%></option>
							<option value="<%=LitMedios%>"><%=LitMedios%></option>
							<option value="<%=LitPagosCargFact%>"><%=LitPagosCargFact%></option>
							<option selected="selected" value=""></option><%
						end if%>
					    </select>
				</div><%
			 DrawInputCelda "width:200px","","",25,0,LitHmanyana,"hmanyana",enc.EncodeForHtmlAttribute(rst("hmanyana")& "")

			 DrawInputCelda "width:200px","","",20,0,LitHTarde,"htarde",enc.EncodeForHtmlAttribute(rst("htarde")& "")

			 DrawInputCelda "width:200px","","",6,0,LitPht,"pht",enc.EncodeForHtmlAttribute(rst("pht")& "")

			 DrawInputCelda "width:200px","","",6,0,LitPkm,"pkm",enc.EncodeForHtmlAttribute(rst("pkm")& "")

			 DrawInputCelda "width:200px","","",6,0,LitPD,"pd",enc.EncodeForHtmlAttribute(rst("pd")& "")
		

			rstSelect.cursorlocation=3
			rstSelect.open "select codigo, descripcion from tipos_entidades with(NOLOCK) where codigo like '" & session("ncliente") & "%' and tipo='" & LitCLIENTE & "' order by descripcion",session("dsn_cliente")
			DrawSelectCeldaResponsive1 "width:200px colspan=2",200,"",0,iif(si_tiene_modulo_agrario<>0,LitTSocio,LitTCliente),"tipo_cliente",rstSelect,rst("tipo_cliente"),"codigo","descripcion","",""
			rstSelect.close

		if si_tiene_modulo_ecomerce<>0 then
            DrawDiv "1","display:" & mostrar_verstock,"mostrar_verstock1"
            DrawLabel "'' id='mostrar_verstock2'","",LitMostStockTienda
            DrawCheck "'' id='mostrar_verstock3'", "", "verstock", iif(nz_b(rst("verstock"))<>0,"-1","")
            CloseDiv
		end if
		if si_asesoria=true then
            DrawInputCelda "width:200px' maxlength='25","","",25,0,LitCarpetaNominas,"dsn_nominaplus",rst("dsn_nominaplus")

		    rstSelect.cursorlocation=3
			rstSelect.open "select codigo, descripcion from periodos with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
			DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitPeriodicidadFact,"periodicidad",rstSelect,rst("periodicidad"),"codigo","descripcion","",""
			rstSelect.close


		    rstSelect.cursorlocation=3
			rstSelect.open "select codigo, descripcion from tiendas with (nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
			DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSucursal,"sucursal",rstSelect,rst("tienda"),"codigo","descripcion","",""
			rstSelect.close

            DrawInputCelda "' readonly='true","","",25,0,LitFPublicacion,"fpubnominas",rst("fpubnominas")

            DrawCheckCelda "","","",0,LitMostrarListPortal,"mostrarListPortal",rst("ASESORIALIST") 

		end if
		if si_tiene_modulo_agrario<>0 then
				 DrawInputCelda "width:200px","","",12,0,LitFNacimiento,"fnacimiento",rst("fnacimiento")
				 DrawInputCelda "width:200px' maxlength='20","","",25,0,LitSegSocial,"segsocial",rst("SegSocial")
		elseif si_asesoria=true then
				 DrawInputCelda "width:200px' maxlength='20","","",25,0,LitSegSocial,"segsocial",rst("SegSocial")
	
		end if
		
		if si_tiene_modulo_CRMComunicacion <> 0 then
            DrawCheckCelda "","","",0,LITDELIVERYADV,"submit_advertising",rst("submit_advertising") 
            DrawCheckCelda "","","",0,LITCOMMEMAIL,"email_communication",rst("email_communication") 
            DrawCheckCelda "","","",0,LITCOMMSMS,"sms_communication",rst("sms_communication") 
		end if
        if si_tiene_modulo_fidelizacion30<>0 then
            DrawDiv "1","",""
            DrawLabel "","",LITROL
            DrawSpan "","",iif(null_z(rst("role"))=0,LITCLIENTEMIN, LITNETWORK),""
            CloseDiv
        end if
		%>
  </table>
		    <%Foto "edit",rst("ncliente")%>
  
  </div>
  </div>

	<% 'CONFIG DOC MODO EDIT %>
      <div  class="Section" id="S_EditDD">
   <a href="#" rel="toggle[EditDD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
   <%=LitConfDoc2%>
   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
      </div>  </a>
     <div class="SectionPanel" id="EditDD" style="display: none">
     
     <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
        
        <%'AMF:2/11/2010:Cambiada la obtencion de las series por una llamada a procedure.
        set connSeries = Server.CreateObject("ADODB.Connection")
	    set commandSeries = Server.CreateObject("ADODB.Command")
	           
	    connSeries.open session("dsn_cliente")
	                    
        commandSeries.ActiveConnection = connSeries
	    commandSeries.CommandTimeout = 0
	    commandSeries.CommandText="ObtenerNombreSeriesParaSDefecto"
	    commandSeries.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	    
	    commandSeries.Parameters.Append commandSeries.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente") & "")
        
        

		rstSelect.Open "select * from documentos_cli where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            EligeCelda "check", mode,"","","",0,LitValorPre,"valorado_pre",0,iif(valorado_pre>"",nz_b(valorado_pre),nz_b(rstSelect("valorado_pre")))

			commandSeries.Parameters.Append commandSeries.CreateParameter("@tipodoc",adVarChar,adParamInput,50,"PRESUPUESTO A CLIENTE")
			set rstAux=commandSeries.Execute
            DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSeriePre,"serie_pre",rstAux,rstSelect("serie_pre"),"nserie","nombre","",""
			rstAux.close

            EligeCelda "check", mode,"","","",0,LitValorPed,"valorado_ped",0,iif(valorado_ped>"",nz_b(valorado_ped),nz_b(rstSelect("valorado_ped")))

	 		commandSeries.Parameters("@tipodoc")="PEDIDO DE CLIENTE"
			set rstAux=commandSeries.Execute
	 		DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSeriePed,"serie_ped",rstAux,rstSelect("serie_ped"),"nserie","nombre","",""
	 		rstAux.close

            EligeCelda "check", mode,"","","",0,LitValorAlb,"valorado_alb",0,iif(valorado_alb>"",nz_b(valorado_alb),nz_b(rstSelect("valorado_alb")))

			
			commandSeries.Parameters("@tipodoc")="ALBARAN DE SALIDA"
			set rstAux=commandSeries.Execute
			DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSerieAlb,"serie_alb",rstAux,rstSelect("serie_alb"),"nserie","nombre","",""
			rstAux.close

			commandSeries.Parameters("@tipodoc")="FACTURA A CLIENTE"
			set rstAux=commandSeries.Execute
			DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSerieFac,"serie_fac",rstAux,rstSelect("serie_fac"),"nserie","nombre","",""
			rstAux.close

		'AMF:29/10/2010:Serie de la incidencia.
		if ModuloContratado(session("ncliente"),ModPostVenta) <> 0 then
			commandSeries.Parameters("@tipodoc")="INCIDENCIA"
			set rstAux=commandSeries.Execute
			DrawSelectCeldaResponsive1 "width:200px",200,"",0,LitSerieIncidencia,"serie_incidencia",rstAux,rstSelect("serie_incidencia"),"nserie","nombre","",""
			rstAux.close
		end if
        
		set commandSeries=nothing
    
        connSeries.close
	    set connSeries = nothing		
		rst.close
		
		%>
  </table>
  
  </div>
  </div>

  
<% 'CAMPOS PERSONALIZABLES MODO EDIT %>
  <div  class="Section" id="S_EditCP">
   <a href="#" rel="toggle[EditCP]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader" >
   <%=LitCampPersoCli%>
   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
      </div>  </a>
     <div class="SectionPanel" id="EditCP" style="display: none">
    
	<table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5"><%

		rst.cursorlocation=3
		rst.open "select * from camposperso with(NOLOCK) where tabla='CLIENTES' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
		if not rst.eof then
			num_campos_existen=rst.recordcount
			DrawFila ""
				num_campo=1
				num_campo2=1
				num_puestos=0
				num_puestos2=0
				while not rst.eof
					if num_puestos2>0 and (num_puestos2 mod 2)=0 then
						num_puestos2=0
					end if

                    if rst("titulo") & "" <>"" then
                       datosTitulo=enc.EncodeForHtmlAttribute(rst("titulo") & "")
                    else
                       datosTitulo=rst("titulo")
                    end if

					if rst("titulo") & "">"" and nz_b(rst("system_reg")) <> -1 then
						num_puestos=num_puestos+1
						num_puestos2=num_puestos2+1
						'DrawCelda "CELDA7 style='width:150px'","","",0,rst("titulo") & " : "
						valor_campo_perso=lista_valores(num_campo)
						if rst("tipo")=1 then
							if isNumeric(rst("tamany")) then
								tamany=rst("tamany")
							else
								tamany=1
							end if
                            DrawDiv "1","",""
                            DrawLabel "","",datosTitulo
                            %>
								<input type="text" name="<%="campo" & enc.EncodeForHtmlAttribute(num_campo)%>" size="35" maxlength="<%=enc.EncodeForHtmlAttribute(tamany)%>" value="<%=enc.EncodeForHtmlAttribute(valor_campo_perso)%>"/>
							<%
                            CloseDiv
                        elseif rst("tipo")=2 then
                            DrawCheckCelda "","","",0,datosTitulo,"campo" & enc.EncodeForHtmlAttribute(num_campo),iif(valor_campo_perso="1",-1,0)
                        elseif rst("tipo")=3 then
							num_campo_str=cstr(num_campo)
							if len(num_campo_str)=1 then
								num_campo_str="0" & num_campo_str
							end if
							
							sem=0
							if (num_campo_str="01" and c01="c") or (num_campo_str="02" and c02="c") or (num_campo_str="03" and c03="c") or (num_campo_str="04" and c04="c") or (num_campo_str="05" and c05="c") or (num_campo_str="06" and c06="c") or (num_campo_str="07" and c07="c") or (num_campo_str="08" and c08="c") or (num_campo_str="09" and c09="c") or (num_campo_str="10" and c10="c") or (num_campo_str="11" and c11="c") or (num_campo_str="12" and c12="c") or (num_campo_str="13" and c13="c") or (num_campo_str="14" and c14="c") or (num_campo_str="15" and c15="c") or (num_campo_str="16" and c16="c") or (num_campo_str="17" and c17="c") or (num_campo_str="18" and c18="c") or (num_campo_str="19" and c19="c") or (num_campo_str="20" and c20="c") then
							    strSelListVal="select dni as ndetlista, nombre as valor from personal with(NOLOCK),comerciales with(NOLOCK) where personal.dni like '" & session("ncliente") & "%' and comerciales.comercial like '" & session("ncliente") & "%' and comerciales.fbaja is null and dni like '" & session("ncliente") & "%' and dni=comercial order by valor,ndetlista"
							    sem=1
							else
							    strSelListVal="select ndetlista,valor from campospersolista with(NOLOCK) where tabla='CLIENTES' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' " &  strWhereListaDep & " order by valor,ndetlista"
							end if
							
							rstAux.cursorlocation=3
							rstAux.open strSelListVal,session("dsn_cliente")
							%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12">
                                <label><%=datosTitulo%></label>
                                <select class="width60" name="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" id="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>">
									<%
									encontrado=0
									while not rstAux.eof
										
										if sem=0 then
										    if valor_campo_perso & "">"" and isnumeric(valor_campo_perso) then
											    valor_campo_perso_aux=clng(valor_campo_perso)
										    else
											    valor_campo_perso_aux=0
										    end if
										    if valor_campo_perso_aux=clng(rstAux("ndetlista")) then
											    texto_selected="selected"
											    if encontrado=0 then encontrado=1
										    else
											    texto_selected=""
										    end if
										else
										    if valor_campo_perso & "">"" then
											    valor_campo_perso_aux=valor_campo_perso
										    else
											    valor_campo_perso_aux="0"
										    end if
										    if valor_campo_perso_aux=rstAux("ndetlista")&"" then
											    texto_selected="selected"
											    if encontrado=0 then encontrado=1
										    else
											    texto_selected=""
										    end if
										end if
										%>
										<option value="<%=enc.EncodeForHtmlAttribute(rstAux("ndetlista") & "")%>"  <%=enc.EncodeForHtmlAttribute(texto_selected & "")%> ><%=enc.EncodeForHtmlAttribute(rstAux("valor") & "")%></option>
										<%rstAux.movenext
									wend%>
									<option <%=iif(encontrado=1,"","selected")%> value=""></option>
								</select>
							</div><%
							rstAux.close
						elseif rst("tipo")=4 then
							if isNumeric(rst("tamany")) then
								tamany=rst("tamany")
							else
								tamany=1
							end if
                            DrawDiv "1","",""
                            DrawLabel "","",datosTitulo
                            %>
								<input type="text" name="<%="campo" & enc.EncodeForHtmlAttribute(num_campo)%>" size="35" maxlength="<%=tamany%>" value="<%=enc.EncodeForHtmlAttribute(valor_campo_perso)%>"/>
							<%
                            CloseDiv
                        elseif rst("tipo")=5 then
							if isNumeric(rst("tamany")) then
								tamany=rst("tamany")
							else
								tamany=1
							end if
                            DrawDiv "1","",""
                            DrawLabel "","",datosTitulo
                            %>
                                <input type="text" name="<%="campo" & enc.EncodeForHtmlAttribute(num_campo)%>" size="30" maxlength="<%=tamany%>" value="<%=enc.EncodeForHtmlAttribute(valor_campo_perso)%>"/>
							<%
                            DrawCalendar "campo"&enc.EncodeForHtmlAttribute(num_campo)
                            CloseDiv
						end if

					else
                        valor_campo_perso=lista_valores(num_campo)
						%><input type="hidden" name="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" id="campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" value="<%=enc.EncodeForHtmlAttribute(valor_campo_perso)%>"/><%
					end if%>
                    <input type="hidden" name="tipo_campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" value="<%=enc.EncodeForHtmlAttribute(rst("tipo") & "")%>"/>
					<input type="hidden" name="titulo_campo<%=enc.EncodeForHtmlAttribute(num_campo)%>" value="<%=datosTitulo%>"/>
                    <%rst.movenext
					num_campo=num_campo+1
					if not rst.eof then
						if rst("titulo") & "">"" then
							num_campo2=num_campo2+1
						end if
					end if
				wend

			num_campos=num_puestos
		else
			num_campos=0
			num_campos_existen=0
		end if
		rst.close
	%></table>
	<input type="hidden" name="num_campos" value="<%=enc.EncodeForHtmlAttribute(num_campos_existen)%>"/>

</div>
</div>

<%
 BarraNavegacion "Edit"
if viene="facturas_cli_E" then
    %>
        <script language=javascript type="text/javascript">
            document.clientes.rsocial.focus();
            document.clientes.rsocial.select();
        </script>
    <%
end if
%>
   <% 
    'mmgBuscar
   elseif mode="search2" then
		lote=limpiaCadena(Request.QueryString("lote"))
		nt=limpiaCadena(request.QueryString("nt"))
		if lote="" then lote=1
		sentido=limpiaCadena(Request.QueryString("sentido"))
		firstReg=limpiaCadena(request.querystring("firstReg"))
		lastReg=limpiaCadena(request.querystring("lastReg"))
        ''ricardo 31-5-2006 encontramos el primer y ultimo ncliente de toda la select
        ''este primer y ultimo sirve para que sepamos cuando quitar el icono de primer o ultimo registro
        ''si no , no habra manera de saberlo
        firstRegAll=limpiaCadena(request.querystring("firstRegAll"))
        lastRegAll=limpiaCadena(request.querystring("lastRegAll"))

		strwhere="where clientes.ncliente like '" & session("ncliente") & "%' And Domicilios.codigo=clientes.dir_principal "

		if agente>"" then
			strwhere= strwhere & " and agente='" & agente & "' "
		end if
		if vienecomercial>"" then
			strwhere= strwhere & " and comercial='" & vienecomercial & "' "
		end if
        if comercialSolSusCli & "">"" then
            strwhere= strwhere & " and comercial='" & comercial & "' "
        end if

        if nt="" then
		    strwhere=strwhere & CadenaBusqueda(campo,criterio,texto,vienecomercial,agente)
		else 
		    strwhere=strwhere & CadenaBusquedaMM(criterio,texto,nt)
		end if
        ''ricardo 31-5-2006 encontramos el primer y ultimo cliente de toda la select
        ''este primer y ultimo sirve para que sepamos cuando quitar el icono de primer o ultimo registro
        ''si no , no habra manera de saberlo
        strwhere2=strwhere

		if sentido="prev" then
			strwhere=strwhere & " and clientes.ncliente<'" & firstReg & "'"
			lote=clng(lote)-1
		elseif sentido="next" then
			strwhere=strwhere & " and clientes.ncliente>'" & lastReg & "'"
			lote=clng(lote)+1
		end if

        strSql2=""
        strSql3=""
		strSeleccion= "select clientes.ncliente,rsocial,comercial,poblacion,telefono,tipo_cliente,tactividad,domicilio,ncomercial,contacto,provincia,telefono,cif "
		strfrom=" from clientes with (NOLOCK),domicilios with (NOLOCK)"

        ''mmg 31-03-2008: Vamos a decuar la select para poder realizar consultas en los campos adicionales
        'dependiendo del tipo de campo adicional realizaremos la consulta de una u otra forma
        if texto <> "" then
            if nt=1 or nt=2 or nt=4 or nt=5 then
                strfrom= strfrom & ",ficha_cli f with(nolock)"
                strwhere= strwhere & " and clientes.ncliente=f.ncliente and f.ncampo='"& session("ncliente")& campo &"'"
            else if nt=3 then
                strfrom= strfrom & ",ficha_cli f with(nolock),campospersolista cper with(nolock)"
                strwhere= strwhere & " and clientes.ncliente=f.ncliente and f.ncampo='"& session("ncliente")& campo &"' and f.ncampo=cper.ncampo and f.valor=cper.ndetlista"
            end if
            end if
       end if
      
        strSelSearch=strSeleccion & strfrom & strwhere & strOrder
  
        ''ricardo 31-5-2006 encontramos el primer y ultimo cliente de toda la select
        ''este primer y ultimo sirve para que sepamos cuando quitar el icono de primer o ultimo registro
        ''si no , no habra manera de saberlo
		strSql2="select top 1 clientes.ncliente "
		strOrder=" order by clientes.ncliente"

        strSql2=strSql2 & strfrom & strwhere
        strSql3=strSql2

		if sentido="prev" or sentido="last" then strOrder=strOrder & " desc"

        ''ricardo 31-5-2006 encontramos el primer y ultimo cliente de toda la select
        ''este primer y ultimo sirve para que sepamos cuando quitar el icono de primer o ultimo registro
        ''si no , no habra manera de saberlo
        if firstRegAll & ""="" then
			strSql2=strSql2 & " order by clientes.ncliente"
			rst.cursorlocation=3
			rst.Open strSql2,session("dsn_cliente")
			if not rst.eof then
				firstRegAll=rst("ncliente")
				strwhere=strwhere & " and clientes.ncliente>='" & firstRegAll & "'"
			end if
			rst.close
        end if
        if lastRegAll & ""="" then
			strSql3=strSql3 & " order by clientes.ncliente desc"
			rst.cursorlocation=3
			rst.Open strSql3,session("dsn_cliente")
			if not rst.eof then
				lastRegAll=rst("ncliente")
				strwhere=strwhere & " and clientes.ncliente<='" & lastRegAll & "'"
			end if
			rst.close
        end if

        strSelSearch=strSeleccion & strfrom & strwhere & strOrder
		rst.cursorlocation=3
		rst.maxrecords=NumReg

		rst.Open strSelSearch,session("dsn_cliente")

		if not rst.EOF then
			
				if sentido="prev" or sentido="last" then rst.sort="ncliente"
				firstReg=rst("ncliente")
				rst.movelast
				lastReg=rst("ncliente")
				rst.movefirst

                if firstRegAll & "">"" then
	                if lote=1 then
		                if firstReg=firstRegAll then
			                lote=1
		                else
			                lote=2
		                end if
	                end if
                end if

                        if lastReg<>lastRegAll or lote<>1 then
				                        NextPrev lote,firstReg,lastReg,campo,criterio,texto,sentido,firstRegAll,lastRegAll
                        end if
						%><br/>
						<%if agente>"" or vienecomercial>"" then%>
							<table width='100%' border='0' cellspacing="1" cellpadding="1">
						<%else%>
							<table width=150% border='0' cellspacing="1" cellpadding="1">
						<%end if
			    	  'Fila de encabezado
					   	if agente>"" or vienecomercial>"" then
					   		'DrawFila color_fondo
                            %><tr bgcolor='<%=color_fondo%>'><%
					    		DrawCelda2 "ENCABEZADOL","","",LitCodigo
								DrawCelda2 "ENCABEZADOL","","",LitCli
								if si_tiene_modulo_comercial<>0 then
									DrawCelda2 "ENCABEZADOL","","",LitComercialModCom
								else
									DrawCelda2 "ENCABEZADOL","","",LitComercial
								end if
								DrawCelda2 "ENCABEZADOL","","",LitPoblacion
								DrawCelda2 "ENCABEZADOL","","",LitTel1
								DrawCelda2 "ENCABEZADOL","","",LitTCliente
								DrawCelda2 "ENCABEZADOC","","",LitTActividad
                            %></tr><%
					   		'CloseFila
					   	end if

						VinculosPagina(MostrarClientes)=1
						CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

						fila=1
						while not rst.EOF and fila<=NumReg
							'Seleccionar el color de la fila.
							if ((fila+1) mod 2)=0 then
								color=color_blau
							else
								color=color_terra
							end if


							if agente>"" or vienecomercial>"" then
								'DrawFila color
                                 %><tr bgcolor='<%=color%>'><%
									%><td class="CELDA" algin="left"><%
										response.write(Hiperv(OBJClientes,rst("ncliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("ncliente")),""))
									%></td><%
									DrawCelda2 "CELDA","","",enc.EncodeForHtmlAttribute(rst("rsocial") & "")
									DrawCelda2 "CELDA", "left", false, d_lookup("nombre","personal","dni like '" & session("ncliente") & "%' and dni='" & rst("comercial") & "'",session("dsn_cliente"))
									DrawCelda2 "CELDA","","",enc.EncodeForHtmlAttribute(rst("poblacion") & "")
									DrawCelda2 "CELDA","","",rst("telefono")
									DrawCelda2 "CELDA","","",d_lookup("descripcion","tipos_entidades","codigo like '" & session("ncliente") & "%' and codigo='" & rst("tipo_cliente") & "'",session("dsn_cliente"))
									DrawCelda2 "CELDA","","",d_lookup("descripcion","tipo_actividad","codigo like '" & session("ncliente") & "%' and codigo='" & rst("tactividad") & "'",session("dsn_cliente"))
								%></tr><%
                                'CloseFila
							end if

							fila=fila+1

							rst.MoveNext
						wend
			 			%></table>
			 	<br/>
			 	<%if lastReg<>lastRegAll or lote<>1 then
				    NextPrev lote,firstReg,lastReg,campo,criterio,texto,sentido,firstRegAll,lastRegAll
                end if
				 rst.close
			
        else 'NO HAY REGISTROS%>
		    <font class='CEROFILAS'><%=LitCeroFilas%></font><%
        end if
    end if

   ' AbrirModal "SELECCIONAR_POBLACION2","../configuracion/poblaciones.asp?mode=buscar&viene=clientes&titulo=SELECCIONAR POBLACION",AnchoVentana,AltoVentana,"no","si","no","si",LitBuscar
    if mode="browse" then
        if confirmChange then%>
            <script language="javascript" type="text/javascript">
                if (confirm("<%=LitConfirmDataCustomer%>")) ConfirmChangesCustomer();
            </script>
        <%end if
    end if%>
 </form>
<%end if
set rst=nothing
set rst2=nothing
set rst3=nothing
set rst4=nothing
set rstAux=nothing
set rstAux2=nothing
set rstAux3=nothing
set rstDomi=nothing
set rstDomi2=nothing
set rstOrcu=nothing
set rstSelect=nothing
set rstCom=nothing
set rstCF=nothing
set rstCP=nothing
set rstComer=nothing
set rstC=nothing
set rstCC=nothing
set rc=nothing
set rc2=nothing
set rc3=nothing
set rstCambioCliente=nothing
set rstNC=nothing
set connSeries=nothing
set conn=nothing
set connId=nothing
set commandSeries=nothing
set command=nothing
set commandId=nothing


%>

<% AbrirModal "SELECCIONAR_POBLACION2","../configuracion/poblaciones.asp?mode=buscar&viene=clientes&titulo=SELECCIONAR POBLACION",AnchoVentana,AltoVentana,"no","si","no","si",LitBuscar%>

     

</body>
</html>
