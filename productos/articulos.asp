<%@ Language=VBScript %>
<% 
dim enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
function pintar_saltos_nuevo(texto)
	texto=Replace(texto,"&#10;","")
	texto=Replace(texto,"&#13;","<br>")
	pintar_saltos_nuevo=texto
end function
%>
<%
'CODIGOS DE AÑADIDURAS/MODIFICACIONES -----------------------------------------------------
'JCI-18122002-01 : Se añade el control de borrado del artículo por si aparece en la tabla
'                  DETALLES_ORDEN_FAB
'      FECHA     : 18/12/2002
'      AUTOR     : JCI
'
'JCI-18022003-01 : Gestión de la asignación de un circuito de fabricación al artículo
'      FECHA     : 18/02/2003
'      AUTOR     : JCI

''ricardo 6/3/2003
''se pone , para que en la ficha almacenes, se pueda listar
''los pedidos pendientes de recibir y de servir

' JCI 17/03/2003 : Hipervínculo en la pestaña de fabricación para ver los sobrantes existentes
' VGR 20/03/2003 : Hipervínculo en la pestaña de Propiedades para ver los datos adicionales.

''ricardo 10/4/2003 - si carga_terminal, creamos un registro en articuloster

''ricardo 17/4/2003 - se pone el tipo de articulo
    
''ricardo 24/4/2003 - se controla si existe en articuloster la referencia

' JCI 13/06/2003 : MIGRACION A MONOBASE

'JCI 20/11/2003 : Nuevo formato de presentación

'ricardo 24-3-2004 se añade pestaña de campos personalizables
'JMA 2/4/04: Se añade pestaña de Datos Contables
'JCI 22/06/2004 : Gestión de diversos códigos de barras de artículo
'RGU 27/12/2005: gestionar ofertas de shades
'EJM 08/06/2006: Inserción de nuevos campos proyecto LENTICOM
'				 Revisión del tamaño de las capa de las distintas opciones de carpeta
'RGU 13/10/2006: Añadir campo pvp+iva en el span de precios

'------------------------------------------------------------------------------------------%>
<% Server.ScriptTimeout = 1200 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<META HTTP-EQUIV="Content-Type" Content="<%=imgtipo%>;charset=<%=session("caracteres")%>">

<LINK REL="styleSHEET" href="../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../impresora.css" MEDIA="PRINT">

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->

<!--#include file="articulos.inc" -->
<!--#include file="../productos/listados/codigo_barras.inc" -->
<!--#include file="propiedadesarticulos_carga.inc" -->

<!--#include file="../varios2.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../common/modal2.inc" --> 

<!--#include file="../tablasResponsive.inc" -->


<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/dropdown.js.inc" -->

<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="articulos_linkextra.inc" --> 
<!--#include file="../styles/formularios.css.inc" --> 
<!--#include file="../styles/dropdown.css.inc" -->



<%  '************************************************************************
	' CARGA DE INFORMACION DE MODULOS
	'************************************************************************
	si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
	si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)
	si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)
	si_tiene_modulo_ecomerce=ModuloContratado(session("ncliente"),ModEComerce)
	si_tiene_modulo_terminales=ModuloContratado(session("ncliente"),ModTerminales)
	si_tiene_modulo_importaciones=ModuloContratado(session("ncliente"),ModImportaciones)
	si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
    ' osm 27/04/15 comprobar si tiene módulo intelitur
    si_tiene_modulo_intelitur=ModuloContratado(session("ncliente"),ModIntelitur)
	' JMMM - 03/11/2010 -> Franquicias
	si_tiene_modulo_franquicia = ModuloContratado(session("ncliente"),ModFranquiciasTiendas)

    si_tiene_modulo_OrCU=ModuloContratado(session("ncliente"),ModOrCU)


	esFranquiciador=d_lookup("franquiciador", "configuracion", "nempresa='"&session("ncliente")&"'", session("dsn_cliente"))
	
	set conn = Server.CreateObject("ADODB.Connection")
	
	''MPC 26/05/2009 Se obtiene el campo horecas de configuración para cambiar el nombre de un campo o dejarlo
    horecas=d_lookup("horecas", "configuracion", "nempresa='"&session("ncliente")&"'", session("dsn_cliente"))
    '*** i AMP 31082011 
    changeprices = d_lookup("CHANGE_PRICES_PRODUCTSTRUCTURE", "configuracion", "nempresa='"&session("ncliente")&"'", session("dsn_cliente")) '*** f    

    ' osm 27/04/15 obtener provincia e id de la provincia del usuario
    if si_tiene_modulo_intelitur <> 0 then
        provinciaUser = d_lookup("provincia", "domicilios", "pertenece='"&session("ncliente")&session("usuario")&"' and tipo_domicilio='personal'", session("dsn_cliente"))
        idProvinciaUser = d_lookup("NDETLISTA", "CAMPOSPERSOLISTA", "valor='"&provinciaUser&"' and tabla='ARTICULOS' and ncampo='"&session("ncliente")&"02'", session("dsn_cliente"))
    end if
%>

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('addDCONT', 'fade=1');
    animatedcollapse.addDiv('addPRO', 'fade=1');
    //animatedcollapse.addDiv('addMD', 'fade=1');
    animatedcollapse.addDiv('addDCONT2', 'fade=1');
    animatedcollapse.addDiv('addPRO2', 'fade=1');
    animatedcollapse.addDiv('addMD2', 'fade=1');
    animatedcollapse.addDiv('addCPerso', 'fade=1');

    animatedcollapse.addDiv('BrowseDCONT', 'fade=1');
    animatedcollapse.addDiv('BrowsePRO', 'fade=1');
    //animatedcollapse.addDiv('BrowseMD', 'fade=1');
    animatedcollapse.addDiv('BrowseDCONT2', 'fade=1');
    animatedcollapse.addDiv('BrowsePRO2', 'fade=1');
    animatedcollapse.addDiv('BrowseMD2', 'fade=1');
    animatedcollapse.addDiv('BrowseFO', 'fade=1');
    animatedcollapse.addDiv('BrowsePRE', 'fade=1');
    animatedcollapse.addDiv('BrowseCPerso', 'fade=1');

    animatedcollapse.addDiv('editDCONT', 'fade=1');
    animatedcollapse.addDiv('editPRO', 'fade=1');
    //animatedcollapse.addDiv('editMD', 'fade=1');
    animatedcollapse.addDiv('editDCONT2', 'fade=1');
    animatedcollapse.addDiv('editPRO2', 'fade=1');
    animatedcollapse.addDiv('editMD2', 'fade=1');
    animatedcollapse.addDiv('editFO', 'fade=1');
    animatedcollapse.addDiv('editCPerso', 'fade=1');

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init();

    jQuery(document).on("ready", function () {
        setTimeout(
                        function () { jQuery("#frPropiedades").height(jQuery("#frPropiedades").contents().find("form[name='PropiedadesArticulo']").height()) },1000);
    })

    
    var es_padreGlobal = 0;
    var reftmp = "";
    var esperar_grabacion_foto = 0;
    var cuanto_espero = 5;

    function VerEquipos(campo, criterio, texto, almacen) {
        document.location = "../mantenimiento/equipos.asp?mode=search&campo=" + campo +
            "&criterio=" + criterio + "&texto=" + texto + "&almacen=" + almacen + "&vengode=articulos";
        parent.botones.document.location = "../mantenimiento/equipos_bt.asp?mode=search";
    }

    //***************************************************************************************************************
    function TraeDivisa() {
        document.articulos.action = "articulos.asp?mode=add&refresco=buscar&t=<%=enc.EncodeForJavascript(tarifa)%>";
        document.articulos.submit();
    }
    //***************************************************************************************************************
    function CalculaPreEsc() {
        document.articulos.i_impesc.value = document.articulos.i_impesc.value.replace(".", ",");
    }

    //*****************************************************************************************************************
    //ricardo 5-7-2012, la siguiente funcion es para las teclas rapidas
    function Insertar() {
        var ref = "";
        var es_padre = 0;
        var pvp = "0";
        var importe = "0";
        try {
            ref = document.articulos.T_referencia.value;
        }
        catch (e) {
            ref = "";
        }
        try {
            es_padre = document.articulos.T_es_padre.value;
        }
        catch (e) {
            es_padre = 0;
        }
        try {
            pvp = document.articulos.T_pvp.value;
        }
        catch (e) {
            pvp = "0";
        }
        try {
            importe = document.articulos.T_importe.value;
        }
        catch (e) {
            importe = "0";
        }
        //window.alert("los datos son-" + ref + "-" + es_padre + "-" + pvp + "-" + importe + "-");
        try {
            AnadirPrecio(ref, es_padre, pvp, importe);
        }
        catch (e) {
        }
    }
    function AnadirPrecio(ref, es_padre, pvp, importe) {
        if (es_padre != 0) {
            if (document.articulos.i_hijotcp.value != "") {
                ref = document.articulos.i_hijotcp.value;
                es_padreGlobal = 0;
            }
            else es_padreGlobal = -1;
        }

        if (es_padreGlobal != 0) {
            if (!confirm("<%=LitMsgCambiosTallColConfirm%>")) {
                return false;
            }
        }
        precio = document.articulos.precio.value.replace(",", ".");
        dto = document.articulos.descuento.value.replace(",", ".");
        dtocoste = document.articulos.descuentocoste.value.replace(",", ".");
        if (precio == "" && dto == "" && dtocoste == "") {
            alert("<%=LitMsgPrecDtoNoNulo%>");
            return false;
        }
        si_tiene_modulo_tiendas = document.articulos.si_tiene_modulo_tiendas.value;

        if (si_tiene_modulo_tiendas != 0) {
            if (document.articulos.tarifa.value == "" && document.articulos.temporada.value == "" && document.articulos.rango.value == "") {
                alert("<%=LitMsgTarTempRanNoNulo%>");
                return false;
            }
        }
        else {
            if (document.articulos.tarifa.value == "") {
                alert("<%=LitMsgTarTempRanNoNulo%>");
                return false;
            }
        }

        if (isNaN(precio) || isNaN(dto) || isNaN(dtocoste)) {
            alert("<%=LitMsgPreDesNumerico%>");
            return false;
        }
        if (parseFloat(precio) < 0) {
            alert("<%=LitMsgPVPNoNegativo%>");
            return false;
        }
        if (parseFloat(dto) < -100) {
            dto = "-100";
        }
        if (parseFloat(dtocoste) < -100) {
            dtocoste = "-100";
        }

        document.articulos.precio.value = precio.replace(".", ",");
        document.articulos.descuento.value = dto.replace(".", ",");
        document.articulos.descuentocoste.value = dtocoste.replace(".", ",");
        cad_src = "PreciosDeArticulo.asp?mode=addprecio&referencia=" + ref;
        cad_src = cad_src + "&tarifa=" + document.articulos.tarifa.value;
        if (si_tiene_modulo_tiendas != 0) {
            cad_src = cad_src + "&temporada=" + document.articulos.temporada.value;
            cad_src = cad_src + "&rango=" + document.articulos.rango.value;
        }
        cad_src = cad_src + "&precio=" + document.articulos.precio.value;
        cad_src = cad_src + "&dtocoste=" + document.articulos.descuentocoste.value;
        cad_src = cad_src + "&pvp=" + pvp;
        cad_src = cad_src + "&importe=" + importe;
        cad_src = cad_src + "&descuento=" + document.articulos.descuento.value + "&espadre=" + es_padreGlobal;
        marcoPreciosArticulo.document.location = cad_src;
        document.articulos.precio.value = "";
        document.articulos.precioiva.value = "";
        document.articulos.descuento.value = "";
        document.articulos.descuentocoste.value = "";
        document.getElementById("incimp").innerHTML = "";
        document.getElementById("incpvp").innerHTML = "";
        document.articulos.preciofinal.value = "";
        document.articulos.tarifa.value = "";
        if (si_tiene_modulo_tiendas != 0) {
            document.articulos.temporada.value = "";
            document.articulos.rango.value = "";
        }
        document.articulos.checkb.checked = true;
    }

    //*****************************************************************************************************************
    function GestionPropiedades(mode, ref, es_padre) {
        if (mode == "delete") {
            if (!confirm("<%=LitMsgEliminarArticuloConfirm%>")) return false;
        }
        f = marcoPropiedades.document.PropiedadesArticulo;

        //ricardo 27-12-2006 se añade los campos para los formatos de etiquetas
        if (isNaN(f.cantidadarticulo.value.replace(",", "."))) {
            window.alert("<%=LitMedNoNum%>");
            f.cantidadarticulo.focus();
            f.cantidadarticulo.select();
            return false;
        }


        if (f.urlRewrite != null && f.urlRewrite.value != "") {
            if (f.urlRewrite.value == "True") {
                if (!checkURLPattern(f.url_mostrar.value)) {
                    window.alert('<%=LITURLMOSTRARINVALIDA%>');
                    try { f.url_mostrar.focus(); } catch (err) { }
                    return false;
                }
            }
        }

        refSinEmp = trimCodEmpresa(f.hreferencia.value);
        if (f.prefcodbarras.value == "CI" && refSinEmp != "" && (isNaN(refSinEmp) || refSinEmp.length > 5)) {
            window.alert("<%=LitMsgCodBarImpRefNum%>");
            return false;
        }

        if (f.prefcodbarras.value == "CI" && es_padre != 0) {
            window.alert("<%=LitMsgCodBarImpTyC%>");
            return false;
        }

        if (es_padre != 0) {
            if (!confirm("<%=LitMsgCambiosTallColConfirm%>")) {
                return false;
            }
        }

        if (f.nombre.value != "") {
            //trim
            f.nombre.value = f.nombre.value.replace(/^\s+|\s+$/g, '');
            //delete tabs at the front and end
            f.nombre.value = f.nombre.value.replace(/^\t+|\t+$/g, '');

            //we dont permit tab char
            if (f.nombre.value.indexOf(String.fromCharCode(9)) >= 0) {
                window.alert("<%=LitMsgRefDesCarNoVal%>");
                return false;
            }
        }

        if (f.nombre.value == "") {
            window.alert("<%=LitMsgNombreNoNulo%>");
            return false;
        }
        

        if (comp_car_ext(f.nombre.value, 2) == 1) {
            window.alert("<%=LitMsgRefDesCarNoVal%>");
            return false;
        }

        if (isNaN(f.spvp.value.replace(",", ".")) || f.spvp.value == "" || f.spvp.value < 0) {
            window.alert("<%=LitMsgPvPNumerico%>");
            return false;
        }
        if (f.divisa.value == "") {
            window.alert("<%=LitMsgDivisaNoNulo%>");
            return false;
        }
        if (f.iva.value == "") {
            window.alert("<%=LitMsgIvaNoNulo%>");
            return false;
        }
        //Descuento
        if (f.descuento.value > 100 || f.descuento.value < 0) {
            window.alert("<%=LITDTORANGE%>");
            return false;
        }

        if (isNaN(f.descuento.value.replace(",", "."))) {
            window.alert("<%=LitMsgDescuentoNumerico%>");
            return false;
        }
         if (f.mandatorySubfamily.value == "True") {
             if (f.auxfamilia.value == "") {
                alert("<%=LITVALSUBFAMILY%>");
                document.getElementByName.auxfamilia.focus();
                return false;
             }
        }

        if (isNaN(f.peso.value.replace(",", "."))) {
            window.alert("<%=LITPESONONUMERICO%>");
            f.peso.focus();
            return false;
        }
        <% 'i(EJM 08/06/2006)%>
            <%if si_tiene_modulo_tiendas = 0 and horecas = 0 then %>
        if (isNaN(f.mesesCaducidad.value)) {
            window.alert("<%=LitMsgMesesNumerico%>");
            return false;
        }
        <% end if %>
        if (isNaN(f.importeABparcial.value.replace(",", "."))) {
                window.alert("<%=LitMsgImporteABparcialNumerico%>");
                return false;
            }
        <% 'fin(EJM 08/06/2006)%>


        si_tiene_modulo_comercial = document.articulos.si_tiene_modulo_comercial.value;
        if (si_tiene_modulo_comercial != 0) {
            if (isNaN(f.porcom.value.replace(",", "."))) {
                window.alert("<%=LitMsgPorComisionNumerico%>");
                return false;
            }
        }
        if (isNaN(f.mesesmo.value.replace(",", ".")) || isNaN(f.mesesde.value.replace(",", ".")) || isNaN(f.mesesmt.value.replace(",", "."))) {
            window.alert("<%=LitMsgMesesNumerico%>");
            return false;
        }
        si_tiene_modulo_terminales = document.articulos.si_tiene_modulo_terminales.value;
        if (si_tiene_modulo_terminales != 0) {
            if ((f.carga_terminal.checked) && (f.codbarras.value == "" && f.prefcodbarras.value == "")) { //&& (f.checkcodbarras.value=="false" || f.checkcodbarras.value==""))){
                window.alert("<%=LitMsgNoCargaTerminal%>");
                return false;
            }
        }
        <% uploadFileLimitj2=204800
        uploadFileLimit2 = 205190 %>
        if (marcoFoto.document.upload.blob1.value != "") {
            //window.alert(navigator.appName);
            if (navigator.appName == "Microsoft Internet Explorer") {
                var fso = new ActiveXObject("Scripting.FileSystemObject");
                //window.alert("1");
                if (!fso.FileExists(marcoFoto.document.upload.blob1.value)) {
                    window.alert("<%=LitFicheroNoExiste%>");
                    return false;
                }
                //window.alert("2-<%=cLng(uploadFileLimitj2)%>-" + fso.GetFile(marcoFoto.document.upload.blob1.value).Size + "-" + (fso.GetFile(marcoFoto.document.upload.blob1.value).Size><%=cLng(uploadFileLimitj2)%>) + "-");
                if (fso.GetFile(marcoFoto.document.upload.blob1.value).Size ><%= cLng(uploadFileLimitj2) %>) {
                    window.alert("<%=LitFicTamGrandeFicArt1%> <%=formatnumber(((uploadFileLimit2-390)/1024),0,-1,0,-1)%><%=LitFicTamGrandeFicArt%>");
                    return false;
                }
                //window.alert("3");
            }
            else {
                //window.alert("4-" + marcoFoto.document.upload.blob1.size + "-" + marcoFoto.document.upload.blob1.files[0].size + "-<%=cLng(uploadFileLimitj2)%>-" + (marcoFoto.document.upload.blob1.size><%=cLng(uploadFileLimitj2)%>) + "-");
                if (marcoFoto.document.upload.blob1.files[0].size ><%= cLng(uploadFileLimitj2) %>) {
                    alert("<%=LitFicTamGrandeFicArt1%> <%=formatnumber(((uploadFileLimit2-390)/1024),0,-1,0,-1)%><%=LitFicTamGrandeFicArt%>");
                    return false;
                }
                //window.alert("5");
            }
        }
        //window.alert("6");

        if (marcoFoto.document.upload.blob2.value != "") {
            if (navigator.appName == "Microsoft Internet Explorer") {
                var fso = new ActiveXObject("Scripting.FileSystemObject");
                if (!fso.FileExists(marcoFoto.document.upload.blob2.value)) {
                    window.alert("<%=LitFicheroNoExiste%>");
                    return false;
                }
                if (fso.GetFile(marcoFoto.document.upload.blob2.value).Size ><%= cLng(uploadFileLimitj2) %>) {
                    window.alert("<%=LitFicTamGrandeFicArt2%> <%=formatnumber(((uploadFileLimit2-390)/1024),0,-1,0,-1)%><%=LitFicTamGrandeFicArt%>");
                    return false;
                }
            }
            else {
                if (marcoFoto.document.upload.blob2.files[0].size ><%= cLng(uploadFileLimitj2) %>) {
                    alert("<%=LitFicTamGrandeFicArt2%> <%=formatnumber(((uploadFileLimit2-390)/1024),0,-1,0,-1)%><%=LitFicTamGrandeFicArt%>");
                    return false;
                }
            }
        }

        if (marcoFoto.document.upload.blob3.value != "") {
            if (navigator.appName == "Microsoft Internet Explorer") {
                var fso = new ActiveXObject("Scripting.FileSystemObject");
                if (!fso.FileExists(marcoFoto.document.upload.blob3.value)) {
                    window.alert("<%=LitFicheroNoExiste%>");
                    return false;
                }
                if (fso.GetFile(marcoFoto.document.upload.blob3.value).Size ><%= cLng(uploadFileLimitj2) %>) {
                    window.alert("<%=LitFicTamGrandeFicArt3%> <%=formatnumber(((uploadFileLimit2-390)/1024),0,-1,0,-1)%><%=LitFicTamGrandeFicArt%>");
                    return false;
                }
            }
            else {
                if (marcoFoto.document.upload.blob3.files[0].size ><%= cLng(uploadFileLimitj2) %>) {
                    alert("<%=LitFicTamGrandeFicArt3%> <%=formatnumber(((uploadFileLimit2-390)/1024),0,-1,0,-1)%><%=LitFicTamGrandeFicArt%>");
                    return false;
                }
            }
        }

        //  GPD (27/02/2007).
        if (isNaN(f.ue.value) || f.ue.value == "") {
            window.alert("<%=LitMsgUnidadEmbalajeNumerico%>");
            return false;
        }
        //  DBS (29/11/2013).
        if (isNaN(f.uv.value) || f.uv.value == "") {
            window.alert("<%=LitMsgUnidadVentaNumerico%>");
            return false;
        }

        if (f.fbaja.value != "" && !checkdate(f.fbaja)) {
            window.alert("<%=LitMsgFechaBajaFecha%>");
            return false;
        }

        f.spvp.value = f.spvp.value.replace(".", ",");
        f.descuento.value = f.descuento.value.replace(".", ",");
        if (si_tiene_modulo_comercial != 0) {
            f.porcom.value = f.porcom.value.replace(".", ",");
        }
    //f.meses.value=f.meses.value.replace(".",",");
    <%if si_tiene_modulo_importaciones = 0 then %>
        if (document.articulos.si_campo_personalizables.value == 1) {
            num_campos = marcoCamposPersonalizables.document.CamposPersonalizablesArt.num_campos.value;

            respuesta = comprobarCampPerso("marcoCamposPersonalizables.", num_campos, "CamposPersonalizablesArt");
            if (respuesta != 0) {
                titulo = "titulo_campo" + respuesta;
                tipo = "tipo_campo" + respuesta;
                titulo = marcoCamposPersonalizables.document.CamposPersonalizablesArt.elements[titulo].value;
                tipo = marcoCamposPersonalizables.document.CamposPersonalizablesArt.elements[tipo].value;
                if (tipo == 4) {
                    nomTipo = "<%=LitTipoNumericoArt%>";
                }
                else if (tipo == 5) {
                    nomTipo = "<%=LitTipoFechaArt%>";
                }

                window.alert("<%=LitMsgCampoArt%> " + titulo + " <%=LitMsgTipoArt%> " + nomTipo);

                return false;
            }
        }
    <% end if%>

        //ricardo 12-11-2007 si tiene valor el parametro ne , no se puede dejar el codigo de barras a nulo
        if (document.articulos.ne.value != "") {
                if (f.codbarras.value == "") {
                    window.alert("<%=LitMsgNoCodBarras%>");
                    return false;
                }
            }

        //ricardo 9-12-2004 se pone aqui , ya que con un modem o rdsi no se guardaban las fotos
        if (marcoFoto.document.upload.blob1.value != "" || marcoFoto.document.upload.pepe1.value == "1" ||
            marcoFoto.document.upload.blob2.value != "" || marcoFoto.document.upload.pepe2.value == "1" ||
            marcoFoto.document.upload.blob3.value != "" || marcoFoto.document.upload.pepe3.value == "1" ||
            marcoFoto.document.upload.litdocti1.value != "" || marcoFoto.document.upload.litdocti2.value != "" ||
            marcoFoto.document.upload.linkweb.value != "") {
            marcoFoto.document.upload.action = "sube_art.asp";
            marcoFoto.document.upload.submit();
        }
        else {
            //si no hay cambios de fotos, no tengo que esperar nada
            esperar_grabacion_foto = 1;
        }
        ///////////////////////////////////

        if (marcoPropiedades.document.PropiedadesArticulo.pvpDependientes.value != "" && marcoPropiedades.document.PropiedadesArticulo.bloquearprecios.checked) {
            if (isNaN(marcoPropiedades.document.PropiedadesArticulo.pvpDependientes.value.replace(",", ".")) || marcoPropiedades.document.PropiedadesArticulo.pvpDependientes.value < 0) {
                window.alert("<%=LitMsgPvPDependientesNumerico%>");
                return false;
            }
            if (!confirm("<%=LitMsgConfirmCambioDependientesa%>" + marcoPropiedades.document.PropiedadesArticulo.nombre.value + "<%=LitMsgConfirmCambioDependientesb%>")) {
                return false;
            }
        }

        grabarImg(mode, ref);
    }

    function grabarImg(mode, ref) {
        if (esperar_grabacion_foto == 0 && cuanto_espero >= 0) {
            var t = setTimeout("grabarImg('" + mode + "','" + ref + "')", 1000);
        }
        else {
            grabarFichArt(mode, ref);

            marcoDatosContaDeArticulo.document.DatosContaDeArticulo.action = "DatosContaDeArticulo.asp?mode=save&referencia=" + ref;
            marcoDatosContaDeArticulo.document.DatosContaDeArticulo.submit();

            h_has_manufacturer_item = document.articulos.h_has_manufacturer_item.value;
            if (h_has_manufacturer_item == 10000) {
                //    marcoDatosFabDeArticulo.document.DatosFabDeArticulo.action="DatosFabDeArticulo.asp?mode=save&referencia=" + ref;
                //    marcoDatosFabDeArticulo.document.DatosFabDeArticulo.submit();
            }
        }
        cuanto_espero--;
    }

    function grabarFichArt(mode, ref) {
        f = marcoPropiedades.document.PropiedadesArticulo;

        if (f.fbaja.value != "" && checkdate(f.fbaja)) {
            if (window.confirm("<%=LitMsgBorrarDeCatalogo%>") == true) {
                marcoPropiedades.document.PropiedadesArticulo.action = "PropiedadesArticulo.asp?mode=" + mode + "&fbaja=" + f.fbaja + "&referencia=" + ref + "&t=<%=enc.EncodeForJavascript(tarifa)%>";
                marcoPropiedades.document.PropiedadesArticulo.submit();
            }
            else {
                marcoPropiedades.document.PropiedadesArticulo.action = "PropiedadesArticulo.asp?mode=" + mode + "&referencia=" + ref + "&t=<%=enc.EncodeForJavascript(tarifa)%>";
                marcoPropiedades.document.PropiedadesArticulo.submit();
            }
        }
        else {
            marcoPropiedades.document.PropiedadesArticulo.action = "PropiedadesArticulo.asp?mode=" + mode + "&referencia=" + ref + "&t=<%=enc.EncodeForJavascript(tarifa)%>";
            marcoPropiedades.document.PropiedadesArticulo.submit();
        }

        <%if si_tiene_modulo_importaciones = 0 then %>
            if (document.articulos.si_campo_personalizables.value == 1) {
            //ricardo 24-3-2004 se graban los datos de los campos personalizados
            marcoCamposPersonalizables.document.CamposPersonalizablesArt.action = "CamposPersonalizablesArt.asp?mode=" + mode + "&referencia=" + ref;
            marcoCamposPersonalizables.document.CamposPersonalizablesArt.submit();
        }

        <% end if%>
        }

    //***************************************************************************************************************
    function GestionPrecios(mode, ref, es_padre) {
        if (mode == "selectC") {
            mode = "select";
            reftmp = document.articulos.i_hijotcp.value;
            if (reftmp != "") {
                ref = reftmp;
                es_padreGlobal = 0;
            }
            else es_padreGlobal = -1;
        }
        if (mode == "delete") {
            if (es_padreGlobal != 0) msg = "<%=LitMsgCambiosTallColConfirm%>";
            else msg = "<%=LitMsgEliminarPrecios%>";
            if (!confirm(msg)) return false;
        }
        if (es_padreGlobal != 0 && mode == "save") {
            if (!confirm("<%=LitMsgCambiosTallColConfirm%>")) return false;
        }
        if (reftmp != "") ref = reftmp;
        marcoPreciosArticulo.document.PreciosDeArticulo.action = "PreciosDeArticulo.asp?mode=" + mode + "&referencia=" + ref + "&espadre=" + es_padreGlobal;
        marcoPreciosArticulo.document.PreciosDeArticulo.submit();
    }

    //***************************************************************************************************************
    function seleccionar(marco, formulario, check) {
        nregistros = eval(marco + ".document." + formulario + ".hNregs.value-1");
        if (eval("document.articulos." + check + ".checked")) {
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
            }
        }
        else {
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
            }
        }
    }
    //***************************************************************************************************************
    function TraerProveedor(ref) {
        document.articulos.i_razon_social.value = "";
        document.articulos.i_su_ref.value = "";
        document.articulos.i_tipogar.value = "";
        document.articulos.i_divisa.value = "";
        document.articulos.i_importe.value = "0";
        document.articulos.i_descuento.value = 0;
        document.articulos.i_descuento2.value = 0;
        document.articulos.i_portes.value = 0;
        document.articulos.i_embalaje.value = 0;
        document.articulos.i_pvd.value = 0;
        document.articulos.i_mgarpro.value = 0;
        document.articulos.i_cod_barras.value = 0;

        document.location.href = "articulos.asp?nproveedor=" + document.articulos.i_proveedor.value + "&mode=traerprov&referencia=" + ref + "&t=<%=enc.EncodeForJavascript(tarifa)%>";
    }
    //***************************************************************************************************************
    function CalculaImporte(ndecimales) {

        strimporte = document.articulos.importe.value.replace(",", ".");
        nimporte = parseFloat(strimporte)

        if (isNaN(nimporte)) {
            window.alert("<%=LitMsgImporteNumerico%>");
            document.articulos.importe.value = "0"
            return false;
        }
        if (nimporte < 0) {
            window.alert("<%=LitMsgPvPNoNegativo%>");
            document.articulos.importe.value = "0"
            return false;
        }
        strrecargo = document.articulos.recargo.value.replace(",", ".");
        nrecargo = parseFloat(strrecargo);
        pelas_recargo = (nimporte * nrecargo) / 100;

        if (isNaN(pelas_recargo)) {
            window.alert("<%=LitMsgRecargoNumerico%>");
            document.articulos.recargo.value = "0"
            return false;
        }

        pvp = nimporte + pelas_recargo;
        strpvp = pvp.toString();
        npvp = strpvp.replace(".", ",");

        document.articulos.importe.value = strimporte.replace(".", ",")
        document.articulos.recargo.value = strrecargo.replace(".", ",")
        document.articulos.pvp.value = npvp
        document.articulos.spvp.value = npvp

        return true;
    }

    //***************************************************************************************************************
    function VentanaPreciosTyC(art_padre) {
        pagina = "../central.asp?pag1=productos/preciosTYC.asp&ndoc=" + art_padre + "&viene=articulos&ndocumento=&tdocumento=&mode=browse&pag2=productos/preciosTYC_bt.asp&titulo=<%=LITPREART%>";
        ven = AbrirVentana(pagina, 'P',<%=altoventana %>,<%=anchoventana %>);
    }


    function IgualRefProv() {
        if (document.articulos.h_refpro.value == "SI") {
            f = parent.pantalla.marcoPropiedades.document.PropiedadesArticulo
            f.action = "articulos.asp?mode=first_save&referencia=" + f.referencia.value + "&h_refpro=SI&t=<%=enc.EncodeForJavascript(tarifa)%>";
            f.submit();
        }
    }

    function ver_imagen(modo) {
        marcoFoto.document.location = "Articulos_Imagen.asp?mode=" + modo + "&referencia=" + document.articulos.hreferencia.value + "&mf=" + document.articulos.hmostrar_foto.value;
    }


    function GestionarTamPropiedades(objMenu, objImage) {
        f = marcoPropiedades.document;
        if (eval("f.getElementById('" + objMenu.id + "2').style.display") == "none") {
            eval("f.getElementById('" + objMenu.id + "2').style.display='';");
            objImage.src = "../Images/CarpetaAbierta.gif";
        }
        else {
            eval("f.getElementById('" + objMenu.id + "2').style.display='none';");
            objImage.src = "../Images/CarpetaCerrada.gif";
        }

        altoPRO2 = 0;
        altoDCONT2 = 0;
        altoMD2 = 0;

        var alto = 0;
        if (parent.document.body.offsetHeight) alto = parent.document.body.offsetHeight;
        else alto = parent.self.innerHeight;

        if (objMenu.id == "addDCONT" || objMenu.id == "addPRO" || objMenu.id == "addMD") {
            if (eval("f.getElementById('addDCONT2').style.display") == "" || eval("f.getElementById('addPRO2').style.display") == "" || eval("f.getElementById('addMD2').style.display") == "") {
                addPRO.style.display = "";

                if (eval("f.getElementById('addDCONT2').style.display") == "")
                    altoDCONT2 = f.getElementById('addDCONT2').offsetHeight;
                else
                    altoDCONT2 = 0;
                if (eval("f.getElementById('addPRO2').style.display") == "")
                    altoPRO2 = f.getElementById('addPRO2').offsetHeight;
                else
                    altoPRO2 = 0;
                if (eval("f.getElementById('addMD2').style.display") == "")
                    altoMD2 = f.getElementById('addMD2').offsetHeight;
                else
                    altoMD2 = 0;

                document.getElementById("frPropiedades").style.height = altoPRO2 + altoDCONT2 + altoMD2 + 30;
                addPRO.style.height = altoPRO2 + altoDCONT2 + altoMD2 + 30;
            }
            else addPRO.style.display = "none";
        }
        else {
            if (objMenu.id == "BrowseDCONT" || objMenu.id == "BrowsePRO" || objMenu.id == "BrowseMD") {
                if (eval("f.getElementById('BrowseDCONT2').style.display") == "" || eval("f.getElementById('BrowsePRO2').style.display") == "" || eval("f.getElementById('BrowseMD2').style.display") == "") {
                    BrowsePRO.style.display = "";
                    sumara = 0;

                    if (eval("f.getElementById('BrowseDCONT2').style.display") == "")
                        altoDCONT2 = f.getElementById('BrowseDCONT2').offsetHeight;
                    else
                        altoDCONT2 = 0;
                    if (eval("f.getElementById('BrowsePRO2').style.display") == "")
                        altoPRO2 = f.getElementById('BrowsePRO2').offsetHeight;
                    else
                        altoPRO2 = 0;
                    if (eval("f.getElementById('BrowseMD2').style.display") == "")
                        altoMD2 = f.getElementById('BrowseMD2').offsetHeight;
                    else
                        altoMD2 = 0;

                    document.getElementById("frPropiedades").style.height = altoPRO2 + altoDCONT2 + altoMD2 + 30;//"60";
                    BrowsePRO.style.height = altoPRO2 + altoDCONT2 + altoMD2 + 30;
                }
                else {
                    BrowsePRO.style.display = "none";
                }

            }
            else {
                if (objMenu.id == "editDCONT" || objMenu.id == "editPRO" || objMenu.id == "editMD") {
                    if (eval("f.getElementById('editDCONT2').style.display") == "" || eval("f.getElementById('editPRO2').style.display") == "" || eval("f.getElementById('editMD2').style.display") == "") {
                        editPRO.style.display = "";

                        if (eval("f.getElementById('editDCONT2').style.display") == "")
                            altoDCONT2 = f.getElementById('editDCONT2').offsetHeight;
                        else
                            altoDCONT2 = 0;
                        if (eval("f.getElementById('editPRO2').style.display") == "")
                            altoPRO2 = f.getElementById('editPRO2').offsetHeight;
                        else
                            altoPRO2 = 0;
                        if (eval("f.getElementById('editMD2').style.display") == "")
                            altoMD2 = f.getElementById('editMD2').offsetHeight;
                        else
                            altoMD2 = 0;

                        document.getElementById("frPropiedades").style.height = altoPRO2 + altoDCONT2 + altoMD2 + 30;
                        editPRO.style.height = altoPRO2 + altoDCONT2 + altoMD2 + 30;
                    }
                    else editPRO.style.display = "none";
                }
            }
        }
    }


    function muestrapreciofinal(tipo, valor, ndec) {
        switch (tipo) {
            case "P":
                precio2Aux = parseFloat(document.articulos.precio.value.replace(",", "."));
                precio2Aux = precio2Aux.toFixed(<%=DEC_PREC %>);

                //precio2=parseFloat(precio2Aux.replace(/[.]/g,"").replace(",",".") )
                precio2 = parseFloat(precio2Aux.replace(",", "."));

                iva = parseFloat(document.articulos.hiva.value.replace(",", "."));
                pvpiva = precio2 + ((precio2 * iva) / 100)
                document.articulos.precio.value = precio2.toString().replace(",", ".");
                document.articulos.precioiva.value = pvpiva.toFixed(ndec);//.replace(/[.]/g,"")
                document.articulos.preciofinal.value = document.articulos.precio.value;//.replace(".",",");
                document.articulos.descuento.value = "";
                document.articulos.descuentocoste.value = "";
                break;
            case "D":
                if (isNaN(document.articulos.descuento.value.replace(",", "."))) {
                    window.alert("<%=LitMsgDescuentoNumerico%>");
                    return false;
                }
                fvalor = parseFloat(valor.replace(",", "."));
                fdto = parseFloat(document.articulos.descuento.value.replace(",", "."));
                fvalor = fvalor + ((fvalor * fdto) / 100);
                iva = parseFloat(document.articulos.hiva.value.replace(",", "."));
                pvpiva = fvalor + ((fvalor * iva) / 100);
                pvpiva = pvpiva.toFixed(ndec);
                document.articulos.precioiva.value = pvpiva;//.replace(/[.]/g,"")
                fvalor = fvalor.toFixed(<%=DEC_PREC %>);
                document.articulos.preciofinal.value = fvalor;//.toString().replace(".",",");
                document.articulos.descuentocoste.value = "";
                document.articulos.precio.value = "";
                break;
            case "DC":
                if (isNaN(document.articulos.descuentocoste.value.replace(",", "."))) {
                    window.alert("<%=LitMsgDescuentoNumerico%>");
                    return false;
                }
                fvalor = parseFloat(valor.replace(",", "."));
                fdto = parseFloat(document.articulos.descuentocoste.value.replace(",", "."));
                fvalor = fvalor + ((fvalor * fdto) / 100);
                iva = document.articulos.hiva.value;
                pvpiva = fvalor + ((fvalor * iva) / 100);
                document.articulos.precioiva.value = pvpiva.toFixed(ndec);//.replace(/[.]/g,"")
                fvalor = fvalor.toFixed(<%=DEC_PREC %>);
                document.articulos.preciofinal.value = fvalor;//.toString().replace(".",",");
                document.articulos.descuento.value = "";
                document.articulos.precio.value = "";
                break;
            case "PI":
                campopvpiva = document.articulos.precioiva.value
                iva = document.articulos.hiva.value
                if (campopvpiva != "" && !isNaN(campopvpiva.replace(",", "."))) {
                    pvpiva = parseFloat(campopvpiva.replace(",", "."));
                    pvpsiniva = pvpiva / parseFloat(1 + (iva / 100));
                    pvpsiniva = pvpsiniva.toFixed(<%=DEC_PREC %>);
                    document.articulos.precioiva.value = pvpiva;//.toString().replace(/[.]/g,'');
                    document.articulos.precio.value = pvpsiniva;//.toString().replace(/[.]/g,'');
                    document.articulos.preciofinal.value = pvpsiniva;
                    document.articulos.descuento.value = "";
                    document.articulos.descuentocoste.value = "";
                    muestrapreciofinal('P', '', ndec);
                }
                break;
        }
        //RGU 17/1/2007
        pvpfinal = parseFloat(document.articulos.preciofinal.value.replace(",", "."));//.toString().replace(",",".");
        impO = parseFloat(document.articulos.IO.value.replace(",", "."));//.toString().replace(".","").replace(",",".");
        pvpO = parseFloat(document.articulos.PO.value.replace(",", "."));//.toString().replace(".","").replace(",",".");
        if (impO != 0) {
            icost = ((pvpfinal - impO) * 100) / impO;
            icost = icost.toFixed(<%=decpor %>);
            document.getElementById("incimp").innerHTML = icost.toString();
        }
        if (pvpO != 0) {
            ipvp = ((pvpfinal - pvpO) * 100) / pvpO;
            ipvp = ipvp.toFixed(<%=decpor %>);
            document.getElementById("incpvp").innerHTML = ipvp.toString();
        }

        //RGU 17/1/2007
    }

    function PrecioBloqueado() {
        alert("<%=LitPrecioBloqueado%>");
    }

    //PBG 28/5/2007 Para ocultar artículos dados de baja
    function CambiarValorCheckOcultarArtBaja() {
        var valor, marcado;
        marcado = document.getElementById('chkBoxOcultarArtBaja').checked;

        if (marcado) valor = 1;
        else valor = 0;

        document.getElementById('ocultarArticulosBaja').value = valor;
    }

    function OpenNonPayment() {
        esconde();
    }

    function SetChangePrice(_nombreFrame, _ndoc) {
        //reloadClass(_referencia,"../central.asp?pag1=productos/ChangePricesHistory.asp?mode=browse"; //&ndoc=" + _ndoc + "&ncliente=" + _cust + "&tdocumento=" + _typedoc + "&pag2=administracion/nonpayment_bt.asp" );			                    
        //alPresionar(_referencia);  

        var timestamp = Number(new Date());
        cambiarTamanyo(_nombreFrame, 450, 800);
        reloadClass(_nombreFrame, "../central.asp?pag1=productos/ChangeProductsHistory.asp&ndoc=" + _ndoc + "&mode=browse&ts=" + timestamp + "&pag2=./ventas/changeOrderState_bt.asp");
        alPresionar(_nombreFrame);

    }



</script>
<%'Everilion Interface Timing%>
<script language="javascript" type="text/javascript" src="/lib/js/InterfaceLoadTime.js"></script>
<script language="javascript" type="text/javascript">
    <% mode=request.querystring("mode") %>
        window.onload = function() {
        //osm 28/04/15 bloquear y poner valores por defecto de provincia en modo edit y add
        <%if si_tiene_modulo_intelitur <> 0 and(mode = "edit" or mode = "add") then %>
            var selects = parent.pantalla.marcoCamposPersonalizables.document.getElementsByName("campo2");
            <% if idProvinciaUser & ""= "" or isnull(idProvinciaUser) then %>
            selects[0].disabled = true
                <%else%>
                <%if idProvinciaUser <> 53 then %>
            selects[0].disabled = true
                <%if mode = "add" then %>
                    selects[0].value = <%=idProvinciaUser %>
                    <% end if %>
                <% end if %>
            <% end if %>
        <% end if %>



        <%if mode = "browse" then %>
        //
        <% elseif mode = "first_save" then %>
            IgualRefProv();
        <%else%>
            self.status='';
        <% end if%>
        <%if tracetime > 0 then %>
                StoreTiming("<%=CarpetaProduccion%>", <%=tracetime %>, "<%=enc.EncodeForJavascript(Request.QueryString("mode"))&""%>", "<%=session("usuario")%>", "<%=session("ncliente")%>", window.location.pathname);
        <% end if %>
        }

</script>

<body class="BODY_ASP">
<%
'PBG 28/05/2007 para ocultar artículos dados de baja
ocultarArticulosBaja = limpiaCadena(Request.Form("ocultarArticulosBaja"))

if ocultarArticulosBaja = "" then
    ocultarArticulosBaja = d_lookup("OCULTARARTBAJA","CONFIGURACION"," nempresa = '" & session("ncliente") & "'" , session("dsn_cliente"))
end if

function CerrarTodo()
	set rst = NOTHING
	set rst2 = NOTHING
	set rstAux = NOTHING
	set rstAux2 = NOTHING
    set rstAux3 = NOTHING
	set rst_almacenar = NOTHING
	set rst_proveer = NOTHING
	set rst_escandallo = NOTHING
	set rsttallas = NOTHING
	set rstcolores = NOTHING
end function

'***********************************************************************************************************
'Genera una referencia con contador de configuración
function AutoRef()
	dim RefLibre

	conn.open session("dsn_cliente")
	conn.CommandTimeout = 0

	''ricardo 12-10-2007 si el parametro ne viene con valor se buscara el contador en la empresa de dicho valor
	if ne & "">"" then
	    set rs = conn.execute("EXEC AutoReferencia @p_nempresa='" & ne & "'")
	else
	    set rs = conn.execute("EXEC AutoReferencia @p_nempresa='" & session("ncliente") & "'")
	end if

	RefLibre=rs(0)
	set rs = nothing
	conn.close
	set conn = nothing

	AutoRef=RefLibre
end function

'****************************************************************************************************************
''ricardo 12-11-2007 si el parametro ne tiene valor se comprobara el codigo de barras en la otra empresa
function ComprobarCodBarrasOtraEmpresa(codbarrasOtraEmpresa,referencia)
	dim CodBarrasLibre

	set conn=Server.CreateObject("ADODB.Connection")
	conn.open session("dsn_cliente")
	conn.CommandTimeout = 0
	
	strSelCBL="EXEC ComprobarCodBarrasOtraEmpresa @referencia='" & referencia & "',@codbarras='" & codbarrasOtraEmpresa & "',@p_nempresa='" & ne & "'"
    set rs = conn.execute(strSelCBL)
	CodBarrasLibre=rs(0)
	set rs = nothing
	conn.close
	set conn = nothing

	ComprobarCodBarrasOtraEmpresa=CodBarrasLibre
end function

''ricardo 12-11-2007 si el parametro ne tiene valor se creara el articulo en la otra empresa
function GuardarArtEnOtraEmpresa(referencia,empresa,otraempresa)
	dim msgerror

	set conn=Server.CreateObject("ADODB.Connection")
	conn.open session("dsn_cliente")
	conn.CommandTimeout = 0
	
	strSelGCI="EXEC GuardarArtEnOtraEmpresa @referencia='" & referencia & "',@nempresa='" & empresa & "',@Otraempresa='" & otraempresa & "'"
    set rs = conn.execute(strSelGCI)
	msgerror=rs(0)
	set rs = nothing
	conn.close
	set conn = nothing

	GuardarArtEnOtraEmpresa=msgerror
end function

'*** i AMP
sub GuardarHistorialCambios(ref,pvpAnt,pvp,mode)
      
     '' comprobar si el usuario esta dado de alta como personal.
    ''existe=d_lookup("login","personal","login='" & session("usuario") &  "' and dni like  '" & session("ncliente") & "%' ",session("dsn_cliente"))&""   
    existe=d_lookup("dni","personal","dni='" & session("ncliente") & session("usuario") & "'",session("dsn_cliente"))&""
    if existe>"" then  
        if mode="first_save" then
            
                strEdit="EXEC SetProductsHistory @ncompany='" & session("ncliente") & "',@reference='" & session("ncliente") & ref & "'" &_
                    ",@user='" &  session("usuario") & "'" &_
                    ",@action='" & LITADDPRODUCT & "'" &_
                    ",@text='" & LITADDPRODUCT & "'" &_
                    ",@beforeValue=''" &_ 
                    ",@afterValue=''"         
              
   	            set connActPrice = Server.CreateObject("ADODB.Connection")
	            connActPrice.open session("dsn_cliente")                     
                set rstActPrice = connActPrice.execute(strEdit)
                if not rstActPrice.eof then
                        result = rstActPrice("result")
                        'if result=0 then
                            %> <!--<script type="text/javascript" language="javascript">alert("OK");</script> --><% 
                        'end if
                end if		                           
                connActPrice.close 
                set rstActPrice = Nothing
                set connActPrice = Nothing   
        else
            if pvpAnt<>pvp then             
                strEdit="EXEC SetProductsHistory @ncompany='" & session("ncliente") & "',@reference='" & session("ncliente") & ref & "'" &_
                    ",@user='" & session("usuario") & "'" &_
                    ",@action='" & LITPRICECHANGE & "'" &_
                    ",@text='" & LITTEXTPCHANGE & "'" &_
                    ",@beforeValue='" & pvpAnt & "'" &_ 
                    ",@afterValue='" & pvp & "'"         
        	     
    	         set connActPrice = Server.CreateObject("ADODB.Connection")
	             connActPrice.open session("dsn_cliente")                        	          
                 set rstActPrice = connActPrice.execute(strEdit)
                 if not rstActPrice.eof then
                        result = rstActPrice("result")
                        'if result=0 then
                            %> <!--<script type="text/javascript" language="javascript">alert("OK");</script> --><% 
                        'end if
                 end if		                       
                 connActPrice.close 
                 set rstActPrice = Nothing
                 set connActPrice = Nothing              
            end if
        end if
   else		   
	    %><script language="javascript" type="text/javascript">alert("<%=LITNOEXISTUSER%>")</script><%
    end if
end sub
'*** f AMP


'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(p_referencia, talla, color, miembro)
    rst.cursorlocation=2
    rst.Open "select * from articulos where referencia='" & session("ncliente")&p_referencia & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    nuevo=0
	if rst.eof then
	   rst.addnew
	   nuevo=1
	end if
	if mode="first_save" then
		if session("ncliente")&p_referencia = rst("referencia") then
	   		TmpRef=rst("referencia")
	   		guarda = false
	   		rst.cancelupdate
	   		rst.close%>
	   		<script type="text/javascript" language="javascript">
                   window.alert("<%=LitMsgReferenciaExiste%>");
                   parent.botones.document.location = "articulos_bt.asp?mode=browse&t=<%=enc.EncodeForJavascript(tarifa)%>";
	   		</script>
   		<%else
	   		guarda = true
		end if
	else
		guarda=true
	end if

   
	if request.form("codbarras")>"" then
		codbarras=request.form("codbarras")
	elseif request.form("prefcodbarras")&""<>""then
		codbarras=genera_codbarras(request.form("prefcodbarras"),p_referencia)
	end if
    
	if (mode="first_save" or mode="save") and guarda=true then
		if codbarras>"" then
			if isnumeric(trim(codbarras)) then
				rstAux.Open "select referencia from articulos with(nolock) where referencia='" & session("ncliente")&codbarras & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if rstAux.eof then
					rstAux.close
					rstAux.Open "select referencia from articulos with(nolock) where referencia like '" & session("ncliente") & "%' and cod_barras='" & codbarras & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					if rstAux.eof then
		   				rstAux.close
						rstAux.open "select referencia from codigos_barras with(nolock) where referencia like '" & session("ncliente") & "%' and cod_barras='" & codbarras & "'", session("dsn_cliente")
						if rstAux.eof then
							rstAux.close
							guarda=true
						else
							msg=LitMsgCodBarrasExiste & " (" & trimCodEmpresa(rstAux("referencia")) & ")"
							rstAux.close
							guarda=false
						end if
					else
						msg=LitMsgCodBarrasExiste & " (" & trimCodEmpresa(rstAux("referencia")) & ")"
						rstAux.close
						if cdbl(d_lookup("ncodbarras","configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente")))=cdbl(mid(codbarras,3,10)) then
							msg=msg & LITCONTCOBBARRART
						end if
						guarda=false
					end if
				else
					rstAux.close
					msg=LITREFIGUCODBARART
					guarda=false
				end if
				
                ''ricardo 12-11-2007 si el parametro ne tiene valor se comprobara el codigo de barras en la otra empresa
                codbarrasenotraempresaexit=""
                if ne & "">"" then
                    codbarrasenotraempresaexit=ComprobarCodBarrasOtraEmpresa(codbarras,p_referencia)
                end if
                if codbarrasenotraempresaexit & "">"" then
                    msg=LitMsgCodBarrasExiste
                    guarda=false
                end if
                
			else
				msg=LITCODBARNUMART
				guarda=false
			end if
		end if
	end if
	'guardamos la url en la tabla urlrewrite_eshop
	if urlRewrite=True and si_tiene_modulo_ecomerce<>0 and request.Form("impr_cat")&""="true" then
	    dsnMixta = ObtenDSNMixta(session("dsn_cliente"),DSNILION)
	    cadena= "exec URLREWITEALTA 'articulo','" &session("ncliente")&p_referencia& "','" &request.Form("url_mostrar")& "','"& session("ncliente")&"'"
        rstAux2.Open cadena,dsnMixta                  
        if not rstAux2.eof then
            if rstAux2("result")=-1 then
                guarda=false
                nuevo=0
                rst.cancelupdate
		        rst.close%>
                <script type="text/javascript" language="javascript">
                    window.alert('<%=LitUrlExistente%>' + '<%=rstAux2("tipo")%>' + " " + '<%=rstAux2("nombre")%>');
                    parent.botones.document.location = "articulos_bt.asp?mode=add&t=<%=enc.EncodeForJavascript(tarifa)%>"
                </script>
            <%
            end if
        end if        
	    rstAux2.Close
	end if
	if ((guarda=false) and nuevo=1) then
		rst.cancelupdate
		rst.close%>
	   	<script type="text/javascript" language="javascript">
               alert("<%=msg%>");
               parent.botones.document.location = "articulos_bt.asp?mode=add&t=<%=enc.EncodeForJavascript(tarifa)%>";
   		</script>
   		<%mode="browse"
	end if

	if guarda = true then

		ndecimales = d_lookup("ndecimales", "divisas", "codigo='" & request.form("divisa") & "'", session("dsn_cliente"))
		NombreArticulo=request.form("nombre")
		if miembro="HIJO" and talla>"" then NombreArticulo=NombreArticulo & " T:" & talla
		if miembro="HIJO" and color>"" then NombreArticulo=NombreArticulo & " C:" & color
		rst("referencia")= Nulear(session("ncliente")&p_referencia)
	 	rst("nombre")	= trim(NombreArticulo)

        'MAP 20/12/2012 - Guardar fecha de creación desde el formulario.
        'response.Write("FECHA CREACION= "+request.Form("h_fcreacion"))
        'response.End
        
      
        rst("fechacreacion")=nulear(replace(replace(Request.Form("h_fcreacion"),"a.m.","am"),"p.m.","pm"))


         'rst("fechacreacion")=request.Form("h_fcreacion")
         'response.Write("Fecha creacion al guardar...:"&request.Form("fechacreacion")&"")
         'response.End
       

	 	rst("importe")	= miround(null_z(request.form("coste")),dec_prec)
	 	rst("recargo")	= miround(null_z(request.form("recargo")),decpor)
		pvpAnt=rst("pvp")
		rst("pvp")		= miround(null_z(request.form("spvp")),dec_prec) 'precioventarecargo(rst("importe"),rst("recargo"))
		rst("margen")   = miround(null_z(request.form("margen")),decpor) 'margenventa(rst("importe"),pvp)
	 	rst("descuento")= miround(null_z(request.form("descuento")),decpor)      

		if pvp&""<>rst("pvp")&""  then
			rst("fechamod")=date
		end if
		pvp = rst("pvp")

		rst("iva")= Nulear(request.form("iva"))
		rst("familia") =  Nulear(request.form("familia"))
		fampadre=""
		categoria=""
		if request.form("familia")&""<>"" then fampadre=d_lookup("padre","familias","codigo='" & request.form("familia") & "'",session("dsn_cliente"))
		if fampadre<>"" then categoria=d_lookup("categoria","familias","codigo='" & request.form("familia") & "'",session("dsn_cliente"))
		rst("familia_padre")=Nulear(fampadre)
		rst("categoria")=Nulear(categoria) 		 
		rst("medida") =  Nulear(request.form("medida"))
        ''ZEK: 19/02/2010
        rst("tipo_conversion") = Nulear(request.form("tipo_conversion"))
		rst("puntos") =  Nulear(request.form("puntos"))
		rst("porcada") =  Nulear(request.form("porcada"))
	' >>> MCA 02/12/04 : Ofrecer los campos <Unidad aux.venta> y <Calcular importe detalle> siempre
	'						con indenpendencia del módulo contratado

		rst("medidaventa") =  Nulear(request.form("medidaventa"))
		if request.form("calculoimporte")="on" then
			rst("calculoimporte") = true
		else
			rst("calculoimporte") = false
		end if

		if 1=2 then		' Para conservar el código anterior sin que se ejecute
			if si_tiene_modulo_produccion<>0 then
				rst("medidaventa") =  Nulear(request.form("medidaventa"))
				if request.form("calculoimporte")="on" then
					rst("calculoimporte") = true
				else
					rst("calculoimporte") = false
				end if
			end if
		end if
       if request.Form("servicio") = "on" then
            rst("tipoproducto") = 6
        else
            rst("tipoproducto") = null
        end if

	' >>> MCA 02/12/04 : Ofrecer los campos <Unidad aux.venta> y <Calcular importe detalle> siempre
	'						con indenpendencia del módulo contratado

		if talla>"" then
			rst("talla") = talla
		else
			rst("talla") =  Nulear(request.form("talla"))
		end if

	 	if color>"" then
	      	rst("color")=color
		else
			rst("color") =  Nulear(request.form("color"))
	 	end if

		rst("modelo")	= request.form("modelo")
		rst("fbaja")	= Nulear(request.form("fbaja"))
		if request.Form("peso")&"" > "" then
            rst("weight") = replace(request.Form("peso"),".",",")
        end if
        if request.form("descatalogado")="on" then
			rst("discontinued") = true
		else
			rst("discontinued") = false
		end if
		rst("tipogar")	   = Nulear(request.form("tipogar"))
		rst("meses")	   = null_z(request.form("meses"))
		rst("tmanoobra")   = null_z(request.form("mesesmo"))
		rst("tdesp")       = null_z(request.form("mesesde"))
		rst("tmateriales") = null_z(request.form("mesesmt"))
		rst("porcom")	 = null_z(request.form("porcom"))

		if request.form("ctrl_nserie")="on" then
			rst("ctrl_nserie") = true
		else
			rst("ctrl_nserie") = false
		end if
request.form("control_stock="&request.form("control_stock")&" calculoimporte="&request.form("calculoimporte"))
		if request.form("control_stock")="on" then
			rst("control_stock") = true
		else
			rst("control_stock") = false
		end if

        lotecompra_old=nz_b(rst("LOTECOMPRA"))
        if lotecompra_old=-1 then lotecompra_old=1
        lotecompra_new=0
        if nz_b(request.form("gestion_lotes")) = (-1) then
            rst("lotecompra") = true
            lotecompra_new=1
        else
            rst("lotecompra") = false
            lotecompra_new=0
        end if

        if request.form("impr_cat")&""="" then
            if request.form("impr_catalogo") = "on" then
			    rst("impr_catalogo") = true
		    else
			    rst("impr_catalogo") = false
		    end if
        else
		    if request.form("impr_cat")="true" then
		        rst("impr_catalogo") = true
		    else
		        rst("impr_catalogo") = false
		    end if
		end if

		'	------------------------------------------------
		'	GPD (26/02/2007) - Añadir campos "novedad" y "Unidad de embalaje"

		if request.form("novedad")="on" then
			rst("novedad") = true
		else
			rst("novedad") = false
		end if

		rst("ue") =  Nulear(request.form("ue"))
        '  DBS (29/11/2013).
        rst("uv") =  Nulear(request.form("uv"))

		'	------------------------------------------------

		rst("tipo_articulo")	 = nulear(request.form("tipo_articulo"))

		rst("observaciones")   = request.form("observaciones")
		rst("cod_barras")      = nulear(codbarras)
		rst("caracteristicas") = nulear(request.form("caracteristicas"))
		rst("divisa")          = request.form("divisa")
		rst("agrtallas")       = nulear(request.form("agrupa_tallas"))
		rst("agrcolores")       = nulear(request.form("agrupa_colores"))
		if request.form("agrupa_colores")&"">"" or request.form("agrupa_tallas")&"">"" then
			rst("ES_PADRE")=true
		end if
		if miembro="PADRE" then
			rst("ES_PADRE")=true
		elseif miembro="HIJO" then
			rst("REF_PADRE")=mid(p_referencia,1,instr(p_referencia,"/")-1)
		end if

		if request.form("bono") = "on" then
			rst("bono") = true
		else
			rst("bono") = false
		end if

        if si_tiene_modulo_OrCU<>0 then  
            if request.form("consigna") = "on" then
			    rst("consigna") = 1
		    else
			    rst("consigna") = 0
		    end if
        end if


		if request.form("carga_terminal") = "on" then
			rst("loadter") = true
		else
			rst("loadter") = false
		end if

		''ricardo 24-3-2004

		dim lista_valores
		num_campos_perso=limpiaCadena(request.querystring("num_campos_perso"))
		if num_campos_perso & "">"" then
			redim lista_valores(num_campos_perso+2)
			for ki=1 to num_campos_perso
				nom_campo_perso="campo" & ki
				valor_form=limpiaCadena(request.querystring(nom_campo_perso))
				if ki<10 then
					tipo_campo_perso=d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "0" & ki & "' and tabla='ARTICULOS'",session("dsn_cliente"))
				else
					tipo_campo_perso=d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & ki & "' and tabla='ARTICULOS'",session("dsn_cliente"))
				end if
				if tipo_campo_perso=2 then
					if ucase(valor_form)="ON" or valor_form="1" then
						valor_form=1
					else
						valor_form=0
					end if
				end if
				num_campo_str=cstr(ki)
				if len(num_campo_str)=1 then
					num_campo_str="0" & num_campo_str
				end if
				nom_campo_perso="campo" & num_campo_str
				rst(nom_campo_perso)=valor_form
			next
			
		end if

        subctaventasF=Nulear(request.form("subctaventas"))
        if subctaventasF & ""="" then
            subctaventasF=Nulear(limpiaCadena(request.QueryString("subctaventas")))
        end if
		rst("subctaventas")=subctaventasF
        ''rst("subctaventas")=Nulear(request.form("subctaventas"))
        subctaabventasF=Nulear(request.form("subctaabventas"))
        if subctaabventasF & ""="" then
            subctaabventasF=Nulear(limpiaCadena(request.QueryString("subctaabventas")))
        end if
		rst("subctaabventas")=subctaabventasF
        ''rst("subctaabventas")=Nulear(request.form("subctaabventas"))

        subctacomprasF=Nulear(request.form("subctacompras"))
        if subctacomprasF & ""="" then
            subctacomprasF=Nulear(limpiaCadena(request.QueryString("subctacompras")))
        end if
		rst("subctacompras")=subctacomprasF
        ''rst("subctacompras")=Nulear(request.form("subctacompras"))

        subctaabcomprasF=Nulear(request.form("subctaabcompras"))
        if subctaabcomprasF & ""="" then
            subctaabcomprasF=Nulear(limpiaCadena(request.QueryString("subctaabcompras")))
        end if
		rst("subctaabcompras")=subctaabcomprasF
        ''rst("subctaabcompras")=Nulear(request.form("subctaabcompras"))

        ''if has_manufacturer_item ="1" then
        ''    pnfF=Nulear(request.form("pnf"))
        ''    if pnfF & ""="" then
        ''        pnfF=Nulear(limpiaCadena(request.QueryString("pnf")))
        ''    end if
        ''    rst("pnf") = pnfF
        ''    nmanufacturerF=Nulear(request.form("nmanufacturer"))
        ''    if nmanufacturerF & ""="" then
        ''        nmanufacturerF=Nulear(limpiaCadena(request.QueryString("nmanufacturer")))
        ''    end if
        ''    rst("nmanufacturer") = nmanufacturerF
        ''end if
		
		'JMMM - 03/01/2011 -> Obtenemos las subctas. de la subfamilia (SÓLO SI NO SE ESPECIFICAN SUBCTAS. CONCRETAS)
		if request.form("familia")&"">"" then
		    strSQL = "select subctaventas,subctaabventas,subctacompras,subctaabcompras from familias with(nolock) where codigo = '"& rst("familia") &"' and codigo like '"& session("ncliente") &"%'"
		    rstAux.Open strSQL, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		    if not rstAux.eof then
		        if Nulear(limpiaCadena(request.QueryString("subctaventas")))&""="" then
                    rst("subctaventas") = rstAux("subctaventas")
                end if
                if Nulear(limpiaCadena(request.QueryString("subctaabventas")))&""="" then
                    rst("subctaabventas") = rstAux("subctaabventas")
                end if
                if Nulear(limpiaCadena(request.QueryString("subctacompras")))&""="" then
                    rst("subctacompras") = rstAux("subctacompras")
                end if
                if Nulear(limpiaCadena(request.QueryString("subctaabcompras")))&""="" then
                    rst("subctaabcompras") = rstAux("subctaabcompras")
                end if
            end if
            rstAux.Close
        end if
		rst("aviso")=Nulear(request.form("aviso"))
		rst("documento")=Nulear(request.form("documento"))

		'i(EJM 08/06/2006)
		''MPC 26/05/2009 Se obtiene el campo horecas de configuración para cambiar el nombre de un campo o dejarlo
		if si_tiene_modulo_tiendas<>0 and horecas <> 0 then
		    if request.form("mesesCaducidad") = "on" then
		        rst("meses") = 1
		    else
		        rst("meses")=0
		    end if
		else
		    ''ricardo 9-6-2009 se cambia el nulear por null_z ya que el campo mes es numerico y daba error al insertar null
		    rst("meses")=null_z(request.form("mesesCaducidad"))
		end if
		if rtrim(request.form("importeabparcial")="") then
			 rst("importeabparcial")=0
		else
			rst("importeabparcial")=replace(Nulear(request.form("importeabparcial")),".",",")
		end if
		'fin(EJM 08/06/2006)

		''ricardo 27-12-2006 se añade los campos para los formatos de etiquetas
		rst("cantidadarticulo")=replace(null_z(request.form("cantidadarticulo")),".",",")
		rst("unidadarticulo")=Nulear(request.form("unidadarticulo"))
    referenciapnf=rst("referencia")

               
		rst.update
		rst.close



        ''Ricardo 24-03-2014 se auditara los cambios en el campo lotecompra
        if lotecompra_old<>lotecompra_new then
            strAuditLC="insert into product_history(ncompany,nproduct,note_date,noted_by,note_action,note_text,modification,beforevalue,aftervalue)"
            strAuditLC=strAuditLC & " values (?,?,?,?,?,?,?,?,?)"
            set connInsA = Server.CreateObject("ADODB.Connection")
            set commandInsA =  Server.CreateObject("ADODB.Command")
            connInsA.open session("dsn_cliente")
            commandInsA.ActiveConnection =connInsA
            commandInsA.CommandTimeout = 0
            commandInsA.CommandText=strAuditLC
            commandInsA.CommandType = adCmdText
            commandInsA.Parameters.Append commandInsA.CreateParameter("@ncompany",adVarChar,adParamInput,5,session("ncliente"))
            commandInsA.Parameters.Append commandInsA.CreateParameter("@nproduct",adVarChar,adParamInput,30,referenciapnf)
            commandInsA.Parameters.Append commandInsA.CreateParameter("@note_date",adDate,adParamInput,,now)
            commandInsA.Parameters.Append commandInsA.CreateParameter("@noted_by",adVarChar,adParamInput,20,session("ncliente") & session("usuario"))
            commandInsA.Parameters.Append commandInsA.CreateParameter("@note_action",adVarChar,adParamInput,250,"CAMBIO ATRIBUTO")
            commandInsA.Parameters.Append commandInsA.CreateParameter("@note_text",adVarChar,adParamInput,250,"CAMBIO MANUAL GESTION DE LOTES")
            commandInsA.Parameters.Append commandInsA.CreateParameter("@modification",adSmallInt,adParamInput,,1)
            commandInsA.Parameters.Append commandInsA.CreateParameter("@beforevalue",adVarChar,adParamInput,50,CSTR(lotecompra_old))
            commandInsA.Parameters.Append commandInsA.CreateParameter("@aftervalue",adVarChar,adParamInput,50,CSTR(lotecompra_new))
            set rstInsA = commandInsA.Execute
            set rstInsA = nothing
            set commandInsA = nothing
            set connInsA = nothing
        end if

	    ''ricardo 23-6-2011 se gestiona la referencia PNF
	    if pnf & "">"" then
''response.Write("el pnf 2 es-" & pnf & "-<br/>")
            num_campoC=replace(pnf,"campo","")
            if num_campoC & "">"" then
                num_campo=cint(num_campoC)
            end if
		    nom_campo="campo" & num_campo
		    valor_form=limpiaCadena(request.querystring(nom_campo))
		    if valor_form & ""="" then
			    valor_form=request.form(nom_campo)
		    end if
''response.Write("el valor_form es-" & valor_form & "-" & referenciapnf & "-<br/>")
''response.end
            
	        strselect="select a.referencia,a.cod_barras,a.descripcion from codigos_barras as a where a.referencia='" & referenciapnf & "' and upper(descripcion)=upper('PNF')"
	        rstAux2.open strselect,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	        if not rstAux2.eof then
	            rstAux2("cod_barras")=valor_form
	            rstAux2.update
	        else
	            if valor_form & "">"" then
	                rstAux2.addnew
	                rstAux2("referencia")=referenciapnf
	                rstAux2("cod_barras")=valor_form
	                rstAux2("descripcion")="PNF"
	                rstAux2.update
	            end if
	        end if
''on error resume next
	        
''if err.number<>0 then
''response.Write(err.Description)
''response.end
''end if
''on error goto 0
	        rstAux2.close
	    end if

		'incrementar el codigo de barras
		if request.form("prefcodbarras")&""="C" then
			increm_codbarras(nulear(codbarras))
		end if

		'guardar en la tabla precios
	    Resultado=0
	    set conn = Server.CreateObject("ADODB.Connection")
	    set command =  Server.CreateObject("ADODB.Command")
	    conn.open session("dsn_cliente")
	    command.ActiveConnection =conn
	    command.CommandTimeout = 0
	    command.CommandText="InsUpdArticlePrices"
	    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	    command.Parameters.Append command.CreateParameter("@ncompany",adVarChar,adParamInput,5,session("ncliente"))
	    Command.Parameters.Append Command.CreateParameter("@reference", adVarChar, adParamInput,30,session("ncliente") & p_referencia)
        Command.Parameters.Append Command.CreateParameter("@pvp", adDouble , adParamInput,10, pvp)
        Command.Parameters.Append Command.CreateParameter("@result", adInteger, adParamOutput, Resultado)
	    Command.Execute,,adExecuteNoRecords
	    Resultado = Command.Parameters("@result").Value
        ''response.write("el Resultado es-" & Resultado & "-<br/>")
        if Resultado<0 then
		    %><script language="javascript" type="text/javascript">
                  window.alert("<%=LITOPERFALLFICART%>");
		    </script><%
        end if
	    conn.close
	    set command=nothing
	    set conn=nothing
	        
		'guardado del almacén--------------------------------------------------------------------------------------

		''ricardo 2-8-2004 si en datos de configuracion pone que ART_CREARARTTODALM=1 se creara el articulo en todos los almacenes
	   	ART_CREARARTTODALM=nz_b(d_lookup("ART_CREARARTTODALM","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))
		if ART_CREARARTTODALM<>0 then
		   	almacenDefecto=d_lookup("almacen","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
		   	rst.Open "select codigo from almacenes where codigo like '" & session("ncliente") & "%' and fbaja is null",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			while not rst.eof
			   	rst2.Open "select * from almacenar where articulo like '" & session("ncliente") & "%' and articulo='" & p_referencia & "' and almacen='" & rst("codigo") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		   		if rst2.EOF then
		   			rst2.addnew
				 	rst2("articulo")=session("ncliente")&p_referencia
					rst2("almacen")=rst("codigo")
					rst2("stock") = 0
					rst2("stock_minimo") = 0
					rst2("reposicion") = 0
					rst2("p_recibir") = 0
					rst2("p_servir") = 0
					rst2("p_min") = 0
					if rst("codigo")=almacenDefecto then
						rst2("predet")=1
					else
						rst2("predet")=0
					end if
					rst2("reserva")=0
					rst2("pteabonocanje")=0
					rst2("sat")=0
					rst2.update
				else
					if rst("codigo")=almacenDefecto then
						rst2("predet")=1
					else
						rst2("predet")=0
					end if
					rst2.update
				end if
				rst2.close
				rst.movenext
			wend
			rst.close
		else
		   	almacen=d_lookup("almacen","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
		   	rst.Open "select * from almacenar where articulo like '" & session("ncliente") & "%' and articulo='" & p_referencia & "' and almacen='" & almacen & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	   		if rst.EOF then
	   			rst.addnew
			 	rst("articulo")=session("ncliente")&p_referencia
				rst("almacen")=almacen
				rst("stock") = 0
				rst("stock_minimo") = 0
				rst("reposicion") = 0
				rst("p_recibir") = 0
				rst("p_servir") = 0
				rst("p_min") = 0
				rst("predet")=true
				rst("reserva")=0
				rst("pteabonocanje")=0
				rst("sat")=0
				rst.update
			else
				rst("predet")=true
				rst.update
			end if
			rst.close
		end if

		'---------------------------------------------------------------------------------------------------------
		'Si el usuario pertenece a varias empresas y tiene activa la opción correspondiente,
		'el artículo se da de alta en el almacén por defecto de cada una de las empresas
		msg=""
		rst.open "select ncliente from clientes_users where usuario='" & session("usuario") & "' and fbaja is null",DSNIlion,adUseClient, adLockReadOnly
		if rst.recordcount>1 then
	   		if nz_b(d_lookup("art_multiempresa","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))<>0 then
				while not rst.EOF
					if rst("ncliente")<>session("ncliente") and rst("ncliente")<>SISTEMA_GESTION then 'La empresa actual no se procesa
						'con=d_lookup("dsn","clientes","ncliente='" & rst("ncliente") & "'",DSNIlion)
						con=session("dsn_cliente")
						rso=d_lookup("rsocial","clientes","ncliente='" & rst("ncliente") & "'",DSNIlion)
						ref=d_lookup("referencia","articulos","referencia='" & rst("ncliente")&p_referencia & "'",con)&""
						if ref<>"" then
							msg=msg & "La referencia " & trimCodEmpresa(ref) & " ya existe en la empresa " & rso & ".\n"
						else
							alm=d_lookup("codigo","almacenes","codigo like '" & rst("ncliente") & "%'",con)&""
							if alm="" then
								msg=msg & LITNOEXISALMDEFART & rso & ".\n"
							else
								alm=d_lookup("almacen","configuracion","nempresa='" & rst("ncliente") & "'",con)&""
								rstAux2.open "select * from articulos where referencia='" & session("ncliente")&p_referencia & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
								if ClavesAjenas(con,rstAux2,rso,rst("ncliente"))="OK" then
									rstAux.open "select * from articulos where referencia='" & rst("ncliente")&p_referencia & "'",con,adOpenKeyset,adLockOptimistic
									rstAux.addnew
										rstAux("REFERENCIA")=rst("ncliente")&right(rstAux2("referencia"),len(rstAux2("referencia"))-5)
										rstAux("NOMBRE")=rstAux2("nombre")
										rstAux("IMPORTE")=null_z(rstAux2("importe"))
										rstAux("RECARGO")=null_z(rstAux2("recargo"))
										rstAux("MARGEN")=null_z(rstAux2("margen"))
										rstAux("DESCUENTO")=null_z(rstAux2("descuento"))
										rstAux("PVP")=null_z(rstAux2("pvp"))
										rstAux("FECHAMOD")=nulear(rstAux2("fechamod"))
										rstAux("IVA")=null_z(rstAux2("iva"))

										rstAux("MEDIDA")=nulear(rstAux2("medida"))

	' >>> MCA 02/12/04 : Ofrecer los campos <Unidad aux.venta> y <Calcular importe detalle> siempre,
	'						con indenpendencia del módulo contratado

										if isnull(rstAux2("medidaventa")) then
											rstAux("MEDIDAVENTA")=nulear(rstAux2("medidaventa"))
										else
											rstAux("MEDIDAVENTA")=rst("ncliente")&right(rstAux2("medidaventa"),len(rstAux2("medidaventa"))-5)
										end if
										rstAux("CALCULOIMPORTE")=null_z(rstAux2("calculoimporte"))

										if 1=2 then	' Para conservar el código anterior
											if si_tiene_modulo_produccion<>0 then
												rstAux("MEDIDAVENTA")=nulear(rstAux2("medidaventa"))
												rstAux("CALCULOIMPORTE")=null_z(rstAux2("calculoimporte"))
											end if
										end if

	' >>> MCA 02/12/04 : Ofrecer los campos <Unidad aux.venta> y <Calcular importe detalle> siempre,
	'						con indenpendencia del módulo contratado


										rstAux("MODELO")=nulear(rstAux2("modelo"))

										rstAux("MESES")=null_z(rstAux2("meses"))
										rstAux("TMANOOBRA")=null_z(rstAux2("tmanoobra"))
										rstAux("TDESP")=null_z(rstAux2("tdesp"))
										rstAux("TMATERIALES")=null_z(rstAux2("tmateriales"))
										rstAux("PORCOM")=null_z(rstAux2("porcom"))
										rstAux("CTRL_NSERIE")=null_z(rstAux2("ctrl_nserie"))
										rstAux("IMPR_CATALOGO")=null_z(rstAux2("impr_catalogo"))
										rstAux("OBSERVACIONES")=nulear(rstAux2("observaciones"))
										rstAux("FOTO")=rstAux2("foto")
										rstAux("TIPO_FOTO")=nulear(rstAux2("tipo_foto"))
										rstAux("COD_BARRAS")=nulear(rstAux2("cod_barras"))
										rstAux("CARACTERISTICAS")=nulear(rstAux2("caracteristicas"))

										rstAux("FBAJA")=nulear(rstAux2("fbaja"))
										rstAux("CONTROL_STOCK")=null_z(rstAux2("control_stock"))
                                        rstAux("LOTECOMPRA")=null_z(rstAux2("LOTECOMPRA"))
										rstAux("ES_PADRE")=null_z(rstAux2("es_padre"))
										rstAux("REF_PADRE")=NULL

										rstAux("subctaventas")=nulear(rstAux2("subctaventas"))
										rstAux("subctaabventas")=nulear(rstAux2("subctaabventas"))
										rstAux("subctacompras")=nulear(rstAux2("subctacompras"))
										rstAux("subctaabcompras")=nulear(rstAux2("subctaabcompras"))

										rstAux("campo01")=nulear(rstAux2("campo01"))
										rstAux("campo02")=nulear(rstAux2("campo02"))
										rstAux("campo03")=nulear(rstAux2("campo03"))
										rstAux("campo04")=nulear(rstAux2("campo04"))
										rstAux("campo05")=nulear(rstAux2("campo05"))
										rstAux("campo06")=nulear(rstAux2("campo06"))
										rstAux("campo07")=nulear(rstAux2("campo07"))
										rstAux("campo08")=nulear(rstAux2("campo08"))
										rstAux("campo09")=nulear(rstAux2("campo09"))
										rstAux("campo10")=nulear(rstAux2("campo10"))

										rstAux("FOTO2")=rstAux2("foto2")
										rstAux("TIPO_FOTO2")=nulear(rstAux2("tipo_foto2"))

										rstAux("FOTO3")=rstAux2("foto3")
										rstAux("TIPO_FOTO3")=nulear(rstAux2("tipo_foto3"))

										rstAux("LINKWEB")=nulear(rstAux2("linkweb"))
                                        rstAux("LINKWEB2")=nulear(rstAux2("linkweb2"))
                                        
										rstAux("DOCUMENTOTIENDA1")=nulear(rstAux2("documentotienda1"))
										rstAux("DOCUMENTOTIENDA2")=nulear(rstAux2("documentotienda2"))

										rstAux("DOCUMENTO")=nulear(rstAux2("documento"))
										rstAux("AVISO")=nulear(rstAux2("aviso"))

										rstAux("LOADTER")=null_z(rstAux2("loadter"))
										rstAux("ESCVARIABLE")=null_z(rstAux2("escvariable"))

										'CLAVES AJENAS

										if isnull(rstAux2("familia")) then
											rstAux("FAMILIA")=nulear(rstAux2("familia"))
										else
											rstAux("FAMILIA")=rst("ncliente")&right(rstAux2("familia"),len(rstAux2("familia"))-5)
										end if

										if isnull(rstAux2("talla")) then
											rstAux("TALLA")=nulear(rstAux2("talla"))
										else
											rstAux("TALLA")=rst("ncliente")&right(rstAux2("talla"),len(rstAux2("talla"))-5)
										end if

										if isnull(rstAux2("color")) then
											rstAux("COLOR")=nulear(rstAux2("color"))
										else
											rstAux("COLOR")=rst("ncliente")&right(rstAux2("color"),len(rstAux2("color"))-5)
										end if

										if isnull(rstAux2("tipogar")) then
											rstAux("TIPOGAR")=nulear(rstAux2("tipogar"))
										else
											rstAux("TIPOGAR")=rst("ncliente")&right(rstAux2("tipogar"),len(rstAux2("tipogar"))-5)
										end if

										if isnull(rstAux2("divisa")) then
											rstAux("DIVISA")=nulear(rstAux2("divisa"))
										else
											rstAux("DIVISA")=rst("ncliente")&right(rstAux2("divisa"),len(rstAux2("divisa"))-5)
										end if

										if rstAux2("agrtallas")&""="" then
											rstAux("AGRTALLAS")=nulear(rstAux2("agrtallas"))
										else
											rstAux("AGRTALLAS")=rst("ncliente")&right(rstAux2("agrtallas"),len(rstAux2("agrtallas"))-5)
										end if

										if rstAux2("agrcolores")&""="" then
											rstAux("AGRCOLORES")=nulear(rstAux2("agrcolores"))
										else
											rstAux("AGRCOLORES")=rst("ncliente")&right(rstAux2("agrcolores"),len(rstAux2("agrcolores"))-5)
										end if

										if rstAux2("ncircuito")&""="" then
											rstAux("NCIRCUITO")=nulear(rstAux2("ncircuito"))
										else
											rstAux("NCIRCUITO")=rst("ncliente")&right(rstAux2("ncircuito"),len(rstAux2("ncircuito"))-5)
										end if

										if rstAux2("tipo_articulo")&""="" then
											rstAux("TIPO_ARTICULO")=nulear(rstAux2("tipo_articulo"))
										else
											rstAux("TIPO_ARTICULO")=rst("ncliente")&right(rstAux2("tipo_articulo"),len(rstAux2("tipo_articulo"))-5)
										end if

									rstAux.update
									rstAux.close
									''ricardo 1-10-2004 como dio un error porque se introdujo un campo nuevo en precios se pone los campos a actualizar
									rstAux.open "insert into precios (referencia,tarifa,rango,temporada,pvpdto,es_dto) values ('" & rst("ncliente")&p_referencia & "','" & rst("ncliente") & "BASE','" & rst("ncliente") & "BASE','" & rst("ncliente") & "BASE'," & reemplazar(pvp,",",".") & ",0)",con,adOpenKeyset,adLockOptimistic
									rstAux.open "insert into almacenar (articulo,almacen,ubicacion,stock,stock_minimo,reposicion,p_recibir,p_servir,p_min,predet) values ('" & rst("ncliente")&p_referencia & "','" & alm & "',NULL,0,0,0,0,0,0,1)",con,adOpenKeyset,adLockOptimistic
								end if
								rstAux2.close
							end if
						end if
					end if
					rst.movenext
				wend
			end if
		end if
		rst.close
		
		''ricardo 12-11-2007 si el parametro ne viene con valor, se guardara el articulo en la otra empresa
        if ne & "">"" then
            msg=GuardarArtEnOtraEmpresa(p_referencia,session("ncliente"),ne)
        end if
        
        'if si_tiene_modulo_franquicia and esFranquiciador=true then
		'    strArtFranq="EXEC GuardarCambiosArtEnFranquicias @nempresaCentral='" & session("ncliente") & "' ,@referencia='" & session("ncliente") & p_referencia & "' ,@mode=1 "
	    '    set conn = Server.CreateObject("ADODB.Connection")
	    '    conn.open DSNImport
	    '    conn.CommandTimeout = 0
	    '    set rst = conn.execute(strArtFranq)
	    '    conn.close
	    '    set conn = nothing
		'end if
        
		if msg>"" then
	   		%><script language="javascript" type="text/javascript">
                     alert("<%=LITPRODERRART%>\n<%=msg%>");
			</script><%
		end if
		''ricardo 10/4/2003 - si carga_terminal, creamos un registro en articuloster
		if request.form("carga_terminal") = "on" then
			guardar_carga_terminal request.form("carga_terminal"),p_referencia
		end if
	end if
	
	'*** i AMP 30/08/2011 Comprobar si existe configuracion de precio por defecto para la categoria , familia o subfamilia introducidas.
		if changeprices and  mode="first_save" then		              
            set connCFSPrice = Server.CreateObject("ADODB.Connection")
            
            strEdit="EXEC CheckCatFamSubFam_SalesList @ncompany='" & session("ncliente") & "', @operation=1 , @productRef='" & session("ncliente") & p_referencia & "'" &_
            ",@category='" & Nulear(categoria) & "'" &_
            ",@family='" & Nulear(fampadre) & "'" &_
            ",@subfamily='" & Nulear(request.form("familia")) & "'" &_
            ",@result=0"           
	        ''response.Write strEdit
	        ''response.end     
	        connCFSPrice.open session("dsn_cliente")                     
            set rstCFS = connCFSPrice.execute(strEdit)
            if not rstCFS.eof then
                result = rstCFS("result")
                'if result=0 then
                    %> <!--<script>alert("OK");</script> --><% 
                'end if
            end if		                       
            connCFSPrice.close 
    
            set rstCFS = Nothing
            set connCFSPrice = Nothing                         
	    end if
	'*** f AMP
	if mode="first_save" then	
	   '' response.Write p_referencia&","&pvpAnt&","&pvp&","&mode	      
	    GuardarHistorialCambios p_referencia,pvpAnt,pvp,mode
	end if
	
end sub

'-------------------------------------------------------------------------------------------------------------
'Comprobacion de varias claves ajenas de la tabla ARTICULOS
'-------------------------------------------------------------------------------------------------------------
function ClavesAjenas(con,rst,rso,cli)
	if rst("agrcolores")&""<>"" then
		if d_lookup("codigo","agrupa_colores","codigo='" & cli & right(rst("agrcolores"),len(rst("agrcolores"))-5) & "'",con)&""="" then
			msg=msg & LITAGRCOLNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("agrtallas")&""<>"" then
		if d_lookup("codigo","agrupa_tallas","codigo='" & cli & right(rst("agrtallas"),len(rst("agrtallas"))-5) & "'",con)&""="" then
			msg=msg & LITAGRTALNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("ncircuito")&""<>"" then
		if d_lookup("codigo","circuitos_fab","codigo='" & cli & right(rst("ncircuito"),len(rst("ncircuito"))-5) & "'",con)&""="" then
			msg=msg & LITCIRCNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("familia")&""<>"" then
		if d_lookup("codigo","familias","codigo='" & cli & right(rst("familia"),len(rst("familia"))-5) & "'",con)&""="" then
			msg=msg & LITFAMNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("divisa")&""<>"" then
		if d_lookup("codigo","divisas","codigo='" & cli & right(rst("divisa"),len(rst("divisa"))-5) & "'",con)&""="" then
			msg=msg & LITDIVNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("iva")&""<>"" then
		if d_lookup("tipo_iva","tipos_iva","tipo_iva='" & rst("iva") & "'",con)&""="" then
			msg=msg & LITTIPIVANOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("talla")&""<>"" then
		if d_lookup("codigo","tallas","codigo='" & cli & right(rst("talla"),len(rst("talla"))-5) & "'",con)&""="" then
			msg=msg & LITTALNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("color")&""<>"" then
		if d_lookup("codigo","colores","codigo='" & cli & right(rst("color"),len(rst("color"))-5) & "'",con)&""="" then
			msg=msg & LITCOLNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("tipogar")&""<>"" then
		if d_lookup("codigo","tipos_garantia","codigo='" & cli & right(rst("tipogar"),len(rst("tipogar"))-5) & "'",con)&""="" then
			msg=msg & LITTIPGARNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("medidaventa")&""<>"" then
		if d_lookup("codigo","medidas","codigo='" & cli & right(rst("medidaventa"),len(rst("medidaventa"))-5) & "'",con)&""="" then
			msg=msg & LITMEDAUXNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	if rst("tipo_articulo")&""<>"" then
		if d_lookup("codigo","tipos_entidades","codigo='" & cli & right(rst("tipo_articulo"),len(rst("tipo_articulo"))-5) & "'",con)&""="" then
			msg=msg & LITTIPARTNOEXISEMPART & rso & ".\n"
			ClavesAjenas="ERROR"
			exit function
		end if
	end if

	ClavesAjenas="OK"
end function

'****************************************************************************************************************
function EliminarRegistro(ref)
	seleccion="select referencia from detalles_ped_pro where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_ped_cli where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_alb_pro where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_alb_cli where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_fac_pro where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_fac_cli where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_dev_cli where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_dev_pro where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_pre_cli where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_orden where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	'-----------------------
	'COD : JCI-18122002-01
	'-----------------------
	seleccion=seleccion & "select referencia from detalles_orden_fab where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	'---------------------------
	'FIN COD : JCI-18122002-01
	'---------------------------

	''ricardo 13-3-2003
	''se controla tambien que el articulo no este en catalogos
		seleccion=seleccion & "select referencia from detalles_cat where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	'''''
	seleccion=seleccion & "select articulo from det_pre_cli_param where articulo ='" & ref & "' or articulo in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select articulo from det_ped_cli_param where articulo ='" & ref & "' or articulo in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select articulo from det_alb_cli_param where articulo ='" & ref & "' or articulo in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select articulo from det_fac_cli_param where articulo ='" & ref & "' or articulo in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select referencia from detalles_tickets where referencia ='" & ref & "' or referencia in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select ref from detalles_movimientos where ref ='" & ref & "' or ref in (select referencia from articulos where ref_padre='" & ref & "') UNION "
	seleccion=seleccion & "select art_hijo from escandallo where art_hijo ='" & ref & "'"

	Resultado=0
	set conn = Server.CreateObject("ADODB.Connection")
	set command =  Server.CreateObject("ADODB.Command")

	conn.open session("dsn_cliente")
	command.ActiveConnection =conn
	command.CommandTimeout = 0
	command.CommandText="BorrarArticulo"
	command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	command.Parameters.Append command.CreateParameter("@Referencia",adVarChar,adParamInput,30,ref)
	Command.Parameters.Append Command.CreateParameter("@resul", adInteger, adParamOutput, Resultado)
	on error resume next
	command.Execute,,adExecuteNoRecords
	Resultado = Command.Parameters("@resul").Value
if err.number<>0 then
		%><script language="javascript" type="text/javascript">
            alert("<%=err.description%>")
		</script><%
end if
	on error goto 0
	conn.close
	set command=nothing
	set conn=nothing

	if Resultado=1 then
		%><script language="javascript" type="text/javascript">
              alert("<%=LitMsgBorrarArticulo%>")
		</script><%
		exit function
	else
		''ricardo 30-4-2004 solamente se puede auditar si se ha borrado el articulo
		''rst.close

		rst.open "select referencia from articulos where ref_padre like '" & session("ncliente") & "%' and ref_padre='" & ref & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if not rst.eof then
			rst.close
			nom_art_pad_hij=ref & "("
			rst.open "select referencia from articulos where ref_padre like '" & session("ncliente") & "%' and ref_padre='" & ref & "' and color is not null",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if not rst.eof then
				nom_art_pad_hij=nom_art_pad_hij & LitAgrcolores & ","
			end if
			rst.close
			rst.open "select referencia from articulos where ref_padre like '" & session("ncliente") & "%' and ref_padre='" & ref & "' and talla is not null",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if not rst.eof then
				nom_art_pad_hij=nom_art_pad_hij & LitAgrTallas & ","
			end if
			rst.close
			nom_art_pad_hij=mid(nom_art_pad_hij,1,len(nom_art_pad_hij)-1) & ")"
			auditar_ins_bor session("usuario"),"","","baja",nom_art_pad_hij,"","articulos"
		else
			rst.close
			auditar_ins_bor session("usuario"),"","","baja",ref,"","articulos"
		end if
	end if
end function

'****************************************************************************************************************
Sub Propiedades(modo)
	response.write("<input type='hidden' name='hes_padre' value='" & EncodeForHtml(nz_b(res_padre)) & "'>")
	response.write("<input type='hidden' name='hreferencia' value='" & EncodeForHtml(rreferencia) & "'>")
	response.write("<input type='hidden' name='hmostrar_foto' value='" & EncodeForHtml(nz_b(mostrar_foto)) & "'>")

    '	------------------------------------------------
    '	GPD (01/03/2007) - Añadir campos "novedad" y "Unidad de embalaje"		
	tam = 600 '260
	if modo = "add" then tam =  680''500 '460
	if modo = "edit" then tam = 760''500 '430
   
	'	------------------------------------------------
	response.write("<input type='hidden' name='modm' value='" & EncodeForHtml(modmargen) & "'>")
	''RGU 27/12/2005
	response.write("<input type='hidden' name='fc01' value='" & EncodeForHtml(c01) & "'>")
	response.write("<input type='hidden' name='fc02' value='" & EncodeForHtml(c02) & "'>")
	response.write("<input type='hidden' name='fc03' value='" & EncodeForHtml(c03) & "'>")

	response.write("<iframe name='marcoPropiedades' id='frPropiedades'  src='PropiedadesArticulo.asp?mode=" & EncodeForHtml(modo) & "&referencia=" & EncodeForHtml(rreferencia) & "&modm=" & EncodeForHtml(modmargen) & "&t="& EncodeForHtml(tarifa) &"&c01="& EncodeForHtml(c01) &"&c02="& EncodeForHtml(c02) &"&c03="& EncodeForHtml(c03) &"' frameborder='0' width='100%' height='" & EncodeForHtml(tam) & "' ></iframe>")
end sub
'************************************************************************************************************

'************************************************************************************************************
sub Foto(modo)
	if modo="browse" then
		iframe_alto=250
	else
		iframe_alto=400
	end if
	response.write("<iframe name='marcoFoto' id='frFoto' src='Articulos_Imagen.asp?mode=" & EncodeForHtml(modo) & "&referencia=" & EncodeForHtml(rreferencia) & "&mf=" & EncodeForHtml(nz_b(mostrar_foto)) &"' frameborder='0' width='100%' height='" & EncodeForHtml(iframe_alto) & "'></iframe>")
end sub

'******************************************************************************************************************
sub DatosContaDeArticulo(modo,referencia)
    %><iframe name='marcoDatosContaDeArticulo' id='frDatosContaDeArticulo' src='DatosContaDeArticulo.asp?mode=<%=EncodeForHtml(modo)%>&referencia=<%=EncodeForHtml(referencia)%>' frameborder='0' width='100%' height="170"></iframe><%
end sub
'******************************************************************************************************************
sub DatosFabDeArticulo(modo,referencia)
    %><iframe name='marcoDatosFabDeArticulo' id='frDatosFabDeArticulo' src='DatosFabDeArticulo.asp?mode=<%=EncodeForHtml(modo)%>&referencia=<%=EncodeForHtml(referencia)%>' frameborder='0' width='100%' height="75"></iframe><%
end sub
'******************************************************************************************************************
sub PreciosArticulo()

//		if si_tiene_modulo_tiendas<>0 then
			'response.write("<table cellpadding=1 cellspacing=1 width=950 border='0'>")
//		else
//			'response.write("<table cellpadding=1 cellspacing=1 width=598 border='0'>")
//		end if
            
            
			if res_padre<>0 then
				
					seleccion="select referencia,substring(referencia,6,len(referencia)-5) as nombre from articulos with(NOLOCK) where ref_padre like '" & session("ncliente") & "%' and referencia like '" & session("ncliente") & "%' and ref_padre='" & rreferencia & "' order by referencia"
                    rstAux.cursorlocation=3
					rstAux.open seleccion,session("dsn_cliente")
		   			DrawSelectCelda "width:200px","200","",0,LitCombinaciones,"i_hijotcp",rstAux,"","referencia","nombre","onchange","GestionPrecios('selectC','" & server.urlencode(rreferencia) & "','" & nz_b(res_padre) & "')"
		   			
                    rstAux.close
				
			end if
			 %>
                <div class="overflowXauto">
                <table class="width90 md-table-responsive bCollapse">
                    <tr>
                    <td class="ENCABEZADOL underOrange width10"><%=LitTarifa %></td>

                    <%if si_tiene_modulo_tiendas <> 0 then%>
                        <td class="ENCABEZADOL underOrange width10"><%=LitTemporada %></td>
                        <td class="ENCABEZADOL underOrange width10"><%=LitRango %></td>
                    <% end if%>

                    <td class="ENCABEZADOC underOrange width10"><%=LitPvp%></td>
                    <td class="ENCABEZADOC underOrange width10"><%=LitPvpIva %></td>
                    <td class="ENCABEZADOC underOrange width10"><%=LitPorcentajePvp %></td>
                    <td class="ENCABEZADOC underOrange width10"><%=LitPorcentajeCoste %></td>
                    <td class="ENCABEZADOC underOrange width10"><%=LitPvpFinal %></td>
                    <td class="ENCABEZADOC underOrange width10" style="text-align:center;"><input type='Checkbox' name='checkb' onclick=seleccionar('marcoPreciosArticulo','PreciosDeArticulo','checkb');></td>
                  </tr>
                    <tr>
                
                <!--<table class="underOrange width90" style="border-collapse:inherit; table-layout: fixed; bgcolor="<%=color_blanco_det%>" cellpadding="0" cellspacing="0">-->
                    
                
    
    <%
                rst.cursorlocation=3
				rst.open "select codigo, descripcion from tarifas where upper(codigo)<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
                DrawSelectCeldaDet "'CELDAL7 underOrange width10'","width100","",0,"","tarifa",rst,"","codigo","descripcion","",""
				rst.close
				if si_tiene_modulo_tiendas<>0 then
                    rst.cursorlocation=3
					rst.open "select codigo, descripcion from temporadas where upper(codigo)<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
					DrawSelectCeldaDet "'CELDAL7 underOrange width10'","width100","",0,"","temporada",rst,"","codigo","descripcion","",""
                    rst.close
                    
                    rst.cursorlocation=3
					rst.open "select codigo, descripcion from rangos where upper(codigo)<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
                    DrawSelectCeldaDet "'CELDAL7 underOrange width10'","width100","",0,"","rango",rst,"","codigo","descripcion","",""
					rst.close
				end if
                %><td class="CELDAR7 underOrange width10"><%
                    DrawInput "'CELDAR7 width65'", "", "precio", "", "onchange=javascript:muestrapreciofinal('P','" & rpvp & "'," & ndecimales & ");"
                %></td><%
                'DrawInputCeldaActionDiv "","","","35",0,"","precio",0, "onchange", "javascript:muestrapreciofinal('P','" & rpvp & "'," & ndecimales & ");",false
				
                %><td class="CELDAR7 underOrange width10">
                    <span class='CELDA7' id="incpvp" style="background-color: transparent; border: 0px;"> </span><%
                    DrawInput "'CELDAR7 width65'", "", "precioiva", "", "onchange=javascript:muestrapreciofinal('PI','" & rpvp & "'," & ndecimales & ");"
                    %><%
                %></td><%
				'DrawInputCeldaActionDiv "","","","35",0,"","precioiva",0, "onchange", "javascript:muestrapreciofinal('PI','" & rpvp & "'," & ndecimales & ");",false
				
                %><!--<div class='CELDAR7' id="incpvp" style="background-color: transparent; border: 0px;"> </div>--><%
				'DrawInputCeldaAction "CELDAR7","","",5,0,"","descuento","","onchange","javascript:muestrapreciofinal('D','" & rpvp & "'," & ndecimales & ");",false

                %><td class="CELDAR7 underOrange width10">
                    <span class='CELDAR7' id="incimp" style="background-color: transparent; border: 0px;"> </span><%
                    DrawInput "'CELDAR7 width65'", "", "descuento", "", "onchange=javascript:muestrapreciofinal('D','" & rpvp & "'," & ndecimales & ");"
                    %><%
                %></td><%
				'DrawInputCeldaActionDiv "","","","35",0,"","descuento",0, "onchange", "javascript:muestrapreciofinal('D','" & rpvp & "'," & ndecimales & ");",false
				
                %><!--<div class='CELDAR7' id="incimp" style="background-color: transparent; border: 0px;"> </div>--><%
				'DrawInputCeldaAction "CELDAR7","","",5,0,"","descuentocoste","","onchange","javascript:muestrapreciofinal('DC','" & rimporte & "'," & ndecimales & ");",false

                %><td class="CELDAR7 underOrange width10"><%
                    DrawInput "'CELDAR7 width65'", "", "descuentocoste", "", "onchange=javascript:muestrapreciofinal('DC','" & rimporte & "'," & ndecimales & ");"
                %></td><%
				'DrawInputCeldaActionDiv "","","","35",0,"","precioiva",0, "onchange", "javascript:muestrapreciofinal('DC','" & rimporte & "'," & ndecimales & ");",false
				
                'DrawInputCelda "CELDAR7  READONLY","","",10,0,"","preciofinal",""
                'EligeCelda "input","add","left","","",0,LitPvpFinal,"preciofinal",35,""
                               
                'DrawInputCeldaImg(estilo, ancho, alto, nchar, tabulacion, etiqueta, name, dato, funcion, mensaje,src)
                'DrawInput(clase, estilo, name, dato, otros)
               
				empresa_sup=d_lookup("empresa_sup","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")) & ""
				bloqueoprecios=""
				colorfilabloqueada=""

                 'DrawInputCeldaImg "","","","",0,LitPvpFinal,"preciofinal","","if(AnadirPrecio('" & server.urlencode(rreferencia) & "','" & nz_b(res_padre) & "','" & rpvp & "','" & rimporte & "'))", "", ImgNuevo & "' " & ParamImgNuevo & " alt='" & LitNuevoPrecio
                %><td class="CELDAR7 underOrange width10"><%
                    DrawInput "'CELDAR7 width65'", "", "preciofinal", "", ""
                %></td><%                

				if nz_b(rpreciosbloqueados)<>0 and session("ncliente")<>empresa_sup and empresa_sup<>"" then
					'response.write("<td class=CELDAREFB align=center><a href=javascript:PrecioBloqueado();><img src='../images/" & ImgNuevo & "' " & ParamImgNuevo & " alt='" & LitNuevoPrecio & "'></a></td>")
                %><td class="CELDAC7 underOrange width10" style="text-align:center;"><%
                    response.write("<a class='ic-accept noMTop' href=javascript:PrecioBloqueado();><img src='" & themeIlion & "" & ImgNuevo & "' " & ParamImgNuevo & " alt='" & LitNuevoPrecio & "'></a>")                    
                %></td><%   

                    
				else
                %><td class="CELDAC7 underOrange width10" style="text-align:center;"><%
                    response.write("<a class='ic-accept noMTop' href=javascript:if(AnadirPrecio('" & server.urlencode(rreferencia) & "','" & nz_b(res_padre) & "','" & rpvp & "','" & rimporte & "'));><img src='" & themeIlion & "" & ImgNuevo & "' " & ParamImgNuevo & " alt='" & LitNuevoPrecio & "'></a>")
                %></td><%   
					'response.write("<td class=CELDAREFB align=center>")
                    response.write("<input type='hidden' name='T_referencia' value='" & EncodeForHtml(server.urlencode(rreferencia)) & "'/>")
                    response.write("<input type='hidden' name='T_es_padre' value='" & EncodeForHtml(nz_b(res_padre)) & "'/>")
                    response.write("<input type='hidden' name='T_pvp' value='" & EncodeForHtml(rpvp) & "'/>")
                    response.write("<input type='hidden' name='T_importe' value='" & EncodeForHtml(rimporte) & "'/>")
				end if
                    
		%></tr></table>
                    
                    
                    <%
		'response.write("</table>")

		response.write("<iframe name='marcoPreciosArticulo' id='frPreciosArticulo' class='width90 iframe-data md-table-responsive' src='PreciosDeArticulo.asp?mode=select&referencia=" & EncodeForHtml(rreferencia) & "' frameborder='0' width='975'></iframe>")
		response.write("<table class='width90 md-table-responsive bCollapse'>")
				response.write("<td class='ENCABEZADOL7 width90'>")
					if res_padre<>0 then%>
                         <label style="width:100px;">
                             <a class='CELDAREF7' href="javascript:VentanaPreciosTyC('<%=server.urlencode(rreferencia)%>');"> <%=LITPRECBASEART2%>  </a>
                         </label>
                        <%
                        DrawSpan "","", EncodeForHtml(formatnumber(rpvp,dec_prec,-1,0,-1)) ,""					    
                        'response.write("<div id='pbase'><a class='CELDAREF7' href=javascript:VentanaPreciosTyC('" & server.urlencode(rreferencia) & "');>" & LITPRECBASEART2 & "</a> : " & formatnumber(rpvp,dec_prec,-1,0,-1) & "</div>")
					else
                        DrawLabel "","", LITPRECBASEART & EncodeForHtml(formatnumber(rpvp,dec_prec,-1,0,-1))
						'response.write("<div id='pbase'>" & LITPRECBASEART & formatnumber(rpvp,dec_prec,-1,0,-1) & "</div>")
					end if
				response.write("</td>")
				'response.write("<td class='ENCABEZADOR7 width30' style='text-align:right;'>")
					if nz_b(rpreciosbloqueados)<>0 and session("ncliente")<>empresa_sup and empresa_sup<>"" then
						response.write("<td class='ENCABEZADOR7 width10' id='IcoBorrModif'  style='text-align: right;'><a href=javascript:PrecioBloqueado();><img src='" & themeIlion & "" & ImgDiskette & "' " & ParamImgDiskette & " alt='" & LITGUARDPREART & "'></a>&nbsp;")
						response.write("<a class='ic-delete noMTop' href=javascript:PrecioBloqueado();><img src='" & themeIlion & "" & ImgEliminarDet & "' " & ParamImgEliminar & " alt='" & LITELIMPREART & "'></a>")
					else
						response.write("<td class='ENCABEZADOR7 width10' id='IcoBorrModif'  style='text-align: right;'><a href=javascript:if(GestionPrecios('save','" & server.urlencode(rreferencia) & "','" & nz_b(res_padre) & "'));><img src='" & themeIlion & "" & ImgDiskette & "' " & ParamImgDiskette & " alt='" & LITGUARDPREART & "'></a>&nbsp;")
						response.write("<a class='ic-delete noMTop' href=javascript:if(GestionPrecios('delete','" & server.urlencode(rreferencia) & "','" & nz_b(res_padre) & "'));><img src='" & themeIlion & "" & ImgEliminarDet & "' " & ParamImgEliminar & " alt='" & LITELIMPREART & "'></a>")
					end if
				response.write("</td>")
		response.write("</table>")
	    %></div><%
end sub

sub CamposPersonalizables(modo,referencia)
    response.write("<iframe name='marcoCamposPersonalizables' id='frCamposPersonalizables' src='CamposPersonalizablesArt.asp?mode=" & EncodeForHtml(modo) & "&referencia=" & EncodeForHtml(referencia) & "&c01="&EncodeForHtml(c01)&"&c02="&EncodeForHtml(c02)&"&c03="&EncodeForHtml(c03)&"' frameborder='0' width='100%'></iframe>")
end sub

function CheckAutoReference ()
    srtselect = "select autoref from configuracion with(nolock) where nempresa = ?"
    CheckAutoReference = DLookUpP1(strselect, session("ncliente")&"", adVarChar, 5, session("dsn_cliente"))
end function


'****************************************************************************************************************
'  CODIGO PRINCIPAL DE LA PAGINA
'****************************************************************************************************************
	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion %>
<form name="articulos" method="post">
    <%
   
    
    Dim msg
    
    campo    = limpiaCadena(request.QueryString("campo"))
	if campo & ""="" then
	    campo = Request.Form("campo")
    end if
    
''response.Write("el campo es-" & campo & "-<br/>")

	texto    = limpiaCadena(request.QueryString("texto"))
	''texto    = limpiaCadena(server.URLEncode(request.QueryString("texto")))
	if texto & ""="" then
	    texto = Request.Form("texto")
    end if
''response.Write("el texto es-" & texto & "-" & limpiaCadena(request.QueryString("texto")) & "-" & request.form("texto") & "-" & server.URLEncode(limpiaCadena(request.QueryString("texto"))) & "-<br/>")
	
	lote=limpiaCadena(Request.QueryString("lote"))
	nt=limpiaCadena(request.QueryString("nt"))
	if lote & ""="" then
	    lote=Request.form("lote")
	end if
	if lote="" then lote=1
	sentido=limpiaCadena(Request.QueryString("sentido"))
	if sentido & ""="" then
	    sentido=Request.form("sentido")
	end if
	criterio=limpiaCadena(Request.QueryString("criterio"))
	if criterio & ""="" then
	    criterio=Request.form("criterio")
	end if
''response.Write("el criterio es-" & criterio & "-<br/>") 
	
	total_paginas=0
	total_paginas=limpiaCadena(Request.QueryString("total_paginas"))
	if total_paginas & ""="" then
	    total_paginas=Request.form("total_paginas")
	end if
	if total_paginas & ""="" then total_paginas=0
	
	npagina=0
	npagina=limpiaCadena(Request.QueryString("npagina"))
	if npagina & ""="" then
	    npagina=Request.form("npagina")
	end if
	if npagina & ""="" then npagina=0%>
	<input type="hidden" name="si_tiene_modulo_comercial" value="<%=EncodeForHtml(si_tiene_modulo_comercial)%>">
	<input type="hidden" name="si_tiene_modulo_terminales" value="<%=EncodeForHtml(si_tiene_modulo_terminales)%>">
	<input type="hidden" name="si_tiene_modulo_tiendas" value="<%=EncodeForHtml(si_tiene_modulo_tiendas)%>"><%
	
    'Recordsets'
    set rst = Server.CreateObject("ADODB.Recordset")
    set rst2 = Server.CreateObject("ADODB.Recordset")
    set rstAux = Server.CreateObject("ADODB.Recordset")
    set rstAux2 = Server.CreateObject("ADODB.Recordset")
    set rstAux3 = Server.CreateObject("ADODB.Recordset")
    set rst_almacenar = Server.CreateObject("ADODB.Recordset")
    set rst_proveer = Server.CreateObject("ADODB.Recordset")
    set rst_escandallo = Server.CreateObject("ADODB.Recordset")
    set rsttallas = Server.CreateObject("ADODB.Recordset")
    set rstcolores = Server.CreateObject("ADODB.Recordset")
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")

    ''conn.open DSNILION
    ''command.ActiveConnection =conn
    ''command.CommandTimeout = 0
    ''command.CommandText= "ContractedItem"
    ''command.CommandType = adCmdStoredProc
    ''    command.Parameters.Append command.CreateParameter("@nempresa", adVarChar, adParamInput, 5, session("ncliente"))
    ''    command.Parameters.Append command.CreateParameter("@objeto", adVarChar, adParamInput, 1000, OBJManufacturer)
    ''command.Execute,,adExecuteNoRecords
    ''set rstAux = command.Execute
    ''
    ''if not rstAux.eof then
    ''    has_manufacturer_item = rstAux("result")
    ''else
        has_manufacturer_item = "0"
    ''end if 
    %><input type="hidden" name="h_has_manufacturer_item" value="<%=EncodeForHtml(has_manufacturer_item)%>"><% 
    
   ''conn.close
   ''command.Parameters.Delete ("@nempresa")
   ''command.Parameters.Delete ("@objeto")
    'urlRewrite es un campo del e-shop que indicará si queremos gestionar la url
    cadena="select urlrewrite from b2b_configuracion with(nolock) where empresa='"&session("ncliente")&"'"
	rst.Open cadena,session("dsn_cliente")
	if not rst.eof then
	    urlRewrite = rst("urlrewrite") 
	else
	    urlRewrite = false
	end if
	%><input type="hidden" name="urlRewrite" value="<%=EncodeForHtml(urlRewrite)%>"><% 
	rst.close

	''JMA 6/4/04. Si tiene asignada alguna de las páginas contables
	si_paginas_contables = VerObjeto(OBJEnlaceContable)
	%><input type="hidden" name="si_paginas_contables" value="<%=EncodeForHtml(si_paginas_contables)%>"><%

	''ricardo 24-3-2004 si existen campos personalizables con titulo no nulo si saldra la pestaña de campos personalizables
	si_campo_personalizables=0
    rst.cursorlocation=3
	rst.open "select ncampo from camposperso with(NOLOCK) where tabla='ARTICULOS' and isnull(titulo,'')<>'' and ncampo like '" & session("ncliente") & "%'",session("dsn_cliente")
	if not rst.eof then
		si_campo_personalizables=1
	else
		si_campo_personalizables=0
	end if
	rst.close
	%><input type="hidden" name="si_campo_personalizables" value="<%=EncodeForHtml(si_campo_personalizables)%>"><%

    '***********************************************************************
    ' TIENE QUE EXISTIR UN ALMACEN OBLIGATORIAMENTE EN EL SISTEMA
    '***********************************************************************
    rst.open "select codigo from almacenes where codigo like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    if not rst.eof then

	'SI QUE HAY DEFINIDO POR LO MENOS UN ALMACEN EL SISTEMA
	rst.close

	Dim Opciones() 'Array con las opciones de menu

	refresco=limpiaCadena(request.querystring("refresco"))
	tmpproveedorR=limpiaCadena(request.form("i_proveedor"))

	if refresco="buscar" and tmpproveedorR>"" then
		tmpdivisa = d_lookup("divisa", "proveedores", "nproveedor='" & tmpproveedorR & "'", session("dsn_cliente"))
		tmpproveedor=tmpproveedorR
	end if

	'Leer parámetros de la página
	mode = Request.QueryString("mode")
	
	%>
    <input type="hidden" name="h_mode" value="<%=EncodeForHtml(mode)%>"/>
    <input type="hidden" name="mode_accesos_tienda" value="<%=EncodeForHtml(mode)%>"/>
    <%

	dim modp,modd,modi,novei,cv,caju,ocb,ps,t,gest,c01,c02,c03
	dim au  ' cag Parametro que trae lista de almacenes para desplegable en pantalla del historial dependiendo del usuario
	dim artsl
	dim surt
	dim ne ''ricardo 9-11-2007 si el ne tiene valor se creara el mismo articulo creado en la empresa que venga en ese parametro ne
	
	dim pnf
    dim showcase
	surt=limpiaCadena(request.querystring("surt")&"")
	if surt = "" then
		surt=limpiaCadena(request.form("surt")&"")
	end if
    
   ' showcase = limpiaCadena(request.QueryString("showcase")&"")

	ObtenerParametros("articulos")%>
	<input type="hidden" name="artsl"       value="<%=EncodeForHtml(artsl)%>">
	<input type="hidden" name="pnf"         value="<%=EncodeForHtml(pnf)%>">
	<input type="hidden" name="surt"        value="<%=EncodeForHtml(surt)%>"> 
	<input type="hidden" name="ne"          value="<%=EncodeForHtml(ne)%>"> 
	<input type="hidden" name="campo"       value="<%=EncodeForHtml(campo)%>"/>
	<input type="hidden" name="texto"       value="<%=EncodeForHtml(texto)%>"/>
	<input type="hidden" name="lote"        value="<%=EncodeForHtml(lote)%>"/>
	<input type="hidden" name="criterio"    value="<%=EncodeForHtml(criterio)%>"/>
	<input type="hidden" name="urlmostrar"  value="<%=EncodeForHtml(request.Form("url_mostrar"))%>"/>
	<input type="hidden" name="showcase"    value="<%=EncodeForHtml(showcase) %>" />
	<%
	''cag 28-03-06 parametro almacenes por usuario
	' esto no se pone
	'au=limpiaCadena(request.querystring("au"))
	'if au="" then au=limpiaCadena(request.form("au"))

	%><input type="hidden" name="au" value="<%=EncodeForHtml(au)%>"> <%

	'obtememos el parámetro gest y lo guardamos en un campo oculto para que pueda ser accedido por el bt y en función de este dibujar o no
	'los botones que permiten modificar los artículos
	%>
	<input type="hidden" name="gestion" value="<%=EncodeForHtml(gest)%>">
	<%p_referencia = limpiaCadena(request.querystring("referencia"))
	if p_referencia="" then
   		p_referencia=limpiaCadena(request.querystring("ndoc")) 'Llamada desde el listado de numeros de serie
	end if
	if mode<>"first_save" then
		CheckCadena p_referencia
	end if

	if request.querystring("ver_historial")>"" then
		ver_historial=limpiaCadena(request.querystring("ver_historial"))
	else
		ver_historial=limpiaCadena(request.form("ver_historial"))
	end if
	%><input type="hidden" name="ver_historial" value="<%=EncodeForHtml(ver_historial)%>"><%

	modmargen=limpiaCadena(request.querystring("modm") & "")
	if modmargen="" then modmargen=limpiaCadena(request.form("modm"))

	'RGU 27/12/2005: gestionar ofertas de shades
	if c01="" then c01= limpiaCadena(Request.QueryString("c01") &"")
	if c01="" then c01=limpiaCadena(request.form("fc01"))
	if c02="" then c02=limpiaCadena(request.QueryString("c02") &"")
	if c02="" then c02=limpiaCadena(request.form("fc02"))
	if c03="" then c02=limpiaCadena(request.QueryString("c03") &"")
	if c03="" then c03=limpiaCadena(request.form("fc03"))

	'rgu

	h_refpro=limpiaCadena(request.querystring("h_refpro"))
	genera=limpiaCadena(request.querystring("genera"))
	agrupa_colores=limpiaCadena(request.form("agrupa_colores"))
	agrupa_tallas=limpiaCadena(request.form("agrupa_tallas"))

	nproveedor=limpiaCadena(request.querystring("nproveedor"))

	' MCA 24/02/05 : Parámetro de usuario indicando tarifa
	tarifa= limpiaCadena(request.querystring("t"))
	if tarifa="" then
		tarifa= limpiaCadena(request.form("h_tarifa"))
	end if
	if tarifa="" then
		tarifa= t
	end if%>
	<input type="hidden" name="h_tarifa" value="<%=EncodeForHtml(tarifa)%>">
	
	<!-- PBG 28/5/2007 Para ocultar artículos dados de baja-->
    <input type="hidden" name="ocultarArticulosBaja" id="ocultarArticulosBaja" value="<%=EncodeForHtml(ocultarArticulosBaja)%>" />
    <%if ps="" then ps=limpiaCadena(request.querystring("ps"))
	if ps="" then ps=limpiaCadena(request.form("ps"))

	ndecimales = d_lookup("ndecimales", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("dsn_cliente"))

	if mode="add" or mode="first_save" then%>
		<input type="hidden" name="h_referencia"    value="<%=EncodeForHtml(p_referencia)%>">
		<input type="hidden" name="h_pvp"           value="<%=EncodeForHtml(limpiaCadena(request.form("spvp")))%>">
		<input type="hidden" name="h_iva"           value="<%=EncodeForHtml(limpiaCadena(request.form("iva")))%>">
		<input type="hidden" name="h_descuento"     value="<%=EncodeForHtml(limpiaCadena(request.form("descuento")))%>">
		<input type="hidden" name="h_codbarras"     value="<%=EncodeForHtml(limpiaCadena(request.form("codbarras")))%>">
		<input type="hidden" name="h_meses"         value="<%=EncodeForHtml(limpiaCadena(request.form("meses")))%>">
		<input type="hidden" name="h_mesesmo"       value="<%=EncodeForHtml(limpiaCadena(request.form("mesesmo")))%>">
		<input type="hidden" name="h_mesesde"       value="<%=EncodeForHtml(limpiaCadena(request.form("mesesde")))%>">
		<input type="hidden" name="h_mesesmt"       value="<%=EncodeForHtml(limpiaCadena(request.form("mesesmt")))%>">
		<input type="hidden" name="h_nombre"        value="<%=EncodeForHtml(limpiaCadena(request.form("nombre")))%>">
		<input type="hidden" name="h_ctrl_nserie"   value="<%=EncodeForHtml(limpiaCadena(request.form("ctrl_nserie")))%>">
		<input type="hidden" name="h_control_stock" value="<%=EncodeForHtml(limpiaCadena(request.form("control_stock")))%>">
        <input type="hidden" name="h_gestion_lotes" value="<%=EncodeForHtml(limpiaCadena(request.form("gestion_lotes")))%>">
		<input type="hidden" name="h_fbaja"         value="<%=EncodeForHtml(limpiaCadena(request.form("fbaja")))%>">
		<input type="hidden" name="h_peso"          value="<%=EncodeForHtml(limpiaCadena(request.form("peso"))) %>" />
		<input type="hidden" name="h_divisa"        value="<%=EncodeForHtml(limpiaCadena(request.form("divisa")))%>">
		<input type="hidden" name="h_medida"        value="<%=EncodeForHtml(limpiaCadena(request.form("medida")))%>">

        <!-- MAP 24/12/2012 - Fecha Creación-->
        <input type="hidden" name="h_fcreacion" value="<%=EncodeForHtml(limpiaCadena(request.form("fechacreacion")))%>">

		<!-- GPD 26/02/2007 -->
		<input type="hidden" name="h_novedad" value="<%=EncodeForHtml(limpiaCadena(request.form("novedad")))%>">
		<input type="hidden" name="h_ue" value="<%=EncodeForHtml(limpiaCadena(request.form("ue")))%>">
        <!-- DBS (29/11/2013).-->
        <input type="hidden" name="h_uv" value="<%=EncodeForHtml(limpiaCadena(request.form("uv")))%>">

<%  ' >>> MCA 02/12/04 : Los campos <Unidad aux.venta> y <Calcular importe detalle> aparezcan siempre
	'						con independencia del módulo contratado%>
		<input type="hidden" name="h_medidaventa" value="<%=EncodeForHtml(limpiaCadena(request.form("medidaventa")))%>">
		<input type="hidden" name="h_calculoimporte" value="<%=EncodeForHtml(limpiaCadena(request.form("calculoimporte")))%>">

        <%if 1=2 then 	' Para conservar el código anterior sin que se ejecute
			if si_tiene_modulo_produccion<>0 then%>
			<input type="hidden" name="h_medidaventa" value="<%=EncodeForHtml(limpiaCadena(request.form("medidaventa")))%>">
			<input type="hidden" name="h_calculoimporte" value="<%=EncodeForHtml(limpiaCadena(request.form("calculoimporte")))%>">
            <%end if
		end if

	' <<< MCA 02/12/04 : Los campos <Unidad aux.venta> y <Calcular importe detalle> aparezcan siempre
	'						con independencia del módulo contratado%>
		<input type="hidden" name="h_color"         value="<%=EncodeForHtml(limpiaCadena(request.form("color")))%>">
		<input type="hidden" name="h_talla"         value="<%=EncodeForHtml(limpiaCadena(request.form("talla")))%>">
		<input type="hidden" name="h_familia"       value="<%=EncodeForHtml(limpiaCadena(request.form("familia")))%>">
		<input type="hidden" name="h_modelo"        value="<%=EncodeForHtml(limpiaCadena(request.form("modelo")))%>">
		<input type="hidden" name="h_tipogar"       value="<%=EncodeForHtml(limpiaCadena(request.form("tipogar")))%>">
		<input type="hidden" name="h_porcom"        value="<%=EncodeForHtml(limpiaCadena(request.form("porcom")))%>">
		<input type="hidden" name="h_agrupa_tallas"     value="<%=EncodeForHtml(agrupa_tallas)%>">
		<input type="hidden" name="h_agrupa_colores"    value="<%=EncodeForHtml(agrupa_colores)%>">
		<input type="hidden" name="h_genera"            value="<%=EncodeForHtml(limpiaCadena(request.form("genera")))%>">
		<input type="hidden" name="h_impr_catalogo"     value="<%=EncodeForHtml(limpiaCadena(request.form("impr_catalogo")))%>">
		<input type="hidden" name="h_carga_terminal"    value="<%=EncodeForHtml(limpiaCadena(request.form("carga_terminal")))%>">

		<input type="hidden" name="h_subctaventas"      value="<%=EncodeForHtml(request.form("subctaventas"))%>">
		<input type="hidden" name="h_subctaabventas"    value="<%=EncodeForHtml(limpiaCadena(request.form("subctaabventas")))%>">
		<input type="hidden" name="h_subctacompras"     value="<%=EncodeForHtml(limpiaCadena(request.form("subctacompras")))%>">
		<input type="hidden" name="h_subctaabcompras"   value="<%=EncodeForHtml(limpiaCadena(request.form("subctaabcompras")))%>">
		<input type="hidden" name="ps" value="<%=EncodeForHtml(limpiaCadena(ps))%>">
	<%end if
       'if mode="first_save" then
       '    response.write("subctaventas: " & enc.EncodeForHtmlAttribute(limpiaCadena(request.form("subctaventas"))))
       '    response.End
       'end if
	'Acción a realizar
	if mode="save" or mode="first_save" then '**********************************************************************
			%><input type="hidden" name="h_refpro" value=""><%
	   		if mode="first_save" then
				'Si la referencia está vacía en este punto, significa que hay que generarla
                gosave=1
				if p_referencia="" then
					p_referencia=AutoRef()
                else
                    p_referencia=trim(p_referencia)
                    if instr(p_referencia," ")>0 then
					    mode="add"
                        p_referencia=""
                        %><script type="text/javascript" language="javascript">
                              window.alert("<%=LitMsgRefDesCarNoVal%>");
					    </script><%
                        gosave=0
                    end if

				end if
                if gosave=1 then
				    rstAux.open "select referencia from articulos where referencia='" & session("ncliente") & p_referencia & "'", _
				    session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if not rstAux.eof then
					    rstAux.close
					    mode="add"
					    ''ricardo 6-5-2004 si no se pone a vacio la variable p_referencia el frame campospersonalizable_art.asp
					    ''en la funcion Checkcadena referencia nos sacara del sistema
					    p_referencia=""%>
					    <script type="text/javascript" language="javascript">
                            window.alert("<%=LitMsgReferenciaExiste%>");
                            parent.botones.document.location = "articulos_bt.asp?mode=add&t=<%=enc.EncodeForJavascript(tarifa)%>";
					    </script><%
		   		    else
		   			    rstAux.Close
					    rstAux.open "select * from proveer where nproveedor like '"&session("ncliente")&"%' and su_ref='" & p_referencia & "' ",session("dsn_cliente")
					    if not rstAux.eof and h_refpro&""="" then
						    rstAux.Close
						    mode="add"
                            ''ricardo 2-3-2005 se pone esto, ya que si no al aceptar o cancelar nos tira del sistema
                            p_referencia=""
						    %><script language="javascript" type="text/javascript">
                                  if (window.confirm("<%=LitMsgRefArtRefProvConfirm%>")) {
                                      document.articulos.h_refpro.value = "SI";
                                      document.articulos.style.display = "none";
                                  }
                                  else parent.botones.document.location = "articulos_bt.asp?mode=add&t=<%=enc.EncodeForJavascript(tarifa)%>";
						    </script><%
					    else
						    rstAux.close

						    if genera>"" then
							    GeneraTallasColores p_referencia
							    rstAux.open "select referencia from articulos where ref_padre='" & p_referencia & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
							    if not rstAux.eof then
								    nom_art_pad_hij=p_referencia & "("
								    if agrupa_colores>"" then
									    nom_art_pad_hij=nom_art_pad_hij & LitAgrcolores & ": " & agrupa_colores & ","
								    end if
								    if agrupa_tallas>"" then
									    nom_art_pad_hij=nom_art_pad_hij & LitAgrTallas & ": " & agrupa_tallas & ","
								    end if
								    nom_art_pad_hij=mid(nom_art_pad_hij,1,len(nom_art_pad_hij)-1) & ")"
								    auditar_ins_bor session("usuario"),"","","alta",nom_art_pad_hij,"","articulos"
							    end if
							    rstAux.close
						    else
                                GuardarRegistro p_referencia, "", "",""							
							    p_referencia=session("ncliente")&p_referencia
							    rstAux.open "select nombre from articulos where referencia='" & p_referencia & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
							    if not rstAux.eof then
								    auditar_ins_bor session("usuario"),"","","alta",p_referencia,"","articulos"
							    end if
							    rstAux.close
						    end if

						    rstAux.open "select nombre from articulos where referencia='" & p_referencia & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						    if not rstAux.eof then
							    mode="browse"
						    else
							    mode="add"
						    end if
						    rstAux.close
	    '''''''''''''''''''''''''''''''''''''''''''''
					    end if
				    end if
                end if
			else
				if genera>"" then
			   		GeneraTallasColores p_referencia
				else
                    GuardarRegistro p_referencia, "", "",""					
				end if
				mode="browse"
			end if
	elseif mode="delete" then '************************************************************************************
		auditar_ins_bor session("usuario"),"","","",p_referencia,"","Inicio baja articulo"
	    EliminarRegistro p_referencia
		p_referencia = ""
		mode="add"%>
		  <script language="javascript" type="text/javascript">
              parent.botones.document.location = "articulos_bt.asp?mode=add";
              SearchPage("articulos_lsearch.asp?mode=init", 0);
		   </script><%
	elseif mode="traerprov" then
		nproveedor = completar(nproveedor,5,"0")
		Error="NO"
		rstAux.open "select * from proveedores where nproveedor='" & nproveedor & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if not rstAux.EOF then
	  		tmp_nproveedor=nproveedor
			tmp_nombre=rstAux("razon_social")
			tmp_divisa=rstAux("divisa")
		else
			Error="SI"
			tmp_nproveedor=""
			%><script language="javascript" type="text/javascript">
                  window.alert("<%=LitMsgProveedorNoExiste%>");
			</script><%
		end if
		rstAux.close
		mode="browse"
	end if

	p_nombre = d_lookup("nombre","articulos","referencia='"+p_referencia+"'",session("dsn_cliente"))

	if mode="edit" or mode="browse" then
		rst.open "select * from articulos with(nolock) where referencia='" + p_referencia + "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if not rst.eof then
			rreferencia      = rst("referencia")
			rnombre          = rst("nombre")
			rimporte         = rst("importe")
			rrecargo         = rst("recargo")
			rdescuento       = rst("descuento")
			rpvp             = rst("pvp")
			riva             = rst("iva")
			rfamilia         = rst("familia")
			rmedida          = rst("medida")

	' >>> MCA 02/12/04 : Ofrecer los campos <Unidad aux.venta> y <Calcular importe detalle> siempre
	'						con indenpendencia del módulo contratado
			rmedidaventa     = rst("medidaventa")
			rcalculoimporte  = rst("calculoimporte")

			if 1=2 then		' Para conservar el código anterior sin que se ejecute
				if si_tiene_modulo_produccion<>0 then
					rmedidaventa     = rst("medidaventa")
					rcalculoimporte  = rst("calculoimporte")
				end if
			end if

	' <<< MCA 02/12/04 : Los campos <Unidad aux.venta> y <Calcular importe detalle> aparezcan siempre
	'						con indenpendencia del módulo contratado

'**RGU 17/1/2007
%>
<input type="hidden" name="IO" value="<%=EncodeForHtml(null_z(rimporte))%>">
<input type="hidden" name="PO" value="<%=EncodeForHtml(null_z(rpvp))%>">
<%
'**RGU 17/1/2007

			rtalla           = rst("talla")
			rcolor           = rst("color")
			rmodelo          = rst("modelo")
			rtipogar         = rst("tipogar")
		 	rmeses           = rst("meses")
		 	rporcom          = rst("porcom")
		 	rctrl_nserie	= rst("ctrl_nserie")
		 	rcontrol_stock	= rst("control_stock")
            rgestion_lotes = rst("LOTECOMPRA")
		 	rfbaja			= rst("fbaja")
		 	rpeso           = rst("weight")
		 	rdescatalogado = rst("discontinued")
		 	rimpr_catalogo   = rst("impr_catalogo")
		 	robservaciones   = rst("observaciones")
		 	rcod_barras      = rst("cod_barras")
		 	rcaracteristicas = rst("caracteristicas")
		 	rdivisa          = rst("divisa")
			rref_padre       = rst("ref_padre") & ""
			res_padre		 = rst("es_padre")
			rcirfab          = rst("ncircuito")

			rsubctaventas		= rst("subctaventas")
			rsubctaabventas		= rst("subctaabventas")
			rsubctacompras		= rst("subctacompras")
			rsubctaabcompras	= rst("subctaabcompras")

			%><script language="javascript" type="text/javascript">es_padreGlobal =<%=nz_b(res_padre) %></script><%

		 	if rst("tipo_foto")>"" and not isnull(rst("tipo_foto")) then
		    	mostrar_foto = true
		 	else
		    	mostrar_foto = false
		 	end if
			rpreciosbloqueados=rst("preciosbloqueados")
			rst.close
		else
			rst.close
			%><script language="javascript" type="text/javascript">
                      window.alert("<%=LitMsgArticuloNoExiste%>");
                  document.location = "articulos.asp?mode=search&t=<%=enc.EncodeForJavascript(tarifa)%>";
                  parent.botones.document.location = "articulos_bt.asp?mode=search&t=<%=enc.EncodeForJavascript(tarifa)%>";
			</script><%
			mode=""
			CerrarTodo()
			response.end
		end if
	end if

	PintarCabecera "articulos.asp"

	'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION--------------------------------------
   
           %><div class="headers-wrapper"><%
                DrawDiv "header-bill","",""
                DrawLabel"","",Litreferencia
                DrawSpan "","",EncodeForHtml(trimCodEmpresa(p_referencia)),""
                CloseDiv

	            DrawDiv "header-bill","",""
                DrawLabel"","",Litnombre
                DrawSpan "","",EncodeForHtml(p_nombre),""
                CloseDiv
            %> </div><table width='100%'></table><%
                    %><!--<span><%=Litreferencia &": "%></span><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(p_referencia))%>-->
                   
               
                    <!--<span><img src="<%=ImgLineVertical %>" id="img" style="vertical-align:top;" class="line_vertical" /></span>-->
                <!--<div class="data_client">
                    <span><%=Litnombre &": "%></span><%=p_nombre%>
                </div>-->
            
			
			<!-- PBG 28/05/2005 Para no mostrar artículos con fecha de baja-->
	
	
	<%'---------------------------------------------------------------------------------------------------------

	alarma "articulos.asp"
	BarraOpciones2 mode,p_referencia,nz_b(res_padre),rref_padre,ps,au,artsl,surt

    AbrirModal "fr_ChangePrices","",800,450,"no","si","no","si","Cambios de Precios"

   
	'---------------------------------------------------------------------------------------------------------------
	'Modo de inserción/edición
	'---------------------------------------------------------------------------------------------------------------
	if mode = "add" or mode="edit" then
      	if mode = "add" then
			rimporte   = 0
			rrecargo   = 0
			rmeses     = 0
			rporcom    = 0
			rdescuento = 0
			rctrl_nserie = false
			rcontrol_stock = false
            rgestion_lotes = false
			rimpr_catalogo = false
			riva = d_lookup("iva","configuracion", "nempresa='" & session("ncliente") & "'", session("dsn_cliente"))
			rpvp = 0

			BarraNavegacion "add",artsl

		    if si_tiene_modulo_ecomerce<>0 then
			    texto_tituloFoto=LitFicArtImg2
		    else
			    texto_tituloFoto=LitFicArtImg
		    end if
			%>
            <div id="CollapseSection"> 
                <!--<a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['addPRO','addDCONT', 'addMD','addCPerso']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> -->
                <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['addPRO','addDCONT','addCPerso']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
                <!--<a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['addPRO','addDCONT', 'addMD','addCPerso']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>-->
                <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['addPRO','addDCONT','addCPerso']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
            </div>

            <div class="Section" id="S_addPRO">
                <a href="#" rel="toggle[addPRO]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitPropiedades%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="addPRO">
	         <!--<br />-->
             <!--<center>-->
                <%Propiedades "add"%>
            <!--</center>-->
            </div>
            </div>

            <div class="Section" id="S_addDCONT">
                <a href="#" rel="toggle[addDCONT]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitDCont%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel"  style="display:none;" id="addDCONT">
	         <!--<br />-->
             <!--<center>-->
                <%DatosContaDeArticulo "add",""%>
            <!--</center>-->
            </div>
            </div>

            <%if has_manufacturer_item ="1_ahora_no" then %>
            <div class="Section" id="S_addMD">
                <a href="#" rel="toggle[addMD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LITMANUFACTURERDATA%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel"  style="display:none;" id="addMD">
	         <!--<br />-->
             <!--<center>-->
                <%DatosFabDeArticulo "add",""%>
            <!--</center>-->
            </div>
            </div>
            <%end if%>

            <%if si_tiene_modulo_importaciones=0 then
                if si_campo_personalizables=1 then
                    %>
                    <div class="Section" id="S_addCPerso">
                        <a href="#" rel="toggle[addCPerso]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                            <div class="SectionHeader">
                                <%=LitCampPerso%>
                                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                            </div>
                        </a>
                    <div class="SectionPanel"  style="display:none;" id="addCPerso">
	                 <!--<br />-->
                     <!--<center>-->
                        <%CamposPersonalizables "add",EncodeForHtml(p_referencia)%>
                    <!--</center>-->
                    </div>
                    </div>
                    <%
                end if
            end if
		else
	  		BarraNavegacion "edit",artsl

		    if si_tiene_modulo_ecomerce<>0 then
			    texto_tituloFoto=LitFicArtImg2
		    else
			    texto_tituloFoto=LitFicArtImg
		    end if
            %>
            <div id="CollapseSection"> 
                <!--<a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['editPRO','editDCONT', 'editMD','editFO','editCPerso']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> -->
                <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['editPRO','editDCONT','editFO','editCPerso']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
                <!--<a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['editPRO','editDCONT', 'editMD','editFO','editCPerso']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>-->
                <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['editPRO','editDCONT','editFO','editCPerso']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
            </div>

            <div class="Section" id="S_editPRO">
                <a href="#" rel="toggle[editPRO]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitPropiedades%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="editPRO">
	         <!--<br />-->
             <!--<center>-->
                <%Propiedades "edit"%>
            <!--</center>-->
            </div>
            </div>

            <div class="Section" id="S_editDCONT">
                <a href="#" rel="toggle[editDCONT]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitDCont%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel"  style="display:none;" id="editDCONT">
	         <!--<br />-->
             <!--<center>-->
                <%DatosContaDeArticulo "edit",EncodeForHtml(p_referencia)%>
            <!--</center>-->
            </div>
            </div>

            <% if has_manufacturer_item ="1_ahora_no" then%>
            <div class="Section" id="S_editMD">
                <a href="#" rel="toggle[editMD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LITMANUFACTURERDATA%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel"  style="display:none;" id="editMD">
	         <!--<br />-->
             <!--<center>-->
                <%DatosFabDeArticulo "edit",EncodeForHtml(p_referencia)%>
            <!--</center>-->
            </div>
            </div>
            <%end if %>
            <div class="Section" id="S_editFO">
                <a href="#" rel="toggle[editFO]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=texto_tituloFoto%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel"  style="display:none;" id="editFO">
	         <!--<br />-->
             <!--<center>-->
                <%Foto "edit"%>
            <!--</center>-->
            </div>
            </div>

            <%if si_tiene_modulo_importaciones=0 then
                if si_campo_personalizables=1 then%>
                    <div class="Section" id="S_editCPerso">
                        <a href="#" rel="toggle[editCPerso]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                            <div class="SectionHeader">
                                <%=LitCampPerso%>
                                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                            </div>
                        </a>
                    <div class="SectionPanel"  style="display:none;" id="editCPerso">
	                 <!--<br />-->
                     <!--<center>-->
                        <%CamposPersonalizables "edit",EncodeForHtml(p_referencia)%>
                    <!--</center>-->
                    </div>
                    </div>
                    <%
                end if
            end if

			
		end if
        strselect = "select autoref from configuracion with(nolock) where nempresa=?"
        refaut = nz_b(DLookupP1(strselect,session("ncliente")&"",adVarChar,5,session("dsn_cliente")))
		%><input type="hidden" name="autoref" value='<%=EncodeForHtml(refaut)%>'><%

	'---------------------------------------------------------------------------------------------------------------
	'Modo de visualizacion
	'---------------------------------------------------------------------------------------------------------------
	elseif mode="browse" then
	  	ndecimales = d_lookup("ndecimales", "divisas", "codigo='" & rdivisa & "'", session("dsn_cliente"))

		BarraNavegacion "Browse",artsl
		if si_tiene_modulo_ecomerce<>0 then
			texto_tituloFoto=LitFicArtImg2
		else
			texto_tituloFoto=LitFicArtImg
		end if

		%><input type="hidden" name="hiva" value="<%=EncodeForHtml(riva)%>"><%

        %>
            <div id="CollapseSection"> 
                <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['BrowsePRO','BrowseDCONT','BrowseFO','BrowsePRE','BrowseCPerso']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
                <!--<a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['BrowsePRO','BrowseDCONT', 'BrowseMD','BrowseFO','BrowsePRE','BrowseCPerso']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> -->
                <!--<a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['BrowsePRO','BrowseDCONT', 'BrowseMD','BrowseFO','BrowsePRE','BrowseCPerso']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>-->
                <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['BrowsePRO','BrowseDCONT','BrowseFO','BrowsePRE','BrowseCPerso']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
            </div>

            <div class="Section" id="S_BrowsePRO">
                <a href="#" rel="toggle[BrowsePRO]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitPropiedades%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="BrowsePRO">
	         <!--<br />-->
             <!--<center>-->
                <%Propiedades "browse"%>
            <!--</center>-->
            </div>
            </div>

            <div class="Section" id="S_BrowseDCONT">
                <a href="#" rel="toggle[BrowseDCONT]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitDCont%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel"  style="display:none;" id="BrowseDCONT">
	         <!--<br />-->
             <!--<center>-->
                <%DatosContaDeArticulo "browse",EncodeForHtml(p_referencia)%>
            <!--</center>-->
            </div>
            </div>

            <% if has_manufacturer_item ="1_ahora_no" then%>
            <div class="Section" id="S_BrowseMD">
                <a href="#" rel="toggle[BrowseMD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LITMANUFACTURERDATA%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" style="display:none;" id="BrowseMD">
	         <!--<br />-->
             <!--<center>-->
                <%DatosFabDeArticulo "browse",EncodeForHtml(p_referencia)%>
            <!--</center>-->
            </div>
            </div>
            <%end if %>

            <div class="Section" id="S_BrowseFO">
                <a href="#" rel="toggle[BrowseFO]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=texto_tituloFoto%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" style="display:none;" id="BrowseFO">
	         <!--<br />-->
             <!--<center>-->
                <%Foto "browse"%>
            <!--</center>-->
            </div>
            </div>
            <%if cstr(artsl)<>"1" then %>
            <div class="Section" id="S_BrowsePRE">
                <a href="#" rel="toggle[BrowsePRE]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitPrecios%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" style="display:none;" id="BrowsePRE">
	         <!--<br />-->
             <!--<center>-->
                <%PreciosArticulo%>
            <!--</center>-->
            </div>
            </div>
            <%end if %>
            <%if si_tiene_modulo_importaciones=0 then
                if si_campo_personalizables=1 then
                    %>
                    <div class="Section" id="S_BrowseCPerso">
                        <a href="#" rel="toggle[BrowseCPerso]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                            <div class="SectionHeader">
                                <%=LitCampPerso%>
                                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                            </div>
                        </a>
                    <div class="SectionPanel" style="display:none;" id="BrowseCPerso">
	                 <!--<br />-->
                     <!--<center>-->
                        <%CamposPersonalizables "browse",EncodeForHtml(p_referencia)%>
                    <!--</center>-->
                    </div>
                    </div>
                    <%
                end if
            end if

		
	elseif mode="search" then


	end if%>
	<input type="hidden" name="total_paginas" value="<%=total_paginas%>"/>
</form><%
   else
   		rst.close
   		'NO EXISTE NINGUN ALMACEN EN EL SISTEMA
   		%><script language="javascript" type="text/javascript">
                 alert("<%=LitMsgAlmNoExiste%>");
                 parent.parent.document.location = "../search_layout.asp?pag1=productos/almacenes.asp?mode=add&pag2=productos/almacenes_bt.asp&pag3=productos/almacenes_lsearch.asp";
		</script><%
   end if
end if
connRound.close
set connRound = Nothing
set conn = Nothing
set command=nothing
CerrarTodo()%>
</body>
</html>