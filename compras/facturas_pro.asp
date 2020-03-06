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

''ricardo 5-6-2003 se añade el parametro novei para que en los formatos de impresion no salga el item

''''ricardo 31/7/2003 comprobamos que existe la factura que se ha pedido ver desde un listado, sino se va al modo add
' FLM : 19/01/2009 : Añadir captura de nproveedor por request
' jcg 20/01/2009: Añadida la columna proyecto al proveedor y tratamiento de la misma.

'mmg: variables para obtener los almacenes por defecto
dim almacenSerie
dim almacenTPV

linea1=session("f_tpv")
linea2=session("f_caja")
linea3=session("f_empr")
strconn=session("dsn_cliente")



'Calculamos el almacen por defecto del TPV
set rsTPV = Server.CreateObject("ADODB.Recordset")

cadena= "select c.almacen from tpv a with(NOLOCK), cajas b with(NOLOCK), tiendas c with(NOLOCK), almacenes alm with(NOLOCK) where a.caja=b.codigo and b.tienda=c.codigo and tpv='" +linea1 +"' and b.codigo='" +linea2+"' and alm.codigo=c.almacen and isnull(alm.fbaja,'')=''"
rsTPV.cursorlocation=3
rsTPV.Open cadena,session("dsn_cliente")
if rsTPV.eof then
	almacenTPV= ""
else
	almacenTPV= rsTPV("almacen")
end if
rsTPV.close

nfactura = limpiaCadena(Request.QueryString("nfactura"))
if nfactura="" then nfactura=limpiaCadena(Request.form("nfactura"))
if nfactura ="" then
	nfactura = limpiaCadena(Request.QueryString("ndoc"))
	if nfactura ="" then
		nfactura = limpiaCadena(Request.form("ndoc"))
	end if
end if

CheckCadena nfactura

DivisaFactura=d_lookup("divisa","facturas_pro","nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "'",session("dsn_cliente"))
NdecDiFactura=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & DivisaFactura & "'",session("dsn_cliente"))
if NdecDiFactura & "" = "" then
    NdecDiFactura=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente"))
end if%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<!--#include file="../Custom/aspJSON1.17.asp" -->
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

<!--#include file="facturas_pro.inc" -->
<!--#include file="compras.inc" -->
<!--#include file="../ventas/documentos.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../varios2.inc" -->
<!--#include file="../perso.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../common/modal2.inc" -->

<!--#include file="../js/animatedCollapse.js.inc" -->

<!--#include file="../js/tabs.js.inc" -->

<!--#include file="../styles/generalData.css.inc" -->

<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc"-->

<!--#include file="../styles/Tabs.css.inc" -->


<!--#include file="../js/dropdown.js.inc" -->

<!--#include file="../common/facturas_proActionDrop.inc" -->
<!--#include file="../common/poner_cajaResponsive.inc" -->

<!--#include file="../js/calendar.inc" -->

<!--#include file="facturas_pro_linkextra.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->


<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('CABECERA', 'fade=1');
    animatedcollapse.addDiv('DATFINAN', 'fade=1');

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init();
</script>
<style type="text/css">
 html .ui-autocomplete {
		height: 100px;
	}
</style>

<%
si_tiene_modulo_21=ModuloContratado(session("ncliente"),"21")
si_tiene_modulo_22=ModuloContratado(session("ncliente"),"22")
si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
si_tiene_modulo_contabilidad=ModuloContratado(session("ncliente"),ModContabilidad)
si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
si_tiene_modulo_bierzo=ModuloContratado(session("ncliente"),ModBierzo)
si_tiene_modulo_ccostes=ModuloContratado(session("ncliente"),ModCcostes_Gestion)
si_tiene_modulo_gesteconomica=ModuloContratado(session("ncliente"),ModGestionEconomica)
si_tiene_modulo_gesteconomica=ModuloContratado(session("ncliente"),ModGestionEconomica) 
themeIlion="/lib/estilos/" & folder & "/"
%>

<script language="javascript" type="text/javascript">
    function cambiarfecha(fecha, modo) {
        var fecha_ar = new Array();

        if (fecha != "") {
            suma = 0;
            fecha_ar[suma] = "";
            l = 0
            while (l <= fecha.length) {
                if (fecha.substring(l, l + 1) == '/') {
                    suma++;
                    fecha_ar[suma] = "";
                }
                else {
                    if (fecha.substring(l, l + 1) != '') fecha_ar[suma] = fecha_ar[suma] + fecha.substring(l, l + 1);
                }
                l++;
            }
            if (suma != 2) {
                window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
                return false;
            }
            else {
                nonumero = 0;
                while (suma >= 0 && nonumero == 0) {
                    if (isNaN(fecha_ar[suma])) nonumero = 1;
                    if (fecha_ar[suma].length > 2 && suma != 2) nonumero = 1;
                    if (fecha_ar[suma].length > 4 && suma == 2) nonumero = 1;
                    suma--;
                }

                if (nonumero == 1) {
                    window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
                    return false;
                }
            }
        }
        return true;
    }

    //Desencadena la búsqueda del artículo cuya referencia se indica
    function TraerSerie(mode) {
        prov_old = document.facturas_pro.nproveedor.value;
        if (prov_old != "" && prov_old.length < 5) {
            for (i = prov_old.length; i < 5; i++) {
                prov_old = "0" + prov_old;
            }
        }
        document.facturas_pro.nproveedor.value = prov_old;
        document.facturas_pro.razon_social.value = "";
        document.facturas_pro.forma_pago.value = "";
        document.facturas_pro.tipo_pago.value = "";
        document.facturas_pro.divisa.value = "";
        document.facturas_pro.divisabis.value = "";
        document.facturas_pro.descuento.value = 0;
        document.facturas_pro.descuento2.value = 0;
        document.facturas_pro.recargo.value = 0;
        document.facturas_pro.ncuentacargo.value = "";

        if (confirm("<%=LitCambiarProPuedCamSer%>")) cambiar_serie = 1;
        else cambiar_serie = 0;

        document.facturas_pro.nproveedor.value = "";
        document.location.href = "facturas_pro.asp?ndoc=" + document.facturas_pro.nfactura.value + "&nproveedor=" + prov_old + "&mode=" + mode + "&prov=" + prov_old + "&observaciones=" + document.facturas_pro.observaciones.value + "&serie=" + document.facturas_pro.serie.value + "&fecha=" + document.facturas_pro.fecha.value +
            "&nfactura_pro=" + document.facturas_pro.nfactura_pro.value + "&viene=" + document.facturas_pro.viene.value
            <%if si_tiene_modulo_proyectos<>0 then%>
                + "&cod_proyecto=" + document.facturas_pro.cod_proyecto.value
                <%end if%>
                    + "&caju=" + document.facturas_pro.caju.value + "&novei=" + document.facturas_pro.novei.value +
                    "&incoterms=" + document.facturas_pro.incoterms.value +
                    "&cambiar_serie=" + cambiar_serie +
                    "&fob=" + document.facturas_pro.fob.value +
                    "&s=" + document.facturas_pro.s.value + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
    }

    function TraerProveedor(mode) {
        prov_old = document.facturas_pro.nproveedor.value;
        cambiar_cliente = "";

        if (confirm("<%=LitCambiarSeriePuedCamPro%>")) cambiar_cliente = 1;
        else cambiar_cliente = 0;

        document.facturas_pro.razon_social.value = "";
        document.facturas_pro.forma_pago.value = "";
        document.facturas_pro.tipo_pago.value = "";
        document.facturas_pro.divisa.value = "";
        document.facturas_pro.divisabis.value = "";
        document.facturas_pro.descuento.value = 0;
        document.facturas_pro.descuento2.value = 0;
        document.facturas_pro.recargo.value = 0;
        document.facturas_pro.ncuentacargo.value = "";

        document.facturas_pro.nproveedor.value = "";

        document.location.href = "facturas_pro.asp?ndoc=" + document.facturas_pro.nfactura.value + "&nproveedor=" + "" + "&mode=" + mode + "&prov=" + prov_old + "&observaciones=" + document.facturas_pro.observaciones.value + "&serie=" + document.facturas_pro.serie.value + "&fecha=" + document.facturas_pro.fecha.value +
            "&nfactura_pro=" + document.facturas_pro.nfactura_pro.value + "&viene=" + document.facturas_pro.viene.value
            <%if si_tiene_modulo_proyectos<>0 then%>
                + "&cod_proyecto=" + document.facturas_pro.cod_proyecto.value
                <%end if%>
                    + "&caju=" + document.facturas_pro.caju.value + "&novei=" + document.facturas_pro.novei.value +
                    "&incoterms=" + document.facturas_pro.incoterms.value +
                    "&fob=" + document.facturas_pro.fob.value +
                    "&cambiar_cliente=" + cambiar_cliente +
                    "&s=" + document.facturas_pro.s.value + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
    }

    //***************************************************************************
    function Precios() {
        if (isNaN(document.facturas_pro.dto1.value.replace(",", ".")) || isNaN(document.facturas_pro.dto2.value.replace(",", ".")) || isNaN(document.facturas_pro.rf.value.replace(",", ".")))
            window.alert("<%=LitMsgDto1Dto2RfNumerico%>");
        else {
            //Preparamos los datos para trabajar***************************************
            dto1SinComas = document.facturas_pro.dto1.value.replace(",", ".");
            dto2SinComas = document.facturas_pro.dto2.value.replace(",", ".");
            rfSinComas = document.facturas_pro.rf.value.replace(",", ".");
            //TOTAL DESCUENTO**********************************************************
            dto1 = (parseFloat(document.facturas_pro.importe_bruto.value.replace(",", ".")) * parseFloat(dto1SinComas)) / 100;
            dto2 = ((parseFloat(document.facturas_pro.importe_bruto.value.replace(",", ".")) - dto1) * parseFloat(dto2SinComas)) / 100;
            dtoTotal = dto1 + dto2;
            c_dtoTotal = dtoTotal.toString();
            document.facturas_pro.total_descuento.value = parseFloat(c_dtoTotal).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.h_total_descuento.value = document.facturas_pro.total_descuento.value;
            //BASE IMPONIBLE***********************************************************
            base_imponible = parseFloat(document.facturas_pro.importe_bruto.value) - parseFloat(document.facturas_pro.total_descuento.value);
            c_base_imponible = base_imponible.toString();
            document.facturas_pro.base_imponible.value = parseFloat(c_base_imponible).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.h_base_imponible.value = document.facturas_pro.base_imponible.value;
            //TOTAL IVA****************************************************************
            dto1 = ((parseFloat(document.facturas_pro.sumadet.value) * parseFloat(dto1SinComas)) / 100);
            dto2 = ((parseFloat(document.facturas_pro.sumadet.value) - dto1) * parseFloat(dto2SinComas)) / 100;
            dtoTotal = dto1 + dto2;
            total_iva = parseFloat(document.facturas_pro.sumadet.value) - dtoTotal;
            c_total_iva = total_iva.toString();
            document.facturas_pro.total_iva.value = parseFloat(c_total_iva).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.h_total_iva.value = document.facturas_pro.total_iva.value;
            //RECARGO FINANCIERO*******************************************************
            total_rf = (parseFloat(document.facturas_pro.base_imponible.value) * rfSinComas) / 100;
            c_total_rf = total_rf.toString();
            document.facturas_pro.total_rf.value = parseFloat(c_total_rf).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.h_total_rf.value = document.facturas_pro.total_rf.value;
            //RECARGO DE EQUIVALENCIA**************************************************
            total_re = parseFloat(document.facturas_pro.sumaRE.value);
            c_total_re = total_re.toString();
            document.facturas_pro.total_re.value = parseFloat(c_total_re).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.h_total_re.value = parseFloat(c_total_re).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            //RETENCIÓN FISCAL*******************************************************
            if (document.facturas_pro.IRPF_Total.value == "True" || document.facturas_pro.IRPF_Total.value == "1") {
                baseImp = parseFloat(document.facturas_pro.base_imponible.value) +
                    parseFloat(document.facturas_pro.total_iva.value) +
                    parseFloat(document.facturas_pro.total_re.value) +
                    parseFloat(document.facturas_pro.total_rf.value);
            }
            else baseImp = document.facturas_pro.base_imponible.value;
            total_irpf = (parseFloat(baseImp) * irpfSinComas) / 100;
            c_total_irpf = total_irpf.toString();
            document.facturas_pro.total_irpf.value = parseFloat(c_total_irpf).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.h_total_irpf.value = document.facturas_pro.total_irpf.value;
            //TOTAL
            total_factura = parseFloat(document.facturas_pro.base_imponible.value.replace(",", ".")) + parseFloat(document.facturas_pro.total_iva.value.replace(",", ".")) + parseFloat(document.facturas_pro.total_re.value.replace(",", ".")) + parseFloat(document.facturas_pro.total_rf.value.replace(",", "."));
            c_total_factura = total_factura.toString();
            document.facturas_pro.total_factura.value = parseFloat(c_total_factura).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.h_total_factura.value = document.facturas_pro.total_factura.value;
            //VOLVEMOS A DEJAR LOS DATOS CAMBIADOS COMO ESTABAN************************
            alert(dto1SinComas)
            document.facturas_pro.dto1.value = dto1SinComas;
            document.facturas_pro.dto2.value = dto2SinComas;
            document.facturas_pro.rf.value = rfSinComas;
        }
    }

    //Calcula el importe de la línea de detalle del concepto.
    function ImporteDetalle() {
        if (parseFloat(document.facturas_pro.pvp.value) < 0) {
            window.alert("<%=LitMsgPvPNoNegativo%>");
            document.facturas_pro.pvp.value = 0;
        }
        if (isNaN(document.facturas_pro.cantidad.value.replace(",", ".")) || isNaN(document.facturas_pro.descuento.value.replace(",", ".")) || isNaN(document.facturas_pro.pvp.value.replace(",", ".")))
            window.alert("<%=LitMsgCanPreDesNumerico%>");
        else {
            if (document.facturas_pro.pvp.value == "") document.facturas_pro.pvp.value = 0;
            if (document.facturas_pro.cantidad.value == "") document.facturas_pro.cantidad.value = 1;
            if (document.facturas_pro.descuento.value == "") document.facturas_pro.descuento.value = 0;
            pvpSinComas = document.facturas_pro.pvp.value.replace(",", ".");
            cantidadSinComas = document.facturas_pro.cantidad.value.replace(",", ".");
            dtoSinComas = document.facturas_pro.descuento.value.replace(",", ".");
            pelas = parseFloat(cantidadSinComas) * parseFloat(pvpSinComas);
            pelas_descuento = (pelas * parseFloat(dtoSinComas)) / 100;
            importe = pelas - pelas_descuento;
            c_importe = importe.toString();
            document.facturas_pro.descuento.value = dtoSinComas;
            document.facturas_pro.importe.value = parseFloat(c_importe).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
            document.facturas_pro.pvp.value = pvpSinComas;
            document.facturas_pro.descuento.value = dtoSinComas;
            document.facturas_pro.cantidad.value = cantidadSinComas;
        }
    }

    //Comprueba si el importe del pago es numerico
    function importepagoComp() {
        if (isNaN(document.facturas_pro.importePago.value.replace(",", "."))) {
            window.alert("<%=LitErrImportePago2%>");
            return;
        }
    }

    //Añade un concepto a la factura
    function addConcepto(nfactura) {
        if (isNaN(document.facturas_pro.cantidad.value.replace(",", ".")) || isNaN(document.facturas_pro.descuento.value.replace(",", ".")) || isNaN(document.facturas_pro.pvp.value.replace(",", "."))) {
            window.alert("<%=LitMsgCanPreDesNumerico%>");
            return;
        }

        if (document.facturas_pro.descripcion.value == "") {
            alert("<%=LitMsgDesVacia%>");
            return;
        }
        if (isNaN(document.facturas_pro.pvp.value.replace(",", "."))) {
            window.alert("<%=LitMsgImporteNumerico%>");
            return;
        }
        //Asignar los valores a los campos del submarco de detalles
        fr_Conceptos.document.facturas_procon.cantidad.value = document.facturas_pro.cantidad.value;
        fr_Conceptos.document.facturas_procon.descripcion.value = document.facturas_pro.descripcion.value;
        fr_Conceptos.document.facturas_procon.pvp.value = document.facturas_pro.pvp.value;
        fr_Conceptos.document.facturas_procon.descuento.value = document.facturas_pro.descuento.value;
        fr_Conceptos.document.facturas_procon.iva.value = document.facturas_pro.iva.value;
        //Recargar el submarco de detalles
        fr_Conceptos.document.facturas_procon.action = "facturas_procon.asp?mode=first_save";
        fr_Conceptos.document.facturas_procon.submit();
        //Limpiar los campos del formulario
        document.facturas_pro.cantidad.value = "1";
        document.facturas_pro.descripcion.value = "";
        document.facturas_pro.pvp.value = "0";
        document.facturas_pro.descuento.value = "0";
        document.facturas_pro.iva.value = document.facturas_pro.defaultIva.value;
        document.facturas_pro.importe.value = "0";
        //Colocar el foco en el campo de cantidad.
        document.facturas_pro.cantidad.focus();
        document.facturas_pro.cantidad.select();
    }

    //Añade un pago a cuenta.
    function addPago(nfactura) {
        if (document.facturas_pro.importePago.value == "") document.facturas_pro.importePago.value = 0;
        if (document.facturas_pro.fechaPago.value == "") {
            window.alert("<%=LitErrFechaPago%>");
            return;
        }

        if (isNaN(document.facturas_pro.importePago.value.replace(",", "."))) {
            window.alert("<%=LitErrImportePago2%>");
            return;
        }
        else {
            if (parseFloat(document.facturas_pro.importePago.value.replace(",", ".")) == 0) {
                window.alert("<%=LitMsgImportePositivo%>");
                return;
            }
        }
        if (document.facturas_pro.descripcionPago.value == "") {
            window.alert("<%=LitMsgDesVacia%>");
            return;
        }
        if (document.facturas_pro.tipoPago.value == "") {
            window.alert("<%=LitMsgTipoPagoNoNulo%>");
            return;
        }
        if (!cambiarfecha(document.facturas_pro.fechaPago.value, "Fecha Pago")) return;
        if (!checkdate(document.facturas_pro.fechaPago)) {
            window.alert("<%=LitMsgFechaFecha%>");
            return;
        }

        //Asignar los valores a los campos del submarco de detalles
        fr_PagosCuenta.document.facturas_propago.fecha.value = document.facturas_pro.fechaPago.value;
        fr_PagosCuenta.document.facturas_propago.importe.value = document.facturas_pro.importePago.value;
        fr_PagosCuenta.document.facturas_propago.descripcion.value = document.facturas_pro.descripcionPago.value;
        fr_PagosCuenta.document.facturas_propago.medio.value = document.facturas_pro.tipoPago.value;
        //Recargar el submarco de pagos a cuenta
        fr_PagosCuenta.document.facturas_propago.action = "facturas_propago.asp?mode=first_save";
        fr_PagosCuenta.document.facturas_propago.submit();

        //Limpiar los campos del formulario
        var hoy = new Date();
        document.facturas_pro.fechaPago.value = hoy.getDate() + "/" + (hoy.getMonth() + 1) + "/" + hoy.getFullYear();
        document.facturas_pro.importePago.value = "0";
        document.facturas_pro.descripcionPago.value = "";
        document.facturas_pro.tipoPago.value = "";
        //Colocar el foco en el campo de cantidad.
        document.facturas_pro.fechaPago.focus();
        document.facturas_pro.fechaPago.select();
    }

    //Añade un pago a cuenta.
    function addVencimiento(nfactura) {
        if (document.facturas_pro.importeVto.value == "") document.facturas_pro.importeVto.value = 0;
        if (document.facturas_pro.fechaVto.value == "") {
            window.alert("<%=LitErrFechaPago%>");
            return;
        }

        if (isNaN(document.facturas_pro.importeVto.value.replace(",", "."))) {
            window.alert("<%=LitErrImportePago%>");
            return;
        }
        else {
            if (parseFloat(document.facturas_pro.importeVto.value.replace(",", ".")) == 0) {
                window.alert("<%=LitMsgImportePositivo%>");
                return;
            }
        }

        if (!cambiarfecha(document.facturas_pro.fechaVto.value, "Fecha Vencimiento")) return;

        if (!checkdate(document.facturas_pro.fechaVto)) {
            window.alert("<%=LitMsgFechaFecha%>");
            return;
        }

        //Asignar los valores a los campos del submarco de detalles
        fr_Vencimientos.document.facturas_proven.fecha.value = document.facturas_pro.fechaVto.value;
        fr_Vencimientos.document.facturas_proven.importe.value = document.facturas_pro.importeVto.value;
        fr_Vencimientos.document.facturas_proven.pagado.checked = document.facturas_pro.pagadoVto.checked;
        //Recargar el submarco de pagos a cuenta
        fr_Vencimientos.document.facturas_proven.action = "facturas_proven.asp?mode=first_save";
        fr_Vencimientos.document.facturas_proven.submit();
        //Limpiar los campos del formulario
        var hoy = new Date();
        document.facturas_pro.fechaVto.value = hoy.getDate() + "/" + (hoy.getMonth() + 1) + "/" + hoy.getFullYear();
        document.facturas_pro.importeVto.value = "0";
        document.facturas_pro.pagadoVto.checked = false;
        //Colocar el foco en el campo de cantidad.
        document.facturas_pro.fechaVto.focus();
        document.facturas_pro.fechaVto.select();
    }

    function Recalcula(total_iva_bruto, total_re_bruto) {
        document.facturas_pro.importe_bruto.value = document.facturas_pro.importe_bruto.value.replace(",", ".");
        document.facturas_pro.descuento.value = document.facturas_pro.descuento.value.replace(",", ".");
        document.facturas_pro.descuento2.value = document.facturas_pro.descuento2.value.replace(",", ".");
        document.facturas_pro.recargo.value = document.facturas_pro.recargo.value.replace(",", ".");
        document.facturas_pro.irpf.value = document.facturas_pro.irpf.value.replace(",", ".");

        document.facturas_pro.base_imponible.value = (parseFloat(document.facturas_pro.importe_bruto.value) * (100 - parseFloat(document.facturas_pro.descuento.value))) / 100;
        document.facturas_pro.base_imponible.value = (parseFloat(document.facturas_pro.base_imponible.value) * (100 - parseFloat(document.facturas_pro.descuento2.value))) / 100;
        document.facturas_pro.base_imponible.value = parseFloat(document.facturas_pro.base_imponible.value).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
        document.facturas_pro.total_descuento.value = parseFloat(document.facturas_pro.importe_bruto.value) - parseFloat(document.facturas_pro.base_imponible.value);
        document.facturas_pro.total_descuento.value = parseFloat(document.facturas_pro.total_descuento.value).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);

        document.facturas_pro.total_iva.value = (parseFloat(total_iva_bruto) * (100 - parseFloat(document.facturas_pro.descuento.value))) / 100;
        document.facturas_pro.total_iva.value = (parseFloat(document.facturas_pro.total_iva.value) * (100 - parseFloat(document.facturas_pro.descuento2.value))) / 100;
        document.facturas_pro.total_iva.value = parseFloat(document.facturas_pro.total_iva.value).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);

        //ricardo 13/5/2008 solo si el proveedor tiene recargo de equivalencia se calculara
        //FLM:20090623:Falla esta llamada. Como no estoy seguro del 100% de los casos... compruebo que no exista, y en tal caso busco otro input.
        //valor_h_re=document.facturas_pro.h_re.value;
        if (document.facturas_pro.h_re != null)
            valor_h_re = document.facturas_pro.h_re.value;
        else if (document.facturas_pro.recargo != null)
            valor_h_re = document.facturas_pro.recargo.value;
        else
            valor_h_re = 0;
        ///
        if (valor_h_re == 0) document.facturas_pro.total_re.value = 0;
        else {
            document.facturas_pro.total_re.value = (parseFloat(total_re_bruto) * (100 - parseFloat(document.facturas_pro.descuento.value))) / 100;
            document.facturas_pro.total_re.value = (parseFloat(document.facturas_pro.total_re.value) * (100 - parseFloat(document.facturas_pro.descuento2.value))) / 100;
            document.facturas_pro.total_re.value = parseFloat(document.facturas_pro.total_re.value).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
        }

        document.facturas_pro.total_recargo.value = (parseFloat(document.facturas_pro.base_imponible.value) * parseFloat(document.facturas_pro.recargo.value)) / 100;
        document.facturas_pro.total_recargo.value = parseFloat(document.facturas_pro.total_recargo.value).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);

        if (document.facturas_pro.IRPF_Total.value == "True" || document.facturas_pro.IRPF_Total.value == "1") {
            baseImp = parseFloat(document.facturas_pro.base_imponible.value) +
                parseFloat(document.facturas_pro.total_iva.value) +
                parseFloat(document.facturas_pro.total_re.value) +
                parseFloat(document.facturas_pro.total_recargo.value);
        }
        else baseImp = document.facturas_pro.base_imponible.value;
        document.facturas_pro.total_irpf.value = (parseFloat(baseImp) * parseFloat(document.facturas_pro.irpf.value)) / 100;
        document.facturas_pro.total_irpf.value = parseFloat(document.facturas_pro.total_irpf.value).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);

        document.facturas_pro.total_factura.value =
            parseFloat(document.facturas_pro.base_imponible.value) +
            parseFloat(document.facturas_pro.total_iva.value) +
            parseFloat(document.facturas_pro.total_re.value) +
            parseFloat(document.facturas_pro.total_recargo.value) -
            parseFloat(document.facturas_pro.total_irpf.value);
        document.facturas_pro.total_factura.value = parseFloat(document.facturas_pro.total_factura.value).toFixed(<%=enc.EncodeForJavascript(null_z(NdecDiFactura)) %>);
    }

    //*****************************************************************************
    function Acaja(nfactura, pendiente) {
        <%'FLM:20090505:cuando algún vencimiento está incluido en alguna remesa no se puede incluir la factura en caja.
        venConRemesa = 0
        p_nfacturaREM = limpiaCadena(Request.QueryString("nfactura"))
        if p_nfacturaREM= "" then
        p_nfacturaREM = limpiaCadena(Request.QueryString("ndoc"))
        end if
         if p_nfacturaREM & "" > "" then
        set rstAuxREM = Server.CreateObject("ADODB.Recordset")
        rstAuxREM.cursorlocation = 3
        rstAuxREM.open "select top 1 r.nremesa from remesas_pro r with(nolock) inner join detalles_rempro dr with(nolock) on dr.nremesa=r.nremesa and (dr.nfacturavto='" & p_nfacturaREM & "' ) where r.nempresa='" & session("ncliente") & "' ", session("dsn_cliente")
        if not rstAuxREM.EOF then
        venConRemesa = 1
        end if
        rstAuxREM.close
            set rstAuxREM= Nothing
        end if
        %>
        if("<%=venConRemesa%>" == "1") {
                window.alert("<%=LitMsgNoPagarInCajaFactRemesa%>");
                return;
            }
        //ricardo 11-2-2003
        //ya que cuando se insertan detalles o conceptos o pagos , no se actualiza la pagina de facturas
        //por lo que el pendiente seguia siendo cero, cuando no era verdad,por lo que no decia
        //que la factura iba a ser pagado en su totalidad, y no ponia los vencimientos a pagados
        pendiente = document.facturas_pro.h_impcaja.value;

        if (document.facturas_pro.impcaja.value == "") document.facturas_pro.impcaja.value = 0;
        if (isNaN(document.facturas_pro.impcaja.value.replace(",", "."))) {
            window.alert("<%=LitMsgImporteNumerico%>");
            return false;
        }
        else {
            if (parseFloat(document.facturas_pro.impcaja.value.replace(",", ".")) == 0) {
                window.alert("<%=LitErrImportePago%>");
                return false;
            }
        }
        pagada = "NO";
        if (parseFloat(document.facturas_pro.impcaja.value.replace(",", ".")) == parseFloat(pendiente.replace(",", "."))) {
            if (!confirm("<%=LitMsgAnotPagadaFactConfirm%>")) {
                return false;
            }
            else {
                if (document.facturas_pro.h_llekoAdmin != null && document.facturas_pro.h_llekoAdmin.value == "SI") {
                    if (!confirm("<%=LitConfirmUpdateStatus2%>")) {
                        return false;
                    }
                    else {
                        pagada = "SI";
                    }
                }
                else {
                    pagada = "SI";
                }
            }
        }
        if (document.facturas_pro.ncaja.value == "") {
            alert("<%=LitMsgCajaNoNulo%>");
            return false;
        }
        else {
            if (document.facturas_pro.i_pago.value == "") alert("<%=LitMsgTipoPagoNoNulo%>");
            else {
                fr_PagosCuenta.document.facturas_propago.action = "facturas_propago.asp?mode=acaja&ndoc=" + nfactura + "&impcaja=" + document.facturas_pro.impcaja.value + "&i_pago=" + document.facturas_pro.i_pago.value + "&ncaja=" + document.facturas_pro.ncaja.value + "&pagada=" + pagada;
                fr_PagosCuenta.document.facturas_propago.submit();
                setTabsSelected(3);
            }
        }
    }

    //Genera los vencimientos de la factura.
    function genVencimiento(nfactura) {
        fr_Vencimientos.document.facturas_proven.action = "facturas_proven.asp?mode=create";
        fr_Vencimientos.document.facturas_proven.submit();
    }

    function seguimientoCobros(viene, nproveedor, ndocumento) {
        AbrirVentana("../central.asp?pag1=administracion/seguimientoCobros.asp&mode=imp&nproveedor=" + trimCodEmpresa(nproveedor) + "&viene=" + viene + "&ndoc=" + ndocumento + "&pag2=administracion/seguimientoCobros_bt.asp", 'P',<%=AltoVentana %>,<%=AnchoVentana %>);
    }

    function facturasVinculadas(viene, ncliente, ndocumento) {
        AbrirVentana("/ilionx45/Custom/RepsolPeru/backoffice/RepsolFacturasProVinculadasNC.aspx?ndoc=" + ndocumento + "&ncliente=" + trimCodEmpresa(ncliente), 'P',<%=AltoVentana %>, <%=AnchoVentana %>);
    }

    //****************************************************************************
    function MasDet(sentido, lote, firstReg, lastReg, campo, criterio, texto, firstRegAll, lastRegAll) {
        fr_Detalles.document.facturas_prodet.action = "facturas_prodet.asp?mode=browse&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&firstReg=" + firstReg + "&lastReg=" + lastReg + "&firstRegAll=" + firstRegAll + "&lastRegAll=" + lastRegAll + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
        fr_Detalles.document.facturas_prodet.submit();
    }
    //**EBF 10/3/2008 Añadir albaranes ********************************************************
    function anyadirAlbaranes(viene, ncliente, ndocumento) {
        AbrirVentana("../central.asp?pag1=compras/albpedpro_facpro_param.asp&mode=select1&nproveedor=" + trimCodEmpresa(ncliente) + "&viene=" + viene + "&ndoc=" + ndocumento + "&pag2=compras/albpedpro_facpro_param_bt.asp", 'P',<%=AltoVentana %>,<%=AnchoVentana %>);
    }

    /*RGU 6/4/2009: Contabilizar una única factura*/
    function ContabilizarFra(nfactura, serie, fecha) {
        if (typeof (fr_Enlace.document.EnlaceContableFra.sin_activo) == "object") alert("<%=LitMsgNoEjercicioActivo %>")
        else {
            if (window.confirm("<%=LitEnlazar1%> " + fr_Enlace.document.EnlaceContableFra.nombreEmpresa.value + " <%=LitEnlazar2%> " + fr_Enlace.document.EnlaceContableFra.ejercicio.value + "<%=LitEnlazar3%>")) {
                fr_Enlace.document.EnlaceContableFra.action = "EnlaceContableFra.asp?mode=enlace&nfactura=" + nfactura + "&nserie=" + serie + "&fecha_fac=" + fecha;
                fr_Enlace.document.EnlaceContableFra.submit();
            }
        }
    }

    //AMP 08102010: Incluimos campo factor de cambio en la cabecera de los presupuestos.
    var ret_tra = "";
    var ret_tra2 = "";
    function cambiardivisa(mBase) {
        document.facturas_pro.divisa.value = document.facturas_pro.divisabis.value;

        var divisa = document.facturas_pro.divisa.value;
        if (divisa == mBase) {
            parent.pantalla.document.getElementById("tdfactcambio").style.display = "none";
            parent.pantalla.document.facturas_pro.nfactcambio.value = "1";
        }
        else {
            parent.pantalla.document.getElementById("tdfactcambio").style.display = "";
            ret_tra = "";
            if (!enProceso && http) {
                var timestamp = Number(new Date());
                var url = "../select_factcambio.asp?divisa=" + divisa;
                http.open("GET", url, false);
                http.onreadystatechange = handleHttpResponse;
                enProceso = false;
                http.send(null);
            }

        }
    }

    function handleHttpResponse() {
        if (http.readyState == 4) {
            if (http.status == 200) {
                if (http.responseText.indexOf('invalid') == -1) {
                    // Armamos un array, usando la coma para separar elementos
                    results = http.responseText;
                    enProceso = false;
                    ret_tra = unescape(results);
                    spfc = ret_tra.split(";");
                    factcambio = spfc[0];
                    parent.pantalla.document.facturas_pro.nfactcambio.value = factcambio;
                    ret_tra2 = "";
                    var divisa = document.facturas_pro.divisa.value;
                    if (!enProceso2 && http2) {
                        var timestamp = Number(new Date());
                        var url = "../select_factcambio.asp?divisa=" + divisa + "&que=abreviatura";
                        http2.open("GET", url, false);
                        http2.onreadystatechange = handleHttpResponse2;
                        enProceso2 = false;
                        http2.send(null);
                    }
                }
            }
        }
    }

    function handleHttpResponse2() {
        if (http2.readyState == 4) {
            if (http2.status == 200) {
                if (http2.responseText.indexOf('invalid') == -1) {
                    // Armamos un array, usando la coma para separar elementos
                    results = http2.responseText;
                    enProceso2 = false;
                    ret_tra2 = unescape(results);
                    spfc = ret_tra2.split(";");
                    otraabrev = spfc[0];
                    parent.pantalla.document.getElementById("idfactcambioexpl").innerHTML = otraabrev;
                }
            }
        }
    }

    function getHTTPObject() {
        var xmlhttp;
        if (!xmlhttp && typeof XMLHttpRequest != 'undefined') {
            try {
                xmlhttp = new XMLHttpRequest();
            }
            catch (e) {
                xmlhttp = false;
            }
        }
        return xmlhttp;
    }
    var enProceso = false; // lo usamos para ver si hay un proceso activo
    var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest
    var enProceso2 = false; // lo usamos para ver si hay un proceso activo
    var http2 = getHTTPObject(); // Creamos el objeto XMLHttpRequest

    function comprobarFactorCambio() {
        ok = 1;
        numero = document.facturas_pro.nfactcambio.value;
        document.facturas_pro.nfactcambio.value = numero.replace(",", ".")
        numero2 = document.facturas_pro.nfactcambio.value;
        if (!/^([0-9])*[.]?[0-9]*$/.test(numero2)) {
            alert("<%=LitMsgFactCambioI%>");
            ok = 0;
        }
        if (document.facturas_pro.nfactcambio.value == "") {
            alert("<%=LitMsgFactCambioI%>");
            ok = 0;
        }
    }
    //*** f AMP

    function Redimensionar() {
        var alto = jQuery(window).height();
        var diference = 425;
        var dir_default = 150;

        if (alto > dir_default) {
            if (alto - diference > dir_default) {
                jQuery("#frDetalles").attr("height", alto - diference);
                jQuery("#frConceptos").attr("height", alto - diference);
                jQuery("#frPagosCuenta").attr("height", alto - diference);
                jQuery("#frVencimientos").attr("height", alto - diference);
            }
            else {
                jQuery("#frDetalles").attr("height", dir_default);
                jQuery("#frConceptos").attr("height", dir_default);
                jQuery("#frPagosCuenta").attr("height", dir_default);
                jQuery("#frVencimientos").attr("height", dir_default);
            }
        }
        else {
            jQuery("#frDetalles").attr("height", dir_default);
            jQuery("#frConceptos").attr("height", dir_default);
            jQuery("#frPagosCuenta").attr("height", dir_default);
            jQuery("#frVencimientos").attr("height", dir_default);
        }
    }

    function RoundNumValue(obj, dec) {
        obj.value = obj.value.replace(',', '.');
        var valor = parseFloat(obj.value);
        if (valor != 0) obj.value = valor.toFixed(dec);
    }

    function ChangeCloseMode() {
        jQuery('.window .close').click(function (e) {
            jQuery('#mask').hide();
            jQuery('.window').hide();
            SelectCloseMode()

        });
        jQuery('#mask').click(function () {
            jQuery('#mask').hide();
            jQuery('.window').hide();
            SelectCloseMode()
        });
    }

    function SelectCloseMode() {
        jQuery('.window .close').click(function (e) {
            jQuery('#mask').hide();
            jQuery('.window').hide();

        });
        jQuery('#mask').click(function () {
            jQuery('#mask').hide();
            jQuery('.window').hide();
        });

        ninvoice = document.facturas_pro.h_nfactura.value;
        document.facturas_pro.action = "facturas_pro.asp?mode=browse&nfactura=" + ninvoice;
        document.facturas_pro.submit();
    }

    function OpenModalWindow(ninvoice, facturapro, base_imp) {
        loadModal = 0
        codPr = document.facturas_pro.cod_proyecto.value;

        if (codPr > "") {
            if (!window.confirm("<%=LITINVOICEHASPROJECT %>")) {
                return;
            }
            else {
                if (!UnlinkProjectAJAX()) {
                    alert("<%=LITERRORUNLINKPROJECT %>");
                    return;
                }
                else {
                    loadModal = 1
                    ChangeCloseMode();
                }
            }
        }

        var ran = Math.random();
        var alto = 0;
        var ancho = 0;
        ancho = jQuery(window).width();
        alto = jQuery(window).height();
        ancho = ancho - (ancho * 50 / 100);
        alto = alto - (alto * 30 / 100);
        paginaModal = "../mantenimiento/projects_costdocs.asp?ninvoice=" + ninvoice + "&facturapro=" + facturapro + "&base_imp=" + document.getElementById("base_imponible").innerHTML.toString().replace(".", "") + "&typedoc=0&loadModal=" + loadModal.toString();
        cambiarTamanyo("#ProjectsCost", "300", "700");
        reloadIframe("#ProjectsCost", "");
        reloadClass("#ProjectsCost", paginaModal);
        alPresionar("#ProjectsCost");
    }

    function UnlinkProjectAJAX() {
        result = "";
        var ndocument = document.facturas_pro.nfactura.value;
        var ncompany = ndocument.substring(0, 5);

        if (!enProcesoUn && httpUn) {
            var timestamp = Number(new Date());
            var url = "../mantenimiento/projects_UnlinkDocAJAX.asp?ncompany=" + ncompany + "&ndoc=" + ndocument + "&typedoc=0";
            httpUn.open("POST", url, false);
            httpUn.onreadystatechange = handleHttpResponse3;
            enProcesoUn = false;
            httpUn.send(null);
        }

        if (result == 0)
            return true;
        else
            return false;
    }

    function handleHttpResponse3() {
        if (httpUn.readyState == 4) {
            if (httpUn.status == 200) {
                if (httpUn.responseText.indexOf('invalid') == -1) {
                    // Armamos un array, usando la coma para separar elementos
                    results = httpUn.responseText;
                    enProcesoUn = false;
                }
            }
        }
        enProcesoUn = false;
    }

    var enProcesoUn = false; // lo usamos para ver si hay un proceso activo
    var httpUn = getHTTPObject(); // Creamos el objeto XMLHttpRequest

    jQuery(window).resize(function () { Redimensionar(); });
</script>
<body onLoad="self.status='';" class="BODY_ASP">
<%
     

'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(nfactura,nserie)
	ModDocumentoEquip=false
	ModDocumento=true

	if len(limpiaCadena(request.form("nproveedor")))<=5 then
		nproveedor=session("ncliente") & completar(limpiaCadena(request.form("nproveedor")),5,"0")
	else
		nproveedor=completar(limpiaCadena(request.form("nproveedor")),5,"0")
	end if

	'Miramos si el proveedor tiene recargo de equivalencia
    rstAux.cursorlocation=3
	rstAux.open "Select re from proveedores with(nolock) where nproveedor='" + Nulear(nproveedor) + "'",session("dsn_cliente")
	if not rstAux.eof then
		TieneRE=rstAux("re")
	else
		TieneRE=0
	end if
	rstAux.close

	if nfactura="" and d_lookup("count(*)","facturas_pro","nfactura_pro='" & Nulear(limpiaCadena(Request.Form("nfactura_pro")))&"' and nproveedor='" & Nulear(nproveedor)&"'  and year(fecha)= year(convert (datetime,'" & Nulear(limpiaCadena(Request.Form("fecha"))) & "' ))",session("dsn_cliente"))<>"0" then

	else
		if nfactura="" then
			'Obtener el último nº de factura de la tabla series.
            rstAux.cursorlocation=2
			rstAux.Open "select contador,ultima_fecha from series with(updlock) where nserie='" & nserie & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

			num=rstAux("contador")+1
            num=string(5-len(cstr(num)),"0") + cstr(num)

			'Actualizar el nº de proveedor de CONFIGURACION.
			rstAux("contador")=rstAux("contador")+1
			rstAux("ultima_fecha")=date
			rstAux.Update
			rstAux.Close
            
            'Crear un nuevo registro.
			rst.AddNew
            rst("nfactura")=nserie + right(Nulear(limpiaCadena(Request.Form("fecha"))),2) + num
			SigDoc=rst("nfactura")
			nfactura=SigDoc
			'******************** Manejo de domicilios
			Dom=Domicilios("COMPRAS","FAC_ENV_PROV",nproveedor,rst)
		else
			mensajeTratEquipos="OK"
			if (rst("nproveedor")<>nproveedor) or (rst("fecha")<>cdate(limpiaCadena(Request.Form("fecha"))&"")) then
				ModDocumentoEquip=true
			end if
			if mid(mensajeTratEquipos,1,2)<>"OK" then
				ModDocumento=false
				%><script language="javascript" type="text/javascript">
                      window.alert("<%=enc.EncodeForJavascript(mensajeTratEquipos)%>");
				</script><%
			else
				ModDocumento=true
			end if
		end if
		if ModDocumento then
		    ndec=d_lookup("ndecimales", "divisas", "codigo like '"&session("ncliente")&"%' and codigo='"&limpiaCadena(request.form("divisa"))&"'", session("dsn_cliente"))
			FechaDoc=rst("fecha")
			ProvDoc=rst("nproveedor")
			DtoGeneral=null_z(rst("descuento"))
			DtoGeneral2=null_z(rst("descuento2"))
			'Asignar los nuevos valores a los campos del recordset.
			rst("nfactura_pro") = Nulear(limpiaCadena(Request.Form("nfactura_pro")))
			rst("serie")		= Nulear(limpiaCadena(Request.Form("serie")))
			cambio_proveedor = false
		   	if rst("nproveedor")<>nproveedor then cambio_proveedor=true
			rst("nproveedor")	= Nulear(nproveedor)
			rst("fecha")		= Nulear(limpiaCadena(Request.Form("fecha")))
			rst("forma_pago")	= Nulear(limpiaCadena(request.form("forma_pago")))
			rst("descuento")	= miround(null_z(limpiaCadena(request.form("descuento"))),decpor)
			rst("descuento2")	= miround(null_z(limpiaCadena(request.form("descuento2"))),decpor)
			rst("descuento3")	= miround(null_z(limpiaCadena(request.form("descuento3"))),decpor)
			rst("recargo")		= miround(null_z(limpiaCadena(request.form("recargo"))),decpor)
			rst("irpf")			= miround(null_z(limpiaCadena(request.form("irpf"))),decpor)
			rst("IRPF_Total")	= nz_b(limpiaCadena(Request.Form("IRPF_Total")))
			rst("observaciones")= Nulear(limpiaCadena(request.form("observaciones")))
			rst("contabilizado")= nz_b(limpiaCadena(request.form("contabilizada")))
			rst("cod_proyecto")=Nulear(limpiaCadena(request.form("cod_proyecto")))
			rst("ncuenta")=Nulear(limpiaCadena(request.form("ncuentacargo")))
			'FLM:120309:cuenta de abono del proveedor y el nombre del banco.
			rst("ncuenta_pro")=Nulear(limpiaCadena(request.form("ncuenta_pro")))
	        if(rst("ncuenta_pro")&""<>"") then 
                banco=d_lookup("Entidad","bancos","codigo='" & mid(trim(rst("ncuenta_pro")),5,4) & "'",DsnIlion)
            else
                banco=null
            end if
            rst("banco")=iif(banco="",NULL,trim(banco))
			rst("incoterms")=nulear(limpiaCadena(request.form("incoterms")))
			rst("fob")=nulear(limpiaCadena(request.form("fob")))
			''MPC 28/11/2007 guardar el campo tienda en las facturas de proveedores
			rst("tienda")=nulear(limpiaCadena(request.form("tienda")))
			''ricardo 28/4/2003 si el usuario ha querido recalcular los importes al cambiar las propiedades de la cabecera
			han_cambiado_importes_proveedor=0
			if limpiaCadena(request.querystring("recalcular_importes"))="1" then
				'Detectamos un cambio de proveedor
				if cambio_proveedor=true then
					han_cambiado_importes_proveedor=1
					set rstMiProveer = Server.CreateObject("ADODB.Recordset")
					'recorremos los detalles modificando precios
                    rstAux.cursorlocation=2
					rstaux.open "select * from detalles_fac_pro with(updlock) where nfactura='" & nfactura & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					while not rstAux.eof
                        rstMiProveer.cursorlocation=3
						rstMiProveer.open "select * from proveer with(nolock) where nproveedor='" & rst("nproveedor") & "' and articulo='" & rstaux("referencia") & "'",session("dsn_cliente")
						if not rstMiProveer.eof then
							TmpPvp=CambioDivisa(rstMiProveer("importe"), rstMiProveer("divisa"),iif(tmp_divisa>"",tmp_divisa,limpiaCadena(request.form("divisa"))))
							impOrig=TmpPvp
							TmpDto=null_z(rstMiProveer("descuento"))
							TmpDto2=null_z(rstMiProveer("descuento2"))
							TmpPVP	= (TmpPVP*rstAux("cantidad"))*(100-null_z(TmpDto))/100
							TmpPVP	= TmpPVP*(100-null_z(TmpDto2))/100
							TmpImporte=miround(TmpPVP,2)
						else
   							TmpPvp=0
							impOrig=0
							TmpDto=0
							TmpDto2=0
							TmpPVP	= 0
							TmpImporte=0
						end if
						rstAux("pvp")=impOrig
						rstAux("descuento")=TmpDto
						rstAux("descuento2")=null_z(TmpDto2)
						rstAux("importe")=TmpImporte
						rstAux.update
						rstAux.movenext
						rstMiProveer.close
					wend
					rstAux.close
					Set rstMiProveer=nothing

					'recorremos ahora los conceptos haciendo cambio de divisa si hace falta
                    rstaux.cursorlocation=2
					rstaux.open "select * from conceptos_fac_pro with(updlock) where nfactura='" & nfactura & "' order by nconcepto",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					while not rstAux.eof
						TmpPVP=CambioDivisa(rstAux("pvp"),rst("divisa"),iif(tmp_divisa>"",tmp_divisa,limpiaCadena(request.form("divisa"))))
						rstAux("pvp")=TmpPVP
						TmpPVP=TmpPVP*rstAux("cantidad")
						TmpPVP	= TmpPVP*(100-null_z(rstAux("descuento")))/100
						rstAux("importe")=miround(TmpPVP,2)
						rstAux.update
						rstAux.movenext
					wend
					rstAux.close
				end if
			end if

			rst("divisa")		= Nulear(limpiaCadena(request.form("divisa")))
			rst("factcambio")=miround(Nulear(limpiaCadena(request.form("nfactcambio"))),DEC_PREC) '*** AMP

			if limpiaCadena(request.form("pagada"))="" and rst("pagada")<>0 then
				'Comprobar la caja.
                rstSelect.cursorlocation=3
				rstSelect.open "select ndocumento from caja with(nolock) where ndocumento='" & nfactura & "' or (tdocumento='VENCIMIENTO_ENTRADA' and ndocumento in (select nfactura+'-'+cast(nvencimiento as varchar(10)) from vencimientos_entrada where nfactura='" & nfactura & "'))",session("dsn_cliente")
				if rstSelect.EOF then
					rstSelect.close
					EnCaja="NO"
					rst("pagada")=0
					rstAux.Open "update vencimientos_entrada with(updlock) set pagado = 0 where nfactura='"+rst("nfactura")+"'" , session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				else
					EnCaja="SI"
					rstSelect.close
					if pagada_uxa=0 then
						%><script language="javascript" type="text/javascript">
                              window.alert("<%=LitMsgNoAnularPago%>");
						</script><%
					end if
					pagada_uxa=1 'para que no saque el mensaje de que no se puede modificar la factura
				end if
			end if
			rst("validada")=0
			rst("ahora")=0
			rst("tipo_pago")=Nulear(limpiaCadena(request.form("tipo_pago")))

			'' JMA 20/12/04 Actualizamos los campos personalizables
			num_campos=limpiaCadena(request.querystring("num_campos"))
			if num_campos="" then
				num_campos=limpiaCadena(request.form("num_campos"))
			end if
			if num_campos & "">"" then
				redim lista_valores(num_campos+10)
				for ki=1 to num_campos
					nom_campo="campo" & ki
					valor_form=Nulear(limpiaCadena(request.querystring(nom_campo)))
					if valor_form & ""="" then
						valor_form=Nulear(limpiaCadena(request.form(nom_campo)))
					end if
					if ki<10 then
						tipo_campo_perso=d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "0" & ki & "' and tabla='DOCUMENTOS COMPRA'",session("dsn_cliente"))
					else
						tipo_campo_perso=d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & ki & "' and tabla='DOCUMENTOS COMPRA'",session("dsn_cliente"))
					end if
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
				redim lista_valores(10+5)
				for ki=1 to 15
					lista_valores(ki)=""
				next
			end if

			rst("campo01")=lista_valores(1)
			rst("campo02")=lista_valores(2)
			rst("campo03")=lista_valores(3)
			rst("campo04")=lista_valores(4)
			rst("campo05")=lista_valores(5)
			rst("campo06")=lista_valores(6)
			rst("campo07")=lista_valores(7)
			rst("campo08")=lista_valores(8)
			rst("campo09")=lista_valores(9)
			rst("campo10")=lista_valores(10)
			'' JMA 20/12/04 Fin actualizar campos personalizables

			rst.update
			rst.close

		if han_cambiado_importes_proveedor=1 then
			TmpIvaProveedor=d_lookup("iva","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente"))
			TmpReProveedor=d_lookup("re","tipos_iva","tipo_iva='" & TmpIvaProveedor & "'",session("dsn_cliente"))
			defaultIVA=d_lookup("iva","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
			TmpReDefaultIva=d_lookup("re","tipos_iva","tipo_iva='" & defaultIVA & "'",session("dsn_cliente"))
			if TmpIvaProveedor & "">"" then
				TmpIva=TmpIvaProveedor
				TmpRe=TmpReProveedor
			else
				TmpIva=defaultIVA
				TmpRe=TmpReDefaultIva
			end if
            rstaux.cursorlocation=2
			rstaux.open "select d.iva,d.re,d.referencia from detalles_fac_pro as d with(updlock) where d.nfactura='" & nfactura & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			while not rstAux.eof
				if TmpIva & "">"" then
					rstAux("iva")=TmpIva
					rstAux("re")=TmpRe
				else
					Tmpivaart=d_lookup("iva","articulos","referencia='" & rstAux("referencia") & "'",session("dsn_cliente"))
					TmpReivaart=d_lookup("re","tipos_iva","tipo_iva='" & Tmpivaart & "'",session("dsn_cliente"))
					if Tmpivaart & "">"" then
						rstAux("iva")=Tmpivaart
						rstAux("re")=TmpReivaart
					else
						rstAux("iva")=defaultIVA
						rstAux("re")=TmpReDefaultIva
					end if
				end if
				rstAux.update
				rstAux.movenext
			wend
			rstAux.close
            rstaux.cursorlocation=2
			rstaux.open "select * from conceptos_fac_pro with(updlock) where nfactura='" & nfactura & "' order by nconcepto",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			while not rstAux.eof
				rstAux("iva")=TmpIva
				rstAux("re")=TmpRe
				rstAux.update
				rstAux.movenext
			wend
			rstAux.close
		end if
            rst.cursorlocation=2
		    rst.open "select * from facturas_pro with(updlock) where nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

			n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))
			if n_decimales = "" then
				n_decimales = 0
			end if

			rst("importe_bruto")	= 0
			rst("base_imponible")	= 0
			rst("total_descuento")	= 0
			rst("total_iva")		= 0
			rst("total_re")			= 0
			rst("total_recargo")	= 0
			rst("total_irpf")	= 0

			SumImporteBruto=0
			SumTotalDto=0
			SumBaseImponible=0
			SumTotalIva=0
			SumTotalIvaTotal=0
			SumTotalRE=0
			SumTotalRETotal=0
			SumTotalRF=0
			SumTotalIRPF=0
			SumTotalImporte=0

			seleccion="select sum(importe) as suma, iva, re from detalles_fac_pro with(nolock) "
			seleccion=seleccion+"where nfactura ='"+rst("nfactura")+"' "
			seleccion=seleccion+"GROUP BY IVA, RE "
			seleccion=seleccion+" union all "
			seleccion=seleccion+"select sum(importe) as suma, iva, re from conceptos_fac_pro with(nolock) "
			seleccion=seleccion+"where nfactura ='"+rst("nfactura")+"' "
			seleccion=seleccion+"GROUP BY IVA, RE ORDER BY IVA"
            rstAux.cursorlocation=3
			rstAux.open seleccion,session("dsn_cliente")
			if not rstAux.EOF then
				ivaAnt=rstAux("iva")
			end if
			while not rstAux.EOF
				if rstAux("iva")<>ivaAnt then
					SumTotalIvaTotal=SumTotalIvaTotal+miround(null_z(SumTotalIva),n_decimales)
					SumTotalRETotal=SumTotalRETotal+miround(null_z(SumTotalRE),n_decimales)
					SumTotalIva=0
					SumTotalRE=0
					ivaAnt=rstAux("iva")
				end if
				SumImporteBruto=SumImporteBruto + rstAux("suma")
				dto1=miround((null_z(rstAux("suma"))*null_z(rst("descuento")))/100,2)
				dto2=miround(((null_z(rstAux("suma"))-dto1)*null_z(rst("descuento2")))/100,2)
				total_descuento=dto1+dto2+dto3
				SumTotalDto=SumTotalDto + null_z(total_descuento)
				base_imponible=null_z(rstAux("suma"))-null_z(total_descuento)
				SumBaseImponible=SumBaseImponible + null_z(base_imponible)

				total_iva=(null_z(base_imponible)*rstAux("iva"))/100
				SumTotalIva=SumTotalIva + null_z(total_iva)
				if TieneRE <> 0 then
					re=d_lookup("re","tipos_iva","tipo_iva='" & rstAux("iva") & "'",session("dsn_cliente"))
				else
					re=0
				end if
				total_re=(null_z(base_imponible)*re)/100
				SumTotalRE=SumTotalRE + null_z(total_re)
				rstAux.moveNext
			wend
			SumTotalIvaTotal=SumTotalIvaTotal+miround(null_z(SumTotalIva),n_decimales)
			SumTotalRETotal=SumTotalRETotal+miround(null_z(SumTotalRE),n_decimales)

			rstAux.close

			rst("importe_bruto")=SumImporteBruto
			rst("total_descuento")=SumTotalDto
			rst("base_imponible")=SumBaseImponible
			rst("total_iva")=miround(SumTotalIvaTotal,n_decimales)
			rst("total_re")=miround(SumTotalRETotal,n_decimales)
			SumTotalRF=(null_z(SumBaseImponible)*null_z(rst("recargo")))/100
			rst("total_recargo")=miround(SumTotalRF,n_decimales)
			if nz_b(rst("IRPF_Total"))=0 then
				baseImp=null_z(SumBaseImponible)
			else
				baseImp=null_z(SumBaseImponible)+null_z(rst("total_iva"))+null_z(rst("total_re"))+null_z(rst("total_recargo"))
			end if
			SumTotalIRPF=(null_z(baseImp)*null_z(rst("irpf")))/100
			rst("total_irpf")=miround(SumTotalIRPF,n_decimales)
			SumTotalImporte=null_z(SumBaseImponible)+null_z(rst("total_iva"))+null_z(rst("total_re"))+null_z(rst("total_recargo"))-null_z(rst("total_irpf"))
			rst("total_factura")=miround(SumTotalImporte,2)

			rst.Update

			if limpiaCadena(request.form("pagada")) & "">"" and rst("pagada")=0 then
                rstAux.cursorlocation=2
				rstAux.Open "update vencimientos_entrada with(updlock) set pagado=1 where pagado=0 and nfactura='"+rst("nfactura")+"'" , session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				rstAux.open "update facturas_pro  with(updlock) set pagada=1 where nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			end if

			ActualizaCostes iif(nfactura="",SigDoc,nfactura),"DOCUMENTO","FACTURA DE PROVEEDOR","",ProvDoc,0,FechaDoc,0,0,DtoGeneral,DtoGeneral2,false,session("dsn_cliente")

			if ModDocumentoEquip=true then
				InsertarHistorialNserie "OK1","","","FACTURA DE PROVEEDOR",iif(nfactura="",SigDoc,nfactura),"","","","","MODIFY",mode
			end if
		end if
	end if
end sub

'Elimina los datos del registro cuando se pulsa BORRAR.
sub BorrarRegistro(nfactura)
    'i *** AMP 17092010 -- Restricciones de borrado si existe lote asignado
    rstAux.cursorlocation=3
    rstAux.Open "select * from lotes_entrada with(nolock) where nfactura='" & nfactura& "'" ,session("dsn_cliente")
	noborrar=false ' valor inicial
	nalbaranLote=""
	if not rstAux.eof then 'Si no existe lote asociado a la línea saltar restricción de borrado
	    while not rstAux.EOF
	        if rstAux("nalbaran")>"" then nalbaranLote=rstAux("nalbaran")
            ndetLote = rstAux("ndet")
            nfacturaLote=rstAux("nfactura")
            ndetfraLote=rstAux("ndetfra")
            if ndetfraLote>"" then
                strSelect="ComprobarUsoLote '"&session("ncliente")&"','"&nfactura&"','"&ndetfraLote&"',''"
                rstMM.Open strSelect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                if not rstMM.eof then
                    usoLote=rstMM(1)
                    rstMM.close
                    if usoLote=1 and nalbaranLote=""  then 'si se utiliza lote y no existe un albaran vinculado a este lote --> no borrar línea
                        noborrar=true
                        rstAux.MoveLast
                    end if
                end if
            end if
            rstAux.MoveNext
        wend
	end if
	rstAux.close
    'fin *** AMP 17092010 -- Restricciones de borrado

	if noborrar then
	     %><script language="javascript" type="text/javascript">
                              window.alert("<%=LitMsgNoBorrarPorLote%>");
                              document.facturas_pro.action = "facturas_pro.asp?nfactura=<%=enc.EncodeForJavascript(nfactura)%>&mode=browse&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
                              document.facturas_pro.submit();
                              parent.botones.document.location = "facturas_pro_bt.asp?mode=browse";
	     </script><%
	else
	    'obtenemos las referencias de los artículos para el control de stock
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open session("dsn_cliente")
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = cn
        cmd.CommandText = "ObtenerReferencias"
        cmd.CommandType = adCmdStoredProc

        cmd.Parameters.Append cmd.CreateParameter("ndocumento", adChar, adParamInput,20)
        cmd.Parameters.Append cmd.CreateParameter("tipo_documento", adChar, adParamInput,50)
        cmd("ndocumento") = nfactura
        cmd("tipo_documento") = "FACTURA DE PROVEEDOR"
        set referencias=cmd.Execute

        'borramos la factura
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open session("dsn_cliente")
        Set cmd1 = Server.CreateObject("ADODB.Command")
        Set cmd1.ActiveConnection = cn
        cmd1.CommandText = "BorrarDocumento"
        cmd1.CommandType = adCmdStoredProc

        cmd1.Parameters.Append cmd1.CreateParameter("result", adInteger, _
        adParamReturnValue)
        cmd1.Parameters.Append cmd1.CreateParameter("ndocumento", adChar, _
          adParamInput,20)
        cmd1.Parameters.Append cmd1.CreateParameter("tipo_documento", adChar, _
             adParamInput,50)

        cmd1("ndocumento") = nfactura
        cmd1("tipo_documento") = "FACTURA DE PROVEEDOR"
        set result=cmd1.Execute

        if result("result") = 0 then
            'Mostramos el control de stock
             if session("control_stock")="activado"  then 'CONTROL DE STOCK A NIVEL DE EMPRESA
                while not referencias.EOF
		            Stock=referencias("stock")
		            StockMin=referencias("stock_minimo")
		            PendRecibir=referencias("p_recibir")
		            PendServir=referencias("p_servir")
		            articulo_stock =referencias("control_stock")
		            referencia =referencias("referencia")
	                if articulo_stock <> 0 then 'CONTROL DE STOCK A NIVEL DE ARTICULO
		                if Stock<0 then%>
			                <script language="javascript" type="text/javascript">alert("<%=LitMsgStockNegativo%> <%=trimCodEmpresa(referencia)%>");</script>
			            <%elseif Stock < StockMin then%>
			                <script language="javascript" type="text/javascript">
                                window.alert("<%=LitMsgStockBajoMin%> <%=trimCodEmpresa(referencia)%>");
			                </script>
			            <%end if
		                if PendRecibir<0 then%>
			                <script language="javascript" type="text/javascript">alert("<%=LitMsgStockPRNegativo%> <%=trimCodEmpresa(referencia)%>");</script>
			            <%end if
		                if PendServir<0 then%>
				            <script language="javascript" type="text/javascript">alert("<%=LitMsgStockPSNegativo%> <%=trimCodEmpresa(referencia)%>");</script>
			            <%end if
	                end if 	'stock articulo'

		            referencias.movenext
	            wend
            end if	'stock empresa'
        else
            ' errores
             mensaje= result("desc_error")

            if result("result") = 1 then
                 mensaje= LITERRORBORRADO
            end if
             if result("result") = 2 then
                 mensaje= LITRELACIONDOCUMENTOS
            end if
             if result("result") = 3 then
                 mensaje= LITERRORBORRARDETALLES
            end if
             if result("result") = 4 then
                 mensaje= LITERRORBORRARPARAM
            end if
             if result("result") = 5 then
                 mensaje= rLITERRORBORRARVENCIM
            end if
             if result("result") = 6 then
                 mensaje= LITERRORBORRARPAGOS
            end if%>
            <script language="javascript" type="text/javascript">alert("<%=enc.EncodeForJavascript(mensaje)%>");</script>
        <%end if
    end if
end sub

'Crea la tabla que contiene la barra de grupos de datos.
sub BarraNavegacion(modo,si_tiene_acceso_detalles)
	if mode="browse" then%>
        <script language="javascript" type="text/javascript">
            //document.getElementById("CABECERA").style.display = "none";
            jQuery("#S_CABECERA").hide();
        </script>
    <%end if
    if modo<>"add" and modo<>"edit" then%>
        <script language="javascript" type="text/javascript">
            jQuery("#S_DATFINAN").hide();
        </script>
	<%end if
end sub

'ega 13/03/2008
function CerrarTodo()
    set rsTPV = nothing
    set rstAuxREM = nothing
    set rstMiProveer = nothing
    set cn  = nothing
    set cn = nothing
    set cn = nothing
    set rstDet= nothing
    set cmd = nothing
    set cmd1 = nothing
    set rstAux = nothing
    set rstAux2 = nothing
    set rstAux3 = nothing
    set rst = nothing
    set rst2 = nothing
    set rstSelect = nothing
    set rstDomi = nothing
    set rstConta = nothing
    set rstMM = nothing 
    set rstObtDocCli = nothing
    set comm_mx = nothing
    set conn = nothing
    set referencias = nothing
    connRound.close
    set connRound = Nothing
end function

'FLM:200309
function ComprobarCuenta(ncuenta)
	if (ncuenta & "">"") then
		cuenta=trim(ncuenta)
		cuentaOK=true
        strIban= Mid(cuenta, 1, 2)
        strPais= Mid(cuenta, 3, 2)
		strBanco = Mid(cuenta, 5, 4)
		strOficina = Mid(cuenta, 9, 4)
		strDC1 = Mid(cuenta, 13, 1)
		strDC2 = Mid(cuenta, 14, 1)
		strCuenta = Mid(cuenta, 15, 10)
		if cuenta="" then
			CuentaOK=false
		else
			prueba1=Validar_cuenta(strBanco & strOficina, strDC1, False, strDC1Bueno)
			prueba2=Validar_cuenta(strCuenta, strDC2, True, strDC2Bueno)

			If prueba1=true and prueba2=true Then
				CuentaOK=true
			Else
				CuentaOK=false
			End If
		end if
	else
		if (ncuenta & ""="") then
			CuentaOK=true
		else
			CuentaOK=false
		end if
	end if
	ComprobarCuenta=CuentaOK
end function

'******************************************************************************************************
        '********actualizar collection recipient de factura
'FLM:200309
function updateFieldCollectionRecInvoice(oldCollectionRecipientId,Getfield,nfactura)

    'Actualizamos el estado del collectionRecipientId
    set rstUpdInv = Server.CreateObject("ADODB.Recordset")
    set connUpdInv = Server.CreateObject("ADODB.Connection")
    set commandUpdInv =  Server.CreateObject("ADODB.Command")
    connUpdInv.open session("dsn_cliente")
    commandUpdInv.ActiveConnection =connUpdInv
    commandUpdInv.CommandTimeout = 0
    commandUpdInv.CommandText= "INVOICE_UPDATE_FIELD"
    commandUpdInv.CommandType = adCmdStoredProc                    
    commandUpdInv.Parameters.Append commandUpdInv.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))                            
    commandUpdInv.Parameters.Append commandUpdInv.CreateParameter("@collectionRecipientId", adInteger, adParamInput, 25, oldCollectionRecipientId)
    commandUpdInv.Parameters.Append commandUpdInv.CreateParameter("@updateField", adInteger, adParamInput, 25, Getfield)
    commandUpdInv.Parameters.Append commandUpdInv.CreateParameter("@invoice", adVarChar, adParamInput, 20, nfactura)
    
    commandUpdInv.Execute,,adExecuteNoRecords
    set rstUpdInv = commandUpdInv.Execute
    if not rstUpdInv.eof then
        update_Invoice_ColleRecId = rstUpdInv("ERROR_CODE")
    else
        update_Invoice_ColleRecId = "-1"
    end if
    if rstUpdInv.state<>0 then rstUpdInv.close
    connUpdInv.close
    set connUpdInv=nothing
    set commandUpdInv =nothing
    set rstUpdInv=nothing 

    updateFieldCollectionRecInvoice=update_Invoice_ColleRecId

end function
'Actualiza el collection recipient si viene de lleko
function updateFieldStatusCollectionRe(llekoAdmin,nfactura,statusCollection,pagada)
    update_Invoice_ColleRecId="-1"
    update_ColleRecId="-1"
    if llekoAdmin="SI" then
        'actualizamos la factura ya que el campo lo ha eliminado
        varListF="<li><a><data>"& nfactura& "</data></a></li>"                     
        'Se llamará a un servicio para obtener el campoPers            
        On Error Resume Next            
        Getfield=GetValueInvoiceCollectionRecipientIdCustomFieldId
        ' Error Handler
        If Err.Number <> 0 Then
            ' Error Occurred / Trap it
            On Error Goto 0 ' But don't let other errors hide!
            ' Code to cope with the error here
            auditar_ins_bor session("usuario"),"Error al llamar al servicio GetInvoiceCollectionRecipientIdCustomFieldId",rst("nproveedor"),"alta","","","facturas_pro"                        
            Getfield=-1
        End If
        On Error Goto 0 ' Reset error handling.  
                                                  
        if Getfield&""<>"" and Getfield&"">0 then
        else
            %><script type="text/javascript">

</script><%
        end if                                                    
        'Return json string        
        update_Invoice_ColleRecId="0"
        'response.Write"<br> pagada : "&pagada
        if Getfield&""<>"" and Getfield&"">0 then
            Resultado=0
            methodName=""
            if pagada&""<>"" and pagada=0 then            
                methodName="SelectedCheckPaymentsHaveBeenTransferredAndUpdatePaid" 
            end if
            if pagada&""<>"" and pagada=1 then   
                methodName="SelectedCheckPaymentsHaveBeenPaidAndUpdateTransferred"                         
            end if
            set rstGet = Server.CreateObject("ADODB.Recordset")
            set connGet = Server.CreateObject("ADODB.Connection")
            set commandGet =  Server.CreateObject("ADODB.Command")
            connGet.open session("dsn_cliente")
            commandGet.ActiveConnection =connGet
            commandGet.CommandTimeout = 0
            commandGet.CommandText= methodName
            commandGet.CommandType = adCmdStoredProc
            commandGet.Parameters.Append commandGet.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))
            commandGet.Parameters.Append commandGet.CreateParameter("@xml", adVarChar, adParamInput, 1000, varListF)
            commandGet.Parameters.Append commandGet.CreateParameter("@field", adInteger, adParamInput, 25, Getfield)
            commandGet.Parameters.Append commandGet.CreateParameter("@pagada", adInteger, adParamInput, 25, pagada)
            commandGet.Parameters.Append commandGet.CreateParameter("@errorCodeOut", adInteger, adParamOutput, Resultado)
    		'on error resume next

		    commandGet.Execute,,adExecuteNoRecords
		    valueError = commandGet.Parameters("@errorCodeOut").Value                            
            if rstGet.state<>0 then rstGet.close
            connGet.close
            set connGet=nothing
            set commandGet =nothing
            set rstGet=nothing 
        end if                           
    end if
    updateFieldStatusCollectionRe=valueError
end function    

function GetDomainHttpOrHttpsUrl()       
        domainUrl=GenerateURL
        if Request.ServerVariables("HTTPS")&""="off" then
            domainUrl=Replace(GenerateURL,"http://","https://")       
        end if
        GetDomainHttpOrHttpsUrl = domainUrl 
end function

function GetValueInvoiceCollectionRecipientIdCustomFieldId()       
         Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
         objHTTP.SetOption 2, objHTTP.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
         objHTTP.Open "GET",GetDomainHttpOrHttpsUrl&"/ilionServices4/lleko/Lleko.svc/getInvoiceCollectionRecipientIdCustomFieldId", false                        
         objHTTP.send ""
         jsonstring = objHTTP.responseText                        
         Set oJSON = New aspJSON
        'Load JSON string
        oJSON.loadJSON(jsonstring)                                                        
        'Get single value        
        GetValueInvoiceCollectionRecipientIdCustomFieldId =  oJSON.data("Value") 
end function
       

'****************************************************************************************************************

'********** CODIGO PRINCIPAL DE LA PÁGINA
set connRound = Server.CreateObject("ADODB.Connection")
connRound.open dsnilion

	dim EnCaja
  	n_decimales = 0%>
	<form name="facturas_pro" method="post">
	    <%' Ocultar detalle de las facturas si se da el caso

        PintarCabecera "facturas_pro.asp"

		Resultado=0
		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")
		conn.open dsnilion
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="compruebaOcultarDetalle"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@usr",adVarChar,adParamInput,50,session("usuario"))
		command.Parameters.Append command.CreateParameter("@emp",adVarChar,adParamInput,10,session("ncliente"))
		Command.Parameters.Append Command.CreateParameter("@resul", adInteger, adParamOutput, Resultado)
		'on error resume next
		command.Execute,,adExecuteNoRecords
		oculta = Command.Parameters("@resul").Value
		conn.close
		set command=nothing
		set conn=nothing

        set rstAux = Server.CreateObject("ADODB.Recordset")
	    set rstAux2 = Server.CreateObject("ADODB.Recordset")
	    set rstAux3 = Server.CreateObject("ADODB.Recordset")
	    set rst = Server.CreateObject("ADODB.Recordset")
	    set rst2 = Server.CreateObject("ADODB.Recordset")
	    set rstSelect = Server.CreateObject("ADODB.Recordset")
	    set rstDomi = Server.CreateObject("ADODB.Recordset")
	    set rstConta = Server.CreateObject("ADODB.Recordset")
        set rstMM = Server.CreateObject("ADODB.Recordset") 'ega 10/04/2008

		'Leer parámetros de la página
		mode	= Request.QueryString("mode")
		nfactura	= limpiaCadena(Request.QueryString("nfactura"))
		if nfactura="" then nfactura=limpiaCadena(Request.form("nfactura"))
		if nfactura ="" then
			nfactura = limpiaCadena(Request.QueryString("ndoc"))
			if nfactura ="" then
				nfactura = limpiaCadena(Request.form("ndoc"))
			end if
		end if
		CheckCadena nfactura

		if request.querystring("nfactura_pro")>"" then
			tnfactura_pro = limpiaCadena(request.querystring("nfactura_pro"))
		else
			tnfactura_pro = ""
		end if

		if request.querystring("cod_proyecto")>"" then
			tmp_cod_proyecto=limpiaCadena(request.querystring("cod_proyecto"))
		else
			tmp_cod_proyecto=limpiaCadena(request.form("cod_proyecto"))
		end if

		viene=limpiaCadena(request.querystring("viene"))
		if viene="" then viene=limpiaCadena(request.form("viene"))
		if viene="" then viene="facturas_pro.asp"
		'if viene="cancelar" then p_nfactura_pro=""

		if request.querystring("caju")>"" then
			caju=limpiaCadena(request.querystring("caju"))
		else
			caju=limpiaCadena(request.form("caju"))
		end if

		if request.QueryString("novei")>"" then
			novei=limpiaCadena(request.QueryString("novei"))
		else
			novei=limpiaCadena(request.form("novei"))
		end if

		'FLM : 19/01/2009: Añadir captura de request para el proveedor.
	    if request("nproveedor")>"" then
		    TraerProveedor=limpiaCadena(request("nproveedor"))
		    nproveedor=TraerProveedor
	    end if

        if request.querystring("cambiar_serie")>"" then
		    cambiar_serie=limpiaCadena(request.querystring("cambiar_serie"))
	    else
		    cambiar_serie=limpiaCadena(request.form("cambiar_serie"))
	    end if

	    if request.querystring("cambiar_cliente")>"" then
		    cambiar_cliente=limpiaCadena(request.querystring("cambiar_cliente"))
	    else
		    cambiar_cliente=limpiaCadena(request.form("cambiar_cliente"))
	    end if

        'ega 13/03/2008 asignacion de la variable pagado
        'esta asignación se hacia en varios apartados, y se ha unificado al inicio
        pagado=0
		'ega 12/03/2008 selecciona solamente un campo, ya que solo interesa saber si hay registros con esas condiciones
		selectVencimiento = "SELECT nvencimiento FROM VENCIMIENTOS_ENTRADA with(nolock) where nfactura='" & nfactura & "' and pagado='1'"
        rst2.cursorlocation=3
		rst2.open selectVencimiento ,session("dsn_cliente")
		if not rst2.eof then
		    pagado=1
		end if
		rst2.close

	    ''ricardo 3-11-2011 si no tiene acceso a la opcion de almacenes , se quitara dicho campo
	    si_tiene_acceso_almacenes=1
	    rst2.Open "exec ContractedItem '" & session("ncliente") & "','" & replace(OBJAlmacenes,"'","''") & "'", dsnilion
	    if not rst2.eof then
	        if rst2(0)=1 then
	            si_tiene_acceso_almacenes=1
	        else
	            si_tiene_acceso_almacenes=0
	        end if
	    end if
	    rst2.close
	    ''ricardo 3-11-2011 si no tiene acceso a la opcion de caja , se quitara dicho campo
	    si_tiene_acceso_caja=1
	    rst2.Open "exec ContractedItem '" & session("ncliente") & "','" & replace(OBJGestionCaja,"'","''") & "'", dsnilion
	    if not rst2.eof then
	        if rst2(0)=1 then
	            si_tiene_acceso_caja=1
	        else
	            si_tiene_acceso_caja=0
	        end if
	    end if
	    rst2.close
	    ''ricardo 3-11-2011 si no tiene acceso a la opcion de articulos o servicios , se quitaran los detalles
	    si_tiene_acceso_detalles=1
	    rst2.Open "exec ContractedItem '" & session("ncliente") & "','" & replace(OBJArticulos,"'","''") & "'", dsnilion
	    if not rst2.eof then
	        if rst2(0)=1 then
	            si_tiene_acceso_detalles=1
	        else
	            rst2.close
	            si_tiene_acceso_detalles=0
                rst2.Open "exec ContractedItem '" & session("ncliente") & "','" & replace(OBJServicios,"'","''") & "'", dsnilion
	            if not rst2.eof then
	                if rst2(0)=1 then
	                    si_tiene_acceso_detalles=1
	                else
	                    si_tiene_acceso_detalles=0
	                end if
	            end if
	        end if
	    end if
	    rst2.close

' >>> MCA 12/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compras.
'					 bloque 1/4 en facturas_pro.asp

        'FLM:200409:traslado esta líneas antes de la llamada a preparar_lista(s) y le filtro que el falor de s sea vacio
        Dim s,vlt,gl,sp,modp,modd,modi,novei,caju,cajd,oeditar,oborrar,llekoAdmin
        ObtenerParametros("facturas_pro_det")

        if s&""="" then
		    s=limpiaCadena(request.querystring("s"))
		    if s="" then s=limpiaCadena(request.form("s"))
		end if
		s=preparar_lista(s)

		dim cf 'Parámetro para mostrar el boton de cuadrar factura

		'FLM:200409:traslado esta líneas antes de la llamada a preparar_lista(s)
        if request.querystring("isCollectionRecipient")>"" then
			isCollectionRecipient=limpiaCadena(request.querystring("isCollectionRecipient"))
		else
			isCollectionRecipient=limpiaCadena(request.form("isCollectionRecipient"))
		end if                

' <<< MCA 12/04/05	 bloque 1/4

		gen_vencimientos=d_lookup("gen_vencimientos","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
		%>
		<input type="hidden" name="viene" value="<%=EncodeForHtml(viene)%>"/>
		<input type="hidden" name="mode" value="<%=EncodeForHtml(mode)%>"/>
		<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>"/>
		<input type="hidden" name="novei" value="<%=EncodeForHtml(novei)%>"/>
		<input type="hidden" name="cf" value="<%=EncodeForHtml(cf)%>"/>
		<input type="hidden" name="gen_vencimientos" value="<%=EncodeForHtml(gen_vencimientos)%>"/>
		<input type="hidden" name="s" value="<%=EncodeForHtml(s)%>"/>
        <input type="hidden" name="h_llekoAdmin" id="h_llekoAdmin" value="<%=EncodeForHtml(llekoAdmin)%>"/>
        <input type="hidden" name="isCollectionRecipient" id="isCollectionRecipient" value="<%=EncodeForHtml(isCollectionRecipient)%>"/>
		<%
' >>> MCA 12/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compras.
'					 bloque 2/4

	if p_nfactura & "">"" then
		if comprobar_LS(s,mode,p_nfactura,"FACTURAS_PRO")=0 then%>
			<script language="javascript" type="text/javascript">
                window.alert("<%=LitMsgDocNoPermAcc%>");
                document.facturas_pro.action = "facturas_pro.asp?nfactura=&mode=add&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
                document.facturas_pro.submit();
                parent.botones.document.location = "facturas_pro_bt.asp?mode=add";
			</script>
            <%'ega 13/03/2008 se crea la funcion CerrarTodo, ya que se le llamaba pero no existía
			CerrarTodo()
			response.end
		end if
	end if

' <<<

		tmp_valorado=limpiaCadena(Request.QueryString("valorado"))
		tmp_fecha=limpiaCadena(Request.QueryString("fecha"))
		tmp_serie=limpiaCadena(Request.QueryString("serie"))
		if tmp_serie="" and mode="add" then
			'Obtener la serie por defecto
			tmp_serie=ObtenerSerieTienda("FACTURA DE PROVEEDOR")
			if tmp_serie & ""="" then
				tmp_serie=d_lookup("nserie","series","tipo_documento='FACTURA DE PROVEEDOR' and pordefecto=1 and nserie like '" & session("ncliente") & "%'", session("dsn_cliente"))
			end if
		end if

		if nproveedor="" and mode="add" then
			'Obtener el proveedor de la serie por defecto.
			nproveedor=d_lookup("cliente","series","nserie='" & tmp_serie & "'",session("dsn_cliente"))
		end if

		gen_vencimiento=nz_b(d_lookup("gen_vencimientos","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))
		%><input type="hidden" name="gen_vencimiento" value="<%=EncodeForHtml(gen_vencimiento)%>"/><%

		nserieR=limpiaCadena(request.form("serie"))
		fechaR=limpiaCadena(Request.Form("fecha"))
		observacionesR=limpiaCadena(request.querystring("observaciones"))
		if request.querystring("continuar")>"" then
			continuar=limpiaCadena(request.querystring("continuar"))
		else
			continuar=1
		end if
		if request.querystring("continuarf")>"" then
			continuarf=limpiaCadena(request.querystring("continuarf"))
		else
			continuarf=1
		end if
		if request.querystring("continuari")>"" then
			continuari=limpiaCadena(request.querystring("continuari"))
		else
			continuari=1
		end if

		if request.querystring("incoterms")>"" then
			tmp_incoterms=limpiaCadena(request.querystring("incoterms"))
		else
			tmp_incoterms=limpiaCadena(request.form("incoterms"))
		end if

		if request.querystring("fob")>"" then
			tmp_fob=limpiaCadena(request.querystring("fob"))
		else
			tmp_fob=limpiaCadena(request.form("fob"))
		end if

		deuda=d_lookup("isnull(deuda,0)","facturas_pro","nfactura='" & nfactura & "'",session("dsn_cliente"))
		deudaVenc=d_lookup("isnull(sum(importe),0)","vencimientos_entrada","nfactura='" & nfactura & "' and pagado=0",session("dsn_cliente"))
		if deuda>"" then
			deudaVto=deuda-deudaVenc
		end if
		%><input type="hidden" name="deudaVto" value="<%=EncodeForHtml(deudaVto)%>"/><%

		if mode="save" or mode="first_save" then
			if request.querystring("forma_pago")>"" then
				forma_pago=limpiaCadena(request.querystring("forma_pago"))
			else
				forma_pago=limpiaCadena(request.form("forma_pago"))
			end if
			if request.querystring("pagada")>"" then
				pagada=limpiaCadena(request.querystring("pagada"))
			else
				pagada=limpiaCadena(request.form("pagada"))
			end if
			if pagada="on" then pagada="true"
		end if
		factura_cli=limpiaCadena(Request.form("factura_cli"))

		nfactura_proR=limpiaCadena(Request.Form("nfactura_pro"))
		h_vpagadaR=limpiaCadena(request.form("h_vpagada"))
		h_pagadaR=limpiaCadena(request.form("h_pagada"))

		if request.querystring("divisa")>"" then
			divisa=limpiaCadena(request.querystring("divisa"))
		else
			divisa=limpiaCadena(request.form("divisa"))
		end if

		if request.querystring("serie")>"" then
			serie=limpiaCadena(request.querystring("serie"))
		else
			serie=limpiaCadena(request.form("serie"))
		end if
		if request.querystring("fecha")>"" then
			fecha=limpiaCadena(request.querystring("fecha"))
		else
			fecha=limpiaCadena(request.form("fecha"))
		end if

		importe_antR=limpiaCadena(request.form("importe_ant"))

        campo    = limpiaCadena(request.QueryString("campo"))
	    if campo & ""="" then
	        campo = limpiaCadena(Request.Form("campo"))
        end if
	    texto    = limpiaCadena(request.QueryString("texto"))
	    if texto & ""="" then
	        texto = limpiaCadena(Request.Form("texto"))
        end if

	    lote=limpiaCadena(Request.QueryString("lote"))
	    if lote & ""="" then
	        lote=limpiaCadena(Request.form("lote"))
	    end if
	    if lote="" then lote=1
	    sentido=limpiaCadena(Request.QueryString("sentido"))
	    if sentido & ""="" then
	        sentido=limpiaCadena(Request.form("sentido"))
	    end if
	    criterio=limpiaCadena(Request.QueryString("criterio"))
	    if criterio & ""="" then
	        criterio=limpiaCadena(Request.form("criterio"))
	    end if

	    total_paginas=0
	    total_paginas=limpiaCadena(Request.QueryString("total_paginas"))
	    if total_paginas & ""="" then
	        total_paginas=limpiaCadena(Request.form("total_paginas"))
	    end if
	    if total_paginas & ""="" then total_paginas=0

	    npagina=0
	    npagina=limpiaCadena(Request.QueryString("npagina"))
	    if npagina & ""="" then
	        npagina=limpiaCadena(Request.form("npagina"))
	    end if
	    if npagina & ""="" then npagina=0

		'JMA 20/12/04. Copiar campos personalizables de los proveedores'
		redim tmp_lista_valores(10)
		for ki=1 to 10
			tmp_lista_valores(ki)=""
		next        

		'JMA 20/12/04. FIN Copiar campos personalizables de los proveedores'

		'JMA 20-12-2004 si existen campos personalizables con titulo no nulo saldrán los campos personalizables'
		si_campo_personalizables=0
        rst.cursorlocation=3
		rst.open "select ncampo from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and titulo is not null and titulo<>'' and ncampo like '" & session("ncliente") & "%'",session("dsn_cliente")
		if not rst.eof then
			si_campo_personalizables=1
		else
			si_campo_personalizables=0
		end if
		rst.close%>
		<input type="hidden" name="si_campo_personalizables" value="<%=EncodeForHtml(si_campo_personalizables)%>"/>
	    <input type="hidden" name="campo" value="<%=EncodeForHtml(campo)%>"/>
	    <input type="hidden" name="texto" value="<%=EncodeForHtml(texto)%>"/>
	    <input type="hidden" name="lote" value="<%=EncodeForHtml(lote)%>"/>
	    <input type="hidden" name="criterio" value="<%=EncodeForHtml(criterio)%>"/>        
        
		<%'JMA 20-12-2004 FIN si existen campos personalizables con titulo no nulo saldrán los campos personalizables'

        if mode="edit" and llekoAdmin="SI" then 
            oldCollectionRecipientId=""                                                   
            varListF="<li><a><data>"& nfactura& "</data></a></li>"
            'response.Write "varListF:"&varListF&"<br>"                     
            'Se llamará a un servicio para obtener el campoPers            
            On Error Resume Next            
            Getfield=GetValueInvoiceCollectionRecipientIdCustomFieldId
            ' Error Handler
            If Err.Number <> 0 Then
                ' Error Occurred / Trap it
                On Error Goto 0 ' But don't let other errors hide!
                ' Code to cope with the error here
                auditar_ins_bor session("usuario"),"Error al llamar al servicio GetInvoiceCollectionRecipientIdCustomFieldId",rst("nproveedor"),"alta","","","facturas_pro"                        
                Getfield=-1
            End If
            On Error Goto 0 ' Reset error handling.  
                                                  
            if Getfield&""<>"" and Getfield&"">0 then
            else
                %><script type="text/javascript">

</script><%
            end if                                           
            'Return json string           
            set rstGet = Server.CreateObject("ADODB.Recordset")
            set connGet = Server.CreateObject("ADODB.Connection")
            set commandGet =  Server.CreateObject("ADODB.Command")
            connGet.open session("dsn_cliente")
            commandGet.ActiveConnection =connGet
            commandGet.CommandTimeout = 0
            commandGet.CommandText= "GetCollectionRecipientInvoice"
            commandGet.CommandType = adCmdStoredProc
            commandGet.Parameters.Append commandGet.CreateParameter("@nempresa", adVarChar, adParamInput, 5, session("ncliente"))
            commandGet.Parameters.Append commandGet.CreateParameter("@xml", adVarChar, adParamInput, 1000, varListF)
            commandGet.Parameters.Append commandGet.CreateParameter("@field", adInteger, adParamInput, 25, Getfield)
    
            commandGet.Execute,,adExecuteNoRecords
            set rstGet = commandGet.Execute
            if not rstGet.eof then
                oldCollectionRecipientId= rstGet("COLLECTION_RECIPIENT_ID") 
            end if

            if rstGet.state<>0 then rstGet.close
            connGet.close
            set connGet=nothing
            set commandGet =nothing
            set rstGet=nothing 
            %>
				<script language="javascript" type="text/javascript">
                    document.getElementById("isCollectionRecipient").value = "<%=enc.EncodeForJavascript(oldCollectionRecipientId)%>";
				</script>
			<%                                           
        end if
		'JMA 20-12-2004 añadir campos personalizables a facturas_pro'
		if mode="browse" or mode="edit" or mode="add" or mode="save" or mode="first_save" then
			num_campos=0
			if mode="add" then
				redim lista_valores(10+2)
				for ki=1 to 12
					lista_valores(ki)=""
				next
				num_campos=10
			else
				rstAux2.cursorlocation=3
				rstAux2.open "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from facturas_pro as p with(nolock) where p.nfactura='" & nfactura & "'",session("dsn_cliente")
				if not rstAux2.eof then
					redim lista_valores(10+2)
					lista_valores(1)=Nulear(rstAux2("campo01"))
					lista_valores(2)=Nulear(rstAux2("campo02"))
					lista_valores(3)=Nulear(rstAux2("campo03"))
					lista_valores(4)=Nulear(rstAux2("campo04"))
					lista_valores(5)=Nulear(rstAux2("campo05"))
					lista_valores(6)=Nulear(rstAux2("campo06"))
					lista_valores(7)=Nulear(rstAux2("campo07"))
					lista_valores(8)=Nulear(rstAux2("campo08"))
					lista_valores(9)=Nulear(rstAux2("campo09"))
					lista_valores(10)=Nulear(rstAux2("campo10"))
					num_campos=10
    			else
					redim lista_valores(10+2)
					for ki=1 to 12
						lista_valores(ki)=""
					next
					num_campos=10
				end if
				rstAux2.close
			end if
		end if
		'JMA 20-12-2004 añadir campos personalizables a facturas_pro'

		'mmg >> se obtiene la serie adecuada si se ha modificado el proveedor
		p_nproveedor=limpiaCadena(Request.Form("nproveedor"))
		provR=limpiaCadena(request.querystring("prov"))
		nprov=limpiaCadena(request.querystring("nproveedor"))
	    p_serie=limpiaCadena(Request.QueryString("serie"))
	    p_serieR=limpiaCadena(Request.form("serie"))

	    if p_serieR<>"" and p_nproveedor<>"" then

            TraerProveedor= Completar(p_nproveedor,5,"0")
            nproveedor=TraerProveedor
            tmp_nproveedor=TraerProveedor
            rstAux.cursorlocation=3
            rstAux.open "select serie_fac from documentos_pro with(nolock) where nproveedor='" & session("ncliente") & nproveedor & "'", session("dsn_cliente")

		    if not rstAux.eof then
			    if rstAux("serie_fac")&"">"" then
				    p_serieR=rstAux("serie_fac")
				    tmp_serie=rstAux("serie_fac")
			    else
				    tmp_serie=p_serieR
			    end if
			else
			    tmp_serie=p_serieR
		    end if
		    rstAux.close
        else
            if provR<>"" and p_serieR="" and limpiaCadena(request.QueryString("nproveedor"))<>"" then
                TraerProveedor= Completar(provR,5,"0")
                nproveedor=TraerProveedor
                tmp_nproveedor=TraerProveedor
                if cint(null_z(cambiar_serie))=1 or cambiar_serie & ""="" then
                    rstAux.cursorlocation=3
                    rstAux.open "select serie_fac from documentos_pro with(nolock) where nproveedor='" & session("ncliente")&TraerProveedor & "'", session("dsn_cliente")
		            if not rstAux.eof then
			            if rstAux("serie_fac")&"">"" then
				            p_serieR=rstAux("serie_fac")
				            tmp_serie=rstAux("serie_fac")
			            else
				            tmp_serie=p_serie
			            end if
			        else
			            tmp_serie=p_serie
		            end if
		            rstAux.close
		        end if
            else
                if TraerProveedor="" and mode="add" then
		            'Obtener el proveedor de la serie por defecto.
    	            TraerProveedor=d_lookup("substring(cliente,6,10)","series","nserie='" & tmp_serie & "'",session("dsn_cliente"))
    	            if TraerProveedor&""="" then
		                TraerProveedor=limpiaCadena(request.querystring("prov"))
		            end if
		            nproveedor=TraerProveedor
		            tmp_nproveedor=TraerProveedor
	            end if
	        end if
        end if

	if (mode="add" or mode="edit") and nproveedor<>"" then
		if len(nproveedor)<=5 then
			nproveedor =session("ncliente") & completar(nproveedor,5,"0")
		end if
        rstAux.cursorlocation=3
		rstAux.open "select fbaja from proveedores with(nolock) where nproveedor='" & nproveedor & "'", session("dsn_cliente")
		if not rstAux.eof then
			if rstAux("fbaja")>"" then%>
				<script language="javascript" type="text/javascript">window.alert("<%=LitProvBaja%>");
				</script>
				<%nproveedor=""
				  tmp_nproveedor=""
				  mmbaja="si"
			end if
		end if
		rstAux.close
	end if

	'Captura de datos del proveedor
	if nproveedor > "" and mode<>"browse" then
		if len(nproveedor)<=5 then
			nproveedor =session("ncliente") & completar(nproveedor,5,"0")
		end if
		Error="NO"
        rstAux.cursorlocation=3
		rstAux.open "select * from proveedores with(nolock) where nproveedor='" & nproveedor & "'",session("dsn_cliente")
		if not rstAux.EOF then
		  	tmp_nproveedor=nproveedor
			tmp_descripcion=rstAux("razon_social")
			tmp_forma_pago=rstAux("forma_pago")
			tmp_tipo_pago=rstAux("tipo_pago")
			tmp_divisa=rstAux("divisa")
			tmp_descuento=rstAux("descuento")
			tmp_descuento2=rstAux("descuento2")
			tmp_recargo=null_z(rstAux("recargo"))
			tmp_irpf=null_z(rstAux("irpf"))
			tmp_IRPF_Total=rstAux("IRPF_Total")
			tmp_observaciones=observacionesR
			tmp_cod_proyecto=rstAux("proyecto")

			'JMA 20/12/04: Captura de los campos personalizables del proveedor'
			tmp_lista_valores(1)=rstAux("campo01")
			tmp_lista_valores(2)=rstAux("campo02")
			tmp_lista_valores(3)=rstAux("campo03")
			tmp_lista_valores(4)=rstAux("campo04")
			tmp_lista_valores(5)=rstAux("campo05")
			tmp_lista_valores(6)=rstAux("campo06")
			tmp_lista_valores(7)=rstAux("campo07")
			tmp_lista_valores(8)=rstAux("campo08")
			tmp_lista_valores(9)=rstAux("campo09")
			tmp_lista_valores(10)=rstAux("campo10")
			'JMA 20/12/04: FIN Captura de los campos personalizables del proveedor'
			if cint(null_z(cambiar_serie))=1 or cambiar_serie & ""="" then
				set rstObtDocCli= Server.CreateObject("ADODB.Recordset")
                dato_doc1="serie_fac"
                dato_doc2=""
                if mode="add" then
	                strselectdoc="select "
	                strselectdoc=strselectdoc & dato_doc1
	                if dato_doc2 & "">"" then
		                strselectdoc=strselectdoc & "," & dato_doc2
	                end if
	                strselectdoc=strselectdoc & ",irpf"
	                strselectdoc=strselectdoc & " from documentos_pro dc left outer join series s on dc." & dato_doc1 & "=s.nserie left outer join empresas e on s.empresa=e.cif where nproveedor='" & tmp_nproveedor & "'"
                    rstObtDocCli.cursorlocation=3
	                rstObtDocCli.open strselectdoc, session("dsn_cliente")
	                if not rstObtDocCli.eof then
		                if rstObtDocCli(dato_doc1) & "">"" then
			                tmp_serie=rstObtDocCli(dato_doc1)
		                end if
		                if dato_doc2 & "">"" then
			                p_valorado=rstObtDocCli(dato_doc2)
			                tmp_valorado=rstObtDocCli(dato_doc2)
		                end if
	                end if
	                rstObtDocCli.close
                end if
                set rstObtDocCli=nothing
			end if
            rstAux.close
		else
            rstAux.close
			Error="SI"%>
			<script language="javascript" type="text/javascript">
                window.alert("<%=LitMsgProveedorNoExiste%>");
			</script>
			<%submode2=mode
			if mode="save" then
				mode="edit"
			elseif mode="" then
				mode="add"
			elseif mode="first_save" then
				mode="add"
				nfactura=""
			end if%>
			<script language="javascript" type="text/javascript">
                document.location = "facturas_pro.asp?mode=<%=enc.EncodeForJavascript(mode)%>&nfactura=<%=enc.EncodeForJavascript(nfactura)%>&s=<%=enc.EncodeForJavascript(s)%>&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>"
                parent.botones.document.location = "facturas_pro_bt.asp?mode=<%=enc.EncodeForJavascript(mode)%>";
			</script>
			<%mode=""
            ''reo 22-5-2012 este END es necesario
            response.end
		end if

	end if

	if mode="first_save" then
		if compNumDocNuevo(nserieR,fechaR,"facturas_pro")=0 then
			%><script language="javascript" type="text/javascript">
                  window.alert("<%=LitMsgDocExistRevCont%>");
                  document.location = "facturas_pro.asp?mode=add&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>"
                  parent.botones.document.location = "facturas_pro_bt.asp?mode=add&s=<%=enc.EncodeForJavascript(s)%>"
			</script><%
			mode=""
		end if
	end if

	'Acción a realizar
	if mode="save" or mode="first_save" then
	    ''FLM:200309 Comprobamos la cuenta de abono del proveedor para esta factura.
        if mid(limpiaCadena(request.form("ncuenta_pro")), 1, 2) = "ES" then
            cuentaOk = ComprobarCuenta(Nulear(limpiaCadena(request.form("ncuenta_pro"))))
        else
            cuentaOk = true
        end if
        if cuentaOk then
		    ''ricardo 12/11/2003 comprobamos que no exista el nfactura_pro para un mismo proveedor
		    no_continuar=0
		    strselect="select count(nfactura) as contador from facturas_pro with(nolock) where nproveedor='" & Nulear(nproveedor) & "' and nfactura_pro='" & Nulear(nfactura_proR) & "' and year(fecha)= year(convert (datetime,'" & fecha & "' )) "
		    if mode="save" then
			    strselect=strselect & " and nfactura<>'" & nfactura & "'"
		    end if
            rst.cursorlocation=3
		    rst.open strselect,session("dsn_cliente")
		    if not rst.eof then
			    cuantas_facturas=null_z(rst("contador"))
			    rst.close
			    if cuantas_facturas>0 then
				    no_continuar=1
				    pagada="NO_CONTINUAR"
				    pagado="NO_CONTINUAR"
				    if mode="first_save" then
					    %><script language="javascript" type="text/javascript">
                              window.alert("<%=LitFacturaComprasYaExisteParaProveedor%>");
                              document.location = "facturas_pro.asp?mode=add&s=<%=enc.EncodeForJavascript(s)%>&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>"
                              parent.botones.document.location = "facturas_pro_bt.asp?mode=add";
					    </script><%
					    mode=""
				    else
					    submode2=mode
					    if mode="save" then
						    mode="edit"
					    elseif mode="" then
						    mode="add"
					    end if
					    %><script language="javascript" type="text/javascript">
                              window.alert("<%=LitMsgNumeroFacturaRepetido%>");
                              document.location = "facturas_pro.asp?mode=<%=enc.EncodeForJavascript(mode)%>&nfactura=<%=enc.EncodeForJavascript(nfactura)%>&s=<%=enc.EncodeForJavascript(s)%>&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>"
                              parent.botones.document.location = "facturas_pro_bt.asp?mode=<%=enc.EncodeForJavascript(mode)%>";
					    </script><%
					    ''ricardo 10-12-2003 se cambia el modo ya que si no da error
					    mode=""
				    end if
			    end if
		    else
			    rst.close
		    end if
		    if no_continuar=0 then
                if llekoAdmin="SI" then                                                
                    varListF="<li><a><data>"& nfactura& "</data></a></li>"
                    'Se llamará a un servicio para obtener el campoPers            
                    On Error Resume Next            
                    Getfield=GetValueInvoiceCollectionRecipientIdCustomFieldId
                    ' Error Handler
                    If Err.Number <> 0 Then
                        ' Error Occurred / Trap it
                        On Error Goto 0 ' But don't let other errors hide!
                        ' Code to cope with the error here
                        auditar_ins_bor session("usuario"),"Error al llamar al servicio GetInvoiceCollectionRecipientIdCustomFieldId",rst("nproveedor"),"alta","","","facturas_pro"                        
                        Getfield=-1
                    End If
                    On Error Goto 0 ' Reset error handling.  
                                                  
                    if Getfield&""<>"" and Getfield&"">0 then
                    else
                        %><script type="text/javascript">                              
</script><%
                    end if                                          
                    'Return json string           
                    set rstGet = Server.CreateObject("ADODB.Recordset")
                    set connGet = Server.CreateObject("ADODB.Connection")
                    set commandGet =  Server.CreateObject("ADODB.Command")
                    connGet.open session("dsn_cliente")
                    commandGet.ActiveConnection =connGet
                    commandGet.CommandTimeout = 0
                    commandGet.CommandText= "GetCollectionRecipientInvoice"
                    commandGet.CommandType = adCmdStoredProc
                    commandGet.Parameters.Append commandGet.CreateParameter("@nempresa", adVarChar, adParamInput, 5, session("ncliente"))
                    commandGet.Parameters.Append commandGet.CreateParameter("@xml", adVarChar, adParamInput, 1000, varListF)
                    commandGet.Parameters.Append commandGet.CreateParameter("@field", adInteger, adParamInput, 25, Getfield)
    
                    commandGet.Execute,,adExecuteNoRecords
                    set rstGet = commandGet.Execute
                    if not rstGet.eof then
                        oldCollectionRecipientId= rstGet("COLLECTION_RECIPIENT_ID") 
                    end if

                    if rstGet.state<>0 then rstGet.close
                    connGet.close
                    set connGet=nothing
                    set commandGet =nothing
                    set rstGet=nothing                            
                end if
                
			    pagada_uxa=0
			    if pagada="true" and h_vpagadaR=1 then
				    if rst.state<>0 then rst.close
                    rst.cursorlocation=2
				    rst.Open "select * from facturas_pro with(updlock) where nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic                    
				    GuardarRegistro nfactura,nserieR
				    nfactura=rst("nfactura")
				    if mode="first_save" then
					    auditar_ins_bor session("usuario"),nfactura,rst("nproveedor"),"alta","","","facturas_pro"
				    end if
				    if rst.state<>0 then rst.close
				    pagada=true
				    pagada_uxa=1 'para que no saque el mensaje de que no se puede modificar la factura
			    end if
			    'cuando no esta cobrada y se pone el cobro
			    if pagada="true" and h_pagadaR=0 then
                    
				    if rst.state<>0 then rst.close
                    rst.cursorlocation=2
				    rst.Open "select * from facturas_pro with(updlock) where nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    GuardarRegistro nfactura,nserieR
				    nfactura=rst("nfactura")
				    if mode="first_save" then
					    auditar_ins_bor session("usuario"),nfactura,rst("nproveedor"),"alta","","","facturas_pro"
				    end if
				    total_factura=reemplazar(rst("total_factura"),",",".")
				    fp=rst("forma_pago")

				    if rst.state<>0 then
					    rst.close
				    end if

				    if (cint(continuar)=0 or cint(continuarf)=0 or cint(continuari)=0) and gen_vencimiento=-1 then
					    if rst2.state<>0 then rst2.close
                        rst.cursorlocation=2
					    rst2.Open "delete from vencimientos_entrada where nfactura='" & nfactura & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					    if rst2.state<>0 then rst2.close
					    if fp>"" then
						    CrearVencimientosC nfactura
					    end if
				    end if
				    if rst2.state<>0 then rst2.close
				    pagada="true"
				    pag_com_aux="true"
				    pagada_uxa=1                              
                    statusCollection=2  
                    paidStatus=1                    
                    updateOK=updateFieldStatusCollectionRe(llekoAdmin,nfactura,statusCollection,paidStatus)
			    end if
                updateStatus="0"' para saber si hay que actualizar el collection
			    'cuando esta cobrada y se quita el cobro
			    if pagada="" and h_pagadaR=1 and (EnCaja="" or EnCaja=0 or EnCaja="NO") then
				    'comprobamos que no tengamos ningun vencimiento cobrado en caja
                    rst2.cursorlocation=3
				    rst2.Open "select v.nvencimiento,f.nfactura from vencimientos_entrada as v with(nolock),facturas_pro as f with(nolock) where f.nfactura=v.nfactura and f.nfactura='" & nfactura & "' order by nvencimiento",session("dsn_cliente")
				    'si hay algun vencimiento cobrado y en caja no se podran editar los vencimientos
				    if not rst2.eof then
					    estaencaja=""
					    while not rst2.eof and estaencaja=""
						    estaencaja=d_lookup("caja","caja","ndocumento='" & rst2("nfactura") & "-" & rst2("nvencimiento") & "'",session("dsn_cliente"))
						    rst2.movenext
					    wend
					    rst2.movefirst
					    if estaencaja="" then estadoencaja=0 else estadoencaja=1
				    end if
				    rst2.close
				    if estadoencaja=0 then
					    estaencaja2=d_sum("importe","caja","ndocumento='" & nfactura & "'",session("dsn_cliente"))
					    if estaencaja2=0 then estadoencaja=0 else estadoencaja="1"
				    end if
				    ''''''''''''''''''''
				    if estadoencaja=0 then
					    pagada=false
					    'anulamos los vencimientos
                        rst2.cursorlocation=2
					    rst2.Open "update vencimientos_entrada with(updlock) set pagado=0 where nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    else
					    EnCaja="SI"
					    if pagada_uxa=0 then
						    %><script language="javascript" type="text/javascript">
                                  window.alert("<%=LitMsgNoAnularPago%>");
						    </script><%
						    pagada=true
					    end if
						pagada_uxa=1 'para que no saque el mensaje de que no se puede modificar la factura
				    end if
			    end if
                rst.cursorlocation=2
			    rst.Open "select * from facturas_pro with(updlock) where nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    if not rst.eof then
				    pag_com=rst("pagada")
			    end if
			    pagado=0
                rst2.cursorlocation=3
			    rst2.open "SELECT * FROM VENCIMIENTOS_ENTRADA with(nolock) where nfactura='" & nfactura & "' and pagado='1'",session("dsn_cliente")
			    if not rst2.eof then
				    pagado=1
			    end if
			    rst2.close
			    'FLM:020209: SI NO ESTA EN CAJA MIRAMOS LOS EFECTOS
			    if EnCaja="NO" or EnCaja="" or EnCaja=false then
                    'Comprobamos los efectos de cliente.
                    rstAux.cursorlocation=3
                    rstAux.open "select top 1 nefecto from detalles_efpro with(nolock) where nefecto like '" & session("ncliente") & "%' and (nfacturavto='" & nfactura & "' or nfactura='" & nfactura & "') ",session("dsn_cliente")
                    if not rstAux.EOF then
                        EnEfecto="SI"
                        pagada_uxa=1 'para que no saque el mensaje de que no se puede modificar la factura
                        pagada=true
                        pagado=1
                        updateStatus="1"%>
                        <script language="javascript" type="text/javascript">
                                  alert("<%=LitMsgNoAnularCobroEfecto%>");
                        </script>
                    <%else
                        EnEfecto="No"
                    end if
                    rstAux.close
                    EnRemesa="No"                    
                end if               
                'FLM:20090429: Comprobamos si algún vencimiento está en una remesa y si se ha modificado la forma de pago.
                EstaEnRemesa=0
                if mode<>"first_save" then
                    if forma_pago<>rst("forma_pago") then
                        rst2.cursorlocation=3
		                rst2.open "select top 1 1 from remesas_pro r with(nolock) inner join detalles_rempro dr with(nolock) on dr.nremesa=r.nremesa and dr.nfacturavto='"&nfactura&"' where r.nempresa='"&session("ncliente")&"'",session("dsn_cliente")
			            if not rst2.eof then
			                EstaEnRemesa=1
			                pagada=true
			                pagada_uxa=1
                            pagado=1
			            end if
			            rst2.close
			        end if
			    end if
			    if pag_com_aux=true then pag_com=true
			        if pagado=0 and (pag_com=false or pagada=false) then
				     GuardarRegistro nfactura,serie
				        nfactura=rst("nfactura")
				        if mode="first_save" then
					        auditar_ins_bor session("usuario"),nfactura,rst("nproveedor"),"alta","","","facturas_pro"
				        end if
				        total_factura=reemplazar(rst("total_factura"),",",".")
				        fp=rst("forma_pago")
				        rst.close
				        if mode="save" then
				            if (cint(continuar)=0 or cint(continuarf)=0 or cint(continuari)=0) and gen_vencimiento=-1 then
                                rst2.cursorlocation=3
					            rst2.open "SELECT * FROM VENCIMIENTOS_ENTRADA with(nolock) where nfactura='" & nfactura & "' and pagado='1'",session("dsn_cliente")
					            if rst2.eof then
						            rst2.close

						            'ega 12/03/2008 aplicar bloqueo de fila en el borrado de vencimientos_entrada with(rowlock)
                                    rst2.cursorlocation=2
						            rst2.Open "delete from vencimientos_entrada with(rowlock) where nfactura='" & nfactura & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						            if fp>"" then
							            CrearVencimientosC nfactura
						            end if
					            else
						            rst2.close
					            end if
				             end if
				        end if                    
			        else
                        updateStatus="1"
				        if pagada_uxa=0 then
					        %><script language="javascript" type="text/javascript">window.alert("<%=LitMsgNoModifFactVencPagado%>");</script><%
				        elseif EstaEnRemesa=1 then
				            %><script language="javascript" type="text/javascript">window.alert("<%=LitMsgNoModifFactRemesa%>");</script><%
				        end if
			        end if
			        ant_mode=mode
			        mode="browse"
		        end if
                if updateStatus="0" then
                    statusCollection=1  
                    paidStatus=0
                    updateOK=updateFieldStatusCollectionRe(llekoAdmin,nfactura,statusCollection,paidStatus)
                end if
	    else 'ELSE DE COMPROBACION CUENTA BANCARIA.
        %><script language="javascript" type="text/javascript">
              alert("<%=LitCuentaAbonoError%>");
              <% if mode<>"save" then%>
                  history.back();
              <%end if %>
                  parent.botones.location="facturas_pro_bt.asp?mode=<%=iif(mode="save","edit","add")%>";
		</script>
        <%if mode="save" then
            rst.cursorlocation=2
            rst.Open "select * from facturas_pro with(updlock) where nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        end if
        mode=iif(mode="save","edit","add")
        end if 'FIN DE COMPROBACION CUENTA BANCARIA.
	end if
	if mode="delete" then
		he_borrado=1
		existe_doc=0
		'comprobamos si existe el documento a borrar
        rstAux.cursorlocation=3
		rstAux.open "select nfactura from facturas_pro with(nolock) where nfactura='" & nfactura & "'",session("dsn_cliente")
		if not rstAux.eof then
			existe_doc=1
		end if
		rstAux.close
		if existe_doc=1 then
			'Comprobar si se puede eliminar la factura.
			'comprobamos si tiene apuntes en caja
			tiene_caja=0
            rstAux.cursorlocation=3
			rstAux.open "select ndocumento from caja with(nolock) where ndocumento='" & nfactura & "'",session("dsn_cliente")
			if not rstAux.eof then
				tiene_caja=1
			end if
			rstAux.close
			if tiene_caja=0 then
				mensajeTratEquipos=TratarEquipos("","","FACTURA DE PROVEEDOR",nfactura,"","","","","",mode)
				if mid(mensajeTratEquipos,1,2)<>"OK" then
					mode="browse"%>
					<script language="javascript" type="text/javascript">
                        window.alert("<%=mensajeTratEquipos%>");
					</script>
				<%else
				    'FLM:20090424:Comprobamos que no haya ninguna remesa con el vencimiento
                    rst.cursorlocation=3
		            rst.open "select top 1 r.nremesa from remesas_pro r with(nolock) inner join detalles_rempro dr with(nolock) on dr.nremesa=r.nremesa and (dr.nfacturavto='" & nfactura & "' or dr.nfactura='" & nfactura & "') where r.nempresa='" & session("ncliente") & "' ",session("dsn_cliente")
		            if rst.EOF then
			            rst.close
				        'FLM:050409: SI está en efecto no se puede borrar.
                        rst.cursorlocation=3
			            rst.open "select nfactura from detalles_efpro with(nolock) where nefecto like '" & session("ncliente") & "%' and ( (nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "') or (nfacturavto like '" & session("ncliente") & "%' and nfacturavto='" & nfactura & "')  )",session("dsn_cliente")
			            if rst.EOF then
			                rst.close
				            'ega 12/03/2008 aplicar el no bloqueo en la seleccion de detalles_dev_pro with(nolock)
                            rstAux.cursorlocation=3
					        rstAux.open "select ndocumento from detalles_dev_pro with(nolock) where ndocumento='" & nfactura & "'",session("dsn_cliente")
					        if rstAux.EOF then
						        rstAux.close
                                rstAux.cursorlocation=3
						        rstAux.open "select nproveedor,nfactura_pro from facturas_pro with(nolock) where nfactura='" & nfactura & "'",session("dsn_cliente")
                                if not rstAux.eof then
						            npro_aux=rstAux("nproveedor")
                                else
                                    npro_aux=""
                                end if
						        rstAux.close
						        auditar_ins_bor session("usuario"),nfactura,npro_aux,"baja","","","facturas_pro"
						        InsertarHistorialNserie mensajeTratEquipos,"","","FACTURA DE PROVEEDOR",nfactura,"","","","","",mode
						         ' >>> MCA 21/04/05 Para cargar el modo add tras el borrado
						        mode="add"
						        BorrarRegistro nfactura
						        nfactura=""
                                %>
                                <script language="javascript" type="text/javascript">
                                    parent.botones.document.location = "facturas_pro_bt.asp?mode=add";
                                    SearchPage("facturas_pro_lsearch.asp?mode=init", 0);
						        </script>
                                <%
					        else
						        mode="browse"%>
						        <script language="javascript" type="text/javascript">
                                    window.alert("<%=LitMsgNoBorrarFactDev%>");
						        </script>
					        <%end if
					    else
			                rst.close
			                mode="browse"%>
			                <script language="javascript" type="text/javascript">
                                    alert("<%=LitMsgNoBorrarFactEfecto%>");
			                </script>
			            <%end if
			        else
			            rst.close
			            mode="browse"%>
			            <script language="javascript" type="text/javascript">
                                    alert("<%=LitMsgNoBorrarFactRemesa%>");
			            </script>
			         <%end if
			         end if
			    else
				    mode="browse"%>
				   <script language="javascript" type="text/javascript">
                                window.alert("<%=LitMsgBorrarCaja%>");
				    </script>
			    <%end if
		else
			mode="add"%>
			<script>
				window.alert("<%=LitMsgDocsNoExiste%>");
			</script>
		<%end if
 	end if

	total_iva_bruto	= null_z(d_sum("(importe*iva)/100","detalles_fac_pro","nfactura='" & nfactura & "'",session("dsn_cliente")))
	total_re_bruto	= null_z(d_sum("(importe*re)/100","detalles_fac_pro","nfactura='" & nfactura & "'",session("dsn_cliente")))
    total_iva_bruto	= replace(total_iva_bruto, ",", ".")
	total_re_bruto = replace(total_re_bruto, ",", ".")
	'Mostrar los datos de la página.

    ''ricardo 31/7/2003 comprobamos que existe la factura
    if mode="browse" and he_borrado<>1 then
        'ega 12/03/2008 aplicar el no bloqueo en la seleccion de facturas_pro with(nolock)
        rstAux.cursorlocation=3
        rstAux.open "select nfactura from facturas_pro with(nolock) where nfactura='" & nfactura & "'", session("dsn_cliente")
        if rstAux.eof then
	        nfactura=""%>
	        <script language="javascript" type="text/javascript">
                            window.alert("<%=LitMsgDocsNoExiste%>");
                            parent.botones.document.location = "facturas_pro_bt.asp?mode=add";
	        </script>
	        <%mode="add"
        end if
        rstAux.close
    end if

	if mode="browse" or mode="edit" then
		if nfactura="" then
			nfactura=d_lookup("nfactura","facturas_pro","nfactura like '" & session("ncliente") & "%'",session("dsn_cliente"))
		end if

        strselect="select almacen,contador,ultima_fecha from series with(updlock) where nserie='" & nserie & "'"
        rstAux.cursorlocation=3
		rstAux.Open strselect,session("dsn_cliente")

        'mmg:calculamos el almacen por defecto de la serie
	    if rstAux.eof then
		    almacenSerie= ""
	    else
		    almacenSerie= rstAux("almacen")
        end if
        rstAux.Close 'ega 09/04/2008 faltaba cerrar la conexion antes de volverla a abrir

		' JMA 20/12/04 Campos personalizables'
		rstAux.cursorlocation=3
		rstAux.open "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from facturas_pro as p with(nolock) where p.nfactura='" & nfactura & "'",session("dsn_cliente")

		if not rstAux.eof then
			redim lista_valores(10+2)
			lista_valores(1)=Nulear(rstAux("campo01"))
			lista_valores(2)=Nulear(rstAux("campo02"))
			lista_valores(3)=Nulear(rstAux("campo03"))
			lista_valores(4)=Nulear(rstAux("campo04"))
			lista_valores(5)=Nulear(rstAux("campo05"))
			lista_valores(6)=Nulear(rstAux("campo06"))
			lista_valores(7)=Nulear(rstAux("campo07"))
			lista_valores(8)=Nulear(rstAux("campo08"))
			lista_valores(9)=Nulear(rstAux("campo09"))
			lista_valores(10)=Nulear(rstAux("campo10"))
			num_campos=10
		else
			redim lista_valores(10+2)
			for ki=1 to 12
				lista_valores(ki)=""
			next
			num_campos=10
		end if
		rstAux.close
		' JMA 20/12/04 FIN Campos personalizables'
		if rst.state<>0 then rst.close

		'ega 12/03/2008 seleccionar solamente los campos necesarios de facturas_pro (antes estaba con f.*)
		' se quita la union con la tabla facturas_cli ya que solo era necesaria con el cliente 00126
		strSelDatos="select f.divisa, f.total_factura, f.pagada, f.forma_pago, f.nfactura, f.deuda, f.serie, f.nproveedor, f.fecha, f.nbalance "
		strSelDatos=strSelDatos & ", f.nfactura_pro, f.contabilizado, f.dir_envio, f.tipo_pago, f.cod_proyecto, f.tienda, f.ncuenta, f.incoterms, f.fob "
		strSelDatos=strSelDatos & ", f.observaciones, f.importe_bruto, f.descuento, f.descuento2, f.total_descuento, f.base_imponible, f.total_iva, f.total_re "
		strSelDatos=strSelDatos & ", f.recargo, f.total_recargo, f.irpf, f.total_irpf, f.irpf_total,c.razon_social,c.re as rec_equiv "
		strSelDatos=strSelDatos & " ,ser.nombre as nomserie,f.ncuenta_pro,f.banco "
		strSelDatos=strSelDatos & " from facturas_pro as f with(nolock),proveedores as c with(nolock),series as ser with(NOLOCK) "
		strSelDatos=strSelDatos & " where c.nproveedor=f.nproveedor " & "and f.nfactura='" & nfactura & "' "
		strSelDatos=strSelDatos & " and c.nproveedor like '"&session("ncliente")&"%' and ser.nserie=f.serie "
        rst.cursorlocation=3
		rst.Open strSelDatos,session("dsn_cliente")

		if not rst.eof then
            n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") &"'",session("dsn_cliente"))

			'ega 14/03/2008 el codigo de la factura de cliente, solo para el cliente 00126
			if(session("ncliente") = "00126") then
			    nfactura_cli = d_lookup("nfactura","facturas_cli","campo01='" & rst("nfactura_pro") &"' and nfactura like '"&session("ncliente")&"%'",session("dsn_cliente"))
            else
                nfactura_cli = ""
		    end if%>
			<input type="hidden" name="importe_ant" value='<%=EncodeForHtml(formatnumber(null_z(rst("total_factura")),n_decimales,-1,0,-1))%>'/>
			<input type="hidden" name="factura_cli" value='<%=EncodeForHtml(nfactura_cli)%>'/>

			<%if instr(1,rst("total_factura"),",")=0 then
				imp_aux=rst("total_factura") & ",00"
			else
				imp_aux=rst("total_factura")
			end if

			''ricardo 13/5/2008 se pone este hidden para cuando se modifique algun dato de cabecera, se sepa si calcular o no el total_re
			if nz_b(rst("rec_equiv"))=-1 then
			    valor_h_re=1
			else
			    valor_h_re=0
			end if%>
			<input type="hidden" name="h_re" value='<%=EncodeForHtml(valor_h_re)%>'/>
			<%if importe_antR<>imp_aux then
				if pagado=0 and rst("pagada")=false then
					if mode="browse" then
						gen_vencimiento=nz_b(d_lookup("gen_vencimientos","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))
						hay_ven=0
						if (cint(continuar)=0 or cint(continuarf)=0 or cint(continuari)=0) and gen_vencimiento=-1 then
						    'ega 12/03/2008 selecciona solamente un campo, ya que solo interesa saber si hay registros con esas condiciones
						    'mmg
						    RegenerarVencimientosC nfactura
						end if
					end if
				end if
			end if
		else%>
			<input type="hidden" name="importe_ant" value='0'/>
		<%end if
	elseif mode="add" then
		rst.Open "select *,'' as razon_social from facturas_pro with(nolock) where nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst.AddNew
	end if
    
    VinculosPagina(MostrarProveedores)=1
	CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
    
	'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION
	if mode="edit" then%>
		<SPAN ID="venci_paga" style="display: none">
			<table width="100%">
		  	 <tr>
				<td class="CELDAREDBOLD" align="center">
					<%=LitVPagado%>
				</td>
			</tr>
			</table>
		</SPAN>
	<%end if
    if mode<> "" then
        %><div class="headers-wrapper"><%
                    DrawDiv "header-date","",""
                    if mode="browse" then
                        DrawLabel "","",LitFecha
                        DrawSpan "","",EncodeForHtml(rst("fecha")),""
                    else
                        DrawLabel "txtMandatory","",LitFecha
                        DrawInput "width50","","fecha",EncodeForHtml(iif(mode="add",iif(tmp_fecha>"",tmp_fecha,date()),rst("fecha"))),""
                        DrawCalendar "fecha"
                    end if
                    CloseDiv
                   
                   ''ricardo 25/4/2003 para que salga el mensaje de que se ha cambiado la fecha
			        if mode="edit" then
                        %><input type="hidden" name="h_fecha" value="<%=EncodeForHtml(rst("fecha"))%>"/><%
                end if
                %><input type="hidden" name="fecha_ant" value="<%=EncodeForHtml(iif(rst("fecha")>"",rst("fecha"),date()))%>"/><%
                        DrawDiv "header-bill","",""
                        if mode="browse" then
                            DrawLabel "","",LitFactura
                        else
                            DrawLabel "txtMandatory","",LitFactura
                        end if
                        if mode="add" or mode="edit" then
                            if mode="add" then
                                DrawInput "width150px","","nfactura_pro",EncodeForHtml(iif(tnfactura_pro>"",tnfactura_pro,"")),""
                            else
                                if not rst.eof then
                                    DrawInput "width150px","","nfactura_pro",EncodeForHtml(iif(tnfactura_pro>"",tnfactura_pro,rst("nfactura_pro"))),""
                                else
                                    DrawInput "width150px","","nfactura_pro",EncodeForHtml(iif(tnfactura_pro>"",tnfactura_pro,"")),""
                                end if
                            end if
                        else

                         DrawSpan "","",EncodeForHtml(rst("nfactura_pro")),""
                %><input type="hidden" name="h_nfactura_pro" value="<%=EncodeForHtml(iif(tnfactura_pro>"",tnfactura_pro,""))%>"/><%
                end if
                    CloseDiv
                    DrawDiv  "header-nproveedor","",""
                    Formulario="facturas_pro"
				    if mode="browse" then
                        DrawLabel "","",LitProveedor
					    if rst("nproveedor")>"" then%>
							    <%=Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("nproveedor")),LitVerProveedor)%>
                        <%end if
                    else
                        DrawLabel "txtMandatory","",LitProveedor
					    if mmbaja="si" then
					        tmp_nproveedor=""
					    end if
						    %><input class="CELDA width20" type="text" maxlength="5" name="nproveedor" value="<%=EncodeForHtml(trimCodEmpresa(iif(tmp_nproveedor>"",tmp_nproveedor,rst("nproveedor"))))%>" size="8" onchange="TraerSerie('<%=enc.EncodeForJavascript(null_s(mode))%>');"/><%
						    %><a class="CELDAREFB" href="javascript:AbrirVentana('proveedores_busqueda.asp?ndoc=<%=Formulario%>&titulo=<%=LitSelProv%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerProveedor%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><%
                    end if
				    nompro   = d_lookup("razon_social","proveedores","nproveedor='" & iif(tmp_nproveedor>"",tmp_nproveedor,rst("nproveedor")) & "'",session("dsn_cliente"))
                    if nompro & ""="" then
                        nompro   = d_lookup("razon_social","proveedores","nproveedor='" & session("ncliente") & iif(tmp_nproveedor>"",tmp_nproveedor,rst("nproveedor")) & "'",session("dsn_cliente"))
                    end if
                    if mode="edit" or mode="add" then
                        %><input class="CELDA width30" type="text" disabled name="razon_social" value="<%=EncodeForHtml(nompro)%>"/><%
                    elseif mode="browse" then
                        DrawSpan "","", "&nbsp;&nbsp;" & EncodeForHtml(nompro) ,""
                    end if
                CloseDiv
            end if
	DrawDiv "header-note","",""
		if mode="browse" then
			if not rst.eof then
				if rst("pagada")=0 then
					MB=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
					n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") &"'",session("dsn_cliente"))
					EnCaja=CambioDivisa(d_sum("importe","caja","ndocumento='" & rst("nfactura") & "'",session("dsn_cliente")),rst("divisa"),rst("divisa"))
					Pendiente=miround(null_z(rst("deuda")),n_decimales)
					defecto=""
                    if si_tiene_acceso_caja=1 then
					    poner_cajasResponsive1 "input-ncaja",defecto,"ncaja","100","codigo","descripcion","","",poner_comillas(caju)
		  	  		     %><span class='header-note-inputCaja'>
				            <input class='CELDAR7' type="Text" name="impcaja" value="<%=EncodeForHtml(replace(Pendiente,",","."))%>"/>
			            </span>
			            <span class='header-note-currency'>
				            <font id="fntAbrev" class='ENCABEZADOR7'><%=d_lookup("abreviatura","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))%></font>
			            </span>
			            <span class='header-note-buttonNote '>
				            <img id="imgAnotar" src="<%=themeIlion %><%=ImgAnotar%>" <%=ParamImgAnotar%> alt="<%=LitAnotarCaja%>" title="<%=LitAnotarCaja%>"  style="cursor:pointer;" onclick="Acaja('<%=enc.EncodeForJavascript(rst("nfactura"))%>','<%=enc.EncodeForJavascript(Pendiente)%>')"/>
						    <input type="hidden" name="h_impcaja" value="<%=EncodeForHtml(Pendiente)%>"/>
	  		            </span><%

		  	  		    'ega 12/03/2008 selecciona solamente los campos necesarios para el desplegable
                        rstAux.cursorlocation=3
			  		    rstAux.Open "SELECT codigo, descripcion FROM Tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
					    DrawSelect "input-i_pago","width:150px;","i_pago",rstAux,session("ncliente") & "01","codigo","Descripcion","",""
					    rstAux.Close
					else%>
                        <span CLASS=CELDAR7 width='100'>
                            &nbsp;
                        </span>
		  	  		     <span class='CELDAL7' width="80">
				            &nbsp;
				            <input type="hidden" name="impcaja" value="<%=EncodeForHtml(Pendiente)%>"/>
			            </span>
			            <span class='CELDAL7' width="10">
				            &nbsp;
			            </span>
			            <span class='CELDAL7' width="25">
				            &nbsp;
						    <input type="hidden" name="h_impcaja" value="<%=EncodeForHtml(Pendiente)%>"/>
	  		            </span>
			            <span class='CELDAL7' width="100">
				            &nbsp;
			            </span>
	  		        <%end if
				end if
			else%>
				<!--<td align="center" width="30%"></td>-->
			<%end if
		else%>
			<!--<td align="center" width="30%"></td>-->
		<%end if
        CloseDiv%>
		<%if mode="browse" or mode="save" then

            ''ricardo 13-3-20003
            ''si la serie tiene un formato de impresion sera este el de por defecto
            ''si no sera el elegido en la tabla formatos impresion de ilion
            if not rst.eof then
	            defecto=obtener_formato_imp(rst("serie"),"FACTURA DE PROVEEDOR")
            end if
            ''''''''

			seleccion = "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='FACTURA DE PROVEEDOR' order by descripcion"
            rstSelect.cursorlocation=3
			rstSelect.Open seleccion, DsnIlion

			if not rstSelect.EOF then
				if rstSelect("personalizacion")&"">"" then
					personalizacion="../Custom/" & rstSelect("personalizacion") & "/compras/"
				end if
			end if

            if si_tiene_modulo_21 = 0 and si_tiene_modulo_22 = 0 then
                DrawDiv "header-resources alignCenter","",""
                %>
                    <a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=Servicios/recursos.asp&pag2=Servicios/recursos_bt.asp&codigo=<%=enc.EncodeForJavascript(rst("nfactura"))%>&codigo_print=<%=enc.EncodeForJavascript(rst("nfactura_pro"))%>&tipo=factura a proveedor&viene=enlaces', 'P', <%=AltoVentana%>, <%=AnchoVentana%>)">&nbsp;&nbsp;&nbsp;<%=LitEnlaces%>&nbsp;&nbsp;&nbsp;</a>
                <%
                CloseDiv
            end if
                DrawDiv "header-print","",""
                %><label><a id="idPrintFormat" class="CELDAREFB" href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(personalizacion)%>' + document.facturas_pro.formato_impresion.value+'nfactura=<%="(\'"+enc.EncodeForJavascript(null_s(nfactura))+"\')"%>&mode=browse&empresa=<%=session("ncliente")%>&novei=<%=enc.EncodeForJavascript(novei)%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitImpresionConFormato%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitImpresionConFormato%></a></label>
			<select class='CELDA' style='width:150px' name="formato_impresion">
			    <%encontrado=0
				while not rstSelect.eof
					if defecto=rstSelect("descripcion") then
						encontrado=1
						if isnull(rstSelect("parametros")) then
							prm=""
						else
							prm=rstSelect("parametros") & "&"
						end if
						%><option selected="selected" value="<%=EncodeForHtml(rstSelect("fichero")) & "?" & EncodeForHtml(prm)%>"><%=EncodeForHtml(rstSelect("descripcion"))%></option><%
					else
						if isnull(rstSelect("parametros")) then
							prm=""
						else
							prm=rstSelect("parametros") & "&"
						end if
						%><option value="<%=EncodeForHtml(rstSelect("fichero")) & "?" & EncodeForHtml(prm)%>"><%=EncodeForHtml(rstSelect("descripcion"))%></option><%
					end if
					rstSelect.movenext
				wend%>
			</select>
			<%rstSelect.close
	 	else%>
	 		<span align="right">
			</span>
	 	<%end if
   	CloseDiv%>
   	<%CloseDiv%>
        <%CloseDiv
    if mode = "browse" then
       BarraOpciones "browse", rst("nfactura")   
    end if
    ActionVersion AltoVentana, AnchoVentana%>

<% Alarma "facturas_pro.asp" %>

 <%if (mode="browse" or mode="edit" or mode="add") then
	if not rst.EOF then
		'ega 13/03/2008 modifico la select para que no sea necesario realizar un bucle con llamadas a la bd
		selectCajas = "select count(*) as contar from caja with(nolock) where ndocumento in (select convert(varchar, f.nfactura) + '-' + convert(varchar, v.nvencimiento) as documento from vencimientos_entrada as v with(nolock),facturas_pro as f with(nolock) where f.nfactura=v.nfactura and f.nfactura='" & rst("nfactura") & "')"
        rst2.cursorlocation=3
		rst2.Open  selectCajas , session("dsn_cliente")
		if not rst2.eof then
		    if rst2("contar") = 0 then estadoencaja = 0 else estadoencaja = 1
		end if
		rst2.close

		if estadoencaja=0 then
			estaencaja2=d_sum("importe","caja","ndocumento='" & rst("nfactura") & "'",session("dsn_cliente"))
			if estaencaja2=0 then estadoencaja=0 else estadoencaja="1"
		end if
		if trim(tmp_nproveedor)="" then tmp_nproveedor=rst("nproveedor")%>

		<table width="100%" bgcolor='<%=color_blau%>' border = "0" cellspacing="1" cellpadding="1">
		    <%DrawFila color_blau%>
                <td>
                    <input type="hidden" name="h_nproveedor" value="<%=EncodeForHtml(rst("nproveedor"))%>"/>
		            <input type="hidden" name="nfactura" value="<%=EncodeForHtml(rst("nfactura"))%>"/>
		            <input type="hidden" name="h_nfactura" value="<%=EncodeForHtml(rst("nfactura"))%>"/>
		            <input type="hidden" name="divisa" value="<%=EncodeForHtml(iif(tmp_divisa>"",tmp_divisa,rst("divisa")))%>"/>
		            <input type="hidden" name="olddivisa" value="<%=EncodeForHtml(rst("divisa"))%>"/>
		            <input type="hidden" name="estadoencaja" value="<%=EncodeForHtml(estadoencaja)%>"/>
		            <%if rst("pagada")=true or rst("pagada")<>0 then%>
			            <input type="hidden" name="h_pagada2" value="1"/>
		            <%else%>
			            <input type="hidden" name="h_pagada2" value="0"/>
		            <%end if

		            if pagado=1 then%>
			            <input type="hidden" name="h_vpagada" value="1"/>
		            <%else%>
			            <input type="hidden" name="h_vpagada" value="0"/>
		            <%end if

				    if pagado=1 then
					    if mode="edit" then%>
						    <script language="javascript" type="text/javascript">
                                venci_paga.style.display = "";
						    </script>
					    <%end if
				    end if%>
				    <input type="hidden" size="3" name="vpagada" value="<%=EncodeForHtml(pagado)%>" />
                </td>
            <%CloseFila %>
		</table>
	<%end if
  end if

	if (mode="browse" or mode="edit" or mode="add") then

	if not rst.EOF then
        BarraNavegacion mode,si_tiene_acceso_detalles
        %>
        <div class="Section" id="S_CABECERA">
            <a href="#" rel="toggle[CABECERA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitCabecera%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display: <%=iif(mode="add" or mode="edit","","none")%>;" id="CABECERA">
        <table width="100%" bgcolor="<%=color_blau%>" border="0"><%
		        if mode="browse" then                   
					DrawDiv 1, "", ""
                    DrawLabel "", "", LitSerie
                    DrawSpan "CELDA", "", EncodeForHtml(trimCodEmpresa(rst("serie"))) & " - " & EncodeForHtml(rst("nomserie")), ""
                    CloseDiv

                    rstMM.cursorlocation=3
					rstMM.open "select almacen from series, almacenes alm where nserie='"&rst("serie")&"' and alm.codigo=almacen and isnull(alm.fbaja,'')=''"& strwhere,session("dsn_cliente")
	                if not rstMM.EOF then
		                almacenSerie= rstMM("almacen")
	                else
		                almacenSerie= ""
	                end if
	                rstMM.close
				else
                    ' >>> MCA 12/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compras.
                    '					 bloque 4/4 en facturas_pro.asp
					strSacSerie=  "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='FACTURA DE PROVEEDOR' and nserie like '" & session("ncliente") & "%'"
					if s & "">"" then
						strSacSerie=strSacSerie & " and nserie in " & s
					end if
					strSacSerie=strSacSerie & " order by nserie"
                    rstAux.cursorlocation=3
					rstAux.open strSacSerie,session("dsn_cliente")

                    ' <<< MCA 12/04/05
                    DrawDiv 1, "", ""
                    if mode="browse" then
                        DrawLabel "", "", LitSerie
                    else
                        DrawLabel "txtMandatory", "", LitSerie
                    end if
					if mode="add" then
                        DrawSelect "", "", "serie", rstAux, iif(tmp_serie>"",tmp_serie,rst("serie")), "nserie", "descripcion", "onchange", "javascript:TraerProveedor('add');"
                    else
                        DrawSelect "", "", "serie", rstAux, iif(tmp_serie>"",tmp_serie,rst("serie")), "nserie", "descripcion", "", ""
                    end if
                    CloseDiv

			 		rstAux.close              
            
				end if

                 monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
                dato=rst("divisa")
                if dato>"" then
		                dato_divisa=dato
                else
		                dato_divisa=tmp_divisa
                end if
                if  dato_divisa<>monedaBase or mode="browse" or mode="edit" then
                    %><!--<td colspan="2" style="width:350px;"><table border='0' cellspacing="0" cellpadding="0" width="100%">--><%
                else
                    %><!--<td colspan="2" style="width:250px;"><table border='0' cellspacing="0" cellpadding="0" width="100%">--><%
                end if
                        DrawDiv "1", "", ""
                            DrawLabel "", "", LitPagada
                            if mode="add" then
                                EligeCeldaResponsive1 "check", mode, "", "", "pagada", EncodeForHtml(iif(tmp_pagada>"",nz_b(tmp_pagada),rst("pagada"))), LitCobrada
                            else
                                EligeCeldaResponsive1 "check", mode, "CELDA", "", "pagada", EncodeForHtml(iif(tmp_pagada>"",nz_b(tmp_pagada),rst("pagada"))), LitCobrada
                            end if
				        
                        %><input type="hidden" name="h_pagada" value="<%=EncodeForHtml(iif(rst("pagada")<>0,"1","0"))%>"/><%
				        %><input type="hidden" name="h_nbalance" value="<%=EncodeForHtml(rst("nbalance"))&""%>"/><%
                        CloseDiv
				%>
            </tr>
            <tr>
                <%if mode<>"edit" then
	                estilo_divisa="CELDA"
                else
	                cuantos_detalles=d_count("item","detalles_fac_pro","nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
	                cuantos_conceptos=d_count("nconcepto","conceptos_fac_pro","nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
	                if cint(cuantos_detalles) + cint(cuantos_conceptos)>0 then
		                estilo_divisa="CELDA DISABLED"
		                tipo_eligecelda="input"
	                else
		                estilo_divisa="CELDA"
		                tipo_eligecelda="select"
	                end if
                end if
            n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))
		    if n_decimales = "" then
			    n_decimales = 0
		    end if

		    dato=rst("divisa")
		    if dato>"" then
				    dato_divisa=dato
		    else
				    dato_divisa=tmp_divisa
		    end if

		    if mode<>"browse" then
			    monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
                
			    if tmp_divisa&""="" and dato&""="" then
				    dato=monedaBase
				    dato_divisa=dato
			    end if
			    if tipo_eligecelda="input" then
				    abrevEdit =  d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dato&"'",session("dsn_cliente"))
				    if mode = "edit" then
                        DrawDiv 1, "", ""
                        DrawLabel "txtMandatory", "", LitDivisa
                        DrawInput "", "", "divisabis", abrevEdit, "size='5' disabled"
                        CloseDiv
                    elseif mode = "browse" then
                        
                    end if
                    %>
                        <input type="hidden" name="h_contabilizada" value="<%=EncodeForHtml(iif(rst("contabilizado")<>0,"1","0"))%>"/>
                    <%
			    else
				    monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
                    rstSelect.cursorlocation=3
				    rstSelect.open "select codigo,abreviatura as descripcion from divisas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
                    DrawDiv 1, "", ""
                    DrawLabel "txtMandatory", "", LitDivisa
                    %>
                        <select class="CELDA" name="divisabis" onchange="javascript:cambiardivisa('<%=enc.EncodeForJavascript(null_s(monedaBase))%>')">
                    <%
                            encontrado=false
                            defecto=iif(tmp_divisa>"",tmp_divisa,dato)
	                        while not rstSelect.EOF
		                        if rstSelect("codigo")=defecto then
			                        encontrado=true
			                        response.write("<option selected='selected' value='" & EncodeForHtml(rstSelect("codigo")) & "'>" & EncodeForHtml(rstSelect("descripcion")) & "</option>")
		                        else
			                        response.write("<option value='" & EncodeForHtml(rstSelect("codigo")) & "'>" & EncodeForHtml(rstSelect("descripcion")) & "</option>")
		                        end if
		                        rstSelect.Movenext
	                        wend
                   %>
                        </select>
                        <input type="hidden" name="h_contabilizada" value="<%=iif(rst("contabilizado")<>0,"1","0")%>"/>
                    <%
                    CloseDiv

                    rstSelect.close
		 	    end if
		    else
			    DrawDiv 1, "", ""
                DrawLabel "", "", LitDivisa
                DrawSpan "CELDA", "", EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" &dato_divisa& "'",session("dsn_cliente"))), ""
                CloseDiv
                %>
                    <input type="hidden" name="h_contabilizada" value="<%=iif(rst("contabilizado")<>0,"1","0")%>"/>
                <%
		    end if

		'**RGU 25/1/2006**
		if (si_tiene_modulo_contabilidad<>0) and mode="browse" and rst("contabilizado")=true then
                if  dato_divisa<>monedaBase or mode="browse" then
                    %><!--<td colspan="2" style="width:350px;"><table border='0' cellspacing="0" cellpadding="0" width="100%">--><%
                else
                    %><!--<td colspan="2" style="width:250px;"><table border='0' cellspacing="0" cellpadding="0" width="100%">--><%
                end if
                    TmpCif=d_lookup("empresa","series","nserie like '" & session("ncliente") & "%' and nserie='"&rst("serie")&"' ",session("dsn_cliente"))
			        TmpCif=trimcodempresa(TmpCif)
			        TmpCif=LimpiarCIF(TmpCif)
                    dateInvoice=rst("fecha")
			        any=year(cstr(rst("fecha")))&""
                    ''Ricardo 15-10-2014 se cambia la manera de obtener el ejercicio activo de contabilidad
                    selectConta="SELECT c.nempresa "
                    selectConta=selectConta & " FROM CONFIGCONTAACTIVO ca with(nolock),CONFIGCONTA c with(nolock) "
                    selectConta=selectConta & " left outer join DIVISAS d with(nolock) on d.codigo=c.divisa "
                    selectConta=selectConta & " left outer join CLIENTES cli with(nolock) on cli.ncliente=c.ncliente "
                    selectConta=selectConta & " WHERE ca.ncliente='" & session("ncliente") & "' and usuario='" & session("usuario") & "' and ca.nempresa=c.nempresa and ca.ejercicio=c.ejercicio "
                    ''fin Ricardo 15-10-2014
			        rstConta.cursorlocation=3
			        rstConta.open selectConta,session("dsn_cliente")
			        if not rstConta.eof then
				        nempresaconta=rstConta("nempresa")
			        end if

			        rstConta.close
                    'dgb 01/02/2013  se compara por la fecha inicio y fin del ejercicio por si el ejercicio contable no va de enero a diciembre
			        if nempresaconta&"">"" then
                                        
                        rstConta.cursorlocation=3
				        rstConta.open "select distinct a.nasiento, a.fecha from detalles_asientos"&nempresaconta&" d with(NOLOCK), asientos"&nempresaconta&" a with(NOLOCK) where a.nasiento=d.nasiento and d.nfacturapro='"&rst("nfactura")&"' ",session("dsn_cliente")
				        if not rstConta.eof then
                        
					        rstAux.cursorlocation=3
					        rstAux.open "select c.NEMPRESA,c.EJERCICIO,p.FINICIO,p.FFIN from configcontaactivo as c with(nolock) inner join configconta as p with(nolock) on c.nempresa=p.nempresa and c.ejercicio=p.ejercicio and c.nempresa='"&nempresaconta&"' and c.ncliente = '" & session("ncliente") & "' and c.usuario='"&session("usuario")&"' ",session("dsn_cliente")
					        if not rstAux.eof then
						        TmpEjercActivo=rstAux("ejercicio")
                                TmpFinit=rstAux("finicio")
                                TmpFfin=rstAux("ffin")
						        TmpContaActivo=rstAux("nempresa")
					        end if
					        rstAux.close
                           
					        if (TmpContaActivo&""=nempresaconta and (cdate(dateInvoice)>=cdate(TmpFinit) and cdate(dateInvoice)<=cdate(TmpFfin))) then
                           
						        emp_mx="0"
                                set comm_mx =  Server.CreateObject("ADODB.Command")
                                set conn =  Server.CreateObject("ADODB.Connection")
	                            conn.open session("dsn_cliente")
	                            conn.cursorlocation=3
	                            comm_mx.activeConnection=conn
	                            comm_mx.CommandType = adCmdStoredProc
                                comm_mx.CommandText= "Emp_MX"
                                comm_mx.Parameters.Append comm_mx.CreateParameter("@p_nempresa",adVarChar,,5, session("ncliente"))
                                comm_mx.Parameters.Append comm_mx.CreateParameter("@p_dev",adVarChar,adParamOutput,4)
                                comm_mx.execute
                                if comm_mx("@p_dev")=1 then
                                    emp_mx="1"
                                end if
                                conn.close
                                set comm_mx=nothing
                                set conn=nothing
                                if emp_mx="1" then
							        VinculosPagina(MostrarAsientoMX)=1
							        CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
							        %><td class="CELDACENTER" width='0%' colspan="2"> <%=Visualizar(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),EncodeForHtml(rst("contabilizado"))))& " ("&null_s(EncodeForHtml(rstConta("fecha")))&LitNAsiento&Hiperv(OBJAsientoMX,EncodeForHtml(rstConta("nasiento"))&"&s="&nempresaconta,"browse","facturas_pro",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(rstConta("nasiento")),LitVerAsiento)&")"%></td><%
						        else
							        VinculosPagina(MostrarAsiento)=1:VinculosPagina(MostrarSubcuenta)=1
							        CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
							        %><td class="CELDACENTER" width='0%' colspan="2"> <%=Visualizar(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),EncodeForHtml(rst("contabilizado"))))& " ("&null_s(EncodeForHtml(rstConta("fecha")))&LitNAsiento&Hiperv(OBJAsiento,EncodeForHtml(rstConta("nasiento"))&"&s="&nempresaconta,"browse","facturas_pro",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(rstConta("nasiento")),LitVerAsiento)&")"%></td><%
						        end if
						        nasiento=null_s(rstConta("nasiento"))
						        rstConta.close
					        else%>
						        <td class="CELDACENTER" width='0%' colspan="2"> <%=Visualizar(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),EncodeForHtml(rst("contabilizado"))))& " ("&null_s(EncodeForHtml(rstConta("fecha")))&LitNAsiento&EncodeForHtml(rstConta("nasiento"))&")"%></td>
					        <%end if
					        else
                            if mode = "browse" then
                                DrawDiv "1", "", ""
                                    DrawLabel "", "", LitContabilizada
                                    EligeCeldaResponsive1 "check", mode, "CELDA", "", "contabilizada", EncodeForHtml(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado"))), LitContabilizada
                                CloseDiv
                            else
                                DrawCheckCelda "check", "", "", 0,LitContabilizada, 0, EncodeForHtml(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado")))
                            end if
				        end if
			        else
                                response.Write("2")
                        if mode = "browse" then
                            DrawDiv "1", "", ""
                                DrawLabel "", "", LitContabilizada
                                EligeCeldaResponsive1 "check", mode, "CELDA", "", "contabilizada", EncodeForHtml(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado"))), LitContabilizada
                            CloseDiv
                        else
                            DrawCheckCelda "check", "", "", 0,LitContabilizada, 0, EncodeForHtml(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado")))
                        end if
			        end if
                    %><input type="hidden" name="nasiento" value="<%=EncodeForHtml(nasiento)%>"/>
            <%
		else
			        '**RGU 6/4/2009
			        if mode="browse" and si_tiene_modulo_contabilidad <> 0 then
				        rstConta.CursorLocation=3
                        str="select * from modulosC_users with(NOLOCK) where usuario='"&session("usuario")&"' and ncliente='"&session("ncliente")&"' and nmodulo in('" & replace(ModContabilidad,",","','") & "') "
                        rstConta.Open str,DSNILION
                        if rstConta.EOF then
                            DrawDiv "1", "", ""
                                DrawLabel "", "", LitContabilizada
                                EligeCeldaResponsive1 "check", mode, "CELDA", "", "contabilizada", EncodeForHtml(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado"))), LitContabilizada
                            CloseDiv
			            else
                            DrawDiv "1", "", ""
				            %><label id="TDContabilizada1" class="CELDA"><a class="CELDAREFB" href="javascript:ContabilizarFra('<%=enc.EncodeForJavascript(rst("nfactura"))%>','<%=enc.EncodeForJavascript(rst("serie"))%>','<%=enc.EncodeForJavascript(rst("fecha"))%>')" OnMouseOver="self.status='<%=LitEnlaceFra%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitContabilizada%></a></label><%
				            %><iframe id='frEnlace' name='fr_Enlace' src='EnlaceContableFra.asp?' width='0' height='0' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
				            %><span class="CELDA" align="left" id="TDContabilizada2">No</span><%
                            CloseDiv
				        end if
				        rstConta.Close
			        else
                        if mode = "browse" then
                            DrawDiv "1", "", ""
                                DrawLabel "", "", LitContabilizada
                                EligeCeldaResponsive1 "check", mode, "CELDA", "", "contabilizada", EncodeForHtml(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado"))), LitContabilizada
                            CloseDiv
                        else
                            EligeCelda "check", mode, "", "", "", 0, LitContabilizada, "contabilizada", 0, EncodeForHtml(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado")))
			            end if
                    end if
			%>
            <input type="hidden" name="nasiento" value="<%=EncodeForHtml(nasiento)%>"/>
            <%
		end if%>
        </tr><%
        '*** AMP: Añadimos factor de cambio.
  	      monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
     	  abrevBase =  d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
    	  factcambio = d_lookup("factcambio","facturas_pro","nfactura='" & rst("nfactura") & "' and nfactura like '" & session("ncliente") & "%'",session("dsn_cliente"))

          HayFactorCambio=0
          
  	      if mode="browse" then
  	        if  dato_divisa<>monedaBase then
                abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dato_divisa&"'",session("dsn_cliente"))
                ''CloseFila
                DrawFila color_blau
                    EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "", LitFactCambio, "", "1" & EncodeForHtml(abrevBase) & " = " & EncodeForHtml(CStr(factcambio)) & EncodeForHtml(abreviaAtDiv)
                    HayFactorCambio=1
                CloseFila
  	        end if
  	      else
  	          ocultar=0
  	          if mode="add" or mode="edit" then
                	if mode="add" then
              	         factcambio = d_lookup("factcambio","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dato_divisa&"'",session("dsn_cliente"))
                        abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dato_divisa&"'",session("dsn_cliente"))
              	         if dato_divisa=monedaBase then ocultar=1 end if
              	     else 'modo edit
                        abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dato_divisa&"'",session("dsn_cliente"))
              	         if dato=monedaBase then ocultar=1 end if
              	     end if
                    
                     DrawDiv "1", "", "tdfactcambio"
                     DrawLabel "", "", LitFactCambio
                     DrawSpan "CELDA", "", "1" & EncodeForHtml(abrevBase) & " = ", ""
                     %>
                     <input type="text" name="nfactcambio" value="<%=EncodeForHtml(CStr(factcambio))%>" size="6" style="text-align:right" onchange="comprobarFactorCambio()"/>
                     <span class="CELDA" id="idfactcambioexpl"><%=EncodeForHtml(abreviaAtDiv)%></span>
                     <%
                     CloseDiv
  	                        
  	                if ocultar=1 then
                        %><script language="javascript" type="text/javascript">
                                    parent.pantalla.document.getElementById("tdfactcambio").style.display = "none";
                        </script><%
                    else
                        HayFactorCambio=1
                    end if
  	             end if
  	        end if
        %>

        <%DrawDiv "3-sub", "background-color: #eae7e3", ""
		    DrawLabel "", "", LITDAGRAL
        CloseDiv
			if mode<>"browse" then
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo, descripcion from formas_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
				DrawSelectCelda "width:200px;","200", "", 0, LitFormaPago, "forma_pago", rstSelect, iif(tmp_forma_pago>"",tmp_forma_pago,rst("forma_pago")), "codigo","descripcion","",""
                rstSelect.close
			else
                 DrawDiv "1", "", ""
                    DrawLabel "", "", LitFormaPago
			        DrawSpan "CELDA", "", EncodeForHtml(d_lookup("descripcion","formas_pago","codigo='" & iif(tmp_forma_pago>"",tmp_forma_pago,rst("forma_pago")) & "'",session("dsn_cliente"))), "" 		
				CloseDiv
            end if
			if mode<>"browse" then
                %><td width="10%"><%
                    if mode="add" then
			            %><input type="hidden" name="forma_pago_ant" value=""/><%
                    else
			            %><input type="hidden" name="forma_pago_ant" value="<%=EncodeForHtml(iif(tmp_forma_pago>"",tmp_forma_pago,rst("forma_pago")))%>"/><%
                    end if
                %></td><%
            else
            end if

			if mode<>"browse" then
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo, descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
				DrawSelectCelda "width:200px;","200", "", 0, LitTipoPago, "tipo_pago", rstSelect, iif(tmp_tipo_pago>"",tmp_tipo_pago,rst("tipo_pago")), "codigo","descripcion","",""
				rstSelect.close
			else
			    DrawDiv "1", "", ""
                    DrawLabel "", "", LitTipoPago
                    DrawSpan "CELDA", "", EncodeForHtml(d_lookup("descripcion","tipo_pago","codigo='" & iif(tmp_tipo_pago>"",tmp_tipo_pago,rst("tipo_pago")) & "'",session("dsn_cliente"))), ""
                CloseDiv
             end if
        
		if si_tiene_modulo_proyectos<>0 then
				if mode <> "browse" then
                    if mode="edit" then
                        hasCosts = d_count("ncode","cost_doc","ncompany='" & session("ncliente") & "' and type_doc=0 and ninvoice = '" & rst("nfactura") &"' ",session("dsn_cliente"))
                        %><input type="hidden" name="invoiceHasCosts" value="<%=EncodeForHtml(hasCosts)%>" /><%
                    end if
                    
                    %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><input class="width15" type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto"))))%>" /><input class="CELDA" type="hidden" name="cod_proyectoOLD" value="<%=EncodeForHtml(iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto"))))%>" /><label><%=LitProyecto%></label><%
                    %><iframe id='frProyecto' src='../mantenimiento/docproyectos_responsive.asp?viene=facturas_pro&mode=<%=enc.EncodeForHtmlAttribute(null_s(mode))%>&cod_proyecto=<%=EncodeForHtml(iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto"))))%>' class="width60 iframe-menu"  frameborder="no" scrolling="no" noresize="noresize"></iframe></div><%
					
				else
                    DrawDiv "1", "", ""
                        DrawLabel "", "", LitProyecto
                        DrawSpan "CELDA", "", EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & rst("cod_proyecto") & "'",session("dsn_cliente"))), ""
                    CloseDiv
                    %><input type="hidden" name="cod_proyecto" value='<%=EncodeForHtml(rst("cod_proyecto"))%>' class="CELDA" /><%
				end if
		  '**RGU 1/9/2009
		end if
		if si_tiene_modulo_ccostes<>0 then '**RGU 1/9/2009
			if si_tiene_modulo_proyectos=0 then
			else
			end if
			    '**rgu 2/9/2009
			    if tmp_serie>"" and mode="add" then
			        tmp_tienda=d_lookup("tienda","series","nserie like '"&session("ncliente")&"%' and nserie='"&tmp_serie&"' ", session("dsn_cliente")) '**rgu 2/9/2009
			    end if
			    '**rgu

				if mode <> "browse" then
                    rstSelect.cursorlocation=3
					rstSelect.open "select codigo, descripcion from tiendas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
				    DrawSelectCelda "width:200px;","200", "", 0, LitTienda, "tienda", rstSelect, iif(tmp_tienda>"",tmp_tienda,rst("tienda")),"codigo","descripcion","",""
                    rstSelect.close
			    else
				    DrawDiv "1", "", ""
                        DrawLabel "", "", LitTienda
                        DrawSpan "CELDA", "", EncodeForHtml(d_lookup("descripcion","tiendas","codigo='" & rst("tienda") & "'",session("dsn_cliente"))), ""
					CloseDiv    
				end if
		end if
        
		    if si_tiene_acceso_caja=1 then
			    if mode<>"browse" then
			        num_cuenta = ""
			        'FLM:170309: si son distintos, es que se ha cambiado de proveedor y debemos tomar la nueva cuenta. Si no la que ya existe.
			        if((tmp_nproveedor & "")<>(rst("nproveedor") & "")) then
			            num_cuenta = d_lookup("cuenta_cargo","proveedores","nproveedor='" & tmp_nproveedor & "'",session("dsn_cliente"))
			        else
			             num_cuenta=rst("ncuenta")&""
			        End if
				    rstSelect.cursorlocation=3
				    rstSelect.open "select distinct ncuenta from bancos with(nolock) where nbanco like '" & session("ncliente") & "%' order by ncuenta",session("dsn_cliente")
                    DrawSelectCelda "width:200px","200","",0,LitNCuentaCargo,"ncuentacargo",rstSelect,num_cuenta,"ncuenta","ncuenta","",""
				    rstSelect.close
			    else
                    DrawDiv "1", "", ""
                        DrawLabel "", "", LitNCuentaCargo
                        DrawSpan "CELDA", "", EncodeForHtml(rst("ncuenta")), ""
			        CloseDiv
                end if
			else
			    %><input type="hidden" name="ncuentacargo" value="" /><%
                DrawDiv "1","",""
                DrawLabel "", "", litBanco
		        if mode="browse" then
					    %><label><%=EncodeForHtml(rst("banco"))&""%></label><%
			    else
			        if(tmp_nproveedor&""="") then
			            banco=rst("banco")&""
			        else
			            banco=d_lookup("Entidad","bancos","codigo='" & left(trim(num_cuenta),4) & "'",DsnIlion)
                    end if
			        %><span class="CELDA" style="width:200px"><%
					    %><input class="CELDALEFT" disabled="disabled" maxlength="50" type="text" size="25" name="banco" value="<%=EncodeForHtml(banco)%>"/><%
				    %></span><%
			    end if
                CloseDiv
			end if
		    'FLM:120309: cuenta de abono del proveedor
			if mode="browse" then
                EligeCeldaResponsive "text",mode,"CELDALEFT","","",0,"", LitCuentaAbono,"", EncodeForHtml(rst("ncuenta_pro"))&""
			else
			    'FLM:170309: si son distintos, es que se ha cambiado de proveedor y debemos tomar la nueva cuenta. Si no la que ya existe.
			    num_cuenta = ""
                if((tmp_nproveedor & "")<>(rst("nproveedor") & "")) then
			        num_cuenta = d_lookup("ncuenta","proveedores","nproveedor='" & tmp_nproveedor & "'",session("dsn_cliente"))
			    else
			         num_cuenta=rst("ncuenta_pro")&""
			    End if
                EligeCelda "input",mode,"","","",0, LitCuentaAbono,"ncuenta_pro", 30, EncodeForHtml(num_cuenta)&""
			end if

		if si_tiene_acceso_caja=1 then
		    'FLM:130309: nombre del banco asociado a la cuenta de abono del proveedor.
		        if mode="browse" then
                    EligeCeldaResponsive "text",mode,"CELDALEFT","","",0,"", LitBanco,"", EncodeForHtml(rst("banco"))&""
			    else
			        if(tmp_nproveedor&""="") then
			            banco=rst("banco")&""
			        else
			            banco=d_lookup("Entidad","bancos","codigo='" & left(trim(num_cuenta),4) & "'",DsnIlion)
                    end if
                    EligeCelda "input",mode,"","","",0, LitBanco,"banco", 25, ""
			    end if
	    end if
		    if si_tiene_acceso_caja=1 then
			        if mode="browse" then
                        EligeCeldaResponsive "text",mode,"CELDALEFT","","",0,"", LitIncotFacPro,"", EncodeForHtml(rst("incoterms"))&""
			        else
				        defecto=iif(tmp_incoterms>"",tmp_incoterms,iif(rst("incoterms")>"",rst("incoterms"),""))
                        rstAux.cursorlocation=3
				        rstAux.open "select codigo,codigo as descripcion from incoterms with(nolock) order by descripcion",session("dsn_cliente")
                        DrawSelectCelda "width:60px","60","",0,LitIncotFacPro,"incoterms",rstAux,defecto,"codigo","descripcion","",""
				        rstAux.close
			        end if
			        if mode="browse" then
                        EligeCeldaResponsive "text",mode,"CELDALEFT","","",0,"", LitIncoPuntEntrFacPro,"", EncodeForHtml(rst("fob"))&""
			        else
                        EligeCelda "input",mode,"","","",0, LitIncoPuntEntrFacPro,"fob", 25, ""
				        %><!--<td class="CELDA" style="width:200px"><%
					        defecto=iif(tmp_fob>"",tmp_fob,iif(rst("fob")>"",rst("fob"),""))
					        %><input class="CELDALEFT" maxlength="50" type="text" size="25" name="fob" value="<%=defecto%>"/>
				        </td>--><%
			        end if
			else
			    %><input type="hidden" name="incoterms" value="" />
			    <input type="hidden" name="fob" value="" /><%
			end if

            if mode<>"browse" then
			DrawDiv "1", "", ""
            DrawLabel "", "", LitObservaciones
                if mode="add" then
                    Valor_Observaciones=iif(tmp_observaciones>"",tmp_observaciones,"")
                else
                    Valor_Observaciones=iif(tmp_observaciones>"",tmp_observaciones,rst("observaciones")&"")
                end if
                    DrawTextarea "width60", "", "observaciones", EncodeForHtml(Valor_Observaciones),"rows='2' cols='30'"
            CloseDiv
			else
                EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, LitObservaciones, LitObservaciones, 20, pintar_saltos_nuevo(EncodeForHtml(iif(tmp_observaciones>"",tmp_observaciones,rst("observaciones")&"")))
            end if
		%>
		<%'************************'
		'JMA 20/12/04 ***********'
        '************************'
		if mode="browse" and si_campo_personalizables=1 then
		      	DrawDiv "3-sub", "background-color: #eae7e3", ""
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                'ega 12/03/2008 seleccionar solamente los campos necesarios de camposperso
                rstAux2.cursorlocation=3
				rstAux2.open "select titulo,tipo,tamany,ncampocopia from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
				if not rstAux2.eof then
					'DrawFila ""
						num_campo=1
						num_campo2=1
						num_puestos=0
						num_puestos2=0
						while not rstAux2.eof
							if num_puestos2>0 and (num_puestos2 mod 2)=0 then
								num_puestos2=0
							end if
							if rstAux2("titulo") & "">"" then
								num_puestos=num_puestos+1
								num_puestos2=num_puestos2+1
								valor_campo_perso=""
								if rstAux2("tipo")=2 then
									DrawCeldaResponsive "CELDA align=left style='width:155px'","","",0, EncodeForHtml(rstAux2("titulo")) & ":",iif(lista_valores(num_campo)=1,LitSi,LitNo)
								elseif rstAux2("tipo")=3 then
									if lista_valores(num_campo) & "">"" then
										num_campo_str=cstr(num_campo)
										if len(num_campo_str)=1 then
											num_campo_str="0" & num_campo_str
										end if
										valor_ListCampPerso=d_lookup("valor","campospersolista","ncampo='" & session("ncliente") & num_campo_str & "' and tabla='DOCUMENTOS COMPRA' and ndetlista=" & lista_valores(num_campo),session("dsn_cliente"))
									else
										valor_ListCampPerso=""
									end if
									DrawCeldaResponsive "CELDA align=left style='width:200px'","","",0,EncodeForHtml(rstAux2("titulo")) & ":",EncodeForHtml(valor_ListCampPerso)
								else
									DrawCeldaResponsive "CELDA align=left style='width:200px'","","",0,EncodeForHtml(rstAux2("titulo")) & ":",EncodeForHtml(lista_valores(num_campo))
								end if
							end if
							rstAux2.movenext
							num_campo=num_campo+1
							if not rstAux2.eof then
								if rstAux2("titulo") & "">"" then
									num_campo2=num_campo2+1
								end if
							end if
						wend
					num_campos=num_puestos
				else
					num_campos=0
				end if
				rstAux2.close
		elseif mode="add" and si_campo_personalizables=1 then
		      	DrawDiv "3-sub", "background-color: #eae7e3",""
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv

                'MPC 13/06/2008 Se redimensiona el array de los campos personalizables para que cuando se elimine una factura no de error por
                ' ser de tamaño fijo en lugar de ser variable en función al número de campos.
                redim lista_valores(num_campos + 30)
                for ki=1 to num_campos + 30
			        lista_valores(ki)=""
		        next
		        'MPC
                'ega 12/03/2008 seleccionar solamente los campos necesarios de camposperso
                rstAux2.cursorlocation=3
				rstAux2.open "select titulo,tipo,tamany,ncampocopia from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
				if not rstAux2.eof then
					num_campos_existen=rstAux2.recordcount
					'DrawFila ""
						num_campo=1
						num_campo2=1
						num_puestos=0
						num_puestos2=0
						while not rstAux2.eof
							if num_puestos2>0 and (num_puestos2 mod 2)=0 then
								num_puestos2=0
							end if
							if rstAux2("titulo") & "">"" then
								num_puestos=num_puestos+1
								num_puestos2=num_puestos2+1
								valor_campo_perso=lista_valores(num_campo)

								'JMA 20/12/04. Copiar campos personalizables de los proveedores'
								'ega 12/03/2008 aplicar el no bloqueo en la seleccion de camposperso with(nolock)
                                rstSelect.cursorlocation=3
								rstSelect.open "select tipo,titulo from camposperso with(nolock) where ncampo='" & rstAux2("ncampocopia") & "' and tabla='PROVEEDORES'",session("dsn_cliente")
								if not rstSelect.eof then
									tipoPro=rstSelect("tipo")
									tituloPro=rstSelect("titulo")
								end if
								rstSelect.close
								if tipoPro=rstAux2("tipo") and tituloPro<>"" then
									if rstAux2("ncampocopia")<>"" then
										numCampoPro=cint(trimCodEmpresa(rstAux2("ncampocopia")))
										valor_campo_perso=tmp_lista_valores(numCampoPro)
									end if
								end if
								'JMA 20/12/04. FIN Copiar campos personalizables de los proveedores'

								if rstAux2("tipo")=1 then
									if isNumeric(rstAux2("tamany")) then
										tamany=rstAux2("tamany")
									else
										tamany=1
									end if
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"%><input type="text" style='width:155px' class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
									<%CloseDiv
								elseif rstAux2("tipo")=2 then
									if valor_campo_perso="" then
										valor_campo_perso=0
									end if
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"
									DrawCheckCelda "","","",0,"","campo" & num_campo,iif(valor_campo_perso=1,-1,0)
									CloseDiv
								elseif rstAux2("tipo")=3 then
									num_campo_str=cstr(num_campo)
									if len(num_campo_str)=1 then
										num_campo_str="0" & num_campo_str
									end if
									strSelListVal="select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
                                    rstAux.cursorlocation=3
									rstAux.open strSelListVal,session("dsn_cliente")
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"
                                    DrawSelect "","width:155px","campo"&num_campo,rstAux,valor_campo_perso,"ndetlista","valor","",""
				 					CloseDiv
				 					rstAux.close
								elseif rstAux2("tipo")=4 then
									if isNumeric(rstAux2("tamany")) then
										tamany=rstAux2("tamany")
									else
										tamany=1
									end if
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"%><input type="text" style='width:155px' class="CELDA" name="<%="campo" & num_campo%>" size="35" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
									<%CloseDiv
								elseif rstAux2("tipo")=5 then
									if isNumeric(rstAux2("tamany")) then
										tamany=rstAux2("tamany")
									else
										tamany=1
									end if
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"%><input type="text" style='width:155px' class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
									<%CloseDiv
								end if
							else
								%><input type="hidden" name="campo<%=num_campo%>" value=""/><%
							end if
							%><input type="hidden" name="tipo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("tipo"))%>"/><%
							%><input type="hidden" name="titulo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("titulo"))%>"/><%
							rstAux2.movenext
							num_campo=num_campo+1
							if not rstAux2.eof then
								if rstAux2("titulo") & "">"" then
									num_campo2=num_campo2+1
								end if
							end if
						wend

					num_campos=num_puestos
				else
					num_campos=0
					num_campos_existen=0
				end if
				rstAux2.close
			%><input type="hidden" name="num_campos" value="<%=num_campos_existen%>"/><%
		elseif mode="edit" and si_campo_personalizables=1 then
			    DrawDiv "3-sub", "background-color: #eae7e3", ""
			    %><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                rstAux2.cursorlocation=3
				rstAux2.open "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")

				if not rstAux2.eof then
					num_campos_existen=rstAux2.recordcount
						num_campo=1
						num_campo2=1
						num_puestos=0
						num_puestos2=0

						while not rstAux2.eof
							if num_puestos2>0 and (num_puestos2 mod 2)=0 then
								num_puestos2=0
							end if
							if rstAux2("titulo") & "">"" then
								num_puestos=num_puestos+1
								num_puestos2=num_puestos2+1
								valor_campo_perso=lista_valores(num_campo)
								'JMA 20/12/04. Copiar campos personalizables de los proveedores'
								if nproveedor > "" then
                                    rstSelect.cursorlocation=3
									rstSelect.open "select tipo,titulo from camposperso with(nolock) where ncampo='" & rstAux2("ncampocopia") & "' and tabla='PROVEEDORES'",session("dsn_cliente")
									if not rstSelect.eof then
										tipoPro=rstSelect("tipo")
										tituloPro=rstSelect("titulo")
									end if
									rstSelect.close
									if tipoPro=rstAux2("tipo") and tituloPro<>"" then
										if rstAux2("ncampocopia")<>"" then
											numCampoPro=cint(trimCodEmpresa(rstAux2("ncampocopia")))
											valor_campo_perso=tmp_lista_valores(numCampoPro)
										end if
									end if
								end if
								'JMA 20/12/04. FIN Copiar campos personalizables de los proveedores'

								if rstAux2("tipo")=1 then
									if isNumeric(rstAux2("tamany")) then
										tamany=rstAux2("tamany")
									else
										tamany=1
									end if
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"%><input type="text" style='width:155px' class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/><% 
									CloseDiv
								elseif rstAux2("tipo")=2 then
									if valor_campo_perso="" then
										valor_campo_perso=0
									end if
									DrawCheckCelda "CELDA align=left style='width:200px' align='left'","","",0,"","campo" & num_campo,iif(valor_campo_perso=1,-1,0)
								elseif rstAux2("tipo")=3 then
									num_campo_str=cstr(num_campo)
									if len(num_campo_str)=1 then
										num_campo_str="0" & num_campo_str
									end if
									strSelListVal="select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
                                    rstAux.cursorlocation=3
									rstAux.open strSelListVal,session("dsn_cliente")
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"
			 						DrawSelect "","width:155px","campo"&num_campo,rstAux,valor_campo_perso,"ndetlista","valor","",""
									CloseDiv
			 						rstAux.close
								elseif rstAux2("tipo")=4 then
									if isNumeric(rstAux2("tamany")) then
										tamany=rstAux2("tamany")
									else
										tamany=1
									end if
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"%><input type="text" style='width:155px' class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/><%
									CloseDiv
								elseif rstAux2("tipo")=5 then
									if isNumeric(rstAux2("tamany")) then
										tamany=rstAux2("tamany")
									else
										tamany=1
									end if
									DrawDiv "1","",""
									DrawLabel "","",EncodeForHtml(rstAux2("titulo")) & ":"%><input type="text" style='width:155px' class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/><%
									CloseDiv
								end if
							else
								%><input type="hidden" name="campo<%=num_campo%>" value=""/><%
							end if
							%><input type="hidden" name="tipo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("tipo"))%>"/><%
							%><input type="hidden" name="titulo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("titulo"))%>"/><%
							rstAux2.movenext
							num_campo=num_campo+1
							if not rstAux2.eof then
								if rstAux2("titulo") & "">"" then
									num_campo2=num_campo2+1
								end if
							end if
						wend
					num_campos=num_puestos
				else
					num_campos=0
					num_campos_existen=0
				end if
				rstAux2.close
			%><input type="hidden" name="num_campos" value="<%=num_campos_existen%>"/><%
		end if
		'************************'
		'FIN JMA 20/12/04 *******'
		'************************'

		%></table>
        </div>
        </div><%

		'Mostrar los Detalles del factura.
		if mode="browse" then%>
            <div class="Section" id="S_DATFINAN">
                <a href="#" rel="toggle[DATFINAN]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LITDATFINAN%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
                <div class="SectionPanel" id="DATFINAN">
                    <div id="tabs" style="display:none">
                    <ul>
                        <li><a href="#tabs1"><%=LitDetalles%></a></li>
                        <li><a href="#tabs2"><%=LitConceptos%></a></li>
                        <li><a href="#tabs3"><%=LitVencimientos%></a></li>
                        <li><a href="#tabs4"><%=LitPagosACuenta%></a></li>
                        <li><a href="#tabs-send"><%=LitDatosEnvio%></a></li>
                    </ul>
            <%if oculta=0 and si_tiene_acceso_detalles=1 then
			'Mostrar los Detalles de la factura.
			%>
			<div id="tabs1" class="overflowXauto">
			<table class="width90 md-table-responsive bCollapse">
                <tr><%
					DrawCeldaDet "'CELDAL7 underOrange width5' ","","",0,LitItem
					DrawCeldaDet "'CELDAL7 underOrange width5' ","","",0,LitCantidad
					DrawCeldaDet "'CELDAL7 underOrange width10' ","","",0,LitReferencia
					DrawCeldaDet "'CELDAL7 underOrange width15' ","","",0,LitDescripcion
					if si_tiene_acceso_almacenes=1 then
					    DrawCeldaDet "'CELDAL7 underOrange width10' ","","",0,LitAlmacen
					end if
					DrawCeldaDet "'CELDAL7 underOrange width5' ","","",0,LitPVP
					DrawCeldaDet "'CELDAL7 underOrange width5' ","","",0,LitDto
					DrawCeldaDet "'CELDAL7 underOrange width5' ","","",0,LitDto2
					DrawCeldaDet "'CELDAL7 underOrange width5' ","","",0,LitIva
					DrawCeldaDet "'CELDAL7 underOrange width5' ","","",0,LitImporte
					DrawCeldaDet "'CELDAL7 underOrange width10' ","","",0,"&nbsp;"
				%></tr>
			</table>
                <!-- Se Agrego validacion si se realizo cierre administrativo(True) o no(False) -->
			    <%if pagado=0 and rst("pagada")=false and rst("nbalance")&""="" AND NOT blnEstadoCierre then
					%><iframe id='frDetallesIns' class="width90 iframe-input md-table-responsive" name='fr_DetallesIns' src='facturas_prodetins.asp?ndoc=<%=EncodeForHtml(rst("nfactura"))%>&modP=<%=EncodeForHtml(modp)%>&nproveedor=<%=EncodeForHtml(rst("nproveedor"))%>&tienda=<%=EncodeForHtml(trimCodEmpresa(rst("tienda")))%>&almacenSerie=<%=EncodeForHtml(almacenSerie) %>&almacenTPV=<%=EncodeForHtml(almacenTPV) %>' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
				end if
				%><iframe id='frDetalles' class="width90 md-table-responsive" name="fr_Detalles" src='facturas_prodet.asp?ndoc=<%=EncodeForHtml(rst("nfactura"))%>&modP=<%=EncodeForHtml(modp)%>tienda=<%=EncodeForHtml(trimCodEmpresa(rst("tienda")))%>&almacenSerie=<%=EncodeForHtml(almacenSerie) %>&almacenTPV=<%=EncodeForHtml(almacenTPV) %>&EstadoCierre=<%=EncodeForHtml(blnEstadoCierre) %>' height="145" frameborder="yes" noresize="noresize"></iframe>
			    <br/>
                <span id="paginacion" style="display: "></span>
			</div><%
        end if 'del oculta

			''ricardo 9/8/2004 se pondra el iva que tiene establecido el cliente
			TmpIvaProveedor=d_lookup("iva","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))
			defaultIVA=d_lookup("iva","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
			if TmpIvaProveedor & "">"" then
				TmpIva=TmpIvaProveedor
			else
				TmpIva=defaultIVA
			end if%>
			<input type="hidden" name="defaultIva" value="<%=EncodeForHtml(TmpIva)%>"/>
			<div id="tabs2" class="overflowXauto" >
			<table class="width90 md-table-responsive bCollapse underOrange">
                <tr><%
				'Fila de encabezado
					DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitItem
					DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitCantidad
					DrawCeldaDet "'CELDAL7 underOrange width15'","","",0,LitDescripcion
					DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitPVP
					DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitDto
					DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitIva
					DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitImporte
					DrawCeldaDet "'CELDAL7 underOrange width10'","","",0,"&nbsp;"
				%></tr><%
				'Linea de inserción de un concepto
				if pagado=0 and rst("pagada")=false then
					%>
                    <tr>
						<td class='CELDAL7 underOrange width5'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						<td class='CELDAR7 underOrange width5'>
							<input class='CELDAR7 width100' type="text" name="cantidad"  value="1" onchange="RoundNumValue(this,<%=DEC_CANT%>);ImporteDetalle();"/>
						</td>
						<td class='CELDAL7 underOrange width15' >
							<textarea class='CELDAL7 width100' name="descripcion" onfocus="lenmensaje(this,0,255,'')" onkeydown="lenmensaje(this,0,255,'')" onkeyup="lenmensaje(this,0,255,'')" onBlur="lenmensaje(this,0,255,'')" rows="2"></textarea>
						</td>
						<td class='CELDAR7 underOrange width5' >
							<input class='CELDAR7 width100' type="text" name="pvp" value="0" onchange="RoundNumValue(this,<%=dec_prec%>);ImporteDetalle();"/>
						</td>
						<td class='CELDAR7 underOrange width5' >
							<input class='CELDAR7 width100' type="text" name="descuento" value="0" onchange="RoundNumValue(this,<%=decpor%>);ImporteDetalle();"/>
						</td>
						<%
                        rstSelect.cursorlocation=3
                        rstSelect.open "select tipo_iva, tipo_iva from tipos_iva with(nolock)",session("dsn_cliente")
						DrawSelectCeldaDet "'CELDAL7 underOrange width5'"," width100","",0,"","iva",rstSelect,TmpIva,"tipo_iva","tipo_iva","",""
						rstSelect.close%>
						<td class='CELDAR7 width5'>
							<input class='CELDAR7 underOrange width100' disabled="disabled" type="text" name="importe" value="0"/>
						</td>
						<td class=" width10 underOrange">
							<a CLASS="ic-accept noMTop" href="javascript:addConcepto('<%=enc.EncodeForJavascript(null_s(p_nfactura))%>');" onblur="javascript:document.facturas_pro.cantidad.focus();">
                                <img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/>
                  			</a>
						</td>
						<%if oculta=1 and si_tiene_acceso_detalles=0 then%>
						  <script language="javascript" type="text/javascript">
                                document.facturas_pro.cantidad.focus();
                                document.facturas_pro.cantidad.select();
						   </script>
						<%end if
					CloseFila
				end if%>
			</table>
			<iframe id="frConceptos" name="fr_Conceptos" class="width90 md-table-responsive" src='facturas_procon.asp?mode=browse&ndoc=<%=EncodeForHtml(rst("nfactura"))%>&almacenSerie=<%=EncodeForHtml(almacenSerie) %>&almacenTPV=<%=EncodeForHtml(almacenTPV) %>' height='80' frameborder="yes" noresize="noresize"></iframe>
			</div>
			<%'***** Vencimientos de la factura%>
			<div id="tabs3" class="overflowXauto" >
			<table class="width90 md-table-responsive bCollapse"><%
				'Fila de encabezado
					%>
                    <tr>
						<td class='CELDAL7 underOrange width20'></td>
						<td class='CELDAL7 underOrange width20 txtMandatory'><%=LitFecha%></td>
						<td class='CELDAL7 underOrange width20'><%=LitImporte%></td>
						<td class='CELDAL7 underOrange width20'><%=LitPagado%></td>
						<td class='CELDAL7 underOrange width20'>&nbsp</td>
                    </tr><%
					if pagado=0 and rst("pagada")=false then
							%>
                         <tr>
                            <td class='CELDAL7 underOrange width20'>
								<a href="javascript:genVencimiento('<%=enc.EncodeForJavascript(null_s(p_nfactura))%>');"><img src="<%=themeIlion %><%=ImgGenerarVencimientos%>" <%=ParamImgGenerarVencimientos%> alt="<%=LitGenerar%>" title="<%=LitGenerar%>"/></a>
							</td>
							<td class='CELDAR7 underOrange width20' >
								<input class='CELDAR7 width60' type="text" name="fechaVto" value="" onchange="cambiarfecha(document.facturas_pro.fechaVto.value,'Fecha Vencimiento')"/><%
                                 DrawCalendar "fechaVto"%>
							</td>
							<td class='CELDAR7 underOrange width20'>
								<input class='CELDAR7 width100' type="text" name="importeVto" value="0" onchange="RoundNumValue(this,<%=enc.EncodeForJavascript(null_z(NdecDiFactura))%>);"/>
							</td>
							<td class="CELDAL7 underOrange width20" >
								<input class="CELDAL7" type="checkbox" name="pagadoVto"/>
							</td>
							<td class="width20  underOrange">
								<a class="ic-accept noMTop" href="javascript:addVencimiento('<%=enc.EncodeForJavascript(null_s(p_nfactura))%>');" onblur="javascript:document.facturas_pro.fechaVto.focus();"><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>
							</td>
                        </tr><%
					else%>
						<input type="hidden" name="fechaVto"/>
						<input type="hidden" name="importeVto"/>
						<input type="checkbox" style="display:none" name="pagadoVto"/><%
					end if%>
				</table><% 'FIN DE TABLA QUE CONTIENE LA TABLAS DE ARTICULOS
				%><iframe id="frVencimientos" class="width90 md-table-responsive" name="fr_Vencimientos" src='facturas_proven.asp?mode=browse&ndoc=<%=EncodeForHtml(rst("nfactura"))%>&nproveedor=<%=EncodeForHtml(iif(tmp_nproveedor>"",tmp_nproveedor,p_nproveedor))%>&almacenSerie=<%=EncodeForHtml(almacenSerie) %>&almacenTPV=<%=EncodeForHtml(almacenTPV) %>' width='410' height='80' frameborder="yes" noresize="noresize"></iframe>
			</div><%

			'***** Pagos a cuenta de la factura
			%><div id="tabs4" class="overflowXauto" >
			<table class="width90 md-table-responsive bCollapse"><%
				'Fila de encabezado
				%><tr>
                    <td class='CELDAL7 underOrange width5'><%=LitN%></td>
					<td class='CELDAL7 underOrange width5 txtMandatory'><%=LitFecha%></td>
					<td class='CELDAL7 underOrange width15'><%=LitDescripcion%></td>
					<td class='CELDAL7 underOrange width5'><%=LitImporte%></td>
					<td class='CELDAL7 underOrange width5'><%=LitTipoPago%></td>
					<td class='CELDAL7 underOrange width5'>&nbsp</td><%
				CloseFila
				'Linea de inserción de un pago a cuenta
				if pagado=0 and rst("pagada")=0 then
						%>
                    <tr>
                        <td class='CELDAL7 underOrange width5'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						<td class='CELDAR7 underOrange width5' >
							<input class='CELDAR7 width65' type="text" name="fechaPago" value="" onchange="cambiarfecha(document.facturas_pro.fechaPago.value,'Fecha Pago')"/><%
                            DrawCalendar "fechaPago"%>
						</td>
						<td class='CELDAL7 underOrange width15'>
							<textarea class='CELDAL7 width100' name="descripcionPago" onFocus="lenmensaje(this,0,50,'')" onKeydown="lenmensaje(this,0,50,'')" onKeyup="lenmensaje(this,0,50,'')" onBlur="lenmensaje(this,0,50,'')" rows="2"></textarea>
						</td>
						<td class='CELDAR7 underOrange width5'>
							<input class='CELDAR7 width100' type="text" name="importePago" value="0" onchange="RoundNumValue(this,<%=enc.EncodeForJavascript(null_z(NdecDiFactura))%>);importepagoComp();"/>
						</td>
						<%
                        rstSelect.cursorlocation=3
                        rstSelect.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by codigo",session("dsn_cliente")
						DrawSelectCeldaDet "'CELDAR7 underOrange width5'","width100","",0,"","tipoPago",rstSelect,"","codigo","descripcion","",""
						rstSelect.close%>
						<td class="width5 underOrange">
							<a CLASS="ic-accept noMTop" href="javascript:addPago('<%=enc.EncodeForJavascript(null_s(p_nfactura))%>');" onblur="javascript:document.facturas_pro.fechaPago.focus();"><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>
						</td>
					<%CloseFila
				end if%>
				</table>
				<iframe id="frPagosCuenta" name="fr_PagosCuenta" class="width90 md-table-responsive" src='facturas_propago.asp?mode=browse&ndoc=<%=EncodeForHtml(rst("nfactura"))%>' width='650' height='80' frameborder="yes" noresize="noresize"></iframe>
                </div>
		        <div id="tabs-send" class="overflowXauto">
                    <%'ega 12/03/2008 seleccionar solamente los campos necesarios de domicilios
                    rstDomi.cursorlocation=3
			        rstDomi.Open "select domicilio,poblacion,cp,provincia from domicilios with(nolock) where codigo='" & rst("dir_envio") & "'",session("dsn_cliente")
			        if rst("dir_envio")>"" then
				        pagina="../central.asp?pag1=./compras/facturas_prodireccion_env.asp&ndoc=" &rst("nfactura") & "&mode=browse&pag2=./compras/facturas_prodireccion_env_bt.asp&titulo=" & ucase(LitDatosEnvio) & " " & rst("nfactura_pro")
			        else
				        pagina="../central.asp?pag1=./compras/facturas_prodireccion_env.asp&ndoc=" &rst("nfactura") & "&mode=edit&pag2=./compras/facturas_prodireccion_env_bt.asp&titulo=" & ucase(LitDatosEnvio) & " " & rst("nfactura_pro")
			        end if%>
                    <table width='100%' border='0' cellspacing="0" cellpadding="0">
				        <%'DrawFila color_terra%>
                        <tr>
					        <td>
	  					        <table BORDER="1" cellspacing="0" cellpadding="0">
                                      <tr>
	  				      		        <td CLASS=ENCABEZADOC height="25">	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=LitDatosEnvio%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='CELDAREFB' href="javascript:AbrirVentana('<%=pagina%>','P',<%=altoventana%>,<%=anchoventana%>)" OnMouseOver="self.status='<%=LitEditar%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitEditar%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				        		      </tr>
				      	        </table>
					        </td>
                        </tr>
				        <%'CloseFila
				        if not rstDomi.eof then
					        %>
                            <tr>
						        <td>
	    				  	        <table width='100%' border='0' cellspacing="1" cellpadding="1">
                                        <tr>
                                            <td class="ENCABEZADOL"><%=LitDomicilio%></td>
                                            <td class="ENCABEZADOL"><%=LitPoblacion%></td>
                                            <td class="ENCABEZADOL"><%=LitCP%></td>
                                            <td class="ENCABEZADOL"><%=LitProvincia%></td>
                                        </tr>
                                        <tr>
                                            <td class="CELDA"><%=EncodeForHtml(rstDomi("domicilio"))%></td>
                                            <td class="CELDA"><%=EncodeForHtml(rstDomi("poblacion"))%></td>
                                            <td class="CELDA"><%=EncodeForHtml(rstDomi("cp"))%></td>
                                            <td class="CELDA"><%=EncodeForHtml(rstDomi("provincia"))%></td>
                                        </tr>
                                    </table>
						        </td>
                            </tr>
					        <%
				        end if
				        rstDomi.close%>
                    </table>
                </div>
                </div>
                </div>
            </div>
		<%end if%>
        <div class="Section" id="S_DATTOTAL">
            <div class="SectionHeader2"><%=ucase(LitTotales)%></div>
            <div class="SectionPanel" id="DATTOTAL">
		        <input type="hidden" name="importe_bruto2" value="<%=EncodeForHtml(rst("importe_bruto"))%>"/>

		        <%n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))
		        if n_decimales = "" then
			        n_decimales = 0
		        end if

		        DIVISA=iif(tmp_divisa>"",tmp_divisa,iif(mode="add",d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")),rst("divisa")))
                
                DrawDiv "4", "", ""
                    DrawLabel "", "", LitAbrevia
                    DrawSpan "ENCABEZADOL", "", EncodeForHtml(d_lookup("abreviatura","divisas",iif(tmp_divisa>"","codigo='" & tmp_divisa & "'",iif(mode="add","moneda_base<>0 and codigo like '" & session("ncliente") & "%'","codigo='" & rst("divisa") & "'")),session("dsn_cliente"))), ""
                CloseDiv%>
                <%DrawDiv "4", "", ""
                    DrawLabel "", "", LitBruto
                    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "importe_bruto", EncodeForHtml(formatnumber(null_z(rst("importe_bruto")),n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='importe_bruto'","disabled")
                CloseDiv%>
                <%if mode<>"browse" then
                    %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total"><%
                        %><label><%=LitDto%></label><%
					    %><input class="ENCABEZADOR" type="text" name="descuento" value="<%=EncodeForHtml(iif(tmp_descuento>0,tmp_descuento,iif(rst("descuento")>"",rst("descuento"),0)))%>" OnChange="RoundNumValue(this,<%=decpor%>);Recalcula('<%=total_iva_bruto%>','<%=total_re_bruto%>')"/><%
                    %></div><%
				else
                    DrawDiv "4", "", ""
                        DrawLabel "", "", LitDto
                        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "descuento", EncodeForHtml(cstr(formatnumber(null_z(rst("descuento")),decpor,-1,0,iif(mode="browse",-1,0)))) + "%", "id='dto'"
                    CloseDiv%>
				<%end if
                if mode<>"browse" then
                    %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total"><%
                        %><label><%=LitDto2%></label><%
					%><input class="ENCABEZADOR" type="text" name="descuento2" value="<%=EncodeForHtml(iif(tmp_descuento2>0,tmp_descuento2,iif(rst("descuento2")>"",rst("descuento2"),0)))%>" OnChange="RoundNumValue(this,<%=decpor%>);Recalcula('<%=total_iva_bruto%>','<%=total_re_bruto%>')"/><%
                    %></div><%
				else
                    DrawDiv "4", "", ""
                        DrawLabel "", "", LitDto2
					    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "dto2", EncodeForHtml(formatnumber(null_z(rst("descuento2")),decpor,-1,0,iif(mode="browse",-1,0))) & "%", "id='dto2'"
                    CloseDiv%>
				<%end if
                DrawDiv "4", "", ""
                    DrawLabel "", "", LitTotalDescuento
                    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_descuento", EncodeForHtml(formatnumber(null_z(rst("total_descuento")),n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='total_descuento'","disabled")
                CloseDiv%>
                <%DrawDiv "4", "", ""
                    DrawLabel "", "", LitImponible
                    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "base_imponible", EncodeForHtml(formatnumber(null_z(rst("base_imponible")),n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='base_imponible'","disabled")
                CloseDiv%>
			    <%DrawDiv "4", "", ""
                    DrawLabel "", "", LitIva
                    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_iva", EncodeForHtml(cstr(formatnumber(Null_z(rst("total_iva")),n_decimales,-1,0,iif(mode="browse",-1,0)))), iif(mode="browse","id='total_iva'","disabled")
                CloseDiv%>
                <%DrawDiv "4", "", ""
                    DrawLabel "", "", LitRe
                    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_re", EncodeForHtml(cstr(formatnumber(Null_z(rst("total_re")),n_decimales,-1,0,iif(mode="browse",-1,0)))), iif(mode="browse","id='total_re'","disabled")
                CloseDiv%>
                <%if ((rst("recargo")<>0) or mode="edit" or mode="add") then
					if mode<>"browse" then
                    %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total"><%
                        %><label><%=LitRecargo%></label><%
						%><input class="ENCABEZADOR" type="text" name="recargo" value="<%=EncodeForHtml(iif(tmp_recargo>0,tmp_recargo,iif(rst("recargo")>"",rst("recargo"),0)))%>" OnChange="RoundNumValue(this,<%=decpor%>);Recalcula('<%=total_iva_bruto%>','<%=total_re_bruto%>')"/><%
                    %></div><%
					else
                        DrawDiv "4", "", ""
                            DrawLabel "", "", LitRecargo
                            EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "recargo", EncodeForHtml(cstr(formatnumber(null_z(rst("recargo")),decpor,-1,0,iif(mode="browse",-1,0)))) + "%", "id='recargo'"
                        CloseDiv%>
					<%end if
                    DrawDiv "4", "", ""
                        DrawLabel "", "", LitTotalRecargo
                        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_recargo", EncodeForHtml(formatnumber(null_z(rst("total_recargo")),n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='total_recargo'","disabled")
                    CloseDiv%>
				<%end if
				if ((rst("irpf")<>0) or mode="edit" or mode="add") then
                        if mode<>"browse" then
                            %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total"><%
                                %><label><%=Litirpf%></label><%
						    %><input class="ENCABEZADOR" type="text" name="irpf" value="<%=EncodeForHtml(iif(tmp_irpf>0,tmp_irpf,iif(rst("irpf")>"",rst("irpf"),0)))%>" OnChange="RoundNumValue(this,<%=decpor%>);Recalcula('<%=total_iva_bruto%>','<%=total_re_bruto%>')"/><%
                            %></div><%
					    else
                            DrawDiv "4", "", ""
                                DrawLabel "", "", Litirpf
                                EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "irpf", EncodeForHtml(cstr(formatnumber(null_z(rst("irpf")),decpor,-1,0,iif(mode="browse",-1,0)))) + "%", "id='irpf'"
                            CloseDiv%>
                        <%end if
                    	%><input class="CELDA" type="hidden" name="IRPF_Total" value="<%=EncodeForHtml(iif(tmp_IRPF_Total>"",tmp_IRPF_Total,iif(isnull(rst("IRPF_Total")),"",rst("IRPF_Total"))))%>"/><%
                    DrawDiv "4", "", ""
                        DrawLabel "", "", LitTotalirpf
                        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_irpf", EncodeForHtml(formatnumber(null_z(rst("total_irpf")),n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='total_irpf'","disabled")
                    CloseDiv%>
				<%end if

                    DrawDiv "4", "", ""
                        DrawLabel "", "", LitTotal
                        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_factura", EncodeForHtml(formatnumber(null_z(rst("total_factura")),n_decimales,-1,0,iif(mode="browse",-1,0))), "id='total_factura'"
                    CloseDiv%>
			        <%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")) then
				        n_decimales2=d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente"))
				        if n_decimales2 = "" then
					        n_decimales2 = 0
				        end if

					        DIVISA=iif(tmp_divisa>"",tmp_divisa,iif(mode="add",d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")),rst("divisa")))
					        DrawCelda "ENCABEZADOL","","",0,EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Pimporte_bruto'","ENCABEZADOR disabled")     ,"","",0,"","Pimporte_bruto",10,    EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("importe_bruto")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Pdto'","ENCABEZADOR disabled")               ,"","",0,"","Pdescuento",3,         EncodeForHtml(formatnumber(null_z(rst("descuento")),decpor,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Pdto2'","ENCABEZADOR disabled")              ,"","",0,"","Pdescuento2",3,        EncodeForHtml(formatnumber(null_z(rst("descuento2")),decpor,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Ptotal_descuento'","ENCABEZADOR disabled")   ,"","",0,"","Ptotal_descuento",10,  EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_descuento")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Pbase_imponible'","ENCABEZADOR disabled")    ,"","",0,"","Pbase_imponible",10,   EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("base_imponible")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Ptotal_iva'","ENCABEZADOR disabled")         ,"","",0,"","Ptotal_iva",10,        EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_iva")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Ptotal_re'","ENCABEZADOR disabled")          ,"","",0,"","Ptotal_re",10,         EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_re")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
					        if ((rst("recargo")<>0) or mode="edit" or mode="add") then
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Precargo'","ENCABEZADOR disabled"),"","",0,"","Precargo",3,                  EncodeForHtml(formatnumber(null_z(rst("recargo")),2,-1,0,iif(mode="browse",-1,0)))
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Ptotal_recargo'","ENCABEZADOR disabled"),"","",0,"","Ptotal_recargo",10,     EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_recargo")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
					        end if
					        if ((rst("irpf")<>0) or mode="edit" or mode="add") then
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Pirpf'","ENCABEZADOR disabled"),"","",0,"","Pirpf",3,                        EncodeForHtml(formatnumber(null_z(rst("irpf")),2,-1,0,iif(mode="browse",-1,0)))
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Ptotal_irpf'","ENCABEZADOR disabled"),"","",0,"","Ptotal_irpf",10,           EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_irpf")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
					        end if
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Ptotal_factura'","ENCABEZADOR disabled"),"","",0,"","Ptotal_factura",10,         EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_factura")),DIVISA,session("ncliente") & "01"),n_decimales2,-1,0,iif(mode="browse",-1,0)))
			        end if%>
                </div>
            </div>
		<%if mode="add" then%>
			<script language="javascript" type="text/javascript">
                              document.facturas_pro.fecha.focus();
                              document.facturas_pro.fecha.select();
			</script>
		<%elseif mode="edit" then%>
			<script language="javascript" type="text/javascript">
                              document.facturas_pro.fecha.focus();
                              document.facturas_pro.fecha.select();
			</script>
		<%elseif mode="browse" then%>
			<script type="text/javascript" language="javascript">
                jQuery(window).load(function () {
                    Redimensionar();
                    try {
                        if (document.getElementById("frDetallesIns").style.display != "none") {
                            fr_DetallesIns.document.facturas_prodetins.cantidad.focus();
                            fr_DetallesIns.document.facturas_prodetins.cantidad.select();
                        }
                    }
                    catch (e) {
                    }
                });
			</script>
		<%end if
	end if
	end if

	if mode="search" then
	else
		if mode & "">"" then
			if rst.EOF then%>
				<script language="javascript" type="text/javascript">
                    parent.botones.document.location = "facturas_pro_bt.asp?mode=search";
				</script>
			<%end if
		end if
  	end if
  	if mode="add" then rst.CancelUpdate%>
	<input type="hidden" name="ant_mode" value="<%=EncodeForHtml(ant_mode)%>"/>
	<input type="hidden" name="total_paginas" value="<%=EncodeForHtml(total_paginas)%>"/>
</form>
<%'ega llamada a la funcion de cerrar conexiones
CerrarTodo()
end if
paginaModal="blanco.asp"
AbrirModal "ProjectsCost",paginaModal,"","","no","si","noresize","S","cerrar"%>
</body>
</html>