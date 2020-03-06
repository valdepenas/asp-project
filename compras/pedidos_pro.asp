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
'mmg: variables para obtener los almacenes por defecto
dim almacenSerie
dim almacenTPV

linea1=session("f_tpv")
linea2=session("f_caja")
strconn=session("dsn_cliente")

themeIlion="/lib/estilos/" & folder & "/"


'Calculamos el almacen por defecto del TPV
'set rsTPV = Server.CreateObject("ADODB.Recordset")            
'cadena= "select c.almacen from tpv a with(NOLOCK), cajas b with(NOLOCK), tiendas c with(NOLOCK), almacenes alm with(NOLOCK) where a.caja=b.codigo and b.tienda=c.codigo and tpv='" +linea1 +"' and b.codigo='" +linea2+"' and alm.codigo=c.almacen and isnull(alm.fbaja,'')=''"
set commandTPV = nothing
set connTPV = nothing
set connTPV = Server.CreateObject("ADODB.Connection")
set commandTPV =  Server.CreateObject("ADODB.Command")
connTPV.open session("dsn_cliente")
connTPV.cursorlocation=3
commandTPV.ActiveConnection =connTPV
commandTPV.CommandTimeout = 60
commandTPV.CommandText= "select c.almacen from tpv a with(NOLOCK), cajas b with(NOLOCK), tiendas c with(NOLOCK), almacenes alm with(NOLOCK) where a.caja=b.codigo and b.tienda=c.codigo and tpv= ? and b.codigo= ? and alm.codigo=c.almacen and isnull(alm.fbaja,'')=''"
commandTPV.CommandType = adCmdText
commandTPV.Parameters.Append commandTPV.CreateParameter("@tpv",adVarChar,adParamInput, 8, linea1)
commandTPV.Parameters.Append commandTPV.CreateParameter("@codigo",adVarChar,adParamInput, 10, linea2)

set rsTPV = commandTPV.Execute

'rsTPV.Open cadena,session("dsn_cliente")
if rsTPV.eof then
	almacenTPV= ""
else
	almacenTPV= rsTPV("almacen")
end if
connTPV.close
set connTPV = nothing
set commandTPV = nothing
'rsTPV.close
set rsTPV =nothing

''ricardo 5-6-2003 se añade el parametro novei para que en los formatos de impresion no salga el item
''''ricardo 31/7/2003 comprobamos que existe el pedido que se ha pedido ver desde un listado, sino se va al modo add
'***RGU 19/12/2005 : Añadir limite de compras para el mes en que vence el pago del pedido
'                    Falta gestionar el incremento, descuento de este limite cuando se introducen borran/detalles
'					 Falta gestionar cuando se supera el limite, si seguir mediante contraseña o impedir mas detalles
' jcg 20/01/2009: Añadida la columna proyecto al proveedor y tratamiento de la misma.
%>
<!DOCTYPE html PUBLIC "-//W3C/DTD/ XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml1-transitional.dtd" />
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>

</head>

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

<!--#include file="pedidos_pro.inc" -->
<!--#include file="compras.inc" -->
<!--#include file="../ventas/documentos.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../varios2.inc" -->
<!--#include file="../perso.inc" -->
<!--#include file="../servicios/importar.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/tabs.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc"-->

<!--#include file="../styles/generalData.css.inc"-->
<!--#include file="../styles/Section.css.inc"-->
<!--#include file="../styles/ExtraLink.css.inc"-->
<!--#include file="../styles/Tabs.css.inc" -->

<!--#include file="../js/calendar.inc" -->

<!--#include file="../styles/formularios.css.inc" -->

<!--#include file="../js/dropdown.js.inc" -->

<!--#include file="../common/poner_cajaResponsive.inc" -->

<!--#include file="../styles/dropdown.css.inc" -->

<!--#include file="../common/pedidos_proActionDrop.inc" -->
<!--#include file="pedidos_pro_linkextra.inc" -->

<%si_tiene_modulo_21=ModuloContratado(session("ncliente"),"21")
si_tiene_modulo_22=ModuloContratado(session("ncliente"),"22")
''si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)



npedido	= limpiaCadena(Request.QueryString("npedido"))
if npedido ="" then
	npedido = limpiaCadena(Request.QueryString("ndoc"))
	if npedido ="" then
		npedido = limpiaCadena(Request.form("ndoc"))
	end if
end if

CheckCadena npedido 


'DivisaPedido=d_lookup("divisa","pedidos_pro","npedido like '" & session("ncliente") & "%' and npedido='" & npedido & "'",session("dsn_cliente"))	
divisaPedidoSelect = "select divisa from pedidos_pro with(nolock) where npedido like ?+'%' and npedido=?"
DivisaPedido=DlookupP2(divisaPedidoSelect, session("ncliente")&"", adVarchar, 20, npedido&"", adVarchar, 20, session("dsn_cliente"))

'NdecDiPedido=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & DivisaPedido & "'",session("dsn_cliente"))
nDecDiPedidoSelect = "select ndecimales from divisas with(nolock) where codigo like ?+'%' and codigo=?"
NdecDiPedido=DLookupP2(nDecDiPedidoSelect, session("ncliente")&"", adVarchar, 15,  DivisaPedido&"", adVarchar, 15, session("dsn_cliente"))

if NdecDiPedido & "" = "" then
    'NdecDiPedido=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente"))
    nDecDiPedidoSelect = "select ndecimales from divisas with(nolock) where codigo like ?+'%' and moneda_base<>0 "
    NdecDiPedido=DLookupP1(nDecDiPedidoSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))
end if%>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('GENERAL_DATA', 'fade=1');
    animatedcollapse.addDiv('FINANCIAL_DATA', 'fade=1');

    animatedcollapse.ontoggle = function (jQuery, divobj, state) {
        //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }
    animatedcollapse.init();
</script>
<script language="javascript" type="text/javascript">
    function addVencimiento(npedido) {
        if (document.pedidos_pro.tantoVto.value == "") document.pedidos_pro.tantoVto.value = 0;
        if (document.pedidos_pro.fechaVto.value == "" && (document.pedidos_pro.DiasFFVto.value == "" || document.pedidos_pro.DiasFFVto.value == "0")) {
            window.alert("<%=LitErrFechaPago%>");
            return;
        }
        else {
            if (document.pedidos_pro.DiasFFVto.value == "0") {
                window.alert("<%=LitErrFechaPago%>");
                return;
            }
        }


        if (isNaN(document.pedidos_pro.tantoVto.value.replace(",", "."))) {
            window.alert("<%=LitErrImportePago%>");
            document.pedidos_pro.tantoVto.value = 0;
            return;
        }
        else {
            if (parseFloat(document.pedidos_pro.tantoVto.value.replace(",", ".")) == 0) {
                window.alert("<%=LitMsgImportePositivo%>");
                document.pedidos_pro.tantoVto.value = 0;
                return;
            }
        }

        if (!cambiarfecha(document.pedidos_pro.fechaVto.value, "Fecha Vencimiento")) return;

        if (!checkdate(document.pedidos_pro.fechaVto)) {
            window.alert("<%=LitMsgFechaFecha%>");
            return;
        }

        //Asignar los valores a los campos del submarco de detalles
        fr_Vencimientos.document.vencimientos_pro_config.h_fecha.value = document.pedidos_pro.fechaVto.value;
        fr_Vencimientos.document.vencimientos_pro_config.h_tanto.value = document.pedidos_pro.tantoVto.value;
        fr_Vencimientos.document.vencimientos_pro_config.h_DiasFF.value = document.pedidos_pro.DiasFFVto.value;
        //Recargar el submarco de pagos a cuenta
        fr_Vencimientos.document.vencimientos_pro_config.action = "vencimientos_pro_config.asp?mode=first_save";
        fr_Vencimientos.document.vencimientos_pro_config.submit();
        //Limpiar los campos del formulario
        var hoy = new Date();
        document.pedidos_pro.fechaVto.value = "";//hoy.getDate() + "/" + (hoy.getMonth()+1) + "/" + hoy.getFullYear();
        document.pedidos_pro.tantoVto.value = "0";
        //Colocar el foco en el campo de cantidad.
        document.pedidos_pro.fechaVto.focus();
        document.pedidos_pro.fechaVto.select();
    }

    function ComprobarPvp(modo, npedido, nempresa, usuario, novei, pagina, nproveedor) {
        if (document.pedidos_pro.h_nalbaran.value == "NO" && document.pedidos_pro.h_nfactura.value == "NO") {
            if (document.pedidos_pro.rpc.value == nproveedor) {
                formato = "";
                if (modo == 2) {
                    formato = formato + pagina;
                    formato = formato + document.pedidos_pro.formato_impresion.value + "npedido=(\'" + npedido + "\')";
                }
                else {
                    formato = formato + document.pedidos_pro.formato_impresion.value + "npedido=(\'" + npedido + "\')&mode=browse&empresa=" + nempresa + "&novei=" + novei + "&usuario=" + usuario;
                }
                document.pedidos_pro.formato_impresionEleg.value = formato;

                document.getElementById("waitBoxOculto").style.visibility = "visible";
                document.pedidos_pro.action = "pedidos_pro.asp?mode=browse&submode=crearPedCentral&npedido=" + npedido;
                document.pedidos_pro.submit();
            }
            else {
                formato = "";
                if (modo == 2) {
                    formato = formato + pagina;
                    formato = formato + document.pedidos_pro.formato_impresion.value + "npedido=(\'" + npedido + "\')";
                }
                else {
                    formato = formato + document.pedidos_pro.formato_impresion.value + "npedido=(\'" + npedido + "\')&mode=browse&empresa=" + nempresa + "&novei=" + novei + "&usuario=" + usuario;
                }
                AbrirVentana(formato, 'I', <%=AltoVentana %>, <%=AnchoVentana %>);
            }
        }
        else {
            formato = "";
            if (modo == 2) {
                formato = formato + pagina;
                formato = formato + document.pedidos_pro.formato_impresion.value + "npedido=(\'" + npedido + "\')";
            }
            else {
                formato = formato + document.pedidos_pro.formato_impresion.value + "npedido=(\'" + npedido + "\')&mode=browse&empresa=" + nempresa + "&novei=" + novei + "&usuario=" + usuario;
            }
            AbrirVentana(formato, 'I', <%=AltoVentana %>, <%=AnchoVentana %>);
        }
    }
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
                    window.alert("<%=LitFechaMall & " " & LitFechaMalCampo%> " + modo);
                    return false;
                }
            }
        }
        return true;
    }
    function cambiarFecVto() {
        if (document.pedidos_pro.fechaVto.value == "") {
            document.pedidos_pro.DiasFFVto.disabled = false;
        }
        else {
            if (cambiarfecha(document.pedidos_pro.fechaVto.value, 'Fecha Vencimiento')) {
                document.pedidos_pro.DiasFFVto.vale = "";
                document.pedidos_pro.DiasFFVto.disabled = true;
            }
        }
    }
    function cambiarDiasFFVto() {
        while (document.pedidos_pro.DiasFFVto.value.search(" ") != -1) {
            document.pedidos_pro.DiasFFVto.value = document.pedidos_pro.DiasFFVto.value.replace(" ", "");
        }
        //if (document.pedidos_pro.DiasFFVto.value=="") document.pedidos_pro.DiasFFVto.value=0;
        if (document.pedidos_pro.DiasFFVto.value == "") {
            //window.alert("<%=LitErrFechaPago%>");
            //return;
            document.pedidos_pro.fechaVto.disabled = false;
        }
        else {
            if (isNaN(document.pedidos_pro.DiasFFVto.value.replace(",", "."))) {
                window.alert("<%=LitMsgImporteNumerico%>");
                document.pedidos_pro.DiasFFVto.value = 0;
                return;
            }
            else {
                document.pedidos_pro.fechaVto.vale = "";
                document.pedidos_pro.fechaVto.disabled = true;
            }
        }
    }

    function tier1Menu(objMenu, objImage, oculta) {
        if (objMenu.style.display == "none") {
            objMenu.style.display = "";
            objImage.src = "<%=themeIlion %><%=ImgCarpetaAbierta%>";
            switch (objMenu.id) {
                case "CABECERA":
                    if (oculta == 0) document.getElementById("DETALLES").style.display = "none";
                    document.getElementById("CONCEPTOS").style.display = "none";
                    document.getElementById("DIRENVIO").style.display = "none";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("PAGOS_CUENTA").style.display = "none";
                    document.getElementById("PEDCLI").style.display = "none";
                    document.getElementById("img6").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    if (oculta == 0) document.getElementById("img2").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("img5").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    break;

                case "DETALLES":
                    if (document.pedidos_pro.h_nalbaran.value == "NO" && document.pedidos_pro.h_nfactura.value == "NO") {
                        fr_DetallesIns.document.pedidos_prodetins.cantidad.focus();
                        fr_DetallesIns.document.pedidos_prodetins.cantidad.select();
                    }
                    document.getElementById("CABECERA").style.display = "none";
                    document.getElementById("CONCEPTOS").style.display = "none";
                    document.getElementById("DIRENVIO").style.display = "none";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("PAGOS_CUENTA").style.display = "none";
                    document.getElementById("PEDCLI").style.display = "none";
                    document.getElementById("img6").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img1").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("img5").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    break;

                case "CONCEPTOS":
                    if (document.pedidos_pro.h_nalbaran.value == "NO" && document.pedidos_pro.h_nfactura.value == "NO") {
                        document.pedidos_pro.cantidad.focus();
                        document.pedidos_pro.cantidad.select();
                    }
                    if (oculta == 0) document.getElementById("DETALLES").style.display = "none";
                    document.getElementById("CABECERA").style.display = "none";
                    document.getElementById("DIRENVIO").style.display = "none";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("PAGOS_CUENTA").style.display = "none";
                    document.getElementById("PEDCLI").style.display = "none";
                    document.getElementById("img6").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    if (oculta == 0) document.getElementById("img2").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img1").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("img5").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    break;

                case "DIRENVIO":
                    if (oculta == 0) document.getElementById("DETALLES").style.display = "none";
                    document.getElementById("CONCEPTOS").style.display = "none";
                    document.getElementById("CABECERA").style.display = "none";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("PAGOS_CUENTA").style.display = "none";
                    document.getElementById("PEDCLI").style.display = "none";
                    document.getElementById("img6").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    if (oculta == 0) document.getElementById("img2").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img1").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("img5").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    break;

                case "PAGOS_CUENTA":
                    if (document.pedidos_pro.h_nalbaran.value == "NO" && document.pedidos_pro.h_nfactura.value == "NO") {
                        var hoy = new Date();
                        document.pedidos_pro.fechaPago.value = hoy.getDate() + "/" + (hoy.getMonth() + 1) + "/" + hoy.getFullYear();
                        document.pedidos_pro.fechaPago.focus();
                        document.pedidos_pro.fechaPago.select();
                    }
                    if (oculta == 0) document.getElementById("DETALLES").style.display = "none";
                    document.getElementById("CONCEPTOS").style.display = "none";
                    document.getElementById("CABECERA").style.display = "none";
                    document.getElementById("DIRENVIO").style.display = "none";
                    document.getElementById("PEDCLI").style.display = "none";
                    document.getElementById("img6").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    if (oculta == 0) document.getElementById("img2").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img1").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    break;
                case "PEDCLI":
                    if (oculta == 0) document.getElementById("DETALLES").style.display = "none";
                    document.getElementById("CONCEPTOS").style.display = "none";
                    document.getElementById("CABECERA").style.display = "none";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("PAGOS_CUENTA").style.display = "none";
                    document.getElementById("DIRENVIO").style.display = "none";
                    if (oculta == 0) document.getElementById("img2").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img1").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    //ricardo 30-5-2007 si el parametro cpc=0 no se podran realizar pagos
                    if (document.pedidos_pro.cpc.value != "0") document.getElementById("img5").src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
                    break;
            }
        }
        else {
            objMenu.style.display = "none";
            objImage.src = "<%=themeIlion %><%=ImgCarpetaCerrada%>";
        }
        Redimensionar();
    }

    function TraerProveedor(mode, modo) {
        prov_old = document.pedidos_pro.nproveedor.value;
        cambiar_cliente = "";

        if (confirm("<%=LitCambiarSeriePuedCamPro%>")) cambiar_cliente = 1;
        else cambiar_cliente = 0;

        if (mode == "add") {
            TmpProveedor = document.pedidos_pro.nproveedor.value;
            document.pedidos_pro.tipo_pago.value = "";
            document.pedidos_pro.divisa.value = "";
            document.pedidos_pro.recargo.value = 0;
            document.pedidos_pro.ncuentacargo.value = "";

            document.pedidos_pro.nproveedor.value = "";

            document.location.href = "pedidos_pro.asp?ndoc=" + document.pedidos_pro.h_npedido.value
                + "&nproveedor=" + ""
                + "&mode=" + mode
                + "&observaciones=" + document.pedidos_pro.observaciones.value
                + "&notas=" + document.pedidos_pro.notas.value
                + "&serie=" + document.pedidos_pro.serie.value + "&fecha=" + document.pedidos_pro.fecha.value
                + "&descuento=" + document.pedidos_pro.dto.value + "&recargo=" + document.pedidos_pro.recargo.value
                + "&irpf=" + document.pedidos_pro.irpf.value
                + "&forma_pago=" + document.pedidos_pro.formas_pago.value + "&tipo_pago=" + document.pedidos_pro.tipo_pago.value
                + "&h_divisa=" + document.pedidos_pro.h_divisa.value + "&valorado=" + document.pedidos_pro.valorado.checked
                <%if si_tiene_modulo_proyectos<>0 then%>
                    + "&cod_proyecto=" + document.pedidos_pro.cod_proyecto.value
                    <%end if%>
                        + "&fecha_entrega=" + document.pedidos_pro.fecha_entrega.value
                        + "&viene=pedidos_pro.asp"
                        + "&cambiar_cliente=" + cambiar_cliente
                        + "&caju=" + document.pedidos_pro.caju.value
                        + "&novei=" + document.pedidos_pro.novei.value
                        + "&incoterms=" + document.pedidos_pro.incoterms.value
                        + "&fob=" + document.pedidos_pro.fob.value
                        + "&s=" + document.pedidos_pro.s.value
                        + "&salida=" + document.pedidos_pro.salida.value
                        + "&prov=" + prov_old
                        + "&portes=" + document.pedidos_pro.portes.value;
        }
    }

    function TraerSerie(mode) {
        prov_old = document.pedidos_pro.nproveedor.value;
        if (prov_old != "" && prov_old.length < 5) {
            for (i = prov_old.length; i < 5; i++) prov_old = "0" + prov_old;
        }
        document.pedidos_pro.nproveedor.value = prov_old;
        if (confirm("<%=LitCambiarProPuedCamSer%>")) cambiar_serie = 1;
        else cambiar_serie = 0;

        //FLM:010409: se debe traer el cliente en caso de crear y de modificar.
        TmpProveedor = document.pedidos_pro.nproveedor.value;
        document.pedidos_pro.tipo_pago.value = "";
        document.pedidos_pro.divisa.value = "";
        document.pedidos_pro.recargo.value = 0;
        document.pedidos_pro.ncuentacargo.value = "";

        document.pedidos_pro.nproveedor.value = "";

        document.location.href = "pedidos_pro.asp?ndoc=" + document.pedidos_pro.h_npedido.value
            + "&nproveedor=" + prov_old
            + "&mode=" + mode
            + "&observaciones=" + document.pedidos_pro.observaciones.value
            + "&notas=" + document.pedidos_pro.notas.value
            + "&serie=" + document.pedidos_pro.serie.value + "&fecha=" + document.pedidos_pro.fecha.value
            + "&descuento=" + document.pedidos_pro.dto.value + "&recargo=" + document.pedidos_pro.recargo.value
            + "&irpf=" + document.pedidos_pro.irpf.value
            + "&forma_pago=" + document.pedidos_pro.formas_pago.value + "&tipo_pago=" + document.pedidos_pro.tipo_pago.value
            + "&h_divisa=" + document.pedidos_pro.h_divisa.value + "&valorado=" + document.pedidos_pro.valorado.checked
            <%if si_tiene_modulo_proyectos<>0 then%>
                + "&cod_proyecto=" + document.pedidos_pro.cod_proyecto.value
                <%end if%>
                    + "&fecha_entrega=" + document.pedidos_pro.fecha_entrega.value
                    + "&viene=pedidos_pro.asp"
                    + "&cambiar_serie=" + cambiar_serie
                    + "&caju=" + document.pedidos_pro.caju.value
                    + "&novei=" + document.pedidos_pro.novei.value
                    + "&incoterms=" + document.pedidos_pro.incoterms.value
                    + "&fob=" + document.pedidos_pro.fob.value
                    + "&s=" + document.pedidos_pro.s.value
                    + "&salida=" + document.pedidos_pro.salida.value;
        + "&prov=" + prov_old
            + "&portes=" + document.pedidos_pro.portes.value;
    }

    function Recalcula(total_iva_bruto, total_re_bruto) {
        document.pedidos_pro.importe_bruto.value = document.pedidos_pro.importe_bruto.value.replace(",", ".");
        document.pedidos_pro.dto.value = document.pedidos_pro.dto.value.replace(",", ".");
        document.pedidos_pro.dto2.value = document.pedidos_pro.dto2.value.replace(",", ".");
        document.pedidos_pro.recargo.value = document.pedidos_pro.recargo.value.replace(",", ".");
        document.pedidos_pro.irpf.value = document.pedidos_pro.irpf.value.replace(",", ".");

        document.pedidos_pro.base_imponible.value = (parseFloat(document.pedidos_pro.importe_bruto.value) * (100 - parseFloat(document.pedidos_pro.dto.value))) / 100;
        document.pedidos_pro.base_imponible.value = (parseFloat(document.pedidos_pro.base_imponible.value) * (100 - parseFloat(document.pedidos_pro.dto2.value))) / 100;
        document.pedidos_pro.base_imponible.value = parseFloat(document.pedidos_pro.base_imponible.value).toFixed(<%=enc.EncodeForJavascript(NdecDiPedido)%>);
        document.pedidos_pro.total_descuento.value = parseFloat(document.pedidos_pro.importe_bruto.value) - parseFloat(document.pedidos_pro.base_imponible.value);
        document.pedidos_pro.total_descuento.value = parseFloat(document.pedidos_pro.total_descuento.value).toFixed(<%=enc.EncodeForJavascript(NdecDiPedido)%>);

        document.pedidos_pro.total_iva.value = (parseFloat(total_iva_bruto) * (100 - parseFloat(document.pedidos_pro.dto.value))) / 100;
        document.pedidos_pro.total_iva.value = (parseFloat(document.pedidos_pro.total_iva.value) * (100 - parseFloat(document.pedidos_pro.dto2.value))) / 100;
        document.pedidos_pro.total_iva.value = parseFloat(document.pedidos_pro.total_iva.value).toFixed(<%=enc.EncodeForJavascript(NdecDiPedido)%>);

        document.pedidos_pro.total_re.value = (parseFloat(total_re_bruto) * (100 - parseFloat(document.pedidos_pro.dto.value))) / 100;
        document.pedidos_pro.total_re.value = (parseFloat(document.pedidos_pro.total_re.value) * (100 - parseFloat(document.pedidos_pro.dto2.value))) / 100;
        document.pedidos_pro.total_re.value = parseFloat(document.pedidos_pro.total_re.value).toFixed(<%=enc.EncodeForJavascript(NdecDiPedido)%>);

        document.pedidos_pro.total_recargo.value = (parseFloat(document.pedidos_pro.base_imponible.value) * parseFloat(document.pedidos_pro.recargo.value)) / 100;
        document.pedidos_pro.total_recargo.value = parseFloat(document.pedidos_pro.total_recargo.value).toFixed(<%=enc.EncodeForJavascript(NdecDiPedido)%>);

        if (document.pedidos_pro.IRPF_Total.value == "True" || document.pedidos_pro.IRPF_Total.value == "1")
            baseImp = parseFloat(document.pedidos_pro.base_imponible.value) + parseFloat(document.pedidos_pro.total_iva.value) +
                parseFloat(document.pedidos_pro.total_re.value) + parseFloat(document.pedidos_pro.total_recargo.value);
        else baseImp = document.pedidos_pro.base_imponible.value;

        document.pedidos_pro.total_irpf.value = (parseFloat(baseImp) * parseFloat(document.pedidos_pro.irpf.value)) / 100;
        document.pedidos_pro.total_irpf.value = parseFloat(document.pedidos_pro.total_irpf.value).toFixed(<%=enc.EncodeForJavascript(NdecDiPedido)%>);

        document.pedidos_pro.total_pedido.value =
            parseFloat(document.pedidos_pro.base_imponible.value) +
            parseFloat(document.pedidos_pro.total_iva.value) +
            parseFloat(document.pedidos_pro.total_re.value) +
            parseFloat(document.pedidos_pro.total_recargo.value) -
            parseFloat(document.pedidos_pro.total_irpf.value);
        document.pedidos_pro.total_pedido.value = parseFloat(document.pedidos_pro.total_pedido.value).toFixed(<%=enc.EncodeForJavascript(NdecDiPedido)%>);
    }

    function Acaja(npedido) {
        if (document.pedidos_pro.impcaja.value == "") document.pedidos_pro.impcaja.value = 0;
        if (isNaN(document.pedidos_pro.impcaja.value.replace(",", "."))) {
            window.alert("<%=LitMsgImporteNumerico%>");
            return false;
        }
        else {
            if (parseFloat(document.pedidos_pro.impcaja.value.replace(",", ".")) == 0) {
                window.alert("<%=LitErrImportePago%>");
                return false;
            }
        }
        if (document.pedidos_pro.ncaja.value == "") alert("<%=LitMsgCajaNoNulo%>");
        else {
            if (document.pedidos_pro.i_pago.value == "") alert("<%=LitMsgTipoPagoNoNulo%>");
            else {
                fr_PagosCuenta.document.pedidos_propagos.action = "pedidos_propagos.asp?mode=acaja&ndoc=" + npedido + "&impcaja=" + document.pedidos_pro.impcaja.value + "&i_pago=" + document.pedidos_pro.i_pago.value + "&ncaja=" + document.pedidos_pro.ncaja.value;
                fr_PagosCuenta.document.pedidos_propagos.submit();

            }
        }
    }

    //Añade un concepto al pedido
    function addConcepto(npedido) {
        if (document.pedidos_pro.descripcion.value == "") {
            alert("<%=LitMsgDesVacia%>");
            return;
        }
        if (isNaN(document.pedidos_pro.pvp.value.replace(",", "."))) {
            window.alert("<%=LitMsgImporteNumerico%>");
            return;
        }

        if (isNaN(document.pedidos_pro.cantidad.value.replace(",", ".")) || isNaN(document.pedidos_pro.descuento.value.replace(",", ".")) || isNaN(document.pedidos_pro.pvp.value.replace(",", "."))) {
            window.alert("<%=LitMsgCanPreDesNumerico%>");
            return;
        }

        //vamos a comprobar si hay límites de compra y si los hay si se cumplen
        alcanzado = document.pedidos_pro.alcanzado.value;    
        if (alcanzado != "a")  //el límite existe
        {
            importe = document.pedidos_pro.importe.value;
            importe = importe.replace(".", "");
            importe = parseFloat(importe.replace(",", "."));
            dto1 = document.pedidos_pro.desc1.value;
            dto22 = document.pedidos_pro.desc2.value;
            iva = document.pedidos_pro.iva.value;
            limite = document.pedidos_pro.limite.value;
            alcanzado = document.pedidos_pro.alcanzado.value;
            si_preguntar_riesgo = document.pedidos_pro.si_preguntar_riesgo.value;
            contrasenya = document.pedidos_pro.contrasenya.value;
            mes = document.pedidos_pro.mes.value;
            anyo = document.pedidos_pro.anyo.value;
            ndoc = document.pedidos_pro.h_npedido.value;

            if (comprobarLimite(limite, alcanzado, importe, dto1, dto22, iva, si_preguntar_riesgo, contrasenya, mes, anyo, ndoc, "NUEVO CONCEPTO") == 0)
                return;
        }

        //Asignar los valores a los campos del submarco de detalles
        fr_Conceptos.document.pedidos_procon.cantidad.value = document.pedidos_pro.cantidad.value;
        fr_Conceptos.document.pedidos_procon.descripcion.value = document.pedidos_pro.descripcion.value;
        fr_Conceptos.document.pedidos_procon.pvp.value = document.pedidos_pro.pvp.value;
        fr_Conceptos.document.pedidos_procon.descuento.value = document.pedidos_pro.descuento.value;
        fr_Conceptos.document.pedidos_procon.iva.value = document.pedidos_pro.iva.value;
        //Recargar el submarco de detalles
        fr_Conceptos.document.pedidos_procon.action = "pedidos_procon.asp?mode=first_save";
        fr_Conceptos.document.pedidos_procon.submit();
        //Limpiar los campos del formulario
        document.pedidos_pro.cantidad.value = "1";
        document.pedidos_pro.descripcion.value = "";
        document.pedidos_pro.pvp.value = "0";
        document.pedidos_pro.descuento.value = "0";
        document.pedidos_pro.iva.value = document.pedidos_pro.defaultIva.value;
        document.pedidos_pro.importe.value = "0";
        //Colocar el foco en el campo de cantidad.
        document.pedidos_pro.cantidad.focus();
        document.pedidos_pro.cantidad.select();
    }

    //Comprueba si el importe del pago es numerico
    function importepagoComp() {
        if (isNaN(document.pedidos_pro.importePago.value.replace(",", "."))) {
            window.alert("<%=LitErrImportePago2%>");
            return;
        }
    }

    //Calcula el importe de la línea de detalle del concepto.
    function ImporteDetalle() {
        if (parseFloat(document.pedidos_pro.pvp.value) < 0) {
            window.alert("<%=LitMsgPvPNoNegativo%>");
            document.pedidos_pro.pvp.value = 0;
        }
        if (isNaN(document.pedidos_pro.cantidad.value.replace(",", ".")) || isNaN(document.pedidos_pro.descuento.value.replace(",", ".")) || isNaN(document.pedidos_pro.pvp.value.replace(",", ".")))
            window.alert("<%=LitMsgCanPreDesNumerico%>");
        else {
            if (document.pedidos_pro.pvp.value == "") document.pedidos_pro.pvp.value = 0;
            if (document.pedidos_pro.cantidad.value == "") document.pedidos_pro.cantidad.value = 1;
            if (document.pedidos_pro.descuento.value == "") document.pedidos_pro.descuento.value = "0";

            pvpSinComas = document.pedidos_pro.pvp.value.replace(",", ".");
            cantidadSinComas = document.pedidos_pro.cantidad.value.replace(",", ".");
            dtoSinComas = document.pedidos_pro.descuento.value.replace(",", ".");
            pelas = parseFloat(cantidadSinComas) * parseFloat(pvpSinComas);
            pelas_descuento = (pelas * parseFloat(dtoSinComas)) / 100;
            importe = pelas - pelas_descuento;
            c_importe = importe.toString();
            document.pedidos_pro.cantidad.value = cantidadSinComas;
            document.pedidos_pro.descuento.value = dtoSinComas;
            document.pedidos_pro.importe.value = c_importe;
            document.pedidos_pro.pvp.value = pvpSinComas;
        }
    }

    //Añade un pago a cuenta.
    function addPago(npedido) {
        if (document.pedidos_pro.importePago.value == "") document.pedidos_pro.importePago.value = 0;
        if (document.pedidos_pro.fechaPago.value == "") {
            window.alert("<%=LitErrFechaPago%>");
            return;
        }

        if (isNaN(document.pedidos_pro.importePago.value.replace(",", "."))) {
            window.alert("<%=LitErrImportePago%>");
            return;
        }
        else {
            if (parseFloat(document.pedidos_pro.importePago.value.replace(",", ".")) == 0) {
                window.alert("<%=LitMsgImportePositivo%>");
                return;
            }
        }
        if (document.pedidos_pro.descripcionPago.value == "") {
            window.alert("<%=LitMsgDesVacia%>");
            return;
        }
        if (document.pedidos_pro.tipoPago.value == "") {
            window.alert("<%=LitMsgTipoPagoNoNulo%>");
            return;
        }

        if (!cambiarfecha(document.pedidos_pro.fechaPago.value, "Fecha Pago")) return;

        if (!checkdate(document.pedidos_pro.fechaPago)) {
            window.alert("<%=LitMsgFechaFecha%>");
            return;
        }

        //Asignar los valores a los campos del submarco de detalles
        fr_PagosCuenta.document.pedidos_propagos.fecha.value = document.pedidos_pro.fechaPago.value;
        fr_PagosCuenta.document.pedidos_propagos.importe.value = document.pedidos_pro.importePago.value;
        fr_PagosCuenta.document.pedidos_propagos.descripcion.value = document.pedidos_pro.descripcionPago.value;
        fr_PagosCuenta.document.pedidos_propagos.medio.value = document.pedidos_pro.tipoPago.value;
        //Recargar el submarco de pagos a cuenta
        fr_PagosCuenta.document.pedidos_propagos.action = "pedidos_propagos.asp?mode=first_save";
        fr_PagosCuenta.document.pedidos_propagos.submit();
        //Limpiar los campos del formulario
        var hoy = new Date();
        document.pedidos_pro.fechaPago.value = hoy.getDate() + "/" + (hoy.getMonth() + 1) + "/" + hoy.getFullYear();
        document.pedidos_pro.importePago.value = "0";
        document.pedidos_pro.descripcionPago.value = "";
        document.pedidos_pro.tipoPago.value = "";
        //Colocar el foco en el campo de cantidad.
        document.pedidos_pro.fechaPago.focus();
        document.pedidos_pro.fechaPago.select();
    }
    /*
    //Comprueba si la tecla pulsada es CTRL+S. Si es así guarda el registro.
    if(window.document.addEventListener)
    {
        window.document.addEventListener("keydown", callkeydownhandler, false);
    }
    else
    {
        window.document.attachEvent("onkeydown", callkeydownhandler);
    }
    
    //Comprueba si la tecla pulsada es CTRL+S. Si es así guarda el registro.
    function callkeydownhandler(evnt)
    {
        ev = (evnt) ? evnt : event;
        keyPressed(ev);
    }
    
    function keyPressed(e) {
        var keycode = e.keyCode;
        if (keycode==<%=TeclaGuardar%>) { //CTRL+S
            if (document.pedidos_pro.mode.value=="add" || document.pedidos_pro.mode.value=="edit") {
                if (document.pedidos_pro.fecha.value=="")
                {
                    window.alert("<%=LitMsgFechaNoNulo%>");
                    return;
                }
                if (document.pedidos_pro.serie.value=="")
                {
                    window.alert("<%=LitMsgSerieNoNulo%>");
                    return;
                }
                if (document.pedidos_pro.divisa.value=="")
                {
                    window.alert("<%=LitMsgDivisaNoNulo%>");
                    return;
                }
                if (document.pedidos_pro.nproveedor.value=="")
                {
                    window.alert("<%=LitMsgClienteNoNulo%>");
                    return;
                }
    
                if (!cambiarfecha(document.pedidos_pro.fecha.value,"FECHA PEDIDO")) return false;
    
                if (!checkdate(document.pedidos_pro.fecha))
                {
                    window.alert("<%=LitMsgFechaFecha%>");
                    return;
                }
    
                if (!cambiarfecha(document.pedidos_pro.fecha_entrega.value,"FECHA ENTREGA")) return false;
    
                if (!checkdate(document.pedidos_pro.fecha_entrega))
                {
                    window.alert("<%=LitMsgFechaFecha%>");
                    return;
                }
            	
                if (!cambiarfecha(document.pedidos_pro.salida.value,"FECHA PAGO")) return false;
    
                if (!checkdate(document.pedidos_pro.salida))
                {
                    window.alert("<%=LitMsgFechaFecha%>");
                    return;
                }
            	
                // AMP: comprobacion campo factor de cambio.
                factcambio=document.pedidos_pro.nfactcambio.value.replace(",","."); 		
                if (!/^([0-9])*[.]?[0-9]*$/.test(factcambio))
                { 
                    alert("<%=LitMsgFactCambioI%>"); 
                    return false;
                }
                if (document.pedidos_pro.nfactcambio.value=="")
                {
                    alert("<%=LitMsgFactCambioI%>"); 
                    return false;
                }
    
                // JMA 16/12/04. Campos personalizables
                if (document.pedidos_pro.si_campo_personalizables.value==1)
                {
                    num_campos=document.pedidos_pro.num_campos.value;
                    respuesta=comprobarCampPerso("",num_campos,"pedidos_pro");
                    if(respuesta!=0)
                    {
                        titulo="titulo_campo" + respuesta;
                        tipo="tipo_campo" + respuesta;
                        titulo=document.pedidos_pro.elements[titulo].value;
                        tipo=document.pedidos_pro.elements[tipo].value;
                        if (tipo==4) nomTipo="<%=LitTipoNumerico%>";
                        else if (tipo==5)
                        {
                            nomTipo="<%=LitTipoFecha%>";
                        }
                        window.alert("<%=LitMsgCampo%> " + titulo + " <%=LitMsgTipo%> " + nomTipo);
                        return false;
                    }
                }
    
                switch (document.pedidos_pro.mode.value)
                {
                    case "add":
                        document.pedidos_pro.action="pedidos_pro.asp?mode=first_save";
                        break;
    
                    case "edit":
                        document.pedidos_pro.action="pedidos_pro.asp?mode=save&ndoc=" + document.pedidos_pro.h_npedido.value;
                        break;
                }
                //ricardo 25/4/2003 para que salga el mensaje de que se ha cambiado las propiedades del documento
                // y que puede afectar al importe de los detalles
                nempresa="<%=session("ncliente")%>";
                recalcular_importes=1;
                if(document.pedidos_pro.mode.value=="edit")
                {
                    if (document.pedidos_pro.h_nproveedor.value!=(nempresa + document.pedidos_pro.nproveedor.value) ||
                        document.pedidos_pro.h_fecha.value!=document.pedidos_pro.fecha.value ||
                        document.pedidos_pro.h_divisa.value!=document.pedidos_pro.olddivisa.value)
                    {
                        if (window.confirm("<%=LitMsgCamPropDocCamPrec%>")==false) recalcular_importes=0;
    
                    }
                }
                document.pedidos_pro.action=document.pedidos_pro.action + "&recalcular_importes=" + recalcular_importes;
                document.pedidos_pro.submit();
                parent.botones.document.location="../compras/pedidos_pro_bt.asp?mode=browse";
            }
            else
            { //Mode=browse.
                numTab = getTabsSelected();
    
                //Comprobamos si estamos añadiendo conceptos.
                if (numTab == 1) addConcepto(document.pedidos_pro.h_nalbaran.value);
    
                //Comprobamos si estamos añadiendo conceptos.
                if (numTab == 2) addPago(document.pedidos_pro.h_nalbaran.value);
            }
        }
    }
    */
    /*RGU 24/5/2006*/
    function pedir(npedido) {
        if (window.confirm("<%=LitConfPedPro%>"))
            if (document.getElementById("waitBoxOculto").style.visibility == "hidden") {
                document.getElementById("redi").style.visibility = "hidden";
                document.getElementById("waitBoxOculto").style.visibility = "visible";
                document.pedidos_pro.action = "pedidos_pro.asp?mode=browse&submode=pedir_pro&npedido=" + npedido;
                document.pedidos_pro.submit();
            }
    }

    //  GPD (05/03/2007).
    function mostrarCondicionesCompra(strCodProveedor) {
        var strCadena = '';

        if (document.frDetallesIns.pedidos_prodetins.valido.value == 1) {
            strCadena = strCadena + '../central.asp?pag1=productos/articulos_compra_condiciones.asp';
            strCadena = strCadena + '&ndoc=<%=session("ncliente")%>' + document.frDetallesIns.pedidos_prodetins.referencia.value;
            strCadena = strCadena + '&nproveedor=' + strCodProveedor;
            strCadena = strCadena + '&mode=browse&pag2=productos/articulos_compra_condiciones_bt.asp&titulo=<%=LitCondicionesCompra%>';
            AbrirVentana(strCadena, 'P', 400, 750);
        }
        else alert('<%=LitMsgReferenciaIncorrecta%>');
    }

    function MasDet(sentido, lote, firstReg, lastReg, campo, criterio, texto, firstRegAll, lastRegAll) {
        fr_Detalles.document.pedidos_prodet.action = "pedidos_prodet.asp?mode=browse&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&firstReg=" + firstReg + "&lastReg=" + lastReg + "&firstRegAll=" + firstRegAll + "&lastRegAll=" + lastRegAll + "&almacenSerie=<%=almacenSerie %>&almacenTPV=<%=almacenTPV %>";
        fr_Detalles.document.pedidos_prodet.submit();
    }

    //FUNCIONES PARA SUMAR DIAS A UNA FECHA
    var aFinMes = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

    function finMes(nMes, nAno) {
        return aFinMes[nMes - 1] + (((nMes == 2) && (nAno % 4) == 0) ? 1 : 0);
    }

    function padNmb(nStr, nLen, sChr) {
        var sRes = String(nStr);
        for (var i = 0; i < nLen - String(nStr).length; i++)
            sRes = sChr + sRes;
        return sRes;
    }

    function makeDateFormat(nDay, nMonth, nYear) {
        var sRes;
        sRes = padNmb(nDay, 2, "0") + "/" + padNmb(nMonth, 2, "0") + "/" + padNmb(nYear, 4, "0");
        return sRes;
    }

    function incDate(sFec0) {
        var nDia = parseInt(sFec0.substr(0, 2), 10);
        var nMes = parseInt(sFec0.substr(3, 2), 10);
        var nAno = parseInt(sFec0.substr(6, 4), 10);
        nDia += 1;
        if (nDia > finMes(nMes, nAno)) {
            nDia = 1;
            nMes += 1;
            if (nMes == 13) {
                nMes = 1;
                nAno += 1;
            }
        }
        return makeDateFormat(nDia, nMes, nAno);
    }

    function decDate(sFec0) {
        var nDia = Number(sFec0.substr(0, 2));
        var nMes = Number(sFec0.substr(3, 2));
        var nAno = Number(sFec0.substr(6, 4));
        nDia -= 1;
        if (nDia == 0) {
            nMes -= 1;
            if (nMes == 0) {
                nMes = 12;
                nAno -= 1;
            }
            nDia = finMes(nMes, nAno);
        }
        return makeDateFormat(nDia, nMes, nAno);
    }

    function addToDate(sFec0, sInc) {
        var nInc = Math.abs(parseInt(sInc));
        var sRes = sFec0;
        if (parseInt(sInc) >= 0) for (var i = 0; i < nInc; i++) sRes = incDate(sRes);
        else for (var i = 0; i < nInc; i++) sRes = decDate(sRes);
        return sRes;
    }

    function recalcF1() {
        with (document.formulario) {
            fecha1.value = addToDate(fecha0.value, increm.value);
        }
    }

    function CalculaFechaPago() {
        document.pedidos_pro.formas_pagodias.value = document.pedidos_pro.formas_pago.value;
        if (document.pedidos_pro.fecha.value != "" && document.pedidos_pro.formas_pagodias.value != "" && checkdate(document.pedidos_pro.fecha)) {
            var list = document.pedidos_pro.formas_pagodias;
            var diasdepago = list.options[list.selectedIndex].text;
            document.pedidos_pro.salida.value = addToDate(document.pedidos_pro.fecha.value, diasdepago);
        }
        else document.pedidos_pro.salida.value = document.pedidos_pro.fecha.value;
    }

    //*** i AMP Nueva función cambiar divisa con factor de cambio incorporado.
    var ret_tra = "";
    var ret_tra2 = "";
    function cambiardivisa(mBase) {
        document.pedidos_pro.h_divisa.value = document.pedidos_pro.divisa.value;

        var divisa = document.pedidos_pro.divisa.value;
        if (divisa == mBase) {
            parent.pantalla.document.getElementById("tdfactcambio").style.display = "none";
            parent.pantalla.document.pedidos_pro.nfactcambio.value = "1";
        }
        else {
            parent.pantalla.document.getElementById("tdfactcambio").style.display = "";
            ret_tra = "";
            if (!enProcesoFC && httpFC) {
                var timestamp = Number(new Date());
                var url = "../select_factcambio.asp?divisa=" + divisa;
                httpFC.open("GET", url, false);
                httpFC.onreadystatechange = handleHttpResponseFC;
                enProcesoFC = false;
                httpFC.send(null);
            }

        }
        parent.pantalla.document.pedidos_pro.h_divisa.value = divisa;
        parent.pantalla.document.pedidos_pro.divisafc.value = divisa;
    }

    function handleHttpResponseFC() {
        if (httpFC.readyState == 4) {
            if (httpFC.status == 200) {
                if (httpFC.responseText.indexOf('invalid') == -1) {
                    // Armamos un array, usando la coma para separar elementos
                    results = httpFC.responseText;
                    enProcesoFC = false;
                    ret_tra = unescape(results)
                    spfc = ret_tra.split(";");
                    factcambio = spfc[0];
                    parent.pantalla.document.pedidos_pro.nfactcambio.value = factcambio;
                    ret_tra2 = "";
                    var divisa = document.pedidos_pro.divisa.value;
                    if (!enProcesoFC2 && httpFC2) {
                        var timestamp = Number(new Date());
                        var url = "../select_factcambio.asp?divisa=" + divisa + "&que=abreviatura";
                        httpFC2.open("GET", url, false);
                        httpFC2.onreadystatechange = handleHttpResponseFC2;
                        enProcesoFC2 = false;
                        httpFC2.send(null);
                    }
                }
            }
        }
    }
    function handleHttpResponseFC2() {
        if (httpFC2.readyState == 4) {
            if (httpFC2.status == 200) {
                if (httpFC2.responseText.indexOf('invalid') == -1) {
                    // Armamos un array, usando la coma para separar elementos
                    results = httpFC2.responseText;
                    enProcesoFC2 = false;
                    ret_tra2 = unescape(results);
                    spfc2 = ret_tra2.split(";");
                    otraabrev = spfc2[0];
                    parent.pantalla.document.getElementById("idfactcambioexpl").innerHTML = otraabrev;
                }
            }
        }
    }

    function getHTTPObjectFC() {
        var xmlhttp;
        if (!xmlhttp && typeof XMLHttpRequest != 'undefined') {
            try {
                xmlhttp = new XMLHttpRequest();
            }
            catch (e) { xmlhttp = false; }
        }
        return xmlhttp;
    }
    var enProcesoFC = false; // lo usamos para ver si hay un proceso activo
    var httpFC = getHTTPObjectFC(); // Creamos el objeto XMLHttpRequest
    var enProcesoFC2 = false; // lo usamos para ver si hay un proceso activo
    var httpFC2 = getHTTPObjectFC(); // Creamos el objeto XMLHttpRequest

    function comprobarFactorCambio() {
        numero = document.pedidos_pro.nfactcambio.value;
        document.pedidos_pro.nfactcambio.value = numero.replace(",", ".")
        numero2 = document.pedidos_pro.nfactcambio.value;
        if (!/^([0-9])*[.]?[0-9]*$/.test(numero2)) { alert("<%=LitMsgFactCambioI%>"); }
        if (document.pedidos_pro.nfactcambio.value == "") { alert("<%=LitMsgFactCambioI%>"); }
    }
    //*** f AMP

    function CrearPedidoCentral(npedido) {
        if (window.confirm("<%=LITMSG_ENVIARPEDACENTRAL%>")) {
            document.getElementById("waitBoxOculto").style.visibility = "visible";
            document.pedidos_pro.action = "pedidos_pro.asp?mode=browse&submode=crearPedCentral&npedido=" + npedido;
            document.pedidos_pro.submit();
        }
    }

    function Redimensionar() {
        var alto = jQuery(window).height();
        var diference = 425;
        var dir_default = 150;

        if (alto > dir_default) {
            if (alto - diference > dir_default) {
                jQuery("#frDetalles").attr("height", alto - diference);
                jQuery("#frConceptos").attr("height", alto - diference);
                jQuery("#frPagosCuenta").attr("height", alto - diference);
                jQuery("#frPedcli").attr("height", alto - diference + 55);
            }
            else {
                jQuery("#frDetalles").attr("height", dir_default);
                jQuery("#frConceptos").attr("height", dir_default);
                jQuery("#frPagosCuenta").attr("height", dir_default);
                jQuery("#frPedcli").attr("height", dir_default);
            }
        }
        else {
            jQuery("#frDetalles").attr("height", dir_default);
            jQuery("#frConceptos").attr("height", dir_default);
            jQuery("#frPagosCuenta").attr("height", dir_default);
            jQuery("#frPedcli").attr("height", dir_default);
        }
    }

    function RoundNumValue(obj, dec) {
        obj.value = obj.value.replace(',', '.');
        var valor = parseFloat(obj.value);
        if (valor != 0) obj.value = valor.toFixed(dec);
    }

    jQuery(window).resize(function () { Redimensionar(); });
</script>

<body onload="self.status='';" class="BODY_ASP">
<%function CalcularNumDocumentoDSN(nserie,fecha,dsnCentral)
	set rstCalcNumDoc = Server.CreateObject("ADODB.Recordset")
	
    set commandCND = nothing
    set connCND = Server.CreateObject("ADODB.Connection")
    set commandCND =  Server.CreateObject("ADODB.Command")

    connCND.open dsnCentral
    connCND.cursorlocation=2
    commandCND.ActiveConnection =connCND
    commandCND.CommandTimeout = 60
    commandCND.CommandText= "select * from series with(updlock) where nserie= ?"
    commandCND.CommandType = adCmdText
    commandCND.Parameters.Append commandCND.CreateParameter("@nserie",adVarChar,adParamInput,10, nserie)

    rstCalcNumDoc.Open commandCND, , adOpenKeyset, adLockOptimistic
    'rstCalcNumDoc.Open strSelect, dsnCentral,adOpenKeyset,adLockOptimistic


	siguiente=rstCalcNumDoc("contador")+1
	'Actualizar el nº de cliente de CONFIGURACION.
	rstCalcNumDoc("contador")=siguiente
	rstCalcNumDoc("ultima_fecha")=date
	rstCalcNumDoc.Update
    connCND.Close
    set connCND = nothing
    set commandCND = nothing
	'rstCalcNumDoc.Close

	set rstCalcNumDoc=nothing

	ano=right(cstr(year(fecha)),2)
	CalcularNumDocumentoDSN=nserie & ano & completar(trim(cstr(siguiente)),6,"0")
end function

'******************************************************************************
'Actualiza los precios del pedido                                              
'******************************************************************************
sub PreciosPed(npedido,dsnCentral)

    set commandS = nothing
    set connS = Server.CreateObject("ADODB.Connection")
    set commandS =  Server.CreateObject("ADODB.Command")

    connS.open dsnCentral
    connS.cursorlocation=2
    commandS.ActiveConnection =connS
    commandS.CommandTimeout = 60
    commandS.CommandText= "select * from pedidos_cli with(rowlock) where npedido= ?"
    commandS.CommandType = adCmdText
    commandS.Parameters.Append commandS.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

    rstSelect.Open commandS, , adOpenKeyset, adLockOptimistic

    'rstSelect.cursorlocation=2
	'rstSelect.open "select * from pedidos_cli with(rowlock) where npedido='" & npedido & "'",dsnCentral,adOpenKeyset,adLockOptimistic
	
	'Miramos si el cliente tiene recargo de equivalencia

    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open dsnCentral
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "Select re from clientes with(nolock) where ncliente= ?"
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@ncliente", adChar,adParamInput,10, rstSelect("ncliente"))

    set rstAux = commandAux.Execute

    'rstAux.cursorlocation=3
	'rstAux.open "Select re from clientes with(nolock) where ncliente='" + rstSelect("ncliente") + "'",dsnCentral
    
    if not rstAux.eof then
	    TieneRE=rstAux("re")
    else
        TieneRE=0
    end if
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	'rstAux.close
	
	'desglose de los detalles por tipo de IVA aplicado
	'seleccion="SELECT IVA, SUM(IMPORTE) AS IMPORTEBRUTO, RE, SUM(((IMPORTE * IVA) / 100)) AS TOTALIVA, "
    'seleccion=seleccion+"SUM((IMPORTE * RE) / 100) AS TOTALRE FROM DETALLES_PED_CLI with(nolock) "
	'seleccion=seleccion+"WHERE (NPEDIDO ='" & npedido & "') "
	'seleccion=seleccion+"GROUP BY IVA, RE ORDER BY IVA"
    'rstIvas.cursorlocation=3
	'rstIvas.open seleccion,dsnCentral

    set commandIvas = nothing
    set connIvas = Server.CreateObject("ADODB.Connection")
    set commandIvas =  Server.CreateObject("ADODB.Command")

    connIvas.open dsnCentral
    connIvas.cursorlocation=3
    commandIvas.ActiveConnection =connIvas
    commandIvas.CommandTimeout = 60
    commandIvas.CommandText= "SELECT IVA, SUM(IMPORTE) AS IMPORTEBRUTO, RE, SUM(((IMPORTE * IVA) / 100)) AS TOTALIVA, SUM((IMPORTE * RE) / 100) AS TOTALRE FROM DETALLES_PED_CLI with(nolock) WHERE (NPEDIDO = ?) GROUP BY IVA, RE ORDER BY IVA"
    commandIvas.CommandType = adCmdText
    commandIvas.Parameters.Append commandIvas.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

    set rstIvas = commandIvas.Execute

	SumImporteBruto=0
	SumTotalDto=0
	SumBaseImponible=0
	SumTotalIva=0
	SumTotalRE=0
	SumTotalRF=0
	SumTotalImporte=0
	while not rstIvas.EOF
		SumImporteBruto=SumImporteBruto + rstIvas("ImporteBruto")
		dto1=miround((null_z(rstIvas("ImporteBruto"))*null_z(rstSelect("descuento")))/100,2)
		dto2=miround(((null_z(rstIvas("ImporteBruto"))-dto1)*null_z(rstSelect("descuento2")))/100,2)
		total_descuento=dto1+dto2
		SumTotalDto=SumTotalDto + null_z(total_descuento)
		
		base_imponible=null_z(rstIvas("ImporteBruto"))-null_z(total_descuento)
		SumBaseImponible=SumBaseImponible + null_z(base_imponible)
		total_iva=miround((null_z(base_imponible)*rstIvas("iva"))/100,2)
		SumTotalIva=SumTotalIva + null_z(total_iva)
		if TieneRE <> 0 then
            're=d_lookup("re","tipos_iva","tipo_iva='" & rstIvas("iva") & "'",dsnCentral)
            reSelect = "select re from tipos_iva with(nolock) where tipo_iva= ?"
            re=DLookupP1(reSelect, rstIvas("iva")&"", 139, 4, dsnCentral)
		else
			re=0
		end if
		total_re=(null_z(base_imponible)*re)/100
		SumTotalRE=SumTotalRE + null_z(total_re)
		total_recargo=miround((null_z(base_imponible)*null_z(rstSelect("recargo")))/100,2)
		SumTotalRF=SumTotalRF + null_z(total_recargo)
		total=null_z(base_imponible)+null_z(total_iva)+null_z(total_re)+null_z(total_recargo)
		SumTotalImporte=SumTotalImporte + null_z(total)
		rstIvas.Movenext
	wend
    connIvas.Close
    set connIvas = nothing
    set commandIvas = nothing
	'rstIvas.close
	
	'OTRA VEZ PERO PARA LA TABLA DE CONCEPTOS DEL PEDIDO
	
	'desglose por tipo de IVA aplicado para conceptos del pedido
	'seleccion="SELECT IVA, SUM(IMPORTE) AS IMPORTEBRUTO, RE, SUM(((IMPORTE * IVA) / 100)) AS TOTALIVA, "
    'seleccion=seleccion+"SUM((IMPORTE * RE) / 100) AS TOTALRE FROM conceptos_ped_cli with(nolock) "
	'seleccion=seleccion+"WHERE (npedido ='" & npedido & "')"
	'seleccion=seleccion+"GROUP BY IVA, RE ORDER BY IVA"
    'rstIvas.cursorlocation=3
	'rstIvas.open seleccion,dsnCentral

    set commandIvas = nothing
    set connIvas = Server.CreateObject("ADODB.Connection")
    set commandIvas =  Server.CreateObject("ADODB.Command")

    connIvas.open dsnCentral
    connIvas.cursorlocation=3
    commandIvas.ActiveConnection =connIvas
    commandIvas.CommandTimeout = 60
    commandIvas.CommandText= "SELECT IVA, SUM(IMPORTE) AS IMPORTEBRUTO, RE, SUM(((IMPORTE * IVA) / 100)) AS TOTALIVA, SUM((IMPORTE * RE) / 100) AS TOTALRE FROM CONCEPTOS_PED_CLI with(nolock) WHERE (NPEDIDO = ?) GROUP BY IVA, RE ORDER BY IVA"
    commandIvas.CommandType = adCmdText
    commandIvas.Parameters.Append commandIvas.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

    set rstIvas = commandIvas.Execute

	while not rstIvas.EOF
		SumImporteBruto=SumImporteBruto + rstIvas("ImporteBruto")
		dto1=miround((null_z(rstIvas("ImporteBruto"))*null_z(rstSelect("descuento")))/100,2)
		dto2=miround(((null_z(rstIvas("ImporteBruto"))-dto1)*null_z(rstSelect("descuento2")))/100,2)
		total_descuento=dto1+dto2
		SumTotalDto=SumTotalDto + null_z(total_descuento)
		
		base_imponible=null_z(rstIvas("ImporteBruto"))-null_z(total_descuento)
		SumBaseImponible=SumBaseImponible + null_z(base_imponible)
		total_iva=miround((null_z(base_imponible)*rstIvas("iva"))/100,2)
		SumTotalIva=SumTotalIva + null_z(total_iva)
		if TieneRE <> 0 then
			''re=d_lookup("re","tipos_iva","tipo_iva='" & rstIvas("iva") & "'",dsnCentral)
            reSelect = "select re from tipos_iva with(nolock) where tipo_iva= ?"
            re=DLookupP1(reSelect, rstIvas("iva")&"", 139, 4, dsnCentral)
		else
			re=0
		end if
		total_re=miround((null_z(base_imponible)*re)/100,2)
		SumTotalRE=SumTotalRE + null_z(total_re)
		total_recargo=miround((null_z(base_imponible)*null_z(rstSelect("recargo")))/100,2)
		SumTotalRF=SumTotalRF + null_z(total_recargo)
		total=null_z(base_imponible)+null_z(total_iva)+null_z(total_re)+null_z(total_recargo)
		SumTotalImporte=SumTotalImporte + null_z(total)
		rstIvas.Movenext
	wend
    connIvas.Close
    set connIvas = nothing
    set commandIvas = nothing
	'rstIvas.close
	
	if not rstSelect.eof then
		rstSelect("importe_bruto")=SumImporteBruto
		rstSelect("total_descuento")=SumTotalDto
		rstSelect("base_imponible")=SumBaseImponible
		rstSelect("total_iva")=SumTotalIva
		rstSelect("total_re")=SumTotalRE
		rstSelect("total_recargo")=SumTotalRF
		rstSelect("total_pedido")=SumTotalImporte
		rstSelect.update
	end if
    connS.Close
    set connS = nothing
    set commandS = nothing
	rstSelect.close
end sub

'***************************************************************************************************************
' Funcion para crear la cabecera de un pedido de cliente
'***************************************************************************************************************

function CabeceraPedido(npedido,nserie,fecha,ncliente,nempresaCentral,dsnCentral)
	if npedido="" then
		'Crear un nuevo registro.
		rst.AddNew
		'******************** Manejo de domicilios
		Dom=Domicilios("VENTAS","PED_ENV_CLI",ncliente,rst)
		
		'Gestionamos las direcciones de envío que marca en el pedido de franquicia
		'domPedCom = d_lookup("codigo","domicilios","tipo_domicilio='PED_ENV_PROV' and codigo='"& rstAux2("dir_envio") &"' ",session("dsn_cliente"))
        domPedComSelect = "select codigo from domicilios with(nolock) where tipo_domicilio='PED_ENV_PROV' and codigo= ?"
        domPedCom = DLookupP1(domPedComSelect, rstAux2("dir_envio")&"", adInteger, 4, session("dsn_cliente"))
        
        'rstCli.cursorlocation=3
		'rstCli.open "select PERTENECE, TIPO_DOMICILIO, DOMICILIO, CP, POBLACION, PROVINCIA, PAIS, TELEFONO, A_LA_ATENCION from domicilios where codigo='"& domPedCom &"'",session("dsn_cliente")

        set commandCli = nothing
        set connCli = Server.CreateObject("ADODB.Connection")
        set commandCli =  Server.CreateObject("ADODB.Command")

        connCli.open session("dsn_cliente")
        connCli.cursorlocation=3
        commandCli.ActiveConnection =connCli
        commandCli.CommandTimeout = 60
        commandCli.CommandText= "select PERTENECE, TIPO_DOMICILIO, DOMICILIO, CP, POBLACION, PROVINCIA, PAIS, TELEFONO, A_LA_ATENCION from domicilios where codigo= ?"
        commandCli.CommandType = adCmdText
        commandCli.Parameters.Append commandCli.CreateParameter("@codigo", adInt,adParamInput, 4, domPedCom)

        set rstCli = commandCli.Execute

		if not rstCli.EOF then
            domicilioEnvio = d_lookup("codigo","domicilios"," tipo_domicilio='PED_ENV_CLI' and domicilio='"& rstCli("DOMICILIO") &"' and cp='"& rstCli("CP") &"' and poblacion='"& rstCli("POBLACION") &"' and pais='"& rstCli("PAIS") &"' and telefono='"& rstCli("TELEFONO") &"' and a_la_atencion='"& rstCli("A_LA_ATENCION") &"' ",dsnCentral)
            'domicilioEnvioSelect = "select codigo from domicilios with(nolock) where tipo_domicilio='PED_ENV_CLI' and domicilio = ? and cp= ? and poblacion= ? and pais= ? and telefono= ? and a_la_atencion= ?"
            'domicilioEnvio = DLookupP6(domicilioEnvioSelect, rstCli("DOMICILIO") &"", adVarchar, 100,  rstCli("CP") &"", adVarchar, 10, rstCli("POBLACION") &"", adVarchar, 50,  rstCli("PAIS") &"", adVarchar, 30, rstCli("TELEFONO") &"", adVarchar, 20, rstCli("A_LA_ATENCION") &"", adVarchar, 50, dsnCentral )
            if domicilioEnvio&"">"" then
                'Asignamos la dirección ya existente
                rst("dir_envio")=domicilioEnvio
            else
                'Creamos nueva dirección en central y asignamos al pedido de venta
                'sql = "insert into domicilios (PERTENECE, TIPO_DOMICILIO, DOMICILIO, CP, POBLACION, PROVINCIA, PAIS, TELEFONO, A_LA_ATENCION) values " & _
                        '"('"&ncliente&"','PED_ENV_CLI','"&rstCli("DOMICILIO")&"','"&rstCli("CP")&"','"&rstCli("POBLACION")&"','"&rstCli("PROVINCIA")&"','"&rstCli("PAIS")&"','"&rstCli("TELEFONO")&"','"&rstCli("A_LA_ATENCION")&"')"
                'rstTMP.open sql,dsnCentral,adOpenKeyset,adLockOptimistic

                
                set commandTMP = nothing
                set connTMP = Server.CreateObject("ADODB.Connection")
                set commandTMP =  Server.CreateObject("ADODB.Command")

                connTMP.open dsnCentral
                connTMP.cursorlocation=3
                commandTMP.ActiveConnection =connTMP
                commandTMP.CommandTimeout = 60
                commandTMP.CommandText= "insert into domicilios (PERTENECE, TIPO_DOMICILIO, DOMICILIO, CP, POBLACION, PROVINCIA, PAIS, TELEFONO, A_LA_ATENCION) values (?,'PED_ENV_CLI', ?, ?, ?, ?, ?, ?, ?)"
                commandTMP.CommandType = adCmdText
                commandTMP.Parameters.Append commandTMP.CreateParameter("@pertenece", adVarChar,adParamInput, 55, ncliente)
                commandTMP.Parameters.Append commandTMP.CreateParameter("@domicilio", adVarchar,adParamInput, 100, rstCli("DOMICILIO"))
                commandTMP.Parameters.Append commandTMP.CreateParameter("@cp", adVarchar,adParamInput, 10, rstCli("CP"))
                commandTMP.Parameters.Append commandTMP.CreateParameter("@poblacion", adVarchar,adParamInput, 50, rstCli("POBLACION"))
                commandTMP.Parameters.Append commandTMP.CreateParameter("@provincia", adVarchar,adParamInput, 50, rstCli("PROVINCIA"))
                commandTMP.Parameters.Append commandTMP.CreateParameter("@pais", adVarchar,adParamInput, 30, rstCli("PAIS"))
                commandTMP.Parameters.Append commandTMP.CreateParameter("@telefono", adVarchar,adParamInput, 20, rstCli("TELEFONO"))
                commandTMP.Parameters.Append commandTMP.CreateParameter("@alaatencion", adVarchar, adParamInput, 50, rstCli("A_LA_ATENCION"))

                rstTMP.Open commandTMP, , adOpenKeyset, adLockOptimistic

                domicilioEnvio = d_max("codigo","domicilios","pertenece='" & ncliente & "'",dsnCentral)
                rst("dir_envio")=domicilioEnvio
            end if
        end if
        connCli.Close
        set connCli = nothing
        set commandCli = nothing
        'rstCli.Close
	end if
	
	'tarifaCli = d_lookup("tarifa","clientes","ncliente='" & ncliente & "'",dsnCentral)
    tarifaCliSelect = "select tarifa from clientes with(nolock) where ncliente= ?"
    tarifaCli = DLookupP1(tarifaCliSelect, ncliente &"", adChar, 10, dsnCentral)
	'fpago = d_lookup("fpago","clientes","ncliente='" & ncliente & "'",dsnCentral)
    fpagoSelect = "select fpago from clientes with(nolock) where ncliente= ?"
    fpago = DLookupP1(fpagoSelect, ncliente &"", adChar, 10, dsnCentral)
	'cuentaBancoCliente = d_lookup("ncuenta","clientes","ncliente='" & ncliente & "'",dsnCentral)
    cuentaBancoClienteSelect = "select ncuenta from clientes with(nolock) where ncliente= ?"
    cuentaBancoCliente = DLookupP1(cuentaBancoClienteSelect, ncliente &"", adChar, 10, dsnCentral)

	'Asignar los nuevos valores a los campos del recordset.
	rst("valorado")=rstAux2("valorado")
	rst("serie")=nserie
	rst("ncliente")=ncliente
	rst("descuento")=rstAux2("descuento")
	rst("descuento2")=rstAux2("descuento2")
	rst("importe_bruto")= rstAux2("importe_bruto")
	rst("total_descuento")=rstAux2("total_descuento")
	rst("base_imponible")=rstAux2("base_imponible")
	rst("total_iva")=rstAux2("total_iva")
	rst("recargo")=rstAux2("recargo")
	rst("total_recargo")=rstAux2("total_recargo")
	rst("total_re")=rstAux2("total_re")
	rst("total_pedido")=rstAux2("total_pedido")
	rst("con_orden")=0
	rst("en_empresa")=0
	rst("facturado")=0
	rst("ahora")=0
	rst("transportista")=rstAux2("transportista")
	rst("portes")=rstAux2("portes")
	rst("su_npedido")=trimCodEmpresa(doc)
	rst("observaciones")=rstAux2("observaciones")
	rst("incoterms")=rstAux2("incoterms")
	rst("fob")=rstAux2("fob")
	rst("tarifa")=tarifaCli
	rst("forma_pago")=fpago
    rst("ncuenta")=cuentaBancoCliente
    
	rst("fecha")=fecha
	rst("divisa")=nempresaCentral & trimCodEmpresa(rstAux2("divisa"))
	rst("edi")=entrada
	'Actualizar el registro.
	'Obtener el siguiente nº de documento de la tabla series.
	if npedido="" then
		SigDoc=CalcularNumDocumentoDSN(nserie,fecha,dsnCentral)
		rst("npedido")=SigDoc
	end if
	rst.Update
	if err.number<>0 then
        ha_habido_un_error=1
    end if
	CabeceraPedido=SigDoc
end function

'***************************************************************************************************************
'Funcion para comprobar detalles de pedido de compra con los precios de la central
'***************************************************************************************************************
Function ComprobarPreciosEnCentral(ndocumento,npedido,dsnCentral,ncliente,rpc)
    'tarifaCliente=d_lookup("tarifa","clientes","ncliente='" & ncliente & "'",dsnCentral)
    tarifaClienetSelect = "select tarifa from clientes with(nolock) where ncliente= ?"
    tarifaCliente= DLookupP1(tarifaClienteSelect, ncliente &"", adChar, 5, dsnCentral)
    StrSelDetPedPro="select * from detalles_ped_pro where npedido= ? "
    if rpc & "">"" then
        StrSelDetPedPro=StrSelDetPedPro & " and cantidadpend<>0" 
    end if
    StrSelDetPedPro=StrSelDetPedPro & " order by item" 
    
    'rstAux.open StrSelDetPedPro,cnn,adOpenKeyset,adLockOptimistic

    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open cnn
    connAux.cursorlocation=2
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= StrSelDetPedPro
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, ndocumento)

    rstAux.Open commandAux, , adOpenKeyset, adLockOptimistic

    if tarifaCliente&"">"" or rpc & "">"" then
	    BDFranq=encontrar_datos_dsn(cnn,"Initial Catalog=")
        BDCentral=encontrar_datos_dsn(dsnCentral,"Initial Catalog=")
        
        eliminar="if exists (select * from sysobjects where id = object_id('egesticet.[?]') and sysstat " & _
				     " & 0xf = 3) drop table egesticet.[" & session("usuario") & "]"
	   'rstTMP.open eliminar,cnn,adUseClient,adLockReadOnly

        set commandDel = nothing
        set connDel = Server.CreateObject("ADODB.Connection")
        set commandDel =  Server.CreateObject("ADODB.Command")

        connDel.open cnn
        'connDel.cursorlocation=2
        commandDel.ActiveConnection =connDel
        commandDel.CommandTimeout = 60
        commandDel.CommandText= eliminar
        commandDel.CommandType = adCmdText
        commandDel.Parameters.Append commandDel.CreateParameter("@id", adInt, adParamInput, 4, session("usuario"))

        commandDel.Execute
        
        connDel.Close
        set connDel = nothing
        set commandDel = nothing
      
        crear ="CREATE TABLE egesticet.[" & session("usuario") & "]" & _
                  "(item smallint, " & _
                  "referencia varchar(30), " & _
                  "pvpdto money)"
        rstTMP.open crear,cnn,adUseClient,adLockReadOnly

        ' Insertamos la fila
        if rpc & ""="" or (rpc & "">"" and tarifaCliente & "">"") then
            strselect = "select dpp.item,dpp.referencia, " & _
	                        "case when es_dto=0 then round(p.pvpdto,"& DEC_PREC &") else  " & _
	                        "case when es_dto=1 then round(a.pvp+((p.pvpdto*0.01)*a.pvp),"& DEC_PREC &") else " & _
	                        "case when es_dto=2 then round(a.importe+((p.pvpdto*0.01)*a.importe),"& DEC_PREC &") else round(a.pvp,"& DEC_PREC &") end end end as pvpdto " & _
                        "from "& BDFranq &"..detalles_ped_pro as dpp with(nolock) " & _
	                        "inner join "& BDCentral &"..articulos as a with(nolock) on a.referencia like '"& empresa_sup &"%' " & _
		                        "and substring(dpp.referencia,6,Len(dpp.referencia))=substring(a.referencia,6,Len(a.referencia)) " & _
		                        "and dpp.referencia like '"& session("ncliente") &"%' and dpp.npedido='"& ndocumento &"' " & _
		                    "left join "& BDCentral &"..precios as p with(nolock) on a.referencia=p.referencia and p.tarifa='"& tarifaCliente &"' and rango='"& empresa_sup &"BASE' and temporada='"& empresa_sup &"BASE' "
            if rpc & "">"" then
                strselect=strselect & " where dpp.cantidadpend<>0" 
            end if
            sql = "insert into egesticet.[" & session("usuario") & "] " & strselect 
        else
            strselect = "select dpp.item,dpp.referencia,a.importe as pvpdto "
            strselect=strselect & "from "& BDFranq &"..detalles_ped_pro as dpp with(nolock) "
	        strselect=strselect & "inner join "& BDCentral &"..articulos as a with(nolock) on a.referencia like '"& empresa_sup &"%' "
		    strselect=strselect & "and substring(dpp.referencia,6,Len(dpp.referencia))=substring(a.referencia,6,Len(a.referencia)) "
		    strselect=strselect & "and dpp.referencia like '"& session("ncliente") &"%' and dpp.npedido='"& ndocumento &"' "
		    strselect=strselect & " where dpp.cantidadpend<>0" 
            sql = "insert into egesticet.[" & session("usuario") & "] " & strselect 
        end if

        rstTMP.open sql,cnn,adUseClient,adLockReadOnly
	end if
	
	'Recorremos detalles del pedido de compra
	while not rstAux.eof
        if tarifaCliente&"">"" or rpc & "">"" then
	            'Comprobamos que el precio en central=franquicia
	            'precioCentral=d_lookup("pvpdto","egesticet.[" & session("usuario") & "]","item=" & rstAux("item") & "",cnn )
                precioCentralSelect = "select pvpdto from egesticet.[" & session("usuario") & "] with(nolock) where item=? "
                precioCentral=DLookupP1(precioCentralSelect, rstAux("item") &"", adSmallInt, 2, cnn)
	            if precioCentral&"">"" then
	                if (precioCentral<>rstAux("pvp")) then
	                    rstAux("pvp")=precioCentral
	                    rstAux("importe")=miround(precioCentral*rstAux("cantidad"),DEC_PREC)
	                end if
	            end if
        end if
	    rstAux.update
	    if err.number<>0 then
            ha_habido_un_error=1
        end if
	    rstAux.movenext
    wend
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
    'rstAux.close
    ComprobarPreciosEnCentral="OK"
    
end function

'***************************************************************************************************************
'Funcion para pasar los detalles a un pedido de cliente
'***************************************************************************************************************

Function PasarDetallesPedido(ndocumento,npedido,dsnCentral,ncliente)
	'DETALLES DE ARTICULOS
	'almacen=d_lookup("almacen","configuracion","nempresa='" & empresa_sup & "'",dsnCentral)
    almacenSelect= "select almacen from configuracion with(nolock) where nempresa=?"
    almacen=DLookupP1(almacenSelect, empresa_sup &"", adChar, 5, dsnCentral)
    'rstAux.cursorlocation=3
	'rstAux.open "select * from detalles_ped_pro with(nolock) where npedido='" & ndocumento & "' order by item",con_emisor

    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open con_emisor
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "select * from detalles_ped_pro with(nolock) where npedido= ?"
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20, ndocumento)

    set rstAux = commandAux.Execute
      
    'rstAux2.cursorlocation=2
	'rstAux2.open "select * from detalles_ped_cli where npedido='" & npedido & "'",dsnCentral,adOpenKeyset,adLockOptimistic
	
    set commandAux2 = nothing
    set connAux2 = Server.CreateObject("ADODB.Connection")
    set commandAux2 =  Server.CreateObject("ADODB.Command")

    connAux2.open dsnCentral
    connAux2.cursorlocation=2
    commandAux2.ActiveConnection =connAux2
    commandAux2.CommandTimeout = 60
    commandAux2.CommandText= "select * from detalles_ped_cli where npedido= ?"
    commandAux2.CommandType = adCmdText
    commandAux2.Parameters.Append commandAux2.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

    rstAux2.Open commandAux2, , adOpenKeyset, adLockOptimistic
	
    while not rstAux.eof
	    precioCentral = 0
		rstAux2.addnew
		rstAux2("npedido")=npedido
		rstAux2("referencia")=empresa_sup & trimcodempresa(rstAux("referencia"))
		rstAux2("almacen")=almacen
		'coste = d_lookup("importe","articulos","referencia='" & empresa_sup & trimcodempresa(rstAux("referencia")) & "' and referencia like '"& empresa_sup &"%' ",dsnCentral)
	    costeSelect = "select importe from articulos with(nolock) where referencia= ? and referencia like ?+'%'"
        coste = DLookupP2(costeSelect, empresa_sup & trimcodempresa(rstAux("referencia")) & "", adVarchar, 30, empresa_sup &"", adVarchar, 30, dsnCentral)

        rstAux2("coste")=coste
		rstAux2("item")=rstAux("item")
		rstAux2("cantidad")=rstAux("cantidad")
		rstAux2("pvp")=rstAux("pvp")
		rstAux2("importe")=rstAux("importe")
		rstAux2("descripcion")=rstAux("descripcion")
		rstAux2("descuento")=rstAux("descuento")
		rstAux2("iva")=rstAux("iva")
		rstAux2("re")=rstAux("re")
		rstAux2.update
		if err.number<>0 then
            ha_habido_un_error=1
        end if
		rstAux.movenext
	wend
    connAux2.Close
    set connAux2 = nothing
    set commandAux2 = nothing
	'rstAux2.close
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	'rstAux.close

	'DETALLES DE CONCEPTOS
    'rstAux.cursorlocation=3
	'rstAux.Open "select * from conceptos_ped_pro with(nolock) where npedido='" & ndocumento & "' order by nconcepto",session("dsn_cliente")
	
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open session("dsn_cliente")
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "select * from conceptos_ped_pro with(nolock) where npedido= ? order by nconcepto"
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

    set rstAux = commandAux.Execute

    'rstAux2.cursorlocation=2
	'rstAux2.open "select * from conceptos_ped_cli where npedido='" & npedido & "'",dsnCentral,adOpenKeyset,adLockOptimistic

    set commandAux2 = nothing
    set connAux2 = Server.CreateObject("ADODB.Connection")
    set commandAux2 =  Server.CreateObject("ADODB.Command")

    connAux2.open dsnCentral
    connAux2.cursorlocation=2
    commandAux2.ActiveConnection =connAux2
    commandAux2.CommandTimeout = 60
    commandAux2.CommandText= "select * from conceptos_ped_cli where npedido= ?"
    commandAux2.CommandType = adCmdText
    commandAux2.Parameters.Append commandAux2.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

    rstAux2.Open commandAux2, , adOpenKeyset, adLockOptimistic

	while not rstAux.eof
		rstAux2.addnew
		rstAux2("nconcepto")=rstAux("nconcepto")
		rstAux2("npedido")=npedido
		rstAux2("descripcion")=rstAux("descripcion")
		rstAux2("cantidad")=rstAux("cantidad")
		rstAux2("pvp")=rstAux("pvp")
		rstAux2("importe")=rstAux("importe")
		rstAux2("descuento")=rstAux("descuento")
		rstAux2("iva")=rstAux("iva")
		rstAux2("re")=rstAux("re")
		rstAux2.update
		if err.number<>0 then
            ha_habido_un_error=1
        end if
		rstAux.movenext
	wend
    connAux2.Close
    set connAux2 = nothing
    set commandAux2 = nothing
	'rstAux2.close
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	'rstAux.close

	'Actualizacion de stocks
    'rstAux2.cursorlocation=3
	'rstAux2.Open "select referencia,almacen,cantidad from detalles_ped_cli with(nolock) where npedido='" & npedido & "' order by item",con_emisor

    set commandAux2 = nothing
    set connAux2 = Server.CreateObject("ADODB.Connection")
    set commandAux2 =  Server.CreateObject("ADODB.Command")

    connAux2.open con_emisor
    connAux2.cursorlocation=3
    commandAux2.ActiveConnection =connAux2
    commandAux2.CommandTimeout = 60
    commandAux2.CommandText= "select referencia,almacen,cantidad from detalles_ped_cli with(nolock) where npedido=? order by item"
    commandAux2.CommandType = adCmdText
    commandAux2.Parameters.Append commandAux2.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

    set rstAux2 = commandAux2.Execute
    
    while not rstAux2.EOF
        'rstAux.cursorlocation=2
		'rstAux.Open "select * from almacenar where articulo='" & rstAux2("referencia") & "' and almacen='" & rstAux2("almacen") & "'",con_emisor,adOpenKeyset,adLockOptimistic
		
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open con_emisor
        connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select * from almacenar where articulo=? and almacen=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@articulo", adVarChar, adParamInput, 30,  rstAux2("referencia"))
        commandAux.Parameters.Append commandAux.CreateParameter("@almacen", adVarChar, adParamInput, 10, rstAux2("almacen"))

        rstAux.Open commandAux, , adOpenKeyset, adLockOptimistic
        
        if not rstAux.eof then
			ActualizaStocks "first_save","PEDIDO DE CLIENTE",rstAux2("referencia"),rstAux2("almacen"),rstAux2("cantidad"),"",con_emisor
			connAux.Close
            set connAux = nothing
            set commandAux = nothing
            'rstAux.Close
		else
			rstAux.addnew
			rstAux("articulo")=rstAux2("referencia")
			rstAux("almacen")=almacen
			rstAux("stock")=0
			rstAux("stock_minimo")=0
			rstAux("reposicion")=0
			rstAux("p_recibir")=0
			rstAux("p_servir")=rstAux2("cantidad")
			rstAux("p_min")=0
			rstAux("predet")=1
			rstAux.Update
			if err.number<>0 then
                ha_habido_un_error=1
            end if
			connAux.Close
            set connAux = nothing
            set commandAux = nothing
			'rstAux.Close
		end if
		rstAux2.movenext
	wend
    connAux2.Close
    set connAux2 = nothing
    set commandAux2 = nothing
	'rstAux2.close
	PasarDetallesPedido="OK"
end function

'******************************************************************************
'Crea la tabla que contiene la barra de grupos de datos.
sub BarraNavegacion(modo)
	if modo="add" or mode="edit" then%>
        <script language="javascript" type="text/javascript">
            jQuery("#CABECERA").show();
        </script>	
    <%else%>
        <script language="javascript" type="text/javascript">
            jQuery("#CABECERA").hide();
        </script>	
	<%end if%>

    <%if modo="add" then%>
        <script language="javascript" type="text/javascript">
            jQuery("#DIRENVIO").hide();
            jQuery("#FINANCIAL_DATA").hide();
            jQuery("#li-payments").hide();
            jQuery("#li-vencimientos").hide();
        </script>
    <%end if

    if mode="edit" then%>
        <script language="javascript" type="text/javascript">
            jQuery("#DIRENVIO").hide();
            jQuery("#li-payments").hide();
            jQuery("#li-vencimientos").hide();
        </script>
        <%if oculta=1 then%>
            <script language="javascript" type="text/javascript">
                jQuery("#FINANCIAL_DATA").hide();
            </script>   
	    <%else%>
            <script language="javascript" type="text/javascript">
                jQuery("#FINANCIAL_DATA").show();
            </script>   
		<%end if
    end if
    if mode="browse" then
        if oculta=1 then%>
            <script language="javascript" type="text/javascript">
                jQuery("#FINANCIAL_DATA").hide();
                jQuery("#GENERAL_DATA").hide();
            </script>   
	    <%else%>
            <script language="javascript" type="text/javascript">
                jQuery("#FINANCIAL_DATA").show();
                jQuery("#GENERAL_DATA").hide();
            </script>   
		<%end if
        if cstr(cpc)="0" then%>
            <script language="javascript" type="text/javascript">
                jQuery("#li-payments").hide();
            </script>
        <%else%>
            <script language="javascript" type="text/javascript">
                jQuery("#li-payments").show();
            </script>
        <%end if%>
			<script type="text/javascript" language="javascript">
                jQuery(window).load(function () {
                    Redimensionar();
                    try {
                        if (document.getElementById("frDetallesIns").style.display != "none") {
                            fr_DetallesIns.document.pedidos_prodetins.cantidad.focus();
                            fr_DetallesIns.document.pedidos_prodetins.cantidad.select();
                        }
                    }
                    catch (e) {
                    }
                });
			</script>
    <%end if
end sub

'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(npedido,nserie)
	p_nproveedor=session("ncliente") & Completar(request.form("nproveedor"),5,"0")

	if npedido="" then
		'Crear un nuevo registro.
		rst.AddNew

		'******************** Manejo de domicilios
		Dom=Domicilios("COMPRAS","PED_ENV_PROV",p_nproveedor,rst)
		if Dom="FALSE" then
			rst.cancel%>
			<script language="javascript" type="text/javascript">
                window.alert("<%=LitMsgDirPrincipalError%>");
                document.location = "pedidos_pro.asp?mode=add"
                parent.botones.document.location = "pedidos_pro_bt.asp?mode=add"
			</script>
		<%end if
	end if
	FechaDoc=rst("fecha")
	ProvDoc=rst("nproveedor")
	DtoGeneral=null_z(rst("descuento"))
	DtoGeneral2=null_z(rst("descuento2"))
	'Asignar los nuevos valores a los campos del recordset.
	rst("serie")		= Nulear(Request.Form("serie"))
	cambio_proveedor = false
	if rst("nproveedor")<>p_nproveedor then cambio_proveedor=true
	rst("nproveedor")	= Nulear(p_nproveedor)
	rst("fecha")		= Nulear(Request.Form("fecha"))
	'rst("validar")		= nz_b(Request.Form("validar"))
	rst("forma_pago")	= Nulear(request.form("formas_pago"))
	'ndec=d_lookup("ndecimales", "divisas", "codigo like '"&session("ncliente")&"%' and codigo = '"&request.form("h_divisa")&"'", session("dsn_cliente"))
    ndecSelect="select ndecimales from divisas with(nolock) where codigo like ?+'%' and codigo = ?"
    ndec=DLookupP2(ndecSelect, session("ncliente")&"", adVarchar, 15, request.form("h_divisa")&"", adVarchar, 15, session("dsn_cliente"))

	rst("descuento")	= miround(null_z(request.form("dto")),decpor)
	rst("descuento2")	= miround(null_z(request.form("dto2")),decpor)
	rst("recargo")		= miround(null_z(request.form("recargo")),decpor)
	rst("irpf")	= miround(null_z(request.form("irpf")),decpor)
	rst("IRPF_Total")	= nz_b(Request.Form("IRPF_Total"))
	rst("observaciones")= Nulear(request.form("observaciones"))
	rst("cod_proyecto")=Nulear(request.form("cod_proyecto"))
	rst("fecha_entrega")=Nulear(request.form("fecha_entrega"))
	rst("ncuenta")=Nulear(request.form("ncuentacargo"))
	rst("incoterms")=nulear(request.form("incoterms"))
	rst("fob")=nulear(request.form("fob"))

    rst("portes")=Nulear(request.form("portes"))

	rst("notas")= Nulear(request.form("notas"))
	rst("valorado")=nz_b(Request.Form("valorado"))
	rst("facturado")=0
	rst("ahora")=0
	rst("tipo_pago")=Nulear(request.form("tipo_pago"))
	rst("salida")=Nulear(request.form("salida"))
	
	'FLM:130309: cargamos el banco, la cuenta de cargo y la cuenta de abono del proveedor SIEMPRE
	'ncuenta=d_lookup("ncuenta","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))
    ncuentaSelect = "select ncuenta from proveedores with(nolock) where nproveedor= ?"
    ncuenta= DLookupP1(ncuentaSelect, rst("nproveedor") & "", adChar, 10, session("dsn_cliente"))
	if(ncuenta&""<>"") then 
        'banco=d_lookup("Entidad","bancos","codigo='" & mid(trim(ncuenta),5,4) & "'",DsnIlion)
        bancoSelect = "select Entidad from bancos with(nolock) where codigo= ?"
        banco=DLookupP1(bancoSelect,  mid(trim(ncuenta),5,4) &"", adVarchar, 4, DsnIlion)
	else
	    banco=null
	end if
	rst("banco")=iif(banco="",NULL,trim(banco))
	rst("ncuenta_pro")=iif(ncuenta="",NULL,ncuenta)


''ricardo 28/4/2003 si el usuario ha querido recalcular los importes al cambiar las propiedades de la cabecera
han_cambiado_importes_proveedor=0
if request.querystring("recalcular_importes")="1" then
	'Detectamos un cambio de proveedor
	if cambio_proveedor=true then
		han_cambiado_importes_proveedor=1
		set rstMiProveer = Server.CreateObject("ADODB.Recordset")
	    'recorremos los detalles modificando precios
        'rstAux.cursorlocation=2
		'rstaux.open "select * from detalles_ped_pro with(updlock) where npedido='" & npedido & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select * from detalles_ped_pro with(updlock) where npedido= ? order by item"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20,  npedido)

        rstAux.Open commandAux, , adOpenKeyset, adLockOptimistic

        while not rstAux.eof
            'rstMiProveer.cursorlocation=3
			'rstMiProveer.open "select * from proveer with(nolock) where nproveedor='" & rst("nproveedor") & "' and articulo='" & rstaux("referencia") & "'",session("dsn_cliente")
			
            set commandMiProveer = nothing
            set connMiProveer = Server.CreateObject("ADODB.Connection")
            set commandMiProveer =  Server.CreateObject("ADODB.Command")

            connMiProveer.open session("dsn_cliente")
            connMiProveer.cursorlocation=3
            commandMiProveer.ActiveConnection =connMiProveer
            commandMiProveer.CommandTimeout = 60
            commandMiProveer.CommandText= "select * from proveer with(nolock) where nproveedor= ? and articulo= ?"
            commandMiProveer.CommandType = adCmdText
            commandMiProveer.Parameters.Append commandMiProveer.CreateParameter("@nproveedor", adChar, adParamInput, 10,  rst("nproveedor"))
            commandMiProveer.Parameters.Append commandMiProveer.CreateParameter("@articulo", adVarChar, adParamInput, 30,  rstaux("referencia"))

            set rstMiProveer = commandMiProveer.Execute
            
            if not rstMiProveer.eof then
				TmpPvp=CambioDivisa(rstMiProveer("importe"), rstMiProveer("divisa"),request.form("h_divisa"))
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
            connMiProveer.Close
            set connMiProveer = nothing
            set commandMiProveer = nothing
			'rstMiProveer.close
		wend
		rstAux.close
		Set rstMiProveer=nothing

		'recorremos ahora los conceptos haciendo cambio de divisa si hace falta
        'rstaux.cursorlocation=2
		'rstaux.open "select * from conceptos_ped_pro where npedido='" & npedido & "' order by nconcepto",session("dsn_cliente")

        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select * from conceptos_ped_pro where npedido=? order by nconcepto"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

        set rstAux = commandAux.Execute

		while not rstAux.eof
			TmpPVP=CambioDivisa(rstAux("pvp"),rst("divisa"),request.form("h_divisa"))
			rstAux("pvp")=TmpPVP
			TmpPVP=TmpPVP*rstAux("cantidad")
			TmpPVP	= TmpPVP*(100-null_z(rstAux("descuento")))/100
			'TmpPVP	= TmpPVP*(100-null_z(rst("descuento2")))/100
			rstAux("importe")=miround(TmpPVP,2)
			rstAux.update
			rstAux.movenext
		wend
                
		rstAux.close
	end if
end if

''ricardo 28/4/2003 si el usuario ha querido recalcular los importes al cambiar las propiedades de la cabecera
if request.querystring("recalcular_importes")="1" then
	'Detectar cambios en la divisa del documento para cambiar la divisa de los detalles(artículos y conceptos)
	'if rst("divisa")<>request.form("h_divisa") and rst("divisa")&"">"" then
	'	rstAux.open "select * from detalles_ped_pro where npedido='" & npedido & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	'	while not rstAux.eof
	'		TmpPVP=CambioDivisa(rstAux("pvp"),rst("divisa"),request.form("h_divisa"))
	'		rstAux("pvp")=TmpPVP
	'		TmpPVP=TmpPVP*rstAux("cantidad")
	'		TmpPVP	= TmpPVP*(100-null_z(rst("descuento")))/100
	'		TmpPVP	= TmpPVP*(100-null_z(rst("descuento2")))/100
	'		rstAux("importe")=miround(TmpPVP,2)
	'		rstAux.update
	'		rstAux.movenext
	'	wend
	'	rstAux.close
	'	rstaux.open "select * from conceptos_ped_pro where npedido='" & npedido & "' order by nconcepto",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	'	while not rstAux.eof
	'		TmpPVP=CambioDivisa(rstAux("pvp"),rst("divisa"),request.form("h_divisa"))
	'		rstAux("pvp")=TmpPVP
	'		TmpPVP=TmpPVP*rstAux("cantidad")
	'		TmpPVP	= TmpPVP*(100-null_z(rst("descuento")))/100
	'		TmpPVP	= TmpPVP*(100-null_z(rst("descuento2")))/100
	'		rstAux("importe")=miround(TmpPVP,2)
	'		rstAux.update
	'		rstAux.movenext
	'	wend
	'	rstAux.close
	'end if
end if

	rst("divisa")= Nulear(request.form("h_divisa"))
	'*** AMP 21102010
    rst("factcambio")=miround(Nulear(limpiaCadena(request.form("nfactcambio"))),DEC_PREC) 
	'Actualizar el registro.
	if npedido="" then
		SigDoc=CalcularNumDocumento(nserie,request.form("fecha"))
		rst("npedido")=SigDoc
		npedido=SigDoc
	end if
	ref_edi=CalculaEDI(nserie,Nulear(p_nproveedor),rst("npedido"),"proveedor")
	if ref_edi>"" then
			rst("EDI")=ref_edi
	else rst("edi")=NULL
	end if

	'' JMA 16/12/04 Actualizamos los campos personalizables
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
                'tipo_campo_perso=d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & "0" & ki & "' and tabla='DOCUMENTOS COMPRA'",session("dsn_cliente"))
			    tipo_campo_perso_SELECT = "select tipo from camposperso with(nolock) where ncampo = ? and tabla='DOCUMENTOS COMPRA'"
                tipo_campo_perso = DLookupP1(tipo_campo_perso_SELECT, session("ncliente") & "0" & ki &"", adChar, 7, session("dsn_cliente"))
    
            else
				'tipo_campo_perso=d_lookup("tipo","camposperso","ncampo='" & session("ncliente") & ki & "' and tabla='DOCUMENTOS COMPRA'",session("dsn_cliente"))
			    tipo_campo_perso_SELECT = "select tipo from camposperso with(nolock) where ncampo = ? and tabla='DOCUMENTOS COMPRA'"
                tipo_campo_perso = DLookupP1(tipo_campo_perso_SELECT, session("ncliente") & ki & "", adChar, 7, session("dsn_cliente"))

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
	'' JMA 28/10/04 Fin actualizar campos personalizables

	rst.update
	rst.close

		if han_cambiado_importes_proveedor=1 then
			'TmpIvaProveedor=d_lookup("iva","proveedores","nproveedor='" & p_nproveedor & "'",session("dsn_cliente"))
            TmpIvaProveedorSelect = "select iva from proveedores with(nolock) where nproveedor= ?"
            TmpIvaProveedor=DLookupP1(TmpIvaProveedorSelect, p_nproveedor & "", adChar, 10, session("dsn_cliente"))

			'TmpReProveedor=d_lookup("re","tipos_iva","tipo_iva='" & TmpIvaProveedor & "'",session("dsn_cliente"))
            TmpReProveedorSelect = "select re from tipos_iva with(nolock) where tipo_iva= ?"
            TmpReProveedor=DLookupP1(TmpReProveedorSelect, TmpIvaProveedor &"", 139, 4, session("dsn_cliente"))

			'defaultIVA=d_lookup("iva","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
            defaultIVASelect = "select iva from configuracion with(nolock) where nempresa= ?"
            defaultIVA = DLookupP1(defaultIVASelect, session("ncliente") &"", adChar, 5, session("dsn_cliente"))

			'TmpReDefaultIva=d_lookup("re","tipos_iva","tipo_iva='" & defaultIVA & "'",session("dsn_cliente"))
            TmpReDefaultIvaSelect = "select re from tipos_iva with(nolock) where tipo_iva= ?"
            TmpReDefaultIva= DLookupP1(TmpReDefaultIvaSelect, defaultIVA &"", 139, 4, session("dsn_cliente"))
			if TmpIvaProveedor & "">"" then
				TmpIva=TmpIvaProveedor
				TmpRe=TmpReProveedor
			else
				TmpIva=defaultIVA
				TmpRe=TmpReDefaultIva
			end if
            'rstaux.cursorlocation=2
			'rstaux.open "select d.iva,d.re,d.referencia from detalles_ped_pro as d where d.npedido='" & npedido & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			
            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open session("dsn_cliente")
            connAux.cursorlocation=2
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText= "select d.iva,d.re,d.referencia from detalles_ped_pro as d where d.npedido= ? order by item"
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

            rstAux.Open commandAux, , adOpenKeyset, adLockOptimistic

            while not rstAux.eof
				if TmpIva & "">"" then
					rstAux("iva")=TmpIva
					rstAux("re")=TmpRe
				else
					'Tmpivaart=d_lookup("iva","articulos","referencia='" & rstAux("referencia") & "'",session("dsn_cliente"))
                    TmpivaartSelect = "select iva from articulos with(nolock) where referencia= ?"
                    Tmpivaart = DLookupP1(TmpivaartSelect, rstAux("referencia") &"", adVarchar, 30, session("dsn_cliente"))
					
                    'TmpReivaart=d_lookup("re","tipos_iva","tipo_iva='" & Tmpivaart & "'",session("dsn_cliente"))
                    TmpReivaartSelect = "select re from tipos_iva with(nolock) where tipo_iva= ?"
                    TmpReivaart = DLookupP1(TmpReivaartSelect,  Tmpivaart &"", 139, 4, session("dsn_cliente"))
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
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
			'rstAux.close
            'rstaux.cursorlocation=2
			'rstaux.open "select * from conceptos_ped_pro where npedido='" & npedido & "' order by nconcepto",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			
            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open session("dsn_cliente")
            connAux.cursorlocation=2
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText= "select * from conceptos_ped_pro where npedido= ? order by nconcepto"
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

            rstAux.Open commandAux, , adOpenKeyset, adLockOptimistic

            while not rstAux.eof
				rstAux("iva")=TmpIva
				rstAux("re")=TmpRe
				rstAux.update
				rstAux.movenext
			wend
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
			'rstAux.close
		end if
        'rst.cursorlocation=2
		'rst.open "select * from pedidos_pro where npedido='" & npedido & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        
        set commandRst = nothing
        set connRst = Server.CreateObject("ADODB.Connection")
        set commandRst =  Server.CreateObject("ADODB.Command")

        connRst.open session("dsn_cliente")
        connRst.cursorlocation=2
        commandRst.ActiveConnection =connRst
        commandRst.CommandTimeout = 60
        commandRst.CommandText= "select * from pedidos_pro where npedido= ?"
        commandRst.CommandType = adCmdText
        commandRst.Parameters.Append commandRst.CreateParameter("@npedido", adVarChar, adParamInput, 20, npedido)

        rst.Open commandRst, , adOpenKeyset, adLockOptimistic
        
	'Precios del pedido.
	'Miramos si el proveedor tiene recargo de equivalencia
    'rstAux.cursorlocation=3
	'rstAux.open "Select re from proveedores with(nolock) where nproveedor='" + rst("nproveedor") + "'",session("dsn_cliente")

    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open session("dsn_cliente")
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "Select re from proveedores with(nolock) where nproveedor= ?"
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@nproveedor", adChar, adParamInput, 10, rst("nproveedor"))

    set rstAux = commandAux.Execute

	if not rstAux.eof then
		TieneRE=rstAux("re")
	else
		TieneRE=0
	end if
	rstAux.close

	'n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))
    n_decimales_select = "select ndecimales from divisas with(nolock) where codigo= ?"
    n_decimales= DLookupP1(n_decimales_select, rst("divisa") & "", adVarchar, 15, session("dsn_cliente"))

	if n_decimales = "" then
		n_decimales = 0
	end if

	rst("importe_bruto")	= 0
	rst("base_imponible")	= 0
	rst("total_descuento")	= 0
	rst("total_iva")		= 0
	rst("total_re")		= 0
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

	seleccion="select sum(importe) as suma, iva, re from detalles_ped_pro with(nolock) "
	seleccion=seleccion+"where npedido =? and mainitem is null "
	seleccion=seleccion+"GROUP BY IVA, RE "
	seleccion=seleccion+" union all "
	seleccion=seleccion+"select sum(importe) as suma, iva, re from conceptos_ped_pro with(nolock) "
	seleccion=seleccion+"where npedido =? "
	seleccion=seleccion+"GROUP BY IVA, RE ORDER BY IVA"
    'rstAux.cursorlocation=3
	'rstAux.open seleccion,session("dsn_cliente")

    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open session("dsn_cliente")
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= seleccion
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20, rst("npedido")&"")
    commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar, adParamInput, 20, rst("npedido")&"")

    set rstAux = commandAux.Execute

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
            're=d_lookup("re","tipos_iva","tipo_iva='" & rstAux("iva") & "'",session("dsn_cliente"))
            reSelect = "select re from tipos_iva with(nolock) where tipo_iva= ?"
            re = DLookupP1(reSelect, rstAux("iva") &"", 139, 4, session("dsn_cliente"))
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
	rst("total_pedido")=miround(SumTotalImporte,2)

	rst.Update
	ActualizaCostes iif(npedido="",SigDoc,npedido),"DOCUMENTO","PEDIDO A PROVEEDOR","",ProvDoc,0,FechaDoc,0,0,DtoGeneral,DtoGeneral2,false,session("dsn_cliente")
end sub

'Elimina los datos del registro cuando se pulsa BORRAR.
function BorrarRegistro(npedido)
	tiene_caja=0
    'rst.cursorlocation=3
	'rst.open "select ndocumento from caja with(nolock) where ndocumento='" & npedido & "'",session("dsn_cliente")

    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open session("dsn_cliente")
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "select ndocumento from caja with(nolock) where ndocumento=?"
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@ndocumento", adVarChar, adParamInput, 22, npedido)

    set rst = commandAux.Execute

	if not rst.eof then
		tiene_caja=1
	end if
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	'rst.close

    'rst.cursorlocation=3
	'rst.open "select * from pedidos_pro where npedido='" & npedido & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

    set commandoRst = nothing
    set connoRst = Server.CreateObject("ADODB.Connection")
    set commandoRst =  Server.CreateObject("ADODB.Command")

    connoRst.open session("dsn_cliente")
    connoRst.cursorlocation=3
    commandoRst.ActiveConnection =connoRst
    commandoRst.CommandTimeout = 60
    commandoRst.CommandText= "select * from pedidos_pro where npedido= ?"
    commandoRst.CommandType = adCmdText
    commandoRst.Parameters.Append commandoRst.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

    rst.Open commandoRst, , adOpenKeyset, adLockOptimistic

	'Se generó factura o albarán. Imposible eliminar
	if rst("nfactura")>"" or rst("nalbaran")>"" or tiene_caja=1 then
		if rst("nfactura")>"" then%>
			<script language="javascript" type="text/javascript">
                //window.alert("<%=LitMsgBorrarPedido%> <%=d_lookup("nfactura_pro","facturas_pro","nfactura= '"&rst("nfactura")&"'",session("dsn_cliente"))%>");
			    <%
                    alertSelect = "select nfactura_pro from facturas_pro with(nolock) where nfactura= ?"
                    alertMessage = DLookupP1(alertSelect, rst("nfactura")&"", adVarchar, 20, session("dsn_cliente"))
                %>
                window.alert("<%=LitMsgBorrarPedido & " " & enc.EncodeForJavascript(alertMessage)%>");
			</script>
			<%BorrarRegistro=false
		elseif rst("nalbaran")>"" then%>
			<script language="javascript" type="text/javascript">
                //window.alert("<%=LitMsgBorrarPedido2%> <%=d_lookup("nalbaran_pro","albaranes_pro","nalbaran= '"&rst("nalbaran")&"'",session("dsn_cliente"))%>");
                <%
                    alertSelect = "select nalbaran_pro from albaranes_pro with(nolock) where nalbaran= ?"
                    alertMessage = DLookupP1(alertSelect,  rst("nalbaran")&"", adVarchar, 20, session("dsn_cliente"))
                %>
                window.alert("<%=LitMsgBorrarPedido2 & " " & enc.EncodeForJavascript(alertMessage)%>");
            </script>
			<%BorrarRegistro=false
		elseif tiene_caja=1 then%>
			<script language="javascript" type="text/javascript">
                window.alert("<%=LitMsgBorrarCaja%>");
			</script>
		<%end if
		BorrarRegistro=false
	'Antes de borrar se modifica el stock pendiente de recibir
	else
		'Miramos si se va a borrar el último generado y si es así se descuenta el contador de documentos
		'FechaPedido=d_lookup("fecha","pedidos_pro","npedido='" & npedido & "'",session("dsn_cliente"))
        fechaPedidoSelect = "select fecha from pedidos_pro with(nolock) where npedido= ?"
        FechaPedido= DLookupP1(fechaPedidoSelect, npedido & "",  adVarchar, 20, session("dsn_cliente"))
		ano=right(cstr(year(FechaPedido)),2)
		'nserie=d_lookup("serie","pedidos_pro","npedido='" & npedido & "'",session("dsn_cliente"))
        nserieSelect = "select serie from pedidos_pro with(nolock) where npedido= ?"
        nserie= DLookupP1(nserieSelect, npedido & "",  adVarchar, 20, session("dsn_cliente"))
        'rstAux.cursorlocation=2
		'rstAux.Open "select * from series with(updlock) where nserie='" & nserie & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		
        set commandRstAux = nothing
        set connRstAux = Server.CreateObject("ADODB.Connection")
        set commandRstAux =  Server.CreateObject("ADODB.Command")

        connRstAux.open session("dsn_cliente")
        connRstAux.cursorlocation=2
        commandRstAux.ActiveConnection =connRstAux
        commandRstAux.CommandTimeout = 60
        commandRstAux.CommandText= "select * from series with(updlock) where nserie= ?"
        commandRstAux.CommandType = adCmdText
        commandRstAux.Parameters.Append commandRstAux.CreateParameter("@nserie", adVarChar,adParamInput,10, nserie)

        rstAux.Open commandRstAux, , adOpenKeyset, adLockOptimistic
            
        ultimo=rstAux("contador")

		'mmg:calculamos el almacen por defecto de la serie 
        if rstAux.eof then
	        almacenSerie= ""
        else
            'comprobamos si el almacen esta dado de baja
            'rstMM.cursorlocation=3
            'rstMM.Open "select codigo from almacenes where codigo='" & rstAux("almacen") & "' and isnull(fbaja,'')=''",session("dsn_cliente")
		    
            set commandRstMM = nothing
            set connRstMM = Server.CreateObject("ADODB.Connection")
            set commandRstMM =  Server.CreateObject("ADODB.Command")

            connRstMM.open session("dsn_cliente")
            connRstMM.cursorlocation=3
            commandRstMM.ActiveConnection =connRstMM
            commandRstMM.CommandTimeout = 60
            commandRstMM.CommandText= "select codigo from almacenes where codigo=? and isnull(fbaja,'')=''"
            commandRstMM.CommandType = adCmdText
            commandRstMM.Parameters.Append commandRstMM.CreateParameter("@codigo", adVarChar,adParamInput,10, rstAux("almacen"))

            set rstMM = commandRstMM.Execute
            
            if rstMM.eof then
	            almacenSerie= ""
            else
	            almacenSerie= rstAux("almacen")
	        end if
        end if
            
		if npedido=nserie+ano+completar(trim(cstr(ultimo)),5,"0") then
			rstAux("contador")=ultimo-1
			rstAux.update
		end if
        connRstAux.Close
        set connRstAux = nothing
        set commandRstAux = nothing
		'rstAux.close
        'rstAux.cursorlocation=2
        'rstAux.Open "delete from VENCIMIENTOS_pedPRO with(rowlock) where npedido='" & npedido & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic

        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "delete from VENCIMIENTOS_pedPRO with(rowlock) where npedido=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing


		'rstAux.Open "delete from conceptos_ped_pro with(rowlock) where npedido='" & npedido & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "delete from conceptos_ped_pro with(rowlock) where npedido=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
            
        'rstAux.Open "select referencia,almacen,cantidad,npedidocli from detalles_ped_pro where npedido='" & npedido & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select referencia,almacen,cantidad,npedidocli from detalles_ped_pro where npedido=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        rstAux.Open commandAux, , adOpenKeyset, adLockOptimistic

        'Borrado y Actualización de stocks
		while not rstAux.EOF
		'Si existen enlaces con pedidos de cliente los eliminamos
			if not isnull(rstAux("npedidocli")) then
				'rstAccion.open "update detalles_ped_cli with(updlock) set npedidopro=null, itempedidopro=0 where npedidopro='"&npedido&"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                
                set commandAux = nothing
                set connAux = Server.CreateObject("ADODB.Connection")
                set commandAux =  Server.CreateObject("ADODB.Command")

                connAux.open session("dsn_cliente")
                'connAux.cursorlocation=2
                commandAux.ActiveConnection =connAux
                commandAux.CommandTimeout = 60
                commandAux.CommandText= "update detalles_ped_cli with(updlock) set npedidopro=null, itempedidopro=0 where npedidopro=?"
                commandAux.CommandType = adCmdText
                commandAux.Parameters.Append commandAux.CreateParameter("@npedidopro", adVarChar,adParamInput,20, npedido)

                set rstAccion = commandAux.Execute
                connAux.Close
                set connAux = nothing
                set commandAux = nothing

			end if
			refST=rstAux("referencia")
			almST=rstAux("almacen")
			canST=rstAux("cantidad")
			rstAux.delete
			ActualizaCostes npedido,"DELETEDETALLE","PEDIDO A PROVEEDOR",refST,npro_aux,0,FechaPedido,0,0,0,0,false,session("dsn_cliente")
			ActualizaStocks "delete","PEDIDO A PROVEEDOR",refST,almST,canST,"",session("dsn_cliente")
			rstAux.MoveNext
		wend
		rstAux.close
		'borramos los pagos si tienen
		'rstAux.open "delete from pagos_ped_pro with(rowlock) where npedido='" & npedido & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic

        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        'connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "delete from pagos_ped_pro with(rowlock) where npedido=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        set rstAux = commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing

		'Actualizar el campo NPEDIDO_BIS de la tabla ALBARANES_PRO por si era un pedido BIS
		'rstAux.open "update albaranes_pro with(updlock) set npedido_bis=null where npedido_bis='" & npedido & "'", _
		'session("dsn_cliente"),adOpenKeyset,adLockOptimistic

        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        'connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "update albaranes_pro with(updlock) set npedido_bis=null where npedido_bis=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido_bis", adVarChar,adParamInput,20, npedido)

        set rstAux = commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing

		'rstAux.Open "delete from pedidos_pro with(rowlock) where npedido='" & npedido & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        'connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "delete from pedidos_pro with(rowlock) where npedido= ?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
            
        'Actualizar el campo NPEDIDO_BIS de la tabla ALBARANES_PRO por si era un pedido BIS
		'rstAux.open "update albaranes_pro with(updlock) set npedido_bis=null where npedido_bis='" & npedido & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        'connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "update albaranes_pro with(updlock) set npedido_bis=null where npedido_bis=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido_bis", adVarChar,adParamInput,20, npedido)

        commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
            
        BorrarRegistro=true
	end if
	rst.close
end function

'****************************************************************************************************************

'******************************************************************************
'Guardar fisicamente el registro de detalle del pedido
sub GuardaDetalle(npedido,nserie)
	'Obtener el último nº de detalle del pedido.
	num=d_max("item","detalles_ped_cli","npedido='" & npedido & "'",session("dsn_cliente"))+1
	'Crear un nuevo registro.
	rstDetPed.AddNew
	rstDetPed("npedido")=npedido
	rstDetPed("item")=num

	'Asignar los nuevos valores a los campos del recordset.
	rstDetPed("referencia")  =rstPed("referencia")
	rstDetPed("almacen")     =rstPed("almacen")
	if nserie="SI" then
		rstDetPed("descuento")   =rstPed("descuento")
		rstDetPed("pvp")         =rstPed("pvp")
		rstDetPed("cantidad")    =1
		temp=rstDetAlb("cantidad")*null_z(rstDetAlb("pvp"))
		temp               		 =formatnumber(null_z(rstDetAlb("pvp")-((rstDetPed("pvp")*rstDetPed("descuento"))/100)),n_decimales,-1,0,iif(mode="browse",-1,0))
		rstDetPed("importe")     =formatnumber((temp*rstDetAlb("cantidad")),n_decimales,-1,0,iif(mode="browse",-1,0))
		rstDetPed("descripcion") =rstPed("descripcion")
		rstDetPed("iva")         =rstPed("iva")
		rstDetPed("re")          =rstPed("re")
		rstDetPed("npedido")     =rstPed("npedido")
	else
		rstDetPed("cantidad")    =rstPed("cantidad")
		rstDetPed("pvp")         =rstPed("pvp")
		rstDetPed("importe")     =rstPed("importe")
		rstDetPed("descripcion") =rstPed("descripcion")
		rstDetPed("descuento")   =rstPed("descuento")
		rstDetPed("iva")         =rstPed("iva")
		rstDetPed("re")          =rstPed("re")
		rstDetPed("npedido")     =rstPed("npedido")
	end if
	rstDetPed.Update
	'Actualizacion de stocks
	ActualizaStocks "first_save","PEDIDO DE CLIENTE",rstDetPed("referencia"),rstDetPed("almacen"),rstDetPed("cantidad"),"",session("dsn_cliente")
end sub

'********** CODIGO PRINCIPAL DE LA PÁGINA
set connRound = Server.CreateObject("ADODB.Connection")
connRound.open dsnilion
	
    n_decimales = 0%>
<form name="pedidos_pro" method="post">
    <%
    PintarCabecera "pedidos_pro.asp"
    
    ' Ocultar detalle de las facturas si se da el caso
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
	command.Execute,,adExecuteNoRecords
	oculta = Command.Parameters("@resul").Value
	conn.close
	set command=nothing
	set conn=nothing

    Dim rt,cpc,rpc,vencesp,modP
    ObtenerParametros("pedidos_pro_det")
    ''response.write("el vencesp es-" & vencesp & "-<br>")
        'modP = limpiaCadena(Request.QueryString("modP"))

	'Leer parámetros de la página
	mode	= Request.QueryString("mode")
	submode	= Request.QueryString("submode")

	mode2=request.querystring("mode2")

	if request.querystring("cod_proyecto")>"" then
		tmp_cod_proyecto=limpiaCadena(request.querystring("cod_proyecto"))
	else
		tmp_cod_proyecto=limpiaCadena(request.form("cod_proyecto"))
	end if

	if request.querystring("fecha_entrega")>"" then
		tmp_fecha_entrega=limpiaCadena(request.querystring("fecha_entrega"))
	else
		tmp_fecha_entrega=limpiaCadena(request.form("fecha_entrega"))
	end if

	viene=limpiaCadena(request.querystring("viene"))
	if viene="" then viene=limpiaCadena(request.form("viene"))
	if viene="" then viene="pedidos_pro.asp"

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
	
	if request.querystring("salida")>"" then
		tmp_salida=limpiaCadena(request.querystring("salida"))
	else
		tmp_salida=limpiaCadena(request.form("salida"))
	end if
	
	if request.querystring("cambiar_serie")>"" then
		cambiar_serie=limpiaCadena(request.querystring("cambiar_serie"))
	else
		cambiar_serie=request.form("cambiar_serie")
	end if

	if request.querystring("cambiar_cliente")>"" then
		cambiar_cliente=limpiaCadena(request.querystring("cambiar_cliente"))
	else
		cambiar_cliente=request.form("cambiar_cliente")
	end if
	
	'*** AMP 
	if Request.QueryString("divisafc")>"" then
		tmpdivisafc=limpiaCadena(Request.QueryString("divisafc"))
	elseif Request.form("divisafc")>"" then
		tmpdivisafc=limpiaCadena(Request.form("divisafc"))
	end if	

    campo    = limpiaCadena(request.QueryString("campo"))
    if campo & ""="" then
        campo = Request.Form("campo")
    end if
    texto    = limpiaCadena(request.QueryString("texto"))
    if texto & ""="" then
        texto = Request.Form("texto")
    end if
	
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
    if npagina & ""="" then npagina=0
    
    formato_impresionEleg=limpiaCadena(Request.QueryString("formato_impresionEleg"))
    if formato_impresionEleg & ""="" then
        formato_impresionEleg=limpiaCadena(Request.form("formato_impresionEleg"))
    end if

    if request.querystring("portes")>"" then
		tmp_portes=limpiaCadena(request.querystring("portes"))
	else
		tmp_portes=limpiaCadena(request.form("portes"))
	end if

' >>> MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compra.
'					 pedidos_pro.asp  bloque 1/4
	s=limpiaCadena(request.querystring("s"))
	if s="" then s=limpiaCadena(request.form("s"))
	s=preparar_lista(s)

'iframe oculto que utilizaremos para comprobar los límites de compras, sólo nos hace falta en mode edit%>
	<iframe id='comprobar_limites' src='comprobar_limites.asp?viene=pedpro' style="display:none;" width='0' height='0' frameborder="no" scrolling="no" noresize="noresize"></iframe>

    <%' <<< MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compra.%>
	<input type="hidden" name="viene" value="<%=EncodeForHtml(viene)%>"/>
	<input type="hidden" name="mode" value="<%=EncodeForHtml(mode)%>"/>
	<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>"/>
	<input type="hidden" name="novei" value="<%=EncodeForHtml(novei)%>"/>
	<input type="hidden" name="s" value="<%=EncodeForHtml(s)%>"/>
	<input type="hidden" name="convertidoPedCli" value="NO"/>
	<input type="hidden" name="cpc" value="<%=cpc%>"/>
    <input type="hidden" name="campo" value="<%=EncodeForHtml(campo)%>"/>
    <input type="hidden" name="texto" value="<%=EncodeForHtml(texto)%>"/>
    <input type="hidden" name="lote" value="<%=EncodeForHtml(lote)%>"/>
    <input type="hidden" name="criterio" value="<%=EncodeForHtml(criterio)%>"/>
    <input type="hidden" name="rpc" value="<%=EncodeForHtml(rpc)%>"/>
    <input type="hidden" name="formato_impresionEleg" value=""/>
  
    <%set rstAux = Server.CreateObject("ADODB.Recordset")
    set rstMM = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstdomi = Server.CreateObject("ADODB.Recordset")
	set rstAccion = Server.CreateObject("ADODB.Recordset")
	set rstLimite = Server.CreateObject("ADODB.Recordset")
	set rstPedirPro = Server.CreateObject("ADODB.Recordset")
	set rstIvas = Server.CreateObject("ADODB.Recordset")
	set rstCli = Server.CreateObject("ADODB.Recordset")
    set rstTMP = Server.CreateObject("ADODB.Recordset")


    
	if request.querystring("prov")>"" then
		nproveedor=limpiaCadena(null_s(request.querystring("prov")))
	else
		nproveedor=limpiaCadena(null_s(request.form("prov")))
	end if
	if nproveedor="" then
		if request.querystring("nproveedor")>"" then
			nproveedor=limpiaCadena(null_s(request.querystring("nproveedor")))
		else
			nproveedor=limpiaCadena(null_s(request.form("nproveedor")))
		end if
	end if
	tmp_valorado=limpiaCadena(null_s(Request.QueryString("valorado")))
	tmp_fecha=limpiaCadena(null_s(Request.QueryString("fecha")))
	tmp_serie=limpiaCadena(null_s(Request.QueryString("serie")))
	if tmp_serie & ""="" then
		tmp_serie=limpiaCadena(null_s(Request.form("serie")))
	end if
	observacionesR=limpiaCadena(null_s(request.querystring("observaciones")))
	notasR=limpiaCadena(null_s(request.querystring("notas")))
	serieR=limpiaCadena(null_s(Request.Form("serie")))
	fechaR=limpiaCadena(null_s(Request.Form("fecha")))
	npedidoH=limpiaCadena(null_s(Request.Form("h_npedido")))
	
    ''ricardo 3-11-2011 si no tiene acceso a la opcion de almacenes , se quitara dicho campo
    si_tiene_acceso_almacenes=1
    rst.Open "exec ContractedItem '" & session("ncliente") & "','" & replace(OBJAlmacenes,"'","''") & "'", dsnilion
    if not rst.eof then
        if rst(0)=1 then
            si_tiene_acceso_almacenes=1
        else
            si_tiene_acceso_almacenes=0
        end if
    end if
    rst.close

' >>> MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compra.
'					 pedidos_pro.asp	bloque 2/4

	if p_npedido & "">"" then
		if comprobar_LS(s,mode,p_npedido,"PEDIDOS_PRO")=0 then%>
			<script language="javascript" type="text/javascript">
                window.alert("<%=LitMsgDocNoPermAcc%>");
                document.pedidos_pro.action = "pedidos_pro.asp?npedido=&mode=add";
                document.pedidos_pro.submit();
                parent.botones.document.location = "pedidos_pro_bt.asp?mode=add";
			</script>
            <%CerrarTodo()
			response.end
		end if
	end if
    ' <<< MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compra.

	'***RGU 26/5/2006***
	WaitBoxOculto LitEsperePorFavor
	if submode="pedir_pro" then
		rstPedirPro.cursorlocation=3
		rstPedirPro.open " exec PediraProveedor '"&npedido&"','"&session("usuario")&"' ",dsnilion,adOpenKeyset,adLockOptimistic
		if rstPedirPro.state <> 0 then
			if not rstPedirPro.eof then%>
				<script>
				  <%if rstPedirPro(0)="TABLA" then
				  	mode="browse"
				  	%>
				  	AbrirVentana("../central.asp?pag1=compras/pedidos_pro_edierr.asp&mode=browse&ndoc=<%=enc.EncodeForJavascript(npedido)%>&pag2=compras/pedidos_pro_edierr_bt.asp","I",<%=AltoVentana%>,<%=AnchoVentana%>);
				  <%else%>
					window.alert("<%=rstPedirPro(0)%>");
				  <%end if%>
				  </script>
			<%end if
			rstPedirPro.close
			set rstPedirPro=nothing
		else%>
			<script>
				window.alert("<%=LitOK%>");
			  </script>
		<%end if
		'mode="browse"
	end if
	'***RGU***
	
	'JMMM - 05/10/2010 -> Crear pedido de venta en central (franquicias)
	
	if submode="crearPedCentral" then
	    if rpc & "">"" then
	        'empresa_sup=d_lookup("empresa_sup", "configuracion", "nempresa='"& session("ncliente") &"'", session("dsn_cliente"))
            empresa_sup_select = "select empresa_sup from configuracion with(nolock) where nempresa = ?"
            empresa_sup=DLookupP1(empresa_sup_select, session("ncliente") &"", adChar, 5, session("dsn_cliente"))
            if empresa_sup & "">"" then
                'dsnCentral=d_lookup("dsn", "clientes", "ncliente='"&empresa_sup&"'", DsnIlion)
                dsnCentralSelect = "select dsn from clientes with(nolock) where ncliente= ?"
                dsnCentral=DLookupP1(dsnCentralSelect, empresa_sup&"", adChar, 10, DsnIlion)
            else
                dsnCentral=""
            end if

            'nproveedor=d_lookup("nproveedor","pedidos_pro","npedido='" & npedido & "'",session("dsn_cliente"))
            nproveedorSelect = "select nproveedor from pedidos_pro with(nolock) where npedido= ?"
            nproveedor= DLookupP1(nproveedorSelect, npedido & "", adVarchar, 20, session("dsn_cliente"))
            if nproveedor & "">"" then
	            'cif_emisor=d_lookup("cifedi","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente"))
                cif_emisor_select = "select cifedi from proveedores with(nolock) where nproveedor= ?"
                cif_emisor= DLookupP1(cif_emisor_select, nproveedor & "", adChar, 10, session("dsn_cliente")) 
	        else
	            cif_emisor=""
	        end if
            if cif_emisor & "">"" and empresa_sup & "">"" and dsnCentral & "">"" then
	            'ncliente_part=d_lookup("ncliente","clientes","cifedi='" & cif_emisor & "' and ncliente like '" & empresa_sup & "%'",dsnCentral)
	            ncliente_part_select = "select ncliente from clientes with(nolock) where cifedi= ? and ncliente like ?+'%'"
                ncliente_part= DLookupP2(ncliente_part_select, cif_emisor & "", adVarchar, 20, empresa_sup & "", adChar, 10, dsnCentral)
            else
	            ncliente_part=""
	        end if
	        doc=npedido
	    else
            'Recuperamos el EDI del pedido_pro
            'edi=d_lookup("edi","pedidos_pro","npedido='" & npedido & "' and npedido like '" & session("ncliente") & "%'",session("dsn_cliente"))
            edi_select = "select edi from pedidos_pro with(nolock) where npedido= ? and npedido like ?+'%'"
            edi = DLookupP2(edi_select, npedido & "", adVarchar, 20, session("ncliente") & "", adVarchar, 20, session("dsn_cliente"))
            nserie_edi=mid(edi,11,4)
            ncliente_origen_edi=left(edi,5)
		    ncliente_destino_edi=mid(edi,6,5)
		    any_serie_edi=mid(edi,15,2)
		    ndocumento_edi=mid(edi,17,6)
    		
		    'con_emisor=d_lookup("dsn","clientes","ncliente='" & ncliente_origen_edi & "'",DSNIlion)
            con_emisor_select = "select dsn from clientes with(nolock) where ncliente= ?"
            con_emisor= DLookupP1(con_emisor_select, ncliente_origen_edi & "", adChar, 5, DSNIlion)
            'cif_emisor=d_lookup("cifedi","clientes","ncliente='" & ncliente_origen_edi & "'",DSNIlion)
            cif_emisor_select="select cifedi from clientes with(nolock) where ncliente= ?"
            cif_emisor=DLookupP1(cif_emisor_select, ncliente_origen_edi & "", adChar, 5, DSNIlion)
            'nserie_doc=d_lookup("nserie","series","edi='" & cint(nserie_edi) & "' and nserie like '" & ncliente_origen_edi & "%'",con_emisor)
            nserie_doc_select= "select nserie from series with(nolock) where edi= ? and nserie like = ?+'%'"
            nserie_doc=DLookupP2(nserie_doc_select, cint(nserie_edi) & "", adSmallInt, 2, ncliente_origen_edi & "", adVarchar, 10, con_emisor)
            'empresa_sup=d_lookup("empresa_sup", "configuracion", "nempresa='"& session("ncliente") &"'", session("dsn_cliente"))
            empresa_sup_select="select empresa_sup from configuracion with(nolock) where nempresa= ?"
            empresa_sup= DLookupP1(empresa_sup_select, session("ncliente") &"", adChar, 5, session("dsn_cliente"))
            'dsnCentral=d_lookup("dsn", "clientes", "ncliente='"&empresa_sup&"'", DsnIlion)
            dsnCentralSelect= "select dsn from clientes with(nolock) where ncliente = ?"
            dsnCentral=DLookupP1(dsnCentralSelect, empresa_sup&"", adChar, 5, DsnIlion)

            'ncliente_part=d_lookup("ncliente","clientes","cifedi='" & cif_emisor & "' and ncliente like '" & empresa_sup & "%'",dsnCentral)
            ncliente_part_select = "select ncliente from clientes with(nolock) where cifedi= ? and ncliente like ?+'%'"
            ncliente_part=DLookupP2(ncliente_part_select, cif_emisor & "", adVarchar, 20, empresa_sup & "", adChar, 10, dsnCentral)
            'nserie_doc=d_lookup("nserie","series","edi='" & cint(nserie_edi) & "' and nserie like '" & ncliente_origen_edi & "%'",con_emisor)
    		nserie_doc_select= "select nserie from series with(nolock) where edi= ? and nserie like ?+'%'"
            nserie_doc= DLookupP2(nserie_doc_select, cint(nserie_edi) & "", adSmallInt, 2, ncliente_origen_edi & "", adVarchar, 10, con_emisor) 
		    'serieCliente=d_lookup("serie_ped","documentos_cli","ncliente='" & ncliente_part & "' and ncliente like '" & empresa_sup & "%'",dsnCentral)
		    serieClienteSelect = "select serie_ped from documentos_cli with(nolock) where ncliente= ? and ncliente like ?+'%'"
            serieCliente= DLookupP2(serieClienteSelect, ncliente_part & "", adChar, 10,  empresa_sup & "", adChar, 10, dsnCentral)
            
            doc=nserie_doc & any_serie_edi & ndocumento_edi
		    if serieCliente&""<="" then
		        'serieCliente=d_lookup("nserie","series","nserie like '" & empresa_sup & "%' and tipo_documento = 'PEDIDO DE CLIENTE' and pordefecto = 1 ",dsnCentral)
		        serieClienteSelect= "select nserie from series with(nolock) where nserie like ?+'%' and tipo_documento = 'PEDIDO DE CLIENTE' and pordefecto = 1"
                serieClente=DLookupP1(serieClienteSelect, empresa_sup & "", adVarchar, 10, dsnCentral)
                if serieCliente&""<="" then
		            %>
		            <script language="javascript" type="text/javascript">
                        window.alert("<%=LITERROR_PED_CENTRAL%> \n <%=LITERROR_SERIEDEFECTO%>");
                        document.location = "pedidos_pro.asp?npedido=<%=enc.EncodeForJavascript(doc)%>&mode=browse"
		            </script><%
		        end if
		    end if
		    if ncliente_part&""<="" then%>
	            <script language="javascript" type="text/javascript">
                        window.alert("<%=LITERROR_PED_CENTRAL%> \n <%=LITERROR_CLIENTEENCENTRAL%>");
                        document.location = "pedidos_pro.asp?npedido=<%=enc.EncodeForJavascript(doc)%>&mode=browse"
	            </script><%
		    end if
		    'nproveedorTmp=d_lookup("nproveedor","pedidos_pro","npedido='" & doc & "'",con_emisor)
            nproveedorTmpSelect = "select nproveedor from pedidos_pro with(nolock) where npedido= ?"
            nproveedorTmp= DLookupP1(nproveedorTmpSelect, doc & "", adVarchar, 20, con_emisor)
		end if

        ha_habido_un_error=0	
        set cnn = Server.CreateObject("ADODB.Connection")
        cnn.open session("dsn_cliente")	
        On Error resume next
        cnn.BeginTrans
        'Actualizamos detalles pedidos de compra
        no_seguir=0

        if dsnCentral & "">"" then
            Resultado=ComprobarPreciosEnCentral(doc,npedido_part,dsnCentral,ncliente_part,rpc)
        else
            ''ha_habido_un_error=1
            if rpc & "">"" then
                no_seguir=1
            end if
        end if

        if ha_habido_un_error<>1 and no_seguir=0 then
            'Actualizamos el total de pedido de compra
		    rst.Open "exec ActualizaTotalesPedidoCompra '"& session("ncliente") &"','"& doc &"',0",cnn,adOpenKeyset,adLockOptimistic
		    if err.number<>0 then
                ha_habido_un_error=1
            end if    
		    rst.Close

		    if ha_habido_un_error=0 then
                cnn.CommitTrans
            else
                cnn.RollbackTrans
            end if
            cnn.Close
            on error goto 0   
        end if
        
        if ha_habido_un_error=0 then
            if rpc & "">"" then
                %><script type="text/javascript" language="javascript">
                        window.onload = function (event) {
                            //window.alert("<%=formato_impresionEleg%>");
                            no_seguir = "<%=no_seguir%>";
                            if (no_seguir == "0" || no_seguir == "") {
                                formatoImpresion = "<%=replace(formato_impresionEleg,"'","\'")%>";
                                if (formatoImpresion.indexOf("crearpdf") != -1) {
                                    window.alert("<%=LITOK_PED_CENTRAL_RPC2%>");
                                }
                                else {
                                    window.alert("<%=LITOK_PED_CENTRAL_RPC%>");
                                }
                            }
                            AbrirVentana('<%=replace(formato_impresionEleg,"'","\'")%>', 'I', <%=AltoVentana %>, <%=AnchoVentana %>);
                        }
                </script><%
            else
                cnn.open dsnCentral	
                On Error resume next
                cnn.BeginTrans
                           
                'Creamos pedido de venta en central
                rst.cursorlocation=2
		        rst.open "select * from pedidos_cli where npedido=''",cnn,adOpenKeyset,adLockOptimistic
		        'rstAux2.open "select * from pedidos_pro with(nolock) where npedido='" & doc & "'",con_emisor,adUseClient,adLockReadOnly
		        
                set commandAux = nothing
                set connAux = Server.CreateObject("ADODB.Connection")
                set commandAux =  Server.CreateObject("ADODB.Command")

                connAux.open con_emisor
                connAux.cursorlocation=2
                commandAux.ActiveConnection =connAux
                commandAux.CommandTimeout = 60
                commandAux.CommandText= "select * from pedidos_pro with(nolock) where npedido=?"
                commandAux.CommandType = adCmdText
                commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, doc)

                set rstAux2 = commandAux.Execute
                            
                npedido_part=CabeceraPedido("",serieCliente,date,ncliente_part,empresa_sup,cnn)
		        rstAux2.close
		        rst.close
		        Resultado=PasarDetallesPedido(doc,npedido_part,cnn,ncliente_part)
		        PreciosPed npedido_part,cnn
		        WaitBoxOculto LitEsperePorFavor

		        if Resultado = "OK" and ha_habido_un_error=0 then
    		        cnn.CommitTrans
		            %><script language="javascript" type="text/javascript">		          window.alert("<%=LITOK_PED_CENTRAL%>")</script><%
		            ' Auditamos en la tabla historial_proveedor
                    nanotacion=d_lookup("isnull(max(nanotacion),0)+1","historial_proveedor","",session("dsn_cliente"))

		            strAuditar = "insert into historial_proveedor (nproveedor,anotacion,fecha,nanotacion) " & _ 
		                         "values ('"&nproveedorTmp&"','Pedido "& right(doc,Len(doc)-5) &" enviado a central','"& date() & " " & iif(Len(Hour(Now))<2,"0"&Hour(Now),Hour(Now)) & ":" & iif(Len(Minute(Now))<2,"0"&Minute(Now),Minute(Now)) & ":" & iif(Len(Second(Now))<2,"0"&Second(Now),Second(Now)) &"',"&nanotacion&")"
                    rst.cursorlocation=2
		            rst.open strAuditar,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

		        else
		            cnn.RollbackTrans
		            %>
		            <script language="javascript" type="text/javascript">
                        window.alert("<%=LITERROR_PED_CENTRAL%>")
                    </script>
		            <%
		        end if
    		    
		        cnn.Close
                set cnn=nothing
                on error goto 0	
                
		        If rstAux.State <> adStateClosed Then
                    rstAux.Close
                End If
		        If rstAux2.State <> adStateClosed Then
                    rstAux2.Close
                End If
            end if
        else
            set cnn =nothing
            if no_seguir=0 then%>
                <script language="javascript" type="text/javascript">
                        window.alert("<%=LITERROR_PED_CENTRAL%>")
                </script>
            <%end if
        end if
	end if

	'JMA 16/12/04. Copiar campos personalizables de los proveedores'
	redim tmp_lista_valores(10)
	for ki=1 to 10
		tmp_lista_valores(ki)=""
	next
	'JMA 16/12/04. FIN Copiar campos personalizables de los proveedores'

	'JMA 16/12/04 si existen campos personalizables con titulo no nulo saldrán los campos personalizables'
	si_campo_personalizables=0
    'rst.cursorlocation=3
	'rst.open "select ncampo from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and titulo is not null and titulo<>'' and ncampo like '" & session("ncliente") & "%'",session("dsn_cliente")

    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open session("dsn_cliente")
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "select ncampo from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and titulo is not null and titulo<>'' and ncampo like ?+'%'"
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@ncampo", adChar,adParamInput,7, session("ncliente"))

    set rst = commandAux.Execute

	if not rst.eof then
		si_campo_personalizables=1
	else
		si_campo_personalizables=0
	end if

    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	'rst.close
    %>
	<input type="hidden" name="si_campo_personalizables" value="<%=si_campo_personalizables%>"/>
	<%'JMA 16/12/04 FIN si existen campos personalizables con titulo no nulo saldrán los campos personalizables'
    
    if tmp_serie="" and mode="add" then
	    'Obtener la serie por defecto
	    tmp_serie=ObtenerSerieTienda("PEDIDO A PROVEEDOR")
		
	    'completamos la cadena del proveedor si es necesario
	    if nproveedor <> "" and nproveedor <> null then
	        if len(nproveedor)<5 then
	            for i=len(nproveedor) to 4
	                nproveedor= "0" & nproveedor
	            next
	        end if
	    end if

	    if tmp_serie & ""="" then
	        'mmg:
	        'tmp_serie=d_lookup("serie_ped","documentos_pro","tipo_documento='PEDIDO A PROVEEDOR' and pordefecto=1 and nserie like '" & session("ncliente") & "%'", session("dsn_cliente"))
		    
	        set rsAux = Server.CreateObject("ADODB.Recordset")
	        cadena= "exec ObtenerSeriePedPro '" & session("ncliente")& nproveedor & "','" & session("ncliente") & "'"
	        rsAux.Open cadena,session("dsn_cliente")
	        tmp_serie = rsAux("serie")
	    end if
	end if

	if nproveedor="" and mode="add"  then
		'Obtener el proveedor de la serie por defecto.
		'mmg: ejecutamos el proc alm ObtenerProveedorPedPro
		set rsAux = Server.CreateObject("ADODB.Recordset")
		cadena= "exec ObtenerProveedorPedPro '" & tmp_serie & "'"
		rsAux.Open cadena,session("dsn_cliente")
		nproveedor = rsAux("proveedor")
	end if

    'mmg:
	set rstMM = Server.CreateObject("ADODB.Recordset")
    'rstMM.cursorlocation=3
	'rstMM.open "select almacen from series with(nolock) where nserie='"&tmp_serie&"' "& strwhereMM,session("dsn_cliente")
    
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open session("dsn_cliente")
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "select almacen from series with(nolock) where nserie=? "& strwhereMM
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@nserie", adVarChar,adParamInput,10, tmp_serie)

    set rstMM = commandAux.Execute

	if not rstMM.EOF then
		almacenSerie= rstMM("almacen")
	else
		almacenSerie= ""
	end if

    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	'rstMM.close
	
	'JMA 16/12/04 añadir campos personalizables a pedidos_pro'
	if mode="browse" or mode="edit" or mode="add" or mode="save" or mode="first_save" then
		num_campos=0
		if mode="add" then
			redim lista_valores(10+2)
			for ki=1 to 12
				lista_valores(ki)=""
			next
			num_campos=10
		else
			'rstAux2.cursorlocation=3
			'rstAux2.open "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from pedidos_pro as p with(nolock) where p.npedido='" & npedido & "'",session("dsn_cliente")
		
            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open session("dsn_cliente")
            connAux.cursorlocation=3
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText=  "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from pedidos_pro as p with(nolock) where p.npedido=?"
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

            set rstAux2 = commandAux.Execute

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
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
			'rstAux2.close
		end if
	end if
	'JMA 16/12/04 añadir campos personalizables a pedidos_pro'
		
    'mmg >> se actualiza la serie cuando se cambia el proveedor
    p_serie=limpiaCadena(request.QueryString("serie"))
    p_serieR=request.Form("serie")
    p_nproveedor=limpiaCadena(Request.Form("nproveedor"))
    provR=limpiaCadena(request.querystring("prov"))
    nprov=limpiaCadena(request.querystring("nproveedor"))

    if p_serieR<>"" and p_nproveedor<>"" then
        TraerProveedor= Completar(p_nproveedor,5,"0")
        nproveedor=TraerProveedor
        tmp_nproveedor=TraerProveedor
        if cint(null_z(cambiar_serie))=1 or cambiar_serie & ""="" then
            'rstAux.cursorlocation=3
            'rstAux.open "select serie_ped from documentos_pro with(nolock) where nproveedor='" & session("ncliente") & nproveedor & "'", session("dsn_cliente")

            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open session("dsn_cliente")
            connAux.cursorlocation=3
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText=  "select serie_ped from documentos_pro with(nolock) where nproveedor=?"
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@nproveedor", adChar,adParamInput,10, session("ncliente") & nproveedor)

            set rstAux = commandAux.Execute

	        if not rstAux.eof then
		        if rstAux("serie_ped")&"">"" then
			        p_serieR=rstAux("serie_ped")
			        tmp_serie=rstAux("serie_ped")
		        else
			        tmp_serie=p_serieR
		        end if
		    else
		        tmp_serie=p_serieR
	        end if
            connAux.Close
            set connAux = nothing
            set commandAux = nothing

	        'rstAux.close
	    end if
    else
        if p_serieR="" and request.QueryString("nproveedor")<>"" then
            TraerProveedor= Completar(nprov,5,"0")
            nproveedor=TraerProveedor
            tmp_nproveedor=TraerProveedor
            if cint(null_z(cambiar_serie))=1 or cambiar_serie & ""="" then
                'rstAux.cursorlocation=3
                'rstAux.open "select serie_ped from documentos_pro with(nolock) where nproveedor='" & session("ncliente")&TraerProveedor & "'", session("dsn_cliente")
	    
                set commandAux = nothing
                set connAux = Server.CreateObject("ADODB.Connection")
                set commandAux =  Server.CreateObject("ADODB.Command")

                connAux.open session("dsn_cliente")
                connAux.cursorlocation=3
                commandAux.ActiveConnection =connAux
                commandAux.CommandTimeout = 60
                commandAux.CommandText=  "select serie_ped from documentos_pro with(nolock) where nproveedor=?"
                commandAux.CommandType = adCmdText
                commandAux.Parameters.Append commandAux.CreateParameter("@nproveedor", adChar,adParamInput,10, session("ncliente")&TraerProveedor)

                set rstAux = commandAux.Execute

                if not rstAux.eof then
		            if rstAux("serie_ped")&"">"" then
			            p_serieR=rstAux("serie_ped")
			            tmp_serie=rstAux("serie_ped")
		            else
			            tmp_serie=p_serie
		            end if
		        else
		            tmp_serie=p_serie
	            end if
                connAux.Close
                set connAux = nothing
                set commandAux = nothing
	            'rstAux.close
	        end if
        else
            if TraerProveedor="" and mode="add" then
	            'Obtener el proveedor de la serie por defecto.
	            'TraerProveedor=d_lookup("substring(cliente,6,10)","series","nserie='" & tmp_serie & "'",session("dsn_cliente"))
                TraerProveedorSelect = "select substring(cliente,6,10) from series with(nolock) where nserie= ?"
                TraerProveedor=DLookupP1(TraerProveedorSelect, tmp_serie & "", adVarchar, 10, session("dsn_cliente"))
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
			nproveedor=session("ncliente") & Completar(nproveedor,5,"0")
		end if
        'rstAux.cursorlocation=3
		'rstAux.open "select fbaja from proveedores with(nolock) where nproveedor='" & nproveedor & "'", session("dsn_cliente")

        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=3
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText=  "select fbaja from proveedores with(nolock) where nproveedor=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@nproveedor", adChar,adParamInput,10, nproveedor)

        set rstAux = commandAux.Execute

		if not rstAux.eof then
			if rstAux("fbaja")>"" then%>
				<script language="javascript" type="text/javascript">window.alert("<%=LitProvDadoBaja%>");</script>
				<%nproveedor=""
			end if
		end if
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
		'rstAux.close
	end if

	'Captura de datos del articulo que se está introduciendo en el detalle
	if nproveedor > "" then
		if len(nproveedor)<=5 then
			nproveedor=session("ncliente") & Completar(nproveedor,5,"0")
		end if
		Error="NO"
        'rstAux.cursorlocation=3
		'rstAux.open "select * from proveedores with(nolock) where nproveedor='" & nproveedor & "'",session("dsn_cliente")

        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=3
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText=  "select * from proveedores with(nolock) where nproveedor=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@nproveedor", adChar,adParamInput,10, nproveedor)

        set rstAux = commandAux.Execute

		if not rstAux.EOF then
	  		tmp_nproveedor=nproveedor
			tmp_descripcion=rstAux("razon_social")
			tmp_forma_pago=rstAux("forma_pago")
			tmp_tipo_pago=rstAux("tipo_pago")
			tmp_divisa=rstAux("divisa")
			tmp_descuento=rstAux("descuento")
			tmp_descuento2=rstAux("descuento2")
			tmp_recargo=rstAux("recargo")
			tmp_irpf=rstAux("irpf")
			tmp_IRPF_Total=rstAux("IRPF_Total")
			tmp_observaciones=observacionesR
			tmp_notas=notasR
			tmp_cod_proyecto=rstAux("proyecto")
            tmp_portes=rstAux("portes")

			'JMA 16/12/04: Captura de los campos personalizables del proveedor'
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
			'JMA 16/12/04: FIN Captura de los campos personalizables del proveedor'
			if cint(null_z(cambiar_serie))=1 or cambiar_serie & ""="" then
				'obtener_doc_cli mode,"albaranes_pro",tmp_nproveedor,p_serie,p_valorado,tmp_valorado,tmp_irpf
				
				set rstObtDocCli= Server.CreateObject("ADODB.Recordset")
                dato_doc1="serie_ped"
                dato_doc2="valorado_ped"
                if mode="add" then
	                strselectdoc="select "
	                strselectdoc=strselectdoc & dato_doc1
	                if dato_doc2 & "">"" then
		                strselectdoc=strselectdoc & "," & dato_doc2
	                end if
	                strselectdoc=strselectdoc & ",irpf"
	                strselectdoc=strselectdoc & " from documentos_pro dc left outer join series s with(nolock) on dc." & dato_doc1 & "=s.nserie left outer join empresas e on s.empresa=e.cif where nproveedor= ?"
                    'rstObtDocCli.cursorlocation=3
	                'rstObtDocCli.open strselectdoc, session("dsn_cliente")

                    set commandDocCli = nothing
                    set connDocCli = Server.CreateObject("ADODB.Connection")
                    set commandDocCli =  Server.CreateObject("ADODB.Command")

                    connDocCli.open session("dsn_cliente")
                    connDocCli.cursorlocation=3
                    commandDocCli.ActiveConnection =connDocCli
                    commandDocCli.CommandTimeout = 60
                    commandDocCli.CommandText= strselectdoc
                    commandDocCli.CommandType = adCmdText
                    commandDocCli.Parameters.Append commandDocCli.CreateParameter("@nproveedor", adChar,adParamInput,10, tmp_nproveedor)

                    set rstObtDocCli = commandDocCli.Execute

	                if not rstObtDocCli.eof then
		                if rstObtDocCli(dato_doc1) & "">"" then
			                tmp_serie=rstObtDocCli(dato_doc1)
		                end if
		                if dato_doc2 & "">"" then
			                tmp_valorado=rstObtDocCli(dato_doc2)
		                end if
		                'FLM:010409: comento esto xq no se debe obtener el irpf de la tabla empresas, si no del proveedor.
		                'if rstObtDocCli("irpf") & "">"" then
			            '    tmp_irpf=rstObtDocCli("irpf")
		                'end if
	                end if
                    connDocCli.Close
                    set connDocCli = nothing
                    set commandDocCli = nothing
	                'rstObtDocCli.close
                end if
                set rstObtDocCli=nothing
			end if
		else
			Error="SI"%>
  			<script language="javascript" type="text/javascript">
                  window.alert("<%=LitMsgProveedorNoExiste%>");
                  history.back();
			</script>
        <%end if
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
	'	rstAux.close
	end if
  'Acción a realizar

    if mode="save" or mode="first_save" then
        ModDocumento=true
        'comprobamos si el npedido existe o no segun el contador de configuracion
        if mode="first_save" then
	        if compNumDocNuevo(serieR,fechaR,"pedidos_pro")=0 then%>
		        <script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgDocExistRevCont%>");
                    document.location = "pedidos_pro.asp?mode=add"
                    parent.botones.document.location = "pedidos_pro_bt.asp?mode=add"
		        </script>
		        <%ModDocumento=false
		        el_pedido_existe=1
	        end if
        end if
        if ModDocumento=true then
            rst.cursorlocation=2
	        'rst.Open "select * from pedidos_pro where npedido='" & npedidoH & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	        
            set commandAuxx = nothing
            set connAuxx = Server.CreateObject("ADODB.Connection")
            set commandAuxx =  Server.CreateObject("ADODB.Command")

            connAuxx.open session("dsn_cliente")
            connAuxx.cursorlocation=2
            commandAuxx.ActiveConnection =connAuxx
            commandAuxx.CommandTimeout = 60
            commandAuxx.CommandText= "select * from pedidos_pro where npedido=? "
            commandAuxx.CommandType = adCmdText
            commandAuxx.Parameters.Append commandAuxx.CreateParameter("@npedido", adVarChar,adParamInput,20, npedidoH)

            rst.Open commandAuxx, , adOpenKeyset, adLockOptimistic

            GuardarRegistro npedidoH,serieR
	        npedido=rst("npedido")
	        if mode="first_save" then
		        auditar_ins_bor session("usuario"),npedido,rst("nproveedor"),"alta","","","pedidos_pro"
	        end if

	        rst.close
        end if
        ant_mode=mode
        mode="browse"
    elseif mode="delete" then
		he_borrado=1
		'rst.open "select nproveedor,nfactura,nalbaran from pedidos_pro with(nolock) where npedido='" & npedidoH & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=2
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select nproveedor,nfactura,nalbaran from pedidos_pro with(nolock) where npedido=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedidoH)
        
        rst.Open commandAux, , adOpenKeyset, adLockOptimistic

        if not rst.eof then
			npro_aux=rst("nproveedor")
			nfac_aux=rst("nfactura")
			nalb_aux=rst("nalbaran")
			connAux.Close
            set connAux = nothing
            set commandAux = nothing
            'rst.close

			if BorrarRegistro(npedidoH)=true then
				if isnull(nfac_aux) and isnull(nalb_aux) then
					auditar_ins_bor session("usuario"),npedidoH,npro_aux,"baja","","","pedidos_pro"
				end if
				mode="add"
				npedido=""
                %><script language="javascript" type="text/javascript">
                      parent.botones.document.location = "pedidos_pro_bt.asp?mode=add";
                      SearchPage("purchaseOrder_lsearch.asp?mode=init", 0);
			    </script><%
			else
				mode="browse"
			end if
' >>> MCA 21/04/05 Para cargar el modo add tras el borrado
			npedidoH= ""
		else
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
			'rst.close%>
			<script language="javascript" type="text/javascript">
                      window.alert("<%=LitMsgDocsNoExiste%>");
                      parent.botones.document.location = "pedidos_pro_bt.asp?mode=add";
			</script>
			<%mode="add"
			npedido=""
		end if
    end if

    total_iva_bruto	= null_z(d_sum("(importe*iva)/100","detalles_ped_pro","npedido='" & npedido & "'",session("dsn_cliente")))
    total_re_bruto	= null_z(d_sum("(importe*re)/100","detalles_ped_pro","npedido='" & npedido & "'",session("dsn_cliente")))

    'MB=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
    MBSelect = "select codigo from divisas with(nolock) where moneda_base<>0 and codigo like ?+'%'"
    MB= DLookupP1(MBSelect, session("ncliente") & "", adVarchar, 15, session("dsn_cliente"))
    'Mostrar los datos de la página.

    ''ricardo 31/7/2003 comprobamos que existe el albaran
    if mode="browse" and he_borrado<>1 and el_pedido_existe<>1 then
        'rstAux.cursorlocation=3
	    'rstAux.open "select npedido from pedidos_pro with(nolock) where npedido='" & npedido & "'", session("dsn_cliente")
	    
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=3
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select npedido from pedidos_pro with(nolock) where npedido=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        set rstAux = commandAux.Execute
                
        if rstAux.eof then
		    npedido=""%>
		    <script language="javascript" type="text/javascript">
                      window.alert("<%=LitMsgDocsNoExiste%>");
                      parent.botones.document.location = "pedidos_pro_bt.asp?mode=add";
		    </script>
		    <%mode="add"
	    end if
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
	    'rstAux.close
    end if

    if mode="browse" or mode="edit" then
	    if npedido="" then
            'rstAux.cursorlocation=3
		    'rstAux.open "select top 1 npedido from pedidos_pro with(nolock) where npedido like '" & session("ncliente") & "%' order by fecha desc,npedido desc", session("dsn_cliente")
            
            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open session("dsn_cliente")
            connAux.cursorlocation=3
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText= "select top 1 npedido from pedidos_pro with(nolock) where npedido like ?+'%' order by fecha desc,npedido desc"
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, session("ncliente"))

            set rstAux = commandAux.Execute

            if not rstAux.eof then npedido=rstAux("npedido")

            connAux.Close
            set connAux = nothing
            set commandAux = nothing
		    'rstAux.close
	    end if

	    ' JMA 16/12/04 Campos personalizables'
	    'rstAux.cursorlocation=3
	    'rstAux.open "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from pedidos_pro as p with(nolock) where p.npedido='" & npedido & "'",session("dsn_cliente")
	    
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=3
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from pedidos_pro as p with(nolock) where p.npedido=?"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        set rstAux = commandAux.Execute

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

        connAux.Close
        set connAux = nothing
        set commandAux = nothing
	    'rstAux.close
	    ' JMA 16/12/04 FIN Campos personalizables'

	    'rst.cursorlocation=3
	    'strselect="select f.*,d.abreviatura,d.ndecimales,c.razon_social from pedidos_pro as f,proveedores as c,divisas as d where c.nproveedor=f.nproveedor " & _
	    '"and npedido='" & npedido & "'and f.divisa=d.codigo"
	    'rst.Open strselect, session("dsn_cliente")

        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux =  Server.CreateObject("ADODB.Command")

        connAux.open session("dsn_cliente")
        connAux.cursorlocation=3
        commandAux.ActiveConnection =connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText= "select f.*,d.abreviatura,d.ndecimales,c.razon_social from pedidos_pro as f,proveedores as c,divisas as d where c.nproveedor=f.nproveedor and npedido=? and f.divisa=d.codigo"
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

        set rst = commandAux.Execute

    elseif mode="add" then
	    'rst.Open "select *,'' as razon_social from pedidos_pro where npedido='" & npedido & "'", _
	    'session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	    
        set commando = nothing
        set conno = Server.CreateObject("ADODB.Connection")
        set commando =  Server.CreateObject("ADODB.Command")

        conno.open session("dsn_cliente")
        conno.cursorlocation=3
        commando.ActiveConnection =conno
        commando.CommandTimeout = 60
        commando.CommandText= "select *,'' as razon_social from pedidos_pro where npedido=?"
        commando.CommandType = adCmdText
        commando.Parameters.Append commando.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido &"")

        rst.Open commando, , adOpenKeyset, adLockOptimistic 

	    rst.AddNew
	    rst("valorado")=1
    elseif mode="search" then
	    if cstr(npagina)<>"0" then
	        lote=npagina
	    else
            if sentido="next" then
                lote=lote+1
            elseif sentido="prev" then
                lote=lote-1
            elseif sentido="first" then
                lote=1
            elseif sentido="last" then
                if total_paginas & ""="" or cstr(total_paginas)="0" then
                    lote=1
                else
                    lote=total_paginas
                end if
            end if
        end if

        set conn1 = Server.CreateObject("ADODB.Connection")
        set command1 =  Server.CreateObject("ADODB.Command")

        conn1.open session("dsn_cliente")
        command1.ActiveConnection =conn1
        command1.CommandTimeout = 0
        command1.CommandText="PedidosProBuscar"
        command1.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        command1.Parameters.Append command1.CreateParameter("@nempresa", adVarChar, adParamInput, 5, session("ncliente"))
        command1.Parameters.Append command1.CreateParameter("@campo", adVarChar, adParamInput, 50, Nulear(campo))
        command1.Parameters.Append command1.CreateParameter("@criterio", adVarChar, adParamInput, 50,Nulear(criterio))
        command1.Parameters.Append command1.CreateParameter("@texto", adVarChar, adParamInput, 50,Nulear(texto))
        command1.Parameters.Append command1.CreateParameter("@PageSize", adInteger, adParamInput, 20,NumReg)
        command1.Parameters.Append command1.CreateParameter("@PageNumber", adInteger, adParamInput, 50,lote)
        command1.Parameters.Append command1.CreateParameter("@series", adVarChar, adParamInput, 500,s)
        command1.Execute,,adExecuteRecords
        set rst=command1.execute
    end if
	sumadet=0
	sumaRE=0

	VinculosPagina(MostrarProveedores)=1
	CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
   
    'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION
    


    if mode="edit" then
	    %><input type="hidden" name="h_fecha" value="<%=EncodeForHtml(rst("fecha"))%>"/><%
	    %>
        <input type="hidden" name="h_formapago" value="<%=EncodeForHtml(rst("forma_pago"))%>"/><%
        ''FALLA EncodeForHtmlAttribute(rst("forma_pago"))
    end if
        
    %><div class="headers-wrapper"><%
        DrawDiv "header-date","",""
            DrawLabel "txtMandatory","",LitFecha
            if mode="browse" then
                DrawSpan "","",EncodeForHtml(rst("fecha")),""
            else
                DrawInput "width50","","fecha",EncodeForHtml(iif(mode="add",iif(tmp_fecha>"",tmp_fecha,date()),rst("fecha"))),"onchange='javasript:CalculaFechaPago();'"
                DrawCalendar "fecha"
            end if
        CloseDiv
        DrawDiv "header-npedido","",""
            DrawLabel "","",Litpedido
            DrawSpan "","",EncodeForHtml(trimCodEmpresa(rst("npedido"))),""
        CloseDiv

        DrawDiv "header-nproveedor","",""
            Drawlabel "txtMandatory","",LitProveedor
            if mode="browse" then
				if rst("nproveedor")>"" then%>
						<%=Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("nproveedor"))),LitVerProveedor)%>
				<%end if
			else
      			%><input class="CELDA width20" type="text" size="8" name="nproveedor" value="<%=EncodeForHtml(trimCodEmpresa(iif(nproveedor>"",nproveedor,rst("nproveedor"))))%>" onchange="TraerSerie('<%=EncodeForHtml(mode)%>');"/>
	      		<a class="CELDAREFB" href="javascript:AbrirVentana('proveedores_busqueda.asp?ndoc=pedidos_pro&titulo=<%=LitSelProv%>&mode=search','P',<%=altoventana%>,<%=anchoventana%>)">
                      <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="" title=""/>
	      		</a>
    		<%end if
            'nompro   = d_lookup("razon_social","proveedores","nproveedor='" & iif(nproveedor>"",nproveedor,rst("nproveedor")) & "'",session("dsn_cliente"))
			nomproSelect = "select razon_social from proveedores with(nolock) where nproveedor= ?"
            nompro = DLookupP1(nomproSelect, iif(nproveedor>"",nproveedor,rst("nproveedor")) & "", adChar, 10, session("dsn_cliente"))
            num_cuenta = ""
			'num_cuenta = d_lookup("cuenta_cargo","proveedores","nproveedor='" & iif(nproveedor>"",nproveedor,rst("nproveedor")) & "'",session("dsn_cliente"))
            num_cuenta_select = "select cuenta_cargo from proveedores with(nolock) where nproveedor= ?"
            num_cuenta = DLookupP1(num_cuenta_select, iif(nproveedor>"",nproveedor,rst("nproveedor")) & "",  adChar, 10, session("dsn_cliente"))

            %><input type="hidden" name="ncuentacargo" value="<%=EncodeForHtml(num_cuenta)%>"/><%
			if mode="edit" or mode="add" then
                %><input class="CELDA width30" type="text" disabled name="nombre" value="<%=EncodeForHtml(nompro)%>"/><%
            elseif mode="browse" then
                DrawSpan "","","&nbsp;&nbsp;" & EncodeForHtml(nompro),""
            end if
        CloseDiv

        if mode="browse" and not rst.EOF then
            DrawDiv "header-note","",""
	        ''ricardo 30-5-2007 si el parametro cpc=0 no se pondran las cajas
		    if (rst("nfactura")&"")="" and (rst("nalbaran")&"")="" and cstr(cpc)<>"0" then
			    'MB=d_lookup("codigo","divisas","moneda_base<>0",session("dsn_cliente"))
			    'n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") &"'",session("dsn_cliente"))
			    EnCaja=CambioDivisa(d_sum("importe","pagos_ped_pro","npedido='" & rst("npedido") & "'",session("dsn_cliente")),rst("divisa"),rst("divisa"))
			    Pendiente=miround(rst("total_pedido")-EnCaja,rst("ndecimales"))
			    defecto=""
			    poner_cajasResponsive1 "input-ncaja",defecto,"ncaja","100","codigo","descripcion","","",poner_comillas(caju)%>
                <span class='header-note-inputCaja'>
                    <input class="CELDAR7" type="Text" name="impcaja" value="<%=EncodeForHtml(Pendiente)%>"/>
                </span>
                <span class='header-note-currency'>
                    <font id="fntAbrev" class=ENCABEZADOR7><%=EncodeForHtml(null_s(rst("abreviatura")))%></font>
                </span>
                <span class="header-note-buttonNote">
                    <img id="imgAnotar" src="<%=themeIlion %><%=ImgAnotar%>" <%=ParamImgAnotar%> alt="<%=LitAnotarCaja%>" title="<%=LitAnotarCaja%>" onclick="Acaja('<%=EncodeForHtml(rst("npedido"))%>')" style="cursor:pointer"/>
                </span>
                <%
                'rstAux.cursorlocation=3
                'rstAux.Open "SELECT * FROM Tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by Descripcion",session("dsn_cliente")
                 
                set commandAux = nothing
                set connAux = Server.CreateObject("ADODB.Connection")
                set commandAux =  Server.CreateObject("ADODB.Command")

                connAux.open session("dsn_cliente")
                connAux.cursorlocation=3
                commandAux.ActiveConnection =connAux
                commandAux.CommandTimeout = 60
                commandAux.CommandText= "SELECT * FROM Tipo_pago with(nolock) where codigo like ?+'%' order by Descripcion"                
                commandAux.CommandType = adCmdText
                commandAux.Parameters.Append commandAux.CreateParameter("@codigo", adVarChar,adParamInput,8, session("ncliente"))

                set rstAux = commandAux.Execute

                DrawSelect "input-i_pago","width:150px;","i_pago",rstAux,session("ncliente") & "01","codigo","Descripcion","",""
                
                connAux.Close
                set connAux = nothing
                set commandAux = nothing
                'rstAux.Close
		    end if
            CloseDiv
	    end if
 		if mode="browse" or mode="save" then
            if si_tiene_modulo_21 = 0 and si_tiene_modulo_22 = 0 then 
                DrawDiv "header-resources alignCenter", "", ""
                    %><a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=Servicios/recursos.asp&pag2=Servicios/recursos_bt.asp&codigo=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>&tipo=pedido a proveedor&viene=enlaces&usuario=<%=session("usuario")%>', 'P', <%=AltoVentana%>, <%=AnchoVentana%>)">&nbsp;&nbsp;&nbsp;<%=LitEnlaces%>&nbsp;&nbsp;&nbsp;</a><%            
                CloseDiv
            end if
			''ricardo 13-3-20003
            ''si la serie tiene un formato de impresion sera este el de por defecto
            ''si no sera el elegido en la tabla formatos impresion de ilion
            if not rst.eof then
	            defecto=obtener_formato_imp(rst("serie"),"PEDIDO A PROVEEDOR")
            end if
            ''''''''
            			 
			'JMMM - 05/11/2010 -> Mostrar crear pedido en central (franquicias)
			'seleccion = "select empresa_sup from configuracion with(nolock) where nempresa = '"&session("ncliente")&"'"
            'rstSelect.cursorlocation=3
			'rstSelect.Open seleccion, session("dsn_cliente")

            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open session("dsn_cliente")
            connAux.cursorlocation=3
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText="select empresa_sup from configuracion with(nolock) where nempresa = ?"                
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@nempresa", adChar,adParamInput,5, session("ncliente"))

            set rstSelect = commandAux.Execute

			if not rstSelect.EOF then
			    empresa_sup = rstSelect("empresa_sup")
			    if empresa_sup&"">"" then
			        ' Si tiene empresa_sup (central) se comprueba si la central tiene el módulo de franquicias y que es franquiciador
			        'dsnCentral=d_lookup("dsn", "clientes", "ncliente='"&empresa_sup&"'", DsnIlion)
                    dsnCentralSelect = "select dsn from clientes with(nolock) where ncliente= ?"
                    dsnCentral = DLookupP1(dsnCentralSelect, empresa_sup&"", adChar, 10, DsnIlion)
			        'esFranquiciador=d_lookup("franquiciador","configuracion","nempresa='" & empresa_sup & "'",dsnCentral)
                    esFranquiciadorSelect = "select franquiciador from configuracion with(nolock) where nempresa= ?"
                    esFranquiciador= DLookupP1(esFranquiciadorSelect, empresa_sup & "", adChar, 5, dsnCentral)

			        si_tiene_modulo_franquicia = ModuloContratado(empresa_sup,ModFranquiciasTiendas)
                    if si_tiene_modulo_franquicia and esFranquiciador then
                        'seleccion = "select empresa_sup from clientes with(nolock) where ncliente = '"&session("ncliente")&"'"
			            'rstSelect.Open seleccion, DsnIlion, adOpenKeyset, adLockOptimistic
			            'CIF_Franq = d_lookup("cifedi","clientes","ncliente='" & session("ncliente") & "'",DsnIlion)
                        CIF_Franq_select = "select cifedi from clientes with(nolock) where ncliente= ?"
                        CIF_Franq = DLookupP1(CIF_Franq_select, session("ncliente") &"", adChar, 5, DsnIlion)

			            'clienteEnFranq = d_lookup("ncliente","clientes","cifedi='" & CIF_Franq & "'",dsnCentral)
                        clienteEnFranqSelect = "select ncliente from clientes with(nolock) where cifedi= ?"
			            clienteEnFranq= DLookupP1(clienteEnFranqSelect, CIF_Franq & "", adVarchar, 20, dsnCentral)

			            'cantPed=d_lookup("count(*)", "pedidos_cli", "su_npedido='"& trimCodEmpresa(rst("npedido")) &"' and ncliente='"& clienteEnFranq &"' ", dsnCentral)
			            cantPedSelect = "select count(*) from pedidos_cli with(nolock) where su_npedido= ? and ncliente= ?"
                        cantPed= DLookupP2(cantPedSelect, trimCodEmpresa(rst("npedido")) &"", adVarchar, 20, clienteEnFranq &"", adChar, 10, dsnCentral)

                        if cint(cantPed) = 0 then
                            %><span align="center" style="border-width: 1px; border-left-style:solid; border-color:#c0c0c0;"><a class="CELDAREFB" href="javascript:CrearPedidoCentral('<%=enc.EncodeForJavascript(npedido)%>');">&nbsp;&nbsp;&nbsp;<%=LITCREAR_PED_CENTRAL%>&nbsp;&nbsp;&nbsp;</a></span><%
                        end if
                    end if
			    end if
			end if
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
			'rstSelect.Close
			'seleccion = "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDO A PROVEEDOR' order by descripcion"
            'rstSelect.cursorlocation=3
			'rstSelect.Open seleccion, DsnIlion
	  		
            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open DsnIlion
            connAux.cursorlocation=3
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText= "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente=? and b.tippdoc='PEDIDO A PROVEEDOR' order by descripcion"                
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@ncliente", adChar,adParamInput,5, session("ncliente"))

            set rstSelect = commandAux.Execute

            'CloseDiv
            DrawDiv "header-print","",""
			if rpc & "">"" then
			    %><label><a id="idPrintFormat" class="CELDAREFB" class='CELDAREFB' href='javascript:ComprobarPvp(1,"<%=enc.EncodeForJavascript(npedido)%>","<%=session("ncliente")%>","<%=session("usuario")%>","<%=enc.EncodeForJavascript(novei)%>","","<%=enc.EncodeForJavascript(trimCodEmpresa(rst("nproveedor")))%>");' OnMouseOver="self.status='<%=LitImpresionConFormato%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitImpresionConFormato%></a></label><%
			else 
			    %><label><a id="idPrintFormat" class="CELDAREFB" href="javascript:AbrirVentana(document.pedidos_pro.formato_impresion.value+'npedido=<%="(\'"+enc.EncodeForJavascript(npedido)+"\')"%>&mode=browse&empresa=<%=session("ncliente")%>&novei=<%=enc.EncodeForJavascript(novei)%>&usuario=<%=session("usuario")%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitImpresionConFormato%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitImpresionConFormato%></a></label><%
			end if
			'DrawSelectCelda "CELDARIGHT","120","",0,"","formato_impresion",rstSelect,defecto,"fichero","descripcion","",""%>
			    <select class='CELDA' width='150px' style='width:150px' name="formato_impresion">
			        <%encontrado=0
				    while not rstSelect.eof
					    if defecto=rstSelect("descripcion") then
						    encontrado=1
						    if isnull(rstSelect("parametros")) then
							    prm=""
						    else
							    prm=null_s(rstSelect("parametros")) & "&"
						    end if%>
						    <option selected value="<%=EncodeForHtml(null_s(rstSelect("fichero"))) & "?" & EncodeForHtml(prm)%>"><%=EncodeForHtml(null_s(rstSelect("descripcion")))%></option>
					    <%else
						    if isnull(rstSelect("parametros")) then
							    prm=""
						    else
							    prm=null_s(rstSelect("parametros")) & "&"
						    end if%>
						    <option value="<%=EncodeForHtml(null_s(rstSelect("fichero")))  & "?" & EncodeForHtml(prm)%>"><%=EncodeForHtml(null_s(rstSelect("descripcion")))%></option>
					    <%end if
					    rstSelect.movenext
				    wend%>
			    </select>
		    <%
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
            'rstSelect.close
			if not rst.eof then
				pagina="../crearpdf.asp?destinatario=" & enc.EncodeForJavascript(null_s(rst("nproveedor"))) & "&ndoc=" & enc.EncodeForJavascript(null_s(rst("npedido"))) & "&tdoc=PEDIDO&dedonde=DOCUMENTOC&empresa=" & session("ncliente") & "&mode=DOC&url=compras/"
                if rpc & "">"" then
			        %><a id="idPrintFormat" class="CELDAREFB" href='javascript:ComprobarPvp(2,"<%=enc.EncodeForJavascript(npedido)%>","<%=session("ncliente")%>","<%=session("usuario")%>","<%=enc.EncodeForJavascript(novei)%>","<%=pagina%>","<%=enc.EncodeForJavascript(trimCodEmpresa(rst("nproveedor")))%>");' onmouseover="self.status='<%=LitEnvMail%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=themeIlion %><%=ImgEnviarEmail%>" <%=ParamImgEnviarEmail%> alt="<%=ucase(LitEnvMail)%>" title="<%=ucase(LitEnvMail)%>"/></a><%
			    else
				    %><a id="idPrintFormat" class="CELDAREFB" href="javascript:AbrirVentana('<%=pagina%>' + document.pedidos_pro.formato_impresion.value + 'npedido=<%="(\'"+enc.EncodeForJavascript(npedido)+"\')"%>','A','<%=AltoVentana+100%>','950')" onmouseover="self.status='<%=LitEnvMail%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=themeIlion %><%=ImgEnviarEmail%>" <%=ParamImgEnviarEmail%> alt="<%=ucase(LitEnvMail)%>" title="<%=ucase(LitEnvMail)%>"/></a><%
				end if
			end if
	        CloseDiv
			if session("version")&"" = "5" then
                DrawDiv "","","" 
                CloseDiv
            end if 
		end if%>
        </div><%
    if mode="browse" then
        BarraOpciones "browse", rst("npedido")
    end if
    ActionVersion Altoventana, AnchoVentana
    %><table class ="width100"></table><%

	    Alarma "pedidos_pro.asp"%>
      
    <%if (mode="browse" or mode="edit" or mode="add") and not rst.EOF then 
    'rstAux.cursorlocation=3
	'rstAux.open "select npedido,referencia,item from detalles_ped_pro with(nolock) where npedido='" & rst("npedido") & "' and mainitem is null and cantidadpend<>cantidad",session("dsn_cliente")
	set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux =  Server.CreateObject("ADODB.Command")

    connAux.open session("dsn_cliente")
    connAux.cursorlocation=3
    commandAux.ActiveConnection =connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText= "select npedido,referencia,item from detalles_ped_pro with(nolock) where npedido=? and mainitem is null and cantidadpend<>cantidad"
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, rst("npedido")&"")

    set rstAux = commandAux.Execute
    
    if rstAux.eof then%>
	    <input type="hidden" name="borrarpedido" value="SI"/>
	<%else%>
		<input type="hidden" name="borrarpedido" value="NO"/>
	<%end if
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	'rstAux.close%>
	<input type="hidden" name="h_nproveedor" value="<%=EncodeForHtml(null_s(rst("nproveedor")))%>"/>
	<input type="hidden" name="h_npedido" value="<%=EncodeForHtml(null_s(rst("npedido")))%>"/>
	<input type="hidden" name="h_formapago" value="<%=EncodeForHtml(null_s(rst("forma_pago")))%>"/>
    <!-- FALLA con enc.EncodeForHtmlAttribute (rst("forma_pago"))-->
	<%if not isnull(rst("nalbaran")) then%>
		<input type="hidden" name="h_nalbaran" value="<%=EncodeForHtml(null_s(rst("nalbaran")))%>"/>
		<%
            'nalbaran_pro=d_lookup("nalbaran_pro","albaranes_pro","nalbaran='" & rst("nalbaran") & "'",session("dsn_cliente"))
            nalbaran_pro_select= "select nalbaran_pro from albaranes_pro with(nolock) where nalbaran= ?"
            nalbaran_pro=DLookupP1(nalbaran_pro_select, rst("nalbaran") & "", adVarchar, 20, session("dsn_cliente"))
        %>
		<input type="hidden" name="h_nalbaranpro" value="<%=EncodeForHtml(null_s(nalbaran_pro))%>"/>
	<%else%>
		<input type="hidden" name="h_nalbaran" value="NO"/>
		<input type="hidden" name="h_nalbaranpro" value="NO"/>
	<%end if%>
    <input type="hidden" name="olddivisa" value="<%=EncodeForHtml(null_s(rst("divisa")))%>"/>
    <%'enc.EncodeForHtmlAttribute(rst("divisa falla")) %>
        
        <!--<div id="CollapseSection">
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['GENERAL_DATA', 'CABECERA', 'DIRENVIO', 'FINANCIAL_DATA','TOTAL']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title=""/></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['GENERAL_DATA', 'CABECERA', 'DIRENVIO', 'FINANCIAL_DATA','TOTAL']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title=""/></a>
        </div>-->
		<div class="Section" id="S_GENERAL_DATA">
            <a href="#" rel="toggle[GENERAL_DATA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitCabecera%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                </div>
            </a>
            <div class="SectionPanel" id="GENERAL_DATA">
                <table width="100%" bgcolor="<%=color_blau %>" border="0"><%
                    'DrawFila color_blau
                        if mode="browse" then
				            'DrawCelda "ENCABEZADOL style='width:130px'","","",0,LitSerie + " :"
				            'DrawCelda "CELDA style='width:200px'","","",0,trimCodEmpresa(rst("serie"))
                            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, LitSerie, "serie", "",EncodeForHtml(trimCodEmpresa(rst("serie")))
				            'mmg:
	                        set rstMM = Server.CreateObject("ADODB.Recordset")
                            'rstMM.cursorlocation=3
	                        'rstMM.open "select almacen from series with(nolock), almacenes alm with(nolock) where nserie='"&rst("serie")&"' and alm.codigo=almacen and isnull(alm.fbaja,'')=''"& strwhereMM,session("dsn_cliente")
                            
                            set commandAux = nothing
                            set connAux = Server.CreateObject("ADODB.Connection")
                            set commandAux = Server.CreateObject("ADODB.Command")

                            connAux.open session("dsn_cliente")
                            connAux.cursorlocation=3
                            commandAux.ActiveConnection =connAux
                            commandAux.CommandTimeout = 60
                            commandAux.CommandText= "select almacen from series with(nolock), almacenes alm with(nolock) where nserie=? and alm.codigo=almacen and isnull(alm.fbaja,'')=''"& strwhereMM
                            commandAux.CommandType = adCmdText
                            commandAux.Parameters.Append commandAux.CreateParameter("@nserie", adVarChar,adParamInput,10, rst("serie")&"")

                            set rstMM = commandAux.Execute

                            if not rstMM.EOF then
		                        almacenSerie= rstMM("almacen")
	                        else
		                        almacenSerie= ""
	                        end if
                            connAux.Close
                            set connAux = nothing
                            set commandAux = nothing
	                        'rstMM.close
			            else
                            
				            strSacSerie="select nserie, case when datalength(right(nserie,len(nserie)-5)++' '+nombre)<=21 then right(nserie,len(nserie)-5)++'-'+nombre else left(right(nserie,len(nserie)-5)++'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='PEDIDO A PROVEEDOR' and nserie like ?+'%'"
				            if s & "">"" then
					            strSacSerie=strSacSerie & " and nserie in " & s
							end if
				            strSacSerie=strSacSerie & " order by nserie"
                            ' <<< MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a documentos de compras.
                            'rstAux.cursorlocation=3
				            'rstAux.open strSacSerie,session("dsn_cliente")
                            
                            set commandAux = nothing
                            set connAux = Server.CreateObject("ADODB.Connection")
                            set commandAux = Server.CreateObject("ADODB.Command")

                            connAux.open session("dsn_cliente")
                            connAux.cursorlocation=3
                            commandAux.ActiveConnection =connAux
                            commandAux.CommandTimeout = 60
                            commandAux.CommandText= strSacSerie
                            commandAux.CommandType = adCmdText
                            commandAux.Parameters.Append commandAux.CreateParameter("@nserie", adVarChar,adParamInput,10, session("ncliente")&"")

                            set rstAux = commandAux.Execute

				            if mode="add" then
                                 DrawSelectCelda "CELDA","","",0,LitSerie,"serie",rstAux,iif(tmp_serie>"",tmp_serie,rst("serie")),"nserie","descripcion","onchange","javascript:TraerProveedor('add','2');"
                            else
                                 DrawSelectCelda "CELDA","","",0,LitSerie,"serie",rstAux,iif(tmp_serie>"",tmp_serie,rst("serie")),"nserie","descripcion","",""
				            end if
				            connAux.Close
                            set connAux = nothing
                            set commandAux = nothing
                            'rstAux.close  
                        '------------------------------------------------------------
			            end if
                        if mode="add" or mode="edit" then
                            EligeCelda "check", mode,"CELDA","","",0,LitValorado,"valorado",0,EncodeForHtml(iif(tmp_valorado>"",nz_b(tmp_valorado),rst("valorado")))
                        else 
                            DrawDiv "1","",""
                            DrawLabel "","",LitValorado
                            EligeCeldaResponsive1 "check", mode, "CELDA","","valorado", EncodeForHtml(iif(tmp_valorado>"",nz_b(tmp_valorado),rst("valorado"))), ""
                            CloseDiv
                        end if
                        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:130px'","","",0,LitValorado+":"
			            
                    'CloseFila
                    'DrawFila ""'color_blau
			        '*** AMP 20102010	
			            campo="codigo"
				        if mode="browse" then
					        campo2="abreviatura"
				        else
					        campo2="abreviatura"
				        end if
						
				        if tmpdivisafc>"" then  tmp_divisa = tmpdivisafc end if
				        DIVISA=iif(tmp_divisa>"",tmp_divisa,rst("divisa"))
				        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA"),"","",0,"<nobr>"+LitDivisa+" :</nobr>"
				        dato_celda=Desplegable(mode,campo,campo2,"divisas",DIVISA,"moneda_base<>0 and codigo like '" & session("ncliente") & "%'")
''response.write("los datos 1 son-" & mode & "-" & tmpdivisafc & "-" & DIVISA & "-" & tmp_divisa & "-" & rst("divisa") & "-<br>")
				        if mode<>"browse" then
                            datoDivisaSelect = "select abreviatura from divisas with(nolock) where codigo= ? and codigo like ?+'%'"
					        ' TRUE d_lookup("abreviatura","divisas","codigo='" & tmp_divisa & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                            ' FALSE d_lookup("abreviatura","divisas","codigo='" & dato_celda & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")))
                            datoDivisa=iif(tmp_divisa>"", _
						        DLookupP2(datoDivisaSelect, tmp_divisa & "", adVarchar, 15, session("ncliente") & "", adVarchar, 15, session("dsn_cliente")), _
						        DLookupP2(datoDivisaSelect, dato_celda & "", adVarchar, 15, session("ncliente") & "", adVarchar, 15, session("dsn_cliente")))
				        else
					        datoDivisa=dato_celda
				        end if
''response.write("los datos 2 son-" & mode & "-" & datoDivisa & "-" & tmp_divisa & "-" & rst("divisa") & "-<br>")
				
				        if mode<>"edit" then
	                        estilo_divisa="CELDA"
	                        if mode="add" then
		                        tipo_eligecelda="select"
	                        else
		                        tipo_eligecelda="input"
	                        end if
                        else
	                        cuantos_detalles=d_count("item","detalles_ped_pro","npedido='" & rst("npedido") & "'",session("dsn_cliente"))
	                        cuantos_conceptos=d_count("nconcepto","conceptos_ped_pro","npedido='" & rst("npedido") & "'",session("dsn_cliente"))
	                        if cint(cuantos_detalles) + cint(cuantos_conceptos)>0 then
		                        estilo_divisa="CELDA DISABLED width60"
		                        tipo_eligecelda="input"
	                        else
		                        estilo_divisa="CELDA width60"
		                        tipo_eligecelda="select"
	                        end if
                        end if
                    
                        if mode="add" or mode="edit" then RstAux.close
				        Estilo=iif(mode="browse","CELDA width60",estilo_divisa)
				        if tipo_eligecelda="input" then
                            if mode="browse" then
                                'd_lookup("abreviatura","divisas","codigo='" & rst("divisa") & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                                divisaSelect = "select abreviatura from divisas with(nolock) where codigo= ? and codigo like ?+'%'"
                                EligeCeldaResponsive "text",mode,"CELDA","","",0,LitDivisa,"divisa",5,EncodeForHtml(iif(mode="add",datoDivisa, DLookupP2(divisaSelect, rst("divisa") & "", adVarchar, 15, session("ncliente") & "", adVarchar, 15, session("dsn_cliente"))))
                            else
                                'd_lookup("abreviatura","divisas","codigo='" & rst("divisa") & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                                divisaSelect = "select abreviatura from divisas with(nolock) where codigo= ? and codigo like ?+'%'"    
                                DrawInputCeldaDisabled "","","",5,0,LitDivisa,"divisa",EncodeForHtml(iif(mode="add",datoDivisa, DLookupP2(divisaSelect, rst("divisa") & "", adVarchar, 15, session("ncliente") & "", adVarchar, 15, session("dsn_cliente"))))
                            end if
                'DrawInputCelda (estilo,ancho,alto,nchar,tabulacion,etiqueta,name,dato)
                        else
				            'monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))	
                            monedaBaseSelect = "select codigo from divisas with(nolock) where codigo like ?+'%' and moneda_base='1'"
                            monedaBase = DLookupP1(monedaBaseSelect, session("ncliente") & "", adVarchar, 15, session("dsn_cliente"))	
                            'rstAux.cursorlocation=3
					        'rstAux.open "select codigo,abreviatura as descripcion from divisas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
                     
                            set commandAux = nothing
                            set connAux = Server.CreateObject("ADODB.Connection")
                            set commandAux =  Server.CreateObject("ADODB.Command")

                            connAux.open session("dsn_cliente")
                            connAux.cursorlocation=3
                            commandAux.ActiveConnection =connAux
                            commandAux.CommandTimeout = 60
                            commandAux.CommandText= "select codigo,abreviatura as descripcion from divisas with(nolock) where codigo like ?+'%' order by descripcion"
                            commandAux.CommandType = adCmdText
                            commandAux.Parameters.Append commandAux.CreateParameter("@codigo", adVarChar,adParamInput,15, session("ncliente")&"")

                            set rstAux = commandAux.Execute

                            if mode="add" then
		 			            DrawSelectCelda "CELDA","","",0,LitDivisa,"divisa",rstAux,iif(mode<>"browse",iif(tmp_divisa>"",tmp_divisa,dato_celda),rst("divisa")),"codigo","descripcion","onchange","javascript:cambiardivisa('"&monedaBase&"');"
                            else
                                'd_lookup("abreviatura","divisas","codigo='" & rst("divisa") & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")
                                divisaSelect = "select abreviatura from divisas with(nolock) where codigo= ? and codigo like ?+'%'"
                                DrawInputCeldaDisabled "","","",5,0,LitDivisa,"divisa", EncodeForHtml(DLookupP2(divisaSelect, rst("divisa") & "", adVarchar, 15,  session("ncliente") & "", adVarchar, 15, session("dsn_cliente")))
                            end if
                            connAux.Close
                            set connAux = nothing
                            set commandAux = nothing
                            'rstAux.close
				        end if
''response.write("los datos 3 son-" & mode & "-" & datoDivisa & "-" & tmp_divisa & "-" & rst("divisa") & "-<br>")
				        %><input type="hidden" name="h_divisa" value="<%=EncodeForHtml(iif(tmp_divisa>"",tmp_divisa,dato_celda))%>"/>
				        <!--<input type="hidden" name="divisafc" value="<%=iif(tmp_divisa>"",tmp_divisa,DIVISA)%>"/>-->
                        <!-- FALLA con el enc.EncodeForHtmlAttribute <input type="hidden" name="divisafc" value="<%=EncodeForHtml(iif(mode="add",iif(tmp_divisa>"",tmp_divisa,DIVISA), rst("divisa")))%>"/> -->
                        <input type="hidden" name="divisafc" value="<%=EncodeForHtml(null_s(iif(mode="add",iif(tmp_divisa>"",tmp_divisa,DIVISA), rst("divisa"))))%>"/><%
				        'n_decimales=d_lookup("ndecimales","divisas","codigo='" & iif(mode="browse",DIVISA,iif(tmp_divisa>"",tmp_divisa,dato_celda)) & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                        n_decimales_select= "select ndecimales from divisas with(nolock) where codigo= ? and codigo like ?+'%'"
                        n_decimales= DLookupP2(n_decimales_select, iif(mode="browse",DIVISA,iif(tmp_divisa>"",tmp_divisa,dato_celda)) & "", adVarchar, 15, session("ncliente") & "", adVarchar, 15, session("dsn_cliente"))
		            
		
		            '*** i AMP 04102010 : Incorporamos campo factor de cambio.
                    'Información sobre la moneda base de la empresa.
                    'monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
                    monedaBaseSelect = "select codigo from divisas with(nolock) where codigo like ?+'%' and moneda_base='1'"
                    monedaBase = DLookupP1(monedaBaseSelect, session("ncliente") & "", adVarchar, 15, session("dsn_cliente"))
   	                'abrevBase =  d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
                    abrevBaseSelect = "select abreviatura from divisas with(nolock) where codigo like ?+'%' and moneda_base='1'"
                    abrevBase = DLookupP1(abrevBaseSelect, session("ncliente") & "", adVarchar, 15, session("dsn_cliente"))
                    'factcambio = d_lookup("factcambio","pedidos_pro","npedido='" & rst("npedido") & "' and npedido like '" & session("ncliente") & "%'",session("dsn_cliente"))
                    factcambioSelect = "select factcambio from pedidos_pro with(nolock) where npedido= ? and npedido like ?+'%'"
                    factcambio = DLookupP2(factcambioSelect, rst("npedido") & "", adVarchar, 20, session("ncliente") & "", adVarchar, 20, session("dsn_cliente"))
''response.write("los datos son-" & mode & "-" & DIVISA & "-" & monedaBase & "-" & dv & "-<br>")
	                'DrawFila "" 'color_blau
	                        if mode="browse" then
	                            ''if  DIVISA<>monedaBase then
                                if  rst("divisa")<>monedaBase then
                                    'abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&rst("divisa")&"'",session("dsn_cliente"))
	                                abreviaAtDivSelect = "select abreviatura from divisas with(nolock) where codigo like ?+'%' and codigo= ?"
                                    abreviaAtDiv = DLookupP2(abreviaAtDivSelect, session("ncliente") & "", adVarchar, 15, rst("divisa")&"", adVarchar, 15, session("dsn_cliente"))
                                    'DrawCelda "ENCABEZADOL style='width:130px'","","",0,LitFactCambio+" :"
	                                ''DrawCelda "CELDA style='width:200px'","","",0,CStr(factcambio)+" "+abrevBase
                                    'DrawCelda "CELDA style='width:200px'","","",0,"1" & abrevBase & " = " & CStr(factcambio) & abreviaAtDiv
                                    EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "", LitFactCambio, "",EncodeForHtml("1" & abrevBase & " = " & CStr(factcambio) & abreviaAtDiv)
	                            end if
	                        else
	                            ocultar=0  
	                            if mode="add" or mode="edit" then	  
              	                    if mode="add" then
              	                        dv=iif(DIVISA>"",DIVISA,dato_celda)              	         
              	                        'factcambio = d_lookup("factcambio","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dv&"'",session("dsn_cliente"))
                                        factcambioSelect = "select factcambio from divisas with(nolock) where codigo like ?+'%' and codigo= ?"
                                        factcambio = DLookupP2(factcambioSelect, session("ncliente") & "", adVarchar, 15, dv &"", adVarchar, 15, session("dsn_cliente"))
                                        'abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dv&"'",session("dsn_cliente"))
                                        abreviaAtDivSelect = "select abreviatura from divisas with(nolock) where codigo like ?+'%' and codigo= ?"
                                        abreviaAtDiv = DLookupP2(abreviaAtDivSelect, session("ncliente") & "", adVarchar, 15, dv&"", adVarchar, 15, session("dsn_cliente"))
              	                        if dv=monedaBase then ocultar=1 end if
              	                    else 'modo edit
                                        dvEdit=iif(tmp_divisa>"",tmp_divisa,dato_celda)
                                        'abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dvEdit&"'",session("dsn_cliente"))
              	                        abreviaAtDivSelect = "select abreviatura from divisas with(nolock) where codigo like ?+'%' and codigo= ?"
                                        abreviaAtDiv = DLookupP2(abreviaAtDivSelect, session("ncliente") & "", adVarchar, 15, dvEdit&"", adVarchar, 15, session("dsn_cliente"))
                                        if dvEdit=monedaBase then ocultar=1 end if              	        
              	                    end if

                                DrawDiv "1", "", "tdfactcambio"
                                DrawLabel "", "", LitFactCambio
                                DrawSpan "CELDA", "", "1" & EncodeForHtml(abrevBase) & " = ", ""
                                %>
                                    <input type="text" name="nfactcambio" value="<%=EncodeForHtml(null_s(CStr(factcambio)))%>" size="6" style="text-align:right" onchange="comprobarFactorCambio()"/>
                                    <span class="CELDA" id="idfactcambioexpl"><%=abreviaAtDiv%></span>
                                <%
                                CloseDiv

	                            if ocultar=1 then
                                    %><script language="javascript" type="text/javascript">
                                        parent.pantalla.document.getElementById("tdfactcambio").style.display = "none";
                                    </script><% 
                                end if
	                        end if	     
	                    end if
	                'CloseFila 
                    'DrawFila ""' color_blau
                        'RGU he copiado la fila que pone decimales que antes estaba mas abajo
			             'n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))
                         n_decimales_select = "select ndecimales from divisas with(nolock) where codigo= ?"
                         n_decimales = DLookupP1(n_decimales_select, rst("divisa") & "", adVarchar, 15, session("dsn_cliente"))
			             if n_decimales = "" then
			                n_decimales = 0
  			             end if

				            '***RGU 19/12/2005 ***
			            %><input type="hidden" name="si_tiene_modulo_ebesa" value="<%=si_tiene_modulo_ebesa%>"/><%
			            'pintamos el límite alcanzado si es
			                fecha=rst("salida")&""
				            if fecha="" then
					            fecha=rst("fecha")
				            end if
				            alcanzado=gestionalcanzado(mode, fecha,"")
				            %><input type="hidden" name="alcanzado" ID='alcanzado' value="<%=EncodeForHtml(alcanzado)%>"/><%
				           'a significa que no ha hay en limite_compras datos para ese mes y año
				            if alcanzado<>"a" then
				             if mode="browse" then ' marca 
                                EligeCeldaResponsive "text", mode,"CELDA","","",0,totalcanzado,LitAlcan,0,EncodeForHtml(formatnumber(null_z(alcanzado),n_decimales,-1,0,-1))

				            end if
					            'alcanzado=gestionalcanzado("add",fecha,1000000)
					            'guardamos los valores en campos ocultos
					            mes=month(fecha)
					            anyo=year(fecha)

					            'el mes tiene que tener dos dígitos porqué en la bd tiene dos dígitos sino luego en la select no lo devolverá
					            if len(mes)=1 then
					              mes="0" & mes
					            end if
					            'JCI
                                'rstLimite.cursorlocation=3
					            'rstLimite.open "select lim.importe as importe from limite_compras as lim with(nolock) where lim.nempresa='" & session("ncliente") & "' and lim.mes='" & mes & "' and lim.anyo='" & anyo & "'", session("dsn_cliente")
					            
                                set commandAux = nothing
                                set connAux = Server.CreateObject("ADODB.Connection")
                                set commandAux =  Server.CreateObject("ADODB.Command")
                            
                                connAux.open session("dsn_cliente")
                                connAux.cursorlocation=3
                                commandAux.ActiveConnection =connAux
                                commandAux.CommandTimeout = 60
                                commandAux.CommandText= "select lim.importe as importe from limite_compras as lim with(nolock) where lim.nempresa=? and lim.mes=? and lim.anyo=?"
                                commandAux.CommandType = adCmdText
                                commandAux.Parameters.Append commandAux.CreateParameter("@nempresa", adVarChar,adParamInput,5, session("ncliente")&"")
                                commandAux.Parameters.Append commandAux.CreateParameter("@mes", adVarChar,adParamInput,2, mes)
                                commandAux.Parameters.Append commandAux.CreateParameter("@anyo", adVarChar,adParamInput,4, anyo)
                                
                                set rstLimite = commandAux.Execute
                                
                                if not rstLimite.eof then
		                           'preguntar_riesgo_conf=nz_b(d_lookup("gestionlimitecompras","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))
                                   preguntar_riesgo_conf_select = "select gestionlimitecompras from configuracion with(nolock) where nempresa= ?"
		                           preguntar_riesgo_conf=nz_b(DLookupP1(preguntar_riesgo_conf_select, session("ncliente") & "", adChar, 5, session("dsn_cliente")))
		                           'contrasenya_riesgo_conf=null_s(d_lookup("pwdlimitecompras","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))
		                           contrasenya_riesgo_conf_select = "select pwdlimitecompras from configuracion with(nolock) where nempresa= ?"
                                   contrasenya_riesgo_conf=null_s(DLookupP1(contrasenya_riesgo_conf_select, session("ncliente") & "", adChar, 5, session("dsn_cliente")))

                                        %><input type="hidden" name="limite" value="<%=EncodeForHtml(null_s(rstLimite("importe")))%>"/>
		                               <input type="hidden" name="si_preguntar_riesgo" value="<%=EncodeForHtml(null_s(preguntar_riesgo_conf))%>"/>
		                               <input type="hidden" name="contrasenya" value="<%=EncodeForHtml(null_s(contrasenya_riesgo_conf))%>"/>
						               <input type="hidden" name="mes" value="<%=EncodeForHtml(null_s(mes))%>"/>
						               <input type="hidden" name="anyo" value="<%=EncodeForHtml(null_s(anyo))%>"/><%
                                end if
                                connAux.Close
                                set connAux = nothing
                                set commandAux = nothing
					            'rstLimite.close
				            '***RGU 19/12/2005 ***
				        end if
                        '''MPC 27/04/2007
				        if rt = "si" and mode <> "add" then
				            %><td class="CELDA"><a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=compras/reparto_pedido.asp&pag2=compras/reparto_pedido_bt.asp&mode=browse&ncliente=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>&ndocumento=<%=enc.EncodeForJavascript(null_s(rst("nproveedor")))%>', 'P', <%=AltoVentana%>, <%=AnchoVentana%>)"><%=LitRepartoPed%></a></td><%
				        end if
				        if esto_se_oculta_ya_que_no_funciona=1 and mode <> "add" then
				            %><td class="CELDA"><a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=compras/repetir_pedido.asp&pag2=compras/repetir_pedido_bt.asp&mode=browse&ncliente=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>&ndocumento=<%=enc.EncodeForJavascript(null_s(rst("serie")))%>&nproveedor=<%=enc.EncodeForJavascript(null_s(rst("nproveedor")))%>', 'P', <%=AltoVentana%>, <%=AnchoVentana%>)"><%=LitRepetirPed%></a></td><%
				        end if
   		            'CloseFila
                    DrawDiv "3-sub","background-color: #eae7e3",""
                        %><label class="ENCABEZADOL", style="text-align:left"><%=LIT_GENERAL_DATA%></label><%
                    CloseDiv
			            'rgu he quitado el codigo que coge el numero de decimales y lo he puesto mas arriba
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:170px'","","",0,LitFormaPago+":"
			            defecto=iif(tmp_forma_pago>"",tmp_forma_pago,Nulear(rst("forma_pago")))
			            if mode <> "browse" then
                            'rstAux.cursorlocation=3
				            'rstAux.open "select codigo,descripcion from formas_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
	 			            
                            set commandAux = nothing
                            set connAux = Server.CreateObject("ADODB.Connection")
                            set commandAux =  Server.CreateObject("ADODB.Command")

                            connAux.open session("dsn_cliente")
                            connAux.cursorlocation=3
                            commandAux.ActiveConnection =connAux
                            commandAux.CommandTimeout = 60
                            commandAux.CommandText= "select codigo,descripcion from formas_pago with(nolock) where codigo like ?+'%' order by descripcion"
                            commandAux.CommandType = adCmdText
                            commandAux.Parameters.Append commandAux.CreateParameter("@codigo", adVarChar,adParamInput,10, session("ncliente")&"")

                            set rstAux = commandAux.Execute
                            
                            'DrawSelectCelda "CELDA",iif(mode<>"browse","200",""),"",0,"","formas_pago",rstAux,defecto,"codigo","descripcion","onchange","javascript:CalculaFechaPago();"
	 			            DrawSelectCelda "CELDA",iif(mode<>"browse","200",""),"",0,LitFormaPago,"formas_pago",rstAux,defecto,"codigo","descripcion","onchange","javascript:CalculaFechaPago();"
                            connAux.Close
                            set connAux = nothing
                            set commandAux = nothing
                            'rstAux.close

                            'rstAux.cursorlocation=3
	 			            'rstAux.open "select codigo,dias from formas_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by dias",session("dsn_cliente")
                            
                            set commandAux = nothing
                            set connAux = Server.CreateObject("ADODB.Connection")
                            set commandAux =  Server.CreateObject("ADODB.Command")

                            connAux.open session("dsn_cliente")
                            connAux.cursorlocation=3
                            commandAux.ActiveConnection =connAux
                            commandAux.CommandTimeout = 60
                            commandAux.CommandText= "select codigo,dias from formas_pago with(nolock) where codigo like ?+'%' order by dias"
                            commandAux.CommandType = adCmdText
                            commandAux.Parameters.Append commandAux.CreateParameter("@codigo", adVarChar,adParamInput,10, session("ncliente")&"")

                            set rstAux = commandAux.Execute
                            
                            DrawSelect1 "width60","display:none","","formas_pagodias",rstAux,defecto,"codigo","dias","","","",""
	 			            'DrawSelectCelda "CELDA style='display:none'",iif(mode<>"browse","200",""),"",0,"","formas_pagodias",rstAux,defecto,"codigo","dias","",""
                            connAux.Close
                            set connAux = nothing
                            set commandAux = nothing
                            'rstAux.close
				            %><input type="hidden" name="h_formas_pago" value="<%=EncodeForHtml(null_s(defecto))%>"/><%
			            else
				            'DrawCelda "CELDA","","",0,d_lookup("descripcion","formas_pago","codigo='" & defecto & "'",session("dsn_cliente"))
                            'EligeCeldaResponsive "text",mode,"CELDA","","",0,LitFormaPago,LitFormaPago,15,d_lookup("descripcion","formas_pago","codigo='" & defecto & "'",session("dsn_cliente"))
                            descripFormaPagoSelect = "select descripcion from formas_pago with(nolock) where codigo= ?"
                            EligeCeldaResponsive "text",mode,"CELDA","","",0,LitFormaPago,LitFormaPago,15, EncodeForHtml(DLookupP1(descripFormaPagoSelect, defecto & "", adVarchar, 10, session("dsn_cliente")))
			            end if
			            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"  "
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:180px'","","",0,LitTipoPago+":"
			            defecto=iif(tmp_tipo_pago>"",tmp_tipo_pago,Nulear(rst("tipo_pago")))
			            if mode <> "browse" then
                            'rstAux.cursorlocation=3
				            'rstAux.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
	 			            
                            set commandAux = nothing
                            set connAux = Server.CreateObject("ADODB.Connection")
                            set commandAux =  Server.CreateObject("ADODB.Command")

                            connAux.open session("dsn_cliente")
                            connAux.cursorlocation=3
                            commandAux.ActiveConnection =connAux
                            commandAux.CommandTimeout = 60
                            commandAux.CommandText= "select codigo,descripcion from tipo_pago with(nolock) where codigo like ?+'%' order by descripcion"
                            commandAux.CommandType = adCmdText
                            commandAux.Parameters.Append commandAux.CreateParameter("@codigo", adVarChar,adParamInput,10, session("ncliente")&"")

                            set rstAux = commandAux.Execute

                            DrawSelectCelda "CELDA",iif(mode<>"browse","200",""),"",0,LitTipoPago+"","tipo_pago",rstAux,defecto,"codigo","descripcion","",""
                            connAux.Close
                            set connAux = nothing
                            set commandAux = nothing
                            'rstAux.close
			            else
                            'EligeCeldaResponsive "text",mode,"CELDA","","",0,LitTipoPago,LitTipoPago,15,d_lookup("descripcion","tipo_pago","codigo='" & defecto & "'",session("dsn_cliente"))
			                tipoPagoSelect = "select descripcion from tipo_pago with(nolock) where codigo= ?"
                            EligeCeldaResponsive "text",mode,"CELDA","","",0,LitTipoPago,LitTipoPago,15,EncodeForHtml(DLookupP1(tipoPagoSelect, defecto & "", adVarchar, 8, session("dsn_cliente")))
                        end if
		            'CloseFila

		            'DrawFila "" 'color_blau
			                'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:170px'","","",0,LitNumAlbaran+":"
			                if mode="browse" then
			 	                'EligeCeldaResponsive "text", mode,"CELDA","","",0,LitNumAlbaran,LitNumAlbaran,0,d_lookup("nalbaran_pro","albaranes_pro","nalbaran='" & iif(tmp_nalbaran>"",tmp_nalbaran,iif(isnull(rst("nalbaran")),"",rst("nalbaran"))) & "'",session("dsn_cliente"))
			                    nalbaranSelect = "select nalbaran_pro from albaranes_pro with(nolock) where nalbaran= ?"
			 	                EligeCeldaResponsive "text", mode,"CELDA","","",0,LitNumAlbaran,LitNumAlbaran,0,EncodeForHtml(DLookupP1(nalbaranSelect, iif(tmp_nalbaran>"",tmp_nalbaran,iif(isnull(rst("nalbaran")),"",rst("nalbaran"))) & "", adVarchar, 20, session("dsn_cliente")))
                            else
			 	            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                %><label><%=LitNumAlbaran%></label><%
                                    'd_lookup("nalbaran_pro","albaranes_pro","nalbaran='" & iif(tmp_nalbaran>"",tmp_nalbaran,iif(isnull(rst("nalbaran")),"",rst("nalbaran"))) & "'",session("dsn_cliente"))
                                    nalbaranProSelect = "select nalbaran_pro from albaranes_pro with(nolock) where nalbaran= ?"
					            %><input class='width60' type="text" name="nalbaran" value="<%=EncodeForHtml(DLookupP1(nalbaranProSelect, iif(tmp_nalbaran>"",tmp_nalbaran,iif(isnull(rst("nalbaran")),"",rst("nalbaran"))) & "", adVarchar, 20, session("dsn_cliente")))%>" disabled="disabled"/><%
				            %></div><%
                            end if
			                'if mode<>"browse" then DrawCelda "CELDA","10","",0,"  "
			                'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:180px'","","",0,LitNumFactura+":"
			                if mode="browse" then
			 	                'EligeCeldaResponsive "text", mode,"CELDA","","",0,LitNumFactura,LitNumFactura,0,d_lookup("nfactura_pro","facturas_pro","nfactura='" & iif(tmp_nfactura>"",tmp_nfactura,iif(isnull(rst("nfactura")),"",rst("nfactura"))) & "'",session("dsn_cliente"))
			                    nfacturaProSelect = "select nfactura_pro from facturas_pro with(nolock) where nfactura= ?"
                                EligeCeldaResponsive "text", mode,"CELDA","","",0,LitNumFactura,LitNumFactura,0,EncodeForHtml(DLookupP1(nfacturaProSelect, iif(tmp_nfactura>"",tmp_nfactura,iif(isnull(rst("nfactura")),"",rst("nfactura"))) & "", adVarchar, 20, session("dsn_cliente")))
                            else
                                %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                    %><label><%=LitNumFactura%></label><%
                                        'd_lookup("nfactura_pro","facturas_pro","nfactura='" & iif(tmp_nfactura>"",tmp_nfactura,iif(isnull(rst("nfactura")),"",rst("nfactura"))) & "'",session("dsn_cliente"))
                                        nfacturaProSelect = "select nfactura_pro from facturas_pro with(nolock) where nfactura= ?"
					                %><input class='width60' type="text" name="nfactura" value="<%=EncodeForHtml(DLookupP1(nfacturaProSelect, iif(tmp_nfactura>"",tmp_nfactura,iif(isnull(rst("nfactura")),"",rst("nfactura"))) & "", adVarchar, 20, session("dsn_cliente")))%>" disabled="disabled"/><%
				                %></div><%
			            end if
		            'CloseFila
		            'DrawFila "" 'color_blau
			            if si_tiene_modulo_proyectos<>0 then
				            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:140px'","","",0,LitProyecto+":"
				            if mode <> "browse" then
					            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                    %><label><%=LitProyecto%></label><%
						            %><input class="width60" type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto"))))%>"/><%
						            %><iframe id='frProyecto' name='fr_Proyecto' src='../mantenimiento/docproyectos.asp?viene=pedidos_pro&mode=<%=enc.EncodeForJavascript(mode)%>&cod_proyecto=<%=enc.EncodeForJavascript(iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto"))))%>' width='250' height='30' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
					            %></div><%
				            else
                                'EligeCeldaResponsive "text", mode,"CELDA","","",0,LitProyecto,LitProyecto,0,d_lookup("nombre","proyectos","codigo='" & rst("cod_proyecto") & "'",session("dsn_cliente"))
				                nombreProyectoSelect = "select nombre from proyectos where codigo= ?"
                                EligeCeldaResponsive "text", mode,"CELDA","","",0,LitProyecto,LitProyecto,0,EncodeForHtml(DLookupP1(nombreProyectoSelect, rst("cod_proyecto") & "", adVarchar, 15, session("dsn_cliente")))
                            end if
				            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"  "
			            end if
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:180px'","","",0,LitFechaEntrega+":"
			            if mode <> "browse" then
				            DrawDiv "1","",""
                            DrawLabel "","",LitFechaEntrega%><input class="CELDA" type="text" name="fecha_entrega" value="<%=EncodeForHtml(iif(tmp_fecha_entrega>"",tmp_fecha_entrega,iif(isnull(rst("fecha_entrega")),"",rst("fecha_entrega"))))%>"/><%
                            DrawCalendar "fecha_entrega"
                            CloseDiv
			            else
                            EligeCeldaResponsive "text",mode,"CELDA","","",0,LitFechaEntrega,LitFechaEntrega,0,EncodeForHtml(rst("fecha_entrega"))
				            'DrawCelda "CELDA","","",0,rst("fecha_entrega")
			            end if
		            'CloseFila
		            if mode="first_save" or mode="browse" then
			            'DrawFila "" 'color_blau
				            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " width='140px'","","",0,LitEDI+":"

				            '**RGU 17/8/2006**
				            'solo se podra pedir a proveedor si la empresa proveedora (CENTRAL COVALDROPER)
				            if mid(rst("edi"),6,5)=LitEmpCovaldroper then
					            'DrawCeldahref "CELDAREF","left","false",rst("edi"),"javascript:pedir('"&npedido&"')"
					            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12" id="redi"><%
                                    %><label><%=LitEDI%></label><%
                                    %><a class='CELDAREFB' href=javascript:pedir("<%=EncodeForHtml(null_s(npedido))%>")><%=EncodeForHtml(null_s(rst("edi")))%></a><%
					            %></div><% 
				            else
                                if mode="first_save" then
                                    EligeCelda "", mode,"CELDA","","",0,LitEDI,"nfactura",0,EncodeForHtml(rst("edi"))
                                else 
                                    EligeCeldaResponsive "text", mode,"CELDA","","",0,LitEdi,LitEdi,0,EncodeForHtml(rst("edi"))
                                end if
					            'DrawCelda "CELDA","","",0,rst("edi")&""
				            end if
				            '**RGU **
			            'CloseFila
		            end if
		            'DrawFila "" 'color_blau
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:140px'","","",0,LitIncotPedPro+":"
			            if mode="browse" then
                            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                    %><label><%=LitIncotPedPro%></label><%
                                    if rst("incoterms")<>"" then
                                        %><a class='CELDAREFB' href=javascript:pedir("<%=EncodeForHtml(null_s(npedido))%>")><%=EncodeForHtml(rst("incoterms"))&""%></a><%
                                    else
                                        DrawSpan "CELDA","","",""
                                    end if%></div><%
			            else
				            defecto=iif(tmp_incoterms>"",tmp_incoterms,iif(rst("incoterms")>"",rst("incoterms"),""))
                            rstAux.cursorlocation=3
				            rstAux.open "select codigo,codigo as descripcion from incoterms with(nolock) order by descripcion",session("dsn_cliente")
				            
                            DrawSelectCelda "CELDALEFT","60","",0,LitIncotPedPro,"incoterms",rstAux,defecto,"codigo","descripcion","",""
				            rstAux.close
			            end if
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:180px'","","",0,LitIncoPuntEntrPedPro+":"
			            if mode="browse" then
				            'DrawCelda "'CELDALEFT' align='left' ","","",0,rst("fob")&""
                            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                    %><label><%=LitIncoPuntEntrPedPro%></label><%
                                    if rst("fob")<>"" then
                                        %><a class='CELDAREFB' href=javascript:pedir("<%=EncodeForHtml(null_s(npedido))%>")><%=EncodeForHtml(rst("fob")&"")%></a><%
                                    else 
                                        DrawSpan "CELDA","","",""
                                    end if%></div><%
			            else
				            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                %><label><%=LitIncoPuntEntrPedPro%></label><%
			                    %><%defecto=iif(tmp_fob>"",tmp_fob,iif(rst("fob")>"",rst("fob"),""))%><%
					            %><input class="width60"  type="text"  name="fob" value="<%=EncodeForHtml(defecto)%>"/></div><%
			            end if
		            'CloseFila
                    'DrawFila "" 'color_blau
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:180px'","","",0,LITFREIGHT+":"
			            if mode="browse" then
				            'DrawCelda "CELDA","","",0,rst("portes")&""
                            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                    %><label><%=LITFREIGHT%></label><%
                                    if rst("portes") <> "" then
                                        %><a class='CELDAREFB' href=javascript:pedir("<%=EncodeForHtml(null_s(npedido))%>")><%=EncodeForHtml(rst("portes"))&""%> </a><%
                                    else
                                        DrawSpan "CELDA","","",""
                                    end if %></div><%
			            else
                            defecto=iif(tmp_portes>"",tmp_portes,rst("portes"))
				            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                %><label><%=LITFREIGHT%></label><%
					        %><select class='width60'  name="portes"><%
						        if defecto=LITOWED then
							        %><option selected="selected" value="<%=LITOWED%>"><%=LITOWED%></option><%
							        %><option value="<%=LITPAID%>"><%=LITPAID%></option><%
							        %><option value=""></option><%
						        elseif defecto=LITPAID then
							        %><option value="<%=LITOWED%>"><%=LITOWED%></option><%
							        %><option selected="selected" value="<%=LITPAID%>"><%=LITPAID%></option><%
							        %><option value=""></option><%
						        else
							        %><option value="<%=LITOWED%>"><%=LITOWED%></option><%
							        %><option value="<%=LITPAID%>"><%=LITPAID%></option><%
							        %><option selected="selected" value=""></option><%
						        end if
					            %></select></div><%
			             end if
		            'CloseFila
		            'DrawFila "" 'color_blau
		                display="none"
		                if si_tiene_modulo_ebesa="1" then
		                    display=""
		                end if
		                'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:180px;display:" & display & "'","","",0,LitFechaDePago+":"
			            if mode="browse" then
				            'DrawCelda "'CELDALEFT' style='display:" & display & "' align='left' ","","",0,rst("salida")&""
                            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                    %><label><%=LitFechaDePago%></label><%
                                    if rst("salida")<>"" then
                                        %><a class='CELDAREFB' href=javascript:pedir('<%=EncodeForHtml(npedido)%>') ><%=EncodeForHtml(rst("salida"))&""%> </a><%
                                    else 
                                        DrawSpan "CELDA","","",""
                                    end if%></div><%
			            else
			                DrawDiv "1","",""
                            DrawLabel "","",LitFechaDePago
                            defecto=iif(tmp_salida>"",tmp_salida,iif(rst("salida")>"",rst("salida"),""))%><input class="CELDA" maxlength="10" type="text" name="salida" value="<%=EncodeForHtml(defecto)%>"/><%
                            DrawCalendar "salida"
                            CloseDiv
			                if mode="add" then
			                    %><script language=javascript type="text/javascript">
                                      CalculaFechaPago()
			                    </script><%
				            end if
			            end if
		            'CloseFila
		            'DrawFila "" 'color_blau
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:140px' valign=top","","",0,LitObservaciones+":"
                        DrawDiv "1","",""
                            Drawlabel "","",LitObservaciones
			            if mode="browse" then
				            ''DrawCelda "CELDA","","",0,pintar_saltos_espacios(rst("observaciones")&"")
                            'DrawCeldaSpan "CELDA","","",0,pintar_saltos_espacios(rst("observaciones")&""),3
                            DrawSpan "CELDA","",pintar_saltos_nuevo(EncodeForHtml(rst("observaciones")&"")), ""
			            else
				            'DrawTextCeldaSpan "CELDA","","",2,100,"","observaciones",iif(rst("observaciones")>"",rst("observaciones"),""),3
                            DrawTextarea "width60","","observaciones",EncodeForHtml(iif(rst("observaciones")>"",rst("observaciones"),"")),""
			            end if
		            CloseDiv

		            'DrawFila color_blau
                    DrawDiv "1","",""
                        DrawLabel "","",LitNotas
			            'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:140px' valign=top","","",0,LitNotas+":"
			            if mode="browse" then
				            'DrawCelda "CELDA","","",0,pintar_saltos_espacios(rst("notas")&"")
				            'DrawCeldaSpan "CELDA","","",0,pintar_saltos_espacios(rst("notas")&""),3
                            DrawSpan "CELDA","",pintar_saltos_nuevo(EncodeForHtml(rst("notas")&"")), ""
			            else
				            'DrawTextCeldaSpan "CELDA","","",2,100,"","notas",iif(rst("notas")>"",rst("notas"),""),3
                            DrawTextarea "width60","","notas",EncodeForHtml(iif(rst("notas")>"",rst("notas"),"")),""
			            end if
		            CloseDiv

	            '************************'
	            'JMA 17/12/04 ***********'
	            '************************'
	            if mode="browse" and si_campo_personalizables=1 then
		      	DrawDiv "3-sub", "background-color: #eae7e3", ""
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                        'rstAux2.cursorlocation=3
			            'rstAux2.open "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
			            
                        set commandAux = nothing
                        set connAux = Server.CreateObject("ADODB.Connection")
                        set commandAux =  Server.CreateObject("ADODB.Command")
                        
                        connAux.open session("dsn_cliente")
                        connAux.cursorlocation=3
                        commandAux.ActiveConnection =connAux
                        commandAux.CommandTimeout = 60
                        commandAux.CommandText= "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like ?+'%' order by ncampo,titulo"
                        commandAux.CommandType = adCmdText
                        commandAux.Parameters.Append commandAux.CreateParameter("@ncampo", adChar,adParamInput,7, session("ncliente")&"")
                        
                        set rstAux2 = commandAux.Execute

                        if not rstAux2.eof then
				            'DrawFila ""
					            num_campo=1
					            num_campo2=1
					            num_puestos=0
					            num_puestos2=0
					            while not rstAux2.eof
						            if num_puestos2>0 and (num_puestos2 mod 2)=0 then
							            ''DrawCelda "CELDA style='width:125px'","","",0," "
							            'CloseFila
							            'DrawFila ""
							            num_puestos2=0
						            end if
						            if rstAux2("titulo") & "">"" then
							            if ((num_puestos-1) mod 2)=0 then
								            ''DrawCelda "CELDA style='width:155px'","","",0," "
							            end if
							            num_puestos=num_puestos+1
							            num_puestos2=num_puestos2+1
                                       
							            'DrawCelda "ENCABEZADOL style='width:155px'","","",0,rstAux2("titulo")
							            if rstAux2("tipo")=2 then
								            DrawCeldaResponsive "CELDA align=left style='width:155px'","","",0, rstAux2("titulo") & ":",iif(lista_valores(num_campo)=1,LitSi,LitNo)
							            elseif rstAux2("tipo")=3 then
								            if lista_valores(num_campo) & "">"" then
									            num_campo_str=cstr(num_campo)
									            if len(num_campo_str)=1 then
										            num_campo_str="0" & num_campo_str
									            end if
									            'valor_ListCampPerso=d_lookup("valor","campospersolista","ncampo='" & session("ncliente") & num_campo_str & "' and tabla='DOCUMENTOS COMPRA' and ndetlista=" & lista_valores(num_campo),session("dsn_cliente"))
								                valor_ListCampPersoSelect = "select valor from campospersolista with(nolock) where ncampo= ? and tabla='DOCUMENTOS COMPRA' and ndetlista= ?"
									            valor_ListCampPerso=DLookupP2(valor_ListCampPersoSelect, session("ncliente") & num_campo_str & "", adChar, 7, lista_valores(num_campo), adInt, 4, session("dsn_cliente"))
                                            else
									            valor_ListCampPerso=""
								            end if
                                            DrawCeldaResponsive "CELDA align=left style='width:200px'","","",0,rstAux2("titulo") & ":",EncodeForHtml(valor_ListCampPerso)
								            'DrawCelda "CELDA align='left' style='width:155px'","","",0,valor_ListCampPerso
							            else
                                            DrawCeldaResponsive "CELDA align=left style='width:200px'","","",0,rstAux2("titulo") & ":",EncodeForHtml(lista_valores(num_campo))
								            'DrawCelda "CELDA align='left' style='width:155px'","","",0,lista_valores(num_campo)
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
				            CloseFila
				            num_campos=num_puestos
			            else
				            num_campos=0
			            end if
                        connAux.Close
                        set connAux = nothing
                        set commandAux = nothing
			            'rstAux2.close
	            elseif mode="add" and si_campo_personalizables=1 then
	                %><!--<table width="100%" bgcolor='<%=color_blau%>' border = 1 cellpadding=2 cellspacing=5>--><%
		      	DrawDiv "3-sub", "background-color: #eae7e3", ""
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                        'rstAux2.cursorlocation=3
			            'rstAux2.open "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
			            
	                    set commandAux = nothing
                        set connAux = Server.CreateObject("ADODB.Connection")
                        set commandAux =  Server.CreateObject("ADODB.Command")

                        connAux.open session("dsn_cliente")
                        connAux.cursorlocation=3
                        commandAux.ActiveConnection =connAux
                        commandAux.CommandTimeout = 60
                        commandAux.CommandText=  "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like ?+'%' order by ncampo,titulo"
                        commandAux.CommandType = adCmdText
                        commandAux.Parameters.Append commandAux.CreateParameter("@ncampo", adChar,adParamInput,7, session("ncliente")&"")

                        set rstAux2 = commandAux.Execute

                        if not rstAux2.eof then
				            num_campos_existen=rstAux2.recordcount
				            'DrawFila ""
					            num_campo=1
					            num_campo2=1
					            num_puestos=0
					            num_puestos2=0
					            while not rstAux2.eof
						            if num_puestos2>0 and (num_puestos2 mod 2)=0 then
							            'DrawCelda "CELDA style='width:125px'","","",0," "
							            'CloseFila
							            'DrawFila ""
							            num_puestos2=0
						            end if
						            if rstAux2("titulo") & "">"" then
							            if ((num_puestos-1) mod 2)=0 then
								            'DrawCelda "CELDA style='width:155px'","","",0," "
							            end if
							            num_puestos=num_puestos+1
							            num_puestos2=num_puestos2+1
							            'DrawCelda "CELDA style='width:155px'","","",0,rstAux2("titulo") & " : "
							            valor_campo_perso=""

							            'JMA 17/12/04. Copiar campos personalizables de los proveedores'
                                        'rstSelect.cursorlocation=3
							            'rstSelect.open "select tipo,titulo from camposperso with(nolock) where ncampo='" & rstAux2("ncampocopia") & "' and tabla='PROVEEDORES'",session("dsn_cliente")
							            
                                        set commandSelect = nothing
                                        set connSelect = Server.CreateObject("ADODB.Connection")
                                        set commandSelect=  Server.CreateObject("ADODB.Command")

                                        connSelect.open session("dsn_cliente")
                                        connSelect.cursorlocation=3
                                        commandSelect.ActiveConnection =connSelect
                                        commandSelect.CommandTimeout = 60
                                        commandSelect.CommandText= "select tipo,titulo from camposperso with(nolock) where ncampo=? and tabla='PROVEEDORES'"
                                        commandSelect.CommandType = adCmdText
                                        commandSelect.Parameters.Append commandSelect.CreateParameter("@ncampo", adChar,adParamInput,7, rstAux2("ncampocopia")&"")

                                        set rstSelect = commandSelect.Execute

                                        if not rstSelect.eof then
								            tipoPro=rstSelect("tipo")
								            tituloPro=rstSelect("titulo")
							            end if
							            'rstSelect.close
                                        connSelect.Close
                                        set connSelect = nothing
                                        set commandSelect = nothing

							            if tipoPro=rstAux2("tipo") and tituloPro<>"" then
								            if rstAux2("ncampocopia")<>"" then
									            numCampoPro=cint(trimCodEmpresa(rstAux2("ncampocopia")))
									            valor_campo_perso=tmp_lista_valores(numCampoPro)
								            end if
							            end if
							            'JMA 17/12/04. FIN Copiar campos personalizables de los proveedores'

							            if rstAux2("tipo")=1 then
								            if isNumeric(rstAux2("tamany")) then
									            tamany=rstAux2("tamany")
								            else
									            tamany=1
								            end if
                          					DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" style="width:155px;" class="CELDA" name="<%="campo" & num_campo%>"  maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>" />
								            <%CloseDiv
							            elseif rstAux2("tipo")=2 then
								            if valor_campo_perso="" then
									            valor_campo_perso=0
								            end if
                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"
								            DrawCheckCelda "CELDA align=left style='width:155px' align='left'","","",0,"","campo" & num_campo,iif(valor_campo_perso=1,-1,0)
                                            CloseDiv
							            elseif rstAux2("tipo")=3 then
								            num_campo_str=cstr(num_campo)
								            if len(num_campo_str)=1 then
									            num_campo_str="0" & num_campo_str
								            end if
								            'strSelListVal="select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
                                            'rstAux.cursorlocation=3
								            'rstAux.open strSelListVal,session("dsn_cliente")

                                            set commandListVal = nothing
                                            set connListVal = Server.CreateObject("ADODB.Connection")
                                            set commandListVal =  Server.CreateObject("ADODB.Command")

                                            connListVal.open session("dsn_cliente")
                                            connListVal.cursorlocation=3
                                            commandListVal.ActiveConnection =connListVal
                                            commandListVal.CommandTimeout = 60
                                            commandListVal.CommandText= "select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo=? and valor is not null and valor<>'' order by valor,ndetlista"
                                            commandListVal.CommandType = adCmdText
                                            commandListVal.Parameters.Append commandListVal.CreateParameter("@ncampo", adChar,adParamInput,7, session("ncliente") & num_campo_str &"")

                                            set rstAux = commandListVal.Execute

                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"
                                            DrawSelect "","width:155px","campo"&num_campo,rstAux,valor_campo_perso,"ndetlista","valor","",""
                                            CloseDiv
                                                'DrawSelectCelda "CELDA style='width:155px'","","",0,"","campo"&num_campo,rstAux,valor_campo_perso,"ndetlista","valor","",""
			 					            	
                                            connListVal.Close
                                            set connListVal = nothing
                                            set commandListVal = nothing
                                            'rstAux.close
							            elseif rstAux2("tipo")=4 then
								            if isNumeric(rstAux2("tamany")) then
									            tamany=rstAux2("tamany")
								            else
									            tamany=1
								            end if
                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" class="CELDA" name="<%="campo" & num_campo%>" style="width:155px;" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
								            <%CloseDiv
							            elseif rstAux2("tipo")=5 then
								            if isNumeric(rstAux2("tamany")) then
									            tamany=rstAux2("tamany")
								            else
									            tamany=1
								            end if
                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" class="CELDA" name="<%="campo" & num_campo%>" style="width:155px;" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
								            <%CloseDiv
							            end if
						            else
							            %><input type="hidden" name="campo<%=num_campo%>" value=""/><%
						            end if
						            %><input type="hidden" name="tipo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("tipo"))%>"/>
						            <input type="hidden" name="titulo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("titulo"))%>"/><%
					                rstAux2.movenext
						            num_campo=num_campo+1
						            if not rstAux2.eof then
							            if rstAux2("titulo") & "">"" then
								            num_campo2=num_campo2+1
							            end if
						            end if
					            wend
				            CloseFila
				            num_campos=num_puestos
			            else
				            num_campos=0
				            num_campos_existen=0
			            end if
                        connAux.Close
                        set connAux = nothing
                        set commandAux = nothing
			            'rstAux2.close
		            %><input type="hidden" name="num_campos" value="<%=EncodeForHtml(num_campos_existen)%>"/><%
	            elseif mode="edit" and si_campo_personalizables=1 then
		      	DrawDiv "3-sub", "background-color: #eae7e3", ""
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                        'rstAux2.cursorlocation=3
			            'rstAux2.open "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
			            
                        set commandAux = nothing
                        set connAux = Server.CreateObject("ADODB.Connection")
                        set commandAux =  Server.CreateObject("ADODB.Command")

                        connAux.open session("dsn_cliente")
                        connAux.cursorlocation=3
                        commandAux.ActiveConnection =connAux
                        commandAux.CommandTimeout = 60
                        commandAux.CommandText= "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like ?+'%' order by ncampo,titulo"
                        commandAux.CommandType = adCmdText
                        commandAux.Parameters.Append commandAux.CreateParameter("@ncampo", adChar,adParamInput,7, session("ncliente")&"")

                        set rstAux2 = commandAux.Execute

                        if not rstAux2.eof then
				            num_campos_existen=rstAux2.recordcount
				            'DrawFila ""
					            num_campo=1
					            num_campo2=1
					            num_puestos=0
					            num_puestos2=0
					            while not rstAux2.eof
						            if num_puestos2>0 and (num_puestos2 mod 2)=0 then
							            'DrawCelda "CELDA style='width:125px'","","",0," "
							            'CloseFila
							            'DrawFila ""
							            num_puestos2=0
						            end if
						            if rstAux2("titulo") & "">"" then
							            num_puestos=num_puestos+1
							            num_puestos2=num_puestos2+1
							            'DrawCelda "CELDA style='width:155px'","","",0,rstAux2("titulo") & " : "
							            valor_campo_perso=lista_valores(num_campo)

							            'JMA 17/12/04. Copiar campos personalizables de los proveedores'
							            if nproveedor > "" then
                                            'rstSelect.cursorlocation=3
								            'rstSelect.open "select tipo,titulo from camposperso with(nolock) where ncampo='" & rstAux2("ncampocopia") & "' and tabla='PROVEEDORES'",session("dsn_cliente")
							    
                                            set commandSelect = nothing
                                            set connSelect = Server.CreateObject("ADODB.Connection")
                                            set commandSelect =  Server.CreateObject("ADODB.Command")

                                            connSelect.open session("dsn_cliente")
                                            connSelect.cursorlocation=3
                                            commandSelect.ActiveConnection =connSelect
                                            commandSelect.CommandTimeout = 60
                                            commandSelect.CommandText= "select tipo,titulo from camposperso with(nolock) where ncampo=? and tabla='PROVEEDORES'"
                                            commandSelect.CommandType = adCmdText
                                            commandSelect.Parameters.Append commandSelect.CreateParameter("@ncampo", adChar,adParamInput,7, rstAux2("ncampocopia")&"")

                                            set rstSelect = commandSelect.Execute

                                            if not rstSelect.eof then
									            tipoPro=rstSelect("tipo")
									            tituloPro=rstSelect("titulo")
								            end if
								            'rstSelect.close
                                            connSelect.Close
                                            set connSelect = nothing
                                            set commandSelect = nothing

								            if tipoPro=rstAux2("tipo") and tituloPro<>"" then
									            if rstAux2("ncampocopia")<>"" then
										            numCampoPro=cint(trimCodEmpresa(rstAux2("ncampocopia")))
										            valor_campo_perso=tmp_lista_valores(numCampoPro)
									            end if
								            end if
							            end if
							            'JMA 17/12/04. FIN Copiar campos personalizables de los proveedores'

							            if rstAux2("tipo")=1 then
								            if isNumeric(rstAux2("tamany")) then
									            tamany=rstAux2("tamany")
								            else
									            tamany=1
								            end if
                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" style="width:155px;" class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
								            <%CloseDiv
							            elseif rstAux2("tipo")=2 then
								            if valor_campo_perso="" then
									            valor_campo_perso=0
								            end if
								            DrawCheckCelda "CELDA style='width:155px' align='left'","","",0,"","campo" & num_campo,iif(valor_campo_perso=1,-1,0)
							            elseif rstAux2("tipo")=3 then
								            num_campo_str=cstr(num_campo)
								            if len(num_campo_str)=1 then
									            num_campo_str="0" & num_campo_str
								            end if
								            'strSelListVal="select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
                                            'rstAux.cursorlocation=3
								            'rstAux.open strSelListVal,session("dsn_cliente")
                                            
                                            set commandListVal = nothing
                                            set connListVal = Server.CreateObject("ADODB.Connection")
                                            set commandListVal =  Server.CreateObject("ADODB.Command")

                                            connListVal.open session("dsn_cliente")
                                            connListVal.cursorlocation=3
                                            commandListVal.ActiveConnection =connListVal
                                            commandListVal.CommandTimeout = 60
                                            commandListVal.CommandText= "select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo=? and valor is not null and valor<>'' order by valor,ndetlista"
                                            commandListVal.CommandType = adCmdText
                                            commandListVal.Parameters.Append commandListVal.CreateParameter("@ncampo", adChar,adParamInput,7, session("ncliente") & num_campo_str & "")

                                            set rstAux = commandListVal.Execute

                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"
			 						        DrawSelect "","width:155px","campo"&num_campo,rstAux,valor_campo_perso,"ndetlista","valor","",""
                                            CloseDiv
			 					            'rstAux.close
                                            connListVal.Close
                                            set connListVal = nothing
                                            set commandListVal = nothing
							            elseif rstAux2("tipo")=4 then
								            if isNumeric(rstAux2("tamany")) then
									            tamany=rstAux2("tamany")
								            else
									            tamany=1
								            end if
                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" style="width:155px;" class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
								            <%CloseDiv
							            elseif rstAux2("tipo")=5 then
								            if isNumeric(rstAux2("tamany")) then
									            tamany=rstAux2("tamany")
								            else
									            tamany=1
								            end if
                                            DrawDiv "1","",""
									        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" style="width:155px;" class="CELDA" name="<%="campo" & num_campo%>" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"/>
								            <%CloseDiv
							            end if
						            else
							            %><input type="hidden" name="campo<%=num_campo%>" value=""/><%
						            end if
						            %><input type="hidden" name="tipo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("tipo"))%>"/>
						            <input type="hidden" name="titulo_campo<%=num_campo%>" value="<%=EncodeForHtml(rstAux2("titulo"))%>"/><%
						            rstAux2.movenext
						            num_campo=num_campo+1
						            if not rstAux2.eof then
							            if rstAux2("titulo") & "">"" then
								            num_campo2=num_campo2+1
							            end if
						            end if
					            wend
				            CloseFila
				            num_campos=num_puestos
			            else
				            num_campos=0
				            num_campos_existen=0
			            end if
                        connAux.Close
                        set connAux = nothing
                        set commandAux = nothing
			            'rstAux2.close
		                %><input type="hidden" name="num_campos" value="<%=EncodeForHtml(num_campos_existen)%>"/><%
	                end if
                            
	            '************************'
	            'FIN JMA 28/10/04 *******'
	            '************************'%>
                </table>
            </div>
        </div>

		<%if mode="browse" or mode ="edit" then
 			'** Campo oculto para controlar si el pedido está facturado.
            'rstAux.cursorlocation=3
			'rstAux.Open "select nfactura from pedidos_pro with(nolock) where npedido='" & rst("npedido") & "'",session("dsn_cliente")

            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux =  Server.CreateObject("ADODB.Command")

            connAux.open session("dsn_cliente")
            connAux.cursorlocation=3
            commandAux.ActiveConnection =connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText= "select nfactura from pedidos_pro with(nolock) where npedido=?"
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, rst("npedido")&"")

            set rstAux = commandAux.Execute
	
			if not isnull(rstAux("nfactura")) then
				%><input type="hidden" name="h_nfactura" value="<%=EncodeForHtml(null_s(rstAux("nfactura")))%>"/><%
				'nfactura_pro=d_lookup("nfactura_pro","facturas_pro","nfactura='" & rstAux("nfactura") & "'",session("dsn_cliente"))
                nfactura_pro_select = "select nfactura_pro from facturas_pro with(nolock) where nfactura= ?"
				nfactura_pro=DLookupP1(nfactura_pro_select, rstAux("nfactura") & "", adVarchar, 20, session("dsn_cliente"))
				%><input type="hidden" name="h_nfacturapro" value="<%=EncodeForHtml(null_s(nfactura_pro))%>"/><%
			else
				%><input type="hidden" name="h_nfactura" value="NO"/>
				<input type="hidden" name="h_nfacturapro" value="NO"/><%
			end if
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
			'rstAux.close
            if mode="browse" then
            %><div class="Section" id="S_FINANCIAL_DATA">
                <a href="#" rel="toggle[FINANCIAL_DATA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LIT_FINANCIAL_DATA %>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
                <div class="SectionPanel" id="FINANCIAL_DATA" style="<%=iif(mode="add" or mode="edit","display: none","display: ")%>">  
                    <div id="tabs" style="display:none">
                        <ul>    
                            <li id="li-details"><a href="#tabs-details"><%=LitTituloDet %></a></li>
                            <li id="li-concepts"><a href="#tabs-concepts"><%=LitConceptos %></a></li>
                            <%if vencesp="1" then%>
                                <li id="li-vencimientos"><a href="#tabs-vencimientos"><%=LitVencimientos%></a></li>
                            <%end if%>
                            <li id="li-payments"><a href="#tabs-payments"><%=LitPagosACuenta %></a></li>
                            <li id="li-pedpro"><a href="#tabs-pedcli"><%=LitPedCli %></a></li>
                            <li id="li-send"><a href="#tabs-send"><%=LitDatosEnvio %></a></li>
                        </ul>
                        <%if oculta=0 then
			            'Mostrar los Detalles del pedido.
                            %><div id="tabs-details" class="overflowXauto">
                                <%set rstDet = Server.CreateObject("ADODB.Recordset")
                                'rstDet.cursorlocation=3
			                    'rstDet.open "select * from detalles_ped_pro where npedido='" & rst("npedido") & "' order by item", session("dsn_cliente")
			                    
                                set commandDet = nothing
                                set connDet = Server.CreateObject("ADODB.Connection")
                                set commandDet =  Server.CreateObject("ADODB.Command")

                                connDet.open session("dsn_cliente")
                                connDet.cursorlocation=3
                                commandDet.ActiveConnection =connDet
                                commandDet.CommandTimeout = 60
                                commandDet.CommandText= "select * from detalles_ped_pro where npedido=? order by item"
                                commandDet.CommandType = adCmdText
                                commandDet.Parameters.Append commandDet.CreateParameter("@ncliente", adVarChar,adParamInput,20, rst("npedido")&"")

                                set rstDet = commandDet.Execute
                                    
                                pagina="../central.asp?pag1=compras/pedidos_prodet.asp&ndoc=" + enc.EncodeForJavascript(rst("npedido")) + "&nproveedor=" + enc.EncodeForJavascript(rst("nproveedor")) + "&mode=add&pag2=compras/pedidos_prodet_bt.asp&titulo=" & LitDetallesPedido & " " & enc.EncodeForJavascript(rst("npedido"))
			                    %>
                                 <!-- GPD (05/03/2007) -->
                                <table width="835" border='0' cellspacing="1" cellpadding="1">
				                    <% DrawFila "" %>
		   		                        <td width="100">
		   		                            <%if si_tiene_modulo_ebesa="1"  and isnull(rst("nfactura")) and isnull(rst("nalbaran")) then%>
		   		                            <a class="CELDAREFB" href="javascript:mostrarCondicionesCompra('<%=enc.EncodeForJavascript(null_s(rst("nproveedor")))%>');"><%=LitCondCompra%></a>		   		        
		   		                            <%end if%>
		   		                        </td>		   		    
		   	                        </tr>
                                </table>
			                    <table class="width90 md-table-responsive bCollapse">
			                        <%DrawFila color_terra
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitItem
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitCantidad
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,"PR/R"
					                    DrawCeldaDet "'CELDAL7 underOrange width10'","","",0,LitReferencia
					                    DrawCeldaDet "'CELDAL7 underOrange width15'","","",0,LitDescripcion
					                    if si_tiene_acceso_almacenes=1 then
					                        DrawCeldaDet "'CELDAL7 underOrange width10'","","",0,LitAlmacen
					                    end if
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitPVP
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitDto
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitDto2
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitIva
					                    DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitImporte
					                    DrawCeldaDet "'CELDAL7 underOrange width10'","","",0,"&nbsp;"
				                    CloseFila%>
			                    </table>
                                <!-- Se Agrego validacion si se realizo cierre administrativo(True) o no(False) -->
			                    <%if isnull(rst("nfactura")) and isnull(rst("nalbaran")) AND NOT blnEstadoCierre then
					                %><iframe id='frDetallesIns' class="width90 iframe-input md-table-responsive" name='fr_DetallesIns' src='pedidos_prodetins.asp?ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>&nproveedor=<%=enc.EncodeForJavascript(null_s(rst("nproveedor")))%>&modp=<%=enc.EncodeForJavascript(null_s(modP))%>&almacenSerie=<%=enc.EncodeForJavascript(null_s(almacenSerie)) %>&almacenTPV=<%=enc.EncodeForJavascript(null_s(almacenTPV)) %>' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
				                end if
				                %><iframe id='frDetalles' class="width90 md-table-responsive" name="fr_Detalles" src='pedidos_prodet.asp?ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>&nproveedor=<%=enc.EncodeForJavascript(null_s(rst("nproveedor")))%>&modp=<%=enc.EncodeForJavascript(null_s(modP))%>&almacenSerie=<%=enc.EncodeForJavascript(null_s(almacenSerie)) %>&almacenTPV=<%=enc.EncodeForJavascript(null_s(almacenTPV)) %>&EstadoCierre=<%=enc.EncodeForJavascript(null_s(blnEstadoCierre)) %>' height='150' frameborder="yes" noresize="noresize"></iframe>
                                <span id="paginacion" style="display: ">
		                        </span>
                            </div>
                        <%end if ' del oculta

			            ''ricardo 9/8/2004 se pondra el iva que tiene establecido el cliente
			            'TmpIvaProveedor=d_lookup("iva","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))
                        TmpIvaProveedorSelect = "select iva from proveedores with(nolock) where nproveedor= ?"
                        TmpIvaProveedor = DLookupP1(TmpIvaProveedorSelect, rst("nproveedor") & "", adChar, 10, session("dsn_cliente"))
			            'defaultIVA=d_lookup("iva","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
                        defaultIVASelect = "select iva from configuracion with(nolock) where nempresa= ?"
                        defaultIVA= DLookupP1(defaultIVASelect, session("ncliente") & "", adChar, 5, session("dsn_cliente"))


			            if TmpIvaProveedor & "">"" then
				            TmpIva=TmpIvaProveedor
			            else
				            TmpIva=defaultIVA
			            end if%>

                        <div id="tabs-concepts" class="overflowXauto">
		                    <input type="hidden" name="defaultIva" value="<%=EncodeForHtml(null_s(TmpIva))%>"/>
		                    <%
                            'rstAux.cursorlocation=3
                            'rstAux.Open "select * from conceptos_ped_pro where npedido='" & npedido & "'",session("dsn_cliente")
			                
                            set commandAux = nothing
                            set connAux = Server.CreateObject("ADODB.Connection")
                            set commandAux =  Server.CreateObject("ADODB.Command")

                            connAux.open session("dsn_cliente")
                            connAux.cursorlocation=3
                            commandAux.ActiveConnection =connAux
                            commandAux.CommandTimeout = 60
                            commandAux.CommandText= "select * from conceptos_ped_pro where npedido=?"
                            commandAux.CommandType = adCmdText
                            commandAux.Parameters.Append commandAux.CreateParameter("@npedido", adVarChar,adParamInput,20, npedido)

                            set rstAux = commandAux.Execute

                            pagina="../central.asp?pag1=compras/pedidos_procon.asp&ndoc=" + enc.EncodeForJavascript(npedido) + "&nproveedor=" + enc.EncodeForJavascript(nproveedor) + "&mode=add&pag2=compras/pedidos_procon_bt.asp&titulo=" &  LitConceptoPedido & " " & enc.EncodeForJavascript(rst("npedido"))
	      	                %>
		   	                <table class="width90 md-table-responsive bCollapse">
		   	                    <%'Fila de encabezado
				                DrawFila color_terra
					                DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitItem
					                DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitCantidad
					                DrawCeldaDet "'CELDAL7 underOrange width15'","","",0,LitDescripcion
					                DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitPVP
					                DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitDto
					                DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitIva
					                DrawCeldaDet "'CELDAL7 underOrange width5'","","",0,LitImporte
                                    DrawCeldaDet "'CELDAL7 underOrange width10'","","",0,"&nbsp;"
				                CloseFila

			     	                if isnull(rst("nfactura")) and isnull(rst("nalbaran")) then
					                'Linea de inserción de un detalle
					                DrawFila color_blau%>
						                <td class='CELDAL7 underOrange width5' >
							                &nbsp;
						                </td>
						                <td class='CELDAR7 underOrange width5'>
							                <input class='CELDAR7 width100' type="text" name="cantidad" value="1" onchange="RoundNumValue(this,<%=EncodeForHtml(dec_cant)%>);ImporteDetalle();"/>
						                </td>
						                <td class='CELDAL7 underOrange width15'>
							                <textarea class='CELDAL7 width100' name="descripcion"  rows="2" cols="15"></textarea>
						                </td>
						                <td class='CELDAR7 underOrange width5'>
							                <input class='CELDAR7 width100' type="text" name="pvp" value="0" onchange="RoundNumValue(this,<%=EncodeForHtml(dec_prec)%>);ImporteDetalle();"/>
						                </td>
						                <td class='CELDAR7 underOrange width5'>
							                <input class='CELDAR7 width100' type="text" name="descuento"  value="0" onchange="RoundNumValue(this,<%=EncodeForHtml(decpor)%>);ImporteDetalle();"/>
						                </td>
					                    <%
                                        rstSelect.cursorlocation=3
                                        rstSelect.open "select tipo_iva, tipo_iva from tipos_iva with(nolock)",session("dsn_cliente")
					                    DrawSelectCeldaDet "'CELDAL7 underOrange width5'","width100","",0,"","iva",rstSelect,TmpIva,"tipo_iva","tipo_iva","",""
					                    rstSelect.close%>
						                <td class='CELDAL7 underOrange width5' >
							                <input class='CELDAL7 width100' disabled="disabled" type="text" name="importe" value="0"/>																					
						                </td>
						                <td class="underOrange width10">
							                <a class='ic-accept noMTop' href="javascript:addConcepto('<%=enc.EncodeForJavascript(npedido)%>');" onblur="javascript:document.pedidos_pro.cantidad.focus();">
                                                <img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>
						                </td>
						                <%if oculta=1 then%>
						                  <script language="javascript" type="text/javascript">
                                              document.pedidos_pro.cantidad.focus();
                                              document.pedidos_pro.cantidad.select();
						                   </script>
						                <%end if
					                CloseFila
				                end if%>
				                </table>
				                <iframe id="frConceptos" name="fr_Conceptos" class="width90 md-table-responsive" src='pedidos_procon.asp?mode=browse&ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>' height='80' frameborder="yes" noresize="noresize"></iframe>
                            </div>
                            <%if vencesp="1" then%>
                                <div id="tabs-vencimientos">
			                        <table class="width90 md-table-responsive bCollapse"><%
				                        'Fila de encabezado
					                        DrawFila color_terra
                                                %>
                                                <td class='CELDAL7 underOrange width20' >&nbsp</td>
						                        <td class='CELDAL7 underOrange width20' ><%=LitFecha%></td>
                                                <td class='CELDAL7 underOrange width20' ><%=LITDIASFF%></td>
						                        <td class='CELDAL7 underOrange width20' >%</td>
						                        <td class='CELDAL7 underOrange width20' >&nbsp</td>
                                                <%
					                        CloseFila
                                            if isnull(rst("nfactura")) and isnull(rst("nalbaran")) then
						                        DrawFila color_blau
							                        %>
							                        <td class='CELDAL7 underOrange width20'>
							                        </td>
							                        <td class='CELDAL7 underOrange width20' >
								                        <input class='CELDAL7 width100' type="text" name="fechaVto" value="" onchange="javascript:cambiarFecVto()"/>
							                        </td>
                                                    <%DrawCalendar "fechaVto"%>
							                        <td class='CELDAL7 underOrange width20' >
								                        <input class='CELDAR7 width100' type="text" name="DiasFFVto" value="" onchange="javascript:cambiarDiasFFVto()"/>
							                        </td>
							                        <td class='CELDAL7 underOrange width20' >
								                        <input class='CELDAR7 width100' type="text" name="tantoVto" value="0" onchange="javascript:RoundNumValue(this,<%=EncodeForHtml(NdecDiPedido)%>);"/>
							                        </td>
							                        <td class='CELDAL7 underOrange width20' >
								                        <a class='ic-accept noMTop' href="javascript:addVencimiento('<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>');" onblur="javascript:document.pedidos_pro.fechaVto.focus();">
                                                            <img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/>
								                        </a>
							                        </td><%
						                        CloseFila
                                            end if
                                            %>
					                        <!--</td>
					                        </tr>-->
				                        </table><% 'FIN DE TABLA QUE CONTIENE LA TABLAS DE ARTICULOS
				                        %><iframe id="frVencimientos" name="fr_Vencimientos" class="width90 md-table-responsive" src='vencimientos_pro_config.asp?mode=browse&tdocumento=pedidos_pro&ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>' style="width:355px;height:200px;" frameborder="yes" noresize="noresize"></iframe>
                                </div>
                            <%end if%>
                        <div id="tabs-payments" class="overflowXauto">
			                <table class="width90 md-table-responsive bCollapse"><%
			                'Fila de encabezado
			                DrawFila color_terra
				                %><td class="CELDAL7 underOrange width5" ><%=LitNumPago%></td>
				                <td class="CELDAL7 underOrange width5" ><%=LitFecha%></td>
				                <td class="CELDAL7 underOrange width15" ><%=LitDescripcion%></td>
				                <td class="CELDAL7 underOrange width5" ><%=LitImporte%></td>
				                <td class="CELDAL7 underOrange width5" ><%=LitTipoPago%></td>
				                <td class="CELDAL7 underOrange width5" >&nbsp</td>
			                <%CloseFila
			                if isnull(rst("nfactura")) and isnull(rst("nalbaran")) then
				                'Linea de inserción de un pago a cuenta
				                DrawFila color_blau%>
					                <td class="CELDAL7 underOrange width5" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					                <td class="CELDAR7 underOrange width5" >
						                <input class="CELDAR7 width60" type="text" name="fechaPago" value="" onchange="cambiarfecha(document.pedidos_pro.fechaPago.value,'Fecha Pago')"/><%
                                        DrawCalendar "fechaPago"%>
					                </td>
					                <td class="CELDAL7 underOrange width15" >
						                <textarea class="CELDAL7 width100" name="descripcionPago" onfocus="lenmensaje(this,0,50,'')" onkeydown="lenmensaje(this,0,50,'')" onkeyup="lenmensaje(this,0,50,'')" onBlur="lenmensaje(this,0,50,'')" rows="2"></textarea>
					                </td>
					                <td class="CELDAR7 underOrange width5">
						                <input class="CELDAR7 width100" type="text" name="importePago" value="0" onchange="RoundNumValue(this,<%=EncodeForHtml(NdecDiPedido)%>);importepagoComp();"/>
					                </td>
					                <%
                                    'rstSelect.cursorlocation=3
                                    'rstSelect.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
					                
                                    set commandSelect = nothing
                                    set connSelect = Server.CreateObject("ADODB.Connection")
                                    set commandSelect =  Server.CreateObject("ADODB.Command")

                                    connSelect.open session("dsn_cliente")
                                    connSelect.cursorlocation=3
                                    commandSelect.ActiveConnection =connSelect
                                    commandSelect.CommandTimeout = 60
                                    commandSelect.CommandText= "select codigo,descripcion from tipo_pago with(nolock) where codigo like ?+'%' order by descripcion"
                                    commandSelect.CommandType = adCmdText
                                    commandSelect.Parameters.Append commandSelect.CreateParameter("@codigo", adVarChar,adParamInput,8, session("ncliente")&"")

                                    set rstSelect = commandSelect.Execute

                                    DrawSelectCeldaDet "'CELDAL7 underOrange width5'","width100","",0,"","tipoPago",rstSelect,"","codigo","descripcion","",""
					                
                                    connSelect.Close
                                    set connSelect = nothing
                                    set commandSelect = nothing
                                    'rstSelect.close
                                    %>
					                <td class="CELDAL7 underOrange width5">
						                <a class='ic-accept noMTop' href="javascript:addPago('<%=enc.EncodeForJavascript(npedido)%>');" onblur="javascript:document.pedidos_pro.fechaPago.focus();">
                                            <img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/>
						                </a>
					                </td>
				                <%CloseFila
			                end if%>
			                </table>
			                <iframe id="frPagosCuenta" name="fr_PagosCuenta"  class="width90 md-table-responsive" src='pedidos_propagos.asp?mode=browse&ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>' height='80' frameborder="yes" noresize="noresize"></iframe>
                        </div>
                        <div id="tabs-pedcli" class="overflowXauto" >
			                <table class="width90 md-table-responsive bCollapse">
			                    <%DrawFila color_terra%>
					                <td class='CELDAL7 underOrange width5' width="30"><%=LitItem%></td>
					                <td class='CELDAL7 underOrange width5' width="50"><%=LitCantidad%></td>
					                <td class='CELDAL7 underOrange width10' width="135"><%=LitReferencia%></td>
					                <td class='CELDAL7 underOrange width15' width="160"><%=LitDescripcion%></td>
					                <td class='CELDAL7 underOrange width5' width="160"><%=LitPedCliente%></td>
					                <td class='CELDAL7 underOrange width5' width="30"><%=LitItem%></td>
				                <%CloseFila%>
			                </table>
				            <iframe id="frPedcli" name="fr_Pedcli" class="width90 md-table-responsive" src='pedidos_propedcli.asp?mode=browse&ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>' width='600' height='120' frameborder="yes" noresize="noresize"></iframe>
                        </div>
                        <div id="tabs-send">
                            <%if mode="browse" then
			                    if rst("dir_envio")>"" then
				                    pagina="../central.asp?pag1=./compras/pedidos_prodireccion_env.asp&ndoc=" & enc.EncodeForJavascript(rst("npedido")) & "&mode=browse&pag2=./compras/pedidos_prodireccion_env_bt.asp&titulo=" & ucase(LitDatosEnvio) & " " & enc.EncodeForJavascript(trimCodEmpresa(rst("npedido")))
			                    else
				                    pagina="../central.asp?pag1=./compras/pedidos_prodireccion_env.asp&ndoc=" & enc.EncodeForJavascript(rst("npedido")) & "&mode=edit&pag2=./compras/pedidos_prodireccion_env_bt.asp&titulo=" & ucase(LitDatosEnvio) & " " & enc.EncodeForJavascript(trimCodEmpresa(rst("npedido")))
			                    end if
                                'rstDomi.cursorlocation=3
			                    'rstDomi.Open "select * from domicilios with(nolock) where codigo='" & rst("dir_envio") & "'",session("dsn_cliente")
			                
                                set commandDomi = nothing
                                set connDomi = Server.CreateObject("ADODB.Connection")
                                set commandDomi =  Server.CreateObject("ADODB.Command")

                                connDomi.open session("dsn_cliente")
                                connDomi.cursorlocation=3
                                commandDomi.ActiveConnection =connDomi
                                commandDomi.CommandTimeout = 60
                                commandDomi.CommandText= "select * from domicilios with(nolock) where codigo=?"
                                commandDomi.CommandType = adCmdText
                                commandDomi.Parameters.Append commandDomi.CreateParameter("@codigo", adInteger,adParamInput,4, rst("dir_envio"))

                                set rstDomi = commandDomi.Execute

                            end if%>
                            <table width='100%' border='0' cellspacing="0" cellpadding="0">
			                    <%'DrawFila color_terra%>
                                <tr>
				                    <td>
      				                    <table border="1" cellspacing="0" cellpadding="0">
        				                    <%'DrawFila color_blau2%>
                                            <tr>
						                        <td class="ENCABEZADOC" height="25">
							                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=LitDatosEnvio%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class='CELDAREFB' href="javascript:AbrirVentana('<%=pagina%>','P',<%=altoventana%>,<%=anchoventana%>)" OnMouseOver="self.status='<%=LitEditar%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitEditar%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                                </td>
                                            </tr>
					                            <%CloseFila%>
				                        </table>
				                    </td>
                                </tr>
	        	                <%'CloseFila
			                    if not rstDomi.eof then %>
			                    <tr>
				                    <td>
				                        <table width='100%' border='0' cellspacing="1" cellpadding="1">
					                        <%'DrawFila color_terra
							                    'DrawCelda "ENCABEZADOL","","",0,LitDomicilio
							                    'DrawCelda "ENCABEZADOL","","",0,LitPoblacion
							                    'DrawCelda "ENCABEZADOL","","",0,LitCP
							                    'DrawCelda "ENCABEZADOL","","",0,LitProvincia
						                    'CloseFila%>

                                            <tr>
                                                <td class="ENCABEZADOL"><%=LitDomicilio%></td>
                                                <td class="ENCABEZADOL"><%=LitPoblacion%></td>
                                                <td class="ENCABEZADOL"><%=LitCP%></td>
                                                <td class="ENCABEZADOL"><%=LitProvincia%></td>
                                            </tr>

                                            <%
						                    'DrawFila color_blau2
							                    'DrawCelda "CELDA","","",0,rstDomi("domicilio")
							                    'DrawCelda "CELDA","","",0,rstDomi("poblacion")
							                    'DrawCelda "CELDA","","",0,rstDomi("cp")
							                    'DrawCelda "CELDA","","",0,rstDomi("provincia")
						                    'CloseFila%>

                                            <tr>
                                                <td class="CELDA"><%=EncodeForHtml(null_s(rstDomi("domicilio")))%></td>
                                                <td class="CELDA"><%=EncodeForHtml(null_s(rstDomi("poblacion")))%></td>
                                                <td class="CELDA"><%=EncodeForHtml(null_s(rstDomi("cp")))%></td>
                                                <td class="CELDA"><%=EncodeForHtml(null_s(rstDomi("provincia")))%></td>
                                            </tr>
                                        </table>
				                    </td>
			                    </tr>
                                <%end if
                                connDomi.Close
                                set connDomi = nothing
                                set commandDomi = nothing
			                    'rstDomi.close
                                %>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
            <%end if 
        end if%>
			
		<input type="hidden" name="sumadet" value="<%=EncodeForHtml(null_s(sumadet))%>"/>
		<input type="hidden" name="sumaRE" value="<%=EncodeForHtml(null_s(sumaRE))%>"/>
		<input type="hidden" name="importe_bruto2" value="<%=EncodeForHtml(null_s(rst("importe_bruto")))%>"/>
  
        <% ''enc.EncodeForHtmlAttribute(rst("descuento")) y enc.EncodeForHtmlAttribute(rst("descuento2")) falla la página. %>
        <input type="hidden" name="desc1" value="<%=EncodeForHtml(null_s(rst("descuento")))%>"/>
		<input type="hidden" name="desc2" value="<%=EncodeForHtml(null_s(rst("descuento2")))%>"/>
		
        <div class="Section" id="S_TOTAL">
            <div class="SectionHeader2">
                <%=LitAbrevia %>
            </div>
            <div class="SectionPanel" id="DATTOTAL"><%
		            'Fila de encabezado
		            'DrawFila color_fondo
			        DrawDiv "4", "", ""
                        DrawLabel "", "",LitAbrevia
                        'd_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente"))
                        abrevDivSelect = "select abreviatura from divisas with(nolock) where codigo= ?"
                        DrawCelda "ENCABEZADOL","","",0, EncodeForHtml(DLookupP1(abrevDivSelect, session("ncliente") & "01" &"", adVarchar, 15, session("dsn_cliente")))

                    CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitBruto
				            'EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='importe_bruto'","ENCABEZADOR disabled"),"","",0,"","importe_bruto",10,formatnumber(null_z(rst("importe_bruto")),n_decimales,-1,0,iif(mode="browse",-1,0))
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR", "","importe_bruto",EncodeForHtml(formatnumber(null_z(rst("importe_bruto")),n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='importe_bruto'","disabled")
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitDto
                            if mode<>"browse" then
					            %><input class="ENCABEZADOR" type="text" name="dto" value="<%=EncodeForHtml(iif(tmp_descuento>0,tmp_descuento,iif(rst("descuento")>"",rst("descuento"),0)))%>" onchange="RoundNumValue(this,<%=EncodeForHtml(DECPOR)%>);Recalcula('<%=EncodeForHtml(total_iva_bruto)%>','<%=EncodeForHtml(total_re_bruto)%>');"/><%
				            else
					            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","dto", EncodeForHtml(cstr(formatnumber(null_z(rst("descuento")),decpor,-1,0,iif(mode="browse",-1,0)))) + "%", "id='dto'"
				            end if
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitDto2
                            if mode<>"browse" then
					            %><input class="ENCABEZADOR" type="text" name="dto2" value="<%=EncodeForHtml(iif(tmp_descuento2>0,tmp_descuento2,iif(rst("descuento2")>"",rst("descuento2"),0)))%>" onchange="RoundNumValue(this,<%=EncodeForHtml(DECPOR)%>);Recalcula('<%=EncodeForHtml(total_iva_bruto)%>','<%=EncodeForHtml(total_re_bruto)%>');" /><%
				            else
					            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","dto2",EncodeForHtml(cstr(formatnumber(null_z(rst("descuento2")),decpor,-1,0,iif(mode="browse",-1,0)))) + "%", "id='dto2'"
				            end if
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitTotalDescuento
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_descuento",EncodeForHtml(formatnumber(null_z(rst("total_descuento")),n_decimales,-1,0,iif(mode="browse",-1,0))),iif(mode="browse","id='total_descuento'","disabled")
				            %><input type="hidden" name="h_total_descuento" value="<%=EncodeForHtml(null_s(rst("total_descuento")))%>"/><%
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitImponible
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","base_imponible",EncodeForHtml(formatnumber(null_z(rst("base_imponible")),n_decimales,-1,0,iif(mode="browse",-1,0))),iif(mode="browse","id='base_imponible'","disabled")
				            %><input type="hidden" name="h_base_imponible" value="<%=EncodeForHtml(null_s(rst("base_imponible")))%>"/><%
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitTotalIva
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_iva",EncodeForHtml(cstr(formatnumber(Null_z(rst("total_iva")),n_decimales,-1,0,iif(mode="browse",-1,0)))),iif(mode="browse","id='total_iva'","disabled")
				            %><input type="hidden" name="h_total_iva" value="<%=EncodeForHtml(null_s(rst("total_iva")))%>"/><%
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitRe
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_re",EncodeForHtml(cstr(formatnumber(Null_z(rst("total_re")),n_decimales,-1,0,iif(mode="browse",-1,0)))),iif(mode="browse","id='total_re'","disabled")
				            %><input type="hidden" name="h_total_re" value="<%=EncodeForHtml(null_s(rst("total_re")))%>"/><%
                        CloseDiv%>
			            <%if ((rst("recargo")<>0) or mode="edit" or mode="add") then%>

				            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitRecargo
                            if mode<>"browse" then
						        %><input class="ENCABEZADOR" type="text" name="recargo" value="<%=EncodeForHtml(iif(tmp_recargo>0,tmp_recargo,iif(rst("recargo")>"",rst("recargo"),0)))%>" onchange="RoundNumValue(this,<%=EncodeForHtml(DECPOR)%>);Recalcula('<%=EncodeForHtml(total_iva_bruto)%>','<%=EncodeForHtml(total_re_bruto)%>');"/><%
					        else
						        EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","recargo",EncodeForHtml(formatnumber(null_z(rst("recargo")),2,-1,0,iif(mode="browse",-1,0))) + "%","id='recargo'"
					        end if
	                        CloseDiv%>

				            <%DrawDiv "4", "", ""
                                DrawLabel "", "",LitTotalRecargo
                                EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_recargo",EncodeForHtml(formatnumber(null_z(rst("total_recargo")),n_decimales,-1,0,iif(mode="browse",-1,0))),iif(mode="browse","id='total_recargo'","disabled")
					            %><input type="hidden" name="h_total_rf" value="<%=EncodeForHtml(rst("total_recargo"))%>"/><%
                            CloseDiv%>
			            <%end if

			            if ((rst("irpf")<>0) or mode="edit" or mode="add") then%>

				            <%DrawDiv "4", "", ""
                                DrawLabel "", "",Litirpf
                                if mode<>"browse" then
						            %><input class="ENCABEZADOR" type="text" name="irpf" value="<%=EncodeForHtml(iif(tmp_irpf>0,tmp_irpf,iif(rst("irpf")>"",rst("irpf"),0)))%>" onchange="RoundNumValue(this,<%=EncodeForHtml(DECPOR)%>);Recalcula('<%=EncodeForHtml(total_iva_bruto)%>','<%=EncodeForHtml(total_re_bruto)%>');"/><%
					            else
						            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","irpf",EncodeForHtml(formatnumber(null_z(rst("irpf")),2,-1,0,iif(mode="browse",-1,0))) + "%","id='irpf'"
					            end if
                            CloseDiv%>

				            <%DrawDiv "4", "", ""
                                DrawLabel "", "",LitTotalirpf
                                EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_irpf",EncodeForHtml(formatnumber(null_z(rst("total_irpf")),n_decimales,-1,0,iif(mode="browse",-1,0))),iif(mode="browse","id='total_irpf'","disabled")
					            %><input type="hidden" name="h_total_irpf" value="<%=EncodeForHtml(rst("total_irpf"))%>"/><%
                            CloseDiv%>
			            <%end if%>


			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitTotal
				            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_pedido",EncodeForHtml(formatnumber(null_z(rst("total_pedido")),n_decimales,-1,0,iif(mode="browse",-1,0))),iif(mode="browse","id='total_pedido'","disabled")
				            %><input type="hidden" name="h_total_pedido" value="<%=EncodeForHtml(rst("total_pedido"))%>"/><%
                        CloseDiv%>
				        <%%><input class="CELDA" type="hidden" name="IRPF_Total" value="<%=EncodeForHtml(iif(tmp_IRPF_Total>"",tmp_IRPF_Total,iif(isnull(rst("IRPF_Total")),"",rst("IRPF_Total"))))%>"/><%


			        'CloseFila
                    'd_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
                    imp_equiv_select = "select imp_equiv from configuracion with(nolock) where nempresa= ?"
			        if DLookupP1(imp_equiv_select, session("ncliente") & "", adChar, 5, session("dsn_cliente")) then
				        'DrawFila color_blau
                        'd_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                        codigoDivisaSelect = "select codigo from divisas with(nolock) where moneda_base<>0 and codigo like ?+'%'"
				        DIVISA=iif(tmp_divisa>"",tmp_divisa,iif(mode="add", DLookupP1(codigoDivisaSelect, session("ncliente") & "", adVarchar, 15, session("dsn_cliente")) ,rst("divisa")))
			                'd_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente"))
                            abrevDivSelect = "select abreviatura from divisas with(nolock) where codigo= ?+'01'"
                            DrawCelda "ENCABEZADOL","","",0, EncodeForHtml(DLookupP1(abrevDivSelect,  session("ncliente") & "", adVarchar, 15, session("dsn_cliente")))
                            'd_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente"))
                            ndecDivSelect = "select ndecimales from divisas with(nolock) where codigo= ?+'01'"
				        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Pimporte_bruto'","ENCABEZADOR disabled"),"","",0,"","Pimporte_bruto",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("importe_bruto")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))
				        EligeCelda "input", mode,"ENCABEZADOR disabled id='Pdescuento'","","",0,"","Pdescuento",3,EncodeForHtml(formatnumber(null_z(rst("descuento")),decpor,-1,0,iif(mode="browse",-1,0)))
				        EligeCelda "input", mode,"ENCABEZADOR disabled id='Pdescuento2'","","",0,"","Pdescuento2",3,EncodeForHtml(formatnumber(null_z(rst("descuento2")),decpor,-1,0,iif(mode="browse",-1,0)))
                       'd_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente"))
                        ndecDivSelect = "select ndecimales from divisas with(nolock) where codigo= ?+'01'"
				        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Ptotal_descuento'","ENCABEZADOR disabled"),"","",0,"","Ptotal_descuento",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_descuento")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))
				        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Pbase_imponible'","ENCABEZADOR disabled"),"","",0,"","Pbase_imponible",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("base_imponible")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))
				        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Ptotal_iva'","ENCABEZADOR disabled"),"","",0,"","Ptotal_iva",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_iva")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))
				        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Ptotal_re'","ENCABEZADOR disabled"),"","",0,"","Ptotal_re",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_re")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))
				        if ((rst("recargo")<>0) or mode="edit" or mode="add") then
					        EligeCelda "input", mode,"ENCABEZADOR disabled id='Precargo'","","",0,"","Precargo",3,EncodeForHtml(formatnumber(null_z(rst("recargo")),2,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Ptotal_recargo'","ENCABEZADOR disabled"),"","",0,"","Ptotal_recargo",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_recargo")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))
				        end if
				        if ((rst("irpf")<>0) or mode="edit" or mode="add") then
					        EligeCelda "input", mode,"ENCABEZADOR disabled id='Pirpf'","","",0,"","Pirpf",3,EncodeForHtml(formatnumber(null_z(rst("irpf")),2,-1,0,iif(mode="browse",-1,0)))
					        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Ptotal_irpf'","ENCABEZADOR disabled"),"","",0,"","Ptotal_irpf",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_irpf")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))
				        end if
				        EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='Ptotal_pedido'","ENCABEZADOR disabled"),"","",0,"","Ptotal_pedido",10,EncodeForHtml(formatnumber(CambioDivisa(null_z(rst("total_pedido")),DIVISA,session("ncliente") & "01"), DLookupP1(ndecDivSelect, session("ncliente"), adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0)))

				        'CloseFila
			        end if
            %></div><%
		%></div><%
    

	if mode="add" then
		%><script language="javascript" type="text/javascript">
                                              document.pedidos_pro.fecha.focus();
                                              document.pedidos_pro.fecha.select();
		</script>
	<%elseif mode="edit" then%>
		<script language="javascript" type="text/javascript">
                                              document.pedidos_pro.fecha.focus();
                                              document.pedidos_pro.fecha.select();
		</script>
	<%end if
elseif rst.EOF then%>
  	<script language="javascript" type="text/javascript">
                                              parent.botones.document.location = "pedidos_pro_bt.asp?mode=search";
	</script>
<%end if
if mode="add" then rst.CancelUpdate
rst.Close%>
<input type="hidden" name="total_paginas" value="<%=EncodeForHtml(total_paginas)%>"/>
</form>
<%'Mostrar la barra de pestañas'
		BarraNavegacion mode
end if
set rstAux = nothing
set rstMM = nothing
set rstAux2 = nothing
set rst = nothing
set rstSelect = nothing
set rstdomi = nothing
set rstAccion = nothing
set rstLimite = nothing
set rstPedirPro = nothing
set rstIvas = nothing
set rstCli = nothing
set rstTMP = nothing
set cnn = nothing
set rstMM = nothing
set rstObtDocCli = nothing
set conn1 = nothing
set command1 = nothing
set rstDet = nothing
connRound.close
set connRound = Nothing
set rsAux=nothing
set rsAux=nothing

%>
</body>
</html>