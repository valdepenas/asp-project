<%@ Language=VBScript %>
<%
''ricardo 5-6-2003 se añade el parametro novei para que en los formatos de impresion no salga el item
''''ricardo 31/7/2003 comprobamos que existe el albaran que se ha pedido ver desde un listado, sino se va al modo add
''MPC 13/12/2008 se añade el combo de la tienda en el caso de tener el mñdulo para tiendas.
''MPC 26/12/2007 se añade el paso de los parñmetros modd, modp y modi a las pñginas de los detalles.
' FLM : 19/01/2009 : Añadir captura de nproveedor por request
' jcg 20/01/2009: Añadida la columna proyecto al proveedor y tratamiento de la misma.  CE 

%>
<!DOCTYPE html PUBLIC "-//W3C/DTD/ XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml1-transitional.dtd" />
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>

</head>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  


<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<% if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->

<!--#include file="albaranes_pro.inc" -->
<!--#include file="compras.inc" -->
<!--#include file="pedbis_pro.inc" -->
<!--#include file="../ventas/documentos.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../js/generic.js.inc"-->

<!--#include file="../js/calendar.inc" -->
<!--#include file="../common/camposperso.inc" -->
<!--#include file="../perso.inc" -->
<!--#include file="../varios2.inc" -->

<!--#include file="../js/animatedCollapse.js.inc"-->

<!--#include file="../js/tabs.js.inc"-->

<!--#include file="../styles/generalData.css.inc"-->

<!--#include file="../styles/Section.css.inc"-->
<!--#include file="../styles/ExtraLink.css.inc"-->

<!--#include file="../styles/Tabs.css.inc" -->

<!--#include file="../styles/formularios.css.inc" -->

<!--#include file="../js/dropdown.js.inc" -->

<!--#include file="../common/poner_cajaResponsive.inc" -->

<!--#include file="../styles/dropdown.css.inc" -->

<!--#include file="../common/albaranes_proActionDrop.inc" -->
<!--#include file="albaranes_pro_linkextra.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<%si_tiene_modulo_21=ModuloContratado(session("ncliente"),"21")
si_tiene_modulo_22=ModuloContratado(session("ncliente"),"22")
''si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
si_tiene_modulo_bierzo=ModuloContratado(session("ncliente"),ModBierzo)
si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
si_tiene_modulo_ccostes=ModuloContratado(session("ncliente"),ModCcostes_Gestion) '**rgu:1/9/2009

p_nalbaran=limpiaCadena(Request.QueryString("nalbaran"))
if p_nalbaran="" then p_nalbaran=limpiaCadena(Request.QueryString("ndoc"))
CheckCadena p_nalbaran

DivisaAlbaran=d_lookup("divisa","albaranes_pro","nalbaran like '" & session("ncliente") & "%' and nalbaran='" & p_nalbaran & "'",session("dsn_cliente"))	
NdecDiAlbaran=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & DivisaAlbaran & "'",session("dsn_cliente"))
if NdecDiAlbaran & "" = "" then
    NdecDiAlbaran=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente"))
end if
    
themeIlion="/lib/estilos/" & folder & "/"
    
%><script language="javascript" type="text/javascript">
      function addVencimiento(nalbaran) {
          if (document.albaranes_pro_pista.tantoVto.value == "") document.albaranes_pro_pista.tantoVto.value = 0;
          if (document.albaranes_pro_pista.fechaVto.value == "" && (document.albaranes_pro_pista.DiasFFVto.value == "" || document.albaranes_pro_pista.DiasFFVto.value == "0")) {
              window.alert("<%=LitErrFechaPago%>");
              return;
          }
          else {
              if (document.albaranes_pro_pista.DiasFFVto.value == "0") {
                  window.alert("<%=LitErrFechaPago%>");
                  return;
              }
          }

          if (isNaN(document.albaranes_pro_pista.tantoVto.value.replace(",", "."))) {
              window.alert("<%=LitErrImportePago%>");
              document.albaranes_pro_pista.tantoVto.value = 0;
              return;
          }
          else {
              if (parseFloat(document.albaranes_pro_pista.tantoVto.value.replace(",", ".")) == 0) {
                  window.alert("<%=LitMsgImportePositivo%>");
                  document.albaranes_pro_pista.tantoVto.value = 0;
                  return;
              }
          }

          if (!cambiarfecha(document.albaranes_pro_pista.fechaVto.value, "Fecha Vencimiento")) return;

          if (!checkdate(document.albaranes_pro_pista.fechaVto)) {
              window.alert("<%=LitMsgFechaFecha%>");
              return;
          }

          //Asignar los valores a los campos del submarco de detalles
          fr_Vencimientos.document.vencimientos_pro_config.h_fecha.value = document.albaranes_pro_pista.fechaVto.value;
          fr_Vencimientos.document.vencimientos_pro_config.h_tanto.value = document.albaranes_pro_pista.tantoVto.value;
          fr_Vencimientos.document.vencimientos_pro_config.h_DiasFF.value = document.albaranes_pro_pista.DiasFFVto.value;
          //Recargar el submarco de pagos a cuenta
          fr_Vencimientos.document.vencimientos_pro_config.action = "vencimientos_pro_config.asp?mode=first_save";
          fr_Vencimientos.document.vencimientos_pro_config.submit();
          //Limpiar los campos del formulario
          var hoy = new Date();
          document.albaranes_pro_pista.fechaVto.value = "";//hoy.getDate() + "/" + (hoy.getMonth()+1) + "/" + hoy.getFullYear();
          document.albaranes_pro_pista.tantoVto.value = "0";
          //Colocar el foco en el campo de cantidad.
          document.albaranes_pro_pista.fechaVto.focus();
          document.albaranes_pro_pista.fechaVto.select();
      }
      function cambiarFecVto() {
          if (document.albaranes_pro_pista.fechaVto.value == "") {
              document.albaranes_pro_pista.DiasFFVto.disabled = false;
          }
          else {
              if (cambiarfecha(document.albaranes_pro_pista.fechaVto.value, 'Fecha Vencimiento')) {
                  document.albaranes_pro_pista.DiasFFVto.vale = "";
                  document.albaranes_pro_pista.DiasFFVto.disabled = true;
              }
          }
      }
      function cambiarDiasFFVto() {
          while (document.albaranes_pro_pista.DiasFFVto.value.search(" ") != -1) {
              document.albaranes_pro_pista.DiasFFVto.value = document.albaranes_pro_pista.DiasFFVto.value.replace(" ", "");
          }
          //if (document.albaranes_pro_pista.DiasFFVto.value=="") document.albaranes_pro_pista.DiasFFVto.value=0;
          if (document.albaranes_pro_pista.DiasFFVto.value == "") {
              //window.alert("<%=LitErrFechaPago%>");
              //return;
              document.albaranes_pro_pista.fechaVto.disabled = false;
          }
          else {
              if (isNaN(document.albaranes_pro_pista.DiasFFVto.value.replace(",", "."))) {
                  window.alert("<%=LitMsgImporteNumerico%>");
                  document.albaranes_pro_pista.DiasFFVto.value = 0;
                  return;
              }
              else {
                  document.albaranes_pro_pista.fechaVto.vale = "";
                  document.albaranes_pro_pista.fechaVto.disabled = true;
              }
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
                      window.alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
                      return false;
                  }
              }
          }
          return true;
      }

      //Desencadena la búsqueda del proveedor cuyo numero se indica
      function TraerProveedor(mode) {
          prov_old = document.albaranes_pro_pista.nproveedor.value;
          cambiar_cliente = "";

          if (confirm("<%=LitCambiarSeriePuedCamPro%>")) cambiar_cliente = 1;
          else cambiar_cliente = 0;

          document.location.href = "albaranes_pro_pista.asp?nproveedor=" + "" + "&mode=" + mode +
              "&nalbaran=" + document.albaranes_pro_pista.h_nalbaran.value +
              "&nalbaran_pro=" + document.albaranes_pro_pista.nalbaran_pro.value +
              "&fecha=" + document.albaranes_pro_pista.fecha.value +
              "&prov=" + prov_old +
              "&serie=" + document.albaranes_pro_pista.serie.value +
              "&valorado=" + document.albaranes_pro_pista.valorado.checked +
              "&nfactura=" + document.albaranes_pro_pista.nfactura.value +
              "&observaciones=" + document.albaranes_pro_pista.observaciones.value
              <%if si_tiene_modulo_proyectos<>0 then%>
                  + "&cod_proyecto=" + document.albaranes_pro_pista.cod_proyecto.value
                  <%end if%>
                      + "&viene=albaranes_pro_pista.asp" +
                      "&cambiar_cliente=" + cambiar_cliente +
                      "&caju=" + document.albaranes_pro_pista.caju.value +
                      "&novei=" + document.albaranes_pro_pista.novei.value +
                      "&incoterms=" + document.albaranes_pro_pista.incoterms.value +
                      "&fob=" + document.albaranes_pro_pista.fob.value +
                      "&modp=" + document.albaranes_pro_pista.modp.value +
                      "&modd=" + document.albaranes_pro_pista.modd.value +
                      "&modi=" + document.albaranes_pro_pista.modi.value +
                      "&s=" + document.albaranes_pro_pista.s.value + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
      }

      function TraerSerie(mode) {
          cambiar_serie = "";
          prov_old = document.albaranes_pro_pista.nproveedor.value;
          if (prov_old != "" && prov_old.length < 5) {
              for (i = prov_old.length; i < 5; i++) {
                  prov_old = "0" + prov_old;
              }
          }
          document.albaranes_pro_pista.nproveedor.value = prov_old;

          if (confirm("<%=LitCambiarProPuedCamSer%>")) cambiar_serie = 1;
          else cambiar_serie = 0;

          document.albaranes_pro_pista.nproveedor.value = "";

          document.location.href = "albaranes_pro_pista.asp?nproveedor=" + prov_old + "&mode=" + mode +
              "&nalbaran=" + document.albaranes_pro_pista.h_nalbaran.value +
              "&nalbaran_pro=" + document.albaranes_pro_pista.nalbaran_pro.value +
              "&fecha=" + document.albaranes_pro_pista.fecha.value +
              "&prov=" + prov_old +
              "&serie=" + document.albaranes_pro_pista.serie.value +
              "&valorado=" + document.albaranes_pro_pista.valorado.checked +
              "&nfactura=" + document.albaranes_pro_pista.nfactura.value +
              "&observaciones=" + document.albaranes_pro_pista.observaciones.value
              <%if si_tiene_modulo_proyectos<>0 then%>
                  + "&cod_proyecto=" + document.albaranes_pro_pista.cod_proyecto.value
                  <%end if%>
                      + "&viene=albaranes_pro_pista.asp" +
                      "&cambiar_serie=" + cambiar_serie +
                      "&caju=" + document.albaranes_pro_pista.caju.value +
                      "&novei=" + document.albaranes_pro_pista.novei.value +
                      "&incoterms=" + document.albaranes_pro_pista.incoterms.value +
                      "&fob=" + document.albaranes_pro_pista.fob.value +
                      "&modp=" + document.albaranes_pro_pista.modp.value +
                      "&modd=" + document.albaranes_pro_pista.modd.value +
                      "&modi=" + document.albaranes_pro_pista.modi.value +
                      "&s=" + document.albaranes_pro_pista.s.value + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
      }

      //Redirecciona a la opcion pulsada en la capa de navegación entre registros
      function Navegar(destino, origen) {
          document.albaranes_pro_pista.action = "albaranes_pro_pista.asp?nalbaran=" + origen + "&donde=" + destino + "&mode=search" + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
          document.albaranes_pro_pista.submit();
      }

      //***************************************************************************
      function Precios() {
          if (isNaN(document.albaranes_pro_pista.dto1.value.replace(",", ".")) || isNaN(document.albaranes_pro_pista.dto2.value.replace(",", ".")) || isNaN(document.albaranes_pro_pista.rf.value.replace(",", ".")) || isNaN(document.albaranes_pro_pista.irpf.value.replace(",", ".")))
              alert("<%=LitMsgDto1Dto2RfNumerico%>");
          else {
              //Preparamos los datos para trabajar***************************************
              dto1SinComas = document.albaranes_pro_pista.dto1.value.replace(",", ".");
              dto2SinComas = document.albaranes_pro_pista.dto2.value.replace(",", ".");
              rfSinComas = document.albaranes_pro_pista.rf.value.replace(",", ".");
              irpfSinComas = document.albaranes_pro_pista.irpf.value.replace(",", ".");

              //TOTAL DESCUENTO**********************************************************
              dto1 = (parseFloat(document.albaranes_pro_pista.importe_bruto.value.replace(",", ".")) * parseFloat(dto1SinComas)) / 100;
              dto2 = ((parseFloat(document.albaranes_pro_pista.importe_bruto.value.replace(",", ".")) - dto1) * parseFloat(dto2SinComas)) / 100;
              dtoTotal = dto1 + dto2;
              c_dtoTotal = dtoTotal.toString();
              document.albaranes_pro_pista.total_descuento.value = parseFloat(c_dtoTotal).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.h_total_descuento.value = document.albaranes_pro_pista.total_descuento.value;

              //BASE IMPONIBLE***********************************************************
              base_imponible = parseFloat(document.albaranes_pro_pista.importe_bruto.value.replace(",", ".")) - parseFloat(document.albaranes_pro_pista.total_descuento.value.replace(",", "."));
              c_base_imponible = base_imponible.toString();
              document.albaranes_pro_pista.base_imponible.value = parseFloat(c_base_imponible).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.h_base_imponible.value = document.albaranes_pro_pista.base_imponible.value;
              //TOTAL IVA****************************************************************
              dto1 = ((parseFloat(document.albaranes_pro_pista.sumadet.value.replace(",", ".")) * parseFloat(dto1SinComas)) / 100);
              dto2 = ((parseFloat(document.albaranes_pro_pista.sumadet.value.replace(",", ".")) - dto1) * parseFloat(dto2SinComas)) / 100;
              dtoTotal = dto1 + dto2;
              total_iva = parseFloat(document.albaranes_pro_pista.sumadet.value) - dtoTotal;
              c_total_iva = total_iva.toString();
              document.albaranes_pro_pista.total_iva.value = parseFloat(c_total_iva).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.h_total_iva.value = document.albaranes_pro_pista.total_iva.value;
              //RECARGO FINANCIERO*******************************************************
              total_rf = (parseFloat(document.albaranes_pro_pista.base_imponible.value) * rfSinComas) / 100;
              c_total_rf = total_rf.toString();
              document.albaranes_pro_pista.total_rf.value = parseFloat(c_total_rf).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.h_total_rf.value = document.albaranes_pro_pista.total_rf.value;
              //RECARGO DE EQUIVALENCIA**************************************************
              total_re = parseFloat(document.albaranes_pro_pista.sumaRE.value);
              c_total_re = total_re.toString();
              document.albaranes_pro_pista.total_re.value = parseFloat(c_total_re).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.h_total_re.value = parseFloat(c_total_re).toFixed(<%=NdecDiAlbaran %>);
              //RETENCIÓN FISCAL*******************************************************
              if (document.albaranes_pro_pista.IRPF_Total.value == "True" || document.albaranes_pro_pista.IRPF_Total.value == "1")
                  baseImp = parseFloat(document.albaranes_pro_pista.base_imponible.value) + parseFloat(document.albaranes_pro_pista.total_iva.value) +
                      parseFloat(document.albaranes_pro_pista.total_re.value) + parseFloat(document.albaranes_pro_pista.total_rf.value);
              else baseImp = document.albaranes_pro_pista.base_imponible.value;
              total_irpf = (parseFloat(baseImp) * irpfSinComas) / 100;
              c_total_irpf = total_irpf.toString();
              document.albaranes_pro_pista.total_irpf.value = parseFloat(c_total_irpf).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.h_total_irpf.value = document.albaranes_pro_pista.total_irpf.value;

              //TOTAL
              total_albaran = parseFloat(document.albaranes_pro_pista.base_imponible.value.replace(",", ".")) + parseFloat(document.albaranes_pro_pista.total_iva.value.replace(",", ".")) + parseFloat(document.albaranes_pro_pista.total_re.value.replace(",", ".")) + parseFloat(document.albaranes_pro_pista.total_rf.value.replace(",", ".")) - parseFloat(document.albaranes_pro_pista.total_irpf.value.replace(",", "."));
              c_total_albaran = total_albaran.toString();
              document.albaranes_pro_pista.total_albaran.value = parseFloat(c_total_albaran).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.h_total_albaran.value = document.albaranes_pro_pista.total_albaran.value;
              //VOLVEMOS A DEJAR LOS DATOS CAMBIADOS COMO ESTABAN************************
              document.albaranes_pro_pista.dto1.value = dto1SinComas;
              document.albaranes_pro_pista.dto2.value = dto2SinComas;
              document.albaranes_pro_pista.rf.value = rfSinComas;
              document.albaranes_pro_pista.irpf.value = irpfSinComas;
          }
      }

      //****************************************************************************
      function CreaPedido(albaran) {
          document.location = "albaranes_pro_pista.asp?nalbaran=" + albaran + "&mode=browse&viene=" + document.albaranes_pro_pista.viene.value + "&seriepedido=" + document.albaranes_pro_pista.seriePB.value + "&almacenSerie=<%=almacenSerie %>&almacenTPV=<%=almacenTPV %>";
          parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=browse";
      }

      //****************************************************************************
      function CancelaPedido(albaran) {
          document.albaranes_pro_pista.action = "albaranes_pro_pista.asp?nalbaran=" + albaran + "&mode=browse" + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
          document.albaranes_pro_pista.submit();
          parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=browse";
      }

      //**********************************************************************************
      function Acaja(nalbaran) {
          if (document.albaranes_pro_pista.impcaja.value == "") document.albaranes_pro_pista.impcaja.value = 0;
          if (isNaN(document.albaranes_pro_pista.impcaja.value.replace(",", "."))) {
              alert("<%=LitMsgImporteNumerico%>");
              return false;
          }
          else {
              if (parseFloat(document.albaranes_pro_pista.impcaja.value.replace(",", ".")) == 0) {
                  alert("<%=LitMsgImporteDisCero%>");
                  return false;
              }
          }
          if (document.albaranes_pro_pista.ncaja.value == "") alert("<%=LitMsgCajaNoNulo%>");
          else {
              if (document.albaranes_pro_pista.i_pago.value == "") alert("<%=LitMsgTipoPagoNoNulo%>");
              else {
                  // FLM : 040309 : Añadir confirm a la hora de incluir en caja.
                  if (!confirm("<%=LitMsgAnotPagadaAlbConfirm%>")) return false;

                  fr_PagosCuenta.document.albaranes_propagos.action = "albaranes_propagos.asp?mode=acaja&ndoc=" + nalbaran + "&impcaja=" + document.albaranes_pro_pista.impcaja.value + "&i_pago=" + document.albaranes_pro_pista.i_pago.value + "&ncaja=" + document.albaranes_pro_pista.ncaja.value;
                  fr_PagosCuenta.document.albaranes_propagos.submit();
                  if (document.getElementById("PAGOS_CUENTA") != null) {
                      if (document.getElementById("PAGOS_CUENTA").style.display == "none") {
                          tier1Menu(PAGOS_CUENTA, document.getElementById("img5"), "<%=oculta%>");
                      }
                  }
              }
          }
      }

      //*****************************************************************************
      //Añade un concepto al albarán
      function addConcepto(nalbaran) {
          if (document.albaranes_pro_pista.descripcion.value == "") {
              alert("<%=LitMsgDesVacia%>");
              return;
          }
          if (isNaN(document.albaranes_pro_pista.pvp.value.replace(",", "."))) {
              window.alert("<%=LitMsgImporteNumerico%>");
              return;
          }

          if (isNaN(document.albaranes_pro_pista.cantidad.value.replace(",", ".")) || isNaN(document.albaranes_pro_pista.descuento.value.replace(",", ".")) || isNaN(document.albaranes_pro_pista.pvp.value.replace(",", "."))) {
              window.alert("<%=LitMsgCanPreDesNumerico%>");
              return;
          }

          //Asignar los valores a los campos del submarco de detalles
          fr_Conceptos.document.albaranes_procon.cantidad.value = document.albaranes_pro_pista.cantidad.value;
          fr_Conceptos.document.albaranes_procon.descripcion.value = document.albaranes_pro_pista.descripcion.value;
          fr_Conceptos.document.albaranes_procon.pvp.value = document.albaranes_pro_pista.pvp.value;
          fr_Conceptos.document.albaranes_procon.descuento.value = document.albaranes_pro_pista.descuento.value;
          fr_Conceptos.document.albaranes_procon.iva.value = document.albaranes_pro_pista.iva.value;
          //Recargar el submarco de detalles
          fr_Conceptos.document.albaranes_procon.action = "albaranes_procon.asp?mode=first_save";
          fr_Conceptos.document.albaranes_procon.submit();
          //Limpiar los campos del formulario
          document.albaranes_pro_pista.cantidad.value = "1";
          document.albaranes_pro_pista.descripcion.value = "";
          document.albaranes_pro_pista.pvp.value = "0";
          document.albaranes_pro_pista.descuento.value = "0";
          document.albaranes_pro_pista.iva.value = document.albaranes_pro_pista.defaultIva.value;
          document.albaranes_pro_pista.importe.value = "0";
          //Colocar el foco en el campo de cantidad.
          document.albaranes_pro_pista.cantidad.focus();
          document.albaranes_pro_pista.cantidad.select();
      }

      //**************************************************************************************************
      //Añade un pago a cuenta.
      function addPago(nalbaran) {
          if (document.albaranes_pro_pista.importePago.value == "") document.albaranes_pro_pista.importePago.value = 0;
          if (document.albaranes_pro_pista.fechaPago.value == "") {
              window.alert("<%=LitErrFechaPago%>");
              return;
          }

          if (!cambiarfecha(document.albaranes_pro_pista.fechaPago.value, "Fecha Pago")) return;

          if (!checkdate(document.albaranes_pro_pista.fechaPago)) {
              window.alert("<%=LitMsgFechaFecha%>");
              return;
          }
          if (isNaN(document.albaranes_pro_pista.importePago.value.replace(",", "."))) {
              window.alert("<%=LitErrImportePago%>");
              return;
          }
          else {
              if (parseFloat(document.albaranes_pro_pista.importePago.value.replace(",", ".")) == 0) {
                  window.alert("<%=LitMsgImportePositivo%>");
                  return;
              }
          }
          if (document.albaranes_pro_pista.descripcionPago.value == "") {
              window.alert("<%=LitMsgDesVacia%>");
              return;
          }
          if (document.albaranes_pro_pista.tipoPago.value == "") {
              window.alert("<%=LitMsgTipoPagoNoNulo%>");
              return;
          }
          //Asignar los valores a los campos del submarco de detalles
          fr_PagosCuenta.document.albaranes_propagos.fecha.value = document.albaranes_pro_pista.fechaPago.value;
          fr_PagosCuenta.document.albaranes_propagos.importe.value = document.albaranes_pro_pista.importePago.value;
          fr_PagosCuenta.document.albaranes_propagos.descripcion.value = document.albaranes_pro_pista.descripcionPago.value;
          fr_PagosCuenta.document.albaranes_propagos.medio.value = document.albaranes_pro_pista.tipoPago.value;
          //Recargar el submarco de pagos a cuenta
          fr_PagosCuenta.document.albaranes_propagos.action = "albaranes_propagos.asp?mode=first_save";
          fr_PagosCuenta.document.albaranes_propagos.submit();
          //Limpiar los campos del formulario
          var hoy = new Date();
          document.albaranes_pro_pista.fechaPago.value = hoy.getDate() + "/" + (hoy.getMonth() + 1) + "/" + hoy.getFullYear();
          document.albaranes_pro_pista.importePago.value = "0";
          document.albaranes_pro_pista.descripcionPago.value = "";
          document.albaranes_pro_pista.tipoPago.value = "";
          //Colocar el foco en el campo de cantidad.
          document.albaranes_pro_pista.fechaPago.focus();
          document.albaranes_pro_pista.fechaPago.select();
      }

      //***************************************************************************************
      //Comprueba si el importe del pago es numerico
      function importepagoComp() {
          if (isNaN(document.albaranes_pro_pista.importePago.value.replace(",", "."))) {
              window.alert("<%=LitErrImportePago%>");
              return;
          }
      }

      //****************************************************************************************
      //Calcula el importe de la línea de detalle del concepto.
      function ImporteDetalle() {
          if (parseFloat(document.albaranes_pro_pista.pvp.value) < 0) {
              window.alert("<%=LitMsgImporteNoNegativo%>");
              document.albaranes_pro_pista.pvp.value = 0;
          }
          if (isNaN(document.albaranes_pro_pista.cantidad.value.replace(",", ".")) || isNaN(document.albaranes_pro_pista.descuento.value.replace(",", ".")) || isNaN(document.albaranes_pro_pista.pvp.value.replace(",", ".")))
              window.alert("<%=LitMsgCanPreDesNumerico%>");
          else {
              if (document.albaranes_pro_pista.pvp.value == "") document.albaranes_pro_pista.pvp.value = 0;
              if (document.albaranes_pro_pista.cantidad.value == "") document.albaranes_pro_pista.cantidad.value = 1;
              if (document.albaranes_pro_pista.descuento.value == "") document.albaranes_pro_pista.descuento.value = 0;
              pvpSinComas = document.albaranes_pro_pista.pvp.value.replace(",", ".");
              cantidadSinComas = document.albaranes_pro_pista.cantidad.value.replace(",", ".");
              dtoSinComas = document.albaranes_pro_pista.descuento.value.replace(",", ".");
              pelas = parseFloat(cantidadSinComas) * parseFloat(pvpSinComas);
              pelas_descuento = (pelas * parseFloat(dtoSinComas)) / 100;
              importe = pelas - pelas_descuento;
              c_importe = importe.toString();
              document.albaranes_pro_pista.cantidad.value = cantidadSinComas;
              document.albaranes_pro_pista.descuento.value = dtoSinComas;
              document.albaranes_pro_pista.importe.value = parseFloat(c_importe).toFixed(<%=NdecDiAlbaran %>);
              document.albaranes_pro_pista.pvp.value = pvpSinComas;
          }
      }
      /*
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
      
      function keyPressed(e)
      {
          var keycode = e.keyCode;
          if (keycode==<%=TeclaGuardar%>)
          { //CTRL+S
              if (document.albaranes_pro_pista.mode.value=="add" || document.albaranes_pro_pista.mode.value=="edit")
              {
                  if (document.albaranes_pro_pista.fecha.value=="")
                  {
                      window.alert("<%=LitMsgFechaNoNulo%>");
                      return;
                  }
      
                  if (!cambiarfecha(document.albaranes_pro_pista.fecha.value,"FECHA ALBARAN")) return;
      
                  if (!checkdate(document.albaranes_pro_pista.fecha))
                  {
                     window.alert("<%=LitMsgFechaFecha%>");
                     return ;
                  }
                  if (document.albaranes_pro_pista.serie.value=="")
                  {
                      window.alert("<%=LitMsgSerieNoNulo%>");
                      return;
                  }
                  if (document.albaranes_pro_pista.divisa.value=="")
                  {
                      window.alert("<%=LitMsgDivisaNoNulo%>");
                      return;
                  }
                  if (document.albaranes_pro_pista.nproveedor.value=="")
                  {
                      window.alert("<%=LitMsgProveedorNoNulo%>");
                      return;
                  }
      
                  if (comp_car_ext(document.albaranes_pro_pista.nalbaran_pro.value,1)==1)
                  {
                      window.alert("<%=LitMsgAlbpDesCarNoVal%>");
                      return;
                  }
              	
                  // AMP: comprobacion campo factor de cambio.
                  factcambio=document.albaranes_pro_pista.nfactcambio.value.replace(",","."); 		
                  if (!/^([0-9])*[.]?[0-9]*$/.test(factcambio))
                  { 
                      alert("<%=LitMsgFactCambioI%>"); 
                      return false;
                  }
                  if (document.albaranes_pro_pista.nfactcambio.value=="")
                  {
                      alert("<%=LitMsgFactCambioI%>"); 
                      return false;
                  }
      
                  // JMA 20/12/04. Campos personalizables
                  if (document.albaranes_pro_pista.si_campo_personalizables.value==1)
                  {
                      num_campos=document.albaranes_pro_pista.num_campos.value;
                      respuesta=comprobarCampPerso("",num_campos,"albaranes_pro");
                      if(respuesta!=0)
                      {
                          titulo="titulo_campo" + respuesta;
                          tipo="tipo_campo" + respuesta;
                          titulo=document.albaranes_pro_pista.elements[titulo].value;
                          tipo=document.albaranes_pro_pista.elements[tipo].value;
                          if (tipo==4) nomTipo="<%=LitTipoNumerico%>";
                          else if (tipo==5)
                          {
                              nomTipo="<%=LitTipoFecha%>";
                          }
      
                          window.alert("<%=LitMsgCampo%> " + titulo + " <%=LitMsgTipo%> " + nomTipo);
                          return false;
                      }
                  }
      	
                  //ricardo 15-1-2008 si editamos , no podremos borrar el nalbaran_pro		
                  if (document.albaranes_pro_pista.mode.value=="edit" && document.albaranes_pro_pista.nalbaran_pro.value=="") {
                      window.alert("<%=LitMsgNalbaranProNoNulo%>");
                      return false;
                  }
              	
                  switch (document.albaranes_pro_pista.mode.value)
                  {
                      case "add":
                          document.albaranes_pro_pista.action="albaranes_pro.asp?mode=first_save"+"&almacenSerie=<%=almacenSerie %>&almacenTPV=<%=almacenTPV %>";
                          break;
      
                      case "edit":
                          document.albaranes_pro_pista.action="albaranes_pro.asp?mode=save&ndoc=" + document.albaranes_pro_pista.h_nalbaran.value+"&almacenSerie=<%=almacenSerie %>&almacenTPV=<%=almacenTPV %>";
                          break;
                  }
      
                  //ricardo 25/4/2003 para que salga el mensaje de que se ha cambiado las propiedades del documento
                  // y que puede afectar al importe de los detalles
                  nempresa="<%=session("ncliente")%>";
                  recalcular_importes=1;
                  if (document.albaranes_pro_pista.mode.value=="edit")
                  {
                      if (document.albaranes_pro_pista.h_nproveedor.value!=(nempresa + document.albaranes_pro_pista.nproveedor.value) ||
                          document.albaranes_pro_pista.h_fecha.value!=document.albaranes_pro_pista.fecha.value ||
                          document.albaranes_pro_pista.h_divisa.value!=document.albaranes_pro_pista.olddivisa.value)
                      {
                          if (window.confirm("<%=LitMsgCamPropDocCamPrec%>")==false) recalcular_importes=0;
                      }
                  }
                  document.albaranes_pro_pista.action=document.albaranes_pro_pista.action + "&recalcular_importes=" + recalcular_importes;
                  document.albaranes_pro_pista.submit();
                  parent.botones.document.location="../compras/albaranes_pro_bt.asp?mode=browse";
              }
              else
              { //Mode=browse.
                  que_pantalla=getTabsSelected();
                  //Comprobamos si estamos añadiendo conceptos.
                  if (que_pantalla==1) addConcepto(document.albaranes_pro_pista.h_nalbaran.value);
      
                  //Comprobamos si estamos añadiendo conceptos.
                  if (que_pantalla==2) addPago(document.albaranes_pro_pista.h_nalbaran.value);
              }
          }
      }
      */
      //***************************************************************************
      function MasDet(sentido, lote, firstReg, lastReg, campo, criterio, texto, firstRegAll, lastRegAll) {
          fr_Detalles.document.albaranes_prodet.action = "albaranes_prodet_pista.asp?mode=browse&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&firstReg=" + firstReg + "&lastReg=" + lastReg + "&firstRegAll=" + firstRegAll + "&lastRegAll=" + lastRegAll + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
          fr_Detalles.document.albaranes_prodet.submit();
      }

      //*** i AMP Nueva función cambiar divisa con factor de cambio incorporado.
      var ret_tra = "";
      var ret_tra2 = "";
      function cambiardivisa(mBase) {
          document.albaranes_pro_pista.h_divisa.value = document.albaranes_pro_pista.divisa.value;

          var divisa = document.albaranes_pro_pista.divisa.value;
          if (divisa == mBase) {
              parent.pantalla.document.getElementById("tdfactcambio").style.display = "none";
              parent.pantalla.document.albaranes_pro_pista.nfactcambio.value = "1";
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
          parent.pantalla.document.albaranes_pro_pista.h_divisa.value = divisa;
          parent.pantalla.document.albaranes_pro_pista.divisafc.value = divisa;
      }

      function handleHttpResponseFC() {
          if (httpFC.readyState == 4) {
              if (httpFC.status == 200) {
                  if (httpFC.responseText.indexOf('invalid') == -1) {
                      // Armamos un array, usando la coma para separar elementos
                      results = httpFC.responseText;
                      enProcesoFC = false;
                      ret_tra = unescape(results);
                      spfc = ret_tra.split(";");
                      factcambio = spfc[0];
                      parent.pantalla.document.albaranes_pro_pista.nfactcambio.value = factcambio;
                      ret_tra2 = "";
                      var divisa = document.albaranes_pro_pista.divisa.value;
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
          ok = 1;
          numero = document.albaranes_pro_pista.nfactcambio.value;
          document.albaranes_pro_pista.nfactcambio.value = numero.replace(",", ".")
          numero2 = document.albaranes_pro_pista.nfactcambio.value;
          if (!/^([0-9])*[.]?[0-9]*$/.test(numero2)) {
              alert("<%=LitMsgFactCambioI%>");
              ok = 0;
          }
          if (document.albaranes_pro_pista.nfactcambio.value == "") {
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
              }
              else {
                  jQuery("#frDetalles").attr("height", dir_default);
                  jQuery("#frConceptos").attr("height", dir_default);
                  jQuery("#frPagosCuenta").attr("height", dir_default);
              }
          }
          else {
              jQuery("#frDetalles").attr("height", dir_default);
              jQuery("#frConceptos").attr("height", dir_default);
              jQuery("#frPagosCuenta").attr("height", dir_default);
          }
      }

      function RoundNumValue(obj, dec) {
          obj.value = obj.value.replace(',', '.');
          var valor = parseFloat(obj.value);
          if (valor != 0) obj.value = valor.toFixed(dec);
      }

      function isTextSelected(input) 
      {
          if (typeof input.selectionStart == "number") {
              return input.selectionStart == 0 &&
              input.selectionEnd == input.value.length;
          } else if (typeof document.selection != "undefined") {
              input.focus();
              return document.selection.createRange().text == input.value;
          }
      }

      jQuery(window).resize(function () { Redimensionar(); });
</script>
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
<%'Everilion Interface Timing%>
<script language="javascript" type="text/javascript" src="/lib/js/InterfaceLoadTime.js"></script>
<script language="javascript" type="text/javascript">

    window.onload = function () {

        self.status = '';
    <%if tracetime> 0 then %>
            StoreTiming("<%=CarpetaProduccion%>", <%=tracetime %>, "<%=enc.EncodeForJavascript(Request.QueryString("mode")&"")%>", "<%=enc.EncodeForJavascript(session("usuario"))%>", "<%=enc.EncodeForJavascript(session("ncliente"))%>", window.location.pathname);
    <%end if %>
 }

</script>
<body class="BODY_ASP">
<%
'******************************************************************************
'Crea la tabla que contiene la barra de grupos de datos.
sub BarraNavegacion(modo)
    if modo="add" or mode="edit" then%>
        <script language="javascript" type="text/javascript">
            jQuery("#GENERAL_DATA").show();
        </script>	
	<%else%>
		<script language="javascript" type="text/javascript">
            jQuery("#GENERAL_DATA").hide();
        </script>
    <%end if
	if modo<>"add" and modo<>"edit" then%>	
		<script type="text/javascript" language="javascript">
            jQuery("#FINANCIAL_DATA").show();
            jQuery(window).load(function () {
                Redimensionar();
                try {
                    if (document.getElementById("frDetallesIns").style.display != "none") {
                        fr_DetallesIns.document.albaranes_prodetins_pista.cantidad.focus();
                        fr_DetallesIns.document.albaranes_prodetins_pista.cantidad.select();
                    }
                }
                catch (e) {
                }
            });
		</script>
	<%end if
end sub

'******************************************************************************
'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(nalbaran,nserie)
    nproveedor_aux=request.form("nproveedor")
    if nproveedor_aux & "">"" and len(nproveedor_aux)<=5 then
        nproveedor_aux=session("ncliente") & nproveedor_aux
    end if

	if nalbaran="" then
		'Crear un nuevo registro.
		rst.AddNew
		'******************** Manejo de domicilios
		Dom=Domicilios("COMPRAS","ALB_ENV_PROV",nproveedor_aux,rst)
	end if
	FechaDoc=rst("fecha")
	ProvDoc=rst("nproveedor")
	DtoGeneral=null_z(rst("descuento"))
	DtoGeneral2=null_z(rst("descuento2"))
	'Asignar los nuevos valores a los campos del recordset.
	rst("valorado")=nz_b(Request.Form("valorado"))
	rst("serie")=Nulear(Request.Form("serie"))
	'Si se le cambia el cliente al proveedor, hay que capturar sus direcciones
	if rst("nproveedor")<>Nulear(nproveedor_aux) and nalbaran<>"" then
		Dom=Domicilios("COMPRAS","ALB_ENV_PROV",nproveedor_aux,rst)
	end if
	'--------------------------------------------------------------------------
	cambio_proveedor = false
	if rst("nproveedor")<>nproveedor_aux then cambio_proveedor=true
	
	ndec=d_lookup("ndecimales", "divisas", "codigo like '"&session("ncliente")&"%' and codigo = '"&request.form("h_divisa")&"'", session("dsn_cliente"))
	rst("nproveedor")=Nulear(nproveedor_aux)
	rst("descuento")=miround(Null_z(request.form("dto1")),decpor)
	
	rst("descuento2")=miround(Null_z(request.form("dto2")),decpor)
	rst("total_descuento")=miround(Null_z(request.form("h_total_descuento")),ndec)
	rst("base_imponible")=miround(Null_z(request.form("h_base_imponible")),ndec)
	rst("total_iva")=miround(Null_z(request.form("h_total_iva")),ndec)
	rst("recargo")=miround(Null_z(request.form("rf")),decpor)
	rst("total_recargo")=miround(Null_z(request.form("h_total_rf")),ndec)
	rst("irpf")=miround(Null_z(request.form("irpf")),decpor)
	rst("IRPF_Total")	= nz_b(Request.Form("IRPF_Total"))
	rst("total_irpf")=miround(Null_z(request.form("h_total_irpf")),ndec)
	rst("total_re")=miround(Null_z(request.form("h_total_re")),ndec)
	rst("total_albaran")=miround(Null_z(request.form("h_total_albaran")),ndec)
	rst("facturado")=0
	rst("ahora")=0
	rst("forma_pago")=Nulear(request.form("forma_pago"))
	rst("tipo_pago")=Nulear(request.form("tipo_pago"))
	rst("nfactura")=Nulear(request.form("numfactura"))
	rst("observaciones")=Nulear(request.form("observaciones"))
	rst("cod_proyecto")=Nulear(request.form("cod_proyecto"))
	rst("ncuenta")=Nulear(request.form("ncuentacargo"))
    
   '** ALBARAN TIPO PISTA
    rst("tipo_albaran") = session("ncliente") + "00002"

	'FLM:120309: añado la cuenta de pago al proveedor para dar soporte a la notrma 34.
	rst("ncuenta_pro")=Nulear(request.form("ncuenta_pro"))
	if(rst("ncuenta_pro")&""<>"") then 
        banco=d_lookup("Entidad","bancos","codigo='" & mid(trim(rst("ncuenta_pro")),5,4) & "'",DsnIlion)
    else
        banco=null
    end if
    rst("banco")=iif(banco="",NULL,trim(banco))
	
	rst("incoterms")=nulear(request.form("incoterms"))
	rst("fob")=nulear(request.form("fob"))
	''MPC 13/12/2007 guardar el campo tienda en los albaranes de proveedores
	rst("tienda")=nulear(request.form("tienda"))

	rst("divisa")= Nulear(request.form("divisa"))
	rst("fecha")=Nulear(Request.Form("fecha"))
	rst("divisa")=Nulear(request.form("h_divisa"))
	rst("factcambio")=miround(Nulear(limpiaCadena(request.form("nfactcambio"))),DEC_PREC) '*** AMP 

	if nalbaran="" then
		'Obtener el último nº de albaran de la tabla series.
		strselect="select almacen,contador,ultima_fecha from series where nserie='" & nserie & "'"
		rstAux.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

        'mmg:calculamos el almacen por defecto de la serie 
        if rstAux.eof then
	        almacenSerie= ""
        else
            'comprobamos si el almacen esta dado de baja
            rstMM.cursorlocation=3
            rstMM.Open "select codigo from almacenes where codigo='" & rstAux("almacen") & "' and isnull(fbaja,'')=''",session("dsn_cliente")
		    if rstMM.eof then
	            almacenSerie= ""
            else
	            almacenSerie= rstAux("almacen")
	        end if
        end if
        
		num=rstAux("contador")+1
		num=string(6-len(cstr(num)),"0") + cstr(num)

		'Actualizar el nº de proveedor de CONFIGURACION.
		rstAux("contador")=rstAux("contador")+1
		rstAux("ultima_fecha")=date
		rstAux.Update
		rstAux.Close

		SigDoc = nserie & right(Nulear(Request.Form("fecha")),2) & num
		rst("nalbaran")=SigDoc
	end if

	''ricardo 20/2/2003
	''si se deja vacio el campo,se rellanara con el nalbaran
		if Request.Form("nalbaran_pro") & "">"" then
			   rst("nalbaran_pro") = Nulear(Request.Form("nalbaran_pro"))
		else
			   rst("nalbaran_pro") = trimCodEmpresa(SigDoc)
		end if
	'''''''

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

	rst.Update

	''ricardo 28/4/2003 si el usuario ha querido recalcular los importes al cambiar las propiedades de la cabecera
	han_cambiado_importes_proveedor=0
	if limpiaCadena(request.querystring("recalcular_importes"))="1" then
		'Detectamos un cambio de proveedor
		if cambio_proveedor=true then
		    TmpIvaProveedor=d_lookup("iva","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))
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

			han_cambiado_importes_proveedor=1
			set rstMiProveer = Server.CreateObject("ADODB.Recordset")
			'recorremos los detalles modificando precios
			rstaux.open "select * from detalles_alb_pro with(updlock) where nalbaran='" & nalbaran & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			while not rstAux.eof
                rstMiProveer.cursorlocation=3
				rstMiProveer.open "select * from proveer with(nolock) where nproveedor='" & rst("nproveedor") & "' and articulo='" & rstaux("referencia") & "'",session("dsn_cliente")
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
				rstMiProveer.close
			wend
			rstAux.close
			Set rstMiProveer=nothing

			'recorremos ahora los conceptos haciendo cambio de divisa si hace falta
			rstaux.open "select * from conceptos_alb_pro with(updlock) where nalbaran='" & nalbaran & "' order by nconcepto",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			while not rstAux.eof
				TmpPVP=CambioDivisa(rstAux("pvp"),rst("divisa"),request.form("h_divisa"))
				rstAux("pvp")=TmpPVP
				TmpPVP=TmpPVP*rstAux("cantidad")
				TmpPVP	= TmpPVP*(100-null_z(rstAux("descuento")))/100
				'TmpPVP	= TmpPVP*(100-null_z(rst("descuento2")))/100
				rstAux("importe")=miround(TmpPVP,2)
			    rstAux("iva")=TmpIva
			    rstAux("re")=TmpRe
				rstAux.update
				rstAux.movenext
			wend
			rstAux.close
		end if
	end if

	if nalbaran="" then
		PreciosAlb SigDoc
	else
		PreciosAlb nalbaran
	end if
	ActualizaCostes iif(nalbaran="",SigDoc,nalbaran),"DOCUMENTO","ALBARAN DE PROVEEDOR","",ProvDoc,0,FechaDoc,0,0,DtoGeneral,DtoGeneral2,false,session("dsn_cliente")
end sub

'******************************************************************************
'Elimina los datos del registro cuando se pulsa BORRAR.
sub BorrarRegistro(nalbaran, nalbaran_pro)
    'i *** AMP 17092010 -- Restricciones de borrado si existe lote asignado 
    rstAux.Open "select * from lotes_entrada with(nolock) where nalbaran='" & nalbaran& "'" ,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	noborrar=false ' valor inicial
	nalbaranLote=""
	if not rstAux.eof then 'Si no existe lote asociado a la línea saltar restricción de borrado	
	    while not rstAux.EOF		    
	        if rstAux("nalbaran")>"" then nalbaranLote=rstAux("nalbaran")
            ndetLote = rstAux("ndet")
           
            if ndetLote>"" then
                strSelect="ComprobarUsoLote '"&session("nempresa")&"','"&nalbaran&"','"&ndetLote&"',''"
                rstAux2.Open strSelect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                if not rstAux2.eof then                             
                    usoLote=rstAux2(0)
                    rstAux2.close    		               
                    if usoLote=1  then 'si se utiliza lote y no existe un albaran vinculado a este lote --> no borrar línea 
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
               document.albaranes_pro_pista.action = "albaranes_pro_pista.asp?nalbaran=<%=enc.EncodeForJavascript(nalbaran)%>&mode=browse&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
               document.albaranes_pro_pista.submit();
               parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=browse";
	     </script><%
	else
	    TieneFactura="NO"
        rstAux.cursorlocation=3
	    rstAux.Open "select nfactura from albaranes_pro with(nolock) where nalbaran='" & nalbaran & "'",session("dsn_cliente")
	    if not rstAux.EOF then
		    if not isnull(rstAux("nfactura")) then
			    TieneFactura="SI"
			    factura=rstAux("nfactura")
		    end if
	    end if
	    rstAux.close

	    if TieneFactura="NO" then
            ''ricardo 17-3-2006 a partir de esta fecha se borrara con procedimiento
	        set conn = Server.CreateObject("ADODB.Connection")
	        set command =  Server.CreateObject("ADODB.Command")

	        conn.open session("dsn_cliente")
	        command.ActiveConnection =conn
	        command.CommandTimeout = 0
	        command.CommandText="BorrarDocumento"
	        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	        command.Parameters.Append command.CreateParameter("@ndocumento",adVarChar,adParamInput,20,nalbaran)
	        command.Parameters.Append command.CreateParameter("@tipo_documento",adVarChar,adParamInput,50,"ALBARAN DE PROVEEDOR")
	        command.Parameters.Append command.CreateParameter("@result",adInteger,adParamOutput)
	        command.Execute,,adExecuteNoRecords
	        resultado=command.Parameters("@result").Value

	        conn.close
	        set command=nothing
	        set conn=nothing
	    else%>
		   <script language="javascript" type="text/javascript">
               window.alert("<%=LitMsgBorrarAlbaran%> <%=trimCodEmpresa(factura)%>");
               document.location = "albaranes_pro_pista.asp?nalbaran=<%=enc.EncodeForJavascript(nalbaran)%>&mode=browse" + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
               parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=browse";
		    </script>
	    <%end if
    end if
end sub

function CerrarTodo()
set rstMiProveer = nothing
set conn = nothing
set command = nothing
set rsTPV = nothing
set rstAux = nothing
set rstMM = nothing
set rstAux2 = nothing
set rstAux3 = nothing
set rstProveedor = nothing
set rst = nothing
set rstPed = nothing
set rstDetAlb = nothing
set rstIvas = nothing
set rstSelect = nothing
set rstDomi = nothing	
set rstObtSer = nothing
set rstObtDocCli = nothing
set rstMM = nothing
set rstMiProveer = nothing
set rstVert = nothing
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

'*****************************************************************************'
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************'
'*****************************************************************************'
const borde=0
set connRound = Server.CreateObject("ADODB.Connection")
connRound.open dsnilion
	%>
    <form name="albaranes_pro_pista" method="post">
    
    <%
    PintarCabecera "Albaranes_pro_pista.asp"
    ' Ocultar detalle de las facturas si se da el caso
    'ebf: 23/6/2009 Estaba al principio de la página, se pasa aqui para que no se haga nada hasta comprobar
    'que el acceso a la página esta permitido
    'mmg: variables para obtener los almacenes por defecto
    dim almacenSerie
    dim almacenTPV

    linea1=session("f_tpv")
    linea2=session("f_caja")
    strconn=session("dsn_cliente")
        
    'Calculamos el almacen por defecto del TPV
    set rsTPV = Server.CreateObject("ADODB.Recordset")            
    	
    cadena= "select c.almacen from tpv a with(nolock), cajas b with(nolock), tiendas c with(nolock), almacenes alm with(nolock) where a.caja=b.codigo and b.tienda=c.codigo and tpv='" +linea1 +"' and b.codigo='" +linea2+"' and alm.codigo=c.almacen and isnull(alm.fbaja,'')=''"
    rsTPV.cursorlocation=3
    rsTPV.Open cadena,session("dsn_cliente")
    if rsTPV.eof then
	    almacenTPV= ""
    else
	    almacenTPV= rsTPV("almacen")
    end if
    rsTPV.close    
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

    Dim vencesp
    ObtenerParametros("albaranes_pro_det")
    ''response.write("el vencesp es-" & vencesp & "-<br>")

	' Cursores'
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstMM = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rstAux3 = Server.CreateObject("ADODB.Recordset")
	set rstProveedor = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstPed = Server.CreateObject("ADODB.Recordset")
	set rstDetAlb = Server.CreateObject("ADODB.Recordset")
	set rstIvas = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstDomi = Server.CreateObject("ADODB.Recordset")

	'Leer parámetros de la página'
	mode=Request.QueryString("mode")
	p_nalbaran_pro = limpiaCadena(request.querystring("nalbaran_pro"))
	if p_nalbaran_pro="" then
	    p_nalbaran_pro = limpiaCadena(request.form("nalbaran_pro"))
	end if
	if p_nalbaran_pro="" then
	    p_nalbaran_pro=d_lookup("nalbaran_pro","albaranes_pro","nalbaran='" & p_nalbaran & "'", session("dsn_cliente"))
	end if
	p_nproveedor=limpiaCadena(Request.QueryString("nproveedor"))
	if p_nproveedor="" then
		p_nproveedor=limpiaCadena(request.form("nproveedor"))
	end if
	if p_nproveedor="" then
		p_nproveedor=limpiaCadena(request.form("h_nproveedor"))
	end if
	if p_nproveedor & "">"" and len(p_nproveedor)<=5 then
		p_nproveedor=session("ncliente") & p_nproveedor
	end if

	p_fecha=limpiaCadena(Request.QueryString("fecha"))
	p_fechaR=limpiaCadena(Request.form("fecha"))
	
	''MPC 26/12/2007
	modp=limpiaCadena(request.querystring("modp"))
	if modp="" then modp=limpiaCadena(request.form("modp"))

	modd=limpiaCadena(request.querystring("modd"))
	if modd="" then modd=limpiaCadena(request.form("modd"))

    modi=limpiaCadena(request.querystring("modi"))
	if modi="" then modi=limpiaCadena(request.form("modi"))

' >>> MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compras.
'					 bloque 1/4 en albaranes_pro.asp

	s=limpiaCadena(request.querystring("s"))
	if s="" then s=limpiaCadena(request.form("s"))
	s=preparar_lista(s)

' <<< MCA 11/04/05

	p_serie=limpiaCadena(Request.QueryString("serie"))
    if p_serie&"" = "" then
    	p_serie=limpiaCadena(Request.form("serie"))
    end if

	'mmg >> cambios serie
	if p_serie="" and mode="add" then
		'Obtener la serie por defecto'
		campo_ser=""
		campo_ser="serieAlbPro"

	    'Poner por defecto la serie correspondiente a la tienda donde esta el tpv y caja del fichero cetel.tpv

	    linea1=session("f_tpv")
	    linea2=session("f_caja")
	    linea3=session("ncliente")
	    if linea1<>"" and linea2<>"" and linea3<>"" and campo_ser & "">"" then
		    set rstObtSer = Server.CreateObject("ADODB.Recordset")
		    strSelect = "select c." & campo_ser & " from tpv a, cajas b, tiendas c where a.caja=b.codigo and b.tienda=c.codigo and tpv='" & linea1 & "' and b.codigo='" & linea2 & "'"
		    rstObtSer.cursorlocation=3
		    rstObtSer.open strSelect,session("dsn_cliente")
		    if linea3=session("ncliente") then
			    if rstObtSer.eof then
				    serie_a_devolver = ""
			    else
				    serie_a_devolver=rstObtSer(campo_ser)
			    end if
			    rstObtSer.close
		    else
			    serie_a_devolver= ""
		    end if
		    set rstObtSer=nothing
	    end if
	    p_serie=serie_a_devolver
		if p_serie & ""="" then
			p_serie=d_lookup("nserie","series","tipo_documento='ALBARAN DE PROVEEDOR' and nserie like '" & session("ncliente") & "%' and pordefecto=1", session("dsn_cliente"))
		end if
	end if


'**rgu 2/9/2009    
    if p_serie>"" and mode="add" then
        p_tienda=d_lookup("tienda","series","nserie like '"&session("ncliente")&"%' and nserie='"&p_serie&"' ", session("dsn_cliente")) '**rgu 2/9/2009
    end if
'**rgu    

	if request.querystring("cod_proyecto")>"" then
		tmp_cod_proyecto=limpiaCadena(request.querystring("cod_proyecto"))
	else
		tmp_cod_proyecto=limpiaCadena(request.form("cod_proyecto"))
	end if

	p_valorado=limpiaCadena(Request.QueryString("valorado"))
	viene=limpiaCadena(request.querystring("viene"))
	if viene="" then viene=limpiaCadena(request.form("viene"))
	if viene="" then viene="albaranes_pro_pista.asp"
	if viene="cancelar" then p_nalbaran_pro=""

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

	provR=limpiaCadena(request.querystring("prov"))
	'FLM : 19/01/2009 : Añadir captura de nproveedor
	if provR="" then	    
	    TraerProveedor =limpiaCadena(request("nproveedor"))  
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
	
	'*** AMP 
	if Request.QueryString("divisafc")>"" then
		tmpdivisafc=limpiaCadena(Request.QueryString("divisafc"))
	elseif Request.form("divisafc")>"" then
		tmpdivisafc=limpiaCadena(Request.form("divisafc"))
	end if	

	'JMA 20/12/04. Copiar campos personalizables de los proveedores'
	redim tmp_lista_valores(10)
	for ki=1 to 10
		tmp_lista_valores(ki)=""
	next
	'JMA 20/12/04. FIN Copiar campos personalizables de los proveedores'

	''JMA 20-12-2004 si existen campos personalizables con titulo no nulo saldrán los campos personalizables
	si_campo_personalizables=0
    rst.cursorlocation=3
	rst.open "select ncampo from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and titulo is not null and titulo<>'' and ncampo like '" & session("ncliente") & "%'",session("dsn_cliente")
	if not rst.eof then
		si_campo_personalizables=1
	else
		si_campo_personalizables=0
	end if
	rst.close
	%><input type="hidden" name="si_campo_personalizables" value="<%=enc.EncodeForHtmlAttribute(si_campo_personalizables)%>"/><%
	''JMA 20-12-2004 FIN si existen campos personalizables con titulo no nulo saldrán los campos personalizables

	''JMA 20-12-2004 añadir campos personalizables a albaranes_pro
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
			rstAux2.open "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from albaranes_pro as p with(nolock) where p.nalbaran='" & p_nalbaran & "'",session("dsn_cliente")
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
	
	'**RGU 2/9/2009: Se comenta por que no se ejecuta nunca. el si_tiene_modulo_tiendas no coge ningun valor y no cumple la condicion <>0
	'if mode="add" and si_tiene_modulo_tiendas<>0 then
    '    rstAux.open "select tienda from cajas with(nolock) where codigo='"&session("f_caja")&"' and codigo like '"&session("ncliente")&"%' ", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    '    tmp_tienda=""
    '    if not rstAux.eof then
    '        tmp_tienda=rstAux("tienda")
    '    end if
    '    rstAux.close
    'end if
    if tmp_tienda&""="" and p_tienda>"" then
        tmp_tienda=p_tienda
    end if
	''JMA 20-12-2004 añadir campos personalizables a albaranes_pro
	
	campo    = limpiaCadena(request.QueryString("campo"))
	if campo & ""="" then
	    campo = Request.Form("campo")
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
	    sentido=Request.form("sentido")
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
	    npagina=Request.form("npagina")
	end if
	if npagina & ""="" then npagina=0
	
    ''ricardo 3-11-2011 si no tiene acceso a la opcion de almacenes , se quitara dicho campo
    si_tiene_acceso_almacenes=1
    rstAux2.Open "exec ContractedItem '" & session("ncliente") & "','" & replace(OBJAlmacenes,"'","''") & "'", dsnilion
    if not rstAux2.eof then
        if rstAux2(0)=1 then
            si_tiene_acceso_almacenes=1
        else
            si_tiene_acceso_almacenes=0
        end if
    end if
    rstAux2.close%>
    
	<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
	<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
	<input type="hidden" name="caju" value="<%=enc.EncodeForHtmlAttribute(caju)%>"/>
	<input type="hidden" name="novei" value="<%=enc.EncodeForHtmlAttribute(novei)%>"/>
	<input type="hidden" name="modp" value="<%=enc.EncodeForHtmlAttribute(modp)%>"/>
	<input type="hidden" name="modd" value="<%=enc.EncodeForHtmlAttribute(modd)%>"/>
	<input type="hidden" name="modi" value="<%=enc.EncodeForHtmlAttribute(modi)%>"/>
	<input type="hidden" name="campo" value="<%=enc.EncodeForHtmlAttribute(campo)%>"/>
	<input type="hidden" name="texto" value="<%=enc.EncodeForHtmlAttribute(texto)%>"/>
	<input type="hidden" name="lote" value="<%=enc.EncodeForHtmlAttribute(lote)%>"/>
	<input type="hidden" name="criterio" value="<%=enc.EncodeForHtmlAttribute(criterio)%>"/>
    <%' >>> MCA 11/04/05: Añadir parámetro de usuario con la(s) serie(s) a los documentos de compras.
    '					bloque 2/4 en albaranes_pro.asp%>
	<input type="hidden" name="s" value="<%=s%>"/>
    <%if p_nalbaran & "">"" then
		if comprobar_LS(s,mode,p_nalbaran,"ALBARANES_PRO")=0 then%>
			<script language="javascript" type="text/javascript">
                window.alert("<%=LitMsgDocNoPermAcc%>");
                document.albaranes_pro_pista.action = "albaranes_pro_pista.asp?nalbaran=&mode=add" + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
                document.albaranes_pro_pista.submit();
                parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=add" + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
			</script>
			<%CerrarTodo()
			response.end
		end if
	end if

' <<< MCA 11/04/05 :

   'Se crea un select para que el usuario seleccione la serie del pedido bis si fuera necesario'
    if mode="pedirserie" then
  	   rst.Open "select p.*,pv.razon_social from albaranes_pro as p,proveedores as pv where pv.nproveedor=p.nproveedor and nalbaran='" & p_nalbaran & "' and pv.nproveedor like '" & session("ncliente") & "%' ", session("dsn_cliente"),adOpenKeyset,adLockOptimistic%>
	   <span id="SeriePedBIS" style="display:" >
  	   <p align="center"><table border='0' cellspacing="3" cellpadding="3">
  	   <%DrawFila color_fondo
	      DrawCeldaSpan "ENCABEZADOC","","",0,LitSeriePedBis,2
	   CloseFila
	   DrawFila color_blau
	      SeriePedidoB=d_lookup("distinct serie","pedidos_pro","",session("dsn_cliente"))
		  DrawCelda "CELDARIGHT","","",0,LitSerie & " : "
          rstAux.cursorlocation=3
		  rstAux.open "select * from series with(nolock) where tipo_documento='PEDIDO A PROVEEDOR' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente")
		  DrawSelectCelda "CELDA","","",0,"","seriePB",rstAux,SeriePedidoB,"nserie","nserie","",""
		  rstAux.close
	   CloseFila
	   DrawFila color_blau%>
	      <td class="CELDA"><input class="CELDA" type="Button" name="SeriePedidoBis" value="Aceptar" onclick="CreaPedido('<%=enc.EncodeForJavascript(p_nalbaran)%>');"/></td>
		  <td class="CELDA"><input class="CELDA" type="Button" name="CSeriePedidoBis" value="Cancelar" onclick="CancelaPedido('<%=enc.EncodeForJavascript(p_nalbaran)%>');"/></td><%
	   CloseFila%>
	   </table></p>
  	   </span>
    <%end if 'mode="pedirserie"

    'mmg >> si se ha modificado el proveedor buscamos la serie del mismo
    if p_nproveedor&""<>"" and p_serie&""<>"" then 'se modifica el proveedor con la lupa
        TraerProveedor= right(p_nproveedor,len(p_nproveedor)-5)
        if cint(null_z(cambiar_serie))=1 then 'or cambiar_serie & ""="" then
            rstAux.cursorlocation=3
            rstAux.open "select serie_alb from documentos_pro with(nolock) where nproveedor='" & p_nproveedor & "'", session("dsn_cliente")
		    if not rstAux.eof then
			    if rstAux("serie_alb")&"">"" then
				    p_serie=rstAux("serie_alb")
			    end if
		    end if
		    rstAux.close
		end if
    else
        if provR<>"" and p_serie="" and limpiaCadena(request.QueryString("nproveedor"))<>"" then 'se modifica el proveedor sin la lupa
            TraerProveedor= Completar(provR,5,"0")
            if cint(null_z(cambiar_serie))=1 then 'or cambiar_serie & ""="" then
                rstAux.cursorlocation=3
                rstAux.open "select serie_alb from documentos_pro with(nolock) where nproveedor='" & session("ncliente")&TraerProveedor & "'", session("dsn_cliente")
		        if not rstAux.eof then
			        if rstAux("serie_alb")&"">"" then
				        p_serie=rstAux("serie_alb")
			        end if
			    end if
		        rstAux.close
		    end if
        else
            if TraerProveedor="" and mode="add" then
		        'Obtener el proveedor de la serie por defecto.
		        if cint(null_z(cambiar_cliente))=1 or cambiar_cliente & ""="" then
		            TraerProveedor=d_lookup("substring(cliente,6,10)","series","nserie='" & p_serie & "'",session("dsn_cliente"))
		            if TraerProveedor&""="" then
		                TraerProveedor=limpiaCadena(request.QueryString("prov"))
		            end if
		        else
		            TraerProveedor=limpiaCadena(request.QueryString("prov"))
		        end if
	        end if
	    end if
    end if

	if (mode="add" or mode="edit") and TraerProveedor<>"" then
        rstAux.cursorlocation=3
		rstAux.open "select fbaja from proveedores with(nolock) where nproveedor='" & session("ncliente") & Completar(TraerProveedor,5,"0") & "'", session("dsn_cliente")
		if not rstAux.eof then
			if rstAux("fbaja")>"" then%>
				<script language="javascript" type="text/javascript">				    window.alert("<%=LitProveedorDadoBaja%>");</script>
				<%TraerProveedor=""
			end if
		end if
		rstAux.close
	end if

	'Captura de datos del proveedor que se está introduciendo en el pedido'
	if TraerProveedor > "" then
		TraerProveedor=session("ncliente") & Completar(TraerProveedor,5,"0")
		Error="NO"
		strselect="select * from proveedores with(nolock) where nproveedor='" & TraerProveedor & "'"
        rstAux.cursorlocation=3
		rstAux.open strselect,session("dsn_cliente")
		if not rstAux.EOF then
	  	    tmp_nproveedor=TraerProveedor
			tmp_nombre=rstAux("razon_social")
			tmp_forma_pago=rstAux("forma_pago")
			tmp_tipo_pago=rstAux("tipo_pago")
			tmp_divisa=rstAux("divisa")
			tmp_dto1=rstAux("descuento")
			tmp_dto2=null_z(rstAux("descuento2"))
			tmp_rf=rstAux("recargo")
			tmp_irpf=rstAux("irpf")
			tmp_IRPF_Total=rstAux("IRPF_Total")
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
			if cint(null_z(cambiar_serie))=1 then 'or cambiar_serie & ""="" then
				'obtener_doc_cli mode,"albaranes_pro",tmp_nproveedor,p_serie,p_valorado,tmp_valorado,tmp_irpf
				set rstObtDocCli= Server.CreateObject("ADODB.Recordset")
                dato_doc1="serie_alb"
                dato_doc2="valorado_alb"
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
			                p_serie=rstObtDocCli(dato_doc1)
		                end if
		                if dato_doc2 & "">"" then
			                p_valorado=rstObtDocCli(dato_doc2)
			                tmp_valorado=rstObtDocCli(dato_doc2)
		                end if
		                'FLM:010409: comento esto xq no se debe obtener el irpf de la tabla empresas, si no del proveedor.
		                'if rstObtDocCli("irpf") & "">"" then
			            '    tmp_irpf=rstObtDocCli("irpf")
		                'end if
	                end if
	                rstObtDocCli.close
                end if
                set rstObtDocCli=nothing
			end if
		else
			Error="SI"%>
			<script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgProveedorNoExiste%>");
			</script>
		<%end if
		rstAux.close
	end if
  'Acción a realizar'

	if mode="save" or mode="first_save" then
        rst.cursorlocation=2
		rst.Open "select * from albaranes_pro where nalbaran='" & p_nalbaran & "'", _
		session("dsn_cliente"),adOpenKeyset,adLockOptimistic

		if p_nproveedor > "" then
            rstProveedor.cursorlocation=3
	      	rstProveedor.open "select nproveedor from proveedores with(nolock) where nproveedor='" & p_nproveedor & "'",session("dsn_cliente")
			if rstProveedor.EOF then%>
		      	<script language="javascript" type="text/javascript">
                      window.alert("<%=LitMsgProveedorNoExiste%>");
                      history.back();
				</script> 
				<%mode="add"
			else
				ModDocumento=true
				ModDocumentoEquip=false
				''ricardo 12/11/2003 comprobamos que no exista el nalbaran_pro para un mismo proveedor
				no_continuar=0 
				strselect="select count(nalbaran) as contador from albaranes_pro with(nolock) where nproveedor='" & p_nproveedor & "' and nalbaran_pro='" & p_nalbaran_pro & "' and nalbaran like '" & session("ncliente") & "%'  and year(fecha)= year(convert (datetime,'" & p_fechaR & "' )) "
				if mode="save" then
					strselect=strselect & " and nalbaran<>'" & p_nalbaran & "'"
				end if
                rstAux.cursorlocation=3
				rstAux.open strselect,session("dsn_cliente")
				if not rstAux.eof then
					cuantos_albaranes=null_z(rstAux("contador"))
					rstAux.close
					if cuantos_albaranes>0 then
						ModDocumento=false%>
						<script language="javascript" type="text/javascript">
                            window.alert("<%=LitMsgNumeroAlbaranRepetido%>");
						</script>
						<%if mode="first_save" then%>
							<script language="javascript" type="text/javascript">
                                document.location = "albaranes_pro_pista.asp?mode=add&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
                                parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=add"
							</script>
						<%else
							submode2=mode
							if mode="save" then
								mode="edit"
							elseif mode="" then
								mode="add"
							end if%>
							<script language="javascript" type="text/javascript">
                                document.location = "albaranes_pro_pista.asp?mode=<%=enc.EncodeForJavascript(mode)%>&nalbaran=<%=enc.EncodeForJavascript(p_nalbaran)%>&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>"
                                parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=<%=enc.EncodeForHtmlAttribute(null_s(mode))%>";
							</script>
							<%''ricardo 10-12-2003 se cambia el modo ya que si no da error
							mode=""
						end if
					end if
				else
					rstAux.close
				end if
				if ModDocumento=true then
					if mode="first_save" then
						ModDocumento=true
						ModDocumentoEquip=false
					else
						ModDocumento=false
						ModDocumentoEquip=false
						mensajeTratEquipos="OK0"

						if p_nalbaran & "">"" and not rst.eof then
							'solamente se comprobara si cambia el cliente,centro o fecha
							if ((rst("nproveedor")&""<>p_nproveedor&"") or (rst("fecha")<>cdate(p_fechaR&""))) then
								ModDocumentoEquip=true
								mensajeTratEquipos=TratarEquipos("","","ALBARAN DE PROVEEDOR",p_nalbaran,"","","","","",mode)
							end if
						end if
						if mid(mensajeTratEquipos,1,2)<>"OK" then
							ModDocumento=false%>
							<script language="javascript" type="text/javascript">
                                window.alert("<%=mensajeTratEquipos%>");
							</script>
						<%else
							ModDocumento=true
						end if
					end if
				end if
				if ModDocumento then
				 ''FLM:200309 Comprobamos la cuenta de abono del proveedor para esta factura.
                   if ComprobarCuenta(Nulear(request.form("ncuenta_pro")))=false then
                        ModDocumento=false%>
                        <script language="javascript" type="text/javascript">
                                alert("<%=LitCuentaAbonoError%>");

                        <%if mode= "first_save" then%>
                                    history.back();
                                history.back();
                                parent.botones.document.location = "albaranes_pro_bt.asp?mode=add"
                                    <%else 
                                submode2 = mode
                                if mode= "save" then
                                mode = "edit"
                                elseif mode= "" then
                                mode = "add"
                                end if%>

                                    document.location="albaranes_pro_pista.asp?mode=<%=enc.EncodeForJavascript(mode)%>&nalbaran=<%=enc.EncodeForJavascript(p_nalbaran)%>&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>"
                                parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=<%=enc.EncodeForJavascript(mode)%>";
                        <%end if%>
                        </script>
                        <%mode=""
                    end if		
				end if
				
				if ModDocumento then
					'comprobamos si el nalbaran existe o no segun el contador de configuracion
					if mode="first_save" then
						p_serieAux=p_serie
						if compNumDocNuevo(p_serieAux,p_fechaR,"albaranes_pro")=0 then%>
							<script language="javascript" type="text/javascript">
                                    window.alert("<%=LitMsgDocExistRevCont%>");
                                document.location = "albaranes_pro_pista.asp?mode=add&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
                                parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=add"
							</script>
							<%ModDocumento=false
						end if			
					end if
				end if
				if ModDocumento then
					GuardarRegistro p_nalbaran,p_serie
					if not rst.eof then
			      		p_nalbaran=rst("nalbaran")
						if ModDocumentoEquip=true then
							InsertarHistorialNserie mensajeTratEquipos,"","","ALBARAN DE PROVEEDOR",p_nalbaran,"","","","","MODIFY",mode
						end if
				      	if mode="first_save" then
			      			auditar_ins_bor session("usuario"),p_nalbaran,rst("nproveedor"),"alta","","","albaranes_pro"
						end if
					else
				  		p_nalbaran=""
					end if
				else
					no_modificado=1
				end if
				ant_mode=mode
				mode="browse"
			end if
			rstProveedor.close
		end if
		rst.close
	elseif mode="delete" then
		he_borrado=1
        rst.cursorlocation=3
		rst.Open "select nfactura from albaranes_pro with(nolock) where nalbaran='" & p_nalbaran & "'",session("dsn_cliente")
		if not rst.eof then
			if isnull(rst("nfactura")) then
				rst.close
				'Comprobar si se puede eliminar el albarán.
				mensajeTratEquipos=TratarEquipos("","","ALBARAN DE PROVEEDOR",p_nalbaran,"","","","","",mode)
				if mid(mensajeTratEquipos,1,2)<>"OK" then
					mode="browse"%>
					<script language="javascript" type="text/javascript">
                                window.alert("<%=mensajeTratEquipos%>");
					</script>
				<%else
                    rst.cursorlocation=3
					rst.open "select ndocumento from caja with(nolock) where tdocumento='ALBARAN DE PROVEEDOR' and ndocumento='" & p_nalbaran & "' and caja like '" & session("ncliente") & "%' ",session("dsn_cliente")
					if rst.EOF then
						rst.close
                        rst.cursorlocation=3
						rst.open "select ndocumento from detalles_dev_pro with(nolock) where ndocumento='" & p_nalbaran & "' and ndevolucion like '" & session("ncliente") & "%' ",session("dsn_cliente")
						if rst.EOF then
							rst.close
                            rst.cursorlocation=3
							rst.open "select nproveedor from albaranes_pro with(nolock) where nalbaran='" & p_nalbaran & "'",session("dsn_cliente")
                            if not rst.eof then
							    nprov_aux=rst("nproveedor")
                            else
                                nprov_aux=""
                            end if
							rst.close
							auditar_ins_bor session("usuario"),p_nalbaran,nprov_aux,"baja","","","albaranes_pro"
							InsertarHistorialNserie mensajeTratEquipos,"","","ALBARAN DE PROVEEDOR",p_nalbaran,"","","","","",mode
							BorrarRegistro p_nalbaran, p_nalbaran_pro
                             %><script language="javascript" type="text/javascript">
                                   parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=add";
                                   SearchPage("deliveryNote_pro_lsearch_pista.asp?mode=init", 0);
		                        </script><%
						else
							rst.close
							mode="browse"%>
							<script language="javascript" type="text/javascript">
                                   window.alert("<%=LitMsgNoBorrarAlbDev%>");
							</script>
						<%end if
                        ' >>> MCA 21/04/05 Para cargar el modo add tras el borrado
						mode="add"
						p_nalbaran=""
					else
						rst.close
						mode="browse"%>
						<script language="javascript" type="text/javascript">
                                   window.alert("<%=LitMsgNoBorrarAlbAnotCaja%>");
						</script>
					<%end if
				end if
			else
				mode="browse"
				fac=rst("nfactura")
				rst.close%>
				<script language="javascript" type="text/javascript">
                                window.alert("<%=LitMsgBorrarAlbaran%><%=trimCodEmpresa(fac)%>");
				</script>
			<%end if
		else
			rst.close%>
			<script language="javascript" type="text/javascript">
                            window.alert("<%=LitMsgDocsNoExiste%>");
                            parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=add";
			</script>
		    <%mode="add"
		end if
    end if

    'Mostrar los datos de la página.
    ''ricardo 31/7/2003 comprobamos que existe el albaran
    if mode="browse" and he_borrado<>1 and no_modificado<>1 then
        rstAux.cursorlocation=3
		rstAux.open "select nalbaran from albaranes_pro with(nolock) where nalbaran='" & p_nalbaran & "'", session("dsn_cliente")
		if rstAux.eof then
			p_nalbaran=""%>
			<script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgDocsNoExiste%>");
                    parent.botones.document.location = "albaranes_pro_bt_pista.asp?mode=add";
			</script>
			<%mode="add"
		end if
		rstAux.close
    end if

    if mode="browse" or mode="edit" then
		if p_nalbaran="" then
            rstAux.cursorlocation=3
			rstAux.open "select top 1 nalbaran from albaranes_pro with(nolock) where nalbaran like '" & session("ncliente") & "%' order by fecha desc,nalbaran desc", session("dsn_cliente")
			if not rstAux.eof then p_nalbaran=rstAux("nalbaran")
			rstAux.close
		end if
		' JMA 20/12/04 Campos personalizables'
		rstAux.cursorlocation=3
		rstAux.open "select p.campo01,p.campo02,p.campo03,p.campo04,p.campo05,p.campo06,p.campo07,p.campo08,p.campo09,p.campo10 from albaranes_pro as p with(nolock) where p.nalbaran='" & p_nalbaran & "'",session("dsn_cliente")
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
        rst.cursorlocation=3
		rst.Open "select p.*,pv.razon_social,ser.nombre as nomserie from albaranes_pro as p with(nolock),proveedores as pv with(nolock),series as ser with(nolock) where ser.nserie=p.serie and pv.nproveedor=p.nproveedor " & _
		"and nalbaran='" & p_nalbaran & "' and pv.nproveedor like '" & session("ncliente") & "%' and ser.nserie like '" & session("ncliente") & "%' ", session("dsn_cliente")
	elseif mode="add" then
		rst.Open "select * from albaranes_pro with(nolock) where nalbaran='" & p_nalbaran & "'", _
		session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst.AddNew
		rst("valorado")=1
	
	end if
	sumadet=0
	sumaRE=0

	'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION

    VinculosPagina(MostrarProveedores)=1
	CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
    %><div class="headers-wrapper"><%
            DrawDiv "header-date","",""
                DrawLabel "txtMandatory","",LitFecha
                    if mode="edit" then
				        %><input type="hidden" name="h_fecha" value="<%=rst("fecha")%>"/><%
			        end if

				    if mode="browse" then 
                        DrawSpan "","",rst("fecha"),""
                    else
                        DrawInput "'width50 fecha'","","fecha",iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha")),""
                        DrawCalendar "fecha"
                    end if
            CloseDiv' fecha%><script type="text/javascript">
            $(".fecha").keypress(function (e) {
                if (isTextSelected(this)) {
                    return;
                }
                if(e.which !== 8) {
                    if (this.value.length > 9)
                        return false;
                    var numChars = $(this).val().length;
                    if(numChars === 2 || numChars === 5) {
                        var thisVal = $(this).val();
                        thisVal += '/';
                        $(this).val(thisVal);
                    }
                }
            });</script><%DrawDiv "header-nalbaran","",""
                DrawLabel "txtMandatory","",LitAlbaran
                if mode="add" or mode="edit" then
                    if mode="add" then
                        DrawInput "width150px","","nalbaran_pro",iif(p_nalbaran_pro>"",p_nalbaran_pro,""),""
                    else
                        if not rst.eof then
                            DrawInput "width150px","","nalbaran_pro",iif(p_nalbaran_pro>"",p_nalbaran_pro,rst("nalbaran_pro")),""
                        else
                            DrawInput "width150px","","nalbaran_pro",iif(p_nalbaran_pro>"",p_nalbaran_pro,""),""
                        end if
                    end if
                else
                    DrawSpan "","",rst("nalbaran_pro"),""
                end if
            CloseDiv
            DrawDiv "header-nproveedor","",""
                DrawLabel "txtMandatory","",LitProveedor
                Formulario="albaranes_pro"
			    if mode="browse" then
				    if rst("nproveedor")>"" then%>
					    <%=Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("nproveedor")),LitVerProveedor)%>
				    <%end if
			    else
				    %><input class="CELDA width20" type="text" name="nproveedor" value="<%=enc.EncodeForHtmlAttribute(TrimCodEmpresa(iif(tmp_nproveedor>"",tmp_nproveedor,rst("nproveedor"))))%>" size="8" onchange="TraerSerie('<%=enc.EncodeForJavascript(mode)%>');"/>
				    <a class="CELDAREFB"  href="javascript:AbrirVentana('proveedores_busqueda.asp?ndoc=albaranes_pro&titulo=<%=LitSelProv%>&mode=search','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" onmouseover="self.status='<%=LitVerProveedor%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion)%><%=enc.EncodeForHtmlAttribute(ImgBuscar)%>" <%=ParamImgBuscar%> alt="<%=LitBuscarProveedor%>" title="<%=LitBuscarProveedor%>"/></a><%
			    end if
			    nompro = d_lookup("razon_social","proveedores","nproveedor='" & iif(tmp_nproveedor>"",tmp_nproveedor,rst("nproveedor")) & "'",session("dsn_cliente"))
                nompro=replace(nompro,"'","")

                if mode="edit" or mode="add" then
                    %><input class="CELDA width30" type="text" disabled name="razon_social" value="<%=nompro%>" /><%
                elseif mode="browse" then
                    DrawSpan "","", "&nbsp;&nbsp;" &nompro ,""
                end if

			    'Aqui iba la capa de navegacion
			    if tmp_nproveedor>"" then
				    proveedor_aux=tmp_nproveedor
			    else
				    proveedor_aux=rst("nproveedor")
			    end if
            CloseDiv
  

	DrawDiv "header-note","",""
	    if not rst.eof then
		    if mode="browse" and (rst("nfactura")&"")="" then
		        MB=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
		        n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") &"'",session("dsn_cliente"))
		        EnCaja=CambioDivisa(d_sum("importe","pagos_alb_pro","nalbaran='" & rst("nalbaran") & "'",session("dsn_cliente")),rst("divisa"),rst("divisa"))
		        Pendiente=miround(rst("total_albaran")-EnCaja,n_decimales)
		        defecto=""
		        poner_cajasResponsive1 "input-ncaja",defecto,"ncaja","100","codigo","descripcion","","",poner_comillas(caju)
  	            %><span class='header-note-inputCaja'>
			        <input class='CELDAL7' type="Text" name="impcaja" value="<%=enc.EncodeForHtmlAttribute(Pendiente)%>"/>
		        </span>
		        <span class='header-note-currency'>
			        <font class="ENCABEZADOR7"><%=d_lookup("abreviatura","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))%></font>
		        </span>
		        <span class='header-note-buttonNote'>
			        <img src="<%=enc.EncodeForHtmlAttribute(themeIlion)%><%=enc.EncodeForHtmlAttribute(ImgAnotar)%>" style="cursor:pointer;" <%=ParamImgAnotar%> alt="<%=LITANOTARCAJA%>" title="<%=LITANOTARCAJA%>" onclick="Acaja('<%=enc.EncodeForJavascript(rst("nalbaran"))%>')"/>
  		        </span><%
                rstAux.cursorlocation=3
	  	        rstAux.Open "SELECT codigo,descripcion FROM Tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
		        'DrawSelectCelda "CELDA7","100","","0", "","i_pago",rstAux,session("ncliente") & "01","codigo","Descripcion","",""
		        DrawSelect "input-i_pago","width:150px;","i_pago",rstAux,session("ncliente") & "01","codigo","Descripcion","",""
                rstAux.Close
	        end if
	    end if
    CloseDiv            
    'otro ?
	if mode="browse" or mode="save" then
	    ''ricardo 13-3-20003
        ''si la serie tiene un formato de impresion sera este el de por defecto
        ''si no sera el elegido en la tabla formatos impresion de ilion
        if not rst.eof then
            defecto=obtener_formato_imp(rst("serie"),"ALBARAN DE PROVEEDOR")
        end if
        ''''''''
	    seleccion = "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ALBARAN DE PROVEEDOR' order by descripcion"
        rstSelect.cursorlocation=3
	    rstSelect.Open seleccion, DsnIlion
	    if si_tiene_modulo_21 = 0 and si_tiene_modulo_22 = 0 then
            DrawDiv "header-resources alignCenter","",""
            %>
                <a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=Servicios/recursos.asp&pag2=Servicios/recursos_bt.asp&codigo=<%=enc.EncodeForJavascript(null_s(nalbaran))%>&codigo_print=<%=enc.EncodeForJavascript(null_s(p_nalbaran_pro))%>&tipo=albaran de proveedor&viene=enlaces', 'P', <%=AltoVentana%>, <%=AnchoVentana%>)">&nbsp;&nbsp;&nbsp;<%=LitEnlaces%>&nbsp;&nbsp;&nbsp;</a>
            <%
            CloseDiv
        end if
        DrawDiv "header-print","",""
	    %><label><a id="idPrintFormat" class="CELDAREFB" href="javascript:AbrirVentana(document.albaranes_pro_pista.formato_impresion.value+'nalbaran=<%="(\'"+enc.EncodeForJavascript(null_s(p_nalbaran))+"\')"%>&mode=browse&empresa=<%=session("ncliente")%>&novei=<%=enc.EncodeForJavascript(null_s(novei))%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" onmouseover="self.status='<%=LitImpresionConFormato%>'; return true;" onmouseout="self.status=''; return true;"><%=LitImpresionConFormato%></a></label><%
	    'DrawSelectCelda "CELDARIGHT","120","",0,"","formato_impresion",rstSelect,defecto,"fichero","descripcion","",""
	    %><select class='CELDA' style='width:150px' name="formato_impresion"><%
	        encontrado=0
		    while not rstSelect.eof
			    if defecto=rstSelect("descripcion") then
				    encontrado=1
				    if isnull(rstSelect("parametros")) then
					    prm=""
				    else
					    prm=rstSelect("parametros") & "&"
				    end if
				    %><option selected="selected" value="<%=rstSelect("fichero") & "?" & prm%>"><%=rstSelect("descripcion")%></option><%
			    else
				    if isnull(rstSelect("parametros")) then
					    prm=""
				    else
					    prm=rstSelect("parametros") & "&"
				    end if
				    %><option value="<%=rstSelect("fichero")  & "?" & prm%>"><%=rstSelect("descripcion")%></option><%
			    end if
			    rstSelect.movenext
		    wend
	    %></select><%
	    rstSelect.close
   	    CloseDiv
        if session("version")&"" = "5" then
            DrawDiv "","","" 
            CloseDiv
        end if
        
	end if%></div><%
    ' cierre wrapper

    if mode="browse" then
        BarraOpciones "browse", rst("nalbaran")
    end if

    ActionVersion Altoventana, AnchoVentana

    %><table class="width100"></table> <%

	'Alarma "albaranes_pro.asp"
	
	if (mode="browse" or mode="edit" or mode="add") and not rst.EOF then%>
		<input type="hidden" name="h_nproveedor" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nproveedor")))%>"/>
		<input type="hidden" name="h_nalbaran" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran")))%>"/>
		<input type="hidden" name="olddivisa" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("divisa")))%>"/>
		<input type="hidden" name="h_nalbaran_pro" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran_pro")))%>"/>
		<input type="hidden" name="h_nbalance" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nbalance")))&""%>"/><%			
			
            
            '*** i AMP 04102010 : Incorporamos campo factor de cambio.
            'Información sobre la moneda base de la empresa.
            monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
   	        abrevBase =  d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))
          	factcambio = d_lookup("factcambio","albaranes_pro","nalbaran='" & rst("nalbaran") & "' and nalbaran like '" & session("ncliente") & "%'",session("dsn_cliente"))
       		    	
	       '*** f AMP 04102010%>

        <!--<div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['GENERAL_DATA','FINANCIAL_DATA', 'CABECERA', 'DIRENVIO', 'TOTAL']); hideNoCollapse();"><img Class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['GENERAL_DATA','FINANCIAL_DATA', 'CABECERA', 'DIRENVIO', 'TOTAL']);hideCollapse();" style="display:none"><img Class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
        </div>-->

        <div class="Section" id="S_GENERAL_DATA" >
            <a href="#" rel="toggle[GENERAL_DATA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader" >
                    <%=LITCABECERA %>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
            <div class="SectionPanel" id="GENERAL_DATA">
                <table width="100%" bgcolor="<%=color_blau %>" border="0">
                <%if mode="browse" then
                    DrawDiv "1","",""
                    DrawLabel "txtMandatory","",LitSerie
                    DrawSpan "CELDA","",trimCodEmpresa(rst("serie")) & " - " & rst("nomserie"),""
                    CloseDiv
					
					  'calculamos el almacen de la serie
					  strconn=session("dsn_cliente")
                    set rstMM = Server.CreateObject("ADODB.Recordset")
                  
                    if rstMM.State<>0 then rsAux.Close
                    rstMM.cursorlocation=3
		            rstMM.open "select almacen from series with(nolock), almacenes alm with(nolock) where nserie='"&rst("serie")&"' and alm.codigo=almacen and isnull(alm.fbaja,'')=''"& strwhere,session("dsn_cliente")
		    		if rstMM.eof then
		    		    almacenSerie= ""
		    		else
		    		    almacenSerie= rstMM("almacen")
		    		end if
				  else
                      ' >>> MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compras.
                      '					 bloque 4/4 en albaranes_pro.asp
					    strSacSerie="select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & session("ncliente") & "%' and tipo_documento ='ALBARAN DE PROVEEDOR'"
					    if s & "">"" then
					        strSacSerie=strSacSerie & " and nserie in " & s
					    end if
					    strSacSerie=strSacSerie & " order by nserie"
                        rstAux.cursorlocation=3
					    rstAux.open strSacSerie,session("dsn_cliente")

                      ' <<< MCA 11/04/05 : Añadir parámetro de usuario con la(s) serie(s) a los documentos de compras.

			    	    if mode="add" then
					        DrawDiv "1","",""
                            DrawLabel "txtMandatory","",LitSerie
                            DrawSelect "width60","","serie",rstAux,iif(p_serie>"",p_serie,rst("serie")),"nserie","descripcion","onchange","javascript:TraerProveedor('add');"
                            CloseDiv
					    else
                            DrawDiv "1","",""
                            DrawLabel "txtMandatory","",LItSerie
                            DrawSelect "width60","","serie",rstAux,iif(p_serie>"",p_serie,rst("serie")),"nserie","descripcion","",""
                            CloseDiv
					    end if
			 		    rstAux.close
				  end if
                  if mode = "browse" then
                      %><div class="hidden"><%
                          DrawLabel "", "", LitValorado
                          EligeCeldaResponsive1 "check", mode, "CELDA", "", "valorado", iif(p_valorado>"",nz_b(p_valorado),rst("valorado")), LitValorado
                      CloseDiv
                  else
                      DrawDiv "hidden","",""
                          EligeCeldaResponsive "check", mode, "hidden", "", "", 0, LitValorado, "valorado", "",iif(p_valorado>"",nz_b(p_valorado),rst("valorado"))
                      CloseDiv
                  end if

                  campo="codigo"
				  if mode="browse" then
				        campo2="abreviatura"
				  else
					    campo2="abreviatura"
				  end if
				  '*** AMP 20102010			
				  if tmpdivisafc>"" then  tmp_divisa = tmpdivisafc end if
                  DIVISA=iif(tmp_divisa>"",tmp_divisa,rst("divisa"))
				  dato_celda=Desplegable(mode,campo,campo2,"divisas",DIVISA,"moneda_base<>0 and codigo like '" & session("ncliente") & "%'")
''response.write("los datos 1 son-" & mode & "-" & tmpdivisafc & "-" & DIVISA & "-" & tmp_divisa & "-" & rst("divisa") & "-<br>")
				
				 if mode<>"browse" then
				    datoDivisa=iif(tmp_divisa>"", _
				        d_lookup("abreviatura","divisas","codigo='" & tmp_divisa & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")), _
				        d_lookup("abreviatura","divisas","codigo='" & dato_celda & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")))
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
	                 cuantos_detalles=d_count("item","detalles_alb_pro","nalbaran='" & rst("nalbaran") & "'",session("dsn_cliente"))
	                 cuantos_conceptos=d_count("nconcepto","conceptos_alb_pro","nalbaran='" & rst("nalbaran") & "'",session("dsn_cliente"))
	                 if cint(cuantos_detalles) + cint(cuantos_conceptos)>0 then
		                 estilo_divisa="CELDA DISABLED"
		                 tipo_eligecelda="input"
	                 else
		                 estilo_divisa="CELDA"
		                 tipo_eligecelda="select"
	                 end if
                 end if
                
                 if mode="add" or mode="edit" then RstAux.close
				 Estilo=iif(mode="browse","CELDA",estilo_divisa)
				 if tipo_eligecelda="input" then
				     if mode = "edit" then
                         DrawDiv "1","",""
                         DrawLabel "txtMandatory","",LitDivisa%>
                         <input disabled size="5" type='text' name="divisa" value='<%=enc.EncodeForHtmlAttribute(null_s(datoDivisa)) %>' /><%
                         CloseDiv
                     elseif mode = "browse" then
                         DrawDiv "1","",""
                         DrawLabel "txtMandatory","",LitDivisa
                         DrawSpan "CELDA","",iif(mode="add",datoDivisa,d_lookup("abreviatura","divisas","codigo='" & rst("divisa") & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))),""
                         CloseDiv
                     end if
                 else
				     monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))	
                     rstAux.cursorlocation=3
				     rstAux.open "select codigo,abreviatura as descripcion from divisas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
                     DrawDiv "1","",""
                     DrawLabel "txtMandatory","",LitDivisa
                     DrawSelect "width20","","divisa",rstAux,iif(mode<>"browse",iif(tmp_divisa>"",tmp_divisa,dato_celda),rst("divisa")),"codigo","descripcion","onchange","javascript:cambiardivisa('"&monedaBase&"');"
                     CloseDiv
			 	     rstAux.close
				 end if
''response.write("los datos 3 son-" & mode & "-" & datoDivisa & "-" & tmp_divisa & "-" & rst("divisa") & "-<br>")
                    'DrawFila color_blau
                        if mode="browse" then
	                        if  rst("divisa")<>monedaBase then
                                abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&rst("divisa")&"'",session("dsn_cliente"))
	                            'DrawCelda "ENCABEZADOL style='width:130px'","","",0,LitFactCambio+" :"
	                            ''DrawCelda "CELDA style='width:200px'","","",0,CStr(factcambio)+" "+abrevBase
                                        ''response.write(CStr(factcambio) & " " & abrevBase)
                                        'response.write("1" & abrevBase & " = " & CStr(factcambio) & abreviaAtDiv)
                                    %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                        %><label><%=LitFactCambio %></label><%
                                        %><span class="CELDA"><%="1" & abrevBase & " = " & CStr(factcambio) & abreviaAtDiv %></span><%
                                      %></div><%

                                    %><input type="hidden" name="h_divisa" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_divisa>"",tmp_divisa,dato_celda)))%>"/>
		                            <!--<input type="hidden" name="divisafc" value="<%=iif(tmp_divisa>"",tmp_divisa,DIVISA)%>"/>-->
                                    <input type="hidden" name="divisafc" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(mode="add",iif(tmp_divisa>"",tmp_divisa,DIVISA),rst("divisa"))))%>"/><%
                                
	                        end if
	                      else
	                          ocultar=0  
	                          if mode="add" or mode="edit" then	  
              	                if mode="add" then
              	                    dv=iif(DIVISA>"",DIVISA,dato_celda)              	         
              	                     factcambio = d_lookup("factcambio","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dv&"'",session("dsn_cliente"))               	                 	         
                                    abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dv&"'",session("dsn_cliente"))
              	                     if dv=monedaBase then ocultar=1 end if
              	                 else 'modo edit
                                        dvEdit=iif(tmp_divisa>"",tmp_divisa,dato_celda)
                                        abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dvEdit&"'",session("dsn_cliente"))
              	                     if dvEdit=monedaBase then ocultar=1 end if              	        
              	                 end if


                                DrawDiv "1", "", "tdfactcambio"
                                DrawLabel "", "", LitFactCambio
                                DrawSpan "CELDA", "", "1" & abrevBase & " = ", ""
                                %>
                                    <input type="text" name="nfactcambio" value="<%=enc.EncodeForHtmlAttribute(null_s(CStr(factcambio)))%>" size="6" style="text-align:right" onchange="comprobarFactorCambio()"/>
                                    <span class="CELDA" id="idfactcambioexpl"><%=enc.EncodeForHtmlAttribute(null_s(abreviaAtDiv))%></span>
                                    <input type="hidden" name="h_divisa" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_divisa>"",tmp_divisa,dato_celda)))%>"/>
		                            <input type="hidden" name="divisafc" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_divisa>"",tmp_divisa,DIVISA)))%>"/>
                                <%
                                CloseDiv
	          
	                            if ocultar=1 then %>
                               <script language="javascript" type="text/javascript">
                                   parent.pantalla.document.getElementById("tdfactcambio").style.display = "none"
                                </script><% 
                                end if
	                          end if	     
	                      end if%>

		                <%n_decimales=d_lookup("ndecimales","divisas","codigo='" & iif(mode="browse",DIVISA,iif(tmp_divisa>"",tmp_divisa,dato_celda)) & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))%>
                       <!--</td>-->
                    <%'CloseFila
                    'DrawFila color_titulo
                    DrawDiv "3-sub", "background-color: #eae7e3", ""
		                DrawLabel "", "", LIT_GENERAL_DATA
                    CloseDiv
			        'CloseFila
                    'DrawFila color_blau
				        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:150px;'","","",0,LitFormaPago+":"
				        campo="codigo"
				        campo2="descripcion"
				        dato_celda=Desplegable(mode,campo,campo2,"formas_pago",iif(tmp_forma_pago>"",tmp_forma_pago,rst("forma_pago")),"")
				        
                        if mode = "browse" then
                        DrawDiv "hidden","",""
                            EligeCeldaResponsive "text", mode,"hidden","","",0,LitFormaPago, LitFormaPago, 15 ,dato_celda
                        CloseDiv
                        else
                        DrawDiv "hidden","",""
                            EligeCelda "select", mode, "hidden", iif(mode<>"browse","200",""), "", 0, LitFormaPago, "forma_pago", 15, dato_celda
                        CloseDiv
                        end if

				        
				        if mode="add" or mode="edit" then RstAux.close
				        ''if mode<>"browse" then 
                            'DrawCelda "CELDA","","",0,"&nbsp;&nbsp;"
                        ''end if
				        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:150px;'","","",0,LitTipoPago+":"
				        dato_celda=Desplegable(mode,campo,campo2,"tipo_pago",iif(tmp_tipo_pago>"",tmp_tipo_pago,rst("tipo_pago")),"")
				        
                        if mode = "browse" then
                        DrawDiv "hidden","",""
                            EligeCeldaResponsive "text", mode, "hidden", "", "", 0, LitTipoPago, LitTipoPago, 15, dato_celda
                        CloseDiv
                        else
                        DrawDiv "hidden","",""
                            EligeCeldaResponsive "select",mode,"hidden","","",0,LitTipoPago,"tipo_pago","",dato_celda
                        CloseDiv
                        end if
                        

				        if mode="add" or mode="edit" then RstAux.close
				        if mode="browse" then
                        DrawDiv "hidden","",""
                            EligeCeldaResponsive "text",mode,"hidden","","",0,LitNumFactura,LitNumFactura,"",d_lookup("nfactura_pro","facturas_pro","nfactura='" & iif(isnull(rst("nfactura")),"",rst("nfactura")) & "'",session("dsn_cliente"))
				        CloseDiv
                        else
                            %><div class="hidden"><%
                                %><label><%=LitNumFactura %></label><%
						        %><input class='width60' type="Text" name="nfactura" value="<%=d_lookup("nfactura_pro","facturas_pro","nfactura='" & iif(isnull(rst("nfactura")),"",rst("nfactura")) & "'",session("dsn_cliente"))%>" disabled="disabled"/><%
                            %></div><%
					        %><!--<td CLASS=CELDA>
						        <input class='CELDA' type="Text" name="nfactura" size="25" value="<%=d_lookup("nfactura_pro","facturas_pro","nfactura='" & iif(isnull(rst("nfactura")),"",rst("nfactura")) & "'",session("dsn_cliente"))%>" disabled="disabled"/>
					        </td>--><%
				        end if
				        ''if mode<>"browse" then 
                            'DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''end if
				        if si_tiene_modulo_proyectos<>0 then
					        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:150px;'","","",0,LitProyecto+":"
					        if mode <> "browse" then
                                %><div class="hidden"><%
                                    %><label><%=LitProyecto%></label><%
						            %><input class="CELDA" type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto")))))%>"/><%
						            %><iframe id='frProyecto' name='fr_Proyecto' src='../mantenimiento/docproyectos.asp?viene=albaranes_pro&mode=<%=enc.EncodeForJavascript(mode)%>&cod_proyecto=<%=enc.EncodeForJavascript(null_s(iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto")))))%>' width='250' height='30' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
					            %></div><%
                            else
                                    DrawDiv "hidden","",""
						        'DrawCelda "CELDA","","",0,d_lookup("nombre","proyectos","codigo='" & rst("cod_proyecto") & "'",session("dsn_cliente"))
                                EligeCeldaResponsive "text",enc.EncodeForHtmlAttribute(null_s(mode)), "hidden", "", "", 0, LitProyecto, LitProyecto, "",d_lookup("nombre","proyectos","codigo='" & rst("cod_proyecto") & "'",session("dsn_cliente"))
                                    CloseDiv
					        end if
				        end if
			        'CloseFila
			        if si_tiene_modulo_ccostes<>0 then '**rgu 1/9/2009
			        'DrawFila color_blau
				        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:150px;'","","",0,LitTienda+":"
				        if mode <> "browse" then
				            defecto=iif(tmp_tienda>"",tmp_tienda,rst("tienda"))
                            rstAux.cursorlocation=3
					        rstAux.open "select codigo, descripcion from tiendas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
                                    DrawDiv "hidden","",""
	    			        DrawSelectCelda "hidden","","",0,LitTienda,"tienda",rstAux,defecto,"codigo","descripcion","",""
                                    CloseDiv
		    		        rstAux.close
			            else
                                    DrawDiv "hidden","",""
				            'DrawCelda "CELDA style='width:200px'","","",0,d_lookup("descripcion","tiendas","codigo='" & rst("tienda") & "'",session("dsn_cliente")) '**rgu 2/9/2009
                            EligeCeldaResponsive "text", mode, "hidden", "", "", 0, LitTienda, LitTienda, "",d_lookup("descripcion","tiendas","codigo='" & rst("tienda") & "'",session("dsn_cliente"))
                                    CloseDiv
				        end if
			        'CloseFila
		            end if
		    
		            'FLM:120309: cuenta de cargo y cuenta de abono del proveeedor.
		            'DrawFila color_blau
			            num_cuenta = ""
			            'Si tmp_nproveedor está lleno es que se ha cargado el proveedor. Si no,tomo lo que hay en bd.
			            if tmp_nproveedor&""<>"" then
			                num_cuenta = d_lookup("cuenta_cargo","proveedores","nproveedor='" & proveedor_aux & "'",session("dsn_cliente"))
			            else
			                num_cuenta=rst("ncuenta")
			            end if
			            if mode<>"browse" then
				            'DrawCelda2 "CELDA style='width:150px;'", "left", false, LitNCuentaCargo + ":"
				            rstSelect.cursorlocation=3
				            rstSelect.open "select distinct ncuenta from bancos with(nolock) where nbanco like '" & session("ncliente") & "%' order by ncuenta",session("dsn_cliente")
                                    DrawDiv "hidden","",""
				            DrawSelectCelda "hidden","","",0,LitNCuentaCargo,"ncuentacargo",rstSelect,num_cuenta,"ncuenta","ncuenta","",""
                                    CloseDiv
				            rstSelect.close
			            else
                                    DrawDiv "hidden","",""
				            'DrawCelda2 "ENCABEZADOL", "left", false,LitNCuentaCargo + ":&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				            'DrawCelda "CELDA","","",0,rst("ncuenta")
                            EligeCeldaResponsive "text", enc.EncodeForHtmlAttribute(null_s(mode)), "hidden", "", "", 0, LitNCuentaCargo, LitNCuentaCargo, "",rst("ncuenta")
                                    CloseDiv
			            end if
			            ''if mode<>"browse" then
                            'DrawCelda "CELDA","","",0,""
                        ''end if
		                'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:150px'","","",0,LitCuentaAbono+":"
			            if mode="browse" then
                                    DrawDiv "hidden","",""
				            'DrawCelda "'CELDALEFT' align='left' style='width:200px;'","","",0,rst("ncuenta_pro")&""
                            EligeCeldaResponsive "text", enc.EncodeForHtmlAttribute(null_s(mode)), "hidden", "", "", 0, LitCuentaAbono, LitCuentaAbono, "",rst("ncuenta_pro")&""
                                    CloseDiv
			            else
			                if mode="add" then
			                    num_cuenta = ""
			                    num_cuenta = d_lookup("ncuenta","proveedores","nproveedor='" & proveedor_aux & "'",session("dsn_cliente"))
			                else
                                num_cuenta=iif(tmp_nproveedor&""="",rst("ncuenta_pro")&"",d_lookup("ncuenta","proveedores","nproveedor='" & proveedor_aux & "'",session("dsn_cliente")))	         
			                end if
                            %><div class="hidden"><%
                                %><label><%=LitCuentaAbono %></label><%
						        %><input class="width60" maxlength="25" type="text" name="ncuenta_pro" value="<%=enc.EncodeForHtmlAttribute(null_s(num_cuenta))%>"/><%
                            %></div><%
                        end if
		            'CloseFila
		            'FLM:170309:Nombre del banco.Se pone, segun la cuenta de abono, el que corresponda en la bd de bancos.
		            'Lo pongo a disabled , ya que no se escribre, es meramente informativo.
		            'DrawFila color_blau
		                'DrawCelda "CELDA colspan='"&iif(mode="browse","3","3")&"'","","",0,""
		                'DrawCelda2 iif(mode="browse","ENCABEZADOL","CELDA")  & " style='width:150px;'", "left", false,LitBanco+":"
		                if mode="browse" then		        
                                DrawDiv "hidden","",""
				            'DrawCelda "'CELDALEFT' align='left'  style='width:200px;'","","",0,rst("banco")&""
                            EligeCeldaResponsive "text", enc.EncodeForHtmlAttribute(null_s(mode)), "hidden", "", "", 0, LitBanco, LitBanco, "",rst("banco")&""
                                CloseDiv
			            else
			                if(tmp_nproveedor&""="") then
			                    banco=rst("banco")&""
			                else
			                    banco=d_lookup("Entidad","bancos","codigo='" & left(trim(num_cuenta),4) & "'",DsnIlion)
                            end if
                            %><div class="hidden"><%
                                %><label><%=LitBanco%></label><%
						        %><input class="width60" disabled type="text" name="banco" value="<%=enc.EncodeForHtmlAttribute(null_s(banco))%>"/><%
                            %></div><%
			            end if
		            'CloseFila
			        'DrawFila color_blau
				        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:150px;'","","",0,LitIncotAlbPro+":"
				        if mode="browse" then
					        'DrawCelda "CELDA style='width:200px;'","","",0,rst("incoterms")&""
                                DrawDiv "hidden","",""
                            EligeCeldaResponsive "text", enc.EncodeForHtmlAttribute(null_s(mode)), "", "", "", 0, LitIncotAlbPro, LitIncotAlbPro, "",rst("incoterms")&""
                                CloseDiv
				        else
					        defecto=iif(tmp_incoterms>"",tmp_incoterms,iif(rst("incoterms")>"",rst("incoterms"),""))
                            rstAux.cursorlocation=3
					        rstAux.open "select codigo,codigo as descripcion from incoterms with(nolock) order by descripcion",session("dsn_cliente")
                                DrawDiv "hidden","",""
					        DrawSelectCelda "hidden","60","",0,LitIncotAlbPro,"incoterms",rstAux,defecto,"codigo","descripcion","",""
                                CloseDiv
					        rstAux.close
				        end if
				        ''if mode<>"browse" then 
                            'DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''end if
				        'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:180px'","","",0,LitIncoPuntEntrAlbPro+":"
				        if mode="browse" then
					        'DrawCelda "'CELDALEFT' align='left' ","","",0,rst("fob")&""
                                DrawDiv "hidden","",""
                            EligeCeldaResponsive "text", enc.EncodeForHtmlAttribute(null_s(mode)), "", "", "", 0, LitIncoPuntEntrAlbPro, LitIncoPuntEntrAlbPro, "",rst("fob")&""
                                CloseDiv
				        else
                            %><div class="hidden"><%
                                %><label><%=LitIncoPuntEntrAlbPro %></label><%
					            defecto=iif(tmp_fob>"",tmp_fob,iif(rst("fob")>"",rst("fob"),""))
						        %><input class="width60" maxlength="50" type="text" size="25" name="fob" value="<%=enc.EncodeForHtmlAttribute(defecto)%>"/><%
					        %></div><%
				        end if
			       ' CloseFila
			        if mode="browse" then
				        'DrawFila color_blau
					        'DrawCelda "ENCABEZADOL valign='top'","","",0,LitObservaciones+":&nbsp;"
					        'DrawCelda "CELDALEFT colspan='3'","","",0,pintar_saltos_espacios(rst("observaciones")&"")
                            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, LitObservaciones, LitObservaciones, "",pintar_saltos_espacios(rst("observaciones")&"")
				        'CloseFila
			        else
				        'DrawFila color_blau
					        'DrawCelda "CELDA","","",0,LitObservaciones+":"
					        'DrawTextCeldaSpan "CELDA","","",2,100,"","observaciones",iif(rst("observaciones")>"",rst("observaciones"),""),3
                            %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><%
                                %><label><%=LitObservaciones %></label><%
					            defecto=iif(tmp_fob>"",tmp_fob,iif(rst("fob")>"",rst("fob"),""))
						        DrawTextarea "width60","","observaciones",iif(rst("observaciones")>"",rst("observaciones"),""),""
					        %></div><%
				        'CloseFila
			        end if

	        '************************'
	        'JMA 20/12/04 ***********'
	        '************************'
	        if mode="browse" and si_campo_personalizables=1 then
		      	%><div class="hidden"><%
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                    rstAux2.cursorlocation=3
			        rstAux2.open "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
			        if not rstAux2.eof then
				        'DrawFila ""
					        num_campo=1
					        num_campo2=1
					        num_puestos=0
					        num_puestos2=0
					        while not rstAux2.eof
						        if num_puestos2>0 and (num_puestos2 mod 2)=0 then
                                    ''CloseFila
						            ''DrawFila color_titulo
							            ''DrawCelda "CELDA style='width:125px'","","",0,"&nbsp;"
                                        %>
							            <!--<td colspan=5>
								            <font class = "ENCABEZADOC">&nbsp;</font>
							            </td>--><%
							        'CloseFila
							        'DrawFila ""
							        num_puestos2=0
						        end if
						        if rstAux2("titulo") & "">"" then
							        num_puestos=num_puestos+1
							        num_puestos2=num_puestos2+1
                                    'if num_puestos2=2 then
                                        'DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                                    'end if
							        'DrawCelda "ENCABEZADOL style='width:155px'","","",0,rstAux2("titulo") & " : "
							        if rstAux2("tipo")=2 then
                                        DrawCeldaResponsive "CELDA align=left style='width:155px'","","",0, rstAux2("titulo") & ":",iif(lista_valores(num_campo)=1,LitSi,LitNo)
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
                                        DrawCeldaResponsive "CELDA align=left style='width:200px'","","",0,rstAux2("titulo") & ":",valor_ListCampPerso
							        else
                                        DrawCeldaResponsive "CELDA align=left style='width:200px'","","",0,rstAux2("titulo") & ":",lista_valores(num_campo)
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
					        'if num_puestos=1 then
						        'DrawCelda "CELDA style='width:150px'","","",0,"&nbsp;"
						        'DrawCelda "CELDA style='width:200px'","","",0,"&nbsp;"
						        ''DrawCelda "CELDA style='width:130px'","","",0,"&nbsp;"
					        'end if
				        'CloseFila
				        num_campos=num_puestos
			        else
				        num_campos=0
			        end if
			        rstAux2.close
	        elseif mode="add" and si_campo_personalizables=1 then
		      	%><div class="hidden"><%
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                    rstAux2.cursorlocation=3
			        rstAux2.open "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
			        if not rstAux2.eof then
				        num_campos_existen=rstAux2.recordcount
				        'DrawFila ""
					        num_campo=1
					        num_campo2=1
					        num_puestos=0
					        num_puestos2=0
					        while not rstAux2.eof
						        if num_puestos2>0 and (num_puestos2 mod 2)=0 then
							        ''DrawCelda "CELDA style='width:125px'","","",0,"&nbsp;"
							        'CloseFila
							        'DrawFila ""
							        num_puestos2=0
						        end if
						        if rstAux2("titulo") & "">"" then
							        if ((num_puestos-1) mod 2)=0 then
								        ''DrawCelda "CELDA style='width:155px'","","",0,"&nbsp;"
                                        'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"
							        end if
							        num_puestos=num_puestos+1
							        num_puestos2=num_puestos2+1
							        'DrawCelda "CELDA style='width:155px'","","",0,rstAux2("titulo") & " : "
							        valor_campo_perso=""

							        'JMA 20/12/04. Copiar campos personalizables de los proveedores'
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
                                        DrawLabel "","",enc.EncodeForHtmlAttribute(null_s(rstAux2("titulo"))) & ":"%><input type="text" style="width:155px;" class="CELDA" name="<%="campo" & enc.EncodeForHtmlAttribute(null_s(num_campo))%>" maxlength="<%=tamany%>" value="<%=enc.EncodeForHtmlAttribute(valor_campo_perso)%>"/>
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
								        strSelListVal="select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
                                        rstAux.cursorlocation=3
								        rstAux.open strSelListVal,session("dsn_cliente")
                                        DrawDiv "1","",""
                                        DrawLabel "","",rstAux2("titulo") & ":"
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
                                        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" style="width:155px;" class="CELDA" name="<%="campo" & enc.EncodeForHtmlAttribute(null_s(num_campo))%>" value="<%=enc.EncodeForHtmlAttribute(valor_campo_perso)%>"/>
								        <%CloseDiv
							        elseif rstAux2("tipo")=5 then
								        if isNumeric(rstAux2("tamany")) then
									        tamany=rstAux2("tamany")
								        else
									        tamany=1
								        end if
                                        DrawDiv "1","",""
                                        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" class="CELDA" name="<%="campo" & enc.EncodeForHtmlAttribute(null_s(num_campo))%>" style='width:155px' value="<%=enc.EncodeForHtmlAttribute(valor_campo_perso)%>"/>
								        <%CloseDiv
							        end if
						        else
							        %><input type="hidden" name="campo<%=enc.EncodeForHtmlAttribute(null_s(num_campo))%>" value=""/><%
						        end if
						        %><input type="hidden" name="tipo_campo<%=enc.EncodeForHtmlAttribute(null_s(num_campo))%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstAux2("tipo")))%>"/>
						        <input type="hidden" name="titulo_campo<%=enc.EncodeForHtmlAttribute(null_s(num_campo))%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstAux2("titulo")))%>"/><%
					            rstAux2.movenext
						        num_campo=num_campo+1
						        if not rstAux2.eof then
							        if rstAux2("titulo") & "">"" then
								        num_campo2=num_campo2+1
							        end if
						        end if
					        wend
					        'if num_puestos=1 then
						        'DrawCelda "CELDA style='width:150px'","","",0,"&nbsp;"
						        'DrawCelda "CELDA style='width:200px'","","",0,"&nbsp;"
						        'DrawCelda "CELDA style='width:130px'","","",0,"&nbsp;"
					        'end if
				        'CloseFila
				        num_campos=num_puestos
			        else
				        num_campos=0
				        num_campos_existen=0
			        end if
			        rstAux2.close
		        %><input type="hidden" name="num_campos" value="<%=num_campos_existen%>"/><%
	        elseif mode="edit" and si_campo_personalizables=1 then
		      	%><div class="hidden"><%
			      	%><label colspan="5" class="ENCABEZADOL"><%=LitCampPersoDocC%></label><%
				CloseDiv
                    rstAux2.cursorlocation=3
			        rstAux2.open "select * from camposperso with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("dsn_cliente")
			        if not rstAux2.eof then
				        num_campos_existen=rstAux2.recordcount
				        'DrawFila ""
					        num_campo=1
					        num_campo2=1
					        num_puestos=0
					        num_puestos2=0
					        while not rstAux2.eof
						        if num_puestos2>0 and (num_puestos2 mod 2)=0 then
							        ''DrawCelda "CELDA style='width:125px'","","",0,"&nbsp;"
							        'CloseFila
							        'DrawFila ""
							        num_puestos2=0
						        end if
						        if rstAux2("titulo") & "">"" then
							        'if ((num_puestos-1) mod 2)=0 then
								        ''DrawCelda "CELDA style='width:155px'","","",0,"&nbsp;"
                                        'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"
							        'end if
							        num_puestos=num_puestos+1
							        num_puestos2=num_puestos2+1
							        'DrawCelda "CELDA style='width:150px'","","",0,rstAux2("titulo") & " : "
							        valor_campo_perso=lista_valores(num_campo)

							        'JMA 20/12/04. Copiar campos personalizables de los proveedores'
							        if TraerProveedor > "" then
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
                                        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" class="CELDA" name="<%="campo" & enc.EncodeForHtmlAttribute(null_s(num_campo))%>" style='width:200px' value="<%=enc.EncodeForHtmlAttribute(null_s(valor_campo_perso))%>" />
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
								        strSelListVal="select ndetlista,valor from campospersolista with(nolock) where tabla='DOCUMENTOS COMPRA' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
                                        rstAux.cursorlocation=3
								        rstAux.open strSelListVal,session("dsn_cliente")
                                        'DrawSelect "","width:155px","campo"&num_campo,rstAux,valor_campo_perso,"ndetlista","valor","",""
                                        DrawDiv "1","",""
									    DrawLabel "","",rstAux2("titulo") & ":"
			 						    DrawSelect "","width:200px","campo"&num_campo,rstAux,valor_campo_perso,"ndetlista","valor","",""
                                        CloseDiv
			 					        rstAux.close
							        elseif rstAux2("tipo")=4 then
								        if isNumeric(rstAux2("tamany")) then
									        tamany=rstAux2("tamany")
								        else
									        tamany=1
								        end if
                                        DrawDiv "1","",""
                                        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" class="CELDA" name="<%="campo" & num_campo%>" style='width:155px'  value="<%=enc.EncodeForHtmlAttribute(null_s(valor_campo_perso))%>"/>
								        <%CloseDiv
							        elseif rstAux2("tipo")=5 then
								        if isNumeric(rstAux2("tamany")) then
									        tamany=rstAux2("tamany")
								        else
									        tamany=1
								        end if
                                        DrawDiv "1","",""
                                        DrawLabel "","",rstAux2("titulo") & ":"%><input type="text" class="CELDA" name="<%="campo" & num_campo%>" style="width:155px" value="<%=valor_campo_perso%>"/>
								        <%CloseDiv
							        end if
						        else
							        %><input type="hidden" name="campo<%=enc.EncodeForHtmlAttribute(null_s(num_campo))%>" value=""/><%
						        end if
						        %><input type="hidden" name="tipo_campo<%=enc.EncodeForHtmlAttribute(null_s(num_campo))%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstAux2("tipo")))%>"/>
						        <input type="hidden" name="titulo_campo<%=enc.EncodeForHtmlAttribute(null_s(num_campo))%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rstAux2("titulo")))%>"/><%
					            rstAux2.movenext
						        num_campo=num_campo+1
						        if not rstAux2.eof then
							        if rstAux2("titulo") & "">"" then
								        num_campo2=num_campo2+1
							        end if
						        end if
					        wend
					        'if num_puestos=1 then
						        'DrawCelda "CELDA style='width:150px'","","",0,"&nbsp;"
						        'DrawCelda "CELDA style='width:200px'","","",0,"&nbsp;"
						        'DrawCelda "CELDA style='width:130px'","","",0,"&nbsp;"
					        'end if
				        'CloseFila
				        num_campos=num_puestos
			        else
				        num_campos=0
				        num_campos_existen=0
			        end if
			        rstAux2.close
		        %><input type="hidden" name="num_campos" value="<%=enc.EncodeForHtmlAttribute(num_campos_existen)%>"/><%
	        end if
	        '************************'
	        'FIN JMA 20/12/04 *******'
	        '************************'
            %></table>
            </div>
		</div><%
		if mode="browse" then
			'*********** Dirección de envío'
            rstDomi.cursorlocation=3
			rstDomi.Open "select * from domicilios with(nolock) where codigo='" & rst("dir_envio") & "'",session("dsn_cliente")
			if rst("dir_envio")>"" then
				pagina="../central.asp?pag1=./compras/albaranes_prodireccion_env.asp&ndoc=" &rst("nalbaran") & "&mode=browse&pag2=./compras/albaranes_prodireccion_env_bt.asp&titulo=" & ucase(LITDATOSENVIO) & " " & rst("nalbaran_pro")
			else
				pagina="../central.asp?pag1=./compras/albaranes_prodireccion_env.asp&ndoc=" &rst("nalbaran") & "&mode=edit&pag2=./compras/albaranes_prodireccion_env_bt.asp&titulo=" & ucase(LITDATOSENVIO) & " " & rst("nalbaran_pro")
			end if

			'** Campo oculto para controlar si el albarán está facturado.
            rstAux.cursorlocation=3
			rstAux.Open "select f.nfactura_pro from albaranes_pro as a with(nolock),facturas_pro as f with(nolock) where a.nalbaran='" & rst("nalbaran") & "' and a.nfactura=f.nfactura and f.nfactura like '" & session("ncliente") & "%' ",session("dsn_cliente")
			if not rstAux.eof then
				if not isnull(rstAux("nfactura_pro")) then
					%><input type="hidden" name="h_nfactura" value="<%=enc.EncodeForHtmlAttribute(null_s(rstAux("nfactura_pro")))%>"/><%
				else
					%><input type="hidden" name="h_nfactura" value="NO"/><%
				end if
			else
				%><input type="hidden" name="h_nfactura" value="NO"/><%
			end if
			rstAux.close

            '** Configuración vertical de la empresa
            set conn2 = server.CreateObject("ADODB.Connection")
            set command2  = server.CreateObject("ADODB.Command")
            set rstVert = Server.CreateObject("ADODB.Recordset")
            conn2.open session("dsn_cliente")
            command2.ActiveConnection = conn2
	        command2.CommandTimeout = 0              
            command2.CommandText="GetAllVerticalConfiguration" 
	        command2.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	        command2.Parameters.Append command2.CreateParameter("@NCOMPANY",adVarChar,adParamInput,5,session("ncliente")&"")
            set rstVert = command2.Execute
            if not rstVert.EOF then
                si_tiene_cant_esp = rstVert("expectedamount")
                si_tiene_importe = rstVert("amounts")
            else
                si_tiene_cant_esp = 0
                si_tiene_importe = 0
            end if
            conn2.Close
            set conn2 = nothing
            set command2 = nothing
            %><div class="Section" id="S_FINANCIAL_DATA" >
                <a href="#" rel="toggle[FINANCIAL_DATA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LIT_FINANCIAL_DATA %>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
                <div class="SectionPanel" id="FINANCIAL_DATA">
                    <div id="tabs" style="display:none">
                        <ul>
                            <li><a href="#tabs-details"><%=LitDetalles %></a></li>
                            <li class="hidden"><a href="#tabs-concepts"><%=LitConceptos %></a></li>
                            <%if vencesp="1" then%>
                                <li class="hidden" id="li-vencimientos"><a href="#tabs-vencimientos"><%=LitVencimientos%></a></li>
                            <%end if%>
                            <li class="hidden" id="li-payments" ><a href="#tabs-payments"><%=LitPagosACuenta %></a></li>
                            <li class="hidden"><a href="#tabs-send"><%=LitDatosEnvio%></a></li>
                        </ul><%

                        if oculta=0 then%>
                            <div id="tabs-details" class="overflowXauto">
		   	                    <table class="width90 md-table-responsive bCollapse" ><%
			                        DrawFila ""'color_terra
				                       %><td class='CELDAR7 underOrange width5'></td>
				                        <td class='CELDAL7 underOrange width5'><b><%=LitCantidad%></b></td>
                                       <%if si_tiene_cant_esp = 1 or si_tiene_cant_esp = true then%>
                                        <td class='CELDAL7 underOrange width5'><b><%=LITCANTESPERADA%></b></td>
                                       <%end if
                                       if si_tiene_acceso_almacenes=1 then%>
				                           <td class='CELDAL7 underOrange width10'><b><%=LitTanque%></b></td>
				                       <%end if%>
				                       <td class='CELDAL7 underOrange width5'><b><%=LitReferencia%></b></td>
				                       <td class='CELDAL7 underOrange width15'><b><%=LitDescripcion%></b></td>
                                       <%if si_tiene_importe = 1 or si_tiene_importe = true then%>
				                           <td class='CELDAR7 underOrange width5'><b><%=LitPVP%></b></td>
				                           <td class='CELDAR7 underOrange width5'><b><%=LitIva%></b></td>
				                           <td class='CELDAR7 underOrange width5'><b><%=LitImporte%></b></td>
                                       <%end if%>
				                       <td class='CELDAR7 underOrange width5'></td>
			                        <%CloseFila%>
				                </table>
				                <%if isnull(rst("nfactura")) and isnull(rst("nbalance")) then
				                    if si_tiene_modulo_ebesa<>0 then%>
					                <iframe id='frDetallesIns' name='fr_DetallesIns' class="width90 iframe-input md-table-responsive" src='albaranes_prodetins_pista.asp?ndoc=<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran")))%>&nserie=<%=enc.EncodeForHtmlAttribute(null_s(rst("serie")))%>&nproveedor=<%=enc.EncodeForHtmlAttribute(null_s(rst("nproveedor")))%>&tienda=<%=enc.EncodeForHtmlAttribute(null_s(trimCodEmpresa(rst("tienda"))))%>&modp=<%=enc.EncodeForHtmlAttribute(modp)%>&modd=<%=enc.EncodeForHtmlAttribute(modd)%>&modi=<%=enc.EncodeForHtmlAttribute(modi)%>&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>' frameborder="no" scrolling="no" noresize="noresize"></iframe>
					                <%else%>
					                <iframe id='frDetallesIns' name='fr_DetallesIns' class="width90 iframe-input md-table-responsive" src='albaranes_prodetins_pista.asp?ndoc=<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran")))%>&nserie=<%=enc.EncodeForHtmlAttribute(null_s(rst("serie")))%>&nproveedor=<%=enc.EncodeForHtmlAttribute(null_s(rst("nproveedor")))%>&tienda=<%=enc.EncodeForHtmlAttribute(null_s(trimCodEmpresa(rst("tienda"))))%>&modp=<%=enc.EncodeForHtmlAttribute(modp)%>&modd=<%=enc.EncodeForHtmlAttribute(modd)%>&modi=<%=enc.EncodeForHtmlAttribute(modi)%>&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>' frameborder="no" scrolling="no" noresize="noresize"></iframe>
					                <%end if%>
				                <%end if%>
				                <iframe id='frDetalles' name="fr_Detalles" class="width90 md-table-responsive" src='albaranes_prodet_pista.asp?ndoc=<%=rst("nalbaran")%>&nserie=<%=rst("serie")%>&nproveedor=<%=rst("nproveedor")%>&tienda=<%=trimCodEmpresa(rst("tienda"))%>&modp=<%=enc.EncodeForHtmlAttribute(modp)%>&modd=<%=enc.EncodeForHtmlAttribute(modd)%>&modi=<%=enc.EncodeForHtmlAttribute(modi)%>&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>' height='150' frameborder="yes" noresize="noresize"></iframe>
			                    <span id="paginacion" style="display: "></span>
                            </div>
                        <%end if 'del oculta

			            ''ricardo 9/8/2004 se pondra el iva que tiene establecido el cliente
			            TmpIvaProveedor=d_lookup("iva","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))
			            defaultIVA=d_lookup("iva","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
			            if TmpIvaProveedor & "">"" then
				            TmpIva=TmpIvaProveedor
			            else
				            TmpIva=defaultIVA
			            end if%>

                        <div  id="tabs-concepts" class="overflowXauto hidden" >
				            <input type="hidden" name="defaultIva" value="<%=enc.EncodeForHtmlAttribute(null_s(TmpIva))%>"/>
		   	                <table class="width90 md-table-responsive bCollapse">
		   	                <%'Fila de encabezado
				            DrawFila color_terra%>
					            <td class='CELDAL7 underOrange width5'><%=LitItem%></td>
					            <td class='CELDAR7 underOrange width5'><%=LitCantidad%></td>
					            <td class='CELDAL7 underOrange width15'><%=LitDescripcion%></td>
					            <td class='CELDAR7 underOrange width5 hidden'><%=LitPVP%></td>
					            <td class='CELDAR7 underOrange width5 hidden'><%=LitDto%></td>
					            <td class='CELDAR7 underOrange width5 hidden'><%=LitIva%></td>
					            <td class='CELDAR7 underOrange width5 hidden'><%=LitImporte%></td>
					            <td class='CELDAR7 underOrange width10'>&nbsp</td>
				            <%CloseFila
				            if isnull(rst("nfactura")) and isnull(rst("nbalance")) then
					            'Linea de inserción de un detalle
					            DrawFila color_blau
						            %><td class='CELDAL7 underOrange width5'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						            <td class='CELDAR7 underOrange width5'>
							            <input class='CELDAR7 width100' type="text" name="cantidad" value="1" onchange="RoundNumValue(this,<%=enc.EncodeForHtmlAttribute(null_s(DEC_CANT))%>);ImporteDetalle();"/>
						            </td>
						            <td class='CELDAL7 underOrange width15'>
							            <textarea class='CELDAL7 width100' name="descripcion" rows="2"></textarea>
						            </td>
						            <td class='CELDAR7 underOrange width5'>
							            <input class='CELDAR7 width100' type="text" name="pvp" value="0" onchange="RoundNumValue(this,<%=enc.EncodeForHtmlAttribute(null_s(dec_prec))%>);ImporteDetalle();"/>
						            </td>
						            <td class='CELDAR7 underOrange width5'>
							            <input class='CELDAR7 width100' type="text" name="descuento" value="0" onchange="RoundNumValue(this,<%=enc.EncodeForHtmlAttribute(null_s(decpor))%>);ImporteDetalle();"/>
						            </td>
						            <%
                                    rstSelect.cursorlocation=3
						            rstSelect.open "select tipo_iva, tipo_iva from tipos_iva with(nolock)",session("dsn_cliente")
						            DrawSelectCeldaDet "'CELDAR7 underOrange width5'","width100","",0,"","iva",rstSelect,TmpIva,"tipo_iva","tipo_iva","",""
						            rstSelect.close
						            %><td class='CELDAR7 underOrange width5'>
							            <input class='CELDAR7 width100' disabled type="text" name="importe" value="0"/>
						            </td>
						            <td class=" underOrange width10 ">
							            <a class="ic-accept noMTop" href="javascript:addConcepto('<%=enc.EncodeForJavascript(nalbaran)%>');" onblur="javascript:document.albaranes_pro_pista.cantidad.focus();"><img src="<%=themeIlion%><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>
						            </td><%
						            if oculta=1 then%>
						              <script language="javascript" type="text/javascript">
                                          document.albaranes_pro_pista.cantidad.focus();
                                          document.albaranes_pro_pista.cantidad.select();
						               </script><%
						            end if
					            CloseFila
				            end if
				            %></table>
				            <iframe id="frConceptos" name="fr_Conceptos" class="width90 md-table-responsive" src='albaranes_procon.asp?mode=browse&ndoc=<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran")))%>' frameborder="yes" noresize="noresize"></iframe>
			            </div>
                        <%if vencesp="1" then%>
                            <div class="hidden" id="tabs-vencimientos">
			                    <table style="border-collapse:collapse;table-layout:fixed;width:355px" cellpadding="1" cellspacing="1" ><%
				                    'Fila de encabezado
					                    DrawFila color_terra
                                            %>
                                            <td class='CELDAL7' style="width:50px;">&nbsp</td>
						                    <td class='CELDAL7' style="width:130px;"><%=LitFecha%></td>
                                            <td class='CELDAL7' style="width:50px;"><%=LITDIASFF%></td>
						                    <td class='CELDAL7' style="width:80px;">%</td>
						                    <td class='CELDAL7' style="width:25px;">&nbsp</td>
                                            <%
					                    CloseFila
                                        if isnull(rst("nfactura")) and isnull(rst("nbalance")) then
						                    DrawFila color_blau
							                    %>
							                        <td class='CELDAL7' style="width:50px;">
							                        </td>
							                        <td class='CELDAL7' style="width:130px;">
								                        <input class='CELDAL7' type="text" name="fechaVto" style="width: 100px;" value="" onchange="javascript:cambiarFecVto()"/>
							                        </td>
                                                    <%DrawCalendar "fechaVto"%>
							                        <td class='CELDAL7' style="width:50px;">
								                        <input class='CELDAL7' type="text" name="DiasFFVto" style="width: 45px;" value="" onchange="javascript:cambiarDiasFFVto()"/>
							                        </td>
							                        <td class='CELDAL7' style="width:80px;">
								                        <input class='CELDAL7' type="text" name="tantoVto" style="width: 78px;" value="0" onchange="javascript:RoundNumValue(this,<%=NdecDiAlbaran%>);"/>
							                        </td>
							                        <td class='CELDAL7' style="width:25px;">
								                        <a class="ic-accept noMTop" href="javascript:addVencimiento('<%=enc.EncodeForJavascript(null_s(rst("nalbaran")))%>');" onblur="javascript:document.albaranes_pro_pista.fechaVto.focus();"><img src="<%=themeIlion%><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>
							                        </td>
                                                <%
						                    CloseFila
                                        end if
                                        %>
					                    <!--</td>
					                    </tr>-->
				                    </table><% 'FIN DE TABLA QUE CONTIENE LA TABLAS DE ARTICULOS
				                    %><iframe id="frVencimientos" name="fr_Vencimientos" class="width90 md-table-responsive" src='vencimientos_pro_config.asp?mode=browse&tdocumento=albaranes_pro&ndoc=<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran")))%>' width='410' height='200' frameborder="yes" noresize="noresize"></iframe>
                            </div>
                        <%end if%>
                        <div  id="tabs-payments" class="overflowXauto hidden">
				            <table class="width90 md-table-responsive bCollapse"><%
				            'Fila de encabezado
				            DrawFila color_terra
					            %><td class='CELDAL7 underOrange width5'><%=LitNumPago%></td>
					            <td class='CELDAR7 underOrange width5'><%=LitFecha%></td>
					            <td class='CELDAL7 underOrange width15'><%=LitDescripcion%></td>
					            <td class='CELDAR7 underOrange width5'><%=LitImporte%></td>
					            <td class='CELDAR7 underOrange width5'><%=LitTipoPago%></td>
					            <td class='CELDAL7 underOrange width10'>&nbsp</td><%
				            CloseFila
				            if isnull(rst("nfactura")) and isnull(rst("nbalance")) then
					            'Linea de inserción de un pago a cuenta
					            DrawFila color_blau
						            %><td class='CELDAL7 underOrange width5'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						            <td class='CELDAR7 underOrange width5'>
							            <input class='CELDAR7 width70' type="text" name="fechaPago" value=""/><%
                                            DrawCalendar "fechaPago"%>
						            </td>
						            <td class='CELDAL7 underOrange width15'>
							            <textarea class='CELDAL7 width100' name="descripcionPago" onFocus="lenmensaje(this,0,50,'')" onKeydown="lenmensaje(this,0,50,'')" onKeyup="lenmensaje(this,0,50,'')" onBlur="lenmensaje(this,0,50,'')" rows="2"></textarea>
						            </td>
						            <td class='CELDAR7 underOrange width5'>
							            <input class='CELDAR7 width100' type="text" name="importePago" value="0" onchange="RoundNumValue(this,<%=enc.EncodeForJavascript(null_s(NdecDiAlbaran))%>);importepagoComp();"/>
						            </td><%
                                    rstSelect.cursorlocation=3
						            rstSelect.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by codigo",session("dsn_cliente")
						            DrawSelectCeldaDet "'CELDAL7 underOrange width5'","width100","",0,"","tipoPago",rstSelect,"","codigo","descripcion","",""
						            rstSelect.close
						            %><td class="underOrange width10">
							            <a class="ic-accept noMTop" href="javascript:addPago('<%=enc.EncodeForJavascript(nalbaran)%>');" onblur="javascript:document.albaranes_pro_pista.fechaPago.focus();"><img src="<%=themeIlion%><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>
						            </td><%
					            CloseFila
				            end if
				            %></table>
				            <iframe id="frPagosCuenta" name="fr_PagosCuenta" class="width90 md-table-responsive" src='albaranes_propagos.asp?mode=browse&ndoc=<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran")))%>' width='650' height='80' frameborder="yes" noresize="noresize"></iframe>
			            </div>
                        <div id="tabs-send" class="overflowXauto hidden">
                            <table class="width90 md-table-responsive bCollapse">
			                    <%'DrawFila color_terra%>
                                <tr>
				                    <td>
  				                        <table BORDER="1" cellspacing="0" cellpadding="0">
  				                            <%'DrawFila color_blau2%>
                                            <tr>
  					                            <td CLASS=ENCABEZADOC height="25">	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=LitDatosEnvio%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  				                                    <%if isnull(rst("nfactura")) then%>
  						                                <a class='CELDAREFB'  href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(pagina)%>','P',<%=enc.EncodeForJavascript(altoventana)%>,<%=enc.EncodeForJavascript(anchoventana)%>)" OnMouseOver="self.status='<%=LitEditar%>'; return true;" OnMouseOut="self.status=''; return true;">
  						                                <%=LitEditar%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  					                                <%end if%>
  					                            </td>
                                            </tr>
  				                            <%'CloseFila%>
				                          </table>
                                    </td>
					            </tr>
			                    <%'CloseFila
			                    if not rstDomi.eof then
				                    'DrawFila color_terra%>
                                    <tr>
					                    <td>
    				                      <table width='100%' border='0' cellspacing="1" cellpadding="1">
    				                            <%'DrawFila color_terra
							                        'DrawCelda "ENCABEZADOL underOrange width25","","",0,LitDomicilio
							                        'DrawCelda "ENCABEZADOL underOrange width25","","",0,LitPoblacion
							                        'DrawCelda "ENCABEZADOL underOrange width25","","",0,LitCP
							                        'DrawCelda "ENCABEZADOL underOrange width25","","",0,LitProvincia
						                        'CloseFila %>

                                                <tr>
                                                    <td class="ENCABEZADOL"><%=LitDomicilio%></td>
                                                    <td class="ENCABEZADOL"><%=LitPoblacion%></td>
                                                    <td class="ENCABEZADOL"><%=LitCP%></td>
                                                    <td class="ENCABEZADOL"><%=LitProvincia%></td>
                                                </tr>

						                        <%'DrawFila color_blau2
							                        'DrawCelda "CELDA underOrange width25","","",0,rstDomi("domicilio")
							                        'DrawCelda "CELDA underOrange width25","","",0,rstDomi("poblacion")
							                        'DrawCelda "CELDA underOrange width25","","",0,rstDomi("cp")
							                        'DrawCelda "CELDA underOrange width25","","",0,rstDomi("provincia")
						                        'CloseFila%>

                                                <tr>
                                                    <td class="CELDA"><%=rstDomi("domicilio")%></td>
                                                    <td class="CELDA"><%=rstDomi("poblacion")%></td>
                                                    <td class="CELDA"><%=rstDomi("cp")%></td>
                                                    <td class="CELDA"><%=rstDomi("provincia")%></td>
                                                </tr>
						                    </table>
					                    </td>
                                    </tr>
				                    <%'Closefila
				                end if
				                rstDomi.close%>
				            </table>
                        </div>
		            <%end if%>
                </div>
		    </div> 
        </div>       
		<%'Mostrar los precios de la factura.%>
		<input type="hidden" name="sumadet" value="<%=enc.EncodeForHtmlAttribute(sumadet)%>"/>
		<input type="hidden" name="sumaRE" value="<%=enc.EncodeForHtmlAttribute(sumaRE)%>"/>
		<input type="hidden" name="importe_bruto2" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("importe_bruto")))%>"/>
		
        <div class="Section" id="S_TOTAL" style="visibility: <%=iif(si_tiene_importe = 1 or si_tiene_importe = true, "","hidden")%>;">
            <div class="SectionHeader2">
                <%=LitAbrevia %>
            </div>
            <div class="SectionPanel" id="DATTOTAL" >
		            <%DrawDiv "4", "", ""
                        DrawLabel "", "",LitAbrevia
        		        DrawSpan "ENCABEZADOL","",d_lookup("abreviatura","divisas",iif(tmp_divisa>"","codigo='" & tmp_divisa & "'",iif(mode="add","moneda_base<>0 and codigo like '" & session("ncliente") & "%'","codigo='" & rst("divisa") & "'")),session("dsn_cliente")),""
                    CloseDiv%>
                    <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitBruto
				            'EligeCelda "input", mode,iif(mode="browse","ENCABEZADOR id='importe_bruto'","ENCABEZADOR disabled"),"","",0,"","importe_bruto",10,formatnumber(null_z(rst("importe_bruto")),n_decimales,-1,0,iif(mode="browse",-1,0))
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR", "","importe_bruto",formatnumber(null_z(rst("importe_bruto")),n_decimales,-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='ImpBruto'","disabled")
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitDto
                            if mode<>"browse" then
					            %><input class="ENCABEZADOR" type="text" name="dto1" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_dto1>0,tmp_dto1,iif(rst("descuento")>"",rst("descuento"),0))))%>" onchange="RoundNumValue(this,<%=decpor%>);Precios();"/><%
				            else
					            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","dto1",cstr(formatnumber(rst("descuento"),decpor,-1,0,iif(mode="browse",-1,0))) + "%", "id='Dto'"
				            end if
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "", "",LitDto2
                            if mode<>"browse" then
					            %><input class=ENCABEZADOR type="text" name="dto2" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_dto2>0,tmp_dto2,iif(rst("descuento2")>"",rst("descuento2"),0))))%>" onchange="RoundNumValue(this,<%=decpor%>);Precios();"/><%
				            else
					            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","dto2",cstr(formatnumber(null_z(rst("descuento2")),decpor,-1,0,iif(mode="browse",-1,0))) + "%", "id='Dto2'"
				            end if
                        CloseDiv%>
			            <%DrawDiv "4", "", ""
                            DrawLabel "","",LitTotalDescuento
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_descuento",formatnumber(null_z(rst("total_descuento")),n_decimales,-1,0,iif(mode="browse",-1,0)),iif(mode="browse","ID='TotalDto'","disabled")
				            %><input type="hidden" name="h_total_descuento" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("total_descuento")))%>"/><%
                        CloseDiv%>
                        <%DrawDiv "4", "", ""
                            DrawLabel "","",LitImponible
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","base_imponible",formatnumber(null_z(rst("base_imponible")),n_decimales,-1,0,iif(mode="browse",-1,0)),iif(mode="browse","ID='BImponible'","disabled")
				            %><input type="hidden" name="h_base_imponible" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("base_imponible")))%>"/><%
                        CloseDiv%>
                        <%DrawDiv "4", "", ""
                            DrawLabel "","",LitTotalIva
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_iva",cstr(formatnumber(Null_z(rst("total_iva")),n_decimales,-1,0,iif(mode="browse",-1,0))),iif(mode="browse","ID='TotalIva'","disabled")
				            %><input type="hidden" name="h_total_iva" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("total_iva")))%>"/><%
                        CloseDiv%>
                        <%DrawDiv "4", "", ""
                            DrawLabel "","",LitTotalRecargo
                            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_re",cstr(formatnumber(Null_z(rst("total_re")),n_decimales,-1,0,iif(mode="browse",-1,0))),iif(mode="browse","ID='TotalRE'","disabled")
				            %><input type="hidden" name="h_total_re" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("total_re")))%>"/><%
                        CloseDiv%>
                        <%if ((rst("recargo")<>0) or mode="edit" or mode="add") then%>
					        <%DrawDiv "4", "", ""
                                DrawLabel "","",LitRF
                                if mode<>"browse" then
						            %><input CLASS=ENCABEZADOR type="text" name="rf" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_rf>0,tmp_rf,iif(rst("recargo")>"",rst("recargo"),0))))%>" onchange="RoundNumValue(this,<%=enc.EncodeForHtmlAttribute(null_s(decpor))%>);Precios();"/><%
					            else
						            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","rf",cstr(formatnumber(rst("recargo"),decpor,-1,0,iif(mode="browse",-1,0))) + "%","ID='RF'"
					            end if%>
                            <%CloseDiv%>
                            <%DrawDiv "4", "", ""
                                DrawLabel "","",LitTotalRF
					            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_rf",formatnumber(null_z(rst("total_recargo")),n_decimales,-1,0,iif(mode="browse",-1,0)),iif(mode="browse","ID='TotalRF'","disabled")
                                %><input type="hidden" name="h_total_rf" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("total_recargo")))%>"/><%
                            CloseDiv%>
				        <%end if
				        if ((rst("irpf")<>0) or mode="edit" or mode="add") then%>
					        <%DrawDiv "4", "", ""
                                DrawLabel "","",Litirpf
                                    if mode<>"browse" then
						                %><input CLASS=ENCABEZADOR type="text" name="irpf" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_irpf>0,tmp_irpf,iif(rst("irpf")>"",rst("irpf"),0))))%>" onchange="RoundNumValue(this,<%=enc.EncodeForHtmlAttribute(null_s(decpor))%>);Precios();"/><%
					                else
						                EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","irpf",cstr(formatnumber(rst("irpf"),decpor,-1,0,iif(mode="browse",-1,0))) + "%","ID='irpf'"
					                end if
                                CloseDiv%>
                                <%DrawDiv "4", "", ""
                                    DrawLabel "","",LitTotalirpf
                                    EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_irpf",formatnumber(null_z(rst("total_irpf")),n_decimales,-1,0,iif(mode="browse",-1,0)),iif(mode="browse","ID='Totalirpf'","disabled")
					                %><input type="hidden" name="h_total_irpf" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("total_irpf")))%>" /><%
                                CloseDiv%>
				        <%end if%>
                        <%DrawDiv "4", "", ""
                            DrawLabel "","",LitTotal
				            EligeCeldaResponsive1 "input", mode,"ENCABEZADOR","","total_albaran",formatnumber(null_z(rst("total_albaran")),n_decimales,-1,0,iif(mode="browse",-1,0)),iif(mode="browse","ID='TotalAlbaran'","disabled")
				            %><input class="CELDA" type="hidden" name="IRPF_Total" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_IRPF_Total>"",tmp_IRPF_Total,iif(isnull(rst("IRPF_Total")),"",rst("IRPF_Total")))))%>"/><%
				            %><input type="hidden" name="h_total_albaran" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("total_albaran")))%>"/><%
                        CloseDiv%>

			        <%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")) then
				        'DrawFila color_blau 'PTAS
					        DIVISA=iif(tmp_divisa>"",tmp_divisa,iif(mode="add",d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")),rst("divisa")))
					        DrawCelda "ENCABEZADOL ID='TotalEq'","","",0,LitPTAS
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='ImpBrutoEq'","ENCABEZADOR disabled"),"","",0,"","Pimporte_bruto",10,formatnumber(CambioDivisa(null_z(rst("importe_bruto")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='DtoEq'","ENCABEZADOR disabled"),"","",0,"","Pdto1",5, formatnumber(null_z(rst("descuento")),decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%","")
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='Dto2Eq'","ENCABEZADOR disabled"),"","",0,"","Pdto2",5, formatnumber(null_z(rst("descuento2")),decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%","")
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='TotalDtoEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_descuento",10,formatnumber(CambioDivisa(null_z(rst("total_descuento")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='BImponibleEq'","ENCABEZADOR disabled"),"","",0,"","Pbase_imponible",10,formatnumber(CambioDivisa(null_z(rst("base_imponible")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='TotalIvaEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_iva",10,formatnumber(CambioDivisa(null_z(rst("total_iva")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='TotalREEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_re",10,formatnumber(CambioDivisa(null_z(rst("total_re")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					        if ((rst("recargo")<>0) or mode="edit" or mode="add") then
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='RFEq'","ENCABEZADOR disabled"),"","",0,"","rf1",5, cstr(formatnumber(null_z(rst("recargo")),decpor,-1,0,iif(mode="browse",-1,0))) & iif(mode="browse","%","")
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='TotalRFEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_rf",10,formatnumber(CambioDivisa(null_z(rst("total_recargo")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					        end if
					        if ((rst("irpf")<>0) or mode="edit" or mode="add") then
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='irpfEq'","ENCABEZADOR disabled"),"","",0,"","irpf1",5, cstr(formatnumber(null_z(rst("irpf")),decpor,-1,0,iif(mode="browse",-1,0))) & iif(mode="browse","%","")
						        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='TotalirpfEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_irpf",10,formatnumber(CambioDivisa(null_z(rst("total_irpf")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					        end if
					        EligeCelda "input", mode,iif(mode="browse","CELDARIGHT ID='TotalAlbaranEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_albaran",15,formatnumber(CambioDivisa(null_z(rst("total_albaran")),DIVISA,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
				        'CloseFila
			        end if%>
            </div>
        </div><%

		if mode="add" then%>
			<script language="javascript" type="text/javascript">
                                          document.albaranes_pro_pista.fecha.focus();
                                          document.albaranes_pro_pista.fecha.select();
			</script>
		<%elseif mode="edit" then%>
			<script language="javascript" type="text/javascript">
                                           document.albaranes_pro_pista.tipo_pago.focus();
			</script>
		<%end if
	
	end if
	if mode="add" then rst.CancelUpdate%>
	<input type="hidden" name="total_paginas" value="<%=enc.EncodeForHtmlAttribute(total_paginas)%>"/>
</form>
<%'Mostrar la barra de pestañas
		BarraNavegacion mode

        CerrarTodo()
end if
%>
</body>
</html>