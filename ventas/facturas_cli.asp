<%@ Language=VBScript %>
<%
    response.buffer=true
    Server.ScriptTimeout = 1200 %>
<%

dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
    'celia
''CODIGOS DE AÑADIDURAS/MODIFICACIONES -----------------------------------------------------
''ricardo 7-4-2006 se comprueba el like de todas las tablas del mismo select
'' VGR 25-02-2003 : Añadir el campo AGENTE en la cabecera del documento
'' VGR 04-03-2003 : Cambios para no +editar la factura si ya está liquidada para su comercial
'' VGR 07-03-2003 : Cambios para no editar la factura si ya está liquidada para su agente
''ricardo 12-3-2003 : se pone la configuracion de documentos
''ricardo 3-6-2003 se pone por parametro que pregunte el cambio del comercial de los vencimientos
''ricardo 5-6-2003 se añade el parametro caju para que solo se pongan las cajas que se diga en el parametro
''ricardo 5-6-2003 se añade el parametro novei para que en los formatos de impresion no salga el item
''ricardo 1-11-2003 se añaden los parametros para cuando se llame con la funcion abrirventana
''ricardo 17-5-2004 se añade el parametro ocb para que no salga el coste/beneficio por documento
''ricardo 02-08-2004 se añade el parametro bcc para bloquear los campos contabilizado y cobrada
''ricardo 20-2-2006 se cambia la condicion para poner el dto general para que en lugar de (tmp_dto1>=0) sea (tmp_dto1>=0 and tmp_dto1<>"" and tmp_dto1<>NULL)
''RGU 27-03-2006 : Añadir gestion del parametro de usuario tfi
'***RGU 26/4/2006 parametro nmc Si = 1 no se puede modificar el comercial si ya hay uno asignado, si no hay ningun o si que se puede modificar

'' IVM 16/06/03: Migración monobase

''ricardo 31/7/2003 comprobamos que existe el albaran que se ha pedido ver desde un listado, sino se va al modo add
'' JPP 20-05-2004 se cambia la forma de introducir el hipervinculo al archivo de la factura. Pasa de ser texto a archivo

'JCI 05/08/2004 : Recuperar precios de artículo con el descuento aplicado o no en funcion de
'                 un parámetro de configuración
''RGU	26/10/2007: Parametro pagsl=1 => no se pueden modificar facturas ya generadas
' FLM : 19/01/2009 : Añadir captura de ncliente por queryString
' jcg   20/01/2009: Añadida la columna proyecto al cliente y tratamiento de la misma.
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
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
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../common/modal2.inc" -->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../common/campospersoResponsive.inc" -->
<!--#include file="../js/animatedCollapse.js.inc"-->
<!--#include file="../js/tabs.js.inc"-->
<!--#include file="../styles/generalData.css.inc"-->
<!--#include file="../styles/Section.css.inc"-->
<!--#include file="../styles/ExtraLink.css.inc"-->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="facturas_cli.inc" -->
<!--#include file="../varios2.inc" -->
<!--#include file="../perso.inc" -->
<!--#include file="ventas.inc" -->
<!--#include file="documentos.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="riesgo.inc" -->

<!--#include file="../styles/formularios.css.inc" -->

<!--#include file="../js/dropdown.js.inc" -->
<!--#include file="../common/facturas_cliActionDrop.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->
<!--#include file="../common/poner_cajaResponsive.inc" -->

<!--DGB RESPONSIVE-->
    

<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">

    animatedcollapse.addDiv('DATFINAN', 'fade=1');
    animatedcollapse.addDiv('DATTOTAL', 'fade=1');
    animatedcollapse.addDiv('DatosGenerales', 'fade=1');

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
</head>
<%ncompany_GALP=""
    folder=Session("folder")&""
    if folder="" then 
	    folder="ilion"
    end if
    themeIlion="/lib/estilos/" & folder & "/"

    style =""
    if session("ncliente") = ncompany_GALP then
        style = " style=""display:none;"
    end if
    set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion

    'mmg: variables para obtener los almacenes por defecto
    dim almacenSerie
    dim almacenTPV

    linea1=session("f_tpv")
    linea2=session("f_caja")
    mejico = "0"
    oculta = "0"
    'Calculamos el almacen por defecto del TPV
    set rstMM = Server.CreateObject("ADODB.Recordset")
    set rsTPV = Server.CreateObject("ADODB.Recordset")

    AbrirModal "fr_NonPayment","",600,240,"no","si","no","si","Gestión de Impagos"		                 						  

    'ega 19/06/2008 union de las tablas con join y with(nolock) y likes
    set connDom = Server.CreateObject("ADODB.Connection")
    set commandDom = Server.CreateObject("ADODB.Command")

    connDom.open session("dsn_cliente")
    connDom.cursorlocation=3

    commandDom.ActiveConnection =connDom
    commandDom.CommandTimeout = 60
    commandDom.CommandText = "select c.almacen "&_
            " from tpv a with(nolock) inner join cajas b with(nolock) on a.caja=b.codigo inner join tiendas c with(nolock) on b.tienda=c.codigo inner join almacenes alm with(nolock) on alm.codigo=c.almacen " & _
            " where  tpv=? and b.codigo=? and isnull(alm.fbaja,'')='' " &_
            " and a.tpv like ?+'%' and b.codigo like ?+'%' and c.codigo like ?+'%' and alm.codigo like ?+'%' "
    commandDom.CommandType = adCmdText
    commandDom.Parameters.Append commandDom.CreateParameter("@tpv",adChar,adParamInput,8,linea1)
    commandDom.Parameters.Append commandDom.CreateParameter("@cod",adChar,adParamInput,10,linea2)
    commandDom.Parameters.Append commandDom.CreateParameter("@tpvlike",adChar,adParamInput,10,session("ncliente"))
    commandDom.Parameters.Append commandDom.CreateParameter("@codlike",adChar,adParamInput,5,session("ncliente"))
    commandDom.Parameters.Append commandDom.CreateParameter("@codlike2",adChar,adParamInput,10,session("ncliente"))
    commandDom.Parameters.Append commandDom.CreateParameter("@almCod",adChar,adParamInput,10,session("ncliente"))

    set rsTPV = commandDom.Execute
    
    'cadena= "select c.almacen "&_
            '" from tpv a with(nolock) inner join cajas b with(nolock) on a.caja=b.codigo inner join tiendas c with(nolock) on b.tienda=c.codigo inner join almacenes alm with(nolock) on alm.codigo=c.almacen " & _
            '" where  tpv='" +linea1 +"' and b.codigo='" +linea2+"' and isnull(alm.fbaja,'')='' " &_
            '" and a.tpv like '"&session("ncliente") & "%' and b.codigo like '"&session("ncliente") & "%' and c.codigo like '"&session("ncliente") & "%' and alm.codigo like '"&session("ncliente") & "%' "
    'rsTPV.Open cadena,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

    if rsTPV.eof then
	    almacenTPV= ""
    else
	    almacenTPV= rsTPV("almacen")
    end if
    rsTPV.close

    si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
    si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
    si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)
    si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)
    si_tiene_modulo_fidelizacion=ModuloContratado(session("ncliente"),ModFidelizacion)
    si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
    ''dgb  10-03-2008  MODULO CENTROXOGO
    si_tiene_modulo_Centroxogo=ModuloContratado(session("ncliente"),ModCentroxogo)

    ''ricardo 3-8-2006 se hace esto para que los recursos no salgan en terra
    si_tiene_modulo_21=ModuloContratado(session("ncliente"),"21")
    si_tiene_modulo_22=ModuloContratado(session("ncliente"),"22")
    'Modulos de gestión administrativa Enviar fax
    si_tiene_modulo_29=ModuloContratado(session("ncliente"),"29")
    si_tiene_modulo_30=ModuloContratado(session("ncliente"),"30")

    si_tiene_modulo_Contabilidad=ModuloContratado(session("ncliente"),ModContabilidad)
    si_tiene_modulo_profesionales=ModuloContratado(session("ncliente"),ModProfesionales)

    SuplidosActivados=0

    set connDom = Server.CreateObject("ADODB.Connection")
    set commandDom = Server.CreateObject("ADODB.Command")

    connDom.open session("dsn_cliente")
    connDom.cursorlocation=3

    commandDom.ActiveConnection =connDom
    commandDom.CommandTimeout = 60
    commandDom.CommandText = "SELECT gestion_folios,USE_SUPLIDOS FROM configuracion with(nolock) where nempresa=?"
    commandDom.CommandType = adCmdText
    commandDom.Parameters.Append commandDom.CreateParameter("@nempresa",adChar,adParamInput,5,session("ncliente"))

    set rst2 = commandDom.Execute

    'rst2.open "SELECT gestion_folios,USE_SUPLIDOS FROM configuracion with(nolock) where nempresa='" & session("ncliente") & "'",session("dsn_cliente"), adOpenKeySet, adlockOptimistic
    if not rst2.EOF and rst2("gestion_folios") = true then
        gestionFolios = true
    else
        gestionFolios = false
    end if
    if not rst2.EOF then
        SuplidosActivados=nz_b2(rst2("USE_SUPLIDOS"))
    end if
    rst2.close

    p_nfactura=limpiaCadena(Request.QueryString("nfactura"))
    if p_nfactura="" then
	    p_nfactura=limpiaCadena(Request.QueryString("ndoc"))
    end if
    NdecDiFacturaSelect= "select ndecimales from divisas with(nolock) where codigo like ?+'%' and codigo=?"

    if p_nfactura & "" <> "" then
        'DivisaFactura=d_lookup("divisa","facturas_cli","nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente"))

        divisaFactSelect = "select divisa from facturas_cli with(nolock) where nfactura like ?+'%' and nfactura=?"

        DivisaFactura=DLookupP2(divisaFactSelect,session("ncliente")&"",adVarchar ,20 , p_nfactura&"",adVarchar ,20 ,session("dsn_cliente"))

        'NdecDiFactura=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & DivisaFactura & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))

        NdecDiFactura=DLookupP2(NdecDiFacturaSelect,session("ncliente")&"",adVarchar ,15, DivisaFactura&"",adVarchar ,15 ,session("dsn_cliente"))
    else
        NdecDiFactura=DLookupP2(NdecDiFacturaSelect,session("ncliente")&"",adVarchar ,15, request.form("h_divisa")&"",adVarchar ,15 ,session("dsn_cliente"))

        'NdecDiFactura=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & request.form("h_divisa") & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
    end if%>
    <script language="javascript" type="text/javascript">
        //ricardo 18/8/2004 , se podra cambiar la divisa al documento si no tiene detalles ni conceptos
        function cambiardivisaAnt() {
            document.facturas_cli.h_divisa.value = document.facturas_cli.divisa.value;
        }

        //*** i AMP Nueva función cambiar divisa con factor de cambio incorporado.
        var ret_tra = "";
        var ret_tra2 = "";
        function cambiardivisa(mBase) {
            document.facturas_cli.h_divisa.value = document.facturas_cli.divisa.value;

            var divisa = document.facturas_cli.divisa.value;
            if (divisa == mBase) {
                parent.pantalla.document.getElementById("tdfactcambio").style.display = "none";
                parent.pantalla.document.facturas_cli.nfactcambio.value = "1";
            }
            else {
                parent.pantalla.document.getElementById("tdfactcambio").style.display = "";
                ret_tra = ""

                if (!enProcesoFC && httpFC) {
                    var timestamp = Number(new Date());
                    var url = "../select_factcambio.asp?divisa=" + divisa;
                    httpFC.open("GET", url, false);
                    httpFC.onreadystatechange = handleHttpResponseFC;
                    enProcesoFC = false;
                    httpFC.send(null);
                }
                /*
                spfc = ret_tra.split(";"); 
                factcambio=spfc[0];  
                parent.pantalla.document.facturas_cli.nfactcambio.value=factcambio;                
                */
            }
            parent.pantalla.document.facturas_cli.h_divisa.value = divisa;
            parent.pantalla.document.facturas_cli.divisafc.value = divisa;
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
                        parent.pantalla.document.facturas_cli.nfactcambio.value = factcambio;
                        ret_tra2 = "";
                        var divisa = document.facturas_cli.divisa.value;
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
            numero = document.facturas_cli.nfactcambio.value;
            document.facturas_cli.nfactcambio.value = numero.replace(",", ".")
            numero2 = document.facturas_cli.nfactcambio.value;
            if (!/^([0-9])*[.]?[0-9]*$/.test(numero2)) {
                alert("<%=LitMsgFactCambioI%>");
                ok = 0;
            }
            if (document.facturas_cli.nfactcambio.value == "") {
                alert("<%=LitMsgFactCambioI%>");
                ok = 0;
            }
        }

        //*** f AMP
        function opendoc(documento) {
            try {
                var objFSO = new ActiveXObject("Scripting.FileSystemObject");
                var objFile = objFSO.GetFile(documento);
                nombrecorto_solo_fichero = objFile.ShortName;
                nombrecorto = objFile.ShortPath;

                var oShell = new ActiveXObject("WScript.Shell");
                oShell.Run(nombrecorto, 1, false);
            }
            catch (err) {
                if (err.message == "") menerr = "<%=LitNoExtRec%>";
                else menerr = err.message;
                alert("<%=LitOperFall%>\r\n<%=LitIncAbriFil%>: \"" + documento + "\"\r\n\r\n<%=LitAbriFilErr%>: " + menerr);
            }
        }

        function mostrar(mode) {
            if (document.getElementById('CapaVTD').innerHTML == "+") {
                document.getElementById('CapaVTD').innerHTML = "-";
                document.getElementById('showpreus').style.display = "none";
                document.getElementById('showagr').style.display = "";
                fr_DetallesIns.document.getElementById('showpreus').style.display = "none";
                fr_DetallesIns.document.getElementById('showagr').style.display = "";
            }
            else {
                document.getElementById('CapaVTD').innerHTML = "+";
                document.getElementById('showpreus').style.display = "";
                document.getElementById('showagr').style.display = "none";
                fr_DetallesIns.document.getElementById('showpreus').style.display = "";
                fr_DetallesIns.document.getElementById('showagr').style.display = "none";
            }
        }

        function VerComVenci(viene, nfactura) {
            AbrirVentana("../central.asp?pag1=ventas/vencicomerc.asp&mode=browse&ndoc=" + nfactura + "&viene=" + viene + "&titulo=<%=LitVerComVenci%>&pag2=ventas/vencicomerc_bt.asp", 'P',<%=enc.EncodeForJavascript(AltoVentana) %>,<%=enc.EncodeForJavascript(AnchoVentana) %>);
        }

        function validarCampoCarta(factura, empresa) {
            if (document.facturas_cli.cartas.value != "") {
                //ricardo 24/4/2003 se cambia todas las carta a un fichero
                AbrirVentana('generar_carta.asp?ncliente=' + document.facturas_cli.h_ncliente.value
                    <%if si_tiene_modulo_mantenimiento<>0 then%>
                        + '&ncentro=' + document.facturas_cli.ncentro.value
                        <%end if%>
                            + '&ndocumento=' + factura + '&mode=browse&ncarta=' + document.facturas_cli.cartas.value + '&empresa=' + empresa + "&tdocumento=facturas_cli", 'I',<%=enc.EncodeForJavascript(AltoVentana) %>,<%=enc.EncodeForJavascript(AnchoVentana) %>);
            }
            else alert("<%=LitNoCarta%>");
        }

        function anyadirAlbaranes(viene, ncliente, ndocumento) {
            AbrirVentana("../central.asp?pag1=ventas/pedcli_faccli_param.asp&mode=add&ncliente=" + trimCodEmpresa(ncliente) + "&viene=" + viene + "&ndoc=" + ndocumento + "&pag2=ventas/pedcli_faccli_param_bt.asp", 'P',<%=enc.EncodeForJavascript(AltoVentana) %>,<%=enc.EncodeForJavascript(AnchoVentana) %>);

        }

        function facturasVinculadas(viene, ncliente, ndocumento) {
            AbrirVentana("/ilionx45/Custom/RepsolPeru/backoffice/RepsolFacturasVinculadasNC.aspx?ndoc=" + ndocumento + "&ncliente=" + trimCodEmpresa(ncliente), 'P',<%=enc.EncodeForJavascript(AltoVentana) %>,<%=enc.EncodeForJavascript(AnchoVentana) %>);
        }

        function seguimientoCobros(viene, ncliente, ndocumento) {
            AbrirVentana("../central.asp?pag1=administracion/seguimientoCobros.asp&mode=imp&ncliente=" + trimCodEmpresa(ncliente) + "&viene=" + viene + "&ndoc=" + ndocumento + "&pag2=administracion/seguimientoCobros_bt.asp", 'P',<%=enc.EncodeForJavascript(AltoVentana) %>,<%=enc.EncodeForJavascript(AnchoVentana) %>);
        }

        function abrir_beneficios(viene, ndocumento) {
            var bloqueado;

            if (document.facturas_cli.h_cobrada.value == 0 && document.facturas_cli.h_vpagada.value == 0) bloqueado = "NO";
            else bloqueado = "SI"
            AbrirVentana("costes_doc.asp?ndoc=" + ndocumento + "&viene=" + viene + "&titulo=" + ndocumento + "&bloqueado=" + bloqueado + "&tf=" + document.facturas_cli.tfi.value, 'P',<%=enc.EncodeForJavascript(AltoVentana) %>,<%=enc.EncodeForJavascript(AnchoVentana) %>);
        }

        function abrir_fidelizacion(viene, ndocumento, puede) {
            if (puede == 1) AbrirVentana("../central.asp?pag1=fidelizacion/atrib_puntos.asp&pag2=fidelizacion/atrib_puntos_bt.asp&mode=add&ndoc=" + ndocumento + "&viene=" + viene + "&titulo=<%=LitAtribPuntosFact%> " + ndocumento, 'P',<%=enc.EncodeForJavascript(AltoVentana)%>,440);
            else alert("<%=LitMsgUsuarioPersonalNoExiste%>");
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
                    alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
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
                        alert("<%=LitFechaMal & " " & LitFechaMalCampo%> " + modo);
                        return false;
                    }
                }
            }
            return true;
        }

        //Desencadena la búsqueda del cliente cuyo numero se indica
        function TraerCliente(mode, modo) {
            cambiar_serie = "";
            cambiar_cliente = "";

            if (mode == "add") {
                if (modo == '2') {
                    if (confirm("<%=LitCambiarSeriePuedCamCli%>") == true) cambiar_cliente = 1;
                    else cambiar_cliente = 0;
                }
                else {
                    if (confirm("<%=LitCambiarCliPuedCamSer%>") == true) cambiar_serie = 1;
                    else cambiar_serie = 0;
                }
            }

            if ((cambiar_cliente == 1 && modo == '2') || modo == '1') {
                document.location.href = "facturas_cli.asp?ncliente=" + document.facturas_cli.ncliente.value + "&mode=" + mode +
                    "&nfactura=" + document.facturas_cli.h_nfactura.value +
                    "&fecha=" + document.facturas_cli.fecha.value +
                    "&cli=" + document.facturas_cli.ncliente.value +
                    "&serie=" + document.facturas_cli.serie.value +
                    "&cobrada=" + document.facturas_cli.cobrada.checked +
                    "&nenvio=" + document.facturas_cli.nenvio.value +
                    "&fechapedido=" + document.facturas_cli.fechapedido.value +
                    "&comercial=" + document.facturas_cli.comercial.value +
                <%if si_tiene_modulo_comercial<>0 then%>
                    "&agenteasignado=" + document.facturas_cli.agenteasignado.value +
                <%end if%>
                <%if si_tiene_modulo_proyectos<>0 then%>
                        "&cod_proyecto=" + document.facturas_cli.cod_proyecto.value +
		        <%end if%>
                    "&fechaenvio=" + document.facturas_cli.fechaenvio.value +
                    "&observaciones=" + document.facturas_cli.observaciones.value +
                    "&notas=" + document.facturas_cli.notas.value +
                    "&forma_pago=" + document.facturas_cli.forma_pago.value +
                    "&tipo_pago=" + document.facturas_cli.tipo_pago.value +
                    "&portes=" + document.facturas_cli.portes.value +
                    "&transportista=" + document.facturas_cli.transportista.value +
                    "&tarifa=" + document.facturas_cli.tarifa.value +
                    "&viene=" + document.facturas_cli.viene.value +
                    "&modp=" + document.facturas_cli.modp.value +
                    "&modd=" + document.facturas_cli.modd.value +
                    "&modi=" + document.facturas_cli.modi.value +
                    "&cambiar_serie=" + cambiar_serie +
                    "&cambiar_cliente=" + cambiar_cliente +
                    "&cv=" + document.facturas_cli.cv.value +
                    "&caju=" + document.facturas_cli.caju.value +
                    "&novei=" + document.facturas_cli.novei.value +
                    "&bcc=" + document.facturas_cli.bcc.value +
                    "&ocb=" + document.facturas_cli.ocb.value +
                    "&incoterms=" + document.facturas_cli.incoterms.value +
                    "&fob=" + document.facturas_cli.fob.value +
                    "&s=" + document.facturas_cli.s.value +
                    "&pciva=" + document.facturas_cli.pciva.value +
		        <% if gestionFolios then %>
                        "&nfolio=" + document.facturas_cli.nfolio.value +
		        <% end if %>
                    "&modn=" + document.facturas_cli.modn.value + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
            }
            if (mode == "add") {
                if (cambiar_cliente == 0 && modo == '2') document.facturas_cli.ncliente.value = document.facturas_cli.h_ncliente.value;
            }
        }

        //***************************************************************************
        function Editar(factura) {
            document.facturas_cli.action = "facturas_cli.asp?nfactura=" + factura + "&mode=browse" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
            document.facturas_cli.submit();
            parent.botones.document.location = "facturas_cli_bt.asp?mode=browse&nfactura=" + factura;
        }

        //***************************************************************************
        function Mas(sentido, lote, campo, texto) {
            document.facturas_cli.campo.value = campo;
            document.facturas_cli.texto.value = texto;
            document.facturas_cli.lote.value = lote;

            document.facturas_cli.action = "facturas_cli.asp?mode=search&viene=" + document.facturas_cli.viene.value +
                "&sentido=" + sentido +
                "&modp=" + document.facturas_cli.modp.value +
                "&modd=" + document.facturas_cli.modd.value +
                "&modi=" + document.facturas_cli.modi.value +
                "&novei=" + document.facturas_cli.novei.value +
                "&bcc=" + document.facturas_cli.bcc.value +
                "&cv=" + document.facturas_cli.cv.value +
                "&caju=" + document.facturas_cli.caju.value +
                "&ocb=" + document.facturas_cli.ocb.value +
                "&s=" + document.facturas_cli.s.value +
                "&pciva=" + document.facturas_cli.pciva.value +
                "&modn=" + document.facturas_cli.modn.value +
                "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
            document.facturas_cli.submit();
        }

        //***************************************************************************
        function Precios() {
            if (isNaN(document.facturas_cli.dto1.value.replace(",", ".")) || isNaN(document.facturas_cli.dto2.value.replace(",", ".")) || isNaN(document.facturas_cli.dto3.value.replace(",", ".")) || isNaN(document.facturas_cli.rf.value.replace(",", ".")))
                alert("<%=LitMsgDto1Dto2RfNumerico%>");
            else {
                if (isNaN(document.facturas_cli.irpf.value.replace(",", "."))) alert("<%=LitMsgIRPFNumerico%>");
                else {
                    //Preparamos los datos para trabajar***************************************
                    if (document.facturas_cli.dto1.value != "") dto1SinComas = document.facturas_cli.dto1.value.replace(",", ".");
                    else dto1SinComas = "0";
                    if (document.facturas_cli.dto2.value != "") dto2SinComas = document.facturas_cli.dto2.value.replace(",", ".");
                    else dto2SinComas = "0";
                    if (document.facturas_cli.dto3.value != "") dto3SinComas = document.facturas_cli.dto3.value.replace(",", ".");
                    else dto3SinComas = "0";
                    if (document.facturas_cli.rf.value != "") rfSinComas = document.facturas_cli.rf.value.replace(",", ".");
                    else rfSinComas = "0";
                    if (document.facturas_cli.irpf.value != "") irpfSinComas = document.facturas_cli.irpf.value.replace(",", ".");
                    else irpfSinComas = "0";

                    //TOTAL DESCUENTO**********************************************************
                    dto1 = (parseFloat(document.facturas_cli.importe_bruto.value.replace(",", ".")) * parseFloat(dto1SinComas)) / 100;
                    dto2 = ((parseFloat(document.facturas_cli.importe_bruto.value.replace(",", ".")) - dto1) * parseFloat(dto2SinComas)) / 100;
                    dto3 = ((parseFloat(document.facturas_cli.importe_bruto.value.replace(",", ".")) - dto1 - dto2) * parseFloat(dto3SinComas)) / 100;
                    dtoTotal = dto1 + dto2 + dto3;
                    c_dtoTotal = dtoTotal.toString();
                    document.facturas_cli.total_descuento.value = c_dtoTotal;
                    document.facturas_cli.h_total_descuento.value = document.facturas_cli.total_descuento.value;
                    //BASE IMPONIBLE***********************************************************
                    //base_imponible=parseFloat(document.facturas_cli.importe_bruto.value.replace(",","."))-parseFloat(document.facturas_cli.total_descuento.value.replace(",","."));
                    base_imponible = parseFloat(document.facturas_cli.importe_bruto.value) - parseFloat(document.facturas_cli.total_descuento.value);
                    c_base_imponible = base_imponible.toString();
                    document.facturas_cli.base_imponible.value = c_base_imponible;
                    document.facturas_cli.h_base_imponible.value = document.facturas_cli.base_imponible.value;
                    //TOTAL IVA****************************************************************
                    dto1 = ((parseFloat(document.facturas_cli.sumadet.value.replace(",", ".")) * parseFloat(dto1SinComas)) / 100);
                    dto2 = ((parseFloat(document.facturas_cli.sumadet.value.replace(",", ".")) - dto1) * parseFloat(dto2SinComas)) / 100;
                    dto3 = ((parseFloat(document.facturas_cli.sumadet.value.replace(",", ".")) - dto1 - dto2) * parseFloat(dto3SinComas)) / 100;
                    dtoTotal = dto1 + dto2 + dto3;
                    total_iva = parseFloat(document.facturas_cli.sumadet.value.replace(",", ".")) - dtoTotal;
                    c_total_iva = total_iva.toString();
                    document.facturas_cli.total_iva.value = c_total_iva;
                    document.facturas_cli.h_total_iva.value = document.facturas_cli.total_iva.value;
                    //RECARGO FINANCIERO*******************************************************
                    total_rf = (parseFloat(document.facturas_cli.base_imponible.value.replace(",", ".")) * rfSinComas) / 100;
                    c_total_rf = total_rf.toString();
                    document.facturas_cli.total_rf.value = c_total_rf;
                    document.facturas_cli.h_total_rf.value = document.facturas_cli.total_rf.value;
                    //RECARGO DE EQUIVALENCIA**************************************************
                    total_re = parseFloat(document.facturas_cli.sumaRE.value.replace(",", "."));
                    c_total_re = total_re.toString();
                    document.facturas_cli.total_re.value = c_total_re;
                    document.facturas_cli.h_total_re.value = c_total_re;
                    //SUPLIDOS
                    try {
                        total_suplidos = parseFloat(document.facturas_cli.sumaSUPLIDOS.value.replace(",", "."));
                        c_total_suplidos = total_suplidos.toString();
                        document.facturas_cli.total_suplidos.value = c_total_suplidos;
                        document.facturas_cli.h_total_suplidos.value = c_total_suplidos;
                    }
                    catch (e) {
                    }
                    //IRPF *******************************************************
                    total_irpf = (parseFloat(document.facturas_cli.base_imponible.value) * irpfSinComas) / 100;
                    c_total_irpf = total_irpf.toString();
                    document.facturas_cli.total_irpf.value = c_total_irpf;
                    document.facturas_cli.h_total_irpf.value = document.facturas_cli.total_irpf.value;
                    //TOTAL
                    total_factura = parseFloat(document.facturas_cli.base_imponible.value.replace(",", ".")) + parseFloat(document.facturas_cli.total_iva.value.replace(",", ".")) + parseFloat(document.facturas_cli.total_re.value.replace(",", ".")) + parseFloat(document.facturas_cli.total_suplidos.value.replace(",", ".")) + parseFloat(document.facturas_cli.total_rf.value.replace(",", ".")) - parseFloat(document.facturas_cli.total_irpf.value.replace(",", "."));
                    c_total_factura = total_factura.toString();
                    document.facturas_cli.total_factura.value = c_total_factura;
                    document.facturas_cli.h_total_factura.value = document.facturas_cli.total_factura.value;
                    //VOLVEMOS A DEJAR LOS DATOS CAMBIADOS COMO ESTABAN************************
                    document.facturas_cli.dto1.value = dto1SinComas;
                    document.facturas_cli.dto2.value = dto2SinComas;
                    document.facturas_cli.dto3.value = dto3SinComas;
                    document.facturas_cli.rf.value = rfSinComas;
                    document.facturas_cli.irpf.value = irpfSinComas;
                }
            }
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

            set connDom = Server.CreateObject("ADODB.Connection")
            set commandDom = Server.CreateObject("ADODB.Command")

            connDom.open session("dsn_cliente")
            connDom.cursorlocation=3

            commandDom.ActiveConnection =connDom
            commandDom.CommandTimeout = 60
            commandDom.CommandText = "select top 1 r.nremesa from remesas r with(nolock) inner join detalles_remcli dr with(nolock) on dr.nremesa=r.nremesa and (dr.nfacturavto=? ) where r.nempresa=?"
            commandDom.CommandType = adCmdText
            commandDom.Parameters.Append commandDom.CreateParameter("@nfacturavto",adVarChar,adParamInput,20,p_nfacturaREM)
            commandDom.Parameters.Append commandDom.CreateParameter("@nempresa",adVarChar,adParamInput,20,session("ncliente"))

            set rstAuxREM = commandDom.Execute


            'set rstAuxREM = Server.CreateObject("ADODB.Recordset")
            'rstAuxREM.open "select top 1 r.nremesa from remesas r with(nolock) inner join detalles_remcli dr with(nolock) on dr.nremesa=r.nremesa and (dr.nfacturavto='" & p_nfacturaREM & "' ) where r.nempresa='" & session("ncliente") & "' ", session("dsn_cliente"), adOpenKeyset, adLockOptimistic
            if not rstAuxREM.EOF then
            venConRemesa = 1
            end if
                rstAuxREM.close
                set rstAuxREM= Nothing
            end if%>
            if("<%=venConRemesa%>" == "1") {
                    alert("<%=LitMsgVencRemesa%>");
                    return;
                }
            //ricardo 11-2-2003
            //ya que cuando se insertan detalles o conceptos o pagos , no se actualiza la pagina de facturas
            //por lo que el pendiente seguia siendo cero, cuando no era verdad,por lo que no decia
            //que la factura iba a ser cobrado en su totalidad, y no ponia los vencimientos a cobrados
            pendiente = document.facturas_cli.h_impcaja.value;

            if (document.facturas_cli.impcaja.value == "") document.facturas_cli.impcaja.value = 0;
            if (isNaN(document.facturas_cli.impcaja.value.replace(",", "."))) {
                alert("<%=LitMsgImporteNumerico%>");
                return false;
            }
            else {
                if (document.facturas_cli.i_pago.value == "") {
                    alert("<%=LitMsgTipoPagoNoNulo%>");
                    return false;
                }
                else {
                    if (parseFloat(document.facturas_cli.impcaja.value.replace(",", ".")) == 0) {
                        alert("<%=LitErrImportePago%>");
                        return false;
                    }
                }
            }
            pagada = "NO";
            if (parseFloat(document.facturas_cli.impcaja.value.replace(",", ".")) == parseFloat(pendiente.replace(",", "."))) {
                if (!confirm("<%=LitMsgAnotCobradaFact%>")) return false;
                else pagada = "SI";
            }
            if (document.facturas_cli.ncaja.value == "") {
                alert("<%=LitMsgCajaNoNulo%>");
                return false;
            }
            else {
                fr_PagosCuenta.document.facturas_clipago.action = "facturas_clipago.asp?mode=acaja&ndoc=" + nfactura + "&impcaja=" + document.facturas_cli.impcaja.value + "&i_pago=" + document.facturas_cli.i_pago.value + "&ncaja=" + document.facturas_cli.ncaja.value + "&pagada=" + pagada;
                fr_PagosCuenta.document.facturas_clipago.submit();
                setTabsSelected(3);
                //if (document.getElementById("PAGOS_CUENTA").style.display == "none") 
                //    tier1Menu(PAGOS_CUENTA,document.getElementById("img5"),'<%=oculta%>');
                //}
            }
        }

        //Comprueba si el importe del pago es numerico
        function importepagoComp() {
            if (isNaN(document.facturas_cli.importePago.value.replace(",", "."))) {
                alert("<%=LitErrImportePago2%>");
                return;
            }
        }

        //Calcula el importe de la línea de detalle del concepto.
        function ImporteDetalle() {
            if (parseFloat(document.facturas_cli.pvp.value) < 0) {
                alert("<%=LitMsgPvPNoNegativo%>");
                document.facturas_cli.pvp.value = 0;
            }
            if (isNaN(document.facturas_cli.cantidad.value.replace(",", ".")) || isNaN(document.facturas_cli.descuento.value.replace(",", ".")) || isNaN(document.facturas_cli.descuento2.value.replace(",", ".")) || isNaN(document.facturas_cli.descuento3.value.replace(",", ".")) || isNaN(document.facturas_cli.pvp.value.replace(",", ".")))
                alert("<%=LitMsgCanPreDesNumerico%>");
            else {
                if (document.facturas_cli.pvp.value == "") document.facturas_cli.pvp.value = 0;
                if (document.facturas_cli.cantidad.value == "") document.facturas_cli.cantidad.value = 1;
                if (document.facturas_cli.descuento.value == "") document.facturas_cli.descuento.value = 0;
                if (document.facturas_cli.descuento2.value == "") document.facturas_cli.descuento2.value = 0;
                if (document.facturas_cli.descuento3.value == "") document.facturas_cli.descuento3.value = 0;
                pvpSinComas = document.facturas_cli.pvp.value.replace(",", ".");
                cantidadSinComas = document.facturas_cli.cantidad.value.replace(",", ".");
                dtoSinComas = document.facturas_cli.descuento.value.replace(",", ".");
                dto2SinComas = document.facturas_cli.descuento2.value.replace(",", ".");
                dto3SinComas = document.facturas_cli.descuento3.value.replace(",", ".");
                pelas = parseFloat(cantidadSinComas) * parseFloat(pvpSinComas);
                pelas_descuento = (pelas * parseFloat(dtoSinComas)) / 100;
                pelas = pelas - pelas_descuento;
                pelas_descuento = (pelas * parseFloat(dto2SinComas)) / 100;
                pelas = pelas - pelas_descuento;
                pelas_descuento = (pelas * parseFloat(dto3SinComas)) / 100;
                importe = pelas - pelas_descuento;
                c_importe = importe.toString();
                document.facturas_cli.cantidad.value = cantidadSinComas;
                document.facturas_cli.descuento.value = dtoSinComas;
                document.facturas_cli.descuento2.value = dto2SinComas;
                document.facturas_cli.descuento3.value = dto3SinComas;
                document.facturas_cli.importe.value = parseFloat(c_importe).toFixed(<%=enc.EncodeForJavascript(NdecDiFactura) %>);
                document.facturas_cli.pvp.value = pvpSinComas;
            }
        }

        //Añade un concepto a la factura
        function addConcepto(nfactura) {
            document.facturas_cli.pvp.value = document.facturas_cli.pvp.value.replace(".", ",");
            document.facturas_cli.descuento.value = document.facturas_cli.descuento.value.replace(".", ",");
            document.facturas_cli.descuento2.value = document.facturas_cli.descuento2.value.replace(".", ",");
            document.facturas_cli.descuento3.value = document.facturas_cli.descuento3.value.replace(".", ",");

            if (isNaN(document.facturas_cli.cantidad.value.replace(",", ".")) || isNaN(document.facturas_cli.descuento.value.replace(",", ".")) || isNaN(document.facturas_cli.descuento2.value.replace(",", ".")) || isNaN(document.facturas_cli.descuento3.value.replace(",", ".")) || isNaN(document.facturas_cli.pvp.value.replace(",", "."))) {
                alert("<%=LitMsgCanPreDesNumerico%>");
                return;
            }

            if (document.facturas_cli.descripcion.value == "") {
                alert("<%=LitMsgDesVacia%>");
                return;
            }
            if (isNaN(document.facturas_cli.pvp.value.replace(",", "."))) {
                alert("<%=LitMsgImporteNumerico%>");
                return;
            }

            if (document.facturas_cli.gestbono.value == 1 && document.facturas_cli.frabono.value == 1 && document.facturas_cli.ndetcon.value > 0) {
                alert("<%=LitMsgFraBonoCon%>");
                return;
            }

            //ricardo 23/4/2004 comprobamos el riesgo
            preguntar_riesgo_conf = document.facturas_cli.prieconf.value;
            contrasenya_riesgo_conf = document.facturas_cli.contrpregries.value;
            texto_aviso1 = "<%=LitHaSupRiegMaxAut%>";
            texto_aviso2 = "<%=LitClRiesgo%> : " + document.facturas_cli.rsocries.value;
            texto_aviso3 = "<%=LitRMaxAut%> : ";
            texto_aviso4 = "<%=LitRAlc%> : ";
            if (comprobarRiesgo("", "facturas_cli", "facturas_cli", preguntar_riesgo_conf, texto_aviso1, texto_aviso2, texto_aviso3, texto_aviso4, "<%=LitHaSupRiegMaxAutConv3%>", contrasenya_riesgo_conf, "<%=LitPregContrRiesgo%>", "FACTURA A CLIENTE", nfactura, 0, "CONCEPTO NUEVO") == false)
                return;

            //Asignar los valores a los campos del submarco de detalles
            fr_Conceptos.document.facturas_clicon.cantidad.value = document.facturas_cli.cantidad.value;
            fr_Conceptos.document.facturas_clicon.descripcion.value = document.facturas_cli.descripcion.value;
            fr_Conceptos.document.facturas_clicon.pvp.value = document.facturas_cli.pvp.value;
            fr_Conceptos.document.facturas_clicon.descuento.value = document.facturas_cli.descuento.value;
            fr_Conceptos.document.facturas_clicon.descuento2.value = document.facturas_cli.descuento2.value;
            fr_Conceptos.document.facturas_clicon.descuento3.value = document.facturas_cli.descuento3.value;
            fr_Conceptos.document.facturas_clicon.iva.value = document.facturas_cli.iva.value;
            //Recargar el submarco de detalles
            fr_Conceptos.document.facturas_clicon.action = "facturas_clicon.asp?mode=first_save";
            fr_Conceptos.document.facturas_clicon.submit();

            //Limpiar los campos del formulario
            document.facturas_cli.cantidad.value = "1";
            document.facturas_cli.descripcion.value = "";
            document.facturas_cli.pvp.value = "0";
            document.facturas_cli.descuento.value = "0";
            document.facturas_cli.descuento2.value = "0";
            document.facturas_cli.descuento3.value = "0";
            document.facturas_cli.iva.value = document.facturas_cli.defaultIva.value;
            document.facturas_cli.importe.value = "0";
            //Colocar el foco en el campo de cantidad.
            document.facturas_cli.cantidad.focus();
            document.facturas_cli.cantidad.select();
        }

        function addSuplido() {
            if (document.facturas_cli.descripcionSup.value == "") {
                alert("<%=LitMsgDesVacia%>");
                return;
            }
            if (isNaN(document.facturas_cli.importeSup.value.replace(",", "."))) {
                alert("<%=LitMsgImporteNumerico%>");
                return;
            }
            if (parseFloat(document.facturas_cli.importeSup.value.replace(",", ".")) < 0) {
                alert("<%=LitMsgImporteNoNegativo%>");
                document.facturas_cli.importeSup.value = 0;
            }
            if (parseFloat(document.facturas_cli.importeSup.value.replace(",", ".")) == 0) {
                alert("<%=LITMSGIMPORTEPOSITIVO%>");
            }

            //Asignar los valores a los campos del submarco de suplidos
            fr_Suplidos.document.facturas_suplidos.descripcion.value = document.facturas_cli.descripcionSup.value;
            fr_Suplidos.document.facturas_suplidos.importe.value = document.facturas_cli.importeSup.value;
            //Recargar el submarco de detalles
            fr_Suplidos.document.facturas_suplidos.action = "facturas_suplidos.asp?mode=first_save";
            fr_Suplidos.document.facturas_suplidos.submit();

            //Limpiar los campos del formulario
            document.facturas_cli.descripcionSup.value = "";
            document.facturas_cli.importeSup.value = "0";
            //Colocar el foco en el campo de cantidad.
            document.facturas_cli.descripcionSup.focus();
            document.facturas_cli.descripcionSup.select();
        }

        //Añade un pago a cuenta.
        function addPago(nfactura) {
            if (document.facturas_cli.importePago.value == "") document.facturas_cli.importePago.value = 0;

            if (document.facturas_cli.fechaPago.value == "") {
                alert("<%=LitErrFechaPago%>");
                return;
            }

            if (!cambiarfecha(document.facturas_cli.fechaPago.value, "Fecha Pago")) return;

            if (!checkdate(document.facturas_cli.fechaPago)) {
                alert("<%=LitMsgFechaFecha%>");
                return;
            }

            if (isNaN(document.facturas_cli.importePago.value.replace(",", "."))) {
                alert("<%=LitErrImportePago2%>");
                return;
            }
            else {
                if (parseFloat(document.facturas_cli.importePago.value.replace(",", ".")) == 0) {
                    alert("<%=LitMsgImportePositivo%>");
                    return;
                }
            }
            if (document.facturas_cli.descripcionPago.value == "") {
                alert("<%=LitMsgDesVacia%>");
                return;
            }
            if (document.facturas_cli.tipoPago.value == "") {
                alert("<%=LitMsgTipoPagoNoNulo%>");
                return;
            }
            //Asignar los valores a los campos del submarco de detalles
            fr_PagosCuenta.document.facturas_clipago.fecha.value = document.facturas_cli.fechaPago.value;
            fr_PagosCuenta.document.facturas_clipago.importe.value = document.facturas_cli.importePago.value;
            fr_PagosCuenta.document.facturas_clipago.descripcion.value = document.facturas_cli.descripcionPago.value;
            fr_PagosCuenta.document.facturas_clipago.medio.value = document.facturas_cli.tipoPago.value;
            //Recargar el submarco de pagos a cuenta
            fr_PagosCuenta.document.facturas_clipago.action = "facturas_clipago.asp?mode=first_save";
            fr_PagosCuenta.document.facturas_clipago.submit();
            //Limpiar los campos del formulario
            var hoy = new Date();
            document.facturas_cli.fechaPago.value = hoy.getDate() + "/" + (hoy.getMonth() + 1) + "/" + hoy.getFullYear();
            document.facturas_cli.importePago.value = "0";
            document.facturas_cli.descripcionPago.value = "";
            document.facturas_cli.tipoPago.value = "";
            //Colocar el foco en el campo de cantidad.
            document.facturas_cli.fechaPago.focus();
            document.facturas_cli.fechaPago.select();
        }

        //Añade un pago a cuenta.
        function addVencimiento(nfactura) {
            if (document.facturas_cli.importeVto.value == "") document.facturas_cli.importeVto.value = 0;

            if (document.facturas_cli.fechaVto.value == "") {

                alert("<%=LitErrFechaPago%>");
                return;
            }

            if (!cambiarfecha(document.facturas_cli.fechaVto.value, "Fecha Vencimiento")) return;

            if (!checkdate(document.facturas_cli.fechaVto)) {

                alert("<%=LitMsgFechaFecha%>");
                return;
            }

            if (isNaN(document.facturas_cli.importeVto.value.replace(",", "."))) {

                alert("<%=LitErrImportePago%>");
                return;
            }
            else {
                if (parseFloat(document.facturas_cli.importeVto.value.replace(",", ".")) == 0) {
                    alert("<%=LitMsgImportePositivo%>");
                    return;
                }
            }

            if ((parseFloat(document.facturas_cli.importeVto.value.replace(",", ".")) == parseFloat(document.facturas_cli.recibidoVto.value.replace(",", "."))) || (document.facturas_cli.cobradoVto.checked)) {

                if (!confirm("<%=LitMsgCobVtoSinCajaConfirm%>")) return;
            }

            //Asignar los valores a los campos del submarco de detalles
            fr_Vencimientos.document.facturas_cliven.fecha.value = document.facturas_cli.fechaVto.value;
            fr_Vencimientos.document.facturas_cliven.importe.value = document.facturas_cli.importeVto.value;
            fr_Vencimientos.document.facturas_cliven.importecob.value = document.facturas_cli.recibidoVto.value;
            fr_Vencimientos.document.facturas_cliven.cobradotodo.checked = document.facturas_cli.cobradoVto.checked;
            fr_Vencimientos.document.facturas_cliven.observaciones.value = document.facturas_cli.obsVto.value;
            //Recargar el submarco de pagos a cuenta
            fr_Vencimientos.document.facturas_cliven.action = "facturas_cliven.asp?mode=first_save";
            fr_Vencimientos.document.facturas_cliven.submit();
            //Limpiar los campos del formulario
            var hoy = new Date();
            document.facturas_cli.fechaVto.value = hoy.getDate() + "/" + (hoy.getMonth() + 1) + "/" + hoy.getFullYear();
            document.facturas_cli.importeVto.value = "0";
            document.facturas_cli.recibidoVto.value = "0";
            document.facturas_cli.cobradoVto.checked = false;
            //Colocar el foco en el campo de cantidad.
            document.facturas_cli.fechaVto.focus();
            document.facturas_cli.fechaVto.select();
        }

        //Genera los vencimientos de la factura.
        function genVencimiento(nfactura) {
            fr_Vencimientos.document.facturas_cliven.action = "facturas_cliven.asp?mode=browse&gen=SI";
            fr_Vencimientos.document.facturas_cliven.submit();
        }

        /*cag*/
        function desbloqueoFactura(factura) {
            if (confirm("<%=LitMsgDesBloqueo%>")) {
                document.facturas_cli.h_ahora.value = 0;
                document.facturas_cli.action = "facturas_cli.asp?nfactura=" + factura + "&mode=desbloqueo&ahora=0" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>&viene=browse";
                document.facturas_cli.submit();
                parent.botones.document.location = "facturas_cli_bt.asp?mode=browse";
            }
        }

        function bloqueoFactura(factura) {
            if (confirm("<%=LitMsgBloqueo%>")) {
                document.facturas_cli.h_ahora.value = 1;
                document.facturas_cli.action = "facturas_cli.asp?nfactura=" + factura + "&mode=bloqueo&ahora=1" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>&viene=browse";
                document.facturas_cli.submit();
                parent.botones.document.location = "facturas_cli_bt.asp?mode=browse";
            }
        }

        function desbloqueoFacturaSearch(factura) {
            if (confirm("<%=LitMsgDesBloqueo%>")) {
                url = "facturas_cli.asp?nfactura=" + factura + "&mode=desbloqueo&ahora=0" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>&viene=search";
                //alert(url);
                location.replace(url);
                parent.botones.document.location = "facturas_cli_bt.asp?mode=search";
            }
        }

        function bloqueoFacturaSearch(factura) {
            if (confirm("<%=LitMsgBloqueo%>")) {
                url = "facturas_cli.asp?nfactura=" + factura + "&mode=bloqueo&ahora=1" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>&viene=search";
                location.replace(url);
                parent.botones.document.location = "facturas_cli_bt.asp?mode=search";
            }
        }


        var numFolioCorrecto = false;

        <% if gestionFolios then %>
            function ComprobarNumFolio() {
                if (IsNumeric(parent.pantalla.document.facturas_cli.nfolio.value)) {
                    nserie = parent.pantalla.document.facturas_cli.serie.value;
                    if (!enProceso && http) {
                        var timestamp = Number(new Date());
                        var url = "Facturas_cli_bt.asp?mode=consultaAJAX&consulta=nfolioMinMax&nserie=" + nserie + "&ts=" + timestamp;
                        http.open("GET", url, false);
                        http.onreadystatechange = handleHttpResponse;
                        enProceso = true;
                        http.send(null);
                    }
                }
                else {
                    alert("<%=LITMSGFOLIONUMERICO%>");
                    numFolioCorrecto = false;
                    parent.pantalla.document.facturas_cli.nfolio.focus();
                }
            }
            <% end if %>

                function handleHttpResponse() {
                    if (http.readyState == 4) {
                        if (http.status == 200) {
                            if (http.responseText.indexOf('invalid') == -1) {
                                // Armamos un array, usando la coma para separar elementos
                                results = http.responseText;
                                enProceso = false;
                                if (results == "" || results == "ERROR") alert("<%=LitErrorNumFolio%>");
                                else {
                                    var retValue = results.split(",");
                                    nfolioIntroducido = parent.pantalla.document.facturas_cli.nfolio.value;
                                    if (nfolioIntroducido < parseInt(retValue[0]) || nfolioIntroducido > parseInt(retValue[1])) {
                                        alert("<%=LITMSGFOLIOINCORRECTO %>" + "(" + retValue[0] + "-" + retValue[1] + ")");
                                        parent.pantalla.document.facturas_cli.nfolio.focus();
                                        numFolioCorrecto = false;
                                    }
                                    else numFolioCorrecto = true;
                                }
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
                    catch (e) { xmlhttp = false; }
                }
                return xmlhttp;
            }

        var enProceso = false; // lo usamos para ver si hay un proceso activo
        var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest


        function inicio() {
        <% 'cag
            ' Ocultar detalle de las facturas
            Resultado = 0
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.open dsnilion
            command.ActiveConnection = conn
            command.CommandTimeout = 0
            command.CommandText = "compruebaOcultarDetalle"
            command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
            command.Parameters.Append command.CreateParameter("@usr", adVarChar, adParamInput, 50, session("usuario"))
            command.Parameters.Append command.CreateParameter("@emp", adVarChar, adParamInput, 10, session("ncliente"))
            Command.Parameters.Append Command.CreateParameter("@resul", adInteger, adParamOutput, Resultado)
            'on error resume next
            command.Execute,,adExecuteNoRecords
            oculta = Command.Parameters("@resul").Value
            conn.close
            set command= nothing
            set conn= nothing
            ' si no se ocultan los detalles, debe añadir linea pedido de gtos envio si procede
            'Si se ocultan los detalles, no se debe añadir tal linea

            if oculta= 0 then
            'i(EJM 04/09/2006)  Si la factura se acaba de crear, añadir una línea de factura con los gtos de envió si procediese
            if Request.QueryString("mode") = "first_save" and Nulear(request.form("portes")) = LitDebidos then

            set connDom = Server.CreateObject("ADODB.Connection")
            set commandDom = Server.CreateObject("ADODB.Command")

            connDom.open session("dsn_cliente")
            connDom.cursorlocation=3

            commandDom.ActiveConnection =connDom
            commandDom.CommandTimeout = 60
            commandDom.CommandText = "exec Insert_GtoEnvioLinPed '" & session("ncliente") & request.form("ncliente") & "', '" & session("f_caja") & "'"
            commandDom.CommandType = adCmdText
            commandDom.Parameters.Append commandDom.CreateParameter("@nempresa",adVarChar,adParamInput,10,session("ncliente")&request.form("ncliente"))
            commandDom.Parameters.Append commandDom.CreateParameter("@nempresa",adVarChar,adParamInput,10,session("f_caja"))

            set rstCrearLineasFacZonas = commandDom.Execute


            'strSelect = "exec Insert_GtoEnvioLinPed '" & session("ncliente") & request.form("ncliente") & "', '" & session("f_caja") & "'"
            'rstCrearLineasFacZonas.open strSelect, session("dsn_cliente"), adOpenKeyset, adLockOptimistic


            if not rstCrearLineasFacZonas.eof then
            datosCrearLineasFacZona = rstCrearLineasFacZonas.Getrows
            totalCrearLineasFacZona = 1
            // DGM 17/01/11 Añadido el parametro "viene=auto" para controlar si la linea es insertada automáticamente
            // por los Gastos de Envío.
            response.write "document.crearLineaFac.action='facturas_clidet.asp?mode=first_save&viene=auto';"
            response.write "document.crearLineaFac.submit();"
			        else
            totalCrearLineasFacZona = 0
            end if
			        set rstCrearLineasFacZonas= Nothing
            'fin(EJM 04/09/2006)
            end if
	        'cag
	        end if%>
        }

        //****************************************************************************
        function MasDet(sentido, lote, firstReg, lastReg, campo, criterio, texto, firstRegAll, lastRegAll) {
            fr_Detalles.document.facturas_clidet.action = "facturas_clidet.asp?mode=browse&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&firstReg=" + firstReg + "&lastReg=" + lastReg + "&firstRegAll=" + firstRegAll + "&lastRegAll=" + lastRegAll + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
            fr_Detalles.document.facturas_clidet.submit();
        }

        // GPD (26/04/2007).
        function GenerarValeDescuento(nfactura) {
            if (document.getElementById('descuento').innerHTML == '') {
                if (confirm('<%= LitMsgSeguroValeDto %>')) document.getElementById("frDescuento").document.location = 'valedto.asp?mode=new&ndoc=' + nfactura;
            }
            else AbrirVentana('../central.asp?pag1=ventas/valedto.asp&pag2=ventas/valedto_bt.asp&mode=browse&ndoc=' + nfactura, 'P', <%=enc.EncodeForJavascript(AltoVentana) %>, <%=enc.EncodeForJavascript(AnchoVentana) %>)
        }

        // GPD (27/04/2007).
        function BuscarValeDescuento(nfactura) {
            AbrirVentana('valedto.asp?mode=search&ndoc=' + nfactura, 'P',<%=enc.EncodeForJavascript(AltoVentana) %>,<%=enc.EncodeForJavascript(AnchoVentana) %>);
        }

        //dgb 18/06/2010   asociado a Factura electronica, desde .Net refresca el IFRAME
        function Refrescar(factura) {
            fr_LFactura.document.location.href = "../../<%=CarpetaProduccion %>/ventas/linkFacturaE.asp?cod=" + factura + "&l=<%=LitLinkFirma %>&l1=<%=LitLinkFirma1%>&l2=<%=LitLinkFirma2 %>&l3=<%=LitLinkFirma3 %>"
        }

        function Redimensionar() {
            var alto = jQuery(window).height();
            var diference = 500;
            var dir_default = 145;

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

        function ColocarCapa() {
            <%''MPC 05/ 10 / 2010 Se ha modificado para que en función del navegador que se utilice se ponga un estilo u otro
            if InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Chrome") <> 0 then%>
                document.getElementById("input_file").style.width="130px";
 	        <%elseif InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Safari") <> 0 then%>
                document.getElementById("input_file").style.width="135px";
	        <%elseif InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Opera") <> 0 then%>
                document.getElementById("input_file").style.width="100px";
	        <%elseif InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Firefox") <> 0 then%>
	            //document.getElementById("input_doc").style.position="absolute";
	            //document.getElementById("input_doc").style.width="150px";
	        <%else%>
                document.getElementById("input_file").style.width="0px";
	        <%end if
	        ''FIN MPC 05/ 10 / 2010 %>
        }

        function RoundNumValue(obj, dec) {
            obj.value = obj.value.replace(',', '.');
            var valor = parseFloat(obj.value);
            if (valor != 0) obj.value = valor.toFixed(dec);
        }


        //*** AMP
        function CloseNonPayment(nfactura, nven, datenote) {
            esconde();

            if (nven == "") {
                document.facturas_cli.elements["cobrada"].checked = false;
                ndoc = nfactura
                alert("<%=LITMSGCHANGESTATE%>");
            }
            else {
                ndoc = nfactura + "-" + nven;
                alert("<%=LITMSGCHANGESTATE2%>");
            }
            if (window.confirm("<%=LITMSGNPINVOICE%>") == true) {

                document.facturas_cli.action = "facturas_cli.asp?mode=add&vienenp=" + ndoc + "&datent=" + datenote;
                document.facturas_cli.submit();

                parent.botones.document.opciones.action = "facturas_cli_bt.asp?mode=add";
                parent.botones.document.opciones.submit();
            }
            else {
                document.facturas_cli.action = "facturas_cli.asp?mode=browse&nfactura=" + nfactura;
                document.facturas_cli.submit();

                parent.botones.document.opciones.action = "facturas_cli_bt.asp?mode=browse";
                parent.botones.document.opciones.submit();
            }
        }

        function SetImpago(_referencia, _ndoc, _cust, _typedoc) {
            //alert("padre..>"+_referencia+"--"+ _ndoc +"--"+_cust+"--"+_typedoc) ;
            reloadClass(_referencia, "../central.asp?pag1=administracion/nonpayment.asp&mode=add&ndoc=" + _ndoc + "&ncliente=" + _cust + "&tdocumento=" + _typedoc + "&pag2=administracion/nonpayment_bt.asp");
            alPresionar(_referencia);
            //setTimeout(function(){document.frames("frame-fr_NonPayment").document.nonpayment.i_text.focus();},1500);
            //     getElementById("frame-fr_NonPayment").i_text.focus();
            //fr_NonPayment.document.nonpayment.i_text.focus();
            //fr_NonPayment.document.getElementById("i_text").focus();
        }

        function ChangeNonPayment() {
            if (document.facturas_cli.elements["cobrada"].checked == true) parent.pantalla.document.getElementById("NonPaymentEnabled").style.display = "";
            else parent.pantalla.document.getElementById("NonPaymentEnabled").style.display = "none";
        }

        jQuery(window).resize(function () { Redimensionar(); });
    </script>
    <%'Everilion Interface Timing%>
<script language="javascript" type="text/javascript" src="/lib/js/InterfaceLoadTime.js"></script>
<script language="javascript" type="text/javascript">

        window.onload = function () {
            self.status = '';
            inicio();
    <%if tracetime> 0 then %>
                StoreTiming("<%=CarpetaProduccion%>", <%=enc.EncodeForJavascript(tracetime) %>, "<%=enc.EncodeForJavascript(Request.QueryString("mode"))&""%>", "<%=enc.EncodeForJavascript(session("usuario"))%>", "<%=enc.EncodeForJavascript(session("ncliente"))%>", window.location.pathname);
    <%end if %>
 }

</script>
    <%
    modoPantalla=Request.QueryString("mode")
    if modoPantalla & ""="" then
        modoPantalla=request.Form("mode")
    end if
    'CuandoRedimensionar=0
    'if modoPantalla="browse" or modoPantalla="save" or modoPantalla="first_save" then
    '    CuandoRedimensionar=1
    'end if
    %>

    <body class="bodycentral">
    <%'******************************************************************************
    function ExisteFolioMejico(oculta)
        '**ASP 31/01/2011
            
        set conn=server.CreateObject("ADODB.Connection")
        set command=server.CreateObject("ADODB.Command")
        conn.open session("dsn_cliente")
        conn.cursorlocation=3
        command.activeConnection=conn
        command.CommandType = adCmdStoredProc
        command.CommandText= "Emp_MX"
        command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,,5, session("ncliente"))
        command.Parameters.Append command.CreateParameter("@p_dev",adVarChar,adParamOutput,4)
        command.execute
        if command("@p_dev")=1 then
            mejico="1"
        end if
        conn.close
        set command=nothing
        set conn=nothing

	    if mejico = "1" then
	        oculta=1
	        set conn = Server.CreateObject("ADODB.Connection")
            conn.cursorlocation=3
            conn.open DSNCronos
            set rstFactura= conn.execute("EXEC existMxCFD "& _
	        "@ncliente='"&session("ncliente")&"',"&_
	        "@nfactura='"&rst("nfactura")&"'")
	        existeFirma=rstFactura(0)&""
	        codigoFirma=rstFactura(1)&""
            if existeFirma = 1 then
                oculta=0
            end if
	    end if
			
        '**ASP 
    end function
    '*******************************************************************************

    '*** i AMP 25/07/2011 Gestión de impagos. Función para generar cabecera de factura vinculada.
    sub GenerarVincFacturaImpago(ndoc,nfactura,datenote)
        company = session("ncliente")     
	    set connNPinv=server.CreateObject("ADODB.Connection")
        set commandNPinv=server.CreateObject("ADODB.Command")
        connNPinv.open session("dsn_cliente")
        connNPinv.cursorlocation=3
        commandNPinv.activeConnection=connNPinv
        commandNPinv.CommandType = adCmdStoredProc
        commandNPinv.CommandText= "makeHeadInvoiceToInv"
        commandNPinv.Parameters.Append commandNPinv.CreateParameter("@ncompany",adVarChar,,5, session("ncliente"))
        commandNPinv.Parameters.Append commandNPinv.CreateParameter("@ndoc",adVarChar,,50, ndoc)   
        commandNPinv.Parameters.Append commandNPinv.CreateParameter("@ninvoice",adVarChar,,20, nfactura)  
        commandNPinv.Parameters.Append commandNPinv.CreateParameter("@datenote",adVarChar,,20, datenote)   
 
        commandNPinv.execute   
        connNPinv.close
        set commandNPinv=nothing
        set connNPinv=nothing
    end sub
    '*** f AMP

    '******************************************************************************
    sub ModificarVencimientos(comercial_ant,old_forma_pago,cambiarcom)
	    'Ahora cambiamos el comercial de los vencimientos si hace falta es decir, solo se cambiara si cambia la forma de pago o el comercial
	    if old_forma_pago<>rst("forma_pago") or comercial_ant<>rst("comercial") or (isnull(old_forma_pago) and rst("forma_pago") & "">"") or (old_forma_pago & "">"" and isnull(rst("forma_pago"))) or (isnull(comercial_ant) and rst("comercial") & "">"") or (comercial_ant & "">"" and isnull(rst("comercial"))) then
		    if cambiarcom=1 then

                set connDom = Server.CreateObject("ADODB.Connection")
                set commandDom = Server.CreateObject("ADODB.Command")

                connDom.open session("dsn_cliente")
                connDom.cursorlocation=3

                commandDom.ActiveConnection =connDom
                commandDom.CommandTimeout = 60
            
			    if rst("comercial") & ""="" then
				    commandDom.CommandText ="update vencimientos_salida with(updlock) set comercial=NULL where nfactura=?"
                    commandDom.CommandType = adCmdText
                    commandDom.Parameters.Append commandDom.CreateParameter("@nfactura",adVarChar,adParamInput,20,rst("nfactura"))
			    else
				    commandDom.CommandText ="update vencimientos_salida with(updlock) set comercial=? where nfactura=?"
                    commandDom.CommandType = adCmdText
                    commandDom.Parameters.Append commandDom.CreateParameter("@comercial",adVarChar,adParamInput,20,rst("comercial"))
                    commandDom.Parameters.Append commandDom.CreateParameter("@nfactura",adVarChar,adParamInput,20,rst("nfactura"))
			    end if

                set rstAux = commandDom.Execute

			    'rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    if rstAux.state<>0 then rstAux.close
		    end if
	    end if
    end sub

    'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
    sub GuardarRegistro(nfactura,nserie,fecha,mode)
	    ModDocumentoEquip=false
	    if nfactura="" then
		    'Crear un nuevo registro.
		    rst.AddNew
		    '******************** Manejo de domicilios
		    Dom=Domicilios("VENTAS","FAC_ENV_CLI",session("ncliente") & request.form("ncliente"),rst)
		    ModDocumento=true
	    else
		    mensajeTratEquipos="OK"
		    if ((rst("ncliente")&""<>session("ncliente") & request.form("ncliente")&"") or (rst("ncentro")&""<>request.form("ncentro")&"") or (rst("fecha")<>cdate(Request.Form("fecha")&""))) then
			    ModDocumentoEquip=true
		    end if
		    if mid(mensajeTratEquipos,1,2)<>"OK" then
			    ModDocumento=false%>
			    <script language="javascript" type="text/javascript">
                    alert("<%=mensajeTratEquipos%>");
			    </script>
		    <%else
			    ModDocumento=true
		    end if
	    end if
	    if ModDocumento then
		    'Asignar los nuevos valores a los campos del recordset.
		    rst("serie")=Nulear(Request.Form("serie"))
		    'Si se le cambia el cliente a la factura, hay que capturar sus direcciones
		    if rst("ncliente")&""<>(session("ncliente") & request.form("ncliente")&"") and nfactura<>"" then
			    Dom=Domicilios("VENTAS","FAC_ENV_CLI",session("ncliente") & request.form("ncliente"),rst)
		    end if
		    '--------------------------------------------------------------------------
		    rst("ncliente")=session("ncliente") & Nulear(request.form("ncliente"))

		    if rst("ncentro")&""<>request.form("ncentro")&"" then
			    if request.form("ncentro")>"" then
				    'DirEnvioCentro=d_lookup("dir_envio","centros","ncentro='" & request.form("ncentro") & "'",session("dsn_cliente"))
                    DirEnvioCentroSelect = "select dir_envio from centros with(nolock) where ncentro=?"
                    DirEnvioCentro= DLookupP1(DirEnvioCentroSelect, request.form("ncentro")&"" ,adVarchar ,10, session("dsn_cliente"))

				    if DirEnvioCentro&"">"" then
					    rst("dir_envio")=Nulear(DirEnvioCentro)
				    else
					    Dom=Domicilios("VENTAS","FAC_ENV_CLI",rst("ncliente"),rst)
				    end if
			    else
				    Dom=Domicilios("VENTAS","FAC_ENV_CLI",rst("ncliente"),rst)
			    end if
		    end if
		    rst("ncentro")=Nulear(request.form("ncentro"))
            se_han_cambiado_los_descuentos=0

            if Clng(replace(Null_z(request.form("dto1")),".",""))<>Clng(null_z(rst("descuento"))) or _
	            Clng(replace(Null_z(request.form("dto2")),".",""))<>Clng(null_z(rst("descuento2"))) or _
		            Clng(replace(Null_z(request.form("dto3")),".",""))<>Clng(null_z(rst("descuento3"))) then
			            se_han_cambiado_los_descuentos=1
            end if
            'ndec=d_lookup("ndecimales","divisas", "codigo like '"&session("ncliente")&"%' and codigo='"&request.form("h_divisa")&"'", session("dsn_cliente"))
            ndecSelect="select ndecimales from divisas with(nolock) where codigo like ?+'%' and codigo = ?"
            ndec=DLookupP2(ndecSelect, session("ncliente")&"",adVarchar,15, request.form("h_divisa")&"", adVarchar, 15, session("dsn_cliente"))
            ''MPC 04/02/2010 Se modifica el guardado de los descuentos de la cabecera para que se pueda guardar con decimales
		    rst("descuento")=miround(Null_z(request.form("dto1")), decpor)
		    rst("descuento2")=miround(Null_z(request.form("dto2")),decpor)
		    rst("descuento3")=miround(Null_z(request.form("dto3")),decpor)
		    ''FIM MPC 04/02/2010
		    rst("importe_bruto")=miround(Null_z(request.form("h_importe_bruto")),ndec)
            ''ricardo 20-5-2003
            ''redondeamos, ya que cuando viene un numero con muchos decimales , al reemplazar
            ''sale un numero mas grande que el tipo de datos real soporta
            ''ademas, como luego se vuelve a calcular con la funcion precios
            ''no pasa nada si desperdiciamos decimales,es simplemente para que no de error
            ''el tipo soporta(en la definicion de tabla), no mas de un numero de 8 digitos
		    rst("total_descuento")=miround(mid(Null_z(request.form("h_total_descuento")),1,8),ndec)
		    rst("base_imponible")=miround(Null_z(request.form("h_base_imponible")),ndec)
            ''ricardo 20-5-2003
		    rst("total_iva")=miround(mid(Null_z(request.form("h_total_iva")),1,8),ndec)
		    rst("recargo")=miround(Null_z(request.form("rf")),decpor)
		    rst("total_recargo")=miround(Null_z(request.form("h_total_rf")), ndec)
		    rst("total_re")=miround(Null_z(request.form("h_total_re")), ndec)
		    rst("irpf")=miround(Null_z(request.form("irpf")),decpor)
		    rst("total_irpf")=miround(Null_z(request.form("h_total_irpf")) , ndec)
            ''ricardo 24/1/2003
		    rst("total_factura")=miround(mid(Null_z(request.form("h_total_factura")),1,8),ndec)
		    old_forma_pago=rst("forma_pago")
		    rst("forma_pago")=Nulear(request.form("forma_pago"))
		    nuev_forma_pago=rst("forma_pago")
		    rst("tipo_pago")=Nulear(request.form("tipo_pago"))
		    rst("transportista")=Nulear(request.form("transportista"))
		    rst("nenvio")=Nulear(request.form("nenvio"))
		    rst("portes")=Nulear(request.form("portes"))
		    rst("fecha_envio")=Nulear(request.form("fechaenvio"))
		    rst("fecha_pedido")=Nulear(request.form("fechapedido"))
		    rst("comercial")=Nulear(request.form("comercial"))
		    rst("agente")=Nulear(request.form("agenteasignado"))
		    rst("cod_proyecto")=Nulear(request.form("cod_proyecto"))
		    rst("observaciones")=Nulear(request.form("observaciones"))
		    rst("notas")=Nulear(request.form("notas"))
		    
            'JFT 23/04/2012 Add fields VALIDATED and VALIDATED_BY
            'if mode="save" and nz_b(request.form("validated")) <> 0 then
            '    rst("validated")=nz_b(request.form("validated"))
            '    rst("validated_by")=session("ncliente") & session("usuario")
            'else
            '    rst("validated_by")=variable_null
            'end if

		    if cstr(bcc)="1" then
			    rst("contabilizado")=Request.Form("h_contabilizada")
		    else
			    rst("contabilizado")=nz_b(Request.Form("contabilizada"))
		    end if
		    rst("contacto")=Nulear(request.form("contacto"))
		    rst("documento")=Nulear(request.form("documento"))
		    ncuenta=Nulear(request.form("ncuenta1")) & Nulear(request.form("ncuenta2")) & Nulear(request.form("ncuenta3")) & Nulear(request.form("ncuenta4"))& Nulear(request.form("ncuenta5"))& Nulear(request.form("ncuenta6"))
		    ncuentra=trim(ncuenta)
            if ncuenta & "">"" then
		        rst("ncuenta")=ncuenta
            else
                rst("ncuenta")=null
            end if
		    if request.form("ncuenta3") & "">"" then
			    'banco=d_lookup("entidad","bancos","codigo='" & Nulear(request.form("ncuenta3")) & "'",dsnilion)

                bancoSelect="select entidad from bancos with(nolock) where codigo = ?"

                banco=DlookupP1(bancoSelect, Nulear(request.form("ncuenta3"))&"",adVarchar ,4 ,dsnilion)

                if banco & "">"" then
                    rst("banco")=trim(banco)
                end if
            else
                rst("banco")=null
		    end if
		    rst("incoterms")=nulear(request.form("incoterms"))
		    rst("fob")=nulear(request.form("fob"))

            ''ricardo 9-1-2007
            'n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))

            n_decimalesSelect="select ndecimales from divisas with(nolock) where codigo=?"
            n_decimales=DlookupP1(n_decimalesSelect, rst("divisa")&"",adVarchar ,15 ,session("dsn_cliente"))

            ''ricardo 28/4/2003 si el usuario ha querido recalcular los importes al cambiar las propiedades de la cabecera
            if request.querystring("recalcular_importes")="1" then
		        'Detectar cambios en la divisa del documento para cambiar la divisa de los detalles(artículos y conceptos)
		        if rst("divisa")<>request.form("h_divisa") and rst("divisa")&"">"" then

                set connDom = Server.CreateObject("ADODB.Connection")
                set commandDom = Server.CreateObject("ADODB.Command")

                connDom.open session("dsn_cliente")
                connDom.cursorlocation=2

                commandDom.ActiveConnection =connDom
                commandDom.CommandTimeout = 60
                commandDom.CommandText = "select * from detalles_fac_cli with(UPDLOCK) where nfactura like ?+'%' and nfactura=? order by item"
                commandDom.CommandType = adCmdText
                commandDom.Parameters.Append commandDom.CreateParameter("@nfacturalike",adVarChar,adParamInput,20,session("ncliente"))
                commandDom.Parameters.Append commandDom.CreateParameter("@nfactura",adVarChar,adParamInput,20, nfactura)

                set rstaux = commandDom.Execute

		            'rstaux.CursorLocation=2
			        'rstaux.open "select * from detalles_fac_cli with(UPDLOCK) where nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic



			        while not rstAux.eof
				        TmpPVP=CambioDivisa(rstAux("pvp"),rst("divisa"),request.form("h_divisa"))
				        rstAux("pvp")=TmpPVP
				        dto1_det=(null_z(TmpPVP)*null_z(rstAux("descuento")))/100
				        dto2_det=((null_z(TmpPVP)-dto1_det)*null_z(rstAux("descuento2")))/100
				        dto3_det=((null_z(TmpPVP)-dto2_det-dto1_det)*null_z(rstAux("descuento3")))/100
				        total_descuento_det=dto1_det+dto2_det+dto3_det

				    ''ricardo 2/5/2008 si en el detalle existe cantidad2 se calculara por esta cantidad
				    if rstAux("cantidad2")<>0 and nz_b(rstAux("calcularimpcantidad2"))=-1 then
				        cantidad_a_coger=rstAux("cantidad2")
				    else
				        cantidad_a_coger=rstAux("cantidad")
				    end if
				        rstAux("importe")=miround((TmpPVP-total_descuento_det)*cantidad_a_coger,n_decimales)
				        rstAux.update
				        rstAux.movenext
			        wend
			        rstAux.close
			        
                    set connDom = Server.CreateObject("ADODB.Connection")
                    set commandDom = Server.CreateObject("ADODB.Command")

                    connDom.open session("dsn_cliente")
                    connDom.cursorlocation=2

                    commandDom.ActiveConnection =connDom
                    commandDom.CommandTimeout = 60
                    commandDom.CommandText = "select * from conceptos with(UPDLOCK) where nfactura like ?+'%' and nfactura=? order by nconcepto"
                    commandDom.CommandType = adCmdText
                    commandDom.Parameters.Append commandDom.CreateParameter("@nfacturalike",adVarChar,adParamInput,20,session("ncliente"))
                    commandDom.Parameters.Append commandDom.CreateParameter("@nfactura",adVarChar,adParamInput,20, nfactura)

                    set rstaux = commandDom.Execute

			        'rstaux.open "select * from conceptos with(UPDLOCK) where nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "' order by nconcepto",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			        while not rstAux.eof
				        TmpPVP=CambioDivisa(rstAux("pvp"),rst("divisa"),request.form("h_divisa"))
				        rstAux("pvp")=TmpPVP

				        ''ricardo 9-1-2007 se cambia la manera de calcular los descuentos
				        dto1_det=(null_z(TmpPVP)*null_z(rstAux("descuento")))/100
				        dto2_det=((null_z(TmpPVP)-dto1_det)*null_z(rstAux("descuento2")))/100
				        dto3_det=((null_z(TmpPVP)-dto2_det-dto1_det)*null_z(rstAux("descuento3")))/100
				        total_descuento_det=dto1_det+dto2_det+dto3_det

				    ''ricardo 2/5/2008 si en el detalle existe cantidad2 se calculara por esta cantidad
				    if rstAux("cantidad2")<>0 and nz_b(rstAux("calcularimpcantidad2"))=-1 then
				        cantidad_a_coger=rstAux("cantidad2")
				    else
				        cantidad_a_coger=rstAux("cantidad")
				    end if
				        rstAux("importe")=miround((TmpPVP-total_descuento_det)*cantidad_a_coger,n_decimales)
				        rstAux.update
				        rstAux.movenext
			        wend
			        rstAux.close
		        end if
            end if

            ''ricardo 28/4/2003 si el usuario ha querido recalcular los importes al cambiar las propiedades de la cabecera
            if request.querystring("recalcular_importes")="1" then
		        'Detectar cambios en la fecha y la tarifa para recalcular los precios de los detalles
		        if (rst("fecha")<>cdate(Request.Form("fecha")&"")) or (rst("tarifa")&""<>request.form("tarifa")&"") then
		            ''eduardo 6-11-2009 se pone el location=2 porque es una transaccion, es un update no una lectura
		            
                    set connDom = Server.CreateObject("ADODB.Connection")
                    set commandDom = Server.CreateObject("ADODB.Command")

                    connDom.open session("dsn_cliente")
                    connDom.cursorlocation=2

                    commandDom.ActiveConnection =connDom
                    commandDom.CommandTimeout = 60
                    commandDom.CommandText = "select * from detalles_fac_cli with(UPDLOCK) where nfactura=? order by item"
                    commandDom.CommandType = adCmdText
                    commandDom.Parameters.Append commandDom.CreateParameter("@nfactura",adVarChar,adParamInput,20,nfactura)

                    rstAux.Open commandDom, , adOpenKeyset, adLockOptimistic
                    'set rstAux = commandDom.Execute

			        'rstAux.open "select * from detalles_fac_cli with(UPDLOCK) where nfactura='" & nfactura & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			        while not rstAux.eof
				        'OrigPVP=d_lookup("pvp","articulos","referencia like '" & session("ncliente") & "%' and referencia='" & rstAux("referencia") & "'",session("dsn_cliente"))

                        OrigPVPSelect="select pvp from articulos with(nolock) where referencia like ? +'%' and referencia = ?"
                        OrigPVP=DlookupP2(OrigPVPSelect, session("ncliente")&"" ,adVarchar ,30 ,rstAux("referencia")&"",adVarchar ,30 ,session("dsn_cliente"))

				        'OrigDIV=d_lookup("divisa","articulos","referencia like '" & session("ncliente") & "%' and referencia='" & rstAux("referencia") & "'",session("dsn_cliente"))

                        OrigDIVSelect="select divisa from articulos with(nolock) where referencia like ? +'%' and referencia = ?"
                        OrigDIV=DlookupP2(OrigDIVSelect, session("ncliente")&"", adVarchar, 30, rstAux("referencia")&"", adVarchar, 30, session("dsn_cliente"))

				        DtoAplicado=0
				        TmpPVP=PrecioArticulo(rstAux("referencia"),cdate(Request.Form("fecha")&""),rstAux("cantidad"),request.form("tarifa")&"",OrigPVP,DtoAplicado)
				        ''ricardo 15-11-2005 si viene un descuento de precioarticulo se pondra este.
				        if DtoAplicado<>0 then
					        rstAux("descuento")=-DtoAplicado
				        end if
				        if OrigDIV<>request.form("h_divisa") then
					        TmpPVP=CambioDivisa(TmpPVP,OrigDIV,request.form("h_divisa"))
				        end if
				        rstAux("pvp")=TmpPVP

				        ''ricardo 9-1-2007 se cambia la manera de calcular los descuentos
				        dto1_det=(null_z(TmpPVP)*null_z(rstAux("descuento")))/100
				        dto2_det=((null_z(TmpPVP)-dto1_det)*null_z(rstAux("descuento2")))/100
				        dto3_det=((null_z(TmpPVP)-dto2_det-dto1_det)*null_z(rstAux("descuento3")))/100
				        total_descuento_det=dto1_det+dto2_det+dto3_det

				    ''ricardo 2/5/2008 si en el detalle existe cantidad2 se calculara por esta cantidad
				    if rstAux("cantidad2")<>0 and nz_b(rstAux("calcularimpcantidad2"))=-1 then
				        cantidad_a_coger=rstAux("cantidad2")
				    else
				        cantidad_a_coger=rstAux("cantidad")
				    end if
				        rstAux("importe")=miround((TmpPVP-total_descuento_det)*cantidad_a_coger,n_decimales)
				        ''''''FIN ricardo 9-1-2007''''''''''''''''''''''''
				        rstAux.update
				        rstAux.movenext
			        wend
			        rstAux.close
		        end if
            end if

            ''ricardo 10-8-2004 al cambiar el cliente se cambiara el iva de los detalles si el cliente tiene su campo iva relleno
            han_cambiado_importes_cliente=0
            if request.querystring("recalcular_importes")="1" then
	            if rst("ncliente")<>request.form("h_ncliente") and rst("ncliente")&"">"" then
		            han_cambiado_importes_cliente=1
	            end if
            end if

		    rst("tarifa")=Nulear(request.form("tarifa"))
		    rst("fecha")=Nulear(Request.Form("fecha"))
		    rst("divisa")=Nulear(request.form("h_divisa"))
		    '*** AMP Añadimos la insercion del campo factor de cambio.
		    rst("factcambio")=miround(Nulear(limpiaCadena(request.form("nfactcambio"))),DEC_PREC)
		
		    rst("nfolio")=Nulear(request.form("nfolio"))

		    '' JMA 28/10/04 Actualizamos los campos personalizables
		    num_campos=limpiaCadena(request.querystring("num_campos"))
		    if num_campos="" then
			    num_campos=limpiaCadena(request.form("num_campos"))
		    end if

	    if num_campos & "">"" then
		    redim lista_valores(num_campos+10,2)
		    for ki=1 to num_campos
			    nom_campo="campo" & ki
			    valor_form=Nulear(limpiaCadena(request.querystring(nom_campo)))
			    if valor_form & ""="" then
				    valor_form=Nulear(limpiaCadena(request.form(nom_campo)))
			    end if
			    cadena_campo="" & replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
			    
                set connDom = Server.CreateObject("ADODB.Connection")
                set commandDom = Server.CreateObject("ADODB.Command")

                connDom.open session("dsn_cliente")
                connDom.cursorlocation=3

                commandDom.ActiveConnection =connDom
                commandDom.CommandTimeout = 60
                commandDom.CommandText = "select titulo,tipo from camposperso with(nolock) where ncampo=? and tabla='DOCUMENTOS VENTA' order by SECCIONCP,ncampo,titulo"
                commandDom.CommandType = adCmdText
                commandDom.Parameters.Append commandDom.CreateParameter("@ncampo",adChar,adParamInput,7,session("ncliente") & cadena_campo)

                set rstaux = commandDom.Execute
			    'rstAux.open "select titulo,tipo from camposperso with(nolock) where ncampo='" & session("ncliente") & cadena_campo & "' and tabla='DOCUMENTOS VENTA' order by SECCIONCP,ncampo,titulo",session("dsn_cliente")
			    if not rstAux.EOF then
			        tipo_campo_perso=rstAux("tipo")
			        titulo_campo_perso=rstAux("titulo")
			    end if
			    rstAux.Close
                lista_valores(ki,2)=titulo_campo_perso
			    if tipo_campo_perso=2 then
				    if valor_form="on" then
					    lista_valores(ki,1)=1
				    else
					    lista_valores(ki,1)=0
				    end if
			    else
		            lista_valores(ki,1)=valor_form
			    end if
		    next
	    else
		    redim lista_valores(num_campos_ventas+5,2)
		    for ki=1 to num_campos_ventas+5
			    lista_valores(ki,1)=""
			    lista_valores(ki,2)=""
		    next
	    end if

            ''ricardo 27-3-2009 se guardaran tantos campos como existan en tabla y con titulo
            ''sera con titutlo, por si utilizamos internamente para alguna cosa, pero claro el usuario no puede tocarla
            ''ricardo 27-3-2009 en lugar de poner una linea por campo, se pone un for que lo hago por la totalidad
            ''ademas solamente se guardara el valor, en aquellos campos que tengan titulo
            for ki_cli=1 to num_campos_ventas
                if lista_valores(ki_cli,2) & "">"" then
                    cadena_campo="campo" & replace(space(2-len(cstr(ki_cli)))," ","0") & cstr(ki_cli)
                    rst(cadena_campo)=lista_valores(ki_cli,1)
                end if
            next

		    '' JMA 28/10/04 Fin actualizar campos personalizables
		    '**RGU 28/6/2007: por alguna razon esto ha dejado de funcionar y lo que devuelve el recordset cuando esta cobrada (verdadero, true) no es mayor que cero(paso de cobrada a descobrada)
		    if nz_b(rst("cobrada"))<>nz_b(Request.Form("cobrada")) then 'ha habido cambios en el estado de la factura
		    '**RGU 28/6/2007***
			    if nz_b(Request.Form("cobrada")) <> 0 then 'DE NO COBRADA A COBRADA
			    
			        'hayVencim=d_lookup("nfactura","vencimientos_salida","nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "' and cobrado=0",session("dsn_cliente"))

                    hayVencimSelect="select nfactura from vencimientos_salida with(nolock) where nfactura like ?+'%' and nfactura=? and cobrado=0"
                    hayVencim=DlookupP2(hayVencimSelect, session("ncliente")&"",adVarchar ,20 , nfactura&"",adVarchar ,20 ,session("dsn_cliente"))

			        'hayVencim2=d_lookup("ndocument","nonpayment","ndocument like '" &  nfactura & "%'",session("dsn_cliente"))

                    hayVencim2Select="select ndocument from nonpayment with(nolock) where ndocument like ?+'%'"
                    hayVencim2=DlookupP1(hayVencim2Select, nfactura&"",adVarchar ,20 , session("dsn_cliente"))

				    if hayVencim<>nfactura then rst("cobrada")=1 
				    if hayVencim2=hayVencim then rst("cobrada")=1
				    'Actualizar el registro.Obtener el siguiente nº de documento de la tabla series.
				    if nfactura="" then
					    SigDoc=CalcularNumDocumento(nserie,fecha)
					    rst("nfactura")=SigDoc
					    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),SigDoc,"cliente")
					    if ref_edi>"" then
						    rst("edi")=ref_edi
					    else
						    rst("edi")=NULL
					    end if
				    else
					    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),nfactura,"cliente")
					    if ref_edi>"" then
						    rst("EDI")=ref_edi
					    else
						    rst("edi")=NULL
					    end if
				    end if
				    rst.Update
				    CobrarVencimientos nfactura
			    else 'DE COBRADA A NO COBRADA
				    'Comprobar la caja.
                    set connDom = Server.CreateObject("ADODB.Connection")
                    set commandDom = Server.CreateObject("ADODB.Command")

                    connDom.open session("dsn_cliente")
                    connDom.cursorlocation=3

                    commandDom.ActiveConnection =connDom
                    commandDom.CommandTimeout = 60
                    commandDom.CommandText = "select ndocumento from caja with(nolock) where ndocumento like ?+'%' and ndocumento=? or (tdocumento='VENCIMIENTO_SALIDA' and exists (select nrecibo from vencimientos_salida with(nolock) where ndocumento = nrecibo and nfactura=?))"
                    commandDom.CommandType = adCmdText
                    commandDom.Parameters.Append commandDom.CreateParameter("@ndoc",adVarChar,adParamInput,22,session("ncliente"))
                    commandDom.Parameters.Append commandDom.CreateParameter("@ndocumento",adVarChar,adParamInput,22,nfactura)
                    commandDom.Parameters.Append commandDom.CreateParameter("@nfactura",adVarChar,adParamInput,20,nfactura)

                    set rstaux = commandDom.Execute
                    
				    'rstAux.open "select ndocumento from caja with(nolock) where ndocumento like '" & session("ncliente") & "%' and ndocumento='" & nfactura & "' or (tdocumento='VENCIMIENTO_SALIDA' and exists (select nrecibo from vencimientos_salida with(nolock) where ndocumento = nrecibo and nfactura='" & nfactura & "'))",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstAux.EOF then
					    EnCaja="NO"
					    rstAux.close
					    'Obtener el siguiente nº de documento de la tabla series.
					    if nfactura="" then
						    SigDoc=CalcularNumDocumento(nserie,fecha)
						    rst("nfactura")=SigDoc
						    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),SigDoc,"cliente")
						    if ref_edi>"" then
							    rst("edi")=ref_edi
						    else
							    rst("edi")=NULL
						    end if
					    else
						    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),nfactura,"cliente")
						    if ref_edi>"" then
							    rst("EDI")=ref_edi
						    else rst("edi")=NULL
						    end if
					    end if
					    rst.update
					    AnularVencimientos nfactura

                        set connDom = Server.CreateObject("ADODB.Connection")
                        set commandDom = Server.CreateObject("ADODB.Command")

                        connDom.open session("dsn_cliente")
                        connDom.cursorlocation=3

                        commandDom.ActiveConnection =connDom
                        commandDom.CommandTimeout = 60
                        commandDom.CommandText = "update facturas_cli with(rowlock) set cobrada=0 where nfactura like ?+'%' and nfactura=?"
                        commandDom.CommandType = adCmdText
                        commandDom.Parameters.Append commandDom.CreateParameter("@nfac",adVarChar,adParamInput,20,session("ncliente"))
                        commandDom.Parameters.Append commandDom.CreateParameter("@nfactura",adVarChar,adParamInput,20,nfactura)

                        set rstaux = commandDom.Execute

					    'rstAux.open "update facturas_cli with(rowlock) set cobrada=0 where nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    else
					    rstAux.close
					    'Obtener el siguiente nº de documento de la tabla series.
					    if nfactura="" then
						    SigDoc=CalcularNumDocumento(nserie,fecha)
						    rst("nfactura")=SigDoc
						    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),SigDoc,"cliente")
						    if ref_edi>"" then
							    rst("edi")=ref_edi
						    else
							    rst("edi")=NULL
						    end if
					    else
						    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),nfactura,"cliente")
						    if ref_edi>"" then
							    rst("EDI")=ref_edi
						    else
							    rst("edi")=NULL
						    end if
					    end if
					    rst.update
					    EnCaja="SI"%>
					    <script language="javascript" type="text/javascript">
                            alert("<%=LitMsgNoAbularCobroAnotCaja%>");
					    </script>
				    <%end if
			    end if
		    else
			    'Obtener el siguiente nº de documento de la tabla series.
			    if nfactura="" then
				    SigDoc=CalcularNumDocumento(nserie,fecha)
				    rst("nfactura")=SigDoc
				    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),SigDoc,"cliente")
				    if ref_edi>"" then
					    rst("edi")=ref_edi
				    else
					    rst("edi")=NULL
				    end if
			    else
				    ref_edi=CalculaEDI(nserie,Nulear(session("ncliente")&request.form("ncliente")),nfactura,"cliente")
				    if ref_edi>"" then
					    rst("EDI")=ref_edi
				    else
					    rst("edi")=NULL
				    end if
			    end if

			    rst.update
		    end if

		    ''ricardo 10-8-2004 al cambiar el cliente se cambiara el iva de los detalles y conceptos si el cliente tiene su campo iva relleno
		    if han_cambiado_importes_cliente=1 then
			    stractdet="update detalles_fac_cli with(rowlock)"
			    stractdet=stractdet & " set iva=isnull((select c2.iva from clientes as c2 with(nolock) where c2.ncliente='" & rst("ncliente") & "'),isnull(a.iva,(select iva from configuracion with(nolock) where nempresa='" & session("ncliente") & "'))) "
			    stractdet=stractdet & " ,re=isnull((select t2.re from clientes as c2 with(nolock) inner join tipos_iva as t2 with(nolock) on t2.tipo_iva=c2.iva where c2.ncliente='" & rst("ncliente") & "' ),isnull(t.re,(select t.re from configuracion as c with(nolock) inner join tipos_iva as t with(nolock) on t.tipo_iva=c.iva where c.nempresa='" & session("ncliente") & "' ))) "
			    stractdet=stractdet & " from detalles_fac_cli as d inner join articulos as a with(nolock) on a.referencia=d.referencia "
			    stractdet=stractdet & " left outer join tipos_iva as t with(nolock) on t.tipo_iva=a.iva "
			    stractdet=stractdet & " where nfactura='" & nfactura & "' "
			    stractdet=stractdet & " and d.nfactura like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' "
			    rstAux.open stractdet,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    if rstAux.state<>0 then rstAux.close

			    stractdet="update conceptos with(rowlock)"
			    stractdet=stractdet & " set iva=isnull((select c2.iva from clientes as c2 with(nolock) where c2.ncliente='" & rst("ncliente") & "'),(select iva from configuracion with(nolock) where nempresa='" & session("ncliente") & "')) "
			    stractdet=stractdet & " ,re=isnull((select t2.re from clientes as c2 with(nolock) inner join tipos_iva as t2 with(nolock) on t2.tipo_iva=c2.iva where c2.ncliente='" & rst("ncliente") & "' ),(select t.re from configuracion as c with(nolock) inner join tipos_iva as t with(nolock) on t.tipo_iva=c.iva where c.nempresa='" & session("ncliente") & "' )) "
			    stractdet=stractdet & " from conceptos as d "
			    stractdet=stractdet & " where nfactura='" & nfactura & "' "
			    stractdet=stractdet & " and d.nfactura like '" & session("ncliente") & "%' "
			    rstAux.open stractdet,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    if rstAux.state<>0 then rstAux.close
		    end if

		    ''ricardo 24/1/2003 cuando habia una factura de abono, no se recalculaban los totales y salian mal por lo que lo pongo que se ejecute siempre, que no sea la primera vez

		    if nfactura  & "">"" then
		    if rstAux.state<>0 then rstAux.close
			    Precios nfactura
                ''ricardo 8-5-2003 solo se actualizara la deuda cuando no pase de no cobrado a cobrado haya habido cambios en la fecha,tarifa,cliente y divisa
                si_actualizar=0
                if nz_b(Request.Form("cobrada"))=0 and request.querystring("recalcular_importes")="1" then
	                si_actualizar=1
                end if

                if si_actualizar=1 then
                elseif se_han_cambiado_los_descuentos=1 then
	                if nz_b(Request.Form("cobrada")) <> 0 then
		                rstAux.open "update facturas_cli with(updlock) set cobrada=1 where nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	                end if
                end if
		    end if
		
		    'ahora modificamos el comercial de los vencimientos , si los hubiera
		    ModificarVencimientos comercial_antR,old_forma_pagoR,cambiarcomR

		    if ModDocumentoEquip=true then
			    InsertarHistorialNserie "OK1","","","FACTURA A CLIENTE",nfactura,"","","","","MODIFY",mode
		    end if
	    end if
	    '*** i AMP 25/07/2011 Gestión de impagos.
	    if mode="first_save" and vienenp>"" then	 
	        GenerarVincFacturaImpago vienenp,sigDoc ,datenote
	    end if
	    '*** f AMP 
    end sub

    '******************************************************************************
    'Da por cobrados los vencimientos de la factura
    sub CobrarVencimientos(nfactura)
	    rstAux.Open "update vencimientos_salida with(rowlock) set importecob=importe,cobrado=1 where nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    end sub

    '******************************************************************************
    'Anula los cobros de vencimientos de la factura
    sub AnularVencimientos(nfactura)
	    rstAux.Open "update vencimientos_salida with(rowlock) set importecob=0,cobrado=0 where nfactura like '" & session("ncliente") & "%' and nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    end sub

    '******************************************************************************
    function ComprobacionAtribPuntos(nfactura)
	    if rst.state<>0 then rst.close
   	    no_se_puede_borrar=0
	    if si_tiene_modulo_fidelizacion<>0 then
		    'miramos si esta factura habia generado puntos, si los hubo , habra que quitarlos
		    if VerObjeto(OBJFidelizacion)=true then
			    rst.cursorlocation=3
			    'ega 19/06/2008 solo el campo que se necesita
			    rst.open "select total_factura from facturas_cli as a with(nolock) where a.nfactura like '" & session("ncliente") & "%' and a.nfactura='" & nfactura & "'",session("dsn_cliente")
			    if not rst.eof then
				    importe=null_z(rst("total_factura"))
			    else
				    importe=0
			    end if
			    rst.close
			    set conn = Server.CreateObject("ADODB.Connection")

			    ''ricardo 16/1/2003
			    ''se pone estos procedimientos, ya que en el procedimiento sp_Guardar_Puntos_Atribucion
			    ''se necesita saber el saldo de la tarjeta, pero como la tabla tarjetas esta en ilion
			    '' y el usuario de la base de datos del cliente, no tiene permisos, para acceder a ilion
			    '' es decir, desde un procedimiento que se ejecuta en la bd del cliente, no tiene acceso a ilion
			    ''por eso estos dos procedimientos obtienen la tarjeta de atrib_puntos en esta factura
			    ''y con esta tarjeta el saldo, que pasamos al procedimiento sp_Guardar_Puntos_Atribucion
			    ''como el parametro puntos, que en este caso no se estaba utilizando

			    conn.open session("dsn_cliente")
			    strselect="Exec obtener_tarjeta @ndocumento='" & nfactura & "',@viene='facturas_cli'"
			    set rs = conn.execute(strselect)
			    if not rs.eof then
				    ntarjeta=rs(0)
			    else
				    ntarjeta=""
			    end if
			    rs.close
			    conn.close
			    if ntarjeta & "">"" then
				    conn.open dsnilion
				    strselect="Exec sp_saldo_tarjeta @ntarjeta='" & ntarjeta & "'"
				    set rs = conn.execute(strselect)
				    if not rs.eof then
					    saldo=rs(0)
				    else
					    saldo=0
				    end if
				    rs.close
				    conn.close

				    conn.open session("dsn_cliente")
				    strselect="Exec sp_Guardar_Puntos_Atribucion @ndocumento='" & nfactura & "',@ntarjeta='',@importe='" & reemplazar(importe,",",".") & "',@usuario='" & session("usuario") & "',@viene='facturas_cli',@modo='BORRAR',@comercio='" & session("ncliente") & "',@puntos='" & saldo & "',@codop='" & codop & "'"
				    set rs = conn.execute(strselect)
				    BorrarRegistro1= 0
				    if not rs.eof then
					    ntarjeta=rs(0)
					    if ntarjeta="-2" then
						    BorrarRegistro1=0
						    no_se_puede_borrar=1%>
						    <script language="javascript" type="text/javascript">
                                alert("<%=LitNoPueBorrHayAtriEnFac%>");
                                document.location = "facturas_cli.asp?mode=browse&nfactura=<%=enc.EncodeForHtmlAttribute(nfactura)%>" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
                                parent.botones.location = "facturas_cli_bt.asp?mode=browse";
						    </script>
					    <%elseif ntarjeta="-1" then
						    BorrarRegistro1=0
						    no_se_puede_borrar=1%>
						    <script language="javascript" type="text/javascript">
                                alert("<%=LitNoPueBorrQueSaldNeg%>");
                                document.location = "facturas_cli.asp?mode=browse&nfactura=<%=enc.EncodeForHtmlAttribute(nfactura)%>" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
                                parent.botones.location = "facturas_cli_bt.asp?mode=<%=iif(mode="save","edit","add")%>";
						    </script>
					    <%elseif ntarjeta="-3" then
						    BorrarRegistro1=0
						    no_se_puede_borrar=1
						    'no existe una tarjeta en la tabla atrib_puntos para esta factura,pero si existe un registro'
					    elseif ntarjeta="-4" then
						    BorrarRegistro1=0
						    no_se_puede_borrar=1%>
						    <script language="javascript" type="text/javascript">
                                alert("<%=LitNoPueBorrHayAtrib%>");
                                document.location = "facturas_cli.asp?mode=browse&nfactura=<%=enc.EncodeForHtmlAttribute(nfactura)%>" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
                                parent.botones.location = "facturas_cli_bt.asp?mode=browse";
						    </script>
					    <%else
						    BorrarRegistro1=1
					    end if
				    else
					    puntos=0
					    ntarjeta=""
					    BorrarRegistro1=0
				    end if
				    conn.close
			    else
				    puntos=0
				    ntarjeta=""
				    BorrarRegistro1=0
			    end if

			    if BorrarRegistro1=1 and ntarjeta & "">"" then
				    'Ahora creamos la operacion
				    conn.open dsnilion
				    strselect="Exec sp_Crear_Operaciones @ntarjeta='" & ntarjeta & "',@modo='BORRAR',@comercio='" & session("ncliente") & "',@puntos='" & puntos & "'"
				    set rs = conn.execute(strselect)
				    BorrarRegistro1= 0
				    if not rs.eof then
					    codop=rs(0)
					    BorrarRegistro1=1
				    else
					    codop=""
					    BorrarRegistro1=0
				    end if
				    rs.close
				    conn.close

				    if BorrarRegistro1=1 then
					    if ntarjeta & "">"" then
						    'ahora actualizamos el saldo de la tarjeta
						    conn.open dsnilion
						    strselect="Exec sp_Actualizar_Tarjetas @ntarjeta='" & ntarjeta & "',@modo='BORRAR',@viene='facturas_cli',@puntos='" & puntos & "'"
						    set rs = conn.execute(strselect)
						    BorrarRegistro1= 0
						    if not rs.eof then
							    BorrarRegistro1= rs(0)
						    else
							    BorrarRegistro1=0
						    end if
						    rs.close
						    conn.close
					    end if
				    end if
			    end if
			    set conn=nothing
		    end if
	    end if
	    ComprobacionAtribPuntos=no_se_puede_borrar
    end function

    'Elimina los datos del registro cuando se pulsa BORRAR.
    sub BorrarRegistro(nfactura)
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
        cmd("tipo_documento") = "FACTURA A CLIENTE"
        set referencias=cmd.Execute
               
        'borramos la factura
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open session("dsn_cliente")
        Set cmd1 = Server.CreateObject("ADODB.Command")
        Set cmd1.ActiveConnection = cn
        cmd1.CommandText = "BorrarDocumento"
        cmd1.CommandType = adCmdStoredProc
        ''cmd1.ConnectionTimeout = 90
        cmd1.CommandTimeout = 90
        cmd1.Parameters.Append cmd1.CreateParameter("result", adInteger,adParamReturnValue)
        cmd1.Parameters.Append cmd1.CreateParameter("ndocumento", adChar,adParamInput,20)
        cmd1.Parameters.Append cmd1.CreateParameter("tipo_documento", adChar, adParamInput,50)

        cmd1("ndocumento") = nfactura
        cmd1("tipo_documento") = "FACTURA A CLIENTE"
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
		                if Stock<0 then %>
			                <script language="javascript" type="text/javascript">alert("<%=LitMsgStockNegativo%> <%=trimCodEmpresa(referencia)%>");</script>
			            <%elseif Stock < StockMin then %>
			                <script language="javascript" type="text/javascript">alert("<%=LitMsgStockBajoMin%> <%=trimCodEmpresa(referencia)%>");</script>
			            <%end if
		                if PendRecibir<0 then %>
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
            <script language="javascript" type="text/javascript">alert("<%=mensaje%>");</script> 
        <%end if
    end sub

    '******************************************************************************
    'Crea la tabla que contiene la barra de grupos de datos.
    sub BarraNavegacion(modo)
        if modo="add" or modo="edit" then%>
            <script language="javascript" type="text/javascript">
                jQuery("#S_DatosGenerales").show();
            </script>
        <%else%>
            <script language="javascript" type="text/javascript">
                jQuery("#S_DatosGenerales").hide();
            </script>
        <%end if
    end sub

    '****************************************************************************************************************
    function CerrarTodo()
        set connRound = nothing
        set rstMM = nothing
        set rsTPV = nothing
        set rst2 = nothing
        set rstAuxREM = nothing
        set con = nothing
        set command = nothing
        set rstCrearLineasFacZonas = nothing
        set connNPinv = nothing
        set commandNPinv = nothing
        set cn = nothing
        set cmd = nothing
        set cmd1 = nothing
        set SoapRequest  = nothing
        set rstAux = nothing
        set rstAux2 = nothing
        set rstCliente = nothing
        set rst = nothing
        set rstDomi = nothing
        set rstSelect = nothing
        set rstIvas = nothing
        set rstConta = nothing
        set rstBlq = nothing
        set rstFactura = nothing
        set rs = nothing
        set referencias = nothing
        set result = nothing
    end function

    Function protegerFacturaSAFT(p_nfactura)    
        Dim returnString 
        Dim SoapRequest 
        Dim SoapURL 

        Set SoapRequest = Server.CreateObject("MSXML2.XMLHTTP") 
  
        SoapURL = "http://localhost/IlionServices/Integracion/IntegracionS.asmx/protegerFactura" 
        Dim DataToSend
        DataToSend="ncliente="&session("ncliente")&"&usuario="&session("usuario")&"&idSesion="&session.sessionid&"&nfactura="&p_nfactura

    
        SoapRequest.Open "POST",SoapURL , False 
        SoapRequest.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
        SoapRequest.Send DataToSend
    
        result=SoapRequest.responseText
    
        Set SoapRequest = Nothing 
  
        protegerFacturaSAFT = result 
    End Function 


    Const NUMBER_PADDING = "000000000000" ' a few zeroes more just to make sure ' 
    Function ZeroPadInteger(i, numberOfDigits) 
      ZeroPadInteger = Right(NUMBER_PADDING & i, numberOfDigits) 
    End Function 

    Function obtenerFacturaAnterior(nfactura)    
        longitud = len(nfactura)
        nfactAnt = cint(mid(nfactura, longitud - 5, longitud)) -1
        if nfactAnt = 0 then
            result=""
        else
            nfactFormateado= ZeroPadInteger (nfactAnt,5)
            result = mid(nfactura, 1, longitud - 5)  & nfactFormateado
        end if
        obtenerFacturaAnterior = result 
    End Function

    Function obtenerFacturaSiguiente(nfactura)    
        longitud = len(nfactura)
        nfactSig = cint(mid(nfactura, longitud - 5, longitud)) +1
        if nfactSig = 0 then
            result=""
        else
            nfactFormateado= ZeroPadInteger (nfactSig,5)
            result = mid(nfactura, 1, longitud - 5)  & nfactFormateado
        end if
        obtenerFacturaSiguiente = result 
    End Function  

    '*****************************************************************************
    '********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
    '*****************************************************************************
    const borde=0
    dim EnCaja, EnEfecto,EnRemesa%>
    <form name="facturas_cli" method="post">
	    <%
        PintarCabecera "facturas_cli.asp"
        Alarma "facturas_cli.asp"
        ' Ocultar detalle de las facturas
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

	    'Leer parámetros de la página
	    mode=Request.QueryString("mode")
	
	    vienenp=limpiaCadena(Request.QueryString("vienenp"))
	    datenote=limpiaCadena(Request.QueryString("datent"))
	    %><input type="hidden" name="h_vienenp" value="<%=enc.EncodeForHtmlAttribute(vienenp)%>"/>
	    <input type="hidden" name="h_datenote" value="<%=enc.EncodeForHtmlAttribute(datenote)%>"/><%	
	
        CheckCadena p_nfactura
	
	    dim modp,modd,modi,novei,cv,caju,ocb,bcc,s,pciva,tfi,cpc, nmc, rnf, pnf,modn,pagsl,wclc
	    ObtenerParametros("facturas_cli_det")

	    'cag
	    set rstAux = Server.CreateObject("ADODB.Recordset")
	    set rstAux2 = Server.CreateObject("ADODB.Recordset")
	    set rstCliente = Server.CreateObject("ADODB.Recordset")
	    set rst = Server.CreateObject("ADODB.Recordset")
	    set rstDomi = Server.CreateObject("ADODB.Recordset")
	    set rstSelect = Server.CreateObject("ADODB.Recordset")
	    set rstIvas = Server.CreateObject("ADODB.Recordset")
	    set rstConta = Server.CreateObject("ADODB.Recordset")
	    set rstBlq = Server.CreateObject("ADODB.Recordset")
	    set rstFactura = Server.CreateObject("ADODB.Recordset")
        'saft = d_lookup("saft","configuracion","nempresa = '" & session("ncliente")&"'",session("dsn_cliente"))

        saftSelect="select saft from configuracion with (nolock) where nempresa = ?"
        saft=DlookupP1(saftSelect,session("ncliente")&"",adChar ,5 ,session("dsn_cliente"))

          saft=nz_b(saft)
        if saft = True then
            'hay casos, por ejemplo al bloquear un albaran y convertirlo a factura donde ésta tiene el campo ahora a 1, es decir esta bloqueada
            ' pero no tiene firma por lo que ponemos el campo ahora a 0
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("dsn_cliente")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="update  [FACTURAS_CLI]  with(updlock) set ahora=0  where nfactura=? and ahora=1 and (campo20 is null or campo20='')"
            command.CommandType = adCmdText 'Consulta Parametrizada
            command.Parameters.Append command.CreateParameter("@nfactura",adVarChar,adParamInput,20,p_nfactura)
        
            set result= command.Execute
        
            conn.close
            set command=nothing
            set conn=nothing
        end if

	    if mode="bloqueo" then
	        saftOK=true
      
	        if saft = True then
	            set conn = Server.CreateObject("ADODB.Connection")
	            set command =  Server.CreateObject("ADODB.Command")
	            conn.open session("dsn_cliente")
	            command.ActiveConnection =conn
	            command.CommandTimeout = 0
	            command.CommandText="SELECT [NFACTURA] , serie ,[DESCUENTO] ,[DESCUENTO2]  ,[DESCUENTO3],Convert(money,IMPORTE_BRUTO) as IMPORTE_BRUTO, campo10 , campo11 , campo12,  fc.ncliente as clienteFactura FROM [FACTURAS_CLI] fc with(nolock)  where nfactura=?"
	            command.CommandType = adCmdText 'Consulta Parametrizada
	            command.Parameters.Append command.CreateParameter("@nfactura",adVarChar,adParamInput,20,p_nfactura)
	        
                set result= command.Execute
                while not result.EOF
                    Dtotemp=result("descuento")
                    Dtotemp2=result("descuento2")
                    Dtotemp3=result("descuento3")
                
                    'Comprobamos los descuentos
                    if Dtotemp < 0 or Dtotemp2 <0 or Dtotemp3 < 0 or Dtotemp > 100 or Dtotemp2 > 100 or Dtotemp3 > 100 then
                        saftOK=false%>
	                    <script language="javascript" type="text/javascript">
                            window.alert("<%=LITDESCUENTOERROR%>");
			            </script> 
                    <%end if
                    ' Nos guardamos el campo de excencion de iva para comprobarlo más tarde en los detalles
                    campo10 = result("campo10")&""
                    serie="FT"
                
                    ' si es una nota de crédito la marcamos para luego
                    if Instr(result("serie"), "NC") then
                        serie="NC"
                        if result("importe_bruto")<0 then
                            'no hay comprobar la referencia por lo que no nos hace falta comprobar el campo 11
                           ' campo11 = result("campo11")&""
                            campo12 = result("campo12")&""
                            if campo12="" then
                                saftOK=false%>
	                            <script language="javascript" type="text/javascript">
                                    window.alert("<%=LITREFERENCIAABONOERROR%>");
			                    </script> 
                            <%end if
                        end if
                    else
                        'comprobamos que la serie FT no contine al cliente contado
                        if Instr(result("serie"), "NC")= false  and Instr(result("serie"), "VD")= false  then
                            if result("clienteFactura") = session("ncliente")&"00000" then
                                saftOK=false%>
                                <script language="javascript" type="text/javascript">
                                    window.alert("<%=LITCLIENTECONTADOFT%>");
	                            </script> 
                            <%end if
                        end if
                    end if
                    'ultima comprobación para ver que realmente el cliente contado solo está en la serie VD o NC
                    if result("clienteFactura") = session("ncliente")&"00000" and saftOK=true and ( Instr(result("serie"), "NC")= false  and Instr(result("serie"), "VD")= false)then
                        saftOK=false%>
                        <script language="javascript" type="text/javascript">
                                    window.alert("aqui" + "<%=LITCLIENTECONTADOFT%>");
                        </script> 
                    <%end if
                    result.MoveNext
               wend
	        
	            conn.close
	            set command=nothing
	            set conn=nothing
	        
	            'comprobamos ahora que las lineas de detalle sean correctas
	            if saftOK=true then
	                set conn = Server.CreateObject("ADODB.Connection")
	                set command =  Server.CreateObject("ADODB.Command")
	                conn.open session("dsn_cliente")
	                command.ActiveConnection =conn
	                command.CommandTimeout = 0
	                command.CommandText="SELECT  [NFACTURA]    ,[IMPORTE] ,[DESCUENTO]    ,[DESCUENTO2]     ,[DESCUENTO3]     ,[IVA]      FROM [DETALLES_FAC_CLI] with(nolock)  where nfactura =? union  SELECT        [NFACTURA]      ,[IMPORTE]      ,[DESCUENTO]      ,[DESCUENTO2]      ,[DESCUENTO3]      ,[IVA]   FROM [CONCEPTOS] with(nolock)	  where nfactura =?"
	                command.CommandType = adCmdText 'Consulta Parametrizada
	                command.Parameters.Append command.CreateParameter("@nfactura",adVarChar,adParamInput,20,p_nfactura)
    	            command.Parameters.Append command.CreateParameter("@nfactura2",adVarChar,adParamInput,20,p_nfactura)
                    set result= command.Execute
                    detalle=false
                    while not result.EOF
                        detalle=true
                        Dtotemp=result("descuento")
                        Dtotemp2=result("descuento2")
                        Dtotemp3=result("descuento3")
                    
                        'Comprobamos los descuentos
                        if Dtotemp < 0 or Dtotemp2 <0 or Dtotemp3 < 0 or Dtotemp > 100 or Dtotemp2 > 100 or Dtotemp3 > 100 then
                            saftOK=false%>
	                        <script language="javascript" type="text/javascript">
                                window.alert("<%=LITDESCUENTOERROR%>");
			                </script>
                        <%end if

                        'Comprobamos la excencion de iva
                        if result("IVA")=0 and campo10 = "" then
                            saftOK=false%>
	                        <script language="javascript" type="text/javascript">
                                window.alert("<%=LITIVAERROR%>");
			                </script> 
                        <%end if 

                        ' Comprobamos las líneas positivas y negativas dependiendo de la serie
                        if  serie = "NC" then
                            if result("importe") > 0 then
                                saftOK=false%>
	                            <script language="javascript" type="text/javascript">
                                    window.alert("<%=LITIMPORTENCERROR%>");
			                    </script> 
	                        <%end if
                        else 
                            if result("importe") < 0 then
                                 saftOK=false%>
	                             <script language="javascript" type="text/javascript">
                                     window.alert("<%=LITIMPORTEFTERROR%>");
			                    </script> 
                            <%end if
                        end if
                    
                        result.MoveNext
                    wend
    	        
	                conn.close
	                set command=nothing
	                set conn=nothing 
	            end if
	            if  detalle =false and saftOK=true then
	                saftok=false%>
                    <script language="javascript" type="text/javascript">
                                     window.alert("<%=LITDETALLESVACIO%>");
	                </script> 
                <%end if
	        

	            factAnt = obtenerFacturaAnterior (p_nfactura)
	            if saftOK=true then
	                if factAnt<>"" then
	                    'hash = d_lookup("campo20","facturas_cli","nfactura = '" & factAnt&"'",session("dsn_cliente"))

                        hashSelect ="select campo20 from facturas_cli with(nolock) where nfactura=?"
                        hash = DlookupP1(hashSelect,factAnt&"",adVarchar ,20 ,session("dsn_cliente"))

	                    if hash <> "" then
	                        protegerFacturaSAFT p_nfactura
	                        saftOK=true
	                    else
	                        'hacemos un hack para la primera factura
                           if hash="aaaaaaaaaaXXXXXXXXXXZZZZZZZZZZ1111111111" then
                                protegerFacturaSAFT p_nfactura
	                            saftOK=true
                            else%>
	                             <script language="javascript" type="text/javascript">
                                     window.alert("<%=LitFactAntNoBloqueada%>");
			                    </script> 
    			                <%saftOK=false
    			            end if
	                    end if
	                else
	                    protegerFacturaSAFT p_nfactura
	                    saftOK=true
	                end if
	           end if 
	        end if
	    
	        if saft <> True or saftOK=true then
	            'ega 19/06/2008 likes
	            rst.Open "select atp,total_factura from facturas_cli fc with(nolock) inner join clientes cl with(nolock) on cl.ncliente like '"&session("ncliente")&"%' and cl.ncliente=fc.ncliente where nfactura='"&p_nfactura&"' and fc.nfactura like '"&session("ncliente")&"%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                atp=cbool(rst("atp"))
                totalFactura=rst("total_factura")
	            rst.Close
	            if si_tiene_modulo_EBESA <> 0 and atp=true then
	                rst.Open "select importe_dto_factura,dto_factura from configuracion with(nolock) where nempresa='"&session("ncliente")&"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	                if totalFactura>rst("importe_dto_factura") then
	                    rstBlq.Open "update detalles_fac_cli with(rowlock) set descuento=" & rst("dto_factura") & " where nfactura='" & p_nfactura & "'",session("dsn_cliente")
	                    'divfactura = d_lookup("divisa","facturas_cli","nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente"))

                        divfacturaSelect="select divisa from facturas_cli with (nolock) where nfactura like ?+'%' and nfactura= ?"
                        divfactura=DlookupP2(divfacturaSelect, session("ncliente")&"", adVarchar, 20, p_nfactura&"", adVarchar, 20, session("dsn_cliente"))

                        'ndec = d_lookup("ndecimales", "divisas", "codigo='" & divfactura & "' and codigo like '" & session("ncliente") & "%'", session("dsn_cliente"))

                        ndecSelect="select ndecimales from divisas with (nolock) where codigo= ? and codigo like ?+'%'"
                        ndec=DlookupP2(ndecSelect, divfactura&"", adVarchar, 15, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))

                        'ega 19/06/2008 obtengo solamente los campos necesarios de detalles_fac_cli
	                    rstSelect.Open "select pvp,descuento,descuento2,descuento3,cantidad,item from detalles_fac_cli with(nolock) where nfactura='" & p_nfactura & "' order by item",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	                    while not rstSelect.EOF
    	                    temp=rstSelect("pvp")
    	                    Dtotemp=(temp*rstSelect("descuento"))/100
		                    temp=cdbl(temp)-cdbl(Dtotemp)
		                    Dtotemp=(temp*rstSelect("descuento2"))/100
		                    temp=cdbl(temp)-cdbl(Dtotemp)
    		                Dtotemp=(temp*rstSelect("descuento3"))/100
	    	                temp=cdbl(temp)-cdbl(Dtotemp)
	                        rstAux2.Open "update detalles_fac_cli with(rowlock) set importe = " & replace(cstr(miround((temp*rstSelect("cantidad")),ndec)),",",".") & " where nfactura='" & p_nfactura & "' and item =" & rstSelect("item"),session("dsn_cliente")
	                        rstSelect.MoveNext
	                    wend
	                    rstSelect.Close
	                    Precios p_nfactura
	                end if
	                rst.Close
	            end if
		        rstBlq.open "update facturas_cli with(rowlock) set ahora=1 where nfactura = '" & p_nfactura & "'",session("dsn_cliente")
		        'ega 19/06/2008 obtengo solo los campos necesarios de facturas_cli
		        rstBlq.Open "select ncliente from facturas_cli with(nolock) where nfactura =  '" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		        auditar_ins_bor session("usuario"),p_nfactura,rstBlq("ncliente"),"bloqueo","","","facturas_cli"
		        rstBlq.Close
		    end if
		    viene=limpiaCadena(request.querystring("viene"))
	        if viene="" then viene=limpiaCadena(request.form("viene"))
		
		    if viene = "search" then
		        mode="search"
		    else
		        mode="browse"
		    end if
	    end if
	    if mode="desbloqueo" then
	        saftOK=false
	        desbloquearSAFT=""
	        if saft = True then
	            factSig = obtenerFacturaSiguiente (p_nfactura)
	            'ultimaFact = d_lookup("campo20","facturas_cli","nfactura = '" & factSig&"'",session("dsn_cliente"))

                ultimaFactSelect="select campo20 from facturas_cli with(nolock) where nfactura=?"
                ultimaFact=DlookupP1(ultimaFactSelect,factSig&"", adVarchar, 20,session("dsn_cliente"))

	            ultimaFact=ultimaFact&""
	            if ultimaFact = "" then
	                saftOK=true
	                desbloquearSAFT=",campo20 = null, campo19=null "
	            else%>
	                <script language="javascript" type="text/javascript">
                                     window.alert("<%=LitUltimaFact%>");
			        </script> 
	            <%end if
	        end if
	
	        if saft <> True or saftOK=true  then
	        'utilizamos la variable desbloquearSAFT para añadir los campos a modificar en caso de que quereamos desbloquear una factura de saft
		        rstBlq.open "update facturas_cli with(rowlock) set ahora=0 "& desbloquearSAFT &" where nfactura = '" & p_nfactura & "'",session("dsn_cliente")
		        'ega 19/06/2008 obtengo solo los campos necesarios de facturas_cli
		        rstBlq.Open "select ncliente from facturas_cli with(nolock) where nfactura =  '" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		        auditar_ins_bor session("usuario"),p_nfactura,rstBlq("ncliente"),"desbloqueo","","","facturas_cli"
		        rstBlq.Close
		    end if
		
		    viene=limpiaCadena(request.querystring("viene"))
	        if viene="" then viene=limpiaCadena(request.form("viene"))
		
		    if viene ="search" then
		        mode="search"
		    else
		        mode="browse"
		    end if
	    end if

        'cag
	    if Request.QueryString("h_ahora")>"" then
		    h_ahora=limpiaCadena(Request.QueryString("h_ahora"))
	    else
		    h_ahora=limpiaCadena(Request.form("h_ahora"))
	    end if
        'fin cag

        ' FLM : 19/01/2009 : Añadir captura de ncliente por queryString
        if Request.QueryString("ncliente")>"" then
	        TraerCliente=limpiaCadena(Request.QueryString("ncliente"))
	    end if
	    p_ncliente = limpiaCadena(Request.Form("ncliente"))

	    if p_ncliente="" then
		    p_ncliente=limpiaCadena(request.form("h_ncliente"))
	    end if

	    if Request.QueryString("fecha")>"" then
		    p_fecha=limpiaCadena(Request.QueryString("fecha"))
	    else
		    p_fecha=limpiaCadena(Request.form("fecha"))
	    end if
	
	    if Request.QueryString("nfolio")>"" then
		    nfolio=limpiaCadena(Request.QueryString("nfolio"))
	    else
		    nfolio=limpiaCadena(Request.form("nfolio"))
	    end if
	
	    if Request.QueryString("cobrada")>"" then
		    p_cobrada=limpiaCadena(Request.QueryString("cobrada"))
	    else
		    p_cobrada=limpiaCadena(Request.form("cobrada"))
	    end if
	
	    if Request.QueryString("contabilizada")>"" then
		    p_contabilizada=limpiaCadena(Request.QueryString("contabilizada"))
	    else
		    p_contabilizada=limpiaCadena(Request.form("contabilizada"))
	    end if

	    if request.querystring("contacto")>"" then
		    tmp_contacto=enc.EncodeForJavascript(limpiaCadena(request.querystring("contacto")))
	    else
		    tmp_contacto=enc.EncodeForJavascript(limpiaCadena(request.form("contacto")))
	    end if

	    if request.querystring("banco")>"" then
		    tmp_banco=limpiaCadena(request.querystring("banco"))
	    else
		    tmp_banco=limpiaCadena(request.form("banco"))
	    end if

	    if request.querystring("ncuenta1")>"" then
		    tmp_ncuenta1=limpiaCadena(request.querystring("ncuenta1"))
	    else
		    tmp_ncuenta1=limpiaCadena(request.form("ncuenta1"))
	    end if

	    if request.querystring("ncuenta2")>"" then
		    tmp_ncuenta2=limpiaCadena(request.querystring("ncuenta2"))
	    else
		    tmp_ncuenta2=limpiaCadena(request.form("ncuenta2"))
	    end if

	    if request.querystring("ncuenta3")>"" then
		    tmp_ncuenta3=limpiaCadena(request.querystring("ncuenta3"))
	    else
		    tmp_ncuenta3=limpiaCadena(request.form("ncuenta3"))
	    end if

	    if request.querystring("ncuenta4")>"" then
		    tmp_ncuenta4=limpiaCadena(request.querystring("ncuenta4"))
	    else
		    tmp_ncuenta4=limpiaCadena(request.form("ncuenta4"))
	    end if

	    if request.querystring("ncuenta5")>"" then
		    tmp_ncuenta5=limpiaCadena(request.querystring("ncuenta5"))
	    else
		    tmp_ncuenta5=limpiaCadena(request.form("ncuenta5"))
	    end if

	    if request.querystring("ncuenta6")>"" then
		    tmp_ncuenta6=limpiaCadena(request.querystring("ncuenta6"))
	    else
		    tmp_ncuenta6=limpiaCadena(request.form("ncuenta6"))
	    end if

	    if request.querystring("documento")>"" then
		    tmp_documento=limpiaCadena(request.querystring("documento"))
	    else
		    tmp_documento=limpiaCadena(request.form("documento"))
	    end if

	    if modp="" then modp=limpiaCadena(request.querystring("modp"))
	    if modp="" then modp=limpiaCadena(request.form("modp"))
	    if modd="" then modd=limpiaCadena(request.querystring("modd"))
	    if modd="" then modd=limpiaCadena(request.form("modd"))
	    if modi="" then modi=limpiaCadena(request.querystring("modi"))
	    if modi="" then modi=limpiaCadena(request.form("modi"))
	    if pciva="" then pciva=limpiaCadena(request.querystring("pciva"))
	    if pciva="" then pciva=limpiaCadena(request.form("pciva"))
	    if modn="" then modn=limpiaCadena(request.querystring("modn"))
	    if modn="" then modn=limpiaCadena(request.form("modn"))
	    if rnf="" then rnf=limpiaCadena(request.querystring("rnf"))
	    if rnf="" then rnf=limpiaCadena(request.form("rnf"))
	    if pnf="" then pnf=limpiaCadena(request.querystring("pnf"))
	    if pnf="" then pnf=limpiaCadena(request.form("pnf"))

	    viene=limpiaCadena(request.querystring("viene"))
	    if viene="" then viene=limpiaCadena(request.form("viene"))

	    if request.querystring("cv") & "">"" then
		    cv=limpiaCadena(request.querystring("cv"))
	    elseif request.form("cv") & "">"" then
		    cv=limpiaCadena(request.form("cv"))
	    end if

	    if request.querystring("caju") & "">"" then
		    caju=limpiaCadena(request.querystring("caju"))
	    elseif request.form("caju") & "">"" then
		    caju=limpiaCadena(request.form("caju"))
	    end if

	    if request.QueryString("novei")& "">"" then
		    novei=limpiaCadena(request.QueryString("novei"))
	    elseif request.form("novei") & "">"" then
		    novei=limpiaCadena(request.form("novei"))
	    end if

	    if request.QueryString("bcc") & "">"" then
		    bcc=limpiaCadena(request.QueryString("bcc"))
	    elseif request.form("bcc") & "">"" then
		    bcc=limpiaCadena(request.form("bcc"))
	    end if
	    if cstr(bcc & "")="1" then
		    texto_bcc=" " & "disabled"
	    else
		    texto_bcc=""
	    end if

	    if request.QueryString("ocb") & "">"" then
		    ocb=limpiaCadena(request.QueryString("ocb"))
	    elseif request.form("ocb") & "">"" then
		    ocb=limpiaCadena(request.form("ocb"))
	    end if

	    ''ricardo 7-12-2004 parametro series unicas al documento por usuario
        if s&""="" then
	        s=limpiaCadena(request.querystring("s"))
	        if s="" then s=limpiaCadena(request.form("s"))
        end if
	    s=preparar_lista(s)

	    if request.querystring("comercial_ant")>"" then
		    comercial_antR=limpiaCadena(request.querystring("comercial_ant"))
	    else
		    comercial_antR=limpiaCadena(request.form("comercial_ant"))
	    end if
	    if request.querystring("old_forma_pago")>"" then
		    old_forma_pagoR=limpiaCadena(request.querystring("old_forma_pago"))
	    else
		    old_forma_pagoR=limpiaCadena(request.form("old_forma_pago"))
	    end if
	    if request.querystring("cambiarcom")>"" then
		    cambiarcomR=limpiaCadena(request.querystring("cambiarcom"))
	    else
		    cambiarcomR=limpiaCadena(request.form("cambiarcom"))
	    end if
	    if cambiarcomR & ""="" then cambiarcomR="0"
	    cambiarcomR=cstr(cambiarcomR)

	    '**RGU 26/4/2006
	    if nmc="" then nmc=limpiaCadena(request.querystring("nmc"))
	    if nmc="" then nmc=limpiaCadena(request.form("nmc"))

        campo    = limpiaCadena(request.QueryString("campo"))
	    if campo & ""="" then
	        campo = Request.Form("campo")
        end if
	    texto    = limpiaCadena(request.QueryString("texto"))
	    if texto & ""="" then
	        texto = Request.Form("texto")
        end if
	
	    lote=limpiaCadena(Request.QueryString("lote"))
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

	    modo_orig=request.form("modo_orig")
	    if modo_orig&""="" then modo_orig=request.QueryString("mode")
	    p_pagsl=0

	    if modo_orig<>"browse" or (pagsl&""<>"1" and modo_orig="browse") then
	        p_pagsl=1
	    end if

        if (mode="save" or mode="first_save" or mode="edit") then
            %><input type="hidden" name="modo_orig" value="<%=enc.EncodeForHtmlAttribute(mode)%>" /><%
        else
           if mode="browse" then
                %><input type="hidden" name="modo_orig" value="<%=enc.EncodeForHtmlAttribute(modo_orig)%>" /><%
           end if
        end if
        '**rgu

        No_act_frabono=1

	    '   GPD (27/04/2007).
	    bolActivoValesDTO = false
	    'strSerie = d_lookup("serie", "facturas_cli", "nfactura like '" & p_nfactura & "'", session("dsn_cliente"))

        strSerieSelect="select serie from facturas_cli with(nolock) where nfactura like ?"
        strSerie=DlookupP1(strSerieSelect, p_nfactura&"", adVarchar, 20, session("dsn_cliente"))

        'strCif = d_lookup("empresa", "series", "nserie like '" & strSerie & "'", session("dsn_cliente"))

        strCifSelect="select empresa from series with(nolock) where nserie like ?"
        strCif=DlookupP1(strCifSelect, strSerie&"", adVarchar, 10, session("dsn_cliente"))

        'bolActivoValesDTO = d_lookup("ACTIVARVALEDTO", "EMPRESAS", "CIF like '" & strCif & "'", session("dsn_cliente"))
                    
        bolActivoValesDTOSelect="select ACTIVARVALEDTO from EMPRESAS with(nolock) where CIF like ?"
        bolActivoValesDTO=DlookupP1(bolActivoValesDTOSelect, strCif&"", adVarchar, 25, session("dsn_cliente"))
                    
                    %>

	    <input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
	    <input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
	    <input type="hidden" name="modp" value="<%=enc.EncodeForHtmlAttribute(modp)%>"/>
	    <input type="hidden" name="modd" value="<%=enc.EncodeForHtmlAttribute(modd)%>"/>
	    <input type="hidden" name="modi" value="<%=enc.EncodeForHtmlAttribute(modi)%>"/>
	    <input type="hidden" name="cv" value="<%=enc.EncodeForHtmlAttribute(cv)%>"/>
	    <input type="hidden" name="caju" value="<%=enc.EncodeForHtmlAttribute(caju)%>"/>
	    <input type="hidden" name="novei" value="<%=enc.EncodeForHtmlAttribute(novei)%>"/>
	    <input type="hidden" name="bcc" value="<%=enc.EncodeForHtmlAttribute(bcc)%>"/>
	    <input type="hidden" name="ocb" value="<%=enc.EncodeForHtmlAttribute(ocb)%>"/>
	    <input type="hidden" name="s" value="<%=enc.EncodeForHtmlAttribute(s)%>"/>
	    <input type="hidden" name="pciva" value="<%=enc.EncodeForHtmlAttribute(pciva)%>"/>
	    <input type="hidden" name="modn" value="<%=enc.EncodeForHtmlAttribute(modn)%>"/>
	    <input type="hidden" name="tfi" value="<%=enc.EncodeForHtmlAttribute(tfi)%>"/>
	    <input type="hidden" name="cpc" value="<%=enc.EncodeForHtmlAttribute(cpc)%>"/>
	    <input type="hidden" name="nmc" value="<%=enc.EncodeForHtmlAttribute(nmc)%>"/>
	    <input type="hidden" name="rnf" value="<%=enc.EncodeForHtmlAttribute(rnf)%>"/>
	    <input type="hidden" name="pnf" value="<%=enc.EncodeForHtmlAttribute(pnf)%>"/>
	    <input type="hidden" name="p_pagsl" value="<%=enc.EncodeForHtmlAttribute(p_pagsl)%>"/>
	    <input type="hidden" name="campo" value="<%=enc.EncodeForHtmlAttribute(campo)%>"/>
	    <input type="hidden" name="texto" value="<%=enc.EncodeForHtmlAttribute(texto)%>"/>
	    <input type="hidden" name="lote" value="<%=enc.EncodeForHtmlAttribute(lote)%>"/>
	    <input type="hidden" name="criterio" value="<%=enc.EncodeForHtmlAttribute(criterio)%>"/>

	    <%''ricardo 7-12-2004 si la serie del documento que tratamos de ver no esta entre las permitidas al usuario no nos dejara verla
	    if p_nfactura & "">"" then
		    if comprobar_LS(s,mode,p_nfactura,"FACTURAS_CLI")=0 then
			    %><script language="javascript" type="text/javascript">
                        alert("<%=LitMsgDocNoPermAcc%>");
                        document.facturas_cli.action = "facturas_cli.asp?nfactura=&mode=add" + "&almacenSerie=<%=enc.EncodeForJavascript(almacenSerie) %>&almacenTPV=<%=enc.EncodeForJavascript(almacenTPV) %>";
                        document.facturas_cli.submit();
                        parent.botones.document.location = "facturas_cli_bt.asp?mode=add";
			    </script><%
			    CerrarTodo()
			    response.end
		    end if
	    end if
	    'EBF 3/1/2006 comprobamos si el usuario tiene activada en la tabla de configuración el campo de aplicar dtos cliente a artículos
	    dto_cli_art=d_lookup("DTO_CLI_ART","configuracion","nempresa='"&session("ncliente")&"'",session("dsn_cliente"))%>
	    <input type="hidden" name="dto_cli_art" value="<%=enc.EncodeForHtmlAttribute(dto_cli_art)%>"/>
	    <%gen_vencimiento=nz_b(d_lookup("gen_vencimientos","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))%>
	    <input type="hidden" name="gen_vencimiento" value="<%=enc.EncodeForHtmlAttribute(gen_vencimiento)%>"/>

	    <%'este siguiente hidden es para saber si al cambiar el comercial ,tenemos que cambiar los vencimientos
	    if p_nfactura & "">"" then
		    comercial_ven=0
		    rst.cursorlocation=3
		    'ega 19/06/2008 obtengo solamente los campos necesarios y with(nolock)
		    rst.open "select comercial from vencimientos_salida with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente")
		    if not rst.eof then
			    si_vencimientos=1
			    if rst("comercial") & ""="" then
				    comercial_ven=1
			    end if
		    else
			    si_vencimientos=0
			    comercial_ven=0
		    end if
		    rst.close%>
		    <input type="hidden" name="si_vencimientos" value="<%=enc.EncodeForHtmlAttribute(si_vencimientos)%>"/>
		    <input type="hidden" name="comercial_ven" value="<%=enc.EncodeForHtmlAttribute(comercial_ven)%>"/>
	    <%else
		    si_vencimientos=0%>
		    <input type="hidden" name="si_vencimientos" value="<%=enc.EncodeForHtmlAttribute(si_vencimientos)%>"/>
	    <%end if

	    if request.querystring("cobrada")>"" then
		    tmp_cobrada=limpiaCadena(request.querystring("cobrada"))
	    else
		    tmp_cobrada=limpiaCadena(request.form("cobrada"))
	    end if
	
	    if request.querystring("portes")>"" then
		    tmp_portes=limpiaCadena(request.querystring("portes"))
	    else
		    tmp_portes=limpiaCadena(request.form("portes"))
	    end if

	    if request.querystring("observaciones")>"" then
		    tmp_observaciones=enc.EncodeForJavascript(limpiaCadena(request.querystring("observaciones")))
	    else
		    tmp_observaciones=enc.EncodeForJavascript(limpiaCadena(request.form("observaciones")))
	    end if

	    if request.querystring("notas")>"" then
		    tmp_notas=enc.EncodeForJavascript(limpiaCadena(request.querystring("notas")))
	    else
		    tmp_notas=enc.EncodeForJavascript(limpiaCadena(request.form("notas")))
	    end if

	    if request.querystring("h_ncliente")>"" then
		    tmp_ncliente=limpiaCadena(request.querystring("h_ncliente"))
	    else
		    tmp_ncliente=limpiaCadena(request.form("h_ncliente"))
	    end if

	    if request.querystring("tipo_pago")>"" then
		    tmp_tipo_pago=limpiaCadena(request.querystring("tipo_pago"))
	    else
		    tmp_tipo_pago=limpiaCadena(request.form("tipo_pago"))
	    end if

	    if request.querystring("forma_pago")>"" then
		    tmp_forma_pago=limpiaCadena(request.querystring("forma_pago"))
	    else
		    tmp_forma_pago=limpiaCadena(request.form("forma_pago"))
	    end if

	    if request.querystring("comercial")>"" then
		    tmp_comercial=limpiaCadena(request.querystring("comercial"))
	    else
		    tmp_comercial=limpiaCadena(request.form("comercial"))
	    end if

	    if request.querystring("agenteasignado")>"" then
		    tmp_agenteasignado=limpiaCadena(request.querystring("agenteasignado"))
	    else
		    tmp_agenteasignado=limpiaCadena(request.form("agenteasignado"))
	    end if

        if request.querystring("cod_proyecto")>"" then
		    tmp_cod_proyecto=limpiaCadena(request.querystring("cod_proyecto"))
	    else
		    tmp_cod_proyecto=limpiaCadena(request.form("cod_proyecto"))
	    end if

	    if request.querystring("tarifa")>"" then
		    tmp_tarifa=limpiaCadena(request.querystring("tarifa"))
	    else
		    tmp_tarifa=limpiaCadena(request.form("tarifa"))
	    end if

	    if request.querystring("fechaenvio")>"" then
		    tmp_fechaenvio=limpiaCadena(request.querystring("fechaenvio"))
	    else
		    tmp_fechaenvio=limpiaCadena(request.form("fechaenvio"))
	    end if

	    if request.querystring("fechapedido")>"" then
		    tmp_fechapedido=limpiaCadena(request.querystring("fechapedido"))
	    else
		    tmp_fechapedido=limpiaCadena(request.form("fechapedido"))
	    end if

	    if request.querystring("transportista")>"" then
		    tmp_transportista=enc.EncodeForJavascript(limpiaCadena(request.querystring("transportista")))
	    else
		    tmp_transportista=enc.EncodeForJavascript(limpiaCadena(request.form("transportista")))
	    end if

	    if request.querystring("nombre")>"" then
		    tmp_nombre=enc.EncodeForJavascript(limpiaCadena(request.querystring("nombre")))
	    else
		    tmp_nombre=enc.EncodeForJavascript(limpiaCadena(request.form("nombre")))
	    end if

	    if request.querystring("nenvio")>"" then
		    tmp_nenvio=enc.EncodeForJavascript(limpiaCadena(request.querystring("nenvio")))
	    else
		    tmp_nenvio=enc.EncodeForJavascript(limpiaCadena(request.form("nenvio")))
	    end if

	    if request.querystring("submode")>"" then
		    submode=request.querystring("submode")
	    else
		    submode=request.form("submode")
	    end if

	    if request.querystring("incoterms")>"" then
		    tmp_incoterms=limpiaCadena(request.querystring("incoterms"))
	    else
		    tmp_incoterms=limpiaCadena(request.form("incoterms"))
	    end if

	    if request.querystring("fob")>"" then
		    tmp_fob=enc.EncodeForJavascript(limpiaCadena(request.querystring("fob")))
	    else
		    tmp_fob=enc.EncodeForJavascript(limpiaCadena(request.form("fob")))
	    end if
	
	    '*** AMP 	 
	    if Request.QueryString("divisafc")>"" then
		    tmpdivisafc=limpiaCadena(Request.QueryString("divisafc"))
	    elseif Request.form("divisafc")>"" then
		    tmpdivisafc=limpiaCadena(Request.form("divisafc"))
	    end if	

	    p_serie=limpiaCadena(Request.QueryString("serie"))
	    if p_serie & ""="" then
		    p_serie=limpiaCadena(Request.form("serie"))
	    end if
	    if p_serie="" and mode="add" then
		    'Obtener la serie por defecto
		    p_serie=ObtenerSerieTienda("FACTURA A CLIENTE")
		    if p_serie & ""="" then
			    strwhere=" and pordefecto=1 "
		    else
			    strwhere=" and nserie='" & p_serie & "'"
		    end if
	    else
		    strwhere=" and nserie='" & p_serie & "'"
	    end if

	    'mmg:calculamos el almacen por defecto de la serie
	    'ega 19/06/2008 like en la tabla empresas, y le pongo top 1 porque solo voy a utilizar el primer registro
	    '' MPC 29/08/2008 Se ha actualizado la select para que filtre por el tipo de documento FACTURA A CLIENTE
        rstAux.cursorlocation=3
	    rstAux.open "select top 1 nserie,irpf,almacen from series with(nolock) left outer join empresas with(nolock) on empresa=cif and cif like '" & session("ncliente") & "%' where nserie like '" & session("ncliente") & "%' and tipo_documento='FACTURA A CLIENTE' " & strwhere,session("dsn_cliente")
	    if not rstAux.EOF then
		    p_serie=rstAux("nserie")
		    tmp_irpf=rstAux("irpf")
		    almacenSerie= rstAux("almacen")
	    else
		    almacenSerie= ""
	    end if

	    rstAux.close
	
	    gestbono=request.Form("gestbono")&""
	    if gestbono="" then 
	        'gestbono=nz_b2(d_lookup("gestbono", "configuracion", "nempresa='"&session("ncliente")&"'", session("dsn_cliente")))

            gestbonoSelect="select gestbono from configuracion with(nolock) where nempresa=?"
            gestbono= nz_b2(DlookupP1(gestbonoSelect, session("ncliente")&"", adChar, 5, session("dsn_cliente")))


	    end if
	
	    ''ricardo 5-9-2005 al añadir mas de 10 campos personalizables, se dimensionara la lista_valores al numero de campos existentes
	    num_campos_ventas=10
	    rstAux.cursorlocation=3
	    ''ricardo 27-3-2009 se contara de otra manera mas efectiva
	    rstAux.open "select max(convert(int,substring(ncampo,6,len(ncampo)))) as contador from camposperso with(nolock) where tabla='DOCUMENTOS VENTA' and ncampo like '" & session("ncliente") & "%' and isnull(titulo,'')<>'' ",session("dsn_cliente")
	    if not rstAux.eof and not isnull(rstAux("contador")) then
		    num_campos_ventas=rstAux("contador")
	    else
	        ''response.Write("he entrado2<BR>")
		    num_campos_ventas=10
	    end if
	    rstAux.close

	    ''ricardo 25-1-2007
	    if num_campos_ventas & ""="" then
		    num_campos_ventas=10
	    end if

        ''ricardo 27-3-2009 se guardaran tantos campos como existan en tabla y con titulo sera con titutlo, por si utilizamos internamente para alguna cosa, pero claro el usuario no puede tocarla
	    num_campos_clientes=10
	    rstAux.cursorlocation=3
	    ''ricardo 27-3-2009 se contara de otra manera mas efectiva
	    ''rstAux.open "select count(*) as contador from camposperso with(nolock) where tabla='CLIENTES' and ncampo like '" & session("ncliente") & "%'  and isnull(titulo,'')<>'' ",session("dsn_cliente")
	    rstAux.open "select max(convert(int,substring(ncampo,6,len(ncampo)))) as contador from camposperso with(nolock) where tabla='CLIENTES' and ncampo like '" & session("ncliente") & "%' and isnull(titulo,'')<>'' ",session("dsn_cliente")
	    if not rstAux.eof then
		    num_campos_clientes=rstAux("contador")
	    else
		    num_campos_clientes=10
	    end if
	    rstAux.close

		'JMA 25/11/04. Copiar campos personalizables de los clientes'
		if num_campos_ventas>=num_campos_clientes then
		    campos_a_dimensionar=num_campos_ventas
		else
		    campos_a_dimensionar=num_campos_clientes
		end if
		if cstr(campos_a_dimensionar & "")="" then campos_a_dimensionar=10
		redim tmp_lista_valores(campos_a_dimensionar+2)
		for ki=1 to campos_a_dimensionar
			tmp_lista_valores(ki)=""
		next
		'JMA 25/11/04. FIN Copiar campos personalizables de los clientes'

	    ''JMA 28-10-2004 si existen campos personalizables con titulo no nulo saldrán los campos personalizables
	    si_campo_personalizables=0
	    rst.cursorlocation=3
	    rst.open "select ncampo from camposperso with(nolock) where tabla='DOCUMENTOS VENTA' and titulo is not null and titulo<>'' and ncampo like '" & session("ncliente") & "%'",session("dsn_cliente")
	    if not rst.eof then
		    si_campo_personalizables=1
	    else
		    si_campo_personalizables=0
	    end if
	    rst.close
	    %><input type="hidden" name="si_campo_personalizables" value="<%=enc.EncodeForHtmlAttribute(si_campo_personalizables)%>"/><%
	    ''JMA 28-10-2004 FIN si existen campos personalizables con titulo no nulo saldrán los campos personalizables	
	    ''ricardo 13-10-2009 se obtiene, si esta pantalla tiene algun limite en el numero de facturas creadas
	    if mode="add" then
		
		    limiteFacturasCreadas=0
		    rst.cursorlocation=3
		    rst.open "exec limitesPagina '" & session("ncliente") & "','facturas_cli.asp'",dsnilion
		    if not rst.eof then
		        if rst("limite")& ""="" or isnumeric(rst("limite"))=0 then
		            limiteFacturasCreadas=0
		        else
		            limiteFacturasCreadas=rst("limite")
		        end if
		    end if
		    rst.close
            %><input type="hidden" name="limiteFacturasCreadas" value="<%=enc.EncodeForHtmlAttribute(limiteFacturasCreadas)%>"/><%
            ''ahora se calculan cuantas facturas hay generadas
            CantidadFacturasCreadas=0
            rst.cursorlocation=3
            StrContFacAnyo="select count(*) as contador from facturas_cli with(NOLOCK) where nfactura like '" & session("ncliente") & "%' and fecha>='1/1/" & year(now) & "' and fecha<='31/12/" & year(now) & "'"
            rst.open StrContFacAnyo,session("dsn_cliente")
            if not rst.EOF then
                CantidadFacturasCreadas=rst("contador")
            end if
            rst.close
            %><input type="hidden" name="CantidadFacturasCreadas" value="<%=enc.EncodeForHtmlAttribute(CantidadFacturasCreadas)%>"/><%
	    end if

	    ''JMA 28-10-2004 añadir campos personalizables a facturas_cli
	    ''ricardo 11-1-2007 se añade el mode delete, ya que si la factura tiene un apunte de caja da error, por no haberse hecho el redim de lista_valores
	    if mode="browse" or mode="edit" or mode="add" or mode="save" or mode="first_save" or mode="delete" then
		    num_campos=0
		    if mode="add" then
		        if cstr(num_campos_ventas & "")="" then num_campos_ventas=10
			    redim lista_valores(num_campos_ventas+2)
			    for ki=1 to num_campos_ventas+2
				    lista_valores(ki)=""
			    next
			    num_campos=num_campos_ventas
		    else

                ''ricardo 27-3-2009 se cambia el select , para que salgan todos los campos perso
                if cstr(num_campos_ventas & "")="" then num_campos_ventas=10
                strGuarFacCli="select p.nfactura,"
                strGuarFacCliAux=""
                for ki_cli=1 to num_campos_ventas
                    cadena_campo="p.campo" & replace(space(2-len(cstr(ki_cli)))," ","0") & cstr(ki_cli)
                    strGuarFacCliAux=strGuarFacCliAux & cadena_campo & ","
                next
                strGuarFacCli=strGuarFacCli & mid(strGuarFacCliAux,1,len(strGuarFacCliAux)-1)
                strGuarFacCli=strGuarFacCli & " from facturas_cli as p with(nolock) where p.nfactura='" & p_nfactura & "'"
			    rstAux2.cursorlocation=3
			    rstAux2.open strGuarFacCli,session("dsn_cliente")
			    if not rstAux2.eof then
                    ''ricardo 27-3-2009 se guardaran tantos campos como existan en tabla y con titulo
                    ''sera con titutlo, por si utilizamos internamente para alguna cosa, pero claro el usuario no puede tocarla

                    if cstr(num_campos_ventas & "")="" then num_campos_ventas=10
                    redim lista_valores(num_campos_ventas+2)
                    ''ricardo 27-3-2009 en lugar de poner una linea por campo, se pone un for que lo hago por la totalidad
                    for ki_cli=1 to num_campos_ventas
                        cadena_campo="campo" & replace(space(2-len(cstr(ki_cli)))," ","0") & cstr(ki_cli)
                        lista_valores(ki_cli)=Nulear(rstAux2(cadena_campo))
                    next
				    num_campos=num_campos_ventas
			    else
				    redim lista_valores(num_campos_ventas+2)
				    for ki=1 to num_campos_ventas+2
					    lista_valores(ki)=""
				    next
				    num_campos=num_campos_ventas
			    end if
			    rstAux2.close
		    end if
	    end if
	    ''JMA 28-10-2004 añadir campos personalizables a facturas_cli

	    if request.querystring("cli")>"" then
	        TraerCliente = session("ncliente") & Completar(limpiaCadena(request.querystring("cli")),5,"0")
	    elseif mode="add" and TraerCliente="" then
		    'Obtener el cliente de la serie por defecto.
		    'TraerCliente=d_lookup("cliente","series","nserie like '" & session("ncliente") & "%' and nserie='" & p_serie & "'",session("dsn_cliente"))

            TraerClienteSelect="select cliente from series with(nolock) where nserie like ?+'%' and nserie=?"
            TraerCliente=DlookupP2(TraerClienteSelect, session("ncliente")&"", adVarchar, 10, p_serie&"", adVarchar, 10,session("dsn_cliente"))

	    end if

	    if (mode="add" or mode="edit") and TraerCliente<>"" then
		    rstAux.open "select fbaja,aviso from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & Completar(TraerCliente,5,"0") & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		    if not rstAux.eof then
			    if rstAux("aviso")>"" then
				    mensaje=reemplazar(reemplazar(reemplazar(rstAux("aviso"),chr(10),""),chr(13),"\n"),chr(34),"\"&chr(34))
			    else
				    mensaje=""
			    end if
			    if rstAux("fbaja")>"" then
				    mensaje=LitClienteDadoBaja & "\n" & mensaje
				    TraerCliente=""%>
				    <script language="javascript" type="text/javascript">				        alert("<%=LitAvisos%>:\n\n<%=mensaje%>");</script>
			    <%else
				    if rstAux("aviso")>"" then%>
					    <script language="javascript" type="text/javascript">					        alert("<%=LitAvisos%>:\n\n<%=mensaje%>");</script>
				    <%end if
			    end if
		    end if
		    rstAux.close
	    end if

	    'Captura de datos del cliente que se está introduciendo en el pedido'
	    if TraerCliente > "" then
		    TraerCliente=Completar(TraerCliente,5,"0")
		    Error="NO"
		    'ega 19/06/2008 seleccionar los campos necesarios
            ''ricardo 27-3-2009 se guardaran tantos campos como existan en tabla y con titulo
            ''sera con titutlo, por si utilizamos internamente para alguna cosa, pero claro el usuario no puede tocarla
		        campos_a_dimensionar=num_campos_clientes

		    if cstr(campos_a_dimensionar & "")="" then campos_a_dimensionar=10
		    'jcg 20/01/2009: añadido el parametro proyecto
            ''ricardo 27-3-2009 se cambia el select , para que salgan todos los campos perso
            strTrarCli="select RSOCIAL,TARIFA,DTO,FPAGO,TPAGO,RECARGO,COMERCIAL,TRANSPORTISTA,PORTES,DTO2,BANCO,NCUENTA,DOMREC,DIVISA,DTO3,AGENTE"
            for ki_cli=1 to campos_a_dimensionar
                cadena_campo="campo" & replace(space(2-len(cstr(ki_cli)))," ","0") & cstr(ki_cli)
                strTrarCli=strTrarCli & "," & cadena_campo
            next
            strTrarCli=strTrarCli & ",proyecto from clientes with(nolock) where ncliente='" & TraerCliente & "'"
            rstAux.cursorlocation=3
		    rstAux.open strTrarCli,session("dsn_cliente")

		    if not rstAux.EOF then
		        tmp_ncliente=TraerCliente
			    tmp_nombre=rstAux("rsocial")
			    tmp_forma_pago=rstAux("fpago")
			    tmp_tipo_pago=rstAux("tpago")
			    tmp_transportista=rstAux("transportista")
			    tmp_portes=rstAux("portes")
			    tmp_divisa=rstAux("divisa")
		        tmp_cod_proyecto=rstAux("proyecto")

			    'EBF 12/2/2005 guarda los descuentos de cabecera en funcion de si hemos seleccionado aplicar descuentos de cliente a documentos de venta
			    tmp_dto1=0
			    tmp_dto2=0
			    tmp_dto3=0
			    if dto_cli_art=true then
				    dto1_cli=rstAux("dto")
				    dto2_cli=rstAux("dto2")
				    dto3_cli=rstAux("dto3")
			    else
				    tmp_dto1=rstAux("dto")
				    tmp_dto2=rstAux("dto2")
				    tmp_dto3=rstAux("dto3")
			    end if
			    tmp_rf=rstAux("recargo")
			    tmp_tarifa=rstAux("tarifa")
			    tmp_comercial=rstAux("comercial")
			    tmp_agenteasignado=rstAux("agente")
			    if rstAux("domrec")=true then
				    tmp_banco=trim(rstAux("banco"))
                    'if len(rstAux("ncuenta"))=24 then
				        tmp_ncuenta1= Mid(rstAux("ncuenta"), 1, 2)
				        tmp_ncuenta2= Mid(rstAux("ncuenta"), 3, 2)
				        tmp_ncuenta3= Mid(rstAux("ncuenta"), 5, 4)
	   			        tmp_ncuenta4= Mid(rstAux("ncuenta"), 9, 4)
                        tmp_ncuenta5= Mid(rstAux("ncuenta"), 13, 2)
                        tmp_ncuenta6= Mid(rstAux("ncuenta"), 15, Len(rstAux("ncuenta"))-14)
                    'else
                        'tmp_ncuenta1= ""
				        'tmp_ncuenta2= ""
                        'tmp_ncuenta3= Mid(rstAux("ncuenta"), 1, 4)
	   			        'tmp_ncuenta4= Mid(rstAux("ncuenta"), 5, 4)
                        'tmp_ncuenta5= Mid(rstAux("ncuenta"), 9, 2)
                        'tmp_ncuenta6= Mid(rstAux("ncuenta"), 11, 10)
                    'end if
			    else
				    tmp_banco=""
				    tmp_ncuenta= ""
                    tmp_ncuenta1= ""
				    tmp_ncuenta2= ""
				    tmp_ncuenta3= ""
	   			    tmp_ncuenta4= ""
                    tmp_ncuenta5= ""
                    tmp_ncuenta6= ""
                    ncuenta1= ""
				    ncuenta2= ""
				    ncuenta3= ""
	   			    ncuenta4= ""
                    ncuenta5= ""
                    ncuenta6= ""
			    end if

			    ''ricardo 26-12-2006 se mostraran los valores de todos los campos perso
			    ''ricardo 26-12-2006 se calcula cuantos campos perso hay en clientes
                ''ricardo 27-3-2009 se guardaran tantos campos como existan en tabla y con titulo
                ''sera con titutlo, por si utilizamos internamente para alguna cosa, pero claro el usuario no puede tocarla
		        campos_a_dimensionar=num_campos_clientes
		        if cstr(campos_a_dimensionar & "")="" then campos_a_dimensionar=10

                'ricardo 27-3-2009 en lugar de poner una linea por campo, se pone un for que lo hago por la totalidad
                for ki_cli=1 to campos_a_dimensionar
                    cadena_campo="campo" & replace(space(2-len(cstr(ki_cli)))," ","0") & cstr(ki_cli)
                    tmp_lista_valores(ki_cli)=rstAux(cadena_campo)
                next
			    ''ricardo 12-3-2003 si el cliente tiene en documentos_cli un registro se pondra este en lugar de la serie por defecto
			    if request.querystring("cambiar_serie")>"" then
				    cambiar_serie=limpiaCadena(request.querystring("cambiar_serie"))
			    else
				    cambiar_serie=limpiaCadena(request.querystring("cambiar_serie"))
			    end if
			    if cint(null_z(cambiar_serie))=1 or cambiar_serie & ""="" then
				    obtener_doc_cli mode,"facturas_cli",tmp_ncliente,p_serie,"","",tmp_irpf
			    end if
			    '''''''''''
		    else
			    Error="SI"%>
			    <script language="javascript" type="text/javascript">
                            alert("<%=LitMsgClienteNoExiste%>");
			        //history.back();
			    </script>
		    <%end if
		    rstAux.close
	    end if

	    'Acción a realizar
	    if mode="first_save" then	   
		    if compNumDocNuevo(p_serie,p_fecha,"facturas_cli")=0 then
			    %><script language="javascript" type="text/javascript">
                      alert("<%=LitMsgDocExistRevCont%>");
                      document.facturas_cli.action = "facturas_cli.asp?nfactura=&mode=add" + "&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>";
                      document.facturas_cli.submit();
                      parent.botones.document.location = "facturas_cli_bt.asp?mode=add"
			    </script><%
			    CerrarTodo()
			    response.end
			    mode=""
		    end if
	    end if

	    if request.querystring("continuar")>"" then
		    continuar=limpiaCadena(request.querystring("continuar"))
	    else
		    continuar=limpiaCadena(request.form("continuar"))
	    end if
	    if continuar>"" then
	    else
		    continuar=1
	    end if
	    if request.querystring("continuarf")>"" then
		    continuarf=limpiaCadena(request.querystring("continuarf"))
	    else
		    continuarf=limpiaCadena(request.form("continuarf"))
	    end if
	    if continuarf>"" then
	    else
		    continuarf=1
	    end if
	    if request.querystring("continuari")>"" then
		    continuari=limpiaCadena(request.querystring("continuari"))
	    else
		    continuari=limpiaCadena(request.form("continuari"))
	    end if
	    if continuari>"" then
	    else
		    continuari=1
	    end if

	    if request.querystring("forma_pago")>"" then
		    forma_pago=limpiaCadena(request.querystring("forma_pago"))
	    else
		    forma_pago=limpiaCadena(request.form("forma_pago"))
	    end if
	    if mode="save" then
		    if request.querystring("pagada")>"" then
			    pagada=limpiaCadena(request.querystring("pagada"))
		    else
			    pagada=limpiaCadena(request.form("pagada"))
		    end if
	    else
		    pagada=0
	    end if

	    if request.querystring("divisa")>"" then
		    divisa=limpiaCadena(request.querystring("divisa"))
	    else
		    divisa=request.form("divisa")
	    end if
	    if p_ncliente>"" then
		    p_ncliente = session("ncliente") & p_ncliente
	    else
		    if request.querystring("p_ncliente")>"" then
			    p_ncliente=limpiaCadena(request.querystring("p_ncliente"))
		    else
			    p_ncliente=session("ncliente") & request.form("p_ncliente")
		    end if
	    end if
	    if p_nfactura>"" then
	    else
		    if request.querystring("p_nfactura")>"" then
			    p_nfactura=limpiaCadena(request.querystring("p_nfactura"))
		    else
			    p_nfactura=request.form("p_nfactura")
		    end if
	    end if

	    if mode="save" or mode="first_save" then
		    tmp_ncuenta=tmp_ncuenta1 & tmp_ncuenta2 & tmp_ncuenta3 & tmp_ncuenta4 & tmp_ncuenta5 & tmp_ncuenta6
            if tmp_ncuenta="" then 
                tmp_banco=""
            end if
            'response.Write ("cuenta: " & tmp_ncuenta & " :banco: " & tmp_banco)
            'response.End
            pagada = false
		    if ComprobarCuenta(tmp_banco,tmp_ncuenta)=true then
			    pagada_uxa=0
			    if p_cobrada="on" and request.form("h_vpagada")=1 then
				    if rst.state<>0 then rst.close
				    rst.CursorLocation=2
				    rst.Open "select * from facturas_cli with(rowlock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    GuardarRegistro p_nfactura,p_serie,p_fecha,mode
				    p_nfactura=rst("nfactura")
				    if mode="first_save" then
					    auditar_ins_bor session("usuario"),p_nfactura,rst("ncliente"),"alta","","","facturas_cli"
				    end if
				    if rst.state<>0 then rst.close
				    pagada=true
				    pagada_uxa=1 'para que no saque el mensaje de que no se puede modificar la factura
			    end if
			    'cuando no esta cobrada y se pone el cobro
			    if p_cobrada="on" and request.form("h_cobrada")=0 then
				    if rst.state<>0 then rst.close
				    rst.CursorLocation=2
				    rst.Open "select * from facturas_cli with(rowlock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    GuardarRegistro p_nfactura,p_serie,p_fecha,mode
				    p_nfactura=rst("nfactura")
				    if mode="first_save" then
					    auditar_ins_bor session("usuario"),p_nfactura,rst("ncliente"),"alta","","","facturas_cli"
				    end if
				    tf=reemplazar(rst("total_factura"),",",".")
				    fp=rst("forma_pago")

				    if rst.state<>0 then
					    rst.close
				    end if
				
				    if (cint(continuar)=0 or cint(continuarf)=0 or cint(continuari)=0) and gen_vencimiento=-1 then
					    if rst2.state<>0 then rst2.close
					    rst2.Open "delete from vencimientos_salida with(rowlock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					    if rst2.state<>0 then rst2.close
					    if fp>"" then
						    CrearVencimientos p_nfactura,p_ncliente
					    end if
				    end if
				    if rst2.state<>0 then rst2.close
				    pagada=true
				    pagada_uxa=1
			    end if

			    'cuando esta cobrada y se quita el cobro
			    if p_cobrada="" and request.form("h_cobrada")=1 and (EnCaja="" or EnCaja="NO" or EnCaja=0) then

				    'comprobamos que no tengamos ningun vencimiento cobrado en caja
				    'ega 19/06/2008 unir tabla con join, with(nolock)
				    rst2.Open "select v.nvencimiento,f.nfactura from vencimientos_salida as v with(nolock) inner join facturas_cli as f with(nolock) on f.nfactura=v.nfactura where f.nfactura like '" & session("ncliente") & "%' and v.nfactura like '" & session("ncliente") & "%' and f.nfactura='" & p_nfactura & "' order by nvencimiento", _
				    session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    			    'si hay algun vencimiento cobrado y en caja no se podran editar los vencimientos
				    if not rst2.eof then
					    estaencaja=""
					    while not rst2.eof and estaencaja=""
						    'estaencaja=d_lookup("caja","caja","ndocumento='" & rst2("nfactura") & "-" & rst2("nvencimiento") & "'",session("dsn_cliente"))

                            estaencajaSelect="select caja from caja with(nolock) where ndocumento=?"
                            estaencaja=DlookupP1(estaencajaSelect, rst2("nfactura") & "-" & rst2("nvencimiento")&"", adVarchar, 22, session("dsn_cliente"))

						    rst2.movenext
					    wend
					    rst2.movefirst
					    if estaencaja="" then estadoencaja=0 else estadoencaja=1
				    end if
				    rst2.close
				    if estadoencaja=0 then
					    estaencaja2=d_sum("importe","caja","ndocumento like '" & session("ncliente") & "%' and ndocumento='" & p_nfactura & "'",session("dsn_cliente"))
					    if estaencaja2=0 then estadoencaja=0 else estadoencaja="1"
				    end if
				    'FLM:020209: SI NO ESTA EN CAJA MIRAMOS LOS EFECTOS
				    'FLM:20090422: SI NO ESTA EN CAJA MIRAMOS LAS REMESAS
				    EnEfecto="No"
				    EnRemesa="No"
				    if EnCaja="" or EnCaja="NO" or EnCaja=0 or EnCaja="null" then
                        'Comprobamos los efectos de cliente.
                        rstAux.open "select top 1 nefecto from detalles_efcli with(nolock) where nefecto like '" & session("ncliente") & "%' and (nfacturavto='" & p_nfactura & "' or nfactura='" & p_nfactura & "') "
                        if not rstAux.EOF then
                            EnEfecto="SI"
                            pagada_uxa=1 'para que no saque el mensaje de que no se puede modificar la factura
                            pagada=true
                            pagado=1%>
                            <script language="javascript" type="text/javascript">
                                alert("<%=LitMsgNoAnularCobroEfecto%>");
                            </script>
                            <%
                        else
                            EnEfecto="No"
                        end if
                        rstAux.close
                        %>
                            <script language="javascript" type="text/javascript">
                                //        alert("Imposible anular el cobro de la factura. La factura o alguno de sus vencimientos están incluidos en alguna remesa");
                            </script>
                            <%
                            EnRemesa="No"
                    end if
                    'FLM:20090424:Añador el control de las remesas.
				    if estadoencaja=0  and EnEfecto="NO" and EnRemesa="NO"  then
					    pagada=false
					    AnularVencimientos p_nfactura
				    elseif EnEfecto="NO" and EnRemesa="NO" then
					    EnCaja="SI"
					    forma_pago=""%>
					    <script language="javascript" type="text/javascript">
                                alert("<%=LITMSGNOANULARCOBROANOTCAJA%>");
					    </script>
				        <%pagada_uxa=1 'para que no saque el mensaje de que no se puede modificar la factura
				    end if

			    end if

			    if rst.state<>0 then rst.close
			    rst.CursorLocation=2
			    rst.Open "select * from facturas_cli with(rowlock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			    if p_ncliente > "" then
				    rstCliente.open "select ncliente from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & p_ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				    if rstCliente.EOF then%>
					    <script language="javascript" type="text/javascript">
                            alert("<%=LitMsgClienteNoExiste%>");
                            history.back();
					    </script>
					    <%
					    mode="add"
				    else
					    pagado=0
					    if rst2.state<>0 then rst2.close
					    rst2.open "select nvencimiento, nfactura from vencimientos_salida with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "' and importecob>0", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					    if not rst2.eof then
						    pagado=1
					    end if
					    rst2.close
					    'FLM:20090429: Comprobamos si algún vencimiento está en una remesa y si se ha modificado la forma de pago.
                        EstaEnRemesa=0
                        if mode<>"first_save" then
                            if forma_pago<>rst("forma_pago") then 
		                        rst2.open "select top 1 1 from remesas r with(nolock) inner join detalles_remcli dr with(nolock) on dr.nremesa=r.nremesa and dr.nfacturavto='"&p_nfactura&"' where r.nempresa='"&session("ncliente")&"'"		           
			                    if not rst2.eof then
			                        EstaEnRemesa=1
			                        pagada=true
                                    pagado=1
                                    pagada_uxa=1
			                    end if
			                    rst2.close
			                end if
                            rst2.open "select sum(importe) as total from caja with(nolock) where ndocumento='"&p_nfactura&"'"		           
			                if not rst2.eof then
                                if rst2("total") > 0 then
			                        pagada=true
                                    pagado=1
                                    pagada_uxa=1%>
					                <script language="javascript" type="text/javascript">
					                    alert("<%=LITMSGNOANULARCOBROANOTCAJA%>");
					                </script>
				                <%end if
			                end if
			                rst2.close
			            end if

					    if pagado=0 and (pagada=false or pagada="") then
						    if serie&""="" then serie=request.form("serie")
						    ''ricardo 20/2/2003 si se cambian los descuentos,tarifa,fechas o clientes , se debera cambiar las compras habituales
						    si_cambiar_habitual=0
						    if mode="save" then
							    if cdbl(null_z(rst("descuento")))<>cdbl(null_z(request.form("dto1"))) or cdbl(null_z(rst("descuento2")))<>cdbl(null_z(request.form("dto2"))) or cdbl(null_z(rst("descuento3")))<>cdbl(null_z(request.form("dto3"))) then
								    si_cambiar_habitual=1
							    end if
							    if (rst("fecha")<>cdate(Request.Form("fecha")&"")) then
								    si_cambiar_habitual=2
							    end if
							    if (rst("tarifa")&""<>request.form("tarifa")&"") then
								    si_cambiar_habitual=3
							    end if
							    if (rst("ncliente")&""<>p_ncliente&"") then
								    si_cambiar_habitual=4
							    end if
						    else
							    si_cambiar_habitual=0
						    end if
						    '''''''''''''''''
						    GuardarRegistro p_nfactura,serie,p_fecha,mode
						    p_nfactura=rst("nfactura")
						    if mode="first_save" then
							    auditar_ins_bor session("usuario"),p_nfactura,rst("ncliente"),"alta","","","facturas_cli"
						    end if
						    ''ricardo 18/7/2003 se cambia esta linea, ya que no cogia bien la deuda
						    'tf=reemplazar(d_lookup("deuda","facturas_cli","nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente")),",",".")
						    
                            tfSelect="select deuda from facturas_cli with(nolock) where nfactura like ?+'%' and nfactura=?"
                            tf=reemplazar(DlookupP2(tfSelect, session("ncliente")&"", advarchar, 20, p_nfactura&"", adVarchar, 20, session("dsn_cliente")),",",".")
                                    
                            ''''''
						    fp=rst("forma_pago")
						    rst.close
						    if mode="save" then
							    if (cint(continuar)=0 or cint(continuarf)=0 or cint(continuari)=0) and gen_vencimiento=-1 then
								    if rst2.state<>0 then rst2.close
								    rst2.open "SELECT nfactura, nvencimiento FROM VENCIMIENTOS_salida with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "' and cobrado=1",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
								    if rst2.eof then
									    rst2.close
									    rst2.Open "DELETE from vencimientos_salida with(rowlock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
									    AnularVencimientos p_nfactura
									    if rst2.state<>0 then rst2.close
									    if fp>"" then
										    CrearVencimientos p_nfactura,p_ncliente
									    end if
								    else
									    rst2.close
								    end if
							    end if
						    end if
					    else
						
						    if pagada_uxa=0 then%>
							    <script language="javascript" type="text/javascript">							        alert("<%=LitMsgNoModifFactVencPagado%>");</script>
						    <% elseif EstaEnRemesa=1 then %>
    				            <script language="javascript" type="text/javascript">    				                alert("<%=LitMsgNoModifFactRemesa%>");</script>
				            <%end if

					    end if
				    end if
				    rstCliente.close
			    end if
			    ant_mode=mode
			    mode="browse"

			    'inicializamos las variables por si habiamos cambiado algun valor antes de grabar
			    tmp_tipo_pago="":tmp_forma_pago="":tmp_comercial="":tmp_cod_proyecto="":tmp_agenteasignado=""
			    tmp_tarifa="":tmp_fechaenvio="":tmp_fechapedido="":tmp_transportista=""
			    tmp_nombre="":tmp_nenvio=""
                tmp_banco="":tmp_ncuenta1="":tmp_ncuenta2="":tmp_ncuenta3="":tmp_ncuenta4="":tmp_ncuenta5="":tmp_ncuenta6=""
		    else
			    'LA CUENTA NO ES VALIDA%>
			    <script language="javascript" type="text/javascript">
                                    alert("<%=LitMsgCuentaBancoIncorrecta%>");
                                    parent.history.back();
                                    parent.history.back();
                    <%if mode= "save" then%>
                                        parent.botones.location="facturas_cli_bt.asp?mode=edit";
			        <%else%>
                                        parent.botones.location="facturas_cli_bt.asp?mode=add";
                    <%end if%>
			    </script>
			    <%if rst.state<>0 then rst.close
			    rst.Open "select * from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		    end if 
                    'Laura ini
     '  elseif mode="finish" then  
            %>
		<!--	    <script language="javascript" type="text/javascript">
			        alert("Prueba botón terminar");
			    </script>-->
              <%      'Laura fin                           
       elseif mode="delete" then
		    rstAux.open "select nfactura from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		    if rstAux.eof then
			    rstAux.close
			    p_nfactura=""%>
			    <script language="javascript" type="text/javascript">
                        alert("<%=LitMsgDocsNoExiste%>");
                    parent.botones.document.location = "facturas_cli_bt.asp?mode=add";
			    </script>
			    <%mode="add"
		    else
			    rstAux.close
			    he_borrado=1
			    if rst.state<>0 then rst.close
			    'Comprobar si se puede eliminar la factura.
			    mensajeTratEquipos=TratarEquipos("","","FACTURA A CLIENTE",p_nfactura,"","","","","",mode)
			    if mid(mensajeTratEquipos,1,2)<>"OK" then
				    mode="browse"%>
				    <script language="javascript" type="text/javascript">
                        alert("<%=mensajeTratEquipos%>");
				    </script>
			    <%else
			        'FLM:20090424:Comprobamos que no haya ninguna remesa con el vencimiento
		            rst.open "select top 1 r.nremesa from remesas r with(nolock) inner join detalles_remcli dr with(nolock) on dr.nremesa=r.nremesa and (dr.nfacturavto='" & p_nfactura & "'  or dr.nfactura='" & p_nfactura & "') where r.nempresa='" & session("ncliente") & "' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		            if rst.EOF then
			            rst.close
		                'FLM:050409: SI está en efecto no se puede borrar.
			            rst.open "select nfactura from detalles_efcli with(nolock) where nefecto like '" & session("ncliente") & "%' and ( (nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "') or (nfacturavto like '" & session("ncliente") & "%' and nfacturavto='" & p_nfactura & "')  )",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			            if rst.EOF then
			                rst.close
			                'ega 19/06/2008 agrego like de la clave primaria
				            rst.open "select ndocumento from caja with(nolock) where caja like '" & session("ncliente") & "%' and ndocumento like '" & session("ncliente") & "%' and ndocumento='" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				            if rst.EOF then
					            rst.close
					            rst.open "select ndocumento from detalles_dev_cli with(nolock) where ndevolucion like '" & session("ncliente") & "%' and ndocumento='" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					            if rst.EOF then
						            rst.close
						            rst.open "select ncliente from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						            ncli_aux=rst("ncliente")
						            rst.close
						            if ComprobacionAtribPuntos(p_nfactura)=0 then
							            auditar_ins_bor session("usuario"),p_nfactura,ncli_aux,"baja","","","facturas_cli"
							            InsertarHistorialNserie mensajeTratEquipos,"","","FACTURA A CLIENTE",p_nfactura,"","","","","",mode
							            BorrarRegistro p_nfactura
						            end if
						            'se pone a add, ya que cuando solo hay una factura, no se puede mostrar ninguna factura
						            mode="add"
						            p_nfactura=""%>
						            <script language="javascript" type="text/javascript">
                                        parent.botones.document.location = "facturas_cli_bt.asp?mode=add";
                                        SearchPage("facturas_cli_lsearch.asp?mode=init", 0);
						            </script>
					            <%else
						            rst.close
						            mode="browse"%>
						            <script language="javascript" type="text/javascript">
                                        alert("<%=LitMsgNoBorrarFactDev%>");
						            </script>
					            <%end if
				            else
					            rst.close
					            mode="browse"%>
					            <script language="javascript" type="text/javascript">
                                        alert("<%=LitMsgNoBorrarFactAnotCaja%>");
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
			    'inicializamos las variables por si habiamos cambiado algun valor antes de grabar
			    tmp_tipo_pago="":tmp_forma_pago="":tmp_comercial="":tmp_cod_proyecto="":tmp_agenteasignado=""
			    tmp_tarifa="":tmp_fechaenvio="":tmp_fechapedido="":tmp_transportista=""
			    tmp_nombre="":tmp_nenvio=""
                tmp_banco="":tmp_ncuenta1="":tmp_ncuenta2="":tmp_ncuenta3="":tmp_ncuenta4="":tmp_ncuenta5="":tmp_ncuenta6=""
		    end if
        end if
        'Mostrar los datos de la página.
        ''ricardo 31/7/2003 comprobamos que existe el albaran
        if (mode="browse" or mode="edit") and he_borrado<>1 then
            rstAux.cursorlocation=3
	        rstAux.open "select nfactura from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'", session("dsn_cliente")
	        if rstAux.eof then
		        p_nfactura=""%>
		        <script language="javascript" type="text/javascript">
                            alert("<%=LitMsgDocsNoExiste%>");
                            parent.botones.document.location = "facturas_cli_bt.asp?mode=add";
		        </script>
		        <%mode="add"
	        end if
	        rstAux.close
        end if

        if mode="browse" or mode="edit" then
		    if p_nfactura="" then
			    set rstAux = Server.CreateObject("ADODB.Recordset")
			    p_nfactura=d_lookup("nfactura","facturas_cli","",session("dsn_cliente"))
       
		    end if

		    ' JMA 28/10/04 Campos personalizables'
		    ''ricardo 26-12-2006 se mostraran los valores de todos los campos perso
            ''ricardo 27-3-2009 se cambia el select , para que salgan todos los campos perso
            if cstr(num_campos_ventas & "")="" then num_campos_ventas=10
            strGuarFacCli="select p.nfactura,"
            strGuarFacCliAux=""
            for ki_cli=1 to num_campos_ventas
                cadena_campo="p.campo" & replace(space(2-len(cstr(ki_cli)))," ","0") & cstr(ki_cli)
                strGuarFacCliAux=strGuarFacCliAux & cadena_campo & ","
            next

            strGuarFacCli=strGuarFacCli & mid(strGuarFacCliAux,1,len(strGuarFacCliAux)-1)
            strGuarFacCli=strGuarFacCli & " from facturas_cli as p with(nolock) where p.nfactura='" & p_nfactura & "'"

		    rstAux.cursorlocation=3
		    rstAux.open strGuarFacCli,session("dsn_cliente")
		    if not rstAux.eof then
                ''ricardo 27-3-2009 se guardaran tantos campos como existan en tabla y con titulo
                ''sera con titutlo, por si utilizamos internamente para alguna cosa, pero claro el usuario no puede tocarla

                if cstr(num_campos_ventas & "")="" then num_campos_ventas=10
                redim lista_valores(num_campos_ventas+2)
                ''ricardo 27-3-2009 en lugar de poner una linea por campo, se pone un for que lo hago por la totalidad
                for ki_cli=1 to num_campos_ventas
                    cadena_campo="campo" & replace(space(2-len(cstr(ki_cli)))," ","0") & cstr(ki_cli)
                    lista_valores(ki_cli)=Nulear(rstAux(cadena_campo))
                next
			        num_campos=num_campos_ventas
		    else
		        if cstr(num_campos_ventas & "")="" then num_campos_ventas=10
			    redim lista_valores(num_campos_ventas+2)
			    for ki=1 to num_campos_ventas+2
				    lista_valores(ki)=""
			    next
			    num_campos=num_campos_ventas
		    end if
		    rstAux.close

		    if rst.state<>0 then rst.close
            StrSelGetFact="select p.*,c.rsocial,c.riesgo1,c.riesgo2,b.nliquidacion,h.nliquidacion as nliquidacionAG,tmp.total_suplidos "
		    StrSelGetFact=StrSelGetFact & " from facturas_cli as p with(nolock) "
            StrSelGetFact=StrSelGetFact &  " left outer join detalles_liq b with(nolock) on p.nfactura=b.nfactura "
            StrSelGetFact=StrSelGetFact &  " left outer join detalles_liq_ag h with(nolock) on p.nfactura=h.nfactura "
            StrSelGetFact=StrSelGetFact &  " left outer join ( "
            StrSelGetFact=StrSelGetFact &  " select s.nfactura,sum(importe) as total_suplidos from SUPLIDOS_FAC_CLI as s with(NOLOCK) where s.nfactura='" & p_nfactura & "' group by s.nfactura "
            StrSelGetFact=StrSelGetFact &  " ) as tmp on tmp.nfactura=p.nfactura "
            StrSelGetFact=StrSelGetFact &  " ,clientes as c with(NOLOCK) "
            StrSelGetFact=StrSelGetFact &  " where c.ncliente=p.ncliente "
		    StrSelGetFact=StrSelGetFact & " and p.nfactura like '" & session("ncliente") & "%' and p.nfactura='" & p_nfactura & "'"
            rst.cursorlocation=3
		    rst.Open StrSelGetFact, session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
		    if not rst.eof then
			    'n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))
                    
                n_decimalesSelect="select ndecimales from divisas with(nolock) where codigo =?"
                n_decimales=DlookupP1(n_decimalesSelect, rst("divisa")&"", adVarchar, 15, session("dsn_cliente"))
                    %>
			    <input type="hidden" name="importe_ant" value="<%=formatnumber(null_z(rst("total_factura")),n_decimales,-1,0,iif(mode="browse",-1,0))%>"/>
		    <%else%>
			    <input type="hidden" name="importe_ant" value="0"/>
		    <%end if%>
		    <input type="hidden" name="nliquidacionAG" value="<%=rst("nliquidacionAG")%>"/>
		    <input type="hidden" name="nliquidacion" value="<%=rst("nliquidacion")%>"/>
		    <%nliquidacionAG=nulear(rst("nliquidacionAG"))
		    nliquidacion=nulear(rst("nliquidacion"))
		    imp_Serie = rst("serie")
		    ''ricardo 26/4/2004
		    if not rst.eof and mode="browse" then
			    'preguntar_riesgo_conf=nz_b(d_lookup("RIESGOOBLIG","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))

                preguntar_riesgo_confSelect="select RIESGOOBLIG from configuracion with(nolock) where nempresa =?"
                preguntar_riesgo_conf=nz_b(DlookupP1(preguntar_riesgo_confSelect, session("ncliente")&"", adChar, 5, session("dsn_cliente")))

			    'contrasenya_riesgo_conf=null_s(d_lookup("contrpermries","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))
                
                contrasenya_riesgo_confSelect="select contrpermries from configuracion with(nolock) where nempresa=?"
                contrasenya_riesgo_conf=null_s(DlookupP1(contrasenya_riesgo_confSelect, session("ncliente")&"", adChar, 5, session("dsn_cliente")))
                %>
			    <input type="hidden" name="rmaxaut" value="<%=enc.EncodeForHtmlAttribute(rst("riesgo1"))%>"/>
			    <input type="hidden" name="ralc" value="<%=enc.EncodeForHtmlAttribute(rst("riesgo2"))%>"/>
			    <input type="hidden" name="prieconf" value="<%=enc.EncodeForHtmlAttribute(preguntar_riesgo_conf)%>"/>
                <%''ricardo 27-12-2006 se pone el replace, ya que si la razon social contiene el caracter chr(34) daba error al intentar acceder al campo contrpregries, a la hora de guardar detalles y conceptos%>
			    <input type="hidden" name="rsocries" value="<%=enc.EncodeForHtmlAttribute(replace(rst("rsocial"),chr(34),""))%>"/>
			    <input type="hidden" name="contrpregries" value="<%=enc.EncodeForHtmlAttribute(contrasenya_riesgo_conf)%>"/>
		    <%end if
	    elseif mode="add" then
		    if rst.state<>0 then rst.close
            rst.cursorlocation=3
		    rst.Open "select * from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'", _
		    session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
		    ''rst.AddNew
	    elseif mode="search" then
		
	    end if

	    sumadet=0
	    sumaRE=0

	    'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION
	    if mode="edit" then%>
		    <div ID="venci_paga" style="display: none">
		        <table width="100%">
		           <tr>
			            <td class="CELDAREDBOLD" align="center">
				            <%=LitVPagado%>
			            </td>
		            </tr>
		        </table>
	        </div>
	    <%end if
	    'EBF En caso de que tenga el campo de aplicar dtos de cliente a detalles guardamos el total de compras que hizo el cliente
	    ' el mes del año anterior a el mes que marque la fecha de la factura y los dtos del cliente
	    if (mode="browse" or mode="save" or mode="first_save") and dto_cli_art=true then
		    fecha_consulta=iif(p_fecha>"" ,p_fecha,rst("fecha"))
		    mes=datepart("m",fecha_consulta)
		    ano=datepart("yyyy",fecha_consulta)
		    if si_tiene_modulo_ebesa<>0 then
			    strselect="  select dto, dto2, dto3, isnull(total,0) as total, isnull(total2,0)as total2 " &_
			               " from clientes with(nolock), " & _
			               " (select sum(total_factura) as total from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and ncliente='"&iif(tmp_ncliente>"",p_ncliente,rst("ncliente"))&"' and fecha>='1/"&mes&"/"&ano-1&"' and fecha <'1/"&iif(mes=12,1,mes+1)&"/"&iif(mes=12,ano,ano-1)&"')as tmp," &_
			               " (select sum(total_factura) as total2 from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and ncliente='"&iif(tmp_ncliente>"",p_ncliente,rst("ncliente"))&"' and fecha>='1/"&mes&"/"&ano&"' and fecha <'"&now&"')as tmp2 "&_
			               " where ncliente like '" & session("ncliente") & "%' and ncliente='"&iif(tmp_ncliente>"",p_ncliente,rst("ncliente"))&"' "
		    else
			    strselect="select dto,dto2,dto3 from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='"&iif(tmp_ncliente>"",p_ncliente,rst("ncliente"))&"' "
		    end if
            rstAux.cursorlocation=3
		    rstAux.open strselect,session("dsn_cliente")
		    if not rstAux.eof then
			    if si_tiene_modulo_ebesa<>0 then
				    ganancia=iif(cdbl(rstAux("total2"))-cdbl(rstAux("total"))>=0,1,0)
			    else
				    ganancia=0
			    end if%>
			    <input type="hidden" name="dto1_cli" value="<%=enc.EncodeForHtmlAttribute(rstAux("dto"))%>"/>
			    <input type="hidden" name="dto2_cli" value="<%=enc.EncodeForHtmlAttribute(rstAux("dto2"))%>"/>
			    <input type="hidden" name="dto3_cli" value="<%=enc.EncodeForHtmlAttribute(rstAux("dto3"))%>"/>
			    <%'el valor de ganancia solo afectará en caso de que se tenga el modulo ebesa%>
			    <input type="hidden" name="ganancia" value="<%=ganancia%>"/>
			    <%dto1_cli=rstAux("dto")
			    dto2_cli=rstAux("dto2")
			    dto3_cli=rstAux("dto3")
		    end if
		    rstAux.close
	    end if
	    if mode="browse" then
            '*** COMENTAR u OCULTAR ANTES DE PASAR A PRODUCCION ***
	        'tipodoc=d_lookup("codigo","tipo_documentos","tippdoc='FACTURA A CLIENTE'",DSNIlion)
	        'if not rst.eof then
	        '	ImprimirFormato p_nfactura,tipodoc,"facturas_cli",rst("serie")
	        'end if
            '***************%>
	            <!--<hr/>-->
	    <%end if

        %>
        <%VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarCentros)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina%>

            <% pagina="../central.asp?pag1=ventas/ventas_habituales.asp&pag2=ventas/ventas_habituales_bt.asp&titulo=" & Ucase(LitVenHab)%>
				    
			<% ''aqui va un /actionVersion %>

        <div class="headers-wrapper">
            <%
                    DrawDiv "header-date","",""
                    DrawLabel "","",LitFecha
                        if mode="browse" then
                            DrawSpan "","",rst("fecha"), ""
                        else
                            DrawInput "width150px", "", "fecha", iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha")), ""
                            DrawCalendar "fecha"
                        end if
                    CloseDiv
                    if vienenp>"" then p_fecha=""
			        if mode="edit" then
                        %><input type="hidden" name="h_fecha" value="<%=enc.EncodeForHtmlAttribute(rst("fecha"))%>"/><%
                    end if

                    if not rst.eof then
                    %><input type="hidden" name="fecha_ant" value="<%=enc.EncodeForHtmlAttribute(iif(tmp_fecha>"",tmp_fecha,rst("fecha")))%>"/><%
                    else
                        if tmp_fecha & "">"" then
                            %><input type="hidden" name="fecha_ant" value="<%=enc.EncodeForHtmlAttribute(iif(tmp_fecha>"",tmp_fecha,date()))%>"/><%
                        else 
                            %><input type="hidden" name="fecha_ant" value="<%=enc.EncodeForHtmlAttribute(date())%>"/><%
                        end if
                    end if
			        codigoMostrarNfolio = ""
			        if gestionFolios then
			            if mode="browse" then
				            codigoMostrarNfolio = " / " & UCASE(LITNUMFOLIO) & " : " & rst("nfolio")
				        else			    
				            if mode="edit" then
				                codigoMostrarNfolio = " / " & UCASE(LITNUMFOLIO) & " : <input style='width:80px' type='text' class='CELDA' name='nfolio' id='nfolioID' value='"& enc.EncodeForHtmlAttribute(rst("nfolio")) &"'/>"
				            elseif mode="add" then
				                codigoMostrarNfolio = " / " & UCASE(LITNUMFOLIO) & " : <input style='width:80px' type='text' class='CELDA' name='nfolio' id='nfolioID' value='"& enc.EncodeForHtmlAttribute(nfolio) &"'/>"
					        end if
					    
				        end if
			        end if

                     DrawDiv "header-bill","",""
                    ' --- Fin JMMM
                    if mode="browse" or mode="edit" then
                        if not rst.eof then
                            DrawLabel "","",Litfactura
                            DrawSpan "","",trimCodEmpresa(rst("nfactura")) & codigoMostrarNfolio, ""
                        else
                            DrawLabel "","",Litfactura
                            DrawSpan "","",codigoMostrarNfolio, ""
                        end if
                    else
                        DrawLabel "","",Litfactura
                        DrawSpan "","",codigoMostrarNfolio, ""
                    end if
                    CloseDiv
                    if session("version")&"" <> "5" then
                        DrawDiv "","","" 
                        CloseDiv
                    end if 
                    DrawDiv "header-client-iframe","",""
                        Formulario="facturas_cli"
                        if mode="browse" then 
                            DrawLabel "","",LitCliente
					            if rst("ncliente")>"" then
                                    DrawSpan "","",Hiperv(OBJClientes,rst("ncliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("ncliente")),LitVerCliente),""
                                    %><input class='' id="input-ncliente" style="width:40px" type="hidden" name="ncliente" value="<%=enc.EncodeForHtmlAttribute(rst("ncliente"))%>" size="10" /><%
                                    nomcli = d_lookup("rsocial","clientes","ncliente='" & iif(tmp_ncliente>"" and mode<>"browse",tmp_ncliente,rst("ncliente")) & "'",session("dsn_cliente"))
				                    
                                    nomcliSelect=" select rsocial from clientes with(nolock) where ncliente=?"
                                    nomcli=DlookupP1(nomcliSelect, iif(tmp_ncliente>"" and mode<>"browse",tmp_ncliente,rst("ncliente")), adChar, 10, session("dsn_cliente"))

                                    ''EligeCelda "input", mode,iif(mode="browse","CELDA style='width:160px'","CELDA DISABLED style='width:160px'"),"","",0,"","nombre",40,nomcli
                                    DrawSpan "","",nomcli,""

                                    
                                end if
                        else
                            if mode="edit" then
                                DrawLabel "","",LitCliente
                            elseif mode="add" and vienenp>"" then
					            docs=split(vienenp,"-") 
  		                        ninv=docs(0)
  		                        if Ubound(docs)>0 then  nven=int(docs(1)) end if
                        
                                'tmp_ncliente = d_lookup("ncliente","facturas_cli","nfactura='" & ninv & "'",session("dsn_cliente"))

                                tmp_nclienteSelect="select ncliente from facturas_cli with(nolock) where nfactura=?"
                                tmp_ncliente=DlookupP1(tmp_nclienteSelect, ninv&"", adVarchar, 20, session("dsn_cliente"))
                                
                                DrawLabel "","",Hiperv(OBJClientes,"","add","facturas_cli",Permisos,Enlaces,session("usuario"),session("ncliente"),LitCliente,LitAnadirCliente)						

                            elseif mode="add" then
                                DrawLabel "","",Hiperv(OBJClientes,"","add","facturas_cli",Permisos,Enlaces,session("usuario"),session("ncliente"),LitCliente,LitAnadirCliente)
                                %>
				            <%end if
                            if mode="edit" then
                                %><input size="5" class="width20" style=" vertical-align: middle;" type="text" name="ncliente" value="<%=iif(tmp_ncliente>"",trimCodEmpresa(tmp_ncliente),trimCodEmpresa(rst("ncliente")))%>" maxlength="5" onchange="TraerCliente('<%=enc.EncodeForJavascript(null_s(mode))%>','1');"/>
						            <a class="CELDAREFB" href="javascript:AbrirVentana('clientes_buscar.asp?ndoc=<%=enc.EncodeForJavascript(Formulario)%>&titulo=<%=LITSELCLIENTE%>&mode=search&dtos=<%=enc.EncodeForJavascript(dto_cli_art)%>','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitVerCliente%>'; return true;" OnMouseOut="self.status=''; return true;"><IMG style=" vertical-align: middle;" SRC="<%=enc.EncodeForHtmlAttribute(ImgBuscarDinamic)%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><%
                            else
                                %><input size="5" class="width20" style=" vertical-align: middle;" type="text" name="ncliente" value="<%=iif(tmp_ncliente>"",trimCodEmpresa(tmp_ncliente),"")%>" maxlength="5" onchange="TraerCliente('<%=enc.EncodeForJavascript(null_s(mode))%>','1');"/>
						            <a class="CELDAREFB" href="javascript:AbrirVentana('clientes_buscar.asp?ndoc=<%=enc.EncodeForJavascript(Formulario)%>&titulo=<%=LITSELCLIENTE%>&mode=search&dtos=<%=enc.EncodeForJavascript(dto_cli_art)%>','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitVerCliente%>'; return true;" OnMouseOut="self.status=''; return true;"><IMG style=" vertical-align: middle;" SRC="<%=enc.EncodeForHtmlAttribute(ImgBuscarDinamic)%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><%
                            end if
                            if mode<>"browse" then
                                if mode="add" then
                                    'nomcli   = d_lookup("rsocial","clientes","ncliente='" & iif(tmp_ncliente>"" and mode<>"browse",tmp_ncliente,"") & "'",session("dsn_cliente"))

                                    nomcliSelect="select rsocial from clientes with(nolock) where ncliente=?"
                                    nomcli=DlookupP1(nomcliSelect, iif(tmp_ncliente>"" and mode<>"browse",tmp_ncliente,""), adChar, 10, session("dsn_cliente"))
                                else
                                    'nomcli   = d_lookup("rsocial","clientes","ncliente='" & iif(tmp_ncliente>"" and mode<>"browse",tmp_ncliente,rst("ncliente")) & "'",session("dsn_cliente"))

                                    nomcliSelect="select rsocial from clientes with(nolock) where ncliente=?"
                                    nomcli=DlookupP1(nomcliSelect, iif(tmp_ncliente>"" and mode<>"browse",tmp_ncliente,rst("ncliente")), adChar, 10, session("dsn_cliente"))
                                end if 
				                
                                %><input class="width60" id="input-nomcli" style="vertical-align: middle;" type="text" size="40" name="nombre" value="<%=enc.EncodeForHtmlAttribute(nomcli)%>" /><%
                            end if
                        end if
                    CloseDiv
                    
                        if si_tiene_modulo_mantenimiento<>0 then
                            DrawDiv "header-center","",""

					        'ZONA DEL CAMPO PARA EL CENTRO
					        if mode="edit" or mode="add" then
                               
                                DrawLabel "","",LitCentro
                                if mode="edit" then
                                    %><input class="CELDA" type="hidden" name="ncentro" value="<%=enc.EncodeForHtmlAttribute(iif(TraerCliente>"","",iif(isnull(rst("ncentro")),"",rst("ncentro"))))%>"/><%
							            %><iframe id='frCentro' name='fr_Centro' src='../mantenimiento/doccentrosResponsive.asp?viene=facturas_cli&mode=<%=enc.EncodeForHtmlAttribute(mode)%>&ncentro=<%=enc.EncodeForHtmlAttribute(iif(TraerCliente>"","",iif(isnull(rst("ncentro")),"",rst("ncentro"))))%>&ncliente=<%=enc.EncodeForHtmlAttribute(iif(tmp_ncliente>"",tmp_ncliente,iif(isnull(rst("ncliente")),"",rst("ncliente"))))%>' style=" vertical-align: middle;" frameborder="no" scrolling="no" noresize="noresize"></iframe><%
                                else
                                    %><input class="CELDA" type="hidden" name="ncentro" value="<%=iif(TraerCliente>"","","")%>"/><%
							            %><iframe id='frCentro' name='fr_Centro' src='../mantenimiento/doccentrosResponsive.asp?viene=facturas_cli&mode=<%=enc.EncodeForHtmlAttribute(mode)%>&ncentro=<%=iif(TraerCliente>"","","")%>&ncliente=<%=enc.EncodeForHtmlAttribute(iif(tmp_ncliente>"",tmp_ncliente,""))%>' style=" vertical-align: middle;" frameborder="no" scrolling="no" noresize="noresize"></iframe><%
                                end if
                            elseif mode="browse" then
                             
                                DrawLabel "","",LitCentro
						        'nomcentro   = d_lookup("rsocial","centros","ncentro='" & iif(tmp_ncentro>"",tmp_ncentro,rst("ncentro")) & "'",session("dsn_cliente"))
                                            
                                nomcentroSelect="select rsocial from centros with(nolock) where ncentro=?"
                                nomcentro=DlookupP1(nomcentroSelect, iif(tmp_ncentro>"",tmp_ncentro,rst("ncentro")), adVarchar, 10, session("dsn_cliente"))
                                            
                                            %>
						        <!--<td class="celda" align="left" style="width:70px">-->
							        <%if rst("ncentro")<>"00000" then
								        DrawSpan "","",Hiperv(OBJCentros,rst("ncentro"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("ncentro")),LitVerCentro),""
							        else
								        DrawSpan "","",rst("ncentro"),""
							        end if
                                        %><input type="hidden" name="ncentro" value="<%=enc.EncodeForHtmlAttribute(iif(tmp_ncentro>"",tmp_ncentro,rst("ncentro")))%>"/><%
                                DrawSpan "","",nomcentro,""
					        end if
                            CloseDiv
				        else
					        %><!--<span class="CELDA width5" >
                            <div class="data_client">-->
                            <%
				        end if
                        
                        %>

        <%
        ''response.end

        if not rst.EOF then
		    'si hay algun vencimiento cobrado y en caja no se podran editar los vencimientos
		    'ega 19/06/2008 solamente los campos necesarios y union de tablas con join
            rst2.cursorlocation=3
		    rst2.Open "select v.nfactura, v.nvencimiento from vencimientos_salida as v with(nolock) inner join facturas_cli as f with(nolock) on f.nfactura=v.nfactura where f.nfactura like '" & session("ncliente") & "%' and f.nfactura='" & rst("nfactura") & "' order by nvencimiento", _
		    session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
		
		    '''ASP 31/01/2011 Firma Factura Mx
            if mejico then
		        if existeFirma = 1 then
		            estadoencaja=0
		        else
		            estadoencaja=1
		        end if
		    else
		        if not rst2.eof then
			        estaencaja=""
			        'comprobamos si hay alguna entrada en caja de esta factura, ya que si la hay no se podra modificar los vencimientos
			        while not rst2.eof and estaencaja=""
				        'estaencaja=d_lookup("caja","caja","ndocumento like '" & session("ncliente") & "%' and ndocumento='" & rst2("nfactura") & "-" & rst2("nvencimiento") & "'",session("dsn_cliente"))
				        
                        estaencajaSelect = "select caja from caja with(nolock) where ndocumento like ?+'%' and ndocumento=?"
                        estaencaja = DlookupP2(estaencajaSelect, session("ncliente")&"", advarchar, 22, rst2("nfactura") & "-" & rst2("nvencimiento"), advarchar, 22, session("dsn_cliente"))
            
                        rst2.movenext
			        wend
			        if estaencaja="" then estadoencaja=0 else estadoencaja=1
		        end if
		        rst2.close
		        if estadoencaja=0 then
			        estaencaja2=d_sum("importe","caja","ndocumento like '" & session("ncliente") & "%' and ndocumento='" & rst("nfactura") & "'",session("dsn_cliente"))
			        if estaencaja2=0 then estadoencaja=0 else estadoencaja="1"
		        end if
            end if
        end if
        
	            if mode="browse" then
                    if session("version")&"" <> "5" then
                     DrawDiv "","","" 
                      CloseDiv
                    end if 
                    DrawDiv "header-note","",""
			        if not rst.eof then
				        pagado=0
				        if rst2.state<>0 then rst2.close
				        'nVencimientos=d_lookup("nfactura","vencimientos_salida","nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") &"'",session("dsn_cliente"))
                        
                        nVencimientosSelect="select nfactura from vencimientos_salida with(nolock) where nfactura like ?+'%' and nfactura=?"
                        nVencimientos=DlookupP2(nVencimientosSelect, session("ncliente")&"", adVarchar, 20, rst("nfactura")&"", adVarchar, 20, session("dsn_cliente"))

				        if nVencimientos>"" then
                            rst2.cursorlocation=3
					        rst2.open "SELECT nfactura, nvencimiento FROM VENCIMIENTOS_SALIDA with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "' and importecob=0",session("dsn_cliente")
					        if rst2.eof then
						        pagado=1
					        end if
					        rst2.close
				        end if

                        ''ricardo 9-11-2009 no saldra la caja para el modulo profesionales
                        if si_tiene_modulo_profesionales=0 then
                            ''ricardo 7-4-2006 si el parametro cpc=0 no se pondran las cajas
				            if rst("cobrada")=0 and cstr(cpc)<>"0" and rst("ahora")=0 and p_pagsl then
					            'MB=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))

                                MBSelect="select codigo from divisas with(nolock) where moneda_base<>0 and codigo like ?+'%'"
                                MB=DlookupP1(MBSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))

					            'n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") &"' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))

                                n_decimalesSelect="select ndecimales from divisas with(nolock) where codigo = ? and codigo like ?+'%'"
                                n_decimales=DlookupP2(n_decimalesSelect, rst("divisa")&"", adVarchar, 15, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))


					            EnCaja=CambioDivisa(d_sum("importe","pagos","nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "'",session("dsn_cliente")),rst("divisa"),rst("divisa"))
					            Pendiente=miround(null_z(rst("deuda")),n_decimales)
					            defecto=""
					            poner_cajasResponsive1 "input-ncaja",defecto,"ncaja","100","codigo","descripcion","","",poner_comillas(caju)
			  		            %><span class="header-note-inputCaja"><%
						            %><input class='CELDAR7' type="Text" id="input-impcaja" name="impcaja" value="<%=enc.EncodeForHtmlAttribute(Pendiente)%>" size="12"/><%
					            %></span><%
					            %><span class="header-note-currency"><%
                                    'd_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & rst("divisa") & "'",session("dsn_cliente"))
                                    currencySelect="select abreviatura from divisas with(nolock) where codigo like ?+'%' and codigo=?"
						            %><font class=ENCABEZADOR7><%=enc.EncodeForHtmlAttribute(DlookupP2(currencySelect, session("ncliente")&"", adVarchar, 15, rst("divisa")&"", adVarchar, 15, session("dsn_cliente")))%></font><%
					            %></span><%
					            %><span class="header-note-buttonNote"><%
						            %><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgAnotar)%>" <%=ParamImgAnotar%> alt="<%=LitAnotarCaja%>" title="<%=LitAnotarCaja%>" onclick="Acaja('<%=enc.EncodeForJavascript(rst("nfactura"))%>','<%=enc.EncodeForJavascript(Pendiente)%>')"><%
						            %><input type="hidden" name="h_impcaja" value="<%=enc.EncodeForHtmlAttribute(Pendiente)%>"/><%
			  		            %></span><%

                                rstAux.cursorlocation=3
			  		            rstAux.Open "SELECT codigo, descripcion FROM Tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
					            DrawSelect "input-i_pago", "width:150px;", "i_pago", rstAux, session("ncliente") & "01", "codigo", "Descripcion","",""
					            rstAux.Close
				            end if
                        else
                          %><span class='CELDAL7' ><%
                                %>&nbsp;<%
                                %><input class='CELDAR7' type="hidden" name="impcaja" value="<%=enc.EncodeForHtmlAttribute(Pendiente)%>" size="12"/><%
                            %></span><%
                            %><span class='CELDAL7' ><%
                                %>&nbsp;<%
                            %></span><%
                            %><span class='CELDAL7' ><%
                                %>&nbsp;<%
                                %><input type="hidden" name="h_impcaja" value="<%=enc.EncodeForHtmlAttribute(Pendiente)%>"/><%
                            %></span>
                        <%end if
			        else%>
				        <!--<span align="center" >
			            </span>-->
			        <%end if
                    CloseDiv
		        else%>
			        <!--<span align="center" >
			        </span>-->
		        <%end if

                if mode="browse" or mode="save" then
                    DrawDiv "header-card","",""
                    's_empresa = d_lookup("nserie","series","nserie like '" & session("ncliente") & "%' and tipo_documento='FACTURA A CLIENTE' and pordefecto=1", session("dsn_cliente"))

                    s_empresaSelect="select nserie from series with(nolock) where nserie like ?+'%' and tipo_documento='FACTURA A CLIENTE' and pordefecto=1"
                    s_empresa=DlookupP1(s_empresaSelect, session("ncliente")&"", adVarchar, 10, session("dsn_cliente"))

			        rstAux.cursorlocation=3
			        rstAux.Open "select codigo, nombre from cartas with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre", session("dsn_cliente")
                    %><label><a class='CELDAREFB' href="javascript:validarCampoCarta('<%=enc.EncodeForJavascript(p_nfactura)%>','<%=enc.EncodeForJavascript(session("ncliente"))%>');" OnMouseOver="self.status='<%=LitCartaPresentacion%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitCartaPresentacion%> </a></label><%

                        DrawSelectHeaderPressLetter "CELDARIGHT","60","",0,"","cartas",rstAux,"","codigo","nombre","",""
                        rstAux.close

                    CloseDiv

			        seleccion = "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a with(nolock) inner join formatos_imp as b with(nolock) on a.nformato=b.nformato where a.ncliente='"&session("ncliente")&"' and b.tippdoc='FACTURA A CLIENTE' order by descripcion"
			        rstSelect.Open seleccion, DsnIlion, adOpenKeyset, adLockOptimistic
			        if not rstSelect.eof then
				        if rstSelect("personalizacion")&"">"" then
					        personalizacion="../Custom/" & rstSelect("personalizacion") & "/ventas/"
					        personalizacionEmail="Custom/" & rstSelect("personalizacion") & "/ventas/"
				        else
					        personalizacionEmail="ventas/"
				        end if
			        else
				        personalizacionEmail="ventas/"
			        end if

                    ''ricardo 13-3-20003 si la serie tiene un formato de impresion sera este el de por defecto si no sera el elegido en la tabla formatos impresion de ilion
		            if not rst.eof then
			            defecto=obtener_formato_imp(rst("serie"),"FACTURA A CLIENTE")
		            end if
                    '''''''''
                    'desde aqui
                    if (rnf="1" or pnf="1") and saft=True  and rst("ahora")=0 then
                    else
                        DrawDiv "header-print header-print-top-fac","",""
                        %><label><a id="idPrintFormat" class="CELDAREFB" href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(personalizacion)%>' + document.facturas_cli.formato_impresion.value+'nfactura=<%="(\'"+enc.EncodeForJavascript(p_nfactura)+"\')"%>&mode=browse&empresa=<%=enc.EncodeForJavascript(session("ncliente"))%>&cajaParam=<%=enc.EncodeForJavascript(session("f_caja"))%>&novei=<%=enc.EncodeForJavascript(novei)%>','I',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitImpresionConFormato%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitImpresionConFormatoP%> </a></label><%
                            %><select class='CELDA' name="formato_impresion"><%
                                    encontrado=0
			                    while not rstSelect.eof
			     
			                         imprimirOK=false
			                         'las facturas
			                         if saft=True then
			                             if rst("importe_bruto")>=0 then
			                                 if Instr(ucase(rstSelect("descripcion")), "NC")= false then
			                                    imprimirOK=true
			                                 end if
			                             else
			                                if Instr(ucase(rstSelect("descripcion")), "NC") then
			                                    imprimirOK=true
			                                 end if
			                             end if
			                        else
			                            imprimirOK=true
			                        end if 
			                        if imprimirOK=true then
				                        if defecto=rstSelect("descripcion") then
					                        encontrado=1
					                        if isnull(rstSelect("parametros")) then
						                        prm=""
					                        else
						                        prm=rstSelect("parametros") & "&"
					                        end if
                                            %><option selected="selected" value="<%=enc.EncodeForHtmlAttribute(rstSelect("fichero")) & "?" & enc.EncodeForHtmlAttribute(prm)%>"><%=enc.EncodeForHtmlAttribute(rstSelect("descripcion"))%></option><%
                                        else
					                        if isnull(rstSelect("parametros")) then
						                        prm=""
					                        else
						                        prm=rstSelect("parametros") & "&"
					                        end if
                                                %><option value="<%=enc.EncodeForHtmlAttribute(rstSelect("fichero")) & "?" & enc.EncodeForHtmlAttribute(prm)%>"><%=enc.EncodeForHtmlAttribute(rstSelect("descripcion"))%></option><%
                                        end if
				                    end if
				    
				                    rstSelect.movenext
			                    wend
                            %></select>

			            <% '''DGB 21/10/2008  firma electronica
			            si_tiene_modulo_facturaElectronica=ModuloContratado(session("ncliente"),ModFirmaElectronica)
			            ''' ASP 31/01/2011 firma Factura MX
		                if si_tiene_modulo_facturaElectronica <> 0 then
		                    ExisteFolioMejico(oculta)
		                    if mejico then
		                        if existeFirma = 0 then
		                                pagina="../netInic.asp?pag=/custom/Mexico/cronos/FactEMX.aspx&nfact=" +p_nfactura+"&f="&rst("fecha")&"&c="&rst("ncliente")&"&url=" & personalizacionEmail
                            %><span class="CELDARIGHT" ><a class='CELDAREFB' href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(pagina)%>'+ document.facturas_cli.formato_impresion.value.replace('?','&') + 'novei=<%=enc.EncodeForJavascript(novei)%>','P','600','800')" OnMouseOver="self.status='<%=LitFactElectr%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitFirmar%></a></span><%
                             end if
		                    else
		                        '''Si es español aun no se controla
		                        pagina="../netInic.asp?pag=/cronos/fact_f.aspx&nfact=" +p_nfactura+"&f="&rst("fecha")&"&c="&rst("ncliente")&"&url=" & personalizacionEmail
                                %><span class="CELDARIGHT" width="25"><a class='CELDAREFB' href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(pagina)%>'+ document.facturas_cli.formato_impresion.value.replace('?','&') + 'novei=<%=enc.EncodeForJavascript(novei)%>','P','600','800')" OnMouseOver="self.status='<%=LitFactElectr%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitFirmar%></a></span><%
                                end if
		                '''ASP
			            end if
			            pagina="../crearpdf.asp?destinatario=" & rst("ncliente") & "&ndoc=" & rst("nfactura") & "&tdoc=FACTURA&dedonde=DOCUMENTOV&empresa=" & session("ncliente") & "&cajaParam=" & session("f_caja") & "&mode=DOC&url=" & personalizacionEmail
                        %><span class="CELDARIGHT" ><a class='CELDAREFB' href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(pagina)%>' + document.facturas_cli.formato_impresion.value.replace('?','&') + 'nfactura=<%="(\'"+enc.EncodeForJavascript(p_nfactura)+"\')"%>&novei=<%=enc.EncodeForJavascript(novei)%>','A','<%=enc.EncodeForJavascript(AltoVentana)+100%>','950')" OnMouseOver="self.status='<%=LitEnvEmail%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgEnviarEmail)%>" <%=ParamImgEnviarEmail%> alt="<%=ucase(LitEnvEmail)%>" title="<%=ucase(LitEnvEmail)%>"></a></span><%
                            ''MPC 18/10/2011 Se pone la condición if false para que no salga en ningún caso
                        'if si_tiene_modulo_29 <>0 or si_tiene_modulo_30 <>0 then
                        if false then
 				            pagina="../crearpdf.asp?enviar=FAX&destinatario=" & rst("ncliente") & "&ndoc=" & rst("nfactura") & "&tdoc=FACTURA&dedonde=DOCUMENTOV&empresa=" & session("ncliente") & "&cajaParam=" & session("f_caja") & "&mode=DOC&url=ventas/"
                            %><span class="CELDARIGHT"><a class='CELDAREFB' href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(pagina)%>' + document.facturas_cli.formato_impresion.value + 'nfactura=<%="(\'"+enc.EncodeForJavascript(p_nfactura)+"\')"%>&novei=<%=enc.EncodeForJavascript(novei)%>','A','<%=enc.EncodeForJavascript(AltoVentana)-200%>','600')" OnMouseOver="self.status='<%=LitEnvFax%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitFAX%></a></span><%
                             end if
                        ''FIN MPC 18/10/2011
                        CloseDiv
 		            end if
                    'hasta aqui
                    '   GPD 26/04/2007.
                    'strSerie = d_lookup("serie", "facturas_cli", "nfactura like '" & p_nfactura & "'", session("dsn_cliente"))

                    strSerieSelect="select serie from facturas_cli with(nolock) where nfactura like ?"
                    strSerie=DlookupP1(strSerieSelect, p_nfactura&"", adVarchar, 20, session("dsn_cliente"))

                    'strCif = d_lookup("empresa", "series", "nserie like '" & strSerie & "'", session("dsn_cliente"))

                    strCifSelect="select empresa from series with(nolock) where nserie like ?"
                    strCif=DlookupP1(strCifSelect, strSerie&"", adVarchar, 10, session("dsn_cliente"))

                    'bolActivoValesDTO = d_lookup("ACTIVARVALEDTO", "EMPRESAS", "CIF like '" & strCif & "'", session("dsn_cliente"))

                    bolActivoValesDTOSelect="select ACTIVARVALEDTO from empresas with(nolock) where cif like ?"
                    bolActivoValesDTO=DlookupP1(bolActivoValesDTOSelect, strCif&"", adVarchar, 25, session("dsn_cliente"))


                    If bolActivoValesDTO Then
                        DrawDiv "col-md-2 col-xxs-7 header-discount","",""
                        %><span class="CELDARIGHT"><%
                        %><a class='CELDAREFB' href="javascript:GenerarValeDescuento('<%=enc.EncodeForJavascript(p_nfactura)%>')" OnMouseOver="self.status='<%=LitGeneraValeDto%>'; return true;" OnMouseOut="self.status=''; return true;"><%= LitValeDto %></a><%
                        'descuento = d_lookup("IMPORTE","VALESDTO","FRAEMISION = '" & p_nfactura & "'",session("dsn_cliente"))

                        descuentoSelect="select importe from VALESDTO with(nolock) where FRAEMISION = ?"
                        descuento=DlookupP1(descuentoSelect, p_nfactura&"", adVarchar, 20, session("dsn_cliente"))

                        If Trim(descuento) = "" Then
                            %><div id="descuento" class="CELDAREFB"></div><%
                            %><iframe name="frDescuento" id="frDescuento" width="1" height="1" frameborder="no" scrolling="no" src="" style="visibility:visible;"></iframe><%
                            Else
                            %><div id="descuento" class="CELDAREFB"><br />(<%= enc.EncodeForHtmlAttribute(descuento) %>&nbsp;&euro;)</div><%
                            %><iframe name="frDescuento" id="frDescuento" width="1" height="1" frameborder="no" scrolling="no" src="" style="visibility:hidden;"></iframe><%
                            End If
                            %></span><%
                        CloseDiv
                    End If
			        rstSelect.close
	             else%>
	 	            <!--<span align="right">
		            </span>-->
	             <%end if

        'dgb: factura electronica enlazar el visor    17/06/2010
		if si_tiene_modulo_facturaElectronica <> 0 then
			if mode="browse" then
                DrawDiv "col-md-5 col-xxs-7 header-ebill","",""
			    if mejico then
			        if existeFirma = 1 then
                        %><span class="ENCABEZADOC" ><%
			                %><iframe id='Iframe2' name="fr_LFactura" src='documentosMX.aspx?cod=<%=enc.EncodeForHtmlAttribute(codigoFirma)%>&l=<%=LitLinkFirma %>&l1=<%=LitLinkFirma1%>&l2=<%=LitLinkFirma2 %>&l3=<%=LitLinkFirma3 %>' width='500' height='20' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
			                %></span><%
			        end if
			    else
                   %><span class="ENCABEZADOC" ><%
			            %><iframe id='frLFactura' name="fr_LFactura" src='linkFacturaE.asp?cod=<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>&l=<%=LitLinkFirma %>&l1=<%=LitLinkFirma1%>&l2=<%=LitLinkFirma2 %>&l3=<%=LitLinkFirma3 %>' width='500' height='20' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
			            %></span><%
			    end if
                CloseDiv
			end if
		end if
        %></div><%

        pagado=0
        if not rst.eof then
            if rst2.state<>0 then rst2.close
            rst2.cursorlocation=3
			rst2.open "SELECT nfactura, nvencimiento FROM VENCIMIENTOS_SALIDA with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "' and importecob>0",session("dsn_cliente")
			if not rst2.eof then
				pagado=1
			end if
			rst2.close
        else
            pagado=0
        end if
		if pagado=1 then
			if mode="edit" then%>
			<script language="javascript" type="text/javascript">
                    venci_paga.style.display = "";
			</script>
			<%end if
		end if
        %>
        <input type="hidden" size=3 name="vpagada" value="<%=enc.EncodeForHtmlAttribute(pagado)%>" />
		    <%
            actionVersion pagina, ImgNoCollapse, AltoVentana, AnchoVentana

                'Mostrar la barra de pestañas

		    BarraNavegacion mode%>
        <!--
            <div id="CollapseSection"> 
                <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['DatosGenerales','CABECERA', 'DIRENVIO', 'DATFINAN','DATTOTAL']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
                <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['DatosGenerales','CABECERA', 'DIRENVIO', 'DATFINAN','DATTOTAL']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
            </div>
        -->
        <table style="width: 100%;"></table>

        <div class="Section" id="S_DatosGenerales">
            <a href="#" rel="toggle[DatosGenerales]" data-openimage="<%=enc.EncodeForHtmlAttribute(ImgNoCollapse) %>" data-closedimage="<%=enc.EncodeForHtmlAttribute(ImgCollapse) %>">
                <div class="SectionHeader <%=enc.EncodeForHtmlAttribute(iif(mode="add" or mode="edit","displayed",""))%>">
                    <%=LITCABECERA%>
                    <img class="btn_folder" src="<%=enc.EncodeForHtmlAttribute(ImgNoCollapse) %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display: <%=iif(mode="add" or mode="edit","","none")%>;" id="DatosGenerales">
         <!--<center>-->

		    <table width="100%" border='<%=enc.EncodeForHtmlAttribute(borde)%>' cellpadding="1" cellspacing="1">
		        <%
                ''if si_tiene_modulo_mantenimiento<>0 then
                ''    DrawFila color_blau
                ''end if
				    Formulario="facturas_cli"
				    if mode="browse" then
				    else
					    if mode="add" and vienenp>"" then
					        docs=split(vienenp,"-") 
  		                    ninv=docs(0)
  		                    if Ubound(docs)>0 then  nven=int(docs(1))
                        
                            'tmp_ncliente = d_lookup("ncliente","facturas_cli","nfactura='" & ninv & "'",session("dsn_cliente"))

                            tmp_nclienteSelect="select ncliente from facturas_cli with(nolock) where nfactura= ?"
                            tmp_ncliente=DlookupP1(tmp_nclienteSelect, ninv&"", adVarchar, 20, session("dsn_cliente"))

                            rstAux.cursorlocation=3
                            rstAux.open "select * from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & tmp_ncliente & "'",session("dsn_cliente")
                            if not rstAux.eof then
					            tmp_forma_pago=rstAux("fpago")
					            tmp_portes=rstAux("PORTES")
					            tmp_transportista=rstAux("TRANSPORTISTA")
					            tmp_tipo_pago=rstAux("tpago")
					            ''tmp_fechaenvio=rstAux("PORTES")
					            ''tmp_nenvio=rstAux("PORTES")
					            tmp_tarifa=rstAux("tarifa")
					            ''tmp_fechapedido=rstAux("PORTES")
					            ''tmp_documento=rstAux("PORTES")
					            ''tmp_contacto=rstAux("CONTACTO")
					            tmp_banco=rstAux("banco")					      
					            tmp_ncuenta=rstAux("ncuenta")
                                tmp_ncuenta1=Mid(rstAux("ncuenta"), 1, 2)
                                tmp_ncuenta2=Mid(rstAux("ncuenta"), 3, 2)
					            tmp_ncuenta3=Mid(rstAux("ncuenta"), 5, 4)
				                tmp_ncuenta4=Mid(rstAux("ncuenta"), 9, 4)
				                tmp_ncuenta5=Mid(rstAux("ncuenta"), 13, 2)
	   			                tmp_ncuenta6=Mid(rstAux("ncuenta"), 15, Len(rstAux("ncuenta"))-14)
					            tmp_agenteasignado=rstAux("agente")
					            tmp_cod_proyecto=rstAux("proyecto")
					            ''tmp_incoterms=rstAux("PORTES")					
                                if ninv & "">"" then    
					                tmp_observaciones=LITOBSNP &" "&trimCodEmpresa(ninv)	
                                else
                                    tmp_observaciones=LITOBSNP
                                end if
					        else
                                tmp_forma_pago=""
					            tmp_portes=""
					            tmp_transportista=""
					            tmp_tipo_pago=""
					            ''tmp_fechaenvio=""
					            ''tmp_nenvio=""
					            tmp_tarifa=""
					            ''tmp_fechapedido=""
					            ''tmp_documento=""
					            ''tmp_contacto=""
					            tmp_banco=""
					            tmp_ncuenta=""
					            tmp_ncuenta1= ""
				                tmp_ncuenta2= ""
				                tmp_ncuenta3= ""
	   			                tmp_ncuenta4= ""
                                tmp_ncuenta5= ""
                                tmp_ncuenta6= ""
                                ncuenta1= ""
				                ncuenta2= ""
				                ncuenta3= ""
	   			                ncuenta4= ""
                                ncuenta5= ""
                                ncuenta6= ""
					            tmp_agenteasignado=""
					            tmp_cod_proyecto=""
					            ''tmp_incoterms=""
                                if ninv & "">"" then    
					                tmp_observaciones=LITOBSNP &" "&trimCodEmpresa(ninv)	
                                else
                                    tmp_observaciones=LITOBSNP
                                end if
                            end if
					        rstAux.close  		                			    

                        end if

                    end if
            if esto_ahora_no=1 then
            %>
            <tr>
                <td style="width:0px; height:0px;"></td>
                <td style="width:0px; height:0px;"></td>
                <td style="width:0px; height:0px;"></td>
                <td style="width:0px; height:0px;"></td>
                <%if mode="add" or mode="edit" then %>
                <td style="width:0px; height:0px;"></td>
                <td style="width:50%; height:0px;"></td>
                <%end if 
            end if
            %>
            </tr>
                <!--<tr>
                    <td>
                    <table cellspacing="0" cellpadding="1" style="table-layout:fixed">
                        <tr>--><%
                            DrawDiv "1", "", ""
				            if mode="browse" then
					            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitSerie + ":"
                                DrawLabel "", "", LitSerie
					            'DrawCelda "CELDA","","",0,trimCodEmpresa(rst("serie"))
                                DrawSpan "CELDA", "", trimCodEmpresa(rst("serie")), ""
					            'mmg:
					            rstMM.open "select almacen from series with(nolock) inner join almacenes alm with(nolock) on alm.codigo=almacen where nserie='"&rst("serie")&"' and alm.codigo like '"&session("ncliente")&"%' and isnull(alm.fbaja,'')=''",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	                            if not rstMM.EOF then
		                            almacenSerie= rstMM("almacen")
	                            else
		                            almacenSerie= ""
	                            end if
	                            rstMM.close
				            else
					            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitSerie + ":" 
                                DrawLabel "", "", LitSerie '& " <span class='txt-primary1'>*</span>"
					            ''ricardo 7-12-2004 solamente se podran ver las series permitidas al usuario
					            strSacSerie="select nserie,almacen, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='FACTURA A CLIENTE' and nserie like '" & session("ncliente") & "%'"
					            if s & "">"" then
						            strSacSerie=strSacSerie & " and nserie in " & s
					            end if
					            strSacSerie=strSacSerie & " order by nserie"
                                rstAux.cursorlocation=3
					            rstAux.open strSacSerie,session("dsn_cliente")
					            if mode="add" then
						            'DrawSelectCeldaSpan "CELDA","200","",0,LitSerie,"serie",rstAux,iif(p_serie>"",p_serie,""),"nserie","descripcion","onchange","javascript:document.facturas_cli.h_ncliente.value=document.facturas_cli.ncliente.value;document.facturas_cli.ncliente.value='';TraerCliente('add','2');",2
                                    'DrawSelect "100", "", "serie", rstAux, iif(p_serie>"",p_serie,""),"nserie","descripcion","onchange","javascript:document.facturas_cli.h_ncliente.value=document.facturas_cli.ncliente.value;document.facturas_cli.ncliente.value='';TraerCliente('add','2');"
                                    DrawSelect1 "", "", "serie", "serie", rstAux, enc.EncodeForHtmlAttribute(iif(p_serie>"",p_serie,"")), "nserie", "descripcion", "onchange", "javascript:document.facturas_cli.h_ncliente.value=document.facturas_cli.ncliente.value;document.facturas_cli.ncliente.value='';TraerCliente('add','2');", "", ""
                                else
						            'DrawSelectCeldaSpan "CELDA","200","",0,LitSerie,"serie",rstAux,iif(p_serie>"",p_serie,rst("serie")),"nserie","descripcion","","",2
			 		                'DrawSelect "100", "", "serie", rstAux, iif(p_serie>"",p_serie,rst("serie")),"nserie","descripcion","",""
                                    DrawSelect1 "", "", "serie", "serie", rstAux, enc.EncodeForHtmlAttribute(iif(p_serie>"",p_serie,rst("serie"))), "nserie", "descripcion", "", "", "", "" 
                                end if
			 		            rstAux.close
				            end if
                            CloseDiv
				         %><!--</tr>
                         </table>
                    </td>-->
                    <%
                    'DrawCelda "","","",0,LitCobrada+":"
				    '<td><table border='0' cellspacing="0" cellpadding="0" width="100%" class="section-table">
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitCobrada
				        'DrawFila color_blau			
                            if mode="add" then
				                'EligeCelda "check", mode,"","0","",0,LitCobrada,"cobrada",0,iif(p_cobrada>"",nz_b(p_cobrada),0)				                   	
                                EligeCeldaResponsive1 "check", mode, "", "", "cobrada", iif(p_cobrada>"",nz_b(p_cobrada),0), LitCobrada
                            else
                                'EligeCelda "check", mode,"","0","",0,LitCobrada,"cobrada",0,iif(p_cobrada>"",nz_b(p_cobrada),rst("cobrada"))
                                EligeCeldaResponsive1 "check", mode, "CELDA", "", "cobrada", iif(p_cobrada>"",nz_b(p_cobrada),enc.EncodeForHtmlAttribute(rst("cobrada"))), LitCobrada
                            end if
				            '*** AMP Gestión de impagos.
				            nvenc=""
				            'nonpayment = d_lookup("nonpayment","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))

                            nonpaymentSelect="select nonpayment from configuracion with(nolock) where nempresa = ?"
                            nonpayment= DlookupP1(nonpaymentSelect, session("ncliente")&"", adChar, 5, session("dsn_cliente"))

                            if mode="add" then
                                nvenc=""
                            else			      				        
				                'nvenc=d_lookup("nfactura","vencimientos_salida","nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))			      				        

                                nvencSelect="select nfactura from vencimientos_salida with(nolock) where nfactura like ?+'%' and nfactura= ?"
                                nvenc=DlookupP2(nvencSelect, session("ncliente")&"", adVarchar, 20, rst("nfactura")&"", adVarchar, 20, session("dsn_cliente"))

                            end if
                            if mode="edit" then
				                if rst("cobrada")<>0 and mode="edit" and nz_b2(nonpayment)=1 and nvenc="" then
				           	        typedoc=0		  
					                   '<td class="CELDA">         
			                        %>
                                            <a class="CELDAREFB7" href="#fr_NonPayment" onmouseover="self.status='Anotar Impago'; return true;" onmouseout="self.status=''; return true;" onclick="SetImpago('#fr_NonPayment','<%=enc.EncodeForJavascript(cstr(rst("nfactura")))%>','<%=enc.EncodeForJavascript(rst("ncliente"))%>','<%=enc.EncodeForJavascript(typedoc)%>')">
                                                <img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgAnotar)%>" <%=ParamImgAnotar%> alt="<%=LITNPNOTE%>" title="<%=LITNPNOTE%>" />
                                            </a>
				                   <%                                        
			                           '</td>
				                end if
                            end if

				            pagado=0	
                            if mode="add" then
                                %><input type="hidden" name="h_cobrada" value="0"/><%				
                            else								        
        		                %><input type="hidden" name="h_cobrada" value="<%=iif(rst("cobrada")<>0,"1","0")%>"/><%				
                            end if
				            if mode="add" then
                                pagado=0
                            else
				                if rst2.state<>0 then rst2.close
                                rst2.cursorlocation=3
				                rst2.open "SELECT nfactura, nvencimiento FROM VENCIMIENTOS_SALIDA with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "' and importecob>0",session("dsn_cliente")
				                if not rst2.eof then
					                pagado=1
				                end if
				                rst2.close
                            end if
				            if pagado=1 then%>
					            <input type="hidden" name="h_vpagada" value="1"/>
				            <%else%>
					            <input type="hidden" name="h_vpagada" value="0"/>
				            <%end if
				        'CloseFila
                        CloseDiv
				    '</table></td>
                    ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                '</tr>
                
				    campo="codigo"
				    if mode="browse" then
					    campo2="abreviatura"
				    else
					    campo2="abreviatura"
				    end if
				    '*** AMP 15102010			
				    if tmpdivisafc>"" then  tmp_divisa = tmpdivisafc end if
                    if mode="add" then
                        DIVISA=iif(tmp_divisa>"",tmp_divisa,"")
                    else
				        DIVISA=iif(tmp_divisa>"",tmp_divisa,rst("divisa"))
                    end if
				    %><!--<td colspan=2 style="width:220px">
                        <table cellspacing="0" cellpadding="1"  style="table-layout:fixed">
                            <tr>--><%
				                'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitDivisa+":"
                                dato_celda=Desplegable(mode,campo,campo2,"divisas",DIVISA,"moneda_base<>0 and codigo like '" & session("ncliente") & "%'")

				                if mode<>"browse" then
                                abreviSelect="select abreviatura from divisas with(nolock) where codigo= ? and codigo like ?+'%'"
                                'd_lookup("abreviatura","divisas","codigo='" & tmp_divisa & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                                'd_lookup("abreviatura","divisas","codigo='" & dato_celda & "' and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
					                datoDivisa=iif(tmp_divisa>"", _
						                DlookupP2(abreviSelect, tmp_divisa&"", adVarchar, 15, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")), _
						                DlookupP2(abreviSelect, dato_celda&"", adVarchar, 15, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))
				                else
					                datoDivisa=dato_celda
				                end if
                                ''ricardo 18/8/2004 , se podra cambiar la divisa al documento si no tiene detalles ni conceptos
                                if mode<>"edit" then
	                                estilo_divisa="CELDA"
                	                if mode="add" then
		                                tipo_eligecelda="select"
	                                else
		                                tipo_eligecelda="input"
	                                end if
                                else
	                                cuantos_detalles=d_count("item","detalles_fac_cli","nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
	                                cuantos_conceptos=d_count("nconcepto","conceptos","nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
	                                if cint(cuantos_detalles) + cint(cuantos_conceptos)>0 then
		                                estilo_divisa="CELDA DISABLED colspan='2'"
		                                tipo_eligecelda="input"
	                                else
		                                estilo_divisa="CELDA"
		                                tipo_eligecelda="select"
	                                end if
                                end if
				                if mode="add" or mode="edit" then RstAux.close
				                    Estilo=iif(mode="browse","CELDA",estilo_divisa)
				                    if tipo_eligecelda="input" then
                                        if mode= "edit" then
                                            DrawInputCeldaDisabled "", "", "", 5, 0, LitDivisa, "divisa", datoDivisa
                                        elseif mode= "browse" then
                                            EligeCeldaResponsive "text",mode,"CELDA","","",0,LitDivisa,"divisa",5,datoDivisa
                                        end if
				                else
				                    'monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))

                                    monedaBaseSelect="select codigo from divisas with(nolock) where codigo like ? +'%' and moneda_base='1'"
                                    monedaBase=DlookupP1(monedaBaseSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))

                                    rstAux.cursorlocation=3
					                rstAux.open "select codigo,abreviatura as descripcion from divisas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
		 			                DrawSelectCelda "CELDA","","",0,LitDivisa,"divisa",rstAux,iif(mode="browse",DIVISA,iif(tmp_divisa>"",tmp_divisa,dato_celda)),"codigo","descripcion","onchange","javascript:cambiardivisa('"&monedaBase&"');"
			 		                'DrawSelect1 "selectmenu", "", "divisa", "divisa", rstAux, iif(mode="browse",DIVISA,iif(tmp_divisa>"",tmp_divisa,dato_celda)),"codigo","descripcion","onchange","javascript:cambiardivisa('"&monedaBase&"');", "", LitDivisa
                                    rstAux.close
				                end if
                                %>
				            <!--</tr>
                        </table>
                    </td>-->
				    <input type="hidden" name="h_divisa" value="<%=enc.EncodeForHtmlAttribute(iif(tmp_divisa>"",tmp_divisa,dato_celda))%>"/>
				    <input type="hidden" name="divisafc" value="<%=enc.EncodeForHtmlAttribute(iif(tmp_divisa>"",tmp_divisa,DIVISA))%>"/>	
			    <%

                    'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitContabilizada+":"
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitContabilizada

			        '**RGU 14/2/2007**
			        if si_tiene_modulo_contabilidad<>0 and mode="browse" then
                    
                        if rst("contabilizado")=true then
				            'TmpCif=d_lookup("empresa","series","nserie like '" & session("ncliente") & "%' and nserie='"&rst("serie")&"' ",session("dsn_cliente"))

                            TmpCifSelect="select empresa from series wirh(nolock) where nserie like ?+'%' and nserie =?"
                            TmpCif=DlookupP2(TmpCifSelect, session("ncliente")&"", adVarchar, 10, rst("serie")&"", adVarchar, 10, session("dsn_cliente"))

				            TmpCif=trimcodempresa(TmpCif)
				            TmpCif=LimpiarCIF(TmpCif)

                            dateInvoice=rst("fecha")
				            any=year(cstr(rst("fecha")))&""

                            ''Ricardo 15-10-2014 se cambia la manera de obtener el ejercicio activo de contabilidad
				            'ega 19/06/2008 union de tablas con join
  				            ''selectConta=" select c.nempresa from configconta c with(nolock) inner join clientes cli with(nolock) on cli.ncliente=c.ncliente"
  				            ''selectConta=selectConta&" where c.nempresa like '"&session("ncliente")&"%' "
  				            ''selectConta=selectConta&" and cli.ncliente like '"&session("ncliente")&"%' and cli.cifedi='"&TmpCif&"' "
                            ''por esta otra select
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
                               'ega 19/06/2008 with(nolock) y union de tablas con join
                                rstConta.cursorlocation=3
					            rstConta.open "select distinct a.nasiento, a.fecha from detalles_asientos"&nempresaconta&" d with(nolock) inner join asientos"&nempresaconta&" a with(nolock) on a.nasiento=d.nasiento where d.nfacturacli='"&rst("nfactura")&"' ",session("dsn_cliente")
					            if not rstConta.eof then
						            rstAux.cursorlocation=3
						            rstAux.open "select c.NEMPRESA,c.EJERCICIO,p.FINICIO,p.FFIN from configcontaactivo as c with(nolock) inner join configconta as p with(nolock) on c.nempresa=p.nempresa and c.ejercicio=p.ejercicio and c.nempresa='"&nempresaconta&"' and  c.ncliente = '" & session("ncliente") & "' and c.usuario='"&session("usuario")&"' ",session("dsn_cliente")
						            if not rstAux.eof then
							            TmpEjercActivo=rstAux("ejercicio")
                                         TmpFinit=rstAux("finicio")
                                        TmpFfin=rstAux("ffin")
							            TmpContaActivo=rstAux("nempresa")
						            end if
						            rstAux.close
                                    
						            if (TmpContaActivo&""=nempresaconta and (cdate(dateInvoice)>=cdate(TmpFinit) and cdate(dateInvoice)<=cdate(TmpFfin))) then
							            VinculosPagina(MostrarAsiento)=1:VinculosPagina(MostrarSubcuenta)=1
							            CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina%>
							            <span class="CELDA <%=enc.EncodeForHtmlAttribute(texto_bcc)%> " width='0%' colspan="2"> <%=Visualizar(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado")))& " ("&null_s(rstConta("fecha"))&LitNAsiento&Hiperv(OBJAsiento,rstConta("nasiento")&"&s="&nempresaconta,"browse","facturas_cli",Permisos,Enlaces,session("usuario"),session("ncliente"),rstConta("nasiento"),LitVerAsiento)&")"%></span>
							            <%nasiento=null_s(rstConta("nasiento"))
						            else%>
							            <span class="CELDA <%=enc.EncodeForHtmlAttribute(texto_bcc)%> " width='0%' colspan="2"> <%=Visualizar(iif(tmp_contabilizada>"",nz_b(tmp_contabilizada),rst("contabilizado")))& " ("&null_s(rstConta("fecha"))&LitNAsiento&rstConta("nasiento")&")"%></span>
						            <%end if
					              else
					  	            EligeCeldaResponsive1 "check", mode, "CELDA" & texto_bcc, "", "contabilizada", iif(p_contabilizada>"",nz_b(p_contabilizada),rst("contabilizado")), LitContabilizada
                                end if
                                rstConta.close
				            else
					            EligeCeldaResponsive1 "check", mode,"CELDA" & texto_bcc,"", "contabilizada",iif(p_contabilizada>"",nz_b(p_contabilizada),rst("contabilizado")), LitContabilizada
				            end if
			            else
				            EligeCeldaResponsive1 "check", mode,"CELDA" & texto_bcc,"", "contabilizada",iif(p_contabilizada>"",nz_b(p_contabilizada),rst("contabilizado")), LitContabilizada
			            end if
                    else
                        if mode="add" then
                            EligeCeldaResponsive1 "check", mode,"CELDA" & texto_bcc,"", "contabilizada",iif(p_contabilizada>"",nz_b(p_contabilizada),0), LitContabilizada
                        else
                            EligeCeldaResponsive1 "check", mode,"CELDA" & texto_bcc,"", "contabilizada",iif(p_contabilizada>"",nz_b(p_contabilizada),rst("contabilizado")), LitContabilizada
                        end if
                    end if
                    CloseDiv
			        %><input type="hidden" name="nasiento" value="<%=enc.EncodeForHtmlAttribute(nasiento)%>"/>
                    <%if mode="add" then %>
			            <input type="hidden" name="h_contabilizada" value="0"/>
                    <%else %>
                        <input type="hidden" name="h_contabilizada" value="<%=iif(rst("contabilizado")<>0,"1","0")%>"/>
                    <%end if %>
			        <%
                    if mode<>"add" and mode <>"edit" then%>
			            <!--<td>&nbsp;</td>-->
			            <%
			            if saft=True and rst("ahora") then
				            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitBloqueada
                            DrawDiv "1", "", ""
                            DrawLabel "", "width:120px", LitBloqueada
			            end if
			            if saft=True and rst("ahora")=0 then
				            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitLibre
                            DrawDiv "1", "", ""
                            DrawLabel "", "width:120px", LitLibre
			            end if%>
                    
                        <%if mode="add" then %>
                            <input type="hidden" name="h_ahora" value="0"/>
                        <%else %>
                            <input type="hidden" name="h_ahora" value="<%=iif(rst("ahora")<>0,"1","0")%>"/>
                        <%end if %>

                        <%if mode<>"add" then
                            if pnf="1" and rst("ahora")=0 then 'permitido bloquear, factura libre%>
					            <span class="CELDARIGHT" style="width:30px"><a href="javascript:bloqueoFactura('<%=enc.EncodeForJavascript(rst("nfactura"))%>')" onmouseover="self.status='<%=LitBloqueoFact%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgBloqueoFacturas)%>" <%=ParamImgBloqueoFacturas%> alt="<%=LitBloqueoFact%>" title="<%=LitBloqueoFact%>"/></a></span>
				            <%
                            end if
				            if rnf="1" and rst("ahora") then 'permitido desbloquear, factura bloqueada %>
					            <span class="CELDARIGHT" style="width:30px"><a href="javascript:desbloqueoFactura('<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>')" onmouseover="self.status='<%=LitDesbloqueoFact%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgValidar)%>" <%=ParamImgBloqueoFacturas%> alt="<%=LitDesbloqueoFact%>" title="<%=LitDesbloqueoFact%>"/></a></span>
				            <%
                            end if
                        else

                        end if
                                
			            if saft=True and rst("ahora") then
                            CloseDiv
			            elseif saft=True and rst("ahora")=0 then
                            CloseDiv
			            end if
                    end if
			
			        '*** AMP Añadimos campo factor de cambio en la cabecera.
			        'monedaBase = d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))

                    monedaBaseSelect="select codigo from divisas with(nolock) where codigo like ?+'%' and moneda_base='1'"
                    monedaBase=DlookupP1(monedaBaseSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))

   	                'abrevBase =  d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and moneda_base='1'",session("dsn_cliente"))   	
                                
                    abrevBaseSelect="select abreviatura from divisas with(nolock) where codigo like ?+'%' and moneda_base='1'"
                    abrevBase=DlookupP1(abrevBaseSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))

                    if mode="add" then
                        factcambio = ""
                    else
	                    'factcambio = d_lookup("factcambio","facturas_cli","nfactura='" & rst("nfactura") & "' and nfactura like '" & session("ncliente") & "%'",session("dsn_cliente"))
                        
                        factcambioSelect="select factcambio from facturas_cli with(nolock) where nfactura=? and nfactura like ?+'%'"
                        factcambio=DlookupP2(factcambioSelect, rst("nfactura")&"", adVarchar, 20, session("ncliente")&"", adVarchar, 20, session("dsn_cliente"))

                    end if
                    if mode="browse" then 
                        if  rst("divisa")<>monedaBase then 
                            
                            
                            DrawDiv "1", "", ""
                            DrawLabel "", "", LitFactCambio

                                    %><!--<td colspan="2"><table width="100%" cellspacing="0" cellpadding="1" style="table-layout:fixed"><tr>--><%
	                                'abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&rst("divisa")&"'",session("dsn_cliente"))

                                    abreviaAtDivSelect="select abreviatura from divisas with(nolock) where codigo like ? +'%' and codigo = ?"
                                    abreviaAtDiv=DlookupP2(abreviaAtDivSelect, session("ncliente"), adVarchar, 15, rst("divisa")&"", adVarchar, 15, session("dsn_cliente"))

                                    'DrawCelda "ENCABEZADOL style='width:130px'","","",0,LitFactCambio+" :"
	                                ''DrawCelda "CELDA","","",0,CStr(factcambio)+" "+abrevBase
	                                %>
                                    <span class="CELDA" style="width:200px" >
                                        <%
                                            ''response.write(CStr(factcambio) & " " & abrevBase)
                                            response.write("1" & abrevBase & " = " & CStr(factcambio) & abreviaAtDiv)
                                        %>
                                    </span>
                                    <%
                            CloseDiv
	                    end if
	                else
                        ocultar=0
                        if mode="add" or mode="edit" then	  
                            if mode="add" then              	        
                                'DIVISA:valor de divisa predefinido en el cliente o en una serie con cliente predeterminado
                                'dato_celda: cuando esta vacio moneda_base.              	        
                                dv=iif(DIVISA>"",DIVISA,dato_celda)              	         
                                'factcambio = d_lookup("factcambio","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dv&"'",session("dsn_cliente"))
                                        
                                factcambioSelect="select factcambio from divisas with(nolock) where codigo like ?+'%' and codigo = ?"
                                factcambio=DlookupP2(factcambioSelect, session("ncliente")&"", adVarchar, 15, dv&"", adVarchar, 15, session("dsn_cliente"))

                                'abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dv&"'",session("dsn_cliente"))

                                abreviaAtDivSelect="select abreviatura from divisas with(nolock) where codigo like ?+'%' and codigo = ?"
                                abreviaAtDiv=DlookupP2(factcambioSelect, session("ncliente")&"", adVarchar, 15, dv&"", adVarchar, 15, session("dsn_cliente"))

                                if dv=monedaBase then ocultar=1 end if
                            else 'modo edit
                                dvEdit=iif(tmp_divisa>"",tmp_divisa,dato_celda)
                                'abreviaAtDiv = d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='"&dvEdit&"'",session("dsn_cliente"))

                                abreviaAtDivSelect="select abreviatura from divisas with(nolock) where codigo like ?+'%' and codigo = ?"
                                abreviaAtDiv=DlookupP2(factcambioSelect, session("ncliente")&"", adVarchar, 15, dvEdit&"", adVarchar, 15, session("dsn_cliente"))

                                if dvEdit=monedaBase then ocultar=1 end if              	        
                            end if
                            
                            

                            DrawDiv "1", "", "tdfactcambio"
                            DrawLabel "", "", LitFactCambio
                            DrawSpan "CELDA", "", "1" & abrevBase & " = ", ""
                            %>
                            <input type="text" name="nfactcambio" value="<%=enc.EncodeForHtmlAttribute(CStr(factcambio))%>" size="6" style="text-align:right" onchange="comprobarFactorCambio()"/>
                            <span id="idfactcambioexpl"><%=abreviaAtDiv%></span>
                            <%

                            CloseDiv
                            if ocultar=1 then %><script language="javascript" type="text/javascript">parent.pantalla.document.getElementById("tdfactcambio").style.display = "none"</script><% end if	          
                            ''end if     
	                    end if
	                    '*** f AMP
                    end if
                    ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                    ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                
                'JFT 23/04/2012 Add fields VALIDATED and VALIDATED_BY
                'DrawFila color_blau
                '    if mode<>"add" then
                '        DrawCelda "ENCABEZADOL style='width:70px'","","",0,LitValidated&":"
                '        if rst("validated") = true then
                '            EligeCelda "check", "browse","CELDA","0","",0,"","validated",0,rst("validated")
                '            VinculosPagina(MostrarPersonal)=1
				'			CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
                '            %>
                <!--            <td class="CELDA">
                                <%
                                'Hiperv(OBJPersonal,rst("validated_by"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),d_lookup("nombre","personal","dni='" & rst("validated_by") & "'",session("dsn_cliente")) + " " & d_lookup("surname1","personal","dni='" & rst("validated_by") & "'",session("dsn_cliente")) + " " & d_lookup("surname2","personal","dni='" & rst("validated_by") & "'",session("dsn_cliente")),"")
                                %>
                            </td>--><%
                            'DrawCelda "CELDA","","",0,d_lookup("nombre","personal","dni='" & rst("validated_by") & "'",session("dsn_cliente")) + " " & d_lookup("surname1","personal","dni='" & rst("validated_by") & "'",session("dsn_cliente")) + " " & d_lookup("surname2","personal","dni='" & rst("validated_by") & "'",session("dsn_cliente"))
                '        else
                '            EligeCelda "check", mode,"CELDA","0","",0,"","validated",0,rst("validated")
                '        end if
                '    end if
                'CloseFila
                %>
            <!--</table>-->
	
	  

	    <%
 
	    if (mode="browse" or mode="edit" or mode="add") then
		
            %>
		    <!--<table border='<%=borde%>' cellspacing="1" cellpadding="1" width="100%">-->
			    <%
            
            '''ASP 
            if mode<>"add" then
		        if trim(p_ncliente)="" then p_ncliente=rst("ncliente")
            else
                ''p_ncliente=""
            end if
            %><!--<td width="0%" colspan="10">--><%
            if mode="add" then
                %>
		        <input type="hidden" name="h_ncliente" value=""/>
		        <input type="hidden" name="h_nfactura" value=""/>
		        <input type="hidden" name="olddivisa" value=""/>
		        <input type="hidden" name="estadoencaja" value=""/>
		        <input type="hidden" name="frabono" value="" />
		        <input type="hidden" name="gestbono" value="" />
		        <input type="hidden" name="ndetcon" value="" />
		        <%
            else
                %>
		        <input type="hidden" name="h_ncliente" value="<%=enc.EncodeForHtmlAttribute(rst("ncliente"))%>"/>
		        <input type="hidden" name="h_nfactura" value="<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>"/>
		        <input type="hidden" name="olddivisa" value="<%=enc.EncodeForHtmlAttribute(rst("divisa"))%>"/>
		        <input type="hidden" name="estadoencaja" value="<%=enc.EncodeForHtmlAttribute(estadoencaja)%>"/>
		        <input type="hidden" name="frabono" value="<%=enc.EncodeForHtmlAttribute(nz_b2(rst("frabono")))%>" />
		        <input type="hidden" name="gestbono" value="<%=enc.EncodeForHtmlAttribute(gestbono)%>" />
		        <input type="hidden" name="ndetcon" value="<%=null_z(d_count("item","detalles_fac_cli","nfactura='" & rst("nfactura") & "'",session("dsn_cliente")))+null_z(d_count("nconcepto","conceptos","nfactura='" & rst("nfactura") & "'",session("dsn_cliente")))%>" />
		        <%
            end if
            %><!--</td>--><%
            ''CloseFila

if esto_ahora_no=1 then
                    ''dgb  10-03-2008 comprobamos si es Centroxogo y tiene edi, mostramos posibilidad de exportar
				    no_mostrar=""
				    if mode="browse" then
				        if si_tiene_modulo_Centroxogo<>0 then
                                rstAux.cursorlocation=3
		                        rstAux.open "select edi from facturas_cli with(nolock) where nfactura like '" & session("ncliente") & "%' and nfactura='" & p_nfactura & "'", session("dsn_cliente")
	                            if not rstAux.eof then
	                                if rstAux("edi")&"">"" then
	                                    'si el proveedor de Centroxogo es Portugal
	                                    if mid(rstAux("edi"),6,5)="00184" then
						                    '<td class="ENCABEZADOC" colspan="5" >
                                            DrawDiv "1", "display:" & no_mostrar, "genfact"
						                        '<div id="genfact" style="display:=no_mostrar">
                                            %>
								                    <a class='CELDAREFB' href="javascript:AbrirVentana('../servicios/exportarFacturasCentroxogo.asp?mode=add&entrada=<%=enc.EncodeForJavascript(rstAux("edi"))%>','P', <%=enc.EncodeForJavascript(AltoVentana)-200%>, <%=enc.EncodeForJavascript(AnchoVentana)-100%>)"><%=LitExportarDoc%>  </a>
  						                    <%
						                        '</div>
                                            CloseDiv
					                        '</td>
                                        else
  						                    ''DrawCelda "ENCABEZADOC","70","",0,""
  						                end if
  						            else
  						                ''DrawCelda "ENCABEZADOC","70","",0,""
  						            end if
	                            else
	                                  ''DrawCelda "ENCABEZADOC","70","",0,""
	                            end if
                                rstAux.close
                                %><!--<td align="right" style="width:30%"></td>--><%
                            
	                    else
	                        ''DrawCelda "ENCABEZADOC","70","",0,""
				        end if
				    else
				        ''DrawCelda "ENCABEZADOC","70","",0,""
				    end if
end if


                %>
		    <!--</table>
               <table width='100%' border="0" cellspacing="1" cellpadding="1">-->
                <%

                    DrawDiv "3-sub", "background-color: #eae7e3", ""
                    %> 
                    <label class="ENCABEZADOC" style="text-align:left"><%=LITDAGRAL%></label>
                    <%
                    CloseDiv
                    '<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12" style="background-color: #ccd4e4"> 
                    '<label style="font-size: 16px">=LITDAGRAL</label>
                    '</div>
                %>
                 <!--</table>-->

		    <!--<table class=TDBORDE width="100%"><tr><td>-->

		        <!--<table width='100%' border='<%=borde%>' cellspacing="1" cellpadding="1">-->
		            <%
			            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitFormaPago+":"
			            DrawDiv "1", "", ""
                            DrawLabel "", "", LitFormaPago

                            campo="codigo"
			                campo2="descripcion"
                            if mode="add" then
                                dato_celda=Desplegable(mode,campo,campo2,"formas_pago",iif(tmp_forma_pago>"",tmp_forma_pago,""),"")
                            else
                                dato_celda=Desplegable(mode,campo,campo2,"formas_pago",iif(tmp_forma_pago>"",tmp_forma_pago,rst("forma_pago")),"")
                            end if

                            claseFormaPago = "CELDA"
                            if mode="add" or mode="edit" then
                                claseFormaPago = "selectmenu"
                            end if

			                EligeCeldaResponsive1 "select",mode,claseFormaPago,"","forma_pago",enc.EncodeForHtmlAttribute(dato_celda), ""

                            if mode="add" or mode="edit" then 
                                RstAux.close
                            end if
			            
                            if mode<>"browse" then 

                                if mode="add" then
                                    %><input type="hidden" name="forma_pago_ant" value=""/><%
                                else
                                    %><input type="hidden" name="forma_pago_ant" value="<%=rst("forma_pago")%>"/><%
                                end if
                            end if
                        CloseDiv

			            'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitPortes+":"
			            DrawDiv "1", "", ""
                        DrawLabel "", "", LitPortes
                        if mode="browse" then
				            EligeCeldaResponsive1 "", mode,"CELDA","","portes",iif(tmp_portes>"",tmp_portes,iif(isnull(rst("portes")),"",rst("portes"))), ""
			            else
                            if mode="edit" then
				                defecto=iif(tmp_portes>"",tmp_portes,rst("portes"))
                            else
                                defecto=iif(tmp_portes>"",tmp_portes,"")
                            end if
                            %><select name="portes" id="portes" class="selectmenu">
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
			            <%end if
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        CloseDiv

			            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitTransportista+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitTransportista

                        claseTransportista = "CELDA"
                        if mode="add" or mode="edit" then
                            claseTransportista = ""
                        end if

                        if mode="add" then
                            'EligeCelda "input", mode,"CELDA","","",0,LitTransportista,"transportista",27,iif(tmp_transportista>"",tmp_transportista,"")
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "", "", "transportista", enc.EncodeForHtmlAttribute(iif(tmp_transportista>"",tmp_transportista,"")), ""
                        else
                            if rst("transportista") & "" <>"" then
                               datosTransportista=enc.EncodeForHtmlAttribute(rst("transportista"))
                            else
                               datosTransportista=rst("transportista")
                            end if
                            'EligeCelda "input", mode,"CELDA","","",0,LitTransportista,"transportista",27,iif(tmp_transportista>"",tmp_transportista,iif(isnull(rst("transportista")),"",rst("transportista")))
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), claseTransportista, "", "transportista", enc.EncodeForHtmlAttribute(iif(tmp_transportista>"",tmp_transportista,iif(isnull(datosTransportista),"",datosTransportista))), ""
                        end if
			            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        CloseDiv


			        'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitTipoPago+":"
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitTipoPago

                    if mode="add" then
                        dato_celda=Desplegable(mode,campo,campo2,"tipo_pago",iif(tmp_tipo_pago>"",tmp_tipo_pago,""),"")
                    else
                        dato_celda=Desplegable(mode,campo,campo2,"tipo_pago",iif(tmp_tipo_pago>"",tmp_tipo_pago,rst("tipo_pago")),"")
                    end if

                    claseTipoPago = "CELDA"
                    if mode="add" or mode="edit" then
                        claseTipoPago = "selectmenu"
                    end if
			        'EligeCelda "select", mode,"CELDA style='width:200px;'",iif(mode<>"browse","200",""),"",0,LitTipoPago,"tipo_pago",15,dato_celda
			        EligeCeldaResponsive1 "select", enc.EncodeForHtmlAttribute(null_s(mode)), enc.EncodeForHtmlAttribute(claseTipoPago), "", "tipo_pago", dato_celda, ""
                    CloseDiv
                    if mode="add" or mode="edit" then RstAux.close
                    ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                    ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                    
			            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitFechaEnvio+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitFechaEnvio
                        claseFechaEnvio = "'CELDA width150px'"
                        if mode="add" or mode="edit" then
                            claseFechaEnvio = "'datepicker special-input  width150px'"
                        end if
                        if mode="add" then
                            'EligeCelda "input", mode,"CELDA","","",0,LitFechaEnvio,"fechaenvio",27,iif(tmp_fechaenvio>"",tmp_fechaenvio,"")
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), enc.EncodeForHtmlAttribute(claseFechaEnvio), "", "fechaenvio", iif(tmp_fechaenvio>"",tmp_fechaenvio,""), ""
                        else
                            'EligeCelda "input", mode,"CELDA","","",0,LitFechaEnvio,"fechaenvio",27,iif(tmp_fechaenvio>"",tmp_fechaenvio,iif(isnull(rst("fecha_envio")),"",rst("fecha_envio")))
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), enc.EncodeForHtmlAttribute(claseFechaEnvio), "", "fechaenvio", iif(tmp_fechaenvio>"",tmp_fechaenvio,iif(isnull(rst("fecha_envio")),"",rst("fecha_envio"))), ""
                        end if
                        if mode="add" or mode="edit" then DrawCalendar "fechaenvio" end if
                        CloseDiv
			            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                     

			        'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitNumEnvio+":"
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitNumEnvio

                    claseNumEnvio = "CELDA"
                    if mode="add" or mode="edit" then
                        claseNumEnvio = ""
                    end if

                    if mode="add" then
                        'EligeCelda "input", mode,"CELDA","","",0,LitNumEnvio,"nenvio",27,iif(tmp_nenvio>"",tmp_nenvio,"")
                        EligeCeldaResponsive1 "input", mode, "", "", "nenvio", iif(tmp_nenvio>"",tmp_nenvio,""), ""
                    else
                        if rst("nenvio") & "" <>"" then
                            datosNenvio=enc.EncodeForHtmlAttribute(rst("nenvio"))
                        else
                            datosNenvio=rst("nenvio")
                        end if
                        'EligeCelda "input", mode,"CELDA","","",0,LitNumEnvio,"nenvio",27,iif(tmp_nenvio>"",tmp_nenvio,iif(isnull(rst("nenvio")),"",rst("nenvio")))
                        EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), enc.EncodeForHtmlAttribute(claseNumEnvio), "", "nenvio", iif(tmp_nenvio>"",tmp_nenvio,iif(isnull(datosNenvio),"",datosNenvio)), ""
                    end if
                    ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                    ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                    CloseDiv



			            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitTarifa+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitTarifa
			            campo="codigo"
			            campo2="descripcion"
			            ''ricardo 25/4/2003 para que salga el mensaje de que se ha cambiado la tarifa
			            if mode="edit" then%>
				            <input type="hidden" name="h_tarifa" value="<%=rst("tarifa")%>"/>
			            <%end if
			            ''''''''''
			            if mode<>"browse" then
				            rstAux.open "select codigo,descripcion from tarifas with(nolock) where codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                            if mode="edit" then
                                defectoTarifa=iif(tmp_tarifa>"",tmp_tarifa,rst("tarifa"))
                            else
                                defectoTarifa=iif(tmp_tarifa>"",tmp_tarifa,"")
                            end if
				            'DrawSelectCeldaResponsive1 "width:200px","200","",0,LitTarifa,"tarifa",rstAux,defectoTarifa,"codigo","descripcion","",""
				            DrawSelect1 "selectmenu width50", "", "tarifa", "tarifa", rstAux, enc.EncodeForHtmlAttribute(defectoTarifa), "codigo","descripcion","","","",""
                            rstAux.close
			            else
				            dato_celda=Desplegable(mode,campo,campo2,"tarifas",iif(tmp_tarifa>"",tmp_tarifa,rst("tarifa")),"codigo like '" & session("ncliente") & "%'")
				            'EligeCelda "select", mode,"CELDA style='width:200px'","","",0,LitTarifa,"tarifa",20,dato_celda
				            EligeCeldaResponsive1 "select", enc.EncodeForHtmlAttribute(null_s(mode)), "CELDA", "width50", "tarifa", enc.EncodeForHtmlAttribute(null_s(dato_celda)), ""
                            if mode="add" or mode="edit" then RstAux.close
			            end if
                        CloseDiv

			            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
			            'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitFechaPedido+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitFechaPedido
                        claseFechaPedido = "'CELDA  width150px'"
                        if mode="add" or mode="edit" then
                            claseFechaPedido = "'datepicker special-input width150px'"
                        end if
                        if mode="add" then
			                'EligeCelda "input", mode,"CELDA","","",0,LitFechaPedido,"fechapedido",27,iif(tmp_fechapedido>"",tmp_fechapedido,"")
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), claseFechaPedido, "", "fechapedido", iif(tmp_fechapedido>"",tmp_fechapedido,""), ""
                        else
                            'EligeCelda "input", mode,"CELDA","","",0,LitFechaPedido,"fechapedido",27,iif(tmp_fechapedido>"",tmp_fechapedido,iif(isnull(rst("fecha_pedido")),"",rst("fecha_pedido")))
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), claseFechaPedido, "", "fechapedido", iif(tmp_fechapedido>"",tmp_fechapedido,iif(isnull(rst("fecha_pedido")),"",rst("fecha_pedido"))), ""
                        end if
                        if mode="add" or mode="edit" then DrawCalendar "fechapedido" end if  
                        CloseDiv
                      
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
		            'HYPERVINCULO
			            'DrawCelda "ENCABEZADOL style='width:140px'","","",0,LitHipervinculo+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitHipervinculo
			            if mode="browse" then
				            doc_aux=iif(tmp_documento>"",tmp_documento,iif(isnull(rst("documento")),"",rst("documento")))
				            if doc_aux="" then
					            'DrawCelda "CELDA7 valign='top'","","",0,""
                                DrawSpan "CELDA", "", "", ""
				            else
					            url="javascript:opendoc('" & reemplazar(reemplazar(doc_aux," ","%20"),"\","\\") & "');"
					            'DrawCeldahref "CELDAREF7 valign='top'","left","false",doc_aux,url
                                DrawHref "CELDAREF7", "", enc.EncodeForHtmlAttribute(null_s(doc_aux)), enc.EncodeForHtmlAttribute(null_s(url))
				            end if
			            else
                            if mode<>"add" then
				                if (tmp_documento="") and (not isnull(rst("documento"))) then
    					            tmp_documento=rst("documento")
	    			            end if
                            end if
                            %><input id="input_doc" readonly="readonly" class="" type="text" name="documento" maxlength="255" size="25" value="<%=enc.EncodeForHtmlAttribute(tmp_documento)%>"/>
							<input id="input_file" class="" type="file" name="h_file" maxlength="255" size="25" style='width:0px;display:none;'
					            onclick="javascript: document.facturas_cli.documento.value = document.facturas_cli.h_file.value;"
					            onchange="javascript:document.facturas_cli.documento.value=document.facturas_cli.h_file.value;" value=""/>
							<input class="" type="button" name="vacprecli" value="<%=LitPreCliVaciar%>" onclick="javascript: document.facturas_cli.documento.value = '';"/>
                            <input class="" type="button" name="examinar" value="Examinar" onclick="javascript: document.getElementById('input_file').click();"/><%
                        end if
                        CloseDiv
			            if mode="browse" then
                            DrawDiv "1", "", ""
                            DrawLabel "", "", LitContacto
                            if rst("contacto") & "" <>"" then
                               datosContacto=enc.EncodeForHtmlAttribute(rst("contacto"))
                            else
                               datosContacto=rst("contacto")
                            end if
                            DrawSpan "CELDA", "", datosContacto&"", ""
                            CloseDiv
			            else
                            DrawDiv "1", "", ""
                            DrawLabel "", "", LitContacto
                            if mode="add" then
                                ValorContacto=iif(tmp_contacto>"",tmp_contacto,"")
                                ValorBusqCont=iif(tmp_ncliente>"",tmp_ncliente,"")
                            else
                                if rst("contacto") & "" <>"" then
                                   datosContacto=enc.EncodeForHtmlAttribute(rst("contacto"))
                                else
                                   datosContacto=rst("contacto")
                                end if
                                ValorContacto=iif(tmp_contacto>"",tmp_contacto,iif(isnull(datosContacto),"",datosContacto))
                                ValorBusqCont=iif(tmp_ncliente>"",tmp_ncliente,iif(isnull(rst("ncliente")),"",rst("ncliente")))
                            end if
                            %><div class="icon-input DisInTab">
                                    <input type="text" placeholder="" class="width50" name="contacto" value="<%=enc.EncodeForHtmlAttribute(ValorContacto)%>"/>
                                    <a class="CELDAREFB" href="javascript:AbrirVentana('busqueda_contactos.asp?viene=facturas_cli&ncliente=<%=enc.EncodeForJavascript(ValorBusqCont)%>&titulo=<%=LitSelContCliente%> <%=enc.EncodeForJavascript(trimCodEmpresa(ValorBusqCont))%>&mode=search','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitVerContacto%>'; return true;" OnMouseOut="self.status=''; return true;">
                                        <img src="<%=enc.EncodeForHtmlAttribute(ImgBuscarDinamic)%>" <%=ParamImgBuscarDinamic%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>">
                                    </a>
                                </div><%
                            CloseDiv
                        end if
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
		                'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitBanco+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitBanco
                        claseBanco = "CELDA"
                        if mode="add" or mode="edit" then
                            claseBanco = ""
                        end if
                        if mode="add" then
			                'EligeCelda "input", mode,"CELDA","","",0,LitBanco,"banco",27,iif(tmp_banco>"",tmp_banco,"") 
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), claseBanco, "", "banco", iif(tmp_banco>"",enc.EncodeForHtmlAttribute(null_s(tmp_banco)),""), ""
                        else
                            'EligeCelda "input", mode,"CELDA","","",0,LitBanco,"banco",27,iif(tmp_banco>"",tmp_banco,trim(iif(isnull(rst("banco")),"",rst("banco"))))
                            EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), claseBanco, "", "banco", enc.EncodeForHtmlAttribute(iif(isnull(rst("ncuenta")) or TraerCliente> "",iif(tmp_banco>"", tmp_banco&"", ""),rst("banco")&"")), ""
                        end if
                        CloseDiv
			            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
			            'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitNCuenta+":"

                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitNCuenta

			            if mode="browse" then

                            if not isnull(rst("ncuenta")) and len(rst("ncuenta"))>14 then
                                    restoCuenta = Len(rst("ncuenta"))-14
                            end if

                            tmp_ncuenta=iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),1,2))
                                tmp_ncuenta=tmp_ncuenta & " " & iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),3,2))
				                tmp_ncuenta=tmp_ncuenta & " " & iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),5,4))
				                tmp_ncuenta=tmp_ncuenta & "-" & iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),9,4))
				                tmp_ncuenta=tmp_ncuenta & "-" & iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),13,2))
                                tmp_ncuenta=tmp_ncuenta & "-" & iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),15,restoCuenta))
                            'if len(rst("ncuenta")) >0 then
				                'tmp_ncuenta=iif(tmp_ncuenta1>"",tmp_ncuenta1,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),1,2)))
                                'tmp_ncuenta=tmp_ncuenta & " " & iif(tmp_ncuenta2>"",tmp_ncuenta2,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),3,2)))
				                'tmp_ncuenta=tmp_ncuenta & " " & iif(tmp_ncuenta3>"",tmp_ncuenta3,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),5,4)))
				                'tmp_ncuenta=tmp_ncuenta & "-" & iif(tmp_ncuenta4>"",tmp_ncuenta4,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),9,4)))
				                'tmp_ncuenta=tmp_ncuenta & "-" & iif(tmp_ncuenta5>"",tmp_ncuenta5,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),13,2)))
                                'tmp_ncuenta=tmp_ncuenta & "-" & iif(tmp_ncuenta6>"",tmp_ncuenta6,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),15,restoCuenta)))
                            'else
                                'tmp_ncuenta=iif(tmp_ncuenta1>"",tmp_ncuenta1,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),1,4)))
				                'tmp_ncuenta=tmp_ncuenta & "-" & iif(tmp_ncuenta2>"",tmp_ncuenta2,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),5,4)))
				                'tmp_ncuenta=tmp_ncuenta & "-" & iif(tmp_ncuenta3>"",tmp_ncuenta3,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),9,2)))
				                'tmp_ncuenta=tmp_ncuenta & "-" & iif(tmp_ncuenta4>"",tmp_ncuenta4,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),11,Len(rst("ncuenta"))-10)))
                            'end if
				            tmp_ncuenta=trim(tmp_ncuenta)
				            if tmp_ncuenta="---" then tmp_ncuenta=""
				            'EligeCelda "input", mode,"CELDA","","",0,LitNCuenta,"ncuenta",4,tmp_ncuenta
                            EligeCeldaResponsive1 "input", mode, "CELDA", "", "ncuenta", tmp_ncuenta, ""
			            else
                            if mode="add" then
                                ncuenta1=iif(tmp_ncuenta1>"",tmp_ncuenta1,"")
                                ncuenta2=iif(tmp_ncuenta2>"",tmp_ncuenta2,"")
                                ncuenta3=iif(tmp_ncuenta3>"",tmp_ncuenta3,"")
                                ncuenta4=iif(tmp_ncuenta4>"",tmp_ncuenta4,"")
                                ncuenta5=iif(tmp_ncuenta5>"",tmp_ncuenta5,"")
                                ncuenta6=iif(tmp_ncuenta6>"",tmp_ncuenta6,"")
                                %><div class="inlineTable width100"><%
                                %><div class="width10 tableCell"><input type="text" data-enhanced="true" class="input-bank-1" placeholder="00"              name="ncuenta1" maxlength="2" value="<%=enc.EncodeForHtmlAttribute(ncuenta1)%>" onkeyup="if (this.value.length==2) document.facturas_cli.ncuenta2.focus()" onchange="document.facturas_cli.ncuenta1.value=document.facturas_cli.ncuenta1.value.toUpperCase();"/></div><%
				 	                %><div class="width10 tableCell"><input type="text" data-enhanced="true" class="input-bank-2" placeholder="00"          name="ncuenta2" maxlength="2" value="<%=enc.EncodeForHtmlAttribute(ncuenta2)%>" onkeyup="if (this.value.length==2) document.facturas_cli.ncuenta3.focus()"/></div><%
				 	                %><div class="width20 tableCell"><input type="text" data-enhanced="true" class="input-bank-3" placeholder="0000"        name="ncuenta3" maxlength="4" value="<%=enc.EncodeForHtmlAttribute(ncuenta3)%>" onkeyup="if (this.value.length==4) document.facturas_cli.ncuenta4.focus()"/></div><%
                                    %><div class="width20 tableCell"><input type="text" data-enhanced="true" class="input-bank-4" placeholder="0000"        name="ncuenta4" maxlength="4" value="<%=enc.EncodeForHtmlAttribute(ncuenta4)%>" onkeyup="if (this.value.length==4) document.facturas_cli.ncuenta5.focus()"/></div><%
                                    %><div class="width10 tableCell"><input type="text" data-enhanced="true" class="input-bank-5" placeholder="00"          name="ncuenta5" maxlength="2" value="<%=enc.EncodeForHtmlAttribute(ncuenta5)%>" onkeyup="if (this.value.length==2) document.facturas_cli.ncuenta6.focus()"/></div><%
			 		                %><div class="width40 tableCell"><input type="text" data-enhanced="true" class="input-bank-6" placeholder="0000000000"  name="ncuenta6" value="<%=enc.EncodeForHtmlAttribute(ncuenta6)%>"/></div><%
				                %></div><%
                            else
                                'ncuenta1=iif(tmp_ncuenta1>"",tmp_ncuenta1,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),1,2)))
                                'ncuenta2=iif(tmp_ncuenta2>"",tmp_ncuenta2,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),3,2)))
                                'ncuenta3=iif(tmp_ncuenta3>"",tmp_ncuenta3,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),5,4)))
                                'ncuenta4=iif(tmp_ncuenta4>"",tmp_ncuenta4,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),9,4)))
                                'ncuenta5=iif(tmp_ncuenta5>"",tmp_ncuenta5,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),13,2)))
                                'ncuenta6=iif(tmp_ncuenta6>"",tmp_ncuenta6,iif(isnull(rst("ncuenta")),"",mid(rst("ncuenta"),15,Len(rst("ncuenta"))-14)))
                                if IsNull(rst("ncuenta")) or rst("ncuenta")&"" = "" or TraerCliente> "" then
                                    ncuenta1=iif(tmp_ncuenta1>"",tmp_ncuenta1,"")
                                    ncuenta2=iif(tmp_ncuenta2>"",tmp_ncuenta2,"")
                                    ncuenta3=iif(tmp_ncuenta3>"",tmp_ncuenta3,"")
                                    ncuenta4=iif(tmp_ncuenta4>"",tmp_ncuenta4,"")
                                    ncuenta5=iif(tmp_ncuenta5>"",tmp_ncuenta5,"")
                                    ncuenta6=iif(tmp_ncuenta6>"",tmp_ncuenta6,"")
                                else
                                    ncuenta1=mid(rst("ncuenta"),1,2)
                                    ncuenta2=mid(rst("ncuenta"),3,2)
                                    ncuenta3=mid(rst("ncuenta"),5,4)
                                    ncuenta4=mid(rst("ncuenta"),9,4)
                                    ncuenta5=mid(rst("ncuenta"),13,2)
                                    ncuenta6=mid(rst("ncuenta"),15,Len(rst("ncuenta"))-14)
                                end if
				                    %><div class="inlineTable width100"><%
				 	                %><div class="width10 tableCell"><input type="text" data-enhanced="true" class="input-bank-1" placeholder="00"          name="ncuenta1" maxlength="2" value="<%=enc.EncodeForHtmlAttribute(ncuenta1)%>" onkeyup="if (this.value.length==2) document.facturas_cli.ncuenta2.focus()" onchange="document.facturas_cli.ncuenta1.value=document.facturas_cli.ncuenta1.value.toUpperCase();"/></div><%
				 	                %><div class="width10 tableCell"><input type="text" data-enhanced="true" class="input-bank-2" placeholder="00"          name="ncuenta2" maxlength="2" value="<%=enc.EncodeForHtmlAttribute(ncuenta2)%>" onkeyup="if (this.value.length==2) document.facturas_cli.ncuenta3.focus()"/></div><%
				 	                %><div class="width20 tableCell"><input type="text" data-enhanced="true" class="input-bank-3" placeholder="0000"        name="ncuenta3" maxlength="4" value="<%=enc.EncodeForHtmlAttribute(ncuenta3)%>" onkeyup="if (this.value.length==4) document.facturas_cli.ncuenta4.focus()"/></div><%
                                    %><div class="width20 tableCell"><input type="text" data-enhanced="true" class="input-bank-4" placeholder="0000"        name="ncuenta4" maxlength="4" value="<%=enc.EncodeForHtmlAttribute(ncuenta4)%>" onkeyup="if (this.value.length==4) document.facturas_cli.ncuenta5.focus()"/></div><%
                                    %><div class="width10 tableCell"><input type="text" data-enhanced="true" class="input-bank-5" placeholder="00"          name="ncuenta5" maxlength="2" value="<%=enc.EncodeForHtmlAttribute(ncuenta5)%>" onkeyup="if (this.value.length==2) document.facturas_cli.ncuenta6.focus()"/></div><%
			 		                %><div class="width40 tableCell"><input type="text" data-enhanced="true" class="input-bank-6" placeholder="0000000000"  name="ncuenta6" value="<%=enc.EncodeForHtmlAttribute(ncuenta6)%>"/></div><%
				                %></div><% 
                            end if
                        end if
                        CloseDiv

		            
		            if mode="browse" then
                        if not isnull(rst("fecharemesa")) then
			                    'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitFechaRemesa
				                'DrawCelda "CELDA","","",0,rst("fecharemesa")
                                DrawDiv "1", "", ""
                                DrawLabel "", "", LitFechaRemesa
                                DrawSpan "CELDA", "", rst("fecharemesa"), ""
                                CloseDiv
			                    'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitDescRemesa
				                'DrawCeldaSpan "CELDA","","",0,d_lookup("descripcion","remesas","nremesa='" & rst("nremesa")&"'",session("dsn_cliente")),4

                                'd_lookup("descripcion","remesas","nremesa='" & rst("nremesa")&"'",session("dsn_cliente"))

                                DivRemesaSelect="select descripcion from remesas with(nolock) where nremesa=?"
			                    DrawDiv "1", "", ""
                                DrawLabel "", "", LitDescRemesa
                                DrawSpan "CELDA", "",enc.EncodeForHtmlAttribute(null_s(DlookupP1(DivRemesaSelect, rst("nremesa"), adInteger, 4, session("dsn_cliente"))))&"", ""
                                CloseDiv
                            
                        end if
		            end if

                    if session("ncliente") <> ncompany_GALP then
                        dim LitComercialAux
			            if si_tiene_modulo_comercial<>0 then
				            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitComercialModCom+":"
                            LitComercialAux = LitComercialModCom
			            else
				            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,Litcomercial+":"
                            LitComercialAux = Litcomercial
			            end if

			            defecto=""
                        if mode="addNNN" then
			                if isnull(rst("comercial")) and tmp_comercial="" then
				                rstAux.open "select comercial from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				                if not rstAux.eof then
					                defecto=rstAux("comercial")
				                end if
				                rstAux.close
			                end if
                        end if

                        if mode<>"add" then
			                if rst("comercial")>"" and tmp_comercial="" then
				                defecto=rst("comercial")
			                end if
                        end if
			            if tmp_comercial>"" then
				            defecto=tmp_comercial
			            end if

			            if mode <> "browse" then
                            '***RGU 25/4/2006
                            condicion_comercial=0
                            if mode<>"add" then
                                if isnull(rst("comercial")) then
                                    condicion_comercial=1
                                end if
                            end if
                            if (mode="edit" and nmc<>"1" ) or mode="add" or condicion_comercial=1 then
                                rstAux.cursorlocation=3
                                rstAux.open "select dni, nombre from personal with(nolock) inner join comerciales with(nolock) on dni=comercial where comerciales.fbaja is null and dni like '" & session("ncliente") & "%' and comercial like '" & session("ncliente") & "%' order by nombre",session("dsn_cliente")
                                'DrawSelectCelda "CELDA","200","",0,LitComercialAux,"comercial",rstAux,defecto,"dni","nombre","",""
                                DrawDiv "1", "", ""
                                DrawLabel "", "", LitComercialAux
                                DrawSelect1 "combobox", "", "comercial", "comercial", rstAux, enc.EncodeForHtmlAttribute(defecto), "dni", "nombre", "", "", "", ""
                                CloseDiv
                                    
                                rstAux.close
                            else
                                'DrawCelda "CELDA","","",0,d_lookup("nombre","personal","dni='" & defecto & "'",session("dsn_cliente"))
                                'd_lookup("nombre","personal","dni='" & defecto & "'",session("dsn_cliente"))
                                ComAuxSelect="select nombre from personal with(nolock) where dni=?"
                                DrawDiv "1", "", ""
                                DrawLabel "", "", LitComercialAux
                                DrawSpan "CELDA", "", DlookupP1(ComAuxSelect, defecto&"", adVarchar, 20, session("dsn_cliente")), ""
                                CloseDiv
                                
                                %>
                                <input type="hidden" name="comercial" value="<%=enc.EncodeForHtmlAttribute(rst("comercial"))%>"/>
                                <%
                            end if
                            '***RGU
		 	            else
		 		            'DrawCelda "CELDA","","",0,d_lookup("nombre","personal","dni like '" & session("ncliente") & "%' and dni='" & defecto & "'",session("dsn_cliente"))

                            'd_lookup("nombre","personal","dni like '" & session("ncliente") & "%' and dni='" & defecto & "'",session("dsn_cliente"))
                            ComAuxSelect="select nombre from personal with(nolock) where dni like ?+'%' and dni=?"
                            
			                DrawDiv "1", "", ""
                            DrawLabel "", "", LitComercialAux
                            DrawSpan "CELDA", "", enc.EncodeForHtmlAttribute(DlookupP2(ComAuxSelect, session("ncliente")&"", adVarchar, 20, defecto&"", adVarchar, 20, session("dsn_cliente"))), ""
                            CloseDiv
                        end if

                        if mode="add" then
                            %>
			                <input type="hidden" name="h_comercial" value=""/>
			                <%
                        else
                            %>
			                <input type="hidden" name="h_comercial" value="<%=enc.EncodeForHtmlAttribute(Nulear(rst("comercial")))%>"/>
			                <%
                        end if

                        'Agente
                        'if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
			            if si_tiene_modulo_comercial<>0 then
				            'DrawCelda "ENCABEZADOL style='width:180px'","","",0,Litagenteasignado+":"
				            defecto=""
                            if mode="addNNN" then
				                if isnull(rst("agente")) and tmp_agenteasignado="" then
					                rstAux.open "select agente from clientes with(nolock) where ncliente='" & rst("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					                if not rstAux.eof then
						                defecto=rstAux("agente")
					                end if
					                rstAux.close
				                end if
                            end if
                            if mode<>"add" then
				                if rst("agente")>"" and tmp_agenteasignado="" then
					                defecto=rst("agente")
				                end if
                            end if
				            if tmp_agenteasignado>"" then
					            defecto=tmp_agenteasignado
				            end if

                            DrawDiv "1", "", ""
                            DrawLabel "", "", Litagenteasignado
				            if mode <> "browse" then
                                rstAux.cursorlocation=3
			 		            rstAux.open "select codigo, nombre from agentes with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre",session("dsn_cliente")
					            'DrawSelectCelda "CELDA","200","",0,Litagenteasignado,"agenteasignado",rstAux,defecto,"codigo","nombre","",""
				 	            DrawSelect1 "combobox", "", "agenteasignado", "agenteasignado", rstAux, enc.EncodeForHtmlAttribute(null_s(defecto)), "codigo", "nombre", "", "", "", ""
                                
                                rstAux.close
				            else
			 		            'DrawCelda "CELDA style='width:200px'","","",0,d_lookup("nombre","agentes","codigo like '" & session("ncliente") & "%' and codigo='" & defecto & "'",session("dsn_cliente"))

                                spaSelect="select nombre from agentes with(nolock) where codigo like ?+'%' and codigo = ?"
                                'd_lookup("nombre","agentes","codigo like '" & session("ncliente") & "%' and codigo='" & defecto & "'",session("dsn_cliente"))
				                DrawSpan "CELDA", "", enc.EncodeForHtmlAttribute(DlookupP2(spaSelect, session("ncliente")&"", adChar, 10, defecto&"", adChar, 10, session("dsn_cliente"))), ""
                            end if
                            CloseDiv

			            end if

                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"


			            if si_tiene_modulo_proyectos<>0 then
				            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitProyecto+":"
                            DrawDiv "1", "", ""
                            DrawLabel "", "", LitProyecto
				            if mode <> "browse" then
                                if mode="add" then
                                    Valor_Proyecto=iif(tmp_cod_proyecto>"",tmp_cod_proyecto,"")
                                    Valor_Proyecto_Cli=iif(tmp_ncliente>"",trimCodEmpresa(tmp_ncliente),"")
                                else
                                    Valor_Proyecto=iif(tmp_cod_proyecto>"",tmp_cod_proyecto,iif(isnull(rst("cod_proyecto")),"",rst("cod_proyecto")))
                                    Valor_Proyecto_Cli=iif(tmp_ncliente>"",trimCodEmpresa(tmp_ncliente),trimCodEmpresa(rst("ncliente")))
                                end if

                                'TODO Revisar responsive
                                %>
						            <input class="CELDA" type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(Valor_Proyecto)%>"/>
						            <iframe id="frProyecto" src='../mantenimiento/docproyectos.asp?viene=facturas_cli&mode=<%=enc.EncodeForHtmlAttribute(mode)%>&cod_proyecto=<%=enc.EncodeForHtmlAttribute(Valor_Proyecto)%>&ncliente=<%=enc.EncodeForHtmlAttribute(Valor_Proyecto_Cli)%>' width='250' height='30' frameborder="no" scrolling="no" noresize="noresize"></iframe>
				                <%
                            else
				                'DrawCelda "CELDA","","",0,d_lookup("nombre","proyectos","codigo like '" & session("ncliente") & "%' and codigo='" & rst("cod_proyecto") & "'",session("dsn_cliente"))
                                'd_lookup("nombre","proyectos","codigo like '" & session("ncliente") & "%' and codigo='" & rst("cod_proyecto") & "'",session("dsn_cliente"))

                                span2Select="select nombre from proyectos with(nolock) where codigo like ?+'%' and codigo= ?"

				                DrawSpan "CELDA", "", DlookupP2(span2Select , session("ncliente")&"", adVarchar, 15, rst("cod_proyecto")&"", adVarchar, 15, session("dsn_cliente")), ""
                            end if
                            CloseDiv

				            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
				            if mode="first_save" or mode="browse" then
					            'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitEDI
                                DrawDiv "1", "", ""
                                DrawLabel "", "", LitEDI
                                if mode="add" then
                                    'DrawCelda "CELDA","","",0,""
                                    DrawSpan "", "", "", ""
                                else
					                'DrawCelda "CELDA","","",0,rst("edi")&""
                                    DrawSpan "CELDA", "", enc.EncodeForHtmlAttribute(rst("edi")&""), ""
                                end if
                                CloseDiv
				            end if
			            else
				            if mode="first_save" or mode="browse" then
					            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LitEDI
                                DrawDiv "1", "", ""
                                DrawLabel "", "", LitEDI
                                if mode="add" then
                                    'DrawCelda "CELDA","","",0,""
                                    DrawSpan "", "", "", ""
                                else
					                'DrawCelda "CELDA","","",0,rst("edi")&""
                                    DrawSpan "CELDA", "", enc.EncodeForHtmlAttribute(rst("edi")&""), ""
                                end if
                                CloseDiv
				            end if
			            end if
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
		            

			            'DrawCelda "ENCABEZADOL style='width:120px'","","",0,LITIINCOTFACCli+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LITIINCOTFACCli
			            if mode="browse" then
				            'DrawCelda "CELDA","",LITIINCOTFACCli,0,rst("incoterms")&""
                            DrawSpan "CELDA", "", enc.EncodeForHtmlAttribute(rst("incoterms")&""), ""
			            else
                            if mode="add" then
				                defecto=iif(tmp_incoterms>"",tmp_incoterms,"")
                            else
                                defecto=iif(tmp_incoterms>"",tmp_incoterms,iif(rst("incoterms")>"",rst("incoterms"),""))
                            end if
				            rstAux.cursorlocation=3
				            rstAux.open "select codigo,codigo as descripcion from incoterms with(nolock) order by descripcion",session("dsn_cliente")
				            'DrawSelectCelda "CELDALEFT","60","",0,LITIINCOTFACCli,"incoterms",rstAux,defecto,"codigo","descripcion","",""
				            DrawSelect1 "combobox", "", "incoterms", "incoterms", rstAux, enc.EncodeForHtmlAttribute(null_s(defecto)), "codigo", "descripcion", "", "", "", ""
                                
                            rstAux.close
			            end if
                        CloseDiv

			            'if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
			            'DrawCelda "ENCABEZADOL style='width:180px'","","",0,LitIncoPuntEntrFacCli+":"
                        DrawDiv "1", "", ""
                        DrawLabel "", "", LitIncoPuntEntrFacCli
			            if mode="browse" then
				            'DrawCelda "'CELDALEFT' align='left' ","","",0,rst("fob")&""
                            DrawSpan "CELDA", "", enc.EncodeForHtmlAttribute(rst("fob")&""), ""
			            else
                            if mode="add" then
                                defecto=iif(tmp_fob>"",tmp_fob,"")
                            else
                                if rst("fob") & "" <>"" then
                                   datosfob=enc.EncodeForHtmlAttribute(rst("fob"))
                                else
                                   datosfob=rst("fob")
                                end if
                                defecto=iif(tmp_fob>"",tmp_fob,iif(datosfob>"",datosfob,""))
                            end if
					        '<input type="text" maxlength="50" size="25" name="fob" value="<=defecto>"/>
                            DrawInput "", "", "fob", enc.EncodeForHtmlAttribute(null_s(defecto)), "maxlength='50'"
                        end if
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        ''if mode<>"browse" then DrawCelda "CELDA","10","",0,"&nbsp;&nbsp;"
                        CloseDiv

		            
                    else%>
                        <input type="hidden" name="comercial" value=""/>
			            <input type="hidden" name="h_comercial" value=""/>
                        <input type="hidden" name="agenteasignado" value=""/>
                        <input type="hidden" name="cod_proyecto" value=""/>
                        <input type="hidden" name="incoterms" value=""/>
                        <input type="hidden" name="fob" value=""/>
                    <%end if

		            if mode="browse" then

                            if rst("observaciones") & "" <>"" then
                               datosObservaciones=enc.EncodeForHtmlAttribute(rst("observaciones"))
                            else
                               datosObservaciones=rst("observaciones")
                            end if

                            'DrawCelda "ENCABEZADOL style='width:140px' valign=top","","",0,LitObservaciones
			                'DrawCelda "CELDA id='observaciones'","","",0,pintar_saltos_espacios(iif(tmp_observaciones>"",tmp_observaciones,rst("observaciones")&""))
			                DrawDiv "1", "", ""
                            DrawLabel "", "", LitObservaciones
                            DrawSpan "CELDA", "", pintar_saltos_espacios(iif(tmp_observaciones>"",enc.EncodeForHtmlAttribute(null_s(tmp_observaciones)),datosObservaciones&"")), "id='observaciones'"
                            CloseDiv
                        
                            if rst("notas") & "" <>"" then
                               datosNotas=enc.EncodeForHtmlAttribute(rst("notas"))
                            else
                               datosNotas=rst("notas")
                            end if

				            'DrawCelda "ENCABEZADOL style='width:140px' valign=top","","",0,LitNotas
                            'DrawCelda "CELDA","","",0,pintar_saltos_espacios(iif(tmp_notas>"",tmp_notas,rst("notas")&""))
                        	DrawDiv "1", "", ""
                            DrawLabel "", "", LitNotas
                            DrawSpan "CELDA", "", pintar_saltos_espacios(iif(tmp_notas>"",enc.EncodeForHtmlAttribute(null_s(tmp_notas)),datosNotas&"")), ""
                            CloseDiv
			            
		            else		   

				            'DrawCelda "ENCABEZADOL style='width:140px'","","",0,LitObservaciones+":"
                            DrawDiv "1", "", ""
                            DrawLabel "", "", LitObservaciones
                            if mode="add" then
                                Valor_Observaciones=iif(tmp_observaciones>"",tmp_observaciones,"")
                            else
                                if rst("observaciones") & "" <>"" then
                                   datosObservaciones=enc.EncodeForHtmlAttribute(rst("observaciones"))
                                else
                                   datosObservaciones=rst("observaciones")
                                end if
                                Valor_Observaciones=iif(tmp_observaciones>"",tmp_observaciones,datosObservaciones&"")
                            end if
				            '<textarea class='CELDA' id='observaciones' name="observaciones" cols="30" rows="2"><=Valor_Observaciones></textarea>    
                            DrawTextarea "width100", "", "observaciones", enc.EncodeForHtmlAttribute(null_s(Valor_Observaciones)), "id='observaciones'"
                            CloseDiv

                            'DrawCelda "ENCABEZADOL style='width:140px'","","",0,LitNotas+":"
                            DrawDiv "1", "", ""
                            DrawLabel "", "", LitNotas
                            if mode="add" then
                                Valor_Notas=iif(tmp_notas>"",tmp_notas,"")
                            else
                                if rst("notas") & "" <>"" then
                                   datosNotas=enc.EncodeForHtmlAttribute(rst("notas"))
                                else
                                   datosNotas=rst("notas")
                                end if
                                Valor_Notas=iif(tmp_notas>"",tmp_notas,datosNotas&"")
                            end if
				            '<textarea class='CELDA' name="notas" cols="30" rows="2"><=Valor_Notas></textarea>
                            DrawTextarea "width100", "", "notas", enc.EncodeForHtmlAttribute(null_s(Valor_Notas)), ""
                            CloseDiv

		            end if

	                '************************'
	                'JMA 28/10/04 ***********'
	                '************************'
	                ''ricardo 3-4-2009 paso el pintado de camposperso en una funcion, dentro del fichero camposperso.inc
	                dim num_campos_existenCP,max_num_camposCP,lista_valoresCP,tmp_lista_valoresCP
	                num_campos_existenCP=0
	                max_num_camposCP=0
	                if si_campo_personalizables=1 then
                        DrawDiv "3-sub", "background-color: #eae7e3", ""
                        %><label class="ENCABEZADOC" style="text-align:left"><%=LITCAMPPERSOLISTFACT %></label><%
                        CloseDiv
                        
	                    campos_a_dimensionarVentaCP=num_campos_ventas
	                    campos_a_dimensionarCliCP=num_campos_clientes
		                if cstr(campos_a_dimensionarVentaCP & "")="" then campos_a_dimensionarVentaCP=10
		                if cstr(campos_a_dimensionarCliCP & "")="" then campos_a_dimensionarCliCP=10
                        redim lista_valoresCP(campos_a_dimensionarVentaCP+5)
                        redim tmp_lista_valoresCP(campos_a_dimensionarCliCP+5)
                        tamany_maximo_lista_valores=ubound(lista_valores)
                        tamany_maximo_tmp_lista_valores=ubound(tmp_lista_valores)
                        for golkio=1 to campos_a_dimensionarVentaCP
                            if golkio>tamany_maximo_lista_valores then
                                lista_valoresCP(golkio)=""
                            else
                                lista_valoresCP(golkio)=lista_valores(golkio)
                            end if
                        next
                        for golkio=1 to campos_a_dimensionarCliCP
                            if golkio>tamany_maximo_tmp_lista_valores then
                                tmp_lista_valoresCP(golkio)=""
                            else
                                tmp_lista_valoresCP(golkio)=tmp_lista_valores(golkio)
                            end if
                        next

                        PintarCamposPersoDoc "PEDIDOS_CLI",mode,lista_valoresCP,tmp_lista_valoresCP,campos_a_dimensionarCP,num_campos_existenCP,max_num_camposCP

                        max_num_campos=max_num_camposCP
                        num_campos_existen=num_campos_existenCP
                        if num_campos_existen<max_num_campos then
                            num_campos_existen=max_num_campos
                        end if
                        %><input type="hidden" name="num_campos" value="<%=enc.EncodeForHtmlAttribute(num_campos_existen)%>"/>
	                <%end if
	                '************************'
	                'FIN JMA 28/10/04 *******'
	                '************************'%>
		        </table>
            <!--</center>-->
            </div>
            </div>

		    <%if mode="browse" then%>
                <!--DATOS DEL DOCUMENTO-->
                <!--
                <div class="accordion clearfix" id="S_DATFINAN" >
                    <h5><%=LITDATDOC%></h5>
                    -->
                    <!--
                    <div class="Section" id="S_DATFINAN" style="float:left">
                        <a href="#" rel="toggle[DATFINAN]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                            <div class="SectionHeader">
                                <%=LITDATDOC%>
                                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />        
                            </div>
                        </a>
                    -->

            <div class="Section" id="S_DATFINAN" >
                <a href="#" rel="toggle[DATFINAN]" data-openimage="<%=enc.EncodeForHtmlAttribute(ImgNoCollapse) %>" data-closedimage="<%=enc.EncodeForHtmlAttribute(ImgCollapse) %>">
                    <div class="SectionHeader displayed">
                        <%=LITDATDOC%>
                        <img class="btn_folder" src="<%=enc.EncodeForHtmlAttribute(ImgNoCollapse) %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>

                <div class="SectionPanel" id="DATFINAN">
                <!--<center>-->
                    <div id="tabs" style="display:run-in" class="ui-tabs ui-widget ui-widget-content ui-corner-all">
                    <ul class="ui-tabs-nav ui-helper-reset ui-helper-clearfix ui-widget-header ui-corner-all">
                        <%if oculta=0 then %>
                        <li class="ui-state-default ui-corner-top"><a href="#tabs1"><%=LitDetalles%></a></li>
                        <%end if %>
                        <li class="ui-state-default ui-corner-top"><a href="#tabs2"><%=LitConceptos%></a></li>
                        <li class="ui-state-default ui-corner-top"><a href="#tabs3"><%=LitVencimientos%></a></li>
                        <li class="ui-state-default ui-corner-top"><a href="#tabs4"><%=LitPagosACuenta%></a></li>
                        <li class="ui-state-default ui-corner-top"><a href="#tabs5"><%=LitDatosEnvio%></a></li>
                        <%if SuplidosActivados=1 then%>
                        <li class="ui-state-default ui-corner-top"><a href="#tabs6"><%=LITFICSUPLIDOSFAC%></a></li>
                        <%end if%>
                    </ul>

			        <%
                    '***** Detalles de la factura'
                    'response.write("Detalles"&oculta)
                    'TAB 1
                    if oculta=0 then
                        'ega 19/06/2008 comento la consulta de rstDet porque no se utiliza en la pagina
                        'set rstDet = Server.CreateObject("ADODB.Recordset")
        		        'rstDet.open "select * from detalles_fac_cli where nfactura like '" & session("ncliente") & "%' and nfactura='" & rst("nfactura") & "' order by item", _
		    	        'session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                        'set rstDet =nothing
                        %>

                        <div id="tabs1" class="overflowXauto">
                        <%if pciva="SI" then%>
				            <table class="width90 md-table-responsive bCollapse" >
                        <%else%>
				            <table class="width90 md-table-responsive bCollapse" >
                        <%end if

				            'DrawFila color_terra%>
                            <tr>
					            <td class="ENCABEZADOL underOrange width5" ><%=LitItem%></td>
					            <td class="ENCABEZADOL underOrange width5" ><%=LitCantidad%></td>
					
                        <%''ricardo 13-10-2009 se quita el almacen para el modulo profesionales
                        if si_tiene_modulo_profesionales=0 then%>
                            <td class="ENCABEZADOL underOrange width10" ><%=LitReferencia%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitAlmacen%></td>
					        <td class="ENCABEZADOL underOrange width15" ><%=LitDescripcion%></td>
                        <%else %>
                            <td class="ENCABEZADOL underOrange width10" ><%=LitReferencia%></td>
                            <td class="ENCABEZADOL underOrange width15" ><%=LitDescripcion%></td>
                        <%end if %>
					
					        <%if pciva="SI" then%>
						        <td class="ENCABEZADOL underOrange width5" ><%=LitPVPIVA%></td>
					        <%end if%>
					        <td class="ENCABEZADOL underOrange width5" ><%=LitPVP%></td>

					        <td class='CELDAL7 underOrange width20' >
						        <div id="showpreus" style="display:'';visibility:visible">
							        <table class="width100 bCollapse">
								        <tr>
									        <td class="ENCABEZADOL width25" ><%=LitDto%></td>
									        <td class="ENCABEZADOL width25" ><%=LitDto2%></td>
									        <td class="ENCABEZADOL width25" ><%=LitDto3%></td>
									        <td class="ENCABEZADOL width25" ><%=LitIva%></td>
								        </tr>
							        </table>
						        </div>
						        <div id="showagr" style="display:none;background-color: <%=color_terra%>; layer-background-color: <%=color_terra%>;visibility:visible">
							        <table>
								        <tr>
									        <td width="140" class="ENCABEZADOL"><%=LitLote%></td>
								        </tr>
							        </table>
						        </div>
					        </td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitImporte%></td>
					        <td class='CELDAR7 underOrange width5' >&nbsp</td>
					        <%if si_tiene_modulo_produccion<>0 then%>
						        <td class="ENCABEZADOL underOrange width5" align="center" >
							        <div id="CapaVTD" class="CELDA7" style=" TEXT-ALIGN: center; width:8px;height:8px; z-index:2;cursor:pointer" onclick="mostrar('<%=enc.EncodeForJavascript(mode)%>')">+</div>&nbsp
						        </td>
					        <%else%>
						        <td class='CELDAR7 underOrange width5' width="13">&nbsp</td>
					        <%end if
				        'CloseFila%>
                        </tr>
				        </table>
				        <%

				        'cag
				        if rst("ahora")=0 and (rst("cobrada")=0 and pagado<>1 and nliquidacion &""=""  and nliquidacionAG &""="") and p_pagsl then
                            if pciva="SI" then
	                            tamano_ancho_iframe="833"
                            else
	                            tamano_ancho_iframe="800"
                            end if

                            %>
				            <iframe id='frDetallesIns' name="fr_DetallesIns" class='width90 iframe-input md-table-responsive' frameborder="0" scrolling="no" noresize="noresize" src='facturas_clidetins.asp?ndoc=<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>&ncliente=<%=enc.EncodeForHtmlAttribute(rst("ncliente"))%>&modp=<%=enc.EncodeForHtmlAttribute(modp)%>&modd=<%=enc.EncodeForHtmlAttribute(modd)%>&modi=<%=enc.EncodeForHtmlAttribute(modi)%>&dto1_cli=<%=enc.EncodeForHtmlAttribute(dto1_cli)%>&dto2_cli=<%=enc.EncodeForHtmlAttribute(dto2_cli)%>&dto3_cli=<%=enc.EncodeForHtmlAttribute(dto3_cli)%>&ganancia=<%=enc.EncodeForHtmlAttribute(ganancia)%>&pciva=<%=enc.EncodeForHtmlAttribute(pciva)%>&modn=<%=enc.EncodeForHtmlAttribute(modn)%>&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV)%>&wclc=<%=enc.EncodeForHtmlAttribute(wclc)%>' ></iframe><%
			            end if
                        if pciva="SI" then
	                        tamano_ancho_iframe2="863"
                        else
	                        tamano_ancho_iframe2="800"
                        end if

				        'i(EJM 04/09/2006) si la factura se acaba de crear, cambiar el modo de apertura de las lineas de detalle
				        if totalCrearLineasFacZona=1 then
                            num_factura=rst("nfactura")%>
				            <iframe id='frDetalles' name="fr_Detalles"  class="width90 iframe-data md-table-responsive" frameborder="yes"></iframe>
			            <%else%>
				            <iframe id='frDetalles' name="fr_Detalles" class="width90 iframe-data md-table-responsive" src='facturas_clidet.asp?ndoc=<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>&ncliente=<%=enc.EncodeForHtmlAttribute(rst("ncliente"))%>&modp=<%=enc.EncodeForHtmlAttribute(modp)%>&modd=<%=enc.EncodeForHtmlAttribute(modd)%>&modi=<%=enc.EncodeForHtmlAttribute(modi)%>&dto1_cli=<%=enc.EncodeForHtmlAttribute(dto1_cli)%>&dto2_cli=<%=enc.EncodeForHtmlAttribute(dto2_cli)%>&dto3_cli=<%=enc.EncodeForHtmlAttribute(dto3_cli)%>&ganancia=<%=enc.EncodeForHtmlAttribute(ganancia)%>&pciva=<%=enc.EncodeForHtmlAttribute(pciva)%>&modn=<%=enc.EncodeForHtmlAttribute(modn)%>&almacenSerie=<%=enc.EncodeForHtmlAttribute(almacenSerie) %>&almacenTPV=<%=enc.EncodeForHtmlAttribute(almacenTPV) %>&wclc=<%=enc.EncodeForHtmlAttribute(wclc)%>' width='100%' height='145' frameborder="yes"></iframe>			
				            <%
				        end if
				        'fin(EJM 04/09/2006)%>
                        <div id="paginacion">
		                </div>
                    </div>

		            <%
                    end if 'del oculta
                    'Fin TAB 1

			        ''ricardo 9/8/2004 se pondra el iva que tiene establecido el cliente
			        'TmpIvaCliente=d_lookup("iva","clientes","ncliente like '" & session("ncliente") & "%' and ncliente='" & rst("ncliente") & "'",session("dsn_cliente"))

                    TmpIvaClienteSelect="select iva from clientes with(nolock) where ncliente like ?+'%' and ncliente=?"
                    TmpIvaCliente=DlookupP2(TmpIvaClienteSelect, session("ncliente")&"", adChar, 10, rst("ncliente")&"", adChar, 10, session("dsn_cliente"))

			        'defaultIVA=d_lookup("iva","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))

                    defaultIVASelect="select iva from configuracion with(nolock) where nempresa= ?"
                    defaultIVA=DlookupP1(defaultIVASelect, session("ncliente")&"", adChar, 5, session("dsn_cliente"))

			        if TmpIvaCliente & "">"" then
				        TmpIva=TmpIvaCliente
			        else
				        TmpIva=defaultIVA
			        end if

			        '***** Conceptos de la factura'
			        %>
               
                    <div id="tabs2" class="overflowXauto" >

				    <input type="hidden" name="defaultIva" value="<%=enc.EncodeForHtmlAttribute(TmpIva)%>"/>
				    <table class="width90 md-table-responsive bCollapse" ><%
				        'Fila de encabezado'
				        'DrawFila color_terra
					        %>
                            <tr>
                            <td class="ENCABEZADOL underOrange width10" ><%=LitItem%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitCantidad%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitDescripcion%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitPVP%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitDto%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitDto2%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitDto3%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitIva%></td>
					        <td class="ENCABEZADOL underOrange width10" ><%=LitImporte%></td>
					        <td class="ENCABEZADOL underOrange width10" >&nbsp</td>
                            </tr>
                            <%
				        'CloseFila
				        'Linea de inserción de un concepto'
				        if rst("ahora")=0 and rst("cobrada")=0 and pagado<>1 and nliquidacion &""=""  and nliquidacionAG &""="" and p_pagsl then
					        'DrawFila color_blau
					        %>
                            <tr>
                                <td class='CELDAL7 underOrange width10' >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						        <td class='CELDAL7 underOrange width10' >
							        <input class="CELDAR7 width100"  type="text" name="cantidad" value="1" onchange="RoundNumValue(this ,<%=enc.EncodeForJavascript(DEC_CANT)%>);ImporteDetalle();"/>
						        </td>
						        <td class='CELDAL7 underOrange width10' >
							        <textarea class="CELDAL7 width100"  name="descripcion" onFocus="lenmensaje(this,0,255,'')" onKeydown="lenmensaje(this,0,255,'')" onKeyup="lenmensaje(this,0,255,'')" onBlur="lenmensaje(this,0,255,'')" rows="2"></textarea>
						        </td>
						        <td class='CELDAL7 underOrange width10' >
							        <input class="CELDAR7 width100"  type="text" name="pvp" value="0" onchange="RoundNumValue(this, <%=enc.EncodeForJavascript(DEC_PREC)%>);ImporteDetalle();"/>
						        </td>
						        <td class='CELDAL7 underOrange width10' >
							        <input class="CELDAR7 width100"  type="text" name="descuento" value="0" onchange="RoundNumValue(this, <%=enc.EncodeForJavascript(DECPOR)%>);ImporteDetalle();"/>
						        </td>
						        <td class='CELDAL7 underOrange width10' >
							        <input class="CELDAR7 width100"  type="text" name="descuento2" value="0" onchange="RoundNumValue(this, <%=enc.EncodeForJavascript(DECPOR)%>);ImporteDetalle();"/>
						        </td>
						        <td class='CELDAL7 underOrange width10'>
							        <input class="CELDAR7 width100"  type="text" name="descuento3" value="0" onchange="RoundNumValue(this, <%=enc.EncodeForJavascript(DECPOR)%>);ImporteDetalle();"/>
						        </td>
                                <td class='CELDAL7 underOrange width10' >
                                <%
                                rstSelect.cursorlocation=3
						        rstSelect.open "select tipo_iva, tipo_iva as descripcion from tipos_iva with(nolock) order by descripcion",session("dsn_cliente")
						        'DrawSelectCelda "CELDAL7 style='font-Family: Verdana;font-size: 7.0pt;text-align: right;color: #000000;width: 50px;'","","",0,"","iva",rstSelect,TmpIva,"tipo_iva","descripcion","",""
                                DrawSelect "'CELDAL7 width100'", "", "iva", rstSelect, enc.EncodeForHtmlAttribute(null_s(TmpIva)), "tipo_iva", "descripcion", "", ""
                                rstSelect.close
						        %>
                                </td>
                                <td class='CELDAL7 underOrange width10'>
							        <input class="CELDAR7 width100"  disabled type="text" name="importe" value="0"/>
							        <input type="hidden" name="importec_ant" value="0"/>
						        </td>
						        <td class="underOrange width10">
							        <a class="ic-accept" href="javascript:addConcepto('<%=enc.EncodeForJavascript(p_nfactura)%>');" onblur="javascript:document.facturas_cli.cantidad.focus();"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgNuevo)%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"></a>
						        </td><%
						        if oculta=1 then%>
						          <script language="javascript" type="text/javascript">
                                      document.facturas_cli.cantidad.focus();
                                      document.facturas_cli.cantidad.select();
						           </script><%
						        end if
                            %></tr><%
					        'CloseFila
				        end if
				    %>
                        </table>
				    <iframe id="frConceptos" class="width90 iframe-data md-table-responsive" name="fr_Conceptos" src='facturas_clicon.asp?mode=browse&ndoc=<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>'  height='80' frameborder="yes" noresize="noresize"></iframe>
			    </div>
                <%

			    '***** Vencimientos de la factura
			    %>
               
                <div id="tabs3" class="overflowXauto">
				    <table class="width90 md-table-responsive" ><%
				        'Fila de encabezado
				        'DrawFila color_terra
					    %>
                        <tr>
                            <td class="ENCABEZADOL width10 underOrange" ><%=LitNVencimiento%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitFecha%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitImporte%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitImpCobrado%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitCobTodo%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitObservaciones%></td>
					        <td class="ENCABEZADOL width10 underOrange" >&nbsp</td>
                        </tr>
                        <%
				        'CloseFila
				        'Linea de inserción de los vencimientos
				        if rst("ahora")=0 and rst("cobrada")=0 and nliquidacion &""=""  and nliquidacionAG &""="" and p_pagsl then
					        'DrawFila color_blau
						        %>
                                <tr>
                                <td class='CELDAL7 width10 underOrange' >
							        <a href="javascript:genVencimiento('<%=enc.EncodeForJavascript(p_nfactura)%>');"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgGenerarVencimientos)%>" <%=ParamImgGenerarVencimientos%> alt="<%=LitGenerar%>" title="<%=LitGenerar%>"></a>
							         <a href="javascript:VerComVenci('<%=enc.EncodeForJavascript(viene)%>','<%=enc.EncodeForJavascript(p_nfactura)%>');"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgVencComercial)%>" <%=ParamImgVencComercial%> alt="<%=LitVerComVenci%>" title="<%=LitVerComVenci%>"></a>
							    
                                    <%if cv="1" then%>
                                     <a href="javascript:VerComVenci('<%=enc.EncodeForJavascript(viene)%>','<%=enc.EncodeForJavascript(p_nfactura)%>');"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgVencComercial)%>" <%=ParamImgVencComercial%> alt="<%=LitVerComVenci%>" title="<%=LitVerComVenci%>"></a>
							    
								           <%end if%>
						        </td>
						        <td class='CELDAL7 width10 underOrange'>
							        <input class="CELDAR7 width65"  type="text" name="fechaVto" value="" onchange="cambiarfecha(document.facturas_cli.fechaVto.value,'Fecha Vencimiento')"/>
                                    <%DrawCalendar "fechaVto" %>
						        </td>
						        <td class='CELDAL7 width10 underOrange'>
							        <input class="CELDAR7 width100"  type="text" name="importeVto" value="0" onchange="RoundNumValue(this, <%=enc.EncodeForJavascript(NdecDiFactura)%>);"/>
						        </td>
						        <td class='CELDAL7 width10 underOrange' >
							        <input class="CELDAR7 width100"  type="text" name="recibidoVto" value="0" onchange="RoundNumValue(this, <%=enc.EncodeForJavascript(NdecDiFactura)%>);"/>
						        </td>
						        <td class="CELDAL7 width10 underOrange">
							        <input class="CELDAL7"  type="checkbox" name="cobradoVto"/>
						        </td>
						        <td class='CELDAL7 width10 underOrange' >
							        <textarea class="CELDAL7 width100"  name="obsVto" rows="2"></textarea>
						        </td>
						        <td class=" width10 underOrange">
							        <a class="ic-accept" href="javascript:addVencimiento('<%=enc.EncodeForJavascript(p_nfactura)%>');" onblur="javascript:document.facturas_cli.fechaVto.focus();"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion)%><%=enc.EncodeForHtmlAttribute(ImgNuevo)%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"></a>
						        </td>
                                </tr><%
					        'CloseFila
				        else
					        'DrawFila color_blau
						        %>
                                <tr>
                                <td class='CELDAL7'>
							        <%if cv="1" then%>
								        <a class="ic-accept" href="javascript:VerComVenci('<%=enc.EncodeForJavascript(viene)%>','<%=enc.EncodeForJavascript(p_nfactura)%>');"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion)%><%=enc.EncodeForHtmlAttribute(ImgVencComercial)%>" <%=ParamImgVencComercial%> alt="<%=LitVerComVenci%>" title="<%=LitVerComVenci%>"></a>
							        <%end if%>
						        </td>
                                </tr>
                                <%
					        'CloseFila
				        end if
				    %></table>
				    <iframe id="frVencimientos" class="width90 iframe-data md-table-responsive" name="fr_Vencimientos" src='facturas_cliven.asp?mode=browse&ndoc=<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>&ncliente=<%=enc.EncodeForHtmlAttribute(tmp_ncliente)%>' height='80' frameborder="yes" noresize="noresize"></iframe>
			
                </div><%
			

			    '***** Pagos a cuenta de la factura GPD (27/04/2007).
			    %>
             
                <div id="tabs4" class="overflowXauto" >
				    <table class="width90 md-table-responsive" border='0' cellspacing="1" cellpadding="1">
				        <% 'DrawFila "" %>
                        <tr>
				            <td width="25%" bgcolor="<%=enc.EncodeForHtmlAttribute(color_fondo)%>">
					        </td>
					        <td width="25%" align="right" bgcolor="<%=color_fondo%>">
					        <% If bolActivoValesDTO Then %>
					        <a class="CELDAREFB" href="javascript:BuscarValeDescuento('<%=enc.EncodeForJavascript(p_nfactura)%>');">Canjear Vale Dto</a>
					        <% End If %>
					        </td>
                        </tr>
				        <% 'Closefila %>
				    </table>
				    <table class="width90 md-table-responsive" style="border-collapse:collapse;table-layout:fixed;" ><%
				        'Fila de encabezado
				        'DrawFila color_terra
					    %>
                        <tr>
                            <td class="ENCABEZADOL width10 underOrange" ><%=LitNPago%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitFecha%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitDescripcion%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitImporte%></td>
					        <td class="ENCABEZADOL width10 underOrange" ><%=LitTipoPago%></td>
					        <td class="ENCABEZADOL width10 underOrange" >&nbsp</td>
                        </tr>
                        <%
				        'CloseFila
				        'Linea de inserción de un pago a cuenta
				        if rst("ahora")=0 and rst("cobrada")=0 and pagado<>1 and nliquidacion &""=""  and nliquidacionAG &""="" and p_pagsl then
					        'DrawFila color_blau
						    %>
                            <tr>
                                <td class='CELDAL7 underOrange width10' >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						        <td class='CELDAL7 underOrange width10'">
							        <input class="CELDAR7 width65"  type="text" name="fechaPago" value="" onchange="cambiarfecha(document.facturas_cli.fechaPago.value,'Fecha Pago')"/>
                                    <%DrawCalendar "fechaPago" %>
						        </td>
						        <td class='CELDAL7 underOrange width10' >
							        <textarea class="CELDAL7 width100"  name="descripcionPago" onFocus="lenmensaje(this,0,50,'')" onKeydown="lenmensaje(this,0,50,'')" onKeyup="lenmensaje(this,0,50,'')" onBlur="lenmensaje(this,0,50,'')" rows="2"></textarea>
						        </td>
						        <td class='CELDAL7 underOrange width10' >
							        <input class="CELDAR7 width100"  type="text" name="importePago" value="0" onchange="RoundNumValue(this, <%=enc.EncodeForJavascript(NdecDiFactura)%>);importepagoComp();"/>
						        </td>
                                <td class='CELDAL7 underOrange width10'>
						        <%
                                rstSelect.cursorlocation=3
                                rstSelect.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
						        'DrawSelectCelda "CELDAL7 style='font-Family: Verdana;font-size: 7.0pt;text-align: right;color: #000000;width: 100px;'","","",0,"","tipoPago",rstSelect,"","codigo","descripcion","",""
						        DrawSelect "'CELDAL7 width100'", "", "tipoPago", rstSelect, "", "codigo", "descripcion", "", ""
                                rstSelect.close%>
                                </td>
						        <td class="underOrange width10">
							        <a class="ic-accept" href="javascript:addPago('<%=enc.EncodeForJavascript(p_nfactura)%>');" onblur="javascript:document.facturas_cli.fechaPago.focus();"><img src="<%=enc.EncodeForHtmlAttribute(themeIlion) %><%=enc.EncodeForHtmlAttribute(ImgNuevo)%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"></a>
						        </td>
                            </tr>
					        <%
                            'CloseFila
				        end if%>
				    </table>
				    <iframe id="frPagosCuenta" class="width90 iframe-data md-table-responsive" name="fr_PagosCuenta" src='facturas_clipago.asp?mode=browse&ndoc=<%=enc.EncodeForHtmlAttribute(rst("nfactura"))%>'  height='80' frameborder="yes" noresize="noresize"></iframe>
                </div>
		        <br/>


                <!-- dir envio-->
                <div id="tabs5"class="overflowXauto" >
                <%'***** Mostrar la dirección de envío de la factura..'
			    'ega 19/06/2008 like con pertenece y solamente los campos necesarios
                rstDomi.cursorlocation=3
			    rstDomi.Open "select domicilio,poblacion,cp,provincia from domicilios with(nolock) where codigo='" & rst("dir_envio") & "' and pertenece like '" & session("ncliente") & "%'",session("dsn_cliente")
			    if rst("dir_envio")>"" then
				    pagina="../central.asp?pag1=./ventas/facturas_clidireccion_env.asp&ndoc=" &rst("nfactura") & "&mode=browse&pag2=./ventas/facturas_clidireccion_env_bt.asp&titulo=" & LITDIRENVIO & " " & trimCodEmpresa(rst("nfactura"))
			    else
				    pagina="../central.asp?pag1=./ventas/facturas_clidireccion_env.asp&ndoc=" &rst("nfactura") & "&mode=edit&pag2=./ventas/facturas_clidireccion_env_bt.asp&titulo=" & LITDIRENVIO & " " & trimCodEmpresa(rst("nfactura"))
			    end if%>

			        <table class="width90 md-table-responsive" border='0' cellspacing="0" cellpadding="0">
			            <%'DrawFila color_terra%>
                        <tr>
					        <td>
						        <table border="1" cellspacing="0" cellpadding="0">
						            <%'DrawFila color_blau2%>
                                    <tr>
								        <td class="ENCABEZADOL" height="25">	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=LitDatosEnvio%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <%
							                if rst("ahora")=0 and rst("cobrada")=0 and pagado<>1 and nliquidacion &""="" and nliquidacionAG &""="" and p_pagsl then
							            %>
									        <a class='CELDAREFB' href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(pagina)%>','P',<%=enc.EncodeForJavascript(altoventana)%>,<%=enc.EncodeForJavascript(anchoventana)%>)" OnMouseOver="self.status='<%=LitEditar%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitEditar%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								        <%end if%>
								        </td>
                                    </tr>
							        <%'CloseFila
						        %></table>
					        </td>
                        </tr>
                        <%
				        'CloseFila
				        if not rstDomi.eof then
					        'DrawFila color_terra %>
                            <tr>
						        <td>
    					        <table width='100%' border='0' cellspacing="1" cellpadding="1">
                                    <%
                                    'DrawFila color_terra
							            'DrawCelda "ENCABEZADOL","","",0,LitDomicilio
							            'DrawCelda "ENCABEZADOL","","",0,LitPoblacion
							            'DrawCelda "ENCABEZADOL","","",0,LitCP
							            'DrawCelda "ENCABEZADOL","","",0,LitProvincia
						            'CloseFila
                                    %>
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
                                    'CloseFila
                                    %>
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
                <!-- fin dir envio-->

                <%if SuplidosActivados=1 then%>
                <!-- suplidos -->
                        <div id="tabs6" class="overflowXauto">
                            <table class="width90 md-table-responsive"><%
				                'Fila de encabezado'
				                'DrawFila color_terra
					            %>
                                <tr>
                                    <td class="ENCABEZADOL underOrange" style="width:50px;"><%=LitItem%></td>
					                <td class="ENCABEZADOL underOrange" style="width:305px;"><%=LitDescripcion%></td>
					                <td class="ENCABEZADOL underOrange" style="width:95px;"><%=LitImporte%></td>
					                <td class="ENCABEZADOL underOrange" style="width:20px;">&nbsp</td>
                                </tr>
                                <%
				                'CloseFila
				                'Linea de inserción de un concepto'
				                if rst("ahora")=0 and rst("cobrada")=0 and pagado<>1 and nliquidacion &""=""  and nliquidacionAG &""="" and p_pagsl then
					                'DrawFila color_blau
					                %>
                                    <tr>
                                        <td class='CELDAL7 underOrange' >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						                <td class='CELDAL7 underOrange' >
							                <textarea class="CELDAL7"  name="descripcionSup" onFocus="lenmensaje(this,0,255,'')" onKeydown="lenmensaje(this,0,255,'')" onKeyup="lenmensaje(this,0,255,'')" onBlur="lenmensaje(this,0,255,'')" rows="2"></textarea>
						                </td>
						                <td class='CELDAL7 underOrange'>
							                <input class="CELDAR7"  type="text" name="importeSup" value="0"/>
							                <input type="hidden" name="importeSup_ant" value="0"/>
						                </td>
						                <td class=" underOrange">
							                <a class="ic-accept" href="javascript:addSuplido('<%=enc.EncodeForJavascript(p_nfactura)%>');" onblur="javascript:document.facturas_cli.descripcion.focus();"><img src="<%=enc.EncodeForJavascript(themeIlion) %><%=enc.EncodeForJavascript(ImgNuevo)%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"></a>
						                </td>
                                    </tr>
                                    <%
					                'CloseFila
				                end if
				            %></table>
				            <iframe id="frSuplidos" name="fr_Suplidos" class="width90 iframe-data md-table-responsive" src='facturas_suplidos.asp?mode=browse&ndoc=<%=enc.EncodeForJavascript(rst("nfactura"))%>' frameborder="yes" noresize="noresize"></iframe>
                        </div>
                <!-- fin suplidos -->
                <%end if%>
                </div>
                </div>
                </div>
                <!--Fin DATOS DEL DOCUMENTO-->
		    <%end if 'If mode=browse%>
		    <!--</td></tr></table>-->
            



            <!--
            <div class="accordion clearfix" id="S_DATTOTAL" >
                <h5><%=ucase(LitTotal)%></h5>
                -->
                <!--
                <a href="#" rel="toggle[DATTOTAL]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                    <div class="SectionHeader2">
                        <%=ucase(LitTotal)%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
                -->

            <div class="Section" id="S_DATTOTAL" style="display: flow-root;" >
                <!--<a href="#" rel="toggle[DATTOTAL]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">-->
                    <!--<div class="SectionHeader">-->
                    <div class="SectionHeader2">
                        <%=ucase(LitTotal)%>
                        <!--<img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />-->
                    </div>
                <!--</a>-->
            <div class="SectionPanel" id="DATTOTAL">
		        <%
                if mode="add" then
                    ValorImporteBruto=0
                else
                    ValorImporteBruto=rst("importe_bruto")
                end if

                'Mostrar los precios de la factura.%>
		        <input type="hidden" name="sumadet" value="<%=enc.EncodeForHtmlAttribute(sumadet)%>"/>
		        <input type="hidden" name="sumaRE" value="<%=enc.EncodeForHtmlAttribute(sumaRE)%>"/>
		        <input type="hidden" name="importe_bruto2" value="<%=enc.EncodeForHtmlAttribute(ValorImporteBruto)%>"/>

		        <!--<table class=TDBORDE width="100%"><tr><td>-->
		        <!--<table class="section-table">-->
                <!--<div class="col-xs-12 overflowX">-->
                    <!--<div class="inlineTable">-->
		                <%
		                'Fila de encabezado'

                            if mode="add" then
                                'ValorAbreviatura=d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%'" & iif(tmp_divisa>""," and codigo='" & tmp_divisa & "'"," and moneda_base<>0 "),session("dsn_cliente"))

                                if tmp_divisa>"" then
                                    ValorAbreviaturaSelect="select abreviatura from divisas with(nolock) where codigo like ? +'%' and codigo = ?"
                                    ValorAbreviatura=DlookupP2(ValorAbreviaturaSelect, session("ncliente")&"", adVarchar, 15, tmp_divisa&"", adVarchar, 15, session("dsn_cliente"))
                                else
                                    ValorAbreviaturaSelect="select abreviatura from divisas with(nolock) where codigo like ? +'%' and moneda_base<>0"
                                    ValorAbreviatura=DlookupP1(ValorAbreviaturaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))
                                end if

                                ValorImporteBruto=0
                                ValorDto1=iif((tmp_dto1>=0 and tmp_dto1<>"" ),tmp_dto1,0)
                                ValorDto2=iif((tmp_dto2>=0 and tmp_dto2<>"" ),tmp_dto2,0)
                                ValorDto3=iif((tmp_dto3>=0 and tmp_dto3<>"" ),tmp_dto3,0)
                                ValorTotDto=0
                                ValorBaseImp=0
                                ValorTotalIva=0
                                ValorTotalRe=0
                                ValorTotalSuplidos=0
                                ValorRecargo=iif((tmp_rf>=0 and tmp_rf<>""),tmp_rf,0)
                                ValorTotalRecargo=0
                                ValorIRPF=iif((tmp_irpf>=0 and tmp_irpf<>""),tmp_irpf,0)
                                ValorTotalIRPF=0
                                ValorTotalFactura=0
                            else
                                'ValorAbreviatura=d_lookup("abreviatura","divisas",iif(tmp_divisa>"","codigo like '" & session("ncliente") & "%' and codigo='" & tmp_divisa & "'",iif(mode="add"," moneda_base<>0 ","codigo='" & rst("divisa") & "'")) & " and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))

                                if tmp_divisa>"" then
                                    ValorAbreviaturaSelect="select abreviatura from divisas with(nolock) where codigo like ?+'%' and codigo=?"
                                    ValorAbreviatura=DlookupP2(ValorAbreviaturaSelect, session("ncliente")&"", adVarchar, 15, tmp_divisa&"", adVarchar, 15, session("dsn_cliente"))
                                else
                                    if mode="add" then
                                        ValorAbreviaturaSelect="select abreviatura from divisas with(nolock) where moneda_base<>0 and codigo like ?+'%'"
                                        ValorAbreviatura=DlookupP1(ValorAbreviaturaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))
                                    else
                                        ValorAbreviaturaSelect="select abreviatura from divisas with(nolock) where codigo=? and codigo like ?+'%'"
                                        ValorAbreviatura=DlookupP2(ValorAbreviaturaSelect, rst("divisa")&"", adVarchar, 15, session("ncliente")&"", adVarchar, 15, session("dsn_cliente"))
                                    end if
                                end if
                    
                                ValorImporteBruto=null_z(rst("importe_bruto"))
                                ValorDto1=iif((tmp_dto1>=0 and tmp_dto1<>"" ),tmp_dto1,iif(rst("descuento")>"",rst("descuento"),0))
                                ValorDto2=iif((tmp_dto2>=0 and tmp_dto2<>"" ),tmp_dto2,iif(rst("descuento2")>"",rst("descuento2"),0))
                                ValorDto3=iif((tmp_dto3>=0 and tmp_dto3<>"" ),tmp_dto3,iif(rst("descuento3")>"",rst("descuento3"),0))
                                ValorTotDto=null_z(rst("total_descuento"))
                                ValorBaseImp=null_z(rst("base_imponible"))
                                ValorTotalIva=Null_z(rst("total_iva"))
                                ValorTotalRe=Null_z(rst("total_re"))
                                ValorTotalSuplidos=Null_z(rst("total_suplidos"))
                                ValorRecargo=iif((tmp_rf>=0 and tmp_rf<>""),tmp_rf,iif(rst("recargo")>"",rst("recargo"),0))
                                ValorTotalRecargo=null_z(rst("total_recargo"))
                                ValorIRPF=iif((tmp_irpf>=0 and tmp_irpf<>""),tmp_irpf,iif(rst("IRPF")>"",rst("IRPF"),0))
                                ValorTotalIRPF=null_z(rst("total_irpf"))
                                ValorTotalFactura=null_z(rst("total_factura"))
                            end if


				            'DrawCelda "ENCABEZADOL","","",0,ValorAbreviatura
                            %>
                            <!--
                                <div class="col-lg-1 col-md-2 col-sm-3 col-xs-6 col-xxs-12">
                                    <label><%=LitTotal%></label>
                                    <label><%=ValorAbreviatura%></label>
                                </div>
                            -->
                            <%
                            DrawDiv "4", "", ""
                            DrawLabel "", "", LitTotal
                            DrawSpan "ENCABEZADOL", "", ValorAbreviatura, ""
                            CloseDiv

				            %><input type="hidden" name="h_importe_bruto" value="<%=enc.EncodeForHtmlAttribute(ValorImporteBruto)%>"/><%
				            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='ImpBruto'","txtRight disabled"),"","",0,LitBruto,"importe_bruto",10,formatnumber(ValorImporteBruto,n_decimales,-1,0,iif(mode="browse",-1,0))
                            DrawDiv "4", "", ""
                            DrawLabel "", "", LitBruto
                            EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "importe_bruto", formatnumber(ValorImporteBruto,n_decimales,-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='ImpBruto'","disabled")
                            CloseDiv

				            if mode<>"browse" then
                            %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total-2"><%
                                    %><label><%=LitDto%></label><%
					                %><input class="txtRight" type="text" name="dto1" value="<%=enc.EncodeForHtmlAttribute(ValorDto1)%>" onchange="Precios();"/><%
                                %></div><%
				            else
					            'EligeCeldaResponsive "input", mode,"txtRight id='Dto'","","",0,LitDto,"dto1",4, cstr(formatnumber(rst("descuento"),decpor,-1,0,iif(mode="browse",-1,0))) + "%"
				                DrawDiv "4", "", ""
                                DrawLabel "", "", LitDto
                                EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "dto1", cstr(formatnumber(rst("descuento"),decpor,-1,0,iif(mode="browse",-1,0))) + "%", "id='Dto'"
                                CloseDiv
                            end if

				            if mode<>"browse" then
                                %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total-2"><%
                                    %><label><%=LitDto2%></label><%
					                %><input class="txtRight" type="text" name="dto2" value="<%=enc.EncodeForHtmlAttribute(ValorDto2)%>" onchange="Precios();"/><%
				                %></div><%
                            else
					            'EligeCeldaResponsive "input", mode,"txtRight id='Dto2'","","",0,LitDto2,"dto2",4, formatnumber(rst("descuento2"),decpor,-1,0,iif(mode="browse",-1,0)) & "%"
				                DrawDiv "4", "", ""
                                DrawLabel "", "", LitDto2
                                EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "dto2", formatnumber(rst("descuento2"),decpor,-1,0,iif(mode="browse",-1,0)) & "%", "id='Dto2'"
                                CloseDiv
                            end if

				            if mode<>"browse" then
                            %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total-2"><%
                                    %><label><%=LitDto3%></label><%
					                %><input class="txtRight" type="text" name="dto3" value="<%=enc.EncodeForHtmlAttribute(ValorDto3)%>" onchange="Precios();"/><%
				                %></div><%
                            else
					            'EligeCeldaResponsive "input", mode,"txtRight id='Dto3'","","",0,LitDto3,"dto3",4, formatnumber(null_z(rst("descuento3")),decpor,-1,0,iif(mode="browse",-1,0)) & "%"
				            	DrawDiv "4", "", ""
                                DrawLabel "", "", LitDto3
                                EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "dto3", formatnumber(null_z(rst("descuento3")),decpor,-1,0,iif(mode="browse",-1,0)) & "%", "id='Dto3'"
                                CloseDiv
                            end if

				            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='TotalDto'","txtRight disabled"),"","",0,LitTotalDescuento,"total_descuento",10,formatnumber(ValorTotDto,n_decimales,-1,0,iif(mode="browse",-1,0))
                            DrawDiv "4", "", ""
                            DrawLabel "", "", LitTotalDescuento
                            EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_descuento", formatnumber(ValorTotDto,n_decimales,-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalDto'","disabled")
                            CloseDiv
                            %><input type="hidden" name="h_total_descuento" value="<%=enc.EncodeForHtmlAttribute(ValorTotDto)%>"/><%
                            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='BImponible'","txtRight disabled"),"","",0,LitImponible,"base_imponible",10,formatnumber(ValorBaseImp,n_decimales,-1,0,iif(mode="browse",-1,0))
                            DrawDiv "4", "", ""
                            DrawLabel "", "", LitImponible
                            EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "base_imponible", formatnumber(ValorBaseImp,n_decimales,-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='BImponible'","disabled")
                            CloseDiv
                            %><input type="hidden" name="h_base_imponible" value="<%=enc.EncodeForHtmlAttribute(ValorBaseImp)%>"/><%
                            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='TotalIva'","txtRight disabled"),"","",0,LitTotalIva,"total_iva",10,cstr(formatnumber(ValorTotalIva,n_decimales,-1,0,iif(mode="browse",-1,0)))
                            DrawDiv "4", "", ""
                            DrawLabel "", "", LitTotalIva
                            EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_iva", cstr(formatnumber(ValorTotalIva,n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='TotalIva'","disabled")
                            CloseDiv
                            %><input type="hidden" name="h_total_iva" value="<%=enc.EncodeForHtmlAttribute(ValorTotalIva)%>"/><%
                            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='TotalRE'","txtRight disabled"),"","",0,LitTotalRE,"total_re",10,cstr(formatnumber(ValorTotalRe,n_decimales,-1,0,iif(mode="browse",-1,0)))
                            DrawDiv "4", "", ""
                            DrawLabel "", "", LitTotalRE
                            EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_re", cstr(formatnumber(ValorTotalRe,n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='TotalRE'","disabled")
                            CloseDiv
                            %><input type="hidden" name="h_total_re" value="<%=enc.EncodeForHtmlAttribute(ValorTotalRe)%>"/><%
                            if SuplidosActivados=1 then
                                'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='totalSuplidos'","txtRight disabled"),"","",0,LITFICSUPLIDOSFAC,"total_suplidos",10,cstr(formatnumber(ValorTotalSuplidos,n_decimales,-1,0,iif(mode="browse",-1,0)))
                                DrawDiv "4", "", ""
                                DrawLabel "", "", LITFICSUPLIDOSFAC
                                EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_suplidos", cstr(formatnumber(ValorTotalSuplidos,n_decimales,-1,0,iif(mode="browse",-1,0))), iif(mode="browse","id='totalSuplidos'","disabled")
                                CloseDiv
                                %><input type="hidden" name="h_total_suplidos" value="<%=enc.EncodeForHtmlAttribute(ValorTotalSuplidos)%>"/><%
                            else
                                %><input type="hidden" class="txtRight" id='totalSuplidos' name="total_suplidos" value="<%=enc.EncodeForHtmlAttribute(ValorTotalSuplidos)%>"/>
                                    <input type="hidden" name="h_total_suplidos" value="<%=enc.EncodeForHtmlAttribute(ValorTotalSuplidos)%>"/><%
                            end if
                            if mode="add" then
                                ValorRecargo=0
                            else
                                ValorRecargo=rst("recargo")
                            end if
                            if ((ValorRecargo<>0) or mode="edit" or mode="add") then
					            if mode<>"browse" then
                                    %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total-2"><%
                                        %><label><%=LitRF%></label><%
					                    %><input class="txtRight" type="text" name="rf" value="<%=enc.EncodeForHtmlAttribute(ValorRecargo)%>" onchange="Precios();"/><%
				                    %></div><%
                                
                                else
						            'EligeCeldaResponsive "input", mode,"txtRight id='RF'","","",0,LitRF,"rf",4, cstr(formatnumber(rst("recargo"),decpor,-1,0,iif(mode="browse",-1,0))) + "%"
					                DrawDiv "4", "", ""
                                    DrawLabel "", "", LitRF
                                    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "rf", cstr(formatnumber(rst("recargo"),decpor,-1,0,iif(mode="browse",-1,0))) + "%", iif(mode="browse","id='RF'","disabled")
                                    CloseDiv
                                end if
					            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='TotalRF'","txtRight disabled"),"","",0,LitTotalRF,"total_rf",10,formatnumber(ValorTotalRecargo,n_decimales,-1,0,iif(mode="browse",-1,0))
				                DrawDiv "4", "", ""
                                DrawLabel "", "", LitTotalRF
                                EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_rf", formatnumber(ValorTotalRecargo,n_decimales,-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalRF'","disabled")
                                CloseDiv
                            end if
                            %><input type="hidden" name="h_total_rf" value="<%=enc.EncodeForHtmlAttribute(ValorTotalRecargo)%>"/><%
                            if mode="add" then
                                ValorIRPF=0
                            else
                                ValorIRPF=rst("irpf")
                            end if
                            if ((ValorIRPF<>0) or mode="edit" or mode="add") then
					            if mode<>"browse" then
                                    %><div class="col-lg-1 col-xs-2 col-xxs-4 col-total-2"><%
                                        %><label><%=LitIRPF%></label><%
					                    %><input class="txtRight" type="text" name="irpf" value="<%=enc.EncodeForHtmlAttribute(ValorIRPF)%>" onchange="Precios();"/><%
				                    %></div><%
					            else
						            'EligeCeldaResponsive "input", mode,"txtRight id='IRPF'","","",0,LitIRPF,"irpf",4, cstr(formatnumber(null_z(rst("IRPF")),2,-1,0,iif(mode="browse",-1,0))) + "%"
					                DrawDiv "4", "", ""
                                    DrawLabel "", "", LitIRPF
                                    EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "irpf", cstr(formatnumber(null_z(rst("IRPF")),2,-1,0,iif(mode="browse",-1,0))) + "%", "id='IRPF'"
                                    CloseDiv
                                end if
					            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='TotalIRPF'","txtRight disabled"),"","",0,LitTotalIRPF,"total_irpf",10,formatnumber(ValorTotalIRPF,n_decimales,-1,0,iif(mode="browse",-1,0))
				            	DrawDiv "4", "", ""
                                DrawLabel "", "", LitTotalIRPF
                                EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_irpf", formatnumber(ValorTotalIRPF,n_decimales,-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalIRPF'","disabled")
                                CloseDiv
                            end if
                            %><input type="hidden" name="h_total_irpf" value="<%=enc.EncodeForHtmlAttribute(ValorTotalIRPF)%>"/><%
                            'EligeCeldaResponsive "input", mode,iif(mode="browse","txtRight id='TotalFactura'","txtRight disabled"),"","",0,LitTotal,"total_factura",15,formatnumber(ValorTotalFactura,n_decimales,-1,0,iif(mode="browse",-1,0))
                            DrawDiv "4", "", ""
                            DrawLabel "", "", LitTotal
                            EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "total_factura", formatnumber(ValorTotalFactura,n_decimales,-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalFactura'","disabled")
                            CloseDiv
                            %><input type="hidden" name="h_total_factura" value="<%=enc.EncodeForHtmlAttribute(ValorTotalFactura)%>"/><%

                            'd_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))

                            ifSelect="select imp_equiv from configuracion with(nolock) where nempresa=?"
                            
                            if DlookupP1(ifSelect, session("ncliente")&"", adChar, 5, session("dsn_cliente")) then
                                
                                DIVISASelect="select codigo from divisas with(nolock) where moneda_base<>0 and codigo like ?+'%'"
                                
                                if mode="add" then
					                DIVISA=iif(tmp_divisa>"",tmp_divisa,iif(mode="add",DlookupP1(DIVISASelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")),""))
                                    'd_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                                else
                                    DIVISA=iif(tmp_divisa>"",tmp_divisa,iif(mode="add",DlookupP1(DIVISASelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")),rst("divisa")))
                                    'd_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
                                end if

                            	DrawDiv "1", "", ""
                                DrawLabel "", "", LitTotal
                                DrawSpan "txtRight", "", LitPTAS, ""
                                CloseDiv

                                celdaSelect="select ndecimales from divisas with(nolock) where codigo= ?+'01'"
                                'd_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("dsn_cliente")) misma en todas als Dlookup siguientes

					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitBruto
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Pimporte_bruto", formatnumber(CambioDivisa(ValorImporteBruto,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='ImpBrutoEq'","disabled")
                                CloseDiv
                            
                                'EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='DtoEq'","ENCABEZADOR disabled"),"","",0,"","Pdto1",4, formatnumber(ValorDto1,decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%","")
					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitDto
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Pdto1", formatnumber(ValorDto1,decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%",""), iif(mode="browse","id='DtoEq'","disabled")
                                CloseDiv
                            
                                'EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Dto2Eq'","ENCABEZADOR disabled"),"","",0,"","Pdto2",4, formatnumber(ValorDto2,decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%","")
					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitDto2
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Pdto2", formatnumber(ValorDto2,decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%",""), iif(mode="browse","id='Dto2Eq'","disabled")
                                CloseDiv

                                'EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='Dto3Eq'","ENCABEZADOR disabled"),"","",0,"","Pdto3",4, formatnumber(ValorDto3,decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%","")
					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitDto3
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Pdto3", formatnumber(ValorDto3,decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%",""), iif(mode="browse","id='Dto3Eq'","disabled")
                                CloseDiv
                            
                                'EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='TotalDtoEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_descuento",10,formatnumber(CambioDivisa(ValorTotDto,DIVISA,session("ncliente") & "01"),DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitTotalDescuento
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Ptotal_descuento", formatnumber(CambioDivisa(ValorTotDto,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalDtoEq'","disabled")
                                CloseDiv
                            
                                'EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='BImponibleEq'","ENCABEZADOR disabled"),"","",0,"","Pbase_imponible",10,formatnumber(CambioDivisa(ValorBaseImp,DIVISA,session("ncliente") & "01"),DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitImponible
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Pbase_imponible", formatnumber(CambioDivisa(ValorBaseImp,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='BImponibleEq'","disabled")
                                CloseDiv
                                
                                'EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='TotalIvaEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_iva",10,formatnumber(CambioDivisa(ValorTotalIva,DIVISA,session("ncliente") & "01"),DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitTotalIva
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Ptotal_iva", formatnumber(CambioDivisa(ValorTotalIva,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalIvaEq'","disabled")
                                CloseDiv
                            
                                'EligeCelda "input", mode,iif(mode="browse","CELDARIGHT id='TotalREEq'","ENCABEZADOR disabled"),"","",0,"","Ptotal_re",10,formatnumber(CambioDivisa(ValorTotalRe,DIVISA,session("ncliente") & "01"),DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")),-1,0,iif(mode="browse",-1,0))
					            DrawDiv "1", "", ""
                                DrawLabel "", "", LitTotalRE
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Ptotal_re", formatnumber(CambioDivisa(ValorTotalRe,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalREEq'","disabled")
                                CloseDiv
                            
                                if ((ValorRecargo<>0) or mode="edit" or mode="add") then

					                DrawDiv "1", "", ""
                                    DrawLabel "", "", LitRF
                                    EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Prf", formatnumber(ValorRecargo,decpor,-1,0,iif(mode="browse",-1,0))  & iif(mode="browse","%",""), iif(mode="browse","id='RFEq'","disabled")
                                    CloseDiv
                                    
                                    DrawDiv "1", "", ""
                                    DrawLabel "", "", LitTotalRF
                                    EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Ptotal_rf", formatnumber(CambioDivisa(ValorTotalRecargo,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalRFEq'","disabled")
                                    CloseDiv
                                end if
					            if ((ValorIRPF<>0) or mode="edit" or mode="add") then

                                    DrawDiv "1", "", ""
                                    DrawLabel "", "", LitIRPF
                                    EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Pirpf", formatnumber(ValorIRPF,decpor,-1,0,iif(mode="browse",-1,0)) & iif(mode="browse","%",""), iif(mode="browse","id='IRPFEq'","disabled")
                                    CloseDiv
                                    
                                    DrawDiv "1", "", ""
                                    DrawLabel "", "", LitTotalIRPF
                                    EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Ptotal_irpf", formatnumber(CambioDivisa(ValorTotalIRPF,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalIRPFEq'","disabled")
                                    CloseDiv
                                end if

				                DrawDiv "1", "", ""
                                DrawLabel "", "", LitTotal
                                EligeCeldaResponsive1 "input", enc.EncodeForHtmlAttribute(null_s(mode)), "txtRight", "", "Ptotal_albaran", formatnumber(CambioDivisa(ValorTotalFactura,DIVISA,session("ncliente") & "01"),enc.EncodeForHtmlAttribute(null_s(DlookupP1(celdaSelect, session("ncliente")&"", adVarchar, 15, session("dsn_cliente")))),-1,0,iif(mode="browse",-1,0)), iif(mode="browse","id='TotalFacturaEq'","disabled")
                                CloseDiv

                        end if%>
            </div>
            </div>
            
		    <%if mode="add" then%>
			    <script language="javascript" type="text/javascript">
                                       document.facturas_cli.fecha.focus();
                                       document.facturas_cli.fecha.select();
                                       ColocarCapa();
			    </script>
		    <%elseif mode="edit" then%>
			    <script language="javascript" type="text/javascript">
                    document.facturas_cli.fecha.focus();
                    document.facturas_cli.fecha.select();
                    ColocarCapa();
			    </script>
		    <%elseif mode="browse" then%>
                <iframe name="frameExportar" style='display:none;' frameborder='0' width='500' height='200'></iframe>
			    <input type="hidden" name="tot_doc_ahora" value="<%=enc.EncodeForHtmlAttribute(null_z(rst("deuda")))%>"/>
			    <input type="hidden" name="tot_doc_ahora_ant" value="<%=enc.EncodeForHtmlAttribute(null_z(rst("deuda")))%>"/>
			    <input type="hidden" name="dto1_riesgo" value="<%=enc.EncodeForHtmlAttribute(null_z(rst("descuento")))%>"/>
			    <input type="hidden" name="dto2_riesgo" value="<%=enc.EncodeForHtmlAttribute(null_z(rst("descuento2")))%>"/>
			    <input type="hidden" name="dto3_riesgo" value="<%=enc.EncodeForHtmlAttribute(null_z(rst("descuento3")))%>"/>
			    <input type="hidden" name="ndecimales_riesgo" value="<%=enc.EncodeForHtmlAttribute(null_z(n_decimales))%>"/>
			    <script type="text/javascript" language="javascript">
                    jQuery(window).load(function () {
                        Redimensionar();
                        try {
                            if (document.getElementById("frDetallesIns").style.display != "none") {
                                fr_DetallesIns.document.facturas_clidetins.cantidad.focus();
                                fr_DetallesIns.document.facturas_clidetins.cantidad.select();
                            }
                        }
                        catch (e) {
                        }
                    });
			    </script>
		    <%end if
	    ''end if
	    end if
	    if mode="search" then
		
	    end if
        ''if mode="add" then rst.CancelUpdate%>
        <input type="hidden" name="total_paginas" value="<%=enc.EncodeForHtmlAttribute(total_paginas)%>"/>
    </form>

    <%'i(EJM 04/09/2006)
    'formulario para enviar los datos a la línea de detalle para grabarlos
    if totalCrearLineasFacZona=1 then%>
		<!--Para insertar los gtos de envío en la línea de pedido si fuese necesario-->
		<form name="crearLineaFac" method="post" target="fr_Detalles">
			<input type="hidden" name="nfactura" value="<%=enc.EncodeForHtmlAttribute(num_factura)%>"/>
			<input type="hidden" name="h_cantidad2" value="0"/>
			<input type="hidden" name="h_calcularimpcantidad2" value="0"/>
			<input type="hidden" name="h_cantidad" value="1"/>
			<input type="hidden" name="h_referencia" value="<%=enc.EncodeForHtmlAttribute(trimCodEmpresa(datosCrearLineasFacZona(0,0)))%>"/>
			<input type="hidden" name="h_almacen" value="<%=enc.EncodeForHtmlAttribute(datosCrearLineasFacZona(2,0))%>"/>
			<input type="hidden" name="h_descripcion" value="<%=enc.EncodeForHtmlAttribute(datosCrearLineasFacZona(1,0))%>"/>
			<%precioTarifa =PrecioArticulo (datosCrearLineasFacZona(0,0),date(),1,Nulear(request.form("tarifa")),precio,precioFinal)%>
    		<input type="hidden" name="h_pvp" value="<%=enc.EncodeForHtmlAttribute(precioTarifa)%>"/>
			<input type="hidden" name="h_descuento" value="<%=enc.EncodeForHtmlAttribute(datosCrearLineasFacZona(7,0))%>"/>
			<input type="hidden" name="h_descuento2" value="0"/>
			<input type="hidden" name="h_descuento3" value="0"/>
			<input type="hidden" name="h_iva" value="<%=enc.EncodeForHtmlAttribute(datosCrearLineasFacZona(5,0))%>"/>
			<input type="hidden" name="h_importe" value="<%=enc.EncodeForHtmlAttribute(precioTarifa)%>"/>
			<input type="hidden" name="h_lote" value=""/>
			<input type="hidden" name="nserie" value=""/>

			<input type="hidden" name="modp" value="<%=enc.EncodeForHtmlAttribute(modp)%>"/>
	        <input type="hidden" name="modd" value="<%=enc.EncodeForHtmlAttribute(modd)%>"/>
	        <input type="hidden" name="modi" value="<%=enc.EncodeForHtmlAttribute(modi)%>"/>
            <input type="hidden" name="pciva" value="<%=enc.EncodeForHtmlAttribute(pciva)%>"/>
            <input type="hidden" name="modn" value="<%=enc.EncodeForHtmlAttribute(modn)%>"/>
		</form>
		<!-- Fin  Para insertar los gtos de envío en la línea de pedido si fuese necesario-->
    <%end if
    'fin(EJM 04/09/2006)
end if
CerrarTodo()

%>


</body>
</html>