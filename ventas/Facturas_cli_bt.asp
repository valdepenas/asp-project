<%@ Language=VBScript %>
<%

    dim  enc
    set enc = Server.CreateObject("Owasp_Esapi.Encoder")

' ############ SOLO PARA DEVOLVER CONSULTAS AJAX SOBRE ESTA MISMA PÁGINA ############
function limpiaCadenaFacBT(strValor)
	dim returnedValue

	returnedValue=strValor

	returnedValue=replace(returnedValue,"'","''")
	returnedValue=replace(returnedValue,"--","")
	returnedValue=replace(returnedValue,";","")
	returnedValue=replace(returnedValue,"select","")
	returnedValue=replace(returnedValue,"drop","")
	returnedValue=replace(returnedValue,"insert","")
	returnedValue=replace(returnedValue,"update","")
	returnedValue=replace(returnedValue,"delete","")
	returnedValue=replace(returnedValue,"xp_","")
	returnedValue=replace(returnedValue,"sp_","")
	returnedValue=replace(returnedValue,"shutdown","")
	returnedValue=replace(returnedValue,"bulk","")
	returnedValue=replace(returnedValue,"bcp","")
	returnedValue=replace(returnedValue,"script","")
	returnedValue=replace(returnedValue,"declare","")
	returnedValue=replace(returnedValue,"exec","")

	'Para evitar tratar los tags de html como tales.
	'returnedValue=server.htmlencode(returnedValue)

	limpiaCadenaFacBT=returnedValue
end function
if request.querystring("mode") = "consultaAJAX" then
    
    ' Consultamos el nfolio1 y nfolio2 de una serie
    if request.querystring("consulta") = "nfolioMinMax" then
        set rstAux = Server.CreateObject("ADODB.Recordset")
        nserie = limpiaCadenaFacBT(request.querystring("nserie"))
        if nserie&"">"" then 
            set connDom = Server.CreateObject("ADODB.Connection")
            set commandDom = Server.CreateObject("ADODB.Command")

            connDom.open session("dsn_cliente")
            connDom.cursorlocation=3

            commandDom.ActiveConnection =connDom
            commandDom.CommandTimeout = 60
            commandDom.CommandText = "select case when nfolio1 is null then 0 else case when nfolio1='' then 0 else nfolio1 end end as nfolio1,case when nfolio2 is null then 0 else case when nfolio2='' then 0 else nfolio2 end end as nfolio2 FROM series with(nolock) where nserie=?"
            commandDom.CommandType = adCmdText
            commandDom.Parameters.Append commandDom.CreateParameter("@nserie",adVarchar,adParamInput,10,nserie)

            set rstaux = commandDom.Execute
            'rstAux.open "select case when nfolio1 is null then 0 else case when nfolio1='' then 0 else nfolio1 end end as nfolio1,case when nfolio2 is null then 0 else case when nfolio2='' then 0 else nfolio2 end end as nfolio2 FROM series with(nolock) where nserie='" & nserie & "'",session("dsn_cliente")
            if not rstAux.EOF then
                nfolio1 = rstAux("nfolio1")
                nfolio2 = rstAux("nfolio2")
            end if
            rstAux.Close
            
            response.Write(nfolio1 & "," & nfolio2)
        else
            response.Write("ERROR")
        end if
        set rstAux =nothing
  
    ' Fin de consulta AJAX
    response.End
    end if
end if
' ################################################

' VGR 04-03-2003 : Cambios para no editar la factura si ya está liquidada para su comercial
' VGR 07-03-2003 : Cambios para no editar la factura si ya está liquidada para su agente
'ricardo 3-6-2003 se pone por parametro que pregunte el cambio del comercial de los vencimientos
%>
<script id="DebugDirectives" runat="server" language="javascript">
    // Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=enc.EncodeForHtmlAttribute(session("lenguaje"))%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=enc.EncodeForHtmlAttribute(session("caracteres"))%>"/>
</head>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<!--#include file="facturas_cli.inc" -->
<% 

set rst = Server.CreateObject("ADODB.Recordset")

set connDom = Server.CreateObject("ADODB.Connection")
set commandDom = Server.CreateObject("ADODB.Command")

connDom.open session("dsn_cliente")
connDom.cursorlocation=3

commandDom.ActiveConnection =connDom
commandDom.CommandTimeout = 60
commandDom.CommandText = "SELECT gestion_folios FROM configuracion with(nolock) where nempresa=?"
commandDom.CommandType = adCmdText
commandDom.Parameters.Append commandDom.CreateParameter("@nempresa",adChar,adParamInput,10,session("ncliente"))

set rst = commandDom.Execute

'rst.open "SELECT gestion_folios FROM configuracion with(nolock) where nempresa='" & session("ncliente") & "'",session("dsn_cliente")
if not rst.EOF and rst("gestion_folios") = true then
    gestionFolios = true
else
    gestionFolios = false
end if
rst.Close
set rst =nothing


%>

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
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
        //si se ha pulsado la tecla enter
        //if (window.event.keyCode==13){
        //document.opciones.criterio.focus();
        Buscar();
        //}
    }

    //FLM:20090506: oculta los botones al pulsar
    function OcultaBotones(){
        var m;
        for(m=0; m<document.all.length; m++){
            if(document.all[m].tagName=="IMG"){
                document.all[m].style.display="none";  
            }        
        }
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
        SearchPage("facturas_cli_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
        "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value,1);
        document.opciones.texto.value = "";
    }

    var numFolioCorrecto = false;

    <% if gestionFolios then %>
    function ComprobarNumFolio(){

        if(IsNumeric(parent.pantalla.document.facturas_cli.nfolio.value)){
            nserie = parent.pantalla.document.facturas_cli.serie.value;
            if (!enProceso && http) 
            {
                var timestamp = Number(new Date()); 
                var url = "Facturas_cli_bt.asp?mode=consultaAJAX&consulta=nfolioMinMax&nserie=" + nserie + "&ts=" + timestamp;
                http.open("GET", url, false); 
                http.onreadystatechange = handleHttpResponse;
                enProceso = true;
                http.send(null);
            }
        }
        else{
            alert("<%=LITMSGFOLIONUMERICO %>");
            numFolioCorrecto = false;
            parent.pantalla.document.facturas_cli.nfolio.focus();
        }
    }
    <% end if %>

    function handleHttpResponse() 
    {
        if (http.readyState == 4) 
    {
            if (http.status == 200) 
    {
                if (http.responseText.indexOf('invalid') == -1) 
    {
        // Armamos un array, usando la coma para separar elementos
                    results = http.responseText;
                    enProceso = false;
             
                    if(results == "" || results == "ERROR") alert("<%=LitErrorNumFolio%>");
    else
    {
                        var retValue = results.split(",");
                        nfolioIntroducido = parent.pantalla.document.facturas_cli.nfolio.value;
                        if(nfolioIntroducido < parseInt(retValue[0]) || nfolioIntroducido > parseInt(retValue[1]))
    {
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

        function getHTTPObject() 
        {
            var xmlhttp;
            if (!xmlhttp && typeof XMLHttpRequest != 'undefined') 
            {
                try 
                {
                    xmlhttp = new XMLHttpRequest();
                } 
                catch (e) { xmlhttp = false; }
            }
            return xmlhttp;
        }

    var enProceso = false; // lo usamos para ver si hay un proceso activo
    var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest


    //Validación de campos numéricos y fechas.
    function ComprobarCuentaBancaria() {
        ok = true;
        while (parent.pantalla.document.facturas_cli.ncuenta1.value.search(" ") != -1) {
            parent.pantalla.document.facturas_cli.ncuenta1.value = parent.pantalla.document.facturas_cli.ncuenta1.value.replace(" ", "");
        }
        while (parent.pantalla.document.facturas_cli.ncuenta2.value.search(" ") != -1) {
            parent.pantalla.document.facturas_cli.ncuenta2.value = parent.pantalla.document.facturas_cli.ncuenta2.value.replace(" ", "");
        }
        while (parent.pantalla.document.facturas_cli.ncuenta3.value.search(" ") != -1) {
            parent.pantalla.document.facturas_cli.ncuenta3.value = parent.pantalla.document.facturas_cli.ncuenta3.value.replace(" ", "");
        }
        while (parent.pantalla.document.facturas_cli.ncuenta4.value.search(" ") != -1) {
            parent.pantalla.document.facturas_cli.ncuenta4.value = parent.pantalla.document.facturas_cli.ncuenta4.value.replace(" ", "");
        }
        while (parent.pantalla.document.facturas_cli.ncuenta5.value.search(" ") != -1) {
            parent.pantalla.document.facturas_cli.ncuenta5.value = parent.pantalla.document.facturas_cli.ncuenta5.value.replace(" ", "");
        }
        while (parent.pantalla.document.facturas_cli.ncuenta6.value.search(" ") != -1) {
            parent.pantalla.document.facturas_cli.ncuenta6.value = parent.pantalla.document.facturas_cli.ncuenta6.value.replace(" ", "");
        }

        if (parent.pantalla.document.facturas_cli.ncuenta3.value != "") {
            if (isNaN(parent.pantalla.document.facturas_cli.ncuenta3.value)) {
                window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
                ok = false;
            }
        }
        if (parent.pantalla.document.facturas_cli.ncuenta4.value != "") {
            if (isNaN(parent.pantalla.document.facturas_cli.ncuenta4.value)) {
                window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
                ok = false;
            }
        }
        if (parent.pantalla.document.facturas_cli.ncuenta5.value != "") {
            if (isNaN(parent.pantalla.document.facturas_cli.ncuenta5.value)) {
                window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
                ok = false;
            }
        }
        if (parent.pantalla.document.facturas_cli.ncuenta6.value != "") {
            if (isNaN(parent.pantalla.document.facturas_cli.ncuenta6.value)) {
                window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
                ok = false;
            }
        }

        SumaC = 0;

        if (parent.pantalla.document.facturas_cli.ncuenta3.value != "") {
            SumaC++;
        }
        if (parent.pantalla.document.facturas_cli.ncuenta4.value != "") {
            SumaC++;
        }
        if (parent.pantalla.document.facturas_cli.ncuenta5.value != "") {
            SumaC++;
        }
        if (parent.pantalla.document.facturas_cli.ncuenta6.value != "") {
            SumaC++;
        }

        if (SumaC != 0 && SumaC != 4) {
            window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
            ok = false;
        }


        var country = parent.pantalla.document.facturas_cli.ncuenta1.value;
        var i_iban = parent.pantalla.document.facturas_cli.ncuenta2.value;
        if (country == "") {
            country = "ES";
            //parent.pantalla.document.facturas_cli.ncuenta1.value = "ES";
        }
        var account = parent.pantalla.document.facturas_cli.ncuenta3.value + parent.pantalla.document.facturas_cli.ncuenta4.value + parent.pantalla.document.facturas_cli.ncuenta5.value + parent.pantalla.document.facturas_cli.ncuenta6.value;

        if (account.toString().length != 0 && account.toString().length < 11) {
            window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
            return false;
        }
        if (country == "ES" && account.toString().length != 0)
        {
            var iban = CreateIBAN(country, account);

            if (parent.pantalla.document.facturas_cli.ncuenta2.value != "" && account.toString().length == 20) {
                var i_iban = parent.pantalla.document.facturas_cli.ncuenta2.value;

                if (!ValidateIBAN(iban, i_iban)) {
                    window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
                    parent.pantalla.document.facturas_cli.ncuenta2.select();
                    parent.pantalla.document.facturas_cli.ncuenta2.focus();
                    return false;
                }
            }
            else {
                    //window.alert("<%=LITMSGCUENTABANCOINCORRECTA%>");
                    parent.pantalla.document.facturas_cli.ncuenta1.value = country;
                    parent.pantalla.document.facturas_cli.ncuenta2.value = iban;
                    //return false;
            }
        }

        if (account.toString().length == 0 && parent.pantalla.document.facturas_cli.banco.value != "") {
            parent.pantalla.document.facturas_cli.banco.value = "";
        }

        return ok;
    }
    function ValidarCampos(mode)
    {
        if (parent.pantalla.document.facturas_cli.cobrada.checked){
            if (parent.pantalla.document.facturas_cli.h_cobrada.value == 0){
                if (!confirm("<%=LitMsgCobFactVencConfirm%>")){
                    return false;
                }
            }
            else{
                alert("<%=LitMsgFacturaCobroNoModif%>");
                return false;
            }
        }
        else{
            if (parent.pantalla.document.facturas_cli.h_cobrada.value == 1){
                if (!confirm("<%=LitMsgAnCobFactVencConfirm%>")){
                    return false
                }
            }
        }

        if (parent.pantalla.document.facturas_cli.fecha.value=="") {
            alert("<%=LitMsgFechaNoNulo%>");
            return false;
        }

        if (!cambiarfecha(parent.pantalla.document.facturas_cli.fecha.value,"FECHA FACTURA")){
            return false;
        }

        if (parent.pantalla.document.facturas_cli.fecha.value!="") {
            if (!checkdate(parent.pantalla.document.facturas_cli.fecha)) {
                alert("<%=LitMsgFechaFecha%>");
                return;
            }
        }

        if (!cambiarfecha(parent.pantalla.document.facturas_cli.fechaenvio.value,"FECHA ENVIO")){
            return false;
        }

        if (parent.pantalla.document.facturas_cli.fechaenvio.value!=""){
            if (!checkdate(parent.pantalla.document.facturas_cli.fechaenvio)) {
                alert("<%=LitMsgFechaFecha%>");
                return;
            }
        }

        if (!cambiarfecha(parent.pantalla.document.facturas_cli.fechapedido.value,"FECHA PEDIDO")){
            return false;
        }

        if (parent.pantalla.document.facturas_cli.fechapedido.value!=""){
            if (!checkdate(parent.pantalla.document.facturas_cli.fechapedido)) {
                alert("<%=LitMsgFechaFecha%>");
                return;
            }
        }

        if (parent.pantalla.document.facturas_cli.serie.value=="") {
            alert("<%=LitMsgSerieNoNulo%>");
            return false;
        }

        if (parent.pantalla.document.facturas_cli.divisa.value=="") {
            alert("<%=LitMsgDivisaNoNulo%>");
            return false;
        }
        if (parent.pantalla.document.facturas_cli.ncliente.value=="") {
            alert("<%=LitMsgClienteNoNulo%>");
            return false;
        }

        if (isNaN(parent.pantalla.document.facturas_cli.dto1.value.replace(",",".")) || isNaN(parent.pantalla.document.facturas_cli.dto2.value.replace(",",".")) || isNaN(parent.pantalla.document.facturas_cli.dto3.value.replace(",",".")) || isNaN(parent.pantalla.document.facturas_cli.rf.value.replace(",","."))){
            alert("<%=LitMsgDto1Dto2RfNumerico%>");
            return false;
        }
	
        factcambio=parent.pantalla.document.facturas_cli.nfactcambio.value.replace(",","."); 		
        if (!/^([0-9])*[.]?[0-9]*$/.test(factcambio))
        { 
            alert("<%=LitMsgFactCambioI%>"); 
            return false;
        }
        if (parent.pantalla.document.facturas_cli.nfactcambio.value=="")
        {
            alert("<%=LitMsgFactCambioI%>"); 
            return false;
        }
    
        <% if gestionFolios then %>
        ComprobarNumFolio()
        if(!numFolioCorrecto)
        {
            return false;
        }
        <% end if %>
    
            // JMA 2/11/04. Campos personalizables.
        if ((mode=="add"||mode=="edit")&&(parent.pantalla.document.facturas_cli.si_campo_personalizables.value==1)){
            num_campos=parent.pantalla.document.facturas_cli.num_campos.value;

            respuesta=comprobarCampPerso("parent.pantalla.",num_campos,"facturas_cli");
            if(respuesta!=0){
                titulo="titulo_campo" + respuesta;
                tipo="tipo_campo" + respuesta;
                titulo=parent.pantalla.document.facturas_cli.elements[titulo].value;
                tipo=parent.pantalla.document.facturas_cli.elements[tipo].value;
                if (tipo==4) nomTipo="<%=LitTipoNumerico%>";
                else if (tipo==5) {
                    nomTipo="<%=LitTipoFecha%>";
                }
                alert("<%=LitMsgCampo%> " + titulo + " <%=LitMsgTipo%> " + nomTipo);
                return false;
            }
        }

        if (ComprobarCuentaBancaria()==false)
        {
            return false;
        }
        return true;
    }

    //Realizar la acción correspondiente al botón pulsado.
    function Accion(mode,pulsado) {

        gen_vencimiento=parent.pantalla.document.facturas_cli.gen_vencimiento.value;

        switch (mode) {
            case "browse":
                switch (pulsado) {
                    case "add": //Nuevo registro
                        if (parent.pantalla.document.facturas_cli.mode.value!="search"){
                            parent.pantalla.document.facturas_cli.h_ncliente.value="";
                            parent.pantalla.document.facturas_cli.ncliente.value="";
                        }
                        //FLM:20090506:oculta botones
                        OcultaBotones();
                        parent.pantalla.document.facturas_cli.action="facturas_cli.asp?mode=" + pulsado;
                        parent.pantalla.document.facturas_cli.submit();
                        document.location="facturas_cli_bt.asp?mode=" + pulsado;
                        break;

                    case "edit": //Editar registro
                        /*alert(parent.pantalla.document.facturas_cli.h_ahora.value);*/
				       
                        if (parent.pantalla.document.facturas_cli.p_pagsl.value==0){
                            alert("<%=LitMsgNoEditar%>");
                            break;
                        }
				    
                        if (parent.pantalla.document.facturas_cli.h_contabilizada.value=="1" && document.opciones.bloqContab.value=="1")
                            alert("<%=Lit_NoModContab%>");
                        else {
                            if (parent.pantalla.document.facturas_cli.h_ahora.value==1) alert("<%=LitMsgFactNoModif%>");
                            else {
                                if (parent.pantalla.document.facturas_cli.nliquidacion.value!="") alert("<%=LitMsgFactLiquidada%>");
                                else {

                                    if (parent.pantalla.document.facturas_cli.nliquidacionAG.value!=""){
                                        window.alert("<%=LitMsgFactLiquidadaAG%>");
                                    }
                                    else {
                                        //FLM:20090506:oculta botones
                                        OcultaBotones();					
                                        parent.pantalla.document.facturas_cli.action="facturas_cli.asp?nfactura=" + parent.pantalla.document.facturas_cli.h_nfactura.value +
                                        "&mode=" + pulsado;
                                        parent.pantalla.document.facturas_cli.submit();
                                        document.location="facturas_cli_bt.asp?mode=" + pulsado;
                                    }
                                }
                            }
                        }
                        break;

                    case "delete": //Eliminar registro
                        if (parent.pantalla.document.facturas_cli.p_pagsl.value==0){
                            alert("<%=LitMsgNoEditar%>");
                            break;
                        }
				
                        if (parent.pantalla.document.facturas_cli.h_contabilizada.value=="1" && document.opciones.bloqContab.value=="1")
                            alert("<%=Lit_NoModContab%>");
                        else {
                            if (parent.pantalla.document.facturas_cli.h_ahora.value==1) {
                                alert("<%=LitMsgFactNoModif%>");
                            }
                            else {
                                if (parent.pantalla.document.facturas_cli.nliquidacion.value!="") alert("<%=LitMsgFactLiquidada%>");
                                else{

                                    if (parent.pantalla.document.facturas_cli.nliquidacionAG.value!="") alert("<%=LitMsgFactLiquidadaAG%>");
                                    else
                                    {
                                        if (parent.pantalla.document.facturas_cli.nasiento.value!="") alert("<%=LitMsgFactTieneAsiento%>");
                                        else
                                        {
                                            if (confirm("<%=LitMsgEliminarFacturasConfirm%>")==true)
                                            {
                                                //FLM:20090506:oculta botones
                                                OcultaBotones();					
                                                parent.pantalla.document.facturas_cli.action="facturas_cli.asp?nfactura=" + parent.pantalla.document.facturas_cli.h_nfactura.value +
                                                "&mode=" + pulsado;
                                                parent.pantalla.document.facturas_cli.submit();
                                                document.location="facturas_cli_bt.asp?mode=browse";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;

                    //MPC
                    case "export": //Export invoice to Excel
                        parent.pantalla.frameExportar.document.location = "invoice_customer_export.asp?ninvoice=" + parent.pantalla.document.facturas_cli.h_nfactura.value;
                        break;
                }
                break;

            case "edit":
                switch (pulsado) {
                    case "save": //Guardar registro
                        cont="SI";
                        if ((parent.pantalla.document.facturas_cli.h_cobrada.value=="0") && (parent.pantalla.document.facturas_cli.cobrada.checked)) {
                            if (!confirm("<%=LitMsgCobFacturaSinCajaConfirm%>")) cont="NO";
                        }
                        if (cont=="SI") {
                            if (ValidarCampos(mode)) {
                                //ricardo 25/4/2003 para que salga el mensaje de que se ha cambiado las propiedades del documento
                                // y que puede afectar al importe de los detalles
                                nempresa="<%=session("ncliente")%>";
                                recalcular_importes=1;
                                if (parent.pantalla.document.facturas_cli.h_ncliente.value!=(nempresa + parent.pantalla.document.facturas_cli.ncliente.value) ||
                                    parent.pantalla.document.facturas_cli.h_fecha.value!=parent.pantalla.document.facturas_cli.fecha.value ||
                                    parent.pantalla.document.facturas_cli.h_divisa.value!=parent.pantalla.document.facturas_cli.olddivisa.value ||
                                    parent.pantalla.document.facturas_cli.h_tarifa.value!=parent.pantalla.document.facturas_cli.tarifa.value)
                                {
                                    if (confirm("<%=LitMsgCamPropDocCamPrec%>")==false) recalcular_importes=0;
                                }
                                continuar=1;
                                continuarf=1;
                                continuari=1;
                                if (gen_vencimiento=='-1' && (parent.pantalla.document.facturas_cli.forma_pago.value!=parent.pantalla.document.facturas_cli.forma_pago_ant.value)){
                                    if(confirm("<%=LITMSGGENVENFORPAGCORFIMR%>")) continuar=0;
                                    else continuar=1;
                                }
                                cambiarcom=0;
                                preguntar_cambio_comercial=parent.pantalla.document.facturas_cli.cv.value;
                                if (parseInt(preguntar_cambio_comercial)==1){
                                    if (parseInt(parent.pantalla.document.facturas_cli.si_vencimientos.value)==1 &&
                                            ((parent.pantalla.document.facturas_cli.forma_pago.value!=parent.pantalla.document.facturas_cli.forma_pago_ant.value &&
                                                gen_vencimiento=='-1')
                                                || parent.pantalla.document.facturas_cli.comercial.value!=parent.pantalla.document.facturas_cli.h_comercial.value)){

                                        //ricardo 6-10-2003 si los vencimientos no tienen comercial y asignamos uno a la factura
                                        //no hace falta que preguntemos, directamente se cambia
                                        if (parseInt(parent.pantalla.document.facturas_cli.comercial_ven.value)==0){
                                            if (window.confirm("<%=LitMsgCambiarVencCom%>")) cambiarcom=1;
                                            else cambiarcom=0;
                                        }
                                        else cambiarcom=1;
                                    }
                                }
                                else cambiarcom=1;
                                if (gen_vencimiento=='-1' && (parent.pantalla.document.facturas_cli.fecha.value!=parent.pantalla.document.facturas_cli.fecha_ant.value)){
                                    if(confirm("<%=LitMsgGenVenFechaConfirm%>")) continuarf=0;
                                    else continuarf=1;
                                }
                                if (gen_vencimiento=='-1' && (parent.pantalla.document.facturas_cli.total_factura.value!=parent.pantalla.document.facturas_cli.importe_ant.value))
                                    continuari=0;
                                else continuari=1;
                                //FLM:20090506:oculta botones
                                OcultaBotones();
					
                                parent.pantalla.document.facturas_cli.action="facturas_cli.asp?nfactura=" + parent.pantalla.document.facturas_cli.h_nfactura.value + "&mode=save&continuar=" + continuar + "&continuarf=" + continuarf + "&continuari=" + continuari + "&cambiarcom=" + cambiarcom;
                                parent.pantalla.document.facturas_cli.action=parent.pantalla.document.facturas_cli.action + "&recalcular_importes=" + recalcular_importes;
                                parent.pantalla.document.facturas_cli.submit();
                                document.location="facturas_cli_bt.asp?mode=browse";
                            }
                        }
                        break;

                    case "cancel": //Cancelar edición
                        //FLM:20090506:oculta botones
                        OcultaBotones();
                        parent.pantalla.document.facturas_cli.divisafc.value="";
                        parent.pantalla.document.facturas_cli.action="facturas_cli.asp?nfactura=" + parent.pantalla.document.facturas_cli.h_nfactura.value +
                        "&mode=browse";
                        parent.pantalla.document.facturas_cli.submit();
                        document.location="facturas_cli_bt.asp?mode=browse";
                        break;
                }
                break;

            case "add":
                switch (pulsado) {
                    case "save": //Guardar registro
                        if (ValidarCampos(mode)) {

                            limiteFacturasCreadas=parseInt(parent.pantalla.document.facturas_cli.limiteFacturasCreadas.value);
                            CantidadFacturasCreadas=parseInt(parent.pantalla.document.facturas_cli.CantidadFacturasCreadas.value);
                            vienenp = parent.pantalla.document.facturas_cli.h_vienenp.value;
                            datenote = parent.pantalla.document.facturas_cli.h_datenote.value;					       			      
                            if (CantidadFacturasCreadas<limiteFacturasCreadas){
                                //FLM:20090506:oculta botones
                                OcultaBotones();
                                parent.pantalla.document.facturas_cli.action="facturas_cli.asp?mode=first_save&vienenp="+vienenp+"&datent="+datenote;
                                parent.pantalla.document.facturas_cli.submit();
                                document.location="facturas_cli_bt.asp?mode=browse";
                            }
                            else{
                                mensaje_a_salir="<%=LitSoloPuedeIns1Fact%>" + limiteFacturasCreadas;
                                if (limiteFacturasCreadas==1) mensaje_a_salir=mensaje_a_salir + " <%=LitSoloPuedeIns2Fact%>";
                                else mensaje_a_salir=mensaje_a_salir + " <%=LitSoloPuedeIns3Fact%>";
                                alert(mensaje_a_salir);
                            }
                        }
                        break;

                    case "cancel": //Cancelar edición
                        //FLM:20090506:oculta botones
                        OcultaBotones();
                        parent.pantalla.document.facturas_cli.divisafc.value="";
                        parent.pantalla.document.facturas_cli.ncliente.value="";
                        parent.pantalla.document.facturas_cli.serie.value="";
                        parent.pantalla.document.facturas_cli.action="facturas_cli.asp?mode=add";
                        parent.pantalla.document.facturas_cli.submit();
                        document.location="facturas_cli_bt.asp?mode=add";
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
    obtenerparametros("facturas_cli_det")%>

<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
<input type="hidden" name="bloqContab" value="<%=enc.EncodeForHtmlAttribute(bloqContab)%>" />
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
                'MPC Export invoice
                if cstr(exportinv) <> "0" then%>
                    <td id="idexport" class="CELDABOT" onclick="javascript:Accion('browse','export');">
					    <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
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
			<!--<td class="CELDABOT" onclick=""><%=LitBuscar2 & ": "%>-->
				<select class="IN_S" name="campos">
          			<option selected value="nfactura"><%=LitFactura%></option>
          			<option value="c.rsocial"><%=LitCliente%></option>
          			<%if gestionFolios = true then
                        %><option value="nfolio"><%=LITNUMFOLIO%></option><%
                    end if%>
        		</select>
        	<!--</td><td class="CELDABOT" onclick="">-->
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContieneB%></option>
					<%'option value="empieza">=LitComienza</option>%>
					<option value="termina"><%=LitTerminaB%></option>
					<option value="igual"><%=LitIgualB%></option>
				</select>
        	<!--</td><td class="CELDABOT" onclick="">-->
				<input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
        	<!--</td><td class="CELDABOT" onclick="">-->
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=enc.EncodeForHtmlAttribute(ImgBuscarLF_bt)%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
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