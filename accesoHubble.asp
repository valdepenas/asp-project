<%@  language="VBScript" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%' JCI 29/04/2003 : Se añade el tratamiento de modo 99 de acceso
' JA  05/05/2003 : Eliminar las tablas temporales del usuario al cerrar la sesión.
''ricardo 16-5-2003 se corrige el error de que se escribira nada en la casilla o que se pusieran caracteres
' MCA 03/12/2004 : Gestión de las franjas horarias de acceso
    'probando bea
const paginaAcceso="accesoHubble.asp"
strCodigo="-1"
hd_Act=0
dim parametroD
parametroD=replace(replace(limpiaCadena(request.querystring("D"))&"","{",""),"}","")

dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")

%>
<!--#include file="constantes.inc"-->
<!--#include file="modulos.inc" -->
<%if request.querystring("mode")&""<>"fin" then %>
<% 
   

Function GenerarURLbt
    if Request.ServerVariables("HTTPS")="on" then
    httpCad="https://"
    else
    httpCad="http://"
    end if 
	
    servidor=request.ServerVariables("server_name")
    ruta=httpCad+servidor 
    puerto=Request.ServerVariables("server_port")
    if(puerto<>"80" and puerto<>"443") then
        ruta=ruta+":"&puerto
    end if
	GenerarURLbt = ruta
End Function
%>
<script language="javascript" type="text/javascript">
    var xmlHttp, ServerResponse = null;
    function netinit() {

        //alert("vamos allá-10");
        xmlHttp = GetXmlHttpObject();
        if (xmlHttp != null) {
            var url = "/<%=CarpetaProduccionX4%>/init.aspx"
            //alert (url)
            xmlHttp.open("GET", url, false);  //synchronous method
            xmlHttp.send(null);
            //if (ServerResponse!=null) return ServerResponse; //asynchronous method
            return xmlHttp.responseText; //synchronous method
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
    //object.onclick=function(){alert("boton pulsado");};
    window.onbeforeunload = function (e) {
        netinit();
    }
    
	      
</script>
<%end if %>
<!--#include file="mensajes.inc"-->
<!--#include file="Gestion/matriz.inc"-->
<!--#include file="cache.inc" -->
<!--#include file="adovbs.inc" -->
<!--#include file="tablas.inc" -->

<%if len(parametroD) > 5 then
    'ndist=d_lookup("ndistribuidor", "distribuidores", "id_partner='" & parametroD & "'", dsnilion)

    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open dsnilion
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="GetDistributorData"
    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    command.Parameters.Append command.CreateParameter("@type", adVarChar, adParamInput, 2, "01")
    command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, "")
    command.Parameters.Append command.CreateParameter("@id", adVarChar, adParamInput, 36, parametroD)

    set rst = command.execute

    if not rst.eof then
        ndist = rst("ndistribuidor")
        session("locale_user") = rst("FLOCALE")
    end if

    rst.close
    conn.close
    set rst = nothing
    set command = nothing
    set conn = nothing

else
    ndist=parametroD
end if

'GET LANGUAGE AND HIDE BUTTON LITERALS

if ndist&"">"" then
    set connection = Server.CreateObject("ADODB.Connection")
    set commandDist =  Server.CreateObject("ADODB.Command")
    connection.open dsnilion
    commandDist.ActiveConnection =connection
    commandDist.CommandTimeout = 0
    commandDist.CommandText="GetDataLanguage"
    commandDist.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    commandDist.Parameters.Append commandDist.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, ndist)

    set rst = commandDist.execute

    if not rst.eof then
        
        session("locale_user") = rst("FLOCALE") & ""
        ndLitOk = rst("USERBUTTONLITERAL") & ""
    end if

    rst.close
    connection.close
    set rst = nothing
    set commandDist = nothing
    set connection = nothing

else
    session("locale_user") = "es-ES"
    ndLitOk = "1"
end if

  'repetir para cargar los  literales 000 con el idioma inicializado  
if session("locale_user") & ""<>"" then
   
	locale =session("locale_user")
    LoadLiterals "BOT", locale	
    LoadLiterals "000", locale
     
end if

'GET CONFIG DISTRIBUTOR

if ndist&"">"" then

    set connDistributor = Server.CreateObject("ADODB.Connection")
    set commandDistributor =  Server.CreateObject("ADODB.Command")
    connDistributor.open dsnilion
    commandDistributor.ActiveConnection =connDistributor
    commandDistributor.CommandTimeout = 0
    commandDistributor.CommandText="GetConfigDataDistributor"
    commandDistributor.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    commandDistributor.Parameters.Append commandDistributor.CreateParameter("@ndistributor", adVarChar, adParamInput,5, ndist)
    commandDistributor.Parameters.Append commandDistributor.CreateParameter("@id", adVarChar, adParamInput, 36, "")
    commandDistributor.Parameters.Append commandDistributor.CreateParameter("@dtid", adSmallint, adParamInput,, 2)

    set rstDistributor = commandDistributor.execute

    if not rstDistributor.eof then
        usecasilla = rstDistributor("usecasilla")&""
        adminpwd = rstDistributor("adminpwd")&""
        newuserliteral = rstDistributor("newuserliteral")&""
    end if

    rstDistributor.close
    connDistributor.close
    set rstDistributor = nothing
    set commandDistributor = nothing
    set connDistributor = nothing

end if

%>


<!--#include file="acceso.inc" -->
<!--#include file="ico.inc" -->
<!--#include file="calculos.inc" -->
<!--#include file="varios.inc"-->
<!--#include file="borrTablTemp.inc" -->
<!--#include file="servicios/mensajes_sms.inc"-->
<!--#include file="ilion.inc" -->


<script language="javascript" type="text/javascript">
    //20131003 DBS Funcion ajax
    // Replacing XMLHttpRequest function, native IE if the "Enable native XMLHTTP support" enabled.

    function CreateXmlHttpNotNative(){
        try{
            xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
        }
        catch(e){
            try{
                xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            catch(oc){
                xmlHttp = null;
            }
        }
        if(!xmlHttp && typeof XMLHttpRequest != "undefined"){
            xmlHttp = new XMLHttpRequest();
        }
        return xmlHttp;
    }
    /*******************************************/
    var http = CreateXmlHttpNotNative();
    //var http = GetXmlHttpObject();
    var enProceso = false;
    var result = false;

    function handleHttpResponseQty() {
        if (http.readyState == 4) {
            if (http.status == 200) {
                if (http.responseText != "") {
                    result = http.responseText;

                }
            }
        }
    }
    /******************************************/
</script>
<script type="text/javascript" language="javascript">




    var ev = null;

    if (window.document.addEventListener) {
        window.document.addEventListener("keydown", callkeydownhandler, false);
    }
    else {
        window.document.attachEvent("onkeydown", callkeydownhandler);
    }

    //Comprueba si la tecla pulsada es CTRL+S. Si es así guarda el registro.
    function callkeydownhandler(evnt) {
        ev = (evnt) ? evnt : event;
        keypressed();
    }


    
    //ricardo 16-5-2003
    function ValidarDatos() {
        ok = true;
        while (document.entrada.vcasilla.value.search(" ") != -1) {
            document.entrada.vcasilla.value = document.entrada.vcasilla.value.replace(" ", "");
        }
        if (isNaN(document.entrada.vcasilla.value)) {
            alert("<%=LitNumTarjetaErr%>");
            ok = false;
        }
        if (ok) {
            document.entrada.action = "accesoHubble.asp?mode=comprueba&d=<%=enc.EncodeForJavascript(parametroD)%>";
            document.entrada.submit();
        }
        else {
            document.entrada.vcasilla.focus();
            document.entrada.vcasilla.select();
        }
    }



    //ricardo 16-5-2003
    function keypressed() {
        //if (document.getElementById("hdmode")!=null)
        //{
        if (document.getElementById("hdmode").value != "showdata") {
            var keycode = ev.keyCode;
            //if (keycode == 13 || keycode == 32) {
            if (keycode == 13) {
                ValidarDatos();
            }
        }
        //}
    }


    function activarCuenta() {
        if (document.entrada.chkContract.checked == false) {
            window.alert("<%=LitMsgErrContract%>")
        } else {
            if (document.entrada.email.value == "") alert("<%=LitFaltaMail%>");
            else {
                if (window.confirm("<%=LitConfirmarActivar%>") == true) {
                    document.entrada.action = "<%=paginaAcceso%>?mode=activar&d=<%=enc.EncodeForJavascript(parametroD)%>";
                    document.entrada.submit();
                }
            }
        }
    }

    function RedirigirEmpresas(modo, param) {
        mostrar_maximizado = "<%=mostrar_maximizado%>";

        if (param == "tyg=1" && modo == "tiendas") modo = "empresas";
        if (param == "tyg=1" && modo == "unaempresa") modo = "empresas";
        if (modo == "app") {

            if (mostrar_maximizado == "-1") {
                //pantallaCompleta("sel_app.asp");
                pantallaCompleta(param);
                CerrarVentanaPrincipalSinPreguntar();

            }
            else {

                parent.location = param + "?d=<%=enc.EncodeForJavascript(parametroD)%>";
            }
        }
        if (modo == "empresas") {
            if (mostrar_maximizado == "-1") {
                pantallaCompleta("/" + "<%=CarpetaProduccionX4%>" + "5/UIPrincipal/selCompanyHubble.aspx?pr=1");
                CerrarVentanaPrincipalSinPreguntar();
            }
            else parent.location = "/" + "<%=CarpetaProduccionX4%>" + "5/UIPrincipal/selCompanyHubble.aspx?pr=1&d=<%=enc.EncodeForJavascript(parametroD)%>";

        }
        if (modo == "unaempresa") {
            if (mostrar_maximizado == "-1") {
                pantallaCompleta("/" + "<%=CarpetaProduccionX4%>" + "5/UIPrincipal/selCompanyHubble.aspx?pr=1");
                CerrarVentanaPrincipalSinPreguntar();
            }
            else parent.location = "/" + "<%=CarpetaProduccionX4%>" + "5/UIPrincipal/selCompanyHubble.aspx?pr=1&d=<%=enc.EncodeForJavascript(parametroD)%>";
        }
        if (modo == "tiendas") {
            if (mostrar_maximizado == "-1") {
                pantallaCompleta("sel_tiendas.asp");
                CerrarVentanaPrincipalSinPreguntar();
            }
            else parent.location = "sel_tiendas.asp";
        }
        if (modo == "asesoria") {
            if (mostrar_maximizado == "-1") {
                pantallaCompleta("asesoria/sel_asesoria.asp");
                CerrarVentanaPrincipalSinPreguntar();
            }
            else parent.location = "asesoria/sel_asesoria.asp";
        }
        if (modo == "tiendasTarjMaestra") {
            if (mostrar_maximizado == "-1") {
                pantallaCompleta("sel_tiendas2.asp?si_tienda_cli=true&ncliente=" + param);
                CerrarVentanaPrincipalSinPreguntar();
            }
            else parent.location = "sel_tiendas2.asp?si_tienda_cli=true&ncliente=" + param;
        }
    }

    function DibujaLogo() {
        leftPos = 0;
        topPos = 0;
        eval("document.all('LayerLogo').style.left=leftPos");
        eval("document.all('LayerLogo').style.top=topPos");
    }

    function RecoveryPass() {
        var ran = Math.random();

        if("<%=enc.EncodeForHtmlAttribute(adminpwd)%>"=="True")
        {
            var p="&p=True";
        }
        else{
            p="&p=False";
        }


        paginaModal = "/<%=CarpetaProduccion%>/recoverypass.asp?d=<%=enc.EncodeForHtmlAttribute(ndist)%>&r=" + ran+p;
        if("<%=enc.EncodeForHtmlAttribute(adminpwd)%>"=="True"){
            cambiarTamanyo("#fr_RecoveryPass", "350", "475");
        }
        else
        {
            cambiarTamanyo("#fr_RecoveryPass", "250", "550");
        }
        reloadClass("#fr_RecoveryPass", paginaModal);
        alPresionar("#fr_RecoveryPass");
    }
    function newOrangePass() {

        var ran = Math.random();
        document.entrada.contrasenya.value="";
        document.entrada.contrasenya.disabled=true;
        document.entrada.nombre.value="";


        if("<%=enc.EncodeForHtmlAttribute(adminpwd)%>"=="True")
        {
            var p="&p=True";
        }
        else{
            p="&p=False";
        }

        paginaModal = "/<%=CarpetaProduccion%>/recoverypass.asp?d=<%=enc.EncodeForHtmlAttribute(ndist)%>&r=" + ran+"&mod=newpass"+p;
        cambiarTamanyo("#fr_RecoveryPass", "330", "450");
        reloadClass("#fr_RecoveryPass", paginaModal);
        alPresionar("#fr_RecoveryPass");
    }

    //MAP 02/10/2013 
    function bloqPassword_orange() {
        var user = document.entrada.nombre.value;
        //document.entrada.action = "<%=paginaAcceso%>?mode=bloqPassword_orange&d=<%=enc.EncodeForJavascript(parametroD)%>";
        //document.entrada.submit();

        if (ValidaNactivacion(user)=="OK"){            
            document.entrada.contrasenya.disabled=false;
            document.entrada.contrasenya.value="";
            document.entrada.nombre.value=user;
            document.entrada.contrasenya.blur();
            document.entrada.contrasenya.focus();
            document.entrada.contrasenya.select();
        } 
        else
        {
            document.entrada.contrasenya.disabled=true;
            document.entrada.contrasenya.value="";
            document.entrada.nombre.value=user;
            document.entrada.nombre.blur();
            document.entrada.nombre.focus();
            document.entrada.nombre.select();
        }        
    }
	
    // 20131003 Acceso a la pagina Orange
    function ValidaNactivacion(user)
    {
        //var user = document.entrada.nombre.value;        
        if (!enProceso && http) {
            var url = "Gestion/BusinessAppOrange/theme/Clientes/checkOrange.asp?lg=" + user + "&mode=ValidaNactivacion";
            http.open("POST", url, false);
            http.onreadystatechange = handleHttpResponseQty;
            enProceso = false;
            http.send(null);
        }
        else
            result = 0;

        return result;
    }


</script>

<% if ndist = "00025" then%>
<script type="text/javascript" language="javascript">
    function Register() {
        location.href = "https://cloud.grantbiomed.com/ilionp/CMS/webtienda/cuerpo.asp?Id={A4EA18E8-E013-4499-BF6B-875E80295772}&idioma=10031ES&ilionteca=step3&p=GPROFESSIONAL";
    }
</script>
<%end if %>

<%'mmg 04/02/2008
dim pruebaMM
Function InvokeWebService (strSoap, strSOAPAction, strURL, xmlResponse)
    Dim xmlhttp
    Dim blnSuccess

        'Creamos el objeto ServerXMLHTTP
        Set xmlhttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
        'Abrimos la conexión con el método POST, ya que estamos enviando una petición.
        xmlhttp.Open "POST", strURL
        xmlhttp.setRequestHeader "Man", "POST " & strURL & " HTTP/1.1"
        xmlhttp.setRequestHeader "Host", "localhost"
        xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        xmlhttp.setRequestHeader "SOAPAction", strSOAPAction

        'El SOAPAction es importante ya que el WebService lo utilizará para verificar qué WebMethod estamos usando en la operación.
        'Enviamos la petición

        xmlhttp.send(strSoap)
        pruebaMM=xmlhttp.Status

        'Verificamos el estado de la comunicación
        If xmlhttp.Status = 200 Then
            'El código 200 implica que la comunicación se puedo establecer y que
            'el WebService se ejecutó con éxito.
            blnSuccess = True
        Else
            'Si el código es distinto de 200, la comunicación falló o el
            'WebService provocó un Error.
            blnSuccess = False
        End If

        'Obtenemos la respuesta del servidor remoto, parseada por el MSXML.
        Set xmlResponse = xmlhttp.responseXML
        InvokeWebService = blnSuccess
        Set xmlhttp = Nothing 'Destruimos el objeto
End Function

'mmg 04/02/2008 : Comprobamos si la entrada se intenta realizar mediante Huella dactilar
'dgb 06/1072010: nuevo version Kyros
if (request.querystring("mode")="hd" or request.querystring("mode")="hd2") and request.querystring("id")<> "" then
    hd_Act=1

'Dimensionamos la variable donde obtendremos la respuesta del WebService
'Dim xmlResponse
'Realizamos la llamada a la función InvokeWebService(), brindándole los parámetros correspondientes
 'InvokeWebService strSoap, strSOAPAction, "http://192.168.0.101/kyros/service.asmx", xmlResponse

if request.querystring("mode")="hd" then
    'Comprobamos que el usuario tiene permiso para acceder al sistema
    strSoap    = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "<soap:Body>" & _
    "<AccesoIlion xmlns=""http://www.cetel.com"">" & _
    "<id>" & request.querystring("id") & "</id>" & _
   "</AccesoIlion>" & _
    "</soap:Body>" & _
    "</soap:Envelope>"

    strSoapAux=strSoap
    strSOAPAction = "http://www.cetel.com/AccesoIlion"
    If InvokeWebService (strSoap, strSOAPAction, "http://localhost/kyros/service.asmx", xmlResponse) Then
    'Si el WebService se ejecutó con éxito, obtenemos la respuesta y la imprimimos utilizando MSXML.DOMDocument
        strCodigo = xmlResponse.documentElement.selectSingleNode("soap:Body/AccesoIlionResponse/AccesoIlionResult").text
        strPass="OK"
    Else
        strPass="NO"

    End If
else 'dgb el nuevo caso para Version Kyros 2010
    'Comprobamos que el usuario tiene permiso para acceder al sistema
    strSoap    = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "<soap:Body>" & _
    "<AccesoIlion xmlns=""http://www.ilionsistemas.com/"">" & _
    "<id>" & request.querystring("id") & "</id>" & _
   "</AccesoIlion>" & _
    "</soap:Body>" & _
    "</soap:Envelope>"

    strSoapAux=strSoap
    strSOAPAction = "http://www.ilionsistemas.com/AccesoIlion"

    If InvokeWebService (strSoap, strSOAPAction, "http://localhost/IlionServices/ServidorKyros2010/Service_Kyros.asmx", xmlResponse) Then
    'Si el WebService se ejecutó con éxito, obtenemos la respuesta y la imprimimos utilizando MSXML.DOMDocument
        strCodigo = xmlResponse.documentElement.selectSingleNode("soap:Body/AccesoIlionResponse/AccesoIlionResult").text
        strPass="OK"
    Else
        strPass="NO"
    End If

end if

'Liberamos la memoria del objeto xmlResponse
Set xmlResponse = Nothing
if strPass="OK" then
    if strCodigo="-1" then
        'el usuario no es valido
        if (mode="hd" or mode="hd2") then
            response.Write("<center><font color='black' face='Verdana, Arial, Helvetica, sans-serif' size='2'><b>")
            response.Write("No se ha podido identificar al usuario, por favor inténtelo de nuevo.")
            response.Write("</b></font></center><br />")
        end if
    else
        if Acceder()="ACTIVAR" then
            mode="showdata"
        end if
    end if
end if
  'aqui va el codigo q cortaste y esta en el documento url
end if
folder = "HUBBLE"
title = ""
if parametroD & "" <> "" then
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open dsnilion
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="GetConfigDataDistributor"
    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    if len(parametroD) > 5 then
        command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, "")
    else
        command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, parametroD)
    end if
    if len(parametroD) > 5 then
        command.Parameters.Append command.CreateParameter("@id", adVarChar, adParamInput, 36, parametroD)
    else
        command.Parameters.Append command.CreateParameter("@id", adVarChar, adParamInput, 36, "")
    end if
    command.Parameters.Append command.CreateParameter("@dtid", adSmallInt, adParamInput, , "2")

    set rst = command.execute

    if not rst.eof then
        folder = rst("FOLDERCSS")
        title = rst("DISTRIBUTORTITLE")
    end if

    rst.close
    conn.close
    set rst = nothing
    set command = nothing
    set conn = nothing
end if
 session("folder")=folder
    %>
<html>
<head>
    <title><%=iif(title & "" <> "", title, TituloVentana)%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

    <meta http-equiv="CACHE-CONTROL" content="NO-CACHE"/>
    <meta http-equiv="PRAGMA" content="NO-CACHE"/>
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1"/>
    <meta name="theme-color" content="#e35e24"/>
    <meta name="msapplication-navbutton-color" content="#e35e24"/>
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent"/>
    <meta name="apple-mobile-web-app-capable" content="yes"/>
    <script type="text/javascript">
        /mobi/i.test(navigator.userAgent) && !location.hash && setTimeout(function () {
            if (!pageYOffset) window.scrollTo(0, 1);
        }, 1000);
    </script>


    <link rel="SHORTCUT ICON" href="/lib/estilos/<%=folder%>/images/caballin.ico"/>
    <script type="text/javascript" language="javascript">
<!--
    function MM_openBrWindow(theURL, winName, features) { //v2.0
        window.open(theURL, winName, features);
    }
    //-->
    </script>
    <!--#include file="styles/access.css.inc"-->
    <link rel="stylesheet" href="estilos.css"/>
    <link href="/lib/estilos/<%=folder%>/font-face.css" rel="stylesheet" type="text/css" />
    <!--#include file="js/generic.js.inc"-->
    <!--#include file="common/modal.inc" -->

    <script language="javascript" type="text/javascript" src="jfunciones.js"></script>
</head>
<%
set rstAux=Server.CreateObject("ADODB.Recordset")
mostrar_cabecera=0
mostrar_maximizado=0
set command = nothing
set conn = Server.CreateObject("ADODB.Connection")
set command =  Server.CreateObject("ADODB.Command")
conn.open dsnilion
command.ActiveConnection =conn
command.CommandTimeout = 0
command.CommandText="select mostrar_cabecera,mostrar_maximizado from configuracion with(nolock) where PATHACCESO =?"
command.CommandType = adCmdText
command.Parameters.Append command.CreateParameter("@carpetaproduccion",adVarChar,adParamInput,20,carpetaproduccion)
set rstAux= command.Execute
if not rstAux.eof then
    mostrar_cabecera=nz_b(rstAux("mostrar_cabecera"))
    mostrar_maximizado=nz_b(rstAux("mostrar_maximizado"))
end if
rstAux.close
conn.close
set conn = nothing
set command = nothing
set rstAux=nothing

''ricardo 23-4-2007 si el usuario tiene puesta ip, solamente podra acceder desde esa ip
function ComprobacionIp(usuario)
    ipacceso=request.ServerVariables(CLIENT_IP)
    si_ipacceso=1

    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    set acceso = Server.CreateObject("ADODB.Recordset")
    conn.open dsnilion
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    if mode="99" or hd_Act=1 then
        command.CommandText="SELECT ipacceso FROM indice with(nolock) where entrada=?"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,usuario)
    else
        command.CommandText="SELECT ipacceso FROM indice with(nolock) where entrada=? or id_usuario=?"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,request.form("nombre"))
        command.Parameters.Append command.CreateParameter("@usr",adVarChar,adParamInput,50,request.form("nombre"))
    end if
    set acceso= command.Execute
    if (acceso.EOF) then
        'El usuario no existe.'
        acceso.close
        conn.close
        set conn = nothing
        set command = nothing
        set acceso=nothing
%><script type="text/javascript" language="javascript">document.location = "<%=paginaAcceso%>?mode=error&d=<%=enc.EncodeForJavascript(parametroD)%>";</script><%
    else
        if acceso("ipacceso") & "">"" then
            if instr(1,acceso("ipacceso"),ipacceso,1)=0 then
                si_ipacceso=0
            end if
        end if
    end if

    acceso.close
    conn.close
    set conn = nothing
    set command = nothing
    set acceso=nothing

    ComprobacionIp=si_ipacceso

end function

' >>> MCA 09/12/04 : Gestión de las franjas horarias de acceso
function HoraCorrecta(usuario)
    session("tienda")=0
    session("usuario2")=""
    if (request.form("nombre") > "" and request.form("contrasenya") > "") or usuario > "" then

        if mode="99" or hd_Act=1 then
            query="SELECT case when accesodesdehora1 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesodesdehora1, 108)) else null end as accesodesdehora1, case when accesohastahora1 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesohastahora1, 108)) else null end as accesohastahora1,  case when accesodesdehora2 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesodesdehora2, 108)) else null end as accesodesdehora2, case when accesohastahora2 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesohastahora2, 108)) else null end as accesohastahora2 FROM indice with(nolock) where entrada=?"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open dsnilion
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText=query
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,usuario)
            set acceso = command.execute
        else
            query="SELECT case when accesodesdehora1 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesodesdehora1, 108)) else null end as accesodesdehora1, case when accesohastahora1 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesohastahora1, 108)) else null end as accesohastahora1,  case when accesodesdehora2 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesodesdehora2, 108)) else null end as accesodesdehora2, case when accesohastahora2 is not null then convert(datetime, '01/01/1900 ' + convert(varchar, accesohastahora2, 108)) else null end as accesohastahora2 FROM indice with(nolock) where entrada=? or id_usuario=?"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open dsnilion
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText=query
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,request.form("nombre"))
            command.Parameters.Append command.CreateParameter("@usr",adVarChar,adParamInput,50,request.form("nombre"))
            'response.End
            set acceso = command.execute
        end if

        if (acceso.EOF) then
            'El usuario no existe.'
            acceso.Close
            conn.close
            set conn = nothing
            set command = nothing
            set acceso = nothing
%><script type="text/javascript" language="javascript">document.location = "<%=paginaAcceso%>?mode=error&d=<%=enc.EncodeForJavascript(parametroD)%>";</script><%
        else
            horaactual = cdate("01/01/1900 " & Time())

            if isNull(acceso("accesodesdehora1")) and isNull(acceso("accesodesdehora2")) then franja=0 end if
            if acceso("accesodesdehora1")>0 and isNull(acceso("accesodesdehora2")) then franja=1 end if
            if isNull(acceso("accesodesdehora1")) and acceso("accesodesdehora2")>0 then franja=2 end if
            if acceso("accesodesdehora1")>0 and acceso("accesodesdehora2")>0 then franja=3 end if

            horacorrecta= true
            if franja>0 then
                if franja=1 then
                    if horaactual < acceso("accesodesdehora1") or horaactual > acceso("accesohastahora1") then
                        horacorrecta= false
                    end if
                elseif franja=2 then
                    if horaactual < acceso("accesodesdehora2") or horaactual > acceso("accesohastahora2") then
                        horacorrecta= false
                    end if
                elseif franja=3 then
                    if horaactual > acceso("accesodesdehora1") and horaactual < acceso("accesohastahora1") then
                        horacorrecta= true
                    elseif horaactual > acceso("accesodesdehora2") and horaactual < acceso("accesohastahora2") then
                        horacorrecta= true
                    else
                        horacorrecta= false
                    end	if
                end if
            end if
            acceso.Close
            conn.close
            set conn = nothing
            set command = nothing
            set acceso = nothing
        end if
    end if
end function

' <<< MCA 09/12/04 : Comprobación de las franjas horarias de acceso
' >>> MCA 20/12/04 : Añadir las comprobaciones de fecha de baja y franjas horarias de acceso al modo 99
function ComprobacionesAcceso(entrada)
    usu_bloq=""
    comprobacionesacceso= false

    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open dsnilion
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="GetIndiceData"
    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    command.Parameters.Append command.CreateParameter("@usr", adVarChar, adParamInput, 50, entrada)

    set rstAux = command.execute

    if not rstAux.eof then
        ''ricardo 23-10-2008 se controlara el campo movil, en la tabla clientes_users para saber si es un socio de covaldroper bloqueado
        usu_bloq=""
        rstAux.close
        set command = nothing

        set command =  Server.CreateObject("ADODB.Command")
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="GetClientesUsersData"
        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@usr", adVarChar, adParamInput, 50, entrada)

        set rstAux = command.execute

        si_ =""
        si_fbaja2=""
        if not rstAux.eof then
            si_fbaja=""
            si_fbaja2=""
            rstAux.close
        else
            ''ricardo 23-10-2008 se controlara el campo movil, en la tabla clientes_users para saber si es un socio de covaldroper bloqueado

            usu_bloq=""
            rstAux.close
            set command = nothing
            set command =  Server.CreateObject("ADODB.Command")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="SELECT movil FROM [CLIENTES_USERS] with(nolock)  where usuario =?"
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,entrada)

            set rstAux= command.Execute
            if not rstAux.eof then
                usu_bloq=rstAux("movil")
            end if
            rstAux.close
            set command = nothing

            set command =  Server.CreateObject("ADODB.Command")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="select * from accesos_int with(nolock) where usuario =? and fbaja is null"
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,entrada)

            set rstAux= command.Execute

            if not rstAux.eof then
                si_fbaja=""
                si_fbaja2=""
            else
                si_fbaja="esta_dado_de_baja"
                si_fbaja2=""
            end if
            rstAux.close
            set command = nothing
        end if

    else
        si_fbaja="esta_dado_de_baja"
        si_fbaja2=""
        rstAux.close
        set command = nothing
    end if
   
    %>
    <script language="javascript" type="text/javascript">
        function showMessage(message)
        {
            alert(message);
            return true;
        }
    </script>
    <%
    if si_fbaja & "">"" and si_fbaja2="" then
        if usu_bloq & ""="BAJA_COVALDROPER" then%>
            <script type="text/javascript" language="javascript">
                window.alert("<%=LitAccUsuCovalBloq%>");
            </script>
    <%else
        if Request.QueryString("d")>"" then%>
            <script type="text/javascript" language="javascript">
                //window.alert("<%=LITUSUARIOBAJAORG%>");
                showMessage("<%=LITUSUARIOBAJAORG%>");
            </script>
        <%else%>
            <script type="text/javascript" language="javascript">
                //window.alert("<%=LitUsuarioBaja%>");
                showMessage("<%=LitUsuarioBaja%>");
            </script>
    <%  end if
    end if
        elseif HoraCorrecta(entrada)=false then%>
            <script type="text/javascript" language="javascript">
                window.alert("<%= LitHAccesoIncorrecta %>");
            </script>
    <%elseif ComprobacionIp(entrada)=0 then%>
        <script type="text/javascript" language="javascript">
            window.alert("<%= LitAccesoIpIncorrecta %>");
        </script>
    <%else
        comprobacionesacceso= true
    end if
    'conn.close
    set rstAux = nothing
    set command = nothing
    set conn = nothing
end function

' <<< MCA 20/12/04 : Añadir las comprobaciones de fecha de baja y hora de acceso al modo 99
'***********************************************************************************************************
'El usuario ya está identificado y se procede a caragr su/s empresa/s

sub RedirigeUser()
    if session("EsAccesotienda")=1 then
        sessionUsuario= session("usuarioMKP")
    else
        sessionUsuario= session("usuario")
    end if

	'Verificar que el dominio de acceso no esta excluido'
	dim strDomain, isDomainExcluded				
	strDomain = Request.ServerVariables("SERVER_NAME")
	
	set conn = Server.CreateObject("ADODB.Connection")
	set command =  Server.CreateObject("ADODB.Command")
	conn.open dsnilion
	command.ActiveConnection = conn
	command.CommandText = "INST_GET_DOMAIN_EXCLUDED"
	command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	command.Parameters.Append command.CreateParameter("@ID", adBigInt, adParamInput, , Null)
	command.Parameters.Append command.CreateParameter("@DOMAIN_DESCRIPTION", adVarChar, adParamInput, 255, Null)
	command.Parameters.Append command.CreateParameter("@DOMAIN", adVarChar, adParamInput, 255, strDomain)

	set rst = command.execute

	isDomainExcluded = 0
	
	if not rst.eof then
		isDomainExcluded = 1
	end if

	rst.close
	conn.close
	
	set rst = nothing
	set command = nothing
	set conn = nothing
	
	'Obtenemos la instancia si el dominio no esta excluido'
	dim ninstance, instanceExits
	instanceExits = 0
										
	if(not CBool(isDomainExcluded)) then
		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")
		conn.open dsnilion
		command.ActiveConnection = conn
		command.CommandText = "INST_GET_INSTANCE_BY_FILTERS"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@NINSTANCE", adInteger, adParamInput, , Null)
		command.Parameters.Append command.CreateParameter("@NAME", adVarChar, adParamInput, 10, Null)
		command.Parameters.Append command.CreateParameter("@NPROJECT", adVarChar, adParamInput, 5, Null)
		command.Parameters.Append command.CreateParameter("@PRODUCT_BBDD", adVarChar, adParamInput, 100, Null)
		command.Parameters.Append command.CreateParameter("@PROJECT_BBDD", adVarChar, adParamInput, 100, Null)
		command.Parameters.Append command.CreateParameter("@URL_SITE", adVarChar, adParamInput, 255, strDomain)
		command.Parameters.Append command.CreateParameter("@PATH", adVarChar, adParamInput, 255, Null)
		command.Parameters.Append command.CreateParameter("@STATE", adInteger, adParamInput, , Null)
		command.Parameters.Append command.CreateParameter("@DSNPROY", adVarChar, adParamInput, 255, Null)
		command.Parameters.Append command.CreateParameter("@DSNPROD", adVarChar, adParamInput, 255, Null)
		command.Parameters.Append command.CreateParameter("@ACTIVE", adBoolean, adParamInput, , Null)
		command.Parameters.Append command.CreateParameter("@CREATE_USER", adVarChar, adParamInput, 80, Null)
		command.Parameters.Append command.CreateParameter("@CREATE_DATE", adDate, adParamInput, , Null)
		command.Parameters.Append command.CreateParameter("@USER_SFTP", adVarChar, adParamInput, 80, Null)
		command.Parameters.Append command.CreateParameter("@PASSWORD_SFTP", adVarChar, adParamInput, 80, Null)
		command.Parameters.Append command.CreateParameter("@DEPLOY", adBoolean, adParamInput, , Null)

		set rst = command.execute

		if not rst.eof then
			instanceExits = 1
			ninstance = rst("Ninstance")
		end if

		rst.close
		conn.close
		
		set rst = nothing
		set command = nothing
		set conn = nothing
	end if
	
	'Contamos empresas por instancia'
	dim client_instance
	cuenta_empresas = 0
	
	if(CBool(instanceExits)) then
		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")
		conn.open dsnilion
		command.ActiveConnection = conn
		command.CommandText = "INST_GET_INSTANCE_CLIENTS"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@NINSTANCE", adInteger, adParamInput, , ninstance)
		command.Parameters.Append command.CreateParameter("@USER_NAME", adVarChar, adParamInput, 15, session("usuario"))
		command.Parameters.Append command.CreateParameter("@PARAMETER_EXISTS", adBoolean, adParamInput, , false)
				
		set rst = command.execute

		while not rst.eof
			cuenta_empresas = cuenta_empresas + 1
			client_instance = rst("Nclient")
			rst.MoveNext
		wend
		
		rst.close
		conn.close
		
		set rst = nothing
		set command = nothing
		set conn = nothing
	end if
	
    'Capturamos las posibles empresas del cliente'
    set rs_empresas = Server.CreateObject("ADODB.Recordset")
    set rs_empresas2 = Server.CreateObject("ADODB.Recordset")
    set conn = Server.CreateObject("ADODB.Connection")
    conn.open dsnilion

	if(CBool(isDomainExcluded) or not CBool(instanceExits)) then 'Si no filtramos clientes por instancia
		set command =  Server.CreateObject("ADODB.Command")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="Select count(*) as num From Clientes_Users with(nolock) Where Usuario=? and cliente_int is null and proveedor_int is null and fbaja is null"
		command.CommandType = adCmdText
		command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,sessionUsuario)
		set rs_empresas = command.Execute
		cuenta_empresas = rs_empresas("num")
		rs_empresas.close
		set command = nothing
	end if

    set commandT =  Server.CreateObject("ADODB.Command")
    commandT.ActiveConnection =conn
    commandT.CommandTimeout = 0
    commandT.CommandText="Select count(*) as num From Clientes_Users with(nolock) Where fbaja is null and Usuario='" + sessionUsuario + "' and (cliente_int is not null or proveedor_int is not null)"
    commandT.CommandType = adCmdText
    commandT.Parameters.Append commandT.CreateParameter("@usuario",adVarChar,adParamInput,50,sessionUsuario)
    set rs_empresasT = commandT.Execute
    cuenta_tiendas=rs_empresasT("num")
    rs_empresasT.close
    set rs_empresasT =nothing
    set commandT = nothing

    set command =  Server.CreateObject("ADODB.Command")
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="Select count(*) as num From accesos_int with(nolock) Where fbaja is null and Usuario=?"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,sessionUsuario)
    set rs_empresas = command.Execute
    cuenta_asesoria=rs_empresas("num")
    rs_empresas.close
    set command = nothing

    ModoAdministrar=true
	
    'Si solo tiene una empresa se accede a sus opciones'
    if cuenta_empresas+cuenta_tiendas+cuenta_asesoria = 1 then

        if cuenta_empresas=1 or cuenta_tiendas=1 then
            set command =  Server.CreateObject("ADODB.Command")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="Select * From Clientes_Users with(nolock) Where fbaja is null and Usuario=?"
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,sessionUsuario)
            set rs_empresas = command.Execute

            if not rs_empresas.eof then
                ModoAdministrar=rs_empresas("administrar")
            else
                ModoAdministrar=true
            end if
        else
            set command =  Server.CreateObject("ADODB.Command")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="Select *  From accesos_int with(nolock) Where fbaja is null and Usuario=?"
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,sessionUsuario)
            set rs_empresas = command.Execute
        end if
		
        'Guardamos el numero del cliente'
		sesionNCliente = rs_empresas("ncliente")
			
		if(CBool(isDomainExcluded) or not CBool(instanceExits)) then 'Si no filtramos clientes por instancia
			if cuenta_empresas = 1 then
				session("ncliente") = rs_empresas("ncliente")
			else
				session("ncliente") = "     "
			end if
		else 'Si filtramos clientes por instancia
			if cuenta_empresas = 1 then
				sesionNCliente = client_instance
				session("ncliente") = client_instance
			else
				session("ncliente") = "     "
			end if
		end if		
        
        rs_empresas.close
        set command = nothing
		
		'Informar sessionStorage & cookies
		%><script type="text/javascript" languaje="javascript">
			if (sessionStorage.getItem("ncompany") === null)
			{
				sessionStorage.setItem("ncompany", "<%=sesionNCliente%>")
			}
			else
			{
				sessionStorage.ncompany = "<%=sesionNCliente%>"
			}
			
			var cookieName = 'companyCookie';
			var cookieValue = "<%=sesionNCliente%>"
			var tomorrow = new Date();
			tomorrow.setDate(tomorrow.getDate() + 1);
			var domain = "<%=Request.ServerVariables("SERVER_NAME")%>";

			document.cookie = cookieName +"=" + cookieValue + ";expires=" + tomorrow + ";domain=" + domain + ";path=/";
		</script><%

        'Buscamos la informacion del cliente'
        set rs_Clientes = Server.CreateObject("ADODB.Recordset")
        set command =  Server.CreateObject("ADODB.Command")
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="Select * From Clientes with(nolock) Where Ncliente=?"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,sesionNCliente)
        set rs_Clientes = command.Execute
        session("empresa") = rs_Clientes("Rsocial")
        'session("dsn_cliente") = rs_Clientes("DSN")
        'session("backendListados") = rs_Clientes("DSNLISTADOS") 'JAR - 18/10/07 DSN para listados.
        '201102

        if session("EsAccesotienda")=1 then
            session("dsn_nclienteMKP") = rs_Clientes("DSN")
            session("backendListados_MKP") = rs_Clientes("DSNLISTADOS") 'JAR - 18/10/07 DSN para listados.
        else
            session("dsn_cliente") = rs_Clientes("DSN")
            session("backendListados") = rs_Clientes("DSNLISTADOS") 'JAR - 18/10/07 DSN para listados.
            session("lenguaje") = rs_Clientes("lenguaje")
            session("caracteres") = rs_Clientes("caracteres")
            'dgb: nueva variable
            session("NetEstilo")= rs_Clientes("personalizacion")
            'dgb 08/11/2010:  asignamos el LOCALE
            session("locale")=rs_Clientes("LOCALE")
        end if

        'JCI 20/09/2010: Asignación al entorno del LCID de la empresa
        session.LCID=rs_Clientes("LCID")
        rs_Clientes.Close
        set command = nothing

        Auditar sesionNCliente,session("usuario"),session("usuario2"),"ENTRADA",Request.ServerVariables(CLIENT_IP),Request.ServerVariables("REMOTE_HOST"),Request.ServerVariables("HTTP_USER_AGENT"),DSNIlion
        entrar=0
        if sesionNCliente<>SISTEMA_GESTION then
            set rs_Clientes = Server.CreateObject("ADODB.Recordset")
            set command =  Server.CreateObject("ADODB.Command")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="select dsn from clientes with(nolock) where ncliente=?"
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,sesionNCliente)
            set rs_Clientes = command.Execute

            if not rs_Clientes.eof then
                ''MPC 29/10/2014 Se quita comprobación de la bbdd para que lo haga siempre, se deja como está la pantalla de sel_empresas.asp
                'if instr(rs_Clientes("dsn"),"ilion_")>0 then
                    entrar=1
                'else
                '    entrar=0
                'end if
            else
                entrar=0
            end if
            rs_Clientes.close
            set command = nothing
        end if

        if entrar=1 and sesionNCliente<>SISTEMA_GESTION then
            'Control de stock a nivel de empresa'

            set connCustomer = Server.CreateObject("ADODB.Connection")

            if session("EsAccesotienda")=1 then
                connCustomer.open session("dsn_nclienteMKP")
                set command =  Server.CreateObject("ADODB.Command")
                command.ActiveConnection =connCustomer
                command.CommandTimeout = 0
                command.CommandText="select control_stock from empresas with(nolock) where cif like ? + '%'"
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,sesionNCliente)
                set rs_Clientes = command.Execute
                'rs_Clientes.open "select control_stock from empresas with(nolock) where cif like '" & sesionNCliente & "%'",   session("dsn_nclienteMKP") , adOpenKeyset, adLockOptimistic
            else
                connCustomer.open session("dsn_cliente")
                set command =  Server.CreateObject("ADODB.Command")
                command.ActiveConnection =connCustomer
                command.CommandTimeout = 0
                command.CommandText="select control_stock from empresas with(nolock) where cif like ? + '%'"
                command.CommandType = adCmdText '
                command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,sesionNCliente)
                set rs_Clientes = command.Execute
                ''rs_Clientes.open "select control_stock from empresas with(nolock) where cif like '" & sesionNCliente & "%'", session("dsn_cliente"), adOpenKeyset, adLockOptimistic
            end if

            if not rs_clientes.eof then
                if rs_Clientes("control_stock")= true then
                    session("control_stock") = "activado"
                else
                    session("control_stock") = "desactivado"
                end if
            end if

            rs_Clientes.close
            connCustomer.close
            set command=nothing
        end if
    end if

    'ParametrosAcceso=d_lookup("parametros","param_usuario","ncliente='" & sesionNCliente & "' and usuario='" & sessionUsuario & "' and objeto='996'",DSNIlion)

    ParametrosAcceso=""
    set rs_parametros = Server.CreateObject("ADODB.Recordset")
    set command =  Server.CreateObject("ADODB.Command")
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="select parametros from param_usuario with(nolock) where ncliente=? and usuario=? and objeto='996'"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,sesionNCliente & "")
    command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,sessionUsuario)

    set rs_parametros = command.Execute

    if not rs_parametros.eof then
        ParametrosAcceso = rs_parametros("parametros")
    end if

    rs_parametros.close
    set rs_parametros = nothing
    set command = nothing

    param=""
    if instr(ParametrosAcceso,"?tyg=1")>0 then
        param="tyg=1"
        'EsComercial="ec=1"
        param=param & EsComercial
        '20110323:Just to users with this param would have session to mkp and ilion.
        if session("EsAccesotienda")=1 then
            session("dsn_cliente")=session("dsn_nclienteMKP")
            session("backendListados")=session("backendListados_MKP")
            session("usuario") =session("usuarioMKP")
        end if
    end if

    conn.close
    set command = nothing
    set conn = nothing
    
    if cuenta_empresas+cuenta_tiendas+cuenta_asesoria = 1 then

        
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open dsnilion
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="GetShowSelApp"
        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@ncompany", adVarChar, adParamInput, 5, sesionNCliente)

        set rs_distrib = command.execute

        paso=false
        pageSelApp=""
        if not rs_distrib.eof then
            select case rs_distrib("SHOWSELAPP")
                case 0
                    paso = false
                case 1
                    paso = true
                    pageSelApp=rs_distrib("PAGESELAPP")
            end select
        end if

        rs_distrib.close
        conn.close
        set rs_distrib = nothing
        set command = nothing
        set conn = nothing
    end if
    ''response.write("los datos son-" & ModoAdministrar & "-" & cuenta_empresas & "-" & cuenta_tiendas & "-" & cuenta_asesoria & "-" & paso & "-<br>")
    ''response.end
    'response.write("Session :"  &session("dsn_cliente"))
        'response.End()

        ' response.write("<br>paso:"+cstr(paso))
        ' response.Write("<br>pageSelApp:"+pageSelApp)
        ' response.end

    if paso then
        'map 16/05/2013 - Abrir página de aplicaciones en función de la página indicada en el distribuidor
        if pageSelApp&""<>"" then
            %><script type="text/javascript" language="javascript">RedirigirEmpresas('app', '<%=pageSelApp %>');</script><%
        else
            %><script type="text/javascript" language="javascript">RedirigirEmpresas('app', 'sel_app.asp');</script><%
        end if
    else
        if cuenta_empresas=1 and cuenta_tiendas=0 and cuenta_asesoria=0 then
            if ModoAdministrar then
                %><script type="text/javascript" language="javascript">RedirigirEmpresas('unaempresa', '<%=param%>');</script><%
            else
                if instr(ParametrosAcceso,"?tyg=1")>0 then
                    %><script type="text/javascript" language="javascript">RedirigirEmpresas('tiendasTarjMaestra', '<%=sesionNCliente%>');</script><%
                else
                    %><script type="text/javascript" language="javascript">RedirigirEmpresas('tiendas', '<%=param%>');</script><%
                end if
            end if
        elseif cuenta_empresas=0 and cuenta_tiendas>=1 and cuenta_asesoria=0  then
            %><script type="text/javascript" language="javascript">RedirigirEmpresas('tiendas', '<%=param%>');</script><%
        elseif cuenta_empresas=0 and cuenta_tiendas=0 and cuenta_asesoria>=1 then
            %><script type="text/javascript" language="javascript">RedirigirEmpresas('asesoria', '<%=param%>');</script><%
        else
            %><script type="text/javascript" language="javascript">RedirigirEmpresas('empresas', '<%=param%>');</script>
        <%end if
    end if
    set rs_empresas = Nothing
    set rs_empresas2 = Nothing
    set rs_Clientes = Nothing
end sub

'***********************************************************************************************************
function Acceder()
    session("tienda")=0
    session("usuario2")=""
    '201102:Nueva variable para identificar con quien estoy accediendo.
    session("EsAccesotienda")=0

    if strCodigo="-1" then
        'el usuario se identifica por el metodo tradicional
        nombreAc=request.form("nombre")
        passAc=request.form("contrasenya")
        hd_Act=0
    else
        'el usuario se identifico mediante huella dactilar
        nombreAc=strCodigo
        passAc="OK"
        hd_Act=1
    end if

    if nombreAc> "" and passAc> "" then
        set acceso = Server.CreateObject("ADODB.Recordset")
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 =  Server.CreateObject("ADODB.Command")
        conn2.open dsnilion
        command2.ActiveConnection =conn2
        command2.CommandTimeout = 0
        command2.CommandText="CheckLoginOrID"
        command2.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        command2.Parameters.Append command2.CreateParameter("@login", adVarChar, adParamInput, 150, nombreAc)
        set acceso = command2.Execute

        if (acceso.EOF) then
            'El usuario no existe.'
            acceso.Close
            %><script type="text/javascript" language="javascript">document.location = "<%=paginaAcceso%>?mode=error&d=<%=enc.EncodeForJavascript(parametroD)%>"</script><%
        elseif ComprobacionesAcceso(acceso("entrada"))=true then
            ' <<< MCA 20/12/04 Comprobar si la cuenta está pendiente de ser activada.
            nombreAc=acceso("entrada")
            if not isnull(acceso("nactivacion")) then
                'Activar la cuenta.
                acceso.Close
                Acceder="ACTIVAR"
            else

                ''Ricardo 19-05-2015 si el usuario existe, se auditara si tiene algun bloqueo de acceso
                cuantos_intentos=0
                maximos_intentos=0
                usuario_bloqueado=0
                set rs_compInt = Server.CreateObject("ADODB.Recordset")
                set commandcompInt =  Server.CreateObject("ADODB.Command")
                commandcompInt.ActiveConnection =conn2
                commandcompInt.CommandTimeout = 0
                commandcompInt.CommandText="GetUserAccessAttempts"
                commandcompInt.CommandType = adCmdStoredProc
                commandcompInt.Parameters.Append commandcompInt.CreateParameter("@nuser",adVarChar,adParamInput,50,acceso("entrada"))
                set rs_compInt = commandcompInt.Execute
                if not rs_compInt.eof then
                    cuantos_intentos=rs_compInt("NATTEMPT")
                    maximos_intentos=rs_compInt("MAXRETRY")
                end if
                rs_compInt.close
                set rs_compInt=nothing
                set commandcompInt = Nothing
                set rs_compInt = Server.CreateObject("ADODB.Recordset")
                set commandcompInt =  Server.CreateObject("ADODB.Command")
                commandcompInt.ActiveConnection =conn2
                commandcompInt.CommandTimeout = 0
                commandcompInt.CommandText="CheckUserIsBlocked"
                commandcompInt.CommandType = adCmdStoredProc
                commandcompInt.Parameters.Append commandcompInt.CreateParameter("@nuser",adVarChar,adParamInput,50,acceso("entrada"))
                set rs_compInt = commandcompInt.Execute
                if not rs_compInt.eof then
                    usuario_bloqueado=rs_compInt("blocked")
                end if
                rs_compInt.close
                set rs_compInt=nothing
                set commandcompInt = Nothing
''response.write("los intentos son-" & cuantos_intentos & "-" & maximos_intentos & "-" & usuario_bloqueado & "-<br>")
                if usuario_bloqueado=0 then
                    'response.Write("<br>VER CONTRASEÑAS:")
                    'response.Write("<br>Confirma-->"+acceso("confirma"))
                    'response.Write("<br>Confirma limpiar-->"+limpiar(acceso("confirma"),acceso("version")))
                    'response.Write("<br>contraseña introducida -->"+passAc)
                    condicion_acceso1=0
                    if limpiar(acceso("confirma"),acceso("version"))=passAc then
                        condicion_acceso1=1
                    end if
                    condicion_acceso2=0
                    if verifica_clave(request.form("ncasilla"),request.form("vcasilla"),acceso("indice"),acceso("version"),acceso("usatarjeta")) then
                        condicion_acceso2=1
                    end if
    ''response.write("las condiciones 1 son-" & acceso("entrada") & "-" & condicion_acceso1 & "-" & condicion_acceso2 & "-<br>")
                    numero_intentos_a_poner=0
                    if condicion_acceso1=0 or condicion_acceso2=0 then
                        numero_intentos_a_poner=cuantos_intentos+1
                    end if
    ''response.write("la contraseña o la casilla es incorrecta-" & acceso("entrada") & "-<br>")
                    ''Ricardo 19-05-2015 se contralara si el usuario se equivoca de contraseña
                    if (cuantos_intentos<maximos_intentos) then
                        set rs_compInt = Server.CreateObject("ADODB.Recordset")
                        set commandcompInt =  Server.CreateObject("ADODB.Command")
                        commandcompInt.ActiveConnection =conn2
                        commandcompInt.CommandTimeout = 0
                        commandcompInt.CommandText="UpdateUserAccessAttempts"
                        commandcompInt.CommandType = adCmdStoredProc
                        commandcompInt.Parameters.Append commandcompInt.CreateParameter("@nuser",adVarChar,adParamInput,150,acceso("entrada"))
                        commandcompInt.Parameters.Append commandcompInt.CreateParameter("@nattempts",adSmallint,adParamInput,,numero_intentos_a_poner)
                        set rs_compInt = commandcompInt.Execute
                        set rs_compInt=nothing
                        set commandcompInt = nothing
                        cuantos_intentos=numero_intentos_a_poner
                        if cuantos_intentos>=maximos_intentos then
                            set rs_compInt = Server.CreateObject("ADODB.Recordset")
                            set commandcompInt =  Server.CreateObject("ADODB.Command")
                            commandcompInt.ActiveConnection =conn2
                            commandcompInt.CommandTimeout = 0
                            commandcompInt.CommandText="BlockUser"
                            commandcompInt.CommandType = adCmdStoredProc
                            commandcompInt.Parameters.Append commandcompInt.CreateParameter("@ncompany",adVarChar,adParamInput,5,"")
	                        commandcompInt.Parameters.Append commandcompInt.CreateParameter("@nuser",adVarChar,adParamInput,50,acceso("entrada"))
                            commandcompInt.Parameters.Append commandcompInt.CreateParameter("@nuserBlq",adVarChar,adParamInput,50,acceso("entrada"))
                            commandcompInt.Parameters.Append commandcompInt.CreateParameter("@ip",adVarChar,adParamInput,75,Request.ServerVariables(CLIENT_IP))
                            commandcompInt.Parameters.Append commandcompInt.CreateParameter("@host",adVarChar,adParamInput,75,Request.ServerVariables("REMOTE_HOST"))
                            set rs_compInt = commandcompInt.Execute
                            set rs_compInt=nothing
                            set commandcompInt = nothing
                        end if
                    end if
    ''response.write("las condiciones 2 son-" & acceso("entrada") & "-" & condicion_acceso1 & "-" & condicion_acceso2 & "-" & cuantos_intentos & "-" & maximos_intentos & "-" & hd_Act & "-<br>")
    ''if acceso("entrada")="12341234" then
    ''    response.end
    ''end if
                    if (cuantos_intentos<maximos_intentos) then
                        if (hd_Act=1 or (condicion_acceso1=1 and condicion_acceso2=1)) then
                            
                            if accesoPagina(session.sessionid,nombreAc)=1 then
                            'if acceso("activo")=0 then
                                'Asignar un valor a la variable de ámbito "usuario" de la sesion.'
                                'session("usuario") = acceso("entrada")
                                'ES necesario saber si es un usuario de mkp o no
                                '201102'''''''''''''''''''

                
                                'Set user associated to the idnet in Redis
                                dim redis
                                set redis = Server.CreateObject("EverilionNetRedis.EverilionNetRedis")
                                dim redisResponse
                                redisResponse = redis.SetValueRedisASP(idnet, nombreAc, true)
                                redisResponse = redis.SetValueRedisASP(nombreAc, session("locale_user"), true)
                                'End set user Redis

                                set rs_empresasMKP = Server.CreateObject("ADODB.Recordset")
                                set commandNew =  Server.CreateObject("ADODB.Command")
                                commandNew.ActiveConnection =conn2
                                commandNew.CommandTimeout = 0
                                commandNew.CommandText="Select count(*) as num From Clientes_Users with(nolock) Where Usuario=? and cliente_int is null and proveedor_int is null"
                                commandNew.CommandType = adCmdText
                                commandNew.Parameters.Append commandNew.CreateParameter("@usuario",adVarChar,adParamInput,50,acceso("entrada"))
                                set rs_empresasMKP = commandNew.Execute
                                cuenta_empresas=rs_empresasMKP("num")
                                rs_empresasMKP.close
                                set commandNew = Nothing

                                set commandNew =  Server.CreateObject("ADODB.Command")
                                commandNew.ActiveConnection =conn2
                                commandNew.CommandTimeout = 0
                                commandNew.CommandText="Select count(*) as num From Clientes_Users with(nolock) Where fbaja is null and Usuario=? and (cliente_int is not null or proveedor_int is not null)"
                                commandNew.CommandType = adCmdText
                                commandNew.Parameters.Append commandNew.CreateParameter("@usuario",adVarChar,adParamInput,50,acceso("entrada"))
                                set rs_empresasMKP = commandNew.Execute
                                cuenta_tiendas=rs_empresasMKP("num")
                                rs_empresasMKP.close
                                set commandNew = Nothing

                                set commandNew =  Server.CreateObject("ADODB.Command")
                                commandNew.ActiveConnection =conn2
                                commandNew.CommandTimeout = 0
                                commandNew.CommandText="Select count(*) as num From accesos_int with(nolock) Where fbaja is null and Usuario=?"
                                commandNew.CommandType = adCmdText
                                commandNew.Parameters.Append commandNew.CreateParameter("@usuario",adVarChar,adParamInput,50,acceso("entrada"))
                                set rs_empresasMKP = commandNew.Execute
                                cuenta_asesoria=rs_empresasMKP("num")
                                rs_empresasMKP.close
                                set commandNew = Nothing
                                set rs_empresasMKP= nothing


                                ''ricardo 17-03-2015 si es un usuario de nettit, que se pase a su pantalla de login
                                set commandNew =  Server.CreateObject("ADODB.Command")
                                commandNew.ActiveConnection =conn2
                                commandNew.CommandTimeout = 0
                                commandNew.CommandText="Select top 1 ncliente From ilion_admin..Clientes_Users with(nolock) Where fbaja is null and Usuario=? order by ncliente"
                                commandNew.CommandType = adCmdText
                                commandNew.Parameters.Append commandNew.CreateParameter("@usuario",adVarChar,adParamInput,50,acceso("entrada"))
                                set rs_empresasMKP = commandNew.Execute
                                paso2=false

                                  
                                if not rs_empresasMKP.eof then
                                    if (ModuloContratado(nclienteMod,ModNettfi)=1) then
                                        paso2=true
                                        pageSelApp=GenerarURLbt & "/" & CarpetaProduccionX4 & "/Custom/RCD/BackOffice/LoginBackOffice.aspx?ncliente=" & sesionNCliente & "&usuario=" & sesionUsuarioMKP & "&mismoncliente=" & mismoncliente & "&mismoproveedor=" & mismoproveedor & "&empresas=varias&mismousuario=" & mismousuario
                                    end if
                                end if

                                rs_empresasMKP.close
                                set commandNew = Nothing
                                if paso2 then
		                            %>
                                    <script type="text/javascript" languaje="javascript">
                                        document.location = "<%=pageSelApp%>";
		                            </script>
                                    <%
                                    response.end
                                end if
                                ''fin ricardo 17-03-2015

                                'ACCESO TIENDA

                                if (cuenta_empresas=0 and cuenta_tiendas>=1 and cuenta_asesoria=0) then
                                    session("usuarioMKP") = acceso("entrada")
                                    session("EsAccesotienda")=1
                                else
                                    session("usuario") = acceso("entrada")
                                    session("EsAccesotienda")=0
                                    session("ncliente")=""
                                end if
                                'fin:201102''''''''''''''''''

                                acceso.Close
                                RedirigeUser
                            else
                                'El usuario aun está en el sistema.'
                                acceso.close
                            end if
                        else
                            acceso.close%>
                            <script type="text/javascript" language="javascript">document.location = "<%=paginaAcceso%>?mode=error&d=<%=enc.EncodeForJavascript(parametroD)%>"</script>
                        <%end if
                    else ''viene del if de intentos<maxintentos
                        acceso.close
                        %>
                            <script type="text/javascript" language="javascript">
                                window.alert("Este usuario no tiene acceso al sistema, póngase en contacto con su administrador para más información");
                                document.location = "<%=paginaAcceso%>?d=<%=enc.EncodeForJavascript(parametroD)%>";
                            </script>
                        <%
                    end if
                else ''viene el if de bloqueado
                    acceso.close
                    %>
                        <script type="text/javascript" language="javascript">
                            window.alert("Este usuario no tiene acceso al sistema, póngase en contacto con su administrador para más información");
                            document.location = "<%=paginaAcceso%>?d=<%=enc.EncodeForJavascript(parametroD)%>";
                        </script>
                    <%
                end if
            end if
        else
            acceso.close%>
            <script type="text/javascript" language="javascript">document.location = "<%=paginaAcceso%>?d=<%=enc.EncodeForJavascript(parametroD)%>";</script>
        <%end if
        set conn2 = nothing
        set command2 =  nothing

    else
        'No se ha especificado el nombre y/o la contaseña.'
    %>
        <script type="text/javascript" language="javascript">document.location = "<%=paginaAcceso%>?d=<%=enc.EncodeForJavascript(parametroD)%>";</script>
    <%end if
end function

function verifica_clave(casilla, valor, tablaClaves, version,usatarjeta)
    if nz_b(usatarjeta)=-1 then
        datosTabla=d_lookup("datos","versiones","version=" & version,DSNIlion)

        datoCasilla=right(left(tablaClaves,casilla),1)
        'posTabla=(fix(valor/10)*9)+(valor mod 10)+1
        ''ricardo 3-6-2010 como se puede dejar vacia la casilla, en ese caso se pondra un valor de 99999
        if valor & ""="" then valor="99999"
        posTabla=valor+1
        vCasilla=right(left(datosTabla,posTabla),1)

        if datoCasilla=vCasilla then
            verifica_clave=true
        else
            verifica_clave=false
        end if
    else
        verifica_clave=true
    end if
end function%>

<body class="login animated fadeIn">

    <%AbrirModal "fr_RecoveryPass",paginaModal,0,0,"no","si","noresize","S","cerrar"


mode=request.querystring("mode")

' José Miguel Martínez --> JMMM --> 04/12/2009
' Nuevo modo acceso a APLICATECA
if request.querystring("app")="ERPAPL" or mode="fin" then
    if mode<>"fin" then mode="sso_aplicateca"
end if
' FIN JMMM

' MPC 27/02/2012
' New access mode ONO Kynesis
if request.querystring("app")="OANPOL" or mode="fin" then
    if mode<>"fin" then mode="sso_kynesis"
end if
' FIN JMMM

' MPC 27/02/2012
' Start / ACCESS THROUGH ONO-KYNESIS

if mode = "sso_kynesis" then
    sessionId = Request("sessionid")
    userName = Request("username")

    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open dsnilion
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="SSO_Kynesis"
    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    command.Parameters.Append command.CreateParameter("@userId", adVarChar, adParamInput, 150, userName)
    command.Parameters.Append command.CreateParameter("@sessionId", adVarChar, adParamInput, 256, sessionId)
    command.Parameters.Append command.CreateParameter("@mode", adVarChar, adParamInput, 3, "acc")

    set rst = command.Execute
    if not rst.eof then
        if cint(rst("response")) = -1 then
            rst.close

            set commandIns =  Server.CreateObject("ADODB.Command")
            commandIns.ActiveConnection =conn
            commandIns.CommandTimeout = 0
            commandIns.CommandText="update indice with(updlock) set claveop=null where id_usuario = ?"
            commandIns.CommandType = adCmdText
            commandIns.Parameters.Append commandIns.CreateParameter("@usuario",adVarChar,adParamInput,150,userName)
            set rst = commandIns.Execute
            set rst = nothing
            set commandIns = nothing
            set conn = nothing%>
    <script type="text/javascript" language="javascript">
        alert("<%=LitMsgErrorAccONO%>");
        window.open('', '_self');
        window.close();
        window.close();
    </script>
    <%elseif cint(rst("response")) = 0 then
            rst.close
            set accesoONO = Server.CreateObject("ADODB.Recordset")
            set commandIns =  Server.CreateObject("ADODB.Command")
            commandIns.ActiveConnection =conn
            commandIns.CommandTimeout = 0
            commandIns.CommandText="select entrada from indice with(nolock) where claveop = ? and id_usuario = ?"
            commandIns.CommandType = adCmdText
            commandIns.Parameters.Append commandIns.CreateParameter("@claveop",adVarChar,adParamInput,50,sessionId)
            commandIns.Parameters.Append commandIns.CreateParameter("@usuario",adVarChar,adParamInput,150,userName)
            set accesoONO = commandIns.Execute

            'accesoONO.Open "select entrada from indice with(nolock) where claveop = '" & sessionId & "' and id_usuario = '" & userName & "'", DsnIlion, adOpenKeyset, adLockOptimistic

            if (accesoONO.EOF) then
                accesoONO.Close
                set commandIns = nothing
                set commandIns =  Server.CreateObject("ADODB.Command")
                commandIns.ActiveConnection =conn
                commandIns.CommandTimeout = 0
                commandIns.CommandText="update indice with(updlock) set claveop=null where id_usuario = ?"
                commandIns.CommandType = adCmdText
                commandIns.Parameters.Append commandIns.CreateParameter("@usuario",adVarChar,adParamInput,150,userName)
                 set accesoONO = commandIns.Execute
                'accesoONO.open "update indice with(updlock) set claveop=null where id_usuario = '" & userName & "'", DsnIlion, adOpenKeyset, adLockOptimistic
                set accesoONO=nothing
                set rst = nothing
                set commandIns = nothing
                set conn = nothing%>
            <script type="text/javascript" language="javascript">
                alert("<%=LitMsgErrorAccONO%>");
                window.open('', '_self');
                window.close();
                window.close();
            </script>
            <%else
                session("usuario") = accesoONO("entrada")
                session("OANPOL") = "yes"
                accesoONO.Close
                set commandIns =  Server.CreateObject("ADODB.Command")
                commandIns.ActiveConnection =conn
                commandIns.CommandTimeout = 0
                commandIns.CommandText="update indice with(updlock) set claveop=null where id_usuario = ?"
                commandIns.CommandType = adCmdText
                commandIns.Parameters.Append commandIns.CreateParameter("@usuario",adVarChar,adParamInput,150,userName)
                 set accesoONO = commandIns.Execute
                'accesoONO.open "update indice with(updlock) set claveop=null where id_usuario = '" & userName & "'", DsnIlion, adOpenKeyset, adLockOptimistic
                set accesoONO=nothing
                set rst = nothing
                set commandIns = nothing
                set conn = nothing
                RedirigeUser
            end if
        end if
    end if
end if

'Gestión de logotipos de cada empresa'
logoDistribuidor="false"
''parametroD=request.querystring("D")
if parametroD>"" and mode<>"demo" and mode<>"fin" then

    if len(parametroD)>5 then

        'nDistribuidor=d_lookup("ndistribuidor", "distribuidores", "id_partner='" & parametroD & "'", dsnilion)
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open dsnilion
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="GetDistributorData"
        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@type", adVarChar, adParamInput, 2, "01")
        command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, "")
        command.Parameters.Append command.CreateParameter("@id", adVarChar, adParamInput, 36, parametroD)

        set rst = command.execute

        if not rst.eof then
            nDistribuidor = rst("ndistribuidor")
        end if

        rst.close
        conn.close
        set rst = nothing
        set command = nothing
        set conn = nothing

    else
        nDistribuidor=completar(parametroD,5,"0")
    end if

end if
'FIN Gestión de logotipo de cada empresa'
if mode<>"demo" and mode<>"fin" and mode<>"error" and mode<>"comprueba" then
    if logoDistribuidor="false" then
       pos=0
    else
       pos=80
    end if
    randomize
    casilla = Int((50 - 1 + 1) * Rnd + 1)
    'casilla=35
    casilla = String(2 - Len(CStr(casilla)), "0") & casilla

    randomize
    id_imagen = Int((3 - 1 + 1) * Rnd + 1)
    if id_imagen=1 then
        img="under_footer1"
    elseif id_imagen=2 then
        img="under_footer2"
    elseif id_imagen=3 then
        img="under_footer3"
    end if

    if parametroD>"" and mode<>"demo" and mode<>"fin" then
        otro_logo="SI"
    end if%>
    <form name="entrada" action="" method="post" onkeypress="javascript:keypressed();">
        <%
        if parametroD>"" and mode<>"demo" and mode<>"fin" then%>
        <div class="bck-wrapper">
            <img class="bck" id="bck">
        </div>
        <div style="overflow: hidden;">
            <table width="100%">
                <tr>
                    <!--<td style="width: 50%">-->
                    <td style="width: 100%">
                        <div id="logos_body_everilion" style="width: 100%" class="logo-login animated slideInDown"></div>
                    </td>
                </tr>
            </table>
        </div>
        <%else%>
            <div class="bck-wrapper">
                <img class="bck" id="bck">
            </div>
            <div id="logos_body"></div>
        <%end if%>
        <table id="form" cellpadding="0" cellspacing="0" class="login-box animated zoomIn">
            <tr>
                <td>
                    <label><%=LitUsuario%></label></td></tr>
            <tr>
                <td colspan="2">
                     <% if adminpwd<>"True" then %>
                        <input type="text" name="nombre" placeholder="<%=LitUsuario%>" class="width100"/></td>
                    <% else %>
                        <input type="text" name="nombre" placeholder="<%=LitUsuario%>" class="width100" onchange="bloqPassword_orange()" /></td>
                    <% end if %>
            </tr>
            <tr>
                <td><label><%=LitContrasenya%> </label></td>
            </tr>
            <tr>
                <td colspan="2">
                    <%
                    if adminpwd<>"True" then %>
                        <input type="password" name="contrasenya"  placeholder="<%=LitContrasenya%>" class="width100" autocomplete="off" /></td>
                    <% else %>
                        <input type="password" name="contrasenya"  placeholder="<%=LitContrasenya%>" class="width100" autocomplete="off" disabled="disabled"/></td>

                    <% end if %>
                    
            </tr>
            <%'Si es ORANGE no muestra la casilla
        'if ndist<>"01000" then %>
            <tr id="casilla">
                <td><label><%=LitCasilla%>&nbsp;<%=casilla%> </label></td></tr>
            <tr>
                <td colspan="2">
                    <input type="text" maxlength="2" name="vcasilla" placeholder="Casilla X" class="width100"/><input type="hidden" value="<%=casilla%>" name="ncasilla" /></td>
            </tr>
            <%'else%>
            <!--<tr id="casilla">
            <td><input type="hidden" value=<%=casilla%> name="vcasilla" /></td>
        </tr>-->

            <%'end if %>
           
                <%'if ndist = "00025" OR ndist="01000" then%>
                <%'MAP 16/05/2013 - Nueva clase recoveryPass para mostrar u ocultar el link de "recordar contraseña" %>
                <% if ndist = "00025" then%>
                <tr>
                    <td class="data recoveryPass">
                        <a href="javascript:RecoveryPass();">Recordar contraseña</a>
                    </td>
                   <% else%>
                    <%if adminpwd="True" then %>

                     <td class="data recoveryPass">
                        <a href="javascript:RecoveryPass();">Olvidé mi contraseña</a>
                    </td>
                    <%end if %>

                </tr>
                <%end if %>
            


             <tr id="buttonenter">
                <!--MAP 16/05/2013 - COMENTADO PARA UTILIZAR SIEMPRE EL BOTÓN ESTÁNDAR DE ACCESO - no vale para GRANT-->
                <%if ndist = "00025" then%>
                    <td colspan="2">
                        <input type="button" name="enter" value="Entrar" onmouseover="javascript:window.status='Entrar';return true;" onmouseout="javascript:window.status='';" onclick="javascript:ValidarDatos();" />
                    </td>
                <%else%>
                <td colspan="2">
                    <a onmouseover="javascript:window.status='Entrar';return true;" onmouseout="javascript:window.status='';" href="javascript:ValidarDatos();" name="entrar" style="cursor: pointer;">
                        <%if ndLitOk = "1" then %>
                            <div id="button_enter" class="ui-button wide"> <%=LitEntrar %></div>
                        <% else %>
                            <div id="button_enter" class="ui-button wide"> </div>
                        <% end if %>
                    </a>
                </td>
                <%end if%>
            </tr>
            <tr><td></td></tr>

             
            <%if adminpwd="True" then %>
                <tr>
                    <td></td>
                    <td class="data registerPass">
                        <div class="uppercase"><b><%=newuserliteral%></b></div>
                        <a href="javascript:newOrangePass();" id="button_enter"></a>
                    </td>
                </tr>
            <%end if %>
        </table>
        <div class="bck-text animated slideInUp">
            <img id="text">
        </div>
        <script type="text/javascript">
            function getRandomInt(min, max) {
                return Math.floor(Math.random() * (max - min)) + min;
            }			
            var imagesArray=["images/img_bck1.jpg", "images/img_bck2.jpg", "images/img_bck4.jpg", "images/img_bck5.jpg", "images/img_bck6.jpg", "images/img_bck7.jpg", "images/img_bck8.jpg"]
            var txtArray=["images/txt_black.png", "images/txt_black.png", "images/txt_white.png", "images/txt_black.png", "images/txt_white.png", "images/txt_white.png","images/txt_black.png"]
            var result=getRandomInt(0, 7)
            document.getElementById("bck").src = "/lib/estilos/<%=folder%>/"+imagesArray[result];
            document.getElementById("text").src ="/lib/estilos/<%=folder%>/"+ txtArray[result];
			
        </script>
        <%if ndist = "00025" then%>
            <div id="text">
                <div id="content"></div>
                <br />
                <div id="content2"></div>
                <br />
                <br />
                <table>
                    <tr>
                        <td>
                            <div id="register">Si todavía no estás registrado</div>
                            <br />
                            <input type="button" style="font-size: 15px !important" name="register" value="Registrate ahora ¡GRATIS!" onclick="javascript:Register();" />
                        </td>
                    </tr>
                </table>
            </div>
            <div id="capa_pie">
                <p style="text-align: center; font-size: 11px; padding-bottom: 4px;">
                    &copy; 2012 Tabarca Technologies. Todos los derechos reservados.
                <a href="http://grantbiomed.com/">CONDICIONES DE USO</a> I <a href="http://grantbiomed.com/">POL&Iacute;TICA DE PRIVACIDAD</a>
                </p>
                <div id="foot-2">
                    <div id="foot-a">
                        <h3>M&aacute;s</h3>
                        <ul>
                            <li></li>
                            <a href="http://grantbiomed.com/">Inicio</a></li>
                        <li><a href="http://grantbiomed.com/#bc7">&iquest;Qu&eacute; es grant?</a></li>
                            <li><a href="http://grantbiomed.com/blog">Blog</a></li>
                            <li><a href="http://grantbiomed.com/#bc2">Planes y Precios</a></li>
                            <li><a href="http://grantbiomed.com/modulos">M&oacute;dulos</a></li>
                            <li><a href="http://grantbiomed.com/quienes-somos/">&iquest;Qui&eacute;nes somos?</a></li>
                            <li>Entrar</li>
                        </ul>
                    </div>
                    <div id="foot-b">
                        <h3>Descargas</h3>
                        <ul>
                            <li><a href="http://grantbiomed.com/">Iphone App (pr&oacute;ximamente)</a></li>
                            <li><a href="http://grantbiomed.com/">Android Market (pr&oacute;ximamente)</a></li>
                        </ul>
                    </div>
                    <div id="foot-d">
                        <h3>Enlaces</h3>
                        <ul>
                            <li><a href="http://grantbiomed.com/">Gal&eacute;nica</a></li>
                            <li><a href="http://tabarcatechnologies.com/">Tabarca</a></li>
                            <li><a href="http://cidiplus.com/">CIDi+</a></li>
                        </ul>
                    </div>
                    <div id="foot-c">
                        <h3>Contacto</h3>
                        <ul>
                            <li><a href="mailto:grant@tabarcatechnologies.com">Email</a></li>
                        </ul>
                    </div>
                </div>
                <div style="margin-top: 10px; vertical-align: middle; text-align: center;">
                    <img style="float: left; border: 0px;" src="/lib/estilos/<%=folder%>/images/zonasegura.png" alt="logos" width="74" height="44" />
                    <img usemap="#Map" src="/lib/estilos/<%=folder%>/images/PIE-logos.jpg" border="0" alt="logos" />
                </div>
            </div>
        <%end if%>
    </form>
    <%end if

if mode="99" then
    set rst = Server.CreateObject("ADODB.Recordset")
    u=request.querystring("u")

    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open dsnilion
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="select activo,sesionid,version_csc from indice with(nolock) where entrada=?"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@entrada",adVarChar,adParamInput,50,u)

    set rst = command.execute

    'rst.open "select activo,sesionid,version_csc from indice with(nolock) where entrada='" & u & "'",DSNIlion
    if not rst.eof then
        if rst("activo")<>3 or rst("sesionid")<>0 or rst("version_csc")&""="" then
            rst.close
            set command = nothing
            mode="error"
        elseif ComprobacionesAcceso(u)=true then ' >>> MCA 20/12/04
            rst.close
            set command = nothing
            set conn = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open dsnilion
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="update indice with(updlock) set sesionid=? ,netsessionid=? where entrada=?"
            command.CommandType = adCmdText
            'dgb session.sessionid change to sessionidNet
            command.Parameters.Append command.CreateParameter("@sesionid",adDouble,adParamInput,14,session.sessionid)
            command.Parameters.Append command.CreateParameter("@netsesionid",adVarChar,adParamInput,40,idnet)
            command.Parameters.Append command.CreateParameter("@entrada",adVarChar,adParamInput,50,u)
            set rst = command.Execute
            'rst.open "update indice with(updlock) set sesionid=" & session.sessionid & " where entrada='" & u & "'",DSNIlion,1,3

            'set command = nothing
            'set command =  Server.CreateObject("ADODB.Command")
            'command.ActiveConnection =conn
            'command.CommandTimeout = 0
            'command.CommandText="select activo,sesionid from indice where entrada=?"
            'command.CommandType = adCmdText
            'command.Parameters.Append command.CreateParameter("@entrada",adVarChar,adParamInput,50,u)
            'set rst = command.Execute
            rst.open "select activo,sesionid,netsessionid from indice where entrada='" & u & "'",DSNIlion,1,3
            ahora=Now
            while rst("activo")<>1 and DateDiff("s", ahora, Now)<30
                rst.requery
            wend
            if rst("activo")<>1 then 'Se salió por timeout
                rst("activo")=0
                rst("sesionid")=0
                rst("netsessionid")=null
                rst.update
                rst.close
                set rst=nothing
                set command = nothing
                set conn = nothing
                mode="error"
            else
                rst.close
                set rst=nothing
                set command = nothing
                set conn = nothing
                'Asignar un valor a la variable de ámbito "usuario" de la sesion.'
                session("usuario") = u
                
                RedirigeUser
            end if
        end if
    else
        rst.close
        mode="error"
    end if
    set rst=nothing
    set command = nothing
    set conn = nothing
end if

' José Miguel Martínez --> JMMM --> 04/12/2009
' ######### Inicio / ACCESO POR MEDIO DE LA APLICATECA #########
if mode="sso_aplicateca" then
    'Recogemos los parámetros
    dim sessionId
    dim userName
    dim realmName
    sessionId = Request("sessionId")
    userName = Request("UserName")
    realmName = Request("RealmName")

    set accesoAPL = Server.CreateObject("ADODB.Recordset")
    'accesoAPL.Open "select * from indice where id_usuario = '" & userName & "' and claveop = '" & sessionId & "'", DsnIlion, adOpenKeyset, adLockOptimistic
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open dsnilion
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="select id_usuario,entrada from indice with(nolock) where claveop = '" & sessionId & "' and id_usuario = '" & userName & "'"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@claveop",adVarChar,adParamInput,50,sessionId)
    command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,150,userName)
    set accesoAPL = command.Execute

    'accesoAPL.Open "select id_usuario from indice with(nolock) where claveop = '" & sessionId & "' and id_usuario = '" & userName & "'", DsnIlion, adOpenKeyset, adLockOptimistic

    if (accesoAPL.EOF) then
        accesoAPL.Close
        %><script type="text/javascript" language="javascript">document.location = "<%=paginaAcceso%>?mode=error&d=<%=enc.EncodeForJavascript(parametroD)%>"</script><%
    else
        userName = accesoAPL("ID_USUARIO")
        usuarioIndice=accesoAPL("entrada") 'd_lookup("entrada","indice","id_usuario='" & userName & "'",DSNIlion)
        session("usuario") = usuarioIndice
        session("ERPAPL") = "yes"
        accesoAPL.Close
        RedirigeUser
    end if
    'conn.close
    set accesoAPL=nothing
    set command = nothing
    set conn = nothing
end if
' ######### FIN / ACCESO POR MEDIO DE LA APLICATECA #########%>

    <script type="text/javascript" language="javascript">
        var bName = navigator.appName;
        var bVer = navigator.appVersion;
        function ComprobarValores() {
            ok = true;
            if (document.demo.empresa.value == "") {
                alert("<%=LitFaltaNombreEmp%>");
                return false;
                ok = false;
            }
            if (document.demo.persona.value == "") {
                alert("<%=LitFaltaContacto%>");
                return false;
                ok = false;
            }
            if (document.demo.poblacion.value == "") {
                alert("<%=LitFaltaPoblacion%>");
                return false;
                ok = false;
            }
            if (document.demo.telefono.value == "") {
                alert("<%=LitFaltaTelefono%>");
                return false;
                ok = false;
            } else {
                /*if (isNaN(document.demo.telefono.value)) {
                alert("Debe proporcionar un número de teléfono correcto");
                return false;
                ok=false;
                }*/
            }
            if (document.demo.email.value == "") {
                alert("<%=LitFaltaMail%>");
                return false;
                ok = false;
            }
            if (ok) {
                document.demo.submit();
            }
        }
    </script>
    <iframe id='frRefresh' src='refreshNet.asp' width="10" height="10" frameborder="no" scrolling="no" noresize="noresize" style="display: none"></iframe>
    <%if DSNCronos & "" <> "" then %>
        <iframe id='frSession' src='/<%=CarpetaProduccionX%>/desactiva.aspx' width="100" height="100" frameborder="no" scrolling="no" noresize="noresize" style="display: none"></iframe>
        <iframe id='frAnulaSessionNet4' src='/<%=CarpetaProduccionX4%>/desactiva.aspx' width="100" height="100" frameborder="0" scrolling="no" style="display: none"></iframe>
        <iframe id='frAnulaSessionNet45' src='/<%=CarpetaProduccionX4%>5/desactiva.aspx' width="100" height="100" frameborder="0" scrolling="no" style="display: none"></iframe>
    <%end if

    '201102
    if limpiaCadena(request("viene"))="tienda" then
        session("usuariosale")=session("usuarioMKP")
    else
        session("usuariosale")=session("usuario")
    end if

    lEncryptedB64SLO_Data= session("slo")&""

    if mode="fin" then

       %>
        <iframe id='frAnulaSessionNet45_' src='/<%=CarpetaProduccionX4%>5/desactiva.aspx' width="100" height="100" frameborder="0" scrolling="no" style="display: none"></iframe>
   
        <%
        urlRedirect=""

        usuario=session("usuario")

        if limpiaCadena(request("viene"))="tienda" then
            ncompany=request.QueryString("ncliente")
        else
            ncompany=session("ncliente")
        end if

        'get distributor by parametroD
        if (usuario&""="" or ncompany&""="") and parametroD&"">"" and len(parametroD)>5 then
      
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open dsnilion
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="select ndistribuidor from distribuidores with(nolock) where id_partner =?"
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@idpartner",adVarChar,adParamInput,50,parametroD)
            set rstD= command.Execute
            if not rstD.eof then
                ndistributor=rstD("ndistribuidor")
            end if
            rstD.close
            conn.close
            set conn = nothing
            set command = nothing
            set rstD=nothing


        end if

        if ndistributor &""=""  then

        'GET DISTRIBUTOR BY COMPANY OR USER
            set rstDistibutor = Server.CreateObject("ADODB.Recordset")
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open dsnilion
            command.ActiveConnection =conn
            command.CommandTimeout = 0

            command.CommandType = adCmdStoredProc
            command.CommandText= "GetDistributorPref"
            command.Parameters.Append command.CreateParameter("@ncompany",adVarChar,adParamInput,5,ncompany&"")
            command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,usuario&"")

            'command.CommandText="select ndistribuidor from clientes with(nolock) where ncliente = ?"
            'command.CommandType = adCmdText
            'command.Parameters.Append command.CreateParameter("@ncompany", adVarChar, adParamInput, 5, ncompany&"")

            set rstDistibutor = command.execute

            if not rstDistibutor.eof then
                ndistributor = rstDistibutor("ndistribuidor")
            end if

            rstDistibutor.close

            set command = nothing
        end if

         'GET URLOUT BY DISTRIBUTOR

         'urlRedirect = d_lookup("urlout", "distributor_roles", "dtid=2 and ndistributor = '" & ndistributor & "'", dsnilion)

        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open dsnilion
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="select urlout from distributor_roles with(nolock) where dtid=2 and ndistributor = ?"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, ndistributor&"")

        set rstDistibutor = command.execute

        if not rstDistibutor.eof then
            urlRedirect = rstDistibutor("urlout")

            'TODO 12/04/2019: Concatenar el session("usuario") en la queryString de la url. Puede ya tener algún parámetro o no, cuidado con el ? o el &

            if lEncryptedB64SLO_Data&"">"" then
                urloutCharParam="?"
                if(InStr(urlRedirect,  "?") > 0) then
                    urloutCharParam="&"
                end if
                urlRedirectJS=urlRedirect+urloutCharParam+"slo="+ enc.EncodeForUrl(enc.DecodeFromUrl(lEncryptedB64SLO_Data))
                urlRedirect=""
            end if

        else
            urlRedirect=paginaAcceso
        end if
        
        'response.write("<br>......urlRedirect="&urlRedirect&"")
        'response.write("<br>---distributor="&ndistributor)
        'response.End


        rstDistibutor.close

        conn.close
        set rstDistibutor = nothing
        set command = nothing
        set conn = nothing

        BorrTablTemp

        'session("nombre")=""
        'session("usuario")=""
        'session("dsn_cliente")=""
        '201102
        if limpiaCadena(request("viene"))="tienda" then
            session("usuarioMKP")=""
            session("dsn_nclienteMKP")=""
            'Guardo el usuario por si existiera otra session
            if session("usuario")&"">"" then
                sessionUsuarioTMP=session("usuario")
            end if
            session("usuario")=""
        else
            session("nombre")=""
            session("usuario")=""
            session("dsn_cliente")=""
        end if
        ' José Miguel Martínez --> JMMM --> 07/12/2009
        ' Cierre de ventana para aplicateca (si tiene flag de session("ERPAPL") = "yes" se cierra la ventana)
        ruta=GenerarURL

        if session("ERPAPL") = "yes" then
            session("ERPAPL") = ""
            session.abandon%>
            <script type="text/javascript" language="javascript">
                CerrarVentanaPrincipalSinPreguntar();
            </script>
        <%elseif session("OANPOL") = "yes" then
            session("OANPOL") = ""
            session.abandon%>
            <script type="text/javascript" language="javascript">
                parent.window.open('', '_self');
                parent.window.close();
                parent.window.close();
            </script>
        <%else
            session("ERPAPL") = ""
            session("OANPOL")  = ""
            session.abandon
            

            if urlRedirectJS&"" = "" then
                urlRedirectJS="/" & CarpetaProduccion & "/" & LitDirSalEgesticet
            end if
        end if

        if urlRedirectJS&"">"" then
            %><script language="javascript" > parent.document.location="<%=urlRedirectJS%>"; </script><%
        end if

        if urlRedirect & "" <> "" then
            response.Redirect urlRedirect
        end if

        response.end
    end if

    if mode="comprueba" then
        if Acceder()="ACTIVAR" then
            mode="showdata"
        end if
    end if


Alarma "acceso.asp"

if mode="showdata" then
    if hd_Act=0 then
        nombreAc=request.form("nombre")
        passAc=request.form("contrasenya")
    else
        nombreAct=strCodigo
        passAc="OK"
    end if

    'rst.Open "SELECT nombre,nactivacion,movil,email, datecontract FROM indice with(nolock) where entrada='" & nombreAc & "'", DsnIlion, adOpenKeyset, adLockOptimistic
    set conn = server.CreateObject("ADODB.Connection")
    conn.open DsnIlion
    conn.cursorlocation=3
    set cmd_indice = server.CreateObject("ADODB.Command")
    cmd_indice.ActiveConnection =conn
    cmd_indice.CommandType = adCmdStoredProc
    cmd_indice.CommandText= "GetDataIndices"
    cmd_indice.Parameters.Append cmd_indice.CreateParameter("@user",adVarChar,,30,nombreAc)
    cmd_indice.Parameters.Append cmd_indice.CreateParameter("@id_user",adVarChar,,150,nombreAc)
    set rst = cmd_indice.Execute%>

    <form method="post" name="entrada">
        <input type="hidden" name="nombre" value="<%=enc.EncodeForHtmlAttribute(nombreAc)%>">
        <input type="hidden" name="contrasenya" value="<%=enc.EncodeForHtmlAttribute(passAc)%>">

        <%

    if passAc<>rst("nactivacion") and hd_Act=0 then

        'MAP 30/06/2013 - Obtiene la urlOut de la tabla distribuidores para mostrar en caso de que el código de activación no sea correcto (proyecto orange)

        'response.write("ndistributorGeneral="+ndist)

        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open dsnilion
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="select urlout from distributor_roles with(nolock) where dtid=2 and ndistributor = ?"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, ndist&"")

        set rstDistibutor = command.execute

        if not rstDistibutor.eof then
            urlRedirect = rstDistibutor("urlout")
        end if

        rstDistibutor.close

        'conn.close
        set rstDistibutor = nothing
        set command = nothing
        'set conn = nothing


        'response.write("<br>urlRedirect="+urlRedirect+"<br>PagAcceso="+paginaAcceso)
        'response.End()


        if urlRedirect&""="" then
            if lEncryptedB64SLO_Data&"">"" then
                urloutCharParam="?"
                if(InStr(urlRedirect,  "?") > 0) then
                    urloutCharParam="&"
                end if
                urlRedirect=urlRedirect+urloutCharParam+"slo="+ enc.EncodeForUrl(enc.DecodeFromUrl(lEncryptedB64SLO_Data))
            end if

            %>
            <script type="text/javascript" language="javascript">
                alert("<%=LitErrNumActivacion%>");
                document.entrada.action = "<%=paginaAcceso%>";
                document.entrada.submit();
            </script>
            <%
        else
            %>
            <script type="text/javascript" language="javascript">
                alert("<%=LitErrNumActivacion%>");
                document.entrada.action = "<%=urlRedirect%>";
                document.entrada.submit();
            </script>
            <%
        end if
   else
        randomize
        id_imagen = Int((3 - 1 + 1) * Rnd + 1)
        if id_imagen=1 then
            img="under_footer1"
        elseif id_imagen=2 then
            img="under_footer2"
        elseif id_imagen=3 then
            img="under_footer3"
        end if

        URLACTIVACION=""
        ''usuario_a_comprobar=usuario
        ''if usuario_a_comprobar & ""="" then
            usuario_a_comprobar=nombreAc
        ''end if
''response.write("el URLACTIVACION 1 es-" & URLACTIVACION & "-" & ndist & "-" & ncompany & "-" & usuario & "-" & nombreAc & "-" & passAc & "-" & folder & "-" & usuario_a_comprobar & "-<br>")
        if ndist &""=""  then
            'GET DISTRIBUTOR BY USER
            set rstDistibutor = Server.CreateObject("ADODB.Recordset")
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open dsnilion
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="select top 1 NDISTRIBUIDOR from INDICE with(nolock) where ENTRADA = ? and NDISTRIBUIDOR is not null and nactivacion is not null and confirma is null order by NDISTRIBUIDOR desc"
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,usuario_a_comprobar&"")
            set rstDistibutor = command.execute
            if not rstDistibutor.eof then
                ndist = rstDistibutor("ndistribuidor")
            end if
            rstDistibutor.close
            set command = nothing
            set rstDistibutor =nothing
        end if
''response.write("el URLACTIVACION 2 es-" & URLACTIVACION & "-" & ndist & "-<br>")
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open dsnilion
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="select URLACTIVACION from distributor_roles with(nolock) where dtid=2 and ndistributor = ?"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, ndist&"")

        set rstDistibutor = command.execute

        if not rstDistibutor.eof then
            URLACTIVACION = rstDistibutor("URLACTIVACION")
        end if
''response.write("el URLACTIVACION 3 es-" & URLACTIVACION & "-" & ndist & "-<br>")

        rstDistibutor.close

        'conn.close
        set rstDistibutor = nothing
        set command = nothing

        if URLACTIVACION & "">"" then
            URLACTIVACION=replace(replace(URLACTIVACION,"##ndist##",ndist&""),"##usuario##",usuario_a_comprobar&"")
            %>
            <script type="text/javascript" language="javascript">
                document.location.href="<%=URLACTIVACION%>&frm=<%=paginaAcceso%>";
            </script>
            <%
        else
            %>
            <div id="line_menu_ppal"></div>
            <div id="logos_body"></div>
            <div class="folders">
                <h4><%=LitDatosCuenta%></h4>
                <div class="table-grid col-xxs-12">
                    <div class="col-xs-4 col-xxs-6 table-grid-item">
                        <h5><%=LitNombre%></h5>
                        <div>
                            <strong><%=rst("nombre")%></strong>
                        </div>
                    </div><div class="col-xs-4 col-xxs-6 table-grid-item">
                        <h5><%=LitUsuario%></h5>
                        <div>
                            <strong><%=nombreAc%></strong>
                        </div>
                    </div><div class="col-xs-4 col-xxs-6 table-grid-item">
                        <h5><%=LitNumActivacion%></h5>
                        <div>
                            <strong><%=passAc%></strong>
                        </div>
                    </div>
                </div>
                <%set con_contract =Server.CreateObject("ADODB.Connection")
                set cmd_contract = server.CreateObject("ADODB.Command")
                set rstContract = server.CreateObject("ADODB.Recordset")
                con_contract.open dsnIlion
                con_contract.cursorlocation=3
                cmd_contract.ActiveConnection =con_contract
                cmd_contract.CommandType = adCmdStoredProc
                cmd_contract.CommandText= "GetContract"
                cmd_contract.Parameters.Append cmd_contract.CreateParameter("@folder",adVarChar,,20,CarpetaProduccion)
                set rstContract = cmd_contract.Execute
                if not rstContract.eof then
                strContact=rstContract("contract")
                end if
                con_contract.close
                set cmd_contract=nothing
                set rstContract=nothing
                set con_contract=nothing
                chkdisabled=""
                if rst("datecontract")&"">"" then
                chkdisabled=" disabled checked"
                end if%>
                <div class="col-xxs-12">
                    <h6>CONTRASEÑA DEL SISTEMA</h6>
                    <%
                        'MAP 11/07/2013 - MODIFICADO PARA ORANGE
                        'response.Write("ndist entradilla="&ndist)
                        if ndist="01000" or ndist="01100" then 
                            %><div class="col-xs-4 col-xxs-6 table-grid-item">
                            <h5><%=LitIntroduceMail%></h5>
                        </div><%else %><div class="col-xs-4 col-xxs-6 table-grid-item">
                                <%=LitDatosCorrectos%><h5><%=LitDatosNoCorrectos%></h5>
                        </div><%end if %>
                    
                    
                        <% 'MAP 11/07/2013 - MODIFICADO PARA ORANGE
                        'response.write("distribuidor="&ndist)
                    
                        if ndist="01100" or ndist="01000" then
                            %>
                            
                             <div class="col-xs-6 col-xxs-12 pBottom20">
                                <label><%=LitDirCorreo%></label>
                                <input type="text" name="email" value="<%=rst("email")%>" />

                                <input type="checkbox" name="chkContract" <%=chkdisabled%> style="display:none" checked>

                                <div class="col-sm-3 col-xs-4 col-xxs-6 noPLeft">
                                <a onmouseover="javascript:window.status='<%=LitActivar%>';return true;" onmouseout="javascript:window.status='';" href="javascript:activarCuenta();" name="activar" style="cursor: pointer;">
                                    <%if ndLitOk = "1" then %>
                                        <div id="button_enter"> <%=LitAceptar2 %></div></a>
                                    <% else %>
                                        <div id="button_enter"> </div></a>
                                    <% end if %> </div><div class="col-sm-3 col-xs-4 col-xxs-6 noPRight">
                                <a onmouseover="javascript:window.status='<%=LitCancelar%>';return true;" onmouseout="javascript:window.status='';" href="javascript:document.entrada.action='<%=paginaAcceso%>?d=<%=enc.EncodeForHtmlAttribute(parametroD)%>';document.entrada.submit();" name="cancelar" style="cursor: pointer;">
                                    <%if ndLitOk = "1" then %>
                                        <div id="button_enter"> <%=LitCancelar2 %></div></a>
                                    <% else %>
                                        <div id="button_enter"> </div></a>
                                    <% end if %></div>
                            </div>

                        <%else%>
                            <div class="col-xs-6 col-xxs-12">
                                <label><%=LitMovil%></label>
                                <input type="text" name="movil" /><br />
                            </div><div class="col-xs-6 col-xxs-12 pBottom20">
                                <label><%=LitDirCorreo%></label>
                                <input type="text" name="email" value="<%=rst("email")%>" />
                            </div><div id="accordin clearfix">
                                    <h5><%=strContact%></h5>
                                </div>
                                <div>
                                    <%=LitchkContract%>
                                </div>
                                <div class="col-xxs-12 noPH">
                                    <label for="chkContract" class="label-checkbox"><%=LitchkContract%></label>
                                    <input type="checkbox" name="chkContract" <%=chkdisabled%>/>
                                </div>
                                <div class="col-sm-3 col-xs-4 col-xxs-6 noPLeft">
                                <a onmouseover="javascript:window.status='<%=LitActivar%>';return true;" onmouseout="javascript:window.status='';" href="javascript:activarCuenta();" name="activar" style="cursor: pointer;">
                                    <%if ndLitOk = "1" then %>
                                        <div id="button_enter"> <%=LitAceptar2 %></div></a>
                                    <% else %>
                                        <div id="button_enter"> </div></a>
                                    <% end if %></div>
                                <div class="col-sm-3 col-xs-4 col-xxs-6 noPRight">
                                    <a onmouseover="javascript:window.status='<%=LitCancelar%>';return true;" onmouseout="javascript:window.status='';" href="javascript:document.entrada.action='<%=paginaAcceso%>?d=<%=enc.EncodeForHtmlAttribute(parametroD)%>';document.entrada.submit();" name="cancelar" style="cursor: pointer;">
                                    <%if ndLitOk = "1" then %>
                                        <div id="button_enter"> <%=LitCancelar2 %></div></a>
                                    <% else %>
                                        <div id="button_enter"> </div></a>
                                    <% end if %></div>      
                        <%end if %>
                    <div class="txt-important">
                        <p><%=LitLlamarTelefono%></p>
                    </div>
            </div>
            <%
        end if
    end if
    conn.close
    set cmd_indice=nothing
    set rst=nothing
    set conn=nothing%>
    </form>
    <script type="text/javascript" language="javascript">document.entrada.movil.focus();</script>
    <%
elseif mode="activar" then
    'rst.Open "SELECT * FROM indice with(updlock) where entrada='" & request.form("nombre") & "'", DsnIlion, adOpenKeyset, adLockOptimistic
    set conn = server.CreateObject("ADODB.Connection")
    conn.open DsnIlion
    conn.cursorlocation=3
    set cmd_indice_act = server.CreateObject("ADODB.Command")
    cmd_indice_act.ActiveConnection =conn
    cmd_indice_act.CommandType = adCmdStoredProc
    cmd_indice_act.CommandText= "GetDataIndices"
    cmd_indice_act.Parameters.Append cmd_indice_act.CreateParameter("@user",adVarChar,,30,request.form("nombre"))
    set rst = cmd_indice_act.Execute
    if not rst.eof then
        version=rst("version")
        session("usuario") = rst("nactivacion")
        nombre=rst("nombre")
        entrada=rst("entrada")
    end if
    set cmd_indice_act=nothing
    set rst=nothing%>

    <div id="Div1" style="position: absolute; left: 64px; top: 148px; width: 636px; height: 144px; z-index: 4">
        <form method="post" name="entrada">

            <input type="hidden" name="nombre" value="<%=enc.EncodeForHtmlAttribute(request.form("nombre"))%>">
            <input type="hidden" name="contrasenya" value="<%=enc.EncodeForHtmlAttribute(request.form("contrasenya"))%>">
            <%'Generar la contraseña
    pwd=newPassword()
    pwdCifrado=cifrar(pwd,version)
    login=request.form("nombre")

    'Enviar contraseña por email.

    conn.cursorlocation=3
    set cmd_indice_act = server.CreateObject("ADODB.Command")
    cmd_indice_act.ActiveConnection =conn
    cmd_indice_act.CommandType = adCmdStoredProc
    cmd_indice_act.CommandText= "GERUSEREMAILPWD"
    cmd_indice_act.Parameters.Append cmd_indice_act.CreateParameter("@login",adVarChar,,30,login)
    cmd_indice_act.Parameters.Append cmd_indice_act.CreateParameter("@pwd",adVarChar,,30,pwd)
    cmd_indice_act.Parameters.Append cmd_indice_act.CreateParameter("@name",adVarChar,,100,nombre)
    cmd_indice_act.Parameters.Append cmd_indice_act.CreateParameter("@distribuidor",adVarChar,,5,ndist)

    on error resume next
    set rst = cmd_indice_act.Execute
    'MAP 011/07/2013 CAMBIO PARA EL ENVÍO DE CONTRASEÑ DE ORANGE
    if ndist="01000" or ndist="01100" then
    remitente_correo=rst("sender")
    asunto_correo = rst("subject")
    strMensaje=rst("body")

    strMensaje=replace(strMensaje,"xxModalAsxx","http"&letra&"://"&Request.ServerVariables("SERVER_NAME")&"")
    strMensaje=replace(strMensaje,"xxImgLogoxx","http"&letra&"://"&Request.ServerVariables("SERVER_NAME")&"")
    strMensaje=replace(strMensaje,"xxUserxx",login)

''response.write("el strMensaje 1 es-" & strMensaje & "-<br>")

    else
        if err.number<>0 then
            on error goto 0
            remitente_correo=LitRemitente
            asunto_correo=LitAsunto & " - " & nombre
            ''MPC 10/09/2008 Nuevo mensaje de correo cuando te das de alta.
            strMensaje="<table width='100%' border='0'><tr><td style='font-Family: Verdana;font-size: 8.0pt;text-align: left;color: #000000;'><img src='http"&letra&"://"&Request.ServerVariables("SERVER_NAME")&"/"&CarpetaProduccion&"/images/everilion_ilionstore.png' /><br><br>Estimado cliente: <br><br>" & _
            LitMensaje1 & nombre & "<br><br>" & _
            LitMensaje2 & "<br><br>" & LitMensaje3 & "<b>" & pwd & _
            "</b><br><br>" & LitMensaje4 & "<br><br>" & LitMensaje5 & ". <br>" & LitSaludos & "<br><br>" & LitEverilion &"</td></tr></table>"
            'strMensaje="&nbsp;&nbsp;&nbsp;&nbsp;" & LitMensaje1 & rst("nombre") & "<br><br>" & _
            'LitMensaje2 & "<br><br>" & "&nbsp;&nbsp;&nbsp;&nbsp;" & LitMensaje3 & "<b>" & pwd & _
            '"</b><br><br>" & LitMensaje4 & "<br><br>" & LitMensaje5 & ". " & LitSaludos & "<br><br>" & _
            '"_____________________________________________" & "<br>" & LitEnviadoDesde & _
            '"&nbsp;<a href=" & LitUrlDesde & ">" & LitUrlDesde & "</a><br>"
        else
            on error goto 0
            ''if not rst.eof then
                on error resume next
                remitente_correo=rst("sender")
                if err.number<>0 then
                    remitente_correo=LitRemitente
                end if
                asunto_correo = rst("subject")
                if err.number<>0 then
                    asunto_correo=LitAsunto & " - " & nombre
                end if
                strMensaje=rst("body")
                if err.number<>0 then
                    ''MPC 10/09/2008 Nuevo mensaje de correo cuando te das de alta.
                    strMensaje="<table width='100%' border='0'><tr><td style='font-Family: Verdana;font-size: 8.0pt;text-align: left;color: #000000;'><img src='http"&letra&"://"&Request.ServerVariables("SERVER_NAME")&"/"&CarpetaProduccion&"/images/everilion_ilionstore.png' /><br><br>Estimado cliente: <br><br>" & _
                    LitMensaje1 & nombre & "<br><br>" & _
                    LitMensaje2 & "<br><br>" & LitMensaje3 & "<b>" & pwd & _
                    "</b><br><br>" & LitMensaje4 & "<br><br>" & LitMensaje5 & ". <br>" & LitSaludos & "<br><br>" & LitEverilion &"</td></tr></table>"
                    'strMensaje="&nbsp;&nbsp;&nbsp;&nbsp;" & LitMensaje1 & rst("nombre") & "<br><br>" & _
                    'LitMensaje2 & "<br><br>" & "&nbsp;&nbsp;&nbsp;&nbsp;" & LitMensaje3 & "<b>" & pwd & _
                    '"</b><br><br>" & LitMensaje4 & "<br><br>" & LitMensaje5 & ". " & LitSaludos & "<br><br>" & _
                    '"_____________________________________________" & "<br>" & LitEnviadoDesde & _
                    '"&nbsp;<a href=" & LitUrlDesde & ">" & LitUrlDesde & "</a><br>"
                end if
                on error goto 0
        ''end if
    end if

    end if
    on error goto 0
    set cmd_indice_act=nothing
    set rst=nothing

''response.write("el strMensaje 2 es-" & strMensaje & "-<br>")

    EnviarCorreo remitente_correo,request.form("email"),asunto_correo,strMensaje,"","",2,"",0
''response.end
    strMensajeOK=LitEnviadoMail
    strMsgAuditar=LitMensajeEnviadoAc & LitDirCorreo & ":" & request.form("email") & "."

    'Comprobar si hay que enviar la contraseña por SMS.
    if request.form("movil")>"" then
        strMensaje=LitGesticetInforma & " : " & LitMensaje3 & " : " & pwd & ". " & LitMensaje5
        EnviarSMS request.form("movil"),strMensaje
        strMensajeOK=strMensajeOK & LitEnviadoSMS
        strMsgAuditar=strMsgAuditar & LitMovil & ":" & request.form("movil")
    end if
    strMensajeOK=strMensajeOK & LitGracias

    'Auditar la activación de la cuenta.
    Auditar d_lookup("ncliente","clientes_users","usuario='" & entrada & "'",DSNIlion), _
    entrada,session("usuario2"),LitActivacion,Request.ServerVariables(CLIENT_IP),Request.ServerVariables("REMOTE_HOST"),strMsgAuditar,DSNIlion

    'Actualizar el registro.Muevo el bloque para que se ejecute después del envío de sms
    set cmd_updIndice = server.CreateObject("ADODB.Command")
    cmd_updIndice.ActiveConnection =conn
    cmd_updIndice.CommandType = adCmdStoredProc
    cmd_updIndice.CommandText= "UpdateDataIndice"
    cmd_updIndice.Parameters.Append cmd_updIndice.CreateParameter("@user",adVarChar,,30,entrada)
    cmd_updIndice.Parameters.Append cmd_updIndice.CreateParameter("@confirma",adVarChar,,50,pwdCifrado)
    cmd_updIndice.Parameters.Append cmd_updIndice.CreateParameter("@email",adVarChar,,100,nulear(request.form("email")))
    'if request.form("movil")>"" then
        cmd_updIndice.Parameters.Append cmd_updIndice.CreateParameter("@movil",adVarChar,,15,nulear(request.form("movil")))
    'END IF
    cmd_updIndice.Parameters.Append cmd_updIndice.CreateParameter("@err",adVarChar,adParamOutput,1)
    cmd_updIndice.Execute,,adExecuteNoRecords
    conn.close
    set cmd_updContract=nothing
    set conn=nothing
    
    'MAP 11/07/2013 MODIFICADO PARA ORANGE
    if ndist="01000" or ndist="01100" then
        %>
        <script type="text/javascript" language="javascript">
            alert("<%=strMensajeOK%>");
            parent.location = "accesoHubble.asp?d=<%=enc.EncodeForJavascript(parametroD)%>";
        </script>
    <%else %>
        <script type="text/javascript" language="javascript">
            alert("<%=strMensajeOK%>");
            parent.location="<%=iif(Request.ServerVariables("HTTPS")="on","https","http")%>://www.ilionsistemas.com";
        </script>
    <%end if %>
        </form>
    </div>
    <%else
    if logoDistribuidor<>"false"  then
        pos=pos-40
    end if
    if mode="error" then
        randomize
        casilla = Int((50 - 1 + 1) * Rnd + 1)
        'casilla=35
        casilla = String(2 - Len(CStr(casilla)), "0") & casilla
        randomize
        id_imagen = Int((3 - 1 + 1) * Rnd + 1)
        if id_imagen=1 then
            img="under_footer1"
            txt="footer_white"
        elseif id_imagen=2 then
            img="under_footer2"
            txt="footer_black"
        elseif id_imagen=3 then
            img="under_footer3"
            txt="footer_black"
        end if%>
    <div id="line_menu_ppal"></div>
    <form name="entrada" action="" method="post" onkeypress="javascript:keypressed();">
        <%if parametroD>"" and mode<>"demo" and mode<>"fin" then%>
        <div class="bck-wrapper">
            <img class="bck" id="bck">
        </div>
        <div class="bck-text animated slideInUp">
            <img id="text">
        </div>
        <script type="text/javascript">
            function getRandomInt(min, max) {
                return Math.floor(Math.random() * (max - min)) + min;
            }
            var imagesArray = ["images/img_bck1.jpg", "images/img_bck2.jpg", "images/img_bck4.jpg", "images/img_bck5.jpg", "images/img_bck6.jpg", "images/img_bck7.jpg", "images/img_bck8.jpg"]
            var txtArray = ["images/txt_black.png", "images/txt_black.png", "images/txt_white.png", "images/txt_black.png", "images/txt_white.png", "images/txt_white.png", "images/txt_black.png"]
            var result = getRandomInt(0, 7)
            document.getElementById("bck").src = "/lib/estilos/<%=folder%>/" + imagesArray[result];
            document.getElementById("text").src = "/lib/estilos/<%=folder%>/" + txtArray[result];

        </script>
        <div style="overflow: hidden;">
            <table width="100%">
                <tr>
                    <!--<td style="width: 50%">-->
                    <td style="width: 100%">
                        <div id="logos_body_everilion" style="width: 100%" class="logo-login animated slideInDown"></div>
                    </td>
                    <!--<td style="width: 50%">
                        <div id="logos_body_ndistributor" style="width: 100%; background: url('configuracion/muestra_logoDist.asp?viene=acceso&distribuidor=<%=enc.EncodeForHtmlAttribute(nDistribuidor)%>') no-repeat left center;">
                        </div>
                    </td>-->
                </tr>
            </table>
        </div>
        <%else%>
        <div class="bck-wrapper">
            <img class="bck" id="bck">
        </div>
        <div id="logos_body"></div>
        <%end if%>
        <div class="folders_error">
            <div id="icon_alert"></div>
            <div id="error">
                <div>
                    <%if request.querystring("causa")="UsuarioActivo" then%>
                    <p><%=LitUsuarioActivo%></p>
                    <!--<p id="tel_error"><%=LitTelIncidencias%></p>-->
                    <%else%>
                    <p><%=LitEntradaKO%></p>
                    <!--<p id="tel_error"><%=LitTelIncidencias%></p>-->
                    <%end if%>
                </div>
            </div>
            <%if ndist = "00025" or ndist="01000" then%>
                <table id="form_error" cellpadding="0" cellspacing="0" class="login-box animated zoomIn">
             <%else%>
                <table id="form" cellpadding="0" cellspacing="0" class="login-box animated zoomIn">
             <%end if%>
                <tr>
                    <td><label><%=LitUsuario%> </label></td></tr>
                    <tr>
                    <td colspan="2">
                        <%if adminpwd<>"True" then%>
                            <input type="text" name="nombre" placeholder="<%=LitUsuario%>" class="width100"/></td>
                        <%else %>
                            <input type="text" name="nombre" placeholder="<%=LitUsuario%>" class="width100" onchange="bloqPassword_orange()"/></td>
                        <%end if %>
                </tr>
                <tr>
                    <td><label><%=LitContrasenya%> </label></td></tr>
                    <tr>
                    <td colspan="2">
                        <!--<input type="password" name="contrasenya" autocomplete="off" /></td>-->

                    <%if adminpwd<>"True" then%>                        
                            <input type="password" name="contrasenya" placeholder="<%=LitContrasenya%>" class="width100" autocomplete="off" /></td>
                    <%else %>
                            <input type="password" name="contrasenya" placeholder="<%=LitContrasenya%>" class="width100" autocomplete="off" disabled="disabled"/></td>
                    <%end if %>

                </tr>
                <tr id="casilla">
                    <td><label><%=LitCasilla%>&nbsp;<%=casilla%> </label></td></tr>
                    <tr>
                    <td colspan="2">
                       <input type="text" maxlength="2" name="vcasilla" placeholder="Casilla X"/><input type="hidden" value="<%=casilla%>" name="ncasilla" />
                    </td>
                </tr>

                <%'MAP 16/05/2013 - Nueva clase recoveryPass para mostrar u ocultar el link de "recordar contraseña" %>
                <% if ndist = "00025" then%>
                    <tr>
                    <td></td>
                    <td class="data recoveryPass">
                        <a href="javascript:RecoveryPass();">Recordar contraseña</a>
                    </td>
                   <% else%>
                    <%if adminpwd="True" then %>

                     <td class="data recoveryPass">
                        <a href="javascript:RecoveryPass();">Olvidé mi contraseña</a>
                    </td>
                    <%end if %>

                </tr>
                <%end if %>
                
                 <tr id="buttonenter">
                    <%if ndist = "00025" then%>
                        <td colspan="2">
                            <input type="button" name="enter" value="Entrar" onmouseover="javascript:window.status='Entrar';return true;" onmouseout="javascript:window.status='';" onclick="javascript:ValidarDatos();" />
                        </td>
                    <%else%>
                    <td colspan="2">
                        <a onmouseover="javascript:window.status='Entrar';return true;" onmouseout="javascript:window.status='';" href="javascript:ValidarDatos();" name="entrar" style="cursor: pointer;">
                            <%if ndLitOk = "1" then %>
                                <div id="button_enter" class="ui-button wide"> <%=LitEntrar %></div>
                            <% else %>
                                <div id="button_enter" class="ui-button wide"> </div>
                            <% end if %>      
                        </a>
                    </td>
                    <%end if%>
                </tr>
                <tr><td></td></tr>


                <%if adminpwd="True" then %>
                <tr>
                    <td></td>
                    <td class="data registerPass">
                        
                        <div class="uppercase"><b><%=newuserliteral%></b></div>
                        <a href="javascript:newOrangePass();" id="button_enter"></a>
                    </td>
                </tr>
                <%end if %>
            
                
     
            </table>
        </div>
        <div class="bck-text animated slideInUp">
            <img id="text">
        </div>
        <%if ndist = "00025" then%>
        <br />
        <br />
        <div id="capa_pie" style="padding-top: 10px">
            <p style="text-align: center; font-size: 11px; padding-bottom: 4px;">
                &copy; 2012 Tabarca Technologies. Todos los derechos reservados.
                <a href="http://grantbiomed.com/">CONDICIONES DE USO</a> I <a href="http://grantbiomed.com/">POL&Iacute;TICA DE PRIVACIDAD</a>
            </p>
            <div id="foot-2">
                <div id="foot-a">
                    <h3>M&aacute;s</h3>
                    <ul>
                        <li><a href="http://grantbiomed.com/">Inicio</a></li>
                        <li><a href="http://grantbiomed.com/#bc7">&iquest;Qu&eacute; es grant?</a></li>
                        <li><a href="http://grantbiomed.com/blog">Blog</a></li>
                        <li><a href="http://grantbiomed.com/#bc2">Planes y Precios</a></li>
                        <li><a href="http://grantbiomed.com/modulos">M&oacute;dulos</a></li>
                        <li><a href="http://grantbiomed.com/quienes-somos/">&iquest;Qui&eacute;nes somos?</a></li>
                        <li>Entrar</li>
                    </ul>
                </div>
                <div id="foot-b">
                    <h3>Descargas</h3>
                    <ul>
                        <li><a href="http://grantbiomed.com/">Iphone App (pr&oacute;ximamente)</a></li>
                        <li><a href="http://grantbiomed.com/">Android Market (pr&oacute;ximamente)</a></li>
                    </ul>
                </div>
                <div id="foot-d">
                    <h3>Enlaces</h3>
                    <ul>
                        <li><a href="http://grantbiomed.com/">Gal&eacute;nica</a></li>
                        <li><a href="http://tabarcatechnologies.com/">Tabarca</a></li>
                        <li><a href="http://cidiplus.com/">CIDi+</a></li>
                    </ul>
                </div>
                <div id="foot-c">
                    <h3>Contacto</h3>
                    <ul>
                        <li><a href="mailto:grant@tabarcatechnologies.com">Email</a></li>
                    </ul>
                </div>
            </div>
            <div style="margin-top: 10px; vertical-align: middle; text-align: center;">
                <img style="float: left; border: 0px;" src="/lib/estilos/<%=folder%>/images/zonasegura.png" alt="logos" width="74" height="44" />
                <img usemap="#Map" src="/lib/estilos/<%=folder%>/images/PIE-logos.jpg" border="0" alt="logos" />
            </div>
        </div>
        <%end if%>
    </form>
    <%end if
end if
if mode<>"showdata" then%>
    <script type="text/javascript" language="javascript">
        if (document.entrada.nombre != null) document.entrada.nombre.focus();
    </script>
    <%end if
set xmlhttp = nothing
set conn = nothing
set command = nothing
set rstAux = nothing
set acceso = nothing
set rs_empresas = nothing
set rs_empresas2 = nothing
set rs_Clientes = nothing
set rs_distrib = nothing
set rs_empresasMKP = nothing
set accesoONO = nothing
set rst = nothing
set accesoAPL = nothing
set cmd_indice = nothing
set con_contract = nothing
set cmd_contract = nothing
set rstContract = nothing
set cmd_indice_act = nothing
set cmd_updIndice = nothing
    %>
    <input type="hidden" id="hdmode" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>">
    <%set enc = nothing %>
</body>
</html>