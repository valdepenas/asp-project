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
end function%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
</HEAD>
<script language="JavaScript" src="../jfunciones.js"></script>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../applets/welcome.inc" -->
<!--#include file="../applets/modalAsistente.inc" -->
<!--#include file="../gestion/clientes/clientes.inc" -->
<%
urlAltaUsu=""
set rstAux = Server.CreateObject("ADODB.Recordset")
rstAux.open "select ID_PARTNER from DISTRIBUIDORES d with(nolock) inner join CLIENTES cl with(nolock) on cl.NDISTRIBUIDOR=d.NDISTRIBUIDOR and cl.NCLIENTE='"&session("ncliente")&"'", DSNILION
if not rstAux.eof then
    id = replace(replace(rstAux("ID_PARTNER"), "{", ""), "}", "")
end if
rstAux.close
set rstAux = nothing%>
<script type="text/javascript">
    function Reload() {
        
    }
    function AbrirCompra(idweb, session_ncliente) {
        var ran = Math.random();
        paginaModal = "/<%=CarpetaProduccion%>/cms/gestor/plantillas/mainIlion.asp?id=" + idweb + "&menu=seccion&ilionteca=step1&idioma=" + session_ncliente + "ES&load_chart=1&ran=" + ran;
        cambiarTamanyo("#fr_Asistente", "400", "1020");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function AbrirBajaUsuarios() {
        paginaModal = "../central.asp?pag1=gestion/clientes/clientesAsistente.asp&id=<%=enc.EncodeForJavascript(id)%>&viene=ContratacionBaja_Baja&pag2=gestion/clientes/clientesAsistente_bt.asp";
        //paginaModal = "../central.asp?pag1=<%=pagAssistant%>&viene=ContratacionBaja_Baja&pag2=<%=pagAssistant_bt%>";
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "<%=AltoVentana %>", "<%=AnchoVentana %>");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function AbrirBajaUsuariosEspc(paginaEspc)
    {
        paginaModal = paginaEspc + "&ndoc=ContratacionBaja_Baja";
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "<%=AltoVentana %>", "<%=AnchoVentana %>");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function AbrirCrearUsuarios() {
        paginaModal = "../central.asp?pag1=gestion/clientes/clientesAsistente.asp&id=<%=enc.EncodeForJavascript(id)%>&viene=ContratacionBaja_Alta&pag2=gestion/clientes/clientesAsistente_bt.asp";
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "<%=AltoVentana %>", "<%=AnchoVentana %>");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function AbrirCrearUsuariosEspc(paginaEspc)
    {
        paginaModal = paginaEspc;
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "<%=AltoVentana %>", "<%=AnchoVentana %>");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function AbrirBajaCompleta() {
        var ran = Math.random();
        paginaModal = "../central.asp?pag1=gestion/usuarios/configILTECA_Baja.asp&mode=browse&pag2=gestion/usuarios/configILTECA_Baja_bt.asp&ran=" + ran;
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "<%=AltoVentana %>", "<%=AnchoVentana %>");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function OpenAdminUsers() {
        var ran = Math.random();
        paginaModal = "../central.asp?pag1=gestion/usuarios/adminUsuarios.asp&mode=edit&OMC=NO&viene=sel_app&pag2=gestion/usuarios/adminUsuarios_bt.asp&ran=" + ran;
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "<%=AltoVentana %>", "<%=AnchoVentana %>");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function OpenChangeModule() {
        paginaModal = "../netInic.asp?pag=/ilionx4/Gestion/Admin/UpgradeDowngradeUsers.aspx";
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "260", "370");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
    function OpenChangeLicence() {
        paginaModal = "../netInic.asp?pag=/ilionx4/Gestion/Admin/ContractLicences.aspx";
        reloadIframe("#fr_Asistente", "")
        cambiarTamanyo("#fr_Asistente", "240", "450");
        reloadClass("#fr_Asistente", paginaModal);
        alPresionar("#fr_Asistente");
    }
</script>
<%viene=Request.QueryString("viene")
if viene <> "app" then
    color = color_blau
else
    color = color_blanc
end if %>
<body>
<%mode=EncodeForHtml(Request.QueryString("mode"))
if accesoPagina(session.sessionid,session("usuario"))=1 then
    if viene <> "app" then
        PintarCabecera "contratacion_baja.asp"
    end if%>
<form name="opciones" method="post">
    <%set rst = Server.CreateObject("ADODB.Recordset")
    ncompanyFact = d_lookup("COMPANY_FACT","ILIONTECA_CUSTOMERS","COMPANY_AFACT='" & session("ncliente") & "'",dsnilion)
    idweb=d_lookup("ID_WEB","clientes","ncliente='" & ncompanyFact & "'",dsnilion)
    rst.open "select company_afact from ilionteca_customers with(nolock) where company_afact='" & session("ncliente") & "'  and isnull(NRECDELIVERY,'')<>'' ", DSNILION
    disabled=""
    if rst.eof then
        disabled="disabled"
    end if
    rst.close
    paginaModal="blanco.asp"
    AbrirModal "fr_Asistente",paginaModal,0,0,"no","si","noresize","S","cerrar"
    
    strSelect = "select ISNULL(ADMINISTRATOR, '111110') as ADMINISTRATOR,URLALTAUSU from DISTRIBUTOR_ROLES dr with(nolock) inner join CLIENTES cl with(nolock) on cl.NCLIENTE = '" & session("ncliente") & "' and cl.NDISTRIBUIDOR = dr.NDISTRIBUTOR where dr.DTID=2"
    rst.open strSelect, DsnIlion, adUseClient, adLockReadOnly
    if not rst.eof then
        administrator = rst("ADMINISTRATOR")
        urlAltaUsu=rst("URLALTAUSU")&""
    else
        administrator = "111110"
    end if
    rst.close
    
    dim permissionAdministrator
    num_parametrosAdm=len(administrator)
    
    redim permissionAdministrator(num_parametrosAdm)
	for i=1 to num_parametrosAdm
		permissionAdministrator(i)=mid(administrator,i,1)
	next%>
    <br />
    <div id="TABLEADM">
    <%cont = 1
    while cont <= ubound(permissionAdministrator)
        if permissionAdministrator(cont) = "1" then

            if cont = 2 then%>
                <div class="imgADM">
                    <a href="javascript:OpenAdminUsers();"><div id="AdminUsers" class="imgContract"></div></a>
                    <a class="LINKADM" href="javascript:OpenAdminUsers();"><%=LitAdminUsers%></a>
                </div>
            <%end if

            if cont = 3 then
                if urlAltaUsu &"">"" then
                    %>
                    <div class="imgADM">
                        <a href="javascript:AbrirCrearUsuariosEspc('<%=replace(urlAltaUsu,"##ID##",id)%>');"><div id="createUser" class="imgContract"></div></a>
                        <a class="LINKADM" href="javascript:AbrirCrearUsuariosEspc('<%=replace(urlAltaUsu,"##ID##",id)%>');"><%=LITCONTRCREAUSU%></a>
                    </div>
                    <%
                else
                    %>
                    <div class="imgADM">
                        <a href="javascript:AbrirCrearUsuarios();"><div id="createUser" class="imgContract"></div></a>
                        <a class="LINKADM" href="javascript:AbrirCrearUsuarios();"><%=LITCONTRCREAUSU%></a>
                    </div>
                    <%
                end if
            end if

            if cont = 4 then
                if urlAltaUsu &"">"" then
                    %>
                    <div class="imgADM">
                        <a href="javascript:AbrirBajaUsuariosEspc('<%=replace(urlAltaUsu,"##ID##",id)%>');"><div id="UnSubsUser" class="imgContract"></div></a>
                        <a class="LINKADM" href="javascript:AbrirBajaUsuariosEspc('<%=replace(urlAltaUsu,"##ID##",id)%>');"><%=LITCONTRBAJAUSU%></a>
                    </div>
                    <%
                else
                    %>
                    <div class="imgADM">
                        <a href="javascript:AbrirBajaUsuarios();"><div id="UnSubsUser" class="imgContract"></div></a>
                        <a class="LINKADM" href="javascript:AbrirBajaUsuarios();"><%=LITCONTRBAJAUSU%></a>
                    </div>
                    <%
                end if
            end if

            if cont = 6 then%>
                <div class="imgADM">
                    <a href="javascript:OpenChangeModule();"><div id="Div1" class="imgContract"></div></a>
                    <a class="LINKADM" href="javascript:OpenChangeModule();"><%=LITTITLEUPGRADEDOWNGRADEUSERS%></a>
                </div>
            <%end if

            if cont = 7 then%>
                <div class="imgADM">
                    <a href="javascript:OpenChangeLicence();"><div id="Div2" class="imgContract"></div></a>
                    <a class="LINKADM" href="javascript:OpenChangeLicence();"><%=LITTITLICENCECONTRACT%></a>
                </div>
            <%end if

        end if
        cont = cont + 1
    wend%>
    </div>
</form>
<%end if
set rst=nothing%>
</BODY>
</HTML>