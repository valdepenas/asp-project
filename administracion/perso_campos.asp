<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>

<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="perso_camposFS.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../XSSProtection.inc" -->
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/listTable.css.inc" -->

<%dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
    function cambiar_det(sop, admin)
{  
	session_ncliente=document.perso_campos.session_ncliente.value;
	tabla=document.perso_campos.tla.value;	 
	
	fr_Tabla.document.location = "perso_campos_det.asp?mode=browse&session_ncliente=" + session_ncliente + "&tla=" + tabla + "&sop=" + sop + "&admin=" + admin; // + "&sentido=&lote=&campo=&criterio=&texto=";

	if (tabla == "<%=LitTablaDocV%>" || tabla == "<%=LitTablaDocC%>" || tabla == "ORDENES") {
		document.all('LitCampoCopiaCli').style.visibility="visible";
		if (tabla == "<%=LitTablaDocV%>") {
		    document.getElementById('LitCampoCopiaCli').innerHTML = "<%=LitCampoCopiaCli%>";
		}
	    if (tabla == "<%=LitTablaDocC%>") {
	        document.getElementById('LitCampoCopiaCli').innerText = "<%=LitCampoCopiaC%>";
	    }
	    if (tabla == "ORDENES") {
	        document.getElementById('LitCampoCopiaCli').innerText = "<%=LITCAMPRELORDINC%>";
	    }
	}
	else document.getElementById('LitCampoCopiaCli').style.visibility="hidden";
}

function Mas(sentido,lote,campo,criterio,texto)
{
    document.getElementById("barras").style.display="none";
    tabla=document.perso_campos.tla.value;
    session_ncliente=document.perso_campos.session_ncliente.value;
    fr_Tabla.document.perso_campos_det.action="perso_campos_det.asp?mode=ver&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&tla=" + tabla + "&session_ncliente=" + session_ncliente;
    fr_Tabla.document.perso_campos_det.submit();
}
</script>
<body bgcolor=<%=color_blau%> onload="javascript:parent.botones.page_loaded=1;" class="BODY_ASP">
<%'***********************************************************************************************************'
' CODIGO PRINCIPAL DE LA PAGINA  ***************************************************************************'
'***********************************************************************************************************'

if accesoPagina(session.sessionid,session("usuario"))=1 then%>
	<form name="perso_campos" method="post" class="col-lg-8 col-md-10 col-sm-12 col-xxs-12">
		<%Alarma "perso_campos.asp"

		mode=request("mode")

        sop=Request.QueryString("ndoc")
        
        admin=request.QueryString("admin")

        folder=Session("folder")&""
        if folder="" then 
	        folder="ilion"
        end if
        themeIlion="/lib/estilos/" & folder & "/"
    
		session_ncliente=request.querystring("ncliente")& ""

        if session_ncliente = "00000" then      
        %><script language="JavaScript" type="text/javascript">
              alert("<%=LitErrCampPersoSistGest%>");
        </script><%
        response.end
        end if
 
		if session_ncliente & ""="" then
			session_ncliente=request.querystring("session_ncliente")
		end if
		if session_ncliente & ""="" then
			session_ncliente=request.form("session_ncliente")
		end if%>
		<input type="hidden" name="session_ncliente" value="<%=enc.EncodeForHtmlAttribute(session_ncliente)%>">
		<%if session_ncliente & "">"" then
			vengo_de_sist_gestion=1
			'cadenaDSNCP=d_lookup("dsn","clientes","ncliente='" & session_ncliente & "'",dsnilion)
            strselect ="select dsn from clientes with(nolock) where ncliente=?"
            cadenaDSNCP= DLookupP1(strselect,session_ncliente&"",adVarchar,5,dsnilion)
		else
			vengo_de_sist_gestion=0
			cadenaDSNCP=session("dsn_cliente")
		end if%>
		<input type="hidden" name="vengo_de_sist_gestion" value="<%=enc.EncodeForHtmlAttribute(vengo_de_sist_gestion)%>"><%

		if vengo_de_sist_gestion=1 then
			pagina="<b>" & LitTitulo & "</b>"
			PintarCabeceraPopUp pagina
		else
			PintarCabecera "perso_campos.asp"
		end if
		'*** i AMP 09022011: Cambiamos value del desplegable por el nombre de las tablas y no el literal de idiomas.
		'*** value="NOMBRE TABLA" desc="LITERAL"
		%>
		<br>
		
        <div class="headers-wrapper">
        <%        
          
          DrawDiv "header-card", "", ""
          DrawLabel "", "", LitCamptabla%><select class="width60" name="tla" onchange="javascript:cambiar_det('<%=enc.EncodeForJavascript(sop)%>','<%=enc.EncodeForJavascript(admin)%>')">
						<option value="ARTICULOS"><%=LitTablaArt%></option>
						<option value="CLIENTES"><%=LitTablaCli%></option>
						<option value="PROVEEDORES"><%=LitTablaPro%></option>
						<option value="DOCUMENTOS VENTA"><%=LitTablaDocV%></option>
						<option value="DOCUMENTOS COMPRA"><%=LitTablaDocC%></option>
						<option value="SEGUIMIENTO COMERCIAL"><%=LitTablaComer%></option>
						<option value="CONTACTOS"><%=LitTablaContactos%></option>						
						<option value="USUARIOS WEB"><%=LitTablaUsuariosWeb%></option>
						<option value="DEVOLUCIONES CLIENTE"><%=LitTablaDevolCliente%></option>
						<option value="CENTROS"><%=LitTablaCentros%></option>
						<%''JMMM 29/01/2010 Se añade la tabla de AGENTES %>
						<option value="AGENTES"><%=LitTablaAgentes%></option>
						<%''MPC 19/08/2008 Se añade la tabla de PROYECTOS %>
						<option value="PROYECTOS"><%=LitTablaProyectos%></option>
						<%''MPC 04/12/2009 Se añade la tabla de TRAMITES %>
						<option value="TRAMITES"><%=LitTablaTramites%></option>
						<%''ZEK 04/05/2010 Se añade la tabla de TARJETAS %>
						<option value="TARJETAS"><%=LITTABLATARGETAS%></option>
						<%'AMF:22/12/2010: Se añade la tabla de INCIDENCIAS %>
						<option value="INCIDENCIAS"><%=LitTablaIncidencias%></option>
						<%'*** i AMP:28012011: Se añade la tabla de EQUIPOS %>
						<option value="EQUIPOS"><%=LitTablaEquipos%></option>
						<option value="ORDENES"><%=LITTABLAORDENES%></option>
                        <option value="COMPANYGROUP"><%=LITGROUP%></option>
						<option value="PROGRAMTYPE"><%=LITPROGRAMTYPE%></option>
                        <option value="PROGRAM"><%=LITPROGRAM%></option>
                        <option value="NETWORK"><%=LITNETWORK%></option>
                        <option value="STORE"><%=LITSTORE%></option>
                        <option value="CONTRACT"><%=LITCONTRACT%></option>
                        <option value="POS"><%=LITPOS%></option>
                        <option value="BIN"><%=LITBIN%></option>
                        <option value="CAMPAIGN"><%=LITCAMPAIGN%></option>
                        <option value="CUSTOMER_EXT"><%=LITCUSTOMER_EXT%></option>
                        <option value="PARTNER"><%=LITPARTNER%></option>
                        <option value="LOYALTY_VOUCHER_PROMOTION"><%=LITLOYALTY_VOUCHER_PROMOTION_PERSO%></option>
                        <option value="LOYALTY_VOUCHER_CATEGORY"><%=LITLOYALTY_VOUCHER_CATEGORY_PERSO%></option>
                        <option value="TRANSACTION"><%=LITTRANSACTION_PERSO%></option>
              
                        <option value="LOYALTY_PROGRAM"><%=LIT_LOYALTY_PROGRAM%></option>
                        <option value="LOYALTY_ASSOCIATION"><%=LIT_LOYALTY_ASSOCIATION%></option>
                        <option value="LOYALTY_ASSOCIATION_VERSION"><%=LIT_LOYALTY_ASSOCIATION_VERSION%></option>
                        <option value="LOYALTY_PRODUCT"><%=LIT_LOYALTY_PRODUCT%></option>
                        <option value="LOYALTY_PRODUCT_TYPE"><%=LIT_LOYALTY_PRODUCT_TYPE%></option>
					</select>
           <%CloseDiv%> 
            

		<br/>
        </div>
		<table class="width100 md-table-responsive"  CELLSPACING="1" CELLPADDING="1">
			<tr><%
				DrawceldaDet "'ENCABEZADOL width10'", "", "left", true,"<b>" & LitNCampo & "</b>"
				DrawceldaDet "'ENCABEZADOL width15'", "", "left", true,"<b>" & LitCamptitulo & "</b>"
				DrawceldaDet "'ENCABEZADOL width15'", "", "left", true,"<b>" & LitCampTipo & "</b>"
				DrawceldaDet "'ENCABEZADOL width15'", "", "left", true,"<b>" & LitCampTamany & "</b>"
                DrawceldaDet "'ENCABEZADOL width15'", "", "left", true,"<b>" & LitTituloDep & "</b>"
                
                %>
				<td class="ENCABEZADOL width15" id="LitCampoCopiaCli" style="visibility='hidden';font-weight:bold;">
					<%=LitCampoCopiaCli%>
				</td>
			    <%
                if admin=1 then
                    DrawceldaDet "'ENCABEZADOL width15'", "", "left", true,"<b>" & LitSystemReg & "</b>"
                end if
                DrawCeldaDet "'ENCABEZADOL width5'", "", "left", true,""%>
            </tr>
		</table>
		
		<iframe class="width100 iframe-data md-table-responsive" id="frtabla" name="fr_Tabla" src='perso_campos_det.asp?mode=browse&session_ncliente=<%=enc.EncodeForHtmlAttribute(session_ncliente)%>&sop=<%=enc.EncodeForHtmlAttribute(sop)%>&admin=<%=enc.EncodeForHtmlAttribute(admin)%>'' width='100%' height='230' frameborder="0" noresize="noresize" style="max-width: 100% !important;"></iframe>
		<div align="center">
			<SPAN ID="barras" STYLE="display:none">
			</SPAN>
		</div>
	</form>
<%end if%>
</BODY>
</HTML>