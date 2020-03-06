<%@ Language=VBScript%>
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
<% ' 15/06/04 JPP  añadidos campos:numero clientes, numero proveedores, autoref, nreferencia y pagos a cta cliente'
   '*** 19/12/2005 RGU :Añadir campos para la gestion del limite alcanzado de compras por mes 
%> 
<!--#include file="../cache.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

<script language="javascript">
function ver_ExcUpdPvpArt()
{
	pagina="../central.asp?pag1=configuracion/ExcUpdPvpArt.asp&mode=browse&pag2=configuracion/ExcUpdPvpArt_bt.asp&titulo=<%=LITNOUPDATEPVPAUTO%>";
	ven=AbrirVentana(pagina,'P',<%=altoventana%>,790);
}

function gestionpvpcoste(obj) 
{
	if (obj.name=="updatepvp")
	{
		if (obj.checked) document.configuracion.updatecoste.checked=true;
	}
	else
	{
		if (!obj.checked) document.configuracion.updatepvp.checked=false;
	}
}

function ver_contra()
{
	if (document.configuracion.noriesmau.checked==true) document.getElementById("conperri").style.display="";
	else document.getElementById("conperri").style.display="none";
}

//***RGU 19/12/2005***
function ver_contralimitecomp()
{
	if (document.configuracion.noalcanzarlimimite.checked==true) document.getElementById("conlimitecomp").style.display="";
	else document.getElementById("conlimitecomp").style.display="none";
}
//***RGU 19/12/2005***

function cambio_punto()
{
	//document.configuracion.puntosporimporte.value=document.configuracion.puntosporimporte.value.replace(".",",");
}

function gest_list_portal(obj){
    if (obj.checked==false){
        document.getElementById("ver_mail").style.display="";
        document.getElementById("tdmail").style.display="";
    }
    else{
        document.getElementById("ver_mail").style.display="none";
        document.getElementById("tdmail").style.display="none";
        document.configuracion.asesoriamail.value=""
    }
}

function gest_Alerta(){
    if (document.configuracion.chkAlertaFinContrato.checked || document.configuracion.chkAltaBaja.checked ){
        document.getElementById("ver_alerta2").style.display="";
        document.getElementById("ver_alerta1").style.display="";
        document.getElementById("tdalerta2").style.display="";
        document.getElementById("tdalerta1").style.display="";
    }
    else{
        document.getElementById("ver_alerta2").style.display="none";
        document.getElementById("ver_alerta1").style.display="none";
        document.getElementById("tdalerta2").style.display="none";
        document.getElementById("tdalerta1").style.display="none";
        document.configuracion.asesoriasms.value=""
        document.configuracion.asesoriamail2.value=""
    }
}

function clickPedidos() {
    document.configuracion.elemento_MKP_03.checked = document.configuracion.elemento_MKP_02.checked;
    if (document.getElementById("elemento_MKP_04") != null)
        document.configuracion.elemento_MKP_04.checked = document.configuracion.elemento_MKP_02.checked;
    document.configuracion.elemento_MKP_05.checked = document.configuracion.elemento_MKP_02.checked;
    document.configuracion.elemento_MKP_06.checked = document.configuracion.elemento_MKP_02.checked;
    document.configuracion.elemento_MKP_07.checked = document.configuracion.elemento_MKP_02.checked;
    if(document.getElementById("elemento_MKP_08") != null)
        document.configuracion.elemento_MKP_08.checked = document.configuracion.elemento_MKP_02.checked;
    if(document.getElementById("elemento_MKP_09") != null)
        document.configuracion.elemento_MKP_09.checked = document.configuracion.elemento_MKP_02.checked;
}

function clickAdministracion() {
    document.configuracion.elemento_MKP_11.checked = document.configuracion.elemento_MKP_10.checked;
    document.configuracion.elemento_MKP_12.checked = document.configuracion.elemento_MKP_10.checked;
    document.configuracion.elemento_MKP_13.checked = document.configuracion.elemento_MKP_10.checked;
    document.configuracion.elemento_MKP_14.checked = document.configuracion.elemento_MKP_10.checked;
    document.configuracion.elemento_MKP_15.checked = document.configuracion.elemento_MKP_10.checked;
}

function clickSoporte() {
    document.configuracion.elemento_MKP_19.checked = document.configuracion.elemento_MKP_18.checked;
    document.configuracion.elemento_MKP_20.checked = document.configuracion.elemento_MKP_18.checked;
}

function clickMaintenance() {
    document.configuracion.elemento_MKP_31.checked = document.configuracion.elemento_MKP_30.checked;
}
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<TITLE><%=LitTitulo%></TITLE>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Style-Type" CONTENT="text/css">
<LINK REL="STYLESHEET" HREF="../pantalla.css" MEDIA="SCREEN">
<LINK REL="STYLESHEET" HREF="../impresora.css" MEDIA="PRINT">
<!--#include file="../ilion.inc" -->
<script language="javascript" src="../jfunciones.js"></script>
<script language="javascript">
function Ocultar()
{
	if(document.configuracion.autoref.checked==true)
	{
		document.configuracion.autoref.value="si";
		var el = document.getElementById("celdaref");
		el.style.visibility="visible";
	}
	else
	{
		document.configuracion.autoref.value="no";
		var el = document.getElementById("celdaref");
		el.style.visibility="hidden";
	}
}
var ret_ajax = "";
function getHTTPObject() {
    var xmlhttp;
    //if (!xmlhttp && XMLHttpRequest != null) {
    if (window.XMLHttpRequest) {
        try {
            xmlhttp = new XMLHttpRequest();
        }
        catch (e) { xmlhttp = false; }
    }
    return xmlhttp;
}

var enProceso = false; // lo usamos para ver si hay un proceso activo
var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest

function handleHttpResponse() {
    if (http.readyState == 4) {
        if (http.status == 200) {
            if (http.responseText.indexOf('invalid') == -1) {
                results = http.responseText;
                enProceso = false;
                ret_ajax = unescape(results);
                if (ret_ajax != "1") {
                    alert("<%=LitMsgBinUsed%>");
                } else {
                    alert("<%=LitMsgBinOk%>");
                }
                
            }
        }
    }
}
function ValidateBin() {
    bincode = document.configuracion.TGBBIN.value;
    if (bincode != document.configuracion.tgbBinOrig.value) {
        if (bincode != "" && (bincode.length == 6) && IsNumeric(bincode)) {
            ret_ajax = "";
            if (!enProceso && http) {
                var timestamp = Number(new Date());
                var url = "ValidateBin.asp?bincode=" + bincode;
                http.open("GET", url, true);
                http.onreadystatechange = handleHttpResponse;
                enProceso = false;
                http.send('');
            }
        } else {
            alert("<%=LitMsgErrBin%>");
        }
    } else {
        alert("<%=LitMsgBinOk%>");
    }

}

</script>
<BODY onload="self.status='';" bgcolor=<%=color_blau%>>
<%
'*************************************************************************************************************'
function tienePagina(pagina)
    ''ricardo 25-9-2009 como se quita la tabla accesos, se cambia el select para saber si el usuario tiene el item para esa empresa
    tienePagina=0
	if VerObjeto(pagina)=true then
		tienePagina=1
	end if
end function
'*************************************************************************************************************'

set connRound = Server.CreateObject("ADODB.Connection")
connRound.open dsnilion
if accesoPagina(session.sessionid,session("usuario"))=1 then %>
	<form name="configuracion" method="post">
	<%si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
	si_tiene_modulo_fidelizacion=ModuloContratado(session("ncliente"),ModFidelizacion)
	si_tiene_modulo_fidelizacion_premium=ModuloContratado(session("ncliente"),ModFidelizacionPremium)
	si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
	si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)
	si_tiene_paginaSMS=tienePagina(OBJMensAMoviles)
	si_tiene_modulo_ecomerce=ModuloContratado(session("ncliente"),ModEComerce)
	si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
	si_tiene_modulo_profesionales=ModuloContratado(session("ncliente"),ModProfesionales)
	si_tiene_modulo_fuerzaventas=ModuloContratado(session("ncliente"),ModFuerzaVentas)
	si_tiene_modulo_fuerzaventaspremium=ModuloContratado(session("ncliente"),ModFuerzaVentasPremium)
	'FLM:20100423
	si_tiene_modulo_fidelizacionpremium=ModuloContratado(session("ncliente"),ModFidelizacionPremium)
	si_tiene_modulo_Asesorias=ModuloContratado(session("ncliente"),ModAsesoriaAdministracion)
	'AMF:24/3/2011:Control del modulo de postventa
	si_tiene_modulo_postventa=ModuloContratado(session("ncliente"),ModPostVenta)
    'RGU 1/3/2012
    si_tiene_modulo_contabilidad=ModuloContratado(session("ncliente"),ModContabilidad)
    si_tiene_modulo_orcu=ModuloContratado(session("ncliente"),ModOrCU)
    si_tiene_modulo_TGB=ModuloContratado(session("ncliente"),ModTGB)
    si_tiene_modulo_GestContactos=ModuloContratado(session("ncliente"),ModGestorContactos)
    si_tiene_modulo_PremiumGrant=ModuloContratado(session("ncliente"),ModPremiumGrant)
    %>

	<input type="hidden" name="si_tiene_modulo_mantenimiento" value="<%=EncodeForHtml(si_tiene_modulo_mantenimiento)%>">
	<input type="hidden" name="si_tiene_modulo_fidelizacion" value="<%=EncodeForHtml(si_tiene_modulo_fidelizacion)%>">
	<input type="hidden" name="si_tiene_modulo_tiendas" value="<%=EncodeForHtml(si_tiene_modulo_tiendas)%>">
	<input type="hidden" name="si_tiene_modulo_produccion" value="<%=EncodeForHtml(si_tiene_modulo_produccion)%>">
	<input type="hidden" name="si_tiene_paginaSMS" value="<%=EncodeForHtml(si_tiene_paginaSMS)%>">
	<input type="hidden" name="si_tiene_modulo_ecomerce" value="<%=EncodeForHtml(si_tiene_modulo_ecomerce)%>">
	<input type="hidden" name="si_tiene_modulo_profesionales" value="<%=EncodeForHtml(si_tiene_modulo_profesionales)%>">
	<input type="hidden" name="si_tiene_modulo_fuerzaventas" value="<%=EncodeForHtml(si_tiene_modulo_fuerzaventas)%>">
	<input type="hidden" name="si_tiene_modulo_fuerzaventaspremium" value="<%=EncodeForHtml(si_tiene_modulo_fuerzaventaspremium)%>">
	<input type="hidden" name="si_tiene_modulo_Asesorias" value="<%=EncodeForHtml(si_tiene_modulo_Asesorias)%>">
    <input type="hidden" name="si_tiene_modulo_contabilidad" value="<%=EncodeForHtml(si_tiene_modulo_contabilidad)%>">
    <input type="hidden" name="si_tiene_modulo_orcu" value="<%=EncodeForHtml(si_tiene_modulo_orcu)%>">
    <input type="hidden" name="si_tiene_modulo_TGB" value="<%=EncodeForHtml(si_tiene_modulo_TGB)%>">
    <input type="hidden" name="si_tiene_modulo_GestContactos" value="<%=EncodeForHtml(si_tiene_modulo_GestContactos)%>">
    <input type="hidden" name="si_tiene_modulo_PremiumGrant" value="<%=EncodeForHtml(si_tiene_modulo_PremiumGrant)%>">
    <%'ASP 21/02/2012 %>
    <!--<input type="hidden" name="si_tiene_modulo_proyectos" value="<%=EncodeForHtml(si_tiene_modulo_proyectos)%>">-->
	<%'FLM:20100423 %>
    <input name="si_tiene_modulo_fidelizacionpremium" type="hidden" value="<%=EncodeForHtml(si_tiene_modulo_fidelizacionpremium)%>" />
    <%set rst = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")

	'Leer parámetros de la página'
	mode = EncodeForHtml(limpiacadena(Request.QueryString("mode")) & "")

    n_decimales=d_lookup("codigo", "divisas", "codigo like '"&session("ncliente")&"%' and moneda_base=1", session("dsn_cliente"))

	'Acción a realizar'
	if mode="save"  then
		rst.Open "select * from configuracion with(updlock) where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if limpiacadena(request.form("iva"))>"" then
			rst("iva") = limpiacadena(request.form("iva"))
		end if
		if request.form("almacen")>"" then
			rst("almacen") = limpiacadena(request.form("almacen"))
		end if
		if limpiacadena(request.form("preg_ped_bis"))="on" then
			rst("preg_ped_bis") = true
		else
   			rst("preg_ped_bis") = false
		end if
		if limpiacadena(request.form("serie"))>"" then
			rst("serie_ped_bis")=trimCodEmpresa(request.form("serie"))
		end if
		if limpiacadena(request.form("carpeta"))>"" then
			rst("path_export")=limpiacadena(request.form("carpeta"))
		end if
		if limpiacadena(request.form("preciokm"))>"" and isnumeric(request.form("preciokm"))then
			rst("preciokm")=miround(request.form("preciokm"),DEC_PREC)
		end if
		if limpiacadena(request.form("mostrarequiv"))="on" then
			rst("imp_equiv") = true
		else
   			rst("imp_equiv") = false
		end if

		if limpiacadena(request.form("autocentro"))="on" then
			rst("autocentro") = true
		else
   			rst("autocentro") = false
		end if

		if limpiacadena(request.form("autoordenes"))>"" then
			rst("autoordenes") = limpiacadena(request.form("autoordenes"))
		else
   			rst("autoordenes") = NULL
		end if

		if limpiacadena(request.form("nseriealb"))="on" then
			rst("nseriealb") = true
		else
   			rst("nseriealb") = false
		end if

        'AMF:2/2/2011:Tiempo refresco monitorizacion incidencias.
        if limpiacadena(request.form("trefrescomonitorizacioninc"))&"">"" then
			rst("trefrescomonitorizacioninc") = cint(request.form("trefrescomonitorizacioninc"))
		else
		    'AMF:2/2/2011:El tiempo por defecto es 60 segundos.
   			rst("trefrescomonitorizacioninc") = 60
		end if

		rst("ley_ticket_cabecera") = nulear(request.form("leyticketcabecera"))

		rst("ley_ticket_despedida") = nulear(request.form("leyticketdespedida"))

		rst("dec_cantidades") = null_z(request.form("deccantidades"))

		rst("dec_precios") = null_z(request.form("decprecios"))

		rst("recargo") = miround(null_z(request.form("recargo")),decpor)

		if limpiacadena(request.form("updatepvp"))="on" then
			rst("updatepvp") = true
		else
			rst("updatepvp") = false
		end if

		if limpiacadena(request.form("updatecoste"))="on" then
			rst("updatecoste") = true
		else
			rst("updatecoste") = false
		end if

		if limpiacadena(request.form("noriesmau"))="on" then
			rst("riesgooblig") = true
		else
			rst("riesgooblig") = false
		end if
		
		'***RGU 19/12/2005***
		if limpiacadena(request.form("noalcanzarlimimite"))="on" then
			rst("gestionlimitecompras") = true
		else
			rst("gestionlimitecompras") = false
		end if
		'***RGU 19/12/2005***

		if si_tiene_paginaSMS<>0 then
			if limpiacadena(request.form("mensajeria_sms"))&""<>"" then
				if nz_b(request.form("mensajeria_sms"))<>0 then
				   rst("mensajeria_sms") = true
				else
	   		   		rst("mensajeria_sms") = false
				end if
			end if
		end if

		if limpiacadena(request.form("artTodasTarifas"))="on" then
			rst("ART_TODAS_TARIFAS") = true
		else
			rst("ART_TODAS_TARIFAS") = false
		end if

		if limpiacadena(request.form("ordenEComm"))>"" then
			rst("ordcamposECommerce") = request.form("ordenEComm")
		end if

		if limpiacadena(request.form("duplicararticulos"))="on" then
			rst("art_multiempresa") = true
		else
			rst("art_multiempresa") = false
		end if

		if limpiacadena(request.form("creararttodalm"))="on" then
			rst("art_creararttodalm") = true
		else
			rst("art_creararttodalm") = false
		end if

		if limpiacadena(request.form("gen_vencimientos"))="on" then
			rst("gen_vencimientos") = true
		else
			rst("gen_vencimientos") = false
		end if

		if limpiacadena(request.form("contgencodbarras"))>"" then
			rst("NCODBARRAS") = request.form("contgencodbarras")
		else
			rst("NCODBARRAS") = request.form("contgencodbarras")
		end if


		rst("articulos_nota")=null_z(request.form("articulos_nota"))
		rst("articulos_nota_min")=null_z(request.form("articulos_nota_min"))

		rst("validez")=null_z(request.form("validez"))
		rst("anulacion")=null_z(request.form("anulacion"))
		rst("valor_ticket")=miround(null_z(request.form("valorTicket")),n_decimales)

		rst("pfinic")=null_s(request.form("pfinic"))
		rst("lon_ceros_izda")=null_z(request.form("lon_ceros_izda"))
		rst("cont_nserie")=null_z(request.form("cont_nserie"))
		rst("pffinal")=null_s(request.form("pffinal"))

		if limpiacadena(request.form("numClientes"))>"" then
			rst("ncliente") = clng(request.form("numClientes"))
		end if
		if limpiacadena(request.form("numProve"))>"" then
			rst("nproveedor") = clng(request.form("numProve"))
		end if

		if limpiacadena(request.form("autoref"))="si" then
			if limpiacadena(request.form("referencia"))>"" then
				rst("nreferencia") = clng(request.form("referencia"))
			end if
			rst("autoref") =1
		else
			rst("autoref")=0
		end if

		if limpiacadena(request.form("pagoscta"))="on" then
			rst("pagosctatickets") = true
		else
			rst("pagosctatickets") = false
		end if

		if limpiacadena(request.form("puc"))="on" then
			rst("precultcomp") = true
		else
			rst("precultcomp") = false
		end if
		if limpiacadena(request.form("pttr"))="on" then
			rst("prectartemra") = true
		else
			rst("prectartemra") = false
		end if
		if limpiacadena(request.form("ppfa"))="on" then
			rst("precpvpficart") = true
		else
			rst("precpvpficart") = false
		end if

		if limpiacadena(request.form("mostrarreftienda"))="on" then
			rst("mostrarreftienda") = true
		else
			rst("mostrarreftienda") = false
		end if

		if limpiacadena(request.form("mostrareantienda"))="on" then
			rst("mostrareantienda") = true
		else
			rst("mostrareantienda") = false
		end if

		if limpiacadena(request.form("precdtosinincluir"))="on" then
			rst("precdtonoincluido") = true
		else
			rst("precdtonoincluido") = false
		end if

		rst("contrpermries") = null_s(request.form("contrpermries"))
		'***RGU 19/12/2005***
			rst("pwdlimitecompras") = null_s(request.form("contrperlimite"))
		'***RGU 19/12/2005***

		rst("literaldoctienda1")=nulear(request.form("LDT1"))
		rst("literaldoctienda2")=nulear(request.form("LDT2"))
		
		if limpiacadena(request.form("descontarAlmVenta"))="on" then
			rst("descontarAlmVenta") = true
		else
			rst("descontarAlmVenta") = false
		end if

		'***EBF 3/2/2005
		if limpiacadena(request.form("dto_cli_art"))="on" then
			rst("dto_cli_art") = true
		else
			rst("dto_cli_art") = false
		end if
		
		if limpiacadena(request.form("avisarcambioprecio"))="on" then
			rst("avisarcambioprecio") = 1
		else
			rst("avisarcambioprecio") = false
		end if
		
		if limpiacadena(request.form("valcliente"))="on" then
			rst("valcliente") = true
		else
			rst("valcliente") = false
		end if
		
		if limpiacadena(request.form("ventabcoste"))="on" then
			rst("ventabcoste") = true
		else
			rst("ventabcoste") = false
		end if
		
		if limpiacadena(request.form("factperiodica"))="on" then
			rst("factperiodica") = true
		else
			rst("factperiodica") = false
		end if		

	
		''JMMM 17/11/09 - Configuración de proyectos
		'' Concatenación de cadena que configura la pantalla de proyectos. Ejemplo: '000011101011011101101'        	
		'rst("configproyectos")=Cstr(nz_b2(request.form("proyectoNivel"))) & CStr(nz_b2(request.form("proyectoPrioridad"))) & CStr(nz_b2(request.form("proyectoFecha"))) & _
		'                       CStr(nz_b2(request.form("proyectoDni"))) & CStr(nz_b2(request.form("proyectoDescripcion"))) & CStr(nz_b2(request.form("proyectoCoste"))) & _
		'                       CStr(nz_b2(request.form("proyectoDni_realizador"))) & CStr(nz_b2(request.form("proyectoDuracion_est"))) & CStr(nz_b2(request.form("proyectoFecha_inicio"))) & _
		'                       CStr(nz_b2(request.form("proyectoFecha_prev_fin"))) & CStr(nz_b2(request.form("proyectoDuracion_real"))) & CStr(nz_b2(request.form("proyectoFecha_fin"))) & _ 
		'                       CStr(nz_b2(request.form("proyectoFecha_prod"))) & CStr(nz_b2(request.form("proyectoVersion"))) & CStr(nz_b2(request.form("proyectoNotas"))) & _ 
		'                       CStr(nz_b2(request.form("proyectoRefa_facturar"))) & CStr(nz_b2(request.form("proyectoPorcenafacturar"))) & CStr(nz_b2(request.form("proyectoAvisar"))) & _
		'                       CStr(nz_b2(request.form("proyectoTarea_padre"))) & CStr(nz_b2(request.form("proyectoArchivo"))) & CStr(nz_b2(request.form("proyectoPorc_realizado")))	                       
		''JMMM Fin de conf. de proyectos ----------

    'JMMM 23/09/10 - Configuración de mostrar/ocultar los elementos del marketplace
		if session("ncliente") = "00112" then
		    rst("elementos_mktplace")=Cstr(nz_b2(request.form("elemento_MKP_01"))) & CStr(nz_b2(request.form("elemento_MKP_02"))) & CStr(nz_b2(request.form("elemento_MKP_03"))) & _
            CStr(nz_b2(request.form("elemento_MKP_04"))) & CStr(nz_b2(request.form("elemento_MKP_05"))) & CStr(nz_b2(request.form("elemento_MKP_06"))) & _
            CStr(nz_b2(request.form("elemento_MKP_07"))) & CStr(nz_b2(request.form("elemento_MKP_08"))) & CStr(nz_b2(request.form("elemento_MKP_09"))) & _
            CStr(nz_b2(request.form("elemento_MKP_10"))) & CStr(nz_b2(request.form("elemento_MKP_11"))) & CStr(nz_b2(request.form("elemento_MKP_12"))) & _ 
            CStr(nz_b2(request.form("elemento_MKP_13"))) & CStr(nz_b2(request.form("elemento_MKP_14"))) & CStr(nz_b2(request.form("elemento_MKP_15"))) & _ 
            CStr(nz_b2(request.form("elemento_MKP_16"))) & CStr(nz_b2(request.form("elemento_MKP_17"))) & CStr(nz_b2(request.form("elemento_MKP_18"))) & _
            CStr(nz_b2(request.form("elemento_MKP_19"))) & CStr(nz_b2(request.form("elemento_MKP_20"))) & CStr(nz_b2(request.Form("elemento_MKP_21"))) & _
            CStr(nz_b2(request.Form("elemento_MKP_22"))) & CStr(nz_b2(request.Form("elemento_MKP_23"))) & CStr(nz_b2(request.Form("elemento_MKP_24"))) & _
            "0" & "0" & "0" & "0" & "0" & CStr(nz_b2(request.Form("elemento_MKP_30"))) & CStr(nz_b2(request.Form("elemento_MKP_31")))
        else
            rst("elementos_mktplace")=Cstr(nz_b2(request.form("elemento_MKP_01"))) & CStr(nz_b2(request.form("elemento_MKP_02"))) & CStr(nz_b2(request.form("elemento_MKP_03"))) & _
            CStr(nz_b2(request.form("elemento_MKP_04"))) & CStr(nz_b2(request.form("elemento_MKP_05"))) & CStr(nz_b2(request.form("elemento_MKP_06"))) & _
            CStr(nz_b2(request.form("elemento_MKP_07"))) & CStr(nz_b2(request.form("elemento_MKP_08"))) & CStr(nz_b2(request.form("elemento_MKP_09"))) & _
            CStr(nz_b2(request.form("elemento_MKP_10"))) & CStr(nz_b2(request.form("elemento_MKP_11"))) & CStr(nz_b2(request.form("elemento_MKP_12"))) & _ 
            CStr(nz_b2(request.form("elemento_MKP_13"))) & CStr(nz_b2(request.form("elemento_MKP_14"))) & CStr(nz_b2(request.form("elemento_MKP_15"))) & _ 
            CStr(nz_b2(request.form("elemento_MKP_16"))) & CStr(nz_b2(request.form("elemento_MKP_17"))) & CStr(nz_b2(request.form("elemento_MKP_18"))) & _
            CStr(nz_b2(request.form("elemento_MKP_19"))) & CStr(nz_b2(request.form("elemento_MKP_20"))) & CStr(nz_b2(request.Form("elemento_MKP_21"))) & _
            CStr(nz_b2(request.Form("elemento_MKP_22"))) & CStr(nz_b2(request.Form("elemento_MKP_23"))) & CStr(nz_b2(request.Form("elemento_MKP_24"))) & _
            CStr(nz_b2(request.Form("elemento_MKP_25"))) & CStr(nz_b2(request.Form("elemento_MKP_26"))) & CStr(nz_b2(request.Form("elemento_MKP_27"))) & _
            CStr(nz_b2(request.Form("elemento_MKP_28"))) & CStr(nz_b2(request.Form("elemento_MKP_29"))) & CStr(nz_b2(request.Form("elemento_MKP_30"))) & _
            CStr(nz_b2(request.Form("elemento_MKP_31")))
        end if
		' Fin JMMM ----------

        ''MPC 05/02/2010 - Configuración por defecto del Seguimiento Comercial
        rst("nivelcontactodf")=nulear(request.form("nivelcontactodf"))
        rst("grupocontactodf")=nulear(request.form("grupocontactodf"))
        ''FIN MPC
        ''MPC 04/03/2010 - Configuración por defecto del Seguimiento Comercial
        rst("grupocontactoop")=nulear(request.form("grupocontactoop"))
        ''FIN MPC
        if limpiacadena(request.form("chkComSolVerSusCli"))="on" then
			rst("CADCOMSOLVERSUSCLI") = true
		else
			rst("CADCOMSOLVERSUSCLI") = false
		end if
        'FLM:20100423
        
        '

        'RGU:20100420
        if si_tiene_modulo_Asesorias<>0 then
            rst("ASESORIALIST")=nulear(nz_b2(request.form("chkMostrarAsesoria")))
            if nz_b2(request.form("chkMostrarAsesoria"))=0 then
                rst("ASESORIAMAIL")=nulear(request.form("asesoriamail"))
            end if
            rst("ASESORIAALERTA1")=nulear(nz_b2(request.form("chkAlertaFinContrato")))
            rst("ASESORIAALERTA2")=nulear(nz_b2(request.form("chkAltaBaja")))
            if nz_b2(request.form("chkAlertaFinContrato"))=1 or nz_b2(request.form("chkAltaBaja"))=1 then
                rst("ASESORIAMAIL2")=nulear(request.form("asesoriamail2"))
                rst("ASESORIASMS")=nulear(request.form("asesoriasms"))
            end if
            rst("PATHNOMINAS")=nulear(request.form("PATHNOMINAS"))
        end if

        '' MPC 17/08/2011
        if si_tiene_modulo_mantenimiento<>0 then
            rst("senderemailincidences") = limpiacadena(request.Form("senderemailincidences"))
        end if

        ''ASP 21/02/2012
        if si_tiene_modulo_proyectos <> 0 then
            rst("PROJECT_PREFIX")=nulear(request.form("project_prefix"))
            rst("PROJECT_COUNT")=nulear(request.form("project_count"))
        end if

        'RGU 1/3/2012
        rst("CONTAADDVTOSENLACE")=nz_b2(request.form("chkcontaAddVtos"))
        rst("CONTACTABANCOSUM")=nulear(request.form("ContaCtaBancoSum"))
        rst("USE_SUPLIDOS")=nz_b2(request.form("chkcontaSuplidos"))

        'RGU 14/8/2012
        rst("TGBBIN")=nulear(request.form("TGBBIN"))
        rst("TGBCAE")=nulear(request.form("TGBCAE"))
        rst("TGBCEP")=nulear(request.form("TGBCEP"))
        rst("TGBCED")=nulear(request.form("TGBCED"))

        'AMF:12/11/2012:Staff Contracts and resources booking notification.
        if si_tiene_modulo_PremiumGrant <> 0 then
            if IsNumeric(request.form("STAFFCONTRACT_NOTIFYDAYS")) then
                rst("STAFFCONTRACT_NOTIFYDAYS") = cint(request.form("STAFFCONTRACT_NOTIFYDAYS"))
            end if

            rst("STAFFCONTRACT_NOTIFYID")=nulear(request.form("STAFFCONTRACT_NOTIFYID"))

            rst("RESOURCES_NOTIFY_MAIL")=nulear(request.form("RESOURCES_NOTIFY_MAIL"))
        end if
        
        'if request.form("TGBBIN")&"">"" and request.form("tgbBinOrig") &"">"" and request.form("TGBBIN")&"" <> request.form("tgbBinOrig") then
        '    rst("TGBIDAUTO")=0
        'end if
        
		rst.update
		rst.close
		if limpiacadena(request.QueryString("act_cli"))&""="1" then
		    valor=0
            if nz_b2(request.form("chkMostrarAsesoria")) =1 then
                valor=1
            end if
            
            set conn = Server.CreateObject("ADODB.Connection")
	        set command =  Server.CreateObject("ADODB.Command")
        		
	        conn.open session("dsn_cliente")
	        command.ActiveConnection =conn
	        command.CommandTimeout = 0
	        command.CommandText="Configuracion_act_ver_listados"
	        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	        command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))
	        command.Parameters.Append command.CreateParameter("@valor",adInteger,adParamInput,,valor)
	        
	        command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
	        'on error resume next
	        command.Execute,,adExecuteNoRecords
	        resultado=command.Parameters("@p_error").Value
	        
	        conn.close
	        set command=nothing
	        set conn=nothing
	        if resultado=1 then
	            %><script> alert("<%=LitMsgErrorSaveCli %>")</script><%
	        end if
        end if
        if si_tiene_modulo_TGB<>0 then
            set connSB = Server.CreateObject("ADODB.Connection")
	        set commandSB =  Server.CreateObject("ADODB.Command")
        		
	        connSB.open DSNILION
	        commandSB.ActiveConnection =connSB
	        commandSB.CommandTimeout = 0
	        commandSB.CommandText="SaveGeneralBin"
	        commandSB.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	        commandSB.Parameters.Append commandSB.CreateParameter("@companyId",adVarChar,adParamInput,5,session("ncliente"))
	        commandSB.Parameters.Append commandSB.CreateParameter("@bin",adVarChar,adParamInput,6,request.form("TGBBIN"))
	        
	        commandSB.Parameters.Append commandSB.CreateParameter("@err",adTinyInt,adParamOutput)
	        'on error resume next
	        commandSB.Execute,,adExecuteNoRecords
	        resultado=commandSB.Parameters("@err").Value
	        connSB.close
	        set commandSB=nothing
	        set connSB=nothing
	        
        end if
        %>
		<script>
		   alert("<%=LitActualizados%>");
		</script>
	<%end if

	PintarCabecera "configuracion.asp"
	alarma "configuracion.asp"

	rst.open "select * from configuracion where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic%>
	<hr><%DrawDiv "3-sub","background-color: #eae7e3",""
                    %><label class="ENCABEZADOL" style="text-align:left"><%=LitValdef%></label><%
                    CloseDiv%>
	<!--<table width=100% border="0">
		<tr>
			<td width=35%><font CLASS=CELDA><b><%=LitValdef%></b></font></td>
		</tr>
	</table>-->

<table width="100%" bgcolor=<%=color_blau%> border = "1"><%
		rstAux.open "select tipo_iva from tipos_iva with(nolock)",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        DrawSelectCelda "width60","","","0",LitIVA,"iva",rstAux,null_s(rst("iva")),"tipo_iva","tipo_iva","",""
	    rstAux.close
	    ''ricardo 20-1-2010 no saldra el almacen para el modulo profesionales
        if si_tiene_modulo_profesionales=0 then
	        rstAux.open "select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
            DrawSelectCelda "width60","","","0",LitAlmacen,"almacen",rstAux,null_s(rst("almacen")),"codigo","descripcion","",""
		    rstAux.close
        else    
            %><input type="hidden" name="almacen" value="<%=EncodeForHtml(null_s(rst("almacen")))%>"><% 
        end if
		DrawInputCelda "CELDA","","",5,0,LitPrecioKM,"preciokm",EncodeForHtml(null_s(rst("preciokm")))
		DrawInputCelda "CELDA","","",2,0,LitDecCantidades,"deccantidades",EncodeForHtml(null_s(rst("dec_cantidades")))
		DrawInputCelda "CELDA","","",2,0,LitDecPrecios,"decprecios",EncodeForHtml(null_s(rst("dec_precios")))
		DrawInputCelda "CELDA","","",6,0,LitRecargo,"recargo",EncodeForHtml(null_s(rst("recargo")))
		DrawDiv "1","",""
        DrawLabel "","",LitActCoste%><input class="CELDA" type="checkbox" name="updatecoste" <%=iif(nz_b(rst("updatecoste"))=-1,"checked","")%> onclick="gestionpvpcoste(this)"><%CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LitActPVP%><input class="CELDA" type="checkbox" name="updatepvp" <%=iif(nz_b(rst("updatepvp"))=-1,"checked","")%> onclick="gestionpvpcoste(this)"><br /><a CLASS=CELDAREF href="javascript:ver_ExcUpdPvpArt()" onmouseover="self.status='<%=LITNOUPDATEPVPAUTO%>';return true;" onmouseout="self.status='';return true;"><%="(" & LITNOUPDATEPVPAUTO & ")"%></a><%CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LitNoPermRiesgoMaxAut%><input class="CELDA" type="checkbox" name="noriesmau" <%=iif(nz_b(rst("riesgooblig"))=-1,"checked","")%> onclick="ver_contra()"><%CloseDiv
        DrawDiv "1",iif(nz_b(rst("riesgooblig"))=-1,"","display:none"),"conperri"
        DrawLabel "","",LitContrPermRiesgo%><input class="CELDA" maxlength=12 type="password" name="contrpermries" value="<%=EncodeForHtml(null_s(rst("contrpermries")))%>"><%CloseDiv

	'***RGU 19/12/2005***
		DrawDiv "1","",""
        DrawLabel "","",LitNoPermAlcancarlimite%><input class="CELDA" type="checkbox" name="noalcanzarlimimite" <%=iif(nz_b(rst("gestionlimitecompras"))=-1,"checked","")%> onclick="ver_contralimitecomp()"><%CloseDiv
        DrawDiv "1",iif(nz_b(rst("gestionlimitecompras"))=-1,"","display:none"),"conlimitecomp"
        DrawLabel "","",LitContrLimiteCompra%><input class="CELDA" maxlength=12 type="password" name="contrperlimite" value="<%=EncodeForHtml(null_s(rst("pwdlimitecompras")))%>"><%CloseDiv

	'***RGU 19/12/2005***
		DrawCheckCelda "CELDA","","",0,LitPrecDtoSinIncluir,"precdtosinincluir",null_s(rst("precdtonoincluido"))
        DrawDiv "1","",""
        DrawLabel "","",LitAvisarCambiosPrecios%><input class="CELDA" type='checkbox' name="avisarcambioprecio" <%=iif(nz_b(rst("avisarcambioprecio"))=-1,"checked","")%>><%CloseDiv
		DrawCheckCelda "CELDA","","",0,LitAplicarDtosCliente,"dto_cli_art",null_s(rst("dto_cli_art"))
		DrawCheckCelda "CELDA","","",0,LIT_VALCLIENTE,"valcliente",null_s(rst("valcliente"))
		DrawCheckCelda "CELDA","","",0,LIT_FACTPERIODICA,"factperiodica",null_s(rst("FACTPERIODICA"))
		DrawCheckCelda "CELDA","","",0,LIT_VENTASDEBAJODELCOSTE,"ventabcoste",null_s(rst("ventabcoste"))
   
    ''ricardo 20-1-2010 no saldra el almacen para el modulo profesionales
    if si_tiene_modulo_profesionales=0 then
	    DrawCheckCelda "CELDA","","",0,LitDescontarAlmVenta,"descontarAlmVenta",null_s(rst("descontarAlmVenta"))
	else
	    ValorDescAlmV=""
		if ucase(rst("descontarAlmVenta"))="ON" or ucase(rst("descontarAlmVenta"))="VERDADERO" or ucase(rst("descontarAlmVenta"))="TRUE" then
		    ValorDescAlmV="on"
	    else
		    ValorDescAlmV=""
	    end if
	    %><input type="hidden" name="descontarAlmVenta" value="<%=EncodeForHtml(ValorDescAlmV)%>"><% 
    end if

	if si_tiene_modulo_ebesa<>1 then
			DrawCheckCelda "CELDA","","",0,LitArtTodasTarifas,"artTodasTarifas",null_s(rst("ART_TODAS_TARIFAS"))
	end if%>
</table>
<br><%DrawDiv "3-sub","background-color: #eae7e3",""
                    %><label class="ENCABEZADOL" style="text-align:left"><%=LitContadores%></label><%
                    CloseDiv%>
<!--<table width=100% border="0">
	<tr>
		<td width="50%"><font CLASS=CELDA><b><%=LitContadores%></b></font></td>
	</tr>
</table>-->
<table bgcolor=<%=color_blau%> border = "1">
	<%
		DrawInputCelda "CELDA","","",4,0,LitNumClientes,"numClientes",EncodeForHtml(null_s(rst("ncliente")))
		DrawInputCelda "CELDA","","",4,0,LitNumProve,"numProve",EncodeForHtml(null_s(rst("nproveedor")))
        DrawDiv "1","",""
        DrawLabel "","",LitAutoRef%>
		<%if rst("autoref")<>0 then %><input type="checkbox" name="autoref" value="si" checked onclick="javascript:Ocultar();">
		<%else%><input type="checkbox" name="autoref" value="no" onclick="javascript:Ocultar();">
		<%end if
        CloseDiv
        DrawDiv "1","","celdaref"
        DrawLabel "","",LitReferencia%><input type="text" name="referencia" value="<%=EncodeForHtml(null_s(rst("nreferencia")))%>" ><%CloseDiv
        %><script>
			Ocultar();
		</script>
</table>
<br>
<%if si_tiene_paginaSMS<>0 then%>
        <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitMensajeriaSMS%></label><%
        CloseDiv%>
<!--	<table width=100% border="0">
	<tr>
		<td width="50%"><font CLASS=CELDA><b><%=LitMensajeriaSMS%></b></font></td>
	</tr>
	</table>-->
	<table bgcolor=<%=color_blau%> border = "1">
		<tr>
			<input type=hidden name="mensajeria_smsbd" value="<%=EncodeForHtml(nz_b(rst("mensajeria_sms")))%>">
			<%DrawDiv "1","",""
              DrawLabel "","",LitPermitirEnvRec
              DrawCheck "'CELDA' " & iif(nz_b(rst("mensajeria_sms"))<>0,"DISABLED",""),"","mensajeria_sms",null_s(rst("mensajeria_sms"))%><br /><a CLASS=CELDAREF href="javascript:AbrirVentana('ContratoSMS.PDF','P','<%=AltoVentana%>','700')" onmouseover="self.status='<%=LitCondicionesGenerales%>';return true;" onmouseout="self.status='';return true;"><%=LitCondicionesGenerales%></a><%CloseDiv
              %>
		</tr>
	</table>
<br>
<%end if

if si_tiene_modulo_tiendas<>0 then%>
	<%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitLEYTICKETCABECERA%></label><%
        CloseDiv%>
<!--	<tr>
		<td width="40%"><font CLASS=CELDA><b><%=LitLEYTICKETCABECERA%></b></font></td>
		
	</tr>-->
	<table bgcolor=<%=color_blau%> border = "1"><%
        ley_ticket_cabecera_temp = null_s(rst("ley_ticket_cabecera"))
		DrawInputCelda "'CELDA' maxlength='50'","","",40,0,LitLEYTICKETCABECERA2,"leyticketcabecera",EncodeForHtml(ley_ticket_cabecera_temp) %>
    </table><%
        DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitLEYTICKETDESPEDIDA%></label><%
        CloseDiv%>
	<table bgcolor=<%=color_blau%> border = "1"><%
		DrawInputCelda "'CELDA' maxlength='50'","","",40,0,LitLEYTICKETDESPEDIDA2,"leyticketdespedida",EncodeForHtml(null_s(rst("ley_ticket_despedida")))
	%></table>
<%end if%>
<br>
<!--
<table width=100% border="0">
<tr>
<td width=35%><font CLASS=CELDA><b><%=LitPedbis%></b></font></td>
</tr>
</table>
<table bgcolor=<%=color_blau%> border = "1">
	<tr>
		<td>
			<table width=100% bgcolor=<%=color_blau%> border = "0"><%
            	'DrawFila color_blau
               		'Drawcelda2 "CELDA", "left", false, LitPideSerie + ": "
               		'DrawCheckCelda "CELDA","","",0,"","preg_ped_bis",rst("preg_ped_bis")
	           		'rstAux.open "select nserie, nombre from series with(nolock) where tipo_documento='PEDIDO DE CLIENTE' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
               		'Drawcelda2 "CELDA", "left", false, LitSerie + ": "
               		'codserie=session("ncliente") & rst("serie_ped_bis")
               		'DrawSelectCelda "CELDA","","","1","","serie",rstAux,codserie,"nserie","nombre","",""
			   	'rstAux.close
            	'CloseFila%>
			</table>
		</td>
	</tr>
</table>
<br>
-->
<%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitExportacion%></label><%
        CloseDiv%>
         <table width=100% bgcolor=<%=color_blau%> border = "0">
            <%DrawInputCelda "CELDA","","",50,0,LitCarpeta,"carpeta",EncodeForHtml(null_s(rst("path_export")))%>
         </table>
<br>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitMostrarEquiv%></label><%
        CloseDiv%>
        <table width=100% bgcolor=<%=color_blau%> border = "0"><%
            DrawDiv "1","",""
            DrawLabel "","",LitMostrarEquiv
		 	DrawCheck "CELDA","","mostrarequiv",null_s(rst("imp_equiv"))
			CloseDiv%>
		</table>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitGenerarVencimiento%></label><%
        CloseDiv%>
	<table><%
		DrawDiv "1","",""
        DrawLabel "","",LitGenerarVencimiento 
		DrawCheck "CELDA","","gen_vencimientos",null_s(rst("gen_vencimientos"))
        CloseDiv%>
    </table>
<br>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitMultiempresa%></label><%
        CloseDiv%>
            <table width=100% bgcolor=<%=color_blau%> border = "0">
                <%DrawDiv "1","",""
                DrawLabel "","",LitDuplicarArticulos
                DrawCheck "CELDA","","duplicararticulos",null_s(rst("art_multiempresa"))
                CloseDiv%>
            	<%''ricardo 20-1-2010 no saldra el almacen para el modulo profesionales
                if si_tiene_modulo_profesionales=0 then
		            DrawDiv "1","",""
                    DrawLabel "","",LitCrearArtEnTodAlm
                    DrawCheck "CELDA","","creararttodalm",null_s(rst("art_creararttodalm"))
                    CloseDiv
                else
	                ValorCrearTodAlm=""
		            if ucase(rst("art_creararttodalm"))="ON" or ucase(rst("art_creararttodalm"))="VERDADERO" or ucase(rst("art_creararttodalm"))="TRUE" then
		                ValorCrearTodAlm="on"
	                else
		                ValorCrearTodAlm=""
	                end if
                    %><input type="hidden" name="creararttodalm" value="<%=EncodeForHtml(ValorCrearTodAlm)%>"><% 
                end if 
			    %></table>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitContGenCodBarras%></label><%
        CloseDiv%>
	        <table width=50% bgcolor=<%=color_blau%> border = "0">
		        <%DrawDiv "1","",""
                DrawLabel "","",LitContGenCodBarrasCont%><input type="text" class="CELDA" size="10" maxlength=10 name="contgencodbarras" value="<%=EncodeForHtml(null_s(rst("NCODBARRAS")))%>" onchange="ValidarCampos()"><%CloseDiv
           	    %></table>
        <br>
<%if si_tiene_modulo_mantenimiento<>0 then%>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitMantenimiento%></label><%
        CloseDiv%>
		<table bgcolor=<%=color_blau%> border = "0">
      	  	<%DrawDiv "1","",""
            DrawLabel "","",LitAutoCentro
        	DrawCheck "CELDA","","autocentro",null_s(rst("autocentro"))
			CloseDiv
            rstAux.open "select nserie, nombre from series with(nolock) where tipo_documento='ORDEN' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
      	 	DrawSelectCelda "width60","","","1",LitAutoOrden,"autoordenes",rstAux,null_s(rst("autoordenes")),"nserie","nombre","",""
		   	rstAux.close
            DrawDiv "1","",""
            DrawLabel "","",LitObliSerieAlb
            DrawCheck "CELDA","","nseriealb",null_s(rst("nseriealb"))
			CloseDiv
            DrawInputCelda "'CELDA' maxlength=10 align='right'","","",10,0,LitTiempoRefrescoMonitorizacion,"trefrescomonitorizacioninc",EncodeForHtml(null_s(rst("trefrescomonitorizacioninc")))
		'MPC 17/08/2011 Se añade email de notificación de incidencias
            DrawInputCelda "'CELDA' maxlength=255 align='right'","","",30,0,LitSenderEmailIncidences,"senderemailincidences",EncodeForHtml(null_s(rst("senderemailincidences")))
			'FIN MPC 17/08/2011%>
	</table>
<%end if%>
<br>
<%if si_tiene_modulo_proyectos<>0 then%>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitProyectos%></label><%
        CloseDiv%>
        <table width=100% bgcolor=<%=color_blau%> border = "0">
            <%DrawInputCelda "'CELDA' maxlength=6 align='right'","","",6,0,LitProjectPrefix,"project_prefix",EncodeForHtml(null_s(rst("project_prefix")))
              DrawInputCelda "'CELDA' maxlength=4 align='right'","","",4,0,LitProjectCount,"project_count",EncodeForHtml(null_s(rst("project_count")))%>
         </table>
<%end if%>
<br>
<%if si_tiene_modulo_produccion<>0 then%>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitFabricacion%></label><%
        CloseDiv%>
	<table width=100% border="0">
      <%DrawInputCelda "'CELDA' maxlength=10 align='right'","","",10,0,LitArticulosPorNota,"articulos_nota",EncodeForHtml(null_s(rst("articulos_nota")))
        DrawInputCelda "'CELDA' maxlength=10 align='right'","","",10,0,LitArticulosPorNotaMin,"articulos_nota_min",EncodeForHtml(null_s(rst("articulos_nota_min")))
		DrawInputCelda "'CELDA' maxlength=10 align='right'","","",10,0,LitPFInicial,"pfinic",EncodeForHtml(null_s(rst("pfinic")))
		DrawInputCelda "'CELDA' maxlength=10 align='right'","","",10,0,LitPFFinal,"pffinal",EncodeForHtml(null_s(rst("pffinal")))
		DrawInputCelda "'CELDA' maxlength=10 align='right' onchange='ValidarCampos()'","","",10,0,LitContadorNSerie,"cont_nserie",EncodeForHtml(null_s(rst("cont_nserie")))
		DrawInputCelda "'CELDA' maxlength=10 align='right'","","",10,0,LitFabCerosContador,"lon_ceros_izda",EncodeForHtml(null_s(rst("lon_ceros_izda")))%>
    </table>
<%end if%>
<br>
<%if si_tiene_modulo_tiendas<>0 then%>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitTpv%></label><%
        CloseDiv%>
	<table bgcolor=<%=color_blau%> border = "1"><%
        DrawInputCelda "'CELDARIGHT' maxlength=4 align='right'","","",4,0,LitValidez & " " & LitMinutos,"validez",iif(rst("validez")&"">"",EncodeForHtml(null_s(rst("validez"))),0)
		DrawInputCelda "'CELDARIGHT' maxlength=4 align='right'","","",4,0,LitAnulacion & " " & LitMinutos,"anulacion",iif(rst("anulacion")&"">"",EncodeForHtml(null_s(rst("anulacion"))),0)
		DrawInputCelda "'CELDARIGHT' maxlength=10 align='right'","","",8,0,LitValorTicket,"valorTicket",iif(rst("valor_ticket")&"">"",EncodeForHtml(null_s(rst("valor_ticket"))),0)
		DrawDiv "1","",""
        DrawLabel "","",LitPagosTicket
        DrawCheck "CELDA","","pagoscta",null_s(rst("pagosctatickets"))
		CloseDiv%>
	</table>
<%end if%>
<br>
<%if si_tiene_modulo_ecomerce<>0 then%>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitPolPrecTieEcom%></label><%
        CloseDiv%>
	<table bgcolor=<%=color_blau%> border = "1"><%
        DrawDiv "1","",""
        DrawLabel "","",LitUltComp
		DrawCheck "CELDA","","puc",null_s(rst("precultcomp"))
        CloseDiv
		DrawDiv "1","",""
        DrawLabel "","",LitTarTempRan
		DrawCheck "CELDA","","pttr",null_s(rst("prectartemra"))
		Closediv
		DrawDiv "1","",""
        DrawLabel "","",LitPvpFicArt
        DrawCheck "CELDA","","ppfa",null_s(rst("precpvpficart"))
		CloseDiv%>
	</table>
<br>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitCamposOpcionalesEcommerce%></label><%
        CloseDiv%>
	<table bgcolor=<%=color_blau%> border="1"><%
        DrawDiv "1","",""
        DrawLabel "","",LitReferenciaEcommerce
		DrawCheck "CELDA","","mostrarreftienda",null_s(rst("mostrarreftienda"))
		CloseDiv
		DrawDiv	"1","",""
        DrawLabel "","",LitCodBarrasEcomerce
        DrawCheck "CELDA","","mostrareantienda",null_s(rst("mostrareantienda"))
		CloseDiv%>
	</table>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitCampoOrdenECommerce%></label><%
        CloseDiv%>
        <table bgcolor=<%=color_blau%> border = "0"><%
            DrawDiv "1","",""
            DrawLabel "","",LitOrdenArtECommerce%><SELECT CLASS="CELDA" name="ordenEComm">
        		<option value="referencia" <%=iif(rst("ordcamposECommerce")="referencia","selected","")%>> <%=LitOrdenPorRef%></option>
				<option value="nombre" <%=iif(rst("ordcamposECommerce")="nombre","selected","")%>><%=LitOrdenPorNombre%></option>
		    </SELECT><%CloseDiv%>
		</table>
        <br>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitDocExtra%></label><%
        CloseDiv%>
	<table bgcolor=<%=color_blau%> border="1"><%
		DrawInputCelda "'CELDA' maxlength=25","","",20,0,LitDocExtra1,"LDT1",iif(rst("literaldoctienda1")&"">"",EncodeForHtml(null_s(rst("literaldoctienda1"))),"")
		DrawInputCelda "'CELDA' maxlength=25","","",20,0,LitDocExtra2,"LDT2",iif(rst("literaldoctienda2")&"">"",EncodeForHtml(null_s(rst("literaldoctienda2"))),"")%>
	</table>
    <br>
<%if si_tiene_modulo_ecomerce<>0 then
    'DGM 28/9/11 Recogemos el último parametro para no modificarlo
    elemento_MKP_21 = mid(rst("elementos_mktplace"),21,1)
    'DGM 23/11/2012 Check for CHANGE PASSWORD in MP
    elemento_MKP_22 = mid(rst("elementos_mktplace"),22,1)
    'DGM 04/12/2012 Check for MARKET PLACE ORDER INVOCIES REPORT
    elemento_MKP_23 = mid(rst("elementos_mktplace"),23,1)
    elemento_MKP_24 = mid(rst("elementos_mktplace"),24,1)

    elemento_MKP_25 = mid(rst("elementos_mktplace"),25,1)
    elemento_MKP_26 = mid(rst("elementos_mktplace"),26,1)
    elemento_MKP_27 = mid(rst("elementos_mktplace"),27,1)
    elemento_MKP_28 = mid(rst("elementos_mktplace"),28,1)
    elemento_MKP_29 = mid(rst("elementos_mktplace"),29,1)
    ''MPC 04/10/2013 Check from MARKET PLACE MAINTENANCE
    elemento_MKP_30 = mid(rst("elementos_mktplace"),30,1)
    elemento_MKP_31 = mid(rst("elementos_mktplace"),31,1)
%>
    <input type="hidden" name="elemento_MKP_21" value="<%=EncodeForHtml(elemento_MKP_21)%>" />
    <input type="hidden" name="elemento_MKP_22" value="<%=EncodeForHtml(elemento_MKP_22)%>" />
    <input type="hidden" name="elemento_MKP_23" value="<%=EncodeForHtml(elemento_MKP_23)%>" />
    <input type="hidden" name="elemento_MKP_24" value="<%=EncodeForHtml(elemento_MKP_24)%>" />
    <input type="hidden" name="elemento_MKP_25" value="<%=EncodeForHtml(elemento_MKP_25)%>" />
    <input type="hidden" name="elemento_MKP_26" value="<%=EncodeForHtml(elemento_MKP_26)%>" />
    <input type="hidden" name="elemento_MKP_27" value="<%=EncodeForHtml(elemento_MKP_27)%>" />
    <input type="hidden" name="elemento_MKP_28" value="<%=EncodeForHtml(elemento_MKP_28)%>" />
    <input type="hidden" name="elemento_MKP_29" value="<%=EncodeForHtml(elemento_MKP_29)%>" />
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LIT_MOSTRAR_ELEMMKP%></label><%
        CloseDiv%>
        <table bgcolor=<%=color_blau%> border = "1"><%
        DrawDiv "1","",""
        DrawLabel "","",LIT_PRESUPUESTO
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_01",CBool(Mid(rst("elementos_mktplace"),1,1))
        CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LIT_PEDIDOS
        DrawCheck "'CELDA' style='text-align:right;' onclick='javascript:clickPedidos();' ","","elemento_MKP_02",CBool(Mid(rst("elementos_mktplace"),2,1))
	    CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LIT_ADMINISTRACION
        DrawCheck "'CELDA' style='text-align:right;' onclick='javascript:clickAdministracion();'","","elemento_MKP_10",CBool(Mid(rst("elementos_mktplace"),10,1))
	    CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LIT_DOCUMENTOS
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_16",CBool(Mid(rst("elementos_mktplace"),16,1))
	    CloseDiv
        if si_tiene_modulo_proyectos then
            DrawDiv "1","",""
            DrawLabel "","",LIT_TRAMITES
            DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_17",CBool(Mid(rst("elementos_mktplace"),17,1))
            CloseDiv
        end if
        if si_tiene_modulo_postventa then
	        DrawDiv "1","",""
            DrawLabel "","",LitSoporte
            DrawCheck "'CELDA' style='text-align:right;'  onclick='javascript:clickSoporte();'","","elemento_MKP_18",CBool(Mid(rst("elementos_mktplace"),18,1))
            CloseDiv
        end if
	    if si_tiene_modulo_ecomerce then
	        DrawDiv "1","",""
            DrawLabel "","",LITMAINTENANCE
            DrawCheck "'CELDA' style='text-align:right;'  onclick='javascript:clickMaintenance();'","","elemento_MKP_30",CBool(Mid(rst("elementos_mktplace"),30,1))
            Closediv
        end if
	    DrawDiv "1","",""
        DrawLabel "","",LIT_CREARPED
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_03",CBool(Mid(rst("elementos_mktplace"),3,1))
	    CloseDiv
        if session("ncliente") = "00112" then
                literal = LIT_ALBARANESPEND
            else
                literal = LIT_ALBARANESPEND_COVAL
            end if
        DrawDiv "1","",""
        DrawLabel "","",literal
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_11",CBool(Mid(rst("elementos_mktplace"),11,1))
        CloseDiv
	    if si_tiene_modulo_postventa then
	        DrawDiv "1","",""
            DrawLabel "","",LitIncidencias
            DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_19",CBool(Mid(rst("elementos_mktplace"),19,1))
            CloseDiv
        end if
	    if si_tiene_modulo_ecomerce then
	        DrawDiv "1","",""
            DrawLabel "","",LITORDERS
            DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_31",CBool(Mid(rst("elementos_mktplace"),31,1))
            CloseDiv
        end if
	    DrawDiv "1","",""
        DrawLabel "","",LIT_HISTORICOPED
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_05",CBool(Mid(rst("elementos_mktplace"),5,1))
	    CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LIT_FACTRECIBIDAS
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_12",CBool(Mid(rst("elementos_mktplace"),12,1))
	    CloseDiv
        if si_tiene_modulo_postventa then
            DrawDiv "1","",""
            DrawLabel "","",LitAnaliticaIncidencias
            DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_20",CBool(Mid(rst("elementos_mktplace"),20,1))
            CloseDiv
        end if
	    DrawDiv "1","",""
        DrawLabel "","",LIT_PEDPENDIENTES
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_06",CBool(Mid(rst("elementos_mktplace"),6,1))
	    CloseDiv
        if session("ncliente") = "00112" then
            literal = LIT_CONSUMOSMENS_COVAL
        else
            literal = LIT_CONSUMOSMENS
        end if
        DrawDiv "1","",""
        DrawLabel "","",literal
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_13",CBool(Mid(rst("elementos_mktplace"),13,1))
	    Closediv
        DrawDiv "1","",""
        DrawLabel "","",LIT_PEDRAPIDO
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_07",CBool(Mid(rst("elementos_mktplace"),7,1))
	    CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LIT_PRECIOSYTARIFAS
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_14",CBool(Mid(rst("elementos_mktplace"),14,1))
	    CloseDiv
        if session("ncliente") = "00112" then ' Sólo se muestra para Covaldroper
            DrawDiv "1","",""
            DrawLabel "","",LIT_CATALOGO
            DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_04",CBool(Mid(rst("elementos_mktplace"),4,1))
            CloseDiv
        end if
	    if session("ncliente") = "00112" then
            literal = LIT_PAGOSPEND_COVAL
        else
            literal = LIT_PAGOSPEND
        end if
        DrawDiv "1","",""
        DrawLabel "","",literal
        DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_15",CBool(Mid(rst("elementos_mktplace"),15,1))
	    CloseDiv
        if session("ncliente") = "00112" then ' Sólo se muestra para Covaldroper
            DrawDiv "1","",""
            DrawLabel "","",LIT_PEDPREDEFINIDO
            DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_08",CBool(Mid(rst("elementos_mktplace"),8,1))
            CloseDiv
        end if
	    if session("ncliente") = "00112" then ' Sólo se muestra para Covaldroper
            DrawDiv "1","",""
            DrawLabel "","",LIT_PEDREFMASPEDIDAS
            DrawCheck "'CELDA' style='text-align:right;'","","elemento_MKP_09",CBool(Mid(rst("elementos_mktplace"),9,1))
            CloseDiv
        end if%>
        </table>
<%end if
end if
    if si_tiene_modulo_fuerzaventas<>0 or si_tiene_modulo_fuerzaventaspremium<>0 or si_tiene_modulo_GestContactos<>0 then%>
    <br />
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitSeguimientoComercial%></label><%
        CloseDiv%>
        <table bgcolor="<%=color_blau%>" border = "1">
	    <%DrawDiv "1","",""
          DrawLabel "","",LitNivelDF%><select class="CELDA" name="nivelcontactodf">
			 <%for niv=0 to 5
			     if not isnull(rst("nivelcontactodf")) then
			        if cint(niv)=cint(rst("nivelcontactodf")) then
			            paso = true%>
				        <option selected value="<%=EncodeForHtml(niv)%>"><%=EncodeForHtml(niv)%></option>
				    <%else%>
				        <option value="<%=EncodeForHtml(niv)%>"><%=EncodeForHtml(niv)%></option>
				    <%end if
				else%>
				    <option value="<%=EncodeForHtml(niv)%>"><%=EncodeForHtml(niv)%></option>
				<%end if
			 next%>
			 <option <%=iif(not paso, "selected", "")%> value=""></option></select><%CloseDiv
         rstAux.cursorlocation=3
         rstAux.open "select codigo, descripcion from tipos_entidades with(nolock) where codigo like '"&session("ncliente")&"%' and tipo = 'GRUPO CONT.COMERCIAL'",session("dsn_cliente")
         DrawSelectCelda "CELDA","","","1",LitGrupoDF,"grupocontactodf",rstAux,null_s(rst("grupocontactodf")),"codigo","descripcion","",""
         rstAux.close
         rstAux.cursorlocation=3
         rstAux.open "select codigo, descripcion from tipos_entidades with(nolock) where codigo like '"&session("ncliente")&"%' and tipo = 'GRUPO CONT.COMERCIAL'",session("dsn_cliente")
         DrawSelectCelda "CELDA","","","1",LitGrupoOP,"grupocontactoop",rstAux,null_s(rst("grupocontactoop")),"codigo","descripcion","",""
         rstAux.close
         DrawDiv "1","",""
         DrawLabel "","",LITCADCOMSOLVERSUSCLI
         DrawCheck "CELDA","","chkComSolVerSusCli",null_s(rst("CADCOMSOLVERSUSCLI"))
         Closediv%>
         </table><%end if
    if si_tiene_modulo_Asesorias<>0 then%>
    <br>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitAsesoria%></label><%
        CloseDiv%>
        <table bgcolor=<%=color_blau%> border="1">
	    <%DrawDiv "1","",""
            DrawLabel "","",LitMostrarListPortal%>
            <input class="CELDA" type="checkbox" name="chkMostrarAsesoria" <%=iif(nz_b2(rst("asesorialist"))=1,"checked","")%> onclick="gest_list_portal(this)">
        <%CloseDiv%>

        <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12" id="ver_mail" style="display:<%=iif(nz_b2(rst("asesorialist"))=0,"","none")%>">
            <%DrawLabel "'CELDA' id='tdmail'",iif(nz_b2(rst("asesorialist"))=0,"","none"),LitMail%>
            <INPUT CLASS='CELDA' type='text' name='asesoriamail' value='<%=iif(nz_b2(rst("asesorialist"))=0,EncodeForHtml(null_s(rst("asesoriamail"))),"") %>' size=40>
        </div>
        <input type="hidden" name="h_mostrarasesoria" value="<%=EncodeForHtml(nz_b2(rst("asesorialist")))%>"><%
   	    
        DrawDiv "1","",""
            DrawLabel "","",LitAlertaContrato%>
            <input class="CELDA" type="checkbox" name="chkAlertaFinContrato" <%=iif(nz_b2(rst("ASESORIAALERTA1"))=1,"checked","")%> onclick="gest_Alerta()">
        <%CloseDiv%>

        <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12" id="ver_alerta1" style="display:<%=iif(nz_b2(rst("ASESORIAALERTA2"))=1 or nz_b2(rst("ASESORIAALERTA1"))=1,"","none")%>" >
            <%DrawLabel "'CELDA' id='tdalerta1'",iif(nz_b2(rst("ASESORIAALERTA2"))=1 or nz_b2(rst("ASESORIAALERTA1"))=1,"","none"),LitMail%>
            <INPUT CLASS='CELDA' type='text' name='asesoriamail2' value='<%=iif(nz_b2(rst("ASESORIAALERTA2"))=1 or nz_b2(rst("ASESORIAALERTA1"))=1,EncodeForHtml(null_s(rst("ASESORIAMAIL2"))),"") %>' size=40>
        </div><%

	    DrawDiv "1","",""
            DrawLabel "","",LitAlertaAltaBaja%>
            <input class="CELDA" type="checkbox" name="chkAltaBaja" <%=iif(nz_b2(rst("ASESORIAALERTA2"))=1,"checked","")%> onclick="gest_Alerta()">
        <%CloseDiv%>

        <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12" id="ver_alerta2" style="display:<%=iif(nz_b2(rst("ASESORIAALERTA2"))=1 or nz_b2(rst("ASESORIAALERTA1"))=1,"","none")%>" >
            <%DrawLabel "'CELDA' id='tdalerta2'",iif(nz_b2(rst("ASESORIAALERTA2"))=1 or nz_b2(rst("ASESORIAALERTA1"))=1,"","none"),LitSMS%>
            <INPUT CLASS='CELDA' type='text' name='asesoriasms' value='<%=iif(nz_b2(rst("ASESORIAALERTA2"))=1 or nz_b2(rst("ASESORIAALERTA1"))=1,EncodeForHtml(null_s(rst("ASESORIASMS"))),"") %>' size=40>
        </div><%
        
        DrawDiv "1","",""
            DrawLabel "","","PATHNOMINAPLUS"%>
            <input class="CELDA" type="text" name="PATHNOMINAS" value="<%=iif(rst("PATHNOMINAS")&"">"",EncodeForHtml(null_s(rst("PATHNOMINAS"))),"") %>" size="60">
        <%CloseDiv%>
        </table>
    <%end if%>
    <%
    'RGU 1/3/2012
    if si_tiene_modulo_contabilidad<>0 then%>
        <br>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitContabilidad%></label><%
        CloseDiv%>
        <table bgcolor=<%=color_blau%> border="1">
	        <%
            puesto_suplido=0
                DrawDiv "1","",""
                DrawLabel "","",LitContaAddvtos%>
                <input class="CELDA" type="checkbox" name="chkcontaAddVtos" <%=iif(nz_b2(rst("CONTAADDVTOSENLACE"))=1,"checked","")%> ><%
                CloseDiv
                '5/3/2012
                if si_tiene_modulo_orcu<>0 then
                    DrawInputCelda "'CELDA' maxlength=12","","",14,0,LitContaCtaBancoSum,"ContaCtaBancoSum",iif(rst("ContaCtaBancoSum")&"">"",EncodeForHtml(null_s(rst("ContaCtaBancoSum"))),"")
                else
                    puesto_suplido=1
                    DrawDiv "1","",""
                    DrawLabel "","",LITACTSUPLICONTA%>
                    <input class="CELDA" type="checkbox" name="chkcontaSuplidos" <%=iif(nz_b2(rst("USE_SUPLIDOS"))=1,"checked","")%> ><%
                    CloseDiv
                end if
            if puesto_suplido=0 then
                DrawDiv "1","",""
                DrawLabel "","",LITACTSUPLICONTA%>
                <input class="CELDA" type="checkbox" name="chkcontaSuplidos" <%=iif(nz_b2(rst("USE_SUPLIDOS"))=1,"checked","")%> ><%
                CloseDiv
            end if
            %>
        </table>
    <%end if
    
    'RGU 14/8/2012
    if si_tiene_modulo_TGB<>0 then%>
        <br>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LitTGB%></label><%
        CloseDiv%>
        <table bgcolor=<%=color_blau%> border="1" ><%
            DrawDiv "1","",""
            DrawLabel "","",LitTGBBIN%>
            <input type="hidden" name="tgbBinOrig" value="<%=EncodeForHtml(null_s(rst("TGBBIN")))&""%>" >
            <INPUT CLASS='CELDA' maxlength=6 type='text' name='TGBBIN' value='<%=iif(rst("TGBBIN")&"">"",EncodeForHtml(null_s(rst("TGBBIN"))),"")%>' size=7>
            <a class="ic-accept inlineBlock floatNone" id="checkBin" href="javascript:ValidateBin();" >
                <img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgNuevo%> >
            </a><%
            Closediv
            DrawInputCelda "'CELDA' maxlength=13","","",14,0,LitTGBCAE,"TGBCAE",iif(rst("TGBCAE")&"">"",EncodeForHtml(null_s(rst("TGBCAE"))),"")
	        DrawInputCelda "'CELDA' maxlength=4","","",5,0,LitTGBCEP,"TGBCEP",iif(rst("TGBCEP")&"">"",EncodeForHtml(null_s(rst("TGBCEP"))),"")
            DrawInputCelda "'CELDA' maxlength=4","","",5,0,LitTGBCED,"TGBCED",iif(rst("TGBCED")&"">"",EncodeForHtml(null_s(rst("TGBCED"))),"")%>
        </table>
        <br /><%
    end if
    if si_tiene_modulo_PremiumGrant<>0 then%>
        <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LITCONTRACTS%></label><%
        CloseDiv%>
	    <table bgcolor=<%=color_blau%> border="1">
		    <%DrawInputCelda "'CELDA' maxlength='4'","","",5,0,LITCONTRACTENDREMINDERDAYS,"STAFFCONTRACT_NOTIFYDAYS",iif(rst("STAFFCONTRACT_NOTIFYDAYS")&"">"",EncodeForHtml(null_s(rst("STAFFCONTRACT_NOTIFYDAYS"))),"5")
              DrawInputCelda "'CELDA' maxlength='255'","","",50,0,LITCONTRACTENDREMINDEREMAIL,"STAFFCONTRACT_NOTIFYID",iif(rst("STAFFCONTRACT_NOTIFYID")&"">"",EncodeForHtml(null_s(rst("STAFFCONTRACT_NOTIFYID"))),"")%>
	    </table>
        <%DrawDiv "3-sub","background-color: #eae7e3",""
           %><label class="ENCABEZADOL" style="text-align:left"><%=LITRESOURCESBOOKINGREQUESTS%></label><%
        CloseDiv%>
	    <table bgcolor=<%=color_blau%> border="1">
		    <%DrawInputCelda "'CELDA' maxlength='8000'","","",50,0,LITSENDRESOURCESBOOKINGNOTIFICATIONTO,"RESOURCES_NOTIFY_MAIL",iif(rst("RESOURCES_NOTIFY_MAIL")&"">"",EncodeForHtml(null_s(rst("RESOURCES_NOTIFY_MAIL"))),"")%>
	    </table>
    <br/>
    <%end if%>

</form>
<%rst.close
end if%>
</body>
</HTML>
