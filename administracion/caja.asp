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
' JCI 13/05/2003 : Al dar las facturas por cobradas/pagadas el medio de pago se ponía siempre fijo a metálico
' JCI 23/06/2003 : MIGRACION A MONOBASE
%>
<% mode = EncodeForHtml(request.QueryString("mode"))%>
<!--#include file="../adovbs.inc" -->
 <%
    
    if mode="ajax" then
        ncodigo = limpiaCadena(request.QueryString("ncodigo"))

        set rst = Server.CreateObject("ADODB.Recordset")
        set conn = Server.CreateObject("ADODB.Connection")
        set command = Server.CreateObject("ADODB.Command")

        conn.open session("dsn_cliente")
	    command.ActiveConnection =conn
	    command.CommandTimeout = 0
	    command.CommandText="getCurrency"
	    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	    command.Parameters.Append command.CreateParameter("@code",adVarChar,adParamInput,15,ncodigo)
        set rst = command.Execute()

        if not rst.eof then
            result=rst("FACTCAMBIO")
        else
            result=-1
        end if 
        set rst = nothing
        set conn = nothing
        set command = nothing       
        response.Write (result)
        response.End()

    end if
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></SCRIPT>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="Ahoja_gastos.inc" -->
<!--#include file="../perso.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../common/poner_cajaResponsive.inc" -->

<script language="javascript" type="text/javascript">
//Recarga los marcos de la caja
function VerCaja(mode) {
	if (!checkdate(document.caja.fdesde) || (document.caja.fdesde.value=="")) {
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(document.caja.fhasta) || (document.caja.fhasta.value=="")) {
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}
	if (document.caja.ncaja.value=="") {
		window.alert("<%=LitMsgCajaNoNulo%>");
		return false;
	}
		
	if (mode=="agrupado") {
	    document.getElementById("idSaldoSal").value="0";
	    document.getElementById("idActSal").value="0";
	    document.getElementById("idSaldoEntr").value="0";
	    document.getElementById("idActEntr").value="0";
        document.getElementById("totalG").innerHTML="";
        document.getElementById("totalI").innerHTML="";
        document.getElementById("saldo").innerHTML="";
        
		if (document.caja.permc.value=="SI"){
            //RGU 2/9/2009
	        //document.getElementById("i_fecha").focus();
            document.caja.i_fecha.focus();
			document.getElementById("frEntradas").src="detalles_caja_mod.asp?mode=entradas&fdesde=" + document.caja.fdesde.value +"&fhasta=" + document.caja.fhasta.value + "&ncaja=" + document.caja.ncaja.value + "&metalico=" + document.caja.metalico.checked + "&nometalico=" + document.caja.nometalico.checked;
			document.getElementById("frSalidas").src="detalles_caja_mod.asp?mode=salidas&fdesde=" + document.caja.fdesde.value +"&fhasta=" + document.caja.fhasta.value + "&ncaja=" + document.caja.ncaja.value + "&metalico=" + document.caja.metalico.checked + "&nometalico=" + document.caja.nometalico.checked;
		}
		else{
		    //RGU 2/9/2009
	        //document.getElementById("i_tanotacion").focus();
            document.caja.i_tanotacion.focus();
			document.getElementById("frEntradas").src="detalles_caja.asp?mode=entradas&fdesde=" + document.caja.fdesde.value +"&fhasta=" + document.caja.fhasta.value + "&ncaja=" + document.caja.ncaja.value + "&metalico=" + document.caja.metalico.checked + "&nometalico=" + document.caja.nometalico.checked;
			document.getElementById("frSalidas").src="detalles_caja.asp?mode=salidas&fdesde=" + document.caja.fdesde.value +"&fhasta=" + document.caja.fhasta.value + "&ncaja=" + document.caja.ncaja.value + "&metalico=" + document.caja.metalico.checked + "&nometalico=" + document.caja.nometalico.checked;
		}
	}
}

function HacerTraspaso(){
	if (document.caja.ncaja.value=="") {
		window.alert("<%=LitMsgCajaNoNulo%>");
		return false;
	}
	AbrirVentana('../search_layout.asp?pag1=administracion/traspasos_caja.asp?ndoc=' + document.caja.ncaja.value + '?viene=caja?mode=add?titulo=<%=enc.EncodeForJavascript(LitRealizarTraspCaja)%>&pag2=administracion/traspasos_caja_bt.asp&pag3=administracion/traspasos_caja_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
}

function VaciarCampo() {
	document.caja.h_docinterno.value="";
	document.caja.i_documento.value="";
}

function Insertar(mode) {
	if (document.caja.permc.value=="SI") {
		if (document.caja.i_fecha.value=="") {
			window.alert("<%=LitMsgFechaInsercionFecha%>");
			document.caja.i_fecha.focus();
			return false;
		}
		fecha=document.caja.i_fecha.value.replace("-","/");

		if (chkdatetime(fecha)==false) {
			alert("<%=LitFechaMal%>");
			document.caja.i_fecha.focus();
			return false;
		}
	}
	if (document.caja.ncaja.value=="") {
		window.alert("<%=LitMsgCajaNoNulo%>");
		document.caja.ncaja.focus();
		return false;
	}
	if (document.caja.i_divisa.value=="") {
		window.alert("<%=LitMsgDivisaNoNulo%>");
		document.caja.i_divisa.focus();
		return false;
	}
	if (document.caja.i_pago.value=="") {
		window.alert("<%=LitMsgTipoPagoNoNulo%>");
		document.caja.i_pago.focus();
		return false;
	}

	if ((document.caja.i_tapunte.value.substring(5)=="01") && (document.caja.i_documento.value=="")) {
		window.alert("<%=LitMsgDocumentoNoNulo%>");
		document.caja.i_tdocumento.focus();
		return false;
	}
	if ((document.caja.i_tdocumento.value!="") && (document.caja.i_documento.value=="")) {
		window.alert("<%=LitMsgDocumentoNoNulo%>");
		document.caja.i_tdocumento.focus();
		return false;
	}
	if ((document.caja.i_documento.value!="") && (document.caja.i_tdocumento.value=="")) {
		window.alert("<%=LitMsgTipoDocumentoNoNulo%>");
		document.caja.i_tdocumento.focus();
		return false;
	}
	if (document.caja.i_tanotacion.value=="") {
		window.alert("<%=LitMsgAnotacionNoNulo%>");
		document.caja.i_tanotacion.focus();
		return false;
	}
	if (document.caja.i_descripcion.value=="") {
		window.alert("<%=LitDescripcionNulo%>");
		document.caja.i_descripcion.focus();
		return false;
	}
	if (document.caja.i_importe.value=="") document.caja.i_importe.value=0;
	while (document.caja.i_importe.value.search(" ")!=-1) document.caja.i_importe.value=document.caja.i_importe.value.replace(" ","");
	if (document.caja.i_importe.value==""){document.caja.i_importe.value=0}
	if (isNaN(document.caja.i_importe.value.replace(",",".")))
	{
		window.alert("<%=LitMsgImporteNumerico%>");
		document.caja.i_importe.focus();
		return false;
	}
	else
	{
		if (parseFloat(document.caja.i_importe.value.replace(",","."))<=0)
		{
			window.alert("<%=LitMsgImportePositivo%>");
			document.caja.i_importe.focus();
			return false;
		}
		else document.caja.i_importe.value=document.caja.i_importe.value.replace(",",".");
	}

    // check currency change

	while (document.caja.i_changeCurrency.value.search(" ")!=-1) document.caja.i_changeCurrency.value=document.caja.i_changeCurrency.value.replace(" ","");
	if (document.caja.i_changeCurrency.value==""){document.caja.i_changeCurrency.value=1}
	if (isNaN(document.caja.i_changeCurrency.value.replace(",",".")))
	{
		window.alert("<%=LitMsgFactorCambioNumerico%>");
		document.caja.i_changeCurrency.focus();
		return false;
	}
	else
	{
	   
	        if (parseFloat(document.caja.i_changeCurrency.value.replace(",","."))<=0)
	        {
	            window.alert("<%=LitErrorFactorDeCambio%>");
	            document.caja.i_changeCurrency.focus();
	            return false;
	        }
	        else document.caja.i_changeCurrency.value=document.caja.i_changeCurrency.value.replace(",",".");
	    
	}

	//ricardo14-1-2005
	if (!checkdate(document.caja.fdesde) || (document.caja.fdesde.value==""))
	{
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		document.caja.fdesde.focus();
		return false;
	}
	if (!checkdate(document.caja.fhasta) || (document.caja.fhasta.value==""))
	{
		window.alert("<%=LitMsgHastaFechaFecha%>");
		document.caja.fhasta.focus();
		return false;
	}
	if (document.caja.ncaja.value=="")
	{
		window.alert("<%=LitMsgCajaNoNulo%>");
		document.caja.ncaja.focus();
		return false;
	}	
	//FLM:180209: se pregunta al usuario si desea continuar. Lo hago aquí para no tener que recargar la página.
	if (document.caja.i_tdocumento.value=="EFECTO CLIENTE" || document.caja.i_tdocumento.value=="EFECTO PROVEEDOR")
	{
		if(confirm("<%=LitAnotCobEfec%>")==false)
		{
		    document.caja.i_tdocumento.focus();
		    return false;
		}
	}
	////////////////////////

	document.caja.action="caja.asp?mode=save"
	document.caja.submit();
}

function BuscarDoc()
{
	if (document.caja.i_tdocumento.value!="")
	{
		pagina="../central.asp?pag1=Administracion/BuscarDoc.asp&viene=CAJA&ndoc=" + document.caja.i_documento.value +
		"&ncliente=" + document.caja.i_tdocumento.value + "&mode=search&pag2=Administracion/BuscarDoc_bt.asp&titulo=<%=enc.EncodeForJavascript(LitSelDocumento)%>";
		ven=AbrirVentana(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
	}
	else window.alert("<%=LitSeleccionarTipoDocumento%>");
}

/*
 Para los documentos de efectos, no se permite modificar el importe de la anotación en caja.
 */
function compruebaTipoDoc()
{
    if(document.caja.i_tdocumento.value=="EFECTO CLIENTE" || document.caja.i_tdocumento.value=="EFECTO PROVEEDOR")
    {
        alert("<%=LitNoModPrecioEfectos%>");
        caja.i_importe.blur();
        return;
    }
}

/*RGU: 2/9/2009*/
function FocoImporte(e)
{
    var keycode = e.keyCode;
    if (keycode==9){
        try{
            document.caja.ir_traspaso.focus();
        }
        catch(e)
        {
        }
    }
}

function FocoInsert(i)
{
    var keycode = ev.keyCode;
    if (i==1)
    {
        if (keycode==9) document.getElementById("a_BuscarDoc").focus();
    }
}

function FocoFecha()
{
    //document.all("i_tanotacion").focus()
}

function FocoTraspaso()
{
    var keycode = ev.keyCode;
    if (keycode==9) document.getElementById("a_insertar").focus();
}

function ChangeCurrency()
{
    indSel = document.caja.i_divisa.selectedIndex;
    valSel = document.caja.i_divisa[indSel].value.toString();
    defaultCurrency = document.getElementById("h_MB").value;

    if (valSel == defaultCurrency)
    {
        parent.pantalla.document.getElementById("tdLit_changeCurrency").style.display = "none";
        if (parent.pantalla.document.getElementById("td1_changeCurrency") != null)
            parent.pantalla.document.getElementById("td1_changeCurrency").style.display = "none";
        if (parent.pantalla.document.getElementById("td2_changeCurrency") != null)
            parent.pantalla.document.getElementById("td2_changeCurrency").style.display = "none";
        parent.pantalla.document.getElementById("td_changeCurrency").style.display = "none";
        parent.pantalla.document.getElementById("i_changeCurrency").value = "";
    }
    else
    {
        parent.pantalla.document.getElementById("tdLit_changeCurrency").style.display = "";
        if (parent.pantalla.document.getElementById("td1_changeCurrency") != null)
            parent.pantalla.document.getElementById("td1_changeCurrency").style.display = "";
        if (parent.pantalla.document.getElementById("td2_changeCurrency") != null)
            parent.pantalla.document.getElementById("td2_changeCurrency").style.display = "";
        parent.pantalla.document.getElementById("td_changeCurrency").style.display = "";
        parent.pantalla.document.getElementById("i_changeCurrency").value = CalculateCurrency(valSel);
    }
}
//Funcion ajax
var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest
var enProceso = false;
var result = false;


function handleHttpResponse() {
    if (http.readyState == 4) {
        if (http.status == 200) {
            if (http.responseText != "") {
                result = http.responseText;                              
            }
        }
    }
}

function CalculateCurrency(valSel) {
    var ncodigo = valSel;
    result = "";
    if (!enProceso && http) {
        var url = "caja.asp?mode=ajax&ncodigo=" + ncodigo;
        http.open("GET", url, false);
        http.onreadystatechange = handleHttpResponse;
        enProceso = false;
        http.send(null);
    }
    else
        result = 0;

    return result;
}

</script>

<body class="BODY_ASP">
<%si_tiene_modulo_ccostes=ModuloContratado(session("ncliente"),ModCcostes_Gestion) '**rgu:2/9/2009

'****************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------
'Código principal de la página
'-------------------------------------------------------------------------------------------------------------

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion

%>
<form name="caja" method="post"><%

  'Recordsets
   set rst = Server.CreateObject("ADODB.Recordset")
   set rstAux = Server.CreateObject("ADODB.Recordset")
   set rstDom = Server.CreateObject("ADODB.Recordset")

	'Leer parámetros de la página
	'mode = Request.QueryString("mode")
	if Request.QueryString("permc")>"" then
		permc=limpiaCadena(Request.QueryString("permc"))
	else
		permc=limpiaCadena(Request.Form("permc"))
	end if
	if Request.QueryString("caju")>"" then
		caju=limpiaCadena(Request.QueryString("caju"))
	else
		caju=limpiaCadena(Request.Form("caju"))
	end if%>

	<input type="hidden" name="permc" value="<%=EncodeForHtml(permc)%>"/>
	<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>"/>

	<%fdesde=limpiaCadena(request.form("fdesde"))
	fhasta=limpiaCadena(request.form("fhasta"))
	ncajaR=limpiaCadena(request.form("ncaja"))

	if mode="save" then
		FechaInsercion=limpiaCadena(request.form("i_fecha"))
		if FechaInsercion&""="" then FechaInsercion=cstr(now)
		ndocumento=limpiaCadena(request.form("i_documento"))
		ndocinterno=limpiaCadena(request.form("h_docinterno"))
		tdocumento=limpiaCadena(request.form("i_tdocumento"))
		importe=limpiaCadena(request.form("i_importe"))
		tanotacion=limpiaCadena(request.form("i_tanotacion"))
		des=replace(replace(limpiaCadena(request.form("i_descripcion")),chr(13),""),chr(10),"")
		tanotacionR=limpiaCadena(request.form("i_tanotacion"))
		pagoR=limpiaCadena(request.form("i_pago"))
		divisaR=limpiaCadena(request.form("i_divisa"))
		tapunteR=limpiaCadena(request.form("i_tapunte"))
		p_ccostes=limpiaCadena(request.form("i_ccostes")) '**rgu 3/9/2009
        changeCurrency=limpiaCadena(request.Form("i_changeCurrency"))
	end if
	if mode="ConfirmaCobro" or mode="ConfirmaPago" then
		FechaInsercion=limpiaCadena(request.form("h_fecha"))
		ndocumento=limpiaCadena(request.form("h_documento"))
		importe=limpiaCadena(request.form("h_importe"))
		descripcion=limpiaCadena(request.form("h_descripcion"))
		ncaja=limpiaCadena(request.form("h_ncaja"))
		pago=limpiaCadena(request.form("h_pago"))
		ndocinterno=limpiaCadena(request.form("h_docint"))
	    tdocumento=limpiaCadena(request.form("h_tdocumento"))
		descripcion=limpiaCadena(request.form("h_descripcion"))
		tanotacion=limpiaCadena(request.querystring("tanotacion"))
        changeCurrency=limpiaCadena(request.Form("h_changeCurrency")) 
        divisaR=limpiaCadena(request.form("h_divisaR"))
	end if

   'MB = d_lookup("codigo", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("dsn_cliente"))
    strselect = "select codigo from divisas with(nolock) where moneda_base <> ? and codigo like ? + '%'"
    MB = DLookupP2(strselect, "0", adVarchar, 1, session("ncliente"), adVarchar, 5, session("dsn_cliente"))

   %><input type="hidden" name="h_MB" id="h_MB" value="<%=EncodeForHtml(MB)%>" /><%
   'ndecimales = null_z(d_lookup("ndecimales", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("dsn_cliente")))
   strselect = "select ndecimales from divisas with(nolock) where moneda_base <> ? and codigo like ? + '%'"
   nddecimales = null_z(DLookupP2(strselect, "0", adVarchar, 1, session("ncliente"), adVarchar, 5, session("dsn_cliente")))

    'Insertar Anotación
    if mode="save" then
		'Comprobación del documento
		%><input type="hidden" name="h_fecha" value="<%=EncodeForHtml(FechaInsercion)%>"/><%
		'Comprobación del documento'
		if tdocumento="ALBARAN DE PROVEEDOR" OR tdocumento="FACTURA DE PROVEEDOR" or tdocumento="EFECTO CLIENTE" or tdocumento="EFECTO PROVEEDOR" then
		elseif tdocumento="VENCIMIENTO_ENTRADA" then
			numDocumento=ndocinterno
			ndocinterno=ndocumento
			ndocumento=numDocumento
		else
			ndocumento=session("ncliente")&ndocumento
		end if
		%><input type="hidden" name="h_documento" value="<%=EncodeForHtml(ndocumento)%>"/>
		<input type="hidden" name="h_docint" value="<%=EncodeForHtml(ndocinterno)%>"/>
		<input type="hidden" name="h_importe" value="<%=EncodeForHtml(importe)%>"/>
		<input type="hidden" name="h_tdocumento" value="<%=EncodeForHtml(tdocumento)%>"/>
		<input type="hidden" name="h_ncaja" value="<%=EncodeForHtml(ncajaR)%>"/>
		<input type="hidden" name="h_descripcion" value="<%=EncodeForHtml(des)%>"/>
		<input type="hidden" name="h_pago" value="<%=EncodeForHtml(pagoR)%>"/>
        <input type="hidden" name="h_divisaR" value="<%=EncodeForHtml(divisaR)%>"/>
        <input type="hidden" name="h_changeCurrency" value="<%=EncodeForHtml(changeCurrency)%>"/>
    <%
''response.write("los datos 0.1 son-" & ndocumento & "-" & tdocumento & "-" & importe & "-" & tanotacion & "-" & ndocinterno & "-<br>")        
		Resultado=ComprobarDocumento(ndocumento,tdocumento,importe,tanotacion,0,ndocinterno)
''response.write("los datos 0.2 son-" & Resultado & "-" & ncajaR & "-" & tanotacionR & "-" & importe & "-" & pagoR & "-" & des & "-" & ndocumento & "-" & tdocumento & "-" & divisaR & "-" & FechaInsercion & "-" & tapunteR & "-" & ndocinterno & "-" & p_ccostes & "-<br>")
		GuardarRegistro Resultado,ncajaR,"",tanotacionR,importe,pagoR,des,ndocumento,tdocumento,divisaR,FechaInsercion,0,tapunteR,ndocinterno, p_ccostes,changeCurrency
		mode="agrupado"
	elseif mode="ConfirmaCobro" then
		''tanotacion=LitEntrada
		tdocumento=LitFacCli
		PasarACobro ncaja,"",importe,tanotacion,tdocumento,descripcion,ndocumento,FechaInsercion,pago,divisaR,changeCurrency
		mode="agrupado"
	elseif mode="ConfirmaPago" then
		''tanotacion=LitSalida
		PasarAPagado ncaja,"",ndocumento,importe,tanotacion,tdocumento,descripcion,ndocinterno,FechaInsercion,pago,divisaR,changeCurrency
		mode="agrupado"
    end if
    PintarCabecera "caja.asp"
        DrawDiv "1","",""
        DrawLabel "txtMandatory","",LitDesde
        DrawInput "","","fdesde",iif(fdesde>"",EncodeForHtml(fdesde),day(date) & "/" & month(date) & "/" & year(date)),"size='12'"
        DrawCalendar "fdesde"
        CloseDiv
        DrawDiv "1","",""
        DrawLabel "txtMandatory","",LitHasta
        DrawInput "","","fhasta",iif(fhasta>"",EncodeForHtml(fhasta),day(date) & "/" & month(date) & "/" & year(date)),"size='12'"
        DrawCalendar "fhasta"
        CloseDiv
        DrawDiv "1","",""
        DrawLabel "txtMandatory","",LitCajam
            if ncajaR&"">"" then
	            defecto=ncajaR
            end if
            poner_cajasResponsive1 "width60","","ncaja","200","codigo","descripcion","","",poner_comillas(caju)
        CloseDiv
        if session("version")&"" <> "5" then
            DrawDiv "","","" 
            CloseDiv
        end if
        DrawDiv "1","",""
        DrawLabel "","",LitSolMet%><input class='CELDA7' type='checkbox' name='metalico'/><%CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LitSolNoMet%><input class='CELDA7' type='checkbox' name='nometalico'/><%CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LitSaldo%><input type="hidden" id="idSaldoEntr" name="SaldoEntr" value="0" />
                <input type="hidden" id="idActEntr" name="ActEntr" value="0" />
                <input type="hidden" id="idSaldoSal" name="SaldoSal" value="0" />
                <input type="hidden" id="idActSal" name="ActSal" value="0" />
                <div onmouseover="document.getElementById('saldoTotalPTS').style.visibility='visible';" onmouseout="document.getElementById('saldoTotalPTS').style.visibility='hidden';" id="saldoTotal" class=CELDABLUEBOLD style="background-color: transparent; border: 0px; text-align: left;"></div>
                <div id="saldoTotalPTS" class="CELDAREDBOLD" style="background-color: transparent; border: 0px; text-align: left; visibility: hidden;"></div><%CloseDiv
        DrawDiv "1","",""%><input type="Button" class='CELDA' name="ir" value="<%=LitVer%>" onclick="VerCaja('<%=enc.EncodeForJavascript(mode)%>');" onkeydown="FocoInsert(<%=iif(permc="SI",0,1)%>)"/>
                <%if permc="SI" then%>
	                <input type="Button" class='CELDA' name="ir_traspaso" value="<%=LitTraspaso%>" onclick="HacerTraspaso();" onkeydown="FocoTraspaso()"/>
                <%end if
        CloseDiv%>
    <hr/>
   <%
   alarma "caja.asp"
   
    '**rgu 3/9/2009
'si_tiene_modulo_ccostes=0
    tablewidth="835"
    tmarco_w="835"
    frmarco_w="855"
    
    if si_tiene_modulo_ccostes<>0 then
         tablewidth=""
         tmarco_w="955"
         frmarco_w="975"         
         if permc="SI" then
             tmarco_w="975"
             frmarco_w="993"
         end if
    else
         if permc="SI" then
             tmarco_w="855"
             frmarco_w="873"
         end if        
    end if

	'Linea de inserción de registros
	if permc="SI" then%>
		<table class="width90 underOrange md-table-responsive" >
            <tr class="underOrange">
				<td class='ENCABEZADOL underOrange width10'><b><%=LitFecha%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitAnotacion%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitDescripcion%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitApunte%></b></td>
				<td id="td1_changeCurrency" class='ENCABEZADOL underOrange width10' style="display:none;"></td>
				<%if si_tiene_modulo_ccostes<>0 then '**rgu 3/9/2009
				    %><td class='ENCABEZADOL underOrange width10' ><b><%=LitCCostes%></b></td><%
				else
				    %><td class='ENCABEZADOL underOrange width10'></td><%
				end if
				%><td class='ENCABEZADOL underOrange width10' ><a class="ic-accept noMTop" id="a_insertar"  href="javascript:if (Insertar('<%=enc.EncodeForJavascript(mode)%>'));" ><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=ucase(LitGuardar)%>" title="<%=ucase(LitGuardar)%>"/></a></td>
            </tr><%
			DrawFila ""
	
				set rs_Docs = server.CreateObject("ADODB.Recordset")%>
            <td class="CELDAL7 underOrange width10"><%
                DrawInput "width70","","i_fecha",cstr(now)," id='i_fecha'"
                DrawCalendar "i_fecha"%></td><%
				%><td class="CELDAL7 underOrange width10"><select class='width100' name="i_tanotacion"><option value="ENTRADA"><%=LitEntrada%></option><option value="SALIDA"><%=LitSalidaMay%></option><option value="" selected></option></select></td>
            <td class="CELDAL7 underOrange width10"><%
                DrawTextarea "width100","max-width: 300px;","i_descripcion","",""%>
            </td><%
                strselect = "SELECT * FROM Tipo_Apuntes with(nolock)  where codigo like ? + '%' order by descripcion"
                
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))

                set rs_Docs = command2.Execute
                %>
            <td class="CELDAL7 underOrange width10"><%
                DrawSelect "width100","","i_tapunte",rs_Docs,"","codigo","descripcion","",""%></td><%
				rs_Docs.Close
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
                set rs_Docs = nothing
				'**rgu 3/9/2009
                %><td class="CELDAL7 underOrange width10" id="td2_changeCurrency" style='display:none;'></td><%
				if si_tiene_modulo_ccostes<>0 then
				    strselect = "select codigo, descripcion from tiendas with(nolock) where codigo like ? + '%' order by descripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))
                    
                    set rs_Docs = command2.Execute
                    
                    %>
                <td class="CELDAL7 underOrange width10"><%
				    DrawSelect "width100","","i_ccostes",rs_Docs,"","codigo","descripcion","",""%></td><%
				    rs_Docs.Close
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
                    set rs_Docs = nothing
				else
				    %><td class="CELDAL7 underOrange width10"></td><%
				end if%>
             <td class="CELDAL7 underOrange width10"></td><%
			CloseFila%>
            <tr class="underOrange">
				<td class='ENCABEZADOL underOrange width10'><b><%=LitTipo%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitDocumento%></b><a class='ic-delete noMTop noMBottom' href="javascript:VaciarCampo();"><img src="<%=themeIlion %><%=ImgVaciarCampo%>" <%=ParamImgVaciarCampo%> alt="<%=LitVaciarCampo%>" title="<%=LitVaciarCampo%>" align="absmiddle"/></a><a href="javascript:BuscarDoc()"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscarDoc%>" title="<%=LitBuscarDoc%>"/></a></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitPago%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitDivisa%></b></td>
                <td class='ENCABEZADOL underOrange width10' id="tdLit_changeCurrency" style="display:none"><b><%=LitFactorDeCambio %></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitImporte%></b></td>
                <td class="ENCABEZADOL underOrange width10"></td>
            </tr><%
			DrawFila ""
                
                '----------------------------------------------------------------------------------------
                'Nueva forma de obtener la descripcion de los tipos de documentos de la tabla lit_typedoc
                '----------------------------------------------------------------------------------------
                set conn = Server.CreateObject("ADODB.Connection")        
                set command =  Server.CreateObject("ADODB.Command")
                conn.open DSNIlion
                command.ActiveConnection = conn
                command.CommandTimeout = 0
                command.CommandText = "ComboBoxDocTypes"
                command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                command.NamedParameters = True 
                listaIn = "'ALBARAN DE PROVEEDOR','ALBARAN DE SALIDA','EFECTO CLIENTE','EFECTO PROVEEDOR','FACTURA A CLIENTE','FACTURA DE PROVEEDOR','HOJA DE GASTOS','PEDIDO A PROVEEDOR','PEDIDO DE CLIENTE','PEDIDOS ENTRE ALMACENES','VENCIMIENTO_ENTRADA','VENCIMIENTO_SALIDA'"
                command.Parameters.Append command.CreateParameter("@inlist",adVarChar,adParamInput,len(listaIn),listaIn)
                command.Parameters.Append command.CreateParameter("@outlist",adVarChar,adParamInput,1,"")
                command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,session("usuario"))
                command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))

                set rstTDTD = Server.CreateObject("ADODB.Recordset")
                set rstTD = command.Execute
                if not rstTD.eof then%>
                <td class="CELDAL7 underOrange width10"><%
                    DrawSelect "width100","","i_tdocumento",rstTD,"","tippdoc","descripcion","onchange","javascript:VaciarCampo()"%></td><%
                end if			
                rstTD.close
                set command=nothing
                conn.close%>
                <td class="CELDAL7 underOrange width10"><%
                    DrawInput "width100","","i_documento","","readonly"%></td><%
                
                    strselect = "SELECT * FROM Tipo_pago with(nolock)  where codigo like ? + '%'"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))
                    
                    set rs_Docs = command2.Execute

				%>
                <td class="CELDAL7 underOrange width10"><%
				DrawSelect "width100", "","i_pago",rs_Docs,"","codigo","Descripcion","",""%></td><%
				    rs_Docs.Close
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
                    set rs_Docs = nothing
                    
                    strselect = "SELECT codigo,descripcion FROM divisas with(nolock) where codigo like ? + '%'"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))
                    
                    set rs_Docs = command2.Execute

				%>
                <td class="CELDAL7 underOrange width10"><%
				DrawSelect "width100", "","i_divisa",rs_Docs,MB,"codigo","descripcion","onchange","javascript:ChangeCurrency();"%></td><%
				    rs_Docs.Close
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
                    set rs_Docs = nothing
                %><td class="CELDAL7 underOrange width10" id="td_changeCurrency" style="display:none">
                    <input type="text" class="width100" name="i_changeCurrency" id="i_changeCurrency" size="10"/>
                </td>
                <td class="CELDAL7 underOrange width10"><%
				DrawInput "width50", "","i_importe","","onFocus='javascript:compruebaTipoDoc();'"%></td>
                <td class="CELDAL7 underOrange width10"></td>
				<input  type="hidden" name="h_docinterno" value=""/><%
			CloseFila
		%></table><%
	else
		%><table class="width90 underOrange md-table-responsive">
            <tr class="underOrange">
				<td class='ENCABEZADOL underOrange width10'><b><%=LitAnotacion%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitDescripcion%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitApunte%></b></td>
				<%if si_tiene_modulo_ccostes<>0 then '**rgu 3/9/2009
				    %><td class='ENCABEZADOL underOrange width10'><b><%=LitCcostes%></b></td><%
				end if%>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitTipo%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitDocumento%></b><a class='ic-delete noMTop noMBottom' href="javascript:VaciarCampo();"><img src="<%=themeIlion %><%=ImgVaciarCampo%>" <%=ParamImgVaciarCampo%> alt="<%=LitVaciarCampo%>" title="<%=LitVaciarCampo%>" align="absmiddle"/></a><a id="a_BuscarDoc" href="javascript:BuscarDoc()" ><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscarDoc%>" title="<%=LitBuscarDoc%>"/></a></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitPago%></b></td>
				<td class='ENCABEZADOL underOrange width10'><b><%=LitDivisa%></b></td>
                <td class='ENCABEZADOR underOrange width5' id="tdLit_changeCurrency" style="display:none"><b><%=LitFactorDeCambio %></b></td>
				<td class='ENCABEZADOR underOrange width5'><b><%=LitImporte%></b></td>
                <td class="ENCABEZADOL underOrange width5"></td>
            </tr>
            <tr><%
				set rs_Docs = server.CreateObject("ADODB.Recordset")
				%><td class="CELDAL7 underOrange width10"><select class='width100' name="i_tanotacion"><option value="ENTRADA"><%=LitEntrada%></option><option value="SALIDA"><%=LitSalidaMay%></option><option value="" selected></option></select></td>
            <td class="CELDAL7 underOrange width10">
                <%DrawTextarea "width100","max-width: 300px;","i_descripcion","","" %>
            </td><%
                    strselect = "SELECT * FROM Tipo_Apuntes with(nolock)  where codigo like ? + '%' order by descripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))
                    
                    set rs_Docs = command2.Execute
				%>
            <td class="CELDAL7 udnerOrange width10">
                <%DrawSelect "width100","","i_tapunte",rs_Docs,"","codigo","descripcion","",""
				    rs_Docs.Close
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
                    set rs_Docs = nothing                    
                    %>
            </td><%

				'**rgu 3/9/2009
				if si_tiene_modulo_ccostes<>0 then
                    strselect = "select codigo, descripcion from tiendas with(nolock) where codigo like ? + '%' order by descripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))
                    
                    set rs_Docs = command2.Execute
				    %>
                    <td class="CELDAL7 udnerOrange width10">
                        <%DrawSelect "width100","","i_ccostes",rs_Docs,"","codigo","descripcion","",""%>
                    </td><%
				    rs_Docs.Close
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
                    set rs_Docs = nothing   
				end if				
				
                '----------------------------------------------------------------------------------------
                'Nueva forma de obtener la descripcion de los tipos de documentos de la tabla lit_typedoc
                '----------------------------------------------------------------------------------------
                set conn = Server.CreateObject("ADODB.Connection")        
                set command =  Server.CreateObject("ADODB.Command")
                conn.open DSNIlion
                command.ActiveConnection = conn
                command.CommandTimeout = 0
                command.CommandText = "ComboBoxDocTypes"
                command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                command.NamedParameters = True 
                listaIn = "'ALBARAN DE PROVEEDOR', 'ALBARAN DE SALIDA','EFECTO CLIENTE','EFECTO PROVEEDOR','FACTURA A CLIENTE','FACTURA DE PROVEEDOR','HOJA DE GASTOS','ORDEN','PEDIDO A PROVEEDOR','PEDIDO DE CLIENTE','PEDIDOS ENTRE ALMACENES','VENCIMIENTO_ENTRADA','VENCIMIENTO_SALIDA'"
                command.Parameters.Append command.CreateParameter("@inlist",adVarChar,adParamInput,len(listaIn),listaIn)
                command.Parameters.Append command.CreateParameter("@outlist",adVarChar,adParamInput,1,"")
                command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,session("usuario"))
                command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))

                set rstTD = Server.CreateObject("ADODB.Recordset")
                set rstTD = command.Execute
                if not rstTD.eof then%>
                    <td class="CELDAL7 udnerOrange width10">
                        <%DrawSelect "width100","","i_tdocumento",rstTD,"","tippdoc","descripcion","onchange","javascript:VaciarCampo();"%>
                    </td><%
                end if			
                rstTD.close
                set command=nothing
                conn.close%>
                <td class="CELDAL7 udnerOrange width10">
                    <%DrawInput "width100","","i_documento","","readonly"%>
                </td><%
                    strselect = "SELECT * FROM Tipo_pago with(nolock) where codigo like ? + '%'"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))

                    set rs_Docs = command2.Execute

				   %>

                    <td class="CELDAL7 udnerOrange width10">
                        <%DrawSelect "width100","","i_pago",rs_Docs,"","codigo","Descripcion","",""%>
                    </td><%
				    rs_Docs.Close
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
                    set rs_Docs = nothing   

                    strselect = "SELECT codigo,descripcion FROM divisas with(nolock) where codigo like ? + '%'"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@sessionNCliente", adVarChar, adParamInput, 5, session("ncliente"))

                    set rs_Docs = command2.Execute

				%>
                    <td class="CELDAL7 udnerOrange width10">
                        <%DrawSelect "width100","","i_divisa",rs_Docs,MB,"codigo","descripcion","onchange","javascript:ChangeCurrency();"%>
                    </td><%
				    rs_Docs.Close
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
                    set rs_Docs = nothing   
                %><td class="CELDAL7 udnerOrange width5" id="td_changeCurrency" style="display:none">
                    <input type="text" class="width100" name="i_changeCurrency" id="i_changeCurrency" size="10" />
                </td>
                <td class="CELDAL7 udnerOrange width5">
                    <%DrawInput "width100","","i_importe","","onFocus='javascript:compruebaTipoDoc();'"%>
                </td>
				<td class="CELDAL7 udnerOrange width5" style="text-align:left" onkeydown="FocoInsert(1)"><a class="ic-accept noMTop" id="a_insertar" href="javascript:if (Insertar('<%=enc.EncodeForJavascript(mode)%>'));"><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=ucase(LitGuardar)%>" title="<%=ucase(LitGuardar)%>"/></a></td> <!-- onkeydown="FocoInsert()"-->
				<input type="hidden" name="h_docinterno" value=""/>
            </tr>
		  </table><%
	end if%>
	<br/><%

	if mode="agrupado" then
		%><table class="width90 md-table-responsive bCollapse"><%
				DrawCeldaDet "'ENCABEZADOL width100'", "","", true,"<b>" & LitEntradas & "</b>"
		%></table>
		<table class="width90 md-table-responsive bCollapse"><%
			Drawfila color_terra
				%><td class="ENCABEZADOL width10"><b><%=LitFecha%></b></td>
				<td class='ENCABEZADOL width15' ><b><%=LitDescripcion%></b></td>
				<td class='ENCABEZADOL width10' ><b><%=LitApunte%></b></td>
				<%if si_tiene_modulo_ccostes<>0  then %>
				    <td class='ENCABEZADOL width10' ><b><%=LitCCostes%></b></td>
				<%end if%>
				<td class='ENCABEZADOL width10' ><b><%=LitTipo%></b></td>
				<td class='ENCABEZADOL width10' ><b><%=LitDocumento%></b></td>
				<td class='ENCABEZADOL width10' ><b><%=LitPago%></b></td>
                <td class="ENCABEZADOL width10" ><b><%=LitFactorDeCambio%></b></td>
				<td class="ENCABEZADOR width5" ><b><%=LitImporte%></b></td>
				<td class="ENCABEZADOC width5" ><b><%=LitC%></b></td>
				<%if permc="SI" then%>
					<td class='ENCABEZADOL width5' ></td>
				<%end if
			CloseFila
		%></table>
		<%if permc="SI" then%>
			<iframe id="frEntradas" class="width90 iframe-data md-table-responsive" src="detalles_caja_mod.asp?mode=entradas&fdesde=<%=EncodeForHtml(fdesde)%>&fhasta=<%=EncodeForHtml(fhasta)%>&ncaja=<%=EncodeForHtml(defecto)%>" height="120">
			</iframe>
		<%else%>
			<iframe id="frEntradas" class="width90 iframe-data md-table-responsive" src="detalles_caja.asp?mode=entradas&fdesde=<%=EncodeForHtml(fdesde)%>&fhasta=<%=EncodeForHtml(fhasta)%>&ncaja=<%=EncodeForHtml(defecto)%>" height="120">
			</iframe>
		<%end if%>
		<table class="width90 md-table-responsive bCollapse">
            <tr>
                <td class="ENCABEZADOL width80" style="text-align:right"><b><%=LitTotal%></b></td>
				<td class="ENCABEZADOL width10" align="right"><div id="totalI" class="CELDABLUEBOLD" style="background-color: transparent; border: 0px; text-align: right;"></div></td>
            </tr>
		</table>
		<table class="width90 md-table-responsive bCollapse"><%
				DrawCeldaDet "'ENCABEZADOL width100'", "","", true,"<b>" & LitSalidas & "</b>"
		%></table>
		<table class="width90 md-table-responsive bCollapse"><%
			Drawfila color_terra
				%><td class='ENCABEZADOL width10' ><b><%=LitFecha%></b></td>
				<td class='ENCABEZADOL width15' ><b><%=LitDescripcion%></b></td>
				<td class='ENCABEZADOL width10' ><b><%=LitApunte%></b></td>
				<%if si_tiene_modulo_ccostes<>0  then %>
				    <td class='ENCABEZADOL width10' ><b><%=LitCCostes%></b></td>
				<%end if %>
				<td class='ENCABEZADOL width10' ><b><%=LitTipo%></b></td>
				<td class='ENCABEZADOL width10' ><b><%=LitDocumento%></b></td>
				<td class='ENCABEZADOL width10' ><b><%=LitPago%></b></td>
                <td class='ENCABEZADOL width10' ><b><%=LitFactorDeCambio%></b></td>
				<td class='ENCABEZADOR width5' ><b><%=LitImporte%></b></td>
				<td class='ENCABEZADOC width5' ><b><%=LitC%></b></td>
				<%if permc="SI" then%>
					<td class='ENCABEZADOL width5' ></td>
				<%end if
			CloseFila
		%></table>
		<%if permc="SI" then%>
			<iframe id="frSalidas" class="width90 iframe-data md-table-responsive" src="detalles_caja_mod.asp?mode=salidas&fdesde=<%=EncodeForHtml(fdesde)%>&fhasta=<%=EncodeForHtml(fhasta)%>&ncaja=<%=EncodeForHtml(defecto)%>" width="<%=frmarco_w%>" height="120">
			</iframe>
		<%else%>
			<iframe id="frSalidas" class="width90 iframe-data md-table-responsive" src="detalles_caja.asp?mode=salidas&fdesde=<%=EncodeForHtml(fdesde)%>&fhasta=<%=EncodeForHtml(fhasta)%>&ncaja=<%=EncodeForHtml(defecto)%>" width="<%=frmarco_w%>" height="120">
			</iframe>
		<%end if%>
		<table class="width90 md-table-responsive bCollapse">
		    <tr>
				<td class="ENCABEZADOL width80" style="text-align:right"><b><%=LitTotal%></b></td>
				<td class="ENCABEZADOL width10"><div id="totalG" class="CELDAREDBOLD" style="background-color: transparent; border: 0px; text-align: right;"></div></td>
			</tr>
		</table>
		<table class="width90 md-table-responsive bCollapse">
            <tr>
				<td class="ENCABEZADOL width80" style="text-align:right"><b><%=LitSaldo%></b></td>
				<td class="ENCABEZADOL width10" align="right"><div id="saldo" class="CELDABLUEBOLD10" style="background-color: transparent; border: 0px; text-align: right;"></div></td>
			</tr>
		</table>
	<%end if
	'**rgu 3/9/2009
    if request.QueryString("mode")<>"agrupado" then
	''if Resultado&""="-1" and request.QueryString("mode")="save" then
	    if permc="SI" then %> 
	        <script type="text/javascript" language="javascript">
	            window.onload = function () {
	                var counter = 0;
	                var entrado_focus2 = 0;
	                var interval1 = setInterval(function () {
	                    try {
	                        //if (document.getElementById("i_fecha") != null) {
                            if (document.caja.i_fecha!=null){
                                //document.getElementById("i_fecha").focus();
                                document.caja.i_fecha.focus();
                                //document.getElementById("i_fecha").select();
                                document.caja.i_fecha.select();
	                            entrado_focus2 = 1;
	                            counter = 16;
	                            //window.alert("adio-4");
	                        }
	                        else {
	                            //window.alert("adio-5");
	                        }
	                    }
	                    catch (e) {
	                        //window.alert("adio-6");
	                        //window.alert(e.Description);
	                    }
	                    counter++;
	                    if (counter > 15) {
	                        clearInterval(interval1);
	                    }
	                }, 75);
	            }
            </script>
	    <%else%>
	        <script type="text/javascript" language="javascript">
	            window.onload = function () {
	                var counter = 0;
	                var entrado_focus2 = 0;
	                var interval1 = setInterval(function () {
	                    try {
	                        //if (document.getElementById("i_tanotacion") != null) {
	                        if (document.caja.i_tanotacion != null) {
	                            //document.getElementById("i_tanotacion").focus();
	                            document.caja.i_tanotacion.focus();
	                            entrado_focus2 = 1;
	                            counter = 16;
	                            //window.alert("adio-1");
	                        }
	                        else {
	                            //window.alert("adio-2");
	                        }
	                    }
	                    catch (e) {
	                        //window.alert("adio-3");
	                    }
	                    counter++;
	                    if (counter > 15) {
	                        clearInterval(interval1);
	                    }
	                }, 75);
	            }
            </script>
	    <%end if
	''end if
    end if
	if request.QueryString("mode")="agrupado" then%>
	    <script type="text/javascript" language="javascript">
	        window.onload = function () {
	            var counter = 0;
	            var entrado_focus2 = 0;
	            var interval1 = setInterval(function () {
	                try {
	                    if (document.caja.fdesde != null) {
	                        document.caja.fdesde.focus();
	                        entrado_focus2 = 1;
	                        counter = 16;
	                        //window.alert("adio-7");
	                    }
	                    else {
	                        //window.alert("adio-8");
	                    }
	                }
	                catch (e) {
	                    //window.alert("adio-9");
	                }
	                counter++;
	                if (counter > 15) {
	                    clearInterval(interval1);
	                }
	            }, 75);
	        }
        </script>
	<%end if%>
	<script type="text/javascript" language="javascript">
	    if (document.caja.i_importe != null)
	    {
	        function i_importe_callblurhandler(evnt) {
                ev = (evnt) ? evnt : event;
                FocoImporte(ev);
                //window.alert("adio-10");
            }    
            if (window.document.caja.i_importe.addEventListener) {
                window.document.caja.i_importe.addEventListener("keydown",i_importe_callblurhandler , false);
            }
            else {
                window.document.caja.i_importe.attachEvent("onkeydown", i_importe_callblurhandler);
            }
        }
	</script>
</form>
<%end if

if Not(connRound IS Nothing) then
    connRound.close
    set connRound = Nothing
end if

set rst = Nothing
set rstAux = Nothing
set rstDom = Nothing
set rs_Docs = Nothing
set conn = Nothing
set rstTDTD = Nothing
set rs_Docs = Nothing
set conn = Nothing
set rstTD = Nothing
%>
</body>
</html>