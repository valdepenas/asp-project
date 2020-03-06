<%@ Language=VBScript %>
<% Server.ScriptTimeout = 300 %>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<meta http-equiv="Content-style-TypeCONTENT="text/css">
<link rel="styleSHEET" href="../pantalla.css" media="SCREEN"/>
<link rel="styleSHEET" href="../impresora.css" media="PRINT"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/animatedCollapse.js.inc"-->
<!--#include file="../js/tabs.js.inc"-->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="cierres_caja.inc" -->

<!--#include file="../perso.inc" -->
<!--#include file="../common/poner_cajaResponsive.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/generalData.css.inc"-->
<!--#include file="../styles/Section.css.inc"-->
<!--#include file="../styles/ExtraLink.css.inc"-->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
function tier1Menu(objMenu2, objImage2)
{
    objMenu = document.getElementById(objMenu2);
    objImage = document.getElementById(objImage2);
    if (objMenu != null)
    {
        if (objImage != null)
        {
            if (objMenu.style.display == "none")
            {
                objMenu.style.display = "";
                objImage.src = "../Images/<%=ImgCarpetaAbierta%>";
            }
            else
            {
                objMenu.style.display = "none";
                objImage.src = "../images/<%=ImgCarpetaCerrada%>";
            }
        }
    }
}

//***************************************************************************
function Mas(sentido, lote, campo, criterio, texto)
{
	document.location="cierres_caja.asp?mode=search&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto;
}

//***************************************************************************
function Editar(cierre)
{
	document.cierres_caja.action="cierres_caja.asp?ncierre=" + cierre + "&mode=browse";
	parent.botones.document.location = "cierres_caja_bt.asp?mode=browse";
    document.cierres_caja.submit();
	
}

//***************************************************************************
function Permisos(mode)
{
    /*ricardo 16-5-2011 se quita de aqui, porque si la pagina de botones se abre mas tarde que la pantalla, dara error javascript*/
    /*
	if (mode=="add") {
		a=0;
		while (parent.botones.document.getElementById("BotGuardar")==null && a<10000) a=a+1;
		if (document.cierres_caja.dni.value!="") parent.botones.document.getElementById("BotGuardar").style.visibility="visible";
		else parent.botones.document.getElementById("BotGuardar").style.visibility="hidden";
	}
	if (mode=="search" || mode=="browse") {
		a=0;
		while (parent.botones.document.getElementById("BotAnadir")==null && a<10000) a=a+1;
		if (document.cierres_caja.dni.value!="") parent.botones.document.getElementById("BotAnadir").style.visibility="visible";
		else parent.botones.document.getElementById("BotAnadir").style.visibility="hidden";
	}
	*/
}

//***************************************************************************
function CalculaDescuadre(sm, desc)
{
	saldometalico=document.getElementById("h_saldometalico").value;
	saldometalico=saldometalico.replace(",",".");
	metalicoreal=document.cierres_caja.metalicoreal.value.replace(",",".");
	if (!isNaN(metalicoreal)) {
		descuadre=parseFloat(metalicoreal) - parseFloat(saldometalico);
		document.cierres_caja.hdescuadre.value=descuadre.toFixed(2);
		document.getElementById("celdadescuadre").innerHTML=descuadre.toFixed(2);
		document.cierres_caja.metalicoreal.value=metalicoreal;
	}
	else {
		alert("<%=LitMsgFormatoNumericoIncorrecto%>");
		document.cierres_caja.hdescuadre.value=0;
		document.cierres_caja.metalicoreal.value=sm;
		document.getElementById("celdadescuadre").innerHTML=desc;
	}
}

//***************************************************************************
function CalculaSaldoCierre(salida,sc) {
	saldototal=document.getElementById("h_saldoinforme").value.replace(",",".");
	saldoapertura=document.cierres_caja.hsaldoapertura.value.replace(",",".");
	if (document.cierres_caja.salidanometalico.checked==true) saldonometalico=document.cierres_caja.hsaldonometalico.value.replace(",",".");
	else saldonometalico=0;
	descuadre=document.getElementById("celdadescuadre").innerHTML;
	descuadre=descuadre.replace(",",".");
	salidacaja=document.cierres_caja.salidacaja.value.replace(",",".");
	if (!isNaN(salidacaja)) {
		saldocierre=parseFloat(saldoapertura) + parseFloat(saldototal) + parseFloat(descuadre) - parseFloat(saldonometalico) - parseFloat(salidacaja);
		if (saldocierre.toString().search("e-")!=-1) saldocierre=0;
		document.cierres_caja.hsaldocierre.value=saldocierre.toFixed(2);
		document.getElementById("celdasaldocierre").innerHTML=saldocierre.toFixed(2);
		document.cierres_caja.salidacaja.value=salidacaja;
	}
	else {
		alert("<%=LitMsgFormatoNumericoIncorrecto%>");
		document.cierres_caja.hsaldocierre.value=sc;
		document.cierres_caja.salidacaja.value=salida;
		document.cierres_caja.salidanometalico.checked=true;
		document.getElementById("celdasaldocierre").innerHTML=parseFloat(sc.replace(",",".")).toFixed(2);
	}
}
</script>

<%mode = Request.QueryString("mode")%>

<body onload="javascript:Permisos('<%=enc.EncodeForJavascript(mode)%>');" class="BODY_ASP">
<%function CadenaBusqueda(campo,criterio,texto)
	condSerie=""
	if bus="TI" then
        strselect = "select tienda from cajas where codigo=?"
        tienda=DLookupP1(strselect,session("f_caja")&"",adVarChar,10,session("dsn_cliente"))
		condSerie=" and CA.codigo in (select codigo from cajas with(nolock) where tienda='" & tienda & "') "
	elseif bus<>"" then
		condSerie=" and CA.codigo='" & bus & "' "
	end if
	if texto>"" then
		strcond=condSerie & " and CI.divisa=D.codigo and CI.operador=P.dni and CI.caja=CA.codigo and CI.codigo like '" & session("ncliente") & "%'"
		select case criterio
			case "contiene"
				CadenaBusqueda=" where " & campo & " like '%" & texto & "%'" & strcond & " order by CI.codigo desc"
			case "empieza"
				CadenaBusqueda=" where " & campo & " like '" & texto & "%'" & strcond & " order by CI.codigo desc"
			case "termina"
				CadenaBusqueda=" where " & campo & " like '%" & texto + "'" & strcond & " order by CI.codigo desc"
			case "igual"
				CadenaBusqueda=" where " & campo & "='" & texto & "'" & strcond & " order by CI.codigo desc"
		end select
	else
		CadenaBusqueda=" where CI.divisa=D.codigo and CI.operador=P.dni and CI.caja=CA.codigo and CI.codigo like '" & session("ncliente") & "%' " & condSerie & " order by CI.codigo desc"
	end if
end function

'********************************************************************************

'Botones de navegación para las búsquedas.
sub NextPrev(lote,lotes,campo,criterio,texto,pos)%>
<table width='100%' border='0' cellspacing="1" cellpadding="1">
	<tr><td class='MAS'><%
	   lote=cint(lote)
	   lotes=cint(lotes)
	    varias=false
		if lote>1 then
			%><a class='CELDAREF' href="javascript:Mas('prev',<%=enc.EncodeForHtmlAttribute(lote)%>,'<%=enc.EncodeForHtmlAttribute(campo)%>','<%=enc.EncodeForHtmlAttribute(criterio)%>','<%=enc.EncodeForHtmlAttribute(texto)%>');">
			<img src="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a><%
			varias=true
		end if
		textopag=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)
		%><font class='CELDA'> <%=textopag%> </font> <%

		if lote<lotes then
			%><a class='CELDAREF' href="javascript:Mas('next',<%=enc.EncodeForHtmlAttribute(lote)%>,'<%=enc.EncodeForHtmlAttribute(campo)%>','<%=enc.EncodeForHtmlAttribute(criterio)%>','<%=enc.EncodeForHtmlAttribute(texto)%>');">
			<img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a><%
			varias=true
		end if
		if varias=true then
	  	  %><font class='CELDA'>&nbsp;&nbsp; <%=LitPagIrA%> <input class='CELDA' type="text" name="SaltoPagina<%=enc.EncodeForHtmlAttribute(pos)%>" size="2">&nbsp;&nbsp;<a class='CELDAREF' href="javascript:IrAPagina(<%=enc.EncodeForJavascript(pos)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>',<%=rst.PageCount%>,'lote');"><%=LitIr%></a></font><%
	  end if
	%></td></tr>
</table>
<%end sub

'*********************************************************************************************************
'Se pintan los datos de la cabecera del cierre
sub CabeceraCierre(modo)
	if modo="first_save" then%>
		<br/>
        <%DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitInformePrevio%></b></label><%
        CloseDiv
        strselect = "select descripcion from cajas with(nolock) where codigo=?"
		EligeCeldaResponsive "text","browse","CELDA","","",0,"","<b>" & LitCaja & "</b>",LitCaja, DLookupP1(strselect,caja&"",adVarChar,10,session("dsn_cliente"))
        if session("version")&"" <> "5" then
                DrawDiv "","","" 
                CloseDiv
        end if 
        EligeCeldaResponsive "text","browse","CELDA","","",0,"","<b>" & LitDesdeFecha & "</b>",LitCaja, dfecha

        EligeCeldaResponsive "text","browse","CELDA","","",0,"","<b>" & LitHastaFecha & "</b>",LitCaja, hfecha%>
		<br/>
	<%elseif mode="browse" then%>
		<input type="hidden" name="ncierre" value="<%=enc.EncodeForHtmlAttribute(ncierre)%>">
		<%'ega 16/06/2008 union con join y with(nolock)
		strselect="select CC.SALIDATODONOMETALICO,CC.saldo,CC.saldoapertura,CC.saldometalico,CC.realmetalico,CC.salidacierre,CC.saldocierre,CC.codigo,CC.divisa,D.ndecimales,D.abreviatura,CC.caja,C.descripcion as Nombrecaja,CC.fecha,CC.operador,P.nombre as NombreOperador,CC.desde,CC.hasta "
		strselect=strselect & " from cierres_caja CC with(nolock) join divisas D with(nolock) on CC.divisa=D.codigo join cajas C with(nolock) on CC.caja=C.codigo join personal P with(nolock) on CC.operador=P.dni where CC.codigo='" & ncierre & "' "
		rst.open strselect,session("dsn_cliente")
		if not rst.eof then
			rs_ncierre=trimCodEmpresa(rst("codigo"))
			rs_fecha=rst("fecha")
			rs_abreviatura=rst("abreviatura")
			AbreviaturaMB=rs_abreviatura
			n_decimales=rst("ndecimales")
			rs_caja=rst("caja")
			rs_NombreCaja=rst("NombreCaja")
			rs_operador=rst("operador")
			rs_NombreOperador=rst("NombreOperador")
			rs_desde=rst("desde")
			rs_hasta=rst("hasta")
			rs_saldoapertura=rst("saldoapertura")
			rs_saldometalico=rst("saldometalico")
			rs_realmetalico=rst("realmetalico")
			rs_salidacierre=rst("salidacierre")
			rs_saldocierre=rst("saldocierre")
			rs_saldoinforme=rst("saldo")
			rs_salidatodonometalico=rst("SALIDATODONOMETALICO")
			rs_descuadre=rs_realmetalico-rs_saldometalico
		end if
		rst.close
        strselect = "select top 1 codigo from cierres_caja with(nolock) where caja=? and codigo like ?+'%' order by codigo desc"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@caja",adVarChar,adParamInput,10,rs_caja)
        command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente"))
        set rst = command2.Execute
		if not rst.eof then
			rs_ultimocierre=trimCodEmpresa(rst("codigo"))
			if rs_ultimocierre=rs_ncierre then
				%><input type="hidden" name="ultimocierre" value="1"><%
			else
				%><input type="hidden" name="ultimocierre" value="0"><%
			end if
		end if
		conn2.Close
        set conn2 = nothing
        set command2 = nothing
        set rst = nothing%>
		<br/>
		<table width="100%">
		    <td align="right" width="33%">
		    	<%a=formato_impresion()%>
		    </td>
		</table><%
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "", "<b>" & ucase(LitFecha) & "</b>",ucase(LitFecha), enc.EncodeForHtmlAttribute(null_s(rs_fecha))
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & ucase(LitCierre) & "</b>", ucase(LitCierre), enc.EncodeForHtmlAttribute(null_s(rs_ncierre))
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & ucase(LitDivisa) & "</b>", ucase(LitDivisa), enc.EncodeForHtmlAttribute(null_s(rs_abreviatura))

            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & ucase(LitOperador) & "</b>", ucase(LitOperador), enc.EncodeForHtmlAttribute(null_s(rs_NombreOperador))
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & ucase(LitCaja) & "</b>", ucase(LitCaja), enc.EncodeForHtmlAttribute(null_s(rs_NombreCaja))

            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & ucase(LitDesde) & "</b>", ucase(LitDesde), enc.EncodeForHtmlAttribute(null_s(rs_desde))
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & ucase(LitHasta) & "</b>", ucase(LitHasta), enc.EncodeForHtmlAttribute(null_s(formatdatetime(rs_hasta,vbShortDate)))

            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & LitSaldoApertura & "</b>", LitSaldoApertura, formatnumber(rs_saldoapertura,n_decimales,-1,0,-1)
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & LitSaldoCierre & "</b>", LitSaldoCierre, formatnumber(rs_saldocierre,n_decimales,-1,0,-1)

            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & LitSaldoMetalico & "</b>", LitSaldoMetalico, formatnumber(rs_saldometalico,n_decimales,-1,0,-1)
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & LitMetalicoReal & "</b>", LitMetalicoReal , formatnumber(rs_realmetalico,n_decimales,-1,0,-1)

            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & LitDescuadre & "</b>", LitDescuadre, formatnumber(rs_descuadre,n_decimales,-1,0,-1)
            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & LitSalidaMetalico & "</b>", LitSalidaMetalico, formatnumber(rs_salidacierre,n_decimales,-1,0,-1)

            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "","<b>" & LitSalidaTodoNoMetalico & "</b>", LitSerie, enc.EncodeForHtmlAttribute(null_s(Visualizar(rs_salidatodonometalico)))%>
    <br/><%
    end if
end sub

'******************************************************************************
'Crea la tabla que contiene la barra de grupos de datos.
sub BarraNavegacion(modo)%>
        <ul>
			<%if modo="first_save" then%>
                <li><a href="#DATCIERRE"><%=LitConfirmacionCierre%></a></li>
                <li><a href="#MOVCAJA"><%=LitMovimientosCaja%></a></li><%
			else%>
                <li><a href="#MOVCAJA"><%=LitMovimientosCaja%></a></li><%
			end if%>
            <li><a href="#TICKEMIT"><%=LitTicketsEmitidos%></a></li>
            <li><a href="#VTPAGO"><%=LitVentasTipoPago%></a></li>
            <li><a href="#VOPERADORES"><%=LitVentasOperadores%></a></li>
            <li><a href="#VTPV"><%=LitVentasTpv%></a></li>
        </ul>
    <%
end sub

'*********************************************************************************************************
'Se pintan los datos de la cabecera del cierre
sub DibujaSpans()%>
		<div ID="MOVCAJA" class="overflowXauto">
			<%DibujaMovCaja%>
		</div>
		<div ID="TICKEMIT" class="overflowXauto" >
			<%DibujaTicketsEmit%>
		</div>
		<div ID="VTPAGO" class="overflowXauto" >
			<%DibujaVentasTipoPago%>
		</div>
		<div ID="VTARTICULO" class="overflowXauto" >
			<%'DibujaVentasTipoArticulo%>
		</div>
		<div ID="VOPERADORES" class="overflowXauto" >
			<%DibujaVentasOperadores%>
		</div>
		<div ID="VTPV" class="overflowXauto" >
			<%DibujaVentasTpv%>
		</div>
<%end sub

'*********************************************************************************************************
'Se pintan los datos del informe previo
sub DibujaSpansPrevios()%>
		<div ID="DATCIERRE" class="overflowXauto" >
			<%DatosConfirmacionCierre%>
		</div>
		<div ID="MOVCAJA" class="overflowXauto" >
			<%DibujaMovCajaPrevio%>
		</div>
		<div ID="TICKEMIT" class="overflowXauto" >
			<%DibujaTicketsEmitPrevio%>
		</div>
		<div ID="VTPAGO" class="overflowXauto" >
			<%DibujaVentasTipoPagoPrevio%>
		</div>
		<div ID="VTARTICULO" class="overflowXauto" >
			<%'DibujaVentasTipoArticuloPrevio%>
		</div>
		<div ID="VOPERADORES" class="overflowXauto" >
			<%DibujaVentasOperadoresPrevio%>
		</div>
		<div ID="VTPV" class="overflowXauto" >
			<%DibujaVentasTpvPrevio%>
		</div>
<%end sub

'******************************************************************************
'Se pintan los datos de confirmación del cierre
sub DatosConfirmacionCierre()
	set conn = Server.CreateObject("ADODB.Connection")
	set command =  Server.CreateObject("ADODB.Command")

    ''*****************************************************
    ''ricardo 4-2-2008 se cambia el el usuario del dsn_cliente por el de dsn_import, ya que si la empresa tiene acceso a fidelizacion
    ''da un error de acceso a ilion_admin
    	initial_catalogC=encontrar_datos_dsn(session("dsn_cliente"),"Initial Catalog=")

		donde=inStr(1,DSNImport,"Initial Catalog=",1)
		donde_fin=InStr(donde,DSNImport,";",1)
		if donde_fin=0 then
			donde_fin=len(DSNImport)
		end if
		cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))

		dsnCliente=cadena_dsn_final
		conn.open dsnCliente

	''*****************************************************

	command.ActiveConnection =conn
	command.CommandTimeout = 0
	command.CommandText="DatosPreviosCierre"
	command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
	command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
	command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
	command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
	command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
	command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
	command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)
	command.Parameters.Append command.CreateParameter("@p_saldo_metalico",adCurrency,adParamOutput)
	command.Parameters.Append command.CreateParameter("@p_saldo_total",adCurrency,adParamOutput)
	command.Parameters.Append command.CreateParameter("@p_saldo_apertura",adCurrency,adParamOutput)
	command.Parameters.Append command.CreateParameter("@p_saldo_nometalico",adCurrency,adParamOutput)
	command.Parameters.Append command.CreateParameter("@p_saldo_metalico_cierre_ant",adCurrency,adParamOutput)
	command.Parameters.Append command.CreateParameter("@p_NumTickets",adInteger,adParamOutput)
	command.Parameters.Append command.CreateParameter("@p_DesdeTicket",adVarChar,adParamOutput,20)
	command.Parameters.Append command.CreateParameter("@p_HastaTicket",adVarChar,adParamOutput,20)

	'on error resume next
	command.Execute,,adExecuteNoRecords
	'on error goto 0
	resultado=command.Parameters("@p_error").Value
	saldometalico=command.Parameters("@p_saldo_metalico").Value
	saldototal=command.Parameters("@p_saldo_total").Value
	saldoapertura=command.Parameters("@p_saldo_apertura").Value
	saldoNOmetalico=command.Parameters("@p_saldo_nometalico").Value
	saldometalicocierreant=command.Parameters("@p_saldo_metalico_cierre_ant").Value
	NumTickets=command.Parameters("@p_NumTickets").Value
	DesdeTicket=command.Parameters("@p_DesdeTicket").Value
	HastaTicket=command.Parameters("@p_HastaTicket").Value

	conn.close
	set command=nothing
	set conn=nothing

	saldometalicoBCK=saldometalico
	saldometalico=saldometalico+saldometalicocierreant%>
    <%DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitConfirmacionCierre%></b></label><%
        CloseDiv
    EligeCeldaResponsive "text","browse","CELDA","","",0,"","<b>" & LitSaldoApertura & "</b>", LitSaldoApertura,  formatnumber(saldoapertura,n_decimales,-1,0,-1) & " (" & formatnumber(saldometalicocierreant,n_decimales,-1,0,-1) & " " & LITDEMETALICO & ")"
	%><input type="hidden" name="hsaldoapertura" value="<%=enc.EncodeForHtmlAttribute(saldoapertura)%>"><%

    EligeCeldaResponsive "text","browse","CELDA","","",0,"","<b>" & LitSaldoInforme & "</b>", "", formatnumber(saldototal,n_decimales,-1,0,-1)
	%><input type="hidden" name="hsaldoinforme" id="h_saldoinforme" value="<%=enc.EncodeForHtmlAttribute(saldototal)%>"><%
	
    EligeCeldaResponsive "text","browse","CELDA","","",0,"","<b>" & LitSaldoMetalico & "</b>", "",formatnumber(saldometalico,n_decimales,-1,0,-1) & " (" & formatnumber(saldometalicoBCK,n_decimales,-1,0,-1) & " " & LITDEINFORME & " + " & formatnumber(saldometalicocierreant,n_decimales,-1,0,-1) & " " & LITDESALDOAPER & ")"
    %><input type="hidden" name="hsaldometalico" id="h_saldometalico" value="<%=enc.EncodeForHtmlAttribute(saldometalico)%>"><%
	
    DrawDiv "1","",""
    DrawLabel "","","<b>" & LitMetalicoReal & "</b>"%><input type="Text" class="CELDAR7" size="13" name="metalicoreal" onchange="CalculaDescuadre('<%=enc.EncodeForJavascript(saldometalico)%>','0');CalculaSaldoCierre('0','<%=saldototal+saldoapertura-saldoNOmetalico%>');" value="<%=enc.EncodeForHtmlAttribute(saldometalico)%>"><%
    CloseDiv
	
    DrawDiv "1","",""
    DrawLabel "","","<b>" & LitDescuadre & "</b>"
    DrawSpan "CELDA","", formatnumber(0,n_decimales,-1,0,-1), "id='celdadescuadre'"%><input type="hidden" name="hdescuadre" value="0"><%
    CloseDiv
    
    strselect = "SELECT * FROM Tipo_pago with(nolock) where codigo like ?+'%'"
    set command2 = nothing
    set conn2 = Server.CreateObject("ADODB.Connection")
    set command2 = Server.CreateObject("ADODB.Command")
    conn2.Open = session("dsn_cliente")
    conn2.CursorLocation = 3
    command2.ActiveConnection = conn2
    command2.CommandTimeout = 60
    command2.CommandText = strselect
    command2.CommandType = adCmdText
    command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,8,session("ncliente"))
    set rst = command2.Execute

    DrawDiv "1","",""
    DrawLabel "","","<b>" & LitSalidacaja & "</b>"%><input type="Text" class="width30" size="13" name="salidacaja" onchange="CalculaSaldoCierre('0','<%=saldototal+saldoapertura-saldoNOmetalico%>');" value="0"><%
    DrawSelect "width30 disabled","margin-left:2px;","i_pago",rst,session("ncliente") & "01","codigo","Descripcion","",""
    CloseDiv

    conn2.Close
    set conn2 = nothing
    set command2 = nothing
    set rst = nothing

    DrawDiv "1","",""
    DrawLabel "","","<b>" & LitTodasEntradasNoMetalico & "</b>"%><input class='CELDA7' type='checkbox' name='salidanometalico' checked onclick="CalculaSaldoCierre('0','<%=saldototal+saldoapertura-saldoNOmetalico%>');"><%DrawSpan "CELDA","margin-left:2px",formatnumber(saldoNOmetalico,n_decimales,-1,0,-1),""%><input type="hidden" name="hsaldonometalico" value="<%=enc.EncodeForHtmlAttribute(saldoNOmetalico)%>"><%
    CloseDiv

    DrawDiv "1","",""
    DrawLabel "","","<b>" & LitFechaSalida & "</b>"%><input type="Text" class="CELDAR7" size="13" name="fechasalida" value="<%=enc.EncodeForHtmlAttribute(hfecha)%>"><%
    DrawCalendar "fechasalida"
	CloseDiv

    DrawDiv "1","",""
    DrawLabel "","","<b>" & LitSaldoCierre & "</b>"
    DrawSpan "CELDA","", formatnumber(saldototal+saldoapertura-saldoNOmetalico,n_decimales,-1,0,-1), "id='celdasaldocierre'"%><input type="hidden" name="hsaldocierre" value="<%=saldototal+saldoapertura-saldoNOmetalico%>"><%
    CloseDiv
end sub

'******************************************************************************
'Se pintan los datos de los movimientos de caja
sub DibujaMovCaja()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitMovimientosCaja%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "left", false,"<b>" & LitTipoPago & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitEntradas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitSalidas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitSaldo & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		'ega 16/06/2008 union de tablas con join y with(nolock)
		strselect="select C.tpago,T.descripcion as NombreTpago,(C.ventas+C.entradas) as entradas,(C.anulaciones+C.salidas) as salidas,(C.ventas+C.entradas)-(C.anulaciones+C.salidas) as saldo "
		strselect=strselect & " from cierres_tpago C with(nolock) inner join tipo_pago T with(nolock) on C.tpago=T.codigo where C.cierre='" & ncierre & "' order by NombreTpago"
		rst.open strselect,session("dsn_cliente")
		SumEntradas=0
		SumSalidas=0
		SumSaldo=0
		if not rst.eof then
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTPago")
					EntradasCierre=d_sum("importe","caja","cierre='" & ncierre & "' and medio='" & rst("tpago") & "' and tanotacion='ENTRADA' and descripcion='SALIDA CIERRE DE CAJA Nº : " & right(ncierre,len(ncierre)-5) & "'",session("dsn_cliente"))
					DescuadresCierreEnt=d_sum("importe","caja","cierre='" & ncierre & "' and medio='" & rst("tpago") & "' and tanotacion='ENTRADA' and descripcion='DESCUADRE CIERRE DE CAJA Nº : " & right(ncierre,len(ncierre)-5) & "'",session("dsn_cliente"))
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatNumber(rst("entradas")-DescuadresCierreEnt-EntradasCierre,n_decimales,-1,0,-1)
					SumEntradas=SumEntradas + rst("entradas")-DescuadresCierreEnt-EntradasCierre
					SalidasCierre=d_sum("importe","caja","cierre='" & ncierre & "' and medio='" & rst("tpago") & "' and tanotacion='SALIDA' and descripcion='SALIDA CIERRE DE CAJA Nº : " & right(ncierre,len(ncierre)-5) & "'",session("dsn_cliente"))
					DescuadresCierre=d_sum("importe","caja","cierre='" & ncierre & "' and medio='" & rst("tpago") & "' and tanotacion='SALIDA' and descripcion='DESCUADRE CIERRE DE CAJA Nº : " & right(ncierre,len(ncierre)-5) & "'",session("dsn_cliente"))
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatNumber(rst("salidas")-SalidasCierre-DescuadresCierre,n_decimales,-1,0,-1)
					SumSalidas=SumSalidas + rst("salidas") - SalidasCierre - DescuadresCierre
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatNumber((rst("entradas")- DescuadresCierreEnt-EntradasCierre) -(rst("salidas")-SalidasCierre-DescuadresCierre),n_decimales,-1,0,-1)
					SumSaldo=SumSaldo + (rst("entradas")- DescuadresCierreEnt-EntradasCierre) - (rst("salidas")-SalidasCierre-DescuadresCierre)
					'DrawCelda2 "CELDAR7", "right", false, formatNumber(rst("saldo"),n_decimales,-1,0,-1)
					'SumSaldo=SumSaldo + rst("saldo")
				CloseFila
				rst.movenext
			wend
		end if
		rst.close
		'Totales
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL7 width40'", "left", false,"<b>" & LitTotales & "</b>"
			DrawCelda2 "'ENCABEZADOR7 width20'", "right", false,"<b>" & formatNumber(SumEntradas,n_decimales,-1,0,-1) & "</b>"
			DrawCelda2 "'ENCABEZADOR7 width20'", "right", false,"<b>" & formatNumber(SumSalidas,n_decimales,-1,0,-1) & "</b>"
			DrawCelda2 "'ENCABEZADOR7 width20'", "right", false,"<b>" & formatNumber(SumSaldo,n_decimales,-1,0,-1) & "</b>"
		CloseFila
	%></table><%
end sub 

'******************************************************************************
'Se pintan los datos de los movimientos de caja del informe previo
sub DibujaMovCajaPrevio()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitMovimientosCaja%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "left", false,"<b>" & LitTipoPago & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" &  LitEntradas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" &  LitSalidas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" &  LitSaldo & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila

		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")

		conn.open session("dsn_cliente")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="MovCajaPrevios"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
		command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
		command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
		command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
		command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
		command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
		command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)
		set rst=command.Execute
		resultado=command.Parameters("@p_error").Value

		SumEntradas=0
		SumSalidas=0
		SumSaldo=0
		if not rst.eof then
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTPago")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatNumber(rst("entradas"),n_decimales,-1,0,-1)
					SumEntradas=SumEntradas + rst("entradas")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatNumber(rst("salidas"),n_decimales,-1,0,-1)
					SumSalidas=SumSalidas + rst("salidas")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatNumber(rst("saldo"),n_decimales,-1,0,-1)
					SumSaldo=SumSaldo + rst("saldo")
				CloseFila
				rst.movenext
			wend
		end if
		rst.close
		conn.close
		set command=nothing
		set conn=nothing
		'Totales
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL7 width40'", "left", false,"<b>" & LitTotales & "</b>"
			DrawCelda2 "'ENCABEZADOR7 width20'", "right", false,"<b>" & formatNumber(SumEntradas,n_decimales,-1,0,-1) & "</b>"
			DrawCelda2 "'ENCABEZADOR7 width20'", "right", false,"<b>" & formatNumber(SumSalidas,n_decimales,-1,0,-1) & "</b>"
			DrawCelda2 "'ENCABEZADOR7 width20'", "right", false,"<b>" & formatNumber(SumSaldo,n_decimales,-1,0,-1) & "</b>"
		CloseFila
	%></table><%
end sub

'******************************************************************************
'Se pintan los datos de los tickets emitidos
sub DibujaTicketsEmit()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitTicketsEmitidos%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		'ega 16/06/2008 union con join y with(nolock)
		strselect="select C.serie,S.nombre as NombreSerie,C.tipoiva,C.base_imponible,C.iva,C.ventas,C.tickets,C.desde,C.hasta "
		strselect=strselect & " from cierres_serie C with(nolock) inner join series S with(nolock) on C.serie=S.nserie where C.cierre='" & ncierre & "' order by serie,tipoiva"
		rst.open strselect,session("dsn_cliente")
		if not rst.eof then
			SerieAnt=""
			SumBi=0
			SumIva=0
			SumTotal=0
			TSumTotal=0
			while not rst.eof
				if rst("serie")<>SerieAnt then
					'Total Serie Anterior
					if SerieAnt<>"" then
						DrawFila color_terra
							DrawCelda2 "'ENCABEZADOL width25'", "left", false,""
				            DrawCelda2 "'ENCABEZADOL width25'", "left", false,"<b>" & LitTotalserie & "</b>"
							DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumBi,n_decimales,-1,0,-1) & "</b>"
							DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumIva,n_decimales,-1,0,-1) & "</b>"
						CloseFila
						SumBi=0
						SumIva=0
						SumTotal=0
					end if
					'Cabecera Serie
					DrawFila color_blau
						DrawCelda2 "'ENCABEZADOL width25' bgcolor=" & color_terra, "left", false,"<b>" & Litserie & "</b>"
						DrawCelda2 "'CELDAL7 width25'", "left", false,rst("NombreSerie")
						DrawCelda2 "'ENCABEZADOL width25' colspan='2' bgcolor=" & color_terra, "left", false,"<b>" & LitNtickets & "</b>"
						DrawCelda2 "'CELDAL7 width25' colspan='2'", "left", false,rst("tickets")
					CloseFila
					DrawFila color_blau
						DrawCelda2 "'CELDAL7 width25'", "left", false,"&nbsp;"
						DrawCelda2 "'CELDAL7 width25'", "left", false,"&nbsp;"
						DrawCelda2 "'ENCABEZADOL width15'", "left", false,"<b>" & LitDesde & "</b>"
						DrawCelda2 "'CELDAL7 width10'", "left", false,trimCodEmpresa(rst("desde"))
						DrawCelda2 "'ENCABEZADOL width15'", "left", false,"<b>" & LitHasta & "</b>"
						DrawCelda2 "'CELDAL7 width10'", "left", false,trimCodEmpresa(rst("hasta"))
					CloseFila
					'Cabecera Tipos de Iva
					DrawFila color_blau
                        DrawCelda2 "'CELDAL7 width25'", "left", false,""
						DrawCelda2 "'ENCABEZADOL width25'", "left", false,"<b>" & LitTipoIva & "</b>"
						DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & LitBaseImponible & "</b>"
						DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & LitIva & "</b>"
					CloseFila
				end if
				DrawFila color_blau
                    DrawCelda2 "'CELDAL7 width25'", "left", false,""
					DrawCelda2 "'CELDAL7 width25'", "left", false, rst("tipoiva")
					DrawCelda2 "'CELDAL7 width25' colspan='2'", "left", false,formatnumber(rst("base_imponible"),n_decimales,-1,0,-1)
					SumBi=SumBi + rst("base_imponible")
					DrawCelda2 "'CELDAL7 width25' colspan='2'", "left", false, formatnumber(rst("iva"),n_decimales,-1,0,-1)
					SumIva=SumIva + rst("iva")
					SumTotal=SumTotal + rst("ventas")
					TSumTotal=TSumTotal + rst("ventas")
				CloseFila
				SerieAnt=rst("serie")
				rst.movenext
			wend
			'Total última serie
			DrawFila color_terra
                DrawCelda2 "'ENCABEZADOL width25'", "left", false,""
				DrawCelda2 "'ENCABEZADOL width25'", "left", false,"<b>" & LitTotalserie & "</b>"
				DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumBi,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumIva,n_decimales,-1,0,-1) & "</b>"
			CloseFila
			SumBi=0
			SumIva=0
			SumTotal=0
			'Total General
			'DrawFila color_terra
			'	DrawCelda2Span "CELDA7", "left", true, LitTotalTicketsEmitidos & " " & AbreviaturaMB,3
			'	DrawCelda2Span "CELDARIGHT", "right", true,formatnumber(TSumTotal,n_decimales,-1,0,-1),3
			'CloseFila
		end if
		rst.close
	%></table><%
end sub

'******************************************************************************
'Se pintan los datos de los tickets emitidos para el informe previo
sub DibujaTicketsEmitPrevio()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitTicketsEmitidos%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%

		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")

		conn.open session("dsn_cliente")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="TicketsEmitPrevios"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
		command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
		command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
		command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
		command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
		command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
		command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)
		set rst=command.Execute
		resultado=command.Parameters("@p_error").Value

		if not rst.eof then
			SerieAnt=""
			SumBi=0
			SumIva=0
			SumTotal=0
			TSumTotal=0
			while not rst.eof
				if rst("serie")<>SerieAnt then
					'Total Serie Anterior
					if SerieAnt<>"" then
						DrawFila color_terra
                            DrawCelda2 "'ENCABEZADOL width25'", "left", false,""
							DrawCelda2 "'ENCABEZADOL width25'", "left", false,"<b>" & LitTotalserie & "</b>"
							DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumBi,n_decimales,-1,0,-1) & "</b>"
							DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumIva,n_decimales,-1,0,-1) & "</b>"
						CloseFila
						SumBi=0
						SumIva=0
						SumTotal=0
					end if
					'Cabecera Serie
					DrawFila color_blau
						DrawCelda2 "'ENCABEZADOL width25' bgcolor=" & color_terra, "left", false,"<b>" &  Litserie & "</b>"
						DrawCelda2 "'CELDAL7 width25'", "left", false,enc.EncodeForHtmlAttribute(null_s(rst("NombreSerie")))
						DrawCelda2 "'ENCABEZADOL width25' colspan='2' bgcolor=" & color_terra, "left", false,"<b>" &  LitNtickets & "</b>"
						DrawCelda2 "'CELDAL7 width25' colspan='2'", "left", false,rst("tickets")
					CloseFila
					DrawFila color_blau
						DrawCelda2 "'CELDAL7 width25'", "left", false,"&nbsp;"
						DrawCelda2 "'CELDAL7 width25'", "left", false,"&nbsp;"
						DrawCelda2 "'ENCABEZADOL width15'", "left", false,"<b>" & LitDesde & "</b>"
						DrawCelda2 "'CELDAL7 width10'", "left", false,trimCodEmpresa(rst("desde"))
						DrawCelda2 "'ENCABEZADOL width15'", "left", false,"<b>" & LitHasta & "</b>"
						DrawCelda2 "'CELDAL7 width10'", "left", false,trimCodEmpresa(rst("hasta"))
					CloseFila
					'Cabecera Tipos de Iva
					DrawFila color_blau
                        DrawCelda2 "'CELDAL7 width25'", "left", false,""
						DrawCelda2 "'ENCABEZADOL width25'", "left", false,"<b>" & LitTipoIva & "</b>"
						DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & LitBaseImponible & "</b>"
						DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & LitIva & "</b>"
					CloseFila
				end if
				DrawFila color_blau
                    DrawCelda2 "'CELDAL7 width25'", "left", false,""
					DrawCelda2 "'CELDAL7 width25'", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("tipoiva")))
					DrawCelda2 "'CELDAL7 width25' colspan='2'", "left", false,formatnumber(rst("base_imponible"),n_decimales,-1,0,-1)
					SumBi=SumBi + rst("base_imponible")
					DrawCelda2 "'CELDAL7 width25' colspan='2'", "left", false, formatnumber(rst("iva"),n_decimales,-1,0,-1)
					SumIva=SumIva + rst("iva")
					SumTotal=SumTotal + rst("ventas")
					TSumTotal=TSumTotal + rst("ventas")
				CloseFila
				SerieAnt=rst("serie")
				rst.movenext
			wend
			'Total última serie
			DrawFila color_terra
                DrawCelda2 "'ENCABEZADOL width25'", "left", false,""
				DrawCelda2 "'ENCABEZADOL width25'", "left", false,"<b>" & LitTotalserie & "</b>"
				DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumBi,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOL width25' colspan='2'", "left", false,"<b>" & formatnumber(SumIva,n_decimales,-1,0,-1) & "</b>"
			CloseFila
			SumBi=0
			SumIva=0
			SumTotal=0
			'Total General
			'DrawFila color_terra
			'	DrawCelda2Span "CELDA7", "left", true, LitTotalTicketsEmitidos & " " & AbreviaturaMB,3
			'	DrawCelda2Span "CELDARIGHT", "right", true,formatnumber(TSumTotal,n_decimales,-1,0,-1),3
			'CloseFila
		end if
		rst.close
		conn.close
		set command=nothing
		set conn=nothing
	%></table><%
end sub


'******************************************************************************
'Se pintan los datos de las ventas por tipo de pago
sub DibujaVentasTipoPago()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasTipoPago%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "left", false,"<b>" & LitTipoPago & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOL width40'", "left", false,""
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila
		'ega 16/06/2008 union con join y with(nolock)
		strselect="select C.tpago,T.descripcion as NombreTpago,C.ventas,C.ticketsventas,C.anulaciones,C.ticketsanul,(C.ticketsventas+C.ticketsanul) as TotalTickets,(C.ventas-C.anulaciones) as TotalImporte "
		strselect=strselect & " from cierres_tpago C with(nolock) inner join tipo_pago T with(nolock) on C.tpago=T.codigo where C.cierre='" & ncierre & "' order by NombreTpago"
		rst.open strselect,session("dsn_cliente")
		if not rst.eof then
			SumTicketsVentas=0
			SumVentas=0
			SumTicketsAnul=0
			SumAnulaciones=0
			SumTotalTickets=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTpago")
					'DrawCelda2 "CELDAR7", "right", false, rst("ticketsventas")
					SumTicketsVentas=SumTicketsVentas + rst("ticketsventas")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					'DrawCelda2 "CELDAR7", "right", false, rst("ticketsanul")
					SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					'DrawCelda2 "CELDAR7", "right", false, rst("TotalTickets")
					SumTotalTickets=SumTotalTickets + rst("TotalTickets")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
	%></table><%
end sub

'******************************************************************************
'Se pintan los datos de las ventas por tipo de pago del informe previo
sub DibujaVentasTipoPagoPrevio()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasTipoPago%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "left", false,"<b>" & LitTipoPago & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOL width40'", "left", false,""
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila

		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")

		conn.open session("dsn_cliente")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="VentasTipoPagoPrevios"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
		command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
		command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
		command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
		command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
		command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
		command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)
		set rst=command.Execute
		resultado=command.Parameters("@p_error").Value

		if not rst.eof then
			SumTicketsVentas=0
			SumVentas=0
			SumTicketsAnul=0
			SumAnulaciones=0
			SumTotalTickets=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTpago")
					'DrawCelda2 "CELDAR7", "right", false, rst("ticketsventas")
					SumTicketsVentas=SumTicketsVentas + rst("ticketsventas")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					'DrawCelda2 "CELDAR7", "right", false, rst("ticketsanul")
					SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					'DrawCelda2 "CELDAR7", "right", false, rst("TotalTickets")
					SumTotalTickets=SumTotalTickets + rst("TotalTickets")
					DrawCelda2 "'CELDAR7 width20'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width20'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
		conn.close
		set command=nothing
		set conn=nothing
	%></table><%
end sub


'******************************************************************************
'Se pintan los datos de las ventas por tipo de articulo
sub DibujaVentasTipoArticulo()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasTipoArticulo%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTipoArticulo & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOR width40'", "right", false,""
			DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & LitCantidad & "</b>"
			DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & LitCantidad & "</b>"
			DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & LitCantidad & "</b>"
			DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila
		'ega 16/06/2008 with(nolock)
		strselect="select C.tipo,isnull(T.descripcion,'SIN TIPO') as NombreTarticulo,C.ventas,C.cantidadventas,C.anulaciones,C.cantidadanul,(C.cantidadventas-C.cantidadanul) as TotalCantidad,(C.ventas-C.anulaciones) as TotalImporte "
		strselect=strselect & " from cierres_tarticulo C with(nolock) left outer join tipos_entidades T with(nolock) ON (C.tipo=T.codigo and T.tipo='ARTICULO') where C.cierre='" & ncierre & "' order by C.ndet"
		rst.open strselect,session("dsn_cliente")
		if not rst.eof then
			SumCantidadVentas=0
			SumVentas=0
			SumCantidadAnul=0
			SumAnulaciones=0
			SumTotalCantidad=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTarticulo")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("cantidadventas")
					SumCantidadVentas=SumCantidadVentas + rst("cantidadventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("cantidadanul")
					SumCantidadAnul=SumcantidadAnul + rst("cantidadanul")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("TotalCantidad")
					SumTotalCantidad=SumTotalCantidad + rst("TotalCantidad")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOR width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & SumCantidadVentas & "</b>"
				DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & SumCantidadAnul & "</b>"
				DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & SumTotalCantidad & "</b>"
				DrawCelda2 "'ENCABEZADOL width10'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
	%></table><%
end sub

'******************************************************************************
'Se pintan los datos de las ventas por tipo de articulo del informe previo
sub DibujaVentasTipoArticuloPrevio()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasTipoArticulo%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTipoArticulo & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,""
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitCantidad & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitCantidad & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitCantidad & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila

		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")

		conn.open session("dsn_cliente")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="VentasTipoArticuloPrevios"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
		command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
		command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
		command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
		command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
		command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
		command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)
		'on error resume next
		set rst=command.Execute
		'on error goto 0
		resultado=command.Parameters("@p_error").Value

		if not rst.eof then
			SumCantidadVentas=0
			SumVentas=0
			SumCantidadAnul=0
			SumAnulaciones=0
			SumTotalCantidad=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTarticulo")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("cantidadventas")))
					SumCantidadVentas=SumCantidadVentas + rst("cantidadventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("cantidadanul")))
					SumCantidadAnul=SumcantidadAnul + rst("cantidadanul")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("TotalCantidad")))
					SumTotalCantidad=SumTotalCantidad + rst("TotalCantidad")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumCantidadVentas & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumCantidadAnul & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTotalCantidad & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
		conn.close
		set command=nothing
		set conn=nothing
	%></table><%
end sub

'******************************************************************************
'Se pintan los datos de las ventas por operario
sub DibujaVentasOperadores()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasOperadores%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTipoArticulo & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,""
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila
		'ega 16/06/2008 union con join y with(nolock)
		strselect="select C.operador,P.nombre as NombreOperador,C.ventas,C.ticketsventas,C.anulaciones,C.ticketsanul,(C.ticketsventas+C.ticketsanul) as TotalTickets,(C.ventas-C.anulaciones) as TotalImporte "
		strselect=strselect & " from cierres_operador C with(nolock) inner join personal P with(nolock) on C.operador=P.dni where C.cierre='" & ncierre & "' order by NombreOperador"
		rst.open strselect,session("dsn_cliente")
		if not rst.eof then
			SumTicketsVentas=0
			SumVentas=0
			SumTicketsAnul=0
			SumAnulaciones=0
			SumTotalTickets=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreOperador")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("ticketsventas")))
					SumTicketsVentas=SumTicketsVentas + rst("ticketsventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("ticketsanul")))
					SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("TotalTickets")))
					SumTotalTickets=SumTotalTickets + rst("TotalTickets")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsVentas & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsAnul & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTotalTickets & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
	%></table><%
end sub

'******************************************************************************
'Se pintan los datos de las ventas por operario del informe previo
sub DibujaVentasOperadoresPrevio()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasOperadores%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTipoArticulo & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,""
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila

		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")

		conn.open session("dsn_cliente")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="VentasOperadoresPrevios"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
		command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
		command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
		command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
		command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
		command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
		command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)
		'on error resume next
		set rst=command.Execute
		'on error goto 0
		resultado=command.Parameters("@p_error").Value

		if not rst.eof then
			SumTicketsVentas=0
			SumVentas=0
			SumTicketsAnul=0
			SumAnulaciones=0
			SumTotalTickets=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreOperador")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("ticketsventas")))
					SumTicketsVentas=SumTicketsVentas + rst("ticketsventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("ticketsanul")))
					SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, enc.EncodeForHtmlAttribute(null_s(rst("TotalTickets")))
					SumTotalTickets=SumTotalTickets + rst("TotalTickets")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsVentas & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsAnul & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTotalTickets & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
		conn.close
		set command=nothing
		set conn=nothing
	%></table>
<%end sub

'******************************************************************************
'Se pintan los datos de las ventas por tpv
sub DibujaVentasTpv()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasTpv%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTpv & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,""
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila
		'ega 16/06/2008 union con join y with(nolock)
		strselect="select C.tpv,T.descripcion as NombreTPV,C.ventas,C.ticketsventas,C.anulaciones,C.ticketsanul,(C.ticketsventas+C.ticketsanul) as TotalTickets,(C.ventas-C.anulaciones) as TotalImporte "
		strselect=strselect & " from cierres_tpv C with(nolock) inner join tpv T with(nolock) on C.tpv=T.tpv where C.cierre='" & ncierre & "' order by NombreTpv"
		rst.open strselect,session("dsn_cliente")
		if not rst.eof then
			SumTicketsVentas=0
			SumVentas=0
			SumTicketsAnul=0
			SumAnulaciones=0
			SumTotalTickets=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTpv")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("ticketsventas")
					SumTicketsVentas=SumTicketsVentas + rst("ticketsventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("ticketsanul")
					SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("TotalTickets")
					SumTotalTickets=SumTotalTickets + rst("TotalTickets")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsVentas & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsAnul & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTotalTickets & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
	%></table>
<%end sub

'******************************************************************************
'Se pintan los datos de las ventas por tpv del informe previo
sub DibujaVentasTpvPrevio()
	DrawDiv "3-sub","background-color: #eae7e3",""
            %><label class="ENCABEZADOL", style="text-align:left"><b><%=LitVentasTpv%></b></label><%
    CloseDiv
        %><table class="iframe-tab-nospace width100"><%
		DrawFila color_terra
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTpv & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitVentas & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitAnulaciones & "</b> <b>" & AbreviaturaMB & "</b>"
			DrawCelda2 "'ENCABEZADOR width20' colspan='2'", "right", false,"<b>" & LitTotal & "</b> <b>" & AbreviaturaMB & "</b>"
		CloseFila
		DrawFila color_blau
			DrawCelda2 "'ENCABEZADOL width40'", "right", false,""
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitTickets & "</b>"
			DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & LitImporte & "</b>"
		CloseFila

		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")

		conn.open session("dsn_cliente")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="VentasTpvPrevios"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
		command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
		command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
		command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
		command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
		command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
		command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)
		'on error resume next
		set rst=command.Execute
		'on error goto 0
		resultado=command.Parameters("@p_error").Value

		if not rst.eof then
			SumTicketsVentas=0
			SumVentas=0
			SumTicketsAnul=0
			SumAnulaciones=0
			SumTotalTickets=0
			SumTotalImporte=0
			while not rst.eof
				DrawFila color_blau
					DrawCelda2 "'CELDAL7 width40'", "left", false, rst("NombreTpv")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("ticketsventas")
					SumTicketsVentas=SumTicketsVentas + rst("ticketsventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("ventas"),n_decimales,-1,0,-1)
					SumVentas=SumVentas + rst("ventas")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("ticketsanul")
					SumTicketsAnul=SumTicketsAnul + rst("ticketsanul")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("anulaciones"),n_decimales,-1,0,-1)
					SumAnulaciones=SumAnulaciones + rst("anulaciones")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, rst("TotalTickets")
					SumTotalTickets=SumTotalTickets + rst("TotalTickets")
					DrawCelda2 "'CELDAR7 width10'", "' style='text-align:right", false, formatnumber(rst("TotalImporte"),n_decimales,-1,0,-1)
					SumTotalImporte=SumTotalImporte + rst("TotalImporte")
				CloseFila
				rst.movenext
			wend
			DrawFila color_terra
				DrawCelda2 "'ENCABEZADOL width40'", "right", false,"<b>" & LitTotales & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsVentas & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumVentas,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTicketsAnul & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumAnulaciones,n_decimales,-1,0,-1) & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & SumTotalTickets & "</b>"
				DrawCelda2 "'ENCABEZADOR width10'", "right", false,"<b>" & formatnumber(SumTotalImporte,n_decimales,-1,0,-1) & "</b>"
			CloseFila
		end if
		rst.close
		conn.close
		set command=nothing
		set conn=nothing%>
	</table>
<%end sub

'******************************************************************************
'Desplegable con los formatos de impresion
function formato_impresion()
    set rst = Server.CreateObject("ADODB.Recordset")
	seleccion = "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a with(nolock) inner join formatos_imp as b with(nolock) on a.nformato=b.nformato where a.ncliente='"&session("ncliente")&"' and b.tippdoc='CIERRE DE CAJA' order by descripcion"
    rst.cursorlocation=3
	rst.Open seleccion, DsnIlion'', adOpenKeyset, adLockOptimistic
	if not rst.eof then
		if rst("personalizacion")&"">"" then
			personalizacion="../Custom/" & rst("personalizacion") & "/ventas/"
			personalizacionEmail="Custom/" & rst("personalizacion") & "/ventas/"
		else
			personalizacionEmail="ventas/"
		end if
	else
		personalizacionEmail="ventas/"
	end if%>

	<table><tr>
	<td align="right">
		<a class="CELDAREFB" href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(personalizacion)%>' + document.cierres_caja.formato_impresion.value+'ncierre=<%=enc.EncodeForJavascript(ncierre)%>&mode=browse&empresa=<%=session("ncliente")%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitImpresionConFormato%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitImpresionConFormato%></a>
	</td>
	<td class="CELDARIGHT" width="150px" style="width:150px;">
		<select class="CELDARIGHT" width="150px" style="width:150px;" name="formato_impresion"><%
			encontrado=0
			while not rst.eof
				if defecto=rst("descripcion") then
					encontrado=1
					if isnull(rst("parametros")) then
						prm=""
					else
						prm=rst("parametros") & "&"
					end if
					%><option selected="selected" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("fichero") & "?" & prm))%>"><%=enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))%></option><%
				else
					if isnull(rst("parametros")) then
						prm=""
					else
						prm=rst("parametros") & "&"
					end if
					%><option value="<%=enc.EncodeForHtmlAttribute(null_s(rst("fichero")  & "?" & prm))%>"><%=enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))%></option><%
				end if
				rst.movenext
			wend
			rst.close%>
		</select>
	</td>
	</tr></table><%
	formato_impresion=""
end function
'-------------------------------------------------------------------------------------------------------------
'Código principal de la página
'-------------------------------------------------------------------------------------------------------------

%><form name="cierres_caja" method="post"><%  
    'Validar si tiene configurado el módulo comercial VerticalEESS
    const ModVerticalEESS = "V0"
    si_tiene_modulo_vertical=ModuloContratado(session("ncliente"),ModVerticalEESS) 
    if (si_tiene_modulo_vertical <> 0) then    
    %>
        <script language="javascript" type="text/javascript">
            parent.botones.document.getElementById("idaccept").style.display = "none";        
	    </script> <%
    end if

	'Variables Globales
	Dim AbreviaturaMB
	Dim n_decimales

	'Recordsets
	set rst = Server.CreateObject("ADODB.Recordset")

	'Leer parámetros de la página--------------------------------------------------------------------------

	if request.querystring("caju")&""<>"" then
		caju=limpiaCadena(request.querystring("caju"))
	else
		caju=limpiaCadena(request.form("caju"))
	end if

	if request.querystring("ncierre")&""<>"" then
		ncierre=limpiaCadena(request.querystring("ncierre"))
	else
		ncierre=limpiaCadena(request.form("ncierre"))
	end if

	checkCadena ncierre

	lote=limpiaCadena(Request.QueryString("lote"))
	if lote="" then lote=1
	sentido=limpiaCadena(Request.QueryString("sentido"))

	campo=limpiaCadena(Request.QueryString("campo"))
	criterio=limpiaCadena(Request.QueryString("criterio"))
	texto=limpiaCadena(Request.QueryString("texto"))

	caja=limpiaCadena(request.form("caja"))
	dfecha=limpiaCadena(request.form("dfecha"))
	hfecha=limpiaCadena(request.form("hfecha"))

	p_ChkSalidaNoMetalico=nz_b(limpiacadena(request.form("salidanometalico")))
	p_SaldoApertura=limpiaCadena(request.form("hsaldoapertura"))
	p_SaldoMetalico=limpiaCadena(request.form("hsaldometalico"))
	p_MetalicoReal=limpiaCadena(request.form("metalicoreal"))
	p_Descuadre=limpiaCadena(request.form("hdescuadre"))
	p_SalidaCaja=limpiaCadena(request.form("salidacaja"))
	p_SaldoCierre=limpiaCadena(request.form("hsaldocierre"))
	p_PagoSalida=limpiaCadena(request.form("i_pago")) & ""
	
	p_SaldoCierre = replace(p_SaldoCierre, ".",",")
	p_MetalicoReal = replace(p_MetalicoReal, ".",",")
	p_Descuadre = replace(p_Descuadre, ".",",")
	p_SalidaCaja = replace(p_SalidaCaja, ".",",")
	
	if p_PagoSalida="" then
		p_PagoSalida=session("ncliente") & "01"
	end if
	p_FechaSalida=limpiaCadena(request.form("fechasalida")) & ""
	if p_FechaSalida="" then
		p_FechaSalida=hfecha
	end if
	if not isdate(p_FechaSalida) then
		p_FechaSalida=hfecha
	end if

	'Restricciones a la busqueda
	if request.querystring("bus")&""<>"" then
		bus=limpiaCadena(request.querystring("bus"))
	else
		bus=limpiaCadena(request.form("bus"))
	end if%>

	<input type="hidden" name="bus" value="<%=enc.EncodeForHtmlAttribute(bus)%>"/>
	<input type="hidden" name="caju" value="<%=enc.EncodeForHtmlAttribute(caju)%>"/><%
	'---------------------------------------------------------------------------------------------
	PintarCabecera "cierres_caja.asp"
	WaitBoxOculto LitEsperePorFavor
	alarma "cierres_caja.asp"

    strselect = "select dni from personal where login=? and dni like ?+'%'"
	DniPer=DLookupP2(strselect,session("usuario"),adVarChar,50,session("ncliente"),adVarChar,20,session("dsn_cliente"))
    strselect = "select nombre from personal where dni=?"
	NombrePer=DLookupP1(strselect,DniPer&"",adVarChar,20,session("dsn_cliente"))

	%><input type="hidden" name="dni" value="<%=enc.EncodeForHtmlAttribute(DniPer)%>"/><%
	if mode="delete" then '********************************************************************************
		set conn = Server.CreateObject("ADODB.Connection")
		set command =  Server.CreateObject("ADODB.Command")

		conn.open session("dsn_cliente")
		command.ActiveConnection =conn
		command.CommandTimeout = 0
		command.CommandText="BorrarCierreCaja"
		command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamInput,10,ncierre)
		'on error resume next
		command.Execute,,adExecuteNoRecords
		'on error goto 0
		conn.close
		set command=nothing
		set conn=nothing
		mode="add"
		auditar_ins_bor session("usuario"),ncierre,"","baja","","","cierrecaja"
	end if
	if mode="add" then '***********************************************************************************
	%>
		<br/><%
        EligeCeldaResponsive "text", "browse", "CELDA", "", "", 0, "", LitOperador, LitOperador,NombrePer
		'DrawCelda2 "CELDA", "left", false, LitCaja + " : "
		defecto=" "
        DrawDiv "1","",""
        DrawLabel "txtMandatory","",LitCaja
		poner_cajasResponsive1 "width60","","caja","200","codigo","descripcion","","",poner_comillas(caju)
        CloseDiv
        DrawDiv "1","",""
        DrawLabel "txtMandatory","",LitDesdeFecha
        DrawInput "CELDA", "","Dfecha",day(date) & "/" & month(Date) & "/" & year(date),""
        DrawCalendar "Dfecha"
        CloseDiv
		DrawDiv "1","",""
        DrawLabel "txtMandatory","",LitHastaFecha
        DrawInput "CELDA", "","Hfecha",day(date) & "/" & month(Date) & "/" & year(date),""
        DrawCalendar "Hfecha"
        CloseDiv                
	end if
	if mode="confirm" then '****************************************************************************
		if (p_SaldoApertura<>"" and p_SaldoMetalico<>"" and p_MetalicoReal<>"" and p_Descuadre<>"" and p_SalidaCaja<>"" and p_SaldoCierre<>"" and p_FechaSalida<>"" and p_PagoSalida<>"" and p_ChkSalidaNoMetalico<>"") then
			set conn2 = Server.CreateObject("ADODB.Connection")
			set command2 =  Server.CreateObject("ADODB.Command")

		    ''ricardo 22-11-2007 antes de cerrar caja se volcaran los cierres pendientes , procedentes del TPV
			conn2.open session("dsn_cliente")
			command2.ActiveConnection =conn2
			command2.CommandTimeout = 0
			command2.CommandText="InsertaCierresPendientes"
			command2.CommandType = adCmdStoredProc 'Procedimiento Almacenado
            command2.Parameters.Append command2.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
			command2.Parameters.Append command2.CreateParameter("@p_caja",adVarChar,adParamInput,10,"")
			command2.Execute,,adExecuteNoRecords
			conn2.close
			set command2=nothing
			set conn2=nothing

			set conn = Server.CreateObject("ADODB.Connection")
			set command =  Server.CreateObject("ADODB.Command")
			conn.open session("dsn_cliente")
			command.ActiveConnection =conn
			command.CommandTimeout = 0
			command.CommandText="CerrarCajaForm"		
			
			command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
			command.Parameters.Append command.CreateParameter("@p_ncierreTPV",adChar,adParamInput,10,"")
			command.Parameters.Append command.CreateParameter("@p_operario",adVarChar,adParamInput,20,DniPer)
			command.Parameters.Append command.CreateParameter("@p_caja",adVarChar,adParamInput,10,caja)
			command.Parameters.Append command.CreateParameter("@p_dfecha",adVarChar,adParamInput,10,dfecha)
			command.Parameters.Append command.CreateParameter("@p_hfecha",adVarChar,adParamInput,10,hfecha)
			command.Parameters.Append command.CreateParameter("@p_nempresa",adVarChar,adParamInput,5,session("ncliente"))
			command.Parameters.Append command.CreateParameter("@p_saldoapertura",adCurrency,adParamInput,,p_SaldoApertura)
			command.Parameters.Append command.CreateParameter("@p_saldometalico",adCurrency,adParamInput,,p_SaldoMetalico)
			command.Parameters.Append command.CreateParameter("@p_metalicoreal",adCurrency,adParamInput,,p_MetalicoReal)
			command.Parameters.Append command.CreateParameter("@p_descuadre",adCurrency,adParamInput,,p_Descuadre)
			command.Parameters.Append command.CreateParameter("@p_salidacaja",adCurrency,adParamInput,,p_SalidaCaja)
			command.Parameters.Append command.CreateParameter("@p_saldocierre",adCurrency,adParamInput,,p_SaldoCierre)
			command.Parameters.Append command.CreateParameter("@p_fechasalida",adVarChar,adParamInput,10,p_FechaSalida)
			command.Parameters.Append command.CreateParameter("@p_pagosalida",adVarChar,adParamInput,8,p_PagoSalida)
			command.Parameters.Append command.CreateParameter("@p_anotarsalidasnometalico",adInteger,adParamInput,,p_ChkSalidaNoMetalico)

			'Parametros de salida
			command.Parameters.Append command.CreateParameter("@p_error",adInteger,adParamOutput)
			command.Parameters.Append command.CreateParameter("@p_ncierre",adChar,adParamOutput,10,ncierre)

			'on error resume next
			command.Execute,,adExecuteNoRecords
			'on error goto 0

			resultado=command.Parameters("@p_error").Value
			ncierre=command.Parameters("@p_ncierre").Value
			conn.close
			set command=nothing
			set conn=nothing
		else
			resultado=1
		end if

		select case resultado
			case 0
				'OK
				mode="browse"
			case 1
				'No hay registros a incluir en el cierre
				%><script language="javascript" type="text/javascript">
					alert("<%=LitMsgNoExistenApuntes%>");
					parent.botones.document.location = "cierres_caja_bt.asp?mode=add";
                    document.location="cierres_caja.asp?mode=add";
				</script><%
			case else
				'Errores de integridad de datos
				%><script language="javascript" type="text/javascript">
					alert("<%=LitErrorServTec%>");
					parent.botones.document.location = "cierres_caja_bt.asp?mode=add";
                    document.location="cierres_caja.asp?mode=add";
				</script><%
		end select
	end if

	if mode="browse" then '****************************************************************************
		CabeceraCierre mode
        %>
    <div id="tabs" style="display:none"><%
		'Mostrar la barra de pestañas
		BarraNavegacion mode
		DibujaSpans%>
        </div><%
	end if

	if mode="first_save" then '************************************************************************
         %>
		<input type="hidden" name="caja" value="<%=enc.EncodeForHtmlAttribute(caja)%>"/>
		<input type="hidden" name="dfecha" value="<%=enc.EncodeForHtmlAttribute(dfecha)%>"/>
		<input type="hidden" name="hfecha" value="<%=enc.EncodeForHtmlAttribute(hfecha)%>"/>
        <%
		
		 'ega 18/06/2008 borrar la tabla temporal del usuario si existe
         eliminar="if exists (select * from dbo.sysobjects where id = object_id(N'[egesticet].["+DniPer+"]') and OBJECTPROPERTY(id, N'IsTable') = 1) drop table [egesticet].["+DniPer+"] "
         rst.open eliminar,session("dsn_cliente"),adUseClient,adLockReadOnly

        strselect = "select caja from caja where caja=? and fecha>= ?+'00:00:00' and fecha<= ?+'23:59:59' and cierre is null"
        HayRegistros=DLookupP3(strselect,caja,adVarChar,10,dfecha,adDate,,hfecha,adDate,,session("dsn_cliente"))
		
        if HayRegistros<>"" then
            strselect = "select ndecimales from divisas where codigo like ?+'%' and moneda_base<>?"
			n_decimales=DLookupP2(strselect,session("ncliente"),adVarChar,15,"0",adVarChar,1,session("dsn_cliente"))
			CabeceraCierre mode%>
        <div id="tabs" style="display:none"><%
			BarraNavegacion mode
			DibujaSpansPrevios%>
        </div><%
		else
			'No hay registros a incluir en el cierre
			%><script language="javascript" type="text/javascript">
				alert("<%=LitMsgNoExistenApuntes%>");
				parent.botones.document.location = "cierres_caja_bt.asp?mode=add";
            	document.location="cierres_caja.asp?mode=add";
			
			</script><%
		end if
	end if

	if mode="search" then '****************************************************************************
		strwhere=CadenaBusqueda(campo,criterio,texto)
		seleccion="select D.abreviatura,D.ndecimales,CI.codigo,CI.fecha,P.nombre as operador,CA.descripcion as caja, CI.saldo "
		seleccion=seleccion & " from cierres_caja CI with(nolock),personal P with(nolock),cajas CA with(nolock),divisas D with(nolock) "
		seleccion=seleccion & strwhere

		rst.cursorlocation=3
		rst.open seleccion,session("dsn_cliente")
		if not rst.eof then
			lotes=rst.RecordCount/NumReg
			if lotes>clng(lotes) then
				lotes=clng(lotes)+1
			else
				lotes=clng(lotes)
			end if

			if sentido="next" then
				lote=lote+1
			elseif sentido="prev" then
				lote=lote-1
			end if

			rst.PageSize=NumReg
			rst.AbsolutePage=lote

			if rst.RecordCount=1 then
				ncierre=rst("codigo")
				rst.close
				%><script language="javascript" type="text/javascript">
					Editar('<%=ncierre%>');
				</script><%
			else
				NextPrev lote,lotes,campo,criterio,texto,1
				%><table width='100%' border='0' cellspacing="1" cellpadding="1"><%
					'Fila de encabezado
					DrawFila color_fondo
						DrawCelda "ENCABEZADOL","","",0,LitCierre
						DrawCelda "ENCABEZADOL","","",0,LitFecha
						DrawCelda "ENCABEZADOL","","",0,LitOperador
						DrawCelda "ENCABEZADOL","","",0,LitCaja
						DrawCelda "ENCABEZADOR","","",0,LitSaldo
					CloseFila
					fila=1
					while not rst.EOF and fila<=NumReg
						'Seleccionar el color de la fila.
						if ((fila+1) mod 2)=0 then
							color=color_blau
							con_negrita=false
						else
							color=color_terra
							con_negrita=false
						end if

						DrawFila color
							DrawCeldahref "CELDAREF","left","false",trimCodEmpresa(rst("codigo")),"javascript:Editar('" & enc.EncodeForJavascript(null_s(rst("codigo"))) & "')"
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(null_s(rst("fecha")))
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(null_s(rst("operador")))
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(null_s(rst("caja")))
							DrawCelda "CELDARIGHT","","",0,enc.EncodeForHtmlAttribute(null_s(formatnumber(rst("saldo"),rst("ndecimales"),-1,0,-1) & rst("abreviatura")))
						CloseFila

						fila=fila+1
						rst.movenext
					wend
				%></table><%
				NextPrev lote,lotes,campo,criterio,texto,2
			end if
		else 'NO HAY REGISTROS
			rst.close
			%><font class='CEROFILAS'><%=LitCeroFilas%></font><%
		end if
	end if%>
</form><%
set rst=nothing
end if%>
</body>
</html>