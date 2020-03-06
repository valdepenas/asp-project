<%@ Language=VBScript %>
<%
'CODIGOS DE AÑADIDURAS/MODIFICACIONES -----------------------------------------------------
'------------------------------------------------------------------------------------------%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

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
<!--#include file="tickets.inc" -->
<!--#include file="documentos.inc" -->    
<!--#include file="../perso.inc" -->
<!--#include file="../generarFtoImpresion.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file= "../CatFamSubResponsive.inc"-->
<!--#include file="../js/generic.js.inc" -->
<!--#include file="../js/tabs.js.inc" -->
<!--#include file="../js/animatedCollapse.js.inc"-->
<!--#include file="../styles/Section.css.inc"-->
<!--#include file="../styles/ExtraLink.css.inc"-->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/dropdown.js.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->
<!--#include file="../common/poner_cajaResponsive.inc" -->

<%si_tiene_modulo_terminales=ModuloContratado(session("ncliente"),ModTerminales)
%>
<%si_tiene_modulo_petroleos=ModuloContratado(session("ncliente"),ModOrCU)
%>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('DatosGenerales', 'fade=1');
    animatedcollapse.addDiv('DATTICKETS', 'fade=1');
    animatedcollapse.addDiv('DATTOTAL', 'fade=1');

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init();
</script>

<script language="javascript" type="text/javascript">

function abrir_beneficios(viene,ndocumento){
	var bloqueado;

	bloqueado=document.tickets.h_nfactura.value;
	AbrirVentana("costes_doc.asp?ndoc=" + ndocumento + "&viene=" + viene + "&titulo=" + ndocumento + "&bloqueado=" + bloqueado,'P',<%=AltoVentana%>,<%=AnchoVentana%>);
}

function ocultar_genfact()
{
	document.getElementById("genfact").style.display="none";
}

function ocultar_datosFuga()
{
    var tipoPago = document.getElementById("medioPagoTicket").innerHTML.trim().toLowerCase()

    if(tipoPago != "fuga")
    {
        document.getElementById("datosFuga").style.display="none";
		document.getElementById("botones").contentWindow.document.getElementById("ideditar").style.display="none";
    }
}

function tier1Menu(objMenu,objImage)
{
	if (objMenu.style.display == "none")
	{
		objMenu.style.display = "";
		objImage.src = "../Images/<%=ImgCarpetaAbierta%>";
		switch (objMenu.id)
		{
			case "CABECERA":
				document.getElementById("DETALLES").style.display="none";
				document.getElementById("PAGOS_CUENTA").style.display="none";
				document.getElementById("img2").src="../Images/<%=ImgCarpetaCerrada%>";
				document.getElementById("img3").src="../Images/<%=ImgCarpetaCerrada%>";
				ocultar_datosFuga();
				break;

			case "DETALLES":
				document.getElementById("CABECERA").style.display="none";
				document.getElementById("PAGOS_CUENTA").style.display="none";
				document.getElementById("img1").src="../Images/<%=ImgCarpetaCerrada%>";
				document.getElementById("img3").src="../Images/<%=ImgCarpetaCerrada%>";
				break;

			case "PAGOS_CUENTA":
				document.getElementById("CABECERA").style.display="none";
				document.getElementById("DETALLES").style.display="none";
				document.getElementById("img1").src="../Images/<%=ImgCarpetaCerrada%>";
				document.getElementById("img2").src="../Images/<%=ImgCarpetaCerrada%>";
				break;
		}
	}
	else
	{
		objMenu.style.display = "none";
		objImage.src = "../images/<%=ImgCarpetaCerrada%>";
	}
	Redimensionar();
}

//***************************************************************************

function GenerarFactura(nticket,factura)
{
	if (factura=='')
	{
		if (document.tickets.serieFAC.value=='') alert("<%=LitMsgSerieNoNulo%>");
		else
		{
			document.tickets.action="tickets.asp?nticket=" + nticket + "&mode=genera&SerieFac=" + document.tickets.serieFAC.value;
			document.tickets.submit();
			document.getElementById("waitBoxOculto").style.visibility="visible";
		}
	}
	else alert("<%=LitMsgAlbTieneFactura%>");
}

function cambiarfecha(fecha,modo)
{
	var fecha_ar=new Array();
	if (fecha!="" && fecha.length<10)
	{
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length)
		{
			if (fecha.substring(l,l+1)=='/')
			{
				suma++;
				fecha_ar[suma]="";
			}
			else
			{
				if (fecha.substring(l,l+1)!='') fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
			}
			l++;
		}
		if (suma!=2)
		{
			window.alert("<%=LitFechaMal%> " + modo );
			return false;
		}
		else
		{
			nonumero=0;
			while (suma>=0 && nonumero==0)
			{
				if (isNaN(fecha_ar[suma])) nonumero=1;
				if (fecha_ar[suma].length>2 && suma!=2) nonumero=1;
				if (fecha_ar[suma].length>4 && suma==2) nonumero=1;
				suma--;
			}

			if (nonumero==1)
			{
				window.alert("<%=LitFechaMal%> " + modo);
				return false;
			}
		}
	}
	return true;
}

//Añade un pago a cuenta.
function addPago(npedido)
{
	if (document.tickets.importePago.value=="") document.tickets.importePago.value=0;
	if (document.tickets.fechaPago.value=="")
	{
		window.alert("<%=LitErrFechaPago%>");
		return;
	}

    if (document.tickets.fechaPago.value.length<10)
   	{
    	if (!checkdate(document.tickets.fechaPago))
	    {
		    window.alert("<%=LitMsgFechaFecha%>");
    		return;
	    }
	}

	if (document.tickets.fechaPago.value.length>10)
   	{
    	if (!chkdatetime(document.tickets.fechaPago.value))
	    {
		    window.alert("<%=LitMsgFechaFecha%>");
    		return;
	    }
	}

	if (!cambiarfecha(document.tickets.fechaPago.value,"Fecha Pago")) return;

	if (isNaN(document.tickets.importePago.value.replace(",",".")))
	{
		window.alert("<%=LitErrImportePago%>");
		return;
	}
	else
	{
		if (parseFloat(document.tickets.importePago.value.replace(",","."))==0)
		{
			window.alert("<%=LitMsgImportePositivo%>");
			return;
		}
	}
	if (document.tickets.descripcionPago.value=="")
	{
		window.alert("<%=LitMsgDesVacia%>");
		return;
	}
	if (document.tickets.tipoPago.value=="")
	{
		window.alert("<%=LitMsgTipoPagoNoNulo%>");
		return;
	}
	//Asignar los valores a los campos del submarco de detalles
	fr_PagosCuenta.document.tickets_pagos.fecha.value=document.tickets.fechaPago.value;
	fr_PagosCuenta.document.tickets_pagos.importe.value=document.tickets.importePago.value;
	fr_PagosCuenta.document.tickets_pagos.descripcion.value=document.tickets.descripcionPago.value;
	fr_PagosCuenta.document.tickets_pagos.medio.value=document.tickets.tipoPago.value;
	fr_PagosCuenta.document.tickets_pagos.ncaja.value=document.tickets.ncaja.value;
	//Recargar el submarco de pagos a cuenta
	fr_PagosCuenta.document.tickets_pagos.action="tickets_pagos.asp?mode=first_save";
	fr_PagosCuenta.document.tickets_pagos.submit();
	//Limpiar los campos del formulario
	var hoy=new Date();
	document.tickets.fechaPago.value=hoy.getDate() + "/" + (hoy.getMonth()+1) + "/" + hoy.getFullYear();
	document.tickets.importePago.value="0";
	document.tickets.descripcionPago.value="";
	document.tickets.tipoPago.value="";
	//Colocar el foco en el campo de cantidad.
	document.tickets.fechaPago.focus();
	document.tickets.fechaPago.select();
}

//Comprueba si el importe del pago es numerico
function importepagoComp()
{
	if (isNaN(document.tickets.importePago.value.replace(",",".")))
	{
		window.alert("<%=LitErrImportePago2%>");
		return;
	}
}

function Acaja(nticket)
{
	if (document.tickets.impcaja.value=="") document.tickets.impcaja.value=0;
	if (isNaN(document.tickets.impcaja.value.replace(",",".")))
	{
		window.alert("<%=LitMsgImporteNumerico%>");
		return false;
	}
	else
	{
		if (parseFloat(document.tickets.impcaja.value.replace(",","."))==0)
		{
			window.alert("<%=LitErrImportePago%>");
			return false;
		}
	}
	if (document.tickets.ncaja.value=="") alert("<%=LitMsgCajaNoNulo%>");
	else
	{	if (document.tickets.i_pago.value=="") alert("<%=LitMsgTipoPagoNoNulo%>");
		else
		{
			fr_PagosCuenta.document.tickets_pagos.action="tickets_pagos.asp?mode=acaja&ndoc=" + nticket + "&impcaja=" + document.tickets.impcaja.value + "&i_pago=" + document.tickets.i_pago.value + "&ncaja=" + document.tickets.ncaja.value;
			fr_PagosCuenta.document.tickets_pagos.submit();
			if (document.getElementById("PAGOS_CUENTA").style.display == "none") tier1Menu(PAGOS_CUENTA,document.getElementById("img3"));
		}
	}
}
//***************************************************************************

function Redimensionar()
{
    var alto = 0;
    if (parent.document.body.offsetHeight) alto = parent.document.body.offsetHeight;
    else alto = parent.self.innerHeight;
	if (document.getElementById("DETALLES").style.display=="")
	{
	    if (alto > 150)
        {
            if (alto - 372 > 150) document.getElementById("frDetalles").style.height = alto - 372;
            else document.getElementById("frDetalles").style.height = 150;
        }
        else document.getElementById("frDetalles").style.height = 150;
    }
    else
    {
        if (document.getElementById("PAGOS_CUENTA").style.display=="")
        {
            if (alto > 150)
            {
                if (alto - 410 > 150) document.getElementById("frPagosCuenta").style.height = alto - 410;
                else document.getElementById("frPagosCuenta").style.height = 150;
            }
            else document.getElementById("frPagosCuenta").style.height = 150;
        }
    }
}
</script>

<body class="BODY_ASP" <%=iif(Request.QueryString("mode")="browse", "onresize='javascript:Redimensionar();'", "")%>>
 <%

'******************************************************************************
'Crea la tabla que contiene la barra de grupos de datos.
sub BarraNavegacion(modo)%>
	<table bgcolor="<%=color_pestanas%>" width="100%">
		<tr>
			<td align="justify">
				<%if modo="add" or mode="edit" then%>
					<div class=CELDAB>
					<img src="../images/<%=ImgCarpetaAbierta%>" <%=ParamImgCarpetaAbierta%> ID="img1" alt="" title=""/>&nbsp;&nbsp;<%=LitCabecera%>&nbsp;&nbsp;
				<%else%>
					<div class=CELDAB onclick="tier1Menu(CABECERA,document.getElementById('img1'))">
					<%if mode="editFuga" then%>
						<img src="../images/<%=ImgCarpetaAbierta%>" <%=ImgCarpetaAbierta%> ID="img1" alt="" title=""/>&nbsp;&nbsp;<%=LitCabecera%>&nbsp;&nbsp;
					<%else%>
						<img src="../images/<%=ImgCarpetaCerrada%>" <%=ParamImgCarpetaCerrada%> ID="img1" alt="" title=""/>&nbsp;&nbsp;<%=LitCabecera%>&nbsp;&nbsp;
					<%end if%>
				<%end if%>
				</div>
			</td>
			<%if modo<>"add" and modo<>"edit" then%>
				<td align="justify">
					<div class=CELDAB onclick="tier1Menu(DETALLES,document.getElementById('img2'));">				
					<%if mode="editFuga" then%>	
						<img src="../images/<%=ImgCarpetaCerrada%>" <%=ParamImgCarpetaCerrada%> ID="img2" alt="" title=""/>&nbsp;&nbsp;<%=LitDetalles%>&nbsp;&nbsp;				
					<%else%>	
						<img src="../images/<%=ImgCarpetaAbierta%>" <%=ParamImgCarpetaAbierta%> ID="img2" alt="" title=""/>&nbsp;&nbsp;<%=LitDetalles%>&nbsp;&nbsp;
					<%end if%>			
					</div>
				</td>
				<td align="justify">
					<div class=CELDAB onclick="tier1Menu(PAGOS_CUENTA,document.getElementById('img3'));">
					<img src="../images/<%=ImgCarpetaCerrada%>" <%=ParamImgCarpetaCerrada%> ID="img3" alt="" title=""/>&nbsp;&nbsp;<%=LitPagosCuenta%>&nbsp;&nbsp;
					</div>
				</td>
			<%end if%>
		</tr>
	</table>
<%end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0
%>
<form name="tickets" method="post">

    <%'Leer parámetros de la página
	mode=Request.QueryString("mode")
    WaitBoxOculto LitEsperePorFavor

	nticket=limpiaCadena(Request.QueryString("nticket"))
	if nticket="" then nticket=limpiaCadena(Request.form("nticket"))
	if nticket="" then nticket=limpiaCadena(Request.QueryString("ndoc"))
	checkCadena nticket

	if request.querystring("caju")>"" then
		caju=limpiaCadena(request.querystring("caju"))
	else
		caju=limpiaCadena(request.form("caju"))
	end if

	campo=limpiaCadena(Request.QueryString("campo"))
  	criterio=limpiaCadena(Request.QueryString("criterio"))
  	texto=limpiaCadena(Request.QueryString("texto"))

	if request.querystring("viene")&""="tpv" then
	    viene=limpiaCadena(request.querystring("viene"))
		tpv=limpiaCadena(request.querystring("u"))
	end if

	serieFAC=limpiaCadena(Request.QueryString("serieFAC"))
	if serieFAC& ""="" then
		serieFAC=limpiaCadena(Request.form("serieFAC"))
	end if
	
	dim mcb

	''ricardo 21-7-2009 se pone esta funcion para el parametro mcb
	ObtenerParametros("tickets")%>
	<input type="hidden" name="h_mode" value="<%=EncodeForHtml(mode)%>"/>
	<input type="hidden" name="h_nticket" value="<%=EncodeForHtml(nticket)%>"/>
	<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>"/>

	<%lote=limpiaCadena(Request.QueryString("lote"))
	if lote="" then lote=1
	sentido=limpiaCadena(Request.QueryString("sentido"))
	donde=limpiaCadena(request.querystring("donde"))

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion

	PintarCabecera "ticketsF.asp"
	Alarma "tickets.asp"

	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstPed = Server.CreateObject("ADODB.Recordset")
	set rstDetTicket = Server.CreateObject("ADODB.Recordset")
	set rstIvas = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")

	n_decimales=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0",session("dsn_cliente"))
	n_abreviatura=d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0",session("dsn_cliente"))

    ' Anulamos el ticket
    if mode="anular" then
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open session("dsn_cliente")
        command.ActiveConnection = conn
        command.CommandTimeout = 0
        command.CommandText="RepsolPeruSunatAnularTicket"
        command.CommandType = adCmdStoredProc 
        command.Parameters.Append command.CreateParameter("@nticket ", adVarChar, adParamInput,20, nticket)
        command.prepared=true
        command.execute

        conn.close
        set command = nothing
        set conn = nothing

        ' Restaruamos el valor del mode.
        mode = "browse"
    end if

	'Guardar datos CABECERA
	if mode="saveFuga" then
	Request.Form("firstname")
		matricula=limpiaCadena(Request.form("matricula"))
		if matricula="" then matricula=limpiaCadena(Request.QueryString("matricula"))
		observacion=limpiaCadena(Request.form("observacion"))
		if observacion="" then observacion=limpiaCadena(Request.QueryString("observacion"))
	
        set connS = Server.CreateObject("ADODB.Connection")
        set commandS = Server.CreateObject("ADODB.Command")
        connS.open session("dsn_cliente")
        commandS.ActiveConnection = connS
        commandS.CommandTimeout = 0
        commandS.CommandText="updateTicketDatosFuga"
        commandS.CommandType = adCmdStoredProc 
		
		commandS.Parameters.Append commandS.CreateParameter("@nticket ", adVarChar, adParamInput,20, nticket)
		commandS.Parameters.Append commandS.CreateParameter("@matricula ", adVarChar, adParamInput,50, matricula)
		commandS.Parameters.Append commandS.CreateParameter("@observaciones ", adVarChar, adParamInput,-1, observacion)
			
        commandS.prepared=true
        commandS.execute

        connS.close
        set commandS = nothing
        set connS = nothing

        ' Restaruamos el valor del mode.
        mode = "browse"
    end if
	
	'Mostrar los datos de la página.

	if mode="genera" then ' modo para generar la factura
        Server.ScriptTimeout = 3000

     	DropTable session("usuario"), session("dsn_cliente")
		crear ="CREATE TABLE [egesticet].[" & session("usuario") & "] (nventa varchar(20),ndocumento varchar(20),fecha smalldatetime,ncliente char(10),rsocial varchar(100),cifedi varchar(20),medio_pago varchar(8),nompago varchar(100),total_ticket money,ndecimales smallint,abreviatura varchar(5),nclienteAdmin varchar(30),seleccionado bit"
        crear=crear & ",descuento real,descuento2 real,descuento3 real, observaciones varchar(max), campo1 nvarchar(50) "
        crear=crear & ")"
        rst.cursorlocation=2
		rst.open crear,session("dsn_cliente") '',adUseClient,adLockReadOnly
		GrantUser session("usuario"), session("dsn_cliente")
        
        cad_con=mid(DsnImport,1,instr(DSNImport,"Initial Catalog=")+15)
        cad_con=cad_con&mid(session("dsn_cliente"),instr(session("dsn_cliente"),"Initial Catalog=")+16,instr(session("dsn_cliente"),";User Id=")-instr(session("dsn_cliente"),"Initial Catalog=")-16 )
        cad_con=cad_con&mid(DSNImport, instr(DSNImport,";User Id="),len(DSNImport)-instr(DSNImport,";User Id=")+1)

		strSelTickets="select t.nventa,t.nticket,convert(varchar,t.fecha,103),t.ncliente,c.rsocial,c.cifedi,t.medio_pago,tp.descripcion as nompago,t.total_ticket,d.ndecimales,d.abreviatura "
		strSelTickets=strSelTickets & ",(select top 1 ncliente from ilion_admin.dbo.clientes where cifedi=c.cifedi) as nclienteAdmin "
		strSelTickets=strSelTickets & ",1 as seleccionado "
        strSelTickets=strSelTickets & ",t.descuento,t.descuento2,t.descuento3, t.observaciones, t.campo1 "
		strSelTickets=strSelTickets & " from tickets as t with(NOLOCK) left outer join clientes as c with(NOLOCK) on c.ncliente=t.ncliente "
		strSelTickets=strSelTickets & " left outer join tipo_pago as tp with(NOLOCK) on tp.codigo=t.medio_pago,divisas as d"
		strSelTickets=strSelTickets & " where t.nticket='" & nticket & "' and d.codigo=t.divisa and t.nfactura is null "

		strselect="insert into [egesticet].[" & session("usuario") & "] " & strSelTickets
        rst.cursorlocation=2
		rst.open strselect, cad_con'',adUseClient,adLockReadOnly
		if rst.state<>0 then rst.close

		nomusu=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",dsnilion)

		strFactTick="exec ConvertirTicketsFtras @serie='" & serieFAC & "',@fechaFactura='" & date & "',@nclienteUni='',@nclienteAdminUni=''"
		strFactTick=strFactTick & ",@unificar=1,@nusuario='" & session("usuario") & "',@session_ncliente='" & session("ncliente") & "',@ip='" & Request.ServerVariables("REMOTE_ADDR") & "',@nomusuario='" & nomusu & "'"
		set conFactTick=Server.CreateObject("ADODB.Connection")
		conFactTick.CommandTimeout=300
		conFactTick.open session("dsn_cliente")
		on error resume next
		set rst=conFactTick.execute(strFactTick)
		if err.number<>0 then
			el_error=err.number
		end if
		on error goto 0
		if rst.state<>0 then
			on error resume next
			Primero=rst("primero")
			if err.number<>0 then
                on error goto 0
                on error resume next
                Primero=rst(0)
                if err.number<>0 then
				    el_error=err.number
                end if
                if Primero = "No se ha creado ninguna factura" then
                    Primero = ""
                end if
                if Primero & "">"" then
                    donde = InStr(1, Primero, "facturas desde ", 1)
                    if donde>0 then
                        donde = donde + Len("facturas desde ")
                        donde2 = InStr(donde + 1, Primero, " hasta ", 1)
                        if donde2=0 then donde2=len(Primero)
                        Primero = session("ncliente") & Mid(Primero, donde, donde2 - donde)
                    end if
                end if
			else
				Ultimo=rst("ultimo")
				el_error=rst("error")
			end if
			on error goto 0
			rst.close
		end if
		conFactTick.close
		set conFactTick=nothing

		if Primero & "">"" then
            ''actualizamos el nfactura al ticket
            set commandUpdT =  Server.CreateObject("ADODB.Command")
            set connT = Server.CreateObject("ADODB.Connection")
            connT.open session("dsn_cliente")
            commandUpdT.ActiveConnection =connT
            commandUpdT.CommandTimeout = 0
            commandUpdT.CommandText="update tickets with(updlock) set nfactura=? where nticket = ?"
            commandUpdT.CommandType = adCmdText
            commandUpdT.Parameters.Append commandUpdT.CreateParameter("@nfactura",adVarChar,adParamInput,20,Primero)
            commandUpdT.Parameters.Append commandUpdT.CreateParameter("@nticket",adVarChar,adParamInput,20,nticket)
            on error resume next
            set rstT = commandUpdT.Execute
            on error goto 0
            set rstT = nothing
            set commandUpdT = nothing
            set connT = nothing

            ''ponemos observaciones a la factura
            ''actualizamos el nfactura al ticket
            set commandUpdT2 =  Server.CreateObject("ADODB.Command")
            set connT2 = Server.CreateObject("ADODB.Connection")
            connT2.open session("dsn_cliente")
            commandUpdT2.ActiveConnection =connT2
            commandUpdT2.CommandTimeout = 0
            commandUpdT2.CommandText="update facturas_cli with(updlock) set observaciones='Factura generada del ticket ' + ? where nfactura = ?"
            commandUpdT2.CommandType = adCmdText
            commandUpdT2.Parameters.Append commandUpdT2.CreateParameter("@nticket",adVarChar,adParamInput,20,trimCodEmpresa(nticket))
            commandUpdT2.Parameters.Append commandUpdT2.CreateParameter("@nfactura",adVarChar,adParamInput,20,Primero)
            on error resume next
            set rstT2 = commandUpdT2.Execute
            on error goto 0
            set rstT2 = nothing
            set commandUpdT2 = nothing
            set connT2 = nothing

			nclienteFac=d_lookup("ncliente","facturas_cli","nfactura='" & Primero & "'",session("dsn_cliente"))
			auditar_ins_bor session("usuario"),Primero,nclienteFac,"alta","","","facturas_cli"

			mode="browse"%>
			<script language="javascript" type="text/javascript">
				if (window.confirm("<%=LitMsgGenUnaFactura%> <%=trimCodEmpresa(Primero)%>. <%=LitMsgDeseaVer%>"))
                    AbrirVentana('../central.asp?pag1=ventas/facturas_cli.asp&ndoc=<%=enc.EncodeForJavascript(Primero)%>&mode=browse&pag2=ventas/facturas_cli_bt.asp&titulo=<%=LitDetallesFact%> <%=enc.EncodeForJavascript(trimCodEmpresa(Primero))%>','P',<%=altoventana%>,<%=anchoventana%>);
			</script>
		<%else
			mode="browse"%>
			<script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgFacturasNoGeneradas%>");
			</script>
		<%end if
	end if

	if (mode="browse") or (mode ="editFuga") then
        rstAux.cursorlocation=3
		rstAux.open "select nticket from tickets with(nolock) where nticket='" & nticket & "'", session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
		if rstAux.eof then
			nticket=""%>
			<script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgDocsNoExiste%>");
				parent.botones.document.location="tickets_bt.asp?mode=search";
			</script>
			<%mode="add"
		end if
		rstAux.close
	end if
	if (mode="browse") or (mode ="editFuga") then
		if nticket="" then
            rstAux.cursorlocation=3
			rstAux.open "select top 1 nticket from tickets with(nolock) where nticket like '" & session("ncliente") & "%' order by fecha desc,nticket desc", session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
			if not rstAux.eof then nticket=rstAux("nticket")
			rstAux.close
		end if

		strselect="select t.divisa,t.nventa,t.nticket,t.fecha,t.total_ticket,t.anulacion,t.nliquidacion,t.nturno,t.nterminal,t.base_imponible,t.total_iva,t.total_ticket,tp.descripcion as tpv,caj.descripcion as caja,tip.descripcion as tpago,t.nfactura,t.serie,d.abreviatura,d.ndecimales,p.dni as operador,p.nombre as nomoperador,c.ncliente,c.rsocial, t.observaciones, t.campo1 "
		strfrom=" from tickets as t with(NOLOCK) "
		strfrom=strfrom & " left outer join divisas as d with(nolock) on d.codigo=t.divisa "
		strfrom=strfrom & " left outer join personal as p with(nolock)  on p.dni=t.usuario "
		strfrom=strfrom & " left outer join clientes as c with(nolock)  on c.ncliente=t.ncliente "
		strfrom=strfrom & " left outer join tpv as tp with(nolock)  on tp.tpv=t.tpv "
		strfrom=strfrom & " left outer join cajas as caj with(nolock)  on caj.codigo=t.caja "
		strfrom=strfrom & " left outer join tipo_pago as tip with(nolock)  on tip.codigo=t.medio_pago "
		strwhere=" where nticket='" & nticket & "'"
        rst.cursorlocation=3
		rst.Open (strselect & strfrom & strwhere), session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
	elseif mode="search" then
	end if

	'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION
	if (mode="browse") or (mode="editFuga") then%>
		<table width="100%" border="0">
   			<tr>
				<%if not rst.eof then
					if mode="browsexx" then
						DrawCelda "CELDA style='width:50px'","","",0,"&nbsp;"
						DrawCelda "CELDA style='width:150px'","","",0,"&nbsp;"
						DrawCelda "ENCABEZADOC","","",0,"&nbsp;"
						''ricardo 13-3-20003
						''si la serie tiene un formato de impresion sera este el de por defecto
						''si no sera el elegido en la tabla formatos impresion de ilion
						if not rst.eof then
							defecto=obtener_formato_imp(rst("serie"),"TICKET")
						end if
						''''''''
						seleccion = "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='TICKET' order by descripcion"
                        rstSelect.cursorlocation=3
						rstSelect.Open seleccion, DsnIlion'', adOpenKeyset, adLockOptimistic%>
						<td align="right" style='width:75px'><a class='CELDAREFB' href="javascript:AbrirVentana(document.tickets.formato_impresion.value+'nticket=<%="(\'"+enc.EncodeForJavascript(p_nticket)+"\')"%>&mode=browse&empresa=<%=session("ncliente")%>&novei=<%=enc.EncodeForJavascript(novei)%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitImpresionConFormato%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitImpresionConFormato%></a></td>
                        <td class='CELDARIGHT' width='150px'>
							<select class='CELDARIGHT' width='150px' style='width:150px' name="formato_impresion">
								<%
								encontrado=0
								while not rstSelect.eof
									if defecto=rstSelect("descripcion") then
										encontrado=1
										if isnull(rstSelect("parametros")) then
											prm=""
										else
											prm=rstSelect("parametros") & "&"
										end if
										%><option selected value="<%=EncodeForHtml(rstSelect("fichero")) & "?" & prm%>"><%=EncodeForHtml(rstSelect("descripcion"))%></option><%
									else
										if isnull(rstSelect("parametros")) then
											prm=""
										else
											prm=rstSelect("parametros") & "&"
										end if
										%><option value="<%=EncodeForHtml(rstSelect("fichero"))  & "?" & prm%>"><%=EncodeForHtml(rstSelect("descripcion"))%></option><%
									end if
									rstSelect.movenext
								wend%>
							</select>
						</td><%
						rstSelect.close
						pagina="../crearpdf.asp?destinatario=" & enc.EncodeForJavascript(rst("ncliente")) & "&ndoc=" & enc.EncodeForJavascript(rst("nticket")) & "&tdoc=ALBARAN&dedonde=DOCUMENTOV&empresa=" & enc.EncodeForJavascript(session("ncliente")) & "&mode=DOC&url=ventas/"
						%><td class=CELDARIGHT style='width:20px'><a class='CELDAREFB' href="javascript:AbrirVentana('<%=pagina%>' + document.tickets.formato_impresion.value + 'nticket=<%="(\'"+enc.EncodeForJavascript(p_nticket)+"\')"%>&novei=<%=enc.EncodeForJavascript(novei)%>','A','<%=AltoVentana%>','600')" OnMouseOver="self.status='<%=LitEnvEmail%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="../images/<%=ImgEnviarEmail%>" <%=ParamImgEnviarEmail%> alt="<%=ucase(LitEnvEmail)%>" title="<%=ucase(LitEnvEmail)%>"/></a></td><%
					else%>
					 	<td align="right"></td>
					<%end if%>
				<%end if%>
			</tr>
		</table>

		<%VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarFacturasCli)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina%>
         <div class="headers-wrapper">
            <%
                DrawDiv "header-date","",""
                DrawLabel "","",LitFecha
                DrawSpan "","",EncodeForHtml(rst("fecha")),""
                CloseDiv


				if mode="browse" or mode="editFuga" then
					no_mostrar=""
					if rst("nfactura") & "">"" then
						no_mostrar="none"
					end if%>
								<%  DrawDiv "header-fact","display: " & no_mostrar,"genfact"                            
                                    %><label><a class='CELDAREFB' href="javascript:if (window.confirm('<%=LitDesTickConvAFact%>')==true){ocultar_genfact();GenerarFactura('<%=enc.EncodeForJavascript(null_s(rst("nticket")))%>','<%=enc.EncodeForJavascript(null_s(rst("nfactura")))%>')}" OnMouseOver="self.status='<%=LitGentTicFact%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitGenFacTick%></a></label>									
  									<%nserie_aux=obtener_serie_doc("TICKET-FAC_CLI","","",0,rst("ncliente"))
                                    rstAux.cursorlocation=3
									rstAux.open "select nserie,right(nserie,len(nserie)-5) as nserie2 from series with(nolock) where nserie like '" & session("ncliente") & "%' and tipo_documento='FACTURA A CLIENTE' order by nserie2",session("dsn_cliente")'',adOpenKeyset,adLockOptimistic                                    
		 							DrawSelect "","width:150px;","serieFAC",rstAux,nserie_aux,"nserie","nserie2","",""
		 							rstAux.close
                                    CloseDiv%>
				<%end if

				if mode="browse" or mode="editFuga" then
                    DrawDiv "header-bill", "", ""
                        DrawLabel "", "", LitSerie
                        DrawSpan "", "", EncodeForHtml(trimCodEmpresa(rst("serie"))), ""
                    CloseDiv
				end if

			        if viene="tpv" then
			            caju=d_lookup("caja","tpv","tpv like '"&session("ncliente")&"%' and tpv='"&tpv&"' ", session("dsn_cliente"))
			        end if

			    ''ricardo 22-6-2007 si el ticket este o no facturado , se podra pagar(dicho por JAR)
					EnCajaEntrada=CambioDivisa(d_sum("importe","caja","ndocumento='" & rst("nticket") & "' and tdocumento='TICKET' and tanotacion='entrada' ",session("dsn_cliente")),rst("divisa"),rst("divisa"))
					EnCajaSalida=CambioDivisa(d_sum("importe","caja","ndocumento='" & rst("nticket") & "' and tdocumento='TICKET' and tanotacion='salida' ",session("dsn_cliente")),rst("divisa"),rst("divisa"))
					EnCaja=EnCajaEntrada-EnCajaSalida
					Pendiente=miround(rst("total_ticket")-EnCaja,rst("ndecimales"))

                    DrawDiv "header-note","","LitPenCobro"
                    DrawLabel "","",LitPenCobEnCaj 
                    penCob = formatnumber(null_z(pendiente),rst("ndecimales"),-1,0,iif(mode="browse",-1,0)) & " " & rst("abreviatura")
                    DrawSpan "","",EncodeForHtml(penCob),""
                    CloseDiv

                    if si_tiene_modulo_petroleos<>0 and not isnull(rst("anulacion")) then
                         DrawDiv "header-note", "", ""
                            DrawLabel "", "", LitFacturaRect
                            DrawSpan "", "", EncodeForHtml(trimCodEmpresa(rst("anulacion"))), ""
                         CloseDiv
                    end if


                DrawDiv "header-note", "", ""
                    DrawLabel "", "", LitTicket
                    DrawSpan "", "", EncodeForHtml(trimCodEmpresa(rst("nventa"))), ""
                CloseDiv

                    DrawDiv "header-note","",""
					defecto=""
                    poner_cajasResponsive1 "input-ncaja",defecto,"ncaja","100","codigo","descripcion","","",poner_comillas(caju)
                    
                %><span class="header-note-inputCaja">
                    <input class='CELDAR7' type="Text" name="impcaja" value="<%=EncodeForHtml(Pendiente)%>" size="12"/>
                  </span>
                
                <span class="header-note-currency">
                    <font class="ENCABEZADOR7" style="vertical-align:super;"><%=EncodeForHtml(rst("abreviatura"))%></font>
                </span>
                 <span class="header-note-buttonNote">
                     <img src="../images/<%=ImgAnotar%>" height="23" width="23" alt="<%=LitAnotCaja%>" title="<%=LitAnotCaja%>" onclick="Acaja('<%=enc.EncodeForJavascript(rst("nticket"))%>')" style="vertical-align:text-bottom;"/>
                 </span>                 
<%
                    rstAux.cursorlocation=3
				  	rstAux.Open "SELECT codigo,descripcion FROM Tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")'',adOpenKeyset, adLockOptimistic

                    DrawSelect "input-i_pago", "width:100px;", "i_pago", rstAux, session("ncliente") & "01", "codigo", "Descripcion","",""
					rstAux.Close
                    CloseDiv

                DrawDiv "header-bill", "", ""
                    DrawLabel "", "", LitDivisa
                    DrawSpan "", "", EncodeForHtml(rst("abreviatura")), ""
                CloseDiv

				Formulario="tickets"
				if mode="browse" or mode="editFuga" then
                    DrawDiv "header-bill", "", ""
                        DrawLabel "", "", LitOperador
                        operadorAux = trimCodEmpresa(rst("operador")) & " - "  & rst("nomoperador")
                        operadorAux = EncodeForHtml(operadorAux)
                        DrawSpan "", "", operadorAux, ""
                    CloseDiv

                    DrawDiv "header-client-iframe","",""
                    DrawLabel "","",LitCliente
					if rst("ncliente")>"" then
					    DrawSpan "", "", Hiperv(OBJClientes,rst("ncliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("ncliente")),LitVerCliente), ""
						DrawSpan "", "", EncodeForHtml(rst("rsocial")), ""
                    else
                        DrawSpan "", "", "", ""
					end if
                    CloseDiv
                                                    %></div><%
					''ricardo 21-7-2009 si el parametro mcb=1 se mostraran los costes/beneficio del documento
					if cstr(mcb)="1" then
                        DrawDiv "header-action","",""
						texto_beneficio="<a href=javascript:abrir_beneficios('tickets','" & rst("nticket") & "')><img src='../images/"&ImgCostesDoc&"' "&ParamImgCostesDoc&" alt='" & LitMsgVerCostesDoc & "'></a>"
                        DrawSpan "","",EncodeForHtml(texto_beneficio),""
                        CloseDiv
					end if
				end if
			%><table style="width: 100%;"></table>
        <div class="Section" id="S_DatosGenerales">
            <a href="#" rel="toggle[DatosGenerales]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader displayed">
                    <%=LITCABECERA%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
            <div class="SectionPanel" style="display: <%=iif(mode="browse" or mode="editfuga","","none")%>;" id="DatosGenerales">
                    <span id="cabecera">
                        <%
                            EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitTPV,"",EncodeForHtml(rst("tpv"))
                            if rst("nfactura") <> "" then
                                EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitFactura,"",Hiperv(OBJFacturasCli,rst("nfactura"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("nfactura")),LitVerFactura)
                            else
                                EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitFactura,"",""
                            end if
                            EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitCaja,"",EncodeForHtml(rst("caja"))
                            
                            DrawDiv "1", "CELDA", "medioPagoTicket"
                            DrawLabel "", "CELDA", LitTipoPago
                            DrawSpan "CELDA", "", EncodeForHtml(rst("tpago")), ""
                            CloseDiv

                            EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitAnulTicket,"",EncodeForHtml(trimCodEmpresa(rst("anulacion")))
                            EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitNLiquidacion,"",EncodeForHtml(trimCodEmpresa(rst("nliquidacion")))
                            if si_tiene_modulo_terminales<>0 then
                                if rst("nturno") & "">"" then
                                    EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitTurno,"",EncodeForHtml(mid(rst("nturno"),len(rst("nterminal"))+1,len(rst("nturno"))))
                                end if                                
                                EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitTerminal,"",EncodeForHtml(trimCodEmpresa(rst("nterminal")))
                            end if
                            %>

                    <div id="datosFuga">

                    <%DrawDiv "3-sub", "background-color: #eae7e3", ""
                    %> 
                    <label class="ENCABEZADOC" style="text-align:left"><%=LITDATOSFUGA%></label>
                    <%
                    CloseDiv%>
                                <%
									if (mode = "editFuga") then
                                        DrawDiv "1","",""
                                        DrawLabel "","",LITMATRICULA%><input class='CELDAR7' type="Text" name="matricula" value="<%=EncodeForHtml(rst("campo1"))%>" size="10" /><%CloseDiv
                                    else
                                        EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LITMATRICULA,0,EncodeForHtml(rst("campo1"))
									end if

									if (mode = "editFuga") then
                                        DrawDiv "1","",""
                                        DrawLabel "","",LITOBSERVACION%><textarea class='CELDAL7' name="observacion" style="width: 300px;" rows="2"><%=EncodeForHtml(rst("observaciones"))%></textarea><%CloseDiv
                                    else
                                        EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LITOBSERVACION,0,EncodeForHtml(rst("observaciones"))
									end if%>
                        </div>
					</span>
            </div>
        </div>
         <div class="Section" id="S_DATTICKETS" >
                <a href="#" rel="toggle[DATTICKETS]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader displayed">
                        <%=LITTITULO%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
                <div class="SectionPanel" id="DATTICKETS">
                  <div id="tabs" style="display:run-in" class="ui-tabs ui-widget ui-widget-content ui-corner-all">
                    <ul class="ui-tabs-nav ui-helper-reset ui-helper-clearfix ui-widget-header ui-corner-all">                       
                        <li class="ui-state-default ui-corner-top"><a href="#tabs1"><%=LitDetalles %></a></li>
                        <li class="ui-state-default ui-corner-top"><a href="#tabs2"><%=LitPagosCuenta %></a></li>                        
                    </ul>
             <div id="tabs1" class="overflowXauto" >
					<%if mode="browse" or mode="editFuga" then
						'** Campo oculto para controlar si el ticket está facturado.
                        rstAux.cursorlocation=3
						rstAux.Open "select nfactura from tickets with(nolock) where nticket='" & rst("nticket") & "'",session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
						if not isnull(rstAux("nfactura")) then%>
							<input type="hidden" name="h_nfactura" value="<%=EncodeForHtml(rstAux("nfactura"))%>" />
						<%else%>
							<input type="hidden" name="h_nfactura" value="NO"/>
						<%end if
						rstAux.close%>

							<%if mode="editFuga" then %>  
								<span id="DETALLES" style="display:none">
							<%else%>
								<span id="DETALLES" style="display:">
							<%end if%>

						   	<table class="width90 md-table-responsive bCollapse">
                                <tr>
                                    <td class='ENCABEZADOR underOrange width5'><%=LitItem%></td>
									<td class='ENCABEZADOR underOrange width5'><%=LitCantidad%></td>
									<td class='ENCABEZADOL underOrange width10'><%=LitReferencia%></td>
									<td class='ENCABEZADOL underOrange width10'><%=LitAlmacen%></td>
									<td class='ENCABEZADOL underOrange width15'><%=LitDescripcion%></td>
									<td class='ENCABEZADOR underOrange width10'><%=LitPVP%></td>
									<td class='ENCABEZADOR underOrange width5'><%=LitDto%></td>
									<td class='ENCABEZADOR underOrange width5'><%=LitTipoIva%></td>
									<td class='ENCABEZADOR underOrange width5'><%=LitImporte%></td>
                                </tr>
							</table>
							<iframe id='frDetalles' name="fr_Detalles" class="width90 iframe-data md-table-responsive" src='tickets_det.asp?nticket=<%=EncodeForHtml(rst("nticket"))%>' width='777' height='150' frameborder="yes" noresize="noresize"></iframe>
						</span>
                   </div>
                   <div id="tabs2" class="overflowXauto">
                       <span id="PAGOS_CUENTA">
                            <table class="width90 md-table-responsive bCollapse">
                                <tr>
                                    <td class='ENCABEZADOL width10 underOrange' ><%=LitNumPago%></td>
					                <td class='ENCABEZADOL width10 underOrange' ><%=LitFecha%></td>
					                <td class='ENCABEZADOL width10 underOrange' ><%=LitDescripcion%></td>
					                <td class='ENCABEZADOL width10 underOrange' ><%=LitImporte%></td>
					                <td class='ENCABEZADOL width10 underOrange' ><%=LitTipoPago%></td>
					                <td class='ENCABEZADOL width10 underOrange' >&nbsp</td>
                                </tr>
								<%
								if isnull(rst("nfactura")) then
									'Linea de inserción de un pago a cuenta
									%>
                                    <tr>
										<td class='CELDAL7 underOrange width10'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                        <td class="CELDAL7 underOrange width10">
											<input class='CELDAR7 width65' type="text" name="fechaPago" value="" onchange="cambiarfecha(document.tickets.fechaPago.value,'Fecha Pago')"/>
                                             <%DrawCalendar "fechaPago" %>
										</td>
                                        <td class="CELDAL7 underOrange width10">
											<textarea class='CELDAL7 width100' name="descripcionPago" onFocus="lenmensaje(this,0,50,'')" onKeydown="lenmensaje(this,0,50,'')" onKeyup="lenmensaje(this,0,50,'')" onBlur="lenmensaje(this,0,50,'')" rows="2"></textarea>
										</td>
                                        <td class="CELDAL7 underOrange width10">
											<input class='CELDAR7 width100' type="text" name="importePago" value="0" onchange="importepagoComp();"/>
										</td>
                                        <td class="CELDAL7 underOrange width10">
										<%
                                        rstSelect.cursorlocation=3
                                        rstSelect.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")'',adOpenKeyset,adLockOptimistic
                                        DrawSelect "'CELDAL7 width100'","","tipoPago",rstSelect,"","codigo","descripcion","",""
										rstSelect.close%>
                                        </td>
                                        <td class="CELDAL7 underOrange width10">
											<a href="javascript:addPago('<%=enc.EncodeForJavascript(nticket)%>');" onblur="javascript:document.tickets.fechaPago.focus();"><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo1%>" title="<%=LitNuevo1%>"/></a>
										</td>
                                    </tr>
									<%
								end if%>
							</table>
							<iframe id="frPagosCuenta" class="width90 iframe-data md-table-responsive" name="fr_PagosCuenta" src='tickets_pagos.asp?mode=browse&ndoc=<%=EncodeForHtml(rst("nticket"))%>' width='650' height='157' frameborder="yes" noresize="noresize"></iframe>
							<script language="javascript" type="text/javascript">Redimensionar();</script>
						</span>
                     </div>
					<%end if%>
                      </div>		
                </div>
            </div>
        <div class="Section" id="S_DATTOTAL" style="display: flow-root;" >
                    <div class="SectionHeader2">
                        <%=ucase(LitTotal)%>
                    </div>
            <div class="SectionPanel" id="DATTOTAL">
        <%if rst("ndecimales") & ""="" then
			ndecimales=n_decimales
		else
			ndecimales=rst("ndecimales")
		end if
		if rst("abreviatura") & ""="" then
			abreviatura=n_abreviatura
		else
			abreviatura=rst("abreviatura")
		end if

        DrawDiv "4", "", ""
        DrawLabel "", "", LitTotal
        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "", EncodeForHtml(abreviatura), ""
        CloseDiv

        DrawDiv "4", "", "BImponible"
        DrawLabel "", "", LitBImponible
        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "", EncodeForHtml(formatnumber(null_z(rst("base_imponible")),ndecimales,-1,0,iif(mode="browse",-1,0))), ""
        CloseDiv

        DrawDiv "4", "", "TotalIva"
        DrawLabel "", "", LitTotalIVA
        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "", EncodeForHtml(formatnumber(null_z(rst("total_iva")),ndecimales,-1,0,iif(mode="browse",-1,0))), ""
        CloseDiv

        DrawDiv "4", "", "TotalTicket"
        DrawLabel "", "", LitTotalTicket
        EligeCeldaResponsive1 "input", mode, "ENCABEZADOR", "", "", EncodeForHtml(formatnumber(null_z(rst("total_ticket")),ndecimales,-1,0,iif(mode="browse",-1,0))), ""
        CloseDiv%>
            </div>
        </div>
		<br/>
	<%elseif mode="search" then
    elseif mode="add" then
        %>
        <script language="javascript" type="text/javascript">
            window.onload = function () {
               var counter = 0;
               var interval1 = setInterval(function () {
                  SearchPage("tickets_lsearch.asp?mode=search&campo=&criterio=&texto=&viene=<%=enc.EncodeForJavascript(viene)%>", 1);
                   counter++;
                   if (counter == 1) {
                       clearInterval(interval1);
                   }
               }, 500);
            }
        </script>
        <%
    end if%>
</form>
    <%set rstAux=Nothing
	set rstAux2=Nothing
	set rst=Nothing
	set rstPed=Nothing
	set rstDetTicket=Nothing
	set rstIvas=Nothing
	set rstSelect=Nothing

	connRound.close
	set connRound = Nothing
end if%>
</body>
</html>