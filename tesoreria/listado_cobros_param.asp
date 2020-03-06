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
%>
<%
''	VGR	09/05/03	: Que aparezca siempre la fecha de factura y cambiar su literal a F.Emision Fra.
''	VGR	15/05/03	: Que salgan las facturas con importes negativos o cero.
''  JA 19/06/03: Migración monobase.
''	MPC 01/03/2007	: Añadir radio button Ticket TPV y montar la select para obtener los datos y mejorar el rendimiento
''	MPC 10/02/2010	: Se cambia el combo de las series para que sea de selección múltiple
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html lang="<%=session("lenguaje")%>">
<head>
<title><%=titulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>

<LINK REL="styleSHEET" href="../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../impresora.css" MEDIA="PRINT">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->

<!--#include file="cobros_param.inc" -->

<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

<%if request.querystring("viene")="tienda" or request.form("viene")="tienda" then
	titulo=LitTitulo2
else
	titulo=LitTituloList
end if%>

<%si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)
si_tiene_modulo_importaciones=ModuloContratado(session("ncliente"),ModImportaciones)%>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
function tier2Menu(objMenu,que)
{
    if (que=="vencimientos") document.listado_cobros_param.h_que.value="vencimientos";
    else document.listado_cobros_param.h_que.value="";
	if (objMenu.id=="albaranes")
	{
		if (albaranes.style.display == "none")
		{
			albaranes.style.display = "";
			facturas.style.display="none";
			tickets.style.display="none";
			ticketsOp.style.display="none";
			Resto.style.display="";
			document.listado_cobros_param.h_tabla.value="albaranes_cli";
			id_totalbmay.style.display="";
			document.getElementById("textoAgruparCom").innerHTML="<label><%=LitAgruparComercialModCom%></label>" + "<input type=\"checkbox\" class=\"CELDA\" name=\"agrupar_comercial\" <%=iif(agrupar_comercial="true","checked","")%> onclick=\"cambiar_comercial('1')\">";
        }
        else
		{
			albaranes.style.display = "none";
			facturas.style.display="";
			tickets.style.display="none";
			ticketsOp.style.display="none";
			Resto.style.display="";
			document.listado_cobros_param.h_tabla.value="facturas_cli";
			id_totalbmay.style.display="none";
			document.getElementById("textoAgruparCom").innerHTML="<label><%=LitAgruparComercialModCom%></label>" + "<input type=\"checkbox\" class=\"CELDA\" name=\"agrupar_comercial\" <%=iif(agrupar_comercial="true","checked","")%> onclick=\"cambiar_comercial('1')\">";
        }
    }
	if (objMenu.id=="facturas")
	{
		if (facturas.style.display == "none" || que=="facturas" || que=="vencimientos")
		{
			facturas.style.display = "";
			albaranes.style.display="none";
			tickets.style.display="none";
			ticketsOp.style.display="none";
			Resto.style.display="";
			//FLM:20090430:el control de la tabla se pasa a la función SonEfectos().
			//document.listado_cobros_param.h_tabla.value="facturas_cli";
			//SonEfectos();
			id_totalbmay.style.display="none";
			document.getElementById("textoAgruparCom").innerHTML="<label><%=LitAgruparComercialModCom%></label>" + "<input type=\"checkbox\" class=\"CELDA\" name=\"agrupar_comercial\" <%=iif(agrupar_comercial="true","checked","")%> onclick=\"cambiar_comercial('1')\">";
        }
        else
		{
			facturas.style.display = "none";
			albaranes.style.display="";
			tickets.style.display="none";
			ticketsOp.style.display="none";
			Resto.style.display="";
			document.listado_cobros_param.h_tabla.value="albaranes_cli";
			id_totalbmay.style.display="";
			document.getElementById("textoAgruparCom").innerHTML="<label><%=LitAgruparComercialModCom%></label>" + "<input type=\"checkbox\" class=\"CELDA\" name=\"agrupar_comercial\" <%=iif(agrupar_comercial="true","checked","")%> onclick=\"cambiar_comercial('1')\">";
        }
    }

	if (objMenu.id=="tickets")
	{
		if (tickets.style.display == "none")
		{
			facturas.style.display = "none";
			albaranes.style.display="none";
			tickets.style.display="";
			ticketsOp.style.display="";
			Resto.style.display="none";
			document.listado_cobros_param.h_tabla.value="facturas_cli";
			id_totalbmay.style.display="none";
			document.getElementById("textoAgruparCom").innerHTML="<label><%=LitAgruparOperador%></label>" + "<input type=\"checkbox\" class=\"CELDA\" name=\"agrupar_comercial\" <%=iif(agrupar_comercial="true","checked","")%> onclick=\"cambiar_comercial('1')\">";
        }
        else
		{
			facturas.style.display = "none";
			albaranes.style.display="none";
			tickets.style.display="";
			ticketsOp.style.display="";
			Resto.style.display="none";
			document.listado_cobros_param.h_tabla.value="tickets_cli";
			id_totalbmay.style.display="";
			document.getElementById("textoAgruparCom").innerHTML="<label><%=LitAgruparOperador%></label>" + "<input type=\"checkbox\" class=\"CELDA\" name=\"agrupar_comercial\" <%=iif(agrupar_comercial="true","checked","")%> onclick=\"cambiar_comercial('1')\">";
        }
    }
}

//FLM:20090428:Si se marca el check, se deben buscar efectos.
/*function SonEfectos(){
    if(document.listado_cobros_param.efectosPend.checked){
        document.listado_cobros_param.h_tabla.value="efectos_cli";
        document.listado_cobros_param.serie_efec.style.display="";
        document.listado_cobros_param.serie_fac.style.display="none";
    }
    else{
        document.listado_cobros_param.h_tabla.value="facturas_cli";
        document.listado_cobros_param.serie_efec.style.display="none";
        document.listado_cobros_param.serie_fac.style.display="";        
    }
        
}
*/

//Desencadena la búsqueda del proveedor cuya referencia se indica
function TraerProveedor(mode)
{
	document.location.href="listado_cobros_param.asp?nproveedor=" + document.listado_cobros_param.nproveedor.value + "&mode=" + mode + "&nserie=" + document.listado_cobros_param.nserie.value;
}

//Desencadena la búsqueda del cliente cuyo numero se indica
function TraerCliente(mode)
{
	document.location.href="listado_cobros_param.asp?ncliente=" + document.listado_cobros_param.ncliente.value + "&mode=" + mode +
			"&serie_fac=" + document.listado_cobros_param.serie_fac.value +
			"&serie_alb=" + document.listado_cobros_param.serie_alb.value +
			"&serie_efec=" + document.listado_cobros_param.serie_efec.value +
			"&serie_tic=" + document.listado_cobros_param.serie_tic.value
			+ "&comercial=" + document.listado_cobros_param.comercial.value
			+ "&agrupar_comercial=" + document.listado_cobros_param.agrupar_comercial.checked
			+ "&poblacion=" + document.listado_cobros_param.poblacion.value +
			"&agrupar_poblacion=" + document.listado_cobros_param.agrupar_poblacion.checked+
			"&actividad=" + document.listado_cobros_param.actividad.value +
			"&h_tabla=" + document.listado_cobros_param.h_tabla.value +
			"&h_que=" + document.listado_cobros_param.h_que.value +
			"&opcclientebaja=" + document.listado_cobros_param.opcclientebaja.checked +
			"&imptotalbmay=" + document.listado_cobros_param.imptotalbmay.value ;
}

function ver_documento(ndocumento,tipodocumento,ncliente)
{
	if (tipodocumento=="albaranes")
	{
		if (ncliente!="") document.location="../ventas/albaranes_cli_imp.asp?nalbaran=('" + ndocumento + "')&mode=browse&empresa="+ncliente.substr(0,5)+"&ncliente="+ncliente;
		else document.location="../ventas/albaranes_cli_imp.asp?nalbaran=('" + ndocumento + "')&mode=browse&empresa=<%=session("ncliente")%>";
	}
	else
	{
		if(tipodocumento=="facturas")
		{
			if (ncliente!="") document.location="../ventas/facturas_cli_imp.asp?nfactura=('" + ndocumento + "')&mode=browse&empresa="+ncliente.substr(0,5)+"&ncliente="+ncliente;
			else document.location="../ventas/facturas_cli_imp.asp?nfactura=('" + ndocumento + "')&mode=browse&empresa=<%=session("ncliente")%>";
		}
		else
		{
		    if(tipodocumento=="efectos")
		    {
		        if (ncliente!="") document.location="../search_layout.asp?pag1=netInic.asp?s=/ventas/efectos/efectos_cli.aspx&cod=" + ndocumento + "&titulo=<%=LitEfectoCli%>&mode=browse&empresa="+ncliente.substr(0,5)+"?ncliente="+ncliente+"&pag2=";			    
				else document.location="../search_layout.asp?pag1=netInic.asp&s=/ventas/efectos/efectos_cli.aspx&pag2=&cod=" + ndocumento + "&titulo=<%=LitEfectoCli%>&mode=browse&empresa=<%=session("ncliente")%>";
		    }
		    else
		    {
			    if (ncliente!="") document.location="../ventas/ticket_imp.asp?nticket=('" + ndocumento + "')&mode=browse&empresa="+ncliente.substr(0,5)+"&ncliente="+ncliente;
			    else document.location="../ventas/ticket_imp.asp?nticket=('" + ndocumento + "')&mode=browse&empresa=<%=session("ncliente")%>";
			}
		}
	}
	parent.parent.topFrame.document.getElementById("regresar").style.display="";
}

function tratar_poblacion(modo)
{
	if (modo=="1") document.listado_cobros_param.agrupar_poblacion.checked=false;
	if (modo=="2") document.listado_cobros_param.poblacion.value="";
}

function cambiar_comercial(modo)
{
	if (modo=="1")
	{
		if (document.listado_cobros_param.agrupar_comercial.checked==true) document.listado_cobros_param.comercial.value="";
	}
	if (modo=="2") document.listado_cobros_param.agrupar_comercial.checked=false;
}

function tratar_imptotalb()
{
	document.listado_cobros_param.imptotalbmay.value=document.listado_cobros_param.imptotalbmay.value.replace(".",",");
	if (isNaN(document.listado_cobros_param.imptotalbmay.value.replace(",",".")))
	{
	<%if (si_tiene_modulo_importaciones<>0) then%>
		alert("<%=LitImpTotEmbDebNum%>");
	<%else%>
		alert("<%=LitImpTotAlbDebNum%>");
	<%end if%>
	}
}


</script>

<body onload="self.status='';" class="BODY_ASP">
<%'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
'campo: Nombre del campo con el cual se realizará la búsqueda
'criterio: Tipo de búsqueda
'texto: Texto a buscar.
function CadenaBusquedaTienda(campo,criterio,texto,modo)
	if modo=1 then
		texto_aux="facturas_cli."
	elseif modo=2 then
		texto_aux="albaranes_cli."
	elseif modo=3 then
		texto_aux="tickets."
	elseif modo=4 then
		texto_aux="efectos_cli."
	end if

	if texto > "" then
		select case criterio
			case "contiene"
				CadenaBusquedaTienda=" where " + texto_aux + campo + " like '%" + texto + "%' and"
			case "empieza"
				CadenaBusquedaTienda=" where " + texto_aux + campo + " like '" + texto + "%' and"
			case "termina"
				CadenaBusquedaTienda=" where " + texto_aux + campo + " like '%" + texto + "' and"
			case "igual"
				CadenaBusquedaTienda=" where " + texto_aux + campo + "='" + texto + "' and"
		end select
	else
		CadenaBusquedaTienda=" where "
	end if
end function

'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
	'ncliente:
function CadenaBusqueda(ncliente,nserie)
	CadenaBusqueda = ""

	if ncliente > "" then
		CadenaBusqueda = " where ncliente='" & ncliente & "' and"
	end if
	if nserie > "" then
		CadenaBusqueda = CadenaBusqueda + " serie='" & nserie & "' and"
	end if
	CadenaBusqueda = CadenaBusqueda + " cobrada = 0 order by nfactura,fecha desc,ncliente"
end function

'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
borde=0
'FLM:20090430: esta variable controla si se debe ejecutar la sql del listado.
noEjecutesSQL=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
%>
<form name="listado_cobros_param" method="post">
<%PintarCabecera "listado_cobros_param.asp"
	'Leer parámetros de la página'
	' No hace falta comprobar con checkCadena puesto que en la consulta sql se utiliza el session("ncliente")'
	mode = EncodeForHtml(Request.QueryString("mode"))
	if ucase(mode) = "BROWSE" then mode ="imp"

	ncliente	= limpiaCadena(Request.QueryString("ncliente"))
	if ncliente ="" then
		ncliente	= limpiaCadena(Request.form("ncliente"))
	end if
	if ncliente="" then
		ncliente	= limpiaCadena(Request.QueryString("h_ncliente"))
		if ncliente ="" then
			ncliente	= limpiaCadena(Request.form("h_ncliente"))
		end if
	end if

	opcclientebaja	= limpiaCadena(Request.QueryString("opcclientebaja"))
	if opcclientebaja="" then
		opcclientebaja	= limpiaCadena(Request.form("opcclientebaja"))
	end if

	actividad	= limpiaCadena(Request.QueryString("actividad"))
	if actividad ="" then
		actividad	= limpiaCadena(Request.form("actividad"))
	end if

	if request.form("serie_alb")>"" then
		nserie=limpiaCadena(request.form("serie_alb"))
	else
		nserie=limpiaCadena(request.querystring("serie_alb"))
	end if
	
	if nserie="" then
		if request.form("serie_fac")>"" then
			nserie=limpiaCadena(request.form("serie_fac"))
		else
			nserie=limpiaCadena(request.querystring("serie_fac"))
		end if

		if nserie="" then
			if request.form("serie_tic")>"" then
				nserie=limpiaCadena(request.form("serie_tic"))
			else
				nserie=limpiaCadena(request.querystring("serie_tic"))
			end if
		end if
	end if
	
	'FLM:20090430:capturo las series de efectos si es una factura.
    if request.form("serie_efec")>"" then
        nserieEfec=replace(limpiaCadena(request.form("serie_efec"))," ","")
    elseif nserieEfec="" and request.form("nserieEfec")>"" then
        nserieEfec=replace(limpiaCadena(request.form("nserieEfec"))," ","")
    else
        nserieEfec=replace(limpiaCadena(request.querystring("serie_efec"))," ","")
     end if		        

	tabla=limpiaCadena(request.form("Documento"))

	if tabla>"" then
	    'FLM:20090430:si viene un documento y es facturas_cli tengo que verificar si en h_tabla viene efectos_cli
	    'if tabla="facturas_cli" and request.form("h_tabla")="efectos_cli" then
	    '    tabla="efectos_cli"
	    'end if
	else
		if request.form("h_tabla")>"" then
			tabla=limpiaCadena(request.form("h_tabla"))
		else
			tabla=limpiaCadena(request.querystring("h_tabla"))
		end if
	end if
	
	if request.form("h_que")>"" then
		h_que=limpiaCadena(request.form("h_que"))
	else
		h_que=limpiaCadena(request.querystring("h_que"))
	end if

    if h_que="vencimientos" and tabla="facturas_cli" then
        tabla="vencimientos"
    end if

	if request.form("comercial")>"" then
		comercial=limpiaCadena(request.form("comercial"))
	else
		comercial=limpiaCadena(request.querystring("comercial"))
	end if

	if comercial=", " then comercial=""
	pos=InStr(1,comercial,", ")
	if pos=1 then
		comercial = trim(right(comercial,len(comercial)-pos))
	elseif pos<>0 then
		comercial = trim(left(comercial,pos-1))
	end if

	if request.form("agrupar_comercial")>"" then
		agrupar_comercial=limpiaCadena(request.form("agrupar_comercial"))
	else
		agrupar_comercial=limpiaCadena(request.querystring("agrupar_comercial"))
	end if
	
	'CCA 09-01-2008: se añade la posibilidad de listado apaisado
	apaisado=iif(limpiaCadena(request.form("apaisado"))>"","SI","")

	if request.form("poblacion")>"" then
		poblacion=limpiaCadena(request.form("poblacion"))
	else
		poblacion=limpiaCadena(request.querystring("poblacion"))
	end if
	if poblacion & "">"" then poblacion=UCase(poblacion)

	if request.form("agrupar_poblacion")>"" then
		agrupar_poblacion=limpiaCadena(request.form("agrupar_poblacion"))
	else
		agrupar_poblacion=limpiaCadena(request.querystring("agrupar_poblacion"))
	end if

	if request.form("imptotalbmay")>"" then
		imptotalbmay=limpiaCadena(request.form("imptotalbmay"))
	else
		imptotalbmay=limpiaCadena(request.querystring("imptotalbmay"))
	end if

	viene	= limpiaCadena(Request.QueryString("viene"))
	if viene="" then
		viene	= limpiaCadena(Request.form("viene"))
	end if

''ricardo 3-12-2007 si viene de la tienda , por defecto veremos las facturas
	if viene="tienda" and tabla="" then tabla="facturas_cli"
	if tabla="" then tabla="facturas_cli"
	
	' IML 28/04/2004 : Validamos si el usuario tiene acceso
	if viene="tienda" then
		sesionNCliente=left(ncliente,5)&""
		if sesionNCliente&""="" then sesionNCliente=session("ncliente")
		checkAccesoTienda sesionNCliente,ncliente,""
		ncliente=trimCodEmpresa(ncliente)
	else
		sesionNCliente=session("ncliente")
	end if
	' FIN IML 28/04/2004 : Validamos si el usuario tiene acceso

	campo=limpiaCadena(Request.querystring("campo"))
	criterio=limpiaCadena(Request.querystring("criterio"))
	texto=limpiaCadena(Request.querystring("texto"))
	p_lote=limpiaCadena(Request.QueryString("lote"))
	p_sentido=limpiaCadena(Request.QueryString("sentido"))

	NumRegs	= limpiaCadena(Request.QueryString("NumRegs"))
	if NumRegs ="" then
		NumRegs	= limpiaCadena(Request.form("NumRegs"))
	end if

	''ricardo 21-3-2005
	DFecha=limpiaCadena(Request.Form("Dfecha"))
	if DFecha & ""="" then DFecha=limpiaCadena(request.querystring("Dfecha"))
	HFecha=limpiaCadena(Request.Form("Hfecha"))
	if Hfecha & ""="" then Hfecha=limpiaCadena(request.querystring("Hfecha"))%>

	<input type="hidden" name="h_tabla" value="<%=EncodeForHtml(tabla)%>">
	<input type="hidden" name="h_que" value="<%=EncodeForHtml(h_que)%>">

	<%set rstAux = Server.CreateObject("ADODB.Recordset")
	moneda_base = d_lookup("codigo","divisas","codigo like '" & sesionNCliente & "%' and moneda_base=1",session("backendlistados"))
	n_decimalesMB=d_lookup("ndecimales","divisas","codigo like '" & sesionNCliente & "%' and moneda_base<>0",session("backendlistados"))
	abreviaturaMB=d_lookup("abreviatura","divisas","codigo like '" & sesionNCliente & "%' and moneda_base<>0",session("backendlistados"))

	MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='048'", DSNIlion)
	MAXPDF=d_lookup("maxpdf", "limites_listados", "item='048'", DSNIlion)%>
	<input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>
	<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'>
	<input type='hidden' name='maxmb' value='<%=EncodeForHtml(moneda_base)%>'>
	<%strwhere=""
	strwhere2=""
	strwhereEfc=""

	Alarma "listado_cobros_param.asp"%>
	<hr/>
	<%set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rstCliente = Server.CreateObject("ADODB.Recordset")
	set rstVencimientos = Server.CreateObject("ADODB.Recordset")

	if mode="select1"then%>
		<input type="hidden" name="nserie" value="<%=EncodeForHtml(nserie)%>">
		<%
        DrawDiv "3", "", ""
			if si_tiene_modulo_importaciones<>0 then
            DrawDiv "6", "", ""
            DrawLabel "", "", LitEmbarques
            %>
				<span class="CELDA"><input type="radio" name="Documento" value="albaranes_cli" <%=iif(tabla="albaranes_cli","checked","")%> onclick="tier2Menu(albaranes,'')" ></span>
			<%
            CloseDiv
            else
            DrawDiv "6", "", ""
            DrawLabel "", "", LitAlbaranes%>
				<span class="CELDA"><input type="radio" name="Documento" value="albaranes_cli" <%=iif(tabla="albaranes_cli","checked","")%> onclick="tier2Menu(albaranes,'')" ></span>
			<%
            CloseDiv    
            end if
            DrawDiv "6", "", ""
            DrawLabel "", "", LitFacturas
                %>
				<span class="CELDA"><input type="radio" name="Documento" value="facturas_cli" <%=iif(tabla="" or tabla="facturas_cli","checked","")%> onclick="tier2Menu(facturas,'facturas')"></span>
            <%CloseDiv
                DrawDiv "6", "", ""
                DrawLabel "", "", LitVencimientos
                %>
				<span class="CELDA"><input type="radio" name="Documento" value="vencimientos" <%=iif(tabla="vencimientos","checked","")%> onclick="tier2Menu(facturas,'vencimientos')"></span>
                <%CloseDiv
                  DrawDiv "6", "", ""
                  DrawLabel "", "", LitTicketsTPV
                    %>
				<span class="CELDA"><input type="radio" name="Documento" value="tickets_cli" <%=iif(tabla="tickets_cli","checked","")%> onclick="tier2Menu(tickets,'')"></span>
			<%CloseDiv
        CloseDiv%>
        <hr/>
    	<%
            diaHoy = day(date)
			mesHoy=month(date)
			fechaHoy=iif(Len(diaHoy)>1,diaHoy,"0"&diaHoy)&"/"&iif(Len(mesHoy)>1,mesHoy,"0"&mesHoy)&"/"&year(date)
            
            EligeCelda "input", "add", "", "", "", 0, LitDesdeFecha, "Dfecha", "", EncodeForHtml(iif(Dfecha>"",Dfecha,"01/01/" & year(date)))
            DrawCalendar "Dfecha"
            EligeCelda "input", "add", "", "", "", 0, LitHastaFecha, "Hfecha", "", EncodeForHtml(fechahoy)
            DrawCalendar "Hfecha"

				if ncliente >"" then
					ncliente=sesionNCliente & Completar(ncliente,5,"0")
					nom_cliente=d_lookup("rsocial","clientes","ncliente='" & ncliente & "'",session("backendlistados"))
				else
					nom_cliente=""
				end if
				DrawDiv "1", "", ""
                DrawLabel "", "", LitCodigo%><input class='width15' type="text" name="ncliente" value="<%=EncodeForHtml(trimCodEmpresa(ncliente))%>" size="8" onchange="TraerCliente('<%=enc.EncodeForJavascript(mode)%>');"><a class='CELDAREFB' href="javascript:AbrirVentana('../ventas/clientes_buscar.asp?ndoc=listado_cobros_param&titulo=<%=LitSelCliente%>&mode=search&viene=listado_cobros_param','P','<%=AltoVentana%>','<%=AnchoVentana%>')"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt=""></a><input class="width40" type="text" name="nombre" size="18" class="CELDA" value="<%=EncodeForHtml(nom_cliente)%>"><%
                CloseDiv
                  
				rstSelect.open "select codigo,descripcion from tipo_actividad with(nolock) where codigo like '" & sesionNCliente & "%' order by descripcion",session("backendlistados"),adOpenKeyset,adLockOptimistic
				DrawSelectCelda "CELDA colspan='3'","175","",0,LitActividad,"actividad",rstSelect,actividad,"codigo","descripcion","",""
				rstSelect.close               
                EligeCelda "check", "add", "", "", "", 0, LitClienteBaja, "opcclientebaja", "", iif(opcclientebaja="true","checked","")%><span id="Resto" style="display:<%=iif(tabla="albaranes_cli" or tabla="facturas_cli" or tabla="vencimientos","","none")%>"><%
					rstAux.open "select dni, nombre from comerciales c with(nolock) left outer join personal p with(nolock) on c.comercial=p.dni where dni like '" & sesionNCliente & "%' and c.fbaja is null",session("backendlistados"),adOpenKeyset,adLockOptimistic
                    if si_tiene_modulo_comercial<>0 then
                        DrawSelectCelda "CELDA colspan='3'","175","",0,LitComercialModCom,"comercial",rstAux,comercial,"dni","nombre","onchange","cambiar_comercial('2')"
					else
                        DrawSelectCelda "CELDA colspan='3'","175","",0,Litcomercial,"comercial",rstAux,comercial,"dni","nombre","onchange","cambiar_comercial('2')"
					end if
					rstAux.close%></span><span id="ticketsOp" style="display:<%=iif(tabla="tickets_cli","","none")%>"><%
							    rstAux.open "select dni, nombre from personal with(nolock) where dni like '" & sesionNCliente & "%'",session("backendlistados"),adOpenKeyset,adLockOptimistic
							    DrawSelectCelda "CELDA","175","",0,LitOperador,"comercial",rstAux,"","dni","nombre","",""
							    rstAux.close%></span><%
                DrawDiv "1", "", "textoAgruparCom"
                if si_tiene_modulo_comercial<>0 then
                    DrawLabel "", "", LitAgruparComercialModCom
			    else
					DrawLabel "", "", LitAgruparComercial
				end if%><input type="checkbox" name="agrupar_comercial" <%=iif(agrupar_comercial="true","checked","")%> onclick="cambiar_comercial('1')"><%
                CloseDiv
			DrawDiv "1", "", ""
            DrawLabel "", "", LitPoblacion%><input type="text" size="20" name="poblacion" value="<%=EncodeForHtml(poblacion)%>" onchange="tratar_poblacion('1')"><a class='CELDAREFB'  href="javascript:AbrirVentana('../configuracion/poblaciones.asp?mode=buscar&viene=listado_cobros_param&titulo=<%=LitSelPoblacion%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPoblaciones%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			<%CloseDiv
			DrawDiv "1", "", ""
            DrawLabel "", "", LitAgruparPoblacion%><input type="checkbox" name="agrupar_poblacion" <%=iif(agrupar_poblacion="true","true","false")%> onclick="tratar_poblacion('2')"><%
            CloseDiv%><span id="facturas" style="display:<%=iif(tabla="" or tabla="facturas_cli" or tabla="vencimientos","","none")%>"><%
            rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & sesionNCliente & "%' and tipo_documento ='EFECTO CLIENTE'",session("backendlistados"),adOpenKeyset,adLockOptimistic
			DrawSelectMultipleCelda "","","",0,LitIncEfectosPend,"serie_efec",rstAux,iif(nserie>"" and (tabla="facturas_cli" or tabla="vencimientos"),nserie,""),"nserie","descripcion","",""
			rstAux.close	
			DrawDiv "1", "",""
            DrawLabel "", "",LitSerie
            rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & sesionNCliente & "%' and tipo_documento ='FACTURA A CLIENTE'",session("backendlistados"),adOpenKeyset,adLockOptimistic%><select name="serie_fac"  multiple="multiple" size="5" class="width60">
			        <%while not rstAux.eof%>
			             <option value="<%=EncodeForHtml(rstAux("nserie"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
			            <%rstAux.movenext
			           wend%><option value=""></option>
			        </select><%
            rstAux.close	
            CloseDiv
			'CCA 09-01-2007: Añade la posibilidad de escoger un formato de impresión apaisado para el listado de cobros pendientes
			EligeCelda "check", "add", "", "", "", 0, LitApaisadoListCobrosParam, "apaisado", "", ""%></span><span id="albaranes" style="display:<%=iif(tabla="albaranes_cli","","none")%>"><%
                DrawDiv "1", "", ""
                DrawLabel "", "", LitSerie
				rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & sesionNCliente & "%' and tipo_documento ='ALBARAN DE SALIDA'",session("backendlistados"),adOpenKeyset,adLockOptimistic
                %><select name="serie_alb"  multiple="multiple" size="5" class="CELDA" style='width:175px'><%
                    while not rstAux.eof%>
				     <option value="<%=EncodeForHtml(rstAux("nserie"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
				    <%rstAux.movenext
				   wend%>
				    <option value=""></option>
				</select><%rstAux.close
				CloseDiv%></span><span id="tickets" style="display:<%=iif(tabla="tickets_cli","","none")%>"><%
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitSerie
					rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & sesionNCliente & "%' and tipo_documento ='TICKET'",session("backendlistados"),adOpenKeyset,adLockOptimistic%><select name="serie_tic"  multiple="multiple" size="5" class="CELDA" style='width:175px'>
				        <%while not rstAux.eof%>
				             <option value="<%=EncodeForHtml(rstAux("nserie"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
				            <%rstAux.movenext
				           wend%>
				            <option value=""></option>
				        </select><%rstAux.close
                    CloseDiv%></span><span id="id_totalbmay" style="display:<%=iif(tabla="albaranes_cli","","none")%>"><%
                    DrawDiv "1", "", ""
                    if si_tiene_modulo_importaciones<>0 then
					    DrawLabel "", "", LitImpTotalEmbarque
				    else
				        DrawLabel "", "", LitImpTotalAlbaran
				    end if%><input class="CELDA" type="text" size="10" name="imptotalbmay" value="<%=iif(imptotalbmay="","0",imptotalbmay)%>" onchange="tratar_imptotalb()"><%CloseDiv%></span>
                    <hr/>
	<%elseif mode="imp" then
		''ricardo 25-5-2006 comienzo de la select
		''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
		auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"inicio_listado_cobros"%>
		<input type="hidden" name="h_ncliente" value="<%=EncodeForHtml(iif(viene="tienda",sesionNCliente&ncliente,ncliente))%>">
		<%if (tabla="facturas_cli" or tabla="vencimientos") then%>
			<input type="hidden" name="serie_fac" value="<%=EncodeForHtml(nserie)%>">
		<%elseif tabla="albaranes_cli" then%>
			<input type="hidden" name="serie_alb" value="<%=EncodeForHtml(nserie)%>">
		<%elseif tabla="efectos_cli" then%>
			<input type="hidden" name="serie_efe" value="<%=EncodeForHtml(nserie)%>">
		<%else%>
			<input type="hidden" name="serie_tic" value="<%=EncodeForHtml(nserie)%>">
		<%end if%>
		<input type="hidden" name="actividad" value="<%=EncodeForHtml(actividad)%>" >
		<input type="hidden" name="comercial" value="<%=EncodeForHtml(comercial)%>" >
		<input type="hidden" name="agrupar_comercial" value="<%=EncodeForHtml(agrupar_comercial)%>" >
		<input type="hidden" name="poblacion" value="<%=EncodeForHtml(poblacion)%>" >
		<input type="hidden" name="agrupar_poblacion" value="<%=EncodeForHtml(agrupar_poblacion)%>" >
		<input type="hidden" name="imptotalbmay" value="<%=EncodeForHtml(imptotalbmay)%>" >
		<input type="hidden" name="dfecha" value="<%=EncodeForHtml(dfecha)%>" >
		<input type="hidden" name="hfecha" value="<%=EncodeForHtml(hfecha)%>" >
		<input type="hidden" name="nserieEfec" value="<%=EncodeForHtml(nserieEfec) %>" />
		
		<!-- CCA 09-01-2007: Añadido para pasar la información de si se desea el formato apaisado o no -->
		<input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>">

		<%if viene="tienda" then%>
			<input type="hidden" name="viene" value="<%=EncodeForHtml(viene)%>">
			<input type="hidden" name="campo" value="<%=EncodeForHtml(campo)%>">
			<input type="hidden" name="criterio" value="<%=EncodeForHtml(criterio)%>">
			<input type="hidden" name="texto" value="<%=EncodeForHtml(texto)%>">
		<%end if

		strwhere=CadenaBusquedaTienda(campo,criterio,texto,1)
		strwhereVenc=CadenaBusquedaTienda(campo,criterio,texto,1)
		strwhere2=CadenaBusquedaTienda(campo,criterio,texto,2)
		strwhere6=CadenaBusquedaTienda(campo,criterio,texto,3)
		strwhereEfc=CadenaBusquedaTienda(campo,criterio,texto,4)

		if comercial>"" then
			strwhere=strwhere & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = facturas_cli.ncliente and divisas.codigo=facturas_cli.divisa and"
			strwhereVenc=strwhereVenc & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = facturas_cli.ncliente and divisas.codigo=facturas_cli.divisa and"
			strwhere2=strwhere2 & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = albaranes_cli.ncliente and divisas.codigo=albaranes_cli.divisa and"
			strwhere6=strwhere6 & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = tickets.ncliente and divisas.codigo=tickets.divisa and"
			strwhereEfc=strwhereEfc & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = efectos_cli.ncliente and divisas.codigo=efectos_cli.divisa and"
		else
			strwhere=strwhere & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = facturas_cli.ncliente and divisas.codigo=facturas_cli.divisa and"
			strwhereVenc=strwhereVenc & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = facturas_cli.ncliente and divisas.codigo=facturas_cli.divisa and"
			strwhere2=strwhere2 & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = albaranes_cli.ncliente and divisas.codigo=albaranes_cli.divisa and"
			strwhere6=strwhere6 & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = tickets.ncliente and divisas.codigo=tickets.divisa and"
			strwhereEfc=strwhereEfc & " clientes.ncliente like ''" & sesionNCliente & "%'' and clientes.ncliente = efectos_cli.ncliente and divisas.codigo=efectos_cli.divisa and"
		end if
		if tabla="facturas_cli" then%>
			<font class="cab"><b><%=LitDocumento%>: </b><%=LitFacturas%></font><br/>
		<%elseif tabla="vencimientos" then%>
			<font class="cab"><b><%=LitDocumento%>: </b><%=LitVencimientos%></font><br/>
		<%elseif tabla="albaranes_cli" then
			if si_tiene_modulo_importaciones<>0 then%>
				<font class="cab"><b><%=LitDocumento%>: </b><%=LitEmbarques%></font><br/>
			<%else%>
				<font class="cab"><b><%=LitDocumento%>: </b><%=LitAlbaranes%></font><br/>
			<%end if		
		else%>
			<font class="cab"><b><%=LitDocumento%>: </b><%=LitTicketsTPV%></font><br/>
		<%end if
		if dfecha & "">"" then
			strwhere=strwhere & " facturas_cli.fecha>=''" & dfecha & " 00:00:00'' and"
			if tabla="facturas_cli" then
			    strwhereVenc=strwhereVenc & " facturas_cli.fecha>=''" & dfecha & " 00:00:00'' and"
			else
			    strwhereVenc=strwhereVenc & " vencimientos_salida.fechav>=''" & dfecha & " 00:00:00'' and"
			end if
			strwhere2=strwhere2 & " albaranes_cli.fecha>=''" & dfecha & " 00:00:00'' and"
			strwhere6=strwhere6 & " tickets.fecha>=''" & dfecha & " 00:00:00'' and"
			strwhereEfc=strwhereEfc & " isnull(efectos_cli.fechavto,efectos_cli.fecha)>=''" & dfecha & "'' and"%>
			<font class="cab"><b><%=LitDesdeFecha%>:&nbsp;</b></font><font class="cab"><%=EncodeForHtml(dfecha)%></font><br/>
		<%end if
		if hfecha & "">"" then
			strwhere=strwhere & " facturas_cli.fecha<=''" & hfecha & " 23:59:00'' and"
			if tabla="facturas_cli" then
			    strwhereVenc=strwhereVenc & " facturas_cli.fecha<=''" & hfecha & " 23:59:00'' and"
			else
			    strwhereVenc=strwhereVenc & " vencimientos_salida.fechav<=''" & hfecha & " 23:59:00'' and"
			end if
			strwhere2=strwhere2 & " albaranes_cli.fecha<=''" & hfecha & " 23:59:00'' and"
			strwhere6=strwhere6 & " tickets.fecha<=''" & hfecha & " 23:59:00'' and"
			strwhereEfc=strwhereEfc & " isnull(efectos_cli.fechavto,efectos_cli.fecha)<=''" & hfecha & "'' and"%>
			<font class="cab"><b><%=LitHastaFecha%>:&nbsp;</b></font><font class="cab"><%=EncodeForHtml(hfecha)%></font><br/>
		<%end if
		if ncliente > "" then
			ncliente=sesionNCliente & ncliente
			strwhere = strwhere & " facturas_cli.ncliente=''" & ncliente & "'' and"
			strwhereVenc = strwhereVenc & " facturas_cli.ncliente=''" & ncliente & "'' and"         
			%>
			<font class="cab"><b><%=LitCliente%>:&nbsp;</b></font><font class="cab"><%=EncodeForHtml(trimCodEmpresa(ncliente))%>&nbsp;<%=EncodeForHtml(d_lookup("rsocial","clientes","ncliente='" & ncliente & "'",session("backendlistados")))%></font><br/>
			<%strwhere2 = strwhere2 & " albaranes_cli.ncliente=''" & ncliente & "'' and"
			strwhere6 = strwhere6 & " tickets.ncliente=''" & ncliente & "'' and"
			strwhereEfc = strwhereEfc & " efectos_cli.ncliente=''" & ncliente & "'' and"
		else%>
			<input type="hidden" name="opcclientebaja" value="<%=EncodeForHtml(opcclientebaja)%>" >
			<%if opcclientebaja="" then
				strbaja=" "
			else
				strbaja=" clientes.fbaja is null"
				strwhere = strwhere & strbaja & " and"
				strwhereVenc = strwhereVenc & strbaja & " and"
				strwhere2 = strwhere2 & strbaja & " and"
				strwhere6 = strwhere6 & strbaja & " and"
				strwhereEfc = strwhereEfc & strbaja & " and"%>
				<font class=cab><b><%=LitClienteBaja%></b></font><br/>
			<%end if
		end if
		if actividad>"" then
			strwhere = strwhere & " clientes.tactividad=''" & actividad & "'' and"
			strwhereVenc = strwhereVenc & " clientes.tactividad=''" & actividad & "'' and"
			strwhere2 = strwhere2 & " clientes.tactividad=''" & actividad & "'' and"
			strwhere6 = strwhere6 & " clientes.tactividad=''" & actividad & "'' and"
			strwhereEfc = strwhereEfc & " clientes.tactividad=''" & actividad & "'' and"%>
			<font class=cab><b><%=LitActividad%>:&nbsp;</b></font><font class=cab><%=EncodeForHtml(d_lookup("descripcion","tipo_actividad","codigo='" & actividad & "'",session("backendlistados")))%></font><br/>
		<%end if
		
		strwhere7=""
		if nserie>"" then
		    'FLM:20090430:adaptamos las series para un múltiple selección.
			strwhere= strWhere & " FACTURAS_CLI.serie in (''" & replace(replace(nserie," ",""),",","'',''") & "'') and"
			strwhereVenc= strwhereVenc & " FACTURAS_CLI.serie in (''" & replace(replace(nserie," ",""),",","'',''") & "'') and"
			strwhere2= strWhere2 & " serie in (''" & replace(replace(nserie," ",""),",","'',''") & "'') and"
			strwhere6= strWhere6 & " serie in (''" & replace(replace(nserie," ",""),",","'',''") & "'') and"
			strwhere7= " fce.serie in (''" & replace(replace(nserie," ",""),",","'',''") & "'') and "			
		end if
		'FLM:20090819:filtro la serie por el valor de nserieEfec
		if nserieEfec&"">"" then
			strwhereEfc = strwhereEfc & " efectos_cli.serie in (''" & replace(nserieEfec,",","'',''") & "'') and"
		end if
		if nserie&"">"" then
		    rstAux.open "select substring(nserie,6,len(nserie)) + ' - ' + nombre as nombre from series with(nolock) where nserie like '" & sesionNCliente & "%' and nserie in ('" & replace(replace(nserie," ",""),",","','") & "')",session("backendlistados")
		    listaSerie=""
		    while not rstAux.eof  
	            listaSerie= listaSerie&rstAux("nombre")&","
	            rstAux.moveNext		          
	        wend 
	        rstAux.Close%>
			<font class="cab"><b><%=LitSerie%>:&nbsp;</b></font><font class="cab"><%=EncodeForHtml(left(listaSerie,len(listaSerie)-1))%></font><br/>
		<%end if
		
		'FLM:20090820:series de efectos
		if (tabla="facturas_cli" or tabla="vencimientos")  and nserieEfec&"">"" then
		    rstAux.open "select substring(nserie,6,len(nserie)) + ' ' + nombre as nombre from series with(nolock) where nserie like '" & sesionNCliente & "%' and nserie in ('" & replace(replace(nserieEfec," ",""),",","','") & "')",session("backendlistados")
		    listaSerie=""
		    while not rstAux.eof  
	            listaSerie= listaSerie&rstAux("nombre")&","
	            rstAux.moveNext		          
	        wend 
	        rstAux.Close() 
		     if listaSerie>"" then %>
		        <font class="cab"><b><%=LitSerieEfec%>:&nbsp;</b></font>
		        <font class="cab"><%=EncodeForHtml(left(listaSerie,len(listaSerie)-1))%></font><br/>
		    <%end if%>
		<%end if
		if agrupar_poblacion>"" then
			poblacion=""
			strwhere= strWhere & " clientes.dir_principal=domicilios.codigo and"
			strwhereVenc= strwhereVenc & " clientes.dir_principal=domicilios.codigo and"
			strwhere2= strWhere2 & " clientes.dir_principal=domicilios.codigo and"
			strwhere6= strWhere6 & " clientes.dir_principal=domicilios.codigo and"
			strwhereEfc = strwhereEfc & " clientes.dir_principal=domicilios.codigo and"
		else
			if poblacion>"" then
				strwhere= strWhere & " domicilios.tipo_domicilio=''PRINCIPAL_CLI'' and domicilios.poblacion like ''%" & poblacion & "%'' and clientes.dir_principal=domicilios.codigo and"
				strwhereVenc= strwhereVenc & " domicilios.tipo_domicilio=''PRINCIPAL_CLI'' and domicilios.poblacion like ''%" & poblacion & "%'' and clientes.dir_principal=domicilios.codigo and"
				strwhere2= strWhere2 & " domicilios.tipo_domicilio=''PRINCIPAL_CLI'' and domicilios.poblacion like ''%" & poblacion & "%'' and clientes.dir_principal=domicilios.codigo and"
				strwhere6= strWhere6 & " domicilios.tipo_domicilio=''PRINCIPAL_CLI'' and domicilios.poblacion like ''%" & poblacion & "%'' and clientes.dir_principal=domicilios.codigo and"
				strwhereEfc = strwhereEfc& " domicilios.tipo_domicilio=''PRINCIPAL_CLI'' and domicilios.poblacion like ''%" & poblacion & "%'' and clientes.dir_principal=domicilios.codigo and"%>
				<font class="cab"><b><%=LitPoblacion%>:&nbsp;</b></font><font class="cab"><%=EncodeForHtml(poblacion)%></font><br/>
			<%end if
		end if

		strwhere3=""
		strwhere33=""
		strwhere34=""
		strwhere35=""
		strwhere36=""
''		strwhere4=""
		strwhere41=""
		strwhere5=""
		strwhere51=""
		strwhere4Venc=""
		if agrupar_comercial & ""="" then
			if comercial>"" then
				if (tabla="facturas_cli" or tabla="vencimientos") then
					strwhere= strWhere & " FACTURAS_CLI.comercial=''" & comercial & "'' and"					
					strwhere4Venc=strwhereVenc & " (FACTURAS_CLI.comercial=''" & comercial & "'' or  vencimientos_salida.comercial = ''" & comercial & "'') and"
					strwhere4Venc=strwhere4Venc & " nrecibo not in (SELECT nrecibo FROM vencimientos_salida with(nolock) WHERE nfactura like ''"&sesionNCliente&"%'' and nfactura = facturas_cli.nfactura AND comercial <> ''" & comercial & "'') and"
					strwhere2= strWhere2 & " FACTURAS_CLI.comercial=''" & comercial & "'' and"
					strwhere3="facturas_cli.comercial as comercial,"
					strwhere333="facturas_cli.comercial,"
					strwhere34="vencimientos_salida.comercial as comercial,"
					strwhere33="comercial,"
				elseif tabla="albaranes_cli" then
					strwhere= strWhere & " albaranes_cli.comercial=''" & comercial & "'' and"
					strwhere2= strWhere2 & " albaranes_cli.comercial=''" & comercial & "'' and"   
				else
					strwhere6= strWhere6 & " tickets.usuario=''" & comercial & "'' and"
				end if%>
				<font class="cab"><b>
				<%if tabla="facturas_cli" or tabla="vencimientos" or tabla="albaranes_cli" or tabla="efectos_cli" then
					if si_tiene_modulo_comercial<>0 then
						response.write(LitComercialModCom)
					else
						response.write(LitComercial)
					end if
				else
					response.write LitOperador
				end if%>
				:&nbsp;</b></font><font class=cab><%=EncodeForHtml(d_lookup("nombre","personal","dni='" & comercial & "'",session("backendlistados")))%></font><br/>
			<%end if
		else		
''		    'FLM:20090430: no hay comerciales
			if (tabla="facturas_cli" or tabla="vencimientos") then
				strwhere3="facturas_cli.comercial as comercial,"
				strwhere333="facturas_cli.comercial,"
				strwhere34="vencimientos_salida.comercial as comercial,"
				strwhere33="comercial,"
			elseif tabla="albaranes_cli" then
				strwhere3="albaranes_cli.comercial as comercial,"
				strwhere333="albaranes_cli.comercial,"
				strwhere33="comercial,"
''			'FLM:20090430: no hay comerciales
			elseif tabla="efectos_cli" then 
                strwhere3="facturas_cli.comercial as comercial,"
				strwhere333="facturas_cli.comercial,"
				strwhere34="vencimientos_salida.comercial as comercial,"
				strwhere33="comercial,"
			else
				strwhere3="tickets.usuario as comercial,"
				strwhere333="tickets.usuario,"
				strwhere33="comercial,"
			end if
		end if

		if agrupar_poblacion>"" or poblacion>"" then
			strwhere35="upper(domicilios.poblacion) as poblacion,"
			strwhere36="upper(domicilios.poblacion) as poblacion,"
			strwhere444="domicilios.poblacion,"
			strwhere41="poblacion,"
			strfrom="domicilios with(nolock),"
			strwhereEfePob="" 
		end if

		if tabla="albaranes_cli" and imptotalbmay>"" then
			strwhere2=strwhere2 & " albaranes_cli.total_albaran>" & replace(imptotalbmay, ",", ".") & " and"%>
			<font class=cab><b>
			<%if si_tiene_modulo_importaciones<>0 then%>
				<%=LitImpTotalEmbarque%>
			<%else%>
				<%=LitImpTotalAlbaran%>
			<%end if%>
			:&nbsp;</b></font><font class=cab><%=EncodeForHtml(imptotalbmay)%></font><br/>             
		<%end if

		strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
		strwhereVenc=mid(strwhereVenc,1,len(strwhereVenc)-4) 'Quitamos el último AND
		strwhere2=mid(strwhere2,1,len(strwhere2)-4) 'Quitamos el último AND
		if strwhere4Venc="" or strwhere4Venc=" where " then
			strwhere4Venc=strwhereVenc
		else
			strwhere4Venc=mid(strwhere4Venc,1,len(strwhere4Venc)-4) 'Quitamos el último AND
		end if

        'FLM:20090820:Caso para los efectos:Une los albaranes y las facturas que no están en los efectos y los efectos.
        if (tabla="facturas_cli" or tabla="vencimientos") and  nserieEfec&""<>"" then
            'FLM:20090820:Asigno la tabla para los efectos
            tablaSP="EFECTOS_CLI"

''ricardo 24-9-2010 se vuelve a poner el "0," que a fecha 11/2/2010 habia sido quitado por MPC, ya que si incluimos una serie de efectos, da error porque no coincide el numero de columnas en la union
            'seleccion="SELECT " & iif(strwhere3&""="","null,",strwhere3) & iif(strwhere36&""="","null,",strwhere36) & "0,divisas.abreviatura,divisas.ndecimales,facturas_cli.divisa as divisa,facturas_cli.ncliente as ncliente,rsocial, ''F'' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, facturas_cli.deuda AS Deuda"
''ricardo 24-9-2010 linea que existia antes de esta fecha
            seleccion="SELECT " & iif(strwhere3&""="","null,",strwhere3) & iif(strwhere36&""="","null,",strwhere36) & "divisas.abreviatura,divisas.ndecimales,facturas_cli.divisa as divisa,facturas_cli.ncliente as ncliente,rsocial, ''F'' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, facturas_cli.deuda AS Deuda"
''ricardo 24-9-2010 linea que se pone a esta fecha
seleccion="SELECT " & iif(strwhere3&""="","null,",strwhere3) & iif(strwhere36&""="","null,",strwhere36) & "0,divisas.abreviatura,divisas.ndecimales,facturas_cli.divisa as divisa,facturas_cli.ncliente as ncliente,rsocial, ''F'' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, facturas_cli.deuda AS Deuda"

            'la siguiente linea se pone para saber quitando los vencimientos, lo que le queda a la factura por cobrar(deuda-pendiente(recibos))
			seleccion=seleccion & ",facturas_cli.deuda-(select isnull(sum(importe - importecob),0) from vencimientos_salida with(nolock) where nfactura like ''"&sesionNCliente&"%'' and nfactura=facturas_cli.nfactura) as pend_fact,null as nventa"
			seleccion4=" FROM " & strfrom & "divisas with (nolock),FACTURAS_CLI with (nolock) left outer join comerciales comerc with(nolock) on comerc.comercial like ''"&sesionNCliente&"%'' and comerc.comercial=facturas_cli.comercial, clientes " & strwhere & " and cobrada=0 and divisas.codigo like ''"&sesionNCliente&"%'' and facturas_cli.nfactura like ''"&sesionNCliente&"%'' "
			seleccion=seleccion & seleccion4
			''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			seleccion=seleccion & " and facturas_cli.deuda<>0"
			seleccion=seleccion & " UNION "
			seleccion=seleccion & "SELECT " & iif(strwhere34&""="","null,",strwhere34) & iif(strwhere36&""="","null,",strwhere36) & "0,divisas.abreviatura,divisas.ndecimales,facturas_cli.divisa as divisa,facturas_cli.ncliente as ncliente,rsocial, ''V'' AS Tipo, vencimientos_salida.nfactura AS Ndoc, nrecibo AS Nvto,fechav AS Fecha, importe AS Total,importe - importecob AS Deuda"
			'la siguiente linea se pone para saber quitando los vencimientos, lo que le queda a la factura por cobrar(deuda-pendiente(recibos))
			seleccion=seleccion & ",0 as pend_fact,null as nventa"
			seleccion5=" FROM " & strfrom & "divisas with (nolock),FACTURAS_CLI with(nolock) left outer join comerciales comerc with (nolock) on comerc.comercial like ''"&sesionNCliente&"%'' and comerc.comercial=facturas_cli.comercial, clientes with (nolock), vencimientos_salida with(nolock)" & strwhere4Venc & " and facturas_cli.nfactura = vencimientos_salida.nfactura and cobrado=0  and divisas.codigo like ''"&sesionNCliente&"%'' and facturas_cli.nfactura like ''"&sesionNCliente&"%'' and clientes.ncliente like ''"&sesionNCliente&"%'' and vencimientos_salida.nfactura like ''"&sesionNCliente&"%''"
			seleccion=seleccion & seleccion5
			''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			seleccion=seleccion & " and (vencimientos_salida.importe - vencimientos_salida.importecob)<>0"
			seleccion = seleccion & " UNION "
			seleccion = seleccion & " SELECT distinct "& replace(iif(strwhere3&""="","null as a,",strwhere3) & iif(strwhere35&""="","null as b,",strwhere35),"facturas_cli.comercial as","null as") & "0 as c,divisas.abreviatura,divisas.ndecimales,efectos_cli.divisa AS divisa,efectos_cli.ncliente AS ncliente,rsocial, ''E'' AS Tipo,efectos_cli.ndocefecto AS Ndoc, NULL AS nvto,efectos_cli.fecha AS fecha,efectos_cli.importe AS Total, "
			seleccion = seleccion & " efectos_cli.importe AS deuda,0 AS pend_fact,efectos_cli.nefecto as nventa "
            seleccion = seleccion & " FROM " & strfrom & " divisas with(nolock),efectos_cli with(nolock) "
            seleccion = seleccion & " inner join clientes with(nolock) on clientes.ncliente like ''"&sesionNCliente&"%''  and clientes.ncliente=efectos_cli.ncliente"
            seleccion = seleccion & "  "&strwhereEfc&"  divisas.codigo like ''"&sesionNCliente&"%'' and  efectos_cli.pendiente=1 /*and efectos_cli.importe>0 */" 
			''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			seleccion=seleccion & " and (efectos_cli.importe)<>0"
		elseif (tabla="facturas_cli" or tabla="vencimientos") then
		    'FLM:20090820:Asigno la tabla para las facturas
            tablaSP=tabla
            
		    seleccion="SELECT " & iif(strwhere3&""="","null,",strwhere3) & iif(strwhere36&""="","null,",strwhere36) & "0,divisas.abreviatura,divisas.ndecimales,facturas_cli.divisa as divisa,facturas_cli.ncliente as ncliente,rsocial, ''F'' AS Tipo, nfactura AS Ndoc, NULL AS Nvto,fecha AS Fecha, total_factura AS Total, facturas_cli.deuda AS Deuda"
		    'la siguiente linea se pone para saber quitando los vencimientos, lo que le queda a la factura por cobrar(deuda-pendiente(recibos))
			seleccion=seleccion & ",facturas_cli.deuda-(select isnull(sum(importe - importecob),0) from vencimientos_salida with(nolock) where nfactura like ''"&sesionNCliente&"%'' and nfactura=facturas_cli.nfactura) as pend_fact,null as nventa"
			seleccion4=" FROM " & strfrom & "divisas with (nolock),FACTURAS_CLI with (nolock) left outer join comerciales comerc with(nolock) on comerc.comercial like ''"&sesionNCliente&"%'' and comerc.comercial=facturas_cli.comercial, clientes " & strwhere & " and cobrada=0 and divisas.codigo like ''"&sesionNCliente&"%'' and facturas_cli.nfactura like ''"&sesionNCliente&"%'' "
			seleccion=seleccion & seleccion4
			''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			seleccion=seleccion & " and facturas_cli.deuda<>0"
			seleccion=seleccion & " UNION "
			seleccion=seleccion & "SELECT " & iif(strwhere34&""="","null,",strwhere34) & iif(strwhere36&""="","null,",strwhere36) & "0,divisas.abreviatura,divisas.ndecimales,facturas_cli.divisa as divisa,facturas_cli.ncliente as ncliente,rsocial, ''V'' AS Tipo, vencimientos_salida.nfactura AS Ndoc, nrecibo AS Nvto,fechav AS Fecha, importe AS Total,importe - importecob AS Deuda"
			'la siguiente linea se pone para saber quitando los vencimientos, lo que le queda a la factura por cobrar(deuda-pendiente(recibos))
			seleccion=seleccion & ",0 as pend_fact,null as nventa"
			seleccion5=" FROM " & strfrom & "divisas with (nolock),FACTURAS_CLI with(nolock) left outer join comerciales comerc with (nolock) on comerc.comercial like ''"&sesionNCliente&"%'' and comerc.comercial=facturas_cli.comercial, clientes with (nolock), vencimientos_salida with(nolock)" & strwhere4Venc & " and facturas_cli.nfactura = vencimientos_salida.nfactura and cobrado=0  and divisas.codigo like ''"&sesionNCliente&"%'' and facturas_cli.nfactura like ''"&sesionNCliente&"%'' and clientes.ncliente like ''"&sesionNCliente&"%'' and vencimientos_salida.nfactura like ''"&sesionNCliente&"%''"
			seleccion=seleccion & seleccion5
			''ricardo 18-12-2009 solamente saldran las facturas cuya deuda sea distinto de cero
			seleccion=seleccion & " and (vencimientos_salida.importe - vencimientos_salida.importecob)<>0"
		elseif tabla="albaranes_cli" then
		    'FLM:20090820:Asigno la tabla para los albaranes
            tablaSP=tabla
			'Albaranes facturados sin pagos a cuenta
			seleccion=seleccion & "SELECT " & iif(strwhere3&""="","null,",strwhere3) & iif(strwhere35&""="","null,",strwhere35) & "0,divisas.abreviatura, divisas.ndecimales,albaranes_cli.divisa AS divisa,albaranes_cli.ncliente AS ncliente, rsocial, ''A'' AS Tipo,albaranes_cli.nalbaran AS Ndoc, NULL AS nvto,albaranes_cli.fecha AS fecha,TOTAL_ALBARAN AS Total, TOTAL_ALBARAN AS deuda,0 AS pend_fact,null as nventa"
			seleccion=seleccion & " FROM " & strfrom & "divisas with(nolock) , albaranes_cli with(nolock) left outer join comerciales comerc with(nolock) on comerc.comercial like ''"&sesionNCliente&"%'' and comerc.comercial=albaranes_cli.comercial, clientes with(nolock) "
			seleccion=seleccion & strwhere2 & " AND albaranes_cli.nfactura is null AND clientes.ncliente = albaranes_cli.ncliente AND divisas.codigo = albaranes_cli.divisa AND divisas.codigo like ''"&sesionNCliente&"%'' and albaranes_cli.nalbaran like ''"&sesionNCliente&"%'' and clientes.ncliente like ''"&sesionNCliente&"%'' and "
			seleccion=seleccion & " albaranes_cli.nalbaran NOT IN (SELECT nalbaran FROM pagos_alb with(nolock) where nalbaran like ''"&sesionNCliente&"%'')"
			seleccion=seleccion & " GROUP BY " & strwhere333 & strwhere444 & "divisas.abreviatura, divisas.ndecimales,albaranes_cli.divisa, albaranes_cli.nalbaran, albaranes_cli.ncliente, rsocial,albaranes_cli.fecha, TOTAL_ALBARAN"
			'Albaranes facturados con pagos a cuenta cuyo importe es inferior al del albaran
			seleccion=seleccion & " UNION "
			seleccion=seleccion & "SELECT " & iif(strwhere3&""="","null,",strwhere3) & iif(strwhere35&""="","null,",strwhere35) & "0,divisas.abreviatura, divisas.ndecimales,albaranes_cli.divisa AS divisa,albaranes_cli.ncliente AS ncliente, rsocial, ''A'' AS Tipo,albaranes_cli.nalbaran AS Ndoc, NULL AS nvto,albaranes_cli.fecha AS fecha,TOTAL_ALBARAN AS Total, total_albaran - SUM(importe) AS deuda,0 AS pend_fact,null as nventa"
			seleccion=seleccion & " FROM " & strfrom & "divisas with(nolock) , albaranes_cli with(nolock) left outer join comerciales comerc with(nolock) on comerc.comercial like ''"&sesionNCliente&"%'' and comerc.comercial=albaranes_cli.comercial, clientes with(nolock) , pagos_alb with(nolock) "
			seleccion=seleccion & strwhere2 & " and albaranes_cli.nfactura is null AND pagos_alb.nalbaran = albaranes_cli.nalbaran AND clientes.ncliente = albaranes_cli.ncliente AND divisas.codigo = albaranes_cli.divisa and divisas.codigo like ''"&sesionNCliente&"%'' and albaranes_cli.nalbaran like ''"&sesionNCliente&"%'' and clientes.ncliente like ''"&sesionNCliente&"%'' and pagos_alb.nalbaran like ''"&sesionNCliente&"%''"
			seleccion=seleccion & " GROUP BY " & strwhere333 & strwhere444 & "divisas.abreviatura, divisas.ndecimales,albaranes_cli.divisa, albaranes_cli.ncliente, rsocial, albaranes_cli.nalbaran,albaranes_cli.fecha, TOTAL_ALBARAN"
			seleccion=seleccion & " HAVING total_albaran > SUM(importe)"
		else'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'FLM:20090820:Asigno la tabla para los tickets
            tablaSP="TICKETS_CLI"            
            
		    seleccion = seleccion & " select * from ( "
		    seleccion = seleccion & " SELECT "& iif(strwhere3&""="","null as a,",strwhere3) & iif(strwhere35&""="","null as b,",strwhere35) & "0 as c,divisas.abreviatura,divisas.ndecimales,tickets.divisa AS divisa,tickets.ncliente AS ncliente,rsocial, ''T'' AS Tipo,tickets.nticket AS Ndoc, NULL AS nvto,tickets.fecha AS fecha,TOTAL_TICKET AS Total, "
		    ''ricardo 11-1-2010 se cambia la forma de calcular la deuda del ticket, que se estaba calculando mal para los tickets negativos
            ''seleccion = seleccion & " total_ticket - isnull(tmp.timporteEntrada,0) - isnull(tmp2.timporteSalida,0) AS deuda "
            seleccion = seleccion & " total_ticket - isnull(tmp.timporteEntrada,0) + isnull(tmp2.timporteSalida,0) AS deuda "
            seleccion = seleccion & ",0 AS pend_fact,tickets.nventa "
            seleccion = seleccion & " FROM " & strfrom & " divisas with(nolock),tickets with(nolock) "
            seleccion = seleccion & " left outer join (select sum(importe) as timporteEntrada,ndocumento from caja pt with(nolock) where pt.ndocumento like ''"&sesionNCliente&"%'' and tdocumento=''TICKET'' and tanotacion=''entrada'' group by ndocumento) tmp on tmp.ndocumento=tickets.nticket "
            seleccion = seleccion & " left outer join (select sum(importe) as timporteSalida,ndocumento from caja pt with(nolock) where pt.ndocumento like ''"&sesionNCliente&"%'' and tdocumento=''TICKET'' and tanotacion=''salida'' group by ndocumento) tmp2 on tmp2.ndocumento=tickets.nticket "
            seleccion = seleccion & " left outer join personal per with(nolock) on per.dni=tickets.usuario and per.dni like ''"&sesionNCliente&"%'', clientes with(nolock) "
            seleccion = seleccion & strwhere6 & " tickets.nfactura is null and divisas.codigo like ''"&sesionNCliente&"%'' and tickets.nticket like ''"&sesionNCliente&"%'' "
            seleccion = seleccion & " and clientes.ncliente like ''"&sesionNCliente&"%'' "
            seleccion = seleccion & " group by " & strwhere333 & strwhere444 & " usuario,divisas.abreviatura,divisas.ndecimales,tickets.divisa,tickets.ncliente,rsocial,tickets.nticket,tickets.fecha,TOTAL_TICKET "
            ''ricardo 11-1-2010 se cambia la forma de calcular la deuda del ticket, que se estaba calculando mal para los tickets negativos
            ''seleccion = seleccion & ",total_ticket - isnull(tmp.timporteEntrada,0) - isnull(tmp2.timporteSalida,0) "
            seleccion = seleccion & ",total_ticket - isnull(tmp.timporteEntrada,0) + isnull(tmp2.timporteSalida,0) "
            seleccion = seleccion & ",tickets.nventa) tmp1 "
            seleccion = seleccion & " where tmp1.deuda<>0 "
		end if

		seleccion=seleccion & " ORDER BY " & strwhere33 & strwhere41 & " rsocial,Ndoc,Nvto "
		rst.cursorlocation=3
	    'FLM:20090820:cambio el parametro tabla por tablaSP que se asigna en el if anterior.
	    ''response.Write("el seleccion es-" & seleccion & "-<br/>")
	    ''response.end
	    rst.Open "set ansi_warnings off exec ListadoCobrosPendientes '" & seleccion & "', '" & session("usuario")&"', '" & tablaSP & "' set ansi_warnings on",session("backendlistados")
	    if rst.EOF then
	        hayRegistros=0
		    rst.Close
		else
		    hayRegistros=1
	    end if
		if hayRegistros=0 then%>
		   <script language="javascript" type="text/javascript">

		       parent.botones.document.location = "listado_cobros_param_bt.asp?mode=select1";
		       alert("<%=LitMsgDatosNoExiste%>");
			</script>
			<%if viene<>"tienda" then%>
				<script>
					document.location="listado_cobros_param.asp?mode=select1";
				</script>
			<%end if
		else
			NumRegsTotalPendFacQuitar=0 ' lineas que quitamos al final por quitar tambien las lineas de factura por la deuda-(importe-importecob)
			NumRegTotal=rst.recordcount
			NumRegsTotal=NumRegTotal
			rst.movefirst

		   	'Calculos de páginas--------------------------
		  	lote=p_lote
			if lote="" then
				lote=1
			end if
			sentido=p_sentido

			lotes=NumRegTotal/MAXPAGINA
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
			rst.movefirst
			rst.PageSize=MAXPAGINA
			rst.AbsolutePage=lote%>
			<hr/>
			<%NavPaginas lote,lotes,campo,criterio,texto,1%>
			<table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
			<%'Fila de encabezado
			DrawFila color_fondo
			if agrupar_comercial>"" then
				if tabla="facturas_cli" or tabla="vencimientos" or tabla="albaranes_cli" or tabla="efectos_cli" then
					if si_tiene_modulo_comercial<>0 then
						DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitComercialModCom & "</b>"
					else
						DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitComercial & "</b>"
					end if
				else
					DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitOperador & "</b>"
				end if
			end if
			if agrupar_poblacion>"" then
				DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitPoblacion & "</b>"
			end if
			DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitCliente & "</b>"
			if (tabla="facturas_cli" or tabla="vencimientos") then
				DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitFacVen & "</b>"
			elseif tabla="albaranes_cli" then
				if si_tiene_modulo_importaciones<>0 then
					DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitEmbVen & "</b>"
				else
					DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitAlbVen & "</b>"
				end if
			else
				DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitTicVen & "</b>"
			end if
			if tabla="facturas_cli" or tabla="vencimientos" or tabla="albaranes_cli" or tabla="efectos_cli" then
				DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitFechaF & "</b>"
			else
				DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitFechaT & "</b>"
			end if
			DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitFechaR & "</b>"
			DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitValorF & "</b>"
			DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitValorR & "</b>"
			DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitPendienteF & "</b>"
			DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitPendienteR & "</b>"
            DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitCurrency & "</b>"
			CloseFila
			Gtotal_valor = 0
			Gtotal_pendiente = 0
			total_valor = 0
			total_pendiente = 0
			Gtotal_valorR = 0
			Gtotal_pendienteR = 0
			total_valorR = 0
			total_pendienteR = 0
			cli=1

			'numero de registros
			if NumRegs="" then NumRegs=0
			ClienteAnt=""
			ComercialAnt=""
			PoblacionAnt=""
			fila=1

			if viene<>"tienda" then
				VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarAlbaranesCli)=1:VinculosPagina(MostrarFacturasCli)=1
				VinculosPagina(MostrarContratos)=1:VinculosPagina(MostrarEmbarques)=1:VinculosPagina(MostrarEfectosCli)=1
				CargarRestricciones session("usuario"),sesionNCliente,Permisos,Enlaces,VinculosPagina
			end if

			comprobar=0
			no_he_escrito=0
			puesto_drawfila=0
			no_se_ha_puesto_cliente=0
			no_se_ha_puesto_comercial=0
			no_se_ha_puesto_poblacion=0
			seleccion2="select * from ["&session("usuario")&"] "
			ordenar = " ORDER BY " & strwhere33 & strwhere41 & "rsocial,Ndoc,Nvto"
			while not rst.EOF and fila<=MAXPAGINA			
				todavia_no_se_ha_impreso_nada=0
				if rst("tipo")="F" then
					continuar=0
					'COMENTADO POR VGR 16/05/03 PARA QUE SALGAN TAMBIEN LAS DE IMPORTE NEGATIVO
					if rst("tipo")="F" and (rst("pend_fact")=0 and rst("pend_fact") & "">"") then
						continuar=1
					end if
					while continuar=1
						fila=fila+1
						rst.movenext
						todavia_no_se_ha_impreso_nada=1
						if not rst.eof then
							'COMENTADO POR VGR 16/05/03 PARA QUE SALGAN TAMBIEN LAS DE IMPORTE NEGATIVO
							if rst("tipo")="F" and (rst("pend_fact")=0 and rst("pend_fact") & "">"") then
								continuar=1
							else
								continuar=0
							end if
						else
							continuar=0
						end if
					wend
				end if
				if rst.eof then
					rst.moveprevious
				end if
				cabecera=0
					ClienteAct=rst("ncliente")
					comprob_com=0
					if agrupar_comercial>"" then
						if fila<=1 or ComercialAnt<>rst("comercial") or (isnull(ComercialAnt) and rst("comercial") & "">"") or (ComercialAnt & "">"" and isnull(rst("comercial")))then
							comprob_com=1
						else
							comprob_com=0
						end if
					end if
					comprob_cli=0
					if fila<=1 or rst("ncliente")<>ClienteAnt or (rst("ncliente") & ""<>ClienteAnt & "") then
						comprob_cli=1
					else
						comprob_cli=0
					end if
					if comprob_cli=1 or comprob_com=1 then
						if ClienteAnt & ""<>"" then
							total_valor=0
							total_pendiente=0
							Gtotal_valor=0
							Gtotal_pendiente=0
							total_valorR=0
							total_pendienteR=0
							Gtotal_valorR=0
							Gtotal_pendienteR=0
							if agrupar_comercial>"" then
								if isnull(ComercialAnt) then
									where = " comercial is null and ncliente='" & ClienteAnt & "'"
								else
									where = "ncliente='" & ClienteAnt & "' and comercial='"&ComercialAnt&"'"
								end if
							else
								where = " ncliente='" & ClienteAnt & "'"
							end if

							if rstAux.state<>0 then rstAux.close
							rstAux.cursorlocation=3
							if where & "" <> "" then strSelect= seleccion2 & " where " & where
							strSelect=strSelect & ordenar
							rstAux.Open strSelect ,session("backendlistados")

							NumRegs=0
							NumRegsTotalPendFacQuitar=0

							while not rstAux.eof
								if rstAux("ncliente")=ClienteAnt then
									if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										if rstAux("divisa")<>moneda_base then
											divisa_ant=rstAux("divisa")
											divisa_act=rstAux("divisa")
											while divisa_ant=divisa_act and not rstAux.eof
												if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E"  then
													if rstAux("Tipo")="F" then
														if rstAux("pend_fact")<>0 then
															total_valor = total_valor + null_z(rstAux("Total"))
														end if
													end if
													if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E"  then
														total_valor = total_valor + null_z(rstAux("Total"))
													end if
													if rstAux("Tipo")="F" then
														if rstAux("pend_fact")<>0 then
															total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
														end if
													end if
													if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E"  then
														total_pendiente = total_pendiente + null_z(rstAux("deuda"))
													end if
												end if
												if rstAux("Tipo")="V" then
													total_valorR = total_valorR + null_z(rstAux("Total"))
													total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
												end if
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														NumRegs=NumRegs+1
													end if
													documento_contado=rstAux("ndoc")
												else
													if rstAux("Tipo")="V" then
														if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
															NumRegs=NumRegs+1
															documento_contado=rstAux("ndoc")
														end if
													end if
													if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E"  then
														NumRegs=NumRegs+1
													end if
												end if
												rstAux.movenext
												if not rstAux.eof then
													divisa_act=rstAux("divisa")
												end if
											wend
											if divisa_ant<>moneda_base then
												total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
												total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
												total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
												total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
											end if
											pasado=1
								  		else
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													total_valor = total_valor + null_z(rstAux("Total"))
													total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
												end if
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E"  then
												total_valor = total_valor + null_z(rstAux("Total"))
												total_pendiente = total_pendiente + null_z(rstAux("deuda"))
											end if
									  	end if
									elseif rstAux("Tipo")="V" then
										if rstAux("divisa")<>moneda_base then
											divisa_ant=rstAux("divisa")
											divisa_act=rstAux("divisa")
											while divisa_ant=divisa_act and not rstAux.eof
												if rstAux("Tipo")="V" then
													total_valorR = total_valorR + null_z(rstAux("Total"))
													total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
												end if
												rstAux.movenext
												NumRegs=NumRegs+1
												if not rstAux.eof then
													divisa_act=rstAux("divisa")
												end if
											wend
											if divisa_ant<>moneda_base then
												total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
												total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
											end if
											pasado=1
								  		else
											total_valorR = total_valorR + null_z(rstAux("Total"))
											total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
									  	end if
									end if
								end if

								'si hay un registro activo no hacemos el movenext
								if pasado=0 then
									if rstAux("Tipo")="F" then
										if rstAux("pend_fact")<>0 then
											NumRegs=NumRegs+1
											documento_contado=rstAux("ndoc")
										else
											NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
										end if
									else
										if rstAux("Tipo")="V" then
											NumRegs=NumRegs+1
											documento_contado=rstAux("ndoc")
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E"  then
											NumRegs=NumRegs+1
										end if
									end if
									rstAux.movenext
									pasado=0
								else
									pasado=0
								end if
								Gtotal_valor = Gtotal_valor + total_valor
								Gtotal_pendiente = Gtotal_pendiente + total_pendiente
								total_valor =0
								total_pendiente =0
								Gtotal_valorR = Gtotal_valorR + total_valorR
								Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
								total_valorR =0
								total_pendienteR =0
							wend
							rstAux.close
							documento_contado=""
							if NumRegs=0 then NumRegs=1
							DrawFila ""
							if agrupar_comercial>"" then
								DrawCelda "TDBORDECELDA7 bgcolor='" & color_blau & "'","","",0,""
							end if
							if agrupar_poblacion & "">"" then
								DrawCelda "TDBORDECELDA7 bgcolor='" & color_blau & "'","","",0,""
							end if

							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<b>" & LitRegistros & ": " & NumRegs+NumRegsClienteAnadir & "</b>"
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotalCliente & "(" & abreviaturaMB & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & formatnumber(Gtotal_pendiente + Gtotal_pendienteR,n_decimalesMB,-1,0,-1) & "</b></td></tr></table>"
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
							DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1) & "</b>"
							DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1) & "</b>"
							DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1) & "</b>"
							DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1) & "</b>"
                            DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
							CloseFila
							NumRegs=0
							NumRegsClienteAnadir=0
							todavia_no_se_ha_impreso_nada=0
				   		end if
						total_valor=0
						total_pendiente=0
					end if
					if agrupar_comercial>"" then
						ComercialAct=rst("comercial")
						if (isnull(ComercialAnt) and rst("comercial") & "">"") or (isnull(rst("comercial")) and ComercialAnt & "">"") or (rst("comercial") & ""<>ComercialAnt & "") then
							if (ComercialAnt & "">"" and todavia_no_se_ha_impreso_nada=0) or (ComercialAnt & ""="" and fila>1 and todavia_no_se_ha_impreso_nada=0) then
								total_valor=0
								total_pendiente=0
								Gtotal_valor=0
								Gtotal_pendiente=0
								total_valorR=0
								total_pendienteR=0
								Gtotal_valorR=0
								Gtotal_pendienteR=0
								if isnull(ComercialAnt) then
									where  = "comercial is null"
								else
									where = " comercial='" & ComercialAnt & "'"
								end if
								if rstAux.state<>0 then rstAux.close
								if where & "" <> "" then strSelect= seleccion2 & " where " & where & ordenar
								rstAux.cursorlocation=3
								rstAux.Open strSelect,session("backendlistados")

								NumRegs=0
								NumRegsTotalPendFacQuitar=0

								while not rstAux.eof
									if (rstAux("comercial")=ComercialAnt) or (isnull(rstAux("comercial")) and isnull(ComercialAnt)) then
										if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											if rstAux("divisa")<>moneda_base then
												divisa_ant=rstAux("divisa")
												divisa_act=rstAux("divisa")
												while divisa_ant=divisa_act and not rstAux.eof
													if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
														if rstAux("Tipo")="F" then
															if rstAux("pend_fact")<>0 then
																total_valor = total_valor + null_z(rstAux("Total"))
															end if
														end if
														if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E"  then
															total_valor = total_valor + null_z(rstAux("Total"))
														end if
														if rstAux("Tipo")="F" then
															if rstAux("pend_fact")<>0 then
																total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
															end if
														end if
														if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
															total_pendiente = total_pendiente + null_z(rstAux("deuda"))
														end if
													end if
													if rstAux("Tipo")="F" then
														if rstAux("pend_fact")<>0 then
															NumRegs=NumRegs+1
														end if
														documento_contado=rstAux("ndoc")
													else
														if rstAux("Tipo")="V" then
															if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
																NumRegs=NumRegs+1
																documento_contado=rstAux("ndoc")
															end if
														end if
														if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
															NumRegs=NumRegs+1
														end if
													end if
													rstAux.movenext
													if not rstAux.eof then
														divisa_act=rstAux("divisa")
													end if
												wend
												if divisa_ant<>moneda_base then
													total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
													total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
													'le decimos que hay un registro activo
												end if
												pasado=1
											else
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_valor = total_valor + null_z(rstAux("Total"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_valor = total_valor + null_z(rstAux("Total"))
												end if
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_pendiente = total_pendiente + null_z(rstAux("deuda"))
												end if
											end if
										elseif rstAux("Tipo")="V" then
											if rstAux("divisa")<>moneda_base then
												divisa_ant=rstAux("divisa")
												divisa_act=rstAux("divisa")
												while divisa_ant=divisa_act and not rstAux.eof
													if rstAux("Tipo")="V" then
														total_valorR = total_valorR + null_z(rstAux("Total"))
														total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
													end if
													rstAux.movenext
													NumRegs=NumRegs+1
													if not rstAux.eof then
														divisa_act=rstAux("divisa")
													end if
												wend
												if divisa_ant<>moneda_base then
													total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
													total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
													'le decimos que hay un registro activo
												end if
												pasado=1
											else
												total_valorR = total_valorR + null_z(rstAux("Total"))
												total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
											end if
										end if
									end if
									'si hay un registro activo no hacemos el movenext
									if pasado=0 then
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												NumRegs=NumRegs+1
												documento_contado=rstAux("ndoc")
											else
												NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
											end if
										else
											if rstAux("Tipo")="V" then
												NumRegs=NumRegs+1
												documento_contado=rstAux("ndoc")
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												NumRegs=NumRegs+1
											end if
										end if
										rstAux.movenext

										pasado=0
									else
										pasado=0
									end if
									Gtotal_valor = Gtotal_valor + total_valor
									Gtotal_pendiente = Gtotal_pendiente + total_pendiente
									total_valor =0
									total_pendiente =0
									Gtotal_valorR = Gtotal_valorR + total_valorR
									Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
									total_valorR =0
									total_pendienteR =0
								wend

								documento_contado=""
								DrawFila ""

								if NumRegs=0 then NumRegs=1
								DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "' colspan=2","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsComercialAnadir) & "</b>"
								DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<b>" + LitTotalComercial & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
								if agrupar_poblacion>"" then
									DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
								end if
								DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
								DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
								DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
								DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
                                DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
								CloseFila
								NumRegs=0
								PoblacionAnt=""
								NumRegsComercialAnadir=0
							end if
							todavia_no_se_ha_impreso_nada=0
							cabecera=1 ' con esto decimos que acabamos de escribir el comercial y el cliente y no debemos hacer un salto de columna
						end if
					end if
					if fila<=1 or (rst("ncliente") & ""<>ClienteAnt & "") or agrupar_comercial>"" or agrupar_poblacion>"" then
						if agrupar_comercial>"" then
							if fila<=1 or ComercialAnt & ""<>(rst("comercial") & "") or (isnull(ComercialAnt) and rst("comercial") & "">"") or (ComercialAnt & "">"" and rst("comercial") & ""="")then
								comprob_com=1
							else
								comprob_com=0
							end if
							comprob_cli=0
							if fila<=1 or rst("ncliente")<>ClienteAnt or (rst("ncliente") & ""<>ClienteAnt & "") then
								comprob_cli=1
							else
								comprob_cli=0
							end if
							if comprob_cli=1 or comprob_com=1 then
								if puesto_drawfila=0 then
									DrawFila ""
									puesto_drawfila=1
								end if
								total_valor=0
								total_pendiente=0
								if agrupar_comercial>"" then
									if rst("comercial") & ""<>(ComercialAnt & "") or (isnull(ComercialAnt) and rst("comercial")>"") or (ComercialAnt>"" and isnull(rst("comercial")))then
										DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(Nulear(d_lookup("nombre","personal","dni='" & rst("comercial")& "'",session("backendlistados"))))
									else
										DrawCelda "TDBORDECELDA7","","",0,""
									end if
									no_se_ha_puesto_comercial=0
								end if
								if agrupar_poblacion>"" then
									if UCase(rst("poblacion"))<>PoblacionAnt or (isnull(PoblacionAnt) and UCase(rst("poblacion"))>"") or (PoblacionAnt>"" and isnull(UCase(rst("poblacion"))))then
										DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(Nulear(UCase(rst("poblacion"))))
									else
										DrawCelda "TDBORDECELDA7","","",0,""
									end if
								end if%>
								<td class=TDBORDECELDA7>
									<%if viene<>"tienda" then%>
										<%=Hiperv(OBJClientes,rst("ncliente"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(d_lookup("rsocial","clientes","ncliente='" & rst("ncliente") & "'",session("backendlistados"))),LitVerCliente)%>
									<%else%>
										<%=EncodeForHtml(d_lookup("rsocial","clientes","ncliente='" & rst("ncliente") & "'",session("backendlistados")))%>
									<%end if%>
								</td>
								<%no_se_ha_puesto_cliente=0
								cabecera=1
							else
								if fila>1 and comprob_com=0 then
									no_se_ha_puesto_comercial=1
								end if
								if fila>1 and comprob_cli=0 then
									no_se_ha_puesto_cliente=1
								end if
								if fila>1 and agrupar_poblacion>"" and (comprob_cli=0 or comprob_com=0) then
									no_se_ha_puesto_poblacion=1
								end if
							end if
						else
							if agrupar_poblacion>"" then
								if puesto_drawfila=0 then
									DrawFila ""
									puesto_drawfila=1
								end if
								if UCase(rst("poblacion"))<>PoblacionAnt or (isnull(PoblacionAnt) and UCase(rst("poblacion"))>"") or (PoblacionAnt>"" and isnull(UCase(rst("poblacion"))))then
									DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(Nulear(UCase(rst("poblacion"))))
								else
									DrawCelda "TDBORDECELDA7","","",0,""
								end if
								cabecera=1
							end if
							if (fila<=1 or rst("ncliente") & ""<>ClienteAnt & "") then
								if puesto_drawfila=0 then
									DrawFila ""
									puesto_drawfila=1
								end if
								total_valor=0
								total_pendiente=0%>
								<td class=TDBORDECELDA7>
									<%if viene<>"tienda" then%>
										<%=Hiperv(OBJClientes,rst("ncliente"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(d_lookup("rsocial","clientes","ncliente='" & rst("ncliente") & "'",session("backendlistados"))),LitVerCliente)%>
									<%else%>
										<%=EncodeForHtml(d_lookup("rsocial","clientes","ncliente='" & rst("ncliente") & "'",session("backendlistados")))%>
									<%end if%>
								</td>
								<%cabecera=1
								no_se_ha_puesto_cliente=0
							else
								no_se_ha_puesto_cliente=1
							end if
						end if
					else
						no_se_ha_puesto_cliente=1
					end if
					if rst("Tipo")="F" then
						'COMENTADO POR VGR : 15/05/03 PARA QUE SALGAN FAC.IMPORTES NEGATIVOS.
						'if (rst("pend_fact")>0 or rst("pend_fact") & ""="") then
						if (rst("pend_fact")<>0 or rst("pend_fact") & ""="") then
							if puesto_drawfila=0 then
								DrawFila ""
								puesto_drawfila=1
							end if
							documento_impreso=""
							if rst("ndoc") & "">"" then
								documento_impreso=rst("ndoc")
							end if
							if no_se_ha_puesto_cliente=1 then
								DrawCelda "TDBORDECELDA7","","",0,""
							end if
							no_se_ha_puesto_cliente=0
							if no_se_ha_puesto_comercial=1 then
								DrawCelda "TDBORDECELDA7","","",0,""
							end if
							no_se_ha_puesto_comercial=0
							if no_se_ha_puesto_poblacion=1 then
								DrawCelda "TDBORDECELDA7","","",0,""                                                                         
							end if
							no_se_ha_puesto_poblacion=0
							comprobar=0%>
							<td class=TDBORDECELDA7>
								<%if viene<>"tienda" then%>
									<%=Hiperv(OBJFacturasCli,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(trimCodEmpresa(rst("ndoc"))),LitVerFactura)%>
								<%else%>
									<a class='CELDAREFB' href=javascript:ver_documento(<%=enc.EncodeForJavascript(rst("ndoc"))%>','facturas','<%=enc.EncodeForJavascript(ncliente)%>') alt='<%=LitVerFacturas%>'><%=EncodeForHtml(trimCodEmpresa(rst("ndoc")))%></a>
								<%end if%>
							</td>
							<%DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("Fecha"))
							DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
							DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1))'' & " " & rst("abreviatura")
							DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
							if (rst("pend_fact")<>0 or rst("pend_fact") & ""="") then
								'DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1) & " "  & rst("abreviatura")
								DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("pend_fact")),rst("ndecimales"),-1,0,-1))'' & " "  & rst("abreviatura")
							else
								DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
							end if
							DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
                            DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("abreviatura"))
							fila=fila+1
							if puesto_drawfila=1 then
								CloseFila
								puesto_drawfila=0
							end if
						else
							documento_impreso=""
							if rst("ndoc") & "">"" then
								documento_impreso=rst("ndoc")
							end if
							no_he_escrito=1
							fila=fila+1
						end if
					end if
					if rst("Tipo")="V" then
						he_entrado=0
						if puesto_drawfila=0 then
							DrawFila ""
							puesto_drawfila=1
						end if
						if no_se_ha_puesto_cliente=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_cliente=0
						if no_se_ha_puesto_comercial=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_comercial=0
						if no_se_ha_puesto_poblacion=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_poblacion=0
						comprobar=0
						if (documento_impreso & "">"" and instr(1,rst("Nvto"),documento_impreso,1)>0) then
							DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(trimCodEmpresa(rst("Nvto")))
							documento_impreso=rst("ndoc")
						else%>
							<td class=TDBORDECELDA7>
							<%if viene<>"tienda" then%>
								<%=Hiperv(OBJFacturasCli,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(trimCodEmpresa(rst("Nvto"))),LitVerFactura)%>
							<%else%>
								<a class='CELDAREFB' href=javascript:ver_documento('<%=enc.EncodeForJavascript(rst("ndoc"))%>','facturas','<%=enc.EncodeForJavascript(ncliente)%>') alt='<%=LitVerFacturas%>'><%=EncodeForHtml(trimCodEmpresa(rst("Nvto")))%></a>
							<%end if%>
							</td>
							<%documento_impreso=""
						end if
						'DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(d_lookup("fecha","facturas_cli","nfactura='" & rst("ndoc") & "'",session("backendlistados")))
						DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
                        DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("abreviatura"))
						fila=fila+1
						cabecera=0
						if puesto_drawfila=1 then
							CloseFila
							puesto_drawfila=0
						end if
					end if
					if rst("Tipo")="A" then
						if puesto_drawfila=0 then
							DrawFila ""
							puesto_drawfila=1
						end if
						if no_se_ha_puesto_cliente=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_cliente=0
						if no_se_ha_puesto_comercial=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_comercial=0
						if no_se_ha_puesto_poblacion=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_poblacion=0
						comprobar=0%>
						<td class=TDBORDECELDA7>
							<%if viene<>"tienda" then
								if si_tiene_modulo_importaciones<>0 then%>
									<%=Hiperv(OBJEmbarques,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(trimCodEmpresa(rst("ndoc"))),LitVerEmbarque)%>
								<%else%>
									<%=Hiperv(OBJAlbaranesCli,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(trimCodEmpresa(rst("ndoc"))),LitVerAlbaran)%>
								<%end if
							else%>
								<a class='CELDAREFB' href=javascript:ver_documento('<%=enc.EncodeForJavascript(rst("ndoc"))%>','albaranes','<%=enc.EncodeForJavascript(ncliente)%>') alt='<%=LitVerAlbaran%>'><%=EncodeForHtml(trimCodEmpresa(rst("ndoc")))%></a>
							<%end if%>
						</td>
						<%DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
                        DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("abreviatura"))
						fila=fila+1
						cabecera=0
						if puesto_drawfila=1 then
							CloseFila
							puesto_drawfila=0
						end if
					end if
					if rst("Tipo")="T" then
						if puesto_drawfila=0 then
							DrawFila ""
							puesto_drawfila=1
						end if
						if no_se_ha_puesto_cliente=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_cliente=0
						if no_se_ha_puesto_comercial=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_comercial=0
						if no_se_ha_puesto_poblacion=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_poblacion=0
						comprobar=0%>
						<td class=TDBORDECELDA7>
							<%if viene<>"tienda" then%>
							    <%=Hiperv(OBJTickets,rst("ndoc"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,EncodeForHtml(trimCodEmpresa(rst("nventa"))),LitVerTicket)%>
							<%else%>
								<a class='CELDAREFB' href=javascript:ver_documento('<%=enc.EncodeForJavascript(rst("ndoc"))%>','tickets','<%=enc.EncodeForJavascript(ncliente)%>') alt='<%=LitVerTicket%>'><%=EncodeForHtml(trimCodEmpresa(rst("nventa")))%></a>
							<%end if%>
						</td>
						<%DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
                        DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("abreviatura"))
						fila=fila+1
						cabecera=0
						if puesto_drawfila=1 then
							CloseFila
							puesto_drawfila=0
						end if
					end if
					'FLM:20090430:añado el caso para los efectos.
					if rst("Tipo")="E" then
						if puesto_drawfila=0 then
							DrawFila ""
							puesto_drawfila=1
						end if
						if no_se_ha_puesto_cliente=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_cliente=0
						if no_se_ha_puesto_comercial=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_comercial=0
						if no_se_ha_puesto_poblacion=1 then
							DrawCelda "TDBORDECELDA7","","",0,""
						end if
						no_se_ha_puesto_poblacion=0
						comprobar=0%>
						<td class=TDBORDECELDA7>
							<%if viene<>"tienda" then%>
									<%=Hiperv(OBJEfectosCli,rst("nventa"),"","",Permisos,Enlaces,session("usuario"),sesionNCliente,(EncodeForHtml(rst("ndoc")) & " (" &LitEfecto&")"),LitVerEfectoCli)%>
							<%else%>
								<a class='CELDAREFB' href=javascript:ver_documento('<%=enc.EncodeForJavascript(rst("nventa"))%>','efectos','<%=enc.EncodeForJavascript(ncliente)%>') alt='<%=LitVerEfectoCli%>'><%=EncodeForHtml(rst("ndoc"))%></a>
							<%end if%>
						</td>
						<%DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("Fecha"))
						DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(d_lookup("fechavto","efectos_cli","nefecto='" & rst("nventa") & "'",session("backendlistados")))
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("Total")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))' & " " & rst("abreviatura")
						DrawCelda "TDBORDECELDA7","","",0,"&nbsp;"
                        DrawCelda "TDBORDECELDA7","","",0,EncodeForHtml(rst("abreviatura"))
						fila=fila+1
						cabecera=0
						if puesto_drawfila=1 then
							CloseFila
							puesto_drawfila=0
						end if
					end if
					ClienteAnt=rst("ncliente")
					if agrupar_comercial>"" then
						ComercialAnt=rst("comercial")
					end if
					if agrupar_poblacion>"" then
						PoblacionAnt=UCase(rst("poblacion"))
					end if
					rst.MoveNext
			wend
			if fila>MAXPAGINA and (ClienteAct & ""<>ClienteAnt & "") then
				if NumRegs=0 then NumRegs=1
				if agrupar_comercial>"" then
					DrawCelda "TDBORDECELDA7 bgcolor='" & color_blau & "'","","",0,""
				end if
				if agrupar_poblacion>"" then
					DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
				end if
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsClienteAnadir) & "</b>"
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotalCliente & "(" & EncodeForHtml(abreviaturaMB) & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(total_pendiente + total_pendienteR,n_decimalesMB,-1,0,-1)) & "</b></td></tr></table>"
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_valor,n_decimalesMB,-1,0,-1)) & "</b>"
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_pendiente,n_decimalesMB,-1,0,-1)) & "</b>"
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_pendienteR,n_decimalesMB,-1,0,-1)) & "</b>"
                DrawCelda "TDBORDECELDA7","","",0,""
				CloseFila
				NumRegs=0
				NumRegsClienteAnadir=0
			end if
			if agrupar_comercial>"" and fila>MAXPAGINA and (ComercialAct & ""<>ComercialAnt & "") then
				if NumRegs=0 then NumRegs=1
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "' colspan=2","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsComercialAnadir) & "</b>"
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<b>" + LitTotalComercial & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
				if agrupar_poblacion>"" then
					DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
				end if
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
				DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_valor,n_decimalesMB,-1,0,-1)) & "</b>"
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_pendiente,n_decimalesMB,-1,0,-1)) & "</b>"
				DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(total_pendienteR,n_decimalesMB,-1,0,-1)) & "</b>"
                DrawCelda "TDBORDECELDA7","","",0,""
				CloseFila
				NumRegs=0
				PoblacionAnt=""
				NumRegsComercialAnadir=0
			end if
			if lote=lotes then
				if agrupar_comercial>"" then
					'imprimimos el total del actual comercial y actual cliente
					total_valor=0
					total_pendiente=0
					Gtotal_valor=0
					Gtotal_pendiente=0
					total_valorR=0
					total_pendienteR=0
					Gtotal_valorR=0
					Gtotal_pendienteR=0

					if ComercialAct & "">"" then
						where=" comercial='" & ComercialAct & "'"
					end if
					if ClienteAct & "">"" then
						if where&"">"" then where=where& " and "
						where=where & " ncliente='" & ClienteAct & "'"
					end if
					if rstAux.state<>0 then rstAux.close
					if where & "">"" then
						strSelect = seleccion2 & " where " & where & ordenar
					else
						strSelect = seleccion2 & ordenar
					end if
					rstAux.cursorlocation=3
					rstAux.Open strSelect,session("backendlistados")

					NumRegsTotalPendFacQuitar=0
					while not rstAux.eof
						if rstAux("comercial")=ComercialAct or (rstAux("comercial") & ""=ComercialAct & "") then
							if rstAux("ncliente")=ClienteAct or (rstAux("ncliente") & ""=ClienteAct & "") then
								if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									if rstAux("divisa")<>moneda_base then
										divisa_ant=rstAux("divisa")
										divisa_act=rstAux("divisa")
										while divisa_ant=divisa_act and not rstAux.eof
											if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_valor = total_valor + null_z(rstAux("Total"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_valor = total_valor + null_z(rstAux("Total"))
												end if
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_pendiente = total_pendiente + null_z(rstAux("deuda"))
												end if
											end if
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													NumRegs=NumRegs+1
												end if
												documento_contado=rstAux("ndoc")
											else
												if rstAux("Tipo")="V" then
													if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
														NumRegs=NumRegs+1
														documento_contado=rstAux("ndoc")
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													NumRegs=NumRegs+1
												end if
											end if
											rstAux.movenext
											if not rstAux.eof then
												divisa_act=rstAux("divisa")
											end if
										wend
										if divisa_ant<>moneda_base then
											total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
											total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
										end if
										'le decimos que hay un registro activo
										pasado=1
									else
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_valor = total_valor + null_z(rstAux("Total"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_valor = total_valor + null_z(rstAux("Total"))
										end if
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_pendiente = total_pendiente + null_z(rstAux("deuda"))
										end if
									end if
								elseif rstAux("Tipo")="V" then
									if rstAux("divisa")<>moneda_base then
										divisa_ant=rstAux("divisa")
										divisa_act=rstAux("divisa")
										while divisa_ant=divisa_act and not rstAux.eof
											if rstAux("Tipo")="V" then
												total_valorR = total_valorR + null_z(rstAux("Total"))
												total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
											end if
											rstAux.movenext
											NumRegs=NumRegs+1
											if not rstAux.eof then
												divisa_act=rstAux("divisa")
											end if
										wend
										if divisa_ant<>moneda_base then
											total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
											total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
										end if
										'le decimos que hay un registro activo
										pasado=1
									else
										total_valorR = total_valorR + null_z(rstAux("Total"))
										total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
									end if
								end if
							end if
						end if
						'si hay un registro activo no hacemos el movenext
						if pasado=0 then
							if rstAux("Tipo")="F" then
								if rstAux("pend_fact")<>0 then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								else
									NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
								end if
							else
								if rstAux("Tipo")="V" then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									NumRegs=NumRegs+1
								end if
							end if
							rstAux.movenext
							pasado=0
						else
							pasado=0
						end if
						Gtotal_valor = Gtotal_valor + total_valor
						Gtotal_pendiente = Gtotal_pendiente + total_pendiente
						total_valor=0
						total_pendiente=0
						Gtotal_valorR = Gtotal_valorR + total_valorR
						Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
						total_valorR=0
						total_pendienteR=0
					wend
					rstAux.close
					documento_contado=""
				DrawFila color_fondo
					if NumRegs=0 then NumRegs=1
					if agrupar_comercial>"" then
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_blau & "'","","",0,""
					end if
					if agrupar_poblacion>"" then
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
					end if
					DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsClienteAnadir) & "</b>"
					DrawCelda "TDBORDECELDA7","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotalCliente & "(" & EncodeForHtml(abreviaturaMB) & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(Gtotal_pendiente + Gtotal_pendienteR,n_decimalesMB,-1,0,-1)) & "</b></td></tr></table>"
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
                    DrawCelda "TDBORDECELDA7","","",0,""
					CloseFila
					'imprimimos el total del actual comercial
					NumRegs=0
					NumRegsClienteAnadir=0
					total_valor=0
					total_pendiente=0
					Gtotal_valor=0
					Gtotal_pendiente=0
					total_valorR=0
					total_pendienteR=0
					Gtotal_valorR=0
					Gtotal_pendienteR=0

					if ComercialAct & "">"" then
						where = " comercial='" & ComercialAct & "'"
					end if

					if rstAux.state<>0 then rstAux.close
					if where&"">"" then strSelect=seleccion2& " where " & where & ordenar
					rstAux.cursorlocation=3
					rstAux.Open strSelect,session("backendlistados")

					NumRegsTotalPendFacQuitar=0
					while not rstAux.eof
						if rstAux("comercial")=ComercialAct or (rstAux("comercial") & ""=ComercialAct & "") then
							if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
								if rstAux("divisa")<>moneda_base then
									divisa_ant=rstAux("divisa")
									divisa_act=rstAux("divisa")
									while divisa_ant=divisa_act and not rstAux.eof
										if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													total_valor = total_valor + null_z(rstAux("Total"))
												end if
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												total_valor = total_valor + null_z(rstAux("Total"))
											end if
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
												end if
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												total_pendiente = total_pendiente + null_z(rstAux("deuda"))
											end if
										end if
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												NumRegs=NumRegs+1
											end if
											documento_contado=rstAux("ndoc")
										else
											if rstAux("Tipo")="V" then
												if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
													NumRegs=NumRegs+1
													documento_contado=rstAux("ndoc")
												end if
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												NumRegs=NumRegs+1
											end if
										end if
										rstAux.movenext
										if not rstAux.eof then
											divisa_act=rstAux("divisa")
										end if
									wend
									if divisa_ant<>moneda_base then
										total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
										total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
									end if
										'le decimos que hay un registro activo
										pasado=1
								else
									if rstAux("Tipo")="F" then
										if rstAux("pend_fact")<>0 then
											total_valor = total_valor + null_z(rstAux("Total"))
										end if
									end if
									if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										total_valor = total_valor + null_z(rstAux("Total"))
									end if
									if rstAux("Tipo")="F" then
										if rstAux("pend_fact")<>0 then
											total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
										end if
									end if
									if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										total_pendiente = total_pendiente + null_z(rstAux("deuda"))
									end if
								end if
							elseif rstAux("Tipo")="V" then
								if rstAux("divisa")<>moneda_base then
									divisa_ant=rstAux("divisa")
									divisa_act=rstAux("divisa")
									while divisa_ant=divisa_act and not rstAux.eof
										if rstAux("Tipo")="V" then
											total_valorR = total_valorR + null_z(rstAux("Total"))
											total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
										end if
										rstAux.movenext
										NumRegs=NumRegs+1
										if not rstAux.eof then
											divisa_act=rstAux("divisa")
										end if
									wend
									if divisa_ant<>moneda_base then
										total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
										total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
									end if
									'le decimos que hay un registro activo
									pasado=1
								else
									total_valorR = total_valorR + null_z(rstAux("Total"))
									total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
								end if
							end if
						end if
						'si hay un registro activo no hacemos el movenext
						if pasado=0 then
							if rstAux("Tipo")="F" then
								if rstAux("pend_fact")<>0 then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								else
									NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
								end if
							else
								if rstAux("Tipo")="V" then
							        NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									NumRegs=NumRegs+1
								end if
							end if
							rstAux.movenext
							pasado=0
						else
							pasado=0
						end if
						Gtotal_valor = Gtotal_valor + total_valor
						Gtotal_pendiente = Gtotal_pendiente + total_pendiente
						total_valor=0
						total_pendiente=0
						Gtotal_valorR = Gtotal_valorR + total_valorR
						Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
						total_valorR=0
						total_pendienteR=0
					wend
					rstAux.close
					DrawFila color_fondo
					documento_contado=""
					if NumRegs=0 then NumRegs=1
					DrawCelda "TDBORDECELDA7 colspan=2","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsComercialAnadir) & "</b>"
					DrawCelda "TDBORDECELDA7","","",0,"<b>" + LitTotalComercial & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
					if agrupar_poblacion>"" then
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
					end if
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
                    DrawCelda "TDBORDECELDA7","","",0,""
					CloseFila
					'imprimos el total de todo
					NumRegs=0
					PoblacionAnt=""
					NumRegsComercialAnadir=0
					if instr(1,seleccion,"ORDER BY comercial," & strwhere41 & "rsocial,divisa,Ndoc,Nvto",1)>0 then
						seleccion=mid(seleccion,1,len(seleccion)-len("ORDER BY comercial," & strwhere41 & "rsocial,divisa,Ndoc,Nvto"))
					else
						seleccion=mid(seleccion,1,len(seleccion)-len("ORDER BY comercial," & strwhere41 & "rsocial,Ndoc,Nvto"))
					end if
					seleccion=seleccion & "ORDER BY comercial," & strwhere41 & "rsocial,divisa,Ndoc,Nvto"
					total_total_valor=0
					total_total_pendiente=0
					total_valor=0
					total_pendiente=0
					Gtotal_valor=0
					Gtotal_pendiente=0
					total_total_valorR=0
					total_total_pendienteR=0
					total_valorR=0
					total_pendienteR=0
					Gtotal_valorR=0
					Gtotal_pendienteR=0

					if rstAux.state<>0 then rstAux.close
					rstAux.cursorlocation=3
					strSelect = seleccion2 & ordenar
					rstAux.Open strSelect,session("backendlistados")
					NumRegsTotalPendFacQuitar=0

					while not rstAux.eof
						if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
							if rstAux("divisa")<>moneda_base then
								divisa_ant=rstAux("divisa")
								divisa_act=rstAux("divisa")
								while divisa_ant=divisa_act and not rstAux.eof
									if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_valor = total_valor + null_z(rstAux("Total"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_valor = total_valor + null_z(rstAux("Total"))
										end if
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_pendiente = total_pendiente + null_z(rstAux("deuda"))
										end if
									end if
									if rstAux("Tipo")="F" then
										if rstAux("pend_fact")<>0 then
											NumRegs=NumRegs+1
										end if
										documento_contado=rstAux("ndoc")
									else
										if rstAux("Tipo")="V" then
											if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
												NumRegs=NumRegs+1
												documento_contado=rstAux("ndoc")
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											NumRegs=NumRegs+1
										end if
									end if
									rstAux.movenext
									if not rstAux.eof then
										divisa_act=rstAux("divisa")
									end if
								wend
								if divisa_ant<>moneda_base then
									total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
									total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
								end if
								'le decimos que hay un registro activo
								pasado=1
						   	else
								if rstAux("Tipo")="F" then
									if rstAux("pend_fact")<>0 then
										total_valor = total_valor + null_z(rstAux("Total"))
									end if
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									total_valor = total_valor + null_z(rstAux("Total"))
								end if
								if rstAux("Tipo")="F" then
									if rstAux("pend_fact")<>0 then
										total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
									end if
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									total_pendiente = total_pendiente + null_z(rstAux("deuda"))
								end if
							end if
						elseif rstAux("Tipo")="V" then
							if rstAux("divisa")<>moneda_base then
								divisa_ant=rstAux("divisa")
								divisa_act=rstAux("divisa")
								while divisa_ant=divisa_act and not rstAux.eof
									if rstAux("Tipo")="V" then
										total_valorR = total_valorR + null_z(rstAux("Total"))
										total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
									end if
									rstAux.movenext
									NumRegs=NumRegs+1
									if not rstAux.eof then
										divisa_act=rstAux("divisa")
									end if
								wend
								if divisa_ant<>moneda_base then
									total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
									total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
								end if
								'le decimos que hay un registro activo
								pasado=1
						   	else
								total_valorR = total_valorR + null_z(rstAux("Total"))
								total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
							end if
						end if
						'si hay un registro activo no hacemos el movenext
						if pasado=0 and (not rstAux.eof) then
							if rstAux("Tipo")="F" then
								if rstAux("pend_fact")<>0 then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								else
									NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
								end if
							else
								if rstAux("Tipo")="V" then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									NumRegs=NumRegs+1
								end if
							end if
							rstAux.movenext
							pasado=0
						else
							pasado=0
						end if
						Gtotal_valor = Gtotal_valor + total_valor
						Gtotal_pendiente = Gtotal_pendiente + total_pendiente
						total_valor=0
						total_pendiente=0
						Gtotal_valorR = Gtotal_valorR + total_valorR
						Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
						total_valorR=0
						total_pendienteR=0
					wend
					rstAux.close
					documento_contado=""
					DrawFila color_fondo
					DrawCelda "TDBORDECELDA7 colspan=2","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegTotal-NumRegsTotalPendFacQuitar+cuantos_tenemos_que_sumar) & "</b>"
					DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotales & "(" & EncodeForHtml(abreviaturaMB) & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(Gtotal_pendiente + Gtotal_pendienteR,n_decimalesMB,-1,0,-1)) & "</b></td></tr></table>"
					if agrupar_poblacion>"" then
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
					end if
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
                    DrawCelda "TDBORDECELDA7","","",0,""
					CloseFila
					if d_lookup("imp_equiv","configuracion","nempresa='" & sesionNCliente & "'",session("backendlistados")) then
						DrawFila color_fondo
						DrawCelda "TDBORDECELDA7 colspan=2","","",0,""
						Gtotal_pendientePTS  = cdbl(CambioDivisa(null_z(Gtotal_pendiente),moneda_base,sesionNCliente & "01"))
						Gtotal_pendienteRPTS = cdbl(CambioDivisa(null_z(Gtotal_pendienteR),moneda_base,sesionNCliente & "01"))
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotales & "(" & EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados"))) & ")</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(Gtotal_pendientePTS + Gtotal_pendienteRPTS,d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b></td></tr></table>"
						if agrupar_poblacion>"" then
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						end if
						DrawCelda "TDBORDECELDA7","","",0,""
						DrawCelda "TDBORDECELDA7","","",0,""
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(Gtotal_valor,moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(Gtotal_valorR,moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(Gtotal_pendiente),moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(Gtotal_pendienteR),moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7","","",0,""
                        CloseFila
					end if
					NumRegs=0
				else
					total_valor=0
					total_pendiente=0
					Gtotal_valor=0
					Gtotal_pendiente=0
					total_valorR=0
					total_pendienteR=0
					Gtotal_valorR=0
					Gtotal_pendienteR=0
					where = " ncliente='" & ClienteAnt & "'"
					strSelect = seleccion2 & " where " & where & ordenar
					if rstAux.state<>0 then rstAux.close
					rstAux.cursorlocation=3
					rstAux.Open strSelect,session("backendlistados")
					NumRegsTotalPendFacQuitar=0

					while not rstAux.eof
						if rstAux("ncliente") & ""=ClienteAct & "" then
							if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
								if rstAux("divisa")<>moneda_base then
									divisa_ant=rstAux("divisa")
									divisa_act=rstAux("divisa")
									while divisa_ant=divisa_act and not rstAux.eof
										if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													total_valor = total_valor + null_z(rstAux("Total"))
												end if
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												total_valor = total_valor + null_z(rstAux("Total"))
											end if
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
												end if
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												total_pendiente = total_pendiente + null_z(rstAux("deuda"))
											end if
										end if
                                        ''esto lo hemos añadido , para cuando hay varias divisas
										if rstAux("Tipo")="V" then
											total_valorR = total_valorR + null_z(rstAux("Total"))
											total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
										end if

										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												NumRegs=NumRegs+1
											end if
											documento_contado=rstAux("ndoc")
										else
											if rstAux("Tipo")="V" then
												if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
													NumRegs=NumRegs+1
													documento_contado=rstAux("ndoc")
												end if
											end if
											if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												NumRegs=NumRegs+1
											end if
										end if
										rstAux.movenext

										if not rstAux.eof then
											divisa_act=rstAux("divisa")
										end if
									wend
									if divisa_ant<>moneda_base then
										total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
										total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
										total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
										total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
									end if
										'le decimos que hay un registro activo
										pasado=1
								else
									if rstAux("Tipo")="F" then
										if rstAux("pend_fact")<>0 then
											total_valor = total_valor + null_z(rstAux("Total"))
										end if
									end if
									if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										total_valor = total_valor + null_z(rstAux("Total"))
									end if
									if rstAux("Tipo")="F" then
										if rstAux("pend_fact")<>0 then
											total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
										end if
									end if
									if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										total_pendiente = total_pendiente + null_z(rstAux("deuda"))
									end if
								end if
							elseif rstAux("Tipo")="V" then
								if rstAux("divisa")<>moneda_base then
									divisa_ant=rstAux("divisa")
									divisa_act=rstAux("divisa")
									while divisa_ant=divisa_act and not rstAux.eof
										if rstAux("Tipo")="V" then
											total_valorR = total_valorR + null_z(rstAux("Total"))
											total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
										end if
										rstAux.movenext
										NumRegs=NumRegs+1
										if not rstAux.eof then
											divisa_act=rstAux("divisa")
										end if
									wend
									if divisa_ant<>moneda_base then
										total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
										total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
									end if
									'le decimos que hay un registro activo
									pasado=1
								else
									total_valorR = total_valorR + null_z(rstAux("Total"))
									total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
								end if
							end if
						end if
						'si hay un registro activo no hacemos el movenext
						if pasado=0 then
							if rstAux("Tipo")="F" then
								if rstAux("pend_fact")<>0 then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								else
									NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
								end if
							else
								if rstAux("Tipo")="V" then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									NumRegs=NumRegs+1
								end if
							end if
							rstAux.movenext
							pasado=0
						else
							pasado=0
						end if
						Gtotal_valor = Gtotal_valor + total_valor
						Gtotal_pendiente = Gtotal_pendiente + total_pendiente
						total_valor=0
						total_pendiente=0
						Gtotal_valorR = Gtotal_valorR + total_valorR
						Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
						total_valorR=0
						total_pendienteR=0
					wend
					rstAux.close
					DrawFila color_fondo
					documento_contado=""
					if NumRegs=0 then NumRegs=1
					if agrupar_comercial>"" then
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_blau & "'","","",0,""
					end if
					if agrupar_poblacion>"" then
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
					end if
					DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsClienteAnadir) & "</b>"
					DrawCelda "TDBORDECELDA7","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotalCliente & "(" & EncodeForHtml(abreviaturaMB) & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(Gtotal_pendiente + Gtotal_pendienteR,n_decimalesMB,-1,0,-1)) & "</b></td></tr></table>"
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7","","",0,""
                    CloseFila
					NumRegs=0
					NumRegsClienteAnadir=0
					if comercial & "">"" then
						texto_select="ORDER BY comercial," & strwhere41 & "rsocial,Ndoc,Nvto"
					else
						texto_select="ORDER BY " & strwhere41 & "rsocial,divisa,Ndoc,Nvto"
					end if
					if instr(1,seleccion,texto_select,1)>0 then
						seleccion=mid(seleccion,1,len(seleccion)-len(texto_select))
					else
						seleccion=mid(seleccion,1,len(seleccion)-len("ORDER BY " & strwhere41 & "rsocial,Ndoc,Nvto"))
					end if
					seleccion=seleccion & texto_select
					total_valor=0
					total_pendiente=0
					Gtotal_valor=0
					Gtotal_pendiente=0
					total_valorR=0
					total_pendienteR=0
					Gtotal_valorR=0
					Gtotal_pendienteR=0
					if rstAux.state<>0 then rstAux.close
					rstAux.cursorlocation=3
					rstAux.Open seleccion2,session("backendlistados")
					NumRegsTotalPendFacQuitar=0
					while not rstAux.eof
						if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
							if rstAux("divisa")<>moneda_base then
								divisa_ant=rstAux("divisa")
								divisa_act=rstAux("divisa")
								while divisa_ant=divisa_act and not rstAux.eof
									if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_valor = total_valor + null_z(rstAux("Total"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_valor = total_valor + null_z(rstAux("Total"))
										end if
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_pendiente = total_pendiente + null_z(rstAux("deuda"))
										end if
									end if
									if rstAux("Tipo")="V" then
										total_valorR = total_valorR + null_z(rstAux("Total"))
										total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
									end if
									if rstAux("Tipo")="F" then
										if rstAux("pend_fact")<>0 then
											NumRegs=NumRegs+1
										end if
										documento_contado=rstAux("ndoc")
									else
										if rstAux("Tipo")="V" then
											if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
												NumRegs=NumRegs+1
												documento_contado=rstAux("ndoc")
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											NumRegs=NumRegs+1
										end if
									end if
									rstAux.movenext
									if not rstAux.eof then
										divisa_act=rstAux("divisa")
									end if
								wend
								if divisa_ant<>moneda_base then
									total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
									total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
									total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
									total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
								end if
								'le decimos que hay un registro activo
								pasado=1
						   	else
								if rstAux("Tipo")="F" then
									if rstAux("pend_fact")<>0 then
										total_valor = total_valor + null_z(rstAux("Total"))
									end if
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									total_valor = total_valor + null_z(rstAux("Total"))
								end if
								if rstAux("Tipo")="F" then
									if rstAux("pend_fact")<>0 then
										total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
									end if
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									total_pendiente = total_pendiente + null_z(rstAux("deuda"))
								end if
							end if
						elseif rstAux("Tipo")="V" then
							if rstAux("divisa")<>moneda_base then
								divisa_ant=rstAux("divisa")
								divisa_act=rstAux("divisa")
								while divisa_ant=divisa_act and not rstAux.eof
									if rstAux("Tipo")="V" then
										total_valorR = total_valorR + null_z(rstAux("Total"))
										total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
									end if
									rstAux.movenext
									NumRegs=NumRegs+1
									if not rstAux.eof then
										divisa_act=rstAux("divisa")
									end if
								wend
								if divisa_ant<>moneda_base then
									total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
									total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
								end if
								'le decimos que hay un registro activo
								pasado=1
						   	else
								total_valorR = total_valorR + null_z(rstAux("Total"))
								total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
							end if
						end if
						'si hay un registro activo no hacemos el movenext
						if pasado=0 and (not rstAux.eof) then
							if rstAux("Tipo")="F" then
								if rstAux("pend_fact")<>0 then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								else
									NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
								end if
							else
								if rstAux("Tipo")="V" then
									NumRegs=NumRegs+1
									documento_contado=rstAux("ndoc")
								end if
								if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									NumRegs=NumRegs+1
								end if
							end if
							rstAux.movenext
							pasado=0
						else
							pasado=0
						end if
						Gtotal_valor = Gtotal_valor + total_valor
						Gtotal_pendiente = Gtotal_pendiente + total_pendiente
						total_valor=0
						total_pendiente=0
						Gtotal_valorR = Gtotal_valorR + total_valorR
						Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
						total_valorR=0
						total_pendienteR=0
					wend
					documento_contado=""
					DrawFila color_fondo
					if agrupar_comercial>"" then
						DrawCelda "TDBORDECELDA7 colspan=2","","",0,"<b>" & LitRegistros & " " & LitTotal & ": " & EncodeForHtml(NumRegTotal-NumRegsTotalPendFacQuitar+cuantos_tenemos_que_sumar) & "</b>"
					else
						DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitRegistros & " " & LitTotal & ": " & EncodeForHtml(NumRegTotal-NumRegsTotalPendFacQuitar+cuantos_tenemos_que_sumar) & "</b>"
					end if
					DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotales & "(" & EncodeForHtml(abreviaturaMB) & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(Gtotal_pendiente + Gtotal_pendienteR,n_decimalesMB,-1,0,-1)) & "</b></td></tr></table>"
					if agrupar_poblacion>"" then
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
					end if
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7","","",0,""
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
					DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
                    DrawCelda "TDBORDECELDA7","","",0,""
					CloseFila
					if d_lookup("imp_equiv","configuracion","nempresa='" & sesionNCliente & "'",session("backendlistados")) then
						DrawFila color_fondo
						DrawCelda "TDBORDECELDA7","","",0,""
						Gtotal_pendientePTS  = cdbl(CambioDivisa(null_z(Gtotal_pendiente),moneda_base,sesionNCliente & "01"))
						Gtotal_pendienteRPTS = cdbl(CambioDivisa(null_z(Gtotal_pendienteR),moneda_base,sesionNCliente & "01"))
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotales & "(" & EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados"))) & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(Gtotal_pendientePTS+Gtotal_pendienteRPTS,d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b></td></tr></table>"
						if agrupar_poblacion>"" then
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						end if
						DrawCelda "TDBORDECELDA7","","",0,""
						DrawCelda "TDBORDECELDA7","","",0,""
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(Gtotal_valor,moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(Gtotal_valorR,moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(Gtotal_pendiente),moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(formatnumber(CambioDivisa(null_z(Gtotal_pendienteR),moneda_base,sesionNCliente & "01"),d_lookup("ndecimales","divisas","codigo='" & sesionNCliente & "01'",session("backendlistados")),-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7","","",0,""
                        CloseFila
					end if
					NumRegs=0
				end if
			end if

			if agrupar_comercial>"" then
				where = "" & " comercial='" & ComercialAnt & "' and ncliente='" & ClienteAnt & "'"
				strSelect = seleccion2 & " where " & where & ordenar
				if rstAux.state<>0 then rstAux.close
				rstAux.cursorlocation=3
				rstAux.Open strSelect,session("backendlistados")
				'rstAux.movefirst
				numr=0
				while not rstAux.eof
					numr=numr+1
					rstAux.movenext
				wend
				rstAux.close
				if not rst.eof and NumRegs<>numr then
					rst.movenext
				end if

				if not rst.EOF and fila>MAXPAGINA and NumRegs=numr then
					DrawFila ""
					if ComercialAnt & ""<>"" then
						total_valor=0
						total_pendiente=0
						Gtotal_valor=0
						Gtotal_pendiente=0
						if instr(1,seleccion,"ORDER BY comercial," & strwhere41 & "rsocial,divisa,Ndoc,Nvto",1)>0 then
							seleccion=mid(seleccion,1,len(seleccion)-len("ORDER BY comercial," & strwhere41 & "rsocial,divisa,Ndoc,Nvto"))
						else
							seleccion=mid(seleccion,1,len(seleccion)-len("ORDER BY comercial," & strwhere41 & "rsocial,Ndoc,Nvto"))
						end if
						seleccion=seleccion & "ORDER BY comercial," & strwhere41 & "rsocial,divisa,Ndoc,Nvto"

						if rstAux.state<>0 then rstAux.close
						rstAux.cursorlocation=3
						rstAux.Open seleccion,session("backendlistados")
						NumRegsTotalPendFacQuitar=0
						while not rstAux.eof
							if rstAux("comercial") & ""=ComercialAnt & "" then
								if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									if rstAux("divisa")<>moneda_base then
										divisa_ant=rstAux("divisa")
										divisa_act=rstAux("divisa")
										while divisa_ant=divisa_act and not rstAux.eof
											if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_valor = total_valor + null_z(rstAux("Total"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_valor = total_valor + null_z(rstAux("Total"))
												end if
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_pendiente = total_pendiente + null_z(rstAux("deuda"))
												end if
											end if
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													NumRegs=NumRegs+1
												end if
												documento_contado=rstAux("ndoc")
											else
												if rstAux("Tipo")="V" then
													if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
														NumRegs=NumRegs+1
														documento_contado=rstAux("ndoc")
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													NumRegs=NumRegs+1
												end if
											end if
											rstAux.movenext
											if not rstAux.eof then
												divisa_act=rstAux("divisa")
											end if
										wend
										if divisa_ant<>moneda_base then
											total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
											total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
											'le decimos que hay un registro activo
										end if
										pasado=1
								   	else
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_valor = total_valor + null_z(rstAux("Total"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_valor = total_valor + null_z(rstAux("Total"))
										end if
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_pendiente = total_pendiente + null_z(rstAux("deuda"))
										end if
									end if
								elseif rstAux("Tipo")="V" then
									if rstAux("divisa")<>moneda_base then
										divisa_ant=rstAux("divisa")
										divisa_act=rstAux("divisa")
										while divisa_ant=divisa_act and not rstAux.eof
											if rstAux("Tipo")="V" then
												total_valorR = total_valorR + null_z(rstAux("Total"))
												total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
											end if
											rstAux.movenext
											NumRegs=NumRegs+1
											if not rstAux.eof then
												divisa_act=rstAux("divisa")
											end if
										wend
										if divisa_ant<>moneda_base then
											total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
											total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
											'le decimos que hay un registro activo
										end if
										pasado=1
								   	else
										total_valorR = total_valorR + null_z(rstAux("Total"))
										total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
									end if
								end if
							end if
							'si hay un registro activo no hacemos el movenext
							if pasado=0 then
								if rstAux("Tipo")="F" then
									if rstAux("pend_fact")<>0 then
										NumRegs=NumRegs+1
										documento_contado=rstAux("ndoc")
									else
										NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
									end if
								else
									if rstAux("Tipo")="V" then
										if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
											NumRegs=NumRegs+1
											documento_contado=rstAux("ndoc")
										end if
									end if
									if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										NumRegs=NumRegs+1
									end if
								end if
								rstAux.movenext
								pasado=0
							else
								pasado=0
							end if
							Gtotal_valor = Gtotal_valor + total_valor
							Gtotal_pendiente = Gtotal_pendiente + total_pendiente
							total_valor=0
							total_pendiente=0
							Gtotal_valorR = Gtotal_valorR + total_valorR
							Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
							total_valorR=0
							total_pendienteR=0
						wend
						rstAux.close
						documento_contado=""
						if NumRegs=0 then NumRegs=1
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "' colspan=2","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsComercialAnadir) & "</b>"
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<b>" + LitTotalComercial & "(" & EncodeForHtml(abreviaturaMB) & ")</b>"
						if agrupar_poblacion>"" then
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						end if
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
                        DrawCelda "TDBORDECELDA7","","",0,""
						NumRegs=0
						poblacionant=""
						NumRegsComercialAnadir=0
						CloseFila
					end if
				end if
			else
				where = " ncliente='"&ClienteAnt&"'"
				strSelect = seleccion2 & " where " & where & ordenar
				if rstAux.state<>0 then rstAux.close
				rstAux.cursorlocation=3
				rstAux.Open strSelect,session("backendlistados")
				numr=0
				while not rstAux.eof
					numr=numr+1
					rstAux.movenext
				wend
				rstAux.close
				if not rst.eof and NumRegs<>numr then
					rst.movenext
				end if

				if not rst.EOF and fila>MAXPAGINA and NumRegs=numr then
					DrawFila ""
					if ClienteAnt & ""<>"" then
						total_valor=0
						total_pendiente=0
						Gtotal_valor=0
						Gtotal_pendiente=0
						if comercial & "">"" then
							texto_select="ORDER BY comercial," & strwhere41 & "rsocial,divisa,Ndoc,Nvto"
						else
							texto_select="ORDER BY " & strwhere41 & "rsocial,divisa,Ndoc,Nvto"
						end if
						if instr(1,seleccion,texto_select,1)>0 then
							seleccion=mid(seleccion,1,len(seleccion)-len(texto_select))
						else
							seleccion=mid(seleccion,1,len(seleccion)-len("ORDER BY " & strwhere41 & "rsocial,Ndoc,Nvto"))
						end if
''ricardo 22/12/2009 se incluyen los efectos
seleccion=replace(seleccion,"''","'")
						seleccion=seleccion & texto_select
''ricardo 22/12/2009 se incluyen los efectos
seleccion=replace(seleccion,"ORORDER","ORDER")
						if rstAux.state<>0 then rstAux.close
						rstAux.cursorlocation=3
						rstAux.Open seleccion,session("backendlistados")
						NumRegsTotalPendFacQuitar=0
						while not rstAux.eof
							if rstAux("ncliente") & ""=ClienteAnt & "" then
								if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
									if rstAux("divisa")<>moneda_base then
										divisa_ant=rstAux("divisa")
										divisa_act=rstAux("divisa")
										while divisa_ant=divisa_act and not rstAux.eof
											if rstAux("Tipo")="F" or rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_valor = total_valor + null_z(rstAux("Total"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_valor = total_valor + null_z(rstAux("Total"))
												end if
												if rstAux("Tipo")="F" then
													if rstAux("pend_fact")<>0 then
														total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													total_pendiente = total_pendiente + null_z(rstAux("deuda"))
												end if
											end if
											if rstAux("Tipo")="F" then
												if rstAux("pend_fact")<>0 then
													NumRegs=NumRegs+1
												end if
												documento_contado=rstAux("ndoc")
											else
												if rstAux("Tipo")="V" then
													if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
														NumRegs=NumRegs+1
														documento_contado=rstAux("ndoc")
													end if
												end if
												if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
													NumRegs=NumRegs+1
												end if
											end if
											rstAux.movenext
											if not rstAux.eof then
												divisa_act=rstAux("divisa")
											end if
										wend
										if divisa_ant<>moneda_base then
											total_valor = CambioDivisa(total_valor,divisa_ant,moneda_base)
											total_pendiente = CambioDivisa(total_pendiente,divisa_ant,moneda_base)
											'le decimos que hay un registro activo
										end if
										pasado=1
								   	else
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_valor = total_valor + null_z(rstAux("Total"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_valor = total_valor + null_z(rstAux("Total"))
										end if
										if rstAux("Tipo")="F" then
											if rstAux("pend_fact")<>0 then
												total_pendiente = total_pendiente + null_z(rstAux("pend_fact"))
											end if
										end if
										if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
											total_pendiente = total_pendiente + null_z(rstAux("deuda"))
										end if
									end if
								elseif rstAux("Tipo")="V" then
									if rstAux("divisa")<>moneda_base then
										divisa_ant=rstAux("divisa")
										divisa_act=rstAux("divisa")
										while divisa_ant=divisa_act and not rstAux.eof
											if rstAux("Tipo")="V" then
												total_valorR = total_valorR + null_z(rstAux("Total"))
												total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
											end if
											rstAux.movenext
											NumRegs=NumRegs+1
											if not rstAux.eof then
												divisa_act=rstAux("divisa")
											end if
										wend
										if divisa_ant<>moneda_base then
											total_valorR = CambioDivisa(total_valorR,divisa_ant,moneda_base)
											total_pendienteR = CambioDivisa(total_pendienteR,divisa_ant,moneda_base)
											'le decimos que hay un registro activo
										end if
										pasado=1
								   	else
										total_valorR = total_valorR + null_z(rstAux("Total"))
										total_pendienteR = total_pendienteR + null_z(rstAux("deuda"))
									end if
								end if
							end if
							'si hay un registro activo no hacemos el movenext
							if pasado=0 then
								if rstAux("Tipo")="F" then
									if rstAux("pend_fact")<>0 then
										NumRegs=NumRegs+1
										documento_contado=rstAux("ndoc")
									else
										NumRegsTotalPendFacQuitar=NumRegsTotalPendFacQuitar+1
									end if
								else
									if rstAux("Tipo")="V" then
										if (documento_contado="" or instr(1,rstAux("Nvto"),documento_contado,1)=0) then
											NumRegs=NumRegs+1
											documento_contado=rstAux("ndoc")
										end if
									end if
									if rstAux("Tipo")="A" or rstAux("Tipo")="T" or rstAux("Tipo")="E" then
										NumRegs=NumRegs+1
									end if
								end if
								rstAux.movenext
								pasado=0
							else
								pasado=0
							end if
							Gtotal_valor = Gtotal_valor + total_valor
							Gtotal_pendiente = Gtotal_pendiente + total_pendiente
							total_valor=0
							total_pendiente=0
							Gtotal_valorR = Gtotal_valorR + total_valorR
							Gtotal_pendienteR = Gtotal_pendienteR + total_pendienteR
							total_valorR=0
							total_pendienteR=0
						wend
						rstAux.close
						documento_contado=""
						if NumRegs=0 then NumRegs=1
						if agrupar_comercial>"" then
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_blau & "'","","",0,""
						end if
						if agrupar_poblacion>"" then
							DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						end if
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<b>" & LitRegistros & ": " & EncodeForHtml(NumRegs+NumRegsClienteAnadir) & "</b>"
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,"<table width='100%'><tr><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='left'><b>" + LitTotalCliente & "(" & EncodeForHtml(abreviaturaMB) & ")" & "</b></td><td class='TDBORDECELDA7' style='BORDER: 0px solid Black;' align='right'><b>" & EncodeForHtml(formatnumber(Gtotal_pendienteR + Gtotal_pendienteR,n_decimalesMB,-1,0,-1)) & "</b></td></tr></table>"
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						DrawCelda "TDBORDECELDA7 bgcolor='" & color_fondo & "'","","",0,""
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valor,n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(Gtotal_valorR,n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendiente),n_decimalesMB,-1,0,-1)) & "</b>"
						DrawCelda "TDBORDECELDA7 ALIGN=RIGHT bgcolor='" & color_fondo & "'","","",0,"<b>" & EncodeForHtml(formatnumber(null_z(Gtotal_pendienteR),n_decimalesMB,-1,0,-1)) & "</b>"
                        DrawCelda "TDBORDECELDA7","","",0,""
						NumRegs=0
						NumRegsClienteAnadir=0
						CloseFila
					end if
				end if
			end if%>
			</table>
			<%rst.Close
		end if%>
	      <hr/>
		<%NavPaginas lote,lotes,campo,criterio,texto,2%>
		<input type="hidden" name="NumRegs" value="<%=EncodeForHtml(NumRegs)%>">
		<input type="hidden" name="NumRegsTotal" value="<%=EncodeForHtml(NumRegsTotal)%>">
		<input type="hidden" name="NumRegsTotalAnadir" value="<%=EncodeForHtml(NumRegsTotalAnadir)%>">
		<input type="hidden" name="NumRegsClienteAnadir" value="<%=EncodeForHtml(NumRegsClienteAnadir)%>">
		<input type="hidden" name="NumRegsComercialAnadir" value="<%=EncodeForHtml(NumRegsComercialAnadir)%>">
        <%''ricardo 25-5-2006 comienzo de la select
        ''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
        auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"fin_listado_cobros"
	end if%>
	</form>
<%end if
connRound.close
set connRound = Nothing
set rstAux= Nothing
set rstSelect= Nothing
set rst= Nothing
set rstAux= Nothing
set rstAux2= Nothing
set rstCliente= Nothing
set rstVencimientos= Nothing%>
</body>
</html>