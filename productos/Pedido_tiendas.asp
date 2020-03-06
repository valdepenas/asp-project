<%@ Language=VBScript %>
<%
' JCI 17/06/2003 : MIGRACION A MONOBASE
'
' JCI 27/01/2004 : Gestión de lotes
%>
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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
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

<!--#include file="Pedido_tiendas.inc" -->
<!--#include file="../ventas/documentos.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/calendar.inc" -->

<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->  

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('AddDG', 'fade=1')
    animatedcollapse.addDiv('BrowseDG', 'fade=1')
    //animatedcollapse.addDiv('BrowseCab', 'fade=1')
    animatedcollapse.addDiv('CABECERA', 'fade=1')
    animatedcollapse.addDiv('DETALLES', 'fade=1')

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()

</script>

<script language="javascript" type="text/javascript">
function cambiarfecha(fecha,modo)
{
	var fecha_ar=new Array();

	if (fecha!="")
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
			window.alert("<%=LitFechaMal%> en el campo " + modo );
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
				window.alert("<%=LitFechaMal%> en el campo " + modo);
				return false;
			}
		}
	}
	return true;
}



//***************************************************************************
function Editar(albaran)
{
	document.pedido_tiendas.action="pedido_tiendas.asp?npedido=" + albaran + "&mode=browse";
	document.pedido_tiendas.submit();
	parent.botones.document.location="pedido_tiendas_bt.asp?mode=browse";
}

//***************************************************************************
function Mas(sentido,lote,campo,criterio,texto)
{
	document.location="pedido_tiendas.asp?mode=search&viene=" + document.pedido_tiendas.viene.value + "&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&mmp=" +document.pedido_tiendas.mmp.value ;
}
//***************************************************************************
/*
if (window.document.addEventListener) {
    window.document.addEventListener("keydown", callkeydownhandler, false);
} else {
    window.document.attachEvent("onkeydown", callkeydownhandler);
}
function callkeydownhandler(evnt) {
    ev = (evnt) ? evnt : event;
    //comprobar_enter(ev);
    keypress2(ev);
}

function keypress2(e){
	tecla=e.keyCode;
	keyPressed(tecla);
}

//Comprueba si la tecla pulsada es CTRL+S. Si es así guarda el registro.
function keyPressed(tecla)
{
	if (tecla==<%=TeclaGuardar%>)
	{ //CTRL+S
		if (document.pedido_tiendas.mode.value=="add" || document.pedido_tiendas.mode.value=="edit")
		{
			if (document.pedido_tiendas.fecha.value=="")
			{
				window.alert("<%=LitMsgFechaNoNulo%>");
				return;
			}

			if (!cambiarfecha(document.pedido_tiendas.fecha.value,"FECHA PEDIDO")) return false;

			if (!checkdate(document.pedido_tiendas.fecha))
			{
				window.alert("<%=LitMsgFechaFecha%>");
				return;
			}

			if (document.pedido_tiendas.responsable.value=="")
			{
				window.alert("<%=LitMsgResponsableNoNulo%>");
				return false;
			}

			if (document.pedido_tiendas.nserie.value=="")
			{
				window.alert("<%=LitMsgSerieNoNulo%>");
				return;
			}

			if (document.pedido_tiendas.almdestino.value=="")
			{
				window.alert("<%=LitMsgAlmDestinoNoNulo%>");
				return false;
			}
			if (document.pedido_tiendas.almorigen.value=="")
			{
				window.alert("<%=LitMsgAlmOrigenNoNulo%>");
				return false;
			}

			if (document.pedido_tiendas.almorigen.value==document.pedido_tiendas.almdestino.value)
			{
				window.alert("<%=LitMsgAlmOAlmDIguales%>");
				return false;
			}			
			switch (document.pedido_tiendas.mode.value)
			{
				case "add":
					document.pedido_tiendas.action="pedido_tiendas.asp?mode=first_save";
					break;

				case "edit":
					document.pedido_tiendas.action="pedido_tiendas.asp?mode=save&ndoc=" + document.pedido_tiendas.h_npedido.value;
					break;
			}
			document.pedido_tiendas.submit();
			parent.botones.document.location="pedido_tiendas_bt.asp?mode=browse";
		}
		//else { //Mode=browse.
		//}
	}
}
*/
//***************************************************************************

function TraerResponsable()
{
	document.pedido_tiendas.action="pedido_tiendas.asp?responsable=" + document.pedido_tiendas.responsable.value + "&mode=traerresponsable&submode=" + document.pedido_tiendas.mode.value + "&npedido=" + document.pedido_tiendas.h_npedido.value;
	document.pedido_tiendas.submit();
}

function MasDet(sentido,lote,firstReg,lastReg,campo,criterio,texto,firstRegAll,lastRegAll)
{
	fr_Detalles.document.pedido_tiendas_det.action="pedido_tiendas_det.asp?mode=browse&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&firstReg=" + firstReg + "&lastReg=" + lastReg + "&firstRegAll=" + firstRegAll + "&lastRegAll=" + lastRegAll;
	fr_Detalles.document.pedido_tiendas_det.submit();
}
function ocultar_genfact(){
	document.getElementById("genmov").style.display="none";
}
function abrir_detalles(npedido,fila,viene,ep){
	pagina="../central.asp?pag1=mantenimiento/conv_ped_alb_completar.asp&ndoc=" + npedido + "&tdocumento=" + fila + "&viene=" + viene + "&ep=" + ep + "&pag2=mantenimiento/conv_ped_alb_completar_bt.asp";
	AbrirVentana(pagina,'P',<%=altoventana%>,<%=anchoventana%>);
}

function GenerarMovimiento(pedido,ep) {
	if (document.pedido_tiendas.h_nmovimiento.value=='' && document.pedido_tiendas.h_nalbaran.value=='')
	{
		if (document.pedido_tiendas.serieMOV.value=='') alert("<%=LitMsgSerieNoNulo%>");
		else 
		{
		    abrir_detalles(pedido,0,"ventasPedTi",ep);
			//document.pedido_tiendas.action="pedido_tiendas.asp?npedido=" + pedido + "&mode=genera&SerieMov=" + document.albaranes_cli.serieMov.value;
			//document.pedido_tiendas.submit();
			//document.pedido_tiendas.all("waitBoxOculto").style.visibility="visible";

		}
	}
	else alert("<%=LitMsgPedTieneMov%>");
}
function Redimensionar()
{
    var alto = 0;
    if (parent.document.body.offsetHeight) alto = parent.document.body.offsetHeight;
    else alto = parent.self.innerHeight;
	if (document.getElementById("frDetalles").style.display == "")
	{
	    if (alto > 140)
        {
            if (alto - 285 > 140) document.getElementById("frDetalles").style.height = alto - 285;
            else document.getElementById("frDetalles").style.height = 140;
        }
        else document.getElementById("frDetalles").style.height = 140;
    }

}
</script>
<%
modoPantalla=EncodeForHtml(Request.QueryString("mode"))
if modoPantalla & ""="" then
    modoPantalla=EncodeForHtml(request.Form("mode"))
end if
CuandoRedimensionar=0
if modoPantalla="browse" or modoPantalla="save" or modoPantalla="first_save" then
    CuandoRedimensionar=1
end if
%>
<%mode=EncodeForHtml(Request.QueryString("mode"))
if mode = "browse" then %>
	<body class="BODY_ASP" <%=iif(CuandoRedimensionar=1, "onresize='javascript:Redimensionar();'","")%>>
<%else%>
	<body onload="self.status='';" class="BODY_ASP" <%=iif(CuandoRedimensionar=1, "onresize='javascript:Redimensionar();'","")%> >
<%end if


'**
'* Obtiene los números de serie de un detalle de un documento.
'* ndocumento: Indica el nº de documento a buscar.
'* ndetalle: Indica el nº de item del documento.
'* return: La lista de nº de serie del item del documento.
'**

'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(npedido,nserie,fecha,mode)
	if npedido & ""="" then
		'Crear un nuevo registro.
		rst.AddNew
		SigDoc=CalcularNumDocumento(nserie,fecha)
		rst("npedido")=SigDoc
		'npedido=SigDoc
	end if

	'Asignar los nuevos valores a los campos del recordset.
    '19/02/2020: Se comprueba que la cadena esté vacía
    if request.form("nserie") & "">"" then
	    rst("nserie")=limpiaCadena(Nulear(request.form("nserie")))
    else
        %><script language="javascript" type="text/javascript">
            window.alert("<%=LitMsgSerieNoNulo%>");
            parent.botones.location="pedido_tiendas_bt.asp?mode=edit"
		</script><%
		
    end if
        

	'Detectar cambios en el almdestino para recalcular los stock de los detalles
	if npedido & "">"" and rst("almdestino")<>(Request.Form("almdestino")&"") then
		'Primero se elimina el stock y los equipos del almdestino anterior
		dim listaNS

		'contamos las cantidades de equipos que hay en este movimiento

		'hacemos las operaciones
		rstAux.cursorlocation=3
		rstAux.open "select ref,almorigen,cantidad,almdestino,nitem from pedidos_tienda as m with (NOLOCK),detalles_ped_tienda as d with (NOLOCK) where m.npedido like '" & session("ncliente") & "%' and d.npedido like '" & session("ncliente") & "%' and m.npedido=d.npedido and d.npedido='" & npedido & "'",session("dsn_cliente")
		while not rstAux.EOF
			refST=rstAux("ref")
			almSTO=rstAux("almorigen")
			almSTD=rstAux("almdestino")
			canST=rstAux("cantidad")

		    'ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTO,canST,"",session("dsn_cliente")
			'ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTD,-canST,"",session("dsn_cliente")
			rstAux.movenext
		wend
		rstAux.close
		'---------------------
		'Ahora se añade el stock y los equipos al almdestino nuevo
	end if
	rst("almdestino")=limpiaCadena(Nulear(Request.Form("almdestino")))
	rst("almorigen")=limpiaCadena(Nulear(Request.Form("almorigen")))
	rst("responsable")=session("ncliente") + limpiaCadena(Nulear(request.form("responsable")))
    '18/02/2020: Se comprueba que la cadena esté vacía
    if request.form("observaciones") & "">"" then
	    rst("observaciones")=limpiaCadena(Nulear(request.form("observaciones")))
    else
		rst("observaciones")=""
    end if
	rst("fecha")=limpiaCadena(Nulear(Request.Form("fecha")))
	'dgb 25/03/2008: anyadir estado_ped
	if request.form("estado_ped") & "">"" then
		rst("estado_ped")=limpiaCadena(Nulear(request.form("estado_ped")))
	else
		rst("estado_ped")=""
	end if
	'rst("npedido")=iif(cstr(Request.Form("merric") & "")="-1" or ucase(cstr(Request.Form("merric") & ""))="ON",1,0)
	rst.Update
end sub

'******************************************************************************
'Elimina los datos del registro cuando se pulsa BORRAR.
sub BorrarRegistro(npedido)
	'Miramos si se va a borrar el último generado y si es así se descuenta el contador de documentos
	ano=right(cstr(year(d_lookup("fecha","pedidos_tienda","npedido like '" & session("ncliente") & "%' and npedido='" & npedido & "'",session("dsn_cliente")))),2)
	nserie=d_lookup("nserie","pedidos_tienda","npedido like '" & session("ncliente") & "%' and npedido='" & npedido & "'",session("dsn_cliente"))

	rstAux.Open "select * from series with(rowlock) where nserie like '" & session("ncliente") & "%' and nserie='" & nserie & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if not rstAux.eof then
		ultimo=rstAux("contador")
		UltimoDocumento=nserie+ano+completar(trim(cstr(ultimo)),6,"0")
		if npedido=UltimoDocumento then
			rstAux("contador")=ultimo-1
			rstAux.update
		end if
	end if
	rstAux.close

	'Luego se eliminan los detalles
	rstAux.open "delete from detalles_ped_tienda where npedido like '" & session("ncliente") & "%' and npedido='" & npedido & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

	'Después la cabecera del pedido
	rstAux.open "delete from pedidos_tienda where npedido like '" & session("ncliente") & "%' and npedido='" & npedido & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
end sub

sub convertirPedidosMov
	DSNCliente=session("dsn_cliente")
	donde=instr(1,DSNCliente,"Initial Catalog=",1)
	donde2=instr(donde,DSNCliente,";",1)
	BD=mid(DSNCliente,donde+16,donde2-donde-16)
	donde=instr(1,DSNCliente,"User id=",1)
	donde2=instr(donde,DSNCliente,";",1)
	Usuario=mid(DSNCliente,donde+8,donde2-donde-8)
	eliminar="if exists (select * from sysobjects where id = object_id('" & session("usuario") & "_temporal') and sysstat " & _
			 " & 0xf = 3) drop table [" & session("usuario") & "_temporal]"
	rst.open eliminar,session("dsn_cliente"),adUseClient,adLockReadOnly
	crear2="CREATE TABLE [" & session("usuario") & "_temporal] (ncliente char(10),nclienteAdmin varchar(15))"
	rst.open crear2,session("dsn_cliente"),adUseClient,adLockReadOnly
	GrantUser session("usuario")&"_temporal", session("dsn_cliente")
	rst.open "insert into " & BD & "." & Usuario & ".[" & session("usuario") & "_temporal] select c.ncliente as ncliente,ca.ncliente as nclienteAdmin from "&BD&"..clientes c , clientes ca where c.cifedi=ca.cifedi and c.ncliente like '"&session("ncliente")&"%' ",DSNIlion
	
	conn.open session("dsn_cliente")
	strselect="EXEC Convertir_PedTi_Mov @nusuario ='"&session("usuario")&"',@session_ncliente ='"&session("ncliente")&"',@serie= '"&serieMov&"',@unificar =0 "

	set rstSelect2 = conn.execute(strselect)
	Primero=rstSelect2(0)
	Primero=rstSelect2(0)
	Ultimo=rstSelect2(1)
	error=rstSelect2(2)
	cadenaAuditoria=rstSelect2(3)

	if Primero>"" and Ultimo>"" and error=0 then
		auditar_ins_bor session("usuario"),cadenaAuditoria,"","alta","","","conver_pedTI_MOV_Masivo"
		%><script language="javascript" type="text/javascript">
			window.alert("<%=LitMsgMovimientosGenerados%> <%=enc.EncodeForJavascript(trimcodEmpresa(Primero))%> al <%=enc.EncodeForJavascript(trimcodEmpresa(Ultimo))%>");
		</script><%
	elseif error=1 then
		%><script language="javascript" type="text/javascript">
			window.alert("<%=LitMsgSinDireccionPrinc%>");
		</script><%
	elseif error=2 then
		%><script language="javascript" type="text/javascript">
			window.alert("<%=LitMsgNumSeriesRepetidos%>");
		</script><%
	elseif error=3 then
		%><script language="javascript" type="text/javascript">
			window.alert("<%=LitMsgDocExistRevContConv%>");
		</script><%
	else
		%><script language="javascript" type="text/javascript">
			window.alert("<%=LitMsgMovimientosNoGenerados%>");
		</script><%
	end if
	conn.close
end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0 
	
	set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeout = 300
	conn.CommandTimeout = 300

 %>
<form name="pedido_tiendas" method="post"><%
    PintarCabecera "pedido_tiendas.asp"
	'Leer parámetros de la página
	mode=Request.QueryString("mode")
	WaitBoxOculto LitEsperePorFavor
	dim bmr,mmt,mmp

	ObtenerParametros("pedido_tiendas")

	p_npedido=limpiaCadena(Request.QueryString("npedido"))
	if p_npedido="" then p_npedido=limpiaCadena(Request.QueryString("ndoc"))
	if p_npedido>"" then CheckCadena p_npedido

	campo=limpiaCadena(Request.QueryString("campo"))
	criterio=limpiaCadena(Request.QueryString("criterio"))
	texto=limpiaCadena(Request.QueryString("texto"))
	if Request.QueryString("fecha")>"" then
		p_fecha=limpiaCadena(Request.QueryString("fecha"))
	else
		p_fecha=limpiaCadena(Request.form("fecha"))
	end if

	viene=limpiaCadena(request.querystring("viene"))
	if viene="" then viene=limpiaCadena(request.form("viene"))

	if request.querystring("submode")>"" then
		submode=request.querystring("submode")
	else
		submode=request.form("submode")
	end if
	
	if request.querystring("serieMOV")>"" then
		serieMOV=limpiaCadena(request.querystring("serieMOV"))
	else
		serieMOV=request.form("serieMOV")
	end if
	
	if request.querystring("observaciones")>"" then
		tmp_observaciones=limpiaCadena(request.querystring("observaciones"))
	else
		tmp_observaciones=limpiaCadena(request.form("observaciones"))
	end if

	if request.querystring("responsable")>"" then
		tmp_responsable=limpiaCadena(request.querystring("responsable"))
	else
		tmp_responsable=limpiaCadena(request.form("responsable"))
	end if

	if request.querystring("almdestino")>"" then
		tmp_almdestino=limpiaCadena(request.querystring("almdestino"))
	else
		tmp_almdestino=limpiaCadena(request.form("almdestino"))
	end if
	
	if request.querystring("almorigen")>"" then
		tmp_almorigen=limpiaCadena(request.querystring("almorigen"))
	else
		tmp_almorigen=limpiaCadena(request.form("almorigen"))
	end if

	if request.querystring("nserie")>"" then
		p_serie=limpiaCadena(request.querystring("nserie"))
	else
		p_serie=limpiaCadena(request.form("nserie"))
	end if

	if bmr & ""="" then
		if request.QueryString("bmr")& "">"" then
			bmr=limpiaCadena(request.QueryString("bmr"))
		elseif request.form("bmr") & "">"" then
			bmr=limpiaCadena(request.form("bmr"))
		end if
	end if

	if request.QueryString("merric")& "">"" then
		merric=limpiaCadena(request.QueryString("merric"))
	elseif request.form("merric") & "">"" then
		merric=limpiaCadena(request.form("merric"))
	end if

	'***RGU 5/1/2006***
	if request.QueryString("mmp")& "">"" then
		mmp=limpiaCadena(request.QueryString("mmp"))
	elseif request.form("mmp") & "">"" then
		mmp=limpiaCadena(request.form("mmp"))
	end if
	'***

	%>
	<input type="hidden" name="mode" value="<%=EncodeForHtml(null_s(mode))%>">
	<input type="hidden" name="viene" value="<%=EncodeForHtml(null_s(viene))%>">
	<input type="hidden" name="mmp" value="<%=EncodeForHtml(null_s(mmp))%>">
	<%

	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rstAux3 = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")

	if p_serie="" and mode="add" then
		'Obtener la serie por defecto
		p_serie=d_lookup("nserie","series","tipo_documento='PEDIDOS ENTRE ALMACENES' and pordefecto=1 and nserie like '" & session("ncliente") & "%'", session("dsn_cliente"))
	end if

	si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)

	if mode="save" or mode="first_save" then
		'ricardo 7-4-2004 miramos si el personal existe o no
		if submode2<>"traerresponsable" then
			no_proveedor=0
			strselect="select dni,fbaja from personal with (NOLOCK) where dni='" & session("ncliente") + Nulear(request.form("responsable")) & "'"
			rst.cursorlocation=3
			rst.Open strselect,session("dsn_cliente")
			if rst.eof then
				no_proveedor=1
			else
				if rst("fbaja") & "">"" then
					no_proveedor=2
				end if
			end if
			rst.close
		else
			no_proveedor=0
		end if

		if no_proveedor=0 then
			no_encontrado=0

			if no_encontrado=0 then
				strselect="select * from pedidos_tienda where npedido like '" & session("ncliente") & "%' and npedido='" & iif(p_npedido&"">"",p_npedido,"NULL") & "'"
				rst.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

				'if not rst.eof then
						ModDocumento=true
						'comprobamos si el nalbaran existe o no segun el contador de configuracion
						if mode="first_save" then
							if compNumDocNuevo(p_serie,p_fecha,"pedidos_tienda")=0 then
								%><script language="javascript" type="text/javascript">
									window.alert("<%=LitMsgDocExistRevCont%>");
									history.back();
									parent.botones.document.location="pedido_tiendas_bt.asp?mode=add"
								</script><%
								ModDocumento=false
							end if
						end if
						if ModDocumento=true then
							if mode<>"first_save" then
								npediod_old=nz_b(rst("npedido"))
							end if
							GuardarRegistro p_npedido,p_serie,p_fecha,mode
							p_npedido=rst("npedido")
							'InsertarHistorialNserie mensajeTratEquipos,"","","MOVIMIENTOS ENTRE ALMACENES",p_npedido,"","","","","MODIFY",mode
							if mode="first_save" then
								auditar_ins_bor session("usuario"),p_npedido,"","alta","","","pedido_almacenes"
							elseif mode="save" then
							end if
						end if
						ant_mode=mode
						mode="browse"
					
				'end if
				rst.close
			else
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgAlmOAlmDIguales2%>");
					parent.botones.location="pedido_tiendas_bt.asp?mode=edit"
				</script><%
				mode="edit"
			end if
		else
			if mode="first_save" then mode="add" else mode="edit"
			if no_proveedor=1 then
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgResponsableNoExiste%>");
				</script><%
			elseif no_proveedor=2 then
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgResponsableDadoBaja%>");
				</script><%
			end if
			%><script language="javascript" type="text/javascript">
				parent.botones.location="pedido_tiendas_bt.asp?mode=<%=enc.EncodeForJavascript(mode)%>"
			</script><%
		end if
	elseif mode="delete" then
		'comprobamos que no hay ningun detalle con nserie que no sea el ultimo documento

			auditar_ins_bor session("usuario"),p_npedido,"","baja","","","pedido_almacenes"
			
			BorrarRegistro p_npedido
			mode="add"
			p_npedido="" %>
		    <script language="javascript" type="text/javascript">
            //dgb: change to add, refresh search page and open it
		        parent.botones.document.location = "pedido_tiendas_bt.asp?mode=add";
			    SearchPage("pedido_tiendas_lsearch.asp?mode=init",0);			    
		    </script><%
  	end if

	if mode="traerresponsable" then
		submode2="traerresponsable"
		if tmp_responsable> "" then
			responsable=d_lookup("nombre","personal","dni='" & session("ncliente")&tmp_responsable & "'",session("dsn_cliente"))
			if responsable="" then
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgResponsableNoExiste%>");
					//document.pedido_tiendas.action="pedido_tiendas.asp?mode=param&responsable=<%=TmpResponsable2%>";
					//document.pedido_tiendas.submit();
				</script><%
				tmp_responsable=tmp_responsable2
			else
				rstAux3.open "select dni,nombre,fbaja from personal with (NOLOCK) where dni='" & session("ncliente")&tmp_responsable & "' and fbaja is null",session("dsn_cliente")
				if rstAux3.eof then
					%><script language="javascript" type="text/javascript">
						window.alert("<%=LitMsgResponsableDadoBaja%>");
					</script><%
					tmp_responsable=tmp_responsable2
				else
					tmp_responsable=session("ncliente")&tmp_responsable
					TmpNombre=responsable
				end if
				rstAux3.close
			end if
			mode=submode
			%><script language="javascript" type="text/javascript">
				document.pedido_tiendas.mode.value="<%=enc.EncodeForJavascript(mode)%>";
			</script><%
		else
			tmp_responsable=""
			TmpNombre=""
			mode=submode
			%><script language="javascript" type="text/javascript">
				document.pedido_tiendas.mode.value="<%=enc.EncodeForJavascript(mode)%>";
			</script><%
		end if
		if mode="first_save" then
			mode="add"
		else
			mode="edit"
		end if
	end if
	if mode="generaMov"  then	'modo para generar Movimientos'
		convertirPedidosMov
		mode="browse"
	end if
	if mode="browse" or mode="search" or mode="add" then
		if mmt=1 then
			linea1=session("f_tpv")
			linea2=session("f_caja")
			linea3=session("f_empr")

			strSelect = "select c.almacen from tpv a with(nolock), cajas b with(nolock), tiendas c with(nolock) where a.caja=b.codigo and b.tienda=c.codigo and tpv='" & linea1 & "' and b.codigo='" & linea2 & "'"
			rstAux3.cursorlocation=3
			rstAux3.open strSelect,session("dsn_cliente")
			if not rstAux3.eof then
				AlmacenTienda=rstAux3("almacen")
			else
				AlmacenTienda=""
			end if
			rstAux3.close
			if AlmacenTienda & "">"" then
				cadena_mov_sol_usu=" and almdestino='" & AlmacenTienda & "'"
				tmp_almdestino=AlmacenTienda
			end if
		else
			cadena_mov_sol_usu=""
		end if
	end if

	'Mostrar los datos de la página.
	if mode="browse" or mode="edit" then
		if p_npedido="" then
			rstAux.open "select top 1 npedido from pedidos_tienda with (NOLOCK) where npedido like '" & session("ncliente") & "%' " & cadena_mov_sol_usu & " order by fecha desc,npedido desc", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if not rstAux.eof then p_npedido=rstAux("npedido")
			rstAux.close
		end if
		strselect="select m.* from pedidos_tienda as m where npedido like '" & session("ncliente") & "%' and npedido='" & p_npedido & "' " & cadena_mov_sol_usu
		rst.Open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if rst.eof then
			mode="add"
			rst.close
			%><script language="javascript" type="text/javascript">
				parent.botones.document.location="pedido_tiendas_bt.asp?mode=add";
			</script><%
			rst.Open "select * from pedidos_tienda where npedido like '" & session("ncliente") & "%' and npedido='" & p_npedido & "'", _
			session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rst.AddNew
		end if
	elseif mode="add" then
		rst.Open "select * from pedidos_tienda where npedido like '" & session("ncliente") & "%' and npedido='" & p_npedido & "'", _
		session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst.AddNew
    end if

	Alarma "pedido_tiendas.asp"

	if (mode="browse" or mode="edit" or mode="add") then
	  if not rst.EOF then
		if mode="browse"  and mode="edit"  then
			texto_disabled=" disabled "
		else
			texto_disabled=""
		end if

		%><input type="hidden" name="h_npedido" value="<%=EncodeForHtml(null_s(rst("npedido")))%>">
        <div class="headers-wrapper">
            <%
                DrawDiv "header-date", "", ""
                    DrawLabel "", "", LitFecha
                     if texto_disabled & "">"" then
                        if mode = "browse" then
                            DrawSpan "", "", EncodeForHtml(rst("fecha")), ""
                        else
                            %><input disabled type="text" name="h_fecha" value="<%=EncodeForHtml(iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha")))%>" size="10"> <%
                            DrawCalendar "h_fecha"
                        end if
                    else
                        if mode = "browse" then
                            DrawSpan "", "", EncodeForHtml(rst("fecha")), ""
                        else
                            %><input type="text" name="fecha" value="<%=EncodeForHtml(iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha")))%>" size="10"> <%
                            DrawCalendar "fecha"
                        end if                        
                    end if
                CloseDiv
                DrawDiv "header-bill", "", ""
                    DrawLabel "", "", LitPedido
                    DrawSpan "", "", EncodeForHtml(trimCodEmpresa(rst("npedido"))), ""
                CloseDiv
                DrawDiv "header-fact", "", ""
                    if mode = "browse" then
            	        no_mostrar=""
				        cantpendiente=d_sum("cantPendiente","detalles_ped_tienda","npedido='" & rst("npedido") & "' ",session("dsn_cliente"))
				        if rst("nmovimiento") & "">"" or rst("nalbaran") & "">"" then
					        no_mostrar="none"
				        end if      
                        %><label><a class="CELDAREFB" href="javascript:if(window.confirm('<%=LitDesPedConvAMov%>')==true){GenerarMovimiento('<%=enc.EncodeForJavascript(rst("npedido"))%>','<%=enc.EncodeForJavascript(ep)%>')}" OnMouseOver="self.status='<%=LitGenMovPed%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitGenerarMov%></a></label><%
                        strSacSerieGenFact="select nserie,right(nserie,len(nserie)-5) as nserie2 from series with(nolock) where nserie like '" & session("ncliente") & "%' and tipo_documento='MOVIMIENTOS ENTRE ALMACENES'"
				        if sFacSerU & "">"" then
					        strSacSerieGenFact=strSacSerieGenFact & " and nserie in " & sFacSerU
				        end if
				        rstAux.cursorlocation=3
				        rstAux.open strSacSerieGenFact,session("dsn_cliente")
                        DrawSelect "","width:150px;","serieMOV",rstAux,nserie_aux,"nserie","nserie2","",""
                        'DrawSelectHeaderPressLetter "CELDARIGHT","60","",0,"","serieMOV",rstAux,nserie_aux,"nserie","nserie2","",""
		 		        rstAux.close
                    end if
                CloseDiv
                DrawDiv "col-sm-3 col-xxs-6 header-print", "", ""
                    if mode = "browse" then
				        if not rst.eof then
					        defecto=obtener_formato_imp(rst("nserie"),"MOVIMIENTOS ENTRE ALMACENES")
				        end if
				        seleccion = "select b.fichero as fichero, a.descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a, formatos_imp as b where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDOS ENTRE ALMACENES' order by descripcion"
				        rstSelect.Open seleccion, DsnIlion, adOpenKeyset, adLockOptimistic
                        %><label><a id="idPrintFormat" class="CELDAREFB" style="margin: 0px;" href="javascript:AbrirVentana(document.pedido_tiendas.formato_impresion.value+'npedido=<%="(\'"+p_npedido+"\')"%>&mode=browse&empresa=<%=session("ncliente")%>&usuario=<%=session("usuario")%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitImpresionConFormato%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitImpresionConFormato%></a></label>
                        <select id="capa_impresion" class='CELDA' style='width:150px' name="formato_impresion"><%
					    encontrado=0
					    while not rstSelect.eof
						    if defecto=rstSelect("descripcion") then
							    encontrado=1
							if isnull(rstSelect("parametros")) then
								prm=""
							else
								prm=rstSelect("parametros") & "&"
							end if%>
							<option selected="selected" value="<%=EncodeForHtml(null_s(rstSelect("fichero") & "?" & EncodeForHtml(prm)))%>"><%=EncodeForHtml(null_s(rstSelect("descripcion")))%></option>
						<%else
							if isnull(rstSelect("parametros")) then
								prm=""
							else
								prm=rstSelect("parametros") & "&"
							end if%>
							<option value="<%=EncodeForHtml(null_s(rstSelect("fichero")  & "?" & EncodeForHtml(prm)))%>"><%=EncodeForHtml(null_s(rstSelect("descripcion")))%></option>
						<%end if
						rstSelect.movenext
					wend%>
				    </select><%
                    rstSelect.close
                    end if
                CloseDiv
                %>
        </div>
        <table width="100%"></table>    
        
        
             
       <% 'Colapso de todos las secciones %>
       <!--
        <div id="CollapseSection">
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['CABECERA', 'DETALLES', 'AddDG', 'BrowseCab', 'BrowseDG']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['CABECERA', 'DETALLES', 'AddDG', 'BrowseCab', 'BrowseDG']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
        </div>
        -->
		<%'Inicio Borde Span%>
        <table width="100%"><tr><td>

    <%'Datos Generales
    if mode="browse" then %> 
        <div class="Section" id="S_BrowseDG">
            <a href="#" rel="toggle[BrowseDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader">
                <%=litcabecera%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
        </div></a>

            <div class="SectionPanel" style="display:none;" id="BrowseDG">
	         
             <table width="70%" border='<%=borde%>' cellspacing="1" cellpadding="1"><%
                    'Almacen Origen
                    DrawDiv "1","",""
                    DrawLabel "", "", LitAlmacenOrigen
					if rst("almdestino")&"">"" then
						%><span class='CELDA'><%=EncodeForHtml(null_s(d_lookup("descripcion","almacenes","codigo='" & rst("almorigen") & "'",session("dsn_cliente"))))%>
							<input class='CELDA' type="hidden" name="almorigen" value="<%=EncodeForHtml(null_s(rst("almorigen")))%>" size=10 >
						</span><%
					else
						%><span class=dato>&nbsp;</span><%
					end if
                    CloseDiv

                    'Almacen Destino
                    DrawDiv "1","",""
                    DrawLabel "","", LitAlmacenDestino
					if rst("almdestino")&"">"" then
						%><span class='CELDA'><%=EncodeForHtml(null_s(d_lookup("descripcion","almacenes","codigo='" & rst("almdestino") & "'",session("dsn_cliente"))))%>
							<input class='CELDA' type="hidden" name="almdestino" value="<%=EncodeForHtml(null_s(rst("almdestino")))%>" size=10 >
						</span><%
					else
						%><td class=dato>&nbsp;</td><%
					end if
                    CloseDiv

                    'El numero de serie
                    EligeCeldaResponsive "text",mode,"CELDA","","",0,"nserie", LitSerie,"", EncodeForHtml(null_s(trimCodEmpresa(rst("nserie"))))
					%><input type="hidden" name="nserie" value="<%=EncodeForHtml(rst("nserie")&"")%>"><%
			%>
		</table><%
                    DrawDiv "3-sub", "background-color: #eae7e3", ""
                    %> 
                    <label class="ENCABEZADOC" style="text-align:left"><%=LITDATOSGENERALES%></label>
                    <%
                    CloseDiv
				'ponemos el responsable por defecto
				if tmp_responsable="" then
					'buscar el usuario en la tabla personal
					rstAux.cursorlocation=3
					rstAux.open "select dni from personal with (NOLOCK) where login='" & session("usuario") & "' and dni like '" & session("ncliente") & "%'",session("dsn_cliente")
					if not rstAux.eof then
						tmp_responsable=rstAux("dni")
					end if
					rstAux.close
				else
					if mid(tmp_responsable,1,5)<>session("ncliente") then
						tmp_responsable=session("ncliente") & tmp_responsable
					end if
				end if

                    DrawDiv "1","",""
                    DrawLabel "","",LitResponsable
					if rst("responsable")&"">"" then
						%><span class='CELDA'><%=EncodeForHtml(null_s(d_lookup("nombre","personal","dni='" & rst("responsable") & "'",session("dsn_cliente"))))%>
							<input class='CELDA' type="hidden" name="responsable" value="<%=EncodeForHtml(null_s(rst("responsable")))%>" size=10 >
						</span><%
					else
						%><span class=dato>&nbsp;</span><%
					end if
                    CloseDiv
                    EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitPedidoProc,"",EncodeForHtml(rst("su_npedido")&"")

                    EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitMovimiento, "",EncodeForHtml(trimcodempresa(rst("nmovimiento")&""))
                    %><input type="hidden" name="h_nmovimiento" value="<%=EncodeForHtml(trimcodempresa(rst("nmovimiento")&""))%>"><%

                    EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitAlbaran, "", EncodeForHtml(trimcodempresa(rst("nalbaran")&""))
                    %><input type="hidden" name="h_nalbaran" value="<%=EncodeForHtml(trimcodempresa(rst("nalbaran")&""))%>"><%

					%><%
					'dgb: 25/03/2008  anyadir ESTADO_PED
						strselect2="select codigo,descripcion from estado_doc with(nolock) where codigo= '" & rst("estado_ped") & "' order by descripcion"
						rstAux2.cursorlocation=3
						rstAux2.open strselect2,session("dsn_cliente")
						if not rstAux2.EOF then
                            EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitSituacionPedido,"",EncodeForHtml(null_s(rstAux2("descripcion")))
						else
                            EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitSituacionPedido,"",""
						end if
						rstAux2.close
					    EligeCeldaResponsive "text",mode, "CELDA","","",0,"",LitObservaciones,"",pintar_saltos_nuevo(EncodeForHtml(null_s(iif(tmp_observaciones>"",tmp_observaciones,rst("observaciones")&""))))
						
			%>
        </div>
    </div>
    <%else%>
        <div class="Section" id="S_AddDG">
            <a href="#" rel="toggle[AddDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader">
                <%=litcabecera%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
            </div></a>
            <div class="SectionPanel" style="display:" id="AddDG">
             <table width="70%" border='<%=EncodeForHtml(borde)%>' cellspacing="1" cellpadding="1"><%
				if mode="browse" or mode="edit" then
                    'Almacen Origen
                    DrawDiv "1","",""
                    DrawLabel "","", LitAlmacenOrigen
					if rst("almdestino")&"">"" then
                        DrawSpan "CELDA","",EncodeForHtml(null_s(d_lookup("descripcion","almacenes","codigo='" & rst("almorigen") & "'",session("dsn_cliente")))),""%>
					    <input class='CELDA' type="hidden" name="almorigen" value="<%=EncodeForHtml(null_s(rst("almorigen")))%>" size=10 ><%
					else
						DrawSpan "CELDA","","",""
					end if
                    CloseDiv
                    'Almacen Destino
                    DrawDiv "1","",""
                    DrawLabel "","",LitAlmacenDestino
					if rst("almdestino")&"">"" then
                        DrawSpan "CELDA","",EncodeForHtml(null_s(d_lookup("descripcion","almacenes","codigo='" & rst("almdestino") & "'",session("dsn_cliente")))),""%>
							<input class='CELDA' type="hidden" name="almdestino" value="<%=EncodeForHtml(null_s(rst("almdestino")))%>" size=10 ><%
					else
						DrawSpan "CELDA","","",""
					end if
                    CloseDiv

                    'El numero de serie
                    DrawDiv "1","",""
                    DrawLabel "","",LitSerie
                    DrawSpan "CELDA","",EncodeForHtml(null_s(trimCodEmpresa(rst("nserie")))),""%><input type="hidden" name="nserie" value="<%=EncodeForHtml(rst("nserie")&"")%>"><%
                    CloseDiv
				else
                    'Almacen Origen
                    DrawDiv "1","",""
                    DrawLabel "","",LitAlmacenOrigen
					campo="codigo"
					campo2="descripcion"
					dato_celda=Desplegable(mode,campo,campo2,"almacenes",iif(tmp_almorigen>"",tmp_almorigen,rst("almorigen")),"")
					if texto_disabled & "">"" then
                        EligeCeldaResponsive1 "select",mode,"CELDA","","almorigen",EncodeForHtml(dato_celda),""
						%><input type="hidden" name="almorigen" value="<%=EncodeForHtml(null_s(iif(tmp_almorigen>"",tmp_almorigen,rst("almorigen"))))%>"><%
					else
                        EligeCeldaResponsive1 "select",mode,"CELDA","","almorigen",EncodeForHtml(dato_celda),""
					end if
                    
					if mode="add" or mode="edit" then RstAux.close
					%><input type="hidden" name="h_almorigen" value="<%=EncodeForHtml(null_s(iif(tmp_almorigen>"",tmp_almorigen,rst("almorigen"))))%>"><%
                    CloseDiv

                    'Almacen Destino
                    DrawDiv "1","",""
                    DrawLabel "","",LitAlmacenDestino
					campo="codigo"
					campo2="descripcion"
					dato_celda=Desplegable(mode,campo,campo2,"almacenes",iif(tmp_almdestino>"",tmp_almdestino,rst("almdestino")),"")
					if texto_disabled & "">"" then
                        EligeCeldaResponsive1 "select",mode,"CELDA","","almdestino",EncodeForHtml(null_s(dato_celda)),""
						%><input type="hidden" name="almdestino" value="<%=EncodeForHtml(null_s(iif(tmp_almdestino>"",tmp_almdestino,rst("almdestino"))))%>"><%
					else
                        EligeCeldaResponsive1 "select",mode,"CELDA","","almdestino",EncodeForHtml(null_s(dato_celda)),""
					end if
					if mode="add" or mode="edit" then RstAux.close
					%><input type="hidden" name="h_almdestino" value="<%=EncodeForHtml(null_s(iif(tmp_almdestino>"",tmp_almdestino,rst("almdestino"))))%>"><%
                    CloseDiv

                    'Numero de serie
					rstAux.cursorlocation=3
					rstAux.open "select nserie, nombre as descripcion from series with (NOLOCK) where tipo_documento ='PEDIDOS ENTRE ALMACENES' and nserie like '" & session("ncliente") & "%' order by nombre",session("dsn_cliente")
					if mode="add" then
                         DrawSelectCelda "width:200px","200","",0,LitSerie,"nserie",rstAux,"","nserie","descripcion","",""
					else
						if texto_disabled & "">"" then                   
                            DrawSelectCelda "width:200px","200","",0,LitSerie,"nserie",rstAux,"","nserie","descripcion","",""
							%><input type="hidden" name="nserie" value="<%=EncodeForHtml(null_s(iif(p_serie>"",p_serie,rst("nserie"))))%>"><%
						else
                            DrawSelectCelda "width:200px","200","",0,LitSerie,"nserie",rstAux,"","nserie","descripcion","",""
						end if
					end if
			 		rstAux.close
				end if	
                'response.Write (Prueba)
			%>
         </table><%
                    DrawDiv "3-sub", "background-color: #eae7e3", ""
                    %> 
                    <label class="ENCABEZADOC" style="text-align:left"><%=LITDATOSGENERALES%></label>
                    <%
                    CloseDiv

				'ponemos el responsable por defecto
				if tmp_responsable="" then
					'buscar el usuario en la tabla personal
					rstAux.cursorlocation=3
					rstAux.open "select dni from personal with (NOLOCK) where login='" & session("usuario") & "' and dni like '" & session("ncliente") & "%'",session("dsn_cliente")
					if not rstAux.eof then
						tmp_responsable=rstAux("dni")
					end if
					rstAux.close
				else
					if mid(tmp_responsable,1,5)<>session("ncliente") then
						tmp_responsable=session("ncliente") & tmp_responsable
					end if
				end if
                    DrawDiv "1","",""
                    DrawLabel "","",LitResponsable
					if texto_disabled & "">"" then
                        %>
                        <input class='width15' type="text" <%=texto_disabled%> name="h_responsable" value="<%=EncodeForHtml(null_s(iif(tmp_responsable& "">"",trimCodEmpresa(tmp_responsable),trimCodEmpresa(rst("responsable")))))%>">
                        <input type="hidden" name="responsable" value="<%=EncodeForHtml(null_s(iif(tmp_responsable& "">"",trimCodEmpresa(tmp_responsable),trimCodEmpresa(rst("responsable")))))%>">
                        <a class="CELDAREFB" href="javascript:" OnMouseOver="self.status='<%=LitVerPersonal%>'; return true;" OnMouseOut="self.status=''; return true;">
                            <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/>
                        </a>
                        <input class="width40" disabled <%=texto_disabled%> type="text" name="h_nomresponsable" size="30" value="<%=EncodeForHtml(null_s(iif(TmpNombre & "">"",TmpNombre,d_lookup("nombre","personal","dni='" & iif(tmp_responsable & "">"",tmp_responsable,rst("responsable")) & "'",session("dsn_cliente")))))%>">
                        <input type="hidden" name="nomresponsable" value="<%=EncodeForHtml(null_s(iif(TmpNombre & "">"",TmpNombre,d_lookup("nombre","personal","dni='" & iif(tmp_responsable & "">"",tmp_responsable,rst("responsable")) & "'",session("dsn_cliente")))))%>">
                        <%
					else
						%>
                        <input class='width15' type="text" <%=texto_disabled%> name="responsable" value="<%=EncodeForHtml(null_s(iif(tmp_responsable& "">"",trimCodEmpresa(tmp_responsable),trimCodEmpresa(rst("responsable")))))%>" onchange="TraerResponsable();">
                        <a class="CELDAREFB" href="javascript:AbrirVentana('../administracion/personal_buscar.asp?viene=pedido_tiendas&titulo=<%=LitSelPersonal%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPersonal%>'; return true;" OnMouseOut="self.status=''; return true;">
                            <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/>
                        </a>
                        <input class="width40" disabled <%=texto_disabled%> type="text" name="nomresponsable" size="30" value="<%=EncodeForHtml(null_s(iif(TmpNombre & "">"",TmpNombre,d_lookup("nombre","personal","dni='" & iif(tmp_responsable & "">"",tmp_responsable,rst("responsable")) & "'",session("dsn_cliente")))))%>">
                        <%
					end if
                    CloseDiv

                    'Cod. Pedido 
                    DrawDiv "1","",""
                    DrawLabel "","",LitPedidoProc
                    DrawSpan "CELDA","",EncodeForHtml(rst("su_npedido")&""),""
                    CloseDiv
                    

                    'Nº Movimiento
                    DrawDiv "1","",""
                    DrawLabel "","",LitMovimiento
                    DrawSpan "CELDA","",EncodeForHtml(trimcodempresa(rst("nmovimiento")&"")),""
                    %><input type="hidden" name="h_nmovimiento" value="<%=EncodeForHtml(trimcodempresa(rst("nmovimiento")&""))%>"><%
                    CloseDiv
                        
                    'Albarán
                    DrawDiv "1","",""
                    DrawLabel "","",LitAlbaran
                    DrawSpan "CELDA","width:120px",EncodeForHtml(trimcodempresa(rst("nalbaran")&"")),""
                    %><input type="hidden" name="h_nalbaran" value="<%=EncodeForHtml(trimcodempresa(rst("nalbaran")&""))%>"><%
                    CloseDiv

				    'dgb: 25/03/2008  anyadir ESTADO_PED
				    if mode="edit" then
                        DrawDiv "1","",""
                        DrawLabel "","",LitSituacionPedido
						campo="codigo"
				        campo2="descripcion"
				        dato_celda=Desplegable(mode,campo,campo2,"estado_doc",EncodeForHtml(null_s(rst("estado_ped"))),"codigo like '" & session("ncliente") & "%'")
                        EligeCeldaResponsive1 "select",mode,"width60","","estado_ped",EncodeForHtml(null_s(dato_celda)), LitSituacionPedido
                        CloseDiv
					 else 
					    if mode="add" then
						    testo_est=" where codigo like '" & session("ncliente") & "%'"
						    rstAux2.cursorlocation=3
				            rstAux2.open "select codigo,descripcion from estado_doc with(nolock) " & testo_est & " order by descripcion",session("dsn_cliente")
                            DrawSelectCelda "width60","","",0,LitSituacionPedido,"estado_ped",rstAux2,"","codigo","descripcion","",""
				            rstAux2.close
					    end if
					 end if

                     'Observaciones
					 if texto_disabled & "">"" then
                        EligeCelda "text",mode,"","","",60,LitObservaciones,"observaciones",2, pintar_saltos_nuevo(EncodeForHtml(null_s(rst("observaciones"))))
					 	%><input type="hidden" name="observaciones" value="<%=pintar_saltos_nuevo(EncodeForHtml(null_s(iif(tmp_observaciones>"",tmp_observaciones,iif(rst("observaciones")>"",rst("observaciones"),"")))))%>"><%
					 else
                        EligeCelda "text",mode,"","","",60,LitObservaciones,"observaciones",2,pintar_saltos_nuevo(EncodeForHtml(null_s(rst("observaciones"))))
					 end if
			%>
        </div>
    </div>
    <% end if %>

		<%if mode="browse" then
			'Detalles %>
            <div   class="Section" id="S_DETALLES">
                <a href="#" rel="toggle[DETALLES]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitTituloDet%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div></a>

                <div class="SectionPanel" style="display: " id="DETALLES">
	             <br />
                    <div class="overflowXauto">
                 <table class="width90 md-table-responsive bCollapse" ><%
							%>
                            <tr>
                            <td class="ENCABEZADOL underOrange width15"><%=LitItem%></td>
							<td class="ENCABEZADOL underOrange width15"><%=LitCantidad%></td>
							<td class="ENCABEZADOL underOrange width15"><%=LitPSS%></td>
							<td class="ENCABEZADOL underOrange width15"><%=LitReferencia%></td>
							<td class="ENCABEZADOL underOrange width10"><%=LitDescripcion%></td>
                            <td class="ENCABEZADOL underOrange width5"></td>
                            </tr><%
					%></table><%
				if rst("nmovimiento")&""="" then
					%><iframe id='frDetallesIns' name="fr_DetallesIns" class='width90 iframe-input md-table-responsive' src='pedido_tiendas_detins.asp?ndoc=<%=enc.EncodeForJavascript(rst("npedido"))%>' style="height:40px;" frameborder="no" scrolling="no" noresize="noresize"></iframe><br /><%
				end if
				if si_tiene_modulo_produccion<>0 then%>
					<iframe id='frDetalles' name="fr_Detalles" class="width90 iframe-data md-table-responsive" src='pedido_tiendas_det.asp?ndoc=<%=enc.EncodeForJavascript(rst("npedido"))%>&almacen=<%=enc.EncodeForJavascript(rst("almorigen"))%>'  height='150' frameborder="yes" noresize="noresize"></iframe>
				<%else%>
					<iframe id='frDetalles' name="fr_Detalles" class="width90 iframe-data md-table-responsive" src='pedido_tiendas_det.asp?ndoc=<%=enc.EncodeForJavascript(rst("npedido"))%>&almacen=<%=enc.EncodeForJavascript(rst("almorigen"))%>'  height='150' frameborder="yes" noresize="noresize"></iframe>
				<%end if%>
                </div></div>   
            </div>
		<%end if%>

		</td></tr></table>

		<span id="paginacion" style="display: ">
		</span>
	    <%if submode2="traerresponsable" then
			if texto_disabled & "">"" and texto_bmr<>" " & "disabled" then%>
				<script language="javascript" type="text/javascript">
					document.pedido_tiendas.merric.focus();
				</script>
			<%else%>
				<script language="javascript" type="text/javascript">
					document.pedido_tiendas.observaciones.focus();
					document.pedido_tiendas.observaciones.select();
				</script>
			<%end if
		elseif mode="add" then%>
			<script language="javascript" type="text/javascript">
				document.pedido_tiendas.fecha.focus();
				document.pedido_tiendas.fecha.select();
			</script>
		<%elseif mode="edit" then
			if texto_disabled & "">"" and texto_bmr<>" " & "disabled" then%>
				<script language="javascript" type="text/javascript">
					document.pedido_tiendas.merric.focus();
				</script>
			<%else%>
				<script language="javascript" type="text/javascript">
					document.pedido_tiendas.fecha.focus();
				</script>
			<%end if
		end if
        if mode="browse" then %>
		    <script language="javascript" type="text/javascript">Redimensionar();</script>
        <%end if
	  end if
	end if %>
</form>
<%
	set rstAux = nothing
	set rstAux2 = nothing
	set rstAux3 = nothing
	set rst = nothing
	set rstSelect = nothing
    set conn = nothing
end if%>
</body>
</html>
