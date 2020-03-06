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
' JCI 21/05/2003 : Hay que marcar/desmarcar la factura como pagada en función de los pagos de vencimientos
'                  Pongo lo de la caché
''ricardo 5-6-2003 se pone el parametro caju para saber a que cajas tiene acceso el usuario
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

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
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="pagos.inc" -->

<!--#include file="../perso.inc" -->
<!--#include file="../common/poner_cajaResponsive.inc" -->

<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript" >
/**FLM:20090529: Suma el total de los documentos marcados.**/
var totalImportePagar=0.00;
var numDecimalesEmpresa='<%=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and Moneda_base<>0",session("dsn_cliente")) %>';

function sumTotREM(fila){
    if(document.pagos.elements["check"+fila].checked==true)
        totalImportePagar+=parseFloat(Redondear(parseFloat(document.pagos.elements["imp"+fila].value.replace(",","."))*parseFloat(document.pagos.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
    else
        totalImportePagar-=parseFloat(Redondear(parseFloat(document.pagos.elements["imp"+fila].value.replace(",","."))*parseFloat(document.pagos.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
    document.getElementById("totalAPagar").innerHTML=totalImportePagar.toFixed(numDecimalesEmpresa);
}

/*********************************************/
function seleccionar() {
	nregistros=document.pagos.h_nfilas.value;
	if (document.pagos.check.checked){
		for (i=1;i<=nregistros;i++) {
		    nombre="check" + i;
		    //FLM:20090529:Solo actualizo si es necesario
		    if(document.pagos.elements[nombre].checked==false){
			    document.pagos.elements[nombre].checked=true;
			    //FLM:20090529: Suma el total de los documentos marcados
	            sumTotREM(i);
	        }
		}
		document.pagos.check.value="yyy"
	}
	else{
		for (i=1;i<=nregistros;i++) {
		    nombre="check" + i;
		    //FLM:20090529:Solo actualizo si es necesario
		    if(document.pagos.elements[nombre].checked==true){
			    document.pagos.elements[nombre].checked=false;
			    //FLM:20090529: Suma el total de los documentos marcados
	            sumTotREM(i);
	        }
		}
		//FLM:20090529:el importe tiene que ser 0 al desmarcar el check.
		totalImportePagar=0.00;
		 document.getElementById("totalAPagar").innerHTML=totalImportePagar.toFixed(numDecimalesEmpresa);
		document.pagos.check.value="xxx"
	}
}

//Desencadena la búsqueda del proveedor cuya referencia se indica
function TraerProveedor(mode) {
	if (document.pagos.tabla[0].checked) tabla = "facturas_pro";
	else tabla = "vencimientos_entrada";

	document.location.href="pagos.asp?nproveedor=" + document.pagos.nproveedor.value + "&mode=" + mode
		+"&serie="+document.pagos.serie.value+"&fdesde="+document.pagos.fdesde.value
		+"&fhasta="+document.pagos.fhasta.value+"&tabla="+tabla
		+"&caju=" + document.pagos.caju.value;
}
</script>
<body onload="self.status='';" class="BODY_ASP">
<%
'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
	'fdesde:
	'fhasta:
	'nproveedor:
	'serie
function CadenaBusqueda(fdesde,fhasta,nproveedor,serie)
	CadenaBusqueda = ""
	if fac = "true" then
		if nproveedor > "" then
			CadenaBusqueda = " nproveedor='" & nproveedor & "' and"
		end if
		if serie > "" then
			CadenaBusqueda = CadenaBusqueda + " serie='" & serie & "' and"
		end if
		CadenaBusqueda = CadenaBusqueda + " fecha>='" & fdesde & "' and fecha<='" & fhasta & "' and pagada = 0 and nfactura like '" & session("ncliente") & "%' order by nfactura, fecha desc,nproveedor"

	else
		CadenaBusqueda = " vencimientos_entrada.nfactura=facturas_pro.nfactura and"
		if nproveedor > "" then
			CadenaBusqueda = CadenaBusqueda + " facturas_pro.nproveedor='" & nproveedor & "' and"
		end if
		if serie > "" then
			CadenaBusqueda = CadenaBusqueda + " facturas_pro.serie='" & serie & "' and"
		end if
		CadenaBusqueda = CadenaBusqueda + " vencimientos_entrada.fecha>='" & fdesde & "' and vencimientos_entrada.fecha<='" & fhasta & "' and vencimientos_entrada.pagado = 0 and facturas_pro.nfactura like '" & session("ncliente") & "%'  order by vencimientos_entrada.fecha, vencimientos_entrada.nfactura, vencimientos_entrada.nvencimiento"
	end if
end function

'*****************************************************************************
'Anota en la caja caj con el medio pag el importe impcaja. Tipo "F"->factura. Tipo "V"->vencimiento
sub AnotarEnCaja(tipo,ndoc,ndoc_pro,impcaja,div,caj,pag,rsoc,fechapago, tienda)
	MB=d_lookup("codigo","divisas","moneda_base<>0",session("dsn_cliente"))
	SigAnotacion=d_max("nanotacion","caja","caja='" & caj & "'",session("dsn_cliente")) + 1
	rstAux.open "select * from caja where caja=''",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	rstAux.addnew
	rstAux("caja")=caj
	rstAux("nanotacion")=SigAnotacion
	rstAux("tanotacion")=iif(impcaja>=0,"SALIDA","ENTRADA")
	rstAux("fecha")=fechapago
	rstAux("importe")=iif(impcaja>=0,impcaja,-impcaja)
	rstAux("medio")=pag
	rstAux("descripcion")=rsoc & " (Desde documento)"
	rstAux("ndocumento")=ndoc
	if ndoc_pro>"" then rstAux("ndocumento_pro")=ndoc_pro
	rstAux("tdocumento")=iif(tipo="F","FACTURA DE PROVEEDOR","VENCIMIENTO_ENTRADA")
	rstAux("divisa")=div
    if tipo = "F" then
        factcambio=d_lookup("factcambio","facturas_pro","nfactura like '"& session("ncliente") &"%' and nfactura='" & ndoc & "'",session("dsn_cliente"))
    elseif tipo = "V" then
        pos = instr(1,ndoc,"-")
        factcambio=d_lookup("factcambio","facturas_pro","nfactura like '"& session("ncliente") &"%' and nfactura='" & mid(ndoc,1,pos-1) & "'",session("dsn_cliente"))
    else
        factcambio = 1
    end if
    rstAux("change_currency")=null_z(factcambio)
	rstAux("tapunte")=session("ncliente") & "01"
	rstAux("tienda")=nulear(tienda)
	rstAux.update
	rstAux.close
end sub

'******************************************************************************
'Convierte los pedidos en facturas
sub pagar_facturas()
	strwhere ="("
	while h_nfilas>0
		x=request.form("check" & h_nfilas)
		if x>"" then
			strwhere = strwhere & "'" & x & "',"
		end if
		h_nfilas = h_nfilas -1
	wend

	cadena = "nfactura in " & strwhere & "'xxxxxx') order by nfactura"
	rst.Open "select nfactura,pagada,deuda,razon_social,facturas_pro.nfactura_pro,facturas_pro.divisa, facturas_pro.tienda from facturas_pro with(nolock),proveedores pro with(nolock)where facturas_pro.nproveedor=pro.nproveedor and " & cadena, session("dsn_cliente"),adOpenKeyset,adLockOptimistic

		while not rst.EOF
			pendiente=rst("deuda")

	''10/1/2003 - puesto por ricardo para cuando no habiendo detalles o conceptos, se borran todos, que tambien se
	'borren los vencimientos, si hay vencimiento automatico
	res_borr_dv=comprob_deuda_venci(rst("nfactura"),"COMPRAS")

			rst.Update
			
			'FLM:20090505:Comprobamos que no haya ninguna remesa con el vencimiento
            rstAux.open "select top 1 r.nremesa from remesas_pro r with(nolock) inner join detalles_rempro dr with(nolock) on dr.nremesa=r.nremesa and (dr.nfacturavto='" & rst("nfactura") & "' ) where r.nempresa='" & session("ncliente") & "' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            if rstAux.EOF then
                venConRemesa=0
            else
                venConRemesa=1
            end if
            rstAux.close

            'FLM:20090507: solo se incluye en caja la factura si no tiene ningún vencimiento remesado
			if caja>"" and venConRemesa=0 then
				AnotarEnCaja "F",rst("nfactura"),rst("nfactura_pro"),pendiente,rst("divisa"),caja,pago,rst("razon_social"),fechapago, rst("tienda")
			end if
			'mmg: solo se borran los vencimientos si se ha seleccionado caja
			if caja&"" <> "" and  venConRemesa=0 then
			    rstAux.open "delete from vencimientos_entrada with(rowlock) where pagado=0 and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			else
			    rstAux.Open "update vencimientos_entrada with(updlock) set pagado = 1 where pagado=0 and nfactura='" & rst("nfactura") & "'" , session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			end if
			
            rstAux.open "update facturas_pro with(updlock) set pagada=1 where nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rst.MoveNext
		wend
	rst.Close
end sub

'******************************************************************************
'Convierte los pedidos en facturas
sub pagar_vencimientos()
	restricciones=0

	strwhere ="("
	while h_nfilas>0
		x=limpiaCadena(request.form("check" & h_nfilas))
		if x>"" then
			strwhere = strwhere & "'" & x & "',"
		end if
		h_nfilas = h_nfilas -1
	wend

	cadena = "nfactura+'-'+cast(nvencimiento as varchar(10)) in " & strwhere & "'xxxxxx') order by nfactura"
	rst.Open "select * from vencimientos_entrada where " & cadena, session("dsn_cliente"),adOpenKeyset,adLockOptimistic

	while not rst.EOF
        ' Comprobar que se cumplen las restricciones: 0<=deuda-recibido<=Total
        rstSelect.Open "select * from vencimientos_entrada with(nolock) where nfactura='" & rst("nfactura") & "' and nvencimiento=" & rst("nvencimiento") & " order by nvencimiento", _
        session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        deuda=d_lookup("deuda","facturas_pro","nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
        totalFactura=null_z(d_lookup("total_factura","facturas_pro","nfactura='" & rst("nfactura") & "'",session("dsn_cliente")))
        importeACobrar=rstSelect("importe")
        rstSelect.close
        deudaResultante=deuda-importeACobrar
        ''ricardo 8-3-2010 se cambia la siguiente condicion para que se puedan pagar por caja los vencimientos con importe negativo
        if deudaResultante>=0 and deudaResultante<=abs(totalFactura) then

        pendiente=rst("importe")
        rst.update

        rstSelect.open "select facturas_pro.*,proveedores.razon_social from facturas_pro with(nolock),proveedores with(nolock) where nfactura='" & rst("nfactura") & "' and facturas_pro.nproveedor=proveedores.nproveedor ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

        if not rstSelect.eof then
            div=rstSelect("divisa")
            rsoc=rstSelect("razon_social")
            fac_pro=rstSelect("nfactura_pro")
            tienda=rstSelect("tienda")
        end if
        rstSelect.close
        deuda=d_lookup("deuda","facturas_pro","nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
        rstSelect.open "select count(*) as total from vencimientos_entrada with(nolock) where pagado=0 and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        SinCobrar=rstSelect("total")
        rstSelect.close
        if caja>"" then AnotarEnCaja "V",rst("nfactura") & "-" & rst("nvencimiento"),fac_pro & "-" & rst("nvencimiento"),pendiente,div,caja,pago,rsoc,fechapago, tienda
            rstSelect.open "update vencimientos_entrada with(rowlock) set pagado=1 where nfactura='" & rst("nfactura") & "' and nvencimiento='" & rst("nvencimiento") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        else
            restricciones=1
        end if
        deuda=d_lookup("deuda","facturas_pro","nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
        if deuda=0 then
            rstSelect.open "update facturas_pro with(rowlock) set pagada=1 where nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            rstSelect.open "delete from vencimientos_entrada with(rowlock) where pagado=0 and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        end if
        rst.MoveNext
	wend
	rst.Close

	if restricciones=1 then
		%><script language="javascript" type="text/javascript">
		    window.alert("<%=LitNoPueOperacionDeudaIncorr%>");
		    parent.document.pagos.action="pagos.asp?mode=select1";
		    parent.document.pagos.submit();
		</script><%
	end if
end sub

'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
const borde=0

 %>
<form name="pagos" method="post">
<% PintarCabecera "pagos.asp"
'Leer parámetros de la página
	mode		= EncodeForHtml(Request.QueryString("mode"))

	caju	= limpiaCadena(Request.QueryString("caju"))
	if caju="" then
		caju	= limpiaCadena(Request.form("caju"))
	end if

	%><input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>"><%

	nproveedor	= limpiaCadena(Request.QueryString("nproveedor"))
	if nproveedor ="" then
		nproveedor	= limpiaCadena(Request.form("nproveedor"))
	end if
	if nproveedor > "" then
		nproveedor = completar(nproveedor,5,"0")
		nproveedor=session("ncliente") & nproveedor
	end if
	fdesde		= limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde	= limpiaCadena(Request.form("fdesde"))
	end if
	fhasta		= limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta	= limpiaCadena(Request.form("fhasta"))
	end if
	serie		= limpiaCadena(Request.QueryString("serie"))
	if serie ="" then
		serie	= limpiaCadena(Request.form("serie"))
	end if
	h_nfilas	= limpiaCadena(Request.QueryString("h_nfilas"))
	if h_nfilas ="" then
		h_nfilas	= limpiaCadena(Request.form("h_nfilas"))
	end if

	tabla	= limpiaCadena(Request.QueryString("tabla"))
	if tabla ="" then
		tabla	= limpiaCadena(Request.form("tabla"))
	end if

	if tabla="facturas_pro" then
		fac="true"
	else
		fac="false"
	end if

	if request.form("opcproveedorbaja")>"" then
		opcproveedorbaja=limpiaCadena(request.form("opcproveedorbaja"))
	else
		opcproveedorbaja=limpiaCadena(request.querystring("opcproveedorbaja"))
	end if

	caja=limpiaCadena(request.form("ncaja"))
	pago=limpiaCadena(request.form("i_pago"))
	fechapago=limpiaCadena(request.form("fechapago"))

	strwhere=""
%>
	<!--<table width='100%'>
   	<tr>
	  <td width="50%" align="center" bgcolor="<%=color_fondo%>"><font class='CABECERA'><b><%=LitTitulo%></b></font></td>
	  <td><font class='CABECERA'><b></b></font>
 	      <font class=CELDA><b></b></font>
	  </td>
   	</tr>
    </table>-->
	<% Alarma "pagos.asp" %>
	<hr/>
<%
	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")

	'Acción a realizar
	if mode="confirm"then
		if fac = "true" then
			pagar_facturas
		else
			pagar_vencimientos
		end if
		mode = "select1"
	end if

	if mode="select1"then

				if tabla <> "facturas_pro" then
                  DrawDiv "1", "", ""
                    DrawLabel "", "", LitFacturas
					%><input type="radio" name="tabla" value="facturas_pro"><%
                  CloseDiv
                  DrawDiv "1", "", ""
                    DrawLabel "", "", LitVencimientos
					%><input type="radio" name="tabla" value="vencimientos_entrada" checked><%
                  CloseDiv
				else
                    DrawDiv "1", "", ""
    					DrawLabel "", "", LitFacturas
					%><input type="radio" name="tabla" value="facturas_pro" checked><%
                    CloseDiv
                    DrawDiv "1", "", ""
                        DrawLabel "", "", LitVencimientos
					%><input type="radio" name="tabla" value="vencimientos_entrada" ><%
                    CloseDiv
				end if
			%>
		<hr/>
        <%
			if fdesde >"" then
			 	EligeCelda "input", "add", "left", "", "", 0, LitDesdeFecha, "fdesde", 10, EncodeForHtml(fdesde)
                DrawCalendar "fdesde"
			else
                EligeCelda "input", "add", "left", "", "", 0, LitDesdeFecha, "fdesde", 10, EncodeForHtml("01/01/"+cstr(year(date)))
                DrawCalendar "fdesde"
			end if
			if fhasta >"" then
                EligeCelda "input", "add", "left", "", "", 0, LitHastaFecha, "fhasta", 10, EncodeForHtml(fhasta)
                DrawCalendar "fhasta"
			else
                EligeCelda "input", "add", "left", "", "", 0, LitHastaFecha, "fhasta", 10, EncodeForHtml(date)
                DrawCalendar "fhasta"    
			end if

            DrawDiv "1", "", ""
                DrawLabel "", "", LitProveedor
			%><input class='width15' type="text" name="nproveedor" value="<%=EncodeForHtml(trimCodEmpresa(nproveedor))%>" size = 10 onchange="TraerProveedor('<%=enc.EncodeForJavascript(mode)%>','<%=EncodeForHtml(ndet)%>');">
			<a class='CELDAREFB' href="javascript:AbrirVentana('../compras/proveedores_busqueda.asp?ndoc=pagos&titulo=S<%=LitSelProveedor%>&mode=search&viene=pagos','P',<%=AltoVentana%>,<%=AnchoVentana%>)"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt=""></a>
            <%if nproveedor >"" then
				DrawInput "width40","","razon_social", EncodeForHtml(null_s(d_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente")))), ""
			else
				DrawInput "width40","","razon_social","", ""
			end if
		    CloseDiv
            EligeCelda "check", "add", "left", "", "", 0, LitProveedorBaja, "opcproveedorbaja", 0, ""

			rstSelect.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series where tipo_documento ='FACTURA DE PROVEEDOR' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "CELDA","150","",0,LitSerieFactura,"serie",rstSelect,serie,"nserie","descripcion","",""

    elseif mode="select2" then%>
		<input type="hidden" name="tabla" value="<%=EncodeForHtml(tabla)%>"><%
		if fac="true" then%>
			<input type="hidden" name="mensaje" value="<%=LitMsgPagarFacConfirm%>"><%
		else%>
			<input type="hidden" name="mensaje" value="<%=LitMsgPagarVenConfirm%>"><%
		end if

		strwhere = CadenaBusqueda(fdesde,fhasta,nproveedor,serie)

	if nproveedor="" then
		if opcproveedorbaja="" then
			strbaja=" "
		else
			strbaja=" and pro.fbaja is null"
			h=instr(1,strwhere," order")
			if h>0 then
				strwhere=mid(strwhere,1,h) & strbaja & mid(strwhere,h,len(strwhere))
			end if
		end if
	end if

	rst.cursorlocation=3
	if fac="true" then
		if tabla="facturas_pro" then
		    'FLM:20090511:Filtro que no se listen facturas cuyos vencimientos estén incluidos en alguna remesa.
			tabla="facturas_pro,proveedores as pro,divisas d "
			strwhere="facturas_pro.nproveedor=pro.nproveedor and facturas_pro.divisa=d.codigo and "&_
			    " not exists(select nfacturavto from remesas_pro r with(nolock) inner join detalles_rempro dr with(nolock) on dr.nfacturavto=facturas_pro.nfactura and dr.nremesa=r.nremesa where r.nempresa = '"&session("ncliente")&"') and "&_
			    strwhere
			
			j=instr(1,strwhere," nproveedor")
			if j>0 then
				strwhere=mid(strwhere,1,j) & "facturas_pro." & mid(strwhere,j+1,len(strwhere))
			end if
			j=instr(1,strwhere," order by nfactura, fecha desc,nproveedor")
			if j>0 then
				strwhere=mid(strwhere,1,j) & " order by fecha,nfactura,facturas_pro.nproveedor"
			end if
		end if
		rst.Open "select facturas_pro.*,pro.razon_social,d.ndecimales,d.abreviatura,(select count(*) from vencimientos_entrada where vencimientos_entrada.nfactura=facturas_pro.nfactura) as NumeroVencimientos,d.factcambio from " & tabla & " where " & strwhere, session("dsn_cliente")
	else
		if tabla="vencimientos_entrada" then
			tabla=",proveedores as pro,divisas d "
			strwhere="facturas_pro.nproveedor=pro.nproveedor and facturas_pro.divisa=d.codigo and " & strwhere
		end if
		rst.Open "select vencimientos_entrada.*, facturas_pro.nproveedor,facturas_pro.divisa, facturas_pro.serie, facturas_pro.nfactura_pro,pro.razon_social,d.ndecimales,d.abreviatura,d.factcambio from vencimientos_entrada, facturas_pro" & tabla & " where " & strwhere, session("dsn_cliente")
	end if%>
	<table border='0' cellspacing="1" cellpadding="1" width="100%">
	    <tr><td>
	    
	    <%
			defecto=" " 'se pone con espacio en blanco, para que no salga ninguna
			poner_cajasResponsive "width60",defecto, LitCaja, "ncaja", "175", "codigo", "descripcion", "", "",EncodeForHtml(poner_comillas(caju))

			rstAux.Open "SELECT * FROM Tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
			DrawSelectCelda "CELDA7","175","","0", LitPago,"i_pago",rstAux,"","codigo","Descripcion","",""
			rstAux.Close
			
            EligeCelda "input", "add", "left", "", "", 0, LitFechaPago, "fechapago", 10, EncodeForHtml(date)
		%>
    
	  </td>
      <td align="right">
	    <%'FLM:20090529:Mostramos el total del importe seleccionado en la remesa.%>
        <table width="100%">
        <tr>
            <td class="CELDARIGHT"><strong><%=LitImpTotPagar%>:&nbsp;</strong><span id="totalAPagar">0.00</span>&nbsp;<%=EncodeForHtml(null_s(d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and Moneda_base<>0",session("dsn_cliente")))) %></td>
        </tr>
        </table>
      <%'''''''''''''''''' %>    
	  </td></tr></table>
	<hr/>
	<table width='100%' border='0' cellspacing="1" cellpadding="1"><%
		'Fila de encabezado
		DrawFila color_fondo
			if fac="true" then
				%><td class=CELDA>
					<input type="checkbox" name="check" value="true" onclick="seleccionar();">
				</td><%
				DrawCelda "ENCABEZADOL","","",0,LitFactura
				DrawCelda "ENCABEZADOR","","",0,LitFecha
				DrawCelda "ENCABEZADOL","","",0,LitProveedor
				DrawCelda "ENCABEZADOR","","",0,LitTotal
				DrawCelda "ENCABEZADOR","","",0,LitDeuda
				DrawCelda "ENCABEZADOL","","",0,LitDivisa
				DrawCelda "ENCABEZADOL","","",0,LitVencimientos
			else
				%><td class=CELDA>
					<input type="checkbox" name="check" value="true" onclick="seleccionar();">
				</td><%
				DrawCelda "ENCABEZADOL","","",0,LitFactura
				DrawCelda "ENCABEZADO","","",0,LitVencimiento
				DrawCelda "ENCABEZADOR","","",0,LitFecha
				DrawCelda "ENCABEZADOL","","",0,LitProveedor
				DrawCelda "ENCABEZADOR","","",0,LitImporte
				DrawCelda "ENCABEZADOL","","",0,LitDivisa
			end if
		CloseFila

	VinculosPagina(MostrarFacturasPro)=1:VinculosPagina(MostrarProveedores)=1
	CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

		fila=1
		while not rst.EOF
			CheckCadena rst("nfactura")
			'Seleccionar el color de la fila.
			if ((fila+1) mod 2)=0 then
				color=color_blau
				con_negrita=false
			else
				color=color_terra
				con_negrita=false
			end if

			DrawFila color
				if fac="true" then
					%><td class=CELDA>
						<input type="checkbox" name="check<%=fila%>" value="<%=EncodeForHtml(rst("nfactura"))%>" onclick="sumTotREM('<%=fila%>');document.pagos.check.value='yyy';">
					</td><%
					%><td class="CELDALEFT" align="left">
						<%=Hiperv(OBJFacturasPro,rst("nfactura"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(null_s(rst("nfactura_pro"))),LitVerFactura)%>
					</td><%
					'DrawCelda "CELDA","","",0,rst("nfactura_pro")
					DrawCelda "CELDARIGHT","","",0,EncodeForHtml(rst("fecha"))
					'DrawCelda "CELDA","","",0,rst("razon_social")
					DrawCelda "CELDA","","",0,Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(null_s(rst("razon_social"))),LitVerProveedor)
					DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("total_factura")),rst("ndecimales"),-1,0,-1))
					DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))
					DrawCelda "CELDALEFT","","",0, EncodeForHtml(null_s(rst("abreviatura")))
					DrawCelda "CELDALEFT","","",0,iif(rst("NumeroVencimientos")>0,LitSi,LitNo)
					%><input type="hidden" name="imp<%=fila%>" value="<%=EncodeForHtml(null_s(rst("deuda")))%>" /><%
				else
					%><td class=CELDA>
						<input type="checkbox" name="check<%=fila%>" value="<%=EncodeForHtml(rst("nfactura") & "-" & cstr(rst("nvencimiento")))%>" onclick="sumTotREM('<%=fila%>');document.pagos.check.value='yyy';">
					</td><%
					%><td class="CELDALEFT" align="left">
						<%=Hiperv(OBJFacturasPro,rst("nfactura"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(null_s(rst("nfactura_pro"))),LitVerFactura)%>
					</td><%
					DrawCelda "CELDARIGHT","","",0,EncodeForHtml(rst("nvencimiento"))
					DrawCelda "CELDARIGHT","","",0,EncodeForHtml(rst("fecha"))
					'DrawCelda "CELDA","","",0,rst("razon_social")
					DrawCelda "CELDA","","",0,Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(null_s(rst("razon_social"))),LitVerProveedor)
					DrawCelda "CELDARIGHT ","","",0,EncodeForHtml(formatnumber(rst("importe"),rst("ndecimales"),-1,0,-1))
					%><input type="hidden" name="imp<%=fila%>" value="<%=EncodeForHtml(null_s(rst("importe")))%>" /><%
					DrawCelda "CELDALEFT","","",0, EncodeForHtml(rst("abreviatura"))
				end if
			    'FLM:20090602:campo que guarda el factor de conversión de las divisas.
			    %><input type="hidden" name="factcambio<%=fila%>" value="<%=EncodeForHtml(null_s(rst("factcambio")))%>" /><%
			CloseFila
			fila=fila+1
			rst.MoveNext
		wend%>
	<input type="hidden" name="h_nfilas" value="<%=rst.recordcount%>">
        <%rst.Close%>
</table>
<%end if%>
</form>
<hr/>
<%end if
set rstSelect=nothing
set rstAux=nothing
set rst=nothing
%>
</body>
</html>