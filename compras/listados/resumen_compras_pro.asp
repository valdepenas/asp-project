<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
    <title><%=LitTituloResCompra%></title>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")

function EncodeForHtml(data)
	if data & "" <> "" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function

Server.ScriptTimeout = 400 
%>
    <meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
    <META HTTP-EQUIV="Content-style-TypeCONTENT="text/css">
    <link rel="stylesheet" href="../../pantalla.css" media="screen">
    <link rel="stylesheet" href="../../impresora.css" media="print">
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<% if accesoPagina(session.sessionid,session("usuario"))=1 then %>
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../modulos.inc" -->

<!--#include file="../facturas_pro.inc" -->
<!--#include file="../compras.inc" -->

<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../../styles/formularios.css.inc" -->
<!--#include file="../../js/calendar.inc" -->    

<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
function BloquearVentasNetas()
{
    if(document.resumen_compras_pro.opc_comprasnetas.checked) {
        document.resumen_compras_pro.opc_cantidad.checked=false;
    }
    else {
        document.resumen_compras_pro.opc_cantidad.checked=true;
    }
}

function loadSeries(serie)
{
    jQuery.ajax({
        url: "lista_facturacion_pro_series.asp?id=" + serie,
        type: "POST",
        success: function (response) {
            jQuery('#selectSeries').html("").fadeIn();
            jQuery('#selectSeries').html(response).fadeIn();
        },
        error: function (data, textStatus) {
            console.log('Error loadSeriesByStore');
        }
    });
}

function keypress2(){
	tecla=window.event.keyCode;
	//keyPressed(tecla);
}

function keyPressed(tecla) {
}

function optionalFields() {
	if (document.resumen_compras_pro.agrupar.value=="<%=LitMeses%>"){
		<%if si_tiene_modulo_proyectos<>0 then%>
			agrOtros.style.display = "none";
		<%end if%>
		agrMeses.style.display = "";
		agrMeses2.style.display = "";
		agrMeses3.style.display = "";
	}
	else {
		<%if si_tiene_modulo_proyectos<>0 then%>
			agrOtros.style.display = "";
		<%end if%>
		agrMeses.style.display = "none";
		agrMeses2.style.display = "none";
		agrMeses3.style.display = "none";
	}
	if (document.resumen_compras_pro.agrupar.value=="<%=ucase(LitProveedor)%>")
	{
		agrProveedor.style.display = "";
	}
	else 
	{
		agrProveedor.style.display = "none";
	}	
	//ricardo 22/1/2003 para que se vea o no lo de clientes en hojas separadas, ya que solo sirve para la agrupacion por cliente
	if (document.resumen_compras_pro.agrupar.value=="<%=ucase(LitProveedor)%>" || (document.resumen_compras_pro.agrupar.value=="<%=LitMeses%>" && document.resumen_compras_pro.mostrarfilas.value=="<%=LitProveedores%>")){
		agrProHojSep.style.display="";
		agrProHojSep2.style.display="none";
	}
	else{
		agrProHojSep.style.display="none";
		agrProHojSep2.style.display="";
	}
}

//Desencadena la búsqueda del proveedor cuya referencia se indica
function TraerProveedor(mode) {
	document.resumen_compras_pro,acion="resumen_compras_pro.asp?nproveedor=" + document.resumen_compras_pro.nproveedor.value +
	"&mode=" + mode;
	document.resumen_compras_pro.submit();

}

function Ver_Conceptos(){
	if(document.resumen_compras_pro.ver_conceptos.checked) {
		document.resumen_compras_pro.familia.value='';
		document.resumen_compras_pro.familia.disabled=true;
		document.resumen_compras_pro.referencia.value='';
		document.resumen_compras_pro.referencia.disabled=true;
		document.resumen_compras_pro.tipo_articulo.value='';
		document.resumen_compras_pro.tipo_articulo.disabled=true;
	}else {
		document.resumen_compras_pro.familia.disabled=false;
		document.resumen_compras_pro.referencia.disabled=false;
		document.resumen_compras_pro.tipo_articulo.disabled=false;
	}
}
</script>
<body onload="self.status='';" class="BODY_ASP">
<%
'RGU 17/11/2007 CAMBIO DSN PARA LISTADOS
''ricardo 22/1/2003
''para que cuando una cantidad sea 0 , que se ponga valor en blanco
'para mostrar bien las cantidades en la agrupacion por meses
sub mostrar_Cant_Meses(clase,opcion,valor1,valor2,ndec)%>
	<td class="<%=EncodeForHtml(clase)%>" style="border: 1px solid Black;">                                      	
		<%if opcion>"" then
			if valor1<>0 then%>
				<%=EncodeForHtml(formatnumber(valor1,ndec,-1,0,-1))%>
			<%else%>
				<%="&nbsp;"%>
			<%end if
		else
			if valor2<>0 then%>
				<%=EncodeForHtml(formatnumber(valor2,ndec,-1,0,-1))%>
			<%else%>
				<%="&nbsp;"%>
			<%end if%>
		<%end if%>
	</td>
<%
end sub

'************************ FUNCIONES ******************************************
'crea la tabla temporal para Proveedor
'sub crearProveedor(p_tsel1, p_tsel2,p_tsel3, p_dfecha, p_hfecha, p_tactividad, p_proveedor, p_serie, p_familia, p_referencia, p_nombreart, p_ordenar, p_mb,p_opcproveedorbaja,p_cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas)
'ega 18/07/2008 se ha sustituido el checkbox de Incluir Albaranes Pendientes de Facturar por la seleccion de series de albaran, por lo que el parametro p_tsel3 desaparece
sub crearProveedor(p_tsel1, p_tsel2, p_dfecha, p_hfecha, p_tactividad, p_proveedor, p_serie, p_familia, p_referencia, p_nombreart, p_ordenar, p_mb,p_opcproveedorbaja,p_cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas)
	set rstFunction = Server.CreateObject("ADODB.Recordset")
	set rstTemp = Server.CreateObject("ADODB.Recordset")
	'creamos la tabla temporal
	'borrar ="if exists (select * from sysobjects where id = object_id('" & session("usuario") & "') and sysstat " & _
	'" & 0xf = 3) drop table " & session("usuario")
	DropTable session("usuario"), session("backendlistados")
	crear =" CREATE TABLE [" & session("usuario") & "] (nproveedor varchar(10) NOT NULL ,"
	crear=crear & "Nombre varchar(50), "
	crear=crear & "Referencia varchar(30), "
	crear=crear & "Descripcion varchar(355), "
	crear=crear & "Cantidad real, "
	crear=crear & "[Compras Netas] money, "
	crear=crear & "[Precio Medio] money, "
	crear=crear & "cod_proyecto varchar(60), "
	crear=crear & "Divisa varchar(15),"
	crear=crear & "tiene_escv smallint,"
	crear=crear & "Acumulador money, "
	crear=crear & "Orden money,"
	crear=crear & "cantidad2 real,"
	crear=crear & "medidaVenta varchar(50),"
	crear=crear & "precioMedio2 money"
	
	if opc_coste="1" then
	    crear=crear & ",coste money "
	end if	
	crear=crear & ")"
	rst.open crear,session("backendlistados"),adUseClient,adLockReadOnly
	GrantUser session("usuario"), session("backendlistados")

	nproveedores = 0
	strwhere    =""
	strwhere2   =""
	strwhereall =""

	if p_dfecha>"" then
      	strwhereall=strwhereall & " and f.fecha>='" & p_dfecha & "' "
	end if
	if p_hfecha>"" then
      	strwhereall=strwhereall & " and f.fecha<='" & p_hfecha & "' "
	end if
				
	if p_tactividad>"" then 'Se selecciono tipo de actividad
      	strwhereall = strwhereall + " and p.tactividad='" & p_tactividad & "' "
	end if
	if p_proveedor>"" then 'Se selecciono Proveedor
      	strwhereall = strwhereall + " and f.nproveedor='" & p_proveedor & "' "
	else
		if p_opcproveedorbaja=1 then
			strwhereall = strwhereall + " and p.fbaja is null "
		else
			strbaja=""
		end if
	end if
	if p_serie>"" then 'Se selecciono serie
      	strwhereall = strwhereall + " and f.serie='" & p_serie & "' "
	end if
	if p_familia>"" then
		p_tsel2 = false
		strwhere = strwhere + " and a.familia='" & p_familia & "' "
	end if
	if p_referencia>"" then
      	strwhere = strwhere + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%' "
		tsel2 = false
	end if
	if p_nombreart>"" then
      	strwhere = strwhere + " and a.nombre like '%" & p_nombreart & "%' "
		strwhere2 = strwhere2 + " and d.descripcion like '%" & p_nombreart & "%' "
	end if

	if p_cod_proyecto>"" then
		strwhere=strwhere & " and cod_proyecto='" & p_cod_proyecto & "' "
		strwhere2=strwhere2 & " and cod_proyecto='" & p_cod_proyecto & "' "
	end if

	if tipo_proveedor & "">"" then
		strwhere = strwhere + " and p.tipo_proveedor='" & tipo_proveedor & "' "
		strwhere2 = strwhere2 + " and p.tipo_proveedor='" & tipo_proveedor & "' "
	end if

	if tipo_articulo & "">"" then
		p_tsel2 = false
		strwhere = strwhere + " and a.tipo_articulo='" & tipo_articulo & "' "
	end if
	if mostrarfilas & "">"" then
		strwhere=strwhere + ""
	end if

	strwhereall=strwhereall & " and f.nfactura like '" & session("ncliente") & "%' "

	seleccion1="select distinct f.nproveedor as nproveedor, "
	seleccion1=seleccion1 & "p.RAZON_SOCIAL as Nombre, "
	seleccion1=seleccion1 & "d.referencia as Referencia, "
	seleccion1=seleccion1 & "a.nombre as Descripcion,"
	seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "

''''	seleccion1=seleccion1 & "sum((((((d.importe*round((100-isnull(convert(money,f.descuento),0)),2))/100)*round((100-isnull(convert(money,f.descuento2),0)),2))/100)*round((100-isnull(convert(money,f.descuento3),0)),2))/100) as [Compras Netas], "
	seleccion1=seleccion1 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Compras Netas], "

''''	seleccion1=seleccion1 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*round((100-isnull(convert(money,f.descuento),0)),2))/100)*round((100-isnull(convert(money,f.descuento2),0)),2))/100)*round((100-isnull(convert(money,f.descuento3),0)),2))/100)/sum(d.cantidad)) end as [Precio Medio], "
	seleccion1=seleccion1 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

	seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto,"
	seleccion1=seleccion1 & "f.divisa as Divisa,"
	seleccion1=seleccion1 & "0 as tiene_escv, "
	'i(EJM 06/11/2006) Mostar los cantidad2
	seleccion1=seleccion1 & "sum(cantidad2) as cantidad2,"
	seleccion1=seleccion1 & "calcularIMPCantidad2,"
	seleccion1=seleccion1 & "(select top 1 descripcion from medidas with(nolock) where codigo=a.medidaVenta) as medidaVenta,"
	seleccion1=seleccion1 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2] "
	if opc_coste="1" then
	    seleccion1=seleccion1 & ",a.importe"
	end if	
	
	'fin(EJM 06/11/2006)
	seleccion1=seleccion1 & " from detalles_fac_pro as d with(nolock) "
	seleccion1=seleccion1 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura "
	seleccion1=seleccion1 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
	seleccion1=seleccion1 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion1=seleccion1 & " where d.mainitem is null "
	'ega 22/07/2008 likes
	seleccion1=seleccion1 & " and d.nfactura like '" & session("ncliente") & "%' "
	seleccion1=seleccion1 & " and a.referencia like '" & session("ncliente") & "%' "
	seleccion1=seleccion1 & " and p.nproveedor like '" & session("ncliente") & "%' "

	seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
	seleccion1=seleccion1 & " group by f.nproveedor, "
	seleccion1=seleccion1 & "p.RAZON_SOCIAL, "
	seleccion1=seleccion1 & "d.referencia, "
	seleccion1=seleccion1 & "a.nombre, "
	seleccion1=seleccion1 & "f.cod_proyecto, "
	seleccion1=seleccion1 & "f.divisa, "
	seleccion1=seleccion1 & "calcularIMPCantidad2, "
	seleccion1=seleccion1 & "medidaVenta "
	if opc_coste="1" then
	    seleccion1=seleccion1 & ",a.importe"
	end if		
	seleccion1=seleccion1 & " union all "
	seleccion1=seleccion1 & " select distinct f.nproveedor as nproveedor, "
	seleccion1=seleccion1 & "p.RAZON_SOCIAL as Nombre, "
	seleccion1=seleccion1 & "d.referencia as Referencia, "
	seleccion1=seleccion1 & "a.nombre as Descripcion,"
	seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "
	seleccion1=seleccion1 & "0 as [Compras Netas], "
	seleccion1=seleccion1 & "0 as [Precio Medio], "
	seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto,"
	seleccion1=seleccion1 & "f.divisa as Divisa,"
	seleccion1=seleccion1 & "1 as tiene_escv, "
	'i(EJM 06/11/2006) Mostar los cantidad2
	seleccion1=seleccion1 & "sum(cantidad2) as cantidad2,"
	seleccion1=seleccion1 & "calcularIMPCantidad2,"
	seleccion1=seleccion1 & "(select top 1 descripcion from medidas with(nolock) where codigo=a.medidaVenta) as medidaVenta,"
	seleccion1=seleccion1 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2] "
	if opc_coste="1" then
	    seleccion1=seleccion1 & ",a.importe"
	end if		
	'fin(EJM 06/11/2006)
	seleccion1=seleccion1 & " from detalles_fac_pro as d with(nolock) "
	seleccion1=seleccion1 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura "
	seleccion1=seleccion1 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
	seleccion1=seleccion1 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion1=seleccion1 & " where d.mainitem is not null "
	'ega 22/07/2008 likes
	seleccion1=seleccion1 & " and d.nfactura like '" & session("ncliente") & "%' "
	seleccion1=seleccion1 & " and a.referencia like '" & session("ncliente") & "%' "
	seleccion1=seleccion1 & " and p.nproveedor like '" & session("ncliente") & "%' "
	seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
	seleccion1=seleccion1 & " group by f.nproveedor, "
	seleccion1=seleccion1 & "p.RAZON_SOCIAL, "
	seleccion1=seleccion1 & "d.referencia, "
	seleccion1=seleccion1 & "a.nombre, "
	seleccion1=seleccion1 & "f.cod_proyecto, "
	seleccion1=seleccion1 & "f.divisa, "
	seleccion1=seleccion1 & "calcularIMPCantidad2, "
	seleccion1=seleccion1 & "medidaVenta "
	if opc_coste="1" then
	    seleccion1=seleccion1 & ",a.importe"
	end if		


	addselect = ""
	addgroup  = ""
	if desglose = true then
		'addselect2="d.nconcepto as Referencia,"
	      addselect = "d.descripcion as Descripcion, "
		addgroup  = "d.descripcion, "
		'addgroup2  = "d.nconcepto, "
	else
		addselect =  "'Concepto' as Descripcion, "
		'addselect2="'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
	end if

	seleccion2="select distinct f.nproveedor as nproveedor, "
	seleccion2=seleccion2 & "p.RAZON_SOCIAL as Nombre, "
	seleccion2=seleccion2 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
	seleccion2=seleccion2 & addselect
	seleccion2=seleccion2 & "sum(d.cantidad) as Cantidad, "

''	seleccion2=seleccion2 & "sum((((((d.importe*round((100-isnull(convert(money,f.descuento),0)),2))/100)*round((100-isnull(convert(money,f.descuento2),0)),2))/100)*round((100-isnull(convert(money,f.descuento3),0)),2))/100) as [Compras Netas], "
	seleccion2=seleccion2 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Compras Netas], "

''	seleccion2=seleccion2 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*round((100-isnull(convert(money,f.descuento),0)),2))/100)*round((100-isnull(convert(money,f.descuento2),0)),2))/100)*round((100-isnull(convert(money,f.descuento3),0)),2))/100)/sum(d.cantidad)) end as [Precio Medio], "
	seleccion2=seleccion2 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

	seleccion2=seleccion2 & " f.cod_proyecto, "
	seleccion2=seleccion2 & "f.divisa as Divisa,"
	seleccion2=seleccion2 & " 0 as tiene_escv, "
	'i(EJM 06/11/2006) Mostar los cantidad2
	seleccion2=seleccion2 & "0 as cantidad2,"
	seleccion2=seleccion2 & "0 as calcularIMPCantidad2,"
	seleccion2=seleccion2 & "'' as medidaVenta,"
	seleccion2=seleccion2 & "0 as [Precio Medio2]"
	if opc_coste="1" then
	    seleccion2=seleccion2 & ",0 as importe"
	end if		
	'fin(EJM 06/11/2006)
	seleccion2=seleccion2 & " from conceptos_fac_pro as d with(nolock)  "
	seleccion2=seleccion2 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura "
	seleccion2=seleccion2 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion2=seleccion2 & " where "
	'ega 22/07/2008 likes
	seleccion2=seleccion2 & "  d.nfactura like '" & session("ncliente") & "%' "
	seleccion2=seleccion2 & " and p.nproveedor like '" & session("ncliente") & "%' "
	seleccion2=seleccion2 & strwhereall & " " & strwhere2 & " "
	seleccion2=seleccion2 & "group by f.nproveedor, "
	seleccion2=seleccion2 & "p.RAZON_SOCIAL, "
	seleccion2=seleccion2 & addgroup & addgroup2
	seleccion2=seleccion2 & "f.cod_proyecto, "
	seleccion2=seleccion2 & "f.divisa "

    ''ega 18/07/2008 si se ha seleccionado alguna serie de albaran pendiente de facturar
	''if p_tsel3="on" or p_tsel3="1" then
	if seriesapf >"" then
		strwhereall3=""
		strwhereall3=strwhereall3 & " and alb.serie in " & lista_series & " "
		if p_dfecha>"" then
      		strwhereall3=strwhereall3 & " and alb.fecha>='" & p_dfecha & "'"
		end if
		if p_hfecha>"" then
      		strwhereall3=strwhereall3 & " and alb.fecha<='" & p_hfecha & "'"
		end if
		if p_tactividad>"" then 'Se selecciono tipo de actividad
      		strwhereall3 = strwhereall3 + " and p.tactividad='" & p_tactividad & "'"
		end if
		if p_proveedor>"" then 'Se selecciono proveedor
      		strwhereall3 = strwhereall3 + " and alb.nproveedor='" & p_proveedor & "'"
		else
		    if p_opcproveedorbaja=1 then
 				strbaja=" "
			else
				strbaja=" and p.fbaja is null "
				strwhereall3 = strwhereall3 + strbaja
			end if
		end if
		if p_familia>"" then
			strwhere3 = strwhere3 + " and a.familia='" & p_familia & "'"
		end if
		if p_referencia>"" then
		     	strwhere3 = strwhere3 + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		end if
		if p_nombreart>"" then
      		strwhere3 = strwhere3 + " and a.nombre like '%" & p_nombreart & "%'"
			strwhere4 = strwhere4 + " and d.descripcion like '%" & p_nombreart & "%'"
		end if
		if p_cod_proyecto>"" then
			strwhere3=strwhere3 & " and cod_proyecto='" & p_cod_proyecto & "'"
			strwhere4=strwhere4 & " and cod_proyecto='" & p_cod_proyecto & "'"
		end if

		if tipo_proveedor & "">"" then
			strwhere3 = strwhere3 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
			strwhere4 = strwhere4 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
		end if

		if tipo_articulo & "">"" then
			strwhere3 = strwhere3 + " and a.tipo_articulo='" & tipo_articulo & "'"
		end if
		if mostrarfilas & "">"" then
			strwhere3=strwhere3 + ""
		end if

		strwhereall3=strwhereall3 & " and alb.nalbaran like '" & session("ncliente") & "%'"

		seleccion3="select distinct alb.nproveedor as NProveedor, "
		seleccion3=seleccion3 & "p.razon_social as Nombre, "
		seleccion3=seleccion3 & "d.referencia as Referencia, "
		seleccion3=seleccion3 & "a.nombre as Descripcion,"
		seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "

	    seleccion3=seleccion3 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Compras Netas], "
	    seleccion3=seleccion3 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

		seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
		seleccion3=seleccion3 & "alb.divisa as Divisa,"
		seleccion3=seleccion3 & "0 as tiene_escv, "
		'i(EJM 06/11/2006) Mostar los cantidad2
		seleccion3=seleccion3 & "sum(cantidad2) as cantidad2,"
		seleccion3=seleccion3 & "calcularIMPCantidad2,"
		seleccion3=seleccion3 & "(select top 1 descripcion from medidas with(nolock) where codigo=a.medidaVenta) as medidaVenta,"
		seleccion3=seleccion3 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2] "
	    if opc_coste="1" then
	       seleccion3=seleccion3 & ",a.importe"
    	end if			
		'fin(EJM 06/11/2006)
		seleccion3=seleccion3 & " from detalles_alb_pro as d with(nolock) "
		seleccion3=seleccion3 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion3=seleccion3 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
		seleccion3=seleccion3 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion3=seleccion3 & " where alb.nfactura is null "
		'ega 22/07/2008 likes
	    seleccion3=seleccion3 & " and d.nalbaran like '" & session("ncliente") & "%' "
	    seleccion3=seleccion3 & " and alb.nalbaran like '" & session("ncliente") & "%' "
	    seleccion3=seleccion3 & " and a.referencia like '" & session("ncliente") & "%' "
	    seleccion3=seleccion3 & " and p.nproveedor like '" & session("ncliente") & "%' "
    	
		seleccion3=seleccion3 & strwhereall3 & " " & strwhere3 & " "
		seleccion3=seleccion3 & " and d.mainitem is null "
		seleccion3=seleccion3 & " group by alb.nproveedor , "
		seleccion3=seleccion3 & "p.razon_social, "
		seleccion3=seleccion3 & "d.referencia, "
		seleccion3=seleccion3 & "a.nombre, "
		seleccion3=seleccion3 & "alb.cod_proyecto, "
		seleccion3=seleccion3 & "alb.divisa, "
		seleccion3=seleccion3 & "calcularIMPCantidad2, "
		seleccion3=seleccion3 & "medidaVenta "
	    if opc_coste="1" then
	       seleccion3=seleccion3 & ",a.importe"
    	end if			
		
		seleccion3=seleccion3 & " union all "
		seleccion3=seleccion3 & " select distinct alb.nproveedor as NProveedor, "
		seleccion3=seleccion3 & "p.razon_social as Nombre, "
		seleccion3=seleccion3 & "d.referencia as Referencia, "
		seleccion3=seleccion3 & "a.nombre as Descripcion,"
		seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "
		seleccion3=seleccion3 & "0 as [Compras Netas], "
		seleccion3=seleccion3 & "0 as [Precio Medio], "
		seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
		seleccion3=seleccion3 & "alb.divisa as Divisa,"
		seleccion3=seleccion3 & "1 as tiene_escv, "
		'i(EJM 06/11/2006) Mostar los cantidad2
		seleccion3=seleccion3 & "sum(cantidad2) as cantidad2,"
		seleccion3=seleccion3 & "calcularIMPCantidad2,"
		seleccion3=seleccion3 & "(select top 1 descripcion from medidas with(nolock) where codigo=a.medidaVenta) as medidaVenta,"
		seleccion3=seleccion3 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2] "
	    if opc_coste="1" then
	       seleccion3=seleccion3 & ",a.importe"
    	end if			
		
		'fin(EJM 06/11/2006)
		seleccion3=seleccion3 & " from detalles_alb_pro as d with(nolock) "
		seleccion3=seleccion3 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion3=seleccion3 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
		seleccion3=seleccion3 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion3=seleccion3 & " where alb.nfactura is null "
		'ega 22/07/2008 likes
	    seleccion3=seleccion3 & " and d.nalbaran like '" & session("ncliente") & "%' "
	    seleccion3=seleccion3 & " and alb.nalbaran like '" & session("ncliente") & "%' "
	    seleccion3=seleccion3 & " and a.referencia like '" & session("ncliente") & "%' "
	    seleccion3=seleccion3 & " and p.nproveedor like '" & session("ncliente") & "%' "
	    
		seleccion3=seleccion3 & strwhereall3 & " " & strwhere3 & " "
		seleccion3=seleccion3 & " and d.mainitem is not null "
		seleccion3=seleccion3 & " group by alb.nproveedor , "
		seleccion3=seleccion3 & "p.razon_social, "
		seleccion3=seleccion3 & "d.referencia, "
		seleccion3=seleccion3 & "a.nombre, "
		seleccion3=seleccion3 & "alb.cod_proyecto, "
		seleccion3=seleccion3 & "alb.divisa, "
		seleccion3=seleccion3 & "calcularIMPCantidad2, "
		seleccion3=seleccion3 & "medidaVenta "
	    if opc_coste="1" then
	       seleccion3=seleccion3 & ",a.importe"
    	end if			

		addselect3 = ""
		addgroup3  = ""
		if desglose = true then
      		addselect3 = "CONVERT(nvarchar,d .descripcion) as Descripcion, "
			addgroup3  = "CONVERT(nvarchar,d .descripcion), "
		else
			addselect3 =  "'Concepto' as Descripcion, "
		end if

		seleccion4="select distinct alb.nproveedor as NProveedor, "
        seleccion4=seleccion4 & "p.razon_social as Nombre, "
        seleccion4=seleccion4 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
        seleccion4=seleccion4 & addselect3
        seleccion4=seleccion4 & "sum(d.cantidad) as Cantidad, "

        seleccion4=seleccion4 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Compras Netas], "
        seleccion4=seleccion4 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

        seleccion4=seleccion4 & " alb.cod_proyecto, "
        seleccion4=seleccion4 & "alb.divisa as Divisa,"
        seleccion4=seleccion4 & "0 as tiene_escv, "
	    'i(EJM 06/11/2006) Mostar los cantidad2
	    seleccion4=seleccion4 & "0 as cantidad2,"
	    seleccion4=seleccion4 & "0 as calcularIMPCantidad2,"
	    seleccion4=seleccion4 & "'' as medidaVenta,"
	    seleccion4=seleccion4 & "0 as [Precio Medio2]"
	    if opc_coste="1" then
	       seleccion4=seleccion4 & ",0 as importe "
    	end if		    
	    'fin(EJM 06/11/2006)
        seleccion4=seleccion4 & " from conceptos_alb_pro as d with(nolock) "
        seleccion4=seleccion4 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
        seleccion4=seleccion4 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
        seleccion4=seleccion4 & " where alb.nfactura is null "
        'ega 22/07/2008 likes
	    seleccion4=seleccion4 & " and d.nalbaran like '" & session("ncliente") & "%' "
	    seleccion4=seleccion4 & " and alb.nalbaran like '" & session("ncliente") & "%' "
	    seleccion4=seleccion4 & " and p.nproveedor like '" & session("ncliente") & "%' "	    
        seleccion4=seleccion4 & strwhereall3 & " " & strwhere4 & " "
        seleccion4=seleccion4 & " group by alb.nproveedor , "
        seleccion4=seleccion4 & "p.razon_social, "
        seleccion4=seleccion4 & addgroup3
        seleccion4=seleccion4 & "alb.cod_proyecto, "
        seleccion4=seleccion4 & "alb.divisa "
	end if
	seleccion = ""
	if p_tsel1=true then
		seleccion = seleccion & seleccion1
	end if
	if p_tsel1=true and p_tsel2=true then
		seleccion = seleccion & " union ALL "
	end if
	if p_tsel2 = true then
		seleccion = seleccion & seleccion2
	end if

	''ega 18/07/2008 si se ha seleccionado alguna serie de albaran pendiente de facturar
	''if p_tsel3="on" or p_tsel3="1" then
	if seriesapf >"" then
		if p_tsel1=true then
			seleccion=seleccion & " union ALL " & seleccion3
		end if
		if p_tsel2 = true then
			seleccion=seleccion & " union ALL " & seleccion4
		end if
	end if

	seleccion = seleccion & " order by f.nproveedor"
	if p_tsel1= true then seleccion = seleccion & ", d.referencia"

    rstFunction.open seleccion, session("backendlistados"),1,3
    if not rstFunction.eof then
         acumulado = 0
	     elTotal = 0
	     proveedor_anterior = ""
	     rstTemp.open "select * from [" & session("usuario") & "]", session("backendlistados"), adOpenKeyset, adLockOptimistic
	     proveedorAnterior = ""
	     while not rstFunction.eof
	       if rstFunction("nproveedor")<>proveedorAnterior then
              ''ricardo 17/12/2003
	          ''elTotal = elTotal + formatnumber(acumulado,n_decimalesMB,-1,0,-1)
	          elTotal = elTotal + acumulado
              ''''''''''
	          acumulado = 0
	       end if
	       proveedorAnterior = rstFunction("nproveedor")
	       rstTemp.Addnew
	       rstTemp("nproveedor") = rstFunction("nproveedor")
	       rstTemp("Nombre") = rstFunction("Nombre")
 	       rstTemp("Referencia") = rstFunction("Referencia")
	       rstTemp("Descripcion") = rstFunction("Descripcion")
	       rstTemp("cod_proyecto") = d_lookup("nombre","proyectos","codigo='" & rstFunction("cod_proyecto") & "'",session("backendlistados"))
	       rstTemp("Cantidad") = rstFunction("Cantidad")
	       rstTemp("Compras Netas") = rstFunction("Compras Netas")
	       rstTemp("Precio Medio") = rstFunction("Precio Medio")
	       rstTemp("Divisa") = rstFunction("Divisa")
	       if rstTemp("Divisa") = p_mb then
		    rstTemp("Acumulador") = acumulado + rstTemp("Compras Netas")
		    acumulado = rstTemp("Acumulador")
	       else
		      rstTemp("Acumulador") = acumulado + CambioDivisa(rstTemp("Compras Netas"), rstTemp("Divisa"), p_mb)
		      acumulado = rstTemp("Acumulador")
	       end if
	       rstTemp("tiene_escv")=rstFunction("tiene_escv")
	       rstTemp("Orden") = 0
	       rstTemp("cantidad2")=rstFunction("cantidad2")
	       rstTemp("medidaVenta")=rstFunction("medidaVenta")
	       rstTemp("precioMedio2")=rstFunction("Precio Medio2")
	        if opc_coste="1" then
	           rstTemp("coste")=rstFunction("importe")
    	    end if			

	          rstTemp.Update
		      rstFunction.movenext
	       wend
	       rstFunction.close
	       rstTemp.close

	       'ordenamos por compras si es necesario
	       if ordenar=true then
 	          rstFunction.open "select distinct nproveedor from [" & session("usuario") & "]", session("backendlistados"),1,3
	          while not rstFunction.eof
			     rstTemp.open "update [" & session("usuario") & "] set orden = (select max(acumulador) from [" & _
			                  session("usuario") & "] where nproveedor = '" & rstFunction("nproveedor") & _
						      "') where nproveedor ='" & rstFunction("nproveedor") & "'", session("backendlistados"), 1, 3
			    rstFunction.movenext
		      wend
		      rstFunction.close
	       end if
	    end if

	    ''ricardo 17/12/2003 el total se calcula de esta manera
	    rstTemp.open "select sum([Compras Netas]/d.factcambio) as total from [" & session("usuario") & "] as temp1,divisas as d where d.codigo=temp1.divisa", session("backendlistados"),1,3
	    if not rstTemp.eof then
		    ''elTotal=rstTemp("total")
		    elTotal=formatnumber(null_z(rstTemp("total")),n_decimalesMB,-1,0,-1)
	    end if
	    rstTemp.close
    ''elTotal=0
    ''rstTemp.open "select sum([Compras Netas]) as total,temp1.divisa from [" & session("usuario") & "] as temp1 group by temp1.divisa", session("backendlistados"),1,3
    ''while not rstTemp.eof
    ''	if rstTemp("Divisa") = p_mb then
    ''		elTotal=elTotal + rstTemp("total")
    ''	else
    ''		elTotal=elTotal + CambioDivisa(rstTemp("total"), rstTemp("Divisa"), p_mb)
    ''	end if
    ''	rstTemp.movenext
    ''wend
    ''rstTemp.close

    ''ricardo 17/12/2003 el total se calcula de esta manera
    '''	elTotal = elTotal + formatnumber(null_z(acumulado),n_decimalesMB,-1,0,-1)
	%><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' /><%                                   
end sub
'-----------------------------------------------------
'crea la tabla temporal para proyectos
'ega 18/07/2008 se ha sustituido el checkbox de Incluir Albaranes Pendientes de Facturar por la seleccion de series de albaran, por lo que el parametro p_tsel3 desaparece
'sub crearProyecto (p_tsel1, p_tsel2,p_tsel3, p_dfecha, p_hfecha, p_tactividad, p_proveedor, p_serie, p_familia, p_referencia, p_nombreart, p_ordenar, p_mb,p_opcproveedorbaja,p_cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas)
sub crearProyecto (p_tsel1, p_tsel2,p_dfecha, p_hfecha, p_tactividad, p_proveedor, p_serie, p_familia, p_referencia, p_nombreart, p_ordenar, p_mb,p_opcproveedorbaja,p_cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas)
	set rstFunction = Server.CreateObject("ADODB.Recordset")
	set rstTemp = Server.CreateObject("ADODB.Recordset")
	'creamos la tabla temporal
	'borrar ="if exists (select * from sysobjects where id = object_id('" & session("usuario") & "') and sysstat " & _
	'" & 0xf = 3) drop table " & session("usuario")
	DropTable session("usuario"), session("backendlistados")
	crear="CREATE TABLE [" & session("usuario") & "] (cod_proyecto varchar(60) not null,NProveedor varchar(10) ,"
	crear=crear & "Nombre varchar(50), "
	crear=crear & "Referencia varchar(30), "
	crear=crear & "Descripcion varchar(355), "
	crear=crear & "Cantidad real, "
	crear=crear & "[Compras Netas] money, "
	crear=crear & "[Precio Medio] money, "
	crear=crear & "Divisa varchar(15), "
	crear=crear & "tiene_escv smallint,"
	crear=crear & "Acumulador money, "
	crear=crear & "Orden money)"

	rst.open crear,session("backendlistados"),adUseClient,adLockReadOnly
	GrantUser session("usuario"), session("backendlistados")

	nproveedores = 0
	strwhere    =""
	strwhere2   =""
	strwhereall =""

	if p_dfecha>"" then
		strwhereall=strwhereall & " and f.fecha>='" & p_dfecha & "'"
	end if
	if p_hfecha>"" then
		strwhereall=strwhereall & " and f.fecha<='" & p_hfecha & "'"
	end if
	if p_tactividad>"" then 'Se selecciono tipo de actividad
		strwhereall = strwhereall + " and p.tactividad='" & p_tactividad & "'"
	end if
	if p_proveedor>"" then 'Se selecciono proveedor
		strwhereall = strwhereall + " and f.nproveedor='" & p_proveedor & "'"
	else
		if p_opcproveedorbaja=1 then
			strbaja=" "
		else
			strbaja=" and p.fbaja is null "
			strwhereall = strwhereall + strbaja
		end if
	end if
	if p_serie>"" then 'Se selecciono serie
		strwhereall = strwhereall + " and f.serie='" & p_serie & "'"
	end if
	if p_familia>"" then
		p_tsel2 = false
		strwhere = strwhere + " and a.familia='" & p_familia & "'"
	end if
	if p_referencia>"" then
      	strwhere = strwhere + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		tsel2 = false
	end if
	if p_nombreart>"" then
		strwhere = strwhere + " and a.nombre like '%" & p_nombreart & "%'"
		strwhere2 = strwhere2 + " and d.descripcion like '%" & p_nombreart & "%'"
	end if
	if p_cod_proyecto>"" then
		strwhere=strwhere & " and cod_proyecto='" & p_cod_proyecto & "'"
		strwhere2=strwhere2 & " and cod_proyecto='" & p_cod_proyecto & "'"
	end if

	if tipo_proveedor & "">"" then
		strwhere = strwhere + " and p.tipo_proveedor='" & tipo_proveedor & "'"
		strwhere2 = strwhere2 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
	end if

	if tipo_articulo & "">"" then
		p_tsel2 = false
		strwhere = strwhere + " and a.tipo_articulo='" & tipo_articulo & "'"
	end if
	if mostrarfilas & "">"" then
		strwhere=strwhere + ""
	end if

	strwhereall=strwhereall & " and f.nfactura like '" & session("ncliente") & "%'"

	seleccion1="select distinct f.cod_proyecto as cod_proyecto,f.nproveedor as nproveedor, "
	seleccion1=seleccion1 & "p.razon_social as Nombre, "
	seleccion1=seleccion1 & "d.referencia as Referencia, "
	seleccion1=seleccion1 & "a.nombre as Descripcion, "
	seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "

	seleccion1=seleccion1 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Compras Netas], "
	seleccion1=seleccion1 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

	seleccion1=seleccion1 & "f.divisa as Divisa,"
	seleccion1=seleccion1 & "0 as tiene_escv "
	seleccion1=seleccion1 & " from detalles_fac_pro as d with(nolock) "
	seleccion1=seleccion1 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura "
	seleccion1=seleccion1 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
	seleccion1=seleccion1 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion1=seleccion1 & " inner join proyectos as pr with(nolock) on pr.codigo=f.cod_proyecto "
	seleccion1=seleccion1 & " where d.mainitem is null "
    'ega 22/07/2008 likes
    seleccion1=seleccion1 & " and d.nfactura like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and a.referencia like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
    seleccion1=seleccion1 & " and pr.codigo like '" & session("ncliente") & "%' "
	    
	seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
	seleccion1=seleccion1 & "group by f.cod_proyecto,f.nproveedor, "
	seleccion1=seleccion1 & "p.razon_social, "
	seleccion1=seleccion1 & "d.referencia, "
	seleccion1=seleccion1 & "a.nombre, "
	seleccion1=seleccion1 & "f.divisa "
	seleccion1=seleccion1 & " union all "
	seleccion1=seleccion1 & "select distinct f.cod_proyecto as cod_proyecto,f.nproveedor as nproveedor, "
	seleccion1=seleccion1 & "p.razon_social as Nombre, "
	seleccion1=seleccion1 & "d.referencia as Referencia, "
	seleccion1=seleccion1 & "a.nombre as Descripcion, "
	seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "
	seleccion1=seleccion1 & "0 as [Compras Netas], "
	seleccion1=seleccion1 & "0 as [Precio Medio], "
	seleccion1=seleccion1 & "f.divisa as Divisa,"
	seleccion1=seleccion1 & "1 as tiene_escv "
	seleccion1=seleccion1 & " from detalles_fac_pro as d with(nolock) "
	seleccion1=seleccion1 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura  "
	seleccion1=seleccion1 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
	seleccion1=seleccion1 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion1=seleccion1 & " inner join proyectos as pr with(nolock) on pr.codigo=f.cod_proyecto "
	seleccion1=seleccion1 & " where d.mainitem is not null "
	'ega 22/07/2008 likes
    seleccion1=seleccion1 & " and d.nfactura like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and a.referencia like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
    seleccion1=seleccion1 & " and pr.codigo like '" & session("ncliente") & "%' "    
	seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
	seleccion1=seleccion1 & " group by f.cod_proyecto,f.nproveedor, "
	seleccion1=seleccion1 & "p.razon_social, "
	seleccion1=seleccion1 & "d.referencia, "
	seleccion1=seleccion1 & "a.nombre, "
	seleccion1=seleccion1 & "f.divisa "

	addselect = ""
	addgroup  = ""
	if desglose = true then
		addselect = "d.descripcion as Descripcion, "
		addgroup  = "d.descripcion, "
	else
		addselect =  "'Concepto' as Descripcion, "
	end if

	seleccion2="select distinct f.cod_proyecto,f.nproveedor as nproveedor, "
	seleccion2=seleccion2 & "p.razon_social as Nombre, "
	seleccion2=seleccion2 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
	seleccion2=seleccion2 & addselect
	seleccion2=seleccion2 & "sum(d.cantidad) as Cantidad, "

	seleccion2=seleccion2 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Compras Netas], "
	seleccion2=seleccion2 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

	seleccion2=seleccion2 & "f.divisa as Divisa,"
	seleccion2=seleccion2 & "0 as tiene_escv "
	seleccion2=seleccion2 & " from conceptos as d with(nolock) "
	seleccion2=seleccion2 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura "
	seleccion2=seleccion2 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion2=seleccion2 & " inner join proyectos as pr on pr.codigo=f.cod_proyecto "
	seleccion2=seleccion2 & " where  "
	'ega 22/07/2008 likes
    seleccion2=seleccion2 & " d.nfactura like '" & session("ncliente") & "%' "
    seleccion2=seleccion2 & " and f.nfactura like '" & session("ncliente") & "%' "
    seleccion2=seleccion2 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
    seleccion2=seleccion2 & " and pr.codigo like '" & session("ncliente") & "%' "
    
	seleccion2=seleccion2 & strwhereall & " " & strwhere2 & " "
	seleccion2=seleccion2 & " group by f.cod_proyecto,f.nproveedor, "
	seleccion2=seleccion2 & "p.razon_social, "
	seleccion2=seleccion2 & addgroup
	seleccion2=seleccion2 & "f.divisa "
    ''ega 18/07/2008 si se ha seleccionado alguna serie de albaran pendiente de facturar
	'if p_tsel3="on" or p_tsel3="1" then
	if seriesapf >"" then

		strwhereall3=""
		if p_dfecha>"" then
			strwhereall3=strwhereall3 & " and alb.fecha>='" & p_dfecha & "'"
		end if
		if p_hfecha>"" then
			strwhereall3=strwhereall3 & " and alb.fecha<='" & p_hfecha & "'"
		end if
		if p_tactividad>"" then 'Se selecciono tipo de actividad
			strwhereall3 = strwhereall3 + " and p.tactividad='" & p_tactividad & "'"
		end if
		if p_proveedor>"" then 'Se selecciono proveedor
			strwhereall3 = strwhereall3 + " and alb.nproveedor='" & p_proveedor & "'"
		else
			if p_opcproveedorbaja=1 then
				strbaja=" "
			else
				strbaja=" and p.fbaja is null "
				strwhereall3 = strwhereall3 + strbaja
			end if
		end if
		if p_serie>"" then 'Se selecciono serie
			strwhereall3 = strwhereall3 + " and alb.serie='" & p_serie & "'"
		end if
		if p_familia>"" then
			strwhere3 = strwhere3 + " and a.familia='" & p_familia & "'"
		end if
		if p_referencia>"" then
	      	strwhere3 = strwhere3 + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		end if
		if p_nombreart>"" then
			strwhere3 = strwhere3 + " and a.nombre like '%" & p_nombreart & "%'"
			strwhere4 = strwhere4 + " and d.descripcion like '%" & p_nombreart & "%'"
		end if
		if p_cod_proyecto>"" then
			strwhere3=strwhere3 & " and cod_proyecto='" & p_cod_proyecto & "'"
			strwhere4=strwhere4 & " and cod_proyecto='" & p_cod_proyecto & "'"
		end if

		if tipo_proveedor & "">"" then
			strwhere3 = strwhere3 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
			strwhere4 = strwhere4 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
		end if

		if tipo_articulo & "">"" then
			strwhere3 = strwhere3 + " and a.tipo_articulo='" & tipo_articulo & "'"
		end if
		if mostrarfilas & "">"" then
			strwhere3=strwhere3 + ""
		end if

		strwhereall3=strwhereall3 & " and alb.nalbaran like '" & session("ncliente") & "%'"

		seleccion3="select distinct alb.cod_proyecto as cod_proyecto,alb.nproveedor as NProveedor, "
		seleccion3=seleccion3 & "p.razon_social as Nombre, "
		seleccion3=seleccion3 & "d.referencia as Referencia, "
		seleccion3=seleccion3 & "a.nombre as Descripcion, "
		seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "

	    seleccion3=seleccion3 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Compras Netas], "
	    seleccion3=seleccion3 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

		seleccion3=seleccion3 & "alb.divisa as Divisa,"
		seleccion3=seleccion3 & "0 as tiene_escv "
		seleccion3=seleccion3 & " from detalles_alb_pro as d with(nolock) "
		seleccion3=seleccion3 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion3=seleccion3 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
		seleccion3=seleccion3 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion3=seleccion3 & " inner join proyectos as pr with(nolock) on pr.codigo=alb.cod_proyecto "
		seleccion3=seleccion3 & " where d.mainitem is null "
		'ega 22/07/2008 likes
        seleccion3=seleccion3 & " and d.nalbaran like '" & session("ncliente") & "%' "
        seleccion3=seleccion3 & " and a.referencia like '" & session("ncliente") & "%' "	  
        seleccion3=seleccion3 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
        seleccion3=seleccion3 & " and pr.codigo like '" & session("ncliente") & "%' "    
		seleccion3=seleccion3 & strwhereall3 & " " & strwhere3 & " "
		seleccion3=seleccion3 & " group by alb.cod_proyecto,alb.nproveedor, "
		seleccion3=seleccion3 & "p.razon_social, "
		seleccion3=seleccion3 & "d.referencia, "
		seleccion3=seleccion3 & "a.nombre, "
		seleccion3=seleccion3 & "alb.divisa "
		seleccion3=seleccion3 & " union all "
		seleccion3=seleccion3 & "select distinct alb.cod_proyecto as cod_proyecto,alb.nproveedor as NProveedor, "
		seleccion3=seleccion3 & "p.razon_social as Nombre, "
		seleccion3=seleccion3 & "d.referencia as Referencia, "
		seleccion3=seleccion3 & "a.nombre as Descripcion, "
		seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "
		seleccion3=seleccion3 & "0 as [Compras Netas], "
		seleccion3=seleccion3 & "0 as [Precio Medio], "
		seleccion3=seleccion3 & "alb.divisa as Divisa,"
		seleccion3=seleccion3 & "1 as tiene_escv "
		seleccion3=seleccion3 & "from detalles_alb_pro as d with(nolock) "
		seleccion3=seleccion3 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion3=seleccion3 & " inner join articulos as a with(nolock) on a.referencia=d.referencia "
		seleccion3=seleccion3 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion3=seleccion3 & " inner join proyectos as pr with(nolock) on pr.codigo=alb.cod_proyecto "
		seleccion3=seleccion3 & "where d.mainitem is not null "
	    'ega 22/07/2008 likes
        seleccion3=seleccion3 & " and d.nalbaran like '" & session("ncliente") & "%' "
        seleccion3=seleccion3 & " and a.referencia like '" & session("ncliente") & "%' "	  
        seleccion3=seleccion3 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
        seleccion3=seleccion3 & " and pr.codigo like '" & session("ncliente") & "%' "
        
		seleccion3=seleccion3 & strwhereall3 & " " & strwhere3 & " "
		seleccion3=seleccion3 & " group by alb.cod_proyecto,alb.nproveedor, "
		seleccion3=seleccion3 & "p.razon_social, "
		seleccion3=seleccion3 & "d.referencia, "
		seleccion3=seleccion3 & "a.nombre, "
		seleccion3=seleccion3 & "alb.divisa "

		addselect3 = ""
		addgroup3  = ""
		if desglose = true then
			addselect3 = "CONVERT(nvarchar,d .descripcion) as Descripcion, "
			addgroup3  = "CONVERT(nvarchar,d .descripcion), "
		else
			addselect3 =  "'Concepto' as Descripcion, "
		end if

		seleccion4="select distinct alb.cod_proyecto,alb.nproveedor as NProveedor, "
		seleccion4=seleccion4 & "p.razon_social as Nombre, "
		seleccion4=seleccion4 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
		seleccion4=seleccion4 & addselect3
		seleccion4=seleccion4 & "sum(d.cantidad) as Cantidad, "

	    seleccion4=seleccion4 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Compras Netas], "
	    seleccion4=seleccion4 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

		seleccion4=seleccion4 & "alb.divisa as Divisa,"
		seleccion4=seleccion4 & "0 as tiene_escv "
		seleccion4=seleccion4 & " from conceptos_alb_pro as d with(nolock) "
		seleccion4=seleccion4 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion4=seleccion4 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion4=seleccion4 & " inner join proyectos as pr with(nolock) on pr.codigo=alb.cod_proyecto "
		seleccion4=seleccion4 & " where alb.nfactura is null "
        'ega 22/07/2008 likes
        seleccion4=seleccion4 & " and d.nalbaran like '" & session("ncliente") & "%' "
        seleccion4=seleccion4 & " and alb.nalbaran like '" & session("ncliente") & "%' "        	  
        seleccion4=seleccion4 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
        seleccion4=seleccion4 & " and pr.codigo like '" & session("ncliente") & "%' "
        
		seleccion4=seleccion4 & strwhereall3 & " " & strwhere4 & " "
		seleccion4=seleccion4 & " group by alb.cod_proyecto,alb.nproveedor, "
		seleccion4=seleccion4 & "p.razon_social, "
		seleccion4=seleccion4 & addgroup3
		seleccion4=seleccion4 & "alb.divisa "
	end if

	seleccion = ""
	if tsel1=true then
		seleccion = seleccion & seleccion1
	end if
	if tsel1=true and tsel2=true then
		seleccion = seleccion & " union ALL "
	end if
	if tsel2 = true then
		seleccion = seleccion & seleccion2
	end if

    ''ega 18/07/2008 si se ha seleccionado alguna serie de albaran pendiente de facturar
	'if p_tsel3="on" or p_tsel3="1" then
    if seriesapf >"" then
		if tsel1=true then
			seleccion=seleccion & " union ALL " & seleccion3
		end if
		if tsel2 = true then
			seleccion=seleccion & " union ALL " & seleccion4
		end if
	end if

	seleccion = seleccion & " order by f.cod_proyecto,f.nproveedor"
	if tsel1= true then seleccion = seleccion & ", d.referencia"

	rstFunction.open seleccion, session("backendlistados"),1,3
	if not rstFunction.eof then
		acumulado = 0
		elTotal = 0
		cod_proyectoAnterior = ""
		rstTemp.open "select * from [" & session("usuario") & "]", session("backendlistados"), adOpenKeyset, adLockOptimistic
		cod_proyectoAnterior = ""
		while not rstFunction.eof
			if ucase(rstFunction("cod_proyecto"))<>ucase(cod_proyectoAnterior) then
				''elTotal = elTotal + formatnumber(acumulado,n_decimalesMB,-1,0,-1)
				elTotal = elTotal + acumulado
				acumulado = 0
			end if
			cod_proyectoAnterior = ucase(rstFunction("cod_proyecto"))
			rstTemp.Addnew
			rstTemp("cod_proyecto") = d_lookup("nombre","proyectos","codigo='" & rstFunction("cod_proyecto") & "'",session("backendlistados"))
			rstTemp("nproveedor") = rstFunction("nproveedor")
			rstTemp("Nombre") = rstFunction("Nombre")
			rstTemp("Referencia") = rstFunction("Referencia")
			rstTemp("Descripcion") = rstFunction("Descripcion")
			rstTemp("Cantidad") = rstFunction("Cantidad")
			rstTemp("Compras Netas") = rstFunction("Compras Netas")
			rstTemp("Precio Medio") = rstFunction("Precio Medio")
			rstTemp("Divisa") = rstFunction("Divisa")
			if rstTemp("Divisa") = p_mb then
				rstTemp("Acumulador") = acumulado + rstTemp("Compras Netas")
				acumulado = rstTemp("Acumulador")
			else
				rstTemp("Acumulador") = acumulado + CambioDivisa(rstTemp("Compras Netas"), rstTemp("Divisa"), p_mb)
				acumulado = rstTemp("Acumulador")
			end if
			rstTemp("tiene_escv")=rstFunction("tiene_escv")
			rstTemp("Orden") = 0
			rstTemp.Update
			rstFunction.movenext
		wend
		rstFunction.close
		rstTemp.close

		'ordenamos por compras si es necesario
		if ordenar=true then
			rstFunction.open "select distinct cod_proyecto from [" & session("usuario") & "]", session("backendlistados"),1,3
			while not rstFunction.eof
				rstTemp.open "update [" & session("usuario") & "] set orden = (select max(acumulador) from [" & _
			              session("usuario") & "] where cod_proyecto = '" & rstFunction("cod_proyecto") & _
						  "') where cod_proyecto ='" & rstFunction("cod_proyecto") & "'", session("backendlistados"), 1, 3
				rstFunction.movenext
			wend
			rstFunction.close
		end if
	end if

	''elTotal = elTotal + formatnumber(acumulado,n_decimalesMB,-1,0,-1)
	elTotal = elTotal + acumulado
	%><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' /><%
end sub
'-----------------------------------------------------
' crea la tabla temporal para articulos
'ega 18/07/2008 se ha sustituido el checkbox de Incluir Albaranes Pendientes de Facturar por la seleccion de series de albaran, por lo que el parametro p_tsel3 desaparece
'sub crearArticulo(p_tsel1, p_tsel2,p_tsel3, p_dfecha, p_hfecha, p_familia, p_referencia, p_nombreart, p_tactividad, p_proveedor, p_serie, p_desglose, p_ordenar, p_mb,p_opcproveedorbaja,p_cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas)
sub crearArticulo(p_tsel1, p_tsel2, p_dfecha, p_hfecha, p_familia, p_referencia, p_nombreart, p_tactividad, p_proveedor, p_serie, p_desglose, p_ordenar, p_mb,p_opcproveedorbaja,p_cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas)

	set rstFunction = Server.CreateObject("ADODB.Recordset")
	set rstTemp = Server.CreateObject("ADODB.Recordset")
	'creamos la tabla temporal
	DropTable session("usuario"), session("backendlistados")
	crear="CREATE TABLE [" & session("usuario") & "] (Ref varchar(30) NOT NULL ,"
	crear=crear & "Descripcion varchar(100), "
	crear=crear & "nproveedor varchar(10), "
	crear=crear & "Nombre varchar(100), "
	crear=crear & "Cantidad real, "
	crear=crear & "[Compras Netas] money, "
	crear=crear & "[Precio Medio] money, "
	crear=crear & "cod_proyecto varchar(60), "
	crear=crear & "Divisa varchar(15), "
	crear=crear & "tiene_escv smallint,"
	crear=crear & "AcumulaCompras money, "
	crear=crear & "AcumulaCantidad real, "
	crear=crear & "Orden money)"

	rst.open crear,session("backendlistados"),adUseClient,adLockReadOnly
	GrantUser session("usuario"), session("backendlistados")

	strwhere     = ""
	strwhereall  = ""
	strwhere2    = ""

	if p_dfecha>"" then
		strwhereall=strwhereall & " and f.fecha>='" & p_dfecha & "'"
	end if

	if p_hfecha>"" then
		strwhereall=strwhereall & " and f.fecha<='" & p_hfecha & "'"
	end if
	PorArticulo = "SI"
	if p_familia>"" then
		p_tsel2 = false
		strwhere = strwhere + " and a.familia='" & p_familia & "'"
	end if
	if p_referencia>"" then
      	strwhere = strwhere + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		p_tsel2=false
	end if
	if p_nombreart>"" then
		strwhere = strwhere + " and a.nombre like '%" & p_nombreart & "%'"
		strwhere2 = strwhere2 + " and d.descripcion like '%" & p_nombreart & "%'"
	end if
	if p_tactividad>"" then 'Se selecciono tipo de actividad
		strwhereall = strwhereall + " and p.tactividad='" & p_tactividad & "'"
	end if
	if p_proveedor>"" then 'Se selecciono proveedor
		strwhereall = strwhereall + " and f.nproveedor='" & p_proveedor & "'"
	else
		if p_opcproveedorbaja=1 then
			strwhereall = strwhereall + " and p.fbaja is null "
		else
			strbaja=""
		end if
   	end if

	if p_cod_proyecto>"" then
		strwhere=strwhere & " and cod_proyecto='" & p_cod_proyecto & "'"
		strwhere2=strwhere2 & " and cod_proyecto='" & p_cod_proyecto & "'"
	end if

 	if p_serie>"" then 'Se selecciono serie
	   strwhereall = strwhereall + " and f.serie='" & p_serie & "'"
	end if

	if tipo_proveedor & "">"" then
		strwhere = strwhere + " and p.tipo_proveedor='" & tipo_proveedor & "'"
		strwhere2 = strwhere2 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
	end if

	if tipo_articulo & "">"" then
		p_tsel2=false
		strwhere = strwhere + " and a.tipo_articulo='" & tipo_articulo & "'"
	end if
	if mostrarfilas & "">"" then
		strwhere=strwhere + ""
	end if

	strwhereall=strwhereall & " and f.nfactura like '" & session("ncliente") & "%'"

	seleccion1="select distinct d.referencia as Ref, "
	seleccion1=seleccion1 & "a.nombre as Descripcion, "
	seleccion1=seleccion1 & "f.nproveedor as nproveedor,"
	seleccion1=seleccion1 & "p.RAZON_SOCIAL as Nombre, "
	seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad,"

	seleccion1=seleccion1 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Compras Netas], "
	seleccion1=seleccion1 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

	seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto, "
	seleccion1=seleccion1 & "f.divisa as Divisa,"
	seleccion1=seleccion1 & "0 as tiene_escv "
	seleccion1=seleccion1 & " from detalles_fac_pro as d with(nolock) "
	seleccion1=seleccion1 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura "
	seleccion1=seleccion1 & " inner join articulos as a with(nolock) on a.referencia = d.referencia "
	seleccion1=seleccion1 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion1=seleccion1 & " where d.mainitem is null "
    'ega 22/07/2008 likes
    seleccion1=seleccion1 & " and d.nfactura like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and f.nfactura like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and a.referencia like '" & session("ncliente") & "%' "        	  
    seleccion1=seleccion1 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
	seleccion1=seleccion1 & strwhereall & strwhere & " "
	seleccion1=seleccion1 & " group by d.referencia, "
	seleccion1=seleccion1 & "a.nombre, "
	seleccion1=seleccion1 & "f.cod_proyecto, "
	seleccion1=seleccion1 & "f.nproveedor, "
	seleccion1=seleccion1 & "p.RAZON_SOCIAL, "
	seleccion1=seleccion1 & "f.divisa "
	seleccion1=seleccion1 & " union all "
	seleccion1=seleccion1 & " select distinct d.referencia as Ref, "
	seleccion1=seleccion1 & "a.nombre as Descripcion, "
	seleccion1=seleccion1 & "f.nproveedor as nproveedor,"
	seleccion1=seleccion1 & "p.RAZON_SOCIAL as Nombre, "
	seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad,"
	seleccion1=seleccion1 & "0 as [Compras Netas],"
	seleccion1=seleccion1 & "0 as [Precio Medio],"
	seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto, "
	seleccion1=seleccion1 & "f.divisa as Divisa,"
	seleccion1=seleccion1 & "1 as tiene_escv "
	seleccion1=seleccion1 & " from detalles_fac_pro as d with(nolock) "
	seleccion1=seleccion1 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura "
	seleccion1=seleccion1 & " inner join articulos as a with(nolock) on a.referencia = d.referencia "
	seleccion1=seleccion1 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
	seleccion1=seleccion1 & " where d.mainitem is not null "
	'ega 22/07/2008 likes
    seleccion1=seleccion1 & " and d.nfactura like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and f.nfactura like '" & session("ncliente") & "%' "
    seleccion1=seleccion1 & " and a.referencia like '" & session("ncliente") & "%' "        	  
    seleccion1=seleccion1 & " and p.nproveedor like '" & session("ncliente") & "%' "	  
	seleccion1=seleccion1 & strwhereall & strwhere & " "
	seleccion1=seleccion1 & " group by d.referencia, "
	seleccion1=seleccion1 & "a.nombre, "
	seleccion1=seleccion1 & "f.cod_proyecto, "
	seleccion1=seleccion1 & "f.nproveedor, "
	seleccion1=seleccion1 & "p.RAZON_SOCIAL, "
	seleccion1=seleccion1 & "f.divisa "

	cabecera = ""
	if p_desglose=false then
      	cabecera = "select distinct 'zzzzzzzzzzz' as Ref, " & _
            "'@concepto@' as Descripcion, "
	else

		''ricardo 18/2/2003
		''se pone esto, ya que puede ocurrir que haya una descripcion grandisima
		''cabecera = "select distinct d.descripcion as Ref, " & _
		cabecera="select distinct case when datalength(d.descripcion)<=20 then d.descripcion else left(d.descripcion,20)+'...' end as Ref, " & _
                                   "'@concepto@' as Descripcion, "
	end if

	seleccion2=cabecera & "f.nproveedor as nproveedor, "
    seleccion2=seleccion2 & "p.RAZON_SOCIAL as Nombre, "
    seleccion2=seleccion2 & "sum(d.cantidad) as Cantidad, "

	seleccion2=seleccion2 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Compras Netas], "
	seleccion2=seleccion2 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

    seleccion2=seleccion2 & "f.cod_proyecto as cod_proyecto,"
    seleccion2=seleccion2 & "f.divisa as Divisa,"
    seleccion2=seleccion2 & "0 as tiene_escv "
    seleccion2=seleccion2 & " from conceptos_fac_pro as d with(nolock) "
    seleccion2=seleccion2 & " inner join facturas_pro as f with(nolock) on d.nfactura = f.nfactura  "
    seleccion2=seleccion2 & " inner join proveedores as p with(nolock) on p.nproveedor = f.nproveedor "
    seleccion2=seleccion2 & " where "
    'ega 22/07/2008 likes
    seleccion2=seleccion2 & " d.nfactura like '" & session("ncliente") & "%' "
    seleccion2=seleccion2 & " and f.nfactura like '" & session("ncliente") & "%' "
    seleccion2=seleccion2 & " and p.nproveedor like '" & session("ncliente") & "%' "
    seleccion2=seleccion2 & strwhereall & strwhere2
    if desglose = false then
		seleccion2 = seleccion2 & " group by f.nproveedor, "
	    seleccion2 = seleccion2 & "p.RAZON_SOCIAL, "
      	seleccion2 = seleccion2 & "f.cod_proyecto, "
	    seleccion2 = seleccion2 & "f.divisa "
	else
		seleccion2 = seleccion2 & " group by d.descripcion, "
	      seleccion2=seleccion2 & "f.cod_proyecto, "
	      seleccion2=seleccion2 & "f.nproveedor, "
	      seleccion2=seleccion2 & "p.RAZON_SOCIAL, "
	      seleccion2=seleccion2 & "f.divisa "
	end if

    ''ega 18/07/2008 si se ha seleccionado alguna serie de albaran pendiente de facturar
	'if p_tsel3="on" or p_tsel3="1" then
	if seriesapf >"" then

		strwhereall3=""
		if p_dfecha>"" then
			strwhereall3=strwhereall3 & " and alb.fecha>='" & p_dfecha & "' "
		end if

		if p_hfecha>"" then
			strwhereall3=strwhereall3 & " and alb.fecha<='" & p_hfecha & "' "
		end if
		PorArticulo = "SI"
		if p_familia>"" then
			strwhere3 = strwhere3 + " and a.familia='" & p_familia & "' "
		end if
		if p_referencia>"" then
	      	strwhere3 = strwhere3 + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		end if
		if p_nombreart>"" then
			strwhere3 = strwhere3 + " and a.nombre like '%" & p_nombreart & "%' "
		end if
		if p_tactividad>"" then 'Se selecciono tipo de actividad
			strwhereall3 = strwhereall3 + " and p.tactividad='" & p_tactividad & "' "
		end if
		if p_proveedor>"" then 'Se selecciono proveedor
			strwhereall3 = strwhereall3 + " and alb.nproveedor='" & p_proveedor & "' "
		else
			if p_opcproveedorbaja=1 then
				strbaja=" "
			else
				strbaja=" and p.fbaja is null "
				strwhereall3 = strwhereall3 + strbaja
			end if
      	end if
		if p_cod_proyecto>"" then
			strwhere3=strwhere3 & " and cod_proyecto='" & p_cod_proyecto & "' "
			strwhere4=strwhere4 & " and cod_proyecto='" & p_cod_proyecto & "' "
		end if

		if tipo_proveedor & "">"" then
			strwhere3 = strwhere3 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
			strwhere4 = strwhere4 + " and p.tipo_proveedor='" & tipo_proveedor & "'"
		end if

		if tipo_articulo & "">"" then
			strwhere3 = strwhere3 + " and a.tipo_articulo='" & tipo_articulo & "'"
		end if
		if mostrarfilas & "">"" then
			strwhere3=strwhere3 + ""
		end if

		strwhereall3=strwhereall3 & " and alb.nalbaran like '" & session("ncliente") & "%'"

		seleccion3="select distinct d.referencia as Ref, "
		seleccion3=seleccion3 & "a.nombre as Descripcion, "
		seleccion3=seleccion3 & "alb.nproveedor as NProveedor,"
		seleccion3=seleccion3 & "p.razon_social as Nombre, "
		seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad,"

	    seleccion3=seleccion3 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Compras Netas], "
	    seleccion3=seleccion3 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

		seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
		seleccion3=seleccion3 & "alb.divisa as Divisa,"
		seleccion3=seleccion3 & "0 as tiene_escv "
		seleccion3=seleccion3 & " from detalles_alb_pro as d with(nolock)"
		seleccion3=seleccion3 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion3=seleccion3 & " inner join articulos as a with(nolock) on a.referencia = d.referencia "
		seleccion3=seleccion3 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion3=seleccion3 & " where alb.nfactura is null "
		'ega 22/07/2008 likes
        seleccion3=seleccion3 & " and d.nalbaran like '" & session("ncliente") & "%' "
        seleccion3=seleccion3 & " and a.referencia like '" & session("ncliente") & "%' "
        seleccion3=seleccion3 & " and p.nproveedor like '" & session("ncliente") & "%' "
		seleccion3=seleccion3 & strwhereall3 & strwhere3 & " "
		seleccion3=seleccion3 & " and d.mainitem is null "
		seleccion3=seleccion3 & " group by d.referencia, "
		seleccion3=seleccion3 & "a.nombre, "
		seleccion3=seleccion3 & "alb.cod_proyecto, "
		seleccion3=seleccion3 & "alb.nproveedor, "
		seleccion3=seleccion3 & "p.razon_social, "
		seleccion3=seleccion3 & "alb.divisa "
		seleccion3=seleccion3 & " union all "
		seleccion3=seleccion3 & "select distinct d.referencia as Ref, "
		seleccion3=seleccion3 & "a.nombre as Descripcion, "
		seleccion3=seleccion3 & "alb.nproveedor as NProveedor,"
		seleccion3=seleccion3 & "p.razon_social as Nombre, "
		seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad,"
		seleccion3=seleccion3 & "0 as [Compras Netas],"
		seleccion3=seleccion3 & "0 as [Precio Medio],"
		seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
		seleccion3=seleccion3 & "alb.divisa as Divisa,"
		seleccion3=seleccion3 & "1 as tiene_escv "
		seleccion3=seleccion3 & " from detalles_alb_pro as d with(nolock) "
		seleccion3=seleccion3 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion3=seleccion3 & " inner join articulos as a with(nolock) on a.referencia = d.referencia "
		seleccion3=seleccion3 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion3=seleccion3 & " where alb.nfactura is null "
		'ega 22/07/2008 likes
        seleccion3=seleccion3 & " and d.nalbaran like '" & session("ncliente") & "%' "
        seleccion3=seleccion3 & " and a.referencia like '" & session("ncliente") & "%' "
        seleccion3=seleccion3 & " and p.nproveedor like '" & session("ncliente") & "%' "        
		seleccion3=seleccion3 & strwhereall3 & strwhere3 & " "
		seleccion3=seleccion3 & " and d.mainitem is not null "
		seleccion3=seleccion3 & " group by d.referencia, "
		seleccion3=seleccion3 & "a.nombre, "
		seleccion3=seleccion3 & "alb.cod_proyecto, "
		seleccion3=seleccion3 & "alb.nproveedor, "
		seleccion3=seleccion3 & "p.razon_social, "
		seleccion3=seleccion3 & "alb.divisa "
 
		cabecera = ""
		if p_desglose=false then
      		cabecera4 = "select distinct 'zzzzzzzzzzz' as Ref, " & _
                                  "'@concepto@' as Descripcion, "
		else
			''ricardo 18/2/2003
			''se pone esto, ya que puede ocurrir que haya una descripcion grandisima
			cabecera4 = "select distinct convert(nvarchar,d.descripcion) as Ref, "
			''cabecera4="select case when datalength(d.descripcion)<=20 then d.descripcion else left(d.descripcion,20)+'...' end as Ref, "
			cabecera4=cabecera4 & "'@concepto@' as Descripcion, "
		end if

		seleccion4=cabecera4 & "alb.nproveedor as NProveedor, "
		seleccion4=seleccion4 & "p.razon_social as Nombre, "
		seleccion4=seleccion4 & "sum(d.cantidad) as Cantidad, "

		seleccion4=seleccion4 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Compras Netas], "
		seleccion4=seleccion4 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "

		seleccion4=seleccion4 & "alb.cod_proyecto as cod_proyecto,"
		seleccion4=seleccion4 & "alb.divisa as Divisa,"
		seleccion4=seleccion4 & "0 as tiene_escv "
		seleccion4=seleccion4 & " from conceptos_alb_pro as d with(nolock) "
		seleccion4=seleccion4 & " inner join albaranes_pro as alb with(nolock) on d.nalbaran = alb.nalbaran "
		seleccion4=seleccion4 & " inner join proveedores as p with(nolock) on p.nproveedor = alb.nproveedor "
		seleccion4=seleccion4 & " where alb.nfactura is null "
		'ega 22/07/2008 likes
        seleccion4=seleccion4 & " and d.nalbaran like '" & session("ncliente") & "%' "
        seleccion4=seleccion4 & " and p.nproveedor like '" & session("ncliente") & "%' "           
		seleccion4=seleccion4 & strwhereall3 & strwhere4
		if desglose = false then
			seleccion4 = seleccion4 & " group by alb.nproveedor, "
			seleccion4=seleccion4 & "p.razon_social, "
			seleccion4=seleccion4 & "alb.cod_proyecto, "
			seleccion4=seleccion4 & "alb.divisa "
		else
			seleccion4 = seleccion4 & " group by convert(nvarchar,d.descripcion), "
			seleccion4=seleccion4 & "alb.cod_proyecto, "
			seleccion4=seleccion4 & "alb.nproveedor, "
			seleccion4=seleccion4 & "p.razon_social, "
			seleccion4=seleccion4 & "alb.divisa "
		end if
	end if

	seleccion = ""
	if p_tsel1=true or p_tsel1="1" then
		seleccion = seleccion & seleccion1
	end if
	if (p_tsel1=true or p_tsel1="1") and (p_tsel2=true or p_tsel2="1") then
		seleccion = seleccion & " union ALL "
	end if
	if p_tsel2 = true or p_tsel2 = "1" then
		seleccion = seleccion & seleccion2
	end if

    ''ega 18/07/2008 si se ha seleccionado alguna serie de albaran pendiente de facturar
	''if p_tsel3="on" or p_tsel3="1" then
	if seriesapf >"" then	
		if p_tsel1=true or p_tsel1="1" then
			seleccion=seleccion & " union ALL " & seleccion3
		end if
		if p_tsel2 = true or p_tsel2="1" then
			seleccion=seleccion & " union ALL " & seleccion4
		end if
	end if

	porcompras = ""
	'if ordenar=true then porcompras = "[Compras Netas] desc, "

	if p_tsel1=true then
	   seleccion = seleccion & " order by " & porcompras & "d.referencia, nproveedor"
	else
	   if p_desglose = false then
 	      seleccion = seleccion & " order by " & porcompras & "nproveedor"
       else
		   seleccion = seleccion & " order by " & porcompras & "descripcion, nproveedor"
	   end if
	end if

	rstFunction.open seleccion, session("backendlistados"),1,3

	if not rstFunction.eof then
      	acumuladoPasta    = 0
		acumuladoCantidad = 0
		elTotal = 0
		rstTemp.open "select * from [" & session("usuario") & "]", session("backendlistados"), adOpenKeyset, adLockOptimistic
		articuloAnterior = ""
		while not rstFunction.eof
			if ucase(rstFunction("Ref"))<>ucase(articuloAnterior) then
                ''ricardo 17/12/2003
			    '''elTotal = elTotal + formatnumber(acumuladoPasta,n_decimalesMB,-1,0,-1)
			    elTotal = elTotal + acumuladoPasta
                '''''''''
	      		acumuladoPasta    = 0
				acumuladoCantidad = 0
			end if
			articuloAnterior = ucase(rstFunction("Ref"))
			rstTemp.Addnew
			if len(rstFunction("Ref"))>20 then
				rstTemp("Ref") =mid(rstFunction("Ref"),1,20) & "..."
			else
				rstTemp("Ref") =rstFunction("Ref")
			end if
			rstTemp("Descripcion") = rstFunction("Descripcion")
			rstTemp("nproveedor") = rstFunction("nproveedor")
			rstTemp("Nombre") = rstFunction("Nombre")
			rstTemp("Cantidad") = rstFunction("Cantidad")
			rstTemp("Compras Netas") = rstFunction("Compras Netas")
			rstTemp("Precio Medio") = rstFunction("Precio Medio")
			rstTemp("cod_proyecto") = d_lookup("nombre","proyectos","codigo='" & rstFunction("cod_proyecto") & "'",session("backendlistados"))
			rstTemp("Divisa") = rstFunction("Divisa")
			if rstTemp("Divisa") = p_mb then
				rstTemp("AcumulaCompras") = acumuladoPasta + rstTemp("Compras Netas")
				acumuladoPasta = rstTemp("AcumulaCompras")
			else
				rstTemp("AcumulaCompras") = acumuladoPasta + CambioDivisa(rstTemp("Compras Netas"), rstTemp("Divisa"), p_mb)
				acumuladoPasta = rstTemp("AcumulaCompras")
			end if
			acumuladoCantidad = acumuladoCantidad + rstFunction("Cantidad")
			rstTemp("AcumulaCantidad") = acumuladoCantidad
			rstTemp("tiene_escv")=rstFunction("tiene_escv")
			rstTemp("Orden") = 0
			rstTemp.Update
			rstFunction.movenext
		wend
		rstFunction.close
		rstTemp.close

		'ordenamos por compras si es necesario
		if ordenar=true then
 	      	rstFunction.open "select distinct ref from [" & session("usuario") & "]", session("backendlistados"),1,3
		      while not rstFunction.eof
				rstTemp.open "update [" & session("usuario") & "] set orden = (select max(acumulacompras) from [" & _
			              session("usuario") & "] where ref = '" & rstFunction("Ref") & _
						  "') where Ref ='" & rstFunction("Ref") & "'", session("backendlistados"), 1, 3
				rstFunction.movenext
			wend
			rstFunction.close
		end if
	end if

''ricardo 17/12/2003 el total se calcula de esta manera
rstTemp.open "select sum([Compras Netas]/d.factcambio) as total from [" & session("usuario") & "] as temp1,divisas as d with(nolock) where d.codigo=temp1.divisa", session("backendlistados"),1,3
if not rstTemp.eof then
	'''elTotal=rstTemp("total")
	elTotal=formatnumber(null_z(rstTemp("total")),n_decimalesMB,-1,0,-1)
'''''''''''''''''''''''''''''
end if
rstTemp.close

''ricardo 17/12/2003 el total se calcula de esta manera
'''	elTotal = elTotal + formatnumber(null_z(acumuladoPasta),n_decimalesMB,-1,0,-1)
	%><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>'><%

end sub


'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
 %>

<form name="resumen_compras_pro" method="post">
<%
		PintarCabecera "resumen_compras_pro.asp"
		'Leer parámetros de la página
		SumaAnt = 0
		SubTotalAnt = 0
  		mode=enc.EncodeForJavascript(Request.QueryString("mode"))
  		campo=limpiaCadena(Request.QueryString("campo"))
  		criterio=limpiaCadena(Request.QueryString("criterio"))
  		texto=limpiaCadena(Request.QueryString("texto"))
		elTotal=limpiaCadena(Request.form("elTotal"))

        seriesapf=cstr(Request.Form("seriesapf"))

        lista_series ="('"
        lista_series2 ="(''"

        if seriesapf > "" then
	        if instr(seriesapf,",")>0 then
		        lista_series = lista_series & replace(replace(seriesapf," ",""),",","','")
		        lista_series2 = lista_series2 & replace(replace(seriesapf," ",""),",","'',''")

	        else
		        lista_series = lista_series & seriesapf
		        lista_series2 = lista_series2 & seriesapf
	        end if
        end if
        if right(lista_series,3) <> "','" then
	        lista_series= lista_series & "','"
        end if
        if right(lista_series2,5) <> "'',''" then
	        lista_series2= lista_series2 & "'',''"
        end if
        lista_series = lista_series +"WvWvW')"
        lista_series2 = lista_series2 +"WvWvW'')"
        
		si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
        si_tiene_modulo_ccostes=ModuloContratado(session("ncliente"),ModCcostes_Gestion)

  		set rstAux = 	Server.CreateObject("ADODB.Recordset")
		set rst = 		Server.CreateObject("ADODB.Recordset")
  		set rst2 = 		Server.CreateObject("ADODB.Recordset")
  		set rstSelect = Server.CreateObject("ADODB.Recordset")
  		set rstTablas = Server.CreateObject("ADODB.Recordset")
%>

	<%if mode="browse" then%>
		<table width='100%'>
		   	<tr>
				<td width="30%" align="left">
					<font class=CELDAL7>&nbsp;(<%=LitEmitido%>&nbsp;<%=day(date)%>/<%=month(date)%>/<%=year(date)%>)</font>
				</td>
				<td>
					<font class='CABECERA'><b></b></font>
					<font class=CELDA><b></b></font>
				</td>
				<td></td>
			</tr>
		</table>
		<hr/>
	<%end if
	Alarma "resumen_compras_pro.asp"

'****************************************************************************************************************
	'Leer parámetros de la página
	mode		= enc.EncodeForJavascript(Request.QueryString("mode"))                                      
	nproveedor	= limpiaCadena(Request.QueryString("nproveedor"))
	if nproveedor ="" then
		nproveedor	= limpiaCadena(Request.form("nproveedor"))
	end if

	if nproveedor & "">"" and mode="select1" then
		nproveedor=session("ncliente") & Completar(nproveedor,5,"0")
		'miramos si existe el proveedor
		rst.open "select nproveedor from proveedores with(nolock) where nproveedor='" & nproveedor & "'",session("backendlistados"),adUseClient,adLockReadOnly
		if rst.eof then
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgProveedorNoExiste%>");
			</script><%
			nproveedor=""
		end if
		rst.close
	end if

	actividad	= limpiaCadena(Request.QueryString("actividad"))
	if actividad ="" then
		actividad	= limpiaCadena(Request.form("actividad"))
	end if

	fdesde		= limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde	= limpiaCadena(Request.form("fdesde"))
	end if

	fhasta		= limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta	= limpiaCadena(Request.form("fhasta"))
	end if

	nserie	= limpiaCadena(Request.QueryString("nserie"))
	if nserie ="" then
		nserie	= limpiaCadena(Request.form("nserie"))
	end if

	falbaran	= limpiaCadena(Request.QueryString("fechaalbaran"))
	if falbaran ="" then
		falbaran	= limpiaCadena(Request.form("fechaalbaran"))
	end if

	tactividad	= limpiaCadena(Request.QueryString("tactividad"))
	if tactividad ="" then
		tactividad	= limpiaCadena(Request.form("tactividad"))
	end if

	referencia	= limpiaCadena(Request.QueryString("referencia"))
	if referencia ="" then
		referencia	= limpiaCadena(Request.form("referencia"))
	end if
	
	if referencia & "">"" and mode="browse" then
		'miramos si existe la referencia
		rst.cursorlocation=3
		rst.open "select referencia from articulos with(nolock) where referencia like '" & session("ncliente") & "%" & referencia & "%'",session("backendlistados")
		if rst.eof then
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgArticuloconRefNoExiste%>");
				parent.pantalla.resumen_compras_pro.action="resumen_compras_pro.asp?mode=select1";
				parent.pantalla.resumen_compras_pro.submit();
				parent.botones.document.location="resumen_compras_pro_bt.asp?mode=select1"
			</script><%
			referencia=""
		end if
		rst.close
	end if

	nombreart	= limpiaCadena(Request.QueryString("nombreart"))
	if nombreart ="" then
		nombreart	= limpiaCadena(Request.form("nombreart"))
	end if

	if nombreart & "">"" and mode="browse" then
		'miramos si existe algun articulo con ese nombreart
		rst.open "select referencia from articulos with(nolock) where referencia like '" & session("ncliente") & "%' and nombre like '%" & nombreart & "%'",session("backendlistados"),adUseClient,adLockReadOnly
		if rst.eof then
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgArticuloConNombreNoExiste%>");
				parent.pantalla.resumen_compras_pro.action="resumen_compras_pro.asp?mode=select1";
				parent.pantalla.resumen_compras_pro.submit();
				parent.botones.document.location="resumen_compras_pro_bt.asp?mode=select1"
			</script><%
			nombreart=""
		end if
		rst.close
	end if

	familia	= limpiaCadena(Request.QueryString("familia"))
	if familia ="" then
		familia	= limpiaCadena(Request.form("familia"))
	end if
	agrupar	= limpiaCadena(Request.QueryString("agrupar"))
	if agrupar ="" then
		agrupar	= limpiaCadena(Request.form("agrupar"))
	end if
	conceptos	= limpiaCadena(Request.QueryString("conceptos"))
	if conceptos ="" then
		conceptos	= limpiaCadena(Request.form("conceptos"))
	end if

	if conceptos>"" then conceptos="1"
	ver_conceptos	= limpiaCadena(Request.QueryString("ver_conceptos"))
	if ver_conceptos ="" then
		ver_conceptos	= limpiaCadena(Request.form("ver_conceptos"))
	end if

	if ver_conceptos>"" then ver_conceptos="1"

  	set_orden	= limpiaCadena(Request.QueryString("ordenar_compras"))
	if set_orden ="" then
		set_orden	= limpiaCadena(Request.form("ordenar_compras"))
	end if

	if set_orden>"" then set_orden="1"

	if request.form("opcproveedorbaja")>"" then
		opcproveedorbaja=limpiaCadena(request.form("opcproveedorbaja"))
	else
		opcproveedorbaja=limpiaCadena(request.querystring("opcproveedorbaja"))
	end if

	if opcproveedorbaja>"" then opcproveedorbaja="1"

	cod_proyecto	= limpiaCadena(Request.QueryString("cod_proyecto"))
	if cod_proyecto="" then
		cod_proyecto	= limpiaCadena(Request.form("cod_proyecto"))
	end if

	if request.form("opc_cod_proyecto")>"" then
		opc_cod_proyecto="1"
	end if

	if request.form("opc_coste")>"" then
		opc_coste="1"
	end if	

	if request.form("seriesapf")>"" then
		seriesapf=limpiaCadena(request.form("seriesapf"))
	else
		seriesapf=limpiaCadena(request.querystring("seriesapf"))
	end if

	prohojassep=limpiaCadena(request.form("prohojassep"))
	if prohojassep="on" then prohojassep="1"

	apaisado=iif(limpiaCadena(request.form("apaisado"))>"","SI","")

	tipo_proveedor=limpiaCadena(Request.QueryString("tipo_proveedor"))
	if tipo_proveedor="" then
		tipo_proveedor= limpiaCadena(Request.form("tipo_proveedor"))
	end if

	tipo_articulo=limpiaCadena(Request.QueryString("tipo_articulo"))
	if tipo_articulo="" then
		tipo_articulo= limpiaCadena(Request.form("tipo_articulo"))
	end if

	mostrarfilas=limpiaCadena(Request.QueryString("mostrarfilas"))
	if mostrarfilas="" then
		mostrarfilas= limpiaCadena(Request.form("mostrarfilas"))
	end if

	opc_cantidad=limpiaCadena(request.form("opc_cantidad"))
	opc_comprasnetas=limpiaCadena(request.form("opc_comprasnetas"))

	WaitBoxOculto LitEsperePorFavor

	if (mode="select1") then%>
    <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitParEstComFac%></h6>
		<%
            EligeCelda "input", "add", "", "", "", 0, LitDesdeFecha, "fdesde", "", iif(fdesde>"",fdesde,"01/01/" & year(date))
            DrawCalendar "fdesde"
            EligeCelda "input", "add", "", "", "", 0, LitHastaFecha, "fhasta", "", iif(fhasta>"",fhasta,day(date) & "/" & month(date) & "/" & year(date))
            DrawCalendar "fhasta"
               
            rstSelect.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='FACTURA DE PROVEEDOR' and nserie like '" & session("ncliente") & "%'  order by nserie",session("backendlistados"),adUseClient,adLockReadOnly
            DrawSelectCelda "CELDA","175","",0,LitSerie,"nserie",rstSelect,nserie,"nserie","descripcion","",""       
            rstSelect.close  
            
            EligeCelda "check", "add", "", "", "", 0, LitFechaAlbaran, "fechaalbaran", "", ""
        %>                                                                                                                                                                                                                                              
     <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitParEstComPro%></h6>                         
        <%
            DrawDiv "1", "", ""
                DrawLabel "", "", LitProveedor%><input class='width15' type="text" name="nproveedor" value="<%=EncodeForHtml(trimCodEmpresa(nproveedor))%>" size = 10 onchange="TraerProveedor('<%=enc.EncodeForJavascript(null_s(mode))%>','<%=enc.EncodeForJavascript(null_s(ndet))%>');"><a class='CELDAREFB' href="javascript:AbrirVentana('../proveedores_busqueda.asp?ndoc=resumen_compras_pro&titulo=<%=LitSelProv%>&mode=search&viene=resumen_compras_pro','P','<%=AltoVentana%>','<%=AnchoVentana%>')" OnMouseOver="self.status='<%=LitVerProveedor%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' disabled type="text" name="nombre" value="<%=EncodeForHtml(iif(nproveedor>"",d_lookup("RAZON_SOCIAL","proveedores","nproveedor='" & nproveedor & "'",session("backendlistados")),""))%>" size="25" /><%
			CloseDiv
			
            rstSelect.open "select codigo, case when datalength(substring(codigo,6,10)+' '+descripcion)<=21 then substring(codigo,6,10)+'-'+descripcion else left(substring(codigo,6,10)+'-'+descripcion,20)+'...' end as descripcion from tipo_actividad with(nolock) where codigo like '" & session("ncliente") & "%' order by codigo",session("backendlistados"),adUseClient,adLockReadOnly
			DrawSelectCelda "CELDA","175","",0,LitActividad,"tactividad",rstSelect,tactividad,"codigo","descripcion","",""
			rstSelect.close
			
            rstSelect.open "select codigo,descripcion from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo ='" & LitProveedor & "' order by descripcion",session("backendlistados"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "CELDA",175,"",0,LitTipoProveedor,"tipo_proveedor",rstSelect,tipo_proveedor,"codigo","descripcion","",""
			rstSelect.close
			
            DrawDiv "1", "", ""
                DrawLabel "", "", LitProveedorBaja%><input type="checkbox" name="opcproveedorbaja" <%=iif(opcproveedorbaja="1","checked","")%>>
			<%
			CloseDiv                                                                                                                 
			if si_tiene_modulo_proyectos<>0 then
                %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><input class="CELDA" type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(cod_proyecto)%>" /><label><%=LitProyecto%></label><%
				%><iframe class="iframe-menu width60" id='frProyecto' src='../../mantenimiento/docproyectos_responsive.asp?viene=resumen_compras_pro&mode=<%=EncodeForHtml(mode)%>&cod_proyecto=<%=EncodeForHtml(cod_proyecto)%>' frameborder="no" scrolling="no" noresize="noresize"></iframe></div><%
			end if%>
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitParEstVenArt%></h6><%                                                                                                         
			                                                                                                                                                          
            EligeCelda "input", "add", "", "", "", 0, LitConrefRC, "referencia", "", referencia
            EligeCelda "input", "add", "", "", "", 0, LitConNombreRC, "nombreart", "", nombreart
			
			rstAux.open " select codigo, nombre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre", session("backendlistados"),adUseClient,adLockReadOnly
			DrawSelectCelda "CELDA","175","",0,LitSubFamilia,"familia",rstAux,familia,"codigo","nombre","",""
			rstAux.close

			rstSelect.open " select codigo, descripcion from tipos_entidades with(nolock) where tipo='ARTICULO' and codigo like '" & session("ncliente") & "%' order by descripcion", session("backendlistados"),adOpenKeyset,adLockOptimistic
	       	DrawSelectCelda "CELDA","175","",0,LitTipoArt,"tipo_articulo",rstSelect,tipo_articulo,"codigo","descripcion","",""
			rstSelect.close
			%>
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitParEstVenOpc%></h6><%

            DrawDiv "1", "", ""
                DrawLabel "", "", LitAgrTpProveedor%><select class='width60' name="agrupar" onchange="javascript:optionalFields();">
						<option <%=iif(agrupar=ucase(LitArticulo),"selected","")%> value="<%=ucase(LitArticulo)%>"><%=ucase(LitArticulo)%></option>
						<option <%=iif(agrupar=LitMeses,"selected","")%> value="<%=LitMeses%>"><%=LitMeses%></option>
						<option <%=iif(agrupar=ucase(LitProveedor) or agrupar="","selected","")%> value="<%=ucase(LitProveedor)%>"><%=ucase(LitProveedor)%></option>
						<%if si_tiene_modulo_proyectos<>0 then%>
							<option <%=iif(agrupar=ucase(LitProyecto),"selected","")%> value="<%=ucase(LitProyecto)%>"><%=ucase(LitProyecto)%></option>
						<%end if%>
					</select>
            <%CloseDiv%>
    <span id="agrMeses3" style="display:none"><span id="agrMeses2" style="display:none"><%
			   DrawDiv "1", "", ""
                    DrawLabel "", "", LitRedMosFil%><select class='width60'  name="mostrarfilas" onchange="javascript:optionalFields();">
						<option <%=iif(mostrarfilas=ucase(LitArticulos),"selected","")%> value="<%=ucase(LitArticulos)%>"><%=ucase(LitArticulos)%></option>
						<option <%=iif(mostrarfilas=LitProveedores or mostrarfilas="","selected","")%> value="<%=LitProveedores%>"><%=LitProveedores%></option>
					</select>
				<%
			CloseDiv%></span></span><%
			
            rstSelect.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as nombre from series with(nolock) where tipo_documento ='ALBARAN DE PROVEEDOR' and nserie like '" & session("ncliente") & "%' order by nserie",session("backendlistados"),adUseClient,adLockReadOnly
            DrawSelectMultipleCelda "'CELDA' align='left' size='5' multiple","200","",0,LitAlbPendFac,"seriesapf",rstSelect,"","nserie","nombre","",""
			rstSelect.close
			
            EligeCelda "check", "add", "", "", "", 0, LitOrdenarCompras, "ordenar_compras", "", iif(set_orden="on" or set_orden="true" or set_orden="1","True","")
            DrawDiv "1", "", ""
                DrawLabel "", "", LitApaisado1%><input type="checkbox" name="apaisado" <%=iif(apaisado="SI" or apaisado="on" or apaisado="true" or apaisado="1","checked","")%> />
			<%CloseDiv
            DrawDiv "1", "", ""
                DrawLabel "", "", LitMostrarConceptos%><input type="checkbox" name="ver_conceptos" <%=iif(ver_conceptos="on" or ver_conceptos="true" or ver_conceptos="1","checked","")%> onclick="javascript:Ver_Conceptos();" /><%
			CloseDiv
            DrawDiv "1", "", "agrProHojSep"
               DrawLabel "", "", LitProveedorPorPagina%><input type="checkbox" name="prohojassep" <%=iif(prohojassep="on" or prohojassep="true" or prohojassep="1","checked","")%> /><%
            CloseDiv%>
            <span id="agrProHojSep2" style="display:none"><%
            EligeCelda "check", "add", "", "", "", 0, LitDesglosarCptos, "conceptos", "", iif(conceptos="on" or conceptos="true" or conceptos="1","True","")
            %></span>
		<%if si_tiene_modulo_proyectos<>0 or agrupar=LitMeses then
			no_mostrar_tit_camp_op=""
		else
			no_mostrar_tit_camp_op="none"
		end if%>
        <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitCamposOpcionales%></h6><%
			if si_tiene_modulo_proyectos<>0 then
			    EligeCelda "check", "add", "", "", "", 0, LitProyecto, "opc_cod_proyecto", "", 	iif(opc_cod_proyecto="on" or opc_cod_proyecto="true" or opc_cod_proyecto="1","True","")
			end if%>
            <div id="agrMeses" style="display: none">
            <%DrawDiv "1", "", ""
                DrawLabel "", "", LitCantidad%><input type="checkbox" name="opc_cantidad" <%=iif(opc_cantidad="on" or opc_cantidad="true" or opc_cantidad="1" or opc_comprasnetas="","checked","")%> onclick="javascript:if(document.resumen_compras_pro.opc_cantidad.checked) document.resumen_compras_pro.opc_comprasnetas.checked=false;else document.resumen_compras_pro.opc_comprasnetas.checked=true;" /><%
			 CloseDiv
             DrawDiv "1", "", ""
                DrawLabel "", "", LitComprasNetas%><input type="checkbox" name="opc_comprasnetas" <%=iif(opc_comprasnetas="on" or opc_comprasnetas="true" or opc_comprasnetas="1","checked","")%> onclick="javascript:BloquearVentasNetas();" /><%
             CloseDiv
                    %>
            </div>
            <div id="agrProveedor" style="display: ">
			<% DrawDiv "1", "", ""
                    DrawLabel "", "", LitCoste%><input type="checkbox" name="opc_coste" <%=iif(opc_coste="on" or opc_coste="true" or opc_coste="1" ,"checked","")%> />
            <% CloseDiv%>
            </div><%
		if agrupar & "">"" then%>
			<script language="javascript" type="text/javascript">
				optionalFields();
				Ver_Conceptos()
			</script>
        <%end if
'****************************************************************************************************************
		'Mostrar el listado.
	elseif mode="browse" then%>
		    <input type="hidden" name="fdesde" value="<%=EncodeForHtml(fdesde)%>" />
		    <input type="hidden" name="fhasta" value="<%=EncodeForHtml(fhasta)%>" />
		    <input type="hidden" name="nserie" value="<%=EncodeForHtml(nserie)%>" />
		    <input type="hidden" name="nproveedor" value="<%=EncodeForHtml(nproveedor)%>" />
		    <input type="hidden" name="actividad" value="<%=EncodeForHtml(actividad)%>" />
	  	    <input type="hidden" name="tactividad" value="<%=EncodeForHtml(tactividad)%>" />
		    <input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>" />
			<input type="hidden" name="nombreart" value="<%=EncodeForHtml(nombreart)%>" />
			<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>" />
			<input type="hidden" name="agrupar" value="<%=EncodeForHtml(agrupar)%>" />
			<input type="hidden" name="conceptos" value="<%=EncodeForHtml(conceptos)%>" />
			<input type="hidden" name="ver_conceptos" value="<%=EncodeForHtml(ver_conceptos)%>" />
			<input type="hidden" name="ordenar_compras" value="<%=EncodeForHtml(set_orden)%>" />
			<input type="hidden" name="opcproveedorbaja" value="<%=EncodeForHtml(opcproveedorbaja)%>" />
			<input type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(cod_proyecto)%>" />
			<input type="hidden" name="opc_cod_proyecto" value="<%=EncodeForHtml(opc_cod_proyecto)%>" />
			<input type="hidden" name="seriesapf" value="<%=EncodeForHtml(seriesapf)%>" />
			<input type="hidden" name="prohojassep" value="<%=EncodeForHtml(prohojassep)%>" />
			<input type="hidden" name="opc_cantidad" value="<%=EncodeForHtml(opc_cantidad)%>" />
			<input type="hidden" name="opc_coste" value="<%=EncodeForHtml(opc_coste)%>" />
			<input type="hidden" name="opc_comprasnetas" value="<%=EncodeForHtml(opc_comprasnetas)%>" />
			<input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>" />
			<input type="hidden" name="tipo_proveedor" value="<%=EncodeForHtml(tipo_proveedor)%>" />
			<input type="hidden" name="tipo_articulo" value="<%=EncodeForHtml(tipo_articulo)%>" />
			<input type="hidden" name="mostrarfilas" value="<%=EncodeForHtml(mostrarfilas)%>" />
<%
			MB=d_lookup("codigo", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("backendlistados"))
			n_decimales = null_z(d_lookup("ndecimales", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("backendlistados")))
			n_decimalesMB = n_decimales
			MB_abrev = d_lookup("abreviatura", "divisas", "codigo='" & MB & "'", session("backendlistados"))
		    MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='108'", DSNIlion)
		    MAXPDF=d_lookup("maxpdf", "limites_listados", "item='108'", DSNIlion)
%>
			<input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>' /><%

			VinculosPagina(MostrarProveedores)=1:VinculosPagina(MostrarArticulos)=1
			CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

			PorProveedor="NO"
			PorGasto="NO"

			tsel1 = true
			tsel2 = true

			prohojassep=request.form("prohojassep")
			if prohojassep="on" then prohojassep="1"

			if conceptos="true" or conceptos = "on" or conceptos="1" then
			   desglose = true
			else
			   desglose = false
			end if

			if ver_conceptos="true" or ver_conceptos = "on" or ver_conceptos = "1" then tsel1=false

			if set_orden="true" or set_orden = "on" or set_orden = "1" then
			   ordenar = true
			else
			   ordenar = false
			end if

			if fdesde>"" then
				%><font class=ENCABEZADO><b><%=LitDesdeFecha%> : </b></font><font class='CELDA'><%=EncodeForHtml(fdesde)%></b></font><br/><%
			end if
			if fhasta>"" then
				%><font class=ENCABEZADO><b><%=LitHastaFecha%> : </b></font><font class='CELDA'><%=EncodeForHtml(fhasta)%></b></font><br/><%
			end if                                                                                                                                                                     

			if tipo_proveedor& "">"" then
				desc_tipo_proveedor= d_lookup("descripcion", "tipos_entidades", "codigo='" & tipo_proveedor & "'", session("backendlistados"))
				%><font class=ENCABEZADO><b><%=LitTipoProveedor%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tipo_proveedor))%>&nbsp;&nbsp;<%=EncodeForHtml(trimCodEmpresa(desc_tipo_proveedor))%></b></font><br/><%
			end if                                                                                 
			if tipo_articulo& "">"" then
				desc_tipo_articulo = d_lookup("descripcion", "tipos_entidades", "codigo='" & tipo_articulo & "'", session("backendlistados"))
				%><font class=ENCABEZADO><b><%=LitTipoArt%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tipo_articulo))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tipo_articulo)%></b></font><br/><%
			end if                                                                                                                                                             
			if mostrarfilas & "">"" and agrupar=LitMeses then
				%><font class=ENCABEZADO><b><%=LitRedMosFil%> : </b></font><font class='CELDA'><%=enc.EncodeForHtmlAttribute(mostrarfilas)%></b></font><br/><%
			end if                                                                                      
			if agrupar=ucase(LitProveedor) then                                                                   
				PorProveedor = "SI"
				if tactividad>"" then 'Se selecciono tipo de actividad
			      	desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
				end if
				if nproveedor>"" then 'Se selecciono proveedor
					nproveedor=session("ncliente") & nproveedor
				      PorProveedor="NO"
					nomcli = d_lookup("RAZON_SOCIAL","proveedores","nproveedor='" & nproveedor & "'",session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitProveedor%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nproveedor))%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><br/><%
				else                                                                                           
					if opcproveedorbaja="" then
						t_opcproveedorbaja=0
					else
						t_opcproveedorbaja=1
					      %><font class=ENCABEZADO><b><%=LitProveedorBaja%></b></font><br/><%
					end if                                                                                
				end if
				if nserie>"" then 'Se selecciono serie
					%><font class=ENCABEZADO><b><%=LitSerie%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nserie))%>&nbsp;&nbsp;<%=EncodeForHtml(d_lookup("nombre","series","nserie='" & nserie & "'",session("backendlistados")))%></b></font><br/><%
				end if
				if familia>"" then                                                                                                                                      
					tsel2 = false
					desc_familia = d_lookup("nombre", "familias", "codigo='" & familia & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitSubFamilia%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(familia))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_familia)%></b></font><br/><%
				end if
				if cod_proyecto>"" then
					%><font class=ENCABEZADO><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",session("backendlistados")))%></font><br/><%
				end if
				if nombreart>"" then 'Se selecciono nombreart
					%><font class=ENCABEZADO><b><%=LitConNombreRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(nombreart)%></b></font><br/><%       
				end if
				if referencia>"" then 'Se selecciono referencia                                                 
					%><font class=ENCABEZADO><b><%=LitConrefRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(referencia)%></b></font><br/><%
				end if
				'if incalbpendfac="on" or incalbpendfac="1" then
				if seriesapf>"" then
					%><font class=ENCABEZADO><b><%=LitIncAlbPenFact%></b></font><br/><%
				end if
				if request.querystring("save")="true" then
					crearProveedor tsel1, tsel2,fdesde, fhasta, tactividad, nproveedor, nserie, familia, referencia, nombreart, ordenar, MB,t_opcproveedorbaja,cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas
				else
					%><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>'><%              
				end if
				%><hr/>
<%
				'if incalbpendfac="on" or incalbpendfac="1" then
				if seriesapf>"" then
					seleccion="SELECT NProveedor, Nombre, Referencia,Descripcion,"
					seleccion=seleccion & "SUM(cantidad) AS Cantidad, SUM([Compras Netas]) AS [Compras Netas],"
					seleccion=seleccion & "SUM([Precio Medio] * cantidad)/ CASE WHEN SUM(cantidad)= 0 THEN 1 ELSE SUM(cantidad) END AS [Precio Medio],"
					if opc_cod_proyecto="1" then
						seleccion=seleccion & "cod_proyecto,"
					end if
					
					seleccion=seleccion & "Divisa,(select sum([Compras Netas]) from [" & session("usuario") & "] WHERE nproveedor = f.NProveedor) as Acumulador,Orden"
					seleccion=seleccion & ",tiene_escv "
					''ricardo 28-8-2007 faltaba el campo cantidad2,medidaventa y preciomedio2
					seleccion=seleccion & ",round(sum(cantidad2)," & n_decimales & ") as cantidad2,medidaVenta,round(sum(precioMedio2)," & n_decimales & ") as precioMedio2 "
					if opc_coste="1" then
						seleccion=seleccion & ",coste "
					end if
					seleccion=seleccion & " FROM [" & session("usuario") & "]  AS f " & strbaja
					seleccion=seleccion & " GROUP BY nproveedor,nombre, referencia,descripcion,divisa,orden,tiene_escv,medidaVenta "
					if opc_cod_proyecto="1" then
						seleccion=seleccion & ",cod_proyecto "
					end if
					if opc_coste="1" then
						seleccion=seleccion & ",coste "
					end if					
					if ordenar = true then
						seleccion=seleccion & "order by orden desc"
					end if
				else
					if ordenar = true then
						seleccion = "select * from [" & session("usuario") & "]" & strbaja & "order by orden desc"
					else
						seleccion = "select * from [" & session("usuario") & "]" & strbaja
					end if
				end if
				'comprobamos si existen varias divisas
				rst.open seleccion,session("backendlistados"),adUseClient,adLockReadOnly
				MostrarDivisa = false
				w_divisa=""
				if not rst.eof then
					w_divisa = rst("Divisa")
					ocultar = 5
				      while (not rst.eof and MostrarDivisa=false)
				      	if w_divisa <> rst("Divisa") or rst("Divisa") <> MB then
							MostrarDivisa=true
							ocultar = 4
						end if
						w_divisa = rst("Divisa")
						rst.movenext
					wend
					rst.movefirst
				end if
				if not rst.eof then
					nproveedor=""
				      'Calculos de páginas--------------------------
					lote=limpiaCadena(Request.QueryString("lote"))
					if lote="" then
						lote=1
					else
						lote = clng(lote)
					end if
					sentido=limpiaCadena(Request.QueryString("sentido"))
					lotes=rst.RecordCount/MAXPAGINA
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

					rst.PageSize=MAXPAGINA
					rst.AbsolutePage=lote
					'-----------------------------------------%>
					<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'>
                    <%NavPaginas lote,lotes,campo,criterio,texto,1%>
					<br/>

					<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">
						<thead>
                        <%if nproveedor="" then
							%><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count+1-ocultar-1-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=LitProveedor%>:
								<%=Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("nproveedor")) & " - " & rst("Nombre")),LitVerProveedor)%>
							</td>
                        <%end if
						DrawFila color_fondo
						    DrawFila color_fondo%>
							    <td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitComprasNetas%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td>
                                <%if opc_cod_proyecto="1" then%>
									<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
								<%end if
								if MostrarDivisa=true then%>
                                    <td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td>
                                <%end if
								if opc_coste="1" then%>
                                    <td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCoste%></td>
                                <%end if
							CloseFila
						CloseFila
						%></thead>
						<tbody><%
							ValorCampoAnt = ""
							AcumuladoAnt  = ""
							ImprimeCabecera = false
							fila = 1
							while not rst.eof and fila<=MAXPAGINA
								CheckCadena rst("nproveedor")
								DrawFila ""
									for each campo in rst.fields
										if campo.name="tiene_escv" then
										else
											if (PorProveedor="SI" and ucase(campo.name)="NPROVEEDOR") then
												if rst(campo.name)<>ValorCampoAnt then
													if ValorCAmpoAnt<>"" then
														'antes de imprimir subtotales imprimimos conceptos (si existen)
														'seleccion2 = "select descripcion, sum(cantidad) as cantidad, sum(importe) as importe, (sum(importe)/sum(cantidad)) as precio from conceptos_fac_pro as c, facturas_pro as f where f.nproveedor='" & proveedor & "' and p.nfactura=f.nfactura group by descripcion"
														'Fila de Subtotal%>
														<td></td>
														<!--<td></td>-->
														<%'if opc_cod_proyecto="1" then%>
															<td></td>
														<%'end if%>
														<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
		            	                            	<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
														<%
														CloseFila
														DrawFila "" 'Fila de separacion
															%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%

														CloseFila
														DrawFila color_fondo
   									                        		%><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=LitProveedor%>:
															<%=Hiperv(OBJProveedores,rst("nproveedor"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("nproveedor")) & " - " & rst("Nombre")),LitVerProveedor)%>
															</td><%
														CloseFila
														DrawFila color_fondo%>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitComprasNetas%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td>
															<%if opc_cod_proyecto="1" then%>
																<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
															<%end if
															if MostrarDivisa=true then
																%><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
															end if
					                                        if opc_coste="1" then
						                                        %><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCoste%></td><%
					                                        end if															
														CloseFila
														ImprimeCabecera=true
													end if
												else
												   ImprimeCabecera=false
												end if
												ValorCampoAnt=rst(campo.name)
											elseif campo.name="Cantidad" or campo.name="Compras Netas" or campo.name="Precio Medio" or campo.name="coste" then 'Formateo del campo con importe
												'ajustamos divisas si es necesario
												if MostrarDivisa=true then
													n_decimales = null_z(d_lookup("ndecimales", "divisas", "codigo='" & rst("Divisa") & "'", session("backendlistados")))
												end if
												if campo.name="Compras Netas" then
 													AcumuladoAnt = rst("Acumulador")
												end if
												if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
													%><td class=tdbordeCELDA7 align="right">
														<%if rst("tiene_escv")<>1 or campo.name="Cantidad" then%>
															<%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",formatnumber(rst(campo.name),iif(campo.name<>"Cantidad",n_decimales,dec_cant),-1,0,-1)))%>
														<%end if%>
														<br/>
														<%if rst("cantidad2")<>0 and rst("cantidad2")<>"" then
															if campo.name="Cantidad" then%>
															<b><%=EncodeForHtml(rst("medidaVenta"))%> :</b> <%=EncodeForHtml(rst("cantidad2"))%>
															<%elseif campo.name="Precio Medio"  then%>
															<b>POR <%=EncodeForHtml(rst("medidaVenta"))%> :</b> <%=EncodeForHtml(formatnumber(rst("preciomedio2"),n_decimales,-1,0,-1))%>
															<%end if
														end if%>
													</td><%
												end if
											elseif  campo.name<>"cantidad2" and  campo.name<>"medidaVenta" and  campo.name<>"precioMedio2" then 'EJM 06/11/2006 Para mostrar las columnas restantes menos las indicadas
												if campo.name<>"Nombre" and ucase(campo.name)<>"NPROVEEDOR" and campo.name<>"Divisa" then
													if campo.name="Referencia" then
												      	if rst(campo.name)="zzzzzzzzzzzzzzzzzzzz" then
															%><td class=tdbordeCELDA7>Concepto</td><%
														else
															%><td class=tdbordeCELDA7>
																<%=Hiperv(OBJArticulos,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerArticulo)%>
															</td><%
														end if
													else
		                                          	    				if campo.name<>"Orden" and campo.name<>"Acumulador" then
															if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
														      	%><td class=tdbordeCELDA7><%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",rst(campo.name)))%></td><%
															end if
														else
															%><!--<td></td>--><%
														end if
													end if
													'Aki
													'end aki
												else
											   		if campo.name="Divisa" and MostrarDivisa=true then
											      		%><td class=TDBORDECELDA7><%=EncodeForHtml(d_lookup("abreviatura", "divisas", "codigo='" & rst(campo.name) & "'", session("backendlistados")))%></td><%
											   		else
						                          				if campo.name = "Divisa" and MostrarDivisa = false then '
									                         		%><!--<td></td>--><%
 									                      		end if
													end if
												end if
											end if
										end if
									next
								CloseFila
								fila = fila +1
								rst.movenext
							wend
							'mostramos subtotales si se alcanzó maxpagina y se va a cambiar de proveedor
							if not rst.eof then
		 			            	if fila>MAXPAGINA and rst("nproveedor")<>ValorCampoAnt then
									'antes de imprimir subtotales imprimimos conceptos (si existen)
									'Fila de Subtotal%>
									<td></td>
									<!--<td></td>-->
									<%'if opc_cod_proyecto="1" then%>
										<td></td>
									<%'end if%>
									<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
			                        <td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
									<%
									CloseFila
								end if
							end if
							'solo mostramos subtotales si se alcanzó el final del rst
							if rst.eof then
								if (PorProveedor="SI") then
									DrawFila "" 'Fila de Subtotal
										%><td></td>
										<!--<td></td>-->
										<%'if opc_cod_proyecto="1" then%>
											<td></td>
										<%'end if%>
										<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
										<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
										<%
									CloseFila
									DrawFila "" 'Fila de separacion
										%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
									CloseFila
								end if
								DrawFila "" 'Fila para el total
									%><td></td>
									<!--<td></td>-->
									<%'if opc_cod_proyecto="1" then%>
										<td></td>
									<%'end if%>
									<% Suma = elTotal%>
									<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
									<td class=tdbordeCELDA7 align="right"><big><b><%=EncodeForHtml(formatnumber(Suma,n_decimalesMB,-1,0,-1))%></b></big></td>
								<%CloseFila%>
								<%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
								   	DrawFila "" 'Fila para el total equivalencia en PTS
									      %><td></td>
										<!--<td></td>-->
										<%'if opc_cod_proyecto="1" then%>
											<td></td>
										<%'end if%>
							      		<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%>&nbsp;<%=EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")))%>:</b></td>
									      <td class=tdbordeCELDA7 align="right"><big><b><%=EncodeForHtml(formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))%></b></big></td>
								      <%CloseFila%>
								<%end if%>
							<%end if%>
						</tbody>
					</table><br/><%
					NavPaginas lote,lotes,campo,criterio,texto,2
				else
					%><script language="javascript" type="text/javascript">
					      alert("<%=LitNoExisteDatos%>");
					      parent.window.frames["botones"].document.location = "resumen_compras_pro_bt.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.action="resumen_compras_pro.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.submit();
					</script><%
				end if
				rst.close
			end if 'end if agrupa por proveedor

			'*************************************** AGRUPACION POR MESES ************************
			if agrupar=LitMeses then
				'***** COD : JCI-090103-01 *****%>
				<input type='hidden' name='elTotal' value='0'>
                <%strWhere=""
				strWhereArt=""
				strWhereSerie=""
				strfAlb= ""
				sin_conceptos=0

				if fdesde>"" then
					strWhere=strWhere & " and doc.fecha>=''" & fdesde & "''"
					if falbaran = "on" then
						strfAlb= strfAlb & " and (al.fecha>=''" & fdesde & "'' or al.fecha is null)"
					end if
				end if
				if fhasta>"" then
					strWhere=strWhere & " and doc.fecha<=''" & fhasta & "''"
					if falbaran = "on" then
						strfAlb= strfAlb & " and (al.fecha<=''" & fhasta & "'' or al.fecha is null)"
					end if
				end if

				if tactividad>"" then 'Se selecciono tipo de actividad
					desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
					strWhere=strWhere & " and c.tactividad=''" & tactividad & "''"
				end if                                                                                                    
				if nproveedor>"" then 'Se selecciono proveedor                                                        
					nproveedor=session("ncliente") & nproveedor
				      PorProveedor="NO"
					nompro = d_lookup("RAZON_SOCIAL","proveedores","nproveedor='" & nproveedor & "'",session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitProveedor%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nproveedor))%>&nbsp;&nbsp;<%=EncodeForHtml(nompro)%></b></font><br/><%
					strWhere=strWhere & " and c.nproveedor=''" & nproveedor & "''"
				else
					if opcproveedorbaja="" then
						t_opcproveedorbaja=0
					else
						t_opcproveedorbaja=1
						strWhere=strWhere & " and c.fbaja is null"%>
						  <font class=ENCABEZADO><b><%=LitProveedorBaja%></b></font><br/>
                    <%end if
				end if
                                                                                      
				if nserie>"" then 'Se selecciono serie
					%><font class=ENCABEZADO><b><%=LitSerie%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nserie))%>&nbsp;&nbsp;<%=EncodeForHtml(d_lookup("nombre","series","nserie='" & nserie & "'",session("backendlistados")))%></b></font><br/><%
					strWhereSerie=nserie
				end if
				if familia>"" then                                                       
					tsel2 = false
					desc_familia = d_lookup("nombre", "familias", "codigo='" & familia & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitSubFamilia%> : </b></font><font class='CELDA'><%=trimCodEmpresa()%>&nbsp;&nbsp;<%=EncodeForHtml(desc_familia)%></b></font><br/><%
					strWhereArt=strWhereArt & " and a.familia=''" & familia & "''"
					sin_conceptos=1                                                                
				end if
				if referencia>"" then
					%><font class=ENCABEZADO><b><%=LitConrefRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(referencia))%></b></font><br/><%
					strWhereArt = strWhereArt + " and substring(a.referencia,6,len(a.referencia)-5) like ''%" & trimCodEmpresa(referencia) & "%''"
					sin_conceptos=1
				end if                                                                                        
				if nombreart>"" then
					%><font class=ENCABEZADO><b><%=LitConNombreRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(nombreart)%></b></font><br/><%
					strWhereArt = strWhereArt + " and a.nombre like ''%" & nombreart & "%''"
	   			end if

				'if incalbpendfac="on" or incalbpendfac="1" then
				if seriesapf>"" then
					%><font class=ENCABEZADO><b><%=LitIncAlbPenFact%></b></font><br/><%
				end if

				if cod_proyecto>"" then
					%><font class=ENCABEZADO><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",session("backendlistados")))%></font><br/><%
					strWhere=strWhere & " and doc.cod_proyecto=''" & cod_proyecto & "''"
				end if

				if tipo_proveedor & "">"" then
					strWhere= strWhere & " and tipo_proveedor=''" & tipo_proveedor & "''"
				end if
				if tipo_articulo & "">"" then
					strWhereArt = strWhereArt & " and a.tipo_articulo=''" & tipo_articulo & "''"
					sin_conceptos=1
				end if
				if opc_coste>"" then
					%><font class=ENCABEZADO><b><%=LitCoste%><br/>
                <%end if
				if opc_cantidad>"" then
					%><font class=ENCABEZADO><b><%=LitCantidad%><br/>
                <%end if

				if opc_comprasnetas>"" then
					%><font class=ENCABEZADO><b><%=LitComprasNetas%><br/>
                <%end if%>
                <hr/>
                <%if Request.QueryString("lote") & "" = "" then
					'Para movimientros entre las páginas no borramos y volvemos a crear la tabla temporal ya que se creo y completó en la primera ejecución

                    ''ricardo 27-8-2007 se cambia el nombre del procedimiento
                    mostrar = "ARTICULOS"
                    if mostrarfilas = LitProveedores then
                        mostrar = "PROVEEDORES"
                    end if
					strQuery="Exec spl_ResumenVentasMesesCompras @p_tablaTemp='[" & session("usuario") & "]', @p_strWhere='" & strWhere & "', @p_strWhereArt='" & strWhereArt & "', @p_serie='" & strWhereSerie & "', @p_soloConceptos=" & iif(ver_conceptos="",0,1) & ", @p_desgloseConceptos=" & iif(conceptos="",0,1) & ", @p_albPendFac=" & iif(seriesapf="",0,1) & ", @p_serieAlbPendFac='"&lista_series2&"',@p_ticPendFac=" & iif(incticpendfac="",0,1) & ",@mostrarfilas='" & mostrar & "',@tipo='COMPRAS',@sin_conceptos=" &  sin_conceptos & ",@p_nempresa='" & session("ncliente") & "',@p_falbaran='"& falbaran &"',@p_strfAlb='"& strfAlb &"'"

					'Llamar al procedimiento almacenado para crear la tabla temporal con los datos del listado.
					set conVentasMeses = Server.CreateObject("ADODB.Connection")
					conVentasMeses.open session("backendlistados")
					conVentasMeses.execute(strQuery)
					conVentasMeses.close
					set conVentasMeses=nothing
				end if

				'''set_orden=request.form("ordenar_compras")

				if mostrarfilas=LitProveedores then
					if ordenar then
						strMeses="select * from [" & session("usuario") & "] order by totalimpcliente desc,nproveedor,nombre"
					else
						strMeses="select * from [" & session("usuario") & "] order by nproveedor,nombre"
					end if
				else
					if ordenar then
						strMeses="select * from [" & session("usuario") & "] order by sumtotalimp desc,nombre"
					else
						strMeses="select * from [" & session("usuario") & "] order by nombre"
					end if
				end if

				rst.cursorlocation=3
				rst.open strMeses,session("backendlistados"),adUseClient,adLockReadOnly
				if not rst.eof then

					maxregistros=rst("maxregistros")

				  	'Calculos de páginas--------------------------
					lote=limpiaCadena(Request.QueryString("lote"))
					if lote="" then
						lote=1
					else
						lote = clng(lote)
					end if
					sentido=limpiaCadena(Request.QueryString("sentido"))
					lotes=maxregistros/MAXPAGINA
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

					'rst.PageSize=MAXPAGINA
					'rst.AbsolutePage=lote

					referencia_actual=rst("referencia")
					referencia_old=rst("referencia")
					if mostrarfilas=LitProveedores then
						proveedor_old=rst("nproveedor")
						proveedor_actual=rst("nproveedor")
					end if
					n_referencia=1
					hasta_referencia=((lote-1)*MAXPAGINA)+1
					while not rst.eof and n_referencia<hasta_referencia
						referencia_old=rst("referencia")
						if mostrarfilas=LitProveedores then
							proveedor_old=rst("nproveedor")
						end if
						rst.movenext
						if not rst.eof then
							referencia_actual=rst("referencia")
							if mostrarfilas=LitProveedores then
								proveedor_actual=rst("nproveedor")
							end if
						else
							referencia_actual="@@@#/===000"
							if mostrarfilas=LitProveedores then
								proveedor_actual="@@@#/===000"
							end if
						end if
						if referencia_actual<>referencia_old or proveedor_actual<>proveedor_old then
							n_referencia=n_referencia+1
						end if
					wend

					'-----------------------------------------%>
					<input type='hidden' name='NumRegs' value='<%=EncodeForHtml(maxregistros)%>'>                   
<%
					NavPaginas lote,lotes,campo,criterio,texto,1
					if lotes>1 then
						%><hr/>
<%					end if

					proTemp=""
					refTemp=""
					fila=1
					totalCant=0
					totalImp=0

					'opc_cantidad=request.form("opc_cantidad")
					'opc_comprasnetas=request.form("opc_comprasnetas")
					verTotales=0

					dia=day(fdesde)
					if len(dia)=1 then
						dia="0" & dia
					end if
					mes=month(fdesde)
					if len(mes)=1 then
						mes="0" & mes
					end if
					ano=year(fdesde)
					fecha_desde=dia & "/" & mes & "/" & ano
					dia=day(fhasta)
					if len(dia)=1 then
						dia="0" & dia
					end if
					mes=month(fhasta)
					if len(mes)=1 then
						mes="0" & mes
					end if
					ano=year(fhasta)
					fecha_hasta=dia & "/" & mes & "/" & ano
					numero_meses=datediff("m",fecha_desde,fecha_hasta,0)+1

					dim tm_lista(12,2)

					if mostrarfilas=LitProveedores then
						nproveedor_old=rst("nproveedor")
					else
						referencia_old=rst("referencia")
					end if

					if mostrarfilas=ucase(LitArticulos) then%>
						<table width="100%" style="border-collapse: collapse;">
						<%'Mostrar la fila de encabezado para los artículos/meses%>
						<tr bgcolor="<%=color_terra%>">
							<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitReferencia%></td>
							<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitDescripcion%></td>

							<%
							poner_ano=0
							if year(fecha_desde)<>year(fecha_hasta) then
								poner_ano=1
							end if
							texto_ano=""
							anyo=""
							fecha_cual_vamos=fecha_desde
							for i=1 to numero_meses
								if i=1 and poner_ano=1 then
									texto_ano=" - " & year(fecha_cual_vamos)
									anyo=year(fecha_cual_vamos)
								end if
								%><td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=EncodeForHtml(ucase(mid(DescMes(month(fecha_cual_vamos)),1,3)) & texto_ano & iif(opc_comprasnetas>""," (" & MB_abrev & ")",""))%></td><%
								fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
								if anyo<>year(fecha_cual_vamos) and poner_ano=1 then
									texto_ano=" - " & year(fecha_cual_vamos)
								else
									texto_ano=""
								end if
								anyo=year(fecha_cual_vamos)
							next%>
							<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalCantidad%></td>
							<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalCompras & " (" & EncodeForHtml(MB_abrev) & ")"%></td>
						</tr>
                    <%end if

					while not rst.eof and fila<=MAXPAGINA
						if mostrarfilas=LitProveedores then
							CheckCadena rst("nproveedor")
							if rst("nproveedor")<>proTemp then
								for i=1 to 12
									tm_lista(i,1)=0
									tm_lista(i,2)=0
								next
								'Fila de encabezado del cliente
								if proTemp="" then%>
									<table width="100%" style="border-collapse: collapse;">
                                <%else
									''ricardo 22/1/2003 para que se veo o no lo de provedores en hojas separadas, ya que solo sirve para la agrupacion por cliente
									if prohojassep="on" or prohojassep="1" then
										%></table>
										<h6 class=SALTO>&nbsp;</h6>
										<table width="100%" style="border-collapse: collapse;">
                                    <%end if
								end if%>
								<tr bgcolor="<%=color_fondo%>">
									<td class="ENCABEZADOL" colspan="16"><%=LitProveedor & ": " & EncodeForHtml(rst("nproveedor")) & " - " & EncodeForHtml(rst("razon_social"))%></td>
								</tr>
                                <%'Mostrar la fila de encabezado para los artículos/meses%>
								<tr bgcolor="<%=color_terra%>">
									<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitReferencia%></td>
									<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitDescripcion%></td>
                                    <%poner_ano=0
									if year(fecha_desde)<>year(fecha_hasta) then
										poner_ano=1
									end if
									texto_ano=""
									anyo=""
									fecha_cual_vamos=fecha_desde
									for i=1 to numero_meses
										if i=1 and poner_ano=1 then
											texto_ano=" - " & year(fecha_cual_vamos)
											anyo=year(fecha_cual_vamos)
										end if%>
										<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=EncodeForHtml(ucase(mid(DescMes(month(fecha_cual_vamos)),1,3)) & texto_ano & iif(opc_comprasnetas>""," (" & MB_abrev & ")",""))%></td>
                                        <%fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
										if anyo<>year(fecha_cual_vamos) and poner_ano=1 then
											texto_ano=" - " & year(fecha_cual_vamos)
										else
											texto_ano=""
										end if
										anyo=year(fecha_cual_vamos)
									next%>
									<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalCantidad%></td>
									<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalCompras & " (" & EncodeForHtml(MB_abrev) & ")"%></td>
								</tr>
                                <%proTemp=rst("nproveedor")
							else
								if rst("referencia")<>refTemp then
									for i=1 to 12
										tm_lista(i,1)=0
										tm_lista(i,2)=0
									next

									refTemp=rst("referencia")
								end if
							end if
						end if%>
						<tr>
							<td class="CELDAL7" style="border: 1px solid Black;"><%=EncodeForHtml(trimCodEmpresa(rst("referencia")))%></td>
							<td class="CELDAL7" style="border: 1px solid Black;"><%=EncodeForHtml(rst("nombre"))%></td>
                            <%sumtotalcant=rst("sumtotalcant")
							sumtotalimp=rst("sumtotalimp")
							fecha_cual_vamos=fecha_desde
							referencia_old2=rst("referencia")
							i=1
							while i<=numero_meses and not rst.eof
								if mostrarfilas=LitProveedores then
									if referencia_old2=EncodeForHtml(rst("referencia")) and proTemp=EncodeForHtml(rst("nproveedor")) then
										if month(fecha_cual_vamos)<>cint(rst("mesdoc")) then%>
											<td class="CELDAR7" style="border: 1px solid Black;"></td>
                                        <%else
											mostrar_Cant_Meses "CELDAR7",opc_cantidad,EncodeForHtml(rst("totalcant")),EncodeForHtml(rst("totalimp")),n_decimalesMB
											rst.movenext
										end if
										fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
										i=i+1
									else
										%><td class="CELDAR7" style="border: 1px solid Black;"></td><%
										i=i+1
									end if
								else
									if referencia_old2=EncodeForHtml(rst("referencia")) then
										if month(fecha_cual_vamos)<>cint(rst("mesdoc")) then
											%><td class="CELDAR7" style="border: 1px solid Black;"></td><%
										else
											mostrar_Cant_Meses "CELDAR7",opc_cantidad,EncodeForHtml(rst("totalcant")),EncodeForHtml(rst("totalimp")),n_decimalesMB
											rst.movenext
										end if
										fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
										i=i+1
									else
										%><td class="CELDAR7" style="border: 1px solid Black;"></td><%
										i=i+1
									end if
								end if
							wend
							for j=i to numero_meses%>
								<td class="CELDAR7" style="border: 1px solid Black;"></td>
                            <%next
							rst.moveprevious

							mostrar_Cant_Meses "CELDAR7","---",sumtotalcant,"0",n_decimalesMB
							mostrar_Cant_Meses "CELDAR7","---",sumtotalimp,"0",n_decimalesMB%>
						</tr>
                        <%totalimpproveedor=rst("totalimpcliente")
						totalcantproveedor=rst("totalcantcliente")
						if mostrarfilas=LitProveedores then
							nproveedor_old=rst("nproveedor")
						else
							referencia_old=rst("referencia")
						end if

						fila=fila+1
						rst.movenext

						if rst.eof then
							verTotales=1
						else
							if mostrarfilas=LitProveedores then
								if EncodeForHtml(rst("nproveedor"))<>proTemp then
									verTotales=1
								end if
							else
								if EncodeForHtml(rst("referencia"))<>refTemp then
									verTotales=0
								end if
							end if
						end if

						if verTotales=1 then
							if mostrarfilas=LitProveedores then
								strPorMeses="select sum(totalimp) as sumtotalimp,sum(totalcant) as sumtotalcant,mesdoc from [" & session("usuario") & "] where nproveedor='" & nproveedor_old & "' group by mesdoc order by convert(int,mesdoc)"
							else
								strPorMeses="select sum(totalimp) as sumtotalimp,sum(totalcant) as sumtotalcant,mesdoc from [" & session("usuario") & "]  group by mesdoc order by convert(int,mesdoc)"
							end if
							rstAux.cursorlocation=3
							rstAux.open strPorMeses,session("backendlistados")
							while not rstAux.eof
								tm_lista(cint(rstAux("mesdoc")),1)=tm_lista(cint(rstAux("mesdoc")),1) + EncodeForHtml(rstAux("sumtotalcant"))
								tm_lista(cint(rstAux("mesdoc")),2)=tm_lista(cint(rstAux("mesdoc")),2) + EncodeForHtml(rstAux("sumtotalimp"))
								rstAux.movenext
							wend
							rstAux.close

							'Mostrar la fila de totales por proveedor.
							%><tr bgcolor="<%=color_terra%>">
								<td class="ENCABEZADOL7" style="border: 1px solid Black;">&nbsp;</td>
								<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitTotales%></td>
								<%
								fecha_cual_vamos=fecha_desde
								for i=1 to numero_meses
									mes_a_poner=month(fecha_cual_vamos)
									mostrar_Cant_Meses "ENCABEZADOR7","---",iif(opc_cantidad>"",tm_lista(mes_a_poner,1),tm_lista(mes_a_poner,2)),"0",n_decimalesMB
									fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
								next%>
								<%mostrar_Cant_Meses "ENCABEZADOR7","---",totalcantproveedor,"0",n_decimalesMB%>
								<%mostrar_Cant_Meses "ENCABEZADOR7","---",totalimpproveedor,"0",n_decimalesMB%>
							</tr><%

							totalCant=0
							totalImp=0

							verTotales=0
						end if
					wend

					''mostramos ahora una fila de totales del listado
					if mostrarfilas=LitProveedores and rst.eof then
						for i=1 to 12
							tm_lista(i,1)=0
							tm_lista(i,2)=0
						next
						strPorMeses="select sum(totalimp) as sumtotalimp,sum(totalcant) as sumtotalcant,mesdoc,(select sum(totalimp) from [" & session("usuario") & "]) as totalimp,(select sum(totalcant) from [" & session("usuario") & "]) as totalcant from [" & session("usuario") & "] group by mesdoc order by convert(int,mesdoc)"
						rstAux.cursorlocation=3
						rstAux.open strPorMeses,session("backendlistados")
						if not rstAux.eof then
							sumatotalimp=rstAux("totalimp")
							sumatotalcant=rstAux("totalcant")
						end if
						while not rstAux.eof
							tm_lista(cint(rstAux("mesdoc")),1)=tm_lista(cint(rstAux("mesdoc")),1) + EncodeForHtml(rstAux("sumtotalcant"))
							tm_lista(cint(rstAux("mesdoc")),2)=tm_lista(cint(rstAux("mesdoc")),2) + EncodeForHtml(rstAux("sumtotalimp"))
							rstAux.movenext
						wend
						rstAux.close
						'Mostrar la fila de totales.
						%><tr bgcolor="<%=color_terra%>">
							<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitResComTotalLis%></td>
							<td class="ENCABEZADOL7" style="border: 1px solid Black;">&nbsp;</td>
<%
							fecha_cual_vamos=fecha_desde
							for i=1 to numero_meses
								mes_a_poner=month(fecha_cual_vamos)
								mostrar_Cant_Meses "ENCABEZADOR7","---",iif(opc_cantidad>"",tm_lista(mes_a_poner,1),tm_lista(mes_a_poner,2)),"0",n_decimalesMB
								fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
							next%>
							<%mostrar_Cant_Meses "ENCABEZADOR7","---",sumatotalcant,"0",n_decimalesMB%>
							<%mostrar_Cant_Meses "ENCABEZADOR7","---",sumatotalimp,"0",n_decimalesMB%>
						</tr>
                    <%end if%>
					</table>
                    <%NavPaginas lote,lotes,campo,criterio,texto,2
				else%>
					<script language="javascript" type="text/javascript">
					    alert("<%=LitNoExisteDatos%>");
					    parent.window.frames["botones"].document.location = "resumen_compras_pro_bt.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.action="resumen_compras_pro.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.submit();
					</script>
                <%end if
				rst.close
			end if 'end if agrupa por meses

	'*************************************** AGRUPACION POR PROYECTOS ********************

			if agrupar=ucase(LitProyecto) then                                                                                              
				PorProyecto = "SI"                                                                           
				opc_cod_proyecto="1"
				if tactividad>"" then 'Se selecciono tipo de actividad
					desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
				end if
				if nproveedor>"" then 'Se selecciono proveedor
					nproveedor=session("ncliente") & completar(nproveedor,5,"0")
					nompro = d_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitProveedor%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nproveedor))%>&nbsp;&nbsp;<%=EncodeForHtml(nompro)%></b></font><br/><%
					t_opcproveedorbaja=1                                                                               
				else                                                                                   
					if opcproveedorbaja="" then
						t_opcproveedorbaja=1
					else
						t_opcproveedorbaja=0
						%><font class=ENCABEZADO><b><%=LitProveedorBaja%></b></font><br/><%
					end if
				end if
                                                                                                     
				if nserie>"" then 'Se selecciono serie
					%><font class=ENCABEZADO><b><%=LitSerie%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nserie))%>&nbsp;&nbsp;<%=EncodeForHtml(d_lookup("nombre","series","nserie='" & nserie & "'",session("backendlistados")))%></b></font><br/><%
				end if
				if familia>"" then
					tsel2 = false
					desc_familia = d_lookup("nombre", "familias", "codigo='" & familia & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitSubFamilia%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(familia))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_familia)%></b></font><br/><%
				end if                                                                                        
				if cod_proyecto>"" then                                                                                                
					PorProyecto="NO"
					%><font class=ENCABEZADO><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",session("backendlistados")))%></font><br/><%
				end if
				if nombreart>"" then 'Se selecciono nombreart
					%><font class=ENCABEZADO><b><%=LitConNombreRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(nombreart)%></b></font><br/><%                  
				end if
				if referencia>"" then 'Se selecciono referencia
					%><font class=ENCABEZADO><b><%=LitConrefRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(referencia))%></b></font><br/><%                       
				end if
				'if incalbpendfac="on" or incalbpendfac="1" then
				if seriesapf>"" then
					%><font class=ENCABEZADO><b><%=LitIncAlbPenFact%></b></font><br/><%
				end if
				if request.querystring("save")="true" then
					crearProyecto tsel1, tsel2, fdesde, fhasta, tactividad, nproveedor, nserie, familia, referencia, nombreart, ordenar, MB,t_opcproveedorbaja,cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas
				else
					%><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' /><%
				end if
				%><hr/><%
				'if incalbpendfac="on" or incalbpendfac="1" then
				if seriesapf>"" then
					seleccion="SELECT cod_proyecto,NProveedor,Nombre,Referencia,Descripcion,"
					seleccion=seleccion & "SUM(cantidad) AS Cantidad, SUM([Compras Netas]) AS [Compras Netas],"
					seleccion=seleccion & "SUM([Precio Medio] * cantidad)/ CASE WHEN SUM(cantidad)= 0 THEN 1 ELSE SUM(cantidad) END AS [Precio Medio],"
					seleccion=seleccion & "Divisa,(select sum([Compras Netas]) from [" & session("usuario") & "] where cod_proyecto=f.cod_proyecto) as Acumulador,Orden"
					seleccion=seleccion & ",tiene_escv "
					seleccion=seleccion & " FROM [" & session("usuario") & "] as f " & strbaja
					seleccion=seleccion & " GROUP BY cod_proyecto,NProveedor,Nombre,Referencia,Descripcion,divisa,orden,tiene_escv "
					if ordenar = true then
						seleccion=seleccion & "order by orden desc,cod_proyecto"
					end if
				else
					if ordenar = true then
						seleccion = "select * from [" & session("usuario") & "]" & strbaja & "order by orden desc,cod_proyecto"
					else
						seleccion = "select * from [" & session("usuario") & "]" & strbaja & "order by cod_proyecto"
					end if
				end if

				'comprobamos si existen varias divisas
				rst.open seleccion,session("backendlistados"),adUseClient,adLockReadOnly
				MostrarDivisa = false
				w_divisa=""
				if not rst.eof then
					w_divisa = EncodeForHtml(rst("Divisa"))
					ocultar = 4
					while (not rst.eof and MostrarDivisa=false)
						if w_divisa <> EncodeForHtml(rst("Divisa")) or EncodeForHtml(rst("Divisa")) <> MB then
							MostrarDivisa=true
							ocultar = 4
						end if
						w_divisa = EncodeForHtml(rst("Divisa"))
						rst.movenext
					wend
					rst.movefirst
				end if
				if not rst.eof then

				      'Calculos de páginas--------------------------
					lote=limpiaCadena(Request.QueryString("lote"))
					if lote="" then
						lote=1
					else
						lote = clng(lote)
					end if
					sentido=limpiaCadena(Request.QueryString("sentido"))
					lotes=rst.RecordCount/MAXPAGINA
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

					rst.PageSize=MAXPAGINA
					rst.AbsolutePage=lote
					'-----------------------------------------%>
					<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'><%
					NavPaginas lote,lotes,campo,criterio,texto,1
					%><br/>


					<table width="100%" style="border-collapse: collapse;">
						<thead><%
							if proyecto="" then
								%><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-1-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=LitProyecto%>:
									<%=EncodeForHtml(rst("cod_proyecto"))%>
								</td><%
							end if
							DrawFila color_fondo
							    DrawFila color_fondo%>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNProveedor%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitComprasNetas%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td><%
								if MostrarDivisa=true then
									%><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
								end if
							CloseFila
						%></thead>
						<tbody><%
							ValorCampoAnt = ""
							AcumuladoAnt  = ""
							ImprimeCabecera = false
							fila = 1
							while not rst.eof and fila<=MAXPAGINA
								CheckCadena rst("nproveedor")
								DrawFila ""
									for each campo in rst.fields
										if campo.name="tiene_escv" then
										else
											if (PorProyecto="SI" and campo.name="cod_proyecto") then
												if rst(campo.name)<>ValorCampoAnt then
													if ValorCAmpoAnt<>"" then
														'antes de imprimir subtotales imprimimos conceptos (si existen)
														'seleccion2 = "select descripcion, sum(cantidad) as cantidad, sum(importe) as importe, (sum(importe)/sum(cantidad)) as precio from conceptos as c, facturas_cli as f where f.nproveedor='" & proveedor & "' and p.nfactura=f.nfactura group by descripcion"
														'Fila de Subtotal%>
														<td></td>
														<td></td>
														<td></td>
														<%if opc_cod_proyecto="1" then%>
															<td></td>
														<%end if%>
														<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
			                                        	<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
														<%
														CloseFila
														DrawFila "" 'Fila de separacion
															%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
														CloseFila
														DrawFila color_fondo
		   									                        %><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-1-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=LitProyecto%>:
																<%=EncodeForHtml(rst("cod_proyecto"))%>
															</td><%
														CloseFila
														DrawFila color_fondo%>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNProveedor%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitComprasNetas%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td><%
															if MostrarDivisa=true then
																%><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
															end if
														CloseFila
														ImprimeCabecera=true
													end if
												else
													ImprimeCabecera=false
												end if
												ValorCampoAnt=rst(campo.name)
											elseif campo.name="Cantidad" or campo.name="Compras Netas" or campo.name="Precio Medio" then 'Formateo del campo con importe
												'ajustamos divisas si es necesario
												if MostrarDivisa=true then
													n_decimales = null_z(d_lookup("ndecimales", "divisas", "codigo='" & rst("Divisa") & "'", session("backendlistados")))
												end if
												if campo.name="Compras Netas" then
	 												AcumuladoAnt = rst("Acumulador")
												end if
												if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
													%><td class=tdbordeCELDA7 align="right">
														<%if rst("tiene_escv")<>1 or campo.name="Cantidad" then%>
															<%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",formatnumber(rst(campo.name),iif(campo.name<>"Cantidad",n_decimales,DEC_CANT),-1,0,-1)))%>
														<%end if%>
													</td><%
												end if
											else
												if campo.name<>"cod_proyecto" and campo.name<>"Divisa" then
													if campo.name="Referencia" then
												      	if rst(campo.name)="zzzzzzzzzzzzzzzzzzzz" then
															%><td class=tdbordeCELDA7>Concepto</td><%
														else
															%><td class=tdbordeCELDA7>
																<%=Hiperv(OBJArticulos,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerArticulo)%>
															</td><%
														end if
													else
	            	                                  	if campo.name<>"Orden" and campo.name<>"Acumulador" then
															if campo.name="NProveedor" or ucase(campo.name)="NPROVEEDOR" then
																%><td class=tdbordeCELDA7>
																	<%=Hiperv(OBJProveedores,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerProveedor)%>
																</td><%
															else
														  		if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
														         		%><td class=tdbordeCELDA7><%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",rst(campo.name)))%></td><%
																end if
															end if
													  	else
													     		%><!--<td>hola2</td>--><%
													  	end if
													end if
													'end aki
												else
													if campo.name="Divisa" and MostrarDivisa=true then
											      		%><td class=tdbordeCELDA7><%=EncodeForHtml(d_lookup("abreviatura", "divisas", "codigo='" & rst(campo.name) & "'", session("backendlistados")))%></td><%
													else
				                          			if campo.name = "Divisa" and MostrarDivisa = false then '
					                         			%><!--<td>hola3</td>--><%
					                      			end if
													end if
												end if
											end if
										end if
									next
								CloseFila
								fila = fila +1
								rst.movenext
							wend

							'solo mostramos subtotales si se alcanzó el final del rst
							if rst.eof then
								if (PorProyecto="SI") then
									DrawFila "" 'Fila de Subtotal
										%><td></td>
										<td></td>
							   			<td></td>
										<%'if opc_cod_proyecto="1" then%>
											<td></td>
										<%'end if%>
										<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
										<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
								   	<%CloseFila
								   	DrawFila "" 'Fila de separacion
										%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
     									CloseFila
								end if
								DrawFila "" 'Fila para el total
									%><td></td>
									<td></td>
							   			<td></td>
									<%'if opc_cod_proyecto="1" then%>
										<td></td>
									<%'end if%>
									<% Suma = elTotal%>
									<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
									<td class=tdbordeCELDA7 align="right"><big><b><%=EncodeForHtml(formatnumber(Suma,n_decimalesMB,-1,0,-1))%></b></big></td>
								<%CloseFila%>
								<%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
									DrawFila "" 'Fila para el total equivalencia en PTS
					      				%><td></td>
							   			<td></td>
							   			<td></td>
										<%'if opc_cod_proyecto="1" then%>
											<td></td>
										<%'end if%>
					      				<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%>&nbsp;<%=EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")))%>:</b></td>
				      					<td class=tdbordeCELDA7 align="right"><big><b><%=EncodeForHtml(formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))%></b></big></td>
									<%CloseFila%>
								<%end if%>
							<%end if%>
						</tbody>
					</table><br/><%
					NavPaginas lote,lotes,campo,criterio,texto,2
				else
					%><script language="javascript" type="text/javascript">
					      alert("<%=LitNoExisteDatos%>");
					      parent.window.frames["botones"].document.location = "resumen_compras_pro_bt.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.action="resumen_compras_pro.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.submit();
						
					</script><%
				end if
				rst.close
			end if 'end if agrupa por proyecto

'********************************************** AGRUPACION POR ARTICULOS *************************
			if agrupar=ucase(LitArticulo) then
				PorArticulo = "SI"                                                                                                   
				if familia>"" then                                                                                                                                           
					tsel2 = false
					desc_familia = d_lookup("nombre", "familias", "codigo='" & familia & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitSubFamilia%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(familia))%>&nbsp;&nbsp;<%=EncodeForHtml(trimCodEmpresa(desc_familia))%></b></font><br/><%
				end if                                                                                    
				if tactividad>"" then 'Se selecciono tipo de actividad                                                                                 
					desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
				end if
				if nproveedor>"" then 'Se selecciono proveedor                                                                                  
					nproveedor=session("ncliente") & completar(nproveedor,5,"0")
					nomcli = d_lookup("RAZON_SOCIAL","proveedores","nproveedor='" & nproveedor & "'",session("backendlistados"))
					%><font class=ENCABEZADO><b><%=LitProveedor%> : </b></font><font class='CELDA'><%=EncodeForHtml(nproveedor)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><br/><%
				else
					if opcproveedorbaja="" then
						t_opcproveedorbaja=0
					else
						t_opcproveedorbaja=1
						%><font class=ENCABEZADO><b><%=LitProveedorBaja%></b></font><br/><%
					end if                                                                         
				end if
				if nserie>"" then 'Se selecciono serie
					%><font class=ENCABEZADO><b><%=LitSerie%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nserie))%>&nbsp;&nbsp;<%=EncodeForHtml(d_lookup("nombre","series","nserie='" & nserie & "'",session("backendlistados")))%></b></font><br/><%
				end if
				if cod_proyecto>"" then 'Se selecciono proyecto
					%><font class=ENCABEZADO><b><%=LitProyecto%> : </b></font><font class='CELDA'><%=EncodeForHtml(cod_proyecto)%>&nbsp;&nbsp;<%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",session("backendlistados")))%></b></font><br/><%
				end if
				if nombreart>"" then 'Se selecciono nombreart                                                                                              
					%><font class=ENCABEZADO><b><%=LitConNombreRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(nombreart)%></b></font><br/><%
				end if
				if referencia>"" then 'Se selecciono referencia                                                          
					''referencia=session("ncliente") & referencia
					%><font class=ENCABEZADO><b><%=LitConrefRC%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(referencia))%></b></font><br/><%
				end if
				'if incalbpendfac="on" or incalbpendfac="1" then
				if seriesapf>"" then
					%><font class=ENCABEZADO><b><%=LitIncAlbPenFact%></b></font><br/><%
				end if
				%><hr/><%
				crearArticulo tsel1, tsel2,fdesde,fhasta, familia, referencia, nombreart, tactividad, nproveedor, nserie, desglose, ordenar, MB,t_opcproveedorbaja,cod_proyecto,tipo_proveedor,tipo_articulo,mostrarfilas

				'if incalbpendfac="on" or incalbpendfac="1" then
				if seriesapf>"" then
					seleccion="SELECT Ref, Descripcion,NProveedor,Nombre,"
					seleccion=seleccion & "SUM(cantidad) AS Cantidad, SUM([Compras Netas]) AS [Compras Netas],"
					seleccion=seleccion & "SUM([Precio Medio] * cantidad)/ CASE WHEN SUM(cantidad)= 0 THEN 1 ELSE SUM(cantidad) END AS [Precio Medio],"
					if opc_cod_proyecto="1" then
						seleccion=seleccion & "cod_proyecto,"
					end if
					seleccion=seleccion & "Divisa,(select sum(CASE WHEN divisa = '" & session("ncliente") & "02' THEN [Compras Netas] ELSE ([Compras Netas] / 166.386) END) from [" & session("usuario") & "] WHERE Ref = f.Ref) as AcumulaCompras,(select sum(cantidad) from [" & session("usuario") & "] WHERE Ref = f.Ref) as AcumulaCantidad,Orden"
					seleccion=seleccion & ",tiene_escv "
					seleccion=seleccion & " FROM [" & session("usuario") & "] as f" & strbaja
					seleccion=seleccion & " GROUP BY ref,NProveedor,Nombre,descripcion,divisa,orden,tiene_escv "
					if opc_cod_proyecto="1" then
						seleccion=seleccion & ",cod_proyecto "
					end if
					if ordenar = true then
						seleccion=seleccion & " order by orden desc"
					end if
				else
					if ordenar=true then
						seleccion = "select * from [" & session("usuario") & "]" & strbaja & " order by orden desc"
					else
						seleccion = "select * from [" & session("usuario") & "]" & strbaja
					end if
			  	end if
				rst.open seleccion,session("backendlistados"),adUseClient, adLockReadOnly
				MostrarDivisa = false
				w_divisa=""
				if not rst.eof then
		 			w_divisa = rst("Divisa")
					ocultar = 3+3
					while (not rst.eof and MostrarDivisa=false)
						if w_divisa <> rst("Divisa")  or rst("Divisa") <> MB then
							MostrarDivisa=true
							ocultar = 2+3
						end if
						w_divisa = rst("Divisa")
						rst.movenext
					wend
					rst.movefirst
				end if
			      if not rst.eof then
					etiq = ""
					desc = ""
					ref = rst("Ref")
					if rst("Descripcion")= "@concepto@" then
						if desglose = false then
							ref = ""
							etiq = LitConceptos
						else
							etiq = LitConcepto
						end if
						desc = ""
					else
						etiq = LitArticulo
						desc = rst("Descripcion")
					end if
					ocultar=ocultar-1
					'Calculos de páginas--------------------------
					lote=limpiaCadena(Request.QueryString("lote"))
					if lote="" then lote=1
					sentido=limpiaCadena(Request.QueryString("sentido"))
					lotes=rst.RecordCount/MAXPAGINA
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

					rst.PageSize=MAXPAGINA
					rst.AbsolutePage=lote                                                                                                            
					'-----------------------------------------%>
					<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'><%
					NavPaginas lote,lotes,campo,criterio,texto,1%><br/>

			 		<table width="100%" style="border-collapse: collapse;">
						<thead>
							<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-1-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=enc.EncodeForHtmlAttribute(etiq)%>:
								<%if etiq=LitConcepto or etiq = LitConceptos then%>
									<%=EncodeForHtml(ref & "  " & desc)%>
								<%else%>
									<%=Hiperv(OBJArticulos,ref,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(ref) & "  " & desc),LitVerArticulo)%>
								<%end if%>
							</td><%
							DrawFila color_fondo
							    DrawFila color_fondo%>
					                <td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitNproveedor%></td>
					                <td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
					                <%'if campo.name="xxxxxxxxxxxxcod_proyecto" or opc_cod_proyecto="1" then
					                if opc_cod_proyecto="1" then%>
						                <td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
					                <%end if%>
					                <td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
					                <td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitComprasNetas%></td>
					                <td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td><%
					                if MostrarDivisa = true then
						                %><td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
					                end if
				                CloseFila
						%></thead>
						<tbody><%
							ValorCampoAnt = ""
							fila = 1
							while not rst.eof and fila<=MAXPAGINA
								CheckCadena rst("nproveedor")
								DrawFila ""
									for each campo in rst.fields
										if campo.name="tiene_escv" then
										else
											if (PorArticulo="SI" and ucase(campo.name)="REF") then
												if ucase(rst(campo.name))<>ValorCampoAnt then
													if ValorCAmpoAnt<>"" then
												   		'antes de imprimir subtotales imprimimos conceptos (si existen)
														'Fila de Subtotal%>
														<!--<td></td>-->
														<%'if opc_cod_proyecto="1" then%>
															<td></td>
														<%'end if%>
														<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
														<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(SubTotalCant,dec_cant,-1,0,-1))%></b></td>
														<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(SubTotal,n_decimalesMB,-1,0,-1))%></b></td>
														<%
														pmTotal=0
														if SubTotalCant<>0 then pmTotal = cdbl(SubTotal)/SubTotalCant
														%>
														<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(pmTotal,n_decimalesMB,-1,0,-1))%></b></td>
														<%
													  	CloseFila
														DrawFila "" 'Fila de separacion
															%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
														CloseFila
														DrawFila color_fondo
														  	etiq = ""
														 	ref = rst("Ref")
															if rst("Descripcion")= "@concepto@" then
																if desglose = false then
																	ref = ""
																	etiq = LitConceptos
																else
																	etiq = LitConcepto
																end if
																desc = ""
															else                                                                                                           
																etiq = LitArticulo
																desc = rst("Descripcion")
															end if
															%>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-1-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=EncodeForHtml(etiq)%>:
																<%if etiq = LitConcepto or etiq = LitConceptos then%>
																	<%=EncodeForHtml(ref & " " & desc)%>
																<%else%>
																	<%=Hiperv(OBJArticulos,ref,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(ref)),LitVerArticulo)%> &nbsp;&nbsp;<%=EncodeForHtml(desc)%>
																<%end if%>
															</td><%
														CloseFila
														DrawFila color_fondo%>
															<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=Litnproveedor%></td>
															<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
															<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
															<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitComprasNetas%></td>
															<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td>
															<%if campo.name="xxxxxxxxxxcod_proyecto" or opc_cod_proyecto="1" then%>
																<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
															<%end if
															if MostrarDivisa = true then
																%><td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
															end if
														CloseFila
													end if 'valorcampoant<>""
												end if 'rst(campo.name)<>ValorCampoAnt
												ValorCampoAnt=ucase(rst(campo.name))
											elseif campo.name="Cantidad" or campo.name="Compras Netas" or campo.name="Precio Medio" then 'Formateo del campo con importe
												'ajustamos divisas si es necesario
												if MostrarDivisa=true then
													n_decimales = null_z(d_lookup("ndecimales", "divisas", "codigo='" & rst("Divisa") & "'", session("backendlistados")))
												end if
												%><td class=tdbordeCELDA7 align="right">
													<%if rst("tiene_escv")<>1 or campo.name="Cantidad" then%>
														<%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",formatnumber(rst(campo.name),iif(campo.name<>"Cantidad",n_decimales,dec_cant),-1,0,-1)))%>
													<%end if%>
												</td><%
											else
												if campo.name<>"Descripcion" and campo.name<>"Ref"  and campo.name<>"Divisa" then
													if campo.name="Nproveedor"  or ucase(campo.name)="NPROVEEDOR" then
														%><td class="tdbordeCELDA7">
															<%=Hiperv(OBJProveedores,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerProveedor)%>
														</td><%
													else
														if campo.name<>"AcumulaCompras" and campo.name<>"AcumulaCantidad" and campo.name<>"Orden" then
															if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
																%><td class=tdbordeCELDA7><%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",rst(campo.name)))%></td><%
															end if
														else
															%><!--<td></td>--><%
														end if
													end if
												else
													if campo.name = "Divisa" and MostrarDivisa = true then '
														if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
															%><td class=tdbordeCELDA7><%=EncodeForHtml(d_lookup("abreviatura", "divisas", "codigo='" & rst(campo.name) & "'", session("backendlistados")))%></td><%
														end if
													else
														if campo.name = "Divisa" and MostrarDivisa = false then '
															%><!--<td></td>--><%
														end if
													end if
												end if
											end if
										end if
									next
								CloseFila
								if ordenar=false then
								     SubTotal = rst("acumulacompras")
								     SubTotalCant = rst("acumulacantidad")
								else
									SubTotal = rst("orden")
									SubTotalCant = rst("acumulacantidad")
								end if
								rst.movenext
								fila = fila + 1
							wend
							if rst.eof then
								if (PorArticulo="SI") then
							      	DrawFila "" 'Fila de Subtotal
					     					%><!--<td></td>-->
										<%'if opc_cod_proyecto="1" then%>
											<td></td>
										<%'end if%>
										<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
										<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(SubTotalCant,dec_cant,-1,0,-1))%></b></td>
										<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(SubTotal,n_decimalesMB,-1,0,-1))%></b></td>
			 							<%
										pmTotal=0
										if SubTotalCant<>0 then pmTotal = cdbl(SubTotal)/SubTotalCant
										%>
										<td class=tdbordeCELDA7 align="right"><b><%=EncodeForHtml(formatnumber(pmTotal,n_decimalesMB,-1,0,-1))%></b></td>
										<%
									CloseFila
									DrawFila "" 'Fila de separacion
										%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
									CloseFila
								end if
								DrawFila "" 'Fila para el total
									Suma = elTotal
									%><td></td>
									<%if opc_cod_proyecto="1" then%>
										<td></td>
									<%end if%>
									<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
									<td></td>
									<td class=tdbordeCELDA7 align="right"><big><b><%=EncodeForHtml(formatnumber(Suma,n_decimalesMB,-1,0,-1))%></b></big></td>
								<%CloseFila%>
								<%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
									DrawFila "" 'Fila para el total equivalencia en PTS
										%><td></td>
										<%if opc_cod_proyecto="1" then%>
											<td></td>
										<%end if%>
										<td class=tdbordeCELDA7 bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotales%>&nbsp;<%=EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")))%>:</b></td>
										<td></td>
										<td class=tdbordeCELDA7 align="right"><big><b><%=EncodeForHtml(formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))%></b></big></td>
									<%CloseFila%>
								<%end if%>
							<%end if%>
						</tbody>
				     </table>
					<br/>
					<%
					NavPaginas lote,lotes,campo,criterio,texto,2
				else
					%><script language="javascript" type="text/javascript">
					      alert("<%=LitNoExisteDatos%>");
					      parent.window.frames["botones"].document.location = "resumen_compras_pro_bt.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.action="resumen_compras_pro.asp?mode=select1";
						parent.pantalla.resumen_compras_pro.submit();
						
					</script><%
				end if 'end if rst.eof (seleccion)
				rst.close
			end if 'end if agrupa por articulo

            %><iframe name="frameExportar" style='display:' src="resumen_compras_pro_exportar.asp?mode=ver" frameborder='0' width='500' height='200'></iframe><%
		end if%>
</form>
<%end if
connRound.close
set connRound = Nothing
set rstFunction = Nothing
set rstTemp = Nothing
set rstFunction = Nothing
set rstTemp = Nothing
set rstFunction = Nothing
set rstTemp = Nothing
set rstAux = Nothing
set rst = Nothing
set rst2 = Nothing
set rstSelect = Nothing
set rstTablas = Nothing
%>
</body>
</html>