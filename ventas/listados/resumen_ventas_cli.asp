<%@ Language=VBScript %>
<%'------------------------------CODIGOS DE AÑADIDURAS/MODIFICACIONES ------------------------
'JCI-090103-01 : Añado lo del total, para que no casque la creacion del pdf al agrupar por meses
'	FECHA :09/01/03
' AUTOR :JCI
'MPC 16/11/2007 CAMBIO DSN PARA LISTADOS
'----------------------------------------------------------------------------------------------%>
<%
''ricardo 23/1/2003
''se añade que se pueda elegir entre un cliente y otro

'JCI 02/03/2003 : Permitir la inclusión de tickets pendientes de facturar. Control de caché
'JA 26/06/03: Migración monobase.'
%>
<% Server.ScriptTimeout = 400 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
    <head>
        <title><%=LitTituloResVent%></title>
        
        <% dim  enc
        set enc = Server.CreateObject("Owasp_Esapi.Encoder")

        function EncodeForHtml(data)
	        if data & "" <> "" then
	          EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	        else
	          EncodeForHtml = data
	        end if
        end function
        %>

        <meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>" />
        <!--#include file="../../constantes.inc" -->
        <!--#include file="../../cache.inc" -->
        <!--#include file="../../calculos.inc" -->
        <%if accesoPagina(session.sessionid,session("usuario"))=1 then%>
        <!--#include file="../../ilion.inc" -->
        <!--#include file="../../mensajes.inc" -->
        <!--#include file="../../adovbs.inc" -->
        <!--#include file="../../varios.inc" -->
        <!--#include file="../../ico.inc" -->
        <!--#include file= "../../tablasResponsive.inc" -->
        <!--#include file="../../js/generic.js.inc"-->
        <!--#include file="../../js/calendar.inc" -->
        <!--#include file="../../modulos.inc" -->
        <!--#include file= "../../CatFamSubResponsive.inc"-->
        <!--#include file="../facturas_cli.inc" -->
        <!--#include file= "../../styles/formularios.css.inc"-->  

        <link rel="stylesheet" href="../../pantalla.css" media="screen" />
        <link rel="stylesheet" href="../../impresora.css" media="print" />
    </head>

    <script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
    <script language="javascript" type="text/javascript">
        function BloquearVentasNetas() {
            if(document.resumen_ventas_cli.opc_ventasnetas.checked) {
                document.resumen_ventas_cli.opc_cantidad.checked=false;
            }
            else {
                document.resumen_ventas_cli.opc_cantidad.checked=true;
            }
        }

        function keypress2() {
	        tecla=window.event.keyCode;
        }

        function keyPressed(tecla) {
        }

        function optionalFields() {
	        if (document.resumen_ventas_cli.agrupar.value=="MESES") {
		        <%if si_tiene_modulo_proyectos<>0 then%>
			        agrOtros.style.display = "none";
		        <%end if%>
		        agrMeses.style.display = "";
		        agrMeses2.style.display = "";
                agrMeses3.style.display = "";
                importIva.style.display = "none";
	        }
	        else {
		        <%if si_tiene_modulo_proyectos<>0 then%>
			        agrOtros.style.display = "";
		        <%end if%>
		        agrMeses.style.display = "none";
		        agrMeses2.style.display = "none";
                agrMeses3.style.display = "none";
                importIva.style.display = "";
	        }
	        //ricardo 22/1/2003 para que se vea o no lo de clientes en hojas separadas, ya que solo sirve para la agrupacion por cliente
	        if (document.resumen_ventas_cli.agrupar.value=="CLIENTE" || (document.resumen_ventas_cli.agrupar.value=="MESES" && document.resumen_ventas_cli.mostrarfilas.value=="CLIENTES")) {
		        agrCliHojSep.style.display="";
	        }
	        else {
		        agrCliHojSep.style.display="none";
	        }
        }

        //Desencadena la búsqueda del proveedor cuya referencia se indica
        function TraerCliente(mode,tipo) {
	        cadena="resumen_ventas_cli.asp?ncliente=" + document.resumen_ventas_cli.ncliente.value;
	        if (tipo=="1") {
		        cadena=cadena + "&acliente=" + document.resumen_ventas_cli.ncliente.value;
	        }
	        else {
		        cadena=cadena + "&acliente=" + document.resumen_ventas_cli.acliente.value;
	        }
	        document.resumen_ventas_cli.action=cadena + "&mode=" + mode;
	        document.resumen_ventas_cli.submit();
        }

        //Desencadena la búsqueda del proveedor cuya referencia se indica
        function TraerProveedor(mode,tipo) {
	        document.resumen_ventas_cli.action="resumen_ventas_cli.asp?nproveedor=" + document.resumen_ventas_cli.nproveedor.value + "&mode=" + mode;
	        document.resumen_ventas_cli.submit();
        }

        function Ver_Conceptos() {
	        if(document.resumen_ventas_cli.ver_conceptos.checked) {
		        document.resumen_ventas_cli.familia.value='';
		        document.resumen_ventas_cli.familia.disabled=true;
		        document.resumen_ventas_cli.familia_padre.value='';
		        document.resumen_ventas_cli.familia_padre.disabled=true;
		        document.resumen_ventas_cli.categoria.value='';
		        document.resumen_ventas_cli.categoria.disabled=true;
		        document.resumen_ventas_cli.referencia.value='';
		        document.resumen_ventas_cli.referencia.disabled=true;
		        document.resumen_ventas_cli.tipo_articulo.value='';
		        document.resumen_ventas_cli.tipo_articulo.disabled=true;
		        document.resumen_ventas_cli.nproveedor.value='';
		        document.resumen_ventas_cli.nproveedor.disabled=true;
		        document.resumen_ventas_cli.razon_social.value='';
		        document.resumen_ventas_cli.razon_social.disabled=true;
		        document.resumen_ventas_cli.actividad_proveedor.value='';
		        document.resumen_ventas_cli.actividad_proveedor.disabled=true;
		        document.resumen_ventas_cli.tipo_proveedor.value='';
		        document.resumen_ventas_cli.tipo_proveedor.disabled=true;
		        document.all("verc").style.display="none";
            }
            else {
		        document.resumen_ventas_cli.familia.disabled=false;
		        document.resumen_ventas_cli.familia_padre.disabled=false;
		        document.resumen_ventas_cli.categoria.disabled=false;
		        document.resumen_ventas_cli.referencia.disabled=false;
		        document.resumen_ventas_cli.tipo_articulo.disabled=false;
		        document.resumen_ventas_cli.nproveedor.disabled=false;
		        document.resumen_ventas_cli.razon_social.disabled=false;
		        document.resumen_ventas_cli.actividad_proveedor.disabled=false;
		        document.resumen_ventas_cli.tipo_proveedor.disabled=false;
		        document.all("verc").style.display="";
	        }
        }
    </script>
<body class="BODY_ASP">
        <%
        '************************ FUNCIONES ******************************************
        ''ricardo 22/1/2003
        ''para que cuando una cantidad sea 0 , que se ponga valor en blanco
        'para mostrar bien las cantidades en la agrupacion por meses
        sub mostrar_Cant_Meses(clase,opcion,valor1,valor2,ndec)
	        %><td class="<%=EncodeForHtml(clase)%>" style="border: 1px solid Black;">
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
	        </td><%
        end sub

        'crea la tabla temporal para clientes
        sub crearCliente(p_tsel1, p_tsel2, p_tsel3, p_dfecha, p_hfecha, p_tactividad, p_cliente, p_acliente, p_serie, p_familia, p_familia_padre, p_categoria, p_referencia, p_nombreart, p_ordenar, p_mb,p_opcclientebaja, p_cod_proyecto, p_ticpenfac, opc_cod_proyecto, nproveedor, actividad_proveedor, tipo_cliente, tipo_proveedor, comercial, tipo_articulo, mostrarfilas, provincia)
	        set rstFunction = Server.CreateObject("ADODB.Recordset")
	        set rstTemp = Server.CreateObject("ADODB.Recordset")
	        'creamos la tabla temporal
	
	        eliminar = "if exists (select * from sysobjects where id = object_id('egesticet.[" & session("usuario") & "]') and sysstat " & _
					         " & 0xf = 3) drop table [" & session("usuario") & "]"					    
	        rstAux.open eliminar,Session("backendlistados"),adUseClient,adLockReadOnly
		
	        crear="CREATE TABLE [" & session("usuario") & "] (NCliente varchar(10) NOT NULL ,"
	        crear=crear & "Nombre varchar(100), "
	        crear=crear & "Referencia varchar(30), "
	        crear=crear & "Descripcion varchar(255), "
	        if opc_cod_proyecto>"" then
		        crear=crear & "cod_proyecto varchar(60), "
	        end if
	        crear=crear & "Cantidad real, "
	        crear=crear & "Cantidad2 real, "
	        crear=crear & "calculoimporte bit, "
	        crear=crear & "tipo_medida varchar(50), "
	        crear=crear & "[Ventas Netas] money, "
	        crear=crear & "[Precio Medio] money, "
	        crear=crear & "[Precio Medio2] money, "
	        crear=crear & "Divisa varchar(15), "
	        crear=crear & "tiene_escv smallint, "
	        crear=crear & "Acumulador money, "
	        crear=crear & "Orden money)"
    
	        rst.open crear, Session("backendlistados"), adUseClient, adLockReadOnly

		    nclientes = 0
		    strwhere    =""
		    strwhere1    =""
		    strwhere2   =""
		    strwhere3   =""
		    strwhere4   =""
		    strwhereall =""

		    if p_dfecha>"" then
			    strwhereall=strwhereall & " and f.fecha>='" & p_dfecha & "'"
		    end if
		    if p_hfecha>"" then
			    strwhereall=strwhereall & " and f.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
		    end if
		    if p_tactividad>"" then 'Se selecciono tipo de actividad
			    strwhereall = strwhereall + " and c.tactividad='" & p_tactividad & "'"
		    end if
		    if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
			    strwhereall = strwhereall + " and f.ncliente='" & p_cliente & "'"
		    else
			    if p_opcclientebaja=1 then
				    strbaja=" "
			    else
				    strbaja=" and c.fbaja is null "
				    strwhereall = strwhereall + strbaja
			    end if
		    end if
		    if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
			    strwhereall = strwhereall + " and f.ncliente>='" & p_cliente & "'"
			    strwhereall = strwhereall + " and f.ncliente<='" & p_acliente & "'"
		    end if
            if provincia & "">"" then
                strwhereall = strwhereall + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
            end if
		    if p_serie<>"" then 'Se selecciono serie
			    if instr(p_serie,",")>0 then
				    strwhereall = strwhereall + " and f.serie in ('" & replace(replace(p_serie," ",""),",","','") & "')"
			    else
				    strwhereall = strwhereall + " and f.serie='" & p_serie & "'"
			    end if
		    end if

		    'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
		    if p_familia<>"" then
			    p_tsel2 = false
			    if instr(p_familia,",")>0 then
				    strwhere = strwhere + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
			    else
				    strwhere = strwhere + " and a.familia='" & p_familia & "'"
			    end if
		    elseif p_familia_padre<>"" then
			    p_tsel2 = false
			    if instr(p_familia_padre,",")>0 then
				    strwhere = strwhere + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
			    else
				    strwhere = strwhere + " and a.familia_padre='" & p_familia_padre & "'"
			    end if
		    elseif p_categoria<>"" then
			    p_tsel2 = false
			    if instr(p_categoria,",")>0 then
				    strwhere = strwhere + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
			    else
				    strwhere = strwhere + " and a.categoria='" & p_categoria & "'"
			    end if
		    end if
		    'FIN FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA *************************************************
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
		    if nproveedor & "">"" then
			    p_tsel2 = false
			    strwhere = strwhere + " and a.referencia in (select articulo from proveer with(NOLOCK) where nproveedor='" & nproveedor & "')"
		    end if
		    if actividad_proveedor & "">"" then
			    p_tsel2 = false
			    strwhere = strwhere + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		    end if
		    if tipo_proveedor & "">"" then
			    p_tsel2 = false
			    strwhere = strwhere + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		    end if
		    if tipo_cliente & "">"" then
			    strwhere = strwhere + " and c.tipo_cliente='" & tipo_cliente & "'"
			    strwhere2 = strwhere2 + " and c.tipo_cliente='" & tipo_cliente & "'"
		    end if

		    if comercial<>"" then 'Se selecciono comercial
			    if instr(comercial,",")>0 then
				    strwhere = strwhere + " and f.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
				    strwhere2 = strwhere2 + " and f.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
			    else
				    strwhere = strwhere + " and f.comercial='" & comercial & "'"
				    strwhere2 = strwhere2 + " and f.comercial='" & comercial & "'"
			    end if
		    end if

		    if tipo_articulo & "">"" then
			    p_tsel2 = false
			    if instr(tipo_articulo,",")>0 then
				    strwhere = strwhere + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
			    else
				    strwhere = strwhere + " and a.tipo_articulo='" & tipo_articulo & "'"
			    end if
		    end if
		    if mostrarfilas & "">"" then
			    strwhere=strwhere + ""
		    end if

		    seleccion1 = "select distinct f.ncliente as NCliente, "
   		    seleccion1=seleccion1 & "c.rsocial as Nombre, "
   		    seleccion1=seleccion1 & "d.referencia as Referencia, "
   		    seleccion1=seleccion1 & "a.nombre as Descripcion,"
		    if opc_cod_proyecto>"" then
            	    seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto, "
		    end if

   		    seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "
   		    seleccion1=seleccion1 & "sum(d.cantidad2) as Cantidad2, "
   		    seleccion1=seleccion1 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		    seleccion1=seleccion1 & "(select descripcion from medidas with(NOLOCK) where codigo=a.medidaventa) as tipo_medida,"

	        seleccion1=seleccion1 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Ventas Netas], "
	        seleccion1=seleccion1 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	        seleccion1=seleccion1 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2], "
   		    seleccion1=seleccion1 & "f.divisa as Divisa,"
		    seleccion1=seleccion1 & "0 as tiene_escv "
   		    seleccion1=seleccion1 & " from detalles_fac_cli as d with(NOLOCK), "
   		    seleccion1=seleccion1 & "facturas_cli as f with(NOLOCK), "
   		    seleccion1=seleccion1 & "articulos as a with(NOLOCK), "
   		    seleccion1=seleccion1 & "clientes as c with(NOLOCK) "
            if provincia & "">"" then
                seleccion1=seleccion1 & ",domicilios as dom with(NOLOCK) "
            end if

   		    seleccion1=seleccion1 & "where  f.nfactura like '" & session("ncliente") & "%' and d.nfactura = f.nfactura and "
   	        seleccion1=seleccion1 & "a.referencia=d.referencia and  d.nfactura like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
   		    seleccion1=seleccion1 & "c.ncliente = f.ncliente "
   		    seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
		    seleccion1=seleccion1 & " and d.mainitem is null "
            if falbaran & "">"" then
                seleccion1=seleccion1 & " and d.nalbaran is null "
            end if
   		    seleccion1=seleccion1 & "group by f.ncliente, "
   		    seleccion1=seleccion1 & "c.rsocial, "
   		    seleccion1=seleccion1 & "d.referencia, "
   		    seleccion1=seleccion1 & "a.nombre, "
		    if opc_cod_proyecto>"" then
			    seleccion1=seleccion1 & "f.cod_proyecto, "
		    end if
		    seleccion1=seleccion1 & "f.divisa "
		    seleccion1=seleccion1 & ",convert(nvarchar,a.calculoimporte) "
   		    seleccion1=seleccion1 & ",a.medidaventa"

		    seleccion1=seleccion1 & " union all "
		    seleccion1=seleccion1 & "select distinct f.ncliente as NCliente, "
   		    seleccion1=seleccion1 & "c.rsocial as Nombre, "
   		    seleccion1=seleccion1 & "d.referencia as Referencia, "
   		    seleccion1=seleccion1 & "a.nombre as Descripcion,"
		    if opc_cod_proyecto>"" then
            	    seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto, "
		    end if
   		    seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "
   		    seleccion1=seleccion1 & "sum(d.cantidad2) as Cantidad2, "
   		    seleccion1=seleccion1 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		    seleccion1=seleccion1 & "(select descripcion from medidas with(NOLOCK) where codigo=a.medidaventa) as tipo_medida,"

   		    seleccion1=seleccion1 & "0 as [Ventas Netas], "
   		    seleccion1=seleccion1 & "0 as [Precio Medio], "
		    ''ricardo 22-3-2004
   		    seleccion1=seleccion1 & "0 as [Precio Medio2], "
   		    seleccion1=seleccion1 & "f.divisa as Divisa,"
   		    seleccion1=seleccion1 & "1 as tiene_escv "
   		    seleccion1=seleccion1 & "from detalles_fac_cli as d with(NOLOCK), "
   		    seleccion1=seleccion1 & "facturas_cli as f with(NOLOCK), "
   		    seleccion1=seleccion1 & "articulos as a with(NOLOCK), "
   		    seleccion1=seleccion1 & "clientes as c with(NOLOCK) "
            if provincia & "">"" then
                seleccion1=seleccion1 & ",domicilios as dom with(NOLOCK) "
            end if
   		    seleccion1=seleccion1 & "where f.nfactura like '" & session("ncliente") & "%' and d.nfactura = f.nfactura and  d.nfactura like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
   		    seleccion1=seleccion1 & "a.referencia=d.referencia and "
   		    seleccion1=seleccion1 & "c.ncliente = f.ncliente "
   		    seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
		    seleccion1=seleccion1 & " and d.mainitem is not null "
            if falbaran & "">"" then
                seleccion1=seleccion1 & " and d.nalbaran is null "
            end if
   		    seleccion1=seleccion1 & "group by f.ncliente, "
   		    seleccion1=seleccion1 & "c.rsocial, "
   		    seleccion1=seleccion1 & "d.referencia, "
   		    seleccion1=seleccion1 & "a.nombre, "
		    if opc_cod_proyecto>"" then
			    seleccion1=seleccion1 & "f.cod_proyecto, "
		    end if
		    seleccion1=seleccion1 & "f.divisa "
   		    seleccion1=seleccion1 & ",convert(nvarchar,a.calculoimporte)"
   		    seleccion1=seleccion1 & ",a.medidaventa"

		    addselect = ""
		    addgroup  = ""
		    if desglose = true then
			    addselect = "d.descripcion as Descripcion, "
			    addgroup  = "d.descripcion, "
		    else
			    addselect =  "'Concepto' as Descripcion, "
		    end if

		    seleccion2="select distinct f.ncliente as NCliente, "
		    seleccion2=seleccion2 & "c.rsocial as Nombre, "
		    seleccion2=seleccion2 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
		    seleccion2=seleccion2 & addselect
		    if opc_cod_proyecto>"" then
			    seleccion2=seleccion2 & " f.cod_proyecto, "
		    end if
		    seleccion2=seleccion2 & "sum(d.cantidad) as Cantidad, "
   		    seleccion2=seleccion2 & "0 as Cantidad2, "
   		    seleccion2=seleccion2 & "0 as calculoimporte,"
   		    seleccion2=seleccion2 & "NULL as tipo_medida,"
	        seleccion2=seleccion2 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Ventas Netas], "
	        seleccion2=seleccion2 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	        ''ricardo 22-3-2004
	        seleccion2=seleccion2 & "0 as [Precio Medio2], "
		    seleccion2=seleccion2 & "f.divisa as Divisa,"
   		    seleccion2=seleccion2 & "0 as tiene_escv "
		    seleccion2=seleccion2 & "from conceptos as d with(NOLOCK), "
		    seleccion2=seleccion2 & "facturas_cli as f with(NOLOCK), "
		    seleccion2=seleccion2 & "clientes as c with(NOLOCK) "
            if provincia & "">"" then
                seleccion2=seleccion2 & ",domicilios as dom with(NOLOCK) "
            end if
		    seleccion2=seleccion2 & "where  f.nfactura like '" & session("ncliente") & "%' and d.nfactura = f.nfactura and   d.nfactura like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
		    seleccion2=seleccion2 & "c.ncliente = f.ncliente "
		    seleccion2=seleccion2 & strwhereall & " " & strwhere2 & " "
		    seleccion2=seleccion2 & "group by f.ncliente, "
		    seleccion2=seleccion2 & "c.rsocial, "
		    seleccion2=seleccion2 & addgroup
		    if opc_cod_proyecto>"" then
			    seleccion2=seleccion2 & "f.cod_proyecto, "
		    end if
            seleccion2=seleccion2 & "f.divisa "
		    if p_tsel3 > "" then
			    strwhereall3=""
			    if p_dfecha>"" then
		      	    strwhereall3=strwhereall3 & " and alb.fecha>='" & p_dfecha & "'"
			    end if
			    if p_hfecha>"" then
		      	    strwhereall3=strwhereall3 & " and alb.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
			    end if
			    if p_tactividad>"" then 'Se selecciono tipo de actividad
		      	    strwhereall3 = strwhereall3 + " and c.tactividad='" & p_tactividad & "'"
			    end if
			    if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
		      	    strwhereall3 = strwhereall3 + " and alb.ncliente='" & p_cliente & "'"
			    else
		      	    if p_opcclientebaja=1 then
					    strbaja=" "
				    else
					    strbaja=" and c.fbaja is null "
					    strwhereall3 = strwhereall3 + strbaja
				    end if
			    end if
			    if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
			            strwhereall3 = strwhereall3 + " and alb.ncliente>='" & p_cliente & "'"
			            strwhereall3 = strwhereall3 + " and alb.ncliente<='" & p_acliente & "'"
			    end if
                if provincia & "">"" then
                    strwhereall3 = strwhereall3 + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
                end if
			    strwhereall3 = strwhereall3 + "  and alb.serie in ('" + replace(replace(p_tsel3," ",""),",","','") + "') "

			    'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
			    if p_familia<>"" then
				    p_tsel2 = false
				    if instr(p_familia,",")>0 then
					    strwhere3 = strwhere3 + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
				    else
					    strwhere3 = strwhere3 + " and a.familia='" & p_familia & "'"
				    end if
			    elseif p_familia_padre<>"" then
				    p_tsel2 = false
				    if instr(p_familia_padre,",")>0 then
					    strwhere3 = strwhere3 + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
				    else
					    strwhere3 = strwhere3 + " and a.familia_padre='" & p_familia_padre & "'"
				    end if
			    elseif p_categoria<>"" then
				    p_tsel2 = false
				    if instr(p_categoria,",")>0 then
					    strwhere3 = strwhere3 + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
				    else
					    strwhere3 = strwhere3 + " and a.categoria='" & p_categoria & "'"
				    end if
			    end if
			    'FIN FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA *************************************************

			    if p_referencia>"" then
		      	    strwhere3 = strwhere + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
			    end if
			    if p_nombreart>"" then
		      	    strwhere3 = strwhere3 + " and a.nombre like '%" & p_nombreart & "%'"
				    strwhere4 = strwhere4 + " and d.descripcion like '%" & p_nombreart & "%'"
			    end if
			    if p_cod_proyecto>"" then
				    strwhere3=strwhere3 & " and cod_proyecto='" & p_cod_proyecto & "'"
				    strwhere4=strwhere4 & " and cod_proyecto='" & p_cod_proyecto & "'"
			    end if

			    if nproveedor & "">"" then
				    strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer with(NOLOCK) where nproveedor='" & nproveedor & "')"
			    end if
			    if actividad_proveedor & "">"" then
				    strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
			    end if
			    if tipo_proveedor & "">"" then
				    strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
			    end if
			    if tipo_cliente & "">"" then
				    strwhere3 = strwhere3 + " and c.tipo_cliente='" & tipo_cliente & "'"
				    strwhere4 = strwhere4 + " and c.tipo_cliente='" & tipo_cliente & "'"
			    end if

			    if comercial<>"" then 'Se selecciono comercial
				    if instr(comercial,",")>0 then
					    strwhere3 = strwhere3 + " and alb.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
					    strwhere4 = strwhere4 + " and alb.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
				    else
					    strwhere3 = strwhere3 + " and alb.comercial='" & comercial & "'"
					    strwhere4 = strwhere4 + " and alb.comercial='" & comercial & "'"
				    end if
			    end if

			    if tipo_articulo & "">"" then
				    if instr(tipo_articulo,",")>0 then
					    strwhere3 = strwhere3 + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
				    else
					    strwhere3 = strwhere3 + " and a.tipo_articulo='" & tipo_articulo & "'"
				    end if
			    end if
			    if mostrarfilas & "">"" then
				    strwhere3=strwhere3 + ""
			    end if

			    seleccion3= "select distinct alb.ncliente as NCliente, "
			    seleccion3=seleccion3 & "c.rsocial as Nombre, "
			    seleccion3=seleccion3 & "d.referencia as Referencia, "
			    seleccion3=seleccion3 & "a.nombre as Descripcion,"
			    if opc_cod_proyecto>"" then
                            seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
			    end if
   			    seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "
   		        seleccion3=seleccion3 & "sum(d.cantidad2) as Cantidad2, "
   		        seleccion3=seleccion3 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		        seleccion3=seleccion3 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
	            seleccion3=seleccion3 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Ventas Netas], "
	            seleccion3=seleccion3 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	            seleccion3=seleccion3 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2], "
			    seleccion3=seleccion3 & "alb.divisa as Divisa,"
	   		    seleccion3=seleccion3 & "0 as tiene_escv "
			    seleccion3=seleccion3 & "from detalles_alb_cli as d with(NOLOCK), "
			    seleccion3=seleccion3 & "albaranes_cli as alb with(NOLOCK), "
			    seleccion3=seleccion3 & "articulos as a with(NOLOCK), "
			    seleccion3=seleccion3 & "clientes as c with(NOLOCK) "
                if provincia & "">"" then
                    seleccion3=seleccion3 & ",domicilios as dom with(NOLOCK) "
                end if
			    seleccion3=seleccion3 & "where alb.nalbaran like '" & session("ncliente") & "%' and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
			    seleccion3=seleccion3 & "a.referencia=d.referencia and "
			    seleccion3=seleccion3 & "c.ncliente = alb.ncliente "
                if falbaran & ""="" then
			        seleccion3=seleccion3 & " and alb.nfactura is null "
                end if
			    seleccion3=seleccion3 & strwhereall3 & " " & strwhere3 & " "
			    seleccion3=seleccion3 & " and d.mainitem is null "
			    seleccion3=seleccion3 & " group by alb.ncliente, "
			    seleccion3=seleccion3 & "c.rsocial, "
			    seleccion3=seleccion3 & "d.referencia, "
			    seleccion3=seleccion3 & "a.nombre, "
			    if opc_cod_proyecto>"" then
				    seleccion3=seleccion3 & "alb.cod_proyecto, "
			    end if
			    seleccion3=seleccion3 & "alb.divisa "
		        ''ricardo 22-3-2004
   		        seleccion3=seleccion3 & ",convert(nvarchar,a.calculoimporte)"
   		        seleccion3=seleccion3 & ",a.medidaventa"
		        ''''''''''''''''''''''''''''
			    seleccion3=seleccion3 & " union all "
			    seleccion3=seleccion3 & "select distinct alb.ncliente as NCliente, "
			    seleccion3=seleccion3 & "c.rsocial as Nombre, "
			    seleccion3=seleccion3 & "d.referencia as Referencia, "
			    seleccion3=seleccion3 & "a.nombre as Descripcion,"
			    if opc_cod_proyecto>"" then
                            seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
			    end if
   			    seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "
		        ''ricardo 22-3-2004
   		        seleccion3=seleccion3 & "sum(d.cantidad2) as Cantidad2, "
   		        seleccion3=seleccion3 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		        seleccion3=seleccion3 & "(select descripcion from medidas with(NOLOCK) where codigo=a.medidaventa) as tipo_medida,"
		        ''''''''''''''''''''''''''''
	   		    seleccion3=seleccion3 & "0 as [Ventas Netas], "
   			    seleccion3=seleccion3 & "0 as [Precio Medio], "
			    ''ricardo 22-3-2004
   			    seleccion3=seleccion3 & "0 as [Precio Medio2], "
   			    seleccion3=seleccion3 & "alb.divisa as Divisa,"
   			    seleccion3=seleccion3 & "1 as tiene_escv "
			    seleccion3=seleccion3 & "from detalles_alb_cli as d with(NOLOCK), "
			    seleccion3=seleccion3 & "albaranes_cli as alb with(NOLOCK), "
			    seleccion3=seleccion3 & "articulos as a with(NOLOCK), "
			    seleccion3=seleccion3 & "clientes as c with(NOLOCK) "
                if provincia & "">"" then
                    seleccion3=seleccion3 & ",domicilios as dom with(NOLOCK) "
                end if
			    seleccion3=seleccion3 & "where alb.nalbaran like '" & session("ncliente") & "%' and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
			    seleccion3=seleccion3 & "a.referencia=d.referencia and "
			    seleccion3=seleccion3 & "c.ncliente = alb.ncliente "
                if falbaran & ""="" then
			        seleccion3=seleccion3 & " and alb.nfactura is null "
                end if
			    seleccion3=seleccion3 & strwhereall3 & " " & strwhere3 & " "
			    seleccion3=seleccion3 & " and d.mainitem is not null "
			    seleccion3=seleccion3 & " group by alb.ncliente, "
			    seleccion3=seleccion3 & "c.rsocial, "
			    seleccion3=seleccion3 & "d.referencia, "
			    seleccion3=seleccion3 & "a.nombre, "
			    if opc_cod_proyecto>"" then
				    seleccion3=seleccion3 & "alb.cod_proyecto, "
			    end if
			    seleccion3=seleccion3 & "alb.divisa "
  		        seleccion3=seleccion3 & ",convert(nvarchar,a.calculoimporte)"
   		        seleccion3=seleccion3 & ",a.medidaventa"

			    addselect3 = ""
			    addgroup3  = ""
			    if desglose = true then
		      	    addselect3 = "CONVERT(nvarchar,d .descripcion) as Descripcion, "
				    addgroup3  = "CONVERT(nvarchar,d .descripcion), "
			    else
				    addselect3 =  "'Concepto' as Descripcion, "
			    end if

			    seleccion4= "select distinct alb.ncliente as NCliente, "
			    seleccion4=seleccion4 & "c.rsocial as Nombre, "
			    seleccion4=seleccion4 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
			    seleccion4=seleccion4 & addselect3
			    if opc_cod_proyecto>"" then
				    seleccion4=seleccion4 & " alb.cod_proyecto, "
			    end if
                seleccion4=seleccion4 & "sum(d.cantidad) as Cantidad, "
   		        seleccion4=seleccion4 & "0 as Cantidad2, "
   		        seleccion4=seleccion4 & "0 as calculoimporte,"
   		        seleccion4=seleccion4 & "NULL as tipo_medida,"
	            seleccion4=seleccion4 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Ventas Netas], "
	            seleccion4=seleccion4 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	            seleccion4=seleccion4 & "0 as [Precio Medio2], "
			    seleccion4=seleccion4 & "alb.divisa as Divisa,"
			    seleccion4=seleccion4 & "0 as tiene_escv "
			    seleccion4=seleccion4 & "from conceptos_alb_cli as d with(NOLOCK), "
			    seleccion4=seleccion4 & "albaranes_cli as alb with(NOLOCK), "
			    seleccion4=seleccion4 & "clientes as c with(NOLOCK) "
                if provincia & "">"" then
                    seleccion4=seleccion4 & ",domicilios as dom with(NOLOCK) "
                end if
			    seleccion4=seleccion4 & "where  alb.nalbaran like '" & session("ncliente") & "%' and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' "
                if falbaran & ""="" then
			        seleccion4=seleccion4 & " and alb.nfactura is null "
                end if
			    seleccion4=seleccion4 & " and c.ncliente = alb.ncliente "
			    seleccion4=seleccion4 & strwhereall3 & " " & strwhere4 & " "
			    seleccion4=seleccion4 & "group by alb.ncliente, "
			    seleccion4=seleccion4 & "c.rsocial, "
			    seleccion4=seleccion4 & addgroup3
			    if opc_cod_proyecto>"" then
				    seleccion4=seleccion4 & "alb.cod_proyecto, "
			    end if
            	seleccion4=seleccion4 & "alb.divisa "
		    end if

	        'Se construye la seleccion de tickets pendientes de facturar si procede
	        if p_ticpenfac>"" then
		        strwhereall3=""
		        if p_dfecha>"" then
      		        strwhereall3=strwhereall3 & " and TI.fecha>='" & p_dfecha & "'"
		        end if
		        if p_hfecha>"" then
      		        strwhereall3=strwhereall3 & " and TI.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
		        end if
		        strwhereall3 = strwhereall3 & " and TI.serie in ('" & replace(replace(p_ticpenfac," ",""),",","','") & "')"
		        if p_tactividad>"" then 'Se selecciono tipo de actividad
      		        strwhereall3 = strwhereall3 + " and c.tactividad='" & p_tactividad & "'"
		        end if
		        if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
      		        strwhereall3 = strwhereall3 + " and TI.ncliente='" & p_cliente & "'"
		        else
		              if p_opcclientebaja=1 then
 				        strbaja=" "
			        else
				        strbaja=" and c.fbaja is null "
				        strwhereall3 = strwhereall3 + strbaja
			        end if
		        end if
		        if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
		              strwhereall3 = strwhereall3 + " and TI.ncliente>='" & p_cliente & "'"
		              strwhereall3 = strwhereall3 + " and TI.ncliente<='" & p_acliente & "'"
		        end if
                if provincia & "">"" then
                    strwhereall3 = strwhereall3 + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
                end if

		        'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
		        if p_familia<>"" then
			        p_tsel2 = false
			        if instr(p_familia,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia='" & p_familia & "'"
			        end if
		        elseif p_familia_padre<>"" then
			        p_tsel2 = false
			        if instr(p_familia_padre,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia_padre='" & p_familia_padre & "'"
			        end if
		        elseif p_categoria<>"" then
			        p_tsel2 = false
			        if instr(p_categoria,",")>0 then
				        strwhere3 = strwhere3 + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.categoria='" & p_categoria & "'"
			        end if
		        end if
		        'FIN FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA *************************************************

		        if p_referencia>"" then
      		        strwhere3 = strwhere + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		        end if
		        if p_nombreart>"" then
      		        strwhere3 = strwhere3 + " and a.nombre like '%" & p_nombreart & "%'"
			        strwhere4 = strwhere4 + " and d.descripcion like '%" & p_nombreart & "%'"
		        end if
		        if nproveedor & "">"" then
			        strwhereall3 = strwhereall3 + " and a.referencia in (select articulo from proveer with(NOLOCK) where nproveedor='" & nproveedor & "')"
		        end if
		        if actividad_proveedor & "">"" then
			        strwhereall3 = strwhereall3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_proveedor & "">"" then
			        strwhereall3 = strwhereall3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_cliente & "">"" then
			        strwhereall3 = strwhereall3 + " and c.tipo_cliente='" & tipo_cliente & "'"
		        end if
		        if tipo_articulo & "">"" then
                    if instr(tipo_articulo,",")>0 then
				        strwhereall3 = strwhereall3 + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
			        else
				        strwhereall3 = strwhereall3 + " and a.tipo_articulo='" & tipo_articulo & "'"
			        end if
		        end if
		        if mostrarfilas & "">"" then
			        strwhereall3 =strwhereall3 + ""
		        end if

		        seleccion5= "select distinct isnull(TI.ncliente,'-----') as NCliente, "
		        seleccion5=seleccion5 & "isnull(c.rsocial,'" & LitSinCliente & "') as Nombre, "
		        seleccion5=seleccion5 & "d.referencia as Referencia, "
		        seleccion5=seleccion5 & "a.nombre as Descripcion,"
		        if opc_cod_proyecto>"" then
			        seleccion5=seleccion5 & " '' as cod_proyecto, "
		        end if
		        seleccion5 = seleccion5 & "sum(d.cantidad) as Cantidad, "
   		        seleccion5=seleccion5 & "0 as Cantidad2, "
   		        seleccion5=seleccion5 & "0 as calculoimporte,"
   		        seleccion5=seleccion5 & "NULL as tipo_medida,"
		        seleccion5=seleccion5 & "sum((((d.importe*(100-convert(money,TI.descuento)))/100)*(100-0))/100) as [Ventas Netas], "
		        seleccion5=seleccion5 & "case when sum(d.cantidad)=0 then 0 else (sum((((d.importe*(100-convert(money,TI.descuento)))/100)*(100-0))/100)/sum(d.cantidad)) end as [Precio Medio], "
		        seleccion5=seleccion5 & "0 as [Precio Medio2], "
		        seleccion5=seleccion5 & "TI.divisa as Divisa,"
		        seleccion5=seleccion5 & "0 as tiene_escv "
		        seleccion5=seleccion5 & "from detalles_tickets as d with(NOLOCK), "
		        seleccion5=seleccion5 & "tickets as TI with(NOLOCK) left outer join clientes as c with(NOLOCK) on c.ncliente=TI.ncliente "
                if provincia & "">"" then
                    seleccion5=seleccion5 & " left outer join domicilios as dom with(NOLOCK) on dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal "
                end if
		        seleccion5=seleccion5 & ",articulos as a with(NOLOCK) "
		        seleccion5=seleccion5 & "where TI.nticket like '" & session("ncliente") & "%' and d.nticket = TI.nticket and d.nticket like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
		        seleccion5=seleccion5 & "a.referencia=d.referencia and "
		        seleccion5=seleccion5 & "TI.nfactura is null "
		        seleccion5=seleccion5 & strwhereall3 & " " & strwhere3 & " "
		        seleccion5=seleccion5 & "group by TI.ncliente, "
		        seleccion5=seleccion5 & "c.rsocial, "
		        seleccion5=seleccion5 & "d.referencia, "
		        seleccion5=seleccion5 & "a.nombre, "
		        seleccion5=seleccion5 & "TI.divisa "
	        end if

	        seleccion = ""
	        if p_tsel1=true then
		        seleccion = seleccion & seleccion1
	        end if
	        if p_tsel1=true and p_tsel2=true then
		        seleccion = seleccion & " union all "
	        end if
	        if p_tsel2 = true then
		        seleccion = seleccion & seleccion2
	        end if
	        if p_tsel3 > "" then
		        if p_tsel1=true then
			        seleccion=seleccion & " union all " & seleccion3
		        end if
		        if p_tsel2 = true then
			        seleccion=seleccion & " union all " & seleccion4
		        end if
	        end if

	        'Se añade la seleccion de tickets pendientes de faturar si procede
	        if p_ticpenfac > "" then
		        seleccion=seleccion & " union all " & seleccion5
	        end if
	        seleccion = seleccion & " order by f.ncliente"
	        if p_tsel1= true then
		        seleccion = seleccion & ", d.referencia"
	        end if

            rstFunction.cursorlocation=3
            rstFunction.open seleccion, Session("backendlistados")
	        if not rstFunction.eof then
		        acumulado = 0
		        elTotal = 0
		        cliente_anterior = ""
		        rstTemp.open "select * from egesticet.[" & session("usuario") & "]",Session("backendlistados"),adOpenKeyset,adLockOptimistic
		        clienteAnterior = ""
		        while not rstFunction.eof
			        if rstFunction("NCliente")<>clienteAnterior then
				        elTotal = elTotal + acumulado
				        acumulado = 0
			        end if
			        clienteAnterior = rstFunction("NCliente")
			        rstTemp.Addnew
			        rstTemp("Ncliente") = rstFunction("NCliente")
			        rstTemp("Nombre") = rstFunction("Nombre")
			        rstTemp("Referencia") = rstFunction("Referencia")
			        rstTemp("Descripcion") = rstFunction("Descripcion")
			        if opc_cod_proyecto>"" then
				        rstTemp("cod_proyecto") = d_lookup("nombre","proyectos","codigo='" & rstFunction("cod_proyecto") & "'",Session("backendlistados"))
			        end if
			        rstTemp("Cantidad") = rstFunction("Cantidad")
			        rstTemp("Cantidad2") = rstFunction("Cantidad2")
			        rstTemp("calculoimporte") = rstFunction("calculoimporte")
			        rstTemp("tipo_medida") = rstFunction("tipo_medida")
			        rstTemp("Ventas Netas") = rstFunction("Ventas Netas")
			        rstTemp("Precio Medio") = rstFunction("Precio Medio")
			        rstTemp("Precio Medio2") = rstFunction("Precio Medio2")
			        rstTemp("Divisa") = rstFunction("Divisa")
			        if rstTemp("Divisa") = p_mb then
				        rstTemp("Acumulador") = acumulado + rstTemp("Ventas Netas")
				        acumulado = rstTemp("Acumulador")
			        else
				        rstTemp("Acumulador") = acumulado + CambioDivisa(rstTemp("Ventas Netas"), rstTemp("Divisa"), p_mb)
				        acumulado = rstTemp("Acumulador")
			        end if
			        rstTemp("tiene_escv")=rstFunction("tiene_escv")
			        rstTemp("Orden") = 0
			        rstTemp.Update
			        rstFunction.movenext
		        wend
		        rstFunction.close
		        rstTemp.close

		        'ordenamos por ventas si es necesario
		        if ordenar=true then
                    rstFunction.cursorlocation=3
 		            rstFunction.open "select distinct ncliente from [" & session("usuario") & "]", Session("backendlistados")
	      	        while not rstFunction.eof
				        rstTemp.open "update [" & session("usuario") & "] set orden = (select max(acumulador) from [" & _
			                      session("usuario") & "] where ncliente = '" & rstFunction("Ncliente") & _
						          "') where ncliente ='" & rstFunction("Ncliente") & "'", Session("backendlistados"), 1, 3
				        rstFunction.movenext
			        wend
			        rstFunction.close
		        end if
	        end if

	        elTotal = elTotal + acumulado
	        %>
            <input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' /><%
	        set rstFunction = Nothing
	        set rstTemp = Nothing
        end sub
        '-----------------------------------------------------

        ' crea la tabla temporal para articulos
        sub crearArticulo(p_tsel1, p_tsel2,p_tsel3, p_dfecha, p_hfecha, p_familia, p_familia_padre, p_categoria, p_referencia, p_nombreart, p_tactividad, p_cliente,p_acliente, p_serie, p_desglose, p_ordenar, p_mb,p_opcclientebaja,p_cod_proyecto,p_ticpenfac,opc_cod_proyecto,nproveedor,actividad_proveedor,tipo_cliente,tipo_proveedor,comercial,tipo_articulo,mostrarfilas,provincia)', p_importeIva)
	        set rstFunction = Server.CreateObject("ADODB.Recordset")
	        set rstTemp = Server.CreateObject("ADODB.Recordset")
	        'creamos la tabla temporal
	        DropTable session("usuario"), Session("backendlistados")
	        crear ="CREATE TABLE [" & session("usuario") & "] (Ref varchar(30) NOT NULL ,"
	        crear=crear & "Descripcion varchar(100), "
	        crear=crear & "Ncliente varchar(10), "
	        crear=crear & "Nombre varchar(100), "
	        if opc_cod_proyecto>"" then
		        crear=crear & "cod_proyecto varchar(60), "
	        end if
	        crear=crear & "Cantidad real, "
	        crear=crear & "Cantidad2 real, "
	        crear=crear & "calculoimporte bit, "
	        crear=crear & "tipo_medida varchar(50), "
	        crear=crear & "[Ventas Netas] money, "
	        crear=crear & "[Precio Medio] money, "
	        crear=crear & "[Precio Medio2] money, "
	        crear=crear & "Divisa varchar(15), "
	        crear=crear & "tiene_escv smallint,"
	        crear=crear & "AcumulaVentas money, "
	        crear=crear & "AcumulaCantidad real, "
	        crear=crear & "Orden money)"

	        rst.open crear,Session("backendlistados"),adUseClient,adLockReadOnly
	        GrantUser session("usuario"), Session("backendlistados")

	        strwhere     = ""
	        strwhereall  = ""
	        strwhere2    = ""

	        if p_dfecha>"" then
		        strwhereall=strwhereall & " and f.fecha>='" & p_dfecha & "'"
	        end if
	        if p_hfecha>"" then
		        strwhereall=strwhereall & " and f.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
	        end if
	        PorArticulo = "SI"
	        'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
	        if p_familia<>"" then
		        p_tsel2 = false
		        if instr(p_familia,",")>0 then
			        strwhere = strwhere + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.familia='" & p_familia & "'"
		        end if
	        elseif p_familia_padre<>"" then
		        p_tsel2 = false
		        if instr(p_familia_padre,",")>0 then
			        strwhere = strwhere + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.familia_padre='" & p_familia_padre & "'"
		        end if
	        elseif p_categoria<>"" then
		        p_tsel2 = false
		        if instr(p_categoria,",")>0 then
			        strwhere = strwhere + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.categoria='" & p_categoria & "'"
		        end if
	        end if
	        if p_referencia>"" then
		        strwhere = strwhere + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		        tsel2=false
	        end if
	        if p_nombreart>"" then
      	        strwhere = strwhere + " and a.nombre like '%" & p_nombreart & "%'"
		        strwhere2 = strwhere2 + " and d.descripcion like '%" & p_nombreart & "%'"
	        end if
	        if p_tactividad>"" then 'Se selecciono tipo de actividad
		        strwhereall = strwhereall + " and c.tactividad='" & p_tactividad & "'"
	        end if
	        if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
		        strwhereall = strwhereall + " and f.ncliente='" & p_cliente & "'"
	        else
		        if p_opcclientebaja=1 then
			        strbaja=" "
		        else
			        strbaja=" and c.fbaja is null "
			        strwhereall = strwhereall + strbaja
		        end if
            end if
	        if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
      	        strwhereall = strwhereall + " and f.ncliente>='" & p_cliente & "'"
	              strwhereall = strwhereall + " and f.ncliente<='" & p_acliente & "'"
	        end if
            if provincia & "">"" then
                strwhereall = strwhereall + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
            end if
	        if p_cod_proyecto>"" then
		        strwhere=strwhere & " and cod_proyecto='" & p_cod_proyecto & "'"
		        strwhere2=strwhere2 & " and cod_proyecto='" & p_cod_proyecto & "'"
	        end if
	        if p_serie<>"" then 'Se selecciono serie
		        if instr(p_serie,",")>0 then
			        strwhereall = strwhereall + " and f.serie in ('" & replace(replace(p_serie," ",""),",","','") & "')"
		        else
			        strwhereall = strwhereall + " and f.serie='" & p_serie & "'"
		        end if
	        end if
	        if nproveedor & "">"" then
		        tsel2 = false
		        strwhere = strwhere + " and a.referencia in (select articulo from proveer where nproveedor='" & nproveedor & "')"
	        end if
	        if actividad_proveedor & "">"" then
		        tsel2 = false
		        strwhere = strwhere + " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
	        end if
	        if tipo_proveedor & "">"" then
		        tsel2 = false
		        strwhere = strwhere + " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
	        end if
	        if tipo_cliente & "">"" then
		        strwhere = strwhere + " and c.tipo_cliente='" & tipo_cliente & "'"
		        strwhere2 = strwhere2 + " and c.tipo_cliente='" & tipo_cliente & "'"
	        end if
	        if comercial<>"" then 'Se selecciono comercial
		        if instr(comercial,",")>0 then
			        strwhere = strwhere + " and f.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
			        strwhere2 = strwhere2 + " and f.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and f.comercial='" & comercial & "'"
			        strwhere2 = strwhere2 + " and f.comercial='" & comercial & "'"
		        end if
	        end if
	        if tipo_articulo & "">"" then
		        tsel2 = false
                if instr(tipo_articulo,",")>0 then
			        strwhere = strwhere + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.tipo_articulo='" & tipo_articulo & "'"
		        end if
	        end if
	        if mostrarfilas & "">"" then
		        strwhere=strwhere & ""
	        end if

	        seleccion1="select distinct d.referencia as Ref, "
	        seleccion1=seleccion1 & "a.nombre as Descripcion, "
	        seleccion1=seleccion1 & "f.ncliente as NCliente,"
	        seleccion1=seleccion1 & "c.rsocial as Nombre, "
	        if opc_cod_proyecto>"" then
	           seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto, "
	        end if
              seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad,"
   	        seleccion1=seleccion1 & "sum(d.cantidad2) as Cantidad2, "
   	        seleccion1=seleccion1 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   	        seleccion1=seleccion1 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
	        seleccion1=seleccion1 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Ventas Netas], "
	        seleccion1=seleccion1 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	        seleccion1=seleccion1 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2], "
	        seleccion1=seleccion1 & "f.divisa as Divisa,"
	        seleccion1=seleccion1 & "0 as tiene_escv "
	        seleccion1=seleccion1 & "from detalles_fac_cli as d with(NOLOCK), "
	        seleccion1=seleccion1 & "facturas_cli as f with(NOLOCK), "
	        seleccion1=seleccion1 & "articulos as a with(NOLOCK), "
	        seleccion1=seleccion1 & "clientes as c with(NOLOCK) "
            if provincia & "">"" then
                seleccion1=seleccion1 & ",domicilios as dom with(NOLOCK) "
            end if
	        seleccion1=seleccion1 & "where f.nfactura like '" & session("ncliente") & "%' and d.nfactura = f.nfactura and d.nfactura like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
	        seleccion1=seleccion1 & "a.referencia = d.referencia and "
	        seleccion1=seleccion1 & "c.ncliente = f.ncliente "
            if falbaran & "">"" then
                seleccion1=seleccion1 & " and d.nalbaran is null "
            end if
	        seleccion1=seleccion1 & strwhereall & strwhere & " "
	        seleccion1=seleccion1 & "and d.mainitem is null "
	        seleccion1=seleccion1 & "group by d.referencia, "
	        seleccion1=seleccion1 & "a.nombre, "
	        if opc_cod_proyecto>"" then
		        seleccion1=seleccion1 & "f.cod_proyecto, "
	        end if
              seleccion1=seleccion1 & "f.ncliente, "
	        seleccion1=seleccion1 & "c.rsocial, "
	        seleccion1=seleccion1 & "f.divisa "
	        seleccion1=seleccion1 & ",convert(nvarchar,a.calculoimporte) "
   	        seleccion1=seleccion1 & ",a.medidaventa"
	        seleccion1=seleccion1 & " union all "
	        seleccion1=seleccion1 & "select distinct d.referencia as Ref, "
	        seleccion1=seleccion1 & "a.nombre as Descripcion, "
	        seleccion1=seleccion1 & "f.ncliente as NCliente,"
	        seleccion1=seleccion1 & "c.rsocial as Nombre, "
	        if opc_cod_proyecto>"" then
	           seleccion1=seleccion1 & "f.cod_proyecto as cod_proyecto, "
	        end if
              seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad,"
   	        seleccion1=seleccion1 & "sum(d.cantidad2) as Cantidad2, "
   	        seleccion1=seleccion1 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   	        seleccion1=seleccion1 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
	        seleccion1=seleccion1 & "0 as [Ventas Netas],"
	        seleccion1=seleccion1 & "0 as [Precio Medio],"
   	        seleccion1=seleccion1 & "0 as [Precio Medio2], "
	        seleccion1=seleccion1 & "f.divisa as Divisa,"
	        seleccion1=seleccion1 & "1 as tiene_escv "
	        seleccion1=seleccion1 & "from detalles_fac_cli as d with(NOLOCK), "
	        seleccion1=seleccion1 & "facturas_cli as f with(NOLOCK), "
	        seleccion1=seleccion1 & "articulos as a with(NOLOCK), "
	        seleccion1=seleccion1 & "clientes as c with(NOLOCK) "
            if provincia & "">"" then
                seleccion1=seleccion1 & ",domicilios as dom with(NOLOCK) "
            end if
	        seleccion1=seleccion1 & "where f.nfactura like '" & session("ncliente") & "%' and d.nfactura = f.nfactura and d.nfactura like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
	        seleccion1=seleccion1 & "a.referencia = d.referencia and "
	        seleccion1=seleccion1 & "c.ncliente = f.ncliente "
	        seleccion1=seleccion1 & strwhereall & strwhere & " "
	        seleccion1=seleccion1 & "and d.mainitem is not null "
            if falbaran & "">"" then
                seleccion1=seleccion1 & " and d.nalbaran is null "
            end if
	        seleccion1=seleccion1 & "group by d.referencia, "
	        seleccion1=seleccion1 & "a.nombre, "
	        if opc_cod_proyecto>"" then
		        seleccion1=seleccion1 & "f.cod_proyecto, "
	        end if
              seleccion1=seleccion1 & "f.ncliente, "
	        seleccion1=seleccion1 & "c.rsocial, "
	        seleccion1=seleccion1 & "f.divisa "
   	        seleccion1=seleccion1 & ",convert(nvarchar,a.calculoimporte)"
   	        seleccion1=seleccion1 & ",a.medidaventa"
	        
            cabecera = ""
	        if p_desglose=false then
      	        cabecera="select distinct 'zzzzzzzzzzz' as Ref, "
      	        cabecera=cabecera & "'@concepto@' as Descripcion, "
	        else
		        cabecera="select distinct d.descripcion as Ref, "
      	        cabecera=cabecera & "'@concepto@' as Descripcion, "
	        end if

	        seleccion2 = cabecera & "f.ncliente as NCliente, "
    	        seleccion2=seleccion2 & "c.rsocial as Nombre, "
	        if opc_cod_proyecto>"" then
		        seleccion2=seleccion2 & "f.cod_proyecto as cod_proyecto,"
	        end if
	        seleccion2=seleccion2 & "sum(d.cantidad) as Cantidad, "
   	        seleccion2=seleccion2 & "0 as Cantidad2, "
   	        seleccion2=seleccion2 & "0 as calculoimporte,"
   	        seleccion2=seleccion2 & "NULL as tipo_medida,"
	        seleccion2=seleccion2 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Ventas Netas], "
	        seleccion2=seleccion2 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	        seleccion2=seleccion2 & "0 as [Precio Medio2], "
            seleccion2=seleccion2 & "f.divisa as Divisa,"
	        seleccion2=seleccion2 & "0 as tiene_escv "
            seleccion2=seleccion2 & "from conceptos as d with(NOLOCK), "
            seleccion2=seleccion2 & "facturas_cli as f with(NOLOCK), "
            seleccion2=seleccion2 & "clientes as c with(NOLOCK) "
            if provincia & "">"" then
                seleccion2=seleccion2 & ",domicilios as dom with(NOLOCK) "
            end if
            seleccion2=seleccion2 & "where f.nfactura like '" & session("ncliente") & "%' and d.nfactura = f.nfactura and d.nfactura like '" & session("ncliente") & "%'  and c.ncliente like '" & session("ncliente") & "%' and "
            seleccion2=seleccion2 & "c.ncliente = f.ncliente "
	        seleccion2=seleccion2 & strwhereall & strwhere2
	        if desglose = false then
		        seleccion2=seleccion2 & "group by f.ncliente, "
		        seleccion2=seleccion2 & "c.rsocial, "
		        if opc_cod_proyecto>"" then
			        seleccion2=seleccion2 & "f.cod_proyecto, "
		        end if
                    seleccion2=seleccion2 & "f.divisa "
	        else
		        seleccion2=seleccion2 & "group by d.descripcion, "
		        if opc_cod_proyecto>"" then
			        seleccion2 = seleccion2 & "f.cod_proyecto, "
		        end if
		        seleccion2=seleccion2 & "f.ncliente, "
                    seleccion2=seleccion2 & "c.rsocial, "
                    seleccion2=seleccion2 & "f.divisa "
	        end if

	        if p_tsel3 > "" then
		        strwhereall3=""
		        if p_dfecha>"" then
			        strwhereall3=strwhereall3 & " and alb.fecha>='" & p_dfecha & "'"
		        end if
		        if p_hfecha>"" then
			        strwhereall3=strwhereall3 & " and alb.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
		        end if
		        strwhereall3 = strwhereall3 & "  and alb.serie in ('" + replace(replace(p_tsel3," ",""),",","','") + "') "

		        PorArticulo = "SI"

		        'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
		        if p_familia<>"" then
			        p_tsel2 = false
			        if instr(p_familia,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia='" & p_familia & "'"
			        end if
		        elseif p_familia_padre<>"" then
			        p_tsel2 = false
			        if instr(p_familia_padre,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia_padre='" & p_familia_padre & "'"
			        end if
		        elseif p_categoria<>"" then
			        p_tsel2 = false
			        if instr(p_categoria,",")>0 then
				        strwhere3 = strwhere3 + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.categoria='" & p_categoria & "'"
			        end if
		        end if

		        if p_referencia>"" then
			        strwhere3 = strwhere3 + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		        end if
		        if p_nombreart>"" then
			        strwhere3 = strwhere3 + " and a.nombre like '%" & p_nombreart & "%'"
		        end if
		        if p_tactividad>"" then 'Se selecciono tipo de actividad
			        strwhereall3 = strwhereall3 + " and c.tactividad='" & p_tactividad & "'"
		        end if
		        if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
			        strwhereall3 = strwhereall3 + " and alb.ncliente='" & p_cliente & "'"
		        else
			        if p_opcclientebaja=1 then
				        strbaja=" "
			        else
				        strbaja=" and c.fbaja is null "
				        strwhereall3 = strwhereall3 + strbaja
			        end if
      	        end if
		        if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
      		        strwhereall3 = strwhereall3 + " and alb.ncliente>='" & p_cliente & "'"
	      	        strwhereall3 = strwhereall3 + " and alb.ncliente<='" & p_acliente & "'"
		        end if
                if provincia & "">"" then
                    strwhereall3 = strwhereall3 + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
                end if
		        if p_cod_proyecto>"" then
			        strwhere3=strwhere3 & " and cod_proyecto='" & p_cod_proyecto & "'"
			        strwhere4=strwhere4 & " and cod_proyecto='" & p_cod_proyecto & "'"
		        end if
		        if nproveedor & "">"" then
			        strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer where nproveedor='" & nproveedor & "')"
		        end if
		        if actividad_proveedor & "">"" then
			        strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_proveedor & "">"" then
			        strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_cliente & "">"" then
			        strwhere3 = strwhere3 + " and c.tipo_cliente='" & tipo_cliente & "'"
			        strwhere4 = strwhere4 + " and c.tipo_cliente='" & tipo_cliente & "'"
		        end if
		        if comercial<>"" then 'Se selecciono comercial
			        if instr(comercial,",")>0 then
				        strwhere3 = strwhere3 + " and alb.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
				        strwhere4 = strwhere4 + " and alb.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and alb.comercial='" & comercial & "'"
				        strwhere4 = strwhere4 + " and alb.comercial='" & comercial & "'"
			        end if
		        end if
		        if tipo_articulo & "">"" then
                    if instr(tipo_articulo,",")>0 then
			            strwhere3 = strwhere3 + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
		            else
			            strwhere3 = strwhere3 + " and a.tipo_articulo='" & tipo_articulo & "'"
		            end if
		        end if
		        if mostrarfilas & "">"" then
			        strwhere3=strwhere3 + ""
		        end if

		        seleccion3="select distinct d.referencia as Ref, "
		        seleccion3=seleccion3 & "a.nombre as Descripcion, "
		        seleccion3=seleccion3 & "alb.ncliente as NCliente,"
		        seleccion3=seleccion3 & "c.rsocial as Nombre, "
		        if opc_cod_proyecto>"" then
			        seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
		        end if
                    seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad,"
   		        seleccion3=seleccion3 & "sum(d.cantidad2) as Cantidad2, "
   		        seleccion3=seleccion3 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		        seleccion3=seleccion3 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
	            seleccion3=seleccion3 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Ventas Netas], "
	            seleccion3=seleccion3 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	            seleccion3=seleccion3 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2], "
		        seleccion3=seleccion3 & "alb.divisa as Divisa,"
		        seleccion3=seleccion3 & "0 as tiene_escv "
		        seleccion3=seleccion3 & "from detalles_alb_cli as d with(NOLOCK), "
		        seleccion3=seleccion3 & "albaranes_cli as alb with(NOLOCK), "
		        seleccion3=seleccion3 & "articulos as a with(NOLOCK), "
		        seleccion3=seleccion3 & "clientes as c with(NOLOCK) "
                if provincia & "">"" then
                    seleccion3=seleccion3 & ",domicilios as dom with(NOLOCK) "
                end if
		        seleccion3=seleccion3 & "where alb.nalbaran like '" & session("ncliente") & "%' and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
		        seleccion3=seleccion3 & "a.referencia = d.referencia "
                if falbaran & ""="" then
		            seleccion3=seleccion3 & " and alb.nfactura is null "
                end if
		        seleccion3=seleccion3 & " and c.ncliente = alb.ncliente "
		        seleccion3=seleccion3 & strwhereall3 & strwhere3 & " "
		        seleccion3=seleccion3 & "and d.mainitem is null "
		        seleccion3=seleccion3 & "group by d.referencia, "
		        seleccion3=seleccion3 & "a.nombre, "
		        if opc_cod_proyecto>"" then
			        seleccion3=seleccion3 & "alb.cod_proyecto, "
		        end if
                    seleccion3=seleccion3 & "alb.ncliente, "
     		        seleccion3=seleccion3 & "c.rsocial, "
		        seleccion3=seleccion3 & "alb.divisa "
   		        seleccion3=seleccion3 & ",convert(nvarchar,a.calculoimporte)"
   		        seleccion3=seleccion3 & ",a.medidaventa"
		        seleccion3=seleccion3 & " union all "
		        seleccion3=seleccion3 & "select distinct d.referencia as Ref, "
		        seleccion3=seleccion3 & "a.nombre as Descripcion, "
		        seleccion3=seleccion3 & "alb.ncliente as NCliente,"
		        seleccion3=seleccion3 & "c.rsocial as Nombre, "
		        if opc_cod_proyecto>"" then
			        seleccion3=seleccion3 & "alb.cod_proyecto as cod_proyecto, "
		        end if
                    seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad,"
   		        seleccion3=seleccion3 & "sum(d.cantidad2) as Cantidad2, "
   		        seleccion3=seleccion3 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		        seleccion3=seleccion3 & "(select descripcion from medidas with(NOLOCK) where codigo=a.medidaventa) as tipo_medida,"
		        seleccion3=seleccion3 & "0 as [Ventas Netas],"
		        seleccion3=seleccion3 & "0 as [Precio Medio],"
   		        seleccion3=seleccion3 & "0 as [Precio Medio2], "
		        seleccion3=seleccion3 & "alb.divisa as Divisa,"
		        seleccion3=seleccion3 & "1 as tiene_escv "
		        seleccion3=seleccion3 & "from detalles_alb_cli as d with(NOLOCK), "
		        seleccion3=seleccion3 & "albaranes_cli as alb with(NOLOCK), "
		        seleccion3=seleccion3 & "articulos as a with(NOLOCK), "
		        seleccion3=seleccion3 & "clientes as c with(NOLOCK) "
                if provincia & "">"" then
                    seleccion3=seleccion3 & ",domicilios as dom with(NOLOCK) "
                end if
		        seleccion3=seleccion3 & "where alb.nalbaran like '" & session("ncliente") & "%' and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
		        seleccion3=seleccion3 & "a.referencia = d.referencia "
                if falbaran & ""="" then
		            seleccion3=seleccion3 & " and alb.nfactura is null "
                end if
		        seleccion3=seleccion3 & " and c.ncliente = alb.ncliente "
		        seleccion3=seleccion3 & strwhereall3 & strwhere3 & " "
		        seleccion3=seleccion3 & "and d.mainitem is not null "
		        seleccion3=seleccion3 & "group by d.referencia, "
		        seleccion3=seleccion3 & "a.nombre, "
		        if opc_cod_proyecto>"" then
			        seleccion3=seleccion3 & "alb.cod_proyecto, "
		        end if
                seleccion3=seleccion3 & "alb.ncliente, "
     		    seleccion3=seleccion3 & "c.rsocial, "
		        seleccion3=seleccion3 & "alb.divisa "
  		        seleccion3=seleccion3 & ",convert(nvarchar,a.calculoimporte)"
   		        seleccion3=seleccion3 & ",a.medidaventa"

		        cabecera = ""
		        if p_desglose=false then
      		        cabecera4 = "select distinct 'zzzzzzzzzzz' as Ref, " & _
                          "'@concepto@' as Descripcion, "
		        else
			        cabecera4 = "select distinct convert(nvarchar,d.descripcion) as Ref, " & _
                          "'@concepto@' as Descripcion, "
		        end if

		        seleccion4=cabecera4 & "alb.ncliente as NCliente, "
		        seleccion4=seleccion4 & "c.rsocial as Nombre, "
		        if opc_cod_proyecto>"" then
			        seleccion4=seleccion4 & "alb.cod_proyecto as cod_proyecto,"
		        end if
		        seleccion4=seleccion4 & "sum(d.cantidad) as Cantidad, "
   		        seleccion4=seleccion4 & "0 as Cantidad2, "
   		        seleccion4=seleccion4 & "0 as calculoimporte,"
   		        seleccion4=seleccion4 & "NULL as tipo_medida,"
	            seleccion4=seleccion4 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Ventas Netas], "
	            seleccion4=seleccion4 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	            seleccion4=seleccion4 & "0 as [Precio Medio2], "
                seleccion4=seleccion4 & "alb.divisa as Divisa,"
		        seleccion4=seleccion4 & "0 as tiene_escv "
                seleccion4=seleccion4 & "from conceptos_alb_cli as d with(NOLOCK), "
                seleccion4=seleccion4 & "albaranes_cli as alb with(NOLOCK), "
                seleccion4=seleccion4 & "clientes as c with(NOLOCK) "
                if provincia & "">"" then
                    seleccion4=seleccion4 & ",domicilios as dom with(NOLOCK) "
                end if
                seleccion4=seleccion4 & "where alb.nalbaran like '" & session("ncliente") & "%' and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%'  and c.ncliente like '" & session("ncliente") & "%' "
                if falbaran &""="" then
		            seleccion4=seleccion4 & " and alb.nfactura is null "
                end if
                seleccion4=seleccion4 & " and c.ncliente = alb.ncliente "
		        seleccion4=seleccion4 & strwhereall3 & strwhere4
		        if desglose = false then
			        seleccion4=seleccion4 & "group by alb.ncliente, "
                    seleccion4=seleccion4 & "c.rsocial, "
			        if opc_cod_proyecto>"" then
				        seleccion4=seleccion4 & "alb.cod_proyecto, "
			        end if
	                    seleccion4=seleccion4 & "alb.divisa "
		        else
			        seleccion4=seleccion4 & "group by convert(nvarchar,d.descripcion), "
			        if opc_cod_proyecto>"" then
				        seleccion4=seleccion4 & "alb.cod_proyecto, "
			        end if
			        seleccion4=seleccion4 & "alb.ncliente, "
                    seleccion4=seleccion4 & "c.rsocial, "
                    seleccion4=seleccion4 & "alb.divisa "
		        end if
	        end if

	        'Se construye la seleccion de tickets pendientes de facturar si procede
	        if p_ticpenfac > "" then
		        strwhereall3=""
		        if p_dfecha>"" then
			        strwhereall3=strwhereall3 & " and TI.fecha>='" & p_dfecha & "'"
		        end if
		        if p_hfecha>"" then
			        strwhereall3=strwhereall3 & " and TI.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
		        end if
		        strwhereall3=strwhereall3 & " and TI.serie in ('" & replace(replace(p_ticpenfac," ",""),",","','") & "')"
		        PorArticulo = "SI"

		        'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
		        if p_familia<>"" then
			        p_tsel2 = false
			        if instr(p_familia,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia='" & p_familia & "'"
			        end if
		        elseif p_familia_padre<>"" then
			        p_tsel2 = false
			        if instr(p_familia_padre,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia_padre='" & p_familia_padre & "'"
			        end if
		        elseif p_categoria<>"" then
			        p_tsel2 = false
			        if instr(p_categoria,",")>0 then
				        strwhere3 = strwhere3 + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.categoria='" & p_categoria & "'"
			        end if
		        end if
		        if p_referencia>"" then
			        strwhere3=strwhere3 + " and d.referencia like '" & session("ncliente") & "%" & p_referencia & "%'"
		        end if
		        if p_nombreart>"" then
			        strwhere3=strwhere3 + " and a.nombre like '%" & p_nombreart & "%'"
		        end if
		        if p_tactividad>"" then 'Se selecciono tipo de actividad
			        strwhereall3=strwhereall3 + " and c.tactividad='" & p_tactividad & "'"
		        end if
		        if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
			        strwhereall3=strwhereall3 + " and TI.ncliente='" & p_cliente & "'"
		        else
			        if p_opcclientebaja=1 then
				        strbaja=" "
			        else
				        strbaja=" and c.fbaja is null "
				        strwhereall3=strwhereall3 + strbaja
			        end if
      	        end if
		        if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
      		        strwhereall3=strwhereall3 + " and TI.ncliente>='" & p_cliente & "'"
	      	        strwhereall3=strwhereall3 + " and TI.ncliente<='" & p_acliente & "'"
		        end if
                if provincia & "">"" then
                    strwhereall3 = strwhereall3 + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
                end if
		        if nproveedor & "">"" then
			        strwhereall3 = strwhereall3 + " and a.referencia in (select articulo from proveer with(NOLOCK) where nproveedor='" & nproveedor & "')"
		        end if
		        if actividad_proveedor & "">"" then
			        strwhereall3 = strwhereall3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_proveedor & "">"" then
			        strwhereall3 = strwhereall3 + " and a.referencia in (select articulo from proveer as pr with(NOLOCK),proveedores as pro with(NOLOCK) where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_cliente & "">"" then
			        strwhereall3 = strwhereall3 + " and c.tipo_cliente='" & tipo_cliente & "'"
		        end if
		        if tipo_articulo & "">"" then
                    if instr(tipo_articulo,",")>0 then
			            strwhereall3 = strwhereall3 + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
		            else
			            strwhereall3 = strwhereall3 + " and a.tipo_articulo='" & tipo_articulo & "'"
		            end if
		        end if
		        if mostrarfilas & "">"" then
			        strwhereall3 =strwhereall3 + ""
		        end if

		        seleccion5 = "select distinct d.referencia as Ref, "
                seleccion5=seleccion5 & "a.nombre as Descripcion, "
                seleccion5=seleccion5 & "isnull(TI.ncliente,'-----') as NCliente,"
                seleccion5=seleccion5 & "isnull(c.rsocial,'" & LitSinCliente & "') as Nombre, "
		        if opc_cod_proyecto>"" then
			        seleccion5=seleccion5 & "'' as cod_proyecto,"
		        end if
		        seleccion5=seleccion5 & "sum(d.cantidad) as Cantidad,"
   		        seleccion5=seleccion5 & "0 as Cantidad2, "
   		        seleccion5=seleccion5 & "0 as calculoimporte,"
   		        seleccion5=seleccion5 & "NULL as tipo_medida,"
                seleccion5=seleccion5 & "sum((((d.importe*(100-convert(money,TI.descuento)))/100)*(100-0))/100) as [Ventas Netas],"
                seleccion5=seleccion5 & "case when sum(d.cantidad)=0 then 0 else (sum((((d.importe*(100-convert(money,TI.descuento)))/100)*(100-0))/100)/sum(d.cantidad)) end as [Precio Medio],"
		        seleccion5=seleccion5 & "0 as [Precio Medio2], "
		        seleccion5=seleccion5 & "TI.divisa as Divisa,"
		        seleccion5=seleccion5 & "0 as tiene_escv "
                seleccion5=seleccion5 & "from detalles_tickets as d with(NOLOCK), "
                seleccion5=seleccion5 & "tickets as TI with(NOLOCK) left outer join clientes c with(NOLOCK) on c.ncliente=TI.ncliente "
                if provincia & "">"" then
                    seleccion5=seleccion5 & " left outer join domicilios dom with(NOLOCK) on dom.codigo=c.dir_principal "
                end if
                seleccion5=seleccion5 & ",articulos as a with(NOLOCK) "
                seleccion5=seleccion5 & "where TI.nticket like '" & session("ncliente") & "%' and d.nticket = TI.nticket and d.nticket like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
                seleccion5=seleccion5 & "a.referencia = d.referencia and "
		        seleccion5=seleccion5 & "TI.nfactura is null "
		        seleccion5=seleccion5 & strwhereall3 & strwhere3 & " "
                seleccion5=seleccion5 & "group by d.referencia, "
                seleccion5=seleccion5 & "a.nombre, "
		        seleccion5=seleccion5 & "TI.ncliente, "
                seleccion5=seleccion5 & "c.rsocial, "
		        seleccion5=seleccion5 & "TI.divisa "
	        end if

	        seleccion = ""
	        if p_tsel1=true then
		        seleccion = seleccion & seleccion1
	        end if
	        if p_tsel1=true and p_tsel2=true then
	           seleccion = seleccion & " union all "
	        end if
	        if p_tsel2 = true then
	           seleccion = seleccion & seleccion2
	        end if

	        if p_tsel3 > "" then
		        if p_tsel1=true then
			        seleccion=seleccion & " union all " & seleccion3
		        end if
		        if p_tsel2 = true then
			        seleccion=seleccion & " union all " & seleccion4
		        end if
	        end if

	        'Se añade la seleccion de tickets pendientes de faturar si procede
	        if p_ticpenfac > "" then
		        seleccion=seleccion & " union all " & seleccion5
	        end if

	        porventas = ""

	        if p_tsel1=true then
		        seleccion = seleccion & " order by " & porventas & "d.referencia, f.ncliente"
	        else
		        if p_desglose = false then
 	      	        seleccion = seleccion & " order by " & porventas & "f.ncliente"
		        else
			        seleccion = seleccion & " order by " & porventas & "d.descripcion, f.ncliente"
		        end if
	        end if

	        rstFunction.open seleccion, Session("backendlistados"),1,3
	        if not rstFunction.eof then
      	        acumuladoPasta    = 0
		        acumuladoCantidad = 0
		        elTotal = 0
		        rstTemp.open "select * from [" & session("usuario") & "]", Session("backendlistados"),adOpenKeyset,adLockOptimistic
		        articuloAnterior = ""
		        while not rstFunction.eof
	    		    if ucase(rstFunction("Ref"))<>ucase(articuloAnterior) then
				        elTotal = elTotal + acumuladoPasta
				        acumuladoPasta    = 0
				        acumuladoCantidad = 0
			        end if
			        articuloAnterior = ucase(rstFunction("Ref"))
			        rstTemp.Addnew
			        rstTemp("Ref") = mid(rstFunction("Ref"),1,25)
			        rstTemp("Descripcion") = rstFunction("Descripcion")
			        if opc_cod_proyecto>"" then
				        rstTemp("cod_proyecto") = d_lookup("nombre","proyectos","codigo='" & rstFunction("cod_proyecto") & "'",Session("backendlistados"))
			        end if
			        rstTemp("NCliente") = rstFunction("NCliente")
			        rstTemp("Nombre") = rstFunction("Nombre")
			        rstTemp("Cantidad") = rstFunction("Cantidad")
			        rstTemp("Cantidad2") = rstFunction("Cantidad2")
			        rstTemp("calculoimporte") = rstFunction("calculoimporte")
			        rstTemp("tipo_medida") = rstFunction("tipo_medida")

			        rstTemp("Ventas Netas") = rstFunction("Ventas Netas")
			        rstTemp("Precio Medio") = rstFunction("Precio Medio")
			        rstTemp("Precio Medio2") = rstFunction("Precio Medio2")
			        rstTemp("Divisa") = rstFunction("Divisa")
			        if rstTemp("Divisa") = p_mb then
				        rstTemp("AcumulaVentas") = acumuladoPasta + rstTemp("Ventas Netas")
				        acumuladoPasta = rstTemp("AcumulaVentas")
			        else
				        rstTemp("AcumulaVentas") = acumuladoPasta + CambioDivisa(rstTemp("Ventas Netas"), rstTemp("Divisa"), p_mb)
				        acumuladoPasta = rstTemp("AcumulaVentas")
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

		        'ordenamos por ventas si es necesario
		        if ordenar=true then
                    rstFunction.cursorlocation=3
 	      	        rstFunction.open "select distinct ref from [" & session("usuario") & "]", Session("backendlistados")
		            while not rstFunction.eof
				         rstTemp.open "update [" & session("usuario") & "] set orden = (select max(acumulaventas) from [" & _
			                      session("usuario") & "] where ref = '" & rstFunction("Ref") & _
						          "') where Ref ='" & rstFunction("Ref") & "'", Session("backendlistados"), 1, 3
				        rstFunction.movenext
			        wend
			        rstFunction.close
		        end if
	        end if

	        elTotal = elTotal + acumuladoPasta
	        %><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' /><%

	        set rstFunction = Nothing
	        set rstTemp = Nothing
        end sub

        'crea la tabla temporal para proyectos
        sub crearProyecto (p_tsel1, p_tsel2,p_tsel3, p_dfecha, p_hfecha, p_tactividad, p_cliente,p_acliente, p_serie, p_familia, p_familia_padre, p_categoria, p_referencia, p_nombreart, p_ordenar, p_mb,p_opcclientebaja,p_cod_proyecto,p_ticpenfac,nproveedor,actividad_proveedor,tipo_cliente,tipo_proveedor,comercial,tipo_articulo,mostrarfilas,provincia)
	        set rstFunction = Server.CreateObject("ADODB.Recordset")
	        set rstTemp = Server.CreateObject("ADODB.Recordset")
	        'creamos la tabla temporal
	        DropTable session("usuario"), Session("backendlistados")
	        crear="CREATE TABLE [" & session("usuario") & "] (cod_proyecto varchar(60) not null,NCliente varchar(10) ,"
	        crear=crear & "Nombre varchar(100), "
	        crear=crear & "Referencia varchar(30), "
	        crear=crear & "Descripcion varchar(255), "
	        crear=crear & "Cantidad real, "
	        crear=crear & "Cantidad2 real, "
	        crear=crear & "calculoimporte bit, "
	        crear=crear & "tipo_medida varchar(50), "
	        crear=crear & "[Ventas Netas] money, "
	        crear=crear & "[Precio Medio] money, "
	        crear=crear & "[Precio Medio2] money, "
	        crear=crear & "Divisa varchar(15), "
	        crear=crear & "tiene_escv smallint,"
	        crear=crear & "Acumulador money, "
	        crear=crear & "Orden money)"

	        rst.open crear,Session("backendlistados"),adUseClient,adLockReadOnly
	        GrantUser session("usuario"), Session("backendlistados")

	        nclientes = 0
	        strwhere    =""
	        strwhere2   =""
	        strwhereall =""

	        if p_dfecha>"" then
		        strwhereall=strwhereall & " and f.fecha>='" & p_dfecha & "'"
	        end if
	        if p_hfecha>"" then
		        strwhereall=strwhereall & " and f.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
	        end if
	        if p_tactividad>"" then 'Se selecciono tipo de actividad
		        strwhereall = strwhereall + " and c.tactividad='" & p_tactividad & "'"
	        end if
	        if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
		        strwhereall = strwhereall + " and f.ncliente='" & p_cliente & "'"
	        else
		        if p_opcclientebaja=1 then
			        strbaja=" "
		        else
			        strbaja=" and c.fbaja is null "
			        strwhereall = strwhereall + strbaja
		        end if
	        end if
	        if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
      	        strwhereall = strwhereall + " and f.ncliente>='" & p_cliente & "'"
	              strwhereall = strwhereall + " and f.ncliente<='" & p_acliente & "'"
	        end if
            if provincia & "">"" then
                strwhereall = strwhereall + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
            end if
	        if p_serie<>"" then 'Se selecciono serie
		        if instr(p_serie,",")>0 then
			        strwhereall = strwhereall + " and f.serie in ('" & replace(replace(p_serie," ",""),",","','") & "')"
		        else
			        strwhereall = strwhereall + " and f.serie='" & p_serie & "'"
		        end if
	        end if
	        'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
	        if p_familia<>"" then
		        p_tsel2 = false
		        if instr(p_familia,",")>0 then
			        strwhere = strwhere + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.familia='" & p_familia & "'"
		        end if
	        elseif p_familia_padre<>"" then
		        p_tsel2 = false
		        if instr(p_familia_padre,",")>0 then
			        strwhere = strwhere + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.familia_padre='" & p_familia_padre & "'"
		        end if
	        elseif p_categoria<>"" then
		        p_tsel2 = false
		        if instr(p_categoria,",")>0 then
			        strwhere = strwhere + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.categoria='" & p_categoria & "'"
		        end if
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
	        if nproveedor & "">"" then
		        p_tsel2 = false
		        strwhere = strwhere + " and a.referencia in (select articulo from proveer where nproveedor='" & nproveedor & "')"
	        end if
	        if actividad_proveedor & "">"" then
		        p_tsel2 = false
		        strwhere = strwhere + " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
	        end if
	        if tipo_proveedor & "">"" then
		        p_tsel2 = false
		        strwhere = strwhere + " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
	        end if
	        if tipo_cliente & "">"" then
		        strwhere = strwhere + " and c.tipo_cliente='" & tipo_cliente & "'"
		        strwhere2 = strwhere2 + " and c.tipo_cliente='" & tipo_cliente & "'"
	        end if
	        if comercial<>"" then 'Se selecciono comercial
		        if instr(comercial,",")>0 then
			        strwhere = strwhere + " and f.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
			        strwhere2 = strwhere2 + " and f.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and f.comercial='" & comercial & "'"
			        strwhere2 = strwhere2 + " and f.comercial='" & comercial & "'"
		        end if
	        end if
	        if tipo_articulo & "">"" then
		        p_tsel2 = false
                if instr(tipo_articulo,",")>0 then
			        strwhere = strwhere + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
		        else
			        strwhere = strwhere + " and a.tipo_articulo='" & tipo_articulo & "'"
		        end if
	        end if
	        if mostrarfilas & "">"" then
		        strwhere=strwhere & ""
	        end if

	        seleccion1="select distinct f.cod_proyecto as cod_proyecto,f.ncliente as NCliente, "
	        seleccion1=seleccion1 & "c.rsocial as Nombre, "
	        seleccion1=seleccion1 & "d.referencia as Referencia, "
	        seleccion1=seleccion1 & "a.nombre as Descripcion, "
	        seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "
   	        seleccion1=seleccion1 & "sum(d.cantidad2) as Cantidad2, "
   	        seleccion1=seleccion1 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   	        seleccion1=seleccion1 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
	        seleccion1=seleccion1 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Ventas Netas], "
	        seleccion1=seleccion1 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	        seleccion1=seleccion1 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2], "
	        seleccion1=seleccion1 & "f.divisa as Divisa,"
	        seleccion1=seleccion1 & "0 as tiene_escv "
	        seleccion1=seleccion1 & "from detalles_fac_cli as d, "
	        seleccion1=seleccion1 & "facturas_cli as f, "
	        seleccion1=seleccion1 & "articulos as a, "
	        seleccion1=seleccion1 & "clientes as c,proyectos as p "
            if provincia & "">"" then
                seleccion1=seleccion1 & ",domicilios as dom "
            end if
	        seleccion1=seleccion1 & "where f.nfactura like '" & session("ncliente") & "%' and p.codigo=f.cod_proyecto and d.nfactura = f.nfactura and d.nfactura like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
	        seleccion1=seleccion1 & "a.referencia=d.referencia and "
	        seleccion1=seleccion1 & "c.ncliente = f.ncliente "
	        seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
	        seleccion1=seleccion1 & " and d.mainitem is null "
            if falbaran & "">"" then
                seleccion1=seleccion1 & " and d.nalbaran is null "
            end if
	        seleccion1=seleccion1 & "group by f.cod_proyecto,f.ncliente, "
	        seleccion1=seleccion1 & "c.rsocial, "
	        seleccion1=seleccion1 & "d.referencia, "
	        seleccion1=seleccion1 & "a.nombre, "
	        seleccion1=seleccion1 & "f.divisa "
	        seleccion1=seleccion1 & ",convert(nvarchar,a.calculoimporte) "
   	        seleccion1=seleccion1 & ",a.medidaventa"
	        seleccion1=seleccion1 & " union all "
	        seleccion1=seleccion1 & "select distinct f.cod_proyecto as cod_proyecto,f.ncliente as NCliente, "
	        seleccion1=seleccion1 & "c.rsocial as Nombre, "
	        seleccion1=seleccion1 & "d.referencia as Referencia, "
	        seleccion1=seleccion1 & "a.nombre as Descripcion, "
	        seleccion1=seleccion1 & "sum(d.cantidad) as Cantidad, "
   	        seleccion1=seleccion1 & "sum(d.cantidad2) as Cantidad2, "
   	        seleccion1=seleccion1 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   	        seleccion1=seleccion1 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
	        seleccion1=seleccion1 & "0 as [Ventas Netas], "
	        seleccion1=seleccion1 & "0 as [Precio Medio], "
	        seleccion1=seleccion1 & "0 as [Precio Medio2], "
	        seleccion1=seleccion1 & "f.divisa as Divisa,"
	        seleccion1=seleccion1 & "1 as tiene_escv "
	        seleccion1=seleccion1 & "from detalles_fac_cli as d, "
	        seleccion1=seleccion1 & "facturas_cli as f, "
	        seleccion1=seleccion1 & "articulos as a, "
	        seleccion1=seleccion1 & "clientes as c,proyectos as p "
            if provincia & "">"" then
                seleccion1=seleccion1 & ",domicilios as dom "
            end if
	        seleccion1=seleccion1 & "where f.nfactura like '" & session("ncliente") & "%' and p.codigo=f.cod_proyecto and d.nfactura = f.nfactura and d.nfactura like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
	        seleccion1=seleccion1 & "a.referencia=d.referencia and "
	        seleccion1=seleccion1 & "c.ncliente = f.ncliente "
	        seleccion1=seleccion1 & strwhereall & " " & strwhere & " "
	        seleccion1=seleccion1 & " and d.mainitem is not null "
            if falbaran & "">"" then
                seleccion1=seleccion1 & " and d.nalbaran is null "
            end if
	        seleccion1=seleccion1 & "group by f.cod_proyecto,f.ncliente, "
	        seleccion1=seleccion1 & "c.rsocial, "
	        seleccion1=seleccion1 & "d.referencia, "
	        seleccion1=seleccion1 & "a.nombre, "
	        seleccion1=seleccion1 & "f.divisa "
   	        seleccion1=seleccion1 & ",convert(nvarchar,a.calculoimporte)"
   	        seleccion1=seleccion1 & ",a.medidaventa"

	        addselect = ""
	        addgroup  = ""
	        if desglose = true then
		        addselect = "d.descripcion as Descripcion, "
		        addgroup  = "d.descripcion, "
	        else
		        addselect =  "'Concepto' as Descripcion, "
	        end if

	        seleccion2="select distinct f.cod_proyecto,f.ncliente as NCliente, "
	        seleccion2=seleccion2 & "c.rsocial as Nombre, "
	        seleccion2=seleccion2 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
	        seleccion2=seleccion2 & addselect
	        seleccion2=seleccion2 & "sum(d.cantidad) as Cantidad, "
   	        seleccion2=seleccion2 & "0 as Cantidad2, "
   	        seleccion2=seleccion2 & "0 as calculoimporte,"
   	        seleccion2=seleccion2 & "NULL as tipo_medida,"
	        seleccion2=seleccion2 & "sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100) as [Ventas Netas], "
	        seleccion2=seleccion2 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,f.descuento),0)))/100)*(100-isnull(convert(money,f.descuento2),0)))/100)*(100-isnull(convert(money,f.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	        seleccion2=seleccion2 & "0 as [Precio Medio2], "
	        seleccion2=seleccion2 & "f.divisa as Divisa,"
	        seleccion2=seleccion2 & "0 as tiene_escv "
	        seleccion2=seleccion2 & "from conceptos as d, "
	        seleccion2=seleccion2 & "facturas_cli as f, "
	        seleccion2=seleccion2 & "clientes as c,proyectos as p "
            if provincia & "">"" then
                seleccion2=seleccion2 & ",domicilios as dom "
            end if
	        seleccion2=seleccion2 & "where f.nfactura like '" & session("ncliente") & "%' and p.codigo=f.cod_proyecto and d.nfactura = f.nfactura and d.nfactura like '" & session("ncliente") & "%'  and c.ncliente like '" & session("ncliente") & "%' and "
	        seleccion2=seleccion2 & "c.ncliente = f.ncliente "
	        seleccion2=seleccion2 & strwhereall & " " & strwhere2 & " "
	        seleccion2=seleccion2 & "group by f.cod_proyecto,f.ncliente, "
	        seleccion2=seleccion2 & "c.rsocial, "
	        seleccion2=seleccion2 & addgroup
	        seleccion2=seleccion2 & "f.divisa "

	        if p_tsel3 > "" then
		        strwhereall3=""
		        if p_dfecha>"" then
			        strwhereall3=strwhereall3 & " and alb.fecha>='" & p_dfecha & "'"
		        end if
		        if p_hfecha>"" then
			        strwhereall3=strwhereall3 & " and alb.fecha<=convert(datetime,'" & p_hfecha & " 23:59:59')"
		        end if
		        if p_tactividad>"" then 'Se selecciono tipo de actividad
			        strwhereall3 = strwhereall3 + " and c.tactividad='" & p_tactividad & "'"
		        end if
		        if p_cliente>"" and p_acliente>"" and p_cliente=p_acliente then 'Se selecciono cliente
			        strwhereall3 = strwhereall3 + " and alb.ncliente='" & p_cliente & "'"
		        else
			        if p_opcclientebaja=1 then
				        strbaja=" "
			        else
				        strbaja=" and c.fbaja is null "
				        strwhereall3 = strwhereall3 + strbaja
			        end if
		        end if
                if provincia & "">"" then
                    strwhereall3 = strwhereall3 + " and dom.pertenece like '" & session("ncliente") & "%' and dom.codigo=c.dir_principal and dom.provincia like '%" & provincia & "%' "
                end if
		        if p_cliente>"" and p_acliente>"" and p_cliente<>p_acliente then
      		        strwhereall3 = strwhereall3 + " and alb.ncliente>='" & p_cliente & "'"
	      	        strwhereall3 = strwhereall3 + " and alb.ncliente<='" & p_acliente & "'"
		        end if
		        if p_serie<>"" then 'Se selecciono serie
			        if instr(p_serie,",")>0 then
				        strwhereall3 = strwhereall3 + " and alb.serie in ('" & replace(replace(p_serie," ",""),",","','") & "')"
			        else
				        strwhereall3 = strwhereall3 + " and alb.serie='" & p_serie & "'"
			        end if
		        end if

		        'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************
		        if p_familia<>"" then
			        p_tsel2 = false
			        if instr(p_familia,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia in ('" & replace(replace(p_familia," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia='" & p_familia & "'"
			        end if
		        elseif p_familia_padre<>"" then
			        p_tsel2 = false
			        if instr(p_familia_padre,",")>0 then
				        strwhere3 = strwhere3 + " and a.familia_padre in ('" & replace(replace(p_familia_padre," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.familia_padre='" & p_familia_padre & "'"
			        end if
		        elseif p_categoria<>"" then
			        p_tsel2 = false
			        if instr(p_categoria,",")>0 then
				        strwhere3 = strwhere3 + " and a.categoria in ('" & replace(replace(p_categoria," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and a.categoria='" & p_categoria & "'"
			        end if
		        end if
		        'FIN FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 13/05/2005 ****************************************

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
		        if nproveedor & "">"" then
			        strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer where nproveedor='" & nproveedor & "')"
		        end if
		        if actividad_proveedor & "">"" then
			        strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tactividad='" & actividad_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_proveedor & "">"" then
			        strwhere3 = strwhere3 + " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor='" & tipo_proveedor & "' and pr.nproveedor like '" & session("ncliente") & "%' and pro.nproveedor like '" & session("ncliente") & "%')"
		        end if
		        if tipo_cliente & "">"" then
			        strwhere3 = strwhere3 + " and c.tipo_cliente='" & tipo_cliente & "'"
			        strwhere4 = strwhere4 + " and c.tipo_cliente='" & tipo_cliente & "'"
		        end if
		        if comercial<>"" then 'Se selecciono comercial
			        if instr(comercial,",")>0 then
				        strwhere3 = strwhere3 + " and alb.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
				        strwhere4 = strwhere4 + " and alb.comercial in ('" & replace(replace(comercial," ",""),",","','") & "')"
			        else
				        strwhere3 = strwhere3 + " and alb.comercial='" & comercial & "'"
				        strwhere4 = strwhere4 + " and alb.comercial='" & comercial & "'"
			        end if
		        end if
		        if tipo_articulo & "">"" then
                    if instr(tipo_articulo,",")>0 then
			            strwhere3 = strwhere3 + " and a.tipo_articulo in ('" & replace(replace(tipo_articulo," ",""),",","','") & "')"
		            else
			            strwhere3 = strwhere3 + " and a.tipo_articulo='" & tipo_articulo & "'"
		            end if
		        end if
		        if mostrarfilas & "">"" then
			        strwhere3=strwhere3 + ""
		        end if

		        seleccion3="select distinct alb.cod_proyecto as cod_proyecto,alb.ncliente as NCliente, "
		        seleccion3=seleccion3 & "c.rsocial as Nombre, "
		        seleccion3=seleccion3 & "d.referencia as Referencia, "
		        seleccion3=seleccion3 & "a.nombre as Descripcion, "
		        seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "
   		        seleccion3=seleccion3 & "sum(d.cantidad2) as Cantidad2, "
   		        seleccion3=seleccion3 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		        seleccion3=seleccion3 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
	            seleccion3=seleccion3 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Ventas Netas], "
	            seleccion3=seleccion3 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	            seleccion3=seleccion3 & "case when sum(d.cantidad2)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad2)) end as [Precio Medio2], "
		        seleccion3=seleccion3 & "alb.divisa as Divisa,"
		        seleccion3=seleccion3 & "0 as tiene_escv "
		        seleccion3=seleccion3 & "from detalles_alb_cli as d, "
		        seleccion3=seleccion3 & "albaranes_cli as alb, "
		        seleccion3=seleccion3 & "articulos as a, "
		        seleccion3=seleccion3 & "clientes as c,proyectos as p "
                if provincia & "">"" then
                    seleccion3=seleccion3 & ",domicilios as dom "
                end if
		        seleccion3=seleccion3 & "where alb.nalbaran like '" & session("ncliente") & "%' and p.codigo=alb.cod_proyecto and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
		        seleccion3=seleccion3 & "a.referencia=d.referencia and "
		        seleccion3=seleccion3 & "c.ncliente = alb.ncliente "
		        seleccion3=seleccion3 &	strwhereall3 & " " & strwhere3 & " "
		        seleccion3=seleccion3 &	" and d.mainitem is null "
                if falbaran & ""="" then
                    seleccion3=seleccion3 & " and alb.nfactura is null "
                end if
		        seleccion3=seleccion3 & "group by alb.cod_proyecto,alb.ncliente, "
		        seleccion3=seleccion3 & "c.rsocial, "
		        seleccion3=seleccion3 & "d.referencia, "
		        seleccion3=seleccion3 & "a.nombre, "
		        seleccion3=seleccion3 & "alb.divisa "
   		        seleccion3=seleccion3 & ",convert(nvarchar,a.calculoimporte)"
   		        seleccion3=seleccion3 & ",a.medidaventa"
		        seleccion3=seleccion3 & " union all "
		        seleccion3=seleccion3 & "select distinct alb.cod_proyecto as cod_proyecto,alb.ncliente as NCliente, "
		        seleccion3=seleccion3 & "c.rsocial as Nombre, "
		        seleccion3=seleccion3 & "d.referencia as Referencia, "
		        seleccion3=seleccion3 & "a.nombre as Descripcion, "
		        seleccion3=seleccion3 & "sum(d.cantidad) as Cantidad, "
   		        seleccion3=seleccion3 & "sum(d.cantidad2) as Cantidad2, "
   		        seleccion3=seleccion3 & "convert(nvarchar,a.calculoimporte) as calculoimporte,"
   		        seleccion3=seleccion3 & "(select descripcion from medidas where codigo=a.medidaventa) as tipo_medida,"
		        seleccion3=seleccion3 & "0 as [Ventas Netas], "
		        seleccion3=seleccion3 & "0 as [Precio Medio], "
		        seleccion3=seleccion3 & "0 as [Precio Medio2], "
		        seleccion3=seleccion3 & "alb.divisa as Divisa,"
		        seleccion3=seleccion3 & "1 as tiene_escv "
		        seleccion3=seleccion3 & "from detalles_alb_cli as d, "
		        seleccion3=seleccion3 & "albaranes_cli as alb, "
		        seleccion3=seleccion3 & "articulos as a, "
		        seleccion3=seleccion3 & "clientes as c,proyectos as p "
                if provincia & "">"" then
                    seleccion3=seleccion3 & ",domicilios as dom "
                end if
		        seleccion3=seleccion3 & "where alb.nalbaran like '" & session("ncliente") & "%' and p.codigo=alb.cod_proyecto and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and a.referencia like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' and "
		        seleccion3=seleccion3 & "a.referencia=d.referencia and "
		        seleccion3=seleccion3 & "c.ncliente = alb.ncliente "
		        seleccion3=seleccion3 &	strwhereall3 & " " & strwhere3 & " "
		        seleccion3=seleccion3 &	" and d.mainitem is not null "
                if falbaran & ""="" then
                    seleccion3=seleccion3 & " and alb.nfactura is null "
                end if
		        seleccion3=seleccion3 & "group by alb.cod_proyecto,alb.ncliente, "
		        seleccion3=seleccion3 & "c.rsocial, "
		        seleccion3=seleccion3 & "d.referencia, "
		        seleccion3=seleccion3 & "a.nombre, "
		        seleccion3=seleccion3 & "alb.divisa "
  		        seleccion3=seleccion3 & ",convert(nvarchar,a.calculoimporte)"
   		        seleccion3=seleccion3 & ",a.medidaventa"

		        addselect3 = ""
		        addgroup3  = ""
		        if desglose = true then
			        addselect3 = "CONVERT(nvarchar,d .descripcion) as Descripcion, "
			        addgroup3  = "CONVERT(nvarchar,d .descripcion), "
		        else
			        addselect3 =  "'Concepto' as Descripcion, "
		        end if

		        seleccion4="select distinct alb.cod_proyecto,alb.ncliente as NCliente, "
		        seleccion4=seleccion4 & "c.rsocial as Nombre, "
		        seleccion4=seleccion4 & "'zzzzzzzzzzzzzzzzzzzz' as Referencia, "
		        seleccion4=seleccion4 & addselect3
		        seleccion4=seleccion4 & "sum(d.cantidad) as Cantidad, "
   		        seleccion4=seleccion4 & "0 as Cantidad2, "
   		        seleccion4=seleccion4 & "0 as calculoimporte,"
   		        seleccion4=seleccion4 & "NULL as tipo_medida,"

	            seleccion4=seleccion4 & "sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100) as [Ventas Netas], "
	            seleccion4=seleccion4 & "case when sum(d.cantidad)=0 then 0 else (sum((((((d.importe*(100-isnull(convert(money,alb.descuento),0)))/100)*(100-isnull(convert(money,alb.descuento2),0)))/100)*(100-isnull(convert(money,alb.descuento3),0)))/100)/sum(d.cantidad)) end as [Precio Medio], "
	            seleccion4=seleccion4 & "0 as [Precio Medio2], "
		        seleccion4=seleccion4 & "alb.divisa as Divisa,"
		        seleccion4=seleccion4 & "0 as tiene_escv "
		        seleccion4=seleccion4 & "from conceptos_alb_cli as d, "
		        seleccion4=seleccion4 & "albaranes_cli as alb, "
		        seleccion4=seleccion4 & "clientes as c,proyectos as p "
                if provincia & "">"" then
                    seleccion4=seleccion4 & ",domicilios as dom "
                end if
		        seleccion4=seleccion4 & "where alb.nalbaran like '" & session("ncliente") & "%' and p.codigo=alb.cod_proyecto and d.nalbaran = alb.nalbaran and d.nalbaran like '" & session("ncliente") & "%' and c.ncliente like '" & session("ncliente") & "%' "
                if falbaran &""="" then
                    seleccion4=seleccion4 & "alb.nfactura is null and "
                end if
		        seleccion4=seleccion4 & " and c.ncliente = alb.ncliente "
		        seleccion4=seleccion4 & strwhereall3 & " " & strwhere4 & " "
		        seleccion4=seleccion4 & "group by alb.cod_proyecto,alb.ncliente, "
		        seleccion4=seleccion4 & "c.rsocial, "
		        seleccion4=seleccion4 & addgroup3
		        seleccion4=seleccion4 & "alb.divisa "
	        end if

	        seleccion = ""
	        if p_tsel1=true then
		        seleccion = seleccion & seleccion1
	        end if
	        if p_tsel1=true and p_tsel2=true then
		        seleccion = seleccion & " union all "
	        end if
	        if p_tsel2 = true then
		        seleccion = seleccion & seleccion2
	        end if

	        if p_tsel3 > "" then
		        if p_tsel1=true then
			        seleccion=seleccion & " union all " & seleccion3
		        end if
		        if p_tsel2 = true then
			        seleccion=seleccion & " union all " & seleccion4
		        end if
	        end if

	        seleccion = seleccion & " order by f.cod_proyecto,f.ncliente"
	        if p_tsel1= true then seleccion = seleccion & ", d.referencia"

	        rstFunction.open seleccion, Session("backendlistados"),1,3
	        if not rstFunction.eof then
		        acumulado = 0
		        elTotal = 0
                cod_proyectoAnterior = ""
		        rstTemp.open "select * from [" & session("usuario") & "]", Session("backendlistados"),adOpenKeyset,adLockOptimistic
		        cod_proyectoAnterior = ""
		        while not rstFunction.eof
			        if ucase(rstFunction("cod_proyecto"))<>ucase(cod_proyectoAnterior) then
				        elTotal = elTotal + acumulado
				        acumulado = 0
			        end if
			        cod_proyectoAnterior = ucase(rstFunction("cod_proyecto"))
			        rstTemp.Addnew
			        rstTemp("cod_proyecto") = d_lookup("nombre","proyectos","codigo='" & rstFunction("cod_proyecto") & "'",Session("backendlistados"))
			        rstTemp("Ncliente") = rstFunction("NCliente")
			        rstTemp("Nombre") = rstFunction("Nombre")
			        rstTemp("Referencia") = rstFunction("Referencia")
			        rstTemp("Descripcion") = rstFunction("Descripcion")
			        rstTemp("Cantidad") = rstFunction("Cantidad")
			        rstTemp("Cantidad2") = rstFunction("Cantidad2")
			        rstTemp("calculoimporte") = rstFunction("calculoimporte")
			        rstTemp("tipo_medida") = rstFunction("tipo_medida")
			        rstTemp("Ventas Netas") = rstFunction("Ventas Netas")
			        rstTemp("Precio Medio") = rstFunction("Precio Medio")
			        rstTemp("Precio Medio2") = rstFunction("Precio Medio2")
			        rstTemp("Divisa") = rstFunction("Divisa")
			        if rstTemp("Divisa") = p_mb then
				        rstTemp("Acumulador") = acumulado + rstTemp("Ventas Netas")
				        acumulado = rstTemp("Acumulador")
			        else
				        rstTemp("Acumulador") = acumulado + CambioDivisa(rstTemp("Ventas Netas"), rstTemp("Divisa"), p_mb)
				        acumulado = rstTemp("Acumulador")
			        end if
			        rstTemp("tiene_escv")=rstFunction("tiene_escv")
			        rstTemp("Orden") = 0
			        rstTemp.Update
			        rstFunction.movenext
		        wend
		        rstFunction.close
		        rstTemp.close

		        'ordenamos por ventas si es necesario
		        if ordenar=true then
                    rstFunction.cursorlocation=3
			        rstFunction.open "select distinct cod_proyecto from [" & session("usuario") & "]", Session("backendlistados")
			        while not rstFunction.eof
				        rstTemp.open "update [" & session("usuario") & "] set orden = (select max(acumulador) from [" & _
			                      session("usuario") & "] where cod_proyecto = '" & rstFunction("cod_proyecto") & _
						          "') where cod_proyecto ='" & rstFunction("cod_proyecto") & "'", Session("backendlistados"), 1, 3
				        rstFunction.movenext
			        wend
			        rstFunction.close
		        end if
	        end if

	        elTotal = elTotal + acumulado
	        %><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' /><%
        end sub

        '*****************************************************************************
        '********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
        '*****************************************************************************
        const borde=0

	    set connRound = Server.CreateObject("ADODB.Connection")
	    connRound.open dsnilion
        %>
        <form name="resumen_ventas_cli" method="post">
        <%
		    PintarCabecera "resumen_ventas_cli.asp"
		    'Leer parámetros de la página
		    SumaAnt = 0
		    SubTotalAnt = 0
  		    mode=Request.QueryString("mode")
  		    campo=limpiaCadena(Request.QueryString("campo"))
  		    criterio=limpiaCadena(Request.QueryString("criterio"))
  		    texto=limpiaCadena(Request.QueryString("texto"))
		    elTotal = limpiaCadena(Request.form("elTotal"))

	        if request.querystring("verdc")>"" then
		        verdc=limpiaCadena(request.querystring("verdc"))
	        else
		        verdc=limpiaCadena(request.form("verdc"))
	        end if
	
	        if verdc & "" = "" then
	            verdc=limpiaCadena(request.querystring("ndoc"))
	        end if

	        %><input type="hidden" name="verdc" value="<%=EncodeForHtml(verdc)%>" /><%

	        si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
	        si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	        si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)

	        ncliente = limpiaCadena(Request.QueryString("ncliente"))
	        if ncliente = "" then
		        ncliente = limpiaCadena(Request.form("ncliente"))
	        end if
	        acliente = limpiaCadena(Request.QueryString("acliente"))
	        if acliente = "" then
		        acliente = limpiaCadena(Request.form("acliente"))
	        end if
	        provincia = limpiaCadena(Request.QueryString("provincia"))
	        if provincia = "" then
		        provincia = limpiaCadena(Request.form("provincia"))
	        end if
	        nproveedor = limpiaCadena(Request.QueryString("nproveedor"))
	        if nproveedor = "" then
		        nproveedor = limpiaCadena(Request.form("nproveedor"))
	        end if
	        actividad = limpiaCadena(Request.QueryString("actividad"))
	        if actividad = "" then
		        actividad = limpiaCadena(Request.form("actividad"))
	        end if
	        fdesde = limpiaCadena(Request.QueryString("fdesde"))
	        if fdesde = "" then
		        fdesde = limpiaCadena(Request.form("fdesde"))
	        end if
	        fhasta = limpiaCadena(Request.QueryString("fhasta"))
	        if fhasta = "" then
		        fhasta = limpiaCadena(Request.form("fhasta"))
	        end if
	        nserie = limpiaCadena(Request.QueryString("nserie"))
	        if nserie = "" then
		        nserie = limpiaCadena(Request.form("nserie"))
	        end if
	        serie=nserie
	        seriesTPF = limpiaCadena(Request.QueryString("seriesTPF"))
	        if seriesTPF = "" then
		        seriesTPF = limpiaCadena(Request.form("seriesTPF"))
	        end if
	        falbaran = limpiaCadena(Request.QueryString("fechaalbaran"))
	        if falbaran = "" then
		        falbaran = limpiaCadena(Request.form("fechaalbaran"))
	        end if
    	    tactividad = limpiaCadena(Request.QueryString("tactividad"))
	        if tactividad = "" then
		        tactividad = limpiaCadena(Request.form("tactividad"))
	        end if
	        referencia = limpiaCadena(Request.QueryString("referencia"))
	        if referencia = "" then
		        referencia = limpiaCadena(Request.form("referencia"))
	        end if
	        nombreart = limpiaCadena(Request.QueryString("nombreart"))
	        if nombreart = "" then
		        nombreart = limpiaCadena(Request.form("nombreart"))
	        end if
	        familia = limpiaCadena(Request.QueryString("familia"))
	        if familia = "" then
		        familia = limpiaCadena(Request.form("familia"))
	        end if
	        familia_padre = limpiaCadena(Request.QueryString("familia_padre"))
	        if familia_padre = "" then
		        familia_padre= limpiaCadena(Request.form("familia_padre"))
	        end if
	        categoria = limpiaCadena(Request.QueryString("categoria"))
	        if categoria = "" then
		        categoria = limpiaCadena(Request.form("categoria"))
	        end if
	        agrupar	= limpiaCadena(Request.QueryString("agrupar"))
	        if agrupar = "" then
		        agrupar	= limpiaCadena(Request.form("agrupar"))
	        end if
    	    conceptos = limpiaCadena(Request.QueryString("conceptos"))
	        if conceptos = "" then
		        conceptos = limpiaCadena(Request.form("conceptos"))
	        end if
            'INICIO AÑADIR IMPORTE MEDIO IVA
            importeMedioIva	= limpiaCadena(Request.QueryString("importeMedioIva"))
	        if importeMedioIva ="" then
		        importeMedioIva	= limpiaCadena(Request.form("importeMedioIva"))
	        end if
            'FIN AÑADIR IMPORTE MEDIO IVA
	        ver_conceptos = limpiaCadena(Request.QueryString("ver_conceptos"))
	        if ver_conceptos = "" then
		        ver_conceptos = limpiaCadena(Request.form("ver_conceptos"))
	        end if
  	        set_orden = limpiaCadena(Request.QueryString("ordenar_ventas"))
	        if set_orden = "" then
		        set_orden = limpiaCadena(Request.form("ordenar_ventas"))
	        end if
	        if request.form("opcclientebaja") > "" then
		        opcclientebaja=limpiaCadena(request.form("opcclientebaja"))
	        else
		        opcclientebaja=limpiaCadena(request.querystring("opcclientebaja"))
	        end if
	        seriesAPF = limpiaCadena(Request.QueryString("seriesAPF"))
	        if seriesAPF = "" then
		        seriesAPF = limpiaCadena(Request.form("seriesAPF"))
	        end if
	        cod_proyecto = limpiaCadena(Request.QueryString("cod_proyecto"))
	        if cod_proyecto = "" then
		        cod_proyecto = limpiaCadena(Request.form("cod_proyecto"))
	        end if
	        clihojassep = limpiaCadena(request.form("clihojassep"))
	        if clihojassep="on" then clihojassep="1"
	        if request.form("opc_cod_proyecto")>"" then
		        opc_cod_proyecto="1"
	        end if
	        apaisado=iif(limpiaCadena(request.form("apaisado"))>"","SI","")
	        comercial=limpiaCadena(Request.QueryString("comercial"))
	        if comercial = "" then
		        comercial = limpiaCadena(Request.form("comercial"))
	        end if
	        tipo_cliente = limpiaCadena(Request.QueryString("tipo_cliente"))
	        if tipo_cliente = "" then
		        tipo_cliente = limpiaCadena(Request.form("tipo_cliente"))
	        end if
	        tipo_proveedor = limpiaCadena(Request.QueryString("tipo_proveedor"))
	        if tipo_proveedor = "" then
		        tipo_proveedor = limpiaCadena(Request.form("tipo_proveedor"))
	        end if
	        actividad_proveedor = limpiaCadena(Request.QueryString("actividad_proveedor"))
	        if actividad_proveedor = "" then
		        actividad_proveedor = limpiaCadena(Request.form("actividad_proveedor"))
	        end if

	        tipo_articulo =limpiaCadena(Request.QueryString("tipo_articulo"))
	        if tipo_articulo = "" then
		        tipo_articulo = limpiaCadena(Request.form("tipo_articulo"))
	        end if
	        mostrarfilas = limpiaCadena(Request.QueryString("mostrarfilas"))
	        if mostrarfilas = "" then
		        mostrarfilas = limpiaCadena(Request.form("mostrarfilas"))
	        end if
	        opc_cantidad = limpiaCadena(request.form("opc_cantidad"))
	        opc_ventasnetas = limpiaCadena(request.form("opc_ventasnetas"))
	        saveR=limpiaCadena(request.querystring("save"))

           'DBS 20130127  
            parametrosBD = obtener_param_obj("088",session("usuario"),session("ncliente"),"mode")
            lista_obt_obj = Split(parametrosBD, "&")
            series_Sal = ""
            if isArray(lista_obt_obj) then
                tamanyo = Ubound(lista_obt_obj)+1 'tamanyo del vector
                i=0
                while i < tamanyo   
                    if Mid(lista_obt_obj(i),1,6)&"" = "SERIES" then
                        lista_obt_obj_Aux=Split(lista_obt_obj(i),"=")
                        series_Sal=lista_obt_obj_Aux(1)
                    end if                        
                    i=i+1
                wend
            end if
             %><input type="hidden" name="series_Sal" id="series_Sal" value="<%=EncodeForHtml(series_Sal)%>" /><%
	        WaitBoxOculto LitEsperePorFavor

            set rstAux = Server.CreateObject("ADODB.Recordset")
            set rst = Server.CreateObject("ADODB.Recordset")
            set rst2 = Server.CreateObject("ADODB.Recordset")
            set rstSelect = Server.CreateObject("ADODB.Recordset")
            set rstTablas = Server.CreateObject("ADODB.Recordset")
            
            ' Fecha actual con formato 0X/0X/YYYY
            currentDayWithTrailingZero = iif(day(Date) < 10, "0" & day(Date), cstr(day(Date)))
            currentMonthWithTrailingZero = iif(month(Date) < 10, "0" & month(Date), cstr(month(Date)))
            currentDate = currentDayWithTrailingZero & "/" & currentMonthWithTrailingZero & "/" & cstr(year(Date))

	        if mode="browse" then%>
		        <table width='100%'>
		   	        <tr>
				        <td width="30%" align="left" >
				  	        <font class="CELDAC7">&nbsp;(<%=LitEmitido & "&nbsp;"%><%=currentDate%>)</font>
				        </td>
				        <td>
					        <font class='CABECERA'><b></b></font>
					        <font class='CELDA'><b></b></font>
				        </td>
				        <td></td>
			        </tr>
	            </table>
		        <hr/>
            <%end if
	        Alarma "resumen_ventas_cli.asp"

	        if (mode="select1") then%>
		        <hr />
                <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitParamFactura%></h6>
                <%DrawDiv "1","","Tdfdesde"
                DrawLabel "","",LitDesdeFecha%>
                    <input type="text" name="fdesde" value="<%=EncodeForHtml(iif(fdesde>"",fdesde,"01/01/" & year(date)))%>" size="0" />
                    <input type="text" name="hfrom" value="" size="5" maxlength="5" autocomplete="off" placeholder="HH:mm" />
                <%CloseDiv
                DrawCalendar "fdesde"
                DrawDiv "1","","Tdfhasta"
                DrawLabel "","",LitHastaFecha%>
                    <input type="text" name="fhasta" value="<%=EncodeForHtml(iif(fhasta>"",fhasta,currentDate))%>" size="0" />
                    <input type="text" name="hto" value="" size="5" maxlength="5" autocomplete="off" placeholder="HH:mm" />
                <%CloseDiv
                DrawCalendar "fhasta"%>
                <%rstSelect.cursorlocation=3
                if series_Sal&"">"" then
				    rstSelect.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(NOLOCK) where tipo_documento ='FACTURA A CLIENTE' and nserie like '" & session("ncliente") & "%' and nserie in "&Replace(Replace(Replace(series_Sal,",","','"),"(","('"),")","')")&" order by descripcion",Session("backendlistados")
                    DrawDiv "1","","Tdnserie"
                    DrawLabel "","",LitSerie%>
                    <select class="CELDA" style='width:200px' multiple="multiple" size="5" name="nserie" id="nserie">
                        <% valor=true    
                        while not rstSelect.eof                                    
                            if valor=true then
                                %><option selected value="<%=enc.EncodeForHtmlAttribute(null_s(rstSelect("nserie"))) %>"><%=enc.EncodeForHtmlAttribute(null_s(rstSelect("descripcion"))) %></option><%                                        
                                valor=false
                            else
                                %><option value="<%=enc.EncodeForHtmlAttribute(null_s(rstSelect("nserie"))) %>"><%=enc.EncodeForHtmlAttribute(null_s(rstSelect("descripcion"))) %></option><% 
                            end if                                       
                            rstSelect.movenext                                
                        wend       
                        %>
                    </select>
                    <%CloseDiv
                else
                    rstSelect.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(NOLOCK) where tipo_documento ='FACTURA A CLIENTE' and nserie like '" & session("ncliente") & "%' order by descripcion",Session("backendlistados")
                    DrawSelectMultipleCelda "","200","",0,LitSerie,"nserie' style='width:200px",rstSelect,enc.EncodeForHtmlAttribute(nserie),"nserie","descripcion","size","8"
                end if
                
				rstSelect.close
                Dim Literal2
				if si_tiene_modulo_comercial<>0 then
					Literal2 = LitComercialModCom
				else
					Literal2 = LitComercial
				end if
                rstSelect.cursorlocation=3
				rstSelect.open "SELECT p.dni, p.nombre FROM PERSONAL AS p with(NOLOCK), comerciales AS c with(NOLOCK) WHERE c.fbaja is null and p.dni = c.comercial and p.dni like '" & session("ncliente") & "%' order by p.nombre", Session("backendlistados")
				DrawSelectMultipleCelda "","200","",0,Literal2,"comercial' style='width:200px",rstSelect,enc.EncodeForHtmlAttribute(comercial),"dni","nombre","size","8"
				rstSelect.close
                EligeCelda "check","add","left","","",0,LitFechaAlbaran,"fechaalbaran",0,""%>
            <hr />
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitParamClientes%></h6>
			<%if ncliente >"" then ncliente=Completar(ncliente,5,"0")
                DrawDiv "1","",""
                DrawLabel "","",LitDesdeCliente
				ncliente_aux=ncliente
				if ncliente & "">"" then
					nombre=d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
				else
					nombre=""
					ncliente=""
				end if
				if nombre & ""="" and ncliente_aux & "">"" then
					%><script language="javascript" type="text/javascript">
						alert("<%=LitMsgClienteNoExiste%>");
					</script><%
				end if
				if acliente >"" then acliente=Completar(acliente,5,"0")
				ncliente_aux=""
				%><input class='width15' type="text" name="ncliente" value="<%=enc.EncodeForHtmlAttribute(ncliente)%>" size='10' onchange="TraerCliente('<%=enc.EncodeForJavascript(null_s(mode))%>','1');"/><a class='CELDAREFB' href="javascript:AbrirVentana('../clientes_buscar.asp?ndoc=resumen_ventas_cli&titulo=<%=LitSelCliente%>&mode=search&viene=resumen_ventas_cli','P','<%=AltoVentana%>','<%=AnchoVentana%>')" onmouseover="self.status='<%=LitVerCliente%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscarDinamic%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' disabled type="text" name="nombre" value="<%=enc.EncodeForHtmlAttribute(nombre)%>" size="20" /><%CloseDiv
                DrawDiv "1","",""
                DrawLabel "","",LitACliente
				ncliente_aux=acliente
				if acliente & "">"" then
					anombre=d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & acliente & "'",Session("backendlistados"))
				else
					anombre=""
					acliente=""
				end if
				if anombre & ""="" and ncliente_aux & "">"" then
					%><script language="javascript" type="text/javascript">
						alert("<%=LitMsgClienteNoExiste%>");
					</script><%
				end if
				%><input class='width15' type="text" name="acliente" value="<%=enc.EncodeForHtmlAttribute(acliente)%>" size='10' onchange="TraerCliente('<%=enc.EncodeForJavascript(null_s(mode))%>','2');"/><a class='CELDAREFB' href="javascript:AbrirVentana('../clientes_buscar.asp?ndoc=resumen_ventas_cli2&titulo=<%=LitSelCliente%>&mode=search&viene=resumen_ventas_cli2','P','<%=AltoVentana%>','<%=AnchoVentana%>')" onmouseover="self.status='<%=LitVerCliente%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscarDinamic%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' disabled type="text" name="anombre" value="<%=enc.EncodeForHtmlAttribute(anombre)%>" size="20" /><%
			    CloseDiv
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo, case when datalength(right(codigo,len(codigo)-5)+' '+descripcion)<=21 then right(codigo,len(codigo)-5)+'-'+descripcion else left(right(codigo,len(codigo)-5)+'-'+descripcion,20)+'...' end as descripcion from tipo_actividad with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",Session("backendlistados")
                DrawSelectCelda "width60","","",0,LitActividad,"tactividad",rstSelect,enc.EncodeForHtmlAttribute(tactividad),"codigo","descripcion","",""
				rstSelect.close
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo,descripcion from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo ='" & LitClienteMin & "' order by descripcion",Session("backendlistados")
				DrawSelectCelda "width60","","",0,LitTipoCliente,"tipo_cliente",rstSelect,enc.EncodeForHtmlAttribute(tipo_cliente),"codigo","descripcion","",""
				rstSelect.close
                DrawDiv "1","",""
                DrawLabel "","",LitClienteBaja%><input type="checkbox" name="opcclientebaja" <%=iif(opcclientebaja="on" or opcclientebaja="true" or opcclientebaja="1","checked","")%>/><%CloseDiv
			    if si_tiene_modulo_proyectos<>0 then
                    %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><input class="CELDA" type="hidden" name="cod_proyecto" value="<%=enc.EncodeForHtmlAttribute(cod_proyecto)%>" /><label><%=LitProyecto%></label><%
                    %><iframe class="width60 iframe-menu" id='frProyecto' src='../../mantenimiento/docproyectos_responsive.asp?viene=resumen_ventas_cli&mode=<%=EncodeForHtml(mode)%>&cod_proyecto=<%=EncodeForHtml(cod_proyecto)%>' frameborder="no" scrolling="no" noresize="noresize"></iframe></div><%
			    end if
                EligeCelda "input","add","left","","",0,LITPROVINCIA,"provincia",0,""%>
		        <hr />
                <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitParamArticulos%></h6><%
                EligeCelda "input","add","left","","",0,LitConref2,"referencia",0,referencia
                EligeCelda "input","add","left","","",0,LitConNombre2,"nombreart",0,nombreart

                set conn = Server.CreateObject("ADODB.Connection")
                set command =  Server.CreateObject("ADODB.Command")
                conn.open session("backendListados")
                command.ActiveConnection =conn
                command.CommandTimeout = 0
                command.CommandText="getAllEntityTypeByType"
                command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                command.Parameters.Append command.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))
                command.Parameters.Append command.CreateParameter("@type", adVarChar, adParamInput, 20, "ARTICULO")

                set rstArtType = command.execute

                DrawSelectMultipleCelda "CELDA",200,"",0,LitTipoArt,"tipo_articulo' style='width:200px",rstArtType,enc.EncodeForHtmlAttribute(tipo_articulo),"codigo","descripcion","size","8"

                rstArtType.close
                conn.close
                set rstArtType = nothing
                set command = nothing
                set conn = nothing

				dim ConfigDespleg (3,13)

				i=0
				ConfigDespleg(i,0)="categoria"
				ConfigDespleg(i,1)="200"
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="CELDA"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitCategoria
				ConfigDespleg(i,10)=categoria
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				i=1
				ConfigDespleg(i,0)="familia_padre"
				ConfigDespleg(i,1)="200"
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="CELDA"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitFamilia
				ConfigDespleg(i,10)=familia_padre
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				i=2
				ConfigDespleg(i,0)="familia"
				ConfigDespleg(i,1)="200"
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre, categoria, padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="CELDA"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitSubFamilia
				ConfigDespleg(i,10)=familia
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				DibujaDesplegables ConfigDespleg,Session("backendlistados")%>
            <hr />
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitVentasArtProveedor%></h6><%

				if nproveedor>"" then nproveedor=Completar(nproveedor,5,"0")
                DrawDiv "1","",""
                DrawLabel "","",LitProveedor 
				nproveedor_aux=nproveedor
				if nproveedor & "">"" then
					razon_social=d_lookup("razon_social","proveedores","nproveedor='" & session("ncliente") & nproveedor & "'",Session("backendlistados"))
				else
					razon_social=""
					nproveedor=""
				end if
				if razon_social & ""="" and nproveedor_aux & "">"" then
					%><script language="javascript" type="text/javascript">
						alert("<%=LitMsgProveedorNoExiste%>");
					</script><%
				end if
				nproveedor_aux=""
				%><input class='width15' type="text" name="nproveedor" value="<%=enc.EncodeForHtmlAttribute(nproveedor)%>" size='10' onchange="TraerProveedor('<%=enc.EncodeForJavascript(null_s(mode))%>','1');" /><a class='CELDAREFB' id="verc" href="javascript:AbrirVentana('../../compras/proveedores_busqueda.asp?ndoc=resumen_ventas_cli&titulo=<%=LitSelProveedor%>&mode=search&viene=resumen_ventas_cli','P','<%=AltoVentana%>','<%=AnchoVentana%>')" onmouseover="self.status='<%=LitVerProveedor%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscarDinamic%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' disabled type="text" name="razon_social" value="<%=enc.EncodeForHtmlAttribute(razon_social)%>" size="20" /><%
                CloseDiv
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo,descripcion from tipo_actividad with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",Session("backendlistados")
				DrawSelectCelda "CELDA",175,"",0,LitActividad,"actividad_proveedor",rstSelect,enc.EncodeForHtmlAttribute(actividad_proveedor),"codigo","descripcion","",""
				rstSelect.close
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo,descripcion from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo ='" & LitProveedor & "' order by descripcion",Session("backendlistados")
				DrawSelectCelda "CELDA",175,"",0,LitTipoProveedor,"tipo_proveedor",rstSelect,enc.EncodeForHtmlAttribute(tipo_proveedor),"codigo","descripcion","",""
				rstSelect.close%>
            <hr />
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitParamGenerales%></h6><%
                DrawDiv "1","",""
                DrawLabel "","",LitAgrTpCliente%><select class="width60" name="agrupar" onchange="javascript:optionalFields();">
						<option <%=iif(agrupar="ARTICULO","selected","")%> value="ARTICULO"><%=LitArticulosMay%></option>
						<option <%=iif(agrupar="CLIENTE" or agrupar="","selected","")%> value="CLIENTE"><%=LitCliente%></option>
						<option <%=iif(agrupar="MESES","selected","")%> value="MESES"><%=LitMeses%></option>
						<%if si_tiene_modulo_proyectos<>0 then%>
							<option <%=iif(agrupar="PROYECTO","selected","")%> value="PROYECTO"><%=ucase(LitProyecto)%></option>
						<%end if%>
					</select><%CloseDiv
				DrawDiv "1","display: none","agrMeses2"
                DrawLabel "","",LitMostrarFilas%><select id="agrMeses3" class="width60" style="display: none" name="mostrarfilas" onchange="javascript:optionalFields();">
						<option <%=iif(mostrarfilas="ARTICULO","selected","")%> value="ARTICULOS"><%=LitArticulosMay%></option>
						<option <%=iif(mostrarfilas="CLIENTES" or mostrarfilas="","selected","")%> value="CLIENTES"><%=LitCliente%></option>
					</select><%CloseDiv
				rstSelect.CursorLocation=3
				
                if series_Sal&"">"" then    
                    rstSelect.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='ALBARAN DE SALIDA' and nserie like '" & session("ncliente") & "%' and nserie in "&Replace(Replace(Replace(series_Sal,",","','"),"(","('"),")","')")&" order by descripcion", Session("backendlistados")
                    DrawDiv "1","","TDseriesAPF"
                    DrawLabel "","",LitAlbPendFac%><select class="width60" style="width:200px;" multiple="multiple" size="5" name="seriesAPF" id="seriesAPF"><%                  
                    valor=true
                    while not rstSelect.eof   
                       if valor=true then
                                %><option selected value="<%=enc.EncodeForHtmlAttribute(null_s(rstSelect("nserie"))) %>"><%=enc.EncodeForHtmlAttribute(null_s(rstSelect("descripcion"))) %></option><%                                        
                                valor=false
                            else
                                %><option value="<%=enc.EncodeForHtmlAttribute(null_s(rstSelect("nserie"))) %>"><%=enc.EncodeForHtmlAttribute(null_s(rstSelect("descripcion"))) %></option><% 
                            end if                                                                                          
                            rstSelect.movenext                                
                        wend                                        
                    %></select><%CloseDiv 
                else
                    rstSelect.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='ALBARAN DE SALIDA' and nserie like '" & session("ncliente") & "%' order by descripcion", Session("backendlistados")
                    DrawSelectMultipleCelda "width60","200","",0,LitAlbPendFac,"seriesAPF' style='width: 200px",rstSelect,"","nserie","descripcion","size","5"
                end if
                
				rstSelect.close
				if si_tiene_modulo_tiendas<>0 then
					rstSelect.CursorLocation=3
					rstSelect.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='TICKET' and nserie like '" & session("ncliente") & "%' order by descripcion", Session("backendlistados")
					DrawSelectMultipleCelda "CELDA",200,"",0,LitTicketsPF,"seriesTPF' style='width: 200px",rstSelect,"","nserie","descripcion","size","5"
					rstSelect.close
				else
                end if
                if session("version")&"" <> "5" then
                    DrawDiv "","","" 
                    CloseDiv
                end if 
                DrawDiv "1","",""
                DrawLabel "","",LitApaisado%><input type="checkbox" name="apaisado" <%=iif(apaisado="SI" or apaisado="on" or apaisado="true" or apaisado="1","checked","")%>/><%CloseDiv
				EligeCelda "check","add","left","","",0,LitOrdenarVentas,"ordenar_ventas",0,iif(set_orden="on" or set_orden="true" or set_orden="1","True","")
            %><span  id="agrCliHojSep" style="display: "><%
                DrawDiv "1","",""
                DrawLabel "","", LitClientePorPagina%><input type="checkbox" name="clihojassep" <%=iif(clihojassep="on" or clihojassep="true" or clihojassep="1","checked","")%>/><%CloseDiv
				DrawDiv "1","",""
                DrawLabel "","",LitMostrarConceptos%><input type="checkbox" name="ver_conceptos" <%=iif(ver_conceptos="on" or ver_conceptos="true" or ver_conceptos="1","checked","")%> onclick="javascript:Ver_Conceptos();"/><%CloseDiv
			    %></span><hr />
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitCamposOpcionales%></h6><%

			if si_tiene_modulo_proyectos<>0 then%>
				<span id="agrOtros" style="display: "><%
                EligeCelda "check","add","left","","",0,LitProyecto,"opc_cod_proyecto",0,iif(opc_cod_proyecto="on" or opc_cod_proyecto="true" or opc_cod_proyecto="1","True","")
                EligeCelda "check","add","left","","",0,LITDESGLOSARCPTOS,"conceptos",0,iif(conceptos="on" or conceptos="true" or conceptos="1","True","")%>
                </span><%
			else
				EligeCelda "check","add","left","","",0,LITDESGLOSARCPTOS,"conceptos",0,iif(conceptos="on" or conceptos="true" or conceptos="1","True","")
			end if
                'INICIO AÑADIR IMPORTE MEDIO IVA
                %><span id="importIva" style="display:"><%
                    EligeCelda "check","add","left","","",0,LITIMPORTEMEDIOIVA,"importeMedioIva",0,iif(importeMedioIva="on" or importeMedioIva="true" or importeMedioIva="1","True","")
                %></span><%
                'FIN AÑADIR IMPORTE MEDIO IVA
            %>
                <span id="agrMeses" style="display: none"><%
                DrawDiv "1","",""
                DrawLabel "", "", LitCantidad%><input type="checkbox" name="opc_cantidad" <%=iif(opc_cantidad="on" or opc_cantidad="true" or opc_cantidad="1" or opc_ventasnetas="","checked","")%> onclick="javascript:if(document.resumen_ventas_cli.opc_cantidad.checked) document.resumen_ventas_cli.opc_ventasnetas.checked=false;else document.resumen_ventas_cli.opc_ventasnetas.checked=true;"/><%CloseDiv
                DrawDiv "1","",""
                DrawLabel "","",LitVentasNetas%><input type="checkbox" name="opc_ventasnetas" <%=iif(opc_ventasnetas="on" or opc_ventasnetas="true" or opc_ventasnetas="1","checked","")%> onclick="javascript:BloquearVentasNetas()"/><%CloseDiv%>
                </span><%
          if agrupar & "">"" then%>
			<script language="javascript" type="text/javascript">
				optionalFields();
			</script>
        <%end if
'****************************************************************************************************************

		'Mostrar el listado.
	elseif mode="browse" then
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"inicio_resumen_ventas"%>
		    <input type="hidden" name="fdesde" value="<%=EncodeForHtml(fdesde)%>" />
		    <input type="hidden" name="fhasta" value="<%=EncodeForHtml(fhasta)%>" />
		    <input type="hidden" name="nserie" value="<%=EncodeForHtml(nserie)%>" />
		    <input type="hidden" name="ncliente" value="<%=EncodeForHtml(ncliente)%>" />
		    <input type="hidden" name="acliente" value="<%=EncodeForHtml(acliente)%>" />
		    <input type="hidden" name="actividad" value="<%=EncodeForHtml(actividad)%>" />
	  	    <input type="hidden" name="tactividad" value="<%=EncodeForHtml(tactividad)%>" />
		    <input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>" />
			<input type="hidden" name="nombreart" value="<%=EncodeForHtml(nombreart)%>" />
			<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>" />
			<input type="hidden" name="familia_padre" value="<%=EncodeForHtml(familia_padre)%>" />
			<input type="hidden" name="categoria" value="<%=EncodeForHtml(categoria)%>" />
			<input type="hidden" name="agrupar" value="<%=EncodeForHtml(agrupar)%>" />
			<input type="hidden" name="conceptos" value="<%=EncodeForHtml(conceptos)%>" />
            <!--INICIO AÑADIR IMPORTE MEDIO IVA-->
            <input type="hidden" name="importeMedioIva" value="<%=EncodeForHtml(importeMedioIva)%>" />
            <!--FIN AÑADIR IMPORTE MEDIO IVA-->
			<input type="hidden" name="ver_conceptos" value="<%=EncodeForHtml(ver_conceptos)%>" />
			<input type="hidden" name="ordenar_ventas" value="<%=EncodeForHtml(set_orden)%>" />
			<input type="hidden" name="opcclientebaja" value="<%=EncodeForHtml(opcclientebaja)%>" />
			<input type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(cod_proyecto)%>" />
			<input type="hidden" name="opc_cod_proyecto" value="<%=EncodeForHtml(opc_cod_proyecto)%>" />
			<input type="hidden" name="clihojassep" value="<%=EncodeForHtml(clihojassep)%>" />
			<input type="hidden" name="opc_cantidad" value="<%=EncodeForHtml(opc_cantidad)%>" />
			<input type="hidden" name="opc_ventasnetas" value="<%=EncodeForHtml(opc_ventasnetas)%>" />
			<input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>" />
			<input type="hidden" name="comercial" value="<%=EncodeForHtml(comercial)%>" />
			<input type="hidden" name="tipo_cliente" value="<%=EncodeForHtml(tipo_cliente)%>" />
			<input type="hidden" name="tipo_proveedor" value="<%=EncodeForHtml(tipo_proveedor)%>" />
			<input type="hidden" name="actividad_proveedor" value="<%=EncodeForHtml(actividad_proveedor)%>" />
			<input type="hidden" name="tipo_articulo" value="<%=EncodeForHtml(tipo_articulo)%>" />
			<input type="hidden" name="nproveedor" value="<%=EncodeForHtml(nproveedor)%>" />
			<input type="hidden" name="mostrarfilas" value="<%=EncodeForHtml(mostrarfilas)%>" />
			<input type="hidden" name="seriesTPF" value="<%=EncodeForHtml(seriesTPF)%>" />
			<input type="hidden" name="seriesAPF" value="<%=EncodeForHtml(seriesAPF)%>" />
            <input type="hidden" name="provincia" value="<%=EncodeForHtml(provincia)%>" />
<%
			MB=d_lookup("codigo", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", Session("backendlistados"))
			n_decimales = null_z(d_lookup("ndecimales", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", Session("backendlistados")))
			n_decimalesMB = n_decimales
			MB_abrev = d_lookup("abreviatura", "divisas", "codigo='" & MB & "' and codigo like '" & session("ncliente") & "%'", Session("backendlistados"))
			MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='088'", DSNIlion)
			MAXPDF=d_lookup("maxpdf", "limites_listados", "item='088'", DSNIlion)
			%><input type='hidden' name='maxpdf' value='<%=enc.EncodeForHtmlAttribute(MAXPDF)%>'/>
<%

			VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarArticulos)=1
			CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

			PorCliente="NO"
			PorGasto="NO"

			tsel1 = true
			tsel2 = true

			if conceptos = "on" then
			   desglose = true
			else
			   desglose = false
			end if
            
            'INICIO AÑADIR IMPORTE MEDIO IVA
            if importeMedioIva = "on" then
			   importeIva = true
			else
			   importeIva = false
			end if   
            'FIN AÑADIR IMPORTE MEDIO IVA

			if ver_conceptos = "on" then tsel1=false

			if set_orden="on" then
			   ordenar=true
			else
			   ordenar=false
			end if

			if fdesde>"" then
				%><font class="ENCABEZADO"><b><%=LitDesdeFecha%> : </b></font><font class='CELDA'><%=EncodeForHtml(fdesde)%></b></font><br/><%
			end if
			if fhasta>"" then
				%><font class="ENCABEZADO"><b><%=LitHastaFecha%> : </b></font><font class='CELDA'><%=EncodeForHtml(fhasta)%></b></font><br/><%
			end if

			if comercial& "">"" then
				desc_comercial= d_lookup("nombre", "PERSONAL", "dni='" & comercial & "'", Session("backendlistados"))
				desc_comercial=NombresEntidades(comercial,"personal","dni","nombre",Session("backendlistados"))
				%><font class="ENCABEZADO"><b><%
					if si_tiene_modulo_comercial<>0 then
						response.write(LitComercialModCom)
					else
						response.write(LitComercial)
					end if
				%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_comercial)%></b></font><br/><%
			end if
			if actividad_proveedor& "">"" then
				desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & actividad_proveedor & "'", Session("backendlistados"))
				%><font class="ENCABEZADO"><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(actividad_proveedor))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
			end if
			if tipo_cliente& "">"" then
				desc_tipo_cliente= d_lookup("descripcion", "tipos_entidades", "codigo='" & tipo_cliente & "'", Session("backendlistados"))
				%><font class="ENCABEZADO"><b><%=LitTipoCliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tipo_cliente))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tipo_cliente)%></b></font><br/><%
			end if
			if tipo_proveedor& "">"" then
				desc_tipo_proveedor= d_lookup("descripcion", "tipos_entidades", "codigo='" & tipo_proveedor & "'", Session("backendlistados"))
				%><font class="ENCABEZADO"><b><%=LitTipoProveedor%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tipo_proveedor))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tipo_proveedor)%></b></font><br/><%
			end if
			if tipo_articulo& "">"" then
                desc_tipo_articulo=NombresEntidades(tipo_articulo,"tipos_entidades","codigo","descripcion",Session("backendlistados"))
				%><font class="ENCABEZADO"><b><%=LitTipoArt%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_tipo_articulo)%></b></font><br/><%
			end if
			if nproveedor & "">"" then
				razon_social = d_lookup("razon_social", "proveedores", "nproveedor='" & session("ncliente") & nproveedor & "'", Session("backendlistados"))
				%><font class="ENCABEZADO"><b><%=LitProveedor%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(nproveedor))%>&nbsp;&nbsp;<%=EncodeForHtml(razon_social)%></b></font><br/><%
			end if
			if mostrarfilas & "">"" and agrupar="MESES" then%>
				<font class="ENCABEZADO"><b><%=LitMostrarFilas%> : </b></font><font class='CELDA'><%=EncodeForHtml(mostrarfilas)%></b></font><br/>
			<%end if
			if agrupar="CLIENTE" then
				PorCliente = "SI"
				if tactividad>"" then 'Se selecciono tipo de actividad
					desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", Session("backendlistados"))
					%><font class="ENCABEZADO"><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
				end if
				if ncliente>"" and acliente>"" and ncliente=acliente then 'Se selecciono cliente
					PorCliente="NO"
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					%><font class="ENCABEZADO"><b><%=LitClienteMin%> : </b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><br/><%
					t_opcclientebaja=1
				else
					if opcclientebaja="" then
						t_opcclientebaja=1
					else
						t_opcclientebaja=0%>
						<font class="ENCABEZADO"><b><%=LitClienteBaja%></b></font><br/>
                    <%end if
				end if
				if ncliente>"" and acliente>"" and ncliente<>acliente then
					t_opcclientebaja=1
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					acliente=Completar(acliente,5,"0")
					nomclihasta=d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & acliente & "'",Session("backendlistados"))
					%><font class="ENCABEZADO"><b><%=LitDesdeCliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><%
					%> - <%
					%><font class="ENCABEZADO"><b><%=LitACliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(acliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomclihasta)%></b></font><br/>
                <%end if
				if serie<>"" then 'Se selecciono serie
					%><font class="ENCABEZADO"><b><%=LitSerie%> : </b></font><font class='CELDA'><%=NombresEntidades(serie,"series","nserie","nombre",Session("backendlistados"))%></b></font><br/><%
				end if
				if familia<>"" then
					tsel2 = false
					desc_familia=NombresEntidades(familia,"familias","codigo","nombre",Session("backendlistados"))
					%><font class="ENCABEZADO"><b><%=LITSUBFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia)%></b></font><br/><%
				elseif familia_padre<>"" then
					tsel2 = false
					desc_familia_padre=NombresEntidades(familia_padre,"familias_padre","codigo","nombre",Session("backendlistados"))
					%><font class="ENCABEZADO"><b><%=LITFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia_padre)%></b></font><br/><%
				elseif categoria<>"" then
					tsel2 = false
					desc_categoria=NombresEntidades(categoria,"categorias","codigo","nombre",Session("backendlistados"))
					%><font class="ENCABEZADO"><b><%=LITCATEGORIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_categoria)%></b></font><br/><%
				end if

				if cod_proyecto>"" then
					%><font class="ENCABEZADO"><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",Session("backendlistados")))%></font><br/><%
				end if
				if provincia>"" then
					%><font class="ENCABEZADO"><b><%=LITPROVINCIA%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(provincia)%></font><br/><%
				end if
				if saveR="true" then
					crearCliente tsel1, tsel2,seriesAPF, fdesde, fhasta, tactividad, iif(ncliente>"",session("ncliente") & ncliente,ncliente),iif(acliente>"",session("ncliente") & acliente,acliente), serie, familia, familia_padre, categoria, referencia, nombreart, ordenar, MB,t_opcclientebaja,cod_proyecto,seriesTPF,opc_cod_proyecto,iif(nproveedor>"",session("ncliente") & nproveedor,nproveedor),actividad_proveedor,tipo_cliente,tipo_proveedor,comercial,tipo_articulo,mostrarfilas,provincia
				else%>
					<input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' />
                <%end if%>
				<hr/>
<%
				if seriesAPF > "" or seriesTPF > "" then
					seleccion="SELECT NCliente, Nombre, Referencia,Descripcion,"
					if opc_cod_proyecto="1" then
						seleccion=seleccion & "cod_proyecto,"
					end if
					seleccion=seleccion & "SUM(cantidad) AS Cantidad"
					seleccion=seleccion & ",SUM(cantidad2) AS Cantidad2"
					seleccion=seleccion & ",case when convert(nvarchar,calculoimporte)=1 then 'VERDADERO' else 'FALSO' end as calculoimporte"
					seleccion=seleccion & ",tipo_medida"
					seleccion=seleccion & ", SUM([Ventas Netas]) AS [Ventas Netas],"
					seleccion=seleccion & "SUM([Precio Medio] * cantidad)/ CASE WHEN SUM(cantidad)= 0 THEN 1 ELSE SUM(cantidad) END AS [Precio Medio],"
					seleccion=seleccion & "SUM([Precio Medio2] * cantidad2)/ CASE WHEN SUM(cantidad2)= 0 THEN 1 ELSE SUM(cantidad2) END AS [Precio Medio2],"
					seleccion=seleccion & "Divisa,(select sum([Ventas Netas]) from [" & session("usuario") & "] WHERE ncliente = f.NCliente) as Acumulador,Orden"
					seleccion=seleccion & ",tiene_escv "
					seleccion=seleccion & " FROM [" & session("usuario") & "]  AS f " & strbaja
					seleccion=seleccion & " GROUP BY ncliente,nombre,referencia,descripcion,divisa,orden,tiene_escv,convert(nvarchar,calculoimporte),tipo_medida "
					if opc_cod_proyecto>"" then
						seleccion=seleccion & ",cod_proyecto "
					end if
					if ordenar = true then
						seleccion=seleccion & " order by orden desc"
					end if
				else
					if ordenar = true then
						seleccion = "select * from [" & session("usuario") & "]" & strbaja & " order by orden desc"
					else
						seleccion = "select * from [" & session("usuario") & "]" & strbaja
					end if
				end if
				'comprobamos si existen varias divisas
                rst.cursorlocation=3
				rst.open seleccion,Session("backendlistados")
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
''ricardo 23/1/2003
''si no se hacia esto, en la primera pagina, no pone el cliente que es.
cliente=""

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
					<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>' />
                    <%NavPaginas lote,lotes,campo,criterio,texto,1%>
					<br/>
					<table width="100%" style="border-collapse: collapse;">
						    <%if cliente="" then
								DrawFila color_fondo
								%><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-3-ocultar-cint(iif(opc_cod_proyecto="1","0","1"))-1%>"><%=LitClienteMin%>:
									<%if left(rst("Ncliente"),5)="-----" then%>
										<%=enc.EncodeForHtmlAttribute(null_s(rst("Nombre")))%>
									<%else%>
										<%=Hiperv(OBJClientes,rst("Ncliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("Ncliente")) & " " & rst("Nombre")),LitVerCliente)%>
									<%end if%>
								</td>
								<%CloseFila
							end if
							DrawFila color_fondo%>
							    <td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
								<%if opc_cod_proyecto="1" then%>
									<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
								<%end if%>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitVentasNetas%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td><%
                                'INICIO AÑADIR IMPORTE MEDIO IVA
                                if importeIva=true then
									%><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LITIMPORTEMEDIOIVA%></td><%
								end if
                                'FIN AÑADIR IMPORTE MEDIO IVA
								if MostrarDivisa=true then
									%><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
								end if                                
							CloseFila

							ValorCampoAnt = ""
							AcumuladoAnt  = ""
							ImprimeCabecera = false
							fila = 1
							while not rst.eof and fila<=MAXPAGINA
								if left(rst("Ncliente"),5)="-----" then
								else
									CheckCadena rst("NCliente")
								end if

								DrawFila ""
									for each campo in rst.fields
										if campo.name="tiene_escv" then
										else
											if (PorCliente="SI" and campo.name="NCliente") then
												if rst(campo.name)<>ValorCampoAnt then
													if ValorCAmpoAnt<>"" then
														'antes de imprimir subtotales imprimimos conceptos (si existen)
														'Fila de Subtotal%>
														<td></td>
														<td></td>
														<%if opc_cod_proyecto="1" then%>
															<td></td>
														<%end if%>
														<td class="tdbordecelda7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
				                                        				<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
														<%CloseFila

														''ricardo 22/1/2003 para que se veo o no lo de clientes en hojas separadas, ya que solo sirve para la agrupacion por cliente
														if clihojassep="on" or clihojassep="1" then
															%></table>
															<h6 class="SALTO">&nbsp;</h6>
															<table width="100%" style="border-collapse: collapse;"><%
														else
															DrawFila "" 'Fila de separacion
																%><td colspan="<%=rst.fields.count-6-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
															CloseFila
														end if

														DrawFila color_fondo
															%><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-4-ocultar-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=LitClienteMin%>:
																<%if left(rst("Ncliente"),5)="-----" then%>
																	<%=enc.EncodeForHtmlAttribute(null_s(rst("Nombre")))%>
																<%else%>
																	<%=Hiperv(OBJClientes,rst("Ncliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("Ncliente")) & " " & rst("Nombre")),LitVerCliente)%>
																<%end if%>
															</td>
                                                        <%CloseFila
														DrawFila color_fondo%>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
															<%if opc_cod_proyecto="1" then%>
																<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
															<%end if%>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitVentasNetas%></td>
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
											elseif campo.name="Cantidad" or campo.name="Cantidad2" or campo.name="calculoimporte" or campo.name="tipo_medida" or campo.name="Precio Medio2" or campo.name="Ventas Netas" or campo.name="Precio Medio" then 'Formateo del campo con importe
												'ajustamos divisas si es necesario
												if MostrarDivisa=true then
													n_decimales = null_z(d_lookup("ndecimales", "divisas", "codigo='" & rst("Divisa") & "'", Session("backendlistados")))
												end if
												if campo.name="Ventas Netas" then
	 												AcumuladoAnt = rst("Acumulador")
												end if
												''ricardo 22-3-2004
												tipo_medida=rst("tipo_medida")
												valor_cantidad2=null_z(rst("Cantidad2"))
												calculoimporte=nz_b(rst("calculoimporte"))
												''ricardo 22-3-2004
												if (campo.name<>"cod_proyecto" or opc_cod_proyecto="1") and campo.name<>"tipo_medida" then
													%>
														<%if rst("tiene_escv")<>1 or campo.name="Cantidad" or campo.name="Cantidad2" then
															''ricardo 22-3-2004
															if campo.name="Cantidad" or campo.name="Cantidad2" then
																if campo.name="Cantidad" then
																	%><td class="tdbordecelda7" align="right"><%
																		if valor_cantidad2<>0 then%>
																			<%=EncodeForHtml(formatnumber(rst(campo.name),DEC_CANT,-1,0,-1))%>
																			<br/>
																			<%=EncodeForHtml(("<b>" & iif(tipo_medida>"",tipo_medida,"") & " : </b>" & formatnumber(valor_cantidad2,DEC_CANT,-1,0,-1)))%>
																		<%else%>
																			<%=EncodeForHtml(formatnumber(rst(campo.name),DEC_CANT,-1,0,-1))%>
																		<%end if
																	%></td><%
																end if
															else
																if campo.name="Precio Medio" or campo.name="Precio Medio2" or campo.name="calculoimporte" then
																	if campo.name="Precio Medio" then%>
																		<td class="tdbordecelda7" align="right">
																			<%=EncodeForHtml(formatnumber(rst(campo.name),n_decimales,-1,0,-1))%>
																			<%if calculoimporte=-1 and rst("Precio Medio2")<>0 then%>
																				<br/>
																				<%=EncodeForHtml(("<b>" & iif(tipo_medida>"",tipo_medida,"") & " : </b>" & formatnumber(rst("Precio Medio2"),n_decimales,-1,0,-1)))%>
																			<%end if%>
																		</td>
																	<%end if
																else%>
																	<td class="tdbordecelda7" align="right">
																		<%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",formatnumber(rst(campo.name),iif(campo.name<>"Cantidad",n_decimales,DEC_CANT),-1,0,-1)))%>
																	</td>
																<%end if
															end if
														else
															if campo.name<>"Cantidad" and campo.name<>"Cantidad2" and campo.name<>"calculoimporte" and campo.name<>"Precio Medio2" then%>
																<td class="tdbordecelda7" align="right">&nbsp;</td>
															<%end if
														end if
												end if
											else
												if campo.name<>"Nombre" and campo.name<>"NCliente" and campo.name<>"Divisa" then
													if campo.name="Referencia" then
												      	if rst(campo.name)="zzzzzzzzzzzzzzzzzzzz" then
															%><td class="tdbordecelda7">Concepto</td><%
														else
															%><td class="tdbordecelda7">
																<%=Hiperv(OBJArticulos,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerArticulo)%>
															</td><%
														end if
													else
				            	                  	                	if campo.name<>"Orden" and campo.name<>"Acumulador" then
													  		if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
													         		%><td class="tdbordecelda7"><%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",rst(campo.name)))%></td><%
															end if
													  	else
													    	%><td></td><%
													  	end if
												   	end if
												   'end aki
												else
													if campo.name="Divisa" and MostrarDivisa=true then
												      	%><td class="tdbordecelda7"><%=EncodeForHtml(d_lookup("abreviatura", "divisas", "codigo='" & rst(campo.name) & "' and codigo like '" & session("ncliente") & "%'", Session("backendlistados")))%></td><%
													else
								                  	        	if campo.name = "Divisa" and MostrarDivisa = false then '
									                  	      	%><td></td><%
 									                      		end if
													end if
												end if
											end if
										end if
									next
								CloseFila
								fila = fila +1
								rst.movenext
								''ricardo 22-3-2004
								valor_cantidad2=0
								tipo_medida=""
							wend

							'solo mostramos subtotales si se alcanzó el final del rst
							if rst.eof then
								if (PorCliente="SI") then
									DrawFila "" 'Fila de Subtotal
										%><td></td>
										<td></td>
										<%if opc_cod_proyecto="1" then%>
											<td></td>
										<%end if%>
										<td class="tdbordecelda7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
										<td class="tdbordecelda7" align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
								   	<%CloseFila
								   	DrawFila "" 'Fila de separacion
										%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
     									CloseFila
								end if
								DrawFila "" 'Fila para el total
									%><td></td>
									<td></td>
									<%if opc_cod_proyecto="1" then%>
										<td></td>
									<%end if
									Suma = elTotal%>
									<td class="tdbordecelda7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
									<td class="tdbordecelda7" align="right"><big><b><%=EncodeForHtml(formatnumber(Suma,n_decimalesMB,-1,0,-1))%></b></big></td>
								<%CloseFila%>
								<%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",Session("backendlistados")) then
									DrawFila "" 'Fila para el total equivalencia en PTS
					      				%><td></td>
							   			<td></td>
										<%if opc_cod_proyecto="1" then%>
											<td></td>
										<%end if%>
					      				<td class="tdbordecelda7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",Session("backendlistados")))%>:</b></td>
				      					<td class="tdbordecelda7" align="right"><big><b><%=EncodeForHtml(formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",Session("backendlistados")),-1,0,-1))%></b></big></td>
									<%CloseFila
								end if
							end if%>

					</table><br/><%
					NavPaginas lote,lotes,campo,criterio,texto,2
				else
					%><script language="javascript" type="text/javascript">
					      alert("<%=LitMsgDatosNoExiste%>");
					      parent.botones.document.location = "resumen_ventas_cli_bt.asp?mode=select1";
						parent.pantalla.resumen_ventas_cli.action="resumen_ventas_cli.asp?mode=select1";
						parent.pantalla.resumen_ventas_cli.submit();
						
					</script><%
				end if
				rst.close
			end if 'end if agrupa por cliente

			'*************************************** AGRUPACION POR MESES ************************
''ricardo 16-12-2003 REO comento con JAR que como en las demas agrupaciones si ponemos alguna parametro de articulo
'no se tienen en cuenta los conceptos, tampoco se tendran en cuenta en esta agrupacion

			if agrupar="MESES" then
				lote=limpiaCadena(Request.QueryString("lote"))
				'***** COD : JCI-090103-01 *****
				%><input type='hidden' name='elTotal' value='0'/><%
				'***** FIN COD : JCI-090103-01 *****
				strWhere=""
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
					strWhere=strWhere & " and doc.fecha<=convert(datetime,''" & fhasta & " 23:59:59'')"
					if falbaran = "on" then
						strfAlb= strfAlb & " and (al.fecha<=''" & fhasta & "'' or al.fecha is null)"
					end if
				end if

				if tactividad>"" then 'Se selecciono tipo de actividad
					desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
					strWhere=strWhere & " and c.tactividad=''" & tactividad & "''"
				end if
				if ncliente>"" and acliente>"" and ncliente=acliente then 'Se selecciono cliente
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitClienteMin%> : </b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><br/><%
					t_opcclientebaja=1
					strWhere=strWhere & " and c.ncliente=''" & session("ncliente") & ncliente & "''"
				else
					if opcclientebaja="" then
						t_opcclientebaja=1
					else
						t_opcclientebaja=0
						%><font class='ENCABEZADO'><b><%=LitClienteBaja%></b></font><br/><%
						strWhere=strWhere & " and c.fbaja is null"
					end if
				end if
				if ncliente>"" and acliente>"" and cliente<>acliente then
					t_opcclientebaja=1
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					acliente=Completar(acliente,5,"0")
					nomclihasta=d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & acliente & "'",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitDesdeCliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></font><%
					%> - <%
					%><font class='ENCABEZADO'><b><%=LitACliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(acliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomclihasta)%></font><br/><%

				      strWhere= strWhere+ " and c.ncliente>=''" & session("ncliente") & ncliente & "''"
				      strWhere= strWhere+ " and c.ncliente<=''" & session("ncliente") & acliente & "''"
				end if
				if provincia & "">"" then
					%><font class="ENCABEZADO"><b><%=LITPROVINCIA%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(provincia)%></font><br/><%
                    strWhere= strWhere+ " and dom.provincia like ''%" & provincia & "%'' "
				end if
				if serie>"" then 'Se selecciono serie
					%><font class='ENCABEZADO'><b><%=LitSerie%> : </b></font><font class='CELDA'><%=NombresEntidades(serie,"series","nserie","nombre",Session("backendlistados"))%></font><br/><%
					if instr(serie,",")>0 then
						strWhereSerie="(''" & replace(replace(serie," ",""),",","'',''") & "'')"
					else
						strWhereSerie=serie
					end if
				end if

				if familia>"" then
					tsel2 = false
					desc_familia=NombresEntidades(familia,"familias","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITSUBFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia)%></font><br/><%
					if instr(familia,",")>0 then
						strWhereArt = strWhereArt & " and a.familia in (''" & replace(replace(familia," ",""),",","'',''") & "'')"
					else
						strWhereArt=strWhereArt & " and a.familia=''" & familia & "''"
					end if
					sin_conceptos=1
				elseif familia_padre<>"" then
					tsel2 = false
					desc_familia_padre=NombresEntidades(familia_padre,"familias_padre","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia_padre)%></font><br/><%
					if instr(familia_padre,",")>0 then
						strWhereArt = strWhereArt & " and a.familia_padre in (''" & replace(replace(familia_padre," ",""),",","'',''") & "'')"
					else
						strWhereArt=strWhereArt & " and a.familia_padre=''" & familia_padre & "''"
					end if
					sin_conceptos=1
				elseif categoria<>"" then
					tsel2 = false
					desc_categoria=NombresEntidades(categoria,"categorias","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITCATEGORIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_categoria)%></font><br/><%
					if instr(categoria,",")>0 then
						strWhereArt = strWhereArt & " and a.categoria in (''" & replace(replace(categoria," ",""),",","'',''") & "'')"
					else
						strWhereArt=strWhereArt & " and a.categoria=''" & categoria & "''"
					end if
					sin_conceptos=1
				end if

				if referencia>"" then
					strWhereArt = strWhereArt + " and a.referencia like ''" & session("ncliente") & "%" & referencia & "%''"
					sin_conceptos=1
				end if
				if nombreart>"" then
					strWhereArt = strWhereArt + " and a.nombre like ''%" & nombreart & "%''"
					sin_conceptos=1
	   			end if

				if cod_proyecto>"" then
					%><font class='ENCABEZADO'><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",Session("backendlistados")))%></font><br/><%
					strWhere=strWhere & " and doc.cod_proyecto=''" & cod_proyecto & "''"
				end if

				if nproveedor & "">"" then
					strWhereArt = strWhereArt & " and a.referencia in (select articulo from proveer where nproveedor=''" & session("ncliente") & nproveedor & "'')"
					sin_conceptos=1
				end if
				if actividad_proveedor & "">"" then
					strWhereArt = strWhereArt & " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tactividad=''" & actividad_proveedor & "'')"
					sin_conceptos=1
				end if
				if tipo_proveedor & "">"" then
					strWhereArt = strWhereArt & " and a.referencia in (select articulo from proveer as pr,proveedores as pro where pr.nproveedor=pro.nproveedor and pro.tipo_proveedor=''" & tipo_proveedor & "'')"
					sin_conceptos=1
				end if
				if tipo_cliente & "">"" then
					strWhere=strWhere & " and c.tipo_cliente=''" & tipo_cliente & "''"
				end if

				if comercial & ""<>"" then
					if instr(comercial,",")>0 then
						strWhere=strwhere & " and doc.comercial in (''" & replace(replace(comercial," ",""),",","'',''") & "'')"
					else
						strWhere=strWhere & " and doc.comercial=''" & comercial & "''"
					end if
				end if

				if tipo_articulo & "">"" then
                    if instr(tipo_articulo,",")>0 then
						strWhereArt=strWhereArt & " and a.tipo_articulo in (''" & replace(replace(tipo_articulo," ",""),",","'',''") & "'')"
					else
						strWhereArt=strWhereArt & " and a.tipo_articulo=''" & tipo_articulo & "''"
					end if
					sin_conceptos=1
				end if

				if opc_cantidad>"" then
					%><font class='ENCABEZADO'><b><%=LitCantidad%><br/>
				<%end if

				if opc_ventasnetas>"" then
					%><font class='ENCABEZADO'><b><%=LitVentasNetas%><br/><%
				end if
				%><hr/><%

				if lote & "" = "" then
	        		'Para movimientros entre las páginas no borramos y volvemos a crear la tabla temporal ya que se creo y completó en la primera ejecución
	        		if instr(seriesTPF,",") > 0 then
						strWhereSeriesTPF ="'''" & replace(replace(seriesTPF," ",""),",","'',''") & "'''"
					else
						strWhereSeriesTPF = "'''" & seriesTPF & "'''"
					end if

					if instr(seriesAPF,",") > 0 then
						strWhereSeriesAPF ="'''" & replace(replace(seriesAPF," ",""),",","'',''") & "'''"
					else
						strWhereSeriesAPF = "'''" & seriesAPF & "'''"
					end if


					strQuery="Exec spl_ResumenVentasMeses @p_tablaTemp='[" & session("usuario") & "]', @p_strWhere='" & strWhere & "', @p_strWhereArt='" & strWhereArt & "', @p_serie='" & strWhereSerie & "', @p_soloConceptos=" & iif(ver_conceptos="",0,1) & ", @p_desgloseConceptos=" & iif(conceptos="",0,1) & ", @p_albPendFac=" & strWhereSeriesAPF & ", @p_ticPendFac=" & strWhereSeriesTPF & ",@mostrarfilas='" & mostrarfilas & "',@tipo='VENTAS',@sin_conceptos=" &  sin_conceptos & ",@p_nempresa='" & session("ncliente") &"',@p_falbaran='"& falbaran &"',@p_strfAlb='"& strfAlb &"'"

					'Llamar al procedimiento almacenado para crear la tabla temporal con los datos del listado.
					set conVentasMeses = Server.CreateObject("ADODB.Connection")
					conVentasMeses.open Session("backendlistados")
					conVentasMeses.CommandTimeout = 0
					conVentasMeses.execute(strQuery)
					conVentasMeses.close
					set conVentasMeses=nothing
				end if

				if mostrarfilas="CLIENTES" then
					if set_orden="on" then
						strMeses="select * from [" & session("usuario") & "] order by totalimpcliente desc,ncliente,nombre"
					else
						strMeses="select * from [" & session("usuario") & "] order by ncliente,nombre"
					end if
				else
					if set_orden="on" then
						strMeses="select * from [" & session("usuario") & "] order by sumtotalimp desc,nombre"
					else
						strMeses="select * from [" & session("usuario") & "] order by nombre"
					end if
				end if
				rst.cursorlocation=3
				rst.open strMeses,Session("backendlistados")
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

					referencia_actual=rst("referencia")
					referencia_old=rst("referencia")
					if mostrarfilas="CLIENTES" then
						cliente_old=rst("ncliente")
						cliente_actual=rst("ncliente")
					end if
					n_referencia=1
					hasta_referencia=((lote-1)*MAXPAGINA)+1
					while not rst.eof and n_referencia<hasta_referencia
						referencia_old=rst("referencia")
						if mostrarfilas="CLIENTES" then
							cliente_old=rst("ncliente")
						end if
						rst.movenext
						if not rst.eof then
							referencia_actual=rst("referencia")
							if mostrarfilas="CLIENTES" then
								cliente_actual=rst("ncliente")
							end if
						else
							referencia_actual="@@@#/===000"
							if mostrarfilas="CLIENTES" then
								cliente_actual="@@@#/===000"
							end if
						end if
						if referencia_actual<>referencia_old or cliente_actual<>cliente_old then
							n_referencia=n_referencia+1
						end if
					wend

					'-----------------------------------------%>
					<input type='hidden' name='NumRegs' value='<%=maxregistros%>'/><%
					NavPaginas lote,lotes,campo,criterio,texto,1
					if lotes>1 then
						%><hr/><%
					end if

					cliTemp=""
					refTemp=""
					fila=1
					totalCant=0
					totalImp=0

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
					numero_anyos=datediff("yyyy",fecha_desde,fecha_hasta,0)+1

					dim tm_lista(20,12,2)

					if mostrarfilas="CLIENTES" then
						ncliente_old=rst("ncliente")
					else
						referencia_old=rst("referencia")
					end if

					if mostrarfilas="ARTICULOS" then
						%><table width="100%" style="border-collapse: collapse;"><%
						'Mostrar la fila de encabezado para los artículos/meses'
						%><tr bgcolor="<%=color_terra%>">
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
								%><td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=EncodeForHtml(ucase(mid(DescMes(month(fecha_cual_vamos)),1,3)) & texto_ano & iif(opc_ventasnetas>""," (" & MB_abrev & ")",""))%></td><%
								fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
								if anyo<>year(fecha_cual_vamos) and poner_ano=1 then
									texto_ano=" - " & year(fecha_cual_vamos)
								else
									texto_ano=""
								end if
								anyo=year(fecha_cual_vamos)
							next%>
							<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalCantidad%></td>
							<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalVentas & " (" & EncodeForHtml(MB_abrev) & ")"%></td>
						</tr><%
					end if

					while not rst.eof and fila<=MAXPAGINA
						if mostrarfilas="CLIENTES" then
							if left(rst("Ncliente"),5)="-----" then
							else
								CheckCadena rst("NCliente")
							end if

							if rst("ncliente")<>cliTemp then
								for i=1 to 12
									for j=1 to numero_anyos
										tm_lista(j,i,1)=0
										tm_lista(j,i,2)=0
									next
								next
								'Fila de encabezado del cliente
								if cliTemp="" then
									%><table width="100%" style="border-collapse: collapse;"><%
								else
									''ricardo 22/1/2003 para que se veo o no lo de clientes en hojas separadas, ya que solo sirve para la agrupacion por cliente
									if clihojassep="on" or clihojassep="1" then
										%></table>
										<h6 class=SALTO>&nbsp;</h6>
										<table width="100%" style="border-collapse: collapse;"><%
									end if
								end if
								%><tr bgcolor="<%=color_fondo%>">
									<td class="ENCABEZADOL" colspan="16"><%=LitClienteMin & ": " & EncodeForHtml(trimCodEmpresa(rst("ncliente")) & " - " & rst("rsocial"))%></td>
								</tr><%
								''if session("ncliente")=PERSO_SIALC then
								if verdc="1" then
                                    rstAux.cursorlocation=3
									rstAux.open "select * from domicilios with(NOLOCK) where tipo_domicilio='PRINCIPAL_CLI' and pertenece='" & rst("ncliente") & "' order by codigo desc",Session("backendlistados")
									%><tr bgcolor="<%=color_fondo%>">
										<td class="CELDA" colspan="16"><%=enc.EncodeForHtmlAttribute(null_s(rstAux("domicilio")))%>&nbsp;<%=enc.EncodeForHtmlAttribute(null_s(rstAux("poblacion")))%>&nbsp;(<%=enc.EncodeForHtmlAttribute(null_s(rstAux("provincia")))%>)&nbsp;<b>Tel : </b><%=enc.EncodeForHtmlAttribute(null_s(rstAux("telefono")))%>
										<%rstAux.close
                                        rstAux.cursorlocation=3
										rstAux.open "select c.ncliente,ta.descripcion from clientes c left outer join tipo_actividad ta on c.tactividad=ta.codigo where c.ncliente='" & rst("ncliente") & "'",Session("backendlistados")
										response.write("&nbsp;<b>" & LitActividad & " : </b>" & rstAux("descripcion")&"")
										rstAux.close%>
										</td>
									</tr><%
								end if

								'Mostrar la fila de encabezado para los artículos/meses'
								%><tr bgcolor="<%=color_terra%>">
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
										end if
										%><td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=EncodeForHtml(ucase(mid(DescMes(month(fecha_cual_vamos)),1,3)) & texto_ano & iif(opc_ventasnetas>""," (" & MB_abrev & ")",""))%></td><%
										fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
										if anyo<>year(fecha_cual_vamos) and poner_ano=1 then
											texto_ano=" - " & year(fecha_cual_vamos)
										else
											texto_ano=""
										end if
										anyo=year(fecha_cual_vamos)
									next%>
									<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalCantidad%></td>
									<td class="ENCABEZADOR7" style="border: 1px solid Black;"><%=LitTotalVentas & " (" & EncodeForHtml(MB_abrev) & ")"%></td>
								</tr><%

								cliTemp=rst("ncliente")
							end if
						else
							if rst("referencia")<>refTemp then
								for i=1 to 12
									for j=1 to numero_anyos
										tm_lista(j,i,1)=0
										tm_lista(j,i,2)=0
									next
								next

								refTemp=rst("referencia")
							end if
						end if

						%><tr>
							<td class="CELDAL7" style="border: 1px solid Black;"><%=EncodeForHtml(trimCodEmpresa(rst("referencia")))%></td>
							<td class="CELDAL7" style="border: 1px solid Black;"><%=enc.EncodeForHtmlAttribute(null_s(rst("nombre")))%></td>
							<%

							sumtotalcant=rst("sumtotalcant")
							sumtotalimp=rst("sumtotalimp")
							fecha_cual_vamos=fecha_desde
							referencia_old2=rst("referencia")
							i=1
							while i<=numero_meses and not rst.eof
								if mostrarfilas="CLIENTES" then
									if referencia_old2=rst("referencia") and cliTemp=rst("ncliente") then
										if (month(fecha_cual_vamos)<>cint(rst("mesdoc")) or year(fecha_cual_vamos)<>cint(rst("anyodoc"))) then
											%><td class="CELDAR7" style="border: 1px solid Black;"></td><%
										else
											mostrar_Cant_Meses "CELDAR7",opc_cantidad,rst("totalcant"),rst("totalimp"),n_decimalesMB
											rst.movenext
										end if
										fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
										i=i+1
									else
										%><td class="CELDAR7" style="border: 1px solid Black;"></td><%
										i=i+1
									end if
								else
									if referencia_old2=rst("referencia") then
										if (month(fecha_cual_vamos)<>cint(rst("mesdoc")) or year(fecha_cual_vamos)<>cint(rst("anyodoc"))) then
											%><td class="CELDAR7" style="border: 1px solid Black;"></td><%
										else
											mostrar_Cant_Meses "CELDAR7",opc_cantidad,rst("totalcant"),rst("totalimp"),n_decimalesMB
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
							for j=i to numero_meses
								%><td class="CELDAR7" style="border: 1px solid Black;"></td><%
							next
							rst.moveprevious
							%>
							<%mostrar_Cant_Meses "CELDAR7","---",sumtotalcant,"0",n_decimalesMB%>
							<%mostrar_Cant_Meses "CELDAR7","---",sumtotalimp,"0",n_decimalesMB%>
						</tr><%

						totalimpcliente=rst("totalimpcliente")
						totalcantcliente=rst("totalcantcliente")
						if mostrarfilas="CLIENTES" then
							ncliente_old=rst("ncliente")
						else
							referencia_old=rst("referencia")
						end if

						fila=fila+1
						rst.movenext

						if rst.eof then
							verTotales=1
						else
							if mostrarfilas="CLIENTES" then
								if rst("ncliente")<>cliTemp then
									verTotales=1
								end if
							else
								if rst("referencia")<>refTemp then
									verTotales=0
								end if
							end if
						end if

						if verTotales=1 then
							if mostrarfilas="CLIENTES" then
								strPorMeses="select sum(totalimp) as sumtotalimp,sum(totalcant) as sumtotalcant,mesdoc,anyodoc from [" & session("usuario") & "] where ncliente='" & ncliente_old & "' group by anyodoc,mesdoc order by convert(int,anyodoc),convert(int,mesdoc)"
							else
								strPorMeses="select sum(totalimp) as sumtotalimp,sum(totalcant) as sumtotalcant,mesdoc,anyodoc from [" & session("usuario") & "]  group by anyodoc,mesdoc order by convert(int,anyodoc),convert(int,mesdoc)"
							end if
							rstAux.cursorlocation=3
							rstAux.open strPorMeses,Session("backendlistados")

							if not rstAux.eof then
								anyo=1
								fecha_cual_vamos=fecha_desde
								anyoActual=year(fecha_cual_vamos)
							end if
							while not rstAux.eof
								if year(fecha_cual_vamos)<>cint(rstAux("anyodoc")) then
									anyo=anyo+1
									fecha_cual_vamos=dateadd("m",12,fecha_cual_vamos)
									anyoActual=year(fecha_cual_vamos)
								else
									tm_lista(anyo,cint(rstAux("mesdoc")),1)=tm_lista(anyo,cint(rstAux("mesdoc")),1) + rstAux("sumtotalcant")
									tm_lista(anyo,cint(rstAux("mesdoc")),2)=tm_lista(anyo,cint(rstAux("mesdoc")),2) + rstAux("sumtotalimp")
									rstAux.movenext
								end if
							wend
							rstAux.close

							'Mostrar la fila de totales por cliente/referencia.
							%><tr bgcolor="<%=color_terra%>">
								<td class="ENCABEZADOL7" style="border: 1px solid Black;">&nbsp;</td>
								<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitTotal%></td>
								<%
								fecha_cual_vamos=fecha_desde
								anyo=1
								anyoActual=year(fecha_cual_vamos)
								for i=1 to numero_meses
									mes_a_poner=month(fecha_cual_vamos)
									mostrar_Cant_Meses "ENCABEZADOR7","---",iif(opc_cantidad>"",tm_lista(anyo,mes_a_poner,1),tm_lista(anyo,mes_a_poner,2)),"0",n_decimalesMB
									fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
									if anyoActual<>year(fecha_cual_vamos) then
										anyo=anyo+1
										anyoActual=year(fecha_cual_vamos)
									end if
								next
								%>
								<%mostrar_Cant_Meses "ENCABEZADOR7","---",totalcantcliente,"0",n_decimalesMB%>
								<%mostrar_Cant_Meses "ENCABEZADOR7","---",totalimpcliente,"0",n_decimalesMB%>
							</tr><%
							totalCant=0
							totalImp=0
							verTotales=0
						end if
					wend

					''mostramos ahora una fila de totales del listado
					if mostrarfilas="CLIENTES" and rst.eof then
						for i=1 to 12
							for j=1 to numero_anyos
								tm_lista(j,i,1)=0
								tm_lista(j,i,2)=0
							next
						next
						strPorMeses="select sum(totalimp) as sumtotalimp,sum(totalcant) as sumtotalcant,mesdoc,anyodoc,(select sum(totalimp) from [" & session("usuario") & "]) as totalimp,(select sum(totalcant) from [" & session("usuario") & "]) as totalcant from [" & session("usuario") & "] group by anyodoc,mesdoc order by convert(int,anyodoc),convert(int,mesdoc)"
						rstAux.cursorlocation=3
						rstAux.open strPorMeses,Session("backendlistados")
						if not rstAux.eof then
							sumatotalimp=rstAux("totalimp")
							sumatotalcant=rstAux("totalcant")
							anyo=1
							fecha_cual_vamos=fecha_desde
							anyoActual=year(fecha_cual_vamos)
						end if
						while not rstAux.eof
							if anyoActual=cint(rstAux("anyodoc")) then
								tm_lista(anyo,cint(rstAux("mesdoc")),1)=tm_lista(anyo,cint(rstAux("mesdoc")),1) + rstAux("sumtotalcant")
								tm_lista(anyo,cint(rstAux("mesdoc")),2)=tm_lista(anyo,cint(rstAux("mesdoc")),2) + rstAux("sumtotalimp")
								rstAux.movenext
							else
								anyo=anyo+1
								fecha_cual_vamos=dateadd("m",12,fecha_cual_vamos)
								anyoActual=year(fecha_cual_vamos)
							end if
						wend
						rstAux.close

						'Mostrar la fila de totales por cliente/referencia.'
						%><tr bgcolor="<%=color_terra%>">
							<td class="ENCABEZADOL7" style="border: 1px solid Black;"><%=LitTotalListado%></td>
							<td class="ENCABEZADOL7" style="border: 1px solid Black;">&nbsp;</td>
							<%
							fecha_cual_vamos=fecha_desde
							anyo=1
							anyoActual=year(fecha_cual_vamos)
							for i=1 to numero_meses
								mes_a_poner=month(fecha_cual_vamos)
								mostrar_Cant_Meses "ENCABEZADOR7","---",iif(opc_cantidad>"",tm_lista(anyo,mes_a_poner,1),tm_lista(anyo,mes_a_poner,2)),"0",n_decimalesMB
								fecha_cual_vamos=dateadd("m",1,fecha_cual_vamos)
								if anyoActual<>year(fecha_cual_vamos) then
									anyo=anyo+1
									anyoActual=year(fecha_cual_vamos)
								end if
							next%>
							<%mostrar_Cant_Meses "ENCABEZADOR7","---",sumatotalcant,"0",n_decimalesMB%>
							<%mostrar_Cant_Meses "ENCABEZADOR7","---",sumatotalimp,"0",n_decimalesMB%>
						</tr><%
					end if

					%></table><%

					NavPaginas lote,lotes,campo,criterio,texto,2
				else
					%><script language="javascript" type="text/javascript">
					      alert("<%=LitMsgDatosNoExiste%>");
					      parent.botones.document.location = "resumen_ventas_cli_bt.asp?mode=select1"
						parent.pantalla.resumen_ventas_cli.action="resumen_ventas_cli.asp?mode=select1";
						parent.pantalla.resumen_ventas_cli.submit();
						
					</script><%
				end if
				rst.close
			end if 'end if agrupa por meses

			'*************************************** AGRUPACION POR PROYECTOS ********************
			if agrupar="PROYECTO" then
				PorProyecto = "SI"
				opc_cod_proyecto="1"
				if tactividad>"" then 'Se selecciono tipo de actividad
					desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
				end if
				if ncliente>"" and acliente>"" and ncliente=acliente then 'Se selecciono cliente
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitClienteMin%> :</b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><br/><%
					t_opcclientebaja=1
				else
					if opcclientebaja="" then
						t_opcclientebaja=1
					else
						t_opcclientebaja=0
						%><font class='ENCABEZADO'><b><%=LitClienteBaja%></b></font><br/><%
					end if
				end if
				if ncliente>"" and acliente>"" and ncliente<>acliente then
					t_opcclientebaja=1
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					acliente=Completar(acliente,5,"0")
					nomclihasta=d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & acliente & "'",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitDesdeCliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><%
					%> - <%
					%><font class='ENCABEZADO'><b><%=LitACliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(acliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomclihasta)%></b></font><br/><%
				end if

				if serie>"" then 'Se selecciono serie
					%><font class='ENCABEZADO'><b><%=LitSerie%> : </b></font><font class='CELDA'><%=NombresEntidades(serie,"series","nserie","nombre",Session("backendlistados"))%></b></font><br/><%
				end if

				if familia>"" then
					tsel2 = false
					desc_familia=NombresEntidades(familia,"familias","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITSUBFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia)%></b></font><br/><%
				elseif familia_padre<>"" then
					tsel2 = false
					desc_familia_padre=NombresEntidades(familia_padre,"familias_padre","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia_padre)%></b></font><br/><%
				elseif categoria<>"" then
					tsel2 = false
					desc_categoria=NombresEntidades(categoria,"categorias","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITCATEGORIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_categoria)%></b></font><br/><%
				end if

				if cod_proyecto>"" then
					PorProyecto="NO"
					%><font class='ENCABEZADO'><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",Session("backendlistados")))%></font><br/><%
				end if
				if provincia>"" then
					%><font class="ENCABEZADO"><b><%=LITPROVINCIA%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(provincia)%></font><br/><%
				end if
				if saveR="true" then
					crearProyecto tsel1, tsel2,seriesAPF, fdesde, fhasta, tactividad, iif(cliente>"",session("ncliente") & cliente,cliente),iif(acliente>"",session("ncliente") & acliente,acliente), serie, familia, familia_padre, categoria, referencia, nombreart, ordenar, MB,t_opcclientebaja,cod_proyecto,seriesTPF,iif(nproveedor>"",session("ncliente") & nproveedor,nproveedor),actividad_proveedor,tipo_cliente,tipo_proveedor,comercial,tipo_articulo,mostrarfilas,provincia
				else
					%><input type='hidden' name='elTotal' value='<%=EncodeForHtml(elTotal)%>' /><%
				end if
				%><hr/><%
				if seriesAPF > "" then
					seleccion="SELECT cod_proyecto,NCliente,Nombre,Referencia,Descripcion,"
					seleccion=seleccion & "SUM(cantidad) AS Cantidad"
					seleccion=seleccion & ",SUM(cantidad2) AS Cantidad2"
					seleccion=seleccion & ",case when convert(nvarchar,calculoimporte)=1 then 'VERDADERO' else 'FALSO' end as calculoimporte"
					seleccion=seleccion & ",tipo_medida"
					seleccion=seleccion & ", SUM([Ventas Netas]) AS [Ventas Netas],"
					seleccion=seleccion & "SUM([Precio Medio] * cantidad)/ CASE WHEN SUM(cantidad)= 0 THEN 1 ELSE SUM(cantidad) END AS [Precio Medio],"
					seleccion=seleccion & "SUM([Precio Medio2] * cantidad2)/ CASE WHEN SUM(cantidad2)= 0 THEN 1 ELSE SUM(cantidad2) END AS [Precio Medio2],"
					seleccion=seleccion & "Divisa,(select sum([Ventas Netas]) from [" & session("usuario") & "] where cod_proyecto=f.cod_proyecto) as Acumulador,Orden"
					seleccion=seleccion & ",tiene_escv "
					seleccion=seleccion & " FROM [" & session("usuario") & "] as f " & strbaja
					seleccion=seleccion & " GROUP BY cod_proyecto,Ncliente,Nombre,Referencia,Descripcion,divisa,orden,tiene_escv,convert(nvarchar,calculoimporte),tipo_medida "
					if ordenar = true then
						seleccion=seleccion & " order by orden desc,cod_proyecto"
					end if
				else
					if ordenar = true then
						seleccion = "select * from [" & session("usuario") & "]" & strbaja & " order by orden desc,cod_proyecto"
					else
						seleccion = "select * from [" & session("usuario") & "]" & strbaja & " order by cod_proyecto"
					end if
				end if
				'comprobamos si existen varias divisas
                rst.cursorlocation=3
				rst.open seleccion,Session("backendlistados")
				MostrarDivisa = false
				w_divisa=""
				if not rst.eof then
					w_divisa = rst("Divisa")
					ocultar = 4
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
					<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'/><%
					NavPaginas lote,lotes,campo,criterio,texto,1
					%><br/>


					<table width="100%" style="border-collapse: collapse;">
						<%
							if proyecto="" then
								%><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-5-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=LitProyecto%>:
									<%=enc.EncodeForHtmlAttribute(null_s(rst("cod_proyecto")))%>
								</td><%
							end if
							DrawFila color_fondo%>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNCliente%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitVentasNetas%></td>
								<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td><%
								if MostrarDivisa=true then
									%><td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
								end if
							CloseFila
							ValorCampoAnt = ""
							AcumuladoAnt  = ""
							ImprimeCabecera = false
							fila = 1
							while not rst.eof and fila<=MAXPAGINA
								if left(rst("Ncliente"),5)="-----" then
								else
									CheckCadena rst("NCliente")
								end if
								DrawFila ""
									for each campo in rst.fields
										if campo.name="tiene_escv" then
										else
											if (PorProyecto="SI" and campo.name="cod_proyecto") then
												if rst(campo.name)<>ValorCampoAnt then
													if ValorCAmpoAnt<>"" then
														'antes de imprimir subtotales imprimimos conceptos (si existen)
														'Fila de Subtotal%>
														<td></td>
														<td></td>
														<td></td>
														<%if opc_cod_proyecto="1" then%>
															<td></td>
														<%end if%>
														<td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
			                                        					<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
														<%
														CloseFila
														DrawFila "" 'Fila de separacion
															%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
														CloseFila
														DrawFila color_fondo
		   									                        %><td class="ENCABEZADOL" style="border: 1px solid Black; "height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-5-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=LitProyecto%>:
																<%=enc.EncodeForHtmlAttribute(null_s(rst("cod_proyecto")))%>
															</td><%
														CloseFila
														DrawFila color_fondo%>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNCliente%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitReferencia%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitDescripcion%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
															<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"><%=LitVentasNetas%></td>
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
											elseif campo.name="Cantidad" or campo.name="Cantidad2" or campo.name="calculoimporte" or campo.name="tipo_medida" or campo.name="Precio Medio2" or campo.name="Ventas Netas" or campo.name="Precio Medio" then 'Formateo del campo con importe
												'ajustamos divisas si es necesario
												if MostrarDivisa=true then
												   n_decimales = null_z(d_lookup("ndecimales", "divisas", "codigo='" & rst("Divisa") & "'", Session("backendlistados")))
												end if
												if campo.name="Ventas Netas" then
	 											   AcumuladoAnt = rst("Acumulador")
												end if
												''ricardo 22-3-2004
												tipo_medida=rst("tipo_medida")
												valor_cantidad2=null_z(rst("Cantidad2"))
												calculoimporte=nz_b(rst("calculoimporte"))
												''ricardo 22-3-2004
												if (campo.name<>"cod_proyecto" or opc_cod_proyecto="1") and campo.name<>"tipo_medida" then
													%><!--<td class='TDBORDECELDA7' align="right">-->
														<%if rst("tiene_escv")<>1 or campo.name="Cantidad" or campo.name="Cantidad2" then
															''ricardo 22-3-2004
															if campo.name="Cantidad" or campo.name="Cantidad2" then
																if campo.name="Cantidad" then
																	%><td class='TDBORDECELDA7' align="right"><%
																		if valor_cantidad2<>0 then%>
																			<%=EncodeForHtml(formatnumber(rst(campo.name),DEC_CANT,-1,0,-1))%>
																			<br/>
																			<%=EncodeForHtml(("<b>" & iif(tipo_medida>"",tipo_medida,"") & " : </b>" & formatnumber(valor_cantidad2,DEC_CANT,-1,0,-1)))%>
																		<%else%>
																			<%=EncodeForHtml(formatnumber(rst(campo.name),DEC_CANT,-1,0,-1))%>
																		<%end if
																	%></td><%
																end if
															else
																if campo.name="Precio Medio" or campo.name="Precio Medio2" or campo.name="calculoimporte" then
																	if campo.name="Precio Medio" then%>
																		<td class='TDBORDECELDA7' align="right">
																			<%=EncodeForHtml(formatnumber(rst(campo.name),n_decimales,-1,0,-1))%>
																			<%if calculoimporte=-1 and rst("Precio Medio2")<>0 then%>
																				<br/>
																				<%=EncodeForHtml(("<b>" & iif(tipo_medida>"",tipo_medida,"") & " : </b>" & formatnumber(rst("Precio Medio2"),n_decimales,-1,0,-1)))%>
																			<%end if%>
																		</td>
																	<%end if
																else%>
																	<td class='TDBORDECELDA7' align="right">
																		<%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",formatnumber(rst(campo.name),iif(campo.name<>"Cantidad",n_decimales,DEC_CANT),-1,0,-1)))%>
																	</td>
																<%end if%>
															<%end if%>
														<%else%>
															<%if campo.name<>"Cantidad" and campo.name<>"Cantidad2" and campo.name<>"calculoimporte" and campo.name<>"Precio Medio2" then%>
																<td class='TDBORDECELDA7' align="right">&nbsp;</td>
															<%end if%>
														<%end if%>
													<!--</td>--><%
												end if
											else
												if campo.name<>"cod_proyecto" and campo.name<>"Divisa" then
													if campo.name="Referencia" then
												      	if rst(campo.name)="zzzzzzzzzzzzzzzzzzzz" then
															%><td class='tdbordecelda7'>Concepto</td><%
														else
															%><td class='tdbordecelda7'>
																<%=Hiperv(OBJArticulos,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerArticulo)%>
															</td><%
														end if
													else
            	                                  						if campo.name<>"Orden" and campo.name<>"Acumulador" then
															if campo.name="NCliente" or ucase(campo.name)="NCLIENTE" then
																%><td class='tdbordecelda7'>
																	<%=Hiperv(OBJClientes,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerCliente)%>
																</td><%
															else
														  		if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
														         		%><td class='tdbordecelda7'><%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",rst(campo.name)))%></td><%
																end if
															end if
													  	else
													     		%><td></td><%
													  	end if
												   	end if
												else
													if campo.name="Divisa" and MostrarDivisa=true then
											      		%><td class='tdbordecelda7'><%=EncodeForHtml(d_lookup("abreviatura", "divisas", "codigo='" & rst(campo.name) & "' and codigo like '" & session("ncliente") & "%'", Session("backendlistados")))%></td><%
													else
							                          	if campo.name = "Divisa" and MostrarDivisa = false then '
								                        	%><td></td><%
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
										<%if opc_cod_proyecto="1" then%>
											<td></td>
										<%end if%>
										<td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
										<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(AcumuladoAnt,n_decimalesMB,-1,0,-1))%></b></td>
								   	<%CloseFila
								   	DrawFila "" 'Fila de separacion
										%><td colspan="<%=rst.fields.count-2-cint(iif(opc_cod_proyecto="1","0","1"))%>">&nbsp;</td><%
     									CloseFila
								end if
								DrawFila "" 'Fila para el total
									%><td></td>
									<td></td>
									<td></td>
									<%if opc_cod_proyecto="1" then%>
										<td></td>
									<%end if%>
									<% Suma = elTotal%>
									<td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
									<td class='TDBORDECELDA7' align="right"><big><b><%=EncodeForHtml(formatnumber(Suma,n_decimalesMB,-1,0,-1))%></b></big></td>
								<%CloseFila%>
								<%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",Session("backendlistados")) then
									DrawFila "" 'Fila para el total equivalencia en PTS
					      				%><td></td>
							   			<td></td>
									<td></td>
										<%if opc_cod_proyecto="1" then%>
											<td></td>
										<%end if%>
					      				<td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",Session("backendlistados")))%>:</b></td>
				      					<td class='TDBORDECELDA7' align="right"><big><b><%=EncodeForHtml(formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",Session("backendlistados")),-1,0,-1))%></b></big></td>
									<%CloseFila%>
								<%end if%>
							<%end if%>

					</table><br/><%
					NavPaginas lote,lotes,campo,criterio,texto,2
				else
					%><script language="javascript" type="text/javascript">
					      alert("<%=LitMsgDatosNoExiste%>");
					      parent.botones.document.location = "resumen_ventas_cli_bt.asp?mode=select1";
                        parent.pantalla.resumen_ventas_cli.action="resumen_ventas_cli.asp?mode=select1";
                        parent.pantalla.resumen_ventas_cli.submit();
						
					</script><%
				end if
				rst.close
			end if 'end if agrupa por proyecto

'********************************************** AGRUPACION POR ARTICULOS *************************
			if agrupar="ARTICULO" then
				PorArticulo = "SI"

				if familia>"" then
					tsel2 = false
					desc_familia=NombresEntidades(familia,"familias","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITSUBFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia)%></b></font><br/><%
				elseif familia_padre<>"" then
					tsel2 = false
					desc_familia_padre=NombresEntidades(familia_padre,"familias_padre","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITFAMILIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia_padre)%></b></font><br/><%
				elseif categoria<>"" then
					tsel2 = false
					desc_categoria=NombresEntidades(categoria,"categorias","codigo","nombre",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LITCATEGORIA%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_categoria)%></b></font><br/><%
				end if

				if tactividad>"" then 'Se selecciono tipo de actividad
					desc_tactividad = d_lookup("descripcion", "tipo_actividad", "codigo='" & tactividad & "'", Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitActividad%> : </b></font><font class='CELDA'><%=EncodeForHtml(trimCodEmpresa(tactividad))%>&nbsp;&nbsp;<%=EncodeForHtml(desc_tactividad)%></b></font><br/><%
	      		end if
				if ncliente>"" and acliente>"" and ncliente=acliente then 'Se selecciono cliente
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitClienteMin%> : </b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><br/><%
					t_opcclientebaja=1
				else
					if opcclientebaja="" then
						t_opcclientebaja=1
					else
						t_opcclientebaja=0
						%><font class='ENCABEZADO'><b><%=LitClienteBaja%></b></font><br/><%
					end if
				end if
				if ncliente>"" and acliente>"" and ncliente<>acliente then
					t_opcclientebaja=1
					nomcli = d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & ncliente & "'",Session("backendlistados"))
					acliente=Completar(acliente,5,"0")
					nomclihasta=d_lookup("rsocial","clientes","ncliente='" & session("ncliente") & acliente & "'",Session("backendlistados"))
					%><font class='ENCABEZADO'><b><%=LitDesdeCliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(ncliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomcli)%></b></font><%
					%> - <%
					%><font class='ENCABEZADO'><b><%=LitACliente%> : </b></font><font class='CELDA'><%=EncodeForHtml(acliente)%>&nbsp;&nbsp;<%=EncodeForHtml(nomclihasta)%></b></font><br/><%
				end if
				if provincia>"" then
					%><font class="ENCABEZADO"><b><%=LITPROVINCIA%>:&nbsp;</b></font><font class='CELDA'><%=EncodeForHtml(provincia)%></font><br/><%
				end if
				if serie>"" then 'Se selecciono serie
					%><font class='ENCABEZADO'><b><%=LitSerie%> : </b></font><font class='CELDA'><%=NombresEntidades(serie,"series","nserie","nombre",Session("backendlistados"))%></b></font><br/><%
				end if
				%><hr/><%
		     lote=limpiaCadena(Request.QueryString("lote"))
		     if lote="" then lote=1
		     sentido=limpiaCadena(Request.QueryString("sentido"))
		     if sentido="next" then
		       lote=lote+1
		     elseif sentido="prev" then
		       lote=lote-1
		    end if
			crearArticulo tsel1, tsel2,seriesAPF, fdesde,fhasta, familia, familia_padre, categoria, referencia, nombreart, tactividad, iif(ncliente>"",session("ncliente") & ncliente,ncliente),iif(acliente>"",session("ncliente") & acliente,acliente), serie, desglose, ordenar, MB,t_opcclientebaja,cod_proyecto,seriesTPF,opc_cod_proyecto,iif(nproveedor>"",session("ncliente") & nproveedor,nproveedor),actividad_proveedor,tipo_cliente,tipo_proveedor,comercial,tipo_articulo,mostrarfilas,provincia',importeIva

			if seriesAPF > "" or seriesTPF > "" then
				seleccion="SELECT Ref, Descripcion,NCliente,Nombre,"
				if opc_cod_proyecto="1" then
					seleccion=seleccion & "cod_proyecto,"
				end if
				seleccion=seleccion & "SUM(cantidad) AS Cantidad"
				seleccion=seleccion & ",SUM(cantidad2) AS Cantidad2"
				seleccion=seleccion & ",case when convert(nvarchar,calculoimporte)=1 then 'VERDADERO' else 'FALSO' end as calculoimporte"
				seleccion=seleccion & ",tipo_medida"
				seleccion=seleccion & ", SUM([Ventas Netas]) AS [Ventas Netas],"
				seleccion=seleccion & "SUM([Precio Medio] * cantidad)/ CASE WHEN SUM(cantidad)= 0 THEN 1 ELSE SUM(cantidad) END AS [Precio Medio],"
				seleccion=seleccion & "SUM([Precio Medio2] * cantidad2)/ CASE WHEN SUM(cantidad2)= 0 THEN 1 ELSE SUM(cantidad2) END AS [Precio Medio2],"
				seleccion=seleccion & "Divisa"
				seleccion=seleccion & ",(select sum(CASE WHEN divisa = '" & session("ncliente") & "02' THEN [Ventas Netas] ELSE ([Ventas Netas] / 166.386) END) "
				seleccion=seleccion & "from [" & session("usuario") & "] WHERE Ref = f.Ref) as AcumulaVentas,(select sum(cantidad) from [" & session("usuario") & "] WHERE Ref = f.Ref) as AcumulaCantidad"
				seleccion=seleccion & ",Orden"
				seleccion=seleccion & ",tiene_escv "
				seleccion=seleccion & " FROM [" & session("usuario") & "] as f" & strbaja
				seleccion=seleccion & " GROUP BY ref,Ncliente,Nombre,descripcion,divisa,orden,tiene_escv,convert(nvarchar,calculoimporte),tipo_medida "
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

			rst.cursorlocation=3
			rst.open seleccion,Session("backendlistados")
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
				<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>' /><%
				NavPaginas lote,lotes,campo,criterio,texto,1%><br/><%

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
				end if%>
			<table width="100%" style="border-collapse: collapse;">
				<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-5-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=EncodeForHtml(etiq)%>:
					<%=Hiperv(OBJArticulos,ref,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(ref) & "&nbsp;&nbsp;" & desc),LitVerArticulo)%>
				</td><%
				DrawFila color_fondo%>
					<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitNCliente%></td>
					<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
					<%'if campo.name="xxxxxxxxxxxxcod_proyecto" or opc_cod_proyecto="1" then
					if opc_cod_proyecto="1" then%>
						<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
					<%end if%>
					<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
					<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitVentasNetas%></td>
					<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td><%
                    'INICIO AÑADIR IMPORTE MEDIO IVA
                    if importeIva=true then
						%><td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LITIMPORTEMEDIOIVA%></td><%
					end if
                    'FIN AÑADIR IMPORTE MEDIO IVA
					if MostrarDivisa = true then
						%><td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
					end if                    
				CloseFila
				ValorCampoAnt = ""
				fila = 1
				while not rst.eof and fila<=MAXPAGINA
					if left(rst("Ncliente"),5)="-----" then
					else
						CheckCadena rst("NCliente")
					end if
					DrawFila ""
						for each campo in rst.fields
							if campo.name="tiene_escv" then
							else
								if (PorArticulo="SI" and campo.name="Ref") then
									if ucase(rst(campo.name))<>ValorCampoAnt then
										if ValorCAmpoAnt<>"" then
											'antes de imprimir subtotales imprimimos conceptos (si existen)
											'Fila de Subtotal%>
											<td></td>
											<%if opc_cod_proyecto="1" then%>
												<td></td>
											<%end if%>
											<td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
											<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(SubTotalCant,DEC_CANT,-1,0,-1))%></b></td>
											<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(SubTotal,n_decimalesMB,-1,0,-1))%></b></td>
											<%
											pmTotal=0
											if SubTotalCant<>0 then pmTotal = cdbl(SubTotal)/SubTotalCant
											%>
											<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(pmTotal,n_decimalesMB,-1,0,-1))%></b></td>
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
												<td class="ENCABEZADOL" style="border: 1px solid Black;" height="15"; bgcolor="<%=color_fondo%>"; colspan="<%=rst.fields.count-ocultar-5-cint(iif(opc_cod_proyecto="1","0","1"))%>"><%=etiq%>:
													<%=Hiperv(OBJArticulos,ref,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(ref) & "&nbsp;&nbsp;" & desc),LitVerArticulo)%>
												</td><%
											CloseFila
											DrawFila color_fondo%>
												<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitNCliente%></td>
												<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitNombre%></td>
												<%if campo.name="xxxxxxxxxxxxcod_proyecto" or opc_cod_proyecto="1" then%>
													<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitProyecto%></td>
												<%end if%>
												<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitCantidad%></td>
												<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitVentasNetas%></td>
												<td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitPrecioMedio%></td><%
                                                'INICIO AÑADIR IMPORTE MEDIO IVA
                                                if importeIva=true then
						                            %><td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LITIMPORTEMEDIOIVA%></td><%
					                            end if
                                                'FIN AÑADIR IMPORTE MEDIO IVA
												if MostrarDivisa = true then
													%><td class="ENCABEZADOL" bgcolor="<%=color_terra%>" style="border: 1px solid Black;" height="15"><%=LitDivisa%></td><%
												end if                                                
											CloseFila
										end if 'valorcampoant<>""
									end if 'rst(campo.name)<>ValorCampoAnt
									ValorCampoAnt=ucase(rst(campo.name))
								elseif campo.name="Cantidad" or campo.name="Cantidad2" or campo.name="calculoimporte" or campo.name="tipo_medida" or campo.name="Precio Medio2" or campo.name="Ventas Netas" or campo.name="Precio Medio" or campo.name="Precio Medio2" then 'Formateo del campo con importe
									'ajustamos divisas si es necesario
									if MostrarDivisa=true then
										n_decimales = null_z(d_lookup("ndecimales", "divisas", "codigo='" & rst("Divisa") & "'", Session("backendlistados")))
									end if
									''ricardo 22-3-2004
									tipo_medida=rst("tipo_medida")
									valor_cantidad2=null_z(rst("Cantidad2"))
									calculoimporte=nz_b(rst("calculoimporte"))
									if campo.name<>"tipo_medida" then
										%><!--<td class='TDBORDECELDA7' align="right">-->
											<%if rst("tiene_escv")<>1 or campo.name="Cantidad" or campo.name="Cantidad2" then
												''ricardo 22-3-2004
												if campo.name="Cantidad" or campo.name="Cantidad2" then
													if campo.name="Cantidad" then
														%><td class='TDBORDECELDA7' align="right"><%
															if valor_cantidad2<>0 then%>
																<%=EncodeForHtml(formatnumber(rst(campo.name),DEC_CANT,-1,0,-1))%>
																<br/>
																<%=EncodeForHtml(("<b>" & iif(tipo_medida>"",tipo_medida,"") & " : </b>" & formatnumber(valor_cantidad2,DEC_CANT,-1,0,-1)))%>
															<%else%>
																<%=EncodeForHtml(formatnumber(rst(campo.name),DEC_CANT,-1,0,-1))%>
															<%end if
														%></td><%
													end if
												else
													if campo.name="Precio Medio" or campo.name="Precio Medio2" or campo.name="calculoimporte" then
														if campo.name="Precio Medio" then%>
															<td class='TDBORDECELDA7' align="right">
																<%=EncodeForHtml(formatnumber(rst(campo.name),n_decimales,-1,0,-1))%>
																<%if calculoimporte=-1 and rst("Precio Medio2")<>0 then%>
																	<br/>
																	<%=EncodeForHtml(("<b>" & iif(tipo_medida>"",tipo_medida,"") & " : </b>" & formatnumber(rst("Precio Medio2"),n_decimales,-1,0,-1)))%>
																<%end if%>
															</td>
														<%end if
													else%>
														<td class='TDBORDECELDA7' align="right">
															<%if rst(campo.name)&""="" then%>
																&nbsp;
															<%else%>
																<%if campo.name<>"Cantidad" then%>
																	<%=EncodeForHtml(formatnumber(rst(campo.name),n_decimales,-1,0,-1))%>
																<%else%>
																	<%=EncodeForHtml(formatnumber(rst(campo.name),DEC_CANT,-1,0,-1))%>
																<%end if%>
															<%end if%>
														</td>
													<%end if%>
												<%end if%>
											<%else%>
												<%if campo.name<>"Cantidad" and campo.name<>"Cantidad2" and campo.name<>"calculoimporte" and campo.name<>"Precio Medio2" then%>
													<td class='TDBORDECELDA7' align="right">&nbsp;</td>
												<%end if%>
											<%end if%>
										<!--</td>-->
									<%end if
								else
									if campo.name<>"Descripcion" and campo.name<>"Ref"  and campo.name<>"Divisa" then
										if campo.name="NCliente" or ucase(campo.name)="NCLIENTE" then
											%><td class='tdbordecelda7'>
												<%if left(rst(campo.name),5)="-----" then%>
													<%=enc.EncodeForHtmlAttribute(null_s(rst("Nombre")))%>
												<%else%>
													<%=Hiperv(OBJClientes,rst(campo.name),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst(campo.name))),LitVerCliente)%>
												<%end if%>
											</td><%
										else
											if campo.name<>"AcumulaVentas" and campo.name<>"AcumulaCantidad" and campo.name<>"Orden" then
												if campo.name<>"cod_proyecto" or opc_cod_proyecto="1" then
													%><td class='tdbordecelda7'><%=EncodeForHtml(iif(rst(campo.name)&""="","&nbsp;",rst(campo.name)))%></td><%
												end if
											else
												%><td></td><%
											end if
										end if
									else
										if campo.name = "Divisa" and MostrarDivisa = true then
											%><td class='tdbordecelda7'><%=EncodeForHtml(d_lookup("abreviatura", "divisas", "codigo='" & rst(campo.name) & "' and codigo like '" & session("ncliente") & "%'", Session("backendlistados")))%></td><%
										else
											if campo.name = "Divisa" and MostrarDivisa = false then '
												%><td></td><%
											end if
										end if
									end if
								end if
							end if
						next
					CloseFila
					if ordenar=false then
					     SubTotal = rst("acumulaventas")
					     SubTotalCant = rst("acumulacantidad")
					else
					     SubTotal = rst("orden")
					     SubTotalCant = rst("acumulacantidad")
					end if
					rst.movenext
					fila = fila + 1
					''ricardo 22-3-2004
					valor_cantidad2=0
					tipo_medida=""
				wend
				if rst.eof then
					if (PorArticulo="SI") then
				      	DrawFila "" 'Fila de Subtotal
							%><td></td>
							<%if opc_cod_proyecto="1" then%>
								<td></td>
							<%end if%>
							<td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
							<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(SubTotalCant,DEC_CANT,-1,0,-1))%></b></td>
							<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(SubTotal,n_decimalesMB,-1,0,-1))%></b></td>
							<%pmTotal=0
							if SubTotalCant<>0 then pmTotal = cdbl(SubTotal)/SubTotalCant%>
							<td class='TDBORDECELDA7' align="right"><b><%=EncodeForHtml(formatnumber(pmTotal,n_decimalesMB,-1,0,-1))%></b></td>
						<%CloseFila
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
				      	<td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%><%=EncodeForHtml(iif(MostrarDivisa=true, " " & MB_abrev & ": ", ": "))%></b></td>
				      	<td></td>
					      <td class='TDBORDECELDA7' align="right"><big><b><%=EncodeForHtml(formatnumber(Suma,n_decimalesMB,-1,0,-1))%></b></big></td>
					<%CloseFila%>
					<%if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",Session("backendlistados")) then
						DrawFila "" 'Fila para el total equivalencia en PTS
						     %><td></td>
							<%if opc_cod_proyecto="1" then%>
								<td></td>
							<%end if%>
						     <td class='TDBORDECELDA7' bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",Session("backendlistados")))%>:</b></td>
						     <td></td>
						     <td class='TDBORDECELDA7' align="right"><big><b><%=EncodeForHtml(formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",Session("backendlistados")),-1,0,-1))%></b></big></td>
					     <%CloseFila%>
					<%end if%>
				<%end if%>
			</table>
			<br/>
			<%NavPaginas lote,lotes,campo,criterio,texto,2
		else%>
			<script language="javascript" type="text/javascript">
			    alert("<%=LitMsgDatosNoExiste%>");
			    parent.botones.document.location = "resumen_ventas_cli_bt.asp?mode=select1";
				parent.pantalla.resumen_ventas_cli.action="resumen_ventas_cli.asp?mode=select1";
				parent.pantalla.resumen_ventas_cli.submit();
				
			</script><%
		  end if 'end if rst.eof (seleccion)
		   rst.close
	   end if 'end if agrupa por articulo

''ricardo 25-5-2006 comienzo de la select
''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"fin_resumen_ventas"%>
    <iframe name="frameExportar" style='display:none;' src="resumen_ventas_cli_exportar.asp?mode=ver" frameborder='0' width='500' height='200'></iframe>
	<%end if%>
        </form>
        <%end if
        set rstAux = Nothing
        set rst = Nothing
        set rst2 = Nothing
        set rstSelect = Nothing
        set rstTablas = Nothing
        connRound.close
        set connRound = Nothing%>
    </body>
</html>