<%@ Language=VBScript %><% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
<%
 Server.ScriptTimeout = 3600
''ricardo 16-11-2007 se cambia la dsn desde dsncliente a backendlistados
%>
<%

''ricardo 1/4/2003 se ha incluido que salga como datos opcionales el coste medio
'' y que a la hora de calcular el coste de compra, se eliga que sea desde el pvd
'' o desde el coste medio de la tabla articulos

''ricardo 28-5-2003 se arregla el problema que da cuando se elige ver_coste_medio, pero no se eliga ver_coste,ver_proveedor,ver_pvd

'**RGU 24/1/2008: añadir campo opcional "Coste ficha Artículo" para ver en una columna el coste actual que tiene el artículo en su ficha.
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo2%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>

<link rel="stylesheet" href="../../pantalla.css" media="SCREEN"/>
<link rel="stylesheet" href="../../impresora.css" media="PRINT"/>
</head>
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
<!--#include file= "../../CatFamSubResponsive.inc"-->
<!--#include file="../../common/campospersoResponsive.inc" -->
<!--#include file="stock_almacenes.inc" -->
<!--#include file="../../modulos.inc" -->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->
<!--#include file= "../../styles/formularios.css.inc"-->  
<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript">


    function cambiar_smi() {
        while (document.stock_almacenes.stockmayoroigual.value.search(" ") != -1) {
            document.stock_almacenes.stockmayoroigual.value = document.stock_almacenes.stockmayoroigual.value.replace(" ", "");
        }
        //ricardo 15-11-2006 se podra dejar el stockmayoroigual a vacio

        if (isNaN(document.stock_almacenes.stockmayoroigual.value.replace(",", "."))) {
            window.alert("<%=LitStAlmNoNum%>");
            document.stock_almacenes.stockmayoroigual.value = "0";
            return;
        }
        document.stock_almacenes.stockmayoroigual.value = document.stock_almacenes.stockmayoroigual.value.replace(".", ",");
    }

    function CambioSinVentas() {
        if (document.stock_almacenes.elements["sinventa"].checked == true) parent.pantalla.document.getElementById("SinVentasEnabled").style.display = "";
        else parent.pantalla.document.getElementById("SinVentasEnabled").style.display = "none";
    }

    function traerproveedor() {
        nproveedor = document.stock_almacenes.nproveedor.value;
        document.stock_almacenes.action = "stock_almacenes.asp?mode=add&nproveedor=" + nproveedor;
        document.stock_almacenes.submit();
    }

    function sel_todos(modo) {
        parent.pantalla.document.stock_almacenes.elements["ver_familia"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_proveedores"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_dto"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_iva"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_pvp"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_pvd"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_coste"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_valor_mercado"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_stock"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_fecha_inventario"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_smin"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_reposicion"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_precibir"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_pservir"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_pmin"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_coste_medio"].checked = modo;
        parent.pantalla.document.stock_almacenes.elements["ver_coste_articulo"].checked = modo;
    }
</script>

<body onload="self.status='';" class="BODY_ASP">
<%

'*****************************************************************************'
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************'
'*****************************************************************************'
const borde=0

  set rstAux = Server.CreateObject("ADODB.Recordset")
  set rstAux2 = Server.CreateObject("ADODB.Recordset")
  set rstSelect = Server.CreateObject("ADODB.Recordset")
  set rst = Server.CreateObject("ADODB.Recordset")
  set conn=Server.CreateObject("ADODB.Connection")

sub CalculaRegistros (p_seleccion)
	NUMREGISTROS = 0

	rst.cursorlocation=3
	rst.open "select * from [" & session("usuario") & "]",session("backendlistados")

   show_almacen = ""
   old_referencia = ""
   old_referencia_or =""
   if rst.eof then RecordsetVacio=true
   while not rst.eof
	  '------------------------------'
	  if rst("almacen")<>show_almacen then
	     NUMREGISTROS = NUMREGISTROS + 1
		 old_referencia = rst("RefPantalla")
		 old_referencia_or = rst("referencia")
		 show_almacen = rst("almacen")
	  else
	     if rst("RefPantalla")<> old_referencia then
		    NUMREGISTROS = NUMREGISTROS + 1
			old_referencia = rst("RefPantalla")
			old_referencia_or = rst("referencia")
		 end if
	  end if
	  rst.movenext
	  '--------------------------'
   wend
rst.close
end sub

sub ImprimeSubTotal()
   numcols = 2
   if ver_coste="on" or ver_stock="on" or ver_valor_mercado="on" then
      DrawFila color_blau
         DrawCelda "TDSINBORDECELDA7 colspan='2'","","",0,"<b>"  & LitTotales  & "</b>"
      if ver_stock="on" then
	     tmpcol = POS_STOCK - numcols -1
		 if tmpcol>0 then
	        	DrawCelda "TDSINBORDECELDA7 colspan='" & tmpcol & "'","","",0,""
		 end if
		 numcols = numcols + tmpcol + 1
		 DrawCelda "DATO align='right'" ,"","",0,"<b>" & formatnumber(null_z(stock_almacen), DEC_CANT, -1, 0, -1) & "</b>"
		 TOTAL_STOCK = TOTAL_STOCK + stock_almacen
	  end if

	  ver_cols=0
	  ver_cols2=0
	  if ver_coste_medio="on" then
	    ver_cols=ver_cols+1
	  end if
	  if ver_coste_Articulo="on" then
	    ver_cols=ver_cols+1
	  end if

      if ver_coste="on" then
	     tmpcol = POS_TOTAL_COSTE - numcols - 1+ver_cols

		 if tmpcol>0 then
		    DrawCelda "TDSINBORDECELDA7 colspan='" & tmpcol & "'","","",0,""
		 end if
         DrawCelda "DATO align='right'","","",0,"<b>"  & formatnumber(null_z(coste_almacen), DEC_PREC, -1,0,-1)  & " " & ABMB & "</b>"
		 numcols = numcols + tmpcol '+ 1
		 TOTAL_COSTE = TOTAL_COSTE + coste_almacen
      end if
      ver_cols2=0
	  if ver_pvp="on" then
	    ver_cols2=ver_cols2+1
	  end if
      if ver_dto="on" then
        ver_cols2=ver_cols2+1
      end if
      if ver_iva="on" then
        ver_cols2=ver_cols2+1
      end if
	  if ver_valor_mercado="on" then

	     'tmpcol = POS_TOTAL_VALOR - numcols - 1+ver_cols2
	     tmpcol = ver_cols2
		 if tmpcol>0 then
		    DrawCelda "TDSINBORDECELDA7 colspan='" & tmpcol & "'","","",0,""
		 end if
         DrawCelda "DATO align='right'","","",0,"<b>"  & formatnumber(null_z(importe_almacen), DEC_PREC, -1,0,-1)  & " " & ABMB & "</b>"
		 numcols = numcols + tmpcol + 1
		 TOTAL_IMPORTE = TOTAL_IMPORTE + importe_almacen
      end if
   end if
	  CloseFila
	  DrawFila color_blau

	     DrawCelda "TDSINBORDECELDA7 colspan='" & COLUMNAS & "'","","",0,"&nbsp;"
	  CloseFila
	  DrawFila color_blau
	     DrawCelda "TDSINBORDECELDA7 colspan='" & COLUMNAS & "'","","",0,"&nbsp;"
	  CloseFila

   coste_almacen = 0
   importe_almacen = 0
   stock_almacen = 0
end sub

sub PasaSubTotal()
   if ver_coste="on" or ver_stock="on" or ver_valor_mercado="on" then
      if ver_stock="on" then
		 TOTAL_STOCK = TOTAL_STOCK + stock_almacen
	  end if
      if ver_coste="on" then
		 TOTAL_COSTE = TOTAL_COSTE + coste_almacen
      end if
	  if ver_valor_mercado="on" then
		 TOTAL_IMPORTE = TOTAL_IMPORTE + importe_almacen
      end if
   end if
   coste_almacen = 0
   importe_almacen = 0
   stock_almacen = 0
end sub

sub ImprimeLinea()
	'imprimimos los valores de old
	DrawCelda "DATO","","",0,Hiperv(OBJArticulos,old_referencia_or,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),old_referencia,LitVerArticulo)
	DrawCelda "DATO","","",0,Hiperv(OBJArticulos,old_referencia_or,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),old_nombre,LitVerArticulo)

	if ver_familia="on" then
		DrawCelda "DATO","","",0, old_familia
	end if
	if ver_stock="on" then
		DrawCelda "DATO align='right'","","",0, iif(old_espadre=true,"<b>"+old_stock+"</b>", old_stock)
		if old_espadre = false then  stock_almacen = stock_almacen + old_stock
	end if
	if ver_fecha_inventario="on" then
		DrawCelda "DATO align='right'","","",0, iif(old_espadre=true,"<b>"+old_fecha_inventario+"</b>", old_fecha_inventario)
	end if
	if ver_smin="on" then
		DrawCelda "DATO align='right'","","",0, iif(old_espadre=true,"<b>"+old_smin+"</b>", old_smin)
	end if
	if ver_reposicion="on" then
		DrawCelda "DATO align='right'","","",0,iif(old_espadre=true,"<b>"+old_reposicion+"</b>", old_reposicion)
	end if
	if ver_precibir="on" then
		DrawCelda "DATO align='right'","","",0,iif(old_espadre=true,"<b>"+old_precibir+"</b>", old_precibir)
	end if
	if ver_pservir="on" then
		DrawCelda "DATO align='right'","","",0,iif(old_espadre=true,"<b>"+old_pservir+"</b>", old_pservir)
	end if
	if ver_pmin="on" then
		DrawCelda "DATO align='right'","","",0,iif(old_espadre=true,"<b>"+old_pmin+"</b>", old_pmin)
	end if
	if ver_proveedores="on" then
		DrawCelda "DATO","","",0,enc.EncodeForHtmlAttribute(null_s(old_proveedores))
		'DrawCelda "DATO","","",0,Hiperv(OBJProveedores,old_proveedores_or,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),old_proveedores,LitVerProveedor)
	end if
	if ver_pvd="on" then
		DrawCelda "DATO align='right'","","",0,old_pvd
		redim old_numprov(1)
		old_numprov(1) = old_prov
		old_index = 1
	end if
	if ver_coste_medio="on" then
		DrawCelda "DATO align='right'","","",0,old_coste_medio
		redim old_numprov_cm(1)
		old_numprov_cm(1) = old_prov
		old_index_cm = 1
	end if
	if ver_coste_Articulo="on" then
		DrawCelda "DATO align='right'","","",0,old_coste_Art
	end if

	if ver_coste="on" then
		DrawCelda "DATO align='right'","","",0,old_coste
		if old_espadre=false then coste_almacen = coste_almacen + old_coste
	end if

	if ver_pvp="on" then
		DrawCelda "DATO align='right'","","",0,old_pvp
	end if

	if ver_dto="on" then
		DrawCelda "DATO align='right'","","",0,old_dto & "%"
	end if
	if ver_iva="on" then
		DrawCelda "DATO align='right'","","",0,old_iva & "%"
	end if
	if ver_valor_mercado="on" then
		DrawCelda "DATO align='right'","","",0,old_tventa
		if old_espadre=false then importe_almacen = importe_almacen + old_tventa
	end if
	CloseFila
	REGISTROS_IMPRESOS = REGISTROS_IMPRESOS + 1
end sub


sub PasaLinea()
	if ver_stock="on" then
		if old_espadre=false then stock_almacen = stock_almacen + old_stock
	end if
	if ver_pvd="on" then
		redim old_numprov(1)
		old_numprov(1) = old_prov
		old_index = 1
	end if
	if ver_coste_medio="on" then
		redim old_numprov_cm(1)
		old_numprov_cm(1) = old_prov
		old_index_cm = 1
	end if
	if ver_coste="on" then
		if old_espadre=false then coste_almacen = coste_almacen + old_coste
	end if
	if ver_valor_mercado="on" then
		if old_espadre = false then importe_almacen = importe_almacen + old_tventa
	end if
	registros_pasados = registros_pasados + 1
	if registros_pasados = MAXPAGINA then
		pagina_actual = pagina_actual + 1
		registros_pasados=0
	end if
end sub

function preparar_lista(valor_PLIST)
	dim aux_valor_PLIST
	aux_valor_PLIST=valor_PLIST
	if aux_valor_PLIST & "">"" then
		if left(aux_valor_PLIST,1)<>"(" then
			aux_valor_PLIST="(" & aux_valor_PLIST
		end if
		if mid(aux_valor_PLIST,2,1)<>"'" then
			aux_valor_PLIST=left(aux_valor_PLIST,1) & "'" & mid(aux_valor_PLIST,2,len(aux_valor_PLIST))
		end if
		if right(aux_valor_PLIST,1)<>")" then
			aux_valor_PLIST=aux_valor_PLIST & ")"
		end if
		if mid(aux_valor_PLIST,len(aux_valor_PLIST)-1,1)<>"'" then
			aux_valor_PLIST=mid(aux_valor_PLIST,1,len(aux_valor_PLIST)-1) & "'" & right(aux_valor_PLIST,1)
		end if
		aux_valor_PLIST=replace(aux_valor_PLIST,",","','")
		''se hara dos veces lo siguiente
		aux_valor_PLIST=replace(aux_valor_PLIST,"''","'")
		aux_valor_PLIST=replace(aux_valor_PLIST,"''","'")
	end if
	preparar_lista=aux_valor_PLIST
end function

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
%>
<form name="stock_almacenes" method="post">
    <%PintarCabecera "stock_almacenes.asp"

	WaitBoxOculto LitEsperePorFavor
    'Leer parámetros de la página'
	mode=enc.EncodeForJavascript(Request.QueryString("mode"))

    'nTiendita=Request.form("tiendas") ' jlvdc

    if mode="ver" then
	    %><script language="javascript" type="text/javascript">
	          //parent.pantallaList.cols = "0,0,*";
	          parent.document.getElementById("marcoset").cols = "0,0,*";
	          //parent.document.getElementsByName("pantallaList").cols = "100,100,100"; //"0,0,*";
	    </script><%
    end if

		COLUMNAS = 2
		TOTAL_COSTE = 0
		TOTAL_IMPORTE = 0
		TOTAL_STOCK = 0
		REGISTROS_IMPRESOS = 0
		NUMREGISTROS = 0
		registros_pasados=0

		coste_almacen = 0
		importe_almacen=0
		stock_almacen=0

		if trim(mode)="browse" then mode="ver"

		apaisado=iif(limpiaCadena(request.form("apaisado"))>"","SI","")

		if enc.EncodeForJavascript(request.querystring("nproveedor")) >"" then
		   nproveedor = limpiaCadena(request.querystring("nproveedor"))
		else
		   nproveedor = limpiaCadena(request.form("nproveedor"))
		end if

		if nproveedor & "">"" then
			nproveedor=session("ncliente") & completar(nproveedor,5,"0")
			if nproveedor & "">"" then
				nomproveedor=d_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("backendlistados"))
				if nomproveedor & ""="" then
					%><script language="javascript" type="text/javascript">
					      window.alert("<%=LitMsgProveedorNoExiste%>");
					</script><%
				end if
			end if
		end if

		if enc.EncodeForJavascript(request.querystring("opcproveedorbaja")) >"" then
		   opcproveedorbaja= limpiaCadena(request.querystring("opcproveedorbaja"))
		else
		   opcproveedorbaja= limpiaCadena(request.form("opcproveedorbaja"))
		end if

		if enc.EncodeForJavascript(request.querystring("referencia")) >"" then
		   referencia = limpiaCadena(request.querystring("referencia"))
		else
		   referencia = limpiaCadena(request.form("referencia"))
		end if
		if enc.EncodeForJavascript(request.form("nombre"))>"" then
			nombre = limpiaCadena(request.form("nombre"))
		else
			nombre=limpiaCadena(request.querystring("nombre"))
		end if

		if enc.EncodeForJavascript(request.form("familia"))>"" then
			familia = limpiaCadena(request.form("familia"))
		else
			familia = limpiaCadena(request.querystring("familia"))
		end if
		CheckCadena familia

		if enc.EncodeForJavascript(request.form("ordenar"))>"" then
			ordenar = limpiaCadena(request.form("ordenar"))
		else
			ordenar = limpiaCadena(request.querystring("ordenar"))
		end if

		if enc.EncodeForJavascript(request.form("almacen"))>"" then
			almacen = limpiaCadena(request.form("almacen"))
		else
			almacen = limpiaCadena(request.querystring("almacen"))
		end if
		CheckCadena almacen

		if enc.EncodeForJavascript(request.form("tipo_articulo"))>"" then
			tipo_articulo = limpiaCadena(request.form("tipo_articulo"))
		else
			tipo_articulo = limpiaCadena(request.querystring("tipo_articulo"))
		end if
		CheckCadena tipo_articulo
		if enc.EncodeForJavascript(request.form("sinventa"))&"">"" then
			CheckSinVenta="on"
			SinVentaDesde=limpiaCadena(request.form("sinventafdesde")&"")
			SinVentaHasta=limpiaCadena(request.form("sinventafhasta")&"")
		else
			if enc.EncodeForJavascript(request.querystring("sinventa")) &"" > "" then
				CheckSinVenta="on"
				SinVentaDesde=limpiaCadena(request.querystring("sinventafdesde")&"")
				SinVentaHasta=limpiaCadena(request.querystring("sinventafhasta")&"")
			else
				CheckSinVenta=""
				SinVentaDesde=""
				SinVentaHasta=""
			end if
		end if

		if enc.EncodeForJavascript(request.form("comocalccostcomp"))>"" then
			comocalccostcomp= limpiaCadena(request.form("comocalccostcomp"))
		else
			comocalccostcomp= limpiaCadena(request.querystring("comocalccostcomp"))
		end if

		if enc.EncodeForJavascript(request.form("stockmayoroigual"))>"" then
			stockmayoroigual = limpiaCadena(request.form("stockmayoroigual"))
		else
			stockmayoroigual = limpiaCadena(request.querystring("stockmayoroigual"))
		end if
		''ricardo 15/11/2006 se podra dejar el campo stockmayoroigual a vacio
		''if stockmayoroigual & ""="" then stockmayoroigual="0"

		if enc.EncodeForJavascript(request.form("ver_familia"))>"" then
			ver_familia = limpiaCadena(request.form("ver_familia"))
		else
			ver_familia = limpiaCadena(request.querystring("ver_familia"))
		end if
		if ver_familia>"" then COLUMNAS = COLUMNAS + 1
'----------------------------------------------------------------------'
		if enc.EncodeForJavascript(request.form("ver_stock"))>"" then
			ver_stock = limpiaCadena(request.form("ver_stock"))
		else
			ver_stock = limpiaCadena(request.querystring("ver_stock"))
		end if
		if ver_stock>"" then
		   COLUMNAS = COLUMNAS + 1
		   POS_STOCK = COLUMNAS
		end if
		if enc.EncodeForJavascript(request.form("ver_fecha_inventario"))>"" then
			ver_fecha_inventario = limpiaCadena(request.form("ver_fecha_inventario"))
		else
			ver_fecha_inventario = limpiaCadena(request.querystring("ver_fecha_inventario"))
		end if
		if ver_fecha_inventario>"" then
		   COLUMNAS = COLUMNAS + 1
		end if

		if enc.EncodeForJavascript(request.form("ver_smin"))>"" then
			ver_smin = limpiaCadena(request.form("ver_smin"))
		else
			ver_smin = limpiaCadena(request.querystring("ver_smin"))
		end if
		if ver_smin>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_reposicion"))>"" then
			ver_reposicion = limpiaCadena(request.form("ver_reposicion"))
		else
			ver_reposicion = limpiaCadena(request.querystring("ver_reposicion"))
		end if
		if ver_reposicion>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_precibir"))>"" then
			ver_precibir = limpiaCadena(request.form("ver_precibir"))
		else
			ver_precibir = limpiaCadena(request.querystring("ver_precibir"))
		end if
		if ver_precibir>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_pservir"))>"" then
			ver_pservir = limpiaCadena(request.form("ver_pservir"))
		else
			ver_pservir = limpiaCadena(request.querystring("ver_pservir"))
		end if
		if ver_pservir>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_pmin"))>"" then
			ver_pmin = limpiaCadena(request.form("ver_pmin"))
		else
			ver_pmin = limpiaCadena(request.querystring("ver_pmin"))
		end if
		if ver_pmin>"" then COLUMNAS = COLUMNAS + 1
'-------------------------------------------------------------------------'
		if enc.EncodeForJavascript(request.form("ver_proveedores"))>"" then
			ver_proveedores = limpiaCadena(request.form("ver_proveedores"))
		else
			ver_proveedores = limpiaCadena(request.querystring("ver_proveedores"))
		end if
		if ver_proveedores>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_pvd"))>"" then
			ver_pvd = limpiaCadena(request.form("ver_pvd"))
		else
			ver_pvd = limpiaCadena(request.querystring("ver_pvd"))
		end if
		if ver_pvd>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_coste"))>"" then
			ver_coste = limpiaCadena(request.form("ver_coste"))
		else
			ver_coste = limpiaCadena(request.querystring("ver_coste"))
		end if
		if ver_coste>"" then
		   COLUMNAS = COLUMNAS + 1
		   POS_TOTAL_COSTE = COLUMNAS
		end if

'--------------------------------------------------------------------------------'
		if enc.EncodeForJavascript(request.form("ver_pvp"))>"" then
			ver_pvp = limpiaCadena(request.form("ver_pvp"))
		else
			ver_pvp = limpiaCadena(request.querystring("ver_pvp"))
		end if
		if ver_pvp>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_dto"))>"" then
			ver_dto = limpiaCadena(request.form("ver_dto"))
		else
			ver_dto = limpiaCadena(request.querystring("ver_dto"))
		end if
		if ver_dto>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_iva"))>"" then
			ver_iva = limpiaCadena(request.form("ver_iva"))
		else
			ver_iva = limpiaCadena(request.querystring("ver_iva"))
		end if
		if ver_iva>"" then COLUMNAS = COLUMNAS + 1

		if enc.EncodeForJavascript(request.form("ver_coste_medio"))>"" then
			ver_coste_medio= limpiaCadena(request.form("ver_coste_medio"))
		else
			ver_coste_medio= limpiaCadena(request.querystring("ver_coste_medio"))
		end if
		if ver_coste_medio>"" then COLUMNAS = COLUMNAS + 1

 	   	if enc.EncodeForJavascript(request.form("ver_valor_mercado"))>"" then
			ver_valor_mercado = limpiaCadena(request.form("ver_valor_mercado"))
		else
			ver_valor_mercado = limpiaCadena(request.querystring("ver_valor_mercado"))
		end if
		if ver_valor_mercado>"" then
		   COLUMNAS = COLUMNAS + 1
		   POS_TOTAL_VALOR = COLUMNAS
		end if

		if enc.EncodeForJavascript(request.form("StockAFecha"))>"" then
			StockAFecha= limpiaCadena(request.form("StockAFecha"))
		else
			StockAFecha= limpiaCadena(request.querystring("StockAFecha"))
		end if


		if enc.EncodeForJavascript(request.form("ver_coste_articulo"))>"" then
			ver_coste_articulo= limpiaCadena(request.form("ver_coste_articulo"))
		else
			ver_coste_articulo= limpiaCadena(request.querystring("ver_coste_articulo"))
		end if
		if ver_coste_articulo>"" then COLUMNAS = COLUMNAS + 1

	dim au  ' cag Parametro que trae lista de almacenes para desplegable dependiendo del usuario

	ObtenerParametros("stockValorAlmacenes")
	au=preparar_lista(au)

	if mode="ver" then%>
		<table width='100%'>
   			<tr>
	  			<td width="30%" align="left">
		      	    	<font class='CABECERA'><b></b></font>
			          	<font class="CELDA7">&nbsp;<%="(" & LitEmitido & " " & day(date) & "/" & month(date) & "/" & year(date) & ")"%></font>
				</td>
			</tr>
	    </table>
		<hr/>
	<%else%>
		<br/>
	<%end if

	if mode="ver" then
		encabezado=0
		%><table><%
			if familia>"" then
				strselect ="select * from familias with(NOLOCK) where codigo='" & familia & "'"
                rst.cursorlocation=3
				rst.open strselect, session("backendlistados")
				if not rst.eof then
					encabezado=1
					DrawFila color_blau
						DrawCelda2 "CELDA", "left", false, LitSubFamilia + ": "
						DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("nombre")))                              
					CloseFila
				end if
				rst.close
			end if
			if almacen>"" then
				strselect ="select * from almacenes with(NOLOCK) where codigo='" & almacen & "'"
				rst.cursorlocation=3
				rst.open strselect, session("backendlistados")
				if not rst.eof then
					encabezado=1
					DrawFila color_blau
      				      DrawCelda2 "CELDA", "left", false, LitAlmacen + ": "
						DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))                   
					CloseFila
				end if
				rst.close
			end if
''ricardo 1/4/2003
			if comocalccostcomp>"" then
				encabezado=1
     	      		DrawCelda2 "CELDA", "left", false, LitCosteMedioComCalcCoste + ": "
     	        if comocalccostcomp="CosteArt" then
     	            DrawCelda2 "CELDA", "left", false, LitCalcCoste
     	        else
				    DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(comocalccostcomp))
				end if
			end if
'''''
			if referencia>"" then
				encabezado=1
				DrawFila color_blau
      	      		DrawCelda2 "CELDA", "left", false, LitConref + ": "
					DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(referencia))
				CloseFila
			end if
			if nombre>"" then
				encabezado=1
				DrawFila color_blau
					DrawCelda2 "CELDA", "left", false, LitConNombre + ": "
					DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(nombre))
				CloseFila
			end if
			if stockmayoroigual>"" and mode<>"add" then
				encabezado=1
				DrawFila color_blau
					DrawCelda2 "CELDA", "left", false, LitStockMayorOIgual + ": "
					DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(stockmayoroigual))
				CloseFila
			end if
			if tipo_articulo>"" and mode<>"add" then
				strselect ="select * from tipos_entidades with(nolock) where codigo='" & tipo_articulo & "' and tipo='ARTICULO'"
				rst.cursorlocation=3
				rst.open strselect, session("backendlistados")
				if not rst.eof then
					encabezado=1
					DrawFila color_blau
					DrawCelda2 "CELDA", "left", false, LitTipoArticulo + ": "
					DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))
					CloseFila
				end if
				rst.close
			end if
			if CheckSinVenta>"" and mode<>"add" then
				encabezado=1
				DrawFila color_blau
					DrawCelda2 "CELDA", "left", false, LitArticulosSinVentas + ": "
					DrawCelda2 "CELDA", "left", false, SinVentaDesde &" - "& SinVentaHasta
				CloseFila
			end if
			if nomproveedor & "">"" then
				encabezado=1
				DrawFila color_blau
					DrawCelda2 "CELDA", "left", false, LitProveedor + ": "
					DrawCelda2 "CELDA", "left", false, trimCodEmpresa(nproveedor) &" - "& enc.EncodeForHtmlAttribute(null_s(nomproveedor))
				CloseFila
			end if
			if StockAFecha & "">"" then
				encabezado=1
				DrawFila color_blau
					DrawCelda2 "CELDA", "left", false, LitStockAFecha + ": "
					DrawCelda2 "CELDA", "left", false, StockAFecha
				CloseFila
			end if
		%></table><%
		if encabezado=1 then
			%><hr/><%
		end if
    	end if

	Alarma "listado_articulo.asp"

  '*********************************************************************************************'
  'Se muestran parametros de seleccion'
  '*********************************************************************************************'

  if mode="add" then
        'DrawCelda2 "CELDA", "left", false, LitProveedor + ": "
        DrawDiv "1","",""
        DrawLabel "","",LitProveedor%><input class='width:15' type="text" name="nproveedor" value="<%=iif(nproveedor>"",trimCodEmpresa(nproveedor),"")%>" size="10"  onchange="traerproveedor()"/><a class='CELDAREFB' href="javascript:AbrirVentana('../../compras/proveedores_busqueda.asp?ndoc=stock_almacenes&titulo=<%=LitSelProveedor%>&mode=search&viene=stock_almacenes','P',<%=AltoVentana%>,<%=AnchoVentana%>)"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width:40' disabled type="text" readonly="readonly" name="razon_social" value="<%=enc.EncodeForHtmlAttribute(nomproveedor)%>" size="18"/><%CloseDiv
	    'DrawCelda2 "CELDA", "left", false, LitProveedorBaja +":"
		'DrawCheckCelda "CELDA","","",0,"","opcproveedorbaja",iif(opcproveedorbaja>"",-1,0)
        EligeCelda "check",mode,"left","","",0,LitProveedorBaja,"opcproveedorbaja",0,iif(opcproveedorbaja>"",-1,0)
        %>
   <hr/>
        <%
        'DrawCelda2 "CELDA width='20%'", "left", false, LitConref + ": "
   	    'DrawInputCelda "CELDA width='30%'","","",22,0,"","referencia",referencia
        EligeCelda "input",mode,"left","","",22,LitConref,"referencia",0,enc.EncodeForHtmlAttribute(null_s(referencia))
        'DrawCelda2 "CELDA width='20%'", "left", false, LitConNombre + ": "
	    'DrawInputCelda "CELDA width='30%'","","",25,0,"","nombre",nombre
        EligeCelda "input",mode,"left","","",22,LitConNombre,"nombre",0,enc.EncodeForHtmlAttribute(null_s(nombre))

        'DrawCelda2 "CELDA", "left", false, LitSubFamilia + ": "
        rstAux.cursorlocation=3
		rstAux.open " select codigo, nombre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre", session("backendlistados")
	    'DrawSelectCelda "CELDA","190","",0,"","familia",rstAux,familia,"codigo","nombre","",""
        DrawSelectCelda "CELDA","190","","",LitSubFamilia,"familia",rstAux,enc.EncodeForHtmlAttribute(null_s(familia)),"codigo","nombre","",""
		rstAux.close
       	
        'DrawCelda2 "CELDA style='width:140px'", "left", false, LITTIENDA + ": "			


        'DrawCelda2 "CELDA", "left", false, LitTipoArticulo + ": "
        rstAux.cursorlocation=3
		rstAux.open " select codigo, descripcion from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='ARTICULO' order by descripcion", session("backendlistados")
	    'DrawSelectCelda "CELDA","190","",0,"","tipo_articulo",rstAux,tipo_articulo,"codigo","descripcion","",""
        DrawSelectCelda "CELDA","190","","",LitTipoArticulo,"tipo_articulo",rstAux,enc.EncodeForHtmlAttribute(null_s(tipo_articulo)),"codigo","descripcion","",""
		rstAux.close            
        'DrawCelda2 "CELDA", "left", false, LitAlmacen + ": "
                %><!--<input value="prueba 10" />--><% 'lbardale

    if au>"" then
			    rstAux.cursorlocation=3
			    rstAux.open " select codigo, descripcion from almacenes with(nolock) where codigo in " & au & " order by descripcion", session("backendlistados")
                DrawDiv "1", "", ""
                DrawLabel "", "", LitAlmacen%><select class="width60" name="almacen"><%
                do while not rstAux.Eof %>
				    <option value="<%=rstAux("codigo")%>"> <%=enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion")))%></option>
				    <%rstAux.moveNext
				loop %>
                </select><%
    		    rstAux.close
                CloseDiv
	else
                rstAux2.cursorlocation=3
				rstAux2.open " select a.codigo, a.descripcion from configuracion c with(nolock), almacenes a with(nolock) where c.almacen=a.codigo and nempresa = '" & session("ncliente") & "'", session("backendlistados")
                'DGB: 30/01/2012  EVITAR ERRROR SI NO HAY ALMACENES
                if not rstAux2.Eof then
                    rstAux.cursorlocation=3
		   		    rstAux.open " select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%'" & " and codigo<>'" & rstAux2("codigo") & "' order by descripcion", session("backendlistados")
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitAlmacen%><select class='width60' name="almacen">
                    <option selected="selected" value="<%= enc.EncodeForHtmlAttribute(null_s(rstAux2("codigo"))) %>"><%=enc.EncodeForHtmlAttribute(null_s(rstAux2("descripcion")))%></option><%           
                    do while not rstAux.Eof %>
				        <option value="<%=rstAux("codigo")%>"> <%=enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) %></option>
				    <%rstAux.moveNext
				    loop%>
				    <option value="">&nbsp;</option>
			        </select><%
                        rstAux.close
                        CloseDiv
                else 'no hay almacen, se pinta vacío
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitAlmacen%><select class='width60' name="almacen">
                    <option value="">&nbsp;</option>
			        </select><%
                    CloseDiv
                end if
                rstAux2.Close
    end if
                    
 

	    'DrawCelda2 "CELDA", "left", false, LitOrdenar + ": "
        DrawDiv "1","",""
        DrawLabel "","",LitOrdenar%><select class='CELDA' name="ordenar" style='width:190px'>
			<option <%=iif(ordenar="Referencia" or ordenar="","selected","")%> value="Referencia"><%=LitRef%></option>
			<option <%=iif(ordenar="Nombre","selected","")%> value="Nombre"><%=LitNombre%></option>
			<option <%=iif(ordenar="SubFamilia","selected","")%> value="SubFamilia"><%=LitSubFamilia%></option>
		</select><%CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LitCosteMedioComCalcCoste%><select class="CELDA" name="comocalccostcomp" style='width:150px'>
				<option <%=iif(comocalccostcomp="Precio Proveedor" or comocalccostcomp="","selected","")%> value="Precio Proveedor"><%=LitCosteMedioPVD%></option>
				<option <%=iif(comocalccostcomp="Coste Medio","selected","")%> value="Coste Medio"><%=LITCOSTEMEDIOSTOCKALM%></option>
				<option  value="CosteArt"><%=LitCalcCoste%></option>
			</select><%CloseDiv
		
        'DrawCelda2 "CELDA", "left", false, LitApaisado & " : "
		'DrawCheckCelda "CELDA","","",0,"","apaisado",iif(apaisado>"",-1,0)
        EligeCelda "check",mode,"left","","",0,LitApaisado,"apaisado",0,iif(apaisado>"",-1,0)
        'DrawCelda2 "CELDA", "left", false, LitStockMayorOIgual + ": "
		'DrawInputCeldaAction "CELDA","","",22,0,"","stockmayoroigual",stockmayoroigual,"onchange","cambiar_smi()",false
        DrawInputCeldaActionDiv "","","","3",0,LitStockMayorOIgual,"stockmayoroigual",enc.EncodeForHtmlAttribute(null_s(stockmayoroigual)),"onchange","cambiar_smi()",false
		
	%><hr/><%
		'DrawCheckCelda "CELDA width='60%' height='30' onclick='CambioSinVentas()'","","",0,LitArticulosSinVentas&" : ","sinventa",""
        %><div class="col-xs-6"><%
        DrawDiv "col-xs-12", "",""
        DrawLabel "","",LitArticulosSinVentas
        DrawCheck "","CELDA width='60%' height='30' onclick='CambioSinVentas()'","sinventa",0
        CloseDiv%><span id="SinVentasEnabled" style="display: none"><%DrawDiv "col-xs-12","",""
        DrawLabel "","",LitDesdeFecha
        DrawInput "","left","sinventafdesde","",""
        DrawCalendar "sinventafdesde"
        CloseDiv
        'EligeCelda "input",mode,"left","","",10,LitDesdeFecha,"sinventafdesde",0,""
        DrawDiv "col-xs-12","",""
		DrawLabel "","",LitHastaFecha
        DrawInput "","left","sinventafhasta","",""
        DrawCalendar "sinventafhasta"
        CloseDiv
        'EligeCelda "input",mode,"left","","",10,LitHastaFecha,"sinventafhasta",0,""%></span><%
		DrawDiv "col-xs-12","",""
        DrawLabel "","",LitStockAFecha
        DrawInput "","left","StockAFecha","",""
        DrawCalendar "StockAFecha"
        CloseDiv
        DrawDiv "col-xs-12","",""
        DrawSpan "CELDA7","",LitStockAFechaAviso,""
        CloseDiv%></div><%
	%><hr/>
    <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitCamposOpcionales%></h6>
	<!--<table border='<%=borde%>' cellspacing="1" cellpadding="1"><%
			DrawFila color_fondo
				DrawCelda2 "ENCABEZADOC style='width:360px'", "left", false, LitCamposOpcionales
			CloseFila%>
		</table>--><%
		    'DrawCelda2 "CELDA", "left", false, LitSubFamilia
		    'DrawCheckCelda "CELDA","","",0,"","ver_familia",iif(ver_familia>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitSubfamilia,"ver_familia",0,iif(ver_familia>"",-1,0)
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false, LitProveedores
		    '''DrawCheckCelda "CELDA","","",0,"","ver_proveedores",cstr(ver_proveedores)
		    'DrawCheckCelda "CELDA","","",0,"","ver_proveedores",true
            EligeCelda "check",mode,"left","","",0,LitProveedores,"ver_proveedores",0,true
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false,LitCosteMedioStockAlm
		    'DrawCheckCelda "CELDA","","",0,"","ver_coste_medio",iif(ver_coste_medio>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitCosteMedioStockAlm,"ver_coste_medio",0,iif(ver_coste_medio>"",-1,0)
		    'DrawCelda2 "CELDA", "left", false, Litdto
		    'DrawCheckCelda "CELDA","","",0,"","ver_dto",iif(ver_dto>"",-1,0)
            EligeCelda "check",mode,"left","","",0,Litdto,"ver_dto",0,iif(ver_dto>"",-1,0)
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false, LitIVA
		    'DrawCheckCelda "CELDA","","",0,"","ver_iva",iif(ver_iva>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitIVA,"ver_iva",0,iif(ver_iva>"",-1,0)
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false,LitVerFechaInv
		    'DrawCheckCelda "CELDA","","",0,"","ver_fecha_inventario",iif(ver_fecha_inventario>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitVerFechaInv,"ver_fecha_inventario",0,iif(ver_fecha_inventario>"",-1,0)
		    'DrawCelda2 "CELDA", "left", false, LitPvp
		    'DrawCheckCelda "CELDA","","",0,"","ver_pvp",iif(ver_pvp>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitPvp,"ver_pvp",0,iif(ver_pvp>"",-1,0)
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false, LitValorMercado
		    'DrawCheckCelda "CELDA","","",0,"","ver_valor_mercado",iif(ver_valor_mercado>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitValorMercado,"ver_valor_mercado",0,iif(ver_valor_mercado>"",-1,0)
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false,LitVerCosteFichArt
		    'DrawCheckCelda "CELDA","","",0,"","ver_coste_articulo",iif(ver_coste_articulo>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitVerCosteFichArt,"ver_coste_articulo",0,iif(ver_coste_articulo>"",-1,0)
		    'DrawCelda2 "CELDA", "left", false, LitPvd2
		    'DrawCheckCelda "CELDA","","",0,"","ver_pvd",true
            EligeCelda "check",mode,"left","","",0,LitPvd2,"ver_pvd",0,true
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false, LitCoste
		    'DrawCheckCelda "CELDA","","",0,"","ver_coste",true
            EligeCelda "check",mode,"left","","",0,LitCoste,"ver_coste",0,true
		    'DrawCelda2 "CELDA", "left", false, LitStock
		    'DrawCheckCelda "CELDA","","",0,"","ver_stock",true
            EligeCelda "check",mode,"left","","",0,LitStock,"ver_stock",0,true
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false, LitSmin
 		    'DrawCheckCelda "CELDA","","",0,"","ver_smin",iif(ver_smin>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitSmin,"ver_smin",0,iif(ver_smin>"",-1,0)
		    'DrawCelda2 "CELDA", "left", false, LitReposicion
		    'DrawCheckCelda "CELDA","","",0,"","ver_reposicion",iif(ver_reposicion>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitReposicion,"ver_reposicion",0,iif(ver_reposicion>"",-1,0)
            'DrawCelda "CELDA","10%","",0," "
            'DrawCelda2 "CELDA", "left", false, LitPrecibir
		    'DrawCheckCelda "CELDA","","",0,"","ver_precibir",iif(ver_precibir>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitPrecibir,"ver_precibir",0,iif(ver_precibir>"",-1,0)
		    'DrawCelda2 "CELDA", "left", false, LitPservir
		    'DrawCheckCelda "CELDA","","",0,"","ver_pservir",iif(ver_pservir>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitPservir,"ver_pservir",0,iif(ver_pservir>"",-1,0)
		    'DrawCelda "CELDA","10%","",0," "
		    'DrawCelda2 "CELDA", "left", false, LitP_min
		    'DrawCheckCelda "CELDA","","",0,"","ver_pmin",iif(ver_pmin>"",-1,0)
            EligeCelda "check",mode,"left","","",0,LitP_min,"ver_pmin",0,iif(ver_pmin>"",-1,0)
       %>
       <hr/><table><%
		DrawFila color_blau
			%>
				<td class="CELDABOT" onclick="javascript:sel_todos(true);">
					<%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,LITBOTSELTODOTITLE%>
				</td>
				<td>&nbsp;</td>
				<td class="CELDABOT" onclick="javascript:sel_todos(false);">
					<%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,LITBOTDSELTODOTITLE%>
				</td>
			<%
		CloseFila
	%></table><%
   end if

    '*********************************************************************************************'
    ' Se muestran los datos de la consulta'
    '*********************************************************************************************'
    if mode="ver" or mode="edit" then
        ''ricardo 25-5-2006 comienzo de la select
        ''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
        auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"inicio_listado_stock"
        sentido=limpiaCadena(Request.QueryString("sentido"))
		MB=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("backendlistados"))
		ABMB=d_lookup("abreviatura","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("backendlistados"))

        strStockAlmacenes="EXEC ListadoStockAlmacenes @NomTabla='" & session("usuario") & "' ,@referencia='" & replace(referencia,"'","''") & "' ,@nombre='" & replace(nombre,"'","''") & "' ,@almacen='" & almacen & "' ,@tipo_articulo='" & tipo_articulo & "', @familia='" & familia & "', @nproveedor='" & nproveedor & "', @stockmayoroigual=" & iif(stockmayoroigual>"",replace(stockmayoroigual,",","."),"NULL") & ", @opcproveedorbaja=" & iif(opcproveedorbaja>"",1,0) & ", @checkSinVenta=" & iif(checkSinVenta>"",1,0) & ", @sinVentaDesde='" & SinVentaDesde & "', @SinVentaHasta='" & SinVentaHasta & "', @ordenar='" & UCASE(ordenar) & "', @session_ncliente='" & session("ncliente") & "',@StockAFecha='" & StockAFecha & "'"

		conn.open session("backendlistados")
		conn.CommandTimeout = 0
		set rstSelect = conn.execute(strStockAlmacenes)

	RecordSetVacio=false
	CalculaRegistros ""

	%><input type="hidden" name="referencia" value="<%=enc.EncodeForHtmlAttribute(referencia)%>"/>
 	<input type="hidden" name="nombre" value="<%=enc.EncodeForHtmlAttribute(nombre)%>"/>
	<input type="hidden" name="familia" value="<%=enc.EncodeForHtmlAttribute(familia)%>"/>
	<input type="hidden" name="almacen" value="<%=enc.EncodeForHtmlAttribute(almacen)%>"/>
	<input type="hidden" name="tipo_articulo" value="<%=enc.EncodeForHtmlAttribute(tipo_articulo)%>"/>
	<input type="hidden" name="comocalccostcomp" value="<%=enc.EncodeForHtmlAttribute(comocalccostcomp)%>"/>
	<input type="hidden" name="nproveedor" value="<%=enc.EncodeForHtmlAttribute(trimCodEmpresa(nproveedor))%>"/>
	<input type="hidden" name="stockmayoroigual" value="<%=enc.EncodeForHtmlAttribute(stockmayoroigual)%>"/>
	<input type="hidden" name="ordenar" value="<%=enc.EncodeForHtmlAttribute(ordenar)%>"/>
	<input type="hidden" name="ver_familia" value="<%=enc.EncodeForHtmlAttribute(ver_familia)%>"/>
	<input type="hidden" name="ver_proveedores" value="<%=enc.EncodeForHtmlAttribute(ver_proveedores)%>"/>
	<input type="hidden" name="ver_coste_medio" value="<%=enc.EncodeForHtmlAttribute(ver_coste_medio)%>"/>
	<input type="hidden" name="ver_dto" value="<%=enc.EncodeForHtmlAttribute(ver_dto)%>"/>
	<input type="hidden" name="ver_iva" value="<%=enc.EncodeForHtmlAttribute(ver_iva)%>"/>
	<input type="hidden" name="ver_pvp" value="<%=enc.EncodeForHtmlAttribute(ver_pvp)%>"/>
	<input type="hidden" name="ver_pvd" value="<%=enc.EncodeForHtmlAttribute(ver_pvd)%>"/>
	<input type="hidden" name="ver_coste" value="<%=enc.EncodeForHtmlAttribute(ver_coste)%>"/>
	<input type="hidden" name="ver_valor_mercado" value="<%=enc.EncodeForHtmlAttribute(ver_valor_mercado)%>"/>
	<input type="hidden" name="ver_stock" value="<%=enc.EncodeForHtmlAttribute(ver_stock)%>"/>
	<input type="hidden" name="ver_smin" value="<%=enc.EncodeForHtmlAttribute(ver_smin)%>"/>
	<input type="hidden" name="ver_reposicion" value="<%=enc.EncodeForHtmlAttribute(ver_reposicion)%>"/>
	<input type="hidden" name="ver_precibir" value="<%=enc.EncodeForHtmlAttribute(ver_precibir)%>"/>
	<input type="hidden" name="ver_pservir" value="<%=enc.EncodeForHtmlAttribute(ver_pservir)%>"/>
	<input type="hidden" name="ver_pmin" value="<%=enc.EncodeForHtmlAttribute(ver_pmin)%>"/>
	<input type="hidden" name="apaisado" value="<%=enc.EncodeForHtmlAttribute(apaisado)%>"/>
	<input type="hidden" name="sinventa" value="<%=enc.EncodeForHtmlAttribute(CheckSinVenta)%>"/>
	<input type="hidden" name="sinventafdesde" value="<%=enc.EncodeForHtmlAttribute(SinVentaDesde)%>"/>
	<input type="hidden" name="sinventafhasta" value="<%=enc.EncodeForHtmlAttribute(SinVentaHasta)%>"/>
	<input type="hidden" name="StockAFecha" value="<%=enc.EncodeForHtmlAttribute(StockAFecha)%>"/>
	<input type="hidden" name="ver_fecha_inventario" value="<%=enc.EncodeForHtmlAttribute(ver_fecha_inventario)%>"/>
	<input type="hidden" name="ver_coste_articulo" value="<%=enc.EncodeForHtmlAttribute(ver_coste_articulo)%>"/>

	<%
	MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='122'", DSNIlion)
	MAXPDF=d_lookup("maxpdf", "limites_listados", "item='122'", DSNIlion)
	%>
	<input type='hidden' name='maxpdf' value='<%=enc.EncodeForHtmlAttribute(MAXPDF)%>'/>
	<input type='hidden' name='maxpagina' value='<%=enc.EncodeForHtmlAttribute(MAXPAGINA)%>'/>
	<input type='hidden' name='totregs' value='<%=enc.EncodeForHtmlAttribute(NUMREGISTROS)%>'/>
	<%

	'COMPRUEBO SI EL RECORDSET QUE HE ABIERTO ANTES ES VACIO Y SI NO TRABAJO CON EL DESDE EL PRIMER REGISTRO'
	if RecordSetVacio then
		  %><font class='CEROFILAS'><%=LitCeroFilas%></font><%
	else
''ricardo 27-12-2005 se pone bien la ordenacion
cadena="select * from [" & session("usuario") & "] order by almacen"
select case ucase(ordenar)
	case "SUBFAMILIA":
		cadena=cadena & ",familia"
	case "NOMBRE":
		cadena=cadena & ",rtrim(ltrim(NomPantalla))"
	case "REFERENCIA":
		cadena=cadena & ",rtrim(ltrim(RefPantalla))"
end select
''''''''''''''''''''''''
''response.write("la cadena es-" & cadena & "-<br>")
        rst.cursorlocation=3
		rst.open cadena,session("backendlistados")

	if not rst.EOF then
		lote=limpiaCadena(Request.QueryString("lote"))

		if lote="" then
			lote=1
		end if
		sentido=limpiaCadena(Request.QueryString("sentido"))
		lotes=(NUMREGISTROS/MAXPAGINA)
		if lotes>(NUMREGISTROS/MAXPAGINA) then
			lotes=(NUMREGISTROS/MAXPAGINA)+1
		else
			lotes=(NUMREGISTROS/MAXPAGINA)
		end if

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

		NavPaginas lote,lotes,campo,criterio,texto,1

		almacenes = false
		if almacen>"" then
		   almacenes = true
		end if
		show_almacen = ""
		dim old_numprov()
		dim old_numprov_cm()
		dim old_numprov_ca()

		registros_pasados = 0
		pagina_actual = 1
		primero =true
		%><br/>

		<table width='100%' style="border-collapse: collapse;" cellspacing="1" cellpadding="1"> <%

		VinculosPagina(MostrarProveedores)=1:VinculosPagina(MostrarArticulos)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina


        impr_ultimalinea=1
''response.write("el inicio 1 es-" & REGISTROS_IMPRESOS & "-" & MAXPAGINA & "-<br>")
		while not rst.eof and REGISTROS_IMPRESOS<MAXPAGINA-1
''response.write("el inicio 2 es-" & REGISTROS_IMPRESOS & "-" & MAXPAGINA & "-<br>")
			'si hay varios almacenes ponemos un encabezado con el almacen'
			if rst("almacen")<>show_almacen then
				if old_referencia>"" then
					if pagina_actual = lote then
						ImprimeLinea()
						ImprimeSubTotal()
					else
						PasaLinea()
						PasaSubTotal()
					end if
				end if
				show_almacen = rst("almacen")
				if pagina_actual = lote then
					DrawFila color_fondo
						primero = false
						DrawCelda "DATO colspan='" & COLUMNAS & "'","","",0,"<b>" & LitAlmacen & ": " & rst("almdesc") & "</b>"
					CloseFila
					'ponemos encabezado'
					DrawFila color_fondo
						DrawCelda "DATO","","",0,"<b>" &LitRef
						DrawCelda "DATO","","",0,"<b>" & LitNombre & "</b>"
						if ver_familia="on" then
							DrawCelda "DATO","","",0,"<b>" & LitSubFamilia & "</b>"
						end if
						if ver_stock="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitStock & "</b>"
						end if
						if ver_fecha_inventario="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitVerFechaInv & "</b>"
						end if
						if ver_smin="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitSMin & "</b>"
						end if
						if ver_reposicion="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitReposicion & "</b>"
						end if
						if ver_precibir="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitPRecibir & "</b>"
						end if
						if ver_pservir="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitPServir & "</b>"
						end if
						if ver_pmin="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitP_min & "</b>"
						end if
						if ver_proveedores="on" then
							DrawCelda "DATO","","",0,"<b>" & LitProveedores & "</b>"
						end if
						if ver_pvd="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitPvd2 & "</b>"
						end if
						if ver_coste_medio="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitCosteMedioStockAlm & "</b>"
						end if
						if ver_coste_articulo="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitVerCosteFichArt & "</b>"
						end if

						if ver_coste="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitCoste & "</b>"
						end if

						if ver_pvp="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitPvp & "</b>"
						end if

						if ver_dto="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & Litdto & "</b>"
						end if
						if ver_iva="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitIva & "</b>"
						end if
						if ver_valor_mercado="on" then
							DrawCelda "DATO align='right'","","",0,"<b>" & LitValorMercado & "</b>"
						end if

					CloseFila
				end if
				'valores anteriores para artículos'
				old_referencia = ""
				old_referencia_or =""
				old_nombre = ""
				old_espadre = ""
				old_familia = ""
				old_stock = 0
				old_fecha_inventario=""
				old_smin = 0
				old_repeticion = 0
				old_precibir=0
				old_pservir = 0
				old_pmin=0
				old_proveedores = ""
				old_pvd = 0
				old_coste = 0
				old_coste_medio=0
				old_pvp=0
				old_Art=0
				old_dto=0
				old_iva = 0
				old_tventa = 0
				redim old_numprov(1)
				old_index = 1
				redim old_numprov_cm(1)
				old_index_cm = 1
				redim old_numprov_ca(1)
				old_index_ca = 1
			end if

			'controlamos cambios de articulos'
''response.write("diferencia-" & old_referencia & "-" & rst("RefPantalla") & "-<br>")
			if old_referencia<>rst("RefPantalla") then
				if old_referencia>"" then
					if pagina_actual = lote then
						if primero = true then
					      	primero = false
							DrawFila color_fondo
				   				primero = false
						            DrawCelda "DATO colspan='" & COLUMNAS & "'","","",0,"<b>" & LitAlmacen & ": " & rst("almdesc") & "</b>"
						      CloseFila
      				            'ponemos encabezado'
							DrawFila color_fondo
								DrawCelda "DATO","","",0,"<b>" &LitRef
								DrawCelda "DATO","","",0,"<b>" & LitNombre & "</b>"
								if ver_familia="on" then
									DrawCelda "DATO","","",0,"<b>" & LitSubFamilia & "</b>"
								end if
								if ver_stock="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitStock & "</b>"
								end if
								if ver_fecha_inventario="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitVerFechaInv & "</b>"
								end if
								if ver_smin="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitSMin & "</b>"
								end if
								if ver_reposicion="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitReposicion & "</b>"
								end if
								if ver_precibir="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitPRecibir & "</b>"
								end if
								if ver_pservir="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitPServir & "</b>"
								end if
								if ver_pmin="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitP_min & "</b>"
								end if
								if ver_proveedores="on" then
									DrawCelda "DATO","","",0,"<b>" & LitProveedores & "</b>"
								end if
								if ver_pvd="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitPvd2 & "</b>"
								end if
								if ver_coste_medio="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitCosteMedioStockAlm & "</b>"
								end if
								if ver_coste_Articulo="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitVerCosteFichArt & "</b>"
								end if
								if ver_coste="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitCoste & "</b>"
								end if
								if ver_pvp="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitPvp & "</b>"
								end if
								if ver_dto="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & Litdto & "</b>"
								end if
								if ver_iva="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitIva & "</b>"
								end if
								if ver_valor_mercado="on" then
									DrawCelda "DATO align='right'","","",0,"<b>" & LitValorMercado & "</b>"
								end if
							CloseFila
						end if
						ImprimeLinea()
					else
						PasaLinea()
					end if
				end if
				'guardamos los nuevos valores'
				old_referencia = rst("RefPantalla")
				old_referencia_or = rst("referencia")
				old_nombre = rst("nomPantalla")
				old_espadre = rst("es_padre")
				if ver_familia = "on" then
					old_familia = rst("nomfamilia")
				else
					old_familia = ""
				end if
				old_stock = formatnumber(null_z(rst("stock")), DEC_CANT, -1, 0, -1)
				old_fecha_inventario=rst("fecha_inventario")
				if ver_smin="on" then
					old_smin = rst("stock_minimo")
				else
					old_smin = 0
				end if
				old_smin = formatnumber(null_z(old_smin), DEC_CANT, -1,0,-1)
				if ver_reposicion="on" then
					old_reposicion = rst("reposicion")
				else
		            	old_reposicion = 0
				end if
				old_reposicion = formatnumber(null_z(old_reposicion), DEC_CANT, -1, 0, -1)
				if ver_precibir="on" then
					old_precibir = rst("p_recibir")
				else
		            	old_precibir=0
				end if
				old_precibir = formatnumber(null_z(old_precibir), DEC_CANT, -1, 0, -1)
				if ver_pservir="on" then
					old_pservir = rst("p_servir")
				else
		            	old_pservir = 0
				end if
				old_pservir = formatnumber(null_z(old_pservir), DEC_CANT, -1, 0, -1)
				if ver_min="on" then
					old_pmin = rst("p_min")
				else
		            	old_pmin=0
				end if
				old_pmin = formatnumber(old_pmin, DEC_CANT, -1, 0, -1)
				if ver_proveedores = "on" then
					old_proveedores = rst("nomprov")
					old_proveedores_or=rst("nprov")
				else
		            	old_proveedores = ""
				end if
				if (ver_pvd = "on" or ver_coste = "on") then ''''and comocalccostcomp="Precio Proveedor" then
					old_pvd = CambioDivisa(rst("pvd"), rst("prdivisa"), MB)
					old_pvd = formatnumber(null_z(old_pvd),DEC_PREC, -1, 0, -1)
				end if

				'**RGU 2/10/2007: se añade la posibilidad de sacar el listado por coste del artículo
				if ver_coste ="on" then
				    old_coste_Art=CambioDivisa(rst("ImporteArt"), rst("prdivisa"), MB)
					old_coste_Art = formatnumber(null_z(old_coste_Art),DEC_PREC, -1, 0, -1)
				end if

				'**rgu**

				redim old_numprov(1)
				old_numprov(1) = old_pvd
				old_index = 1
				if (ver_coste_medio="on" or ver_coste = "on") then '''and comocalccostcomp="Coste Medio" then
					old_coste_medio = CambioDivisa(rst("coste_medio"), rst("prdivisa"), MB)
					old_coste_medio = formatnumber(null_z(old_coste_medio),DEC_PREC, -1, 0, -1)
				end if

				redim old_numprov_cm(1)
				old_numprov_cm(1) = old_coste_medio
				old_index_cm = 1

				redim old_numprov_ca(1)
				old_numprov_ca(1) = old_coste_Art
				old_index_ca = 1

				if ver_coste = "on" then
					if comocalccostcomp="Precio Proveedor" then
						old_coste = miround(rst("stock"), DEC_CANT) * round(old_pvd, DEC_PREC)
					elseif comocalccostcomp="Coste Medio" then
					    old_coste=miround(rst("stock"), DEC_CANT) * round(old_coste_medio, DEC_PREC)
				    else
				        old_coste=miround(rst("stock"), DEC_CANT) * round(old_coste_Art, DEC_PREC)
					end if
					old_coste = formatnumber(null_z(old_coste),DEC_PREC, -1, 0, -1)
				else
	            	old_coste = 0
				end if

           		if ver_pvp ="on" or ver_valor_mercado ="on" then
					old_pvp = CambioDivisa(rst("pvp"), rst("divisa"), MB)
					old_pvp = formatnumber(null_z(old_pvp),DEC_PREC, -1, 0, -1)
				else
         		    old_pvp=0
				end if

				if ver_dto="on" then
					old_dto = rst("dto")
				else
 		            	old_dto=0
				end if
				if ver_iva = "on" then
					old_iva = rst("iva")
				else
 		            old_iva = 0
				end if
				if ver_valor_mercado = "on" then
					old_tventa = old_stock * old_pvp
					old_tventa = formatnumber(null_z(old_tventa),DEC_PREC, -1, 0, -1)
				else
		            	old_tventa = 0
				end if
			 	 'end if 'if referencia>""
			else
''REGISTROS_IMPRESOS=REGISTROS_IMPRESOS+1
''response.write("en proveedores he entrado<br>")
				'actualizamos los valores'
                if ver_proveedores = "on" then
					if old_proveedores >"" then old_proveedores = old_proveedores & ", "
					old_proveedores = old_proveedores & rst("nomprov")
				end if
				if ver_pvd = "on" or ver_coste ="on" then
					old_index = old_index + 1
					old_pvd = CambioDivisa(rst("pvd"), rst("prdivisa"), MB)
					redim preserve old_numprov(old_index)
					old_numprov(old_index)=old_pvd
					suma = 0
					for i=0 to old_index
						suma = suma + old_numprov(i)
					next
					old_pvd = suma/old_index
					old_pvd = formatnumber(null_z(old_pvd),DEC_PREC, -1, 0, -1)
				end if

				if ver_coste_medio = "on" or ver_coste ="on" then
					old_index_cm = old_index_cm + 1
					old_coste_medio = CambioDivisa(rst("coste_medio"), rst("prdivisa"), MB)
					redim preserve old_numprov_cm(old_index_cm)
					old_numprov_cm(old_index_cm)=old_coste_medio
					suma = 0
					for i=0 to old_index_cm
						suma = suma + old_numprov_cm(i)
					next
					old_coste_medio = suma/old_index_cm
					old_coste_medio = formatnumber(null_z(old_coste_medio),DEC_PREC, -1, 0, -1)
				end if

				if ver_coste = "on" then
					if comocalccostcomp="Precio Proveedor" then
						old_coste = miround(rst("stock"), DEC_CANT) * round(old_pvd, DEC_PREC)
					elseif comocalccostcomp="Coste Medio" then
						    old_coste = miround(rst("stock"), DEC_CANT) * round(old_coste_medio, DEC_PREC)
						else
					        old_index_ca = old_index_ca + 1
					        old_coste_Art = CambioDivisa(rst("ImporteArt"), rst("prdivisa"), MB)
					        redim preserve old_numprov_ca(old_index_ca)
					        old_numprov_ca(old_index_ca)=old_coste_Art
					        suma = 0
					        for i=0 to old_index_ca
						        suma = suma + old_numprov_ca(i)
					        next
					        old_coste_Art = suma/old_index_ca
					        old_coste_Art = formatnumber(null_z(old_coste_Art),DEC_PREC, -1, 0, -1)

						    old_coste = miround(rst("stock"), DEC_CANT) * round(old_coste_Art, DEC_PREC)
					end if
					old_coste = formatnumber(null_z(old_coste),DEC_PREC, -1, 0, -1)
				end if
			end if
tmprefOld=rst("refpantalla")&""
			rst.movenext
''response.write("los datos 1 son-" & tmprefOld & "-" & old_proveedores & "-" & pagina_actual & "-" & lote & "-" & lotes & "-" & impr_ultimalinea & "-" & (rst.eof) & "-" & REGISTROS_IMPRESOS & "-" & MAXPAGINA & "-<br>")
			
''ricardo 17-12-2012 se arregla la obtencion de los proveedores de un articulo
esto_esta_mal=0
if esto_esta_mal=1 then
            '**RGU 13/2/2008: Si la última linea tenía mas de un proveedor se estaban calculando mal los datos
			if REGISTROS_IMPRESOS>=MAXPAGINA-1 then
		        if not rst.EOF then
                    tmpref=rst("refpantalla")
                    rst.MoveNext
                    if not rst.EOF then
		                if tmpref = rst("refpantalla") then
		                    REGISTROS_IMPRESOS=REGISTROS_IMPRESOS-1
		                    impr_ultimalinea=0
		                end if
                    end if
		            if not rst.BOF then
		                rst.MovePrevious
		            end if
		        end if
            end if
else
    if REGISTROS_IMPRESOS>=MAXPAGINA-1 then
        if not rst.eof then
            refActual=rst("refpantalla")&""
        else
            refActual=""
        end if
''response.write("los datos 1.1 entrado-" & tmprefOld & "-" & refActual & "-" & REGISTROS_IMPRESOS & "-<br>")
        ''if not rst.eof and tmprefOld=refActual then
        ''    ''REGISTROS_IMPRESOS=REGISTROS_IMPRESOS-1
''response.write("los datos 1.3 entrado-" & tmprefOld & "-" & refActual & "-" & REGISTROS_IMPRESOS & "-<br>")
        ''end if

            while not rst.eof and tmprefOld=refActual
				if old_proveedores >"" then old_proveedores = old_proveedores & ", "
                if ver_proveedores = "on" then
					old_proveedores = old_proveedores & rst("nomprov")
				end if
				if ver_pvd = "on" or ver_coste ="on" then
					old_index = old_index + 1
					old_pvd = CambioDivisa(rst("pvd"), rst("prdivisa"), MB)
					redim preserve old_numprov(old_index)
					old_numprov(old_index)=old_pvd
					suma = 0
					for i=0 to old_index
						suma = suma + old_numprov(i)
					next
					old_pvd = suma/old_index
					old_pvd = formatnumber(null_z(old_pvd),DEC_PREC, -1, 0, -1)
				end if

				if ver_coste_medio = "on" or ver_coste ="on" then
					old_index_cm = old_index_cm + 1
					old_coste_medio = CambioDivisa(rst("coste_medio"), rst("prdivisa"), MB)
					redim preserve old_numprov_cm(old_index_cm)
					old_numprov_cm(old_index_cm)=old_coste_medio
					suma = 0
					for i=0 to old_index_cm
						suma = suma + old_numprov_cm(i)
					next
					old_coste_medio = suma/old_index_cm
					old_coste_medio = formatnumber(null_z(old_coste_medio),DEC_PREC, -1, 0, -1)
				end if

				if ver_coste = "on" then
					if comocalccostcomp="Precio Proveedor" then
						old_coste = miround(rst("stock"), DEC_CANT) * round(old_pvd, DEC_PREC)
					elseif comocalccostcomp="Coste Medio" then
						    old_coste = miround(rst("stock"), DEC_CANT) * round(old_coste_medio, DEC_PREC)
						else
					        old_index_ca = old_index_ca + 1
					        old_coste_Art = CambioDivisa(rst("ImporteArt"), rst("prdivisa"), MB)
					        redim preserve old_numprov_ca(old_index_ca)
					        old_numprov_ca(old_index_ca)=old_coste_Art
					        suma = 0
					        for i=0 to old_index_ca
						        suma = suma + old_numprov_ca(i)
					        next
					        old_coste_Art = suma/old_index_ca
					        old_coste_Art = formatnumber(null_z(old_coste_Art),DEC_PREC, -1, 0, -1)

						    old_coste = miround(rst("stock"), DEC_CANT) * round(old_coste_Art, DEC_PREC)
					end if
					old_coste = formatnumber(null_z(old_coste),DEC_PREC, -1, 0, -1)
				end if
                rst.movenext
                if not rst.eof then
                    refActual=rst("refpantalla")&""
                else
                    refActual=""
                end if
            wend
    end if
end if
		wend
''response.write("los datos 2 son-" & pagina_actual & "-" & lote & "-" & lotes & "-" & impr_ultimalinea & "-" & (rst.eof) & "-" & REGISTROS_IMPRESOS & "-" & MAXPAGINA & "-<br>")
		if pagina_actual = lote   and impr_ultimalinea=1 then
    		ImprimeLinea()
		else
			PasaLinea()
		end if

		'si hemos llegado al final'
		if rst.eof then
			ImprimeSubTotal()
			'Imprimimos totales'
			if ver_stock="on" then
				DrawFila color_fondo
					DrawCelda "TDSINBORDECELDA7 colspan='" & COLUMNAS & "'","","",0,"<b>" & LitTotalStock & ": "& formatnumber(null_z(TOTAL_STOCK), DEC_CANT, -1, 0, -1) &  "</b>"
				CloseFila
			end if
			if ver_coste="on" then
				DrawFila color_fondo
					DrawCelda "TDSINBORDECELDA7 colspan='" & COLUMNAS & "'","","",0,"<b>" & LitTotalCoste & ": " & formatnumber(null_z(TOTAL_COSTE), DEC_PREC, -1, 0, -1) & " " & ABMB & "</b>"
				CloseFila
			end if
		      if ver_valor_mercado="on" then
		      	DrawFila color_fondo
					DrawCelda "TDSINBORDECELDA7 colspan='" & COLUMNAS & "'","","",0,"<b>" & LitTotalVenta & ": " & formatnumber(null_z(TOTAL_IMPORTE), DEC_PREC, -1, 0, -1) & " " & ABMB & "</b>"
				CloseFila
			end if
		end if

		rst.close%>
	</table>
	<br/><%

      NavPaginas lote,lotes,campo,criterio,texto,2
	else
		  %><font class='CEROFILAS'><%=LitCeroFilas%></font><%
      end if
end if

''ricardo 25-5-2006 comienzo de la select
''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"fin_listado_stock"
end if%>
<iframe name="frameExportar" style='display:none;' src="stock_almacenes_pdf.asp?mode=ver" frameborder='0' width='500' height='200'></iframe>
</form>
<%end if
connRound.close
set connRound = Nothing
set rstAux = Nothing
set rstAux2 = Nothing
set rstSelect = Nothing
set rst = Nothing
set conn = Nothing
%>
</body>
</html>