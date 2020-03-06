<%@ Language=VBScript %>
<%
'**RGU 7/12/2006 : Se añaden a la tabla temporal los campos cantidadarticulo, unidadarticulo para que aparezcan en el listado18 en el literal "El Kg le sale a..."
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file= "../CatFamSubResponsive.inc"-->
<!--#include file="codigo_barras.inc" -->
<!--#include file="../styles/formularios.css.inc" -->  
<!--#include file ="../common/campospersoResponsive.inc"-->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc"-->
<title><%=LitTitulo2%></title> 

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>

<link rel="stylesheet" href="../pantalla.css" media="SCREEN"/>
<link rel="stylesheet" href="../impresora.css" media="PRINT"/>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function%>  

</head>
<%
    si_tiene_modulo_Centroxogo=ModuloContratado(session("ncliente"),ModCentroxogo)
%>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
function ValidarCampos() {
	ok=1;

	cantHMax=0
	cantVMax=0
	maxpagina=0
////Se pone esto, ya que si se introduce uno o varios espacios en blanco el isNaN no lo detecta
	while (document.codigo_barras.cantidad.value.search(" ")!=-1){
	    document.codigo_barras.cantidad.value=document.codigo_barras.cantidad.value.replace(" ","");
	}
	if (ok==1 && document.codigo_barras.cantidad.value=='' && document.codigo_barras.cant_doc.checked==false) {
        	window.alert("<%=LitMsgCantidadNoNulo%>");
		ok=0;
	}
	if(ok==1 && isNaN(document.codigo_barras.cantidad.value)){
		window.alert("<%=LitMsgCantidadNoCaracter%>");
		ok=0;
		}
////Se pone esto, ya que si se introduce uno o varios espacios en blanco el isNaN no lo detecta
	while (document.codigo_barras.imprimir_listado_horizontal.value.search(" ")!=-1){
	    document.codigo_barras.imprimir_listado_horizontal.value=document.codigo_barras.imprimir_listado_horizontal.value.replace(" ","");
	}
	if(ok==1 && isNaN(document.codigo_barras.imprimir_listado_horizontal.value)){
		window.alert("<%=LitMsgHorizontalNoCaracter%>");
		ok=0;
	}
////Se pone esto, ya que si se introduce uno o varios espacios en blanco el isNaN no lo detecta
	while (document.codigo_barras.imprimir_listado_vertical.value.search(" ")!=-1){
	    document.codigo_barras.imprimir_listado_vertical.value=document.codigo_barras.imprimir_listado_vertical.value.replace(" ","");
	}
	if(ok==1 && isNaN(document.codigo_barras.imprimir_listado_vertical.value)){
		window.alert("<%=LitMsgVerticalNoCaracter%>");
		ok=0;
	}
	if(ok==1 && document.codigo_barras.imprimir_listado_horizontal.value==''){
		window.alert("<%=LitMsgHorizontalNoNulo%>");
		ok=0;
	}
	if(ok==1 && document.codigo_barras.imprimir_listado_vertical.value==''){
		window.alert("<%=LitMsgVerticalNoNulo%>");
		ok=0;
	}
	if (ok==1 && parseInt(document.codigo_barras.imprimir_listado_vertical.value)<1 || parseInt(document.codigo_barras.imprimir_listado_horizontal.value)<1) {
        window.alert("<%=LitMsgHORVERMINPAGNoNulo%>");
		ok=0;
	}
	if(ok==1 && document.codigo_barras.numdoc.value!='' && document.codigo_barras.tipodoc.value==''){
		window.alert("<%=LitnumdoctipodocNulo%>");
		ok=0;
		}
	if(ok==1 && document.codigo_barras.numdoc.value=='' && document.codigo_barras.tipodoc.value!='' && document.codigo_barras.tipodoc.value!='ASIGNACION MASIVA'){
		window.alert("<%=LitnumdocnumdocNulo%>");
		ok=0;
	}

	if(ok==1 && document.codigo_barras.fmpc.value!=""){
	    if (!checkdate(document.codigo_barras.fmpc)){
			window.alert("<%=LitAMPFPFechMal%>");
			ok=0;
		}
	}

	if(ok==1 && document.codigo_barras.tarifaex.value==''){
		window.alert("<%=LitTarifaNoNula%>");
		ok=0;
	}

	if(ok==1 && document.codigo_barras.stockmayoroigual.value!=''){
	    if (isNaN(document.codigo_barras.stockmayoroigual.value.replace(",","."))){
			window.alert("<%=LitMsgStockNumerico%>");
			ok=0;
		}
	}

	if (ok==1) return true;
	else return false;
}

function WinArticulos() {
	Ven=AbrirVentana("../productos/articulos_buscar.asp?ndoc=codigo_barras&titulo=<%=LitSelArticulo%>&mode=search","P",<%=AltoVentana%>,<%=AnchoVentana%>);
}

//Desencadena la búsqueda del artículo cuya referencia se indica
function TraerArticulo(mode,ndet) {
    if (document.codigo_barras.referencia.value!="") {
        //document.codigo_barras.refrescar.value="NO";
        //document.location.href="codigo_barras.asp?ndoc=" + document.albaranes_clidet.nalbaran.value + "&ncliente=" + document.albaranes_clidet.ncliente.value + "&mode=" + mode +"&fye=" + fye + "&ref=" + document.albaranes_clidet.referencia.value + "&cant=" + document.albaranes_clidet.cantidad.value + "&ndet=" + ndet;
	}
}

var apretado_boton_derecho=0;

function mostrar_cantidad(modo){
    if (modo==0){
        document.codigo_barras.numdoc.readOnly = "";
        document.codigo_barras.numdoc.value = "";
        if (document.codigo_barras.tipodoc.value=='ASIGNACION MASIVA') {
            tarifa=document.codigo_barras.tarifa.value;
            noref=document.codigo_barras.noref.value;
		    //dgb  08/04/2008 se anyade un parametro para indicar desde que pagina se llama
		    Ven=AbrirVentana("./listaasignacionmasiva.asp?tarifa=" + tarifa+"&pag=nuevo&noref="+noref,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
	    }
    }

    if (document.codigo_barras.numdoc.value!="" || document.codigo_barras.tipodoc.value!=""){
        document.codigo_barras.cant_doc.checked=true;
        document.codigo_barras.cantidad.value="";
        document.codigo_barras.cantidad.disabled=true;
		document.getElementById("idcant_doc2").style.display="";
		document.getElementById("id_fmpc").style.display="none";

		//ricardo 21-2-2007 se añade el stock
		document.getElementById("stock1").style.display="none";
		document.getElementById("stock2").style.display="none";
		document.getElementById("stock3").style.display="none";
		document.codigo_barras.stockmayoroigual.value="";
		document.codigo_barras.almacen.value="";
		//document.codigo_barras.tienda.value="";
	}
	else{
        document.codigo_barras.cantidad.value="1";
        document.codigo_barras.cantidad.disabled=false;
        document.codigo_barras.cant_doc.checked=false;
		document.getElementById("idcant_doc2").style.display="none";
		if (document.codigo_barras.cod_temporada.value == "") document.getElementById("id_fmpc").style.display="";
		else document.getElementById("id_fmpc").style.display="none";

		//ricardo 21-2-2007 se añade el stock
		document.getElementById("stock1").style.display="";
		document.getElementById("stock2").style.display="";
		document.getElementById("stock3").style.display="";

	}
}

function control_cantidad(){
    if (document.codigo_barras.cant_doc.checked==true){
        document.codigo_barras.cantidad.value="";
        document.codigo_barras.cantidad.disabled=true;
	}
	else{
        document.codigo_barras.cantidad.value="1";
        document.codigo_barras.cantidad.disabled=false;
	}
}

function Cambio(){
   
    if (document.codigo_barras.formato_impresion.value=="listado_codigo_barras4.asp" ||
        document.codigo_barras.formato_impresion.value=="listado_codigo_barrasSolred3x7.asp"){
        document.codigo_barras.ver_referencia.checked=false;
        document.codigo_barras.ver_referencia.disabled=true;       
        document.codigo_barras.ver_empresa.checked=false;
        document.codigo_barras.ver_empresa.disabled=true;
        if (document.codigo_barras.si_tiene_modulo_terminales.value==1){
            document.codigo_barras.ver_codTerminal.checked=false;
            document.codigo_barras.ver_codTerminal.disabled=true;
		}
		document.getElementById("IMPORTESADICIONALES").style.display="none";
	}
	else{
        if (document.codigo_barras.ver_referencia.disabled==true){
            document.codigo_barras.ver_referencia.disabled=false;
            document.codigo_barras.ver_referencia.checked=true;
		}
        if (document.codigo_barras.ver_refProv.disabled==true){
            document.codigo_barras.ver_refProv.disabled=false;
            document.codigo_barras.ver_refProv.checked=true;
		}
        if (document.codigo_barras.ver_empresa.disabled==true){
            document.codigo_barras.ver_empresa.disabled=false;
            document.codigo_barras.ver_empresa.checked=true;
		}
        if (document.codigo_barras.si_tiene_modulo_terminales.value==1){
            if (document.codigo_barras.ver_codTerminal.disabled==true) document.codigo_barras.ver_codTerminal.disabled=false;
		}
        if (document.codigo_barras.formato_impresion.value=="../custom/listado_codigo_barrasn.asp"){
            if (document.codigo_barras.cod_temporada.value == "") document.getElementById("IMPORTESADICIONALES").style.display="";
		}
		else document.getElementById("IMPORTESADICIONALES").style.display="none";
	}
}

function MuestraFechaMod(obj) {
	if (obj.checked) {
	    document.codigo_barras.fmpc.value="";
	    document.codigo_barras.fmpc.disabled=true;
	}
	else {
	    document.codigo_barras.fmpc.disabled=true;
	    document.codigo_barras.fmpc.value="";
	    document.codigo_barras.fmpc.disabled=false;
	}
}

function MuestraFechaMod3(obj)
{
    if (obj.value!="") document.codigo_barras.solopreciocambiado.checked="";
    else document.codigo_barras.solopreciocambiado.disabled=false;
}

function MuestraFechaMod2(obj) {
    act=0;
    for (i=0;i<obj.options.length;i++)
    {
	    if((obj.options[i].selected) && (obj.options[i].value)) act=1;
    }
	if (act==1) {
	    document.codigo_barras.fmpc.value="";
	    document.codigo_barras.fmpc.disabled=true;
	}
	else {
	    document.codigo_barras.fmpc.value="";
	    document.codigo_barras.fmpc.disabled=false;
	}
}

function control_temporada()
{
    if (document.codigo_barras.cod_temporada.value != "")
	{
		document.getElementById("IMPORTESADICIONALES").style.display="none";
		document.getElementById("id_fmpc").style.display="none";
		document.getElementById("id_fin_temp").style.display="";
	}
	else
	{
		document.getElementById("id_fin_temp").style.display="none";
		document.codigo_barras.fin_temp.checked=false;
		mostrar_cantidad();
		if (document.codigo_barras.formato_impresion.value=="../custom/listado_codigo_barrasn.asp")
			document.getElementById("IMPORTESADICIONALES").style.display="";
		else document.getElementById("IMPORTESADICIONALES").style.display="none";
	}
}
</script>
<body onload="self.status='';" class="BODY_ASP">
<%
''*****************************************************************************
''********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
''*****************************************************************************
const borde=0

%>
<form name="codigo_barras" method="post"><%
		PintarCabecera "codigo_barras_nuevo.asp"

		si_tiene_modulo_terminales=ModuloContratado(session("ncliente"),ModTerminales)

		set rstAux = Server.CreateObject("ADODB.Recordset")
		set rst = Server.CreateObject("ADODB.Recordset")
		set rstSelect = Server.CreateObject("ADODB.Recordset")
		set rstpro = Server.CreateObject("ADODB.Recordset")
        set rstPrecioSinIva = Server.CreateObject("ADODB.Recordset")
        set rstPrecioTarifaLocal = Server.CreateObject("ADODB.Recordset")
        set insertTablaCalculoPrecio = Server.CreateObject("ADODB.Recordset")
        set pruebaSelect = Server.CreateObject("ADODB.Recordset")
        set sumTablaCalculoPrecio = Server.CreateObject("ADODB.Recordset")
        set deleteTablaCalculoPrecio = Server.CreateObject("ADODB.Recordset")

		''Leer parámetros de la página
		mode=EncodeForHtml(Request.QueryString("mode"))

		if ucase(mode) = "BROWSE" then mode ="ver"
		if enc.EncodeForJavascript(request.querystring("referencia")) >"" then
			referencia = limpiaCadena(request.querystring("referencia"))
		else
			referencia = limpiaCadena(request.form("referencia"))
		end if

		dim fye, tu, vatlcb,TC, isSolred3x7

		if enc.EncodeForJavascript(request.querystring("fye")) >"" then
			fye = limpiaCadena(request.querystring("fye"))
		else
			fye = limpiaCadena(request.form("fye"))
		end if
		if enc.EncodeForJavascript(request.querystring("TC")) >"" then
			TC = limpiaCadena(request.querystring("TC"))
		else
			TC = limpiaCadena(request.form("TC"))
		end if

	    ''ricardo 26-4-2007 parametro para solamente trabajar con campo01=1
	    dim mcp1, od, noref

		ObtenerParametros "codigo_barras"

		if enc.EncodeForJavascript(request.querystring("cantidad")) >"" then
			cantidad = limpiaCadena(request.querystring("cantidad"))
		else
		    cantidad = limpiaCadena(request.form("cantidad"))
		end if
		if enc.EncodeForJavascript(request.querystring("articulo")) >"" then
			articulo = limpiaCadena(request.querystring("articulo"))
		else
			articulo = limpiaCadena(request.form("articulo"))
		end if

		if enc.EncodeForJavascript(request.querystring("stockmayoroigual")) >"" then
			stockmayoroigual= limpiaCadena(request.querystring("stockmayoroigual"))
		else
			stockmayoroigual= limpiaCadena(request.form("stockmayoroigual"))
		end if

		if enc.EncodeForJavascript(request.querystring("almacen")) >"" then
			almacen= limpiaCadena(request.querystring("almacen"))
		else
			almacen= limpiaCadena(request.form("almacen"))
		end if

        ''if request.querystring("tienda") >"" then
		''	tienda= limpiaCadena(request.querystring("tienda"))
		''else
		''	tienda=limpiaCadena(request.form("tienda"))
		''end if
		
		'dgb  08/04/2008  anyadirmo Almacen para ASIGNACION MASIVA
		if enc.EncodeForJavascript(request.querystring("almacenmasivo")) >"" then
			almacenmasivo= limpiaCadena(request.querystring("almacenmasivo"))
		else
			almacenmasivo= limpiaCadena(request.form("almacenmasivo"))
		end if

		nombre	= limpiaCadena(request.form("nombre"))
		referencia	= limpiaCadena(request.form("referencia"))

		familia	= limpiaCadena(request.form("familia"))
		familia_padre	= limpiaCadena(request.form("familia_padre"))
		categoria	= limpiaCadena(request.form("categoria"))

		ordenar	= limpiaCadena(request.form("ordenar"))
		cod_temporada	= limpiaCadena(request.form("cod_temporada"))
        filtro_temporada	= limpiaCadena(request.form("filtro_temporada"))
		tipodoc	= limpiaCadena(request.form("tipodoc"))
		numdoc	= limpiaCadena(request.form("numdoc"))

		ver_referencia	=	limpiaCadena(request.form("ver_referencia"))
        ver_refProv	=	limpiaCadena(request.form("ver_refProv"))
		ver_nombre		=	limpiaCadena(request.form("ver_nombre"))
		ver_empresa		=	limpiaCadena(request.form("ver_empresa"))
		ver_lineas		=	limpiaCadena(request.form("ver_lineas"))
		ver_precios		=	limpiaCadena(request.form("ver_precios"))
		ver_codTerminal	=	limpiaCadena(request.form("ver_codTerminal"))
		imprimir_listado_horizontal	=	limpiaCadena(request.form("imprimir_listado_horizontal"))
		imprimir_listado_vertical	=	limpiaCadena(request.form("imprimir_listado_vertical"))
		formato_impresion			=	limpiaCadena(request.form("formato_impresion"))
		cant_doc					=	limpiaCadena(request.form("cant_doc"))
		fin_temp					=	limpiaCadena(request.form("fin_temp"))

		fechamodprec=limpiaCadena(request.form("fmpc"))
		solopreciocambiado=limpiaCadena(request.form("solopreciocambiado"))
		opcprec1=limpiaCadena(request.form("opcprec1"))
		opcprec2=limpiaCadena(request.form("opcprec2"))
		tarifa1=limpiaCadena(request.form("tarifa1"))
		tarifa2=limpiaCadena(request.form("tarifa2"))
		tarifaex=limpiaCadena(request.form("tarifaex"))
		tarifaiva1=limpiaCadena(request.form("tarifaiva1"))
		tarifaiva2=limpiaCadena(request.form("tarifaiva2"))
		tarifa=limpiaCadena(request.form("tarifa"))

        ndoc = limpiaCadena(request.QueryString("ndoc"))

		Alarma "codigo_barras_nuevo.asp"%>
		<input type="hidden" name="si_tiene_modulo_terminales" value="<%=EncodeForHtml(si_tiene_modulo_terminales)%>"/>
		<input type="hidden" name="tu" value="<%=EncodeForHtml(tu)%>" />		
		<%'dgb  08/04/2008  anyadimos Almacen para ASIGNACION MASIVA  %>
		<input type="hidden" name="almacenmasivo" value="<%=EncodeForHtml(almacenmasivo)%>"/>
        <input type="hidden" name="noref" value="<%=EncodeForHtml(noref)%>"/>
		<br/>

	<%if ucase(mode) = "EXPORTAR" then%>
		<input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>"/>
		<input type="hidden" name="nombre" value="<%=EncodeForHtml(nombre)%>"/>
		<input type="hidden" name="ordenar" value="<%=EncodeForHtml(ordenar)%>"/>
		<input type="hidden" name="cantidad" value="<%=EncodeForHtml(cantidad)%>"/>
		<input type="hidden" name="tipodoc" value="<%=EncodeForHtml(tipodoc)%>"/>
		<input type="hidden" name="numdoc" value="<%=EncodeForHtml(numdoc)%>"/>
		<input type="hidden" name="ver_referencia" value="<%=EncodeForHtml(ver_referencia)%>"/>
        <input type="hidden" name="ver_refProv" value="<%=EncodeForHtml(ver_refProv)%>"/>
		<input type="hidden" name="fye" value="<%=EncodeForHtml(fye)%>"/>
		<input type="hidden" name="ver_nombre" value="<%=EncodeForHtml(ver_nombre)%>"/>
		<input type="hidden" name="ver_empresa" value="<%=EncodeForHtml(ver_empresa)%>"/>
		<input type="hidden" name="ver_lineas" value="<%=EncodeForHtml(ver_lineas)%>"/>
		<input type="hidden" name="ver_precios" value="<%=EncodeForHtml(ver_precios)%>"/>
		<input type="hidden" name="ver_codterminal" value="<%=EncodeForHtml(ver_codterminal)%>"/>
		<input type="hidden" name="imprimir_listado_horizontal" value="<%=EncodeForHtml(imprimir_listado_horizontal)%>"/>
		<input type="hidden" name="imprimir_listado_vertical" value="<%=EncodeForHtml(imprimir_listado_vertical)%>"/>
		<input type="hidden" name="formato_impresion" value="<%=EncodeForHtml(formato_impresion)%>"/>
		<input type="hidden" name="cant_doc" value="<%=EncodeForHtml(cant_doc)%>"/>
		<input type="hidden" name="fin_temp" value="<%=EncodeForHtml(fin_temp)%>"/>
		<input type="hidden" name="fmpc" value="<%=EncodeForHtml(fechamodprec)%>"/>
		<input type="hidden" name="solopreciocambiado" value="<%=EncodeForHtml(solopreciocambiado)%>"/>
		<input type="hidden" name="tarifa" value="<%=EncodeForHtml(tarifa)%>"/>
		<input type="hidden" name="stockmayoroigual" value="<%=EncodeForHtml(stockmayoroigual)%>"/>
		<input type="hidden" name="almacen" value="<%=EncodeForHtml(almacen)%>"/>
        <!--<input type="hidden" name="tienda" value="<%=EncodeForHtml(tienda)%>"/>-->
		<input type="hidden" name="opcprec1" value="<%=EncodeForHtml(opcprec1)%>"/>
		<input type="hidden" name="opcprec2" value="<%=EncodeForHtml(opcprec2)%>"/>
		<input type="hidden" name="tarifa1" value="<%=EncodeForHtml(tarifa1)%>"/>
		<input type="hidden" name="tarifa2" value="<%=EncodeForHtml(tarifa2)%>"/>
		<input type="hidden" name="tarifaex" value="<%=EncodeForHtml(tarifaex)%>"/>
		<input type="hidden" name="tarifaiva1" value="<%=EncodeForHtml(tarifaiva1)%>"/>
		<input type="hidden" name="tarifaiva2" value="<%=EncodeForHtml(tarifaiva2)%>"/>
		<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>"/>
		<input type="hidden" name="familia_padre" value="<%=EncodeForHtml(familia_padre)%>"/>
		<input type="hidden" name="categoria" value="<%=EncodeForHtml(categoria)%>"/>
		<input type="hidden" name="cod_temporada" value="<%=EncodeForHtml(cod_temporada)%>"/>
        <input type="hidden" name="filtro_temporada" value="<%=EncodeForHtml(filtro_temporada)%>"/>
		<input type="hidden" name="TC" value="<%=EncodeForHtml(TC)%>"/>
        <input type="hidden" name="mcp1" value="<%=EncodeForHtml(mcp1)%>"/>

		<script language="javascript" type="text/javascript">
		    document.codigo_barras.action = "exportar_codigo_barrasn.asp";
		    document.codigo_barras.submit();
		</script><%
	end if

	''*********************************************************************************************
	''Se muestran parametros de seleccion
	''*********************************************************************************************
	if mode="param" then%>
        <%'DrawCelda2 "CELDA style='width:130px'", "left", false, LitConref + ": "
            DrawDiv "1","",""
            DrawLabel "","",LitConRef
            %><input class="CELDA" type="text" name="referencia" value="<%=iif(ndoc & "" <> "", EncodeForHtml(trimcodempresa(ndoc)), "")%>" size="15" onchange="TraerArticulo('<%=enc.EncodeForJavascript(mode)%>','<%=enc.EncodeForJavascript(ndet)%>');"/><a class="CELDAREFB" href="javascript:WinArticulos()"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscarDinamic%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><%CloseDiv
            'DrawCelda2 "CELDA style='width:50px'", "left", false,""
			'DrawCelda2 "CELDA style='width:150px'", "left", false, LitConNombre + ": "
			'DrawInputCelda "CELDA","","",25,0,"","nombre",nombre
            EligeCelda "input","add","left","","",25,LitConNombre,"nombre",0,EncodeForHtml(nombre)
			'DrawCelda2 "CELDA style='width:130px'", "left", false, LitOrdenar2 + ": "
            DrawDiv "1","",""
            DrawLabel "","",LitOrdenar2%><select class="width60" name="ordenar" >
					<option selected="selected" value="REFERENCIA"><%=ucase(LitRef)%></option>
					<option value="NOMBRE"><%=ucase(LitNombre2)%></option>
				</select><%CloseDiv

            ' Matias Lozano Añado un salto para CAPONE
            if session("version")&"" <> "5" then
                DrawDiv "","","" 
                CloseDiv
            end if 
			'DrawCelda2 "CELDA style='width:130px'", "left", false, LitArtTemp + ": "
            rstAux.cursorlocation=3
			rstAux.open " select codigo, descripcion from temporadas with(nolock) where codigo like '" & session("ncliente") & "%' and codigo <> '" & session("ncliente") & "BASE' order by descripcion", session("backendlistados")
			'DrawSelectCelda "CELDA","190","",0,"","cod_temporada",rstAux,cod_temporada,"codigo","descripcion","onchange","control_temporada()"
            'DrawSelectCelda "","","","",LitArtTemp,"cod_temporada",rstAux,cod_temporada,"codigo","descripcion","onchange","control_temporada()"
            DrawDiv "1","",""
            DrawLabel "","",LitArtTemp
            DrawSelect "width30","margin-right:5px;","cod_temporada",rstAux,cod_temporada,"codigo","descripcion","onchange","control_temporada()"
			rstAux.close
            if si_tiene_modulo_Centroxogo<>0 then%><select class="width30" name="filtro_temporada">
                    <option selected="selected" value="incluir"><%=ucase(LitIncluir)%></option>
                    <option value="excluir"><%=ucase(LitExcluir)%></option>
                </select>
            <%else%><select class="width30" name="filtro_temporada">
                    <option selected="selected" value="blanco"></option>
                    <option value="incluir"><%=ucase(LitIncluir)%></option>
                    <option value="excluir"><%=ucase(LitExcluir)%></option>
                </select>
            <%End If
            CloseDiv
            DrawDiv "1","display:none","id_fin_temp"
            DrawLabel "","",LitFinTemp%><input type="checkbox" name="fin_temp"/><%CloseDiv

		
			'DrawCelda2 "CELDA style='width:130px'", "left", false, Littipodoc + ": "
            ''MPC 22/12/2011 If you have a parameter d = 1 assignment is just massive.
			''rstAux.Open "SELECT * FROM Tipo_Documentos with(nolock) where tippdoc in ('ALBARAN DE SALIDA','ALBARAN DE PROVEEDOR','FACTURA A CLIENTE','FACTURA DE PROVEEDOR','HOJA DE GASTOS','PEDIDO A PROVEEDOR','PEDIDO DE CLIENTE','PRESUPUESTO A CLIENTE','ORDEN DE FABRICACION','MOVIMIENTOS ENTRE ALMACENES') union select'ASIGNACION MASIVA', '666' ",DsnIlion,adOpenKeyset, adLockOptimistic
			if od = "1" then
                strSelect = "select 'ASIGNACION MASIVA' as TippDoc, '666' as codigo"
                rstAux.cursorlocation=3
                rstAux.Open strSelect,DsnIlion
                'DrawSelectCelda "CELDA","190","","0","","tipodoc",rstAux,tipodoc,"TippDoc","TippDoc","onchange","mostrar_cantidad(0)"
                DrawSelectCelda "width60","","","0",Littipodoc,"tipodoc",rstAux,tipodoc,"TippDoc","TippDoc","onchange","mostrar_cantidad(0)"
			    rstAux.close
            else
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

                listaIn = "'ALBARAN DE SALIDA','ALBARAN DE PROVEEDOR','FACTURA A CLIENTE','FACTURA DE PROVEEDOR','HOJA DE GASTOS','PEDIDO A PROVEEDOR','PEDIDO DE CLIENTE','PRESUPUESTO A CLIENTE','ORDEN DE FABRICACION','MOVIMIENTOS ENTRE ALMACENES'"
                addListaIn = "'ASIGNACION MASIVA'"
                command.Parameters.Append command.CreateParameter("@inlist",adVarChar,adParamInput,len(listaIn),listaIn)
                command.Parameters.Append command.CreateParameter("@outlist",adVarChar,adParamInput,1,"")
                command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,session("usuario"))
                command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))
                command.Parameters.Append command.CreateParameter("@addlist",adVarChar,adParamInput,len(addListaIn),addListaIn)

                set rstTD = Server.CreateObject("ADODB.Recordset")
                set rstTD = command.Execute
                if not rstTD.eof then
                    'DrawSelectCelda "CELDA","190","","0","","tipodoc",rstTD,tipodoc,"tippdoc","descripcion","onchange","mostrar_cantidad(0)"
                    DrawSelectCelda "CELDA","190","","0",Littipodoc,"tipodoc",rstTD,tipodoc,"tippdoc","descripcion","onchange","mostrar_cantidad(0)"
                end if			
                rstTD.close
                conn.close
                set command=nothing
                set rstTD =nothing       
            end if
			
			'DrawCelda2 "CELDA style='width:50px'", "left", false,""
			'DrawCelda2 "CELDA style='width:150px'", "left", false, LitNumDoc + ": "
            DrawDiv "1","",""
            DrawLabel "","",LitNumDoc%><input class="CELDA" type="text" name="numdoc" size="25" maxlength="20" value="<%=EncodeForHtml(numdoc)%>" onkeyup="javascript:mostrar_cantidad(1)" onmousedown="javascript:mostrar_cantidad(2)" onblur="javascript:mostrar_cantidad(3)"/><%CloseDiv
		
			'DrawCelda2 "CELDA style='width:130px'", "left", false, LitCantidad2 + ": "
			if cantidad="" then cantidad="1"
		      'DrawInputCelda "CELDA","","",3,0,"","cantidad",cantidad
                EligeCelda "input","add","","","",3,LitCantidad2,"cantidad",0,EncodeForHtml(cantidad)
			'DrawCelda2 "CELDA style='width:50px'", "left", false,""
            DrawDiv "1","display:none","idcant_doc2"
            DrawLabel "","",LitCodBarrImpCantDoc%><input type="checkbox" name="cant_doc" onclick="javascript:control_cantidad()"/><%CloseDiv

        
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
				ConfigDespleg(i,9)=LitCategoria & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
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
				ConfigDespleg(i,9)=LitFamilia2
				ConfigDespleg(i,10)=familia_padre
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""
				i=2
				ConfigDespleg(i,0)="familia"
				ConfigDespleg(i,1)="200"
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="CELDA"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitSubFamilia
				ConfigDespleg(i,10)=familia
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				DibujaDesplegables ConfigDespleg,session("backendlistados")
            ''Se recupera el select multiple de tarifa
            'DrawCelda2 "CELDA", "left", false, LITTARIFALISTADOCODIGOBARRAS & ": "
			StrIn=""
			if tu>"" then
				StrIn=StrIn&" and codigo in "& replace(replace(replace(tu,",","','"),"(","('"),")","')")
			end if
            rstAux.cursorlocation=3
			rstAux.Open "SELECT codigo, descripcion FROM tarifas with(nolock) where codigo like '"&session("ncliente")&"%' "&strIn&" and codigo <>'"&session("ncliente")&"BASE'  order by descripcion",session("backendlistados")
            DrawDiv "1","",""
            DrawLabel "","",LITTARIFALISTADOCODIGOBARRAS%><select class='CELDA' multiple="multiple" size='8' style='width:200' name='tarifa' onchange="MuestraFechaMod2(this)" onfocus="MuestraFechaMod2(this)">
			    <%while not rstAux.eof %>
			    	<option value='<%=EncodeForHtml(rstAux("codigo"))%>'><%=EncodeForHtml(rstAux("descripcion"))%></option>
			    	<%rstAux.movenext
			    wend%>
			    <option selected="selected" value=''></option>
	            </select><%CloseDiv
            rstAux.close

			''ricardo 21-2-2007 se añade el stock y el almacen
			DrawDiv "1","","stock1"
            DrawLabel "","",LitStockMayorOIgualCodBarras%><input class='CELDA' type='text' name='stockmayoroigual' size="7" maxlength="10" value="0"/><%CloseDiv

			Tienda_defecto=""
			rstAux.cursorlocation=3
            rstAux.open "select codigo, descripcion from almacenes with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
            DrawDiv "1","","stock3"
            DrawLabel "","display: ' id='stock2'",LitAlmacenCodBarras%><select class='width60' name="almacen" >
			    <%while not rstAux.eof
                    if Almacen_defecto=rstAux("descripcion") then%>
						<option selected="selected" value="<%=EncodeForHtml(rstAux("codigo"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
					<%else%>
						<option value="<%=EncodeForHtml(rstAux("codigo"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
					<%end if
					rstAux.movenext
				wend
                if Almacen_defecto & ""="" then%>
					<option selected="selected" value=""></option>
				<%end if%>
			</select><%rstAux.close
                CloseDiv
			''DrawCelda2 "CELDA", "left", false, LitSoloPrecioCambiado & ": "
			DrawDiv "1","",""
            DrawLabel "","",LitSoloPrecioCambiado%><input class='CELDA' type='checkbox' name='solopreciocambiado' onclick="MuestraFechaMod(this);"/><%CloseDiv
		
		    DrawDiv "1","","id_fmpc"
            DrawLabel "","",LitArtModPrecFecPos%><%if fye="0" then%><input type="text" class="CELDA" size="13" maxlength="10" name="fmpc" value="" onchange="MuestraFechaMod3(this);"/>
				<%else%><input type="text" class="CELDA" size="13" maxlength="10" name="fmpc" value="<%=iif(fechamodprec>"",EncodeForHtml(fechamodprec),date-1)%>" onchange="MuestraFechaMod3(this);"/>
				<%end if%><%CloseDiv
            DrawCalendar "fmpc"%>
	<hr/>
	<% 
			'DrawCelda2 "CELDA", "left", false,LitFormImprCodBarras
        

''ricardo 19-7-2005 se pone pasar_a_terra=0 para pasar a Terra
if pasar_a_terra=0 then

end if ''ricardo 19-7-2005
        EligeCelda "check","add","left","","",0,LitRef,"ver_referencia",0,"True"
		EligeCelda "check","add","left","","",0,LitNombre2,"ver_nombre",0,"True"
		if fye="0" then
			'DrawCheckCelda "CELDA","","",0,"","ver_empresa","False"
            EligeCelda "check","add","left","","",0,LitEmpresa,"ver_empresa",0,"False"
		else
			'DrawCheckCelda "CELDA","","",0,"","ver_empresa","True"
            EligeCelda "check","add","left","","",0,LitEmpresa,"ver_empresa",0,"True"
		end if
		'DrawCelda2 "CELDA", "left", false, LitEmpresa
		'DrawCheckCelda "CELDA","","",0,"","ver_lineas","True"
        EligeCelda "check","add","left","","",0,LitLineas,"ver_lineas",0,"True"
		'DrawCelda2 "CELDA", "left", false, LitLineas
		'DrawCheckCelda "CELDA","","",0,"","ver_precios","True"
        EligeCelda "check","add","left","","",0,LitPrecios2,"ver_precios",0,"True"
		'DrawCelda2 "CELDA", "left", false, LitPrecios2
		if si_tiene_modulo_terminales<>0 then
			'DrawCheckCelda "CELDA","","",0,"","ver_codTerminal","False"
			'DrawCelda2 "CELDA", "left", false, LitCodTerminal
            EligeCelda "check","add","left","","",0,LitCodTerminal,"ver_codTerminal",0,"False"
		end if
        'DrawCheckCelda "CELDA","","",0,"","ver_refProv","False"
		'DrawCelda2 "CELDA", "left", false, LitRefProv
        EligeCelda "check","add","left","","",0,LitRefProv,"ver_refProv",0,"False"
   %><span id="IMPORTESADICIONALES" style="display:none"><hr/><%

        DrawDiv "col-lg-6 col-md-12 col-xs-12 col-sm-12","",""
        DrawLabel "","",litOpcImporteAdicional1
        CloseDiv

		
        DrawDiv "col-lg-6 col-md-12 col-xs-12 col-sm-12","",""
        DrawLabel "","",litOpcImporteAdicional2
        CloseDiv


        rstAux.cursorlocation=3
		rstAux.open " select codigo, descripcion from tarifas with(nolock) where codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", session("backendlistados")
		haytarifas=0

		if not rstAux.eof then haytarifas=1
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
		    DrawDiv "4","",""%><input type="radio" name="opcprec1" value="tarifa"/><%CloseDiv
            DrawSelectCelda "","","",0,LitTarifaListadoCodigoBarras,"tarifa1",rstAux,tarifa1,"codigo","descripcion","",""
        %></div><%

        if haytarifas=1 then rstaux.movefirst   
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4","",""%><input type="radio" name="opcprec2" value="tarifa"/><%CloseDiv
            DrawSelectCelda "","","",0,LitTarifaListadoCodigoBarras,"tarifa2",rstAux,tarifa2,"codigo","descripcion","",""
        %></div><%     

        if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4","",""%><input type="radio" name="opcprec1" value="tarifaiva"/><%CloseDiv 
            DrawSelectCelda "","","",0,LitTarifaIvaListadoCodigoBarras,"tarifa1",rstAux,tarifaiva1,"codigo","descripcion","",""
        %></div><%

        if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4","",""%><input type="radio" name="opcprec2" value="tarifaiva"/><%CloseDiv
            DrawSelectCelda "","","",0,LitTarifaIvaListadoCodigoBarras,"tarifa2",rstAux,tarifaiva2,"codigo","descripcion","",""
        %></div><%

        if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4", "", ""%><input type="radio" name="opcprec1" value="coste"/><%CloseDiv
				DrawDiv "1", "", ""
					DrawLabel "", "", LitCosteListadoCodigoBarras 
				CloseDiv
        %></div><%

		if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4", "", ""%><input type="radio" name="opcprec2" value="coste"/><%CloseDiv
				DrawDiv "1", "", ""
					DrawLabel "", "", LitCosteListadoCodigoBarras 
				CloseDiv
        %></div><%

	    if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%   
            DrawDiv "4", "", ""%><input type="radio" name="opcprec1" value="none" checked="checked"/><%CloseDiv
				DrawDiv "1", "", ""
					DrawLabel "", "", LitNingunoListadoCodigoBarras 
				CloseDiv
        %></div><%		    

		if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%   
            DrawDiv "4", "", ""%><input type="radio" name="opcprec2" value="none" checked="checked"/><%CloseDiv
				DrawDiv "1", "", ""
					DrawLabel "", "", LitNingunoListadoCodigoBarras 
				CloseDiv
	    %></div><%	

		rstAux.close
	%>
	</span><hr/>
	<% 
		'DrawCelda2 "CELDA style='width:150px'", "left", false, LitMsgHorizontal
		'DrawInputCelda "CELDA style='width:30px'","","",3,0,"","imprimir_listado_horizontal","1"
        EligeCelda "input","add","left","","",3,LitMsgHorizontal,"imprimir_listado_horizontal",0,"1"
		'DrawCelda2 "CELDA style='width:150px'", "left", false, LitMsgVertical
		'DrawInputCelda "CELDA style='width:30px'","","",3,0,"","imprimir_listado_vertical","1"
        EligeCelda "input","add","left","","",3,LitMsgVertical,"imprimir_listado_vertical",0,"1"

''ricardo 19-7-2005 se pone pasar_a_terra=0 para pasar a Terra
if pasar_a_terra=0 then
            rstAux.cursorlocation=3
			rstAux.open " select codigo, case when codigo='" & session("ncliente") & "BASE' then 'PRECIO FICHA ARTICULO' else descripcion end as desccrip from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by desccrip", session("backendlistados")
			tarifaex = session("ncliente") & "BASE"
            DrawDiv "1", "", ""%><label><%DrawHref "CELDAREFB", "", LitExportar, "javascript:if(ValidarCampos()){codigo_barras.action='codigo_barras_nuevo.asp?mode=exportar';codigo_barras.submit();parent.botones.document.location='codigo_barras_bt.asp?mode=exportar';}"%></label><%DrawSelect "100","","tarifaex",rstAux,tarifaex,"codigo","desccrip","",""
			rstAux.close
        %><!--<a class="CELDAREFB" href="javascript:if(ValidarCampos()){codigo_barras.action='codigo_barras.asp?mode=exportar';codigo_barras.submit();parent.botones.document.location='codigo_barras_bt.asp?mode=exportar';}"><%=LitExportar%></a>--><%
            CloseDiv

else
	%><%
end if ''ricardo 19-7-2005

	%>
	<hr/>
	<table><%
		DrawFila color_blau
			%><input type="hidden" name="fye" value="<%=EncodeForHtml(fye)%>"/>
			<input type="hidden" name="maxpagina_form1" value="<%=EncodeForHtml(cuantasEtiqPorPag24)%>"/>
			<input type="hidden" name="cantHMax_form1" value="<%=EncodeForHtml(AnchoForm1)%>"/>
			<input type="hidden" name="cantVMax_form1" value="<%=EncodeForHtml(AltoForm1)%>"/>
			<input type="hidden" name="maxpagina_form2" value="<%=EncodeForHtml(cuantasEtiqPorPag18)%>"/>
			<input type="hidden" name="cantHMax_form2" value="<%=EncodeForHtml(AnchoForm2)%>"/>
			<input type="hidden" name="cantVMax_form2" value="<%=EncodeForHtml(AltoForm2)%>"/>
			<input type="hidden" name="maxpagina_form4" value="<%=EncodeForHtml(cuantasEtiqPorPag36)%>"/>
			<input type="hidden" name="cantHMax_form4" value="<%=EncodeForHtml(AnchoForm4)%>"/>
			<input type="hidden" name="cantVMax_form4" value="<%=EncodeForHtml(AltoForm4)%>"/>
			<input type="hidden" name="maxpagina_form5" value="<%=EncodeForHtml(cuantasEtiqPorPag40)%>"/>
			<input type="hidden" name="cantHMax_form5" value="<%=EncodeForHtml(AnchoForm5)%>"/>
			<input type="hidden" name="cantVMax_form5" value="<%=EncodeForHtml(AltoForm5)%>"/>
			<input type="hidden" name="maxpagina_form6" value="<%=EncodeForHtml(cuantasEtiqPorPag8x3)%>"/>
			<input type="hidden" name="cantHMax_form6" value="<%=EncodeForHtml(AnchoForm6)%>"/>
			<input type="hidden" name="cantVMax_form6" value="<%=EncodeForHtml(AltoForm6)%>"/>
			<input type="hidden" name="maxpagina_form7" value="<%=EncodeForHtml(cuantasEtiqPorPag8x3P)%>"/>
			<input type="hidden" name="cantHMax_form7" value="<%=EncodeForHtml(AnchoForm7)%>"/>
			<input type="hidden" name="cantVMax_form7" value="<%=EncodeForHtml(AltoForm7)%>"/>
			<input type="hidden" name="maxpagina_form8" value="<%=EncodeForHtml(cuantasEtiqPorPag8x3PP)%>"/>
			<input type="hidden" name="cantHMax_form8" value="<%=EncodeForHtml(AnchoForm8)%>"/>
			<input type="hidden" name="cantVMax_form8" value="<%=EncodeForHtml(AltoForm8)%>"/>
			<input type="hidden" name="maxpagina_form9" value="<%=EncodeForHtml(cuantasEtiqPorPag7x2)%>"/>
			<input type="hidden" name="cantHMax_form9" value="<%=EncodeForHtml(AnchoForm9)%>"/>
			<input type="hidden" name="cantVMax_form9" value="<%=EncodeForHtml(AltoForm9)%>"/>
			<input type="hidden" name="maxpagina_form10" value="<%=EncodeForHtml(cuantasEtiqPorPag7x2)%>"/>
			<input type="hidden" name="cantHMax_form10" value="<%=EncodeForHtml(AnchoForm10)%>"/>
			<input type="hidden" name="cantVMax_form10" value="<%=EncodeForHtml(AltoForm10)%>"/>
			<input type="hidden" name="maxpagina_formCHACAL" value="<%=EncodeForHtml(cuantasEtiqPorPag24)%>"/>
			<input type="hidden" name="cantHMax_formCHACAL" value="<%=EncodeForHtml(AnchoForm1)%>"/>
			<input type="hidden" name="cantVMax_formCHACAL" value="<%=EncodeForHtml(AltoForm1)%>"/>

			<input type="hidden" name="maxpagina_form11" value="<%=EncodeForHtml(cuantasEtiqPorPag44)%>"/>
			<input type="hidden" name="cantHMax_form11" value="<%=EncodeForHtml(AnchoForm11)%>"/>
			<input type="hidden" name="cantVMax_form11" value="<%=EncodeForHtml(AltoForm11)%>"/>

			<input type="hidden" name="maxpagina_form12" value="<%=EncodeForHtml(cuantasEtiqPorPag44)%>"/>
			<input type="hidden" name="cantHMax_form12" value="<%=EncodeForHtml(AnchoForm12)%>"/>
			<input type="hidden" name="cantVMax_form12" value="<%=EncodeForHtml(AltoForm12)%>"/>

			<input type="hidden" name="maxpagina_form13" value="<%=EncodeForHtml(cuantasEtiqPorPag24Margen)%>"/>
			<input type="hidden" name="cantHMax_form13" value="<%=EncodeForHtml(AnchoForm13)%>"/>
			<input type="hidden" name="cantVMax_form13" value="<%=EncodeForHtml(AltoForm13)%>"/>

			<input type="hidden" name="maxpagina_form14" value="<%=EncodeForHtml(cuantasEtiqPorPag24Margen)%>"/>
			<input type="hidden" name="cantHMax_form14" value="<%=EncodeForHtml(AnchoForm14)%>"/>
			<input type="hidden" name="cantVMax_form14" value="<%=EncodeForHtml(AltoForm14)%>"/>

			<input type="hidden" name="maxpagina_form15" value="<%=EncodeForHtml(cuantasEtiqPorPag7x2)%>"/>
			<input type="hidden" name="cantHMax_form15" value="<%=EncodeForHtml(AnchoForm15)%>"/>
			<input type="hidden" name="cantVMax_form15" value="<%=EncodeForHtml(AltoForm15)%>"/>

			<input type="hidden" name="maxpagina_form19" value="<%=EncodeForHtml(cuantasEtiqPorPag45)%>"/>
			<input type="hidden" name="cantHMax_form19" value="<%=EncodeForHtml(AnchoForm19)%>"/>
			<input type="hidden" name="cantVMax_form19" value="<%=EncodeForHtml(AltoForm19)%>"/>

			<input type="hidden" name="maxpagina_form20" value="<%=EncodeForHtml(cuantasEtiqPorPagRollo)%>"/>
			<input type="hidden" name="cantHMax_form20" value="<%=EncodeForHtml(AnchoForm20)%>"/>
			<input type="hidden" name="cantVMax_form20" value="<%=EncodeForHtml(AltoForm20)%>"/>

			<input type="hidden" name="maxpagina_form21" value="<%=EncodeForHtml(cuantasEtiqPorPagRollo)%>"/>
			<input type="hidden" name="cantHMax_form21" value="<%=EncodeForHtml(AnchoForm21)%>"/>
			<input type="hidden" name="cantVMax_form21" value="<%=EncodeForHtml(AltoForm21)%>"/>

			<input type="hidden" name="maxpagina_form22" value="<%=EncodeForHtml(cuantasEtiqPorPag20)%>"/>
			<input type="hidden" name="cantHMax_form22" value="<%=EncodeForHtml(AnchoForm22)%>"/>
			<input type="hidden" name="cantVMax_form22" value="<%=EncodeForHtml(AltoForm22)%>"/>

			<input type="hidden" name="maxpagina_form23" value="<%=EncodeForHtml(cuantasEtiqPorPag40)%>"/>
			<input type="hidden" name="cantHMax_form23" value="<%=EncodeForHtml(AnchoForm23)%>"/>
			<input type="hidden" name="cantVMax_form23" value="<%=EncodeForHtml(AltoForm23)%>"/>

			<input type="hidden" name="maxpagina_form24" value="<%=EncodeForHtml(cuantasEtiqPorPag16)%>"/>
			<input type="hidden" name="cantHMax_form24" value="<%=EncodeForHtml(AnchoForm24)%>"/>
			<input type="hidden" name="cantVMax_form24" value="<%=EncodeForHtml(AltoForm24)%>"/>

			<input type="hidden" name="maxpagina_form25" value="<%=EncodeForHtml(cuantasEtiqPorPagRollo)%>"/>
			<input type="hidden" name="cantHMax_form25" value="<%=EncodeForHtml(AnchoForm25)%>"/>
			<input type="hidden" name="cantVMax_form25" value="<%=EncodeForHtml(AltoForm25)%>"/>
			
			<input type="hidden" name="maxpagina_form26" value="<%=EncodeForHtml(cuantasEtiqPorPag18)%>"/>
			<input type="hidden" name="cantHMax_form26" value="<%=EncodeForHtml(AnchoForm26)%>"/>
			<input type="hidden" name="cantVMax_form26" value="<%=EncodeForHtml(AltoForm26)%>"/>

			<input type="hidden" name="maxpagina_form27" value="<%=EncodeForHtml(cuantasEtiqPorPag39)%>"/>
			<input type="hidden" name="cantHMax_form27" value="<%=EncodeForHtml(AnchoForm27)%>"/>
			<input type="hidden" name="cantVMax_form27" value="<%=EncodeForHtml(AltoForm27)%>"/>
			
			<input type="hidden" name="maxpagina_form28" value="<%=EncodeForHtml(cuantasEtiqPorPag18)%>"/>
			<input type="hidden" name="cantHMax_form28" value="<%=EncodeForHtml(AnchoForm28)%>"/>
			<input type="hidden" name="cantVMax_form28" value="<%=EncodeForHtml(AltoForm28)%>"/>

            <input type="hidden" name="maxpagina_form32" value="<%=EncodeForHtml(cuantasEtiqPorPag21)%>"/>
            <input type="hidden" name="cantHMax_form32" value="<%=EncodeForHtml(AnchoForm32)%>"/>
            <input type="hidden" name="cantVMax_form32" value="<%=EncodeForHtml(AltoForm32)%>"/>

            <input type="hidden" name="maxpagina_formSolred3x7" value="<%=EncodeForHtml(cuantasEtiqPorPag21)%>" />
            <input type="hidden" name="cantHMax_formSolred3x7" value="<%=EncodeForHtml(AnchoFormSolred3x7)%>" />
            <input type="hidden" name="cantVMax_formSolred3x7" value="<%=EncodeForHtml(AltoFormSolred3x7)%>" />

<%
	%></table><%

        DrawDiv "1", "", ""
			
''ricardo 13-3-20003
''si la serie tiene un formato de impresion sera este el de por defecto
''si no sera el elegido en la tabla formatos impresion de ilion
		defecto=""
		if nserie & "">"" then
			defecto=obtener_formato_imp(nserie,"ETIQUETAS DE ARTICULOS")
		end if

			seleccion = "select b.fichero as fichero, a.descripcion as descripcion,a.personalizacion,b.parametros as parametros,a.defecto from clientes_formatos_imp as a with(NOLOCK), formatos_imp as b with(NOLOCK) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ETIQUETAS DE ARTICULOS' order by descripcion"
			rstSelect.cursorlocation=3
			rstSelect.Open seleccion, DsnIlion'', adOpenKeyset, adLockOptimistic
			DrawLabel "", "",LitFormImprCodBarras%><select class='width60'  name="formato_impresion" onchange="Cambio();">
			<%
				no_habia_fin=0
				if not rstSelect.eof then
					no_habia_fin=1
				end if

				while not rstSelect.eof and defecto & ""=""
					if rstSelect("defecto")<>0 then
						defecto=rstSelect("descripcion")
					end if
					rstSelect.movenext
				wend
				if no_habia_fin=1 then
					rstSelect.movefirst
				end if
				encontrado=0
				while not rstSelect.eof
					if defecto=rstSelect("descripcion") then
						encontrado=1
						if isnull(rstSelect("parametros")) then
							prm=""
						else
							prm=rstSelect("parametros") & "&"
						end if
						%><option selected="selected" value="<%=EncodeForHtml(rstSelect("fichero") & iif(prm>"","?" & prm,""))%>"><%=EncodeForHtml(rstSelect("descripcion"))%></option><%
					else
						if isnull(rstSelect("parametros")) then
							prm=""
						else
							prm=rstSelect("parametros") & "&"
						end if
						%><option value="<%=EncodeForHtml(rstSelect("fichero") & iif(prm>"","?" & prm,""))%>"><%=EncodeForHtml(rstSelect("descripcion"))%></option><%
					end if
					rstSelect.movenext
				wend

			%></select><%
			rstSelect.close
			set rstSelect=nothing
		CloseDiv
		%><script language="javascript" type="text/javascript">Cambio();</script><%
	end if

	''*********************************************************************************************
	'' Se muestran los datos de la consulta
	''*********************************************************************************************

	if mode="ver" then
	    if enc.EncodeForJavascript(request.querystring("nodoc"))=1 then
	        rstSelect.open "insert into auditoria (nempresa,login,usuario,fecha,ip,accion,descripcion) values('"&session("ncliente")&"','"&session("usuario")&"','','"&now&"','','Listado codigo barras','impresion sin seleccion de documento')",session("dsn_cliente")
	    end if

		%><input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>"/>
		<input type="hidden" name="nombre" value="<%=EncodeForHtml(nombre)%>"/>
		<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>"/>
		<input type="hidden" name="familia_padre" value="<%=EncodeForHtml(familia_padre)%>"/>
		<input type="hidden" name="categoria" value="<%=EncodeForHtml(categoria)%>"/>
		<input type="hidden" name="ordenar" value="<%=EncodeForHtml(ordenar)%>"/>
		<input type="hidden" name="cantidad" value="<%=EncodeForHtml(cantidad)%>"/>
		<input type="hidden" name="tipodoc" value="<%=EncodeForHtml(tipodoc)%>"/>
		<input type="hidden" name="numdoc" value="<%=EncodeForHtml(numdoc)%>"/>
		<input type="hidden" name="ver_referencia" value="<%=EncodeForHtml(ver_referencia)%>"/>
        <input type="hidden" name="ver_refProv" value="<%=EncodeForHtml(ver_refProv)%>"/>
		<input type="hidden" name="fye" value="<%=EncodeForHtml(fye)%>"/>
		<input type="hidden" name="ver_nombre" value="<%=EncodeForHtml(ver_nombre)%>"/>
		<input type="hidden" name="ver_empresa" value="<%=EncodeForHtml(ver_empresa)%>"/>
		<input type="hidden" name="ver_lineas" value="<%=EncodeForHtml(ver_lineas)%>"/>
		<input type="hidden" name="ver_precios" value="<%=EncodeForHtml(ver_precios)%>"/>
		<input type="hidden" name="ver_codterminal" value="<%=EncodeForHtml(ver_codterminal)%>"/>
		<input type="hidden" name="imprimir_listado_horizontal" value="<%=EncodeForHtml(imprimir_listado_horizontal)%>"/>
		<input type="hidden" name="imprimir_listado_vertical" value="<%=EncodeForHtml(imprimir_listado_vertical)%>"/>
		<input type="hidden" name="formato_impresion" value="<%=EncodeForHtml(formato_impresion)%>"/>
		<input type="hidden" name="cant_doc" value="<%=EncodeForHtml(cant_doc)%>"/>
		<input type="hidden" name="fin_temp" value="<%=EncodeForHtml(fin_temp)%>"/>
		<input type="hidden" name="fmpc" value="<%=EncodeForHtml(fechamodprec)%>"/>
		<input type="hidden" name="solopreciocambiado" value="<%=EncodeForHtml(solopreciocambiado)%>"/>
		<input type="hidden" name="tarifa" value="<%=EncodeForHtml(tarifa)%>"/>
		<input type="hidden" name="stockmayoroigual" value="<%=EncodeForHtml(stockmayoroigual)%>"/>
		<input type="hidden" name="almacen" value="<%=EncodeForHtml(almacen)%>"/>
        <!--<input type="hidden" name="tienda" value="<%=EncodeForHtml(tienda)%>"/>-->

		<input type="hidden" name="opcprec1" value="<%=EncodeForHtml(opcprec1)%>"/>
		<input type="hidden" name="opcprec2" value="<%=EncodeForHtml(opcprec2)%>"/>
		<input type="hidden" name="tarifa1" value="<%=EncodeForHtml(tarifa1)%>"/>
		<input type="hidden" name="tarifa2" value="<%=EncodeForHtml(tarifa2)%>"/>
		<input type="hidden" name="tarifaex" value="<%=EncodeForHtml(tarifaex)%>"/>
		<input type="hidden" name="tarifaiva1" value="<%=EncodeForHtml(tarifaiva1)%>"/>
		<input type="hidden" name="tarifaiva2" value="<%=EncodeForHtml(tarifaiva2)%>"/>
		<input type="hidden" name="cod_temporada" value="<%=EncodeForHtml(cod_temporada)%>"/>
        <input type="hidden" name="filtro_temporada" value="<%=EncodeForHtml(filtro_temporada)%>"/>
<%

''RGU
		''Creamos el select el from y el where del procedimiento que creará la tabla con los datos de las etiquetas
		devolver_error=0

		seriedoc=""
		tabladoc=""
        excluir=0

 		if tipodoc>"" or numdoc>"" then
	 		 select case tipodoc
				case "ALBARAN DE SALIDA":
					tipo_doc_tabla="detalles_alb_cli"
					tabladoc="albaranes_cli"
				case "ALBARAN DE PROVEEDOR":
					tipo_doc_tabla="detalles_alb_pro"
					tabladoc="albaranes_pro"
				case "DEVOLUCION A PROVEEDOR":
					tipo_doc_tabla="detalles_dev_pro"
					tabladoc="devoluciones_pro"
				case "DEVOLUCION DE CLIENTE":
					tipo_doc_tabla="detalles_dev_cli"
					tabladoc="devoluciones_cli"
				case "FACTURA A CLIENTE":
					tipo_doc_tabla="detalles_fac_cli"
					tabladoc="facturas_cli"
				case "FACTURA DE PROVEEDOR":
					tipo_doc_tabla="detalles_fac_pro"
					tabladoc="facturas_pro"
				case "HOJA DE GASTOS":
					devolver_error=1
					tabladoc="hojas_gastos"
				case "PEDIDO A PROVEEDOR":
					tipo_doc_tabla="detalles_ped_pro"
					tabladoc="pedidos_pro"
				case "PEDIDO DE CLIENTE":
					tipo_doc_tabla="detalles_ped_cli"
					tabladoc="pedidos_cli"
				case "PRESUPUESTO A CLIENTE":
					'devolver_error=1
					tipo_doc_tabla="detalles_pre_cli"
					tabladoc="presupuestos_cli"
				case "TICKET":
					tipo_doc_tabla="detalles_tickets"
					tabladoc="tickets"
				case "ASIGNACION MASIVA":
					'numdoc = "00000" + session("ncliente")
					numdoc = "00000" + almacenmasivo
					tipo_doc_tabla="detalles_pre_cli"
					tabladoc="presupuestos_cli"
				case "ORDEN DE FABRICACION":
					tipo_doc_tabla="detalles_orden_fab"
					tabladoc="ordenes_fab"
				case "MOVIMIENTOS ENTRE ALMACENES":
				    tipo_doc_tabla="detalles_movimientos"
				    tabladoc="movimientos"
			 end select
			 if numdoc>"" then
			 	select case tipodoc
			 			case "ALBARAN DE SALIDA":
			 				num_doc_tabla="nalbaran"
			 			case "ALBARAN DE PROVEEDOR":
                            rst.cursorlocation=3
			 				rst.open "select nalbaran from albaranes_pro with(nolock) where nalbaran_pro='" & numdoc &"' and nalbaran like '" & session("ncliente") & "%'", session("backendlistados")
			 	           	if not rst.EOF then
		 						numdoc=trimCodEmpresa(rst("nalbaran"))
		 						num_doc_tabla="nalbaran"
			 				else
		 						devolver_error=1
			 				end if
			 				rst.close
			 			case "DEVOLUCION A PROVEEDOR":
			 				num_doc_tabla="ndocumento"
			 			case "DEVOLUCION DE CLIENTE":
			 				num_doc_tabla="ndocumento"
			 			case "FACTURA A CLIENTE":
			 				num_doc_tabla="nfactura"
			 			case "FACTURA DE PROVEEDOR":
                            rst.cursorlocation=3
			 				rst.open "select nfactura from facturas_pro with(nolock) where nfactura_pro='" & numdoc &"' and nfactura like '" & session("ncliente") & "%'", session("backendlistados")
			 	            if not rst.EOF then
			 					numdoc=trimCodEmpresa(rst("nfactura"))
			 					num_doc_tabla="nfactura"
			 				else
			 					devolver_error=1
			 				end if
			 				rst.close
			 			case "HOJA DE GASTOS":
			 				devolver_error=1
			 			case "PEDIDO A PROVEEDOR":
			 				num_doc_tabla="npedido"
			 			case "PEDIDO DE CLIENTE":
			 				num_doc_tabla="npedido"
			 			case "PRESUPUESTO A CLIENTE":
			 				'devolver_error=1
			 				num_doc_tabla="npresupuesto"
			 			case "TICKET":
			 				num_doc_tabla="nticket"
			 			case "ASIGNACION MASIVA":
			 				num_doc_tabla="npresupuesto"
			 			case "ORDEN DE FABRICACION":
			 				num_doc_tabla="norden"
			 			case "MOVIMIENTOS ENTRE ALMACENES":
			 			    num_doc_tabla="nmovimiento"
			 	end select
			 end if
			 if tabladoc & "">"" and numdoc & "">"" and num_doc_tabla & "">"" then
			 	if tipodoc = "ASIGNACION MASIVA" then
					'seriedoc=d_lookup("serie",tabladoc,num_doc_tabla & "='00000" & almacenmasivo & "'",session("backendlistados"))
                    strselectD="select serie from " & tabladoc & " where " & num_doc_tabla & " = ? "
                    seriedoc=DLookupP1(strselectD, "00000" & almacenmasivo, adVarChar,10,session("backendlistados"))
				elseif tipodoc = "MOVIMIENTOS ENTRE ALMACENES" then
				    'seriedoc=d_lookup("nserie",tabladoc,num_doc_tabla & "='" & session("ncliente") & numdoc & "'",session("backendlistados"))
                    strselectD="select nserie from " & tabladoc & " where " & num_doc_tabla & " = ? "
                    seriedoc=DLookupP1(strselectD, session("ncliente") & numdoc, adVarChar,10,session("backendlistados"))
				else
				    'seriedoc=d_lookup("serie",tabladoc,num_doc_tabla & "='" & session("ncliente") & numdoc & "'",session("backendlistados"))
                    strselectD="select serie from " & tabladoc & " where " & num_doc_tabla & " = ? "
                    seriedoc=DLookupP1(strselectD, session("ncliente") & numdoc, adVarChar,10,session("backendlistados"))
				end if
			 end if
		end if
		        
            ''Matias Lozano 29-08-2018 se recupera el filtro tarifa
			if (cod_temporada&"">"" and fin_temp<>"on") or (tarifa&"" > "") then
				''ricardo 22-2-2007 si es asignacion masiva se podran los precios del presupuesto
				if tipodoc = "ASIGNACION MASIVA" then
					strselect= "select p.pvp  as pvp "
				else
					strselect= "select isnull((case when precios.es_dto=0 then precios.pvpdto else case when precios.es_dto=1 then (a.pvp + ((a.pvp*precios.pvpdto)/100)) else (a.importe + ((a.importe*precios.pvpdto)/100)) end end),a.pvp)  as pvp "
				end if
			else
				''ricardo 22-2-2007 si es asignacion masiva se podran los precios del presupuesto
				if tipodoc = "ASIGNACION MASIVA" then
					strselect = "select p.pvp "
				else
					strselect = "select a.pvp "
				end if
			end if

			strselect=strselect & ",a.divisa,a.cod_barras "
			''ricardo 22-2-2007 si es asignacion masiva se podran los precios del presupuesto, por lo que no hay que sumar el iva al pvp

			if tipodoc = "ASIGNACION MASIVA" then
			    strselect=strselect & ",0 as iva "
			else
				strselect=strselect & ",a.iva "
			end if

			strselect=strselect & ",isnull(ter.codterminal,'''') as codterminal,d.abreviatura,d.ndecimales,"
			strselect=strselect & " round(a.importe,d.ndecimales) as importe, "
			strselect=strselect & " tall.descripcion as talla,col.descripcion as color, "

			if formato_impresion="listado_codigo_barras10.asp" or formato_impresion="listado_codigo_barras25.asp" then
				strselect=strselect & " case when a.ref_padre is not null then (select referencia from articulos where referencia=a.ref_padre) else a.referencia end as referencia,case when a.ref_padre is not null then (select nombre from articulos where referencia=a.ref_padre) else a.nombre end as nombre,  "
			else
				strselect=strselect & "a.referencia,a.nombre,"
			end if
''response.write("los datos son-" & tarifa1 & "-" & tarifaiva1 & "-" & opcprec1 & "-<br>")
            filtrotarifa1=""
            if tarifa1 & "">"" then
                filtrotarifa1=tarifa1
            else
                if tarifaiva1 & "">"" then
                    filtrotarifa1=tarifaiva1
                else
                    if opcprec1="coste" then
                        filtrotarifa1=""
                    else
                        if opcprec1="none" then
                            filtrotarifa1=""
                        else
                            filtrotarifa1=""
                        end if
                    end if
                end if
            end if
''response.write("los datos son-" & tarifa2 & "-" & tarifaiva2 & "-" & opcprec2 & "-<br>")
            filtrotarifa2=""
            if tarifa2 & "">"" then
                filtrotarifa2=tarifa2
            else
                if tarifaiva2 & "">"" then
                    filtrotarifa2=tarifaiva2
                else
                    if opcprec1="coste" then
                        filtrotarifa2=""
                    else
                        if opcprec1="none" then
                            filtrotarifa2=""
                        else
                            filtrotarifa2=""
                        end if
                    end if
                end if
            end if
            if tipodoc = "ASIGNACION MASIVA" and tarifaiva1 & "">"" then
                strselect=strselect & " isnull((select case when precios.es_dto=0 then round(pvpdto,d.ndecimales) else case when precios.es_dto=1 then round(a.pvp + ((a.pvp*pvpdto)/100),d.ndecimales) else round(a.importe + ((a.importe*pvpdto)/100),d.ndecimales) end end from precios with(nolock) where referencia=a.referencia and tarifa=''" & filtrotarifa1 & "'' and rango=''" & session("ncliente") & "BASE'' and temporada=''" & session("ncliente") & "BASE''),0) *(1+a.iva/100) as pvptarifa1, "
            else
		  	    strselect=strselect & " isnull((select case when precios.es_dto=0 then round(pvpdto,d.ndecimales) else case when precios.es_dto=1 then round(a.pvp + ((a.pvp*pvpdto)/100),d.ndecimales) else round(a.importe + ((a.importe*pvpdto)/100),d.ndecimales) end end from precios with(nolock) where referencia=a.referencia and tarifa=''" & filtrotarifa1 & "'' and rango=''" & session("ncliente") & "BASE'' and temporada=''" & session("ncliente") & "BASE''),0) as pvptarifa1, "
            end if
            if tipodoc = "ASIGNACION MASIVA" and tarifaiva2 & "">"" then
			    strselect=strselect & " (isnull((select case when precios.es_dto=0 then round(pvpdto,d.ndecimales) else case when precios.es_dto=1 then round(a.pvp + ((a.pvp*pvpdto)/100),d.ndecimales) else round(a.importe + ((a.importe*pvpdto)/100),d.ndecimales) end end from precios with(nolock) where referencia=a.referencia and tarifa=''" & filtrotarifa2 & "'' and rango=''" & session("ncliente") & "BASE'' and temporada=''" & session("ncliente") & "BASE''),0)) *(1+a.iva/100) as pvptarifa2 "
            else
                strselect=strselect & " isnull((select case when precios.es_dto=0 then round(pvpdto,d.ndecimales) else case when precios.es_dto=1 then round(a.pvp + ((a.pvp*pvpdto)/100),d.ndecimales) else round(a.importe + ((a.importe*pvpdto)/100),d.ndecimales) end end from precios with(nolock) where referencia=a.referencia and tarifa=''" & filtrotarifa2 & "'' and rango=''" & session("ncliente") & "BASE'' and temporada=''" & session("ncliente") & "BASE''),0) as pvptarifa2 "
            end if
''response.write("el strselect es-" & strselect & "-<br>")
''response.end

			if tipo_doc_tabla&"">"" and cant_doc="on" then
				strselect=strselect&" , abs(p.cantidad) as ''cantidad'' "
			else
				strselect= strselect&" , "&cantidad&" as ''cantidad''  "
			end if

			if tipodoc>"" or numdoc>"" then
				strselect= strselect&" , 1 as ''condicion1'' "
			else
				strselect= strselect&" , 0 as ''condicion1'' "
			end if
			if cant_doc="on" then
				strselect= strselect&" , 1 as ''condicion2'' "
			else
				strselect= strselect&" , 0 as ''condicion2'' "
			end if

			strselect= strselect+" , ''"&opcprec1&"'' as ''opcprec1'' "
			strselect= strselect+" , ''"&opcprec2&"'' as ''opcprec2'' "

			strselect= strselect+" , "&imprimir_listado_horizontal&" as ''imprimir_listado_horizontal'' "
			strselect= strselect+" , "&imprimir_listado_vertical&" as ''imprimir_listado_vertical'' "

			if ver_referencia="on" then
				strselect = strselect+", 1 as ver_referencia "
			else
				strselect = strselect+", 0 as ver_referencia "
			end if
            if ver_refProv="on" then
				strselect = strselect+", 1 as ver_refProv "
			else
				strselect = strselect+", 0 as ver_refProv "
			end if
			if ver_nombre="on" then
				strselect = strselect+", 1 as ver_nombre "
			else
				strselect = strselect+", 0 as ver_nombre "
			end if
			if ver_empresa="on" then
				strselect = strselect+", 1 as ver_empresa "
			else
				strselect = strselect+", 0 as ver_empresa "
			end if
			if ver_lineas="on" then
				strselect = strselect+", 1 as ver_lineas "
			else
				strselect = strselect+", 0 as ver_lineas "
			end if
			if ver_precios="on" then
				strselect = strselect+", 1 as ver_precios "
			else
				strselect = strselect+", 0 as ver_precios "
			end if  
			if ver_codterminal="on" then
				strselect = strselect+", 1 as ver_codterminal "
			else
				strselect = strselect+", 0 as ver_codterminal "
			end if

			'***RGU 12/5/2006 ***
			strselect = strselect+" ,a.medida, a.medidaventa, a.campo01"
			'***
			'**RGU 7/12/2006 nota:las fechas estan bien
			strselect = strselect+" , a.unidadarticulo, a.cantidadarticulo, med.valor as valorunidadarticulo "
			'***

			''ricardo 20-12-2006 se añade el codigo de la tarifa
			strselect = strselect & " ," & iif(tarifa>"","''" & tarifa & "''","NULL") & " as cod_tarifa_elegida "

			''ricardo 07-03-2007 se añade el codigo del almacen elegido
            ''belen manso 16-04-2018 el codigo del almacen lo obtenemos segun la tienda elegida
            ''stralmacen = "select almacen from tiendas where codigo like '" & tienda & "'"
			if tipodoc & "">"" then
			    if tipodoc="MOVIMIENTOS ENTRE ALMACENES" then
			        strselect = strselect & " ," & iif(almacen>"","''" & almacen & "''","(select mp.almdestino from movimientos as mp with(NOLOCK) where mp.nmovimiento like ''" & session("ncliente") & "%'' and mp.nmovimiento=p.nmovimiento)") & " as cod_almacen_elegido "
			    else
			        strselect = strselect & " ," & iif(almacen>"","''" & almacen & "''","p.almacen") & " as cod_almacen_elegido "
			    end if
			else
			    strselect = strselect & " ," & iif(almacen>"","''" & almacen & "''","null") & " as cod_almacen_elegido "
			end if


			'if nz_b(solopreciocambiado)<>0 and vatlcb="1" then
			'	strselect=strselect+", 1 as ''tratar'' "
			'elseif tarifa&"" > "" and vatlcb="1" then
			'	strselect=strselect+", 2 as ''tratar'', tarif.tarifa as tarifa "
			'else
			'	strselect=strselect+", 0 as ''tratar'' "
			'end if
			'**RGU19/2/2007
			if nz_b(solopreciocambiado)<>0 and vatlcb="1" then
				if tarifa&"">"" then
					strselect=strselect+", 2 as ''tratar'', tarif.tarifa as tarifa "
				else
					strselect=strselect+", 1 as ''tratar'' "
				end if
			else
				strselect=strselect+", 0 as ''tratar'' "
			end if
			'**rgu**
			''ricardo 3-4-2007 se añade el campo su_ref
			strselect=strselect+",(select top 1 prov.su_ref from proveer as prov with(NOLOCK) where prov.articulo like ''"&session("ncliente")&"%'' and prov.articulo=a.referencia) as su_ref"

			''ricardo 17-10-2007 se añaden el campo05 y el tdocumento,ndocumento,serie del documento y precio de temporada'
			strselect=strselect+",a.campo05"
			if tipodoc & "">"" then
			    strselect=strselect+",''" + tipodoc + "'' as tdocumento "
			else
			    strselect=strselect+",null as tdocumento "
			end if
			if numdoc & "">"" then
				if tipodoc = "ASIGNACION MASIVA" then
					'strselect=strselect+",''00000" + session("ncliente") + "'' as ndocumento "
					strselect=strselect+",''00000" + almacenmasivo + "'' as ndocumento "
				else
				    strselect=strselect+",''" + session("ncliente") + numdoc + "'' as ndocumento "
				end if
			else
			    strselect=strselect+",null as ndocumento "
			end if
			if seriedoc & "">"" then
			    strselect=strselect+",''" + seriedoc + "'' as serie_documento "
			else
			    strselect=strselect+",null as serie_documento "
			end if
		    strselect=strselect + ",round( "
			strselect=strselect + "     ( "
			strselect=strselect + "     isnull( "
            strselect=strselect + "             ( "
            strselect=strselect + "             (case when isnull(tt.es_dto,0)=0 then "
    		strselect=strselect + "                 convert(money,tt.pvpdto) "
			strselect=strselect + "              else case when isnull(tt.es_dto,0)=1 then "
			strselect=strselect + "                         convert(money,A.pvp)+((convert(money,A.pvp)*convert(money,tt.pvpdto))/convert(money,100)) "
            strselect=strselect + "                     else "
            strselect=strselect + "                         convert(money,A.importe)+((convert(money,A.importe)*convert(money,tt.pvpdto))/convert(money,100)) "
            strselect=strselect + "                     end "
            strselect=strselect + "             end "
            strselect=strselect + "             ) "
             if tipodoc = "ASIGNACION MASIVA" then
                strselect=strselect + "          *(1+a.iva/100) "
             end if                 
            strselect=strselect + "         ) "
            strselect=strselect + "     ,"
            if tipodoc = "ASIGNACION MASIVA" then
				strselect=strselect & "p.pvp "
			else
				strselect=strselect & "a.pvp "
			end if
            strselect=strselect + "       ))," & dec_prec & ") as pvp_temporada "

            ''fin ricardo 17-10-2007
	''FROM
            
			strfrom=" from articulos as a with(nolock) "
			strfrom=strfrom+ " left outer join articuloster as ter with(nolock) on ter.referencia=a.referencia and ter.referencia like ''"+session("ncliente")+"%'' "
			strfrom=strfrom+ " left outer join tallas as tall with(nolock) on tall.codigo=a.talla and tall.codigo like ''"+session("ncliente")+"%'' left outer join colores as col with(nolock) on col.codigo=a.color and col.codigo like ''"&session("ncliente")&"%'' "
			strfrom=strfrom+ " left outer join medidas as med with(nolock) on med.codigo like ''"+session("ncliente")+"%'' and med.descripcion=a.unidadarticulo "
			''ricardo 18-10-2007 se añade esto para obtener el precio de temporada
			strfrom=strfrom+ " left outer join ( "
			strfrom=strfrom+ "      select pre.es_dto,pre.pvpdto,pre.referencia "
			strfrom=strfrom+ "      from temporadas as tem with (NOLOCK),precios as pre with (NOLOCK) "
            fecha_hoy=day(date) & "/" & month(date) & "/" & year(date)
			strfrom=strfrom & "	    where tem.codigo like ''" & session("ncliente") & "%'' and tem.f_min<=''" & fecha_hoy & "'' and tem.f_max>=''" & fecha_hoy & "'' "
			strfrom=strfrom+ "      and pre.referencia like ''"+session("ncliente")+"%'' and pre.tarifa=''" + session("ncliente") + "BASE'' and pre.temporada=tem.codigo and pre.rango=''" + session("ncliente") + "BASE'' "
			strfrom=strfrom+ " ) as tt on tt.referencia=a.referencia "
            '--------------------------------------------------------------------
            if tipodoc = "ASIGNACION MASIVA" then
                if cod_temporada&""<>"" and filtro_temporada<>"blanco" then
                    if filtro_temporada="incluir" then
                        strfrom = strfrom + "INNER JOIN precios on precios.referencia = a.referencia"
                        strfrom = strfrom + " and precios.temporada =''"&cod_temporada&"''"
                        strfrom = strfrom + " and precios.rango=''" & session("ncliente") & "BASE'' "
                        strfrom = strfrom + " and precios.tarifa=''" & session("ncliente") & "BASE''"
                    end if
                    if filtro_temporada="excluir" then
                        excluir = 1
                    end if
                end if
            end if
            '--------------------------------------------------------------------
			strfrom= strfrom+", divisas as d with(nolock) "

			if tipo_doc_tabla&"">"" then
				strfrom =strfrom & ", " & tipo_doc_tabla & " as p with(nolock)"
		  	end if
			''ricardo 28-2-2007 si el documento es Asignacion masiva debera salir lo del presupuesto sin ningun filtro
			if tipodoc = "ASIGNACION MASIVA" then
			else
				if cod_temporada&"">"" or tarifa&"" > "" then
					strfrom=strfrom &", precios  with(nolock) "
				end if
				if tarifa&"" > "" then
					strfrom=strfrom&" ,articulos_tarifa tarif with(nolock) "
				end if
			end if


	''WHERE
			strwhere =" where a.cod_barras<>'''' and a.referencia like ''" & session("ncliente") & "%'' and d.codigo=a.divisa and d.codigo like ''"&session("ncliente")&"%'' and a.fbaja is null "


			'filtro por referencia del articulo
		    'if referencia > "" and lista="" then
			if referencia > ""  then
			 	strwhere = strwhere + " and substring(a.referencia,6,30) like ''%" + referencia + "%'' "
		    end if

			'filtro por el tipo de documento
			if tipodoc>"" or numdoc>"" then
			    if tipodoc="MOVIMIENTOS ENTRE ALMACENES" then
			        strwhere = strwhere + " and p.ref = a.referencia "
			    else
				    strwhere = strwhere + " and p.referencia = a.referencia and p.mainitem is null "
				end if

				if tipodoc = "ASIGNACION MASIVA" then
					'strwhere = strwhere + " and p." & num_doc_tabla & "=''00000" & session("ncliente") & "'' "
					'dgb 09/04/2008  anyado el filtro del almacen
					strwhere = strwhere + " and p." & num_doc_tabla & "=''00000" & almacenmasivo & "'' and p.almacen=''"& almacenmasivo&"'' "
				else
					strwhere = strwhere + " and p." & num_doc_tabla & "=''" & session("ncliente") & numdoc & "'' "
				end if
			end if

			'filtro por el nombre del articulo
		    if nombre > "" then
                nombre = replace(nombre, "'", "##")
				strwhere = strwhere + " and nombre like ''%" + nombre + "%'' "
		  	end if

			'filtro por la familia del articulo
		  	if familia > "" then
				strwhere = strwhere + " and familia in  "
				cadfam=familia
			else
				if familia_padre>"" then
					strwhere= strwhere & " and familia_padre in  "
					cadfam=familia_padre
				else
					if categoria> "" then
						strwhere = strwhere + " and categoria in  "
						cadfam=categoria
					end if
				end if
		  	end if
			if  familia>"" or familia_padre>"" or categoria>"" then
				if mid(cadfam,len(cadfam)-1,len(cadfam)) = ", " then
					cadfam=mid(cadfam,1,len(cadfam)-2)
				end if
				cadfam="''"&cadfam&"''"
				contcom=1
				while instr(contcom,cadfam,",")
					contcom=instr(contcom,cadfam,",")
					cadfam=mid(cadfam,1,contcom-1)&"'',''" & mid(cadfam,contcom+2,len(cadfam))
					contcom=contcom+5
				wend
				strwhere=strwhere + " (" &cadfam & ")"
			end if

			'filtro para sacar solo los articulos que hayan cambiado de precio desde la fecha
			'' solo si no hay temporada seleccionada
			if fechamodprec & "">"" and (tipodoc & ""="" and numdoc & ""="") and cod_temporada&""="" then
				strwhere = strwhere + " and fechamod>=''" + fechamodprec + "'' "
			end if

			'filtro solo los articulos que han sido cambiados de precio y no han sido imprimidos
			if nz_b(solopreciocambiado)<>0 then
				'**RGU 12/2/2007: Si se ha elegido tarifa no se aplica este filtro, se aplica el pendientetratar del articulos_tarifa
				if tarifa&"">"" then
				else
					strwhere = strwhere + " and pendienteimpresion<>0 "
				end if
				'**rgu
				'strselect= strselect&" , 1 as ''condicion3'' "
				strselect= strselect&" , 0 as ''condicion3'' "
			else
				strselect= strselect&" , 0 as ''condicion3'' "
			end if

			''ricardo 28-2-2007 si el documento es Asignacion masiva debera salir lo del presupuesto sin ningun filtro
            if tipodoc = "ASIGNACION MASIVA" then
                
			else
				if tarifa&"" > "" or cod_temporada&"">"" then
					strwhere= strwhere + " and precios.referencia=a.referencia and precios.referencia like ''"&session("ncliente")&"%'' "
				end if

				if tarifa&"" > "" then
					strtarifa="(''"
					strtarifa=strtarifa&replace(tarifa,", ","'',''")
					strtarifa=strtarifa&"'')"
					strwhere = strwhere + " and precios.tarifa in " & strtarifa & " and precios.rango=''" & session("ncliente") & "BASE'' "
					if cod_temporada&""="" then
						strwhere = strwhere + " and precios.temporada=''" & session("ncliente") & "BASE'' "
					end if
					strwhere = strwhere + " and tarif.referencia= a.referencia and tarif.referencia like ''"&session("ncliente")&"%'' "
					strwhere = strwhere + " and tarif.tarifa=precios.tarifa"

					''ricardo 29-11-2006
					if nz_b(solopreciocambiado)<>0 then
						strwhere = strwhere + " and tarif.pendientetratar=1 "
					end if
				end if

                'filtro por temporada
                if cod_temporada&""<>"" and filtro_temporada<>"blanco" then
                    if filtro_temporada="incluir" then
                        strwhere= strwhere + " and precios.temporada=''"&cod_temporada&"'' "
                    else
                        if filtro_temporada="excluir" then
                           excluir = 1
                        end if
                    end if

                    strwhere = strwhere + " and precios.rango=''" & session("ncliente") & "BASE'' "

                    if tarifa&""="" then
                        strwhere= strwhere + " and precios.tarifa=''" & session("ncliente") & "BASE''"
                    end if
                end if
            end if

			''ricardo 20-2-2007 se añade el stock mayor o igual a
			if stockmayoroigual & "">"" or almacen & "">"" then
				strwhere= strwhere + " and a.referencia in "
				strwhere= strwhere + " (select alma.articulo from almacenar as alma WITH(NOLOCK) "
				strwhere= strwhere + " where alma.articulo like ''" + session("ncliente") + "%'' and alma.articulo=a.referencia "
				if almacen & "">"" then
					strwhere= strwhere + " and alma.almacen=''" & almacen & "'' "
				end if
				if stockmayoroigual & "">"" then
					strwhere= strwhere + " group by alma.articulo "
					strwhere= strwhere + " having sum(alma.stock)>= " + replace(stockmayoroigual,",",".")
				end if
				strwhere= strwhere + " ) "
			end if

			''ricardo 26-4-2007 parametro para solamente trabajar con campo01=1
			if cstr(mcp1)="1" then
			    strwhere= strwhere + " and isnull(a.campo01,'''')=''1'' "
			end if

            ''pascual 29/03/2017 add familia
            strselect = strselect + ",a.familia as familia "  

			''esto se hace siempre que no haya que imprimir codigos de barra por no cumplir los parametros
			condicion=1
			if (tipodoc>"" and strwhere="") or devolver_error=1 then
				condicion=0 'no los cumple
			else
				condicion=1
			end if

''RGU

		'JMA 30/10/05: Quitamos de la condicion el formato "listado_codigo_barras.asp"'
		if  formato_impresion="listado_codigo_barras4.asp" then
			formato_impresion="..\\productos\\listados\\" & formato_impresion
		end if
		if formato_impresion="..\\..\\custom\\listado_codigo_barras5.asp" then
			formato_impresion="listado_codigo_barras5.asp"
		end if
		if formato_impresion="../custom/listado_codigo_barrasn.asp" then
			formato_impresion="listado_codigo_barrasn.asp"
		end if
''rgu

			set conn = Server.CreateObject("ADODB.Connection")
			conn.cursorlocation=3
			donde=inStr(1,formato_impresion,"zebra",1)
			if donde=1 then
			    conn.open session("dsn_cliente")
			else
			    conn.open session("backendlistados")
			end if
            
''response.Write(strSelect)
''response.Write(strfrom)
''response.Write(strWhere)
''response.end

            if excluir = 1 then
                llamada="EXEC CodigosBarras @strSelect='"&strSelect&"', @strFrom='"&strfrom&"',@strWhere='"&strWhere&"',@order='"&ordenar&"', @Nomtabla ='["& session("usuario") &"]', @condicion='"&condicion&"', @temporada='"&cod_temporada&"'"
             else
                llamada="EXEC CodigosBarras @strSelect='"&strSelect&"', @strFrom='"&strfrom&"',@strWhere='"&strWhere&"',@order='"&ordenar&"', @Nomtabla ='["& session("usuario") &"]', @condicion='"&condicion&"'"    
             end if

            set result=conn.execute(replace(llamada, "##","''"))
            deleteTablaCalculoPrecio.open "delete  from sumaPrecioLocal_etiquetasMoney", session("backendlistados")
			if result(0) then%>
				<script language="javascript" type="text/javascript">
				    document.codigo_barras.action = "<%=formato_impresion%>";
				    document.codigo_barras.submit();
				</script><%
			else
				%><script language="javascript" type="text/javascript">window.Alert("<%=LitError%>");</script><%
			end if

			result.close
			conn.close

           
''RGU
	end if

	%></form><%

		set rstAux = Nothing
		set rst = Nothing
        set rstPrecioSinIva = Nothing
        set rstPrecioTarifaLocal = Nothing
        set insertTablaCalculoPrecio = Nothing
        set pruebaSelect=Nothing
		set rstSelect = Nothing
		set rstpro = Nothing
        set conn =Nothing
        set deleteTablaCalculoPrecio=Nothing

end if%>
</body>
</html>