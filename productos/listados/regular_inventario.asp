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
''ricardo 16-11-2007 se cambia la dsn desde dsncliente a backendlistados
%>
<%
' AUTOR :CAG
'----------------------------------------------------------------------------------------------%>

<% Server.ScriptTimeout = 400 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloRegInv%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>

<LINK REL="styleSHEET" href="../../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../../impresora.css" MEDIA="PRINT">
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../modulos.inc" -->
<!--#include file="../reg_inventario.inc" -->
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../../catFamSubResponsive.inc" -->
<!--#include file="../../styles/formularios.css.inc" -->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->

<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
    //ricardo 20-4-2007 se añade esta funcion
    function cambiarDesv(boton_radio) {
        if (boton_radio == '2') {
            document.getElementById("mostSolDesv1").style.display = "";
            //document.getElementById("mostSolDesv2").style.display = "";
            document.regular_inventario.verSolDesv.checked = false;
            //dgb 31-03-2008 anyadir Ordenar por
            document.getElementById("ordenar1").style.display = "";
            //document.getElementById("ordenar2").style.display = "";
        }
        else {
            document.getElementById("mostSolDesv1").style.display = "none";
            //document.getElementById("mostSolDesv2").style.display = "none";
            document.regular_inventario.verSolDesv.checked = false;
            //dgb 31-03-2008 anyadir Ordenar por
            document.getElementById("ordenar1").style.display = "none";
            //document.getElementById("ordenar2").style.display = "none";
        }
    }

    function keypress2() {
        tecla = window.event.keyCode;
        //keyPressed(tecla);
    }

    function tier2Menu(objMenu) {

        if (objMenu == "albaranes") {
            serFact.style.display = "none";
            serAlb.style.display = "";
        }
        if (objMenu == "facturas") {
            serFact.style.display = "";
            serAlb.style.display = "none";
        }
    }

    function keyPressed(tecla) {
    }

    //Desencadena la búsqueda del proveedor cuya referencia se indica
    function TraerProveedor(mode, tipo) {
        document.ventas_tienda.action = "ventas_tienda.asp?nproveedor=" + document.ventas_tienda.nproveedor.value + "&mode=" + mode;
        document.ventas_tienda.submit();
    }
</script>

<body class="BODY_ASP">
<iframe name="frameExportar" style='display:none;' src="regular_inventario_pdf.asp?mode=ver" frameborder='0' width='500' height='200'></iframe>
<%
'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
%>
<form name="regular_inventario" method="post">
<%
		PintarCabecera "regular_inventario.asp"
		'Leer parámetros de la página
		SumaAnt = 0
		SubTotalAnt = 0
  		mode=EncodeForHtml(Request.QueryString("mode"))
  		campo=limpiaCadena(Request.QueryString("campo"))
  		criterio=limpiaCadena(Request.QueryString("criterio"))
  		texto=limpiaCadena(Request.QueryString("texto"))
		elTotal = limpiaCadena(Request.form("elTotal"))

	dim vcri  ' cag Parametro que trae lista de almacenes para desplegable en pantalla del historial dependiendo del usuario

	ObtenerParametros "listado_regular_inventario.asp"

	%><input type="hidden" name="vcri" value="<%=EncodeForHtml(vcri)%>"><%

	documento	= limpiaCadena(Request.QueryString("documento"))
	if documento="" then
		documento	= limpiaCadena(Request.form("documento"))
	end if

	if enc.EncodeForJavascript(request.querystring("familia")) >"" then
	   familia = limpiaCadena(request.querystring("familia"))
	else
	   familia = limpiaCadena(request.form("familia"))
	end if

	si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
	si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)

	fdesde		= limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde	= limpiaCadena(Request.form("fdesde"))
	end if

	fhasta		= limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta	= limpiaCadena(Request.form("fhasta"))
	end if

	if enc.EncodeForJavascript(request.querystring("almacen")) >"" then
	   almacen = limpiaCadena(request.querystring("almacen"))
	else
	   almacen = limpiaCadena(request.form("almacen"))
	end if

	comercial	= limpiaCadena(Request.QueryString("comercial"))
	if comercial ="" then
		comercial	= limpiaCadena(Request.form("comercial"))
	end if

	if enc.EncodeForJavascript(request.querystring("numDocum")) >"" then
	   numDocum = limpiaCadena(request.querystring("numDocum"))
	else
	   numDocum = limpiaCadena(request.form("numDocum"))
	end if

	if enc.EncodeForJavascript(request.querystring("referencia")) >"" then
	   referencia = limpiaCadena(request.querystring("referencia"))
	else
	   referencia = limpiaCadena(request.form("referencia"))
	end if

	if enc.EncodeForJavascript(request.querystring("nombre")) >"" then
	   nombre = limpiaCadena(request.querystring("nombre"))
	else
	   nombre = limpiaCadena(request.form("nombre"))
	end if

	categoria	= limpiaCadena(Request.QueryString("categoria"))
	if categoria ="" then
		categoria	= limpiaCadena(Request.form("categoria"))
	end if

	familia_padre	= limpiaCadena(Request.QueryString("familia_padre"))
	if familia_padre ="" then
		familia_padre	= limpiaCadena(Request.form("familia_padre"))
	end if

	familia	= limpiaCadena(Request.QueryString("familia"))
	if familia ="" then
		familia	= limpiaCadena(Request.form("familia"))
	end if

	agrupar	= limpiaCadena(Request.QueryString("agrupar"))
	if agrupar ="" then
		agrupar	= limpiaCadena(Request.form("agrupar"))
	end if

	detalle	= limpiaCadena(Request.QueryString("detalle"))
	if detalle ="" then
		detalle	= limpiaCadena(Request.form("detalle"))
	end if

	verCostes	= limpiaCadena(Request.QueryString("verCostes"))
	if verCostes ="" then
		verCostes	= limpiaCadena(Request.form("verCostes"))
	end if

    ''ricardo 20-4-2007
	verSolDesv	= limpiaCadena(Request.QueryString("verSolDesv"))
	if verSolDesv ="" then
		verSolDesv	= limpiaCadena(Request.form("verSolDesv"))
	end if
	
	''dgb 31-03-2008
	ordenar=limpiaCadena(Request.QueryString("ordenar"))
	if ordenar ="" then
		ordenar	= limpiaCadena(Request.form("ordenar"))
	end if

	WaitBoxOculto LitEsperePorFavor

  set rstAux = Server.CreateObject("ADODB.Recordset")
  set rst = Server.CreateObject("ADODB.Recordset")
  set rst2 = Server.CreateObject("ADODB.Recordset")
  set rstSelect = Server.CreateObject("ADODB.Recordset")
  set rstTablas = Server.CreateObject("ADODB.Recordset")

	if mode="browse" then%>
		<table width='100%'>
		   	<tr>
				<td width="30%" align="left" >
				  	<font class=CELDAC7>&nbsp;(<%=LitEmitido%>&nbsp; <%=day(date)%>/<%=month(date)%>/<%=year(date)%>)</font>
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
	Alarma "regular_inventario.asp"

	if (mode="select1") then%>
		<!--<table border='<%=borde%>'>--><br/>
            <%TmpAlmacenDef= ""
			'Poner por defecto el almacen correspondiente a la tienda donde esta el tpv y caja del fichero cetel.tpv'
			linea1=session("f_tpv")
			linea2=session("f_caja")
			linea3=session("f_empr")
			'Obtenemos la tienda y el almacen correspondiente.
			if linea1<>"" and linea2<>"" and linea3<>"" then
				strSelect = "select c.almacen from tpv a, cajas b, tiendas c where a.caja=b.codigo and b.tienda=c.codigo and tpv='" & linea1 & "' and b.codigo='" & linea2 & "'"
				rstAux.cursorlocation=3
				rstAux.open strSelect,session("backendlistados")

				if linea3=session("ncliente") then
					if rstAux.eof then
						TmpAlmacenDef = ""
					else
						TmpAlmacenDef = rstAux("almacen")
					end if
				else
					TmpAlmacenDef = ""
				end if
				rstAux.close
			end if

                DrawDiv "1","",""
                DrawLabel "","",LitDesdeFecha
                DrawInput "", "", "fdesde",EncodeForHtml(iif(fdesde>"",fdesde,"01/01/" & year(date))), "" 
                DrawCalendar "fdesde"
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitHastaFecha
                DrawInput "", "", "fhasta",EncodeForHtml(iif(fhasta>"",fhasta,day(date) & "/" & month(date) & "/" & year(date))), "" 
                DrawCalendar "fhasta"
                CloseDiv
			
                TmpAlmacen=""
                rstSelect.cursorlocation=3
				rstSelect.open "select codigo, descripcion from almacenes with(nolock) where tienda is null and codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
				DrawSelectMultipleCelda "width60","","",0,LitAlmacen,"almacen",rstSelect,TmpAlmacen,"codigo","descripcion","",""
				rstSelect.close
				
				rstSelect.cursorlocation=3
				rstSelect.open "SELECT p.dni, p.nombre FROM PERSONAL AS p with(nolock) WHERE p.dni like '" & session("ncliente") & "%' order by p.nombre", session("backendlistados")
				DrawSelectMultipleCelda "width60","","",0,LitResponsable,"comercial",rstSelect,comercial,"dni","nombre","",""
				rstSelect.close
			%><br/>
			
			<%
                EligeCelda "input","add","left","","",0,LitNumDocum,"numDocum",0,EncodeForHtml(numDocum)
			
                EligeCelda "input","add","left","","",0,LitConref,"referencia",0,EncodeForHtml(referencia)
				
                EligeCelda "input","add","left","","",0,LitConNombre,"nombre",0,EncodeForHtml(nombre)
			
					dim ConfigDespleg (3,13)

					i=0
					ConfigDespleg(i,0)="categoria"
					ConfigDespleg(i,1)=""
					ConfigDespleg(i,2)="5"
					ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
					ConfigDespleg(i,4)=1
					ConfigDespleg(i,5)="width60"
					ConfigDespleg(i,6)="MULTIPLE"
					ConfigDespleg(i,7)="codigo"
					ConfigDespleg(i,8)="nombre"
					ConfigDespleg(i,9)=LitCategoria
					ConfigDespleg(i,10)=categoria
					ConfigDespleg(i,11)=""
					ConfigDespleg(i,12)=""

					i=1
					ConfigDespleg(i,0)="familia_padre"
					ConfigDespleg(i,1)=""
					ConfigDespleg(i,2)="5"
					ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
					ConfigDespleg(i,4)=1
					ConfigDespleg(i,5)="width60"
					ConfigDespleg(i,6)="MULTIPLE"
					ConfigDespleg(i,7)="codigo"
					ConfigDespleg(i,8)="nombre"
					ConfigDespleg(i,9)=LitFamilia
					ConfigDespleg(i,10)=familia_padre
					ConfigDespleg(i,11)=""
					ConfigDespleg(i,12)=""

					i=2
					ConfigDespleg(i,0)="familia"
					ConfigDespleg(i,1)=""
					ConfigDespleg(i,2)="5"
					ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
					ConfigDespleg(i,4)=1
					ConfigDespleg(i,5)="width60"
					ConfigDespleg(i,6)="MULTIPLE"
					ConfigDespleg(i,7)="codigo"
					ConfigDespleg(i,8)="nombre"
					ConfigDespleg(i,9)=LitSubFamilia
					ConfigDespleg(i,10)=familia
					ConfigDespleg(i,11)=""
					ConfigDespleg(i,12)=""
					DibujaDesplegables ConfigDespleg,session("backendlistados")%><br/>
     <%			
             DrawDiv "1", "", ""
             DrawLabel "", "", LitAgrupar%><select class="width60" name="agrupar"> 		
                        <option <%=iif(agrupar="agr_Documento" or agrupar>"","selected","")%> value="agr_Documento"><%=LitDocumento%></option>
						<option <%=iif(agrupar="agr_Articulo","selected","")%> value="agr_Articulo"><%=LitArticulo%></option>
						<option <%=iif(agrupar="agr_SubFamilia","selected","")%> value="agr_SubFamilia"><%=LITSUBFAMILIA%></option>
					</select>					
			<%CloseDiv
             DrawDiv "4","",""
             DrawLabel "","", LitSinDetalle%><input  type="radio" name="detalle" value="sin" checked  onmouseup="cambiarDesv('1')"><%CloseDiv
				if cstr(vcri)<>"1" then%>
					<input type="hidden" name='verCostes' value=""><%
				else				
					
                EligeCelda "check","add","","","",0,LitMostCostes,"verCostes",0, ""
				end if
                         
             DrawDiv "4","",""
             DrawLabel "","", LitConDetalle%><input type="radio" name="detalle" value="con"  onmouseup="cambiarDesv('2')"><%	   
			 CloseDiv''ricardo 20-4-2007 se pone el filtro de mostrar solo desviaciones%>
            <br />
            <span id="mostSolDesv1" style="display: none"><%
             EligeCelda "check","add","","","",0,LitMostSolDesv,"verSolDesv",0, "" 
				'dgb 31-03-2008  anyadir Ordenar por 
			%></span><%
                DrawDiv "1","display:none","ordenar1"
                  DrawLabel "","",LitOrdenado  
                    %><select class="width60" name="ordenar" >                   
                        <option value="descripcion" selected ><%=LitDescripcion%></option>
                        <option value="referencia" ><%=LitReferencia%></option>
                    </select><%CloseDiv

'****************************************************************************************************************
		'Mostrar el listado.
	elseif mode="browse" then%>
		    <input type="hidden" name="fdesde" value="<%=EncodeForHtml(fdesde)%>">
		    <input type="hidden" name="fhasta" value="<%=EncodeForHtml(fhasta)%>">
		    <input type="hidden" name="almacen" value="<%=EncodeForHtml(almacen)%>">
			<input type="hidden" name="comercial" value="<%=EncodeForHtml(comercial)%>">
	    	<input type="hidden" name="numDocum" value="<%=EncodeForHtml(numDocum)%>">
	    	<input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>">
	    	<input type="hidden" name="categoria" value="<%=EncodeForHtml(categoria)%>">
	    	<input type="hidden" name="familia_padre" value="<%=EncodeForHtml(familia_padre)%>">
	    	<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>">
	    	<input type="hidden" name="nombre" value="<%=EncodeForHtml(nombre)%>">
			<input type="hidden" name="agrupar" value="<%=EncodeForHtml(agrupar)%>">
	    	<input type="hidden" name="detalle" value="<%=EncodeForHtml(detalle)%>">
			<input type="hidden" name="verCostes" value="<%=EncodeForHtml(verCostes)%>">
			<input type="hidden" name="verSolDesv" value="<%=EncodeForHtml(verSolDesv)%>">
			<input type="hidden" name="ordenar" value="<%=EncodeForHtml(ordenar)%>">

<%
			rst.cursorlocation=3
			rst.open "select codigo,factcambio,ndecimales,abreviatura from divisas with(NOLOCK) where moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("backendlistados")
			if not rst.eof then
			    MB=rst("codigo")
			    factcambio_MB=rst("factcambio")
			    n_decimales=rst("ndecimales")
			    ndecimales_MB=n_decimales
			    MB_abrev=rst("abreviatura")
			end if
			rst.close
			
			''MB=d_lookup("codigo", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("backendlistados"))
			''n_decimales = null_z(d_lookup("ndecimales", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("backendlistados")))
			''n_decimalesMB = n_decimales
			''MB_abrev = d_lookup("abreviatura", "divisas", "codigo='" & MB & "' and codigo like '" & session("ncliente") & "%'", session("backendlistados"))
			MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='453'", DSNIlion)
			MAXPDF=d_lookup("maxpdf", "limites_listados", "item='453'", DSNIlion)
			''DEC_CANT=conseguir_dec(session("backendlistados"),session("ncliente"),"cantidad")
			''DEC_PREC=conseguir_dec(session("backendlistados"),session("ncliente"),"precios")
			%><input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>
<%
			VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarArticulos)=1
			CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina                                     

			if fdesde>"" then
				%><font class=ENCABEZADO><b><%=LitDesdeFecha%> : </b></font><font class='CELDA'><%=EncodeForHtml(fdesde)%></b></font><br/><%
			end if
			if fhasta>"" then
				%><font class=ENCABEZADO><b><%=LitHastaFecha%> : </b></font><font class='CELDA'><%=EncodeForHtml(fhasta)%></b></font><br/><%
			end if
			if almacen<>"" then 'Se selecciono uno o varios almacenes
				%><font class=ENCABEZADO><b><%=LitAlmacen%> : </b></font><font class='CELDA'><%=EncodeForHtml(NombresEntidades(almacen,"almacenes","codigo","descripcion",session("backendlistados")))%></b></font><br/><%
			end if
			if comercial<>"" and si_tiene_modulo_comercial<>0 then
				%> <font class=ENCABEZADO><b> <%=LITRESPONSABLE%>  : </b></font><font class='CELDA'> <%=EncodeForHtml(NombresEntidades(comercial,"personal","dni","nombre",session("backendlistados")))%> </b></font><br/> <%
			end if
			if comercial<>"" and si_tiene_modulo_comercial=0 then 'Se selecciono uno o varios comerciales
				%> <font class=ENCABEZADO><b>  <%=LITRESPONSABLE%>  : </b></font><font class='CELDA'>  <%=EncodeForHtml(NombresEntidades(comercial,"personal","dni","nombre",session("backendlistados")))%> </b></font><br/> <%
			end if

			if numDocum>"" then
					%><font class=ENCABEZADO><b><%=LitNumDocum%> : </b></font><font class='CELDA'><%=EncodeForHtml(numDocum)%></b></font><br/><%
			end if
			if referencia>"" then
					%><font class=ENCABEZADO><b><%=LitConref%> : </b></font><font class='CELDA'><%=EncodeForHtml(referencia)%></b></font><br/><%
			end if
			if nombre>"" then
					%><font class=ENCABEZADO><b><%=LitConNombre%> : </b></font><font class='CELDA'><%=EncodeForHtml(nombre)%></b></font><br/><%
			end if

			'Mostrar nombre de familia, familia_padre ó gama (categoria)                                                                                
			if familia<>"" then
				desc_familia=NombresEntidades(familia,"familias","codigo","nombre",session("backendlistados"))
				%><font class=ENCABEZADO><b><%=LitSubFamilia%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia)%></b></font><br/><%
			elseif familia_padre<>"" then
				desc_familia_padre=NombresEntidades(familia_padre,"familias_padre","codigo","nombre",session("backendlistados"))
				%><font class=ENCABEZADO><b><%=LitFamilia%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_familia_padre)%></b></font><br/><%
			elseif categoria<>"" then
				desc_categoria=NombresEntidades(categoria,"categorias","codigo","nombre",session("backendlistados"))
				%><font class=ENCABEZADO><b><%=LitCategoria%> : </b></font><font class='CELDA'><%=EncodeForHtml(desc_categoria)%></b></font><br/><%
			end if

			if agrupar="agr_Documento" then
					%><font class=ENCABEZADO><b><%=LitAgrupar%> : </b></font><font class='CELDA'><%=LitDocumento%></b></font><%
					if detalle="sin" then%>
						<font class='CELDA'><%=LitSinDetalle%></b></font><%
					elseif detalle="con" then%>
						<font class='CELDA'><%=LitConDetalle%></b></font><%
					'dgb 31-03-2008
					%>
					<br /><font class=ENCABEZADO><b><%=LitOrdenado%> : </b></font>
					<%
					if ordenar="descripcion" then%>
					    <font class='CELDA'><%=LitDescripcion%></font>
					<%elseif ordenar="referencia" then%>
					    <font class='CELDA'><%=LitReferencia%></font>
					<%end if
					end if
			elseif agrupar="agr_Articulo" then
					%><font class=ENCABEZADO><b><%=LitAgrupar%> : </b></font><font class='CELDA'><%=LitArticulo%></b></font><%
					if detalle="sin" then%>
						<font class='CELDA'><%=LitSinDetalle%></b></font><%
					elseif detalle="con" then%>
						<font class='CELDA'><%=LitConDetalle%></b></font><%
						'dgb 31-03-2008
					    %>
					    <br /><font class=ENCABEZADO><b><%=LitOrdenado%> : </b></font>
					    <%
					    if ordenar="descripcion" then%>
					        <font class='CELDA'><%=LitDescripcion%></font>
					    <%elseif ordenar="referencia" then%>
					     <font class='CELDA'><%=LitReferencia%></font>
					    <%end if
					end if
			elseif agrupar="agr_SubFamilia" then
					%><font class=ENCABEZADO><b><%=LitAgrupar%> : </b></font><font class='CELDA'><%=LitSubFamiliaRegInv%></b></font><%
					if detalle="sin" then%>
						<font class='CELDA'><%=LitSinDetalle%></b></font><%
					elseif detalle="con" then%>
						<font class='CELDA'><%=LitConDetalle%></b></font><%
						'dgb 31-03-2008
					    %>
					    <br /><font class=ENCABEZADO><b><%=LitOrdenado%> : </b></font>
					    <%
					    if ordenar="descripcion" then%>
					      <font class='CELDA'><%=LitDescripcion%></font>
					    <%elseif ordenar="referencia" then%>
					       <font class='CELDA'><%=LitReferencia%></font>
					    <%end if
					end if
			end if
			%><br/><%

			if verCostes>"" then
					%><font class=ENCABEZADO><b><%=LitMostCostes%> </b></font><br/><%
			end if

            ''ricardo 20-4-2007
			if verSolDesv & "">"" then
			    %><font class=ENCABEZADO><b><%=LitMostSolDesv%></font><br/><%
			end if

''''''''''''
'gestion de la almacen y del comercial (PASO DE PARAMETROS A PROCEDIMIENTO)
''''''''''''
			if almacen>"" then 'Se selecciono almacen
				if instr(almacen,",")>0 then
					listaAlmacen="(''" & replace(replace(almacen," ",""),",","'',''") & "'')"
				else
					listaAlmacen="(''" & replace(replace(almacen," ",""),",","'',''") & "'')"
				end if
			end if

			if comercial>"" then 'Se selecciono comercial
				if instr(comercial,",")>0 then
					listaComercial="(''" & replace(replace(comercial," ",""),",","'',''") & "'')"
				else
					listaComercial="(''" & replace(replace(comercial," ",""),",","'',''") & "'')"
				end if
			end if

			if categoria>"" then 'Se selecciono categoria
				if instr(categoria,",")>0 then
					listaCategoria="(''" & replace(replace(categoria," ",""),",","'',''") & "'')"
				else
					listaCategoria="(''" & replace(replace(categoria," ",""),",","'',''") & "'')"
				end if
			end if
			if familia_padre>"" then 'Se selecciono familia_padre
				if instr(familia_padre,",")>0 then
					listaFamilia_padre="(''" & replace(replace(familia_padre," ",""),",","'',''") & "'')"
				else
					listaFamilia_padre="(''" & replace(replace(familia_padre," ",""),",","'',''") & "'')"
				end if
			end if
			if familia>"" then 'Se selecciono familia
				if instr(comercial,",")>0 then
					listaFamilia="(''" & replace(replace(familia," ",""),",","'',''") & "'')"
				else
					listaFamilia="(''" & replace(replace(familia," ",""),",","'',''") & "'')"
				end if
			end if
			 %><hr/> <%

		if verCostes>"" then
			verCost=1
		else
			verCost=0
		end if

		''ricardo 20-4-2007
		if verSolDesv>"" then
			verSolDesv=1
		else
			verSolDesv=0
		end if

		strQuery="Exec listadoRegularizacionInventario "
		strQuery=strQuery & "@fdesde='" & fdesde & "',"
		'strQuery=strQuery & "@fhasta='" & fhasta & " 23:59:59',"
		strQuery=strQuery & "@fhasta='" & fhasta & "',"
		strQuery=strQuery & "@listaAlmacen='" & listaAlmacen & "',"
		strQuery=strQuery & "@comercial='" & listaComercial & "',"
		strQuery=strQuery & "@numDocumento='" & numDocum & "',"
		strQuery=strQuery & "@conRef='" & referencia & "',"
		strQuery=strQuery & "@conNombre='" & nombre & "',"
		strQuery=strQuery & "@categoria='" & listaCategoria & "',"
		strQuery=strQuery & "@familia='" & listaFamilia_padre & "',"
		strQuery=strQuery & "@subfamilia='" & listaFamilia & "',"
		strQuery=strQuery & "@agrupar='" & agrupar & "',"
		strQuery=strQuery & "@verdetalle='" & detalle & "',"
		strQuery=strQuery & "@mostrarCoste=" & verCost & ","
		strQuery=strQuery & "@verSolDesv=" & verSolDesv & ","
		strQuery=strQuery & "@usuario='" & session("usuario") & "',"
		strQuery=strQuery & "@sesion_ncliente='" & session("ncliente") & "'"

		lote=limpiaCadena(Request.QueryString("lote"))

		if lote="" then
			set conVentas = Server.CreateObject("ADODB.Connection")
			conVentas.open session("backendlistados")
			conVentas.execute(strQuery)
			conVentas.close
			set conVentas=nothing
		end if
		strSelect="select * from [" & session("usuario") & "]"
		if (agrupar="agr_Articulo" or agrupar="agr_Documento") and detalle="con" then
		    'dgb
		    if ordenar="descricpion" then
		        strSelect=strSelect+" order by documento,nombreArticulo"
		    elseif ordenar="referencia" then
		        strSelect=strSelect+" order by documento,referencia"
		    end if
		else 
            strSelect=strSelect+" order by num"
        end if
		rstAux.cursorlocation=3
		rstAux.open strSelect ,session("backendlistados")
		NUMREGISTROS=rstAux.recordcount

  	  if rstAux.EOF then
			rstAux.Close
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitMsgDatosNoExiste%>");
			      parent.window.frames["botones"].document.location = "regular_inventario_bt.asp?mode=select1";
			      document.regular_inventario.action = "regular_inventario.asp?mode=select1";
			      document.regular_inventario.submit();
			</script>
      <%else%>
		<input type="hidden" name="NumRegsTotal" value="<%=EncodeForHtml(NUMREGISTROS)%>">
        <%if lote="" then
				lote=1
		else
			lote = clng(lote)
		end if
		sentido=limpiaCadena(Request.QueryString("sentido"))
		lotes=NUMREGISTROS/MAXPAGINA
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

		rstAux.PageSize=MAXPAGINA
		rstAux.AbsolutePage=lote

		NavPaginas lote,lotes,campo,criterio,texto,1 %>

		<table  borde="0" width="100%" style="border-collapse: collapse;"><%
		'Dibujar cabecera de listado
    	DrawFila color_fondo
		if agrupar="agr_Articulo" and detalle="sin"  then
				DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", true, LitArticulo
		end if
		if agrupar="agr_SubFamilia" and detalle="sin"  then
				DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", true, LitSubFamiliaRegInv
		end if
		if (agrupar="agr_Documento" and detalle="sin") or  (agrupar="agr_Articulo" and detalle="sin") or  (agrupar="agr_SubFamilia" and detalle="sin") then
					DrawCelda2 "DATO ALIGN=LEFT ", "left", true, LitDocumento
					DrawCelda2 "DATO ALIGN=LEFT ", "left", true, LitFecha
					DrawCelda2 "DATO ALIGN=LEFT ", "left", true, LitResponsable
					DrawCelda2 "DATO ALIGN=LEFT ", "left", true, LitObservaciones

					if verCost=1 and agrupar="agr_Documento" then
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", true, LitValStockNue
					end if
					if verCost=1 and (agrupar="agr_Articulo" or agrupar="agr_SubFamilia")  then
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", true, LitCoste
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", true, LitValStockNue
					end if
					if verCost=1 then
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", true, LitValStockDesv
					end if
			    CloseFila
		end if
		'elseif agrupar="agr_Documento" and detalle="con" then
		if agrupar="agr_Documento" and detalle="con" then
		    	DrawFila color_fondo
				    %>	<td class="TDBORDECELDAB8" rowspan="2"><%=LitDocumento%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitFecha%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitResponsable%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitArticulo%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="3" align="center"><%=LitStock%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitMin%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitRep%></td><%
						if verCost=1 then%>
							<td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitCoste%></td>
						    <td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitValStockNue%></td>
						    <td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitValStockDesv%></td>
						<%
						end if%>
					</tr>
					<tr bgcolor=<%=color_fondo%>>
					    <td class="TDBORDECELDAB8" align="center"><%=LitReferencia%></td>
					    <td class="TDBORDECELDAB8" align="center"><%=LitDescripcion%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td><td class=tdbordeCELDAB8 align="center"><%=LitStockDesv%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td>
					</tr><%
			    CloseFila
		end if
		if agrupar="agr_Articulo" and detalle="con" then
		    	DrawFila color_fondo
				    %>	<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitArticulo%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitDocumento%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitFecha%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitResponsable%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="3" align="center"><%=LitStock%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitMin%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitRep%></td><%
						if verCost=1 then%>
						<td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitCoste%></td>
						<td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitValStockNue%></td>
						<td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitValStockDesv%></td>
						<%
						end if%>
					</tr>
					<tr bgcolor=<%=color_fondo%>>
					    <td class="TDBORDECELDAB8" align="center"><%=LitReferencia%></td>
					    <td class="TDBORDECELDAB8" align="center"><%=LitDescripcion%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td><td class=tdbordeCELDAB8 align="center"><%=LitStockDesv%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td>
					</tr><%
			    CloseFila
		end if
		if agrupar="agr_SubFamilia" and detalle="con" then
		    	DrawFila color_fondo
				    %>	<td class="TDBORDECELDAB8" rowspan="2"><%=LitSubFamiliaRegInv%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitDocumento%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitFecha%></td>
						<td class="TDBORDECELDAB8" rowspan="2"><%=LitResponsable%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="3" align="center"><%=LitStock%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitMin%></td>
						<td class="TDBORDECELDAB8" rowspan="1" colspan="2" align="center"><%=LitRep%></td><%
						if verCost=1 then%>
						<td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitCoste%></td>
						<td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitValStockNue%></td>
						<td class="TDBORDECELDAB8" align="right" rowspan="2"><%=LitValStockDesv%></td>
						<%
						end if%>
					</tr>
					<tr bgcolor=<%=color_fondo%>>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td><td class=tdbordeCELDAB8 align="center"><%=LitStockDesv%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td>
						<td class="TDBORDECELDAB8" align="center"><%=LitAnt%></td><td class="TDBORDECELDAB8" align="center"><%=LitSNuevo%></td>
					</tr><%
			    CloseFila
		end if

		'Representar datos de la tabla de usuario
		fila=1

		if agrupar="agr_Documento" and detalle="sin" then
		   while not rstAux.eof and fila<=MAXPAGINA
			    DrawFila color_blau
					DrawCelda2 "DATO ALIGN=LEFT", "left", false, EncodeForHtml(trimCodEmpresa(rstAux("documento")))
					DrawCelda2 "DATO ALIGN=LEFT", "left", false, EncodeForHtml(rstAux("fecha"))
					DrawCelda2 "DATO ALIGN=LEFT", "left", false, EncodeForHtml(rstAux("responsable"))
					DrawCelda2 "DATO ALIGN=LEFT", "left", false, EncodeForHtml(rstAux("observaciones"))
					if verCost=1 then
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", false, EncodeForHtml(formatnumber(rstAux("coste"),n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", false, EncodeForHtml(formatnumber(rstAux("costeDesv"),n_decimales,-1,0,-1))
					end if
					fila=fila+1
				CloseFila
				''rstAux.movenext

    			documAnt=rstAux("documento")

				if verCost="1" then
					''ricardo 20-4-2007 se cambia la forma de calcular los totales, de esta
				    ''	totaldocumAnt=rstAux("totalAcum")
				    '' a esta otra
				    totaldocumAnt=rstAux("totalDoc")
				    totaldocumStockDesvAnt=rstAux("TotalDocDesv")
				end if


				rstAux.movenext

				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotal&"</b>",1
						'DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,formatnumber(totalDocum,DEC_PREC,-1,0,-1)
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if

			wend
		end if
                                                                                                                                                                                                                                                                                             
                                                                                                                                                                           
		if agrupar="agr_Documento" and detalle="con" then
		   documAnt=""
'		   if totDoc>0 then                                                                                                  
'		      totalDocum=totDoc
'		   end if                                                                                                
		   while not rstAux.eof and fila<=MAXPAGINA                                                        
			    DrawFila color_blau
				    if documAnt<>rstAux("documento") then
						DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(trimCodEmpresa(rstAux("documento"))))
						DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("fecha")))
						DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("responsable")))
					else
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, ""
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, ""
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, ""
					end if
					DrawCelda2 "DATO ALIGN=LEFT width='80'", "left", false, EncodeForHtml(null_s(trimCodEmpresa(rstAux("referencia"))))
					DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("nombreArticulo")))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stockAnt")) ,DEC_CANT,-1,0,-1))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stockNew")) ,DEC_CANT,-1,0,-1))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("StockDif")) ,DEC_CANT,-1,0,-1))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stMinAnt")) ,DEC_CANT,-1,0,-1))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stMinNew")) ,DEC_CANT,-1,0,-1))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stRepAnt")) ,DEC_CANT,-1,0,-1))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stRepNew")) ,DEC_CANT,-1,0,-1))
					if verCost=1 then
					    DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("coste")),DEC_PREC,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("total")),n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("totalDesv")),n_decimales,-1,0,-1))
						''totalDocum=totalDocum+rstAux("total")
					end if
					fila=fila+1                                                                           
				CloseFila

				documAnt=EncodeForHtml(null_s(rstAux("documento")))                    

				if verCost="1" then
					''ricardo 20-4-2007 se cambia la forma de calcular los totales, de esta
				    ''	totaldocumAnt=rstAux("totalAcum")
				    '' a esta otra
				    totaldocumAnt=EncodeForHtml(null_s(rstAux("totalDoc")))              
				    totaldocumStockDesvAnt=EncodeForHtml(null_s(rstAux("TotalDocDesv")))

				    totaltotaldocumAnt=EncodeForHtml(null_s(rstAux("TotalTotalstockAnt")))
				    totaltotaldocumStockDesvAnt=EncodeForHtml(null_s(rstAux("TotalTotalStockDif")))

				end if
				totaldocumStockAnt1=EncodeForHtml(null_s(rstAux("TotalstockAnt")))        
				totaldocumStockNue1=EncodeForHtml(null_s(rstAux("TotalstockNew")))
				totaldocumStockDesv1=EncodeForHtml(null_s(rstAux("TotalStockDif")))


				totaltotaldocumStockAnt1=EncodeForHtml(null_s(rstAux("totaltotalDesv")))
				totaltotaldocumStockNue1=EncodeForHtml(null_s(rstAux("totaltotalDoc")))
				totaltotaldocumStockDesv1=EncodeForHtml(null_s(rstAux("totalTotalDocDesv")))

				rstAux.movenext

'			    if documAnt<>rstAux("documento") and not rstAux.eof then
			    if not rstAux.eof then
				    if documAnt<>rstAux("documento") and verCost=1  then
						DrawFila color_fondo
					    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
					    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotDocum&"</b>",2
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
						CloseFila
						totalDocum=0
					end if
				    if documAnt<>rstAux("documento") and verCost=0  then
						DrawFila color_fondo
					    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
					    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotDocum&"</b>",2
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
						CloseFila
						totalDocum=0
					end if
				end if
				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotDocum&"</b>",2
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if
				if rstAux.eof and verCost=0 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotDocum&"</b>",2
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
					CloseFila
				end if
			wend

			''ahora se vera una linea de totales de todo el listado
				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotListRegInv&"</b>",2
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaltotaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaltotaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if
				if rstAux.eof and verCost=0 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotListRegInv&"</b>",2
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
					CloseFila
				end if

		end if

		'Agrupados por Artículo
		if agrupar="agr_Articulo" and detalle="sin" then
		   articleAnt=""
		   while not rstAux.eof  and fila<=MAXPAGINA
			    DrawFila color_blau
				    if articleAnt<>rstAux("referencia") then
				        if rstAux("nombreArticulo")=" " then
							DrawCelda2 "DATO ALIGN=LEFT ", "left", false, LitRefSinNombre
						else
							DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("nombreArticulo")))                   
						end if
					else
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", false, ""
					end if
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(trimCodEmpresa(rstAux("documento"))))                    
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("fecha")))
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("responsable")))
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("observaciones")))
					if verCost=1 then
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("coste"),DEC_PREC,-1,0,-1)))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("total"),n_decimales,-1,0,-1)))      
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("totalDesv"),n_decimales,-1,0,-1)))   
					end if
					fila=fila+1                                                      
				CloseFila                                                                            
				articleAnt=rstAux("referencia")
				''rstAux.movenext

				if verCost="1" then
					''ricardo 20-4-2007 se cambia la forma de calcular los totales, de esta
				    ''	totaldocumAnt=rstAux("totalAcum")
				    '' a esta otra
				    totaldocumAnt=rstAux("totalDoc")
				    totaldocumStockDesvAnt=rstAux("TotalDocDesv")
				end if
				rstAux.movenext

				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",6
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotal&"</b>",1
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if
			wend
		end if

		if agrupar="agr_Articulo" and detalle="con" then
		   articleAnt=""
		   while not rstAux.eof  and fila<=MAXPAGINA                 
			    DrawFila color_blau
				    if articleAnt<>rstAux("referencia") then
				        DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(trimCodEmpresa(rstAux("referencia")))) 
						if rstAux("nombreArticulo")=" " then
							DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, LitRefSinNombre
						else
							DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("nombreArticulo")))   
						end if
					else                                                                                     
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, ""                       
                        DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, ""
					end if
					DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(trimCodEmpresa(rstAux("documento")))) 
					DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("fecha")))
					DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("responsable")))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stockAnt"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stockNew"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("StockDif"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stMinAnt"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stMinNew"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stRepAnt"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stRepNew"),DEC_CANT,-1,0,-1)))
					if verCost=1 then
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("coste"),DEC_PREC,-1,0,-1)))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("total"),n_decimales,-1,0,-1)))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("totalDesv"),n_decimales,-1,0,-1))) 
					end if
					fila=fila+1
				CloseFila                                                                                    
				articleAnt=rstAux("referencia")
				''rstAux.movenext

				'''''''''''''''''''''''''
    			if verCost="1" then
					''ricardo 20-4-2007 se cambia la forma de calcular los totales, de esta
				    ''	totaldocumAnt=rstAux("totalAcum")
				    '' a esta otra                                                                             
				    totaldocumAnt=EncodeForHtml(rstAux("totalDoc"))                                          
				    totaldocumStockDesvAnt=EncodeForHtml(rstAux("TotalDocDesv"))        
				end if
				totaldocumStockAnt1=EncodeForHtml(rstAux("TotalstockAnt"))
				totaldocumStockNue1=EncodeForHtml(rstAux("TotalstockNew"))
				totaldocumStockDesv1=EncodeForHtml(rstAux("TotalStockDif"))
				rstAux.movenext

				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotal&"</b>",1
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
						DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if
				if rstAux.eof and verCost=0 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotal&"</b>",1
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
						DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
					CloseFila
				end if
				'''''''''''''''''''''''''
			wend
		end if

		'Agrupados por SubFamilia
		if agrupar="agr_SubFamilia" and detalle="sin" then
		   SubFamiliaAnt=""
		   while not rstAux.eof  and fila<=MAXPAGINA
			    DrawFila color_blau
				    if SubFamiliaAnt & ""<>rstAux("subfamilia") & "" then
						if rstAux("nombreSubFamilia")=" " then
							DrawCelda2 "DATO ALIGN=LEFT ", "left", false, LitRefSinNombre
						else
							DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("nombreSubFamilia")))              
						end if
					else                                                                                
						DrawCelda2 "DATO ALIGN=RIGHT ", "left", false, ""                                         
					end if
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(trimCodEmpresa(rstAux("documento"))))
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("fecha")))
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("responsable")))
					DrawCelda2 "DATO ALIGN=LEFT ", "left", false, EncodeForHtml(null_s(rstAux("observaciones")))
					if verCost=1 then
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("coste"),DEC_PREC,-1,0,-1)))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("total"),n_decimales,-1,0,-1)))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("totalDesv"),n_decimales,-1,0,-1)))
					end if                                                                   
					fila=fila+1
				CloseFila
				SubFamiliaAnt=EncodeForHtml(null_s(rstAux("subfamilia"))) & ""                  
				''rstAux.movenext

				if verCost="1" then
					''ricardo 20-4-2007 se cambia la forma de calcular los totales, de esta
				    ''	totaldocumAnt=rstAux("totalAcum")
				    '' a esta otra
				    totaldocumAnt=rstAux("totalDoc")
				    totaldocumStockDesvAnt=rstAux("TotalDocDesv")
				end if
				rstAux.movenext

				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotal&"</b>",1
						'DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,formatnumber(totalDocum,DEC_PREC,-1,0,-1)
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if
			wend
		end if

		if agrupar="agr_SubFamilia" and detalle="con" then
		   SubFamiliaAnt=""
		   while not rstAux.eof  and fila<=MAXPAGINA
			    DrawFila color_blau
				    if SubFamiliaAnt & ""<>rstAux("subfamilia") & "" then
						if rstAux("nombreSubFamilia")=" " then
							DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, LitSubFamSinNombre
						else
							DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(rstAux("nombreSubFamilia"))
						end if
					else
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, ""
					end if
					DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(trimCodEmpresa(rstAux("documento")))) 
					DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("fecha")))
					DrawCelda2 "DATO ALIGN=LEFT width='150'", "left", false, EncodeForHtml(null_s(rstAux("responsable")))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stockAnt"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stockNew"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("StockDif"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stMinAnt"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stMinNew"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stRepAnt"),DEC_CANT,-1,0,-1)))
					DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("stRepNew"),DEC_CANT,-1,0,-1)))
					if verCost=1 then
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("coste"),DEC_PREC,-1,0,-1)))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("total"),n_decimales,-1,0,-1)))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false, EncodeForHtml(null_s(formatnumber(rstAux("totalDesv"),n_decimales,-1,0,-1))) 
					end if
					fila=fila+1
				CloseFila
				SubFamiliaAnt=rstAux("subfamilia") & ""
				''rstAux.movenext

				'''''''''''''''''''''''''
    			if verCost="1" then
					''ricardo 20-4-2007 se cambia la forma de calcular los totales, de esta
				    ''	totaldocumAnt=rstAux("totalAcum")
				    '' a esta otra
				    totaldocumAnt=EncodeForHtml(rstAux("totalDoc"))                                                          
				    totaldocumStockDesvAnt=EncodeForHtml(rstAux("TotalDocDesv"))                                   

				    totaltotaldocumAnt=EncodeForHtml(rstAux("TotalTotalstockAnt"))                                    
				    totaltotaldocumStockDesvAnt=EncodeForHtml(rstAux("TotalTotalStockDif"))
				end if
				totaldocumStockAnt1=EncodeForHtml(rstAux("TotalstockAnt"))                                                 
				totaldocumStockNue1=EncodeForHtml(rstAux("TotalstockNew"))                                        
				totaldocumStockDesv1=EncodeForHtml(rstAux("TotalStockDif"))                      

				totaltotaldocumStockAnt1=EncodeForHtml(rstAux("totaltotalDesv"))                           
				totaltotaldocumStockNue1=EncodeForHtml(rstAux("totaltotalDoc"))                     
				totaltotaldocumStockDesv1=EncodeForHtml(rstAux("totalTotalDocDesv"))
				rstAux.movenext

			    if not rstAux.eof then
				    if SubFamiliaAnt & ""<>rstAux("subfamilia") & "" and verCost=1  then
						DrawFila color_fondo
					    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
					    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotSubFam&"</b>",1
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
						DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
						CloseFila
						totalDocum=0
					end if
				    if SubFamiliaAnt & ""<>rstAux("subfamilia") & "" and verCost=0  then
						DrawFila color_fondo
					    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
					    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotSubFam&"</b>",1
						    DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
						    DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
						    DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
						    DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
						CloseFila
						totalDocum=0
					end if
				end if

				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotSubFam&"</b>",1
						    DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
						    DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
						    DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
						    DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if
				if rstAux.eof and verCost=0 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotSubFam&"</b>",1
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockAnt1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockNue1),DEC_CANT,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaldocumStockDesv1),DEC_CANT,-1,0,-1))
						DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
					CloseFila
				end if
				'''''''''''''''''''''''''

			''ahora se vera una linea de totales de todo el listado
				if rstAux.eof and verCost=1 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotListRegInv&"</b>",1
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",5
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaltotaldocumAnt,n_decimales,-1,0,-1))
						DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(totaltotaldocumStockDesvAnt,n_decimales,-1,0,-1))
					CloseFila
				end if
				if rstAux.eof and verCost=0 then
					DrawFila color_fondo
				    	DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",3
				    	DrawCeldaSpan "DATO ALIGN=RIGHT ","","",0,"<b>"&LitTotListRegInv&"</b>",1
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockAnt1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockNue1),DEC_CANT,-1,0,-1))
							DrawCelda2 "DATO ALIGN=RIGHT width='150'", "left", false,EncodeForHtml(formatnumber(null_z(totaltotaldocumStockDesv1),DEC_CANT,-1,0,-1))
							DrawCeldaSpan "DATO ALIGN=LEFT","","",0,"",4
					CloseFila
				end if
			wend
		end if
		NavPaginas lote,lotes,campo,criterio,texto,2
	 end if
%>
			</table><%
	end if 'fin del mode browse
%></form>
<%
end if
set rstAux = Nothing
set rst = Nothing
set rst2 = Nothing
set rstSelect = Nothing
set rstTablas = Nothing

%>
<%connRound.close
set connRound = Nothing%>
</body>

</html>