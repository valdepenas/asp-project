<%@ Language=VBScript %>
<%
''ricardo 16-11-2007 se cambia la dsn desde dsncliente a backendlistados
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
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
<!--#include file="../../modulos.inc" -->
<!--#include file= "../../CatFamSubResponsive.inc"-->
<!--#include file="listado_escandallo.inc" -->
<!--#include file= "../../styles/formularios.css.inc"-->  
<!--#include file="../../common/campospersoResponsive.inc" -->
<title><%=LitTitulo2%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>

<link rel="STYLESHEET" href="../../pantalla.css" media="SCREEN"/>
<link rel="STYLESHEET" href="../../impresora.css" media="PRINT"/>
</head>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<body onload="self.status='';" class="BODY_ASP">

<%
'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0

set conn = Server.CreateObject("ADODB.Connection")
conn.open session("backendlistados")

set rstAux = Server.CreateObject("ADODB.Recordset")
set rstAux2 = Server.CreateObject("ADODB.Recordset")
set rst = Server.CreateObject("ADODB.Recordset")
set rst1 = Server.CreateObject("ADODB.Recordset")


    si_tiene_modulo_importaciones=ModuloContratado(session("ncliente"),ModImportaciones)%>
<form name="listado_escandallo" method="post">
    <%PintarCabecera "listado_escandallo.asp"
    
    WaitBoxOculto LitEsperePorFavor

    'Leer parámetros de la página
    mode=Request.QueryString("mode")
    if trim(mode)="browse" then mode="ver"

    if request.querystring("referencia") >"" then
       referencia = limpiaCadena(request.querystring("referencia"))
    else
       referencia = limpiaCadena(request.form("referencia"))
    end if
    if referencia>"" then
	    referencia_aux = "%" & referencia & "%"
    end if

    if request.form("nombre")>"" then
	    nombre = limpiaCadena(request.form("nombre"))
    else
	    nombre=limpiaCadena(request.querystring("nombre"))
    end if
    if nombre>"" then
	    nombre_aux = "%" & nombre & "%"
    end if

    ''MPC 19/05/2009 Recogida de los nuevos filtros añadidos en el listado
    if request.form("familia")>"" then
		familia = limpiaCadena(request.form("familia"))
	else
		familia = limpiaCadena(request.querystring("familia"))
	end if

	if request.form("familia_padre")>"" then
		familia_padre = limpiaCadena(request.form("familia_padre"))
	else
		familia_padre = limpiaCadena(request.querystring("familia_padre"))
	end if

	if request.form("categoria")>"" then
		categoria = limpiaCadena(request.form("categoria"))
	else
		categoria = limpiaCadena(request.querystring("categoria"))
	end if
    
    if request.form("tipoarticulo")>"" then
		tipoarticulo = limpiaCadena(request.form("tipoarticulo"))
	else
		tipoarticulo = limpiaCadena(request.querystring("tipoarticulo"))
	end if
	
	if request.form("agrupar")>"" then
		agrupar = limpiaCadena(request.form("agrupar"))
	else
		agrupar = limpiaCadena(request.querystring("agrupar"))
	end if

	if request.form("ver_coste")>"" then
		ver_coste = limpiaCadena(request.form("ver_coste"))
	else
		ver_coste = limpiaCadena(request.querystring("ver_coste"))
	end if
    
    if request.form("ver_precio")>"" then
		ver_precio = limpiaCadena(request.form("ver_precio"))
	else
		ver_precio = limpiaCadena(request.querystring("ver_precio"))
	end if
	
	if request.form("ver_agrtalla")>"" then
		ver_agrtalla = limpiaCadena(request.form("ver_agrtalla"))
	else
		ver_agrtalla = limpiaCadena(request.querystring("ver_agrtalla"))
	end if
	
	if request.form("ver_agrcolor")>"" then
		ver_agrcolor = limpiaCadena(request.form("ver_agrcolor"))
	else
		ver_agrcolor = limpiaCadena(request.querystring("ver_agrcolor"))
	end if
	
	if request.form("ver_agrupa")>"" then
		ver_agrupa = limpiaCadena(request.form("ver_agrupa"))
	else
		ver_agrupa = limpiaCadena(request.querystring("ver_agrupa"))
	end if
    ''FIN MPC

    if request.form("ordenar")>"" then
	    ordenar = limpiaCadena(request.form("ordenar"))
    else
	    ordenar = limpiaCadena(request.querystring("ordenar"))
    end if

    if request.form("almacen")>"" then
	    almacen = limpiaCadena(request.form("almacen"))
    else
	    almacen = limpiaCadena(request.querystring("almacen"))
    end if
    CheckCadena almacen%>
	<table width='100%'>
   	    <tr>
            <td><font class='CABECERA'><b></b></font>
                <font class="CELDA"><b></b></font>
            </td>
            <td width="20%">
                <font class='CABECERA'><b></b></font>
            </td>
            <td width="30%">
                <font class='CABECERA'><b></b></font>
            <%if mode="ver" then%>
                <font class="CELDA"><b>&nbsp;<%="(" & LitEmitido & " " & day(date) & "/" & month(date) & "/" & year(date) & ")"%></b></font>
            <%end if%>
        </tr>
    </table>
    <hr/>
    <table>
    <%' ********  ESCRIBIR LAS OPCIONES DE LISTADO **********
		if referencia>"" then
			DrawFila color_blau
      	      DrawCelda2 "CELDA", "left", false, "<B>" + LitConref + ": </B>"
			DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(referencia))
			CloseFila
		end if
		if nombre>"" then
			DrawFila color_blau
			DrawCelda2 "CELDA", "left", false, "<B>" + LitConNombre + ": </B>"
			DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(nombre))
			CloseFila
		end if
		if familia>"" then
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false,"<b>" & LitSubFamilia + ": </b>"
				DrawCelda2 "CELDA", "left", false, NombresEntidades(familia,"familias","codigo","nombre",session("backendlistados"))
			CloseFila
		elseif familia_padre<>"" then
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false,"<b>" & LitFamilia + ": </b>"
				DrawCelda2 "CELDA", "left", false, NombresEntidades(familia_padre,"familias_padre","codigo","nombre",session("backendlistados"))
			CloseFila
		elseif categoria<>"" then
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false,"<b>" & LitCategoria + ": </b>"
				DrawCelda2 "CELDA", "left", false, NombresEntidades(categoria,"categorias","codigo","nombre",session("backendlistados"))
			CloseFila
		end if
		if tipoarticulo>"" then
			strselect ="select descripcion from tipos_entidades with(nolock) where codigo='" & tipoarticulo & "' and tipo='ARTICULO'"
            rst.cursorlocation=3
			rst.open strselect, session("backendlistados")
			if not rst.eof then
				DrawFila color_blau
				DrawCelda2 "CELDA", "left", false,"<b>" & LitTipoArticulo + ": </b>"
				DrawCelda2 "CELDA", "left", false, rst("descripcion")
				CloseFila
			end if
			rst.close
		end if
		if almacen>"" then
			strselect ="select descripcion from almacenes with(nolock) where codigo='" & almacen & "'"
			rst.cursorlocation=3
			rst.open strselect, session("backendlistados")
			if not rst.eof then
				DrawFila color_blau
				if si_tiene_modulo_importaciones=0 then
      		      DrawCelda2 "CELDA", "left", false, "<B>" + LitAlmacen + ": </B>"
      		    end if
				DrawCelda2 "CELDA", "left", false, rst("descripcion")
				CloseFila
			end if
			rst.close
		end if
		if ordenar>"" then
		    DrawFila color_blau
			    DrawCelda2 "CELDA", "left", false,"<b>" & LitOrdenar + ": </b>"
			    DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(ordenar))
			CloseFila
		end if
		if agrupar>"" then
		    DrawFila color_blau
			    DrawCelda2 "CELDA", "left", false,"<b>" & LitAgruparEscandallo + ": </b>"
			    DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(agrupar))
			CloseFila
		end if%>
    </table>
	<%Alarma "listado_articulo.asp"

  '*********************************************************************************************
  'Se muestran parametros de seleccion
  '*********************************************************************************************

  if mode="add" then%>
        <%' **** PEDIR LOS DATOS PARA FILTRAR EL LISTADO *****
          'DrawCelda2 "CELDA", "left", false, "<B>" & LitFiltrosParam & "</b>"
        %>
    <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitFiltrosParam%></h6>
        <%
        'DrawCelda2 "CELDA width='11%'", "left", false, LitConref + ": "
   	    'DrawInputCelda "CELDA width='20%'","","",25,0,"","referencia",referencia
        EligeCelda "input","add","left","","",0,LitConref,"referencia",35,referencia
        'EligeCelda "input","",""
        'DrawCelda2 "CELDA width='11%'", "left", false, LitConNombre + ": "
	    'DrawInputCelda "CELDA","","",25,0,"","nombre",nombre
        EligeCelda "input","add","left","","",0,LitConNombre,"nombre",35,nombre

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
			ConfigDespleg(i,9)=LitCategoria & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
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

		    'DrawCelda2 "CELDA width='11%'", "left", false, LitTipoArticulo + ": "
		    rstAux.cursorlocation=3
	   		rstAux.open " select codigo, descripcion from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='articulo' order by descripcion", session("backendlistados")'',adOpenKeyset,adLockOptimistic
       		'DrawSelectCelda "CELDA width='20%'","150","",0,"","tipoarticulo",rstAux,tipoarticulo,"codigo","descripcion","",""
            'DrawSelectCelda(estilo,ancho,alto,tabulacion,etiqueta,name,reg,value,campo,campo2,evento,funcion)
            DrawSelectCelda "CELDA","","",0,LitTipoArticulo,"tipoarticulo",rstAux,tipoarticulo,"codigo","descripcion","",""
	   		rstAux.close
			if si_tiene_modulo_importaciones=0 then
           	    'DrawCelda2 "CELDA width='11%'", "left", false, LitAlmacen + ": "
           		rstAux.cursorlocation=3
		   		rstAux.open " select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion", session("backendlistados")'',adOpenKeyset,adLockOptimistic
	       		'DrawSelectCelda "CELDA","150","",0,"","almacen",rstAux,almacen,"codigo","descripcion","",""
                DrawSelectCelda "CELDA","","",0,LitAlmacen,"almacen",rstAux,almacen,"codigo","descripcion","",""
		   		rstAux.close
			end if
	        'DrawCelda2 "CELDA width='11%'", "left", false, LitOrdenar + ": "
            DrawDiv "1","",""
            DrawLabel "","",LitOrdenar%><select class='width60' name="ordenar">
			        <option selected="selected" style='width:175px'  value="Referencia"><%=ucase(LitRef)%></option>
				    <option value="Nombre"><%=ucase(LitNombre)%></option>
				    <option value="Familia"><%=ucase(LitFamilia)%></option>
		   	    </select>
		    <%
            CloseDiv
            'DrawCelda2 "CELDA width='11%'", "left", false, LitAgruparEscandallo + ": "
            DrawDiv "1","",""
            DrawLabel "","",LitAgruparEscandallo%><select class='width60' style='width:175px' name="agrupar">
		        <option value="SUBFAMILIA"><%=LitSubFam%></option>
		        <option selected="selected" value=""></option>
		    </select>
            <%
		    CloseDiv
	    '******************************************************%>
	<hr/>
	    <%'DrawFila color_fondo
			'DrawCelda2 "ENCABEZADOL", "left", false, LitCamposOpcionales
		'CloseFila%>
    <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitCamposOpcionales%></h6>
	    <%
			'DrawCelda2 "CELDA", "left", false, LitCostes
			'DrawCheckCelda "CELDA","","",0,"","ver_coste",cstr(ver_coste)
            EligeCelda "check",mode,"","","",0,LitCostes,"ver_coste",0,cstr(ver_coste)
			'DrawCelda "CELDA","7%","",0," "
			'DrawCelda2 "CELDA", "left", false, LitPVPVenta
			'DrawCheckCelda "CELDA","","",0,"","ver_precio",cstr(ver_precio)
            EligeCelda "check",mode,"","","",0,LitPVPVenta,"ver_precio",0,cstr(ver_precio)
			'DrawCelda "CELDA","7%","",0," "
			'DrawCelda2 "CELDA", "left", false, LitAgrTallas
			'DrawCheckCelda "CELDA","","",0,"","ver_agrtalla",cstr(ver_agrtalla)
            EligeCelda "check",mode,"","","",0,LitAgrTallas,"ver_agrtalla",0,cstr(ver_agrtalla)
			'DrawCelda "CELDA","7%","",0," "
			'DrawCelda2 "CELDA", "left", false, LitAgrColores
			'DrawCheckCelda "CELDA","","",0,"","ver_agrcolor",cstr(ver_agrcolor)
            EligeCelda "check",mode,"","","",0,LitAgrColores,"ver_agrcolor",0,cstr(ver_agrcolor)
			'DrawCelda "CELDA","7%","",0," "
			'DrawCelda2 "CELDA", "left", false, LitAgrupa
			'DrawCheckCelda "CELDA","","",0,"","ver_agrupa",cstr(ver_agrupa)
            EligeCelda "check",mode,"","","",0,LitAgrupa,"ver_agrupa",0,cstr(ver_agrupa)
            %>
    <%end if

   '*********************************************************************************************
   ' Se muestran los datos de la consulta
   '*********************************************************************************************

    if mode="ver" or mode="edit" then
        sentido=limpiaCadena(Request.QueryString("sentido"))
        'CREAR EN LA TABLA TEMPORAL EL LISTADO DE ESCANDALLO.
        set rst1 = conn.execute("EXEC sp_ListadoEscandallo @NomTabla='" & session("usuario") & "' , @ParamRef='" & referencia_aux & "' , @ParamNombre='" & nombre_aux & "' , @ParamCat='" & categoria & "', @ParamFamPadre='" & familia_padre & "', @ParamFamilia='" & familia & "', @ParamTipoArt='" & tipoarticulo & "', @ParamAlmacen='" & almacen & "' , @ParamOrdenar='" & ordenar & "', @ParamAgrupar='" & agrupar & "', @session_empresa='" & session("ncliente") & "'")

  %><hr/>
        <input type="hidden" name="referencia" value="<%=enc.EncodeForHtmlAttribute(null_s(referencia))%>"/>
 	    <input type="hidden" name="nombre" value="<%=enc.EncodeForHtmlAttribute(null_s(nombre))%>"/>
		<input type="hidden" name="familia" value="<%=enc.EncodeForHtmlAttribute(null_s(familia))%>"/>
		<input type="hidden" name="familia_padre" value="<%=enc.EncodeForHtmlAttribute(null_s(familia_padre))%>"/>
		<input type="hidden" name="categoria" value="<%=enc.EncodeForHtmlAttribute(null_s(categoria))%>"/>
		<input type="hidden" name="tipoarticulo" value="<%=enc.EncodeForHtmlAttribute(null_s(tipoarticulo))%>"/>
		<input type="hidden" name="agrupar" value="<%=enc.EncodeForHtmlAttribute(null_s(agrupar))%>"/>
		<input type="hidden" name="ver_coste" value="<%=enc.EncodeForHtmlAttribute(null_s(ver_coste))%>"/>
		<input type="hidden" name="ver_precio" value="<%=enc.EncodeForHtmlAttribute(null_s(ver_precio))%>"/>
		<input type="hidden" name="ver_agrtalla" value="<%=enc.EncodeForHtmlAttribute(null_s(ver_agrtalla))%>"/>
		<input type="hidden" name="ver_agrcolor" value="<%=enc.EncodeForHtmlAttribute(null_s(ver_agrcolor))%>"/>
		<input type="hidden" name="ver_agrupa" value="<%=enc.EncodeForHtmlAttribute(null_s(ver_agrupa))%>"/>
		<input type="hidden" name="almacen" value="<%=enc.EncodeForHtmlAttribute(null_s(almacen))%>"/>
		<input type="hidden" name="ordenar" value="<%=enc.EncodeForHtmlAttribute(null_s(ordenar))%>"/>
		<%MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='105'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='105'", DSNIlion)%>
		<input type='hidden' name='maxpdf' value='<%=MAXPDF%>'/>
		<input type='hidden' name='maxpagina' value='<%=MAXPAGINA%>'/>

		<%if ordenar="REFERENCIA" then
	        orden =  ""
        elseif ordenar="NOMBRE" then
	        orden =  " NOMBRE,"
        elseif ordenar="FAMILIA" then
	        orden =  "FAMILIA,"
        end if

        orden1=""
        if agrupar="SUBFAMILIA" then
            orden =  "FAMILIA,"
            orden1= "SUBFAMILIA,"
        end if

        rst.cursorlocation=3
        rst.open "select * from egesticet.[" & session("usuario") & "] ORDER BY " & orden & " ARTICULO," & orden1 & "NDET,PARAMETRO", session("backendlistados")%>

        <input type="hidden" name="nRegs" value="<%=rst.RecordCount%>"/>

        <%if not rst.EOF then
	        lote=limpiaCadena(Request.QueryString("lote"))

		    if lote="" then
		        lote=1
		    end if

		    sentido=limpiaCadena(Request.QueryString("sentido"))
            lotes=(rst.RecordCount/MAXPAGINA)
            if lotes>(rst.RecordCount/MAXPAGINA) then
                lotes=(rst.RecordCount/MAXPAGINA)+1
            else
                lotes=(rst.RecordCount/MAXPAGINA)
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

            rst.PageSize=MAXPAGINA
            rst.AbsolutePage=lote

            NavPaginas lote,lotes,campo,criterio,texto,1
            VinculosPagina(MostrarArticulos)=1
            CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

            fila=1

            articulo_ant=""
            total_coste=0
            total_cantidad=0

%><table width='100%'  border='0' cellspacing="1" cellpadding="1" style="border-collapse: collapse;"><%
		while not rst.EOF and fila<=MAXPAGINA
			CheckCadena rst("articulo")

				if rst("parametro")>"" then
					tipo_art="PARAMETRO"
					if tipo_art_Ant<>tipo_art then
						cab_param="1"
					else
						cab_param="0"
					end if
				else
					if rst("articulo_esc")>"" then
						tipo_art="ESCANDALLO"
						if tipo_art_Ant<>tipo_art then
							cab_esc="1"
						else
						    cab_esc="0"
						end if
					else
						tipo_art="PRINCIPAL"
						if tipo_art_Ant<>tipo_art then
							cab_prin="1"
						else
							cab_prin="0"
						end if
					end if
				end if

				'Pintar encabezados
				if tipo_art="PARAMETRO" and cab_param="1" then%>
					<!--<table width='100%'  BORDER="1" cellspacing="1" cellpadding="1">-->
					    
					    <%DrawFila color_blau
					        DrawCelda "CELDA width='3%'","","",0,"&nbsp;"
		    			    DrawCelda "CELDA width='10%'","","",0,"&nbsp;"
		    			    DrawCelda "CELDA7 width='15%'","","",0,"<b>" &LitParametro & "</b>"
	    			        DrawCelda "CELDA7 width='20%' colspan='2'","","",0,"<b>" & LitValor & "</b>"
	    	    	 	    DrawCelda "CELDA7 width='13%'","","",0,"<b>" & LitMedida & "</b>"
	    	    	 	    DrawCelda "CELDAC7 width='5%'","","",0,"<b>" & LitDepPrecio & "</b>"
			  		    CloseFila
				elseif tipo_art="ESCANDALLO" and cab_esc="1" then%>
					<!--<table width='100%'  BORDER="1" cellspacing="1" cellpadding="1">-->
					    <%DrawFila color_terra
						    DrawCelda "CELDA7 width='3%' bgcolor=" & color_blau,"","",0,"&nbsp;"
						    DrawCelda "CELDA7 width='10%'","","",0,"<b>" &LitRef & "</b>"
						    DrawCelda "CELDA7 width='15%'","","",0,"<b>" & LitNombre & "</b>"
						    DrawCelda "CELDA7 width='15%'","","",0,"<b>" & LitSubFamilia & "</b>"
						    DrawCelda "CELDAR7 width='5%'","","",0,"<b>" & LitCantidad & "</b>"
						    if si_tiene_modulo_importaciones=0 then
							    DrawCelda "CELDA7 width='13%'","","",0,"<b>" & LitAlmacen & "</b>"
						    end if
						    if ver_coste="on" then
    						    DrawCelda "CELDAR7 width='5%'","","",0,"<b>" & LitCoste & "</b>"
    						end if
    						if ver_precio="on" then
    						    DrawCelda "CELDAR7 width='5%'","","",0,"<b>" & LitPrecio & "</b>"
    						end if
						    if ver_agrtalla="on" then
		    				    DrawCelda "CELDA7 width='10%'","","",0,"<b>" & LitAgrTalla & "</b>"
						    end if
						    if ver_agrcolor="on" then
	    					    DrawCelda "CELDA7 width='10%'","","",0,"<b>" & LitAgrColor & "</b>"
						    end if
						    if ver_agrupa="on" then
    						    DrawCelda "CELDA7 width='10%'","","",0,"<b>" & LitAgrupa & "</b>"
                            end if
						    DrawCelda "CELDAC7 width='5%'","","",0,"<b>" & LitTYC & "</b>"
					    CloseFila
				elseif tipo_art="PRINCIPAL" and cab_prin="1" then
					if articulo_Ant<>"" and articulo_ant<>rst("articulo") then%>
					    <%if ver_coste="on" then
					        rstAux.cursorlocation=3
						    rstAux.open "SELECT sum(isnull(cast(coste as real), 0)) as coste from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "'",session("backendlistados")
						    if not rstAux.eof then
						        total_coste = null_z(rstAux("coste"))
						    else
						        total_coste=0
						    end if
						    rstAux.close
                        end if

						rstAux.cursorlocation=3
						rstAux.open "SELECT sum(isnull(cast(cantidad as real), 0)) as cantidad from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "'",session("backendlistados")
						if not rstAux.eof then
						    total_cantidad = null_z(rstAux("cantidad"))
						else
						    total_cantidad=0
						end if
						rstAux.close

						if ver_precio="on" then
    						rstAux.cursorlocation=3
	    					rstAux.open "SELECT sum(isnull(cast(precio as real), 0)) as precio from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "'",session("backendlistados")
	    					if not rstAux.eof then
		    				    total_precio = null_z(rstAux("precio"))
		    				else
		    				    total_precio=0
		    				end if
			    			rstAux.close
			    		end if

						DrawFila color_fondo
						    DrawCelda "CELDA7 width='3%'","","",0,"&nbsp;"
						    if si_tiene_modulo_importaciones<>0 then
						        DrawCelda "CELDA width='40%' colspan='3'","","",0,"<b>" & LitTotalCar & ":</b>"
	            		    else
						        DrawCelda "CELDA width='40%' colspan='3'","","",0,"<b>" & LitTotalEsc & ":</b>"
						    end if
						    DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_cantidad,DEC_CANT,-1,0,-1) & "</b>"
						    if si_tiene_modulo_importaciones=0 then
						        DrawCelda "CELDA7 width='13%'","","",0,"&nbsp;"
					        end if
						    if ver_coste="on" then
						    DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_coste,DEC_PREC,-1,0,-1) & "</b>"
						    end if
						    if ver_precio="on" then
						    DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_precio,DEC_PREC,-1,0,-1) & "</b>"
						    end if
		    	     	    suma_colspan=1
		    	     	    if ver_agrupa="on" then
		    	     	        suma_colspan=suma_colspan+1
		    	     	    end if
		    	     	    if ver_agrtalla="on" then
		    	     	        suma_colspan=suma_colspan+1
		    	     	    end if
		    	     	    if ver_agrcolor="on" then
		    	     	        suma_colspan=suma_colspan+1
		    	     	    end if
						    DrawCelda "CELDA7 width='5%' colspan='" & suma_colspan & "'","","",0,"&nbsp;"
						CloseFila
					%><!--</table>-->
					<tr><td><br /></td></tr><%
					end if

				    %><!--<table width='100%'  style="border: 1px solid Black;" cellspacing="0" cellpadding="0">-->
				    <tr bgcolor='<%=color_fondo%>' style="border:1px solid black;">
				    <%
					''DrawFila color_fondo
			    		DrawCelda "CELDA style='width:20%;border-bottom:1px solid black;border-left:1px solid black;border-top:1px solid black;' colspan='2'","","",0,"<b>" &LitRefPadre & "</b>"
		    		    DrawCelda "CELDA style='width:25%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,"<b>" & LitNombre & "</b>"
		    	     	DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,"<b>" & LitSubFamilia & "</b>"
		    	     	if ver_agrtalla="on" then
		    	     	DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,"<b>" & LitAgrTalla & "</b>"
		    	     	end if
		    	     	if ver_agrcolor="on" then
		    	     	DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,"<b>" & LitAgrColor & "</b>"
		    	     	end if
		    	     	suma_colspan=3
		    	     	if ver_coste="on" then
		    	     	    suma_colspan=suma_colspan+1
		    	     	end if
		    	     	if ver_precio="on" then
		    	     	    suma_colspan=suma_colspan+1
		    	     	end if
		    	     	if ver_agrupa="on" then
		    	     	    suma_colspan=suma_colspan+1
		    	     	end if
		    	     	DrawCelda "CELDA align=center style='width:10%;border-bottom:1px solid black;border-right:1px solid black;border-top:1px solid black;' colspan='" & suma_colspan & "'","","",0,"<b>" & LitVariable & "</b>"
			  		CloseFila
				end if

				''DrawFila color_blau
					if tipo_art="PRINCIPAL" then
					    
					    %><tr bgcolor="<%=color_blau%>" style="border:1px solid black;"><%
					    ''DrawFila color_blau
						DrawCelda "CELDA style='width:20%;border-bottom:1px solid black;border-left:1px solid black;border-top:1px solid black;' colspan='2'","","",0,Hiperv(OBJArticulos,rst("articulo"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("articulo")),LitVerArticulo)
						DrawCelda "CELDA style='width:25%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,rst("nombre")

						if rst("familia")="" then
							DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,"&nbsp;"
						else
							DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,iif(rst("familia")="","",enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre", "familias", "codigo='" & rst("familia") & "'", session("backendlistados")))))
						end if

                        if ver_agrtalla="on" then
						    if rst("agrtallas")="" then
							    if rst("talla")="" then
								    DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,"&nbsp;"
							    else
								    DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","tallas","codigo='" & rst("talla") & "'", session("backendlistados"))))
							    end if
						    else
							    DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","agrupa_tallas","codigo='" & rst("agrtallas") & "'", session("backendlistados"))))
						    end if
						end if

                        if ver_agrcolor="on" then
    						if rst("agrcolores")="" then
	    						if rst("color")="" then
		    						DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,"&nbsp;"
			    				else
				    				DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","colores","codigo='" & rst("color") & "'",session("backendlistados"))))
					    		end if
						    else
							    DrawCelda "CELDA style='width:15%;border-bottom:1px solid black;border-top:1px solid black;'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","agrupa_colores","codigo='" & rst("agrcolores") & "'", session("backendlistados"))))
    						end if
    					end if

		    	     	suma_colspan=3
		    	     	if ver_coste="on" then
		    	     	    suma_colspan=suma_colspan+1
		    	     	end if
		    	     	if ver_precio="on" then
		    	     	    suma_colspan=suma_colspan+1
		    	     	end if
		    	     	if ver_agrupa="on" then
		    	     	    suma_colspan=suma_colspan+1
		    	     	end if
		    	     	
						if rst("variable")=true then
							DrawCelda "CELDA align='CENTER' style='width:10%;border-bottom:1px solid black;border-right:1px solid black;border-top:1px solid black;' colspan='" & suma_colspan & "'","","",0,"Si"
						else
							DrawCelda "CELDA align='CENTER' style='width:10%;border-bottom:1px solid black;border-right:1px solid black;border-top:1px solid black;' colspan='" & suma_colspan & "'","","",0,"No"
						end if
					elseif tipo_art="ESCANDALLO" then
					    DrawFila color_blau
						DrawCelda "CELDA7 width='3%'","","",0,"&nbsp;"
						DrawCelda "CELDA7 width='10%'","","",0,Hiperv(OBJArticulos,rst("articulo_esc"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("articulo_esc")),LitVerArticulo)
						DrawCelda "CELDA7 width='15%'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("nombre_esc")))
						DrawCelda "CELDA7 width='15%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre","familias","codigo='" & rst("subfamilia") & "'", session("backendlistados"))))
						DrawCelda "'CELDAR7' width='5%'","","",0,formatnumber(rst("cantidad"),DEC_CANT,-1,0,-1)
						if si_tiene_modulo_importaciones=0 then
							DrawCelda "CELDA7 width='13%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","almacenes","codigo='" & rst("almacen") & "'", session("backendlistados"))))
						end if
						'DrawCelda "CELDAR7 width='5%'","","",0,rst("coste")
						if ver_coste="on" then
						DrawCelda "'CELDAR7' width='5%'","","",0,formatnumber(rst("coste"),DEC_PREC,-1,0,-1)
						end if
						if ver_precio="on" then
						DrawCelda "'CELDAR7' width='5%'","","",0,formatnumber(rst("precio"),DEC_PREC,-1,0,-1)
						end if
						if ver_agrtalla="on" then
    						if rst("agrtallas_esc")="" then
	    						if rst("talla_esc")="" then
		    						DrawCelda "CELDA7 width='10%'","","",0,"&nbsp;"
			    				else
				    				DrawCelda "CELDA7 width='10%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","tallas","codigo='" & rst("talla_esc") & "'", session("backendlistados"))))
					    		end if
						    else
							    DrawCelda "CELDA7 width='10%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","agrupa_tallas","codigo='" & rst("agrtallas_esc") & "'", session("backendlistados"))))
						    end if
						end if

                        if ver_agrcolor="on" then
    						if rst("agrcolores_esc")="" then
	    						if rst("color_esc")="" then
		    						DrawCelda "CELDA7 width='10%'","","",0,"&nbsp;"
			    				else
				    				DrawCelda "CELDA7 width='10%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","colores","codigo='" & rst("color_esc") & "'",session("backendlistados"))))
					        	end if
    						else
	    						DrawCelda "CELDA7 width='10%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","agrupa_colores","codigo='" & rst("agrcolores_esc") & "'", session("backendlistados"))))
		    				end if
		    			end if
		    			if ver_agrupa="on" then
						DrawCelda "CELDA7 width='10%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre","agrupaciones","codigo='" & rst("agrupacion") & "'", session("backendlistados"))))
					    end if
						if rst("tallacolor")=true then
							DrawCelda "CELDAC7 width='5%'","","",0,"Sí"
						else
							DrawCelda "CELDAC7 width='5%'","","",0,"No"
						end if
					elseif tipo_art="PARAMETRO" then
					    DrawFila color_blau
					    DrawCelda "CELDA7 width='3%'","","",0,"&nbsp;"
						DrawCelda "CELDA7 width='10%'","","",0,"&nbsp;"
						DrawCelda "CELDA7 width='15%'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("parametro")))
						DrawCelda "CELDA7 width='20%' colspan='2'","","",0,rst("valor")
						DrawCelda "CELDA7 width='13%'","","",0,enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion","medidas","codigo='" & rst("medida") & "'", session("backendlistados"))))
						if rst("Dep_Precio")=true then
							DrawCelda "CELDAC7 width='5%'","","",0,"Sí"
						else
							DrawCelda "CELDAC7 width='5%'","","",0,"No"
						end if
					end if
				CloseFila
			articulo_ant=rst("articulo")
			subfamilia_ant=rst("subfamilia")
			articulo_esc_ant=rst("articulo_esc")
			familia_ant=rst("familia")
			tipo_art_Ant=tipo_art
			fila=fila+1
			rst.MoveNext
			paso = false
    		if agrupar="SUBFAMILIA" then
    		    if rst.EOF then
    		        paso = true
    		    elseif not isnull(articulo_esc_ant) and subfamilia_ant&"" <> rst("subfamilia")&"" then
    		        paso=true
    		    end if
    		    if paso then
    		        DrawFila color_terra
	                if ver_coste="on" then
	                    rstAux.cursorlocation=3
	                    if subfamilia_ant&"" <> "" then
	                        rstAux.cursorlocation=3
		                    rstAux.open "SELECT sum(isnull(cast(coste as real), 0)) as coste from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "' and subfamilia='" & subfamilia_ant & "' and articulo_esc is not null",session("backendlistados")
		                else
		                    rstAux.cursorlocation=3
		                    rstAux.open "SELECT sum(isnull(cast(coste as real), 0)) as coste from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "' and subfamilia is null and articulo_esc is not null",session("backendlistados")
		                end if
		                if not rstAux.eof then
		                    total_coste = rstAux("coste")
		                else
		                    total_coste=0
                        end if
		                if total_coste & "" = "" then total_coste=0
		                rstAux.close
                    end if

		            rstAux.cursorlocation=3
		            if subfamilia_ant&"" <> "" then
		                rstAux.cursorlocation=3
			            rstAux.open "SELECT sum(isnull(cast(cantidad as real), 0)) as cantidad from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "' and subfamilia='" & subfamilia_ant & "' and articulo_esc is not null",session("backendlistados")
			        else
			            rstAux.cursorlocation=3
	                    rstAux.open "SELECT sum(isnull(cast(cantidad as real), 0)) as cantidad from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "' and subfamilia is null and articulo_esc is not null",session("backendlistados")
	                end if
	                if not rstAux.eof then
		                total_cantidad = rstAux("cantidad")
		            else
		                total_cantidad=0
		            end if
		            if total_cantidad & "" = "" then total_cantidad=0
		            rstAux.close

		            if ver_precio="on" then
			            rstAux.cursorlocation=3
			            if subfamilia_ant&"" <> "" then
			                rstAux.cursorlocation=3
			                rstAux.open "SELECT sum(isnull(cast(precio as real), 0)) as precio from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "' and subfamilia='" & subfamilia_ant & "' and articulo_esc is not null",session("backendlistados")
			            else
			                rstAux.cursorlocation=3
	                        rstAux.open "SELECT sum(isnull(cast(precio as real), 0)) as precio from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "' and subfamilia is null and articulo_esc is not null",session("backendlistados")
	                    end if
	                    if not rstAux.eof then
			                total_precio = rstAux("precio")
			            else
			                total_precio=0
			            end if
			            if total_precio & "" = "" then total_precio=0
			            rstAux.close
		            end if

		            DrawFila color_terra
		                DrawCelda "CELDA7 width='3%'","","",0,"&nbsp;"
		                DrawCelda "CELDA7 width='40%' colspan='3'","","",0,"<b>" & LitTotales & " " & LitSubFam & " " & enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre", "familias", "codigo like '"&session("ncliente")&"%' and codigo='"&subfamilia_ant&"'", session("dsn_cliente")))) & ":</b>"
		                DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_cantidad,DEC_CANT,-1,0,-1) & "</b>"
		                if si_tiene_modulo_importaciones=0 then
						    DrawCelda "CELDA7 width='13%'","","",0,"&nbsp;"
					    end if
			            if ver_coste="on" then
			            DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_coste,DEC_PREC,-1,0,-1) & "</b>"
			            end if
						if ver_precio="on" then
		                DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_precio,DEC_PREC,-1,0,-1) & "</b>"
		                end if
	    	     	    suma_colspan=1
	    	     	    if ver_agrupa="on" then
	    	     	        suma_colspan=suma_colspan+1
	    	     	    end if
	    	     	    if ver_agrtalla="on" then
	    	     	        suma_colspan=suma_colspan+1
	    	     	    end if
	    	     	    if ver_agrcolor="on" then
	    	     	        suma_colspan=suma_colspan+1
	    	     	    end if
		                DrawCelda "CELDA7 width='5%' colspan='" & suma_colspan & "'","","",0,"&nbsp;"
		            CloseFila
		        end if
		    end if
		wend

		if rst.eof then
			if ver_coste="on" then
	            rstAux.cursorlocation=3
		        rstAux.open "SELECT sum(isnull(cast(coste as real), 0)) as coste from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "'",session("backendlistados")
		        if not rstAux.eof then
		            total_coste = null_z(rstAux("coste"))
		        else
		            total_coste=0
		        end if
		        rstAux.close
            end if

			rstAux.cursorlocation=3
			rstAux.open "SELECT sum(isnull(cast(cantidad as real), 0)) as cantidad from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "'",session("backendlistados")
			if not rstAux.eof then
			    total_cantidad = null_z(rstAux("cantidad"))
			else
			    total_cantidad=0
			end if
			rstAux.close

            if ver_precio="on" then
				rstAux.cursorlocation=3
				rstAux.open "SELECT sum(isnull(cast(precio as real), 0)) as precio from egesticet.[" & session("usuario") & "] where articulo='" & articulo_ant & "'",session("backendlistados")
				if not rstAux.eof then
				    total_precio = null_z(rstAux("precio"))
				else
				    total_precio=0
				end if
    			rstAux.close
    		end if
			    DrawFila color_fondo
				    DrawCelda "CELDA7 width='3%'","","",0,"&nbsp;"
				    if si_tiene_modulo_importaciones<>0 then
				        DrawCelda "CELDA width='40%' colspan='3'","","",0,"<b>" & LitTotalCar & ":</b>"
        		    else
				        DrawCelda "CELDA width='40%' colspan='3'","","",0,"<b>" & LitTotalEsc & ":</b>"
				    end if
				        DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_cantidad,DEC_CANT,-1,0,-1) & "</b>"
				    if si_tiene_modulo_importaciones=0 then
				        DrawCelda "CELDA7 width='13%'","","",0,"&nbsp;"
			        end if
				    if ver_coste="on" then
				    DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_coste,DEC_PREC,-1,0,-1) & "</b>"
				    end if
				    if ver_precio="on" then
				    DrawCelda "'CELDAR7' width='5%'","","",0,"<b>" & formatnumber(total_precio,DEC_PREC,-1,0,-1) & "</b>"
				    end if
    	     	    suma_colspan=1
    	     	    if ver_agrupa="on" then
    	     	        suma_colspan=suma_colspan+1
    	     	    end if
    	     	    if ver_agrtalla="on" then
    	     	        suma_colspan=suma_colspan+1
    	     	    end if
    	     	    if ver_agrcolor="on" then
    	     	        suma_colspan=suma_colspan+1
    	     	    end if
				    DrawCelda "CELDA7 width='5%' colspan='" & suma_colspan & "'","","",0,"&nbsp;"
				CloseFila
        end if

		rst.close%>
		<tr><td><br /></td></tr>
		<%NavPaginas lote,lotes,campo,criterio,texto,2
	else%>
		<font class='CEROFILAS'><%=LitCeroFilas%></font>
		<%end if
   end if%>
</form>
<%
set conn =nothing
set rst=nothing
set rst1=nothing
set rstAux=nothing
set rstAux2=nothing
end if
%>
</body>
</html>