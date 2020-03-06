<%@ Language=VBScript %>
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
<!--#include file="codigo_barras.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>

<link rel="stylesheet" href="../pantalla.css" media="SCREEN"/>
<link rel="stylesheet" href="../impresora.css" media="PRINT"/>
</head>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function%> 

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    function ValidarCampos() {
        ok=1;

        cantHMax=0;
        cantVMax=0;
        maxpagina=0;

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

        if (ok==1)
            return true;
        else
            return false;

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
                Ven=AbrirVentana("./listaasignacionmasiva.asp","P",<%=AltoVentana%>,<%=AnchoVentana%>);
            }
        }

        if (document.codigo_barras.numdoc.value!="" || document.codigo_barras.tipodoc.value!=""){
            document.codigo_barras.cant_doc.checked=true;
            document.codigo_barras.cantidad.value="";
            document.codigo_barras.cantidad.disabled=true;
            document.getElementById("idcant_doc2").style.display="";
            document.getElementById("id_fmpc").style.display="none";
        }
        else{
            document.codigo_barras.cantidad.value="1";
            document.codigo_barras.cantidad.disabled=false;
            document.codigo_barras.cant_doc.checked=false;
            document.getElementById("idcant_doc2").style.display="none";
            document.getElementById("id_fmpc").style.display="";
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
        if (document.codigo_barras.formato_impresion.value=="listado_codigo_barras4.asp"){
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
            if (document.codigo_barras.ver_empresa.disabled==true){
                document.codigo_barras.ver_empresa.disabled=false;
                document.codigo_barras.ver_empresa.checked=true;
            }
            if (document.codigo_barras.si_tiene_modulo_terminales.value==1){
                if (document.codigo_barras.ver_codTerminal.disabled==true){
                    document.codigo_barras.ver_codTerminal.disabled=false;
                }
            }
            if (document.codigo_barras.formato_impresion.value=="../custom/listado_codigo_barras.asp"){
                document.getElementById("IMPORTESADICIONALES").style.display="";
            } else {
                document.getElementById("IMPORTESADICIONALES").style.display="none";
            }
        }
    }

    function MuestraFechaMod(obj) {
        if (obj.checked) {
            document.codigo_barras.fmpc.value="";
            document.codigo_barras.fmpc.disabled=true;
        } else {
            document.codigo_barras.fmpc.value="";
            document.codigo_barras.fmpc.disabled=false;
        }
    }

</script>
<body onload="self.status='';" class="BODY_ASP">
<%

''*****************************************************************************
''********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
''*****************************************************************************
const borde=0


%><form name="codigo_barras" method="post"><%
		PintarCabecera "codigo_barras.asp"

		si_tiene_modulo_terminales=ModuloContratado(session("ncliente"),ModTerminales)

		set rstAux = Server.CreateObject("ADODB.Recordset")
		set rst = Server.CreateObject("ADODB.Recordset")
		set rstSelect = Server.CreateObject("ADODB.Recordset")


		''Leer parámetros de la página
		mode=enc.EncodeForJavascript(Request.QueryString("mode"))

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

		ObtenerParametros "codigos_barras"

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

		nombre	= limpiaCadena(request.form("nombre"))
		referencia	= limpiaCadena(request.form("referencia"))

		familia	= limpiaCadena(request.form("familia"))
		ordenar	= limpiaCadena(request.form("ordenar"))

		tipodoc	= limpiaCadena(request.form("tipodoc"))
		numdoc	= limpiaCadena(request.form("numdoc"))

		ver_referencia	=	limpiaCadena(request.form("ver_referencia"))
		ver_nombre		=	limpiaCadena(request.form("ver_nombre"))
		ver_empresa		=	limpiaCadena(request.form("ver_empresa"))
		ver_lineas		=	limpiaCadena(request.form("ver_lineas"))
		ver_precios		=	limpiaCadena(request.form("ver_precios"))
		ver_codTerminal	=	limpiaCadena(request.form("ver_codTerminal"))
		imprimir_listado_horizontal	=	limpiaCadena(request.form("imprimir_listado_horizontal"))
		imprimir_listado_vertical	=	limpiaCadena(request.form("imprimir_listado_vertical"))
		formato_impresion			=	limpiaCadena(request.form("formato_impresion"))
		cant_doc					=	limpiaCadena(request.form("cant_doc"))

		fechamodprec=limpiaCadena(request.form("fmpc"))
		solopreciocambiado=limpiaCadena(request.form("solopreciocambiado"))
		opcprec1=limpiaCadena(request.form("opcprec1"))
		opcprec2=limpiaCadena(request.form("opcprec2"))
		tarifa1=limpiaCadena(request.form("tarifa1"))
		tarifa2=limpiaCadena(request.form("tarifa2"))
		tarifaex=limpiaCadena(request.form("tarifaex"))
		tarifaiva1=limpiaCadena(request.form("tarifaiva1"))
		tarifaiva2=limpiaCadena(request.form("tarifaiva2"))

		Alarma "codigo_barras.asp"
		%><input type="hidden" name="si_tiene_modulo_terminales" value="<%=EncodeForHtml(si_tiene_modulo_terminales)%>"/><%
		%><br/><%

	if ucase(mode) = "EXPORTAR" then
		%><input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>"/>
		<input type="hidden" name="nombre" value="<%=EncodeForHtml(nombre)%>"/>
		<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>"/>
		<input type="hidden" name="ordenar" value="<%=EncodeForHtml(ordenar)%>"/>
		<input type="hidden" name="cantidad" value="<%=EncodeForHtml(cantidad)%>"/>
		<input type="hidden" name="tipodoc" value="<%=EncodeForHtml(tipodoc)%>"/>
		<input type="hidden" name="numdoc" value="<%=EncodeForHtml(numdoc)%>"/>
		<input type="hidden" name="ver_referencia" value="<%=EncodeForHtml(ver_referencia)%>"/>
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
		<input type="hidden" name="fmpc" value="<%=EncodeForHtml(fechamodprec)%>"/>
		<input type="hidden" name="solopreciocambiado" value="<%=EncodeForHtml(solopreciocambiado)%>"/>

		<input type="hidden" name="opcprec1" value="<%=EncodeForHtml(opcprec1)%>"/>
		<input type="hidden" name="opcprec2" value="<%=EncodeForHtml(opcprec2)%>"/>
		<input type="hidden" name="tarifa1" value="<%=EncodeForHtml(tarifa1)%>"/>
		<input type="hidden" name="tarifa2" value="<%=EncodeForHtml(tarifa2)%>"/>
		<input type="hidden" name="tarifaex" value="<%=EncodeForHtml(tarifaex)%>"/>
		<input type="hidden" name="tarifaiva1" value="<%=EncodeForHtml(tarifaiva1)%>"/>
		<input type="hidden" name="tarifaiva2" value="<%=EncodeForHtml(tarifaiva2)%>"/>

		<script language="javascript" type="text/javascript">
		    document.codigo_barras.action = "exportar_codigo_barras.asp";
		    document.codigo_barras.submit();
		</script><%
	end if

	''*********************************************************************************************
	''Se muestran parametros de seleccion
	''*********************************************************************************************
	if mode="param" then
	    DrawDiv "1", "", ""
            DrawLabel "", "", LitConref%><input class="width40" type="text" name="referencia" size="25" onchange="TraerArticulo('<%=enc.EncodeForJavascript(mode)%>','<%=enc.EncodeForJavascript(ndet)%>');"/><a class="CELDAREFB" href="javascript:WinArticulos()"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
        <%CloseDiv
        EligeCelda "input", "add", "", "", "", 0, LitConNombre, "nombre", "", nombre
        
        DrawDiv "1", "", ""
            DrawLabel "", "", LitOrdenar2%><select class="width60" name="ordenar" >
					<option selected="selected" value="REFERENCIA"><%=ucase(LitRef)%></option>
					<option value="NOMBRE"><%=ucase(LitNombre2)%></option>
				</select><%
        CloseDiv

        rstAux.cursorlocation=3
		rstAux.open " select codigo, nombre from familias with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre", session("backendlistados")
		DrawSelectCelda "CELDA","190","",0,LitSubFamilia,"familia",rstAux,familia,"codigo","nombre","",""
		rstAux.close
'----------------------------------------------------------------------------------------
'Nueva forma de obtener la descripcion de los tipos de documentos de la tabla lit_typedoc
'----------------------------------------------------------------------------------------
inList = "'ALBARAN DE SALIDA','ALBARAN DE PROVEEDOR','FACTURA A CLIENTE','FACTURA DE PROVEEDOR','HOJA DE GASTOS','PEDIDO A PROVEEDOR','PEDIDO DE CLIENTE','PRESUPUESTO A CLIENTE'"
if session("ncliente")="00012" or session("ncliente")="00182" or session("ncliente")="00180" then
    inList=inList & ",'MOVIMIENTOS ENTRE ALMACENES'"
end if
addList = "'ASIGNACION MASIVA'"
set conn = Server.CreateObject("ADODB.Connection")
set command =  Server.CreateObject("ADODB.Command")
conn.open DSNIlion
command.ActiveConnection = conn
command.CommandTimeout = 0
command.CommandText = "ComboBoxDocTypes"
command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
command.NamedParameters = True 
command.Parameters.Append command.CreateParameter("@inlist",adVarChar,adParamInput,len(inList),inList)
command.Parameters.Append command.CreateParameter("@outlist",adVarChar,adParamInput,1,"")
command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,session("usuario"))
command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,50,session("ncliente"))
command.Parameters.Append command.CreateParameter("@addlist",adVarChar,adParamInput,len(addList),addList)
set rstTD = Server.CreateObject("ADODB.Recordset")
set rstTD = command.Execute
if not rstTD.eof then
		DrawSelectCelda "CELDA","190","","0",Littipodoc,"tipodoc",rstTD,tipodoc,"tippdoc","descripcion","onchange","mostrar_cantidad(0)"
end if	
rstTD.close
conn.close
set rstTD =nothing
set conn =nothing
set command=nothing

DrawDiv "1", "", ""
    DrawLabel "", "", LitNumDoc%><input type="text" name="numdoc" size="25" maxlength="20" value="<%=EncodeForHtml(numdoc)%>" onkeyup="javascript:mostrar_cantidad(1)" onmousedown="javascript:mostrar_cantidad(2)" onblur="javascript:mostrar_cantidad(3)"/>
	<%
CloseDiv

if cantidad="" then cantidad="1"

EligeCelda "input", "add", "", "", "", 0, LitCantidad2, "cantidad", "", cantidad            
    %><span id="idcant_doc2" style='display:none'><%
        DrawDiv "1", "", ""
        DrawLabel "", "", LitCodBarrImpCantDoc%>
<input type="checkbox" name="cant_doc" onclick="javascript:control_cantidad()"/>
<%CloseDiv%> 
</span>
<%

    		''ricardo 19-7-2005 se pone avisarcambioprecio=0 para pasar a Terra
			''avisarcambioprecio=nz_b(d_lookup("avisarcambioprecio","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")))
			avisarcambioprecio=1
			''''''''''''''''''''''ricardo 19-7-2005''''''''''''''''''''''''''''''''''''''''''''''
			if avisarcambioprecio<>0 then
                DrawDiv "1", "", ""
	                DrawLabel "", "", LitSoloPrecioCambiado%><input type='checkbox' name='solopreciocambiado' onclick="MuestraFechaMod(this);"/><%
                CloseDiv
			end if
		
			%>
			<span id="id_fmpc"><%
                DrawDiv "1", "", ""
				if fye="0" then
					DrawLabel "", "", LitArtModPrecFecPos%><input type="text"  size="13" maxlength="10" name="fmpc" value=""/>
					<%
				else
					DrawLabel "", "", LitArtModPrecFecPos%><input type="text"  size="13" maxlength="10" name="fmpc" value="<%=EncodeForHtml(iif(fechamodprec>"",fechamodprec,date-1))%>"/>
					<%
				end if
				CloseDiv%>
			</span>
	<hr/>
	<%
        EligeCelda "check", "add", "", "", "", 0, LitRef, "ver_referencia", "", "True"
        EligeCelda "check", "add", "", "", "", 0, LitNombre2, "ver_nombre", "", "True"
		if fye="0" then
            EligeCelda "check", "add", "", "", "", 0, LitEmpresa, "ver_empresa", "", "False"    
        else
            EligeCelda "check", "add", "", "", "", 0, LitEmpresa, "ver_empresa", "", "True"
        end if
        EligeCelda "check", "add", "", "", "", 0, LitLineas, "ver_lineas", "", "True"
        EligeCelda "check", "add", "", "", "", 0, LitPrecios2, "ver_precios", "", "True"
		if si_tiene_modulo_terminales<>0 then
            EligeCelda "check", "add", "", "", "", 0, LitCodTerminal, "ver_codTerminal", "", "False"
		end if
	%>
	<span id="IMPORTESADICIONALES" style="display:none" >
	<hr/>
    <%
		DrawDiv "col-lg-6 col-md-12 col-xs-12 col-sm-12", "", ""
			DrawLabel "", "", litOpcImporteAdicional1
		CloseDiv
		DrawDiv "col-lg-6 col-md-12 col-xs-12 col-sm-12", "", ""
			DrawLabel "", "", litOpcImporteAdicional2
		CloseDiv
        
		rstAux.cursorlocation=3
		rstAux.open " select codigo, descripcion from tarifas with(NOLOCK) where codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", session("backendlistados")
		haytarifas=0

		if not rstAux.eof then haytarifas=1
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4", "", ""%><input type="radio" name="opcprec1" value="tarifa"/><%CloseDiv    
            DrawSelectCelda "","","",0,LitTarifaListadoCodigoBarras,"tarifa1",rstAux,tarifa1,"codigo","descripcion","",""
        %></div><%

		if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
		    DrawDiv "4", "", ""%><input type="radio" name="opcprec2" value="tarifa"/><%CloseDiv
            DrawSelectCelda "","","",0,LitTarifaListadoCodigoBarras,"tarifa2",rstAux,tarifa2,"codigo","descripcion","",""
        %></div><%

		if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4", "", ""%><input type="radio" name="opcprec1" value="tarifaiva"/><%CloseDiv    
            DrawSelectCelda "","","",0,LitTarifaIvaListadoCodigoBarras,"tarifaiva1",rstAux,tarifaiva1,"codigo","descripcion","",""
        %></div><%

		if haytarifas=1 then rstaux.movefirst
        %><div class="col-lg-6 col-md-12 col-xs-12 col-sm-12"><%
            DrawDiv "4", "", ""%><input type="radio" name="opcprec2" value="tarifaiva"/><%CloseDiv    
            DrawSelectCelda "","","",0,LitTarifaIvaListadoCodigoBarras,"tarifaiva2",rstAux,tarifaiva2,"codigo","descripcion","",""
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
	</span>
	<hr/>
	<%
	    EligeCelda "input", "add", "", "", "", 0, LitMsgHorizontal, "imprimir_listado_horizontal", "", "1"
		EligeCelda "input", "add", "", "", "", 0, LitMsgVertical, "imprimir_listado_vertical", "", "1"	
''ricardo 19-7-2005 se pone pasar_a_terra=0 para pasar a Terra
if pasar_a_terra=0 then
            rstAux.cursorlocation=3
			rstAux.open " select codigo, case when codigo='" & session("ncliente") & "BASE' then 'PRECIO FICHA ARTICULO' else descripcion end as desccrip from tarifas with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by desccrip", session("backendlistados")
			tarifaex = session("ncliente") & "BASE"
            DrawDiv "1", "", ""%><label><%
            DrawHref "CELDAREFB", "", LitExportar, "javascript:if(ValidarCampos()){codigo_barras.action='codigo_barras.asp?mode=exportar';codigo_barras.submit();parent.botones.document.location='codigo_barras_bt.asp?mode=exportar';}"%></label><%
			DrawSelect "100","","tarifaex",rstAux,tarifaex,"codigo","desccrip","",""
			rstAux.close
            CloseDiv
else
	%><%
end if ''ricardo 19-7-2005
	%>
	<hr/>
	
		
			<input type="hidden" name="fye" value="<%=EncodeForHtml(fye)%>"/>
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

			<input type="hidden" name="maxpagina_form16" value="<%=EncodeForHtml(cuantasEtiqPorPag1MargenG)%>"/>
			<input type="hidden" name="cantHMax_form16" value="<%=EncodeForHtml(AnchoForm16)%>"/>
			<input type="hidden" name="cantVMax_form16" value="<%=EncodeForHtml(AltoForm16)%>"/>

			<input type="hidden" name="maxpagina_form17" value="<%=EncodeForHtml(cuantasEtiqPorPag1MargenP)%>"/>
			<input type="hidden" name="cantHMax_form17" value="<%=EncodeForHtml(AnchoForm17)%>"/>
			<input type="hidden" name="cantVMax_form17" value="<%=EncodeForHtml(AltoForm17)%>"/>

			<input type="hidden" name="maxpagina_form5ALHILO" value="<%=EncodeForHtml(cuantasEtiqPorPag40)%>"/>
			<input type="hidden" name="cantHMax_form5ALHILO" value="<%=EncodeForHtml(AnchoForm5)%>"/>
			<input type="hidden" name="cantVMax_form5ALHILO" value="<%=EncodeForHtml(AltoForm5)%>"/>

            <input type="hidden" name="maxpagina_formSolred3x7" value="<%=EncodeForHtml(cuantasEtiqPorPag21)%>" />
            <input type="hidden" name="cantHMax_formSolred3x7" value="<%=EncodeForHtml(AnchoFormSolred3x7)%>" />
            <input type="hidden" name="cantVMax_formSolred3x7" value="<%=EncodeForHtml(AltoFormSolred3x7)%>" />



<%
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

		%><input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>"/>
		<input type="hidden" name="nombre" value="<%=EncodeForHtml(nombre)%>"/>
		<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>"/>
		<input type="hidden" name="ordenar" value="<%=EncodeForHtml(ordenar)%>"/>
		<input type="hidden" name="cantidad" value="<%=EncodeForHtml(cantidad)%>"/>
		<input type="hidden" name="tipodoc" value="<%=EncodeForHtml(tipodoc)%>"/>
		<input type="hidden" name="numdoc" value="<%=EncodeForHtml(numdoc)%>"/>
		<input type="hidden" name="ver_referencia" value="<%=EncodeForHtml(ver_referencia)%>"/>
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
		<input type="hidden" name="fmpc" value="<%=EncodeForHtml(fechamodprec)%>"/>
		<input type="hidden" name="solopreciocambiado" value="<%=EncodeForHtml(solopreciocambiado)%>"/>

		<input type="hidden" name="opcprec1" value="<%=EncodeForHtml(opcprec1)%>"/>
		<input type="hidden" name="opcprec2" value="<%=EncodeForHtml(opcprec2)%>"/>
		<input type="hidden" name="tarifa1" value="<%=EncodeForHtml(tarifa1)%>"/>
		<input type="hidden" name="tarifa2" value="<%=EncodeForHtml(tarifa2)%>"/>
		<input type="hidden" name="tarifaex" value="<%=EncodeForHtml(tarifaex)%>"/>
		<input type="hidden" name="tarifaiva1" value="<%=EncodeForHtml(tarifaiva1)%>"/>
		<input type="hidden" name="tarifaiva2" value="<%=EncodeForHtml(tarifaiva2)%>"/>



		<%'JMA 30/10/05: Quitamos de la condicion el formato "listado_codigo_barras.asp"'
		if formato_impresion="listado_codigo_barras2.asp" or formato_impresion="listado_codigo_barras4.asp" then
			formato_impresion="..\\productos\\listados\\" & formato_impresion
		end if
		if formato_impresion="..\\..\\custom\\listado_codigo_barras5.asp" then
			formato_impresion="listado_codigo_barras5.asp"
		end if
		if formato_impresion="../custom/listado_codigo_barras.asp" then
			formato_impresion="listado_codigo_barras.asp"
		end if
		if formato_impresion="..\\..\\custom\\listado_codigo_barras5ALHILO.asp" then
			formato_impresion="listado_codigo_barras5ALHILO.asp"
		end if


		%><script language="javascript" type="text/javascript">
		      document.codigo_barras.action = "<%=formato_impresion%>";
		      document.codigo_barras.submit();
		</script><%
	end if
		set rstAux = Nothing
		set rst = Nothing
		set rstSelect = Nothing

	%></form><%

end if
%>
</body>
</html>
