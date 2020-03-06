<%@ Language=VBScript %>

<% dim  enc
    set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<!--#include file="../constantes.inc" -->
<%
' ############ SOLO PARA DEVOLVER CONSULTAS AJAX ############
if request.querystring("mode") = "consultaAJAX" then
    if request.querystring("consulta") = "actualizarGrupos" then
        set rstAux = Server.CreateObject("ADODB.Recordset")
        sql = "EXEC ActualizarGruposEnFranquicias @nempresaCentral='" & session("ncliente") & "'"	
		rstAux.open sql,DSNImport
        if rstAux("devolver") = "0" then
            response.Write("OK")
        else
            response.Write("ERROR")
        end if
        rstAux.close
    end if
   
    ' Fin de consulta AJAX
    response.End
end if
' ################################################
'' JCI 18/06/2003 : MIGRACION A MONOBASE
'RGU 13/10/2006: Añadir campo pvp+iva en el span de precios
%>
<%response.buffer=true%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
    

<!--#include file="../mensajes.inc" -->
<!--#include file="grupos_ofertas.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../CatFamSubResponsive.inc" -->
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../varios.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/tabs.js.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">

animatedcollapse.addDiv('CABECERA', 'fade=1')
animatedcollapse.addDiv('ARTICULOS', 'fade=1')

animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
    //$: Access to jQuery
    //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
    //state: "block" or "none", depending on state
}

animatedcollapse.init()

if (window.document.addEventListener) {
    window.document.addEventListener("keydown", callkeydownhandler, false);
} else {
    window.document.attachEvent("onkeydown", callkeydownhandler);
}

function callkeydownhandler(evnt) {
    ev = (evnt) ? evnt : event;
    tecla_pulsada=enc.EncodeForJavascript(ev.keyCode);
}

function CampoRefPulsado(mode,marco,formulario,queordenar,comoordenar){
    if (tecla_pulsada==13){
        continuar=0;
        if (mode=="ALTA"){
            if (document.grupos_ofertas.RefPro.value!="") continuar=1;
        }
        if (mode=="BAJA"){
            if (document.grupos_ofertas.bRefPro.value!="") continuar=1;
        }
        if (continuar==1) Insertar(mode,'1','insertar',queordenar,comoordenar);
    }
}

function OrdenarDatos(mode,marco,formulario,campo){
    campo=campo.toUpperCase();
    eval("queordenar=" + marco + ".document." + formulario + ".queordenar.value.toUpperCase()");
    eval("comoordenar=" + marco + ".document." + formulario + ".comoordenar.value.toUpperCase()");
    if (campo!=queordenar || comoordenar==""){
        queordenar=campo;
        comoordenar="ASC";
    }
    else{
        if (campo==queordenar && comoordenar=="ASC") comoordenar="DESC";
        else comoordenar="ASC";
    }
    queimagen1="";
    queimagen2="";
    queimagen3="";
    comoimagen="";
    if (comoordenar=="ASC") comoimagen="&darr;";
    if (comoordenar=="DESC") comoimagen="&uarr;";
    if(queordenar=="A.REFERENCIA"){
        queimagen1=comoimagen;
        queimagen2="&harr;";
        queimagen3="&harr;";
    }
    if(queordenar=="A.NOMBRE"){
        queimagen2=comoimagen;
        queimagen1="&harr;";
        queimagen3="&harr;";
    }
    if(queordenar=="F.NOMBRE"){
        queimagen3=comoimagen;
        queimagen2="&harr;";
        queimagen1="&harr;";
    }
    if (mode=="ALTA"){
        document.getElementById("OD1A").innerHTML=queimagen1;
        document.getElementById("OD2A").innerHTML=queimagen2;
        document.getElementById("OD3A").innerHTML=queimagen3;
    }
    else{
        document.getElementById("OD1B").innerHTML=queimagen1;
        document.getElementById("OD2B").innerHTML=queimagen2;
        document.getElementById("OD3B").innerHTML=queimagen3;
    }
    Insertar(mode,'1','first',queordenar,comoordenar);
}

function Insertar(mode, pag, sentido, queordenar, comoordenar) {
    var familia = "";
    var categoria = "";
    var subfamilia = "";
	switch (mode) {
	    case "ALTA":
	        mod = "save";
	        if (sentido == "first") {
	            sentido = "&submode=first";
	            mod = "first";
	            document.grupos_ofertas.condbase.value = "0";
	        }

	        if (sentido == "insertar") {
	            sentido = "&submode=insertar";
	            mod = "insertar";
	        }

	        for (i = 0; i < document.grupos_ofertas.familia.options.length; i++) {
	            if (document.grupos_ofertas.familia.options[i].selected)
	                subfamilia = subfamilia + document.grupos_ofertas.familia.options[i].value + ",";
	        }
	        if (subfamilia != "" && subfamilia != "undefined") subfamilia = subfamilia.substring(0, subfamilia.length - 1);
	        else subfamilia = document.grupos_ofertas.familia.value;
	        for (i = 0; i < document.grupos_ofertas.familia_padre.options.length; i++) {
	            if (document.grupos_ofertas.familia_padre.options[i].selected)
	                familia = familia + document.grupos_ofertas.familia_padre.options[i].value + ",";
	        }
	        if (familia != "" && familia != "undefined") familia = familia.substring(0, familia.length - 1);
	        else familia = document.grupos_ofertas.familia_padre.value;
	        for (i = 0; i < document.grupos_ofertas.categoria.options.length; i++) {
	            if (document.grupos_ofertas.categoria.options[i].selected)
	                categoria = categoria + document.grupos_ofertas.categoria.options[i].value + ",";
	        }
	        if (categoria != "" && categoria != "undefined") categoria = categoria.substring(0, categoria.length - 1);
	        else categoria = document.grupos_ofertas.categoria.value;

	        pagina = "ArticulosDeGrupo.asp?mode=" + mod + "&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.grupos_ofertas.htarifa.value
			+ "&ref=" + document.grupos_ofertas.refcontiene.value + "&familia=" + subfamilia + "&categoria=" + categoria + "&familia_padre=" + familia +
            "&tipoarticulo=" + document.grupos_ofertas.tipoarticulo.value + "&desc=" + document.grupos_ofertas.descontiene.value + "&nproveedor="
			+ document.grupos_ofertas.proveedor.value + "&queordenar=" + queordenar + "&comoordenar=" + comoordenar
			+ "&rangodesde=" + document.grupos_ofertas.rangodesde.value + "&rangohasta=" + document.grupos_ofertas.rangohasta.value + "&pvpiva=" + document.grupos_ofertas.pvpiva.checked
			+ "&viene=ALTA";

	        if (mod == "first") {
	            marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
	            document.getElementById("frArticulosAdd").src = pagina;
	        }
	        else {
	            marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
	            marcoArticulosAdd.document.ArticulosDeGrupo.action = pagina;
	            marcoArticulosAdd.document.ArticulosDeGrupo.submit();
	        }
	        document.grupos_ofertas.check.checked = false;
	        break;
		case "BAJA":
			mod="save2";
			if(sentido=="first"){
				sentido="&submode=first";
				mod="first2";
			}
			
			if(sentido=="insertar"){
				sentido="&submode=insertar";
				mod="insertar";
			}
            for (i = 0; i < document.grupos_ofertas.familia1.options.length; i++) {
                if (document.grupos_ofertas.familia1.options[i].selected)
                    subfamilia = subfamilia + document.grupos_ofertas.familia1.options[i].value + ",";
            }
            if (subfamilia != "" && subfamilia != "undefined") subfamilia = subfamilia.substring(0, subfamilia.length - 1);
            else subfamilia = document.grupos_ofertas.familia1.value;
            for (i = 0; i < document.grupos_ofertas.familia_padre1.options.length; i++) {
                if (document.grupos_ofertas.familia_padre1.options[i].selected)
                    familia = familia + document.grupos_ofertas.familia_padre1.options[i].value + ",";
            }
            if (familia != "" && familia != "undefined") familia = familia.substring(0, familia.length - 1);
            else familia = document.grupos_ofertas.familia_padre1.value;
            for (i = 0; i < document.grupos_ofertas.categoria1.options.length; i++) {
                if (document.grupos_ofertas.categoria1.options[i].selected)
                    categoria = categoria + document.grupos_ofertas.categoria1.options[i].value + ",";
            }
            if (categoria != "" && categoria != "undefined") categoria = categoria.substring(0, categoria.length - 1);
            else categoria = document.grupos_ofertas.categoria1.value;

			pagina="ArticulosDeGrupo.asp?mode="+mod+"&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.grupos_ofertas.htarifa.value +
			"&bref=" + document.grupos_ofertas.brefcontiene.value + "&bfamilia=" + subfamilia + "&bcategoria=" + categoria + "&bfamilia_padre="
            + familia + "&btipoarticulo=" + document.grupos_ofertas.btipoarticulo.value + "&bdesc=" +
            document.grupos_ofertas.bdescontiene.value + "&bnproveedor=" + document.grupos_ofertas.bproveedor.value +
            "&queordenar=" + queordenar + "&comoordenar=" + comoordenar + "&brangodesde=" + document.grupos_ofertas.brangodesde.value + "&brangohasta=" 
            + document.grupos_ofertas.brangohasta.value + "&bpvpiva=" + document.grupos_ofertas.bpvpiva.value + "&viene=BAJA";
			if (mod=="first2"){
			    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
				document.getElementById("frArticulosBorrar").src=pagina;
			}else{
                marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
				marcoArticulosBorrar.document.ArticulosDeGrupo.action=pagina;
				marcoArticulosBorrar.document.ArticulosDeGrupo.submit();
			}

			document.grupos_ofertas.checkb.checked = false;
			break;
	}
}

function Editar(p_codigo, p_npagina, p_campo, p_criterio, p_texto) {
	document.location="grupos_ofertas.asp?mode=edit&p_codigo=" + p_codigo +"&npagina="+ p_npagina +"&campo="  + p_campo +"&texto="  + p_texto +"&criterio=" + p_criterio;
	parent.botones.document.location="grupos_ofertas_bt.asp?mode=edit";
}

function GuardarArticulos(mode) {
	if (mode=="ALTA") 
	{
		marcoArticulosAdd.document.ArticulosDeGrupo.action="ArticulosDeGrupo.asp?mode=save&submode=all&tarifa=" + document.grupos_ofertas.htarifa.value
		marcoArticulosAdd.document.ArticulosDeGrupo.submit();
	}
}

function BorrarArticulos() {
	if (confirm("<%=LitMsgEliminarRefTarifaConfirm%>")) {
		marcoArticulosBorrar.document.ArticulosDeGrupo.action="ArticulosDeGrupo.asp?mode=delete&npagina=" + marcoArticulosBorrar.document.ArticulosDeGrupo.hnpagina.value + "&tarifa=" + document.grupos_ofertas.htarifa.value;
		marcoArticulosBorrar.document.ArticulosDeGrupo.submit();
	}
}

function seleccionar(marco,formulario,check) {
	nregistros=eval(marco + ".document." + formulario + ".hNregs.value-1");
	if (eval("document.grupos_ofertas." + check+ ".checked")){
		for (i=1;i<=nregistros;i++) {
			nombre="check" + i;
			eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
		}
	}
	else {
		for (i=1;i<=nregistros;i++) {
			nombre="check" + i;
			eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
		}
	}
}

// ############################ FUNCIONES AJAX ############################

function handleHttpResponse() {
    document.getElementById("waitBoxOculto").style.visibility = "visible";
    if (http.readyState == 4) {
        if (http.status == 200) {
            if (http.responseText.indexOf('invalid') == -1) {
                results = http.responseText;
                enProceso = false;
                if (results != "OK") {
                    alert("<%=LITMSG_GRUPOSACTUALIZADO_ERROR %>");
                }
                else {
                    alert("<%=LITMSG_GRUPOSACTUALIZADO_OK %>");
                }
            }
        }
        else {
            alert("<%=LITMSG_GRUPOSACTUALIZADO_ERROR %>");
        }
        document.getElementById("waitBoxOculto").style.visibility = "hidden";
    }
}

function ActualizarGrupos() {
    if (!enProceso && http) {
        var timestamp = Number(new Date());
        var url = "grupos_ofertas.asp?mode=consultaAJAX&consulta=actualizarGrupos&ts=" + timestamp;
        http.open("GET", url, false);
        http.onreadystatechange = handleHttpResponse;
        enProceso = true;
        http.send(null);
    }
}

function getHTTPObject() {
    var xmlhttp;
    if (!xmlhttp && typeof XMLHttpRequest != 'undefined') {
        try {
            xmlhttp = new XMLHttpRequest();
        }
        catch (e) { xmlhttp = false; }
    }
    return xmlhttp;
}

var enProceso = false; // lo usamos para ver si hay un proceso activo
var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest

// ############################ FIN DE AJAX ############################
</script>
<body class="BODY_ASP" bgcolor=<%=color_blau%>>
<%
themeIlion="/lib/estilos/" & folder & "/"

'******************************************************************************************************************
'                                             FUNCIONES ASP
'******************************************************************************************************************
sub BarraNavegacion()	
    %>
        <script language="javascript" type="text/javascript">
            jQuery("#S_CABECERA").hide();
            jQuery("#ARTICULOS").show();
        </script>
    <%
end sub

'**********************************************************************************************************
sub SpanAltasArticulos(tar)
	dis=""
	'TieneArticulos=d_lookup("referencia","articulos_grupos_oferta","codigo='" & tar & "'",session("dsn_cliente"))&""'

    TieneArticulosSelect= "Select referencia from articulos_grupos_oferta where codigo='"& tar & "'"
    TieneArticulos= DLookupP1(TieneArticulosSelect, tar, adVarChar,10,  session("dsn_cliente"))&""

	if TieneArticulos="" then dis="disabled"
	'Línea para establecer los parámetros de relleno de iframe
    if session("version") <> "5" then%>
	    <div style="border: 1px solid Black;"><%
    else%>
        <div><%
    end if
		'Drawfila ""
        DrawDiv "9", "", ""
            DrawLabel "'CELDAB7'", "", LitRefContiene
            DrawInput "CELDA7", "width:130px","refcontiene","",""      
        CloseDiv
        DrawDiv "9", "", ""
			DrawLabel "'CELDAB7 text-bold'", "", LitDesContiene
            DrawInput "CELDA7", "width:130px","descontiene","",""      
        CloseDiv
			'strselect ="select * from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='ARTICULO' order by descripcion "
            strselect ="select * from tipos_entidades with(nolock) where codigo like ? + '%' and tipo=? order by descripcion "
            set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 =  Server.CreateObject("ADODB.Command")
        conn2.open session("dsn_cliente")
        conn2.cursorlocation=3
        command2.ActiveConnection =conn2
        command2.CommandTimeout = 60
        command2.CommandText=strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,5,session("ncliente")&"")
        command2.Parameters.Append command2.CreateParameter("@tipo",adVarChar,adParamInput,50,"ARTICULO")
                
        set rstAux= command2.Execute
	    

			'rstAux.cursorlocation=3
			'rstAux.open strselect, session("dsn_cliente")
        DrawDiv "9", "", ""
			DrawLabel "'CELDAB7'", "", LitTipoArt%>
		    <select style="display:; width:130px;" class=CELDAL7 name="tipoarticulo">
			    <%if tipoarticulo="" then %>
			    	<option selected value=""> </option>
				<%else%>
				    <option selected value="<%=enc.EncodeForHtmlAttribute(tipoarticulo)%>"> <%=enc.EncodeForHtmlAttribute(trimCodEmpresa(null_s(tipoarticulo)))%></option>
			    	<option value=""> </option>
				<%end if
			 	while not rstAux.eof%>
		   			<option value="<%=enc.EncodeForHtmlAttribute(rstAux("codigo"))%>"><%=enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion")))%></option>
					<%rstAux.movenext
				wend%>                                                                  
		   	</select></td><%
        CloseDiv
		'rstAux.close
        conn2.close
        set conn2    =  nothing
        set command2 =  nothing
        set rstAux  =  nothing	


        %><table></table><%
        
		Drawfila ""
			dim ConfigDespleg (3,13)
			i=0
			ConfigDespleg(i,0)="categoria"
			ConfigDespleg(i,1)="130"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="CELDA7 colspan=2"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)="<b>" & LitCategoria & "</b>"
			ConfigDespleg(i,10)=categoria
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=1
			ConfigDespleg(i,0)="familia_padre"
			ConfigDespleg(i,1)="130"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="CELDA7"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)="<b>" & LitFamilia & "</b>"
			ConfigDespleg(i,10)=familia_padre
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""
 
			i=2
			ConfigDespleg(i,0)="familia"
			ConfigDespleg(i,1)="130"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="CELDA7"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)="<b>" & LitSubFamilia & "</b>"
			ConfigDespleg(i,10)=familia
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			DibujaDesplegablesCustom ConfigDespleg,session("dsn_cliente")
		CloseFila
        %><table></table><%

        DrawDiv "col-xxs-12", "", ""
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", LitImporteRango
            CloseDiv
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", LitDesde
                DrawInput "CELDA7", "width:130px;","rangodesde","",""
            CloseDiv
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", LitHasta
                DrawInput "CELDA7", "width:130px;","rangohasta","",""
            CloseDiv
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", litPvpIva
                %><input class="" type="Checkbox" name="pvpiva"><%
            CloseDiv
        CloseDiv

        DrawDiv "9", "", ""
            DrawLabel "CELDAB7", "", LitProveedor
            'rstAux.open "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like '" & session("ncliente") & "%' order by razon_social",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            strselect= "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like ? + '%' order by razon_social"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,5,session("ncliente")&"")
            
            set rstAux= command2.Execute
		 	
			DrawSelectCeldaDet "","CELDA7 style='width:130px'","",0,"","proveedor",rstAux,"","nproveedor","razon_social","",""
			'rstAux.close
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing	

        CloseDiv

        DrawDiv "col-xxs-12", "text-align: right;", ""
            DrawLabel "CELDAB7", "", LitCargarArticulos%>
			<a class="ic-accept floatRight" href="javascript:if(Insertar('ALTA','1','first','A.referencia','asc'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>"></a>
		<%CloseDiv %>
	</div>
	<table class="width100 md-table-responsive">
	    <%escribe1="&darr;"
	    escribe2="&harr;"
	    escribe3="&harr;"
        colorflecha="blue"
						
		Drawfila color_terra
			%><td class="CELDAC7 underOrange width5" width="25"><input class="" type="Checkbox" name="check" onClick="seleccionar('marcoArticulosAdd','ArticulosDeGrupo','check');" ></td>
			<td class="CELDAC7 underOrange width15">
			    <b><%=LitReferencia%></b>
			    <a CLASS="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>; display: inline-block;" id="OD1A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDeGrupo','A.referencia');" title="<%=LitOrdTarRef & " " & LitOrdSentidoD%>"><b><%=escribe1%></b></a>
			</td>
			<td class="CELDAC7 underOrange width25">
			    <b><%=LitDescripcion%></b>
			    <a CLASS="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>; display: inline-block;" id="OD2A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDeGrupo','A.nombre');" title="<%=LitOrdTarDesc%>" ><b><%=escribe2%></b></a>
			</td>
			<td class="CELDAC7 underOrange width20">
			    <b><%=LitSubFamilia%></b>
			    <a CLASS="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>; display: inline-block;" id="OD3A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDeGrupo','F.nombre');" title="<%=LitOrdTarSubf%>" ><b><%=escribe3%></b></a>
			</td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvp%></b></td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvpIva%></b></td>
		<%CloseFila
	%></table>
	<iframe name="marcoArticulosAdd" id='frArticulosAdd' src='ArticulosDeGrupo.asp' class="width100 md-table-responsive"<!--width='<% response.Write(iif(si_tiene_modulo_credito,"810","810"))  %>'-->></iframe>
	<table style="width:100%;" border="0" cellpadding="0" cellspacing="0"><%
		DrawFila ""
			%><td class=CELDA7 style="width: 140px;">
				<div align="left" valign="center" id="Nregs" style="width: 140px; font-weight: bold;"></div>
			</td>
			<td class=CELDAC7>
			</td>
			<td class=CELDAR7 width="20">
			<div id="IcoIns" style="visibility: hidden;">
			<a href="javascript:if(GuardarArticulos('ALTA'));"><img src="<%=themeIlion %><%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitGuardarArt%>"></a>
			</div></td><%
		CloseFila
	%></table><%
end sub

'**********************************************************************************************************
sub SpanBajasArticulos()
	'Línea para establecer los parámetros de relleno de iframe
    if session("version") <> "5" then%>
	    <div style="border: 1px solid Black;"><%
    else%>
        <div><%
    end if
		'Drawfila ""
        DrawDiv "9", "", ""
            DrawLabel "CELDAB7", "", LitRefContiene
            DrawInput "CELDA7", "width:130px","brefcontiene","",""     
        CloseDiv
        DrawDiv "9", "", ""
			DrawLabel "CELDAB7", "", LitDesContiene
            DrawInput "CELDA7", "width:130px", "bdescontiene","",""
        CloseDiv
			'strselect ="select * from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='ARTICULO' order by descripcion "
			'rstAux.cursorlocation=3
			'rstAux.open strselect, session("dsn_cliente")
             strselect ="select * from tipos_entidades with(nolock) where codigo like ? + '%' and tipo=? order by descripcion "
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,5,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo",adVarChar,adParamInput,50,"ARTICULO")
                
        set rstAux= command2.Execute
        DrawDiv "9", "", ""
			DrawLabel "CELDAB7", "", LitTipoArt%>
		    <select style="width:130px;" class=CELDAL7 name="btipoarticulo">
			    <%if tipoarticulo="" then %>                                                            null_s(
			    	<option selected value=""> </option>
				<%else%>
				    <option selected value="<%=enc.EncodeForHtmlAttribute(tipoarticulo)%>"><%=enc.EncodeForHtmlAttribute(trimCodEmpresa(null_s(tipoarticulo)))%></option>
			    	<option value=""> </option>
				<%end if
			 	while not rstAux.eof%>
		   			<option value="<%=enc.EncodeForHtmlAttribute(rstAux("codigo"))%>"><%=enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion")))%></option>
					<%rstAux.movenext
				wend%>
		   	</select><%                                                         
        CloseDiv
        conn2.close
        set conn2    =  nothing
        set command2 =  nothing
        set rstAux  =  nothing	
		'CloseFila
        %><table></table><%

		dim ConfigDespleg (3,13)

		i=0
		ConfigDespleg(i,0)="categoria1"
		ConfigDespleg(i,1)="130"
		ConfigDespleg(i,2)="6"
		ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
		ConfigDespleg(i,4)=1
		ConfigDespleg(i,5)="CELDA7 colspan=2"
		ConfigDespleg(i,6)="MULTIPLE"
		ConfigDespleg(i,7)="codigo"
		ConfigDespleg(i,8)="nombre"
		ConfigDespleg(i,9)="<b>" & LitCategoria & "</b>"
		ConfigDespleg(i,10)=categoria
		ConfigDespleg(i,11)=""
		ConfigDespleg(i,12)=""

		i=1
		ConfigDespleg(i,0)="familia_padre1"
		ConfigDespleg(i,1)="130"
		ConfigDespleg(i,2)="6"
		ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
		ConfigDespleg(i,4)=1
		ConfigDespleg(i,5)="CELDA7"
		ConfigDespleg(i,6)="MULTIPLE"
		ConfigDespleg(i,7)="codigo"
		ConfigDespleg(i,8)="nombre"
		ConfigDespleg(i,9)="<b>" & LitFamilia & "</b>"
		ConfigDespleg(i,10)=familia_padre
		ConfigDespleg(i,11)=""
		ConfigDespleg(i,12)=""

		i=2
		ConfigDespleg(i,0)="familia1"
		ConfigDespleg(i,1)="130"
		ConfigDespleg(i,2)="6"
		ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
		ConfigDespleg(i,4)=1
		ConfigDespleg(i,5)="CELDA7"
		ConfigDespleg(i,6)="MULTIPLE"
		ConfigDespleg(i,7)="codigo"
		ConfigDespleg(i,8)="nombre"
		ConfigDespleg(i,9)="<b>" & LitSubFamilia & "</b>"
		ConfigDespleg(i,10)=familia
		ConfigDespleg(i,11)=""
		ConfigDespleg(i,12)=""

		DibujaDesplegablesCustom ConfigDespleg,session("dsn_cliente")

        %><table></table><%

        DrawDiv "col-xxs-12", "", ""
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", LitImporteRango
            CloseDiv
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", LitDesde
                DrawInput "CELDA7", "width:130px;","brangodesde","",""
            CloseDiv
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", LitHasta
                DrawInput "CELDA7", "width:130px;","brangohasta","",""
            CloseDiv
            DrawDiv "10", "", ""
                DrawLabel "CELDAB7", "", litPvpIva
                %><input class="" type="Checkbox" name="bpvpiva"><%
            CloseDiv 
        CloseDiv
        DrawDiv "9", "", ""
            DrawLabel "CELDAB7", "", LitProveedor
		 	'rstAux.cursorlocation=3
		 	'rstAux.open "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like '" & session("ncliente") & "%' order by razon_social",session("dsn_cliente")
            strselect= "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like ? + '%' order by razon_social"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,5,session("ncliente")&"")
            
            set rstAux= command2.Execute
			DrawSelectCeldaDet "","CELDA7 style='width:130px'","",0,"","bproveedor",rstAux,"","nproveedor","razon_social","",""
			'rstAux.close
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing
        CloseDiv

        DrawDiv "col-xxs-12", "text-align: right;", ""
            DrawLabel "CELDAB7", "", LitCargarArticulos%>
			<a class="ic-accept floatRight" href="javascript:if(Insertar('BAJA','1','first','A.referencia','asc'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>"></a>
		<%CloseDiv%>
	</div>
	<table class="width100 md-table-responsive">
        <%escribe1="&darr;"
        escribe2="&harr;"
        escribe3="&harr;"
        colorflecha="blue"

		Drawfila color_terra%>
			<td class="CELDAC7 underOrange width5"><input class="" type="Checkbox" name="checkb" onClick="seleccionar('marcoArticulosBorrar','ArticulosDeGrupo','checkb');"></td>
			<td class="CELDAC7 underOrange width15">
			    <b><%=LitReferencia%></b>
			    <a CLASS="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>; display: inline-block;" id="OD1B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDeGrupo','A.referencia');" title="<%=LitOrdTarRef & " " & LitOrdSentidoD%>"><b><%=escribe1%></b></a>
			</td>
			<td class="CELDAC7 underOrange width25">
			    <b><%=LitDescripcion%></b>
			    <a CLASS="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>; display: inline-block;" id="OD2B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDeGrupo','A.nombre');" title="<%=LitOrdTarDesc%>" ><b><%=escribe2%></b></a>
			</td>
			<td class="CELDAC7 underOrange width20">
			    <b><%=LitSubFamilia%></b>
			    <a CLASS="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>; display: inline-block;" id="OD3B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDeGrupo','F.nombre');" title="<%=LitOrdTarSubf%>" ><b><%=escribe3%></b></a>
			</td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvp%></b></td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvpIva%></b></td>
		<%CloseFila%>
	</table>
	<iframe name="marcoArticulosBorrar" id='frArticulosBorrar' src='ArticulosDeGrupo.asp' class="width100 md-table-responsive"<!--width='<% response.Write(iif(si_tiene_modulo_credito,"810","810"))  %>'-->></iframe>
	<table style="width:100%;" border="0" cellpadding="0" cellspacing="0"><%
		DrawFila ""
			%><td class=CELDA7 style="width: 140px;">
				<div align="left" id="NregsB" style="width: 140px; font-weight: bold;"></div>
			</td>
			<td class=CELDAC7>
			</td>
			<td class=CELDAR7 width="48">
				<!--<div id="IcoBorrModif" style="visibility: hidden;"><a href="javascript:if(GuardarArticulos('BAJA'));"><img src="../images/<%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitGuardarArt%>"></a>&nbsp;-->
				<a href="javascript:if(BorrarArticulos());"><img src="<%=themeIlion %><%=ImgEliminar%>" <%=ParamImgEliminar%> alt="<%=LitEliminarArt%>"></div></a>
			</td><%
		CloseFila
	%></table><%
end sub

'**************************************************************************************************
'                                   Código principal de la página
'**************************************************************************************************

if accesoPagina(session.sessionid,session("usuario"))=1 then
    
   %><form name="grupos_ofertas" method="post" action="grupos_ofertas.asp"><%

	si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	'JMMMM - 28/01/2010 --> Se añade modulo línea de crédito
	si_tiene_modulo_credito=ModuloContratado(session("ncliente"),ModLineaCredito)
	' JMMM - 03/11/2010 -> Franquicias
	si_tiene_modulo_franquicia = ModuloContratado(session("ncliente"),ModFranquiciasTiendas)
	'esFranquiciador=d_lookup("franquiciador", "configuracion", "nempresa='"&session("ncliente")&"'", session("dsn_cliente"))'
    
    esFranquiciadorSelect= "select franquiciador from configuracion with(nolock) where nempresa =?"
    esFranquiciadorReal=DlookupP1(esFranquiciadorSelect, session("ncliente")&"",adVarChar, 5, session("dsn_cliente"))

	set rst = server.CreateObject("ADODB.Recordset")
	set rstAux = server.CreateObject("ADODB.Recordset")
	set rstAux2 = server.CreateObject("ADODB.Recordset")
	set rstAux3 = server.CreateObject("ADODB.Recordset")

	mode=enc.EncodeForJavascript(request.querystring("mode"))
	mode2=enc.EncodeForJavascript(request.querystring("mode2"))

	codigoI=left(limpiaCadena(Request.Form("i_codigo")),5)
	descripcionI=limpiaCadena(request.form("i_descripcion"))
	observacionesI = nulear(limpiaCadena(request.form("i_observaciones")))

	codigoE=limpiaCadena(Request.Form("e_codigo"))
	CheckCadena codigoE
	descripcionE=limpiaCadena(request.form("e_descripcion"))
	observacionesE = nulear(limpiaCadena(request.form("e_observaciones")))

	condbase=enc.EncodeForJavascript(request.form("condbase"))
	bcondbase=enc.EncodeForJavascript(request.form("bcondbase"))

	if condbase&""=""then
		condbase="0"
	end if
	if bcondbase&""=""then
		bcondbase="0"
	end if
	
	WaitBoxOculto LitEsperePorFavor
				
	%><input type="hidden" name="condbase" value="<%=enc.EncodeForHtmlAttribute(condbase)%>">
	  <input type="hidden" name="bcondbase" value="<%=enc.EncodeForHtmlAttribute(bcondbase)%>"><%

	if mode="delete" then
		p_codigo=limpiaCadena(request("codigo"))
''ricardo 28-1-2008 solamente se concatenara el nempresa cuando venga del mode=add
        'mmg:evita que casque el CheckCadena y te expulse del sistema
		if mode2="add" then
			p_codigo=session("ncliente")& p_codigo
		end if
	else
		p_codigo=limpiaCadena(request.form("codigo"))
		if p_codigo="" then
			p_codigo=limpiaCadena(request.querystring("codigo"))
		end if
		if p_codigo="" then p_codigo=limpiaCadena(request("p_codigo"))
	end if
	CheckCadena p_codigo

	'insertamos si nos llegan los valores
	if codigoI>"" and descripcionI>"" then
		'rst.Open "select * from grupos_oferta where codigo='" & session("ncliente")&codigoI & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        strselect="select * from grupos_oferta  where codigo=?"

        set conn = Server.CreateObject("ADODB.Connection")
        set rst = Server.CreateObject("ADODB.Recordset")
	    set command =  Server.CreateObject("ADODB.Command")
	    conn.open session("dsn_cliente")
        'conn.cursorlocation=3

	    command.ActiveConnection =conn
	    command.CommandTimeout = 60
	    command.CommandText=strselect
	    command.CommandType = adCmdText 'CONSULTA
        command.Parameters.Append command.CreateParameter("@codigo", adVarChar, adParamInput, 10, session("ncliente")&codigoI)
        rst.CursorLocation = adUseClient
        rst.Open command, , adOpenKeyset, adLockOptimistic
		if rst.EOF then
			rst.AddNew
			rst("codigo")  = session("ncliente")&codigoI
			rst("nombre")   = descripcionI
            

   			'rst.Update
            rst.Update


		else %>
			<script>
				window.alert("<%=LitMsgCodigoExiste%>");
				history.back();
			</script>
		<%end if
		'rst.Close
        rst.close
        conn.close
        set rst = nothing
        set command = nothing
        set conn = nothing

	end if

	'actualizamos valores
	if codigoE>"" and descripcionE>"" and mode<>"delete" then
		'rst.Open "select * from grupos_oferta with(rowlock) where codigo='" & codigoE & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        strselect="select * from grupos_oferta with(rowlock) where codigo=?"

        set conn = Server.CreateObject("ADODB.Connection")
        set rst = Server.CreateObject("ADODB.Recordset")
	    set command =  Server.CreateObject("ADODB.Command")
	    conn.open session("dsn_cliente")
        'conn.cursorlocation=3

	    command.ActiveConnection =conn
	    command.CommandTimeout = 60
	    command.CommandText=strselect
	    command.CommandType = adCmdText 'CONSULTA
        command.Parameters.Append command.CreateParameter("@codigo", adVarChar, adParamInput, 10, codigoE&"")
        rst.CursorLocation = adUseClient
        rst.Open command, , adOpenKeyset, adLockOptimistic
		if not rst.EOF then
			'rst("codigo")  = codigoE
			rst("nombre")   = descripcionE

			rst.Update
		else %>
			<script>
				window.alert("<%=LitMsgCodigoNoExiste%>");
				history.back();
			</script>
		<%end if
		'rst.Close
        rst.close
        conn.close
        set rst = nothing
        set command = nothing
        set conn = nothing
	end if

	'eliminamos valores
	if mode="delete" and p_codigo>"" then
		'miramos a ver si esta puesta en algun documento
		no_borrar=0
		rst.cursorlocation=3
		rst.open "select tarifa from facturas_cli with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")
        'strselect= "select tarifa from facturas_cli with(nolock) where tarifa=?" 
        'set command2 = nothing
        'set conn2 = Server.CreateObject("ADODB.Connection")
        'set command2 =  Server.CreateObject("ADODB.Command")
        'conn2.open session("dsn_cliente")
        'conn2.cursorlocation=3
        'command2.ActiveConnection =conn2
        'command2.CommandTimeout = 60
        'command2.CommandText=strselect
        'command2.CommandType = adCmdText
        'command2.Parameters.Append command2.CreateParameter("@tarifa",adVarChar,adParamInput,10,p_codigo&"")
        'set rst= command2.Execute
		if not rst.eof then
			no_borrar=1
		end if
		rst.close
        'set conn2    =  nothing
        'set command2 =  nothing
        'set rst  =  nothing
		rst.cursorlocation=3
		rst.open "select tarifa from pedidos_cli with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")
        'strselect= "select tarifa from pedidos_cli with(nolock) where tarifa=?"
        'set command2 = nothing
        'set conn2 = Server.CreateObject("ADODB.Connection")
        'set command2 =  Server.CreateObject("ADODB.Command")
        'conn2.open session("dsn_cliente")
        'conn2.cursorlocation=3
        'command2.ActiveConnection =conn2
        'command2.CommandTimeout = 60
        'command2.CommandText=strselect
        'command2.CommandType = adCmdText
        'command2.Parameters.Append command2.CreateParameter("@tarifa",adVarChar,adParamInput,10,p_codigo&"")
        'set rst= command2.Execute
		if not rst.eof then
			no_borrar=1
		end if
		rst.close
        'set conn2    =  nothing
        'set command2 =  nothing
        'set rst  =  nothing	

		rst.cursorlocation=3
		rst.open "select tarifa from albaranes_cli with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")
        'strselect= "select tarifa from albaranes_cli with(nolock) where tarifa=?"
        'set command2 = nothing
        'set conn2 = Server.CreateObject("ADODB.Connection")
        'set command2 =  Server.CreateObject("ADODB.Command")
        'conn2.open session("dsn_cliente")
        'conn2.cursorlocation=3
        'command2.ActiveConnection =conn2
        'command2.CommandTimeout = 60
        'command2.CommandText=strselect
        'command2.CommandType = adCmdText
        'command2.Parameters.Append command2.CreateParameter("@tarifa",adVarChar,adParamInput,10,p_codigo&"")
        'set rst= command2.Execute
		if not rst.eof then
			no_borrar=1
		end if
		rst.close
        'set conn2    =  nothing
        'set command2 =  nothing
        'set rst  =  nothing	
		rst.cursorlocation=3
		rst.open "select tarifa from clientes with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")
        'strselect= "select tarifa from clientes with(nolock) where tarifa=?"
        'set command2 = nothing
        'set conn2 = Server.CreateObject("ADODB.Connection")
        'set command2 =  Server.CreateObject("ADODB.Command")
        'conn2.open session("dsn_cliente")
        'conn2.cursorlocation=3
        'command2.ActiveConnection =conn2
        'command2.CommandTimeout = 60
        'command2.CommandText=strselect
        'command2.CommandType = adCmdText
        'command2.Parameters.Append command2.CreateParameter("@tarifa",adVarChar,adParamInput,10,p_codigo&"")
        'set rst= command2.Execute
		if not rst.eof then
			no_borrar=1
		end if
		rst.close
        'set conn2    =  nothing
        'set command2 =  nothing
        'set rst  =  nothing	

		if no_borrar=0 then
			%>
                <SCRIPT LANGUAGE="JavaScript">
                      document.getElementById("waitBoxOculto").style.visibility = "visible";
			    </script>
            <%
			response.flush
			rst.open "delete from articulos_grupos_oferta with(rowlock) where codigo='" & p_codigo & "' ",session("dsn_cliente"),adUseClient, adLockReadOnly
            'strselect= "delete from articulos_grupos_oferta with(rowlock) where codigo=?" 
            'set command2 = nothing
            'set conn2 = Server.CreateObject("ADODB.Connection")
            'set command2 =  Server.CreateObject("ADODB.Command")
            'conn2.open session("dsn_ncliente")
            'conn2.cursorlocation=3
            'command2.ActiveConnection =conn2
            'command2.CommandTimeout = 60
            'command2.CommandText=strselect
            'command2.CommandType = adCmdText
            'command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,p_codigo&"")
            'set rst= command2.Execute
            'rst.close
            'set conn2    =  nothing
            'set command2 =  nothing
            'set rst  =  nothing
			 'AHORA SE BORRA LA TARIFA
 	        rst.Open "delete from grupos_oferta with(rowlock) where codigo='" & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
            'strselect= "delete from grupos_oferta with(rowlock) where codigo=?"
            'set command2 = nothing
            'set conn2 = Server.CreateObject("ADODB.Connection")
            'set command2 =  Server.CreateObject("ADODB.Command")
            'conn2.open session("dsn_ncliente")
            'conn2.cursorlocation=3
            'command2.ActiveConnection =conn2
            'command2.CommandTimeout = 60
            'command2.CommandText=strselect
            'command2.CommandType = adCmdText
            'command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,p_codigo&"")
            'set rst= command2.Execute
            'rst.close
            'set conn2    =  nothing
            'set command2 =  nothing
            'set rst  =  nothing
			 %><SCRIPT LANGUAGE="JavaScript">
			       document.getElementById("waitBoxOculto").style.visibility = "hidden";
			</script><%
		else
			%><SCRIPT LANGUAGE="JavaScript">
				window.alert("<%=LitTarifaNoBorrarAsocDoc%>");
			</script><%
		end if
	end if

      p_criterio=limpiaCadena(request("criterio"))
      p_campo=limpiaCadena(request("campo"))
      p_texto=limpiaCadena(request("texto"))
      p_npagina=limpiaCadena(request("npagina"))

      if p_texto>"" then
	  	 if p_campo="codigo" then p_campo="substring(codigo,6,10)"
         c_where=" where " & p_campo & " "
      else
         c_where=""
      end if

      if c_where>"" then
         select case p_criterio
            case "contiene"
               c_where=c_where+ "like '%" & p_texto & "%'"
            case "termina"
               c_where=c_where+ "like '%" & p_texto & "'"
            case "empieza"
               c_where=c_where+ "like '" & p_texto & "%'"
            case "igual"
              c_where=c_where + "='" & p_texto & "'"
         end select
		 c_where=c_where & " and codigo like '" & session("ncliente") & "%' "
	  else
	  	 c_where=" where codigo like '" & session("ncliente") & "%' "
      end if

   PintarCabecera "grupos_ofertas.asp"
 Alarma "grupos_ofertas.asp" %>
   <%
    c_select="select * from grupos_oferta with(nolock)"

        if c_where>"" then
           c_select=c_select & c_where
        end if

        if p_npagina="" then
           p_npagina=1
        end if

        select case request("pagina")
           case "siguiente"
              p_npagina=p_npagina+1
           case "anterior"
              p_npagina=p_npagina-1
        end select%>
  <input type="hidden" name="h_npagina" value="<%=enc.EncodeForHtmlAttribute(cstr(p_npagina))%>">
	<%
        set rst = Server.CreateObject("ADODB.Recordset")
        rst.Open c_select,session("dsn_cliente"),adUseClient, adLockReadOnly

        if not rst.EOF then
           rst.PageSize=NumReg
           rst.AbsolutePage=p_npagina
        end if

  if mode<>"edit" and rst.RecordCount>NumReg then                                 
      if clng(p_npagina) >1 then %>
		 <a class=CABECERA href="grupos_ofertas.asp?pagina=anterior&npagina=<%=enc.EncodeForJavascript(cstr(p_npagina))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
  		 <IMG SRC="<%=themeIlion %><%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></a>
  	<%end if

    texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	<font class=CELDA> <%=texto%> </font> <%

     if clng(p_npagina)<rst.PageCount then %>
		<a class=CABECERA href="grupos_ofertas.asp?pagina=siguiente&npagina=<%=enc.EncodeForJavascript(cstr(p_npagina))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
  		<IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></a>
  	<%end if

	%><font class=CELDA>&nbsp;&nbsp; <%=LitPagIrA%> <input class=CELDA type="text" name="SaltoPagina1" size="2">&nbsp;&nbsp;<a class=CELDAREF href="javascript:IrAPagina(1,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=enc.EncodeForJavascript(rst.PageCount)%>,'npagina');"><%=LitIr%></a></font><%
  end if

    ' Botón para actualizar todos los grupos en todas las franquicias
    if mode<>"edit" and si_tiene_modulo_franquicia and esFranquiciador=true then%>
       <br />
        <table style="border : 1px solid black; border-collapse : collapse;" cellpadding="3">
            <tr>
                <td style="border : 1px solid black;" align="center" class="CELDABOT" onmouseover="this.className='TDACTIVO8'" onmouseout="this.className='CELDABOT'" bgcolor="<%=color_blau%>"><a class="CELDAREFB7" href="javascript:ActualizarGrupos();" OnMouseOver="self.status='<%=LITACTUALIZAR_GRUPOS_FRANQ%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LITACTUALIZAR_GRUPOS_FRANQ%>&nbsp;&nbsp;&nbsp;</a></td>
            </tr>
        </table>
       
       <%
    end if
    
    if mode<>"edit" then %>
   <br />
   <table BORDER="0" CELLSPACING="1" CELLPADDING="1">
   <%drawfila color_terra
     Drawcelda2 "'CELDA underOrange NO_BORDER_H' style='width:502px'", "left", true, LitNBregistro %>
   </table>
		<table BORDER="0" CELLSPACING="1" CELLPADDING="1">
		<%Drawfila color_terra
        	Drawcelda2 "'CELDA underOrange NO_BORDER_H' style='width:100px'", "left", true, LitCodigo
        	Drawcelda2 "'CELDA underOrange NO_BORDER_H' style='width:400px'", "left", true, LitDescripcion
        	
		CloseFila
		Drawfila color_blau
            DrawInputCeldaSpan "ENCABEZADOL maxlength='5' style='width:100px'","","",7,0,"","i_codigo","",""
            DrawInputCeldaSpan "ENCABEZADOL maxlength='50' style='width:400px''","","",57,0,"","i_descripcion","",""
		CloseFila%>
      </table>
      
   <%end if%>
    <br />
    <table class="searchResult" style="width:auto;" BORDER="0" CELLSPACING="1" CELLPADDING="1">
    <%if mode<>"edit" then
        Drawfila color_terra
            DrawCelda2 "'CELDA underOrange' valign='top' style='width:100px'","left",true,LitCodigo
            DrawCelda2 "'CELDA underOrange' valign='top' style='width:400px'","left",true,LitDescripcion  
        CloseFila
    end if
      Drawfila color_fondo
        par=false
        i=1

        while not rst.EOF and i<=NumReg
           if mode="edit" and p_codigo=rst("codigo") then

           elseif mode<>"edit" then
				h_ref="javascript:Editar('" & enc.EncodeForJavascript(null_s(rst("codigo"))) & "'," & enc.EncodeForJavascript(null_s(p_npagina)) & ",'" & enc.EncodeForJavascript(null_s(p_campo)) & "','" & enc.EncodeForJavascript(null_s(p_criterio)) & "','" & enc.EncodeForJavascript(null_s(p_texto)) & "');"
				if ucase(rst("codigo"))<>session("ncliente") & "BASE" then
					if par then
						Drawfila color_terra
						par=false
					else
            			Drawfila color_blau
	            	  	par=true
					end if
                    %><td class="CELDAREF">              
                    <%DrawHref "CELDAREF valign='top' valign='top' style='width:100px'","left",enc.EncodeForHtmlAttribute(trimCodEmpresa(rst("codigo"))),h_ref%></td><%
					DrawCelda2 "CELDA valign='top' style='width:400px'", "left", false, rst("nombre")
					
					CloseFila
				end if
           end if

           i = i + 1
           rst.MoveNext
        wend%>
    </table>

    <%if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
		 <a class=CABECERA href="grupos_ofertas.asp?pagina=anterior&npagina=<%=enc.EncodeForJavascript(cstr(p_npagina))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
  		 <IMG SRC="<%=themeIlion %><%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></a>
  	<%end if

    texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	<font class=CELDA> <%=texto%> </font> <%

     if clng(p_npagina)<rst.PageCount then %>
		<a class=CABECERA href="grupos_ofertas.asp?pagina=siguiente&npagina=<<%=enc.EncodeForJavascript(cstr(p_npagina))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
  		<IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></a>
  	<%end if%>
	<font class=CELDA>&nbsp;&nbsp; <%=LitPagIrA%> <input class=CELDA type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;<a class=CELDAREF href="javascript:IrAPagina(2,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=enc.EncodeForJavascript(rst.PageCount)%>,'npagina');"><%=LitIr%></a></font><%
	rst.close
  end if   

   '***************************************************************************
   'Zona de código para la gestión de artículos de la tarifa
   '***************************************************************************

   if mode="edit" then
   		%><br><%BarraNavegacion%>        

        <div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['CABECERA', 'ARTICULOS']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['CABECERA', 'ARTICULOS']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
        </div>

        <div class="Section" id="S_CABECERA">
            <a href="#" rel="toggle[CABECERA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader"><%=LitCabecera%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
            <div class="SectionPanel" id="CABECERA" style="display:none;">
                <input type="Hidden" name="htarifa" value="<%=enc.EncodeForHtmlAttribute(p_codigo)%>">
                <table width=750 BORDER="0" CELLSPACING="1" CELLPADDING="1">
			      <%
                    dim color_terrap
                    color_terrap = color_terra

                    if session("version")&"" = "5" then  
                        color_terrap = "#fdfdfd"
                    end if

                    Drawfila color_fondo
                    Drawcelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitCodigo
				    Drawcelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitDescripcion

				    par=false                                                                          
				    i=1
				    rst.movefirst                                                                              
				    while not rst.EOF and i<=NumReg
				       if mode="edit" and p_codigo=enc.EncodeForHtmlAttribute(null_s(rst("codigo"))) then
					      Drawfila color_terrap
						      DrawCelda2 "CABECERA", "center", true, enc.EncodeForHtmlAttribute(null_s(trimCodEmpresa(rst("codigo"))))
						      %><input type="Hidden" name="e_codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("codigo")))%>"><%
                              DrawInputCeldaSpan "CELDA maxlength='50'","","",58,0,"","e_descripcion",enc.EncodeForHtmlAttribute(null_s(rst("nombre"))),""
					      CloseFila
				       end if
				       i = i + 1
				       rst.MoveNext
				    wend
				    'rst.Close %>
		       </table>
               <br />
            </div>
        </div>

        <div class="Section" id="S_ARTICULOS">
            <a href="#" rel="toggle[ARTICULOS]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader"><%=LITNOMBREGRUPOOFERTA%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
		    <div class="SectionPanel" id="ARTICULOS">
                <div id="tabs" style="display:none">
                    <ul>
                        <li><a href="#tabs1"><%=LitAnadirArticulos%></a></li>
                        <li><a href="#tabs2"><%=LitBorrModifArticulos%></a></li>

                    </ul>
                    <div id="tabs1">
                        <%SpanAltasArticulos p_codigo%>
                    </div>
                    <div id="tabs2">
                        <%SpanBajasArticulos%>
                    </div>
                </div>
            </div>
        </div>	              	
	<%end if%>
   </form>
<%end if%>
</BODY>
</HTML>