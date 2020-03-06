<%@ Language=VBScript %>
 <% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>

<% 
    folder = session("folder")
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
<% if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="categorias.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../js/generic.js.inc"-->

<!--#include file="../styles/formularios.css.inc" -->  

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
function Insertar2() {
    if (document.categorias.i_codigo.value==""){
        window.alert("<%=LitMsgCodigoNoNulo%>");
        return;
    }

    if (comp_car_ext(document.categorias.i_codigo.value,1)==1){
		window.alert("<%=LitMsgTipoADesCarNoVal%>");
		return;
   }

   if (document.categorias.i_descripcion.value==""){
        window.alert("<%=LitMsgDescripcionNoNulo%>");
        return;
   }

    if (comp_car_ext(document.categorias.i_descripcion.value,0)==1){
		window.alert("<%=LitMsgTipoADesCarNoVal%>");
		return;
   }

	//Recargar el submarco de detalles
	fr_Tabla.document.categorias_det.action="categorias_det.asp?mode=save&i_codigo=" + document.categorias.i_codigo.value + "&i_descripcion=" + document.categorias.i_descripcion.value;
	fr_Tabla.document.categorias_det.submit();
	//Limpiar los campos del formulario
	document.categorias.i_descripcion.value="";
	document.categorias.i_codigo.value="";
	//Colocar el foco en el campo de codigo
	document.categorias.i_codigo.focus();
}

function Mas(sentido,lote, texto) {
	document.getElementById("barras").style.display="none";
	
	fr_Tabla.document.categorias_det.action="categorias_det.asp?mode=ver&sentido=" + sentido + "&lote=" + lote + "&texto=" + texto;
	fr_Tabla.document.categorias_det.submit();
}

function Insertar() {
    Insertar2();
}

function Resize()
{
    var alto = jQuery(window).height();
    var diference = 175;
    var dir_default = 200;

    if (alto > dir_default)
    {
        if (alto - diference > dir_default) jQuery("#frtabla").attr("height", alto - diference);
        else jQuery("#frtabla").attr("height", dir_default);
    }
    else jQuery("#frtabla").attr("height", dir_default);
}

jQuery(window).resize(function () { Resize(); });
</script>

<body class="BODY_ASP">
<%
'***********************************************************************************************************
' CODIGO PRINCIPAL DE LA PAGINA  ***************************************************************************
'***********************************************************************************************************'

	set rstselect = Server.CreateObject("ADODB.Recordset")%>
<form name="categorias" method="post" action="categorias.asp">
    <%PintarCabecera "familias.asp"%>
    
	<br/><table bgcolor="<%=color_fondo%>">
   		<tr>
			<td class=CABECERA width="33%" style="text-align: center">
				<b><%=ucase(LitCategorias)%></b>
			</td>
			<td width="25%" style="text-align: center" class="CABECERA" onmouseover="this.className='tdactivo10'" onmouseout="this.className='CABECERA'" onclick="document.location='familias_padre.asp';parent.botones.document.location='familias_padre_bt.asp';" bgcolor="<%=color_blau%>">
				<b><%=ucase(LitFamilias)%></b>
			</td>
			<td width="25%" style="text-align: center" class="CABECERA" onmouseover="this.className='tdactivo10'" onmouseout="this.className='CABECERA'" onclick="document.location='familias.asp?mode=add';parent.botones.document.location='familias_bt.asp?mode=add';" bgcolor="<%=color_blau%>">
				<b><%=ucase(LitSubFamilias)%></b>
			</td>
   	    </tr>
    </table><%
	
	Alarma "familias.asp"

	mode=enc.EncodeForHtmlAttribute(null_s(request("mode")))%>
    <input type="hidden" name="mode_accesos_tienda" value="<%=mode%>" />
    <%p_codigo=enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("p_codigo"))))%>

    <table class="underOrange bCollapse" width="460">
        <%
        Drawfila color_terra
        %>
            <td class="ENCABEZADOL underOrange" width="55" ><b><%=LitCodigo%></b></td>
		    <td class="ENCABEZADOL underOrange" ><b><%=LitDescripcion%></b></td>
            <td class="ENCABEZADOL underOrange" width="40">&nbsp</td>
        <%
        CloseFila
        %>
        <tr>
            <td class="CELDAL7 underOrange" width="55" >
                <input class="CELDAL7 width100" name="i_codigo">
            </td>
			<td class="CELDAL7 underOrange" >					
                <input class="CELDAL7 width100" name="i_descripcion">
			</td>
            <td class="CELDAR7 underOrange width5" width="40">
			    <a href="javascript:Insertar();" class="ic-accept NoMTop" ><img src="/lib/estilos/<%=folder%>/<%=ImgAplicar%>" <%=ParamImgNuevo%> alt="<%=LitNuevo1%>"></a>
			</td>
        </tr>
    </table>
	<%
        
			'Drawcelda2 "CELDA", "left", false, LitCodigo
        	'Drawcelda2 "CELDA", "left", false, LitDescripcion
            ''''EligeCelda "input","add","left","","",0,LitCodigo,"i_codigo",35,i_codigo
            'EligeCelda "input","add","left","","",0,LitDescripcion,"i_descripcion",35,i_descripcion
            ''''DrawInputCeldaImg "", "", "", 35, 0, LitDescripcion, "i_descripcion", i_descripcion, "Insertar2();", LitNuevo1,ImgNuevo
            
       
			'DrawInputCelda "ENCABEZADOL maxlength='5'","","",5,0,"","i_codigo",""
        	'DrawInputCelda "ENCABEZADOL maxlength='50'","","",50,0,"","i_descripcion",""
    %>
		    
		   	<!--<a href="javascript:Insertar2();" ><img src="../images/<%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo1%>"></a>-->
	
	
	<script language="javascript" type="text/javascript">
		document.categorias.i_codigo.focus();
	</script>
    <%DrawDiv "3", "", ""%>
    <iframe id="frtabla" name="fr_Tabla" src='categorias_det.asp?mode=browse' width='450' height='200' frameborder="yes" noresize="noresize"></iframe>
    <%CloseDiv %>
   	<table width="750"><%
		DrawFila ""
			%><td class='CELDA7' width="250">
				<span id="barras" style="display:none">
				</span>
			</td><%
		CloseFila%>
   	</table>
    <script language="javascript" type="text/javascript">Resize();</script>
</form>
<%end if
set rstselect = Nothing%>
</body>
</html>