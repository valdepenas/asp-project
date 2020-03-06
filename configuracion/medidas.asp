<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<TITLE><%=LitTituloUM%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
function Insertar()
{
    if (document.medidas.i_codigo.value==""){
        window.alert("<%=LitMsgAbreviaturaNoNulo%>");
        return;
    }

    if (document.medidas.i_descripcion.value==""){
        window.alert("<%=LitMsgDescripcionNoNulo%>");
        return;
    }

    if (comp_car_ext(document.medidas.i_descripcion.value,0)==1){
		window.alert("<%=LitMsgTipoADesCarNoVal%>");
		return;
   }

	//Recargar el submarco de detalles
	fr_Tabla.document.medidas_det.h_i_codigo.value=document.medidas.i_codigo.value;
	fr_Tabla.document.medidas_det.h_i_descripcion.value=document.medidas.i_descripcion.value;
	fr_Tabla.document.medidas_det.action="medidas_det.asp?mode=save";
	fr_Tabla.document.medidas_det.submit();
	//Limpiar los campos del formulario
	document.medidas.i_descripcion.value="";
	document.medidas.i_codigo.value="";
	//Colocar el foco en el campo de descripcion
	document.medidas.i_codigo.focus();
}

function Mas(sentido,lote, texto) {
	document.getElementById("barras").style.display="none";
	fr_Tabla.document.medidas_det.action="medidas_det.asp?mode=ver&sentido=" + sentido + "&lote=" + lote + "&texto=" + texto;
	fr_Tabla.document.medidas_det.submit();
}

if(window.document.addEventListener)
{
    window.document.addEventListener("keydown", callkeydownhandler, false);
}
else
{
    window.document.attachEvent("onkeydown", callkeydownhandler);
}

var ev = null;

function callkeydownhandler(evnt)
{
    ev = (evnt) ? evnt : event;
    keyPressed();
}

//Comprueba si la tecla pulsada es CTRL+S. Si es así guarda el registro.
function keyPressed()
{
    var keycode = ev.keyCode;
	if (keycode==19) //CTRL+S
		Insertar();
}
</script>
<body bgcolor=<%=color_blau%> onkeypress="keyPressed();">
<%
'***********************************************************************************************************'
' CODIGO PRINCIPAL DE LA PAGINA  ***************************************************************************'
'***********************************************************************************************************'
if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<form name="medidas" method="post" action="medidas.asp">
    <%'Leer parámetros de la página'
	mode=request("mode")

   PintarCabecera "medidas.asp"%>
	<% Alarma "medidas.asp"

	 if mode<>"edit" then %>
   		<hr>
		<table class="width100 underOrange md-table-responsive" BORDER="0" CELLSPACING="1" CELLPADDING="1">
            <tr class="underOrange"><%
				DrawceldaDet "'ENCABEZADOL underOrange width5'", "", "left", true,"<b>" & LitAbreviatura & "</b>"
        		DrawceldaDet "'ENCABEZADOL underOrange width20'", "", "left", true,"<b>" & LitDescripcion & "</b>"
                DrawceldaDet "'ENCABEZADOL underOrange width50'", "", "left", true, ""%>
            </tr>
            <tr>
                <td class="CELDAL7 underOrange width5">
                    <input type="text" class="width100" name="i_codigo" maxlength="5" />
                </td>
                <td class="CELDAL7 underOrange width20">
                    <input type="text" class="width100" name="i_descripcion" maxlength="50" />
                </td>
				<td class="CELDAL7 underOrange width50">
		   			<a class="ic-accept noMTop" href="javascript:Insertar();" ><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo1%>"></a>
				</td>
            </tr>
		</table>
		<script language="javascript">
			document.medidas.i_codigo.focus();
		</script>
   <%end if

   if mode<>"edit" then%>
      <iframe class="width100 iframe-data md-table-responsive" id="frtabla" name="fr_Tabla" src='medidas_det.asp?mode=browse' width='100%' height='250' frameborder="no" noresize="noresize"></iframe><%
   end if%>
   <table width="750">
        <%DrawFila ""%>
			<td class=CELDA7 width="250">
				<SPAN ID="barras" STYLE="display:none">
				</SPAN>
			</td>
		<%CloseFila%>
   </table>
</form>
<%end if%>
</BODY>
</HTML>