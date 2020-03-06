<%@ Language=VBScript %>
<%
'**RGU 31/10/2006: Añadir campo tipo apunte si se tiene el módulo ebesa
%>
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
<!--#include file="../modulos.inc"-->
<!--#include file="../styles/formularios.css.inc" -->
<TITLE><%=LitTituloMP%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
function Insertar()
{
    if (isNaN(document.tipo_pago.i_codigo.value))
    {
        alert("<%=LitMsgCodigoNoNulo1%>");
        return;
    }

    if (document.tipo_pago.i_descripcion.value=="")
    {
        alert("<%=LitMsgDescripcionNoNulo%>");
        return;
    }

    if (isNaN(document.tipo_pago.i_gasto.value.replace(",","."))||document.tipo_pago.i_gasto.value=="")
    {
        alert("<%=LitMsgGastoNoValido%>");
        return;
    }

    if (comp_car_ext(document.tipo_pago.i_descripcion.value,0)==1)
    {
	    window.alert("<%=LitMsgTipoADesCarNoVal%>");
	    return;
    }
    
	//Recargar el submarco de detalles
     var uri = "tipo_pago_det.asp?mode=save&cod_tipo=" + encodeURIComponent(document.tipo_pago.i_codigo.value) +
        "&desc_tipo=" + encodeURIComponent(document.tipo_pago.i_descripcion.value) +
        "&gasto_tipo=" + encodeURIComponent(document.tipo_pago.i_gasto.value.replace(".", ",")) +
        "&tapunte=" + encodeURIComponent(document.tipo_pago.i_tapunte.value) +
        "&web=" + encodeURIComponent(document.tipo_pago.i_web.checked);
    fr_Tabla.document.tipo_pago_det.action = uri;
	fr_Tabla.document.tipo_pago_det.submit();

	//Limpiar los campos del formulario
	document.tipo_pago.i_codigo.value="";
	document.tipo_pago.i_descripcion.value="";
	document.tipo_pago.i_gasto.value="0";
	document.tipo_pago.i_tapunte.value=""

	//Colocar el foco en el campo de cantidad.
	document.tipo_pago.i_codigo.focus();
}

function Mas(sentido,lote,campo,criterio, texto)
{
    document.getElementById("barras").style.display = "none";
    var uri = "tipo_pago_det.asp?mode=ver&sentido=" + encodeURIComponent(sentido) +
        "&lote=" + encodeURIComponent(lote) +
        "&campo=" + encodeURIComponent(campo) +
        "&criterio=" + encodeURIComponent(criterio) +
        "&texto=" + encodeURIComponent(texto);
	fr_Tabla.document.tipo_pago_det.action = uri;
	fr_Tabla.document.tipo_pago_det.submit();
}

//Comprueba si la tecla pulsada es CTRL+S. Si es así guarda el registro.
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

function keyPressed()
{
    var keycode = ev.keyCode;
	if (keycode==19) //CTRL+S
		Insertar();
}
</script>
<body bgcolor=<%=color_blau%>  onkeypress="keyPressed();">
<%
'***********************************************************************************************************
' CODIGO PRINCIPAL DE LA PAGINA  ***************************************************************************
'***********************************************************************************************************
if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<form name="tipo_pago" method="post" action="tipo_pago.asp">
<%
    ' ¿Tiene el módulo de tienda contratado? Recogemos el dato correspondiente.
    contratado_tienda = ModuloContratado(session("ncliente"), ModTiendaWeb)
    si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
    'contratado_tienda=false
    'si_tiene_modulo_ebesa=false

	PintarCabecera "tipo_pago.asp"
	Alarma "tipo_pago.asp"

	mode=request("mode")

	set rstAux = server.CreateObject("ADODB.Recordset")

	if mode<>"edit" then%>
    <hr>
        <table class="width100 underOrange md-table-responsive" BORDER="0" CELLSPACING="1" CELLPADDING="1">
        <tr class="underOrange">
           <%DrawCeldaDet "'ENCABEZADOL underOrange width5'", "", "", 0,"<b>" & LitCodigo & "</b>"
             DrawCeldaDet "'ENCABEZADOL underOrange width20'", "", "", 0,"<b>" & LitDescripcion & "</b>"
             DrawCeldaDet "'ENCABEZADOL underOrange width10'", "", "", 0,"<b>" & LitGasto & "</b>"
             if si_tiene_modulo_ebesa then
                DrawCeldaDet "'ENCABEZADOL underOrange width10", "left", true,"<b>" & LitTapunteCC & "</b>"
             end if
             if contratado_tienda then
                DrawCeldaDet "'ENCABEZADOL underOrange width5'", "", "", 0,"<b>" & LitWeb & "</b>"
		     end if
             DrawCeldaDet "'ENCABEZADOL underOrange width5'", "", "", 0, ""%>
        </tr>
        <tr>
            <td class="CELDAL7 underOrange width5">
		        <%DrawInput "width50","","i_codigo","","size=5"%>
            </td>                
            <td class="CELDAL7 underOrange width20">
                <%DrawInput "width70","","i_descripcion","","size=5"%>
            </td>
            <td class="CELDAL7 underOrange width10">
                <%DrawInput "width50","","i_gasto","","size=5"%>
            </td><%
            if si_tiene_modulo_ebesa then
		        rstAux.open "select codigo,descripcion, cuenta from tipo_apuntes with(nolock) where codigo like '"&session("ncliente")&"%' order by descripcion", session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		        %><td class="CELDAL7 underOrange width10">
                    <select CLASS=ENCABEZADOL name='i_tapunte'>
		        	<option name='selnull' value=''></option>
		        <%if not rstAux.eof then
		        	while not rstAux.eof%>
		        		<option value='<%=enc.EncodeForHtmlAttribute(rstAux("codigo") & "")%>'><%=enc.EncodeForHtmlAttribute(rstAux("Descripcion")&" - "&rstAux("cuenta"))%></option>
		        		<%rstAux.movenext
		        	wend
		        	rstAux.close
		        end if%>
		        </select>
                </td><%
            else
               %><input type="hidden" name="i_tapunte" value=""><%
            end if
            if contratado_tienda then %>
                <td class="CELDAL7 underOrange width5"> 
                    <%DrawCheck "","","i_web",""%>
                </td><%
		    else%>
			<input type="hidden" name="i_web" value="0">
            <%end if%>
            <td class="CELDAL7 underOrange width5">
		        <a href="javascript:Insertar();" class="ic-accept noMTop"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgNuevo%> alt="<%=LitNuevo1%>"></a>
            </td>
        </tr>
        </table>
        <script language="javascript">
	        document.tipo_pago.i_codigo.focus();
	    </script>
   <%end if
   if mode<>"edit" then
   		if contratado_tienda <> 0 then
			width = "750"
		else
			width = "670"
		end if
		if si_tiene_modulo_ebesa then
			width=cint(width)+250
		end if%>
      <iframe id="frtabla" name="fr_Tabla" src='tipo_pago_det.asp?mode=browse' class="width100 iframe-data md-table-responsive" height='340' frameborder="no" noresize="noresize"></iframe><%
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
<%set rstAux=Nothing
end if%>
</BODY>
</HTML>