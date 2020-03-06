<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->

<TITLE><%=LitTituloTEnt%></TITLE>
    <% 
        dim  enc
        set enc = Server.CreateObject("Owasp_Esapi.Encoder")
    %>  

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="javascript" src="../jfunciones.js"></script>
<script Language="javascript">
function Guardar(param)
{
    ok=1;
    switch(param)
    {
        case 1:
            if (parent.pantalla.document.tipo_entidad.e_codigo.value=="")
            {
                alert ("<%=LitMsgCodigoNoNulo%>");
                ok=0;
            }
            if (parent.pantalla.document.tipo_entidad.e_descripcion.value=="")
            {
                alert ("<%=LitMsgDescripcionNoNulo%>");
                ok=0;
            }
            if (comp_car_ext(parent.pantalla.document.tipo_entidad.e_descripcion.value,0)==1)
            {
                alert("<%=LitMsgTipoEnDesCarNoVal%>");
                ok=0;
                break;
            }
	        break;

        case 2:
            if (parent.pantalla.document.tipo_entidad.i_codigo.value=="")
            {
                alert ("<%=LitMsgCodigoNoNulo%>");
                ok=0;
            }
            if (comp_car_ext(parent.pantalla.document.tipo_entidad.i_codigo.value,1)==1)
            {
                alert("<%=LitMsgTipoEnDesCarNoVal%>");
                ok=0;
                break;
		    }

            if (parent.pantalla.document.tipo_entidad.i_descripcion.value=="")
            {
                alert ("<%=LitMsgDescripcionNoNulo%>");
                ok=0;
            }
            
	        if (comp_car_ext(parent.pantalla.document.tipo_entidad.i_descripcion.value,0)==1)
	        {
		        alert("<%=LitMsgTipoEnDesCarNoVal%>");
		        ok=0;
		        break;
		    }
            break;
    }
    if (ok==1)
    {
        parent.pantalla.document.tipo_entidad.submit();
        if (param==1) document.location="tipo_entidad_bt.asp";
    }
}

function Buscar()
{
	parent.pantalla.document.location="tipo_entidad.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next";
	document.location="tipo_entidad_bt.asp";
}

function Eliminar(param)
{
    if (window.confirm("<%=LitMsgEliminarEntidadConfirm%>")==true)
    {
        switch (param)
        {
            case 1:
	            if (parent.pantalla.document.tipo_entidad.e_codigo.value=="") alert ("<%=LitMsgCodigoNoNulo%>");
                else parent.pantalla.document.location="tipo_entidad.asp?mode=delete&codigo=" + parent.pantalla.document.tipo_entidad.e_codigo.value;
    		    break;

            case 2:
	            if (parent.pantalla.document.tipo_entidad.i_codigo.value=="") alert ("<%=LitMsgCodigoNoNulo%>");
                else parent.pantalla.document.location="tipo_entidad.asp?mode=delete&codigo=" + parent.pantalla.document.tipo_entidad.i_codigo.value;
			    break;
        }
        document.location="tipo_entidad_bt.asp";
    }
}

function Cancelar()
{
	parent.pantalla.document.location="tipo_entidad.asp";
	document.location="tipo_entidad_bt.asp";
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
    comprobar_enter();
}

//****************************************************************************************************************
function comprobar_enter()
{
    var keycode = ev.keyCode;
	//si se ha pulsado la tecla enter
	if (keycode==13)
	{
		document.opciones.criterio.focus();
		Buscar();
	}
}
</script>
<body class="body_master_ASP">
<form name="opciones" method="post">
<%if request("mode")="edit" then
    param=1
else
    param=2
end if%>
    <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_left_ASP" >
	        <table id="BUTTONS_CENTER_ASP">
		        <tr>
    		        <td CLASS="CELDABOT" onclick="javascript:Guardar(<%=enc.EncodeForJavascript(param)%>);">
				        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
			        </td>
		            <td CLASS="CELDABOT" onclick="javascript:Eliminar(<%=enc.EncodeForJavascript(param)%>);">
			            <%PintarBotonBTRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,""%>
		            </td>
		            <td CLASS="CELDABOT" onclick="javascript:Cancelar('noasist');">
			            <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
		            </td>
                </tr>
            </table>
        </div>
        <div id="FILTERS_MASTER_ASP">
			<!--<td class="CELDABOT"><%=LitBuscar & ": "%>-->
				<select class="IN_S" name="campos">
					<option value="codigo"><%=LitCodigo%></option>
					<option selected value="descripcion"><%=LitDescripcion%></option>
					<option value="tipo"><%=LitTipoDe%></option>
				</select>
			<!--</td>
			<td class="CELDABOT">-->
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<!--<OPTION value="empieza"><%=LitComienza%></OPTION>-->
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
				</select>
			<!--</td>
			<td class="CELDABOT">-->
				<input class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" onkeypress="javascript:comprobar_enter();">
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>"></a>
			<!--</td>-->
		</div>
	</div>
<%ImprimirPie_bt%>
</form>
</BODY>
</HTML>