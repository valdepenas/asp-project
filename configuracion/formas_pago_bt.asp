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

<TITLE><%=LitTituloFP%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
function Guardar(param,viene) {
   ok=1;
   switch(param){
        case 1:
            if (parent.pantalla.document.formas_pago.e_codigo.value=="") {
                window.alert ("<%=LitMsgCodigoNoNulo%>");
                ok=0;
            }
            if (parent.pantalla.document.formas_pago.e_descripcion.value=="")  {
                window.alert ("<%=LitMsgDescripcionNoNulo%>");
                ok=0;
            }
		    if (comp_car_ext(parent.pantalla.document.formas_pago.e_descripcion.value,0)==1){
		        window.alert("<%=LitMsgFormPDesCarNoVal%>");
		        ok=0;
		        break;
		    }
		    break;

        case 2:
            if (parent.pantalla.document.formas_pago.i_codigo.value=="") {
                window.alert ("<%=LitMsgCodigoNoNulo%>");
                ok=0;
		        break;
            }
		    if (comp_car_ext(parent.pantalla.document.formas_pago.i_codigo.value,1)==1){
		        window.alert("<%=LitMsgFormPDesCarNoVal%>");
		        ok=0;
		        break;
		    }
	        if (comp_car_ext(parent.pantalla.document.formas_pago.i_descripcion.value,0)==1){
		        window.alert("<%=LitMsgFormPDesCarNoVal%>");
		        ok=0;
		        break;
		    }
            if (parent.pantalla.document.formas_pago.i_descripcion.value=="")  {
                window.alert ("<%=LitMsgDescripcionNoNulo%>");
                ok=0;
            }
		    break;
    }
    if (ok==1) {
        if (viene=="asistente"){            
            parent.pantalla.document.formas_pago.submit();
	        document.opciones.submit();
	        	        
	        parent.pantalla.document.location="formas_pago.asp?viene=asistente";
	        document.location="formas_pago_bt.asp?viene=asistente";  	    
	    }
	    else{
	        parent.pantalla.document.formas_pago.submit();
	        document.opciones.submit();
	    }       
    }
}

function Buscar() {
	parent.pantalla.document.location="formas_pago.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next";
	document.location="formas_pago_bt.asp";
}

function Eliminar(param) {
    if (window.confirm("<%=LitMsgEliminarFormaPagoConfirm%>")==true) {
        switch (param){
            case 1:
	            if (parent.pantalla.document.formas_pago.e_codigo.value=="") window.alert ("<%=LitMsgCodigoNoNulo%>");
                else
                    parent.pantalla.document.location="formas_pago.asp?mode=delete&codigo=" + parent.pantalla.document.formas_pago.e_codigo.value;
    		    break;

            case 2:
	            if (parent.pantalla.document.formas_pago.i_codigo.value=="") window.alert ("<%=LitMsgCodigoNoNulo%>");
                else
                    parent.pantalla.document.location="formas_pago.asp?mode=delete&codigo=" + parent.pantalla.document.formas_pago.i_codigo.value;
			    break;
        }
        document.location="formas_pago_bt.asp";
    }
}

function Cancelar(viene){
    if (viene=="asistente") {
	    parent.pantalla.document.location="formas_pago.asp?viene=asistente";
	    document.location="formas_pago_bt.asp?viene=asistente"; 
	}
	else{
	    parent.pantalla.document.location="formas_pago.asp";
	    document.location="formas_pago_bt.asp";  
	}
}

//****************************************************************************************************************
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

function comprobar_enter(){
    var keycode = ev.keyCode;
	//si se ha pulsado la tecla enter
	if (keycode==13){
		document.opciones.criterio.focus();
		Buscar();
	}
}

function MoverPagPM(ruta,rbt){
    if (rbt=="1") parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"1";        
  	else parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"2";	  
}

function Cerrar(){
    parent.location="../Applets/asistentePM.asp?mode=cancel";
}
</script>
<body class="body_master_ASP">
<%viene=Request.QueryString("viene")
 tipo="2"
 'ruta=ObtenerNombreFichero(tipo)
%>
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
		            <%if viene="asistente" then %>
    		            <td CLASS="CELDABOT" onclick="javascript:Guardar(<%=param%>,'asistente');">
				            <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
			            </td>
			            <%if param=1 then %>
			                <td CLASS="CELDABOT" onclick="javascript:Cancelar('asistente');">
				                <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			                </td>
			            <%else %>
			                <td CLASS="CELDABOT" onclick="javascript:Cerrar();">
				                <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			                </td>
			            <%end if %>
			    
		                <td CLASS=CELDABOT>
				            <A CLASS=CELDAREF href="javascript:MoverPagPM('<%=ruta%>','<2');"><IMG SRC="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></A>
			            </td> 
		                <td CLASS=CELDABOT>
				            <A CLASS=CELDAREF href="javascript:MoverPagPM('<%=ruta%>','1');"><IMG SRC="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></A>
			            </td>
		            <%else%>
    		            <td CLASS="CELDABOT" onclick="javascript:Guardar(<%=param%>,'noasist');">
				            <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,""%>
			            </td>
			            <td CLASS="CELDABOT" onclick="javascript:Eliminar(<%=param%>);">
				            <%PintarBotonBTRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,""%>
			            </td>
			            <td CLASS="CELDABOT" onclick="javascript:Cancelar('noasist');">
				            <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			            </td>
		            <%end if%>
		        </tr>
            </table>
        </div>
        <%if viene<>"asistente" then %>
            <div id="FILTERS_MASTER_ASP">
			    <!--<td CLASS=CELDABOT><%=LitBuscar & ": "%>-->
				    <SELECT class="IN_S" name="campos">
					    <OPTION  value="codigo"><%=LitCodigo%></OPTION>
					    <OPTION selected value="descripcion"><%=LitDescripcion%></OPTION>
				    </SELECT>
			    <!--</td>
			    <td CLASS=CELDABOT>-->
				    <SELECT class="IN_S" name="criterio">
					    <OPTION value="contiene"><%=LitContiene%></OPTION>
					    <!--<OPTION value="empieza"><%=LitComienza%></OPTION>-->
					    <OPTION value="termina"><%=LitTermina%></OPTION>
					    <OPTION value="igual"><%=LitIgual%></OPTION>
				    </SELECT>
			    <!--</td>
			    <td CLASS=CELDABOT>-->
				    <INPUT class="IN_S" type="text" name="texto" size=20 maxLength=20 value="" onKeyPress="javascript:comprobar_enter();">
				    <A CLASS=CELDAREF href="javascript:Buscar();"><IMG src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> ALT="<%=LitBuscar%>"></A>
			    <!--</td>-->
            </div>
        <%end if%>
    </div>
<%ImprimirPie_bt%>
</form>
</BODY>
</HTML>