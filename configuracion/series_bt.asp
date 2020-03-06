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

<TITLE><%=LitTituloSD%></TITLE>
    <% 
        ' ' 07/05/2019 Se realiza cambios de ciberseguridad
        ' - enc.EncodeForJavascript(param) -> Cross Site Scripting (XSS)
        ' - enc.EncodeForHtmlAttribute(param) -> Cross Site Scripting (XSS)
        ' - limpiaCadena() -> Inyección SQL
        dim  enc
        set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
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
            if (parent.pantalla.document.series.e_Nserie.value=="") {
                window.alert ("<%=LitMsgCodigoNoNulo%>");
                ok=0;
            }
	        if (ok==1 && (comp_car_ext(parent.pantalla.document.series.e_Nserie.value,1)==1 || parent.pantalla.document.series.e_Nserie.value.indexOf(" ")!=-1)){
		        window.alert("<%=LitMsgCodSerieDesCarNoVal%>");
		        ok=0;
	        }
            if (ok==1 && parent.pantalla.document.series.e_Nombre.value=="")  {
                window.alert ("<%=LitMsgNombreNoNulo%>");
                ok=0;
            }

            if (ok==1 && parent.pantalla.document.series.e_empresa.value=="")  {
                window.alert ("<%=LitMsgEmpresaNoNulo%>");
                ok=0;
            }

	        if (ok==1 && parent.pantalla.document.series.e_documento.value=="")  {
                window.alert ("<%=LitMsgTipoDocumentoNoNulo%>");
                ok=0;
            }
		    if (ok==1 && parent.pantalla.document.series.e_Contador.value=="")  {
		 	    window.alert ("<%=LitMsgContadorNoNulo%>");
                ok=0;
		    }
		    else {
		 	    if (isNaN(parent.pantalla.document.series.e_Contador.value)) {
		   		    window.alert("<%=LitMsgContadorNumerico%>");
		   		    ok=0;
		 	    }
		    }
            break;

        case 2:
            if (parent.pantalla.document.series.i_Nserie.value=="") {
                window.alert ("<%=LitMsgCodigoNoNulo%>");
                ok=0;
            }
	        if (ok==1 && (comp_car_ext(parent.pantalla.document.series.i_Nserie.value,1)==1 || parent.pantalla.document.series.i_Nserie.value.indexOf(" ")!=-1)){
		        window.alert("<%=LitMsgCodSerieDesCarNoVal%>");
		        ok=0;
	        }
            if (ok==1 && parent.pantalla.document.series.i_Nombre.value=="")  {
                window.alert ("<%=LitMsgNombreNoNulo%>");
                ok=0;
            }

            if (ok==1 && parent.pantalla.document.series.i_empresa.value=="")  {
                window.alert ("<%=LitMsgEmpresaNoNulo%>");
                ok=0;
            }
	        if (ok==1 && parent.pantalla.document.series.i_documento.value=="")  {
                window.alert ("<%=LitMsgTipoDocumentoNoNulo%>");
                ok=0;
            }
		    if (ok==1 && parent.pantalla.document.series.i_Contador.value=="")  {
    		 	window.alert ("<%=LitMsgContadorNoNulo%>");
                ok=0;
		    }
		    else {
		 	    if (isNaN(parent.pantalla.document.series.i_Contador.value)) {
		   		    window.alert("<%=LitMsgContadorNumerico%>");
		   		    ok=0;
		 	    }
		    }
            break;
    }
    if (ok==1) {
        if (viene=="asistente"){
            parent.pantalla.document.series.action="series.asp?viene=asistente";
		    parent.pantalla.document.series.submit();
            document.location="series_bt.asp";		
	        document.opciones.submit();	    
	    }
	    else{
	        parent.pantalla.document.series.submit();
	        document.opciones.submit();	 
	    }
    }
}

function Buscar() {
	parent.pantalla.document.location="series.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next";
	document.location="series_bt.asp";
}

function Eliminar(param) {
    if (window.confirm("<%=LitMsgEliminarSerieConfirm%>")==true) {
        switch (param){
            case 1:
    	        if (parent.pantalla.document.series.e_Nserie.value=="") window.alert ("<%=LitMsgCodigoNoNulo%>");
                else parent.pantalla.document.location="series.asp?mode=delete&Nserie=" + parent.pantalla.document.series.e_Nserie.value;
    		    break;

            case 2:
	            if (parent.pantalla.document.series.i_Nserie.value=="") window.alert ("<%=LitMsgCodigoNoNulo%>");
                else parent.pantalla.document.location="series.asp?mode=delete&Nserie=" + parent.pantalla.document.series.i_Nserie.value;
			    break;
        }
        document.location="series_bt.asp";
    }
}

function Cancelar()
{
	parent.pantalla.document.location="series.asp?npagina="+parent.pantalla.document.series.h_npagina.value;
	document.location="series_bt.asp";
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
<form name="opciones" method="post">
<% viene=Request.QueryString("viene")
 tipo="1"
 'ruta=ObtenerNombreFichero(tipo)
  
if request("mode")="edit" then
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
		                <td CLASS="CELDABOT" onclick="javascript:Cerrar();">
			                <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
		                </td>
		                <td CLASS=CELDABOT>
				            <A CLASS=CELDAREF href="javascript:MoverPagPM('<%=ruta%>','2');"><IMG SRC="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></A>
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
		                <td CLASS="CELDABOT" onclick="javascript:Cancelar();">
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
					    <OPTION value="nserie"><%=LitSerie%></OPTION>
					    <OPTION selected value="nombre"><%=LitNombre%></OPTION>
					    <OPTION value="tipo_documento"><%=LitDocumento%></OPTION>
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