<%@ Language=VBScript %>
<%'JMAN 13-06-03: Migración a monobase'%>
<% dim enc
   set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html LANG="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<title><%=LitTituloCaj%></title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta HTTP-EQUIV="Content-Type" Content="text/html; charset=<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
function Buscar() {
	parent.pantalla.fr_Tabla.document.cajas_det.action="cajas_det.asp?mode=search&lote=1&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value;
	parent.pantalla.fr_Tabla.document.cajas_det.submit();
	document.location="cajas_bt.asp";
}

function Cancelar()
{
	parent.pantalla.fr_Tabla.document.cajas_det.action="cajas_det.asp?mode=browse?lote=1";
	parent.pantalla.fr_Tabla.document.cajas_det.submit();
}

//****************************************************************************************************************
function comprobar_enter(){
		document.opciones.criterio.focus();
		Buscar();
}

function MoverPagPM(ruta,rbt){   	
  	if (rbt=="1") parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"1";        
  	else parent.location="../Applets/asistentePM.asp?mode=sig&modd="+ruta+"-"+"2";	      
}
</script>
<body class="body_master_ASP">
<form name="opciones" method="post">
<%if request("mode")="edit" then
    param=1
else
    param=2
end if
    viene=enc.EncodeForJavascript(Request.QueryString("viene"))
'    tipo="1"    
'    ruta=ObtenerNombreFichero(tipo) 
%>
<div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_left_ASP" >
        <table id="BUTTONS_CENTER_ASP">
		    <tr>
                <%if viene<>"asistente" then %>
			    <td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar();">
				    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			    </td>
                <%else%>
			        <td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar();">
                        <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,""%>
			        </td>	
			         <td id="idprevious" class="CELDABOT" onclick="javascript:MoverPagPM('<%=enc.EncodeForJavascript(ruta)%>','2');">
                        <%PintarBotonBTLeft LITBOTANTERIOR,ImgAnterior,ParamImgAnterior,""%>
			         </td> 
		             <td id="idnext" class="CELDABOT" onclick="javascript:MoverPagPM('<%=enc.EncodeForJavascript(ruta)%>','1');">
                        <%PintarBotonBTLeft LITBOTSIGUIENTE,ImgSiguiente,ParamImgSiguiente,""%>
			         </td> 		   
                <%end if%>
            </tr>
        </table>
    </div>
    <%if viene<>"asistente" then %>
        <div id="FILTERS_MASTER_ASP">
		    <select class="IN_S" name="campos">
			    <option  value="codigo"><%=LitCodigo%></option>
			    <option  selected value="descripcion"><%=LitDescripcion%></option>
			    <option  value="serie"><%=LitSerie%></option>
			    <option  value="contable"><%=LitCuenta%></option>
			    <option  value="tienda"><%=LitTienda%></option>
		    </select>
		    <select class="IN_S" name="criterio">
			    <option value="contiene"><%=LitContiene%></option>
			    <option value="termina"><%=LitTermina%></option>
			    <option value="igual"><%=LitIgual%></option>
		    </select>
		    <input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
		    <a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>"/></a>
        </div>
</div>
    <%end if%>
    <table style="width:100%;height:42px;vertical-align:bottom;" align="center">
    <tr>
    <td style="width:100%;height:42px; vertical-align:bottom; text-align:center;">
        <%ImprimirPie_bt%>
    </td>
    </tr>
    </table>
</form>
</body>
</html>