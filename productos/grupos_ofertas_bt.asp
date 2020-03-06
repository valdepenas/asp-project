<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
    <% dim  enc
    set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupa_colores.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<TITLE><%=LitTituloColor%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script Language="JavaScript">
function Guardar(param) {
   ok=1;
   switch(param){
      case 1:
         if (parent.pantalla.document.grupos_ofertas.e_codigo.value=="") {
            window.alert ("<%=LitMsgCodigoNoNulo%>");
            ok=0;
		break;
         }
		if (comp_car_ext(parent.pantalla.document.grupos_ofertas.e_codigo.value,1)==1){
		window.alert("<%=LitMsgColoDesCarNoVal%>");
		ok=0;
			break;
		}

         if (parent.pantalla.document.grupos_ofertas.e_descripcion.value=="")  {
            window.alert ("<%=LitMsgDescripcionNoNulo%>");
            ok=0;
		break;
         }
	if (comp_car_ext(parent.pantalla.document.grupos_ofertas.e_descripcion.value,0)==1){
		window.alert("<%=LitMsgColoDesCarNoVal%>");
		ok=0;
			break;
		}

		 break;
      case 2:
	  
         if (parent.pantalla.document.grupos_ofertas.i_codigo.value=="") {
            window.alert ("<%=LitMsgCodigoNoNulo%>");
            ok=0;
			break;
         }
		if (comp_car_ext(parent.pantalla.document.grupos_ofertas.i_codigo.value,3)==1){
			window.alert("<%=LitMsgColoDesCarNoVal%>");
			ok=0;
			break;
		}

        if (parent.pantalla.document.grupos_ofertas.i_descripcion.value=="")  {
            window.alert ("<%=LitMsgDescripcionNoNulo%>");
            ok=0;
			break;
         }
		if (comp_car_ext(parent.pantalla.document.grupos_ofertas.i_descripcion.value,0)==1){
			window.alert("<%=LitMsgColoDesCarNoVal%>");
			ok=0;
			break;
		}

         break;
   }
   if (ok==1) {
      parent.pantalla.document.grupos_ofertas.submit();
	  document.opciones.submit();
   }

}

function Buscar() {
	parent.pantalla.document.location="grupos_ofertas.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next";
	document.location = "grupos_ofertas_bt.asp";
}

function Eliminar(param) {
   if (window.confirm("<%=LitMsgEliminarColorConfirm%>")==true) {
      switch (param){
         case 1:
	        if (parent.pantalla.document.grupos_ofertas.e_codigo.value=="") {
               window.alert ("<%=LitMsgCodigoNoNulo%>");
            }
            else {
               parent.pantalla.document.location="grupos_ofertas.asp?mode=delete&codigo=" + parent.pantalla.document.grupos_ofertas.e_codigo.value;
            }
    		break;

         case 2:
	        if (parent.pantalla.document.grupos_ofertas.i_codigo.value=="") {
               window.alert ("<%=LitMsgCodigoNoNulo%>");
            }
            else {
               parent.pantalla.document.location="grupos_ofertas.asp?mode=delete&codigo=" + parent.pantalla.document.grupos_ofertas.i_codigo.value;
            }
			break;
      }
        document.location = "grupos_ofertas_bt.asp";
   }
}

function Cancelar() {
	parent.pantalla.document.location="grupos_ofertas.asp?npagina="+parent.pantalla.document.grupos_ofertas.h_npagina.value;
	document.location = "grupos_ofertas_bt.asp";
}

//****************************************************************************************************************
function comprobar_enter(){
	//si se ha pulsado la tecla enter
	
		document.opciones.criterio.focus();
		Buscar();
	
}
</script>
<body class="body_master_ASP">
<form name="opciones" method="post">
 <%if request("mode")="edit" then
			              param=1
			           else
			              param=2
			           end if
     
mode=enc.EncodeForJavascript(Request.QueryString("mode"))
%>
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
          <td id="idsave" class="CELDABOT" onclick="javascript:Guardar(<%=enc.EncodeForJavascript(param)%>);">
				    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDAR%>
			    </td>
            <%
                visible=""
                if mode<>"edit" then
                    visible=" style='display:none;' "
                end if%>
			<td id="iddelete" class="CELDABOT" <%=visible%> onclick="javascript:Eliminar(<%=enc.EncodeForJavascript(param)%>);">
				    <%PintarBotonBTLeft LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
			    </td>
			<td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar();">
				    <%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
			    </td>
			 </table>
        </div>
    
        <div id="FILTERS_MASTER_ASP">
			
				<SELECT class="IN_S"  name="campos">
					<OPTION value="codigo"><%=LitCodigo%></OPTION>
					<OPTION selected value="nombre"><%=LitDescripcion%></OPTION>
				</SELECT>
		
				<select class="IN_S" name="criterio">
					<OPTION value="contiene"><%=LitContiene%></OPTION>
					<!--<OPTION value="empieza"><%=LitComienza%></OPTION>-->
					<OPTION value="termina"><%=LitTermina%></OPTION>
					<OPTION value="igual"><%=LitIgual%></OPTION>
				</SELECT>
			
			
				<input id="KeySearch" class="IN_S" type="texto" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();">
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			
		</div>
    </div>
<%ImprimirPie_bt%>
</form>
</BODY>
</HTML>