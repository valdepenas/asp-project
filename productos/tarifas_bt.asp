<%@ Language=VBScript %>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<%'' JCI 18/06/2003 : MIGRACION A MONOBASE%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../XSSProtection.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../calculos.inc" -->

<!--#include file="tarifas.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
function Guardar(param) {
   ok=1;
   switch(param){
      case 1:
         if (parent.pantalla.document.tarifas.e_codigo.value=="") {
            window.alert ("<%=LitMsgCodigoNoNulo%>");
            ok=0;
			break;
         }

	    if (comp_car_ext(parent.pantalla.document.tarifas.e_codigo.value,1)==1){
		    window.alert("<%=LitMsgTarDesCarNoVal%>");
		    ok=0;
			break;
	    }

        if (parent.pantalla.document.tarifas.e_descripcion.value=="")  {
            window.alert ("<%=LitMsgDescripcionNoNulo%>");
            ok=0;
            break;
        }
	    if (comp_car_ext(parent.pantalla.document.tarifas.e_descripcion.value,0)==1){
		    window.alert("<%=LitMsgTarDesCarNoVal%>");
		    ok=0;
			break;
	    }
		break;
		
      case 2:
        if (parent.pantalla.document.tarifas.i_codigo.value=="") {
            window.alert ("<%=LitMsgCodigoNoNulo%>");
            ok=0;
			break;
        }

		if (comp_car_ext(parent.pantalla.document.tarifas.i_codigo.value,1)==1){
		    window.alert("<%=LitMsgTarDesCarNoVal%>");
		    ok=0;
			break;
		}

         if (parent.pantalla.document.tarifas.i_descripcion.value=="")  {
            window.alert ("<%=LitMsgDescripcionNoNulo%>");
            ok=0;
			break;
         }

		if (comp_car_ext(parent.pantalla.document.tarifas.i_descripcion.value,0)==1){
		    window.alert("<%=LitMsgTarDesCarNoVal%>");
		    ok=0;
			break;
		}
		break;
   }
   if (ok==1) {
      parent.pantalla.document.tarifas.submit();
	  document.opciones.submit();
   }
}

function comprobar_enter() {
    //si se ha pulsado la tecla enter
    //if (window.event.keyCode==13){
    document.opciones.criterio.focus();
    Buscar();
    //}
}

function Buscar() {
	parent.pantalla.document.location="tarifas.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next";
	document.location="tarifas_bt.asp";
}

function Eliminar(param) {
   if (window.confirm("<%=LitMsgEliminarTarifaConfirm%>")==true) {
	      switch (param){
    	     case 1:
	    	    if (parent.pantalla.document.tarifas.e_codigo.value=="") window.alert ("<%=LitMsgCodigoNoNulo%>");
   		        else parent.pantalla.document.location="tarifas.asp?mode=delete&mode2=edit&codigo=" + parent.pantalla.document.tarifas.e_codigo.value;
    			break;

	         case 2:
		        if (parent.pantalla.document.tarifas.i_codigo.value=="")  window.alert ("<%=LitMsgCodigoNoNulo%>");
	            else parent.pantalla.document.location="tarifas.asp?mode=delete&mode2=add&codigo=" + parent.pantalla.document.tarifas.i_codigo.value;
				break;
	      }
      document.location="tarifas_bt.asp";
   }
}

function Cancelar() {
	parent.pantalla.document.location="tarifas.asp?mode=add";
	document.location="tarifas_bt.asp";
}


</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<%
modo=request.QueryString("mode")


    if modo="aplicaTots" or modo="aplicaTots2" then

        esdto=limpiaCadena(request.QueryString("esdto"))
        pvpdto=limpiaCadena(request.QueryString("dto")) 
        precio= limpiaCadena(request.QueryString("precio"))  
        if precio>"" and pvpdto="" then
            pvpdto=precio
        end if   
        bloqueo="CHECKED"
        if modo="aplicaTots" then temp="_temp1"
        if modo="aplicaTots2" then temp="_temp2"

       'Calculamos el campo precio.
       select case esdto
        case 0
            sql2=" ,precio="& replace(precio,",",".") 
        case 1
            sql2=" ,precio= round(PVPORIGEN+(PVPORIGEN*" & replace(pvpdto,",",".") & ")/100,ndecimales ) "                      
        case 2
            sql2=" ,precio= round(coste+(coste*" & replace(pvpdto,",",".") & ")/100,ndecimales ) "                         
       end select    
              
        set rstApliTot = server.CreateObject("ADODB.Recordset")
        sql = "update [" & session("usuario") & temp & "] set pvpdto="&replace(pvpdto,",",".")&",esdto="&esdto&", bloqueo='"&bloqueo&"'" 
        sql = sql & sql2 ''&" where BLOQUEO='' "       
	    rstApliTot.open sql,session("dsn_cliente")
        if rstApliTot.state<>0 then rstApliTot.close
        set rstApliTot=nothing
	    
    end if

    if request("mode")="edit" then
        param=1
    else
        param=2
    end if%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
                <%if not mode="XXXX" then%>
			            <td id="idsave" class="CELDABOT" onclick="javascript:Guardar(<%=param%>);">
					        <%PintarBotonBT LITBOTGUARDARCAB,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				        </td>
				        <td id="iddelete" class="CELDABOT" onclick="javascript:Eliminar(<%=param%>);">
					        <%PintarBotonBTRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				        </td>
				        <td id="idcancel" class="CELDABOT" onclick="javascript:Cancelar();">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
                <%end if%>    
        </tr>
          </table>
        </div>

         <div id="FILTERS_MASTER_ASP">
			    <!--<td class=CELDABOT><%=LitBuscar & ": "%>-->
				    <select class="IN_S" name="campos">
					    <option  value="codigo"><%=LitCodigo%></option>
					    <option selected value="descripcion"><%=LitDescripcion%></option>
					    <option  value="observaciones"><%=LitObservaciones%></option>
				    </select>
			    <!--</td>
			    <td class=CELDABOT>-->
				    <select class="IN_S" name="criterio">
					    <option value="contiene"><%=LitContiene%></option>
					    <!--<option value="empieza"><%=LitComienza%></option>-->
					    <option value="termina"><%=LitTermina%></option>
					    <option value="igual"><%=LitIgual%></option>
				    </select>
			    <!--</td>
			    <td class=CELDABOT>-->
				    <input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				    <a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			    <!--</td>
		    </tr>
	    </table>-->
    </div>
</div>
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