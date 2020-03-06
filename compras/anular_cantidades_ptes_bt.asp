<%@ Language=VBScript %>

<script id="DebugDirectives" runat="server" language="javascript">
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../styles/Master.css.inc" -->

<!--#include file="pedidos_pro.inc" -->
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<%

	if request.querystring("viene")>"" then
		viene=limpiaCadena(request.querystring("viene"))
	elseif request.form("viene")>"" then
		viene=request.form("viene")
	end if

	if request.querystring("referencia")>"" then
		referencia=limpiaCadena(request.querystring("referencia"))
	elseif request.form("referencia")>"" then
		referencia=request.form("referencia")
	end if
%>
<script type="text/javascript" language="javascript" src="../jfunciones.js"></script>
<script type="text/javascript" language="javascript" src="/lib/js/shortKey.js"></script>
<script type="text/javascript" language="javascript">
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1);

function Imprimir() {
	if (pr) //NS4, IE5
		parent.pantalla.print()
		//vbImprimir()
	else if (da && !mac) // IE4 (Windows)
		//vbImprimir()
		alert("<%=LitNoImprime%>");
	else // Otros Navegadores
		alert("<%=LitNoImprime%>");
	return false;
}
function comprobar_enter() {
    document.opciones.criterio.focus();
    Buscar();
}

//Validación de campos numéricos y fechas.
function ValidarCampos() {
	ok=1
	if (!checkdate(parent.pantalla.document.anular_cantidades_ptes.fdesde)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}
	if (!checkdate(parent.pantalla.document.anular_cantidades_ptes.fhasta)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}
	if (!checkdate(parent.pantalla.document.anular_cantidades_ptes.fentregadesde)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}
	if (!checkdate(parent.pantalla.document.anular_cantidades_ptes.fentregahasta)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}
	if (ok==1) return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
	    case "select1":
	        switch (pulsado) {
	            case "imp": //Aceptar
	                if (ValidarCampos()) {
	                    parent.pantalla.document.anular_cantidades_ptes.action = "anular_cantidades_ptes.asp?mode=" + pulsado;
	                    parent.pantalla.document.anular_cantidades_ptes.submit();
	                    document.location = "anular_cantidades_ptes_bt.asp?mode=" + pulsado;
	                }
	                break;
	            case "select1": //Cancelar
	                parent.pantalla.document.location = "anular_cantidades_ptes.asp?mode=select1";
	                document.location = "anular_cantidades_ptes_bt.asp?mode=select1";
	                break;
	        }
	        break;
		case "imp":
			switch (pulsado) {
				case "todos": //Seleccionar todos los registros
					nregistros=parent.pantalla.document.anular_cantidades_ptes.h_nregistros.value;
					for (i=1;i<=nregistros;i++) {
						nombre="checkElim" + i;
						parent.pantalla.document.anular_cantidades_ptes.elements[nombre].checked=true;
					}
					break;

				case "ninguno": //Editar registro
					nregistros=parent.pantalla.document.anular_cantidades_ptes.h_nregistros.value;
					for (i=1;i<=nregistros;i++) {
						nombre="checkElim" + i;
						parent.pantalla.document.anular_cantidades_ptes.elements[nombre].checked=false;
					}
					break;

				case "delete": 
					if (window.confirm("<%=LitMsgBorrarSel%>")==true) {
						nregistros=parent.pantalla.document.anular_cantidades_ptes.h_nregistros.value;
						parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						parent.pantalla.document.anular_cantidades_ptes.action="anular_cantidades_ptes.asp?mode=delete";
						parent.pantalla.document.anular_cantidades_ptes.submit();
						document.location="anular_cantidades_ptes_bt.asp?mode=oculto";
					}
					break;
			
				case "imprimir": //Volver atrás
					if (parent.pantalla.document.anular_cantidades_ptes.NumRegsTotal.value>0){
						parent.pantalla.focus();
						Imprimir();
						//parent.pantalla.document.location="anular_cantidades_ptes.asp?mode=select1";
						//document.location="anular_cantidades_ptes_bt.asp?mode=select1";
					}
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.listado_pedidos_pendientes.NumRegsTotal.value)>=parseInt(parent.pantalla.document.listado_pedidos_pendientes.maxpdf.value))
						alert("<%=LitMsgRegPDF%>");
					else {
						if (parent.pantalla.document.listado_pedidos_pendientes.NumRegsTotal.value>0){
							parent.pantalla.document.listado_pedidos_pendientes.action="listado_pedidos_pendientes_pdf.asp?mode=browse";
							parent.pantalla.document.listado_pedidos_pendientes.submit();
							document.location="listado_pedidos_pendientes_bt.asp?mode=pdf&viene=<%=enc.EncodeForJavascript(viene)%>&referencia=" + parent.pantalla.document.listado_pedidos_pendientes.referencia.value;
						}
					}
					break;
				case "cancel": //Volver atrás
					parent.pantalla.document.location="anular_cantidades_ptes.asp?mode=select1";
					document.location="anular_cantidades_ptes_bt.asp?mode=select1";
					break;
			}
			break;
	}
}

function Buscar() {
	parent.pantalla.document.location="anular_cantidades_ptes.asp?mode=imp&campo=" + document.opciones.campos.value + 
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=" + "&nproveedor=" + parent.pantalla.document.anular_cantidades_ptes.nproveedor.value + "&viene=" + parent.pantalla.document.anular_cantidades_ptes.viene.value;
	document.location="anular_cantidades_ptes_bt.asp?viene=" + parent.pantalla.document.anular_cantidades_ptes.viene.value;	
}
</script>
<%if viene="tienda" then
	cadena="topmargin='0'"
else
	cadena=""
end if%>
<body class="body_master_ASP" <%=cadena%>>
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
    <input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
	<%if viene="tienda" then%>
        <div id="PageFooter_ASP">
        <div id="FILTERS_MASTER_ASP">
            <!--<table id="BUTTONS_CENTER_ASP">
		        <tr>
				<td class=CELDABOT><%=LitBuscar & ": "%>-->
					<select class="IN_S" name="campos">
						<option selected="selected" value="npedido"><%=LitPedido%></option>
						<option value="fecha"><%=LitFecha%></option>
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
				<!--</td>
				<td class=CELDABOT>-->
                   <a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
				<!--</td>
			</tr>
		</table>-->
        </div>
        </div>
	<%else%>
        <div id="PageFooter_ASP">
        <div id="ControlPanelFooter_ASP">
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
            <%
				if mode="select1" then
					%>
				    <td id="idaccept" class="CELDABOT" onclick="javascript:Accion('select1','imp');">
					    <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				    </td>
                    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('select1','select1');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				    <%
                elseif mode="imp" then
					%>
				    <td id="idSelectAll" class="CELDABOT" onclick="javascript:Accion('imp','todos');">
					    <%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,LITBOTSELTODOTITLE%>
				    </td>
				    <td id="idSelectNothing" class="CELDABOT" onclick="javascript:Accion('imp','ninguno');">
					    <%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,LITBOTDSELTODOTITLE%>
				    </td>
				    <td id="idaccept" class="CELDABOT" onclick="javascript:Accion('imp','delete');">
					    <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				    </td>
					<%if viene<>"articulos" then%>
				        <td id="idreturn" class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
					        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				        </td>
					<%end if%>
				<%elseif mode="pdf" then%>
			        <td id="idreturn" class="CELDABOT" onclick="javascript:Accion('pdf','back');">
				        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			        </td>
				<%elseif mode="oculto" then	%>
				    <td class="CELDABOT" align="center">
					    &nbsp;
				    </td>
				<%end if%>
			</tr>
		</table>
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