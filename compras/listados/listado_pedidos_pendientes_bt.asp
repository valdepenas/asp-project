<%@ Language=VBScript %>
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
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
</head>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  
<!--#include file="../../calculos.inc" -->
<!--#include file="../../constantes.inc" -->

<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->

<!--#include file="../pedidos_pro.inc" -->

<!--#include file="../../styles/Master.css.inc" -->

<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
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
//Validación de campos numéricos y fechas.
function ValidarCampos() {

	ok=1
	if (!checkdate(parent.pantalla.document.listado_pedidos_pendientes.fdesde)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}
	if (!checkdate(parent.pantalla.document.listado_pedidos_pendientes.fhasta)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}
	if (!checkdate(parent.pantalla.document.listado_pedidos_pendientes.fentregadesde)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}
	if (!checkdate(parent.pantalla.document.listado_pedidos_pendientes.fentregahasta)) {
		window.alert("<%=LitMsgFechaFecha%>");
		ok=0;
	}

	if (ok==1){
		return true;
	}
}


//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "select1":
			switch (pulsado) {
				case "imp": //Aceptar
				    if (ValidarCampos()) {
				        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
						parent.pantalla.document.listado_pedidos_pendientes.action="listado_pedidos_pendientesResultado.asp?mode=" + pulsado;
						parent.pantalla.document.listado_pedidos_pendientes.submit();
						document.location = "listado_pedidos_pendientes_bt.asp?mode=" + pulsado;
					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.location="listado_pedidos_pendientes.asp?mode=select1";
					document.location="listado_pedidos_pendientes_bt.asp?mode=select1";
					break;
			}
			break;
		case "imp":
			switch (pulsado) {
				case "imprimir": //Volver atrás
					if (parent.pantalla.document.listado_pedidos_pendientesResultado.NumRegsTotal.value>0){
						parent.pantalla.focus();
						Imprimir();
						//parent.pantalla.document.location="listado_pedidos_pendientes.asp?mode=select1";
						//document.location="listado_pedidos_pendientes_bt.asp?mode=select1";
					}
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.listado_pedidos_pendientesResultado.NumRegsTotal.value)>=parseInt(parent.pantalla.document.listado_pedidos_pendientesResultado.maxpdf.value)) {
						alert("<%=LitMsgRegPDF%>");
					} else {
						if (parent.pantalla.document.listado_pedidos_pendientesResultado.NumRegsTotal.value>0){
							parent.pantalla.document.listado_pedidos_pendientesResultado.action="listado_pedidos_pendientes_pdf.asp?mode=browse";
							parent.pantalla.document.listado_pedidos_pendientesResultado.submit();
							document.location = "listado_pedidos_pendientes_bt.asp?mode=pdf&viene=<%=enc.EncodeForJavascript(viene)%>&referencia=" + parent.pantalla.document.listado_pedidos_pendientesResultado.referencia.value;
						}
					}
					break;
				case "cancel": //Volver atrás
					parent.pantalla.document.location="listado_pedidos_pendientes.asp?mode=select1";
					document.location="listado_pedidos_pendientes_bt.asp?mode=select1";
					break;
				
			}
			break;
		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
					viene="<%=viene%>";
					if (viene!="articulos"){
						//parent.pantalla.document.location="listado_pedidos_pendientes.asp?mode=select1";
						//document.location="listado_pedidos_pendientes_bt.asp?mode=select1";
						parent.document.location="../../central.asp?pag1=compras/listados/listado_pedidos_pendientes.asp&mode=select1&pag2=compras/listados/listado_pedidos_pendientes_bt.asp";
					}
					else{
					    parent.pantalla.document.location = "listado_pedidos_pendientes.asp?mode=imp&referencia=<%=enc.EncodeForJavascript(referencia)%>&viene=articulos";
						document.location="listado_pedidos_pendientes_bt.asp?mode=imp&viene=articulos";
					}
					break;
			}
			break;

	}
}

function Buscar() {
		parent.pantalla.document.location="listado_pedidos_pendientes.asp?mode=imp&campo=" + document.opciones.campos.value + 
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=" + "&nproveedor=" + parent.pantalla.document.listado_pedidos_pendientes.nproveedor.value + "&viene=" + parent.pantalla.document.listado_pedidos_pendientes.viene.value;
		document.location="listado_pedidos_pendientes_bt.asp?viene=" + parent.pantalla.document.listado_pedidos_pendientes.viene.value;	
}

</script>

<%
if viene="tienda" then
	cadena="topmargin='0'"
else
	cadena=""
end if%>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post" action="javascript:if ('<%=enc.EncodeForJavascript(viene)%>'=='tienda') {document.opciones.criterio.focus();Buscar()}">
	<%if viene="tienda" then%>
		<table width='100%' border='0' cellspacing="1" cellpadding="1">
			<tr>
				<td class=CELDABOT><%=LitBuscar & ": "%>
					<select class=INPUT name="campos">
						<option selected value="npedido"><%=LitPedido%></option>
						<option value="fecha"><%=LitFecha%></option>
					</select>
				</td>
				<td class=CELDABOT>
					<select class=INPUT name="criterio">
						<option value="contiene"><%=LitContiene%></option>
						<!--<option value="empieza"><%=LitComienza%></option>-->
						<option value="termina"><%=LitTermina%></option>
						<option value="igual"><%=LitIgual%></option>
					</select>
				</td>
				<td class=CELDABOT>
					<input class=INPUT type="text" name="texto" size=20 maxlength=20 value="">
				</td>
				<td class=CELDABOT>
				   <a class='CELDAREF' href="javascript:Buscar();"><img src="../images/<%=ImgBuscar_bt%>" <%=ParamImgBuscar_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
				</td>
			</tr>
		</table>
	<%else%>
		<div id="PageFooter_ASP" >
            <div id="ControlPanelFooter_ASP" >
	            <table id="BUTTONS_CENTER_ASP" >
			        <tr><%
				        if mode="select1" then
					        %>
				            <td class="CELDABOT" onclick="javascript:Accion('select1','imp');">
					            <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				            </td>
				            <td class="CELDABOT" onclick="javascript:Accion('select1','select1');">
					            <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				            </td>
				        <%elseif mode="imp" then
					        %>	
				            <td class="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					            <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				            </td>
				            <td class="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					            <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,IMPRIMIRLISTADOTITLE%>
				            </td>		
					        <%if viene<>"articulos" then%>
                                <td class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
				                    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			                    </td>
					        <%end if%>
				        <%elseif mode="pdf" then
					        %>
				            <td class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				            </td>
				        <%end if%>
			        </tr>
		        </table>
            </div>
        </div>
    <%end if%>
</form>
</body>
</html>