<%@ Language=VBScript %>
<%
' JCI 17/04/2003 : Se añade control de la caché
'                  Se añade el include del control de impresion
' IML : 27/11/03 : Control de Impresion (controlimpresion.inc)
'JMG 21/04/2004 : Se añade el mensaje de "espere por favor"

dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../../calculos.inc" -->
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../tickets.inc" -->
<!--#include file="../../styles/Master.css.inc" -->

<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1);

function cambiarfecha(fecha,modo){
	var fecha_ar=new Array();

	if (fecha!=""){
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length){
			if (fecha.substring(l,l+1)=='/'){
				suma++;
				fecha_ar[suma]="";
			}
			else{
				if (fecha.substring(l,l+1)!='') fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
			}
			l++;
		}
		if (suma!=2) {
			window.alert("<%=LitFechaMal%> " + modo );
			return false;
		}
		else {
			nonumero=0;
			while (suma>=0 && nonumero==0){
				if (isNaN(fecha_ar[suma])) nonumero=1;
				if (fecha_ar[suma].length>2 && suma!=2) nonumero=1;
				if (fecha_ar[suma].length>4 && suma==2) nonumero=1;
				suma--;
			}

			if (nonumero==1){
				window.alert("<%=LitFechaMal%> " + modo);
				return false;
			}
		}
	}
	return true;
}

function Imprimir() {
	if (pr) //NS4, IE5
		parent.pantalla.print()
	else if (da && !mac) // IE4 (Windows)
		alert("<%=LitNoImprime%>");
	else // Otros Navegadores
		alert("<%=LitNoImprime%>");
	return false;
}

//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if(parent.pantalla.document.estad_dia_tickets.Dfecha.value=="" || parent.pantalla.document.estad_dia_tickets.Hfecha.value==""){
		window.alert("<%=LitDebeExistirPeriodoFechas%>");
		return false;
	}
    
    //FLM:20090618:esta comprobación estaba comentada. la descomento ya que si la fecha es incorrecta el procedimiento del listado falla
	if(!checkdate(parent.pantalla.document.estad_dia_tickets.Dfecha)){
		window.alert("<%=LitMsgFechaDesde%>");
		return false;
	}
    
    //FLM:20090618:esta comprobación estaba comentada. la descomento ya que si la fecha es incorrecta el procedimiento del listado falla
	if(!checkdate(parent.pantalla.document.estad_dia_tickets.Hfecha)){
		window.alert("<%=LitMsgFechaHasta%>");
		return false;
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "param":
			switch (pulsado) {
				case "select": //Aceptar
					if (ValidarCampos()) {
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						parent.pantalla.document.estad_dia_tickets.action="estad_dia_ticketsResultado.asp?mode=imp";
						parent.pantalla.document.estad_dia_tickets.submit();
						document.location="estad_dia_tickets_bt.asp?mode=imp";
					}
					break;
			}
			break;
		case "imp":
			switch (pulsado) {
				case "cancel": //Volver atrás
					parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
					cadena="estad_dia_tickets.asp?mode=param";
					parent.pantalla.document.estad_dia_ticketsResultado.action=cadena;
					parent.pantalla.document.estad_dia_ticketsResultado.submit();
					document.location="estad_dia_tickets_bt.asp?mode=param";
					break;
					
				case "imprimir": //Volver atrás
					parent.pantalla.focus();
					parent.pantalla.print();
					break;
					
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.estad_dia_ticketsResultado.NumRegs.value)>=parseInt(parent.pantalla.document.estad_dia_ticketsResultado.maxpdf.value))
						alert("<%=LitLimitePDF%>");
					else {
						parent.pantalla.document.estad_dia_ticketsResultado.action = "estad_dia_tickets_pdf.asp?mode=param";
						parent.pantalla.document.estad_dia_ticketsResultado.submit();
						cadena="estad_dia_tickets_bt.asp?mode=pdf";
						document.location=cadena;
					}
					break;
			}
			break;

		case "pdf":
			switch (pulsado) {
				case "back": //Volver a la pantalla anterior
					parent.document.location = "../../central.asp?pag1=ventas/listados/estad_dia_tickets.asp&pag2=ventas/listados/estad_dia_tickets_bt.asp&mode=param";
					break;
			}
			break;
	}
}
</script>

<body class="body_master_ASP">
<%
if request.querystring("caju")>"" then
	caju=limpiaCadena(request.querystring("caju"))
else
	caju=request.Form("caju")
end if

'**RGU 23/1/2007**
if request.querystring("ndoc")&"">"" then
	tpv=limpiaCadena(request.querystring("ndoc"))
else
	tpv=request.Form("tpv")
end if%>
<input type="hidden" name="tpv" value="<%=enc.EncodeForHtmlAttribute(null_s(tpv))%>" />
<%'**RGU
mode=enc.EncodeForHtmlAttribute(null_s(Request.QueryString("mode")))%>                                                                         
<form name="opciones" method="post">
	<div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
		            <%if mode="param" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('param','select');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
			        <%elseif mode="imp" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <%if tpv&""="" then%>
				            <td class="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					            <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				            </td>
			                <td class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
				                <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			                </td>
			            <%end if
			        elseif mode="pdf" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				        </td>
			        <%end if%>
		        </tr>
	        </table>
        </div>
    </div>                                                                                                   
	<input type="hidden" name="caju" value="<%=enc.EncodeForHtmlAttribute(null_s(caju))%>" />
	<table style="width:100%; height:42px; vertical-align:bottom;" align="center">
        <tr>
            <td style="width:100%; height:42px; vertical-align:bottom; text-align:center;">
                <%ImprimirPie_bt%>
            </td>
        </tr>
    </table>
</form>
</body>
</html>