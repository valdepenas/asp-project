<%@ Language=VBScript %>
<%
'' JCI 30/04/2003 : Soluci�n de problemas y errores varios en la conversi�n de pedidos con n�meros de serie
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">

</head>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  


<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->

<!--#include file="../ilion.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../varios.inc" -->

<!--#include file="pedpro_albpro_param.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
function cambiarfecha(fecha,modo)
{
	var fecha_ar=new Array();

	if (fecha!="")
	{
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length)
		{
			if (fecha.substring(l,l+1)=='/')
			{
				suma++;
				fecha_ar[suma]="";
			}
			else
			{
				if (fecha.substring(l,l+1)!='') fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
			}
			l++;
		}
		if (suma!=2)
		{
			window.alert("<%=LitFechaMal%> en el campo " + modo );
			return false;
		}
		else
		{
			nonumero=0;
			while (suma>=0 && nonumero==0)
			{
				if (isNaN(fecha_ar[suma])) nonumero=1;
				if (fecha_ar[suma].length>2 && suma!=2) nonumero=1;
				if (fecha_ar[suma].length>4 && suma==2) nonumero=1;
				suma--;
			}

			if (nonumero==1)
			{
				window.alert("<%=LitFechaMal%> en el campo " + modo);
				return false;
			}
		}
	}
	return true;
}

//Validaci�n de campos num�ricos y fechas.
function ValidarCampos() {
	if (parent.pantalla.document.pedpro_albpro_param.fdesde.value=="") {
		window.alert("<%=LitMsgDesdeFechaNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.pedpro_albpro_param.nproveedor.value=="") {
		window.alert("<%=LitMsgProveedorNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.pedpro_albpro_param.fhasta.value=="") {
		window.alert("<%=LitMsgHastaFechaNoNulo%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.pedpro_albpro_param.fdesde)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return;
	}
	if (!cambiarfecha(parent.pantalla.document.pedpro_albpro_param.fdesde.value,"Desde Fecha")) return false;
	if (!checkdate(parent.pantalla.document.pedpro_albpro_param.fhasta)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return;
	}
	if (!cambiarfecha(parent.pantalla.document.pedpro_albpro_param.fhasta.value,"Hasta Fecha")) return false;

	return true;
}

//MPC 11/09/2008 Si existe el albaran se lanza el procedimiento de actualizaci�n
/*
function existeAlbaranPro()
{
	// Use the native cross-browser nitobi Ajax object
	var myAjaxRequest = new nitobi.ajax.HttpRequest();

	// Define the url for your generatekey script
	myAjaxRequest.handler = "existeAlbaranPro.asp?nalbaran=" + parent.pantalla.document.pedpro_albpro_param.nalbaran_pro.value + 
	"&fecha=" + parent.pantalla.document.pedpro_albpro_param.falbaran.value+
	"&nproveedor=" + parent.pantalla.document.pedpro_albpro_param.nproveedor2.value;
	myAjaxRequest.async = false;
	myAjaxRequest.get();

	// return the result to the grid
	return myAjaxRequest.httpObj.responseText;
}
*/
var xmlHttp, ServerResponse = null;
function existeAlbaranPro()
{
    xmlHttp = GetXmlHttpObject();
    if (xmlHttp != null) 
    {
        var url = "existeAlbaranPro.asp?nalbaran=" + parent.pantalla.document.pedpro_albpro_param.nalbaran_pro.value + 
        "&fecha=" + parent.pantalla.document.pedpro_albpro_param.falbaran.value+
        "&nproveedor=" + parent.pantalla.document.pedpro_albpro_param.nproveedor2.value;
        //window.aler(url);
        xmlHttp.open("GET",url,false);  //synchronous method
        xmlHttp.send(null);
        return xmlHttp.responseText; //synchronous method
    }
}
function getData ()
{ 
    if (xmlHttp.readyState == 4 || xmlHttp.readyState == "complete")
    { 
        ServerResponse = xmlHttp.responseText;
    } 
 } 
   
function GetXmlHttpObject()
{ 
   var xmlHttp=null; 
   try
   { 
        // Firefox, Opera 8.0+, Safari 
        xmlHttp=new XMLHttpRequest(); 
   }
   catch (e)
   { 
        //Internet Explorer 
       try
       { 
           xmlHttp=new ActiveXObject("Msxml2.XMLHTTP"); 
       }
       catch (e)
       { 
           xmlHttp=new ActiveXObject("Microsoft.XMLHTTP"); 
       } 
   } 
   return xmlHttp; 
 } 

function ValidarCampos2() {
//ricardo 10-3-2003
//se comprueba si se han insertado todos los nserie que hacen falta
	si_obligado_nserie=parent.pantalla.document.pedpro_albpro_param.si_obligado_nserie.value;
	if (si_obligado_nserie==1){
		nregs=parent.pantalla.document.pedpro_albpro_param.h_nfilas.value;
		hay_poner_serie=0;
		tiene_nserie=0;
		pedidosKO="";
		for (lk=1;lk<=nregs;lk++){
			if (eval("parent.pantalla.document.pedpro_albpro_param.check" + lk + ".checked==true")){
				if (eval("parent.pantalla.document.pedpro_albpro_param.hace_falta_nserie" + lk + ".value=='1'")){
					hay_poner_serie=hay_poner_serie+1;
					if (eval("parent.pantalla.document.pedpro_albpro_param.nserie" + lk + ".value=='OK'"))
						tiene_nserie=tiene_nserie+1;
					else
						eval("pedidosKO=pedidosKO + " + "\n" + "parent.pantalla.document.pedpro_albpro_param.check" + lk + ".value;");
				}
			}
		}
		if (hay_poner_serie!=tiene_nserie){
			window.alert("<%=litMsgFaltanNserie%>" + trimCodEmpresa(pedidosKO));
			return false;
		}
	}
//////////
	nregistros=parent.pantalla.document.pedpro_albpro_param.h_nfilas.value;
	seleccionado=false
	for (i=1;i<=nregistros;i++) {
		nombre="check" + i;
		if(parent.pantalla.document.pedpro_albpro_param.elements[nombre].checked==true) seleccionado=true;
	}
	if (seleccionado==false)
	{
		window.alert("<%=LitMsgSeleccionNula%>");
		return false;
	}
	if (parent.pantalla.document.pedpro_albpro_param.nserie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.pedpro_albpro_param.falbaran.value=="") {
		window.alert("<%=LitMsgFechaNoNulo%>");
		return false;
	}

	if (!checkdate(parent.pantalla.document.pedpro_albpro_param.falbaran)) {
		window.alert("<%=LitMsgFechaFecha%>");
		return false;
    }

	if (!window.confirm("<%=LitMsgConvPedidosConfirm%>")) return false;
	if (!cambiarfecha(parent.pantalla.document.pedpro_albpro_param.falbaran.value,"Fecha Albaran")) return false;
	
	//MPC 11/09/2008 Si existe el albaran se lanza el procedimiento de actualizaci�n
	existe = 0;
	if (existeAlbaranPro() != "") {
	    if (!confirm("<%=LitExisteAlbaran%>")) {
	        return false;
	    }
	    else {
	        existe = 1;
	    }
	}

    return true;

}

//Realizar la acci�n correspondiente al bot�n pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "select1":
			switch (pulsado) {
				case "select2": //Aceptar
					if (ValidarCampos()) {
						parent.pantalla.document.pedpro_albpro_param.action="pedpro_albpro_paramResultado.asp?mode=" + pulsado;
						parent.pantalla.document.pedpro_albpro_param.submit();
						document.location="pedpro_albpro_param_bt.asp?mode=" + pulsado;
					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.pedpro_albpro_param.action="pedpro_albpro_param.asp?mode=" + pulsado;
					parent.pantalla.document.pedpro_albpro_param.submit();
					document.location="pedpro_albpro_param_bt.asp?mode=" + pulsado;
					break;
			}
			break;
		case "select2":
			switch (pulsado) {
				case "todos": //Seleccionar todos los registros
					nregistros=parent.pantalla.document.pedpro_albpro_param.h_nfilas.value;
					for (i=1;i<=nregistros;i++) {
						nombre="check" + i;
						parent.pantalla.document.pedpro_albpro_param.elements[nombre].checked=true;
					}
					break;

				case "ninguno": //No seleccionar ningun registro
					nregistros=parent.pantalla.document.pedpro_albpro_param.h_nfilas.value;
					for (i=1;i<=nregistros;i++) {
						nombre="check" + i;
						parent.pantalla.document.pedpro_albpro_param.elements[nombre].checked=false;
					}
					break;
				case "confirm": //Aceptar
					if (ValidarCampos2()) {
						genpedbis=0;
						parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
						parent.pantalla.document.pedpro_albpro_param.action="pedpro_albpro_param.asp?mode=confirm&genpedbis=" + genpedbis+"&existe="+existe;
						parent.pantalla.document.pedpro_albpro_param.submit();
						document.location="pedpro_albpro_param_bt.asp?mode=select1";
					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.pedpro_albpro_param.action="pedpro_albpro_param.asp?mode=select1";
					parent.pantalla.document.pedpro_albpro_param.submit();
					document.location="pedpro_albpro_param_bt.asp?mode=select1";
					break;
			}
			break;

		case "imp":
			switch (pulsado) {
				case "cancel": //Volver atr�s
					parent.pantalla.document.location=history.back();
					history.back();
					break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_ASP" >
    <table id="BUTTONS_CENTER_ASP">
        <tr><%
			if mode="select1" then
				%>
				<td id="idaccept" class="CELDABOT" onclick="javascript:Accion('select1','select2');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('select1','select1');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
				<%
			elseif mode="select2" then
				%>
				<td id="idSelectAll" class="CELDABOT" onclick="javascript:Accion('select2','todos');">
					<%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,LITBOTSELTODOTITLE%>
				</td>
				<td id="idSelectNothing" class="CELDABOT" onclick="javascript:Accion('select2','ninguno');">
					<%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,LITBOTDSELTODOTITLE%>
				</td>
				<td id="idaccept" class="CELDABOT" onclick="javascript:Accion('select2','confirm');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				</td>

				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('select2','select1');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
				<%
			elseif mode="imp" then
				%>
				<td id="idreturn" class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
					<%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				</td>
				<%
			end if%>
		</tr>
	</table>
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