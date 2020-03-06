<%@ Language=VBScript %>
<%
  ' JMG 26/03/2004 : Aceptación de parámetros en las consultas personalizadas.
  ' JMG 26/04/2004 : Muestra de parámetros personalizados en el listado.
  ' JMG 06/05/2004 : Permitir la gestión de las consultas personalizadas a través del sistema de gestión.
  ' JMG 19/05/2004 : Gestionar los parámetros personalizados.
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<!--#include file="../controlimpresion.inc" -->
<!--#include file="exportacion.inc" -->
<!--#include file="listado_personal_param.inc" -->
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=Session("caracteres")%>"/>
 <% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
/*
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

function printWindow() {
	factory.printing.header = "<%=LitTitulo%>. -- Registros : " + parent.pantalla.document.listado_personal_param.NumRegs.value + "&bPágina &p de &P"
	factory.printing.footer = "<%=PieListados%>&bFecha : &D"
	factory.printing.portrait = false
	factory.printing.leftMargin = 19.0
	factory.printing.topMargin = 19.0
	factory.printing.rightMargin = 19.0
	factory.printing.bottomMargin = 19.0
	factory.printing.Print(false, parent.pantalla)
}
*/
function Buscar(desdeGestion,ges) {
	parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&ges=" + ges + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
	parent.pantalla.document.listado_personal_param.submit();
	document.location="listado_personal_param_bt.asp?mode=search&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
}

function BuscarUsuario(mode) {
	parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=" + mode + "&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
	parent.pantalla.document.listado_personal_param.submit();
	document.location="listado_personal_param_bt.asp?mode=" + mode + "&desdeGestion=true" + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
}

function BuscarParametro() {
	parent.pantalla.fr_Tabla.document.listado_personal_param_det.action="listado_personal_param_det.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
	parent.pantalla.fr_Tabla.document.listado_personal_param_det.submit();
	document.location="listado_personal_param_bt.asp?mode=configParam&desdeGestion=true" + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
}

//function comprobar_enter(e, modo)
function comprobar_enter(modo)
{
    //var keycode = e.keyCode;
	//si se ha pulsado la tecla enter

	//if (keycode==13) {
		if ((modo.indexOf("asignUser")<0) && (modo.indexOf("asignSave")<0) && (modo.indexOf("configParam")<0)) {
			document.opciones.criterio.focus();
			parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=search&campo=" + document.opciones.campos.value +
			"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
			parent.pantalla.document.listado_personal_param.submit();
			document.location="listado_personal_param_bt.asp?mode=search" + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
		}
		else if (modo.indexOf("configParam")==0)
			BuscarParametro();
		else BuscarUsuario(modo);
	//}
}

// Comprueba si la cadena tiene parametros
function tieneParams(cadena) {
	expReg=new RegExp("arg[0-9]+");

	if (cadena.search(expReg)!=-1)  //Tiene parametros
		tiene=true;
	else tiene=false;

	return tiene;
}

//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if (parent.pantalla.document.listado_personal_param.consulta.value=="") {
		window.alert("<%=LitMsgConsultaNulo%>");
		return false;
	}
	if (parent.pantalla.document.listado_personal_param.descripcion.value=="") {
		window.alert("<%=LitMsgDescripcionNoNulo%>");
		return false;
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado,desdeGestion,ges) {
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "imprimir": //Imprimir Listado
					if (parent.pantalla.document.listado_personal_param.NumRegs) {
						parent.pantalla.focus();
						//printWindow();
                        parent.pantalla.print();
					}
					else alert("<%=LitMsgNoDevuelveRegistros%>");
					break;
				case "imprimirp": //Imprimir pdf
					if (parent.pantalla.document.listado_personal_param.NumRegs) {
						if (parseInt(parent.pantalla.document.listado_personal_param.NumRegs.value)>=parseInt(parent.pantalla.document.listado_personal_param.maxpdf.value))
							alert("<%=LitDemasiadosRegistros%>");
						else {
							parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
							parent.pantalla.document.listado_personal_param.action="listado_personal_param_pdf.asp";
							parent.pantalla.document.listado_personal_param.submit();

							document.location="listado_personal_param_bt.asp?mode=pdf&ges=" + ges +"&acc=<%= Request.QueryString("acc")%>" + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						}
					}
					else alert("<%=LitMsgNoDevuelveRegistros%>");
					break;
				case "cancelar": //Cancelar operacion
					if (parent.pantalla.document.listado_personal_param.ges.value.indexOf("SI")>=0) {
                        cadena="listado_personal_param.asp?fecha=" + parent.pantalla.document.listado_personal_param.fecha.value
                        + "&mode=add" 
                        + "&descripcion=" + parent.pantalla.document.listado_personal_param.descripcion.value
                        + "&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						parent.pantalla.document.listado_personal_param.action=cadena;
						parent.pantalla.document.listado_personal_param.submit();
						document.location="listado_personal_param_bt.asp?mode=add&ges=" + document.opciones.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					else {
                        cadena="listado_personal_param.asp?fecha=" + parent.pantalla.document.listado_personal_param.fecha.value
                        + "&mode=search"
                        //+ "&consulta=" + parent.pantalla.document.listado_personal_param.consulta.value 
                        + "&descripcion=" + parent.pantalla.document.listado_personal_param.descripcion.value 
                        + "&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						parent.pantalla.document.listado_personal_param.action=cadena
						parent.pantalla.document.listado_personal_param.submit();
						document.location="listado_personal_param_bt.asp?mode=search&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					break;
				case "add": //Añadir consulta
					parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=add&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					parent.pantalla.document.listado_personal_param.submit();
					document.location="listado_personal_param_bt.asp?mode=add&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					break;
				case "save": //Guardar consulta
					if (parent.pantalla.document.listado_personal_param.fecha.value=="") {
						parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=first_save&confirma=SI&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						parent.pantalla.document.listado_personal_param.submit();
						document.location="listado_personal_param_bt.asp?mode=browse&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					else {
						if (confirm("<%=LitMsgSobConsultaConfirm%>" + parent.pantalla.document.listado_personal_param.descripcion.value + ".?")) {
							parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=first_save&confirma=SI&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
							parent.pantalla.document.listado_personal_param.submit();
							document.location="listado_personal_param_bt.asp?mode=browse&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						}
					}
					break;
				case "exportar":
					if (parent.pantalla.document.listado_personal_param.NumRegs) {
						if (parseInt(parent.pantalla.document.listado_personal_param.NumRegs.value)>=parseInt(parent.pantalla.document.listado_personal_param.maxpdf.value))
							alert("<%=LitDemasiadosRegistrosExp%>");
						else
						{
						    cadena = parent.pantalla.document.listado_personal_param.descripcion.value.replace("á", "a");
						    cadena = cadena.replace("é", "e");
						    cadena = cadena.replace("í", "i");
						    cadena = cadena.replace("ó", "o");
						    cadena = cadena.replace("ú", "u");
						    cadena = cadena.replace("Á", "A");
						    cadena = cadena.replace("É", "E");
						    cadena = cadena.replace("Í", "I");
						    cadena = cadena.replace("Ó", "O");
						    parent.pantalla.document.listado_personal_param.descripcion.value = cadena.replace("Ú", "U");
							parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
                            parent.pantalla.exportar.document.listado_personal_param_exportar.params.value = parent.pantalla.document.listado_personal_param.params.value;
							parent.pantalla.exportar.document.listado_personal_param_exportar.fecha.value = parent.pantalla.document.listado_personal_param.fecha.value;
                            parent.pantalla.exportar.document.listado_personal_param_exportar.sesion_usuario.value = parent.pantalla.document.listado_personal_param.sesion_usuario.value;
                            parent.pantalla.exportar.document.listado_personal_param_exportar.ncliente.value = parent.pantalla.document.listado_personal_param.ncliente.value;
							parent.pantalla.exportar.document.listado_personal_param_exportar.descripcion.value = parent.pantalla.document.listado_personal_param.descripcion.value;
							parent.pantalla.exportar.document.listado_personal_param_exportar.action="listado_personal_param_exportar.asp?mode=exportar";
							parent.pantalla.exportar.document.listado_personal_param_exportar.submit();
						}
					}
					else alert("<%=LitMsgNoDevuelveRegistros%>");
					break;
			}
			break;

		case "browseConsulta":
			switch (pulsado) {
				case "add": //Nueva Consulta
					parent.pantalla.document.listado_personal_param.fecha.value="";
					parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=add&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					parent.pantalla.document.listado_personal_param.submit();
					document.location="listado_personal_param_bt.asp?mode=add&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					break;
				case "delete": //Borrar Consulta
					if (confirm("<%=LitMsgEliminarConsultaConfirm%>")) {
						if ((desdeGestion=="true") && (confirm("<%=LitMsgEliminarTodas%>"))) {
							parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=delete&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&cantidad=mas&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
							parent.pantalla.document.listado_personal_param.submit();
							document.location="listado_personal_param_bt.asp?mode=search&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						}
						else {
							parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=delete&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&cantidad=1&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
							parent.pantalla.document.listado_personal_param.submit();
							document.location="listado_personal_param_bt.asp?mode=search&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						}
					}
					break;
				// Modificado por jmg: Cambio de modo, de browse a getParams
				case "run": //Ejecutar Consulta
					if (!tieneParams(parent.pantalla.document.listado_personal_param.consulta.value.toLowerCase())) {    //parent.pantalla.document.listado_personal_param.consulta.value.toLowerCase().indexOf("arg")==-1) {
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
					    parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=browse&confirma=NO&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					    parent.pantalla.document.listado_personal_param.submit();
					    document.location="listado_personal_param_bt.asp?mode=browse&ges=" + document.opciones.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					else {
					    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility="visible";
					    parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=getParams&confirma=NO&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					    document.location="listado_personal_param_bt.asp?mode=getParams&ges=" + document.opciones.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					    parent.pantalla.document.listado_personal_param.submit();
					}
					break;
				case "save":
					if (ValidarCampos()) {
						if (parent.pantalla.document.listado_personal_param.consulta.value.indexOf(parent.pantalla.document.listado_personal_param.ncliente.value)>=0) {
							if (confirm("<%=LitMsgModificarTodas%>")) {
								parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=save&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&cantidad=mas&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
								parent.pantalla.document.listado_personal_param.submit();
								document.location="listado_personal_param_bt.asp?mode=save&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
							}
							else {
								parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=save&confirma=SI&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&cantidad=1&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
								parent.pantalla.document.listado_personal_param.submit();
								document.location="listado_personal_param_bt.asp?mode=save&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
							}
						}
						else alert("<%=LitMsgPropiaEmpresa%>");
					}
					break;
			}
			break;

		case "save":
			switch (pulsado) {
				case "imprimir": //Imprimir Listado
					parent.pantalla.focus();
					Imprimir();
					break;
				case "confirm": //Generar fichero
					if (confirm("<%=LitMsgRemesaConfirm%>")) {
						parent.pantalla.genera();
						parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=confirm&confirmar=SI&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						parent.pantalla.document.listado_personal_param.submit();
						document.location="listado_personal_param_bt.asp?mode=add&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					else {
						parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=confirm&confirmar=NO&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						parent.pantalla.document.listado_personal_param.submit();
						document.location="listado_personal_param_bt.asp?mode=add&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?fecha=" + parent.pantalla.document.listado_personal_param.fecha.value
					 + "&mode=add&consulta=" + parent.pantalla.document.listado_personal_param.consulta.value + "&descripcion=" + parent.pantalla.document.listado_personal_param.descripcion.value + "&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					parent.pantalla.document.listado_personal_param.submit();
					document.location="listado_personal_param_bt.asp?mode=add&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					break;
			}
			break;

		case "add":
			switch (pulsado) {
			    // Modificado por jmg: Cambio de modo, de browse a getParams
				case "run": //Ejecutar registro
					if (ValidarCampos()) {
						parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=getParams&confirma=NO&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						parent.pantalla.document.listado_personal_param.submit();
						document.location="listado_personal_param_bt.asp?mode=getParams&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					break;
				case "save": //Guardar registro
					if (ValidarCampos()) {
						if (parent.pantalla.document.listado_personal_param.consulta.value.indexOf(parent.pantalla.document.listado_personal_param.ncliente.value)>=0) {
							parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=save&confirma=SI&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
							parent.pantalla.document.listado_personal_param.submit();
							document.location="listado_personal_param_bt.asp?mode=save&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						}
						else alert("<%=LitMsgPropiaEmpresa%>");

					}
					break;
			}
			break;
		case "imp":
			switch (pulsado) {
				case "cancel": //Volver atrás
					parent.pantalla.document.location=history.back();
					history.back();
				break;
			}
			break;
		case "asignUser":
			switch (pulsado) {
				case "guardar": //Asigna la consulta
					if (parent.pantalla.document.listado_personal_param.consulta.value.indexOf(parent.pantalla.document.listado_personal_param.ncliente.value)>=0) {
						parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=save&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
						parent.pantalla.document.listado_personal_param.submit();
						document.location="listado_personal_param_bt.asp?mode=save&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					}
					else {
						alert("<%=LitMsgPropiaEmpresa%>");
						parent.close();
					}
				break;
				case "cancelar":
					parent.close();
				break;
			}
			break;
		case "asignSave":
			switch (pulsado) {
				case "guardar": //Asigna la consulta
					parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=save&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					parent.pantalla.document.listado_personal_param.submit();
					document.location="listado_personal_param_bt.asp?mode=save&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
				break;
				case "cancelar":
					parent.pantalla.document.listado_personal_param.action="listado_personal_param.asp?mode=add&confirma=SI&ges=" + parent.pantalla.document.listado_personal_param.ges.value + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
					parent.pantalla.document.listado_personal_param.submit();
					document.location="listado_personal_param_bt.asp?mode=add&ges=" + document.opciones.ges.value + "&desdeGestion=" + desdeGestion + "&sop=" + parent.pantalla.document.listado_personal_param.sop.value;
				break;
			}
			break;
		case "configParam":
			switch (pulsado) {
				case "cancelar":
					parent.close();
				break;
			}
		case "pdf":
			switch (pulsado) {
				case "volver":
				    parent.pantalla.document.location="listado_personal_param.asp?mode=search&ges=" + ges + "&sop=";// + parent.pantalla.document.listado_personal_param.sop.value;
					document.location="listado_personal_param_bt.asp?mode=search&ges=" + ges + "&sop=";// + parent.pantalla.document.listado_personal_param.sop.value;
				break;
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%''EscribeControlImpresion "listado_personal_param.asp"

mode=enc.EncodeForJavascript(Request.QueryString("mode"))
ges =limpiaCadena(Request.QueryString("ges"))
sop = limpiaCadena(Request.QueryString("ndoc"))
if sop & "" = "" then sop = limpiaCadena(Request.QueryString("sop"))
if sop & "" = "" then sop = Request.form("sop")
if ges="" then ges = request.form("ges")

accede= limpiaCadena(Request.QueryString("acc"))
if accede="" then
	accede="otro"
end if

if mode="gestion" or mode="asignSave" or mode="asignUser" or mode="configParam" then
	desdeGestion="true"
	if mode="gestion" then
		mode="search"
	end if
	ges="SI"
else
	desdeGestion=limpiaCadena(Request.QueryString("desdeGestion"))
end if

if desdeGestion="true" and mode="save" then
	mode="search"
end if

dim hil
obtenerParametros("ConsultaPer")
%>
<form name="opciones" method="post" action="">
<input type="hidden" name="ges" value="<%=enc.EncodeForHtmlAttribute(ges)%>"/>
<input type="hidden" name="sop" value="<%=enc.EncodeForHtmlAttribute(sop)%>"/>
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
<div id="PageFooter_ASP" >
<%if mode="imp" or mode="pdf" then%>
    <div id="ControlPanelFooter_ASP" >
<%else%>
    <div id="ControlPanelFooter_left_ASP" >
<%end if %>
    <table id="BUTTONS_CENTER_ASP">
		<tr>
		    <%if mode="browse" then%>
				<td id="idprint" class="CELDABOT" onclick="javascript:Accion('browse','imprimir','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				</td>
                <%if hil&""<>"1" then %>
				    <td id="idprintlist" class="CELDABOT" onclick="javascript:Accion('browse','imprimirp','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					    <%PintarBotonBTLeft LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				    </td>
				<%
                end if
                if ges="SI" then%>
				    <td id="idsave" class="CELDABOT" onclick="javascript:Accion('browse','save','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				<%end if%>

				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('browse','cancelar','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
				<td id="idexport" class="CELDABOT" onclick="javascript:Accion('browse','exportar','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
				</td>
			<%elseif mode="browseConsulta" then
				if ges="SI" then
					if desdeGestion<>"true" then%>
				        <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browseConsulta','add','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					        <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				        </td>
					<%else%>
				        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('browseConsulta','save','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					        <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				        </td>
					<%end if%>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browseConsulta','delete','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
    					<%PintarBotonBTLeft LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				<%end if
				if desdeGestion<>"true" then%>
				    <td id="idaccept" class="CELDABOT" onclick="javascript:Accion('browseConsulta','run','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
    					<%PintarBotonBTLeft LITBOTACEPTAR,ImgEjecutar,ParamImgEjecutar,LITBOTACEPTARTITLE%>
				    </td>
				<%end if
			elseif mode="search" then
				if (ges="SI" or desdeGestion="true") and sop <> "1" then%>
				    <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				<%end if
			elseif mode="save" then%>
				<td id="idprint" class="CELDABOT" onclick="javascript:Accion('save','imprimir','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('save','cancel','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="add" then
				if desdeGestion<>"true" then%>
				    <td id="idaccept" class="CELDABOT" onclick="javascript:Accion('add','run','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
    					<%PintarBotonBTLeft LITBOTACEPTAR,ImgEjecutar,ParamImgEjecutar,LITBOTACEPTARTITLE%>
				    </td>
				<%else%>
				    <td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				<%end if
			elseif mode="imp" then%>
				<td id="idreturn" class="CELDABOT" onclick="javascript:Accion('imp','cancel','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBT LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				</td>
			<%elseif mode="asignUser" then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('asignUser','guardar','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asignUser','cancelar','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="asignSave" then%>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('asignSave','guardar','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('asignSave','cancelar','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					<%PintarBotonBTLeft LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="configParam" then%>
				<td id="idaccept" class="CELDABOT" onclick="javascript:Accion('configParam','cancelar','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
    				<%PintarBotonBTLeft LITBOTACEPTAR,ImgEjecutar,ParamImgEjecutar,LITBOTACEPTARTITLE%>
				</td>
			<%elseif mode="pdf" then%>
			    <%if accede<>"cab" and accede<>"link" then%>
				    <td id="idreturn" class="CELDABOT" onclick="javascript:Accion('pdf','volver','<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');">
					    <%PintarBotonBT LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				    </td>
			    <%else%>
				    <td class="CELDABOT">
				    </td>
			    <%end if%>
			<%end if
%>
		</tr>
	</table>
    </div>
<%
			if mode<>"pdf" and accede<>"cab" and accede<>"link" then%>
                <div id="FILTERS_MASTER_ASP">
				<!--<td class=CELDABOT><%=LitBuscar & ": "%>-->
					<select class="IN_S" name="campos">
					<%if mode="asignUser" or mode="asignSave" then%>
						<option selected="selected" value="nombre"><%=LitNombre%></option>
				       	<option value="i.entrada"><%=LitLogin%></option>
       				<%elseif mode="configParam" then%>
      					<option selected="selected" value="titulo"><%=LitTitulo%></option>
      					<!--<option value="campotabla"><%=LitCampoTabla%></option>-->
                        <option value="nparam"><%=LitNParam%></option>
	          		<%else%>
      					<option selected="selected" value="descripcion"><%=LitDescripcion%></option>
      					<option value="consulta"><%=LitConsulta%></option>
	          		<%end if%>
	        		</select>
				<!--</td>
				<td class=CELDABOT>-->
					<select class="IN_S" name="criterio">
						<option value="contiene"><%=LitContiene%></option>
						<option value="empieza"><%=LitComienza%></option>
						<option value="termina"><%=LitTermina%></option>
						<option value="igual"><%=LitIgual%></option>
					</select>
				<!--</td>
				<td class=CELDABOT>-->
					<%if tipodato="" then%>
                        <input id="KeySearch" class="IN_S" type="text" name="texto" size="15" maxlength="20" value="" runat="javascript:comprobar_enter('<%=enc.EncodeForJavascript(mode)%>');"/>
					<%else%>
						<select class="IN_S" name="texto">
							<option value="0"><%=LitTexto%></option>
							<option value="1"><%=LitNumero%></option>
							<option value="2"><%=LitFecha%></option>
						</select>
					<%end if%>
				<!--</td>
				<td class=CELDABOT>-->
					<%if mode="asignUser" or mode="asignSave" then%>
						<a class="CELDAREF" href="javascript:BuscarUsuario('<%=enc.EncodeForJavascript(mode)%>');"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
					<%elseif mode="configParam" then%>
						<a class="CELDAREF" href="javascript:BuscarParametro();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
					<%else%>
                        <a class="CELDAREF" href="javascript:Buscar('<%=enc.EncodeForJavascript(desdeGestion)%>','<%=enc.EncodeForJavascript(ges)%>');"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
					<%end if%>
				<!--</td>
		</tr>
	</table>
                -->
                </div>
			<%end if%>
            </div>
<%if desdeGestion="true" then
    %>
    <table style="width:100%;height:42px;vertical-align:bottom;" align="center">
    <tr>
    <td style="width:100%;height:42px; vertical-align:bottom; text-align:center;">
    <%ImprimirPiePopUp_bt%>
    </td>
    </tr>
    </table>
    <%
else
    %>
    <table style="width:100%;height:42px;vertical-align:bottom;" align="center">
    <tr>
    <td style="width:100%;height:42px; vertical-align:bottom; text-align:center;">
    <%ImprimirPie_bt%>
    </td>
    </tr>
    </table>
    <%
end if%>
</form>
</body>
</html>