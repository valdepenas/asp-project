<%@ Language=VBScript %>
<%
''IML 23/06/2003: Migración a monobase
%>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>

<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="Ahoja_gastos.inc" -->

<!--#include file="../styles/Master.css.inc" -->
<!--#include file="../styles/FootButton.css.inc" -->

<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<script language="javascript" type="text/javascript">

    window.addEventListener("load", function hideFilterMaster() {
        var classNameLeft = parent.parent.document.getElementById('left').className;
        if (classNameLeft == "search-panel open") {
            document.getElementById("FILTERS_MASTER_ASP").setAttribute("style", "display:none")
        }
    });

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
			window.alert("<%=LitFechaMal%>");
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

function Buscar()
{
    SearchPage("traspasos_caja_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value, 1);
    document.opciones.texto.value = "";
}

//Validación de campos numéricos y fechas.
function ValidarCampos() {
	if (parent.pantalla.document.traspasos_caja.serie.value=="") {
		window.alert("<%=LitMsgSerieNoNulo%>");
		return false;
	}

	if (!cambiarfecha(parent.pantalla.document.traspasos_caja.fecha.value,"FECHA TRASPASO")){
		return false;
	}

	if(!checkdate(parent.pantalla.document.traspasos_caja.fecha)){
		window.alert("<%=LitMsgFechaFecha%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.descripcion.value=="") {
		window.alert("<%=LitMsgDescripcionNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.responsable.value=="") {
		window.alert("<%=LitMsgResponsableNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.cajaorg.value=="") {
		window.alert("<%=LitMsgcajaorgNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.cajadest.value=="") {
		window.alert("<%=LitMsgcajadestNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.cajaorg.value==parent.pantalla.document.traspasos_caja.cajadest.value){
		window.alert("<%=LitMsgCajasIguales%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.divisa.value=="") {
		window.alert("<%=LitMsgDivisaNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.medio.value=="") {
		window.alert("<%=LitMsgMedioNoNulo%>");
		return false;
	}

	if (parent.pantalla.document.traspasos_caja.importeEsc.value!="SI") {
		if (isNaN(parent.pantalla.document.traspasos_caja.importe.value.replace(",","."))) {
			window.alert("<%=LitMsgImporteNumerico%>");
			return false;
		}

		if (parent.pantalla.document.traspasos_caja.importe.value.replace(",",".")<0) {
			window.alert("<%=LitMsgImportePositivo%>");
			return false;
		}

		if (parent.pantalla.document.traspasos_caja.importe.value=="") {
			window.alert("<%=LitMsgImporteNoNulo%>");
			return false;
		}
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "browse":
			switch (pulsado) {
			    case "add": //Nuevo registro
					parent.pantalla.document.location="traspasos_caja.asp?mode=" + pulsado;
					document.location="traspasos_caja_bt.asp?mode=" + pulsado;
					break;

				case "edit": //Editar registro
					if (parent.pantalla.document.traspasos_caja.encierre.value=="") {
						parent.pantalla.document.location="traspasos_caja.asp?ntraspaso=" + parent.pantalla.document.traspasos_caja.ntraspaso.value + "&mode=" + pulsado;
						document.location="traspasos_caja_bt.asp?mode=" + pulsado;
					}
					else alert("<%=LitMsgEditarTraspasoCierre%> " + parent.pantalla.document.traspasos_caja.encierre.value);
					break;

				case "delete": //Eliminar registro
					if (parent.pantalla.document.traspasos_caja.encierre.value=="") {
						if (parent.pantalla.document.traspasos_caja.contab.value=="True") window.alert("<%=LitMsgTraspasoContabilizado%>");
						else {
							if (window.confirm("<%=LitMsgEliminarTraspasoConfirm%>")==true) {
								parent.pantalla.document.traspasos_caja.action="traspasos_caja.asp?mode=" + pulsado + "&ntraspaso=" + parent.pantalla.document.traspasos_caja.ntraspaso.value;
								parent.pantalla.document.traspasos_caja.submit();
								document.location="traspasos_caja_bt.asp?mode=search";
							}
						}
					}
					else alert("<%=LitMsgEditarTraspasoCierre%> " + parent.pantalla.document.traspasos_caja.encierre.value);
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos()) {
						parent.pantalla.document.traspasos_caja.action="traspasos_caja.asp?ntraspaso=" + parent.pantalla.document.traspasos_caja.ntraspaso.value + "&mode=save";
						parent.pantalla.document.traspasos_caja.submit();
						document.location="traspasos_caja_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.location="traspasos_caja.asp?ntraspaso=" + parent.pantalla.document.traspasos_caja.ntraspaso.value + "&mode=browse";
					document.location="traspasos_caja_bt.asp?mode=browse";
					break;

				case "search": //Buscar datos
					break;
			}
			break;

case "add":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos()) {
						parent.pantalla.document.traspasos_caja.action="traspasos_caja.asp?ntraspaso=" + parent.pantalla.document.traspasos_caja.ntraspaso.value + "&mode=save";
						parent.pantalla.document.traspasos_caja.submit();
						document.location="traspasos_caja_bt.asp?mode=browse";
					}
					break;

				case "cancel": //Cancelar edición
				    parent.pantalla.document.traspasos_caja.action = "traspasos_caja.asp?mode=add";
					parent.pantalla.document.traspasos_caja.submit();
					document.location = "traspasos_caja_bt.asp?mode=add";
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "search":
			switch (pulsado) {
				case "search": //Buscar datos
					break;
			}
			break;
	}
}

//****************************************************************************************************************
function comprobar_enter()
{
	//document.opciones.criterio.focus();
	Buscar();
}
</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")
%>
<form name="opciones" method="post">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>">
<div id="PageFooter_ASP" >
    <div id="ControlPanelFooter_left_ASP" >
        <table id="BUTTONS_CENTER_ASP">
		<tr>
		    <%if mode="browse" then%>
                   <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBTLeft LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				    </td>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeftRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
			<%elseif mode="search" then%>
                   <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
			<%elseif mode="edit" then%>
			        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
			<%elseif mode="add" then%>
			        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					    <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
			<%end if%>
           </tr>
        </table>
    </div>

    <div id="FILTERS_MASTER_ASP">
		<select class="IN_S" name="campos">
		  <option value="t.ntraspaso"><%=LitNTraspaso%></option>
          <option value="t.descripcion"><%=LitDescripcion%></option>
          <option value="t.responsable"><%=LitResponsable%></option>
          <option value="t.cajaorg"><%=LitCajaOrg%></option>
          <option value="t.cajadest"><%=LitCajaDes%></option>
        </select>
				<select class="IN_S" name="criterio">
					<option value="contiene"><%=LitContiene%></option>
					<option value="termina"><%=LitTermina%></option>
					<option value="igual"><%=LitIgual%></option>
				</select>
				<input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
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