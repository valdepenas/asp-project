<%@ Language=VBScript %>
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloList%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../calculos.inc" -->

<!--#include file="cobros_param.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
//Validación de campos numéricos y fechas.
function ValidarCampos()
{
	if (!checkdate(parent.pantalla.document.listado_cobros_param.Dfecha))
	{
		window.alert("<%=LitMsgDesdeFechaFecha%>");
		return false;
	}
	if (!checkdate(parent.pantalla.document.listado_cobros_param.Hfecha))
	{
		window.alert("<%=LitMsgHastaFechaFecha%>");
		return false;
	}
	/*var diferencia = DiferenciaTiempo(parent.pantalla.document.listado_cobros_param.Hfecha.value,parent.pantalla.document.listado_cobros_param.Dfecha.value, "dias");
	if (diferencia>365)
	{
		window.alert("<%=LitDiferenciaFechas%>");
		return false;
	}*/
	parent.pantalla.document.listado_cobros_param.imptotalbmay.value=parent.pantalla.document.listado_cobros_param.imptotalbmay.value.replace(".",",");
	if (isNaN(parent.pantalla.document.listado_cobros_param.imptotalbmay.value.replace(",",".")))
	{
		window.alert("<%=LitImpTotAlbDebNum%>");
		return false;
	}
	return true;
}

function Buscar()
{
		parent.pantalla.document.location="listado_cobros_param.asp?mode=imp&campo=" + document.opciones.campos.value +
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=" +
		"&ncliente=" + parent.pantalla.document.listado_cobros_param.h_ncliente.value + "&viene=" + parent.pantalla.document.listado_cobros_param.viene.value +
		"&h_tabla=" + document.opciones.tabla.value;
		document.location="listado_cobros_param_bt.asp?viene=" + parent.pantalla.document.listado_cobros_param.viene.value + "&h_tabla=" + document.opciones.tabla.value;
}

function Ver(tipodoc)
{
	cadena="listado_cobros_param.asp?mode=imp&campo=" + document.opciones.campos.value;
	cadena +="&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value;
	cadena +="&sentido="
	cadena +="&ncliente=" + parent.pantalla.document.listado_cobros_param.h_ncliente.value;
	cadena +="&viene=" + parent.pantalla.document.listado_cobros_param.viene.value;

	if (tipodoc=="albaranes_cli") cadena += "&h_tabla=albaranes_cli";
	else
	{
		 if (tipodoc=="facturas_cli"){
		    if(parent.pantalla.document.listado_cobros_param.efectosPend.checked)
		        cadena += "&h_tabla=efectos_cli";
		    else
		        cadena += "&h_tabla=facturas_cli";
		 }
		 else cadena += "&h_tabla=tickets_cli";
	}
	parent.pantalla.document.location=cadena;
}

//FLM:20090818:si se ha seleccionado alguna serie se pone en h_tabla=efectos_cli.
function ComprobarEfectos(){
    var obj=parent.pantalla.document.listado_cobros_param.serie_efec;
    for(i=0;i<obj.length;i++){
        if(obj[i].selected==true){
            parent.pantalla.document.listado_cobros_param.h_tabla.value="efectos_cli";
            return;
        }            
    }
}
					    
//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado)
{
	switch (mode)
	{
		case "select1":
			switch (pulsado)
			{
				case "imp": //Aceptar
					if (ValidarCampos())
					{
					    //FLM:20090818:si se ha seleccionado alguna serie se pone en h_tabla=efectos_cli.
					    ComprobarEfectos();
						if (parent.pantalla.document.listado_cobros_param.poblacion.value!="" && parent.pantalla.document.listado_cobros_param.agrupar_poblacion.checked==true)
							parent.pantalla.document.listado_cobros_param.agrupar_poblacion.checked=false;

						parent.pantalla.document.listado_cobros_param.action="listado_cobros_paramResultado.asp?mode=" + pulsado;
						parent.pantalla.document.listado_cobros_param.submit();
						document.location="listado_cobros_param_bt.asp?mode=" + pulsado;
					}
					break;
				case "select1": //Cancelar
					parent.pantalla.document.listado_cobros_param.action="listado_cobros_param.asp?mode=" + pulsado;
					parent.pantalla.document.listado_cobros_param.submit();
					document.location="listado_cobros_param_bt.asp?mode=" + pulsado;
					break;
			}
			break;
		case "imp":
			switch (pulsado)
			{
				case "cancel": //Volver atrás
					//parent.pantalla.document.location=history.back();
					//history.back();
					parent.pantalla.document.location="listado_cobros_param.asp?mode=select1";
					document.location="listado_cobros_param_bt.asp?mode=select1";
					break;
				case "imprimir": //Volver atrás
					parent.pantalla.focus();
					// CCA 09-01-2008: Cambio para añadir el apaisado
					parent.pantalla.print(parent.pantalla.document.listado_cobros_paramResultado.apaisado.value,parent.pantalla.document.listado_cobros_paramResultado.NumRegs.value);
					break;
				case "imprimirp": //Imprimir Listado en PDF
					if (parseInt(parent.pantalla.document.listado_cobros_paramResultado.NumRegsTotal.value)>=parseInt(parent.pantalla.document.listado_cobros_paramResultado.maxpdf.value))
						alert("<%=LitMsgRegistros%>");
					else
					{
						parent.pantalla.document.listado_cobros_paramResultado.action="listado_cobros_param_pdf.asp?mode=browse&xls=0"; //&ncliente=" + parent.pantalla.document.listado_cobros_param.h_ncliente.value + "&nserie=" + parent.pantalla.document.listado_cobros_param.h_nserie.value + "&actividad=" + parent.pantalla.document.listado_cobros_param.h_actividad.value;
						parent.pantalla.document.listado_cobros_paramResultado.submit();
						document.location="listado_cobros_param_bt.asp?mode=pdf";
					}
					break;
			    case "exportar":
			        cadena = "";
			        cadena = cadena + "&dfecha=" + parent.pantalla.document.listado_cobros_paramResultado.dfecha.value;
			        cadena = cadena + "&hfecha=" + parent.pantalla.document.listado_cobros_paramResultado.hfecha.value;
			        cadena = cadena + "&h_ncliente=" + parent.pantalla.document.listado_cobros_paramResultado.h_ncliente.value;
			        cadena = cadena + "&actividad=" + parent.pantalla.document.listado_cobros_paramResultado.actividad.value;
			        if (parent.pantalla.document.listado_cobros_paramResultado.serie_alb != null)
			            cadena = cadena + "&serie_alb=" + parent.pantalla.document.listado_cobros_paramResultado.serie_alb.value;
			        if (parent.pantalla.document.listado_cobros_paramResultado.serie_fac != null)
			            cadena = cadena + "&serie_fac=" + parent.pantalla.document.listado_cobros_paramResultado.serie_fac.value;
			        if (parent.pantalla.document.listado_cobros_paramResultado.serie_tic != null)
			            cadena = cadena + "&serie_tic=" + parent.pantalla.document.listado_cobros_paramResultado.serie_tic.value;
			        if (parent.pantalla.document.listado_cobros_paramResultado.Documento != null)
			            cadena = cadena + "&Documento=" + parent.pantalla.document.listado_cobros_paramResultado.Documento.value;
			        cadena = cadena + "&h_tabla=" + parent.pantalla.document.listado_cobros_paramResultado.h_tabla.value;
			        cadena = cadena + "&h_que=" + parent.pantalla.document.listado_cobros_paramResultado.h_que.value;
			        cadena = cadena + "&comercial=" + parent.pantalla.document.listado_cobros_paramResultado.comercial.value;
			        cadena = cadena + "&agrupar_comercial=" + parent.pantalla.document.listado_cobros_paramResultado.agrupar_comercial.value;
			        cadena = cadena + "&poblacion=" + parent.pantalla.document.listado_cobros_paramResultado.poblacion.value;
			        cadena = cadena + "&agrupar_poblacion=" + parent.pantalla.document.listado_cobros_paramResultado.agrupar_poblacion.value;
			        cadena = cadena + "&imptotalbmay=" + parent.pantalla.document.listado_cobros_paramResultado.imptotalbmay.value;
			        if (parent.pantalla.document.listado_cobros_paramResultado.opcclientebaja != null)
			            cadena = cadena + "&opcclientebaja=" + parent.pantalla.document.listado_cobros_paramResultado.opcclientebaja.value;
			        if (parent.pantalla.document.listado_cobros_paramResultado.nserieEfec != null)
			            cadena = cadena + "&nserieEfec=" + parent.pantalla.document.listado_cobros_paramResultado.nserieEfec.value;
                    parent.pantalla.marcoExportar.document.location = "listado_cobros_param_pdf.asp?mode=browse&xls=1" + cadena;
			        break;
				}
			break;
		case "pdf":
			switch (pulsado)
			{
				case "back": //Volver a la pantalla anterior
				            //FLM:20090916:arreglo botón volver.
							//parent.pantalla.document.location="listado_cobros_param.asp?mode=select1";
							parent.pantalla.location.href="listado_cobros_param.asp?mode=select1";
							document.location="listado_cobros_param_bt.asp?mode=select1";
							break;
			}
			break;
		}
}

function cambiar()
{
	if (document.opciones.tabla.value=="albaranes_cli")
	{
		document.opciones.campos.item(0).value="nalbaran";
		document.opciones.campos.item(0).text="<%=LitAlbVen%>";
	}
	else
	{
		document.opciones.campos.item(0).value="nfactura";
		document.opciones.campos.item(0).text="<%=LitFactura%>";
	}
	Ver(document.opciones.tabla.value);
}

function comprobar_enter() {
    document.opciones.criterio.focus();
    Buscar();
}
</script>
<%if request.querystring("viene")>"" then
	viene=limpiaCadena(request.querystring("viene"))
elseif request.form("viene")>"" then
	viene=request.form("viene")
end if
if viene="tienda" then
	cadena="topmargin='0'"
else
	cadena=""
	'if nserieEfec&""<>"" then
end if%>
<body class="body_master_ASP">
<%mode=enc.EncodeForJavascript(Request.QueryString("mode"))

if request.querystring("tabla")>"" then
	tabla=limpiaCadena(request.querystring("tabla"))
elseif request.form("tabla")>"" then
	tabla=request.form("tabla")
end if
if tabla="" then
	if request.querystring("h_tabla")>"" then
		tabla=limpiaCadena(request.querystring("h_tabla"))
	elseif request.form("h_tabla")>"" then
		tabla=request.form("h_tabla")
	end if
end if

''ricardo 3-12-2007 si viene de la tienda , por defecto veremos las facturas
if viene="tienda" and tabla="" then tabla="facturas_cli"%>
<form name="opciones" method="post" action="javascript:if ('<%=viene%>'=='tienda') {document.opciones.criterio.focus();Buscar()}">
    <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
	        <%if viene="tienda" then%>
		        <div id="FILTERS_MASTER_ASP">
					<font class="CELDABOT" color="<%=color_blanco_det%>"><%=LitVer%></font>
					<select class="IN_S" name="tabla" onchange="cambiar()">
						<option <%if tabla="albaranes_cli" then response.write("selected")%> value="albaranes_cli"><%=LitAlbaranes%></option>
						<option <%if tabla="facturas_cli" then response.write("selected")%> value="facturas_cli"><%=LitFacturas%></option>
					</select>
					<select class="IN_S" name="campos">
						<%if tabla="albaranes_cli" then%>
							<option selected value="nalbaran"><%=LitAlbVen%></option>
						<%else%>
							<option selected value="nfactura"><%=LitFactura%></option>
						<%end if%>
						<option value="fecha"><%=LitFecha%></option>
					</select>

					<select class="IN_S" name="criterio">
						<option value="contiene"><%=LitContiene%></option>
						<option value="termina"><%=LitTermina%></option>
						<option value="igual"><%=LitIgual%></option>
					</select>

					<input class="IN_S" type="text" name="texto" size=20 maxlength=20 value="" runat="javascript:comprobar_enter();">
					<a class='CELDAREF' href="javascript:Buscar();"><img src="../images/<%=ImgBuscar_bt%>" <%=ParamImgBuscar_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
                </div>
	        <%else%>
		        <table id="BUTTONS_CENTER_ASP">
			        <tr>
				        <%if mode="select1" then%>
				            <td class="CELDABOT" onclick="javascript:Accion('select1','imp');">
					            <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				            </td>
				            <td class="CELDABOT" onclick="javascript:Accion('select1','select1');">
					            <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				            </td>
				        <%elseif mode="imp" then%>
				            <td class="CELDABOT" onclick="javascript:Accion('imp','imprimir');">
					            <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				            </td>
				            <td class="CELDABOT" onclick="javascript:Accion('imp','imprimirp');">
					            <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				            </td>
                            <td class="CELDABOT" onclick="javascript:Accion('imp','exportar');">
				                <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			                </td>
				            <td class="CELDABOT" onclick="javascript:Accion('imp','cancel');">
					            <%PintarBotonBTRed LITBOTCANCELAR,ImgVolver,ParamImgVolver,LITBOTCANCELARTITLE%>
				            </td>
				        <%elseif mode="pdf" then%>
				            <td class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				            </td>
				        <%end if%>
			        </tr>
		        </table>
	        <%end if%>
        </div>
    </div>
    <table style="width:100%;height:30px;vertical-align:bottom;" align="center">
        <tr>
            <td style="width:100%;height:30px; vertical-align:bottom; text-align:center;">
                <%ImprimirPie_bt%>
            </td>
        </tr>
    </table>
</form>
</body>
</html>
