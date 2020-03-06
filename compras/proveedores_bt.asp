<%@ Language=VBScript %>
<%
dim enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
%>
<%
' RGU 8/10/2007: Si el parametro pagsl=1 no se pueden editar los datos del proveedor
 %>
<script id=DebugDirectives runat=server language=javascript>
// Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->

<!--#include file="../ilion.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="../varios.inc" -->
<!--#include file="../calculos.inc" -->

<!--#include file="proveedores.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<%'' MPC 08/10/208 Se obtiene el parámetro cifrepe para controlar si se puede insertar cif repetidos
dim noadd, pagsl, cifrepe
mode=Request.QueryString("mode")
obtenerparametros("proveedores")
noadd = noadd&""%>

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/iban.js"></script>
<script language="javascript" type="text/javascript">

    window.addEventListener("load", function hideFilterMaster() {
        var classNameLeft = parent.parent.document.getElementById('left').className;
        if (classNameLeft == "search-panel open") {
            document.getElementById("FILTERS_MASTER_ASP").setAttribute("style", "display:none")
        }
    });

function comprobar_enter(){
	//si se ha pulsado la tecla enter
	//if (window.event.keyCode==13){
		//document.opciones.criterio.focus();
		Buscar();
	//}
}

var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1);
function Imprimir() {
	if (pr) //NS4, IE5
		parent.pantalla.print()
	else if (da && !mac) // IE4 (Windows)
		alert("<%=LitNoImprime%>");
	else // Otros Navegadores
		alert("<%=LitNoImprime%>");
	return false;
}

//DGB: change to page search
function Buscar() {
	SearchPage("proveedores_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value,1);

    document.opciones.texto.value="";
}

// 01/10/2012 DGB: eliminar nitobi 
var xmlHttp, ServerResponse = null;
function ComprobarExisteCIF()
{
    xmlHttp = GetXmlHttpObject();
    if (xmlHttp != null) 
    {
        var nproveedor="";
        try{
            nproveedor=parent.pantalla.document.proveedores.hnproveedor.value;
        }
        catch(e)
        {
        }
        var url = "existeCIFProveedor.asp?cif=" + parent.pantalla.document.proveedores.cif.value + "&nproveedor=" + nproveedor;
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


//Validación de campos numéricos y fechas.
function ValidarCampos(mode) {
	if (parent.pantalla.document.proveedores.falta.value==""){
		window.alert("<%=LitMsgFaltaNoNulo%>");
		return false;
	}
	else{
		if (!checkdate(parent.pantalla.document.proveedores.falta)){
			window.alert("<%=LitMsgFechaAltaFecha%>");
			return false;
		}
	}
	if (parent.pantalla.document.proveedores.fbaja.value!=""){
		if (!checkdate(parent.pantalla.document.proveedores.fbaja)){
			window.alert("<%=LitMsgFechaBajaFecha%>");
			return false;
		}
	}
	if (parent.pantalla.document.proveedores.razon_social.value=="") {
		window.alert("<%=LitMsgRsocialNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.proveedores.nombre.value=="") {
		window.alert("<%=LitMsgNombreNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.proveedores.domicilio.value=="") {
		window.alert("<%=LitMsgDireccionNoNulo%>");
		return false;
	}
	if (parent.pantalla.document.proveedores.cif.value=="") {
		window.alert("<%=LitMsgCifNoNulo%>");
		return false;
	}

	if (isNaN(parent.pantalla.document.proveedores.e_primer_ven.value))  {
		window.alert("<%=LitPrimerVenNoNum%>");
		parent.pantalla.document.proveedores.e_primer_ven.focus();
		return false;		
	}
	else if (isNaN(parent.pantalla.document.proveedores.e_segundo_ven.value))  {
		window.alert("<%=LitSegundoVenNoNum%>");
		parent.pantalla.document.proveedores.e_segundo_ven.focus();
		return false;		
	}
	else if (isNaN(parent.pantalla.document.proveedores.e_tercer_ven.value))  {
		window.alert("<%=LitTercerVenNoNum%>");
		parent.pantalla.document.proveedores.e_tercer_ven.focus();
		return false;
    }
    var country = parent.pantalla.document.proveedores.country.value;
    if (country == "") 
        country ="ES";

    if (country == "ES"){
        //var account = parent.pantalla.document.proveedores.NEntidad.value + parent.pantalla.document.proveedores.Oficina.value + parent.pantalla.document.proveedores.DC.value + parent.pantalla.document.proveedores.Cuenta.value;
        var account = parent.pantalla.document.proveedores.ncuenta.value;
        var iban = CreateIBAN(country, account);
        
            
	    if (parent.pantalla.document.proveedores.iban.value!="") {
            var i_iban = parent.pantalla.document.proveedores.iban.value;

            if (!ValidateIBAN(iban, i_iban)) {
                alert("<%=LitCodeIBANIncorrect%>");
                return false;
            }
        }
        else 
        {
            if (isNaN(iban) == true)
                parent.pantalla.document.proveedores.iban.value = "";
            else
                parent.pantalla.document.proveedores.iban.value = iban;
        }
    }
	if ((mode=="add"||mode=="edit")&&(parent.pantalla.document.proveedores.si_campo_personalizables.value==1)){
	    //FLM:130309:Verificamos la cuenta de abono //si está marcada la domiciliación.
	    //if(parent.pantalla.document.proveedores.Domiciliacion.checked){
	      //  alert(parent.pantalla.document.proveedores.ncuenta.value);
	    //}
		num_campos=parent.pantalla.document.proveedores.num_campos.value;

		respuesta=comprobarCampPerso("parent.pantalla.",num_campos,"proveedores");
		if(respuesta!=0){
			titulo="titulo_campo" + respuesta;
			tipo="tipo_campo" + respuesta;
			titulo=parent.pantalla.document.proveedores.elements[titulo].value;
			tipo=parent.pantalla.document.proveedores.elements[tipo].value;
			if (tipo==4) {
				nomTipo="<%=LitTipoNumericoPro%>";
			}
			else if (tipo==5) {
				nomTipo="<%=LitTipoFechaPro%>";
			}

			window.alert("<%=LitMsgCampoPro%> " + titulo + " <%=LitMsgTipoPro%> " + nomTipo);

			return false;
		}
	}
	return true;
}
    <%  
        viene2 = EncodeForHtml(limpiaCadena(Request.QueryString("modp") & ""))
        if viene2 = "" then viene = EncodeForHtml(limpiaCadena(Request.Form("modp") & ""))
    %>

// ASP 09/1/2012
 function reloadPanelGlobal(viene)
 {
    if(viene == "GlobalAgenda")
    {
        parent.window.opener.__doPostBack("reload",""); 
    }
 }
 //FIN ASP 09/1/2012

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Nuevo registro
					parent.pantalla.document.proveedores.action="proveedores.asp?mode=" + pulsado;
					parent.pantalla.document.proveedores.submit();
					document.location="proveedores_bt.asp?mode=" + pulsado + "&modp=<%=viene2%>";
					break;

				case "edit": //Editar registro
					parent.pantalla.document.proveedores.action="proveedores.asp?nproveedor=" + parent.pantalla.document.proveedores.hnproveedor.value +
					"&mode=" + pulsado;
					parent.pantalla.document.proveedores.submit();
					document.location="proveedores_bt.asp?mode=" + pulsado+"&modp=<%=viene2%>";
					break;

				case "delete": //Eliminar registro
					if (window.confirm("<%=LitMsgEliminarProveedorConfirm%>")==true) {
						parent.pantalla.document.proveedores.action="proveedores.asp?mode=" + pulsado + "&nproveedor=" + parent.pantalla.document.proveedores.hnproveedor.value;
						parent.pantalla.document.proveedores.submit();
                        reloadPanelGlobal("<%=viene2%>");
						document.location="proveedores_bt.asp?mode=browse&modp=<%=viene2%>";
					}
					break;
				case "print": //Imprimir ficha
					parent.pantalla.focus();
					Imprimir();
					break;
				case "search": //Buscar datos
					break;
			}
			break;

		case "edit":
			switch (pulsado) {
				case "save": //Guardar registro
					if (ValidarCampos(mode))
					{
			            var cade="save&nproveedor=" + parent.pantalla.document.proveedores.hnproveedor.value;
					    repe =0;
                        <%if cifrepe = "1" then%>
                            existe =ComprobarExisteCIF();
                            if (existe!="")
                            {
                                if (confirm("<%=LitCIFRepe%>")) repe = 1;
                            }
                        <%end if%>
			            
			            <%if cifrepe = "1" then%>
                        //window.alert(repe + "-" + existe);
			            if (repe == 1 ||existe == "")
			            {
                            parent.pantalla.document.proveedores.action="proveedores.asp?mode=" + cade + "&repe=" + repe;
			                parent.pantalla.document.proveedores.submit();
                            //ASP 09/01/2012                            
                             reloadPanelGlobal("<%=viene2%>");
                            //FIN ASP 09/01/2012
			                document.location="proveedores_bt.asp?mode=browse";
			            }
			            <%else%>
                            parent.pantalla.document.proveedores.action="proveedores.asp?mode=" + cade + "&repe=" + repe;
			                parent.pantalla.document.proveedores.submit();
			                document.location="proveedores_bt.asp?mode=browse";
			            <%end if%>

					}
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.proveedores.action="proveedores.asp?nproveedor=" + parent.pantalla.document.proveedores.hnproveedor.value +
					"&mode=browse";
					parent.pantalla.document.proveedores.submit();
					document.location="proveedores_bt.asp?mode=browse";
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "add":
			switch (pulsado)
			{
				case "save": //Guardar registro
					if (ValidarCampos(mode))
					{
                        var cade="save";
					    repe =0;
                        <%if cifrepe = "1" then%>
                            existe =ComprobarExisteCIF();
                            if (existe!="")
                            {
                                if (confirm("<%=LitCIFRepe%>")) repe = 1;
                            }
                        <%end if%>
                        <%if cifrepe = "1" then%>
                            //window.alert(repe + "-" + existe);
			                if (repe == 1 ||existe == "")
			                {
                                parent.pantalla.document.proveedores.action="proveedores.asp?mode=" + cade + "&repe=" + repe;
			                    parent.pantalla.document.proveedores.submit();
                                //ASP 09/01/2012                            
                                 reloadPanelGlobal("<%=viene2%>");
                                //FIN ASP 09/01/2012
			                    document.location="proveedores_bt.asp?mode=browse";
			                }
			            <%else%>
                            parent.pantalla.document.proveedores.action="proveedores.asp?mode=" + cade + "&repe=" + repe;
			                parent.pantalla.document.proveedores.submit();
			                document.location="proveedores_bt.asp?mode=browse";
			            <%end if%>
					}
					break;

				case "cancel": //Cancelar edición
                    if(document.opciones.viene.value=="")
                    {
					    parent.pantalla.document.proveedores.action="proveedores.asp?mode=add";
					    parent.pantalla.document.proveedores.submit();
					    document.location="proveedores_bt.asp?mode=add";
                    }
                    else parent.window.close();
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "search":
			switch (pulsado)
			{
				case "search": //Buscar datos
					break;
			}
			break;
	}
}

</script>
<body class="body_master_ASP">
<%mode=Request.QueryString("mode")
if request.QueryString("noadd")>"" then
	viene= limpiaCadena(request.QueryString("noadd"))
else
	viene=request.form("noadd")
end if

if request.QueryString("viene")>"" then
	viene= limpiaCadena(request.QueryString("viene"))
else
	viene=request.form("viene")
end if%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%EncodeForHtml(mode)%>" />
<input type="hidden" name="noadd" value="<%EncodeForHtml(noadd)%>">
<input type="hidden" name="viene" value="<%EncodeForHtml(viene)%>">
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >

        <%
		if noadd="1" and mode="add" then mode="search"
		if request.querystring("viene")&"">"" then
			if mode="add" then
				%>
                <table id="BUTTONS_CENTER_ASP">
		            <tr>
        				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save');">
		        			<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				        </td>
				        <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					        <%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
		            </tr>
	            </table>
                </div>
			    <%
			end if
		else
            %>
            <table id="BUTTONS_CENTER_ASP">
		        <tr>
            <%
			if mode="browse" then
				if noadd<>"1" then
				    %>
				    <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				    <%
				end if
				if pagsl&""<>"1" then%>
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBTLeft LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				    </td>
				    <td id="iddelete" class="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBTLeftRed LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
				    <%
				end if
				%>
				<td id="idprint" class="CELDABOT" onclick="javascript:Accion('browse','print');">
					<%PintarBotonBTLeft LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				</td>
			    <%
			elseif mode="search" then
				if noadd<>"1" then
				    %>
				    <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBTLeft LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				    <%
				end if
			elseif mode="edit"  then
			    if pagsl&""<>"1" then
			        %>
				    <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					    <%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
    				<%
    			end if
    			%>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					<%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			    <%
			elseif mode="add" then
			    %>
				<td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save');">
					<%PintarBotonBTLeft LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				</td>
				<td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					<%PintarBotonBTLeftRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			    <%
			end if%>
		        </tr>
	        </table>
            </div>
    
    <div id="FILTERS_MASTER_ASP">
      
		<!--<td class=CELDABOT><%=LitBuscar & ": "%>-->
		<select class="IN_S" name="campos">
			<option value="razon_social"><%=LitRSocial%></option>
			<option value="nproveedor"><%=LitNProveedor%></option>
			<option value="nombre"><%=LitNombre%></option>
			<option value="cif"><%=LitCif%></option>
			<option value="contacto"><%=LitContacto%></option>
			<option value="domicilio"><%=LitDomicilio%></option>
			<option value="cp"><%=LitCp%></option>
			<option value="provincia"><%=LitProvincia%></option>
			<option value="pais"><%=LitPais%></option>
			<option value="telefono"><%=LitTel1%></option>
			<option value="telefono2"><%=LitTel2%></option>
			<option value="fax"><%=LitFax%></option>
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
            <input id="KeySearch" class="IN_S" type="text" name="texto" size="15" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
		<!--</td>
		<td class=CELDABOT>-->
			<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
		<!--</td>-->
        <!--</tr>
        </table>-->
    </div>
    <%end if%>

    
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