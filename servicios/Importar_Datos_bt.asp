<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" --> 
<!--#include file="Importar_Datos.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../controlimpresion.inc" -->
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/Master.css.inc" -->
</HEAD>
<script language="JavaScript" src="../jfunciones.js"></script>
<script language="JavaScript" type="text/javascript">
extArray = new Array('dbf','csv','txt'); // <---- Extensiones válidas

function extension(file) {
	allowSubmit = false;
	if (!file) return true;
	path=file.substring(0,file.lastIndexOf("\\")+1);
	file = file.slice(file.lastIndexOf("\\")+1);
	// Sacamos el nombre del archivo (y solucionamos bug Opera 6)
	if (file.indexOf('"') != -1) {
		var archivo = file.substring(0,file.indexOf('"'));
		file = file.substring(0,file.indexOf('"'));
	}
	else var archivo = file;
	// Sacamos la extension del archivo y la pasamos a minusculas
	file = file.slice(file.lastIndexOf(".")+1);
	var ext = file.toLowerCase();
	// Comparamos con los elementos del array
	for (var i = 0; i < extArray.length; i++) {
		if (extArray[i] == ext) { 
			allowSubmit = true;
			break;
		}
	}
	return allowSubmit;
}

//-------------------------------------------------------------------------------

function fichero(f) {
	path=f.substring(0,f.lastIndexOf("\\")+1);
	file = f.slice(f.lastIndexOf("\\")+1);
	// Sacamos el nombre del archivo (y solucionamos bug Opera 6)
	if (file.indexOf('"') != -1) {
		var archivo = file.substring(0,file.indexOf('"'));
		file = file.substring(0,file.indexOf('"'));
	}
	else var archivo = file;
	return archivo.toLowerCase();
}

//-------------------------------------------------------------------------------

function Verificar(proceso,formato) {
	switch (proceso) {
		case "C" : //CLIENTES-------------------------------------------------------------------------------
			if (formato=="dbf") {
				if (fichero(parent.pantalla.document.importar_datos.fichero01.value)=="clientes.dbf") {
					if (parent.pantalla.document.importar_datos.fichero02.value=="" || parent.pantalla.document.importar_datos.fichero03.value=="" || parent.pantalla.document.importar_datos.fichero04.value=="")
						return confirm("<%=LitFaltanFicheros%>");
					else {
						if (fichero(parent.pantalla.document.importar_datos.fichero02.value)!="fpago.dbf" || fichero(parent.pantalla.document.importar_datos.fichero03.value)!="agentes.dbf" || fichero(parent.pantalla.document.importar_datos.fichero04.value)!="provinc.dbf")
							alert("<%=LitFicherosComp%> <%=LitfpagoDBF%> , <%=LitagentesDBF%> y <%=LitprovincDBF%>")
						else return true;
					}
				}
				else {
					alert("<%=LitDebeSelArchivo%> <%=LitclientesDBF%>");
					return false;
				}
			}
			else {
				if (fichero(parent.pantalla.document.importar_datos.fichero01.value)=="clientes.csv") {
					if (parent.pantalla.document.importar_datos.fichero02.value=="" || parent.pantalla.document.importar_datos.fichero03.value=="")
						return confirm("<%=LitFaltanFicheros%>");
					else {
						if (fichero(parent.pantalla.document.importar_datos.fichero02.value)!="fpago.csv" || fichero(parent.pantalla.document.importar_datos.fichero03.value)!="comercial.csv")
							alert("<%=LitFicherosComp%> <%=LitfpagoCSV%> y <%=LitcomercialCSV%>")
						else return true;
					}
				}
				else {
					alert("<%=LitDebeSelArchivo%> <%=LitclientesCSV%>");
					return false;
				}
			}
			break;
			
		case "P" : //PROVEEDORES----------------------------------------------------------------------------
			if (formato=="dbf") {
				if (fichero(parent.pantalla.document.importar_datos.fichero01.value)=="proveedo.dbf") {
					if (parent.pantalla.document.importar_datos.fichero02.value=="") return confirm("<%=LitFaltanFicheros%>");
					else {
						if (fichero(parent.pantalla.document.importar_datos.fichero02.value)!="fpago.dbf") alert("<%=LitFicherosComp%> <%=LitfpagoDBF%>")
						else return true;
					}
				}
				else {
					alert("<%=LitDebeSelArchivo%> <%=LitproveedoDBF%>");
					return false;
				}
			}
			else {
				if (fichero(parent.pantalla.document.importar_datos.fichero01.value)=="proveedo.csv") {
					if (parent.pantalla.document.importar_datos.fichero02.value=="") return confirm("<%=LitFaltanFicheros%>");
					else
					{
						if (fichero(parent.pantalla.document.importar_datos.fichero02.value)!="fpago.csv") alert("<%=LitFicherosComp%> <%=LitfpagoCSV%>")
						else return true;
					}
				}
				else {
					alert("<%=LitDebeSelArchivo%> <%=LitproveedoCSV%>");
					return false;
				}
			}
			break;
		
		case "A" : //ARTICULOS------------------------------------------------------------------------------
			if (formato=="dbf") {
				if (fichero(parent.pantalla.document.importar_datos.fichero01.value)=="articulo.dbf") {
					if (parent.pantalla.document.importar_datos.fichero02.value=="" || parent.pantalla.document.importar_datos.fichero03.value=="")
						return confirm("<%=LitFaltanFicheros%>");
					else {
						if (fichero(parent.pantalla.document.importar_datos.fichero02.value)!="familias.dbf" || fichero(parent.pantalla.document.importar_datos.fichero03.value)!="stocks.dbf")
							alert("<%=LitFicherosComp%> <%=LitfamiliasDBF%> y <%=LitstocksDBF%>")
						else return true;
					}
				}
				else {
					alert("<%=LitDebeSelArchivo%> <%=LitarticuloDBF%>");
					return false;
				}
			}
			else {
				if (fichero(parent.pantalla.document.importar_datos.fichero01.value)=="articulo.csv") {
					if (parent.pantalla.document.importar_datos.fichero02.value=="" || parent.pantalla.document.importar_datos.fichero03.value=="")
						return confirm("<%=LitFaltanFicheros%>");
					else {
						if (fichero(parent.pantalla.document.importar_datos.fichero02.value)!="familias.csv" || fichero(parent.pantalla.document.importar_datos.fichero03.value)!="stocks.csv")
							alert("<%=LitFicherosComp%> <%=LitfamiliasCSV%> y <%=LitstocksCSV%>")
						else return true;
					}
				}
				else {
					alert("<%=LitDebeSelArchivo%> <%=LitarticuloCSV%>");
					return false;
				}
			}
			break;
			
		case "S" : //SUMINISTROS------------------------------------------------------------------------------
			return true;
        case "F" : //ARTICULOS------------------------------------------------------------------------------
			if (fichero(parent.pantalla.document.importar_datos.fichero01.value)=="<%=LitTarjetasCSV%>") {
                return true;
			}
			else {
				alert("<%=LitDebeSelArchivo%> <%=LitTarjetasCSV%>");
				return false;
			}
			break;
	}
}

//---------------------------------------------------------------------------------------------------------

function ValidarCampos()
{
	if (parent.pantalla.document.importar_datos.fichero01.value=="") {
		alert("<%=LitObligatorio%>");
		return false;
	}
	
	if (parent.pantalla.document.importar_datos.s.value=="1") { //SUMINISTROS
		if (parent.pantalla.document.importar_datos.proceso.checked) var proceso="S";
		
		if (parent.pantalla.document.importar_datos.formato.checked) var formato="txt";
	}
	else { //FICHEROS DE CLIENTES
		if (parent.pantalla.document.importar_datos.proceso[0].checked) var proceso="C";
		else {
			if (parent.pantalla.document.importar_datos.proceso[1].checked) var proceso="P";
			else if (parent.pantalla.document.importar_datos.proceso[2].checked) var proceso = "A";
                else  var proceso = "F";
		}

		if (parent.pantalla.document.importar_datos.formato[0].checked) var formato="dbf";
		else var formato="csv";
	}

	switch (proceso) {
		case "C" : //CLIENTES
			if (formato=="dbf") {
				if (extension(parent.pantalla.document.importar_datos.fichero01.value) && extension(parent.pantalla.document.importar_datos.fichero02.value) &&
					extension(parent.pantalla.document.importar_datos.fichero03.value) && extension(parent.pantalla.document.importar_datos.fichero04.value)) 
					return Verificar("C","dbf");
				else {
					alert("<%=LitSoloArchivosDBF%>");
					return false;
				}
			}
			else {
				if (extension(parent.pantalla.document.importar_datos.fichero01.value) && extension(parent.pantalla.document.importar_datos.fichero02.value) &&
					extension(parent.pantalla.document.importar_datos.fichero03.value)) 
					return Verificar("C","csv");
				else {
					alert("<%=LitSoloArchivosCSV%>");
					return false;
				}
			}
			break;
			
		case "P" : //PROVEEDORES
			if (formato=="dbf") {
				if (extension(parent.pantalla.document.importar_datos.fichero01.value) && extension(parent.pantalla.document.importar_datos.fichero02.value))
					return Verificar("P","dbf");
				else {
					alert("<%=LitSoloArchivosDBF%>");
					return false;
				}
			}
			else {
				if (extension(parent.pantalla.document.importar_datos.fichero01.value) && extension(parent.pantalla.document.importar_datos.fichero02.value))
					return Verificar("P","csv");
				else {
					alert("<%=LitSoloArchivosCSV%>");
					return false;
				}
			}
			break;

		case "A" : //ARTICULOS
			if (formato=="dbf") {
				if (extension(parent.pantalla.document.importar_datos.fichero01.value) && extension(parent.pantalla.document.importar_datos.fichero02.value) &&
					extension(parent.pantalla.document.importar_datos.fichero03.value))
					return Verificar("A","dbf");
				else {
					alert("<%=LitSoloArchivosDBF%>");
					return false;
				}
			} else {
				if (extension(parent.pantalla.document.importar_datos.fichero01.value) && extension(parent.pantalla.document.importar_datos.fichero02.value) &&
					extension(parent.pantalla.document.importar_datos.fichero03.value))
					return Verificar("A","csv");
				else {
					alert("<%=LitSoloArchivosCSV%>");
					return false;
				}
			}
			break;
		
		case "S" : //SUMINISTROS
			if (formato=="txt") {
				if (extension(parent.pantalla.document.importar_datos.fichero01.value)) return Verificar("S","txt");
				else {
					alert("<%=LitSoloArchivosTXT%>");
					return false;
				}
			}
			break;
        case "F" : //TARJETAS FIDELIZACION
				if (extension(parent.pantalla.document.importar_datos.fichero01.value) )
					return Verificar("F","csv");
				else {
					alert("<%=LitSoloArchivosCSV%>");
					return false;
				}
			
			break;
	}
	return true;
}

//---------------------------------------
//Comprobar los tamaños -----------------
//---------------------------------------
function SizeOK() {
 

    maximo_length=parseFloat("<%=cLng(LimiteUploadImp)%>)");
    if (!checkFileSize(parent.pantalla.document.importar_datos.fichero01, maximo_length)) {
        window.alert(parent.pantalla.document.importar_datos.fichero01.value + ".\n<%=LitLimiteMaximoUpload%><%=formatnumber(((LimiteUploadImp-390)/1024),0,-1,0,-1)%>Kb");
        return false;
    }


	if (parent.pantalla.document.importar_datos.fichero02!=null) {
		if (parent.pantalla.document.importar_datos.fichero02.value!="") {
		    if (!checkFileSize(parent.pantalla.document.importar_datos.fichero02, maximo_length)) {
				window.alert(parent.pantalla.document.importar_datos.fichero02.value + ".\n<%=LitLimiteMaximoUpload%><%=formatnumber(((LimiteUploadImp-390)/1024),0,-1,0,-1)%>Kb");
				return false;
			}
		}
	}
	if (parent.pantalla.document.importar_datos.fichero03!=null) {
		if (parent.pantalla.document.importar_datos.fichero03.value!="") {
		    if (!checkFileSize(parent.pantalla.document.importar_datos.fichero03, maximo_length)) {
				window.alert(parent.pantalla.document.importar_datos.fichero03.value + ".\n<%=LitLimiteMaximoUpload%><%=formatnumber(((LimiteUploadImp-390)/1024),0,-1,0,-1)%>Kb");
				return false;
			}
		}
	}
	if (parent.pantalla.document.importar_datos.fichero04!=null) {
		if (parent.pantalla.document.importar_datos.fichero04.value!="") {
		    if (!checkFileSize(parent.pantalla.document.importar_datos.fichero04, maximo_length)) {
				window.alert(parent.pantalla.document.importar_datos.fichero04.value + ".\n<%=LitLimiteMaximoUpload%><%=formatnumber(((LimiteUploadImp-390)/1024),0,-1,0,-1)%>Kb");
				return false;
			}
		}
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {
	switch (mode) {
		case "select":
			switch (pulsado) {
				case "importar": //enviar registro
					if (ValidarCampos()) {
						if (SizeOK()) {
							parent.pantalla.document.getElementById("waitBoxOculto").style.visibility='visible';
							parent.pantalla.importar_datos.action = "Importar_Datos.asp?mode=procesa&ncliente=" + parent.pantalla.document.importar_datos.codempresa.value + "&tdocumento=" + parent.pantalla.document.importar_datos.tituloDoc.value;
							parent.pantalla.document.importar_datos.submit();
							document.location="importar_datos_bt.asp?mode=procesa";
						}
					}
					break;
			}
			break;
		case "informe":
			switch (pulsado) {
				case "imprimir" :
					parent.pantalla.focus();
					printWindow();
					break
				case "cancel" :
				    parent.pantalla.document.location = "importar_datos.asp?mode=select&ncliente=" + parent.pantalla.document.importar_datos.codempresa.value + "&borra=SI&s=" + parent.pantalla.document.importar_datos.s.value + "&tdocumento=" + parent.pantalla.document.importar_datos.tituloDoc.value;
					document.location="importar_datos_bt.asp?mode=select";
			}
			break;
	}
}
</script>
<body class="body_master_ASP">
<%EscribeControlImpresion "Importar_Datos.asp"
mode=Request.QueryString("mode")%>
<form name="opciones" method="post">
    <div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
            <table id="BUTTONS_CENTER_ASP">
                <tr>
		            <%if mode="select" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('select','importar');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTAR%>
				        </td>
			        <%elseif mode="informe" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('informe','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIR%>
				        </td>
			            <td class="CELDABOT" onclick="javascript:Accion('informe','cancel');">
				            <%PintarBotonBTRed LITBOTCANCELAR,ImgVolver,ParamImgVolver,LITBOTCANCELAR%>
			            </td>
			        <%end if%>
		        </tr>
	        </table>
        </div>
    </div>
<%ImprimirPie_bt%>	
</form>
</BODY>
</HTML>
