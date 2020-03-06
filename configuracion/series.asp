<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablasresponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<TITLE><%=LitTituloSD%></TITLE>
    <% 
        ' ' 07/05/2019 Se realiza cambios de ciberseguridad
        ' - enc.EncodeForJavascript(param) -> Cross Site Scripting (XSS)
        ' - enc.EncodeForHtmlAttribute(param) -> Cross Site Scripting (XSS)
        ' - limpiaCadena() -> Inyección SQL
        dim  enc
        set enc = Server.CreateObject("Owasp_Esapi.Encoder")
     %>  

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
</HEAD>
<% 

set rst = server.CreateObject("ADODB.Recordset")
si_tiene_gestionFolios = false
rst.cursorlocation=3
rst.open "SELECT gestion_folios FROM configuracion with(nolock) where nempresa='" & session("ncliente") & "'",session("dsn_cliente"), adOpenKeySet, adlockOptimistic
if not rst.EOF and rst("gestion_folios") = true then
    si_tiene_gestionFolios = true
end if
rst.Close 

%>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
function deshabilitaCampo(){
  	document.series.e_facturable.disabled=true;
	document.series.e_facturable.value=1;
	document.series.e_facturable.checked=true;
}

function comprobar_tipodoc(modo)
{
	if(modo=="1")
	{
        if (document.series.e_documento.value=="FACTURA A CLIENTE" || document.series.e_documento.value=="FACTURA DE PROVEEDOR")
        {
	        document.series.e_facturable.value=1;
	        document.series.e_facturable.checked=true;
	        document.series.e_facturable.disabled=true;
        }
        else document.series.e_facturable.disabled=false;

		if(document.series.e_documento.value=="HOJA DE GASTOS" || document.series.e_documento.value=="MOVIMIENTOS ENTRE ALMACENES"
		    || document.series.e_documento.value=="PEDIDOS ENTRE ALMACENES"
			|| document.series.e_documento.value=="CATALOGO" || document.series.e_documento.value=="ORDEN"
			|| document.series.e_documento.value=="PARTE" || document.series.e_documento.value=="INCIDENCIA"
			|| document.series.e_documento.value=="ORDEN DE FABRICACION"){
			document.series.h_cliente.value="";
			document.series.nombre_cli.value="";
		}
		if (document.series.e_documento.value==""){
		}
		else{
			//ocultamos todos los formatos de impresion
			document.getElementById("formatos_imp_ninguno").style.display="none";
			document.getElementById("formatos_imp_alb_cli").style.display="none";
			document.getElementById("formatos_imp_dev_cli").style.display="none";
			document.getElementById("formatos_imp_fac_cli").style.display="none";
			document.getElementById("formatos_imp_ped_cli").style.display="none";
			document.getElementById("formatos_imp_pre_cli").style.display="none";
			document.getElementById("formatos_imp_fac_pro").style.display="none";
			document.getElementById("formatos_imp_alb_pro").style.display="none";
			document.getElementById("formatos_imp_ped_pro").style.display="none";
			document.getElementById("formatos_imp_dev_pro").style.display="none";
			document.getElementById("formatos_imp_ord_fab").style.display="none";
			document.getElementById("formatos_imp_ord").style.display="none";
			document.getElementById("formatos_imp_mov").style.display="none";
			document.getElementById("formatos_imp_ped_ti").style.display="none";
			document.getElementById("formatos_imp_parte").style.display="none";
			document.getElementById("formatos_imp_cat").style.display="none";
			document.getElementById("formatos_imp_inc").style.display="none";
			//ahora ponemos el formato de impresion adecuado
			elegido=0;
			if (document.series.e_documento.value=="ALBARAN DE SALIDA"){
				document.getElementById("formatos_imp_alb_cli").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="DEVOLUCION DE CLIENTE"){
				document.getElementById("formatos_imp_dev_cli").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="FACTURA A CLIENTE"){
				document.getElementById("formatos_imp_fac_cli").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="PEDIDO DE CLIENTE"){
				document.getElementById("formatos_imp_ped_cli").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="PRESUPUESTO A CLIENTE"){
				document.getElementById("formatos_imp_pre_cli").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="FACTURA DE PROVEEDOR"){
				document.getElementById("formatos_imp_fac_pro").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="ALBARAN DE PROVEEDOR"){
				document.getElementById("formatos_imp_alb_pro").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="PEDIDO A PROVEEDOR"){
				document.getElementById("formatos_imp_ped_pro").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="DEVOLUCION A PROVEEDOR"){
				document.getElementById("formatos_imp_dev_pro").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="ORDEN DE FABRICACION"){
				document.getElementById("formatos_imp_ord_fab").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="ORDEN"){
				document.getElementById("formatos_imp_ord").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="MOVIMIENTOS ENTRE ALMACENES"){
				document.getElementById("formatos_imp_mov").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="PEDIDOS ENTRE ALMACENES"){
				document.getElementById("formatos_imp_ped_ti").style.display="";
				elegido=1;
			}				
			if (document.series.e_documento.value=="PARTE DE TRABAJO"){
				document.getElementById("formatos_imp_parte").style.display="";
				elegido=1;
			}
			if (document.series.e_documento.value=="CATALOGO"){
				document.getElementById("formatos_imp_cat").style.display="";
				elegido=1;
			}
			
			if (document.series.e_documento.value=="INCIDENCIA"){
				document.getElementById("formatos_imp_inc").style.display="";
				elegido=1;
			}
			if (elegido==0) document.getElementById("formatos_imp_ninguno").style.display="";

			if (document.series.h_cliente.value=="") document.series.e_doc_aux.value=document.series.e_documento.value;
			else{
				no_error=0;
				if ((document.series.e_doc_aux.value=="ALBARAN DE SALIDA" ||
				document.series.e_doc_aux.value=="DEVOLUCION DE CLIENTE" ||
				document.series.e_doc_aux.value=="FACTURA A CLIENTE" ||
				document.series.e_doc_aux.value=="HOJA DE GASTOS" ||
				document.series.e_doc_aux.value=="PEDIDO DE CLIENTE" ||
				document.series.e_doc_aux.value=="PRESUPUESTO A CLIENTE" ||
				document.series.e_doc_aux.value=="TICKET") &&
				(document.series.e_documento.value=="FACTURA DE PROVEEDOR" ||
				document.series.e_documento.value=="ALBARAN DE PROVEEDOR" ||
				document.series.e_documento.value=="PEDIDO A PROVEEDOR" ||
				document.series.e_documento.value=="DEVOLUCION A PROVEEDOR")){
					window.alert("<%=LitNoCambiarTipodoc%>");
					document.series.e_documento.value=document.series.e_doc_aux.value;
					no_error=1;
				}
				if ((document.series.e_documento.value=="ALBARAN DE SALIDA" ||
				document.series.e_documento.value=="DEVOLUCION DE CLIENTE" ||
				document.series.e_documento.value=="FACTURA A CLIENTE" ||
				document.series.e_documento.value=="HOJA DE GASTOS" ||
				document.series.e_documento.value=="PEDIDO DE CLIENTE" ||
				document.series.e_documento.value=="PRESUPUESTO A CLIENTE" ||
				document.series.e_documento.value=="TICKET") &&
				(document.series.e_doc_aux.value=="FACTURA DE PROVEEDOR" ||
				document.series.e_doc_aux.value=="ALBARAN DE PROVEEDOR" ||
				document.series.e_doc_aux.value=="PEDIDO A PROVEEDOR" ||
				document.series.e_doc_aux.value=="DEVOLUCION A PROVEEDOR")){
					window.alert("<%=LitNoCambiarTipodoc%>");
					document.series.e_documento.value=document.series.e_doc_aux.value;
					no_error=1;
				}

				if (no_error==0) document.series.e_doc_aux.value=document.series.e_documento.value;
			}
		}
	}
	else{
        if (document.series.i_documento.value=="FACTURA A CLIENTE" || document.series.i_documento.value=="FACTURA DE PROVEEDOR")
        {
	        document.series.i_facturable.value=1;
	        document.series.i_facturable.checked=true;
	        document.series.i_facturable.disabled=true;
        }
        else
        {
	        document.series.i_facturable.disabled=false;
        }

		if(document.series.i_documento.value=="HOJA DE GASTOS" || document.series.i_documento.value=="MOVIMIENTOS ENTRE ALMACENES"
			|| document.series.i_documento.value=="PEDIDOS ENTRE ALMACENES"
			|| document.series.i_documento.value=="CATALOGO" || document.series.i_documento.value=="ORDEN"
			|| document.series.i_documento.value=="PARTE" || document.series.i_documento.value=="INCIDENCIA"
			|| document.series.i_documento.value=="ORDEN DE FABRICACION"){
			document.series.h_cliente.value="";
			document.series.nombre_cli.value="";
		}

		if (document.series.i_documento.value==""){
		}
		else{
				document.getElementById("formatos_imp_ninguno").style.display="none";
				document.getElementById("formatos_imp_alb_cli").style.display="none";
				document.getElementById("formatos_imp_dev_cli").style.display="none";
				document.getElementById("formatos_imp_fac_cli").style.display="none";
				document.getElementById("formatos_imp_ped_cli").style.display="none";
				document.getElementById("formatos_imp_pre_cli").style.display="none";
				document.getElementById("formatos_imp_fac_pro").style.display="none";
				document.getElementById("formatos_imp_alb_pro").style.display="none";
				document.getElementById("formatos_imp_ped_pro").style.display="none";
				document.getElementById("formatos_imp_dev_pro").style.display="none";
				document.getElementById("formatos_imp_ord_fab").style.display="none";
				document.getElementById("formatos_imp_ord").style.display="none";
				document.getElementById("formatos_imp_mov").style.display="none";
				document.getElementById("formatos_imp_ped_ti").style.display="none";
				document.getElementById("formatos_imp_parte").style.display="none";
				document.getElementById("formatos_imp_cat").style.display="none";
				document.getElementById("formatos_imp_inc").style.display="none";
				elegido=0;
				//ahora ponemos el formato de impresion adecuado
				if (document.series.i_documento.value=="ALBARAN DE SALIDA"){
					document.getElementById("formatos_imp_alb_cli").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="DEVOLUCION DE CLIENTE"){
					document.getElementById("formatos_imp_dev_cli").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="FACTURA A CLIENTE"){
					document.getElementById("formatos_imp_fac_cli").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="PEDIDO DE CLIENTE"){
					document.getElementById("formatos_imp_ped_cli").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="PRESUPUESTO A CLIENTE"){
					document.getElementById("formatos_imp_pre_cli").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="FACTURA DE PROVEEDOR"){
					document.getElementById("formatos_imp_fac_pro").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="ALBARAN DE PROVEEDOR"){
					document.getElementById("formatos_imp_alb_pro").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="PEDIDO A PROVEEDOR"){
					document.getElementById("formatos_imp_ped_pro").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="DEVOLUCION A PROVEEDOR"){
					document.getElementById("formatos_imp_dev_pro").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="ORDEN DE FABRICACION"){
					document.getElementById("formatos_imp_ord_fab").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="ORDEN"){
					document.getElementById("formatos_imp_ord").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="MOVIMIENTOS ENTRE ALMACENES"){
					document.getElementById("formatos_imp_mov").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="PEDIDOS ENTRE ALMACENES"){
					document.getElementById("formatos_imp_ped_ti").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="PARTE DE TRABAJO"){
					document.getElementById("formatos_imp_parte").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="CATALOGO"){
					document.getElementById("formatos_imp_cat").style.display="";
					elegido=1;
				}
				if (document.series.i_documento.value=="INCIDENCIA"){
					document.getElementById("formatos_imp_inc").style.display="";
					elegido=1;
				}
				if (elegido==0){
					document.getElementById("formatos_imp_ninguno").style.display="";
				}

			if (document.series.h_cliente.value==""){
				document.series.i_doc_aux.value=document.series.i_documento.value;
			}
			else{
				no_error=0;

				if ((document.series.i_doc_aux.value=="ALBARAN DE SALIDA" ||
				document.series.i_doc_aux.value=="DEVOLUCION DE CLIENTE" ||
				document.series.i_doc_aux.value=="FACTURA A CLIENTE" ||
				document.series.i_doc_aux.value=="HOJA DE GASTOS" ||
				document.series.i_doc_aux.value=="PEDIDO DE CLIENTE" ||
				document.series.i_doc_aux.value=="PRESUPUESTO A CLIENTE" ||
				document.series.i_doc_aux.value=="TICKET") &&
				(document.series.i_documento.value=="FACTURA DE PROVEEDOR" ||
				document.series.i_documento.value=="ALBARAN DE PROVEEDOR" ||
				document.series.i_documento.value=="PEDIDO A PROVEEDOR" ||
				document.series.i_documento.value=="DEVOLUCION A PROVEEDOR")){
					window.alert("<%=LitNoCambiarTipodoc%>");
					document.series.i_documento.value=document.series.i_doc_aux.value;
					no_error=1;
				}
				if ((document.series.i_documento.value=="ALBARAN DE SALIDA" ||
				document.series.i_documento.value=="DEVOLUCION DE CLIENTE" ||
				document.series.i_documento.value=="FACTURA A CLIENTE" ||
				document.series.i_documento.value=="HOJA DE GASTOS" ||
				document.series.i_documento.value=="PEDIDO DE CLIENTE" ||
				document.series.i_documento.value=="PRESUPUESTO A CLIENTE" ||
				document.series.i_documento.value=="TICKET") &&
				(document.series.i_doc_aux.value=="FACTURA DE PROVEEDOR" ||
				document.series.i_doc_aux.value=="ALBARAN DE PROVEEDOR" ||
				document.series.i_doc_aux.value=="PEDIDO A PROVEEDOR" ||
				document.series.i_doc_aux.value=="DEVOLUCION A PROVEEDOR")){
					window.alert("<%=LitNoCambiarTipodoc%>");
					document.series.i_documento.value=document.series.i_doc_aux.value;
					no_error=1;
				}


				if (no_error==0){
					document.series.i_doc_aux.value=document.series.i_documento.value;
				}

			}
		}
	}
}
function buscar_clipro(modo){
    if(modo=="1")
    {
		if (document.series.e_documento.value=="") window.alert("<%=LitNoTipDocElegido%>");
		else{
			if (document.series.e_documento.value!="HOJA DE GASTOS" && document.series.e_documento.value!="MOVIMIENTOS ENTRE ALMACENES"){
				if (document.series.e_documento.value=="ALBARAN DE SALIDA" ||
					document.series.e_documento.value=="DEVOLUCION DE CLIENTE" ||
					document.series.e_documento.value=="FACTURA A CLIENTE" ||
					document.series.e_documento.value=="PEDIDO DE CLIENTE" ||
					document.series.e_documento.value=="PRESUPUESTO A CLIENTE" ||
					document.series.e_documento.value=="TICKET"){

					AbrirVentana('../ventas/clientes_buscar.asp?ndoc=<%=Formulario%>&titulo=<%=LitSelCliente%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>);
				}
				else{
					if (document.series.e_documento.value=="FACTURA DE PROVEEDOR" ||
		    			document.series.e_documento.value=="ALBARAN DE PROVEEDOR" ||
	    				document.series.e_documento.value=="PEDIDO A PROVEEDOR" ||
    					document.series.e_documento.value=="DEVOLUCION A PROVEEDOR"){
					    AbrirVentana('../compras/proveedores_busqueda.asp?ndoc=<%=Formulario%>&titulo=<%=LitSelProveedor%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>);
					}
					else window.alert("<%=LitTipodocNoCliPro%>");
				}
			}
			else window.alert("<%=LitTipodocNoCliPro%>");
		}
	}
    else{
	    if (document.series.i_documento.value=="") window.alert("<%=LitNoTipDocElegido%>");
	    else{
		    if (document.series.i_documento.value!="HOJA DE GASTOS" && document.series.i_documento.value!="MOVIMIENTOS ENTRE ALMACENES"){
			    if (document.series.i_documento.value=="ALBARAN DE SALIDA" ||
				    document.series.i_documento.value=="DEVOLUCION DE CLIENTE" ||
				    document.series.i_documento.value=="FACTURA A CLIENTE" ||
				    document.series.i_documento.value=="HOJA DE GASTOS" ||
				    document.series.i_documento.value=="PEDIDO DE CLIENTE" ||
				    document.series.i_documento.value=="PRESUPUESTO A CLIENTE" ||
				    document.series.i_documento.value=="TICKET"){

				    AbrirVentana('../ventas/clientes_buscar.asp?ndoc=<%=Formulario%>&titulo=<%=LitSelCliente%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>);
			    }
			    else{
				    if (document.series.i_documento.value=="FACTURA DE PROVEEDOR" ||
		    		    document.series.i_documento.value=="ALBARAN DE PROVEEDOR" ||
	    			    document.series.i_documento.value=="PEDIDO A PROVEEDOR" ||
    				    document.series.i_documento.value=="DEVOLUCION A PROVEEDOR"){
				        AbrirVentana('../compras/proveedores_busqueda.asp?ndoc=<%=Formulario%>&titulo=<%=LitSelProveedor%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>);
				    }
				    else window.alert("<%=LitTipodocNoCliPro%>");
			    }
		    }
		    else window.alert("<%=LitTipodocNoCliPro%>");
	    }
    }
}

function Editar(p_codigo, p_npagina, p_campo, p_criterio, p_texto,viene) {
	
	if(viene=="asistente"){
	    document.location="series.asp?mode=edit&p_codigo=" + p_codigo
                                             +"&npagina="+ p_npagina
                                             +"&campo="  + p_campo
                                             +"&texto="  + p_texto
                                             +"&criterio=" + p_criterio
                                             +"&viene=" + viene;
	    parent.botones.document.location="series_bt.asp?mode=edit&viene=asistente"; 
    }
	else{
	    document.location="series.asp?mode=edit&p_codigo=" + p_codigo
                                             +"&npagina="+ p_npagina
                                             +"&campo="  + p_campo
                                             +"&texto="  + p_texto
                                             +"&criterio=" + p_criterio
                                             +"#" + p_codigo; 
	    parent.botones.document.location="series_bt.asp?mode=edit";
	}
}

function limpiarCliente() {
	document.series.h_cliente.value="";
	document.series.nombre_cli.value="";
}

<% if si_tiene_gestionFolios then %>
function mostrarNfolios(mode){

    if(mode == 1){
        nfolio1 = "e_nfolio1"
        nfolio2 = "e_nfolio2"
        elementoSelect = "e_documento"
    }
    else{
        nfolio1 = "i_nfolio1"
        nfolio2 = "i_nfolio2"
        elementoSelect = "i_documento"
    }
    if (document.getElementById(elementoSelect).value=="FACTURA A CLIENTE"){
		document.getElementById(nfolio1).style.display="";
		document.getElementById(nfolio2).style.display="";
	}
	else{
	    if(mode == 1){
	        document.getElementById(nfolio1).value = "";
            document.getElementById(nfolio2).value = "";
        }
        document.getElementById(nfolio1).style.display="none";
        document.getElementById(nfolio2).style.display="none";
    }
}

<% end if%>

</script>
<body bgcolor=<%=color_blau%>>
<%
Function GetLanguage(usuario, ncliente)
        selectQuery = "SELECT IDIOMA FROM CLIENTES AS C WITH(NOLOCK) LEFT JOIN CLIENTES_USERS CU ON C.NCLIENTE = CU.NCLIENTE AND CU.USUARIO = '"& usuario &"' WHERE C.NCLIENTE = '"& ncliente &"'"
        set rstI=server.CreateObject("ADODB.Recordset")
        rstI.Open selectQuery,DSNIlion,adOpenKeyset, adLockOptimistic
        if not rstI.EOF then
            idiomaUsuario = rstI("IDIOMA")
        end if
        rstI.close
        GetLanguage = idiomaUsuario
End Function
idiomaUser = GetLanguage(session("usuario"), session("ncliente"))

set rstAux = server.CreateObject("ADODB.Recordset")

si_tiene_modulo_ccostes=ModuloContratado(session("ncliente"),ModCcostes_Gestion) '**rgu:2/9/2009
si_tiene_modulo_ecomerce=ModuloContratado(session("ncliente"),ModEComerce) ' -- 26/02/2010 Albert





'*********************************************************************************************************'
' CODIGO PRINCIPAL DE LA PAGINA  *************************************************************************'
'*********************************************************************************************************'

if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<form name="series" method="post" action="series.asp">
    <%Formulario="series"
    PintarCabecera "series.asp"
    

    'insertamos si nos llegan los valores'
	'Leer parámetros de la página'
	    mode=request("mode")	   
	'AMP 28/07/2010 : Añadimos parametro viene para adaptar la configuración de series al asistente de puesta en marcha.
		viene=Request.QueryString("viene")

		p_i_Nserie=limpiaCadena(Request.Form("i_Nserie"))
		p_i_Nombre=limpiaCadena(request.form ("i_Nombre"))
		p_i_empresa=limpiaCadena(request.form ("i_Empresa"))
		p_i_Contador=limpiaCadena(Request.form("i_Contador"))
		p_i_nfolio1=limpiaCadena(Request.form("i_nfolio1"))
		p_i_nfolio2=limpiaCadena(Request.form("i_nfolio2"))
		p_i_FechaModificacion = Date
		p_i_Documento=limpiaCadena(Request.form("i_Documento"))
		p_i_almacen=limpiaCadena(Request.form("i_almacen"))
		p_i_cta_ventas=limpiaCadena(Request.form("i_cta_ventas"))
		p_i_cta_compras=limpiaCadena(Request.form("i_cta_compras"))
		p_i_cta_aventas=limpiaCadena(Request.form("i_cta_aventas"))
		p_i_cta_acompras=limpiaCadena(Request.form("i_cta_acompras"))
		p_i_cta_caja=limpiaCadena(Request.form("i_cta_caja"))
		p_i_cta_rfventas=limpiaCadena(Request.form("i_cta_rfventas"))
		p_i_cta_rfcompras=limpiaCadena(Request.form("i_cta_rfcompras"))
		p_i_cta_retventas=limpiaCadena(Request.form("i_cta_retventas"))
		p_i_cta_retcompras=limpiaCadena(Request.form("i_cta_retcompras"))
		p_i_pordefecto=limpiaCadena(request.form("i_pordefecto"))
		p_i_facturable=limpiaCadena(request.form("i_facturable"))
		p_i_ocultarmkp=limpiaCadena(request.form("i_ocultarmkp"))
		p_i_cliente=limpiaCadena(request.form("h_cliente"))
		p_i_formato_imp_alb_cli=limpiaCadena(request.form("i_formato_imp_alb_cli"))
		p_i_formato_imp_dev_cli=limpiaCadena(request.form("i_formato_imp_dev_cli"))
		p_i_formato_imp_fac_cli=limpiaCadena(request.form("i_formato_imp_fac_cli"))
		p_i_formato_imp_ped_cli=limpiaCadena(request.form("i_formato_imp_ped_cli"))
		p_i_formato_imp_pre_cli=limpiaCadena(request.form("i_formato_imp_pre_cli"))
		p_i_formato_imp_fac_pro=limpiaCadena(request.form("i_formato_imp_fac_pro"))
		p_i_formato_imp_alb_pro=limpiaCadena(request.form("i_formato_imp_alb_pro"))
		p_i_formato_imp_ped_pro=limpiaCadena(request.form("i_formato_imp_ped_pro"))
		p_i_formato_imp_dev_pro=limpiaCadena(request.form("i_formato_imp_dev_pro"))
		p_i_formato_imp_ord_fab=limpiaCadena(request.form("i_formato_imp_ord_fab"))
		p_i_formato_imp_ord=limpiaCadena(request.form("i_formato_imp_ord"))
		p_i_formato_imp_mov=limpiaCadena(request.form("i_formato_imp_mov"))
		p_i_formato_imp_mov=limpiaCadena(request.form("i_formato_imp_ped_ti"))
		p_i_formato_imp_parte=limpiaCadena(request.form("i_formato_imp_parte"))
		p_i_formato_imp_cat=limpiaCadena(request.form("i_formato_imp_cat"))
		p_i_formato_imp_inc=limpiaCadena(request.form("i_formato_imp_inc"))

		p_h_codigo=limpiaCadena(Request.Form("h_codigo"))
		checkCadena p_h_codigo

		p_e_Nserie=limpiaCadena(Request.Form("e_Nserie"))
		p_e_Nombre=limpiaCadena(request.form ("e_Nombre"))
		p_e_Empresa=limpiaCadena(request.form("e_Empresa"))
		p_e_Contador=limpiaCadena(Request.form("e_Contador"))
		p_e_nfolio1=limpiaCadena(Request.form("e_nfolio1"))
		p_e_nfolio2=limpiaCadena(Request.form("e_nfolio2"))
		p_e_FechaModificacion = Date
		p_e_Documento=limpiaCadena(Request.form("e_Documento"))
		p_e_ccostes=limpiaCadena(Request.form("e_ccostes")) '**rgu 4/9/2009
		p_e_almacen=limpiaCadena(Request.form("e_almacen"))
		p_e_cta_ventas=limpiaCadena(Request.form("e_cta_ventas"))
		p_e_cta_compras=limpiaCadena(Request.form("e_cta_compras"))
		p_e_cta_aventas=limpiaCadena(Request.form("e_cta_aventas"))
		p_e_cta_acompras=limpiaCadena(Request.form("e_cta_acompras"))
		p_e_cta_caja=limpiaCadena(Request.form("e_cta_caja"))
		p_e_cta_rfventas=limpiaCadena(Request.form("e_cta_rfventas"))
		p_e_cta_rfcompras=limpiaCadena(Request.form("e_cta_rfcompras"))
		p_e_cta_retventas=limpiaCadena(Request.form("e_cta_retventas"))
		p_e_cta_retcompras=limpiaCadena(Request.form("e_cta_retcompras"))
		p_e_pordefecto=limpiaCadena(request.form("e_pordefecto"))
		p_e_facturable=limpiaCadena(request.form("e_facturable"))
		p_e_cliente=limpiaCadena(request.form("h_cliente"))
		p_e_formato_imp_alb_cli=limpiaCadena(request.form("e_formato_imp_alb_cli"))
		p_e_formato_imp_dev_cli=limpiaCadena(request.form("e_formato_imp_dev_cli"))
		p_e_formato_imp_fac_cli=limpiaCadena(request.form("e_formato_imp_fac_cli"))
		p_e_formato_imp_ped_cli=limpiaCadena(request.form("e_formato_imp_ped_cli"))
		p_e_formato_imp_pre_cli=limpiaCadena(request.form("e_formato_imp_pre_cli"))
		p_e_formato_imp_fac_pro=limpiaCadena(request.form("e_formato_imp_fac_pro"))
		p_e_formato_imp_alb_pro=limpiaCadena(request.form("e_formato_imp_alb_pro"))
		p_e_formato_imp_ped_pro=limpiaCadena(request.form("e_formato_imp_ped_pro"))
		p_e_formato_imp_dev_pro=limpiaCadena(request.form("e_formato_imp_dev_pro"))
		p_e_formato_imp_ord_fab=limpiaCadena(request.form("e_formato_imp_ord_fab"))
		p_e_formato_imp_ord=limpiaCadena(request.form("e_formato_imp_ord"))
		p_e_formato_imp_mov=limpiaCadena(request.form("e_formato_imp_mov"))
		p_e_formato_imp_mov=limpiaCadena(request.form("e_formato_imp_ped_ti"))
		p_e_formato_imp_parte=limpiaCadena(request.form("e_formato_imp_parte"))
		p_e_formato_imp_cat=limpiaCadena(request.form("e_formato_imp_cat"))
		p_e_formato_imp_inc=limpiaCadena(request.form("e_formato_imp_inc"))
        p_e_ocultarmkp=limpiaCadena(request.Form("e_ocultarmkp"))
        
		p_criterio=limpiaCadena(request("criterio"))
		p_campo=limpiaCadena(request("campo"))
		p_texto=limpiaCadena(request("texto"))
		p_npagina=limpiaCadena(request("npagina"))
		p_Nserie=limpiaCadena(request("Nserie"))
		p_pagina=limpiaCadena(request("pagina"))
		p_p_codigo=limpiaCadena(request("p_codigo"))
		checkCadena p_p_codigo

	if p_i_Nserie>"" and p_i_Nombre>"" then
		p_codigo=session("ncliente") & p_i_Nserie
		p_Nombre=p_i_Nombre
		p_pordefecto=nz_b(p_i_pordefecto)
		if p_pordefecto<>0 then p_pordefecto=1
		p_facturable=nz_b(p_i_facturable)
		if p_facturable<>0 then p_facturable=1
		
		p_ocultarmkp=nz_b(p_i_ocultarmkp)
		if p_ocultarmkp<>0 then p_ocultarmkp=1

		'''se grabara los datos segun el tipo de documento
		select case p_i_Documento
			case "ALBARAN DE SALIDA"
				p_formato_imp=p_i_formato_imp_alb_cli
			case "DEVOLUCION DE CLIENTE"
				p_formato_imp=p_i_formato_imp_dev_cli
			case "FACTURA A CLIENTE"
				p_formato_imp=p_i_formato_imp_fac_cli
				p_facturable=1
			case "PEDIDO DE CLIENTE"
				p_formato_imp=p_i_formato_imp_ped_cli
			case "PRESUPUESTO A CLIENTE"
				p_formato_imp=p_i_formato_imp_pre_cli
			case "FACTURA DE PROVEEDOR"
				p_formato_imp=p_i_formato_imp_fac_pro
				p_facturable=1
			case "ALBARAN DE PROVEEDOR"
				p_formato_imp=p_i_formato_imp_alb_pro
			case "PEDIDO A PROVEEDOR"
				p_formato_imp=p_i_formato_imp_ped_pro
			case "DEVOLUCION A PROVEEDOR"
				p_formato_imp=p_i_formato_imp_dev_pro
			case "ORDEN DE FABRICACION"
				p_formato_imp=p_i_formato_imp_ord_fab
			case "ORDEN"
				p_formato_imp=p_i_formato_imp_ord
			case "MOVIMIENTOS ENTRE ALMACENES"
				p_formato_imp=p_i_formato_imp_mov
			case "PEDIDOS ENTRE ALMACENES"
				p_formato_imp=p_i_formato_imp_ped_ti				
			case "PARTE DE TRABAJO"
				p_formato_imp=p_i_formato_imp_parte
			case "CATALOGO"
				p_formato_imp=p_i_formato_imp_cat
			case "INCIDENCIA"
				p_formato_imp=p_i_formato_imp_inc
			case else:
				p_formato_imp=""
		end select

'''''''
		rst.Open "select * from series where nserie='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		if rst.EOF then
        	rst.Close

            strselect = "select * from series with(ROWLOCK) WHERE nserie=?"
            set conn = Server.CreateObject("ADODB.Connection")
            set rst = Server.CreateObject("ADODB.Recordset")
	        set command = Server.CreateObject("ADODB.Command")
	        conn.open session("dsn_cliente")
            'conn.cursorlocation = 3
	        command.ActiveConnection = conn
	        command.CommandTimeout = 60
	        command.CommandText = strselect
	        command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@p_codigo", adVarChar, adParamInput, 10, p_codigo)
            rst.CursorLocation = adUseClient
            rst.Open command, , adOpenKeyset, adLockOptimistic

			rst.AddNew
			rst("nserie")=p_codigo
			rst("nombre")=p_Nombre
			rst("empresa")=p_i_Empresa
			rst("contador")=p_i_Contador
			rst("nfolio1")=p_i_nfolio1
			rst("nfolio2")=p_i_nfolio2
			rst("ultima_fecha")=p_i_FechaModificacion
			rst("tipo_documento")=p_i_Documento
			'rst("tienda")=nulear(p_e_ccostes)'**rgu 4/9/2009
            if isnull(p_e_ccostes) or p_e_ccostes & ""="" then
		        'rst("tienda")=null
            else
                rst("tienda")=p_e_ccostes
            end if

			rst("almacen")=nulear(p_e_almacen)
			if p_i_cta_ventas>"" then rst("cta_ventas")=p_i_cta_ventas
			if p_i_cta_compras>"" then rst("cta_compras")=p_i_cta_compras
			if p_i_cta_aventas>"" then rst("cta_aventas")=p_i_cta_aventas
			if p_i_cta_acompras>"" then rst("cta_acompras")=p_i_cta_acompras
			if p_i_cta_caja>"" then rst("cta_caja")=p_i_cta_caja
			if p_i_cta_rfventas>"" then rst("cta_rfventas")=p_i_cta_rfventas
			if p_i_cta_rfcompras>"" then rst("cta_rfcompras")=p_i_cta_rfcompras
			if p_i_cta_retventas>"" then rst("cta_retventas")=p_i_cta_retventas
			if p_i_cta_retcompras>"" then rst("cta_retcompras")=p_i_cta_retcompras
			rst("pordefecto")=p_pordefecto
			if p_pordefecto<>0 and d_lookup("nserie","series","nserie like '" & session("ncliente") & "%' and pordefecto=1 and tipo_documento='" & p_Documento & "'",session("dsn_cliente"))>"" then rst("pordefecto")=0
			rst("facturable")=p_facturable
			''if p_facturable<>0 and d_lookup("nserie","series","nserie like '" & session("ncliente") & "%' and facturable=1 and tipo_documento='" & p_Documento & "'",session("dsn_cliente"))>"" then rst("facturable")=0
			if p_i_cliente>"" then
				rst("cliente")=p_i_cliente
			else
				'rst("cliente")=null
			end if
			if p_formato_imp<>"" then
				rst("formato_imp")=p_formato_imp
			else
				'rst("formato_imp")=null
			end if
			rst("totalizar") = 0
			rst("EDI")=null_z(d_max("edi","series","nserie like '" & session("ncliente") & "%'",session("dsn_cliente")))+1
			rst("ocultarmkp")=p_ocultarmkp
			rst.Update

            rst.close
            conn.close
            set rst = nothing
            set command = nothing
            set conn = nothing

		else %>
		    rst.Close
			<script>
			    window.alert("<%=LitMsgCodigoExiste%>");
			    history.back();
	      	</script>
	   	<%end if
	end if


	'actualizamos valores'
	if p_e_Nserie>"" or p_e_Nombre>"" then
		p_codigoAnt = p_h_codigo
		p_codigo=session("ncliente") & p_e_Nserie
		p_Nombre=p_e_Nombre
		p_pordefecto=nz_b(p_e_pordefecto)
		if p_pordefecto<>0 then p_pordefecto=1
		p_facturable=nz_b(p_e_facturable)
		if p_facturable<>0 then p_facturable=1
        '26/02/2010 Albert: Añadimos valores para editar campo ocultar mkp.
        p_ocultarmkp=nz_b(p_e_ocultarmkp)
		if p_ocultarmkp<>0 then p_ocultarmkp=1

		'se grabara los datos segun el tipo de documento'
		select case p_e_Documento
			case "ALBARAN DE SALIDA"
				p_formato_imp=p_e_formato_imp_alb_cli
			case "DEVOLUCION DE CLIENTE"
				p_formato_imp=p_e_formato_imp_dev_cli
			case "FACTURA A CLIENTE"
				p_formato_imp=p_e_formato_imp_fac_cli
				p_facturable=1
			case "PEDIDO DE CLIENTE"
				p_formato_imp=p_e_formato_imp_ped_cli
			case "PRESUPUESTO A CLIENTE"
				p_formato_imp=p_e_formato_imp_pre_cli
			case "FACTURA DE PROVEEDOR"
				p_formato_imp=p_e_formato_imp_fac_pro
				p_facturable=1
			case "ALBARAN DE PROVEEDOR"
				p_formato_imp=p_e_formato_imp_alb_pro
			case "PEDIDO A PROVEEDOR"
				p_formato_imp=p_e_formato_imp_ped_pro
			case "DEVOLUCION A PROVEEDOR"
				p_formato_imp=p_e_formato_imp_dev_pro
			case "ORDEN DE FABRICACION"
				p_formato_imp=p_e_formato_imp_ord_fab
			case "ORDEN"
				p_formato_imp=p_e_formato_imp_ord
			case "MOVIMIENTOS ENTRE ALMACENES"
				p_formato_imp=p_e_formato_imp_mov
			case "PEDIDOS ENTRE ALMACENES"
				p_formato_imp=p_e_formato_imp_ped_ti				
			case "PARTE DE TRABAJO"
				p_formato_imp=p_e_formato_imp_parte
			case "CATALOGO"
				p_formato_imp=p_e_formato_imp_cat
			case "INCIDENCIA"
				p_formato_imp=p_e_formato_imp_inc
			case else:
				p_formato_imp=""
		end select

		'''''''
		if p_codigo<>p_codigoAnt then
		  	rst.Open "select * from series with(nolock) where nserie='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		  	if not rst.EOF then
				rst.close
				'ya existe el nuevo codigo que se quiere asignar a esta serie %>
				<SCRIPT language="JavaScript">
				    window.alert("<%=LitMsgCodigoExiste%>")
				    document.location = "series.asp"
				</script><%
			else
				rst.close
                resultado=0
                set conn = Server.CreateObject("ADODB.Connection")
                set command =  Server.CreateObject("ADODB.Command")

                conn.open session("dsn_cliente")
                command.ActiveConnection =conn
                command.CommandTimeout = 0
                command.CommandText="DeleteSeries"
                command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                command.Parameters.Append command.CreateParameter("@nserie",adVarChar,adParamInput,10,p_codigoAnt)
                command.Parameters.Append command.CreateParameter("@result",adInteger,adParamOutput)
                command.Execute,,adExecuteNoRecords
                resultado=command.Parameters("@result").Value
                if resultado & ""="" then
                    resultado=0
                end if
                conn.close
                set command=nothing
                set conn=nothing
				if resultado<>1 then
					'existen documentos con el codigo anterior de la forma de pago'
					%>
		 			<script type="text/javascript" language="javascript">
		 			    window.alert("<%=LitMsgModifSerie%>")
		 			    document.location = "series.asp"
					</script>
                    <%
				else
				 	rst.Open "select * from series where nserie='" & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
					rst.AddNew
					rst("nserie")=p_codigo
					rst("nombre")=p_Nombre
					rst("empresa")=p_e_Empresa
			 		rst("contador")=p_e_Contador
			 		rst("nfolio1")=p_e_nfolio1
			 		rst("nfolio2")=p_e_nfolio2
					rst("ultima_fecha")=p_e_FechaModificacion
			 		rst("tipo_documento")=p_e_Documento
			 		rst("tienda")=nulear(p_e_ccostes)		'**rgu 4/9/2009
			 		rst("almacen")=nulear(p_e_almacen)
					rst("cta_ventas")=nulear(p_e_cta_ventas)
					rst("cta_compras")=nulear(p_e_cta_compras)
					rst("cta_aventas")=nulear(p_e_cta_aventas)
					rst("cta_acompras")=nulear(p_e_cta_acompras)
					rst("cta_caja")=nulear(p_e_cta_caja)
					rst("cta_rfventas")=nulear(p_e_cta_rfventas)
					rst("cta_rfcompras")=nulear(p_e_cta_rfcompras)
					rst("cta_retventas")=nulear(p_e_cta_retventas)
					rst("cta_retcompras")=nulear(p_e_cta_retcompras)
					rst("pordefecto")=p_pordefecto
					if p_pordefecto<>0 and d_lookup("nserie","series","nserie like '" & session("ncliente") & "%' and pordefecto=1 and tipo_documento='" & p_Documento & "'",session("dsn_cliente"))>"" then rst("pordefecto")=0
					rst("facturable")=p_facturable
					''if p_facturable<>0 and d_lookup("nserie","series","nserie like '" & session("ncliente") & "%' and facturable=1 and tipo_documento='" & p_Documento & "'",session("dsn_cliente"))>"" then rst("facturable")=0

					if p_e_cliente>"" then
						rst("cliente")=p_e_cliente
					else
						rst("cliente")=null
					end if
					if p_formato_imp<>"" then
						rst("formato_imp")=p_formato_imp
					else
						rst("formato_imp")=null
					end if
					rst("totalizar") = 0
					rst("EDI")=null_z(d_max("edi","series","nserie like '" & session("ncliente") & "%'",session("dsn_cliente")))+1
					rst("ocultarmkp")=p_ocultarmkp '26/02/2010 Albert: Asignacion para grabar check ocultarmkp en bd.
	        		rst.Update
					rst.close
				end if
			end if
		else ' los codigos son iguales
		  	rst.Open "select * from series with(rowlock) where nserie='" & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
			rst("nserie")  = p_codigo
			rst("nombre")=p_Nombre
			rst("empresa")=p_e_Empresa
			rst("contador")=p_e_Contador  
			rst("nfolio1")=p_e_nfolio1
			rst("nfolio2")=p_e_nfolio2
			rst("ultima_fecha")=p_e_FechaModificacion
			rst("tipo_documento")=p_e_Documento
			rst("almacen")=nulear(p_e_almacen)
			rst("tienda")=nulear(p_e_ccostes) '**rgu 4/9/2009
			rst("cta_ventas")=nulear(p_e_cta_ventas)
			rst("cta_compras")=nulear(p_e_cta_compras)
			rst("cta_aventas")=nulear(p_e_cta_aventas)
			rst("cta_acompras")=nulear(p_e_cta_acompras)
			rst("cta_caja")=nulear(p_e_cta_caja)
			rst("cta_rfventas")=nulear(p_e_cta_rfventas)
			rst("cta_rfcompras")=nulear(p_e_cta_rfcompras)
			rst("cta_retventas")=nulear(p_e_cta_retventas)
			rst("cta_retcompras")=nulear(p_e_cta_retcompras)
			rst("pordefecto")=p_pordefecto
			if p_pordefecto<>0 and d_lookup("nserie","series","nserie like '" & session("ncliente") & "%' and pordefecto=1 and tipo_documento='" & p_Documento & "' and nserie<>'" & p_codigo & "'",session("dsn_cliente"))>"" then rst("pordefecto")=0
			rst("facturable")=p_facturable

			if p_e_cliente>"" then
				rst("cliente")=p_e_cliente
			else
				rst("cliente")=null
			end if
			if p_formato_imp<>"" then
				rst("formato_imp")=p_formato_imp
			else
				rst("formato_imp")=null
			end if
			rst("totalizar") = 0
			rst("ocultarmkp")=p_ocultarmkp '26/02/2010 Albert: Asignacion para grabar valor ocultarmkp en bd
			rst.Update
			rst.close
		end if
	end if

  'eliminamos valores
  if mode="delete" and p_Nserie>"" then
    p_codigo=session("ncliente") & p_Nserie
    resultado=0
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")

    conn.open session("dsn_cliente")
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText="DeleteSeries"
    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
    command.Parameters.Append command.CreateParameter("@nserie",adVarChar,adParamInput,10,p_codigo)
    command.Parameters.Append command.CreateParameter("@result",adInteger,adParamOutput)
    command.Execute,,adExecuteNoRecords
    resultado=command.Parameters("@result").Value
    if resultado & ""="" then
        resultado=0
    end if
    conn.close
    set command=nothing
    set conn=nothing
    if resultado<>1 then
        %>
	 	<script type="text/javascript" language="javascript">
	 	    window.alert("<%=LitMsgBorrarSerie%>")
	 	    document.location = "series.asp"
		</script>
        <%
    else
        %>
	 	<script type="text/javascript" language="javascript">
	 	    window.alert("<%=LITSERIEBORRACORR%>")
		</script>
        <%
    end if

  end if
  ''response.write("los datos son-" & p_campo & "-" & p_texto & "-" & p_criterio & "-<br>")
        p_campo2=""
       aux = p_campo
       if p_texto>"" then
            if p_campo="tipo_documento" then
                aux = "tipo_documento"
                p_campo2 = "(s.tipo_documento"
            else
                if mid(p_campo,1,len("s."))<>"s." then
                    p_campo2 = "s." + p_campo
                end if
			end if
         c_where=" where nserie like '" & session("ncliente") & "%' and " + p_campo2 + " "
      else
         c_where=" where nserie like '" & session("ncliente") & "%' "
      end if
  
      if p_texto>"" then
         select case p_criterio
            case "contiene"
               c_where=c_where+ "like '%"+p_texto+"%'"
               if aux="tipo_documento" then                    
                    c_where=c_where + " OR ldt.value LIKE " +  "'%" + p_texto + "%') "
                end if
            case "termina"
               c_where=c_where+ "like '%"+p_texto+"'"
               if aux="tipo_documento" then                    
                    c_where=c_where + " OR ldt.value LIKE " +  "'%" + p_texto + "') "
                end if
            case "empieza"
               c_where=c_where+ "like '" + p_texto + "%'"
               if aux="tipo_documento" then                    
                    c_where=c_where + " OR ldt.value LIKE " +  "'" + p_texto + "%') "
                end if
            case "igual"
                if aux = "nserie" then
                    c_where=c_where + "='" + session("ncliente") + p_texto + "' "  
                end if            
                if aux="tipo_documento" then                    
                    c_where=c_where + "='" + p_texto + "' OR ldt.value = '" + p_texto + "') "
                end if   
                if aux = "nombre" then
                    c_where=c_where + "='" + p_texto + "' " 
                end if
            end select
      end if

	Alarma "series.asp" %>
	<hr>
	<%
        'si_tiene_modulo_ccostes=1
        'si_tiene_modulo_ecomerce=1
        ''ricardo 7-12-2004 cuando editamos , guardamos e intentamos guardar una serie nueva da error de javascript
        'AMP 28/07/2010: Añadimos opcion asistente puesta en marcha con parametro viene=asistente.
        if mode<>"edit" and mode<>"search" then
            if viene="asistente" then
	            %><script language="javascript">	          parent.botones.document.location = "series_bt.asp?mode=browse&viene=asistente";</script><%
	        else
	            %><script language="javascript">	          parent.botones.document.location = "series_bt.asp?mode=browse";</script><%
	        end if
        end if
	    set conn = Server.CreateObject("ADODB.Connection")

        initial_catalogC=encontrar_datos_dsn(session("dsn_cliente"),"Initial Catalog=")

		donde=inStr(1,DSNImport,"Initial Catalog=",1)
		donde_fin=InStr(donde,DSNImport,";",1)
		if donde_fin=0 then
			donde_fin=len(DSNImport)
		end if
		cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))

		dsnCliente=cadena_dsn_final
		conn.open dsnCliente

        c_select = "SELECT s.*, ISNULL(ldt.value,s.tipo_documento) as descripcion FROM series s with(NOLOCK) LEFT JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON s.tipo_documento = td.tippdoc LEFT JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser
	    c_order = " ORDER BY s.nserie"

        if c_where>"" then
           c_select=c_select+c_where+c_order
        end if
        if p_npagina="" then
           p_npagina=1
        end if

        select case p_pagina
           case "siguiente"
              p_npagina=p_npagina+1
           case "anterior"
              p_npagina=p_npagina-1
        end select
  		%>
  		<input type="hidden" name="h_npagina" value="<%=enc.EncodeForHtmlAttribute(null_s(cstr(p_npagina)))%>">
		<%
        
        set rst = Server.CreateObject("ADODB.Recordset")

        'response.Write(c_select)
        'response.end
        rst.Open c_select,dsnCliente,adUseClient, adLockReadOnly

        if not rst.EOF then
           rst.PageSize=NumReg
           rst.AbsolutePage=p_npagina
        end if

  if mode<>"edit" and rst.RecordCount>NumReg then
      if p_npagina >1 then %>
	  		<a class=CABECERA href="series.asp?pagina=anterior&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
		<IMG SRC="<%=themeIlion %><%=ImgAnterior%>" align='top' ALT="<%=LitAnterior%>"></a>
  	<%end if
      texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	  <font class=CELDA> <%=texto%> </font> <%
      if clng(p_npagina) < rst.PageCount then %>
	  		<a class=CABECERA href="series.asp?pagina=siguiente&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
		<IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" align='top' ALT="<%=LitSiguiente%>"></a>
  	<%end if

	%><font class=CELDA>&nbsp;&nbsp; Ir a Pag. <input class=CELDA type="text" name="SaltoPagina1" size="2">&nbsp;&nbsp;<a class=CELDAREF href="javascript:IrAPagina(1,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina');">Ir</a></font><%

 end if%>
 <table class="width100 lg-table-responsive bCollapse" BORDER="0" CELLSPACING="1" CELLPADDING="1">
      <%Drawfila color_fondo
        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitSerie & "</b>"
        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitNombre & "</b>"
        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitEmpresa & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "right", true,"<b>" & LitContador & "</b>"
		if si_tiene_gestionFolios then
		    DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LIT_DESDE_NFOLIO & "</b>"
		    DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LIT_HASTA_NFOLIO & "</b>"
		end if
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "right", true,"<b>" & LitFechaModificacion & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width10'","", "left", true,"<b>" & LitDocumento & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitSeriePorDefecto & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width10'","", "left", true,"<b>" & LitClientePorDefecto & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitFormatoImpresionDefecto & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitFacturable & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitAlmacen & "</b>"
		if si_tiene_modulo_ccostes <> 0 then
		    DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCCostes & "</b>"
		end if		
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaVentas & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaCompras & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaAVentas & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaACompras & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaCaja & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRFVentas & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRFCompras & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRETVentas & "</b>"
		DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRETCompras & "</b>"
		'-- 26/02/2010 Albert: Ocultarmkp Añadimos literal definicio en series.ini
		if si_tiene_modulo_ecomerce <> 0 then
		   DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitOcultarMKP & "</b>"
		end if		
		
        par=false
        i=1

        while not rst.EOF and i<=NumReg
           if par then
              Drawfila color_terra
              par=false
           else
              Drawfila color_blau
              par=true
           end if
           
           
           if mode="edit" and p_p_codigo=rst("nserie") then
		      set rs_Docs = server.CreateObject("ADODB.Recordset")
              %><a id="<%=enc.EncodeForHtmlAttribute(null_s(rst("nserie")))%>"></a><%
			  %><input type="hidden" name="h_codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nserie")))%>"/><%

			   'DrawInputCelda "CELDA7 maxlength='5'","","","5",0,"","e_Nserie",trimCodEmpresa(rst("nserie"))
                DrawceldaDet "'CELDAL7 width5'","", "left", false, trimCodEmpresa(rst("nserie"))

                %>
                <input type="hidden" name="e_Nserie" value="<%=trimCodEmpresa(rst("nserie"))%>"/>
                <%

                %><td class="CELDAL7 width5"><%
                    'DrawInputCelda "CELDA7 maxlength='50'","","","25",0,"","e_Nombre",rst("nombre")
                    %><input type="text" name="e_Nombre" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nombre")))%>"/>
                 </td><%
				

			   'DrawInputCelda "CELDA7","","","25",0,"","e_empresa",d_lookup("nombre","empresas","cif='" & rst("empresa") & "'",session("dsn_cliente"))
		   	    rs_Docs.Open "SELECT * FROM empresas with(nolock) where cif like '" & session("ncliente") & "%' order by nombre",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
			    'DrawSelectCelda "CELDA7","125","",0,"","e_empresa",rs_Docs,rst("empresa"),"cif","nombre","",""
                %><td class="CELDAL7 width5"><%                
                DrawSelect "'width100'", "","e_empresa",rs_Docs,enc.EncodeForHtmlAttribute(null_s(rst("empresa"))),"cif","nombre","",""
			    rs_Docs.Close
                %></td>
                    <td class="CELDAL7 width5">
                        <input class="width100" type="text" name="e_Contador" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("contador")))%>"/>
                    </td><%

                if si_tiene_gestionFolios then
                
                    if rst("tipo_documento")="FACTURA A CLIENTE" then
                        visibleFolio = "style='width:45px;'"
                    else
                        visibleFolio = "style='display:none;width:45px;'"
                    end if
                    
                    %><td class="CELDAL7 width5" ><input class="width100" <%=visibleFolio %> type="text" name="e_nfolio1" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nfolio1")))%>"></td><%
                    %><td class="CELDAL7 width5" ><input class="width100" <%=visibleFolio %> type="text" name="e_nfolio2" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nfolio2")))%>"></td><%

                end if

                
                DrawCeldaDet "'CELDAL7 width5'","", "right", false, enc.EncodeForHtmlAttribute(null_s(rst("ultima_fecha")))

                strselectTD = "SELECT * FROM Tipo_Documentos with(nolock) where tippdoc not like '%vencimiento%' and tippdoc not like '%e-mail%' and tippdoc not like '%carta%' order by tippdoc"
                if si_tiene_gestionFolios then
                    mostrarFolios = "mostrarNfolios(1);"
                else
                    mostrarFolios = ""
                end if
                
                rstAux.open strselectTD,DsnIlion,adOpenKeyset, adLockOptimistic                
                Dim listauxE
                While Not rstAux.EOF
                    listauxE = listauxE & "'" & rstAux("tippdoc") & "', "
                    rstAux.MoveNext
                wend
                rstAux.Close                
                listaInE = Left(listauxE,len(listauxE)-2)

                set conn = Server.CreateObject("ADODB.Connection")        
                set command =  Server.CreateObject("ADODB.Command")
                conn.open DSNIlion
                command.ActiveConnection = conn
                command.CommandTimeout = 0
                command.CommandText = "ComboBoxDocTypes"
                command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                command.NamedParameters = True 
                command.Parameters.Append command.CreateParameter("@inlist",adVarChar,adParamInput,len(listaInE),listaInE)
                command.Parameters.Append command.CreateParameter("@outlist",adVarChar,adParamInput,1,"")
                command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,session("usuario"))
                command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))

                set rstTD = Server.CreateObject("ADODB.Recordset")
                set rstTD = command.Execute

                %><td class="CELDAL7 width10"><%  
                if not rstTD.eof then
	                'DrawSelectCelda "CELDA7","125","",0,"","e_documento",rstTD,rst("tipo_documento"),"tippdoc","descripcion","onchange","comprobar_tipodoc('1');" & mostrarFolios &""
                     DrawSelect "'width100'", "","e_documento",rstTD,rst("tipo_documento"),"descripcion","tippdoc","onchange","comprobar_tipodoc('1');" & mostrarFolios &""
                end if			
                rstTD.close
                 %></td><%

                set command=nothing
                %><input type="hidden" name="e_doc_aux" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("tipo_documento")))%>"><%

                %><td class="CELDAC7 width5"><%  
                     DrawCheck "","","e_pordefecto",rst("pordefecto")
                %></td><%

			  %><td class="CELDAL7 width10">
				<%if (rst("tipo_documento")="ALBARAN DE SALIDA" or _
					rst("tipo_documento")="DEVOLUCION DE CLIENTE" or _
					rst("tipo_documento")="FACTURA A CLIENTE" or _
					rst("tipo_documento")="HOJA DE GASTOS" or _
					rst("tipo_documento")="MOVIMIENTOS ENTRE ALMACENES" or _
					rst("tipo_documento")="PEDIDOS ENTRE ALMACENES" or _
					rst("tipo_documento")="PEDIDO DE CLIENTE" or _
					rst("tipo_documento")="PRESUPUESTO A CLIENTE" or _
					rst("tipo_documento")="TICKET") then%>
					<input CLASS="width50" type="input" name="nombre_cli" value="<%=d_lookup("rsocial","clientes","ncliente='" & rst("cliente") & "'",session("dsn_cliente"))%>" disabled>
				<%else%>                
					<input CLASS="width50" type="input" name="nombre_cli" value="<%=d_lookup("razon_social","proveedores","nproveedor='" & rst("cliente") & "'",session("dsn_cliente"))%>" disabled>
				<%end if%>
				<a CLASS=CELDAREFB  href="javascript:buscar_clipro('1')" OnMouseOver="self.status='<%=LitBuscarCliPro%>'; return true;" OnMouseOut="self.status=''; return true;"><IMG SRC="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> ALT="<%=LitBuscar%>"></a>
				<a CLASS=CELDAREFB  href="javascript:limpiarCliente();" OnMouseOver="self.status=''; return true;" OnMouseOut="self.status=''; return true;"><img align="center" src="<%=themeIlion %><%=ImgEliminarDet%>" <%=ParamImgVaciarCampo%> ALT="<%=LitBorrar%>"></a>
				<input type="hidden" name="h_cliente" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cliente")))%>">
		    	</td>
	
       		<td class="CELDAL7 width5">
					<%
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ALBARAN DE SALIDA' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                    'DrawSelectCelda "CELDA","95","' id='formatos_imp_alb_cli' style='display:" & iif(rst("tipo_documento")="ALBARAN DE SALIDA","","none") & ";' height='10",0,"","e_formato_imp_alb_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="ALBARAN DE SALIDA","","none") & ";' id='formatos_imp_alb_cli' ","e_formato_imp_alb_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
                    
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='DEVOLUCION DE CLIENTE' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_dev_cli' style='display:" & iif(rst("tipo_documento")="DEVOLUCION DE CLIENTE","","none") & ";' height='10",0,"","e_formato_imp_dev_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="DEVOLUCION DE CLIENTE","","none") & ";' id='formatos_imp_dev_cli' ","e_formato_imp_dev_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
					
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='FACTURA A CLIENTE' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_fac_cli' style='display:" & iif(rst("tipo_documento")="FACTURA A CLIENTE","","none") & ";' height='10",0,"","e_formato_imp_fac_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="FACTURA A CLIENTE","","none") & ";' id='formatos_imp_fac_cli' ","e_formato_imp_fac_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDO DE CLIENTE' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ped_cli' style='display:" & iif(rst("tipo_documento")="PEDIDO DE CLIENTE","","none") & ";' height='10",0,"","e_formato_imp_ped_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="PEDIDO DE CLIENTE","","none") & ";' id='formatos_imp_ped_cli' ","e_formato_imp_ped_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PRESUPUESTO A CLIENTE' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_pre_cli' style='display:" & iif(rst("tipo_documento")="PRESUPUESTO A CLIENTE","","none") & ";' height='10",0,"","e_formato_imp_pre_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="PRESUPUESTO A CLIENTE","","none") & ";' id='formatos_imp_pre_cli' ","e_formato_imp_pre_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='MOVIMIENTOS ENTRE ALMACENES' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_mov' style='display:" & iif(rst("tipo_documento")="MOVIMIENTOS ENTRE ALMACENES","","none") & ";' height='10",0,"","e_formato_imp_mov",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="MOVIMIENTOS ENTRE ALMACENES","","none") & ";' id='formatos_imp_mov' ","e_formato_imp_mov",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDOS ENTRE ALMACENES' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ped_ti' style='display:" & iif(rst("tipo_documento")="PEDIDOS ENTRE ALMACENES","","none") & ";' height='10",0,"","e_formato_imp_ped_ti",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="PEDIDOS ENTRE ALMACENES","","none") & ";' id='formatos_imp_ped_ti' ","e_formato_imp_ped_ti",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='FACTURA DE PROVEEDOR' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_fac_pro' style='display:" & iif(rst("tipo_documento")="FACTURA DE PROVEEDOR","","none") & ";' height='10",0,"","e_formato_imp_fac_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="FACTURA DE PROVEEDOR","","none") & ";' id='formatos_imp_fac_pro' ","e_formato_imp_fac_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ALBARAN DE PROVEEDOR' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                        
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_alb_pro' style='display:" & iif(rst("tipo_documento")="ALBARAN DE PROVEEDOR","","none") & ";' height='10",0,"","e_formato_imp_alb_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="ALBARAN DE PROVEEDOR","","none") & ";' id='formatos_imp_alb_pro' ","e_formato_imp_alb_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDO A PROVEEDOR' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                        
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ped_pro' style='display:" & iif(rst("tipo_documento")="PEDIDO A PROVEEDOR","","none") & ";' height='10",0,"","e_formato_imp_ped_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="PEDIDO A PROVEEDOR","","none") & ";' id='formatos_imp_ped_pro' ","e_formato_imp_ped_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='DEVOLUCION A PROVEEDOR' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_dev_pro' style='display:" & iif(rst("tipo_documento")="DEVOLUCION A PROVEEDOR","","none") & ";' height='10",0,"","e_formato_imp_dev_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="DEVOLUCION A PROVEEDOR","","none") & ";' id='formatos_imp_dev_pro' ","e_formato_imp_dev_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='CATALOGO' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					

                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_cat' style='display:" & iif(rst("tipo_documento")="CATALOGO","","none") & ";' height='10",0,"","e_formato_imp_cat",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="CATALOGO","","none") & ";' id='formatos_imp_cat' ","e_formato_imp_cat",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""


                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ORDEN DE FABRICACION' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                        
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ord_fab' style='display:" & iif(rst("tipo_documento")="ORDEN DE FABRICACION","","none") & ";' height='10",0,"","e_formato_imp_ord_fab",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="ORDEN DE FABRICACION","","none") & ";' id='formatos_imp_ord_fab' ","e_formato_imp_ord_fab",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ORDEN' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					

                    'DrawSelectCelda "CELDA","95","' id='formatos_imp_ord' style='display:" & iif(rst("tipo_documento")="ORDEN","","none") & ";' height='10",0,"","e_formato_imp_ord",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="ORDEN","","none") & ";' id='formatos_imp_ord' ","e_formato_imp_ord",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                   
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='INCIDENCIA' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_inc' style='display:" & iif(rst("tipo_documento")="INCIDENCIA","","none") & ";' height='10",0,"","e_formato_imp_inc",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="INCIDENCIA","","none") & ";' id='formatos_imp_inc' ","e_formato_imp_inc",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PARTE DE TRABAJO' order by descripcion"
					defecto=rst("FORMATO_IMP")
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                    
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_parte' style='display:" & iif(rst("tipo_documento")="PARTE DE TRABAJO","","none") & ";' height='10",0,"","e_formato_imp_parte",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:" & iif(rst("tipo_documento")="PARTE DE TRABAJO","","none") & ";' id='formatos_imp_parte' ","e_formato_imp_parte",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
   
                    rstAux.close
					%><div id='formatos_imp_ninguno' class="CELDA" style="display:none">
					</div>
			</td><%

            %><td class="CELDAC7 width5"><%  
            DrawCheck "","","e_facturable",rst("facturable")
            %></td><%  

				
			  if (rst("tipo_documento")="FACTURA A CLIENTE" or rst("tipo_documento")="FACTURA DE PROVEEDOR") then
			  %><script>
			        deshabilitaCampo();
			  	</script>
			  <%
			  end if
            %><td class="CELDAL7 width5"><%  
			    rstAux.Open "select codigo, descripcion from almacenes with(nolock) where codigo like '"&session("ncliente")&"%' and isnull(fbaja,'')='' order by descripcion",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
                'DrawSelectCelda "CELDA7","125","",0,"","e_almacen",rstAux,rst("almacen"),"codigo","descripcion","",""
                 DrawSelect "'CELDAL7 width100'","","e_almacen",rstAux,enc.EncodeForHtmlAttribute(null_s(rst("almacen"))),"codigo","descripcion","",""
                rstAux.Close
            %></td><%  
			  '**rgu 4/9/2009
            %><td class="CELDAL7 width5"><%  
		      if si_tiene_modulo_ccostes <> 0 then
		        strselect="select codigo, descripcion from tiendas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion"
		        rstAux.CursorLocation=3
		        rstAux.Open strselect, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
				'DrawSelectCelda "CELDA","100","",0,"","e_ccostes",rstAux,rst("tienda"),"codigo","descripcion","",""
                 DrawSelect "'CELDAL7 width100'","","e_ccostes",rstAux,enc.EncodeForHtmlAttribute(null_s(rst("tienda"))),"codigo","descripcion","",""
				rstAux.close
		      end if
            %></td><%  
                  %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_ventas" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_ventas")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_compras" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_compras")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_aventas" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_aventas")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_acompras" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_acompras")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_caja" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_caja")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_rfventas" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_rfventas")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_rfcompras" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_rfcompras")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_retventas" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_retventas")))%>"/>
                    </td><% 

                    %><td class="CELDAL7 width5">
		                <input type="text" class="width100" ,maxlength="5", name="e_cta_retcompras" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("cta_retcompras")))%>"/>
                    </td><% 
                    if si_tiene_modulo_ecomerce <> 0 then 
                    %><td class="CELDAC7 width5"><%
		                DrawCheck "","","e_ocultarmkp",rst("ocultarmkp")
                    %></td><%
                    end if
			  
			  'DrawInputCelda "CELDA7 maxlength="&rst("cta_ventas").DefinedSize,"","","9",0,"","e_cta_ventas",rst("cta_ventas")
			  'DrawInputCelda "CELDA7 maxlength="&rst("cta_compras").DefinedSize,"","","9",0,"","e_cta_compras",rst("cta_compras")
			  'DrawInputCelda "CELDA7 maxlength="&rst("cta_aventas").DefinedSize,"","","9",0,"","e_cta_aventas",rst("cta_aventas")
			  'DrawInputCelda "CELDA7 maxlength="&rst("cta_acompras").DefinedSize,"","","9",0,"","e_cta_acompras",rst("cta_acompras")
			  'DrawInputCelda "CELDA7 maxlength="&rst("cta_caja").DefinedSize,"","","9",0,"","e_cta_caja",rst("cta_caja")
			  'DrawInputCelda "CELDA7 maxlength="&rst("cta_rfventas").DefinedSize,"","",9,0,"","e_cta_rfventas",rst("cta_rfventas")
		   	  'DrawInputCelda "CELDA7 maxlength="&rst("cta_rfcompras").DefinedSize,"","",9,0,"","e_cta_rfcompras",rst("cta_rfcompras")
			  'DrawInputCelda "CELDA7 maxlength="&rst("cta_retventas").DefinedSize,"","",9,0,"","e_cta_retventas",rst("cta_retventas")
		   	  'DrawInputCelda "CELDA7 maxlength="&rst("cta_retcompras").DefinedSize,"","",9,0,"","e_cta_retcompras",rst("cta_retcompras")
		   	  ' OcultarMKP
		   	  'if si_tiene_modulo_ecomerce <> 0 then ' -- 26/02/2010 Albert: Comprobamos que la empresa tenga modulo mkp
		   	    'DrawCheckCelda "CELDACENTER","","","0","","e_ocultarmkp",rst("ocultarmkp")
		   	  'end if

        ' >>> MCA 16/08/05 : Las series de regularización de inventario no deben aparecer.
		'elseif rst("Tipo_Documento") <> "REGULARIZACION DE INVENTARIO" then
        '***RGU 30/03/06: Las series de Cambio de Precios no deben aparecer
 		elseif rst("Tipo_Documento") <> "REGULARIZACION DE INVENTARIO"  and rst("Tipo_Documento")<>"CAMBIO DE PRECIOS" then          

              h_ref="javascript:Editar('" & reemplazar(rst("nserie")," ","%20") & "'," & _
			                           p_npagina & ",'" & _
									   p_campo & "','" & _
									   p_criterio & "','" & _
									   reemplazar(p_texto," ","%20") & "','" & viene & "');"%>
            <td class="CELDAL7 width5">
                <%DrawHref "CELDAREF","",trimCodEmpresa(rst("nserie")),h_ref%>
            </td><%
              DrawCeldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("nombre")))
              DrawCeldaDet "'CELDAL7 width5'","", "left", false, d_lookup("nombre","empresas","cif='" & rst("empresa") & "'",session("dsn_cliente"))
			  DrawCeldaDet "'CELDAL7 width5'","", "right", false, enc.EncodeForHtmlAttribute(null_s(rst("contador")))
			  if si_tiene_gestionFolios then
			    DrawCeldaDet "'CELDAL7 width5'","", "right", false, enc.EncodeForHtmlAttribute(null_s(rst("nfolio1")))
			    DrawCeldaDet "'CELDAL7 width5'","", "right", false, enc.EncodeForHtmlAttribute(null_s(rst("nfolio2")))
			  end if
			  DrawCeldaDet "'CELDAL7 width5'","","", false, enc.EncodeForHtmlAttribute(null_s(rst("ultima_fecha")))
			  DrawCeldaDet "'CELDAL7 width10'","","", false, enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))
			  if rst("pordefecto")=true then
						DrawCeldaDet "'CELDAC7 width5'","", "center", false, "<IMG SRC='../images/" & ImgSeriePorDefectoSi & "' " & ParamImgSeriePorDefectoSi & ">"
				else
						DrawCeldaDet "'CELDAC7 width5'","", "center", false, "<IMG SRC='../images/" & ImgSeriePorDefectoNo & "' " & ParamImgSeriePorDefectoNo & ">"
			  end if
			  if (rst("tipo_documento")="ALBARAN DE SALIDA" or _
					rst("tipo_documento")="DEVOLUCION DE CLIENTE" or _
					rst("tipo_documento")="FACTURA A CLIENTE" or _
					rst("tipo_documento")="HOJA DE GASTOS" or _
					rst("tipo_documento")="MOVIMIENTOS ENTRE ALMACENES" or _
					rst("tipo_documento")="PEDIDOS ENTRE ALMACENES" or _
					rst("tipo_documento")="PEDIDO DE CLIENTE" or _
					rst("tipo_documento")="PRESUPUESTO A CLIENTE" or _
					rst("tipo_documento")="TICKET") then
					DrawceldaDet "'CELDAL7 width10'","", "left", false, d_lookup("rsocial","clientes","ncliente='" & rst("cliente") & "'",session("dsn_cliente"))
			  else
					DrawceldaDet "'CELDAL7 width10'","", "left", false, d_lookup("razon_social","proveedores","nproveedor='" & rst("cliente") & "'",session("dsn_cliente"))
			  end if
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, d_lookup("descripcion","clientes_formatos_imp","ncliente='" & session("ncliente") & "' and nformato='" & rst("formato_imp") & "'",dsnilion)
			  if rst("facturable")=true then
						DrawCeldaDet "'CELDAC7 width5'","", "center", false, "<IMG SRC='../images/" & ImgSeriePorDefectoSi & "' " & ParamImgSeriePorDefectoSi & ">"
				else
						DrawCeldaDet "'CELDAC7 width5'","", "center", false, "<IMG SRC='../images/" & ImgSeriePorDefectoNo & "' " & ParamImgSeriePorDefectoNo & ">"
			  end if
			  
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, d_lookup("descripcion","almacenes","codigo like '"&session("ncliente")&"%' and codigo='" & rst("almacen") & "'",session("dsn_cliente"))
			  
		      '**rgu 4/9/2009
		      if si_tiene_modulo_ccostes <> 0 then
		        DrawceldaDet "'CELDAL7 width5'","", "left", false, d_lookup("descripcion","tiendas","codigo like '"&session("ncliente")&"%' and codigo='" & rst("tienda") & "'",session("dsn_cliente"))
		      end if		
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_ventas")))
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_compras")))
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_aventas")))
		      DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_acompras")))
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_caja")))
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_rfventas")))
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_rfcompras")))
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_retventas")))
			  DrawceldaDet "'CELDAL7 width5'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("cta_retcompras")))
			  
			  ' OcultarMkp
			  if si_tiene_modulo_ecomerce <> 0 then '26/02/2010 Albert: Asignamos icononos para campo ocultarmkp
			    if rst("ocultarmkp")=true then
						DrawCeldaDet "'CELDAC7 width5'","", "center", false, "<IMG SRC='../images/" & ImgSeriePorDefectoSi & "' " & ParamImgSeriePorDefectoSi & ">"
				else
						DrawCeldaDet "'CELDAC7 width5'","", "center", false, "<IMG SRC='../images/" & ImgSeriePorDefectoNo & "' " & ParamImgSeriePorDefectoNo & ">"
			    end if
			  end if
			  
		    else 'Si no se imprime nada, se alterna el color de la fila
		    if par=true then
		        par = false
		    else
		        par = true 
		    end if
		    	  
           end if
           i = i + 1
           rst.MoveNext
        wend
        'rst.Close %>
   </table>

  <%if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
	  		<a class=CABECERA href="series.asp?pagina=anterior&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
		<IMG SRC="<%=themeIlion %><%=ImgAnterior%>" align='top' ALT="<%=LitAnterior%>"></a>
  	<%end if
      texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	  <font class=CELDA> <%=texto%> </font> <%
      if clng(p_npagina) < rst.PageCount then %>
	  		<a class=CABECERA href="series.asp?pagina=siguiente&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
		<IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" align='top' ALT="<%=LitSiguiente%>"></a>
    <%end if

	 %><font class=CELDA>&nbsp;&nbsp; Ir a Pag. <input class=CELDA type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;<a class=CELDAREF href="javascript:IrAPagina(2,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina');">Ir</a></font><%

    rst.Close

 end if%>

   <%if mode<>"edit" then 
       if session("version")&"" <> "5" then%>
   		    <hr style="min-width:1225px">
   		    <table class="width100 lg-table-responsive" BORDER="0" CELLSPACING="1" CELLPADDING="1" style="min-width:1225px"><%
       else%>
             <hr>
   		    <table class="width100 lg-table-responsive" BORDER="0" CELLSPACING="1" CELLPADDING="1" >   
      <%end if
            DrawceldaDet "'ENCABEZADOL width50'", "", "left", true,"<b>" & LitNBregistro & "</b>"%>
            </table><%
        if session("version")&"" <> "5" then%>
            <table class="width100 underOrange lg-table-responsive bCollapse" style="min-width:1225px"><%
        else%>
           <table class="width100 underOrange lg-table-responsive bCollapse"><%
        end if%>
            <tr class="underOrange"><%
                DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitSerie & "</b>"
                DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitNombre & "</b>"
                DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitEmpresa & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitContador & "</b>"
		        if si_tiene_gestionFolios then
		            DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LIT_DESDE_NFOLIO & "</b>" 
		            DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LIT_HASTA_NFOLIO & "</b>"
		        end if
		        DrawceldaDet "'ENCABEZADOL underOrange width10'","", "left", true,"<b>" & LitDocumento & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitSeriePorDefecto & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width10'","", "left", true,"<b>" & LitClientePorDefecto & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitFormatoImpresionDefecto & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitFacturable & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitAlmacen & "</b>"
		        if si_tiene_modulo_ccostes <> 0 then
		            DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCCostes & "</b>"
		        end if				
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaVentas & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaCompras & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaAVentas & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaACompras & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaCaja & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRFVentas & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRFCompras & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRETVentas & "</b>"
		        DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCtaRETCompras & "</b>"
		        if si_tiene_modulo_ecomerce <> 0 then '26/02/2010 Albert: Añadimos literal ocultarmkp definido en series.ini
		            DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitOcultarMKP & "</b>"
		        end if%>
            </tr><%
            Drawfila color_blau
		    set rs_Docs = server.CreateObject("ADODB.Recordset")

			'DrawInputCelda "CELDA7 maxlength='5'","","",5,0,"","i_Nserie",""
                %><td class="CELDA underOrange width5">
                  <input type="text" class="width100" ,maxlength="5", name="i_Nserie" size="5"/>
              </td> <%

			'DrawInputCelda "CELDA7 maxlength='50'","","",25,0,"","i_Nombre",""
                %><td class="CELDA underOrange width5">
                <input type="text" class="width100" ,maxlength="50", name="i_Nombre" size="25"/>
                </td><%
			rs_Docs.Open "SELECT * FROM empresas with(nolock) where cif like '" & session("ncliente") & "%' order by nombre",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
			
            'DrawSelectCelda "CELDA7","125","",0,"","i_empresa",rs_Docs,"","cif","nombre","",""
                %><td class="CELDA underOrange width5">
                    <select class="CELDAL7 width100" name="i_empresa">
                        <% while not rs_Docs.EOF
                            %><option value="<%=enc.EncodeForHtmlAttribute(null_s(rs_Docs("cif")))%>"><%=enc.EncodeForHtmlAttribute(null_s(rs_Docs("nombre"))) %></option>              			               
	                    <% rs_Docs.MoveNext
                        wend %>
                    </select>     
               </td><%
			
            rs_Docs.Close
			'DrawInputCelda "CELDA7","","",5,0,"","i_Contador",""    
                %><td class="CELDA underOrange width5">
                <input type="text" class="width100" ,maxlength="5", name="i_Contador" size="5"/><%
             %></td><%
		    
		    if si_tiene_gestionFolios then      
                %><td class="CELDA underOrange width5" ><input class="CELDA7" style="width:45px;" type="text" name="i_nfolio1" value=""></td><%
                %><td class="CELDA underOrange width5" ><input class="CELDA7" style="width:45px;" type="text" name="i_nfolio2" value=""></td><%
            end if
		    
			strselectTD = "SELECT * FROM Tipo_Documentos with(nolock) where tippdoc not like '%vencimiento%' and tippdoc not like '%e-mail%' and tippdoc not like '%carta%' order by tippdoc"
			if si_tiene_gestionFolios then
                mostrarFolios = "mostrarNfolios(2);"
            else
                mostrarFolios = ""
            end if

            rstAux.open strselectTD,DsnIlion,adOpenKeyset, adLockOptimistic                
            Dim listaux2
            While Not rstAux.EOF
                listaux2 = listaux2 & "'" & rstAux("tippdoc") & "', "
                rstAux.MoveNext
            wend
            rstAux.Close                
            listaIn2 = Left(listaux2,len(listaux2)-2)

            set conn = Server.CreateObject("ADODB.Connection")        
            set command =  Server.CreateObject("ADODB.Command")
            conn.open DSNIlion
            command.ActiveConnection = conn
            command.CommandTimeout = 0
            command.CommandText = "ComboBoxDocTypes"
            command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
            command.NamedParameters = True 
            command.Parameters.Append command.CreateParameter("@inlist",adVarChar,adParamInput,len(listaIn2),listaIn2)
            command.Parameters.Append command.CreateParameter("@outlist",adVarChar,adParamInput,1,"")
            command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,session("usuario"))
            command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))

            set rstTD = Server.CreateObject("ADODB.Recordset")
            set rstTD = command.Execute
            if not rstTD.eof then
            %><td class="CELDA underOrange width10"><%
                 DrawSelect "'CELDAL7 width100'","","i_documento",rstTD,"FACTURA","tippdoc","descripcion","onchange","comprobar_tipodoc('2');"& mostrarFolios &""
            end if			
            rstTD.close
            set command=nothing
            %></td><%   
            
            %><td class="CELDAC7 underOrange width5"><%
			    %><input type="hidden" name="i_doc_aux" value=""><%
			    DrawCheck "","","i_pordefecto",false
            %></td><%  

			%><td class="CELDA underOrange width10">
				<input CLASS="width50" type="input" name="nombre_cli" value="" disabled><a CLASS=CELDAREFB  href="javascript:buscar_clipro('2')" OnMouseOver="self.status='<%=LitBuscarCliPro%>'; return true;" OnMouseOut="self.status=''; return true;"><IMG SRC="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> ALT="<%=LitBuscar%>"></a><a CLASS=CELDAREFB  href="javascript:limpiarCliente();" OnMouseOver="self.status=''; return true;" OnMouseOut="self.status=''; return true;"><img align="center" src="<%=themeIlion %><%=ImgEliminarDet%>" <%=ParamImgVaciarCampo%> ALT="<%=LitBorrar%>"></a>
				<input type="hidden" name="h_cliente" value="">
			</td><%
		   %><td class="CELDA underOrange width5">
					<%
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ALBARAN DE SALIDA' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic

			       'DrawSelectCelda "CELDA","95","' id='formatos_imp_alb_cli' style='display:none;' height='10",0,"","i_formato_imp_alb_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_alb_cli' height='10","i_formato_imp_alb_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

					rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='DEVOLUCION DE CLIENTE' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                        
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_dev_cli' style='display:none;' height='10",0,"","i_formato_imp_dev_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_dev_cli' height='10","i_formato_imp_dev_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
                    
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='FACTURA A CLIENTE' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
				
                    
                    'DrawSelectCelda "CELDA","95","' id='formatos_imp_fac_cli' style='display:none;' height='10",0,"","i_formato_imp_fac_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_fac_cli' height='10","i_formato_imp_fac_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDO DE CLIENTE' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ped_cli' style='display:none;' height='10",0,"","i_formato_imp_ped_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_ped_cli' height='10","i_formato_imp_ped_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PRESUPUESTO A CLIENTE' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_pre_cli' style='display:none;' height='10",0,"","i_formato_imp_pre_cli",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_pre_cli' height='10","i_formato_imp_pre_cli",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='MOVIMIENTOS ENTRE ALMACENES' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                        
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_mov' style='display:none;' height='10",0,"","i_formato_imp_mov",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_mov' height='10","i_formato_imp_mov",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDOS ENTRE ALMACENES' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                        
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ped_ti' style='display:none;' height='10",0,"","i_formato_imp_ped_ti",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_ped_ti' height='10","i_formato_imp_ped_ti",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                        
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='FACTURA DE PROVEEDOR' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                    'DrawSelectCelda "CELDA","95","' id='formatos_imp_fac_pro' style='display:none;' height='10",0,"","i_formato_imp_fac_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_fac_pro' height='10","i_formato_imp_fac_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ALBARAN DE PROVEEDOR' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_alb_pro' style='display:none;' height='10",0,"","i_formato_imp_alb_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_alb_pro' height='10","i_formato_imp_alb_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PEDIDO A PROVEEDOR' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
				
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ped_pro' style='display:none;' height='10",0,"","i_formato_imp_ped_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_ped_pro' height='10","i_formato_imp_ped_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='DEVOLUCION A PROVEEDOR' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_dev_pro' style='display:none;' height='10",0,"","i_formato_imp_dev_pro",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_dev_pro' height='10","i_formato_imp_dev_pro",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='CATALOGO' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
				
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_cat' style='display:none;' height='10",0,"","i_formato_imp_cat",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_cat' height='10","i_formato_imp_cat",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
   
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ORDEN DE FABRICACION' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
				
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ord_fab' style='display:none;' height='10",0,"","i_formato_imp_ord_fab",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_ord_fab' height='10","i_formato_imp_ord_fab",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
 
                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='ORDEN' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_ord' style='display:none;' height='10",0,"","i_formato_imp_ord",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_ord' height='10","i_formato_imp_ord",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='INCIDENCIA' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                   'DrawSelectCelda "CELDA","95","' id='formatos_imp_inc' style='display:none;' height='10",0,"","i_formato_imp_inc",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_inc' height='10","i_formato_imp_inc",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""

                    rstAux.close
					strselect="select a.nformato,a.descripcion from clientes_formatos_imp as a with(nolock), formatos_imp as b with(nolock) where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='PARTE DE TRABAJO' order by descripcion"
					defecto=""
					rstAux.Open strselect, DsnIlion, adOpenKeyset, adLockOptimistic
					
                    'DrawSelectCelda "CELDA","95","' id='formatos_imp_parte' style='display:none;' height='10",0,"","i_formato_imp_parte",rstAux,defecto,"nformato","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","display:none;' id='formatos_imp_parte' height='10","i_formato_imp_parte",rstAux,enc.EncodeForHtmlAttribute(defecto),"nformato","descripcion","",""
  
                    rstAux.close
					%><div id='formatos_imp_ninguno' class="CELDA" style="display:">
					</div>
			</td><%

                    %><td class="CELDAC7 underOrange width5"><%
		                DrawCheck "","","i_facturable",true
                    %></td><%  

                    %><td class="CELDA underOrange width5"><%
                    rstAux.Open "select codigo, descripcion from almacenes with(nolock) where codigo like '"&session("ncliente")&"%' and isnull(fbaja,'')='' order by descripcion",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
                    'DrawSelectCelda "CELDA7","125","",0,"","e_almacen",rstAux,"","codigo","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","","e_almacen",rstAux,"","codigo","descripcion","",""
                    rstAux.Close
                    %></td><% 

                    %><td class="CELDA underOrange width5"><%
                    if si_tiene_modulo_ccostes <> 0 then
		            strselect="select codigo, descripcion from tiendas with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion"
		            rstAux.CursorLocation=3
		            rstAux.Open strselect, session("dsn_cliente"), adOpenKeyset, adLockOptimistic
				    'DrawSelectCelda "CELDA","100","",0,"","e_ccostes",rstAux,"","codigo","descripcion","",""
                    DrawSelect "'CELDAL7 width100'","","e_ccostes",rstAux,"","codigo","descripcion","",""
				    rstAux.close
		            end if
                    %></td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_ventas" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_compras" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_aventas" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_acompras" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_caja" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_rfventas" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_rfcompras" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_retventas" size="5"/>
                    </td><% 

                    %><td class="CELDA underOrange width5">
		                <input type="text" class="width100" ,maxlength="5", name="i_cta_retcompras" size="5"/>
                    </td><% 
                    if si_tiene_modulo_ecomerce <> 0 then 
                    %><td class="CELDAC7 underOrange width5"><%
		                DrawCheck "","","i_ocultarmkp",false
                    %></td><% 
                    end if
                closefila%>
       </table>
   <%end if%>
   </form>
<%'No hay sesión
else
	MsgError LitSinSesion%>
	<br><a href="../" target="_top">Iniciar sesión</a>
<%end if%>
</BODY>
</HTML>
