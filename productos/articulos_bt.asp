<%@ Language=VBScript %>
<%
''EJM 08/06/2006: Inserción de nuevos campos proyecto LENTICOM
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  

<title><%=LitTitulo%></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="../adovbs.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->

<!--#include file="articulos.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<%si_tiene_modulo_importaciones=ModuloContratado(session("ncliente"),ModImportaciones)
si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
si_tiene_modulo_ecomerce=ModuloContratado(session("ncliente"),ModEComerce)
' osm 28/04/15 comprobar si tiene módulo intelitur
si_tiene_modulo_intelitur=ModuloContratado(session("ncliente"),ModIntelitur)
' osm 28/04/15 obtener provincia e id de la provincia del usuario con intelitur
if si_tiene_modulo_intelitur <> 0 then 
    provinciaUser = d_lookup("provincia", "domicilios", "pertenece='"&session("ncliente")&session("usuario")&"' and tipo_domicilio='personal'", session("dsn_cliente"))
    idProvinciaUser = d_lookup("NDETLISTA", "CAMPOSPERSOLISTA", "valor='"&provinciaUser&"' and tabla='ARTICULOS' and ncampo='"&session("ncliente")&"02'", session("dsn_cliente"))
end if
''MPC 26/05/2009 Se obtiene el campo horecas de configuración para cambiar el nombre de un campo o dejarlo
horecas=d_lookup("horecas", "configuracion", "nempresa='"&session("ncliente")&"'", session("dsn_cliente"))%>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
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
/*
if (window.document.addEventListener) {
    window.document.addEventListener("keydown", callkeydownhandler, false);
}
else
{
    window.document.attachEvent("onkeydown", callkeydownhandler);
}

function callkeydownhandler(evnt) {
    ev = (evnt) ? evnt : event;
    //Texto_onkeypress(ev);
}

function Texto_onkeypress(e)
{
    var keycode = e.keyCode;
    if (keycode == 13) Buscar('<%=formulario%>');
}
*/
function comprobar_enter(){
	//si se ha pulsado la tecla enter
	//if (window.event.keyCode==13){
		//document.opciones.criterio.focus();
		Buscar()
	//}
}
function Buscar() {
	SearchPage("articulos_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value,1);

    document.opciones.texto.value = "";
}

//Validación de campos numéricos y fechas.
function ValidarCampos(){    
	f=parent.pantalla.marcoPropiedades.document.PropiedadesArticulo;	
	

    if(f.urlRewrite!=null && f.urlRewrite.value!="")
    {
        if(f.urlRewrite.value=="True" )
        {
            
            if(!checkURLPattern(f.url_mostrar.value))
            {
                window.alert("<%=LITURLMOSTRARINVALIDA%>");
                try{f.url_mostrar.focus();}catch(err){}
                return false;
            }     
        }
    }
    
	//ricardo 27-12-2006 se añade los campos para los formatos de etiquetas
	if (isNaN(f.cantidadarticulo.value.replace(",","."))) {
		window.alert("<%=LitMedNoNum%>");
		f.cantidadarticulo.focus();
		f.cantidadarticulo.select();
		return false;
	}


    //trim
    f.referencia.value=f.referencia.value.replace(/^\s+|\s+$/g, '');

	if (f.referencia.value=="" && parent.pantalla.document.articulos.autoref.value=="0") {
		window.alert("<%=LitMsgReferenciaNoNulo%>");
		return false;
    }

	if (f.prefcodbarras.value=="CI" && f.referencia.value!="" && (isNaN(f.referencia.value) || f.referencia.value.length>5)) {
		window.alert("<%=LitMsgCodBarImpRefNum%>");
		return false;
	}

	if (f.prefcodbarras.value=="CI" && (f.agrupa_tallas.value!="" || f.agrupa_colores.value!="")) {
		window.alert("<%=LitMsgCodBarImpTyC%>");
		return false;
	}

	if (f.nombre.value=="") {
	   window.alert("<%=LitMsgNombreNoNulo%>");
	   return false;
    }

	if (comp_car_ext(f.referencia.value,1)==1 || f.referencia.value.indexOf("/")!=-1){
		window.alert("<%=LitMsgRefDesCarNoVal%>");
		return false;
	}

    if (f.nombre.value!=""){
        //trim
        f.nombre.value=f.nombre.value.replace(/^\s+|\s+$/g, '');
        //delete tabs at the front and end
        f.nombre.value=f.nombre.value.replace(/^\t+|\t+$/g, '');

        //we dont permit tab char
        if ( f.nombre.value.indexOf(String.fromCharCode(9))>=0 ) {
		    window.alert("<%=LitMsgRefDesCarNoVal%>");
		    return false;
	    }
    }

    if (comp_car_ext(f.nombre.value,2)==1){
		window.alert("<%=LitMsgRefDesCarNoVal%>");
		return false;
	}

	if (isNaN(f.spvp.value.replace(",",".")) || f.spvp.value=="" || f.spvp.value<0) {
		window.alert("<%=LitMsgPvPNumerico%>");
		return false;
	}
	if (f.divisa.value=="") {
		window.alert("<%=LitMsgDivisaNoNulo%>");
		return false;
	}
	if (f.iva.value=="") {
		window.alert("<%=LitMsgIvaNoNulo%>");
		return false;
    }
    if (f.mandatorySubfamily.value == "True") {
      if (f.auxfamilia.value == "") {
          window.alert("<%=LITVALSUBFAMILY%>");
          return false;
        }
    }


    //Descuento
	if(f.descuento.value > 100 || f.descuento.value < 0)
	{
	    window.alert("<%=LITDTORANGE%>");
	    return false;
	}

	if (isNaN(f.descuento.value.replace(",","."))) {
		window.alert("<%=LitMsgDescuentoNumerico%>");
		return false;
	}
	
	
	if (isNaN(f.peso.value.replace(",","."))) {
		window.alert("<%=LITPESONONUMERICO%>");
		f.peso.focus();
		return false;
	}
	
	
	
	<%'i(EJM 08/06/2006)
	''MPC 26/05/2009 Se obtiene el campo horecas de configuración para cambiar el nombre de un campo o dejarlo
	if si_tiene_modulo_tiendas=0 and horecas = 0 then%>
	if (isNaN(f.mesesCaducidad.value)) {
		window.alert("<%=LitMsgMesesNumerico%>");
		return false;
	}
	<%end if%>
	if (isNaN(f.importeABparcial.value.replace(",","."))) {
		window.alert("<%=LitMsgImporteABparcialNumerico%>");
		return false;
	}	
	<%'fin(EJM 08/06/2006)%>
	
	si_tiene_modulo_comercial=parent.pantalla.document.articulos.si_tiene_modulo_comercial.value;
	if (si_tiene_modulo_comercial!=0){
		if (isNaN(f.porcom.value.replace(",","."))) {
			window.alert("<%=LitMsgPorComisionNumerico%>");
			return false;
		}
	}
	if (isNaN(f.mesesmo.value.replace(",",".")) || isNaN(f.mesesde.value.replace(",",".")) || isNaN(f.mesesmt.value.replace(",","."))) {
		window.alert("<%=LitMsgMesesNumerico%>");
		return false;
	}
			
	//  GPD (27/02/2007).
	if (isNaN(f.ue.value) || f.ue.value=="") {		
		window.alert("<%=LitMsgUnidadEmbalajeNumerico%>");
		return false;
	}
	//DBS (29/11/2013).
	if (isNaN(f.uv.value) || f.uv.value=="") {		
	    window.alert("<%=LitMsgUnidadVentaNumerico%>");
	    return false;
	}

	if (f.codbarras.value!="" && isNaN(f.codbarras.value)) {
		//alert(f.codbarras.value);
		window.alert("<%=LITCODBARNUMART%>");
		return false;
	}

// && !(f.checkcodbarras2.checked)
	si_tiene_modulo_terminales=parent.pantalla.document.articulos.si_tiene_modulo_terminales.value;
	if (si_tiene_modulo_terminales!=0){
		if (f.carga_terminal.checked && (f.codbarras.value=="" ) && f.prefcodbarras.value==""){
			window.alert("<%=LitMsgNoCargaTerminal%>");
			return false;
		}
	}
	if (f.fbaja.value!="" && !checkdate(f.fbaja)) {
		window.alert("<%=LitMsgFechaBajaFecha%>");
		return false;
	}

	////////
	b = 1;
	c=0;

/*
if (f.codbarras.value.length == 13){

longitud=f.codbarras.value.length;

    		for(a=0;a<longitud-1;a++){
     		   if(b == 1){
	            c = c + parseInt(f.codbarras.value.substring(a,a+1));
      	      b = 0;
			}
	        else{
      	      c = c + (parseInt(f.codbarras.value.substring(a,a+1)) * 3);
            	b = 1;
			};
		};

	    h = 0;
	    g=c + h;
	    ff=g % 10;

	    while (ff!=0) {
      	  h++;
		  g=c + h;
		  ff=g % 10;
	    }
	    CrearControl = h;
	  }
	else CrearControl = 0;

	if (f.codbarras.value.length>1 && (f.codbarras.value.length!=13 || CrearControl!=parseInt(f.codbarras.value.substring(longitud-1,longitud))))
	 {
		window.alert("<%=LitMsgCod_BarrasMal%>");
		return false;
	}
*/
	///////

	/*if (f.genera.checked){
		if (f.agrupa_tallas.value=="") {
			if (f.agrupa_colores.value=="") {
				window.alert("Debe seleccionar una agrupación de tallas o de colores o ambas");
				return false;
			}
		}
	}*/

<%if si_tiene_modulo_importaciones=0 then%>
	if (parent.pantalla.document.articulos.si_campo_personalizables.value==1){
		num_campos=parent.pantalla.marcoCamposPersonalizables.document.CamposPersonalizablesArt.num_campos.value;

		respuesta=comprobarCampPerso("parent.pantalla.marcoCamposPersonalizables.",num_campos,"CamposPersonalizablesArt");
		if(respuesta!=0){
			titulo="titulo_campo" + respuesta;
			tipo="tipo_campo" + respuesta;
			titulo=parent.pantalla.marcoCamposPersonalizables.document.CamposPersonalizablesArt.elements[titulo].value;
			tipo=parent.pantalla.marcoCamposPersonalizables.document.CamposPersonalizablesArt.elements[tipo].value;
			if (tipo==4) {
				nomTipo="<%=LitTipoNumericoArt%>";
			}
			else if (tipo==5) {
				nomTipo="<%=LitTipoFechaArt%>";
			}

			window.alert("<%=LitMsgCampoArt%> " + titulo + " <%=LitMsgTipoArt%> " + nomTipo);

			return false;
		}
	}
<%end if%>

    //ricardo 12-11-2007 si tiene valor el parametro ne , no se puede dejar el codigo de barras a nulo
//window.alert(parent.pantalla.document.articulos.ne.value + "-" + f.codbarras.value + "-")
//return false;
    if (parent.pantalla.document.articulos.ne.value!=""){
        if (f.codbarras.value==""){
		    window.alert("<%=LitMsgNoCodBarras%>");
		    return false;
        }
    }

	return true;
}

    function ValidarCampos2() {

	if (parent.pantalla.document.articulos.referencia.value=="") {
		window.alert("<%=LitMsgReferenciaNoNulo%>");
		return false;
	}
	else{
	   if (parent.pantalla.document.articulos.nombre.value=="") {
		   window.alert("<%=LitMsgNombreNoNulo%>");
		   return false;
	   }
	   else{
	      if (parent.pantalla.document.articulos.divisa.value=="") {
		      window.alert("<%=LitMsgDivisaNoNulo%>");
		      return false;
           }
           if (f.mandatorySubfamily.value == "True") {
                alert("<%=LITVALSUBFAMILY%>");
                document.getElementByName.auxfamilia.focus();
                return false;
            }
	   }

	}
	
	//  GPD (27/02/2007).
	if (isNaN(f.ue.value) || f.ue.value=="") {		
		window.alert("<%=LitMsgUnidadEmbalajeNumerico%>");
		return false;
	}
	//DBS (29/11/2013).
	if (isNaN(f.uv.value) || f.uv.value=="") {		
	    window.alert("<%=LitMsgUnidadVentaNumerico%>");
	    return false;
	}
	if (isNaN(parent.pantalla.document.articulos.descuento.value)) {
		window.alert("<%=LitMsgDescuentoNumerico%>");
		return false;
	}
	else{
	   if (isNaN(parent.pantalla.document.articulos.porcom.value)) {
		     window.alert("<%=LitMsgPorComisionNumerico%>");
		     return false;
	   }
	   else{
	      if (isNaN(parent.pantalla.document.articulos.meses.value)) {
		        window.alert("<%=LitMsgMesesNumerico%>");
		        return false;
	      }
	   }
	}
	return true;
}

//Realizar la acción correspondiente al botón pulsado.
function Accion(mode,pulsado) {

    //osm 28/04/15 bloquear botones añadir y guardar en modo browse y guardar en modo edit y add
    var bloquear = false;
    <%mode=request.querystring("mode")%>
    <%if si_tiene_modulo_intelitur <> 0 and (mode = "edit" or mode = "add" or mode = "browse") then %>
        <% if idProvinciaUser&""="" or isnull(idProvinciaUser) then %>
            bloquear = true
        <% end if %>
    <% end if %>
    
	switch (mode) {
		case "browse":
			switch (pulsado) {
				case "add": //Nuevo registro
                    if(bloquear) {
                         window.alert("Debes añadir o corregir tu provincia en tu perfil");
                    }
                    else {
                        if(parent.pantalla.document.articulos.gestion.value=="0")
                        {
                            window.alert("<%=LitMsgImposibleAnyadir%>");
                        }
                        else
                        {
					        parent.pantalla.document.articulos.action="articulos.asp?mode=" + pulsado;
                            parent.pantalla.document.articulos.submit();
                            document.location="articulos_bt.asp?mode=" + pulsado;
                        }
                    }
				   
					break;

				case "edit": //Editar registro					
				    if(bloquear) {
				        window.alert("Debes añadir o corregir tu provincia en tu perfil");
				    }
				    else {
				        //ricardo 8-9-2006 si artsl=1 el articulo una vez creado solamente se pueda ver 
				        if(parent.pantalla.document.articulos.artsl.value!="1")
				        {
				            if(parent.pantalla.document.articulos.gestion.value=="0")
				            {
				                window.alert("<%=LitMsgEditarArt%>");
				            }
				            else
				            {
				                parent.pantalla.document.articulos.action="articulos.asp?referencia=" + parent.pantalla.document.articulos.hreferencia.value +
                                "&mode=" + pulsado+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
				                parent.pantalla.document.articulos.submit();
				                document.location="articulos_bt.asp?mode=" + pulsado;
				            }
				        }
				        else{
				            window.alert("<%=LitArtNoEdit%>");
				        }
				    }
				    				
					
					break;

				case "delete": //Eliminar registro
				//ricardo 8-9-2006 si artsl=1 el articulo una vez creado solamente se pueda ver 
				    if(parent.pantalla.document.articulos.artsl.value!="1")
				    	{
						if (parent.pantalla.document.articulos.hes_padre.value!=0) {
							msg="<%=LitMsgEliminarArticuloTyCConfirm%>";
						}
						else {
							msg="<%=LitMsgEliminarArticuloConfirm%>"
						}
						if (window.confirm(msg)==true) {
							parent.pantalla.document.articulos.action="articulos.asp?mode=" + pulsado + "&referencia=" + parent.pantalla.document.articulos.hreferencia.value+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
							parent.pantalla.document.articulos.submit();
							document.location="articulos_bt.asp?mode=browse";
						}
					}
					else{
						window.alert("<%=LitArtNoBorr%>");
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
			        if(bloquear) {
			            window.alert("Debes añadir o corregir tu provincia en tu perfil");
			        }
			        else {
			            if(parent.pantalla.document.articulos.gestion.value=="0")
                        {
			                window.alert("<%=LitMsgImposibleAnyadir%>");
			            }
			            else
			            {
			                if(parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.url_mostrar!=null)
			                    var url2 = parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.url_mostrar.value;
			                if(parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.impr_cat!=null)
			                    var tienda2 = parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.impr_cat.value;
			                if(parent.pantalla.document.articulos.hreferencia!=null)
			                {
			                    var referencia = parent.pantalla.document.articulos.hreferencia.value;
			                    var ref_final = referencia.substring(5,referencia.lenght);
			                }
                        
                            if (tienda2) {
                                    <%if si_tiene_modulo_ecomerce <> 0 then %>
                                SigueGuardando(url2, ref_final);
			                    <%else %>
                                parent.pantalla.GestionPropiedades('save', parent.pantalla.document.articulos.hreferencia.value, parent.pantalla.document.articulos.hes_padre.value)
                                    <% end if%>
                               }
                            else {
                                parent.pantalla.GestionPropiedades('save',parent.pantalla.document.articulos.hreferencia.value,parent.pantalla.document.articulos.hes_padre.value)				        
                            }
			                    
			            }
			        }
				   
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.articulos.action="articulos.asp?referencia=" + parent.pantalla.document.articulos.hreferencia.value +
					"&mode=browse";
					parent.pantalla.document.articulos.submit();
					document.location="articulos_bt.asp?mode=browse"+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
					break;

				case "search": //Buscar datos
					break;
			}
			break;

		case "add":
			switch (pulsado) {
			    case "save": //Guardar registro
			        if(bloquear) {
			            window.alert("Debes añadir o corregir tu provincia en tu perfil");
			        }
			        else {
		                if(parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.url_mostrar!=null)
		                    var url = parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.url_mostrar.value;
                        if(parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.impr_cat!=null)
		                    var tienda = parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.impr_cat.value;
		                if(parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.referencia!=null)
	                        var referencia = parent.pantalla.marcoPropiedades.document.PropiedadesArticulo.referencia.value;
                    
	                    if(parent.pantalla.document.articulos.gestion.value=="0")
		                {
		                    window.alert("<%=LitMsgImposibleAnyadir%>");
		                }
		                else
		                {
		                    f=parent.pantalla.marcoPropiedades.document.PropiedadesArticulo
		                    if (ValidarCampos()) 
		                    {
		                        f=parent.pantalla.marcoPropiedades.document.PropiedadesArticulo;
		                        <%if si_tiene_modulo_importaciones=0 then %>
                                    if (parent.pantalla.document.articulos.si_campo_personalizables.value==1){
						
                                        //ricardo 24-3-2004 copiamos los campos del marco campospersonalizadosart
                                        cadena_campos_perso="";
                                        num_campos_perso=parent.pantalla.marcoCamposPersonalizables.document.CamposPersonalizablesArt.num_campos.value;
                                        cadena_campos_perso=cadena_campos_perso + "&num_campos_perso=" + num_campos_perso;
                                        for (ki=1;ki<=num_campos_perso;ki++)
                                        {
                                            if (eval("parent.pantalla.marcoCamposPersonalizables.document.CamposPersonalizablesArt.campo" + ki + ".value=='on'")){
                                                if (eval("parent.pantalla.marcoCamposPersonalizables.document.CamposPersonalizablesArt.campo" + ki + ".checked==true")){
                                                    cadena_campos_perso=cadena_campos_perso + "&campo" + ki + "=1";
                                                }
                                                else
                                                {
                                                    cadena_campos_perso=cadena_campos_perso + "&campo" + ki + "=0";
                                                }
                                            }
                                            else
                                            {
                                                cadena_campos_perso=cadena_campos_perso + "&campo" + ki + "=" + eval("parent.pantalla.marcoCamposPersonalizables.document.CamposPersonalizablesArt.campo" + ki + ".value")
                                            }
                                        }
                                    }
		                            else{
                                                cadena_campos_perso="";
		                            }
						
		                        <%else%>
                                    cadena_campos_perso="";
		                        <%end if%>

                                cadena_subcta="";
		                        cadena_subcta=cadena_subcta + "&subctaventas=" + parent.pantalla.marcoDatosContaDeArticulo.document.DatosContaDeArticulo.subctaventas.value;
		                        cadena_subcta=cadena_subcta + "&subctaabventas=" + parent.pantalla.marcoDatosContaDeArticulo.document.DatosContaDeArticulo.subctaabventas.value;
		                        cadena_subcta=cadena_subcta + "&subctacompras=" + parent.pantalla.marcoDatosContaDeArticulo.document.DatosContaDeArticulo.subctacompras.value;
		                        cadena_subcta=cadena_subcta + "&subctaabcompras=" + parent.pantalla.marcoDatosContaDeArticulo.document.DatosContaDeArticulo.subctaabcompras.value;

		                        cadena_pnf="";
		                        h_has_manufacturer_item=parent.pantalla.document.articulos.h_has_manufacturer_item.value;
		                        if (h_has_manufacturer_item==1)
		                        {
		                            if (parent.pantalla.marcoDatosFabDeArticulo!=null){
		                                cadena_pnf=cadena_pnf + "&nmanufacturer=" + parent.pantalla.marcoDatosFabDeArticulo.document.DatosFabDeArticulo.nmanufacturer.value;
		                                cadena_pnf=cadena_pnf + "&pnf=" + parent.pantalla.marcoDatosFabDeArticulo.document.DatosFabDeArticulo.pnf.value;    
		                            }
		                        }
                        
		                    if(tienda=="true")
		                    {
		                        <%if si_tiene_modulo_ecomerce<>0 then %>
                                SigueGuardando(url,referencia);
		                        <%else %>
		                        //if (f.genera.checked){
                                    if (false) 
		                            {
		                                if (window.confirm("<%=LitMsgArtGenConfirm%>")==true){
		                                    f.action="articulos.asp?mode=first_save&referencia=" + f.referencia.value +
                                            "&genera=si" + cadena_subcta + cadena_pnf + cadena_campos_perso+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
		                                    f.submit();
		                                    document.location="articulos_bt.asp?mode=browse";
		                                }
		                            }
		                            else
		                            {
                                                f.action="articulos.asp?mode=first_save&referencia="+f.referencia.value + cadena_subcta + cadena_pnf + cadena_campos_perso+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
		                                f.submit();
		                                document.location="articulos_bt.asp?mode=browse";
		                            }
		                        <%end if%>
		                    }
		                    else
		                    {
					            if (false) 
		                        {
		                            if (window.confirm("<%=LitMsgArtGenConfirm%>")==true){
		                                f.action="articulos.asp?mode=first_save&referencia=" + f.referencia.value +
                                        "&genera=si" + cadena_subcta + cadena_pnf + cadena_campos_perso+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
		                                f.submit();
		                                document.location="articulos_bt.asp?mode=browse";
		                            }
		                        }
		                        else
		                        {
						  	        f.action="articulos.asp?mode=first_save&referencia="+f.referencia.value + cadena_subcta + cadena_pnf + cadena_campos_perso+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
		                            f.submit();
		                            document.location="articulos_bt.asp?mode=browse";
		                        }
                            }
                        }
                        }
			        }
				    
				    
					break;

				case "cancel": //Cancelar edición
					parent.pantalla.document.articulos.action="articulos.asp?mode=add"+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
					parent.pantalla.document.articulos.submit();
					document.location="articulos_bt.asp?mode=add";
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

//FUNCIONES AJAX URL
var xmlHttp;
function CreateXmlHttp()
{

    // Probamos con IE
    try
    {
        // Funcionará para JavaScript 5.0
        xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
    }
    catch(e)
    {
        try
        {
            xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        }
        catch(oc)
        {
            xmlHttp = null;
        }
    }

    // Si no se trataba de un IE, probamos con esto
    if(!xmlHttp && typeof XMLHttpRequest != "undefined")
    {
        xmlHttp = new XMLHttpRequest();
    }

    return xmlHttp;
}

function recogeInfo()
{
    if(xmlHttp.readyState == 4 && xmlHttp.status == 200)
    {
        
       results = xmlHttp.responseText.split(",");    
       tipo = results[1]; 
       nombre = results[2]; 
       if(results[0]==1)
       {
            
           <%if Request.QueryString("mode")="add" then %>
               if (false) 
		        {
		      	    if (window.confirm("<%=LitMsgArtGenConfirm%>")==true)
		      	    {
			     	    f.action="articulos.asp?mode=first_save&referencia=" + f.referencia.value +
			     	    "&genera=si" + cadena_campos_perso+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
				 	    f.submit();
			     	    document.location="articulos_bt.asp?mode=browse";
			  	    }
		        }
		        else
		        {
			  	    f.action="articulos.asp?mode=first_save&referencia="+f.referencia.value + cadena_campos_perso+"&c01="+parent.pantalla.document.articulos.fc01.value+"&c02="+parent.pantalla.document.articulos.fc02.value+"&c03="+parent.pantalla.document.articulos.fc03.value;
			  	    f.submit();
			  	    document.location="articulos_bt.asp?mode=browse";
		        }
		   <%end if%>
		   <%if Request.QueryString("mode")="edit" then %>
		         parent.pantalla.GestionPropiedades('save',parent.pantalla.document.articulos.hreferencia.value,parent.pantalla.document.articulos.hes_padre.value)
		   <%end if%>
       }
       else
       {
            alert("<%=LitUrlExistente%>"+" "+tipo+" "+nombre);
       }
    }
}

function SigueGuardando(url,referencia)
{
    // 1.- Creamos el objeto xmlHttpRequest
    CreateXmlHttp();
    cod=new Date().getTime();
    // 2.- Definimos la llamada para hacer un simple GET.
    var ajaxRequest = 'getUrlRewrite.asp?url='+url+'&code='+cod+'&ref='+referencia;

    // 3.- Marcar qué función manejará la respuesta
    xmlHttp.onreadystatechange = recogeInfo;

    // 4.- Enviar
    xmlHttp.open("GET", ajaxRequest, true);
    xmlHttp.send("");  
}

</script>

<body class="body_master_ASP">

<%
mode=Request.QueryString("mode")
%>

<%
''ricardo 19-6-2006 se recoge el parametro de articulos_buscar para poner por defecto el campo a buscar
	dim bp
	dim pnf
    
	ObtenerParametros("articulos_buscar")
    titulo_campo_perso=""
	if pnf & "">"" then
	    numero_campo=replace(pnf,"campo","")
	    titulo_campo_perso=d_lookup("titulo","camposperso","ncampo='" & session("ncliente") & numero_campo & "' and tabla='ARTICULOS'",session("dsn_cliente"))
	end if
%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>" />
    <input type="hidden" name="respuesta" id="respuesta" value=""/>
    <%
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    set rstAux = Server.CreateObject("ADODB.Recordset")
    conn.open DSNILION
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandText= "ContractedItem"
    command.CommandType = adCmdStoredProc
        command.Parameters.Append command.CreateParameter("@nempresa", adVarChar, adParamInput, 5, session("ncliente"))
        command.Parameters.Append command.CreateParameter("@objeto", adVarChar, adParamInput, 1000, OBJManufacturer)
    command.Execute,,adExecuteNoRecords
    set rstAux = command.Execute

    if not rstAux.eof then
        has_manufacturer_item = rstAux("result")
    else
        has_manufacturer_item = "0"
    end if
    
   conn.close
     %>

<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		<tr>
        <%
			if mode="browse" then
			    
				%>
                   <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
<!-- ricardo 25-5-2006 se oculta el boton de borrar , dicho por JAR-->
<!--
				    <td CLASS="CELDABOT" onclick="javascript:Accion('browse','delete');">
					    <%PintarBotonBT LITBOTBORRAR,ImgBorrar,ParamImgBorrar,LITBOTBORRARTITLE%>
				    </td>
-->
				    <td id="idedit" class="CELDABOT" onclick="javascript:Accion('browse','edit');">
					    <%PintarBotonBT LITBOTEDITAR,ImgEditar,ParamImgEditar,LITBOTEDITARTITLE%>
				    </td>
				    <td id="idprint" class="CELDABOT" onclick="javascript:Accion('browse','print');">
					    <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				    </td>
				<%
			elseif mode="search" then

				%>
                   <td id="idadd" class="CELDABOT" onclick="javascript:Accion('browse','add');">
					    <%PintarBotonBT LITBOTANADIR,ImgAdd,ParamImgAdd,LITBOTANADIRTITLE%>
				    </td>
				<%

			elseif mode="edit" then
				%>
			        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('edit','save');">
					    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				<%

			elseif mode="add" then
				%>
			        <td id="idsave" class="CELDABOT" onclick="javascript:Accion('add','save');">
					    <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				    </td>
				    <td id="idcancel" class="CELDABOT" onclick="javascript:Accion('add','cancel');">
					    <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				    </td>
				<%
			end if
			%>
		        </tr>
	        </table>
            </div>
    
    <div id="FILTERS_MASTER_ASP">
            <!--<td CLASS=CELDABOT><%=LitBuscar & ": "%></td>
			<td CLASS=CELDABOT>-->
				<select class="IN_S" name="campos">
				  <option value="a.referencia" <%=iif(ucase(bp)=ucase(Litreferencia),"selected","")%>><%=Litreferencia%></option>
        		  <option value="a.nombre" <%=iif(ucase(bp)=ucase(Litnombre) or bp="" or (bp<>"" and ucase(bp)<>ucase(Litreferencia)),"selected","")%>><%=Litnombre%></option>
<%if si_tiene_modulo_importaciones=0 then%>
				  <option value="a.cod_barras" <%=iif(ucase(bp)=ucase("Cod.Barras"),"selected","")%>><%=LITCODBARRAS%></option>
<%end if%>
				  <option value="p.su_ref" <%=iif(ucase(bp)=ucase(LitRefProveedor),"selected","")%>><%=LitRefProveedor%></option>
<%if has_manufacturer_item <> 0 then %>
				  <option value="a.pnf" <%=iif(ucase(bp)=ucase("pnf"),"selected","")%>><%=LITPNF%></option>
<%end if %>
<%if pnf&"">"" then %>
    <option value="<%="a." & pnf%>" <%=iif(ucase(bp)=ucase("pnf"),"selected","")%>><%=titulo_campo_perso%></option>
<%end if 

%>
		        </select>
			<!--</td>
			<td CLASS=CELDABOT>-->
				<select class="IN_S" name="criterio">
					<OPTION value="contiene"><%=LitContiene%></OPTION>
					<!--<OPTION value="empieza"><%=LitComienza%></OPTION>-->
					<OPTION value="termina"><%=LitTermina%></OPTION>
					<OPTION value="igual"><%=LitIgual%></OPTION>
				</select>
			<!--</td>
			<td CLASS=CELDABOT>-->
                <input id="KeySearch" class="IN_S" type="text" name="texto" size="15" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
			<!--</td>
			<td CLASS=CELDABOT>-->
				<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			<!--</td>
		</tr>
	</table>-->
    </div>
    </div>
    <table style="width:100%;height:42px;vertical-align:bottom;" align="center">
    <tr>
    <td style="width:100%;height:42px; vertical-align:bottom; text-align:center;">
    <%ImprimirPie_bt
    set conn=nothing
set command=nothing
set rstAux=nothing
    %>
    </td>
    </tr>
    </table>
</form>
</body>
</html>
