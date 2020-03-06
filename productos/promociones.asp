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
' ############ SOLO PARA DEVOLVER CONSULTAS AJAX ############
if request.querystring("mode") = "consultaAJAX" then
    if request.querystring("consulta") = "actualizarGrupos" then
        set rstAux = Server.CreateObject("ADODB.Recordset")
        sql = "EXEC ActualizarGruposEnFranquicias @nempresaCentral='" & session("ncliente") & "'"	
		rstAux.open sql,DSNImport
        if rstAux("devolver") = "0" then
            response.Write("OK")
        else
            response.Write("ERROR")
        end if
        rstAux.close
        set rstAux =nothing
    end if
   
    ' Fin de consulta AJAX
    response.End
end if
' ################################################
'' JCI 18/06/2003 : MIGRACION A MONOBASE
'RGU 13/10/2006: Añadir campo pvp+iva en el span de precios
%>
<%response.buffer=true%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../modulos.inc" -->

<!--#include file="../CatFamSubResponsive.inc" -->

<!--Para Nuevo Estilo -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../js/tabs.js.inc" -->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="promociones.inc" -->

<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">

    animatedcollapse.addDiv('CABECERA', 'fade=1')
    animatedcollapse.addDiv('ARTICULOS', 'fade=1')
    animatedcollapse.addDiv('CONDICIONES', 'fade=1')
    animatedcollapse.addDiv('CONDICIONESAPLICACION', 'fade=1')

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()

    if (window.document.addEventListener) {
        window.document.addEventListener("keydown", callkeydownhandler, false);
    } else {
        window.document.attachEvent("onkeydown", callkeydownhandler);
    }

    function callkeydownhandler(evnt) {
        ev = (evnt) ? evnt : event;
        tecla_pulsada = ev.keyCode;
    }

    function CampoRefPulsado(mode, marco, formulario, queordenar, comoordenar) {
        if (tecla_pulsada == 13) {
            continuar = 0;
            if (mode == "ALTA") {
                if (document.promociones.RefPro.value != "") continuar = 1;
            }
            if (mode == "BAJA") {
                if (document.promociones.bRefPro.value != "") continuar = 1;
            }
            if (continuar == 1) Insertar(mode, '1', 'insertar', queordenar, comoordenar);
        }
    }

    function OrdenarDatos(mode, marco, formulario, campo) {
        campo = campo.toUpperCase();
        eval("queordenar=" + marco + ".document." + formulario + ".queordenar.value.toUpperCase()");
        eval("comoordenar=" + marco + ".document." + formulario + ".comoordenar.value.toUpperCase()");
        if (campo != queordenar || comoordenar == "") {
            queordenar = campo;
            comoordenar = "ASC";
        }
        else {
            if (campo == queordenar && comoordenar == "ASC") comoordenar = "DESC";
            else comoordenar = "ASC";
        }
        queimagen1 = "";
        queimagen2 = "";
        queimagen3 = "";
        comoimagen = "";
        if (comoordenar == "ASC") comoimagen = "&darr;";
        if (comoordenar == "DESC") comoimagen = "&uarr;";
        if (queordenar == "A.REFERENCIA") {
            queimagen1 = comoimagen;
            queimagen2 = "&harr;";
            queimagen3 = "&harr;";
        }
        if (queordenar == "A.NOMBRE") {
            queimagen2 = comoimagen;
            queimagen1 = "&harr;";
            queimagen3 = "&harr;";
        }
        if (queordenar == "F.NOMBRE") {
            queimagen3 = comoimagen;
            queimagen2 = "&harr;";
            queimagen1 = "&harr;";
        }
        if (mode == "ALTA") {
            document.getElementById("OD1A").innerHTML = queimagen1;
            document.getElementById("OD2A").innerHTML = queimagen2;
            document.getElementById("OD3A").innerHTML = queimagen3;
        }
        else {
            document.getElementById("OD1B").innerHTML = queimagen1;
            document.getElementById("OD2B").innerHTML = queimagen2;
            document.getElementById("OD3B").innerHTML = queimagen3;
        }
        Insertar(mode, '1', 'first', queordenar, comoordenar);
    }

    //              Insertar('ALTAGRUPO', '1', 'first', 'A.referencia', 'asc')
    function Insertar(mode, pag, sentido, queordenar, comoordenar) {
        switch (mode) {
            case "ALTA":
                mod = "save";
                if (sentido == "first") {
                    sentido = "&submode=first";
                    mod = "first";
                    document.promociones.condbase.value = "0";
                }

                if (sentido == "insertar") {
                    sentido = "&submode=insertar";
                    mod = "insertar";
                }

                pagina = "ArticulosDePromocion.asp?mode=" + mod + "&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.promociones.htarifa.value
			+ "&ref=" + document.promociones.refcontiene.value + "&familia=" + document.promociones.familia.value + "&categoria="
			+ document.promociones.categoria.value + "&familia_padre=" + document.promociones.familia_padre.value + "&tipoarticulo="
			+ document.promociones.tipoarticulo.value + "&desc=" + document.promociones.descontiene.value + "&nproveedor="
			+ document.promociones.proveedor.value + "&queordenar=" + queordenar + "&comoordenar=" + comoordenar
			+ "&rangodesde=" + document.promociones.rangodesde.value + "&rangohasta=" + document.promociones.rangohasta.value + "&pvpiva=" + document.promociones.pvpiva.checked
			+ "&viene=ALTA";

                if (mod == "first") {
                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    document.getElementById("frArticulosAdd").src = pagina;
                }
                else {
                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    marcoArticulosAdd.document.ArticulosDePromocion.action = pagina;
                    marcoArticulosAdd.document.ArticulosDePromocion.submit();
                }
                document.promociones.check.checked = true;
                break;
            case "ALTAGRUPO":
                mod = "save";
                if (sentido == "first") {
                    sentido = "&submode=first";
                    mod = "first3";
                    document.promociones.condbase.value = "0";
                }

                if (sentido == "insertar") {
                    sentido = "&submode=insertar";
                    mod = "insertar";
                }

                pagina = "ArticulosDePromocion.asp?mode=" + mod + "&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.promociones.htarifa.value
			+ "&ref=" + document.promociones.refcontiene.value + "&familia=" + document.promociones.familia.value + "&categoria="
			+ document.promociones.categoria.value + "&familia_padre=" + document.promociones.familia_padre.value + "&tipoarticulo="
			+ document.promociones.tipoarticulo.value + "&desc=" + document.promociones.descontiene.value + "&nproveedor="
			+ document.promociones.proveedor.value + "&queordenar=" + queordenar + "&comoordenar=" + comoordenar
			+ "&rangodesde=" + document.promociones.rangodesde.value + "&rangohasta=" + document.promociones.rangohasta.value + "&pvpiva=" + document.promociones.pvpiva.checked
			+ "&viene=ALTAGRUPO" + "&codigog=" + document.promociones.codigog.value + "&descripciong=" + document.promociones.descripciong.value;

                if (mod == "first3") {
                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    document.getElementById("frArticulosAdd").src = pagina;
                }
                else {
                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    marcoArticulosAdd.document.ArticulosDePromocion.action = pagina;
                    marcoArticulosAdd.document.ArticulosDePromocion.submit();
                }
                document.promociones.check.checked = true;
                break;

            case "BAJA":
                mod = "save2";
                if (sentido == "first") {
                    sentido = "&submode=first";
                    mod = "first2";
                }

                if (sentido == "insertar") {
                    sentido = "&submode=insertar";
                    mod = "insertar";
                }

                pagina = "ArticulosDePromocion.asp?mode=" + mod + "&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.promociones.htarifa.value +
			"&bref=" + document.promociones.brefcontiene.value + "&bfamilia=" + document.promociones.familia1.value + "&bcategoria=" +
			document.promociones.categoria1.value + "&bfamilia_padre=" + document.promociones.familia_padre1.value + "&btipoarticulo=" +
			document.promociones.btipoarticulo.value + "&bdesc=" + document.promociones.bdescontiene.value + "&bnproveedor=" +
			document.promociones.bproveedor.value + "&queordenar=" + queordenar + "&comoordenar=" + comoordenar +
			"&brangodesde=" + document.promociones.brangodesde.value + "&brangohasta=" + document.promociones.brangohasta.value +
			"&bpvpiva=" + document.promociones.bpvpiva.value + "&viene=BAJA"; ;
                if (mod == "first2") {
                    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    document.getElementById("frArticulosBorrar").src = pagina;
                } else {
                    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    marcoArticulosBorrar.document.ArticulosDePromocion.action = pagina;
                    marcoArticulosBorrar.document.ArticulosDePromocion.submit();
                }

                document.promociones.checkb.checked = true;
                break;

            case "BAJAGRUPO":
                mod = "save2";
                if (sentido == "first") {
                    sentido = "&submode=first";
                    mod = "first4";
                }

                if (sentido == "insertar") {
                    sentido = "&submode=insertar";
                    mod = "insertar";
                }

                pagina = "ArticulosDePromocion.asp?mode=" + mod + "&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.promociones.htarifa.value +
			        "&bref=" + document.promociones.brefcontiene.value + "&bfamilia=" + document.promociones.familia1.value + "&bcategoria=" +
			        document.promociones.categoria1.value + "&bfamilia_padre=" + document.promociones.familia_padre1.value + "&btipoarticulo=" +
			        document.promociones.btipoarticulo.value + "&bdesc=" + document.promociones.bdescontiene.value + "&bnproveedor=" +
			        document.promociones.bproveedor.value + "&queordenar=" + queordenar + "&comoordenar=" + comoordenar +
			        "&brangodesde=" + document.promociones.brangodesde.value + "&brangohasta=" + document.promociones.brangohasta.value +
			        "&bpvpiva=" + document.promociones.bpvpiva.value + "&viene=BAJAGRUPO" + "&bcodigog=" + document.promociones.bcodigog.value +
                    "&bdescripciong=" + document.promociones.bdescripciong.value;

                if (mod == "first4") {
                    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    document.getElementById("frArticulosBorrar").src = pagina;
                } else {
                    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    marcoArticulosBorrar.document.ArticulosDePromocion.action = pagina;
                    marcoArticulosBorrar.document.ArticulosDePromocion.submit();
                }

                document.promociones.checkb.checked = true;
                break;

        }
    }

    function Editar(p_codigo, p_npagina, p_campo, p_criterio, p_texto) {
        document.location = "promociones.asp?mode=edit&p_codigo=" + p_codigo + "&npagina=" + p_npagina + "&campo=" + p_campo + "&texto=" + p_texto + "&criterio=" + p_criterio;
        parent.botones.document.location = "promociones_bt.asp?mode=edit";
    }
    function TraerArticulo(mode,p_codigo,p_npagina) {
        if (mode == "add") {
            if (document.promociones.referencia.value != "") {
                document.promociones.action = "promociones.asp?mode=" + mode + "&submode=traerarticulo&npagina=" + p_npagina + "&i_codigo=" + document.promociones.i_codigo.value + "&i_descripcion=" + document.promociones.i_descripcion.value +
                    "&i_descripcion_tpv=" + document.promociones.i_descripcion_tpv.value + "&i_v_from=" + document.promociones.i_v_from.value + "&i_v_to=" + document.promociones.i_v_to.value +
                    "&i_qt_total=" + document.promociones.i_qt_total.value + "&i_qt_dtos=" + document.promociones.i_qt_dtos.value + "&i_dto=" + document.promociones.i_dto.value +
                    "&i_dto2=" + document.promociones.i_dto2.value + "&i_TypePromotion=" + document.promociones.i_TypePromotion.value + "&referencia=" + document.promociones.referencia.value;
                document.promociones.submit();
            }
        } else {
            if (document.promociones.referencia.value != "") {
                document.promociones.action = "promociones.asp?mode=" + mode + "&p_codigo=" + p_codigo +"&submode=traerarticulo&npagina=" + p_npagina + "&e_codigo=" + document.promociones.e_codigo.value + "&e_descripcion=" + document.promociones.e_descripcion.value +
                    "&e_descripcion_tpv=" + document.promociones.e_description_tpv.value + "&e_v_from=" + document.promociones.e_v_from.value + "&e_v_to=" + document.promociones.e_v_to.value +
                    "&e_qt_total=" + document.promociones.e_qt_total.value + "&e_qt_dtos=" + document.promociones.e_qt_dtos.value + "&e_dto=" + document.promociones.e_dto.value +
                    "&e_dto2=" + document.promociones.e_dto2.value + "&e_TypePromotion=" + document.promociones.e_TypePromotion.value + "&referencia=" + document.promociones.referencia.value;
                document.promociones.submit();
            }
        }
    }

    function GuardarArticulos(mode) {
        if (mode == "ALTA") {
            marcoArticulosAdd.document.ArticulosDePromocion.action = "ArticulosDePromocion.asp?mode=save&submode=all&tarifa=" + document.promociones.htarifa.value
            marcoArticulosAdd.document.ArticulosDePromocion.submit();
        }
    }

    function GuardarCondicion(mode) {
        if (mode == "ALTA") {
            marcoCondicionesAdd.document.CondicionesPromocion.action = "CondicionesPromocion.asp?mode=save&tarifa=" + document.promociones.htarifa.value;
            marcoCondicionesAdd.document.CondicionesPromocion.submit();
        }
    }

    function BorrarArticulos() {
        if (confirm("<%=LitMsgEliminarRefTarifaConfirm%>")) {
            marcoArticulosBorrar.document.ArticulosDePromocion.action = "ArticulosDePromocion.asp?mode=delete&npagina=" + marcoArticulosBorrar.document.ArticulosDePromocion.hnpagina.value + "&tarifa=" + document.promociones.htarifa.value;
            marcoArticulosBorrar.document.ArticulosDePromocion.submit();
        }
    }
    function WinArticulos() {
        Ven = AbrirVentana("articulosTienda_buscar.asp?ndoc=promociones&titulo=<%=LitSelArticulo%>&mode=search", "P",<%=AltoVentana %>,<%=AnchoVentana %>);

    }

    function seleccionar(marco, formulario, check) {
        nregistros = eval(marco + ".document." + formulario + ".hNregs.value-1");
        if (eval("document.promociones." + check + ".checked")) {
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
            }
        }
        else {
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
            }
        }
    }
    function typeDevelopmentonchange(val,num) {
        if (val == 0) {
            document.getElementById("id_Reward").style = "display:none";
            document.getElementById("c_Reward").style = "display:none";
            document.getElementById("c_qt_dtos").style = "display:";
            document.getElementById("id_qt_dtos").style = "display:";
            document.getElementById("c_dto").style = "display:";
            document.getElementById("id_dto").style = "display:";
            document.getElementById("c_dto2").style = "display:";
            document.getElementById("id_dto2").style = "display:";
        }
        else {
            document.getElementById("id_Reward").style = "display:";
            document.getElementById("c_Reward").style = "display:";
            document.getElementById("c_qt_dtos").style = "display:none";
            document.getElementById("id_qt_dtos").style = "display:none";
            document.getElementById("c_dto").style = "display:none";
            document.getElementById("id_dto").style = "display:none";
            document.getElementById("c_dto2").style = "display:none";
            document.getElementById("id_dto2").style = "display:none";
        }
    }
    function isTextSelected(input) 
    {
        if (typeof input.selectionStart == "number") {
            return input.selectionStart == 0 &&
            input.selectionEnd == input.value.length;
        } else if (typeof document.selection != "undefined") {
            input.focus();
            return document.selection.createRange().text == input.value;
        }
    }

    // ############################ FUNCIONES AJAX ############################

    function handleHttpResponse() {
        promociones.document.getElementById("waitBoxOculto").style.visibility = "visible";
        if (http.readyState == 4) {
            if (http.status == 200) {
                if (http.responseText.indexOf('invalid') == -1) {
                    results = http.responseText;
                    enProceso = false;
                    if (results != "OK") {
                        alert("<%=LITMSG_GRUPOSACTUALIZADO_ERROR %>");
                    }
                    else {
                        alert("<%=LITMSG_GRUPOSACTUALIZADO_OK %>");
                    }
                }
            }
            else {
                alert("<%=LITMSG_GRUPOSACTUALIZADO_ERROR %>");
            }
            promociones.document.getElementById("waitBoxOculto").style.visibility = "hidden";
        }
    }

    function ActualizarGrupos() {
        if (!enProceso && http) {
            var timestamp = Number(new Date());
            var url = "promociones.asp?mode=consultaAJAX&consulta=actualizarGrupos&ts=" + timestamp;
            http.open("GET", url, false);
            http.onreadystatechange = handleHttpResponse;
            enProceso = true;
            http.send(null);
        }
    }

    function getHTTPObject() {
        var xmlhttp;
        if (!xmlhttp && typeof XMLHttpRequest != 'undefined') {
            try {
                xmlhttp = new XMLHttpRequest();
            }
            catch (e) { xmlhttp = false; }
        }
        return xmlhttp;
    }

    var enProceso = false; // lo usamos para ver si hay un proceso activo
    var http = getHTTPObject(); // Creamos el objeto XMLHttpRequest

    // ############################ FIN DE AJAX ############################
</script>
<body class="BODY_ASP">
<%
'******************************************************************************************************************
'                                             FUNCIONES ASP
'******************************************************************************************************************
sub BarraNavegacion()
    %>
        <script language="javascript" type="text/javascript">
            jQuery("#S_CABECERA").hide();
            jQuery("#ARTICULOS").show();
        </script>
    <%

end sub

'******************************************************************************************************************
sub SpanCondiciones(tarifa)
	%><iframe name="marcoCondicionesAdd" id="frCondicionesAdd" src="CondicionesPromocion.asp?mode=show&tarifa=<%= enc.EncodeForHtmlAttribute(null_s(tarifa)) %>" style="margin: 0px; padding: 0px; width: 100%; border: 0px none;"></iframe><%
end sub

'******************************************************************************************************************

sub SpanCondicionesAplicacion(tarifa)
    'Divisa= d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente"))
    strselect= "select abreviatura from divisas where codigo like ?+'%' and moneda_base<>0 "
     Divisa= DLookupP1(strselect, session("ncliente")&"", adVarChar, 15, session("dsn_cliente"))

    DrawDiv "3","",""
    DrawLabel "","display:inline-block;margin-right: 4px;", LIT_LIMIT_CASH2
    DrawInput "","","impMinimo",enc.EncodeForHtmlAttribute(null_s(importeE)),""
    DrawSpan "CELDA","",Divisa & " " & LITIVAINC & " <b>" & LIT_LIMIT_CASH3 & "</b>",""
    CloseDiv
end sub

sub SpanAltasArticulos(tar)
	dis=""
	'TieneArticulos=d_lookup("referencia","articulos_grupos_oferta","codigo='" & tar & "'",session("dsn_cliente"))&""
    TieneArticulosSelect= "Select referencia from articulos_grupos_oferta where codigo=?"
    TieneArticulos= DLookupP1(TieneArticulosSelect, tar&"", adVarChar,10,  session("dsn_cliente"))&""

	if TieneArticulos="" then dis="disabled"
	'Línea para establecer los parámetros de relleno de iframe
        EligeCelda "input-detail",mode,"left","","",0,LitRefContiene,"refcontiene",20,""
        EligeCelda "input-detail",mode,"left","","",0,LitDesContiene,"descontiene",20,""

		strselect ="select * from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='ARTICULO' order by descripcion "
		rstAux.cursorlocation=3
		rstAux.open strselect, session("dsn_cliente")

        DrawDiv "1-detail", "", ""
			DrawLabel "", "", LitTipoArt%>
		    <select style="display:; width:175px;" class="CELDAL7" name="tipoarticulo">
			    <%if tipoarticulo="" then %>
			    	<option selected="selected" value=""> </option>
				<%else%>
				    <option selected="selected" value="<%=enc.EncodeForHtmlAttribute(tipoarticulo)%>"> <%=trimCodEmpresa(tipoarticulo)%></option>
			    	<option value=""> </option>
				<%end if
			 	while not rstAux.eof%>
		   			<option value="<%=rstAux("codigo")%>"><%=enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion")))%></option>
					<%rstAux.movenext
				wend%>
		   	</select></td>
			<%
        CloseDiv
        rstAux.close

			dim ConfigDespleg (3,13)
			i=0
			ConfigDespleg(i,0)="categoria"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
		    ConfigDespleg(i,9)=LitCategoria
			ConfigDespleg(i,10)=categoria
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=1
			ConfigDespleg(i,0)="familia_padre"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitFamilia
			ConfigDespleg(i,10)=familia_padre
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""
 
			i=2
			ConfigDespleg(i,0)="familia"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitSubFamilia
			ConfigDespleg(i,10)=familia
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			DibujaDesplegablesDetail ConfigDespleg,session("dsn_cliente")
                
            DrawDiv "1-detail", "", ""
            DrawLabel "", "", LitImporteRango
            DrawSpan "CELDA", "", "", ""
            CloseDiv
            DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitDesde & " " & LITFORMATDATE
                DrawInput "rangodesde", "","rangodesde","",""
                DrawCalendar "rangodesde"
            CloseDiv
            DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitHasta & " " & LITFORMATDATE
                DrawInput "rangohasta", "","rangohasta","",""
                DrawCalendar "rangohasta"
            CloseDiv%><script type="text/javascript">
                        $(".rangodesde").keypress(function (e) {
                            if (isTextSelected(this)) {
                                return;
                            }
                            if(e.which !== 8) {
                                if (this.value.length > 9)
                                    return false;
                                var numChars = $(this).val().length;
                                if(numChars === 2 || numChars === 5) {
                                    var thisVal = $(this).val();
                                    thisVal += '/';
                                    $(this).val(thisVal);
                                }
                            }
                        });
                            $(".rangohasta").keypress(function (e) {
                            if (isTextSelected(this)) {
                                return;
                            }
                            if(e.which !== 8) {
                                if (this.value.length > 9)
                                    return false;
                                var numChars = $(this).val().length;
                                if(numChars === 2 || numChars === 5) {
                                    var thisVal = $(this).val();
                                    thisVal += '/';
                                    $(this).val(thisVal);
                                }
                            }
                        });
                       </script><%
            DrawDiv "1-detail", "", ""
                DrawLabel "", "", litPvpIva
                %><input type="checkbox" name="pvpiva"><%
            CloseDiv
                    
            DrawDiv "1-detail", "", ""
            DrawLabel "", "", LitProveedor
            set rstAux = Server.CreateObject("ADODB.Recordset")
		 	    'rstAux.cursorlocation=3
		 	    'rstAux.open "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like '" & session("ncliente") & "%' order by razon_social",session("dsn_cliente")
            strselect= "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like ? + '%' order by razon_social"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,session("ncliente")&"")
            set rstAux= command2.Execute

            DrawSelect "width60","","proveedor",rstAux,"","nproveedor","razon_social","",""
            CloseDiv
            
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing	

                    
        DrawDiv "1-detail", "", ""
        DrawLabel "", "", LitCargarArticulos%>
			<a class="ic-accept" href="javascript:if(Insertar('ALTA','1','first','A.referencia','asc'));"><img src="<%=themeIlion %><%=ImgAplicar%>"<%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>"></a>
		<%CloseDiv
            
        DrawDiv "1-detail", "", ""
            DrawLabel "", "", LitCodigo
            DrawInput "", "","codigog","",""      
        CloseDiv
        DrawDiv "1-detail", "", ""
			DrawLabel "", "", LitDescripcion
            DrawInput "", "","descripciong","",""      
        CloseDiv

        DrawDiv "1-detail", "", ""
            DrawLabel "", "", LitCargarGrupos%><a class="ic-accept" href="javascript:if(Insertar('ALTAGRUPO','1','first','A.referencia','asc'));"><img src="<%=themeIlion %><%=ImgAplicar%>"<%=ParamImgAplicar%> alt="<%=LitCargarGrupos%>" title="<%=LitCargarGrupos%>"/></a>
		<%CloseDiv
        %>
	<table class="width100 md-table-responsive">
	    <%escribe1="&darr;"
	    escribe2="&harr;"
	    escribe3="&harr;"
        colorflecha="blue"
						
		Drawfila color_terra
			%><td class="CELDAC7 underOrange width5"><input type="Checkbox" name="check" onclick="seleccionar('marcoArticulosAdd','ArticulosDePromocion','check');" /></td>
			<td class="CELDAC7 underOrange width15">
			    <b><%=LitReferencia%></b>
			    <a class="CELDAREF7" style="font-size: larger; display: inline-block;" color: <%=colorflecha%>;" id="OD1A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDePromocion','A.referencia');" title="<%=LitOrdTarRef & " " & LitOrdSentidoD%>"><b><%=escribe1%></b></a>
			</td>
			<td class="CELDAC7 underOrange width25">
			    <b><%=LitDescripcion%></b>
			    <a class="CELDAREF7" style="font-size: larger; display: inline-block;" color: <%=colorflecha%>;" id="OD2A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDePromocion','A.nombre');" title="<%=LitOrdTarDesc%>" ><b><%=escribe2%></b></a>
			</td>
			<td class="CELDAC7 underOrange width20">
			    <b><%=LitSubFamilia%></b>
			    <a class="CELDAREF7" style="font-size: larger; display: inline-block;" color: <%=colorflecha%>;" id="OD3A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDePromocion','F.nombre');" title="<%=LitOrdTarSubf%>" ><b><%=escribe3%></b></a>
			</td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvp%></b></td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvpIva%></b></td>
		<%CloseFila
	%></table>
	<iframe name="marcoArticulosAdd" id='frArticulosAdd' src='ArticulosDePromocion.asp'  class="width100 md-table-responsive" <!--width='<% response.Write(iif(si_tiene_modulo_credito,"815","715"))  %>'-->></iframe>
	<table style="width:100%;" border="0" cellpadding="0" cellspacing="0"><%
		DrawFila ""
			%><td class="CELDA7" style="width: 140px;">
				<div align="left" valign="center" id="Nregs" style="width: 140px; font-weight: bold;"></div>
			</td>
			<td class="CELDAC7" width="890">
			</td>
			<td class="CELDAR7" width="20">
			<div id="IcoIns" style="visibility: hidden;">
			<a href="javascript:if(GuardarArticulos('ALTA'));"><img src="<%=themeIlion %><%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitGuardarArt%>"></a>
			</div></td><%
		CloseFila
	%></table><%
end sub

'**********************************************************************************************************
sub SpanBajasArticulos()
	'Línea para establecer los parámetros de relleno de iframe
        EligeCelda "input-detail",mode,"left","","",0,LitRefContiene,"brefcontiene",20,""
        EligeCelda "input-detail",mode,"left","","",0,LitDesContiene,"bdescontiene",20,""

		'strselect ="select * from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='ARTICULO' order by descripcion "
		'rstAux.cursorlocation=3
		'rstAux.open strselect, session("dsn_cliente")
        strselect ="select * from tipos_entidades with(nolock) where codigo like ? + '%' and tipo=? order by descripcion "
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo",adVarChar,adParamInput,50,"ARTICULO")
                
            set rstAux= command2.Execute

        DrawDiv "1-detail", "", ""
			DrawLabel "", "", LitTipoArt%>
		    <select style="display:;width:175px;" class="CELDAL7" name="btipoarticulo">
			    <%if tipoarticulo="" then %>
			    	<option selected="selected" value=""> </option>
				<%else%>
				    <option selected="selected" value="<%=enc.EncodeForHtmlAttribute(tipoarticulo)%>"> <%=trimCodEmpresa(tipoarticulo)%></option>
			    	<option value=""> </option>
				<%end if
			 	while not rstAux.eof%>
		   			<option value="<%=enc.EncodeForHtmlAttribute(null_s(rstAux("codigo")))%>"><%=enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion")))%></option>
					<%rstAux.movenext
				wend%>
		   	</select> <%
        CloseDiv
        conn2.close
        set conn2    =  nothing
        set command2 =  nothing
        set rstAux  =  nothing
        'rstAux.close

			dim ConfigDespleg (3,13)

			i=0
			ConfigDespleg(i,0)="categoria1"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitCategoria
			ConfigDespleg(i,10)=categoria
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=1
			ConfigDespleg(i,0)="familia_padre1"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitFamilia
			ConfigDespleg(i,10)=familia_padre
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=2
			ConfigDespleg(i,0)="familia1"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitSubFamilia
			ConfigDespleg(i,10)=familia
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			DibujaDesplegablesDetail ConfigDespleg,session("dsn_cliente")
		    
            %><table></table><%
            
            DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitImporteRango
                DrawSpan "CELDA","","",""
            CloseDiv
            DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitDesde  & " " & LITFORMATDATE
                DrawInput "brangodesde", "","brangodesde","",""
                DrawCalendar "brangodesde"
            CloseDiv
            DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitHasta & " " & LITFORMATDATE
                DrawInput "brangohasta", "","brangohasta","",""
                DrawCalendar "brangohasta"
            CloseDiv%><script type="text/javascript">
                        $(".brangodesde").keypress(function (e) {
                            if (isTextSelected(this)) {
                                return;
                            }
                            if(e.which !== 8) {
                                if (this.value.length > 9)
                                    return false;
                                var numChars = $(this).val().length;
                                if(numChars === 2 || numChars === 5) {
                                    var thisVal = $(this).val();
                                    thisVal += '/';
                                    $(this).val(thisVal);
                                }
                            }
                        });
                        $(".brangohasta").keypress(function (e) {
                        if (isTextSelected(this)) {
                            return;
                        }
                        if(e.which !== 8) {
                            if (this.value.length > 9)
                                return false;
                            var numChars = $(this).val().length;
                            if(numChars === 2 || numChars === 5) {
                                var thisVal = $(this).val();
                                thisVal += '/';
                                $(this).val(thisVal);
                            }
                        }
                    });</script><%
            DrawDiv "1-detail", "", ""
                DrawLabel "", "", litPvpIva
                %><input type="checkbox" name="bpvpiva"><%
            CloseDiv 

            DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitProveedor
		 	    'rstAux.cursorlocation=3
		 	    'rstAux.open "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like '" & session("ncliente") & "%' order by razon_social",session("dsn_cliente")
            strselect= "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like ? + '%' order by razon_social"
            set rstAux = Server.CreateObject("ADODB.Recordset")
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,session("ncliente")&"")
            
            set rstAux= command2.Execute
                DrawSelect "width60","","bproveedor",rstAux,"","nproveedor","razon_social","",""
			'rstAux.close
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing
        CloseDiv

        DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitCargarArticulos%>
			<a class="ic-accept" href="javascript:if(Insertar('BAJA','1','first','A.referencia','asc'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>"></a>
		<%CloseDiv

        DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitCodigo
            DrawInput "", "","bcodigog","",""      
        CloseDiv

        DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitDescripcion
            DrawInput "", "","bdescripciong","",""      
        CloseDiv

        DrawDiv "1-detail", "", ""
                DrawLabel "", "", LitCargarGrupos%>
			<a class="ic-accept" href="javascript:if(Insertar('BAJAGRUPO','1','first','A.referencia','asc'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>" title="<%=LitCargarArticulos%>"/></a>
		<%CloseDiv
        %>
	<table class="width100 md-table-responsive">
        <%escribe1="&darr;"
        escribe2="&harr;"
        escribe3="&harr;"
        colorflecha="blue"

		Drawfila color_terra%>
			<td class="CELDAC7 underOrange width5"><input type="Checkbox" name="checkb" onclick="seleccionar('marcoArticulosBorrar','ArticulosDePromocion','checkb');"/></td>
			<td class="CELDAC7 underOrange width15">
			    <b><%=LitReferencia%></b>
			    <a class="CELDAREF7" style="font-size: larger; display: inline-block;" color: <%=colorflecha%>;" id="OD1B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDePromocion','A.referencia');" title="<%=LitOrdTarRef & " " & LitOrdSentidoD%>"><b><%=escribe1%></b></a>
			</td>
			<td class="CELDAC7 underOrange width25">
			    <b><%=LitDescripcion%></b>
			    <a class="CELDAREF7" style="font-size: larger; display: inline-block;" color: <%=colorflecha%>;" id="OD2B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDePromocion','A.nombre');" title="<%=LitOrdTarDesc%>" ><b><%=escribe2%></b></a>
			</td>
			<td class="CELDAC7 underOrange width20">
			    <b><%=LitSubFamilia%></b>
			    <a class="CELDAREF7" style="font-size: larger; display: inline-block;" color: <%=colorflecha%>;" id="OD3B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDePromocion','F.nombre');" title="<%=LitOrdTarSubf%>" ><b><%=escribe3%></b></a>
			</td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvp%></b></td>
			<td class="CELDAC7 underOrange width10"><b><%=LitPvpIva%></b></td>
		<%CloseFila%>
	</table>
	<iframe name="marcoArticulosBorrar" id='frArticulosBorrar' src='ArticulosDePromocion.asp' class="width100 md-table-responsive"<!--width='<% response.Write(iif(si_tiene_modulo_credito,"815","715"))  %>'-->></iframe>
	<table style="width:100%;" cellpading="0" cellspacing="0"><%
		DrawFila ""
			%><td class="CELDA7" style="width: 140px;">
				<div align="left" id="NregsB" style="width: 140px; font-weight: bold;"></div>
			</td>
			<td class="CELDAC7" width="890">
			</td>
			<td class="CELDAR7" width="48">
				<!--<div id="IcoBorrModif" style="visibility: hidden;"><a href="javascript:if(GuardarArticulos('BAJA'));"><img src="../images/<%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitGuardarArt%>"></a>&nbsp;-->
				<a class="ic-delete" href="javascript:if(BorrarArticulos());"><img src="<%=themeIlion %><%=ImgEliminar%>" <%=ParamImgEliminar%> alt="<%=LitEliminarArt%>" title="<%=LitEliminarArt%>"/></div></a>
			</td><%
		CloseFila
	%></table><%
end sub

'**************************************************************************************************
'                                   Código principal de la página
'**************************************************************************************************

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
   %><form name="promociones" method="post" action="promociones.asp"><%

	si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	'JMMMM - 28/01/2010 --> Se añade modulo línea de crédito
	si_tiene_modulo_credito=ModuloContratado(session("ncliente"),ModLineaCredito)
	' JMMM - 03/11/2010 -> Franquicias
	si_tiene_modulo_franquicia = ModuloContratado(session("ncliente"),ModFranquiciasTiendas)
	'esFranquiciador=d_lookup("franquiciador", "configuracion", "nempresa='"&session("ncliente")&"'", session("dsn_cliente"))
    esFranquiciadorSelect= "select franquiciador from configuracion with(nolock) where nempresa =?"
    esFranquiciadorReal=DlookupP1(esFranquiciadorSelect, session("ncliente")&"",adVarChar, 5, session("dsn_cliente"))


	set rst = server.CreateObject("ADODB.Recordset")
	set rstAux = server.CreateObject("ADODB.Recordset")
	set rstAux2 = server.CreateObject("ADODB.Recordset")
	set rstAux3 = server.CreateObject("ADODB.Recordset")
    set rst2 = server.CreateObject("ADODB.Recordset")
    
    'Leer parámetros de la página
    mode=enc.EncodeForJavascript(request.querystring("mode"))
	mode2=enc.EncodeForJavascript(request.querystring("mode2"))
    submode = enc.EncodeForJavascript(request.querystring("submode"))


    'Parametros Inserción
    if Request.QueryString("i_codigo")>"" then
       codigoI=limpiaCadena(Request.QueryString("i_codigo"))
    else
        codigoI=left(limpiaCadena(Request.Form("i_codigo")),5)
    end if
    if Request.QueryString("i_descripcion")>"" then
        descripcionI=limpiaCadena(request.QueryString("i_descripcion"))
    else
        descripcionI=limpiaCadena(request.form("i_descripcion"))
    end if
    if Request.QueryString("i_descripcion_tpv")>"" then
        descripcion_tpvI=limpiaCadena(request.QueryString("i_descripcion_tpv"))
    else
        descripcion_tpvI=nulear(limpiaCadena(request.form("i_descripcion_tpv")))
    end if
    if Request.QueryString("i_v_from")>"" then
        v_fromI=limpiaCadena(request.QueryString("i_v_from"))
    else
        v_fromI=nulear(limpiaCadena(request.form("i_v_from")))
    end if
    if Request.QueryString("i_v_to")>"" then
        v_toI=limpiaCadena(request.QueryString("i_v_to"))
    else
        v_toI=nulear(limpiaCadena(request.form("i_v_to")))
    end if
    if Request.QueryString("i_qt_total")>"" then
        qt_totalI=limpiaCadena(request.QueryString("i_qt_total"))
    else
        qt_totalI=nulear(limpiaCadena(request.form("i_qt_total")))
    end if
    if Request.QueryString("i_qt_dtos")>"" then
        qt_dtosI=limpiaCadena(request.QueryString("i_qt_dtos"))
    else
        qt_dtosI=nulear(limpiaCadena(request.form("i_qt_dtos")))
    end if
    if Request.QueryString("i_dto")>"" then
        dtoI=limpiaCadena(request.QueryString("i_dto"))
    else
        dtoI=nulear(limpiaCadena(request.form("i_dto")))
    end if
    if Request.QueryString("i_dto2")>"" then
        dto2I=limpiaCadena(request.QueryString("i_dto2"))
    else
        dto2I=nulear(limpiaCadena(request.form("i_dto2")))
    end if
    TypePromotionI = 0
    if Request.QueryString("i_TypePromotion")>"" then
        TypePromotionI=limpiaCadena(request.QueryString("i_TypePromotion"))
    else
        TypePromotionI=nulear(limpiaCadena(request.form("i_TypePromotion")))
    end if
    if Request.QueryString("referencia")>"" then
        refI=limpiaCadena(request.QueryString("referencia"))
    else
        refI=nulear(limpiaCadena(request.form("referencia")))
    end if

    If dtoI&""<>"" and dto2I&""="" then
        typeI = 1
        dscntoI = dtoI
    else if dtoI&""="" and dto2I&""<>"" then
        typeI = 0
        dscntoI = dto2I
        End If
    End If

    'Parametro Edicion
    if Request.QueryString("e_codigo")>"" then
       codigoE=limpiaCadena(Request.QueryString("e_codigo"))
    else
        codigoE=limpiaCadena(Request.Form("e_codigo"))
    end if
	CheckCadena codigoE
    if Request.QueryString("e_descripcion")>"" then
       descripcionE=limpiaCadena(Request.QueryString("e_descripcion"))
    else
        descripcionE=limpiaCadena(Request.Form("e_descripcion"))
    end if
    if Request.QueryString("impMinimo")>"" then
       impMinimoE=trim(limpiaCadena(Request.QueryString("impMinimo")))
    else
        impMinimoE=trim(limpiaCadena(Request.Form("impMinimo")))
    end if
    if Request.QueryString("e_observaciones")>"" then
       observacionesE=limpiaCadena(Request.QueryString("e_observaciones"))
    else
        observacionesE= nulear(limpiaCadena(request.form("e_observaciones")))
    end if
    if Request.QueryString("e_description_tpv")>"" then
       description_tpvE=limpiaCadena(Request.QueryString("e_description_tpv"))
    else
        description_tpvE= nulear(limpiaCadena(request.form("e_description_tpv")))
    end if
    if Request.QueryString("e_v_from")>"" then
       v_fromE=limpiaCadena(Request.QueryString("e_v_from"))
    else
        v_fromE= nulear(limpiaCadena(request.form("e_v_from")))
    end if
    if Request.QueryString("e_v_to")>"" then
       v_toE=limpiaCadena(Request.QueryString("e_v_to"))
    else
        v_toE= nulear(limpiaCadena(request.form("e_v_to")))
    end if
    if Request.QueryString("e_qt_total")>"" then
       qt_totalE = limpiaCadena(Request.QueryString("e_qt_total"))
    else
        qt_totalE = nulear(limpiaCadena(request.form("e_qt_total")))
    end if
    if Request.QueryString("e_qt_dtos")>"" then
       qt_dtosE = limpiaCadena(Request.QueryString("e_qt_dtos"))
    else
        qt_dtosE = nulear(limpiaCadena(request.form("e_qt_dtos")))
    end if
    if Request.QueryString("e_dto")>"" then
       dtoE = limpiaCadena(Request.QueryString("e_dto"))
    else
        dtoE = nulear(limpiaCadena(request.form("e_dto")))
    end if
    if Request.QueryString("e_dto2")>"" then
       dto2E = limpiaCadena(Request.QueryString("e_dto2"))
    else
        dto2E = nulear(limpiaCadena(request.form("e_dto2")))
    end if
    if Request.QueryString("e_TypePromotion")>"" then
       TypePromotionE = limpiaCadena(Request.QueryString("e_TypePromotion"))
    else
        TypePromotionE = limpiaCadena(request.Form("e_TypePromotion"))
    end if
    if Request.QueryString("referencia")>"" then
       refE = limpiaCadena(Request.QueryString("referencia"))
    else
        refE = limpiaCadena(request.Form("referencia"))
    end if

    If dtoE&""<>"" and dto2E&""="" then
        typeE = 1
        dscntoE = dtoE
    else if dtoE&""="" and dto2E&""<>"" then
        typeE = 0
        dscntoE = dto2E
        End If
    End If

	condbase=request.form("condbase")
	bcondbase=request.form("bcondbase")

	if condbase&""=""then
		condbase="0"
	end if
	if bcondbase&""=""then
		bcondbase="0"
	end if
	
	WaitBoxOculto LitEsperePorFavor
		
    mode_acceso=mode        		
    if mode_acceso & ""="" then
        mode_acceso="add"
    end if
    if mode_acceso & ""="delete" then
        mode_acceso="add"
    end if
    if mode_acceso & ""="save" then
        mode_acceso="add"
    end if
       

    if submode="traerarticulo" then
       if mode="add" then
           strselect ="select nombre from articulos with(nolock) where referencia=? and (TipoProducto<>1 OR TipoProducto is NULL)"
           set command2 = nothing
           set conn2 = Server.CreateObject("ADODB.Connection")
           set command2 =  Server.CreateObject("ADODB.Command")
           conn2.open session("dsn_cliente")
           conn2.cursorlocation=3
           command2.ActiveConnection =conn2
           command2.CommandTimeout = 60
           command2.CommandText=strselect
           command2.CommandType = adCmdText
           command2.Parameters.Append command2.CreateParameter("@referencia",adVarChar,adParamInput,30,session("ncliente")&refI&"")
           
           set rst2= command2.Execute
           if not rst2.EOF then
                nombreI = rst2("nombre")
           else
                if refI&"">"" then%>
                    <script language="javascript" type="text/javascript">
                        window.alert("<%=LITARTICLENOTFOUND%>");
			        </script><%
                    refI = ""
                else
                    nombreI = ""
                end if
           end if
           conn2.close
           set conn2    =  nothing
           set command2 =  nothing
           set rst2  =  nothing
        else if mode="edit" then
           strselect ="select nombre from articulos with(nolock) where referencia=? and (TipoProducto<>1 OR TipoProducto is NULL)"
           set command2 = nothing
           set conn2 = Server.CreateObject("ADODB.Connection")
           set command2 =  Server.CreateObject("ADODB.Command")
           conn2.open session("dsn_cliente")
           conn2.cursorlocation=3
           command2.ActiveConnection =conn2
           command2.CommandTimeout = 60
           command2.CommandText=strselect
           command2.CommandType = adCmdText
           command2.Parameters.Append command2.CreateParameter("@referencia",adVarChar,adParamInput,30,session("ncliente")&refE&"")
           
           set rst2= command2.Execute
           if not rst2.EOF then
                nombreE = rst2("nombre")
           else
                if refE&"">"" then%>
                    <script language="javascript" type="text/javascript">
                        window.alert("<%=LITARTICLENOTFOUND%>");
			        </script><%
                    refE = ""
                else
                    nombreE = ""
                end if
           end if
           conn2.close
           set conn2    =  nothing
           set command2 =  nothing
           set rst2  =  nothing
           end if
        end if
    end if
	%>
    <input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode_acceso)%>"/>
    <script language="javascript" type="text/javascript">
        window.onload = function () {
            mode_acceso = "<%=enc.EncodeForJavascript(mode_acceso)%>";
            try {
                parent.botones.document.opciones.mode.value = mode_acceso;
            }
            catch (e) {
            }
            //window.alert("el mode_acceso es-" + mode_acceso + "-");
            if (mode_acceso == "edit") {
                try {
                    parent.botones.document.getElementById("iddelete").style.display = "";
                }
                catch (e) {
                }
            }
            else {
                try {
                    parent.botones.document.getElementById("iddelete").style.display = "none";
                }
                catch (e) {
                }
            }
        }
    </script>
    <input type="hidden" name="condbase" value="<%=enc.EncodeForHtmlAttribute(condbase)%>"/>
	<input type="hidden" name="bcondbase" value="<%=enc.EncodeForHtmlAttribute(bcondbase)%>"/>
    <%

	if mode="delete" then
		p_codigo=limpiaCadena(request("codigo"))
''ricardo 28-1-2008 solamente se concatenara el nempresa cuando venga del mode=add
        'mmg:evita que casque el CheckCadena y te expulse del sistema
		if mode2="add" then
			p_codigo=session("ncliente")& p_codigo
		end if
	else
		p_codigo=limpiaCadena(request.form("codigo"))
		if p_codigo="" then
			p_codigo=limpiaCadena(request.querystring("codigo"))
		end if
		if p_codigo="" then p_codigo=limpiaCadena(request("p_codigo"))
	end if
	CheckCadena p_codigo
    if submode <> "traerarticulo" then
	    'insertamos si nos llegan los valores
	    if codigoI>"" and descripcionI>"" then
            strSource = "select * from PROMOTIONS where code='" & session("ncliente")&codigoI & "'"
		    rst.Open strSource, session("dsn_cliente"),adOpenKeyset, adLockOptimistic 
		    if rst.EOF then
			    rst.AddNew
			    rst("code")             = session("ncliente")&codigoI   'codigo => FC
			    rst("description")      = descripcionI                  'descripcion => Prom primavera 3x2
			    rst("description_tpv")  = descripcion_tpvI              'descripcion ticket
                rst("v_from")           = v_fromI                       'desde => fecha inicio
                rst("v_to")             = v_toI                         'hasta => fecha fin
                rst("importe_minimo")   = miround(0,dec_prec)           'importe mínimo = 0,00 
                if TypePromotionI = 1 then
                    rst("type_promotion") = 1
			        rst("qt_total")         = qt_totalI                     
			        rst("qt_dtos")          = 1                      
			        rst("dto")              = miround(0,dec_prec)    
                    rst("type")             = 0                        
                    rst("reward") = session("ncliente")&refI
                else
                    rst("type_promotion") = 0
			        rst("qt_total")         = qt_totalI                     'Artículos Necesarios => 3
			        rst("qt_dtos")          = qt_dtosI                      'Con descuento => 1 con descuento
			        rst("dto")              = miround(dscntoI,dec_prec)     'Importe Fijo o Descuento Valor
                    rst("type")             = typeI                         'Si Importe Fijo => 1, Si Descuento Valor => 0
                    rst("reward")           = null
                end if
			    rst.Update%>
			    <script language="javascript" type="text/javascript">
                    document.location = "promociones.asp?npagina=1";
                    parent.botones.document.location = "promociones_bt.asp?mode=add";
			    </script>
		    <%
		    else %>
			    <script language="javascript" type="text/javascript">
			        window.alert("<%=LitMsgCodigoExiste%>");
                    document.location = "promociones.asp?npagina=1";
                    parent.botones.document.location = "promociones_bt.asp?mode=add";
			    </script>
		    <%end if
		    rst.Close
	    end if
    end if
    if submode <> "traerarticulo" then
	    'actualizamos valores
	    if codigoE>"" and descripcionE>"" and mode<>"delete" then
		    rst.Open "select * from PROMOTIONS with(rowlock) where code='" & codigoE & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		    if not rst.EOF then
			    'rst("codigo")           = codigoE
			    rst("description")      = descripcionE
			    rst("description_tpv")  = description_tpvE
                rst("v_from")           = v_fromE
                rst("v_to")             = v_toE
                rst("importe_minimo")   = iif(impMinimoE&"" > "", impMinimoE, 0)
                if TypePromotionE = 1 then
                    rst("type_promotion")   = 1        
                    rst("reward")           = session("ncliente")&refE
			        rst("qt_total")         = qt_totalE                     
			        rst("qt_dtos")          = 1    
			        rst("dto")              = miround(0,dec_prec)    
                    rst("type")             = 0                     
                else
			        rst("qt_total")         = qt_totalE
			        rst("qt_dtos")          = qt_dtosE
                    rst("dto")              = miround(dscntoE,dec_prec)
                    rst("type")             = typeE
                    rst("type_promotion")   = 0   
                    rst("reward")           = null
                end if
			    rst.Update
		    else %>
			    <script language="javascript" type="text/javascript">
			        window.alert("<%=LitMsgCodigoNoExiste%>");
                    document.location = "promociones.asp?npagina=1";
                    parent.botones.document.location = "promociones_bt.asp?mode=add";
			    </script>
		    <%end if
		    rst.Close
	    end if
    end if

	'eliminamos valores
	if mode="delete" and p_codigo>"" then
		'miramos a ver si esta puesta en algun documento
		

        'miramos si existe algun ticket usando la promocion
        strselect = "select * from promotion_ticket where promotionID=?"
        set rstAux4 = Server.CreateObject("ADODB.Recordset")
        set command4 = nothing
        set conn4 = Server.CreateObject("ADODB.Connection")
        set command4 =  Server.CreateObject("ADODB.Command")
        conn4.open session("dsn_cliente")
        conn4.cursorlocation=3
        command4.ActiveConnection =conn4
        command4.CommandTimeout = 60
        command4.CommandText=strselect
        command4.CommandType = adCmdText
        command4.Parameters.Append command4.CreateParameter("@promotionID",adVarChar,adParamInput,10,p_codigo&"")
                
        set rstAux4 = command4.Execute
        if rstAux4.eof then
		    rst.open "DELETE FROM PROMOTIONS_ARTICLE WITH(ROWLOCK) WHERE codprom='" & p_codigo & "' ",session("dsn_cliente"),adUseClient, adLockReadOnly
		    rst.Open "DELETE FROM PROMOTIONS WITH(ROWLOCK) WHERE code='" & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
            %><script language="javascript" type="text/javascript">
                  window.alert("<%=LITMSGPROMBORRAR%>");
			</script><%
		
			%><script language="javascript" type="text/javascript">
			      //window.alert("<%=LitPromNoBorrar%>");
			</script><%
        else
           %><script language="javascript" type="text/javascript">
                 window.alert("<%=LITMSGPROTICKNODELETE%>");
                 document.location = "promociones.asp?mode=edit&p_codigo=<%=enc.EncodeForJavascript(p_codigo)%>";
                 parent.botones.document.location = "promociones_bt.asp?mode=edit";
			</script><%
        end if
	end if


      p_criterio=limpiaCadena(request("criterio"))
      p_campo=limpiaCadena(request("campo"))
      p_texto=limpiaCadena(request.QueryString("texto"))
      p_npagina=limpiaCadena(request.QueryString("npagina"))
                

      if p_texto>"" then
	  	 if p_campo="code" then 
            p2_campo="substring(code,6,10)"          
            c_where=" where " & p2_campo & " "
         else
            c_where=" where description "
        end if
      end if 
                    
      if c_where>"" then
         select case p_criterio
            case "contiene"
               c_where=c_where+ "like '%" & p_texto & "%'"
            case "termina"
               c_where=c_where+ "like '%" & p_texto & "'"
            case "empieza"
               c_where=c_where+ "like '" & p_texto & "%'"
            case "igual"
              c_where=c_where + "='" & p_texto & "'"
         end select
		 c_where=c_where & " and code like '" & session("ncliente") & "%' "
	  else
	  	 c_where=" where code like '" & session("ncliente") & "%' "
      end if
   PintarCabecera "promociones.asp"
   Alarma "promociones.asp" %>
   <hr/>
   <%
    c_select="select * from promotions with(nolock)"

        if c_where>"" then
           c_select=c_select & c_where
        end if

        if p_npagina="" then
           p_npagina=1
        end if

        select case request("pagina")
           case "siguiente"
              p_npagina=p_npagina+1
           case "anterior"
              p_npagina=p_npagina-1
        end select%>
  <input type="hidden" name="h_npagina" value="<%=cstr(p_npagina)%>"/>
	<%
        rst.cursorlocation=3
        rst.Open c_select,session("dsn_cliente")

        if not rst.EOF then
           rst.PageSize=NumReg
           rst.AbsolutePage=p_npagina
        end if

  if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
		 <a class="CABECERA" href="promociones.asp?pagina=anterior&npagina=<%=cstr(p_npagina)%>&campo=<%=enc.EncodeForHtmlAttribute(p_campo)%>&criterio=<%=enc.EncodeForHtmlAttribute(p_criterio)%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
  		 <img src="<%=themeIlion %><%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a>
  	<%end if

    texto=LitPagina & " " & cstr(p_npagina) & " " & LitDe & " " & cstr(rst.PageCount)%>
  	<font class="CELDA"> <%=texto%> </font> <%

     if clng(p_npagina)<rst.PageCount then %>
		<a class="CABECERA" href="promociones.asp?pagina=siguiente&npagina=<%=cstr(p_npagina)%>&campo=<%=enc.EncodeForHtmlAttribute(p_campo)%>&criterio=<%=enc.EncodeForHtmlAttribute(p_criterio)%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
  		<img src="<%=themeIlion %><%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a>
  	<%end if

	%><font class="CELDA">&nbsp;&nbsp; <%=LitPagIrA%> <input class="CELDA" type="text" name="SaltoPagina1" size="2"/>&nbsp;&nbsp;<a class="CELDAREF" href="javascript:IrAPagina(1,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina');"><%=LitIr%></a></font><%
  end if%>


  <%if mode<>"edit" then                'Codigo cuando se va a dar de alta una promocion%>
        <table width="700" border="0" cellspacing="1" cellpadding="1">
        <%
            Drawfila color_terra
                Drawcelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitNBregistro
            CloseFila
        %>
        </table>
		<table class="width90 table-component">
		<tr class="underOrange"><%
        	DrawceldaDet "'CELDA underOrange width5'","", "left", true, LitCodigo
        	DrawceldaDet "'CELDA underOrange width10'","", "left", true, LitDescripcion
            DrawCeldaDet "'CELDA underOrange width10'","", "left", true, LitDescripcionImpresion
            DrawCeldaDet "'CELDA underOrange width10'","", "left", true, LitDesde & " " & LITFORMATDATE
            DrawCeldaDet "'CELDA underOrange width10'","", "left", true, LitHasta & " " & LITFORMATDATE
            DrawCeldaDet "'CELDA underOrange width10'","","left", true, LIT_TYPE_OF_DEVELOPMENT
            DrawCeldaDet "'CELDA underOrange width5' style='display:' id='c_qt_total'","", "left", true, LitArticulosTotales
            If TypePromotionI&"">"" then
                if TypePromotionI = 0  then
                    display = "style='display:none'"
                    display2 = "style='display:;'"
                else 
                    display = "style='display:;'"
                    display2 = "style='display:none'"
                end if
            else
                display = "style='display:none'"
                display2 = "style='display:;'"
            end if%>
            <td class="CELDA underOrange width15" <%=display%> id="c_Reward" colspan="3" height="leftpx"><%=LIT_GIFT_ARTICLE%></td>
            <td class="CELDA underOrange width5" <%=display2%> id="c_qt_dtos" height="leftpx"><%=LitArticulosDescontados%></td>
            <td class="CELDA underOrange width5" <%=display2%> id="c_dto" height="leftpx"><%=LitImporteFijo%></td>
            <td class="CELDA underOrange width5" <%=display2%> id="c_dto2" height="leftpx"><%=LitDescuento%></td>
            </tr>
            <tr class="underOrange">
                <td class="CELDAL7 width5">
                    <input type="text" class="width100" name="i_codigo" maxlength="5" value="<%=enc.EncodeForHtmlAttribute(null_s(codigoI))%>" /></td>
                <td class="CELDAL7 width5">
                    <input type="text" class="width100" maxlength="50" name="i_descripcion" value="<%=enc.EncodeForHtmlAttribute(null_s(descripcionI))%>"/></td>
                <td class="CELDAL7 width5">
                    <textarea class="CELDAL7 width100" name="i_descripcion_tpv" maxlength="100" value="<%=descripcion_tpvI%>"><%=enc.EncodeForHtmlAttribute(null_s(descripcion_tpvI))%></textarea></td>
                <td class="CELDAL7 width5">
                    <input type="text" class="width65 i_v_from" maxlength="15" name="i_v_from" value="<%=enc.EncodeForHtmlAttribute(null_s(v_fromI))%>"/><%
                    DrawCalendar "i_v_from"
                %></td>
                <td class="CELDAL7 width5">
                    <input type="text" class="width65 i_v_to" maxlength="15" name="i_v_to" value="<%=enc.EncodeForHtmlAttribute(null_s(v_toI))%>"/><%
                    DrawCalendar "i_v_to"
                %></td><script type="text/javascript">
                        $(".i_v_from").keypress(function (e) {
                            if (isTextSelected(this)) {
                                return;
                            }
                            if(e.which !== 8) {
                                if (this.value.length > 9)
                                    return false;
                                var numChars = $(this).val().length;
                                if(numChars === 2 || numChars === 5) {
                                    var thisVal = $(this).val();
                                    thisVal += '/';
                                    $(this).val(thisVal);
                                }
                            }
                        });
                        $(".i_v_to").keypress(function (e) {
                            if (isTextSelected(this)) {
                                return;
                            }
                            if(e.which !== 8) {
                                if (this.value.length > 9)
                                    return false;
                                var numChars = $(this).val().length;
                                if(numChars === 2 || numChars === 5) {
                                    var thisVal = $(this).val();
                                    thisVal += '/';
                                    $(this).val(thisVal);
                                }
                            }
                        });
                       </script>
                <td class="CELDAL7 width5">
                    <select name="i_TypePromotion" class="width80" onchange="typeDevelopmentonchange(this.value,1)">
                        <option <%=iif(TypePromotionI=0,"selected","") %> value="0"><%=LIT_DISCOUNT %></option>
                        <option <%=iif(TypePromotionI=1,"selected","") %> value="1"><%=LIT_GIFT %></option>
                    </select>
                </td>
                <td id="id_qt_total" class="CELDAL7 width5" style="display:;" >
                    <input type="text" class="width100" name="i_qt_total" maxlength="5" value="<%=qt_totalI%>" />
                 </td>
                <td id="id_Reward" <%=display%> class="CELDAL7 width15">
                    <input class="width20" type="text" name="referencia" size="25" value="<%=enc.EncodeForHtmlAttribute(null_s(refI)) %>" onchange="TraerArticulo('add','<%=enc.EncodeForJavascript(null_s(p_codigo))%>','<%=enc.EncodeForJavascript(null_s(p_npagina))%>')" />
                     <a class="CELDAREFB" href="javascript:WinArticulos()"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
                    <input class="width60" type="text" name="nombre" value="<%=enc.EncodeForHtmlAttribute(null_s(nombreI))%>" />
                  </td>
                <td id="id_qt_dtos" class="CELDAL7 width5" <%=display2%> >
                    <input type="text" class="width100" name="i_qt_dtos" maxlength="5" value="<%=enc.EncodeForHtmlAttribute(null_s(qt_dtosI))%>" />
                </td>
                <td id="id_dto" class="CELDAL7 width5" <%=display2%> >
                    <input type="text" class="width100" name="i_dto" maxlength="5" value="<%=enc.EncodeForHtmlAttribute(null_s(dtoI))%>" />
                </td>
                <td id="id_dto2" class="CELDAL7 width5" <%=display2%> >
                    <input type="text" class="width100" name="i_dto2" maxlength="5" value="<%=enc.EncodeForHtmlAttribute(null_s(dto2I))%>" />
                </td>
            </tr>
      </table><%
      if mode="edit" then %>
      <%end if%>
   <%end if

   '***************************************************************************
   'Zona de código para la gestión de artículos de la tarifa
   '***************************************************************************

   if mode="edit" then                  'Codigo cuando se elige una promocion para editarla
   	    BarraNavegacion%>

        <div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['CABECERA', 'ARTICULOS', 'CONDICIONES','CONDICIONESAPLICACION']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['CABECERA', 'ARTICULOS', 'CONDICIONES','CONDICIONESAPLICACION']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
        </div>

        <div class="Section" id="S_CABECERA">
            <a href="#" rel="toggle[CABECERA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader"><%=LitCabecera%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
    <div class="SectionPanel" id="CABECERA" style="display:;">
	    <!--<table class=TBORDE width="100%"><tr><td>-->
		        <input type="hidden" name="htarifa" value="<%=enc.EncodeForHtmlAttribute(p_codigo)%>"/>	
                <table class="width90 table-component">
			      <tr class="underOrange">
			      <%
                    dim color_terrap
                    color_terrap = color_terra

                    if session("version")&"" = "5" then  
                        color_terrap = "#fdfdfd"
                    end if  
                      
                    DrawceldaDet "'CELDA width5'","", "left", true, LitCodigo
        	        DrawceldaDet "'CELDA width10'","", "left", true, LitDescripcion
                    DrawCeldaDet "'CELDA width10'","", "left", true, LitDescripcionImpresion
                    DrawCeldaDet "'CELDA width10'","", "left", true, LitDesde & " " & LITFORMATDATE
                    DrawCeldaDet "'CELDA width10'","", "left", true, LitHasta & " " & LITFORMATDATE
                    DrawCeldaDet "'CELDA width10'","", "left", true, LIT_TYPE_OF_DEVELOPMENT
                    DrawCeldaDet "'CELDA width5' style='display:' id='c_qt_total'","", "left", true, LitArticulosTotales
                    DrawCeldaDet "'CELDA width15' style='display:none' id='c_Reward' colspan='3'","","left",true,LIT_GIFT_ARTICLE
                    DrawCeldaDet "'CELDA width5' style='display:' id='c_qt_dtos'","", "left", true, LitArticulosDescontados
                    DrawCeldaDet "'CELDA width5' style='display:' id='c_dto'","", "left", true, LitImporteFijo
                    DrawCeldaDet "'CELDA width5' style='display:' id='c_dto2'","", "left", true, LitDescuento%>
                   </tr><%
				    par=false
				    i=1
                    if not rst.EOF then
				    ''rst.movefirst
				    while not rst.EOF and i<=NumReg
				       if mode="edit" and p_codigo=rst("code") then
                          importeE = rst("importe_minimo")
                          TmpDescripArtselect= "select nombre from articulos where referencia=?"
						  DrawCeldaDet "'CELDAL7 underOrange width5'","", "center", true, trimCodEmpresa(rst("code"))
						  %><input type="hidden" name="e_codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("code")))%>"/>
                            <td class="CELDAL7 underOrange width10"><%
						        DrawInput "width100","","e_descripcion",enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo",descripcionE,rst("description")))),"maxlength='50'"%>
                            </td>
                            <td class="CELDAL7 underOrange width10">
                                <textarea class="CELDAL7 width100" name="e_description_tpv" value="<%=iif(submode="traerarticulo",description_tpvE,rst("description_tpv"))%>" maxlength="100"><%=enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo",description_tpvE,rst("description_tpv"))))%></textarea>
                            </td>
                            <td class="CELDAL7 underOrange width10"><%
						        DrawInput "'width65 e_v_from'","","e_v_from", enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo", v_fromE,rst("v_from")))),"maxlength='15'"
                                DrawCalendar "e_v_from"%>
                            </td>
                            <td class="CELDAL7 underOrange width10"><%
						        DrawInput "'width65 e_v_to'","","e_v_to", enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo",v_toE,rst("v_to")))),"maxlength='15'"
                                DrawCalendar "e_v_to"%>
                            </td><script type="text/javascript">
                                $(".e_v_from").keypress(function (e) {
                                    if (isTextSelected(this)) {
                                        return;
                                    }
                                    if(e.which !== 8) {
                                        if (this.value.length > 9)
                                            return false;
                                        var numChars = $(this).val().length;
                                        if(numChars === 2 || numChars === 5) {
                                            var thisVal = $(this).val();
                                            thisVal += '/';
                                            $(this).val(thisVal);
                                        }
                                    }
                                });
                                $(".e_v_to").keypress(function (e) {
                                    if (isTextSelected(this)) {
                                        return;
                                    }
                                    if(e.which !== 8) {
                                        if (this.value.length > 9)
                                            return false;
                                        var numChars = $(this).val().length;
                                        if(numChars === 2 || numChars === 5) {
                                            var thisVal = $(this).val();
                                            thisVal += '/';
                                            $(this).val(thisVal);
                                        }
                                    }
                                });
                               </script>
                            <td class="CELDAL7 underOrange width5">
                                <select name="e_TypePromotion" class="width80" onchange="typeDevelopmentonchange(this.value,1)">
                                    <option <%=iif(TypePromotionE="",iif(rst("type_promotion") = 1, "","selected"),iif(TypePromotionE = "1", "","selected"))%> value="0"><%=LIT_DISCOUNT %></option>
                                    <option <%=iif(TypePromotionE="",iif(rst("type_promotion") = 1, "selected",""),iif(TypePromotionE = "1", "selected",""))%> value="1"><%=LIT_GIFT %></option>
                                </select>
                            </td>
                            <td id="id_qt_total" class="CELDAL7 underOrange width5"><%DrawInput "width100","","e_qt_total",enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo", qt_totalE,rst("qt_total")))),"maxlength='5'"%></td>
                            <td id="id_Reward" style="display:none" class="CELDAL7 underOrange width20">
                                <input class="width20" type="text" name="referencia" size="25" value="<%=iif(refE="",trimCodEmpresa(rst("reward")),refE)%>" onchange="TraerArticulo('edit','<%=enc.EncodeForJavascript(null_s(p_codigo))%>','<%=enc.EncodeForJavascript(null_s(p_npagina))%>')" />
                                 <a class="CELDAREFB" href="javascript:WinArticulos()"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
                                <input class="width60" type="text" name="nombre" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(refE="",DLookupP1(TmpDescripArtselect, rst("reward")&"",adVarChar, 30, session("dsn_cliente"))&"",nombreE)))%>"/>
                             </td>
                            <td id="id_qt_dtos" class="CELDAL7 underOrange width5"><%DrawInput "width100","","e_qt_dtos",enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo", qt_dtosE,rst("qt_dtos")))),"maxlength='5'"%></td><%
                             if rst("type")=1 then%>
                                <td id="id_dto" class="CELDAL7 underOrange width5"><%DrawInput "width100","","e_dto",enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo", dtoE,rst("dto")))),"maxlength='5'"%></td>
                                <td id="id_dto2" class="CELDAL7 underOrange width5"><%DrawInput "width100","","e_dto2", "","maxlength='5'"%></td><%
                             else%>
                                <td id="id_dto" class="CELDAL7 underOrange width5"><%DrawInput "width100","","e_dto","","maxlength='5'"%></td>
                                <td id="id_dto2" class="CELDAL7 underOrange width5"><%DrawInput "width100","","e_dto2",enc.EncodeForHtmlAttribute(null_s(iif(submode="traerarticulo", dto2E,rst("dto")))),"maxlength='5'"%></td><%
                             end if
                             if TypePromotionE="" then
                                 if rst("type_promotion") = 1 then
                                       %><script type="text/javascript">
                                             typeDevelopmentonchange(1, 1);
                                         </script><%
                                 end if
                             else
                                 if TypePromotionE = "1" then
                                       %><script type="text/javascript">
                                             typeDevelopmentonchange(1, 1);
                                         </script><%
                                 end if
                             end if
                          CloseFila
				       end if

				       i = i + 1
				       rst.MoveNext
				    wend
				    'rst.Close
                    end if %>
                </table>
            </div>
        </div>
       <div class="Section" id="S_CONDICIONESAPLICACION" >
            <a href="#" rel="toggle[CONDICIONESAPLICACION]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader"><%=LIT_LIMIT_CASH%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
		    <div class="SectionPanel" id="CONDICIONESAPLICACION">
				<% SpanCondicionesAplicacion(p_codigo) %>
            </div>
        </div>
        <div class="Section" id="S_CONDICIONES" style="display: none;">
            <a href="#" rel="toggle[CONDICIONES]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader"><%=LITCONDPROM%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
		    <div class="SectionPanel" id="CONDICIONES">
				<%SpanCondiciones(p_codigo) %>
            </div>
        </div>
        <div class="Section" id="S_ARTICULOS">
            <a href="#" rel="toggle[ARTICULOS]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader"><%=LitArtProm%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
		    <div class="SectionPanel" id="ARTICULOS">
                <div id="tabs" style="display:none">
                    <ul>
                        <li><a href="#tabs1"><%=LitAnadirArticulosGrupos%></a></li>
                        <li><a href="#tabs2"><%=LitBorrModifArticulosGrupos%></a></li>
                    </ul>
                    <div id="tabs1">
			            <%SpanAltasArticulos p_codigo%>
		            </div>
                    <div id="tabs2">
			                <%SpanBajasArticulos%>
		            </div>
                </div>
            </div>
        </div>
<!--    </td></tr></table>-->
	<%end if%>

    <%if mode<>"edit" then 'Valores que se muestran despues de insertar o editar una promocion
        
        %>&nbsp
       <table class="table-component table-responsive width90">
           <tr class="underOrange"><%
        		Drawcelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitCodigo
				Drawcelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitDescripcion
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitDescripcionImpresion
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitDesde  & " " & LITFORMATDATE
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitHasta  & " " & LITFORMATDATE
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'","left", true, LIT_TYPE_OF_DEVELOPMENT
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitArticulosTotales
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitArticulosDescontados
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitImporteFijo
                DrawCelda2 "'CELDA underOrange NO_BORDER_H'", "left", true, LitDescuento
                DrawCelda2 "'CELDA underOrange NO_BORDER_H' style='text-align:left'","left", true, LIT_GIFT_ARTICLE%>
            </tr><%
    end if
      Drawfila ""
        par=false
        i=1

        while not rst.EOF and i<=NumReg
           if mode="edit" and p_codigo=rst("code") then

           elseif mode<>"edit" then
				 'h_ref = "javascript:Editar('" & rst("code") & "'," & p_npagina & ",'" & p_campo & "','" & p_criterio & "','" & p_texto & "');"
				
                h_ref="javascript:Editar('" & enc.EncodeForJavascript(rst("code")) & "'," & _
			                            enc.EncodeForJavascript(p_npagina) & ",'" & _
				  					    enc.EncodeForJavascript(p_campo) & "','" & _
									    enc.EncodeForJavascript(p_criterio) & "','" & _
									    enc.EncodeForJavascript(p_texto) & "');"

                if ucase(rst("code"))<>session("ncliente") & "BASE" then
					if par then
						Drawfila color_blau
						par=false
					else
            			Drawfila color_terra
	            	  	par=true
					end if      
                    DrawCeldahrefTd "CELDAREF valign='top'","left",false,trimCodEmpresa(rst("code")),enc.EncodeForHtmlAttribute(h_ref)
                    DrawCelda2 "CELDA maxlength='10'", "", false, EncodeForHtml(rst("description"))
					DrawCelda2 "CELDA maxlength='10'", "", false, EncodeForHtml(rst("description_tpv"))
                    DrawCelda2 "CELDA maxlength='15'", "", false, EncodeForHtml(rst("v_from"))
                    DrawCelda2 "CELDA maxlength='15'", "", false, EncodeForHtml(rst("v_to"))
                    DrawCelda2 "CELDA maxlength='5'", "", false, iif(rst("type_promotion")=1,LIT_GIFT,LIT_DISCOUNT) 
					DrawCelda2 "CELDA maxlength='5'", "", false, EncodeForHtml(rst("qt_total"))
					DrawCelda2 "CELDA maxlength='5'", "", false, EncodeForHtml(rst("qt_dtos"))
                    if rst("type")=1 then
					    DrawCelda2 "CELDA maxlength='5'", "", false, iif(rst("type_promotion") = 1,"-",EncodeForHtml(rst("dto")))
                        DrawCelda2 "CELDA maxlength='5'", "", false, "-"
                    else
                        DrawCelda2 "CELDA maxlength='5'", "", false, "-"
                        DrawCelda2 "CELDA maxlength='5'", "", false, iif(rst("type_promotion") = 1,"-",EncodeForHtml(rst("dto")))
                    end if
                    DrawCelda2 "CELDA","",false,iif(rst("type_promotion") = 1, d_lookup("nombre","articulos","referencia='" & rst("reward") & "'",session("dsn_cliente"))," - ")
					CloseFila
				end if
           end if

           i = i + 1
           rst.MoveNext
        wend%>
      <%if mode<>"edit" then %>
      </table>
      <%end if%>

    <%if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
		 <a class="CABECERA" href="promociones.asp?pagina=anterior&npagina=<%=cstr(p_npagina)%>&campo=<%=enc.EncodeForHtmlAttribute(p_campo)%>&criterio=<%=enc.EncodeForHtmlAttribute(p_criterio)%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
  		 <img src="<%=themeIlion %><%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a>
  	<%end if

    texto=LitPagina & " " & cstr(p_npagina) & " " & LitDe & " " & cstr(rst.PageCount)%>
  	<font class="CELDA"> <%=texto%> </font> <%

     if clng(p_npagina)<rst.PageCount then %>
		<a class="CABECERA" href="promociones.asp?pagina=siguiente&npagina=<%=cstr(p_npagina)%>&campo=<%=enc.EncodeForHtmlAttribute(p_campo)%>&criterio=<%=enc.EncodeForHtmlAttribute(p_criterio)%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
  		<img src="<%=themeIlion %><%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a>
  	<%end if%>
	<font class="CELDA">&nbsp;&nbsp; <%=LitPagIrA%> <input class="CELDA" type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;<a class="CELDAREF" href="javascript:IrAPagina(2,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina');"><%=LitIr%></a></font><%
	rst.close
  end if%>
   </form>
<%
	set rst = nothing
    set rst2  =  nothing
	set rstAux = nothing
	set rstAux2 = nothing
	set rstAux3 = nothing
    set rstAux4 = nothing
connRound.close
set connRound = Nothing
end if%>
</body>
</html>