<%@ Language=VBScript %><% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  
<script id="DebugDirectives" runat="server" language="javascript">
    // Set these to true to enable debugging or tracing
@set @debug=false
@set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../XSSProtection.inc" -->
<!--#include file="../../ilion.inc" -->

<!--#include file="../../tablas.inc" -->
<!--#include file="../../varios_bt.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../mensajes.inc" -->

<!--#include file="listado_articulos.inc" -->

<!--#include file="../../styles/Master.css.inc" -->
<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    //Validacion de campos
    function ValidarCampos() {
        if (parseInt(parent.pantalla.document.listado_articulos.NumRegs.value)>parseInt(parent.pantalla.document.listado_articulos.maxpdf.value)) {
            window.alert("<%=LitMsgLimitePdf%>");
            return false;
        }

        if (parseInt(parent.pantalla.document.listado_articulos.NumRegs.value)<=0){
            window.alert("<%=LitNoHayRegistrosParaImp%>");
            return false;
        }

        return true;
    }

    function ValidarCampos2(){
        if (parent.pantalla.document.listado_articulos.stockmayoroigual.value==" " || isNaN(parent.pantalla.document.listado_articulos.stockmayoroigual.value.replace(",","."))){
            window.alert("<%=LitMsgStockNumerico%>");
            return false
        }
        if (!checkdate(parent.pantalla.document.listado_articulos.desde_fb)){
            window.alert("<%=LitMsgDesdeFechaBaja%>");
            parent.pantalla.document.listado_articulos.desde_fb.focus();
            return false
        }
        if (!checkdate(parent.pantalla.document.listado_articulos.hasta_fb)){
            window.alert("<%=LitMsgHastaFechaBaja%>");
            parent.pantalla.document.listado_articulos.hasta_fb.focus();
            return false
        }

        if (!checkdate(parent.pantalla.document.listado_articulos.desde_fc)){
            window.alert("<%=LITFORMATOFECHADINCORR%>");
            parent.pantalla.document.listado_articulos.desde_fc.focus();
            return false
        }
        if (!checkdate(parent.pantalla.document.listado_articulos.hasta_fc)){
            window.alert("<%=LITFORMATOFECHAHINCORR%>");
            parent.pantalla.document.listado_articulos.hasta_fc.focus();
            return false
        }
        if(parent.pantalla.document.listado_articulos.lista_obt_objok.value=="1"){
            if(parent.pantalla.document.listado_articulos.tarifa.value=="")
            {
                window.alert("<%=LITFORMATOFILIALINCORR%>");
                parent.pantalla.document.listado_articulos.tarifa.focus();
                return false
            }
        }

        if (DiferenciaTiempo(parent.pantalla.document.listado_articulos.hasta_fb.value, parent.pantalla.document.listado_articulos.desde_fb.value, "dias") < 0) {
            window.alert("<%=LitMsgErrorDate%>");
            parent.pantalla.document.listado_articulos.desde_fb.focus();
            parent.pantalla.document.listado_articulos.desde_fb.select();
            return false;
        }

        if (parent.pantalla.document.listado_articulos.si_campo_personalizables.value==1){
            num_campos=parent.pantalla.document.listado_articulos.num_campos.value;
            respuesta=comprobarCampPerso("parent.pantalla.",num_campos,"listado_articulos");
            if(respuesta!=0){
                titulo="titulo_campo" + respuesta;
                tipo="tipo_campo" + respuesta;
                titulo=parent.pantalla.document.listado_articulos.elements[titulo].value;
                tipo=parent.pantalla.document.listado_articulos.elements[tipo].value;
                if (tipo==4) nomTipo="<%=LitTipoNumericoListArt%>";
                else if (tipo==5) {
                    nomTipo="<%=LitTipoFechaListArt%>";
                }

                window.alert("<%=LitMsgCampoListArt%> " + titulo + " <%=LitMsgTipoListArt%> " + nomTipo);

                return false;
            }
        }

        var cuantos = 0;
        for (var i = 0; opt = parent.pantalla.document.listado_articulos.tarifa.options[i]; i++) {
            if (opt.selected) cuantos++;
        }

        if (cuantos > 10) {
            window.alert("<%=LitTooManySelectedRates%>");
            return false;
        }

        return true;
    }

    //Realizar la acción correspondiente al botón pulsado.
    function Accion(mode,pulsado) {
        switch (mode) {
            case "add":
                switch (pulsado) {
                    case "aceptar": //Aceptar
                        if (ValidarCampos2()){
                            parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            parent.pantalla.document.listado_articulos.action="listado_articulosResultado.asp?mode=ver";
                            parent.pantalla.document.listado_articulos.submit();
                            document.location="listado_articulos_bt.asp?mode=ver";
                        }
                        break;
                    case "cancelar": //Cancelar
                        parent.pantalla.document.listado_articulos.action="listado_articulos.asp?mode=add";
                        parent.pantalla.document.listado_articulos.submit();
                        document.location="listado_articulos_bt.asp?mode=add";
                        break;
                }
                break;
            case "ver":
                switch (pulsado) {
                    case "volver": //Volver atrás
                        parent.pantalla.document.location = "listado_articulos.asp?mode=add";
                        document.location = "listado_articulos_bt.asp?mode=add";
                        break;
                    case "imprimir": //Volver atrás
                        parent.pantalla.focus();
                        parent.pantalla.print(parent.pantalla.document.listado_articulosResultado.apaisado.value, parent.pantalla.document.listado_articulosResultado.maxpagina.value);
                        break;

                    case "imprimirp": //Imprimir Listado en PDF
                        if (parseInt(parent.pantalla.document.listado_articulosResultado.NumRegs.value)>=parseInt(parent.pantalla.document.listado_articulosResultado.maxpdf.value))
                            alert("<%=LitMsgLimitePdf%>");
                        else
                        {
                            
                            pagina="../../crearpdf.asp?apaisado=" + parent.pantalla.document.listado_articulosResultado.apaisado.value + "&destinatario=&ndoc=&tdoc=&dedonde=&empresa=<%=session("ncliente")%>&impusuario=<%=session("usuario")%>&cajaParam=&mode=LISTADO_CLIENTES&url=productos/listados/listado_articulos_pdf.asp";
                                    
                            parent.pantalla.document.location=pagina;
                            document.location="listado_articulos_bt.asp?mode=pdf";
                           
                        }
                        break;
                    case "exportar2":
                        parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";

                        cadena = ""
                        cadena = cadena + "&ver_almacen=" + parent.pantalla.document.listado_articulosResultado.ver_almacen.value;
                        cadena = cadena + "&ver_familia=" + parent.pantalla.document.listado_articulosResultado.ver_familia.value;
                        cadena = cadena + "&ver_coste=" + parent.pantalla.document.listado_articulosResultado.ver_coste.value;
                        cadena = cadena + "&ver_dto=" + parent.pantalla.document.listado_articulosResultado.ver_dto.value;
                        cadena = cadena + "&ver_recargo=" + parent.pantalla.document.listado_articulosResultado.ver_recargo.value;
                        //i(EJM 22/02/2007) Nuevos campos para la importación
                        cadena = cadena + "&ver_CodSub=" + parent.pantalla.document.listado_articulosResultado.ver_CodSub.value;
                        cadena = cadena + "&ver_Embalaje=" + parent.pantalla.document.listado_articulosResultado.ver_Embalaje.value;
                        //fin(EJM 22/02/2007) Nuevos campos para la importación
                        cadena = cadena + "&ver_margen=" + parent.pantalla.document.listado_articulosResultado.ver_margen.value;
                        cadena = cadena + "&ver_pvp=" + parent.pantalla.document.listado_articulosResultado.ver_pvp.value;
                        cadena = cadena + "&ver_iva=" + parent.pantalla.document.listado_articulosResultado.ver_iva.value;
                        cadena = cadena + "&ver_divisa=" + parent.pantalla.document.listado_articulosResultado.ver_divisa.value;
                        cadena = cadena + "&ver_codbarras=" + parent.pantalla.document.listado_articulosResultado.ver_codbarras.value;
                        cadena = cadena + "&ver_stock=" + parent.pantalla.document.listado_articulosResultado.ver_stock.value;
                        cadena = cadena + "&ver_smin=" + parent.pantalla.document.listado_articulosResultado.ver_smin.value;
                        cadena = cadena + "&ver_smax=" + parent.pantalla.document.listado_articulosResultado.ver_smax.value;
                        cadena = cadena + "&ver_reposicion=" + parent.pantalla.document.listado_articulosResultado.ver_reposicion.value;
                        cadena = cadena + "&ver_precibir=" + parent.pantalla.document.listado_articulosResultado.ver_precibir.value;
                        cadena = cadena + "&ver_pservir=" + parent.pantalla.document.listado_articulosResultado.ver_pservir.value;
                        cadena = cadena + "&ver_pmin=" + parent.pantalla.document.listado_articulosResultado.ver_pmin.value;
                        cadena = cadena + "&ver_coste_medio=" + parent.pantalla.document.listado_articulosResultado.ver_coste_medio.value;
                        cadena = cadena + "&ver_pvpiva=" + parent.pantalla.document.listado_articulosResultado.ver_pvpiva.value;
                        cadena = cadena + "&ver_codTerminal=" + parent.pantalla.document.listado_articulosResultado.ver_codTerminal.value;
                        cadena = cadena + "&ver_nomTerminal=" + parent.pantalla.document.listado_articulosResultado.ver_nomTerminal.value;
                        cadena = cadena + "&ver_desAmpliada=" + parent.pantalla.document.listado_articulosResultado.ver_desAmpliada.value;
                        cadena = cadena + "&ver_tipoArticulo=" + parent.pantalla.document.listado_articulosResultado.ver_tipoArticulo.value;
                        cadena = cadena + "&tarifa=" + parent.pantalla.document.listado_articulosResultado.tarifa.value.replace("#coma#", ",");
                        cadena = cadena + "&ver_medida=" + parent.pantalla.document.listado_articulosResultado.ver_medida.value;
                        cadena = cadena + "&ver_peso=" + parent.pantalla.document.listado_articulosResultado.ver_peso.value;
                        cadena = cadena + "&ver_medidaventa=" + parent.pantalla.document.listado_articulosResultado.ver_medidaventa.value;

                        num_campos_perso = parent.pantalla.document.listado_articulosResultado.h_num_campos_articulos.value;
                        for (ki = 1; ki <= num_campos_perso; ki++) {
                            cadena = cadena + "&ver_campo" + ki + "=" + eval("parent.pantalla.document.listado_articulosResultado.elements['ver_campo" + ki + "'].value");
                        }

                        cadena = cadena + "&ver_PLU=" + parent.pantalla.document.listado_articulosResultado.ver_PLU.value;
                        cadena = cadena + "&ver_GRPPLU=" + parent.pantalla.document.listado_articulosResultado.ver_GRPPLU.value;
                        if (pulsado == "exportar2") cadena = cadena + "&forma_exportar=2";
                        else cadena = cadena + "&forma_exportar=1";

                        //parent.pantalla.marcoExportar.document.location = "listado_articulos_exportar.asp?mode=exportar" + cadena;

                        parent.pantalla.document.listado_articulosResultado.action = "listado_articulos_pdf.asp?mode=xls&xls=1&impusuario=<%=session("usuario")%>&empresa=<%=session("ncliente")%>&usuario=<%=session("usuario")%>";
                        parent.pantalla.document.listado_articulosResultado.submit();
                        document.location = "listado_articulos_bt.asp?mode=pdf";
                                    
                        //parent.pantalla.document.location=pagina;
                        //document.location="listado_articulos_bt.asp?mode=pdf";
                        break;
                }
                break;
            case "pdf":
                switch (pulsado) {
                    case "back": //Volver a la pantalla anterior
                        parent.document.location = "../../central.asp?pag1=productos/listados/listado_articulos.asp&pag2=productos/listados/listado_articulos_bt.asp&mode=add";
                        break;
                }
                break;

            case "edit":
                switch (pulsado) {
                    case "cancel": //Volver atrás
                        parent.pantalla.document.listado_articulos.action="listado_articulos.asp?mode=ver";
                        parent.pantalla.document.listado_articulos.submit();
                        document.location="listado_articulos_bt.asp?mode=ver";
                        break;
                    case "save": //Almacenar
                        if(ValidarCampos()){
                            parent.pantalla.document.listado_articulos.action="listado_articulos.asp?mode=save";
                            parent.pantalla.document.listado_articulos.submit();
                            document.location="listado_articulos_bt.asp?mode=ver";
                        }
                        break;
                }

                break;
        }
    }
</script>
<body class="body_master_ASP">
<%mode=enc.EncodeForJavascript(Request.QueryString("mode"))%>
<form name="opciones" method="post">
	<div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
	        <table id="BUTTONS_CENTER_ASP">
		        <tr>
		            <%if mode="add" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('add','aceptar');">
					        <%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('add','cancelar');">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
			        <%elseif mode="ver" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('ver','imprimir');">
					        <%PintarBotonBT LITBOTIMPRIMIR,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('ver','imprimirp');">
					        <%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				        </td>
			            <td class="CELDABOT" onclick="javascript:Accion('ver','volver');">
				            <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			            </td>
			            <td class="CELDABOT" onclick="javascript:Accion('ver','exportar2');">
				            <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			            </td>
			        <%elseif mode="pdf" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					        <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				        </td>
			        <%elseif mode="edit" then%>
				        <td class="CELDABOT" onclick="javascript:Accion('edit','save');">
					        <%PintarBotonBT LITBOTGUARDAR,ImgGuardar,ParamImgGuardar,LITBOTGUARDARTITLE%>
				        </td>
				        <td class="CELDABOT" onclick="javascript:Accion('edit','cancel');">
					        <%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				        </td>
			        <%end if%>
		        </tr>
	        </table>
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