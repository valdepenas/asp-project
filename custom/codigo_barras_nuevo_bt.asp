<%@ Language=VBScript %><% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%> 
<script id='DebugDirectives' runat='server' language='javascript'>
    // Set these to true to enable debugging or tracing
    @set @debug=false
    @set @trace=false
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head> 
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/Master.css.inc" -->
<!--#include file="../controlimpresion.inc" -->
<!--#include file="codigo_barras.inc" -->
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    nodoc = 0
    function printWindowCodBarras(margenizquierdo, margenderecho, margensuperior, margeninferior) {
        /*pasamos los milimetros a pulgadas*/
        mizq = parseFloat(margenizquierdo) * 0.0393700787401575;
        mizq = mizq.toFixed(3);
        msup = parseFloat(margensuperior) * 0.0393700787401575;
        msup = msup.toFixed(3);
        mdcho = parseFloat(margenderecho) * 0.0393700787401575;
        mdcho = mdcho.toFixed(3);
        minf = parseFloat(margeninferior) * 0.0393700787401575;
        minf = minf.toFixed(3);

        factory.printing.portrait = true;
        factory.printing.leftMargin = mizq;
        factory.printing.topMargin = msup;
        factory.printing.rightMargin = mdcho;
        factory.printing.bottomMargin = minf;
        //ricardo 8-9-2009 se comenta la siguiente linea ya que en la nueva version del meadco, se ha de tener licencia para poder utilizar dicha funcion

        factory.printing.Print(true, parent.pantalla.mainFrame);
    }

    //Validacion de campos
    function ValidarCampos() {
        ok = 1;

        cantHMax = 0;
        cantVMax = 0;
        maxpagina = 0;

        //JMA 30/10/05: Pasamos listado_codigo_barras_nuevo.asp a CUSTOM
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("../custom/listado_codigo_barrasn.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form1.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form1.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form1.value;
        }
        //if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("../custom/listado_codigo_barras2n.asp")!=-1){
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras2n.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form2.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form2.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form2.value;
            //rgu cambio para que sean 16 igual que se hizo en el listado_codigo_barras2.asp para mercalopez
            maxpagina = maxpagina - cantHMax
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras4.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form4.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form4.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form4.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("..\\\\..\\\\custom\\\\listado_codigo_barras5.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form5.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form5.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form5.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras6.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form6.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form6.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form6.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras7.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form7.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form7.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form7.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras8.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form8.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form8.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form8.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras9.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form9.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form9.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form9.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value == "listado_codigo_barras9SP.asp") {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form9.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form9.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form9.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras9THT.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form9.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form9.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form9.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras10.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form10.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form10.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form10.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("../custom/listado_codigo_barras11.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form11.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form11.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form11.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barrasCHACAL.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_formCHACAL.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_formCHACAL.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_formCHACAL.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras12.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form12.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form12.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form12.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras13.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form13.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form13.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form13.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras14.asp") != -1) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form14.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form14.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form14.value;
        }

        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras15.asp") != -1) || (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras18.asp") != -1) || (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras29.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form15.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form15.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form15.value;
        }

        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras19.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form19.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form19.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form19.value;
        }

        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras20.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form20.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form20.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form20.value;
        }

        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras21.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form21.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form21.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form21.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras22.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form22.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form22.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form22.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras23.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form23.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form23.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form23.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras24.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form24.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form24.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form24.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras25.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form25.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form25.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form25.value;
        }
        if (parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf('zebra') != -1) {
            if (parent.pantalla.document.codigo_barras.tipodoc.value == "") {
                if (!window.confirm("<%=LitSinSeleccion%>")) return;
                else nodoc = 1;
            }
            maxpagina = 10000;
            cantHMax = 10000;
            cantVMax = 10000;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras26.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form26.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form26.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form26.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras27.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form27.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form27.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form27.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras28.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form28.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form28.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form28.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barras32.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_form32.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_form32.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_form32.value;
        }
        if ((parent.pantalla.document.codigo_barras.formato_impresion.value.indexOf("listado_codigo_barrasSolred3x7.asp") != -1)) {
            maxpagina = parent.pantalla.document.codigo_barras.maxpagina_formSolred3x7.value;
            cantHMax = parent.pantalla.document.codigo_barras.cantHMax_formSolred3x7.value;
            cantVMax = parent.pantalla.document.codigo_barras.cantVMax_formSolred3x7.value;
        }

        ////Se pone esto, ya que si se introduce uno o varios espacios en blanco el isNaN no lo detecta
        while (parent.pantalla.document.codigo_barras.cantidad.value.search(" ") != -1) {
            parent.pantalla.document.codigo_barras.cantidad.value = parent.pantalla.document.codigo_barras.cantidad.value.replace(" ", "");
        }
        if (ok == 1 && parent.pantalla.document.codigo_barras.cantidad.value == '' && parent.pantalla.document.codigo_barras.cant_doc.checked == false) {
            window.alert("<%=LitMsgCantidadNoNulo%>");
            ok = 0;
        }
        if (ok == 1 && isNaN(parent.pantalla.document.codigo_barras.cantidad.value)) {
            window.alert("<%=LitMsgCantidadNoCaracter%>");
            ok = 0;
        }
        ////Se pone esto, ya que si se introduce uno o varios espacios en blanco el isNaN no lo detecta
        while (parent.pantalla.document.codigo_barras.imprimir_listado_horizontal.value.search(" ") != -1) {
            parent.pantalla.document.codigo_barras.imprimir_listado_horizontal.value = parent.pantalla.document.codigo_barras.imprimir_listado_horizontal.value.replace(" ", "");
        }
        if (ok == 1 && isNaN(parent.pantalla.document.codigo_barras.imprimir_listado_horizontal.value)) {
            window.alert("<%=LitMsgHorizontalNoCaracter%>");
            ok = 0;
        }
        ////Se pone esto, ya que si se introduce uno o varios espacios en blanco el isNaN no lo detecta
        while (parent.pantalla.document.codigo_barras.imprimir_listado_vertical.value.search(" ") != -1) {
            parent.pantalla.document.codigo_barras.imprimir_listado_vertical.value = parent.pantalla.document.codigo_barras.imprimir_listado_vertical.value.replace(" ", "");
        }
        if (ok == 1 && isNaN(parent.pantalla.document.codigo_barras.imprimir_listado_vertical.value)) {
            window.alert("<%=LitMsgVerticalNoCaracter%>");
            ok = 0;
        }
        if (ok == 1 && parent.pantalla.document.codigo_barras.imprimir_listado_horizontal.value == '') {
            window.alert("<%=LitMsgHorizontalNoNulo%>");
            ok = 0;
        }
        if (ok == 1 && parent.pantalla.document.codigo_barras.imprimir_listado_vertical.value == '') {
            window.alert("<%=LitMsgVerticalNoNulo%>");
            ok = 0;
        }
        if (ok == 1 && ((parseInt(parent.pantalla.document.codigo_barras.imprimir_listado_vertical.value) - 1) * cantHMax) + (parseInt(parent.pantalla.document.codigo_barras.imprimir_listado_horizontal.value) - 1) > (maxpagina - 1)) {
            window.alert("<%=LitMsgHORVERMAXPAGNoNulo%>");
            ok = 0;
        }
        if (ok == 1 && parseInt(parent.pantalla.document.codigo_barras.imprimir_listado_vertical.value) < 1 || parseInt(parent.pantalla.document.codigo_barras.imprimir_listado_horizontal.value) < 1) {
            window.alert("<%=LitMsgHORVERMINPAGNoNulo%>");
            ok = 0;
        }
        if (ok == 1 && parent.pantalla.document.codigo_barras.numdoc.value != '' && parent.pantalla.document.codigo_barras.tipodoc.value == '') {
            window.alert("<%=LitnumdoctipodocNulo%>");
            ok = 0;
        }
        if (ok == 1 && parent.pantalla.document.codigo_barras.numdoc.value == '' && parent.pantalla.document.codigo_barras.tipodoc.value != '' && parent.pantalla.document.codigo_barras.tipodoc.value != 'ASIGNACION MASIVA') {
            window.alert("<%=LitnumdocnumdocNulo%>");
            ok = 0;
        }

        if (ok == 1 && parent.pantalla.document.codigo_barras.fmpc.value != "") {
            if (!checkdate(parent.pantalla.document.codigo_barras.fmpc)) {
                window.alert("<%=LitAMPFPFechMal%>");
                ok = 0;
            }
        }

        if (ok == 1 && parent.pantalla.document.codigo_barras.stockmayoroigual.value != '') {
            if (isNaN(parent.pantalla.document.codigo_barras.stockmayoroigual.value.replace(",", "."))) {
                window.alert("<%=LitMsgStockNumerico%>");
                ok = 0;
            }
        }

        if (ok == 1)
            return true;
        else
            return false;
    }


    //Realizar la acción correspondiente al botón pulsado.
    function Accion(mode, pulsado, filtro) {
        switch (mode) {
            case "param":
                switch (pulsado) {
                    case "aceptar": //Aceptar


                        if (ValidarCampos()) {
                            codigo_temp = parent.pantalla.document.codigo_barras.cod_temporada.value;
                            filtro_temp = parent.pantalla.document.codigo_barras.filtro_temporada.value;

                            if ((filtro_temp == "incluir") && (codigo_temp != 0))
                                alert("Se ha seleccionado imprimir las etiquetas de temporada. \n Compruebe el papel");
                            else if ((filtro_temp == "excluir") && (codigo_temp != 0))
                                alert("Se ha seleccionado no imprimir las etiquetas de temporada. \n Compruebe el papel");

                            parent.pantalla.document.codigo_barras.action = "codigo_barras_nuevoResultado.asp?mode=ver&nodoc=" + nodoc;
                            parent.pantalla.document.codigo_barras.submit();
                            document.location = "codigo_barras_nuevo_bt.asp?mode=ver";
                        }

                        break;
                    case "cancelar": //Cancelar
                        parent.pantalla.document.codigo_barras.action = "codigo_barras_nuevo.asp?mode=param";
                        parent.pantalla.document.codigo_barras.submit();
                        document.location = "codigo_barras_nuevo_bt.asp?mode=param";
                        break;
                }
                break;
            case "exportar":
                switch (pulsado) {
                    case "aceptar": //Aceptar
                        parent.pantalla.document.location = "codigo_barras_nuevo.asp?mode=param";
                        document.location = "codigo_barras_nuevo_bt.asp?mode=param";
                        break;
                }
                break;
            case "ver":
                switch (pulsado) {
                    case "volver": //Volver atrás
                        parent.pantalla.document.location = "codigo_barras_nuevo.asp?mode=param";
                        document.location = "codigo_barras_nuevo_bt.asp?mode=param";
                        break;
                    case "imprimir": //Imprimir
                        parent.pantalla.mainFrame.focus();
                        valor_form = parent.pantalla.mainFrame.document.forms[0].name;
                        if (valor_form == "listado21_codigo_barras" || valor_form == "listado23_codigo_barras" || valor_form == "listado25_codigo_barras") {
                            printWindowCodBarras(0, 0, 0, 0);
                        }
                        else {
                            if (valor_form == "listado22_codigo_barras") {
                                printWindowCodBarras(75, 75, 75, 0);
                            }
                            else {
                                if (navigator.appName == "Microsoft Internet Explorer") {
                                    printWindow('0', 0);
                                }
                                else {
                                    parent.pantalla.mainFrame.print();
                                }
                            }
                        }
                        break;
                    case "imprimirp": //Imprimir en PDF

                        nlotes = parent.pantalla.mainFrame.document.forms[0].h_lotes.value;
                        valor_form = parent.pantalla.mainFrame.document.forms[0].name;

                        //alert(valor_form);

                        maxpdf = parseInt(parent.pantalla.mainFrame.document.forms[0].maxpdf.value);
                        ok = 1;
                        if (valor_form == "listado4_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        else {
                            if (valor_form == "listado4_codigo_barras") {
                                valor_form = "..\\productos\\listados\\" + valor_form;
                            }
                        }
                        if (valor_form == "listado3n_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }

                        // JMA 30/10/05: Pasar listado_codigo_barras a CUSTOM
                        if (valor_form == "listado2n_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        //if (valor_form=="listado2_codigo_barras" && nlotes>maxpdf){
                        //	ok=0;
                        //}
                        //else{
                        //	if (valor_form=="listado2_codigo_barras"){
                        //		valor_form="..\\productos\\listados\\" + valor_form;
                        //	}
                        //}
                        if (valor_form == "listado5_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado6_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado7_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado8_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado9_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado9SP_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado9_codigo_barrasTHT" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado10_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listado11_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
                        if (valor_form == "listadoSolred3x7_codigo_barras" && nlotes > maxpdf) {
                            ok = 0;
                        }
						if (ok == 1) {
                            //alert(valor_form + "_pdf.asp");
                            cadena = valor_form + "_pdf.asp";

							//ricardo 16-5-2007 para que funcione correctamente el listado 11
							if (valor_form == "listado11_codigo_barras" || valor_form == "listado12_codigo_barras" || valor_form == "listado14_codigo_barras" || valor_form == "listado13_codigo_barras" || valor_form == "..\\productos\\listados\\listado4_codigo_barras" || valor_form == "listado2_codigo_barrasCHACAL" || valor_form == "listadoSolred3x7_codigo_barras" || valor_form == "listado5_codigo_barras") {
                                cadena = cadena + "?cantidad=" + parent.pantalla.mainFrame.document.forms[0].cantidad.value;
                                cadena = cadena + "&nombre=" + parent.pantalla.mainFrame.document.forms[0].nombre.value;
                                cadena = cadena + "&referencia=" + parent.pantalla.mainFrame.document.forms[0].referencia.value;
                                cadena = cadena + "&familia=" + parent.pantalla.mainFrame.document.forms[0].familia.value;
                                cadena = cadena + "&ordenar=" + parent.pantalla.mainFrame.document.forms[0].ordenar.value;
                                cadena = cadena + "&maxpdf=" + parent.pantalla.mainFrame.document.forms[0].maxpdf.value;
                                cadena = cadena + "&maxpagina=" + parent.pantalla.mainFrame.document.forms[0].maxpagina.value;
                                cadena = cadena + "&tipodoc=" + parent.pantalla.mainFrame.document.forms[0].tipodoc.value;
                                cadena = cadena + "&numdoc=" + parent.pantalla.mainFrame.document.forms[0].numdoc.value;
                                cadena = cadena + "&ver_referencia=" + parent.pantalla.mainFrame.document.forms[0].ver_referencia.value;
                                //						    cadena=cadena + "&fye=" + parent.pantalla.document.frames("mainFrame").document.forms[0].fye.value;
                                cadena = cadena + "&ver_nombre=" + parent.pantalla.mainFrame.document.forms[0].ver_nombre.value;
                                cadena = cadena + "&ver_empresa=" + parent.pantalla.mainFrame.document.forms[0].ver_empresa.value;
                                cadena = cadena + "&ver_lineas=" + parent.pantalla.mainFrame.document.forms[0].ver_lineas.value;
                                cadena = cadena + "&ver_precios=" + parent.pantalla.mainFrame.document.forms[0].ver_precios.value;
                                //if (parent.pantalla.document.codigo_barras.si_tiene_modulo_terminales.value==1){
                                //    cadena=cadena + "&ver_codTerminal=" + parent.pantalla.document.frames("mainFrame").document.forms[0].ver_codTerminal.value;
                                //}
                                cadena = cadena + "&imprimir_listado_horizontal=" + parent.pantalla.mainFrame.document.forms[0].imprimir_listado_horizontal.value;
                                cadena = cadena + "&imprimir_listado_vertical=" + parent.pantalla.mainFrame.document.forms[0].imprimir_listado_vertical.value;
                                cadena = cadena + "&fmpc=" + parent.pantalla.mainFrame.document.forms[0].fmpc.value;
                                //						    cadena=cadena + "&solopreciocambiado=" + parent.pantalla.document.frames("mainFrame").document.forms[0].solopreciocambiado.value;
                                cadena = cadena + "&cant_doc=" + parent.pantalla.mainFrame.document.forms[0].cant_doc.value;
                                // JMA 30/10/05: Pasar listado_codigo_barras_nuevo a CUSTOM
                                //if (valor_form=="..\\productos\\listados\\listado2_codigo_barras") {
                                //if (valor_form=="listado2n_codigo_barras") {
                                //	cadena=cadena + "&opcprec1=" + parent.pantalla.document.frames("mainFrame").document.forms[0].opcprec1.value;
                                //	cadena=cadena + "&opcprec2=" + parent.pantalla.document.frames("mainFrame").document.forms[0].opcprec2.value;
                                //	cadena=cadena + "&tarifa1=" + parent.pantalla.document.frames("mainFrame").document.forms[0].tarifa1.value;
                                //	cadena=cadena + "&tarifa2=" + parent.pantalla.document.frames("mainFrame").document.forms[0].tarifa2.value;
                                //	cadena=cadena + "&tarifaiva1=" + parent.pantalla.document.frames("mainFrame").document.forms[0].tarifaiva1.value;
                                //	cadena=cadena + "&tarifaiva2=" + parent.pantalla.document.frames("mainFrame").document.forms[0].tarifaiva2.value;
                                //}
                            }

                            if (valor_form == "listado9_codigo_barras" || valor_form == "listado9SP_codigo_barras" || valor_form == "listado9_codigo_barrasTHT") {
                                //pagina="../crearpdf.asp?mode=LISTADO_CLIENTES&empresa=<%=session("ncliente")%>&impusuario=<%=session("usuario")%>&url=custom/listado9_codigo_barras_pdf.asp";
                                //pagina="listado9_codigo_barras_pdf.asp";
                                if (valor_form == "listado9SP_codigo_barras") {
                                    cadena = cadena + "?cantidad=" + parent.pantalla.mainFrame.document.forms[0].cantidad.value;
                                    cadena = cadena + "&nombre=" + parent.pantalla.mainFrame.document.forms[0].nombre.value;
                                    cadena = cadena + "&referencia=" + parent.pantalla.mainFrame.document.forms[0].referencia.value;
                                    cadena = cadena + "&familia=" + parent.pantalla.mainFrame.document.forms[0].familia.value;
                                    cadena = cadena + "&ordenar=" + parent.pantalla.mainFrame.document.forms[0].ordenar.value;
                                    cadena = cadena + "&maxpdf=" + parent.pantalla.mainFrame.document.forms[0].maxpdf.value;
                                    cadena = cadena + "&maxpagina=" + parent.pantalla.mainFrame.document.forms[0].maxpagina.value;
                                    cadena = cadena + "&tipodoc=" + parent.pantalla.mainFrame.document.forms[0].tipodoc.value;
                                    cadena = cadena + "&numdoc=" + parent.pantalla.mainFrame.document.forms[0].numdoc.value;
                                    cadena = cadena + "&ver_referencia=" + parent.pantalla.mainFrame.document.forms[0].ver_referencia.value;
                                    cadena = cadena + "&ver_nombre=" + parent.pantalla.mainFrame.document.forms[0].ver_nombre.value;
                                    cadena = cadena + "&ver_empresa=" + parent.pantalla.mainFrame.document.forms[0].ver_empresa.value;
                                    cadena = cadena + "&ver_lineas=" + parent.pantalla.mainFrame.document.forms[0].ver_lineas.value;
                                    cadena = cadena + "&ver_precios=" + parent.pantalla.mainFrame.document.forms[0].ver_precios.value;
                                    cadena = cadena + "&imprimir_listado_horizontal=" + parent.pantalla.mainFrame.document.forms[0].imprimir_listado_horizontal.value;
                                    cadena = cadena + "&imprimir_listado_vertical=" + parent.pantalla.mainFrame.document.forms[0].imprimir_listado_vertical.value;
                                    cadena = cadena + "&fmpc=" + parent.pantalla.mainFrame.document.forms[0].fmpc.value;
                                    cadena = cadena + "&cant_doc=" + parent.pantalla.mainFrame.document.forms[0].cant_doc.value;
                                }

                                pagina = cadena;
                                parent.pantalla.document.location = pagina;
                                document.location = "codigo_barras_nuevo_bt.asp?mode=pdf";
                            }
                            else {
                                if (valor_form == "listado23_codigo_barras") {
                                    cadena = cadena + "?tipodoc=" + parent.pantalla.mainFrame.document.forms[0].h_tipodoc.value;
                                    cadena = cadena + "&numdoc=" + parent.pantalla.mainFrame.document.forms[0].h_numdoc.value;
                                }
                                if (valor_form == "listado26_codigo_barras") {
                                    cadena = cadena + "?ordenar=" + parent.pantalla.mainFrame.document.forms[0].ordenar.value;
                                }

                                parent.pantalla.document.location = cadena;
                                document.location = "codigo_barras_nuevo_bt.asp?mode=pdf";
                            }

                        }
                        else window.alert("<%=LitDemRegPDFListCodBarras%>");
                        break;
                }
                break;
            case "pdf":
                switch (pulsado) {
                    case "back": //Volver a la pantalla anterior
                        //parent.pantalla.document.location="codigo_barras_nuevo.asp?mode=param";
                        //document.location="codigo_barras_nuevo_bt.asp?mode=param";
                        parent.document.location = "../central.asp?pag1=custom/codigo_barras_nuevo.asp&mode=param&pag2=custom/codigo_barras_nuevo_bt.asp";

                        break;
                }
                break;
        }
    }
</script>
<body class="body_master_ASP">

<%EscribeControlImpresion "codigo_barras_nuevo.asp"
mode=enc.EncodeForJavascript(Request.QueryString("mode"))%>
<form name="opciones" method="post">
	<div id="PageFooter_ASP" >
        <div id="ControlPanelFooter_ASP" >
	        <table id="BUTTONS_CENTER_ASP">
		<tr>
		    <%if mode="param" then%>
				<td class="CELDABOT" onclick="javascript:Accion('param','aceptar');">
					<%PintarBotonBT LITBOTACEPTAR,ImgAceptar,ParamImgAceptar,LITBOTACEPTARTITLE%>
				</td>
				<td class="CELDABOT" onclick="javascript:Accion('param','cancelar');">
					<%PintarBotonBTRed LITBOTCANCELAR,ImgCancelar,ParamImgCancelar,LITBOTCANCELARTITLE%>
				</td>
			<%elseif mode="ver" then%>
				<td class="CELDABOT" onclick="javascript:Accion('ver','imprimir');">
					<%PintarBotonBT LITBOTIMPRIMIRPAG,ImgImprimir,ParamImgImprimir,LITBOTIMPRIMIRPAGTITLE%>
				</td>
				<td class="CELDABOT" onclick="javascript:Accion('ver','imprimirp');">
					<%PintarBotonBT LITBOTIMPRIMIRLISTADO,ImgImprimir_list,ParamImgImprimir_list,LITBOTIMPRIMIRLISTADOTITLE%>
				</td>
			    <td class="CELDABOT" onclick="javascript:Accion('ver','volver');">
				    <%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
			    </td>
			<%elseif mode="exportar" then%>
			    <!--<td class="CELDABOT" onclick="javascript:Accion('exportar','aceptar');">
				    <%PintarBotonBT LITBOTEXPORTAR,ImgExportar,ParamImgExportar,LITBOTEXPORTARTITLE%>
			    </td>-->
				<td class="CELDABOT" onclick="javascript:Accion('exportar','aceptar');">
					<%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				</td>
			<%elseif mode="pdf" then%>
				<td class="CELDABOT" onclick="javascript:Accion('pdf','back');">
					<%PintarBotonBTRed LITBOTVOLVER,ImgVolver,ParamImgVolver,LITBOTVOLVERTITLE%>
				</td>
			<%end if%>
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