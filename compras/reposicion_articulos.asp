<%@ Language=VBScript %>
<!DOCTYPE html PUBLIC "-//W3C/DTD/ XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml1-transitional.dtd" />
<html LANG="<%=session("lenguaje")%>">
<HEAD>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  

<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="reposicion_articulos.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../CatFamSubResponsive.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/tabs.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc"-->
<!--#include file="../styles/Section.css.inc"-->
<!--#include file="../styles/font-face.css.inc"-->
<!--#include file="../styles/generalData.css.inc"-->
<!--#include file="../styles/ExtraLink.css.inc"-->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->
<!--#include file="../js/dropdown.js.inc" -->

<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>"/>
</head>
<script type="text/javascript" language="javascript" src="../jfunciones.js"></script>
<script type="text/javascript" language="javascript">
    function tratar_importe() {
        if (isNaN(document.reposicion_articulos.stock_mayor.value.replace(",", "."))) {
            window.alert("<%=LitMsgStockMayqueNumerico%>");
            document.reposicion_articulos.stock_mayor.focus();
            document.reposicion_articulos.stock_mayor.select();
        }
    }

    //Desencadena la búsqueda del proveedor cuyo numero se indica
    function TraerProveedor(mode) {
        document.location.href = "reposicion_articulos.asp?nuevoproveedor=" + document.reposicion_articulos.nproveedor.value + "&mode=" + mode +
	"&nserie=" + document.reposicion_articulos.nserie.value +
	"&fpedido=" + document.reposicion_articulos.fpedido.value +
	"&solicitarproveedor=" + document.reposicion_articulos.solicitarProveedor.checked +
	"&referencia=" + document.reposicion_articulos.referencia.value +
	"&nombre=" + document.reposicion_articulos.nombre.value +
	"&proveedor=" + document.reposicion_articulos.proveedor.value +
	"&categoria=" + document.reposicion_articulos.categoria.value +
	"&familia_padre=" + document.reposicion_articulos.familia_padre.value +
	"&familia=" + document.reposicion_articulos.familia.value +
	"&almacen=" + document.reposicion_articulos.almacen.value +
	"&almacenDefecto=" + document.reposicion_articulos.almacenDefecto.checked +
	"&stock_min=" + document.reposicion_articulos.stock_min.checked +
	"&stock_rep=" + document.reposicion_articulos.stock_rep.checked +
	"&mspmb=" + document.reposicion_articulos.mspmb.checked +
	"&reposicion=" + document.reposicion_articulos.reposicion.checked +
	"&reposiciondias=" + document.reposicion_articulos.reposiciondias.value +
	"&reposiciondesde=" + document.reposicion_articulos.reposiciondesde.value +
	"&reposicionhasta=" + document.reposicion_articulos.reposicionhasta.value +
	"&stock_mayor=" + document.reposicion_articulos.stock_mayor.value +
	"&estimacion=" + document.reposicion_articulos.estimacion.checked +
	"&estimaciondias=" + document.reposicion_articulos.estimaciondias.value +
	"&estimaciondesde=" + document.reposicion_articulos.estimaciondesde.value +
	"&estimacionhasta=" + document.reposicion_articulos.estimacionhasta.value +
	"&viene=reposicion_articulos.asp" +
	"&listar=NO";
    }

    function PonerHtml(sentido, lote) {
        cadena = "";
        cadena = "<table width='100%' border='0' cellspacing='1' cellpadding='1'>";
        cadena = cadena + "<tr><td class='MAS'>";
        if (sentido == "next") lote = parseInt(marcoStock.document.reposicion_articulos.lote.value) + 1;
        if (sentido == "prev") lote = parseInt(marcoStock.document.reposicion_articulos.lote.value) - 1;
        if (sentido == "nulo") lote = parseInt(marcoStock.document.reposicion_articulos.lote.value);

        lotes = parseInt(marcoStock.document.reposicion_articulos.lotes.value);
        varias = false
        if (lote > 1) {
            cadena = cadena + "<a class='CELDAREF' href=\"javascript:Mas('prev'," + lote + ");\">";
            cadena = cadena + "<img src='../images/<%=ImgAnterior%>' <%=ParamImgAnterior%> alt='<%=LitAnterior%>' title='<%=LitAnterior%>'/></a>";
            varias = true;
        }
        texto = "<%=LitPagina%>" + " " + lote + " " + "<%=litDe%>" + " " + lotes;
        cadena = cadena + "<font class='CELDA'>" + texto + "</font>";

        if (lote < lotes) {
            cadena = cadena + "<a class='CELDAREF' href=\"javascript:Mas('next'," + lote + ");\">";
            cadena = cadena + "<img src='../images/<%=ImgSiguiente%>' <%=ParamImgSiguiente%> alt='<%=LitgSiguiente%>' title='<%=LitgSiguiente%>'/></a>";
            varias = true;
        }

        cadena = cadena + "</td></tr>";
        cadena = cadena + "</table>";
        document.getElementById("barras").innerHTML = cadena;
    }

    function Mas(sentido, lote) {
        document.getElementById("barras").style.display = "none";
        marcoStock.document.reposicion_articulos.action = "reposicion_articulos_datos.asp?mode=ver&sentido=" + sentido + "&lote=" + lote;
        marcoStock.document.reposicion_articulos.submit();
    }

    function Mostrar() {
        if (isNaN(document.reposicion_articulos.stock_mayor.value.replace(",", "."))) {
            window.alert("<%=LitMsgStockMayqueNumerico%>");
            document.reposicion_articulos.stock_mayor.focus();
            document.reposicion_articulos.stock_mayor.select();
        }
        else if (document.reposicion_articulos.estimacion.checked == true) {
            if ((isNaN(document.reposicion_articulos.estimaciondias.value.replace(",", "."))) || (document.reposicion_articulos.estimaciondias.value == "")) {
                window.alert("<%=LitMsgEstimacionDiasNumerico%>")
                document.reposicion_articulos.estimaciondias.focus();
                document.reposicion_articulos.estimaciondias.select();
            }
            else if ((!checkdate(document.reposicion_articulos.estimaciondesde)) || (document.reposicion_articulos.estimaciondesde.value == "")) {
                window.alert("<%=LitMsgEstimacionDesdeDate%>")
                document.reposicion_articulos.estimaciondesde.focus();
                document.reposicion_articulos.estimaciondesde.select();
            }
            else if ((!checkdate(document.reposicion_articulos.estimacionhasta)) || (document.reposicion_articulos.estimacionhasta.value == "")) {
                window.alert("<%=LitMsgEstimacionHastaDate%>")
                document.reposicion_articulos.estimacionhasta.focus();
                document.reposicion_articulos.estimacionhasta.select();
            }
            else if (DiferenciaTiempo(document.reposicion_articulos.estimacionhasta.value, document.reposicion_articulos.estimaciondesde.value, "dias") < 0) {
                window.alert("<%=LitMsgEstimacionDesdeMayorHastaDate%>")
                document.reposicion_articulos.estimaciondesde.focus();
                document.reposicion_articulos.estimaciondesde.select();
            }
            else
            {
                try
                {
                    //chrome
                    document.getElementById("frstock").document.reposicion_articulos.document.getElementById("waitBoxOculto").style.visibility = "visible";
                }
                catch(e)
                {
                    //IE
                    try
                    {
                        marcoStock.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    }
                    catch (e2)
                    {
                        document.getElementById("frstock").document.getElementById("waitBoxOculto").style.visibility = "visible";
                    }
                }
                setTabsSelected(1);
                try
                {
                    document.reposicion_articulos.target = marcoStock.name;
                }
                catch (e)
                {
                    document.reposicion_articulos.target = document.getElementById("frstock").name;
                }
                document.reposicion_articulos.action = "reposicion_articulos_datos.asp?listar=SI&mode=ver";
                document.reposicion_articulos.submit();
            }
        }
        else if (document.reposicion_articulos.reposicion.checked == true) {
            var longFDesde = document.reposicion_articulos.reposiciondesde.value.search(" ");
            var longFHasta = document.reposicion_articulos.reposicionhasta.value.search(" ");
            var paso = false;
            var paso1 = false;
            if (longFDesde == 0) paso = checkdate(document.reposicion_articulos.reposiciondesde);
            else paso = chkdatetime(document.reposicion_articulos.reposiciondesde.value);
            if (longFHasta == 0) paso1 = checkdate(document.reposicion_articulos.reposicionhasta);
            else paso1 = chkdatetime(document.reposicion_articulos.reposicionhasta.value);
            if ((isNaN(document.reposicion_articulos.reposiciondias.value.replace(",", "."))) || (document.reposicion_articulos.reposiciondias.value == "")) {
                alert("<%=LitMsgReponerDiasNumerico%>")
                document.reposicion_articulos.estimaciondias.focus();
                document.reposicion_articulos.estimaciondias.select();
            }
            else if (!paso || document.reposicion_articulos.reposiciondesde.value == "") {
                alert("<%=LitMsgReponerDesdeDate%>");
                document.reposicion_articulos.reposiciondesde.focus();
                document.reposicion_articulos.reposiciondesde.select();
            }
            else if (!paso1 || document.reposicion_articulos.reposicionhasta.value == "") {
                alert("<%=LitMsgReponerHastaDate%>");
                document.reposicion_articulos.reposicionhasta.focus();
                document.reposicion_articulos.reposicionhasta.select();
            }
            else if (DiferenciaTiempo(document.reposicion_articulos.reposicionhasta.value, document.reposicion_articulos.reposiciondesde.value, "dias") < 0) {
                alert("<%=LitMsgReponerDesdeMayorHastaDate%>");
                document.reposicion_articulos.reposiciondesde.focus();
                document.reposicion_articulos.reposiciondesde.select();
            }
            else {
                try
                {
                    //chrome
                    document.getElementById("frstock").document.reposicion_articulos.document.getElementById("waitBoxOculto").style.visibility = "visible";
                }
                catch (e)
                {
                    //IE
                    try
                    {
                        marcoStock.document.getElementById("waitBoxOculto").style.visibility = "visible";
                    }
                    catch (e2)
                    {
                        document.getElementById("frstock").document.getElementById("waitBoxOculto").style.visibility = "visible";
                    }
                }
                setTabsSelected(1);
                try {
                    document.reposicion_articulos.target = marcoStock.name;
                }
                catch (e) {
                    document.reposicion_articulos.target = document.getElementById("frstock").name;
                }
                document.reposicion_articulos.action = "reposicion_articulos_datos.asp?listar=SI&mode=ver";
                document.reposicion_articulos.submit();
            }
        }
        else {
            try
            {
                //chrome
                document.getElementById("frstock").document.reposicion_articulos.document.getElementById("waitBoxOculto").style.visibility = "visible";
            }
            catch (e)
            {
                //IE
                try
                {
                    marcoStock.document.getElementById("waitBoxOculto").style.visibility = "visible";
                }
                catch (e2)
                {
                    document.getElementById("frstock").document.getElementById("waitBoxOculto").style.visibility = "visible";
                }
            }
            setTabsSelected(1);
            try {
                document.reposicion_articulos.target = marcoStock.name;
            }
            catch (e) {
                document.reposicion_articulos.target = document.getElementById("frstock").name;
            }
            document.reposicion_articulos.action = "reposicion_articulos_datos.asp?listar=SI&mode=ver";
            document.reposicion_articulos.submit();
        }
    }

    function seleccionar(marco, formulario, check) {
        nregistros = eval(marco + ".document." + formulario + ".hNRegs.value-1");
        if (eval("document.reposicion_articulos." + check + ".checked")) {
            var preguntar = 1
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                nombre2 = "cantidad" + i;
                nombre4 = "cpedmin" + i;
                nombre3 = "razonSocial" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");

                str = eval(marco + ".document." + formulario + ".elements[nombre2].value");
                str = (str.split(",").join("."));

                var nuevaCant = parseFloat(str);

                if ((eval(marco + ".document." + formulario + ".elements[nombre2].value") == "") || (nuevaCant <= 0))
                    eval(marco + ".document." + formulario + ".elements[nombre2].value=" + marco + ".document." + formulario + ".elements[nombre4].value");

                if ((eval(marco + ".document." + formulario + ".elements[nombre3].value") == "") && (document.reposicion_articulos.nproveedor.value == "")) {
                    if (preguntar == 1) var mostrarMsg = 1;

                    eval(marco + ".document." + formulario + ".elements[nombre2].value=''");
                    eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
                }

                if (mostrarMsg == 1) {
                    preguntar = 0;
                    mostrarMsg = 0;
                }
            }
        }
        else {
            for (i = 1; i <= nregistros; i++) {
                nombre = "check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
            }
        }
    }

    function seleccionEstimacion(tipo) {
        if (document.reposicion_articulos.estimacion.checked == true || document.reposicion_articulos.reposicion.checked == true) {
            if (document.reposicion_articulos.almacen.value == "") document.reposicion_articulos.almacenDefecto.checked = true;
            if (document.reposicion_articulos.proveedor.value == "") document.reposicion_articulos.mspmb.checked = true;
        }
        if (tipo == "estimacion") document.reposicion_articulos.reposicion.checked = false;

        else if (tipo == "reposicion") {
            document.reposicion_articulos.estimacion.checked = false;
        }
    }

    function AbrirVen() {
        document.getElementById("barras").style.display = "none";
        try {
            marcoStock.document.reposicion_articulos.action = "reposicion_articulos_datos.asp?mode=ver&Prev=SI";
            marcoStock.document.reposicion_articulos.submit();
        }
        catch (e) {
            document.getElementById("frstock").document.reposicion_articulos.action = "reposicion_articulos_datos.asp?mode=ver&Prev=SI";
            document.getElementById("frstock").document.reposicion_articulos.submit();
        }

    }

    //----------------------------------------
    //Funcion para generar pedidos
    //----------------------------------------
    function GenerarPedido() {
        // Validar el campo Serie Pedido
        if (document.reposicion_articulos.nserie.value == "") {
            window.alert("<%=LitMsgSerieNula%>");
            return;
        }

        // Validar el campo Fecha Pedido
        if (document.reposicion_articulos.fpedido.value == "") {
            window.alert("<%=LitMsgFechaNula%>");
            return;
        }
        if (!checkdate(document.reposicion_articulos.fpedido)) {
            window.alert("<%=LitFechaNoVal%>");
            return;
        }

        //Validar que hay algun artículo seleccionado para generar el pedido
        try {
            nregistros = marcoStock.document.reposicion_articulos.hNRegs.value-1;
        }
        catch (e) {
            nregistros = document.getElementById("frstock").document.reposicion_articulos.hNRegs.value - 1;
        }

        var haySel = 0;
        var CantVal = 1;

        for (i = 1; i <= nregistros; i++) {
            nombre = "check" + i;
            nombre1 = "cantidad" + i;

            try {
                if (marcoStock.document.reposicion_articulos.elements[nombre].checked) haySel = 1;
            }
            catch (e) {
                if (document.getElementById("frstock").document.reposicion_articulos.elements[nombre].checked) haySel = 1;
            }

            // Comprobar que las cantidades no son nulas o negativas.
            try {
                var str = eval("marcoStock.document.reposicion_articulos.elements[nombre1].value");
            }
            catch (e) {
                var str = eval("document.getElementById('frstock').document.reposicion_articulos.elements[nombre1].value");
            }
            str = (str.split(",").join("."));

            var nuevaCant = parseFloat(str);

            try {
                if ((marcoStock.document.reposicion_articulos.elements[nombre].checked)) {
                    if ((eval("marcoStock.document.reposicion_articulos.elements[nombre1].value") == "") || (nuevaCant <= 0)) {
                        CantVal = 0;
                        break;
                    }
                }
            }
            catch (e) {
                if ((document.getElementById("frstock").document.reposicion_articulos.elements[nombre].checked)) {
                    if ((eval("document.getElementById('frstock').document.reposicion_articulos.elements[nombre1].value") == "") || (nuevaCant <= 0)) {
                        CantVal = 0;
                        break;
                    }
                }
            }
        }

        if (haySel == 1) {
            try {
                if (CantVal == 1) {
                    if (marcoStock.document.reposicion_articulos.pendiente.value == -1) {
                        if (confirm("<%=LitMsgGenerarPedidoConPendiente%>")) {
                            marcoStock.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            marcoStock.document.reposicion_articulos.action = "reposicion_articulos_datos.asp?mode=generar&fpedido=" + document.reposicion_articulos.fpedido.value + "&nserie=" + document.reposicion_articulos.nserie.value + "&npro=" + document.reposicion_articulos.nproveedor.value + "&solicitarPro=" + document.reposicion_articulos.solicitarProveedor.checked + "&pendiente=1"
                            marcoStock.document.reposicion_articulos.submit();
                        }
                    }
                    else {
                        if (confirm("<%=LitMsgGenerarPedido%>")) {
                            marcoStock.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            marcoStock.document.reposicion_articulos.action = "reposicion_articulos_datos.asp?mode=generar&fpedido=" + document.reposicion_articulos.fpedido.value + "&nserie=" + document.reposicion_articulos.nserie.value + "&npro=" + document.reposicion_articulos.nproveedor.value + "&solicitarPro=" + document.reposicion_articulos.solicitarProveedor.checked + "&pendiente=0";
                            marcoStock.document.reposicion_articulos.submit();
                        }
                    }
                }
                else window.alert("<%=LitMsgCantNoVal%>");
            }
            catch (e) {
                if (CantVal == 1) {
                    if (document.getElementById("frstock").document.reposicion_articulos.pendiente.value == -1) {
                        if (confirm("<%=LitMsgGenerarPedidoConPendiente%>")) {
                            document.getElementById("frstock").document.getElementById("waitBoxOculto").style.visibility = "visible";
                            document.getElementById("frstock").document.reposicion_articulos.action = "reposicion_articulos_datos.asp?mode=generar&fpedido=" + document.reposicion_articulos.fpedido.value + "&nserie=" + document.reposicion_articulos.nserie.value + "&npro=" + document.reposicion_articulos.nproveedor.value + "&solicitarPro=" + document.reposicion_articulos.solicitarProveedor.checked + "&pendiente=1"
                            document.getElementById("frstock").document.reposicion_articulos.submit();
                        }
                    }
                    else {
                        if (confirm("<%=LitMsgGenerarPedido%>")) {
                            document.getElementById("frstock").document.getElementById("waitBoxOculto").style.visibility = "visible";
                            document.getElementById("frstock").document.reposicion_articulos.action = "reposicion_articulos_datos.asp?mode=generar&fpedido=" + document.reposicion_articulos.fpedido.value + "&nserie=" + document.reposicion_articulos.nserie.value + "&npro=" + document.reposicion_articulos.nproveedor.value + "&solicitarPro=" + document.reposicion_articulos.solicitarProveedor.checked + "&pendiente=0";
                            document.getElementById("frstock").document.reposicion_articulos.submit();
                        }
                    }
                }
                else window.alert("<%=LitMsgCantNoVal%>");
            }

        }
        else window.alert("<%=LitMsgNoHaySel%>");
    }

    function calculaDias(tipo) {
        //Set 1 day in milliseconds
        var dia = 1000 * 60 * 60 * 24
        if (tipo == "estimacion") {
            if (checkdate(document.reposicion_articulos.estimacionhasta) && checkdate(document.reposicion_articulos.estimaciondesde)) {
                //Calculate difference btw the two dates, and convert to days
                var arraydesde = document.reposicion_articulos.estimaciondesde.value.split("/")
                var arrayhasta = document.reposicion_articulos.estimacionhasta.value.split("/")
                var desde = new Date(arraydesde[1] + "/" + arraydesde[0] + "/" + arraydesde[2]);
                var hasta = new Date(arrayhasta[1] + "/" + arrayhasta[0] + "/" + arrayhasta[2]);

                dias = Math.ceil((hasta.getTime() - desde.getTime()) / (dia)) + 1;

                document.reposicion_articulos.estimaciondias.value = dias;
                document.reposicion_articulos.reposicion.checked = false;
            }
        }
        else if (tipo == "reposicion") {
            if (checkdate(document.reposicion_articulos.reposicionhasta) && checkdate(document.reposicion_articulos.reposiciondesde)) {
                //Calculate difference btw the two dates, and convert to days
                var arraydesde = document.reposicion_articulos.reposiciondesde.value.split("/")
                var arrayhasta = document.reposicion_articulos.reposicionhasta.value.split("/")
                var desde = new Date(arraydesde[1] + "/" + arraydesde[0] + "/" + arraydesde[2]);
                var hasta = new Date(arrayhasta[1] + "/" + arrayhasta[0] + "/" + arrayhasta[2]);

                dias = Math.ceil((hasta.getTime() - desde.getTime()) / (dia)) + 1;

                document.reposicion_articulos.reposiciondias.value = dias;
                document.reposicion_articulos.estimacion.checked = false;
            }
        }
    }

    function tier1Menu(objMenu, objImage) {
        if (objMenu.style.display == "none") {
            objMenu.style.display = "";
            objImage.src = "../Images/<%=ImgCarpetaAbierta%>";
            switch (objMenu.id) {
                case "detalles":
                    document.getElementById("parametros").style.display = "none";
                    document.getElementById("img1").src = "../Images/<%=ImgCarpetaCerrada%>";
                    break;

                case "parametros":
                    document.getElementById("detalles").style.display = "none";
                    document.getElementById("img2").src = "../Images/<%=ImgCarpetaCerrada%>";
                    document.location = "reposicion_articulos.asp?mode=param";
                    break;
            }
        }
        else {
            objMenu.style.display = "none";
            objImage.src = "../images/<%=ImgCarpetaCerrada%>";
        }
    }

    //  GPD (15/05/2007).
    function isObject(a) {
        return (typeof a == 'object' && !!a) || isFunction(a);
    }

    //  GPD (15/05/2007).
    function SeleccionarListasCompletas() {
        if (isObject(document.reposicion_articulos.seriefactura)) {
            for (i = 0; i <= document.reposicion_articulos.seriefactura.options.length; i++) document.reposicion_articulos.seriefactura.options[i].selected = true;
        }

        if (isObject(document.reposicion_articulos.serieticket)) {
            for (j = 0; j <= document.reposicion_articulos.serieticket.options.length; j++) document.reposicion_articulos.serieticket.options[j].selected = true;
        }

        if (isObject(document.reposicion_articulos.seriealbaran)) {
            for (k = 0; k <= document.reposicion_articulos.seriealbaran.options.length; k++) document.reposicion_articulos.seriealbaran.options[k].selected = true;
        }
    }
</script>

<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('ORDERS', 'fade=1')

    animatedcollapse.ontoggle = function ($, divobj, state) {
        //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }
    animatedcollapse.init();

</script>

<body onload="self.status='';" bgcolor="<%=color_blau%>">
<%
'----------------------------------------------------------------------------
'Funciones
'----------------------------------------------------------------------------

'Botones de navegación para las búsquedas.
sub SpanNextPrev(lote,lotes,pos)%>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
	<tr><td class="MAS"><%
	   lote=cint(lote)
	   lotes=cint(lotes)
	    varias=false
		if lote>1 then
			%><a class="CELDAREF" href="javascript:Mas('prev',<%=enc.EncodeForJavascript(lote)%>);">
			<img src="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a><%
			varias=true
		end if
		texto=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)
		%><font class="CELDA"> <%=texto%> </font> <%

		if lote<lotes then
			%><a class="CELDAREF" href="javascript:Mas('next',<%=enc.EncodeForJavascript(lote)%>);">
			<img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a><%
			varias=true
		end if%>
	</td></tr>
</table>
<%end sub

sub SpanStock()%>
	<table class="width90 lg-table-responsive bCollapse">
        <tr>
			<td class="CELDA underOrange ENCABEZADOL width5"><input type="Checkbox" name="check" onClick="seleccionar('marcoStock','reposicion_articulos','check');"></td>
			<td class="CELDA underOrange ENCABEZADOL width5"><%=LitCantidad%></td>			
			<td class="CELDA underOrange ENCABEZADOL width10"><%=LitCUE%></td>
			<td class="CELDA underOrange ENCABEZADOL width10"><%=LitProveedor%></td>
			<td class="CELDA underOrange ENCABEZADOL width10"><%=LitRef%></td>
			<td class="CELDA underOrange ENCABEZADOL width10"><%=LitNombre%></td>
			<td class="CELDACENTER underOrange ENCABEZADOL width10"><%=LitAlmacen%></td>
			<td class="CELDARIGHT underOrange ENCABEZADOL width5"><%=LitPVD%></td>
			<td class="CELDACENTER underOrange ENCABEZADOL width5"><%=LitStock%></td>
			<td class="CELDACENTER underOrange ENCABEZADOL width5"><%=LitStockMin%></td>
			<td class="CELDACENTER underOrange ENCABEZADOL width5"><%=LitPRecibir%></td>
			<td class="CELDACENTER underOrange ENCABEZADOL width5"><%=LitPServir%></td>
			<td class="CELDACENTER underOrange ENCABEZADOL width5"><%=LitPrep%></td>
        </tr>
    </table>
    
    <!-- GPD (08/05/2007) -->
	<iframe name="marcoStock" id="frstock" src='reposicion_articulos_datos.asp?mode=vacio' class='width90 lg-table-responsive iframe-tab-space iframe-data' height="400px;"></iframe>
<%
    DrawDiv "1","","" %><span id="barras" style="display:none"></span>
    <%CloseDiv
      DrawDiv "col-lg-3 col-md-6 col-sm-6 col-xs-12","",""
      %><a class="CELDAREFB" href="javascript:AbrirVen()"><%=LitPrevisualizar%></a>
    <%CloseDiv
      DrawDiv "col-lg-3 col-md-6 col-sm-6 col-xs-12","",""
      %><a href="javascript:GenerarPedido();"><img src="../images/<%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitGenPedido%>" title="<%=LitGenPedido%>"/></a>
    <%CloseDiv
end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
borde=0

if accesoPagina(session.sessionid,session("usuario"))=1 then  %>
<form name="reposicion_articulos" method="post">
	<%WaitBoxOculto LitEsperePorFavor
	PintarCabecera "reposicion_articulos.asp"

		'Leer parámetros de la página
		listar=limpiaCadena(request.querystring("listar"))
  		mode=Request.QueryString("mode")
		if request.querystring("lote") >"" then
		   lote = limpiaCadena(request.querystring("lote"))
		elseif request.form("lote")>"" then
		   lote = limpiaCadena(request.form("lote"))
		else
		   lote = 1
		end if
		if request.querystring("referencia") >"" then
		   referencia = limpiaCadena(request.querystring("referencia"))
		else
		   referencia = limpiaCadena(request.form("referencia"))
		end if
		if request.querystring("nombre") >"" then
		   nombre = limpiaCadena(request.querystring("nombre"))
		else
		   nombre = limpiaCadena(request.form("nombre"))
		end if
		if request.querystring("almacen") >"" then
		   almacen = limpiaCadena(request.querystring("almacen"))
		else
		   almacen = limpiaCadena(request.form("almacen"))
		end if
		if request.querystring("categoria") >"" then
		   categoria = limpiaCadena(request.querystring("categoria"))
		else
		   categoria = limpiaCadena(request.form("categoria"))
		end if
		if request.querystring("familia_padre") >"" then
		   familia_padre = limpiaCadena(request.querystring("familia_padre"))
		else
		   familia_padre = limpiaCadena(request.form("familia_padre"))
		end if
		if request.querystring("familia") >"" then
		   familia = limpiaCadena(request.querystring("familia"))
		else
		   familia = limpiaCadena(request.form("familia"))
		end if
		if request.querystring("ordenar") >"" then
		   ordenar = limpiaCadena(request.querystring("ordenar"))
		else
		   ordenar = limpiaCadena(request.form("ordenar"))
		end if
		if request.querystring("estimacion")>"" then
		   estimacion=limpiaCadena(request.querystring("estimacion"))
		else
		   estimacion=limpiaCadena(request.form("estimacion"))
		end if
		if request.querystring("estimaciondias")>"" then
		   estimaciondias=limpiaCadena(request.querystring("estimaciondias"))
		else
		   estimaciondias=limpiaCadena(request.form("estimaciondias"))
		end if
		if request.querystring("estimaciondesde")>"" then
		   estimaciondesde=limpiaCadena(request.querystring("estimaciondesde"))
		else
		   estimaciondesde=limpiaCadena(request.form("estimaciondesde"))
		end if
		if request.querystring("estimacionhasta")>"" then
		   estimacionhasta=limpiaCadena(request.querystring("estimacionhasta"))
		else
		   estimacionhasta=limpiaCadena(request.form("estimacionhasta"))
		end if

		if request.querystring("reposicion")>"" then
		   reposicion=limpiaCadena(request.querystring("reposicion"))
		else
		   reposicion=limpiaCadena(request.form("reposicion"))
		end if
		if request.querystring("reposiciondias")>"" then
		   reposiciondias=limpiaCadena(request.querystring("reposiciondias"))
		else
		   reposiciondias=limpiaCadena(request.form("reposiciondias"))
		end if
		if request.querystring("reposiciondesde")>"" then
		   reposiciondesde=limpiaCadena(request.querystring("reposiciondesde"))
		else
		   reposiciondesde=limpiaCadena(request.form("reposiciondesde"))
		end if
		if request.querystring("reposicionhasta")>"" then
		   reposicionhasta=limpiaCadena(request.querystring("reposicionhasta"))
		else
		   reposicionhasta=limpiaCadena(request.form("reposicionhasta"))
		end if
		
		if request.querystring("proveedor") >"" then
		   proveedor = limpiaCadena(request.querystring("proveedor"))
		else
		   proveedor = limpiaCadena(request.form("proveedor"))
		end if

		'Formateamos el campo a 5 digitos
		if proveedor>"" then
			proveedor=session("ncliente") & completar(proveedor,5,"0")
		end if

		if request.querystring("nproveedor") >"" then
		   nuevoProveedor = limpiaCadena(request.querystring("nproveedor"))
		else
		   nuevoProveedor = limpiaCadena(request.form("nproveedor"))
		end if

		'Formateamos el campo a 5 digitos
		if nuevoProveedor>"" and false then
			nuevoProveedor=session("ncliente") & completar(nuevoProveedor,5,"0")
		end if

		if request.querystring("nserie") >"" then
		   nserie = limpiaCadena(request.querystring("nserie"))
		else
		   nserie = limpiaCadena(request.form("nserie"))
		end if

		if request.querystring("solicitarproveedor") >"" then
		   solicitarproveedor = limpiaCadena(request.querystring("solicitarproveedor"))
		else
		   solicitarproveedor = limpiaCadena(request.form("solicitarproveedor"))
		end if

		if request.querystring("almacenDefecto") >"" then
		   almacenDefecto = limpiaCadena(request.querystring("almacenDefecto"))
		else
		   almacenDefecto = limpiaCadena(request.form("almacenDefecto"))
		end if

		if request.querystring("stock_min") >"" then
		   stock_min = limpiaCadena(request.querystring("stock_min"))
		else
		   stock_min = limpiaCadena(request.form("stock_min"))
		end if

		if request.querystring("stock_rep") >"" then
		   stock_rep = limpiaCadena(request.querystring("stock_rep"))
		else
		   stock_rep = limpiaCadena(request.form("stock_rep"))
		end if

		if request.querystring("stock_mayor") >"" then
		   stock_mayor= limpiaCadena(request.querystring("stock_mayor"))
		else
		   stock_mayor= limpiaCadena(request.form("stock_mayor"))
		end if

		if request.querystring("mspmb") >"" then
		   mspmb= limpiaCadena(request.querystring("mspmb"))
		else
		   mspmb= limpiaCadena(request.form("mspmb"))
		end if

		if request.querystring("cantpendiente") >"" then
		   cantpendiente= limpiaCadena(request.querystring("cantpendiente"))
		else
		   cantpendiente= limpiaCadena(request.form("cantpendiente"))
		end if%>
    <input type="hidden" name="p_referencia" value="<%=enc.EncodeForHtmlAttribute(referencia)%>"/>
	<%Alarma "reposicion_articulos.asp" %>

	<%set rstAux = Server.CreateObject("ADODB.Recordset")
    set rst = Server.CreateObject("ADODB.Recordset")

  '*********************************************************************************************
  'Se muestran parametros de seleccion
  '*********************************************************************************************
  if mode="param" then
	'************* PARAMETROS PARA LOS DATOS DEL PEDIDO **************
    %><div class="headers-wrapper"><%
            rstAux.cursorlocation=3
			rstAux.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where tipo_documento ='PEDIDO A PROVEEDOR' and nserie like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
            DrawDiv "col-sm-4 col-xs-6 col-xxs-12","",""
            DrawLabel "","",LitSeriePedido
            DrawSelect "width60","","nserie",rstAux,nserie,"nserie","descripcion","",""
            CloseDiv
		 	rstAux.close
            '------------------------------------------------------------
            DrawDiv "header-date","",""
                DrawLabel "","",LitFechaPedido      
			if fpedido>"" then
                DrawInput "","width:80px","fpedido",fpedido,""
                DrawCalendar "fpedido"
			else
                DrawInput "","width:80px","fpedido",date,""
                DrawCalendar "fpedido"
			end if
            CloseDiv
            if session("version")&"" <> "5" then
                DrawDiv "","","" 
                CloseDiv
            end if
            DrawDiv "header-prov","",""
                DrawLabel "","",LitPedirProveedor%><input class="CELDA" type="hidden" name="nproveedor" value="<%=enc.EncodeForHtmlAttribute(nuevoproveedor)%>"/><input class="CELDA" type="hidden" name="nom_proveedor" value=""/><iframe id='frProveedor' class="width60 iframe-menu" src='../compras/docproveedor_responsive.asp?viene=reposicion_articulos&anterior=fpedido&siguiente=solicitarProveedor&nproveedor=<%=enc.EncodeForHtmlAttribute(nuevoproveedor)%>' frameborder="no" scrolling="no" noresize="noresize"></iframe><%CloseDiv
            if session("version")&"" <> "5" then
                DrawDiv "","","" 
                CloseDiv
            end if
            DrawDiv "1","",""
                DrawLabel "","",LitSolicitarAProveedor
			if solicitarProveedor="true" then%><input class="CELDA" type='checkbox' name='solicitarProveedor' checked/>
			<%else%><input class="CELDA" type='checkbox' name='solicitarProveedor'/>
			<%end if
		    CloseDiv%>
	    </div>
        <table class="width100"></table>
<%end if 

'************* PARAMETROS PARA LA SELECCION DE ARTICULOS **************%>
  	    
    <div class="Section" id="S_ORDERS" >
        <a href="#" rel="toggle[ORDERS]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader">
                <%=LitTitle %>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
            </div>
        </a>
        <div class="SectionPanel" id="ORDERS" >
            <div id="tabs" style="display:none">
                <ul>
                    <% if mode="param" then %>
                        <li id="li-parameters"><a href="#tabs-parameters"><%=LitParamBusqueda %></a></li>
                    <%end if %>
                    <li id="li-results"><a href="#tabs-results"><%=LitResultado %></a></li>
                </ul>
	            <%if mode ="param" then %>
                    <div id="tabs-parameters" class="overflowXauto">
                            <%  EligeCelda "input","add","left","","",0,LitConref,"referencia",35,trimCodEmpresa(referencia)                           
                                EligeCelda "input","add","left","","",0,LitConNombre,"nombre",35,nombre
                                nomproSELECT="select razon_social from proveedores with(nolock) where nproveedor = ?"
                                nompro = DlookupP1(nomproSELECT, proveedor&"", adchar, 10, session("dsn_cliente"))
                                if nompro = "" and proveedor > "" then
	                                proveedor = ""%><script type="text/javascript" language="javascript">alert("<%=LitMsgNoProveedor%>");</script>
                                <%end if%>
                                <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><label><%=LitProveedor%></label><%
                             %><input class="width15" type="text" name="proveedor" value="<%=enc.EncodeForHtmlAttribute(null_s(trimCodEmpresa(proveedor)))%>" size="10" onchange="TraerProveedor('<%=enc.EncodeForJavascript(mode)%>');"><a class="CELDAREFB"  href="javascript:AbrirVentana('../compras/proveedores_busqueda.asp?ndoc=reposicion_articulos1&titulo=<%=LitSelProveedor%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerProveedor%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscarProveedor%>" title="<%=LitBuscarProveedor%>"/></a><input class="width40" type="text" size="40" disabled name="razon_social1" value="<%=enc.EncodeForHtmlAttribute(null_s(nompro))%>"/></div><%
                                rstAux.cursorlocation=3
                                rstAux.open " select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion", session("dsn_cliente")
                                DrawSelectCelda "width60","","",0,LitAlmacen,"almacen",rstAux,almacen,"codigo","descripcion","onclick","javascript:seleccionEstimacion()"
                                rstAux.close	
			                    dim ConfigDespleg (3,13)

			                    i=0
			                    ConfigDespleg(i,0)="categoria"
			                    ConfigDespleg(i,1)="200"
			                    ConfigDespleg(i,2)="8"
			                    ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			                    ConfigDespleg(i,4)=1
			                    ConfigDespleg(i,5)="width60"
			                    ConfigDespleg(i,6)="MULTIPLE"
			                    ConfigDespleg(i,7)="codigo"
			                    ConfigDespleg(i,8)="nombre"
			                    ConfigDespleg(i,9)=LitCategoria
			                    ConfigDespleg(i,10)=categoria
			                    ConfigDespleg(i,11)=""
			                    ConfigDespleg(i,12)=""

			                    i=1
			                    ConfigDespleg(i,0)="familia_padre"
			                    ConfigDespleg(i,1)="200"
			                    ConfigDespleg(i,2)="8"
			                    ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			                    ConfigDespleg(i,4)=1
			                    ConfigDespleg(i,5)="width60"
			                    ConfigDespleg(i,6)="MULTIPLE"
			                    ConfigDespleg(i,7)="codigo"
			                    ConfigDespleg(i,8)="nombre"
			                    ConfigDespleg(i,9)=LitFamilia
			                    ConfigDespleg(i,10)=familia_padre
			                    ConfigDespleg(i,11)=""
			                    ConfigDespleg(i,12)=""

			                    i=2
			                    ConfigDespleg(i,0)="familia"
			                    ConfigDespleg(i,1)="200"
			                    ConfigDespleg(i,2)="8"
			                    ConfigDespleg(i,3)="select codigo, nombre, categoria, padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			                    ConfigDespleg(i,4)=1
			                    ConfigDespleg(i,5)="width60"
			                    ConfigDespleg(i,6)="MULTIPLE"
			                    ConfigDespleg(i,7)="codigo"
			                    ConfigDespleg(i,8)="nombre"
			                    ConfigDespleg(i,9)=LitSubFamilia
			                    ConfigDespleg(i,10)=familia
			                    ConfigDespleg(i,11)=""
			                    ConfigDespleg(i,12)=""
			                    DibujaDesplegables ConfigDespleg,session("dsn_cliente")
                            
                                DrawInputCeldaActionDiv "","","",10,0,LitSTockMayor,"stock_mayor",iif(stock_mayor>"",stock_mayor,""),"onchange","tratar_importe()",false

		                '   GPD (28/02/2007)
                            DrawDiv "1", "", ""
                            DrawLabel "", "", LitIncluirCantPendiente%><input class="checkbox-label" type="checkbox" name="cantpendiente"/><%CloseDiv
		                    DrawDiv "1", "", ""
                            DrawLabel "", "", LitAlmacenDefecto
			                if almacenDefecto="true" then
				                %><input class="checkbox-label" type="checkbox" name="almacenDefecto" checked onclick="javascript: seleccionEstimacion()"/><%
			                else
				                %><input class="checkbox-label" type="checkbox" name="almacenDefecto" onclick="javascript:seleccionEstimacion()"/><%
			                end if
                            CloseDiv
                            DrawDiv "1","",""
                            DrawLabel "","",LitStockMin2
			                if stock_min="true" then
				                %><input class="checkbox-label" type="checkbox" name="stock_min" checked/><%
			                else
				                %><input class="checkbox-label" type="checkbox" name="stock_min"/><%
			                end if
                            CloseDiv
			
			                DrawDiv "1","",""
                            DrawLabel "","",LitPrep2
			                if stock_rep="true" then
				                %><input class="checkbox-label" type="checkbox" name="stock_rep" checked/><%
			                else
				                %><input class="checkbox-label" type="checkbox" name="stock_rep"/><%
			                end if
                            CloseDiv
		                    'sacamos la fecha de hace 3 meses
		                    edesde = date - 14
		                    mes1 = month(edesde)
		                    ano1 = year(edesde)
		                    dia1 = day(edesde)

		                    fecha1 = iif(len(dia1)=1, "0"&dia1,dia1) & "/" & iif(len(mes1)=1, "0"&mes1,mes1) & "/" & ano1
		                    marcado = ""
		    
		                    if estimacion = "on" or estimacion = "true" then
			                    marcado = "checked"
		                    end if
                            DrawDiv "3", "", ""
                            CloseDiv
                                DrawDiv "4", "", ""
                                    %><input type="checkbox" name="estimacion" <%=marcado%> onclick="javascript:seleccionEstimacion('estimacion')"><%
                                CloseDiv
                                DrawDiv "6", "", ""
                                    DrawLabel "", "", LitEstimacion1
                                    %><input type="text" name="estimaciondias" value="<%=iif(estimaciondias>"",enc.EncodeForHtmlAttribute(estimaciondias),15)%>"/><%
                                    DrawSpan "CELDA", "", Mid(LitEstimacion2, 1, 5) , ""
                                CloseDiv
                                DrawDiv "6", "", ""
                                    DrawLabel "", "", Mid(LitEstimacion2, 6)
                                    %><input type="text" name="estimaciondesde" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(estimaciondesde>"",estimaciondesde,fecha1)))%>" onchange="javascript:calculaDias('estimacion')"/><%
                                    DrawCalendar "estimaciondesde"
                                CloseDiv
                                DrawDiv "6", "", ""
                                    DrawLabel "", "", LitEstimacion3
                                    %><input type="text" name="estimacionhasta" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(estimacionhasta>"",estimacionhasta,date)))%>" onchange="javascript:calculaDias('estimacion')"/><%
                                    DrawCalendar "estimacionhasta"
                                CloseDiv                            

                            DrawDiv "1", "", ""
                            DrawLabel "","",LitMostProvMasBar
			                %><input class="CELDA" type="checkbox" name="mspmb" onclick="javascript:seleccionEstimacion()"/><%		
			                if Reposicion="on" or Reposicion="true" then			
			                    marcado="checked"		
		                    end if
                            CloseDiv
                            DrawDiv "3", "", ""
                            CloseDiv
                            DrawDiv "4", "", ""
                                %><input type="checkbox" name="reposicion" <%=marcado%> onclick="javascript:seleccionEstimacion('reposicion')"/><%
                            CloseDiv
                            DrawDiv "6", "", ""
                                DrawLabel "", "", LitReposicion
                                %><input type="hidden" name="reposiciondias" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(reposiciondias>"",reposiciondias,15)))%>"/><%
                                %><input type="text" name="reposiciondesde" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(reposiciondesde>"",reposiciondesde,fecha1)))%>" onchange="javascript:calculaDias('reposicion')"/><%
                                DrawCalendar "reposiciondesde"
                            CloseDiv
                            DrawDiv "6", "", ""
                                DrawLabel "", "", LitEstimacion3
                                %><input type="text" name="reposicionhasta" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(reposicionhasta>"",reposicionhasta,date)))%>" onchange="javascript:calculaDias('reposicion')"/><%
                                DrawCalendar "reposicionhasta"
                            CloseDiv
                            DrawDiv "3", "", ""
                                 DrawLabel "", "", LitMensajeHoras + LitMensajeHoras1
                            CloseDiv 
        		            DrawDiv "3", "", ""
                            CloseDiv
		
		                if estimacion="on" or estimacion="true" then%>
			                <script type="text/javascript" language=javascript>seleccionEstimacion('estimacion');</script>
		                <%else%>
			                <script type="text/javascript" language=javascript>
			                    document.reposicion_articulos.reposicion.checked = true;
			                    seleccionEstimacion('reposicion');
			                </script>		
		                <%end if
		
		                Dim arrDesplegables (3,13)

		                i = 0
		                arrDesplegables(i,0) = "seriefactura"
		                arrDesplegables(i,1) = ""
		                arrDesplegables(i,2) = "8"
		                arrDesplegables(i,3) = "select SUBSTRING(NSERIE,6,LEN(NSERIE) - 5) + '-' + NOMBRE as TITULO, NSERIE as CODIGO from SERIES with(nolock) where NSERIE like '" & session("ncliente") & "%' and TIPO_DOCUMENTO = 'FACTURA A CLIENTE' order by TITULO"
		                arrDesplegables(i,4) = 1
		                arrDesplegables(i,5) = "width60"
		                arrDesplegables(i,6) = "MULTIPLE"
		                arrDesplegables(i,7) = "CODIGO"
		                arrDesplegables(i,8) = "TITULO"
		                arrDesplegables(i,9) = LitFacturas
		                arrDesplegables(i,10) = ""
		                arrDesplegables(i,11) = ""
		                arrDesplegables(i,12) = ""

		                i = 1
		                arrDesplegables(i,0) = "serieticket"
		                arrDesplegables(i,1) = ""
		                arrDesplegables(i,2) = "8"
		                arrDesplegables(i,3) = "select SUBSTRING(NSERIE,6,LEN(NSERIE) - 5) + '-' + NOMBRE as TITULO, NSERIE as CODIGO from SERIES with(nolock) where NSERIE like '" & session("ncliente") & "%' and TIPO_DOCUMENTO = 'TICKET' order by TITULO"
		                arrDesplegables(i,4) = 2
		                arrDesplegables(i,5) = "width60"
		                arrDesplegables(i,6) = "MULTIPLE"
		                arrDesplegables(i,7) = "CODIGO"
		                arrDesplegables(i,8) = "TITULO"
		                arrDesplegables(i,9) = LitTicketsPendientes
		                arrDesplegables(i,10) = ""
		                arrDesplegables(i,11) = ""
		                arrDesplegables(i,12) = ""

		                i = 2
		                arrDesplegables(i,0) = "seriealbaran"
                        arrDesplegables(i,1) = ""
		                arrDesplegables(i,2) = "8"
		                arrDesplegables(i,3) = "select SUBSTRING(NSERIE,6,LEN(NSERIE) - 5) + '-' + NOMBRE as TITULO, NSERIE as CODIGO from SERIES with(nolock) where NSERIE like '" & session("ncliente") & "%' and TIPO_DOCUMENTO = 'ALBARAN DE SALIDA' order by TITULO"
		                arrDesplegables(i,4) = 2
		                arrDesplegables(i,5) = "width60"
		                arrDesplegables(i,6) = "MULTIPLE"
		                arrDesplegables(i,7) = "CODIGO"
		                arrDesplegables(i,8) = "TITULO"
		                arrDesplegables(i,9) = LitAlbaranesPendientes
		                arrDesplegables(i,10) = ""
		                arrDesplegables(i,11) = ""
		                arrDesplegables(i,12) = ""
				
			            DibujaDesplegablesSeleccionados arrDesplegables, session("dsn_cliente")
                        if session("version")&"" <> "5" then
                            DrawDiv "","","" 
                            CloseDiv
                        end if
		                DrawDiv "1","",""
                            DrawLabel "","",LitEnviarConsulta				
			                %><a href="javascript:Mostrar();" class="ic-accept"><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgAplicar%> alt="<%=LitCargarInv%>" title="<%=LitCargarInv%>"/></a><%
		                CloseDiv%>
                    </div ><%        
                
                else
                    %><div id="tabs-parameters"></div><%
                end if

                maxpagina=25
                '*********************************************************************************************
                ' Se muestran los datos de la consulta
                '*********************************************************************************************%>
	            <div id="tabs-results" class="overflowXauto">
                    <%SpanStock%>
                </div>
            </div>
        </div>
    </div>
    </form>
<%end if%> 
</body>
</html>