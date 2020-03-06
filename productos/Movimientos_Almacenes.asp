<%@ Language=VBScript %>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>

<%
' ############ SOLO PARA DEVOLVER CONSULTAS AJAX SOBRE ESTA MISMA PÁGINA ############
function limpiaCadenaDetins(strValor)
	dim returnedValue

	returnedValue=strValor

	returnedValue=replace(returnedValue,"'","''")
	returnedValue=replace(returnedValue,"--","")
	returnedValue=replace(returnedValue,";","")
	returnedValue=replace(returnedValue,"select","")
	returnedValue=replace(returnedValue,"drop","")
	returnedValue=replace(returnedValue,"insert","")
	returnedValue=replace(returnedValue,"update","")
	returnedValue=replace(returnedValue,"delete","")
	returnedValue=replace(returnedValue,"xp_","")
	returnedValue=replace(returnedValue,"sp_","")
	returnedValue=replace(returnedValue,"shutdown","")
	returnedValue=replace(returnedValue,"bulk","")
	returnedValue=replace(returnedValue,"bcp","")
	returnedValue=replace(returnedValue,"script","")
	returnedValue=replace(returnedValue,"declare","")
	returnedValue=replace(returnedValue,"exec","")

	'Para evitar tratar los tags de html como tales.
	'returnedValue=server.htmlencode(returnedValue)

	limpiaCadenaDetins=returnedValue
end function
if request.querystring("mode") = "consultaAJAX" then
    if request.querystring("consulta")="introducirCodeAt" then
        p_nmovimiento = limpiaCadena(request.querystring("nmovimiento"))
        p_codeat = limpiaCadena(request.querystring("codeat"))
        if p_nmovimiento & "">"" then
            if p_codeat & "">"" then
                set rstBlq = Server.CreateObject("ADODB.Recordset")
                rstBlq.open "update movimientos with(UPDLOCK) set SAFTATDOCCODEID='" & p_codeat & "' where nmovimiento='" & p_nmovimiento & "'",session("dsn_cliente")
                if rstBlq.state<>0 then rstBlq.close
                set rstBlq =nothing
                response.write("OK")
            else
                response.write("NOOK1")
            end if
        else
            response.write("NOOK2")
        end if
        response.end
    end if
    if request.querystring("consulta")="MarcarMercanciaRecibida" then
        p_nmovimiento = limpiaCadena(request.querystring("nmovimiento"))
        if p_nmovimiento & "">"" then
            set rstBlq = Server.CreateObject("ADODB.Recordset")
            rstBlq.open "update movimientos with(UPDLOCK) set mercrecibida=case when mercrecibida=1 then 0 else 1 end where nmovimiento='" & p_nmovimiento & "'",session("dsn_cliente")
            if rstBlq.state<>0 then rstBlq.close
            set rstBlq =nothing
            response.write("OK")
        else
            response.write("NOOK")
        end if
        response.end
    end if
    if request.querystring("consulta")="bloquearMovimiento" then
        errorAjax=0
        p_nmovimiento = limpiaCadena(request.querystring("nmovimiento"))
        p_serie=limpiaCadena(request.querystring("serie"))
        saft=limpiaCadena(request.querystring("saft"))
        lsearch=limpiaCadena(request.querystring("lsearch"))
''response.write("los datos 1 son-" & p_nmovimiento & "-" & p_serie & "-" & saft & "-" & lsearch & "-<br>")
        if p_nmovimiento & "">"" then
            saftOK=true
''response.write("los datos 2 son-" & cstr(saft) & "-<br>")
            if cstr(saft)="1" then
                'comprobamos ahora que las lineas de detalle sean correctas
                detalle =false
                if saftOK=true then
                    set rstBlq = Server.CreateObject("ADODB.Recordset")
                    rstBlq.cursorlocation=3
                    rstBlq.open "select top 1 nmovimiento from detalles_movimientos with(NOLOCK) where nmovimiento='" & p_nmovimiento & "'",session("dsn_cliente")
                    if not rstBlq.eof then
                        detalle =true
                    end if
                    rstBlq.close
                    set rstBlq =nothing
                end if
''response.write("los datos 3 son-" & detalle & "-" & saftOK & "-<br>")
                if  detalle =false and saftOK=true then
                    saftOK=false
                    %>
                    <script language="javascript" type="text/javascript">
                        //window.alert("<%=LITDETALLESVACIO%>");
                    </script> 
                    <%
                    errorAjax=5
                end if
''response.write("los datos 4 son-" & p_nmovimiento & "-" & factAnt & "-<br>")
                ''on error resume next
                factAnt = obtenerFacturaAnterior (p_nmovimiento)
                ''if err.number<>0 then
                ''    response.write(5/0)
                ''    response.end
                ''end if
                ''on error goto 0
''response.write("los datos 5 son-" & saftOK & "-" & p_nmovimiento & "-" & factAnt & "-<br>")
                if saftOK=true then
                    if factAnt<>"" then
                        if p_nmovimiento & "">"" then
                            ''on error resume next
                            anyo = d_lookup("YEAR(FECHA)","movimientos","nmovimiento = '" & p_nmovimiento & "'",session("dsn_cliente"))
                            ''if err.number<>0 then
                            ''    response.write(1/0)
                            ''    response.end
                            ''end if
                            ''on error goto 0
                        else
                            anyo=cstr(year(now()))
                        end if
                        ''on error resume next
                        factAnt=D_LOOKUP("MAX(NMOVIMIENTO)","MOVIMIENTOS","NSERIE='" & p_serie & "' AND NMOVIMIENTO<'" & p_nmovimiento & "' AND NMOVIMIENTO LIKE '" & SESSION("NCLIENTE") & "%' AND YEAR(FECHA)=" & anyo,session("dsn_cliente"))
                        ''if err.number<>0 then
                        ''    response.write(2/0)
                        ''    response.end
                        ''end if
                        ''on error goto 0
                        if factAnt&""="" then ''Si es la primera factura
                            hash="aaaaaaaaaaXXXXXXXXXXZZZZZZZZZZ1111111111" 
                        else
                            ''on error resume next
                            hash = d_lookup("hash","movimientos","nmovimiento = '" & factAnt & "'",session("dsn_cliente"))
                            ''if err.number<>0 then
                            ''    response.write(3/0)
                            ''    response.end
                            ''end if
                            ''on error goto 0
                        end if 
	                    
                        if hash <> "" then
                            ''on error resume next
                            protegerFacturaSAFT p_nmovimiento
                            ''if err.number<>0 then
                            ''    response.write(6/0)
                            ''    response.end
                            ''end if
                            ''on error goto 0
                            saftOK=true
                        else
                            'hacemos un hack para la primera factura
                            if hash="aaaaaaaaaaXXXXXXXXXXZZZZZZZZZZ1111111111" then
                                ''on error resume next
                                protegerFacturaSAFT p_nmovimiento
                                ''if err.number<>0 then
                                ''    response.write(7/0)
                                ''    response.end
                                ''end if
                                ''on error goto 0
                                saftOK=true
                            else
                                %>
                                <script language="javascript" type="text/javascript">
                                    //window.alert("<%=LitFactAntNoBloqueada%>");
                                </script> 
                                <%
                                errorAjax=6
                                saftOK=false
                            end if
                        end if
                    else
''response.write("los datos 6 son-" & p_nmovimiento & "-<br>")
                        ''on error resume next
                        protegerFacturaSAFT p_nmovimiento
                        ''if err.number<>0 then
                        ''    response.write(8/0)
                        ''    response.end
                        ''end if
                        ''on error goto 0
                        saftOK=true
''response.write("los datos 7 son-" & p_nmovimiento & "-<br>")
                        ''response.Write saft&"-5-"& saftOK
                        errorAjax=7
                    end if
                end if 
            end if
''response.write("los datos 8 son-" & cstr(saft) & "-" & saftOK & "-" & p_nmovimiento & "-<br>")
            if cstr(saft)<>"1" or saftOK=true then
                set rstBlq = Server.CreateObject("ADODB.Recordset")
                ''on error resume next
                rstBlq.open "update movimientos with(rowlock) set bloqueado=1 where nmovimiento = '" & p_nmovimiento & "'",session("dsn_cliente")
                ''if err.number<>0 then
                ''    response.write(9/0)
                ''    response.end
                ''end if
                ''on error goto 0
                set rstBlq = nothing
''response.write("los datos 9 son-" & session("usuario") & "-" & p_nmovimiento & "-<br>")
''response.write("los datos 10 son-" & dsnilion & "-<br>")
''response.end
                auditar_ins_bor session("usuario"),p_nmovimiento,"defecto","bloqueo","","","movimientos_almacenes"
            end if
            ''viene=limpiaCadena(request.querystring("viene"))
            ''if viene="" then viene=limpiaCadena(request.form("viene"))
            ''if viene = "search" then
            ''    mode="search"
            'else
            ''    mode="browse"
            ''end if
            response.write("RESBLOQ=" & errorAjax & "LSEARCH=" & lsearch & "&P_NMOVIMIENTO=" & mid(p_nmovimiento,6,len(p_nmovimiento)))
        else
            response.write("ERROR-LSEARCH=" & lsearch)
        end if
        ' Fin de consulta AJAX
        response.End
    end if
    if request.querystring("consulta")="desbloquearMovimiento" then
        errorAjax=0
        p_nmovimiento = limpiaCadena(request.querystring("nmovimiento"))
        p_serie=limpiaCadena(request.querystring("serie"))
        saft=limpiaCadena(request.querystring("saft"))
        lsearch=limpiaCadena(request.querystring("lsearch"))
        if p_nmovimiento & "">"" then
            saftOK=false
            desbloquearSAFT=""
            if cstr(saft)="1" then
                MovSig = obtenerFacturaSiguiente (p_nmovimiento)
                ''on error resume next
                ultimoMov = d_lookup("hash","movimientos","nmovimiento='" & MovSig & "'",session("dsn_cliente"))
                ''if err.number<>0 then
                ''    response.write(10/0)
                ''    response.end
                ''end if
                ''on error goto 0
                if ultimoMov & "" = "" then
                    saftOK=true
                    desbloquearSAFT=",hash = null, saftnmovimiento=null "
                    errorAjax=1
                else
                    %>
                    <script language="javascript" type="text/javascript">
                        //window.alert("d1<%=LITULTIMOMOV%>");
                    </script> 
                    <%
                    errorAjax=2
                end if
            end if
	
            if cstr(saft)<>"1" or saftOK=true  then
                'utilizamos la variable desbloquearSAFT para añadir los campos a modificar en caso de que quereamos desbloquear una factura de saft
                set rstBlq = Server.CreateObject("ADODB.Recordset")
                rstBlq.open "update movimientos with(rowlock) set bloqueado=0 where nmovimiento = '" & p_nmovimiento & "'",session("dsn_cliente")
                set rstBlq = nothing
                auditar_ins_bor session("usuario"),p_nmovimiento,"defecto","desbloqueo","","","movimientos"
            end if
		
            ''viene=limpiaCadena(request.querystring("viene"))
            ''if viene="" then viene=limpiaCadena(request.form("viene"))
            ''if viene ="search" then
            ''    mode="search"
            ''else
            ''    mode="browse"
            ''end if
            response.write("RESDESBLOQ=" & errorAjax & "LSEARCH=" & lsearch & "&P_NMOVIMIENTO=" & mid(p_nmovimiento,6,len(p_nmovimiento)))
        else
            response.write("ERROR-LSEARCH=" & lsearch)
        end if
        ' Fin de consulta AJAX
        response.End
    end if
    if request.querystring("consulta")="ComprobarVentasConLoteCab" then
        errorAjax=0
        p_nmovimiento = limpiaCadena(request.querystring("nmovimiento"))
        modo = limpiaCadena(request.querystring("modo"))
        if p_nmovimiento & "">"" then
            set rstBlq = Server.CreateObject("ADODB.Recordset")
            StrSelect="select distinct dm.idlotecab,da.nalbaran as ndocumento "
            StrSelect=StrSelect & " from detalles_movimientos as dm with(NOLOCK),detalles_alb_cli as da with(NOLOCK) "
            StrSelect=StrSelect & " where dm.nmovimiento like '" & session("ncliente") & "%' and da.nalbaran like '" & session("ncliente") & "%' "
            StrSelect=StrSelect & " and dm.nmovimiento = '" & p_nmovimiento & "' and dm.idlotecab is not null "
            StrSelect=StrSelect & " and da.idlote=dm.idlotecab "
            StrSelect=StrSelect & " union all "
            StrSelect=StrSelect & " select distinct dm.idlotecab,da.nfactura as ndocumento " 
            StrSelect=StrSelect & " from detalles_movimientos as dm with(NOLOCK),detalles_fac_cli as da with(NOLOCK) "
            StrSelect=StrSelect & " where dm.nmovimiento like '" & session("ncliente") & "%' and da.nfactura like '" & session("ncliente") & "%' "
            StrSelect=StrSelect & " and dm.nmovimiento = '" & p_nmovimiento & "' and dm.idlotecab is not null "
            StrSelect=StrSelect & " and da.idlote=dm.idlotecab "
            StrSelect=StrSelect & " union all "
            StrSelect=StrSelect & " select distinct dm.idlotecab,da.nmovimiento as ndocumento "
            StrSelect=StrSelect & " from detalles_movimientos as dm with(NOLOCK),detalles_movimientos as da with(NOLOCK) "
            StrSelect=StrSelect & " where dm.nmovimiento like '" & session("ncliente") & "%' and da.nmovimiento like '" & session("ncliente") & "%' "
            StrSelect=StrSelect & " and dm.nmovimiento = '" & p_nmovimiento & "' and dm.idlotecab is not null "
            StrSelect=StrSelect & " and da.idlote=dm.idlotecab"
            rstBlq.cursorlocation=3
            rstBlq.open StrSelect,session("dsn_cliente")
            if not rstBlq.eof then
                response.write("NOOK=" & modo)
            else
                response.write("OK=" & modo)
            end if
            rstBlq.close
            set rstBlq = nothing
        end if
        ' Fin de consulta AJAX
        response.End
    end if
end if
%>
<%
' JCI 17/06/2003 : MIGRACION A MONOBASE
'
' JCI 27/01/2004 : Gestión de lotes
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>



<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->

<!--#include file="movimientos_almacenes.inc" -->
<!--#include file="../ventas/documentos.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../js/dropdown.js.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('AddDG', 'fade=1')
    animatedcollapse.addDiv('BrowseDG', 'fade=1')
    //animatedcollapse.addDiv('BrowseCab', 'fade=1')
    animatedcollapse.addDiv('CABECERA', 'fade=1')
    animatedcollapse.addDiv('DETALLES', 'fade=1')

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()

</script>
<script language="javascript" type="text/javascript">
    comprobacionVentasEnLotes = 0;
    function MarcarMecrecibida(movimiento) {
        h_mercrecibida = document.movimientos_almacenes.h_mercrecibida.value;
        if (h_mercrecibida == "-1") {
            mercrecibida = false;
        }
        else {
            mercrecibida = true;
        }
        texto_aviso_MR = "";

        almdestino_nom = "";
        if (mercrecibida == false && h_mercrecibida == "-1") {
            if (document.movimientos_almacenes.almdestino != null) {
                try {
                    almdestino_nom = document.movimientos_almacenes.nom_almacen.value;
                }
                catch (e) {
                    almdestino_nom = "";
                }
            }
            if (almdestino_nom != null && almdestino_nom != "") {
                texto_aviso_MR = "<%=LitEstSegMarcRec3%>" + almdestino_nom + ".<%=LitEstSegMarcRec2%>";
            }
            else {
                texto_aviso_MR = "";
            }
        }
        de_mercrecibida_a_mercrecibida = 0;
        if (mercrecibida == true && h_mercrecibida == "0") {
            de_mercrecibida_a_mercrecibida = 1;
            if (document.movimientos_almacenes.almdestino != null) {
                try {
                    almdestino_nom = document.movimientos_almacenes.nom_almacen.value;
                }
                catch (e) {
                    almdestino_nom = "";
                }
            }
            if (almdestino_nom != null && almdestino_nom != "") {
                texto_aviso_MR = "<%=LitEstSegMarcRec1%>" + almdestino_nom + ".<%=LitEstSegMarcRec2%>";
            }
            else {
                texto_aviso_MR = "";
            }
        }
        if (texto_aviso_MR != "") {
            if (window.confirm(texto_aviso_MR) == true) {
                if (!enProceso && http) {
                    //document.getElementById("waitBoxOculto").style.visibility = "visible";
                    var url = "movimientos_almacenes.asp?mode=consultaAJAX&consulta=MarcarMercanciaRecibida&nmovimiento=" + movimiento;
                    http.open("GET", url, true);
                    http.onreadystatechange = handleHttpResponseCoste2;
                    enProceso = true;
                    http.send(null);
                    //document.getElementById("waitBoxOculto").style.visibility = "hidden";
                }
            }
        }
    }
    function IntroducirCodeAt(movimiento) {
        var CodeAt = prompt("Introduce el codigo AT", "");
        if (CodeAt != null && CodeAt != "") {
            if (!enProceso && http) {
                //document.getElementById("waitBoxOculto").style.visibility = "visible";
                var url = "movimientos_almacenes.asp?mode=consultaAJAX&consulta=introducirCodeAt&nmovimiento=" + movimiento + "&codeat=" + CodeAt;
                http.open("GET", url, true);
                http.onreadystatechange = handleHttpResponseCoste2;
                enProceso = true;
                http.send(null);
                //document.getElementById("waitBoxOculto").style.visibility = "hidden";
            }
        }
    }
    function desbloqueoMovimiento(movimiento, serie, saft, lsearch) {
        if (confirm("<%=LitMsgDesBloqueo%>")) {
            try {
                document.movimientos_almacenes.bloqueado.value = 0;
            }
            catch (e) {
                try {
                    parent.window.main.pantalla.document.movimientos_almacenes.bloqueado.value = 0;
                }
                catch (e) {
                    try {
                        parent.window.pantalla.document.movimientos_almacenes.bloqueado.value = 0;
                    }
                    catch (e) {
                    }
                }
            }
            //document.movimientos_almacenes.action = "movimientos_almacenes.asp?nmovimiento=" + movimiento + "&mode=bloqueo";
            //document.movimientos_almacenes.submit();
            //parent.botones.document.location = "movimientos_almacenes_bt.asp?mode=browse";
            if (!enProceso && http) {
                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                //date = new Date().getTime();
                var url = "movimientos_almacenes.asp?mode=consultaAJAX&consulta=desbloquearMovimiento&nmovimiento=" + movimiento + "&serie=" + serie + "&saft=" + saft + "&lsearch=" + lsearch;
                http.open("GET", url, true);
                http.onreadystatechange = handleHttpResponseCoste;
                enProceso = true;
                http.send(null);
                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "hidden";
            }
        }
    }
    function bloqueoMovimiento(movimiento, serie, saft, lsearch) {
        if (confirm("<%=LitMsgBloqueo%>")) {
            try {
                document.movimientos_almacenes.bloqueado.value = 1;
            }
            catch (e) {
                try {
                    parent.window.main.pantalla.document.movimientos_almacenes.bloqueado.value = 1;
                }
                catch (e) {
                    try {
                        parent.window.pantalla.document.movimientos_almacenes.bloqueado.value = 1;
                    }
                    catch (e) {
                    }
                }
            }
            //document.movimientos_almacenes.action = "movimientos_almacenes.asp?nmovimiento=" + movimiento + "&mode=bloqueo";
            //document.movimientos_almacenes.submit();
            //parent.botones.document.location = "movimientos_almacenes_bt.asp?mode=browse";

            if (!enProceso && http) {
                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "visible";
                //date = new Date().getTime();
                var url = "movimientos_almacenes.asp?mode=consultaAJAX&consulta=bloquearMovimiento&nmovimiento=" + movimiento + "&serie=" + serie + "&saft=" + saft + "&lsearch=" + lsearch;
                http.open("GET", url, true);
                http.onreadystatechange = handleHttpResponseCoste;
                enProceso = true;
                http.send(null);
                parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "hidden";
            }
        }
    }

    function ComprobarVentasConLoteCab(movimiento, modo) {
        if (!enProceso && http) {
            document.getElementById("waitBoxOculto").style.visibility = "visible";
            var url = "movimientos_almacenes.asp?mode=consultaAJAX&consulta=ComprobarVentasConLoteCab&nmovimiento=" + movimiento + "&modo=" + modo;
            http.open("GET", url, false);
            http.onreadystatechange = handleHttpResponseCoste3;
            enProceso = true;
            http.send(null);
            document.getElementById("waitBoxOculto").style.visibility = "hidden";
        }
    }
    function handleHttpResponseCoste3() {
        if (http.readyState == 4) {
            if (http.status == 200) {
                //window.alert("paso 1");
                if (http.responseText != "") {
                    results = http.responseText;
                    if (results == "" || results.toUpperCase().search("ERROR") != -1 || results.toUpperCase().search("NOOK") != -1) {
                        //alert("<%=LitError%>");
                        donde = 0;
                        mensajeDesBloq = "";
                        donde = results.toUpperCase().search("NOOK");
                        if (donde != -1) {
                            mensajeDesBloq = results.substring(donde + "NOOK=".length);
                        }
                        if (mensajeDesBloq.toUpperCase().search("DELETE") != -1) {
                            //window.alert("<%=LITMSGNOBORRARPORLOTE%>");
                            comprobacionVentasEnLotes = 1;
                        }
                        else {
                            //window.alert("<%=LITNODESBLMARCRECPORLOTE%>");
                            comprobacionVentasEnLotes = 2;
                        }
                    }
                    else {
                        comprobacionVentasEnLotes = 0;
                    }
                    enProceso = false;
                }
            }
        }
    }
    function handleHttpResponseCoste2() {
        if (http.readyState == 4) {
            if (http.status == 200) {
                //window.alert("paso 1");
                if (http.responseText != "") {
                    results = http.responseText;
                    if (results == "" || results.toUpperCase().search("ERROR") != -1) {
                        //alert("<%=LitError%>");
                    }
                    else {
                    }
                    enProceso = false;
                }
            }
        }
    }
    function handleHttpResponseCoste() {
        if (http.readyState == 4) {
            if (http.status == 200) {
                //window.alert("paso 1");
                if (http.responseText != "") {
                    parent.pantalla.document.getElementById("waitBoxOculto").style.visibility = "hidden";
                    results = http.responseText;
                    //window.alert(results);
                    if (results == "" || results.toUpperCase().search("ERROR") != -1) {
                        //alert("<%=LitError%>");
                    }
                    else {
                        LSEARCH = "0";
                        LSEARCHPOS = 0;
                        errorBloD = 0;
                        LSEARCHPOS = results.toUpperCase().search("LSEARCH");
                        //window.alert("LSEARCHPOS-" + LSEARCHPOS + "-");
                        if (LSEARCHPOS != 0) {
                            LSEARCH = results.substring(LSEARCHPOS + "LSEARCH=".length, LSEARCHPOS + "LSEARCH=".length + 1);
                        }
                        //window.alert("LSEARCH-" + LSEARCH + "-");
                        //if (LSEARCH == "0") {
                        //ocultar detalles
                        ocultar_detalles = 0;
                        bloqueoODes = 0;
                        mensajeResDesBloq = "";
                        donde = results.toUpperCase().search("RESDESBLOQ");
                        if (donde != -1) {
                            mensajeResDesBloq = results.substring(donde + "RESDESBLOQ=".length, donde + "RESDESBLOQ=".length + 1);
                        }
                        if (mensajeResDesBloq != "") {
                            bloqueoODes = 0;
                            if (mensajeResDesBloq == "2") {
                                window.alert("<%=LITULTIMOMOV%>");
                                ocultar_detalles = 1;
                                errorBloD = 1;
                            }
                            if (mensajeResDesBloq == "1") {
                                //todo ok
                                ocultar_detalles = 0;
                                errorBloD = 0;
                            }
                        }

                        mensajeResBloq = "";
                        donde = 0;
                        donde = results.toUpperCase().search("RESBLOQ");
                        if (donde != -1) {
                            mensajeResBloq = results.substring(donde + "RESBLOQ=".length, donde + "RESBLOQ=".length + 1);
                        }
                        //window.alert(mensajeResBloq);
                        if (mensajeResBloq != "") {
                            bloqueoODes = 1;
                            if (mensajeResBloq == "5") {
                                window.alert("<%=LITDETALLESVACIO%>");
                                ocultar_detalles = 0;
                                errorBloD = 1;
                            }
                            if (mensajeResBloq == "6") {
                                window.alert("<%=LitFactAntNoBloqueada%>");
                                ocultar_detalles = 0;
                                errorBloD = 1;
                            }
                            if (mensajeResBloq == "7") {
                                //todo ok
                                ocultar_detalles = 1;
                                errorBloD = 0;
                            }
                        }


                        mensajeMov = "";
                        donde = 0;
                        donde = results.toUpperCase().search("P_NMOVIMIENTO");
                        if (donde != -1) {
                            mensajeMov = results.substring(donde + "P_NMOVIMIENTO=".length);
                        }

                        if (LSEARCH == "0") {
                            //window.alert(ocultar_detalles);
                            if (ocultar_detalles == 1) {
                                if (document.getElementById("frDetallesIns") != null) {
                                    document.getElementById("frDetallesIns").style.display = "none";
                                }
                                if (document.getElementById("frDetallesInsT") != null) {
                                    document.getElementById("frDetallesInsT").style.display = "none";
                                }
                                h_almacenSerie = document.movimientos_almacenes.h_almacenSerie.value;
                                if (h_almacenSerie != "") {
                                    if (document.getElementById("capa_impresion") != null) {
                                        document.getElementById("capa_impresion").style.display = "";
                                    }
                                    if (document.getElementById("idPrintFormat") != null) {
                                        document.getElementById("idPrintFormat").style.display = "";
                                    }
                                }
                                else {
                                    if (document.getElementById("capa_impresion") != null) {
                                        document.getElementById("capa_impresion").style.display = "none";
                                    }
                                    if (document.getElementById("idPrintFormat") != null) {
                                        document.getElementById("idPrintFormat").style.display = "none";
                                    }
                                }

                                if (document.getElementById("celdaBloquear") != null) {
                                    document.getElementById("celdaBloquear").style.display = "none";
                                }
                                if (document.getElementById("celdaDesBloquear") != null) {
                                    document.getElementById("celdaDesBloquear").style.display = "";
                                }
                                if (document.getElementById("textobloqueado") != null) {
                                    document.getElementById("textobloqueado").style.display = "";
                                }
                                if (document.getElementById("textodebloqueado") != null) {
                                    document.getElementById("textodebloqueado").style.display = "none";
                                }
                                try {
                                    fr_Detalles.document.bloqueado.value = 1;
                                }
                                catch (e) {
                                }
                                try {
                                    fr_Detalles.document.movimiento_bloqueado.value = 1;
                                }
                                catch (e) {
                                }
                                if (document.getElementById("linksaft") != null) {
                                    document.getElementById("linksaft").style.display = "";
                                }
                            }
                            else {
                                if (document.getElementById("frDetallesIns") != null) {
                                    document.getElementById("frDetallesIns").style.display = "";
                                }
                                if (document.getElementById("frDetallesInsT") != null) {
                                    document.getElementById("frDetallesInsT").style.display = "";
                                }
                                if (document.getElementById("capa_impresion") != null) {
                                    document.getElementById("capa_impresion").style.display = "none";
                                }
                                if (document.getElementById("idPrintFormat") != null) {
                                    document.getElementById("idPrintFormat").style.display = "none";
                                }
                                if (document.getElementById("celdaBloquear") != null) {
                                    document.getElementById("celdaBloquear").style.display = "";
                                }
                                if (document.getElementById("celdaDesBloquear") != null) {
                                    document.getElementById("celdaDesBloquear").style.display = "none";
                                }
                                if (document.getElementById("textobloqueado") != null) {
                                    document.getElementById("textobloqueado").style.display = "none";
                                }
                                if (document.getElementById("textodebloqueado") != null) {
                                    document.getElementById("textodebloqueado").style.display = "";
                                }
                                try {
                                    fr_Detalles.document.bloqueado.value = 0;
                                }
                                catch (e) {
                                }
                                try {
                                    fr_Detalles.document.movimiento_bloqueado.value = 0;
                                }
                                catch (e) {
                                }
                                if (document.getElementById("linksaft") != null) {
                                    document.getElementById("linksaft").style.display = "none";
                                }
                            }
                        }
                        if (errorBloD == 0) {
                            if (bloqueoODes == 1) {
                                window.alert("<%=LITMOVBLOQEXI%>".replace("XXXXXXXX", mensajeMov));
                            }
                            else {
                                window.alert("<%=LITMOVDESBLOQEXI%>".replace("XXXXXXXX", mensajeMov));
                            }
                        }
                        //}
                    }
                    enProceso = false;
                }
            }
        }
    }

    if (!window.XMLHttpRequest) {
        window.XMLHttpRequest = function () {
            return new ActiveXObject('Microsoft.XMLHTTP');
        }
    }

    var http = new XMLHttpRequest();
    var enProceso = false; // lo usamos para ver si hay un proceso activo
</script>

<%
''Dim bma
''bma=limpiaCadena(request.querystring("bma")&"")
''if bma = "" then
''	bma=limpiaCadena(request.form("bma")&"")
''end if
''ObtenerParametros("movimientos_almacenes")

''dgb  11-04-2008  MODULO CENTROXOGO
si_tiene_modulo_Centroxogo=ModuloContratado(session("ncliente"),ModCentroxogo)%>

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    function cambiarfecha(fecha, modo) {
        var fecha_ar = new Array();

        if (fecha != "") {
            suma = 0;
            fecha_ar[suma] = "";
            l = 0
            while (l <= fecha.length) {
                if (fecha.substring(l, l + 1) == '/') {
                    suma++;
                    fecha_ar[suma] = "";
                }
                else {
                    if (fecha.substring(l, l + 1) != '') fecha_ar[suma] = fecha_ar[suma] + fecha.substring(l, l + 1);
                }
                l++;
            }
            if (suma != 2) {
                window.alert("<%=LitFechaMal%> en el campo " + modo);
                return false;
            }
            else {
                nonumero = 0;
                while (suma >= 0 && nonumero == 0) {
                    if (isNaN(fecha_ar[suma])) nonumero = 1;
                    if (fecha_ar[suma].length > 2 && suma != 2) nonumero = 1;
                    if (fecha_ar[suma].length > 4 && suma == 2) nonumero = 1;
                    suma--;
                }

                if (nonumero == 1) {
                    window.alert("<%=LitFechaMal%> en el campo " + modo);
                    return false;
                }
            }
        }
        return true;
    }


    //***************************************************************************
    /*
    function Editar(albaran)
    {
    document.movimientos_almacenes.action="movimientos_almacenes.asp?nmovimiento=" + albaran + "&mode=browse";
    document.movimientos_almacenes.submit();
    parent.botones.document.location="movimientos_almacenes_bt.asp?mode=browse&bma=<%=bma%>";
    }
    */
    //***************************************************************************
    /*
    function Mas(sentido,lote,campo,criterio,texto)
    {
    document.location="movimientos_almacenes.asp?mode=search&viene=" + document.movimientos_almacenes.viene.value + "&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&mmp=" +document.movimientos_almacenes.mmp.value ;
    }
    */
    //***************************************************************************
    /*
    if (window.document.addEventListener) {
    window.document.addEventListener("keydown", callkeydownhandler, false);
    } else {
    window.document.attachEvent("onkeydown", callkeydownhandler);
    }
    function callkeydownhandler(evnt) {
    ev = (evnt) ? evnt : event;
    //comprobar_enter(ev);
    keypress2(ev);
    }

    function keypress2(e)
    {
    tecla=e.keyCode;
    keyPressed(tecla);
    }

    //Comprueba si la tecla pulsada es CTRL+S. Si es así guarda el registro.
    function keyPressed(tecla)
    {
    if (tecla==<%=TeclaGuardar%>)
    { //CTRL+S
    if (document.movimientos_almacenes.mode.value=="add" || document.movimientos_almacenes.mode.value=="edit")
    {
    if (document.movimientos_almacenes.fecha.value=="")
    {
    window.alert("<%=LitMsgFechaNoNulo%>");
    return;
    }

    if (!cambiarfecha(document.movimientos_almacenes.fecha.value,"FECHA MOVIMIENTO")) return false;

    if (!checkdate(document.movimientos_almacenes.fecha))
    {
    window.alert("<%=LitMsgFechaFecha%>");
    return;
    }

    if (document.movimientos_almacenes.responsable.value=="")
    {
    window.alert("<%=LitMsgResponsableNoNulo%>");
    return false;
    }

    if (document.movimientos_almacenes.nserie.value=="")
    {
    window.alert("<%=LitMsgSerieNoNulo%>");
    return;
    }

    if (document.movimientos_almacenes.almdestino.value=="")
    {
    window.alert("<%=LitMsgAlmDestinoNoNulo%>");
    return false;
    }

    switch (document.movimientos_almacenes.mode.value)
    {
    case "add":
    document.movimientos_almacenes.action="movimientos_almacenes.asp?mode=first_save";
    break;

    case "edit":
    document.movimientos_almacenes.action="movimientos_almacenes.asp?mode=save&ndoc=" + document.movimientos_almacenes.h_nmovimiento.value;
    break;
    }
    document.movimientos_almacenes.submit();
    parent.botones.document.location="movimientos_almacenes_bt.asp?mode=browse&bma=<%=bma%>";
    }
    //else { //Mode=browse.
    //}
    }
    }
    */
    //***************************************************************************

    function TraerResponsable() {
        document.movimientos_almacenes.action = "movimientos_almacenes.asp?responsable=" + document.movimientos_almacenes.responsable.value + "&mode=traerresponsable&submode=" + document.movimientos_almacenes.mode.value + "&nmovimiento=" + document.movimientos_almacenes.h_nmovimiento.value;
        document.movimientos_almacenes.submit();
    }

    function MasDet(sentido, lote, firstReg, lastReg, campo, criterio, texto, firstRegAll, lastRegAll) {
        frDetalles.document.movimientos_almacenes_det.action = "movimientos_almacenes_det.asp?mode=browse&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&firstReg=" + firstReg + "&lastReg=" + lastReg + "&firstRegAll=" + firstRegAll + "&lastRegAll=" + lastRegAll;
        frDetalles.document.movimientos_almacenes_det.submit();
    }
    function Redimensionar(modo) {
        //window.alert("los datos son-" + modo + "-");
        if (modo == "browse" || modo == "save" | modo == "first_save") {
            var alto = 0;
            if (parent.document.body.offsetHeight) alto = parent.document.body.offsetHeight;
            else alto = parent.self.innerHeight;
            if (document.getElementById("frDetalles").style.display == "") {
                if (alto > 140) {
                    if (alto - 310 > 140) {
                        //window.alert("1--->" + (alto - 310) + "-");
                        document.getElementById("frDetalles").style.height = alto - 310;
                    }
                    else {
                        //window.alert("2--->140-");
                        document.getElementById("frDetalles").style.height = 140;
                    }
                }
                else {
                    //window.alert("3--->140-");
                    document.getElementById("frDetalles").style.height = 140;
                }
            }
        }
    }
</script>
<%
mode=request.querystring("mode")
%>
<body class="BODY_ASP" onresize="javascript:Redimensionar('<%=enc.EncodeForJavascript(mode)%>');">
<%

'**
'* Obtiene los números de serie de un detalle de un documento.
'* ndocumento: Indica el nº de documento a buscar.
'* ndetalle: Indica el nº de item del documento.
'* return: La lista de nº de serie del item del documento.
'**

'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(nmovimiento,nserie,fecha,mode)
	if nmovimiento & ""="" then
		'Crear un nuevo registro.
		rst.AddNew
		SigDoc=CalcularNumDocumento(nserie,fecha)
		rst("nmovimiento")=SigDoc
		'nmovimiento=SigDoc
	end if

	'Asignar los nuevos valores a los campos del recordset.
	rst("nserie")=Nulear(limpiaCadena(Request.Form("nserie")))
	'Detectar cambios en el almdestino para recalcular los stock de los detalles
	if nmovimiento & "">"" and rst("almdestino")<>(limpiaCadena(Request.Form("almdestino"))&"") then
		'Primero se elimina el stock y los equipos del almdestino anterior
		dim listaNS

		'contamos las cantidades de equipos que hay en este movimiento
		strselect1="select max(nitem) as total from movimientos as m with (NOLOCK),detalles_movimientos as d with (NOLOCK) where m.nmovimiento like '" & session("ncliente") & "%' and d.nmovimiento like '" & session("ncliente") & "%' and m.nmovimiento=d.nmovimiento and d.nmovimiento='" & nmovimiento & "'"
		rstAux.cursorlocation=3
		rstAux.open strselect1,session("dsn_cliente")
		if not rstAux.EOF then
			if rstAux("total") & "">"" then
				redim listaNS(rstAux("total") + 1)
			end if
		end if
		rstAux.close

		'hacemos las operaciones
		rstAux.cursorlocation=3
		rstAux.open "select ref,almorigen,cantidad,almdestino,nitem from movimientos as m with (NOLOCK),detalles_movimientos as d with (NOLOCK) where m.nmovimiento like '" & session("ncliente") & "%' and d.nmovimiento like '" & session("ncliente") & "%' and m.nmovimiento=d.nmovimiento and d.nmovimiento='" & nmovimiento & "'",session("dsn_cliente")
		while not rstAux.EOF
			refST=rstAux("ref")
			almSTO=rstAux("almorigen")
			almSTD=rstAux("almdestino")
			canST=rstAux("cantidad")

			ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTO,canST,"",session("dsn_cliente")
			ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTD,-canST,"",session("dsn_cliente")
			rstAux.movenext
		wend
		rstAux.close
		'---------------------
		'Ahora se añade el stock y los equipos al almdestino nuevo
		rstAux.cursorlocation=3
		rstAux.open "select ref,almorigen,cantidad,almdestino,nitem from movimientos as m with (NOLOCK),detalles_movimientos as d with (NOLOCK) where m.nmovimiento like '" & session("ncliente") & "%' and d.nmovimiento like '" & session("ncliente") & "%' and m.nmovimiento=d.nmovimiento and d.nmovimiento='" & nmovimiento & "'",session("dsn_cliente")
		while not rstAux.EOF
			refST=rstAux("ref")
			almSTO=rstAux("almorigen")
			almSTD=limpiaCadena(Request.Form("almdestino"))
			canST=rstAux("cantidad")

			ListaN3=split(listaNS(rstAux("nitem")),chr(13)& chr(10),-1,1)
			k3=1
			Redim listaNumeros3(k3)
			for i=0 to ubound(ListaN3)
				if len(ListaN3(i))>0 then
					listaNumeros3(k3)=ListaN3(i)
					k3=k3+1
					Redim Preserve listaNumeros3(k3)
				end if
			next

			ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTO,-canST,"",session("dsn_cliente")
			ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTD,canST,"",session("dsn_cliente")
			rstAux.movenext
		wend
		rstAux.close
	end if

	rst("almdestino")=Nulear(limpiaCadena(Request.Form("almdestino")))
	rst("responsable")=session("ncliente") + Nulear(limpiaCadena(request.form("responsable")))
	rst("observaciones")=Nulear(limpiaCadena(request.form("observaciones")))
	rst("fecha")=Nulear(limpiaCadena(Request.Form("fecha")))
	rst("mercrecibida")=iif(cstr(Request.Form("merric") & "")="-1" or ucase(cstr(Request.Form("merric") & ""))="ON",1,0)
    rst("SAFTMOVEMENTSTARTTIME")=Nulear(limpiaCadena(Request.Form("SAFTMOVEMENTSTARTTIME")))
    rst("SAFTATDOCCODEID")=Nulear(limpiaCadena(Request.Form("SAFTATDOCCODEID")))
	rst.Update
end sub

'******************************************************************************
'Elimina los datos del registro cuando se pulsa BORRAR.
sub BorrarRegistro(nmovimiento)

	'Miramos si se va a borrar el último generado y si es así se descuenta el contador de documentos
	ano=right(cstr(year(d_lookup("fecha","movimientos","nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & nmovimiento & "'",session("dsn_cliente")))),2)
	nserie=d_lookup("nserie","movimientos","nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & nmovimiento & "'",session("dsn_cliente"))
    rstAux.cursorlocation=2
	rstAux.Open "select * from series where nserie like '" & session("ncliente") & "%' and nserie='" & nserie & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if not rstAux.eof then
		ultimo=rstAux("contador")
		UltimoDocumento=nserie+ano+completar(trim(cstr(ultimo)),6,"0")
		if nmovimiento=UltimoDocumento then
			rstAux("contador")=ultimo-1
			rstAux.update
		end if
	end if
	rstAux.close

	'Primero se actualiza el stock
	rstAux.cursorlocation=3
	rstAux.open "select ref,almorigen,cantidad,almdestino,nitem from movimientos as m with (NOLOCK),detalles_movimientos as d with (NOLOCK) where m.nmovimiento like '" & session("ncliente") & "%' and d.nmovimiento like '" & session("ncliente") & "%' and  m.nmovimiento=d.nmovimiento and d.nmovimiento='" & nmovimiento & "'",session("dsn_cliente")
	while not rstAux.EOF
		refST=rstAux("ref")
		almSTO=rstAux("almorigen")
		almSTD=rstAux("almdestino")
		canST=rstAux("cantidad")

		ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTO,canST,"",session("dsn_cliente")
		ActualizaStocks mode,"MOVIMIENTOS ENTRE ALMACENES",refST,almSTD,-canST,"",session("dsn_cliente")
		rstAux.movenext
	wend
	rstAux.close
	'Luego se eliminan los detalles
    rstAux.cursorlocation=2
	rstAux.open "delete from detalles_movimientos with(rowlock) where nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & nmovimiento & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    
    if rstAux.State<>0 then
        rstAux.close
    end if
	'Después la cabecera del pedido
    rstAux.cursorlocation=2
	rstAux.open "delete from movimientos with(rowlock) where nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & nmovimiento & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
end sub

'Formar el recordset para los cuadros desplegables.
function Desplegable(mode,campo,campo2,tabla,dato,por_defecto)
	if mode="add" or mode="edit" then
		if mode="add" and por_defecto<>"" then
			'Valor por defecto al añadir un registro
			Desplegable=d_lookup(campo,tabla,por_defecto,session("dsn_cliente"))
		end if
		
		if campo=campo2 then
            RstAux.cursorlocation=3
			RstAux.open "select " + campo + " from " + tabla + " where " + campo + " like '" & session("ncliente") & "%' order by " & campo,session("dsn_cliente")
		else
            RstAux.cursorlocation=3
			RstAux.open "select " + campo + "," + campo2 + " from " + tabla + " where " + campo + " like '" & session("ncliente") & "%' order by " & campo2,session("dsn_cliente")
		end if

		if mode="edit" or por_defecto="" then Desplegable=dato
	elseif mode="browse" then
		if campo=campo2 then
			Desplegable=dato
		else
			Desplegable=d_lookup(campo2,tabla,campo + "='" & dato & "'",session("dsn_cliente"))
		end if
	end if
end function

    Function protegerFacturaSAFT(p_nmovimiento)    
        Dim returnString 
        Dim SoapRequest 
        Dim SoapURL 

        Set SoapRequest = Server.CreateObject("MSXML2.XMLHTTP") 
        Dim DataToSend
        DataToSend="ncliente="&session("ncliente")&"&usuario="&session("usuario")&"&idSesion="&session.sessionid&"&nmovimiento="&p_nmovimiento
        'response.Write p_nmovimiento&"--"
        'response.End
        SoapURL = "http://81.19.108.49/IlionServices/Integracion/IntegracionS.asmx/protegerMovimiento" 
        
    
        SoapRequest.Open "POST",SoapURL , False 
        SoapRequest.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
        SoapRequest.Send DataToSend
    
        result=SoapRequest.responseText
    ''response.Write "result:"&result
        Set SoapRequest = Nothing 
  
        protegerFacturaSAFT = result 
    End Function 


    Const NUMBER_PADDING = "000000000000" ' a few zeroes more just to make sure ' 
    Function ZeroPadInteger(i, numberOfDigits) 
      ZeroPadInteger = Right(NUMBER_PADDING & i, numberOfDigits) 
    End Function 

    Function obtenerFacturaAnterior(nfactura)    
        longitud = len(nfactura)
        nfactAnt = cint(mid(nfactura, longitud - 5, longitud)) -1
        if nfactAnt = 0 then
            result=""
        else
            nfactFormateado= ZeroPadInteger (nfactAnt,5)
            result = mid(nfactura, 1, longitud - 5)  & nfactFormateado
        end if
        obtenerFacturaAnterior = result 
    End Function

    Function obtenerFacturaSiguiente(nfactura)    
        longitud = len(nfactura)
        nfactSig = cint(mid(nfactura, longitud - 5, longitud)) +1
        if nfactSig = 0 then
            result=""
        else
            nfactFormateado= ZeroPadInteger (nfactSig,5)
            result = mid(nfactura, 1, longitud - 5)  & nfactFormateado
        end if
        obtenerFacturaSiguiente = result 
    End Function  
'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0
	%>
<form name="movimientos_almacenes" method="post"><%
    PintarCabecera "movimientos_almacenes.asp"

    WaitBoxOculto LitEsperePorFavor

	'Leer parámetros de la página
	mode=Request.QueryString("mode")
    %><input type="hidden" name="mode_accesos_tienda" value="<%=enc.EncodeForHtmlAttribute(mode)%>" /><%

	dim bmr,mmt,mmp,rnf, pnf,s,bma

	ObtenerParametros("movimientos_almacenes")
    'ebf 4/9/2013 Parámetro unico de series para visualizar y crear
    if s&""="" then
	     s=limpiaCadena(request.querystring("s"))
	     if s="" then s=limpiaCadena(request.form("s"))
    end if
	s=preparar_lista(s)

    ''response.write("el parametro s es-" & s & "-<br>")

    if bma & ""="" then
        bma=limpiaCadena(request.querystring("bma")&"")
        if bma = "" then
	        bma=limpiaCadena(request.form("bma")&"")
        end if
    end if

	p_nmovimiento=limpiaCadena(Request.QueryString("nmovimiento"))
	if p_nmovimiento="" then p_nmovimiento=limpiaCadena(Request.QueryString("ndoc"))
	if p_nmovimiento>"" then CheckCadena p_nmovimiento

	campo=limpiaCadena(Request.QueryString("campo"))
	criterio=limpiaCadena(Request.QueryString("criterio"))
	texto=limpiaCadena(Request.QueryString("texto"))
	if Request.QueryString("fecha")>"" then
		p_fecha=limpiaCadena(Request.QueryString("fecha"))
	else
		p_fecha=limpiaCadena(Request.form("fecha"))
	end if

	viene=limpiaCadena(request.querystring("viene"))
	if viene="" then viene=limpiaCadena(request.form("viene"))

	if request.querystring("submode")>"" then
		submode=request.querystring("submode")
	else
		submode=request.form("submode")
	end if

	if request.querystring("observaciones")>"" then
		tmp_observaciones=limpiaCadena(request.querystring("observaciones"))
	else
		tmp_observaciones=limpiaCadena(request.form("observaciones"))
	end if

	if request.querystring("responsable")>"" then
		tmp_responsable=limpiaCadena(request.querystring("responsable"))
	else
		tmp_responsable=limpiaCadena(request.form("responsable"))
	end if

	if request.querystring("almdestino")>"" then
		tmp_almdestino=limpiaCadena(request.querystring("almdestino"))
	else
		tmp_almdestino=limpiaCadena(request.form("almdestino"))
	end if

	if request.querystring("nserie")>"" then
		p_serie=limpiaCadena(request.querystring("nserie"))
	else
		p_serie=limpiaCadena(request.form("nserie"))
	end if

	if bmr & ""="" then
		if request.QueryString("bmr")& "">"" then
			bmr=limpiaCadena(request.QueryString("bmr"))
		elseif request.form("bmr") & "">"" then
			bmr=limpiaCadena(request.form("bmr"))
		end if
	end if

	if request.QueryString("merric")& "">"" then
		merric=limpiaCadena(request.QueryString("merric"))
	elseif request.form("merric") & "">"" then
		merric=limpiaCadena(request.form("merric"))
	end if

	'***RGU 5/1/2006***
	if request.QueryString("mmp")& "">"" then
		mmp=limpiaCadena(request.QueryString("mmp"))
	elseif request.form("mmp") & "">"" then
		mmp=limpiaCadena(request.form("mmp"))
	end if
	'***

	%>
	<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
	<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
	<input type="hidden" name="mmp" value="<%=enc.EncodeForHtmlAttribute(mmp)%>"/>
    <input type="hidden" name="s" value="<%=enc.EncodeForHtmlAttribute(s)%>"/>
	<%
    saft = d_lookup("saft","configuracion","nempresa = '" & session("ncliente")&"'",session("dsn_cliente"))
        saft=nz_b(saft)
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstAux3 = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")
    set rstBlq = Server.CreateObject("ADODB.Recordset")
    if p_nmovimiento & "">"" then
		    if comprobar_LS(s,mode,p_nmovimiento,"MOVIMIENTOS")=0 then
			    %><script language="javascript" type="text/javascript">
			          alert("<%=LitMsgDocNoPermAcc%>");
			          document.movimientos_almacenes.action = "movimientos_almacenes.asp?nmovimiento=&mode=add";
			          document.movimientos_almacenes.submit();
			          parent.botones.document.location = "facturas_cli_bt.asp?mode=add";
			    </script><%
			    CerrarTodo()
			    response.end
		    end if
	    end if
	if p_serie="" and mode="add" then
		'Obtener la serie por defecto
		p_serie=d_lookup("nserie","series","tipo_documento='MOVIMIENTOS ENTRE ALMACENES' and pordefecto=1 and nserie like '" & session("ncliente") & "%'", session("dsn_cliente"))
	end if
    ''response.Write saft&"-1-"& saftOK

	si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)

	if mode="save" or mode="first_save" then
		'ricardo 7-4-2004 miramos si el personal existe o no
		if submode2<>"traerresponsable" then
			no_proveedor=0
			strselect="select dni,fbaja from personal with (NOLOCK) where dni='" & session("ncliente") + Nulear(limpiaCadena(request.form("responsable"))) & "'"
			rst.cursorlocation=3
			rst.Open strselect,session("dsn_cliente")
			if rst.eof then
				no_proveedor=1
			else
				if rst("fbaja") & "">"" then
					no_proveedor=2
				end if
			end if
			rst.close
		else
			no_proveedor=0
		end if

		if no_proveedor=0 then
			no_encontrado=0
			if tmp_almdestino&"">"" and mode="save" then
				strselect="select almorigen from detalles_movimientos with (NOLOCK) where nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & iif(p_nmovimiento&"">"",p_nmovimiento,"NULL") & "' and almorigen='" & tmp_almdestino & "'"
				rst.cursorlocation=3
				rst.Open strselect,session("dsn_cliente")
				if not rst.eof then
					no_encontrado=1
				end if
				rst.close
			end if
			if no_encontrado=0 then
			    rst.CursorLocation=2
				strselect="select * from movimientos where nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & iif(p_nmovimiento&"">"",p_nmovimiento,"NULL") & "'"
				rst.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

				mensajeTratEquipos=TratarEquipos("","","MOVIMIENTOS ENTRE ALMACENES",p_nmovimiento,"","","","","",mode)
				if mid(mensajeTratEquipos,1,2)<>"OK" then
					%><script language="javascript" type="text/javascript">
					      window.alert("<%=mensajeTratEquipos%>");
					      parent.botones.document.location = "movimientos_almacenes_bt.asp?mode=edit&bma=<%=enc.EncodeForJavascript(bma)%>";
					</script><%
					ant_mode=mode
					mode="edit"
				else
					ModDocumento=true
					'comprobamos si el nalbaran existe o no segun el contador de configuracion
					if mode="first_save" then
						if compNumDocNuevo(p_serie,p_fecha,"movimientos")=0 then%>
							<script language="javascript" type="text/javascript">
							    window.alert("<%=LitMsgDocExistRevCont%>");
							    history.back();
							    parent.botones.document.location = "movimientos_almacenes_bt.asp?mode=add&bma=<%=enc.EncodeForJavascript(bma)%>"
							</script>
							<%ModDocumento=false
						end if
					end if
					if ModDocumento=true then
						if mode<>"first_save" then
							mercrecibida_old=nz_b(rst("mercrecibida"))
						end if
						GuardarRegistro p_nmovimiento,p_serie,p_fecha,mode
						p_nmovimiento=rst("nmovimiento")
                        ''ricardo 12-08-2013 creamos ahora los lotes de compra
                        if mode & ""="save" then
                            mercrecibida_new=iif(cstr(Request.Form("merric") & "")="-1" or ucase(cstr(Request.Form("merric") & ""))="ON",1,0)
                            FaltaGrabarLotesCompra=cstr(Request.Form("FaltaGrabarLotesCompra"))
                            if FaltaGrabarLotesCompra & ""="" then
                                FaltaGrabarLotesCompra="0"
                            end if
                            ''response.write("los datos de grabar lotes son-" & p_nmovimiento & "-" & FaltaGrabarLotesCompra & "-" & mercrecibida_old & "-" & mercrecibida_new & "-<br>")
                            if cstr(FaltaGrabarLotesCompra & "")="1" then
                                if mercrecibida_old=0 and mercrecibida_new=1 then
                                    ''response.write("ejecutamos el procedimientos SaveLotsMovs-" & p_nmovimiento & "-<br>")
                                    set rstSL = Server.CreateObject("ADODB.Recordset")
                                    set connSL = server.CreateObject("ADODB.Connection")
                                    set commandSL  = server.CreateObject("ADODB.Command")
                                    connSL.open session("dsn_cliente")
                                    commandSL.ActiveConnection = connSL
	                                commandSL.CommandTimeout = 0              
                                    commandSL.CommandText="SaveLotsMovs" 
	                                commandSL.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	                                commandSL.Parameters.Append commandSL.CreateParameter("@ncompany",adVarChar,adParamInput,5,session("ncliente"))
                                    commandSL.Parameters.Append commandSL.CreateParameter("@nmovement",adVarChar,adParamInput,20,p_nmovimiento)
                                    set rstSL = commandSL.Execute
                                    if rstSL.state<>0 then rstSL.close
                                    set rstSL = nothing
                                    set connSL = nothing
                                    set commandSL  = nothing
                                end if
                            end if
                        end if
						InsertarHistorialNserie mensajeTratEquipos,"","","MOVIMIENTOS ENTRE ALMACENES",p_nmovimiento,"","","","","MODIFY",mode
						if mode="first_save" then
							auditar_ins_bor session("usuario"),p_nmovimiento,"","alta","","","movimientos_almacenes"
						elseif mode="save" then
							if nz_b(rst("mercrecibida"))<>mercrecibida_old then
								if nz_b(rst("mercrecibida"))=-1 then
									auditar_ins_bor session("usuario"),p_nmovimiento,rst("almdestino"),"alta","","","mercancia_recibida"
								else
									auditar_ins_bor session("usuario"),p_nmovimiento,rst("almdestino"),"baja","","","mercancia_no_recibida"
								end if
							end if
						end if
					end if
					ant_mode=mode
					mode="browse"
				end if
				rst.close
			else
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgAlmOAlmDIguales2%>");
				      parent.botones.location = "movimientos_almacenes_bt.asp?mode=edit&bma=<%=enc.EncodeForJavascript(bma)%>"
				</script><%
				mode="edit"
			end if
		else
			if mode="first_save" then mode="add" else mode="edit"
			if no_proveedor=1 then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgResponsableNoExiste%>");
				</script><%
			elseif no_proveedor=2 then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgResponsableDadoBaja%>");
				</script><%
			end if
			%><script language="javascript" type="text/javascript">
			      parent.botones.location = "movimientos_almacenes_bt.asp?mode=<%=enc.EncodeForJavascript(mode)%>&bma=<%=enc.EncodeForJavascript(bma)%>"
			</script><%
		end if
	elseif mode="delete" then
		'comprobamos que no hay ningun detalle con nserie que no sea el ultimo documento
		mensajeTratEquipos=TratarEquipos("","","MOVIMIENTOS ENTRE ALMACENES",p_nmovimiento,"","","","","",mode)
		if mid(mensajeTratEquipos,1,2)<>"OK" then
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=mensajeTratEquipos%>");
			</script><%
			if submode>"" then
				mode=submode
			else
				mode="browse"
			end if
			%><script language="javascript" type="text/javascript">
			      document.movimientos_almacenes.mode.value = "<%=enc.EncodeForJavascript(mode)%>";
			</script><%
		else
			auditar_ins_bor session("usuario"),p_nmovimiento,"","baja","","","movimientos_almacenes"
			InsertarHistorialNserie mensajeTratEquipos,"","","MOVIMIENTOS ENTRE ALMACENES",p_nmovimiento,"","","","","",mode
			BorrarRegistro p_nmovimiento
			mode="add"
			p_nmovimiento="" %>
            <script language="javascript" type="text/javascript">
                //dgb: change to add, refresh search page and open it
                parent.botones.document.location = "Movimientos_Almacenes_bt.asp?mode=add";
                SearchPage("Movimientos_Almacenes_lsearch.asp?mode=init", 0);			    
		    </script>
        <%
		end if
  	end if

	if mode="traerresponsable" then
		submode2="traerresponsable"
		if tmp_responsable> "" then
			responsable=d_lookup("nombre","personal","dni='" & session("ncliente")&tmp_responsable & "'",session("dsn_cliente"))
			if responsable="" then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgResponsableNoExiste%>");
				</script><%
				tmp_responsable=tmp_responsable2
			else
                rstAux3.cursorlocation=3
				rstAux3.open "select dni,nombre,fbaja from personal with (NOLOCK) where dni='" & session("ncliente")&tmp_responsable & "' and fbaja is null",session("dsn_cliente")
				if rstAux3.eof then
					%><script language="javascript" type="text/javascript">
					      window.alert("<%=LitMsgResponsableDadoBaja%>");
					</script><%
					tmp_responsable=tmp_responsable2
				else
					tmp_responsable=session("ncliente")&tmp_responsable
					TmpNombre=responsable
				end if
				rstAux3.close
			end if
			mode=submode
			%><script language="javascript" type="text/javascript">
			      document.movimientos_almacenes.mode.value = "<%=enc.EncodeForJavascript(mode)%>";
			</script><%
		else
			tmp_responsable=""
			TmpNombre=""
			mode=submode
			%><script language="javascript" type="text/javascript">
			      document.movimientos_almacenes.mode.value = "<%=enc.EncodeForJavascript(mode)%>";
			</script><%
		end if
		if mode="first_save" then
			mode="add"
		else
			mode="edit"
		end if
	end if

	if mode="browse" or mode="search" or mode="add" then
		if mmt=1 then
			linea1=session("f_tpv")
			linea2=session("f_caja")
			linea3=session("f_empr")

			strSelect = "select c.almacen from tpv a with(nolock), cajas b with(nolock), tiendas c with(nolock) where a.caja=b.codigo and b.tienda=c.codigo and tpv='" & linea1 & "' and b.codigo='" & linea2 & "'"
			rstAux3.cursorlocation=3
			rstAux3.open strSelect,session("dsn_cliente")
			if not rstAux3.eof then
				AlmacenTienda=rstAux3("almacen")
			else
				AlmacenTienda=""
			end if
			rstAux3.close
			if AlmacenTienda & "">"" then
				cadena_mov_sol_usu=" and almdestino='" & AlmacenTienda & "'"
				tmp_almdestino=AlmacenTienda
			end if
            
		else
			cadena_mov_sol_usu=""
		end if

	end if

	'Mostrar los datos de la página.
	if mode="browse" or mode="edit" then
		if p_nmovimiento="" then
            rstAux.cursorlocation=3
			rstAux.open "select top 1 nmovimiento from movimientos with (NOLOCK) where nmovimiento like '" & session("ncliente") & "%' " & cadena_mov_sol_usu & " order by fecha desc,nmovimiento desc", session("dsn_cliente")
			if not rstAux.eof then p_nmovimiento=rstAux("nmovimiento")
			rstAux.close
		end if
        ''Ricardo 01-08-2013 comprobamos si el movimiento tiene articulos insertados con lote de compra
        existen_lotes_compra=0
        strselect="select d.idlote from detalles_movimientos as d with(NOLOCK) where d.nmovimiento like '" & session("ncliente") & "%' and d.nmovimiento='" & p_nmovimiento & "' and d.idlote is not null "
        rst.cursorlocation=3
		rst.Open strselect, session("dsn_cliente")
        if not rst.eof then
            existen_lotes_compra=1
        end if
        rst.close

		strselect="select m.*,convert(varchar,SAFTMOVEMENTSTARTTIME,103) + ' ' + convert(varchar,SAFTMOVEMENTSTARTTIME,108) as fecha_saft from movimientos as m with(NOLOCK) where nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & p_nmovimiento & "' " & cadena_mov_sol_usu
        rst.cursorlocation=3
		rst.Open strselect, session("dsn_cliente")
		if rst.eof then
			mode="add"
			rst.close
			%><script language="javascript" type="text/javascript">
			      parent.botones.document.location = "Movimientos_Almacenes_bt.asp?mode=add&bma=<%=enc.EncodeForJavascript(bma)%>";
			</script><%
            rst.cursorlocation=2
			rst.Open "select *,convert(varchar,SAFTMOVEMENTSTARTTIME,103) + ' ' + convert(varchar,SAFTMOVEMENTSTARTTIME,108) as fecha_saft from movimientos where nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & p_nmovimiento & "'", _
			session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rst.AddNew
		end if
	elseif mode="add" then
        rst.cursorlocation=2
		rst.Open "select *,convert(varchar,SAFTMOVEMENTSTARTTIME,103) + ' ' + convert(varchar,SAFTMOVEMENTSTARTTIME,108) as fecha_saft from movimientos where nmovimiento like '" & session("ncliente") & "%' and nmovimiento='" & p_nmovimiento & "'", _
		session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst.AddNew
    end if
	

	Alarma "movimientos_almacenes.asp"
	if mode="browse" or mode="save" then%>
	<%end if

	if (mode="browse" or mode="edit" or mode="add") then
	  if not rst.EOF then
		if nz_b(rst("mercrecibida"))=-1 and mode="edit" and submode2<>"traerresponsable" then
			texto_disabled=" disabled "
		else
            texto_disabled=""
		end if
        if texto_disabled & ""="" then
            if existen_lotes_compra=1 then
                texto_disabledAlm=" disabled "
            else
    			texto_disabledAlm=""
            end if
        end if
        
		%>
        <input type="hidden" name="h_nmovimiento" value="<%=enc.EncodeForHtmlAttribute(rst("nmovimiento"))%>"/>
        <input type="hidden" name="existen_lotes_compra" value="<%=enc.EncodeForHtmlAttribute(existen_lotes_compra)%>"/>
        <input type="hidden" name="FaltaGrabarLotesCompra" value="0"/>
        <%

        if mode="browse" or mode="save" then %>
           <div class="headers-wrapper"><%
             DrawDiv "header-date","",""
             DrawLabel "","",LitFecha
             DrawSpan "","",enc.EncodeForHtmlAttribute(iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha"))), ""
             CloseDiv

             DrawDiv "header-bill","",""
             DrawLabel "","",LitMovimiento
             DrawSpan "","",enc.EncodeForHtmlAttribute(trimCodEmpresa(rst("nmovimiento"))), ""
             CloseDiv%><%	
       'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION %>

                <%''ricardo 13-3-20003
				''si la serie tiene un formato de impresion sera este el de por defecto
				''si no sera el elegido en la tabla formatos impresion de ilion
                campo_bloqueado=0
                almacenSerie=""
                campo_bloqueadoAux=0
				if not rst.eof then
					defecto=obtener_formato_imp(rst("nserie"),"MOVIMIENTOS ENTRE ALMACENES")
                    campo_bloqueado=null_z(rst("bloqueado"))
                    campo_bloqueadoAux=campo_bloqueado
                    almacenSerie=d_lookup("almacen","series","nserie='" & rst("nserie") & "'",session("dsn_cliente"))
                    if (saft=-1 or saft=true) then
                        if (campo_bloqueado=-1 or campo_bloqueado=1 or campo_bloqueado=true) and almacenSerie & "">"" then
                            campo_bloqueado=0
                        else
                            campo_bloqueado=1
                        end if
                    end if
				end if
                %>
                <input type="hidden" name="h_almacenSerie" value="<%=enc.EncodeForHtmlAttribute(null_s(almacenSerie))%>" />
                <%
				''''''''
''response.write("los datos para la impresion son-" & cstr(rnf) & "-" & cstr(pnf) & "-" & saft & "-" & almacenSerie & "-" & campo_bloqueado & "-" & campo_bloqueadoAux & "-<br>")
                ver_impresion=""
                if ((cstr(rnf)="1" or cstr(pnf)="1") and (saft=-1 or saft=true) and (campo_bloqueado=-1 or campo_bloqueado=1 or campo_bloqueado=true)) and almacenSerie&""<>"" then ''or ((saft=-1 or saft=true) and campo_bloqueado=0)  then
                    ver_impresion="none"
                else
                    ver_impresion=""
                end if
				seleccion = "select b.fichero as fichero, a.descripcion,a.personalizacion,b.parametros as parametros from clientes_formatos_imp as a, formatos_imp as b where a.nformato=b.nformato and a.ncliente='"&session("ncliente")&"' and b.tippdoc='MOVIMIENTOS ENTRE ALMACENES' order by descripcion"
                rstSelect.cursorlocation=3
				rstSelect.Open seleccion, DsnIlion

                DrawDiv "col-md-3 col-sm-4 col-xxs-6 header-print","",""
                    %><label><a id="idPrintFormat" style=" display:<%=ver_impresion%>;" class='CELDAREFB' href="javascript:AbrirVentana(document.movimientos_almacenes.formato_impresion.value+'nmovimiento=<%="(\'"+enc.EncodeForJavascript(p_nmovimiento)+"\')"%>&mode=browse&empresa=<%=session("ncliente")%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" onmouseover="self.status='<%=LitImpresionConFormatoB%>'; return true;" onmouseout="self.status=''; return true;"><%=LitImpresionConFormatoB%></a></label>
				<select id="capa_impresion" class='CELDA' style='width:150px;display:<%=ver_impresion%>;' name="formato_impresion"><%
					encontrado=0
					while not rstSelect.eof
						if defecto=rstSelect("descripcion") then
							encontrado=1
							if isnull(rstSelect("parametros")) then
								prm=""
							else
								prm=rstSelect("parametros") & "&"
							end if%>
							<option selected="selected" value="<%=enc.EncodeForHtmlAttribute(rstSelect("fichero") & "?" & prm)%>"><%=enc.EncodeForHtmlAttribute(null_s(rstSelect("descripcion")))%></option>
						<%else
							if isnull(rstSelect("parametros")) then
								prm=""
							else
								prm=rstSelect("parametros") & "&"
							end if%>
							<option value="<%=enc.EncodeForHtmlAttribute(rstSelect("fichero")  & "?" & prm)%>"><%=enc.EncodeForHtmlAttribute(null_s(rstSelect("descripcion")))%></option>
						<%end if
						rstSelect.movenext
					wend%>
				</select>
			    <%
                rstSelect.close
                CloseDiv
                %></div><%
        else 
            'Encabezados en modo add o edit%>
            <div class="headers-wrapper"><%
                DrawDiv "header-date","",""
                DrawLabel "","",LitFecha
                if texto_disabled & "">"" then		
                      DrawInput "CELDA" & texto_disabled, "", "h_fecha", iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha")), ""
				    %><input type="hidden" name="fecha" value="<%=iif(mode="add",iif(p_fecha>"",enc.EncodeForHtmlAttribute(p_fecha),date()),enc.EncodeForHtmlAttribute(rst("fecha")))%>"/><%
			    else
                      DrawInput "CELDA", "", "fecha", enc.EncodeForHtmlAttribute(null_s(iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha")))), ""
                      DrawCalendar "fecha"				

                end if
                CloseDiv
                DrawDiv "header-bill","",""
                DrawLabel "","",LitMovimiento
                DrawSpan "","",trimCodEmpresa(rst("nmovimiento")), ""
                CloseDiv%></div><%
        end if

        'Inicio Borde Span%>
        <table width="100%"><tr><td>

        <% 'Para desplegar o comprimir todo %>
        <!--
        <div id="CollapseSection">
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['CABECERA', 'DETALLES', 'AddDG', 'BrowseCab', 'BrowseDG']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['CABECERA', 'DETALLES', 'AddDG', 'BrowseCab', 'BrowseDG']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
        </div>
        -->
        <%if mode="browse" then %>
        <div   class="Section" id="S_BrowseDG">
            <a href="#" rel="toggle[BrowseDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader">
                <%=litcabecera%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />   
            </div></a> 

            <div class="SectionPanel" style="display:none;" id="BrowseDG">
           <%			
				if mode="browse" then					
                    DrawDiv "1","",""
                    DrawLabel "","",LitAlmacenDestino
					if rst("almdestino")&"">"" then
						%><span class="CELDA"><%nom_almacen=d_lookup("descripcion","almacenes","codigo='" & rst("almdestino") & "'",session("dsn_cliente"))
                            response.write(nom_almacen)
                            %>
                            <input type="hidden" name="nom_almacen" value="<%=enc.EncodeForHtmlAttribute(null_s(nom_almacen))%>" />
							<input class="CELDA" type="hidden" name="almdestino" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("almdestino")))%>" size="10" />
						</span><%
					else
						DrawSpan "CELDA","","",""
					end if
                    CloseDiv

                    EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitSerie,"", trimCodEmpresa(rst("nserie"))
			    end if
			
            ''''''''''''''''''SAFT
			if mode<>"add" and mode <>"edit" then
                    campo_saft=0
                    if saft=True then
                        campo_saft=1
                    end if
                    if saft=True then
                        cadena_bloqueado=""
                        if rst("bloqueado")=0 then
                            cadena_bloqueado1=""
                            cadena_bloqueado2="none"
                        else
                            cadena_bloqueado1="none"
                            cadena_bloqueado2=""
                        end if
                    else
                        cadena_bloqueado1="none"
                        cadena_bloqueado2="none"
                    end if
                    
			        if saft=True then
                        %><div id="textobloqueado" class="col-lg-4 col-md-6 col-sm-6 col-xs-12" style="display:<%=cadena_bloqueado2%>;"><%DrawLabel "", "",LitBloqueada%></div>
                        <%
			        end if
			        if saft=True then
                        %><div id="textodebloqueado" class="col-lg-4 col-md-6 col-sm-6 col-xs-12" style="display:<%=cadena_bloqueado%>;"><%DrawLabel "", "",LitLibre%></div>
                        <%
			        end if
                    %>
                    
                    <%if mode="add" then %>
                        <input type="hidden" name="bloqueado" value="0"/>
                    <%else %>
                        <input type="hidden" name="bloqueado" value="<%=iif(rst("bloqueado")<>0,"1","0")%>"/>
                    <%end if %>
                    <%if mode<>"add" then

                        %>
                        <%if (pnf="1") then 'permitido bloquear, factura libre%>
                            <div id="celdaBloquear" class="col-lg-4 col-md-6 col-sm-6 col-xs-12" style="display:<%=cadena_bloqueado1%>;"><a href="javascript:bloqueoMovimiento('<%=enc.EncodeForJavascript(rst("nmovimiento"))%>','<%=enc.EncodeForJavascript(rst("nserie"))%>','<%=campo_saft%>','0')" onmouseover="self.status='<%=LITBLOQUEOMOV%>'; return true;" onmouseout="self.status=''; return true;"><img src="../images/<%=ImgBloqueoFacturas%>" <%=ParamImgBloqueoFacturas%> alt="<%=LITBLOQUEOMOV%>" title="<%=LITBLOQUEOMOV%>"/></a></div>
				        <%end if
				        if (rnf="1") then 'permitido desbloquear, factura bloqueada %>
                            <div id="celdaDesBloquear" class="col-lg-4 col-md-6 col-sm-6 col-xs-12" style="display:<%=cadena_bloqueado2%>;"><a href="javascript:desbloqueoMovimiento('<%=enc.EncodeForJavascript(rst("nmovimiento"))%>','<%=enc.EncodeForJavascript(rst("nserie"))%>','<%=campo_saft%>','0')" onmouseover="self.status='<%=LITDESBLOQUEOMOV%>'; return true;" onmouseout="self.status=''; return true;"><img src="../images/<%=ImgValidar%>" <%=ParamImgBloqueoFacturas%> alt="<%=LITDESBLOQUEOMOV%>" title="<%=LITDESBLOQUEOMOV%>"/></a></div>
				        <%end if%>
			        <%
                    else

                    end if
            end if
            ''''''''''''''''''FIN SAFT
			if mode<>"add" then

				if cstr(bmr & "")="1" then
					texto_bmr=" " & "disabled"
				else
					texto_bmr=""
				end if
				
					%><input type="hidden" name="h_mercrecibida" value="<%=enc.EncodeForHtmlAttribute(nz_b(rst("mercrecibida")))%>"/><%
						
						EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitMercRecMov,"", iif(nz_b(rst("mercrecibida"))=-1,"Sí","No")                       
                        EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitFechaRecepcion,"", rst("fecharecepcion")
				                      
                if saft=True then
                   EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitSaftAT,"", enc.EncodeForHtmlAttribute(null_s(rst("SAFTATDOCCODEID")))
                   EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitSaftFechSal,"", enc.EncodeForHtmlAttribute(null_s(rst("fecha_saft")))
                   
                    campo_bloqueado=nz_b(rst("bloqueado"))
                    ver_links=""
                    if (campo_bloqueado=-1 or campo_bloqueado=1 or campo_bloqueado=true) then
                        ver_links=""
                    else
                        ver_links="none"
                    end if
                    %>
                        <div id="linksaft" class="col-lg-12 col-md-12 col-sm-12 col-xs-12" style="display:<%=ver_links%>;">
                            <!--<td class="CELDA" colspan="2" ><a class='CELDAREFB' href="javascript:MarcarMecrecibida('<%=rst("nmovimiento")%>')">Mercancia Recibida</a></td>-->
                            <a href="javascript:MarcarMecrecibida('<%=enc.EncodeForJavascript(rst("nmovimiento"))%>')">Mercanc&iacute;a Recibida</a>
                            <table class="width100"></table>
                            <%
				            
                            %>
                           
                            <a href="javascript:IntroducirCodeAt('<%=enc.EncodeForJavascript(rst("nmovimiento"))%>')">Introducir C&oacute;digo At</a>
                        </div>
                                              
                    <%
                end if
		    end if
			%>
		<!--</table>-->
        <!--<table width='100%' border="0" cellspacing="1" cellpadding="1">-->
        <%
                    'DrawFila color_titulo
                        %>
                            <!--<td colspan="5">
                                <font class = "ENCABEZADOC"><%=LITDATOSGENERALES%></font>
                            </td>--><%
                    'CloseFila
                     DrawDiv "3-sub", "background-color: #eae7e3", ""
                    %> 
                    <label class="ENCABEZADOC" style="text-align:left"><%=LITDATOSGENERALES%></label>
                    <%
                    CloseDiv
        

				'ponemos el responsable por defecto
				if tmp_responsable="" then
					'buscar el usuario en la tabla personal
					rstAux.cursorlocation=3
					rstAux.open "select dni from personal with (NOLOCK) where login='" & session("usuario") & "' and dni like '" & session("ncliente") & "%'",session("dsn_cliente")
					if not rstAux.eof then
						tmp_responsable=rstAux("dni")
					end if
					rstAux.close
				else
					if mid(tmp_responsable,1,5)<>session("ncliente") then
						tmp_responsable=session("ncliente") & tmp_responsable
					end if
				end if

				if rst("responsable")&"">"" then
					%><input class='CELDA' type="hidden" name="responsable" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("responsable")))%>" size="10" /><%
                    EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitResponsable,"", enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre","personal","dni='" & rst("responsable") & "'",session("dsn_cliente"))))
				else
					%><!--<td class=dato>&nbsp;</td>--><%
				end if				
				if session("version")&"" <> "5" then
                    DrawDiv "","","" 
                    CloseDiv
                end if 
				EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitObservaciones,"", enc.EncodeForHtmlAttribute(pintar_saltos_espacios(iif(tmp_observaciones>"",tmp_observaciones,null_s(rst("observaciones")))))
			%></div>   
        </div>

        <% 'Detalles	%>         
        <div class="Section" id="S_DETALLES">
            <a href="#" rel="toggle[DETALLES]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader">
                <%=LitTituloDet%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
            </div></a>

            <div class="SectionPanel" style="display: " id="DETALLES">
            <br />
	         <%
                    campo_bloqueado=nz_b(rst("bloqueado"))
''response.write("los datos para los detalles son-" & cstr(rnf) & "-" & cstr(pnf) & "-" & saft & "-" & almacenSerie & "-" & campo_bloqueado & "-" & campo_bloqueadoAux & "-<br>")
                    if nz_b(rst("mercrecibida"))=0 then
                        ver_detins=""
                        if (cstr(rnf)="1" or cstr(pnf)="1") and (saft=-1 or saft=true) and (campo_bloqueado=-1 or campo_bloqueado=1 or campo_bloqueado=true) then
                            ver_detins="none"
                        else
                            ver_detins=""
                        end if
				   	    %><table id="frDetallesInsT" class="width90 md-table-responsive bCollapse" style="display:<%=ver_detins%>;">
                               <tr>
                                    <td class='ENCABEZADOL underOrange width5'><%=LitItem%></td>
							        <td class='ENCABEZADOL underOrange width5'><%=LitCantidad%></td>
							        <%'dgb 11/04/2008  se muestran los datos Enviado/Recibidos 
							        if si_tiene_modulo_Centroxogo<>0 then  %>
							            <td class='ENCABEZADOL underOrange width5'><%=LitER%></td>
							        <%end if %>
							        <td class='ENCABEZADOL underOrange width10'><%=LitReferencia%></td>
							        <td class='ENCABEZADOL underOrange width15'><%=LitDescripcion%></td>
							        <td class='ENCABEZADOL underOrange width10'><%=LitAlmacenOrigen%></td>
							        <%if si_tiene_modulo_produccion<>0 then%>
								        <td class='ENCABEZADOL underOrange width10'><%=LitLote%></td>
							        <%end if%>
							        <td class='ENCABEZADOL underOrange width5'>&nbsp</td>
							        <td class='ENCABEZADOL underOrange width5'>&nbsp</td>
                                </tr>
				   	      </table>
                            <iframe style="display:<%=ver_detins%>;" id='frDetallesIns' name="fr_DetallesIns" src='movimientos_almacenes_detins.asp?ndoc=<%=enc.EncodeForJavascript(rst("nmovimiento"))%>' class="width90 md-table-responsive" height='70' frameborder="no" scrolling="no" noresize="noresize"></iframe><br /><%
				    end if
				    'dgb 11/04/2008  se muestran los datos Enviado/Recibidos 
				    if si_tiene_modulo_Centroxogo<>0 then 
				        if si_tiene_modulo_produccion<>0 then%>
					       <iframe id='frDetalles' name="fr_Detalles" src='movimientos_almacenes_det.asp?ndoc=<%=enc.EncodeForJavascript(rst("nmovimiento"))%>&EstadoCierre=<%=enc.EncodeForJavascript(blnEstadoCierre)%>' class="width90 md-table-responsive" height='150' frameborder="yes" noresize="noresize"></iframe>
				        <%else%>
                            <iframe id='frDetalles' name="fr_Detalles" src='movimientos_almacenes_det.asp?ndoc=<%=enc.EncodeForJavascript(rst("nmovimiento"))%>&EstadoCierre=<%=enc.EncodeForJavascript(blnEstadoCierre)%>' class="width90 md-table-responsive" height='150' frameborder="yes" noresize="noresize"></iframe>
				        <%end if
				    else
				        if si_tiene_modulo_produccion<>0 then%>
                            <iframe id='frDetalles' name="fr_Detalles" src='movimientos_almacenes_det.asp?ndoc=<%=enc.EncodeForJavascript(rst("nmovimiento"))%>&EstadoCierre=<%=enc.EncodeForJavascript(blnEstadoCierre)%>' class="width90 md-table-responsive" height='150' frameborder="yes" noresize="noresize"></iframe>
				        <%else%>
                            <iframe id='frDetalles' name="fr_Detalles" src='movimientos_almacenes_det.asp?ndoc=<%=enc.EncodeForJavascript(rst("nmovimiento"))%>&EstadoCierre=<%=enc.EncodeForJavascript(blnEstadoCierre)%>' class="width90 md-table-responsive" height='150' frameborder="yes" noresize="noresize"></iframe>
				        <%end if%>
				    <%end if %>
            </div>   
         </div>

		<%
        else 
             'Datos Generales %>
            <div   class="Section" id="S_AddDG">
                <a href="#" rel="toggle[AddDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=litcabecera%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />  
                </div></a>  
                <div class="SectionPanel" style="display: " id="AddDG">
                 <!--<table width="70%" border='0' cellspacing="1" cellpadding="1">--><%			    
				    'DrawCelda "ENCABEZADOL","","",0,LitAlmacenDestino+":"
				    campo="codigo"
				    campo2="descripcion"
				    dato_celda=Desplegable(mode,campo,campo2,"almacenes",enc.EncodeForHtmlAttribute(iif(tmp_almdestino>"",tmp_almdestino,null_s(rst("almdestino")))),"")

				    if texto_disabled & "">"" or texto_disabledAlm & "">"" then
					    EligeCelda "select-disabled", mode,"CELDA" & iif(Repes>""," disabled",iif(texto_disabled>"",texto_disabled,texto_disabledAlm)),iif(mode<>"browse","180",""),"",0,LitAlmacenDestino,"h_almdestino2",15,dato_celda
					    %><input type="hidden" name="almdestino" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_almdestino>"",tmp_almdestino,rst("almdestino"))))%>"/><%
				    else
					    EligeCelda "select", mode,"CELDA" & iif(Repes>""," disabled",""),iif(mode<>"browse","180",""),"",0,LitAlmacenDestino,"almdestino",15,dato_celda
				    end if
				    if mode="add" or mode="edit" then RstAux.close
				    %><input type="hidden" name="h_almdestino" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_almdestino>"",tmp_almdestino,rst("almdestino"))))%>"/><%
				    
				    rstAux.cursorlocation=3
				    strSacSerie="select nserie, nombre as descripcion from series with (NOLOCK) where tipo_documento ='MOVIMIENTOS ENTRE ALMACENES' and nserie like '" & session("ncliente") & "%'"
					if s & "">"" then
						  strSacSerie=strSacSerie & " and nserie in " & s
					end if
                    strSacSerie=strSacSerie & " order by nombre"
                    rstAux.open strSacSerie,session("dsn_cliente")
                    if mode="add" then
					    DrawSelectCelda "CELDA","185","",0,LitSerie,"nserie",rstAux,enc.EncodeForHtmlAttribute(null_s(iif(p_serie>"",p_serie,rst("nserie")))),"nserie","descripcion","",""
				    else
					    if texto_disabled & "">"" then
						    DrawSelectCeldaDisabled "","","",0,LitSerie,"h_nserie",rstAux,enc.EncodeForHtmlAttribute(null_s(iif(p_serie>"",p_serie,rst("nserie")))),"nserie","descripcion","",""
						    %><input type="hidden" name="nserie" value="<%=enc.EncodeForHtmlAttribute(iif(p_serie>"",p_serie,rst("nserie")))%>"/><%
					    else
						    DrawSelectCelda "CELDA","185","",0,LitSerie,"nserie",rstAux,enc.EncodeForHtmlAttribute(null_s(iif(p_serie>"",p_serie,rst("nserie")))),"nserie","descripcion","",""
					    end if
				    end if
			 	    rstAux.close
			    
			    if mode<>"add" then
				    if cstr(bmr & "")="1" then
					    texto_bmr=" " & "disabled"
				    else
					    texto_bmr=""
				    end if
				    
					    %><input type="hidden" name="h_mercrecibida" value="<%=enc.EncodeForHtmlAttribute(nz_b(rst("mercrecibida")))%>"/><%
					    if mode="browse" then						   
						   
                            EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitMercRecMov,"", iif(nz_b(rst("mercrecibida"))=-1,"Sí","No")   
                          
                            EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitFechaRecepcion,"", enc.EncodeForHtmlAttribute(null_s(rst("fecharecepcion")))
					    else						    
						    if texto_bmr=" " & "disabled" then
							    EligeCelda "check", mode,"' " & texto_bmr,"0","",0,LitMercRecMov,"h_merric",0,enc.EncodeForHtmlAttribute(iif(merric>"",nz_b(merric),nz_b(rst("mercrecibida"))))
							    %><input type="hidden" name="merric" value="<%=iif(merric>"",nz_b(merric),nz_b(rst("mercrecibida")))%>"/><%
						    else
							    EligeCelda "check", mode,"' " & texto_bmr,"0","",0,LitMercRecMov,"merric",0,enc.EncodeForHtmlAttribute(iif(merric>"",nz_b(merric),nz_b(rst("mercrecibida"))))
						    end if
                            
					    end if
				    
                                  
                    if saft=True then                       
                            DrawDiv "1","",""                           
                            DrawLabel "","",LitSaftAT
                            ''DrawCelda "CELDA","","",0,rst("SAFTATDOCCODEID")
							%><input class='CELDA' type="text" name="SAFTATDOCCODEID" value="<%=enc.EncodeForHtmlAttribute(rst("SAFTATDOCCODEID"))%>" size="30" maxlength="200" />
							<%CloseDiv
                            
                            DrawDiv "1","",""
                            DrawLabel "","",LitSaftFechSal
							%><input class='CELDA' type="text" name="SAFTMOVEMENTSTARTTIME" value="<%=enc.EncodeForHtmlAttribute(rst("fecha_saft"))%>" size="20" maxlength="20" /><%                     
                            DrawCalendar "SAFTMOVEMENTSTARTTIME"
                            CloseDiv                       
                    end if
		    end if		   
                     
                  DrawDiv "3-sub", "background-color: #eae7e3", ""
                  %> 
                  <label class="ENCABEZADOC" style="text-align:left"><%=LITDATOSGENERALES%></label>
                  <%
                  CloseDiv            

				    'ponemos el responsable por defecto
				    if tmp_responsable="" then
					    'buscar el usuario en la tabla personal
					    rstAux.cursorlocation=3
					    rstAux.open "select dni from personal with (NOLOCK) where login='" & session("usuario") & "' and dni like '" & session("ncliente") & "%'",session("dsn_cliente")
					    if not rstAux.eof then
						    tmp_responsable=rstAux("dni")
					    end if
					    rstAux.close
				    else
					    if mid(tmp_responsable,1,5)<>session("ncliente") then
						    tmp_responsable=session("ncliente") & tmp_responsable
					    end if
				    end if

					if mode="browse" then						  
					    if rst("responsable")&"">"" then
						    %> <input class='CELDA' type="hidden" name="responsable" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("responsable")))%>" size="10" /><%                                
                             EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitResponsable,"", enc.EncodeForHtmlAttribute(null_s(d_lookup("nombre","personal","dni='" & rst("responsable") & "'",session("dsn_cliente"))))
					    else
						    %><!--<td class="dato">&nbsp;</td>--><%
					    end if
					else
					    if texto_disabled & "">"" then
                                DrawDiv "1","",""
						        DrawLabel "","",LitResponsable
							    %><input class='width15' type="text" <%=texto_disabled%> name="h_responsable" size="10" value="<%=enc.EncodeForHtmlAttribute(iif(tmp_responsable& "">"",trimCodEmpresa(tmp_responsable),trimCodEmpresa(rst("responsable"))))%>"/>
                                <input type="hidden" name="responsable" value="<%=enc.EncodeForHtmlAttribute(iif(tmp_responsable& "">"",trimCodEmpresa(tmp_responsable),trimCodEmpresa(rst("responsable"))))%>"/>
							    <a class="CELDAREFB" href="javascript:" onmouseover="self.status='<%=LitVerPersonal%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
                                <input class="width40" disabled <%=texto_disabled%> type="text" name="h_nomresponsable" size="48" value="<%=enc.EncodeForHtmlAttribute(iif(TmpNombre & "">"",TmpNombre,d_lookup("nombre","personal","dni='" & iif(tmp_responsable & "">"",tmp_responsable,rst("responsable")) & "'",session("dsn_cliente"))))%>"/>
							    <input type="hidden" name="nomresponsable" value="<%=enc.EncodeForHtmlAttribute(iif(TmpNombre & "">"",TmpNombre,d_lookup("nombre","personal","dni='" & iif(tmp_responsable & "">"",tmp_responsable,rst("responsable")) & "'",session("dsn_cliente"))))%>"/>
						    <%CloseDiv
					    else
						    DrawDiv "1","",""
                            DrawLabel "","",LitResponsable
							    %><input class='width15' type="text" <%=texto_disabled%> name="responsable" size=10 value="<%=enc.EncodeForHtmlAttribute(iif(tmp_responsable& "">"",trimCodEmpresa(tmp_responsable),trimCodEmpresa(rst("responsable"))))%>" onchange="TraerResponsable();"/>
                                <a class="CELDAREFB" href="javascript:AbrirVentana('../administracion/personal_buscar.asp?viene=movimientos_almacenes&titulo=<%=LitSelPersonal%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" onmouseover="self.status='<%=LitVerPersonal%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
                                <input class="width40" disabled <%=texto_disabled%> type="text" name="nomresponsable" size="48" value="<%=enc.EncodeForHtmlAttribute(iif(TmpNombre & "">"",TmpNombre,d_lookup("nombre","personal","dni='" & iif(tmp_responsable & "">"",tmp_responsable,rst("responsable")) & "'",session("dsn_cliente"))))%>"/>
						    <%CloseDiv
					    end if
					end if
                
				    if mode="browse" then
                        EligeCeldaResponsive "text","browse","CELDA","","",0,"",LitObservaciones,"", pintar_saltos_espacios(iif(tmp_observaciones>"",tmp_observaciones,rst("observaciones")&""))
				    else
                        if session("version")&"" <> "5" then
                            DrawDiv "","","" 
                            CloseDiv
                        end if 
						if texto_disabled & "">"" then
                            DrawDiv "1","",""
                            DrawLabel "","",LitObservaciones%><textarea disabled="disabled" class="width60" name="<%=h_observaciones%>" ><%=enc.EncodeForHtmlAttribute(iif(tmp_observaciones>"",tmp_observaciones,iif(rst("observaciones")>"",rst("observaciones"),"")))%></textarea><%
                            CloseDiv
							'EligeCelda "text","add","' " & texto_disabled,"","",0,LitObservaciones,"h_observaciones","",iif(tmp_observaciones>"",tmp_observaciones,iif(rst("observaciones")>"",rst("observaciones"),""))
                        %><input type="hidden" name="observaciones" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(tmp_observaciones>"",tmp_observaciones,iif(rst("observaciones")>"",rst("observaciones"),""))))%>" /><%
						else
							EligeCelda "text","add","CELDA","","",0,LitObservaciones,"observaciones","",enc.EncodeForHtmlAttribute(null_s(iif(tmp_observaciones>"",tmp_observaciones,iif(rst("observaciones")>"",rst("observaciones"),""))))
						end if						    
				    end if%></div>   
            </div>

        <%end if %>
		</td></tr></table>

		<span id="paginacion" style="display: ">
		</span>
	    <%if submode2="traerresponsable" then
			if texto_disabled & "">"" and texto_bmr<>" " & "disabled" then%>
				<script language="javascript" type="text/javascript">
				    document.movimientos_almacenes.merric.focus();
				</script>
			<%else%>
				<script language="javascript" type="text/javascript">
				    document.movimientos_almacenes.observaciones.focus();
				    document.movimientos_almacenes.observaciones.select();
				</script>
			<%end if
		elseif mode="add" then%>
			<script language="javascript" type="text/javascript">
			    document.movimientos_almacenes.fecha.focus();
			    document.movimientos_almacenes.fecha.select();
			</script>
		<%elseif mode="edit" then
			if texto_disabled & "">"" and texto_bmr<>" " & "disabled" then%>
				<script language="javascript" type="text/javascript">
				    document.movimientos_almacenes.merric.focus();
				</script>
			<%else%>
				<script language="javascript" type="text/javascript">
				    document.movimientos_almacenes.fecha.focus();
				</script>
			<%end if
		end if
        if mode="browse" then %>
		    <script language="javascript" type="text/javascript">		        Redimensionar();</script>
        <%end if
	  end if
	end if %>
</form>
<%
	set rstAux = nothing
	set rstAux3 = nothing
	set rst = nothing
	set rstSelect = nothing
 end if %>
</body>
</html>
