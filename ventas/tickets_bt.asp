<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>">
</head>

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  

<!--#include file="../calculos.inc" -->
<!--#include file="../constantes.inc" -->

<!--#include file="../cache.inc" -->
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->

<!--#include file="tickets.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="../styles/Master.css.inc" -->

<!--#include file="../varios.inc" -->



<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">

        window.addEventListener("load", function hideFilterMaster() {
        var classNameLeft = parent.parent.document.getElementById('left').className;
        if (classNameLeft == "search-panel open") {
            document.getElementById("FILTERS_MASTER_ASP").setAttribute("style", "display:none")
        }
    });

    function comprobar_enter() {
        //si se ha pulsado la tecla enter
        //if (window.event.keyCode==13){
        //document.opciones.criterio.focus();
        Buscar();
        //}
    }

    function Buscar() {
	    SearchPage("tickets_lsearch.asp?mode=search&campo=" + document.opciones.campos.value +
	    "&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value, 1);
	    document.opciones.texto.value = "";
    }

    // Realizar la acción correspondiente al botón pulsado.
    function Accion(pulsado) {
        switch (pulsado) {
            case "Anular": // Anular ticket.
                if (confirm('¿Desea anular el Ticket?')) {
                    parent.pantalla.document.location = "tickets.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=anular";
                    document.location = "tickets_bt.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=browse";
                }
                break;
            
            case "Editar": //Editar cabecera ticket
                parent.pantalla.document.location = "tickets.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=editFuga";
                document.location = "tickets_bt.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=edit";
                break;
				
			case "Cancelar": //Cancelar edicion cabecera
                parent.pantalla.document.location = "tickets.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=browse";
                document.location = "tickets_bt.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=browse";
                break;
				
			case "Guardar": //Guardar edicion cabecera

                parent.pantalla.document.location = "tickets.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=saveFuga&matricula="+parent.pantalla.tickets.matricula.value+"&observacion="+parent.pantalla.tickets.observacion.value;
                document.location = "tickets_bt.asp?nticket=" + parent.pantalla.tickets.h_nticket.value + "&mode=browse";
                break;
                
        }
    }
        
</script>

<body class="body_master_ASP">

<%mode=limpiaCadena(Request.QueryString("mode"))%>
<form name="opciones" method="post" action="">
<input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
<div id="PageFooter_ASP" >
<div id="ControlPanelFooter_left_ASP" >
    <table id="BUTTONS_CENTER_ASP">
		    <tr>
                <%
                ' SSR 02/01/17 Recogida de parametros para ocultar el boton Anular.
                dim opAnular
                opAnular = "0"
                                    
                bit_cerrar=d_lookup("CERRARDOC","CONFIGURACION"," NEMPRESA = '" & session("ncliente") & "'",session("dsn_cliente"))
                if ucase(bit_cerrar)="ON" or ucase(bit_cerrar)="VERDADERO" or ucase(bit_cerrar)="TRUE" or ucase(bit_cerrar)="1" then
                    opAnular = "1"
               

                    'Comprobamos si el ticket esta anulado
                    set conn = Server.CreateObject("ADODB.Connection")
                    set command =  Server.CreateObject("ADODB.Command")
                    conn.open session("dsn_cliente")
                    command.ActiveConnection = conn
                    command.CommandTimeout = 0
                    command.CommandText="select NTICKET from TICKETS with(nolock) where TOTAL_TICKET='0.00' and NTICKET ='" + limpiaCadena(Request.QueryString("nticket")) + "'"

                    set result=command.Execute

                    if not result.EOF then
                        opAnular = "0"
                    end if

                    conn.close
                    set command = nothing
                    set conn = nothing

                    'Mostramos los botones
                    if (opAnular="1") And (mode="browse") then%>
                        <td id="idanular" class="CELDABOT" onclick="javascript:Accion('Anular');">
                            <%PintarBotonBTLeftRed "ANULAR",ImgBorrar,ParamImgBorrar,"ANULARTITLE"%>
                        </td>
			    <%  end if
                 end if

				'Comprobamos si el ticket es de tipo FUGA
				dim esFuga
                esFuga = "0"

				set c = Server.CreateObject("ADODB.Connection")
                set com =  Server.CreateObject("ADODB.Command")
                c.open session("dsn_cliente")
                com.ActiveConnection = c
                com.CommandTimeout = 0
				
				nticket=limpiaCadena(Request.QueryString("nticket"))
				if nticket="" then nticket=limpiaCadena(Request.form("nticket"))
				if nticket="" then nticket=limpiaCadena(Request.QueryString("ndoc"))
				checkCadena nticket
				
                com.CommandText="select tip.descripcion from tickets as t with(NOLOCK) left outer join tipo_pago as tip with(nolock) on tip.codigo=t.medio_pago where t.NTICKET ='" + nticket + "' and tip.descripcion='FUGA'"

                set r=com.Execute
				
                if not r.EOF then
                    esFuga = "1"
                end if

				c.close
                set com = nothing
                set c = nothing

				if (esFuga="1") And (mode="browse") then%>
				
                    <td id="ideditar" class="CELDABOT" onclick="javascript:Accion('Editar');">
                        <%PintarBotonBTLeft LITBTNEDITARCABECERA,ImgEditar,ParamImgEditar,"EDITARTITLE"%>
                    </td>
				<%end if%>
				<%
				if (mode="edit") then%>
				
                    <td id="idguardar" class="CELDABOT" onclick="javascript:Accion('Guardar');">
                        <%PintarBotonBTLeft LITBTNGUARDAR,ImgGuardar,ParamImgGuardar,"GUARDARTITLE"%>
                    </td>
				<%end if
				
				if (mode="edit") then%>
				
                    <td id="idcancelar" class="CELDABOT" onclick="javascript:Accion('Cancelar');">
                        <%PintarBotonBTLeftRed LITBTNCANCELAR,ImgCancelar,ParamImgCancelar,"CANCELARTITLE"%>
                    </td>
				<%end if%>

            </tr>
        </table>
    </div>
    
    <!--<p>Prueba de modo:</p>
    <input type="text" name="mode2" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>-->

    <div id="FILTERS_MASTER_ASP">
		<select class="IN_S" name="campos">
				<option selected value="t.nventa"><%=LitTicket%></option>
				<option value="p.nombre"><%=LitOperador%></option>
		</select>
		<select class="IN_S" name="criterio">
				<option value="contiene"><%=LitContiene%></option>
				<option value="termina"><%=LitTermina%></option>
				<option value="igual"><%=LitIgual%></option>
		</select>
		<input id="KeySearch" class="IN_S" type="text" name="texto" size="20" maxlength="20" value="" runat="javascript:comprobar_enter();"/>
		<a class="CELDAREF" href="javascript:Buscar();"><img src="<%=ImgBuscarLF_bt%>" <%=ParamImgBuscarLF_bt%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
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
