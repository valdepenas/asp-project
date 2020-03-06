<%@ Language=VBScript %>
<%
'CODIGOS DE AÑADIDURAS/MODIFICACIONES -----------------------------------------------------
'JCI-14022003-01 : Modificaciones diversas para gestionar al personal como operario de
'                  fabricación
'      FECHA     : 14/02/2003
'      AUTOR     : JCI
'VGR 7/3/03 	 : Añadir BarraOpciones en Datos de Comerciales.
'VGR 28/03/03 	 : Quitar espacios en blanco de los campos hcomercial,htecnico,hoperario
'IVM 02/04/03 	 : Se añade gestión de penalizaciones
'IVM 04/04/03 	 : Se añade pestaña para ver listado de ventas TPV
'IVM 04/04/03 	 : Se añade pestaña para ver incentivos
'IVM 11/04/03    : Se añade el campo Código de Operador
'IVM 11/04/03    : Se modifica la página para que en el caso de que al guardar un registro no
'                  se pueda guardar por algún motivo se vuelva al modo add o browse con los
'                  datos anteriores
'JCI 22/04/03    : Añadir opción de ver los tickets anulados de talonarios
'------------------------------------------------------------------------------------------%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->

<!--#include file="personal.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/tabs.js.inc" -->
<!--#include file="../js/calendar.inc"-->

<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
    
<!--#include file="personal_linkextra.inc"-->
<!--#include file="../styles/dropdown.css.inc" -->
<!--#include file="../js/dropdown.js.inc" -->

<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('addDG', 'fade=1')
    animatedcollapse.addDiv('browseDG', 'fade=1')
    animatedcollapse.addDiv('browseDC', 'fade=1')
    animatedcollapse.addDiv('browseDT', 'fade=1')
    animatedcollapse.addDiv('browseDO', 'fade=1')
    animatedcollapse.addDiv('editDG', 'fade=1')
    animatedcollapse.addDiv('editDC', 'fade=1')
    animatedcollapse.addDiv('editDT', 'fade=1')
    animatedcollapse.addDiv('editDO', 'fade=1')

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()
</script>

<script language="javascript" type="text/javascript">
    //Abre la ventana de comisiones por articulos que tiene un comercial
    function Comisiones(dni,nombre)
    {
        Ven=AbrirVentana("../administracion/comisiones_articulos.asp?mode=edit&ndoc=personal&dni=" + dni + "&nombre=" + nombre,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function Penalizaciones(dni,nombre)
    {
        Ven=AbrirVentana("../administracion/personal_pen.asp?mode=edit&ndoc=personal&dni=" + dni + "&nombre=" + nombre,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }


    function Incentivos(dni,nombre)
    {
        Ven=AbrirVentana("../administracion/personal_incen.asp?mode=edit&ndoc=personal&dni=" + dni + "&nombre=" + nombre,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function Anulaciones(dni,nombre)
    {
        Ven=AbrirVentana('../central.asp?pag1=ventas/Lista_Anulaciones.asp&pag2=ventas/Lista_Anulaciones_bt.asp&ndoc=' + dni + '&viene=personal&titulo=<%=LitAnulaciones%>','P',<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function ComisionesExtras(dni,nombre)
    {
        Ven=AbrirVentana("../central.asp?pag1=administracion/comisiones_extras.asp&pag2=administracion/comisiones_extras_bt.asp&ndoc=personal&dni=" + dni + "&nombre=" + nombre,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function Comercial(modo,si_tiene_modulo_comercial){
        if (modo=="alta")
        {
            if(si_tiene_modulo_comercial!=0) texto_confirm="<%=LitMsgAltaComercialConfirmModCom%>";
            else texto_confirm="<%=LitMsgAltaComercialConfirmSinModCom%>";
            if (window.confirm(texto_confirm)==true)
                document.location="personal.asp?dni="+document.personal.hdni.value+"&domicilio="+document.personal.hdomicilio.value+"&comercial=alta";
        }
        else
        {
            if(si_tiene_modulo_comercial!=0) texto_confirm="<%=LitMsgEliminarComercialConfirmModCom%>";
            else texto_confirm="<%=LitMsgEliminarComercialConfirm%>";
            if (window.confirm(texto_confirm)==true) 
                document.location="personal.asp?dni="+document.personal.hdni.value+"&domicilio="+document.personal.hdomicilio.value+"&comercial=baja";
        }
    }

    function Tecnico(modo)
    {
        if (modo=="alta")
        {
            if (window.confirm("<%=LitMsgAltaTecnicoConfirm%>")==true) 
                document.location="personal.asp?dni=" + document.personal.hdni.value + "&domicilio="+document.personal.hdomicilio.value+"&tecnico=alta";
        }
        else
        {
            if (window.confirm("<%=LitMsgBajaTecnicoConfirm%>")==true) 
                document.location="personal.asp?dni=" + document.personal.hdni.value + "&domicilio="+document.personal.hdomicilio.value+"&tecnico=baja";
        }
    }

    <%'** COD JCI-14022003-01 **%>
    function Operario(modo)
    {
        if (modo=="alta")
        {
            if (window.confirm("<%=LitMsgAltaOperarioConfirm%>")==true)
                document.location="personal.asp?dni=" + document.personal.hdni.value + "&domicilio="+document.personal.hdomicilio.value+"&operario=alta";
        }
        else
        {
            if (window.confirm("<%=LitMsgBajaOperarioConfirm%>")==true)
                document.location="personal.asp?dni=" + document.personal.hdni.value + "&domicilio="+document.personal.hdomicilio.value+"&operario=baja";
        }
    }

    function seleccionar(marco,formulario,check)
    {
        nregistros=eval(marco + ".document." + formulario + ".hNRegs.value-1");
        if (eval("document.personal." + check + ".checked"))
        {
            for (i=1;i<=nregistros;i++)
            {
                nombre="check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
            }
        }
        else
        {
            for (i=1;i<=nregistros;i++)
            {
                nombre="check" + i;
                eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
            }
        }
    }

    function GestionCostes(mode,operario)
    {
        if (mode=="delete")
        {
            if(!confirm("<%=LitMsgEliminarCostesConfirm%>")) return false;
        }
        fr_OperarioFases.document.operario_fases.action="operario_fases.asp?mode=" + mode + "&operario=" + operario;
        fr_OperarioFases.document.operario_fases.submit();
    }
    <%'** FIN COD JCI-14022003-01 **%>
</script>
<%mode=enc.EncodeForHtmlAttribute(null_s(request.querystring("mode")))%>

<body class="BODY_ASP">

<%viene=enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("viene"))))%>

<%Sub BarraOpciones2(modo,codigo,rnombre,si_tiene_modulo_comercial)%>
	<table width="100%">
	 	<tr>
			<td width="98%">
				<table  id="enlaces_extra"  cellpadding="0" cellspacing="0">
  	                <tr>
		                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=ventas/lista_presupuestos_cli.asp&pag2=ventas/lista_presupuestos_cli_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&viene=comercial&titulo=<%=LitListaPresupuestos%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrPresupuestos%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitPresupuestos%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=ventas/Lista_pedidos_cli.asp&pag2=ventas/Lista_pedidos_cli_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&viene=comercial&titulo=<%=LitListaPedidos%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrPedidos%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitPedidos%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=ventas/Lista_albaranes_cli.asp&pag2=ventas/Lista_albaranes_cli_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&viene=comercial&titulo=<%=LitListaAlbaranes%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrAlbaranes%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitAlbaranes%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=ventas/Lista_facturas_cli.asp&pag2=ventas/Lista_facturas_cli_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&viene=comercial&titulo=<%=LitListaFacturas%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrFacturas%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitFacturas%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <!--<td style="border : 1px solid black;" align="center" class="CELDACENTERB" onmouseover="this.className='TDACTIVO8'" onmouseout="this.className='CELDACENTERB'" bgcolor="<%=color_blau%>"><a class="CELDAREFB7" href="javascript:AbrirVentana('../ventas/clientes_buscar.asp?ndoc=<%=codigo%>&viene=comercial&titulo=LISTA DE CLIENTES','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='Ir a Clientes'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitClientes%>&nbsp;&nbsp;&nbsp;</a></td>-->
		                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=ventas/clientes.asp&pag2=ventas/clientes_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&viene=comercial&titulo=<%=LitListaClientes%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrClientes%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitClientes%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <%if si_tiene_modulo_comercial<>0 then%>
			                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=administracion/liqcomercial.asp&pag2=administracion/liqcomercial_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&viene=comercial&titulo=<%=LitListaLiquidaciones%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrLiquidaciones%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitLiquidaciones%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
			                <td><div><a class="CELDAREF7" href="javascript:Comisiones('<%=enc.EncodeForJavascript(null_s(codigo))%>','<%=enc.EncodeForJavascript(null_s(rnombre))%>');" OnMouseOver="self.status='<%=LitIrComisionArt%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitComisionesArt%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
			                <td><div><a class="CELDAREF7" href="javascript:ComisionesExtras('<%=enc.EncodeForJavascript(null_s(codigo))%>','<%=enc.EncodeForJavascript(null_s(rnombre))%>');" OnMouseOver="self.status='<%=LitIrComisionExt%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitComisionesExt%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <%end if%>
	                </tr>
                </table>
            </td>
        </tr>
    </table>
<%end sub

Sub BarraOpcionesGen(modo,codigo,rnombre)%>
	<table width="100%">
	 	<tr>
			<td width="98%">
				<table  id="enlaces_extra"  cellpadding="0" cellspacing="0">
					<tr>
	                    <%if (VerObjetoUsuario(OBJTpv, codigo)) then%>
			                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=ventas/Lista_tickets.asp&pag2=ventas/Lista_tickets_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&viene=personal&titulo=<%=enc.EncodeForJavascript(LitVentasTPV)%>','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitIrVentasTPV%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitVentasTPV%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <%end if

		                if (VerObjetoUsuario(OBJIncentivos, codigo)) then%>
			                <td><div><a class="CELDAREF7" href="javascript:Incentivos('<%=enc.EncodeForJavascript(null_s(codigo))%>','<%=enc.EncodeForJavascript(null_s(rnombre))%>');"','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitIrIncentivos%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitIncentivos%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <%end if
		                if (VerObjetoUsuario(OBJPenalizaciones, codigo)) then%>
			                <td><div><a class="CELDAREF7" href="javascript:Penalizaciones('<%=enc.EncodeForJavascript(null_s(codigo))%>','<%=enc.EncodeForJavascript(null_s(rnombre))%>');"','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitIrPenalizaciones%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitPenalizaciones%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <%end if
		                if si_tiene_modulo_tiendas<>0 then%>
			                <td><div><a class="CELDAREF7" href="javascript:Anulaciones('<%=enc.EncodeForJavascript(null_s(codigo))%>','<%=enc.EncodeForJavascript(null_s(rnombre))%>');"','P',<%=enc.EncodeForJavascript(AltoVentana)%>,<%=enc.EncodeForJavascript(AnchoVentana)%>)" OnMouseOver="self.status='<%=LitIrAnulaciones%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitAnulaciones%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
		                <%end if

	  	                'DGB: 26-03-10 enlace para el Control Horario
	  	                if si_tiene_modulo_ControlPresencia <> 0 then
	  	                    pagina="../netInic.asp?pag=/ControlHorario/CH_administracion_acceso.aspx&ndoc="&rdni%>
		                    <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(null_s(pagina))%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrAdmAcceso%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitAdmAcceso%></a></div></td><td class="der">|</td>
	       
	                        <% pagina="../netInic.asp?pag=/ControlHorario/CH_configuracion_calendario.aspx&ndoc="&rdni%>
	                        <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(null_s(pagina))%>','P',<%=AltoVentana+90%>,<%=AnchoVentana+90%>)" OnMouseOver="self.status='<%=LitIrAdmAcceso%>'; return true;" OnMouseOut="self.status=''; return true;"><%=LitCHCalendario%></a></div></td><td class="der">|</td>
	                    <%end if%>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
<%end sub

Sub BarraOpcionesOperario(modo,codigo,rnombre)%>
	<table width="100%">
	 	<tr>
			<td width="98%">
				<table  id="enlaces_extra"  cellpadding="0" cellspacing="0">
  	                <tr>
		                <td><div><a class="CELDAREF7" href="javascript:AbrirVentana('../central.asp?pag1=fabricacion/lista_notas_operario.asp&pag2=fabricacion/lista_notas_operario_bt.asp&ndoc=<%=enc.EncodeForJavascript(null_s(codigo))%>&titulo=<%=LitListadoNotas%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitIrNotas%>'; return true;" OnMouseOut="self.status=''; return true;">&nbsp;&nbsp;&nbsp;<%=LitNotas%>&nbsp;&nbsp;&nbsp;</a></div></td><td class="der">|</td>
	                </tr>
                </table>
            </td>
        </tr>
    </table>

<%end sub

'*************************************************************************************************************
'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(p_dni)

    'rst.Open "select * from personal where dni='" & p_dni & "'", _
	'	session("dsn_cliente"),adOpenKeyset,adLockOptimistic

    set conn = Server.CreateObject("ADODB.Connection")
    set rst = Server.CreateObject("ADODB.Recordset")
	set command =  Server.CreateObject("ADODB.Command")
	conn.open session("dsn_cliente")
    'conn.cursorlocation=3
	command.ActiveConnection =conn
	command.CommandTimeout = 60

    strselect="select * from personal where dni=?"

	command.CommandText=strselect
	command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@dni", adVarChar, adParamInput, 20, p_dni&"")

    rst.Open command, , adOpenKeyset, adLockOptimistic

	if rst.eof then
	   rst.addnew
       strselect="select usuario from clientes_users with(nolock) where usuario=? and ncliente=?"
       EsUsuarioGesticet= DLookupP2(strselect,p_dni&"",adVarChar,50,session("ncliente")&"",adVarchar,5,DSNIlion)                              
	else
       strselect="select usuario from clientes_users with(nolock) where usuario=? and ncliente=?"
       EsUsuarioGesticet= DLookupP2(strselect,rst("login")&"",adVarChar,50,session("ncliente")&"",adVarchar,5,DSNIlion)                              
	end if

   if mode="first_save" then
		submode="add&submode=add"
	else
		submode="edit&dni=" + p_dni
	end if

	if mode="save" then
		if p_dni <> rst("dni") then
	   		guarda = false
	   		rst.cancelupdate
	   		rst.close

            conn.close
            set conn    =  nothing
            set command =  nothing
            set rst  =  nothing	
            %>
	   		<script language="javascript" type="text/javascript">
	   		    window.alert("<%=LitMsgModifDni%>");
	   		    document.personal.action="personal.asp?mode=<%=enc.EncodeForJavascript(null_s(submode))%>";
	   		    parent.pantalla.document.personal.submit();
	   		    parent.botones.document.location="personal_bt.asp?mode=<%=enc.EncodeForJavascript(null_s(submode))%>";
	   		</script>
	   	<%else
			 '2010072:COMENTO ESTA ACTUALIZACIÓN, DE TODOS LOS CONTACTOS DE LA PERSONA QUE SE ACTUALIZA, PARA QUE NO SE MODIFIQUE EL NIVEL DE SUS CONTACTOS. 
			'rstAux.Open "update contactoscomercial with(rowlock) set nivel=" & null_z(request.form("nivel")) & " where comercial='" & p_dni & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			'''''''''''''''''''''
			guarda = true
		end if
	else
		guarda=true
	end if

	if si_tiene_paginaSMS<>0 then
		if guarda=true and EsUsuarioGesticet="" and nz_b(request.form("mensajeria_sms"))<>0 then
			guarda=false
			rst.cancelupdate
		   	rst.close

            conn.close
            set conn    =  nothing
            set command =  nothing
            set rst  =  nothing	
			%><script language="javascript" type="text/javascript">
			      alert("<%=LitUsuarioEgesticet%>");
			      parent.window.frames["botones"].document.location = "personal_bt.asp?mode=add";
			      document.personal.action="personal.asp?mode=<%=enc.EncodeForJavascript(null_s(submode))%>";
			      parent.pantalla.document.personal.submit();
			</script><%
		end if
	end if

	if guarda=true then
		codigoop=request.form("codigo")
		'Si el codigo es vacio se pone el dni (sin codigo de empresa)'
		if codigoop&""="" then
			codigoop=trimCodEmpresa(Nulear(p_dni))
		end if

		if codigoop>"" then
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60

			'rstAux.open "select codigo from personal with(nolock) where codigo='" & codigoop & "' and dni like '" & session("ncliente") & "%' and dni<>'"&Nulear(p_dni)&"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            strselect = "select codigo from personal with(nolock) where codigo=? and dni like ?+'%' and dni<>?"
            command2.CommandText=strselect
            command2.CommandType = adCmdText

            command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,15,codigoop&"")
            command2.Parameters.Append command2.CreateParameter("@dni1",adVarChar,adParamInput,20,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@dni2",adVarChar,adParamInput,20,Nulear(p_dni)&"")
                
            set rstAux3= command2.Execute

			if not rstAux3.eof then
				rstAux3.close
                conn2.close
                set conn2    =  nothing
                set command2 =  nothing
                set rstAux3  =  nothing	
				guarda = false
		   		rst.cancelupdate
		   		rst.close%>
	   			<script language="javascript" type="text/javascript">
	   			    window.alert("<%=LitMsgCodigoOpExiste%><%=codigoop%>");
	   			    document.personal.action="personal.asp?mode=<%=enc.EncodeForJavascript(null_s(submode))%>";
	   			    parent.pantalla.document.personal.submit();
	   			    parent.botones.document.location="personal_bt.asp?mode=<%=enc.EncodeForJavascript(null_s(submode))%>";
		   		</script>
		   	<%else
				rstAux3.close
                conn2.close
                set conn2    =  nothing
                set command2 =  nothing
                set rstAux3  =  nothing	
			end if
		end if
	end if

	if guarda=true then
		rst("dni")        =  Nulear(p_dni)
		if request.form("antiguedad")>"" then
	      	rst("antiguedad") = cdate(request.form("antiguedad"))
		else
			rst("antiguedad")=nulear(Request.form("antiguedad"))
		end if
		rst("observaciones") = request.form("observaciones")
		rst("sueldo") = miround(Null_z(request.form("sueldo")),2)
		rst("jornada") = reemplazar(null_z(request.form("jornada")),".",",")
		rst("ss") = request.form("ss")
		rst("nombre") = request.form("nombre")
		rst("alias") = request.form("alias")
		rst("codigo") = codigoop&""   'nulear(request.form("codigo"))
		rst("phextra")= miround(null_z(request.form("phextra")),dec_prec)
		rst("nivel")= null_z(request.form("nivel"))
		rst("email")=nulear(request.form("email"))
		rst("caja")=Nulear(request.form("caja"))
		rst("telefono2")      = Nulear(request.form("telefono2"))
		rst("fax")            = Nulear(request.form("fax"))
    	rst("fbaja")		  = Nulear(request.form("fbaja"))
    	rst("hora_ini_ma")	  = Nulear(request.form("horaIniMa"))
    	rst("hora_fin_ma")	  = Nulear(request.form("horaFinMa"))
    	rst("hora_ini_ta")	  = Nulear(request.form("horaIniTa"))
    	rst("hora_fin_ta")	  = Nulear(request.form("horaFinTa"))
    	rst("IRPF")           =miround(Null_z(request.form("IRPF")),2)
    	rst("importe_ss")     =miround(null_z(request.form("segsocial")),2)
        rst("maxamount")      =miround(null_z(request.form("maxamount")),2)
    	'dgb control horario modulo Kyros  19/11/2009
        'dgm The Department of time control is always displayed.
    	if si_tiene_modulo_ControlPresencia <> 0 then
    	    rst("fdepartamento")=Nulear(request.Form("departamento"))
    	else
            rst("department")=Nulear(request.Form("departamento"))
        end if
    			
  	    if (si_tiene_modulo_bierzo<>0) then 
		    rst("campo01") = limpiaCadena(request.form("tarifa"))
		    rst("campo02") = limpiaCadena(request.form("porctarifa"))
		    rst("campo03") = limpiaCadena(request.form("simayor"))
		end if

		if mode="first_save" then
			rst("login")=trimCodEmpresa(Nulear(p_dni))
		end if
		rst("tipo") = Nulear(request.form("tipo"))
		''MPC 18/06/2010 Se comenta por petición de JAR
		''No tiene sentido ya que se puede o no actualizar este campo
		''if si_tiene_paginaSMS<>0 then
			''if nz_b(request.form("mensajeria_sms"))<>0 and rst("telefono2")&""<>"" then
	        'set command2 = nothing
            'set conn2 = Server.CreateObject("ADODB.Connection")
            'set command2 =  Server.CreateObject("ADODB.Command")
            'conn2.open DSNIlion
            'conn2.cursorlocation=3
            'command2.ActiveConnection =conn2
            'command2.CommandTimeout = 60

            set rstAux = Server.CreateObject("ADODB.Recordset")

			if nz_b(request.form("mensajeria_sms"))<>0 then
				'rstAux.open "update clientes_users with(updlock) set movil='" & iif(instr(rst("telefono2"),"+34"),rst("telefono2"),"+34" & rst("telefono2")) & "' where usuario='" & rst("login") & "' and ncliente='" & session("ncliente") & "'",DSNIlion
                rst("mensajeria_sms")=nz_b(request.form("mensajeria_sms"))
                strupdate = "update clientes_users with(rowlock) set movil=? where usuario=? and ncliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 =  Server.CreateObject("ADODB.Command")
                conn2.Open = DSNIlion
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText=strupdate
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@movil",adVarChar,adParamInput,18,iif(instr(rst("telefono2"),"+34"),rst("telefono2")&"","+34" & rst("telefono2"))&"")
                command2.Parameters.Append command2.CreateParameter("@usuario",adVarChar,adParamInput,50,rst("login")&"")
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,session("ncliente")&"")
                set rstAux= command2.Execute

                'if not rstAux.eof then 
                '  'rstAux.close
                'else
                '  'rstAux.close                   
                'end if
	            conn2.close
	            set conn2    =  nothing
	            set command2 =  nothing
	            set rstAux  =  nothing					
			else
				rst("mensajeria_sms")=0
				'rstAux.open "update clientes_users with(updlock) set movil=NULL where usuario='" & rst("login") & "' and ncliente='" & session("ncliente") & "'",DSNIlion
                strupdate = "update clientes_users with(updlock) set movil=NULL where usuario=? and ncliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 =  Server.CreateObject("ADODB.Command")
                conn2.Open = DSNIlion
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText=strupdate
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@usuario",adVarChar,adParamInput,50,rst("login")&"")
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
                set rstAux= command2.Execute

                 'if not rstAux.eof then 
                   'rstAux.close
                'else
                   'rstAux.close                   
                'end if
	            conn2.close
	            set conn2    =  nothing
	            set command2 =  nothing
	            set rstAux  =  nothing	

			end if

           
		''else
            'set command2 = nothing
            'set conn2 = Server.CreateObject("ADODB.Connection")
            'set command2 =  Server.CreateObject("ADODB.Command")
            'conn2.open DSNIlion
            'conn2.cursorlocation=3
            'command2.ActiveConnection =conn2
            'command2.CommandTimeout = 6
			if rst("telefono2")&""<>"" then
				'rstAux.open "update clientes_users with(updlock) set movil='" & iif(instr(rst("telefono2"),"+34"),rst("telefono2"),"+34" & rst("telefono2")) & "' where usuario='" & rst("login") & "' and ncliente='" & session("ncliente") & "'",DSNIlion
                strupdate= "update clientes_users with(updlock) set movil=? where usuario=? and ncliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 =  Server.CreateObject("ADODB.Command")
                conn2.Open = DSNIlion
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText=strupdate
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@movil",adVarChar,adParamInput,18,iif(instr(rst("telefono2"),"+34"),rst("telefono2")&"","+34" & rst("telefono2")) &"")
                command2.Parameters.Append command2.CreateParameter("@usuario",adVarChar,adParamInput,50,rst("login")&"")
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
                set rstAux= command2.Execute
			else
				'rstAux.open "update clientes_users with(updlock) set movil=NULL where usuario='" & rst("login") & "' and ncliente='" & session("ncliente") & "'",DSNIlion
                strupdate= "update clientes_users with(updlock) set movil=NULL where usuario=? and ncliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 =  Server.CreateObject("ADODB.Command")
                conn2.Open = DSNIlion
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText=strupdate
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@usuario",adVarChar,adParamInput,50,rst("login")&"")
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
                set rstAux= command2.Execute
			end if
	        conn2.close
	        set conn2    =  nothing
	        set command2 =  nothing
	        'set rstAux  =  nothing	
		''end if
		''FIN MPC 18/06/2010
		rst.update
		rst.close
        'conn.close
        'set conn    =  nothing
        'set command =  nothing
        'set rst  =  nothing	

        'dgb 21/01/2010  CONTROL HORARIO
        if si_tiene_modulo_ControlPresencia <> 0 then
            set command2 = nothing
    	    set connCH = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            connCH.open session("dsn_cliente")
            command2.ActiveConnection =connCH
            command2.CommandTimeout = 0
            command2.CommandText="CH_PersonalCalendario"
            command2.CommandType = adCmdStoredProc            

            command2.Parameters.Append command2.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@dni",adVarChar,adParamInput,20,Nulear(p_dni)&"")

			set rstCH = command2.Execute
			rstCH.close
			connCH.Close
			set rstCH=nothing                     
    	end if
    	
        set conn = Server.CreateObject("ADODB.Connection")
        set rst = Server.CreateObject("ADODB.Recordset")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open session("dsn_cliente")
        'conn.cursorlocation=3
        command.ActiveConnection =conn
        command.CommandTimeout = 60

        strselect="select * from domicilios where pertenece=? and tipo_domicilio='PERSONAL'"	

        command.CommandText=strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@pertenece", adVarChar, adParamInput, 20, p_dni&"")

        rst.Open command, , adOpenKeyset, adLockOptimistic

		if rst.eof then
	   		rst.addnew
		end if

		rst("pertenece")      = Nulear(p_dni)
		rst("tipo_domicilio") = "PERSONAL"
		rst("domicilio")      = Nulear(request.form("domicilio"))
		rst("cp")             = Nulear(request.form("cp"))
		rst("poblacion")      = Nulear(request.form("poblacion"))
    	rst("provincia")      = Nulear(request.form("provincia"))
		rst("pais")           = Nulear(request.form("pais"))
    	rst("telefono")       = Nulear(request.form("telefono"))
		rst.Update
		rst.Close
        conn.close
	    set conn    =  nothing
	    set command =  nothing
	    set rst  =  nothing	

        strselect = "select comercial from comerciales with(nolock) where comercial=?"
        comercial = DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente"))

        strselect = "select dni from tecnicos where dni=?"
        tecnico = DLookupP1(strselect, p_dni&"", adVarChar,20,session("dsn_cliente"))

        strselect= "select operario from operarios where operario = ?"
        operario = DLookupP1(strselect, p_dni&"", adVarChar,20,session("dsn_cliente"))
	   

		if comercial>"" then	   		
            set conn = Server.CreateObject("ADODB.Connection")
            set rst = Server.CreateObject("ADODB.Recordset")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("dsn_cliente")
            'conn.cursorlocation=3
            command.ActiveConnection =conn
            command.CommandTimeout = 60            

            strselect="select * from comerciales with(rowlock) where comercial=?"          
            command.CommandText=strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@comercial", adVarChar, adParamInput, 20, comercial&"")

            rst.Open command, , adOpenKeyset, adLockOptimistic

	   		rst("cventas")   = reemplazar(null_z(request.form("cventas")),".",",")
	   		rst("mganancia") = reemplazar(null_z(request.form("mganancia")),".",",")
	   		rst("objetivo")  = reemplazar(null_z(request.form("objetivo")),".",",")
	   		rst("per_ob")    = reemplazar(null_z(request.form("per_ob")),".",",")
	   		rst("combase")   = miround(null_z(request.form("cbase")),decpor)
	   		rst("comconceptos") = reemplazar(null_z(request.form("cconcepto")),".",",")
	   		rst("PENALIZACION")  = miround(null_z(request.form("pena")),decpor)
	   		rst("superior")    = nulear(request.form("superior"))
	   		if nulear(request.form("fbaja"))>"" then
				' Desasignamos el comercial de todos los clientes y de todos los centros'
                strselect = "select ncliente,comercial from clientes with(nolock) where ncliente like ?+'%' and comercial=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,session("ncliente")&"")
                command2.Parameters.Append command2.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
                rstAux.CursorLocation = adUseClient
                rstAux.Open command2, ,adOpenKeyset, adLockOptimistic
				while not rstAux.eof
                    strupdate = "update clientes with(updlock) set comercial=NULL where ncliente=?"
                    set command3 = nothing
                    set conn3 = Server.CreateObject("ADODB.Connection")
                    set command3 = Server.CreateObject("ADODB.Command")
                    conn3.Open = session("dsn_cliente")
                    conn3.CursorLocation = 3
                    command3.ActiveConnection = conn3
                    command3.CommandTimeout = 60
                    command3.CommandText = strupdate
                    command3.CommandType = adCmdText
                    command3.Parameters.Append command3.CreateParameter("@ncliente",adVarChar,adParamInput,10,session("ncliente")&"")
                    set rstAux2 = command3.Execute
                    conn3.Close
                    set conn3 = nothing
                    set command3 = nothing
					rstAux.movenext
				wend
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
                strselect = "select ncentro,comercial from centros with(nolock) where ncentro like ?+'%' and comercial=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@ncentro",adVarChar,adParamInput,10,session("ncliente")&"")
                command2.Parameters.Append command2.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
                rstAux.CursorLocation = adUseClient
                rstAux.Open command2, ,adOpenKeyset, adLockOptimistic
				while not rstAux.eof
                    strupdate = "update centros with(updlock) set comercial=NULL where ncentro=?"
                    set command3 = nothing
                    set conn3 = Server.CreateObject("ADODB.Connection")
                    set command3 = Server.CreateObject("ADODB.Command")
                    conn3.Open = session("dsn_cliente")
                    conn3.CursorLocation = 3
                    command3.ActiveConnection = conn3
                    command3.CommandTimeout = 60
                    command3.CommandText = strupdate
                    command3.CommandType = adCmdText
                    command3.Parameters.Append command3.CreateParameter("@ncentro",adVarChar,adParamInput,10,rstAux("ncentro")&"")
                    set rstAux2 = command3.Execute
                    conn3.Close
                    set conn3 = nothing
                    set command3 = nothing
					rstAux.movenext
				wend
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
	   			rst("fbaja") = nulear(request.form("fbaja"))
	   		end if
	   		rst.update
	   		rst.close
            conn.close
	        set conn    =  nothing
	        set command =  nothing
	        set rst  =  nothing	
		end if

		if tecnico>"" then

            set conn = Server.CreateObject("ADODB.Connection")
            set rst = Server.CreateObject("ADODB.Recordset")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("dsn_cliente")
            'conn.cursorlocation=3
            command.ActiveConnection =conn
            command.CommandTimeout = 60

			'rst.open "select * from tecnicos with(rowlock) where dni='" & tecnico & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            strselect="select * from tecnicos with(rowlock) where dni=?"
            command.CommandText=strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@dni", adVarChar, adParamInput, 20, tecnico&"")

            rst.Open command, , adOpenKeyset, adLockOptimistic

	   		rst("comision")  = miround(null_z(request.form("tcomision")),decpor)
	   		rst("phextralab") = miround(null_z(request.form("tphextralab")),dec_prec)
			rst("phextrafes") = miround(null_z(request.form("tphextrafes")),dec_prec)
			rst("phlaboral") = miround(null_z(request.form("tphlaboral")),dec_prec)
	   		rst("incentivo1")  = miround(null_z(request.form("tincentivo1")),dec_prec)
	   		rst("incentivo2")    = miround(null_z(request.form("tincentivo2")),dec_prec)
			rst("almacen")    = nulear(request.form("talmacen"))
			rst("vehiculo")    = nulear(request.form("tvehiculo"))
	   		if nulear(request.form("fbaja"))>"" then
	   			rst("fbaja") = nulear(request.form("fbaja"))
	   		end if
	   		rst.update
	   		rst.close
            conn.close
	        set conn    =  nothing
	        set command =  nothing
	        set rst  =  nothing	
		end if

		if operario>"" then

            set conn = Server.CreateObject("ADODB.Connection")
            set rst = Server.CreateObject("ADODB.Recordset")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("dsn_cliente")
            'conn.cursorlocation=3
            command.ActiveConnection =conn
            command.CommandTimeout = 60
		
            strselect="select * from operarios with(updlock) where operario=?"
            command.CommandText=strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@operario", adVarChar, adParamInput, 20, operario&"")

            rst.Open command, , adOpenKeyset, adLockOptimistic

	   		rst("coste_hora")  = miround(null_z(request.form("ocoste_hora")),dec_prec)
	   		if nulear(request.form("fbaja"))>"" then
	   			rst("fbaja") = nulear(request.form("fbaja"))
	   		end if
	   		rst.update
	   		rst.close
            conn.close
	        set conn    =  nothing
	        set command =  nothing
	        set rst  =  nothing	
		end if
	end if
end sub

'*************************************************************************************************************
sub EliminarRegistro()
    w_dni = enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("dni"))))

    rstAux.cursorlocation=3
    rstAux.open "select comercial from pedidos_cli with(nolock) where comercial='" & w_dni & "' union " & _
	            "select comercial from albaranes_cli with(nolock) where comercial='" & w_dni & "' union " & _
                "select comercial from facturas_cli with(nolock) where comercial='" & w_dni & "' union " & _
                "select comercial from presupuestos_cli with(nolock) where comercial='" & w_dni & "' union " & _
				"select dni from hojas_gastos with(nolock) where dni='" & w_dni & "' union " & _
				"select responsable from movimientos with(nolock) where responsable='" & w_dni & "' union " & _
				"select comercial from contactoscomercial with(nolock) where comercial='" & w_dni & "' union " & _
				"select dni from grupos_pro with(nolock) where dni='" & w_dni & "' union " & _
				"select tecnico from trabajo with(nolock) where tecnico='" & w_dni & "' union " & _
				"select tecnico from parte with(nolock) where tecnico='" & w_dni & "' union " & _
				"select tecnico from tecauxiliar with(nolock) where tecnico='" & w_dni & "' union " & _
				"select tecnico from ordenes with(nolock) where tecnico='" & w_dni & "' union " & _
				"select operario from fabricar with(nolock) where operario='" & operario & "' union " & _
				"select comercial from comerciales with(nolock) where superior='" & w_dni & "' union " & _
				"select personal from detalles_pen with(nolock) where personal ='" & W_dni & "' union " & _
				"select comercial from vencimientos_salida with(nolock) where comercial='" & W_dni & "' union " & _
				"select codigo from detalles_inc with(nolock) where personal='" & W_dni & "' union " & _
				"select comercial from clientes with(nolock) where comercial='" & W_dni & "'",session("dsn_cliente")
	if rstAux.eof then
        rstAux.Close
        rstAux.cursorlocation=2
        strdelete = "delete from costes_fases with(rowlock) where operario=?"
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux = Server.CreateObject("ADODB.Command")
        connAux.Open = session("dsn_cliente")
        connAux.CursorLocation = 3
        commandAux.ActiveConnection = connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText = strdelete
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@operario",adVarChar,adParamInput,20,w_dni&"")
        set rstAux = commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
        strdelete = "delete from operarios with(rowlock) where operario=?"
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux = Server.CreateObject("ADODB.Command")
        connAux.Open = session("dsn_cliente")
        connAux.CursorLocation = 3
        commandAux.ActiveConnection = connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText = strdelete
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@operario",adVarChar,adParamInput,20,w_dni&"")
        set rstAux = commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
        strdelete = "delete from comerciales with(rowlock) where comercial=?"
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux = Server.CreateObject("ADODB.Command")
        connAux.Open = session("dsn_cliente")
        connAux.CursorLocation = 3
        commandAux.ActiveConnection = connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText = strdelete
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@comercial",adVarChar,adParamInput,20,w_dni&"")
        set rstAux = commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
       	strdelete = "delete from tecnicos with(rowlock) where dni=?"
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux = Server.CreateObject("ADODB.Command")
        connAux.Open = session("dsn_cliente")
        connAux.CursorLocation = 3
        commandAux.ActiveConnection = connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText = strdelete
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@dni",adVarChar,adParamInput,20,w_dni&"")
        set rstAux = commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
		strdelete = "delete from domicilios with(rowlock) where pertenece=? and tipo_domicilio='PERSONAL'"
        set command = nothing
        set conn = Server.CreateObject("ADODB.Connection")
        set command = Server.CreateObject("ADODB.Command")
        conn.Open = session("dsn_cliente")
        conn.CursorLocation = 3
        command.ActiveConnection = conn
        command.CommandTimeout = 60
        command.CommandText = strdelete
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@pertenece",adVarChar,adParamInput,55,w_dni&"")
        set rst = command.Execute
        conn.Close
        set conn = nothing
        set command = nothing
        strdelete = "delete from personal with(rowlock) where dni=?"
        set command = nothing
        set conn = Server.CreateObject("ADODB.Connection")
        set command = Server.CreateObject("ADODB.Command")
        conn.Open = session("dsn_cliente")
        conn.CursorLocation = 3
        command.ActiveConnection = conn
        command.CommandTimeout = 60
        command.CommandText = strdelete
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@dni",adVarChar,adParamInput,55,w_dni&"")
        set rst = command.Execute
        conn.Close
        set conn = nothing
        set command = nothing
	else
        rstAux.Close%>
		<script>
			window.alert("<%=LitMsgBorrarPersona%>");
			document.location="personal.asp?mode=browse&dni=<%=enc.EncodeForHtmlAttribute(null_s(w_dni))%>";
		</script>
	<%end if
end sub


'*************************************************************************************************************'
function tienePagina(pagina)
    ''ricardo 25-9-2009 como se quita la tabla accesos, se cambia el select para saber si el usuario tiene el item para esa empresa
    tienePagina=0
	if VerObjeto(pagina)=true then
		tienePagina=1
	end if
end function
'*************************************************************************************************************'

'*************************************************************************************************************
'-----------------------------------------------------
' FUNCION PARA DAR DE ALTA UNA PERSONA COMO COMERCIAL
'-----------------------------------------------------
sub AltaComercial(p_dni)
    strselect = "select * from comerciales where comercial=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
    if rstAux.eof then
    	    rstAux.addnew
    	    rstAux("comercial")=p_dni
            rstAux.Update
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
    else
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
            strupdate = "update comerciales with(updlock) set fbaja=null where comercial=?"
            set commandAux = nothing
            set connAux = Server.CreateObject("ADODB.Connection")
            set commandAux = Server.CreateObject("ADODB.Command")
            connAux.Open = session("dsn_cliente")
            connAux.CursorLocation = 3
            commandAux.ActiveConnection = connAux
            commandAux.CommandTimeout = 60
            commandAux.CommandText = strupdate
            commandAux.CommandType = adCmdText
            commandAux.Parameters.Append commandAux.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
            set rstAux = commandAux.Execute
            connAux.Close
            set connAux = nothing
            set commandAux = nothing
    end if%>
   <script language="javascript" type="text/javascript">
       document.location="personal.asp?dni="+document.personal.hcomercial.value+"&mode=edit";
       parent.botones.location="personal_bt.asp?mode=edit";
    </script>
<%end sub

'*************************************************************************************************************
'-----------------------------------------------------
' FUNCION PARA DAR DE baja UNA PERSONA COMO COMERCIAL
'-----------------------------------------------------
sub BajaComercial(p_dni)
	' Desasignamos el comercial de todos los clientes y de todos los centros'
    strselect = "select ncliente,comercial from clientes with(nolock) where ncliente like ?+'%' and comercial=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@ncliente",adVarChar,adParamInput,10,session("ncliente")&"")
    commandAux.Parameters.Append commandAux.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
	while not rstAux.eof
        strupdate = "update clientes with(updlock) set comercial=NULL where ncliente=?"
        set commandAux2 = nothing
        set connAux2 = Server.CreateObject("ADODB.Connection")
        set commandAux2 = Server.CreateObject("ADODB.Command")
        connAux2.Open = session("dsn_cliente")
        connAux2.CursorLocation = 3
        commandAux2.ActiveConnection = connAux2
        commandAux2.CommandTimeout = 60
        commandAux2.CommandText = strupdate
        commandAux2.CommandType = adCmdText
        commandAux2.Parameters.Append commandAux2.CreateParameter("@ncliente",adVarChar,adParamInput,10,rstAux("ncliente")&"")
        set rstAux2 = commandAux2.Execute
        connAux2.Close
        set connAux2 = nothing
        set commandAux2 = nothing
		rstAux.movenext
	wend
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
    strselect = "select ncentro,comercial from centros with(nolock) where ncentro like ?+'%' and comercial=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@ncentro",adVarChar,adParamInput,10,session("ncliente")&"")
    commandAux.Parameters.Append commandAux.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
	while not rstAux.eof
        strupdate = "update centros with(updlock) set comercial=NULL where ncentro=?"
        set commandAux2 = nothing
        set connAux2 = Server.CreateObject("ADODB.Connection")
        set commandAux2 = Server.CreateObject("ADODB.Command")
        connAux2.Open = session("dsn_cliente")
        connAux2.CursorLocation = 3
        commandAux2.ActiveConnection = connAux2
        commandAux2.CommandTimeout = 60
        commandAux2.CommandText = strupdate
        commandAux2.CommandType = adCmdText
        commandAux2.Parameters.Append commandAux2.CreateParameter("@ncentro",adVarChar,adParamInput,10,rstAux("ncentro")&"")
        set rstAux2 = commandAux2.Execute
        connAux2.Close
        set connAux2 = nothing
        set commandAux2 = nothing
		rstAux.movenext
	wend
    connAux.Close
    set connAux = nothing
    set commandAux = nothing
	' No eliminamos el registro de la base de datos, sino que le asignamos una fecha de baja y recargamos la página'
    strselect = "select * from comerciales with(rowlock) where comercial=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
	if not rstAux.eof then
		rstAux("fbaja")=day(date) & "/" & month(date) & "/" & year(date)
		rstAux.update
	end if
    connAux.Close
    set connAux = nothing
    set commandAux = nothing%>
   <script language="javascript" type="text/javascript">
       document.location="personal.asp?dni="+document.personal.hcomercial.value+"&mode=browse";
       parent.botones.location="personal_bt.asp?mode=browse";
    </script>
<%end sub

'*************************************************************************************************************
'-----------------------------------------------------
' FUNCION PARA DAR DE ALTA UNA PERSONA COMO TECNICO
'-----------------------------------------------------
sub AltaTecnico(p_dni)
    strselect  = "select * from tecnicos where dni=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@dni",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
    if rstAux.eof then
    	rstAux.addnew
    	rstAux("dni")=p_dni
        rstAux.Update
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
    else
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
        strupdate = "update tecnicos with(updlock) set fbaja=NULL where dni=?"
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux = Server.CreateObject("ADODB.Command")
        connAux.Open = session("dsn_cliente")
        connAux.CursorLocation = 3
        commandAux.ActiveConnection = connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText = strupdate
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@dni",adVarChar,adParamInput,20,p_dni&"")
        set rstAux = commandAux.Execute
        connAux.Close
        set connAux = nothing
        set commandAux = nothing
    end if%>
    <script language="javascript" type="text/javascript">
        document.location="personal.asp?dni=" + document.personal.htecnico.value + "&mode=edit";
        parent.botones.location="personal_bt.asp?mode=edit";
    </script>
<%end sub

'*************************************************************************************************************
'-----------------------------------------------------
' FUNCION PARA DAR DE baja UNA PERSONA COMO TECNICO
'-----------------------------------------------------
sub BajaTecnico(p_dni)
	'' No eliminamos el registro de la base de datos, sino que le asignamos una fecha de baja y recargamos la página
    strselect = "select * from tecnicos with(rowlock) where dni=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@dni",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
	if not rstAux.eof then
		rstAux("fbaja")=day(date) & "/" & month(date) & "/" & year(date)
		rstAux.update
	end if
	connAux.Close
    set connAux = nothing
    set commandAux = nothing%>
   <script language="javascript" type="text/javascript">
       document.location="personal.asp?dni=" + document.personal.htecnico.value + "&mode=browse";
       parent.botones.location="personal_bt.asp?mode=browse";
    </script>
<%end sub

'** COD JCI-14022003-01 **
'*************************************************************************************************************
'-----------------------------------------------------
' FUNCION PARA DAR DE ALTA UNA PERSONA COMO OPERARIO
'-----------------------------------------------------'
sub AltaOperario(p_dni)
    strselect = "select operario,fbaja from operarios where operario=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@operario",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
    if rstAux.eof then
    	rstAux.addnew
    	rstAux("operario")=p_dni
        rstAux.Update
	    connAux.Close
        set connAux = nothing
        set commandAux = nothing
    else
	    connAux.Close
        set connAux = nothing
        set commandAux = nothing
        strupdate = "update operarios with(updlock) set fbaja=NULL where operario=?"
        set commandAux = nothing
        set connAux = Server.CreateObject("ADODB.Connection")
        set commandAux = Server.CreateObject("ADODB.Command")
        connAux.Open = session("dsn_cliente")
        connAux.CursorLocation = 3
        commandAux.ActiveConnection = connAux
        commandAux.CommandTimeout = 60
        commandAux.CommandText = strupdate
        commandAux.CommandType = adCmdText
        commandAux.Parameters.Append commandAux.CreateParameter("@operario",adVarChar,adParamInput,20,p_dni&"")
    	set rstAux = commandAux.Execute
	    connAux.Close
        set connAux = nothing
        set commandAux = nothing
    end if%>
    <script language="javascript" type="text/javascript">
        document.location="personal.asp?dni=" + document.personal.hoperario.value + "&mode=edit";
        parent.botones.location="personal_bt.asp?mode=edit";
    </script>
<%end sub

'*************************************************************************************************************
'-----------------------------------------------------
' FUNCION PARA DAR DE baja UNA PERSONA COMO OPERARIO
'-----------------------------------------------------
sub BajaOperario(p_dni)
	'' No eliminamos el registro de la base de datos, sino que le asignamos una fecha de baja y recargamos la página
    strselect = "select * from operarios with(rowlock) where operario=?"
    set commandAux = nothing
    set connAux = Server.CreateObject("ADODB.Connection")
    set commandAux = Server.CreateObject("ADODB.Command")
    connAux.Open = session("dsn_cliente")
    connAux.CursorLocation = 3
    commandAux.ActiveConnection = connAux
    commandAux.CommandTimeout = 60
    commandAux.CommandText = strselect
    commandAux.CommandType = adCmdText
    commandAux.Parameters.Append commandAux.CreateParameter("@operario",adVarChar,adParamInput,20,p_dni&"")
    rstAux.CursorLocation = adUseClient
    rstAux.Open commandAux, ,adOpenKeyset, adLockOptimistic
	if not rstAux.eof then
		rstAux("fbaja")=day(date) & "/" & month(date) & "/" & year(date)
		rstAux.update
	end if
	connAux.Close
    set connAux = nothing
    set commandAux = nothing%>
    <script language="javascript" type="text/javascript">
        document.location="personal.asp?dni=" + document.personal.hoperario.value + "&mode=browse";
        parent.botones.location="personal_bt.asp?mode=browse";
    </script>
<%end sub

'** FIN COD JCI-14022003-01 **

'*************************************************************************************************************

'Crea la tabla que contiene la barra de grupos de datos (Generales,Comerciales,etc)
sub BarraNavegacion(modo)
	%>
	<script language="javascript" type="text/javascript">
	    jQuery("#<%="S_"&modo & "DG"%>").show();
	    jQuery("#<%="S_"&modo & "DC"%>").show();
        
	    <%if si_tiene_modulo_mantenimiento<>0 then%>
		    jQuery("#<%="S_"&modo & "DT"%>").show();
	    <%else%>
            jQuery("#<%="S_"&modo & "DT"%>").hide();
	    <%end if
        if si_tiene_modulo_produccion<>0 then%>
		    jQuery("#<%="S_"&modo & "DO"%>").show();
	    <%else %>
            jQuery("#<%="S_"&modo & "DO"%>").hide();
	    <%end if %>
    </script>
    <%
end sub

'****************************************************************************************************************
'---------------------------------------------
'Código de la página
'---------------------------------------------
set connRound = Server.CreateObject("ADODB.Connection")
connRound.open dsnilion

%>
<form name="personal" method="post">
    <%si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)
	si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)
	si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
	si_tiene_modulo_bierzo=ModuloContratado(session("ncliente"),ModBierzo)
	'dgb: control Presencia asociado con Kyros
	si_tiene_modulo_ControlPresencia=ModuloContratado(session("ncliente"),ModControlPresencia)
    %>

	<input type="hidden" name="h_si_tiene_modulo_comercial" value="<%=enc.EncodeForHtmlAttribute(si_tiene_modulo_comercial)%>">
	<input type="hidden" name="h_si_tiene_modulo_produccion" value="<%=enc.EncodeForHtmlAttribute(si_tiene_modulo_produccion)%>">
	<input type="hidden" name="h_si_tiene_modulo_mantenimiento" value="<%=enc.EncodeForHtmlAttribute(si_tiene_modulo_mantenimiento)%>">
    
	<%'Recordsets
    set rst = Server.CreateObject("ADODB.Recordset")
    set rstAux = Server.CreateObject("ADODB.Recordset")
    set rstAux2 = Server.CreateObject("ADODB.Recordset")
    set rstDom = Server.CreateObject("ADODB.Recordset")

    'Leer parámetros de la página'
    mode = enc.EncodeForHtmlAttribute(null_s(Request.QueryString("mode")))
    submode = enc.EncodeForHtmlAttribute(null_s(Request.QueryString("submode")))
    p_dni = enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("dni"))))
    p_domicilio = enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("domicilio"))))

	si_tiene_paginaSMS=tienePagina(OBJMensAMoviles)%>
    <input type="hidden" name="mode_accesos_tienda" value="<%=enc.EncodeForHtmlAttribute(null_s(mode))%>">
	<input type=hidden name="si_tiene_paginaSMS" value="<%=enc.EncodeForHtmlAttribute(null_s(si_tiene_paginaSMS))%>">
	<%'Informacion de la mensajería SMS de la empresa
    strselect = "select mensajeria_sms from configuracion with(Nolock) where nempresa=?"%>

  	<input type=hidden name="mensajeria_smsbd" value="<%=nz_b(DLookupP1(strselect,session("ncliente")&"",adVarChar,5,session("dsn_cliente")))%>">
  	

  	<%if mode="first_save" or submode="add" then
		rdni           = limpiaCadena(request.form("dni"))
		rnombre        = limpiaCadena(request.form("nombre"))
		ralias         = limpiaCadena(request.form("alias"))
		rcodigo        = limpiaCadena(request.form("codigo"))
		rtipo          = limpiaCadena(request.form("tipo"))
		rantiguedad    = limpiaCadena(request.form("antiguedad"))
		rhorainima	   = limpiaCadena(request.form("horaIniMa"))
		rhorafinma	   = limpiaCadena(request.form("horaFinMa"))
		rhorainita	   = limpiaCadena(request.form("horaIniTa"))
		rhorafinta	   = limpiaCadena(request.form("horaFinTa"))
		rss            = limpiaCadena(request.form("ss"))
		rjornada       = limpiaCadena(request.form("jornada"))
		rsueldo        = limpiaCadena(request.form("sueldo"))
		rphextra       = limpiaCadena(request.form("phextra"))
		rdomicilio     = limpiaCadena(request.form("domicilio"))
		rpoblacion     = limpiaCadena(request.form("poblacion"))
		rcp            = limpiaCadena(request.form("cp"))
		rprovincia     = limpiaCadena(request.form("provincia"))
		rpais          = limpiaCadena(request.form("pais"))
		rtelefono      = limpiaCadena(request.form("telefono"))
		rtelefono2     = limpiaCadena(request.form("telefono2"))
		rfax           = limpiaCadena(request.form("fax"))
		remail         = limpiaCadena(request.form("email"))
		rnivel         = limpiaCadena(request.form("nivel"))
		robservaciones = limpiaCadena(request.form("observaciones"))
		rcaja          = limpiaCadena(request.form("caja"))
		rirpf          = limpiaCadena(request.form("irpf"))
		rsegsocial     = limpiaCadena(request.form("segsocial"))
		'dgb control horario modulo Kyros  19/11/2009
        'dgm The Department of time control is always displayed.
        'if si_tiene_modulo_ControlPresencia <> 0 then
    		rdepartamento= limpiaCadena(request.Form("departamento"))
    	
    	'end if	
        rmaxamount= limpiaCadena(request.Form("maxamount"))
  	    if (si_tiene_modulo_bierzo<>0) then 
		    rtarifa        = limpiaCadena(request.form("tarifa"))
		    rporctarifa     = limpiaCadena(request.form("poctarifa"))
		    rsimayor       = limpiaCadena(request.form("simayor"))
		end if
		rfbaja	       = limpiaCadena(request.form("fbaja"))
	end if
	comercialR=enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("comercial"))))
	tecnicoR=enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("tecnico"))))
	operarioR=enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("operario"))))

	'Cargamos valores'
	if mode="first_save" or submode = "add" then
		if mode="first_save" then%>
			<input type="hidden" name="dni" value="<%=enc.EncodeForHtmlAttribute(null_s(rdni))%>">
			<input type="hidden" name="nombre" value="<%=enc.EncodeForHtmlAttribute(null_s(rnombre))%>">
			<input type="hidden" name="alias" value="<%=enc.EncodeForHtmlAttribute(null_s(ralias))%>">
			<input type="hidden" name="codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rcodigo))%>">
			<input type="hidden" name="tipo" value="<%=enc.EncodeForHtmlAttribute(null_s(rtipo))%>">
			<input type="hidden" name="antiguedad" value="<%=enc.EncodeForHtmlAttribute(null_s(rantiguedad))%>">
			<input type="hidden" name="departamento" value="<%=enc.EncodeForHtmlAttribute(null_s(rdepartamento))%>">
			<input type="hidden" name="horaIniMa" value="<%=enc.EncodeForHtmlAttribute(null_s(rhorainima))%>">
			<input type="hidden" name="horaFinMa" value="<%=enc.EncodeForHtmlAttribute(null_s(rhorafinma))%>">
			<input type="hidden" name="horaIniTa" value="<%=enc.EncodeForHtmlAttribute(null_s(rhorainita))%>">
			<input type="hidden" name="horafinTa" value="<%=enc.EncodeForHtmlAttribute(null_s(rhorafinta))%>">
			<input type="hidden" name="ss" value="<%=enc.EncodeForHtmlAttribute(null_s(rss))%>">
			<input type="hidden" name="jornada" value="<%=enc.EncodeForHtmlAttribute(null_s(rjornada))%>">
			<input type="hidden" name="sueldo" value="<%=enc.EncodeForHtmlAttribute(null_s(rsueldo))%>">
			<input type="hidden" name="phextra" value="<%=enc.EncodeForHtmlAttribute(null_s(rphextra))%>">
			<input type="hidden" name="domicilio" value="<%=enc.EncodeForHtmlAttribute(null_s(rdomicilio))%>">
			<input type="hidden" name="irpf" value="<%=enc.EncodeForHtmlAttribute(null_s(rirpf))%>">
			<input type="hidden" name="segsocial" value="<%=enc.EncodeForHtmlAttribute(null_s(rsegsocial))%>">
			<%if (si_tiene_modulo_bierzo<>0) then%>
			<input type="hidden" name="tarifa" value="<%=enc.EncodeForHtmlAttribute(null_s(rtarifa))%>">
			<input type="hidden" name="poctarifa" value="<%=enc.EncodeForHtmlAttribute(null_s(rporctarifa))%>">
			<input type="hidden" name="rsimayor" value="<%=enc.EncodeForHtmlAttribute(null_s(rsimayor))%>">
			<%end if%>		
			<input type="hidden" name="poblacion" value="<%=enc.EncodeForHtmlAttribute(null_s(rpoblacion))%>">
			<input type="hidden" name="cp" value="<%=enc.EncodeForHtmlAttribute(null_s(rcp))%>">
			<input type="hidden" name="provincia" value="<%=enc.EncodeForHtmlAttribute(null_s(rprovincia))%>">
			<input type="hidden" name="pais" value="<%=enc.EncodeForHtmlAttribute(null_s(rpais))%>">
		    <input type="hidden" name="telefono" value="<%=enc.EncodeForHtmlAttribute(null_s(rtelefono))%>">
		    <input type="hidden" name="telefono2" value="<%=enc.EncodeForHtmlAttribute(null_s(rtelefono2))%>">
		    <input type="hidden" name="fax" value="<%=enc.EncodeForHtmlAttribute(null_s(rfax))%>">
		    <input type="hidden" name="rmail" value="<%=enc.EncodeForHtmlAttribute(null_s(remail))%>">
		    <input type="hidden" name="nivel" value="<%=enc.EncodeForHtmlAttribute(null_s(rnivel))%>">
		    <input type="hidden" name="observaciones" value="<%=enc.EncodeForHtmlAttribute(null_s(robservaciones))%>">
		    <input type="hidden" name="caja" value="<%=enc.EncodeForHtmlAttribute(null_s(rcaja))%>">
		    <input type="hidden" name="fbaja" value="<%=enc.EncodeForHtmlAttribute(null_s(rfbaja))%>">
	   <%end if
	end if

    strselect = "select ndecimales from divisas with(Nolock) where moneda_base<>0 and codigo like ?+'%'"
    ndecimales = null_z(DLookupP1(strselect,session("ncliente")&"",adVarChar,15,session("dsn_cliente")))

    'Actualiza la variable domicilio si esta no fue pasada por parametro'
    if p_domicilio ="" then
        strselect = "select domicilio from domicilios with(Nolock) where pertenece=? and tipo_domicilio='PERSONAL'"
        p_domicilio=DLookupP1(strselect,p_dni&"",adVarChar,55,session("dsn_cliente"))
    end if

    'Da de alta/baja un comercial si así es requerido'
    if comercialR>"" then%>
        <input type="hidden" name="hcomercial" value="<%=enc.EncodeForHtmlAttribute(null_s(p_dni))%>">
        <%if comercialR="alta" then
            AltaComercial(p_dni)
        elseif comercialR="baja" then
            BajaComercial(p_dni)
	    end if
   end if

   'Da de alta/baja un técnico si así es requerido'
   if tecnicoR>"" then
      %><input type="hidden" name="htecnico" value="<%=enc.EncodeForHtmlAttribute(null_s(p_dni))%>"><%
	  if tecnicoR="alta" then
          AltaTecnico(p_dni)
      elseif tecnicoR="baja" then
          BajaTecnico(p_dni)
	  end if
   end if

   '** COD JCI-14022003-01 **'
   'Da de alta/baja un operario si así es requerido'
   if operarioR>"" then
      %><input type="hidden" name="hoperario" value="<%=enc.EncodeForHtmlAttribute(null_s(p_dni))%>"><%
	  if operarioR="alta" then
          AltaOperario(p_dni)
      elseif operarioR="baja" then
          BajaOperario(p_dni)
	  end if
   end if
    '** COD JCI-14022003-01 **'

   'Comprobacion de si la persona es un comercial y/o técnico y/o operario'
  	'*** VGR 28/03/03 : He añadido ltrim(rtrim()) para evitar posibles errores en la comprobación de campos.'
    strselect = "select comercial from comerciales with(Nolock) where comercial=?"
  	comercial = rtrim(ltrim(DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente")) & ""))
    strselect = "select dni from tecnicos with(Nolock) where dni=?"
	tecnico = rtrim(ltrim(DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente")) & ""))
    strselect = "select operario from operarios with(Nolock) where operario=?"
	operario = rtrim(ltrim(DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente")) & "")) %>
	<input type="hidden" name="hcomercial" value="<%=enc.EncodeForHtmlAttribute(null_s(comercial))%>">
	<input type="hidden" name="htecnico" value="<%=enc.EncodeForHtmlAttribute(null_s(tecnico))%>">
	<input type="hidden" name="hoperario" value="<%=enc.EncodeForHtmlAttribute(null_s(operario))%>">
	<%
	'Comprobación de si la persona está dada de alta o baja como comercial, técnico y operario'
	if comercial>"" then
        strselect = "select fbaja from comerciales with(Nolock) where comercial=?"
        fbajacomercial=DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente"))
		if fbajacomercial>"" then
			%><input type="hidden" name="hcomerc" value=""><%
		else
			%><input type="hidden" name="hcomerc" value="<%=enc.EncodeForHtmlAttribute(null_s(fbajacomercial))%>"><%
		end if
	else
		%><input type="hidden" name="hcomerc" value=""><%
	end if
	if tecnico>"" then
        strselect = "select fbaja from tecnicos with(Nolock) where dni=?"
        fbajatecnico=DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente"))
		if fbajatecnico>"" then
			%><input type="hidden" name="htecnic" value=""><%
		else
			%><input type="hidden" name="htecnic" value="<%=enc.EncodeForHtmlAttribute(null_s(fbajatecnico))%>"><%
		end if
	else
		%><input type="hidden" name="htecnic" value=""><%
	end if
	if operario>"" then
        strselect = "select fbaja from operarios with(Nolock) where operario=?"
        fbajaoperario=DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente"))
		if fbajaoperario>"" then
			%><input type="hidden" name="hoper" value=""><%
		else
			%><input type="hidden" name="hoper" value="<%=enc.EncodeForHtmlAttribute(null_s(fbajaoperario))%>"><%
		end if
	else
		%><input type="hidden" name="hoper" value=""><%
	end if

    'Acción a realizar'
    if mode="save" or mode="first_save" then
		if mode="first_save" then
			p_dni = session("ncliente") & p_dni
            strselect = "select * from personal where dni=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@dni",adVarChar,adParamInput,20,p_dni&"")
            set rstAux = command2.Execute
			if not rstAux.eof then
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
				mode="add"
				p_dni=""
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgDniExiste%>");
				      document.personal.action="personal.asp?mode=add";
				      document.personal.submit();
				      parent.botones.document.location="personal_bt.asp?mode=add";
	   			</script><%
			else
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
				GuardarRegistro p_dni
				mode="browse"
			end if
		else
			p_dni = session("ncliente") & p_dni
			GuardarRegistro p_dni
			mode="browse"
		end if

	elseif mode="delete" then
	    EliminarRegistro
		p_dni = ""
		p_domicilio = ""
		mode="add"
		%><script language="javascript" type="text/javascript">
		      parent.botones.document.location = "personal_bt.asp?mode=add";
		      SearchPage("personal_lsearch.asp?mode=init", 0);
		</script><%
    end if

	if mode="edit" or mode="browse" then
        set rst = Server.CreateObject("ADODB.Recordset")

        rst.cursorlocation=3
        strselect = "select * from personal with(NOLOCK) where dni=?"
        set command = nothing
        set conn = Server.CreateObject("ADODB.Connection")
        set command = Server.CreateObject("ADODB.Command")
        conn.Open = session("dsn_cliente")
        conn.CursorLocation = 3
        command.ActiveConnection = conn
        command.CommandTimeout = 60
        command.CommandText = strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@dni",adVarChar,adParamInput,20,p_dni&"")
        set rst = command.Execute
		if rst.eof then
			personal_no_existe=1
			p_dni=""
			mode="add"
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitMsgRegsNoExiste%>");
			      parent.botones.document.location = "personal_bt.asp?mode=add";
			      SearchPage("personal_lsearch.asp?mode=init", 0);
			</script><%
		else
			personal_no_existe=0
		end if
        conn.Close
        set conn = nothing
        set command = nothing
	end if

   strselect = "select nombre from personal with(Nolock) where dni=?"
   p_nombre = DLookupP1(strselect,p_dni&"",adVarChar,20,session("dsn_cliente"))

   if mode="edit" or mode="browse" then
   	  rnivel=""
      rst.cursorlocation=3
      strselect = " select * from personal with(nolock) where dni=?"
      set command = nothing
      set conn = Server.CreateObject("ADODB.Connection")
      set command = Server.CreateObject("ADODB.Command")
      conn.Open = session("dsn_cliente")
      conn.CursorLocation = 3
      command.ActiveConnection = conn
      command.CommandTimeout = 60
      command.CommandText = strselect
      command.CommandType = adCmdText
      command.Parameters.Append command.CreateParameter("@dni",adVarChar,adParamInput,20,p_dni&"")
      set rst = command.Execute

      rstDom.cursorlocation=3
      strselect = "select * from domicilios with(nolock) where pertenece=? and tipo_domicilio='PERSONAL'"
      set commandDom = nothing
      set connDom = Server.CreateObject("ADODB.Connection")
      set commandDom = Server.CreateObject("ADODB.Command")
      connDom.Open = session("dsn_cliente")
      connDom.CursorLocation = 3
      commandDom.ActiveConnection = connDom
      commandDom.CommandTimeout = 60
      commandDom.CommandText = strselect
      commandDom.CommandType = adCmdText
      commandDom.Parameters.Append commandDom.CreateParameter("@pertenece",adVarChar,adParamInput,55,p_dni&"")
      set rstDom = commandDom.Execute
      if not rst.eof then
         rdni        = rst("dni")
	     rantiguedad   = rst("antiguedad")
	     rhorainima    = rst("hora_ini_ma")
	     rhorafinma    = rst("hora_fin_ma")
	     rhorainita    = rst("hora_ini_ta")
	     rhorafinta    = rst("hora_fin_ta")
   	     robservaciones = rst("observaciones")
		 rsueldo        = null_z(rst("sueldo"))
		 rjornada       = null_z(rst("jornada"))
		 rss            = rst("ss")
		 rnombre        = rst("nombre")
		 ralias			 = rst("alias")
		 rcodigo        = rst("codigo")
		 rphextra       = rst("phextra")
		 rnivel			= "" & rst("nivel")
		 remail         = rst("email")
		 rcaja          = rst("caja")
		 rtelefono2     = rst("telefono2")
	     rfax           = rst("fax")
	     rtipo			= rst("tipo")

	     if si_tiene_paginaSMS<>0 then
		 	rsms           = rst("mensajeria_sms")
		 end if
		 rfbaja			= rst("fbaja")
		 rirpf          = rst("irpf")
		 rsegsocial     = rst("importe_ss")
         rmaxamount     = rst("maxamount")
		 
		 if (si_tiene_modulo_bierzo<>0) then  
		    rtarifa        = rst("campo01")
		    rporctarifa     = rst("campo02")
		    rsimayor       = rst("campo03")
		 end if
		 
		 'dgb control horario modulo Kyros  19/11/2009
         'dgm The Department of time control is always displayed.
    	 if si_tiene_modulo_ControlPresencia <> 0 then
    	    rdepartamento= rst("fdepartamento")
    	 else
            rdepartamento= rst("department")
         end if
      end if
      conn.Close
      set conn = nothing
      set command = nothing

      if not rstdom.eof then
         rdomicilio = rstDom("domicilio")
    	 rpoblacion = rstDom("poblacion")
    	 rcp        = rstDom("cp")
	     rprovincia = rstDom("provincia")
	     rpais      = rstDom("pais")
	     rtelefono  = rstDom("telefono")
      end if
      connDom.Close
      set connDom = nothing
      set commandDom = nothing

	  if comercial>"" then
         rstAux.cursorlocation=3
         strselect = "select * from comerciales with(nolock) where comercial=?"
         set commandAux = nothing
         set connAux = Server.CreateObject("ADODB.Connection")
         set commandAux = Server.CreateObject("ADODB.Command")
         connAux.Open = session("dsn_cliente")
         connAux.CursorLocation = 3
         commandAux.ActiveConnection = connAux
         commandAux.CommandTimeout = 60
         commandAux.CommandText = strselect
         commandAux.CommandType = adCmdText
         commandAux.Parameters.Append commandAux.CreateParameter("@comercial",adVarChar,adParamInput,20,p_dni&"")
         set rstAux = commandAux.Execute
		 if not rstAux.eof then
		    rcventas = rstAux("cventas")
			rmganancia = rstAux("mganancia")
			robjetivo = rstAux("objetivo")
			rper_ob = rstAux("per_ob")
			rcbase = rstAux("combase")
			rcconcepto = rstAux("comconceptos")
			rcpena = rstAux("penalizacion")
			rsuperior = rstAux("superior")
			rcfbaja	= rstAux("fbaja")
			if rcfbaja & "">"" then
				existe_comercial="0"
			else
				existe_comercial="1"
			end if
		else
			existe_comercial="0"
		 end if
          connAux.Close
          set connAux = nothing
          set commandAux = nothing
	  else
	  	rcfbaja = " "
		existe_comercial="0"
	  end if

		%><input type="hidden" name="existe_comercial" value="<%=existe_comercial%>"><%

	  if tecnico>"" then
         rstAux.cursorlocation=3
         strselect = "select tecnicos.*,descripcion,marca,modelo from tecnicos with(nolock) left outer join almacenes with(nolock) on almacen=codigo left outer join vehiculos with(nolock) on matricula=vehiculo where dni=?"
         set commandAux = nothing
         set connAux = Server.CreateObject("ADODB.Connection")
         set commandAux = Server.CreateObject("ADODB.Command")
         connAux.Open = session("dsn_cliente")
         connAux.CursorLocation = 3
         commandAux.ActiveConnection = connAux
         commandAux.CommandTimeout = 60
         commandAux.CommandText = strselect
         commandAux.CommandType = adCmdText
         commandAux.Parameters.Append commandAux.CreateParameter("@dni",adVarChar,adParamInput,20,p_dni&"")
         set rstAux = commandAux.Execute
		 if not rstAux.eof then
		    Tcomision=rstAux("comision")
		 	Tphextralab=rstAux("phextralab")
		 	Tphextrafes=rstAux("phextrafes")
		 	Tphlaboral=rstAux("phlaboral")
		 	Tincentivo1=rstAux("incentivo1")
		 	Tincentivo2=rstAux("incentivo2")
		 	Talmacen=rstAux("descripcion")
		 	TcodAlmacen=rstAux("almacen")
		 	Tmatricula=rstAux("vehiculo")
		 	Tmarca=rstAux("marca")
		 	Tmodelo=rstAux("modelo")
		 	Tvehiculo=rstAux("vehiculo") & " " & rstAux("marca") & " " & rstAux("modelo")
		 	Tfbaja=rstAux("fbaja")
		 	if Tfbaja & "">"" then
		 		existe_tecnico="0"
		 	else
		 		existe_tecnico="1"
		 	end if
		 else
		 	existe_tecnico="0"
		 end if
         connAux.Close
         set connAux = nothing
         set commandAux = nothing
	  else
	  	Tfbaja=" "
		existe_tecnico="0"
	  end if

		%><input type="hidden" name="existe_tecnico" value="<%=existe_tecnico%>"><%

	  '** COD JCI-14022003-01 **
	  if operario<>"" then
         rstAux.cursorlocation=3
         strselect = "select coste_hora,fbaja from operarios with(nolock) where operario=?"
         set commandAux = nothing
         set connAux = Server.CreateObject("ADODB.Connection")
         set commandAux = Server.CreateObject("ADODB.Command")
         connAux.Open = session("dsn_cliente")
         connAux.CursorLocation = 3
         commandAux.ActiveConnection = connAux
         commandAux.CommandTimeout = 60
         commandAux.CommandText = strselect
         commandAux.CommandType = adCmdText
         commandAux.Parameters.Append commandAux.CreateParameter("@operario",adVarChar,adParamInput,20,p_dni&"")
         set rstAux = commandAux.Execute
		 if not rstAux.eof then
		 	Ocoste_hora=rstAux("coste_hora")
		 	Ofbaja=rstAux("fbaja")
			if Ofbaja & "">"" then
				existe_operario="0"
			else
				existe_operario="1"
			end if
		else
			existe_operario="0"
		 end if
         connAux.Close
         set connAux = nothing
         set commandAux = nothing
	  else
	    Ofbaja=" "
		existe_operario="0"
	  end if
	  '** FIN COD JCI-14022003-01 **

		%><input type="hidden" name="existe_operario" value="<%=existe_operario%>"><%

   end if
   PintarCabecera "personal.asp"
        %><div class="headers-wrapper"><%
            DrawDiv "header-center","",""
            DrawLabel "","",Litdni%><span class="CELDA"><%=trimCodEmpresa(p_dni)%></span><%CloseDiv%>
            <!--<td width="1%"><span><img src="<%=ImgLineVertical %>" id="img" style="vertical-align:top;" class="line_vertical" /></span></td>-->
            <%DrawDiv "header-name","",""
              DrawLabel "","",LitNombre
              if mode="browse" then%> 
                    <span class="CELDA"><%=enc.EncodeForHtmlAttribute(null_s(p_nombre))%></span><%
                else
                    %><span class="CELDA"><%=enc.EncodeForHtmlAttribute(null_s(p_nombre))%></span><%
              end if%>

            <%CloseDiv%>

            <%if mode="browse" then 
             DrawDiv "col-lg-3 col-md-6 col-sm-6 col-xs-12","",""%>
                <a class="CELDAREFB" href="javascript:AbrirVentana('../central.asp?pag1=Servicios/recursos.asp&pag2=Servicios/recursos_bt.asp&codigo=<%=enc.EncodeForJavascript(p_dni)%>&tipo=personal&viene=enlaces', 'P', <%=AltoVentana%>, <%=AnchoVentana%>)">&nbsp;<%=LitRecurso %></a>
             <%CloseDiv
            end if%>
        </div>
    <table style="width: 100%;"></table>
   <%alarma "personal.asp"

   	'---------------------------------
   	'Modo de inserción
   	'---------------------------------
	if mode = "add" then
		 %>
   		<% 'DATOS GENERALES MODO AÑADIR
	' Inicio Borde Span
	%>
    <div id="CollapseSection"> 
        <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['addDG']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
        <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['addDG']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
    </div>

    <!--<table class=TBORDE width="100%"><tr><td>-->
        <div class="Section" id="S_addDG">
            <a href="#" rel="toggle[addDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitDatosGenerales%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display: ;" id="addDG">
		    <input type="hidden" name="hdni" value="<%=enc.EncodeForHtmlAttribute(null_s(rdni))%>"> 
     		<table  bgcolor='<%=color_blau%>' border="0" cellpadding="2" cellspacing="2"><%


                DrawInputCeldaLabel "CELDA' maxlength='15","txtMandatory",15,Litdni,"dni",enc.EncodeForHtmlAttribute(null_s(rdni))

                DrawInputCeldaLabel "CELDA","txtMandatory",30,LitNombre,"nombre",enc.EncodeForHtmlAttribute(null_s(rnombre))

                EligeCelda "input",mode,"CELDA' maxlength='50'","","",0,LitAlias,"alias",15,enc.EncodeForHtmlAttribute(null_s(ralias))

                EligeCelda "input",mode,"CELDA' maxlength='15'","","",0,LitCodigoOp,"codigo",15,enc.EncodeForHtmlAttribute(null_s(rcodigo))

				rstAux.cursorlocation=3
                strselect = "select codigo,descripcion from tipos_entidades with(nolock) where tipo='PERSONAL' and codigo like ?+'%' order by descripcion"
                set command = nothing
                set conn = Server.CreateObject("ADODB.Connection")
                set command = Server.CreateObject("ADODB.Command")
                conn.Open = session("dsn_cliente")
                conn.CursorLocation = 3
                command.ActiveConnection = conn
                command.CommandTimeout = 60
                command.CommandText = strselect
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
                set rstAux = command.Execute
                DrawSelectCelda "CELDA","","",0,LitTipo,"tipo",rstAux,rtipo,"codigo","descripcion","",""
                conn.Close
                set conn = nothing
                set command = nothing

                EligeCelda "input",mode,"CELDA' maxlength='30'","","",0,LitSS,"ss",20,rss

                DrawInputCeldaActionDiv "","","",10,0,LitAntiguedad,"antiguedad",enc.EncodeForHtmlAttribute(null_s(rantiguedad)),"onblur","checkdate(this)", false
                DrawCalendar "antiguedad"

                EligeCelda "input",mode,"CELDA' maxlength='5'","","",0,Litjornada,"jornada",2,enc.EncodeForHtmlAttribute(null_s(rjornada))

    		'dgb control horario modulo Kyros  19/11/2009
            'dgm 09.07.12 The Department of time control is always displayed.

				rstAux.cursorlocation=3
                if si_tiene_modulo_ControlPresencia <> 0 then
                    strselect = "select FidDepartamento as code, Fdescripcion as descrip from ch_departamentos with(nolock) where fiddepartamento like ?+'%' order by fdescripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@fiddepartamento",adVarChar,adParamInput,8,session("ncliente")&"")
                    set rstAux = command2.Execute
                else
                    strselect = "select codigo as code, descripcion as descrip from departamentos with(nolock) where codigo like ?+'%' order by descripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,7,session("ncliente")&"")
                    set rstAux = command2.Execute
                end if
                DrawSelectCelda "CELDA","","",0,LitCHDepartamento,"departamento",rstAux,rdepartamento,"code","descrip","",""
                conn2.Close
                set conn2 = nothing
                set command2 = nothing

                 EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitMaxAmount,"maxamount",10,enc.EncodeForHtmlAttribute(null_s(rmaxamount))
            if si_tiene_modulo_ControlPresencia <> 0 then%>
    			<input type="hidden" id="horaIniMa" value="" />
    			<input type="hidden" id="horaFinMa" value="" />
    			<input type="hidden" id="horaIniTa" value="" />
    			<input type="hidden" id="horaFinTa" value="" /><%
            else
                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraIniMa,"horaIniMa",enc.EncodeForHtmlAttribute(null_s(rhorainima)),"","", false

                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraFinMa,"horaFinMa",enc.EncodeForHtmlAttribute(null_s(rhorafinma)),"","", false

                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraIniTa,"horaIniTa",enc.EncodeForHtmlAttribute(null_s(rhorainita)),"","", false

                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraFinTa,"horaFinTa",enc.EncodeForHtmlAttribute(null_s(rhorafinta)),"","", false
    			%>
                <script language="javascript" type="text/javascript">
                    function horaIniMa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaIniMa, false, ev);
                    }

                    if (window.document.personal.horaIniMa.addEventListener) {
                        window.document.personal.horaIniMa.addEventListener("keyup", horaIniMa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaIniMa.attachEvent("onkeyup", horaIniMa_callkeyuphandler);
                    }

                    function horaFinMa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaFinMa, false, ev);
                    }

                    if (window.document.personal.horaFinMa.addEventListener) {
                        window.document.personal.horaFinMa.addEventListener("keyup", horaFinMa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaFinMa.attachEvent("onkeyup", horaFinMa_callkeyuphandler);
                    }

                    function horaIniTa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaIniTa, false, ev);
                    }

                    if (window.document.personal.horaIniTa.addEventListener) {
                        window.document.personal.horaIniTa.addEventListener("keyup", horaIniTa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaIniMa.attachEvent("onkeyup", horaIniTa_callkeyuphandler);
                    }

                    function horaFinTa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaFinTa, false, ev);
                    }

                    if (window.document.personal.horaFinTa.addEventListener) {
                        window.document.personal.horaFinTa.addEventListener("keyup", horaFinTa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaFinTa.attachEvent("onkeyup", horaFinTa_callkeyuphandler);
                    }
	            </script>
    	    <%end if

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitSueldo,"sueldo",10,enc.EncodeForHtmlAttribute(null_s(rsueldo))

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitPhextra,"phextra",6,enc.EncodeForHtmlAttribute(null_s(rphextra))

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitIRPF,"irpf",10,enc.EncodeForHtmlAttribute(null_s(rirpf))

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitSegSocial,"segsocial",10,enc.EncodeForHtmlAttribute(null_s(rsegsocial))

     	        if (si_tiene_modulo_bierzo<>0) then   			    
                    DrawDiv "1","",""
                    DrawLabel "","",LitTarifa%><input class="CELDA" type="text" name="porctarifa" size="5" value="<%=enc.EncodeForHtmlAttribute(null_s(rporctarifa))%>"/><span class="CELDA"><%=LitPorcentaje%></span><%CloseDiv
                    DrawDiv "1","",""
                    DrawLabel "","",LitTarifaSi%><select class='width10' name="tarifa">
				        <%if rtarifa="" then 
				            defecto="selected"
				          else 
				            defecto=""
				          end if %>
				    	<option <%=defecto%> value=""></option>
				        <%if rtarifa=BB then 
				            defecto="selected"
				          else 
				            defecto=""
				          end if %>
				    	<option <%=defecto%> value="<%=enc.EncodeForHtmlAttribute(null_s(BB))%>"><%=BB%></option>
				        <%if rtarifa=BL then 
				            defecto="selected"
				          else 
				            defecto=""
				          end if %>
				    	<option <%=defecto%> value="<%=enc.EncodeForHtmlAttribute(null_s(BL))%>"><%=BL%></option>
				    </select><label class="width5" style="text-align:center; display: inherit;"><%=LitTarifaMayor%></label><input type="text" size=5 class="width10" name="simayor" value="<%=enc.EncodeForHtmlAttribute(null_s(rsimayor))%>"/><%CloseDiv
    			end if
                    DrawInputCeldaLabel "CELDA","txtMandatory",35,LitDireccion,"domicilio",enc.EncodeForHtmlAttribute(null_s(rdomicilio))

                    DrawDiv "1","",""
                    DrawLabel "","",LitPoblacion%><input class="CELDA" type="text" maxlength='50' size=25 name="poblacion" value="<%=enc.EncodeForHtmlAttribute(null_s(rpoblacion))%>"><a class='CELDAREFB'  href="javascript:AbrirVentana('../configuracion/poblaciones.asp?mode=buscar&viene=personal&titulo=<%=LitSelPoblacion%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPoblaciones%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"></a><%CloseDiv

                    EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitCP,"cp",5,enc.EncodeForHtmlAttribute(null_s(rcp))

                    EligeCelda "input",mode,"CELDA' maxlength='50'","","",0,LitProvincia,"provincia",25,enc.EncodeForHtmlAttribute(null_s(rprovincia))

                    EligeCelda "input",mode,"CELDA' maxlength='30'","","",0,LitPais,"pais",30,enc.EncodeForHtmlAttribute(null_s(rpais))

                    EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitTel1,"telefono",20,enc.EncodeForHtmlAttribute(null_s(rtelefono))

                    EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitTel2,"telefono2",20,enc.EncodeForHtmlAttribute(null_s(rtelefono2))

                    EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitFax,"fax",20,enc.EncodeForHtmlAttribute(null_s(rfax))

                    EligeCelda "input",mode,"left","","",0,LitEmail,"email",30,enc.EncodeForHtmlAttribute(null_s(remail))
				
                    if mode="add" then rnivel=0

                    DrawDiv "1","",""
                    DrawLabel "","",LitNivel%><select class='CELDA' name="nivel">
					<%for niv=0 to 5
						if cint(niv)=cint(rnivel) then%>
							<option selected value="<%=enc.EncodeForHtmlAttribute(null_s(niv))%>"><%=niv%></option>
						<%else%>
							<option value="<%=enc.EncodeForHtmlAttribute(null_s(niv))%>"><%=niv%></option>
						<%end if
					next%></select><%CloseDiv
                    EligeCelda "input",mode,"left","","",0,LitFBaja,"fbaja",10,rfbaja
                    DrawCalendar "fbaja"
                    strselect = "select codigo,descripcion from cajas with(nolock) where codigo like ?+'%' order by descripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
                    set rstAux = command2.Execute

                    DrawSelectCelda "CELDA style='width:200px' ","","",0,LitCaja,"caja",rstAux,rcaja,"codigo","descripcion","",""
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing

					if si_tiene_paginaSMS<>0 then
                        EligeCelda "check",mode,"left","","",0,LitMensajeriaSMS,"mensajeria_sms",0,""
					end if

                    EligeCelda "text",mode,"left","","",0,LitObservaciones,"observaciones",2,enc.EncodeForHtmlAttribute(null_s(robservaciones))
			%></table>
		<!--</center>-->
        </div>
        </div>
		<!--</td></tr></table>-->
	<%
    BarraNavegacion mode
    
    elseif mode="edit" then
		
        BarraOpciones mode,rdni,rnombre,si_tiene_modulo_comercial
		'DATOS GENERALES MODO EDIT
		' Inicio Borde Span
		%>
    <div id="CollapseSection"> 
        <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['editDG','editDC','editDT','editDO']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
        <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['editDG','editDC','editDT','editDO']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
    </div>

        <!--<table class=TBORDE width="100%"><tr><td>-->
        <div class="Section" id="S_editDG">
            <a href="#" rel="toggle[editDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitDatosGenerales%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display: ;" id="editDG">
	     <!--<br />-->
         <!--<center>-->
		    <input type="hidden" name="hdni" value="<%=enc.EncodeForHtmlAttribute(null_s(rdni))%>">
			<input type="hidden" name="hdomicilio" value="<%=enc.EncodeForHtmlAttribute(null_s(rdomicilio))%>">
			<input type="hidden" name="hcodigoop" value="<%=enc.EncodeForHtmlAttribute(null_s(rcodigo))%>">
     			<table  bgcolor='<%=color_blau%>' border='0' cellpadding=2 cellspacing=2> <%
		      	%><td colspan="4"><%
		       		'BarraOpcionesGen mode,rdni,rnombre
		       	%></td>
                </table><%
                DrawInputCeldaLabel "CELDA' maxlength='15","txtMandatory",15,Litdni,"dni",trimCodEmpresa(rdni)

                DrawInputCeldaLabel "CELDA","txtMandatory",30,LitNombre,"nombre",enc.EncodeForHtmlAttribute(null_s(rnombre))

                EligeCelda "input",mode,"CELDA' maxlength='50'","","",0,LitAlias,"alias",15,enc.EncodeForHtmlAttribute(null_s(ralias))

                EligeCelda "input",mode,"CELDA' maxlength='15'","","",0,LitCodigoOp,"codigo",15,enc.EncodeForHtmlAttribute(null_s(rcodigo))

				rstAux.cursorlocation=3
                strselect = "select codigo,descripcion from tipos_entidades with(nolock) where tipo='PERSONAL' and codigo like ?+'%' order by descripcion"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
                set rstAux = command2.Execute
                DrawSelectCelda "CELDA","","",0,LitTipo,"tipo",rstAux,rtipo,"codigo","descripcion","",""
                conn2.Close
                set conn2 = nothing
                set command2 = nothing

                EligeCelda "input",mode,"CELDA' maxlength='30'","","",0,LitSS,"ss",20,rss

                DrawInputCeldaActionDiv "","","",10,0,LitAntiguedad,"antiguedad",enc.EncodeForHtmlAttribute(null_s(rantiguedad)),"onblur","checkdate(this)", false
                DrawCalendar "antiguedad"

                EligeCelda "input",mode,"CELDA' maxlength='5'","","",0,Litjornada,"jornada",2,enc.EncodeForHtmlAttribute(null_s(rjornada))

    		'dgb control horario modulo Kyros  19/11/2009
            'dgm 09.07.12 The Department of time control is always displayed.

				rstAux.cursorlocation=3
                if si_tiene_modulo_ControlPresencia <> 0 then
                    strselect = "select FidDepartamento as code,Fdescripcion as descrip from ch_departamentos with(nolock) where FidDepartamento like ?+'%' order by Fdescripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@FidDepartamento",adVarChar,adParamInput,8,session("ncliente")&"")
                    set rstAux = command2.Execute
                else
                    strselect = "select codigo as code,descripcion as descrip from departamentos with(nolock) where codigo like ?+'%' order by descripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,7,session("ncliente")&"")
                    set rstAux = command2.Execute
                end if
                DrawSelectCelda "CELDA","","",0,LitCHDepartamento,"departamento",rstAux,rdepartamento,"code","descrip","",""
                conn2.Close
                set conn2 = nothing
                set command2 = nothing

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitMaxAmount,"maxamount",10,rmaxamount
    		if si_tiene_modulo_ControlPresencia <> 0 then%>
    			<input type="hidden" id="horaIniMa" value="" />
    			<input type="hidden" id="horaFinMa" value="" />
    			<input type="hidden" id="horaIniTa" value="" />
    			<input type="hidden" id="horaFinTa" value="" />
   			<%else
                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraIniMa,"horaIniMa",enc.EncodeForHtmlAttribute(null_s(rhorainima)),"","", false

                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraFinMa,"horaFinMa",enc.EncodeForHtmlAttribute(null_s(rhorafinma)),"","", false

                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraIniTa,"horaIniTa",enc.EncodeForHtmlAttribute(null_s(rhorainita)),"","", false

                DrawInputCeldaActionDiv "CELDA' maxlength='5'","","",5,0,LitHoraFinTa,"horaFinTa",enc.EncodeForHtmlAttribute(null_s(rhorafinta)),"","", false%><script language="javascript" type="text/javascript">
                    function horaIniMa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaIniMa, false, ev);
                    }

                    if (window.document.personal.horaIniMa.addEventListener) {
                        window.document.personal.horaIniMa.addEventListener("keyup", horaIniMa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaIniMa.attachEvent("onkeyup", horaIniMa_callkeyuphandler);
                    }

                    function horaFinMa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaFinMa, false, ev);
                    }

                    if (window.document.personal.horaFinMa.addEventListener) {
                        window.document.personal.horaFinMa.addEventListener("keyup", horaFinMa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaFinMa.attachEvent("onkeyup", horaFinMa_callkeyuphandler);
                    }

                    function horaIniTa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaIniTa, false, ev);
                    }

                    if (window.document.personal.horaIniTa.addEventListener) {
                        window.document.personal.horaIniTa.addEventListener("keyup", horaIniTa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaIniMa.attachEvent("onkeyup", horaIniTa_callkeyuphandler);
                    }

                    function horaFinTa_callkeyuphandler(evnt) {
                        ev = (evnt) ? evnt : event;
                        formatHora(document.personal.horaFinTa, false, ev);
                    }

                    if (window.document.personal.horaFinTa.addEventListener) {
                        window.document.personal.horaFinTa.addEventListener("keyup", horaFinTa_callkeyuphandler, false);
                    }
                    else {
                        window.document.personal.horaFinTa.attachEvent("onkeyup", horaFinTa_callkeyuphandler);
                    }
	            </script><%end if
		   		EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitSueldo,"sueldo",10,enc.EncodeForHtmlAttribute(null_s(rsueldo))

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitPhextra,"phextra",6,enc.EncodeForHtmlAttribute(null_s(rphextra))

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitIRPF,"irpf",10,enc.EncodeForHtmlAttribute(null_s(rirpf))

                EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitSegSocial,"segsocial",10,enc.EncodeForHtmlAttribute(null_s(rsegsocial))
     			if (si_tiene_modulo_bierzo<>0) then
					DrawDiv "1","",""
                    DrawLabel "","",LitTarifa%><input  class="CELDA" type="text" name="porctarifa" size=5 value="<%=enc.EncodeForHtmlAttribute(null_s(rporctarifa))%>"/><span class="CELDA"><%=LitPorcentaje%></span><%CloseDiv
                    DrawDiv "1","",""
                    DrawLabel "","",LitTarifaSi%><select class='width10' name="tarifa">
					    <%if rtarifa="" then 
					        defecto="selected"
					      else 
					        defecto=""
					      end if %>
						<option <%=defecto%> value=""></option>
					    <%if rtarifa=BB then 
					        defecto="selected"
					      else 
					        defecto=""
					      end if %>
						<option <%=defecto%> value="<%=enc.EncodeForHtmlAttribute(null_s(BB))%>"><%=BB%></option>
					    <%if rtarifa=BL then 
					        defecto="selected"
					      else 
					        defecto=""
					      end if %>
						<option <%=defecto%> value="<%=enc.EncodeForHtmlAttribute(null_s(BL))%>"><%=BL%></option>
					</select><label class="width5" style="text-align:center;display: inherit;"><%=LitTarifaMayor%></label><input type="text"  size=5 class="width10" name="simayor" value="<%=enc.EncodeForHtmlAttribute(null_s(rsimayor))%>"/><%CloseDiv
    			end if    			
                    DrawInputCeldaLabel "CELDA","txtMandatory",35,LitDireccion,"domicilio",rdomicilio

                    DrawDiv "1","",""
                    DrawLabel "","",LitPoblacion%><input class="CELDA" type="text" maxlength='50' size=25 name="poblacion" value="<%=enc.EncodeForHtmlAttribute(null_s(rpoblacion))%>"><a class='CELDAREFB'  href="javascript:AbrirVentana('../configuracion/poblaciones.asp?mode=buscar&viene=personal&titulo=<%=LitSelPoblacion%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPoblaciones%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"></a><%CloseDiv

                    EligeCelda "input",mode,"CELDA' maxlength='5'","","",0,LitCP,"cp",5,enc.EncodeForHtmlAttribute(null_s(rcp))

                    EligeCelda "input",mode,"CELDA' maxlength='50'","","",0,LitProvincia,"provincia",25,enc.EncodeForHtmlAttribute(null_s(rprovincia))

                    EligeCelda "input",mode,"CELDA","","",0,LitPais,"pais",30,enc.EncodeForHtmlAttribute(null_s(rpais))

                    EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitTel1,"telefono",20,enc.EncodeForHtmlAttribute(null_s(rtelefono))

                    EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitTel2,"telefono2",20,enc.EncodeForHtmlAttribute(null_s(rtelefono2))

                    EligeCelda "input",mode,"CELDA' maxlength='10'","","",0,LitFax,"fax",20,enc.EncodeForHtmlAttribute(null_s(rfax))

                    EligeCelda "input",mode,"CELDA","","",0,LitEmail,"email",30,enc.EncodeForHtmlAttribute(null_s(remail))
					if mode="add" then rnivel=0
                    DrawDiv "1","",""
                    DrawLabel "","",LitNivel%><select class='CELDA' name="nivel">
					<%for niv=0 to 5
						if cint(niv)=cint(rnivel) then%>
							<option selected value="<%=enc.EncodeForHtmlAttribute(null_s(niv))%>"><%=niv%></option>
						<%else%>
							<option value="<%=enc.EncodeForHtmlAttribute(null_s(niv))%>"><%=niv%></option>
						<%end if
					next%>
					</select><%CloseDiv

                    EligeCelda "input",mode,"left","","",0,LitFBaja,"fbaja",10,enc.EncodeForHtmlAttribute(null_s(rfbaja))
                    DrawCalendar "fbaja"
                    strselect = "select codigo,descripcion from cajas with(nolock) where codigo like ?+'%' order by descripcion"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
                    set rstAux = command2.Execute
                    DrawSelectCelda "CELDA style='width:200px' ","","",0,LitCaja,"caja",rstAux,rcaja,"codigo","descripcion","",""
					conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
					if si_tiene_paginaSMS<>0 then
                        EligeCelda "check",mode,"left","","",0,LitMensajeriaSMS,"mensajeria_sms",10,enc.EncodeForHtmlAttribute(null_s(rsms))
					end if
                    EligeCelda "text",mode,"left","","",0,LitObservaciones,"observaciones",2,enc.EncodeForHtmlAttribute(null_s(robservaciones))
			%>
        <!--</center>-->
        </div>
        </div><%
		'DATOS COMERCIALES MODO EDIT'%>
        <div class="Section" id="S_editDC">
            <a href="#" rel="toggle[editDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=iif(si_tiene_modulo_comercial<>0,LitDatosComercialesModCom,LitDatosComerciales)%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:none;" id="editDC"><%
				''if comercial>"" then
			  if rfbaja<>"" then%>
				<input type="hidden" name="cventas" value="0">
				<input type="hidden" name="mganancia" value="0">
				<input type="hidden" name="per_ob" value="0">
				<input type="hidden" name="objetivo" value="0">
				<input type="hidden" name="cconcepto" value="0">
				<input type="hidden" name="cbase" value="0">
				<input type="hidden" name="pena" value="0">
				<input type="hidden" name="tcomision" value="0">
				<input type="hidden" name="tphextralab" value="0">
				<input type="hidden" name="tphextrafes" value="0">
				<input type="hidden" name="tphlaboral" value="0">
				<input type="hidden" name="tincentivo1" value="0">
				<input type="hidden" name="tincentivo2" value="0">
			  <%else
				if rcfbaja<>"" then
				else
		       		DrawFila color_blau
		       			%><td colspan="4"><%
		       			'BarraOpciones mode,rdni,rnombre,si_tiene_modulo_comercial
		       			%></td><%
					CloseFila
				end if

				''if comercial>"" then
				if rcfbaja<>"" then
					DrawFila color_blau
						%><td colspan=4><a class="CELDAREF7" href="javascript:Comercial('alta',<%=enc.EncodeForJavascript(null_s(si_tiene_modulo_comercial))%>);"><%=iif(si_tiene_modulo_comercial<>0,LitDarAltaComComercModCom,LitDarAltaComComerc)%></a></td><%
					CloseFila
				else%>
			<!-- Quitar cuando se descomenten los campos InputText-->
				<input type="hidden" name="cventas" value="0">
				<input type="hidden" name="mganancia" value="0">
				<input type="hidden" name="per_ob" value="0">
				<input type="hidden" name="objetivo" value="0">
				<input type="hidden" name="cconcepto" value="0"><%
					if si_tiene_modulo_comercial<>0 then
                        EligeCelda "input",mode,"left","","",0,LitPenalizacion,"pena",20,enc.EncodeForHtmlAttribute(null_s(rcpena))
					end if
                    Dim Literal
					if si_tiene_modulo_comercial<>0 then
                        Literal = LitSuperiorModCom
					else
                        Literal = LitSuperior
					end if
					rstAux.cursorlocation=3
                    strselect = "select dni,nombre from personal a with(nolock), comerciales b with(nolock) where a.dni=b.comercial and a.dni like ?+'%' order by nombre"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@dni",adVarChar,adParamInput,20,session("ncliente")&"")
                    set rstAux = command2.Execute
                    DrawSelectCelda "CELDA","","",0,Literal,"superior",rstAux,rsuperior,"dni","nombre","",""
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
				if si_tiene_modulo_comercial<>0 then
                    EligeCelda "input",mode,"left","","",0,LitComBase,"cbase",20,rcbase
				end if%>
                    <table style="width: 100%;"></table>
					<td colspan=1><a class="CELDAREF7" href="javascript:Comercial('baja',<%=enc.EncodeForJavascript(null_s(si_tiene_modulo_comercial))%>);"><%=iif(si_tiene_modulo_comercial<>0,LitDarBajaComComercModCom,LitDarBajaComComerc)%></a></td><%
				end if
			  end if
			%>
        <!--</center>-->
        </div>
        </div>
		<%'DATOS TECNICOS MODO EDIT'%>
        <div class="Section" id="S_editDT">
            <a href="#" rel="toggle[editDT]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitDatosTecnicos%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:none;" id="editDT"><%
			  if rfbaja<>"" then
			  else
				''if tecnico>"" then
				if Tfbaja<>"" then
						%><td colspan=4><a class="CELDAREF7" href="javascript:Tecnico('alta');"><%=LitAltaTecnico %></a></td><%
				else
                        EligeCelda "input",mode,"left","","",0,LitComision,"tcomision",20,enc.EncodeForHtmlAttribute(null_s(Tcomision))

                        EligeCelda "input",mode,"left","","",0,LIT_PHLABORAL,"tphlaboral",20,enc.EncodeForHtmlAttribute(null_s(Tphlaboral))

                        EligeCelda "input",mode,"left","","",0,LitIncentivo1,"tincentivo1",20,enc.EncodeForHtmlAttribute(null_s(Tincentivo1))

                        EligeCelda "input",mode,"left","","",0,LitHoraExtraLab,"tphextralab",20,enc.EncodeForHtmlAttribute(null_s(Tphextralab))

                        EligeCelda "input",mode,"left","","",0,LitIncentivo2,"tincentivo2",20,enc.EncodeForHtmlAttribute(null_s(Tincentivo2))

                        EligeCelda "input",mode,"left","","",0,LitHoraExtraFes,"tphextrafes",20,enc.EncodeForHtmlAttribute(null_s(Tphextrafes))

                        strselect = "select matricula,(right(matricula,len(matricula)-5) + ' ' + marca + ' ' + modelo) as descripcion from vehiculos with(nolock) where matricula like ?+'%' order by (marca + ' ' + modelo)"
                        set command2 = nothing
                        set conn2 = Server.CreateObject("ADODB.Connection")
                        set command2 = Server.CreateObject("ADODB.Command")
                        conn2.Open = session("dsn_cliente")
                        conn2.CursorLocation = 3
                        command2.ActiveConnection = conn2
                        command2.CommandTimeout = 60
                        command2.CommandText = strselect
                        command2.CommandType = adCmdText
                        command2.Parameters.Append command2.CreateParameter("@matricula",adVarChar,adParamInput,13,session("ncliente")&"")
                        set rstAux = command2.Execute
                        DrawSelectCelda "CELDA style='width:140px'","","",0,LitVehiculo,"tvehiculo",rstAux,Tmatricula,"matricula","descripcion","",""
						conn2.Close
                        set conn2 = nothing
                        set command2 = nothing

                        strselect = "select codigo,descripcion from almacenes with(nolock) where codigo like ?+'%' order by descripcion"
                        set command2 = nothing
                        set conn2 = Server.CreateObject("ADODB.Connection")
                        set command2 = Server.CreateObject("ADODB.Command")
                        conn2.Open = session("dsn_cliente")
                        conn2.CursorLocation = 3
                        command2.ActiveConnection = conn2
                        command2.CommandTimeout = 60
                        command2.CommandText = strselect
                        command2.CommandType = adCmdText
                        command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
                        set rstAux = command2.Execute
                        DrawSelectCelda "CELDA style='width:140px'","","",0,LitAlmacen,"talmacen",rstAux,TCodAlmacen,"codigo","descripcion","",""
						conn2.Close
                        set conn2 = nothing
                        set command2 = nothing

					DrawFila color_blau%>
                        <table style="width: 100%;"></table>
                        <td colspan=4><a class="CELDAREF7" href="javascript:Tecnico('baja');"><%=LitBajaTecnico %></a></td><%
					CloseFila
				end if
			  end if
			%>
        <!--</center>-->
        </div>
        </div><%

		'** COD JCI-14022003-01 **
		'DATOS OPERARIO MODO EDIT %>
        <div class="Section" id="S_editDO">
            <a href="#" rel="toggle[editDO]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitDatosOperario%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:none;" id="editDO"><%
			  if rfbaja<>"" then
			  else
				''if operario<>"" then
				if Ofbaja<>"" then
					%><td colspan=4><a class="CELDAREF7" href="javascript:Operario('alta');"><%=LitAltaOper %></a></td><%
				else
	       			%><td colspan="4"><%
	       			'BarraOpcionesOperario mode,rdni,rnombre
	       			%></td><%
                    EligeCelda "input",mode,"left","","",0,LitCosteHora,"ocoste_hora",20,Ocoste_hora
					DrawFila color_blau
						%><table style="width: 100%;"></table>
                        <td colspan=4><a class="CELDAREF7" href="javascript:Operario('baja');"><%=LitBajaOper %></a></td><%
					CloseFila
				end if
			  end if
			%>
		<!--</center>-->
        </div>
        </div><%

        BarraNavegacion mode
		'** FIN COD JCI-14022003-01 **'

	elseif mode="browse" then
   		BarraOpciones mode,rdni,rnombre,si_tiene_modulo_comercial
		'DATOS GENERALES MODO BROWSE
		' Inicio Borde Span
		%>
        <div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['browseDG','browseDC','browseDT','browseDO']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['browseDG','browseDC','browseDT','browseDO']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
        </div>
        <div class="Section" id="S_browseDG">
            <a href="#" rel="toggle[browseDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitDatosGenerales%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:;" id="browseDG">
	  		<input type="hidden" name="hdni" value="<%=enc.EncodeForHtmlAttribute(rdni)%>">
	  		<input type="hidden" name="hdomicilio" value="<%=enc.EncodeForHtmlAttribute(null_s(rdomicilio))%>">
      		<table border='0' cellpadding=3 cellspacing=1 width='100%'><%
		      	%><td colspan="4"><%
		       		'BarraOpcionesGen mode,rdni,rnombre
		       	%></td></table><%

                DrawCeldaResponsiveLabel "CELDA","txtMandatory",Litdni,trimCodEmpresa(rdni)

                DrawCeldaResponsiveLabel "CELDA","txtMandatory",Litnombre,enc.EncodeForHtmlAttribute(null_s(rnombre))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitAlias,LitAlias,"",enc.EncodeForHtmlAttribute(null_s(ralias))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitCodigoOp,LitCodigoOp,"",enc.EncodeForHtmlAttribute(null_s(rcodigo))

                strselect = "select descripcion from tipos_entidades with(Nolock) where codigo=?"
                EligeCeldaResponsive "text",mode,"CELDA","","","",LitTipo,LitTipo,"",enc.EncodeForHtmlAttribute(null_s(DLookupP1(strselect,rtipo&"",adVarChar,10,session("dsn_cliente"))))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitSS,LitSS,"",enc.EncodeForHtmlAttribute(null_s(rss))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitAntiguedad,LitAntiguedad,"",enc.EncodeForHtmlAttribute(null_s(rantiguedad))

                EligeCeldaResponsive "text",mode,"CELDA","","","",Litjornada,Litjornada,"",enc.EncodeForHtmlAttribute(null_s(rjornada))
    		'dgb control horario modulo Kyros  19/11/2009
    		'dgm 09.07.12 The Department of time control is always displayed.
                if si_tiene_modulo_ControlPresencia <> 0 then
                    strselect = "select fdescripcion from ch_departamentos with(Nolock) where fiddepartamento=?"
                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitCHDepartamento,LitCHDepartamento,"",enc.EncodeForHtmlAttribute(null_s(DLookupP1(strselect,rdepartamento&"",adVarChar,20,session("dsn_cliente"))))
                else
                    strselect = "select descripcion from departamentos with(Nolock) where codigo=?"
                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitCHDepartamento,LitCHDepartamento,"", enc.EncodeForHtmlAttribute(null_s(DLookupP1(strselect,rdepartamento&"",adVarChar,8,session("dsn_cliente"))))
                end if

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitMaxAmount,LitMaxAmount,"",formatnumber(null_z(rmaxamount),ndecimales,-2,-2,-2)
    	     if si_tiene_modulo_ControlPresencia = 0 then
                EligeCeldaResponsive "text",mode,"CELDA","","","",LitHoraIniMa,LitHoraIniMa,"",enc.EncodeForHtmlAttribute(null_s(rhorainima))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitHoraFinMa,LitHoraFinMa,"",enc.EncodeForHtmlAttribute(null_s(rhorafinma))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitHoraIniTa,LitHoraIniTa,"",enc.EncodeForHtmlAttribute(null_s(rhorainita))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitHoraFinTa,LitHoraFinTa,"",enc.EncodeForHtmlAttribute(null_s(rhorafinta))
    	      end if

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitSueldo,LitSueldo,"",formatnumber(null_z(rsueldo),ndecimales,-2,-2,-2)

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitPhextra,LitPhextra,"",formatnumber(null_z(rphextra),ndecimales,-2,-2,-2)

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitIRPF,LitIRPF,"",formatnumber(null_z(rirpf),ndecimales,-2,-2,-2)

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitSegSocial,LitSegSocial,"",formatnumber(null_z(rsegsocial),ndecimales,-2,-2,-2)
     			if (si_tiene_modulo_bierzo<>0) then   			    
                    DrawDiv "1","",""
                    DrawLabel "","",LitTarifa%><span class="CELDA"><%=enc.EncodeForHtmlAttribute(null_s(rporctarifa))%> <%=LitPorcentaje%> &nbsp;<%=LitTarifaSi %><%=enc.EncodeForHtmlAttribute(null_s(rtarifa)) %> <%=LitTarifaMayor %><%=enc.EncodeForHtmlAttribute(null_s(rsimayor))%></span><%CloseDiv
    			end if    			

                DrawCeldaResponsiveLabel "CELDA","txtMandatory",LitDireccion,enc.EncodeForHtmlAttribute(null_s(rdomicilio))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitPoblacion,LitPoblacion,"",enc.EncodeForHtmlAttribute(null_s(rpoblacion))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitCP,LitCP,"",enc.EncodeForHtmlAttribute(null_s(rcp))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitProvincia,LitProvincia,"",enc.EncodeForHtmlAttribute(null_s(rprovincia))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitPais,LitPais,"",enc.EncodeForHtmlAttribute(null_s(rpais))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitTel1,LitTel1,"",enc.EncodeForHtmlAttribute(null_s(rtelefono))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitTel2,LitTel2,"",enc.EncodeForHtmlAttribute(null_s(rtelefono2))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitFax,LitFax,"",enc.EncodeForHtmlAttribute(null_s(rfax))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitEmail,LitEmail,"",enc.EncodeForHtmlAttribute(null_s(remail))

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitNivel,LitNivel,"",enc.EncodeForHtmlAttribute(null_s(rnivel))

                EligeCeldaResponsive "text",mode,"CELDAREDBOLD","","","",LitFBaja,LitFBaja,"",enc.EncodeForHtmlAttribute(null_s(rfbaja))
                        
                strselect = "select descripcion from cajas with(Nolock) where codigo=?"
                EligeCeldaResponsive "text",mode,"CELDA","","","",LitCaja,LitCaja,"", enc.EncodeForHtmlAttribute(null_s(DLookupP1(strselect,rcaja&"",adVarChar,10,session("dsn_cliente"))))
		   		if si_tiene_paginaSMS<>0 then
                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitMensajeriaSMS,LitMensajeriaSMS,"",Visualizar(rsms)
			 	end if

                EligeCeldaResponsive "text",mode,"CELDA","","","",LitObservaciones,LitObservaciones,"",pintar_saltos_espacios(null_s(robservaciones))
				%>
        <!--</center>-->
        </div>
        </div>
		<%'DATOS COMERCIALES MODO BROWSE'%>
        <div class="Section" id="S_browseDC">
            <a href="#" rel="toggle[browseDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=iif(si_tiene_modulo_comercial<>0,LitDatosComercialesModCom,LitDatosComerciales)%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:none;" id="browseDC">
	     <!--<br />-->
         <!--<center>-->
			    <%
			  if rfbaja<>"" then
			  else
	       		if rcfbaja<>"" then
	       		else
	       			DrawFila color_blau%>
	       				<td colspan="4">
	       				<%'BarraOpciones mode,rdni,rnombre,si_tiene_modulo_comercial%>
	       				</td><%
	       			CloseFila
				end if
	       		if rcfbaja<>"" then
					DrawFila color_blau
						%><td colspan=4><a class="CELDAREF7" href="javascript:Comercial('alta',<%=enc.EncodeForJavascript(null_s(si_tiene_modulo_comercial))%>);"><%=iif(si_tiene_modulo_comercial<>0,LitDarAltaComComercModCom,LitDarAltaComComerc)%></a></td></table><%
					CloseFila
				else%><%
						if si_tiene_modulo_comercial<>0 then
                            EligeCeldaResponsive "text",mode,"CELDA","","","",LitPenalizacion,LitPenalizacion,"",enc.EncodeForHtmlAttribute(null_s(rcpena))
						end if
						if si_tiene_modulo_comercial<>0 then
                            strselect = "select nombre from personal with(Nolock) where dni=?"
                            EligeCeldaResponsive "text",mode,"CELDA","","","",LitSuperiorModCom,LitSuperiorModCom,"",enc.EncodeForHtmlAttribute(null_s(DLookupP1(strselect,rsuperior&"",adVarChar,20,session("dsn_cliente"))))
						else
                            strselect = "select nombre from personal with(Nolock) where dni=?"
                            EligeCeldaResponsive "text",mode,"CELDA","","","",LitSuperior,LitSuperior,"",enc.EncodeForHtmlAttribute(null_s(DLookupP1(strselect,rsuperior&"",adVarChar,20,session("dsn_cliente"))))
						end if
					if si_tiene_modulo_comercial<>0 then
                        EligeCeldaResponsive "text",mode,"CELDA","","","",LitComBase,LitComBase,"",enc.EncodeForHtmlAttribute(null_s(rcbase))
					end if%>
						<table style="width: 100%;"></table>
                        <td colspan=4><a class="CELDAREF7" href="javascript:Comercial('baja',<%=enc.EncodeForJavascript(null_s(si_tiene_modulo_comercial))%>);"><%=iif(si_tiene_modulo_comercial<>0,LitDarBajaComComercModCom,LitDarBajaComComerc)%></a>
					<%
				end if
			  end if%>
        <!--</center>-->
        </div>
        </div>
	  	<%'DATOS TECNICO MODO BROWSE'%>
        <div class="Section" id="S_browseDT">
            <a href="#" rel="toggle[browseDT]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitDatosTecnicos%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:none;" id="browseDT">
			    <%
			  if rfbaja<>"" then
			  else
				''if tecnico>"" then
				if Tfbaja<>"" then
					DrawFila color_blau%>
						<td colspan=4><a class="CELDAREF7" href="javascript:Tecnico('alta');"><%=LitAltaTecnico %></a></td>
					<%CloseFila
				else
                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitComision,LitComision,"",formatnumber(Tcomision,decpor,-1,0,-1) & "%"

                    EligeCeldaResponsive "text",mode,"CELDA","","","",LIT_PHLABORAL,LIT_PHLABORAL,"",formatnumber(Tphlaboral,ndecimales,-1,0,-1)

                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitIncentivo1,LitIncentivo1,"",formatnumber(Tincentivo1,ndecimales,-1,0,-1)

                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitHoraExtraLab,LitHoraExtraLab,"",formatnumber(Tphextralab,ndecimales,-1,0,-1)

                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitIncentivo2,LitIncentivo2,"",formatnumber(Tincentivo2,ndecimales,-1,0,-1)

                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitHoraExtraFes,LitHoraExtraFes,"",formatnumber(Tphextrafes,ndecimales,-1,0,-1)

                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitVehiculo,LitVehiculo,"",trimCodEmpresa(Tvehiculo)

                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitAlmacen,LitAlmacen,"",enc.EncodeForHtmlAttribute(null_s(Talmacen))%>
						<table style="width: 100%;"></table>
						<td colspan=4><a class="CELDAREF7" href="javascript:Tecnico('baja');"><%=LitBajaTecnico %></a></td>
					<%
				end if
			  end if%>
        </div>
        </div>
	  	<%'** COD JCI-14022003-01 **
		'DATOS OPERARIO MODO BROWSE %>
        <div class="Section" id="S_browseDO">
            <a href="#" rel="toggle[browseDO]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitDatosOperario%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:none;" id="browseDO"><%

			  if rfbaja<>"" then
			  else
				if Ofbaja<>"" then
				else%><td colspan="4">
	       			<%'BarraOpcionesOperario mode,rdni,rnombre%>
	       			</td><%
                    EligeCeldaResponsive "text",mode,"CELDA","","","",LitCosteHora,LitCosteHora,"",formatnumber(Ocoste_hora,ndecimales,-1,0,-1)
				end if
			  end if%>
		  <%if rfbaja<>"" then
		  else
			if Ofbaja<>"" then%>
				<table style="width: 100%;"></table>
				<a class="CELDAREF7" href="javascript:Operario('alta');"><%=LitAltaOper %></a>
			<%else%>
                <div class="overflowXauto">
				<table class="width90 md-table-responsive bCollapse"><%
					Drawfila color_terra%>
						<td class="ENCABEZADOL underOrange width10"><input type="Checkbox" name="check" onclick="seleccionar('fr_OperarioFases','operario_fases','check');" checked></td>
						<td class="ENCABEZADOL underOrange width40"><%=LitFase%></td>
						<td class="ENCABEZADOL underOrange width40"><%=LitCoste%></td>
					<%CloseFila%>
				</table>
				<iframe class='width90 iframe-data md-table-responsive' id='frOperarioFases' name="fr_OperarioFases" src='operario_fases.asp?mode=browse&operario=<%=enc.EncodeForJavascript(null_s(operario))%>' frameborder="yes" noresize="noresize"></iframe>
                <table class="width90 md-table-responsive bCollapse">
                    <tr>
                        <td class='CELDAL7 width70'><a class="CELDAREF7" href="javascript:Operario('baja');"><%=LitBajaOper %></a></td>
				        <td class='CELDAR7 width20' style="text-align:right;"><a class="" href="javascript:if(GestionCostes('save','<%=enc.EncodeForJavascript(null_s(operario))%>'));"><img src='../images/<%=ImgDiskette%>' <%=ParamImgDiskette%> alt='<%=LitGuardarCostes%>' title='<%=LitGuardarCostes%>'  align="center"></a>&nbsp;<a class="ic-delete" href="javascript:if(GestionCostes('delete','<%=enc.EncodeForJavascript(operario)%>'));"><img src='<%=themeIlion %><%=ImgEliminarDet%>' <%=ParamImgEliminar%> alt='<%=LitEliminarCoste%>' title='<%=LitEliminarCoste%>'  align="center"></a></td>
                    </tr>
                </table>
                </div>
				<br/>
			<%end if
		  end if%>
        <!--</center>-->
        </div>
        </div>
	  	<%'** FIN COD JCI-14022003-01 **'
        BarraNavegacion mode
        %>
	<% elseif mode="search" then
		
    end if%>
</form>
<%
end if
'** COD JCI-14022003-01 **
set rst = Nothing
set rstAux = Nothing
set rstDom = Nothing
set rst = Nothing
set rstAux2 = Nothing
connRound.close
set connRound = nothing
set connCH=nothing
'** FIN COD JCI-14022003-01 **%>
</body>
</html>