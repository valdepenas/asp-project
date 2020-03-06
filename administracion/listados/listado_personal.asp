<%@ Language=VBScript %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloLP%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<META HTTP-EQUIV="Content-style-TypeCONTENT="text/css">
<LINK REL="styleSHEET" href="../../pantalla.css" MEDIA="SCREEN">
<LINK REL="styleSHEET" href="../../impresora.css" MEDIA="PRINT">
</head>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
%>  

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
 
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>

<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../modulos.inc" -->

<!--#include file="../personal.inc" -->

<!--#include file="../../tablasResponsive.inc" -->
 <!--#include file="../../styles/formularios.css.inc" --> 


<script language="javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
function sel_todos(modo)
{
	document.listado_personal.opcdomicilio.checked=modo;
	document.listado_personal.opccodigopostal.checked=modo;
	document.listado_personal.opcpoblacion.checked=modo;
	document.listado_personal.opcprovincia.checked=modo;
	document.listado_personal.opctelefono.checked=modo;
	document.listado_personal.opcalias.checked=modo;
	document.listado_personal.opcmovils.checked=modo;
	document.listado_personal.opcantiguedad.checked=modo;
	document.listado_personal.opcnumsegsocial.checked=modo;
	document.listado_personal.opcemail.checked=modo;
	document.listado_personal.opcnivel.checked=modo;
	document.listado_personal.opccaja.checked=modo;
	document.listado_personal.opchoras.checked=modo;
	document.listado_personal.opcsalario.checked=modo;
	document.listado_personal.opcvalorhextra.checked=modo;
}

function Traerpersonal() {
	document.listado_personal.action="listado_personal.asp?dni=" + document.listado_personal.dni.value + "&mode=traerpersonal";
	document.listado_personal.submit();
}

function tratarsolo(que){
	if (que=='1'){
		document.listado_personal.soltec.checked=false;
		document.listado_personal.solope.checked=false;
	}
	if (que=='2'){
		document.listado_personal.solope.checked=false;
		document.listado_personal.solcom.checked=false;
	}
	if (que=='3'){
		document.listado_personal.soltec.checked=false;
		document.listado_personal.solcom.checked=false;
	}
}
</script>

<body onload="self.status='';" class="BODY_ASP">
<%
'RGU 16/11/2007 CAMBIO DSN PARA LISTADOS

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
borde=0

  set rst = Server.CreateObject("ADODB.Recordset")
  set rst2 = Server.CreateObject("ADODB.Recordset")
  set rstAux = Server.CreateObject("ADODB.Recordset")
  set rstAux2 = Server.CreateObject("ADODB.Recordset")

%>
<form name="listado_personal" method="post">                                                             
	<%  PintarCabecera "listado_personal.asp"

		'Leer parámetros de la página
  		mode=Request.QueryString("mode")
%><input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"><%

		if Request.QueryString("dni")>"" then
			dni=limpiaCadena(Request.QueryString("dni"))
		else
			dni=limpiaCadena(Request.Form("dni"))
		end if

		if request.querystring("nombre")>"" then
			nombre=limpiaCadena(request.querystring("nombre"))
		else
			nombre=limpiaCadena(request.Form("nombre"))
		end if

		if request.querystring("tpersonal")>"" then
			tpersonal=limpiaCadena(request.querystring("tpersonal"))
		else
			tpersonal=limpiaCadena(request.Form("tpersonal"))
		end if


		if request.querystring("tiposolo")>"" then
			tiposolo=limpiaCadena(request.querystring("tiposolo"))
		else
			tiposolo=limpiaCadena(request.Form("tiposolo"))
		end if

		if request.querystring("orden")>"" then
			orden=limpiaCadena(request.querystring("orden"))
		else
			orden=limpiaCadena(request.Form("orden"))
		end if

	if request.querystring("opcdomicilio")>"" then
		opcdomicilio=limpiaCadena(request.querystring("opcdomicilio"))
	else
		opcdomicilio=limpiaCadena(request.form("opcdomicilio"))
	end if

	if opcdomicilio>"" then opcdomicilio=1

	if request.querystring("opccodigopostal")>"" then
		opccodigopostal=limpiaCadena(request.querystring("opccodigopostal"))
	else
		opccodigopostal=limpiaCadena(request.form("opccodigopostal"))
	end if

	if opccodigopostal>"" then opccodigopostal=1

	if request.querystring("opcpoblacion")>"" then
		opcpoblacion=limpiaCadena(request.querystring("opcpoblacion"))
	else
		opcpoblacion=limpiaCadena(request.form("opcpoblacion"))
	end if

	if opcpoblacion>"" then	opcpoblacion=1

	if request.querystring("opcprovincia")>"" then
		opcprovincia=limpiaCadena(request.querystring("opcprovincia"))
	else
		opcprovincia=limpiaCadena(request.form("opcprovincia"))
	end if

	if opcprovincia>"" then	opcprovincia=1

	if request.querystring("opctelefono")>"" then
		opctelefono=limpiaCadena(request.querystring("opctelefono"))
	else
		opctelefono=limpiaCadena(request.form("opctelefono"))
	end if

	if opctelefono>"" then	opctelefono=1

	if request.querystring("opcalias")>"" then
		opcalias=limpiaCadena(request.querystring("opcalias"))
	else
		opcalias=limpiaCadena(request.form("opcalias"))
	end if

	if opcalias>"" then opcalias=1

	if request.querystring("opcmovil")>"" then
		opcmovil=limpiaCadena(request.querystring("opcmovil"))
	else
		opcmovil=limpiaCadena(request.form("opcmovil"))
	end if

	if opcmovil>"" then opcmovil=1

	if request.querystring("opcantiguedad")>"" then
		opcantiguedad=limpiaCadena(request.querystring("opcantiguedad"))
	else
		opcantiguedad=limpiaCadena(request.form("opcantiguedad"))
	end if

	if opcantiguedad>"" then opcantiguedad=1

	if request.querystring("opcnumsegsocial")>"" then
		opcnumsegsocial=limpiaCadena(request.querystring("opcnumsegsocial"))
	else
		opcnumsegsocial=limpiaCadena(request.form("opcnumsegsocial"))
	end if

	if opcnumsegsocial>"" then opcnumsegsocial=1

	if request.querystring("opcemail")>"" then
		opcemail=limpiaCadena(request.querystring("opcemail"))
	else
		opcemail=limpiaCadena(request.form("opcemail"))
	end if

	if opcemail>"" then opcemail=1

	if request.querystring("opcnivel")>"" then
		opcnivel=limpiaCadena(request.querystring("opcnivel"))
	else
		opcnivel=limpiaCadena(request.form("opcnivel"))
	end if

	if opcnivel>"" then opcnivel=1

	if request.querystring("opccaja")>"" then
		opccaja=limpiaCadena(request.querystring("opccaja"))
	else
		opccaja=limpiaCadena(request.form("opccaja"))
	end if

	if opccaja>"" then opccaja=1

	if request.querystring("opchoras")>"" then
		opchoras=limpiaCadena(request.querystring("opchoras"))
	else
		opchoras=limpiaCadena(request.form("opchoras"))
	end if

	if opchoras>"" then opchoras=1

	if request.querystring("opcsalario")>"" then
		opcsalario=limpiaCadena(request.querystring("opcsalario"))
	else
		opcsalario=limpiaCadena(request.form("opcsalario"))
	end if

	if opcsalario>"" then opcsalario=1

	if request.querystring("opcvalorhextra")>"" then
		opcvalorhextra=limpiaCadena(request.querystring("opcvalorhextra"))
	else
		opcvalorhextra=limpiaCadena(request.form("opcvalorhextra"))
	end if

	if opcvalorhextra>"" then opcvalorhextra=1

	apaisado=iif(limpiaCadena(request.form("apaisado"))>"","SI","")

	mostrarclientescom=iif(limpiaCadena(request.form("mostrarclientescom"))>"","SI","")
	
	'jcg 26/02/2009
	if request.querystring("checkClientesBaja")>"" then
		checkClientesBaja=limpiaCadena(request.querystring("checkClientesBaja"))
	else
		checkClientesBaja=limpiaCadena(request.form("checkClientesBaja"))
	end if
	
	si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
	si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)
	si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)

	Alarma "listado_personal.asp"

  '*********************************************************************************************
  'Se muestran parametros de seleccion
  '*********************************************************************************************

	if mode="traerpersonal" then
		if dni> "" then
			dni=session("ncliente") & dni

			'nombre=d_lookup("nombre","personal","dni='" & dni & "'",session("backendlistados"))

            strselect1="select nombre from personal with(nolock) where dni=?"
            nombre=DLookupP1(strselect1,dni&"",adVarchar,20,session("backendlistados"))

			if nombre="" then
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgPersonalNoExiste%>");
				</script><%
			else
			    rst2.open "select dni,nombre,fbaja from personal with(nolock) where dni='" & dni & "' and fbaja is null",session("backendlistados"),adOpenKeyset,adLockOptimistic
				if rst2.eof then
					%><script language="javascript" type="text/javascript">
						window.alert("<%=LitMsgPersonalDadoBaja%>");
					</script><%
					dni=""
					nombre=""
				else
					nombre=nombre
				end if
				rst2.close
			end if
			mode="param"
			%><script language="javascript" type="text/javascript">
				document.listado_personal.mode.value="<%=enc.EncodeForJavascript(mode)%>";
			</script><%
		else
			dni=""
			nombre=""
			mode="param"
			%><script language="javascript" type="text/javascript">
				document.listado_personal.mode.value="<%=enc.EncodeForJavascript(mode)%>";
			</script><%
		end if
	end if

  if mode="param" then
		%><table><%
			
				'DrawCelda2 "CELDA width=200", "left", false,LitEmpleado+":"
                DrawDiv "1", "", ""
                DrawLabel "", "", LitEmpleado
				%>
					<input class='width20' type="text" name="dni" size=10 value="<%=iif(dni & "">"",trimCodEmpresa(dni) ,"")%>" onchange="Traerpersonal();">
					<a class='CELDAREFB'  href="javascript:AbrirVentana('../../administracion/personal_buscar.asp?viene=listado_personal&titulo=<%=LitSelEmpleado%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitBuscarEmpleado%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
					<input class="width60" style="width:180px" readonly type="text" name="nombre" size="48" value="<%=iif(nombre & "">"",enc.EncodeForHtmlAttribute(nombre),"")%>">
				<%
                CloseDiv
			
			
				'DrawCelda2 "CELDA width=200", "left", false, LitTipo & " : "
                'DrawDiv "1", "", ""
                'DrawLabel "", "", LitTipo
				rstAux.cursorlocation=3
				rstAux.open "select codigo,descripcion from tipos_entidades with(nolock) where tipo='PERSONAL' and codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
				DrawSelectCelda "CELDA","150","",0,LitTipo,"tpersonal",rstAux,tpersonal,"codigo","descripcion","",""
				rstAux.Close
				'DrawCelda2 "CELDA", "left", false,"&nbsp;"
			    'CloseDiv

			DrawDiv "1", "", ""
				'DrawCelda2 "CELDA  width=200", "left", false, LitMostrar & " : "
                 DrawLabel "", "", LitMostrar   

				%>
					<select name="tiposolo" class="CELDA" style='width:150px'>
						<%if si_tiene_modulo_comercial<>0 then%>
							<option <%=iif(tiposolo=LitSoloComModCom,"selected","")%> value="<%=LitSoloComModCom%>"><%=LitSoloComModCom%></option>
						<%else%>
							<option <%=iif(tiposolo=LitSoloCom,"selected","")%> value="<%=LitSoloCom%>"><%=LitSoloCom%></option>
						<%end if%>
						<%if si_tiene_modulo_mantenimiento<>0 then%>
							<option <%=iif(tiposolo=LitSoloTec,"selected","")%> value="<%=LitSoloTec%>"><%=LitSoloTec%></option>
						<%end if%>
						<%if si_tiene_modulo_produccion<>0 then%>
							<option <%=iif(tiposolo=LitSoloOpe,"selected","")%> value="<%=LitSoloOpe%>"><%=LitSoloOpe%></option>
						<%end if%>
						<option <%=iif(tiposolo=LitTodos or tiposolo="","selected","")%> value="<%=LitTodos%>"><%=LitTodos%></option>
					</select>
				<%
			 CloseDiv
                    
			    DrawDiv "1", "", ""
				'DrawCelda2 "CELDA  width=200", "left", false, LitOrden & " : "
                DrawLabel "", "", LitOrden 
				%><select name="orden" class="CELDA" style='width:150px'>
						<option <%=iif(orden=LitNombre or orden="","selected","")%> value="<%=LitNombre%>"><%=LitNombre%></option>
						<option <%=iif(orden=LitDni,"selected","")%> value="<%=LitDni%>"><%=LitDni%></option>
					</select>
			<%	
			CloseDiv
                DrawDiv "1", "", ""
				'DrawCelda2 "CELDA width=200", "left", false, LitApaisado & " : "
                DrawLabel "", "", LitApaisado
				DrawCheckCelda "CELDA","","",0,"","apaisado",iif(apaisado>"",-1,0)
			    CloseDiv
        %>
        </table>
        <table>
        <%			
			    DrawDiv "1", "", ""
				if si_tiene_modulo_comercial<>0 then
					'DrawCelda2 "CELDA width=200", "left", false, LitMostrarClientesComModCom & " : "
                    DrawLabel "", "", LitMostrarClientesComModCom
				else
					'DrawCelda2 "CELDA width=200", "left", false, LitMostrarClientesCom & " : "
                    DrawLabel "", "", LitMostrarClientesCom
				end if
				DrawCheckCelda "CELDA width=50","","",0,"","mostrarclientescom",iif(mostrarclientescom>"",-1,0)
                CloseDiv
                ' jcg 25/02/2009
				'DrawCelda2 "CELDA width=230", "left", false, LitNoMostarPersonalDadoDeBaja & " : "
                DrawDiv "1", "", ""
                DrawLabel "", "", LitNoMostarPersonalDadoDeBaja
				%><input type='checkbox' name='checkClientesBaja' checked="checked"><%
                CloseDiv
                
                %>
               
        </table>
		<hr/>
		<table border='<%=borde%>' cellspacing="1" cellpadding="1"><%
			DrawDiv "3-sub", "background-color: #eae7e3", ""
				'DrawCelda2 "ENCABEZADOL", "left", false, LitCamposOpcionales
                DrawLabel "", "", LitCamposOpcionales
			CloseDiv %>
		</table>
		<table border='<%=borde%>' cellspacing="1" cellpadding="1"><%
			
				%>
				<%
					DrawDiv "7", "", ""
                    DrawDiv "3-sub", "background-color: #eae7e3", ""
						'DrawCelda2Span "ENCABEZADOL", "left", false, LitDatosGenerales,5
                        DrawLabel "", "", LitDatosGenerales
                    CloseDiv
					
					
						'DrawCelda2 "CELDA style='width:130px'", "left", false, LitDomicilio
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitDomicilio, "opcdomicilio", "", iif(opcdomicilio="1","True","0") 
						'DrawCheckCelda "CELDA","","",0,"","opcdomicilio",iif(opcdomicilio="1","True","0")                        
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA style='width:100px'", "left", false, LitCp1
						'DrawCheckCelda "CELDA","","",0,"","opccodigopostal",iif(opccodigopostal="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitCp1, "opccodigopostal", "", iif(opccodigopostal="1","True","0")
                      
					
					
						'DrawCelda2 "CELDA", "left", false, LitPoblacion
						'DrawCheckCelda "CELDA","","",0,"","opcpoblacion",iif(opcpoblacion="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitPoblacion, "opcpoblacion", "", iif(opcpoblacion="1","True","0")
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitProvincia
						'DrawCheckCelda "CELDA","","",0,"","opcprovincia",iif(opcprovincia="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitProvincia, "opcprovincia", "", iif(opcprovincia="1","True","0")
					
					
						'DrawCelda2 "CELDA", "left", false, LitTelefono
						'DrawCheckCelda "CELDA","","",0,"","opctelefono",iif(opctelefono="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitTelefono, "opctelefono", "", iif(opctelefono="1","True","0")
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitAlias
						'DrawCheckCelda "CELDA","","",0,"","opcalias",iif(opcalias="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitAlias, "opcalias", "", iif(opcalias="1","True","0")
					
					
						'DrawCelda2 "CELDA", "left", false, LitMovil
						'DrawCheckCelda "CELDA","","",0,"","opcmovil",iif(opcmovil="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitMovil, "opcmovils", "", iif(opcmovil="1","True","0")
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitAntiguedad
						'DrawCheckCelda "CELDA","","",0,"","opcantiguedad",iif(opcantiguedad="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitAntiguedad, "opcantiguedad", "", iif(opcantiguedad="1","True","0")
					
					
						'DrawCelda2 "CELDA", "left", false, LitNSSocial
						'DrawCheckCelda "CELDA","","",0,"","opcnumsegsocial",iif(opcnumsegsocial="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitNSSocial, "opcnumsegsocial", "", iif(opcnumsegsocial="1","True","0")
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitEmail1
						'DrawCheckCelda "CELDA","","",0,"","opcemail",iif(opcemail="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitEmail1, "opcemail", "", iif(opcemail="1","True","0")
					
					
						'DrawCelda2 "CELDA", "left", false, LitNivel
						'DrawCheckCelda "CELDA","","",0,"","opcnivel",iif(opcnivel="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitNivel, "opcnivel", "", iif(opcnivel="1","True","0")
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitCaja
						'DrawCheckCelda "CELDA","","",0,"","opccaja",iif(opccaja="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitCaja, "opccaja", "", iif(opccaja="1","True","0")
					
					
						'DrawCelda2 "CELDA", "left", false, LitHoras
						'DrawCheckCelda "CELDA","","",0,"","opchoras",iif(opchoras="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitHoras, "opchoras", "", iif(opchoras="1","True","0")
						'DrawCelda "CELDA","10%","",0," "
						'DrawCelda2 "CELDA", "left", false, LitSueldo
						'DrawCheckCelda "CELDA","","",0,"","opcsalario",iif(opcsalario="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitSueldo, "opcsalario", "", iif(opcsalario="1","True","0")
					
					
						'DrawCelda2 "CELDA", "left", false, LitValorHExt
						'DrawCheckCelda "CELDA","","",0,"","opcvalorhextra",iif(opcvalorhextra="1","True","0")
                        EligeCelda "check-listado", "edit", "", "", "", 0, LitValorHExt, "opcvalorhextra", "", iif(opcvalorhextra="1","True","0")
					%>
				
			<%CloseDiv%>
		</table>
		<hr/>
		<table>
			<tr bgcolor="<%=color_blau%>">
				<td class="CELDABOT" onclick="javascript:sel_todos(true);">
					<%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,""%>
				</td>
				<td>&nbsp;</td>
				<td class="CELDABOT" onclick="javascript:sel_todos(false);">
					<%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,""%>
				</td>
			</tr>
		</table><%
   end if

   '*********************************************************************************************
   ' Se muestran los datos de la consulta
   '*********************************************************************************************

   if mode="browse" then
		'MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='234'", DSNIlion)
		'MAXPDF=d_lookup("maxpdf", "limites_listados", "item='234'", DSNIlion)

        strselectMAXPAG="select maxpagina from limites_listados where item=?"
        MAXPAGINA=DlookupP1(strselectMAXPAG,"234",AdVarChar,3,DSNIlion)

        strselectMAXPDF="select maxpdf from limites_listados where item=?"
        MAXPDF=DlookupP1(strselectMAXPDF,"234",AdVarChar,3,DSNIlion)

	   if dni<>"" then
	     dni=session("ncliente") & dni                                                                                    
	   end if
	   %>
		<input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>
		<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'>
		<input type="hidden" name="dni" value="<%=EncodeForHtml(dni)%>">
		<input type="hidden" name="tpersonal" value="<%=EncodeForHtml(tpersonal)%>">
		<input type="hidden" name="tiposolo" value="<%=EncodeForHtml(tiposolo)%>">
		<input type="hidden" name="orden" value="<%=EncodeForHtml(orden)%>">
		<input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>">
		<input type="hidden" name="mostrarclientescom" value="<%=EncodeForHtml(mostrarclientescom)%>">
		
		<%'jcg %>
		<input type="hidden" name="checkClientesBaja" value="<%=EncodeForHtml(checkClientesBaja)%>">
           
        <input type="hidden" name="opcdomicilio" value="<%=EncodeForHtml(opcdomicilio)%>">
		<input type="hidden" name="opccodigopostal" value="<%=EncodeForHtml(opccodigopostal)%>">
		<input type="hidden" name="opcpoblacion" value="<%=EncodeForHtml(opcpoblacion)%>">
		<input type="hidden" name="opcprovincia" value="<%=EncodeForHtml(opcprovincia)%>">
		<input type="hidden" name="opctelefono" value="<%=EncodeForHtml(opctelefono)%>">
		<input type="hidden" name="opcalias" value="<%=EncodeForHtml(opcalias)%>">
		<input type="hidden" name="opcmovil" value="<%=EncodeForHtml(opcmovil)%>">
		<input type="hidden" name="opcantiguedad" value="<%=EncodeForHtml(opcantiguedad)%>">
		<input type="hidden" name="opcnumsegsocial" value="<%=EncodeForHtml(opcnumsegsocial)%>">
		<input type="hidden" name="opcemail" value="<%=EncodeForHtml(opcemail)%>">
		<input type="hidden" name="opcnivel" value="<%=EncodeForHtml(opcnivel)%>">
		<input type="hidden" name="opccaja" value="<%=EncodeForHtml(opccaja)%>">
		<input type="hidden" name="opchoras" value="<%=EncodeForHtml(opchoras)%>">
		<input type="hidden" name="opcsalario" value="<%=EncodeForHtml(opcsalario)%>">
		<input type="hidden" name="opcvalorhextra" value="<%=EncodeForHtml(opcvalorhextra)%>">

	   <%VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarPersonal)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

		encabezado=0

		strwhere=" where p.dni like '" & session("ncliente") & "%' and"                                          

		if dni & "">"" then
			strwhere=strwhere & " p.dni='" & dni & "' and"
			%><font class="CELDA"><%=LitEmpleado & " : " & EncodeForHtml(trimCodEmpresa(dni)) & " - " & d_lookup("nombre","personal","dni='" & dni & "'",session("backendlistados"))%></font><br/><%
			encabezado=1
		else
			''strwhere=strwhere & " p.fbaja is null and"
		end if

		if tpersonal & "">"" then
				rstAux2.open "select codigo,descripcion from tipos_entidades with(nolock) where tipo='PERSONAL' and codigo like '" & tpersonal & "' order by descripcion",session("backendlistados")
				nomPersonal=rstAux2("descripcion")
				rstAux2.Close                                            
			%><font class="CELDA"><%=LitTipo & " : " & EncodeForHtml(nomPersonal)%></font><br/><%
			strwhere=strwhere & " p.tipo='" & EncodeForHtml(tpersonal) & "' and"
			encabezado=1                                                       
		end if

		if tiposolo & "">"" then
			select case tiposolo
				case LitSoloCom,LitSoloComModCom:
					if si_tiene_modulo_comercial<>0 then
						%><font class="CELDA"><%=LitSoloComModCom%></font><br/><%
					else
						%><font class="CELDA"><%=LitSoloCom%></font><br/><%
					end if
					strwhere=strwhere & " c.comercial is not null and"
					encabezado=1
				case LitSoloTec:
					%><font class="CELDA"><%=LitSoloTec%></font><br/><%
					strwhere=strwhere & " t.dni is not null and"
					encabezado=1
				case LitSoloOpe:
					%><font class="CELDA"><%=LitSoloOpe%></font><br/><%
					strwhere=strwhere & " o.operario is not null and"
					encabezado=1
				case LitTodos & "ESTE NO SE MUESTRA":
					%><font class="CELDA"><%=LitTodos%></font><br/><%
					encabezado=1
			end select
		end if

			strwhere=mid(strwhere,1,len(strwhere)-4) 'para quitar el ultimo and'

		if encabezado=1 then
			%><hr/><%
		end if

		if Request.QueryString("lote")=1 or Request.QueryString("lote")="" then

			suma_opciones=0
			if null_z(opcdomicilio)=1 then
				suma_opciones=1
			end if
			if null_z(opccodigopostal)=1 then
				suma_opciones=1
			end if
			if null_z(opcpoblacion)=1 then
				suma_opciones=1
			end if
			if null_z(opcprovincia)=1 then
				suma_opciones=1
			end if
			if null_z(opctelefono)=1 then
				suma_opciones=1
			end if
			if null_z(opcmovil)=1 then
				suma_opciones=1
			end if
			if null_z(opcemail)=1 then
				suma_opciones=1
			end if

		set conn = Server.CreateObject("ADODB.Connection")
		conn.open session("backendlistados")
		
		
		'jcg 25/02/2009
		strselect="EXEC ListadoPersonal @strwhere='" & reemplazar(strwhere,"'","''") & "',@orden='" & orden & "',@num_columna='',@modo='CREAR',@mostrarclientes='" & mostrarclientescom & "',@opciones=" & suma_opciones & ",@usuario='" & session("usuario") & "'"
		set rs = conn.execute(strselect)
		
		set rs = nothing
		conn.close
		set conn = nothing
	end if

	strselect="select max(num_columna) as maximo from [" & session("usuario") & "]"

		rst.Open strselect, session("backendlistados"),adUseClient, adLockReadOnly
		if rst.eof then
			rst.Close
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitMsgDatosNoExiste%>");
			      parent.window.frames["botones"].document.location = "listado_personal_bt.asp?mode=param";
				document.location="listado_personal.asp?mode=param";
				
			</script><%
		else

			'Calculos de páginas--------------------------'
			lote=limpiaCadena(Request.QueryString("lote"))
			if lote="" then
				lote=1
			end if
			sentido=limpiaCadena(Request.QueryString("sentido"))

			if rst("maximo")>"" then
			  max=rst("maximo")
			else
			  max=0
			end if

			lotes=max/MAXPAGINA

			if rst("maximo")>"" then
			%><script language="javascript" type="text/javascript">
			</script><%
			else
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgDatosNoExiste%>");
				document.location="listado_personal.asp?mode=param";
				parent.botones.document.location="listado_personal_bt.asp?mode=param";
			</script><%
			end if

			if lotes>clng(lotes) then
				lotes=clng(lotes)+1
			else
				lotes=clng(lotes)
			end if
			if sentido="next" then
				lote=lote+1
			elseif sentido="prev" then
				lote=lote-1
			end if
			rst.close
		end if

    ' jcg 26/02/2009: Filtro por clientes dados de baja
    strfiltroClientesBaja = ""
    if(checkClientesBaja = "on") then
        strfiltroClientesBaja = " fbaja is null AND "
    end if
    
    strselect="select * from [" & session("usuario") & "] where " + strfiltroClientesBaja + "num_columna>" & ((lote-1)*MAXPAGINA) & " and num_columna<=" & ((lote)*MAXPAGINA)
    	
    rst.Open strselect, session("backendlistados"),adUseClient, adLockReadOnly
		%><input type="hidden" name="NumRegs" value="<%=lotes*MAXPAGINA%>"><%
		if rst.EOF then
			rst.Close
		else

			NavPaginas lote,lotes,campo,criterio,texto,1

			'Fila de encabezado
			fila=1

	if request.form("campo_en_el_que_estamos")>"" then
		campo_en_el_que_estamos=limpiaCadena(request.form("campo_en_el_que_estamos"))
	else
		campo_en_el_que_estamos=""
	end if

	if request.form("dni_donde_estamos")>"" then
		dni_donde_estamos=limpiaCadena(request.form("dni_donde_estamos"))
	else
		dni_donde_estamos=""
	end if

no_encabezado=1

			while not rst.EOF

				CheckCadena rst("dni")

if (campo_en_el_que_estamos="escliente" and dni_donde_estamos<>rst("dni")) then
	campo_en_el_que_estamos=""
end if

				if campo_en_el_que_estamos="" then
						sumaarriba=0
						espacio_anchoa=0
						if null_z(opcalias)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+10
						end if
						if null_z(opcantiguedad)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+8
						end if
						if null_z(opcnumsegsocial)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+12
						end if
						if null_z(opcnivel)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+8
						end if
						if null_z(opccaja)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+12
						end if
						if null_z(opchoras)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+5
						end if
						if null_z(opcsalario)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+9
						end if
						if null_z(opcvalorhextra)=1 then
						else
							sumaarriba=sumaarriba+1
							espacio_anchoa=espacio_anchoa+10
						end if

						sumaabajo=0
						espacio_anchob=0
						if null_z(opcdomicilio)=1 then
						else
							sumaabajo=sumaabajo+1
							espacio_anchob=espacio_anchob+26
						end if
						if null_z(opccodigopostal)=1 then
						else
							sumaabajo=sumaabajo+1
							espacio_anchob=espacio_anchob+10
						end if
						if null_z(opcpoblacion)=1 then
						else
							sumaabajo=sumaabajo+1
							espacio_anchob=espacio_anchob+8
						end if
						if null_z(opcprovincia)=1 then
						else
							sumaabajo=sumaabajo+1
							espacio_anchob=espacio_anchob+12
						end if
						if null_z(opctelefono)=1 then
						else
							sumaabajo=sumaabajo+1
							espacio_anchob=espacio_anchob+8
						end if
						if null_z(opcmovil)=1 then
						else
							sumaabajo=sumaabajo+1
							espacio_anchob=espacio_anchob+12
						end if
						if null_z(opcemail)=1 then
						else
							sumaabajo=sumaabajo+3
							espacio_anchob=espacio_anchob+24
						end if


if no_encabezado=1 then
					%><table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
						DrawFila color_fondo
							DrawCelda "TDBORDECELDA7 style='width:16%'","","",0,"<b>" & LitNombre & "</b>"
							DrawCelda "TDBORDECELDA7 style='width:10%'","","",0,"<b>" & LitDni & "</b>"

							if null_z(opcalias)=1 then
								DrawCelda "TDBORDECELDA7 style='width:10%'","","",0,"<b>" & LitAlias & "</b>"
							end if
							if null_z(opcantiguedad)=1 then
								DrawCelda "TDBORDECELDA7 style='width:8%'","","",0,"<b>" & LitAntiguedad & "</b>"
							end if
							if null_z(opcnumsegsocial)=1 then
								DrawCelda "TDBORDECELDA7 style='width:12%'","","",0,"<b>" & LitNSSocial & "</b>"
							end if
							if null_z(opcnivel)=1 then
								DrawCelda "TDBORDECELDA7 style='width:8%' align='right'","","",0,"<b>" & LitNivel & "</b>"
							end if
							if null_z(opccaja)=1 then
								DrawCelda "TDBORDECELDA7 style='width:12%'","","",0,"<b>" & LitCaja & "</b>"
							end if
							if null_z(opchoras)=1 then
								DrawCelda "TDBORDECELDA7 style='width:5%' align='right'","","",0,"<b>" & LitHoras & "</b>"
							end if
							if null_z(opcsalario)=1 then
								DrawCelda "TDBORDECELDA7 style='width:9%' align='right'","","",0,"<b>" & LitSueldo & "</b>"
							end if
							if null_z(opcvalorhextra)=1 then
								DrawCelda "TDBORDECELDA7 style='width:10%' align='right'","","",0,"<b>" & LitValorHExt & "</b>"
							end if
							if sumaarriba>0 then
								DrawCelda "TDBORDECELDA7 style='width:" & espacio_ancho & "%' colspan='" & sumaarriba & "'","","",0,"&nbsp;"
							end if
						CloseFila
						if sumaabajo=0 then
							DrawFila color_fondo
						end if
							if null_z(opcdomicilio)=1 then
								DrawCelda "TDBORDECELDA7 colspan='2'","","",0,"<b>" & LitDomicilio & "</b>"
							end if
							if null_z(opccodigopostal)=1 then
								DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitCp1 & "</b>"
							end if
							if null_z(opcpoblacion)=1 then
								DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitPoblacion & "</b>"
							end if
							if null_z(opcprovincia)=1 then
								DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitProvincia & "</b>"
							end if
							if null_z(opctelefono)=1 then
								DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitTelefono & "</b>"
							end if
							if null_z(opcmovil)=1 then
								DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitMovil & "</b>"
							end if
							if null_z(opcemail)=1 then
								DrawCelda "TDBORDECELDA7 colspan='3'","","",0,"<b>" & LitEmail1 & "</b>"
							end if
							if sumaabajo>0 and sumaabajo<7 then
								DrawCelda "TDBORDECELDA7 style='width:" & espacio_anchob & "%' colspan='" & sumaabajo & "'","","",0,"&nbsp;"
							end if
						if sumaabajo=0 then
							CloseFila
						end if
	no_encabezado=0
end if
						DrawFila color_blau
							%><td class='TDBORDECELDA7' style='width:16%'>
								<b>
									<%=Hiperv(OBJPersonal,rst("dni"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("nombre"),LitVerPersonal)%>
								</b>
							</td>
							<!--<td class='TDBORDECELDA7' style='width:10%'>-->
								<!--<%=Hiperv(OBJPersonal,rst("dni"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("dni"),LitVerPersonal)%>-->
							<!--</td>--><%
							DrawCelda "TDBORDECELDA7 style='width:10%'","","",0,EncodeForHtml(trimCodEmpresa(rst("dni")))
							if null_z(opcalias)=1 then
								DrawCelda "TDBORDECELDA7 style='width:10%'","","",0,EncodeForHtml(rst("alias"))
							end if
							if null_z(opcantiguedad)=1 then
								DrawCelda "TDBORDECELDA7 style='width:8%'","","",0,EncodeForHtml(rst("antiguedad"))
							end if
							if null_z(opcnumsegsocial)=1 then
								DrawCelda "TDBORDECELDA7 style='width:12%'","","",0,EncodeForHtml(rst("ss"))
							end if
							if null_z(opcnivel)=1 then
								DrawCelda "TDBORDECELDA7 style='width:8%' align='right'","","",0,EncodeForHtml(rst("nivel"))
							end if
							if null_z(opccaja)=1 then
                                strselectDESC="select descripcion from cajas where codigo=?"
								'DrawCelda "TDBORDECELDA7 style='width:12%'","","",0,d_lookup("descripcion","cajas","codigo='" & rst("caja") & "'",session("backendlistados"))
                                 DrawCelda "TDBORDECELDA7 style='width:12%'","","",0,DLookupP1(strselectDESC,rst("caja")&"",adVarChar,10,session("backendlistados"))
							end if
							if null_z(opchoras)=1 then
								DrawCelda "TDBORDECELDA7 style='width:5%' align='right'","","",0,EncodeForHtml(rst("jornada"))
							end if
							if null_z(opcsalario)=1 then
								DrawCelda "TDBORDECELDA7 style='width:9%' align='right'","","",0,EncodeForHtml(rst("sueldo"))
							end if
							if null_z(opcvalorhextra)=1 then
								DrawCelda "TDBORDECELDA7 style='width:10%' align='right'","","",0,EncodeForHtml(rst("phextra"))
							end if
							if sumaarriba>0 then
								DrawCelda "TDBORDECELDA7 style='width:" & espacio_anchoa & "%' colspan='" & sumaarriba & "'","","",0,"&nbsp;"
							end if
						CloseFila

						if sumaabajo>0 and sumaabajo<7 then
							DrawFila color_blau
						end if
							if null_z(opcdomicilio)=1 then
								DrawCelda "TDBORDECELDA7 colspan='2'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("domicilio")))
							end if
							if null_z(opccodigopostal)=1 then
								DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("cp")))
							end if
							if null_z(opcpoblacion)=1 then
								DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("poblacion")))
							end if
							if null_z(opcprovincia)=1 then
								DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("provincia")))
							end if
							if null_z(opctelefono)=1 then
								DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("telefono")))
							end if
							if null_z(opcmovil)=1 then
								DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("telefono2")))
							end if
							if null_z(opcemail)=1 then
								DrawCelda "TDBORDECELDA7 colspan='3'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("email")))
							end if
							if sumaabajo>0 and sumaabajo<7 then
								DrawCelda "TDBORDECELDA7 style='width:" & espacio_anchob & "%' colspan='" & sumaabajo & "'","","",0,"&nbsp;"
							end if
						if sumaabajo>0 and sumaabajo<7 then
							CloseFila
						end if
if no_encabezado=1 then
					%></table><%

end if

					if sumaabajo>0 and sumaabajo<7 then
						fila=fila+2
					else
						fila=fila+1
					end if

					campo_en_el_que_estamos=""
				end if


				if campo_en_el_que_estamos="" then

					if rst("escomercial") & "">"" and si_tiene_modulo_comercial<>0 then
						%><table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
							DrawFila color_blau
								if si_tiene_modulo_comercial<>0 then
									DrawCelda "TDBORDECELDA7 align='center' colspan='3'","","",0,"<b>" & LitDatComComerModCom & "</b>"
								else
									DrawCelda "TDBORDECELDA7 align='center' colspan='3'","","",0,"<b>" & LitDatComComer & "</b>"
								end if
							CloseFila
							DrawFila color_blau
								DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitPorComBase & "</b>"
								DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitPorPen & "</b>"
								DrawCelda "TDBORDECELDA7 align='right' style='width:60%'","","",0,"&nbsp;"
							CloseFila
							DrawFila color_blau
								DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("combase")))
								DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("penalizacion")))
								DrawCelda "TDBORDECELDA7 align='right' style='width:60%'","","",0,"&nbsp;"
							CloseFila
						%></table><%
						fila=fila+1
no_encabezado=1
					end if
					campo_en_el_que_estamos=""
				end if


				if mostrarclientescom="SI" and (campo_en_el_que_estamos="" or campo_en_el_que_estamos="escliente") then

						dni=rst("dni")
						if rst("escliente") & "">"" then
							%><table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
								DrawFila color_blau
									if si_tiene_modulo_comercial<>0 then
										DrawCelda "TDBORDECELDA7 align='center' colspan='6'","","",0,"<b>" & LitCliComModCom & "</b>"
									else
										DrawCelda "TDBORDECELDA7 align='center' colspan='6'","","",0,"<b>" & LitCliCom & "</b>"
									end if
								CloseFila
								DrawFila color_blau
									DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitCodigo & "</b>"
									DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitCliente & "</b>"
									DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitPoblacion & "</b>"
									DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitTelefono & "</b>"
									DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitTCliente & "</b>"
									DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitActividad & "</b>"
								CloseFila
								continuar=0
								if not rst.eof then
									if dni=rst("dni") and rst("escliente") & "">"" then
										continuar=1
									end if
								else
									continuar=0
								end if
								while continuar=1 and not rst.eof ''fila<=MAXPAGINA
									DrawFila color_blau
										%><td class=TDBORDECELDA7>
											<%=Hiperv(OBJClientes,rst("escliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("escliente")),LitVerCliente)%>
										</td>
										<!--<td class=TDBORDECELDA7>-->
											<!--<%=Hiperv(OBJClientes,rst("escliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("rsocial"),LitVerCliente)%>-->
										<!--</td>--><%
										DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("rsocial")))
										DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("poblacionc")))
										DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("telefonoc")))
										DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("tipo_cliente")))
										DrawCelda "TDBORDECELDA7","","",0,enc.EncodeForHtmlAttribute(null_s(rst("tactividad")))
									CloseFila
									rst.movenext
									if not rst.eof then
										if dni=rst("dni") and rst("escliente") & "">"" then
											continuar=1
										else
		campo_en_el_que_estamos=""
		dni_donde_estamos=""
											continuar=0
										end if
									else
		campo_en_el_que_estamos="escliente"
		dni_donde_estamos=dni
										continuar=0
									end if
									fila=fila+1
								wend
							%></table><%
no_encabezado=1
						end if
						if not rst.eof then
							if dni<>rst("dni") then
								rst.moveprevious
							end if
						else
							rst.moveprevious
						end if
						dni=""
				end if

				if campo_en_el_que_estamos="" then
						if rst("estecnico") & "">"" then
							%><table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
								DrawFila color_blau
									DrawCelda "TDBORDECELDA7 align='center' colspan='7'","","",0,"<b>" & LitDatComTec & "</b>"
								CloseFila
								DrawFila color_blau
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitComision & "</b>"
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitHoraExtraDL & "</b>"
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitHoraExtraDF & "</b>"
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitIncentivo1 & "</b>"
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitIncentivo2 & "</b>"
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitAlmacen & "</b>"
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitVehiculo & "</b>"
								CloseFila
								DrawFila color_blau
									DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("comision")))
									DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("phextralab")))
									DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("phextrafes")))
									DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("incentivo1")))
									DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("incentivo2")))
									DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("almacen")))
									DrawCelda "TDBORDECELDA7 align='right'","","",0,EncodeForHtml(trimCodEmpresa(rst("vehiculo")))
								CloseFila
							%></table><%
							fila=fila+1
                            no_encabezado=1
						end if
					campo_en_el_que_estamos=""
				end if
				if campo_en_el_que_estamos="" then
						if rst("esoperario") & "">"" then
							%><table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
								DrawFila color_blau
									DrawCelda "TDBORDECELDA7 align='center' colspan='2'","","",0,"<b>" & LitDatComOpe & "</b>"
								CloseFila
								DrawFila color_blau
									DrawCelda "TDBORDECELDA7 align='right'","","",0,"<b>" & LitCospHora & "</b>"
									DrawCelda "TDBORDECELDA7 align='right' style='width:85%'","","",0,"&nbsp;"
								CloseFila
								DrawFila color_blau
									DrawCelda "TDBORDECELDA7 align='right'","","",0,enc.EncodeForHtmlAttribute(null_s(rst("coste_hora")))
									DrawCelda "TDBORDECELDA7 align='right' style='width:85%'","","",0,"&nbsp;"
								CloseFila
							%></table><%
							fila=fila+1
                            no_encabezado=1
						end if
					campo_en_el_que_estamos=""
				end if

				rst.movenext

				'if not rst.eof and fila<MAXPAGINA then
					%><!--<br/>--><%
				'end if
			wend
			rst.close
			if lote=lotes then
				strselect="select count(distinct dni) as suma from [" & session("usuario") & "]"
				
				'jcg 09/03/2009
				if(checkClientesBaja = "on") then
				    strselect = strselect & "where fbaja is null "
				end if
				
				rst.Open strselect, session("backendlistados"),adUseClient, adLockReadOnly
				if not rst.eof then
					total_registros=rst("suma")
				end if
				rst.close

				%><table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
					DrawFila color_fondo
						DrawCelda "TDBORDECELDA7","","",0,"<b>" & LitTotalReg & " : </b>" & EncodeForHtml(total_registros)
					CloseFila
				%></table>
			<%end if
                                                                                                                                            
			NavPaginas lote,lotes,campo,criterio,texto,2%>
			<input type="hidden" name="campo_en_el_que_estamos" value="<%=EncodeForHtml(campo_en_el_que_estamos)%>">
			<input type="hidden" name="dni_donde_estamos" value="<%=EncodeForHtml(dni_donde_estamos)%>">
		<%end if
   end if%>
</form>
<%end if
set rst=nothing
set rst2=nothing
set rstAux=nothing
set rstAux2=nothing
%>
</body>
</html>