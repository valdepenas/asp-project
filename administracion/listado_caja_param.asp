<%@ Language=VBScript %><% 
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
<%Dim CodigoHTML%>

<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTituloExt%></title>
<meta http-equiv='Content-Type' content='text/html'; charset="<%=session("caracteres")%>">
<meta http-equiv='Content-style-Type' content='text/css'>
<link rel='styleSHEET' href='../pantalla.css' media='SCREEN'>
<link rel='styleSHEET' href='../impresora.css' media='PRINT'>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->

<!--#include file="Ahoja_gastos.inc" -->

<!--#include file="../perso.inc" -->

<!--#include file="../tablasResponsive.inc" -->
<!--#include file= "../CatFamSubResponsive.inc"-->
<!--#include file="../common/poner_cajaResponsive.inc"-->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" --> 

<script language="javascript" type="text/javascript" src='../jfunciones.js'></script>
<script language="javascript" type="text/javascript">
    function GestionarAgrupaciones() {
        if (document.getElementById("mostrarSaldo").checked == true) {
            document.getElementById("agrAnotacion").checked = false;
            document.getElementById("agrTipoPago").checked = false;
            document.getElementById("agrAnotacion").disabled = true;
            document.getElementById("agrTipoPago").disabled = true;
        }
        else {
            document.getElementById("agrAnotacion").disabled = false;
            document.getElementById("agrTipoPago").disabled = false;
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
</script>

<body onload="self.status='';" class="BODY_ASP">
<%sub CalculoPaginacion()
	if lote="" then lote=1

	lotes=rst.RecordCount/MAXPAGINA
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

	rst.PageSize=MAXPAGINA
	rst.AbsolutePage=lote
end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************

    Function GetTDocumento(tdocumento, idioma)
        selectQuery = "select ISNULL(ldt.value, td.tippdoc) as descr from ilion_admin..tipo_documentos td with(NOLOCK) LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode	AND ldt.[language] = " & idioma & " where td.tippdoc = '" & tdocumento & "'"

        set rstI=server.CreateObject("ADODB.Recordset")
        rstI.cursorlocation=3
        rstI.Open selectQuery,DSNIlion
        if not rstI.EOF then
            valor = rstI("descr")
        end if
        rstI.close
        set rstI = nothing
        GetTDocumento = valor
    End Function

    Function GetLanguage(usuario, ncliente)
        selectQuery = "SELECT IDIOMA FROM CLIENTES AS C WITH(NOLOCK) LEFT JOIN CLIENTES_USERS CU ON C.NCLIENTE = CU.NCLIENTE AND CU.USUARIO = '"& usuario &"' WHERE C.NCLIENTE = '"& ncliente &"'"
        set rstI=server.CreateObject("ADODB.Recordset")
        rstI.cursorlocation=3
        rstI.Open selectQuery,DSNIlion
        if not rstI.EOF then
            idiomaUsuario = rstI("IDIOMA")
        end if
        rstI.close
        set rstI = nothing
        GetLanguage = idiomaUsuario
    End Function        
    idiomaUser = GetLanguage(session("usuario"), session("ncliente"))

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
	
	%><form name='listado_caja_param' method='post'><%

	PintarCabecera "listado_caja_param.asp"

	'Leer parámetros de la página'
  	mode=enc.EncodeForJavascript(Request.QueryString("mode"))
  	campo=limpiaCadena(Request.QueryString("campo"))
  	criterio=limpiaCadena(Request.QueryString("criterio"))
  	texto=limpiaCadena(Request.QueryString("texto"))
	lote=limpiaCadena(Request.QueryString("lote"))
	sentido=limpiaCadena(Request.QueryString("sentido"))

	dfecha = limpiaCadena(Request.form("Dfecha"))
	hfecha = limpiaCadena(Request.form("Hfecha"))
	tdocumento = limpiaCadena(Request.form("tdocumento"))
	ndocumento = limpiaCadena(Request.form("ndocumento"))
	tanotacion = limpiaCadena(Request.form("tanotacion"))
	
    tDocumentoLdt = GetTDocumento(tdocumento, idiomaUser)

	aux=enc.EncodeForJavascript(Request.Form("agrAnotacion"))
	if aux<>"0" and aux<> "-1" then
	   if ucase(aux)="ON" or ucase(aux)="VERDADERO" or ucase(aux)="TRUE" then
	      AgrAnotacion=-1
	   else
	      AgrAnotacion=0
	   end if
	 else  
	   AgrAnotacion=aux
	 end if
	
	aux=enc.EncodeForJavascript(Request.Form("agrTipoPago"))
	if aux<>"0" and aux<> "-1" then
	   if ucase(aux)="ON" or ucase(aux)="VERDADERO" or ucase(aux)="TRUE" then
	      AgrTipoPago=-1
	   else
	      AgrTipoPago=0
	   end if
	 else  
	   AgrTipoPago=aux
	 end if

	descripcion=limpiaCadena(request.form("descripcion"))
	tpago=limpiaCadena(request.form("tpago"))
	tapunte=limpiaCadena(request.form("tapunte"))
	apaisado=enc.EncodeForJavascript(request.form("apaisado"))
	caja = limpiaCadena(Request.form("caja"))

	if enc.EncodeForJavascript(request.querystring("caju"))>"" then
		cajau=limpiaCadena(request.querystring("caju"))
	else
		cajau=limpiaCadena(request.form("caju"))
	end if
	if cajau & ""="" then
		dim caju
		ObtenerParametros("listado_caja_param")
		if caju & "">"" then
			cajau=caju
		end if
	end if
	%><input type="hidden" name="caju" value="<%=EncodeForHtml(cajau)%>"/><%

    'convertimos los valores de los checkbox a cadenas de texto
	if enc.EncodeForJavascript(request.querystring("mostrarSaldo"))>"" then
		mostrarSaldo=limpiaCadena(request.querystring("mostrarSaldo"))
	else
		mostrarSaldo=limpiaCadena(request.form("mostrarSaldo"))
	end if

	if enc.EncodeForJavascript(request.querystring("mostrarSolTras"))>"" then
		mostrarSolTras=limpiaCadena(request.querystring("mostrarSolTras"))
	else
		mostrarSolTras=limpiaCadena(request.form("mostrarSolTras"))
	end if
	
	if mostrarSolTras="on" then mostrarSolTras="1"

	'cag
	if enc.EncodeForJavascript(request.querystring("mostrarGasto"))>"" then
		mostrarGasto=limpiaCadena(request.querystring("mostrarGasto"))
	else
		mostrarGasto=limpiaCadena(request.form("mostrarGasto"))
	end if
	'fin cag
	
	lotes=1

	si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
	si_tiene_modulo_produccion=ModuloContratado(session("ncliente"),ModProduccion)

  	set conn = Server.CreateObject("ADODB.Connection")
  	conn.open session("backendlistados")

  	set rstAux = Server.CreateObject("ADODB.Recordset")
  	set rst = Server.CreateObject("ADODB.Recordset")

	WaitBoxOculto LitEsperePorFavor

	Alarma "listado_caja_param.asp"
	%><hr/><%
'****************************************************************************************************************'
	if (mode="add") then
                DrawDiv "col-lg-2 col-md-3 col-sm-3 col-xs-6","",""
                DrawLabel "","",LitDesdeFecha
                DrawInput "'width70 dfecha'", "", "Dfecha", iif(TmpDfecha>"",TmpDfecha,"01/01/" & year(date)), ""
                DrawCalendar "Dfecha"
                CloseDiv
                DrawDiv "col-lg-2 col-md-3 col-sm-3 col-xs-6","",""
                DrawLabel "","",LitDesdeHora
                DrawInput "width50","left","Dhora","00:00","id='Dhora'"
                CloseDiv%><script type="text/javascript">
                $("#Dhora").keypress(function (e) {
                    if (isTextSelected(this)) {
                            return;
                        }
                        if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) 
                        {
                            return false;
                        }
                        else
                        {
                            var reg = /[0-9]/;
                            if (this.value.length == 2 && reg.test(this.value))
                                this.value = this.value + ":";
                            if (this.value.length > 4)
                            return false;
                        }
                });
                $(".dfecha").keypress(function (e) {
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
				
                DrawDiv "col-lg-2 col-md-3 col-sm-3 col-xs-6","",""
                DrawLabel "","",LitHastaFecha    
                DrawInput "'width70 hfecha'", "", "Hfecha", iif(TmpHfecha>"",TmpHfecha,day(date) & "/" & month(date) & "/" & year(date)), ""
                DrawCalendar "Hfecha"
                CloseDiv
                DrawDiv "col-lg-2 col-md-3 col-sm-3 col-xs-6","",""
                DrawLabel "","",LitHastaHora
                DrawInput "width50","left","Hhora","23:59","id='Hhora'"
                CloseDiv%><script type="text/javascript">
                $("#Hhora").keypress(function (e) {
                    if (isTextSelected(this)) {
                            return;
                        }
                        if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) 
                        {
                            return false;
                        }
                        else
                        {
                            var reg = /[0-9]/;
                            if (this.value.length == 2 && reg.test(this.value))
                                this.value = this.value + ":";
                            if (this.value.length > 4)
                            return false;
                        }
                });
                $(".hfecha").keypress(function (e) {
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
		
				defecto=""
                Drawdiv "1","",""
                DrawLabel "txtMandatory","",LitCajaM
				poner_cajasResponsive1 "width60",defecto,"caja",0,"codigo","descripcion","","",poner_comillas(cajau)
                CloseDiv
				strtipdoc=" where "
				if si_tiene_modulo_mantenimiento=0 then
					strtipdoc=strtipdoc & " TIPPDOC<>'INCIDENCIA' and TIPPDOC<>'ORDEN' and TIPPDOC<>'PARTE DE TRABAJO' and "
				end if
				if si_tiene_modulo_produccion=0 then
					strtipdoc=strtipdoc & " TIPPDOC<>'ORDEN DE FABRICACION' and TIPPDOC<>'NOTA DE FABRICACION' and "
				end if
				if strtipdoc=" where " then
					strtipdoc=""
				else
					strtipdoc=mid(strtipdoc,1,len(strtipdoc)-4)
				end if

                '----------------------------------------------------------------------------------------
                'Nueva forma de obtener la descripcion de los tipos de documentos de la tabla lit_typedoc
                '----------------------------------------------------------------------------------------
                set conn = Server.CreateObject("ADODB.Connection")        
                set command =  Server.CreateObject("ADODB.Command")
                conn.open DSNIlion
                command.ActiveConnection = conn
                command.CommandTimeout = 0
                command.CommandText = "ComboBoxDocTypes"
                command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                command.NamedParameters = True 
                command.Parameters.Append command.CreateParameter("@inlist",adVarChar,adParamInput,5,"todos")
                command.Parameters.Append command.CreateParameter("@outlist",adVarChar,adParamInput,1,"")
                command.Parameters.Append command.CreateParameter("@user",adVarChar,adParamInput,50,session("usuario"))
                command.Parameters.Append command.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))
                set rstTD = Server.CreateObject("ADODB.Recordset")
                set rstTD = command.Execute
                if not rstTD.eof then
		                DrawSelectCelda "","","","0",LitTipoDocumento,"tdocumento",rstTD,"","tippDoc","descripcion","",""
                end if			
                rstTD.close
                set command=nothing
                conn.close
			
                DrawDiv "1", "", ""
                DrawLabel "", "", LitAnotacion%><select class='width60' name="tanotacion">
						<option value="ENTRADA"><%=LitEntrada%></option>
						<option value="SALIDA"><%=LitSalida2%></option>
						<option value="" selected></option>
					</select><%
                CloseDiv
                EligeCelda "check","add","","","",0,LitAgrupar,"agrAnotacion",0,""
                CloseDiv
			
                rstAux.cursorlocation=3
				rstAux.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
				DrawSelectCelda "CELDA","190","",0,LitTipoPago,"tpago",rstAux,"","codigo","descripcion","",""
				rstAux.close
			
                DrawDiv "1", "", ""
                DrawLabel "", "", LitAgrupar			
                DrawCheck "", "", "agrTipoPago", ""
                CloseDiv
			
                 EligeCelda "input","add","left","","",0,LitDescripcion,"descripcion","",""
				
                 EligeCelda "input","add","left","","",0,LitNDocumento,"ndocumento","",""
			
                rstAux.cursorlocation=3
			 	rstAux.open "select codigo,descripcion from tipo_apuntes with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("backendlistados")
                DrawDiv "1", "", "ocultartapunte"
                DrawLabel "", "", LitTipoApunte%><select class="width60" name="tapunte">
					    <%while not rstAux.eof%>
							<option <%=iif(tapunte=rstAux("codigo"),"selected","")%> value="<%=EncodeForHtml(rstAux("codigo"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option><%
							rstAux.movenext
						wend%>
						<option <%=iif(tapunte="","selected","")%> value=""></option>
					</select>
				<%rstAux.close
                CloseDiv
                DrawDiv "1", "", ""
                DrawLabel "","",LitMostrarSaldo%><input type='checkbox' name='mostrarSaldo' onclick="javascript: GestionarAgrupaciones();"/><%
                CloseDiv
                EligeCelda "check","add","","","",0,LitMosNoTras,"mostrarSolTras",0, ""  
                EligeCelda "check","add","","","",0,LitMostrarGasto,"mostrarGasto",0, ""
	            EligeCelda "check","add","","","",0,LitApaisado,"apaisado",0, ""
		    	%><hr/><%
'****************************************************************************************************************'
		'Mostrar el listado.'
	elseif mode="browse" then

            set conn = Server.CreateObject("ADODB.Connection")
            initial_catalogC=encontrar_datos_dsn(session("backendlistados"),"Initial Catalog=")

	        donde=inStr(1,DSNImport,"Initial Catalog=",1)
	        donde_fin=InStr(donde,DSNImport,";",1)

	        if donde_fin=0 then
		        donde_fin=len(DSNImport)
	        end if
	        cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))

	        dsnCliente=cadena_dsn_final

	        conn.open dsnCliente

			MB=""''d_lookup("codigo", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("backendlistados"))
            n_decimales=0''null_z(d_lookup("ndecimales", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("backendlistados")))
			AbreviaturaMB=""''d_lookup("abreviatura", "divisas", "moneda_base<>0 and codigo like '" & session("ncliente") & "%'", session("backendlistados"))
            rst.cursorlocation=3
            rst.open "select codigo,ndecimales,abreviatura from divisas with(NOLOCK) where codigo like '" & session("ncliente") & "%' and moneda_base<>0",session("backendlistados")
            if not rst.eof then
                n_decimales=rst("ndecimales")
                MB=rst("codigo")
                AbreviaturaMB=rst("abreviatura")
            end if
            rst.close

			MAXPAGINA=0''d_lookup("maxpagina", "limites_listados", "item='086'", DSNIlion)
			MAXPDF=0''d_lookup("maxpdf", "limites_listados", "item='086'", DSNIlion)
            rst.cursorlocation=3
            rst.open "select maxpagina,maxpdf from limites_listados with(NOLOCK) where item='086'",DSNIlion
            if not rst.eof then
                MAXPAGINA=rst("maxpagina")
                MAXPDF=rst("maxpdf")
            end if
            rst.close
			%><input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'/><%
			
			PorCliente="NO"
			PorGasto="NO"

			%><input type='hidden' name='Dfecha' value='<%=EncodeForHtml(dfecha)%>'/>
			<input type='hidden' name='Hfecha' value='<%=EncodeForHtml(hfecha)%>'/>
            <input type='hidden' name='Dhora' value='<%=EncodeForHtml(dhora)%>'/>
			<input type='hidden' name='Hhora' value='<%=EncodeForHtml(hhora)%>'/>
			<input type='hidden' name='caja' value='<%=EncodeForHtml(caja)%>'/>
			<input type='hidden' name='tdocumento' value='<%=EncodeForHtml(tDocumentoLdt)%>'/>
			<input type='hidden' name='ndocumento' value='<%=EncodeForHtml(ndocumento)%>'/>
			<input type='hidden' name='tanotacion' value='<%=EncodeForHtml(tanotacion)%>'/>
			<input type='hidden' name='agrAnotacion' value='<%=EncodeForHtml(agrAnotacion)%>'/>
		    <input type='hidden' name='agrTipoPago' value='<%=EncodeForHtml(agrTipoPago)%>'/>
			<input type='hidden' name='descripcion' value='<%=EncodeForHtml(descripcion)%>'/>
			<input type='hidden' name='tpago' value='<%=EncodeForHtml(tpago)%>'/>
			<input type='hidden' name='tapunte' value='<%=EncodeForHtml(tapunte)%>'/>
			<input type='hidden' name='mostrarSaldo' value='<%=EncodeForHtml(mostrarSaldo)%>'/>
			<input type='hidden' name='mostrarSolTras' value='<%=EncodeForHtml(mostrarSolTras)%>'/>
		    <%'cag%>
			<input type='hidden' name='mostrarGasto' value='<%=EncodeForHtml(mostrarGasto)%>'/>
			<%'fin cag%>
			<input type='hidden' name='apaisado' value='<%=EncodeForHtml(apaisado)%>'/>
			<font class="cab"><b><%=LitCajaM%> : </b></font><font class="cab"><%=d_lookup("descripcion","cajas","codigo='" & caja & "'",session("backendlistados"))%></b></font><br/><%
			if tanotacion>"" then
				%><font class="cab"><b><%=LitAnotacion%> : </b></font><font class="cab"><%=EncodeForHtml(tanotacion)%></font><br/><%
			end if
			if tpago>"" then
				%><font class="cab"><b><%=LitTipoPago%> : </b></font><font class="cab"><%=d_lookup("descripcion","tipo_pago","codigo='" & tpago & "'",session("backendlistados"))%></b></font><br/><%
			end if
			if tapunte>"" then
				%><font class="cab"><b><%=LitTipoApunte%> : </b></font><font class="cab"><%=d_lookup("descripcion","tipo_apuntes","codigo='" & tapunte & "'",session("backendlistados"))%></b></font><br/><%
			end if
			if tdocumento>"" then
				%><font class="cab"><b><%=LitTipoDocumento%> : </b></font><font class="cab"><%=EncodeForHtml(tDocumentoLdt)%></b></font><br/><%
			end if
			if ndocumento>"" then
				%><font class="cab"><b><%=LitNDocumento%> : </b></font><font class="cab"><%=EncodeForHtml(ndocumento)%></b></font><br/><%
			end if
			if descripcion>"" then
				%><font class="cab"><b><%=LitDescripcion%> : </b></font><font class="cab"><%=EncodeForHtml(descripcion)%></b></font><br/><%
			end if
			if dfecha>"" then
				%><font class="cab"><b><%=LitDesdeFecha%> : </b></font><font class="cab"><%=EncodeForHtml(dfecha)%></b></font><br/><%
			end if
			if hfecha>"" then
				%><font class="cab"><b><%=LitHastaFecha%> : </b></font><font class="cab"><%=EncodeForHtml(hfecha)%></b></font><br/><%
			end if
			if mostrarSolTras="1" then
				%><font class="cab"><b><%=LitMosNoTras%></b></font><br/><%
			end if
			
			saldoReal=d_lookup("saldo","cajas","codigo='" & caja & "'",session("backendlistados"))
			saldoReal=formatnumber(saldoReal,n_decimales,-1,0,-1)
			if saldoReal>=0 then
				colorSaldoReal=color_azul
			else
				colorSaldoReal=color_rojo
			end if
            
			strwhere="where caja='" & caja & "' and "
			if dfecha>"" then strwhere=strwhere & " fecha>='" & dfecha & " 00:00:00' and"
			if hfecha>"" then strwhere=strwhere & " fecha<='" & hfecha & " 23:59:59' and"
			if tpago>"" then strwhere=strwhere & " medio='" & tpago & "' and"
			if tapunte>"" then strwhere=strwhere & " c.tapunte='" & tapunte & "' and"
			if tdocumento>"" then strwhere=strwhere & " tdocumento='" & tdocumento & "' and"
			if ndocumento>"" then
				if tdocumento="ALBARAN DE PROVEEDOR" then
					'ndoc_aux=d_lookup("nalbaran","albaranes_pro","nalbaran_pro='" & ndocumento & "'",session("backendlistados"))
					rst.cursorlocation=3
					'ega 18/03/2008 like con el código del cliente
					rst.open "select nalbaran from albaranes_pro with(nolock) where nalbaran = '"& session("ncliente") &"%' and nalbaran_pro like '%" & ndocumento & "%'",session("backendlistados")
					lista=""
					while not rst.eof
						lista=lista & rst("nalbaran") & "','"
						rst.movenext
					wend
					if lista & "">"" then
						lista="('" & mid(lista,1,len(lista)-2) & ")"

					else
						lista="('@#@#@#xx2er---|||')"
					end if
					rst.close
					strwhere=strwhere & " ndocumento in " & lista & " and"
				elseif tdocumento="FACTURA DE PROVEEDOR" then
					'ndoc_aux=d_lookup("nfactura","facturas_pro","nfactura_pro='" & ndocumento & "'",session("backendlistados"))
					rst.cursorlocation=3
					'ega 18/03/2008 like con el código del cliente
					rst.open "select nfactura from facturas_pro with(nolock) where nfactura='"& session("ncliente") &"%' and nfactura_pro like '%" & ndocumento & "%'",session("backendlistados")
					lista=""
					while not rst.eof
						lista=lista & rst("nfactura") & "','"
						rst.movenext
					wend
					if lista & "">"" then
						lista= "('" & mid(lista,1,len(lista)-2) & ")"
					else
						lista="('@#@#@#xx2er---|||')"
					end if
					rst.close
					strwhere=strwhere & " ndocumento in " & lista & " and"
				else
					strwhere=strwhere & " ndocumento like '%" & ndocumento & "%' and"
				end if
			end if

			if descripcion>"" then strwhere=strwhere & " c.descripcion like '%" & descripcion & "%' and"

			if mostrarSolTras="1" then
				strwhere=strwhere & " ntraspaso is null and"
			end if

			if tanotacion>"" then 'Se ha seleccionado anotacion. Se ignora agrupacion
				if AgrTipoPago<>0 and tpago&""="" then 'Se agrupa por el tipo de pago
					strwhere=strwhere & " tanotacion='" & tanotacion & "' and"
					strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
					'ega 18/03/2008 with(nolock) y likes con código del cliente
					'ega 27/03/2008 si el gasto es nulo, se le pone 0, si el medio es nulo lo pone a vacio
					seleccion="select isnull(tipo_pago.gasto,0) as gasto,caja,isnull(medio,'') as medio,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,ndocumento,ndocumento_pro,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                    ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                    seleccion=seleccion & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                    seleccion=seleccion & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                    seleccion=seleccion & " from caja c with(nolock) "
					seleccion=seleccion & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                    seleccion=seleccion & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                    seleccion=seleccion & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                    seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                    seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = "& idiomaUser
                    seleccion=seleccion & " " & strwhere
                    seleccion=seleccion & " order by tipo_pago.descripcion,fecha"
					rst.cursorlocation=3
''response.write("el seleccion 1 es-" & seleccion & "-<br>")
					rst.open seleccion,dsnCliente
					if not rst.eof then
						%><font class="cab"><b><%=LitApuntes%> : </b></font><font class="cab"><%=rst.recordcount%></font><br/>
						<font class="cab"><b><%=LitSaldoReal%> : </b></font><font style='color: <%=colorSaldoReal%>' class=cab><%=EncodeForHtml(saldoReal)%>&nbsp; <%=EncodeForHtml(AbreviaturaMB)%></b></font><br/><hr/><br/>
						<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'/><%
						'Calculos de páginas--------------------------
						CalculoPaginacion()
						NavPaginas lote,lotes,campo,criterio,texto,1
						'-----------------------------------------%><br/>
						<table width='100%' style='border-collapse: collapse;'>
							<thead>
								<tr bgcolor="<%=color_fondo%>">
									<td class="tdbordeCELDA7"><b><%=LitTipoPago%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitFecha%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitDescripcion%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitTipoDocumento%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitNDocumento%></b></td>
									<!-- cag --> <%
									if  mostrarGasto="on" then%>
											<td class="tdbordeCELDA7"><b><%=LitGasto%></b></td> <%
									end if %>
									<!-- fin cag-->
									<td class="tdbordeCELDA7"><b><%=LitTipoApunte%></b></td>
									<td class="tdbordeCELDA7" align="right"><b><%=LitImporte%></b></td>
									
								</tr>
							</thead>
							<tbody><%
								Suma=0
								SubTotal=0
								SumaGasto=0
								SubTotalGasto=0
								MedioAnterior=""
								fila=1
								while not rst.eof and fila<=MAXPAGINA%>
									<%CheckCadena rst("caja")%>
									<tr><%
										dato="&nbsp;"
										if rst("medio")<>MedioAnterior then
											If MedioAnterior<>"" then
												SubTotal=0
												SubTotalGasto=0
												'ega 18/03/2008 with(nolock) y likes
												'ega 27/03/2008 si el gasto es nulo, se le pone 0
												strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                                                ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                                strselect=strselect & " ,isnull(c.factcambio,divisas.factcambio) as factcambio "
                                                strselect=strselect & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                                                strselect=strselect & " from caja c with(nolock)"
												strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                                strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                                strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                                strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                                strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = "& idiomaUser
                                                strselect=strselect & " "  & strwhere & " and medio='" & MedioAnterior & "' "
                                                strselect=strselect & " order by tanotacion desc,fecha"
												rstAux.cursorlocation=3
''response.write("el seleccion 2 es-" & strselect & "-<br>")
												rstAux.open strselect,dsnCliente
												for k=1 to rstAux.recordcount
													Nmedio=rstAux("nmedio")
													if rstAux("tanotacion")="ENTRADA" then
                                                        if rstAux("factcambio") & "">"" then
														    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
														    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                        else
														    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
														    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                        end if
													else
                                                        if rstAux("factcambio") & "">"" then
														    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
														    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                        else
														    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
														    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB) '-
                                                        end if
													end if
													rstAux.movenext
												next
												rstAux.close
												if mostrarGasto="on" then %>
													<td class="tdbordeCELDA7" colspan="4">&nbsp;</td>
													<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>													
													<td class="tdbordeCELDA7" align="right"><div><b><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
													<td class="tdbordeCELDA7">&nbsp;</td>
													<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>												
												<%else
												%><td class="tdbordeCELDA7" colspan="5">&nbsp;</td>
													<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=Nmedio%> :&nbsp;</b></td>
													<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												<%end if%>
												</tr>
												<tr> <%'Fila de separacion%>
												<% if mostrarGasto="on" then%>
													<td class="tdbordeCELDA7" colspan="8">&nbsp;</td>									
												<%else%>
													<td class="tdbordeCELDA7" colspan="7">&nbsp;</td>
												<%end if%>
												</tr><%
												SubTotal=0
												SubTotalGasto=0
											end if
											dato=rst("Nmedio")
										end if
										if rst("tanotacion")="ENTRADA" then
											color=color_azul
										else
											color=color_rojo
										end if%>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(dato)%></td>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("fecha"))%></td>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descripcion"))%></td>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descr"))%></td>
										<%if rst("ndocumento_pro")&"">"" then%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("ndocumento_pro"))%></td>
										<%else%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(trimCodEmpresa(rst("ndocumento")))%></td>
										<%end if%>
										<%'cag
	 									if  mostrarGasto="on" then
											impGasto=rst("Importe")*rst("gasto")/100
											if rst("tanotacion")="ENTRADA" then%>
												<td class="tdbordeCELDA7" align="right"><%=formatnumber(impGasto,rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
											<%else%>
												<td class="tdbordeCELDA7" align="right"><%=formatnumber(abs(impGasto),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>												
											<%end if
										 end if
										'fin cag
										%>										
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("tapuntdesc"))%></td>
										<%if rst("tanotacion")="ENTRADA" then%>
											<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></div></td>
										<%else%>
											<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'>-<%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></div></td>
										<%end if%>
									</tr><%
									MedioAnterior=rst("medio")
									rst.movenext
									fila=fila+1
								wend
								if not rst.eof then 'escritos el número máximo de filas pero quedan registros
									'Se ha llegado al final de las filas permitidas justo
									'cuando se cambiaba de pago.Escribir subtotal
									if rst("medio")<>MedioAnterior then
										SubTotal=0
										SubTotalGasto=0
										'ega 18/03/2008 with(nolock) y likes
										'ega 27/03/2008 si el gasto es nulo, se le pone 0
										strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                                        ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                        strselect=strselect & " ,isnull(c.factcambio,divisas.factcambio) as factcambio "
                                        strselect=strselect & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr, ISNULL(ldt.value,c.tdocumento) as descr "
                                        strselect=strselect & " from caja c with(nolock) "
										strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                        strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                        strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                        strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                        strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser
                                        strselect=strselect & " " & strwhere & " and medio='" & MedioAnterior & "' "
                                        strselect=strselect & " order by tanotacion desc,fecha"
										rstAux.cursorlocation=3
''response.write("el seleccion 3 es-" & strselect & "-<br>")
                              			rstAux.open strselect,dsnCliente
										''rstAux.open strselect
										for k=1 to rstAux.recordcount
											Nmedio=rstAux("nmedio")
											if rstAux("tanotacion")="ENTRADA" then
                                                if rstAux("factcambio") & "">"" then
												    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
												    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                end if
											else
                                                if rstAux("factcambio") & "">"" then
												    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
												    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB) '-
                                                end if
											end if
											rstAux.movenext
										next
										rstAux.close
										color=color_azul
										'cag
										if mostrarGasto="on" then%>
										<tr>
											<td class="tdbordeCELDA7" colspan="4">&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><div><b><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>										
											<td class="tdbordeCELDA7">&nbsp;</td>
											<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
										</tr>										
										<%else%>
										<tr>
											<td class="tdbordeCELDA7" colspan="5">&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
										</tr><%
										end if
									end if
								else 'Se llegó al final de los registros
									'Fila para el Subtotal
									SubTotal=0
									Suma=0
									SubTotalGasto=0
									SumaGasto=0
									rstAux.cursorlocation=3
									'rstAux.open "select tanotacion,Fecha,caja.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as medio,Importe,divisas.abreviatura,divisa,ndecimales from caja,divisas,tipo_pago " & strwhere & " and medio='" & MedioAnterior & "' order by tanotacion desc,fecha",session("backendlistados")
									rst.movefirst
									for k=1 to rst.recordcount
										Nmedio=rst("nmedio")
										if rst("medio")=MedioAnterior then
											if rst("tanotacion")="ENTRADA" then
                                                if rst("factcambio") & "">"" then
												    SubTotal=SubTotal + rst("importe")/rst("factcambio")
												    Suma=Suma + rst("importe")/rst("factcambio")
												    SubTotalGasto=SubTotalGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    SubTotal=SubTotal + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    Suma=Suma + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
											else
                                                if rst("factcambio") & "">"" then
												    SubTotal=SubTotal - rst("importe")/rst("factcambio")
												    Suma=Suma - rst("importe")/rst("factcambio")
												    SubTotalGasto=SubTotalGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    SubTotal=SubTotal - CambioDivisa(rst("importe"),rst("divisa"),MB)
												    Suma=Suma - CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB) '-
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)'-
                                                end if
											end if
										else
											if rst("tanotacion")="ENTRADA" then
                                                if rst("factcambio") & "">"" then
												    Suma=Suma + rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
											else
                                                if rst("factcambio") & "">"" then
												    Suma=Suma - rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma - CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB) '-
                                                end if
											end if
										end if
										rst.movenext
									next%>
									<tr>
										<%if mostrarGasto="on" then%>
											<td class="tdbordeCELDA7" colspan=4>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>
										<%else%>
											<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>										
								        <%end if%>
										
										<%if SubTotal=0 then 
										    if mostrarGasto="on" then %>
												<td class="tdbordeCELDA7" align="right"><b><div ><%=formatnumber(abs(SumaGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>											
											<%else%>
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%end if%>
										<%else 
										    if mostrarGasto="on" then %>
												<td class="tdbordeCELDA7" align="right"><b><div><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>											
											<%else%>
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%end if
										 end if%>
									</tr>
									<tr> <%'Fila de separacion%>
									<% if mostrarGasto="on" then%>
										<td class="tdbordeCELDA7" colspan=8>&nbsp;</td>									
									<%else%>
										<td class="tdbordeCELDA7" colspan=7>&nbsp;</td>
									<%end if%>
									</tr>
									<tr> <%'Fila para el total
										if tanotacion="ENTRADA" then
											color=color_azul
										else
											color=color_rojo
										end if
										if mostrarGasto="on" then %>
											<td class="tdbordeCELDA7" colspan=4>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div><%=formatnumber(abs(SumaGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
											<td class="tdbordeCELDA7">&nbsp;</td>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>										
										<%else%>
											<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
										<%end if%>
									</tr><%
									'EQUIVALENCIA EN PTAS
									if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
										DrawFila color_blau
										if tanotacion="ENTRADA" then
											color=color_azul
										else
											color=color_rojo
										end if
										if mostrarGasto="on" then%>
											<td class="tdbordeCELDA7" colspan=4>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div><%=cstr((formatnumber(CambioDivisa(abs(SumaGasto),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
											<td class="tdbordeCELDA7">&nbsp;</td>										
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
										<%else%>
											<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
										<%end if
										CloseFila
									end if
								end if%>
							</tbody>
						</table><br/><%
						NavPaginas lote,lotes,campo,criterio,texto,2

					else%>
						<script language="javascript" type="text/javascript">
						    alert("<%=LitNoExisteDatos%>");
						    parent.window.frames["botones"].document.location = "listado_caja_param_bt.asp?mode=add";
						    document.location = "listado_caja_param.asp?mode=add&caju=<%=enc.EncodeForJavascript(cajau)%>";
						</script><%
					end if
					rst.close
				else 'Se ha seleccionado anotacion pero no se agrupa por el tipo de pago
					strwhere=strwhere & " tanotacion='" & tanotacion & "' and"
					strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
					'cag
					'seleccion="select caja,nanotacion,tanotacion,Fecha,c.descripcion as Descripcion,isnull(tdocumento,'') as tdocumento,isnull(ndocumento,'') as ndocumento,isnull(ndocumento_pro,'') as ndocumento_pro,tipo_pago.descripcion as medio,Importe,divisas.abreviatura,divisa,ndecimales,factcambio,isnull(tipo_apuntes.descripcion,'') as tapuntdesc,0 as saldo from caja c "
					'ega 18/03/2008 with(nolock) y likes con el codigo del cliente
					'ega 27/03/2008 si el gasto es nulo, se le pone 0
					seleccion="select isnull(tipo_pago.gasto,0) as gasto,caja,nanotacion,tanotacion,Fecha,c.descripcion as Descripcion,isnull(tdocumento,'') as tdocumento,isnull(ndocumento,'') as ndocumento,isnull(ndocumento_pro,'') as ndocumento_pro,isnull(tipo_pago.descripcion,'') as medio,Importe,divisas.abreviatura,divisa,ndecimales "
                    ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                    seleccion=seleccion & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                    
                    seleccion=seleccion & " ,isnull(tipo_apuntes.descripcion,'') as tapuntdesc,0 as saldo from caja c with(nolock) "
					'fin cag
					seleccion = seleccion & "left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%'  left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' " & strwhere & "  order by fecha,nanotacion"

    				set command =  Server.CreateObject("ADODB.Command")
		            
                    command.ActiveConnection =conn
                    command.CommandTimeout = 0
                    command.CommandText="ListadoExtractoCaja"
                    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    command.Parameters.Append command.CreateParameter("@NomTabla",adVarChar,adParamInput,50,session("usuario"))
                    command.Parameters.Append command.CreateParameter("@seleccion",adVarChar,adParamInput,3000,seleccion)
                    command.Parameters.Append command.CreateParameter("@mostrarSaldo",adVarChar,adParamInput,10,mostrarSaldo)
                    command.Parameters.Append command.CreateParameter("@fechaInicio",adDate,adParamInput,,dfecha)
                    command.Parameters.Append command.CreateParameter("@saldoInicial",adCurrency ,adParamOutput,,SaldoInicial)
                    
                    command.Execute
                    
                    SaldoInicial = Command.Parameters("@saldoInicial").Value
        			set command=nothing

                    seleccion="select c.* "
                    ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                    seleccion=seleccion & " ,isnull(c.factcambio,divisas.factcambio) as factcambio "
                    seleccion=seleccion & " , ISNULL(ldt.value,c.tdocumento) as descr "
                    seleccion=seleccion & " from [egesticet].[" & session("usuario") & "] as c "
                    seleccion=seleccion & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                    seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                    seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser
                    seleccion=seleccion & " order by c.fecha"
		            rst.cursorlocation=3
		            rst.open seleccion,dsnCliente

					if not rst.eof then
						%><font class="cab"><b><%=LitApuntes%> : </b></font><font class="cab"><%=rst.recordcount%></font><br/>
						<font class="cab"><b><%=LitSaldoReal%> : </b></font><font style='color: <%=colorSaldoReal%>' class=cab><%=EncodeForHtml(saldoReal)%>&nbsp; <%=EncodeForHtml(AbreviaturaMB)%></b></font><br/><hr/><br/>
						<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'/><%
						'Calculos de páginas--------------------------
						CalculoPaginacion()
						NavPaginas lote,lotes,campo,criterio,texto,1
						'-----------------------------------------%><br/>
						<table width='100%' style='border-collapse: collapse;'>
							<thead>
								<tr bgcolor="<%=color_fondo%>">
									<td class="tdbordeCELDA7"><b><%=LitFecha%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitAnotacion%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitDescripcion%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitTipoDocumento%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitNDocumento%></b></td>
									<td class="tdbordeCELDA7"><b><%=LitTipoPago%></b></td>
									<!-- cag --> <%
									if  mostrarGasto="on" then%>
											<td class="tdbordeCELDA7"><b><%=LitGasto%></b></td> <%
									end if %>
									<!-- fin cag-->									
									<td class="tdbordeCELDA7"><b><%=LitTipoApunte%></b></td>
									<td class="tdbordeCELDA7" align="right"><b><%=LitImporte%></b></td>
									<%
									if mostrarSaldo="on" then
										%>
										<td class="tdbordeCELDA7" align="right"><b><%=LitSaldo%>(<%=EncodeForHtml(AbreviaturaMB)%>)</b></td>
										<%
									end if
									%>
								</tr>
							</thead>
							<tbody><%
								Suma=0
								fila=1
								' Ponemos la linea de saldo inicial'
								if mostrarSaldo="on" and lote=1 then
									if SaldoInicial>=0 then
										colorSaldoInicial=color_azul
									else
										colorSaldoInicial=color_rojo
									end if
									%>
									<tr>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7"><%=LitSaldoInicial%></td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7" style='color:<%=colorSaldoInicial%>' align="right"><%=formatnumber(SaldoInicial,rst("ndecimales"),-1,0,-1)%></td>
									</tr>
								<%
								end if
								' FIN de la linea de saldo inicial'
								while not rst.eof and fila<=MAXPAGINA
									%>
									<tr>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("fecha"))%></td>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("tanotacion"))%></td>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descripcion"))%></td>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descr"))%></td><%
										if rst("ndocumento_pro")&"">"" then%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("ndocumento_pro"))%></td>
										<%else%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(trimCodEmpresa(rst("ndocumento")))%></td>
										<%end if%>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("medio"))%></td> 
										<%'cag
	 									if  mostrarGasto="on" then
											impGasto=rst("Importe")*rst("gasto")/100
											if rst("tanotacion")="ENTRADA" then%>
												<td class="tdbordeCELDA7" align="right"><%=formatnumber(impGasto,rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
											<%else%>
												<td class="tdbordeCELDA7" align="right"><%=formatnumber(abs(impGasto),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>											
											<%end if%>
										<%end if
										'fin cag
										%>
										<td class="tdbordeCELDA7"><%=rst("tapuntdesc")%></td>
										<%if rst("tanotacion")="ENTRADA" then%>
											<td class="tdbordeCELDA7" style='color:<%=color_azul%>' align="right"><%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
										<%else%>
											<td class="tdbordeCELDA7" style='color:<%=color_rojo%>' align="right">-<%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
										<%end if
										if mostrarSaldo="on" then
											if rst("saldo")>=0 then%>
												<td class="tdbordeCELDA7" style='color:<%=color_azul%>' align="right"><%=formatnumber(rst("saldo"),n_decimales,-1,0,-1)%></td>
											<%else%>
												<td class="tdbordeCELDA7" style='color:<%=color_rojo%>' align="right"><%=formatnumber(rst("saldo"),n_decimales,-1,0,-1)%></td>
											<%end if
										end if%>
									</tr><%
									saldoFinal=rst("saldo")
									rst.movenext
									fila=fila+1
								wend
								if rst.eof then 'Se llegó al final de los registros%>
									<tr><% 'Fila para el total
										Suma=0
										SumaGasto=0
										colVacias=6
                                        rstAux.cursorlocation=3
''response.write("el seleccion 4 es-" & seleccion & "-<br>")
										rstAux.open seleccion,dsnCliente
										for k=1 to rstAux.recordcount
											if rstAux("tanotacion")="ENTRADA" then
                                                if rstAux("factcambio") & "">"" then
                                                    Suma=Suma + rstAux("importe")/rstAux("factcambio")
                                                    SumaGasto=SumaGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    Suma=Suma + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    'cag
												    SumaGasto=SumaGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
												    'fin cag
                                                end if
											else
                                                if rstAux("factcambio") & "">"" then
                                                    Suma=Suma - rstAux("importe")/rstAux("factcambio")
                                                    SumaGasto=SumaGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    Suma=Suma - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    'cag
												    SumaGasto=SumaGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
												    'fin cag
                                                end if
											end if
											rstAux.movenext
										next
										rstAux.close%>
										<!-- cag -->
										<!-- <td class="tdbordeCELDA7" colspan="<%=colVacias%>">&nbsp;</td> -->
										<%if mostrarGasto="on" then%>										
											<td class="tdbordeCELDA7" colspan="<%=colVacias-1%>">&nbsp;</td>
										  <%else %>
											<td class="tdbordeCELDA7" colspan="<%=colVacias%>">&nbsp;</td>
										<%end if%>
										<!-- fin cag-->
										<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitTotal%> :&nbsp;</b></big></td>
										<!-- cag -->
										<%if mostrarGasto="on" then%>
										<td class="tdbordeCELDA7" align="right"><big><b><div><%=formatnumber(abs(SumaGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></div></b></big></td>									
										<td class="tdbordeCELDA7">&nbsp;</td>
										<%end if%>
										<!-- fin cag-->										
										<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=iif(Suma>=0,color_azul,color_rojo)%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></div></b></big></td>
										<%if mostrarSaldo="on" then%>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=iif(saldoFinal>=0,color_azul,color_rojo)%>'><%=formatnumber(saldoFinal,n_decimales,-1,0,-1)%></div></b></big></td>
										<%end if%>
									</tr><%
									'EQUIVALENCIA EN PTAS'
									if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
										DrawFila color_blau %>
										    <!--cag -->
											<!-- <td class="tdbordeCELDA7" colspan=<%=colVacias%>>&nbsp;</td> -->
											<%if mostrarGasto="on" then%>										
												<td class="tdbordeCELDA7" colspan="<%=colVacias-1%>">&nbsp;</td>
										    <%else %>
												<td class="tdbordeCELDA7" colspan="<%=colVacias%>">&nbsp;</td>
											<%end if%>
											<!-- fin cag-->											
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitTotal + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
											<!-- cag -->
											<%if mostrarGasto="on" then%>
												<td class="tdbordeCELDA7" align="right"><big><b><div><%=cstr((formatnumber(CambioDivisa(abs(SumaGasto),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
												<td class="tdbordeCELDA7">&nbsp;</td>
											<%end if%>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=iif(Suma>=0,color_azul,color_rojo)%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
											<%if mostrarSaldo="on" then%>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=iif(saldoFinal>=0,color_azul,color_rojo)%>'><%=cstr((formatnumber(CambioDivisa(SaldoFinal,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1)))%></div></b><big></td>
											<%end if%>
										<%
										CloseFila
									end if
								end if%>
							</tbody>
						</table><br/><%
						NavPaginas lote,lotes,campo,criterio,texto,2
					else%>
						<script language="javascript" type="text/javascript">
						    alert("<%=LitNoExisteDatos%>");
						    parent.window.frames["botones"].document.location = "listado_caja_param_bt.asp?mode=add";
						    document.location = "listado_caja_param.asp?mode=add&caju=<%=enc.EncodeForJavascript(cajau)%>";
						</script>
                    <%end if
					rst.close
				end if
			else 'No se ha seleccionado anotación
				if AgrAnotacion<>"0" then 'Se agrupa por anotacion
					if AgrTipoPago<>0 and tpago&""="" then 'Se agrupa por el tipo de pago
						strwhereTotal=strwhere
						'strwhere=strwhere & " divisa=divisas.codigo and medio=tipo_pago.codigo and"
						strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
						'ega 18/03/2008 with(nolock)
						'ega 27/03/2008 si el gasto es nulo, se le pone 0, si el medio es nulo lo pone a vacio
						seleccion="select isnull(tipo_pago.gasto,0) as gasto,isnull(medio,'') as medio,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,ndocumento,ndocumento_pro,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                        ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                        seleccion=seleccion & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                        seleccion=seleccion & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                        seleccion=seleccion & " from caja c with(nolock) "
						seleccion=seleccion & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                        seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser
                        seleccion=seleccion & " " & strwhere
                        seleccion=seleccion & " order by tipo_pago.descripcion,tanotacion desc,fecha"
						rst.cursorlocation=3
''response.write("el seleccion 5 es-" & seleccion & "-<br>")
						rst.open seleccion,dsnCliente
						if not rst.eof then
							%><font class="cab"><b><%=LitApuntes%> : </b></font><font class="cab"><%=rst.recordcount%></font><br/>
							<font class="cab"><b><%=LitSaldoReal%> : </b></font><font style='color: <%=colorSaldoReal%>' class=cab><%=EncodeForHtml(saldoReal)%>&nbsp; <%=EncodeForHtml(AbreviaturaMB)%></b></font><br/><hr/><br/>
							<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'/><%
							'Calculos de páginas--------------------------
							CalculoPaginacion()
							NavPaginas lote,lotes,campo,criterio,texto,1
							'-----------------------------------------%><br/>
							<table width='100%' style='border-collapse: collapse;'>
								<thead>
									<tr bgcolor="<%=color_fondo%>">
										<td class="tdbordeCELDA7"><b><%=LitTipoPago%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitAnotacion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitFecha%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitDescripcion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitTipoDocumento%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitNDocumento%></b></td>
										<!-- cag --> <%
										if  mostrarGasto="on" then%>
											<td class="tdbordeCELDA7"><b><%=LitGasto%></b></td> <%
										end if %>
										<!-- fin cag-->											
										<td class="tdbordeCELDA7"><b><%=LitTipoApunte%></b></td>
										<td class="tdbordeCELDA7" align="right"><b><%=LitImporte%></b></td>
									</tr>
								</thead>
								<tbody><%
									Suma=0
									SubTotal=0
									SumaGasto=0
									SubTotalGasto=0
									AnotAnterior=""
									MedioAnterior=""
									fila=1
									while not rst.eof and fila<=MAXPAGINA%>
										<tr><%
											dato="&nbsp;"
											if (rst("tanotacion")<>AnotAnterior) or (rst("medio")<>MedioAnterior) then
												If AnotAnterior<>"" then
													'Fila para el Subtotal
													SubTotal=0
													SubTotalGasto=0
													rstAux.cursorlocation=3
													'ega with(nolock)
													'ega 27/03/2008 si el gasto es nulo, se le pone 0
													strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales " 
                                                    ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                                    strselect=strselect & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                                                    strselect=strselect & " ,ISNULL(ldt.value,c.tdocumento) as descr "
                                                    strselect=strselect & " from caja c with(nolock) "
													strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                                    strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                                    strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                                    strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                                    strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser
                                                    strselect=strselect & " " & strwhere & " and tanotacion='" & AnotAnterior & "' and medio='" & MedioAnterior & "' "
                                                    strselect=strselect & " order by tanotacion desc,fecha"
''response.write("el seleccion 6 es-" & strselect & "-<br>")
													rstAux.open strselect,dsnCliente
													for k=1 to rstAux.recordcount
														Nmedio=rstAux("nmedio")
														if rstAux("tanotacion")="ENTRADA" then
                                                            if rstAux("factcambio") & "">"" then
															    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
															    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                            else
															    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
															    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                            end if
														else
                                                            if rstAux("factcambio") & "">"" then
															    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
															    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                            else
															    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
															    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                            end if
														end if
														rstAux.movenext
													next
													rstAux.close
													if AnotAnterior="ENTRADA" then
														color=color_azul
													else
														color=color_rojo
													end if
													if mostrarGasto="on" then%>
														<td class="tdbordeCELDA7" colspan="5">&nbsp;</td>
														<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(AnotAnterior)%>&nbsp;<%=EncodeForHtml(Nmedio)%>:&nbsp;</b></td>
														<td class="tdbordeCELDA7" align="right"><div><b><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
														<td class="tdbordeCELDA7">&nbsp;</td>
														<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
													<%else%>
														<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
														<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=AnotAnterior%>&nbsp;<%=EncodeForHtml(Nmedio)%>:&nbsp;</b></td>
														<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
													<%end if%>
													</tr>
													<tr> <%'Fila de separacion%>
													<% if mostrarGasto="on" then%>
														<td class="tdbordeCELDA7" colspan="9">&nbsp;</td>									
													<%else%>
														<td class="tdbordeCELDA7" colspan="8">&nbsp;</td>
													<%end if%>
													</tr><%
													SubTotal=0
													SubTotalGasto=0
												end if
												dato=rst("tanotacion")
											end if
											if rst("tanotacion")="ENTRADA" then
                                                if rst("factcambio") & "">"" then
												    Suma=Suma + rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
												color=color_azul
											else
                                                if rst("factcambio") & "">"" then
												    Suma=Suma - rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma - CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB) '-
                                                end if
												color=color_rojo
											end if
											if (rst("medio")<>MedioAnterior) then
												%><td class="tdbordeCELDA7"><%=EncodeForHtml(rst("Nmedio"))%></td><%
											else%>
												<td class="tdbordeCELDA7">&nbsp;</td>
											<%end if%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(dato)%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("fecha"))%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descripcion"))%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descr"))%></td><%
											if rst("ndocumento_pro")&"">"" then%>
												<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("ndocumento_pro"))%></td>
											<%else%>
												<td class="tdbordeCELDA7"><%=EncodeForHtml(trimCodEmpresa(rst("ndocumento")))%></td>
											<%end if%>
											<%'cag
		 									if mostrarGasto="on" then
											   impGasto=rst("Importe")*rst("gasto")/100
											   if rst("tanotacion")="ENTRADA" then%>
												<td class="tdbordeCELDA7" align="right"><%=formatnumber(impGasto,rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
											   <%else%>
												<td class="tdbordeCELDA7" align="right"><%=formatnumber(abs(impGasto),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
											   <%end if	
											end if
											'fin cag
											%>										
											<td class="tdbordeCELDA7"><%=rst("tapuntdesc")%></td>
											<%if rst("tanotacion")="ENTRADA" then%>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></div></td>
											<%else%>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'>-<%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></div></td>
											<%end if

											'SubTotal=SubTotal + CambioDivisa(rst("importe"),rst("divisa"),MB)%>
										</tr><%
										AnotAnterior=rst("tanotacion")
										MedioAnterior=rst("medio")
										rst.movenext
										fila=fila+1
									wend
									if not rst.eof then 'escritos el número máximo de filas pero quedan registros
										'Se ha llegado al final de las filas permitidas justo
										'cuando se cambiaba de tipo de anotacion.Escribir subtotal
										if rst("tanotacion")<>AnotAnterior then
											if ucase(rst("tanotacion"))="SALIDA" then
												Med=rst("medio")
											else
												Med=MedioAnterior
											end if
											SubTotal=0
											SubTotalGasto=0
											'ega 18/03/2008 with(nolock)
											'ega 27/03/2008 si el gasto es nulo, se le pone 0
											strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales "
                                            ''Ricardo 07-10-2013 se añade el campo factcambio de los documentos
                                            strselect=strselect & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                                            strselect=strselect & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                                            strselect=strselect & " from caja c with(nolock) "
											strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser
                                            strselect=strselect & " "  & strwhere & " and tanotacion='" & AnotAnterior & "' and medio='" & Med & "' "
                                            strselect=strselect & " order by tanotacion desc,fecha"
											rstAux.cursorlocation=3
''response.write("el seleccion 7 es-" & strselect & "-<br>")
											rstAux.open strselect,dsnCliente
											for k=1 to rstAux.recordcount
												Nmedio=rstAux("nmedio")
												if rstAux("tanotacion")="ENTRADA" then
                                                    if rstAux("factcambio") & "">"" then
													    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
													    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
													    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                    end if
												else
                                                    if rstAux("factcambio") & "">"" then
													    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
													    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
													    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                    end if
												end if
												rstAux.movenext
											next
											rstAux.close
											color=color_azul
											if mostrarGasto="on" then%>
												<tr>
													<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
													<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=AnotAnterior%>&nbsp;<%=EncodeForHtml(Nmedio)%>:&nbsp;</b></td>
													<td class="tdbordeCELDA7" align="right"><div><b><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
													<td class="tdbordeCELDA7">&nbsp;</td>													
													<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												</tr><%
											else%>
												<tr>
													<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
													<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(AnotAnterior)%>&nbsp;<%=EncodeForHtml(Nmedio)%>:&nbsp;</b></td>
													<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												</tr><%
											end if
										end if
									else 'Se llegó al final de los registros
										'Fila para el Subtotal
										SubTotal=0
										Suma=0
										SubTotalGasto=0
										SumaGasto=0
										rst.movefirst
										for k=1 to rst.recordcount
											if rst("tanotacion")="ENTRADA" then
                                                if rst("factcambio") & "">"" then
												    Suma=Suma + rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
											else
                                                if rst("factcambio") & "">"" then
												    Suma=Suma - rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma - CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
											end if
											rst.movenext
										next
										'ega 18/03/2008 with(nolock)
										'ega 27/03/2008 si el gasto es nulo, se le pone 0
										strsel="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                                        ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                        strsel=strsel & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                                        strsel=strsel & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                                        strsel=strsel & " from caja c with(nolock) "
										strsel=strsel & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                        strsel=strsel & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                        strsel=strsel & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                        strsel=strsel & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                        strsel=strsel & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode	AND ldt.[language] = "& idiomaUser
                                        strsel=strsel & " " & strwhere
                                        strsel=strsel & " and tanotacion='" & AnotAnterior & "' and medio='" & MedioAnterior & "' "
                                        strsel=strsel & " order by tanotacion desc,fecha"
										rstAux.cursorlocation=3
''response.write("el seleccion 8 es-" & strsel & "-<br>")
										rstAux.open strsel,dsnCliente
										for k=1 to rstAux.recordcount
											Nmedio=rstAux("nmedio")
											if rstAux("tanotacion")="ENTRADA" then
                                                if rstAux("factcambio") & "">"" then
												    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
												    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                end if
											else
                                                if rstAux("factcambio") & "">"" then
												    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
												    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                end if
											end if
											rstAux.movenext
										next
										rstAux.close
										if mostrarGasto="on" then%>
										<tr>
											<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(AnotAnterior)%>&nbsp;<%=EncodeForHtml(Nmedio)%>:&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><b><div><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<td class="tdbordeCELDA7">&nbsp;</td>
											<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
										</tr>									
										<%else%>
										<tr>
											<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(AnotAnterior)%>&nbsp;<%=EncodeForHtml(Nmedio)%>:&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
										</tr>
										<%end if%>
										<tr> <%'Fila de separacion%>
										<%if mostrarGasto="on" then %>
											<td class="tdbordeCELDA7" colspan=9>&nbsp;</td>										
										<%else%>
											<td class="tdbordeCELDA7" colspan=8>&nbsp;</td>
										<%end if%>
										</tr>
										<tr> <%'Fila para el total
											'Suma=Suma-SubTotal
											if Suma>=0 then
												color=color_azul
											else
												color=color_rojo
											end if
											if mostrarGasto="on" then%>
												<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
												<td class="tdbordeCELDA7" align="right"><big><b><div><%=formatnumber(abs(SumaGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>											
												<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>											
											<%else%>
											<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
											<%end if%>
										</tr><%
										'EQUIVALENCIA EN PTAS
										if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
											DrawFila color_blau
											if Suma>=0 then
												color=color_azul
											else
												color=color_rojo
											end if
											if mostrarGasto="on" then%>
												<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><big><b><div><%=cstr((formatnumber(CambioDivisa(abs(SumaGasto),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>										
												<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
											<%else%>
												<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
											<%end if
											CloseFila
										end if
									end if%>
								</tbody>
							</table><br/><%
							NavPaginas lote,lotes,campo,criterio,texto,2
						else%>
							<script language="javascript" type="text/javascript">
							    alert("<%=LitNoExisteDatos%>");
							    parent.window.frames["botones"].document.location = "listado_caja_param_bt.asp?mode=add";
							    document.location = "listado_caja_param.asp?mode=add&caju=<%=enc.EncodeForJavascript(cajau)%>";
							</script><%
						end if
						rst.close
					else 'Se agrupa por anotacion pero NO por tipo de pago
						strwhereTotal=strwhere
						'strwhere=strwhere & " divisa=divisas.codigo and medio=tipo_pago.codigo and tipo_apuntes.codigo=tapunte and"
						strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
						'ega 18/03/2008 with(nolock)
						'ega 27/03/2008 si el gasto es nulo, se le pone 0, si el medio es nulo lo pone a vacio
						seleccion="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,ndocumento,ndocumento_pro,isnull(tipo_pago.descripcion,'') as medio,Importe,divisas.abreviatura,divisa,ndecimales "
                        ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                        seleccion=seleccion & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                        seleccion=seleccion & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                        seleccion=seleccion & " from caja c with(nolock) "
						seleccion=seleccion & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                        seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser 
                        seleccion=seleccion & " " & strwhere
                        seleccion=seleccion & " order by tanotacion desc,fecha"
                    'response.Write(seleccion)
                    'response.end
                        rst.cursorlocation=3
''response.write("el seleccion 9 es-" & seleccion & "-<br>")
						rst.open seleccion,dsnCliente
						if not rst.eof then
							%><font class="cab"><b><%=LitApuntes%> : </b></font><font class="cab"><%=rst.recordcount%></font><br/>
							<font class="cab"><b><%=LitSaldoReal%> : </b></font><font style='color: <%=colorSaldoReal%>' class=cab><%=EncodeForHtml(saldoReal)%>&nbsp; <%=EncodeForHtml(AbreviaturaMB)%></b></font><br/><hr/><br/>
							<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'/><%
							'Calculos de páginas--------------------------
							CalculoPaginacion()
							NavPaginas lote,lotes,campo,criterio,texto,1
							'-----------------------------------------%><br/>
							<table width='100%' style='border-collapse: collapse;'>
								<thead>
									<tr bgcolor="<%=color_fondo%>">
										<td class="tdbordeCELDA7"><b><%=LitAnotacion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitFecha%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitDescripcion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitTipoDocumento%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitNDocumento%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitTipoPago%></b></td>
										<!-- cag --> <%
										if  mostrarGasto="on" then%>
												<td class="tdbordeCELDA7"><b><%=LitGasto%></b></td> <%
										end if %>
										<!-- fin cag-->
										<td class="tdbordeCELDA7"><b><%=LitTipoApunte%></b></td>
										<td class="tdbordeCELDA7" align="right"><b><%=LitImporte%></b></td>
									</tr>
								</thead>
								<tbody><%
									Suma=0
									SumaGasto=0
									SubTotal=0
									AnotAnterior=""
									fila=1
									while not rst.eof and fila<=MAXPAGINA%>
										<tr><%
											dato="&nbsp;"
											if rst("tanotacion")&""<>AnotAnterior&"" then
												If AnotAnterior&""<>"" then
													'Fila para el Subtotal
															SubTotal=0
															SubTotalGasto=0
''response.write("los subtotal 10.1 son-" & SubTotal & "-" & SubTotalGasto & "-" & AnotAnterior & "-<br>")
															'ega 18/03/2008 with(nolock)
															'ega 27/03/2008 si el gasto es nulo, se le pone 0, si el medio es nulo lo pone a vacio
															strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,isnull(tipo_pago.descripcion,'') as medio,Importe,divisas.abreviatura,divisa,ndecimales "
                                                            ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                                            strselect=strselect & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                                                            strselect=strselect & " , ISNULL(ldt.value,c.tdocumento) as descr "
                                                            strselect=strselect & " from caja c with(nolock) "
															strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "  
                                                            strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' " 
                                                            strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser 
                                                            strselect=strselect & " " & strwhere & " and tanotacion='" & AnotAnterior & "' order by tanotacion desc,fecha"
															rstAux.cursorlocation=3
''response.write("el seleccion 10 es-" & strselect & "-<br>")
															rstAux.open strselect,dsnCliente
															for k=1 to rstAux.recordcount
																if rstAux("tanotacion")="ENTRADA" then
                                                                    if rstAux("factcambio") & "">"" then
																	    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
																	    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                                    else
																	    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
																	    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                                    end if
																else
                                                                    if rstAux("factcambio") & "">"" then
																	    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
																	    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                                    else
																	    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
																	    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB) '-
                                                                    end if
																end if
																rstAux.movenext
															next
															rstAux.close
														if AnotAnterior&""="ENTRADA" then
															color=color_azul
														else
															color=color_rojo
														end if
														if mostrarGasto="on" then%>
															<td class="tdbordeCELDA7" colspan="5">&nbsp;</td>
															<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%> :&nbsp;</b></td>
															<td class="tdbordeCELDA7" align="right"><div><b><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
															<td class="tdbordeCELDA7">&nbsp;</td>
															<td class="tdbordeCELDA7" align="right"><div style="color: <%=color%>;"><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
														<%else%>
															<td class="tdbordeCELDA7" colspan="6">&nbsp;</td>
															<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%> :&nbsp;</b></td>
															<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>														
														<%end if%>
													</tr>
													<tr> <%'Fila de separacion%>
													<% if mostrarGasto="on" then%>
														<td class="tdbordeCELDA7" colspan="9">&nbsp;</td>									
													<%else%>
														<td class="tdbordeCELDA7" colspan="8">&nbsp;</td>
													<%end if%>
													</tr><%
													SubTotal=0
												end if
												dato=rst("tanotacion")
											end if
											if rst("tanotacion")="ENTRADA" then
                                                if rst("factcambio") & "">"" then
												    Suma=Suma + rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
												color=color_azul
												signo=""
											else
                                                if rst("factcambio") & "">"" then
												    Suma=Suma - rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma - CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB) '-
                                                end if
												color=color_rojo
												signo="-"
											end if%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(dato)%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("fecha"))%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descripcion"))%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descr"))%></td><%
											if rst("ndocumento_pro")&"">"" then%>
												<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("ndocumento_pro"))%></td>
											<%else%>
												<td class="tdbordeCELDA7"><%=EncodeForHtml(trimCodEmpresa(rst("ndocumento")))%></td>
											<%end if%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("medio"))%></td>
											<%'cag
	 										if  mostrarGasto="on" then
												impGasto=rst("Importe")*rst("gasto")/100%>
												<td class="tdbordeCELDA7"  align="right"><%=formatnumber(abs(null_z(impGasto)),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
											<%end if
											'fin cag
											%>										
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("tapuntdesc"))%></td>
											<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><%=signo%><%=formatnumber(null_z(rst("Importe")),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></div></td><%
											'SubTotal=SubTotal + CambioDivisa(rst("importe"),rst("divisa"),MB)%>
										</tr><%
										AnotAnterior=rst("tanotacion")
''response.write("los subtotal 10.2 son-" & SubTotal & "-" & SubTotalGasto & "-" & AnotAnterior & "-<br>")
										rst.movenext
										fila=fila+1
									wend
									if not rst.eof then 'escritos el número máximo de filas pero quedan registros
										'Se ha llegado al final de las filas permitidas justo
										'cuando se acaban los ENTRADAS.Escribir subtotal
										if ucase(rst("tanotacion"))="SALIDA" and AnotAnterior="ENTRADA" then
											SubTotal=0
											SubTotalGasto=0
											'ega 18/03/2008 with(nolock)
											'ega 27/03/2008 si el gasto es nulo, se le pone 0, si el medio es nulo lo pone a vacio
											strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,isnull(tipo_pago.descripcion,'') as medio,Importe,divisas.abreviatura,divisa,ndecimales"
                                            ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                            strselect=strselect & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                                            strselect=strselect & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                                            strselect=strselect & " from caja c with(nolock) "
											strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = '" & idiomaUser & "' "
                                            strselect=strselect & " " & strwhere & " and tanotacion='ENTRADA' "
                                            strselect=strselect & " order by tanotacion desc,fecha"
											
											rstAux.cursorlocation=3
''response.write("el seleccion 11 es-" & strselect & "-<br>")
											rstAux.open strselect,dsnCliente
											for k=1 to rstAux.recordcount
												if rstAux("tanotacion")="ENTRADA" then
                                                    if rstAux("factcambio") & "">"" then
													    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
													    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
													    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                    end if
												else
                                                    if rstAux("factcambio") & "">"" then
													    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
													    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
													    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB) '-
                                                    end if
												end if
												rstAux.movenext
											next
											rstAux.close
											color=color_azul
											if mostrarGasto="on" then%>
											<tr>
												<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%> :&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(abs(null_z(SubTotalGasto)),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												<td class="tdbordeCELDA7">&nbsp;aaa</td>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(null_z(SubTotal),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%else%>
											<tr>
												<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%> :&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(null_z(SubTotal),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>											
											<%end if%>
											</tr><%
										end if
									else 'Se llegó al final de los registros
										'Fila para el Subtotal
										SubTotal=0
										Suma=0
										SubTotalGasto=0
										SumaGasto=0
										'rstAux.open "select tanotacion,Fecha,caja.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as medio,Importe,divisas.abreviatura,divisa,ndecimales,tipo_apuntes.descripcion as tapuntdesc from caja,divisas,tipo_pago,tipo_apuntes " & strwhere & " and tanotacion='SALIDA' order by tanotacion desc,fecha",session("backendlistados"),adUseClient, adLockReadOnly
										rst.movefirst
										for k=1 to rst.recordcount
											if rst("tanotacion")="ENTRADA" then
                                                if rst("factcambio") & "">"" then
												    Suma=Suma + rst("importe")/rst("factcambio")
												    SumaGasto=SumaGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    Suma=Suma + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
											else
                                                if rst("factcambio") & "">"" then
												    SubTotal=SubTotal + rst("importe")/rst("factcambio")
												    SubTotalGasto=SubTotalGasto + (rst("importe")*rst("gasto")/100)/rst("factcambio")
                                                else
												    SubTotal=SubTotal + CambioDivisa(rst("importe"),rst("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rst("importe")*rst("gasto")/100,rst("divisa"),MB)
                                                end if
											end if
											rst.movenext
										next
										if mostrarGasto="on" then%>
										<tr>
											<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%> :&nbsp;</b></td>
											<%if Suma>0 then %>
												<td class="tdbordeCELDA7" align="right"><b><div><%=formatnumber(abs(null_z(SumaGasto)),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												<td class="tdbordeCELDA7">&nbsp;</td>												
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(null_z(Suma),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%else%>
												<td class="tdbordeCELDA7" align="right"><b><div><%=formatnumber(abs(null_z(SubTotalGasto)),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
												<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'>-<%=formatnumber(null_z(SubTotal),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%end if%>
										</tr>
										<%else%>
										<tr>
											<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%> :&nbsp;</b></td>
											<%'if SubTotal=0 then %>
											<%if Suma>0 then %>
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'><%=formatnumber(null_z(Suma),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%else%>
												<td class="tdbordeCELDA7" align="right"><b><div style='color: <%=color%>'>-<%=formatnumber(null_z(SubTotal),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%end if%>
										</tr>
										<%end if%>
										<tr> <%'Fila de separacion%>
										<% if mostrarGasto="on" then%>
											<td class="tdbordeCELDA7" colspan="9">&nbsp;</td>									
										<%else%>
											<td class="tdbordeCELDA7" colspan="8">&nbsp;</td>
										<%end if%>
										</tr>
										<tr> <%'Fila para el total
											Suma=Suma-SubTotal
											SumaGasto=SumaGasto+SubTotalGasto
											if Suma>=0 then
												color=color_azul
											else
												color=color_rojo
											end if
											if mostrarGasto="on" then%>
												<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
												<td class="tdbordeCELDA7" align="right"><big><b><div ><%=formatnumber(abs(null_z(SumaGasto)),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
												<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(null_z(Suma),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
											<%else%>
												<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(null_z(Suma),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
											<%end if%>
										</tr><%
										'EQUIVALENCIA EN PTAS
										if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
											DrawFila color_blau
											if Suma>=0 then
												color=color_azul
											else
												color=color_rojo
											end if
											if mostrarGasto="on" then%>
												<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><big><b><div><%=cstr((formatnumber(CambioDivisa(abs(null_z(SumaGasto)),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
												<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(null_z(Suma),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>												
											<%else%>
												<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(null_z(Suma),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
											<%end if
											CloseFila
										end if
									end if%>
								</tbody>
							</table><br/><%
							NavPaginas lote,lotes,campo,criterio,texto,2
						else%>
							<script language="javascript" type="text/javascript">
							    alert("<%=LitNoExisteDatos%>");
							    parent.window.frames["botones"].document.location = "listado_caja_param_bt.asp?mode=add";
							    document.location = "listado_caja_param.asp?mode=add&caju=<%=enc.EncodeForJavascript(cajau)%>";
							</script><%
						end if
						rst.close
					end if
				else 'No se agrupa por anotacion
					if AgrTipoPago<>0 and tpago&""="" then 'Se agrupa por el tipo de pago
						strwhereTotal=strwhere
						'strwhere=strwhere & " divisa=divisas.codigo and medio=tipo_pago.codigo and"
						strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
						'ega 18/03/2008 with(nolock)
						'ega 27/03/2008 si el gasto es nulo, se le pone 0
						seleccion="select isnull(tipo_pago.gasto,0) as gasto,isnull(medio,'') as medio,Fecha,tanotacion,c.descripcion as Descripcion,tdocumento,ndocumento,ndocumento_pro,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                        ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                        seleccion=seleccion & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                        seleccion=seleccion & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                        seleccion=seleccion & " from caja c with(nolock) "
						seleccion=seleccion & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                        seleccion=seleccion & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = '" & idiomaUser & "' "
                        seleccion=seleccion & " " & strwhere
                        seleccion=seleccion & "  order by tipo_pago.descripcion,fecha"
						'response.Write(seleccion)						
                      
						rst.cursorlocation=3
''response.write("el seleccion 12 es-" & seleccion & "-<br>")
						rst.open seleccion,dsnCliente
						if not rst.eof then
							%><font class="cab"><b><%=LitApuntes%> : </b></font><font class="cab"><%=rst.recordcount%></font><br/>
							<font class="cab"><b><%=LitSaldoReal%> : </b></font><font style='color: <%=colorSaldoReal%>' class=cab><%=EncodeForHtml(saldoReal)%>&nbsp; <%=EncodeForHtml(AbreviaturaMB)%></b></font><br/><hr/><br/>
							<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'/><%
							'Calculos de las páginas-----------------------------
							CalculoPaginacion()
							NavPaginas lote,lotes,campo,criterio,texto,1
							'----------------------------------------------------%>
							<br/>
							<table width='100%' style='border-collapse: collapse;'>
								<thead>
									<tr bgcolor="<%=color_fondo%>">
										<td class="tdbordeCELDA7"><b><%=LitTipoPago%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitFecha%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitAnotacion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitDescripcion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitTipoDocumento%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitNDocumento%></b></td>
										<!-- cag --> <%
										if  mostrarGasto="on" then%>
											<td class="tdbordeCELDA7"><b><%=LitGasto%></b></td> <%
										end if %>									
										<td class="tdbordeCELDA7"><b><%=LitTipoApunte%></b></td>
										<td class="tdbordeCELDA7" align="right"><b><%=LitImporte%></b></td>
									</tr>
								</thead>
								<tbody><%
								Suma=0
								fila=1
								MedioAnterior=""
								While Not rst.eof and fila<=MAXPAGINA
									%><tr><%
									dato="&nbsp;"
									if rst("medio")<>MedioAnterior then
										If MedioAnterior<>"" then
											SubTotal=0
											SubTotalGasto=0
											'ega 18/03/2008 with(nolock)
											'ega 27/03/2008 si el gasto es nulo, se le pone 0
											strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                                            ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                            strselect=strselect & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                                            strselect=strselect & " ,ISNULL(ldt.value,c.tdocumento) as descr "
                                            strselect=strselect & " from caja c with(nolock) "
											strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                            strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = "& idiomaUser
                                            strselect=strselect & " " & strwhere & " and medio='" & MedioAnterior & "' "
                                            strselect=strselect & " order by tipo_pago.descripcion,fecha"
											rstAux.cursorlocation=3
''response.write("el seleccion 13 es-" & strselect & "-<br>")
											rstAux.open strselect,dsnCliente
											for k=1 to rstAux.recordcount
												Nmedio=rstAux("nmedio")
												if rstAux("tanotacion")="ENTRADA" then
                                                    if rstAux("factcambio") & "">"" then
                                                        SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
                                                        SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
													    'cag
													    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
													    'fin cag
                                                    end if
												else
                                                    if rstAux("factcambio") & "">"" then
                                                        SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
                                                        SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
													    'cag
    													SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
	    												'fin cag
                                                    end if
												end if
												rstAux.movenext
											next
											rstAux.close
											color=color_azul
											if SubTotal<0 then color=color_rojo
											'cag
											if mostrarGasto="on" then%>
											 	<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><div><b><%=formatnumber(abs(SubTotalGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>												
											 	<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%else%>
											 <td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%end if%>	
											</tr>											
											<tr> <%'Fila de separacion%>
											<% if mostrarGasto="on" then%>
												<td class="tdbordeCELDA7" colspan=9>&nbsp;</td>									
											<%else%>
												<td class="tdbordeCELDA7" colspan=8>&nbsp;</td>
											<%end if%>
											</tr><%
											SubTotal=0
										end if
										dato=rst("Nmedio")
									end if%>
									<td class="tdbordeCELDA7"><%=EncodeForHtml(dato)%></td>
									<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("fecha"))%></td>
									<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("tanotacion"))%></td>
									<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descripcion"))%></td>
									<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descr"))%></td><%
									if rst("ndocumento_pro")&"">"" then%>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("ndocumento_pro"))%></td>
									<%else%>
										<td class="tdbordeCELDA7"><%=EncodeForHtml(trimCodEmpresa(rst("ndocumento")))%></td>
									<%end if
									'cag
	 								if  mostrarGasto="on" then
											impGasto=rst("Importe")*rst("gasto")/100
											if rst("tanotacion")="ENTRADA" then
											  color=color_azul
											else
											  color=color_rojo
											end if
											%>
											<td class="tdbordeCELDA7"  align="right"><%=formatnumber(abs(impGasto),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
									<%end if
									'fin cag
									%><td class="tdbordeCELDA7"><%=EncodeForHtml(rst("tapuntdesc"))%></td><%
									if rst("tanotacion")="ENTRADA" then
										%><td class="tdbordeCELDA7" align="right"><div style='color: <%=color_azul%>'><%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1) & "&nbsp;" & EncodeForHtml(rst("abreviatura"))%></div></td><%
									else
										%><td class="tdbordeCELDA7" align="right"><div style='color: <%=color_rojo%>'>-<%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1) & "&nbsp;" & EncodeForHtml(rst("abreviatura"))%></div></td><%
									end if
									MedioAnterior=rst("medio")
									rst.movenext
									fila=fila+1
								wend
								if rst.eof then 'Se llegó al final de los registros%>
									<tr><% 'Fila para el subtotal
										SubTotal=0
										SubTotalGasto=0
										'ega 18/03/2008 with(nolock)
										'ega 27/03/2008 si el gasto es nulo, se le pone 0
										strselect="select isnull(tipo_pago.gasto,0) as gasto,tanotacion,Fecha,c.descripcion as Descripcion,tdocumento,tipo_pago.descripcion as Nmedio,Importe,divisas.abreviatura,divisa,ndecimales"
                                        ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                                        strselect=strselect & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                                        strselect=strselect & " ,tipo_apuntes.descripcion as tapuntdesc, ISNULL(ldt.value,c.tdocumento) as descr "
                                        strselect=strselect & " from caja c with(nolock) "
										strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                                        strselect=strselect & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                                        strselect=strselect & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                                        strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                                        strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = "& idiomaUser
                                        strselect=strselect & " " & strwhere & " and medio='" & MedioAnterior & "' "
                                        strselect=strselect & " order by tipo_pago.descripcion,fecha"
										rstAux.cursorlocation=3
''response.write("el seleccion 14 es-" & strselect & "-<br>")
										rstAux.open strselect,dsnCliente
										for k=1 to rstAux.recordcount
											Nmedio=rstAux("nmedio")
											if rstAux("tanotacion")="ENTRADA" then
                                                if rstAux("factcambio") & "">"" then
                                                    SubTotal=SubTotal + rstAux("importe")/rstAux("factcambio")
                                                    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    SubTotal=SubTotal + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    'cag
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
												    'fin cag
                                                end if
											else
                                                if rstAux("factcambio") & "">"" then
												    SubTotal=SubTotal - rstAux("importe")/rstAux("factcambio")
												    SubTotalGasto=SubTotalGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    SubTotal=SubTotal - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    SubTotalGasto=SubTotalGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                end if
											end if
											rstAux.movenext
										next
										rstAux.close
										color=color_azul
										if SubTotal<0 then color=color_rojo
										    'cag
											if  mostrarGasto="on" then %>
										  		<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
											<%else%>
												<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
											<%end if
											'fin cag
											%>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><b><%=LitTotal%>&nbsp;<%=EncodeForHtml(Nmedio)%> :&nbsp;</b></td>
											<%' cag
											if  mostrarGasto="on" then 
											    if SubTotalGasto<0 then
												   SubTotalGasto=abs(SubTotalGasto)
												 end if%>
												<td class="tdbordeCELDA7" align="right"><div><b><%=formatnumber(SubTotalGasto,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											  	<td class="tdbordeCELDA7">&nbsp;</td>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%else
											'fin cag%>
											<td class="tdbordeCELDA7" align="right"><div style='color: <%=color%>'><b><%=formatnumber(SubTotal,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></div></td>
											<%end if%>
										</tr>
										<tr> <%'Fila de separacion%>
										<% if mostrarGasto="on" then%>
											<td class="tdbordeCELDA7" colspan="10">&nbsp;</td>									
										<%else%>
											<td class="tdbordeCELDA7" colspan="9">&nbsp;</td>
										<%end if%>
										</tr>
									<tr><% 'Fila para el total
										Suma=0
										SumaGasto=0
										rstAux.cursorlocation=3
''response.write("el seleccion 15 es-" & seleccion & "-<br>")
										rstAux.open seleccion,dsnCliente
										for k=1 to rstAux.recordcount
											if rstAux("tanotacion")="ENTRADA" then
                                                if rstAux("factcambio") & "">"" then
												    Suma=Suma + rstAux("importe")/rstAux("factcambio")
												    SumaGasto=SumaGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    Suma=Suma + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                end if
											else
                                                if rstAux("factcambio") & "">"" then
												    Suma=Suma - rstAux("importe")/rstAux("factcambio")
												    SumaGasto=SumaGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                else
												    Suma=Suma - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
												    SumaGasto=SumaGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB) 
                                                end if
											end if
											rstAux.movenext
										next
										rstAux.close
										if Suma>=0 then
											color=color_azul
										else
											color=color_rojo
										end if
										'cag
										if  mostrarGasto="on" then %>
											<td class="tdbordeCELDA7" colspan="5">&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div><%=formatnumber(abs(SumaGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
											<td class="tdbordeCELDA7">&nbsp;</td>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>											
										<%else%>
										<td class="tdbordeCELDA7" colspan="6">&nbsp;</td>
										<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
										<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
										<%end if%>
									</tr><%
									'EQUIVALENCIA EN PTAS
									if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
										DrawFila color_blau
										if Suma>=0 then
											color=color_azul
										else
											color=color_rojo
										end if
										if mostrarGasto="on" then %>
											<td class="tdbordeCELDA7" colspan=5>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div><%=cstr((formatnumber(CambioDivisa(abs(SumaGasto),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
											<td class="tdbordeCELDA7">&nbsp;</td>											
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td><%
										else%>
											<td class="tdbordeCELDA7" colspan=6>&nbsp;</td>
											<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td><%
										end if
										CloseFila
									end if
								end if%>
								</tbody>
							</table><br/><%
							NavPaginas lote,lotes,campo,criterio,texto,2
						else%>
							<script language="javascript" type="text/javascript">
							    alert("<%=LitNoExisteDatos%>");
							    parent.window.frames["botones"].document.location = "listado_caja_param_bt.asp?mode=add";
							    document.location = "listado_caja_param.asp?mode=add&caju=<%=enc.EncodeForJavascript(cajau)%>";
							</script><%
						end if
						rst.close
					else 'No se agrupa por anotacion y no se agrupa por tipo de pago
						strwhereTotal=strwhere
						'strwhere=strwhere & " divisa=divisas.codigo and medio=tipo_pago.codigo and tapunte=tipo_apuntes.codigo and"
						strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
						seleccion="select isnull(tipo_pago.gasto,0) as gasto,caja,nanotacion,tanotacion,Fecha,c.descripcion as Descripcion,isnull(tdocumento,'') as tdocumento,isnull(ndocumento,'') as ndocumento,isnull(ndocumento_pro,'') as ndocumento_pro,isnull(tipo_pago.descripcion,'') as medio,Importe,divisas.abreviatura,divisa,ndecimales "
                        ''Ricardo 07-10-2013 se añade el campo change_currency de los documentos
                        seleccion=seleccion & " ,isnull(c.change_currency,divisas.factcambio) as factcambio "
                        seleccion=seleccion & " ,isnull(tipo_apuntes.descripcion,'') as tapuntdesc,0 as saldo "
                        seleccion=seleccion & " from caja c with(nolock) "
                        seleccion=seleccion & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_pago with(nolock) on c.medio=tipo_pago.codigo and tipo_pago.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & " left outer join tipo_apuntes with(nolock) on c.tapunte=tipo_apuntes.codigo and tipo_apuntes.codigo like '"& session("ncliente") &"%' "
                        seleccion=seleccion & strwhere
                        seleccion=seleccion & " order by fecha,nanotacion"
					    
					    'ega 28/03/2008 modificado el procedimiento para que devuelva el saldo inicial
						llamadaProc="EXEC ListadoExtractoCaja @NomTabla='" & session("usuario") & "' ,@seleccion='" & reemplazar(seleccion,"'","''") & "' ,@mostrarSaldo='" & mostrarSaldo & "',@fechaInicio='" & dfecha & "',@saldoInicial="&SaldoInicial
''response.write("el llamadaProc es-" & llamadaProc & "-<br>")
''response.end
                        'llamar al procedimiento ListadoExtractoCaja 
				        set command =  Server.CreateObject("ADODB.Command")
    		    
                        command.ActiveConnection =conn
                        command.CommandTimeout = 0
                        command.CommandText="ListadoExtractoCaja"
                        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado

                        command.Parameters.Append command.CreateParameter("@NomTabla",adVarChar,adParamInput,50,session("usuario"))
                        command.Parameters.Append command.CreateParameter("@seleccion",adVarChar,adParamInput,3000,seleccion)
                        command.Parameters.Append command.CreateParameter("@mostrarSaldo",adVarChar,adParamInput,10,mostrarSaldo)
                        command.Parameters.Append command.CreateParameter("@fechaInicio",adDate,adParamInput,,dfecha)
                        command.Parameters.Append command.CreateParameter("@saldoInicial",adCurrency ,adParamOutput,,SaldoInicial)
                        
                        command.Execute
                        
                        SaldoInicial = Command.Parameters("@saldoInicial").Value
    			        set command=nothing

                        strselect="select c.* "
                        ''Ricardo 07-10-2013 se añade el campo factcambio de los documentos
                        strselect=strselect & " ,isnull(c.factcambio,divisas.factcambio) as factcambio "
                        strselect=strselect & " ,ISNULL(ldt.value,tdocumento) as descr "
                        strselect=strselect & " from [egesticet].[" & session("usuario") & "] as c "
                        strselect=strselect & " LEFT OUTER JOIN ilion_admin..tipo_documentos td with(NOLOCK) ON c.tdocumento = td.tippdoc "
                        strselect=strselect & " LEFT OUTER JOIN ilion_admin..lit_doctypes ldt ON td.codigo = ldt.doccode AND ldt.[language] = " & idiomaUser
                        strselect=strselect & " left outer join divisas with(nolock) on c.divisa = divisas.codigo and divisas.codigo like '"& session("ncliente") &"%' "
                        strselect=strselect & " order by fecha"
						rst.cursorlocation=3
						rst.open strselect,dsnCliente

						if not rst.eof then
							%><font class="cab"><b><%=LitApuntes%> : </b></font><font class="cab"><%=rst.recordcount%></font><br/>
							<font class="cab"><b><%=LitSaldoReal%> : </b></font><font style='color: <%=colorSaldoReal%>' class=cab><%=EncodeForHtml(saldoReal)%>&nbsp; <%=EncodeForHtml(AbreviaturaMB)%></b></font><br/><hr/><br/>
							<input type='hidden' name='NumRegs' value='<%=rst.recordcount%>'/><%
							'Calculos de las páginas-----------------------------
							CalculoPaginacion()
							NavPaginas lote,lotes,campo,criterio,texto,1
							'----------------------------------------------------%>
							<br/>
							<table width='100%' style='border-collapse: collapse;'>
								<thead>
									<tr bgcolor="<%=color_fondo%>">
										<td class="tdbordeCELDA7"><b><%=LitFecha%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitAnotacion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitDescripcion%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitTipoDocumento%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitNDocumento%></b></td>
										<td class="tdbordeCELDA7"><b><%=LitTipoPago%></b></td>
										<!-- cag --> <%
										if  mostrarGasto="on" then%>
												<td class="tdbordeCELDA7"><b><%=LitGasto%></b></td> <%
										end if %>
										<!-- fin cag-->									
										<td class="tdbordeCELDA7"><b><%=LitTipoApunte%></b></td>
										<td class="tdbordeCELDA7" align="right"><b><%=LitImporte%></b></td>
										<%
										if mostrarSaldo="on" then
											%>
											<td class="tdbordeCELDA7" align="right"><b><%=LitSaldo%>(<%=EncodeForHtml(AbreviaturaMB)%>)</b></td>
											<%
										end if
										%>
									</tr>
								</thead>
								<tbody><%
								Suma=0
								fila=1
								' Ponemos la linea de saldo inicial'
								if mostrarSaldo="on" and lote=1 then
									'SaldoInicialEntradas=d_lookup("isnull(round(sum(importe/factcambio),"&rst("ndecimales")&"),0)","caja left outer join divisas ON divisa = divisas.codigo","tanotacion='ENTRADA' and caja='"&caja&"' and fecha<'" & dfecha & "'",session("backendlistados"))
									'SaldoInicialSalidas=d_lookup("isnull(round(sum(importe/factcambio),"&rst("ndecimales")&"),0)","caja left outer join divisas ON divisa = divisas.codigo","tanotacion='SALIDA' and caja='"&caja&"' and fecha<'" & dfecha & "'",session("backendlistados"))
									'SaldoInicial=SaldoInicialEntradas-SaldoInicialSalidas

                                    'ega 28/03/2008 no se consulta a la BD porque el saldo inicial lo devuelve el procedimiento ListadoExtractoCaja
									'strTotales="select Entradas.total-Salidas.Total as SaldoInicial "
									'strTotales=strTotales & " from (select isnull(round(sum(importe/factcambio),"&rst("ndecimales")&"),0) as Total from caja with (NOLOCK) left outer join divisas with (NOLOCK) ON divisa = divisas.codigo where tanotacion='ENTRADA' and caja='"&caja&"' and fecha<'" & dfecha & "') as Entradas, "
     								'strTotales=strTotales & "      (select isnull(round(sum(importe/factcambio),"&rst("ndecimales")&"),0) as Total from caja with (NOLOCK) left outer join divisas with (NOLOCK) ON divisa = divisas.codigo where tanotacion='SALIDA' and caja='"&caja&"' and fecha<'" & dfecha & "') as Salidas"
                                    'rstAux.open strTotales,session("backendlistados")

									'if not rstAux.eof then
									'	SaldoInicial=rstAux("SaldoInicial")
									'end if
									'rstAux.close									
									if SaldoInicial>=0 then
										colorSaldoInicial=color_azul
									else
										colorSaldoInicial=color_rojo
									end if
									%>
									<tr>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7"><%=LitSaldoInicial%></td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7">&nbsp;</td>
										<td class="tdbordeCELDA7" style='color:<%=colorSaldoInicial%>' align="right"><%=formatnumber(SaldoInicial,rst("ndecimales"),-1,0,-1)%></td>
									</tr>
								<%
								end if
								' FIN de la linea de saldo inicial'
									while not rst.eof and fila<=MAXPAGINA%>
										<tr>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("fecha"))%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("tanotacion"))%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descripcion"))%></td>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("descr"))%></td>
											<%if rst("ndocumento_pro")&"">"" then%>
												<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("ndocumento_pro"))%></td>
											<%else%>
												<td class="tdbordeCELDA7"><%=EncodeForHtml(trimCodEmpresa(rst("ndocumento")))%></td>
											<%end if%>
											<td class="tdbordeCELDA7"><%=EncodeForHtml(rst("medio"))%></td>
											<%'cag
		 									if  mostrarGasto="on" then
												impGasto=rst("Importe")*rst("gasto")/100
												if ucase(rst("tanotacion"))="SALIDA" then%>
												   <td class="tdbordeCELDA7" align="right"><%=formatnumber(abs(impGasto),rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>
												<%else%>
												   <td class="tdbordeCELDA7" align="right"><%=formatnumber(impGasto,rst("ndecimales"),-1,0,-1)%>&nbsp;<%=EncodeForHtml(rst("abreviatura"))%></td>																								   
												<%end if   
											end if
											'fin cag%>											
											<td class="dato"><%=EncodeForHtml(rst("tapuntdesc"))%></td><%
											if ucase(rst("tanotacion"))="SALIDA" then%>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color_rojo%>'>-<%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1) & "&nbsp;" & EncodeForHtml(rst("abreviatura"))%></div></td><%
												'Suma=Suma - CambioDivisa(rst("importe"),rst("divisa"),MB)
											else%>
												<td class="tdbordeCELDA7" align="right"><div style='color: <%=color_azul%>'><%=formatnumber(rst("Importe"),rst("ndecimales"),-1,0,-1) & "&nbsp;" & EncodeForHtml(rst("abreviatura"))%></div></td><%
												'Suma=Suma + CambioDivisa(rst("importe"),rst("divisa"),MB)
											end if
											if mostrarSaldo="on" then
												if rst("saldo")>=0 then
													colorSaldo=color_azul
												else
													colorSaldo=color_rojo
												end if
												%>
												<td class="tdbordeCELDA7" style='color:<%=colorSaldo%>' align="right"><%=formatnumber(rst("saldo"),n_decimales,-1,0,-1)%></td>
												<%
											end if
											%>
										</tr><%
										saldoFinal=rst("saldo")
										rst.movenext
										fila=fila+1
									wend
									if rst.eof then 'Se llegó al final de los registros%>
										<tr><% 'Fila para el total
											Suma=0
											SumaGasto=0
											colVacias=6
											rstAux.cursorlocation=3
''response.write("el seleccion 16 es-" & seleccion & "-<br>")
											rstAux.open seleccion,session("backendlistados")
											for k=1 to rstAux.recordcount
''response.write("el gasto 16 es-" & rstAux("tanotacion") & "-" & Suma & "-" & SumaGasto & "-" & rstAux("importe") & "-" & rstAux("gasto") & "-" & rstAux("divisa") & "-" & MB & "-<br>")
												if rstAux("tanotacion")="ENTRADA" then
                                                    if rstAux("factcambio") & "">"" then
                                                        Suma=Suma + rstAux("importe")/rstAux("factcambio")
                                                        SumaGasto=SumaGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    Suma=Suma + CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
                                                        SumaGasto=SumaGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                    end if
												else
                                                    if rstAux("factcambio") & "">"" then
                                                        Suma=Suma - rstAux("importe")/rstAux("factcambio")
                                                        SumaGasto=SumaGasto + (rstAux("importe")*rstAux("gasto")/100)/rstAux("factcambio")
                                                    else
													    Suma=Suma - CambioDivisa(rstAux("importe"),rstAux("divisa"),MB)
                                                        SumaGasto=SumaGasto + CambioDivisa(rstAux("importe")*rstAux("gasto")/100,rstAux("divisa"),MB)
                                                    end if
												end if
												rstAux.movenext
											next
											rstAux.close
''response.write("el SumaGasto 16 es-" & SumaGasto & "-<br>")
											if Suma>=0 then
												color=color_azul
											else
												color=color_rojo
											end if%>
											<!-- cag -->
											<!-- <td class="tdbordeCELDA7" colspan=<%=colVacias%>>&nbsp;</td> -->
											<%if mostrarGasto="on" then%>										
												<td class="tdbordeCELDA7" colspan=<%=colVacias-1%>>&nbsp;</td>
											  <%else %>
												<td class="tdbordeCELDA7" colspan=<%=colVacias%>>&nbsp;</td>
											<%end if%>
											<!-- fin cag-->
											<td class="tdbordeCELDA7" bgcolor=<%=color_fondo%> align="right"><big><b><%=LitSaldo%> :&nbsp;</b></big></td>
											<!-- cag -->
											<%if mostrarGasto="on" then%>
											<td class="tdbordeCELDA7" align="right"><big><b><div><%=formatnumber(abs(SumaGasto),n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></div></b></big></td>									
											<td class="tdbordeCELDA7">&nbsp;</td>
											<%end if%>
											<!-- fin cag-->																					
											<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=formatnumber(Suma,n_decimales,-1,0,-1)%>&nbsp;<%=EncodeForHtml(AbreviaturaMB)%></b></big></div></td>
											<%if mostrarSaldo="on" then%>
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=iif(saldoFinal>=0,color_azul,color_rojo)%>'><%=formatnumber(saldoFinal,n_decimales,-1,0,-1)%></b></big></div></td>
											<%end if%>
										</tr><%
										'EQUIVALENCIA EN PTAS
										if d_lookup("imp_equiv","configuracion","nempresa='" & session("ncliente") & "'",session("backendlistados")) then
											DrawFila color_blau
												if Suma>=0 then
													color=color_azul
												else
													color=color_rojo
												end if
												%>											
												<!-- cag -->
												<!-- <td class="tdbordeCELDA7" colspan="<%=colVacias%>">&nbsp;</td> -->
												<%if mostrarGasto="on" then%>										
													<td class="tdbordeCELDA7" colspan="<%=colVacias-1%>">&nbsp;</td>
												  <%else %>
													<td class="tdbordeCELDA7" colspan="<%=colVacias%>">&nbsp;</td>
												<%end if%>
												<!-- fin cag-->												
												<td class="tdbordeCELDA7" bgcolor="<%=color_fondo%>" align="right"><big><b><%=LitSaldo + " " + d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%>:&nbsp;</b></td>
												<!-- cag -->
												<%if mostrarGasto="on" then%>
													<td class="tdbordeCELDA7" align="right"><big><b><div><%=cstr((formatnumber(CambioDivisa(abs(SumaGasto),MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
													<td class="tdbordeCELDA7">&nbsp;</td>
												<%end if%>
												<!--fin cag-->
												<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=color%>'><%=cstr((formatnumber(CambioDivisa(Suma,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1))) & "&nbsp;" & d_lookup("abreviatura","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados"))%></div></b><big></td>
												<%if mostrarSaldo="on" then%>
													<td class="tdbordeCELDA7" align="right"><big><b><div style='color: <%=iif(saldoFinal>=0,color_azul,color_rojo)%>'><%=cstr((formatnumber(CambioDivisa(SaldoFinal,MB,session("ncliente") & "01"),d_lookup("ndecimales","divisas","codigo='" & session("ncliente") & "01'",session("backendlistados")),-1,0,-1)))%></div></b><big></td>
												<%end if%>
											<%
											CloseFila
										end if
									end if%>
								</tbody>
							</table><br/><%
							NavPaginas lote,lotes,campo,criterio,texto,2
						else%>
							<script language="javascript" type="text/javascript">
							    alert("<%=LitNoExisteDatos%>");
							    parent.window.frames["botones"].document.location = "listado_caja_param_bt.asp?mode=add";
							    document.location = "listado_caja_param.asp?mode=add&caju=<%=enc.EncodeForJavascript(cajau)%>";
							</script><%
						end if
						rst.close
					end if
				end if
			end if
	end if%>
</form>
<%connRound.close
set connRound = Nothing
set rst=nothing
set rstAux=nothing
set rstTD = nothing
set conn = nothing
set conn = nothing
set conn = nothing
end if%>
</body>
</html>