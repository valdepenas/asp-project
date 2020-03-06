<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->    
<!--#include file="../ico.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<TITLE><%=LitTituloDiv%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
</HEAD>

<%
dim enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>

<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">

    function Editar(p_codigo, p_npagina, p_campo, p_criterio, p_texto)
    {
        document.location="divisas.asp?mode=edit&p_codigo=" + p_codigo
                                                    +"&npagina="+ p_npagina
                                                    +"&campo="  + p_campo
                                                    +"&texto="  + p_texto
                                                    +"&criterio=" + p_criterio;

        parent.botones.document.location="divisas_bt.asp?mode=edit";
    }


    function isNumberKey(evt) {
        var charCode = (evt.which) ? evt.which : evt.keyCode;
        if (charCode > 31 && (charCode != 44 && charCode != 46 && (charCode < 48 || charCode > 57)))
            return false;
        return true;
    }

    function calcularFactor() {
        var val1            = document.getElementById("calcvalor1").value;
        var valMonedaDesde  = document.getElementById("calcmoneda1").value;
        var valMonedaHasta  = document.getElementById("calcmoneda2").value;
        val1            = val1.replace(",", ".");
        valMonedaDesde  = valMonedaDesde.replace(",", ".");
        valMonedaHasta  = valMonedaHasta.replace(",", ".");
        if (!Number.isNaN(val1) && !Number.isNaN(valMonedaDesde) && !Number.isNaN(valMonedaHasta)) {
            var result = (valMonedaHasta / valMonedaDesde) * val1;
            document.getElementById("calcvalor2").value = result.toFixed(14);
        }
        else {
            document.getElementById("calcvalor2").value = "";
        }

    }

    function calcularFactorFinal() {
        var cantidadMoneda  = document.getElementById("cantidadMoneda").value;
        var monedaFactor    = document.getElementById("monedaFactor").value;
        cantidadMoneda  = cantidadMoneda.replace(",", ".");
        monedaFactor    = monedaFactor.replace(",", ".");
        if (!Number.isNaN(monedaFactor) && !Number.isNaN(cantidadMoneda)) {
            var result = cantidadMoneda/monedaFactor;
            document.getElementById("e_FactorCambio").value = result.toFixed(14);
        }
        else {
            document.getElementById("e_FactorCambio").value = "";
        }
    }

</script>

<body bgcolor="<%=color_blau%>">
<%
    'Metodo que deja que solo haya 1 moneda base
    sub NormalizaTablaDivisas(codigoCompleto)
        'Primero comprobamos si existe la divisa con codigo codigoCompleto y es moneda_base
        strselectNTD1 = "select moneda_base from divisas with(NOLOCK) WHERE codigo = ? ;"
        set connNTD1     = Server.CreateObject("ADODB.Connection")
        set rstNTD1      = Server.CreateObject("ADODB.Recordset")
	    set commandNTD1  = Server.CreateObject("ADODB.Command")
        connNTD1.open session("dsn_cliente")
        connNTD1.cursorlocation        = 3
        commandNTD1.ActiveConnection   = connNTD1
        commandNTD1.CommandTimeout     = 60
        commandNTD1.CommandText        = strselectNTD1
        commandNTD1.CommandType        = adCmdText
        commandNTD1.Parameters.Append commandNTD1.CreateParameter("@codigoCompleto",adVarChar,adParamInput,15,codigoCompleto)
        set rstNTD1 = commandNTD1.Execute

        if not rstNTD1.EOF then 'si existe divisa
	        if rstNTD1("moneda_base")<>0 then 'si existe moneda base
            
                strselectNTD2 = "select moneda_base from divisas with(ROWLOCK) WHERE codigo like ?+'%' and codigo <> ? ;"
                set connNTD2     = Server.CreateObject("ADODB.Connection")
                set rstNTD2      = Server.CreateObject("ADODB.Recordset")
	            set commandNTD2  = Server.CreateObject("ADODB.Command")
	            connNTD2.open session("dsn_cliente")
                'connNTD2.cursorlocation=3
	            commandNTD2.ActiveConnection    = connNTD2
	            commandNTD2.CommandTimeout      = 60
	            commandNTD2.CommandText         = strselectNTD2
	            commandNTD2.CommandType         = adCmdText 
                commandNTD2.Parameters.Append commandNTD2.CreateParameter("@sesionNCliente", adVarChar, adParamInput, 5, session("ncliente"))
                commandNTD2.Parameters.Append commandNTD2.CreateParameter("@codigoCompleto", adVarChar, adParamInput, 15, codigoCompleto)
                rstNTD2.Open commandNTD2, , adOpenKeyset, adLockOptimistic

                while not rstNTD2.EOF
		            rstNTD2("moneda_base") = 0
		            rstNTD2.Update
		            rstNTD2.movenext
	            wend

                rstNTD2.close
                connNTD2.close
                set rstNTD2      = nothing
                set commandNTD2  = nothing
                set connNTD2     = nothing
            else
                %>
			    <script type="text/javascript" language="JavaScript">
                    window.alert("Error normalizando divisas. No se ha detectado moneda base con codigo: <%=enc.EncodeForHtmlAttribute(null_s(codigoCompleto))%>");
                    document.location = "divisas.asp";
			    </script>
                <%
            end if
        else
            %>
			<script type="text/javascript" language="JavaScript">
                window.alert("Error normalizando divisas. No se ha detectado divisas con codigo: <%=enc.EncodeForHtmlAttribute(null_s(codigoCompleto))%>");
                document.location = "divisas.asp";
			</script>
            <%
        end if

        rstNTD1.close
        connNTD1.close
        set rstNTD1      = nothing
        set commandNTD1  = nothing
        set connNTD1     = nothing

    end sub

    sub ResetMonedaBase()
        strselectRMB = "select moneda_base from divisas with(ROWLOCK) WHERE codigo like ?+'%';"

        set connRMB     = Server.CreateObject("ADODB.Connection")
        set rstRMB      = Server.CreateObject("ADODB.Recordset")
	    set commandRMB  = Server.CreateObject("ADODB.Command")
	    connRMB.open session("dsn_cliente")
        'connRMB.cursorlocation=3
	    commandRMB.ActiveConnection    = connRMB
	    commandRMB.CommandTimeout      = 60
	    commandRMB.CommandText         = strselectRMB
	    commandRMB.CommandType         = adCmdText 
        commandRMB.Parameters.Append commandRMB.CreateParameter("@sesionNCliente", adVarChar, adParamInput, 5, session("ncliente"))
        rstRMB.Open commandRMB, , adOpenKeyset, adLockOptimistic

        while not rstRMB.EOF
            if rstRMB("moneda_base") <> 0 then
		        rstRMB("moneda_base") = 0
		        rstRMB.Update
            end if
		    rstRMB.movenext
	    wend

        rstRMB.close
        connRMB.close
        set rstRMB      = nothing
        set commandRMB  = nothing
        set connRMB     = nothing

	end sub

    sub ActualizarSaldos()
        set rstAux = server.CreateObject("ADODB.Recordset")
	    rstAux.open "select saldo from cajas with(rowlock)",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	    while not rstAux.eof
		    rstAux("saldo")=CambioDivisa(rstAux("saldo"),MBAnterior,MBNueva)
		    rstAux.update
		    rstAux.movenext
	    wend
	    rstAux.close
    end sub

'**********************************************************************************************************'
' CODIGO PRINCIPAL DE LA PAGINA  **************************************************************************'
'**********************************************************************************************************'

if accesoPagina(session.sessionid,session("usuario"))=1 then
	'set connRound = Server.CreateObject("ADODB.Connection")
	'connRound.open dsnilion
    %>
	<form name="divisas" method="post" action="divisas.asp">
	<%'Leer parámetros de la página'
	mode=request("mode")

		p_i_codigo=limpiaCadena(Request.Form("i_codigo"))
		p_i_descripcion=limpiaCadena(request.form ("i_descripcion"))
		p_i_FactorCambio=limpiaCadena(Request.form("i_FactorCambio"))
		p_i_Ndecimales=limpiaCadena(request.form("i_Ndecimales"))
		p_i_FechaCotizacion=limpiaCadena(Request.form("i_FechaCotizacion"))
		p_i_Abreviatura=limpiaCadena(Request.form("i_Abreviatura"))
		p_i_AbreviaturaFE=limpiaCadena(Request.form("i_AbreviaturaFE"))
		p_i_MonedaBase=limpiaCadena(Request.form("i_MonedaBase"))
		p_CambiarMoneda=limpiaCadena(request("CambiarMoneda"))
		p_e_codigo=limpiaCadena(request.form("e_codigo"))
		p_p_codigo=limpiaCadena(request("p_codigo"))

		checkCadena p_p_codigo
		p_e_descripcion=limpiaCadena(request.form("e_descripcion"))
		p_e_FactorCambio=limpiaCadena(Request.form("e_FactorCambio"))
		p_e_Ndecimales=limpiaCadena(Request.form("e_Ndecimales"))
		p_e_FechaCotizacion=limpiaCadena(Request.form("e_FechaCotizacion"))
		p_e_Abreviatura=limpiaCadena(Request.form("e_Abreviatura"))
		p_e_AbreviaturaFE=limpiaCadena(Request.form("e_AbreviaturaFE"))
		p_e_MonedaBase=limpiaCadena(Request.form("e_MonedaBase"))
		p_c_codigo=limpiaCadena(request("codigo"))

		checkCadena p_c_codigo
		p_criterio=limpiaCadena(request("criterio"))
		p_campo=limpiaCadena(request("campo"))
		p_texto=limpiaCadena(request("texto"))
		p_npagina=limpiaCadena(request("npagina"))
		p_pagina=limpiaCadena(request("pagina"))

		set rst = server.CreateObject("ADODB.Recordset")
		'set rstAux = server.CreateObject("ADODB.Recordset")

  		PintarCabecera "divisas.asp"

'-------Crear
		'insertamos si nos llegan los valores del formulario'
		if p_i_codigo>"" and p_i_descripcion>"" then

			ErrorFecha = 0
			p_codigo=p_i_codigo
			p_descripcion=p_i_descripcion
			p_FactorCambio=p_i_FactorCambio
			p_Ndecimales=p_i_Ndecimales
			p_FechaCotizacion=p_i_FechaCotizacion
			if not isdate(p_FechaCotizacion) then
				ErrorFecha = 1
			end if
			p_Abreviatura=p_i_Abreviatura
			p_AbreviaturaFE=p_i_AbreviaturaFE
			p_MonedaBase=nz_b2(p_i_MonedaBase)

			rst.cursorlocation=3
			rst.Open "select * from divisas with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente")
			if rst.EOF then
				registro="PRIMERO"
			end if
			rst.Close

			puesto_moneda_base=0

            'rst.cursorlocation=2
			'rst.Open "select * from divisas with(rowlock) where codigo='" + session("ncliente") + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
            strselectCrear = "select * from divisas with(rowlock) where codigo = ? ;"
            set connCrear     = Server.CreateObject("ADODB.Connection")
            set rstCrear      = Server.CreateObject("ADODB.Recordset")
	        set commandCrear  = Server.CreateObject("ADODB.Command")
	        connCrear.open session("dsn_cliente")
            connCrear.cursorlocation=2
	        commandCrear.ActiveConnection    = connCrear
	        commandCrear.CommandTimeout      = 60
	        commandCrear.CommandText         = strselectCrear
	        commandCrear.CommandType         = adCmdText 
            commandCrear.Parameters.Append commandCrear.CreateParameter("@codigoCompleto", adVarChar, adParamInput, 15, session("ncliente") & p_codigo)
            rstCrear.Open commandCrear, , adOpenKeyset, adLockOptimistic

			if rstCrear.EOF then 'comprobar que codigo no exista
				if ErrorFecha = 0 then 'LA FECHA ESTA BIEN
        
			        if p_MonedaBase <> 0 then
				        if p_CambiarMoneda = "SI" then
					        ResetMonedaBase
				        else
					        p_MonedaBase = 0
				        end if
			        end if
            
					rstCrear.AddNew
					rstCrear("codigo")=session("ncliente") & p_codigo
					rstCrear("descripcion")=p_descripcion
					rstCrear("Factcambio")=replace(p_FactorCambio,".",",")
					rstCrear("Ndecimales")=p_Ndecimales
					rstCrear("fultrev")=p_FechaCotizacion
					rstCrear("abreviatura")=p_Abreviatura
					rstCrear("abreviatura_FE")=p_AbreviaturaFE
					if registro="PRIMERO" then
				 		rstCrear("moneda_base")=1
				 	else
				 	    if p_MonedaBase <> 0 then
				 	        puesto_moneda_base=1
                        end if
		 				rstCrear("moneda_base")=p_MonedaBase
				 	end if
					if err.number = -2147352571  then 
                        %>
						<script type="text/javascript" language="JavaScript">
                            window.alert("<%=LitMsgFactorCambioNumerico%>");
                            document.location = "divisas.asp";
						</script>
                        <%
						rstCrear.delete
						rstCrear.close
                        connCrear.close
                        set rstCrear      = nothing
                        set commandCrear  = nothing
                        set connCrear     = nothing
					else
						rstCrear.Update
					end if
				else 'LA FECHA ESTA MAL
                    %>
					<script type="text/javascript" language="JavaScript">
                        window.alert("<%=LitMsgFechaCotizacionFecha%>");
				 	</script>
                    <%
				end if
			else 'Codigo ya existe
                %>
				<script type="text/javascript">
			        window.alert("<%=LitMsgCodigoExiste%>");
				</script>
			    <%
            end if

 			'rst.Close
            rstCrear.close
            connCrear.close
            set rstCrear      = nothing
            set commandCrear  = nothing
            set connCrear     = nothing

			''ricardo 13-8-2010
		    if puesto_moneda_base=1 then
			    'rst.open "update divisas with(updlock) set moneda_base=0 where codigo like '" & session("ncliente") & "%' and codigo<>'" & session("ncliente") & p_codigo & "'", session("dsn_cliente"), adOpenKeyset, adLockOptimistic
                codigoCompl = session("ncliente") & p_codigo
                NormalizaTablaDivisas codigoCompl
            end if
		end if

'-------Editar 
		'actualizamos valores
		if p_e_codigo>"" or p_e_descripcion>"" then

			ErrorFecha = 0
			p_codigo=p_e_codigo
			p_descripcion=p_e_descripcion
			p_FactorCambio=p_e_FactorCambio
			p_Ndecimales=p_e_Ndecimales
			p_FechaCotizacion=p_e_FechaCotizacion
			if not isdate(p_FechaCotizacion) then
				ErrorFecha = 1
			end if
			p_Abreviatura=p_e_Abreviatura
			p_AbreviaturaFE=p_e_AbreviaturaFE
			p_MonedaBase=nz_b2(p_e_MonedaBase)
			puesto_moneda_base=0
			if ErrorFecha = 0 then 'LA FECHA ESTA BIEN
				'rst.Open "select * from divisas with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
                strselectEditar = "select * from divisas with(rowlock) where codigo = ?;"

                set connEditar     = Server.CreateObject("ADODB.Connection")
                set rstEditar      = Server.CreateObject("ADODB.Recordset")
	            set commandEditar  = Server.CreateObject("ADODB.Command")
	            connEditar.open session("dsn_cliente")
                'connEditar.cursorlocation=3
	            commandEditar.ActiveConnection    = connEditar
	            commandEditar.CommandTimeout      = 60
	            commandEditar.CommandText         = strselectEditar
	            commandEditar.CommandType         = adCmdText 
                commandEditar.Parameters.Append commandEditar.CreateParameter("@codigoCompleto", adVarChar, adParamInput, 15, p_codigo)
                rstEditar.Open commandEditar, , adOpenKeyset, adLockOptimistic

				if not rstEditar.EOF then
		  	 		if rstEditar("moneda_base") <> 0 then 'SE HA SELECCIONADO LA MONEDA BASE
			 			if p_MonedaBase = 0 then 'SE INTENTA QUITARLA COMO MONEDA BASE %>
	  		 				<script type="text/javascript" language="JavaScript">
		 						window.alert("<%=LitMsgModifMonedaBase%>")
		 					</script><%
						else
							on error resume next
      		   				rstEditar("descripcion")=p_descripcion
		 					rstEditar("Factcambio")=replace(p_FactorCambio,".",",")
							rstEditar("Ndecimales")=p_Ndecimales
		 					rstEditar("fultrev")=p_FechaCotizacion
			 				rstEditar("abreviatura")=p_Abreviatura
			 				rstEditar("abreviatura_FE")=p_AbreviaturaFE
				 	        if p_MonedaBase <> 0 then
				 	            puesto_moneda_base=1
                            end if
			 				rstEditar("moneda_base")=p_MonedaBase
         					if err.number = -2147352571 then
								rstEditar.close
                                connEditar.close
                                set rstEditar      = nothing
                                set commandEditar  = nothing
                                set connEditar     = nothing
                                %>
								<script type="text/javascript" language="JavaScript">
                                    window.alert("<%=LitMsgFactorCambioNumerico%>");
                                    document.location = "divisas.asp";
								</script>
                                <%
							else
      	   						rstEditar.Update
							end if
						end if
			 		else 'NO SE HA SELECCIONADO LA MONEDA BASE
			 			if p_MonedaBase = 0 then 'NO SE HA MARCADO COMO MONEDA BASE
							on error resume next
         					rstEditar("descripcion")=p_descripcion
		 					rstEditar("Factcambio")=replace(p_FactorCambio,".",",")
							rstEditar("Ndecimales")=p_Ndecimales
			 				rstEditar("fultrev")=p_FechaCotizacion
			 				rstEditar("abreviatura")=p_Abreviatura
			 				rstEditar("abreviatura_FE")=p_AbreviaturaFE
				 	        if p_MonedaBase <> 0 then
				 	            puesto_moneda_base=1
                            end if
		 					rstEditar("moneda_base")=p_MonedaBase
         					if err.number = -2147352571 then
								rstEditar.close
                                connEditar.close
                                set rstEditar      = nothing
                                set commandEditar  = nothing
                                set connEditar     = nothing
                                %>
								<script type="text/javascript" language="JavaScript">
                                    window.alert("<%=LitMsgFactorCambioNumerico%>");
                                    document.location = "divisas.asp";
								</script>
                                <%
							else
         						rstEditar.Update
							end if
						else ' SE MARCA COMO MONEDA BASE
							if p_CambiarMoneda="SI" then
                                
                                'rst.close
                                MBanteriorSelect = "select codigo from divisas with(nolock) where moneda_base<>0 and codigo like ?+'%'"
                                MBanterior=DLookupP1(MBanteriorSelect,session("ncliente")&"",adVarChar,5,session("dsn_cliente"))
								'MBanterior=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
								MBNueva=p_codigo
								ResetMonedaBase

								'rst.Open "select * from divisas with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
								on error resume next
      	   						rstEditar("descripcion")=p_descripcion
			 					rstEditar("Factcambio")=replace(p_FactorCambio,".",",")
								rstEditar("Ndecimales")=p_Ndecimales
		 						rstEditar("fultrev")=p_FechaCotizacion
		 						rstEditar("abreviatura")=p_Abreviatura
		 						rstEditar("abreviatura_FE")=p_AbreviaturaFE
				 	            if p_MonedaBase <> 0 then
				 	                puesto_moneda_base=1
                                end if
		 						rstEditar("moneda_base")=p_MonedaBase
	      	   					if err.number = -2147352571 then
									rstEditar.close
                                    connEditar.close
                                    set rstEditar      = nothing
                                    set commandEditar  = nothing
                                    set connEditar     = nothing
                                    %>
									<script type="text/javascript" language="JavaScript">
                                        window.alert("<%=LitMsgFactorCambioNumerico%>");
                                        document.location = "divisas.asp";
									</script>
                                    <%
								else
		         					rstEditar.Update
									ActualizarSaldos
								end if
							end if
	  					end if
					end if
      			else 
                    %>
         			<script type="text/javascript">
            			window.alert("<%=LitMsgCodigoNoExiste%>");
         			</script>
                    <%
	      		end if
      			'rst.Close

                rstEditar.close
                connEditar.close
                set rstEditar      = nothing
                set commandEditar  = nothing
                set connEditar     = nothing

			    ''ricardo 13-8-2010
		        if puesto_moneda_base=1 then
			        'rst.open "update divisas with(updlock) set moneda_base=0 where codigo like '" & session("ncliente") & "%' and codigo<>'" & p_codigo & "'", session("dsn_cliente"), adOpenKeyset, adLockOptimistic
                    NormalizaTablaDivisas p_codigo
                end if
			else 'LA FECHA ESTA MAL%>
	  			<script type="text/javascript" language="JavaScript">
		 			window.alert("<%=LitMsgModifFecha%>")
			 	</script><%
			end if
		end if

'-------Eliminar
		'eliminamos valores
		if mode="delete" and p_c_codigo>"" then
			p_codigo=p_c_codigo
                
			'rst.Open "select * from divisas with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
            strselectEliminar = "select * from divisas with(rowlock) where codigo = ? ;"
            set connEliminar     = Server.CreateObject("ADODB.Connection")
            set rstEliminar      = Server.CreateObject("ADODB.Recordset")
	        set commandEliminar  = Server.CreateObject("ADODB.Command")
	        connEliminar.open session("dsn_cliente")
            'connEliminar.cursorlocation=3
	        commandEliminar.ActiveConnection    = connEliminar
	        commandEliminar.CommandTimeout      = 60
	        commandEliminar.CommandText         = strselectEliminar
	        commandEliminar.CommandType         = adCmdText 
            commandEliminar.Parameters.Append commandEliminar.CreateParameter("@codigoCompleto", adVarChar, adParamInput, 25, p_codigo)
            rstEliminar.Open commandEliminar, , adOpenKeyset, adLockOptimistic

			if rstEliminar("moneda_base") <> 0 then 
                %>
				<script type="text/javascript" language="JavaScript">
					window.alert("<%=LitMsgBorrarMonedaBase%>")
				</script>
                <%
			else
				rstEliminar.Delete
			end if
			'rst.Close
            rstEliminar.close
            connEliminar.close
            set rstEliminar      = nothing
            set commandEliminar  = nothing
            set connEliminar     = nothing

		end if


        'Calculadora de divisas
        strselectCalc1      = "select * from divisas with(nolock) where codigo like ?+'%'"
        set connCalc1       = Server.CreateObject("ADODB.Connection")
        set rstCalc1        = Server.CreateObject("ADODB.Recordset")
	    set commandCalc1    = Server.CreateObject("ADODB.Command")
        connCalc1.open session("dsn_cliente")
        connCalc1.cursorlocation        = 3
        commandCalc1.ActiveConnection   = connCalc1
        commandCalc1.CommandTimeout     = 60
        commandCalc1.CommandText        = strselectCalc1
        commandCalc1.CommandType        = adCmdText
        commandCalc1.Parameters.Append commandCalc1.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente"))
        set rstCalc1 = commandCalc1.Execute

        strselectMB1 = "select count(*) from divisas with(nolock) where codigo like ?+'%'"
        contaDivisas=DLookupP1(strselectMB1,session("ncliente"),adVarChar,5,session("dsn_cliente"))

        'Cuando hay mas de una divisa
        if contaDivisas > 1 then

        %><hr/>
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitConsultaDivisas%></h6>
        <%
            DrawDiv "col-lg-3 col-md-3 col-sm-3 col-xs-6", "", ""
            DrawLabel "", "", LitCantidad
            response.write("<input type='text' class='width80' name='calcvalor1' id='calcvalor1' onkeypress='return isNumberKey(event)' value='1' oninput='calcularFactor()' />")
            CloseDiv

            DrawDiv "col-lg-3 col-md-3 col-sm-3 col-xs-6", "", ""
            DrawLabel "", "", LitMoneda
            response.write("<select class='width80' name='calcmoneda1' id='calcmoneda1' onchange='calcularFactor()' >")
            while not rstCalc1.EOF
		        if rstCalc1("MONEDA_BASE")<>0 then
			        encontrado=true
			        response.write("<option selected='selected' value='" & rstCalc1("FACTCAMBIO") & "'>" & rstCalc1("DESCRIPCION") & "</option>")
		        else
			        on error resume next
			        response.write("<option value='" & rstCalc1("FACTCAMBIO") & "'>" & rstCalc1("DESCRIPCION") & "</option>")
			        if err.number<>0 then
				        response.write("el error es-" & err.description & "-" & "DESCRIPCION" & "-<br>")
				        response.end
			        end if
			        on error goto 0
		        end if
		        rstCalc1.Movenext
	        wend
            response.write("</select>")
            CloseDiv
                
            DrawDiv "col-lg-3 col-md-3 col-sm-3 col-xs-6", "", ""
            DrawLabel "", "", LitEquivalenA
            response.write("<input type='text' class='width80' name='calcvalor2' id='calcvalor2' onkeypress='return isNumberKey(event)' value='1' disabled='disabled' />")
            CloseDiv
                        
            DrawDiv "col-lg-3 col-md-3 col-sm-3 col-xs-6", "", ""
            DrawLabel "", "", LitMoneda
            rstCalc1.MoveFirst()
            response.write("<select class='width80' name='calcmoneda2' id='calcmoneda2' onchange='calcularFactor()' >")
            while not rstCalc1.EOF
		        if rstCalc1("MONEDA_BASE")<>0 then
			        encontrado=true
			        response.write("<option selected='selected' value='" & rstCalc1("FACTCAMBIO") & "'>" & rstCalc1("DESCRIPCION") & "</option>")
		        else
			        on error resume next
			        response.write("<option value='" & rstCalc1("FACTCAMBIO") & "'>" & rstCalc1("DESCRIPCION") & "</option>")
			        if err.number<>0 then
				        response.write("el error es-" & err.description & "-" & "DESCRIPCION" & "-<br>")
				        response.end
			        end if
			        on error goto 0
		        end if
		        rstCalc1.Movenext
	        wend
            response.write("</select>")
            CloseDiv
    
        end if

        %>
        <hr/>
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitConfigDivisas%></h6>
        <%

		if p_texto>"" then
			if p_campo="codigo" then
				c_where=" where " + p_campo + " like '" + session("ncliente")
			else
				c_where=" where " + p_campo + " like '"
			end if
		else
			c_where=""
		end if

		if c_where>"" then
			select case p_criterio
				case "contiene"
					c_where=c_where + "%' + ? + '%'"
				case "termina"
					c_where=c_where + "%' + ? "
				case "empieza"
					c_where=c_where + "' + ? + '%'"
				case "igual"
					if p_campo="codigo" then
						c_where=" where " + p_campo + " = '" + session("ncliente") + "' + ? "
					else
						c_where=" where " + p_campo + " = ? "
					end if
			end select
		end if

		Alarma "divisas.asp"

        %>
		<hr/>
		<%

        c_select="select * from divisas with(nolock)"
		if c_where>"" then
			c_where=c_where + " and codigo like '" + session("ncliente") + "%'"
		else
			c_where=" where codigo like '" + session("ncliente") + "%'"
		end if
		c_select=c_select + c_where + " order by codigo"

		if p_npagina="" then
			p_npagina=1
		end if

		select case p_pagina
			case "siguiente"
				p_npagina=p_npagina+1
			case "anterior"
				p_npagina=p_npagina-1
		end select

        %>
  		<input type="hidden" name="h_npagina" value="<%=enc.EncodeForHtmlAttribute(null_s(cstr(p_npagina)))%>"/>
		<%

        strselectListar = c_select
        set connListar     = Server.CreateObject("ADODB.Connection")
        set rstListar      = Server.CreateObject("ADODB.Recordset")
	    set commandListar  = Server.CreateObject("ADODB.Command")
        connListar.open session("dsn_cliente")
        connListar.cursorlocation        = 3
        commandListar.ActiveConnection   = connListar
        commandListar.CommandTimeout     = 60
        commandListar.CommandText        = strselectListar
        commandListar.CommandType        = adCmdText
        if p_texto>"" then
            commandListar.Parameters.Append commandListar.CreateParameter("@ptexto",adVarChar,adParamInput,25,p_texto)
		end if
        set rstListar = commandListar.Execute

        'rst.Open c_select,session("dsn_cliente"),adUseClient, adLockReadOnly

		if not rstListar.EOF then
			rstListar.PageSize=NumReg
			rstListar.AbsolutePage=p_npagina
		end if

		if mode<>"edit" and rstListar.RecordCount>NumReg then
			if clng(p_npagina) >1 then%>
				<a class="CABECERA" href="divisas.asp?pagina=anterior&npagina=<%=enc.EncodeForHtmlAttribute(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForHtmlAttribute(null_s(p_campo))%>&criterio=<%=enc.EncodeForHtmlAttribute(null_s(p_criterio))%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
				<IMG SRC="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></a>
			<%end if
			textopag=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rstListar.PageCount)%>
			<font class="CELDA"> <%=textopag%> </font> <%
			if clng(p_npagina) < rstListar.PageCount then %>
	  			<a class="CABECERA" href="divisas.asp?pagina=siguiente&npagina=<%=enc.EncodeForHtmlAttribute(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForHtmlAttribute(null_s(p_campo))%>&criterio=<%=enc.EncodeForHtmlAttribute(null_s(p_criterio))%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
				<IMG SRC="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></a>
			<%end if%>
			<font class="CELDA">&nbsp;&nbsp; Ir a Pag. <input class="CELDA" type="text" name="SaltoPagina1" size="2">&nbsp;&nbsp;<a class="CELDAREF" style="display: inline" href="javascript:IrAPagina(1,'<%=enc.EncodeForHtmlAttribute(null_s(p_campo))%>','<%=enc.EncodeForHtmlAttribute(null_s(p_criterio))%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=enc.EncodeForHtmlAttribute(null_s(rstListar.PageCount))%>,'npagina');">Ir</a></font>
		<%end if

        if ModuloContratado(session("ncliente"),"FA") = 1 then
        %>
            <a class="CELDA" style="text-align:center" href="javascript:void(window.open('../../ILIONX45/LoyaltyBase/Management/LoyaltyCurrencyEvaluationRuleType.aspx','Evaluation','width=1020,height=500'))">
		        <img alt="btn_save_img" style="float: right;padding-right: 20px;" src="../images/images/configuracion.png" title="Gestionar tipos de reglas de evaluación para la divisa">
            </a>
        <%
        end if
        %>

		<table class="width100 md-table-responsive bCollapse" BORDER="0" CELLSPACING="1" CELLPADDING="1">
			<%Drawfila color_fondo
				DrawCeldaDet "'ENCABEZADOL underOrange width10' align='left'    style='text-align: left;' ", "", "", 0, "<b>" & LitCodigo & "</b>"
				DrawCeldaDet "'ENCABEZADOL underOrange width15' align='left'    style='text-align: left;' ", "", "", 0, "<b>" & LitDescripcion & "</b>"
				DrawCeldaDet "'ENCABEZADOL underOrange width25' align='right'   style='text-align: left;' ", "", "", 0, "<b>" & LitFactorDeCambio & "</b>"
				DrawCeldaDet "'ENCABEZADOL underOrange width5'  align='right'   style='text-align: left;' ", "", "", 0, "<b>" & LitDecimales & "</b>"
				DrawCeldaDet "'ENCABEZADOL underOrange width10' align='right'   style='text-align: left;' ", "", "", 0, "<b>" & LitFechaCotizacion & "</b>"
				DrawCeldaDet "'ENCABEZADOL underOrange width5'  align='center'  style='text-align: left;' ", "", "", 0, "<b>" & LitAbreviatura & "</b>"
				DrawCeldaDet "'ENCABEZADOL underOrange width5'  align='center'  style='text-align: left;' ", "", "", 0, "<b>" & LitISO & "</b>"
				DrawCeldaDet "'ENCABEZADOC underOrange width10' align='center'  style='text-align: left;' ", "", "", 0, "<b>" & LitMonedaBase & "</b>"
                
                par=false
				i=1

				while not rstListar.EOF and i<=NumReg
					if par then
						Drawfila color_terra
						par=false
					else
						Drawfila color_blau
						par=true
					end if

                    'Formulario modo edit
					if mode="edit" and p_p_codigo=rstListar("codigo") then
						DrawCeldaDet "'CELDAL7 width10' align='left' style='text-align: left;'", "", "", 0, trimCodEmpresa(rstListar("codigo"))%>
						<input type="hidden" name="e_codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rstListar("codigo")))%>">
						    <%
                            'DrawInputCelda "CELDALEFT","","","15",0,"","e_descripcion",rstListar("descripcion")
                            %>
                        <td class="CELDAL7 width15" align="left" style="text-align: left;">
                            <%
                            DrawInput "'width90'","","e_descripcion",enc.EncodeForHtmlAttribute(null_s(rstListar("descripcion"))),"maxlength='25'"
                            %>
                        </td>
                        <td class="CELDAL7 width25" align="right" style="text-align: left;">
                            <%
						    'DrawInputCelda "CELDARIGHT","","","7",0,"","e_FactorCambio",rstListar("factcambio")
                            DrawInput "'CELDAL7 width50'","","cantidadMoneda",enc.EncodeForHtmlAttribute(null_s(rstListar("factcambio"))),"oninput='calcularFactorFinal()' id='cantidadMoneda' onkeypress='return isNumberKey(event)' maxlength='50'"
                        
                            response.write(LitEn & " &nbsp;")
                            rstCalc1.MoveFirst()
                            response.write("<select class='width40' name='monedaFactor' id='monedaFactor' onchange='calcularFactorFinal()' >")
                            while not rstCalc1.EOF
		                        if rstCalc1("MONEDA_BASE")<>0 then
			                        encontrado=true
			                        response.write("<option selected='selected' value='" & rstCalc1("FACTCAMBIO") & "'>" & rstCalc1("DESCRIPCION") & "</option>")
		                        else
			                        on error resume next
			                        response.write("<option value='" & rstCalc1("FACTCAMBIO") & "'>" & rstCalc1("DESCRIPCION") & "</option>")
			                        if err.number<>0 then
				                        response.write("el error es-" & err.description & "-" & "DESCRIPCION" & "-<br>")
				                        response.end
			                        end if
			                        on error goto 0
		                        end if
		                        rstCalc1.Movenext
	                        wend
                            response.write("</select>")
                            %>
						    <input type="hidden" name="e_FactorCambio" id="e_FactorCambio" value="<%=enc.EncodeForHtmlAttribute(null_s(rstListar("factcambio")))%>">

                        </td>
                        <td class="CELDAL7 width5" align="right" style="text-align: left;">
                            <%
						    'DrawInputCelda "CELDARIGHT","","","2",0,"","e_Ndecimales",rstListar("ndecimales")
                            DrawInput "'CELDAL7 width90'","","e_Ndecimales",enc.EncodeForHtmlAttribute(null_s(rstListar("ndecimales"))),"maxlength='3'"
                            %>
                        </td>
                        <td class="CELDAL7 width10" align="right" style="text-align: left;">
                            <input class="width65" type="text" size="8" value="<%=enc.EncodeForHtmlAttribute(null_s(rstListar("fultrev")))%>" name="e_FechaCotizacion"/>
                            <%
                            DrawCalendar "e_FechaCotizacion"
						    %>
                        </td>
					    <td class="CELDAL7 width5" align="center" style="text-align: left;">
					        <INPUT CLASS="CELDAL7 width90" type="text" name="e_Abreviatura" value="<%=enc.EncodeForHtmlAttribute(trim(null_s(rstListar("abreviatura"))))%>" size="5" maxlength="5">
					    </td>
						<td class="CELDAL7 width5" align="center" style="text-align: left;">
						    <INPUT CLASS="CELDAL7 width90" type="text" name="e_AbreviaturaFE" value="<%=enc.EncodeForHtmlAttribute(null_s(rstListar("abreviatura_fe")))%>" size="5" maxlength="3">
						</td>
                        <td class="CELDAC7 width10" align="center" style="text-align: left;">
						    <%
                            'DrawCheckCelda "CELDACENTER","","","0","","e_MonedaBase",rstListar("moneda_base")
                            DrawCheck "","","e_MonedaBase",rstListar("moneda_base")
                            %>
                        </td>
                        <%
						'CAMPO OCULTO PARA LA CONFIRMACION DE CAMBIO DE MONEDA BASE EN LA EDICION
						if rstListar("moneda_base") <> 0 then%>
							<INPUT type="hidden" id="text1" name="h_MonedaBase" value="SI">
						<%else%>
							<INPUT type="hidden" id="text1" name="h_MonedaBase" value="NO">
						<%end if



					else
			            h_ref="javascript:Editar('" & rstListar("codigo") & "'," & _
			                                        p_npagina & ",'" & _
									                p_campo & "','" & _
									                p_criterio & "','" & _
									                enc.EncodeForJavascript(p_texto) & "');"
						'DrawCeldaHref "CELDAREF","left",false,trimCodEmpresa(rstListar("codigo")),h_ref
                        %>
                        <td class="CELDAL7 width10" align="left" style="text-align: left;">                
                            <%
                            DrawHref "CELDAREF","",trimCodEmpresa(rstListar("codigo")),h_ref
                            %>
                        </td><%                       
						DrawCeldaDet "'CELDAL7 width15' align='left'    style='text-align: left;'   ",       "", "", 0, enc.EncodeForHtmlAttribute(null_s(rstListar("descripcion")))
						DrawCeldaDet "'CELDAL7 width25' align='right'   style='text-align: left;'   ",       "", "", 0, enc.EncodeForHtmlAttribute(null_s(rstListar("factcambio")))
						DrawCeldaDet "'CELDAL7 width5'  align='right'   style='text-align: left;'   ",       "", "", 0, enc.EncodeForHtmlAttribute(null_s(rstListar("ndecimales")))
						DrawCeldaDet "'CELDAL7 width10' align='right'   style='text-align: left;'   ",       "", "", 0, enc.EncodeForHtmlAttribute(null_s(rstListar("fultrev")))
						DrawCeldaDet "'CELDAL7 width5'  align='center'  style='text-align: left;'   ",       "", "", 0, enc.EncodeForHtmlAttribute(null_s(rstListar("abreviatura")))
						DrawCeldaDet "'CELDAL7 width5'  align='center'  style='text-align: left;'   ",       "", "", 0, enc.EncodeForHtmlAttribute(null_s(rstListar("abreviatura_fe")))
						if rstListar("moneda_base") <> 0 then
							DrawCeldaDet "'CELDAC7 width10' align='center'  style='text-align: left;'", "", "", 0, "<IMG SRC='../images/" & ImgSeriePorDefectoSi & "' " & ParamImgSeriePorDefectoSi & ">"
						else
							DrawCeldaDet "'CELDAC7 width10' align='center'  style='text-align: left;'", "", "", 0, "<IMG SRC='../images/" & ImgSeriePorDefectoNo & "' " & ParamImgSeriePorDefectoNo & ">"
					end if                            
					CloseFila
				end if
				i = i + 1
				rstListar.MoveNext
			wend%>
		</table>
		<%if mode<>"edit" and rstListar.RecordCount>NumReg then
			if clng(p_npagina) >1 then %>
				<a class=CABECERA href="divisas.asp?pagina=anterior&npagina=<%=enc.EncodeForHtmlAttribute(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForHtmlAttribute(null_s(p_campo))%>&criterio=<%=enc.EncodeForHtmlAttribute(null_s(p_criterio))%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
				<IMG SRC="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></a>
			<%end if
			textopag=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rstListar.PageCount)%>
			<font class=CELDA> <%=textopag%> </font> <%
			if clng(p_npagina) < rstListar.PageCount then %>
				<a class=CABECERA href="divisas.asp?pagina=siguiente&npagina=<%=enc.EncodeForHtmlAttribute(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForHtmlAttribute(null_s(p_campo))%>&criterio=<%=enc.EncodeForHtmlAttribute(null_s(p_criterio))%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
				<IMG SRC="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></a>
			<%end if
			%><font class=CELDA>&nbsp;&nbsp; Ir a Pag. <input class=CELDA type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;<a class="CELDAREF" style="display: inline" href="javascript:IrAPagina(2,'<%=enc.EncodeForHtmlAttribute(null_s(p_campo))%>','<%=enc.EncodeForHtmlAttribute(null_s(p_criterio))%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=enc.EncodeForHtmlAttribute(null_s(rstListar.PageCount))%>,'npagina');">Ir</a></font><%
			'rst.Close
            rstListar.close
            connListar.close
            set rstListar      = nothing
            set commandListar  = nothing
            set connListar     = nothing

		end if%>
		<br>
		<%if mode<>"edit" then %>

			<hr>
			<table class="width100 md-table-responsive" width="100%" BORDER="0" CELLSPACING="1" CELLPADDING="1">
           <%DrawCeldaDet "'ENCABEZADOL underOrange width50' align='left' style='text-align: left;'", "", "", 0,"<b>" & LitNBregistro & "</b>"%>
            </table>
			<table class="width100 md-table-responsive underOrange bCollapse" BORDER="0" CELLSPACING="1" CELLPADDING="1">
				<tr class="underOrange">
					<%
                    DrawCeldaDet "'ENCABEZADOL width10' align='left'    style='text-align: left;' ", "","", 0, "<b>" & LitCodigo & "</b>"
					DrawCeldaDet "'ENCABEZADOL width15' align='left'    style='text-align: left;' ", "","", 0, "<b>" & LitDescripcion & "</b>"
					DrawCeldaDet "'ENCABEZADOL width25' align='right'   style='text-align: left;' ", "","", 0, "<b>" & LitFactorDeCambio & "</b>"
					DrawCeldaDet "'ENCABEZADOL width5'  align='right'   style='text-align: left;' ", "","", 0, "<b>" & LitDecimales & "</b>"
					DrawCeldaDet "'ENCABEZADOL width10' align='right'   style='text-align: left;' ", "","", 0, "<b>" & LitFechaCotizacion & "</b>"
					DrawCeldaDet "'ENCABEZADOL width5'  align='center'  style='text-align: left;' ", "","", 0, "<b>" & LitAbreviatura & "</b>"
					DrawCeldaDet "'ENCABEZADOL width5'  align='center'  style='text-align: left;' ", "","", 0, "<b>" & LitISO & "</b>"
					DrawCeldaDet "'ENCABEZADOC width10' align='center'  style='text-align: left;' ", "","", 0, "<b>" & LitMonedaBase & "</b>"
				%></tr>
				<tr><%	
                ''para poder crear una nueva divisa%>
                    <td class="CELDAL7 underOrange width10" style="text-align: left;">
                        <input class="width50" type="text" size="3" name="i_codigo" maxlength="10" />
                    </td>
                    <td class="CELDAL7 underOrange width15" style="text-align: left;">
                        <input class="width90" type="text" size="7" name="i_descripcion" maxlength="25" />
                    </td>
                    <td class="CELDAL7 underOrange width25" style="text-align: left;">
                        <input class="width90" type="text" size="3" name="i_FactorCambio" maxlength="50" />
                    </td>
                    <td class="CELDAL7 underOrange width5" style="text-align: left;">
                        <input class="width90" type="text" size="2" name="i_Ndecimales" maxlength="3" />
                    </td>
                    <td class="CELDAL7 underOrange width10" style="text-align: left;">
                        <input class="width65" type="text" size="8" name="i_FechaCotizacion" />
                        <%DrawCalendar "i_FechaCotizacion"%>
                    </td>
                    <td class="CELDAL7 underOrange width5" style="text-align: left;">
					    <input class="width90" type="text" size="5" name="i_Abreviatura" maxlength="5" value="" />
					</td>
					<td class="CELDAL7 underOrange width5" style="text-align: left;">
					    <input class="width90" type="text" size="3" name="i_AbreviaturaFE" maxlength="3" value="" />
					</td>
                    <td class="CELDAC7 underOrange width10" style="text-align: left;">
					    <input class="width10" type="checkbox" name="i_MonedaBase" />
					</td>
				</tr>
			</table>

		<%end if%>
	</form>
    <%

    rstCalc1.close
    connCalc1.close
    set rstCalc1      = nothing
    set commandCalc1  = nothing
    set connCalc1     = nothing

    'connRound.close
	'set connRound = Nothing
end if%>
</BODY>
</HTML>
