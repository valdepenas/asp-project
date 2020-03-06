<%@ Language=VBScript %>
<% 
'**RGU 8/3/2007: Añadir campo para series de facturas de abono de ventas (Por ahora no se muestra)
'' Toni Climent 14-01-2009: Modificación para distingir los TPV´s de empresas de hostelería del resto
'' TCD 21-01-2009: Modificacion del proceso de volcado de OBJETOTPVOFF para que se realice solo al crear
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../CatFamSubResponsive.inc"--> 
<!--#include file="../js/calendar.inc" -->
<!--#include file="../varios2.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="tiendas.inc" -->
<!--#include file="tiendas_register_datetime_changes.inc" -->
<!--#include file="../js/animatedCollapse.js.inc"-->
<!--#include file="../js/tabs.js.inc"-->
<!--#include file="../styles/modal.css.inc"-->
<!--#include file="../styles/generalData.css.inc"-->
<!--#include file="../styles/Section.css.inc"-->
<!--#include file="../styles/ExtraLink.css.inc"-->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="tiendas_linkextra.inc" -->
<!--#include file="../js/dropdown.js.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->
<!--#include file="../common/clientesActionDrop.inc" -->

<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script type="text/javascript" src="https://maps.googleapis.com/maps/api/js?key=AIzaSyCCRB21T7f3jp6c2UxH2vtBlzmZDivWzd0&sensor=true"></script>
<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('DatosDG', 'fade=1')
    animatedcollapse.addDiv('DatosDV', 'fade=1')
    animatedcollapse.addDiv('DatosDC', 'fade=1')
    animatedcollapse.addDiv('DatosCtr', 'fade=1')
    animatedcollapse.addDiv('DatosAgro', 'fade=1')
    animatedcollapse.addDiv('DatosCPC', 'fade=1')
    animatedcollapse.addDiv('DatosDSC', 'fade=1')
    animatedcollapse.addDiv('DatosOBJ', 'fade=1')
    animatedcollapse.addDiv('DatosSAL', 'fade=1')
    animatedcollapse.addDiv('DatosPRM', 'fade=1')

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }
    animatedcollapse.init()
</script>
<%
    function GenerateGoogleMapsLink(addressInputName, cityInputName, postalCodeInputName, regionInputName, countryInputName)
    end function

    dim enc
    set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
</head>
<%
cod=limpiaCadena(request.QueryString("codigo"))

set connRound = Server.CreateObject("ADODB.Connection")
connRound.open dsnilion%>
<script language="javascript" type="text/javascript">
    function SaveFile(alm,caj,ti)
    {
        if (caj=="") alert("<%=LitAsinarTiendaCaja%>");
        else
        {
            if (ExisteFichero("C:\\","cetel.tpv"))
            {
                if (confirm("<%=LitMsgSobreEscFichero%>")) document.getElementById("IdFichConf").src="fichconf.asp?a=" + alm + "&c=" + caj + "&t=" + ti;
            }
            else document.getElementById("IdFichConf").src="fichconf.asp?a=" + alm + "&c=" + caj + "&t=" + ti;
        }
    }

    function AnadirObj()
    {
        if (isNaN(document.tiendas.obj.value.replace(',','.')) || document.tiendas.obj.value=="")
        {
            window.alert("<%=LitObjNoNum%>");
            return;
        }

        if ( parseInt(document.tiendas.mes.value) <1 || parseInt(document.tiendas.mes.value) >12  || isNaN(document.tiendas.mes.value ) || document.tiendas.mes.value==""  )
        {
            window.alert("<%=LitMesNoNum%>");
            return;
        }

        if (document.tiendas.any.value.length != 4 || isNaN(document.tiendas.any.value ) || document.tiendas.any.value=="" )
        {
            window.alert("<%=LitAnyNoNum%>");
            return;
        }
        marcoObjetivos.document.tiendas_obj.action="tiendas_obj.asp?mode=add&tienda="+document.tiendas.hcodigo.value+"&mes="+parseInt(document.tiendas.mes.value)+"&any="+parseInt(document.tiendas.any.value)+"&objetivo="+document.tiendas.obj.value;
        marcoObjetivos.document.tiendas_obj.submit();

        document.tiendas.mes.value="";
        document.tiendas.any.value="";
        document.tiendas.obj.value="";
    }

    function AbrirCal()
    {
        if ( parseInt(document.tiendas.emes.value) <1 || parseInt(document.tiendas.emes.value) >12 || isNaN(document.tiendas.emes.value)  || document.tiendas.emes.value==""  )
        {
            window.alert("<%=LitMesNoNum%>");
            document.tiendas.emes.select();
            document.tiendas.emes.focus();
            return;
        }
        if (document.tiendas.eany.value.length != 4 || isNaN(document.tiendas.eany.value) || document.tiendas.eany.value=="" )
        {
            window.alert("<%=LitAnyNoNum%>");
            document.tiendas.eany.select();
            document.tiendas.eany.focus();
            return;
        }
        AbrirVentana("tiendas_ent.asp?mode=add&tienda="+document.tiendas.hcodigo.value+"&mes="+parseInt(document.tiendas.emes.value)+"&any="+parseInt(document.tiendas.eany.value) ,'P',<%=AltoVentana%>,<%=AnchoVentana%>);
    }
</script>
<body class="BODY_ASP">
<script type="text/javascript" language="javascript">
    function Abrir(pagina, p2,p3,p4,p5)
    {
        switch(p5)
        {
            case "1":
                p1="../central.asp?pag1="+pagina+"&pag2=productos/articulosPetroleo_bt.asp&mode=add"
                AbrirVentana(p1,p2,p3,p4)
                break;
            case "2":
                p1 = "../central.asp?pag1=" + pagina + "&pag2=productos/tanquesPetroleo_bt.asp&mode=add&cod=<%=enc.EncodeForJavascript(cod)%>"
                AbrirVentana(p1,p2,p3,p4)
                break;
            case "3":
                p1 = "../central.asp?pag1=" + pagina + "&pag2=productos/boquerelPetroleo_bt.asp&mode=add&cod=<%=enc.EncodeForJavascript(cod)%>"
                AbrirVentana(p1,p2,p3,p4)
                break;
            default:
                p1 = "../central.asp?pag1=" + pagina + "&pag2=productos/boquerelPetroleo_bt.asp&mode=add&cod=<%=enc.EncodeForJavascript(cod)%>"
                AbrirVentana(p1,p2,p3,p4)
                break;
        }
    }
</script>
<%function Pwd(pass)
    Dim cad
    cad=""
    
    if len(pass) = 0 or pass&""="" then
        cad=""
    else
        For cuen = 1 To len(pass) Step 1
            cad=cad & "*"
        next
    end if
    
    Pwd=cad
end function

'Crea la tabla que contiene la barra de grupos de datos (Generales,Series Ventas,Series Compras)
sub BarraNavegacion(modo,tieneTPV)
        %>
        <script language="javascript" type="text/javascript">
            jQuery("#DatosDG").show();
            jQuery("#DatosDV").hide();
            jQuery("#DatosDC").hide();
            jQuery("#DatosCtr").hide();
            jQuery("#DatosAgro").hide();
            jQuery("#DatosCPC").hide();
            jQuery("#DatosPRM").hide();
            jQuery("#DatosDSC").hide();
            jQuery("#DatosOBJ").hide();
            jQuery("#DatosSAL").hide();
        </script>
        <%
end sub

'*************************************************************************************************************
'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro(p_mode)

    p_codigo = limpiaCadena(request.querystring("codigo"))
	if p_codigo & "">"" and p_mode="first_save" then
		p_codigo=session("ncliente") & p_codigo
	end if

    set connr = nothing
    set commandr = nothing

    set connr = Server.CreateObject("ADODB.Connection")
    set commandr =  Server.CreateObject("ADODB.Command")

    connr.open session("dsn_cliente")
    connr.CursorLocation = 3
    commandr.ActiveConnection = connr
    commandr.CommandTimeout = 60
    commandr.CommandText="select * from tiendas where codigo=?"
    commandr.CommandType = adCmdText
    commandr.Parameters.Append commandr.CreateParameter("@codigo",adVarChar,adParamInput,10,p_codigo)

    rst.Open commandr, , adOpenKeyset, adLockOptimistic

	'Comprobamos que la tienda no existe ya
	if rst.eof then
		if p_mode="first_save" then
			rst.addnew
   			guarda=true
   		end if
	end if

	if p_mode="first_save" then
		if p_codigo = rst("codigo") then
	   		guarda = false
	   		rst.cancelupdate
	   		rst.close
	   		%><script language="javascript" type="text/javascript">
	   		      window.alert("<%=LitMsgTiendaYaEx%>");
	   		      document.tiendas.action="tiendas.asp?mode=add";
	   		      parent.pantalla.document.tiendas.submit();
	   		      parent.botones.document.location="tiendas_bt.asp?mode=add";
	   		</script><%
		else
	   		guarda = true
		end if
	else
		guarda=true
	end if

	if guarda=true then

        set conn = nothing
        set command = nothing

        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")

        conn.open session("dsn_cliente")
        conn.cursorlocation=3
        command.ActiveConnection =conn
        command.CommandTimeout = 60
        command.CommandText="select * from domicilios where pertenece=? and tipo_domicilio='TIENDA'"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,55,p_codigo)

        rstAux.open command, , adOpenKeyset, adLockOptimistic

		if rstAux.eof then
	   		rstAux.addnew
		end if

		rstAux("pertenece")      = Nulear(p_codigo)
		rstAux("tipo_domicilio") = "TIENDA"            
        rstAux("domicilio")      = Nulear(enc.EncodeForJavascript(request.form("domicilio")))
		rstAux("cp")             = Nulear(enc.EncodeForJavascript(request.form("cp")))
		rstAux("poblacion")      = Nulear(enc.EncodeForJavascript(request.form("poblacion")))       
	    rstAux("provincia")      = Nulear(enc.EncodeForJavascript(request.form("provincia")))
	    rstAux("pais")           = Nulear(enc.EncodeForJavascript(request.form("pais")))
	    rstAux("telefono")       = Nulear(enc.EncodeForJavascript(request.form("telefono")))

		rstAux.Update

		cod_tienda = rstAux("codigo")

		rstAux.Close

        conn.close
        set conn = nothing
        set command = nothing

		rst("codigo")=p_codigo
		rst("descripcion")= Nulear(enc.EncodeForJavascript(request.form("descripcion")))
		rst("domicilio") = cod_tienda
		rst("fax") = Nulear(enc.EncodeForJavascript(request.form("fax")))
		rst("observaciones") = Nulear(enc.EncodeForJavascript(request.form("observaciones")))

		rst("almacen") = request.form("almacen")
		rst("tarifa") = Nulear(request.form("tarifa"))
		rst("email") = Nulear(enc.EncodeForJavascript(email))
		rst("serieAlbCli")       = Nulear(seralbcli)
		rst("serieFacCli")       = Nulear(serfaccli)
		rst("seriePedCli")       = Nulear(serpedcli)
		rst("seriePreCli")       = Nulear(serprecli)
		rst("serieAlbPro")       = Nulear(seralbpro)
		rst("serieFacPro")       = Nulear(serfacpro)
		rst("seriePedPro")       = Nulear(serpedpro)
		rst("ncliente")		 = Nulear(request.form("ncliente"))

		'FLM:20100215: filtro por el módulo de fidelización premium.
		if si_tiene_modulo_ModFidelizacionPremium <> 0 then
		    rst("codcentro")    = Nulear(request.form("ncentro"))
        end if
        
		'***RGU 2/5/06 ***
		'Toni Climent 14-01-2009 Reutlizacion primacrecimiento para salones
		rst("porcrecimiento1")=miround(null_z(request.Form("pc1")),decpor)
		rst("porcrecimiento2")=miround(null_z(request.Form("pc2")),decpor)
		rst("porcrecimiento3")=miround(null_z(request.Form("pc3")),decpor)
		rst("porcrecimiento4")=miround(null_z(request.Form("pc4")),decpor)
		rst("porcrecimiento5")=miround(null_z(request.Form("pc5")),decpor)
		
		if not es_hostelera then
		    rst("primacrecimiento1")=miround(null_z(request.Form("pri1")),decpor)
		    rst("primacrecimiento2")=miround(null_z(request.Form("pri2")),decpor)
		    rst("primacrecimiento3")=miround(null_z(request.Form("pri3")),decpor)
		    rst("primacrecimiento4")=miround(null_z(request.Form("pri4")),decpor)
		    rst("primacrecimiento5")=miround(null_z(request.Form("pri5")),decpor)
		else
		    rst("primacrecimiento1")=miround(null_z(request.Form("sal1")),decpor)
		    rst("primacrecimiento2")=miround(null_z(request.Form("sal2")),decpor)
		    rst("primacrecimiento3")=miround(null_z(request.Form("sal3")),decpor)
		    rst("primacrecimiento4")=miround(null_z(request.Form("sal4")),decpor)
		    rst("primacrecimiento5")=miround(null_z(request.Form("sal5")),decpor)
		end if

		rst("dtol")=miround(null_z(request.Form("dtol")),decpor)
		rst("dtodia")=miround(null_z(request.Form("dtodia")),decpor)

		fdesde=nulear(request.form("dtodiadesde"))
		if len(fdesde) >0 then
			fhdesde=nulear(request.form("dtohoradesde"))
			if len(fhdesde)>0 then
				fdesde=fdesde&" "&fhdesde
			end if
		end if
		rst("dtodiadesde")=nulear(fdesde)
		fhasta=nulear(request.form("dtodiahasta"))
		if len(fhasta) >0 then
			fhhasta=nulear(request.form("dtohorahasta"))
			if len(fhhasta)>0 then
				fhasta=fhasta&" "&fhhasta
			end if
		end if
		rst("dtodiahasta")=nulear(fhasta)
		hdtodial=request.form("dtodial")
		rst("dtodial")=iif(hdtodial="on" or hdtodial="1",-1,0)
		rst("pwdpvpmanual")=nulear(request.form("pwdpvpmanual"))
		rst("pwddtoregalo")=request.form("pwdregalo")
		rst("dtoregalo")=miround(null_z(request.Form("dtoregalo")),decpor)
		rst("pwddtoencargado")=nulear(request.form("pwdencargado"))
		rst("dtoencargado")=miround(null_z(request.Form("dtoencargado")),decpor)

        if see_coordenates=1 then
            rst("x_coordenate") = Nulear(enc.EncodeForJavascript(request.form("x_coordenate")))
            rst("y_coordenate") = Nulear(enc.EncodeForJavascript(request.form("y_coordenate")))
        end if

         errorControlador=0
        'mmg: OrCU
        if si_tiene_modulo_OrCU <> 0 then
            if (Nulear(request.form("codCont"))<>Nulear(request.form("hcodCont"))) then
            if (Nulear(request.form("codCont"))&""<>"") then
		    'guardamos la configuracion de Ilion_admin con la ip_controlador

            set commandO = nothing
            set connO = Server.CreateObject("ADODB.Connection")
            set commandO =  Server.CreateObject("ADODB.Command")

            connO.open DsnIlion
            connO.cursorlocation=3
            commandO.ActiveConnection =connO
            commandO.CommandTimeout = 60
            commandO.CommandText="select codigo from controladores with(nolock) where codigo=?"
            commandO.CommandType = adCmdText
            commandO.Parameters.Append commandO.CreateParameter("@codigo",adVarChar,adParamInput,20,Nulear(request.form("codCont")))

            set rstAux = commandO.Execute

		    if not rstAux.EOF then
		         errorControlador=1
		         rstAux.Close
                 connO.Close
                 set connO = nothing
                 set commandO = nothing
		    else
		        rstAux.Close
                connO.Close
                set connO = nothing
                set commandO = nothing

                set connO = Server.CreateObject("ADODB.Connection")
                set commandO =  Server.CreateObject("ADODB.Command")

                connO.open DsnIlion
                connO.cursorlocation=3
                commandO.ActiveConnection =connO
                commandO.CommandTimeout = 60
                commandO.CommandText="select codigo from controladores with(nolock) where codigo=? and nempresa=?"
                commandO.CommandType = adCmdText
                commandO.Parameters.Append commandO.CreateParameter("@codigo",adVarChar,adParamInput,20,Nulear(request.form("codCont")))
                commandO.Parameters.Append commandO.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))

                set rstAux = command0.Execute

		         if not rstAux.EOF then
		            errorControlador=1
		            rstAux.Close
                    connO.Close
                    set connO = nothing
                    set commandO = nothing
		        else
		            rstAux.Close
                    connO.Close
                    set connO = nothing
                    set commandO = nothing

		            rst("Cod_controlador")     = Nulear(request.form("codCont"))
		            rst("Ip_controlador")      = Nulear(request.form("ipCon"))
		            rst("Usuario_controlador") = Nulear(request.form("login"))
		            rst("Pwd_controlador")     = Nulear(request.form("pass"))
		            'insertamos la nueva linea

		             if (Nulear(request.form("codCont"))&""<>"") then

                        set command = nothing
                        set conn = server.CreateObject("ADODB.Connection")
                        set command = Server.CreateObject("ADODB.Command")

                        conn.open session("dsn_cliente")

                        command.ActiveConnection = conn
                        command.CommandTimeout = 0
                        command.CommandType = adCmdText
		              
		                if (Nulear(request.form("hcodCont"))&""<>"") then
                            command.CommandText = "update controladores with(updlock) set codigo=? from controladores where codigo=? and nempresa=?"
                            command.Parameters.Append command.CreateParameter("@codigo", adVarChar, adParamInput, 20, Nulear(request.form("codCont")))
                            command.Parameters.Append command.CreateParameter("@codigoh", adVarChar, adParamInput, 20, Nulear(request.form("hcodCont")))
                            command.Parameters.Append command.CreateParameter("@ncliente", adVarChar, adParamInput, 5, session("ncliente"))
		                else
                            command.CommandText = "insert into controladores (nempresa,codigo) values (?,?)"
                            command.Parameters.Append command.CreateParameter("@ncliente", adVarChar, adParamInput, 5, session("ncliente"))
                            command.Parameters.Append command.CreateParameter("@codigo", adVarChar, adParamInput, 20, Nulear(request.form("codCont")))
		                end if

		                command.Execute
                        conn.Close

                        set conn = nothing
                        set command = nothing
                        set conn = nothing
		            end if
		        end if
		    end if
		   else
	            rst("Cod_controlador")     = Nulear(request.form("codCont"))
	            rst("Ip_controlador")      = Nulear(request.form("ipCon"))
	            rst("Usuario_controlador") = Nulear(request.form("login"))
	            rst("Pwd_controlador")     = Nulear(request.form("pass"))
		   end if
		  else
		     rst("Cod_controlador")     = Nulear(request.form("codCont"))
		     rst("Ip_controlador")      = Nulear(request.form("ipCon"))
		     rst("Usuario_controlador") = Nulear(request.form("login"))
		     rst("Pwd_controlador")     = Nulear(request.form("pass"))

		     if (Nulear(request.form("codCont"))&""<>"" and Nulear(request.form("hcodCont"))&""="") then
                set commandT = nothing
                set connT = Server.CreateObject("ADODB.Connection")
                set commandT =  Server.CreateObject("ADODB.Command")

                connT.open DsnIlion
                connT.cursorlocation=3
                commandT.ActiveConnection =connT
                commandT.CommandTimeout = 60
                commandT.CommandType = adCmdText
                commandT.CommandText="insert into controladores (nempresa,codigo) values (?,?)"

                commandT.Parameters.Append commandT.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))
                commandT.Parameters.Append commandT.CreateParameter("@codigo",adVarChar,adParamInput,20,Nulear(request.form("codCont")))

                commandT.Execute
                connT.Close
                set connT = nothing
                set commandT = nothing
		     end if
		  end if
		  rst("comunidad") = Nulear(request.Form("comunidad"))
		    
		  'dgb: 07-01-2009  Agroclub
            if si_tiene_modulo_Agroclub <> 0 then
		    'se guardan los datos
		        rst("agroclubnum")=Nulear(request.form("ncooperativa"))
		        rst("agroclublote")=Nulear(request.form("nlote"))
		        rst("agroclubversion")=Nulear(request.form("nversion"))
		    end if  
		end if

		''ricardo 14-8-2008 se guardaran los campos iptpv,porttpv,objetotpvoff
		ObtBDActual=encontrar_datos_dsn(session("dsn_cliente"),"Initial Catalog=")
		rstAux.cursorlocation=3

		''Toni Climent 14-01-2009 Distinguimos si se trata de una empresa del sector hostelero o no, obteniendo el campo OBJETOTPVOFF o OBJETOTPVHORECAS
		''TCD 21-01-2009 Modificamos para que la accion solo se realice al crear

        set commandO = nothing
        set connO = Server.CreateObject("ADODB.Connection")
        set commandO =  Server.CreateObject("ADODB.Command")

        connO.open DsnIlion
        connO.cursorlocation=3
        commandO.ActiveConnection =connO
        commandO.CommandTimeout = 60
        commandO.CommandType = adCmdText
        commandO.Parameters.Append commandO.CreateParameter("@bbdd",adVarChar,adParamInput,50,ObtBDActual)

		if p_mode="first_save" then		
		    if not es_hostelera then

                commandO.CommandText="select IPTPV,PORTTPV,OBJETOTPVOFF from CONFIGURACIONMEZCLA with(NOLOCK) where BBDD=?"

                set rstAux = commandO.Execute

		        if not rstAux.EOF then
		            rst("IPTPV")=rstAux("IPTPV")
		            rst("PORTTPV")=rstAux("PORTTPV")
		            rst("OBJETOTPVOFF")=rstAux("OBJETOTPVOFF")
		        end if
		        rstAux.Close
		    else

                commandO.CommandText="select IPTPV,PORTTPV,OBJETOTPVHORECAS from CONFIGURACIONMEZCLA with(NOLOCK) where BBDD=?"

                set rstAux = commandO.Execute

		        if not rstAux.EOF then
	    	        rst("IPTPV")=rstAux("IPTPV")
    		        rst("PORTTPV")=rstAux("PORTTPV")
		            rst("OBJETOTPVOFF")=rstAux("OBJETOTPVHORECAS")
		        end if
		        rstAux.Close
		    end if
        end if

		rst.Update
		rst.close
        connO.Close
        set connO = nothing
        set commandO = nothing

        resultSave=SaveStoreChanged(p_mode, p_codigo) 'added 2015-08-07 [JJC]

        if errorControlador=1 then
            mensaje=LitMsgNoGuardaCont & Nulear(request.form("codCont"))
            %><script language="javascript" type="text/javascript">
                  alert("<%=enc.EncodeForJavascript(mensaje)%>");
		    </script><%
        end if
		p_ntienda=p_codigo

        if si_tiene_modulo_OrCU <> 0 then
            'mmg OrCU >> Actualizamos la tabla ORCU_DATOS_SINCRONIZAR con los nuevos precios que se deban insertar
            if request.form("codCont")&"" <> "" then
		        if rstOrCU.state<>0 then rstOrCU.close

                set commandT = nothing
                set connT = Server.CreateObject("ADODB.Connection")
                set commandT =  Server.CreateObject("ADODB.Command")

                connT.open session("dsn_cliente")
                connT.cursorlocation=3
                commandT.ActiveConnection =connT
                commandT.CommandTimeout = 60
                commandT.CommandType = adCmdText
                commandT.CommandText="insert into ORCU_DATOS_SINCRONIZAR(NEmpresa,Objeto,Id,Instalacion,TipoCambio,Fecha) select ?,3,referencia,?,2,getdate() from articulos with (nolock) where referencia like ? + '%' and tipoProducto=1"

                commandT.Parameters.Append commandT.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente"))
                commandT.Parameters.Append commandT.CreateParameter("@pntienda",adVarChar,adParamInput,10,p_ntienda)
                commandT.Parameters.Append commandT.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente"))

                commandT.Execute
                connT.Close
                set connT = nothing
                set commandT = nothing
            end if
        end if
	end if

    connr.Close
    set connr = nothing
    set commandr = nothing

end sub

'*************************************************************************************************************
sub GuardarConfPC(puerto,numPuerto)
	if puerto<>"" and numPuerto<>"" then

        set commandG= nothing
        set connG = Server.CreateObject("ADODB.Connection")
        set commandG =  Server.CreateObject("ADODB.Command")

        connG.open session("dsn_cliente")
        connG.cursorlocation=3
        commandG.ActiveConnection =connG
        commandG.CommandTimeout = 60
        commandG.CommandType = adCmdText
        commandG.CommandText="select t.descripcion,t.visor from tpv as t where t.tpv like ? + '%' and t.tpv=? and t.caja=?"

        commandG.Parameters.Append commandG.CreateParameter("@fempr",adVarChar,adParamInput,10,session("f_empr"))
        commandG.Parameters.Append commandG.CreateParameter("@ftpv",adVarChar,adParamInput,8,session("f_tpv"))
        commandG.Parameters.Append commandG.CreateParameter("@fcaja",adVarChar,adParamInput,10,session("f_caja"))
        
        set rstAux = commandG.Execute

		if not rstAux.eof then
			visor=rstAux("visor")
			descripcionTPV=rstAux("descripcion")

            set commandT= nothing
            set connT = Server.CreateObject("ADODB.Connection")
            set commandT =  Server.CreateObject("ADODB.Command")

            connT.open session("dsn_cliente")
            connT.cursorlocation=3
            commandT.ActiveConnection =connT
            commandT.CommandTimeout = 60
            commandT.CommandType = adCmdText
            commandT.CommandText="select isnull(max(right(codigo,len(codigo)-5)),0) as mayor from visores where codigo like ? + '%'"

            commandT.Parameters.Append commandT.CreateParameter("@codigo",adVarChar,adChar,5,session("ncliente"))

            set rst = commandT.Execute
			if rst.eof then
				mayor="001"
			else
				mayor=CStr(CInt(rst("mayor"))+1)
				do while len(mayor)<3
					mayor="0" & mayor
				loop
			end if

			mayor=session("f_empr") & mayor

			rst.close
            connT.Close
            set connT = nothing
            set commandT = nothing

            set connT = Server.CreateObject("ADODB.Connection")
            set commandT =  Server.CreateObject("ADODB.Command")

            connT.open session("dsn_cliente")
            connT.cursorlocation=3
            commandT.ActiveConnection =connT
            commandT.CommandTimeout = 60
            commandT.CommandType = adCmdText
            commandT.CommandText="select v.codigo,v.puerto,v.npuerto,v.nombre from visores as v where codigo=?"

            commandT.Parameters.Append commandT.CreateParameter("@codigo",adVarChar,adChar,8,iif(isnull(visor),mayor,visor))

            set rst = commandT.Execute
			if rst.eof then
				rst.addnew
				rst("codigo")=mayor
				rst("nombre")=descripcionTPV
				rstAux("visor")=mayor
			end if
			rst("puerto")=puerto
			rst("npuerto")=numPuerto
			rst.update
			rstAux.update

			rst.close
            connT.Close
            set connT = nothing
            set commandT = nothing
		end if
		rstAux.close
        connG.Close
        set connG = nothing
        set commandG = nothing
	else
		EliminarConfPC
	end if
end sub

'*************************************************************************************************************
sub EliminarConfPC()
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")

    conn.open session("dsn_cliente")
    conn.cursorlocation=3
    command.ActiveConnection =conn
    command.CommandTimeout = 60
    command.CommandText="select t.descripcion,t.visor from tpv as t with(nolock) where t.visor is not null and t.tpv like ? + '%' and t.tpv=? and t.caja=?"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@fempr",adVarChar,adParamInput,8,session("f_empr"))
    command.Parameters.Append command.CreateParameter("@ftpv",adVarChar,adParamInput,8,session("f_tpv"))
    command.Parameters.Append command.CreateParameter("@tcaja",adVarChar,adParamInput,10,session("f_caja"))

    set rstAux = command.Execute

    if not rstAux.eof then
		visor=rstAux("visor")

		rstAux.close
        conn.close
        set conn = nothing   
        set command = nothing

        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open session("dsn_cliente")
        conn.cursorlocation=3
        command.ActiveConnection =conn
        command.CommandTimeout = 60
        command.CommandText="select * from tpv where visor=?"
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@visor",adChar,adParamInput,8,visor)

        set rst = command.Execute

		if not rst.eof then
			if rst.RecordCount=1 then
				rst("visor")=null
				rst.Update

                conn.close
                set conn = nothing
                set command = nothing
                rst.close

                set conn = Server.CreateObject("ADODB.Connection")
                set command =  Server.CreateObject("ADODB.Command")
                conn.open session("dsn_cliente")
                conn.cursorlocation=3
                command.ActiveConnection =conn
                command.CommandTimeout = 60
                command.CommandText="delete from visores with(rowlock) where codigo=?"
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@visor",adChar,adParamInput,8,visor)

                command.Execute

                conn.close
                set conn = nothing
                set command = nothing
			end if
        else
            conn.Close
            set conn = nothing
            set command = nothing
		    rst.close

		end if
	else
        conn.close
        set conn = nothing   
        set command = nothing
	end if
end sub

'*************************************************************************************************************
sub EliminarRegistro(codigo)
	'CONDICIONES PARA PODER BORRAR UNA TIENDA
    
    set command = nothing
    set conn = server.CreateObject("ADODB.Connection")
    set command = Server.CreateObject("ADODB.Command")

    conn.open session("dsn_cliente")

    command.ActiveConnection = conn
    command.CommandTimeout = 0
    command.CommandText =  "delete from domicilios with(rowlock) where pertenece = ? and tipo_domicilio='TIENDA'"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@codigo", adVarChar, adParamInput, 55, codigo)
    command.Execute

    conn.Close

    set command = nothing
    set conn = nothing

    set conn = server.CreateObject("ADODB.Connection")
    set command = Server.CreateObject("ADODB.Command")

    conn.open session("dsn_cliente")

    command.ActiveConnection = conn
    command.CommandTimeout = 0
    command.CommandText =  "delete from envios_tienda with(rowlock) where ntienda=?"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@codigo", adVarChar, adParamInput, 10, codigo)
    command.Execute

    conn.Close

    set command = nothing
    set conn = nothing

    set conn = server.CreateObject("ADODB.Connection")
    set command = Server.CreateObject("ADODB.Command")

    conn.open session("dsn_cliente")

    command.ActiveConnection = conn
    command.CommandTimeout = 0
    command.CommandText =  "delete from tiendas with(rowlock) where codigo=?"
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@codigo", adVarChar, adParamInput, 10, codigo)
    command.Execute

    conn.Close

    set command = nothing
    set conn = nothing
end sub

function DeleteUsrStore(store)

    xmlDelete="<CompanyUsersShop><us><UsersShop><SHOP>00012</SHOP><USER></USER></UsersShop></us></CompanyUsersShop>"

    err=0
    set connDU=server.CreateObject("ADODB.Connection")
    set cmdDU=server.CreateObject("ADODB.Command")
    
    connDU.open session("dsn_cliente")
    connDU.cursorlocation=3
    cmdDU.ActiveConnection =connDU
    cmdDU.CommandType = adCmdStoredProc
    cmdDU.CommandText="ILI_MANAGE_USER_SHOP"
    cmdDU.Parameters.Append cmdDU.CreateParameter("@companyId",adVarchar,,5,session("ncliente"))
    cmdDU.Parameters.Append cmdDU.CreateParameter("@action",adTinyInt,,,3)
    cmdDU.Parameters.Append cmdDU.CreateParameter("@data",adVarchar,,-1,xmlDelete)
    cmdDU.Parameters.Append cmdDU.CreateParameter("@store",adVarchar,,10,store)
    cmdDU.Parameters.Append cmdDU.CreateParameter("@result",adTinyInt,adParamOutput,,0)
    cmdDU.Execute,,adExecuteNoRecords
    if cint(null_z(cmdDU("@result")))<0 then
        %><script language="javascript" type="text/javascript">
              window.alert("<%=LitErrDelUS%>");
              document.location = "tiendas.asp?codigo=<%=enc.EncodeForJavascript(p_ntienda)%>&mode=browse";
              parent.botones.document.location = "tiendas_bt.asp?mode=browse";
		</script>
		<%
        err=1
    end if
    connDU.close
    set cmdDU=nothing
    set connDU=nothing
    DeleteUsrStore=err
end function


'*************************************************************************************************************
'*********   CODIGO PRINCIPAL DE LA PAGINA *******************************************************************
'*************************************************************************************************************

 %>

<form name="tiendas" method="post">
<div style="display:none;" id="mapToSelect">
    <div class="backGround" onclick="jQuery('#mapToSelect').hide(500);"></div>
    <div class="backGroundClose" onclick="jQuery('#mapToSelect').hide(500);"></div>
    <div id="map-canvas" class="map-canva"></div>
</div>                        

    <%
PintarCabecera "tiendas.asp"

    ''Toni Climent 14-01-2008: Discriminador entre empresas hosteleras y no hosteleras
    Dim es_hostelera  'Indica si la empresa pertenece al sector hostelero o no
    es_hostelera = false
    
	'Recordsets
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstDom = Server.CreateObject("ADODB.Recordset")
	set rstOrCU = Server.CreateObject("ADODB.Recordset")
	
	'modulos contratados
	si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
	si_tiene_modulo_OrCU=ModuloContratado(session("ncliente"),ModOrCU)
    si_tiene_modulo_Teekit=ModuloContratado(session("ncliente"),ModTeekit)

	'FLM:20100215:módulo fidelización premium
	si_tiene_modulo_ModFidelizacionPremium =  ModuloContratado(session("ncliente"),ModFidelizacionPremium)
    see_coordenates=0
    if si_tiene_modulo_ModFidelizacionPremium or ModuloContratado(session("ncliente"),ModFidelizacion) or ModuloContratado(session("ncliente"),ModEComerce) or ModuloContratado(session("ncliente"),ModCMS) then
        see_coordenates=1
    end if

    dim strselectp1
    strselectp1 = "select re from clientes with (NOLOCK) where ncliente=?"
	si_tiene_modulo_Agroclub=DLookupP1(strselectp1, session("ncliente")&"", adChar, 10, DSNIlion)

    %>
    <input type="hidden" name="si_tiene_modulo_ebesa" value="<%=enc.EncodeForHtmlAttribute(si_tiene_modulo_ebesa)%>" />
    <input type="hidden" name="si_tiene_modulo_OrCU" value="<%=enc.EncodeForHtmlAttribute(si_tiene_modulo_OrCU)%>" />
    <input type="hidden" name="si_tiene_modulo_ModFidelizacionPremium" value="<%=enc.EncodeForHtmlAttribute(si_tiene_modulo_ModFidelizacionPremium)%>" />
    <input type="hidden" name="si_tiene_modulo_Agroclub" value="<%=nz_b(si_tiene_modulo_Agroclub)%>" />
    <%

	'Leer parámetros de la página
	mode = Request.QueryString("mode")
	viene = limpiaCadena(Request.QueryString("viene"))
	
	dim VerAlmPetroleo
	ObtenerParametros("tiendas")
	
	
	if request.querystring("ndoc")="" then
		p_ntienda = limpiaCadena(request.querystring("codigo"))
	else
		p_ntienda = limpiaCadena(request.querystring("ndoc"))
	end if
	if mode<>"first_save" then
		CheckCadena p_ntienda
	end if

	p_domicilio = limpiaCadena(request.querystring("domicilio"))

	codigoR=limpiaCadena(request.querystring("codigo"))
	if mode<>"first_save" then
		CheckCadena codigoR
	end if

	if request.querystring("em") & "">"" then

		email=limpiaCadena(replace(request.querystring("em"),";","%3B"))
	else
		email=limpiaCadena(replace(request.form("em"),";","%3B"))
	end if
    email=replace(email,"%3B",";")

	if request.querystring("sac") & "">"" then
		seralbcli=limpiaCadena(request.querystring("sac"))
	else
		seralbcli=limpiaCadena(request.form("sac"))
	end if
	if request.querystring("sfc") & "">"" then
		serfaccli=limpiaCadena(request.querystring("sfc"))
	else
		serfaccli=limpiaCadena(request.form("sfc"))
	end if
	if request.querystring("sfabc") & "">"" then
		serfacAbcli=limpiaCadena(request.querystring("sfabc"))
	else
		serfacAbcli=limpiaCadena(request.form("sfabc"))
	end if
	if request.querystring("spc") & "">"" then
		serpedcli=limpiaCadena(request.querystring("spc"))
	else
		serpedcli=limpiaCadena(request.form("spc"))
	end if
	if request.querystring("sprc") & "">"" then
		serprecli=limpiaCadena(request.querystring("sprc"))
	else
		serprecli=limpiaCadena(request.form("sprc"))
	end if
	if request.querystring("sap") & "">"" then
		seralbpro=limpiaCadena(request.querystring("sap"))
	else
		seralbpro=limpiaCadena(request.form("sap"))
	end if
	if request.querystring("sfp") & "">"" then
		serfacpro=limpiaCadena(request.querystring("sfp"))
	else
		serfacpro=request.form("sfp")
	end if
	if request.querystring("spp") & "">"" then
		serpedpro=limpiaCadena(request.querystring("spp"))
	else
		serpedpro=limpiaCadena(request.form("spp"))
	end if

	if session("f_tpv") & "">"" then
		tpv=limpiaCadena(session("f_tpv"))
		tieneTPV="SI"
	else
		tpv=""
		tieneTPV="NO"
	end if
	if len(session("f_caja"))>5 then
		if right(session("f_caja"),len(session("f_caja"))-5) & "">"" and tieneTPV="SI" then
			caja=limpiaCadena(session("f_caja"))
			tieneTPV="SI"
		else
			caja=""
			tieneTPV="NO"
		end if
	else
		caja=""
		tieneTPV="NO"
	end if
	if session("f_empr") & "">"" and tieneTPV="SI" then
		codEmpresa=limpiaCadena(session("f_empr"))
		if codEmpresa=session("ncliente") then
			tieneTPV="SI"
		else
			codEmpresa=""
			tieneTPV="NO"
		end if
	else
		codEmpresa=""
		tieneTPV="NO"
	end if

	if request.form("puerto") & "">"" then
		puerto=limpiaCadena(request.form("puerto"))
	end if
	if request.form("numPuerto") & "">"" then
		numPuerto=limpiaCadena(request.form("numPuerto"))
	end if

	ncliente=limpiaCadena(request.querystring("ncliente"))
	if ncliente="" then ncliente=request.form("ncliente")

   'Acción a realizar
   
    set connT = nothing
    set commandT = nothing

    set connT = Server.CreateObject("ADODB.Connection")
    set commandT =  Server.CreateObject("ADODB.Command")

    connT.open session("dsn_cliente")
    connT.cursorlocation=3
    commandT.ActiveConnection =connT
    commandT.CommandTimeout = 60
    commandT.CommandType = adCmdText
    commandT.CommandText="select HORECAS from CONFIGURACION with(nolock) where nempresa=?"
    commandT.Parameters.Append commandT.CreateParameter("@nempresa", adChar, adParamInput, 5, session("ncliente"))

    set rstAux = commandT.Execute

    if rstAux("HORECAS") <> 0 then
		es_hostelera = true
	end if

	rstAux.Close
    connT.Close
    set connT = nothing
    set commandT = nothing

	%><input type="hidden" name="mode" value="<%=enc.EncodeForHtmlAttribute(null_s(mode))%>"/>
    <input type="hidden" name="h_tieneTPV" value="<%=enc.EncodeForHtmlAttribute(null_s(tieneTPV))%>"/>
    <input type="hidden" name="h_es_hostelera" value="<%=nz_b(es_hostelera)%>"/>
    <%


	if mode="save" then
		GuardarRegistro("save")
		GuardarConfPC puerto,numPuerto
		mode="browse"
	elseif mode="first_save" then
		GuardarRegistro("first_save")
		GuardarConfPC puerto,numPuerto
		mode="browse"
	elseif mode="delete" then
		'comprobamos si la tienda esta asignada a una caja

        set connT = Server.CreateObject("ADODB.Connection")
        set commandT =  Server.CreateObject("ADODB.Command")

        connT.open session("dsn_cliente")
        connT.cursorlocation=3
        commandT.ActiveConnection =connT
        commandT.CommandTimeout = 60
        commandT.CommandType = adCmdText
        commandT.CommandText="select tienda from cajas with(nolock) where tienda=?"
        commandT.Parameters.Append commandT.CreateParameter("@tienda", adVarChar, adParamInput, 10, p_ntienda)

        set rst = commandT.Execute

		if not rst.eof then
			rst.close
            connT.Close
            set connT = nothing
            set commandT = nothing

			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitTieAsignCaja%>");
			      document.location = "tiendas.asp?codigo=<%=enc.EncodeForJavascript(p_ntienda)%>&mode=browse";
			      parent.botones.document.location="tiendas_bt.asp?mode=browse";
			</script>
			<%
		else
			rst.close
            connT.Close
            set connT = nothing
            set commandT = nothing

            errdel=0
            if request.QueryString("del_st")&""="1" then
                errdel=DeleteUsrStore(codigoR)
            end if
            if errdel=0 then
                resultSave=SaveStoreChanged("delete", codigoR) 'added 2015-08-07 [JJC]
		        EliminarConfPC
		        EliminarRegistro codigoR
			    p_ntienda = ""
			    p_domicilio = ""
			    mode="add"
                %>
			    <script language="javascript" type="text/javascript">
			        parent.botones.document.location = "tiendas_bt.asp?mode=add";
			        SearchPage("tiendas_lsearch.asp?mode=init", 0);
			    </script>
                <%
            end if
		end if
	end if

    if p_ntienda & "">"" then
        strselectp1 = "select descripcion from tiendas with (NOLOCK) where codigo=?"
        p_nombre = DLookupP1(strselectp1, p_ntienda&"", adVarChar, 10, session("dsn_cliente")&"")
    end if

	if mode="edit" or mode="browse" then

        set connT = Server.CreateObject("ADODB.Connection")
        set commandT = Server.CreateObject("ADODB.Command")

        connT.open session("dsn_cliente")
        connT.cursorlocation=3

        commandT.ActiveConnection =connT
        commandT.CommandTimeout = 60
        commandT.CommandText = "select * from tiendas with(nolock) where codigo=?"
        commandT.CommandType = adCmdText
        commandT.Parameters.Append commandT.CreateParameter("@codigo",adVarChar,adParamInput,10,p_ntienda)

        set rst = commandT.Execute

 		if not rst.eof then
  	  	    rcodigo        = rst("codigo")
     	    rdescripcion   = rst("descripcion")
		    robservaciones = rst("observaciones")
		    rfax           = rst("fax")
	 	    ralmacen       = rst("almacen")
		    rtarifa        = rst("tarifa")
		    email=rst("email")
		    rncliente=rst("ncliente")&""
		    seralbcli=rst("serieAlbCli")
		    serfaccli=rst("serieFacCli")
		    serfacAbcli=rst("serieFacAbCli")
		    serpedcli=rst("seriePedCli")
		    serprecli=rst("seriePreCli")
		    seralbpro=rst("serieAlbPro")
		    serfacpro=rst("serieFacPro")
		    serpedpro=rst("seriePedPro")
		    'FLM:20100215: filtro por el módulo de fidelización premium.
	        if si_tiene_modulo_ModFidelizacionPremium <> 0 then
		        rcodcentro=rst("codcentro")
		    end if

		    rpc1=rst("porcrecimiento1")
		    rpri1=rst("primacrecimiento1")
		    rpc2=rst("porcrecimiento2")
		    rpri2=rst("primacrecimiento2")
		    rpc3=rst("porcrecimiento3")
		    rpri3=rst("primacrecimiento3")
		    rpc4=rst("porcrecimiento4")
		    rpri4=rst("primacrecimiento4")
		    rpc5=rst("porcrecimiento5")
		    rpri5=rst("primacrecimiento5")
			
		    'mmg:OrCU
		    if si_tiene_modulo_OrCU <> 0 then
		        codCont=rst("Cod_controlador")
		        login=rst("Usuario_controlador")
		        ipCon=rst("Ip_controlador")
		        pass=rst("Pwd_controlador")
		        civmdh=rst("comunidad")
			
		        'dgb: 07-01-2009 Agroclub  
		        if si_tiene_modulo_Agroclub <> 0 then
		            ncooperativa=rst("agroclubnum")
		            nlote=rst("agroclublote")
		            nversion=rst("agroclubversion")
		        end if
		    end if
			
		    rdtol=rst("dtol")
		    rdtodia=rst("dtodia")
		    rdtodiadesde=rst("dtodiadesde")
		    rdtodiahasta=rst("dtodiahasta")
		    rdtodial=rst("dtodial")
		    rpwdpvpmanual=rst("pwdpvpmanual")
		    rpwddtoregalo=rst("pwddtoregalo")
		    rdtoregalo=rst("dtoregalo")
		    rpwddtoencargado=rst("pwddtoencargado")
		    rdtoencargado=rst("dtoencargado")
		    rhdesde=hour(rdtodiadesde)
		    rhhasta=hour(rdtodiahasta)
		    rmdesde=minute(rdtodiadesde)
		    rmhasta=minute(rdtodiahasta)
            x_coordenate=rst("x_coordenate")
            y_coordenate=rst("y_coordenate")

		    if len(rhdesde)=1 then rhdesde="0"&rhdesde
		    if len(rhhasta)=1 then rhhasta="0"&rhhasta
		    if len(rmdesde)=1 then rmdesde="0"&rmdesde
		    if len(rmhasta)=1 then rmhasta="0"&rmhasta
 		end if
 		rst.close
        connT.Close
        set connT = nothing
        set commandT = nothing
	    strselect="select * from domicilios with(nolock) where pertenece=? and tipo_domicilio='TIENDA'"
                      
        set connDom = Server.CreateObject("ADODB.Connection")
        set commandDom = Server.CreateObject("ADODB.Command")

        connDom.open session("dsn_cliente")
        connDom.cursorlocation=3

        commandDom.ActiveConnection =connDom
        commandDom.CommandTimeout = 60
        commandDom.CommandText = strselect
        commandDom.CommandType = adCmdText
        commandDom.Parameters.Append commandDom.CreateParameter("@pntienda",adVarChar,adParamInput,55,p_ntienda)

        set rstDom = commandDom.Execute

		if not rstdom.eof then
			rdomicilio    = rstDom("domicilio")
			rpoblacion    = rstDom("poblacion")
			rcp           = rstDom("cp")
	   		rprovincia    = rstDom("provincia")
	   		rpais         = rstDom("pais")
   			rtelefono     = rstDom("telefono")
	    end if

     	rstDom.close
        connDom.Close
        set connDom = nothing
        set commandDom = nothing

        set connDom = Server.CreateObject("ADODB.Connection")
        set commandDom = Server.CreateObject("ADODB.Command")

        connDom.open session("dsn_cliente")
        connDom.cursorlocation=3

        commandDom.ActiveConnection =connDom
        commandDom.CommandTimeout = 60
        commandDom.CommandText = "select t.tpv,v.puerto,v.npuerto from tpv as t with(nolock) left join visores as v with(nolock) on t.visor=v.codigo where t.tpv like ? + '%' and t.tpv=? and t.caja=?"
        commandDom.CommandType = adCmdText
        commandDom.Parameters.Append commandDom.CreateParameter("@codEmpresa",adVarChar,adParamInput,8,codEmpresa)
        commandDom.Parameters.Append commandDom.CreateParameter("@tpv",adVarChar,adParamInput,8,tpv)
        commandDom.Parameters.Append commandDom.CreateParameter("@caja",adVarChar,adParamInput,10,caja)

        set rst = commandDom.Execute

		if not rst.eof then
			puerto=rst("puerto")
			numPuerto=rst("npuerto")
		end if

		rst.close
        connDom.Close
        set connDom = nothing
        set commandDom = nothing

	    end if
                    %>
	<div class="headers-wrapper">
    <%
            DrawDiv "header-client","",""
            DrawLabel "headerLabel","",LitCodigo
            DrawSpan "","",enc.EncodeForHtmlAttribute(trimCodEmpresa(p_ntienda)),""
            CloseDiv 
           
            DrawDiv "header-rsocial","",""
            DrawLabel "headerLabel","",iif(si_tiene_modulo_OrCU<>0,LitInstalacion,LitDescripcion)
            DrawSpan "","",p_nombre,""
            CloseDiv          
        %></div>
	
	<%'mmg: OrCU
	if si_tiene_modulo_OrCU <> 0 and mode <> "add" and mode <> "search" then
        BarraOpciones rcodigo%>
		
	<%end if
    alarma "tiendas.asp"

	if mode = "add" or mode="edit" or mode="browse" then
		BarraNavegacion mode,tieneTPV
	end if
 'estilos generico
 clase="span-browser"
	'Modo de inserción
	if mode = "add" or mode="edit" then
        %>
        <table style="width: 100%;"></table>
        <div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['DatosDG','DatosDV','DatosDC','DatosCtr','DatosAgro','DatosCPC','DatosPRM','DatosDSC','DatosSAL']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['DatosDG','DatosDV','DatosDC','DatosCtr','DatosAgro','DatosCPC','DatosPRM','DatosDSC','DatosSAL']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
        </div>

		<input type="hidden" name="hcodigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rcodigo))%>"/>
		<%' Inicio Borde Span%>       
        <div class="Section" id="S_DatosDG">
            <a href="#" rel="toggle[DatosDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitTiDatGen%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>                
            </a>
        <div class="SectionPanel" id="DatosDG">	   
				<%			   		
			   		if mode="edit" then
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitCodigo,"",trimCodEmpresa(rcodigo)
						%><input type="hidden" name="codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rcodigo))%>"/><%
		   			else			   		
                        EligeCelda "input","add","CELDA' maxlength='5'","","",0,LitCodigo,"codigo",5,trimCodEmpresa(rcodigo)
			   		end if		   			
		   			'mmg:OrCU
		   			if si_tiene_modulo_OrCU <> 0 then		   			   
                         EligeCelda "input","add","left","","",0,LitInstalacion,"descripcion",30,replace(rdescripcion,chr(39),"&#39;")
		   			else			   		   
                        EligeCelda "input","add","left","","",0,LitDescripcion,"descripcion",30,replace(rdescripcion,chr(39),"&#39;")
			   		end if			   		
				
					if rncliente>"" then
                        strselectp1 = "select rsocial from clientes with (NOLOCK) where ncliente=?"
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitCliente,"",trimcodempresa(rncliente) & " "&DLookupP1(strselectp1, rncliente&"", adChar, 10, session("dsn_cliente")&"")
						%><input class="CELDA" type="hidden" name="ncliente" value="<%=enc.EncodeForHtmlAttribute(null_s(rncliente))%>"/><%
					else
						%><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><label><%=LitCliente%></label><input class="CELDA" type="hidden" name="ncliente" value="<%=enc.EncodeForHtmlAttribute(null_s(iif(ncliente>"", ncliente, "")))%>"/><iframe id='frProyecto' src='docclientesResponsive.asp?viene=tiendas&siguiente=domicilio&mode=<%=enc.EncodeForHtmlAttribute(null_s(mode))%>&ncliente=<%=enc.EncodeForHtmlAttribute(null_s(ncliente))%>' class="width60 iframe-menu" height='30' frameborder="no" scrolling="no" noresize="noresize"></iframe></div>
  
					<%end if
                  
		   			'FLM:20100215: filtro por el módulo de fidelización premium.
		            if si_tiene_modulo_ModFidelizacionPremium <> 0 then		
		   			    %><div class="col-lg-4 col-md-6 col-sm-6 col-xs-12"><label><%=LitCentro%></label><input class="CELDA" type="hidden" name="ncentro" value="<%=enc.EncodeForHtmlAttribute(null_s(rcodcentro))%>"/><iframe id='Iframe1' src='../mantenimiento/doccentrosResponsive.asp?viene=tiendas&mode=<%=enc.EncodeForHtmlAttribute(null_s(mode))%>&ncentro=<%=enc.EncodeForHtmlAttribute(null_s(rcodcentro))%>&ncliente=<%=enc.EncodeForHtmlAttribute(null_s(ncliente))%>' class="width60 iframe-menu" height='30' frameborder="no" scrolling="no" noresize="noresize"></iframe></div>						
					<%  
                    else
					    '
					end if
				
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitDireccion

                    valor_rdomicilio=""
                    if rdomicilio & "">"" then
                        valor_rdomicilio=replace(rdomicilio,"'","&apos;")
                    else
                        valor_rdomicilio=""
                    end if %><input class="CELDA" size="35" maxlength="100" id="domicilio" name="domicilio" value="<%=enc.EncodeForHtmlAttribute(null_s(valor_rdomicilio))%>" onblur="loadCoordenate();" />                                

                    <%
                    CloseDiv
                    if si_tiene_modulo_Teekit <> 0 then %>
                    <div class="col-lg-4 col-md-6 col-sm-6 col-xs-12" map" style="display:inline" id="shopAddress_map" onclick="showMaps();" title="<%=LitViewMap %>">
                        <img width="16" hspace="2" height="16" border="0" align="top" title="Ver Mapa" alt="Ver Mapa" src="../images/images/mapa.png" style="cursor: pointer;" />
                    </div>
                    <% end if
                    if (session("version")&"" <> "5") then
                        DrawDiv "1","",""
                        CloseDiv
                    end if

                    DrawDiv "1","",""
                    DrawLabel "","",LitPoblacion
                    valor_rpoblacion=""
                    if rpoblacion & "">"" then
                        valor_rpoblacion=replace(rpoblacion,"'","&apos;")
                    else
                        valor_rpoblacion=""
                    end if %><input size="25" maxlength="50" id="poblacion" name="poblacion" value="<%=enc.EncodeForHtmlAttribute(null_s(valor_rpoblacion))%>" onblur="loadCoordenate();" /> 
            
                    <%CloseDiv

                    DrawDiv "1","",""
                    DrawLabel "","",LitCP%><input class="CELDA" size="5" maxlength="10" id="cp" name="cp" value="<%=enc.EncodeForHtmlAttribute(null_s(rcp))%>" onblur="loadCoordenate();" />    
          
                   <%CloseDiv
			
                    DrawDiv "1","",""
                    DrawLabel "","",LitProvincia
                    valor_rprovincia=""
                    if rprovincia & "">"" then
                        valor_rprovincia=replace(rprovincia,"'","&apos;")
                    else
                        valor_rprovincia=""
                    end if%><input size="25" maxlength="50" id="provincia" name="provincia" value="<%=enc.EncodeForHtmlAttribute(null_s(valor_rprovincia))%>" />                
                <%  CloseDiv
		   			
                    EligeCelda "input","add","left","","",0,LitPais,"pais",30,rpais
                    EligeCelda "input","add","left","","",0,LitTel1,"telefono",20,rtelefono
                    EligeCelda "input","add","left","","",0,LitFax,"fax",20,rfax
                    EligeCelda "input","add","left","","",0,LitTiEmail,"em",35,email
		   			
		   			''dgb: 29/09/2008 se anyade un campo para modulo Orcu del impuesto IVMDH
		   			if si_tiene_modulo_OrCU <> 0 then
                        set connT = nothing
                        set commandT = nothing

                        set connT = Server.CreateObject("ADODB.Connection")
                        set commandT = Server.CreateObject("ADODB.Command")

                        connT.open session("dsn_cliente")
                        connT.cursorlocation=3

                        commandT.ActiveConnection =connT
                        commandT.CommandTimeout = 60
                        commandT.CommandText = "select codigo, nombre from comunidad_orcu with(nolock) where codigo like ? + '%' order by nombre"
                        commandT.CommandType = adCmdText
                        commandT.Parameters.Append commandT.CreateParameter("@codigo",adVarChar,adParamInput,5,session("ncliente"))

                        set rstAux = commandT.Execute

				 	    DrawDiv "col-lg-4 col-md-6 col-sm-6 col-xs-12 visibilityHidden","",""
                        DrawLabel "","",LitIVMDH
                        DrawSelectCeldaWDNOLIT "CELDA","200","",0,"comunidad",rstAux,civmdh,"codigo","nombre","",""
                        CloseDiv
                    
                        rstAux.Close
                        connT.Close
                        set connT = nothing
                        set commandT = nothing
		   			end if
      		 
                    EligeCelda "text",mode,"","","",0,LitObservaciones,"observaciones","",robservaciones
		   			
                    if see_coordenates=1 then                       
                         DrawDiv "1","",""
                         DrawLabel "","",LitCoordenates%>                      
                        <%                             
                        DrawSpan clase,"","X ",""%><input class="CELDA" size="20" maxlength="25" id="x_coordenate" name="x_coordenate" value="<%=enc.EncodeForHtmlAttribute(null_s(x_coordenate))%>" />                       
                        <% DrawSpan clase,""," Y ",""%><input class="CELDA" size="20" maxlength="25" id="y_coordenate" name="y_coordenate" value="<%=enc.EncodeForHtmlAttribute(null_s(y_coordenate))%>" />
                        <%                             
                        CloseDiv
                    elseif (session("version")&"" <> "5") then
                        DrawDiv "1","",""
                        CloseDiv
                    end if					

                    set connT = nothing
                    set commandT = nothing

                    set connT = Server.CreateObject("ADODB.Connection")
                    set commandT = Server.CreateObject("ADODB.Command")

                    connT.open session("dsn_cliente")
                    connT.cursorlocation=3

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandType = adCmdText        
				
					if si_tiene_modulo_OrCU <> 0 and VerAlmPetroleo<>1 then
                        commandT.CommandText = "select codigo, descripcion from almacenes with(nolock) where codigo like ? +'%' and Tienda is null order by descripcion"
				 	else
                        commandT.CommandText = "select codigo, descripcion from almacenes with(nolock) where codigo like ? +'%' order by descripcion"                       
				 	end if

                    commandT.Parameters.Append commandT.CreateParameter("@codigo",adVarChar,adParamInput,5,session("ncliente"))
                    set rstAux = commandT.Execute

				 	DrawSelectCelda "CELDA","200","",0,LitAlmacen,"almacen",rstAux,ralmacen,"codigo","descripcion","",""

					rstAux.close
                    connT.Close
                    set connT = nothing
                    set commandT = nothing

                    set connT = Server.CreateObject("ADODB.Connection")
                    set commandT = Server.CreateObject("ADODB.Command")

                    connT.open session("dsn_cliente")
                    connT.cursorlocation=3

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60                 
                    commandT.CommandType = adCmdText
                    
					if si_tiene_modulo_OrCU <> 0 then
                        commandT.CommandText = "select codigo, descripcion from tarifas with(nolock) where codigo like ? +'%' and codigo <> ? + 'BASE' and tarifaCliente is null order by descripcion"
				 	else
                        commandT.CommandText = "select codigo, descripcion from tarifas with(nolock) where codigo like ? +'%' and codigo <> ? + 'BASE' order by descripcion"
				 	end if

                    commandT.Parameters.Append commandT.CreateParameter("@codigo",adVarChar,adParamInput,5,session("ncliente"))
                    commandT.Parameters.Append commandT.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente"))

                    set rstAux = commandT.Execute

                    DrawDiv "1","",""
                    CloseDiv

                    DrawDiv "1-detail","",""
                    DrawLabel "","",LitTarifa
                    DrawSelectCeldaWDNOLIT "CELDA","200","",0,"tarifa",rstAux,rtarifa,"codigo","descripcion","",""
                    DrawLabel "","","("&LitTarifaAviso&")"
                    CloseDiv

					rstAux.close	
                    connT.Close
                    set connT = nothing
                    set commandT = nothing
			%>			
        </div>
        </div>
        <div class="Section" id="S_DatosDV">
            <a href="#" rel="toggle[DatosDV]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitTiSerVen%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" id="DatosDV" style="display:none;">	   
				<%		
                    set connT = nothing
                    set commandT = nothing

                    set connT = Server.CreateObject("ADODB.Connection")
                    set commandT = Server.CreateObject("ADODB.Command")

                    connT.open session("dsn_cliente")
                    connT.cursorlocation=3

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandText = "select nserie,(substring(nserie,6,len(nserie)) + ' - ' + nombre) as descripcion from series with(nolock) where nserie like ? +'%' and tipo_documento='ALBARAN DE SALIDA'"
                    commandT.CommandType = adCmdText
                    commandT.Parameters.Append commandT.CreateParameter("@nserie",adVarChar,adParamInput,5,session("ncliente"))

                    set rstAux = commandT.Execute

				 	DrawSelectCelda "CELDA","200","",0,LitTiSeAlbC,"sac",rstAux,seralbcli,"nserie","descripcion","",""
					rstAux.close
		   			
                    set commandT = nothing
                    set commandT = Server.CreateObject("ADODB.Command")

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandText = "select nserie,(substring(nserie,6,len(nserie)) + ' - ' + nombre) as descripcion from series with(nolock) where nserie like ? + '%' and tipo_documento='FACTURA A CLIENTE'"
                    commandT.CommandType = adCmdText
                    commandT.Parameters.Append commandT.CreateParameter("@nserie",adVarChar,adParamInput,5,session("ncliente"))

                    set rstAux = commandT.Execute

				 	DrawSelectCelda "CELDA","200","",0,LitTiSeFacC,"sfc",rstAux,serfaccli,"nserie","descripcion","",""
					rstAux.close
		   			
                    set commandT = nothing
                    set commandT = Server.CreateObject("ADODB.Command")

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandText = "select nserie,(substring(nserie,6,len(nserie)) + ' - ' + nombre) as descripcion from series with(nolock) where nserie like ? + '%' and tipo_documento='PEDIDO DE CLIENTE'"
                    commandT.CommandType = adCmdText
                    commandT.Parameters.Append commandT.CreateParameter("@nserie",adVarChar,adParamInput,5,session("ncliente"))

                    set rstAux = commandT.Execute

				 	DrawSelectCelda "CELDA","200","",0,LitTiSePedC,"spc",rstAux,serpedcli,"nserie","descripcion","",""
					rstAux.close
		   			
                    set commandT = nothing
                    set commandT = Server.CreateObject("ADODB.Command")

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandText = "select nserie,(substring(nserie,6,len(nserie)) + ' - ' + nombre) as descripcion from series with(nolock) where nserie like ? + '%' and tipo_documento='PRESUPUESTO A CLIENTE'"
                    commandT.CommandType = adCmdText
                    commandT.Parameters.Append commandT.CreateParameter("@nserie",adVarChar,adParamInput,5,session("ncliente"))

                    set rstAux = commandT.Execute

				 	DrawSelectCelda "CELDA","200","",0,LitTiSePreC,"sprc",rstAux,serprecli,"nserie","descripcion","",""
					rstAux.close	
                    
                    connT.Close
                    set connT = nothing
                    set commandT = nothing
				%>			
        </div>
        </div>
        <div class="Section" id="S_DatosDC">
            <a href="#" rel="toggle[DatosDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitTiSerCom%> 
                   <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" id="DatosDC" style="display:none;">
	   
				<%
                    set connT = nothing
                    set commandT = nothing

                    set connT = Server.CreateObject("ADODB.Connection")
                    set commandT = Server.CreateObject("ADODB.Command")

                    connT.open session("dsn_cliente")
                    connT.cursorlocation=3

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandType = adCmdText
                    commandT.Parameters.Append commandT.CreateParameter("@nserie",adVarChar,adParamInput,5,session("ncliente"))
				
                    commandT.CommandText = "select nserie,(substring(nserie,6,len(nserie)) + ' - ' + nombre) as descripcion from series with(nolock) where nserie like ? + '%' and tipo_documento='ALBARAN DE PROVEEDOR'"
					set rstAux = commandT.Execute
				 	
                    DrawSelectCelda "CELDA","200","",0,LitTiSeAlbP,"sap",rstAux,seralbpro,"nserie","descripcion","",""
					rstAux.close

                    set commandT = nothing
                    set commandT = Server.CreateObject("ADODB.Command")

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandType = adCmdText
                    commandT.Parameters.Append commandT.CreateParameter("@nserie",adVarChar,adParamInput,5,session("ncliente"))
						   			
                    commandT.CommandText = "select nserie,(substring(nserie,6,len(nserie)) + ' - ' + nombre) as descripcion from series with(nolock) where nserie like ? + '%' and tipo_documento='FACTURA DE PROVEEDOR'"
					set rstAux = commandT.Execute
				 	
                    DrawSelectCelda "CELDA","200","",0,LitTiSeFacP,"sfp",rstAux,serfacpro,"nserie","descripcion","",""
					rstAux.close
		   			
                    set commandT = nothing
                    set commandT = Server.CreateObject("ADODB.Command")

                    commandT.ActiveConnection =connT
                    commandT.CommandTimeout = 60
                    commandT.CommandType = adCmdText
                    commandT.Parameters.Append commandT.CreateParameter("@nserie",adVarChar,adParamInput,5,session("ncliente"))
						   			
                    commandT.CommandText = "select nserie,(substring(nserie,6,len(nserie)) + ' - ' + nombre) as descripcion from series with(nolock) where nserie like ? + '%' and tipo_documento='PEDIDO A PROVEEDOR'"
					set rstAux = commandT.Execute
				 	
                    DrawSelectCelda "CELDA","200","",0,LitTiSePedP,"spp",rstAux,serpedpro,"nserie","descripcion","",""
					rstAux.close
                    connT.Close
                    set connT = nothing
                    set commandT = nothing
				%>			
        </div>
        </div>
		
		<%'mmg: modulo OrCU 
        if si_tiene_modulo_OrCU <> 0 then  
            %>
            <div class="Section" id="S_DatosCtr">
                <a href="#" rel="toggle[DatosCtr]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=litConfCont%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosCtr" style="display:none;">
	       
	                <input type="hidden" name="hcodCont" value="<%=enc.EncodeForHtmlAttribute(null_s(codCont))%>"/>				  
					    <%
                            EligeCelda "input","add","left","","",0,LitCodCont,"codCont",20,codCont			   			
                            EligeCelda "input","add","left","","",0,LitLogin,"login",20,login
                            EligeCelda "input","add","left","","",0,LitIpCon,"ipCon",20,ipCon

			   			    DrawPasswordCeldaResponsive "CELDAR7' maxlength='20' align='right'" ,"","",20,0,LitPass,"pass",pass
					   %>
            </div>
            </div>

		    <%
        end if

        'dgb: 07-01-2009 modulo Agroclub
        if si_tiene_modulo_Agroclub <> 0 then
        %>
            <div class="Section" id="S_DatosAgro">
                <a href="#" rel="toggle[DatosAgro]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=litAgroclub%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosAgro" style="display:none;">	        
					    <%
                            EligeCelda "input","add","left","","",0,litNCooperativa,"ncooperativa",10,ncooperativa
                            EligeCelda "input","add","left","","",0,litNLote,"nlote",10,nlote
                            EligeCelda "input","add","left","","",0,litVersion,"nversion",10,nversion
					    %>
            </div>
            </div>
        <%end if 
        if tieneTPV="SI" then
        %>
            <div class="Section" id="S_DatosCPC">
                <a href="#" rel="toggle[DatosCPC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitTiConfPC%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosCPC" style="display:none;">
	       
				    <%
				   
                         DrawDiv "3-sub", "background-color: #eae7e3", ""
                        %><label class="ENCABEZADOC" style="text-align:left"><%=LitImpR%></label>
                        
                        <%
                        CloseDiv
                        EligeCelda "input","add","left","","",0,LitPuerto,"puerto",5,puerto
                        EligeCelda "input","add","left","","",0,LitNum,"numPuerto",5,numPuerto
				    %>
            </div>
            </div>
		<%
        end if

        if si_tiene_modulo_ebesa<> 0 then
			%>
            <div class="Section" id="S_DatosPRM">
                <a href="#" rel="toggle[DatosPRM]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitPrim%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosPRM" style="display:none;">
					<%
                        EligeCelda "input","add","left","","",0,LitPc1,"pc1",5,null_z(rpc1)
                        EligeCelda "input","add","left","","",0,LitPri1,"pri1",5,null_z(rpri1)
                        EligeCelda "input","add","left","","",0,LitPc2,"pc2",5,null_z(rpc2)
                        EligeCelda "input","add","left","","",0,LitPri2,"pri2",5,null_z(rpri2)
                        EligeCelda "input","add","left","","",0,LitPc3,"pc3",5,null_z(rpc3)
                        EligeCelda "input","add","left","","",0,LitPri3,"pri3",5,null_z(rpri3)
                        EligeCelda "input","add","left","","",0,LitPc4,"pc4",5,null_z(rpc4)
                        EligeCelda "input","add","left","","",0,LitPri4,"pri4",5,null_z(rpri4)
                        EligeCelda "input","add","left","","",0,LitPc5,"pc5",5,null_z(rpc5)
                        EligeCelda "input","add","left","","",0,LitPri5,"pri5",5,null_z(rpri5)
					%>
            </div>
            </div>
		<%end if

        if si_tiene_modulo_ebesa<> 0 then
            %>  
            <div class="Section" id="S_DatosDSC">
                <a href="#" rel="toggle[DatosDSC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitDescuentos%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosDSC" style="display:none;">
	       
				<%
                    EligeCelda "input","add","left","","",0,LitDtol,"dtol",5,null_z(rdtol)
                    EligeCelda "input","add","left","","",0,LitDtoDia,"dtodia",5,null_z(rdtodia)
                    
                    DrawInputCeldaInput LitDiaDesde, "dtodiadesde","dtohoradesde", iif(isnull(rdtodiadesde) or rdtodiadesde="","",day(rdtodiadesde)&"/"&month(rdtodiadesde)&"/"&year(rdtodiadesde)),iif(isnull(rdtodiadesde) or rdtodiadesde="","",rhdesde&":"&rmdesde), "15", "15"
				    DrawCalendar "dtodiadesde"                        
                   
                    DrawInputCeldaInput LitDiaHasta, "dtodiahasta","dtohorahasta", iif(isnull(rdtodiahasta) or rdtodiahasta="","",day(rdtodiahasta)&"/"&month(rdtodiahasta)&"/"&year(rdtodiahasta)), iif(isnull(rdtodiahasta) or rdtodiahasta="","",rhhasta&":"&rmhasta), "15", "15"
				    DrawCalendar "dtodiahasta"
				
                    EligeCelda "check","add","","","",0,LitAplicar,"dtodial",8,""
                    EligeCelda "input","add","left","","",0,LitPwdManual,"pwdpvpmanual",25,rpwdpvpmanual
                    EligeCelda "input","add","left","","",0,LitDtoRegalo,"dtoregalo",25,null_z(rdtoregalo)
                    EligeCelda "input","add","left","","",0,LitPwdRegalo,"pwdregalo",25,rpwddtoregalo
                    EligeCelda "input","add","left","","",0,LitDtoEncargado,"dtoencargado",25,null_z(rdtoencargado)
                    EligeCelda "input","add","left","","",0,LitPwdEncargado,"pwdencargado",25,rpwddtoencargado
				%>
			
            </div>
            </div>
        <%
        end if
        '' Toni Climent 14-01-2009: Agergar SPAN Gestion de salones
		''TCD: 21-01-2009 Cargar o no valores si HORECAs en salones
        if es_hostelera <> 0 and mode <> "add" then
            
            %>
            <div class="Section" id="S_DatosSAL">
                <a href="#" rel="toggle[DatosSAL]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitSal%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosSAL" style="display:none;">
					<%
                        EligeCelda "input","add","left","","",0,LitSal1,"sal1",5,iif(es_hostelera = false,"0",null_z(rpri1))
                        EligeCelda "input","add","left","","",0,LitSal2,"sal2",5,iif(es_hostelera = false,"0",null_z(rpri2))
                        EligeCelda "input","add","left","","",0,LitSal3,"sal3",5,iif(es_hostelera = false,"0",null_z(rpri3))
                        EligeCelda "input","add","left","","",0,LitSal4,"sal4",5,iif(es_hostelera = false,"0",null_z(rpri4))
                        EligeCelda "input","add","left","","",0,LitSal5,"sal5",5,iif(es_hostelera = false,"0",null_z(rpri5))
					%>
            </div>
            </div>
		<%
        end if
        %>
		
	<%elseif mode="browse" then %>
        <table style="width: 100%;"></table>
        <div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['DatosDG','DatosDV','DatosDC','DatosCtr','DatosAgro','DatosCPC','DatosPRM','DatosOBJ','DatosDSC','DatosSAL']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['DatosDG','DatosDV','DatosDC','DatosCtr','DatosAgro','DatosCPC','DatosPRM','DatosOBJ','DatosDSC','DatosSAL']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
        </div>
		<input type="hidden" name="hcodigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rcodigo))%>" />
		<input type="hidden" name="hdomicilio" value="<%=enc.EncodeForHtmlAttribute(null_s(rdomicilio))%>" />
		<%' Inicio Borde Span%>		
            <div class="Section" id="S_DatosDG">
                <a href="#" rel="toggle[DatosDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitTiDatGen%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosDG">
	       
					<%
					
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitCodigo,"",trimCodEmpresa(rcodigo)				   		
				   		'mmg:OrCU
		   			    if si_tiene_modulo_OrCU <> 0 then		   			       
                             EligeCeldaResponsive "text","browse",clase,"","",0,"",LitInstalacion,"",rdescripcion
		   			    else			   		        
                             EligeCeldaResponsive "text","browse",clase,"","",0,"",LitDescripcion,"",rdescripcion
			   		    end if
		           			
                        strselectp1 = "select rsocial from clientes with (NOLOCK) where ncliente=?"
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitCliente,"",trimcodempresa(rncliente) & " "&DLookupP1(strselectp1, rncliente&"", adChar, 10, session("dsn_cliente")&"")
					   
						'FLM:20100215: filtro por el módulo de fidelización premium.
		                if si_tiene_modulo_ModFidelizacionPremium <> 0 then
						     strselectp1 = "select rsocial from centros with (NOLOCK) where ncentro=?"
                             EligeCeldaResponsive "text","browse",clase,"","",0,"",LitCentro,"",DLookupP1(strselectp1, rcodcentro&"", adVarChar, 10, session("dsn_cliente")&"")
						end if
					
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitDireccion,"",rdomicilio
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPoblacion,"",rpoblacion
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitProvincia,"",rprovincia
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPais,"",rpais
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTel1,"",rtelefono
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitFax,"",rfax
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiEmail,"",email
						
						''dgb: 29/09/2008 se anyade un campo para modulo Orcu del impuesto IVMDH
		   			    if si_tiene_modulo_OrCU <> 0 then
                            strselectp1 = "select nombre from comunidad_orcu with (NOLOCK) where codigo=?"
                            DrawCeldaResponsiveCustomClass "col-lg-4 col-md-6 col-sm-6 col-xs-12 visibilityHidden", "",clase,"","","",LitIVMDH,DLookupP1(strselectp1, civmdh&"", adVarChar, 10, session("dsn_cliente")&"")
						end if
					
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitObservaciones,"",pintar_saltos_espacios(robservaciones & "")						

                        if see_coordenates=1 then                          
                                DrawDiv "1","",""
                                DrawLabel "","",LitCoordenates                    
                                                   
                                DrawSpan clase,"","X ",""  
                                DrawSpan clase,"",iif(x_coordenate&""<>"",x_coordenate,"&nbsp;&nbsp;"),""
                                DrawSpan clase,""," Y ",""
                                DrawSpan clase,"",iif(y_coordenate&""<>"",y_coordenate,"&nbsp;&nbsp;"),""
                                                    
                                CloseDiv
                        end if
					
                        strselectp1 = "select descripcion from almacenes with (NOLOCK) where codigo=?"
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitAlmacen,"",DLookupP1(strselectp1, ralmacen&"", adVarChar, 10, session("dsn_cliente")&"")					
					
                        strselectp1 = "select descripcion from tarifas with (NOLOCK) where codigo=?"
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTarifa,"",DLookupP1(strselectp1, rtarifa&"", adVarChar, 10, session("dsn_cliente")&"")

                        set conn4 = nothing
                        set command4 = nothing

                        set conn4 = Server.CreateObject("ADODB.Connection")
                        set command4 =  Server.CreateObject("ADODB.Command")

                        conn4.open session("dsn_cliente")
                        conn4.cursorlocation=3

                        command4.ActiveConnection =conn4
                        command4.CommandTimeout = 60
                        command4.CommandText="select codigo from cajas where tienda=?"
                        command4.CommandType = adCmdText
                        command4.Parameters.Append command4.CreateParameter("@codigo",adVarChar,adParamInput,10,rcodigo)

                        set rstAux = command4.Execute

						if rstaux.eof then
							cajapaltpv=""
						else
							cajapaltpv=rstaux("codigo")
						end if

						rstaux.close
                        conn4.Close
                        set conn4 = nothing
                        set command4 = nothing

						DrawDiv "3","",""%><label><a class="CELDAREF" href="javascript:SaveFile('<%=enc.EncodeForJavascript(null_s(ralmacen))%>','<%=enc.EncodeForJavascript(null_s(cajapaltpv))%>','<%=enc.EncodeForJavascript(null_s(rcodigo))%>')"><%=LitAsignarPcTienda%></a></label><%
                        CloseDiv

                    set connVL=server.CreateObject("ADODB.Connection")
                    set cmdVL=server.CreateObject("ADODB.Command")
                    set rstVL=server.CreateObject("ADODB.recordset")
                    dsnM=ObtenDSNMixta(session("dsn_cliente"), dsnIlion)
                    connVL.open dsnM
	                connVL.cursorlocation=3
                    cmdVL.ActiveConnection = connVL
                    cmdVL.CommandType = adCmdStoredProc
                    cmdVL.CommandText="ValidateLoyaltyDeleteStore"
                    cmdVL.Parameters.Append cmdVL.CreateParameter("@companyId",adVarchar,,5,session("ncliente"))
                    cmdVL.Parameters.Append cmdVL.CreateParameter("@store",adVarchar,,10,p_ntienda)
                    set rstVL =cmdVL.Execute
	                if not rstVL.eof then
		               countusers=null_z(rstVL("countusers"))
                       countoperations=null_z(rstVL("countoperations"))
                       %>
                       <input type="hidden" name="countusers" value="<%=enc.EncodeForHtmlAttribute(null_s(countusers))%>" />
                       <input type="hidden" name="countoperations" value="<%=enc.EncodeForHtmlAttribute(null_s(countoperations))%>" />
                       <%
	                end if
	                connVL.close
                    set rstVL=nothing
                    set cmdVL=nothing
                    set connVL=nothing
				%>
				<iframe name="frFichConf" id="IdFichConf" style="visibility:hidden;width:200px;height:200px"></iframe>
            </div>
            </div>
            <div class="Section" id="S_DatosDV">
                <a href="#" rel="toggle[DatosDV]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitTiSerVen%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosDV" style="display:none;">
					<%
					    strselectp1 = "select nombre from series with (NOLOCK) where nserie=?"
			   			EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiSeAlbC,"",iif(seralbcli>"",trimCodEmpresa(seralbcli) & " - " & DLookupP1(strselectp1, seralbcli&"", adVarChar, 10, session("dsn_cliente")&""),"")
                       
						strselectp1 = "select nombre from series with (NOLOCK) where nserie=?"
		   				EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiSeFacC,"",iif(serfaccli>"",trimCodEmpresa(serfaccli) & " - " & DLookupP1(strselectp1, serfaccli&"", adVarChar, 10, session("dsn_cliente")&""),"")
                        
						strselectp1 = "select nombre from series with (NOLOCK) where nserie=?"
			   			EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiSePedC,"",iif(serpedcli>"",trimCodEmpresa(serpedcli) & " - " & DLookupP1(strselectp1, serpedcli&"", adVarChar, 10, session("dsn_cliente")&""),"")
                       
						strselectp1 = "select nombre from series with (NOLOCK) where nserie=?"
			   			EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiSePreC,"",iif(serprecli>"",trimCodEmpresa(serprecli) & " - " &DLookupP1(strselectp1, serprecli&"", adVarChar, 10, session("dsn_cliente")&""),"")
					%>
            </div>
            </div>
            <div class="Section" id="S_DatosDC">
                <a href="#" rel="toggle[DatosDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitTiSerCom%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosDC" style="display:none;">
					<%
					    strselectp1 = "select nombre from series with (NOLOCK) where nserie=?"
			   			EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiSeAlbP,"",iif(seralbpro>"",trimCodEmpresa(seralbpro) & " - " & DLookupP1(strselectp1, seralbpro, adVarChar, 10, session("dsn_cliente")&""),"")
                        
			   			strselectp1 = "select nombre from series with (NOLOCK) where nserie=?"
			   			EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiSeFacP,"",iif(serfacpro>"",trimCodEmpresa(serfacpro) & " - " & DLookupP1(strselectp1, serfacpro, adVarChar, 10, session("dsn_cliente")&""),"")
                       
			   	        strselectp1 = "select nombre from series with (NOLOCK) where nserie=?"
		   				EligeCeldaResponsive "text","browse",clase,"","",0,"",LitTiSePedP,"",iif(serpedpro>"",trimCodEmpresa(serpedpro) & " - " & DLookupP1(strselectp1, serpedpro, adVarChar, 10, session("dsn_cliente")&""),"")
					%>
            </div>
            </div>
		<%'mmg: modulo OrCU
        if si_tiene_modulo_OrCU <> 0 then
        %>
            <div class="Section" id="S_DatosCtr">
                <a href="#" rel="toggle[DatosCtr]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=litConfCont%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosCtr" style="display:none;">
					<%
			   			EligeCeldaResponsive "text","browse",clase,"","",0,"",LitCodCont,"",codCont
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitLogin,"",login
			   			EligeCeldaResponsive "text","browse",clase,"","",0,"",LitIpCon,"",ipCon
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPass,"",Pwd(pass)
					%>
            </div>
            </div>
		<%
        end if

        'dgb: 07-01-2009 modulo Agroclub
        if si_tiene_modulo_Agroclub <> 0 then
        %>
            <div class="Section" id="S_DatosAgro">
                <a href="#" rel="toggle[DatosAgro]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=litAgroclub%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosAgro" style="display:none;">
					<%
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",litNCooperativa,"",ncooperativa
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",litNLote,"",nlote
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",litVersion,"",nversion			   			
					%>
            </div>
            </div>
        <%end if 
        if tieneTPV="SI" then
        %>
            <div class="Section" id="S_DatosCPC">
                <a href="#" rel="toggle[DatosCPC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitTiConfPC%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosCPC" style="display:none;">
			    <%
                    DrawDiv "3-sub", "background-color: #eae7e3", ""
                        %><label class="ENCABEZADOC" style="text-align:left"><%=LitImpR%></label>
                        
                        <%
                    CloseDiv
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPuerto,"",puerto
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitNum,"",numPuerto
			   %>
            </div>
            </div>
        <%end if 
        if si_tiene_modulo_ebesa<> 0 then
        %>
            <div class="Section" id="S_DatosPRM">
                <a href="#" rel="toggle[DatosPRM]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitPrim%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosPRM" style="display:none;">
				    <%
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPc1,"",null_z(rpc1)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPri1,"",null_z(rpri1)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPc2,"", null_z(rpc2)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPri2,"", null_z(rpri2)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPc3,"", null_z(rpc3)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPri3,"", null_z(rpri3)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPc4,"", null_z(rpc4)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPri4,"", null_z(rpri4)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPc5,"", null_z(rpc5)
                        EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPri5,"", null_z(rpri5)
				    %>
            </div>
            </div>
        <%end if 
        if si_tiene_modulo_ebesa<> 0 then
            %>
            <div class="Section" id="S_DatosOBJ">
                <a href="#" rel="toggle[DatosOBJ]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitObj%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosOBJ" style="display:none;">
				    <%
                         EligeCelda "input","add","left","","",0,Litmes,"mes",5,""
                         EligeCelda "input","add","left","","",0,Litany,"any",5,""
                         DrawInputCeldaInsertar "", "", "", 15, 0, Litobjt, "obj", obj, "AnadirObj();", LitNuevoObj
                        %>
			    <iframe name='marcoObjetivos' id='frObjetivos' src='tiendas_obj.asp?mode=select&tienda=<%=enc.EncodeForHtmlAttribute(null_s(rcodigo))%>'></iframe>
            </div>
            </div>
        <%
        end if
        if si_tiene_modulo_ebesa<> 0 then
            
            %>
            <div class="Section" id="S_DatosDSC">
                <a href="#" rel="toggle[DatosDSC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitDescuentos%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosDSC" style="display:none;">
			    <%
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitDtol,"",null_z(rdtol)
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitDtoDia,"",null_z(rdtodia)
                    DrawTwoCeldasResponsive clase,"","","", LitDiaDesde, iif(isnull(rdtodiadesde) or rdtodiadesde="","",day(rdtodiadesde)&"/"&month(rdtodiadesde)&"/"&year(rdtodiadesde)), iif(isnull(rdtodiadesde) or rdtodiadesde="","",rhdesde&":"&rmdesde)
                    DrawTwoCeldasResponsive clase,"","","", LitDiaHasta, iif(isnull(rdtodiahasta) or rdtodiahasta="","",day(rdtodiahasta)&"/"&month(rdtodiahasta)&"/"&year(rdtodiahasta)), iif(isnull(rdtodiahasta) or rdtodiahasta="","",rhhasta&":"&rmhasta)
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitAplicar,"",iif(rdtodial,"Si","No")
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPwdManual,"",null_z(rpwdpvpmanual)
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitDtoRegalo,"",null_z(rdtoregalo)
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPwdRegalo,"",null_s(rpwddtoregalo)
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPwdRegalo,"",null_s(rpwddtoregalo)
                    EligeCeldaResponsive "text","browse",clase,"","",0,"",LitPwdEncargado,"",null_s(rpwddtoencargado)
			    %>
            </div>
            </div>
		<%
        end if
        ''Toni Climent 14-01-2009 SPAN de salones
		''TCD: 21-01-2009 Cargar o no valores si HORECAs en salones
        if es_hostelera <> 0 and mode <> "add" then
            %>
            <div class="Section" id="S_DatosSAL">
                <a href="#" rel="toggle[DatosSAL]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitSal%>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>
                </a>
            <div class="SectionPanel" id="DatosSAL" style="display:none;">
					<%
                        EligeCeldaResponsive "text","browse",clase,"","",0,"","sal1","",iif(es_hostelera = false,"0",null_z(rpri1)) 
                        EligeCeldaResponsive "text","browse",clase,"","",0,"","sal2","",iif(es_hostelera = false,"0",null_z(rpri2)) 
                        EligeCeldaResponsive "text","browse",clase,"","",0,"","sal3","",iif(es_hostelera = false,"0",null_z(rpri3))
                        EligeCeldaResponsive "text","browse",clase,"","",0,"","sal4","",iif(es_hostelera = false,"0",null_z(rpri4))
                        EligeCeldaResponsive "text","browse",clase,"","",0,"","sal5","",iif(es_hostelera = false,"0",null_z(rpri5))
					%>
            </div>
            </div>
        <%end if
	elseif mode="search" then
   end if
	set rst = nothing
	set rstAux = nothing
	set rstDom = nothing
	set rstOrCU = nothing
   %>
</form>
<%end if
connRound.close
set connRound = Nothing%>
<script type="text/javascript" id="mapas">
    var loadMaps=0;
    var marker ="";
    var x,y;

    var maps=""; 

    function InitMaps(){
    
        var mapOptions = {zoom: 15,center: new google.maps.LatLng(y,x)};
        maps = new google.maps.Map(document.getElementById("map-canvas"),mapOptions);
        maps.setCenter(new google.maps.LatLng(y,x));
        marker = new google.maps.Marker({
            position: new google.maps.LatLng(y,x),
            map: maps,
            draggable: true,
            title: "<%=LitPressLocation %>"               
        });

        myListener = google.maps.event.addListener(maps, 'click', function(event) {
            placeMarker(event.latLng);
            google.maps.event.removeListener(myListener);
        });
        myListener =  google.maps.event.addListener(marker, 'drag', function(event) {
            placeMarker(event.latLng);
            google.maps.event.removeListener(myListener);
        });

        function placeMarker(location) {
         
            marker.setPosition(location);
            maps.setCenter(location);
            var markerPosition = marker.getPosition();
   
            populateInputs(markerPosition);
            google.maps.event.addListener(marker, "drag", function (mEvent) {
                populateInputs(mEvent.latLng);
            });
        }

        function populateInputs(pos) {
            jQuery("#x_coordenate").val(pos.lng());
            jQuery("#y_coordenate").val(pos.lat());
        }
    }

    function showMaps(){
        var T = (jQuery(window).height() / 2) - (jQuery("#map-canvas").height() / 2) - parseInt(jQuery("#map-canvas").css("paddingTop"));
        var L = (jQuery(window).width() / 2) - (jQuery("#map-canvas").width() / 2) - parseInt(jQuery("#map-canvas").css("paddingLeft"));
        if (L <= 0) L = 0;
        var L2 = L+(jQuery("#map-canvas").width())+5;
        if (T <= 0) T = 0;
        jQuery("#mapToSelect").slideToggle(0);
        jQuery("#map-canvas").animate({ top: T + "px", left: L + "px" }, 500);
        jQuery(".backGroundClose").animate({ top: T + "px", left: L2 + "px" }, 500);
        if (loadMaps==0){
            InitMaps();
            loadMaps=1;
        }
        setTimeout('x=jQuery("#x_coordenate").val();y=jQuery("#y_coordenate").val();maps.setCenter(new google.maps.LatLng(y,x));marker.setPosition(new google.maps.LatLng(y,x));','300');
    }

    function loadCoordenate(){
        <% if si_tiene_modulo_Teekit <> 0 then %>
        var address="";
        var cY="",cX="";
        address=jQuery("#domicilio").val()+" "+jQuery("#poblacion").val()+" "+jQuery("#CP").val()
    
        if ((jQuery("#domicilio").val()+jQuery("#poblacion").val()+jQuery("#CP").val())!=""){
            if (loadMaps==0){
                InitMaps();
                loadMaps=1;
            }
            geocoder = new google.maps.Geocoder();
            geocoder.geocode({ 'address': address }, function (results, status) { jQuery("#x_coordenate").val(results[0].geometry.location.lng()); jQuery("#y_coordenate").val(results[0].geometry.location.lat()); });
                    
        }else{
            jQuery("#y_coordenate").val("");
            jQuery("#x_coordenate").val("");
        }
        <% end if %>
    }

    function coordenada(position) {
        x=position.coords.longitude;
        y=position.coords.latitude;
        jQuery("#x_coordenate").val(x);
        jQuery("#y_coordenate").val(y);
    
    }
</script>
</body>
</html>