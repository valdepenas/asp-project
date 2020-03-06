<%@ Language=VBScript %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="ilion.inc" -->
<!--#include file="cache.inc" -->
<!--#include file="adovbs.inc" -->
<!--#include file="constantes.inc" -->
<!--#include file="sel_empresas.inc" -->
<!--#include file="tablas.inc" -->
<!--#include file="mensajes.inc" -->
<!--#include file="calculos.inc" -->
<!--#include file="varios2.inc" -->
<!--#include file="varios.inc" -->
<%dim parametroD
parametroD=replace(replace(request.querystring("D")&"","{",""),"}","")
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
if session("folder") & "" = "" then
    session("folder") = "ilion"
    title = ""
    if parametroD & "" <> "" then
        set conn = Server.CreateObject("ADODB.Connection")
	    set command =  Server.CreateObject("ADODB.Command")
	    conn.open dsnilion
	    command.ActiveConnection =conn
	    command.CommandTimeout = 0
	    command.CommandText="GetConfigDataDistributor"
	    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        if len(parametroD) > 5 then
	        command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, "")
        else
            command.Parameters.Append command.CreateParameter("@ndistributor", adVarChar, adParamInput, 5, parametroD)
        end if
        if len(parametroD) > 5 then
            command.Parameters.Append command.CreateParameter("@id", adVarChar, adParamInput, 36, parametroD)
        else
            command.Parameters.Append command.CreateParameter("@id", adVarChar, adParamInput, 36, "")
        end if
        command.Parameters.Append command.CreateParameter("@dtid", adSmallInt, adParamInput, , "2")

        set rst = command.execute

        if not rst.eof then
            session("folder") = rst("FOLDERCSS")
            title = rst("DISTRIBUTORTITLE")
        end if

	    rst.close
        conn.close
        set rst = nothing
        set command = nothing
        set conn = nothing
    end if
end if%>
<!--#include file="borrTablTemp.inc" -->
<!--#include file="cabecera.inc" -->
<!--#include file="ico.inc" -->
<%' VGR  06/05/03	: Mostrar solo las empresas cuya FBAJA sea nula.

if request.querystring("exFich")>"" then
	exFich=request.querystring("exFich")
else
	exFich="0"
end if

if request.querystring("linea1")<>"" then
	tpv=request.querystring("linea1")
else
	tpv=""
end if

if request.querystring("linea2")<>"" then
	caja=request.querystring("linea2")
else
	caja=""
end if

if request.querystring("linea3")<>"" then
	empresa=request.querystring("linea3")
else
	empresa=""
end if

if exFich="1" then
	session("f_tpv") = session("ncliente") & tpv
	session("f_caja")= session("ncliente") & caja
	session("f_empr")= empresa
else
	session("f_tpv") = ""
	session("f_caja")= ""
	session("f_empr")= ""
end if%>
<html>
<head>
<title><%=TituloVentana%></title>
<style type="text/css">
    .bienvenido {font-Family: <%=lletra%>;font-size: 16.0pt;text-align:center;font-weight: bold;color:rgb(246,165,0)}
	.mensaje {font-Family: <%=lletra%>;font-size: 10.0pt;text-align:center; color: white;}
	.TITULAR {font-Family: <%=lletra%>;font-size: 16.0pt;text-align:center;font-weight: bold;}
	.CABECERA {font-Family: <%=lletra%>;font-size: 8.0pt;}
	.CELDA {font-Family: <%=lletra%>;font-size: 8.0pt;}
	.BOLDLEFT {font-Family: <%=lletra%>;font-size: 8.0pt;text-align: left;font-weight: bold;}
	.NUEVO {font-Family: <%=lletra%>;font-size: 8.0pt;text-align: left;	font-weight: bold;}
</style>
<!--#include file="styles/access.css.inc"-->
<script type="text/javascript" language="javascript" src="jfunciones.js"></script>
<script type="text/javascript" language="JavaScript">
var fallo=false;
function DoTheRefresh() 
{ 
    document.location=document.location;
} 
</script>
</head>
<body>
<%if accesoPagina(session.sessionid,session("usuario"))=1 then
	set rs_stock = Server.CreateObject("ADODB.Recordset")

    go = request.QueryString("go")
    id_user = request.QueryString("id_user")
''ricardo 13-5-2003 se borraran las tablas temporales para cuando salgamos de una empresa
if request("viene")="otra_empresa" then
	BorrTablTemp
end if

if request("Ncliente") = "" then
	''ricardo 13/11/2003

    set rstAux = Server.CreateObject("ADODB.Recordset")
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open dsnilion
    conn.cursorlocation=3
    command.ActiveConnection =conn
    command.CommandTimeout = 60

    strmaxcab="select mostrar_cabecera,mostrar_maximizado from configuracion with(nolock) where PATHACCESO=?"
    command.CommandText=strmaxcab
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@path",adVarChar,adParamInput,20,carpetaproduccion)
    
    set rstAux= command.Execute               

	mostrar_cabecera=0
	mostrar_maximizado=0
	
	if not rstAux.eof then
		mostrar_cabecera=nz_b(rstAux("mostrar_cabecera"))
		mostrar_maximizado=nz_b(rstAux("mostrar_maximizado"))
	end if
	rstAux.close
    conn.close
    set conn    =  nothing
    set command =  nothing      
	set rstAux=nothing

    mostrar_cabecera=1

	if mostrar_cabecera<>0 then
		MuestraCabecera "sel_empresas"
	end if

    ruta=GenerarURL
    Session("net")=""
    %>
    <iframe id='frAnulaSessionNet' src='/<%=CarpetaProduccionX%>/desactiva.aspx' width="100" height="100" frameborder="0" scrolling="no" style="display: none"></iframe>	
    <iframe id='frAnulaSessionNet45' src='/<%=CarpetaProduccionX4%>5/desactiva.aspx' width="100" height="100" frameborder="0" scrolling="no"style="display: none"></iframe>	
    <iframe id="initSe4" src="/<%=CarpetaProduccionX4%>/init.aspx" width="100" height="100" frameborder="no" scrolling="no" noresize="noresize" style="display: none"></iframe>
    <br /><br /><br /><br />
	<table width="100%" border="0" cellspacing="1" cellpadding="1">
		<tr><td></td></tr>
		<tr>
			<td width="2px">&nbsp;</td>
			<td colspan="2">
				<%PintarCabecera "sel_empresas.asp"%>
			</td>
			<td width="2px">&nbsp;</td>
		</tr>
		<tr valign="top">
			<td width="2px">&nbsp;</td>
			<td align="left"><img src="images/<%=ImgSelEmpresa%>" <%=ParamImgSelEmpresa%> alt="" title=""/></td>
			<td width="93%" class="CELDA" align="left">
				<table width="100%" border="0" cellspacing="1" cellpadding="1">				
				  <%'Recuperar parametro de usuario'
					strselect_l = "select parametros from param_usuario with(nolock) where usuario=? and objeto=?"
                    ParametrosSelEmpresas = DLookupP2(strselect_l, session("usuario"), adVarChar, 50, "998", adVarChar, 3, DSNIlion)					
					
				    'Verificar que el dominio de acceso no esta excluido'
					dim strDomain, isDomainExcluded				
					strDomain = Request.ServerVariables("SERVER_NAME")
					
					set conn = Server.CreateObject("ADODB.Connection")
					set command =  Server.CreateObject("ADODB.Command")
					conn.open dsnilion
					command.ActiveConnection = conn
					command.CommandText = "INST_GET_DOMAIN_EXCLUDED"
					command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
					command.Parameters.Append command.CreateParameter("@ID", adBigInt, adParamInput, , Null)
					command.Parameters.Append command.CreateParameter("@DOMAIN_DESCRIPTION", adVarChar, adParamInput, 255, Null)
					command.Parameters.Append command.CreateParameter("@DOMAIN", adVarChar, adParamInput, 255, strDomain)

					set rst = command.execute

					isDomainExcluded = 0
					
					if not rst.eof then
						isDomainExcluded = 1
					end if

					rst.close
					conn.close
					
					set rst = nothing
					set command = nothing
					set conn = nothing
					
					'Obtenemos la instancia si el dominio no esta excluido'
					dim ninstance, instanceExits
					instanceExits = 0
														
					if(not CBool(isDomainExcluded)) then
						set conn = Server.CreateObject("ADODB.Connection")
						set command =  Server.CreateObject("ADODB.Command")
						conn.open dsnilion
						command.ActiveConnection = conn
						command.CommandText = "INST_GET_INSTANCE_BY_FILTERS"
						command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
						command.Parameters.Append command.CreateParameter("@NINSTANCE", adInteger, adParamInput, , Null)
						command.Parameters.Append command.CreateParameter("@NAME", adVarChar, adParamInput, 10, Null)
						command.Parameters.Append command.CreateParameter("@NPROJECT", adVarChar, adParamInput, 5, Null)
						command.Parameters.Append command.CreateParameter("@PRODUCT_BBDD", adVarChar, adParamInput, 100, Null)
						command.Parameters.Append command.CreateParameter("@PROJECT_BBDD", adVarChar, adParamInput, 100, Null)
						command.Parameters.Append command.CreateParameter("@URL_SITE", adVarChar, adParamInput, 255, strDomain)
						command.Parameters.Append command.CreateParameter("@PATH", adVarChar, adParamInput, 255, Null)
						command.Parameters.Append command.CreateParameter("@STATE", adInteger, adParamInput, , Null)
						command.Parameters.Append command.CreateParameter("@DSNPROY", adVarChar, adParamInput, 255, Null)
						command.Parameters.Append command.CreateParameter("@DSNPROD", adVarChar, adParamInput, 255, Null)
						command.Parameters.Append command.CreateParameter("@ACTIVE", adBoolean, adParamInput, , Null)
						command.Parameters.Append command.CreateParameter("@CREATE_USER", adVarChar, adParamInput, 80, Null)
						command.Parameters.Append command.CreateParameter("@CREATE_DATE", adDate, adParamInput, , Null)
						command.Parameters.Append command.CreateParameter("@USER_SFTP", adVarChar, adParamInput, 80, Null)
						command.Parameters.Append command.CreateParameter("@PASSWORD_SFTP", adVarChar, adParamInput, 80, Null)
						command.Parameters.Append command.CreateParameter("@DEPLOY", adBoolean, adParamInput, , Null)

						set rst = command.execute

						if not rst.eof then
							instanceExits = 1
							ninstance = rst("Ninstance")
						end if

						rst.close
						conn.close
						
						set rst = nothing
						set command = nothing
						set conn = nothing
					end if
					
					'Filtrado de empresas por instancia'
					dim anyCompanyObtained
					anyCompanyObtained = 0
					
					if(CBool(instanceExits)) then
						set conn = Server.CreateObject("ADODB.Connection")
						set command =  Server.CreateObject("ADODB.Command")
						conn.open dsnilion
						command.ActiveConnection = conn
						command.CommandText = "INST_GET_INSTANCE_CLIENTS"
						command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
						command.Parameters.Append command.CreateParameter("@NINSTANCE", adInteger, adParamInput, , ninstance)
						command.Parameters.Append command.CreateParameter("@USER_NAME", adVarChar, adParamInput, 15, session("usuario"))
						command.Parameters.Append command.CreateParameter("@PARAMETER_EXISTS", adBoolean, adParamInput, , instr(ParametrosSelEmpresas,"?tyg=1")>0)
						
						set rst = command.execute

						while not rst.eof
							anyCompanyObtained = 1
							
							DrawFila color_blau
							%><td width="10px"><font face='Arial'> ></font></td><%
							if go = "open" and rst(0)<> "00000" then
							    h_ref = "http://79.125.49.133:8080/openbravo/com.everis.initiative.cloud.sso/Login.html?user=" & enc.EncodeForHtmlAttribute(id_user)
							elseif rst(0)="01232" then
                                h_ref = "http://79.125.49.133:8080/openbravo/com.everis.initiative.cloud.sso/Login.html?user=Openbravo"
                            else
							    h_ref = "sel_empresas.asp?Ncliente=" & rst(0) & "&d=" & enc.EncodeForHtmlAttribute(parametroD)
							end if
							DrawCeldaHref "CELDA","left",false,rst("Rsocial"),h_ref
							%></tr><%
							
							rst.MoveNext
						wend

						if(not CBool(anyCompanyObtained)) then
							response.Write(LITINSTANCENOTCOMPANYS)
						end if
						
						rst.close
						conn.close
						
						set rst = nothing
						set command = nothing
						set conn = nothing
					end if
					
					'Capturamos las empresas del cliente'
					if(CBool(isDomainExcluded) or not CBool(instanceExits)) then
						set rs_empresas = Server.CreateObject("ADODB.Recordset")
					   
						param=""
						set command = nothing
						set conn = Server.CreateObject("ADODB.Connection")
						set command =  Server.CreateObject("ADODB.Command")
						conn.open DsnIlion
						conn.cursorlocation=3
						command.ActiveConnection =conn
						command.CommandTimeout = 60
						if instr(ParametrosSelEmpresas,"?tyg=1")>0 then                        
							'rs_empresas.open "Select CU.*,C.* From Clientes_Users As CU with(nolock),Clientes As C with(nolock) Where CU.fbaja is null and CU.Usuario='" + session("usuario") + "' And C.Ncliente=CU.NCliente  Order By Rsocial", DsnIlion, adOpenKeyset, adLockOptimistic
							strselect="Select CU.*,C.* From Clientes_Users As CU with(nolock),Clientes As C with(nolock) Where CU.fbaja is null and CU.Usuario=? And C.Ncliente=CU.NCliente Order By Rsocial"
							command.CommandText=strselect
							command.CommandType = adCmdText
							command.Parameters.Append command.CreateParameter("@usu",adVarChar,adParamInput,50,session("usuario"))    
							set rs_empresas= command.Execute               
						else
							strselect="Select CU.*,C.* From Clientes_Users As CU with(nolock),Clientes As C with(nolock) Where CU.fbaja is null and CU.Usuario=? And C.Ncliente=CU.NCliente and (CU.cliente_int is null and CU.proveedor_int is null) Order By Rsocial"
							command.CommandText=strselect
							command.CommandType = adCmdText
							command.Parameters.Append command.CreateParameter("@usu",adVarChar,adParamInput,50,session("usuario"))    
							set rs_empresas= command.Execute       

						end if

						while not rs_empresas.EOF
							DrawFila color_blau
								%><td width="10px"><font face='Arial'> ></font></td><%
								if go = "open" and rs_empresas(0)<> "00000" then
									h_ref = "http://79.125.49.133:8080/openbravo/com.everis.initiative.cloud.sso/Login.html?user=" & enc.EncodeForHtmlAttribute(id_user)
								elseif rs_empresas(0)="01232" then
									h_ref = "http://79.125.49.133:8080/openbravo/com.everis.initiative.cloud.sso/Login.html?user=Openbravo"
								else
									h_ref = "sel_empresas.asp?Ncliente=" & rs_empresas(0) & "&d=" & enc.EncodeForHtmlAttribute(parametroD)
								end if
								DrawCeldaHref "CELDA","left",false,rs_empresas("Rsocial"),h_ref
							%></tr><%
							rs_empresas.MoveNext
						wend
						rs_empresas.close                            
						conn.close
						set conn    =  nothing
						set command =  nothing 
					end if%>
				</table>
			</td>
			<td width="2px">&nbsp;</td>
		</tr>
	</table><br/><br/><br/><%
	si_tienda_cli=false
	si_tienda_pro=false
	
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open DsnIlion
    conn.cursorlocation=3
    command.ActiveConnection =conn
    command.CommandTimeout = 60

    strselect="Select CU.*,C.* From Clientes_Users As CU,Clientes As C Where CU.Usuario=? And C.Ncliente=CU.NCliente and CU.cliente_int is null and CU.proveedor_int is null and CU.ncliente<>'00000' Order By Rsocial"

    command.CommandText=strselect
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@usu",adVarChar,adParamInput,50,session("usuario"))
    
    set rs_empresas= command.Execute               
	
	if not rs_empresas.eof then
		while not rs_empresas.eof
			if VerObjeto2(OBJAccesosCli,rs_empresas("ncliente"))=true then
				si_tienda_cli=true
			end if
			if VerObjeto2(OBJAccesosPro,rs_empresas("ncliente"))=true then
				si_tienda_pro=true
			end if
			rs_empresas.movenext
		wend
	end if
	rs_empresas.close
    conn.close
    set conn    =  nothing
    set command =  nothing      
	%><table width="40%" border="0" cellspacing="1" cellpadding="1" align="center"><%
	DrawFila color_blau

    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open DsnIlion
    conn.cursorlocation=3
    command.ActiveConnection =conn
    command.CommandTimeout = 60

	if instr(ParametrosSelEmpresas,"?tyg=1")>0 then
		strselect="Select * From Clientes_Users with(nolock) Where Usuario=?"
        command.CommandText=strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@usu",adVarChar,adParamInput,50,session("usuario"))
        set rs_empresas= command.Execute

	else
		strselect="Select * From Clientes_Users with(nolock) Where Usuario=? and (cliente_int is not null or proveedor_int is not null)"
        command.CommandText=strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@usu",adVarChar,adParamInput,50,session("usuario"))
        set rs_empresas= command.Execute

	end if
	'rs_empresas.open strselect, DsnIlion, adUseClient, adLockReadOnly
	if not rs_empresas.eof or si_tienda_cli=true or si_tienda_pro=true then
		%><td width="50%" class="CELDA" align="center">
			<a class="CELDA" href="sel_tiendas.asp?si_tienda_cli=<%=si_tienda_cli%>&si_tienda_pro=<%=si_tienda_pro%>">
				<img src="images/<%=ImgIrATiendas%>" <%=ParamImgIrATiendas%> alt="IR A TIENDAS" title="IR A TIENDAS"/></a>
		</td><%
	end if
	rs_empresas.close            	
    conn.close
    set conn    =  nothing
    set command =  nothing      
	'IML 23/01/2004  : Acceso Asesoria
	ncliente_int=""
	ncliente=""

    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open DsnIlion
    conn.cursorlocation=3
    command.ActiveConnection =conn
    command.CommandTimeout = 60

	strselect="select a.*,c.cif from accesos_int a with(nolock), clientes c with(nolock) where a.ncliente=c.ncliente and a.usuario=? and a.fbaja is null order by a.ncliente"

    command.CommandText=strselect
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@usu",adVarChar,adParamInput,50,session("usuario"))
    set rs_empresas= command.Execute

	'rs_empresas.open strselect, DsnIlion, adUseClient, adLockReadOnly
	if not rs_empresas.eof then
		href_logo="configuracion/muestra_logo.asp?cif=" & rs_empresas("ncliente")&rs_empresas("cif") & "&empresa=" & rs_empresas("ncliente")
		href_asesoria="asesoria/sel_asesoria.asp?ncliente="&ncliente&"&cliente_int="&ncliente_int
		%><td width="50%" class="CELDA" align="center">
			<a href="<%=href_asesoria%>"><img src="<%=href_logo%>" align="center" width="100" height="50" border="0" alt="IR A ASESORÍA" title="IR A ASESORÍA"/></a>
		</td><%
	end if
	%></tr></table><%
	'FIN IML 23/01/2004  : Acceso Asesoria
	rs_empresas.close
    conn.close
    set conn    =  nothing
    set command =  nothing     
else

	p_Ncliente=request("Ncliente")

	'Se vuelve a comprobar que el par NCLIENTE-USUARIO es el correcto	
    set rs_check = Server.CreateObject("ADODB.Recordset")
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")
    conn.open DsnIlion
    conn.cursorlocation=3
    command.ActiveConnection =conn
    command.CommandTimeout = 60

	strselect="Select * From Clientes_Users with(nolock) Where Ncliente=? And Usuario=? and fbaja is null"

    command.CommandText=strselect
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@ncli",adChar,adParamInput,5,p_Ncliente)
    command.Parameters.Append command.CreateParameter("@usu",adVarChar,adParamInput,50,session("usuario"))
    set rs_check= command.Execute

	if rs_check.EOF then
		rs_check.Close
        conn.close
        set conn    =  nothing
        set command =  nothing     
        CheckCadena ""
		%><script type="text/javascript" languaje="javascript">
		      document.location = "/<%=carpetaproduccion%>/desactiva.asp?mode=12&d=<%=enc.EncodeForJavascript(parametroD)%>";
		</script><%
	else
		'Permisos de ejecucion
		rs_check.Close
        conn.close
        set conn    =  nothing
        set command =  nothing     

		set rs_Clientes = Server.CreateObject("ADODB.Recordset")
		set rs_Clientes2 = Server.CreateObject("ADODB.Recordset")

        set command = nothing
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open DsnIlion
        conn.cursorlocation=3
        command.ActiveConnection =conn
        command.CommandTimeout = 60

        strselect="Select * From Clientes with(nolock) Where Ncliente=?"

        command.CommandText=strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@ncli",adChar,adParamInput,5,p_Ncliente)
        set rs_Clientes= command.Execute

		session("Ncliente") = rs_Clientes("ncliente")
		
		'Informar sessionStorage & cookies
		%><script type="text/javascript" languaje="javascript">
			if (sessionStorage.getItem("ncompany") === null)
			{
				sessionStorage.setItem("ncompany", "<%=p_Ncliente%>");
			}
			else
			{
				sessionStorage.ncompany = "<%=p_Ncliente%>";
			}
			
			var cookieName = 'companyCookie';
			var cookieValue = "<%=p_Ncliente%>"
			var tomorrow = new Date();
			tomorrow.setDate(tomorrow.getDate() + 1);
			var domain = "<%=Request.ServerVariables("SERVER_NAME")%>";

			document.cookie = cookieName +"=" + cookieValue + ";expires=" + tomorrow + ";domain=" + domain + ";path=/";
		</script><%
	
		session("dsn_cliente") = rs_Clientes("DSN")
		session("backendListados") = rs_Clientes("DSNLISTADOS") 'JAR - 18/10/07 DSN para listados.
		session("empresa") = ""
		session("empresa") = rs_Clientes("Rsocial")
		session("lenguaje") = rs_Clientes("lenguaje")
		session("caracteres") = rs_Clientes("caracteres")
		'dgb: nueva variable
		session("NetEstilo")= rs_Clientes("personalizacion")
		
		'JCI 20/09/2010: Asignación al entorno del LCID de la empresa
		session.LCID=rs_Clientes("LCID")
		'dgb 08/11/2010:  asignamos el LOCALE
		session("locale")=rs_Clientes("LOCALE")
		
		Auditar session("ncliente"),session("usuario"),session("usuario2"),"ENTRADA",Request.ServerVariables(CLIENT_IP),Request.ServerVariables("REMOTE_HOST"),Request.ServerVariables("HTTP_USER_AGENT"),DSNIlion
		rs_Clientes.Close
        conn.close
        set conn    =  nothing
        set command =  nothing     

		''JA: 12/05/07 CREAR REGISTRO EN TABLA VARSID
		'rs_Clientes.open "delete from VARSID where ncliente='" & session("Ncliente") & "' and usuario='" & session("usuario") & "'", DsnIlion, adOpenKeyset, adLockOptimistic
		'rs_Clientes.open "insert into VARSID values ('" & session("usuario") & "','" & session("ncliente") & "'," & session.sessionid & ",0)", DsnIlion, adOpenKeyset, adLockOptimistic
		''FIN

        set conn = Server.CreateObject("ADODB.Connection")
	    set command =  Server.CreateObject("ADODB.Command")
	    conn.open dsnilion
	    command.ActiveConnection =conn
	    command.CommandTimeout = 0
	    command.CommandText="GetShowSelApp"
	    command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))

        set rs_distrib = command.execute

        paso=false
        pageSelApp=""

	    if not rs_distrib.eof then
            select case rs_distrib("SHOWSELAPP")
                case 0
                    paso = false
                case 1
                    paso = true
            end select

            pageSelApp=rs_distrib("PAGESELAPP")

	    end if
		
	    rs_distrib.close
        conn.close
        set rs_distrib = nothing
        set command = nothing
        set conn = nothing

        if paso then
        
            if pageSelApp<>"" then

                %><script type="text/javascript" language="javascript">location.href = "<%=pageSelApp %>?d=<%=enc.EncodeForJavascript(parametroD)%>";</script><%
            else
                %><script type="text/javascript" language="JavaScript">location.href = "sel_app.asp?d=<%=enc.EncodeForJavascript(parametroD)%>";</script><%
            end if

        end if

        version_tpv=null
		tiene_tpv=0
		if session("ncliente")<>SISTEMA_GESTION then
			entrar=0
			if session("ncliente")<>SISTEMA_GESTION then
				
                set command = nothing
                set conn = Server.CreateObject("ADODB.Connection")
                set command =  Server.CreateObject("ADODB.Command")
                conn.open DsnIlion
                conn.cursorlocation=3
                command.ActiveConnection =conn
                command.CommandTimeout = 60

	            strselect="select dsn from clientes with(nolock) where ncliente=?"

                command.CommandText=strselect
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@ncli",adChar,adParamInput,5,p_Ncliente)
                set rs_Clientes= command.Execute

				if not rs_Clientes.eof then
					entrar=1
				else
					entrar=0
				end if
				rs_Clientes.close
                conn.close
                set rs_distrib = nothing
                set command = nothing
                set conn = nothing
			end if
			if entrar=1 and session("ncliente")<>SISTEMA_GESTION and ComprobarSiEsAsoc()=0 then
				
                set command = nothing
                set conn = Server.CreateObject("ADODB.Connection")
                set command =  Server.CreateObject("ADODB.Command")
                conn.open session("dsn_cliente")
                conn.cursorlocation=3
                command.ActiveConnection =conn
                command.CommandTimeout = 60

                strselect = "select control_stock from empresas with(nolock) where cif like ?+'%'"
                command.CommandText=strselect
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@cif",adVarChar,adParamInput,20,session("ncliente"))

                set rs_stock = command.Execute

				if not rs_stock.eof then
					 if rs_stock("control_stock")= true then
							session("control_stock") = "activado"
					 else
							session("control_stock") = "desactivado"
					 end if
				end if
				rs_stock.close
                conn.close
                set rs_distrib = nothing
                set command = nothing
                set conn = nothing

			end if

			'Comprobar que existen tpv para esa empresa, entonces leeremos la configuracion que tenga en el fichero cetel.tpv

			if entrar=1 and session("ncliente")<>SISTEMA_GESTION and session("dsn_cliente")<>"" and ComprobarSiEsAsoc()=0 then
				set rs_tpv = Server.CreateObject("ADODB.Recordset")
				set command = nothing
                set conn = Server.CreateObject("ADODB.Connection")
                set command =  Server.CreateObject("ADODB.Command")
                conn.open session("dsn_cliente")
                conn.cursorlocation=3
                command.ActiveConnection =conn
                command.CommandTimeout = 60
                                        
                strselect = "Select * from tpv with(nolock) where tpv like ?+'%'"
                command.CommandText=strselect
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@tpv",adVarChar,adParamInput,8,session("ncliente"))

                set rs_tpv = command.Execute

				if rs_tpv.eof then
					tiene_tpv=0
				else
					tiene_tpv=1

                    'sacar la version menor
                    version_aux=null
                    while not rs_tpv.EOF
                        version_aux = rs_tpv("VERSION_TPV")
                        if not isNull(version_aux) and version_aux<>"" then
                            version_aux=replace(version_aux&"",",",".")
                                        
                            if InStr(version_aux,".")>0 then
                                version_aux=Mid(version_aux,1,InStr(version_aux,".")-1)
                            end if

                            if isNull(version_tpv) or version_aux < version_tpv then
                                version_tpv = version_aux
                            end if
                        end if
                        rs_tpv.movenext
		            wend

				end if
				rs_tpv.close
                conn.close
                set command = nothing
                set conn = nothing
				set rs_tpv=nothing                    
			end if
		end if
		        ''ricardo 23-4-2010 los navegadores que no soporte la lectura de fichero que no lo hagan
                dim u,b,v
                set u=Request.ServerVariables("HTTP_USER_AGENT")
                set b=new RegExp
                set v=new RegExp
                b.Pattern="android|avantgo|blackberry|blazer|compal|elaine|fennec|hiptop|ip(hone|od)|iris|kindle|lge |maemo|midp|mmp|mobile|o2|opera m(ob|in)i|palm( os)?|p(ixi|re)\/|plucker|pocket|psp|smartphone|symbian|treo|up\.(browser|link)|vodafone|wap|windows ce; (iemobile|ppc)|xiino"
                v.Pattern="1207|6310|6590|3gso|4thp|50[1-6]i|770s|802s|a wa|abac|ac(er|oo|s\-)|ai(ko|rn)|al(av|ca|co)|amoi|an(ex|ny|yw)|aptu|ar(ch|go)|as(te|us)|attw|au(di|\-m|r |s )|avan|be(ck|ll|nq)|bi(lb|rd)|bl(ac|az)|br(e|v)w|bumb|bw\-(n|u)|c55\/|capi|ccwa|cdm\-|cell|chtm|cldc|cmd\-|co(mp|nd)|craw|da(it|ll|ng)|dbte|dc\-s|devi|dica|dmob|do(c|p)o|ds(12|\-d)|el(49|ai)|em(l2|ul)|er(ic|k0)|esl8|ez([4-7]0|os|wa|ze)|fetc|fly(\-|_)|g1 u|g560|gene|gf\-5|g\-mo|go(\.w|od)|gr(ad|un)|haie|hcit|hd\-(m|p|t)|hei\-|hi(pt|ta)|hp( i|ip)|hs\-c|ht(c(\-| |_|a|g|p|s|t)|tp)|hu(aw|tc)|i\-(20|go|ma)|i230|iac( |\-|\/)|ibro|idea|ig01|ikom|im1k|inno|ipaq|iris|ja(t|v)a|jbro|jemu|jigs|kddi|keji|kgt( |\/)|klon|kpt |kwc\-|kyo(c|k)|le(no|xi)|lg( g|\/(k|l|u)|50|54|e\-|e\/|\-[a-w])|libw|lynx|m1\-w|m3ga|m50\/|ma(te|ui|xo)|mc(01|21|ca)|m\-cr|me(di|rc|ri)|mi(o8|oa|ts)|mmef|mo(01|02|bi|de|do|t(\-| |o|v)|zz)|mt(50|p1|v )|mwbp|mywa|n10[0-2]|n20[2-3]|n30(0|2)|n50(0|2|5)|n7(0(0|1)|10)|ne((c|m)\-|on|tf|wf|wg|wt)|nok(6|i)|nzph|o2im|op(ti|wv)|oran|owg1|p800|pan(a|d|t)|pdxg|pg(13|\-([1-8]|c))|phil|pire|pl(ay|uc)|pn\-2|po(ck|rt|se)|prox|psio|pt\-g|qa\-a|qc(07|12|21|32|60|\-[2-7]|i\-)|qtek|r380|r600|raks|rim9|ro(ve|zo)|s55\/|sa(ge|ma|mm|ms|ny|va)|sc(01|h\-|oo|p\-)|sdk\/|se(c(\-|0|1)|47|mc|nd|ri)|sgh\-|shar|sie(\-|m)|sk\-0|sl(45|id)|sm(al|ar|b3|it|t5)|so(ft|ny)|sp(01|h\-|v\-|v )|sy(01|mb)|t2(18|50)|t6(00|10|18)|ta(gt|lk)|tcl\-|tdg\-|tel(i|m)|tim\-|t\-mo|to(pl|sh)|ts(70|m\-|m3|m5)|tx\-9|up(\.b|g1|si)|utst|v400|v750|veri|vi(rg|te)|vk(40|5[0-3]|\-v)|vm40|voda|vulc|vx(52|53|60|61|70|80|81|83|85|98)|w3c(\-| )|webc|whit|wi(g |nc|nw)|wmlb|wonu|x700|xda(\-|2|g)|yas\-|your|zeto|zte\-"
                b.IgnoreCase=true
                v.IgnoreCase=true
                b.Global=true
                v.Global=true
                if b.test(u) or v.test(Left(u,4)) then
                    tiene_tpv=0
                end if

       if tiene_tpv=1 and (isNull(version_tpv) or (version_tpv>-1 and version_tpv<24)) then
		    if (InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") <> 0) or InStr(Request.ServerVariables("HTTP_USER_AGENT"), "Trident") <> 0then
            'if (InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") <> 0) then
                randomize
	            id_imagen = Int((3 - 1 + 1) * Rnd + 1)
	            if id_imagen=1 then
                    img="img_back1.jpg"
                elseif id_imagen=2 then
                    img="img_back2.jpg"
                elseif id_imagen=3 then
                    img="img_back3.jpg"
                end if%>
            <script type="text/javascript" languaje="javascript">
                //continuamos comprobando si existe el fichero solo si no va a fallar
                    try {
                        //Leemos el fichero cetel.tpv si existe.
                        if (ExisteFichero("C:\\CETEL\\", "cetel.tpv")) {

                            Cadena = ReadTodoFichero("C:\\CETEL\\", "cetel.tpv");
                            lineas = Cadena.split("\r\n");
                            linea1 = lineas[0];
                            linea2 = lineas[1];
                            linea3 = lineas[2];
                        }
                        else {
                            if (ExisteFichero("C:\\", "cetel.tpv")) {
                                Cadena = ReadTodoFichero("C:\\", "cetel.tpv");

                                lineas = Cadena.split("\r\n");
                                linea1 = lineas[0];
                                linea2 = lineas[1];
                                linea3 = lineas[2];
                            }
                            else {
                                linea1 = "";
                                linea2 = "";
                                linea3 = "";
                            }

                        }

                        document.location = "default.asp?empresas=varias&exFich=1&linea1=" + linea1 + "&linea2=" + linea2 + "&linea3=" + linea3;
                    }
                    catch (ex) {

                        fallo= true;
                       
                       //Esta función pone en marcha la descarga del .reg y desconecta de la página 
                       function redir(){
                            if (confirm("<%=LitConfModRegistro %>")) {
                                location = "../controles/ConfigIE.reg";
                                setTimeout("location = 'acceso.asp?mode=fin' ",1000);
                            }
                       }

                        var pagina='';

                        if(navigator.appName.match("Internet Explorer")!=null){

                            pagina = '<div id="line_menu_ppal"></div>' +
                                    '<div id="logos_body_error"></div>' +
                                    '<div id="icon_alert"></div>' +
                                    '<div id="welcome">' +
        	                            '<div id="welcome_title"><%=LITBIENVENIDO%></span></div>' +
                                        '<div>' +
            	                            '<%=LITNavNoConfigurado %>' +
                                            '<p><%=LitManualConfiguracion%></a></p>' +
                                        '</div>' +
                                    '</div>' +
                                    '<div id="expl">' +
        	                            '<div id="expl_title"><%=LitActCmpActX%></div>' +
                                            '<div class="steps">' +
            	                            '<table><tr><td><div id="step1"></div></td>' +
            	                            '<td><div id="step2"></div></td>' +
            	                            '<td><div id="step3"></div></td></tr></table>' +
                                        '</div>' +
                                    '</div>' +
                                    '<div id="footer_wrapper">' +
                                        '<img src="/lib/estilos/ilion/images/<%=Img%>" />' +
                                        '<div id="footer">' +
                                    '</div>';
                        }
                        else
                        {
                            pagina = '<div id="line_menu_ppal"></div>' +
                                    '<div id="logos_body_error"></div>' +
                                    '<div id="icon_alert"></div>' +
                                    '<div id="welcome">' +
        	                            '<div id="welcome_title"><%=LITBIENVENIDO%></span></div>' +
                                        '<div>' +
                                            '<p><%=LITNavNoIE %></a></p>' +
                                        '</div>' +
                                    '</div>' +
                                    '<div id="footer_wrapper">' +
                                        '<img src="/lib/estilos/ilion/images/<%=Img%>" />' +
                                        '<div id="footer">' +
                                    '</div>';
                        }
                }
			</script>
		    <%else%>
		        <script type="text/javascript" language="javascript">document.location = "default.asp?empresas=varias";</script>
		    <%end if
		else%>
			<script type="text/javascript" language="javascript">document.location = "default.asp?empresas=varias";</script>
		<%end if
	end if
end if
set rs_stock = Nothing
set rstAux = Nothing
set rs_empresas = Nothing
set rs_check = Nothing
set rs_clientes = Nothing
set rs_clientes2 = Nothing
set rs_tpv = Nothing
'No Hay Sesion
end if%>
<script type="text/javascript" language="javascript">
    //si falla reemplazamos el body
    if(fallo)
    {
        document.body.innerHTML=pagina;
    }
</script>
</body>
</html>
