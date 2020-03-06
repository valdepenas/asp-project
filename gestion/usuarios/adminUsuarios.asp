<%@ Language=VBScript %>
<%
'' VGR	12/05/03	:	Añadir la condicion para mostrar las opciones de menu que no sean NULL(pag.Inicio).

' >>> MCA 21/12/04 : Incorporar gestión de franjas horarias de acceso a la administración de usuarios
%>
  <% dim enc
     set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
  %>
<!DOCTYPE html PUBLIC "-//W3C/DTD/ XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml1-transitional.dtd" />
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../../styles/ilionp.css.inc" -->
<!--#include file="../../styles/ExtraLink.css.inc" -->
<!--#include file="../../modulos.inc" -->
<!--#include file="adminUsuarios.inc" -->
<!--#include file="../../styles/listTable.css.inc" -->
<!--#include file="../../styles/formularios.css.inc" -->

<meta http-equiv="Content-Type" content="text/html"; charset="iso-8859-1"/>
<meta http-equiv="Content-style-Type" content="text/css"/>
<link rel="stylesheet" href="../../pantalla.css" media="screen"/>
<link rel="stylesheet" href="../../impresora.css" media="print"/>
<script type="text/javascript" language="javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

</head>
<%'Leer parámetros de la página'

  mode         = request.Querystring("mode")
  if mode="search" then
    mode="edit"
  end if
  campo        = limpiaCadena(request.QueryString("campo"))
  criterio     = limpiaCadena(request.QueryString("criterio"))
  texto        = limpiaCadena(request.QueryString("texto"))
	cliente=limpiaCadena(request.querystring("ncliente"))
  	if mode="adminedit" then
  	  mode="edit"
  	  permiso="admin"
  	end if
  	if mode="adminsave" then
  		mode="save"
  		permiso="admin"
  	end if
	if cliente="" then
  		cliente = session("ncliente")
  	end if

  nModulos	= limpiaCadena(Request.QueryString("nmodulos"))
  nUsuarios = limpiaCadena(Request.QueryString("nusuarios"))
  nUsuario1=limpiaCadena(Request.QueryString("hnusuario1"))
  nUsuario2=limpiaCadena(Request.Querystring("hnusuario2"))

	OMC=limpiaCadena(Request.QueryString("OMC"))%>

<body class="BODY_ASP">
<script type="text/javascript" language="javascript">
<% ' >>> MCA 21/12/04 : Función modificada para la gestión de franjas horarias %>
function Mas(sentido,lote,campo,criterio,texto,modo)
{
	if (modo=="gestionhorarios")
	{
		if (document.adminUsuarios.hmodificado.value=="1")
		{
			if (window.confirm("<%=LitDeseaGuardar%>")==true){
				document.adminUsuarios.action="adminUsuarios.asp?mode=guardahorarios&hnusuario1=" + document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + document.adminUsuarios.hnusuario2.value + "&ncliente=" + document.adminUsuarios.hncliente.value + "&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&nusuarios=" + document.adminUsuarios.hnusuarios.value + "&nmodulos=" + document.adminUsuarios.hnmodulos.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
				document.adminUsuarios.submit();
			}
			else document.location="adminUsuarios.asp?mode=gestionhorarios&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&ncliente=" + document.adminUsuarios.hncliente.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
		}
		else document.location="adminUsuarios.asp?mode=gestionhorarios&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&ncliente=" + document.adminUsuarios.hncliente.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
	}
	else
	{
		if (document.adminUsuarios.hmodificado.value=="1")
		{
			if (window.confirm("<%=LitDeseaGuardar%>")==true){
				document.adminUsuarios.action="adminUsuarios.asp?hnusuario1=" + document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + document.adminUsuarios.hnusuario2.value + "&ncliente=" + document.adminUsuarios.hncliente.value + "&mode=save&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&nusuarios=" + document.adminUsuarios.hnusuarios.value + "&nmodulos=" + document.adminUsuarios.hnmodulos.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
				document.adminUsuarios.submit();
			}
			else document.location="adminUsuarios.asp?mode=edit&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&ncliente=" + document.adminUsuarios.hncliente.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
		}
		else document.location="adminUsuarios.asp?mode=edit&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&ncliente=" + document.adminUsuarios.hncliente.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
	}
}
<% ' >>> MCA 21/12/04 : Función modificada para la gestión de franjas horarias %>

function AbreGestionHorarios(lote,campo,criterio,texto,modo)
{
	if (document.adminUsuarios.hmodificado.value=="1")
	{
		if (window.confirm("<%=LitDeseaGuardar%>")==true)
		{
			document.adminUsuarios.action="adminUsuarios.asp?mode=save&hnusuario1=" + document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + document.adminUsuarios.hnusuario2.value + "&ncliente=" + document.adminUsuarios.hncliente.value + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&nusuarios=" + document.adminUsuarios.hnusuarios.value + "&nmodulos=" + document.adminUsuarios.hnmodulos.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
			document.adminUsuarios.submit();
		}
	}

	document.location= "adminUsuarios.asp?mode=gestionhorarios&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&ncliente=" + document.adminUsuarios.hncliente.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
	parent.botones.document.location=  "adminUsuarios_bt.asp?mode=edit&OMC=<%= enc.EncodeForjavascript(OMC)%>&hmode=gestionhorarios";
}

function AbreGestionModulos(lote,campo,criterio,texto,modo)
{
	if (document.adminUsuarios.hmodificado.value=="1")
	{
		if (window.confirm("<%=LitDeseaGuardar%>")==true)
		{
			document.adminUsuarios.action="adminUsuarios.asp?mode=guardahorarios&hnusuario1=" + document.adminUsuarios.hnusuario1.value + "&hnusuario2=" + document.adminUsuarios.hnusuario2.value + "&ncliente=" + document.adminUsuarios.hncliente.value + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&nusuarios=" + document.adminUsuarios.hnusuarios.value + "&nmodulos=" + document.adminUsuarios.hnmodulos.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
			document.adminUsuarios.submit();
		}
	}

	document.location= "adminUsuarios.asp?mode=edit&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto + "&ncliente=" + document.adminUsuarios.hncliente.value + "&OMC=" + document.adminUsuarios.hOMC.value + "&ver=" + document.adminUsuarios.ver.value;
	parent.botones.document.location=  "adminUsuarios_bt.asp?mode=edit&OMC=<%= enc.EncodeForjavascript(OMC) %>&hmode=edit";
}

function AbreteSesamo(usr,cliente,ver)
{
    //ricardo 4-3-2010 se cambia el "-" por "$" ya que con los usuarios de shades hay problemas
    AbrirVentana("../../central.asp?pag1=Gestion/usuarios/pagusuario.asp&modd=1&ndoc=" + cliente + "$" + usr + "&mode=edit&pag2=gestion/usuarios/PagUsuario_bt.asp&titulo=<%=LitPagUsuario%>&ndocumento=" + ver,"P","<%=altoventana%>","<%=anchoventana%>");
}

function AbreteSesamo2(cliente) {
    AbrirVentana("../../central.asp?pag1=Gestion/usuarios/ConsultaModulos.asp&modd=1&ndoc=" + cliente + "&mode=edit&pag2=gestion/usuarios/ConsultaModulos_bt.asp&titulo=<%=LitAyudaInfoModDisp%>","P","<%=altoventana%>","<%=anchoventana%>");
}

function ActualizaLicencias(numUser,usuario,numModulos) {
	nUsuarios=document.adminUsuarios.hnusuarios.value;
	nModulos=document.adminUsuarios.hnmodulos.value;

	moduloCheck="check"+usuario+"i"+numModulos;
	moduloLicencias="licencia"+numModulos;

	document.adminUsuarios.hmodificado.value="1";

	if (document.adminUsuarios.elements[moduloCheck].checked==1)
	{
		valor=parseInt(document.getElementById(moduloLicencias).innerHTML)-1;
  		if (valor==0){
			for (i=numUser;i<=nUsuarios-1;i++) {
				nCheck="check"+i+"i"+numModulos;
				if (document.adminUsuarios.elements[nCheck].checked!=1) document.adminUsuarios.elements[nCheck].disabled=true;
			}
  		}
  		document.adminUsuarios.elements[moduloCheck].value="yyy";
	  	document.getElementById(moduloLicencias).innerHTML=valor;
	}
	else
	{
	    valor=parseInt(document.getElementById(moduloLicencias).innerHTML)+1;
	    if (valor==1){
			for (i=numUser;i<=nUsuarios-1;i++) {
				nCheck="check"+i+"i"+numModulos;
				if ((document.adminUsuarios.elements[nCheck].disabled==true)&&(document.adminUsuarios.elements[nCheck].value!="BAJA"))
					document.adminUsuarios.elements[nCheck].disabled=false;
			}
	    }
	    document.adminUsuarios.elements[moduloCheck].value="";
	    document.getElementById(moduloLicencias).innerHTML=valor;
	}
}

function OpenSpecialParameters(){
    document.adminUsuarios.action="SpecialParameters.asp?mode=browse";
    document.adminUsuarios.submit(); 
    parent.botones.document.location="SpecialParameters_bt.asp?"
}
</script>
<%
'********** FUNCIONES'
'****************************************************************************************************************
'Botones de navegación para las búsquedas.
sub NextPrev(lote,lotes,campo,criterio,texto,pos,mode)%>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
	<tr><td class="MAS">
        <%lote=cint(lote)
	    lotes=cint(lotes)
	    varias=false
		if lote>1 then%>
			<a class="CELDAREF" href="javascript:Mas('prev',<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(mode)%>');">
			<img id="BtnAnt" src="<%=ImgAnteriorLF%>" <%=ParamImgAnteriorLF%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a>
            <%varias=true
		end if
		textopag=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)%>
		<font class="CELDA"> <%=textopag%> </font>
        <%if lote<lotes then%>
			<a class="CELDAREF" href="javascript:Mas('next',<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(mode)%>');">
			<img id="BtnSig" src="<%=ImgSiguienteLF%>" <%=ParamImgSiguienteLF%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a><%
			varias=true
		end if%>
	</td></tr>
</table>
<%end sub

'****************************************************************************************************************
'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas
'de datos.
	'campo: Nombre del campo con el cual se realizará la búsqueda
	'criterio: Tipo de búsqueda
	'texto: Texto a buscar.
function CadenaBusqueda(campo,criterio,texto)
   if texto > "" then
	  select case criterio
		  case "contiene"
			  CadenaBusqueda=" and " + campo + " like '%" + texto + "%'"
  		  case "empieza"
			  CadenaBusqueda=" and " + campo + " like '" + texto + "%'"
		  case "termina"
			  CadenaBusqueda=" and " + campo + " like '%" + texto + "'"
		  case "igual"
			  CadenaBusqueda=" and " + campo + "='" + texto + "'"
	  end select
	  ''07/05/2009 MPC Se cambia la ordenación del listado de los cliente
	  CadenaBusqueda=CadenaBusqueda & " order by c.fbaja, i.nombre"
   else
      CadenaBusqueda=" order by c.fbaja, i.nombre"
   end if
      ''FIN MPC
end function

' >>> MCA 21/12/04 : Incorporar gestión de franjas horarias a la administración de usuarios

sub GuardarHorarios_masivo(nusuario1,nusuario2)
	strwhere= CadenaBusqueda(campo,criterio,texto)
    rstAux2.cursorlocation=3
	rstAux2.open "SELECT * FROM CLIENTES_USERS as c with(nolock), indice as i with(nolock) WHERE usuario=entrada and administrar=1 and ncliente = '" & cliente & "' " & strwhere,DSNIlion
	for index=0 to nusuario1-1
		rstAux2.movenext
	next

	for usuario=nusuario1 to nusuario2
        ListaHorarios=""
        rst.cursorlocation=2
		rst.Open "select * from indice with(updlock) where entrada='" & rstAux2("usuario") & "'",DSNIlion,adOpenKeyset,adLockOptimistic
		if not rst.eof then
			if request.form("desdehora1_u"& usuario)>"" then
				rst("accesodesdehora1")= TimeValue(Nulear(request.form("desdehora1_u"& usuario)))
                ListaHorarios=ListaHorarios & "accesodesdehora1=" & rst("accesodesdehora1")
			else
				rst("accesodesdehora1")= null
                ListaHorarios=ListaHorarios & "accesodesdehora1=null"
			end if
			if request.form("hastahora1_u"& usuario)>"" then
				rst("accesohastahora1")= TimeValue(Nulear(request.form("hastahora1_u"& usuario)))
                ListaHorarios=ListaHorarios & ",accesohastahora1=" & rst("accesohastahora1")
			else
				rst("accesohastahora1")= null
                ListaHorarios=ListaHorarios & ",accesohastahora1=null"
			end if
			if request.form("desdehora2_u"& usuario)>"" then
				rst("accesodesdehora2")= TimeValue(Nulear(request.form("desdehora2_u"& usuario)))
                ListaHorarios=ListaHorarios & ",accesodesdehora2=" & rst("accesodesdehora2")
			else
				rst("accesodesdehora2")= null
                ListaHorarios=ListaHorarios & ",accesodesdehora2=null"
			end if
			if request.form("hastahora2_u"& usuario)>"" then
				rst("accesohastahora2")= TimeValue(Nulear(request.form("hastahora2_u"& usuario)))
                ListaHorarios=ListaHorarios & ",accesohastahora2=" & rst("accesohastahora2")
			else
				rst("accesohastahora2")= null
                ListaHorarios=ListaHorarios & ",accesohastahora2=null"
			end if
			rst.update
			rst.close
            ''Ricardo 29-01-2014 se audita las acciones de modificacion de horarios a un usuario
            if ListaHorarios & "">"" then
                set connIR = Server.CreateObject("ADODB.Connection")
                set comdIR =  Server.CreateObject("ADODB.Command")
                connIR.Open DSNIlion
                connIR.cursorlocation=3
                comdIR.ActiveConnection=connIR
                comdIR.CommandTimeout = 120
                comdIR.CommandText = "insert into trazas_servicios(fecha,servicio,evento,texto) select getdate(),'GestionUsuarios','ModifHorarios','Gestor:' + ? + ',Usuario:' + ? + ',Empresa:' + ? + ',Horario:' + ?"
                comdIR.CommandType = adCmdText
                comdIR.Parameters.Append comdIR.CreateParameter("@entrada",adVarChar,,50,session("usuario")&"")
                comdIR.Parameters.Append comdIR.CreateParameter("@usuario",adVarChar,,50,rstAux2("usuario")&"")
                comdIR.Parameters.Append comdIR.CreateParameter("@ncliente",adVarChar,,5,cliente&"")
                comdIR.Parameters.Append comdIR.CreateParameter("@ListaHorarios",adVarChar,,8000,ListaHorarios&"")
                comdIR.Execute
                connIR.close
                set comdIR=nothing
                set connIR=nothing
            end if
            ''fin Ricardo 29-01-2014
		end if
		rstAux2.movenext
	next
	rstAux2.close
end sub

' >>> MCA 21/12/04 : Incorporar gestión de franjas horarias a la administración de usuarios

'***************************************************************************************************************'
'*********************************  CODIGO PRINCIPAL DE LA PÁGINA  *********************************************'
'***************************************************************************************************************'
const borde=0

    ver = limpiaCadena(request.QueryString("ndoc"))
    if ver & "" = "" then ver = limpiaCadena(request.QueryString("ver"))
    if ver & "" = "" then ver = request.form("ver")

    verEncHTMLJS =enc.EncodeForJavascript(enc.EncodeForHtmlAttribute(ver))

    viene = limpiaCadena(request.QueryString("viene"))
    'Conexion y cursores'
    set conn = Server.CreateObject("ADODB.Connection")
    set command=  Server.CreateObject("ADODB.Command")
    set rst = Server.CreateObject("ADODB.Recordset")
    set rstAux = Server.CreateObject("ADODB.Recordset")
    set rstAux2 = Server.CreateObject("ADODB.Recordset")
    set rstAux3 = Server.CreateObject("ADODB.Recordset")
    set rstAux4 = Server.CreateObject("ADODB.Recordset")
    set rstAux5 = Server.CreateObject("ADODB.Recordset")
    set rstAux6 = Server.CreateObject("ADODB.Recordset")
    set rstAux7 = Server.CreateObject("ADODB.Recordset")
    set rstTemp = Server.CreateObject("ADODB.Recordset")
    set rstSelect = Server.CreateObject("ADODB.Recordset")
    set rstConn = Server.CreateObject("ADODB.Recordset")
    set rsIndice = Server.CreateObject("ADODB.Recordset")
    set rstFirma = Server.CreateObject("ADODB.Recordset")%>
<form name="adminUsuarios" method="post">
    <%
    ''response.write("el viene es-" & viene & "-<br/>")
    if viene <> "sel_app" then
        'sistema de gestion
        if session("ncliente")&""="00000" then
            PaintHeaderPopUp "ges_clientes.asp", LITADMUSUCLI
        else
            PintarCabecera "adminUsuarios.asp"
        end if
    end if
    %>
<input type="hidden" name="hcampo" value="<%=enc.EncodeForHtmlAttribute(campo)%>"/>
<input type="hidden" name="hcriterio" value="<%=enc.EncodeForHtmlAttribute(criterio)%>"/>
<input type="hidden" name="htexto" value="<%=enc.EncodeForHtmlAttribute(texto)%>"/>
<input type="hidden" name="hncliente" value="<%=enc.EncodeForHtmlAttribute(cliente)%>"/>
<input type="hidden" name="hmodificado" value="0"/>
<input type="hidden" name="hpermiso" value="<%=enc.EncodeForHtmlAttribute(permiso)%>"/>
<input type="hidden" name="hOMC" value="<%=enc.EncodeForHtmlAttribute(OMC)%>"/>
<input type="hidden" name="recargar" value="SI"/>
<input type="hidden" name="hmode" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
<input type="hidden" name="ver" value="<%=verEncHTMLJS%>"/>
<%WaitBoxOculto LitEsperePorFavor

 'Acción a realizar'

  '***** MODO SAVE ***************************'
  if mode="save" then
	nUser1=cint(nUsuario1)
	nUser2=cint(nUsuario2)
	 
  	'Guardamos los módulos para los usuarios'
	strwhere= CadenaBusqueda(campo,criterio,texto)
    rstAux2.cursorlocation=3
	rstAux2.open "SELECT * FROM CLIENTES_USERS as c with(NOLOCK), indice as i with(NOLOCK) WHERE usuario=entrada and administrar=1 and ncliente = '" & cliente & "' " & strwhere,DSNIlion
	for index=0 to nUser1-1
		rstAux2.movenext
	next

'***********'
	' Guardamos los cambios en las páginas'
'***********'
	for usuario=nUser1 to nUser2
		listaAlta=""
		listaBaja=""
        set connRSave = Server.CreateObject("ADODB.Connection")
        set comdRSave =  Server.CreateObject("ADODB.Command")
        connRSave.Open DSNIlion
        connRSave.cursorlocation=3
        comdRSave.ActiveConnection=connRSave
        comdRSave.CommandTimeout = 120
        StrObtLic="SELECT l.nmodulo, nombre FROM LICENCIAS AS l with(nolock) LEFT OUTER JOIN clientes AS c with(nolock) ON c.ncliente = l.ncliente LEFT OUTER JOIN modulos_comerciales AS m with(nolock) ON m.nmodulo = l.nmodulo where c.ncliente='" & cliente & "' and l.visible=1 ORDER BY l.nmodulo"
        comdRSave.CommandText = StrObtLic
        comdRSave.CommandType = adCmdText
        ''comdRSave.Parameters.Append comdRSave.CreateParameter("@ncliente",adVarChar,adParamInput,5,cliente&"")
        set rstAux=comdRSave.Execute

		for modulo=0 to nModulos-1
			checkmodulo="check"&usuario&"i"&modulo
            rstAux3.cursorlocation=3
			rstAux3.open "SELECT * from modulosc_users with(nolock) where usuario='" & rstAux2("usuario") & "' and ncliente='" & cliente & "' and nmodulo='" & rstAux("nmodulo") & "'",DSNIlion

			if request.form(checkmodulo)>"" then
			  if rstAux3.eof then
             
'************ INSERCIÓN DE UN MÓDULO COMERCIAL ENTERO **********'
                    ''ahora se da de alta en todos los modulos
 	                Resultado=0
	                set commandAux =  Server.CreateObject("ADODB.Command")
	                conn.open dsnilion
	                commandAux.ActiveConnection =conn
	                commandAux.CommandTimeout = 0
	                commandAux.CommandText="AltaUnUsurioEnUnModuloComercial"
	                commandAux.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	                commandAux.Parameters.Append commandAux.CreateParameter("nempresa",adVarChar,adParamInput,5,cliente)
	                commandAux.Parameters.Append commandAux.CreateParameter("@usuario",adVarChar,adParamInput,50,rstAux2("usuario"))
	                commandAux.Parameters.Append commandAux.CreateParameter("@nmodulo",adVarChar,adParamInput,5,rstAux("nmodulo"))
	                commandAux.Parameters.Append commandAux.CreateParameter("@dsn_cronos",adVarChar,adParamInput,500,DSNCronos & "")
	                commandAux.Parameters.Append commandAux.CreateParameter("@resultado", adInteger, adParamOutput, Resultado)
                    commandAux.Parameters.Append commandAux.CreateParameter("@admin",adVarChar,adParamInput,5,"0")
                    commandAux.Parameters.Append commandAux.CreateParameter("@activelicensecode",adInteger,adParamInput,,0)
                    commandAux.Parameters.Append commandAux.CreateParameter("@usuario2", adVarChar, adParamInput,50, session("usuario"))
                    commandAux.Parameters.Append commandAux.CreateParameter("@ip", adVarChar, adParamInput,75, Request.ServerVariables(CLIENT_IP))
                    commandAux.Parameters.Append commandAux.CreateParameter("@host", adVarChar, adParamInput,75, Request.ServerVariables("REMOTE_HOST"))
                    commandAux.Execute,,adExecuteNoRecords

	                Resultado = commandAux.Parameters("@resultado").Value
	                conn.close
	                set commandAux=nothing

                    strSelectModuloAlta = "select nombre from modulos_comerciales with(nolock) where nmodulo = ?"
                    moduloAlta = DLookupP1(strSelectModuloAlta,rstAux("nmodulo")&"",adVarchar,2,dsnilion&"")                   
			  	    listaAlta=listaAlta&"Módulo "&moduloAlta&","

                    'listaAlta=listaAlta&"Módulo "&d_lookup("nombre","modulos_comerciales","nmodulo='"+rstAux("nmodulo")+"'",dsnilion)&","
			  end if
			else
			  if not rstAux3.eof then
'************ BORRADO DE UN MÓDULO COMERCIAL ENTERO **********'
                    Resultado=0
                    set commandAux =  Server.CreateObject("ADODB.Command")
                    conn.open dsnilion
                    commandAux.ActiveConnection =conn
                    commandAux.CommandTimeout = 0
                    commandAux.CommandText="BajaUnUsurioEnUnModuloComercial"
                    commandAux.CommandType = adCmdStoredProc 'Procedimiento Almacenado
                    commandAux.Parameters.Append commandAux.CreateParameter("nempresa",adVarChar,adParamInput,5,cliente)
                    commandAux.Parameters.Append commandAux.CreateParameter("@usuario",adVarChar,adParamInput,50,rstAux2("usuario"))
                    commandAux.Parameters.Append commandAux.CreateParameter("@nmodulo",adVarChar,adParamInput,5,rstAux("nmodulo"))
                    commandAux.Parameters.Append commandAux.CreateParameter("@resultado", adInteger, adParamOutput, Resultado)
                    commandAux.Parameters.Append commandAux.CreateParameter("@activelicensecode",adInteger,adParamInput,,0)
                    commandAux.Parameters.Append commandAux.CreateParameter("@usuario2", adVarChar, adParamInput,50, session("usuario"))
                    commandAux.Parameters.Append commandAux.CreateParameter("@ip", adVarChar, adParamInput,75, Request.ServerVariables(CLIENT_IP))
                    commandAux.Parameters.Append commandAux.CreateParameter("@host", adVarChar, adParamInput,75, Request.ServerVariables("REMOTE_HOST"))
                    commandAux.Execute,,adExecuteNoRecords

                    Resultado = commandAux.Parameters("@resultado").Value
                    conn.close
                    set commandAux=nothing

'************ FIN BORRADO DE UN MÓDULO COMERCIAL ENTERO **********'
                strSelectModuloBaja = "select nombre from modulos_comerciales with(nolock) where nmodulo = ?"
                moduloBaja = DLookupP1(strSelectModuloBaja,rstAux("nmodulo")&"",adVarchar,2,dsnilion&"")                
			  	listaBaja=listaBaja&"Módulo "&moduloBaja&","
                'listaBaja=listaBaja&"Módulo "&d_lookup("nombre","modulos_comerciales","nmodulo='"+rstAux("nmodulo")+"'",dsnilion)&","
			  end if
			end if
			rstAux3.close
			rstAux.movenext
		next

        ''Ricardo 29-01-2014 se audita las acciones de alta y baja de modulos comerciales a un usuario
        rstAux3.CursorLocation=2
        StrAudit=""
        if listaAlta & "">"" then
            set connIR = Server.CreateObject("ADODB.Connection")
            set comdIR =  Server.CreateObject("ADODB.Command")
            connIR.Open DSNIlion
            connIR.cursorlocation=3
            comdIR.ActiveConnection=connIR
            comdIR.CommandTimeout = 120
            comdIR.CommandText = "insert into trazas_servicios(fecha,servicio,evento,texto) select getdate(),'GestionUsuarios','AltaModComercial','Gestor:' + ? + ',Usuario:' + ? + ',Empresa:' + ? + ',Modulos:' + ?"
            comdIR.CommandType = adCmdText
            comdIR.Parameters.Append comdIR.CreateParameter("@entrada",adVarChar,,50,session("usuario")&"")
            comdIR.Parameters.Append comdIR.CreateParameter("@usuario",adVarChar,,50,rstAux2("usuario")&"")
            comdIR.Parameters.Append comdIR.CreateParameter("@ncliente",adVarChar,,5,cliente&"")
            comdIR.Parameters.Append comdIR.CreateParameter("@listaAlta",adVarChar,,8000,listaAlta&"")
            comdIR.Execute
            connIR.close
            set comdIR=nothing
            set connIR=nothing
        end if
        if listaBaja & "">"" then
            set connIR = Server.CreateObject("ADODB.Connection")
            set comdIR =  Server.CreateObject("ADODB.Command")
            connIR.Open DSNIlion
            connIR.cursorlocation=3
            comdIR.ActiveConnection=connIR
            comdIR.CommandTimeout = 120
            comdIR.CommandText = "insert into trazas_servicios(fecha,servicio,evento,texto) select getdate(),'GestionUsuarios','BajaModComercial','Gestor:' + ? + ',Usuario:' + ? + ',Empresa:' + ? + ',Modulos:' + ?"
            comdIR.CommandType = adCmdText
            comdIR.Parameters.Append comdIR.CreateParameter("@entrada",adVarChar,,50,session("usuario")&"")
            comdIR.Parameters.Append comdIR.CreateParameter("@usuario",adVarChar,,50,rstAux2("usuario")&"")
            comdIR.Parameters.Append comdIR.CreateParameter("@ncliente",adVarChar,,5,cliente&"")
            comdIR.Parameters.Append comdIR.CreateParameter("@listaBaja",adVarChar,,8000,listaBaja&"")
            comdIR.Execute
            connIR.close
            set comdIR=nothing
            set connIR=nothing
        end if
        ''fin Ricardo 29-01-2014

        listaRestricciones=""
		''si el usuario ya no tiene ningun modulo asignado , se le quitaran todas las restricciones de todos los modulos
		'', ya que para ello se le han quitado los modulos
		rstAux3.CursorLocation=3
		rstAux3.open "SELECT * from modulosc_users with(nolock) where usuario='" & rstAux2("usuario") & "' and ncliente='" & cliente & "' ",DSNIlion
		if rstAux3.EOF then
		    rstAux3.close
            ''Ricardo 29-01-2014 se audita las acciones eliminacion de restricciones a un usuario
            set connR = Server.CreateObject("ADODB.Connection")
            set comdR =  Server.CreateObject("ADODB.Command")
            connR.Open DSNIlion
            connR.cursorlocation=3
            comdR.ActiveConnection=connR
            comdR.CommandTimeout = 120
            StrAudit=""
            StrAudit=StrAudit & " declare @lista varchar(MAX) "
            StrAudit=StrAudit & " set @lista='' "
            StrAudit=StrAudit & " select @lista=@lista + r.item + ',' from restricciones as r with(NOLOCK) "
            StrAudit=StrAudit & " where r.entrada='?'  and r.ncliente='?' "
            StrAudit=StrAudit & " and r.item not in (select par.item from parametrizaciones as par with(NOLOCK) where par.ncliente=r.ncliente) "
            StrAudit=StrAudit & " if @lista<>'' "
            StrAudit=StrAudit & " 	BEGIN "
            StrAudit=StrAudit & " 		set @lista=SUBSTRING(@lista,1,len(@lista)-1) "
            StrAudit=StrAudit & " 	END "
            StrAudit=StrAudit & " select @lista as lista"
            comdR.CommandText = StrAudit
            comdR.CommandType = adCmdText
            comdR.Parameters.Append comdR.CreateParameter("@entrada",adVarChar,,50,rstAux2("usuario")&"")
            comdR.Parameters.Append comdR.CreateParameter("@ncliente",adVarChar,,5,cliente&"")
            set rstAux3=comdR.Execute
            if not rstAux3.EOF then
                rstInsert=1
                listaRestricciones=rstAux3("lista") & ""
                if listaRestricciones & ""="" then
                    listaRestricciones="no hay restricciones a eliminar"
                end if
            else
                listaRestricciones="no hay restricciones a eliminar"
            end if
            if rstAux3.state<>0 then rstAux3.close
            set comdR=nothing
            set connR=nothing
            ''fin Ricardo 29-01-2014

            ''ricardo 01-03-2013 se eliminaran las restricciones menos las parametrizaciones de la empresa
            StrDelRes="delete r from restricciones as r with(ROWLOCK) where r.entrada='" & rstAux2("usuario") & "'  and r.ncliente='" & cliente & "' and r.item not in (select par.item from parametrizaciones as par with(NOLOCK) where par.ncliente=r.ncliente)"
            rstAux3.cursorlocation=2
            rstAux3.open StrDelRes,DSNIlion,adOpenKeyset,adLockOptimistic
            if rstAux3.State<>0 then rstAux3.close
            ''Ricardo 29-01-2014 se audita las acciones eliminacion de restricciones a un usuario
            if listaRestricciones & "">"" then
                set connIR = Server.CreateObject("ADODB.Connection")
                set comdIR =  Server.CreateObject("ADODB.Command")
                connIR.Open DSNIlion
                connIR.cursorlocation=3
                comdIR.ActiveConnection=connIR
                comdIR.CommandTimeout = 120
                comdIR.CommandText = "insert into trazas_servicios(fecha,servicio,evento,texto) select getdate(),'GestionUsuarios','DelAllRestricciones','Gestor:' + ? + ',Usuario:' + ? + ',Empresa:' + ? + ',listaRestricciones:' + ?"
                comdIR.CommandType = adCmdText
                comdIR.Parameters.Append comdIR.CreateParameter("@entrada",adVarChar,,50,session("usuario")&"")
                comdIR.Parameters.Append comdIR.CreateParameter("@usuario",adVarChar,,50,rstAux2("usuario")&"")
                comdIR.Parameters.Append comdIR.CreateParameter("@ncliente",adVarChar,,5,cliente&"")
                comdIR.Parameters.Append comdIR.CreateParameter("@listaRestricciones",adVarChar,,8000,listaRestricciones&"")
                comdIR.Execute
                connIR.close
                set comdIR=nothing
                set connIR=nothing
            end if
            ''fin Ricardo 29-01-2014
        else
            rstAux3.close
		end if
		if rstAux.state<>0 then rstAux.close
        set comdRSave=nothing
        set connRSave=nothing
		rstAux2.movenext
	next
	rstAux2.close
    %>
	<script type="text/javascript" language="javascript">alert("<%=LitDatosOK%>");</script>
	<%if Request.Form("hmode")="gestionhorarios" then
		mode= "gestionhorarios"
	else
		mode="edit"
	end if
  end if
  on error resume next
  '***** FIN modo save ***********************'%>
   <table border="0" width="100%" cellpadding="2" cellspacing="0"  style=" border:0px solid black;border-collapse:collapse;" >
   <tr>
	  <td width="40%" class="CELDA7" align="left">
        <% 
            strSelectRSocial = "select rsocial from clientes with(nolock) where ncliente = ?"
            RSocial = DLookupP1(strSelectRSocial,session("ncliente")&"",adVarchar,5,DSNIlion&"") 
        %>
        <font class="CELDA7"><%=LitEmpresa%> </font><font class="CELDAB7"><%=enc.EncodeForHtmlAttribute(null_s(RSocial))%></font>
	  </td>
      <td style="width:60%; text-align:right;">
		<table style="width:100%; text-align:right;" id="enlaces_extra"  cellpadding="0" cellspacing="0">
  	        <tr>

<%
if err.number<>0 then
    response.end
end if
on error goto 0
    '***** MODO EDIT ***********************'
  if mode="edit" then
		sentido=limpiaCadena(Request.QueryString("sentido"))
		lote=limpiaCadena(Request.QueryString("lote"))
		if lote="" then lote=1

		if sentido="next" then
			lote=lote+1
		elseif sentido="prev" then
			lote=lote-1
		end if

		strwhere= CadenaBusqueda(campo,criterio,texto)
        rstTemp.cursorlocation=3
		rstTemp.open "SELECT c.usuario,i.nombre,c.fbaja FROM CLIENTES_USERS as c with(nolock), indice as i with(nolock) WHERE c.usuario=i.entrada and c.administrar=1 and c.ncliente = '" & cliente & "' " & strwhere,DSNIlion
		lotes=rstTemp.RecordCount/NumReg
		if lotes>clng(lotes) then
			lotes=clng(lotes)+1
		else
			lotes=clng(lotes)
		end if

		nModulos=0
		nUsuarios=((lote-1)*NumReg)

    'if permiso<>"admin" then 
    if session("ncliente") <> "00000" then
        %>
        <!--<td align="center" style="border : 1px solid black; width:150px;"  class="CELDARIGHTB" onmouseover="this.className='TDACTIVO8'" onmouseout="this.className='CELDARIGHTB'"  bgcolor="<%=color_blau%>" >-->
                <td style="width:150px"><div>
                    <a class="CELDAREF" align="center" href="javascript:OpenSpecialParameters();"><%=LitSpecialParam%></a>
                </div>
                </td>
                <td class="der" style="width:3px">|</td>
        <!--</td>--><%
    end if

	if permiso="admin" then %>
      <!--<td align="right">
		&nbsp;-->
<%	elseif mode="edit" or mode="save" then 
        %><!--<td align="center" style="border : 1px solid black;width:200px;" class="CELDARIGHTB" onmouseover="this.className='TDACTIVO8'" onmouseout="this.className='CELDARIGHTB'"  bgcolor="<%=color_blau%>">--><%
        if ver <> "1" then%>
            <td style="width:230px"><div>
		        <a class="CELDAREF" href="javascript:AbreGestionHorarios(<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(mode)%>');"><%=LitGestionHoraria%></a>
            </div></td>
		<%end if%>
<% 	elseif mode="gestionhorarios" or mode="guardahorarios"  then
        %><!--<td align="center" style="border : 1px solid black;width:200px;" class="CELDARIGHTB" onmouseover="this.className='TDACTIVO8'" onmouseout="this.className='CELDARIGHTB'"  bgcolor="<%=color_blau%>">--><%
        if ver <> "1" then%>
            <td style="width:230px"><div>
    		    <a class="CELDAREF" href="javascript:AbreGestionModulos(<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(mode)%>');"><%=LitGestionModulos%></a>
            </div></td>
        <%end if
	end if %>
    </tr>
    </table>
	</td>
  </tr>
  </table>
<br/>
		<input type="hidden" name="hlote" value="<%=enc.EncodeForHtmlAttribute(lote)%>"/>
		<input type="hidden" name="sinUsuarios" value=""/>
        <%if not rstTemp.eof then
			rstTemp.PageSize=NumReg
			rstTemp.AbsolutePage=lote

		NextPrev lote,lotes,campo,criterio,texto,1,mode%>
		<table class="width100 xs-table-responsive">
			<tr class="underOrange"><%
				if OMC<>"SI" then
					%><td class="CELDAB7"><%=LitUsuario%></td><%
                    set connRSave1 = Server.CreateObject("ADODB.Connection")
                    set comdRSave1 =  Server.CreateObject("ADODB.Command")
                    set rstAux1 = Server.CreateObject("ADODB.Recordset")
                    connRSave1.Open DSNIlion
                    connRSave1.cursorlocation=3
                    comdRSave1.ActiveConnection=connRSave1
                    comdRSave1.CommandTimeout = 120
                    StrObtLic="SELECT l.nmodulo, nombre FROM LICENCIAS AS l with(nolock) LEFT OUTER JOIN clientes AS c with(nolock) ON c.ncliente = l.ncliente LEFT OUTER JOIN modulos_comerciales AS m with(nolock) ON m.nmodulo = l.nmodulo where c.ncliente='" & cliente & "' and l.visible=1 ORDER BY l.nmodulo"
''response.write("el StrObtLic 1 es-" & StrObtLic & "-" & cliente & "-<br>")
                    comdRSave1.CommandText = StrObtLic
                    comdRSave1.CommandType = adCmdText
                    ''comdRSave1.Parameters.Append comdRSave1.CreateParameter("@ncliente",adVarChar,adParamInput,5,cliente&"")
                    set rstAux1=comdRSave1.Execute
					while not rstAux1.eof
						%><td class="CELDAC7"><b><%=EncodeForHtml(rstAux1("nombre"))%></b></td><%
						nModulos=nModulos+1
						rstAux1.movenext
					wend
					if rstAux1.state<>0 then rstAux1.close
                    ''comdRSave1.close
                    ''connRSave1.close
                    set rstAux1=nothing
                    set comdRSave1=nothing
                    set connRSave1=nothing
				else
					%><td class="CELDAB7"><%=LitUsuarios%></td><%
				end if %>
			</tr>
			<input type="hidden" name="hnmodulos" value="<%=enc.EncodeForHtmlAttribute(nModulos)%>"/><%

			 while not rstTemp.eof and fila<NumReg
				DrawFila color_blau
					nombreUsuario=rstTemp("nombre")
					if rstTemp("fbaja")>"" then
						%><td class="CELDA"><%=enc.EncodeForHtmlAttribute(null_s(nombreUsuario))%> (<%=ucase(LitBaja)%>)</td><%
					else
						%><td><a class="CELDAREF" href="javascript:AbreteSesamo('<%=enc.EncodeForJavascript(rstTemp("usuario"))%>','<%=enc.EncodeForJavascript(cliente)%>', '<%=verEncHTMLJS%>');"><%=enc.EncodeForHtmlAttribute(nombreUsuario)%></a></td><%
					end if

					if OMC<>"SI" then
                        set connRSave2 = Server.CreateObject("ADODB.Connection")
                        set comdRSave2 =  Server.CreateObject("ADODB.Command")
                        connRSave2.Open DSNIlion
                        connRSave2.cursorlocation=3
                        comdRSave2.ActiveConnection=connRSave2
                        comdRSave2.CommandTimeout = 120
                        StrObtLic="SELECT l.nmodulo, nombre FROM LICENCIAS AS l with(nolock) LEFT OUTER JOIN clientes AS c with(nolock) ON c.ncliente = l.ncliente LEFT OUTER JOIN modulos_comerciales AS m with(nolock) ON m.nmodulo = l.nmodulo where c.ncliente='" & cliente & "' and l.visible=1 ORDER BY l.nmodulo"
''response.write("el StrObtLic 2 es-" & StrObtLic & "-" & cliente & "-<br>")
                        comdRSave2.CommandText = StrObtLic
                        comdRSave2.CommandType = adCmdText
                        ''comdRSave2.Parameters.Append comdRSave2.CreateParameter("@ncliente",adVarChar,adParamInput,5,cliente&"")
                        set rstAux=comdRSave2.Execute

						for modulo=0 to nModulos-1
							numUser=((lote-1)*NumReg)
							%><td class="CELDAC7" style="text-align:center;"><input type="checkbox" class="CELDA" name="check<%=enc.EncodeForHtmlAttribute(nUsuarios)%>i<%=modulo%>" value="" onclick="ActualizaLicencias(<%=enc.EncodeForJavascript(numUser)%>,'<%=enc.EncodeForJavascript(nUsuarios)%>','<%=enc.EncodeForJavascript(modulo)%>');"/></td><%
							if rstTemp("fbaja")>"" then
							  %><script type="text/javascript" language="javascript">
							  	document.adminUsuarios.check<%=nUsuarios%>i<%=modulo%>.disabled=true;
							  	document.adminUsuarios.check<%=nUsuarios%>i<%=modulo%>.value="BAJA";
							  </script><%
							end if
                            rstAux2.cursorlocation=3
							rstAux2.open "SELECT ncliente from modulosc_users with(nolock) where ncliente='"&cliente&"' and nmodulo='"&rstAux("nmodulo")&"' and usuario='"&rstTemp("usuario")&"'",DSNIlion
							if not rstAux2.eof then%>
							    <script type="text/javascript" language="javascript">
							        document.adminUsuarios.check<%=nUsuarios%>i<%=modulo%>.checked=1;
							  		document.adminUsuarios.check<%=nUsuarios%>i<%=modulo%>.value="yyy";
							  	</script>
							<%end if
							rstAux2.close
							rstAux.movenext
						next
						if rstAux.state<>0 then rstAux.close
                        ''comdRSave2.close
                        ''connRSave2.close
                        set comdRSave2=nothing
                        set connRSave2=nothing
					end if
				CloseFila
				fila=fila+1
				rstTemp.movenext
				nUsuarios=nUsuarios+1
			 wend%>
		<input type="hidden" name="hnusuario1" value="<%=enc.EncodeForHtmlAttribute(numUser)%>"/>
		<input type="hidden" name="hnusuario2" value="<%=enc.EncodeForHtmlAttribute(nUsuarios)-1%>"/>
		<input type="hidden" name="hnusuarios" value="<%=enc.EncodeForHtmlAttribute(nUsuarios)%>"/>
            <%if OMC<>"SI" then%>
				<tr class="underOrange"><%
					%><td class="CELDAB7"><b><%=LitLicenciasDisp%></b></td><%
                    set connRSave3 = Server.CreateObject("ADODB.Connection")
                    set comdRSave3 =  Server.CreateObject("ADODB.Command")
                    connRSave3.Open DSNIlion
                    connRSave3.cursorlocation=3
                    comdRSave3.ActiveConnection=connRSave3
                    comdRSave3.CommandTimeout = 120
                    StrObtLic="SELECT l.nmodulo, nombre FROM LICENCIAS AS l with(nolock) LEFT OUTER JOIN clientes AS c with(nolock) ON c.ncliente = l.ncliente LEFT OUTER JOIN modulos_comerciales AS m with(nolock) ON m.nmodulo = l.nmodulo where c.ncliente='" & cliente & "' and l.visible=1 ORDER BY l.nmodulo"
''response.write("el StrObtLic 3 es-" & StrObtLic & "-" & cliente & "-<br>")
                    comdRSave3.CommandText = StrObtLic
                    comdRSave3.CommandType = adCmdText
                    ''comdRSave3.Parameters.Append comdRSave3.CreateParameter("@ncliente",adVarChar,adParamInput,5,cliente&"")
                    set rstAux=comdRSave3.Execute

					for modulo=0 to nModulos-1
						' **** Abro la conexión obteniendo las licencias disponibles para cada módulo'
						conn.open DSNIlion
						strselect="EXEC sp_ObtieneLicenciasDispMod @ncliente='" & cliente & "',@modulo='" & rstAux("nmodulo") & "'"
						set rstConn = conn.execute(strselect)
						if not rstConn.EOF then
							licenciasDisp=rstConn("LICENCIAS_DISP")
						else
							licenciasDisp=0
						end if
						conn.close
						if licenciasDisp=0 then
							primero=((lote-1)*NumReg)
				 			for usuario=primero to nUsuarios-1
			 					%><script type="text/javascript" language="javascript">
			 						if (document.adminUsuarios.check<%=usuario%>i<%=modulo%>.value=="")
			 							document.adminUsuarios.check<%=usuario%>i<%=modulo%>.disabled=true;
			 					</script><%
				 			next
						end if
						%><td class="CELDAC7" style="text-align:center; padding-left:0px;"><b><div id="licencia<%=modulo%>"><%=licenciasDisp%></div></b></td><%
						rstAux.movenext
					next
					if rstAux.state<>0 then rstAux.close
                    ''comdRSave3.close
                    ''connRSave3.close
                    set comdRSave3=nothing
                    set connRSave3=nothing %>
				</tr><%
			end if
		%></table><br/><%
			 NextPrev lote,lotes,campo,criterio,texto,2,mode
		else
  			%><script type="text/javascript" language="javascript">document.adminUsuarios.sinUsuarios.value="SI";</script><%
		end if

		rstTemp.close
	'***** FIN modo edit ***********************'
   end if %>

<%' >>> MCA 21/12/04 : Incorporar gestión de franjas horarias de acceso a la administración de usuarios

if mode="guardahorarios" then
	nUser1=cint(nUsuario1)
	nUser2=cint(nUsuario2)

	GuardarHorarios_masivo nUser1,nUser2%>
	<script type="text/javascript" language="javascript">
	    window.alert("<%=LitDatosOK%>");
	</script>
	<%if Request.Form("hmode")="edit" then
		mode= "edit"
	else
		mode="gestionhorarios"
	end if
end if

if mode="gestionhorarios" then
		sentido=limpiaCadena(Request.QueryString("sentido"))
		lote=limpiaCadena(Request.QueryString("lote"))
		if lote="" then lote=1

		if sentido="next" then
			lote=lote+1
		elseif sentido="prev" then
			lote=lote-1
		end if

		strwhere= CadenaBusqueda(campo,criterio,texto)
        rstTemp.cursorlocation=3
		rstTemp.open "SELECT c.usuario,i.nombre,c.fbaja FROM CLIENTES_USERS as c with(nolock), indice as i with(nolock) WHERE c.usuario=i.entrada and c.administrar=1 and c.ncliente = '" & cliente & "' " & strwhere,DSNIlion

		lotes=rstTemp.RecordCount/NumReg
		if lotes>clng(lotes) then
			lotes=clng(lotes)+1
		else
			lotes=clng(lotes)
		end if

		nModulos=0
		nUsuarios=((lote-1)*NumReg)%>
  	<td align="right">
<% 	if permiso="admin" then %>
		&nbsp;
<%	elseif mode="edit" or mode="save" then %>
		<a class="CELDAREF" href="javascript:AbreGestionHorarios(<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(mode)%>');">
		<%=LitGestionHoraria%></a>
<% 	elseif (mode="gestionhorarios" or mode="guardahorarios") and OMC<>"SI" then %>
		<a class="CELDAREF" href="javascript:AbreGestionModulos(<%=enc.EncodeForJavascript(lote)%>,'<%=enc.EncodeForJavascript(campo)%>','<%=enc.EncodeForJavascript(criterio)%>','<%=enc.EncodeForJavascript(texto)%>','<%=enc.EncodeForJavascript(mode)%>');">
		<%=LitGestionModulos%></a>
<% end if %>
	</td>
  </tr>
  </table>
<br/>
		<input type="hidden" name="hlote" value="<%=enc.EncodeForHtmlAttribute(lote)%>"/>
		<input type="hidden" name="sinUsuarios" value=""/>
        <%if not rstTemp.eof then
			rstTemp.PageSize=NumReg
			rstTemp.AbsolutePage=lote

		NextPrev lote,lotes,campo,criterio,texto,1,mode%>
		<table width="100%" bgcolor="<%=color_blau%>" border="0" cellpading="2" cellspacing="2">
            <%DrawFila color_blau%>
				<td class="CELDAB7"></td>
				<td class="CELDAC7" colspan="2"><b> <%=LitFranjaHoraria1%></b></td>
				<td class="CELDAC7" colspan="2"><b> <%=LitFranjaHoraria2%></b></td>
            <%CloseFila
			DrawFila color_blau%>
					<td class="CELDAB7"></td>
					<td class="CELDAC7" ><b> <%=LitDesde%></b></td>
					<td class="CELDAC7" ><b> <%=LitHasta%></b></td>
					<td class="CELDAC7" ><b> <%=LitDesde%></b></td>
					<td class="CELDAC7" ><b> <%=LitHasta%></b></td>
            <%CloseFila%>
		<input type="hidden" name="hnmodulos" value="<%=enc.EncodeForHtmlAttribute(nModulos)%>"/>
        <%while not rstTemp.eof and fila<NumReg
			DrawFila color_blau
				nombreUsuario=rstTemp("nombre")
				if rstTemp("fbaja")>"" then%>
				<td class="CELDA"><%=enc.EncodeForHtmlAttribute(null_s(nombreUsuario))%> (<%=ucase(LitBaja)%>)</td><%
				else%>
				<td><a class="CELDAREF" href="javascript:AbreteSesamo('<%=enc.EncodeForJavascript(rstTemp("usuario"))%>','<%=enc.EncodeForJavascript(cliente)%>', '<%=verEncHTMLJS%>');"><%=nombreUsuario%></a></td>
                <%end if

				    seleccion="select * from indice with(nolock) where entrada='" & rstTemp("usuario") & "'"
                    rsIndice.cursorlocation=3
					rsIndice.Open seleccion, DSNIlion

					if Hour(rsIndice("accesodesdehora1"))>0 or Minute(rsIndice("accesodesdehora1"))>0 then
						if Minute(rsIndice("accesodesdehora1"))=0 then
							desdehora1= Hour(rsIndice("accesodesdehora1")) &":00"
						else
						    desdehora1= Hour(rsIndice("accesodesdehora1")) &":"& Minute(rsIndice("accesodesdehora1"))
					    end if
					else
						desdehora1= ""
					end if

					if Hour(rsIndice("accesohastahora1"))>0 or Minute(rsIndice("accesohastahora1"))>0 then
						if Minute(rsIndice("accesohastahora1"))=0 then
							hastahora1= Hour(rsIndice("accesohastahora1")) &":00"
						else
							hastahora1= Hour(rsIndice("accesohastahora1")) &":"& Minute(rsIndice("accesohastahora1"))
						end if
					else
						hastahora1= ""
					end if

					if Hour(rsIndice("accesodesdehora2"))>0 or Minute(rsIndice("accesodesdehora2"))>0 then
						if Minute(rsIndice("accesodesdehora2"))=0 then
							desdehora2= Hour(rsIndice("accesodesdehora2")) &":00"
						else
							desdehora2= Hour(rsIndice("accesodesdehora2")) &":"& Minute(rsIndice("accesodesdehora2"))
						end if
					else
						desdehora2= ""
					end if

					if Hour(rsIndice("accesohastahora2"))>0 or Minute(rsIndice("accesohastahora2"))>0 then
						if Minute(rsIndice("accesohastahora2"))=0 then
							hastahora2= Hour(rsIndice("accesohastahora2")) &":00"
						else
							hastahora2= Hour(rsIndice("accesohastahora2")) &":"& Minute(rsIndice("accesohastahora2"))
						end if
					else
						hastahora2= ""
					end if

					if desdehora1=":00" then desdehora1="" end if
					if hastahora1=":00" then hastahora1="" end if
					if desdehora2=":00" then desdehora2="" end if
					if hastahora2=":00" then hastahora2="" end if

					numUser=((lote-1)*NumReg)%>
				<td class=""CELDAC7"">
				<input type="text" class="CELDA" size="10" name="desdehora1_u<%=enc.EncodeForHtmlAttribute(nUsuarios)%>" value="<%= desdehora1%>" onchange="javascript:document.adminUsuarios.hmodificado.value=1;"/>
				<font class="CELDA7"><%=LitHoras%></font>
				</td>
				<td class="CELDAC7">
				<input type="text" class="CELDA" size="10" name="hastahora1_u<%=enc.EncodeForHtmlAttribute(nUsuarios)%>" value="<%= hastahora1%>" onchange="javascript:document.adminUsuarios.hmodificado.value=1;"/>
				<font class="CELDA7"><%=LitHoras%></font>
				</td>
				<td class="CELDAC7">
				<input type="text" class="CELDA" size="10" name="desdehora2_u<%=enc.EncodeForHtmlAttribute(nUsuarios)%>" value="<%= desdehora2%>" onchange="javascript:document.adminUsuarios.hmodificado.value=1;"/>
				<font class="CELDA7"><%=LitHoras%></font>
				</td>
				<td class="CELDAC7">
				<input type="text" class="CELDA" size="10" name="hastahora2_u<%=enc.EncodeForHtmlAttribute(nUsuarios)%>" value="<%= hastahora2%>" onchange="javascript:document.adminUsuarios.hmodificado.value=1;"/>
				<font class="CELDA7"><%=LitHoras%></font>
				</td>
                <%rsIndice.close
		CloseFila
		fila=fila+1
		rstTemp.movenext
		nUsuarios=nUsuarios+1
	 wend
 	    %>
		<input type="hidden" name="hnusuario1" value="<%=enc.EncodeForHtmlAttribute(numUser)%>"/>
		<input type="hidden" name="hnusuario2" value="<%=enc.EncodeForHtmlAttribute(nUsuarios)-1%>"/>
		<input type="hidden" name="hnusuarios" value="<%=enc.EncodeForHtmlAttribute(nUsuarios)%>"/>
	</table><br/>
            <%NextPrev lote,lotes,campo,criterio,texto,2,mode
		else%>
	<script type="text/javascript" language="javascript">document.adminUsuarios.sinUsuarios.value="SI";</script><%
		end if
		rstTemp.close
end if%>
 </form>
<%' <<< MCA 21/12/04 : Incorporar gestión de franjas horarias de acceso a la administración de usuarios
    set conn = Nothing
    set command= Nothing
    set rst = Nothing
    set rstAux =Nothing
    set rstAux2 = Nothing
    set rstAux3 = Nothing
    set rstAux4 = Nothing
    set rstAux5 = Nothing
    set rstAux6 = Nothing
    set rstAux7 = Nothing
    set rstTemp = Nothing
    set rstSelect = Nothing
    set rstConn = Nothing
    set rsIndice = Nothing
    set rstFirma = Nothing

    set enc=Nothing
end if	' accesoPagina correcto

%>
</body>
</html>