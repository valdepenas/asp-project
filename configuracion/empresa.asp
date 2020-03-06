<%@ Language=VBScript%>
   <% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>
<%

mode = Request.QueryString("mode")
if mode & ""="delete" or mode & ""="save" or mode & ""="first_save" then
	Response.Buffer = TRUE
	Response.Clear
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title></title>
<meta http-equiv="Content-Type" content="<%=imgtipo%>"; charset="<%=session("caracteres")%>"/>
<meta http-equiv="Content-style-Type" content="text/css"/>
<link rel="styleSHEET" href="../pantalla.css" media="screen"/>
<link rel="styleSHEET" href="../impresora.css" media="print"/>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../upload.asp"-->
<!--#include file="empresa.inc" -->
<!--#include file="../styles/listTable.css.inc" -->

</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">

function cifvalidation(){ 
   var fallo = false;
    cif= document.getElementsByName("cif2")[0].value;
   if(cif.length > 25 || cif.indexOf ( ' ' )!=-1){
    fallo=true;
   }
   
   
   if(fallo){
        alert("<%=LitCifCharNoValidos %>");
        return false;
    } 
    
    else {
        return true;
    }
    

}

function cambiar() {
	if (document.empresa.check1.checked){
		document.empresa.elements["pepe"].value="1";
		document.empresa.check1.value="yyy"
	}
	else{
		document.empresa.elements["pepe"].value="0";
		document.empresa.check1.value="xxx"
	}

	//ricardo 24-7-2006 se añaden dos fotos mas a empresas
	if (document.empresa.check12.checked){
		document.empresa.elements["pepe2"].value="1";
		document.empresa.check12.value="yyy"
	}
	else{
		document.empresa.elements["pepe2"].value="0";
		document.empresa.check12.value="xxx"
	}

	//ricardo 24-7-2006 se añaden dos fotos mas a empresas
	if (document.empresa.check13.checked){
		document.empresa.elements["pepe3"].value="1";
		document.empresa.check13.value="yyy"
	}
	else{
		document.empresa.elements["pepe3"].value="0";
		document.empresa.check13.value="xxx"
	}
}
function cambiar2() {
	if (document.empresa.check2.checked){
		document.empresa.elements["impresion"].value="1";
		document.empresa.check2.value="yyy"
	}
	else{
		document.empresa.elements["impresion"].value="0";
		document.empresa.check2.value="xxx"
	}
}
function cambiar3() {
	if (document.empresa.check3.checked){
		document.empresa.elements["stock"].value="1";
		document.empresa.check3.value="yyy"
	}
	else{
		document.empresa.elements["stock"].value="0";
		document.empresa.check3.value="xxx"
	}
}

function cambiar4() {
	if (document.empresa.check4.checked){
		document.empresa.elements["impresion_empresa"].value="1";
		document.empresa.check4.value="yyy";
		document.empresa.check5.disabled=false;
	}
	else{
		document.empresa.elements["impresion_empresa"].value="0";
		document.empresa.check4.value="xxx";
		document.empresa.check5.checked=false;
		document.empresa.elements["impresion_tienda"].value="0";
		document.empresa.check5.value="xxx";
		document.empresa.check5.disabled=true;
	}
}

function cambiar5()
{
	if (document.empresa.check5.checked)
	{
		document.empresa.elements["impresion_tienda"].value="1";
		document.empresa.check5.value="yyy";
	}
	else
	{
		document.empresa.elements["impresion_tienda"].value="0";
		document.empresa.check5.value="xxx";
	}
}
//FLM:20091123:función del check de logo_email
function cambiar6()
{
	if (document.empresa.check6.checked)
	{
		document.empresa.elements["logo_email"].value="1";
		document.empresa.check6.value="yyy";
	}
	else
	{
		document.empresa.elements["logo_email"].value="0";
		document.empresa.check6.value="xxx";
	}
}

function cambiar7()
{
    if (document.empresa.checkRecc.checked)
    {
        document.empresa.elements["recc"].value="1";
        document.empresa.checkRecc.value="yyy";
    }
    else
    {
        document.empresa.elements["recc"].value="0";
        document.empresa.checkRecc.value="xxx";
    }
}
//AMP 28/07/2010 : Añadimos parametro viene para controlar cuando se abre pantalla empresa para asistente de puesta en marcha.
function Editar(cif,viene)
{ 
	document.location="empresaFormulario.asp?mode=edit&cif=" + cif;
	parent.botones.document.location="empresa_bt.asp?mode=edit&viene="+viene;	
}

function VerDescuentos(ncliente)
{
	AbrirVentana("../central.asp?pag1=configuracion/valesdto.asp&pag2=configuracion/valesdto_bt.asp&ndoc=" + ncliente + "&viene=clientes&titulo=<%=LitCutlDelCli%> : " + trimCodEmpresa(ncliente),'P',<%=AltoVentana%>,<%=AnchoVentana%>);
}

function VerComisiones(ncliente)
{
	AbrirVentana("../central.asp?pag1=configuracion/comisiones.asp&pag2=configuracion/comisiones_bt.asp&ndoc=" + ncliente + "&viene=clientes&titulo=<%=LitComisionMarca%> : " + trimCodEmpresa(ncliente),'P',<%=AltoVentana%>,<%=AnchoVentana%>);
}

function ValoresDefecto(cif,ncliente)
{
    if (window.confirm("<%=LitSeguroValoresDefect%>"))
    {
	    document.location="empresa.asp?mode=valoredefecto&cif=" + cif+"&ncliente="+ncliente;
	    parent.botones.document.location="empresa_bt.asp?mode=edit";
	}
}
</script>


<%'***RGU 24/7/2006: Añadir gestion en el formulario de los campos Administrador, cargo, texto1-2-3-4-5(campo01-02-03-04-05)
%>

<% 
	'Leer parámetros de la página'
	mode = Request.QueryString("mode")
	viene=limpiaCadena(Request.QueryString("viene"))
	
  

	if Request.QueryString("cif")>"" then
		TmpCif=limpiaCadena(Request.QueryString("cif"))
		if  mode<>"first_save" then
		checkCadena TmpCif
		end if
	end if
	if mode="delete" then
		''Response.Buffer = TRUE
		''Response.Clear

		byteCount = Request.TotalBytes

		RequestBin = Request.BinaryRead(byteCount)
		Dim UploadRequest
		Set UploadRequest = CreateObject("Scripting.Dictionary")

		BuildUploadRequest  RequestBin

		TmpCif=Nulear(UploadRequest.Item("cif").Item("Value"))
		TmpCif=limpiaCadena(TmpCif)
		'checkCadena TmpCif

        'primeros miramos si hay alguna serie con esta empresa
        strselect="select top 1 nserie from series with(nolock) where empresa=?"
        existnserie= DLookupP1(strselect,TmpCif&"",adVarChar,25,session("dsn_cliente"))
		set rst = Server.CreateObject("ADODB.Recordset")
   
        if existnserie&""="" then
            'delete
            set conn=server.CreateObject("ADODB.connection")
			set cmd=server.CreateObject("ADODB.Command")
			conn.open session("dsn_cliente")
			conn.cursorlocation=2
			cmd.ActiveConnection=conn
			cmd.commandText= "deleteCompany"
			cmd.CommandType = adCmdStoredProc
			cmd.Parameters.Append cmd.CreateParameter("@code",adVarChar,adParamInput,25, TmpCif&"") 
			set rst = cmd.Execute
			rst.close
            conn.close
            set rst=nothing
            set command = nothing
			set conn = nothing	
		else
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitNosepuedeborrar%>");
			</script><%
		end if

		mode="browse"
		TmpCif=""
	end if

	if mode="save" or mode="first_save" then
		''Response.Buffer = TRUE
		''Response.Clear

		byteCount = Request.TotalBytes

		RequestBin = Request.BinaryRead(byteCount)
		'Dim UploadRequest'
		Set UploadRequest = CreateObject("Scripting.Dictionary")

		BuildUploadRequest  RequestBin

		TmpCif=Nulear(UploadRequest.Item("cif").Item("Value"))
		TmpCif=limpiaCadena(TmpCif)
		'checkCadena TmpCif
		guarda=true

		'set rst = Server.CreateObject("ADODB.Recordset")
        'rst.cursorlocation=2
		'rst.Open "select * from empresas where cif='" & iif(mode="first_save",session("ncliente"),"") & TmpCif & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        
        strselect="select * from empresas where cif=?"
        set conn = Server.CreateObject("ADODB.Connection")
        set rst = Server.CreateObject("ADODB.Recordset")
	    set command =  Server.CreateObject("ADODB.Command")
	    conn.open session("dsn_cliente")
        
	    command.ActiveConnection =conn
	    command.CommandTimeout = 60
	    command.CommandText=strselect
	    command.CommandType = adCmdText 'CONSULTA
        
        if mode="fisrt_save" then
            ncif=session("ncliente")&TmpCif
        else
           ncif=TmpCif
        end if
        command.Parameters.Append command.CreateParameter("@cif", adVarChar, adParamInput, 25, ncif&"")
       ' rst.CursorLocation = adUseClient
        rst.Open command, , adOpenKeyset, adLockOptimistic

		if mode="first_save" then
			if rst.eof then
				rst.addnew
				rst("cif")=session("ncliente") & TmpCif
			else
				guarda=false
			end if
		end if

		if guarda=true then
			picture = UploadRequest.Item("logotipo").Item("Value")
			contentType = UploadRequest.Item("logotipo").Item("ContentType")
			filepathname = UploadRequest.Item("logotipo").Item("FileName")

''ricardo 24-7-2006 se añaden dos fotos mas a empresas
			picture2 = UploadRequest.Item("logotipo2").Item("Value")
			contentType2 = UploadRequest.Item("logotipo2").Item("ContentType")
			filepathname2 = UploadRequest.Item("logotipo2").Item("FileName")

			picture3 = UploadRequest.Item("logotipo3").Item("Value")
			contentType3 = UploadRequest.Item("logotipo3").Item("ContentType")
			filepathname3 = UploadRequest.Item("logotipo3").Item("FileName")
			

			
'''''''''''''''''''''''''''''''''''''''''''''''''

			rst("nombre")		= Nulear(UploadRequest.Item("nombre").Item("Value"))
			rst("nombrecom")	= Nulear(UploadRequest.Item("nombrecom").Item("Value"))
			rst("direccion")	= Nulear(UploadRequest.Item("direccion").Item("Value"))
			rst("poblacion")	= Nulear(UploadRequest.Item("poblacion").Item("Value"))
			rst("provincia")	= Nulear(UploadRequest.Item("provincia").Item("Value"))
			rst("cp")			= Nulear(UploadRequest.Item("cp").Item("Value"))
			rst("pais")			= Nulear(UploadRequest.Item("pais").Item("Value"))
			rst("telefono")		= Nulear(UploadRequest.Item("telefono").Item("Value"))
			rst("telefono2")	= Nulear(UploadRequest.Item("telefono2").Item("Value"))
			rst("fax")			= Nulear(UploadRequest.Item("fax").Item("Value"))
			rst("email")		= Nulear(UploadRequest.Item("email").Item("Value"))
			rst("leyenda")		= Nulear(UploadRequest.Item("leyenda").Item("Value"))
			rst("pie_mail")		= Nulear(UploadRequest.Item("pie_mail").Item("Value"))
			rst("objeto_social")= Nulear(UploadRequest.Item("objeto_social").Item("Value"))
			rst("IRPF")			= Null_z(UploadRequest.Item("IRPF").Item("Value"))
			
            ''DGB: RECC
                'FLM:20091123
			if UploadRequest.Item("recc").Item("Value") = "1" then
				rst("RECC")=1
			else
				rst("RECC")=0
			end if

			if UploadRequest.Item("impresion").Item("Value") = "1" then
				rst("print_logo")=1
			else
				rst("print_logo")=0
			end if
			if UploadRequest.Item("impresion_empresa").Item("Value") = "1" then
				rst("print_empresa")=1
			else
				rst("print_empresa")=0
			end if

			if UploadRequest.Item("stock").Item("Value") = "1" then
				rst("control_stock")=1
			else
				rst("control_stock")=0
			end if
			'FLM:20091123
			if UploadRequest.Item("logo_email").Item("Value") = "1" then
				rst("logo_email")=1
			else
				rst("logo_email")=0
			end if

			if UploadRequest.Item("impresion_tienda").Item("Value") = "1" then
				rst("print_tienda")=1
			else
				rst("print_tienda")=0
			end if

			if UploadRequest.Item("pepe").Item("Value") = "1" then
				rst("logotipo")= pepito
				rst("tipo_logo")= pepito
			else
				if picture >"" then
					picturechunk =  picture & chrB(0)
					rst("logotipo").appendChunk picturechunk
					rst("tipo_logo") = contentType
				end if
			end if
''ricardo 24-7-2006 a añaden dos fotos mas a empresas
			if UploadRequest.Item("pepe2").Item("Value") = "1" then
				rst("logotipo2")= pepito2
				rst("tipo_logo2")= pepito2
			else
				if picture2 >"" then
					picturechunk2 =  picture2 & chrB(0)
					rst("logotipo2").appendChunk picturechunk2
					rst("tipo_logo2") = contentType2
				end if
			end if

			if UploadRequest.Item("pepe3").Item("Value") = "1" then
				rst("logotipo3")= pepito3
				rst("tipo_logo3")= pepito3
			else
				if picture3 >"" then
					picturechunk3 =  picture3 & chrB(0)
					rst("logotipo3").appendChunk picturechunk3
					rst("tipo_logo3") = contentType3
				end if
			end if

            'dfs 02/07/2014. Se comprueba si el cliente tiene el modulo de Teekit para actualizar la fecha de modificacion del logo si lo cambia
            'if (picture > "" or picture2 > "" or picture3 > "") and (ModuloContratado (session("ncliente"), ModTeekit)) then
                'rst("LOGO_MODIFY_DATE") = Date
            'end if 


'''''''''''''''''''''''''''''''

			'**RGU 24/7/2006
			rst("administrador")= Nulear(UploadRequest.Item("admin").Item("Value"))
			rst("cargo")= Nulear(UploadRequest.Item("cargo").Item("Value"))
			rst("campo01")= Nulear(UploadRequest.Item("T1").Item("Value"))
			rst("campo02")= Nulear(UploadRequest.Item("T2").Item("Value"))
			rst("campo03")= Nulear(UploadRequest.Item("T3").Item("Value"))
			rst("campo04")= Nulear(UploadRequest.Item("T4").Item("Value"))
			rst("campo05")= Nulear(UploadRequest.Item("T5").Item("Value"))
			'**RGU

            '**ASP 23/12/2010
            'mejico=d_lookup("gestion_folios","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")) & ""
            strselect="select gestion_folios from configuracion with(nolock) where nempresa=?"
            mejico= DLookupP1(strselect,session("ncliente")&"",adVarChar,5,session("dsn_cliente"))
			if mejico then
            rst("calle")	    = Nulear(UploadRequest.Item("calle").Item("Value"))
			rst("ninterior")	= Nulear(UploadRequest.Item("ninterior").Item("Value"))
			rst("nexterior")	= Nulear(UploadRequest.Item("nexterior").Item("Value"))
			rst("colonia")		= Nulear(UploadRequest.Item("colonia").Item("Value"))
			rst("municipio")	= Nulear(UploadRequest.Item("municipio").Item("Value"))
			end if
            '**ASP
			rst.update
			rst.close   
            conn.close      
			set rst=nothing
            set command = nothing
			set conn = nothing	
			
			
			'Comprobamos si se ha cambiado el cif para lanzar el procedimiento que realizará el cambio en la DB
			cifold=limpiaCadena(Nulear(UploadRequest.Item("cif").Item("Value")))
			cifnew=cifold
			if   mode<>"first_save" then 
			cifnew=session("ncliente") & limpiaCadena(Nulear(UploadRequest.Item("cif2").Item("Value")))
			if  mode<>"first_save" then
		        checkCadena cifnew
		    end if
			end if

			if cifold <> cifnew and  InStr(cifnew, " ") = 0 AND len(cifnew)<26 then
			   
			   
    			
			   
			    'llamada procedimiento
                set conn = Server.CreateObject("ADODB.Connection")
                conn.open session("dsn_cliente")
                
		        set rstCif = Server.CreateObject("ADODB.Recordset")
		        set command =  Server.CreateObject("ADODB.Command")
			    command.ActiveConnection =conn
		        command.CommandTimeout = 0
		        command.CommandText="ModificacionCifEmpresa"
		        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		        command.Parameters.Append command.CreateParameter("@CifOld",adVarChar,adParamInput,25,cifold)
		        command.Parameters.Append command.CreateParameter("@CifNew",adVarChar,adParamInput,25,cifnew)
		        set rstCif=command.Execute
			    if not rstCif.eof then
			        resultado=rstCif("Respuesta")
			        if (resultado=0) then
			            
			            auditar_ins_bor session("usuario"),cifold, cifnew, "modificar", "","","empresa"
			        else
			             if (resultado=1) then
			         
			                %><script language="javascript" type="text/javascript">
				            window.alert("<%=LitEmpresaExiste %>");
            				
			                </script><%  
			                
			          
			            else 'resultado = 2
    			            
			            end if
			        end if
    			  
			    end if
			    rstCif.close
			    set  rstCif=nothing
			    set command=nothing
                
			else
			end if   
            if viene="asistente" then
			  %><script language="javascript" type="text/javascript">
				window.alert("<%=LitDatosGuardados%>");
				parent.botones.document.location="empresa_bt.asp?mode=browse&viene=asistente";
			</script><% 
			else			 
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitDatosGuardados%>");
				parent.botones.document.location="empresa_bt.asp?mode=browse";
			</script><%
			end if
		else
			rst.close
			%><script language="javascript" type="text/javascript">
				window.alert ("<%=LitEmpresaExiste%>")
		  		parent.botones.document.location="empresa_bt.asp?mode=browse";
			</script><%
		end if

		TmpCif=""
		mode="browse"
	end if%>


<body  class="BODY_ASP">
	<form name="empresa" method="post" enctype="multipart/form-data">
        <input type="hidden" name="mode_accesos_tienda" value="<%=enc.EncodeForHtmlAttribute(mode)%>"/>
	    <input type="hidden" name="pepe" value="0"/>
	    <input type="hidden" name="pepe2" value="0"/>
	    <input type="hidden" name="pepe3" value="0"/>
<%
	'Recordsets'
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")
    
   PintarCabecera "empresa.asp"
   alarma "empresa.asp"

    if mode="valoredefecto" then
    		DSNCliente=session("dsn_cliente")
			donde=instr(1,DSNCliente,"Initial Catalog=",1)
			donde2=instr(donde,DSNCliente,";",1)
			BD=mid(DSNCliente,donde+16,donde2-donde-16)
            'rst.cursorlocation=2
            'rst.open "exec [ValoresDefectoEmpresa]	@nempresa ='"&limpiaCadena(request.QueryString("cif"))&"',@ncliente ='"&session("ncliente")&"' ",replace(DSNILION,"ilion_admin",BD)
            set conn=server.CreateObject("ADODB.connection")
			set cmd=server.CreateObject("ADODB.Command")
			conn.open replace(DSNILION,"ilion_admin",BD)
			conn.cursorlocation=3
			cmd.ActiveConnection=conn
		    cmd.CommandText="ValoresDefectoEmpresa"
		    cmd.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		    cmd.Parameters.Append cmd.CreateParameter("@nempresa",adVarChar,adParamInput,50,limpiaCadena(request.QueryString("cif")))
		    cmd.Parameters.Append cmd.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
		    set rst=cmd.Execute
        if not rst.eof then
            if rst("error")=0 then
                %><script language="javascript" type="text/javascript">
                    alert ("<%=LitMsgDatosCreados %>");
                </script><%
            else 
                %><script language="javascript" type="text/javascript">
                    alert ("<%=LitError%> "&rst("error")&":<%=LitMsgDatosNoCreados %>");
                </script><%
            end if
        end if
        rst.close
        conn.close
		set rst = nothing
		set cmd = nothing
		set conn = nothing	
        mode="edit"
    end if
    
    ''ricardo 5-11-2009 se obtiene el parametro cpo
    dim cpo
    ObtenerParametros("empresas")
	       
	if mode="browse" then%>
			<hr/>
			<%
			
			strselect="select * from empresas with(NOLOCK) where cif like ?+'%'"
            set conn=server.CreateObject("ADODB.connection")
            set cmd=server.CreateObject("ADODB.Command")
			conn.open session("dsn_cliente")
			conn.cursorlocation=3
			cmd.ActiveConnection=conn
			cmd.commandText= strselect
			cmd.CommandType = adCmdText
			cmd.Parameters.Append cmd.CreateParameter("@nempresa",adVarChar,adParamInput,5, session("ncliente")&"") 
            set rst=cmd.Execute
			if not rst.eof then
				%>
				<table width="100%" bgcolor="<%=color_blau%>" border="0"><%				  
					DrawFila color_fondo
						DrawCelda "CELDA","","",0,"<b>" & LitNombre & "</b>"
						DrawCelda "CELDA","","",0,"<b>" & LitCif & "</b>"
						DrawCelda "CELDA","","",0,"<b>" & LitDireccion & "</b>"
						DrawCelda "CELDA","","",0,"<b>" & LitPoblacion & "</b>"
						DrawCelda "CELDA","","",0,"<b>" & LitCp & "</b>"
						DrawCelda "CELDA","","",0,"<b>" & LitProvincia & "</b>"
						DrawCelda "CELDA","","",0,"<b>" & LitTelefono & "</b>"
					CloseFila
					
					while not rst.eof
						DrawFila color_blau
							'DrawCelda "CELDA","","",0,rst("nombre")'
							ref="javascript:Editar('" & rst("cif") & "','"&enc.EncodeForJavascript(viene)&"');"
							DrawCeldahref "CELDAREF7","left",false,enc.EncodeForHtmlAttribute(replace(rst("nombre"),"'","&#39;")),ref
							DrawCelda "CELDA","","",0,trimCodEmpresa(rst("cif"))
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(""&rst("direccion"))
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(""&rst("poblacion"))
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(""&rst("cp"))
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(""&rst("provincia"))
							DrawCelda "CELDA","","",0,enc.EncodeForHtmlAttribute(""&rst("telefono"))
						CloseFila
						rst.movenext
					wend%>
				</table>
				<input type="hidden" name="cantidad" value="<%=enc.EncodeForHtmlAttribute(rst.recordcount)%>"/>
			<%else%>
				<input type="hidden" name="cantidad" value="0"/>
			<%end if
		rst.close
		conn.close
		set rst = nothing
		set command = nothing
		set conn = nothing	
		''ricardo 8-10-2009 se obtiene, si esta pantalla tiene algun limite en el numero de empresas creadas
		limiteEmpresasCreadas=0
		
		'rst.open "exec limitesPagina '" & session("ncliente") & "','empresa.asp'",dsnilion
        set conn=server.CreateObject("ADODB.connection")
        set cmd=server.CreateObject("ADODB.Command")
		conn.open DsnIlion
		conn.cursorlocation=3
		cmd.ActiveConnection=conn
		cmd.commandText= "limitesPagina"
		cmd.CommandType = adCmdStoredProc
		cmd.Parameters.Append cmd.CreateParameter("@ncliente",adVarChar,adParamInput,5, session("ncliente")&"") 
        cmd.Parameters.Append cmd.CreateParameter("@pagina_asp",adVarChar,adParamInput,100, "empresa.asp") 
        set rst=cmd.Execute
		if not rst.eof then
		    if rst("limite")& ""="" or isnumeric(rst("limite"))=0 then
		        limiteEmpresasCreadas=0
		    else
		        limiteEmpresasCreadas=rst("limite")
		    end if
		end if
		rst.close
        conn.close
		set rst = nothing
		set command = nothing
		set conn = nothing	        
                %>
        <input type="hidden" name="limiteEmpresasCreadas" value="<%=enc.EncodeForHtmlAttribute(limiteEmpresasCreadas)%>"/>
    <%elseif mode="add" then
          
            set conn=server.CreateObject("ADODB.connection")
			set command=server.CreateObject("ADODB.Command")
			conn.open DSNIlion
			conn.cursorlocation=3
			command.ActiveConnection=conn
		    command.CommandText="select cif,rsocial,ncomercial,domicilio,poblacion,provincia,cp,pais,telefono,telefono2,fax,email,web,lenguaje,caracteres from clientes with(NOLOCK) where ncliente=?"
		    command.CommandType = adCmdText 'Procedimiento Almacenado
		    command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
		    set rstAux=command.Execute
             if not rstAux.eof then      
                strselect="select cif from empresas with(NOLOCK) where cif=?"
                existe= DLookupP1(strselect,session("ncliente")&rstAux("cif"),adVarChar,25,session("dsn_cliente"))
                
                if existe&""="" then
                    set conn2=server.CreateObject("ADODB.connection")
			        set command2=server.CreateObject("ADODB.Command")         
			        conn2.open session("dsn_cliente")
			        conn2.cursorlocation=2
			        command2.ActiveConnection=conn2
		            command2.CommandText="setCompanyBasic"
		            command2.CommandType = adCmdStoredProc 'Procedimiento Almacenado
		            command2.Parameters.Append command2.CreateParameter("@code",adVarChar,adParamInput,25,session("ncliente")&rstAux("cif"))
		            command2.Parameters.Append command2.CreateParameter("@name",adVarChar,adParamInput,50,rstAux("rsocial"))
                    command2.Parameters.Append command2.CreateParameter("@name_comercial",adVarChar,adParamInput,45,left(rstAux("ncomercial"),45))
                    command2.Parameters.Append command2.CreateParameter("@address",adVarChar,adParamInput,50,rstAux("domicilio"))
                    command2.Parameters.Append command2.CreateParameter("@town",adVarChar,adParamInput,50,rstAux("poblacion"))
                    command2.Parameters.Append command2.CreateParameter("@state",adVarChar,adParamInput,30,rstAux("provincia"))
                    command2.Parameters.Append command2.CreateParameter("@cp",adVarChar,adParamInput,10,rstAux("cp"))
                    command2.Parameters.Append command2.CreateParameter("@country",adVarChar,adParamInput,30,rstAux("pais"))
                    command2.Parameters.Append command2.CreateParameter("@phone",adVarChar,adParamInput,30,rstAux("telefono"))
                    command2.Parameters.Append command2.CreateParameter("@phone2",adVarChar,adParamInput,30,rstAux("telefono2"))
                    command2.Parameters.Append command2.CreateParameter("@fax",adVarChar,adParamInput,30,rstAux("fax"))
                    command2.Parameters.Append command2.CreateParameter("@email",adVarChar,adParamInput,50,rstAux("email"))
                    command2.Parameters.Append command2.CreateParameter("@internet",adVarChar,adParamInput,50,rstAux("web"))
                    command2.Parameters.Append command2.CreateParameter("@language",adVarChar,adParamInput,10,rstAux("lenguaje"))
                    command2.Parameters.Append command2.CreateParameter("@characters",adVarChar,adParamInput,20,rstAux("caracteres"))
                    command2.Parameters.Append command2.CreateParameter("@legends",adVarChar,adParamInput,250,null)
                    set rst2=command2.Execute  
                    
                    'rst2.close 
                    conn2.close
                    set command2=nothing
                    set rst2 = nothing
                    set conn2= nothing

                end if
		      
			end if
            rstAux.close
            conn.close
            set command=nothing
            set rstAux = nothing
            set conn= nothing
           
        %>
		<hr/>
		<table width="100%" bgcolor="<%=color_blau%>" border="0"><%
			DrawFila color_blau
				DrawCelda "CELDA","","",0,LitLogotipo1 & LitLogotipomaxtam1 & LitTamnyFotoEmp1 & LitLogotipomaxtam2 & "):"%>
				<td><input class="celda" type="file" name="logotipo"/></td>
				<%
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA","","",0,LitLogotipo2 & LitLogotipomaxtam1 & LitTamnyFotoEmp2 & LitLogotipomaxtam2 & "):"%>
				<td><input class="celda" type="file" name="logotipo2"/></td>
				<%
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA","","",0,LitLogotipo3 & LitLogotipomaxtam1 & LitTamnyFotoEmp3 & LitLogotipomaxtam2 & "):"%>
				<td><input class="celda" type="file" name="logotipo3"/></td>
				<%
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitCif + ":"
				DrawInputCelda "CELDA maxlength='20'","","",40,0,"","cif",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitNombre + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","nombre",""
				DrawCelda "CELDA style='width:100px'","","",0,LitNombreCom + ":"
				DrawInputCelda "CELDA maxlength='45'","","",40,0,"","nombrecom",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA","","",0,LitAdmin + ":"
				DrawInputCelda "CELDA maxlength='100'","","",40,0,"","admin",""
				DrawCelda "CELDA","","",0,LitCargo + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","cargo",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA","","",0,LitDireccion + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","direccion",""
			CloseFila
			'mejico=d_lookup("gestion_folios","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")) & ""
            strselect="select gestion_folios from configuracion with(nolock) where nempresa=?"
            mejico= DLookupP1(strselect,session("ncliente")&"",adVarChar,5,session("dsn_cliente"))
			if mejico then
			    DrawFila color_blau%><td colspan='4'><table width="100%" bgcolor="<%=color_blau%>" border="0"><%
			    DrawFila color_blau
			        DrawCelda "CELDA colspan='6'","","",0,LitDireccionFacturaElectronica & ":"
			    CloseFila
			    DrawFila color_blau
			        DrawCelda "CELDA ","12","",2,LitCalle & ":"
			        DrawInputCelda "CELDA maxlength='50'","29","",50,0,"","calle",""
			        DrawCelda "CELDA ","","",0,LitNExterior & ":"
			        DrawInputCelda "CELDA maxlength='20'","","",20,0,"","nexterior",""
			        DrawCelda "CELDA ","","",2,LitNInterior & ":"
			        DrawInputCelda "CELDA maxlength='20'","","",20,2,"","ninterior",""
			    CloseFila
			    DrawFila color_blau
			        DrawCelda "CELDA","","",2,LitColonia & ":"
			        DrawInputCelda "CELDA maxlength='50'","","",50,0,"","colonia",""
			        DrawCelda "CELDA","","",0,LitMunicipio & ":"
			        DrawInputCelda "CELDA maxlength='50' colspan='4'","","",50,0,"","municipio",""
			    CloseFila
			    %></table></td><%
			    CloseFila
		    end if
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitPoblacion + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","poblacion",""
				DrawCelda "CELDA","","",0,LitProvincia + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","provincia",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitCp + ":"
				DrawInputCelda "CELDA maxlength='10'","","",10,0,"","cp",""
				DrawCelda2 "CELDA", "left", false, LitPais  + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","pais",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitTelefono + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","telefono",""
				DrawCelda "CELDA","","",0,LitTelefono2 + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","telefono2",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitFax + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","fax",""
				DrawCelda "CELDA","","",0,LitEmail + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","email",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:150px'","","",0,LitIRPF + ":"
				DrawInputCelda "CELDA maxlength='3'","","",5,0,"","IRPF",""
                DrawCelda "CELDA style='width:150px'","","",0,LitRECC + ":"
                %>
               <td><input type="hidden" name="recc" value="0"/>
                   <input type="checkbox" name="checkRecc" value="true" onclick="cambiar7();"/>
               </td>
            <%
			CloseFila
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitImpresionLogo + ":"%>
				<td class="CELDA">
					<input type="hidden" name="impresion" value="0"/>
					<input type="checkbox" name="check2" value="true" onclick="cambiar2();"/><%
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class='CELDA' align='left'>" & LitImpresionEmpresa & ":</font>")%>
					<input type="hidden" name="impresion_empresa" value="0"/>
					<input type="checkbox" name="check4" value="true" onclick="cambiar4();"/>
				</td>
				<%DrawCelda2 "CELDA", "left", false, LitImpresionTienda + ":"%>
				<td class="CELDA">
					<input type="hidden" name="impresion_tienda" value="0"/>
					<input type="checkbox" name="check5" value="false" onclick="cambiar5();" disabled="disabled"/><%
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class='CELDA' align='left'>" & LitControlStock & ":</font>")
					session("control_stock") = "desactivado"%>
					<input type="hidden" name="stock" value="0"/>
					<input type="checkbox" name="check3" value="true" onclick="cambiar3();"/>
					<%response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class='CELDA' align='left'>" & LitLogoEmail & ":</font>")%>
					<input type="hidden" name="logo_email" value="0"/>
					<input type="checkbox" name="check6" value="true" onclick="cambiar6();"/>
				</td>
			<%CloseFila%>
		</table>
		<table width="100%" bgcolor="<%=color_blau%>" border="0">
		    <%DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitLeyenda + ":"
				DrawTextCelda "CELDA maxlength='250'","","",5,50,"","leyenda",""
				DrawCelda "CELDA style='width:100px'","","",0,LitT1 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T1",""
				DrawCelda "CELDA style='width:100px'","","",0,LitT2 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T2",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitObjetoSocial + ":"
				DrawTextCelda "CELDA","","",5,50,"","objeto_social",""
				DrawCelda "CELDA style='width:100px'","","",0,LitT3 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T3",""
				DrawCelda "CELDA style='width:100px'","","",0,LitT4 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T4",""
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitPieMail + ":"
				DrawTextCelda "CELDA","","",5,50,"","pie_mail",""
				DrawCelda "CELDA style='width:100px'","","",0,LitT5 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T5",""

			CloseFila%>
		</table>
	<%elseif mode="edit" then
					
            strselect="select * from empresas with(NOLOCK) where cif like ?+'%'"
            set conn=server.CreateObject("ADODB.connection")
            set cmd=server.CreateObject("ADODB.Command")
			conn.open session("dsn_cliente")
			conn.cursorlocation=3
			cmd.ActiveConnection=conn
			cmd.commandText= strselect
			cmd.CommandType = adCmdText
			cmd.Parameters.Append cmd.CreateParameter("@nempresa",adVarChar,adParamInput,5, session("ncliente")&"") 
            set rst=cmd.Execute
			if not rst.eof then%>
				<input type="hidden" name="cantidad" value="<%=enc.EncodeForHtmlAttribute(rst.recordcount)%>"/>
			<%else%>
				<input type="hidden" name="cantidad" value="0"/>
			<%end if
			rst.close
            
			'rst.open "select * from empresas with(NOLOCK) where cif='" & TmpCif & "'",session("dsn_cliente")
            strselect="select * from empresas with(NOLOCK) where cif= ?"
            set conn=server.CreateObject("ADODB.connection")
            set cmd=server.CreateObject("ADODB.Command")
			conn.open session("dsn_cliente")
			conn.cursorlocation=3
			cmd.ActiveConnection=conn
			cmd.commandText= strselect
			cmd.CommandType = adCmdText
			cmd.Parameters.Append cmd.CreateParameter("@cif",adVarChar,adParamInput,25, TmpCif &"") 
            set rst=cmd.Execute
			
			if not rst.eof then
				if rst("tipo_logo")>"" and not isnull(rst("tipo_logo")) then
					mostrar_logo1 = true
				else
					mostrar_logo1 = false
				end if
				if rst("tipo_logo2")>"" and not isnull(rst("tipo_logo2")) then
					mostrar_logo2 = true
				else
					mostrar_logo2 = false
				end if
				if rst("tipo_logo3")>"" and not isnull(rst("tipo_logo3")) then
					mostrar_logo3 = true
				else
					mostrar_logo3 = false
				end if
			else
				mostrar_logo1 = false
				mostrar_logo2 = false
				mostrar_logo3 = false
			end if%>

			<input type="hidden" name="cif" value="<%=enc.EncodeForHtmlAttribute(""&rst("cif"))%>"/>
            <table width="100%">
                <tr>                    
                    <td valign="middle" align="left">
                        <font class="CELDA7"><%=LitCif%>:</font>
                        <font class="CELDAB7"><%=trimCodEmpresa(rst("cif"))%></font>
				    </td>
				    <td valign="middle" align="left">
			            <font class="CELDA7"><%=LitNombre%>:</font>
			            <font class="CELDAB7"><%=enc.EncodeForHtmlAttribute(""&rst("nombre"))%></font>
				    </td>
				    <%
                   
                    strselect= "select count(*) as series from series with(nolock) where empresa like ?+'%' and empresa is not null"
                    set conn2=server.CreateObject("ADODB.connection")
                    set cmd2=server.CreateObject("ADODB.Command")
			        conn2.open session("dsn_cliente")
			        conn2.cursorlocation=3
			        cmd2.ActiveConnection=conn2
			        cmd2.commandText= strselect
			        cmd2.CommandType = adCmdText
			        cmd2.Parameters.Append cmd2.CreateParameter("@ncompany",adVarChar,adParamInput,5, session("ncliente") &"") 
                    set rstAux=cmd2.Execute
				    if not rstAux.EOF then 
				        if rstAux("series")=0 then%>
				        <td><a class="CELDAREF" href="javascript:ValoresDefecto('<%=enc.EncodeForJavascript(""&rst("cif"))%>','<%=enc.EncodeForJavascript(cliente)%>');"><%=LitValoresDefecto%></a></td>			  				    
				        <%end if
				    end if
				    rstAux.close
                    conn2.Close
                    set cmd2=nothing
                    set conn2=nothing
                    set rstAux=nothing
                            %>
				    <td valign="middle" align="right">
				        <%  '------------------------------------------------
				            '   GPD (14/03/2007)
				            if rst("ACTIVARVALEDTO") then%>				        
                        <table style=" border-collapse : collapse;" cellspacing="0" cellpadding="3">
                            <tr>
                                <td style="border : 1px solid black;" align="center" class="CELDACENTERB" onmouseover="this.className='TDACTIVO8'" onmouseout="this.className='CELDACENTERB'" bgcolor="<%=color_blau%>">
                                <a class="CELDAREFB7" href="javascript:VerDescuentos('<%=enc.EncodeForJavascript(TmpCif)%>');" onmouseover="self.status='<%=LitIrCorreo%>'; return true;" onmouseout="self.status=''; return true;">
                                &nbsp;&nbsp;&nbsp;<%=ListConfigValeDto%>&nbsp;&nbsp;&nbsp;
                                </a>
                                </td>
                            </tr>
                        </table>                                                        
                        
                        </td>
                            <%end if
                            ''ricardo 4-11-2009 se añaden las comisiones extras por empresa
                        if cstr(cpo & "")="1" then%>				        
				            <td valign="middle" align="right">
                                <table style=" border-collapse : collapse;" cellspacing="0" cellpadding="3">
                                    <tr>
                                        <td style="border : 1px solid black;" align="center" class="CELDACENTERB" onmouseover="this.className='TDACTIVO8'" onmouseout="this.className='CELDACENTERB'" bgcolor="<%=color_blau%>">
                                        <a class="CELDAREFB7" href="javascript:VerComisiones('<%=enc.EncodeForJavascript(TmpCif)%>');" onmouseover="self.status='<%=LitIrComision%>'; return true;" onmouseout="self.status=''; return true;">
                                        &nbsp;&nbsp;&nbsp;<%=ListComisiones%>&nbsp;&nbsp;&nbsp;
                                        </a>
                                        </td>
                                    </tr>
                                </table>                                                        
				            </td>				  
				       <%end if %>
			   </tr>
			</table>
		<hr/>
		<table width="100%" bgcolor="<%=color_blau%>" border="0"><%
			DrawFila color_blau
				if mostrar_logo1 = true then
					%><td>
						<img src="muestra_logo.asp?cif=<%=enc.EncodeForJavascript(""&rst("cif"))%>&empresa=<%=session("ncliente")%>&viene=datos_empresa1" width="200" height="100" alt="" title=""/>
					</td><%
				else
					%><td class="ENCABEZADOC" style='border: 1px solid Black;' width='200' height='100' valign="middle" bgcolor="<%=color_blanc%>">
						<%=LitSinImagenEmpresa%>
					</td><%
				end if%>
				<td>
					<table width="100%" bgcolor="<%=color_blau%>" border="0">
						<tr><%
							DrawCelda "CELDA","","",0,LitLogotipo1 & LitLogotipomaxtam1 & LitTamnyFotoEmp1 & LitLogotipomaxtam2 & ":"%>
							<td><input type="file" name="logotipo"/></td>
						</tr>
						<tr><%
							DrawCelda2 "CELDA", "left", false, LitBorrar + ":"%>
							<td class="CELDA" >
								<input type="checkbox" name="check1" value="true" onclick="cambiar();"/>
							</td>
						</tr>
					</table>
				</td><%
			CloseFila
			DrawFila color_blau
				if mostrar_logo2 = true then
					%><td>
						<img src="muestra_logo.asp?cif=<%=enc.EncodeForJavascript(""&rst("cif"))%>&empresa=<%=session("ncliente")%>&viene=datos_empresa2" width="200" height="100" alt="" title=""/>
					</td><%
				else
					%><td class="ENCABEZADOC" style='border: 1px solid Black;' width='200' height='100' valign="middle" bgcolor="<%=color_blanc%>">
						<%=LitSinImagenEmpresa%>
					</td><%
				end if%>
				<td>
					<table width="100%" bgcolor="<%=color_blau%>" border="0">
						<tr><%
							DrawCelda "CELDA","","",0,LitLogotipo2 & LitLogotipomaxtam1 & LitTamnyFotoEmp2 & LitLogotipomaxtam2 & ":"%>
							<td><input type="file" name="logotipo2"/></td>
						</tr>
						<tr><%
							DrawCelda2 "CELDA", "left", false, LitBorrar + ":"%>
							<td class="CELDA" >
								<input type="checkbox" name="check12" value="true" onclick="cambiar();"/>
							</td>
						</tr>
					</table>
				</td><%
			CloseFila
			DrawFila color_blau
				if mostrar_logo3 = true then
					%><td>
						<img src="muestra_logo.asp?cif=<%=enc.EncodeForJavascript(""&rst("cif"))%>&empresa=<%=session("ncliente")%>&viene=datos_empresa3" width="200" height="100" alt="" title=""/>
					</td><%
				else
					%><td class="ENCABEZADOC" style='border: 1px solid Black;' width='200' height='100' valign="middle" bgcolor="<%=color_blanc%>">
						<%=LitSinImagenEmpresa%>
					</td><%
				end if%>
				<td>
					<table width="100%" bgcolor="<%=color_blau%>" border="0">
						<tr><%
							DrawCelda "CELDA","","",0,LitLogotipo3 & LitLogotipomaxtam1 & LitTamnyFotoEmp3 & LitLogotipomaxtam2 & ":"%>
							<td><input type="file" name="logotipo3"/></td>
						</tr>
						<tr><%
							DrawCelda2 "CELDA", "left", false, LitBorrar + ":"%>
							<td class="CELDA" >
								<input type="checkbox" name="check13" value="true" onclick="cambiar();"/>
							</td>
						</tr>
					</table>
				</td><%
			CloseFila%>
		</table>
		<table width="100%" bgcolor="<%=color_blau%>" border="0"><%
		    DrawFila color_blau
		        DrawCelda "CELDA style='width:100px'","","",0,LitCif + ":"
		        DrawInputCelda "CELDA maxlength='20' onchange=""cifvalidation();"" ","","",40,0,"","cif2",enc.EncodeForHtmlAttribute(replace(trimcodempresa(rst("cif"))&"","'","&#39;"))
		    CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitNombre + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","nombre",enc.EncodeForHtmlAttribute(replace(rst("nombre")&"","'","&#39;"))
				DrawCelda "CELDA style='width:100px'","","",0,LitNombreCom + ":"
				DrawInputCelda "CELDA maxlength='45'","","",40,0,"","nombrecom",enc.EncodeForHtmlAttribute(replace(rst("nombrecom")&"","'","&#39;"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA","","",0,LitAdmin + ":"
				DrawInputCelda "CELDA maxlength='100'","","",40,0,"","admin",enc.EncodeForHtmlAttribute(""&rst("administrador"))
				DrawCelda "CELDA","","",0,LitCargo + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","cargo",enc.EncodeForHtmlAttribute(""&rst("cargo"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA","","",0,LitDireccion + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","direccion",enc.EncodeForHtmlAttribute(""&rst("direccion"))
			CloseFila
			'mejico=d_lookup("gestion_folios","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")) & ""
             strselect="select gestion_folios from configuracion with(nolock) where nempresa=?"
            mejico= DLookupP1(strselect,session("ncliente")&"",adVarChar,5,session("dsn_cliente"))
			if mejico then
			    DrawFila color_blau%><td colspan='4'><table width="100%" bgcolor="<%=color_blau%>" border="0"><%
			    DrawFila color_blau
			        DrawCelda "CELDA colspan='6'","","",0,LitDireccionFacturaElectronica & ":"
			    CloseFila
			    DrawFila color_blau
			        DrawCelda "CELDA style='width:145px'","","",2,LitCalle & ":"
			        DrawInputCelda "CELDA maxlength='50'","29","",40,0,"","calle",enc.EncodeForHtmlAttribute(""&rst("calle"))
			        DrawCelda "CELDA ","","",0,LitNExterior & ":"
			        DrawInputCelda "CELDA maxlength='20'","","",20,0,"","nexterior",enc.EncodeForHtmlAttribute(""&rst("nexterior"))
			        DrawCelda "CELDA ","","",2,LitNInterior & ":"
			        DrawInputCelda "CELDA maxlength='20'","","",20,2,"","ninterior",enc.EncodeForHtmlAttribute(""&rst("ninterior"))
			    CloseFila
			    DrawFila color_blau
			        DrawCelda "CELDA style='width:145px'","","",2,LitColonia & ":"
			        DrawInputCelda "CELDA maxlength='50'","","",40,0,"","colonia",enc.EncodeForHtmlAttribute(""&rst("colonia"))
			        DrawCelda "CELDA","","",0,LitMunicipio & ":"
			        DrawInputCelda "CELDA maxlength='50' colspan='4'","","",40,0,"","municipio",enc.EncodeForHtmlAttribute(""&rst("municipio"))
			    CloseFila
			    %></table></td><%
			    CloseFila
		    end if
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitPoblacion + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","poblacion",enc.EncodeForHtmlAttribute(""&rst("poblacion"))
				DrawCelda "CELDA","","",0,LitProvincia + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","provincia",enc.EncodeForHtmlAttribute(""&rst("provincia"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitCp + ":"
				DrawInputCelda "CELDA maxlength='10'","","",10,0,"","cp",enc.EncodeForHtmlAttribute(""&rst("cp"))
				DrawCelda2 "CELDA", "left", false, LitPais  + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","pais",enc.EncodeForHtmlAttribute(""&rst("pais"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitTelefono + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","telefono",enc.EncodeForHtmlAttribute(""&rst("telefono"))
				DrawCelda "CELDA","","",0,LitTelefono2 + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","telefono2",enc.EncodeForHtmlAttribute(""&rst("telefono2"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitFax + ":"
				DrawInputCelda "CELDA maxlength='30'","","",30,0,"","fax",enc.EncodeForHtmlAttribute(""&rst("fax"))
				DrawCelda "CELDA","","",0,LitEmail + ":"
				DrawInputCelda "CELDA maxlength='50'","","",40,0,"","email",enc.EncodeForHtmlAttribute(""&rst("email"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:150px'","","",0,LitIRPF + ":"
				DrawInputCelda "CELDA maxlength='3'","","",5,0,"","IRPF",enc.EncodeForHtmlAttribute(""&rst("IRPF"))
                DrawCelda "CELDA style='width:150px'","","",0,LitRECC + ":"
               %>
            <td>
            <%    if rst("RECC")=-1 or rst("RECC")="True" then%>
						<input type="hidden" name="recc" value="1"/>
						<input type="checkbox" name="checkRecc" value="true" onclick="cambiar7();" checked="checked"/> <%
					else%>
						<input type="hidden" name="recc" value="0"/>
						<input type="checkbox" name="checkRecc" value="true" onclick="cambiar7();"/><%
					end if%>
                </td>
                
            <%
			CloseFila
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitImpresionLogo + ":"%>
				<td class="CELDA"><%
					if rst("print_logo")=-1 or rst("print_logo")="True" then%>
						<input type="checkbox" name="check2" value="true" onclick="cambiar2();" checked="checked"/>
						<input type="hidden" name="impresion" value="1"/> <%
					else%>
						<input type="hidden" name="impresion" value="0"/>
						<input type="checkbox" name="check2" value="true" onclick="cambiar2();"/><%
					end if
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class='CELDA' align='left'>" & LitImpresionEmpresa & ":</font>")
					if rst("print_empresa")=-1 or rst("print_empresa")="True" then%>
						<input type="checkbox" name="check4" value="true" onclick="cambiar4();" checked="checked"/>
						<input type="hidden" name="impresion_empresa" value="1"/> <%
					else%>
						<input type="hidden" name="impresion_empresa" value="0"/>
						<input type="checkbox" name="check4" value="true" onclick="cambiar4();"/><%
					end if%>
				</td>
				<%
				DrawCelda2 "CELDA", "left", false, LitImpresionTienda + ":"%>
				<td class="CELDA" ><%
					if rst("print_tienda")=-1 or rst("print_tienda")="True" then%>
						<input type="checkbox" name="check5" value="true" onclick="cambiar5();" checked="checked"/>
						<input type="hidden" name="impresion_tienda" value="1"/> <%
					else%>
						<input type="hidden" name="impresion_tienda" value="0"/>
						<input type="checkbox" name="check5" value="true" onclick="cambiar5();" <%=iif(rst("print_empresa")=-1 or rst("print_empresa")="True","","disabled")%>/><%
					end if
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class='CELDA' align='left'>" & LitControlStock & ":</font>")
					if rst("control_stock")=-1 or rst("control_stock")="True" then
						session("control_stock") = "activado"%>
						<input type="hidden" name="stock" value="1"/>
						<input type="checkbox" name="check3" value="true" onclick="cambiar3();" checked="checked"/> <%
					else
						session("control_stock") = "desactivado"%>
						<input type="hidden" name="stock" value="0"/>
						<input type="checkbox" name="check3" value="true" onclick="cambiar3();"/><%
					end if%>
					<%'FLM:20091123
					response.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class='CELDA' align='left'>" & LitLogoEmail & ":</font>")
					if rst("logo_email")=-1 or rst("logo_email")="True" then%>
						<input type="hidden" name="logo_email" value="1"/>
						<input type="checkbox" name="check6" value="true" onclick="cambiar6();" checked="checked"/> <%
					else%>
						<input type="hidden" name="logo_email" value="0"/>
						<input type="checkbox" name="check6" value="true" onclick="cambiar6();"/><%
					end if%>
				</td><%
			CloseFila%>
		</table>
		<table width="100%" bgcolor="<%=color_blau%>" border="0"><%
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitLeyenda + ":"
				DrawTextCelda "CELDA maxlength='250'","","",5,50,"","leyenda",enc.EncodeForHtmlAttribute(""&rst("leyenda"))
				DrawCelda "CELDA style='width:100px'","","",0,LitT1 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T1",enc.EncodeForHtmlAttribute(""&rst("campo01"))
				DrawCelda "CELDA style='width:100px'","","",0,LitT2 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T2",enc.EncodeForHtmlAttribute(""&rst("campo02"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitObjetoSocial + ":"
				DrawTextCelda "CELDA","","",5,50,"","objeto_social",enc.EncodeForHtmlAttribute(""&rst("objeto_social"))
				DrawCelda "CELDA style='width:100px'","","",0,LitT3 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T3",enc.EncodeForHtmlAttribute(""&rst("campo03"))
				DrawCelda "CELDA style='width:100px'","","",0,LitT4 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T4",enc.EncodeForHtmlAttribute(""&rst("campo04"))
			CloseFila
			DrawFila color_blau
				DrawCelda "CELDA style='width:100px'","","",0,LitPieMail + ":"
				DrawTextCelda "CELDA","","",5,50,"","pie_mail",enc.EncodeForHtmlAttribute(""&rst("pie_mail"))
				DrawCelda "CELDA style='width:100px'","","",0,LitT5 + ":"
				DrawTextCelda "CELDA maxlength='500'","","",5,50,"","T5",enc.EncodeForHtmlAttribute(""&rst("campo05"))
			CloseFila%>
		</table>
		<%rst.close
           conn.close
          set cmd=nothing
          set conn= nothing
	else
	end if%>
	</form>
<%
	set rst = nothing
	set rstAux = nothing
	set rstSelect = nothing
    set conn = nothing
    set rstCif = nothing

end if%>
</body>
</html>