<%@ Language=VBScript %>
<%
''''ricardo 31/7/2003 comprobamos que existe el proveedor que se ha pedido ver desde un listado, sino se va al modo add

''ricardo 1/9/2003 se modifica para que cuando el proveedor este dado de baja o no tenga el modulo de clientes,
'' no pueda convertirse en distribuidor
''''''''''

'JCI 09/07/2004 : Borrar bonificaciones y registros de proveer antes de eliminar el proveedor

'ilionUA 06/04/2012 : Nuevo Diseño
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
function pintar_saltos_nuevo(texto)
	texto=Replace(texto,"&#10;","")
	texto=Replace(texto,"&#13;","<br>")
	pintar_saltos_nuevo=texto
end function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>
<style type="text/css">
.el0
{
	display:inline-block;
	/*padding: 0px;*/
	/*margin: 0px -3px -3px 0px;*/
	/*margin: 0px -3px -3px 0px;*/
	float:left;
	/*border: 1px solid #ff0000;*/
}
</style>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<meta http-equiv="Content-style-Type" content="text/css"/>
<link rel="styleSHEET" href="../pantalla.css" media="SCREEN"/>
<link rel="styleSHEET" href="../impresora.css" media="PRINT"/>
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
<!--#include file="proveedores.inc" -->
<!--#include file="../common/googlemaps.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
    
<!--#include file="../tablasResponsive.inc" -->

  
<!--#include file="proveedores_linkextra.inc" -->
<!--#include file="../js/dropdown.js.inc" -->


<!--#include file="../js/calendar.inc" -->
    
<!--#include file="../common/proveedoresActionDrop.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/dropdown.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script type="text/javascript" language="javascript">
    animatedcollapse.addDiv('AddDG', 'fade=1')
    animatedcollapse.addDiv('AddDC', 'fade=1')
    animatedcollapse.addDiv('AddDB', 'fade=1')
    animatedcollapse.addDiv('AddCD', 'fade=1')
    animatedcollapse.addDiv('AddOD', 'fade=1')
    animatedcollapse.addDiv('AddDD', 'fade=1')
    animatedcollapse.addDiv('AddCP', 'fade=1')
    animatedcollapse.addDiv('BrowseDG', 'fade=1')
    animatedcollapse.addDiv('BrowseDC', 'fade=1')
    animatedcollapse.addDiv('BrowseDB', 'fade=1')
    animatedcollapse.addDiv('BrowseCD', 'fade=1')
    animatedcollapse.addDiv('BrowseOD', 'fade=1')
    animatedcollapse.addDiv('BrowseDD', 'fade=1')
    animatedcollapse.addDiv('BrowseCP', 'fade=1')
    animatedcollapse.addDiv('EditDG', 'fade=1')
    animatedcollapse.addDiv('EditDC', 'fade=1')
    animatedcollapse.addDiv('EditDB', 'fade=1')
    animatedcollapse.addDiv('EditCD', 'fade=1')
    animatedcollapse.addDiv('EditOD', 'fade=1')
    animatedcollapse.addDiv('EditDD', 'fade=1')
    animatedcollapse.addDiv('EditCP', 'fade=1')


    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //jQuery: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()

</script>

<script language="javascript" type="text/javascript">
<!--
//jcg 02/02/2008: necesaria para unificar con las demas paginas de ventas.
/*
function keyPressed(tecla)
{
	return;
}
*/
function abrir_acceso(nproveedor,viene,titulo,correo)
{
	if (correo=='') alert("<%=LitNoAccesoTienNoMail%>");
	else AbrirVentana("../central.asp?pag1=tiendas/accesos_tienda.asp&pag2=tiendas/accesos_tienda_bt.asp&ndoc=" + nproveedor + "&viene=" + viene + "&titulo=" + titulo,'P',<%=AltoVentana%>,<%=AnchoVentana%>);
}

function convprodist(nproveedor)
{
	if(confirm("<%=LitConvertirProDis%>")) document.location="proveedores.asp?mode=convertirprodist&nproveedor=" + nproveedor + "&salto=no";
}

function tier1Menu(objMenu,objImage)
{
    if (objMenu.style.display == "none")
    {
        objMenu.style.display = "";
        objImage.src = "../Images/<%=ImgCarpetaAbierta%>";
    }
    else
    {
        objMenu.style.display = "none";
		objImage.src = "../images/<%=ImgCarpetaCerrada%>";
    }
}

function Inicia()
{
   parent.document.location="default.htm";
}

function CopiarCampos()
{
	document.proveedores.de_domicilio.value=document.proveedores.domicilio.value;
	document.proveedores.de_poblacion.value=document.proveedores.poblacion.value;
	document.proveedores.de_cp.value=document.proveedores.cp.value;
	document.proveedores.de_provincia.value=document.proveedores.provincia.value;
	document.proveedores.de_pais.value=document.proveedores.pais.value;
	document.proveedores.de_telefono.value=document.proveedores.telefono.value;
	document.proveedores.de_telefono2.value=document.proveedores.telefono2.value;
	document.proveedores.de_fax.value=document.proveedores.fax.value;
}

function EliminarDirEnvio(nproveedor)
{
	if (confirm("<%=LitMsgElimDirEnvioConfirm%>"))
	{
		document.proveedores.action="proveedores.asp?mode=borrardirenvio&nproveedor=" + nproveedor;
		document.proveedores.submit();
	}
}

function comprobar()
{
	if (parseInt(document.proveedores.e_primer_ven.value)>31) document.proveedores.e_primer_ven.value=31;
	if (parseInt(document.proveedores.e_segundo_ven.value)>31) document.proveedores.e_segundo_ven.value=31;
	if (parseInt(document.proveedores.e_tercer_ven.value)>31) document.proveedores.e_tercer_ven.value=31;
	if ((parseInt(document.proveedores.e_segundo_ven.value)<= parseInt(document.proveedores.e_primer_ven.value))
	    && (parseInt(document.proveedores.e_segundo_ven.value)>0 )
	   )
	{
	   alert("<%=LitDiaPago1InfIgualDiaPago2%>");
	   document.proveedores.e_segundo_ven.focus();
	}
	else
	if ((parseInt(document.proveedores.e_tercer_ven.value)<= parseInt(document.proveedores.e_segundo_ven.value))
	     && (parseInt(document.proveedores.e_tercer_ven.value)>0 )
	   )
	{
	   alert("<%=LitDiaPago2InfIgualDiaPago3%>");
	   document.proveedores.e_tercer_ven.focus();
	}
}

function Editar(prov)
{
	document.proveedores.action="proveedores.asp?nproveedor=" + prov + "&mode=browse";
	document.proveedores.submit();
	//document.location="proveedores.asp?nproveedor=" + prov + "&mode=browse";
	parent.botones.document.location="proveedores_bt.asp?mode=browse&noadd=<%=enc.EncodeForJavascript(noadd)%>";
}

function Mas(sentido,lote,campo,criterio,texto)
{
	document.location="proveedores.asp?mode=search&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto;
}

function iva0()
{
    if(document.proveedores.intra.checked || document.proveedores.invsp.checked ) {
		document.proveedores.iva.value=0;
		document.proveedores.iva.disabled=true;
	}
	else document.proveedores.iva.disabled="";
}

//AMF:17/12/2010:Lleva el proveedor creado a la incidencia.
function llevarProveedorIncidencia(nproveedor) {
    eval("window.top.opener.parent.pantalla.document.incidencias.encargadopor.value=trimCodEmpresa(nproveedor);");

    window.top.opener.parent.pantalla.TraerProveedor("../mantenimiento/");
	window.close();
	return true;
}

//-->
</script>

<body onload="self.status='';" class="BODY_ASP">
<%''MPC 28/04/2009 Se añade el parámetro de usuario cifrepe
dim cifrepe
''DBS 07/10/2014 obtenermos el parametro de usuario
dim llekoAdmin
ObtenerParametros("proveedores")
'cag
	si_tiene_modulo_21=ModuloContratado(session("ncliente"),"21")
	si_tiene_modulo_22=ModuloContratado(session("ncliente"),"22")
	si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
'fin cag
    si_tiene_modulo_proyectos = ModuloContratado(session("ncliente"),ModProyectos) 'jcg
    si_tiene_modulo_OrCU=ModuloContratado(session("ncliente"),ModOrCU)
    si_tiene_modulo_SANTOS=ModuloContratado(session("ncliente"),ModCentroxogo)

    set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
'response.Write "parametro1:"&llekoAdmin
'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
function GuardarRegistro(nproveedor)
 if (CInt(request.form("e_primer_ven")) >= CInt(request.form("e_segundo_ven")) and CInt(request.form("e_segundo_ven"))>0) or _
     (CInt(request.form("e_segundo_ven")) >= CInt(request.form("e_tercer_ven")) and CInt(request.form("e_tercer_ven"))>0) then%>
    <script language="javascript" type="text/javascript">
        window.alert("<%=LitMsgDiasMal%>");
        history.back();
        history.back();
        parent.botones.document.location = "proveedores_bt.asp?mode=edit&noadd=<%=enc.EncodeForJavascript(noadd)%>";
   </script>
 <%else
	errorp= "no"
	crear_proveedor=1
	if Request.Form("cif")>"" then
		CIF=LimpiarCIF(Request.form("cif"))
		if CIF>"" then
            rstAux.cursorlocation=3
            if nproveedor & "">"" then
                'strselect = "select nproveedor from proveedores where nproveedor<>'" & nproveedor & "' and cifedi='" & CIF & "' and nproveedor like '" & session("ncliente") & "%'", session("dsn_cliente")
                strselect = "select nproveedor from proveedores where nproveedor<>? and cifedi=? and nproveedor like ?+'%'"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
                command2.Parameters.Append command2.CreateParameter("@cifedi",adVarChar,adParamInput,50,CIF&"")
                command2.Parameters.Append command2.CreateParameter("@nproveedor1",adVarChar,adParamInput,10,session("ncliente")&"")
                rstAux.CursorLocation = adUseClient
                rstAux.Open command2, , adOpenKeyset, adLockOptimistic
            else
                'strselect = "select nproveedor from proveedores where cifedi='" & CIF & "' and nproveedor like '" & session("ncliente") & "%'", session("dsn_cliente")
                strselect = "select nproveedor from proveedores where cifedi=? and nproveedor like ?+'%'"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@cifedi",adVarChar,adParamInput,50,CIF&"")
                command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,session("ncliente")&"")
                rstAux.CursorLocation = adUseClient
                rstAux.Open command2, , adOpenKeyset, adLockOptimistic
            end if
            '' MPC 08/10/2008 Si tiene el parámetro cifrepe y a la pregunta responde que si entonces inserta el proveedor aunque exista el cif
			if not rstAux.EOF and cstr(repe)<>"1" then
				if nproveedor="" then
					submode="add"
				else
					submode="edit"
				end if
				errorp = "si"
				crear_proveedor=0
                conn2.close
                set conn2    =  nothing
                set command2 =  nothing
                %>
				<script language="javascript" type="text/javascript">
				    alert("<%=LitCIFExistente%>");
				    parent.botones.document.location = "proveedores_bt.asp?mode=<%=enc.EncodeForJavascript(submode)%>&noadd=<%=enc.EncodeForJavascript(noadd)%>";
				    document.location = "proveedores.asp?mode=<%=enc.EncodeForJavascript(submode)%>&noadd=<%=enc.EncodeForJavascript(noadd)%>&repe=<%=enc.EncodeForJavascript(repe)%>&nproveedor=<%=enc.EncodeForJavascript(nproveedor)%>";
				</script>
			<%end if
			if rstAux.state <> 0 then rstAux.Close
		else
			if nproveedor="" then
				submode="add"
			else
				submode="edit"
			end if
			errorp = "si"
			crear_proveedor=0%>
			<script language="javascript" type="text/javascript">
			    alert("<%=LitCIFIncorrecto%>");
			    history.back();
			    parent.botones.document.location = "proveedores_bt.asp?mode=<%=enc.EncodeForJavascript(submode)%>&noadd=<%=enc.EncodeForJavascript(noadd)%>";
			</script>
		<%end if
	end if

	if errorp ="no" then
		if nproveedor="" then
			'Obtener el último nº de proveedores de CONFIGURACION.
			'rstAux.Open "select nproveedor from configuracion where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            strselect = "select nproveedor from configuracion where nempresa=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente")&"")
            rstAux.CursorLocation = adUseClient
            rstAux.Open command2, , adOpenKeyset, adLockOptimistic

			if not rstAux.eof then
				num=rstAux("nproveedor")+1
			else
				rstAux.addnew
				num=1
			end if
			num=string(5-len(cstr(num)),"0") + cstr(num)

			'Actualizar el nº de proveedor de CONFIGURACION.
			rstAux("nproveedor")=rstAux("nproveedor")+1
			rstAux.Update
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing


			''ricardo 10-1-2005 se comprobara que no existe el nproveedor segun el contador de datos de configuracion
			nproveedor_a_buscar=session("ncliente") & num
			rstAux.cursorlocation=3
			'rstAux.open "select nproveedor from proveedores with(nolock) where nproveedor like '"&session("ncliente")&"%' and nproveedor='" & nproveedor_a_buscar & "'",session("dsn_cliente")

            strselect = "select nproveedor from proveedores with(nolock) where nproveedor like ?+'%' and nproveedor=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@nproveedor1",adVarChar,adParamInput,10,nproveedor_a_buscar&"")
            rstAux.CursorLocation = adUseClient
            rstAux.Open command2, , adOpenKeyset, adLockOptimistic

			if not rstAux.eof then
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing

				'rstAux.open "update configuracion with(updlock) set nproveedor=nproveedor-1 where nempresa='" & session("ncliente") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

                strselect = "update configuracion with(updlock) set nproveedor=nproveedor-1 where nempresa=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente")&"")
                rstAux.CursorLocation = adUseClient
                rstAux.Open command2, , adOpenKeyset, adLockOptimistic

				crear_proveedor=0%>
				<script language="javascript" type="text/javascript">
				    window.alert("<%=LitMsgProvExistRevCont%>");
				    history.back();
				    parent.botones.document.location = "proveedores_bt.asp?mode=add&noadd=<%=enc.EncodeForJavascript(noadd)%>"
				</script>
			<%else

			end if
            rstAux.close
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
      
			if crear_proveedor=1 then
				'Crear un nuevo registro.
				rst.AddNew
				rst("nproveedor")=session("ncliente") & num
			end if
		end if

		if crear_proveedor=1 then
			'Datos del domicilio de los DATOS GENERALES y los DATOS DE ENVIO
			if nproveedor="" then 'MODO AÑADIR NUEVO
				'Abrimos la tabla de domicilios y creamos un registro nuevo para proveedor

				'rstDomi.Open "select * from domicilios where pertenece='" + rst("nproveedor") +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

                strselect = "select * from domicilios where pertenece=?" 
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@pertenece",adVarChar,adParamInput,55,rst("nproveedor")&"")
                rstDomi.CursorLocation = adUseClient
                rstDomi.Open command2, , adOpenKeyset, adLockOptimistic

				rstDomi.AddNew
				rstDomi("pertenece") = rst("nproveedor")
				rstDomi("tipo_domicilio") = "PRINCIPAL_PROV"
				rstDomi("domicilio") = Nulear(left(request.form("domicilio"),100))
				'rstDomi("cp")        = Nulear(request.form("cp"))
				if Request.Form("poblacion")>"" and Request.Form("cp")="" then
                    strselect = "select cod_postal from poblaciones with(Nolock) where poblacion=?"
                    cp_aux=DLookupP1(strselect,request.form("poblacion")&"",adVarChar,50,DsnIlion)
					if cp_aux="00000" then cp_aux=""
					rstDomi("cp")     		=Nulear(cp_aux)
				else
					rstDomi("cp")     		=Request.Form("cp")
				end if
				TmpCP=rstDomi("cp") 'este valor se utiliza para la creacion del contacto
				rstDomi("poblacion") = Nulear(request.form("poblacion"))
				rstDomi("provincia") = Nulear(request.form("provincia"))
				rstDomi("pais")      = Nulear(request.form("pais"))
				rstDomi("telefono")  = Nulear(request.form("telefono"))
				rstDomi.Update

				rst("dir_principal") = rstDomi("codigo")
                conn2.close
                set conn2    =  nothing
                set command2 =  nothing


				'Guardamos la direccion de delegacion caso de que exista
				if request.form("del_domicilio") > "" then
					'Abrimos la tabla de domicilios y creamos un registro nuevo
					'rstDomi.Open "select * from domicilios where pertenece='" + rst("nproveedor") +"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

                    strselect = "select * from domicilios where pertenece=?"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@pertenece",adVarChar,adParamInput,55,rst("nproveedor")&"")
                    rstDomi.CursorLocation = adUseClient
                    rstDomi.Open command2, , adOpenKeyset, adLockOptimistic

					rstDomi.AddNew
					rstDomi("pertenece") = rst("nproveedor")
					rstDomi("tipo_domicilio") = "ENVIO_PROV"
					rstDomi("domicilio") = Nulear(request.form("del_domicilio"))
					rstDomi("cp")        = Nulear(request.form("del_cp"))
					if Request.Form("del_poblacion")>"" and Request.Form("del_cp")="" then
						   strselect = "select cod_postal from poblaciones with(nolock) where poblacion=?"
                           cp_aux=DLookupP1(strselect,request.form("poblacion")&"",adVarChar,50,DsnIlion)
						if cp_aux="00000" then cp_aux=""
						rstDomi("cp")     		=Nulear(cp_aux)
					else
						rstDomi("cp")     		=Request.Form("del_cp")
					end if
					rstDomi("poblacion") = Nulear(request.form("del_poblacion"))
					rstDomi("provincia") = Nulear(request.form("del_provincia"))
					rstDomi("pais")      = Nulear(request.form("del_pais"))
					rstDomi("telefono")  = Nulear(request.form("del_telefono"))
					rstDomi("a_la_atencion")  = Nulear(request.form("del_contacto"))
					rstDomi.Update
					rst("dir_envio") = rstDomi("codigo")
                    conn2.close
                    set conn2    =  nothing
                    set command2 =  nothing
				end if
			else 'MODO EDITAR
				'cogemos el anterior DirPrincipal para poder modificar el del contacto
				TmpDirPrincipal_ant=rst("dir_principal")

				'Abrimos la tabla de domicilios y modificamos el registro para PROVEEDOR
				'Seleccion="SELECT * FROM domicilios WHERE codigo ='"+cstr(null_z(rst("dir_principal")))+"'"
				'rstDomi.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

                strselect = "SELECT * FROM domicilios WHERE codigo =?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60 
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@codigo",adInteger,adParamInput,10,cstr(null_z(rst("dir_principal")))&"")
                rstDomi.CursorLocation = adUseClient
                rstDomi.Open command2, , adOpenKeyset, adLockOptimistic

				if not (rstDomi("domicilio")&"" = request.form("domicilio")&"" and _
					rstDomi("cp")&""        = request.form("cp")&"" and _
					rstDomi("poblacion")&"" = request.form("poblacion")&"" and _
					rstDomi("provincia")&"" = request.form("provincia")&"" and _
					rstDomi("pais")&""      = request.form("pais")&"" and _
					rstDomi("telefono")&""  = request.form("telefono")&"") then
						rstDomi.AddNew
						rstDomi("pertenece") = rst("nproveedor")
						rstDomi("tipo_domicilio") = "PRINCIPAL_PROV"
						rstDomi("domicilio") = Nulear(left(request.form("domicilio"),100))
						'rstDomi("cp")        = Nulear(request.form("cp"))
						if Request.Form("poblacion")>"" and Request.Form("cp")="" then
						        strselect = "select cod_postal from poblaciones with(nolock) where poblacion=?"
                                cp_aux=DLookupP1(strselect,request.form("poblacion")&"",adVarChar,50,DsnIlion)
							if cp_aux="00000" then cp_aux=""
							rstDomi("cp")     		=Nulear(cp_aux)
						else
							rstDomi("cp")     		=Request.Form("cp")
						end if

					TmpCP_Nuevo=rstDomi("cp")
					rstDomi("poblacion") = Nulear(request.form("poblacion"))
					rstDomi("provincia") = Nulear(request.form("provincia"))
					rstDomi("pais")      = Nulear(request.form("pais"))
					rstDomi("telefono")  = Nulear(request.form("telefono"))

					rstDomi.Update
					rst("dir_principal") = rstDomi("codigo")
				end if
                conn2.close
                set conn2    =  nothing
                set command2 =  nothing
				'Modificamos la direccion de envío caso de que exista
				if request.form("del_domicilio") > "" then
					'Abrimos la tabla de domicilios y modificamos el registro para ENVIO
					nuevo = "false"

					'Seleccion="SELECT * FROM domicilios WHERE codigo ='" +cstr(null_z(rst("dir_envio"))) + "'"
					'rstDomi.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

                    strselect = "SELECT * FROM domicilios WHERE codigo=?"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@codigo",adInteger,adParamInput,10,cstr(null_z(rst("dir_envio")))&"")
                    rstDomi.CursorLocation = adUseClient
                    rstDomi.Open command2, , adOpenKeyset, adLockOptimistic


					if not rstDomi.EOF then
						if not (rstDomi("domicilio")&"" = request.form("del_domicilio")&"" and _
							rstDomi("cp")&""        = request.form("del_cp")&"" and _
							rstDomi("poblacion")&"" = request.form("del_poblacion")&"" and _
							rstDomi("provincia")&"" = request.form("del_provincia")&"" and _
							rstDomi("pais")&""      = request.form("del_pais")&"" and _
							rstDomi("a_la_atencion")&""      = request.form("del_contacto")&"" and _
							rstDomi("telefono")&""  = request.form("del_telefono")&"") then
								nuevo = "true"
						end if
					else
						nuevo = "true"
					end if
					if nuevo = "true" then
						rstDomi.AddNew
						rstDomi("pertenece") = rst("nproveedor")
						rstDomi("tipo_domicilio") = "ENVIO_PROV"
						rstDomi("domicilio") = Nulear(request.form("del_domicilio"))
						'rstDomi("cp")        = Nulear(request.form("del_cp"))
						if Request.Form("del_poblacion")>"" and Request.Form("del_cp")="" then
						        strselect = "select cod_postal from poblaciones with(nolock) where poblacion=?"
                                cp_aux=DLookupP1(strselect,request.form("poblacion")&"",adVarChar,50,DsnIlion)
							if cp_aux="00000" then cp_aux=""
							rstDomi("cp")     		=Nulear(cp_aux)
						else
							rstDomi("cp")     		=Request.Form("del_cp")
						end if
						rstDomi("poblacion") = Nulear(request.form("del_poblacion"))
						rstDomi("provincia") = Nulear(request.form("del_provincia"))
						rstDomi("pais")      = Nulear(request.form("del_pais"))
						rstDomi("telefono")  = Nulear(request.form("del_telefono"))
						rstDomi("a_la_atencion")  = Nulear(request.form("del_contacto"))
						rstDomi.Update
						rst("dir_envio") = rstDomi("codigo")
					end if
                    conn2.close
                    set conn2    =  nothing
                    set command2 =  nothing
				end if
			end if

			'Asignar los nuevos valores a los campos del recordset.
			'DATOS GENERALES
			rst("cif")           = Nulear(Request.Form("cif"))
			rst("cifedi")        = CIF

			'cogemos el anterior razon_social para poder modificar el del contacto
			TmpRazon_Social_Ant=rst("razon_social")

			rst("razon_social")  = Nulear(Request.Form("razon_social"))
			rst("falta")   		 = Nulear(request.form("falta"))
			rst("fbaja")  		 = Nulear(request.form("fbaja"))
			rst("nombre")        = Nulear(Request.Form("nombre"))

			'cogemos el anterior contacto para poder modificar el del contacto
			TmpContacto_Ant=rst("contacto")

			rst("contacto")      = Nulear(request.form("contacto"))
			rst("web")           = Nulear(request.form("web"))

			'cogemos el anterior email para poder modificar el del contacto
			TmpEmail_Ant=rst("email")

			rst("email")         = Nulear(request.form("email"))
			rst("observaciones") = Nulear(request.form("observaciones"))

            if si_tiene_modulo_OrCU <> 0 then
                rst("cae") = Nulear(request.form("cae"))
            end if

			'cogemos el anterior movil para poder modificar el del contacto
			TmpMovil_Ant=rst("telefono2")

			rst("telefono2")     = Nulear(request.form("telefono2"))

			'cogemos el anterior fax para poder modificar el del contacto
			TmpFax_Ant=rst("fax")

			rst("fax")           = Nulear(request.form("fax"))
			'DATOS COMERCIALES
			rst("descuento")     = miround(replace(Null_z(request.form("descuento")),",","."), decpor)
			rst("descuento2")    = miround(replace(Null_z(request.form("descuento2")),",","."), decpor)
			rst("forma_pago")    = Nulear(request.form("forma_pago"))
			rst("tipo_pago")     = Nulear(request.form("tipo_pago"))
			rst("primer_ven")    = Nulear(request.form("e_primer_ven"))
			rst("segundo_ven")   = Nulear(request.form("e_segundo_ven"))
			rst("tercer_ven")    = Nulear(request.form("e_tercer_ven"))
			rst("recargo")       = miround(replace(Null_z(request.form("recargo")),",","."), decpor)
			rst("re")            = Nz_b(request.form("re"))
			rst("cuenta_contable") = Nulear(request.form("cuenta_contable"))
			rst("ccontable_efecto") = Nulear(request.form("cuenta_contable_efecto"))
            strselect = "select codigo from divisas with(nolock) where moneda_base <> 0 and codigo like ?+'%'"
			rst("divisa") 		= iif(request.form("divisa")>"",Nulear(request.form("divisa")),DLookupP1(strselect,session("ncliente")&"",adVarChar,15,session("dsn_cliente")))
			rst("IRPF")	 		= miround(replace(Null_z(request.form("IRPF")),",","."), decpor)
			rst("IRPF_Total")	= Nz_b(request.form("IRPF_Total"))
			''rst("exento_iva")	= Nz_b(request.form("exento_iva"))
			'DATOS BANCARIOS
			'FLM:130309:Tomo el nombre del banco de la tabla bancos.
            ''response.write("la cuentabanco es-" & request.form("ncuenta") & "-" & left(trim(request.form("ncuenta")),4) & "-<br>")
            ncuenta_nom=""
            if request.form("ncuenta") & "">"" then
                ncuenta_nom=trim(request.form("ncuenta"))
            else
                ncuenta_nom=cuenta
            end if
                strselect = "select entidad from bancos with(nolock) where codigo=?"
			    rst("banco")	     = Nulear(left(DLookupP1(strselect,left(ncuenta_nom,4)&"",adVarChar,50,DsnIlion), 40))
			    rst("banco_dom")     = Nulear(request.form("banco_dom"))
                p_pais=Request.Form("country")
                ''response.write("la p_pais es-" & p_pais & "-<br>")
            if p_pais & ""="" and isnumeric(cuenta) then
                p_pais="ES"
                existe_pais=1
            else
                existe_pais=0
                set rstPais = server.CreateObject("ADODB.Recordset")
                rstPais.cursorlocation=3
                strselect = "select iso2 from paises with(NOLOCK) where iso2=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = dsnilion
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@iso2",adVarChar,adParamInput,2,p_pais&"")
                set rstPais = command2.Execute
                if not rstPais.eof then
                    existe_pais=1
                end if
				conn2.Close
                set conn2 = nothing
                set command2 = nothing
                set rstPais=nothing
            end if
            

            if existe_pais=1 and cuenta & "" <> "" then
                p_iban=Request.Form("IBAN")

                if p_iban & ""="" then
                    set conn=  server.CreateObject("ADODB.Connection")
                    set cmd=  server.CreateObject("ADODB.Command")
                    conn.open session("dsn_cliente")
                    cmd.ActiveConnection=conn
                    conn.cursorlocation=3
                    cmd.CommandText="CalculateIBAN"
	                cmd.CommandType = adCmdStoredProc 
                    cmd.Parameters.Append cmd.CreateParameter("@banknumber", adVarChar, ,20,cuenta)
                    cmd.Parameters.Append cmd.CreateParameter("@country", adVarChar, ,20,p_pais)
                    set rstGetIban=cmd.execute
                    if not rstGetIban.eof then
                        p_iban=rstGetIban("result")
                    end if
                    conn.close
                    if rstGetIban.state<>0 then rstGetIban.close
                    set rstGetIban=nothing
                    set cmd=nothing
                    set conn=nothing
                end if
                ''response.write("el p_iban es-" & p_iban & "-" & cuenta & "-<br>")
                entidadSW=""
                
                if cuenta & "">"" then
                    entidadSW=left(cuenta,4)
                end if
                ncuenta=p_pais & p_iban & cuenta

                
                if swift & ""="" then
                    set conn=  server.CreateObject("ADODB.Connection")
                    set cmd=  server.CreateObject("ADODB.Command")
                    conn.open dsnilion
                    cmd.ActiveConnection=conn
                    conn.cursorlocation=3
                    cmd.CommandText="GetDefaultDataBank"
	                cmd.CommandType = adCmdStoredProc 
                    cmd.Parameters.Append cmd.CreateParameter("@bankCode", adVarChar, , 4,entidadSW)
                    set rstGetSwift=cmd.execute
                    if not rstGetSwift.eof then
                        swift=rstGetSwift("swift_code")
                    end if
                    conn.close
                    if rstGetSwift.state<>0 then rstGetSwift.close
                    set rstGetSwift=nothing
                    set cmd=nothing
                    set conn=nothing
                end if
            else
                ncuenta = request.form("ncuenta") 
            end if
            
            
			rst("ncuenta")       = Nulear(ncuenta)
			rst("cuenta_cargo")  = Nulear(request.form("ncuentacargo"))
            if swift & "" <> "" then
                rst("swift_code")       = Nulear(swift)
            else
                rst("swift_code")       = Nulear(request.Form("bic"))
            end if
			'FLM:130309: domiciliación bancaria.
			rst("domrec")        = Nz_b(request.form("Domiciliacion"))
			'OTROS DATOS
			rst("tactividad")     = Nulear(request.form("tactividad"))
			rst("transportista")  = Nulear(request.form("transportista"))
			rst("proyecto")       = Nulear(request.form("cod_proyecto")) 'jcg
			rst("portes")         = Nulear(request.form("portes"))
			rst("tipo_proveedor") = Nulear(request.form("tipo_proveedor"))
			rst("iva")=			nulear(request.form("iva"))

			'**RGU 13/6/2006**
			if request.form("intra") = "on" then
				rst("intra") = 1
				rst("iva")=0
			else
				rst("intra") = 0
			end if
			'**RGU**
            'dgb
            if request.form("recc") = "on" then
				rst("recc") = 1
			else
				rst("recc") = 0
			end if
            if nz_b2(request.form("invsp")) =1 then
                rst("INVSUJETOPASIVO")=1
                rst("iva")=0
            else
                rst("INVSUJETOPASIVO")=0
            end if
            rst("CCONTABLE_PAGOVENC")=Nulear(request.form("CCONTABLE_PAGOVENC"))
			'ricardo 20-5-2004 actualizamos los campos personalizables
			num_campos=limpiaCadena(request.querystring("num_campos"))
			if num_campos="" then
				num_campos=limpiaCadena(request.form("num_campos"))
			end if
			if num_campos & "">"" then
				redim lista_valores(num_campos+5)
				for ki=1 to num_campos
					nom_campo="campo" & ki
					valor_form=Nulear(limpiaCadena(request.querystring(nom_campo)))
					if valor_form & ""="" then
						valor_form=Nulear(limpiaCadena(request.form(nom_campo)))
					end if
					tipo_campo_perso=request.form("tipo_campo" & ki)
					if tipo_campo_perso=2 then
						if valor_form="on" then
							lista_valores(ki)=1
						else
							lista_valores(ki)=0
						end if
					else
						lista_valores(ki)=valor_form
					end if
				next
			else
				redim lista_valores(10+5)
				for ki=1 to 15
					lista_valores(ki)=""
				next
			end if
			rst("campo01")=lista_valores(1)
			rst("campo02")=lista_valores(2)
			rst("campo03")=lista_valores(3)
			rst("campo04")=lista_valores(4)
			rst("campo05")=lista_valores(5)
			rst("campo06")=lista_valores(6)
			rst("campo07")=lista_valores(7)
			rst("campo08")=lista_valores(8)
			rst("campo09")=lista_valores(9)
			rst("campo10")=lista_valores(10)

			'Actualizar el registro.
			rst.Update
			'mmg
			'PONEMOS LOS DATOS EN DOCUMENTO_PRO
		    if ncliente & ""="" then
			    ncliente_doc_cli=rst("nproveedor")
		    else
			    ncliente_doc_cli=ncliente
		    end if
            
		    'cadena="select * from documentos_pro where ncliente = '" & session("ncliente") & "' and nproveedor='" & rst("nproveedor") & "'"
		    'rstAux.open cadena,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
            
            strselect = "select * from documentos_pro where ncliente=? and nproveedor=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,rst("nproveedor")&"")
            rstAux.CursorLocation = adUseClient
            rstAux.Open command2, , adOpenKeyset, adLockOptimistic

		    if rstAux.eof then
			    rstAux.addnew
		        rstAux("nproveedor")= rst("nproveedor")
		        rstAux("ncliente")=session("ncliente")
		    end if
            rstAux("valorado_ped")=nz_b(request.form("valorado_ped"))
            rstAux("valorado_alb")=nz_b(request.form("valorado_alb"))
            rstAux.update
            strupdate = "update documentos_pro with(rowlock) set serie_ped=?, serie_alb=?, serie_fac=? where ncliente=? and nproveedor=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strupdate
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@serie_ped",adVarChar,adParamInput,10,iif(request.form("serie_ped")>"",request.form("serie_ped"),null))
            command2.Parameters.Append command2.CreateParameter("@serie_alb",adVarChar,adParamInput,10,iif(request.form("serie_alb")&""<>"",request.form("serie_alb"),null))
            command2.Parameters.Append command2.CreateParameter("@serie_fac",adVarChar,adParamInput,10,iif(request.form("serie_fac")&""<>"",request.form("serie_fac"),null))
            command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,rst("nproveedor")&"")
            set rstAux = command2.Execute

		    'rstAux("valorado_ped")=nz_b(request.form("valorado_ped"))
            'if request.form("serie_ped")&"">"" then
            '    rstAux("serie_ped") = request.form("serie_ped")
            'end if
		    'rstAux("serie_ped")=iif(request.form("serie_ped")>"",request.form("serie_ped"),null)
		    'rstAux("valorado_alb")=nz_b(request.form("valorado_alb"))
		    'rstAux("serie_alb")=iif(request.form("serie_alb")&""<>"",request.form("serie_alb"),null)
		    'rstAux("serie_fac")=iif(request.form("serie_fac")&""<>"",request.form("serie_fac"),null)

		    'rstAux.update
		    conn2.close
            set conn2    =  nothing
            set command2 =  nothing
		    rstAux.cursorlocation=2
			if nproveedor="" then 'MODO AÑADIR NUEVO
				'creamos un contacto, en caso de que exista contacto o email
				if request.form("contacto")>"" or request.form("email")>"" then
					TmpNContacto=d_max("substring(ncontacto,6,10)","contactos_pro","ncontacto like '" & session("ncliente") & "%'",session("dsn_cliente"))+1
					TmpNContacto=session("ncliente") & completar(TmpNContacto,5,"0")
					strselect="select * from domicilios where tipo_domicilio='CONTACTO_PRO' and pertenece=?"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@pertenece",adVarChar,adParamInput,55,TmpNContacto&"")
                    rstAux.CursorLocation = adUseClient
                    rstAux.Open command2, ,adOpenKeyset, adLockOptimistic
					if rstAux.eof then
						rstAux.AddNew
						' el codigo no se pone, ya que esta automatico en el diseño de la tabla
						'num_dom=d_max("codigo","domicilios","",session("dsn_cliente"))+1
						'rstAux("codigo")=cint(num_dom)
						rstAux("pertenece")=TmpNContacto
						rstAux("tipo_domicilio")="CONTACTO_PRO"
					end if
					rstAux("domicilio")=Nulear(left(request.form("domicilio"),100))
					rstAux("CP")=TmpCP
					rstAux("poblacion")=Nulear(request.form("poblacion"))
					rstAux("provincia")=Nulear(request.form("provincia"))
					'rstAux("pais")=
					rstAux("telefono")=Nulear(request.form("telefono"))
					'rstAux("a_la_atencion")=
					rstAux.update
					num_dom=rstAux("codigo")
					conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
					'ahora grabamos el contacto
					strselect="select * from contactos_pro where ncontacto=?"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strselect
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@ncontacto",adVarChar,adParamInput,10,session("ncliente")&"")
                    rstAux.CursorLocation = adUseClient
                    rstAux.Open command2, ,adOpenKeyset, adLockOptimistic
					if rstAux.eof then
						rstAux.AddNew
						rstAux("ncontacto")=TmpNContacto
						rstAux("nproveedor")=rst("nproveedor")
						rstAux("domicilio")=num_dom
						if request.form("contacto")>"" then
							rstAux("nombre")=Nulear(request.form("contacto"))
						else
							rstAux("nombre")=Nulear(Request.Form("razon_social"))
						end if
					end if
					rstAux("cargo")=""
					'if TmpDepartamento & "">"" then
					'	rstAux("departamento")=TmpDepartamento
					'end if
					rstAux("movil")=Nulear(left(request.form("telefono2"),15))
					rstAux("fax")=Nulear(left(request.form("fax"),15))
					rstAux("mail")=left(Nulear(request.form("email")),50)
					rstAux.update
					conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
				end if
			else 'modo modificar direccion de los conctactos que tengan los mismos datos
                strselect = "select ncontacto,nombre,mail,movil,fax from contactos_pro where nproveedor=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
                rstAux2.CursorLocation = adUseClient
                rstAux2.Open command2, ,adOpenKeyset, adLockOptimistic
				if not rstAux2.eof then
					'cogemos los datos de la direccion del proveedor anterior al cambio
					strselect="select * from domicilios where codigo=?"
                    set command = nothing
                    set conn = Server.CreateObject("ADODB.Connection")
                    set command = Server.CreateObject("ADODB.Command")
                    conn.Open = session("dsn_cliente")
                    conn.CursorLocation = 3
                    command.ActiveConnection = conn
                    command.CommandTimeout = 60
                    command.CommandText = strselect
                    command.CommandType = adCmdText
                    command.Parameters.Append command.CreateParameter("@codigo",adInteger,adParamInput,4,TmpDirPrincipal_ant&"")
                    set rstAux = command.Execute
					if not rstAux.eof then
						TmpDomicilio_Ant=rstAux("domicilio")
						TmpCP_Ant=rstAux("cp")
						TmpPoblacion_Ant=rstAux("poblacion")
						TmpProvincia_Ant=rstAux("provincia")
						TmpTelefono_Ant=rstAux("telefono")
					end if
				    conn.Close
                    set conn = nothing
                    set command = nothing

					while not rstAux2.eof
                        strselect = "select * from domicilios where tipo_domicilio='CONTACTO_PRO' and pertenece=?"
                        set command = nothing
                        set conn = Server.CreateObject("ADODB.Connection")
                        set command = Server.CreateObject("ADODB.Command")
                        conn.Open = session("dsn_cliente")
                        conn.CursorLocation = 3
                        command.ActiveConnection = conn
                        command.CommandTimeout = 60
                        command.CommandText = strselect
                        command.CommandType = adCmdText
                        command.Parameters.Append command.CreateParameter("@pertenece",adVarChar,adParamInput,55,rstAux2("ncontacto")&"")
                        rstAux.CursorLocation = adUseClient
                        rstAux.Open command, ,adOpenKeyset, adLockOptimistic
						if not rstAux.eof then
							if null_s(TmpDomicilio_Ant)=null_s(rstAux("domicilio")) and null_s(TmpCP_Ant)=null_s(rstAux("cp")) and null_s(TmpPoblacion_Ant)=null_s(rstAux("poblacion")) and null_s(TmpProvincia_Ant)=null_s(rstAux("provincia")) and null_s(TmpTelefono_Ant)=null_s(rstAux("telefono")) then
								rstAux("domicilio")=Nulear(request.form("domicilio"))
								rstAux("cp")=TmpCP_Nuevo
								rstAux("poblacion")=Nulear(request.form("poblacion"))
								rstAux("provincia")=Nulear(request.form("provincia"))
								rstAux("telefono")=Nulear(request.form("telefono"))
								rstAux.update
							end if
						end if
				        conn.Close
                        set conn = nothing
                        set command = nothing

						if null_s(TmpRazon_Social_Ant)=null_s(rstAux2("nombre")) then
							rstAux2("nombre")=Nulear(Request.Form("razon_social"))
						end if
						if null_s(TmpContacto_Ant)=null_s(rstAux2("nombre")) then
							rstAux2("nombre")=Nulear(request.form("contacto"))
						end if
						if null_s(TmpEmail_Ant)=null_s(rstAux2("mail")) then
							rstAux2("mail")=left(Nulear(request.form("email")),50)
						end if
						if null_s(TmpMovil_Ant)=null_s(rstAux2("movil")) then
							rstAux2("movil")=Nulear(request.form("telefono2"))
						end if
						if null_s(TmpFax_Ant)=null_s(rstAux2("fax")) then
							rstAux2("fax")=Nulear(request.form("fax"))
						end if

						rstAux2.update
						rstAux2.movenext
					wend
				end if
			end if
			npro=rst("nproveedor")
		end if
	end if

	if viene>"" then
		if viene="subcuentas" then
			pagina="subcuentas"
		end if
        rstAux.cursorlocation=3
        strselect = "select * from proveedores with(nolock) where nproveedor like ?+'%' and nproveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,session("ncliente")&"")
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,npro&"")
        set rstAux = command2.Execute
		if not rstAux.eof then
		    'AMF:17/12/2010:Comprobamos que no venga de incidencias porque provoca error.
		    if viene<>"incidencias" then
			    %><script language="javascript" type="text/javascript">
			            alert("PASO");
				        window.top.opener.parent.pantalla.document.subcuentas.nproveedor.value="<%=enc.EncodeForJavascript(npro)%>";
                        window.top.opener.parent.pantalla.fr_Proveedor.document.docproveedor.nproveedor.value ="<%=enc.EncodeForJavascript(trimCodEmpresa(npro))%>";
                        window.top.opener.parent.pantalla.fr_Proveedor.document.docproveedor.nom_proveedor.value ="<%=enc.EncodeForJavascript(rstAux("razon_social"))%>";
			    </script>
			    <%
			end if

			'jcg 02/02/2008
            if si_tiene_modulo_proyectos<>0 then%>
		        <script language="javascript" type="text/javascript">
                        window.top.opener.parent.pantalla.document.<%=enc.EncodeForJavascript(viene) %>.cod_proyecto.value="<%=enc.EncodeForJavascript(rst("proyecto"))%>";
                </script>
            <%end if%>
			<script language="javascript" type="text/javascript">
			    parent.window.close();
			</script><%
		else
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitNoPuedeCrearCliente%>");
			      parent.window.close();
			</script><%
		end if
		conn2.Close
        set conn2 = nothing
        set command2 = nothing
	end if
	GuardarRegistro=crear_proveedor
end if
end function

function GuardarCliente(ncliente,nproveedor)
	continuar=0
	strselect="select * from proveedores where nproveedor=?"
    set command4 = nothing
    set conn4 = Server.CreateObject("ADODB.Connection")
    set command4 = Server.CreateObject("ADODB.Command")
    conn4.Open = session("dsn_cliente")
    conn4.CursorLocation = 3
    command4.ActiveConnection = conn4
    command4.CommandTimeout = 60
    command4.CommandText = strselect
    command4.CommandType = adCmdText
    command4.Parameters.Append command4.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
    set rst4 = command4.Execute
	if not rst4.eof then
		if ncliente & ""="" then
			'Obtener el último nº de clientes de CONFIGURACION.

            strselect = "select ncliente from configuracion where nempresa=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nempresa",adVarChar,adParamInput,5,session("ncliente")&"")
            rstAux.CursorLocation = adUseClient
            rstAux.open command2, ,adOpenKeyset, adLockOptimistic

			if not rstAux.EOF then
				num=rstAux("ncliente")+1
				num=string(5-len(cstr(num)),"0") + cstr(num)

				'Actualizar el nº de cliente de CONFIGURACION.
				rstAux("ncliente")=rstAux("ncliente")+1
				rstAux.Update
				conn2.close
                set conn2    =  nothing
                set command2 =  nothing
			else

				rstAux.addnew
				rstAux("ncliente")="1"
				rstAux.Update
				conn2.close
                set conn2    =  nothing
                set command2 =  nothing
				num=1
				num=string(5-len(cstr(num)),"0") + cstr(num)
			end if

            ''ricardo 10-3-2009 se cambia este select, ya que da tiempo de espera, por estar mal construida
			''strselect="select * from clientes"
			strselect="select top 1 * from clientes where ncliente like ?+'%'"
            set command3 = nothing
            set conn3 = Server.CreateObject("ADODB.Connection")
            set command3 = Server.CreateObject("ADODB.Command")
            conn3.Open = session("dsn_cliente")
            conn3.CursorLocation = 3
            command3.ActiveConnection = conn3
            command3.CommandTimeout = 60
            command3.CommandText = strselect
            command3.CommandType = adCmdText
            command3.Parameters.Append command3.CreateParameter("@ncliente",adVarChar,adParamInput,10,session("ncliente")&"")
            rst3.CursorLocation = adUseClient
            rst3.open command3, ,adOpenKeyset, adLockOptimistic
			'Crear un nuevo registro.
			rst3.AddNew
			rst3("ncliente")=session("ncliente") & num
			continuar=1
		else
			strselect="select * from clientes where ncliente=?"
            set command3 = nothing
            set conn3 = Server.CreateObject("ADODB.Connection")
            set command3 = Server.CreateObject("ADODB.Command")
            conn3.Open = session("dsn_cliente")
            conn3.CursorLocation = 3
            command3.ActiveConnection = conn3
            command3.CommandTimeout = 60
            command3.CommandText = strselect
            command3.CommandType = adCmdText
            command3.Parameters.Append command3.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
            rst3.CursorLocation = adUseClient
            rst3.open command3, ,adOpenKeyset, adLockOptimistic
			if not rst3.eof then
				continuar=1
			end if
		end if

		if continuar=1 then
			continuar=0
			'ahora actualizamos los datos generales
			rst3("cif")			= Nulear(rst4("cif"))
			rst3("cifedi")		= Nulear(rst4("cifedi"))
			rst3("rsocial")		= Nulear(rst4("razon_social"))
			rst3("ncomercial")	= Nulear(rst4("nombre"))
			rst3("contacto")    = Nulear(rst4("contacto"))
			rst3("web")			= Nulear(rst4("web"))
			rst3("email")		= Nulear(rst4("email"))
		    	rst3("observaciones")	= Nulear(rst4("observaciones"))
			'rst3("aviso")		= Nulear(rst4("aviso"))
			if ncliente="" then
				rst3("falta")   		 = day(date) & "/" & month(date) & "/" & year(date)
				rst3("fbaja")  		 = NULL
			else
				'rst3("falta")		= Nulear(rst4("falta"))
				'rst3("fbaja")		= Nulear(rst4("fbaja"))
			end if
			rst3("telefono2")	= Nulear(rst4("telefono2"))
			rst3("fax")		= Nulear(rst4("fax"))

			'DATOS COMERCIALES
			'rst3("dto")     = Null_z(rst4("descuento"))
			'rst3("dto2")    = Null_z(rst4("descuento2"))
			'rst3("fpago")    = Nulear(rst4("forma_pago"))
			'rst3("tpago")     = Nulear(rst4("tipo_pago"))
			'rst3("recargo")       = Null_z(rst4("recargo"))
			'rst3("re")            = Nz_b(rst4("re"))
			'rst3("CCONTABLE") = Nulear(rst4("cuenta_contable")) 
            strselect = "select codigo from divisas with(nolock) where moneda_base <> 0 and codigo like ?+'%'"
			rst3("divisa") = iif(rst4("divisa")>"",Nulear(rst4("divisa")),DLookupP1(strselect,session("ncliente")&"",adVarChar,15,session("dsn_cliente")))
			'rst3("IRPF")	= Nulear(rst4("IRPF"))
			'rst3("IRPF_Total")	= Nz_b(rst4("IRPF_Total"))
			''rst3("exento_iva")	= Nz_b(rst4("exento_iva"))

			'DATOS BANCARIOS
			'rst3("banco")	     = Nulear(rst4("banco"))
			'rst3("bancodom")     = Nulear(rst4("banco_dom"))
			'rst3("bancopob")     = Nulear(rst4("banco_pob"))
			'rst3("ncuenta")       = Nulear(rst4("ncuenta"))
			'OTROS DATOS
			'rst3("tactividad")     = Nulear(rst4("tactividad"))
			'rst3("transportista")  = Nulear(rst4("transportista"))
			'rst3("portes")         = Nulear(rst4("portes"))
			'rst3("tipo_cliente") = NULL


			if rst4("dir_principal")&"">"" then
				'Abrimos la tabla de domicilios y modificamos el registro para PROVEEDOR
				Seleccion="SELECT * FROM domicilios WHERE codigo =?"
                set commandDomi = nothing
                set connDomi = Server.CreateObject("ADODB.Connection")
                set commandDomi = Server.CreateObject("ADODB.Command")
                connDomi.Open = session("dsn_cliente")
                connDomi.CursorLocation = 3
                commandDomi.ActiveConnection = connDomi
                commandDomi.CommandTimeout = 60
                commandDomi.CommandText = Seleccion
                commandDomi.CommandType = adCmdText
                commandDomi.Parameters.Append commandDomi.CreateParameter("@codigo",adInteger,adParamInput,4,cstr(null_z(rst4("dir_principal")))&"")
                set rstDomi = commandDomi.Execute
				if not rstDomi.eof then
					if ncliente & ""="" then
						'Abrimos la tabla de domicilios y modificamos el registro para CLIENTE
                        strselect = "select * from domicilios where pertenece=?"
                        set commandDomi2 = nothing
                        set connDomi2 = Server.CreateObject("ADODB.Connection")
                        set commandDomi2 = Server.CreateObject("ADODB.Command")
                        connDomi2.Open = session("dsn_cliente")
                        connDomi2.CursorLocation = 3
                        commandDomi2.ActiveConnection = connDomi2
                        commandDomi2.CommandTimeout = 60
                        commandDomi2.CommandText = strselect
                        commandDomi2.CommandType = adCmdText
                        commandDomi2.Parameters.Append commandDomi2.CreateParameter("@pertenece",adVarChar,adParamInput,55,rst3("ncliente")&"")
                        rstDomi2.CursorLocation = adUseClient
                        rstDomi2.Open command2, ,adOpenKeyset, adLockOptimistic
						rstDomi2.AddNew
						rstDomi2("pertenece") = rst3("ncliente")
						rstDomi2("tipo_domicilio") = "PRINCIPAL_CLI"
						continuar=1
					else
						'Abrimos la tabla de domicilios y modificamos el registro para CLIENTE
						Seleccion="SELECT * FROM domicilios WHERE codigo =?"
                        set commandDomi2 = nothing
                        set connDomi2 = Server.CreateObject("ADODB.Connection")
                        set commandDomi2 = Server.CreateObject("ADODB.Command")
                        connDomi2.Open = session("dsn_cliente")
                        connDomi2.CursorLocation = 3
                        commandDomi2.ActiveConnection = connDomi2
                        commandDomi2.CommandTimeout = 60
                        commandDomi2.CommandText = Seleccion
                        commandDomi2.CommandType = adCmdText
                        commandDomi2.Parameters.Append commandDomi2.CreateParameter("@codigo",adInteger,adParamInput,55,cstr(null_z(rst3("dir_principal")))&"")
                        rstDomi2.CursorLocation = adUseClient
                        rstDomi2.Open command2, ,adOpenKeyset, adLockOptimistic
						if not rstDomi2.eof then
							continuar=1
						end if
					end if
					if continuar=1 then
						rstDomi2("domicilio") = Nulear(rstDomi("domicilio"))
						if Nulear(rstDomi("poblacion"))>"" and Nulear(rstDomi("cp"))="" then
                            strselect = "select cod_postal from poblaciones with(nolock) where poblacion=?"
                            cp_aux=DLookupP1(strselect, Nulear(rstDomi("poblacion"))&"",adVarChar,15,DsnIlion)
							if cp_aux="00000" then cp_aux=""
							rstDomi2("cp")     		=Nulear(cp_aux)
						else
							rstDomi2("cp")     		=Nulear(rstDomi("cp"))
						end if
						rstDomi2("poblacion") = Nulear(rstDomi("poblacion"))
						rstDomi2("provincia") = Nulear(rstDomi("provincia"))
						rstDomi2("pais")      = Nulear(rstDomi("pais"))
						rstDomi2("telefono")  = Nulear(rstDomi("telefono"))
						rstDomi2.Update
						rst3("dir_principal")=Nulear(rstDomi2("codigo"))
					end if
				    connDomi2.Close
                    set connDomi2 = nothing
                    set commandDomi2 = nothing
				end if
				connDomi.Close
                set connDomi = nothing
                set commandDomi = nothing
			end if
			continuar=0
			'Modificamos la direccion de envío caso de que exista
			if rst4("dir_envio")&"">"" then
				'Abrimos la tabla de domicilios para PROVEEDOR
				Seleccion="SELECT * FROM domicilios WHERE codigo =?"
                set commandDomi = nothing
                set connDomi = Server.CreateObject("ADODB.Connection")
                set commandDomi = Server.CreateObject("ADODB.Command")
                connDomi.Open = session("dsn_cliente")
                connDomi.CursorLocation = 3
                commandDomi.ActiveConnection = connDomi
                commandDomi.CommandTimeout = 60
                commandDomi.CommandText = Seleccion
                commandDomi.CommandType = adCmdText
                commandDomi.Parameters.Append commandDomi.CreateParameter("@codigo",adInteger,adParamInput,4,cstr(null_z(rst4("dir_envio")))&"")
                set rstDomi = commandDomi.Execute
				if not rstDomi.eof then
					if rstDomi2.state<>0 then rstDomi2.close
					if ncliente & ""="" or rst3("dir_envio") & ""=""  then
						strselect = "select * from domicilios where pertenece=?"
                        set commandDomi2 = nothing
                        set connDomi2 = Server.CreateObject("ADODB.Connection")
                        set commandDomi2 = Server.CreateObject("ADODB.Command")
                        connDomi2.Open = session("dsn_cliente")
                        connDomi2.CursorLocation = 3
                        commandDomi2.ActiveConnection = connDomi2
                        commandDomi2.CommandTimeout = 60
                        commandDomi2.CommandText = strselect
                        commandDomi2.CommandType = adCmdText
                        commandDomi2.Parameters.Append commandDomi2.CreateParameter("@pertenece",adVarChar,adParamInput,55,rst3("ncliente")&"")
                        rstDomi2.CursorLocation = adUseClient
						rstDomi2.AddNew
						rstDomi2("pertenece") = rst3("ncliente")
						rstDomi2("tipo_domicilio") = "ENVIO_CLI"
						continuar=1
					else
						'Abrimos la tabla de domicilios y modificamos el registro para CLIENTE
						Seleccion="SELECT * FROM domicilios WHERE codigo =?"
                        set commandDomi2 = nothing
                        set connDomi2 = Server.CreateObject("ADODB.Connection")
                        set commandDomi2 = Server.CreateObject("ADODB.Command")
                        connDomi2.Open = session("dsn_cliente")
                        connDomi2.CursorLocation = 3
                        commandDomi2.ActiveConnection = connDomi2
                        commandDomi2.CommandTimeout = 60
                        commandDomi2.CommandText = Seleccion
                        commandDomi2.CommandType = adCmdText
                        commandDomi2.Parameters.Append commandDomi2.CreateParameter("@codigo",adInteger,adParamInput,4,cstr(null_z(rst3("dir_envio")))&"")
                        rstDomi2.CursorLocation = adUseClient
                        rstDomi2.Open command2, ,adOpenKeyset, adLockOptimistic
						if not rstDomi2.eof then
							continuar=1
						end if
					end if
					if continuar=1 then
						if rst3("dir_envio")>"" and ncliente>"" then
							rstDomi2("domicilio") = Nulear(rstDomi("domicilio"))
							if Nulear(rstDomi("poblacion"))>"" and Nulear(rstDomi("cp"))="" then
                                strselect = "select cod_postal from poblaciones with(nolock) where poblacion=?"
                                cp_aux=DLookupP1(strselect,Nulear(rstDomi("poblacion"))&"",adVarChar,15,DsnIlion)
								if cp_aux="00000" then cp_aux=""
								rstDomi2("cp")     		=Nulear(cp_aux)
							else
								rstDomi2("cp")     		=Nulear(rstDomi("cp"))
							end if
							rstDomi2("poblacion") = Nulear(rstDomi("poblacion"))
							rstDomi2("provincia") = Nulear(rstDomi("provincia"))
							rstDomi2("pais")      = Nulear(rstDomi("pais"))
							rstDomi2("telefono")  = Nulear(rstDomi("telefono"))
							rstDomi2("A_LA_ATENCION")  = Nulear(rstDomi("A_LA_ATENCION"))
							rstDomi2.Update
						else
							'Abrimos la tabla de domicilios y creamos un registro nuevo
							'Seleccion="select * from domicilios where pertenece='" + rst3("ncliente") +"'"
							'rstDomi2.Open Seleccion,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
							'rstDomi2.AddNew
							rstDomi2("pertenece") = rst3("ncliente")
							rstDomi2("tipo_domicilio") = "ENVIO_CLI"
							rstDomi2("domicilio") = Nulear(rstDomi("domicilio"))
							if Nulear(rstDomi("poblacion"))>"" and Nulear(rstDomi("cp"))="" then
                                strselect = "select cod_postal from poblaciones with(nolock) where poblacion=?"
                                cp_aux=DLookupP1(strselect,Nulear(rstDomi("poblacion"))&"",adVarChar,15,DsnIlion)
								if cp_aux="00000" then cp_aux=""
								rstDomi2("cp")     		=Nulear(cp_aux)
							else
								rstDomi2("cp")     		=Nulear(rstDomi("cp"))
							end if
							rstDomi2("poblacion") = Nulear(rstDomi("poblacion"))
							rstDomi2("provincia") = Nulear(rstDomi("provincia"))
							rstDomi2("pais")      = Nulear(rstDomi("pais"))
							rstDomi2("telefono")  = Nulear(rstDomi("telefono"))
							rstDomi2("A_LA_ATENCION")  = Nulear(rstDomi("A_LA_ATENCION"))
							rstDomi2.Update
							rst3("dir_envio") = Nulear(rstDomi2("codigo"))
						end if
					end if
				    connDomi2.Close
                    set connDomi2 = nothing
                    set commandDomi2 = nothing
				end if
				connDomi.Close
                set connDomi = nothing
                set commandDomi = nothing
			end if
			rst3.update
			ncliente=rst3("ncliente")

			''ricardo 29-4-2003
			'PONEMOS LOS DATOS EN DOCUMENTO_CLI
            strselect = "select * from documentos_cli where ncliente=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
            rstAux.CursorLocation = adUseClient
            rstAux.Open command2, ,adOpenKeyset, adLockOptimistic

			if rstAux.eof then
                rstAux.addnew
			end if
            rstAux.update
            strupdate = "update documentos_cli with(rowlock) set valorado_pre=-1, serie_pre=null, valorado_ped=-1, serie_ped=null, valorado_alb=-1, serie_alb=null,serie_fac=null where ncliente=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strupdate
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
            set rstAux = command2.Execute

			'rstAux("ncliente")=ncliente
			'rstAux("valorado_pre")=-1
			'rstAux("serie_pre")=NULL
			'rstAux("valorado_ped")=-1
			'rstAux("serie_ped")=NULL
			'rstAux("valorado_alb")=-1
			'rstAux("serie_alb")=NULL
			'rstAux("serie_fac")=NULL
			'rstAux.update
			'''''''''
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
		end if
		conn3.Close
        set conn3 = nothing
        set command3 = nothing
	end if
	conn4.Close
    set conn4 = nothing
    set command4 = nothing
	GuardarCliente=ncliente
end function


'ega 15/04/2008 quita las comillas dobles, simples y los salto de linea del texto
function LimpiarTexto(texto)
    texto = Replace(texto, chr(34), "")'quitar comillas dobles
    texto = Replace(texto, chr(39), "")'quitar comillas simples
    texto = Replace(texto, chr(10), "")'quitar salto de linea
    texto = Replace(texto, chr(13), "")'quitar salto de carro
    LimpiarTexto=texto
end function


'Elimina los datos del registro cuando se pulsa BORRAR.
Function BorrarRegistro(nproveedor)
	''on error resume next

	se_puede_borrar=0
	' si existe el cliente en algun documento no se podra borrar, por lo que tampoco se podra borrar de distribuidores
	strselect="SELECT a.nproveedor FROM albaranes_pro AS a with(nolock) WHERE a.nproveedor = '" & nproveedor & "'"
	strselect=strselect & " union all "
	strselect=strselect & "SELECT f.nproveedor FROM facturas_pro AS f with(nolock) WHERE f.nproveedor = '" & nproveedor & "'"
	strselect=strselect & " union all "
	strselect=strselect & "SELECT p.nproveedor FROM pedidos_pro AS p with(nolock) WHERE p.nproveedor = '" & nproveedor & "'"
	strselect=strselect & " union all "
	strselect=strselect & "SELECT d.nproveedor FROM devoluciones_pro AS d with(nolock) WHERE d.nproveedor = '" & nproveedor & "'"
	'rst.open "select nproveedor from albaranes_pro,facturas_pro,pedidos_pro,devoluciones_pro where nproveedor='" & nproveedor & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	rst.open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if rst.eof then
		rst.close
		se_puede_borrar=1
	else
		se_puede_borrar=0
		mens_err = LitMsgNoBorrarDistProvDocCompra
		rst.close
	end if

	' si existe el cliente en algun documento no se podra borrar, por lo que tampoco se podra borrar de distribuidores
	strselect = "select destino from mailing_proveedores with(nolock) where destino=?"
    set command2 = nothing
    set conn2 = Server.CreateObject("ADODB.Connection")
    set command2 = Server.CreateObject("ADODB.Command")
    conn2.Open = session("dsn_cliente")
    conn2.CursorLocation = 3
    command2.ActiveConnection = conn2
    command2.CommandTimeout = 60
    command2.CommandText = strselect
    command2.CommandType = adCmdText
    command2.Parameters.Append command2.CreateParameter("@destino",adVarChar,adParamInput,10,nproveedor&"")
    set rst = command2.Execute

	if rst.eof then
		conn2.close
        set conn2    =  nothing
        set command2 =  nothing
		if se_puede_borrar=1 then
			se_puede_borrar=1
		end if
	else
		se_puede_borrar=0
		mens_err = LitMsgNoBorrarDistProvMailProv
		conn2.close
        set conn2    =  nothing
        set command2 =  nothing

	end if
	strselect = "select ncliente from distribuidores where nproveedor=?"
    set command2 = nothing
    set conn2 = Server.CreateObject("ADODB.Connection")
    set command2 = Server.CreateObject("ADODB.Command")
    conn2.Open = session("dsn_cliente")
    conn2.CursorLocation = 3
    command2.ActiveConnection = conn2
    command2.CommandTimeout = 60
    command2.CommandText = strselect
    command2.CommandType = adCmdText
    command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
    set rst = command2.Execute

	if not rst.eof then
		ncliente=rst("ncliente")
        conn2.close
        set conn2    =  nothing
        set command2 =  nothing

		' si existe el cliente en algun documento no se podra borrar, por lo que tampoco se podra borrar de distribuidores
		strselect = "select destino from mailing_clientes with(nolock) where destino=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@destino",adVarChar,adParamInput,10,ncliente&"")
        set rst = command2.Execute

		if rst.eof then
	    conn2.close
        set conn2    =  nothing
        set command2 =  nothing
			if se_puede_borrar=1 then
				se_puede_borrar=1
			end if
		else
			se_puede_borrar=0
			mens_err = LitMsgNoBorrarCliMailCli
			conn2.close
            set conn2    =  nothing
            set command2 =  nothing
		end if

		strselect="SELECT c.ncliente FROM clientes AS c with(nolock) WHERE c.ndist =(SELECT ndist FROM distribuidores AS d WHERE d.ncliente = '" & ncliente & "')"
		rst.open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if rst.eof then
			rst.close
			if se_puede_borrar=1 then
				se_puede_borrar=1
			end if
		else
			se_puede_borrar=0
			mens_err = LitMsgNoBorrarDistCliAsoc
			rst.close
		end if
		strselect = "select * from contactos_cli with(nolock) where ncliente=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
        set rst = command2.Execute

		if not rst.eof then
			se_puede_borrar=0
			mens_err = LitMsgNoBorrarDistContAsoc
		end if
		conn2.close
        set conn2    =  nothing
        set command2 =  nothing

		strselect = "select top 1 ncliente from centros with(nolock) where ncliente=?" 
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
        set rst = command2.Execute

		if not rst.eof then
			se_puede_borrar=0
			mens_err = LitMsgNoBorrarDistCentroAsoc
		end if
		conn2.close
        set conn2    =  nothing
        set command2 =  nothing
	''ricardo 16-5-2003
	strselect = "select * from HISTORIAL_CLIENTE with(nolock) where ncliente=?"
    set command2 = nothing
    set conn2 = Server.CreateObject("ADODB.Connection")
    set command2 = Server.CreateObject("ADODB.Command")
    conn2.Open = session("dsn_cliente")
    conn2.CursorLocation = 3
    command2.ActiveConnection = conn2
    command2.CommandTimeout = 60
    command2.CommandText = strselect
    command2.CommandType = adCmdText
    command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
    set rst = command2.Execute
	if not rst.eof then
		se_puede_borrar=0
		mens_err = LitMsgNoBorrarDistHistAsoc
	end if
	rst.close

	''ricardo 25/11/2003 comprobamos que no existan datos del cliente en estas tablas:
	strselect = "select * from ventas with(nolock) where ncliente=?"
    set command2 = nothing
    set conn2 = Server.CreateObject("ADODB.Connection")
    set command2 = Server.CreateObject("ADODB.Command")
    conn2.Open = session("dsn_cliente")
    conn2.CursorLocation = 3
    command2.ActiveConnection = conn2
    command2.CommandTimeout = 60
    command2.CommandText = strselect
    command2.CommandType = adCmdText
    command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
    set rst = command2.Execute

	if not rst.eof then
		se_puede_borrar=0
		mens_err = LitMsgNoBorrarDistDocVenta
	end if
	conn2.close
    set conn2    =  nothing
    set command2 =  nothing


	strselect = "select * from tickets with(nolock) where ncliente=?"
    set command2 = nothing
    set conn2 = Server.CreateObject("ADODB.Connection")
    set command2 = Server.CreateObject("ADODB.Command")
    conn2.Open = session("dsn_cliente")
    conn2.CursorLocation = 3
    command2.ActiveConnection = conn2
    command2.CommandTimeout = 60
    command2.CommandText = strselect
    command2.CommandType = adCmdText
    command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
    set rst = command2.Execute

	if not rst.eof then
		se_puede_borrar=0
		mens_err = LitMsgNoBorrarDistDocVenta
	end if
	conn2.close
    set conn2    =  nothing
    set command2 =  nothing

	strselect = "select * from conf_clientes with(nolock) where cliente=?"
    set command2 = nothing
    set conn2 = Server.CreateObject("ADODB.Connection")
    set command2 = Server.CreateObject("ADODB.Command")
    conn2.Open = session("dsn_cliente")
    conn2.CursorLocation = 3
    command2.ActiveConnection = conn2
    command2.CommandTimeout = 60
    command2.CommandText = strselect
    command2.CommandType = adCmdText
    command2.Parameters.Append command2.CreateParameter("@cliente",adVarChar,adParamInput,10,ncliente&"")
    set rst = command2.Execute
	if not rst.eof then
		se_puede_borrar=0
		mens_err = LitMsgNoBorrarDistConfMaqFab
	end if
	conn2.close
    set conn2    =  nothing
    set command2 =  nothing


		strselect="SELECT a.ncliente FROM albaranes_cli AS a with(nolock) WHERE a.ncliente = '" & ncliente & "'"
		strselect=strselect & " union all "
		strselect=strselect &"SELECT f.ncliente FROM facturas_cli AS f with(nolock) WHERE f.ncliente = '" & ncliente & "'"
		strselect=strselect & " union all "
		strselect=strselect &"SELECT p.ncliente FROM pedidos_cli AS p with(nolock) WHERE p.ncliente = '" & ncliente & "'"
		strselect=strselect & " union all "
		strselect=strselect &"SELECT pre.ncliente FROM presupuestos_cli AS pre with(nolock) WHERE pre.ncliente = '" & ncliente & "'"
		strselect=strselect & " union all "
		strselect=strselect &"SELECT d.ncliente FROM devoluciones_cli AS d with(nolock) WHERE d.ncliente = '" & ncliente & "'"
		'rst.open "select ncliente from albaranes_cli,facturas_cli,pedidos_cli,presupuestos_cli,devoluciones_cli where ncliente='" & ncliente & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst.open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if rst.eof then
			rst.close
			if se_puede_borrar=1 then
				'Borramos sus registros de la tabla de documentos_cli
				strselect="delete from documentos_cli with(rowlock) where ncliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strselect
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
                set rst = command2.Execute
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
				'Si se ha creado a partir de un centro borramos del codcliente del centro
                strupdate = "update centros with(rowlock) set codcliente=null where codcliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strupdate
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@codcliente",adVarChar,adParamInput,10,ncliente&"")
                set rst = command2.Execute
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
                strdelete= "delete from distribuidores with(rowlock) where ncliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strdelete
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
                set rst = command2.Execute
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
                strdelete = "delete from clientes with(rowlock) where ncliente=?"
                set command2 = nothing
                set conn2 = Server.CreateObject("ADODB.Connection")
                set command2 = Server.CreateObject("ADODB.Command")
                conn2.Open = session("dsn_cliente")
                conn2.CursorLocation = 3
                command2.ActiveConnection = conn2
                command2.CommandTimeout = 60
                command2.CommandText = strdelete
                command2.CommandType = adCmdText
                command2.Parameters.Append command2.CreateParameter("@ncliente",adVarChar,adParamInput,10,ncliente&"")
                set rst = command2.Execute
                conn2.Close
                set conn2 = nothing
                set command2 = nothing
				if err.number=-2147217900 then
					%><script language="javascript" type="text/javascript">
						alert("<%=LitMsgNoBorrarCliDocAsoc%>");
					</script><%
				else
                    strdelete= "delete from domicilios with(rowlock) where pertenece=? and tipo_domicilio in ('PRINCIPAL_CLI','ENVIO_CLI')"
                    set command2 = nothing
                    set conn2 = Server.CreateObject("ADODB.Connection")
                    set command2 = Server.CreateObject("ADODB.Command")
                    conn2.Open = session("dsn_cliente")
                    conn2.CursorLocation = 3
                    command2.ActiveConnection = conn2
                    command2.CommandTimeout = 60
                    command2.CommandText = strdelete
                    command2.CommandType = adCmdText
                    command2.Parameters.Append command2.CreateParameter("@pertenece",adVarChar,adParamInput,55,ncliente&"")
                    set rst = command2.Execute
                    conn2.Close
                    set conn2 = nothing
                    set command2 = nothing
				end if
				se_puede_borrar=1
			end if
		else
			se_puede_borrar=0
			mens_err = LitMsgNoBorrarDistProvDocVenta
            conn2.Close
            set conn2 = nothing
            set command2 = nothing
		end if
	else
        conn2.Close
        set conn2 = nothing
        set command2 = nothing
	end if

	''ricardo 18-4-2006 si el proveedor tiene historial no se podra borrar
	if se_puede_borrar=1 then
		strselect = "select * from historial_proveedor with(nolock) where nproveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
        set rst = command2.Execute
		if not rst.eof then
			se_puede_borrar=0
			mens_err = LitMsgNoBorrarProvHistAsoc
		end if
    conn2.close
    set conn2    =  nothing
    set command2 =  nothing
	end if

	if se_puede_borrar=1 then
		strselect = "delete from contactos_pro with(rowlock) where nproveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
        set rst = command2.Execute
	else
		strselect = "select * from contactos_pro with(nolock) where nproveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor)
        set rst = command2.Execute
		if not rst.eof then
			se_puede_borrar=0
			if mens_err="" or isnull(mens_err) or mens_err2="" or isnull(mens_err2) then
				mens_err = LitMsgNoBorrarProvContAsoc
			end if
		end if
		conn2.close
        set conn2    =  nothing
        set command2 =  nothing
	end if

	if se_puede_borrar=1 then
        'mmg 21/07/2008
        strselect = "select * from documentos_pro with(nolock) where nproveedor like ?+'%' and nproveedor=?"
		set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,session("ncliente")&"")
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
        set rst = command2.Execute

        if not rst.eof then
		    strselect = "delete from documentos_pro with(rowlock) where nproveedor=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
            set rst = command2.Execute
		end if
		conn2.close
        set conn2    =  nothing
        set command2 =  nothing

		'Se borran las posibles bonificaciones
		strselect = "delete from bonificaciones with(rowlock) where proveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@destino",adVarChar,adParamInput,10,nproveedor&"")
        set rst = command2.Execute

        conn2.close
        set conn2    =  nothing
        set command2 =  nothing

		'Se borra la informacion de proveer
		strselect = "delete from proveer with(rowlock) where nproveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
        set rst = command2.Execute

        conn2.close
        set conn2    =  nothing
        set command2 =  nothing

		'Borramos sus registros de la tabla de domicilios
		strselect = "delete from domicilios with(rowlock) where pertenece=? and tipo_domicilio in ('PRINCIPAL_PROV','ENVIO_PROV')"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@pertenece",adVarChar,adParamInput,55,nproveedor&"")
        set rst = command2.Execute

        conn2.close
        set conn2    =  nothing
        set command2 =  nothing

		strselect = "delete from proveedores with(rowlock) where nproveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
        set rst = command2.Execute
        conn2.close
        set conn2    =  nothing
        set command2 =  nothing

		BorrarRegistro=""
	else
		if mens_err="" then mens_err = LitMsgNoBorrarDistProvDocAsoc
		%><script language="javascript" type="text/javascript">
			alert("<%=mens_err%>");
		</script>
		<%BorrarRegistro=nproveedor
	end if
end Function

'********** CODIGO PRINCIPAL DE LA PÁGINA
const borde=0

'Conexion y cursores
   set rst = Server.CreateObject("ADODB.Recordset")
   set rst2 = Server.CreateObject("ADODB.Recordset")
   set rst3 = Server.CreateObject("ADODB.Recordset")
   set rst4 = Server.CreateObject("ADODB.Recordset")
   set rstAux = Server.CreateObject("ADODB.Recordset")
   set rstAux2 = Server.CreateObject("ADODB.Recordset")
   set rstSelect = Server.CreateObject("ADODB.Recordset")
   set rstDomi = Server.CreateObject("ADODB.Recordset")
   set rstDomi2 = Server.CreateObject("ADODB.Recordset")%>
<form name="proveedores" method="post">
    <%PintarCabecera "proveedores.asp"

	'Leer parámetros de la página
	dim noadd
	mode     = Request.QueryString("mode")
	nproveedor = limpiaCadena(Request.QueryString("nproveedor"))
	if nproveedor="" then
  		nproveedor = limpiaCadena(Request.QueryString("ndoc"))
	end if
	CheckCadena nproveedor
	campo    = limpiaCadena(request.QueryString("campo"))
	criterio = limpiaCadena(request.QueryString("criterio"))
	texto    = limpiaCadena(request.QueryString("texto"))
	salto    = limpiaCadena(request.QueryString("salto"))

	viene     = limpiaCadena(Request.QueryString("viene")&"")
	if viene = "" then viene=limpiaCadena(Request.Form("viene")&"")

	npro=limpiaCadena(request("nproveedor"))

	si_tiene_modulo_importaciones=ModuloContratado(session("ncliente"),ModImportaciones)

    Country = limpiaCadena(request.Form("country"))
	'NEntidad=limpiaCadena(request.form("NEntidad"))
	'Oficina=limpiaCadena(request.form("Oficina"))
	'DC=limpiaCadena(request.form("DC"))
	'Cuenta=limpiaCadena(request.form("Cuenta"))
    Cuenta=limpiaCadena(request.form("ncuenta"))

    '' ASP  10/1/2012 contactos global
	viene2     = limpiaCadena(Request.QueryString("modp")&"")
	if viene2 = "" then viene2=limpiaCadena(Request.Form("modp")&"")

	'' MPC 28/04/2009 Lectura del parámetro cifrepe
    repe=limpiaCadena(request.QueryString("repe"))

	noadd=limpiaCadena(Request.Querystring("noadd"))&""
	if noadd = "" then noadd=limpiaCadena(request.form("noadd"))&""%>
    <input type="hidden" name="mode_accesos_tienda" value="<%=EncodeForHtml(mode)%>">
	<input type="hidden" name="viene" value="<%=EncodeForHtml(viene)%>">
    <input type="hidden" name="viene2" value="<%=EncodeForHtml(viene2)%>">
	<input type="hidden" name="noadd" value="<%=EncodeForHtml(noadd)%>">
	<%''ricardo 1-6-2004 si existen campos personalizables con titulo no nulo si saldra la pestaña de campos personalizables
		si_campo_personalizables=0
        strselect = "select ncampo from camposperso with(nolock) where tabla='PROVEEDORES' and titulo is not null and titulo<>'' and ncampo like ?+'%'"
        set command = nothing
        set conn = Server.CreateObject("ADODB.Connection")
        set command = Server.CreateObject("ADODB.Command")
        conn.Open = session("dsn_cliente")
        conn.CursorLocation = 3
        command.ActiveConnection = conn
        command.CommandTimeout = 60
        command.CommandText = strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@ncampo",adVarChar,adParamInput,7,session("ncliente")&"")
        set rst = command.Execute
		if not rst.eof then
			si_campo_personalizables=1
		else
			si_campo_personalizables=0
		end if
		conn.Close
        set conn = nothing
        set command = nothing
		%><input type="hidden" name="si_campo_personalizables" value="<%=EncodeForHtml(si_campo_personalizables)%>"><%

	''ricardo 1-6-2004 añadir campos personalizables a proveedores
	if mode="browse" or mode="edit" or mode="add" or mode="delete" or mode="save" or mode="convertirprodist" or mode="borrardirenvio" then
		num_campos=0
		if mode="add" then
			redim lista_valores(10+2)
			for ki=1 to 12
				lista_valores(ki)=""
			next
			num_campos=10
		else
			rstAux2.cursorlocation=3
            strselect = "select c.campo01,c.campo02,c.campo03,c.campo04,c.campo05,c.campo06,c.campo07,c.campo08,c.campo09,c.campo10 from proveedores as c with(nolock) where c.nproveedor=?"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
            set rstAux2 = command.Execute
			if not rstAux2.eof then
				redim lista_valores(10+2)
				lista_valores(1)=Nulear(rstAux2("campo01"))
				lista_valores(2)=Nulear(rstAux2("campo02"))
				lista_valores(3)=Nulear(rstAux2("campo03"))
				lista_valores(4)=Nulear(rstAux2("campo04"))
				lista_valores(5)=Nulear(rstAux2("campo05"))
				lista_valores(6)=Nulear(rstAux2("campo06"))
				lista_valores(7)=Nulear(rstAux2("campo07"))
				lista_valores(8)=Nulear(rstAux2("campo08"))
				lista_valores(9)=Nulear(rstAux2("campo09"))
				lista_valores(10)=Nulear(rstAux2("campo10"))
				num_campos=10
			else
				redim lista_valores(10+2)
				for ki=1 to 12
					lista_valores(ki)=""
				next
				num_campos=10
			end if
		    conn.Close
            set conn = nothing
            set command = nothing
		end if
	end if

	if mode="borrardirenvio" then
        strselect = "select dir_envio from proveedores with(Nolock) where nproveedor=?"
        dir_envio=DLookupP1(strselect,nproveedor&"",adVarChar,10,session("dsn_cliente"))
		if dir_envio & "">"" then
            strsdelete = "delete from domicilios with(rowlock) where codigo=?"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strsdelete
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@codigo",adInteger,adParamInput,4,dir_envio&"")
            set rstAux = command.Execute
			conn.Close
            set conn = nothing
            set command = nothing
            strupdate = "update proveedores with(rowlock) set dir_envio=NULL where nproveedor=?"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strupdate
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
            set rstAux = command.Execute
			conn.Close
            set conn = nothing
            set command = nothing
            strselect = "select ncliente,nproveedor from distribuidores with(nolock) where nproveedor=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
            set rst2 = command2.Execute
			if not rst2.eof then
                strupdate = "update clientes with(rowlock) set dir_envio=NULL where ncliente=?"
                set command = nothing
                set conn = Server.CreateObject("ADODB.Connection")
                set command = Server.CreateObject("ADODB.Command")
                conn.Open = session("dsn_cliente")
                conn.CursorLocation = 3
                command.ActiveConnection = conn
                command.CommandTimeout = 60
                command.CommandText = strupdate
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,10,rst2("ncliente")&"")
                set rstAux = command.Execute
			    conn.Close
                set conn = nothing
                set command = nothing
			end if
			conn2.Close
            set conn2 = nothing
            set command2 = nothing
		end if
		mode="edit"
	end if

  if mode="convertirprodist" then
	rst.Open "select * from proveedores where nproveedor='" & nproveedor & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if rst.eof then
		'no se puede añadir un distribuidor sin estar dato de alta como proveedor
	else
		ndist=""
		error= "no"
		p_ncliente=""
		if rst("cifedi")>"" then
			if salto="no" then
				rstAux.Open "select ncliente from clientes with(nolock) where ncliente like '" & session("ncliente") & "%' and cifedi='" & rst("cifedi") & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if not rstAux.EOF then
					ncliente=rstAux("ncliente")
					error = "si"
					%><script language="javascript" type="text/javascript">
						if (window.confirm("<%=LitCrearDistribuidorClienteExistente%>")==false){
							document.proveedores.action="proveedores.asp?nproveedor=<%=enc.EncodeForJavascript(nproveedor)%>&mode=browse";
							document.proveedores.submit();
						}
						else{
							document.proveedores.action="proveedores.asp?nproveedor=<%=enc.EncodeForJavascript(nproveedor)%>&mode=convertirprodist&salto=si&ncliente=<%=enc.EncodeForJavascript(ncliente)%>";
							document.proveedores.submit();
						}
					</script>
				<%end if
				rstAux.Close
			else
				p_ncliente=limpiaCadena(request.querystring("ncliente"))
			end if
		else
			error = "si"
			%><script language="javascript" type="text/javascript">
				alert("<%=LitCIFIncorrecto%>");
			</script><%
		end if
		if error= "no" then
			ncliente=GuardarCliente(p_ncliente,nproveedor)
			rst2.open "select * from distribuidores",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			num=d_max("ndist","distribuidores","ndist like '" & session("ncliente") & "%'",session("dsn_cliente"))
			if num & "">"" then
				num=num+1
			else
				num=session("ncliente") & "00001"
			end if
			rst2.AddNew
			rst2("ndist")=completar(cstr(num),10,"0")
			ndist=rst2("ndist")
			rst2("ncliente")=ncliente
			rst2("nproveedor")=nproveedor
			rst2.update
			rst2.close
		end if

		if error= "no" and ndist>"" then
			'rst("ndist")=ndist
			'rst.update
			auditar_ins_bor session("usuario"),ndist,ncliente,"alta","cliente","","distribuidores"
		end if
		ndist=""
	end if
	rst.close
	mode="browse"
  end if
 'Acción a realizar
  if mode="save" then

    if country&"" = "ES" then
     
        'cuenta=trim(NEntidad&Oficina&DC&Cuenta)
        'FLM:130309: comprobamos el número de cuenta de abono.
        strBanco = Mid(cuenta, 1, 4)
	    strOficina = Mid(cuenta, 5, 4)
	    strDC1 = Mid(cuenta, 9, 1)
	    strDC2 = Mid(cuenta, 10, 1)
	    strCuenta = Mid(cuenta, 11, 10)
    
    
	    if cuenta&""="" then
		    CuentaOK=true
	    elseif not Validar_cuenta(strBanco & strOficina, strDC1, False, strDC1Bueno) Then
		    CuentaOK=false
	    elseif not Validar_cuenta(strCuenta, strDC2, True, strDC2Bueno) Then
		    CuentaOK=false
	    else
		    CuentaOK=true
	    end if
    else
        'cuenta=trim(NEntidad&Oficina&DC&Cuenta)
        CuentaOK = true
    end if
    
    
    if CuentaOK= true  then
        rst.Open "select * from proveedores where nproveedor='" & npro & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	    if GuardarRegistro(npro)=1 then
		    'Si este proveedor es tambien distribuidor entonces se guardara las modificaciones tambien en clientes
            strselect = "select * from distribuidores where nproveedor=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,npro&"")
            rst2.CursorLocation = adUseClient
            rst2.Open command2, ,adOpenKeyset, adLockOptimistic
		    if not rst2.eof then
			    ncliente=GuardarCliente(rst2("ncliente"),npro)
		    end if
		    conn2.Close
            set conn2 = nothing
            set command2 = nothing

		    if npro<>"" then
			    auditar_ins_bor session("usuario"),"",rst("nproveedor"),"alta","","","proveedores"
		    end if

		    'AMF:17/12/2010:Si viene de incidencias tengo que volver con el numero de proveedor
		    if viene = "incidencias" then%>
		        <script language="javascript" type="text/javascript">
                    llevarProveedorIncidencia("<%=enc.EncodeForJavascript(rst("nproveedor"))%>");
		        </script>
		    <%end if
	    end if
		rst.Close
	else%>
	    <script language="javascript" type="text/javascript">
			alert("<%=LitCuentaAbonoError%>");
			parent.botones.location="proveedores_bt.asp?mode=edit";
			document.location = "proveedores.asp?mode=edit&noadd=<%=enc.EncodeForJavascript(noadd)%>&repe=<%=enc.EncodeForJavascript(repe)%>&nproveedor=<%=enc.EncodeForJavascript(nproveedor)%>";
		</script>
	<%end if

	if npro<>"" then
		mode="browse"
	else
		if noadd="1" then
			mode="search"
		else
			mode="add"
		end if
	end if
  elseif mode="delete" then
	he_borrado=1
    strselect = "select razon_social from proveedores with(Nolock) where nproveedor=?"
    nomprov=DLookupP1(strselect,nproveedor&"",adVarChar,10,session("dsn_cliente"))
    strselect = "select ndist,ncliente from distribuidores with(nolock) where nproveedor=?"
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command = Server.CreateObject("ADODB.Command")
    conn.Open = session("dsn_cliente")
    conn.CursorLocation = 3
    command.ActiveConnection = conn
    command.CommandTimeout = 60
    command.CommandText = strselect
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@nproveedor",adVarChar,adParamInput,10,npro&"")
    set rstAux = command.Execute
	if not rstAux.eof then
		ndist=rstAux("ndist")
		ncliente=rstAux("ncliente")
	else
		ndist=""
		ncliente=""
	end if
	conn.Close
    set conn = nothing
    set command = nothing
	pro=BorrarRegistro(npro)
	if pro>"" then
		mode="browse"
		nproveedor=pro
	else
		if ndist>"" then
			auditar_ins_bor session("usuario"),ndist,ncliente,"baja",npro,"","distribuidores"
		else
			auditar_ins_bor session("usuario"),"",nproveedor,"baja",nomprov,"","proveedores"
		end if
		mode="add"
		nproveedor="" %>
        <script language="javascript" type="text/javascript">
            //dgb: change to add, refresh search page and open it
			    parent.botones.document.location = "proveedores_bt.asp?mode=add";
			    SearchPage("proveedores_lsearch.asp?mode=init",0);
		</script>
        <%
	end if
  end if
    strselect = "select ndist from distribuidores with(nolock) where nproveedor=?"
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command = Server.CreateObject("ADODB.Command")
    conn.Open = session("dsn_cliente")
    conn.CursorLocation = 3
    command.ActiveConnection = conn
    command.CommandTimeout = 60
    command.CommandText = strselect
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
    set rst2 = command.Execute
	if not rst2.eof then
		titulo2=LitProveedores2
	else
		titulo2=LitProveedores
	end if
	conn.Close
    set conn = nothing
    set command = nothing
	if noadd="1" and mode="add" then
		mode="search"
	end if

	Alarma "Proveedores.asp"

    if noadd="1" and mode="add" then
        mode="search"
    end if
    if mode="add" then 
    %><table width="100%"><tr><td>
        <div class="headers-wrapper">
        <%

            DrawDiv "header-nproveedor","",""
            DrawLabel "headerLabel","",LitNproveedor
            DrawSpan "","",EncodeForHtml(trimCodEmpresa(nproveedor)), ""
            CloseDiv

            DrawDiv "header-rsocial","",""
            DrawLabel "headerLabel","",LitRSocial

            strselect = "select razon_social from proveedores with(nolock) where nproveedor=?"
            DrawSpan "","",EncodeForHtml(DLookupP1(strselect,nproveedor&"",adVarChar,50,session("dsn_cliente"))), ""
            CloseDiv
        %>
        </div>
    </td></tr></table><%
	' Inicio Borde Span
	%><table width="100%"><tr><td>

    <%'Colapsa o despliega todas las secciones %>
    <div id="CollapseSection">
    <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['AddDG', 'AddDC', 'AddDB','AddCD','AddOD','AddDD','AddCP']);hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll %>" alt="" <%=ParamImgCollapse %> title=""/></a>
    <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['AddDG', 'AddDC', 'AddDB','AddCD','AddOD','AddDD','AddCP']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll %>" alt="" <%=ParamImgCollapse %> title=""/></a>
    </div>

<%	 'DATOS GENERALES MODO AÑADIR %>
<div class="Section" id="S_AddDG">
    <a href="#" rel="toggle[AddDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitDatosGenerales%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
    </div></a>

    <div class="SectionPanel" style="" id="AddDG">
        <table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="2">
            <%DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitRSocial
            DrawInput "", "", "razon_social", "", "maxlength='50' size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitNombre
            DrawInput "", "", "nombre", "", "maxlength='50' size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitCIF
            DrawInput "", "", "cif", "", "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitContacto
            DrawInput "", "", "contacto", "", "maxlength='50' size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFalta
            DrawInput "", "", "falta", date, "size='10'"
            DrawCalendar "falta"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFbaja
            DrawInput "", "", "fbaja", "", "size='10'"
            DrawCalendar "fbaja"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitDomicilio
            DrawInput "", "", "domicilio", "", "size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPoblacion
            DrawInput "", "", "poblacion", "", "size='25'"
			%>
			    <a class='' href="javascript:AbrirVentana('../configuracion/poblaciones.asp?mode=buscar&viene=proveedores&titulo=<%=LitSelPobla%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPobla%>'; return true;" OnMouseOut="self.status=''; return true;">
                    <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>">
			    </a>
            <%CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitCP
            DrawInput "", "", "cp", "", "maxlength='10' size='5'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitProvincia
            DrawInput "", "", "provincia", "", "maxlength='50' size='25'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPais
            DrawInput "", "", "pais", "", "maxlength='30' size='30'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel1
            DrawInput "", "", "telefono", "", "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel2
            DrawInput "", "", "telefono2", "", "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFax
            DrawInput "", "", "fax", "", "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitEMail
            DrawInput "", "", "email", "", "maxlength='255' size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitWEB
            DrawInput "", "", "web", "", "maxlength='100' size='35'"
            CloseDiv

            if si_tiene_modulo_OrCU <> 0 then  
                DrawDiv "1","",""
                DrawLabel "","",LITCAE
                DrawInput "", "", "cae", "", "maxlength='20' size='25'"
                CloseDiv
	        end if

            DrawDiv "1","",""
            DrawLabel "","",LitObservaciones
            DrawTextarea "width60", "", "observaciones", "", "rows='2' cols='30'"
            CloseDiv
        %>
        </table>
    </div>
</div>

<%'DATOS DELEGACION MODO AÑADIR %>
<div class="Section" id="S_AddDD">
    <a href="#" rel="toggle[AddDD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitDatosEnvio%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
    </div></a>

    <div class="SectionPanel" style="display:none " id="AddDD">
        <table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="2">
		    <%DrawDiv "1","",""
            DrawLabel "","",LitDomicilio
            DrawInput "", "", "del_domicilio", "", "size='35'"
            CloseDiv
                
            DrawDiv "1","",""
            DrawLabel "","",LitPoblacion
            DrawInput "", "", "del_poblacion", "", "size='25'"
            %>
				<a class='' href="javascript:AbrirVentana('../configuracion/poblaciones.asp?mode=buscar&viene=proveedores2&titulo=<%=LitSelPobla%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPobla%>'; return true;" OnMouseOut="self.status=''; return true;">
                    <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>">
				</a>
            <%
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitCP
            DrawInput "", "", "del_cp", "", " maxlength='10' size='5'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitProvincia
            DrawInput "", "", "del_provincia", "", "size='25'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPais
            DrawInput "", "", "del_pais", "", "maxlength='30' size='30'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel1
            DrawInput "", "", "del_telefono", "", "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitContacto
            DrawInput "", "", "del_contacto", "", "size='20'"
            CloseDiv%>
        </table>
    </div>
</div>

<% 'DATOS COMERCIALES MODO AÑADIR %>
<div class="Section" id="S_AddDC">
    <a href="#" rel="toggle[AddDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitDatosComerciales%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
    </div></a>

    <div class="SectionPanel" style="display:none " id="AddDC">
        <table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="2">
        <%
            DrawDiv "1","",""
            DrawLabel "","","% " + LitDescuento 
            DrawInput "", "", "descuento", "", "size='4'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","","% " + LitDescuento2 
            DrawInput "", "", "descuento2", "", "size='4'"
            CloseDiv

            rstSelect.cursorlocation=3
            strselect = "select codigo, descripcion from formas_pago with(nolock) where codigo like ?+'%' order by descripcion"
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
            set rstSelect = command.Execute
		    DrawSelectCelda "CELDA colspan=2",200,"",0,LitFormaPago,"forma_pago",rstSelect,"","codigo","descripcion","","" 
			conn.Close
            set conn = nothing
            set command = nothing

            rstSelect.cursorlocation=3
            strselect = "select codigo, descripcion from tipo_pago with(nolock) where codigo like ?+'%' order by descripcion"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,8,session("ncliente")&"")
            set rstSelect = command.Execute
		    DrawSelectCelda "CELDA colspan=2",200,"",0,LitTipoPago,"tipo_pago",rstSelect,"","codigo","descripcion","",""
		    conn.Close
            set conn = nothing
            set command = nothing

            DrawDiv "1","",""
            DrawLabel "","",LitPrimerVen 
            DrawInput "", "", "e_primer_ven", 0, "maxlength='2' size='3' onchange='comprobar();'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitSegunVen 
            DrawInput "", "", "e_segundo_ven", 0, "maxlength='2' size='3' onchange='comprobar();'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTercerVen 
            DrawInput "", "", "e_tercer_ven", 0, "maxlength='2' size='3' onchange='comprobar();'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","","% " + LitRFinanciero 
            DrawInput "", "", "recargo", "", "size='4'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitREquivalencia
            DrawCheck "","","re",""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","","% " + LitIRPF
            DrawInput "", "", "IRPF", "", "size='4'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitIRPF_Total
            DrawCheck "","","IRPF_Total",""
            CloseDiv

			rstSelect.open "select tipo_iva as codigo,tipo_iva as descripcion from tipos_iva with(nolock) order by tipo_iva",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "CELDA colspan=2",50,"",0,LitTipIvaPro,"iva",rstSelect,"","codigo","descripcion","",""
			rstSelect.close

            strselect = "select codigo, abreviatura from divisas with(nolock) where codigo like ?+'%'"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,15,session("ncliente")&"")
            set rstSelect = command.Execute
			DrawSelectCelda "CELDA colspan=2","","",0,LitDivisa,"divisa",rstSelect,"","codigo","abreviatura","",""
			conn.Close
            set conn = nothing
            set command = nothing%>
        </table>
        <%'DGB 
            DrawDiv "3-sub","",""
                    DrawLabel "","",LITCONTA
                    CloseDiv
          %><table class="DataTable"><%
            DrawDiv "1","",""
            DrawLabel "","",LitCContable
            DrawInput "", "", "cuenta_contable", "", "size='25'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitCContable_efecto
            DrawInput "", "", "cuenta_contable_efecto", "", "size='25'"
            CloseDiv

		    '**RGU 13/6/2006
            DrawDiv "1","",""
            DrawLabel "","",LitIntracomunitario
            %><input class='' type='checkbox' name='intra' <%=EncodeForHtml(chd)%> onclick="javascript:iva0()" /><%
            CloseDiv
			'**RGU**
            DrawDiv "1","",""
            DrawLabel "","",LITRECC	 
            %><input class='' type='checkbox' name='recc' <%=EncodeForHtml(chd2)%> /><%          
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LITINVSUJETOPASIVO
            %><input class='' type='checkbox' name='invsp' onclick="javascript:iva0()" /><%
            CloseDiv
            if si_tiene_modulo_SANTOS<>0 then
                DrawDiv "1","",""
                DrawLabel "","",LITCCONTPAGOVENC
                DrawInput "", "", "CCONTABLE_PAGOVENC", "", "size='25'"
                CloseDiv
            end if%></table>
        </div>
    </div>

<%'DATOS BANCARIOS MODO AÑADIR %>
<div class="Section" id="S_AddDB">
    <a href="#" rel="toggle[AddDB]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitDatosBancarios%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
    </div></a>

    <div class="SectionPanel" style="display:none " id="AddDB">
        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
        <%
                DrawDiv "1","",""
                DrawLabel "","",LitBanco
                DrawInput "", "", "banco", "", "size='33'"
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitBancoDom
                DrawInput "", "", "banco_dom", "", "size='33'"
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitNCuenta%><input class='' type="text" name="country" value='' maxlength="2" size="2"  onkeyup="if (this.value.length==2) document.proveedores.iban.focus()" onblur="this.value=this.value.toUpperCase();"/>
                    <input class='' type="text" name="iban" value='' maxlength="2" size="2"  onkeyup="if (this.value.length==2) document.proveedores.ncuenta.focus()"/>
                    <!--<input class='CELDA' type="hidden" name="NEntidad" value='' maxlength="4" size="3"  onkeyup="if (this.value.length==4) document.proveedores.Oficina.focus()"/>-->
                    <!--<input class='CELDA' type="hidden" name="Oficina" value='' maxlength="4" size="3" onkeyup="if (this.value.length==4) document.proveedores.DC.focus()"/>-->
                    <!--<input class='CELDA' type="hidden" name="DC" value='' maxlength="2" size="1" onkeyup="if (this.value.length==2) document.proveedores.Cuenta.focus()"/>-->
                    <!--<input class='CELDA' type="hidden" name="Cuenta" value='' maxlength="10" size="10"/>-->
                    <input class='' type="text" name="ncuenta" value='' maxlength="28" size="20" /><%
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitBICSWIFT
                DrawInput "", "", "bic", "", "maxlength='11' size='11'"
                CloseDiv

			    rstSelect.cursorlocation=3
                strselect = "select distinct ncuenta from bancos with(nolock) where nbanco like ?+'%'"
                set command = nothing
                set conn = Server.CreateObject("ADODB.Connection")
                set command = Server.CreateObject("ADODB.Command")
                conn.Open = session("dsn_cliente")
                conn.CursorLocation = 3
                command.ActiveConnection = conn
                command.CommandTimeout = 60
                command.CommandText = strselect
                command.CommandType = adCmdText
                command.Parameters.Append command.CreateParameter("@nbanco",adVarChar,adParamInput,10,session("ncliente")&"")
                set rstSelect = command.Execute
			    DrawSelectCelda "","","",0,LitNCuentaCargo,"ncuentacargo",rstSelect,"","ncuenta","ncuenta","",""
			    
				conn.Close
                set conn = nothing
                set command = nothing

                DrawDiv "1","",""
                DrawLabel "","",LitDomiciliacion
                DrawCheck "", "", "Domiciliacion", ""
                CloseDiv%>
        </table>
    </div>
</div>

<% 'OTROS DATOS MODO AÑADIR %>
<div class="Section" id="S_AddOD">
    <a href="#" rel="toggle[AddOD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitOtrosDatos%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
    </div></a>

    <div class="SectionPanel" style="display:none " id="AddOD">
        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="2">
        <%
        'jcg
        if si_tiene_modulo_proyectos<>0 then
                DrawDiv "1","",""
                DrawLabel "","",LitProyecto
                    'frProyecto
                    %><input class="" type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(tmpProyecto)%>"/>
                    <iframe id='frProyecto' src='../mantenimiento/docproyectos.asp?viene=proveedores&mode=<%=EncodeForHtml(mode)%>&cod_proyecto=<%=EncodeForHtml(tmpProyecto)%>' class="width60 iframe-menu" frameborder="no" scrolling="no" noresize="noresize"></iframe><%
                CloseDiv
        end if
            strselect = "select codigo, descripcion from tipo_actividad with(nolock) where codigo like ?+'%'"
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
            set rstSelect = command.Execute
			DrawSelectCelda "CELDA colspan=2",200,"",0,LitTActividad,"tactividad",rstSelect,"","codigo","descripcion","",""
			conn.Close
            set conn = nothing
            set command = nothing
            strselect = "select codigo, descripcion from tipos_entidades with(nolock) where tipo = ? and codigo like ?+'%'"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@tipo",adVarChar,adParamInput,20,LitPROVEEDOR&"")
            command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
            set rstSelect = command.Execute
			DrawSelectCelda "CELDA colspan=2",200,"",0,LitTProveedor,"tipo_proveedor",rstSelect,"","codigo","descripcion","",""
			conn.Close
            set conn = nothing
            set command = nothing

            DrawDiv "1","",""
            DrawLabel "","",LitTransportista
            DrawInput "", "", "transportista", "", "size='25'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPortes
			defecto=""
            %><select class="width20" name="portes">
			 	<%if defecto=LitDebidos then%>
					<option selected value="<%=LitDebidos%>"><%=LitDebidos%></option>
                    <option value="<%=LitPagados%>"><%=LitPagados%></option>
                    <option value=""></option>
                <%elseif defecto=LitPagados then%>
                    <option value="<%=LitDebidos%>"><%=LitDebidos%></option>
                    <option selected value="<%=LitPagados%>"><%=LitPagados%></option>
                    <option value=""></option>
                <%else%>
                    <option value="<%=LitDebidos%>"><%=LitDebidos%></option>
                    <option value="<%=LitPagados%>"><%=LitPagados%></option>
                    <option selected value=""></option>
                <%end if%>
			    </select><%
            CloseDiv
		'CloseFila%>
        </table>
    </div>
</div>

<% 'CONFIG DOC MODO AÑADIR %>
<div class="Section" id="S_AddCD">
    <a href="#" rel="toggle[AddCD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitConfDoc2%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
    </div></a>

    <div class="SectionPanel" style="display:none " id="AddCD">
        <table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="2">
            <%DrawDiv "1","",""
            DrawLabel "","",LitValorPed
            DrawCheck "", "", "valorado_ped", iif(valorado_ped="",-1,nz_b(valorado_ped))
            CloseDiv

			rstAux.cursorlocation=3
			strselect = "select nserie,(right(nserie,len(nserie)-5) + ' - ' + nombre) as nombre from series with(nolock) where nserie like ?+'%' and tipo_documento =?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nserie",adVarChar,adParamInput,10,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo_documento",adVarChar,adParamInput,50,"PEDIDO A PROVEEDOR")
            set rstAux = command2.Execute
	 		DrawSelectCelda "CELDA","200","",0,LitSeriePed,"serie_ped",rstAux,serie_ped,"nserie","nombre","",""
	 		conn2.close
            set conn2    =  nothing
            set command2 =  nothing

            DrawDiv "1","",""
            DrawLabel "","",LitValorAlb
            DrawCheck "", "", "valorado_alb", iif(valorado_alb="",-1,nz_b(valorado_alb))
            CloseDiv

			rstAux.cursorlocation=3
			strselect = "select nserie,(right(nserie,len(nserie)-5) + ' - ' + nombre) as nombre from series with(nolock) where nserie like ?+'%' and tipo_documento=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nserie",adVarChar,adParamInput,10,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo_documento",adVarChar,adParamInput,50,"ALBARAN DE PROVEEDOR")
            set rstAux = command2.Execute
			DrawSelectCelda "CELDA","200","",0,LitSerieAlb,"serie_alb",rstAux,serie_alb,"nserie","nombre","",""
			conn2.close
            set conn2    =  nothing
            set command2 =  nothing

			rstAux.cursorlocation=3
			strselect = "select nserie,(right(nserie,len(nserie)-5) + ' - ' + nombre) as nombre from series with(nolock) where nserie like ?+'%' and tipo_documento=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nserie",adVarChar,adParamInput,10,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo_documento",adVarChar,adParamInput,50,"FACTURA DE PROVEEDOR")
            set rstAux = command2.Execute
			DrawSelectCelda "CELDA","200","",0,LitSerieFac,"serie_fac",rstAux,serie_fac,"nserie","nombre","",""
			conn2.close
            set conn2    =  nothing
            set command2 =  nothing%>
        </table>
    </div>
</div>

<%'CAMPOS PERSONALIZABLES MODO AÑADIR %>
<%if si_campo_personalizables=1 then%>
<div class="Section" id="S_AddCP">
    <a href="#" rel="toggle[AddCP]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
    <div class="SectionHeader">
        <%=LitCampPersoPro%>
        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
    </div></a>

    <div class="SectionPanel" style="display:none " id="AddCP">
	    <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5"><%

            strselect = "select * from camposperso with(nolock) where tabla='PROVEEDORES' and ncampo like ?+'%' order by ncampo,titulo"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@ncampo",adVarChar,adParamInput,7,session("ncliente")&"")
            set rst2 = command2.Execute

		    if not rst2.eof then
			    num_campos_existen=rst2.recordcount
				    num_campo=1
				    num_campo2=1
				    num_puestos=0
				    num_puestos2=0
				    while not rst2.eof
					    if num_puestos2>0 and (num_puestos2 mod 2)=0 then
						    num_puestos2=0
					    end if
					    if rst2("titulo") & "">"" then
						    if ((num_puestos-1) mod 2)=0 then
						    end if
						    num_puestos=num_puestos+1
						    num_puestos2=num_puestos2+1
                            DrawDiv "1","",""
                            DrawLabel "","",EncodeForHtml(rst2("titulo"))

						    valor_campo_perso=""
						    if rst2("tipo")=1 then
							    if isNumeric(rst2("tamany")) then
								    tamany=rst2("tamany")
							    else
								    tamany=1
							    end if
							    %><!--<td class="CELDA7" style='width:200px' align="left">-->
								    <input type="text" class="" name="<%="campo" & num_campo%>" size="35" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>">
							    <!--</td>--><%
                                CloseDiv
						    elseif rst2("tipo")=2 then
						        DrawCheck "", "", "campo" & num_campo, iif(valor_campo_perso="on",-1,0)
                                CloseDiv
                            elseif rst2("tipo")=3 then
							    num_campo_str=cstr(num_campo)
							    if len(num_campo_str)=1 then
								    num_campo_str="0" & num_campo_str
							    end if

                                strselect = "select ndetlista,valor from campospersolista with(nolock) where tabla='PROVEEDORES' and ncampo=? and valor is not null and valor<>'' order by valor,ndetlista"
                                set command2 = nothing
                                set conn2 = Server.CreateObject("ADODB.Connection")
                                set command2 = Server.CreateObject("ADODB.Command")
                                conn2.Open = session("dsn_cliente")
                                conn2.CursorLocation = 3
                                command2.ActiveConnection = conn2
                                command2.CommandTimeout = 60
                                command2.CommandText = strselect
                                command2.CommandType = adCmdText
                                command2.Parameters.Append command2.CreateParameter("@ncampo",adVarChar,adParamInput,8,session("ncliente")&num_campo_str&"")
                                set rstAux = command2.Execute%>
								    <select class="" name="campo<%=EncodeForHtml(num_campo)%>" style='width:200px'>  
									    <%encontrado=0
									    while not rstAux.eof
										    if valor_campo_perso & "">"" and isnumeric(valor_campo_perso) then
											    valor_campo_perso_aux=cint(valor_campo_perso)
										    else
											    valor_campo_perso_aux=0
										    end if
										    if valor_campo_perso_aux=cint(rstAux("ndetlista")) then
											    texto_selected="selected"
											    if encontrado=0 then encontrado=1
										    else
											    texto_selected=""
										    end if%>
										    <option value="<%=EncodeForHtml(rstAux("ndetlista"))%>"  <%=EncodeForHtml(texto_selected)%> ><%=EncodeForHtml(rstAux("valor"))%></option>       
										    <%rstAux.movenext
									    wend%>
									    <option <%=iif(encontrado=1,"","selected")%> value=""></option>
								    </select>
                                <%
                                CloseDiv
						    elseif rst2("tipo")=4 then
							    if isNumeric(rst2("tamany")) then
								    tamany=rst2("tamany")
							    else
								    tamany=1
							    end if
							    %><input type="text" class="" name="<%="campo" & EncodeForHtml(num_campo)%>" size="35" maxlength="<%=EncodeForHtml(tamany)%>" value="<%=EncodeForHtml(valor_campo_perso)%>">
							    <%CloseDiv
						    elseif rst2("tipo")=5 then
							    if isNumeric(rst2("tamany")) then
								    tamany=rst2("tamany")
							    else
								    tamany=1
							    end if
							    %><input type="text" class="" name="<%="campo" & EncodeForHtml(num_campo)%>" size="35" maxlength="<%=EncodeForHtml(tamany)%>" value="<%=EncodeForHtml(valor_campo_perso)%>">
							    <%CloseDiv
						    end if
					    else
						    %><input type="hidden" name="campo<%=EncodeForHtml(num_campo)%>" value=""><%
					    end if
					    %><input type="hidden" name="tipo_campo<%=EncodeForHtml(num_campo)%>" value="<%=EncodeForHtml(rst2("tipo"))%>"><%
					    %><input type="hidden" name="titulo_campo<%=EncodeForHtml(num_campo)%>" value="<%=EncodeForHtml(rst2("titulo"))%>"><%
					    rst2.movenext
					    num_campo=num_campo+1
					    if not rst2.eof then
						    if rst2("titulo") & "">"" then
							    num_campo2=num_campo2+1
						    end if
					    end if
				    wend
			    num_campos=num_puestos
		    else
			    num_campos=0
			    num_campos_existen=0
		    end if
		    conn2.close
            set conn2    =  nothing
            set command2 =  nothing
	    %></table>
	    <input type="hidden" name="num_campos" value="<%=EncodeForHtml(num_campos_existen)%>">
    </div>
    </div>
    <%end if %>

	</td></tr></table>
    <%elseif mode="browse" then
''ricardo 31/7/2003 comprobamos que existe el pedido
  if mode="browse" and he_borrado<>1 then
		strselect = "select nproveedor from proveedores with(nolock) where nproveedor=?"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nproveedor&"")
        set rstAux = command2.Execute

		if rstAux.eof then
			nproveedor=""
			%><script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgDocsNoExiste%>");
				document.proveedores.action="proveedores.asp?mode=add"
				document.proveedores.submit();
				parent.botones.document.location="proveedores_bt.asp?mode=add&noadd=<%=enc.EncodeForJavascript(noadd)%>";
			</script><%
			mode="add"
		end if
		conn2.close
        set conn2    =  nothing
        set command2 =  nothing
  end if

	if rst.state<>0 then rst.close
      rst.Open "select * from proveedores inner join domicilios on domicilios.codigo=proveedores.dir_principal where nproveedor='" & nproveedor & "' ", session("dsn_cliente"),adOpenKeyset,adLockOptimistic


	%><input type="hidden" name="hnproveedor" value="<%=EncodeForHtml(rst("nproveedor"))%>">

    <%'DATOS GENERALES MODO BROWSE %>
        <div class="headers-wrapper width90">
        <%
            DrawDiv "header-nproveedor","",""
            DrawLabel "headerLabel","",LitNproveedor
            DrawSpan "","",EncodeForHtml(trimCodEmpresa(nproveedor)), ""
            CloseDiv

            DrawDiv "header-rsocial","",""
            DrawLabel "headerLabel","",LitRSocial
            strselect = "select razon_social from proveedores with(nolock) where nproveedor=?"
            DrawSpan "","",EncodeForHtml(DLookupP1(strselect,nproveedor&"",adVarChar,50,session("dsn_cliente"))), ""
            CloseDiv
        %></div><%

    DrawProveedoresAction

    'Pinta la barra de opciones
    BarraOpciones "browse", nproveedor

    claseSpan="span-browser"
    'Inicio Borde Span
	%><table width="100%"><tr><td>

    <%'Colapsa o despliega todas las secciones %>
    <div id="CollapseSection">
    <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['BrowseDG', 'BrowseDC', 'BrowseDB','BrowseCD','BrowseOD','BrowseDD','BrowseCP']);hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll %>" alt="" <%=ParamImgCollapse %> title=""/></a>
    <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['BrowseDG', 'BrowseDC', 'BrowseDB','BrowseCD','BrowseOD','BrowseDD','BrowseCP']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll %>" alt="" <%=ParamImgCollapse %> title=""/></a>
    </div>

    <div class="Section" id="S_BrowseDG">
        <a href="#" rel="toggle[BrowseDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosGenerales%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="" id="BrowseDG">
        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpading="2" cellspacing="5">
		    <%DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitNombre
            DrawSpan claseSpan,"",EncodeForHtml(rst("nombre")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitCIF
            DrawSpan claseSpan,"",EncodeForHtml(rst("cif")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel1
            DrawSpan claseSpan,"",EncodeForHtml(rst("telefono")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFalta
            DrawSpan claseSpan,"",EncodeForHtml(rst("falta")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFbaja
            DrawSpan claseSpan,"",EncodeForHtml(rst("fbaja")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitContacto
            DrawSpan claseSpan,"",EncodeForHtml(rst("contacto")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel2
            DrawSpan claseSpan,"",EncodeForHtml(rst("telefono2")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitDomicilio + LinkGoogleMap(rst("razon_social"),rst("domicilio"), rst("poblacion"), rst("cp"), rst("provincia"), rst("pais"),0)
            DrawSpan claseSpan,"",EncodeForHtml(rst("domicilio")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFAX
            DrawSpan claseSpan,"",EncodeForHtml(rst("fax")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPoblacion
            DrawSpan claseSpan,"",EncodeForHtml(rst("poblacion")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitCP
            DrawSpan claseSpan,"",EncodeForHtml(rst("cp")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitProvincia
            DrawSpan claseSpan,"",EncodeForHtml(rst("provincia")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitWeb
            DrawSpan claseSpan,"",EncodeForHtml(rst("web")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPais
            DrawSpan claseSpan,"",EncodeForHtml(rst("pais")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitEMail
            DrawSpan claseSpan,"",EncodeForHtml(rst("email")),""
            CloseDiv

            if si_tiene_modulo_OrCU <> 0 then
                DrawDiv "1","",""
                DrawLabel "","",LITCAE
                DrawSpan claseSpan,"",EncodeForHtml(rst("cae")),""
                CloseDiv
            end if 

            DrawDiv "1","",""
            DrawLabel "","",LitObservaciones
            'DrawSpan claseSpan,"",pintar_saltos_espacios(EncodeForHtml(rst("observaciones")&"")),""
            DrawSpan claseSpan,"",pintar_saltos_nuevo(EncodeForHtml(rst("observaciones")&"")),""
            CloseDiv%>
	    </table>
        </div>
    </div>

	<%'DATOS DELEGACION MODO BROWSE
    strselect = "select * from domicilios with(nolock) where codigo=?"
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command = Server.CreateObject("ADODB.Command")
    conn.Open = session("dsn_cliente")
    conn.CursorLocation = 3
    command.ActiveConnection = conn
    command.CommandTimeout = 60
    command.CommandText = strselect
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@codigo",adInteger,adParamInput,4,cstr(null_z(rst("dir_envio")))&"")
    set rstSelect = command.Execute
	if not rstSelect.EOF then
	    if isnull(rst("razon_social")) then razon_social = "" else razon_social = rst("razon_social") end if
	    if isnull(rstSelect("domicilio")) then del_domicilio = "" else del_domicilio = rstSelect("domicilio") end if
	    if isnull(rstSelect("telefono")) then del_telefono = "" else del_telefono = rstSelect("telefono") end if
	    if isnull(rstSelect("poblacion")) then del_poblacion = "" else del_poblacion = rstSelect("poblacion") end if
	    if isnull(rstSelect("cp")) then del_cp = "" else del_cp = rstSelect("cp") end if
	    if isnull(rstSelect("provincia")) then del_provincia = "" else del_provincia = rstSelect("provincia") end if
	    if isnull(rstSelect("pais")) then del_pais = "" else del_pais = rstSelect("pais") end if
	    if isnull(rstSelect("a_la_atencion")) then del_contacto = "" else del_contacto = rstSelect("a_la_atencion") end if
	    'del_telefono2=rstSelect("telefono2")
	    'del_fax=rstSelect("fax")
	end if
	conn.Close
    set conn = nothing
    set command = nothing%>

    <div class="Section" id="S_BrowseDD">
        <a href="#" rel="toggle[BrowseDD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosEnvio%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="BrowseDD">
        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpading="2" cellspacing="5">

		    <%DrawDiv "1","",""
            DrawLabel "","",LitTel1
            DrawSpan claseSpan,"",EncodeForHtml(del_telefono),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitDomicilio + LinkGoogleMap(razon_social,del_domicilio, del_poblacion, del_cp, del_provincia, del_pais,0)
            DrawSpan claseSpan,"",EncodeForHtml(del_domicilio),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPoblacion
            DrawSpan claseSpan,"",EncodeForHtml(del_poblacion),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitCP
            DrawSpan claseSpan,"",EncodeForHtml(del_cp),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitProvincia
            DrawSpan claseSpan,"",EncodeForHtml(del_provincia),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPais
            DrawSpan claseSpan,"",EncodeForHtml(del_pais),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitContacto
            DrawSpan claseSpan,"",EncodeForHtml(del_contacto),""
            CloseDiv%>
        </table>
        </div>
    </div>


	<%'DATOS COMERCIALES MODO BROWSE%>
    <div class="Section" id="S_BrowseDC">
        <a href="#" rel="toggle[BrowseDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosComerciales%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="BrowseDC">
	    <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpading="2" cellspacing="5">
            <%DrawDiv "1","",""
            DrawLabel "","","% " + LitDescuento
            DrawSpan claseSpan,"",EncodeForHtml(rst("descuento")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","","% " + LitDescuento2
            DrawSpan claseSpan,"",EncodeForHtml(rst("descuento2")),""
            CloseDiv

 			DrawDiv "1","",""
            DrawLabel "","",LitFormaPago
            strselect = "select descripcion from formas_pago with(nolock) where codigo=?"
            DrawSpan claseSpan,"",EncodeForHtml(DLookupP1(strselect,null_s(rst("forma_pago"))&"",adVarChar,50,session("dsn_cliente"))),""
            CloseDiv

			DrawDiv "1","",""
            DrawLabel "","",LitTipoPago
            strselect = "select descripcion from tipo_pago with(nolock) where codigo=?"
            DrawSpan claseSpan,"",EncodeForHtml(DLookupP1(strselect,null_s(rst("tipo_pago"))&"",adVarChar,50,session("dsn_cliente"))),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPrimerVen
            DrawSpan claseSpan,"",EncodeForHtml(rst("primer_ven")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitSegunVen
            DrawSpan claseSpan,"",EncodeForHtml(rst("segundo_ven")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTercerVen
            DrawSpan claseSpan,"",EncodeForHtml(rst("tercer_ven")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","","% " + LitRFinanciero
            DrawSpan claseSpan,"",EncodeForHtml(rst("recargo")),""
            CloseDiv

			if rst("re") =  0 then
			    sino = LitNo
			else
			    sino = LitSi
			end if

            DrawDiv "1","",""
            DrawLabel "","",LitREquivalencia
            DrawSpan claseSpan,"",sino,""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","","% " + LitIRPF
            DrawSpan claseSpan,"",EncodeForHtml(rst("IRPF")),""
            CloseDiv

			if rst("IRPF_Total") =  0 then
			    sino = LitNo
			else
			    sino = LitSi
			end if

            DrawDiv "1","",""
            DrawLabel "","",LitIRPF_Total
            DrawSpan claseSpan,"",sino,""
            CloseDiv

			DrawDiv "1","",""
            DrawLabel "","",LitTipIvaPro
            DrawSpan claseSpan,"",EncodeForHtml(rst("iva")),""
            CloseDiv

            strselect = "select codigo, abreviatura from divisas with(nolock) where codigo=?"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,15,rst("divisa")&"")
            set rstSelect = command.Execute
			DrawDiv "1","",""
            DrawLabel "","",LitDivisa
            DrawSpan claseSpan,"",EncodeForHtml(rstSelect("abreviatura")),""
            CloseDiv
			conn.Close
            set conn = nothing
            set command = nothing
		%>
	    </table>
        <%'DGB                     
                 DrawDiv "3-sub","",""
                    DrawLabel "","",LITCONTA
                    CloseDiv%>
            <table class="DataTable">
            <%
                DrawDiv "1","",""
                DrawLabel "","",LitCContable
                DrawSpan claseSpan,"",EncodeForHtml(rst("cuenta_contable")),""
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitCContable_Efecto
                DrawSpan claseSpan,"",EncodeForHtml(rst("ccontable_efecto")),""
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitIntracomunitario
                DrawSpan claseSpan,"",EncodeForHtml(visualizar(rst("intra"))),""
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LITRECC
                DrawSpan claseSpan,"",EncodeForHtml(visualizar(rst("recc"))),""
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LITINVSUJETOPASIVO
                DrawSpan claseSpan,"",EncodeForHtml(visualizar(rst("INVSUJETOPASIVO"))),""
                CloseDiv
                if si_tiene_modulo_SANTOS<>0 then
                    DrawDiv "1","",""
                    DrawLabel "","",LITCCONTPAGOVENC
                    DrawSpan claseSpan,"",EncodeForHtml(rst("CCONTABLE_PAGOVENC")),""
                    CloseDiv
                end if
            %>
            </table>
        </div>
        </div>
    </div>

	<% 'DATOS BANCARIOS MODO BROWSE %>
    <div class="Section" id="S_BrowseDB">
        <a href="#" rel="toggle[BrowseDB]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosBancarios%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="BrowseDB">
        <table width="100% "bgcolor='<%=color_blau%>' border='0' cellpading=2 cellspacing=5>
        <%DrawDiv "1","",""
            DrawLabel "","",LitBanco
            DrawSpan claseSpan,"",EncodeForHtml(rst("banco")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitBancoDom
            DrawSpan claseSpan,"",EncodeForHtml(rst("banco_dom")),""
            CloseDiv

            if rst("ncuenta")&"">"" then
                if len(rst("ncuenta"))=24 then
                    iban=mid(rst("ncuenta"),3,2)
                    pais=left(rst("ncuenta"),2)
                    strBanco = Mid(rst("ncuenta"), 5, 4)
	    	        strOficina = Mid(rst("ncuenta"), 9, 4)
    		        strDC = Mid(rst("ncuenta"), 13, 2)
    		        strCuenta = Mid(rst("ncuenta"), 15, 10)
                    ncuenta = pais & " " & iban & " " & strBanco & "-" & strOficina & "-" & strDC & "-" & strCuenta
                else
                    strBanco = Mid(rst("ncuenta"), 1, 4)
	    	        strOficina = Mid(rst("ncuenta"), 5, 4)
    		        strDC = Mid(rst("ncuenta"), 9, 2)
    		        strCuenta = Mid(rst("ncuenta"), 11, 10)
                    ncuenta = strBanco & "-" & strOficina & "-" & strDC & "-" & strCuenta
                end if
            end if
            if ncuenta&"" > "" then
                if CBool(isNumeric(left(rst("ncuenta"),2))) = true then
                    ncuenta = rst("ncuenta")
                else
                    ncuenta = left(rst("ncuenta"),2) & "-" & mid(rst("ncuenta"),3,2) & "-" & right(rst("ncuenta"),len(rst("ncuenta"))-4)
                end if
            else
                ncuenta= ""
            end if

            DrawDiv "1","",""
            DrawLabel "","",LitNCuenta
            DrawSpan claseSpan,"",EncodeForHtml(ncuenta),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitBICSWIFT
            DrawSpan claseSpan,"",EncodeForHtml(rst("swift_code")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitNCuentaCargo
            strselect = "select ncuenta from bancos with(nolock) where ncuenta=?"
            DrawSpan claseSpan,"", EncodeForHtml(DLookupP1(strselect,rst("cuenta_cargo")&"",adVarChar,25,session("dsn_cliente"))),""
            CloseDiv

			if rst("domrec") =  0 then
			   sino = LitNo
			else
			   sino = LitSi
			end if

            DrawDiv "1","",""
            DrawLabel "","",LitDomiciliacion
            DrawSpan claseSpan,"",sino,""
            CloseDiv
		%>
        </table>
        </div>
    </div>

    <% 'OTROS DATOS MODO BROWSE%>
    <div class="Section" id="S_BrowseOD">
        <a href="#" rel="toggle[BrowseOD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitOtrosDatos%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="BrowseOD">
        <table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="5">
		<%'jcg
        if si_tiene_modulo_proyectos<>0 then
                DrawDiv "1","",""
                DrawLabel "","",LitProyecto
                strselect = "select nombre from proyectos with(nolock) where codigo like ?+'%' and codigo=?"
                DrawSpan claseSpan,"",EncodeForHtml(DLookupP2(strselect,session("ncliente"),adVarChar,60,rst("proyecto"),adVarChar,15,session("dsn_cliente"))),""
                CloseDiv
        end if
			if rst("tactividad")>"" then

            strselect = "select descripcion from tipo_actividad with(nolock) where codigo=?"
            
			    Actividad = DLookupP1(strselect,rst("tactividad")&"",adVarChar,150,session("dsn_cliente"))
			else
			    Actividad = ""
			end if

            DrawDiv "1","",""
            DrawLabel "","",LitTActividad
            DrawSpan claseSpan,"",EncodeForHtml(Actividad),""
            CloseDiv

			if rst("tipo_proveedor")>"" then
            strselect = "select descripcion from tipos_entidades with(nolock) where codigo=?"
			    TProveedor = DLookupP1(strselect,rst("tipo_proveedor")&"",adVarChar,50,session("dsn_cliente"))
			else
			    TProveedor = ""
			end if

            DrawDiv "1","",""
            DrawLabel "","",LitTProveedor
            DrawSpan claseSpan,"",EncodeForHtml(TProveedor),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTransportista
            DrawSpan claseSpan,"",EncodeForHtml(rst("transportista")),""
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPortes
            DrawSpan claseSpan,"",EncodeForHtml(rst("portes")),""
            CloseDiv%>
        </table>
        </div>
    </div>

    <% 'CONFIG DOC MODO BROWSE%>
    <div class="Section" id="S_BrowseCD">
        <a href="#" rel="toggle[BrowseCD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitConfDoc2%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="BrowseCD">
	    <table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="5">
	        <%rstSelect.cursorlocation=3

	        nprov=limpiaCadena(request.querystring("nproveedor"))
	        if nprov & "" = "" then nprov = nproveedor
            strselect = "select * from documentos_pro where ncliente=? and nproveedor=?"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
            command.Parameters.Append command.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nprov&"")
            set rstSelect = command.Execute

			    if (rstSelect.EOF=false) then
			        dato=iif(rstSelect("valorado_ped")<>0,LitSi,LitNo)
                    strselect = "select nombre from series with(nolock) where nserie like ?+'%' and nserie=?"
			        dato2=trimCodEmpresa(rstSelect("serie_ped")) & " - " & DLookupP2(strselect,session("ncliente")&"",adVarChar,50, rstSelect("serie_ped")&"",adVarChar,10,session("dsn_cliente"))&""
			    else
			        dato=LitNo
			        dato2="-"
			    end if

                DrawDiv "1","",""
                DrawLabel "","",LitValorPed
                DrawSpan claseSpan,"",EncodeForHtml(dato),""
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitSeriePed
                DrawSpan claseSpan,"",EncodeForHtml(dato2),""
                CloseDiv

			    if (rstSelect.EOF=false) then
			        dato2=iif(rstSelect("valorado_alb")<>0,LitSi,LitNo)
                    strselect = "select nombre from series with(nolock) where nserie like ?+'%' and nserie=?"
			        dato=trimCodEmpresa(rstSelect("serie_alb")) & " - " & DLookupP2(strselect,session("ncliente")&"",adVarChar,10,rstSelect("serie_alb")&"",adVarChar,10,session("dsn_cliente"))&""
			    else
			        dato="-"
			        dato2=LitNo
			    end if

                DrawDiv "1","",""
                DrawLabel "","",LitValorAlb
                DrawSpan claseSpan,"",EncodeForHtml(dato2),""
                CloseDiv

                DrawDiv "1","",""
                DrawLabel "","",LitSerieAlb
                DrawSpan claseSpan,"",EncodeForHtml(dato),""
                CloseDiv

			    if (rstSelect.EOF=false) then
                strselect = "select nombre from series with(nolock) where nserie like ?+'%' and nserie=?"
			        dato=trimCodEmpresa(rstSelect("serie_fac")) & " - " & DLookupP2(strselect,session("ncliente")&"",adVarChar,10,rstSelect("serie_fac")&"",adVarChar,10,session("dsn_cliente"))&""
			    else
			        dato="-"
			    end if

                DrawDiv "1","",""
                DrawLabel "","",LitSerieFac
                DrawSpan claseSpan,"",EncodeForHtml(dato),""
                CloseDiv
                
				conn.Close
                set conn = nothing
                set command = nothing
			rst.Close%>
		</table>
        </div>
    </div>



	<% 'CAMPOS PERSONALIZABLES MODO BROWSE%>
    <%if si_campo_personalizables=1 then%>
    <div class="Section" id="S_BrowseCP">
        <a href="#" rel="toggle[BrowseCP]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitCampPersoPro%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="BrowseCP">
        <table width="100%" bgcolor='<%=color_blau%>' border='0' cellpadding="2" cellspacing="5"><%


        strselect = "select * from camposperso with(nolock) where tabla='PROVEEDORES' and ncampo like ?+'%' order by ncampo,titulo"
        set command2 = nothing
        set conn2 = Server.CreateObject("ADODB.Connection")
        set command2 = Server.CreateObject("ADODB.Command")
        conn2.Open = session("dsn_cliente")
        conn2.CursorLocation = 3
        command2.ActiveConnection = conn2
        command2.CommandTimeout = 60
        command2.CommandText = strselect
        command2.CommandType = adCmdText
        command2.Parameters.Append command2.CreateParameter("@ncampo",adVarChar,adParamInput,8,session("ncliente")&"")
        set rst2 = command2.Execute

        if not rst2.eof then
	        'DrawFila ""
		        num_campo=1
		        num_campo2=1
		        num_puestos=0
		        num_puestos2=0
		        while not rst2.eof
			        if num_puestos2>0 and (num_puestos2 mod 2)=0 then
				        num_puestos2=0
			        end if
			        if rst2("titulo") & "">"" then
                        DrawDiv "1","",""
				        if ((num_puestos-1) mod 2)=0 then
                            DrawLabel "","",EncodeForHtml(rst2("titulo"))
				        else
                            DrawLabel "","",EncodeForHtml(rst2("titulo"))
				        end if
				        num_puestos=num_puestos+1
				        num_puestos2=num_puestos2+1
				        if rst2("tipo")=2 then
				            DrawSpan claseSpan,"",iif(lista_valores(num_campo)=1,LitSi,LitNo),""
                        elseif rst2("tipo")=3 then
					        if lista_valores(num_campo) & "">"" then
						        num_campo_str=cstr(num_campo)
						        if len(num_campo_str)=1 then
							        num_campo_str="0" & num_campo_str
						        end if
                                strselect = "select valor from campospersolista with(nolock) where ncampo=? and tabla=? and ndetlista=?"
						        valor_ListCampPerso=DLookupP3(strselect,session("ncliente") & num_campo_str&"",adVarChar,60,"PROVEEDORES",AdVarChar,11,lista_valores(num_campo),adInteger,4,session("dsn_cliente"))
					        else
						        valor_ListCampPerso=""
					        end if
                            DrawSpan claseSpan,"",EncodeForHtml(valor_ListCampPerso),""
				        else
                            DrawSpan claseSpan,"",EncodeForHtml(lista_valores(num_campo)),""
				        end if
                        CloseDiv
			        end if
			        rst2.movenext
			        num_campo=num_campo+1
			        if not rst2.eof then
				        if rst2("titulo") & "">"" then
					        num_campo2=num_campo2+1
				        end if
			        end if
		        wend
	        num_campos=num_puestos
        else
	        num_campos=0
        end if
        conn2.close
        set conn2    =  nothing
        set command2 =  nothing
        %>
        </table>
        </div>
    </div>

    <%end if%>
	</td></tr></table>

    <%elseif mode="edit" then

        rst.Open "select * from proveedores,domicilios where nproveedor='" & nproveedor & "' and domicilios.codigo=proveedores.dir_principal", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	%>
    <input type="hidden" name="hnproveedor" value="<%=EncodeForHtml(rst("nproveedor"))%>"> 
    
    <% 'DATOS GENERALES MODO EDIT %>
    <table width="100%"><tr><td>
        <div class="headers-wrapper">
        <%
            DrawDiv "header-nproveedor","",""
            DrawLabel "headerLabel","",LitNproveedor
            DrawSpan "","",EncodeForHtml(trimCodEmpresa(nproveedor)), ""
            CloseDiv

            DrawDiv "header-rsocial","",""
            DrawLabel "headerLabel","",LitRSocial

            strselect = "select razon_social from proveedores with(nolock) where nproveedor=?"
            DrawSpan "","",EncodeForHtml(DLookupP1(strselect,nproveedor&"",adVarChar,50,session("dsn_cliente"))), ""
            CloseDiv
        %>
        </div>
    </td></tr></table><%

    BarraOpciones "edit",nproveedor

	' Inicio Borde Span
	%><table width="100%"><tr><td>

    <%'Colapsa o despliega todas las secciones %>
    <div id="CollapseSection">
    <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['EditDG', 'EditDC', 'EditDB','EditCD','EditOD','EditDD','EditCP']);hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll %>" alt="" <%=ParamImgCollapse %> title=""/></a>
    <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['EditDG', 'EditDC', 'EditDB','EditCD','EditOD','EditDD','EditCP']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll %>" alt="" <%=ParamImgCollapse %> title=""/></a>
    </div>

    <div class="Section" id="S_EditDG">
        <a href="#" rel="toggle[EditDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosGenerales%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="" id="EditDG">
        <table width="100%" border='0' cellpadding="2" cellspacing="2">
            <%DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitRSocial
            DrawInput "", "", "razon_social", EncodeForHtml(rst("razon_social")), "maxlength='50' size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitNombre
			if rst("nombre") & "">"" then
                DrawInput "", "", "nombre", EncodeForHtml(replace(rst("nombre"),"'","&#39")), "maxlength='50' size='35'"
			else
                DrawInput "", "", "nombre", "", "maxlength='50' size='35'"
			end if
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitCIF
            DrawInput "", "", "cif", EncodeForHtml(rst("cif")), "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitContacto
            DrawInput "", "", "contacto", EncodeForHtml(rst("contacto")), "maxlength='50' size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFalta
            DrawInput "", "", "falta", EncodeForHtml(rst("falta")), "size='10'"
            DrawCalendar "falta"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFbaja
            DrawInput "", "", "fbaja", EncodeForHtml(rst("fbaja")), "size='10'"
            DrawCalendar "fbaja"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "txtMandatory","",LitDomicilio
            DrawInput "", "", "domicilio", EncodeForHtml(rst("domicilio")), "size='35'"
            CloseDiv

		    DrawDiv "1","",""
            DrawLabel "","",LitPoblacion
            DrawInput "", "", "poblacion", EncodeForHtml(rst("poblacion")), "size='25'"
			%><a class='' href="javascript:AbrirVentana('../configuracion/poblaciones.asp?mode=buscar&viene=proveedores&titulo=<%=LitSelPobla%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPobla%>'; return true;" OnMouseOut="self.status=''; return true;">
                    <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>">
			    </a><%CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitCP
            DrawInput "", "", "cp", EncodeForHtml(rst("cp")), "maxlength='10' size='5'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitProvincia
            DrawInput "", "", "provincia", EncodeForHtml(rst("provincia")), "maxlength='50' size='25'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPais
            DrawInput "", "", "pais", EncodeForHtml(rst("pais")), "maxlength='30' size='30'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel1
            DrawInput "", "", "telefono", EncodeForHtml(rst("telefono")), "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel2
            DrawInput "", "", "telefono2", EncodeForHtml(rst("telefono2")), "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitFax
            DrawInput "", "", "fax", EncodeForHtml(rst("fax")), "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitEMail
            DrawInput "", "", "email", EncodeForHtml(rst("email")), "maxlength='255' size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitWEB
            DrawInput "", "", "web", EncodeForHtml(rst("web")), "maxlength='100' size='35'"
            CloseDiv

            if si_tiene_modulo_OrCU <> 0 then
                DrawDiv "1","",""
                DrawLabel "","",LITCAE
                DrawInput "", "", "cae", EncodeForHtml(rst("cae")), "maxlength='20' size='25'"
                CloseDiv
            end if

            DrawDiv "1","",""
            DrawLabel "","",LitObservaciones
            DrawTextarea "width60", "", "observaciones", EncodeForHtml(rst("observaciones")), "rows='2' cols='60'"
            CloseDiv%>
        </table>
        </div>
    </div>

    <%'DATOS DELEGACION MODO EDIT
    strselect = "select * from domicilios with(nolock) where codigo=?"
    set command = nothing
    set conn = Server.CreateObject("ADODB.Connection")
    set command = Server.CreateObject("ADODB.Command")
    conn.Open = session("dsn_cliente")
    conn.CursorLocation = 3
    command.ActiveConnection = conn
    command.CommandTimeout = 60
    command.CommandText = strselect
    command.CommandType = adCmdText
    command.Parameters.Append command.CreateParameter("@codigo",adInteger,adParamInput,4,cstr(null_z(rst("dir_envio")))&"")
    set rstSelect = command.Execute
	if not rstSelect.EOF then
	    del_domicilio=rstSelect("domicilio")
	    del_telefono=rstSelect("telefono")
	    del_poblacion=rstSelect("poblacion")
	    del_cp=rstSelect("cp")
	    del_provincia=rstSelect("provincia")
	    del_pais=rstSelect("pais")
	    del_contacto=rstSelect("a_la_atencion")
	end if
	conn.Close
    set conn = nothing
    set command = nothing%>
    <div class="Section" id="S_EditDD">
        <a href="#" rel="toggle[EditDD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosEnvio%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="EditDD">
        <table width="100%" border='0' cellpadding="2" cellspacing="2">
        <%
        DrawDiv "3","",""
        %>
            <a class="CELDAREFR7" href="javascript:EliminarDirEnvio('<%=enc.EncodeForJavascript(rst("nproveedor"))%>')" > 
                <img src="../images/<%=ImgVaciarCampo%>" <%=ParamImgVaciarCampo%> alt="<%=LitElimiDirEnvio%>" title="<%=LitElimiDirEnvio%>">
            </a>
        <%
        CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitDomicilio
            DrawInput "", "", "del_domicilio", EncodeForHtml(del_domicilio), "size='32'"
            CloseDiv
            
            DrawDiv "1","",""
            DrawLabel "","",LitPoblacion
            DrawInput "", "", "del_poblacion", EncodeForHtml(del_poblacion), "size='25'"
            %><a class='CELDAREFB'  href="javascript:AbrirVentana('../configuracion/poblaciones.asp?mode=buscar&viene=proveedores2&titulo=<%=LitSelPobla%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPobla%>'; return true;" OnMouseOut="self.status=''; return true;">
                    <img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>">
				</a><%CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitCP
            DrawInput "", "", "del_cp", EncodeForHtml(del_cp), "maxlength='10' size='5'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitProvincia
            DrawInput "", "", "del_provincia", EncodeForHtml(del_provincia), "size='25'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitPais
            DrawInput "", "", "del_pais", EncodeForHtml(del_pais), "maxlength='30' size='30'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitTel1
            DrawInput "", "", "del_telefono", EncodeForHtml(del_telefono), "maxlength='20' size='20'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitContacto
            DrawInput "", "", "del_contacto", EncodeForHtml(del_contacto), "size='20'"
            CloseDiv%>
        </table>
        </div>
    </div>

	<% 'DATOS COMERCIALES MODO EDIT %>
    <div class="Section" id="S_EditDC">
        <a href="#" rel="toggle[EditDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosComerciales%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="EditDC">
        <table width="100%" border='0' cellpadding="2" cellspacing="2">
            <%DrawDiv "1","",""
            DrawLabel "","","% " + LitDescuento 
            DrawInput "", "", "descuento", EncodeForHtml(rst("descuento")), "size='4'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","","% " + LitDescuento2 
            DrawInput "", "", "descuento2", EncodeForHtml(rst("descuento2")), "size='4'"
            CloseDiv

            strselect = "select codigo, descripcion from formas_pago with(nolock) where codigo like ?+'%'"
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
            set rstSelect = command.Execute
			DrawSelectCelda "CELDA",200,"",0,LitFormaPago,"forma_pago",rstSelect,rst("forma_pago"),"codigo","descripcion","",""
			conn.Close
            set conn = nothing
            set command = nothing

            strselect = "select codigo, descripcion from tipo_pago with(nolock) where codigo like ?+'%'"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,8,session("ncliente")&"")
            set rstSelect = command.Execute
			DrawSelectCelda "CELDA",200,"",0,LitTipoPago,"tipo_pago",rstSelect,rst("tipo_pago"),"codigo","descripcion","",""
			conn.Close
            set conn = nothing
            set command = nothing
		'cag
			DrawDiv "1","",""
            DrawLabel "","",LitPrimerVen 
            DrawInput "", "", "e_primer_ven", EncodeForHtml(rst("primer_ven")), "maxlength='2' size='3' onchange='comprobar();'"
            CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LitSegunVen 
            DrawInput "", "", "e_segundo_ven", EncodeForHtml(rst("segundo_ven")), "maxlength='2' size='3' onchange='comprobar();'"
            CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LitTercerVen 
            DrawInput "", "", "e_tercer_ven", EncodeForHtml(rst("tercer_ven")), "maxlength='2' size='3' onchange='comprobar();'"
            CloseDiv
		'fin cag
            DrawDiv "1","",""
            DrawLabel "","","% " + LitRFinanciero
            DrawInput "", "", "recargo", EncodeForHtml(rst("recargo")), "size='4'"
            CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LitREquivalencia
            DrawCheck "","","re",rst("re")
            CloseDiv
            DrawDiv "1","",""
            DrawLabel "","","% " + LitIRPF
            DrawInput "", "", "IRPF", EncodeForHtml(rst("IRPF")), "size='4'"
            CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LitIRPF_Total
            DrawCheck "","","IRPF_Total",rst("IRPF_Total")
            CloseDiv
		    defecto=iif(rst("iva")>"",rst("iva"),"")

		    rstSelect.open "select tipo_iva as codigo,tipo_iva as descripcion from tipos_iva with(nolock) order by tipo_iva",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		    DrawSelectCelda "CELDA "&dsb,50,"",0,LitTipIvaPro,"iva",rstSelect,defecto,"codigo","descripcion","",""
		    rstSelect.close

            strselect = "select codigo, abreviatura from divisas with(nolock) where codigo like ?+'%'"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,15,session("ncliente")&"")
            set rstSelect = command.Execute
			DrawSelectCelda "CELDA","","",0,LitDivisa,"divisa",rstSelect,rst("divisa"),"codigo","abreviatura","",""
			conn.Close
            set conn = nothing
            set command = nothing%>
        </table>
        <%'DGB  
            DrawDiv "3-sub","",""
                    DrawLabel "","",LITCONTA
                    CloseDiv%>
            <table class="DataTable">
            <%DrawDiv "1","",""
            DrawLabel "","",LitCContable
            DrawInput "", "", "cuenta_contable", EncodeForHtml(rst("cuenta_contable")), "size='25'"
            CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LitCContable_efecto
            DrawInput "", "", "cuenta_contable_efecto", EncodeForHtml(rst("ccontable_efecto")), "size='25'"
            CloseDiv
		    chd=""
		    if cint(rst("intra"))=-1 then
			    chd="checked"
		    end if
		    '**RGU 12/6/2006
            DrawDiv "1","",""
            DrawLabel "","",LitIntracomunitario
		    %><!--<td class="CELDA">-->
                <input class='' type='checkbox' name='intra' <%=EncodeForHtml(chd)%> onclick="javascript:iva0()" />
		    <!--</td>-->
            <%
		    CloseDiv
            '**RGU**
            chd2=""
		    if cint(rst("recc"))=-1 then
			    chd2="checked"			    
		    end if
            DrawDiv "1","",""
            DrawLabel "","",LITRECC
		    %>
                <input class='' type='checkbox' name='recc' <%=EncodeForHtml(chd2)%> />
            <%
            CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LITINVSUJETOPASIVO
		    %><!--<td class="CELDA">-->
                <input class='' type='checkbox' name='invsp' <%=iif(nz_b2(rst("INVSUJETOPASIVO"))=1, "checked","")%> onclick="javascript:iva0()" />
            <!--</td>-->
            <%
            CloseDiv

            if si_tiene_modulo_SANTOS<>0 then
                DrawDiv "1","",""
                DrawLabel "","",LITCCONTPAGOVENC
                DrawInput "", "", "CCONTABLE_PAGOVENC", EncodeForHtml(rst("CCONTABLE_PAGOVENC")), "size='25'"
                CloseDiv
            end if
        'CloseFila
        if nz_b2(rst("INVSUJETOPASIVO"))=1 or nz_b2(rst("intra"))=1 then
            %><script>iva0();</script><%
        end if
        %>
            </table>
        </div>

        </div>

  	<%'DATOS BANCARIOS MODO EDIT %>
    <div class="Section" id="S_EditDB">
        <a href="#" rel="toggle[EditDB]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitDatosBancarios%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="EditDB">
        <table width="100%" border='0' cellpadding="2" cellspacing="2">
            <%DrawDiv "1","",""
            DrawLabel "","",LitBanco
            DrawInput "", "", "banco", EncodeForHtml(rst("banco")), "size='33'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitBancoDom
            DrawInput "", "", "banco_dom", EncodeForHtml(rst("banco_dom")), "size='35'"
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitNCuenta
            if rst("ncuenta")&"">"" then
                if len(rst("ncuenta"))=24 then
                    iban=mid(rst("ncuenta"),3,2)
                    pais=left(rst("ncuenta"),2)
                    strBanco = Mid(rst("ncuenta"), 5, 4)
			        strOficina = Mid(rst("ncuenta"), 9, 4)
			        strDC = Mid(rst("ncuenta"), 13, 2)
			        strCuenta = Mid(rst("ncuenta"), 15)
                else
                    strBanco = Mid(rst("ncuenta"), 1, 4)
			        strOficina = Mid(rst("ncuenta"), 5, 4)
			        strDC = Mid(rst("ncuenta"), 9, 2)
			        strCuenta = Mid(rst("ncuenta"), 11)
                end if
            end if
            
            if rst("ncuenta")&"" > "" and not isnumeric(left(rst("ncuenta"),2))   then
                pais = left(rst("ncuenta"),2)
                iban = mid(rst("ncuenta"),3,2)
                strCuenta = right(rst("ncuenta"),len(rst("ncuenta"))-4)
            else
                pais = ""
                iban = ""
                strCuenta = rst("ncuenta")&""
            end if
            
            %>  <input class='CELDA' type="text" name="country" value='<%=EncodeForHtml(pais)%>' maxlength="2" size="2"  onkeyup="if (this.value.length==2) document.proveedores.iban.focus()" onblur="this.value=this.value.toUpperCase();"/>
                <input class='CELDA' type="text" name="iban"    value='<%=EncodeForHtml(iban)%>' maxlength="2" size="2"  onkeyup="if (this.value.length==2) document.proveedores.ncuenta.focus()"/>
                <input class='CELDA' type="text" name="ncuenta" value='<%=EncodeForHtml(strCuenta)%>' maxlength="28" size="20" /><%
            CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitBICSWIFT
            DrawInput "", "", "bic", EncodeForHtml(rst("swift_code")), "maxlength='11' size='11'"
            CloseDiv

			rstSelect.cursorlocation=3
            strselect = "select distinct ncuenta from bancos with(nolock) where nbanco like ?+'%'"
            set command = nothing
            set conn = Server.CreateObject("ADODB.Connection")
            set command = Server.CreateObject("ADODB.Command")
            conn.Open = session("dsn_cliente")
            conn.CursorLocation = 3
            command.ActiveConnection = conn
            command.CommandTimeout = 60
            command.CommandText = strselect
            command.CommandType = adCmdText
            command.Parameters.Append command.CreateParameter("@nbanco",adVarChar,adParamInput,10,session("ncliente")&"")
            set rstSelect = command.Execute
			DrawSelectCelda "CELDA","175","",0,LitNCuentaCargo,"ncuentacargo",rstSelect,rst("cuenta_cargo"),"ncuenta","ncuenta","",""
			conn.Close
            set conn = nothing
            set command = nothing

            DrawDiv "1","",""
            DrawLabel "","",LitDomiciliacion
            DrawCheck "", "", "Domiciliacion", rst("domrec")
            CloseDiv%>
        </table>
        </div>
    </div>

    <% 'OTROS DATOS MODO EDIT %>
    <div class="Section" id="S_EditOD">
        <a href="#" rel="toggle[EditOD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitOtrosDatos%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="EditOD">
        <table width="100%" border='0' cellpadding="2" cellspacing="2">
		<%'jcg
        if si_tiene_modulo_proyectos<>0 then
                DrawDiv "1","",""
                DrawLabel "","",LitProyecto%>
                    <input class="CELDA" type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(rst("proyecto"))%>">
                    <iframe id='Iframe1' src='../mantenimiento/docproyectos.asp?viene=proveedores&mode=<%=EncodeForHtml(mode)%>&cod_proyecto=<%=EncodeForHtml(rst("proyecto"))%>' width='250' height='30' frameborder="no" scrolling="no" noresize="noresize"></iframe>
				<%CloseDiv
        end if
        strselect = "select codigo, descripcion from tipo_actividad with(nolock) where codigo like ?+'%'"
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
        set rstSelect = command.Execute
		DrawSelectCelda "CELDA",200,"",0,LitTActividad,"tactividad",rstSelect,rst("tactividad"),"codigo","descripcion","",""
		conn.Close
        set conn = nothing
        set command = nothing

        strselect = "select codigo, descripcion from tipos_entidades with(nolock) where tipo=? and codigo like ?+'%'"
        set command = nothing
        set conn = Server.CreateObject("ADODB.Connection")
        set command = Server.CreateObject("ADODB.Command")
        conn.Open = session("dsn_cliente")
        conn.CursorLocation = 3
        command.ActiveConnection = conn
        command.CommandTimeout = 60
        command.CommandText = strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@tipo",adVarChar,adParamInput,20,LitPROVEEDOR&"")
        command.Parameters.Append command.CreateParameter("@codigo",adVarChar,adParamInput,10,session("ncliente")&"")
        set rstSelect = command.Execute
		DrawSelectCelda "CELDA",200,"",0,LitTProveedor,"tipo_proveedor",rstSelect,rst("tipo_proveedor"),"codigo","descripcion","",""
		conn.Close
        set conn = nothing
        set command = nothing

        DrawDiv "1","",""
        DrawLabel "","",LitTransportista
        DrawInput "", "", "transportista", EncodeForHtml(rst("transportista")), "size='25'"
        CloseDiv

        DrawDiv "1","",""
        DrawLabel "","",LitPortes%><select class='width20' name="portes">
		    <%if rst("portes")=LitDebidos then%>
                <option selected value="<%=LitDebidos%>"><%=LitDebidos%></option>
                <option value="<%=LitPagados%>"><%=LitPagados%></option>
                <option value=""></option>
            <%elseif rst("portes")=LitPagados then%>
                <option value="<%=LitDebidos%>"><%=LitDebidos%></option>
                <option selected value="<%=LitPagados%>"><%=LitPagados%></option>
                <option value=""></option>
            <%else%>
                <option value="<%=LitDebidos%>"><%=LitDebidos%></option>
                <option value="<%=LitPagados%>"><%=LitPagados%></option>
                <option selected value=""></option>
            <%end if%>
		    </select><%
        CloseDiv
		rst.close
		%></table>
        </div>
    </div>

    <% 'CONFIG DOC MODO EDIT %>
    <div class="Section" id="S_EditCD">
        <a href="#" rel="toggle[EditCD]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitConfDoc2%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="EditCD">
        <table width="100%" bgcolor="<%=color_blau%>" border="0" cellpadding="2" cellspacing="2">
        <%
        nprov=limpiaCadena(request.querystring("nproveedor"))
		rstSelect.cursorlocation=3
        strselect = "select * from documentos_pro where ncliente=? and nproveedor=?"
        set command = nothing
        set conn = Server.CreateObject("ADODB.Connection")
        set command = Server.CreateObject("ADODB.Command")
        conn.Open = session("dsn_cliente")
        conn.CursorLocation = 3
        command.ActiveConnection = conn
        command.CommandTimeout = 60
        command.CommandText = strselect
        command.CommandType = adCmdText
        command.Parameters.Append command.CreateParameter("@ncliente",adVarChar,adParamInput,5,session("ncliente")&"")
        command.Parameters.Append command.CreateParameter("@nproveedor",adVarChar,adParamInput,10,nprov&"")
        set rstSelect = command.Execute

		    if (rstSelect.EOF=true) then
		        dato1=false
		        dato2=""
		    else
		        dato1=rstSelect("valorado_ped")
		        dato2=rstSelect("serie_ped")
		    end if

			DrawDiv "1","",""
            DrawLabel "","",LitValorPed
            DrawCheck "", "", "valorado_ped", iif(valorado_ped>"",nz_b(valorado_ped),nz_b(dato1))
            CloseDiv
			rstAux.cursorlocation=3
			
            strselect = "select nserie,(right(nserie,len(nserie)-5) + ' - ' + nombre) as nombre from series with(nolock) where nserie like ?+'%' and tipo_documento=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nserie",adVarChar,adParamInput,8,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo_documento",adVarChar,adParamInput,50,"PEDIDO A PROVEEDOR")
            set rstAux = command2.Execute

	 		DrawSelectCelda "CELDA","200","",0,LitSeriePed,"serie_ped",rstAux,dato2,"nserie","nombre","",""
	 		conn2.close
            set conn2    =  nothing
            set command2 =  nothing

		    if (rstSelect.EOF=true) then
		        dato1=false
		        dato2=""
		    else
		        dato1=rstSelect("valorado_alb")
		        dato2=rstSelect("serie_alb")
		    end if

            DrawDiv "1","",""
            DrawLabel "","",LitValorAlb
            DrawCheck "", "", "valorado_alb", iif(valorado_alb>"",nz_b(valorado_alb),nz_b(dato1))
            CloseDiv

			rstAux.cursorlocation=3

            strselect = "select nserie,(right(nserie,len(nserie)-5) + ' - ' + nombre) as nombre from series with(nolock) where nserie like ?+'%' and tipo_documento=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nserie",adVarChar,adParamInput,10,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo_documento",adVarChar,adParamInput,50,"ALBARAN DE PROVEEDOR")
            set rstAux = command2.Execute

			DrawSelectCelda "CELDA","200","",0,LitSerieAlb,"serie_alb",rstAux,dato2,"nserie","nombre","",""
			conn2.close
            set conn2    =  nothing
            set command2 =  nothing

		    if (rstSelect.EOF=true) then
		        dato1=""
		    else
		        dato1=rstSelect("serie_fac")
		    end if

			rstAux.cursorlocation=3

            strselect = "select nserie,(right(nserie,len(nserie)-5) + ' - ' + nombre) as nombre from series with(nolock) where nserie like ?+'%' and tipo_documento=?"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nserie",adVarChar,adParamInput,10,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@tipo_documento",adVarChar,adParamInput,50,"FACTURA DE PROVEEDOR")
            set rstAux = command2.Execute

			DrawSelectCelda "CELDA","200","",0,LitSerieFac,"serie_fac",rstAux,dato1,"nserie","nombre","",""
			conn2.close
            set conn2    =  nothing
            set command2 =  nothing%>
        </table>
        </div>
    </div>

    <% 'CAMPOS PERSONALIZABLES MODO EDIT %>
    <%if si_campo_personalizables=1 then%>
    <div class="Section" id="S_EditCP">
        <a href="#" rel="toggle[EditCP]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
        <div class="SectionHeader">
            <%=LitCampPersoPro%>
            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> title=""/>
        </div></a>

        <div class="SectionPanel" style="display:none " id="EditCP">
	    <table width="100%" border='0' cellpadding="2" cellspacing="2"><%

            strselect = "select * from camposperso with(nolock) where tabla='PROVEEDORES' and ncampo like ?+'%' order by ncampo,titulo"
            set command3 = nothing
            set conn3 = Server.CreateObject("ADODB.Connection")
            set command3 = Server.CreateObject("ADODB.Command")
            conn3.Open = session("dsn_cliente")
            conn3.CursorLocation = 3
            command3.ActiveConnection = conn3
            command3.CommandTimeout = 60
            command3.CommandText = strselect
            command3.CommandType = adCmdText
            command3.Parameters.Append command3.CreateParameter("@ncampo",adVarChar,adParamInput,7,session("ncliente")&"")
            set rst2 = command3.Execute

		    if not rst2.eof then
			    num_campos_existen=rst2.recordcount
			    'DrawFila ""
				    num_campo=1
				    num_campo2=1
				    num_puestos=0
				    num_puestos2=0
				    while not rst2.eof
					    if num_puestos2>0 and (num_puestos2 mod 2)=0 then
						    num_puestos2=0
					    end if
					    if rst2("titulo") & "">"" then
						    if ((num_puestos-1) mod 2)=0 then
						    end if
						    num_puestos=num_puestos+1
						    num_puestos2=num_puestos2+1

                            DrawDiv "1","",""
                            DrawLabel "","",EncodeForHtml(rst2("titulo"))

						    valor_campo_perso=lista_valores(num_campo)
						    if rst2("tipo")=1 then
							    if isNumeric(rst2("tamany")) then
								    tamany=rst2("tamany")
							    else
								    tamany=1
							    end if
							    %><input type="text" class="" name="<%="campo" & EncodeForHtml(num_campo)%>" size="35" maxlength="<%=tamany%>" value="<%=EncodeForHtml(valor_campo_perso)%>"> 
							    <%
                                CloseDiv
						    elseif rst2("tipo")=2 then
						        DrawCheck "", "", "campo" & num_campo,iif(valor_campo_perso=1,-1,0)
                                CloseDiv
                            elseif rst2("tipo")=3 then
							    num_campo_str=cstr(num_campo)
							    if len(num_campo_str)=1 then
								    num_campo_str="0" & num_campo_str
							    end if
							    

                                strselect = "select ndetlista,valor from campospersolista where tabla=? and ncampo=? and valor is not null and valor<>'' order by valor,ndetlista"
                                set command2 = nothing
                                set conn2 = Server.CreateObject("ADODB.Connection")
                                set command2 = Server.CreateObject("ADODB.Command")
                                conn2.Open = session("dsn_cliente")
                                conn2.CursorLocation = 3
                                command2.ActiveConnection = conn2
                                command2.CommandTimeout = 60
                                command2.CommandText = strselect
                                command2.CommandType = adCmdText
                                command2.Parameters.Append command2.CreateParameter("@tabla",adVarChar,adParamInput,50,"PROVEEDORES")
                                command2.Parameters.Append command2.CreateParameter("@ncampo",adVarChar,adParamInput,7,session("ncliente")&num_campo_str&"")
                                set rstAux = command2.Execute%>
								    <select class="" name="campo<%=EncodeForHtml(num_campo)%>" style='width:175px'>
									    <%
									    encontrado=0
									    while not rstAux.eof
										    if valor_campo_perso & "">"" and isnumeric(valor_campo_perso) then
											    valor_campo_perso_aux=cint(valor_campo_perso)
										    else
											    valor_campo_perso_aux=0
										    end if
										    if valor_campo_perso_aux=cint(rstAux("ndetlista")) then
											    texto_selected="selected"
											    if encontrado=0 then encontrado=1
										    else
											    texto_selected=""
										    end if
										    %>
										    <option value="<%=EncodeForHtml(rstAux("ndetlista"))%>"  <%=EncodeForHtml(texto_selected)%> ><%=EncodeForHtml(rstAux("valor"))%></option>
										    <%rstAux.movenext
									    wend
                                        %>
									    <option <%=iif(encontrado=1,"","selected")%> value=""></option>
								    </select>
                                <%
							    conn2.close
                                set conn2    =  nothing
                                set command2 =  nothing
                                CloseDiv
						    elseif rst2("tipo")=4 then
							    if isNumeric(rst2("tamany")) then
								    tamany=rst2("tamany")
							    else
								    tamany=1
							    end if
							    %><input type="text" class="" name="<%="campo" & EncodeForHtml(num_campo)%>" size="35" maxlength="<%=EncodeForHtml(tamany)%>" value="<%=EncodeForHtml(valor_campo_perso)%>">
							    <%
                                CloseDiv
						    elseif rst2("tipo")=5 then
							    if isNumeric(rst2("tamany")) then
								    tamany=rst2("tamany")
							    else
								    tamany=1
							    end if
							    %><input type="text" class="" name="<%="campo" & EncodeForHtml(num_campo)%>" size="35" maxlength="<%=EncodeForHtml(tamany)%>" value="<%=EncodeForHtml(valor_campo_perso)%>">
							    <%CloseDiv
						    end if
					    else
						    %><input type="hidden" name="campo<%=EncodeForHtml(num_campo)%>" value=""><%
					    end if
					    %><input type="hidden" name="tipo_campo<%=EncodeForHtml(num_campo)%>" value="<%=EncodeForHtml(rst2("tipo"))%>">    
					    <input type="hidden" name="titulo_campo<%=EncodeForHtml(num_campo)%>" value="<%=EncodeForHtml(rst2("titulo"))%>"><%
					    rst2.movenext
					    num_campo=num_campo+1
					    if not rst2.eof then
						    if rst2("titulo") & "">"" then
							    num_campo2=num_campo2+1
						    end if
					    end if
				    wend
			    num_campos=num_puestos
		    else
			    num_campos=0
			    num_campos_existen=0
		    end if
		    conn3.close
            set conn3    =  nothing
            set command3 =  nothing

	    %>
	    </table>
	    <input type="hidden" name="num_campos" value="<%=EncodeForHtml(num_campos_existen)%>" /> 
        </div>
    </div>
    <%end if%>

</td>
</tr>
</table>

<%end if %>
</form>
<%end if
   set rst = nothing
   set rst2 = nothing
   set rst3 = nothing
   set rst4 = nothing
   set rstAux = nothing
   set rstAux2 = nothing
   set rstSelect = nothing
   set rstDomi = nothing
   set rstDomi2 = nothing
connRound.close
set connRound = Nothing%>
</body>
</html>
