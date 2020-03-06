<%@ Language=VBScript %>
<% Server.ScriptTimeout = 300 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>  


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
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="pedpro_albpro_param.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
    function tier1Menu(objMenu, objImage) {
        if (objMenu.style.display == "none") {
            objMenu.style.display = "";
            objImage.src = "../Images/<%=ImgCarpetaAbierta%>";
        }
        else {
            objMenu.style.display = "none";
            objImage.src = "../images/<%=ImgCarpetaCerrada%>";
        }
    }

    function tier2Menu(objMenu) {
        albaranes.style.display = "none";
        pedidos.style.display = "none";
        objMenu.style.display = "";
    }

    function tier3Menu(objImage) {
        if (objImage.src == "../images/<%=ImgCarpetaCerrada%>") objImage.src = "../Images/<%=ImgCarpetaAbierta%>";
        else objImage.src = "../images/<%=ImgCarpetaCerrada%>";
    }

    function abrir_detalles(ndoc, fila, viene) {
        pagina = "../central.asp?pag1=mantenimiento/conv_ped_alb_completar.asp&ndoc=" + ndoc + "&tdocumento=" + fila + "&viene=compras&s=pedpro_facpro&pag2=mantenimiento/conv_ped_alb_completar_bt.asp";
        AbrirVentana(pagina, 'P', 400, 750);
    }

    function abrir_detallesEBESA(ndoc, fila, viene) {
        pagina = "../central.asp?pag1=mantenimiento/conv_ped_alb_completar_E.asp&ndoc=" + ndoc + "&tdocumento=" + fila + "&viene=devolucionPro&s=pedpro_facpro&pag2=mantenimiento/conv_ped_alb_completar_bt_E.asp";
        AbrirVentana(pagina, 'P', 400, 750);
    }

    //Desencadena la búsqueda del proveedor cuyo numero se indica
    function TraerProveedor() {
        //document.location="albpedpro_facpro_param.asp?nproveedor=" + document.albpedpro_facpro_param.nproveedor.value + "&mode=select1&fdesde=" + document.albpedpro_facpro_param.fdesde.value + "&fhasta=" + document.albpedpro_facpro_param.fhasta.value + "&tabla=" + document.albpedpro_facpro_param.htabla.value;
        document.albpedpro_facpro_param.action = "albpedpro_facpro_param.asp?nproveedor=" + document.albpedpro_facpro_param.nproveedor.value + "&mode=select1&fdesde=" + document.albpedpro_facpro_param.fdesde.value + "&fhasta=" + document.albpedpro_facpro_param.fhasta.value + "&tabla=" + document.albpedpro_facpro_param.htabla.value;
        document.albpedpro_facpro_param.submit();
    }

    function seleccionar() {
        nregistros = document.albpedpro_facpro_param.h_nfilas.value;
        if (document.albpedpro_facpro_param.check.checked)
        {
            for (i = 1; i <= nregistros; i++)
            {
                nombre = "hace_falta_lote" + i;
                nombre2 = "hace_falta_serie" + i;
                nombre3 = "check" + i;
                //if (document.albpedpro_facpro_param.elements[nombre].value==0 && document.albpedpro_facpro_param.elements[nombre2].value==0)
                if (document.albpedpro_facpro_param.elements[nombre3].disabled==false)
                {
                    nombre = "check" + i;
                    document.albpedpro_facpro_param.elements[nombre].checked = true;
                }
            }
            document.albpedpro_facpro_param.check.value = "yyy"
        }
        else
        {
            for (i = 1; i <= nregistros; i++)
            {
                nombre = "hace_falta_lote" + i;
                nombre2 = "hace_falta_serie" + i;
                nombre3 = "check" + i;
                //if (document.albpedpro_facpro_param.elements[nombre].value==0 && document.albpedpro_facpro_param.elements[nombre2].value==0)
                if (document.albpedpro_facpro_param.elements[nombre3].disabled==false)
                {
                    nombre = "check" + i;
                    document.albpedpro_facpro_param.elements[nombre].checked = false;
                }
            }
            document.albpedpro_facpro_param.check.value = "xxx"
        }
    }
</script>
<body onload="self.status='';" bgcolor=<%=color_blau%>>
<%'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
	'fdesde:
	'fhasta:
	'nproveedor:
function CadenaBusqueda(fdesde,fhasta,nproveedor)
	if alb = "true" then
	    seriea=request.Form("seriea")
	    cadserie=""
	    if seriea&"">"" then
	        cadserie=" and serie='"&seriea&"' "
	    end if
	    
		if nproveedor > "" then
			CadenaBusqueda = " nproveedor='" & nproveedor & "' and fecha>='" & fdesde & "' and fecha<='" & fhasta & "' "&cadserie&"  "
		else
			CadenaBusqueda = " fecha>='" & fdesde & "' and fecha<='" & fhasta & "' "&cadserie&" "
		end if
	else
	    seriep=request.Form("seriep")
	    cadserie=""
	    if seriep&"">"" then
	        cadserie=" and serie='"&seriep&"' "
	    end if
	    
		if nproveedor > "" then
			CadenaBusqueda = " nproveedor='" & nproveedor & "' and fecha>='" & fdesde & "' and fecha<='" & fhasta & "' "&cadserie&" and nalbaran is null and nfactura is null "
		else
			CadenaBusqueda = " fecha>='" & fdesde & "' and fecha<='" & fhasta & "' "&cadserie&" and nfactura is null and nalbaran is null "
		end if
	end if
    'Lo movemos
	'if centro>"" then
	'    CadenaBusqueda = CadenaBusqueda & " and tienda='"&centro&"'"
	'end if
end function

function Limpiar(cadena)
	dim temp
	temp=cadena
	temp=Reemplazar(temp,"/","")
	temp=Reemplazar(temp,"-","")
	temp=Reemplazar(temp,"_","")
	temp=Reemplazar(temp,"$","")
	temp=Reemplazar(temp,"%","")
	temp=Reemplazar(temp,":","")
	Limpiar=temp
end function

'******************************************************************************
'Validar las lineas con numeros de serie de los pedidos
function validarNSerie(tabla)
	if tabla="pedidos_pro" then
		tabla_det="detalles_ped_pro"
		ncampo="npedido"
		mens_err_VN="Pedido"
	elseif tabla="albaranes_pro" then
		tabla_det="detalles_alb_pro"
		ncampo="nalbaran"
		mens_err_VN="Albaran"
	end if

	seleccion="select * from " & tabla_det & " as d with(nolock) inner join articulos a with(nolock) on a.referencia=d.referencia and a.referencia like '"&session("ncliente")&"%' and a.ctrl_nserie=1"
	seleccion=seleccion & "," & tabla & " as p with (nolock) where " & strwhere
	'seleccion=seleccion & " and d.referencia in (select referencia from articulos where ctrl_nserie=1) "
	seleccion=seleccion & " and d." & ncampo & "=p." & ncampo & " and d." & ncampo & " like '"&session("ncliente")&"%' and p." & ncampo & " like '"&session("ncliente")&"%' "
	seleccion=seleccion & " order by p.fecha desc,d." & ncampo & " desc,d.item asc"

	if tabla_det & "">"" and ncampo & "">"" then
		Correcto=true
		rst.Open seleccion, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		while not rst.EOF
			if tabla="pedidos_pro" then
				
				rstAux2.Open "select nserie from egesticet.[" & session("usuario") & "] where ndocumento='" & rst("npedido") & "' and item = " & rst("item"),session("dsn_cliente")
                if not rstAux2.eof then
                    listaTEquipos=rstAux2("nserie")
                    if listaTEquipos & "">"" then
                        listaequipos=split(listaTEquipos,",")
                        p=ubound(listaequipos)+2
                    else
                        p=0
                    end if
                end if
                
                rstAux2.Close
			else
				listaTEquipos=rst("nserie")
				p=2
			end if

			mensajeTratEquipos="OK1"
			if tabla="pedidos_pro"  then
				if rst("cantidadPend")<>0 then
					mensajeTratEquipos=TratarEquipos(listaTEquipos,p-1,"PED-FAC-PRO",rst("npedido"),rst("item"),rst("referencia"),rst("cantidadPend"),rst("almacen"),"","first_save")
				end if
			elseif tabla="albaranes_pro" and listaTEquipos & ""="" then
				mensajeTratEquipos=LitMsgErrCantNSerie
			end if
			'EBF Insertamos los numeros de serie en la tabla temporal
			if mid(mensajeTratEquipos,1,2)<>"OK" then
				Correcto=false
				if tabla="pedidos_pro" then
					%><SCRIPT language="JavaScript">
						window.alert("<%=mensajeTratEquipos%>\n<%=mens_err_VN & ": "%><%=trimCodEmpresa(rst("npedido"))%> <%=LitEnLaLinea%>: <%=rst("item")%>");
					</script><%
				else
					%><SCRIPT language="JavaScript">
						window.alert("<%=mensajeTratEquipos%>\n<%=mens_err_VN & ": "%><%=rst("nalbaran_pro")%> <%=LitEnLaLinea%>: <%=rst("item")%>");
					</script><%
				end if
				rst.movelast
			'elseif listaTEquipos>"" then
''ricardo 1/8/2006 se añade la actualizacion del campo mi_nserie
			'	rstAux.open "update [" & session("usuario") & "] set nserie='"&replace(listaTEquipos,"'","''")&"',mi_nserie='" & replace(listaTEquipos,"'","''") & "' where ndocumento='"&rst(ncampo)&"' and item="&rst("item") , session("dsn_cliente")
			end if
			rst.movenext
		wend
		
		rst.close
	else
		Correcto=false
	end if
	validarNSerie=Correcto
end function

'******************************************************************************
'Valida los descuentos de los pedidos
function validar_pedidos()
    validar_pedidos=true
	'rst.open "delete from [" & session("usuario") & "] ",session("dsn_cliente")
	rst.open "select count(*) from [" & session("usuario") & "] ",session("dsn_cliente")
	total = rst(0)
	rst.close
	strwhere ="("
	while h_nfilas>0
		npedido=limpiaCadena(request.form("check"+cstr(h_nfilas)))
		checkCadena(npedido)
		if npedido>"" and total =0 then
			rst.open "insert into [" & session("usuario") & "] (ndocumento,item,cantidad,referencia,almacen) select npedido,item,cantidadPend,referencia,almacen from detalles_ped_pro with(nolock) where npedido='"&npedido&"' ",session("dsn_cliente")
			rst.open "insert into [" & session("usuario") & "] (ndocumento) select npedido from pedidos_pro with(nolock) where npedido='"&npedido&"' and npedido not in (select npedido from detalles_ped_pro with(nolock) where npedido='"&npedido&"') ",session("dsn_cliente")
            npedidos=npedidos+1
		end if
		strwhere = strwhere +"'"+npedido+"',"
		h_nfilas = h_nfilas -1
	wend
	strwhereAux=strwhere
	strwhere = "p.npedido in "+strwhereAux+"'xxxxxx')"
	strwhere2 = "p.npedido in "+strwhereAux+"'xxxxxx')"

	''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
	unidoccompdto=nz_b(d_lookup("unidoccompdto","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))

	StrCAbecera="select distinct divisa"
	if unidoccompdto=0 then
		StrCAbecera=StrCAbecera & ",descuento,descuento2"
	end if
	StrCAbecera=StrCAbecera & ",recargo,irpf from pedidos_pro as p with(nolock) where " + strwhere

	rst_cabecera.cursorlocation=3
	rst_cabecera.Open StrCAbecera, session("dsn_cliente") '',adUseClient, adLockReadOnly
	if rst_cabecera.recordCount <> 1 then
		%><script language="javascript">
		      window.alert("<%=LitMsgPedAlbComprasIncong%>");
		</script><%
		strwhere = "p.npedido in ('xxxxxx')"
		validar_pedidos=false
	end if
	rst_cabecera.Close

	if validar_pedidos=true then
		rst_cabecera.Open "select distinct irpf_total from pedidos_pro as p with(nolock) where " & strwhere2, session("dsn_cliente"),adUseClient, adLockReadOnly
		if rst_cabecera.recordCount <> 1 then
			%><script language="javascript">
			      window.alert("<%=LitMsgPedAlbComprasIncong2%>");
			</script><%
			strwhere2 = "p.npedido in ('xxxxxx')"
			validar_pedidos=false
		end if
		rst_cabecera.Close
	end if

	if validar_pedidos=true then
		validar_pedidos=validar_albped_vencimiento("",1,"comprasP","")
	end if
end function

'******************************************************************************
'Convierte los pedidos en facturas
function convertir_pedidos()
    ''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
    unidoccompdto=nz_b(d_lookup("unidoccompdto","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))

	strselect="select * from pedidos_pro as p where " + strwhere
	if strwhere & "">"" then
		strselect=strselect & " and nalbaran is null and nfactura is null"
	else
		strselect=strselect & " nalbaran is null and nfactura is null"
	end if
	rst_cabecera.Open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic

	if not rst_cabecera.eof then
		forma_pago=rst_cabecera("forma_pago")
		divisa=iif(rst_cabecera("divisa")>"",rst_cabecera("divisa"),d_lookup("codigo","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")))
		descuento=null_z(rst_cabecera("descuento"))
		descuento2=null_z(rst_cabecera("descuento2"))
		cod_proyecto=null_s(rst_cabecera("cod_proyecto"))
		nerror=0
		nerror2=0
		nerror3=0
		nerror4=0
		lista_pedidos="('"
		while not rst_cabecera.EOF and nerror=0 and nerror2=0 and nerror3=0 and nerror4=0
			if rst_cabecera("forma_pago")<>forma_pago or (isnull(rst_cabecera("forma_pago")) and forma_pago>"") or (rst_cabecera("forma_pago")>"" and isnull(forma_pago)) then
				nerror=1
			end if
			if rst_cabecera("divisa")<>divisa or (isnull(rst_cabecera("divisa")) and divisa>"") or (rst_cabecera("divisa")>"" and isnull(divisa))then
				nerror2=1
			end if

			''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
			if unidoccompdto=0 then
				if rst_cabecera("descuento")<>descuento or (isnull(rst_cabecera("descuento")) and descuento>"") or (rst_cabecera("descuento")>"" and isnull(descuento))then
					nerror3=1
				end if
			end if
			''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
			if unidoccompdto=0 then
				if rst_cabecera("descuento2")<>descuento2 or (isnull(rst_cabecera("descuento2")) and descuento2>"") or (rst_cabecera("descuento2")>"" and isnull(descuento2))then
					nerror3=1
				end if
			end if
			if ucase(rst_cabecera("cod_proyecto"))<>ucase(cod_proyecto) or (isnull(rst_cabecera("cod_proyecto")) and cod_proyecto>"") or (rst_cabecera("cod_proyecto")>"" and isnull(cod_proyecto)) then
				nerror4=1
			end if
			lista_pedidos=lista_pedidos & trimcodempresa(rst_cabecera("npedido")) & "','"
			rst_cabecera.MoveNext
		wend
		if lista_pedidos="('" then
			lista_pedidos=""
		else
			lista_pedidos=mid(lista_pedidos,1,len(lista_pedidos)-2) & ")"
		end if
		if nerror=1 or nerror2=1 or nerror3=1 or nerror4=1 then
			rst_cabecera.Close
			if nerror=1 then
				%><script language="javascript">
				      window.alert("<%=LitMsgPedAlbComprasNoConv3%>");
				</script><%
			end if
			if nerror2=1 then%>
				<script language="javascript">
				    window.alert("<%=LitMsgPedAlbComprasNoConv4%>");
				</script><%
			end if
			if nerror3=1 then
				%><script language="javascript">
				      window.alert("<%=LitMsgPedAlbComprasNoConv5%>");
				</script><%
			end if
			if nerror4=1 then
				%><script language="javascript">
				      window.alert("<%LitMsgPedAlbComprasNoConv6%>");
				</script><%
			end if
			convertir_pedidos=""
		else
			rst_cabecera.movefirst
			nfactura = ""
			ncuentaAnt=""
			variasCuentas=0
			IncotermsDistintos=CalcularIncotermsDistintos(lista_pedidos,")","pedidos_pro",rst_cabecera("nproveedor"),0,1)
			FobDistintos=CalcularFobDistintos(lista_pedidos,")","pedidos_pro",rst_cabecera("nproveedor"),0,1)
			'gen_vencimiento=d_lookup("gen_vencimientos","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente"))
			'EBF convertimos a masiva la conversion de albaranes a facturas
			strselect="EXEC ConvertirPedAlbDev_Fac @serie='"&nserie&"',@fechaFactura='"&ffactura&"',@nfactura_pro='"&nfactura_pro&"',@nusuario='"&session("usuario")&"',@session_ncliente='"&session("ncliente")&"',@tipo='PEDIDO',@lista='"&replace(replace(replace(lista_pedidos,"'",""),"(",""),")","")&"',@modulo=''"
	        set conn2 = Server.CreateObject("ADODB.Connection")
	        set command2 =  Server.CreateObject("ADODB.Command")

	        conn2.open session("dsn_cliente")			
	        command2.ActiveConnection =conn2
	        command2.CommandTimeout = 90
	        conn2.cursorlocation=3
	        command2.CommandText=strselect
   	        set rstAux=Command2.Execute			
			
			Primero=rstAux(0)
			Ultimo=rstAux(1)
			error=rstAux(2)
			cadenaAuditoria=rstAux(3)

			if Primero>"" and Ultimo>"" and error=0 then
				nfactura_pro=d_lookup("nfactura_pro","facturas_pro","nfactura='"&Primero&"'",session("dsn_cliente"))
				%><SCRIPT>
					window.alert("<%=LitMsgFacturaGenerado%> <%=nfactura_pro%> ");
					if (window.confirm("<%=LitMsgDeseaVer%>")==true)
						AbrirVentana('../search_layout.asp?pag1=compras/facturas_pro.asp?ndoc=<%=primero%>?mode=browse?titulo=<%=LitDetallesFac%> <%=nfactura_pro%>&pag2=compras/facturas_pro_bt.asp&pag3=compras/facturas_pro_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
				</script><%
			elseif error=1 then
				%><SCRIPT>
				      window.alert("<%=LitMsgSinDireccionPrinc%>");
				</script><%
			elseif error=2 then
				%><SCRIPT>
				      window.alert("<%=LitMsgNumSeriesRepetidos%>");
				</script><%
			elseif error=3 then
				%><SCRIPT>
				      window.alert("<%=LitMsgDocExistRevContConv%>");
				</script><%
			else
				%><SCRIPT>
				      window.alert("<%=LitMsgFacturaNoGenerado%>");
				</script><%
			end if
			rstAux.close
			conn2.close
	        set conn2=nothing		
	        set command2=nothing

			'**************************'
			' Ahora guardamos los pagos y los vencimientos'
			rst_cabecera.movefirst
			'**************************'

			rst_cabecera.Close
			nfactura_param = nfactura
			'vamos a auditar
			auditar_ins_bor session("usuario"),cadenaAuditoria,"","alta","","","conver_alb_ped_dev_fac_proMASIVO"			
			auditar_ins_bor session("usuario"),Primero,"","alta","","","equipos_conver_alb_ped_fac_pro"
			convertir_pedidos="OK"
		end if
	else
		rst_cabecera.Close
	end if
end function

'******************************************************************************
'Valida los descuentos de los albaranes
function validar_albaranes()
	validar_albaranes=true
    h_nfilas2=h_nfilas
	strwhere ="("
	strwhere2 = "("
	while h_nfilas>0
		nalbaran=limpiaCadena(request.form("check"+cstr(h_nfilas)))
		tipo=limpiaCadena(request.form("tipo"+cstr(h_nfilas)))
		checkCadena(nalbaran)
		if nalbaran>"" and tipo="ALB" then
			rst.open "insert into [" & session("usuario") & "] (ndocumento,item,cantidad,referencia) select nalbaran,item,cantidad,referencia from detalles_alb_pro where nalbaran='"&nalbaran&"' ",session("dsn_cliente")
			rst.open "insert into [" & session("usuario") & "] (ndocumento) select nalbaran from albaranes_pro where nalbaran='"&nalbaran&"' and nalbaran not in (select nalbaran from detalles_alb_pro where nalbaran='"&nalbaran&"') ",session("dsn_cliente")
			strwhere = strwhere +"'"+nalbaran+"',"
			strwhere2 = strwhere2 & "'" & d_lookup("nalbaran_pro","albaranes_pro","nalbaran='" & nalbaran & "'",session("dsn_cliente")) & "',"
			nalbaranes=nalbaranes+1
		end if
		h_nfilas = h_nfilas -1
	wend
	h_nfilas=h_nfilas2
	while h_nfilas>0
		nalbaran=limpiaCadena(request.form("check"+cstr(h_nfilas)))
		tipo=limpiaCadena(request.form("tipo"+cstr(h_nfilas)))
		checkCadena(nalbaran)
		if nalbaran>"" and tipo="DEV" then
			rst.open "insert into [" & session("usuario") & "_dev] (ndevolucion,item,cantidad,referencia) select ndevolucion,item,cantidad,referencia from detalles_dev_pro with(nolock) where ndevolucion='"&nalbaran&"' and ndevolucion not in( select ndevolucion from [" & session("usuario") & "_dev]) ",session("dsn_cliente")
			strwhere = strwhere +"'"+nalbaran+"',"
			nalbaranes=nalbaranes+1
		end if
		h_nfilas = h_nfilas -1
	wend

	strwhereAux=strwhere
	strwhereAux2=strwhere2
	strwhere2 = "ndocumento in " & strwhereAux2 & "'xxxxxx')"
	strwhere = "p.nalbaran in "+strwhereAux+"'xxxxxx')"
	strwhere2_2 = "ndocumento in " & strwhereAux2 & "'xxxxxx')"
	strwhere_2 = "p.nalbaran in "+strwhereAux+"'xxxxxx')"

	''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
	unidoccompdto=nz_b(d_lookup("unidoccompdto","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))

	StrCAbecera="select distinct divisa"
	if unidoccompdto=0 then
		StrCAbecera=StrCAbecera & ",descuento,descuento2"
	end if
	StrCAbecera=StrCAbecera & ",recargo,irpf from albaranes_pro as p where " + strwhere

	rst_cabecera.cursorlocation=3
	rst_cabecera.Open StrCAbecera, session("dsn_cliente") '',adUseClient, adLockReadOnly
	if rst_cabecera.recordCount > 1 then
		%><script language="javascript">
		      window.alert("<%=LitMsgPedAlbComprasIncong%>");
		</script><%
		strwhere2 = "ndocumento in ('xxxxxx')"
		strwhere = "p.nalbaran in ('xxxxxx')"
		validar_albaranes=false
	end if
	rst_cabecera.Close

	if validar_albaranes=true then
		rst_cabecera.Open "select distinct irpf_total from albaranes_pro as p where " + strwhere_2, session("dsn_cliente"),adUseClient, adLockReadOnly
		if rst_cabecera.recordCount > 1 then
			%><script language="javascript">
			      window.alert("<%=LitMsgPedAlbComprasIncong2%>");
			</script><%
			strwhere2_2 = "ndocumento in ('xxxxxx')"
			strwhere_2 = "p.nalbaran in ('xxxxxx')"
			validar_albaranes=false
		end if
		rst_cabecera.Close
	end if

	if validar_albaranes=true then
		validar_albaranes=validar_albped_vencimiento("",1,"comprasA","")
	end if
end function

'******************************************************************************
'Convierte los albaranes en facturas
function convertir_albaranes()
    ''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
    unidoccompdto=nz_b(d_lookup("unidoccompdto","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))

	strselect="select * from albaranes_pro as p with(nolock) where " + strwhere
	if strwhere & "">"" then
		strselect=strselect & " and p.nfactura is null"
	else
		strselect=strselect & " p.nfactura is null"
	end if
	rst_cabecera.Open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic

    rstAux.Open "select ndevolucion from devoluciones_pro p with(nolock) where "+ replace(strwhere,"nalbaran","ndevolucion") ,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
   
	if not rst_cabecera.eof or not rstAux.EOF then
	    if not rst_cabecera.EOF then
		    forma_pago=rst_cabecera("forma_pago")
		    divisa=iif(rst_cabecera("divisa")>"",rst_cabecera("divisa"),d_lookup("codigo","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")))
		    descuento=null_z(rst_cabecera("descuento"))
		    descuento2=null_z(rst_cabecera("descuento2"))
		    cod_proyecto=null_s(rst_cabecera("cod_proyecto"))
        end if
		nerror=0
		nerror2=0
		nerror3=0
		nerror4=0
		lista_albaranes="('"
		while not rst_cabecera.EOF and nerror=0 and nerror2=0 and nerror3=0 and nerror4=0

			if rst_cabecera("forma_pago")<>forma_pago or (isnull(rst_cabecera("forma_pago")) and forma_pago>"") or (rst_cabecera("forma_pago")>"" and isnull(forma_pago))then
				nerror=1
			end if
			if rst_cabecera("divisa")<>divisa or (isnull(rst_cabecera("divisa")) and divisa>"") or (rst_cabecera("divisa")>"" and isnull(divisa))then
				nerror2=1
			end if
			''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
			if unidoccompdto=0 then
				if rst_cabecera("descuento")<>descuento or (isnull(rst_cabecera("descuento")) and descuento>"") or (rst_cabecera("descuento")>"" and isnull(descuento))then
					nerror3=1
				end if
			end if
			''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
			if unidoccompdto=0 then
				if rst_cabecera("descuento2")<>descuento2 or (isnull(rst_cabecera("descuento2")) and descuento2>"") or (rst_cabecera("descuento2")>"" and isnull(descuento2))then
					nerror3=1
				end if
			end if
			if ucase(rst_cabecera("cod_proyecto"))<>ucase(cod_proyecto) or (isnull(rst_cabecera("cod_proyecto")) and cod_proyecto>"") or (rst_cabecera("cod_proyecto")>"" and isnull(cod_proyecto)) then
				nerror4=1
			end if
			lista_albaranes=lista_albaranes & rst_cabecera("nalbaran_pro") & "','"
			rst_cabecera.MoveNext
			'if not rst_cabecera.eof then
				'forma_pago=rst_cabecera("forma_pago")
				'divisa=iif(rst_cabecera("divisa")>"",rst_cabecera("divisa"),d_lookup("codigo","divisas","moneda_base=1",session("dsn_cliente")))
			'end if
		wend
		while not rstaux.EOF
		    lista_albaranes=lista_albaranes & trimcodempresa(rstAux("ndevolucion")) & "','"
		    rstAux.MoveNext 
		wend
		rstAux.Close
		
		if lista_albaranes="('" then
			lista_albaranes=""
		else
			lista_albaranes=mid(lista_albaranes,1,len(lista_albaranes)-2) & ")"
		end if
		if nerror=1 or nerror2=1 or nerror3=1 or nerror4=1 then
			rst_cabecera.Close
			if nerror=1 then
				%><script language="javascript">
				      window.alert("<%=LitMsgDistFormPago%>");
				</script><%
			end if
			if nerror2=1 then
				%><script language="javascript">
				      window.alert("LitMsgDistDivisa");
				</script><%
			end if
			if nerror3=1 then
				%><script language="javascript">
				      window.alert("LitMsgDistDescuentos");
				</script><%
			end if
			if nerror4=1 then
				%><script language="javascript">
				      window.alert("LitMsgDistProyectos");
				</script><%
			end if
			convertir_albaranes=""
		else
			nfactura = ""
			ncuentaAnt=""
			variasCuentas=0

		    if not rst_cabecera.EOF then
    			rst_cabecera.movefirst
	    		IncotermsDistintos=CalcularIncotermsDistintos(lista_albaranes,")","albaranes_pro",rst_cabecera("nproveedor"),0,1)
		    	FobDistintos=CalcularFobDistintos(lista_albaranes,")","albaranes_pro",rst_cabecera("nproveedor"),0,1)
            end if
            		    	
			'EBF convertimos a masiva la conversion de albaranes a facturas
			if si_tiene_modulo_EBESA<>0 then
			    strselect="EXEC ConvertirPedAlbDev_Fac @serie='"&nserie&"',@fechaFactura='"&ffactura&"',@nfactura_pro='"&nfactura_pro&"',@nusuario='"&session("usuario")&"',@session_ncliente='"&session("ncliente")&"',@tipo='ALBARAN',@lista='"&replace(replace(replace(lista_albaranes,"'",""),"(",""),")","")&"',@modulo='EBESA'"
			else
			    strselect="EXEC ConvertirPedAlbDev_Fac @serie='"&nserie&"',@fechaFactura='"&ffactura&"',@nfactura_pro='"&nfactura_pro&"',@nusuario='"&session("usuario")&"',@session_ncliente='"&session("ncliente")&"',@tipo='ALBARAN',@lista='"&replace(replace(replace(lista_albaranes,"'",""),"(",""),")","")&"',@modulo=''"
			end if
	        set conn = Server.CreateObject("ADODB.Connection")
	        set command =  Server.CreateObject("ADODB.Command")

	        conn.open session("dsn_cliente")			
	        command.ActiveConnection =conn
	        command.CommandTimeout = 90
	        conn.cursorlocation=3
	        command.CommandText=strselect
   	        set rstAux=Command.Execute

			Primero=rstAux(0)
			Ultimo=rstAux(1)
			error=rstAux(2)
			cadenaAuditoria=rstAux(3)

			if Primero>"" and Ultimo>"" and error=0 then
				nfactura_pro=d_lookup("nfactura_pro","facturas_pro","nfactura='"&Primero&"'",session("dsn_cliente"))
				%><SCRIPT>
					window.alert("<%=LitMsgFacturaGenerado%> <%=nfactura_pro%>");
					if (window.confirm("<%=LitMsgDeseaVer%>")==true)
						AbrirVentana('../search_layout.asp?pag1=compras/facturas_pro.asp?ndoc=<%=primero%>?mode=browse?titulo=<%=LitDetallesFac%> <%=nfactura_pro%>&pag2=compras/facturas_pro_bt.asp&pag3=compras/facturas_pro_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
				</script><%
			elseif error=1 then
				%><SCRIPT>
				      window.alert("<%=LitMsgSinDireccionPrinc%>");
				</script><%
			elseif error=2 then
				%><SCRIPT>
				      window.alert("<%=LitMsgNumSeriesRepetidos%>");
				</script><%
			elseif error=3 then
				%><SCRIPT>
				      window.alert("<%=LitMsgDocExistRevContConv%>");
				</script><%
			else
				%><SCRIPT>
				      window.alert("<%=LitMsgFacturaNoGenerado%>");
				</script><%
			end if
			rstAux.close
	        set command=nothing
	        set conn=nothing			
			'**************************'

			rst_cabecera.Close
			nfactura_param = nfactura
			'vamos a auditar
			auditar_ins_bor session("usuario"),Primero,"","alta","","","conver_alb_ped_dev_fac_proMASIVO"
			'no se audita los equipos, ya que no se crea ninguno, al estar ya creados por el albaran
			'auditar_ins_bor session("usuario"),nfactura,"","alta","","","equipos_conver_alb_ped_fac_pro"
			convertir_albaranes="OK"
		end if
	else
	    rstAux.Close
		rst_cabecera.Close
	end if
end function

function anyadir_albaranes(ndoc)
    ''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
    unidoccompdto=nz_b(d_lookup("unidoccompdto","configuracion","nempresa='" & session("ncliente") & "'",session("dsn_cliente")))

	strselect="select * from albaranes_pro as p with(nolock)  where " + strwhere
	if strwhere & "">"" then
		strselect=strselect & " and p.nfactura is null"
	else
		strselect=strselect & " p.nfactura is null"
	end if
	rst_cabecera.Open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic

	if not rst_cabecera.eof  then
		forma_pago=rst_cabecera("forma_pago")
		divisa=iif(rst_cabecera("divisa")>"",rst_cabecera("divisa"),d_lookup("codigo","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")))
		descuento=null_z(rst_cabecera("descuento"))
		descuento2=null_z(rst_cabecera("descuento2"))
		cod_proyecto=null_s(rst_cabecera("cod_proyecto"))
		nerror=0
		nerror2=0
		nerror3=0
		nerror4=0
		lista_albaranes="('"
		while not rst_cabecera.EOF and nerror=0 and nerror2=0 and nerror3=0 and nerror4=0
			if rst_cabecera("forma_pago")<>forma_pago or (isnull(rst_cabecera("forma_pago")) and forma_pago>"") or (rst_cabecera("forma_pago")>"" and isnull(forma_pago))then
				nerror=1
			end if
			if rst_cabecera("divisa")<>divisa or (isnull(rst_cabecera("divisa")) and divisa>"") or (rst_cabecera("divisa")>"" and isnull(divisa))then
				nerror2=1
			end if
			''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
			if unidoccompdto=0 then
				if rst_cabecera("descuento")<>descuento or (isnull(rst_cabecera("descuento")) and descuento>"") or (rst_cabecera("descuento")>"" and isnull(descuento))then
					nerror3=1
				end if
			end if
			''ricardo 20-2-2006 si en configuracion esta el campo unidoccompdto puesto a 1 dara igual como esten los descuentos
			if unidoccompdto=0 then
				if rst_cabecera("descuento2")<>descuento2 or (isnull(rst_cabecera("descuento2")) and descuento2>"") or (rst_cabecera("descuento2")>"" and isnull(descuento2))then
					nerror3=1
				end if
			end if
			if ucase(rst_cabecera("cod_proyecto"))<>ucase(cod_proyecto) or (isnull(rst_cabecera("cod_proyecto")) and cod_proyecto>"") or (rst_cabecera("cod_proyecto")>"" and isnull(cod_proyecto)) then
				nerror4=1
			end if
			lista_albaranes=lista_albaranes & rst_cabecera("nalbaran_pro") & "','"
			rst_cabecera.MoveNext
		wend
		if lista_albaranes="('" then
			lista_albaranes=""
		else
			lista_albaranes=mid(lista_albaranes,1,len(lista_albaranes)-2) & ")"
		end if
		
		if nerror=1 or nerror2=1 or nerror3=1 or nerror4=1 then
			rst_cabecera.Close
			if nerror=1 then
				%><script language="javascript">
				      window.alert("<%=LitMsgAlbFacComprasNoConv3%>");
				</script><%
			end if
			if nerror2=1 then
				%><script language="javascript">
				      window.alert("<%=LitMsgAlbFacComprasNoConv4%>");
				</script><%
			end if
			if nerror3=1 then
				%><script language="javascript">
				      window.alert("<%=LitMsgAlbFacComprasNoConv5%>");
				</script><%
			end if
			if nerror4=1 then
				%><script language="javascript">
				      window.alert("<%=LitMsgAlbFacComprasNoConv6%>");
				</script><%
			end if
			anyadir_albaranes=""
		else
			rst_cabecera.movefirst
			nfactura = ""
			ncuentaAnt=""
			variasCuentas=0
			IncotermsDistintos=CalcularIncotermsDistintos(lista_albaranes,")","albaranes_pro",rst_cabecera("nproveedor"),0,1)
			FobDistintos=CalcularFobDistintos(lista_albaranes,")","albaranes_pro",rst_cabecera("nproveedor"),0,1)
			'EBF convertimos a masiva la conversion de albaranes a facturas
            strselect="EXEC AnyadirAlbaranesFacPro @ndoc='" & ndoc & "', @nusuario='"&session("usuario")&"',@session_ncliente='"&session("ncliente")&"',@lista='"&replace(replace(replace(lista_albaranes,"'",""),"(",""),")","")&"'"
			rstAux.open strselect,session("dsn_cliente")
			Primero=rstAux(0)
			Ultimo=rstAux(1)
			error=rstAux(2)
			cadenaAuditoria=rstAux(3)

			if Primero>"" and Ultimo>"" and error=0 then
				nfactura_pro=d_lookup("nfactura_pro","facturas_pro","nfactura='"&Primero&"'",session("dsn_cliente"))
				
				%><SCRIPT>
					window.alert("<%=LitMsgFacturaGenerado%> <%=nfactura_pro%>");
					<%if viene<>"facturas_pro" then %>
					if (window.confirm("<%=LitMsgDeseaVer%>")==true)
						AbrirVentana('../search_layout.asp?pag1=compras/facturas_pro.asp?ndoc=<%=primero%>?mode=browse?titulo=<%=LitDetallesFac%> <%=nfactura_pro%>&pag2=compras/facturas_pro_bt.asp&pag3=compras/facturas_pro_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
					<%else%>
					window.top.opener.parent.pantalla.document.facturas_pro.action="facturas_pro.asp?nfactura=<%=primero%>&mode=browse";
     				window.top.opener.parent.pantalla.document.facturas_pro.submit()
					    top.window.close();
					<% end if %>
				</script><%
			elseif error=1 then
				%><SCRIPT>
				      window.alert("<%=LitMsgSinDireccionPrinc%>");
				</script><%
			elseif error=2 then
				%><SCRIPT>
				      window.alert("<%=LitMsgNumSeriesRepetidos%>");
				</script><%
			elseif error=3 then
				%><SCRIPT>
				      window.alert("<%=LitMsgDocExistRevContConv%>");
				</script><%
			else
				%><SCRIPT>
				      window.alert("<%=LitMsgFacturaNoGenerado%>");
				</script><%
			end if
			rstAux.close

			'**************************'
			' Ahora guardamos los pagos y los vencimientos'
			rst_cabecera.movefirst
			'**************************'

			rst_cabecera.Close
			nfactura_param = nfactura
			'vamos a auditar
			auditar_ins_bor session("usuario"),Primero,"","alta","","","conver_alb_ped_dev_fac_proMASIVO"
			'no se audita los equipos, ya que no se crea ninguno, al estar ya creados por el albaran
			'auditar_ins_bor session("usuario"),nfactura,"","alta","","","equipos_conver_alb_ped_fac_pro"
			anyadir_albaranes="OK"
		end if
	else
		rst_cabecera.Close
	end if
end function

'*****************************************************************************
'Dibuja las diferentes capas para introducir los nserie de los pedidos que lo requieran
sub DibujarSpanSeries(fdesde,fhasta,nproveedor)
     seriep=request.Form("seriep")
	    cadserie=""
	    if seriep&"">"" then
	        cadserie=" and serie='"&seriep&"' "
	    end if

	if nproveedor > "" then
		Condiciones = " nproveedor='" & nproveedor & "' and fecha>='" & fdesde & "' and fecha<='" & fhasta & "' "&cadserie&" and nalbaran is null and nfactura is null"
	else
		Condiciones = " fecha>='" & fdesde & "' and fecha<='" & fhasta & "' and nalbaran is null and nfactura is null"
	end if

	seleccion="select d.npedido,d.item,d.referencia,d.cantidadpend,d.descripcion,d.almacen,a.ctrl_nserie,a.lotecompra from detalles_ped_pro as d with(nolock),pedidos_pro as p with(nolock),articulos a with(nolock) "
	seleccion=seleccion & " where  "
	seleccion=seleccion & " d.referencia=a.referencia and  (a.ctrl_nserie=1 or a.lotecompra=1) and a.referencia like '"&session("ncliente")&"%' "
	seleccion=seleccion & " and d.npedido=p.npedido and  "& Condiciones 
	seleccion=seleccion & " order by p.fecha desc,d.npedido desc,d.item asc"
	rstAux.Open seleccion, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	npedidoAnt=""
	cont=0
	while not rstAux.eof
		nombreSpan="A" & Limpiar(rstAux("npedido")) & "B"
		nombreCampo=nombreSpan & completar(cstr(rstAux("item")),3,"0")
		if rstAux("npedido")<>npedidoAnt then
			if npedidoAnt<>"" then
				%></table>
				</span>
				<br><%
			end if
			%><SPAN ID="<%=nombreSpan%>" STYLE="display:none" height="0px">
				<table width=100% BORDER="0" CELLSPACING="1" CELLPADDING="1"><%
					'Fila de encabezado
					DrawFila color_fondo
						DrawCelda "ENCABEZADOL","","",0,LitPedido
						DrawCelda "ENCABEZADOL","","",0,LitItem
						DrawCelda "ENCABEZADOL","","",0,LitReferencia
						DrawCelda "ENCABEZADOC","","",0,LitCantidad
						DrawCelda "ENCABEZADOL","","",0,LitDescripcion
						DrawCelda "ENCABEZADOC","","",0,LitNSerie
					CloseFila
					'Escribir linea de detalle de nserie
					DrawFila color_blau
						DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rstAux("npedido"))
						DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rstAux("item")))
						DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rstAux("referencia"))
						DrawCelda2 "CELDA", "center", false, enc.EncodeForHtmlAttribute(null_s(rstAux("cantidadPend")))
						DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion")))
						pagina="../Mantenimiento/wizard_nserie.asp?mode=add&campo=" & nombreCampo & "&referencia=" & rstAux("referencia") & "&almacen=" & rstAux("almacen")
						%><input type="hidden" name="nserie_alb" value="<%=enc.EncodeForHtmlAttribute(null_s(nserie_alb))%>">
						<td CLASS=CELDACENTER><TEXTAREA CLASS=CELDA name="<%=enc.EncodeForHtmlAttribute(null_s(nombreCampo))%>" rows="2" cols="35"></TEXTAREA>
							<A CLASS=CELDAREFB href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(null_s(pagina))%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)"><IMG SRC="../images/<%=ImgVarita%>" <%=ParamImgVarita%> ALT="<%=LitAsistente%>" title="<%=LitAsistente%>"></A></td><%
					CloseFila
		else
			'Escribir linea de detalle de nserie
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rstAux("npedido"))
				DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rstAux("item")))
				DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rstAux("referencia"))
				DrawCelda2 "CELDA", "center", false, enc.EncodeForHtmlAttribute(null_s(rstAux("cantidadPend")))
				DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion")))
				cadena=cadena & "C"
				pagina="../Mantenimiento/wizard_nserie.asp?mode=add&campo=" & nombreCampo & "&referencia=" & rstAux("referencia") & "&almacen=" & rstAux("almacen")
				%><td CLASS=CELDACENTER><TEXTAREA CLASS=CELDA name="<%=enc.EncodeForHtmlAttribute(null_s(nombreCampo))%>" rows="2" cols="35"></TEXTAREA>
					<A CLASS=CELDAREFB href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(null_s(pagina))%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)"><IMG SRC="../images/<%=ImgVarita%>" <%=ParamImgVarita%> ALT="<%=LitAsistente%>" title="<%=LitAsistente%>"></A></td><%
			CloseFila
		end if
		npedidoAnt=rstAux("npedido")
	 	rstAux.movenext
		cont=cont+1
		if rstAux.eof then
			%></table>
			</span>
			<br><%
		end if
	wend
	rstAux.close
end sub

function validar_pro_fac(nproveedor,nfactura,ffactura)
	ModDocumento=true
	''ricardo 12/11/2003 comprobamos que no exista el nalbaran_pro para un mismo proveedor
	no_continuar=0
	strselect="select count(nfactura) as contador from facturas_pro where nproveedor='" & nproveedor & "' and nfactura_pro='" & nfactura & "'  and year(fecha)= year(convert (datetime,'" & ffactura & "' ))"
	rst_cabecera.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if not rst_cabecera.eof then
		cuantas_facturas=null_z(rst_cabecera("contador"))
		rst_cabecera.close
		if cuantas_facturas>0 then
			ModDocumento=false
		end if
	else
		rst_cabecera.close
	end if
	if ModDocumento=true then validar_pro_fac="OK" else validar_pro_fac="NO_OK"
end function

'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************

borde=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion

	Dim ListaN,ListaNumeros,k%>
<form name="albpedpro_facpro_param" method="post">
<% PintarCabecera "albpedpro_facpro_param.asp"
'Leer parámetros de la página
    si_tiene_modulo_EBESA=ModuloContratado(session("ncliente"),ModEBESA)
    si_tiene_modulo_ccostes=ModuloContratado(session("ncliente"),ModCcostes_Gestion) 
    si_tiene_modulo_fabricacion=ModuloContratado(session("ncliente"),ModProduccion)
	mode		= Request.QueryString("mode")

	si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)

	nproveedor	= limpiaCadena(Request.QueryString("nproveedor"))
	if nproveedor ="" then
		nproveedor	= limpiaCadena(Request.form("nproveedor"))
	end if
	if nproveedor > "" then
		nproveedor=session("ncliente") & completar(nproveedor,5,"0")
	end if
	checkCadena(nproveedor)

	'AKI
	if nproveedor>"" then
	   if d_lookup("nproveedor", "proveedores", "nproveedor='" & nproveedor & "'",session("dsn_cliente"))>"" then
	      nombre = d_lookup("razon_social", "proveedores", "nproveedor='" & nproveedor & "'",session("dsn_cliente"))
	   else
	      nproveedor=""%>
		  <script>
		      window.alert("<%=LitNoExisteProveedor%>");
		      parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select1";
              document.location.href = "albpedpro_facpro_param.asp?mode=select1";
		      
		  </script>
	   <%end if
	end if
	viene= Request.QueryString("viene")
	if viene ="" then
		viene	= Request.form("viene")
	end if	
	ndoc= Request.QueryString("ndoc")
	if ndoc ="" then
		ndoc	= Request.form("ndoc")
	end if%>
	<input type="hidden" name="nproveedor2" value="<%=enc.EncodeForHtmlAttribute(null_s(nproveedor))%>">
	<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>">
	<input type="hidden" name="ndoc" value="<%=enc.EncodeForHtmlAttribute(ndoc)%>">

	<%fdesde		= limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde	= limpiaCadena(Request.form("fdesde"))
	end if
	fhasta		= limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta	= limpiaCadena(Request.form("fhasta"))
	end if

	if request.querystring("tabla")>"" then tabla = request.querystring("tabla")

	if tabla="" then
	   if request.form("tabla")="pedidos_pro" then
		   alb="false"
		   tabla = "pedidos_pro"
	   else
		   alb="true"
		   tabla = "albaranes_pro"
	   end if
	end if

	nserie_fac	= Request.QueryString("nserie_fac")
	if nserie_fac ="" then
		nserie_fac	= Request.form("nserie_fac")
	end if
	fecha_fac	= Request.QueryString("fecha_fac")
	if fecha_fac ="" then
		fecha_fac	= Request.form("fecha_fac")
	end if
	nfactura_pro_fac= Request.QueryString("nfactura_pro_fac")
	if nfactura_pro_fac ="" then
		nfactura_pro_fac	= Request.form("nfactura_pro_fac")
	end if
	inclDevol=nz_b2(Request.QueryString("inclDevol"))
	if inclDevol ="" or inclDevol =0 then
		inclDevol= nz_b2(Request.form("inclDevol"))
	end if

    nserie		= Request.form("nserie")&""
	centro      = Request.Form("centroa")&""
	ffactura	= Request.form("ffactura")&""
	h_nfilas	= Request.form("h_nfilas")&""
	nfactura_pro = request.form("nfactura_pro")&""
	nfactura_pro2=request.form("nfactura_pro")
	nproveedor2=limpiaCadena(request.form("nproveedor2"))
	checkCadena(nproveedor2)

	strwhere=""
	strwhere2=""

	WaitBoxOculto LitEsperePorFavor
	Alarma "albpedpro_facpro_param.asp"%>
<%	set rst_factura_det = Server.CreateObject("ADODB.Recordset")
	set rst_factura = Server.CreateObject("ADODB.Recordset")
	set rst_det = Server.CreateObject("ADODB.Recordset")
	set rst_cabecera = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rst_albaran_con = Server.CreateObject("ADODB.Recordset")
	set rst_pedido_con = Server.CreateObject("ADODB.Recordset")
	set rst_factura_con = Server.CreateObject("ADODB.Recordset")
	set rst_factura_pagos = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rstDomi = Server.CreateObject("ADODB.Recordset")
	set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeout = 300
	conn.CommandTimeout = 300	
	nfactura_param = ""

    


	'REALIZAR LA CONVERSION
	if mode="confirm"then
		mode="select1"
		if viene="facturas_pro" then
		    if validar_albaranes then
		        if validarNSerie("albaranes_pro") then
		            anyadir_albaranes(ndoc)
		        end if
		    end if
		else
		    if validar_pro_fac(nproveedor2,nfactura_pro2,ffactura)="NO_OK" then
		        if alb = "true" then
		            strwhere=request.Form("strwhere")%>
			        <input type="hidden" name="strwhere" value="<%=strwhere%>">
			        <SCRIPT>
			            window.alert("<%=LitMsgNumeroFacturaRepetido%>");
			            parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select2";
                        document.albpedpro_facpro_param.action = "albpedpro_facpro_param.asp?mode=select2&nproveedor=<%=trimcodempresa(nproveedor2)%>&submod=err";
			            document.albpedpro_facpro_param.submit();
			            
			        </script><%
			    else
			        %><script>
			              window.alert("<%=LitMsgNumeroFacturaRepetido%>");
			              history.back();
			              history.back();
			              parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select2";
				    </script><%
			    end if
		    else
			    if alb = "true" then
			        nalbtot=h_nfilas
			        nalbaranes=0
				    if validar_albaranes then
					    if validarNSerie("albaranes_pro") then
						    if convertir_albaranes="OK" and false then
							    %><script language="javascript">
							          if (window.confirm("<%=LitMsgDeseaVer%>") == true)
									        AbrirVentana('../search_layout.asp?pag1=compras/facturas_pro.asp?ndoc=<%=enc.EncodeForJavascript(null_s(nfactura_param))%>?mode=browse?titulo=<%=LitDetallesFac%> <%=nfactura_pro%>&pag2=compras/facturas_pro_bt.asp&pag3=compras/facturas_pro_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
							    </script><%
						    end if
						    if cint(nalbaranes)<cint(nalbtot) then
						        mode="select2"
						        tabla=request.Form("tabla")&""
						        strwhere=request.Form("strwhere")&""
						        %><script language="javascript">
						              parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select2";
						        </script><%
						    end if
					    else
						    %><SCRIPT>
						          history.back();
						          history.back();
						          parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select2";
						    </script><%
					    end if
				    end if
			    else
                    npedtot=h_nfilas
			        npedidos=0
				    if validar_pedidos then
					    if validarNSerie("pedidos_pro") then
						    if convertir_pedidos="OK"  and false then
							    %><script language="javascript">
							          if (window.confirm("<%=LitMsgDeseaVer%>") == true)
									    AbrirVentana('../search_layout.asp?pag1=compras/facturas_pro.asp?ndoc=<%=enc.EncodeForJavascript(null_s(nfactura_param))%>?mode=browse?titulo=<%=LitDetallesFac%> <%=nfactura_pro%>&pag2=compras/facturas_pro_bt.asp&pag3=compras/facturas_pro_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
							    </script><%
						    end if
                            if cint(npedidos)<cint(npedtot) then
						        mode="select2"
						        tabla=request.Form("tabla")&""
						        strwhere=request.Form("strwhere")&""
						        %><script language="javascript">
						              parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select2";
						        </script><%
						    end if
					    else
                            mode="select2"
                            tabla=request.Form("tabla")&""
                            strwhere=request.Form("strwhere")&""
						    %><SCRIPT>
						          //history.back();
						          //history.back();
						          parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select2";
						    </script><%
					    end if
				    end if
			    end if
		    end if
		end if
	end if

	'PARAMETROS DE ENTRADA
	if mode="select1" then%>
		<input type="hidden" name="nserie_fac" value="<%=enc.EncodeForHtmlAttribute(null_s(nserie_fac))%>">
		<input type="hidden" name="fecha_fac" value="<%=enc.EncodeForHtmlAttribute(null_s(fecha_fac))%>">
		<input type="hidden" name="nfactura_pro_fac" value="<%=enc.EncodeForHtmlAttribute(null_s(nfactura_pro_fac))%>">
		<input type="hidden" name="htabla" value="<%=enc.EncodeForHtmlAttribute(tabla)%>">
		<table BORDER="0" CELLSPACING="1" CELLPADDING="1">
        <hr style="display:block"/>
        <%if viene="facturas_pro" then%>
                <script  language=javascript>
                    document.albpedpro_facpro_param.htabla.value = 'albaranes_pro';
                </script>
        <%else
            DrawDiv "1","",""
            DrawLabel "","",LitPedidos%><input type="radio" name="tabla" value="pedidos_pro" onClick="document.albpedpro_facpro_param.htabla.value = 'pedidos_pro';tier2Menu(pedidos)" <%=iif(tabla="pedidos_pro","checked","")%>><%CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LitAlbaranes%><input type="radio" name="tabla" value="albaranes_pro" onClick="document.albpedpro_facpro_param.htabla.value = 'albaranes_pro';tier2Menu(albaranes)" <%=iif(tabla="albaranes_pro" or tabla="", "checked", "")%>><%CloseDiv
		end if%>
		</table><hr style="display:block"><%
			'DrawCelda2 "CELDA width=213px", "left", false, LitDesdeFecha + ":"
		 	'DrawInputCelda "CELDA","","",10,0,"","fdesde",iif(fdesde="","01/01/"+right(cstr(year(date)),2),fdesde)
            EligeCelda "input","add","left","","",0,LitDesdeFecha,"fdesde",10,iif(fdesde="","01/01/"+right(cstr(year(date)),2),fdesde)
            DrawCalendar "fdesde"
			'DrawCelda2 "CELDA", "left", false, LitHastaFecha + ":"
		 	'DrawInputCelda "CELDA","","",10,0,"","fhasta",iif(fhasta>"",fhasta,date)
            EligeCelda "input","add","left","","",0,LitHastaFecha,"fhasta",10,iif(fhasta>"",fhasta,date)
            DrawCalendar "fhasta"
		%>
		<%if viene="facturas_pro" then%>
		    <input type="hidden" name="nproveedor" value="<%=trimCodEmpresa(nproveedor)%>" />
		<%else
			'DrawCelda2 "CELDA width=213px", "left", false, LitProveedor + ":"
			'DrawInputCeldaBuscar "CELDA","","",5,0,"","nproveedor",nproveedor,"AbrirVentana('proveedores_busqueda.asp?ndoc=albpedpro_facpro_param&titulo=SELECCIONAR PROVEEDOR&mode=search','P',"+cstr(altoventana)+","+cstr(anchoventana)+")",""
            DrawDiv "1","",""
            DrawLabel "","",LitProveedor
            DrawInput "CELDA maxlenght='5'","","nproveedor",trimCodEmpresa(nproveedor), "onchange = 'TraerProveedor()'"%><a CLASS=CELDAREFB href="javascript:AbrirVentana('proveedores_busqueda.asp?ndoc=albpedpro_facpro_param&titulo=<%=LitSelProveedor%>&mode=search','P',<%=cstr(altoventana)%>,<%=cstr(anchoventana)%>)" OnMouseOver="self.status='<%=LitVerProveedor%>'; return true;" OnMouseOut="self.status=''; return true;"><IMG SRC="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> ALT=""></A><%DrawInput "CELDA disabled","","razon_social",nombre,""
			CloseDiv%>
		<%end if
		'**RGU 20/6/07 %>
	    <SPAN ID="albaranes" STYLE="display:<%=iif(tabla="albaranes_pro","","none")%> "><%
			if si_tiene_modulo_ccostes<>0 then
	                h_centroa=request.Form("centroa")&""
	                strselect="select codigo, descripcion from tiendas with(nolock) where codigo like '"&session("ncliente")&"%'  order by descripcion "
	                rstAux.CursorLocation=3
    	            rstAux.Open  strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                    DrawSelectCelda "CELDA7","","",0,LitCentroCostes,"centroa",rstAux,h_centroa,"codigo","descripcion","",""
    	            rstAux.Close
	        end if
                    h_seriea=request.Form("seriea")&""
	                strselect="select nserie,right(nserie,len(nserie)-5)+'-'+nombre as descripcion from series with(nolock) where nserie like '"&session("ncliente")&"%' and tipo_documento='ALBARAN DE PROVEEDOR' order by nserie "
                    rstAux.CursorLocation=3
    	            rstAux.Open  strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                    DrawSelectCelda "CELDA","175","",0,LitSerie,"seriea",rstAux,h_seriea,"nserie","descripcion","",""
    	            rstAux.Close


                '------------------------------------------------------------
	            if viene<>"facturas_pro" then
					    rstAux.open "select nserie,right(nserie,len(nserie)-5)+'-'+nombre as descripcion from series where tipo_documento ='DEVOLUCION A PROVEEDOR' and nserie like '" & session("ncliente") & "%' order by nserie",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                        DrawSelectCelda "CELDA onchange='if (this.value==""""){document.albpedpro_facpro_param.inclDevol.checked=false}else{document.albpedpro_facpro_param.inclDevol.checked=true}'","200","",0,LitSerieDev,"serie_dev",rstAux,doc_serie_dev,"nserie","descripcion","",""
					    rstAux.close
                        DrawDiv "1","",""
                        DrawLabel "","",LitInclDevol%><input type="checkbox" name="inclDevol" value="1" onClick="if (!this.checked) {document.pedcli_faccli_param.serie_dev.value=''}"><%CloseDiv
				 end if%>
	    </SPAN>
		<SPAN ID="pedidos" STYLE="display:<%=iif(tabla="pedidos_pro","","none")%>">
			<%
	        h_seriep=request.Form("seriep")&""
	        strselect="select nserie,right(nserie,len(nserie)-5)+'-'+nombre as descripcion from series with(nolock) where nserie like '"&session("ncliente")&"%' and tipo_documento='PEDIDO A PROVEEDOR' "
            rstAux.CursorLocation=3
    	    rstAux.Open  strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
    	    DrawSelectCelda "CELDA","","",0,LitSerie,"seriep",rstAux,h_seriep,"nserie","descripcion","",""
    	    rstAux.Close

	        %>
	    </SPAN>
	    <%'**RGU ***

	'LISTA DE DOCUMENTOS SELECCIONADOS
	elseif mode="select2" then%>
        <input type="hidden" name="nproveedor" value="<%=trimcodempresa(nproveedor2)%>">
        <input type="hidden" name="inclDevol" value="<%=enc.EncodeForHtmlAttribute(inclDevol)%>">
        <input type="hidden" name="nserie_fac" value="<%=enc.EncodeForHtmlAttribute(null_s(nserie_fac))%>">
        <input type="hidden" name="fecha_fac" value="<%=enc.EncodeForHtmlAttribute(null_s(fecha_fac))%>">
        <input type="hidden" name="nfactura_pro_fac" value="<%=enc.EncodeForHtmlAttribute(null_s(nfactura_pro_fac))%>">

        <%eliminar="if exists (select * from sysobjects where id = object_id('egesticet.[" & session("usuario") & "_temporal_dev]') and sysstat " & _
			         " & 0xf = 3) drop table egesticet.[" & session("usuario") & "_temporal_dev]"
        rst.CursorLocation=2
        rst.open eliminar,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        eliminar="if exists (select * from sysobjects where id = object_id('egesticet.[" & session("usuario") &  "_dev]') and sysstat " & _
			         " & 0xf = 3) drop table egesticet.[" & session("usuario")&"_dev]"
        rst.CursorLocation=2
        ''response.write("el eliminar es-" & eliminar & "-<br>")
        rst.open eliminar,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

        crear="CREATE TABLE [" & session("usuario") &"_dev] ("
        crear=crear & "ndevolucion varchar(20),item smallint,cantidad real,cantidad2 real,referencia varchar(30),mi_nserie varchar(3000),nserie varchar(2000), pvp money, importe money"
        crear=crear & ")"
        
        crear2="CREATE TABLE [" & session("usuario") & "_temporal_dev] ("
        crear2=crear2 & "ndevolucion varchar(20),item smallint,referencia varchar(30),mi_nserie varchar(3000),nserie varchar(2000), pvp money, importe money"
        crear2=crear2 & ")"
        rst.CursorLocation=2
        rst.open crear,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        rst.open crear2,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
        GrantUser session("usuario")&"_dev", session("dsn_cliente")
        GrantUser session("usuario")&"_temporal_dev" , session("dsn_cliente")

	    DropTable session("usuario"), session("dsn_cliente")
    	
	    crear="CREATE TABLE [" & session("usuario") & "] ("
	    crear=crear & "ndocumento varchar(20),item smallint,cantidad real,referencia varchar(30),mi_nserie varchar(8000),nserie varchar(8000),lote varchar(100),almacen varchar(10)"
	    crear=crear & ")"
        rst.CursorLocation=2
	    rst.open crear,session("dsn_cliente"),adOpenKeyset,adLockOptimistic

	    GrantUser session("usuario"), session("dsn_cliente")
        
        IsCadenaBusqueda = false
        if request.QueryString("mode")="confirm" and mode="select2" then
        else
            if request.QueryString("submod")&""="err" then
                 strwhere=request.Form("strwhere")&""
            else
                IsCadenaBusqueda = true
		        strwhere = CadenaBusqueda(fdesde,fhasta,nproveedor)
                strwhereSinCentro = strwhere
	            if centro>"" then
	                strwhere = strwhere & " and tienda='"&centro&"'"
	            end if
		    end if
		end if
 
		if inclDevol=1 then
            'ALBARAN o GUÍA DE REMISIÓN CON CHECK DE INCLUIR DEVOLUCIONES
            strselect="select 'ALB' as tipo,nalbaran_pro,nalbaran,serie,fecha,nproveedor,total_albaran,divisa from "+tabla+" with(nolock) where nalbaran like '"&session("ncliente")&"%' and " + strwhere +" and nfactura is null "
		    strselect=strselect+" UNION select 'DEV' as tipo,right(d.ndevolucion,len(d.ndevolucion)-5) as nalbaran_pro,d.ndevolucion,serie,fecha,nproveedor,0,null "
		    strselect=strselect+" from devoluciones_pro d with(nolock), (select ndevolucion from detalles_dev_pro where ndevolucion like '"&session("ncliente")&"%' group by ndevolucion having min(cantidadpend)<>0 or max(cantidadpend)<>0) dd "
		                
            if IsCadenaBusqueda=true then
                strwhere = strwhereSinCentro
            end if
            strselect=strselect+" where d.ndevolucion like '"&session("ncliente")&"%' and d.ndevolucion=dd.ndevolucion  and "+ strwhere
		    strselect=strselect+" order by  fecha desc,nalbaran_pro desc,nproveedor"
		else
		    if alb="true" then
                'ALBARAN o GUÍA DE REMISIÓN
                strselect="select 'ALB' as tipo, nalbaran_pro,nalbaran,serie,fecha,nproveedor,total_albaran,divisa from "+tabla+" where nalbaran like '"&session("ncliente")&"%' and " + strwhere +" and nfactura is null order by fecha desc,nalbaran_pro desc,nproveedor  "
		    else
                'PEDIDOS
                strselect="select npedido,serie,fecha,nproveedor,total_pedido,divisa from "+tabla+" where npedido like '"&session("ncliente")&"%' and " + strwhere +"  order by fecha desc,npedido desc,nproveedor "
                strselect="select p.npedido,p.serie,p.fecha,p.nproveedor,p.total_pedido,p.divisa "
                strselect=strselect & " ,CASE when exists( "
			    strselect=strselect & "             select top 1 d.REFERENCIA "
			    strselect=strselect & "             from DETALLES_PED_PRO as d with(NOLOCK),articulos as a with(NOLOCK) "
			    strselect=strselect & "             where d.NPEDIDO like '"&session("ncliente")&"%' and a.REFERENCIA like '"&session("ncliente")&"%' "
			    strselect=strselect & "             and d.NPEDIDO=p.NPEDIDO and a.REFERENCIA=d.REFERENCIA "
			    strselect=strselect & "             and a.CTRL_NSERIE=1 "
			    strselect=strselect & "             ) "
                strselect=strselect & " then 1 "
                strselect=strselect & " else 0 "
                strselect=strselect & " end as requiere_nserie "
                strselect=strselect & " ,CASE when exists( "
			    strselect=strselect & "             select top 1 d.REFERENCIA "
			    strselect=strselect & "             from DETALLES_PED_PRO as d with(NOLOCK),articulos as a with(NOLOCK) "
			    strselect=strselect & "             where d.NPEDIDO like '"&session("ncliente")&"%' and a.REFERENCIA like '"&session("ncliente")&"%' "
			    strselect=strselect & "             and d.NPEDIDO=p.NPEDIDO and a.REFERENCIA=d.REFERENCIA "
			    strselect=strselect & "             and a.LOTECOMPRA=1 "
			    strselect=strselect & "             ) "
                strselect=strselect & " then 1 "
                strselect=strselect & " else 0 "
                strselect=strselect & " end as requiere_nlote "
                strselect=strselect & " from "+tabla+" as p with(NOLOCK) "
                strselect=strselect & " where p.npedido like '"&session("ncliente")&"%' "
                strselect=strselect & " and " & strwhere
                strselect=strselect & " order by p.fecha desc,p.npedido desc,p.nproveedor"

		    end if
		end if
        %>
		<input type="hidden" name="strwhere" value="<%=enc.EncodeForHtmlAttribute(null_s(strwhere))%>">
		<%
        ''response.write(strselect)
        ''response.end
        rst.CursorLocation=3
        rst.Open strselect, session("dsn_cliente")'',adUseClient, adLockReadOnly
		
		if rst.eof then
			rst.close
		   	%><script type="text/javascript" language="javascript">
		   	      window.alert("<%=LitNoExisteRegProveedor%>");
		   	      parent.botones.document.location = "albpedpro_facpro_param_bt.asp?mode=select1";
                  document.albpedpro_facpro_param.action = "albpedpro_facpro_param.asp?mode=select1&viene=<%=enc.EncodeForJavascript(viene)%>&nproveedor=<%=enc.EncodeForJavascript(trimcodempresa(nproveedor))%>&";
		   	      document.albpedpro_facpro_param.submit();
		   	      
			</script><%
		else
			%><input type="hidden" name="tabla" value="<%=enc.EncodeForHtmlAttribute(tabla)%>"><%
			if alb="true" then
				%><input type="hidden" name="mensaje" value="<%=LitMsgConvAlbaranesConfirm%>"><%
			else
				%><input type="hidden" name="mensaje" value="<%=LitMsgConvPedidosConfirm%>"><%
			end if
      		%><table width="100%" border="<%=borde%>" cellspacing="1" cellpadding="1"><%
      		 	if viene="facturas_pro" then%>
				 	<input type="hidden" name="nfactura" value="<%=enc.EncodeForHtmlAttribute(null_s(ndoc))%>" />
				<%else
		    		DrawFila color_blau
					    DrawCelda2 "CELDA", "left", false, LitSerieFactura + ":"
					    rstAux.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series where tipo_documento ='FACTURA DE PROVEEDOR' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					    'DrawSelectCelda "CELDA","","",0,"","nserie",rstAux,nserie_fac,"nserie","descripcion","",""

					        DrawCelda2 "CELDA", "left", false, LitFechaFactura + ":"
					        if fecha_fac>"" then
					 	        DrawInputCelda "CELDA maxlength='10'","","",10,0,"","ffactura",fecha_fac
					        else
					 	        DrawInputCelda "CELDA maxlength='10'","","",10,0,"","ffactura",date
					        end if
					        DrawCelda2 "CELDA", "left", false, LitNFactura + ":"
				 	        DrawInputCelda "CELDA maxlength='20'","","",20,0,"","nfactura_pro",nfactura_pro_fac
    				    CloseFila
    				end if%>
			</table>

			<%if si_tiene_modulo_mantenimiento<>0 then
				if alb="false" then
					DibujarSpanSeries fdesde,fhasta,nproveedor
				end if
			end if%>

			<table width="100%" border="0" cellspacing="1" cellpadding="1">
			    <%'Fila de encabezado
				DrawFila color_fondo
					%><td class="CELDA">
						<input type="checkbox" name="check" value="true" onClick="seleccionar();">
					</td><%
					if alb="true" then
					    if viene="facturas_pro" then
					        DrawCelda "ENCABEZADOL","","",0,LitAlbaran
					    else
					        DrawCelda "ENCABEZADOL","","",0,LitAlbaranDev
					    end if
						DrawCelda "ENCABEZADOL","","",0,LitNserie
						if viene<>"facturas_pro" then
						    DrawCelda "ENCABEZADOL","","",0,LitCompletar
						end if
					else
						DrawCelda "ENCABEZADOL","","",0,LitPedido
						if si_tiene_modulo_mantenimiento<>0 or si_tiene_modulo_produccion<>0 then
							DrawCelda "ENCABEZADOL","","",0,LitCompletar
						end if
					end if
					DrawCelda "ENCABEZADOR","","",0,LitFecha
					DrawCelda "ENCABEZADOL","","",0,LitProveedor
					DrawCelda "ENCABEZADOR","","",0,LitTotal
				CloseFila

				VinculosPagina(MostrarPedidosPro)=1:VinculosPagina(MostrarAlbaranesPro)=1:VinculosPagina(MostrarDevolucionesPro)=1
				CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

				MB=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))

				fila=1
				while not rst.EOF
					n_decimales=d_lookup("ndecimales","divisas","codigo='" & iif(isnull(rst("divisa")), MB, rst("divisa")) & "'",session("dsn_cliente"))
					'Seleccionar el color de la fila.
					if ((fila+1) mod 2)=0 then
						color=color_blau
					else
						color=color_terra
					end if

					DrawFila color
						if alb="true" then
							%><td class="CELDA">
								<input type="checkbox" name="check<%=fila%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("nalbaran")))%>">
								<input type="hidden" name="tipo<%=fila%>" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("tipo")))%>">
								
							</td>
							<td class="CELDALEFT" align="left"><%
							    if rst("tipo")="DEV" then
								    response.write(Hiperv(OBJDevolucionesPro,rst("nalbaran"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("nalbaran_pro"),LitVerDevolucion))
								else
								    response.write(Hiperv(OBJAlbaranesPro,rst("nalbaran"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("nalbaran_pro"),LitVerAlbaran))
								end if%>
							</td><% 
							DrawCelda "CELDALEFT","","",0,trimCodEmpresa(rst("serie"))
   							if viene<>"facturas_pro" then 
                                if si_tiene_modulo_EBESA<>0 then%>
						            <td class="CELDACENTER" style='width:60px'>
						                <%if rst("tipo")="DEV" then %>
									        <div class=CELDAB2 onMouseOver="this.className='CELDAB2'" onMouseOut="this.className='CELDAB2'" onclick="abrir_detallesEBESA('<%=enc.EncodeForJavascript(null_s(rst("nalbaran")))%>','<%=fila %>');">
		              						        <img src="../Images/<%=ImgCarpetaCerrada%>" height="17" width="17" border="0" id="imgDev<%=fila%>" alt="<%=LitCompletar%>" title="<%=LitCompletar%>">
			      		       			        </div>							
								        <%end if%>
    						        </td><%
    						    else%>
						            <td class="CELDACENTER" style='width:60px'>
						                <%if rst("tipo")="DEV" then %>
									        <div class="CELDAB2" onMouseOver="this.className='CELDAB2'" onMouseOut="this.className='CELDAB2'" onclick="abrir_detalles('<%=enc.EncodeForJavascript(null_s(rst("nalbaran")))%>','<%=fila %>');">
	              						        <img src="../Images/<%=ImgCarpetaCerrada%>" height="17" width="17" border="0" id="img1" alt="<%=LitCompletar%>" title="<%=LitCompletar%>">
		      		       			        </div>							
								        <%end if %>
    						        </td><%
    						    end if
   						    end if  
							DrawCelda "CELDARIGHT","","",0,rst("fecha")
							DrawCelda "CELDA","","",0,d_lookup("razon_social","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))
							abreviatura = d_lookup("abreviatura", "divisas", "codigo='" + rst("divisa") + "'", session("dsn_cliente"))
							if rst("tipo")="DEV" then
							    DrawCelda "CELDARIGHT","","",0,""
							else
							    DrawCelda "CELDARIGHT","","",0,formatnumber(rst("total_albaran"),n_decimales,-1,0,-1) & "  " & abreviatura
							end if 
						else
							%><td class="CELDA">
                            	<%
                                hace_falta_serie=rst("requiere_nserie")
                                hace_falta_lote=rst("requiere_nlote")
                                %>
							    <input type="hidden" name="hace_falta_serie<%=fila%>" value="<%=enc.EncodeForHtmlAttribute(hace_falta_serie)%>"/>
						        <%
                                %>
							    <input type="hidden" name="hace_falta_lote<%=fila%>" value="<%=enc.EncodeForHtmlAttribute(hace_falta_lote)%>"/>
						        <%
                                bloquear_ck=""
                                if hace_falta_serie=1 or hace_falta_lote=1 then
                                    bloquear_ck="disabled='disabled'"
                                end if
                                %>
								<input type="checkbox" name="check<%=fila%>" <%=bloquear_ck%> value="<%=enc.EncodeForHtmlAttribute(null_s(rst("npedido")))%>">
								<input type="hidden" name="nserie<%=fila%>" value="">
							</td>
							<td class="CELDALEFT" align="left">
								<%=Hiperv(OBJPedidosPro,rst("npedido"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("npedido")),LitVerPedido)%>
							</td><%
							if si_tiene_modulo_mantenimiento<>0 or si_tiene_modulo_fabricacion then
								seleccion="select a.referencia from detalles_ped_pro dpp with(nolock) inner join articulos a with(nolock) on (a.ctrl_nserie=1 or a.lotecompra=1) and a.referencia =dpp.referencia and a.referencia like '"+session("ncliente")+"%' where dpp.npedido='" & rst("npedido") & "'"
								rstAux.open seleccion, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
								if rstAux.eof then
									DrawCelda "CELDARIGHT","","",0,""
								else
									cadena="A" & Limpiar(rst("npedido")) & "B"
									%><td>
									    <div class="CELDAB2" onMouseOver="this.className='CELDAB2'" onMouseOut="this.className='CELDAB2'" onclick="javascript:abrir_detalles('<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>',<%=fila%>,'pedpro_facpro_param');tier3Menu(document.getElementById('imgDev<%=fila%>'));">
		              						<IMG SRC="../Images/<%=ImgCarpetaCerrada%>" height="17" width="17" border="0" id="imgDev<%=fila%>" alt="<%=LitCompletarTit%>" alt="<%=LitCompletar%>">
			      		       			</div>
									</td><%
								end if
								rstAux.close
							end if
							DrawCelda "CELDARIGHT","","",0,rst("fecha")
							DrawCelda "CELDA","","",0,d_lookup("razon_social","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))

							DrawCelda "CELDARIGHT","","",0,formatnumber(rst("total_pedido"),n_decimales,-1,0,-1)
						end if
					CloseFila
					fila=fila+1
					rst.MoveNext
				wend%>
				<input type="hidden" name="h_nfilas" value="<%=rst.recordcount%>">
				<%rst.Close%>
			</table>
		<%end if
	end if%>
</form>
<hr>
<%end if
connRound.close
set connRound = Nothing
set connRound = Nothing
set rst_factura_det = Nothing
set rst_factura = Nothing
set rst_det = Nothing
set rst_cabecera = Nothing
set rstAux = Nothing
set rst = Nothing
set rst_albaran_con = Nothing
set rst_pedido_con = Nothing
set rst_factura_con = Nothing
set rst_factura_pagos = Nothing
set rstAux2 = Nothing
set rstDomi = Nothing
%>
</BODY>
</HTML>
