<%@ Language=VBScript %>
<%
' JCI 30/04/2003 : Solución de problemas y errores varios en la conversión de pedidos con números de serie
response.buffer=true
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>

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

<!--#include file="pedpro_albpro_param.inc" -->
<!--#include file="pedbis_pro.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->

<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">

sAgent = navigator.userAgent;

function tier1Menu(objImage) {

	if (objImage.src== "../Images/<%=ImgCarpetaAbierta%>"){
		objImage.src = "../images/<%=ImgCarpetaCerrada%>";
	}
	else {
		objImage.src = "../Images/<%=ImgCarpetaAbierta%>";
        }
}

function abrir_detalles(npedido,fila){

	//pagina="../central.asp?pag1=compras/pedpro_albpro_param_completar.asp&ndoc=" + npedido + "&tdocumento=" + fila + "&pag2=compras/pedpro_albpro_param_completar_bt.asp";
	
	pagina="../central.asp?pag1=mantenimiento/conv_ped_alb_completar.asp&ndoc=" + npedido + "&tdocumento=" + fila + "&viene=compras&pag2=mantenimiento/conv_ped_alb_completar_bt.asp";
	AbrirVentana(pagina,'P',<%=altoventana%>,<%=anchoventana%>);
}

function seleccionar() {
	nregistros=document.pedpro_albpro_param.h_nfilas.value;
	if (document.pedpro_albpro_param.check.checked)
	{
	    for (i=1;i<=nregistros;i++)
	    {
	        nombre = "hace_falta_lote" + i;
	        nombre2 = "hace_falta_serie" + i;
	        nombre3="check" + i;
	        //if (document.pedpro_albpro_param.elements[nombre].value==0 && document.pedpro_albpro_param.elements[nombre2].value==0)
	        if (document.pedpro_albpro_param.elements[nombre3].disabled==false)
	        {
	            nombre="check" + i;
	            document.pedpro_albpro_param.elements[nombre].checked=true;
	        }
		}
		document.pedpro_albpro_param.check.value="yyy"
	}
	else
	{
	    for (i=1;i<=nregistros;i++)
	    {
	        nombre = "hace_falta_lote" + i;
	        nombre2 = "hace_falta_serie" + i;
	        nombre3="check" + i;
	        //if (document.pedpro_albpro_param.elements[nombre].value==0 && document.pedpro_albpro_param.elements[nombre2].value==0)
	        if (document.pedpro_albpro_param.elements[nombre3].disabled==false)
	        {
	            nombre="check" + i;
	            document.pedpro_albpro_param.elements[nombre].checked=false;
	        }
		}
		document.pedpro_albpro_param.check.value="xxx"
	}
}

function TraerProveedor(mode){

	document.location.href="pedpro_albpro_param.asp?nproveedor=" + document.pedpro_albpro_param.nproveedor.value
	+ "&mode=" + mode + "&fdesde=" + document.pedpro_albpro_param.fdesde.value
	+ "&fhasta=" + document.pedpro_albpro_param.fhasta.value
	+ "&viene=pedpro_albpro_param.asp";
}

</script>
<body onload="self.status='';" class="BODY_ASP">
<%sub actualizar_totales(nalbaran)
    rst_albaran.cursorlocation=2
	rst_albaran.Open "select * from albaranes_pro where nalbaran='"+nalbaran+"'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst_albaran("divisa") & "'",session("dsn_cliente"))
	if n_decimales = "" then
		n_decimales = 0
	end if

	'Miramos si el proveedor tiene recargo de equivalencia
    rstAux.cursorlocation=3
	rstAux.open "Select re from proveedores with(NOLOCK) where nproveedor='" + rst_albaran("nproveedor") + "'",session("dsn_cliente")
	if not rstAux.eof then
		TieneRE=rstAux("re")
	else
		TieneRE=0
	end if
	rstAux.close

	rst_albaran("importe_bruto")	= 0
	rst_albaran("base_imponible")	= 0
	rst_albaran("total_descuento")	= 0
	rst_albaran("total_iva")		= 0
	rst_albaran("total_re")			= 0
	rst_albaran("total_recargo")	= 0
	rst_albaran("total_irpf")	= 0
	SumImporteBruto=0
	SumTotalDto=0
	SumBaseImponible=0
	SumTotalIva=0
	SumTotalIvaTotal=0
	SumTotalRE=0
	SumTotalRETotal=0
	SumTotalRF=0
	SumTotalIRPF=0
	SumTotalImporte=0

	seleccion="select sum(importe) as suma, iva, re from detalles_alb_pro with(NOLOCK) "
	seleccion=seleccion+"where nalbaran ='"+rst_albaran("nalbaran")+"' and mainitem is null "
	seleccion=seleccion+"GROUP BY IVA, RE "
	seleccion=seleccion+" union all "
	seleccion=seleccion+"select sum(importe) as suma, iva, re from conceptos_alb_pro with(NOLOCK) "
	seleccion=seleccion+"where nalbaran ='"+rst_albaran("nalbaran")+"' "
	seleccion=seleccion+"GROUP BY IVA, RE ORDER BY IVA"
    rstAux.cursorlocation=3
	rstAux.open seleccion,session("dsn_cliente")

	if not rstAux.EOF then
		ivaAnt=rstAux("iva")
	end if
	while not rstAux.EOF
		if rstAux("iva")<>ivaAnt then
			SumTotalIvaTotal=SumTotalIvaTotal+miround(null_z(SumTotalIva),n_decimales)
			SumTotalRETotal=SumTotalRETotal+miround(null_z(SumTotalRE),n_decimales)
			SumTotalIva=0
			SumTotalRE=0
			ivaAnt=rstAux("iva")
		end if
		SumImporteBruto=SumImporteBruto + rstAux("suma")
		dto1=miround((null_z(rstAux("suma"))*null_z(rst_albaran("descuento")))/100,2)
		dto2=miround(((null_z(rstAux("suma"))-dto1)*null_z(rst_albaran("descuento2")))/100,2)
		total_descuento=dto1+dto2+dto3
		SumTotalDto=SumTotalDto + null_z(total_descuento)
		base_imponible=null_z(rstAux("suma"))-null_z(total_descuento)
		SumBaseImponible=SumBaseImponible + null_z(base_imponible)

		total_iva=(null_z(base_imponible)*rstAux("iva"))/100
		SumTotalIva=SumTotalIva + null_z(total_iva)
		if TieneRE <> 0 then
			re=d_lookup("re","tipos_iva","tipo_iva='" & rstAux("iva") & "'",session("dsn_cliente"))
		else
			re=0
		end if
		total_re=(null_z(base_imponible)*re)/100
		SumTotalRE=SumTotalRE + null_z(total_re)
		rstAux.moveNext
	wend
	SumTotalIvaTotal=SumTotalIvaTotal+miround(null_z(SumTotalIva),n_decimales)
	SumTotalRETotal=SumTotalRETotal+miround(null_z(SumTotalRE),n_decimales)

	rstAux.close

	rst_albaran("importe_bruto")=SumImporteBruto
	rst_albaran("total_descuento")=SumTotalDto
	rst_albaran("base_imponible")=SumBaseImponible
	rst_albaran("total_iva")=miround(SumTotalIvaTotal,n_decimales)
	rst_albaran("total_re")=miround(SumTotalRETotal,n_decimales)
	SumTotalRF=(null_z(SumBaseImponible)*null_z(rst_albaran("recargo")))/100
	rst_albaran("total_recargo")=miround(SumTotalRF,n_decimales)
	if nz_b(rst_albaran("IRPF_Total"))=0 then
		baseImp=null_z(SumBaseImponible)
	else
		baseImp=null_z(SumBaseImponible)+null_z(rst_albaran("total_iva"))+null_z(rst_albaran("total_re"))+null_z(rst_albaran("total_recargo"))
	end if
	SumTotalIRPF=(null_z(baseImp)*null_z(rst_albaran("irpf")))/100
	rst_albaran("total_irpf")=miround(SumTotalIRPF,n_decimales)
	SumTotalImporte=null_z(SumBaseImponible)+null_z(rst_albaran("total_iva"))+null_z(rst_albaran("total_re"))+null_z(rst_albaran("total_recargo"))-null_z(rst_albaran("total_irpf"))
	rst_albaran("total_albaran")=miround(SumTotalImporte,2)

	rst_albaran.Update
	rst_albaran.Close
end sub

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


'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas de datos.
	'fdesde:
	'fhasta:
	'nproveedor:
function CadenaBusqueda(fdesde,fhasta,nproveedor)
	if nproveedor > "" then
		CadenaBusqueda = " nproveedor='" & nproveedor & "' and fecha>='" & fdesde & "' and fecha<='" & fhasta & "'and nalbaran is null and nfactura is null order by fecha desc,npedido desc,nproveedor"
	else
		CadenaBusqueda = " fecha>='" & fdesde & "' and fecha<='" & fhasta & "' and nalbaran is null and nfactura is null order by fecha desc,npedido desc,nproveedor"
	end if
end function

function GuardarCabeceraAlbaran(nalbaran,nserie,falbaran, nalbaran_pro,ncuentaAnt,variasCuentas,NumIncoterms,NumFob)
	if nalbaran="" then
		'Obtener el último nº de albaran de la tabla series.
        rstAux.cursorlocation=2
		rstAux.Open "select contador,ultima_fecha from series where nserie='" & nserie & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic

		num=rstAux("contador")+1
		num=string(5-len(cstr(num)),"0") + cstr(num)

		'Actualizar el nº de proveedor de CONFIGURACION.
		rstAux("contador")=rstAux("contador")+1
		rstAux("ultima_fecha")=date
		rstAux.Update
		rstAux.Close

		'Crear un nuevo registro.
        rst_albaran.cursorlocation=2
		rst_albaran.Open "select * from albaranes_pro", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst_albaran.AddNew
		rst_albaran("nalbaran")=nserie+ right(falbaran,2) + num
	GuardarCabeceraAlbaran = nserie+ right(falbaran,2) + num

''ricardo 20/2/2003
''si no se pone nalbaran_pro, se pondra el nalbaran
		if nalbaran_pro & "">"" then
			rst_albaran("nalbaran_pro") = nalbaran_pro
		else
			rst_albaran("nalbaran_pro") =trimCodEmpresa(rst_albaran("nalbaran"))
		end if
		rst_albaran("serie")		= nserie
		rst_albaran("nproveedor")	= rst_pedido("nproveedor")
		rst_albaran("fecha")		= falbaran
		'nfactura lo dejamos con valor nulo
		rst_albaran("descuento")	= rst_pedido("descuento")
		rst_albaran("descuento2")	= rst_pedido("descuento2")
		rst_albaran("descuento3")	= rst_pedido("descuento3")
		rst_albaran("importe_bruto")	= rst_pedido("importe_bruto")
		rst_albaran("base_imponible")	= rst_pedido("base_imponible")
		rst_albaran("total_descuento")	= rst_pedido("total_descuento")
		rst_albaran("total_iva")		= rst_pedido("total_iva")
		rst_albaran("total_re")			= rst_pedido("total_re")
		rst_albaran("recargo")			= rst_pedido("recargo")
		rst_albaran("total_recargo")	= rst_pedido("total_recargo")
		rst_albaran("irpf")				= rst_pedido("irpf")
		rst_albaran("irpf_total")		= rst_pedido("irpf_total")
		rst_albaran("total_irpf")		= rst_pedido("total_irpf")
		rst_albaran("total_albaran")	= rst_pedido("total_pedido")
		rst_albaran("divisa")			= rst_pedido("divisa")
		rst_albaran("forma_pago")		= rst_pedido("forma_pago")
		rst_albaran("facturado")=0
		rst_albaran("observaciones")= "Albarán generado del/los pedidos :"
		rst_albaran("cod_proyecto")= rst_pedido("cod_proyecto")
		rst_albaran("ncuenta")= rst_pedido("ncuenta")
		'FLM:170309:añado cuenta de abono del proveedor
		'rst_albaran("ncuenta_pro")= rst_pedido("ncuenta_pro")


		rst_albaran("valorado")=0
		rst_albaran("ahora")=0
		rst_albaran("tipo_pago")	= rst_pedido("tipo_pago")
		if NumIncoterms=1 then
			rst_albaran("incoterms")=rst_pedido("incoterms")
		end if
		if NumFob=1 then
			rst_albaran("fob")=rst_pedido("fob")
		end if

		' JMA 20/12/04. Copiar campos personalizables
		rst_albaran("campo01")=rst_pedido("campo01")
		rst_albaran("campo02")=rst_pedido("campo02")
		rst_albaran("campo03")=rst_pedido("campo03")
		rst_albaran("campo04")=rst_pedido("campo04")
		rst_albaran("campo05")=rst_pedido("campo05")
		rst_albaran("campo06")=rst_pedido("campo06")
		rst_albaran("campo07")=rst_pedido("campo07")
		rst_albaran("campo08")=rst_pedido("campo08")
		rst_albaran("campo09")=rst_pedido("campo09")
		rst_albaran("campo10")=rst_pedido("campo10")
		' JMA 20/12/04. FIN Copiar campos personalizables.

		'******************** Manejo de domicilios
		Dom=Domicilios("COMPRAS","ALB_ENV_PROV",rst_albaran("nproveedor"),rst_albaran)
	else
		rst_albaran.Open "select * from albaranes_pro where nalbaran='"+nalbaran+"'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst_albaran("importe_bruto")	= rst_albaran("importe_bruto")+rst_pedido("importe_bruto")
		rst_albaran("base_imponible")	= rst_albaran("base_imponible")+rst_pedido("base_imponible")
		rst_albaran("total_descuento")	= rst_albaran("total_descuento")+rst_pedido("total_descuento")
		rst_albaran("total_iva")		= rst_albaran("total_iva")+rst_pedido("total_iva")
		rst_albaran("total_re")			= rst_albaran("total_re")+rst_pedido("total_re")
		rst_albaran("recargo")			= rst_albaran("recargo")
		rst_albaran("total_recargo")	= rst_albaran("total_recargo")+rst_pedido("total_recargo")
		rst_albaran("irpf")				= rst_albaran("irpf")
		rst_albaran("total_irpf")	= rst_albaran("total_irpf")+rst_pedido("total_irpf")
		rst_albaran("total_albaran")	= rst_albaran("total_albaran")+rst_pedido("total_pedido")

		if variasCuentas=0 then
			if ncuentaAnt<>"" and rst_pedido("ncuenta")<>ncuentaAnt then
				ncuentaPro = d_lookup("cuenta_cargo","proveedores","nproveedor='" & rst_pedido("nproveedor") & "'",session("dsn_cliente"))
				rst_albaran("ncuenta") = ncuentaPro
				ncuentaAnt = ncuentaPro
				variasCuentas=1
			else
				rst_albaran("ncuenta") = rst_pedido("ncuenta")
				ncuentaAnt = rst_pedido("ncuenta")
			end if
		end if

		if rst_albaran("forma_pago")<>rst_pedido("forma_pago") then
			rst_albaran("forma_pago")	= limpiaCadena(request.form("banana"))
		end if
		if rst_albaran("tipo_pago")<>rst_pedido("tipo_pago") then
			rst_albaran("tipo_pago")	= limpiaCadena(request.form("banana"))
		end if

		' JMA 20/12/04. Copiar campos personalizables
		if rst_albaran("campo01")<>rst_pedido("campo01") then rst_albaran("campo01")=p_parametro_nulo
		if rst_albaran("campo02")<>rst_pedido("campo02") then rst_albaran("campo02")=p_parametro_nulo
		if rst_albaran("campo03")<>rst_pedido("campo03") then rst_albaran("campo03")=p_parametro_nulo
		if rst_albaran("campo04")<>rst_pedido("campo04") then rst_albaran("campo04")=p_parametro_nulo
		if rst_albaran("campo05")<>rst_pedido("campo05") then rst_albaran("campo05")=p_parametro_nulo
		if rst_albaran("campo06")<>rst_pedido("campo06") then rst_albaran("campo06")=p_parametro_nulo
		if rst_albaran("campo07")<>rst_pedido("campo07") then rst_albaran("campo07")=p_parametro_nulo
		if rst_albaran("campo08")<>rst_pedido("campo08") then rst_albaran("campo08")=p_parametro_nulo
		if rst_albaran("campo09")<>rst_pedido("campo09") then rst_albaran("campo09")=p_parametro_nulo
		if rst_albaran("campo10")<>rst_pedido("campo10") then rst_albaran("campo10")=p_parametro_nulo
		' JMA 20/12/04. Copiar campos personalizables

	GuardarCabeceraAlbaran = nalbaran
	end if

	'Actualizar el registro.
	rst_albaran.Update
	rst_albaran.Close

end function

'****************************************************************************************************************

sub GuardarDetalleAlbaran(nalbaran,npedido,ndecimales)
    rst_pedido_det.cursorlocation=2
	rst_pedido_det.Open "select * from albaranes_pro where nalbaran='"+nalbaran+"'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		rst_pedido_det("observaciones")=rst_pedido_det("observaciones") & " " & trimCodEmpresa(npedido)
		rst_pedido_det.update
	rst_pedido_det.close


	''ricardo 7-3-2003
	''para la personalizacion 253 de JUAREZ RECHE
	''se cambia el select

	'''strselect="select * from detalles_ped_pro where npedido='"+npedido+"' order by item"

	''por este otro
	strselect="select d.npedido,d.referencia,d.almacen,d.item"
	strselect=strselect & ",d.cantidad as cant_anterior,case when e.cantidad is not null then e.cantidad else d.cantidadpend end as cantidad"
	strselect=strselect & ",d.pvp,d.importe,d.descripcion,d.descuento,d.descuento2,d.iva,d.re,d.mainitem"
	strselect=strselect & ",e.mi_nserie,e.nserie "
	strselect=strselect & " from detalles_ped_pro as d with(NOLOCK) "
	strselect=strselect & " left outer join [" & session("usuario") & "] as e on e.item=d.item and e.npedido=d.npedido "
	strselect=strselect & ",pedidos_pro as p "
	strselect=strselect & " where p.npedido=d.npedido and d.npedido='" & npedido & "'"
	strselect=strselect & " order by d.item"
    rst_pedido_det.cursorlocation=3
	rst_pedido_det.Open strselect, session("dsn_cliente")
	while not rst_pedido_det.EOF

		varios = true

		'Obtener el último nº de detalle del albaran.
		item=d_max("item","detalles_alb_pro","nalbaran='" & nalbaran & "'",session("dsn_cliente"))+1


			'Crear un nuevo registro.
			rst_albaran_det.Open "select * from detalles_alb_pro where nalbaran='" & nalbaran & "'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rst_albaran_det.AddNew
			rst_albaran_det("nalbaran")		= nalbaran
			rst_albaran_det("referencia")	= rst_pedido_det("referencia")
			rst_albaran_det("almacen")		= rst_pedido_det("almacen")
			rst_albaran_det("item")			= item
			rst_albaran_det("cantidad")		= rst_pedido_det("cantidad")
			rst_albaran_det("descripcion")	= rst_pedido_det("descripcion")
			rst_albaran_det("pvp")			= rst_pedido_det("pvp")
			rst_albaran_det("descuento")	= rst_pedido_det("descuento")
			rst_albaran_det("descuento2")	= rst_pedido_det("descuento2")

''ricardo 10/3/2003
''el importe sera de calcular con los descuentos
			temp=rst_albaran_det("pvp")*rst_albaran_det("cantidad")
			temp2=temp-(temp*rst_albaran_det("descuento"))/100
			temp3=temp2-(temp2*rst_albaran_det("descuento2"))/100

			'rst_albaran_det("importe")= rst_pedido_det("importe")
			rst_albaran_det("importe")=miround(temp3,ndecimales)

'''''''''''''''''

			rst_albaran_det("iva")			= rst_pedido_det("iva")
			rst_albaran_det("mainitem")			= rst_pedido_det("mainitem")

			rst_albaran_det("re")			= null_z(rst_pedido_det("re"))
			rst_albaran_det("npedido")		= rst_pedido_det("npedido")
			rst_albaran_det("itempedido")   = rst_pedido_det("item")

''ricardo 10-3-2003
''ahora la cantidad,el nserie se sacan de la tabla temporal,si hubo modificaciones
			if rst_pedido_det("nserie") & "">"" then
				nserie=mid(rst_pedido_det("nserie"),3,len(rst_pedido_det("nserie"))-4) 'quitamos los parentesis y la primera y ultima comilla
				ListaN=split(nserie,"','",-1,1)
				listaTEquipos="('"
				k2=1
				Redim ListaNumeros(k2)
				for i=0 to ubound(ListaN)
					if len(ListaN(i))>0 then
						listaTEquipos=listaTEquipos & ListaN(i) & "','"
						ListaNumeros(k2)=ListaN(i)
						k2=k2+1
						Redim Preserve ListaNumeros(k2)
					end if
				next
				if listaTEquipos="('" then
					listaTEquipos=""
				else
					listaTEquipos=mid(listaTEquipos,1,len(listaTEquipos)-2) & ")" 'Quitamos la última coma y comilla y cerramos el paréntesis
				end if
			else
				nserie=NULL
				listaTEquipos=""
				k2=0
				varios=false
			end if


'''''''''''''''
        rstAux.cursorlocation=3
	    rstAux.Open "select referencia from articulos with(NOLOCK) where referencia='" & rst_pedido_det("referencia") & "'and ctrl_nserie=1", session("dsn_cliente")
		if rstAux.eof then 'no obliga a numero de serie
			varios = false
		else
			if rst_pedido_det("cantidad")=1 then varios = false
		end if
		rstAux.close

			if varios=true then
				rst_albaran_det("mi_nserie")	= "VARIOS"
				rst_albaran_det("nserie")		= "VARIOS"
			else
				if nserie & "">"" then
					rst_albaran_det("mi_nserie")	= rst_pedido_det("referencia")  & nserie
					rst_albaran_det("nserie")		= nserie
				end if
			end if

			'Actualizar el registro.
			rst_albaran_det.Update
			rst_albaran_det.Close

			'Actualizar Stock
''ricardo 10-3-2003
''se llamara a la funcion stock, con las cantidades sin modificar
			ActualizaStocks "first_save","PEDIDO_PRO->ALBARAN_PRO",rst_pedido_det("referencia"),rst_pedido_det("almacen"),rst_pedido_det("cant_anterior"),"",session("dsn_cliente")
''''''
''ahora llamaremos al stock con las cantidades nuevas
			if rst_pedido_det("cant_anterior")<>rst_pedido_det("cantidad") then
				cant_a_poner=rst_pedido_det("cant_anterior")-rst_pedido_det("cantidad")
''ricardo 3-9-2003 se quita esta, ya que si no el stock sale mal con el albaran
'' ahora se controla desde los triggers
				'rstAux.Open "select p_recibir,p_servir,stock,stock_minimo from almacenar where articulo='" & rst_pedido_det("referencia") & "' and almacen='" & rst_pedido_det("almacen") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				'rstAux("stock")=rstAux("stock")-cant_a_poner
				'rstAux.update
				'rstAux.close
'''''''
				'Comprobamos si el artículo tiene escandallo fijo para descontar las unidades del stock
				escvariable=d_lookup("escvariable","articulos","referencia='" & rst_pedido_det("referencia") & "'",session("dsn_cliente"))
				if escvariable="1" or escvariable=1 or escvariable="Verdadero" or escvariable="true" or escvariable=true then
				else
					'StocksEscandallo rst_pedido_det("referencia"),rst_pedido_det("almacen"),cant_a_poner,"stock","-",session("dsn_cliente")
				end if
			end if
'''''

			'Vamos a meter los numeros de serie en equipos

			InsertarHistorialNserie "OK1",listaTEquipos,k2-1,"ALBARAN DE PROVEEDOR",nalbaran,item,rst_pedido_det("referencia"),rst_pedido_det("cantidad"),rst_pedido_det("almacen"),"","first_save"

		item = item + 1
		rst_pedido_det.MoveNext
	wend
	rst_pedido_det.Close
end sub

'****************************************************************************************************************

sub GuardarConceptoAlbaran(nalbaran,npedido,ndecimales)
    rst_pedido_con.cursorlocation=3
	rst_pedido_con.Open "select * from conceptos_ped_pro with(NOLOCK) where npedido='"+npedido+"'", session("dsn_cliente")
	while not rst_pedido_con.EOF

		'Obtener el último nº de detalle del albaran.
		nconcepto=d_max("nconcepto","conceptos_alb_pro","nalbaran='" & nalbaran & "'",session("dsn_cliente"))+1

		'Crear un nuevo registro.
			rst_albaran_con.Open "select * from conceptos_alb_pro", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rst_albaran_con.AddNew
			rst_albaran_con("nconcepto")	= nconcepto
			rst_albaran_con("nalbaran")		= nalbaran
			rst_albaran_con("descripcion")	= rst_pedido_con("descripcion")
			rst_albaran_con("cantidad")		= rst_pedido_con("cantidad")
			rst_albaran_con("pvp")			= rst_pedido_con("pvp")
			rst_albaran_con("importe")		= rst_pedido_con("importe")
			rst_albaran_con("descuento")	= rst_pedido_con("descuento")
			rst_albaran_con("iva")			= rst_pedido_con("iva")
			rst_albaran_con("re")			= rst_pedido_con("re")
			nconcepto = nconcepto + 1
			'Actualizar el registro.
			rst_albaran_con.Update
			rst_albaran_con.Close
		rst_pedido_con.MoveNext
	wend
	rst_pedido_con.Close
end sub

'******************************************************************************

sub GuardarPagosAlbaran(nalbaran,npedido,ndecimales)
    rst_pedido_con.cursorlocation=3
	rst_pedido_con.Open "select * from pagos_ped_pro with(NOLOCK) where npedido='"+npedido+"'", session("dsn_cliente")
	while not rst_pedido_con.EOF

		'Obtener el último nº de detalle del albaran.
		npago=d_max("npago","pagos_alb_pro","nalbaran='" & nalbaran & "'",session("dsn_cliente"))+1

		'Crear un nuevo registro.
			rst_albaran_con.Open "select * from pagos_alb_pro", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			rst_albaran_con.AddNew
			rst_albaran_con("npago")= npago
			rst_albaran_con("nalbaran")= nalbaran
			rst_albaran_con("fecha")= rst_pedido_con("fecha")
			rst_albaran_con("importe")= rst_pedido_con("importe")
			rst_albaran_con("descripcion")= rst_pedido_con("descripcion")
			rst_albaran_con("ncaja")=rst_pedido_con("ncaja")
			rst_albaran_con("nanotacion")=rst_pedido_con("nanotacion")
			rst_albaran_con("medio")=rst_pedido_con("medio")
			npago = npago + 1
			'Actualizar el registro.
			rst_albaran_con.Update
			rst_albaran_con.Close
		rst_pedido_con.MoveNext
	wend
	rst_pedido_con.Close
end sub

'******************************************************************

'Validar las lineas con numeros de serie
function validarNSerie()

	'seleccion="select * from detalles_ped_pro where " & strwhere
	'seleccion=seleccion & " and referencia in (select referencia from articulos where ctrl_nserie=1) order by npedido,item"

	seleccion="select d.npedido,d.referencia,d.almacen,d.item"
	seleccion=seleccion & ",d.cantidad as cant_anterior,case when e.cantidad is not null then e.cantidad else d.cantidad end as cantidad"
	seleccion=seleccion & ",d.pvp,d.importe,d.descripcion,d.descuento,d.descuento2,d.iva,d.re,d.mainitem"
	seleccion=seleccion & ",e.mi_nserie,e.nserie "
	seleccion=seleccion & " from detalles_ped_pro as d with(NOLOCK) "
	seleccion=seleccion & " left outer join [" & session("usuario") & "] as e on e.item=d.item and e.npedido=d.npedido "
	seleccion=seleccion & ",articulos as a "
	seleccion=seleccion & " where " & strwhere
	seleccion=seleccion & " and d.referencia=a.referencia and a.ctrl_nserie=1 "
	seleccion=seleccion & " order by d.npedido,d.item"

	Correcto=true
    rst.cursorlocation=3
	rst.Open seleccion, session("dsn_cliente")
	while not rst.EOF
		'construimos el nombre del campo a leer los numeros de serie
		'nombreSpan="A" & Limpiar(rst("npedido")) & "B"
		'nombreCampo=nombreSpan & completar(cstr(rst("item")),3,"0")
		'p_nserie=limpiaCadena(request.form(nombreCampo))

		'ListaN=split(p_nserie,chr(13)&chr(10),-1,1)
		'listaTEquipos="('"
		'k=1
		'Redim ListaNumeros(k)
		'for i=0 to ubound(ListaN)
		'	if len(ListaN(i))>0 then
		'		listaTEquipos=listaTEquipos & ListaN(i) & "','"
		'		ListaNumeros(k)=ListaN(i)
		'		k=k+1
		'		Redim Preserve ListaNumeros(k)
		'	end if
		'next
		'if listaTEquipos="('" then
		'	listaTEquipos=""
		'else
		'	listaTEquipos=mid(listaTEquipos,1,len(listaTEquipos)-2) & ")" 'Quitamos la última coma y comilla y cerramos el paréntesis
		'end if

		listaTEquipos=rst("nserie")
		p=rst("cantidad")
		mensajeTratEquipos=TratarEquipos(listaTEquipos,p,"PED-ALB-PRO",rst("npedido"),rst("item"),rst("referencia"),rst("cantidad"),rst("almacen"),"","first_save")
		if mid(mensajeTratEquipos,1,2)<>"OK" then
			Correcto=false
			rst.movelast
			%><script language="javascript" type="text/javascript">
			window.alert("<%=mensajeTratEquipos%>\n<%=LitPedido%> <%=trimCodEmpresa(rst("npedido"))%> <%=LitEnLaLinea%> <%=rst("item")%>");
			</script><%
		end if
		rst.movenext
	wend
	rst.close
	validarNSerie=Correcto
end function


'******************************************************************************
'Valida los descuentos de los pedidos
sub validar()

	valido=true
	strwhere ="("
	while h_nfilas>0
		npedido=limpiaCadena(request.form("check"+cstr(h_nfilas)))
		if npedido>"" then
			strwhere = strwhere +"'"+npedido+"',"
		end if
		h_nfilas = h_nfilas -1
	wend

	strwhereAux=strwhere
	strwhere = "npedido in "+strwhereAux+"'xxxxxx')"
	strwhere2 = "npedido in "+strwhereAux+"'xxxxxx')"
	strwhere3 = "npedido not in "+strwhereAux+"'xxxxxx')"
    rst_pedido.cursorlocation=3
	rst_pedido.Open "select distinct divisa,descuento,descuento2,recargo,irpf from pedidos_pro with(NOLOCK) where " + strwhere, session("dsn_cliente")
	if 	rst_pedido.recordCount <> 1 then%>
		<script language="javascript" type="text/javascript">
		    window.alert("<%=LitMsgAlbComprasIncong%>");
		</script>
		<%strwhere = "npedido in ('xxxxxx')"
		valido=false
	end if
	rst_pedido.close

    if valido=true then
        rst_pedido.cursorlocation=3
        rst_pedido.Open "select distinct divisa, descuento, descuento2,recargo,irpf from(select divisa, descuento, descuento2,recargo,irpf from albaranes_pro with(NOLOCK) where nalbaran like '"&session("ncliente")&"%' and fecha='"&falbaran&"' and nalbaran_pro ='"&nalbaran_pro&"' and nproveedor ='"&nproveedor2&"' union all select distinct divisa,descuento,descuento2,recargo,irpf from pedidos_pro with(NOLOCK) where " + strwhere+") as tabla", session("dsn_cliente")
        'select distinct divisa,descuento,descuento2,recargo,irpf from pedidos_pro where " + strwhere, session("dsn_cliente"),adUseClient, adLockReadOnly
		if 	rst_pedido.recordCount <> 1 then%>
			<script language="javascript" type="text/javascript">
			    window.alert("<%=LitMsgAlbComprasIncong%>");
			</script>
			<%strwhere = "npedido in ('xxxxxx')"
			valido=false
		end if
		rst_pedido.Close
    end if

	if valido=true then
        rst_pedido.cursorlocation=3
		rst_pedido.Open "select distinct irpf_total from pedidos_pro with(NOLOCK) where " + strwhere2, session("dsn_cliente")
		if 	rst_pedido.recordCount <> 1 then%>
			<script language="javascript" type="text/javascript">
			    window.alert("<%=LitMsgAlbComprasIncong2%>");
			</script>
			<%strwhere = "npedido in ('xxxxxx')"
			valido=false
		end if
		rst_pedido.Close
	end if
end sub

'******************************************************************************
'Convierte los pedidos en albaranes
function convertir(existe)
	strselect="select * from pedidos_pro where " + strwhere
	if strwhere & "">"" then
		strselect=strselect & " and nalbaran is null and nfactura is null"
	else
		strselect=strselect & " nalbaran is null and nfactura is null"
	end if
	
	rst.open "delete from [" & session("usuario") & "] where "+strwhere3 ,session("dsn_cliente")
	rst_pedido.Open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if not rst_pedido.eof then
		forma_pago=rst_pedido("forma_pago")
		divisa=iif(rst_pedido("divisa")>"",rst_pedido("divisa"),d_lookup("codigo","divisas","moneda_base=1 and codigo like'" & session("ncliente") & "%'",session("dsn_cliente")))
		descuento=null_z(rst_pedido("descuento"))
		descuento2=null_z(rst_pedido("descuento2"))
		cod_proyecto=null_s(rst_pedido("cod_proyecto"))
		nerror=0
		nerror2=0
		nerror3=0
		nerror4=0
		lista_pedidos="('"
		while not rst_pedido.EOF and nerror=0 and nerror2=0

			if rst_pedido("forma_pago")<>forma_pago or (isnull(rst_pedido("forma_pago")) and forma_pago>"") or (rst_pedido("forma_pago")>"" and isnull(forma_pago))then
				nerror=1
			end if
			if rst_pedido("divisa")<>divisa or (isnull(rst_pedido("divisa")) and divisa>"") or (rst_pedido("divisa")>"" and isnull(divisa))then
				nerror2=1
			end if
			if rst_pedido("descuento")<>descuento or (isnull(rst_pedido("descuento")) and descuento>"") or (rst_pedido("descuento")>"" and isnull(descuento))then
				nerror3=1
			end if
			if rst_pedido("descuento2")<>descuento2 or (isnull(rst_pedido("descuento2")) and descuento2>"") or (rst_pedido("descuento2")>"" and isnull(descuento2))then
				nerror3=1
			end if
			if ucase(rst_pedido("cod_proyecto"))<>ucase(cod_proyecto) or (isnull(rst_pedido("cod_proyecto")) and cod_proyecto>"") or (rst_pedido("cod_proyecto")>"" and isnull(cod_proyecto))then
				nerror4=1
			end if
			lista_pedidos=lista_pedidos & trimcodempresa(rst_pedido("npedido")) & "','"
			pedido=d_lookup("npedido","[" & session("usuario") & "] ","npedido='"&rst_pedido("npedido")&"'",session("dsn_cliente"))
			if pedido="" then
				'EJM 30/10/2006. Se tiene encuenta el campo cantidad2
				'rst.open "insert into [" & session("usuario") & "] (npedido,item,cantidad,referencia) select npedido,item,cantidadPend,referencia from detalles_ped_pro where npedido='"&rst_pedido("npedido")&"' ",session("dsn_cliente")
				rst.open "insert into [" & session("usuario") & "] (npedido,item,cantidad,referencia) select npedido,item,cantidadPend,referencia from detalles_ped_pro where npedido='"&rst_pedido("npedido")&"' and cantidadPend<>0 ",session("dsn_cliente")				
                rst.open "insert into [" & session("usuario") & "] (npedido) select npedido from pedidos_pro where npedido='"&rst_pedido("npedido")&"' and npedido not in (select npedido from detalles_ped_pro where npedido='"&rst_pedido("npedido")&"') ",session("dsn_cliente")
				rst.open "insert into [" & session("usuario") & "] (npedido) select distinct npedido from conceptos_ped_pro where npedido='"&rst_pedido("npedido")&"' and npedido not in (select npedido from detalles_ped_pro where npedido='"&rst_pedido("npedido")&"' and cantidadPend<>0) ",session("dsn_cliente")
			end if
			rst_pedido.MoveNext
			'if not rst_pedido.eof then
				'forma_pago=rst_pedido("forma_pago")
				'divisa=iif(rst_pedido("divisa")>"",rst_pedido("divisa"),d_lookup("codigo","divisas","moneda_base=1",session("dsn_cliente")))
			'end if
		wend
		if lista_pedidos="('" then
			lista_pedidos=""
		else
			lista_pedidos=mid(lista_pedidos,1,len(lista_pedidos)-2) & ")"
		end if
		if nerror=1 or nerror2=1 or nerror3=1 or nerror4=1 then
			rst_pedido.Close
			if nerror=1 then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitNoConvDistFormPagoConvPedAlb%>");
				</script><%
			end if
			if nerror2=1 then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitNoConvDistDivConvPedAlb%>");
				</script><%
			end if
			if nerror3=1 then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitNoConvDistDescConvPedAlb%>");
				</script><%
			end if
			if nerror4=1 then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitNoConvDistProyConvPedAlb%>");
				</script><%
			end if
			convertir=""
		else
			rst_pedido.movefirst
			nalbaran =""
			ncuentaAnt=""
			IncotermsDistintos=CalcularIncotermsDistintos(lista_pedidos,")","pedidos_pro",rst_pedido("nproveedor"),0,1)
			FobDistintos=CalcularFobDistintos(lista_pedidos,")","pedidos_pro",rst_pedido("nproveedor"),0,1)
			variasCuentas=0
			conn.open session("dsn_cliente")
			'''EBF Añadido para hacer la conversion de pedidos a albaranes de forma masiva 17/01/2006'''
			''MPC 11/09/2008 Si existe el albaran se lanza el procedimiento de actualización
			if existe = 0 then
    			strselect="EXEC ConvertirPedDev_Alb @serie='"&nserie&"',@fechaAlbaran='"&falbaran&"',@nalbaran_pro='"&nalbaran_pro&"',@nusuario='"&session("usuario")&"',@session_ncliente='"&session("ncliente")&"',@tipo='PEDIDO',@lista='"&replace(replace(replace(lista_pedidos,"'",""),"(",""),")","")&"',@modulo=''"
    	    else
    	        strselect="EXEC ActualizarPedDev_Alb @serie='"&nserie&"',@fechaAlbaran='"&falbaran&"',@nalbaran_pro='"&nalbaran_pro&"',@nusuario='"&session("usuario")&"',@session_ncliente='"&session("ncliente")&"',@tipo='PEDIDO',@lista='"&replace(replace(replace(lista_pedidos,"'",""),"(",""),")","")&"',@modulo=''"
    		end if
			set rstAux = conn.execute(strselect)
			
			Primero=rstAux(0)
			Ultimo=rstAux(1)
			error=rstAux(2)
			cadenaAuditoria=rstAux(3)

			if Primero>"" and Ultimo>"" and error=0 then
				nalbaran_pro=d_lookup("nalbaran_pro","albaranes_pro","nalbaran='"&Primero&"'",session("dsn_cliente"))
				auditar_ins_bor session("usuario"),cadenaAuditoria,"","alta","","","conver_dev_ped_alb_proMASIVO"
				
''ricardo 29-1-2009 se añaden los pametros modd,modi,modp que no se estaban pasando actualmente
error_aux222=error
existe_2222=existe
ObtenerParametros("albaranes_pro_det")
if modd & ""="" then
    modd="SI"
end if
if modp & ""="" then
    modp="SI"
end if
if modi & ""="" then
    modi=""
end if
error=error_aux222
existe=existe_2222
				%>
				<script type="text/javascript" language="javascript">
				    <%if existe = 0 then%>
					alert("<%=LitMsgAlbaranGenerado%> <%=nalbaran_pro%> ");
					<%else%>
					alert("<%=LitMsgAlbaranActualizado%> <%=nalbaran_pro%> ");
					<%end if%>
					if (document.pedpro_albpro_param.cvsimp.value!="SI")
					{
						if (window.confirm("<%=LitMsgDeseaVer%>")==true)
						{
						    //ricardo 29-1-2009 se añaden los pametros modd,modi,modp que no se estaban pasando actualmente
						    cadena_a_abrir="../search_layout.asp?pag1=compras/albaranes_pro.asp?ndoc=<%=primero%>";
						    cadena_a_abrir=cadena_a_abrir + "?mode=browse";
						    cadena_a_abrir=cadena_a_abrir + "?modp=<%=enc.EncodeForJavascript(null_s(modp))%>?modd=<%=enc.EncodeForJavascript(null_s(modd))%>?modi=<%=enc.EncodeForJavascript(null_s(modi))%>?titulo=<%=LitDatosAlbaran%> <%=enc.EncodeForJavascript(null_s(nalbaran_pro))%>";
						    cadena_a_abrir=cadena_a_abrir + "&pag2=compras/albaranes_pro_bt.asp";
						    cadena_a_abrir=cadena_a_abrir + "&pag3=compras/deliveryNote_pro_lsearch.asp";
							AbrirVentana(cadena_a_abrir,'P',<%=altoventana%>,<%=anchoventana%>) 
						}
					}
				</script>
			<%elseif error=1 then%>
				<script type="text/javascript" language="javascript">
				    window.alert("<%=LitMsgSinDireccionPrinc%>");
				</script>
			<%elseif error=2 then%>
				<script type="text/javascript" language="javascript">
				    window.alert("<%=LitMsgNumSeriesRepetidos%>");
				</script>
			<%elseif error=3 then%>
				<script type="text/javascript" language="javascript">
				    window.alert("<%=LitMsgDocExistRevContConv%>");
				</script>
			<%else%>
				<script type="text/javascript" language="javascript">
				    window.alert("<%=LitMsgAlbaranNoGenerado%>");
				</script>
			<%end if
			rstAux.close
			mode="select1"

			while not rst_pedido.EOF and false
				nalbaran = GuardarCabeceraAlbaran(nalbaran,nserie,falbaran,nalbaran_pro,ncuentaAnt,variasCuentas,IncotermsDistintos,FobDistintos)
				if variasCuentas=0 then
					if ncuentaAnt<>"" and rst_pedido("ncuenta")<>ncuentaAnt then
						ncuentaPro = d_lookup("cuenta_cargo","proveedores","nproveedor='" & rst_pedido("nproveedor") & "'",session("dsn_cliente"))
						ncuentaAnt = ncuentaPro
						variasCuentas=1
					else
						ncuentaAnt = rst_pedido("ncuenta")
					end if
				end if

''ricardo 13-3-2003
''se pone esto,ya que el valor de ndecimales , no se calculaba por ningun sitio
				if isnull(rst_pedido("divisa")) then
					ndecimales=d_lookup("ndecimales","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
				else
					ndecimales=d_lookup("ndecimales","divisas","codigo='" & rst_pedido("divisa") & "'",session("dsn_cliente"))
				end if
'''''''''
				GuardarDetalleAlbaran nalbaran,rst_pedido("npedido"),ndecimales
				GuardarConceptoAlbaran nalbaran,rst_pedido("npedido"),ndecimales
				GuardarPagosAlbaran nalbaran,rst_pedido("npedido"),ndecimales
				actualizar_totales nalbaran
				
				cantpendiente=d_sum("cantidadpend","detalles_ped_pro","npedido='" & rst_pedido("npedido") & "' and mainitem is null",session("dsn_cliente"))
				if cantpendiente=0 then
					rst_pedido("nalbaran")=nalbaran
				end if
				rst_pedido.Update
				rst_pedido.MoveNext
			wend
			rst_pedido.Close
			nalbaran_param = nalbaran
			'vamos a auditar
			'auditar_ins_bor session("usuario"),nalbaran,"","alta","","","conver_ped_alb_pro"
			'auditar_ins_bor session("usuario"),nalbaran,"","alta","","","equipos_albaranes_pro"
			convertir="OK"
''ricardo 10-3-2003
''si hay que hacer el pedido bis

		if genpedbis=1 then
			rst.open "select * from albaranes_pro where nalbaran='" & nalbaran & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			PreguntarSerie=0
			rstPedidos.open "select distinct npedido from detalles_alb_pro where nalbaran='" & rst("nalbaran") & "' and npedido is not null ",session("dsn_cliente"),adUseClient, adLockReadOnly
			lista = "('xxxxxx'"
			while not rstPedidos.eof
				lista = lista & ",'" & rstPedidos("npedido") & "'"
				rstPedidos.movenext
			wend
			lista=lista & ")"
			rstPedidos.movefirst
			'if rstPedidos.Recordcount=1 then
				PedidoAlb=d_lookup("npedido","detalles_alb_pro","nalbaran='" & rst("nalbaran") & "' and npedido is not null",session("dsn_cliente"))
				SeriePedido=d_lookup("serie","pedidos_pro","npedido='" & PedidoAlb & "' and npedido is not null",session("dsn_cliente"))
			'else
			'	rstSeries.Open "select distinct serie from pedidos_pro where npedido in " & lista,session("dsn_cliente"),adUseClient, adLockReadOnly
			'	if rstSeries.Recordcount=1 then
			'		SeriePedido=rstSeries("serie")
			'	else
			'		PreguntarSerie=d_lookup("preg_ped_bis","configuracion","",session("dsn_cliente"))
			'		SeriePedido=d_lookup("serie_ped_bis","configuracion","",session("dsn_cliente"))
			'	end if
			'	rstSeries.close
			'end if
			rstPedidos.close
			if not isnull(rst("npedido_bis")) then 'si ya existe pedido bis para este albaran
				CrearDetalles lista,rst("nalbaran"),rst("npedido_bis"),"" 'se actualizan los detalles del pedido bis
				Precios rst("npedido_bis") 'Calculo de totales del pedido
				%><script language="javascript" type="text/javascript">
					if (window.confirm("<%=LitMsgActPedidoBis%><%=trimCodEmpresa(rst("npedido_bis"))%>. <%=LitMsgDeseaVer%>"))
						AbrirVentana('../search_layout.asp?pag1=compras/pedidos_pro.asp?ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido_bis")))%>?mode=browse?titulo=<%=LitPedidoBis%> <%=trimCodEmpresa(rst("npedido_bis"))%>&pag2=compras/pedidos_pro_bt.asp&pag3=compras/purchaseOrder_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
				</script><%
			else
				if nz_b(PreguntarSerie)<>0 then 'hay que consultar al usuario el nº serie del pedido bis
					%><script language="javascript" type="text/javascript">
					      //document.albaranes_prodet.action="albaranes_prodet.asp?nalbaran=<%=nalbaran%>&mode=pedirserie";
					      //document.albaranes_prodet.submit();
					</script><%
				else
					npedidoBIS=CabeceraPedido(SeriePedido,rst("fecha")) 'se genera la cabecera del pedido bis
					CrearDetalles lista,rst("nalbaran"),npedidoBIS,"" 'se actualizan los detalles del pedido bis
					rst("npedido_bis")=npedidoBIS 'se indica que el albarán generó un pedido bis
					rst.update
					Precios npedidoBIS 'Calculo de totales del pedido
					'A los pedidos implicados en la generacion del albarán se les indica que hay un pedido bis
					rstAux.open "update pedidos_pro set npedido_bis='" & npedidoBIS & "' where npedido in " & lista,session("dsn_cliente"),adOpenKeyset,adLockOptimistic%>
					<script language="javascript" type="text/javascript">
						if (window.confirm("<%=LitMsgGenPedidoBis%> <%=trimCodEmpresa(npedidoBIS)%>. <%=LitMsgDeseaVer%>"))
							AbrirVentana('../search_layout.asp?pag1=compras/pedidos_pro.asp?ndoc=<%=enc.EncodeForJavascript(null_s(rst("npedido_bis")))%>?mode=browse?titulo=<%=LitPedidoBis%> <%=trimCodEmpresa(rst("npedido_bis"))%>&pag2=compras/pedidos_pro_bt.asp&pag3=compras/purchaseOrder_lsearch.asp','P',<%=altoventana%>,<%=anchoventana%>);
					</script>
				<%end if
			end if
			rst.close
		end if

'''''''
		end if
	else
		rst_pedido.Close
	end if
end function

'*****************************************************************************
'Dibuja las diferentes capas para introducir los nserie de los pedidos que lo requieran
sub DibujarSpanSeries(fdesde,fhasta,nproveedor)
	if nproveedor > "" then
		Condiciones = " nproveedor='" & nproveedor & "' and fecha>='" & fdesde & "' and fecha<='" & fhasta & "' and nalbaran is null and nfactura is null"
	else
		Condiciones = " fecha>='" & fdesde & "' and fecha<='" & fhasta & "' and nalbaran is null and nfactura is null"
	end if

	seleccion="select * from detalles_ped_pro where npedido in "
	seleccion=seleccion & "(select npedido from pedidos_pro where " & Condiciones & ") and referencia in "
	seleccion=seleccion & "(select referencia from articulos where ctrl_nserie=1) order by npedido,item"
	rst.Open seleccion, session("dsn_cliente"),adOpenKeyset,adLockOptimistic

	npedidoAnt=""
	cont=0
	while not rst.eof
		nombreSpan="A" & Limpiar(rst("npedido")) & "B"
		nombreCampo=nombreSpan & completar(cstr(rst("item")),3,"0")
		if rst("npedido")<>npedidoAnt then
			if npedidoAnt<>"" then%>
				</table>
				</span>
				<br/>
			<%end if%>
			<span id="<%=nombreSpan%>" style="display:none">
				<table width='100%' border='0' cellspacing="1" cellpadding="1">
				    <%'Fila de encabezado
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
						DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rst("npedido"))
						DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("item")))
						DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rst("referencia"))
						DrawCelda2 "CELDA", "center", false, enc.EncodeForHtmlAttribute(null_s(rst("cantidad")))
						DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))
						pagina="../Mantenimiento/wizard_nserie.asp?mode=add&campo=" & nombreCampo & "&referencia=" & rst("referencia") & "&almacen=" & rst("almacen")
						%><input type="hidden" name="nserie_alb" value="<%=enc.EncodeForHtmlAttribute(nserie_alb)%>">
						<td class="CELDACENTER"><textarea class='CELDA' name="<%=enc.EncodeForHtmlAttribute(null_s(nombreCampo))%>" rows="2" cols="35"></textarea>
							<a class='CELDAREFB'  href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(null_s(pagina))%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)"><img src="../images/<%=ImgVarita%>" <%=ParamImgVarita%> alt="<%=LitAsistente%>" title="<%=LitAsistente%>"/></a></td><%
					CloseFila
		else
			'Escribir linea de detalle de nserie
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rst("npedido"))
				DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("item")))
				DrawCelda2 "CELDA", "left", false, trimCodEmpresa(rst("referencia"))
				DrawCelda2 "CELDA", "center", false, enc.EncodeForHtmlAttribute(null_s(rst("cantidad")))
				DrawCelda2 "CELDA", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))
				cadena=cadena & "C"
				pagina="../Mantenimiento/wizard_nserie.asp?mode=add&campo=" & nombreCampo & "&referencia=" & rst("referencia") & "&almacen=" & rst("almacen")
				%><td class="CELDACENTER"><textarea class='CELDA' name="<%=enc.EncodeForHtmlAttribute(null_s(nombreCampo))%>" rows="2" cols="35"></textarea>
					<a class='CELDAREFB'  href="javascript:AbrirVentana('<%=enc.EncodeForJavascript(null_s(pagina))%>','P',<%=AltoVentana%>,<%=AnchoVentana%>)"><img src="../images/<%=ImgVarita%>" <%=ParamImgVarita%> alt="<%=LitAsistente%>" title="<%=LitAsistente%>"/></a></td><%
			CloseFila
		end if
		npedidoAnt=rst("npedido")
	 	rst.movenext
		cont=cont+1
		if rst.eof then
			%></table>
			</span>
			<br/><%
		end if
	wend
	rst.close
end sub

function validar_pro_alb(nproveedor,nalbaran,falbaran)
	ModDocumento=true
	''ricardo 12/11/2003 comprobamos que no exista el nalbaran_pro para un mismo proveedor
	no_continuar=0
	strselect="select count(nalbaran) as contador from albaranes_pro where nproveedor='" & nproveedor & "' and nalbaran_pro='" & nalbaran & "' and year(fecha)= year(convert (datetime,'" & falbaran & "' )) "
	rstAux.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	if not rstAux.eof then
		cuantos_albaranes=null_z(rstAux("contador"))
		rstAux.close
		if cuantos_albaranes>0 then
			ModDocumento=false
		end if
	else
		rstAux.close
	end if
	if ModDocumento=true then validar_pro_alb="OK" else validar_pro_alb="NO_OK"
end function

'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************

borde=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion

	Dim ListaN,ListaNumeros,k

	set rst_albaran_det = Server.CreateObject("ADODB.Recordset")
	set rst_albaran = Server.CreateObject("ADODB.Recordset")
	set rst_albaran_con = Server.CreateObject("ADODB.Recordset")
	set rst_pedido_det = Server.CreateObject("ADODB.Recordset")
	set rst_pedido_con = Server.CreateObject("ADODB.Recordset")
	set rst_pedido = Server.CreateObject("ADODB.Recordset")
	set rstdomi = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstPedidos = Server.CreateObject("ADODB.Recordset")
	set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeout = 300
	conn.CommandTimeout = 300

 %>
<form name="pedpro_albpro_param" method="post">
<% PintarCabecera "pedpro_albpro_param.asp"

dim modd,modp,modi,modi2
'Leer parámetros de la página
	mode = Request.QueryString("mode")
	nproveedor	= limpiaCadena(Request.QueryString("nproveedor"))
	if nproveedor ="" then
		nproveedor	= limpiaCadena(Request.form("nproveedor"))
	end if
	if nproveedor > "" then
		nproveedor = session("ncliente") & completar(nproveedor,5,"0")
	end if
	'checkCadena(nproveedor)
	%><input type="hidden" name="nproveedor2" value="<%=enc.EncodeForHtmlAttribute(nproveedor)%>"><%

	fdesde		= limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde	= limpiaCadena(Request.form("fdesde"))
	end if
	fhasta		= limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta	= limpiaCadena(Request.form("fhasta"))
	end if

	nserie_alb	= limpiaCadena(Request.QueryString("nserie_alb"))
	if nserie_alb ="" then
		nserie_alb	= limpiaCadena(Request.form("nserie_alb"))
	end if
	fecha_alb	= limpiaCadena(Request.QueryString("fecha_alb"))
	if fecha_alb ="" then
		fecha_alb	= limpiaCadena(Request.form("fecha_alb"))
	end if
	nalbaran_pro_alb= limpiaCadena(Request.QueryString("nalbaran_pro_alb"))
	if nalbaran_pro_alb ="" then
		nalbaran_pro_alb	= limpiaCadena(Request.form("nalbaran_pro_alb"))
	end if
	viene=limpiaCadena(request.querystring("viene"))
	if viene="" then viene=limpiaCadena(request.form("viene"))
	if viene="" then viene="pedpro_albpro_param.asp"
	'if viene="cancelar" then p_npedido_pro=""

	genpedbis= limpiaCadena(Request.QueryString("genpedbis"))
	if genpedbis ="" then
		genpedbis	= limpiaCadena(Request.form("genpedbis"))
	end if
	
	existe = limpiaCadena(request.QueryString("existe"))
	nserie		= limpiaCadena(Request.form("nserie")&"")
	falbaran	= limpiaCadena(Request.form("falbaran")&"")
	h_nfilas	= limpiaCadena(Request.form("h_nfilas")&"")
	nalbaran_pro = limpiaCadena(request.form("nalbaran_pro")&"")
	fdesde2 = limpiaCadena(request.form("fdesde2")&"")
	fhasta2 = limpiaCadena(request.form("fhasta2")&"")
	nalbaran_pro2 = limpiaCadena(request.form("nalbaran_pro")&"")
	nproveedor2 = limpiaCadena(request.form("nproveedor2")&"")
	checkCadena(nproveedor2)

	dim cvsimp'',modi,modd,modp
	ObtenerParametros("convpedalbcomp")

	%><input type="hidden" name="fdesde2" value="<%=enc.EncodeForHtmlAttribute(fdesde)%>"/>
	<input type="hidden" name="fhasta2" value="<%=enc.EncodeForHtmlAttribute(fhasta)%>"/>
	<input type="hidden" name="generarpedidobis" value="0"/>
	<input type="hidden" name="viene" value="<%=enc.EncodeForHtmlAttribute(viene)%>"/>
	<input type="hidden" name="cvsimp" value="<%=enc.EncodeForHtmlAttribute(ucase(cvsimp))%>"/><%

	si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)

	if nproveedor<>"" then
		strselect="select fbaja from proveedores where nproveedor='" & nproveedor & "'"
		rstAux.open strselect, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if not rstAux.eof then
			if rstAux("fbaja")&"">"" then
				%><script language="javascript" type="text/javascript">				      window.alert("<%=LitProvDadoBajaConvPedAlb%>");</script><%
				nproveedor=""
			end if
		else
			%><script language="javascript" type="text/javascript">			      window.alert("<%=LitProvNoExiConvPedAlb%>");</script><%
			nproveedor=""
		end if
		rstAux.close
	end if

	strwhere=""
	strwhere3=""

	WaitBoxOculto LitEsperePorFavor
	Alarma "pedpro_albpro_param.asp"%>
	<hr/>
<%nalbaran_param = ""

	'Acción a realizar
	if mode="confirm" then
		mode = "select1"
		if validar_pro_alb(nproveedor2,nalbaran_pro2,falbaran)="NO_OK"  and false then
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitMsgNumeroAlbaranRepetido%>");
			      parent.pantalla.document.location = "pedpro_albpro_param.asp?mode=select2&nproveedor=<%=trimCodEmpresa(nproveedor2)%>&fdesde=<%=fdesde2%>&fhasta=<%=fhasta2%>";
			      parent.botones.document.location = "pedpro_albpro_param_bt.asp?mode=select2"
			</script><%
			mode=""
		else
			validar
			''ricardo 10-3-2003
			''ya no hace falta validar aqui,ya que se valido,cuando se introdujeron los nserie
			'if validarNSerie then
				if convertir(existe)="OK" then%>
					<script language="javascript" type="text/javascript">
					    //if (window.confirm("<%=LitMsgDeseaVer%>")==true){
					    //	parent.pantalla.document.location="albaranes_pro.asp?nalbaran=<%=nalbaran_param%>&mode=browse";
					    //	parent.botones.document.location="albaranes_pro_bt.asp?mode=browse";
					    //}
					</script>
				<%end if
			'end if
		end if
	end if

	if mode="select1"then%>
		<input type="hidden" name="nserie_alb" value="<%=enc.EncodeForHtmlAttribute(nserie_alb)%>">
		<input type="hidden" name="fecha_alb" value="<%=enc.EncodeForHtmlAttribute(fecha_alb)%>">
		<input type="hidden" name="nalbaran_pro_alb" value="<%=enc.EncodeForHtmlAttribute(nalbaran_pro_alb)%>">

        <%			
            'DrawCelda2 "CELDA", "left", false, LitDesdeFecha + ":"
		 	'DrawInputCelda "CELDA","","",10,0,"","fdesde",iif(fdesde>"",fdesde,"01/01/"+cstr(year(date)))
            EligeCelda "input","add","left","","",0,LitDesdeFecha,"fdesde",10,iif(fdesde>"",fdesde,"01/01/"+cstr(year(date)))
            DrawCalendar "fdesde"
			'DrawCelda2 "CELDA", "left", false, LitHastaFecha + ":"
		 	'DrawInputCelda "CELDA","","",10,0,"","fhasta",iif(fhasta>"",fhasta,date)
            EligeCelda "input","add","left","","",0,LitHastaFecha,"fhasta",10,iif(fhasta>"",fhasta,date)
            DrawCalendar "fhasta"
        %>
        <%        
			'DrawCelda2 "CELDA", "left", false, LitProveedor + ":"
			'DrawInputCeldaBuscar "CELDA","","",5,0,"","nproveedor",trimCodEmpresa(nproveedor),"AbrirVentana('proveedores_busqueda.asp?ndoc=pedpro_albpro_param&titulo=SELECCIONAR PROVEEDOR&mode=search&','P',"+cstr(altoventana)+","+cstr(anchoventana)+")",""
            DrawDiv "1","",""
            provSELECT = "select razon_social from proveedores with (nolock) where nproveedor=?"
            DrawLabel "","",LitProveedor%><input class='width15' type="text" name="nproveedor" value="<%=trimCodEmpresa(nproveedor)%>" onchange="TraerProveedor('<%=enc.EncodeForJavascript(mode)%>');"/><a class='CELDAREFB'  href="javascript:AbrirVentana('proveedores_busqueda.asp?ndoc=pedpro_albpro_param&titulo=<%=LitSelProveedor%>&mode=search','P',<%=altoventana%>,<%=anchoventana%>)"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt=""></a>
            <%DrawInput "width40","","razon_social",iif(nproveedor>"",DLookupP1(provSELECT,nproveedor&"",adchar,10,session("dsn_cliente")),""),"disabled"
                'd_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente")),"")
			'DrawInputCelda "CELDA disabled","","",40,0,"","razon_social",iif(nproveedor>"",d_lookup("razon_social","proveedores","nproveedor='" & nproveedor & "'",session("dsn_cliente")),"")
		%>

    <%elseif mode="select2" then%>
		<table width='100%' border="<%=borde%>" cellspacing="1" cellpadding="1"><%
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false, LitSerieAlbaran +":"
                rstAux.open "select nserie, case when datalength(substring(nserie,6,10)+' '+nombre)<=21 then substring(nserie,6,10)+'-'+nombre else left(substring(nserie,6,10)+'-'+nombre,20)+'...' end as descripcion from series where tipo_documento ='ALBARAN DE PROVEEDOR' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		        DrawSelectCelda "CELDA","","",0,"","nserie",rstAux,nserie,"nserie","descripcion","",""
		        rstAux.close

				DrawCelda2 "CELDA", "left", false, LitFechaAlbaran + ":"
				if fecha_alb>"" then
				 	DrawInputCelda "CELDA","","",10,0,"","falbaran",fecha_alb
				else
				 	DrawInputCelda "CELDA","","",10,0,"","falbaran",date
				end if
				DrawCelda2 "CELDA", "left", false, LitNAlbaran + ":"
			 	DrawInputCelda "CELDA","","",10,0,"","nalbaran_pro",enc.EncodeForHtmlAttribute(null_s(nalbaran_pro_alb))
			CloseFila%>
		</table>
		<br/>
        <%strwhere = CadenaBusqueda(fdesde,fhasta,nproveedor)

		''ricardo 7/3/2003
		''a partir de esta fecha, se mostrar en ventana aparte
		''''DibujarSpanSeries fdesde,fhasta,nproveedor
		'en su lugar tambien se creara un tabla aparte

		DropTable session("usuario"), session("dsn_cliente")
		crear="CREATE TABLE [" & session("usuario") & "] ("
''ricardo 31/7/2006 se cambia el tamaño de los campos mi_nserie y nserie de varchar a text
		crear=crear & "npedido varchar(20),item smallint,cantidad real,referencia varchar(30),mi_nserie text,nserie text, lote varchar(100),almacen varchar(10)"
		crear=crear & ")"
		rst.open crear,session("dsn_cliente"),adUseClient,adLockReadOnly

		GrantUser session("usuario"), session("dsn_cliente")

''''''''''''''''''''''''''''''''
        rst.cursorlocation=3
		rst.Open "select * from pedidos_pro with(NOLOCK) where " + strwhere, session("dsn_cliente")%>
<table width='100%' border='0' cellspacing="1" cellpadding="1">
        <%'Fila de encabezado
		DrawFila color_fondo%>
			<td class='CELDA' >
				<input type="checkbox" name="check" value="true" onclick="seleccionar();"/>
			</td>
			<%DrawCelda "ENCABEZADOL","","",0,LitPedido
			DrawCelda "ENCABEZADOL","","",0,LitCompletarTit
			DrawCelda "ENCABEZADOR","","",0,LitFecha
			DrawCelda "ENCABEZADOL","","",0,LitProveedor
			if cstr(ucase(cvsimp))<>"SI" then
				DrawCelda "ENCABEZADOR","","",0,LitTotal
			end if
		CloseFila

	VinculosPagina(MostrarPedidosPro)=1
	CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

		seleccion="select nseriealb from configuracion where nempresa='" & session("ncliente") & "'"
		rstAux.open seleccion, session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		if not rstAux.eof then
			if rstAux("nseriealb")<>0 then
				si_obligado_nserie=1
			else
				si_obligado_nserie=0
			end if
		else
			si_obligado_nserie=0
		end if
		rstAux.close%>
		
		<input type="hidden" name="si_obligado_nserie" value="<%=enc.EncodeForHtmlAttribute(si_obligado_nserie)%>"/>
		
		<%fila=1
		while not rst.EOF
			if isnull(rst("divisa")) then
				n_decimales=d_lookup("ndecimales","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente"))
			else
				n_decimales=d_lookup("ndecimales","divisas","codigo='" & rst("divisa") & "'",session("dsn_cliente"))
			end if

			'Seleccionar el color de la fila.
			if ((fila+1) mod 2)=0 then
				color=color_blau
			else
				color=color_terra
			end if

			'if si_obligado_nserie=1 then
				seleccion="select referencia from articulos with(NOLOCK) where ctrl_nserie=1 and referencia in (select referencia from detalles_ped_pro with(NOLOCK) where npedido='" & rst("npedido") & "')"
                rstAux.cursorlocation=3
				rstAux.open seleccion, session("dsn_cliente")
				if rstAux.eof then
					si_poner_nserie=0
				else
					si_poner_nserie=1
				end if
				rstAux.close
			'else
			'	si_poner_nserie=0
			'end if

            si_obligado_lotes=0
			seleccion="select referencia from articulos with(NOLOCK) where lotecompra=1 and referencia in (select referencia from detalles_ped_pro with(NOLOCK) where npedido='" & rst("npedido") & "')"
            rstAux.cursorlocation=3
			rstAux.open seleccion, session("dsn_cliente")
			if rstAux.eof then
				si_obligado_lotes=0
			else
				si_obligado_lotes=1
			end if
			rstAux.close

			if si_tiene_modulo_mantenimiento=0 then
				si_poner_nserie=0
				si_obligado_nserie=0
			end if

            ''if si_tiene_modulo_fabricacion=0 then
            ''    si_obligado_lotes=0
            ''end if

			DrawFila color%>
				<td class="CELDA">
                    <%
                        if si_obligado_nserie=1 and si_poner_nserie=1 then
                            hace_falta_serie=1
                        else
                            hace_falta_serie=0
                        end if
                        hace_falta_lote=si_obligado_lotes
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
					<input type="checkbox" name="check<%=fila%>" <%=bloquear_ck%> value="<%=rst("npedido")%>"/>
				</td>
				<td class="CELDALEFT" align="left">
					<%if cstr(ucase(cvsimp))<>"SI" then%>
						<%=Hiperv(OBJPedidosPro,rst("npedido"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("npedido")),LitVerPedido)%>
					<%else%>
						<%=trimCodEmpresa(rst("npedido"))%>
					<%end if%>
				</td>
				<%'if si_obligado_nserie=1 then
					if si_poner_nserie=0 then
						'DrawCelda "CELDA style='width:60px'","","",0,""
						cadena="A" & Limpiar(rst("npedido")) & "B"%>
						<td style='width:60px' align="center">
							<div class="CELDAB2" onmouseover="this.className='CELDAB2'" onmouseout="this.className='CELDAB2'" onclick="abrir_detalles('<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>','<%=fila%>');tier1Menu(img<%=fila%>)">
								<input type="hidden" name="hace_falta_nserie<%=fila%>" value="0">
              						<img src="../images/<%=ImgCarpetaCerrada%>" <%=ParamImgCarpetaCerrada%> id="img<%=fila%>" alt="<%=LitCompletarDet%>" title="<%=LitCompletarDet%>"/>
	              				</div>
						</td>
					<%else
						cadena="A" & Limpiar(rst("npedido")) & "B"%>
						<td style='width:60px' align="center">
							<div class="CELDAB2" onmouseover="this.className='CELDAB2'" onmouseout="this.className='CELDAB2'" onclick="abrir_detalles('<%=enc.EncodeForJavascript(null_s(rst("npedido")))%>','<%=fila%>','<%=enc.EncodeForJavascript(null_s(cadena))%>')">
								<%''ricardo 7-3-2003
								''a partir de esta fecha se mostrara en ventana aparte
								'''onClick="tier1Menu(<cadena>,img<fila>)">
								%>
								<input type="hidden" name="hace_falta_nserie<%=fila%>" value="1"/>
								<input type="hidden" name="nserie<%=fila%>" value=""/>
            	  					<img src="../images/<%=ImgNSerie_det%>" <%=ParamImgNserie_det%> id="img<%=fila%>" alt="<%=LitCompletarDetSerie%>" title="<%=LitCompletarDetSerie%>"/>
	              				</div>
						</td>
					<%end if
				'else
				'	DrawCelda "CELDA style='width:60px'","","",0,""
				'end if
				DrawCelda "CELDARIGHT","","",0,rst("fecha")
				DrawCelda "CELDA","","",0,d_lookup("razon_social","proveedores","nproveedor='" & rst("nproveedor") & "'",session("dsn_cliente"))
				if cstr(ucase(cvsimp))<>"SI" then
					DrawCelda "CELDARIGHT","","",0,formatnumber(rst("total_pedido"),n_decimales,-1,0,-1)
				end if
			CloseFila
			fila=fila+1
			rst.MoveNext
		wend%>
		<input type="hidden" name="h_nfilas" value="<%=rst.recordcount%>"/>
		<%rst.Close%>
		</table>
	<%end if%>
</form>
<%end if
connRound.close
set connRound = Nothing
set rst_albaran_det = Nothing
set rst_albaran = Nothing
set rst_albaran_con = Nothing
set rst_pedido_det = Nothing
set rst_pedido_con = Nothing
set rst_pedido = Nothing
set rstdomi = Nothing
set rstAux = Nothing
set rstSelect = Nothing
set rst = Nothing
set rstPedidos = Nothing
set conn = Nothing
%>
</body>
</html>