<%@ Language=VBScript %>
<% 
dim enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
%>
<%
' JCI 21/05/2003 : Hay que marcar/desmarcar la factura como cobrada en función de los cobros de vencimientos
'                  Pongo lo de la caché
''ricardo 5-6-2003 se pone el parametro caju para saber que cajas debo mostrar
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->

<!--#include file="cobros_param.inc" -->

<!--#include file="../perso.inc" -->

<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript" >
/**FLM:20090529: Suma el total de los documentos marcados.**/
var totalImporteCobrar=0.00;
var numDecimalesEmpresa='<%=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and Moneda_base<>0",session("dsn_cliente")) %>';

function sumTotREM(){
    var fila;
    totalImporteCobrar=0.00;
    for(fila=1;fila<=document.cobros_param.h_nfilas.value;fila++)
        if(document.cobros_param.elements["check"+fila].checked==true)
        {
            if (document.cobros_param.h_tabla.value=="vencimientos_salida" ||document.cobros_param.h_tabla.value=="tickets_cli")
            {
                //window.alert("los datos 1 son-" + document.cobros_param.elements["importecob"+fila].value + "-" + document.cobros_param.elements["factcambio"+fila].value + "-" + numDecimalesEmpresa + "-");
                totalImporteCobrar+=parseFloat(Redondear(parseFloat(document.cobros_param.elements["importecob"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
            }
            else
            {
                totalImporteCobrar+=parseFloat(Redondear(parseFloat(document.cobros_param.elements["h_deuda"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
            }
        }
       
    document.getElementById("totalACobrar").innerHTML=totalImporteCobrar.toFixed(numDecimalesEmpresa);
}

/*********************************************/
function cambiar_importecob(fila)
{
    //FLM:20090529:añado el else y el filtro de los tickets.
	if (document.cobros_param.h_tabla.value=="vencimientos_salida" )
	{
		importe=eval("document.cobros_param.h_importecob" + fila + ".value");
		importe=importe.replace(".","").replace(",",".");

		if (eval("document.cobros_param.check" + fila + ".checked==true")){
		    eval("document.cobros_param.importecob" + fila + ".value=document.cobros_param.h_deuda" + fila + ".value;");
		    totalImporteCobrar+=parseFloat(Redondear(parseFloat(document.cobros_param.elements["importecob"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
		}
		else{
		    totalImporteCobrar-=parseFloat(Redondear(parseFloat(document.cobros_param.elements["importecob"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
		    eval("document.cobros_param.importecob" + fila + ".value=document.cobros_param.cero" + fila + ".value;");
		}
	}
	else if (document.cobros_param.h_tabla.value=="tickets_cli"){
	    if (eval("document.cobros_param.check" + fila + ".checked==true"))
		    totalImporteCobrar+=parseFloat(Redondear(parseFloat(document.cobros_param.elements["importecob"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
		else
		    totalImporteCobrar-=parseFloat(Redondear(parseFloat(document.cobros_param.elements["importecob"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));
	}
	else{
	    if (eval("document.cobros_param.check" + fila + ".checked==true"))
		    totalImporteCobrar+=parseFloat(Redondear(parseFloat(document.cobros_param.elements["h_deuda"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));	
		else
		    totalImporteCobrar-=parseFloat(Redondear(parseFloat(document.cobros_param.elements["h_deuda"+fila].value.replace(",","."))*parseFloat(document.cobros_param.elements["factcambio"+fila].value.replace(",",".")),numDecimalesEmpresa).replace(".","").replace(",","."));	
	}
	document.getElementById("totalACobrar").innerHTML=totalImporteCobrar.toFixed(numDecimalesEmpresa);
}

function cambiar_importe(fila,deuda,importecob)
{
	if (document.cobros_param.h_tabla.value=="vencimientos_salida" ||document.cobros_param.h_tabla.value=="tickets_cli")
	{
		ok=1;
		importe_a_cobrar=eval("document.cobros_param.importecob" + fila + ".value");
		importe_a_cobrar=importe_a_cobrar.replace(',','.');
		if (isNaN(importe_a_cobrar)==true)
		{
			window.alert("<%=LitImporteCobNoNumero%>");
			ok=0;
		}

		if (importe_a_cobrar<=0)
		{
			window.alert("<%=LitImporteCobNoNegativo%>");
			ok=0;
		}
		if (ok==1 && importe_a_cobrar>deuda)
		{
			window.alert("<%=LitImportCobMayImpor%>");
			ok=0;
		}
		if (ok==1)
		{
			//eval("document.cobros_param.importecob" + fila + ".value=document.cobros_param.importecob" + fila + ".value.replace('.',',')");
			eval("document.cobros_param.check" + fila + ".checked=true");
		}
		else
		{
			eval("document.cobros_param.check" + fila + ".checked=false");
			eval("document.cobros_param.importecob" + fila + ".value=document.cobros_param.cero" + fila + ".value;");
		}		
	}
	//FLM:20090529: Suma el total de los documentos marcados.
	sumTotREM();
}

//Desencadena la búsqueda del cliente cuyo numero se indica
function TraerCliente(mode)
{
	if (document.all("radio_1").checked==true) documento_elegido=document.all("radio_1").value;
	if (document.all("radio_2").checked==true) documento_elegido=document.all("radio_2").value;
	if (document.all("radio_3").checked==true) documento_elegido=document.all("radio_3").value;
	
	document.location.href="cobros_param.asp?ncliente=" + document.cobros_param.ncliente.value + "&mode=" + mode
		+ "&Dfecha=" + document.cobros_param.Dfecha.value + "&Hfecha=" + document.cobros_param.Hfecha.value
		+ "&caju=" + document.cobros_param.caju.value + "&Documento=" + documento_elegido;
}

//**************************************************************************************************************
function seleccionar()
{
    //FLM:20090529: Suma el total de los documentos marcados.
    totalImporteCobrar=0.00;
    
	nregistros=document.cobros_param.h_nfilas.value;
	if (document.cobros_param.check.checked)
	{
		for (i=1;i<=nregistros;i++)
		{
			nombre="check" + i;
			document.cobros_param.elements[nombre].checked=true;
			cambiar_importecob(i);
		}
		document.cobros_param.check.value="yyy"
	}
	else
	{
		for (i=1;i<=nregistros;i++)
		{
			nombre="check" + i;
			document.cobros_param.elements[nombre].checked=false;
			cambiar_importecob(i);
		}
		//FLM:20090529:al desmarcar el check de todos el importe debe ser 0. Lo ponemos a 0 y actualizamos el total.
		totalImporteCobrar=0.00;
		document.getElementById("totalACobrar").innerHTML=totalImporteCobrar.toFixed(numDecimalesEmpresa);
		
		document.cobros_param.check.value="xxx"
	}
}

//****************************************************************************

function Cambiar(tipo)
{
	if (tipo=="facturas_cli" || tipo=="vencimientos_salida")
	{
		document.getElementById("venFact").style.display="";
		document.getElementById("tickets").style.display="none";
		document.getElementById("venFactCom").style.display="";
		document.getElementById("ticketsOp").style.display="none";
        document.getElementById("numDocumento").style.display="none"; 
	}
	else
	{
		if (tipo == "tickets_cli")
		{
			document.getElementById("venFact").style.display="none";
			document.getElementById("tickets").style.display="";
			document.getElementById("venFactCom").style.display="none";
			document.getElementById("ticketsOp").style.display="";
            var per = document.getElementById("peru").value;
            if( per == 0){
                document.getElementById("numDocumento").style.display=""; 
            }
            else{
                document.getElementById("numDocumento").style.display="none"; 
                }
         
		}
	}
}
</script>

<body onload="self.status='';" class="BODY_ASP">
<%
'Forma la cadena de búsqueda de la instrucción SQL para realizar las búsquedas
'de datos.
	'campo: Nombre del campo con el cual se realizará la búsqueda
	'criterio: Tipo de búsqueda
	'texto: Texto a buscar.
function CadenaBusqueda(campo,criterio,texto)
	strcond="and p.ncliente=c.ncliente"
	select case criterio
		case "contiene"
			CadenaBusqueda=campo + " like '%" + texto + "%'" + strcond + " order by " + campo
		case "empieza"
			CadenaBusqueda=campo + " like '" + texto + "%'" + strcond + " order by " + campo
		case "termina"
			CadenaBusqueda=campo + " like '%" + texto + "'" + strcond + " order by " + campo
		case "igual"
			CadenaBusqueda=campo + "='" + texto + "'" + strcond + " order by " + campo
	end select
end function
'******************************************************************************

'******************************************************************************
'Da por cobrados los vencimientos de la factura
sub CobrarVencimientos(nfactura)
	rstAux.Open "update vencimientos_salida with (updlock) set importecob=importe,cobrado=1 where nfactura like '"&session("ncliente")&"%' and nfactura='" & nfactura & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
end sub

'*****************************************************************************
'Anota en la caja caj con el medio pag el importe impcaja. Tipo "F"->factura. Tipo "V"->vencimiento. "T"->ticket.
sub AnotarEnCaja(tipo,ndoc,impcaja,div,caj,pag,rsoc,fechacobro, tienda)
	MB=d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0",session("dsn_cliente"))
	SigAnotacion=d_max("nanotacion","caja","caja='" & caj & "'",session("dsn_cliente")) + 1
	rstAux2.open "select * from caja where caja=''",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	rstAux2.addnew
	rstAux2("caja")=caj
	rstAux2("nanotacion")=SigAnotacion
	rstAux2("tanotacion")=iif(impcaja>=0,"ENTRADA","SALIDA")
	rstAux2("fecha")=fechacobro
    impcaja = replace(impcaja,".",",")
	rstAux2("importe")=iif(impcaja>=0,impcaja,-impcaja)
	rstAux2("medio")=pag
	if tabla="tickets_cli" then
		rstAux2("descripcion")="PAGO A CTA. TICKET TPV. OPERADOR: " & session("usuario")
	else
		rstAux2("descripcion")=rsoc & " (Desde documento)"
	end if
	rstAux2("ndocumento")=ndoc
    factcambio=0
	if tipo="F" then
		tip="FACTURA A CLIENTE"
        factcambio=d_lookup("factcambio","facturas_cli","nfactura like '"& session("ncliente") &"%' and nfactura='" & ndoc & "'",session("dsn_cliente"))
	elseif tipo="T" then
		tip="TICKET"
        factcambio=d_lookup("factcambio","divisas","codigo like '"& session("ncliente") &"%' and codigo='" & div & "'",session("dsn_cliente"))
	elseif tipo="V" then
		tip="VENCIMIENTO_SALIDA"
        pos = instr(1,ndoc,"-")
        factcambio=d_lookup("factcambio","facturas_cli","nfactura like '"& session("ncliente") &"%' and nfactura='" & mid(ndoc,1,pos-1) & "'",session("dsn_cliente"))
	end if
	rstAux2("tdocumento")=tip
	rstAux2("divisa")=div
    rstAux2("change_currency")=null_z(factcambio)
	rstAux2("tapunte")=session("ncliente") & "01"
	rstAux2("tienda")=nulear(tienda)
	rstAux2.update
	rstAux2.close
end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
borde=0

%>
<form name="cobros_param" method="post">
	<%PintarCabecera "cobros_param.asp"
	WaitBoxOculto LitEsperePorFavor
	'Leer parámetros de la página
  	mode=EncodeForHtml(Request.QueryString("mode"))
	if mode="search" then mode="browse"

	if request.querystring("caju")>"" then
		caju=limpiaCadena(request.querystring("caju"))
	else
		caju=limpiaCadena(request.form("caju"))
	end if

	tabla=limpiaCadena(request.form("h_tabla"))
	if tabla="" then tabla=limpiaCadena(request.form("documento"))
	Nregistros=limpiaCadena(request.form("h_nfilas"))
	caja=limpiaCadena(request.form("ncaja"))
	pago=limpiaCadena(request.form("i_pago"))
	fechacobro=limpiaCadena(request.form("fechacobro"))

	DFecha=limpiaCadena(Request.Form("Dfecha"))
	if DFecha & ""="" then DFecha=limpiaCadena(request.querystring("Dfecha"))
	HFecha=limpiaCadena(Request.Form("Hfecha"))
	if Hfecha & ""="" then Hfecha=limpiaCadena(request.querystring("Hfecha"))

	TmpDfecha=limpiaCadena(request.querystring("Dfecha"))
	TmpHfecha=limpiaCadena(request.querystring("Hfecha"))

	ncliente=limpiaCadena(Request.Form("ncliente"))
	if trim(ncliente) = "" then ncliente=limpiaCadena(request.querystring("ncliente"))
	ncliente=trim(ncliente)

	Serie=limpiaCadena(Request.Form("serie"))
	comercial=limpiaCadena(request.form("comercial"))
	Serie_tic=limpiaCadena(Request.Form("serie_tic"))
    NumDocumento=limpiaCadena(Request.Form("numDocumento"))
	operador=limpiaCadena(request.form("operador"))

	if request.form("opcclientebaja")>"" then
		opcclientebaja=limpiaCadena(request.form("opcclientebaja"))
	else
		opcclientebaja=limpiaCadena(request.querystring("opcclientebaja"))
	end if

	si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)
	si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)%>
	<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>">
	<%Alarma "cobros_param.asp"%>
	<hr/>
	<%set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstIvas = Server.CreateObject("ADODB.Recordset")
	set rstAlb = Server.CreateObject("ADODB.Recordset")
	set rstDetAlb = Server.CreateObject("ADODB.Recordset")
	set rstPed = Server.CreateObject("ADODB.Recordset")
    set rst = Server.CreateObject("ADODB.Recordset")
	set rstCofo = Server.CreateObject("ADODB.Recordset")
    Dim peru
    peru = 0
    if rstCofo.state<>0 then rstCofo.close
        rstCofo.open "select tankmargen from configuracion with (nolock) where ncliente like '"&session("ncliente")&"%' or almacen like '"&session("ncliente")&"%' ", session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        if not rstCofo.EOF then	
            peru = CInt(rstCofo("tankmargen"))
           %> <input type="hidden" id="peru" value="<%=EncodeForHtml(peru)%>">
           <%
        end if
	'Fragmento de codigo que realiza el cobro
	if mode="save" then
		DropTable session("usuario"), session("dsn_cliente")
		crear ="CREATE TABLE [" & session("usuario") & "] (ndocumento varchar(22),importecob money)"
		rst.open crear,session("dsn_cliente"),adUseClient,adLockReadOnly
		GrantUser session("usuario"), session("dsn_cliente")
		lista="("
		NregistrosSel=0
  		for i=1 to Nregistros
			nombre="check" & i
			if request.form(nombre) > "" then ' DOCUMENTO SELECCIONADO
				NregistrosSel=NregistrosSel+1
				ndocumento=trim(limpiaCadena(request.form(nombre)))
				lista=lista & "'" & ndocumento & "',"
				if tabla="vencimientos_salida" or tabla="tickets_cli" then
					strselect="insert into [" & session("usuario") & "] (ndocumento,importecob) values('" & ndocumento & "'," & replace(limpiaCadena(request.form("importecob" & i)),",",".") & ")"
					rst.open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					if rst.state<>0 then rst.close
				end if
			end if
		next
		lista=mid(lista,1,len(lista)-1) & ")" 'Quitamos la última coma y cerramos el paréntesis
		' Si se ha seleccionado caja, comprobamos que ninguno de los vencimientos esté cobrado parcialmente sin incluir en caja'
		' Si no se ha seleccionado caja, comprobamos que ninguno de los vencimientos esté cobrado parcialmente incluido en caja'
		if lista<>")" then
			if tabla<>"facturas_cli" then
				if  tabla="vencimientos_salida" then
					if caja>"" then
						strselect="select '' union all "
						strselect=strselect & "SELECT DISTINCT nrecibo FROM vencimientos_salida AS a with (nolock) LEFT OUTER JOIN "
						strselect=strselect & "caja AS caj with (nolock) ON a.nrecibo = caj.ndocumento and caj.caja like '"&session("ncliente")&"%' WHERE a.nfactura like '"&session("ncliente")&"%' and a.nrecibo IN " & lista & " AND "
						strselect=strselect & "(a.importecob = 0 OR (a.importecob <> 0 AND caj.importe IS NOT NULL))"
						rst.cursorlocation=3
						rst.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						if not rst.eof then
							if (rst.recordcount-1)<>NregistrosSel then
								lista="ERROR2"
							end if
						end if
					else
						strselect="SELECT * FROM caja with (nolock) WHERE caja like '"&session("ncliente")&"%' and ndocumento IN " & lista
						rst.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						if not rst.eof then
							lista="ERROR1"
						end if
					end if
					rst.close
				end if
			end if
		end if

		if lista=")" then%>
			<script language="javascript" type="text/javascript">
				window.alert("<%=LitMsgDocumentoNoSel%>");
				parent.botones.document.location = "cobros_param_bt.asp?mode=add";
                parent.pantalla.document.location="cobros_param.asp?mode=add";
				
			</script>
		<%elseif lista="ERROR1" then%>
			<script language="javascript" type="text/javascript">
				window.alert("<%=LitVencimEnCaja%>");
				parent.botones.document.location = "cobros_param_bt.asp?mode=add";
                document.location="cobros_param.asp?mode=add";
				
			</script>
		<%elseif lista="ERROR2" then%>
			<script language="javascript" type="text/javascript">
				window.alert("<%=LitVencimSinCaja%>");
				parent.botones.document.location = "cobros_param_bt.asp?mode=add";
                document.location="cobros_param.asp?mode=add";
				
			</script>
		<%else
			'Recorrido para el cobro de facturas por cliente'
			Dim ListaFacturas()

			if tabla="facturas_cli" then
				strselect="select nfactura,facturas_cli.deuda,cobrada,facturas_cli.divisa,rsocial, s.tienda from facturas_cli with (nolock) left outer join series s with(nolock) on s.nserie=facturas_cli.serie,clientes with (nolock) where nfactura like '"&session("ncliente")&"%' and clientes.ncliente like '"&session("ncliente")&"%' and nfactura IN " & lista & " and facturas_cli.ncliente=clientes.ncliente order by nfactura, fecha desc,facturas_cli.ncliente"
				rst.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				campo="nfactura"
			elseif tabla="vencimientos_salida" then
				strselect="select v.*,f.divisa,cl.rsocial,[" & session("usuario") & "].importecob as importe_a_cobrar, s.tienda from vencimientos_salida v with (nolock),facturas_cli f with (nolock) left outer join series s with(nolock) on s.nserie=f.serie,clientes cl with (nolock),[" & session("usuario") & "] where v.nfactura like '"&session("ncliente")&"%' and f.nfactura like '"&session("ncliente")&"%' and cl.ncliente like '"&session("ncliente")&"%' and nrecibo IN " & lista & " and f.nfactura=v.nfactura and f.ncliente=cl.ncliente and [" & session("usuario") & "].ndocumento=v.nrecibo order by nrecibo"
				rst.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				campo="nrecibo"
			elseif tabla="tickets_cli" then
				strselect="select t.*,cl.rsocial,[" & session("usuario") & "].importecob as importe_a_cobrar, null as tienda from tickets t with (nolock),clientes cl with (nolock),[" & session("usuario") & "] where t.nticket like '"&session("ncliente")&"%' and cl.ncliente like '"&session("ncliente")&"%' and t.nticket IN " & lista & " and t.ncliente=cl.ncliente and [" & session("usuario") & "].ndocumento=t.nticket order by t.nticket"
				rst.Open strselect,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				campo="nticket"
			end if
			ReDim ListaFacturas(rst.recordcount)
			k=1
			lista="("
			while not rst.EOF
				if tabla="facturas_cli" then
					pendiente=rst("deuda")

					''10/1/2003 - puesto por ricardo para cuando no habiendo detalles o conceptos, se borran todos, que tambien se
					''borren los vencimientos, si hay vencimiento automatico
					res_borr_dv=comprob_deuda_venci(rst("nfactura"),"VENTAS")

					rst.update
					
					'FLM:20090505:Comprobamos que no haya ninguna remesa con el vencimiento
		            rstAux.open "select top 1 r.nremesa from remesas r with(nolock) inner join detalles_remcli dr with(nolock) on dr.nremesa=r.nremesa and (dr.nfacturavto='" & rst("nfactura") & "' ) where r.nempresa='" & session("ncliente") & "' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
		            if rstAux.EOF then
		                venConRemesa=0
		            else
		                venConRemesa=1
		            end if
		            rstAux.close
		            
		            if caja>"" and venConRemesa=0 then AnotarEnCaja "F",rst("nfactura"),pendiente,rst("divisa"),caja,pago,rst("rsocial"),fechacobro, rst("tienda")
		            'FLM:20090505:Comprobamos que no haya ninguna remesa con el vencimiento para hacer la anotación en caja
					if caja>"" and venConRemesa=0 then 
					    'FLM:20090506:si incluyo en caja y no hay vencimientos en remesas se deben borrar los vencimientos.
					    rstAux.open "delete from vencimientos_salida with(rowlock) where importecob=0 and nfactura='" & rst("nfactura") & "' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					else
					    rstAux.open "update vencimientos_salida with(updlock) set importecob=importe,cobrado=1 where cobrado=0 and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					end if    

                    'rstAux.open "update vencimientos_salida with(updlock) set importecob=importe,cobrado=1 where nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					rstAux.open "update facturas_cli with (updlock) set cobrada=1 where nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

					'FLM:20090505:No se deben borrar los vencimientos, el comportamiento debe ser igual que al marcar a cobrada una factura desde la propia factura.					
					'''CobrarVencimientos rst(campo)
					''Como la factura se dio por cobrada, hay que eliminar los vencimientos que no estén cobrados y cobrar los que tengan un cobro parcial
					'FLM:20090424:Comprobamos que no haya ninguna remesa con el vencimiento
					 'if venConRemesa=0 then
					 '   rstAux.open "delete from vencimientos_salida where importecob=0 and nfactura='" & rst("nfactura") & "' ",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					'else
					    'rstAux.open "update vencimientos_salida with(updlock) set importecob=importe,cobrado=1 where cobrado=0 and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
                    'end if
                    'rstAux.open "update vencimientos_salida with(updlock) set importe=importecob,cobrado=1 where importecob<>0 and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic

					lista=lista & "'" & trimCodEmpresa(rst("nfactura")) & "',"
				elseif tabla ="vencimientos_salida" then
					rstAux.open "select * from vencimientos_salida where nfactura like '"&session("ncliente")&"%' and nrecibo='" & rst("nrecibo") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					if not rstAux.eof then
						importe_cob_ant=rst("importecob")
						pendiente=rst("importe_a_cobrar")
						importecob_a_poner=rst("importe_a_cobrar") + rst("importecob")
						rstAux("importecob")=importecob_a_poner
						if rst("importe")=rstAux("importecob") then
							rstAux("cobrado")=1
						end if
						rstSelect.cursorlocation=3
						rstSelect.open "select s.tienda from facturas_cli f with(nolock) inner  join series s with(nolock) on s.nserie=f.serie where f.nfactura='"&rstAux("nfactura")&"'", session("dsn_cliente"), adOpneKeyse, adLockOptimistic
                        if not rstSelect.eof then
                            tienda=rstSelect("tienda")
                        end if
                        rstSelect.close
						if caja>"" then AnotarEnCaja "V",rst("nrecibo"),pendiente,rst("divisa"),caja,pago,rst("rsocial"),fechacobro, tienda
						rstAux.update
					end if
					rstAux.close
					DeudaFactura=d_lookup("deuda","facturas_cli","nfactura='" & rst("nfactura") & "'",session("dsn_cliente"))
					rstAux.open "select count(*) as total from vencimientos_salida with (nolock) where nfactura like '"&session("ncliente")&"%' and cobrado=0 and nfactura='" & rst("nfactura") & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					SinCobrar=rstAux("total")
					rstAux.close
					if (DeudaFactura<=0) and (SinCobrar=0) then
						lista=lista & "'" & trimCodEmpresa(rst("nfactura")) & "',"
					end if
				elseif tabla ="tickets_cli" then
					rstAux.open "select importecob from ["&session("usuario")&"] where ndocumento='"&rst("nticket")&"'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					if not rstAux.eof then importeacobrar=rstAux("importecob")
					rstAux.close
					AnotarEnCaja "T",rst("nticket"),replace(importeacobrar,",","."),rst("divisa"),caja,pago,rst("rsocial"),fechacobro, null
					lista=lista & "'" & trimCodEmpresa(rst("nticket")) & "',"
				end if
				k=k+1
				rst.MoveNext
			wend
			lista=mid(lista,1,len(lista)-1) & ")" 'Quitamos la última coma y cerramos el paréntesis
			if lista=")" then%>
				<script>
					<%if tabla <> "tickets_cli" then%>
					window.alert("<%=LitMsgFacturasNoCobradas%>");
					<%else%>
					window.alert("<%=LitMsgTicketsNoCobradas%>");
					<%end if%>
					caju=document.cobros_param.caju.value;
					parent.botones.document.location="cobros_param_bt.asp?mode=add";
                    document.location="cobros_param.asp?mode=add&caju=" + caju;
					
				</script>
			<%else%>
				<script>
					<%if tabla <> "tickets_cli" then%>
					    window.alert("<%=LitMsgFacturasCobradas%><%=Lista%>");
					    <%else%>
                            <%if peru = 0 then%>
					            window.alert("<%=fechacobro%>\n<%=LitMsgTicketsCobradas%><%=Lista%>");
                            <%else%>
                                window.alert("<%=LitMsgTicketsCobradas%><%=Lista%>");
                            <%end if %>
                    <%end if%>
					caju=document.cobros_param.caju.value;
					parent.botones.document.location="cobros_param_bt.asp?mode=add";
                    document.location="cobros_param.asp?mode=add&caju=" + caju;
					
				</script>
			<%end if
			rst.close
		end if 'lista=)
	end if 'mode=save
'****************************************************************************************************************

	'Comprobamos la validez de datos.
	if mode="browse" then
		if not isdate(DFecha) or not isdate(Hfecha) then
			if not isdate(DFecha) then%>
				<script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgDesdeFechaFecha%>");
					parent.botones.document.location="cobros_param_bt.asp?mode=add";
				</script>
			<%else%>
				<script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgHastaFechaFecha%>");
					parent.botones.document.location="cobros_param_bt.asp?mode=add";
				</script>
			<%end if
			mode="add"
		else
			ncliente=trim(ncliente)
			if ncliente > "" then
				rstAux.open "select ncliente from clientes with (nolock) where ncliente like '"&session("ncliente")&"%' and ncliente='" & session("ncliente") & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if rstAux.EOF then%>
					<script language="javascript" type="text/javascript">
						window.alert("<%=LitMsgClienteNoExiste%>");
						parent.botones.document.location="cobros_param_bt.asp?mode=add";
					</script>
					<%mode="add"
				end if
				rstaux.close
			end if
		end if
	end if

	if mode="add" then
		'FRAGMENTO DE CODIGO ENCARGADO DE CAPTURAR EL NOMBRE DEL CLIENTE
		TmpNCliente=""
		TmpNombre=""
		TmpDfecha=""
		TmpHfecha=""
		if trim(ncliente) > "" then
			ncliente=session("ncliente") & completar(trim(ncliente),5,"0")
			rstAux.open "select rsocial from clientes with (nolock) where ncliente like '"&session("ncliente")&"%' and ncliente='" & ncliente & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			if rstAux.EOF then%>
				<script language="javascript" type="text/javascript">
					window.alert("<%=LitMsgClienteNoExiste%>");
				</script>
			<%else
				TmpNCliente=ncliente
				TmpNombre=rstAux("rsocial")%>
				<script language="javascript" type="text/javascript">
					parent.botones.document.anchors("Seleccionar").focus();
				</script>
			<%end if
			rstAux.close
		end if%>
		<table width="100%" border='0' cellspacing="1" cellpadding="1">
			<tr><%
                DrawDiv "6","",""
                DrawLabel "","",LitFacturas%><input type="radio" id="radio_1" name="Documento" value="facturas_cli" <%=iif(limpiaCadena(request.querystring("Documento"))="facturas_cli","checked","")%> onclick="javascript:Cambiar('facturas_cli');"><%CloseDiv
                DrawDiv "6","",""
                DrawLabel "","",LitVencimientos%><input type="radio" id="radio_2" name="Documento" value="vencimientos_salida" <%=iif(limpiaCadena(request.querystring("Documento"))="vencimientos_salida" or request.querystring("Documento")="","checked","")%> onclick="javascript:Cambiar('vencimientos_salida');"><%CloseDiv
                DrawDiv "6","",""
                DrawLabel "","",LitTickets%><input type="radio" id="radio_3" name="Documento" value="tickets_cli" <%=iif(limpiaCadena(request.querystring("Documento"))="tickets_cli","checked","")%> onclick="javascript:Cambiar('tickets_cli');"><%CloseDiv
                %></tr>
            <tr><%
				'DrawCelda2 "CELDA style='width:120px'", "left", false, LitDesdeFecha +":"
			 	'DrawInputCelda "CELDA","","",10,0,"","Dfecha",iif(TmpDfecha>"",TmpDfecha,"01/01/" & year(date))
                EligeCelda "input",mode,"left","","",0,LitDesdeFecha,"Dfecha",10,EncodeForHtml(iif(TmpDfecha>"",TmpDfecha,"01/01/" & year(date)))
                DrawCalendar "Dfecha"
				diaHoy = day(date)
				mesHoy=month(date)
				fechaHoy=iif(Len(diaHoy)>1,diaHoy,"0"&diaHoy)&"/"&iif(Len(mesHoy)>1,mesHoy,"0"&mesHoy)&"/"&year(date)
				'DrawCelda2 "CELDA style='width:120px'", "left", false, LitHastaFecha +":"
			 	'DrawInputCelda "CELDA","","",10,0,"","Hfecha",iif(TmpHfecha>"",TmpHfecha,fechaHoy)
                EligeCelda "input",mode,"left","","",0,LitHastaFecha,"Hfecha",10,EncodeForHtml(iif(TmpHfecha>"",TmpHfecha,fechaHoy))
                DrawCalendar "Hfecha"
				'DrawCelda2 "CELDA", "left", false,"&nbsp;"
				'DrawCelda2 "CELDA", "left", false,"&nbsp;"
				'DrawCelda2 "CELDA style='width:120px'", "left", false, LitCodigo +":"
				Formulario="cobros_param"
                DrawDiv "1","",""
                DrawLabel "","",LitCodigo
                %><input type="text" class="width15" maxlength="5" name="ncliente" value="<%=EncodeForHtml(trimCodEmpresa(TmpNcliente))%>" onchange="TraerCliente('<%=EncodeForHtml(mode)%>')"><a class='CELDAREFB' href="javascript:AbrirVentana('../ventas/clientes_buscar.asp?ndoc=<%=Formulario%>&titulo=<%=LitSelCliente%>&mode=search&viene=cobros_param','P',400,700)" OnMouseOver="self.status='<%=LitVerCliente%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input type="text" class="width40" disabled name="nombre" value="<%=EncodeForHtml(TmpNombre)%>"><%CloseDiv
               	%><span ID="venFact" style="display:"><%
					'DrawCelda2 "CELDA style='width:120px'", "left", false,LitSerieFactura +": "
					rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & session("ncliente") & "%' and tipo_documento ='FACTURA A CLIENTE'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					'DrawSelectCelda "CELDA","175","",0,"","serie",rstAux,iif(nserie>"" and tabla="facturas_cli",nserie,""),"nserie","descripcion","",""
                    DrawSelectCelda "CELDA","175","",0,LitSerieFactura,"serie",rstAux,iif(nserie>"" and tabla="facturas_cli",nserie,""),"nserie","descripcion","",""
					rstAux.close
				%></span><span ID="tickets" style="display:none"><%
					'DrawCelda2 "CELDA style='width:120px'", "left", false,LitSerieTicket +": "
					rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(nolock) where nserie like '" & session("ncliente") & "%' and tipo_documento ='TICKET'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					'DrawSelectCelda "CELDA","175","",0,"","serie_tic",rstAux,iif(nserie>"" and tabla="tickets_cli",nserie,""),"nserie","descripcion","",""
                    DrawSelectCelda "CELDA","175","",0,LitSerieTicket,"serie_tic",rstAux,iif(nserie>"" and tabla="tickets_cli",nserie,""),"nserie","descripcion","",""
					rstAux.close
				%></span><%
                ''Campo nuevo 
             %><SPAN ID="numDocumento" style="display:none"><%
                 'DrawCelda2 "CELDA style='width:120px'", "left", false,LitNumDocumento +": "
				  'DrawInputCeldaLen "CELDA style='width:175px'","","",22,0,"","numDocumento","",22
                 DrawDiv "1","",""
                 DrawLabel "","",LitNumDocumento%><input type="text" name="numDocumento" maxlength="22" size="22" /><%CloseDiv
              %></span><%
			DrawDiv "1","",""
            DrawLabel "","",LitClienteBaja%><input type="checkbox" class="CELDA" name="opcclientebaja"><%
            CloseDiv
            %><SPAN ID="venFactCom" style="display:"><%
                Dim Literal
				if si_tiene_modulo_comercial<>0 then
					'DrawCelda2 "CELDA style='width:120px'", "left", false, LitComercialModCom +":"
                    Literal = LitComercialModCom
				else
					'DrawCelda2 "CELDA style='width:120px'", "left", false, LitComercial +":"
                    Literal = LitComercial
				end if
				rstAux.open "select dni, nombre from personal with (nolock),comerciales with(nolock) where comerciales.fbaja is null and dni=comercial and dni like '" & session("ncliente") & "%' and comercial like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				'DrawSelectCelda "CELDA","160","",0,"","comercial",rstAux,"","dni","nombre","",""
                DrawSelectCelda "CELDA","160","",0,Literal,"comercial",rstAux,"","dni","nombre","",""
				rstAux.close
				'DrawCelda2 "CELDA", "left", false,"&nbsp;"
				'DrawCelda2 "CELDA", "left", false,"&nbsp;"
			%></span><SPAN ID="ticketsOp" style="display:none"><%
				'DrawCelda2 "CELDA style='width:120px'", "left", false,LitOperador +": "
				rstAux.open "select dni, nombre from personal with(nolock) where dni like '" & session("ncliente") & "%'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				'DrawSelectCelda "CELDA","175","",0,"","operador",rstAux,"","dni","nombre","",""
                DrawSelectCelda "CELDA","175","",0,LitOperador,"operador",rstAux,"","dni","nombre","",""
				rstAux.close
				'DrawCelda2 "CELDA", "left", false,"&nbsp;"
				'DrawCelda2 "CELDA", "left", false,"&nbsp;"
			%></span><%CloseFila%>
        </tr>
		</table><hr/>
		<%'Mostrar los documentos.
	elseif mode="browse" then
		MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='286'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='286'", DSNIlion)%>
		<input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>
		<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'>
		<input type="hidden" name="Documento" value="<%=EncodeForHtml(tabla)%>">
		<input type="hidden" name="ncliente" value="<%=EncodeForHtml(ncliente)%>">
		<input type="hidden" name="serie" value="<%=EncodeForHtml(Serie)%>">
		<input type="hidden" name="comercial" value="<%=EncodeForHtml(comercial)%>">
		<input type="hidden" name="serie_tic" value="<%=EncodeForHtml(Serie_tic)%>">
        <input type="hidden" name="numDocumento" value="<%=EncodeForHtml(NumDocumento)%>">
		<input type="hidden" name="operador" value="<%=EncodeForHtml(operador)%>">
		<input type="hidden" name="opcclientebaja" value="<%=EncodeForHtml(opcclientebaja)%>">
		<input type="hidden" name="Dfecha" value="<%=EncodeForHtml(DFecha)%>">
		<input type="hidden" name="Hfecha" value="<%=EncodeForHtml(HFecha)%>">
		<%if ncliente > "" then
			ncliente=session("ncliente") & ncliente
			if tabla="facturas_cli" then 'FACTURAS
				seleccion = "select nfactura,fecha,f.ncliente,cl.rsocial,total_factura,f.deuda,f.divisa,d.ndecimales,d.abreviatura,(select count(*) from vencimientos_salida with (nolock) where vencimientos_salida.nfactura like '"&session("ncliente")&"%' and vencimientos_salida.nfactura=f.nfactura) as NumeroVencimientos,d.factcambio from facturas_cli f with (nolock) left outer join comerciales as co with (nolock) on f.comercial=co.comercial and co.comercial like '"&session("ncliente")&"%', clientes cl with (nolock), divisas d with (nolock) where f.nfactura like '"&session("ncliente")&"%' and cl.ncliente like '"&session("ncliente")&"%' and d.codigo like '"&session("ncliente")&"%' and f.ncliente='" & ncliente & "' and fecha>='" & DFecha & "' and fecha<='" & HFecha & "' and cobrada=0 and f.deuda<>0 and f.ncliente=cl.ncliente and f.divisa=d.codigo "
				if Serie>"" then
				   seleccion = seleccion + " and serie='" + Serie + "'"
				end if
				if comercial & "">"" then
					seleccion = seleccion + " and f.comercial='" + comercial + "'"
				'FLM:20090702:por JCI, no se debe hacer este filtro.
				'else
				'	seleccion = seleccion + " and co.fbaja is null "
				end if
				orden=" order by fecha,nfactura"
				'FLM:20090511:Filtro que no se listen facturas cuyos vencimientos estén incluidos en alguna remesa.
				where1 = " and (exists (select * from vencimientos_salida with (nolock) left join detalles_remcli dr with(nolock) on dr.nfacturavto=f.nfactura  where vencimientos_salida.nfactura like '"&session("ncliente")&"%' and vencimientos_salida.nfactura=f.nfactura and importecob=0 and dr.nfacturavto is null) OR (SELECT COUNT(*) FROM vencimientos_salida with (nolock) WHERE nfactura like '"&session("ncliente")&"%' and nfactura = f.nfactura) = 0) "
				''seleccion = seleccion + orden
				seleccion = seleccion + where1 + orden		
			elseif tabla="vencimientos_salida" then 'VENCIMIENTOS
				seleccion = "select nrecibo,f.nfactura,f.divisa,fechav as fecha,f.ncliente,cl.rsocial,importe,importe-importecob as deuda,importecob,d.ndecimales,d.abreviatura,d.factcambio from facturas_cli f with (nolock) left outer join comerciales as co with (nolock) on f.comercial=co.comercial and co.comercial like '"&session("ncliente")&"%',vencimientos_salida with (nolock),divisas d with (nolock),clientes cl with (nolock) "
				seleccion = seleccion & "where vencimientos_salida.nfactura like '"&session("ncliente")&"%' and cl.ncliente like '"&session("ncliente")&"%' and d.codigo like '"&session("ncliente")&"%' and f.ncliente='" & ncliente & "' and f.nfactura=vencimientos_salida.nfactura and fechav>='" & DFecha & "' and fechav<='" & HFecha & "' and cobrado=0 and importe>importecob and f.ncliente=cl.ncliente and f.divisa=d.codigo and d.codigo like '" & session("ncliente") & "%'"
				if Serie>"" then
				   seleccion = seleccion + " and f.serie='" + Serie + "'"
				end if
				if comercial & "">"" then
					seleccion = seleccion + " and vencimientos_salida.comercial='" + comercial + "'"
				'FLM:20090702:por JCI, no se debe hacer este filtro.
				'else
				'	seleccion = seleccion + " and co.fbaja is null "
				end if
				orden=" order by fecha,nrecibo"
				seleccion = seleccion + orden
			elseif tabla="tickets_cli" then
				seleccion = "select * from (select t.nticket, t.nventa, t.divisa, t.fecha, t.ncliente, cl.rsocial, t.total_ticket as importe,isnull(pt.totalPagado,0) as cobrado, t.total_ticket-isnull(pt.totalPagado,0) as deuda,d.ndecimales,d.abreviatura,d.factcambio from tickets t with(nolock) left outer join personal p with(nolock) on p.dni= t.usuario and p.dni like '"&session("ncliente")&"%' inner join divisas d with(nolock) on d.codigo=t.divisa and d.codigo like '"&session("ncliente")&"%' inner join clientes cl with(nolock) on cl.ncliente=t.ncliente and cl.ncliente like '"&session("ncliente")&"%' left outer join (select sum (case when tanotacion='ENTRADA' then importe else -importe end) as totalPagado, ndocumento from caja with(nolock) where ndocumento like '"&session("ncliente")&"%' and tdocumento='TICKET' "
				seleccion = seleccion & " group by ndocumento) pt on pt.ndocumento=t.nticket where t.nticket like '"&session("ncliente")&"%' and cl.ncliente like '"&session("ncliente")&"%' and t.fecha>='" & DFecha & " 00:00:00.000' and t.fecha<='" & HFecha & " 23:59:00.000' and t.total_ticket<>0 "
				seleccion1 = seleccion1 & " ) tmp where tmp.importe<>tmp.cobrado order by tmp.fecha,tmp.nventa"
				if Serie_tic>"" then
				   where = where + " and t.serie='" + Serie_tic + "'"
				end if
                if NumDocumento >"" then
				   where = where + " and t.nventa='" + session("ncliente") + NumDocumento + "'"
				end if
				if ncliente & "">"" then
					where = where + " and t.ncliente='" + ncliente + "'"
				end if
				if operador & "" > "" then
				    where = where + " and t.usuario='" + operador + "'"
				end if
				
				seleccion = seleccion + where + seleccion1
			end if
		else
			if request.form("opcclientebaja")>"" then
				opcclientebaja=request.form("opcclientebaja")
			else
				opcclientebaja=limpiaCadena(request.querystring("opcclientebaja"))
			end if

			if opcclientebaja="" then
				strbaja=" "
			else
				strbaja=" and cl.fbaja is null"
			end if

			if tabla="facturas_cli" then 'FACTURAS
				seleccion = "select nfactura,fecha,p.ncliente,cl.rsocial,total_factura,p.deuda,p.divisa,d.ndecimales,d.abreviatura,(select count(*) from vencimientos_salida with (nolock) where vencimientos_salida.nfactura like '"&session("ncliente")&"%' and vencimientos_salida.nfactura=p.nfactura) as NumeroVencimientos,d.factcambio from facturas_cli as p with (nolock) left outer join comerciales as co with (nolock) on p.comercial=co.comercial and co.comercial like '"&session("ncliente")&"%',clientes as cl with (nolock),divisas as d with (nolock) where p.nfactura like '"&session("ncliente")&"%' and cl.ncliente like '"&session("ncliente")&"%' and d.codigo like '"&session("ncliente")&"%' and fecha>='" & DFecha & "' and fecha<='" & HFecha & "' and cobrada=0 and p.divisa=d.codigo "
				'FLM:20090511:Filtro que no se listen facturas cuyos vencimientos estén incluidos en alguna remesa.
				where1 = " and (exists (select * from vencimientos_salida with (nolock) left join detalles_remcli dr with(nolock) on dr.nfacturavto=p.nfactura  where vencimientos_salida.nfactura like '"&session("ncliente")&"%' and vencimientos_salida.nfactura=p.nfactura and importecob=0 and dr.nfacturavto is null) OR (SELECT COUNT(*) FROM vencimientos_salida with (nolock) WHERE vencimientos_salida.nfactura like '"&session("ncliente")&"%' and nfactura = p.nfactura) = 0) "
				orden = " order by fecha,nfactura, p.ncliente"
				if Serie>"" then
				   seleccion = seleccion + " and serie='" + Serie + "'"
				end if
				if comercial & "">"" then
					seleccion = seleccion + " and p.comercial='" + comercial + "'"
				'FLM:20090702:por JCI, no se debe hacer este filtro.
				'else
				'	seleccion = seleccion + " and co.fbaja is null"
				end if
				''seleccion = seleccion + strbaja + " and p.ncliente=cl.ncliente" + orden
				seleccion = seleccion + strbaja + " and p.ncliente=cl.ncliente" + where1 + orden
			elseif tabla="vencimientos_salida" then 'VENCIMIENTOS
				seleccion = "select nrecibo,p.nfactura,p.divisa,fechav as fecha,p.ncliente,cl.rsocial,importe,importe-importecob as deuda,importecob,d.ndecimales,d.abreviatura,d.factcambio from facturas_cli as p with (nolock) left outer join comerciales as co with (nolock) on p.comercial=co.comercial and co.comercial like '"&session("ncliente")&"%',clientes as cl with (nolock),vencimientos_salida with (nolock),divisas as d with (nolock) "
				seleccion = seleccion & "where p.nfactura like '" & session("ncliente") & "%' and cl.ncliente like '"&session("ncliente")&"%' and d.codigo like '"&session("ncliente")&"%' and p.nfactura=vencimientos_salida.nfactura and fechav>='" & DFecha & "' and fechav<='" & HFecha & "' and cobrado=0 and importe>importecob and p.divisa=d.codigo and d.codigo like '" & session("ncliente") & "%' "
				orden = " order by fecha,nrecibo"
				if Serie>"" then
				   seleccion = seleccion + " and p.serie='" + serie + "'"
				end if
				if comercial & "">"" then
					seleccion = seleccion + " and vencimientos_salida.comercial='" + comercial + "'"
				'FLM:20090702:por JCI, no se debe hacer este filtro.
				'else
				'	seleccion = seleccion + " and co.fbaja is null"
				end if
				seleccion = seleccion + strbaja + " and p.ncliente=cl.ncliente" + orden
			elseif tabla="tickets_cli" then
				seleccion = "select * from (select t.nticket, t.nventa, t.divisa, t.fecha, t.ncliente, cl.rsocial, t.total_ticket as importe,isnull(pt.totalPagadoEntrada,0)-isnull(pt1.totalPagadoSalida,0) as cobrado, t.total_ticket-isnull(pt.totalPagadoEntrada,0)+isnull(pt1.totalPagadoSalida,0) as deuda,d.ndecimales,d.abreviatura,d.factcambio from tickets t with(nolock) left outer join personal p with(nolock) on p.dni= t.usuario and p.dni like '"&session("ncliente")&"%' inner join divisas d with(nolock) on d.codigo=t.divisa and d.codigo like '"&session("ncliente")&"%' inner join clientes cl with(nolock) on cl.ncliente=t.ncliente and cl.ncliente like '"&session("ncliente")&"%' left outer join (select sum (importe) as totalPagadoEntrada, ndocumento from caja with(nolock) where ndocumento like '"&session("ncliente")&"%' and tdocumento='TICKET' and tanotacion='entrada' "
				seleccion = seleccion & " group by ndocumento) pt on pt.ndocumento=t.nticket left outer join (select sum(importe) as totalPagadoSalida, ndocumento from caja with(nolock) where ndocumento like '"&session("ncliente")&"%' and tdocumento='TICKET' and tanotacion='SALIDA' group by ndocumento) pt1 on pt1.ndocumento=t.nticket where t.nticket like '"&session("ncliente")&"%' and cl.ncliente like '"&session("ncliente")&"%' and t.fecha>='" & DFecha & " 00:00:00.000' and t.fecha<='" & HFecha & " 23:59:00.000' and t.total_ticket<>0 and t.nfactura is null "
				seleccion1 = seleccion1 & " ) tmp where tmp.importe<>tmp.cobrado order by tmp.fecha,tmp.nventa"
				if Serie_tic>"" then
				   where = where + " and t.serie='" + Serie_tic + "'"
				end if
                if NumDocumento >"" then
				   where = where + " and t.nventa='" + session("ncliente") + NumDocumento + "'"
				end if
				if ncliente & "">"" then
					where = where + " and t.ncliente='" + ncliente + "'"
				end if
				if operador & "" > "" then
				    where = where + " and t.usuario='" + operador + "'"
				end if
				seleccion = seleccion + where +  strbaja + seleccion1
			end if
		end if

		rst.cursorlocation=3
		rst.Open seleccion,session("dsn_cliente")
		if not rst.EOF then 'HAY REGISTROS
		   	'Calculos de páginas--------------------------
			NumRegTotal=rst.recordcount
		  	lote=limpiaCadena(Request.QueryString("lote"))
			if lote="" then
				lote=1
			end if
			sentido=limpiaCadena(Request.QueryString("sentido"))
			lotes=NumRegTotal/MAXPAGINA
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
			rst.AbsolutePage=lote%>
			<table border='0' cellspacing="1" cellpadding="1">
				<%DrawFila color_blau
					DrawCelda2 "CELDA", "left", true, LitCaja
					defecto=" "
					poner_cajas "CELDA7",defecto,"ncaja","","codigo","descripcion","","",poner_comillas(caju)
					DrawCelda2 "CELDA", "left", true, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					DrawCelda2 "CELDA", "left", true, LitPago
					''ricardo 28-7-2008 si tiene el modulo ebesa solamente saldran los tipos en donde copiasticket=2
					if si_tiene_modulo_ebesa<>0 then
					    StrSelPagAdd="select t.codigo,t.descripcion "
                        StrSelPagAdd=StrSelPagAdd & " from tipo_pago as t with(NOLOCK) "
                        StrSelPagAdd=StrSelPagAdd & " where t.codigo like '" & session("ncliente") & "%' "
                        StrSelPagAdd=StrSelPagAdd & " and t.copiasticket=2 "
                        StrSelPagAdd=StrSelPagAdd & " order by t.descripcion"
					else
					    StrSelPagAdd="SELECT * FROM Tipo_pago with (nolock) where codigo like '" & session("ncliente") & "%' order by descripcion"
					end if
					rstAux.cursorlocation=3
					rstAux.Open StrSelPagAdd,session("dsn_cliente")
					DrawSelectCelda "CELDA7","","","0", "","i_pago",rstAux,"","codigo","Descripcion","",""
					rstAux.Close
					DrawCelda2 "CELDA", "left", true, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					DrawCelda2 "CELDA", "left", true, LitFechaCobro
					DrawInputCelda "CELDA","","",10,0,"","fechacobro",EncodeForHtml(date)
				CloseFila%>
				</table>
				<hr/>
		    <%'FLM:20090529:pongo table.%>
			<table width="100%">
			<tr><td>
			    <%NavPaginas lote,lotes,campo,criterio,texto,1%>
			</td><td>
			<%'FLM:20090529:Mostramos el total del importe seleccionado en la remesa.%>
	            <table width="100%">
	            <tr>
	                <td class="CELDARIGHT" ><strong><%=LitImpTotCobrar%>:</strong>&nbsp;<span id="totalACobrar"><%=EncodeForHtml(formatnumber(0,rst("ndecimales"),-1,0,-1))%></span>&nbsp;<%=EncodeForHtml(d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and Moneda_base<>0",session("dsn_cliente"))) %></td>
	            </tr>
	            </table>
	        </td></tr></table>	
	        <%'''''''''''''''''' %>        

				<table width='100%' border='0' cellspacing="1" cellpadding="1">
				<%DrawFila color_blau
					DrawCelda2 "CABECERA", "left", true, LitSelPedidos
				CloseFila%>
				</table>
				<table width='100%' border='0' cellspacing="1" cellpadding="1">
				<%'Fila de encabezado
				DrawFila color_fondo%>
					<td class=CELDALEFT>
						<input type="checkbox" name="check" value="true" onclick="seleccionar();">
					</td>
					<%if tabla="vencimientos_salida" then DrawCelda "ENCABEZADOL","","",0,LitNumRecibo
					if tabla <> "tickets_cli" then 
						DrawCelda "ENCABEZADOL","","",0,LitNumFactura
					else
						DrawCelda "ENCABEZADOL","","",0,LitNumTicket
					end if
					DrawCelda "ENCABEZADOR","","",0,LitFecha
					DrawCelda "ENCABEZADOL","","",0,LitCliente
					DrawCelda "ENCABEZADOR","","",0,LitImporte
					if tabla="vencimientos_salida" or tabla="tickets_cli" then
						DrawCelda "ENCABEZADOR","","",0,LitImporteCobrado
					end if
					DrawCelda "ENCABEZADOR","","",0,LitDeuda
					DrawCelda "ENCABEZADOL","","",0,LitDivisa
					if tabla="facturas_cli" then DrawCelda "ENCABEZADOL","","",0,LitVencimientos
				CloseFila

				VinculosPagina(MostrarFacturasCli)=1:VinculosPagina(MostrarClientes)=1
				CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

				fila=1
				while not rst.eof and fila<=MAXPAGINA
					'Seleccionar el color de la fila.
					if ((fila+1) mod 2)=0 then
						color=color_blau
					else
						color=color_terra
					end if
					if tabla="facturas_cli" then
						NDocumento=rst("nfactura")
					elseif tabla="vencimientos_salida" then
						NDocumento=rst("nrecibo")
					elseif tabla="tickets_cli" then
						NDocumento=rst("nticket")
					end if
					CheckCadena NDocumento
					DrawFila color%>
						<td class=CELDALEFT>
							<%'if tabla="vencimientos_salida" or tabla="tickets_cli" then
								'texto_check="onclick='cambiar_importecob(" & fila & "," & reemplazar(reemplazar(formatnumber(null_z(rst("importecob")),rst("ndecimales"),-1,0,-1),".",""),",",".") & ")'"
								''texto_check="onclick='cambiar_importecob(" & fila & ")'"
							'else
								'texto_check=""
							'end if
							'FLM:20090529:Saco esto xq se tiene que ejecutar siempre..
							texto_check="onclick='cambiar_importecob(" & fila & ")'"
							%>
							<input type="hidden" name="cero<%=fila%>" value="0">
							<input type="checkbox" name="check<%=fila%>" value="<%=EncodeForHtml(NDocumento)%>" <%=texto_check%>>
						</td>
						<%if tabla="vencimientos_salida" then
							DrawCelda "CELDALEFT","","",0,EncodeForHtml(trimCodEmpresa(NDocumento))%>
							<td class="CELDALEFT" align="left">
								<%=Hiperv(OBJFacturasCli,rst("nfactura"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("nfactura"))),LitVerFactura)%>
							</td>
						<%elseif tabla="facturas_cli" then%>
							<td class="CELDALEFT" align="left">
								<%=Hiperv(OBJFacturasCli,rst("nfactura"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("nfactura"))),LitVerFactura)%>
							</td>
						<%elseif tabla="tickets_cli" then%>
							<td class="CELDALEFT" align="left">
								<%=Hiperv(OBJTickets,rst("nventa"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(trimCodEmpresa(rst("nventa"))),LitVerFactura)%>
							</td>
						<%end if
						DrawCelda "CELDARIGHT","","",0,EncodeForHtml(rst("fecha"))
						'DrawCelda "CELDALEFT","","",0,rst("rsocial")
						DrawCelda "CELDALEFT","","",0,Hiperv(OBJClientes,rst("ncliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),EncodeForHtml(rst("rsocial")),LitVerCliente)
						if tabla="facturas_cli" then
							TotalDocumento=formatnumber(null_z(rst("total_factura")),rst("ndecimales"),-1,0,-1)
						elseif tabla="vencimientos_salida" then
							TotalDocumento=formatnumber(null_z(rst("importe")),rst("ndecimales"),-1,0,-1)
						elseif tabla="tickets_cli" then
							TotalDocumento=formatnumber(null_z(rst("importe")),rst("ndecimales"),-1,0,-1)
						end if
						DrawCelda "CELDARIGHT","","",0,EncodeForHtml(TotalDocumento)
						if tabla="vencimientos_salida" then%>
							<input type="hidden" name="h_deuda<%=fila%>" value="<%=EncodeForHtml(replace(null_z(rst("deuda")),",","."))%>">
							<%DrawInputCeldaAction "CELDARIGHT","","",10,0,"","importecob" & fila,0,"onchange","cambiar_importe(" & fila & "," & EncodeForHtml(replace(null_z(rst("deuda")),",",".")) & "," & EncodeForHtml(null_z(rst("importecob"))) & ")",false%>
							<input type="hidden" name="h_importecob<%=fila%>" value="<%=EncodeForHtml(rst("importecob"))%>">
						<%elseif tabla="tickets_cli" then%>
							<input type="hidden" name="h_deuda<%=fila%>" value="<%=EncodeForHtml(replace(null_z(rst("deuda")),",","."))%>">
							<%DrawInputCeldaAction "CELDARIGHT","","",10,0,"","importecob" & fila,EncodeForHtml(replace(null_z(rst("deuda")),",",".")),"onchange","cambiar_importe(" & fila & "," & EncodeForHtml(replace(null_z(rst("deuda")),",",".")) & "," & EncodeForHtml(null_z(rst("cobrado"))) & ")",false%>
							<input type="hidden" name="h_importecob<%=fila%>" value="<%=EncodeForHtml(rst("cobrado"))%>">
						<%else %>
						<input type="hidden" name="h_deuda<%=fila%>" value="<%=EncodeForHtml(replace(rst("deuda"),",","."))%>">
						<%end if
						'FLM:20090602:campo que guarda el factor de conversión de las divisas.
				        %><input type="hidden" name="factcambio<%=fila%>" value="<%=EncodeForHtml(rst("factcambio"))%>" /><%
				        
						DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("deuda")),rst("ndecimales"),-1,0,-1))
						DrawCelda "CELDALEFT","","",0, EncodeForHtml(rst("abreviatura"))
						if tabla="facturas_cli" then DrawCelda "CELDALEFT","","",0,iif(rst("NumeroVencimientos")>0,LitSi,LitNo)
					CloseFila
					fila=fila+1
					rst.MoveNext
				wend%>
				</table>
				<input type="hidden" name="h_nfilas" value="<%=fila-1%>">
				<input type="hidden" name="h_tabla" value="<%=EncodeForHtml(tabla)%>">
				<%rst.close%>
				<hr/>
				<%NavPaginas lote,lotes,campo,criterio,texto,2
			else%>
				<script>
					window.alert("<%=LitMsgNoDocumentos%>");
					caju = document.cobros_param.caju.value;
					parent.window.frames["botones"].document.location = "cobros_param_bt.asp?mode=add";
					document.location="cobros_param.asp?mode=add&caju=" + caju;
				</script>
			<%end if
	end if%>
</form>
<%end if
set rstAux=nothing
set rstAux2=nothing
set rstSelect=nothing
set rstIvas=nothing
set rstAlb=nothing
set rstDetAlb=nothing
set rstPed=nothing
set rst=nothing
set rstCofo=nothing
%>
</body>
</html>