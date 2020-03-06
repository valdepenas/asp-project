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
function pintar_saltos_nuevo(texto)
	texto=Replace(texto,"&#10;","")
	texto=Replace(texto,"&#13;","<br>")
	pintar_saltos_nuevo=texto
end function
%>
<%
''IML 23/06/03: Migracion a monobase
''IML 04/11/03: Modificación en la función de eliminar registros
''JCI 03/04/04: No se pueden modificar traspasos cuyos apuntes de caja están en un cierre
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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

<!--#include file="Ahoja_gastos.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/tabs.js.inc" -->
<!--#include file="../js/calendar.inc" -->

<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

<script language="javascript" type="text/javascript" src="/lib/js/ListSearch.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
animatedcollapse.addDiv('CABECERA', 'fade=1');
animatedcollapse.addDiv('DETALLES', 'fade=1');

animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
    //$: Access to jQuery
    //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
    //state: "block" or "none", depending on state
}

animatedcollapse.init();

function AnyadirApuntes(ntraspaso)
{
	pagina="../central.asp?pag1=administracion/anyadirApuntes.asp&mode=select1&viene=" + ntraspaso + "&ndoc=traspasos_caja&pag2=administracion/anyadirApuntes_bt.asp";
	ven=AbrirVentana(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
}

function cambiarfecha(fecha,modo)
{
	var fecha_ar=new Array();
	if (fecha!="")
	{
		suma=0;
		fecha_ar[suma]="";
		l=0
		while (l<=fecha.length)
		{
			if (fecha.substring(l,l+1)=='/')
			{
				suma++;
				fecha_ar[suma]="";
			}
			else
			{
				if (fecha.substring(l,l+1)!='') fecha_ar[suma]=fecha_ar[suma] + fecha.substring(l,l+1);
			}
			l++;
		}
		if (suma!=2)
		{
			window.alert("<%=LitFechaMal%> en el campo " + modo );
			return false;
		}
		else
		{
			nonumero=0;
			while (suma>=0 && nonumero==0)
			{
				if (isNaN(fecha_ar[suma])) nonumero=1;
				if (fecha_ar[suma].length>2 && suma!=2) nonumero=1;
				if (fecha_ar[suma].length>4 && suma==2) nonumero=1;
				suma--;
			}

			if (nonumero==1)
			{
				window.alert("<%=LitFechaMal%> en el campo " + modo);
				return false;
			}
		}
	}
	return true;
}

//Funciones para 'perseguir' la barra de scroll
function setVariables()
{
	if (navigator.appName == "Netscape")
	{
		v=".top=";
		dS="document.";
		sD="";
		y="window.pageYOffset";
	}
	else
	{
		v=".pixelTop=";
		dS="";
		sD=".style";
		y="document.body.scrollTop";
   }
}

function checkLocation()
{
	object="navegar";
	yy=eval(y);
	eval(dS+object+sD+v+yy);
	setTimeout("checkLocation()",10);
}

//Redirecciona a la opcion pulsada en la capa de navegación entre registros
function Navegar(destino,origen)
{
	document.traspasos_caja.action="traspasos_caja.asp?nmovimiento=" + origen + "&donde=" + destino + "&mode=search";
	document.traspasos_caja.submit();
}

function TraerResponsable()
{
	document.traspasos_caja.action="traspasos_caja.asp?responsable=" + document.traspasos_caja.responsable.value + "&mode=traerresponsable&submode=" + document.traspasos_caja.mode.value;
	document.traspasos_caja.submit();
}

function cambiarimporte()
{
	//document.traspasos_caja.importe.value=document.traspasos_caja.importe.value.replace(".",",");
}

function Redimensionar()
{
    var alto = jQuery(window).height();
    var diference = 240;
    var dir_default = 140;

    if (alto > dir_default)
    {
        if (alto - diference > dir_default) jQuery("#frDetalles").attr("height", alto - diference);
        else jQuery("#frDetalles").attr("height", dir_default);
    }
    else jQuery("#frDetalles").attr("height", dir_default);
}

jQuery(window).resize(function (){Redimensionar();});
</script>
<%modoPantalla=EncodeForHtml(Request.QueryString("mode"))
if modoPantalla & ""="" then
    modoPantalla=EncodeForHtml(request.Form("mode"))
end if%>

<body class="BODY_ASP">
<%Function Posicion(ntraspaso)
	if ntraspaso & "">"" then
	    rstAux.CursorLocation=3
		rstAux.Open "select ntraspaso from traspasos with(NOLOCK) where ntraspaso like '"+Session("ncliente")+"%' order by fecha desc,ntraspaso desc",session("dsn_cliente")
		cont=1
		if not rstAux.EOF then
			while (rstAux("ntraspaso")<>ntraspaso) and (not rstAux.eof)
				rstAux.movenext
				cont=cont + 1
			wend
			temp="" & cont & " de " & rstAux.recordcount
		end if
		rstAux.close
		Posicion=temp
	end if
end Function

'******** GUARDAR REGISTRO *******
'Actualiza los datos del registro cuando se pulsa el botón de 'GUARDAR'
sub GuardarRegistro()

    'ega 05/06/2008 quitar los rst.addnew y rst.update y hacer uso de command
    command.CommandText= "select ntraspaso, fecha, descripcion, responsable, cajaorg, cajadest, importe, divisa, medio, serie, contabilizado from traspasos where ntraspaso like '"+Session("ncliente")+"%' and ntraspaso='" & TmpNTraspaso & "'"
    set rstTraspaso = command.Execute()
   
	if rstTraspaso.eof then
        modo="añadir"
        ''dig=mid(year(date),3,2)
        anyo=right(TmpFecha,2)
        
        command.CommandText= "select max(ntraspaso) as total from traspasos with(nolock) where ntraspaso like '" & session("ncliente") & trimCodEmpresa(TmpSerie) & anyo & "%'"
        set rstNumeroTraspasos = command.Execute()
    
        if not rstNumeroTraspasos.eof then
			if rstNumeroTraspasos("total") & "">"" then
				num=rstNumeroTraspasos("total")
				num=cint(right(num,6))+1
			else
				num="1"
			end if
		else
			num="1"
		end if
		rstNumeroTraspasos.close
		ntraspaso = session("ncliente") & trimCodEmpresa(TmpSerie) & anyo & completar(num,6,"0")
        fecha = "" 
        descripcion = "" 
        responsable = ""
        cajaorg =""
        cajadest =""
        importe = ""
        divisa =""
        medio =""
        serie =""
        contabilizado = 0
		TmpNTraspaso=session("ncliente") & trimCodEmpresa(TmpSerie) & anyo & completar(num,6,"0")
    else
    	modo="editar"
      	ntraspaso = rstTraspaso("ntraspaso")
        fecha = rstTraspaso("fecha")
        descripcion = rstTraspaso("descripcion")
        responsable = rstTraspaso("responsable")
        cajaorg =rstTraspaso("cajaorg")
        cajadest =rstTraspaso("cajadest")
        importe = rstTraspaso("importe")
        divisa = rstTraspaso("divisa")
        medio =rstTraspaso("medio")
        serie =rstTraspaso("serie")
        contabilizado = rstTraspaso("contabilizado")    
                 	
    	TmpNTraspaso = rstTraspaso("ntraspaso")
		importe_ant=CambioDivisa(importe,divisa,d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente")))
    end if
    
    if nz_b(contabilizado)=-1 then
		contabilizado=-1
	else
		serie = TmpSerie
		fecha = Nulear(TmpFecha)
		descripcion=Nulear(TmpDescripcion)
		responsable=Session("ncliente")+(Nulear(TmpResponsable))
		' Comprobamos si el traspaso tiene apuntes asociados, en cuyo caso no podrá modificarse el importe'
		'ega 04/06/2008 agrego with(nolock) y cuento el numero de caja que cumplen las condiciones
		command.CommandText= "select count(caja) as numero from caja with(nolock) where caja like '" & session("ncliente") & "%' and ntraspaso like '"+Session("ncliente")+"%' and ntraspaso='" & TmpNTraspaso & "'"
        set rstNumeroCaja = command.Execute()
        
		if rstNumeroCaja("numero")=0 then
			importe=null_z(TmpImporte)
			cajaorg=Nulear(TmpCajaOrg)
			cajadest=Nulear(TmpCajaDest)
			medio=Nulear(TmpMedio)
		end if
		rstNumeroCaja.close

		divisa=iif(TmpDivisa & "">"",TmpDivisa,d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente")))
	end if

    'ega 10/06/2008 pongo la consulta de actualizacion fuera del if y asignandole la variable contabilizado
    if nz_b(TmpContabilizado)=-1 then
		contabilizado=1
	else
		contabilizado=0
	end if

    'ega 04/05/2008 agrego with(rowlock) y realizo solamente un update, incluyendo las dos condiciones de tanotacion mediante un in (es más optimo que OR)
    command.CommandText="update caja with(updlock) set contabilizado="&contabilizado&" where caja like '" & session("ncliente") & "%' and ndocumento like '"+Session("ncliente")+"%' and ndocumento='" & TmpNTraspaso & "' and tanotacion in ('ENTRADA','SALIDA')"
    Command.Execute , , adExecuteNoRecords


	TmpImporteAux=CambioDivisa(TmpImporte,iif( TmpDivisa & "">"",TmpDivisa,d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente"))),d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("dsn_cliente")))

  	if rstTraspaso.eof then   ''insert
        command.CommandText="insert into traspasos (ntraspaso, fecha, descripcion, responsable, cajaorg, cajadest, importe, divisa, medio, serie, contabilizado) " & _
                            " values ('"&ntraspaso&"','"&fecha&"','"&descripcion&"','"&responsable&"','"&cajaorg&"','"&cajadest&"',convert(money, replace('"&importe&"', ',','.')),'"&divisa&"','"&medio&"','"&serie&"',"&contabilizado&")"
    else  ''update
        command.CommandText="update traspasos with(updlock) set fecha = '"&fecha&"', descripcion = '"&descripcion&"',responsable = '"&responsable&"', cajaorg= '"&cajaorg&"', cajadest= '"&cajadest&"', importe= convert(money, replace('"&importe&"', ',','.')), divisa= '"&divisa&"', medio= '"&medio&"', serie= '"&serie&"', contabilizado= "&contabilizado&"" & _ 
                            " where ntraspaso like '" & session("ncliente") & "%' and ntraspaso = '"&TmpNTraspaso &"'"
    end if        

    command.Execute , , adExecuteNoRecords
    rstTraspaso.close
    
	if modo="añadir" then
		'vamos a auditar
		auditar_ins_bor session("usuario"),ntraspaso,"","alta","","","traspasos_cajas"
	end if

	' *** Añadimos/Modificamos un gasto a la caja origen cuyo importe, divisa y medio de pago sea el del traspaso
	'ega 04/06/2008 agrego las columnas necearias
	
	command.CommandText= "select caja, nanotacion, tanotacion, fecha, divisa, descripcion, ndocumento, tdocumento, importe, medio from caja with(nolock) where caja like '" & session("ncliente") & "%' and ndocumento like '"+Session("ncliente")+"%' and ndocumento='" & TmpNTraspaso & "' and tanotacion='SALIDA'"
    set rstCaja = command.Execute()

	if rstCaja.eof then
		nanotacion=d_max("nanotacion","caja","caja like '" & session("ncliente") & "%' and caja='" & Nulear(TmpCajaOrg) & "'",session("dsn_cliente")) + 1
	else
		borrar=""
		caja_ant=""
		importeC=null_z(rstCaja("importe"))
		medioC=rstCaja("medio")	
		cajaC=rstCaja("caja")	
		if rstCaja("caja")<>Nulear(TmpCajaOrg) then
			borrar=rstCaja("nanotacion")
			caja_ant=rstCaja("caja")
			nanotacion=d_max("nanotacion","caja","caja like '" & session("ncliente") & "%' and caja='" & Nulear(TmpCajaOrg) & "'",session("dsn_cliente")) + 1
		else
			nanotacion=rstCaja("nanotacion")
		end if
	end if
    
    insertCajaOri = ""
    insertCajaDest = ""

    if contabilizado=0 then
		tanotacionC="SALIDA"
		fechaC=Nulear(TmpFecha) & " " & replace(replace(time, "a.m.", "am"), "p.m.", "pm")
		' Comprobamos si el traspaso tiene apuntes asociados, en cuyo caso no podrá modificarse el importe'
		'ega 04/06/2008 agrego with(nolock) y cuento el numero de caja que cumplen las condiciones
		command.CommandText= "select count(caja) as numero from caja with(NOLOCK) where caja like '" & session("ncliente") & "%' and ntraspaso like '"+Session("ncliente")+"%' and ntraspaso='" & TmpNTraspaso & "'"
        set rstCaja2 = command.Execute()
		
		if rstCaja2("numero")=0 then
		    importeC=null_z(TmpImporte)
			cajaC=Nulear(TmpCajaOrg)
			medioC=Nulear(TmpMedio)
		end if
		rstCaja2.close
		divisaC=Nulear(divisa)
		descripcionC=TmpDescripcion
		ndocumentoC=TmpNTraspaso
		tdocumentoC="TRASPASO ENTRE CAJAS"
			
        if rstCaja.eof or borrar &"">"" then 'insert
            'command.CommandText="insert into caja (caja,nanotacion,tanotacion,fecha, importe, descripcion,ndocumento, tdocumento, medio, divisa ) values " & _ 
            '                "('"&cajaC &"','"&nanotacion &"','"&tanotacionC&"','"&fechaC&"', convert(money, replace('"&importeC&"', ',','.')), '"&descripcionC&"','"&ndocumentoC&"','"&tdocumentoC&"', '"&medioC&"', '"&divisaC&"')"
            insertCajaOri = "(select '"&cajaC &"','"&nanotacion &"','"&tanotacionC&"', '"&fechaC&"', convert(money, replace('"&importeC&"', ',','.')), '"&descripcionC&"','"&ndocumentoC&"','"&tdocumentoC&"', '"&medioC&"', '"&divisaC&"')"                            
        else ''update
            command.CommandText="update caja with(updlock) set tanotacion = '"&tanotacionC&"',fecha = '"&fechaC&"', importe=convert(money, replace('"&importeC&"', ',','.')), descripcion='"&descripcionC&"',ndocumento='"&ndocumentoC&"', tdocumento='"&tdocumentoC&"', medio='"&medioC&"', divisa='"&divisaC&"'" & _ 
                                " where caja like '" & session("ncliente") & "%' and caja = '"&cajaC&"' and nanotacion ='"&nanotacion &"'"
               
            on error resume next
		        Command.Execute , , adExecuteNoRecords

                if err.number<>0 then
    	        %><script language="javascript" type="text/javascript">
	    	        window.alert("<%=LitErrInsTras%>");
	            </script><%
                end if
             on error goto 0
        end if
		rstCaja.close
		
		if borrar & "">"" and caja_ant & "">"" then
			command.CommandText= "delete from caja with(rowlock) where caja like '" & session("ncliente") & "%' and ndocumento like '"+Session("ncliente")+"%' and ndocumento='" & TmpNTraspaso & "' and caja='" & caja_ant & "' and nanotacion='" & borrar & "'"
			command.Execute , , adExecuteNoRecords
		end if
   	else
		rstCaja.close
	end if

	' *** Añadimos/Modificamos un ingreso a la caja destino cuyo importe, divisa y medio de pago sea el del traspaso
	'ega 04/05/2008 agrego las columnas necearias
	command.CommandText= "select caja,nanotacion,tanotacion,fecha,importe,medio,descripcion,ndocumento,tdocumento,divisa from caja with(nolock) where caja like '" & session("ncliente") & "%' and ndocumento like '"+Session("ncliente")+"%' and ndocumento='" & TmpNTraspaso & "' and tanotacion='ENTRADA'"
	set rstCaja = command.Execute()
    
	if rstCaja.eof then
		nanotacion=d_max("nanotacion","caja","caja like '" & session("ncliente") & "%' and caja='" & Nulear(TmpCajaDest) & "'",session("dsn_cliente")) + 1
	else
		borrar=""
		caja_ant=""
		importeC=null_z(rstCaja("importe"))
		medioC=rstCaja("medio")	
		cajaC=rstCaja("caja")	


		if rstCaja("caja")<>Nulear(TmpCajaDest) then
			borrar=rstCaja("nanotacion")
			caja_ant=rstCaja("caja")
			nanotacion=d_max("nanotacion","caja","caja like '" & session("ncliente") & "%' and caja='" & Nulear(TmpCajaDest) & "'",session("dsn_cliente")) + 1
			cajaC=rstCaja("caja")
		else
			nanotacion=rstCaja("nanotacion")
			cajaC=rstCaja("caja")
		end if
	end if
		 
		 
	if contabilizado=0 then
		tanotacionC="ENTRADA"
		fechaC=Nulear(TmpFecha) & " " & replace(replace(time, "a.m.", "am"), "p.m.", "pm")
		' Comprobamos si el traspaso tiene apuntes asociados, en cuyo caso no podrá modificarse el importe'
		'ega 04/06/2008 cuento el numero de caja que cumplen las condiciones
		command.CommandText= "select count(caja) as numero from caja with(NOLOCK) where caja like '" & session("ncliente") & "%' and ntraspaso like '"+Session("ncliente")+"%' and ntraspaso='" & TmpNTraspaso & "'"
        set rstCaja2 = command.Execute()
		
		'if rstAux2.eof then
		if rstCaja2("numero")=0 then
			importeC=null_z(TmpImporte)
			cajaC=Nulear(TmpCajaDest)
			medioC=Nulear(TmpMedio)
		end if
		rstCaja2.close
		
		divisaC=Nulear(divisa)
		descripcionC=TmpDescripcion
		ndocumentoC=TmpNTraspaso
		tdocumentoC="TRASPASO ENTRE CAJAS"
    
      	if rstCaja.eof or borrar &"">"" then 'insert
            'command.CommandText="insert into caja (caja,nanotacion,tanotacion,fecha, importe, descripcion,ndocumento, tdocumento, medio, divisa ) values " & _ 
            '                "('"&cajaC &"','"&nanotacion &"','"&tanotacionC&"','"&fechaC&"', convert(money, replace('"&importeC&"', ',','.')), '"&descripcionC&"','"&ndocumentoC&"','"&tdocumentoC&"', '"&medioC&"', '"&divisaC&"')"
            insertCajaDest = "(select '"&cajaC &"','"&nanotacion &"','"&tanotacionC&"','"&fechaC&"',  convert(money, replace('"&importeC&"', ',','.')), '"&descripcionC&"','"&ndocumentoC&"','"&tdocumentoC&"', '"&medioC&"', '"&divisaC&"')"
        else ''update
            command.CommandText="update caja with(updlock) set tanotacion = '"&tanotacionC&"',fecha = '"&fechaC&"', importe=convert(money, replace('"& importeC&"', ',','.')), descripcion='"&descripcionC&"',ndocumento='"&ndocumentoC&"', tdocumento='"&tdocumentoC&"', medio='"&medioC&"', divisa='"&divisaC&"'" & _ 
                                " where caja like '" & session("ncliente") & "%' and caja = '"&cajaC&"' and nanotacion ='"&nanotacion &"'"
             on error resume next
                    command.Execute , , adExecuteNoRecords    
             if err.number<>0 then
	                %><script language="javascript" type="text/javascript">
		                window.alert("<%=LitErrInsTras%>");
	                </script><%
            end if
            on error goto 0
        end if	
		rstCaja.close

		if borrar & "">"" and caja_ant & "">"" then
			command.CommandText= "delete from caja with(rowlock) where caja like '" & session("ncliente") & "%' and ndocumento like '"+Session("ncliente")+"%' and ndocumento='" & TmpNTraspaso & "' and caja='" & caja_ant & "' and nanotacion='" & borrar & "'"
			command.Execute , , adExecuteNoRecords
		end if
	else
		rstCaja.close
	end if

    'ega 06/06/08 hago las dos inserciones en la tabla caja en una sola ejecución
    if contabilizado  = 0 and insertCajaOri &"" > "" and insertCajaDest &"" > "" then
       sqlTemporal=  "insert into caja (caja,nanotacion,tanotacion,fecha, importe, descripcion,ndocumento, tdocumento, medio, divisa ) " & _ 
                             insertCajaOri & " union " & insertCajaDest

        command.CommandText= sqlTemporal
        on error resume next
            command.Execute , , adExecuteNoRecords    
        if err.number<>0 then
	        %><script language="javascript" type="text/javascript">
		        window.alert("<%=LitErrInsTras%>");
	        </script><%
        end if
        on error goto 0
    end if

    set rstCaja = nothing
    set rstCaja2 = nothing
    set rstTraspaso = nothing
end sub
'******** FIN GUARDAR REGISTRO *******

'******** ELIMINAR REGISTRO *******
function EliminarRegistro()
	'vamos a auditar
	auditar_ins_bor session("usuario"),TmpNTraspaso,"","baja","","","traspasos_cajas"

    ''MPC 30/11/2011 Se cambia todo por un procedimiento almacenado
    set connDelete = Server.CreateObject("ADODB.Connection")
	set commandDelete =  Server.CreateObject("ADODB.Command")

	connDelete.open session("dsn_cliente")
	commandDelete.ActiveConnection =conn
	commandDelete.CommandTimeout = 0
	commandDelete.CommandText="DeleteTransfer"
	commandDelete.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	commandDelete.Parameters.Append command.CreateParameter("@ncompany",adVarChar,adParamInput,5,session("ncliente"))
	commandDelete.Parameters.Append command.CreateParameter("@ntransfer",adVarChar,adParamInput,20,TmpNTraspaso)
	commandDelete.Parameters.Append command.CreateParameter("@return",adInteger,adParamOutput)
	commandDelete.Execute,,adExecuteNoRecords
	result=commandDelete.Parameters("@return").Value
	connDelete.close
	set commandDelete=nothing
	set connDelete=nothing

    EliminarRegistro = result
end function

'******************************************************************************
'Crea la tabla que contiene la barra de grupos de datos.
sub BarraNavegacion(modo)
    if modo="add" or mode="edit" then%>
        <script language="javascript" type="text/javascript">
            jQuery("#S_CABECERA").show();
            jQuery("#S_DETALLES").hide();
        </script>
    <%else%>
        <script language="javascript" type="text/javascript">
            jQuery("#S_CABECERA").hide();
            jQuery("#S_DETALLES").show();
        </script>
    <%end if
end sub
'******************************************************************************
'*****************************************************************************'
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************'
'*****************************************************************************'
const borde=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
	
	'ega 09/06/08
    set conn = Server.CreateObject("ADODB.Connection")
    set command =  Server.CreateObject("ADODB.Command")

    conn.open session("dsn_cliente")
    		    
    command.ActiveConnection =conn
    command.CommandTimeout = 0
    command.CommandType = adCmdText

%>
	<form name="traspasos_caja" method="post">
		<% PintarCabecera "traspasos_caja.asp"

		'********** '
		' Cursores y conexiones '
		set rst = Server.CreateObject("ADODB.Recordset")
		set rst2= Server.CreateObject("ADODB.Recordset")
		set rstAux = Server.CreateObject("ADODB.Recordset")
		set rstAux2 = Server.CreateObject("ADODB.Recordset")
		set rstDet = Server.CreateObject("ADODB.Recordset")

		'********** '
		'Leer parámetros de la página '
		mode = Request.QueryString("mode")
            %>
			<input type="hidden" name="mode" value="<%=EncodeForHtml(mode)%>">
			<input type="hidden" name="campo" value="<%=EncodeForHtml(campo)%>">
			<input type="hidden" name="criterio" value="<%=EncodeForHtml(criterio)%>">
			<input type="hidden" name="texto" value="<%=EncodeForHtml(texto)%>">
		<%if Request.QueryString("submode")>"" then
			submode=Request.QueryString("submode")
		else
			submode=Request.Form("submode")
		end if

		if Request.QueryString("ndoc")>"" then
			ndoc=limpiaCadena(Request.QueryString("ndoc"))
		else
			ndoc=limpiaCadena(Request.Form("ndoc"))
		end if

		if ndoc="" then
			if Request.QueryString("ntraspaso")>"" then
				TmpNTraspaso=limpiaCadena(Request.QueryString("ntraspaso"))
			else
				TmpNTraspaso=limpiaCadena(Request.Form("ntraspaso"))
			end if
		else
			TmpNTraspaso=ndoc
		end if
		CheckCadena TmpNtraspaso%>
		<input type="hidden" name="ntraspaso" value="<%=EncodeForHtml(TmpNTraspaso)%>">
		<%if Request.QueryString("fecha")>"" then
			TmpFecha=limpiaCadena(Request.QueryString("fecha"))
		else
			TmpFecha=limpiaCadena(Request.Form("fecha"))
		end if

		if request.querystring("contabilizado")>"" then
			TmpContabilizado=limpiaCadena(request.querystring("contabilizado"))
		else
			TmpContabilizado=limpiaCadena(request.form("contabilizado"))
		end if

		if request.querystring("serie")>"" then
			TmpSerie=limpiaCadena(request.querystring("serie"))
		else
			TmpSerie=limpiaCadena(request.form("serie"))
		end if
		if TmpSerie="" and mode="add" then
			' Obtenemos la serie por defecto
			TmpSerie=d_lookup("nserie","series","nserie like '" & session("ncliente") & "%' and tipo_documento='TRASPASO ENTRE CAJAS' and  pordefecto=1",session("dsn_cliente"))
		end if

		if Request.QueryString("divisa")>"" then
			TmpDivisa=limpiaCadena(Request.QueryString("divisa"))
		else
			TmpDivisa=limpiaCadena(Request.Form("divisa"))
		end if

		if Request.QueryString("responsable")>"" then
			TmpResponsable=limpiaCadena(Request.QueryString("responsable"))
		else
			TmpResponsable=limpiaCadena(Request.Form("responsable"))
			if TmpResponsable<>"" and mode="add" then TmpResponsable=Session("ncliente")&TmpResponsable
		end if

		if Request.QueryString("descripcion")>"" then
			TmpDescripcion=limpiaCadena(Request.QueryString("descripcion"))
		else
			TmpDescripcion=limpiaCadena(Request.Form("descripcion"))
		end if

		if Request.QueryString("cajaorg")>"" then
			TmpCajaOrg=limpiaCadena(Request.QueryString("cajaorg"))
		else
			TmpCajaOrg=limpiaCadena(Request.Form("cajaorg"))
		end if

		if Request.QueryString("cajadest")>"" then
			TmpCajaDest=limpiaCadena(Request.QueryString("cajadest"))
		else
			TmpCajaDest=limpiaCadena(Request.Form("cajadest"))
		end if

		if Request.QueryString("medio")>"" then
			TmpMedio=limpiaCadena(Request.QueryString("medio"))
		else
			TmpMedio=limpiaCadena(Request.Form("medio"))
		end if

		if Request.QueryString("importe")>"" then
			TmpImporte=limpiaCadena(Request.QueryString("importe"))
		else
			TmpImporte=limpiaCadena(Request.Form("importe"))
		end if

		if request.querystring("nomoperador")>"" then
			TmpNombre=limpiaCadena(request.querystring("nomoperador"))
		else
			TmpNombre=limpiaCadena(request.Form("nomoperador"))
		end if

		if Request.QueryString("viene")>"" then
			viene=limpiaCadena(Request.QueryString("viene"))
		else
			viene=limpiaCadena(Request.Form("viene"))
		end if
		if viene & "">"" then
			if Request.QueryString("ndoc")>"" then
				TmpCajaOrg=limpiaCadena(Request.QueryString("ndoc"))
			else
				TmpCajaOrg=limpiaCadena(Request.Form("ndoc"))
			end if
			TmpNTraspaso=""
		end if

		fechaR=limpiaCadena(request.form("fecha"))

		alarma "traspasos_caja.asp"

 		'ACCIONES A REALIZAR '
		'********** '
 		' 1ª acción: TRAERRESPONSABLE '
		if mode="traerresponsable" then
			submode2="traerresponsable"
			if TmpResponsable> "" then
				TmpResponsable=Session("ncliente")+TmpResponsable
				responsable=d_lookup("nombre","personal","dni like '" & Session("ncliente") & "%' and dni='" & TmpResponsable & "'",session("dsn_cliente"))
				if responsable="" then
					%><script language="javascript" type="text/javascript">
						window.alert("<%=LitMsgResponsableNoExiste%>");
					</script><%
					TmpResponsable=TmpResponsable2
				else
				    'ega 09/06/2008 uso command
				    command.CommandText= "select dni,nombre,fbaja from personal with(NOLOCK) where dni like '" & Session("ncliente") & "%' and dni='" & TmpResponsable & "' and fbaja is null"
                    set rst2 = command.Execute()
					
					if rst2.eof then
						%><script language="javascript" type="text/javascript">
							window.alert("<%=LitMsgResponsableDadoBaja%>");
						</script><%
						TmpResponsable=TmpResponsable2
					else
						TmpNombre=responsable
					end if
					rst2.close
				end if
				mode=submode
				%><script language="javascript" type="text/javascript">
					document.traspasos_caja.mode.value="<%=enc.EncodeForJavascript(mode)%>";
				</script><%
			else
				TmpResponsable=""
				TmpNombre=""
				mode=submode
				%><script language="javascript" type="text/javascript">
					document.traspasos_caja.mode.value="<%=enc.EncodeForJavascript(mode)%>";
				</script><%
			end if
		end if
		
		'********** '
		'2ª acción: MODO SAVE, BROWSE, ADD, EDIT, SEARCH '
		if mode="save" then
			GuardarRegistro
			mode="browse"
		elseif mode="delete" then
			result = EliminarRegistro
            if result = 0 then%>
                <script language="javascript" type="text/javascript">alert("<%=LitErrortDeleteTransfer%>");</script>
                <%mode = "browse"
            else
			    mode = "add"
                TmpNTraspaso=""
                TmpResponsable=""
                %>
                <script language="javascript" type="text/javascript">
                    parent.botones.document.location = "traspasos_caja_bt.asp?mode=add";
                    SearchPage("traspasos_caja_lsearch.asp?mode=init", 0);
                </script>
            <%end if
		end if

		if mode="add" then
		    'ega 04/06/2008 agrego with(NOLOCK)
            rst.cursorlocation=3
			rst.Open "select ntraspaso,descripcion,fecha,importe,contabilizado,serie,divisa,responsable,cajaorg,cajadest,medio from traspasos with(NOLOCK) where ntraspaso like '" & Session("ncliente") & "%' and ntraspaso='" & TmpNTraspaso & "'",session("dsn_cliente")
			''rst.AddNew
		elseif mode="browse" then
			rst.cursorlocation=3
			rst.Open "select ntraspaso,descripcion,fecha,importe,contabilizado,serie,divisa,responsable,cajaorg,cajadest,medio from traspasos with(NOLOCK) where ntraspaso like '" & Session("ncliente") & "%' and ntraspaso='" & TmpNTraspaso & "'",session("dsn_cliente")
		elseif mode="edit" then
            rst.cursorlocation=3
			rst.Open "select ntraspaso,descripcion,fecha,importe,contabilizado,serie,divisa,responsable,cajaorg,cajadest,medio from traspasos with(NOLOCK) where ntraspaso like '" & Session("ncliente") & "%' and ntraspaso='" & TmpNTraspaso & "'",session("dsn_cliente")
		elseif mode="search" then
			
		end if%>
		<script language="javascript" type="text/javascript">
			document.traspasos_caja.ntraspaso.value="<%=enc.EncodeForJavascript(TmpNTraspaso)%>";
		</script>

		<%'********** '
		'CABECERA CON EL TITULO Y LOS FORMATOS DE IMPRESION Y LA CAPA DE NAVEGACION '%>
		<!--<br/>-->
		<%if mode<>"search" then
			if (mode="browse" or mode="edit" or mode="add") then
                VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarCentros)=1:VinculosPagina(MostrarContacto)=1
				CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
				' Mostrar la barra de pestañas '
				BarraNavegacion mode%>
        <div class="headers-wrapper"><%
            DrawDiv "header-date","",""
            DrawLabel "","",LitFecha
               if mode="browse" then
                   DrawSpan "", "", EncodeForHtml(rst("fecha")), "" 
                   'response.write(rst("fecha"))
               else
                   ''EligeCelda "input", mode,"CELDA valing='top'","","",0,"","fecha",10,iif(mode="add",iif(p_fecha>"",p_fecha,date()),rst("fecha"))
                   DrawInput "width150px", "", "fecha", EncodeForHtml(iif(mode="add",iif(TmpFecha>"",TmpFecha,date()),rst("fecha"))), ""
                   DrawCalendar "fecha"
               end if
            CloseDiv
               %><input type="hidden" name="h_fecha" value="<%=EncodeForHtml(fechaR)%>"><%
               DrawDiv "header-move", "",""
               if mode="browse" or mode="edit" then
                   DrawLabel "","CELDA",LitNTraspaso
                   DrawSpan "CELDA","",EncodeForHtml(trimCodEmpresa(rst("ntraspaso"))),""
               else
                   DrawLabel "","CELDA",LitNTraspaso
               end if
                CloseDiv%>
            </div>
        <table style="width: 100%;"></table>
                <%if not rst.EOF or mode="add" then
                    if mode="add" then
                        %><input type="hidden" name="contab" value="0"><%
                        esDisabled2=""
                    else
				        %><input type="hidden" name="contab" value="<%=EncodeForHtml(nz_b(rst("contabilizado")))%>"><%
				        esDisabled=iif(nz_b(rst("contabilizado")),"DISABLED","")
				        if esDisabled="DISABLED" then
					        esDisabled2="DISABLED"
				        else
				            rst2.CursorLocation=3
					        rst2.open "select caja,nanotacion from caja with(NOLOCK) where caja like '" & Session("ncliente") & "%' and ntraspaso='" & TmpNTraspaso & "'",session("dsn_cliente")
					        if not rst2.eof then
						        esDisabled2="disabled"
					        else
						        esDisabled2=""
					        end if
					        rst2.close
				        end if
                    end if%>

                    <div class="Section" id="S_CABECERA">
                        <a href="#" rel="toggle[CABECERA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                            <div class="SectionHeader">
                                <%=LitCabecera%>
                                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                            </div>
                        </a>
                    <div class="SectionPanel" style="<%=iif(mode="add" or mode="edit","display: ","display: none")%>" id="CABECERA"><%
					    if mode="browse" then
						    'DrawCelda "ENCABEZADOL","","",0,LitSerie + ": "
						    'DrawCelda "CELDA","","",0,trimCodEmpresa(rst("serie"))
                            EligeCeldaResponsive "text", mode, "CELDA", "", "", 0, "", LitSerie, LitSerie, EncodeForHtml(trimCodEmpresa(rst("serie")))
					    else
						    'DrawCelda "CELDA","","",0,LitSerie + ": "
						    rstAux.CursorLocation=3
						    rstAux.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as descripcion from series with(NOLOCK) where nserie like '" & Session("ncliente") & "%' and tipo_documento ='TRASPASO ENTRE CAJAS' and nserie like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
                            if mode="add" then
                                'DrawSelectCelda "CELDA " & iif(mode="edit",esDisabled,""),"","",0,"","serie",rstAux,iif(TmpSerie>"",TmpSerie,""),"nserie","descripcion","",""
                                DrawSelectCelda "CELDA " & iif(mode="edit",esDisabled,""),"CELDA","",0,LitSerie,"serie",rstAux,iif(TmpSerie>"",TmpSerie,""),"nserie","descripcion","",""
                            else
                                'DrawSelectCelda "CELDA " & iif(mode="edit",esDisabled,""),"","",0,"","serie",rstAux,iif(TmpSerie>"",TmpSerie,rst("serie")),"nserie","descripcion","",""
                                DrawSelectCelda "CELDA " & iif(mode="edit",esDisabled,""),"CELDA","",0,LitSerie,"serie",rstAux,iif(TmpSerie>"",TmpSerie,rst("serie")),"nserie","descripcion","",""
                            end if
			 			    rstAux.close
					    end if

                        'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"

                        campo="codigo"
					    campo2="abreviatura"

                        if mode="add" then
                            DIVISA=TmpDivisa
                        else
					        DIVISA=iif(TmpDivisa>"",TmpDivisa,rst("divisa"))
                        end if
					    'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA"),"","",0,LitDivisa+":"
                        if DIVISA & "">"" then
					        dato_celda=Desplegable(mode,campo,campo2,"divisas",DIVISA,"codigo like '" & session("ncliente") & "%' and moneda_base<>0 ")
                        else
                            dato_celda=d_lookup("codigo","divisas","codigo like '" & session("ncliente") & "%' and moneda_base<>0 ",session("dsn_cliente"))
                        end if
					    Estilo=iif(mode="browse","CELDA","CELDA DISABLED")
					    if mode<>"browse" then
						    datoDivisa=iif(TmpDivisa>"", _
						    d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & TmpDivisa & "' ",session("dsn_cliente")), _
						    d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & dato_celda & "' ",session("dsn_cliente")))
                            'EligeCelda "input",mode,"CELDA DISABLED","","",0,LitDivisa,"divisa",0,datoDivisa
                            DrawInputCeldaDisabled "CELDA","","","",0,LitDivisa,"divisa",EncodeForHtml(datoDivisa)
                            
					    else
						    datoDivisa=dato_celda
                            EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitDivisa,LitDivisa,EncodeForHtml(datoDivisa)
					    end if
					    'EligeCelda "input", mode,Estilo,"","",0,"","divisa",5,datoDivisa
                        %><input type="hidden" name="h_divisa" value="<%=iif(TmpDivisa>"",EncodeForHtml(TmpDivisa),EncodeForHtml(dato_celda))%>"><%
                        if mode="add" or mode="edit" then
                            if RstAux.state<>0 then RstAux.close
                        end if
					    n_decimales=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & iif(mode="browse",DIVISA,iif(TmpDivisa>"",TmpDivisa,dato_celda)) & "' ",session("dsn_cliente"))
                        a_breviatura=d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & iif(mode="browse",DIVISA,iif(TmpDivisa>"",TmpDivisa,dato_celda)) & "' ",session("dsn_cliente"))
					    'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA"),"","",0,LitContabilizado+":"
                        if mode="add" then
                            'EligeCelda "check", mode,"CELDA","","",0,"","contabilizado",0,false
                            EligeCelda "check",mode,"CELDA","","",0,LitContabilizado,"contabilizado",0,false
                        else
                            if mode="browse" then
                                EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitContabilizado,LitContabilizado,Visualizar(EncodeForHtml(iif(TmpContabilizado>"",nz_b(TmpContabilizado),nz_b(rst("contabilizado")))))
                            else
					        'EligeCelda "check", mode,"CELDA","","",0,"","contabilizado",0,iif(TmpContabilizado>"",nz_b(TmpContabilizado),nz_b(rst("contabilizado")))
                            EligeCelda "check", mode,"CELDA","","",0,LitContabilizado,"contabilizado",0,EncodeForHtml(iif(TmpContabilizado>"",nz_b(TmpContabilizado),nz_b(rst("contabilizado"))))
                            end if
                        end if

                        'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"

                        if TmpResponsable="" then
							' Buscar el usuario en la tabla personal '
							rstAux.CursorLocation=3
							rstAux.open "select dni from personal with(NOLOCK) where dni like '"+Session("ncliente")+"%' and login='" & session("usuario") & "' ",session("dsn_cliente")
							if not rstAux.eof then
								TmpResponsable=rstAux("dni")
							end if
							rstAux.close
						end if
						if mode="browse" then
							'DrawCelda "ENCABEZADOL style='width:100px'","","",0,LitResponsable+":"
                            DrawDiv "1","",""
                            DrawLabel "","",LitResponsable
							if rst("responsable")&"">"" then
                                DrawSpan "CELDA","",EncodeForHtml(d_lookup("nombre","personal","dni like '"+Session("ncliente")+"%' and dni='" & rst("responsable") & "'",session("dsn_cliente"))),""%><input class='CELDA' type="hidden" name="responsable" value="<%=EncodeForHtml(rst("responsable"))%>" size=10 >
							<%else
                                DrawSpan "","","",""
                              end if
                            CloseDiv
						else
							'DrawCelda "CELDA style='width:100px'","","",0,LitResponsable+":"
                            DrawDiv "1","",""
                            DrawLabel "","",LitResponsable%>
                                <%if mode="add" then%><input class='width15' <%=iif(mode="edit",esDisabled,"")%> type="text" name="responsable" value="<%=iif(TmpResponsable & "">"",EncodeForHtml(trimCodEmpresa(TmpResponsable)) ,"")%>" onchange="TraerResponsable();">
                                <%else%><input class='width15' type="text" name="responsable" size=10 value="<%=EncodeForHtml(iif(TmpResponsable & "">"",trimCodEmpresa(TmpResponsable) ,trimCodEmpresa(rst("responsable"))))%>" onchange="TraerResponsable();">
								<%end if
                                if mode<>"add" then
                                    if nz_b(rst("contabilizado"))=-1 then%>
    									&nbsp;
	    							<%else%><a class='CELDAREFB' href="javascript:AbrirVentana('../administracion/personal_buscar.asp?viene=traspasos_caja&titulo=<%=LitSelPersonal%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPersonal%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
			    					<%end if
                                else%><a class='CELDAREFB' href="javascript:AbrirVentana('../administracion/personal_buscar.asp?viene=traspasos_caja&titulo=<%=LitSelPersonal%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPersonal%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a>
                                <%end if
                                'if session("version")&"" = "5" then
                                    %><input class="width40"  type="text" name="nomresponsable"  disabled value="<%=EncodeForHtml(iif(mode="add",iif(TmpNombre & "">"",TmpNombre,""),iif(TmpNombre & "">"",TmpNombre,d_lookup("nombre","personal","dni='" & iif(TmpResponsable & "">"",TmpResponsable ,rst("responsable")) & "'",session("dsn_cliente")))))%>"><%
                                'end if
                                CloseDiv
                                %>
						<%end if
						'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:100px'","","",0,LitDescripcion +":"
						if mode="add" or mode="edit" then%>
								<%
                                if mode="add" then
                                    textdesc=""
                                else
                                    textdesc=iif(TmpDescripcion & "">"",TmpDescripcion,rst("descripcion"))
                                end if
								if textdesc="" then
									textdesc=LitTextTraspaso
								end if
                                if mode="add" then
                                    DrawDiv "1","",""
                                        DrawLabel "","",LitDescripcion%><textarea class='CELDA' <%=iif(mode="edit",esDisabled,"")%> name="descripcion" onFocus="lenmensaje(this,0,255,'')" onKeydown="lenmensaje(this,0,255,'')" onKeyup="lenmensaje(this,0,255,'')" onBlur="lenmensaje(this,0,255,'')" rows="3" cols="100"><%=EncodeForHtml(textdesc)%></textarea><%
                                    CloseDiv
                                elseif mode="edit" then
                                    DrawDiv "1", "", ""
                                        DrawLabel "","",LitDescripcion%><textarea class='CELDA' disabled name="descripcion" onFocus="lenmensaje(this,0,255,'')" onKeydown="lenmensaje(this,0,255,'')" onKeyup="lenmensaje(this,0,255,'')" onBlur="lenmensaje(this,0,255,'')" rows="3" cols="100"><%=EncodeForHtml(textdesc)%></textarea><%
                                    CloseDiv
                                end if
                                %>
						<%else
                            EligeCeldaResponsive "text",mode,"CELDA","",0,"",LitDescripcion,LitDescripcion,20,EncodeForHtml(pintar_saltos_espacios(iif(TmpDescripcion & "">"",TmpDescripcion,rst("descripcion"))))
							'DrawCelda "CELDA colspan='4'","","",0,pintar_saltos_espacios(iif(TmpDescripcion & "">"",TmpDescripcion,rst("descripcion")))
						end if
                        ''DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"
                        ''DrawCelda "CELDA","","",0,"&nbsp;"
                        ''DrawCelda "CELDA","","",0,"&nbsp;"
						'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:100px'","","",0,LitCajaOrg +":"
						campo="codigo"
						campo2="descripcion"
                        if mode="add" then
                            dato_celda=Desplegable(mode,campo,campo2,"cajas","","") 
                        else
							dato_celda=Desplegable(mode,campo,campo2,"cajas",iif(TmpCajaOrg>"",TmpCajaOrg,rst("cajaorg")),"")
                        end if
						if viene & "">"" then
                            'DrawSelectCelda(estilo,ancho,alto,tabulacion,etiqueta,name,reg,value,campo,campo2,evento,funcion)
                            'DrawSelectCelda "CELDA","","",0,LitCajaOrg,"cajaorg2",rst,iif(TmpCajaOrg>"",TmpCajaOrg,""),campo,campo2,"",""
                            if mode="browse" then
                                EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitCajaOrg,LitCajaOrg,EncodeForHtml(dato_celda)
                            elseif mode = "edit" then
							    EligeCelda "select", mode,"CELDA style='width:180px'" & iif(viene>""," disabled",""),iif(mode<>"browse","180",""),"",0,LitCajaOrg,"cajaorg2",15,EncodeForHtml(dato_celda)
                            elseif mode = "add" then
                                EligeCelda "select", mode,"CELDA style='width:180px'" & iif(viene>""," disabled",""),iif(mode<>"browse","180",""),"",0,LitCajaOrg,"cajaorg2",15,EncodeForHtml(dato_celda)
                            end if
							%><input type="hidden" name="cajaorg" value="<%=EncodeForHtml(TmpCajaOrg)%>"><%
						else
                            if mode="browse" then
                                EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitCajaOrg,LitCajaOrg,EncodeForHtml(dato_celda)
                            elseif mode = "edit" then
							    EligeCelda "select", mode,"CELDA " & iif(mode="edit",esDisabled2,"") & " style='width:180px'",iif(mode<>"browse","180",""),"",0,LitCajaOrg,"cajaorg",15,EncodeForHtml(dato_celda)
                            elseif mode = "add" then
                                EligeCelda "select", mode,"CELDA " & iif(mode="edit",esDisabled2,"") & " style='width:180px'",iif(mode<>"browse","180",""),"",0,LitCajaOrg,"cajaorg",15,EncodeForHtml(dato_celda)
                            end if
                            'DrawSelectCelda "CELDA","","",0,LitCajaOrg,"cajaorg",rst,iif(TmpCajaOrg>"",TmpCajaOrg,""),campo,campo2,"",""
						end if
						if mode="add" or mode="edit" then
                            if RstAux.state<>0 then RstAux.close
                        end if
						'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"
						'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:100px'","","",0,LitCajaDes +":"
						campo="codigo"
						campo2="descripcion"
                        if mode="add" then
                            dato_celda=Desplegable(mode,campo,campo2,"cajas","","") 
                        else
							dato_celda=Desplegable(mode,campo,campo2,"cajas",iif(TmpCajaDest>"",TmpCajaDest,rst("cajadest")),"")
                        end if
                        if mode="browse" then
                            EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitCajaDes,LitCajaDes,EncodeForHtml(dato_celda)
                        elseif mode="add" then
                            EligeCelda "select", mode,"CELDA " & iif(mode="edit",esDisabled2,""),iif(mode<>"browse","180",""),"",0,LitCajaDes,"cajadest",15,EncodeForHtml(dato_celda)
                        elseif mode="edit" then
						    EligeCelda "select", mode,"CELDA " & iif(mode="edit",esDisabled2,""),iif(mode<>"browse","180",""),"",0,LitCajaDes,"cajadest",15,EncodeForHtml(dato_celda)
                        end if
                        'DrawSelectCelda "CELDA","","",0,LitCajaDes,"cajadest",rst,iif(TmpCajaDest>"",TmpCajaDest,""),campo,campo2,"",""
						if mode="add" or mode="edit" then
                            if RstAux.state<>0 then RstAux.close
                        end if
						'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"

						'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:100px'","","",0,LitMedio +":"
                        if mode="add" then
                            defecto=""
                        else
                            defecto=iif(TmpMedio>"",TmpMedio,rst("medio"))
                        end if
                        if mode="add" or mode="edit" then
                             rstAux.cursorlocation=3
			                 rstAux.open "select codigo, descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
                             
                            if mode = "edit" then
                                DrawSelectCelda "CELDA " & iif(mode="edit",esDisabled2,""),iif(mode<>"browse","180",""),"",0,LitMedio,"medio",rstAux,defecto,"codigo","descripcion","",""
                            else
                                DrawSelectCelda "CELDA " & iif(mode="edit",esDisabled2,""),iif(mode<>"browse","180",""),"",0,LitMedio,"medio",rstAux,defecto,"codigo","descripcion","",""
                            end if
			                 rstAux.close
                         else
                            'DrawCelda "CELDA style='width:180px'","","",0,d_lookup("descripcion","tipo_pago","codigo='" & rst("medio") & "'",session("dsn_cliente"))
                                EligeCeldaResponsive "text",mode,"CELDA","",0,"",LitMedio,LitMedio,20,EncodeForHtml(d_lookup("descripcion","tipo_pago","codigo='" & rst("medio") & "'",session("dsn_cliente")))
                         end if

						if mode="add" or mode="edit" then
                            if RstAux.state<>0 then RstAux.close
                        end if
						'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"
						'DrawCelda iif(mode="browse","ENCABEZADOL","CELDA") & " style='width:100px'","","",0,LitImporte +":"
                        if mode="add" then
                            importe_aux2=miround(0, n_decimales)
                        else
							importe_aux2=miround(iif(TmpImporte & "">"",TmpImporte,rst("importe")), n_decimales)
                        end if
						if mode="add" or mode="edit" then
                            if mode="edit" then
								' Comprobamos si el traspaso tiene apuntes asociados, en cuyo caso no podrá modificarse el importe'
								'ega 11/06/2008 uso de command y contar el numero de caja para saber que celda dibujar
                                command.CommandText= "select count (caja) as numero from caja with(NOLOCK) where caja like '" & session("ncliente") & "%' and ntraspaso like '"+Session("ncliente")+"%' and ntraspaso='" & TmpNTraspaso & "'"
                                set rstAux2 = command.Execute()								    

								'rstAux2.open "select * from caja with(NOLOCK) ,session("dsn_cliente")
								'if rstAux2.eof then
								if rstAux2("numero")=0 then   
									%><input type="hidden" name="importeEsc" value="NO"><%
									'DrawInputCeldaAction "CELDA " & iif(mode="edit",esDisabled2,"") & " maxlength='10' ID='importe'","","",17,0,"","importe",iif(importe_aux2>"",importe_aux2,"0"),"onchange","cambiarimporte()",false
                                    DrawInputCeldaActionDiv "CELDA " & iif(mode="edit",esDisabled2,"") & " maxlength='10' ID='importe'","","",17,0,LitImporte,"importe",iif(importe_aux2>"",EncodeForHtml(importe_aux2),"0"),"onchange","cambiarimporte()",false
								else
									%><input type="hidden" name="importeEsc" value="SI"><%
									''DrawInputCeldaAction "CELDA disabled maxlength='10' ID='importe'","","",17,0,"","importe",iif(importe_aux2>"",importe_aux2,"0"),"onchange","cambiarimporte()",false
									 DrawInputCeldaDisabled "", "", "", 10, 0, LitImporte, "importe", EncodeForHtml(formatnumber(iif(importe_aux2>"",importe_aux2,"0"),n_decimales,-1,0,-1) & " " & a_breviatura)
								end if
								rstAux2.close
                            else
								%><input type="hidden" name="importeEsc" value="NO"><%
								'DrawInputCeldaAction "CELDA " & iif(mode="edit",esDisabled2,"") & " maxlength='10' ID='importe'","","",17,0,"","importe",iif(importe_aux2>"",importe_aux2,"0"),"onchange","cambiarimporte()",false
                                DrawInputCeldaActionDiv "CELDA " & iif(mode="edit",esDisabled2,"") & " maxlength='10' ID='importe'","","",17,0,LitImporte,"importe",iif(importe_aux2>"",EncodeForHtml(importe_aux2),"0"),"onchange","cambiarimporte()",false
                            end if
						else
							'DrawCelda2 "CELDA name='importe' ID='importe'", "left", false, formatnumber(iif(importe_aux2>"",importe_aux2,"0"),n_decimales,-1,0,-1) & " " & a_breviatura
                            EligeCeldaResponsive "text",mode,"CELDA","",0,"",LitImporte,LitImporte,"",EncodeForHtml(formatnumber(iif(importe_aux2>"",importe_aux2,"0"),n_decimales,-1,0,-1) & " " & a_breviatura)
						end if
                        'DrawCelda "CELDA style='width:10%'","","",0,"&nbsp;"
					%>
                    </div>
                    </div>
				    <%'Mostrar los Detalles del pedido.'
						    if mode="browse" then
							    if mode="browse" then
								    enCierre=d_lookup("cierre","caja","caja like '" & session("ncliente") & "%' and cierre is not null and (ntraspaso='" & rst("ntraspaso") & "' or (ndocumento='" & rst("ntraspaso") & "' and tdocumento='TRASPASO ENTRE CAJAS'))",session("dsn_cliente")) & ""
								    %><input type="hidden" name="encierre" value="<%=EncodeForHtml(trimCodEmpresa(enCierre))%>"><%
							    end if
							
							    set rstDet = Server.CreateObject("ADODB.Recordset")%>
                                <div class="Section" id="S_DETALLES">
                                    <a href="#" rel="toggle[DETALLES]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                                        <div class="SectionHeader">
                                            <%=LitDetalles%>
                                            <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                                        </div>
                                    </a>
                                <div class="SectionPanel" style="<%=iif(mode="add" or mode="edit","display: none","display: ")%>;" id="DETALLES">
							        <%''ricardo 26-2-2008 se comentan estas dos lineas, ya que no hacen nada.
							        ''rstDet.cursorlocation=3
							        ''rstDet.open "select * from caja with(NOLOCK) where caja like '" & session("ncliente") & "%' and ntraspaso like '"+Session("ncliente")+"%' and ntraspaso='" & rst("ntraspaso") & "' order by nanotacion",session("dsn_cliente")
							        if nz_b(rst("contabilizado"))=-1 then
								        ' No aparece el icono para añadir '%>
								        <table width="725" border='0' cellspacing="1" cellpadding="1">
								          <tr>
									        <td width="25%" bgcolor="<%=color_fondo%>">
										        <font class = "ENCABEZADOC"><%=LitTituloDet%></font>
				   					        </td>
				   				          </tr>
				   				        </table>
				   			        <%else
								        ' Aparece el icono para añadir '%>
								        <table width="725" border='0' cellspacing="1" cellpadding="1">
								          <tr>
									        <td width="25%" bgcolor="<%=color_fondo%>">
										        <!--<font class = "ENCABEZADOC"><%=LitTituloDet%></font>-->
										        <a href="javascript:AnyadirApuntes('<%=EncodeForHtml(TmpNTraspaso)%>')"><img src="../images/<%=ImgAnyadir%>" <%=ParamImgAnyadir%> alt="<%=LitAddApuntes%>"></a>
				   					        </td>
				   				          </tr>
				   				        </table>
				   			        <%end if%>
                                    <div class="overflowXauto">
							        <table class="width90 lg-table-responsive bCollapse">
							            <%DrawFila color_terra%>
								            <td class='ENCABEZADOL  underOrange width20'><%=LitFecha%></td>
								            <td class='ENCABEZADOL  underOrange width20'><%=LitDescripcion%></td>
								            <td class='ENCABEZADOL  underOrange width20'><%=LitDocumento%></td>
								            <td class='ENCABEZADOL  underOrange width20'><%=LitTipo%></td>
								            <td class='ENCABEZADOL  underOrange width15'><%=LitImporte%></td>
                                            <td class='ENCABEZADOL  underOrange width5'></td>
               				            <%CloseFila%>
							        </table>
							        <iframe class='width90 iframe-data lg-table-responsive' id='frDetalles' name="fr_Detalles" src='traspasos_cajadet.asp?ndoc=<%=EncodeForHtml(rst("ntraspaso"))%>' height="100px"  frameborder="yes" noresize="noresize"></iframe>
                                    
				                      <%if mode="browse" then
                                            abreviatura=d_lookup("abreviatura","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & rst("divisa") & "' ",session("dsn_cliente"))
					                        ndecimales=d_lookup("ndecimales","divisas","codigo like '" & session("ncliente") & "%' and codigo='" & rst("divisa") & "' ",session("dsn_cliente"))%>
				                              <table class="width90 lg-table-responsive bCollapse">
				  	                            <tr>
                                                    <td class='CELDAR7 width80'></td>
				  		                            <td class='CELDAL7 width10' id='litImporteTotal'><%=LitImporte%></td>
				  		                            <td class='CELDAL7 width10' id='importeTotal' style="background-color: transparent; border: 0px; text-align: right;"><%=EncodeForHtml(formatnumber(rst("importe"),ndecimales,-1,0,-1))%>&nbsp; <%=EncodeForHtml(abreviatura)%></td>
				  	                            </tr>
				                              </table>
				                      <%end if%>
                                    </div>
                                </div>
                                </div>
						    <%end if
				    if submode2="traerresponsable" then%>
					    <script language="javascript" type="text/javascript">
						    document.traspasos_caja.descripcion.focus();
						    document.traspasos_caja.descripcion.select();
					    </script>
				    <%end if
                    if mode="browse" then %>
    		            <script language="javascript" type="text/javascript">jQuery(window).load(function () { Redimensionar(); });</script>
                    <%end if
                end if
			end if
		elseif mode="search" then
        end if%>
	</form>
<%end if
set rst = nothing
set rst2 = nothing
set rstAux = nothing
set rstDet = nothing
set command=nothing
conn.close
set conn = Nothing
connRound.close
set connRound = Nothing
set rstAux2 = Nothing
%>
</body>
</html>
