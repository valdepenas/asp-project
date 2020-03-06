<%@ Language=VBScript %>
<%Server.ScriptTimeout = 1200

dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function

''ricardo 30-5-2003 se pone como personalizacion que salgan los descuentos para ciertas familias segun viene del parametro del menu
''ricardo 5-6-2003 se añade el parametro caju para que solo se pongan las cajas que se diga en el parametro
'JCI : 24/02/2005
'	se añade la gestión del parámetro de usuario i para mostrar los importes con iva
'EJM 10/10/2006 Se sustituye la operación importe + iva por la columna importeIVA 
'MPC 16/11/2007 CAMBIO DSN PARA LISTADOS
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
    <title><%=LitTituloRVT%></title>
    <meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>" />
    <link rel="stylesheet" href="../../pantalla.css" media="screen" />
    <link rel="stylesheet" href="../../impresora.css" media="print" />
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<% if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../tickets.inc" -->
<!--#include file="../../perso.inc" -->
<!--#include file="../../CatFamSubResponsive.inc"-->
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../ventas.inc" -->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->
<!--#include file="../../common/campospersoResponsive.inc" -->
<!--#include file="../../common/poner_cajaResponsive.inc" -->
<!--#include file="../../styles/formularios.css.inc"-->  

<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>

<script language="javascript" type="text/javascript">
    function AbrirElegir(modo){
        pagina="../../central.asp?pag1=ventas/listados/listado_tickets_elegir.asp&mode=add&ndoc=listado_tickets&viene=" + modo + "&pag2=ventas/listados/listado_tickets_elegir_bt.asp";
        ven=AbrirVentana(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function poner_alt(ver,texto){
        if (ver==1){
            //se vera el texto
            document.getElementById("ver_alt").innerHTML=unescape(texto);
            var x=event.clientX+15;
            var y=event.clientY+15;
            document.getElementById("ver_alt").style.left=x;
            document.getElementById("ver_alt").style.top=y;
            document.getElementById("ver_alt").style.display="";
        }
        else{
            //no se vera
            document.getElementById("ver_alt").style.display="none";
        }
    }

    function TraerOperador() {
        document.listado_tickets.action="listado_tickets.asp?operador=" + document.listado_tickets.operador.value + "&mode=traeroperador";
        document.listado_tickets.submit();
    }

    function ver_tours(){
        fdesde=document.listado_tickets.dfecha.value;
        fhasta=document.listado_tickets.hfecha.value;
        operador=document.listado_tickets.operador.value;
        pagina="../../central.asp?pag1=mantenimiento/listados/tours.asp&mode=imp&ndoc=" + fdesde + "&ndocumento=" + fhasta + "&tdocumento=" + operador + "&viene=listado_tickets&pag2=mantenimiento/listados/tours_bt.asp";
        ven=AbrirVentana(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function ver_penalizaciones(){
        fdesde=document.listado_tickets.dfecha.value;
        fhasta=document.listado_tickets.hfecha.value;
        operador=document.listado_tickets.operador.value;
        pagina="../../central.asp?pag1=administracion/listados/listado_penalizaciones.asp&mode=imp&ndoc=" + fdesde + "&ndocumento=" + fhasta + "&tdocumento=" + operador + "&viene=listado_tickets&pag2=administracion/listados/listado_penalizaciones_bt.asp";
        ven=AbrirVentana(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function ver_comisiones(){
        fdesde=document.listado_tickets.dfecha.value;
        fhasta=document.listado_tickets.hfecha.value;
        pagina="../../central.asp?pag1=ventas/listados/descuentos_docs_cli.asp&mode=browse&ndoc=" + fdesde + "&ndocumento=" + fhasta + "&viene=listado_tickets&pag2=ventas/listados/descuentos_docs_cli_bt.asp";
        ven=AbrirVentana(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }

    function paramListado(listado){

        valor_i=document.listado_tickets.i.value;
        if (listado=="tickets"){
            document.location.href="tickets.asp?mode=param&i=" +  valor_i;
            parent.botones.document.location="tickets_bt.asp?mode=param"
        }
        else{
            document.location.href="listado_tickets.asp?mode=param&i=" +  valor_i;
            parent.botones.document.location="listado_tickets_bt.asp?mode=param"
        }
	
    }

    function ValidarFecha(num){
        // Se entiende que si el usuario deja la hora vacía, quiere decir que no tiene límite, es decir, será 00:00 o 23:59
        switch (num){
            case 1:
                if (document.listado_tickets.dhora.value != "")
                {
                    if (document.listado_tickets.dhora.value.length == 4){
                        var h1 = document.listado_tickets.dhora.value.substring(0,2);
                        var m1 = document.listado_tickets.dhora.value.substring(2);
                        document.listado_tickets.dhora.value = h1 +":"+ m1;
                        if (!checkhora(document.listado_tickets.dhora))
                        {
                            alert("<%=LitMsgHoraNumerico%>");
                            document.listado_tickets.dhora.value="";
                            return false;		
                        }
                    }
                    else if (document.listado_tickets.dhora.value.length == 3){
                        document.listado_tickets.dhora.value = "0" + document.listado_tickets.dhora.value;
                        ValidarFecha(1);    
                    }
                    else {
                        if (!checkhora(document.listado_tickets.dhora))
                        {
                            alert("<%=LitMsgHoraNumerico%>");
                            document.listado_tickets.dhora.value="";
                            return false;		
                        }
                    }
		        
                }
                else
                {
                    document.listado_tickets.dhora.value = "00:00";
                }
                break;
            case 2:
                if (document.listado_tickets.hhora.value != "")
                {
                    if (document.listado_tickets.hhora.value.length==4)
                    {
                        var h2 = document.listado_tickets.hhora.value.substring(0,2);
                        var m2 = document.listado_tickets.hhora.value.substring(2);
                        document.listado_tickets.hhora.value = h2 +":"+ m2;
                        if (!checkhora(document.listado_tickets.hhora))
                        {
                            alert("<%=LitMsgHoraNumerico%>");
                            document.listado_tickets.hhora.value ="";
                            document.listado_tickets.hhora.focus();
                            return false;		
                        }
                    }
                    else if (document.listado_tickets.hhora.value.length==3)
                    {
                        document.listado_tickets.hhora.value = "0" + document.listado_tickets.hhora.value;
                        ValidarFecha(2);
                    }
                    else
                    {
                        if (!checkhora(document.listado_tickets.hhora))
                        {
                            alert("<%=LitMsgHoraNumerico%>");
                            document.listado_tickets.hhora.value ="";
                            document.listado_tickets.hhora.focus();
                            return false;		
                        }
                    }
                }
                else
                {
                    document.listado_tickets.hhora.value = "23:59";
                }
                break;   
        }
    }
</script>
<body onload="self.status='';" class="BODY_ASP">
<iframe name="frameExport" style='display:none;' src="listado_tickets_pdf.asp?mode=ver" frameborder='0' width='500' height='200'></iframe>
<%
sub CreatablaTemporalIva

        strdrop ="if exists (select * from sysobjects where id = object_id('[" & session("usuario") & "_desglose_iva]') and sysstat " & _
		" & 0xf = 3) drop table [" & session("usuario") & "_desglose_iva]"
		rstAux.open strdrop,Session("backendlistados"),adUseClient,adLockReadOnly
		
        if rstAux.state<>0 then rstAux.close

		strDesgloseIva="CREATE TABLE [" & session("usuario") & "_desglose_iva] (num int identity(1,1)"
		strDesgloseIva=strDesgloseIva & ",referencia varchar(30),total_cant real,total_imp money,total_imp_iva money,divisa varchar(15)"

		select case mostrarcolum
			case ucase(LitTicketTpv):
		        strDesgloseIva=strDesgloseIva & ",tpv varchar(8),nombretpv varchar(50)"
			case ucase(LitOperador):
		        strDesgloseIva=strDesgloseIva & ",usuario varchar(20),nompersonal varchar(50)"
			case ucase(LitCaja):
		        strDesgloseIva=strDesgloseIva & ",caja varchar(10),nombrecaja varchar(50)"
			case ucase(LitTicketTienda):
		        strDesgloseIva=strDesgloseIva & ",tienda varchar(10),nombretienda varchar(50)"
        end select

		strDesgloseIva=strDesgloseIva & ",nomarticulo varchar(100),familia varchar(10),nomfamilia varchar(50),fampadre varchar(10),art_cod_categoria varchar(50)"

		strDesgloseIva=strDesgloseIva & ", porcentaje_iva real, importe_base_total money , factorcambioOrigen float)"

		rstAux.open strDesgloseIva,Session("backendlistados"),adUseClient,adLockReadOnly

		if rstAux.State<>0 then rstAux.close
		GrantUser session("usuario"), Session("backendlistados")

        strDesgloseIvaInserta =""
		strDesgloseIvaInserta=" insert into [" & session("usuario") & "_desglose_iva](referencia,total_cant,total_imp,total_imp_iva,divisa " 

		select case mostrarcolum
			case ucase(LitTicketTpv):
		        strDesgloseIvaInserta=strDesgloseIvaInserta & ",tpv,nombretpv"
			case ucase(LitOperador):
		        strDesgloseIvaInserta=strDesgloseIvaInserta & ",usuario,nompersonal"
			case ucase(LitCaja):
		        strDesgloseIvaInserta=strDesgloseIvaInserta & ",caja,nombrecaja"
			case ucase(LitTicketTienda):
		        strDesgloseIvaInserta=strDesgloseIvaInserta & ",tienda,nombretienda"
        end select

		strDesgloseIvaInserta = strDesgloseIvaInserta & ",nomarticulo,familia,nomfamilia,fampadre,art_cod_categoria "

		strDesgloseIvaInserta = strDesgloseIvaInserta & ", porcentaje_iva , importe_base_total , factorcambioOrigen) "
   
		strDesgloseIvaInserta = strDesgloseIvaInserta & strSelectIva & strfrom & " CROSS join DIVISAS as div with (NOLOCK) " 
        strDesgloseIvaInserta = strDesgloseIvaInserta & strwhere & " AND div.CODIGO = ti.DIVISA AND div.CODIGO like '" & session("ncliente") & "%' " 
        strDesgloseIvaInserta = strDesgloseIvaInserta & strGroupIva & " ,div.FACTCAMBIO "
             
        rstDesgloseIva.open strDesgloseIvaInserta,Session("backendlistados")
end sub

sub calculaIva()    
    ''DBG 19/07/2016 Version 5 de muestra totales de IVA
    ''Utilizo el dato importe o importeIVA segun la configuración del servidor si coniva esta activado o no
    if coniva="SI" then 
        strIVAMuestraDesglose = " SELECT (tmp.porcentaje_iva) AS porcentajeiva "

        strIVAMuestraDesglose = strIVAMuestraDesglose  & " ,CASE MAX(tmp.divisa) WHEN '" & mb & "' THEN SUM(tmp.importe_base_total )"
	    strIVAMuestraDesglose = strIVAMuestraDesglose  & " ELSE SUM((tmp.importe_base_total/(tmp.factorcambioOrigen))*(" & mbFactorCambioDestino & "))   END AS baseimponible "

	    strIVAMuestraDesglose = strIVAMuestraDesglose  & " ,CASE MAX(tmp.divisa) WHEN '"& mb & "' THEN SUM( tmp.total_imp_iva - tmp.importe_base_total) "
	    strIVAMuestraDesglose = strIVAMuestraDesglose  & " ELSE SUM(( (tmp.total_imp_iva/tmp.factorcambioOrigen)*(" & mbFactorCambioDestino & ") )  - (tmp.importe_base_total/tmp.factorcambioOrigen)*(" & mbFactorCambioDestino & "))  END AS importeiva "
	
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " ,CASE MAX(tmp.divisa) WHEN '" & mb &  "' THEN SUM( total_imp_iva ) "
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " ELSE SUM((tmp.total_imp_iva/tmp.factorcambioOrigen)*(" & mbFactorCambioDestino & "))  END AS total "
    	
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " FROM [" & session("usuario") & "_desglose_iva] as tmp "
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " GROUP BY tmp.porcentaje_iva"
    else
        strIVAMuestraDesglose = " SELECT (tmp.porcentaje_iva) AS porcentajeiva "

        strIVAMuestraDesglose = strIVAMuestraDesglose  & " ,CASE MAX(tmp.divisa) WHEN '" & mb & "' THEN SUM(tmp.importe_base_total )"
	    strIVAMuestraDesglose = strIVAMuestraDesglose  & " ELSE SUM((tmp.importe_base_total/(tmp.factorcambioOrigen ))*(" & mbFactorCambioDestino & "))   END AS baseimponible "

	    strIVAMuestraDesglose = strIVAMuestraDesglose  & " ,CASE MAX(tmp.divisa) WHEN '"& mb & "' THEN SUM( tmp.total_imp_iva - tmp.importe_base_total) "
	    strIVAMuestraDesglose = strIVAMuestraDesglose  & " ELSE SUM(( (tmp.total_imp_iva/tmp.factorcambioOrigen)*(" & mbFactorCambioDestino & ") )  - (tmp.importe_base_total/tmp.factorcambioOrigen)*(" & mbFactorCambioDestino & "))  END AS importeiva "
	
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " ,CASE MAX(tmp.divisa) WHEN '" & mb &  "' THEN SUM( total_imp_iva ) "
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " ELSE SUM((tmp.total_imp_iva/tmp.factorcambioOrigen)*(" & mbFactorCambioDestino & "))  END AS total "
    	
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " FROM [" & session("usuario") & "_desglose_iva] as tmp "
        strIVAMuestraDesglose = strIVAMuestraDesglose  & " GROUP BY tmp.porcentaje_iva"
    end if

    ''Imprimo los resultados en una tabla
    ''Imprimo los encabezados de la tabla Desglose IVA
   %>
   <tr>
        <td>&nbsp;</td>
   </tr>
<%
    DrawFila color_fondo
    DrawCelda "DATO colspan='4' align='left' style='width:150px'","","",0,"<b>" & LitDesgloseSegunIva & "</b>" 
    CloseFila
    DrawFila color_fondo
    DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & LitPorcIva & "</b>" 
    DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & LitBaseImp & "</b>"
    DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & LitImporteIva1 & "</b>"
    DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & LitTotal & "</b>"
    CloseFila
    
    rstDesgloseIvaRecorreBucle.cursorlocation = 3	

    ' Totales
    baseimponible = 0
    importeiva = 0
	total = 0
    fila = 0

    '' Lanzo la consulta aqui
    rstDesgloseIvaRecorreBucle.open strIVAMuestraDesglose,Session("backendlistados")

    while not rstDesgloseIvaRecorreBucle.eof
        fila=fila+1
        if ((fila+1) mod 2)=0 then
			color=color_blau
		else
			color=color_terra
		end if

		' Calculo los totales  --> 16/06/2016 DBG Calculo de los totales del desglose de IVA
		baseimponible = baseimponible + rstDesgloseIvaRecorreBucle("baseimponible")
		importeiva = importeiva + rstDesgloseIvaRecorreBucle("importeiva")
		total = total + rstDesgloseIvaRecorreBucle("total")

        ''DBG 1/07/2016 Imprime filas de la tabla Desglose IVA  
        ''DBG 1/07/2016 dec_cant es el literal hace que me muestre hasta 2 decimales en las instrucciones formatnumber
        ''DBG 19/07/2016 Cambio el orden de impresion de los campos segun este activo o no el valor del campo coniva

        DrawFila color
        DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & rstDesgloseIvaRecorreBucle("porcentajeiva") & "%" & "</b>"
        DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & formatnumber(rstDesgloseIvaRecorreBucle("baseimponible"),dec_mb,-1,0,-1) & " " & abreviatura & "</b>"
        DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & formatnumber(rstDesgloseIvaRecorreBucle("importeiva"),dec_mb,-1,0,-1) & " " & abreviatura & "</b>"
        DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & formatnumber(rstDesgloseIvaRecorreBucle("total"),dec_mb,-1,0,-1) & " " & abreviatura & "</b>"
        CloseFila

        rstDesgloseIvaRecorreBucle.movenext
    wend

	rstDesgloseIvaRecorreBucle.close
    %>
	<!-- Fila de totales de desglose de Iva -->
    <%
    DrawFila color_fondo
    DrawCelda "DATO align='left' style='width:150px'","","",0,"<b> Totales </b>"
    DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & formatnumber(baseimponible, dec_mb, -1, 0, -1) & " " & abreviatura & "</b>"
    DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & formatnumber(importeiva, dec_mb, -1, 0, -1) & " " & abreviatura & "</b>"
    DrawCelda "DATO align='right' style='width:150px'","","",0,"<b>" & formatnumber(total, dec_mb, -1, 0, -1) & " " & abreviatura & "</b>"
    CloseFila
%>
    <tr>
        <td>&nbsp;</td>
    </tr>
    <% ''01/07/2016 DBG: Fin de Desglose IVA %>      
<%
end sub

function Formatear(texto)
	if texto & "">"" then
		texto=server.urlencode(texto)
		texto=replace(texto,"+","%20")
	end if
	Formatear=texto
end function

function Escribirhref(ndoc,tdoc,texto_a_poner)
	if tdoc="operador" then
		url=Hiperv(OBJPersonal,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(texto_a_poner),LitVerPersonal)
	elseif tdoc="tienda" then
		url=Hiperv(OBJTiendas,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(texto_a_poner),LitVerTienda)
	elseif tdoc="articulo" then
		url=Hiperv(OBJArticulos,ndoc,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),texto_a_poner,LitVerArticulo)
	else
		url=trimCodEmpresa(ndoc)
	end if

	Escribirhref=url
end function

sub CabeceraListado(lista,contador,tdocumento,mostrarcolum,mostrardesc,mostrarinfo)
	if mostrardesc="on" then
		DrawFila color_fondo
			DrawCelda "DATO align='left' style='width:150px'","","",0,"<b>" & mostrarcolum & "</b>"
			DrawCelda "DATO align='left' style='width:350px'","","",0,"<b>" & ucase(LitNombreListTpv) & "</b>"
			DrawCelda "CELDA align='left' style='background-color:" & color_blau & "'","","",0,"&nbsp;"
		CloseFila
		fila=1
		while fila<=contador
			if ((fila+1) mod 2)=0 then
				color=color_blau
			else
				color=color_terra
			end if
			DrawFila color
				DrawCelda "DATO align='left'","","",0,"<b>" & Escribirhref(lista(fila,1),tdocumento,lista(fila,1)) & "</b>"
				DrawCelda "DATO align='left'","","",0,"<b>" & lista(fila,2) & "</b>"
				DrawCelda "CELDA align='left' style='background-color:" & color_blau & "'","","",0,"&nbsp;"
				fila=fila+1
			CloseFila
		wend
		DrawFila ""
		CloseFila
		DrawFila ""
		CloseFila
		%></table>
		<table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
	end if

    ancho=65/contador
    if ancho<clng(ancho) then
    	ancho=clng(ancho)-1
    else
	    ancho=clng(ancho)
    end if

	DrawFila color_fondo
		DrawCelda "DATO width='15%'","","",0,"&nbsp;"
		fila=1
		while fila<=contador
			%><td class="dato" width='<%=ancho%>%' align='right' onmouseover=poner_alt(1,'<%=Formatear(lista(fila,2))%>') onmouseout=poner_alt(0,'')>
				<b>
				<%=Escribirhref(lista(fila,1),tdocumento,lista(fila,1))%>
				<%if mostrarinfo=ucase(LitImporte) then%>
					<%="(" & enc.EncodeForHtmlAttribute(null_s(abrev_mb)) & ")"%>
				<%end if%>
				</b>
			</td><%
			fila=fila+1
		wend

		DrawCelda "DATO align='right' width='10%'","","",0,"<b>" & LitTotalCant & "</b>"
		DrawCelda "DATO align='right' width='10%'","","",0,"<b>" & LitTotalImpte & "(" & enc.EncodeForHtmlAttribute(null_s(abrev_mb)) & ")</b>"
	CloseFila
end sub

sub DibujaLineaSubTotales(lista_columnas,contador,cadena,familia,columna_a_mostrar,mostrarinfo,fampadre,categoria,opcion)
    ''hacemos una copia en otra variable para no modificar la que ya existe
    dim lista
    lista=lista_columnas

    dim lista2()
    contador_lista2=0

    dim conjunto
    dim cond

    if opcion & ""="categoria" then
        literaltot="Totales de la categoría"
        cond="tmp.art_cod_categoria"
        if categoria & ""="" then
            filtro=" is null"
        else
            filtro="='" & categoria & "'"
        end if
        cad_aux=" and "& cond & filtro

        cadena4=",tmp.art_cod_categoria,"
        cadena5=",tmp.art_cod_categoria,tmp.divisa"
    end if

    if opcion & ""="fampadre" then
        literaltot="Totales de la familia"
        cond="tmp.fampadre"
        if fampadre & ""="" then
            filtro=" is null"
        else
            filtro="='" & fampadre & "'"
        end if
        cad_aux=" and "& cond & filtro

        cadena4=",tmp.art_cod_categoria,tmp.fampadre,"
        cadena5=",tmp.art_cod_categoria,tmp.fampadre,tmp.divisa"
    end if

    if opcion & ""="familia" then
        literaltot="Totales de la subfamilia"
        cond="tmp.familia"
        if familia & ""="" then
            filtro=" is null"
        else
            filtro="='" & familia & "'"
        end if
        cad_aux=" and "& cond & filtro

        cadena4=",tmp.art_cod_categoria,tmp.fampadre,tmp.familia,"
        cadena5=",tmp.art_cod_categoria,tmp.fampadre,tmp.familia,tmp.divisa"
    end if

    if opcion & ""="" then
        literaltot="Totales sin categoría/familia/subfamilia"
        cad_aux= " and tmp.familia is null and tmp.fampadre is null and tmp.art_cod_categoria is null"
        cadena4=",tmp.art_cod_categoria,tmp.fampadre,tmp.familia,"
        cadena5=",tmp.art_cod_categoria,tmp.fampadre,tmp.familia,tmp.divisa"
    end if

	if coniva="SI" then 
		cadena2="sum(tmp.total_cant) as total_cant,sum(round(tmp.total_imp," & dec_mb &")) as total_imp,tmp.divisa "		
		cadena4=cadena4 &"tmp.divisa,sum(tmp.total_cant) as cantidad,sum(round(tmp.total_imp," & dec_mb &")) as importe"		
	else
		cadena2="sum(tmp.total_cant) as total_cant,sum(tmp.total_imp) as total_imp,tmp.divisa "
		cadena4=cadena4 &"tmp.divisa,sum(tmp.total_cant) as cantidad,sum(tmp.total_imp) as importe"
	end if

    cadena2=cadena2 & cadena 
	cadena4=cadena4 & cadena 

    cadena4=cadena4 & " and tmp.familia in " & poner_comillas(dtofam)
    
    cadena4=cadena4 & cad_aux
    cadena2=cadena2 & cad_aux & " group by tmp.divisa"           

	select case columna_a_mostrar
		case ucase(LitTicketTpv):
			cadena2="select tmp.tpv," & cadena2 & " ,tmp.tpv"
			cadena4="select distinct tmp.tpv " & cadena4 & " group by tmp.tpv" & cadena5 & " order by tmp.tpv" &cadena5
			cadena_a_comparar2="tpv"
		case ucase(LitOperador):
			cadena2="select tmp.usuario," & cadena2 & " ,tmp.usuario"
			cadena4="select distinct tmp.usuario " & cadena4 & " group by tmp.usuario" & cadena5 & " order by tmp.usuario"&cadena5
			cadena_a_comparar2="usuario"
		case ucase(LitCaja):
			cadena2="select tmp.caja," & cadena2 & " ,tmp.caja"
			cadena4="select distinct tmp.caja " & cadena4 & " group by tmp.caja" & cadena5 & " order by tmp.caja,"
			cadena_a_comparar2="caja"
		case ucase(LitTicketTienda):
			cadena2="select tmp.tienda," & cadena2 & " ,tmp.tienda"
			cadena4="select distinct tmp.tienda " & cadena4 & " group by tmp.tienda" & cadena5 & " order by tmp.tienda"&cadena5
			cadena_a_comparar2="tienda"
	end select

	''calculamos ahora las referencias que hay
	cadena3="select count(distinct upper(tmp.referencia)) as regtotales " & cadena
    cadena3=cadena3 & cad_aux

	rstcursor.cursorlocation=3
	rstcursor.open cadena3, Session("backendlistados")
	if not rstcursor.eof then
		regtotales=rstcursor("regtotales")
	else
		regtotales=0
	end if
	rstcursor.close

	si_dto=0
	if dtofam & "">"" then
		'miramos si hay algun articulo que pertenezca a la familia del dtofam
		rstcursor.cursorlocation=3
		rstcursor.open cadena4, Session("backendlistados")
		if not rstcursor.eof then
			redim preserve lista2(rstcursor.recordcount+2,3)
			si_dto=1
			contador_lista2=1
			campo_ant=rstcursor(cadena_a_comparar2)
			while not rstcursor.eof
					lista2(contador_lista2,1)=rstcursor(cadena_a_comparar2)
					dto=(miround(rstcursor("cantidad"),dec_cant)*precio_ticket)-cambiodivisa(rstcursor("importe"),rstcursor("divisa"),mb)
					if lista2(contador_lista2,2) & "">"" then
						lista2(contador_lista2,2)=lista2(contador_lista2,2) + dto
					else
						lista2(contador_lista2,2)=dto
					end if

				campo_ant=rstcursor(cadena_a_comparar2)
				rstcursor.movenext
				if not rstcursor.eof then
					if (campo_ant & "")=(rstcursor(cadena_a_comparar2) & "") then
					else
						contador_lista2=contador_lista2+1
					end if
				end if
			wend
		else
			si_dto=0
		end if
		rstcursor.close
	end if
	rstcursor.cursorlocation=3

	rstcursor.open cadena2, Session("backendlistados")
	if not rstcursor.eof then
		for i=1 to contador_columnas
			lista(i,3)=0
			lista(i,4)=0
		next
		sumatotalimp=0
		sumatotalcant=0
		while not rstcursor.eof
			if rstcursor("total_cant") & "">"" then
				cantidad_a_sumar=rstcursor("total_cant")
			else
				cantidad_a_sumar=0
			end if
			if rstcursor("total_imp") & "">"" then
				importe_a_sumar=cambiodivisa(rstcursor("total_imp"),rstcursor("divisa"),mb)
			else
				importe_a_sumar=0
			end if
			if importe_a_sumar & "">"" and cantidad_a_sumar & "">"" then
				dto_a_sumar=(cantidad_a_sumar*precio_ticket)-importe_a_sumar
			end if
			if cantidad_a_sumar<>0 or importe_a_sumar<>0 then
				for i=1 to contador_columnas
					if lista(i,1) & ""=(rstcursor(cadena_a_comparar2)) & "" then
						lista(i,3)=lista(i,3) + cantidad_a_sumar
						lista(i,4)=lista(i,4) + importe_a_sumar
					end if
				next
			end if
			rstcursor.movenext
		wend
	end if
	rstcursor.close

	dto_total_a_poner=0

	DrawFila color_fondo
		DrawCelda2 "dato align='left' width='10%' ", "left", false,"<b>" & literaltot & " : " & regtotales & "</b>"
		sumatotalimp=0
		sumatotalcant=0
		for i=1 to contador_columnas
			if lista(i,3) & "">"" then
				if lista(i,3)<>0 then
					cantidad_a_poner=formatnumber(lista(i,3),dec_cant,-1,0,-1)
					sumatotalcant=sumatotalcant + cantidad_a_poner
				else
					cantidad_a_poner=""
					sumatotalcant=sumatotalcant + 0
				end if
			else
				cantidad_a_poner=""
				sumatotalcant=sumatotalcant + 0
			end if
			if lista(i,4) & "">"" then
				if lista(i,4)<>0 then
					importe_a_poner=formatnumber(lista(i,4),dec_mb,-1,0,-1)
					sumatotalimp=sumatotalimp + importe_a_poner
				else
					importe_a_poner=""
					sumatotalimp=sumatotalimp + 0
				end if
			else
				importe_a_poner=""
				sumatotalimp=sumatotalimp + 0
			end if
			if importe_a_poner & "">"" and cantidad_a_poner & "">"" then
				dto_a_poner=(cantidad_a_poner*precio_ticket)-importe_a_poner
				dto_a_poner=formatnumber(dto_a_poner,decpor,-1,0,-1)
			else
				dto_a_poner=0
			end if

			if mostrarinfo=ucase(LitCantidad) then
				DrawCelda2 "dato align='right' ", "right", false,"<b>" & cantidad_a_poner & "</b>"
			else
				texto_a_poner=importe_a_poner
				if si_dto=1 then
					if dto_a_poner & "">"" and dto_a_poner<>0 then
						j=1
						si_poner=0
						while si_poner=0 and j<=contador_lista2
							if lista(i,1) & ""=lista2(j,1) & "" then
								si_poner=1
								''si queremos poner el dto total de la columna se deja como esta
								'' y si queremos solo los de las familias de la variable dtofam
								dto_a_poner=formatnumber(lista2(j,2),decpor,-1,0,-1)
								dto_total_a_poner=dto_total_a_poner + dto_a_poner
							end if
							j=j+1
						wend
						if si_poner=1 then
							texto_a_poner=texto_a_poner & "<br/>" & LitDescuento & " : " & dto_a_poner
						end if
					end if
				end if
				DrawCelda2 "dato align='right' ", "right", false,"<b>" & texto_a_poner & "</b>"
			end if
		next
		if sumatotalcant & "">"" then
			sumatotalcant=formatnumber(sumatotalcant,dec_cant,-1,0,-1)
		end if
		DrawCelda2 "dato align='right' width='10%' ", "right", false,"<b>" & sumatotalcant & "</b>"
		if sumatotalimp & "">"" then
			sumatotalimp=formatnumber(sumatotalimp,dec_mb,-1,0,-1)
			sumadto=(sumatotalcant*precio_ticket)-sumatotalimp
			dto_a_poner=formatnumber(sumadto,decpor,-1,0,-1)
		end if
		texto_a_poner=sumatotalimp
		if si_dto=1 then
			'si queremos poner el dto total de la columna se deja como esta
			' y si queremos solo los de las familias de la variable dtofam
			dto_a_poner=dto_total_a_poner
			if dto_a_poner & "">"" and dto_a_poner<>0 then
				texto_a_poner=texto_a_poner & "<br/>" & LitDescuento & " : " & dto_a_poner
			end if
		end if
		DrawCelda2 "dato align='right' width='10%' ", "right", false,"<b>" & texto_a_poner & "</b>"
	CloseFila

end sub

sub DibujaLineaTotales(lista,contador,cadena,cambiar_literal,columna_a_mostrar,mostrarinfo)
    dim lista2()
    contador_lista2=0

	if coniva="SI" then
		cadena2="sum(tmp.total_cant) as total_cant,sum(round(tmp.total_imp," & dec_mb & ")) as total_imp,tmp.divisa " & cadena
		cadena4=",tmp.familia,tmp.divisa,sum(tmp.total_cant) as cantidad,sum(round(tmp.total_imp," & dec_mb &")) as importe"
	else
		cadena2="sum(tmp.total_cant) as total_cant,sum(tmp.total_imp) as total_imp,tmp.divisa " & cadena
		cadena4=",tmp.familia,tmp.divisa,sum(tmp.total_cant) as cantidad,sum(tmp.total_imp) as importe"
	end if
	cadena5=",tmp.familia,tmp.divisa"
	cadena4=cadena4 & cadena & " and tmp.familia in " & poner_comillas(dtofam)

	select case columna_a_mostrar
		case ucase(LitTicketTpv):
			cadena2="select tmp.tpv," & cadena2 & " group by tmp.divisa,tmp.tpv"
			cadena4="select distinct tmp.tpv " & cadena4 & " group by tmp.tpv" & cadena5 & " order by tmp.tpv,tmp.familia,tmp.divisa"
			cadena_a_comparar2="tpv"
		case ucase(LitOperador):
			cadena2="select tmp.usuario," & cadena2 & " group by tmp.divisa,tmp.usuario"
			cadena4="select distinct tmp.usuario " & cadena4 & " group by tmp.usuario" & cadena5 & " order by tmp.usuario,tmp.familia,tmp.divisa"
			cadena_a_comparar2="usuario"
		case ucase(LitCaja):
			cadena2="select tmp.caja," & cadena2 & " group by tmp.divisa,tmp.caja"
			cadena4="select distinct tmp.caja " & cadena4 & " group by tmp.caja" & cadena5 & " order by tmp.caja,tmp.familia,tmp.divisa"
			cadena_a_comparar2="caja"
		case ucase(LitTicketTienda):
			cadena2="select tmp.tienda," & cadena2 & " group by tmp.divisa,tmp.tienda"
			cadena4="select distinct tmp.tienda " & cadena4 & " group by tmp.tienda" & cadena5 & " order by tmp.tienda,tmp.familia,tmp.divisa"
			cadena_a_comparar2="tienda"
	end select

	'calculamos ahora las referencias que hay
	cadena3="select count(distinct upper(tmp.referencia)) as regtotales " & cadena
	rstcursor.cursorlocation=3
	rstcursor.open cadena3, Session("backendlistados")
	if not rstcursor.eof then
		regtotales=rstcursor("regtotales")
	else
		regtotales=0
	end if
	rstcursor.close

	si_dto=0
	if dtofam & "">"" then
		'miramos si hay algun articulo que pertenezca a la familia del dtofam
		rstcursor.cursorlocation=3
		rstcursor.open cadena4, Session("backendlistados")
		if not rstcursor.eof then
			redim preserve lista2(rstcursor.recordcount+2,3)
			si_dto=1
			contador_lista2=1
			campo_ant=rstcursor(cadena_a_comparar2)
			while not rstcursor.eof
					lista2(contador_lista2,1)=rstcursor(cadena_a_comparar2)
					dto=(miround(rstcursor("cantidad"),dec_cant)*precio_ticket)-cambiodivisa(rstcursor("importe"),rstcursor("divisa"),mb)					
					if lista2(contador_lista2,2) & "">"" then
						lista2(contador_lista2,2)=lista2(contador_lista2,2) + dto
					else
						lista2(contador_lista2,2)=dto
					end if

				campo_ant=rstcursor(cadena_a_comparar2)
				rstcursor.movenext
				if not rstcursor.eof then
					if (campo_ant & "")=(rstcursor(cadena_a_comparar2) & "") then
					else
						contador_lista2=contador_lista2+1
					end if
				end if
			wend
		else
			si_dto=0
		end if
		rstcursor.close
	end if

	'calculamos los totales
	rstcursor.cursorlocation=3
	rstcursor.open cadena2, Session("backendlistados")
	if not rstcursor.eof then
		for i=1 to contador_columnas
			lista(i,3)=0
			lista(i,4)=0
		next
		sumatotalimp=0
		sumatotalcant=0
		while not rstcursor.eof
			if rstcursor("total_cant") & "">"" then
				cantidad_a_sumar=rstcursor("total_cant")
			else
				cantidad_a_sumar=0
			end if
			if rstcursor("total_imp") & "">"" then
				importe_a_sumar=cambiodivisa(rstcursor("total_imp"),rstcursor("divisa"),mb)
			else
				importe_a_sumar=0
			end if
			if importe_a_sumar & "">"" and cantidad_a_sumar & "">"" then
				dto_a_sumar=(cantidad_a_sumar*precio_ticket)-importe_a_sumar
			end if
			if cantidad_a_sumar<>0 or importe_a_sumar<>0 then
				for i=1 to contador_columnas
					if lista(i,1) & ""=(rstcursor(cadena_a_comparar2)) & "" then
						lista(i,3)=lista(i,3) + cantidad_a_sumar
						lista(i,4)=lista(i,4) + importe_a_sumar
					end if
				next
			end if
			rstcursor.movenext
		wend
	end if
	rstcursor.close

	if cambiar_literal=0 then
		cadena_total=LitTotales
	else
		cadena_total=LitTotales3
		DrawFila ""
		CloseFila
		DrawFila ""
		CloseFila
	end if

	dto_total_a_poner=0

    mostrarinfo2=mostrarinfo
    for colum=1 to 2
	    if colum=1 then
		    mostrarinfo=mostrarinfo2
	    else
		    if mostrarinfo2=ucase(LitCantidad) then
			    mostrarinfo=ucase(LitImporte)
		    else
			    mostrarinfo=ucase(LitCantidad)
		    end if
	    end if
	    if mostrarinfo=ucase(LitImporte) then
		    texto_total=LitTotalImpte & "(" & enc.EncodeForHtmlAttribute(null_s(abrev_mb)) & ") (" & regtotales & " " & LitRegTot & ")"
	    else
		    texto_total=LitTotalCant & " (" & regtotales & " " & LitRegTot & ")"
	    end if
	    DrawFila color_fondo
		    DrawCelda2 "dato align='left' width='10%' ", "left", false,"<b>" & texto_total & "</b>"
		    sumatotalimp=0
		    sumatotalcant=0
		    cantidad_a_poner=0
		    importe_a_poner=0
		    for i=1 to contador_columnas
			    if lista(i,3) & "">"" then
				    if lista(i,3)<>0 then
					    cantidad_a_poner=formatnumber(lista(i,3),dec_cant,-1,0,-1)
					    sumatotalcant=sumatotalcant + cantidad_a_poner
				    else
					    cantidad_a_poner=""
					    sumatotalcant=sumatotalcant + 0
				    end if
			    else
				    cantidad_a_poner=""
				    sumatotalcant=sumatotalcant + 0
			    end if
			    if lista(i,4) & "">"" then
				    if lista(i,4)<>0 then
					    importe_a_poner=formatnumber(lista(i,4),dec_mb,-1,0,-1)
					    sumatotalimp=sumatotalimp + importe_a_poner
				    else
					    importe_a_poner=""
					    sumatotalimp=sumatotalimp + 0
				    end if
			    else
				    importe_a_poner=""
				    sumatotalimp=sumatotalimp + 0
			    end if

			    if importe_a_poner & "">"" and cantidad_a_poner & "">"" then
				    dto_a_poner=(cantidad_a_poner*precio_ticket)-importe_a_poner
				    dto_a_poner=formatnumber(dto_a_poner,decpor,-1,0,-1)
			    else
				    dto_a_poner=0
			    end if

			    if mostrarinfo=ucase(LitCantidad) then
				    DrawCelda2 "dato align='right' ", "right", false,"<b>" & cantidad_a_poner & "</b>"
			    else
				    texto_a_poner=importe_a_poner
				    if si_dto=1 then
					    if dto_a_poner & "">"" and dto_a_poner<>0 then
						    j=1
						    si_poner=0
						    while si_poner=0 and j<=contador_lista2
							    if lista(i,1) & ""=lista2(j,1) & "" then
								    si_poner=1
								    'si queremos poner el dto total de la columna se deja como esta
								    ' y si queremos solo los de las familias de la variable dtofam
								    dto_a_poner=formatnumber(lista2(j,2),decpor,-1,0,-1)
								    dto_total_a_poner=dto_total_a_poner + dto_a_poner
							    end if
							    j=j+1
						    wend
						    if si_poner=1 then
							    texto_a_poner=texto_a_poner & "<br/>" & LitDescuento & " : " & dto_a_poner
						    end if
					    end if
				    end if
				    DrawCelda2 "dato align='right' ", "right", false,"<b>" & texto_a_poner & "</b>"
			    end if
		    next
	    if colum=1 then
		    if sumatotalcant & "">"" then
			    sumatotalcant=formatnumber(sumatotalcant,dec_cant,-1,0,-1)
		    end if
		    DrawCelda2 "dato align='right' width='10%' ", "right", false,"<b>" & sumatotalcant & "</b>"
		    if sumatotalimp & "">"" then
			    sumatotalimp=formatnumber(sumatotalimp,dec_mb,-1,0,-1)
			    sumadto=(sumatotalcant*precio_ticket)-sumatotalimp
			    dto_a_poner=formatnumber(sumadto,decpor,-1,0,-1)
		    end if
		    texto_a_poner=sumatotalimp
		    if si_dto=1 then
			    'si queremos poner el dto total de la columna se deja como esta
			    ' y si queremos solo los de las familias de la variable dtofam
			    dto_a_poner=dto_total_a_poner
			    if dto_a_poner & "">"" and dto_a_poner<>0 then
				    texto_a_poner=texto_a_poner & "<br/>" & LitDescuento & " : " & dto_a_poner
			    end if
		    end if
		    DrawCelda2 "dato align='right' width='10%' ", "right", false,"<b>" & texto_a_poner & "</b>"
	    else
		    DrawCelda2 "dato align='right' width='10%' ", "right", false,"&nbsp;"
		    DrawCelda2 "dato align='right' width='10%' ", "right", false,"&nbsp;"
	    end if
	    CloseFila
    next
end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************

const borde=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion

    'MAP 21/12/2012 - Get serie UserParams (Resumen de Ventas > Listado de Tickets)
    dim s
    ObtenerParametros("list_tickets")
%>
	<div id="ver_alt" class=ETIQUETATPV style="display:none;position:absolute;width:150px;z-index:200;"></div>

   <form name="listado_tickets" method="post">
	<% PintarCabecera "listado_tickets.asp"
	WaitBoxOculto LitEsperePorFavor

	'Leer parámetros de la página
    mode=Request.QueryString("mode")

	if mode="browse" then mode="imp"

	if request.querystring("dtofam")>"" then
		dtofam = enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.querystring("dtofam"))))
	else
		dtofam = enc.EncodeForHtmlAttribute(null_s(limpiaCadena(request.Form("dtofam"))))
	end if
	if dtofam & ""="" then dtofam="@@%%!!ÑÑ>><<{}[]ªº?¿¡"

	if request.querystring("verp")>"" then
		verp=limpiaCadena(request.querystring("verp"))
	else
		verp=limpiaCadena(request.Form("verp"))
	end if

	if request.querystring("vert")>"" then
		vert=limpiaCadena(request.querystring("vert"))
	else
		vert=limpiaCadena(request.Form("vert"))
	end if

	if request.querystring("verc")>"" then
		verc=limpiaCadena(request.querystring("verc"))
	else
		verc=limpiaCadena(request.Form("verc"))
	end if

	if request.querystring("caju")>"" then
		caju=limpiaCadena(request.querystring("caju"))
	else
		caju=limpiaCadena(request.form("caju"))
	end if
	
	if request.querystring("i")>"" then
		coniva=limpiaCadena(request.querystring("i"))
	else
		coniva=limpiaCadena(request.form("i"))
	end if

	if request.querystring("operador")>"" then
		operador=limpiaCadena(request.querystring("operador"))
	else
		operador=limpiaCadena(request.Form("operador"))
	end if
	TmpOperador2=TmpOperador

	if request.querystring("tipoperador")>"" then
		tipoperador=limpiaCadena(request.querystring("tipoperador"))
	else
		tipoperador=limpiaCadena(request.Form("tipoperador"))
	end if
	CheckCadena tipoperador

	if request.querystring("serie")>"" then
		serie=limpiaCadena(request.querystring("serie"))
	else
		serie=limpiaCadena(request.Form("serie"))
	end if
	CheckCadena serie

	if request.querystring("tienda")>"" then
		tienda=limpiaCadena(request.querystring("tienda"))
	else
		tienda=limpiaCadena(request.Form("tienda"))
	end if
	CheckCadena tienda

	if request.querystring("caja")>"" then
		caja=limpiaCadena(request.querystring("caja"))
	else
		caja=limpiaCadena(request.Form("caja"))
	end if
	CheckCadena caja

	if request.querystring("tpv")>"" then
		tpv=limpiaCadena(request.querystring("tpv"))
	else
		tpv=limpiaCadena(request.Form("tpv"))
	end if
	CheckCadena tpv

	if request.querystring("agruparpor")>"" then
		agruparpor=limpiaCadena(request.querystring("agruparpor"))
	else
		agruparpor=limpiaCadena(request.Form("agruparpor"))
	end if

	if request.querystring("dfecha")>"" then
		dfecha=limpiaCadena(request.querystring("dfecha"))
	else
		dfecha=limpiaCadena(request.Form("dfecha"))
	end if

	if request.querystring("hfecha")>"" then
		hfecha=limpiaCadena(request.querystring("hfecha"))
	else
		hfecha=limpiaCadena(request.Form("hfecha"))
	end if

	if request.querystring("mediopago")>"" then
		mediopago=limpiaCadena(request.querystring("mediopago"))
	else
		mediopago=limpiaCadena(request.Form("mediopago"))
	end if

	if request.querystring("conref")>"" then
		conref=limpiaCadena(request.querystring("conref"))
	else
		conref=limpiaCadena(request.Form("conref"))
	end if

	if request.querystring("connombre")>"" then
		connombre=limpiaCadena(request.querystring("connombre"))
	else
		connombre=limpiaCadena(request.Form("connombre"))
	end if

	if request.querystring("tiparticulo")>"" then
		tiparticulo=limpiaCadena(request.querystring("tiparticulo"))
	else
		tiparticulo=limpiaCadena(request.Form("tiparticulo"))
	end if
	CheckCadena tiparticulo

	if request.querystring("familia")>"" then
		familia=limpiaCadena(request.querystring("familia"))
	else
		familia=limpiaCadena(request.Form("familia"))
	end if
	CheckCadena familia

	if request.querystring("apaisado")>"" then
		apaisado=limpiaCadena(request.querystring("apaisado"))
	else
		apaisado=limpiaCadena(request.Form("apaisado"))
	end if

	if request.querystring("mostrarinfo")>"" then
		mostrarinfo=limpiaCadena(request.querystring("mostrarinfo"))
	else
		mostrarinfo=limpiaCadena(request.Form("mostrarinfo"))
	end if

	if request.querystring("mostrardesc")>"" then
		mostrardesc=limpiaCadena(request.querystring("mostrardesc"))
	else
		mostrardesc=limpiaCadena(request.Form("mostrardesc"))
	end if

	if request.querystring("mostrarcolum")>"" then
		mostrarcolum=limpiaCadena(request.querystring("mostrarcolum"))
	else
		mostrarcolum=limpiaCadena(request.Form("mostrarcolum"))
	end if

    'DGM 22/02/2011 Recogemos los datos nuevos de la mejora
    if request.QueryString("dhora") > "" then
        dhora = limpiaCadena(request.QueryString("dhora"))
    else
        dhora = limpiaCadena(request.Form("dhora"))
    end if
        
    if request.QueryString("hhora") > "" then
        hhora = limpiaCadena(request.QueryString("hhora"))
    else
        hhora = limpiaCadena(request.Form("hhora"))
    end if
       
    if request.QueryString("familia_padre") > "" then
        familia_padre = limpiaCadena(request.QueryString("familia_padre"))
    else
        familia_padre = limpiaCadena(request.Form("familia_padre"))
    end if
        
    if request.QueryString("categoria") > "" then
        categoria = limpiaCadena(request.QueryString("categoria"))
    else
        categoria = limpiaCadena(request.Form("categoria"))
    end if
    
    if request.QueryString("detalle") > "" then
        detalle = limpiaCadena(request.QueryString("detalle"))
    else
        detalle = limpiaCadena(request.Form("detalle"))
    end if
    '01/07/2016 DBG 
    if request.querystring("opcdesgloseiva")>"" then
		TmpOpcDesgloseiva=limpiaCadena(request.querystring("opcdesgloseiva"))
	else
		TmpOpcDesgloseiva=limpiaCadena(request.Form("opcdesgloseiva"))
	end if
    'MAP 03/01/2013 - preparar lista de series con los datos indicados en el parámetro de usuario
     if s&""="" then
	    s=limpiaCadena(request.querystring("s"))
	    if s="" then s=limpiaCadena(request.form("s"))
        end if
	    s=preparar_lista(s)

	set rst = Server.CreateObject("ADODB.Recordset")
	set rst2 = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstcursor=Server.CreateObject("ADODB.Recordset")
    set rstDesgloseIva = Server.CreateObject("ADODB.Recordset")
    set rstDesgloseIvaRecorreBucle = Server.CreateObject("ADODB.Recordset")

	'se crea la tabla temporal para elegir varios operadores,cajas,tiendas y tpv
	if mode="param" then

		strdrop ="if exists (select * from sysobjects where id = object_id('[" & session("usuario") & "-temporal]') and sysstat " & _
		" & 0xf = 3) drop table [" & session("usuario") & "-temporal]"
		rstAux.open strdrop,Session("backendlistados"),adUseClient,adLockReadOnly
		if rstAux.state<>0 then rstAux.close

		strselect="CREATE TABLE [" & session("usuario") & "-temporal] (viene varchar(15),codigo varchar(50),seleccionado bit)"
		rstAux.open strselect,Session("backendlistados"),adUseClient,adLockReadOnly
		GrantUser session("usuario") & "-temporal", Session("backendlistados")
	end if

	if mode="param" or mode="traeroperador" then                                                                               
		%><input type="hidden" name="dtofam" value="<%=dtofam%>">
		<input type="hidden" name="verp" value="<%=EncodeForHtml(verp)%>">
		<input type="hidden" name="vert" value="<%=EncodeForHtml(vert)%>">
		<input type="hidden" name="verc" value="<%=EncodeForHtml(verc)%>">
		<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>">
		<input type="hidden" name="i" value="<%=EncodeForHtml(coniva)%>"><%
	end if

	strwhere=""%>

	<table width='100%' cellspacing="1" cellpadding="1">
   		<tr>
				<%if mode="browse" then%>
					<td width="40%" align="center" bgcolor="<%=color_fondo%>">
				  	<font class=CELDAC7>&nbsp;(Emitido el <%=day(date)%>/<%=month(date)%>/<%=year(date)%>)</font>
					</td>
				  <%end if%>
			<!--</td>-->
   			<td><font class='CABECERA'><b></b></font>
 	     			<font class=CELDA><b></b></font>
			</td>                                                                                                                
			<%if mode="imp" then%>
				<td class=CELDARIGHT bgcolor="">
					<%fdesde=dfecha
					fhasta=hfecha
					if fdesde>"" then
						if fhasta>"" then
							%><%=LitPeriodoFechas%> : <%=EncodeForHtml(fdesde)%> - <%=EncodeForHtml(fhasta)%><%
						else
							%><%=LitPeriodoFechas%> : <%=LitDesde%>&nbsp;<%=EncodeForHtml(fdesde)%><%
						end if
					else
						if fhasta>"" then
							%><%=LitPeriodoFechas%> : <%=LitHasta%>&nbsp;<%=EncodeForHtml(fhasta)%><%
						else
						end if
					end if%>
				</td><%
			else%>
				<td></td><%
			end if%>
   		</tr>
	</table>
	<hr/>
	<%if mode="param" then
      DrawDiv "1", "", ""
            DrawLabel "", "", LitListTicket%><input type="radio" name="listaTicket" onclick="javascript:paramListado('tickets')"><%
      CloseDiv
      DrawDiv "1", "", ""
            DrawLabel "", "", LitListResumen%><input type="radio" name="listaTicket" onclick="javascript:paramListado('resumen')" checked><%
      CloseDiv
        %>
	<hr/>
	<%end if%>

	<% Alarma "listado_tickets.asp"

	'*************************************************************************************************************'
	if mode="traeroperador" then
		if operador> "" then
			operador=session("ncliente") & operador
			nomoperador=d_lookup("nombre","personal","dni='" & operador & "'",Session("backendlistados"))
			if nomoperador="" then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgOperadorNoExiste%>");
				</script><%
				operador=""
			else
				rst2.open "select dni,nombre,fbaja from personal with(nolock) where dni like '"& session("ncliente")&"%' and dni='" & operador & "' and fbaja is null",Session("backendlistados"),adOpenKeyset,adLockOptimistic
				if rst2.eof then
					%><script language="javascript" type="text/javascript">
					      window.alert("<%=LitMsgOperadorDadoBaja%>");
					</script><%
					nomoperador=""
					operador=""
				else
					nomoperador=rst2("nombre")
				end if
				rst2.close
			end if
			mode="param"
		else
			operador=""
			nomoperador=""
			mode="param"
		end if
	end if

	if mode="param" then
            EligeCelda "input","add","left","","",0,LitDesdeFecha,"dfecha",0,iif(dfecha>"",dfecha,"01/01/" & year(date))
            DrawCalendar "dfecha"

            EligeCelda "input","add","left","","",0,LitDHora,"dhora",0,iif(dhora>"",dhora,"00:00")

            EligeCelda "input","add","left","","",0,LitHastaFecha,"hfecha",0,CDate(iif(hfecha>"",hfecha,day(date) & "/" & month(date) & "/" & year(date)))
            DrawCalendar "hfecha"

            EligeCelda "input","add","left","","",0,LitHHora,"hhora",0,iif(hhora>"",hhora,"23:59")
            %>
	    <script language="javascript" type="text/javascript">
	        function dhora_callkeyuphandler(evnt)
	        {
	            ev = (evnt) ? evnt : event;
	            formatHora(document.listado_tickets.dhora, false, ev);
	        }

	        if(window.document.listado_tickets.dhora.addEventListener)
	        {
	            window.document.listado_tickets.dhora.addEventListener("keyup", dhora_callkeyuphandler,false);
	        }
	        else
	        {
	            window.document.listado_tickets.attachEvent("onkeyup", dhora_callkeyuphandler);
	        }

	        function hhora_callkeyuphandler(evnt)
	        {
	            ev = (evnt) ? evnt : event;
	            formatHora(document.listado_tickets.hhora,false,ev);
	        }

	        if(window.document.listado_tickets.hhora.addEventListener)
	        {
	            window.document.listado_tickets.hhora.addEventListener("keyup", hhora_callkeyuphandler, false);
	        }
	        else
	        {
	            window.document.listado_tickets.hhora.attachEvent("onkeyup", hhora_callkeyuphandler);
	        }
        </script>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
        	    
		<%
			DrawDiv "1","",""%><label><a class="CELDAREF7" style="margin: 0px;" href="javascript:AbrirElegir('operador')"><%=LitOperador%></a></label><font class="CELDAREDBOLD" id='multioper' style='display:none'>(<%=LitListTicMulti%>)</font><%
			if nomoperador & ""="" then
				nomoperador=d_lookup("nombre","personal","dni='" & operador & "'",Session("backendlistados"))
			end if
			%><input class='width15' type="text" name="operador" size=10 value="<%=EncodeForHtml(trimCodEmpresa(operador))%>" onchange="TraerOperador();"><a class='CELDAREFB'  href="javascript:AbrirVentana('../../administracion/personal_buscar.asp?viene=listado_tickets&titulo=<%=LitSelPersonal%>&mode=search','P',<%=AltoVentana%>,<%=AnchoVentana%>)" OnMouseOver="self.status='<%=LitVerPersonal%>'; return true;" OnMouseOut="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscarDinamic%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class="width40" type="text" name="nomoperador" size="25" value="<%=EncodeForHtml(nomoperador)%>"><%CloseDiv

			rstAux.open "select codigo,descripcion from tipos_entidades with(nolock) where tipo='PERSONAL' and codigo like '" & session("ncliente") & "%'",Session("backendlistados"),adOpenKeyset,adLockOptimistic
            DrawSelectCelda "CELDA","","",0,LitTipOperador,"tipoperador",rstAux,iif(tipoperador>"",tipoperador,""),"codigo","descripcion","",""
			rstAux.close

            strSacSerie= "select nserie,nombre as descripcion from series with(nolock) where tipo_documento='TICKET' and nserie like '" & session("ncliente") & "%'"
			
            if s & "">"" then
				strSacSerie=strSacSerie & " and nserie in "+ s
			end if
			strSacSerie=strSacSerie & " order by nserie"

            rstAux.cursorlocation=3
			rstAux.open strSacSerie,session("backendlistados"),adOpenKeyset,adLockOptimistic

            DrawSelectCelda "CELDA","","",0,LitSerie,"serie",rstAux,iif(serie>"",serie,""),"nserie","descripcion","",""
			rstAux.close

			rstAux.open "select codigo,descripcion from tipo_pago with(nolock) where codigo like '" & session("ncliente") & "%'",Session("backendlistados"),adOpenKeyset,adLockOptimistic
            DrawSelectCelda "CELDA","","",0,LitMPago,"mediopago",rstAux,iif(mediopago>"",mediopago,""),"codigo","descripcion","",""
			rstAux.close

			rstAux.open "select codigo,descripcion from tiendas with(nolock) where codigo like '" & session("ncliente") & "%'",Session("backendlistados"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "CELDA","","",0,LitTicketTienda,"tienda",rstAux,tienda,"codigo","descripcion","",""
			rstAux.close

			if caja="" then
				caja=" " 'se pone espacio en blanco, para que salga vacio el campo caja
			end if
            poner_cajasResponsive "CELDA",caja,LitCaja,"caja","150","codigo","descripcion","","",poner_comillas(caju)

			rstAux.open "select tpv,descripcion from tpv with(nolock) where tpv like '" & session("ncliente") & "%'",Session("backendlistados"),adOpenKeyset,adLockOptimistic
            DrawSelectCelda "CELDA","","",0,LitTicketTpv,"tpv",rstAux,tpv,"tpv","descripcion","",""
			rstAux.close
		'DGM 21/02/2011 Añadir filtro CATEGORIA-FAMILIA-SUBFAMILIA
			dim ConfigDespleg (3,13)

			i=0
			ConfigDespleg(i,0)="categoria"
			ConfigDespleg(i,1)="200"
			ConfigDespleg(i,2)="5"
			ConfigDespleg(i,3)="select codigo, nombre from categorias with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="CELDA"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitCategoria
			ConfigDespleg(i,10)=categoria
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=1
			ConfigDespleg(i,0)="familia_padre"
			ConfigDespleg(i,1)="200"
			ConfigDespleg(i,2)="5"
			ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="CELDA"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitFamilia
			ConfigDespleg(i,10)=familia_padre
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=2
			ConfigDespleg(i,0)="familia"
			ConfigDespleg(i,1)="200"
			ConfigDespleg(i,2)="5"
			ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="CELDA"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitSubFamilia
			ConfigDespleg(i,10)=familia
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			DibujaDesplegables ConfigDespleg,session("backendlistados")
		
            EligeCelda "input","add","left","","",0,LitRefListTpv,"conref",38,conref
            EligeCelda "input","add","left","","",0,LitNomListTpv,"connombre",38,connombre

            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("backendListados")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="getAllEntityTypeByType"
            command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
            command.Parameters.Append command.CreateParameter("@ncompany", adVarChar, adParamInput, 5, session("ncliente"))
            command.Parameters.Append command.CreateParameter("@type", adVarChar, adParamInput, 20, "ARTICULO")

            set rstArtType = command.execute

            DrawSelectCelda "CELDA","180","multiple",0,LitTipArticulo,"tiparticulo",rstArtType,tiparticulo,"codigo","descripcion","",""

            rstArtType.close
            conn.close
            set rstArtType = nothing
            set command = nothing
            set conn = nothing
            DrawDiv "1","",""
            DrawLabel "","",LitInfAMostrar%><select name="mostrarinfo" class="width60" >
					<option <%=iif(ucase(mostrarinfo)=ucase(LitImporte) or mostrarinfo="","selected","")%> value="<%=ucase(LitImporte)%>"><%=ucase(LitImporte)%></option>
					<option <%=iif(ucase(mostrarinfo)=ucase(LitCantidad),"selected","")%> value="<%=ucase(LitCantidad)%>"><%=ucase(LitCantidad)%></option>
				</select><%CloseDiv

            DrawDiv "1","",""
            DrawLabel "","",LitAgrupListTpv%><select name="agruparpor" class="width60" >
					<option <%=iif(ucase(agruparpor)=ucase(LitSubFamilia),"selected","")%> value="<%=ucase(LitSubFamilia)%>"><%=ucase(LitSubFamilia)%></option>
					<option <%=iif(ucase(agruparpor)=ucase(LitFamilia),"selected","")%> value="<%=ucase(LitFamilia)%>"><%=ucase(LitFamilia) %></option>
					<option <%=iif(ucase(agruparpor)=ucase(LitCategoria),"selected","")%> value="<%=ucase(LitCategoria)%>"><%=ucase(LitCategoria)%> </option>
					<option <%=iif(agruparpor="","selected","")%> value=""></option>
				</select><%CloseDiv
            DrawDiv "1","",""
            DrawLabel "","",LitMostrarEnCol%><select name="mostrarcolum" class="width60" >
					<option <%=iif(ucase(mostrarcolum)=ucase(LitTicketTpv) or mostrarcolum="","selected","")%> value="<%=ucase(LitTicketTpv)%>"><%=ucase(LitTicketTpv)%></option>
					<option <%=iif(ucase(mostrarcolum)=ucase(LitOperador),"selected","")%> value="<%=ucase(LitOperador)%>"><%=ucase(LitOperador)%></option>
					<option <%=iif(ucase(mostrarcolum)=ucase(LitCaja),"selected","")%> value="<%=ucase(LitCaja)%>"><%=ucase(LitCaja)%></option>
					<option <%=iif(ucase(mostrarcolum)=ucase(LitTicketTienda),"selected","")%> value="<%=ucase(LitTicketTienda)%>"><%=ucase(LitTicketTienda)%></option>
				</select><%CloseDiv
		%><hr /><%
		'DGM 22/02/2011 Añadimos la opcion de mostrar con detalle o sin detalle  
		DrawDiv "1","",""
        DrawLabel "","",LitConDetalle%><input type="radio" name="detalle" value="conDet" checked="checked" /><%CloseDiv
        DrawDiv "1","",""
        DrawLabel "","",LitSinDetalle%><input type="radio" name="detalle" value="sinDet" /><%CloseDiv
        %> 
		<%
        DrawDiv "1","",""
        DrawLabel "","",LitMostrarDesc%><input type="checkbox" name="mostrardesc" <%=iif(mostrardesc=-1,"checked","")%> ><%CloseDiv
        EligeCelda "check","add","left","","",0,LitApaisadoListTpv,"apaisado",0,iif(apaisado=-1 or apaisado="" or apaisado="on",-1,0)
        EligeCelda "check","add","left","","",0,LitDesgloseIva,"opcdesgloseiva",0,iif(TmpOpcDesgloseiva="on","True","")
        EligeCelda "check","add","left","","",0,LITIMPORTEIVA,"opcprecioiva",0,checked
        %>
       
	<%elseif mode="imp" then
		''ricardo 12-12-2007 si viene de paginacion se debera quitar el nempresa al operador
		lotePag=limpiaCadena(Request.QueryString("lote"))
		lotesPag=limpiaCadena(Request.form("lotesPag"))
		if lotesPag="" then
		    lotesPag=limpiaCadena(Request.QueryString("lotesPag"))
		end if
		condicionlotes=0
		if lotePag & "">"" and lotesPag & "">"" then
		    if cint(lotePag)=cint(lotesPag) then
		        condicionlotes=1
		    end if
		end if
		sentidoPag=limpiaCadena(Request.QueryString("sentido"))
		if condicionlotes=1 or sentidoPag & "">"" then
		    operador=trimCodEmpresa(operador)
		end if
		'''''''''''''''''''''''''''''''''''''''

		if operador & "">"" then%>
			<input type="hidden" name="operador" value="<%=session("ncliente") & EncodeForHtml(operador)%>" />
		<%else%>
			<input type="hidden" name="operador" value="<%=EncodeForHtml(operador)%>" />
		<%end if%>
		<input type="hidden" name="dfecha" value="<%=EncodeForHtml(dfecha)%>" />
		<input type="hidden" name="hfecha" value="<%=EncodeForHtml(hfecha)%>" />
		<input type="hidden" name="serie" value="<%=EncodeForHtml(serie)%>" />
		<input type="hidden" name="tienda" value="<%=EncodeForHtml(tienda)%>" />
		<input type="hidden" name="caja" value="<%=EncodeForHtml(caja)%>" />
		<input type="hidden" name="tpv" value="<%=EncodeForHtml(tpv)%>" />
		<input type="hidden" name="mediopago" value="<%=EncodeForHtml(mediopago)%>" />
		<input type="hidden" name="tipoperador" value="<%=EncodeForHtml(tipoperador)%>" />
		<input type="hidden" name="agruparpor" value="<%=EncodeForHtml(agruparpor)%>" />
		<input type="hidden" name="conref" value="<%=EncodeForHtml(conref)%>" />
		<input type="hidden" name="connombre" value="<%=EncodeForHtml(connombre)%>" />
		<input type="hidden" name="tiparticulo" value="<%=EncodeForHtml(tiparticulo)%>" />
		<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>" />
		<input type="hidden" name="mostrarcolum" value="<%=EncodeForHtml(mostrarcolum)%>" />
		<input type="hidden" name="mostrardesc" value="<%=EncodeForHtml(mostrardesc)%>" />
		<input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>" />
		<input type="hidden" name="mostrarinfo" value="<%=EncodeForHtml(mostrarinfo)%>" />
		<input type="hidden" name="dtofam" value="<%=dtofam%>" />
		<input type="hidden" name="verp" value="<%=EncodeForHtml(verp)%>" />
		<input type="hidden" name="vert" value="<%=EncodeForHtml(vert)%>" />
		<input type="hidden" name="verc" value="<%=EncodeForHtml(verc)%>" />
		<input type="hidden" name="caju" value="<%=EncodeForHtml(caju)%>" />
		<input type="hidden" name="i" value="<%=EncodeForHtml(coniva)%>" />
		<% 'DGM 22/02/2011
		   ' El campo famiila pertenece en el esquema a subfamilia
		   ' Añadimos los nuevos valores %>
		<input type="hidden" name="familia_padre" value="<%=EncodeForHtml(familia_padre)%>" />
		<input type="hidden" name="categoria" value="<%=EncodeForHtml(categoria)%>" />
		<input type="hidden" name="dhora" value="<%=EncodeForHtml(dhora)%>" />
		<input type="hidden" name="hhora" value="<%=EncodeForHtml(hhora)%>" />
		<input type="hidden" name="detalle" value="<%=EncodeForHtml(detalle)%>" />
        <input type="hidden" name="opcdesgloseiva" value="<%=EncodeForHtml(TmpOpcDesgloseiva)%>" />
       
        <%
        '---------------------------cambios ILYA-----------------------------------------------------------------------------------------
        dec_mb=d_lookup("ndecimales","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",Session("backendlistados"))
        strselect=""
		strselect="select upper(d.referencia) as referencia"
		strselect=strselect & ",sum(d.cantidad) as total_cant"
		if coniva="SI" then
			strselect=strselect & ",sum(round(d.importeiva,"& dec_mb &")) as total_imp"		
            strselect=strselect & ",sum(round(d.importeiva,"& dec_mb &")) as total_imp_iva"	
		else
			strselect=strselect & ",sum(round(d.importe,"& dec_mb &")) as total_imp"
            strselect=strselect & ",sum(round(d.importeiva,"& dec_mb &")) as total_imp_iva"	
		end if

		strselect=strselect & ",ti.divisa"
		encabezado=0

		%><table width="100%"><tr><td width="85%" align="left" valign="top"><%

		select case mostrarinfo
			case ucase(LitImporte):
				%><font class=cab><b><%=LitInfAMostrar%> : </b></font><font class=cab><%=ucase(LitImporte)%></font><br/><%
				encabezado=1
			case ucase(LitCantidad):
				%><font class=cab><b><%=LitInfAMostrar%> : </b></font><font class=cab><%=ucase(LitCantidad)%></font><br/><%
				encabezado=1
		end select

		select case mostrarcolum
			case ucase(LitTicketTpv):
				strselect=strselect & ",ti.tpv,tpv.descripcion as nombretpv"
				%><font class=cab><b><%=LitMostrarEnCol%> : </b></font><font class=cab><%=ucase(LitTicketTpv)%></font><br/><%
				encabezado=1
			case ucase(LitOperador):
				strselect=strselect & ",ti.usuario,p.nombre as nompersonal"
				%><font class=cab><b><%=LitMostrarEnCol%> : </b></font><font class=cab><%=ucase(LitOperador)%></font><br/><%
				encabezado=1
			case ucase(LitCaja):
				strselect=strselect & ",ti.caja,c.descripcion as nomcaja"
				%><font class=cab><b><%=LitMostrarEnCol%> : </b></font><font class=cab><%=ucase(LitCaja)%></font><br/><%
				encabezado=1
			case ucase(LitTicketTienda):
				strselect=strselect & ",c.tienda,tda.descripcion as nomtienda"
				%><font class=cab><b><%=LitMostrarEnCol%> : </b></font><font class=cab><%=ucase(LitTicketTienda)%></font><br/><%
				encabezado=1
		end select

		strselect=strselect & ",a.nombre as nomarticulo"

        strselect=strselect & ",a.familia,f.nombre as nomfamilia,a.familia_padre as fampadre, a.categoria as art_cod_categoria"
        
        ''DBG  04/07/2016 en strSelectIva me guardo los mismo select de la consulta para rellenar la tabla temporal principal 
        strSelectIva=""
        strSelectIva = strselect & ",d.iva as porcentaje_iva, sum(round(d.importe,"& dec_mb &")) as importe_base_total , div.FACTCAMBIO as factorcambioOrigen "

		strfrom=""
		strfrom=strfrom & " from tickets as ti with(nolock)"
        if mediopago>"" then
            strfrom=strfrom & " cross apply (select distinct c.ndocumento from CAJA c with(nolock) where c.caja like '"&session("ncliente") &"%' and c.TDOCUMENTO='TICKET' and c.MEDIO='" & mediopago & "' and c.ndocumento=ti.nticket) as CMP "
        end if
		strfrom=strfrom & " left outer join personal as p with(nolock) on ti.usuario=p.dni and p.dni like '" & session("ncliente") & "%' "
		strfrom=strfrom & " left outer join tpv with(nolock) on tpv.tpv=ti.tpv and tpv.tpv like '" & session("ncliente") & "%' "
		strfrom=strfrom & " left outer join cajas as c with(nolock) on c.codigo=ti.caja and c.codigo like '" & session("ncliente") & "%' "
		strfrom=strfrom & " left outer join tiendas as tda with(nolock) on tda.codigo=c.tienda and  tda.codigo like '" & session("ncliente") & "%' "
		strfrom=strfrom & " ,detalles_tickets as d with(nolock),articulos as a with(nolock)"
		strfrom=strfrom & " left outer join familias as f with(nolock) on f.codigo=a.familia and f.codigo like '" & session("ncliente") & "%' "
		strwhere=" where "
		strwhere=strwhere & "ti.nticket like '" & session("ncliente") & "%' and d.nticket like '" & session("ncliente") & "%' and "
		strwhere=strwhere & "a.referencia like '" & session("ncliente") & "%' and "

		if serie>"" then
			strwhere=strwhere & " ti.serie='" & serie & "' and"
			%><font class=cab><b><%=LitSerie%> : </b></font><font class=cab><%=EncodeForHtml(trimCodEmpresa(serie)) & " - " & d_lookup("nombre","series","nserie='" & serie & "'",Session("backendlistados"))%></font><br/><%
			encabezado=1
                                                                               
        elseif s>"" then                                                                            
        'map 15/01/2013 - Serie as filter
            strwhere=strwhere + " ti.serie in "+ s + " and"
		end if

		
		if operador >"" then
			strwhere=strwhere & " ti.usuario='" & session("ncliente") & operador & "' and"
			%><font class=cab><b><%=LitOperador%> : </b></font><font class=cab><%=EncodeForHtml(operador)%> - <%=d_lookup("nombre","personal","dni='" & session("ncliente") & operador & "'",Session("backendlistados"))%></font><br/><%
			encabezado=1
		else
			strselect2="select codigo from [" & session("usuario") & "-temporal] where viene='operador' and seleccionado<>0"
			rst.cursorlocation=3
			rst.Open strselect2, Session("backendlistados")
			listaCampo=""
			if not rst.eof then
				strwhere=strwhere & " ti.usuario in (" & strselect2 & ") and"
				while not rst.eof
					listaCampo=listaCampo & trimCodEmpresa(rst("codigo")) & " - " & d_lookup("nombre","personal","dni='" & rst("codigo") & "'",Session("backendlistados")) & ","
					rst.movenext
				wend
			end if
			rst.close                                                                                            
			if listaCampo<>"" and mostrarcolum<>ucase(LitOperador) then
				listaCampo=mid(listaCampo,1,len(listaCampo)-1)
				%><font class=cab><b><%=LitOperador%> : </b></font><font class=cab><%=EncodeForHtml(listaCampo)%></font><br/><%
				encabezado=1
			else
				strwhere=strwhere & " ti.usuario like '" & session("ncliente") & "%' and"
			end if
		end if

		if tipoperador >"" then
			strwhere=strwhere & " p.tipo='" & tipoperador & "' and"
			%><font class=cab><b><%=LitTipOperador%> : </b></font><font class=cab><%=d_lookup("descripcion","tipos_entidades","codigo='" & tipoperador & "'",Session("backendlistados"))%></font><br/><%
			encabezado=1
		end if

		if mediopago>"" then
			%><font class=cab><b><%=LitMPago%> : </b></font><font class=cab><%=d_lookup("descripcion","tipo_pago","codigo='" & mediopago & "'",Session("backendlistados"))%></font><br/><%
			encabezado=1
		end if

		if tienda>"" then
			strwhere=strwhere & " tda.codigo='" & tienda & "' and"
			%><font class=cab><b><%=LitTicketTienda%> : </b></font><font class=cab><%=d_lookup("descripcion","tiendas","codigo='" & tienda & "'",Session("backendlistados"))%></font><br/><%
			encabezado=1
		else
			strselect2="select codigo from [" & session("usuario") & "-temporal] where viene='tiendas' and seleccionado<>0"
			rst.cursorlocation=3
			rst.Open strselect2, Session("backendlistados")
			listaCampo=""
			if not rst.eof then
				strwhere=strwhere & " tda.codigo in (" & strselect2 & ") and"
				while not rst.eof
					listaCampo=listaCampo & d_lookup("descripcion","tiendas","codigo='" & rst("codigo") & "'",Session("backendlistados")) & ","
					rst.movenext
				wend
			end if
			rst.close
			if listaCampo<>"" and mostrarcolum<>ucase(LitTicketTienda) then
				listaCampo=mid(listaCampo,1,len(listaCampo)-1)
				%><font class=cab><b><%=LitTicketTienda%> : </b></font><font class=cab><%=listaCampo%></font><br/><%
				encabezado=1
			end if
		end if

		if caja>"" then
			strwhere=strwhere & " c.codigo='" & caja & "' and"
			%><font class=cab><b><%=LitCaja%> : </b></font><font class=cab><%=d_lookup("descripcion","cajas","codigo='" & caja & "'",Session("backendlistados"))%></font><br/><%
			encabezado=1
		else
			strselect2="select codigo from [" & session("usuario") & "-temporal] where viene='caja' and seleccionado<>0"
			rst.cursorlocation=3
			rst.Open strselect2, Session("backendlistados")
			listaCampo=""
			if not rst.eof then
				strwhere=strwhere & " c.codigo in (" & strselect2 & ") and"
				while not rst.eof
					listaCampo=listaCampo & d_lookup("descripcion","cajas","codigo='" & rst("codigo") & "'",Session("backendlistados")) & ","
					rst.movenext
				wend
			end if
			rst.close
			if listaCampo<>"" and mostrarcolum<>ucase(LitCaja) then
				listaCampo=mid(listaCampo,1,len(listaCampo)-1)
				%><font class=cab><b><%=LitCaja%> : </b></font><font class=cab><%=listaCampo%></font><br/><%
				encabezado=1
			end if
		end if

		if tpv>"" then
			strwhere=strwhere & " tpv.tpv='" & tpv & "' and"
			%><font class=cab><b><%=LitTicketTpv%> : </b></font><font class=cab><%=d_lookup("descripcion","tpv","tpv='" & tpv & "'",Session("backendlistados"))%></font><br/><%
			encabezado=1
		else
			strselect2="select codigo from [" & session("usuario") & "-temporal] where viene='tpv' and seleccionado<>0"
			rst.cursorlocation=3
			rst.Open strselect2, Session("backendlistados")
			listaCampo=""
			if not rst.eof then
				strwhere=strwhere & " tpv.tpv in (" & strselect2 & ") and"
				while not rst.eof
					listaCampo=listaCampo & d_lookup("descripcion","tpv","tpv='" & rst("codigo") & "'",Session("backendlistados")) & ","
					rst.movenext
				wend
			end if                                                                                                 
			rst.close
			if listaCampo<>"" and mostrarcolum<>ucase(LitTicketTpv) then
				listaCampo=mid(listaCampo,1,len(listaCampo)-1)
				%><font class=cab><b><%=LitTicketTpv%> : </b></font><font class=cab><%=EncodeForHtml(listaCampo)%></font><br/><%
				encabezado=1
			end if
		end if

		if conref>"" then
			strwhere=strwhere & " d.referencia like '%" & conref & "%' and"
			%><font class=cab><b><%=LitRefListTpv%> : </b></font><font class=cab><%=EncodeForHtml(conref)%></font><br/><%
			encabezado=1
		end if

		if connombre>"" then
			strwhere=strwhere & " d.referencia in (select referencia from articulos with(nolock) where nombre like '%" & connombre & "%' and referencia like '" & session("ncliente") & "%') and"
			%><font class=cab><b><%=LitNomListTpv%> : </b></font><font class=cab><%=EncodeForHtml(connombre)%></font><br/><%
			encabezado=1
		end if

		if tiparticulo >"" then                                                                                                                          
            tiparticuloR = replace(tiparticulo,", ","','")
		    strwhere = strwhere & " a.tipo_articulo in ('" & tiparticuloR & "') and"
            desc_tipoarticulo=NombresEntidades(tiparticulo,"tipos_entidades","codigo","descripcion",session("backendListados"))
			%><font class=cab><b><%=LitTipArticulo%> : </b></font><font class=cab><%=EncodeForHtml(desc_tipoarticulo)%></font><br/><%
			encabezado=1
		end if

		if categoria >"" then
		    categoriaR = replace(categoria,", ","','")
		    strwhere = strwhere & " a.categoria in ('" & categoriaR & "') and"

            categoria = replace(categoria,", ","','")
            categoria = "('" & categoria & "')"
		end if 
		
		if familia_padre > "" then
		    familia_padreR = replace(familia_padre,", ","','")
		    strwhere = strwhere & " a.familia_padre in ('" & familia_padreR & "') and"

            familia_padre = replace(familia_padre,", ","','")
            familia_padre = "('" & familia_padre & "')"
		end if
		
		if familia > "" then
		    familiaR = replace(familia,", ","','")
		    strwhere = strwhere & " a.familia in ('" & familiaR & "') and"

            familia = replace(familia,", ","','")
            familia = "('" & familia & "')"
		end if
		
		' Ponemos las fechas que ponga el usuario (nunca será vacio, en el onchange de las horas, se controlará)
		if dhora = "" then
		    dhora = "00:00"
		end if
		if hhora = "" then
		    hhora = "23:59"
		end if
		strwhere=strwhere & " ti.fecha >= '" & dfecha & " " & dhora & ":00' and ti.fecha <= '" & hfecha & " " & hhora & ":59' and"
		strwhere=strwhere & " d.nticket=ti.nticket and d.referencia=a.referencia "

        select case agruparpor
		    case ucase(LitSubFamilia):  str_select = "a.familia"
            case ucase(LitFamilia):     str_select = "a.familia_padre"
            case ucase(LitCategoria):   str_select = "a.categoria"
        end select

		'para que salgan las familias padre aunque no tengan articulos
		if agruparpor=ucase(LitSubFamilia) or agruparpor=ucase(LitFamilia) or agruparpor=ucase(LitCategoria) then
			listaFam="('"
			cadena="select distinct " & str_select & " as campo " & strfrom & strwhere & " and " & str_select & " is not null order by campo"
			rst.cursorlocation=3
			rst.Open cadena, Session("backendlistados")
			while not rst.eof
				listaFam=listaFam & rst("campo") & "','"
				rst.movenext
			wend
			rst.close
			if listaFam="('" then
				listaFam=""
			else
				listaFam=mid(listaFam,1,len(listaFam)-2) & ")"
			end if
		end if

		'para que salgan las subfamilias de las familias que salen en el select
		if agruparpor=ucase(LitSubFamilia) and familia="XXXXXXXX" then
			lista="('"
			cadena=strfrom & strwhere
			cadena="select distinct a.familia as campo " & cadena & " and a.familia is not null order by campo"
			rst.cursorlocation=3
			rst.Open cadena, Session("backendlistados")
			while not rst.eof
				lista=lista & rst("campo") & "','"
				rst.movenext
			wend
			rst.close
			if lista="('" then
				lista=""
			else
				lista=mid(lista,1,len(lista)-2) & ")"
			end if
			strwhere=strwhere & " and a.familia in (select codigo from familias with(nolock) where codigo like '" &session("ncliente") & "%' and padre in " & lista
			strwhere=strwhere & " union "
			strwhere=strwhere & " select codigo from familias with(nolock) where codigo like '" &session("ncliente") & "%' and codigo in " & lista
			strwhere=strwhere & " )"
		end if

        set command3 = nothing
        set conn3 = Server.CreateObject("ADODB.Connection")
        set command3 =  Server.CreateObject("ADODB.Command")
        conn3.open Session("backendlistados")
        conn3.cursorlocation=3
        command3.ActiveConnection =conn3
        command3.CommandTimeout = 60

        if agruparpor=ucase(LitSubFamilia) and familia & "">"" then
            %><font class=cab><b><%=LitSubFamilia%> : </b></font>
            <font class=cab>
                <%
                    strselectAgrup="select nombre from familias with(nolock) where codigo in " & familia & " "
                    command3.CommandText=strselectAgrup
                    command3.CommandType = adCmdText
                    set rst3=command3.Execute
                    while not rst3.eof
				        listaGrupo=listaGrupo & rst3("nombre")
				        rst3.movenext
                        if not rst3.eof then
                            listaGrupo=listaGrupo & ", "
                        end if
			        wend
                    rst3.close
                    response.Write(listaGrupo)
                %>
            </font><br/><%
			encabezado=1
			strwhere=strwhere & " and a.familia in (select codigo from familias with(nolock) where codigo like '" &session("ncliente") & "%' and codigo in " & familia & " ) "
        elseif agruparpor=ucase(LitFamilia) and familia_padre & "">"" then
            %><font class=cab><b><%=LitFamilia%> : </b></font>
            <font class=cab>
                <%
                    strselectAgrup="select nombre from familias_padre with(nolock) where codigo in " & familia_padre & " "
                    command3.CommandText=strselectAgrup
                    command3.CommandType = adCmdText
                    set rst3=command3.Execute
                    while not rst3.eof
				        listaGrupo=listaGrupo & rst3("nombre")
				        rst3.movenext
                        if not rst3.eof then
                            listaGrupo=listaGrupo & ", "
                        end if
			        wend
                    rst3.close
                    response.Write(listaGrupo)
                %>
            </font><br/><%
			encabezado=1
			strwhere=strwhere & " and a.familia_padre in (select codigo from familias_padre with(nolock) where codigo like '" &session("ncliente") & "%' and codigo in " & familia_padre & " ) "
        elseif agruparpor=ucase(LitCategoria) and categoria & "">"" then
            %><font class=cab><b><%=LitCategoria%> : </b></font>
            <font class=cab>
                <%
                    strselectAgrup="select nombre from categorias with(nolock) where codigo in " & categoria & " "
                    command3.CommandText=strselectAgrup
                    command3.CommandType = adCmdText
                    set rst3=command3.Execute
                    while not rst3.eof
				        listaGrupo=listaGrupo & rst3("nombre")
				        rst3.movenext
                        if not rst3.eof then
                            listaGrupo=listaGrupo & ", "
                        end if
			        wend
                    rst3.close
                    response.Write(listaGrupo)
                %>
            </font><br/><%
			encabezado=1
			strwhere=strwhere & " and a.categoria in (select codigo from categorias with(nolock) where codigo like '" &session("ncliente") & "%' and codigo in " & categoria & " ) "
        end if

        conn3.close
        set conn3       =  nothing
        set command3    =  nothing
        set rst3        =  nothing	
        
		'ahora ponemos los hipervinculos a los listados
		%></td>
		<td class='CELDA' width='15%' align='left' valign="top">
			<%if vert="1" then%>
				<a class='CELDAREFB7' href="javascript:ver_tours();" onmouseover="self.status='<%=LitVerTours%>';return true;" onmouseout="self.status='';return true;"><%=LitListadoTours%></a>
			<%end if
			if (verp="1" or verc="1") and vert="1" then
				%><br/><%
			end if
			if verc="1" then%>
				<a class='CELDAREFB7' href="javascript:ver_comisiones();" onmouseover="self.status='<%=LitVerComisiones%>';return true;" onmouseout="self.status='';return true;"><%=LitListadoComisiones%></a>
			<%end if
			if verp="1" and (verc="1" or vert="1") then
				%><br/><%
			end if
			if verp="1" then%>
				<a class='CELDAREFB7' href="javascript:ver_penalizaciones();" onmouseover="self.status='<%=LitVerPenalizaciones%>';return true;" onmouseout="self.status='';return true;"><%=LitListadoPenalizaciones%></a>
			<%end if%>
		</td>
		</tr></table>
		<%

		strgroup=""
		strorder=" order by "

		select case agruparpor
			case ucase(LitSubFamilia):
				''ricardo 14-11-2007 se ordena tambien por el nombre de la familia padre, ya que si no, pueden salir las familias hijas en distintas paginas
				                        strorder=strorder & "art_cod_categoria,fampadre,f.nombre,"
            case ucase(LitFamilia):     strorder=strorder & "art_cod_categoria,fampadre,f.nombre,"
            case ucase(LitCategoria):   strorder=strorder & "art_cod_categoria,fampadre,f.nombre,"
		end select

		strorder=strorder & "a.nombre,a.referencia,"
		select case mostrarcolum
			case ucase(LitTicketTpv):
				strorder=strorder & "ti.tpv,"
				strgroup=strgroup & "ti.tpv,tpv.descripcion "
				cadena_a_comparar="referencia"
				cadena_a_comparar2="tpv"
				cadena_a_comparar3="nomarticulo"
			case ucase(LitOperador):
				strorder=strorder & "ti.usuario,"
				strgroup=strgroup & "ti.usuario,p.nombre"
				cadena_a_comparar="referencia"
				cadena_a_comparar2="usuario"
				cadena_a_comparar3="nomarticulo"
			case ucase(LitCaja):
				strorder=strorder & "ti.caja,"
				strgroup=strgroup & "ti.caja,c.descripcion"
				cadena_a_comparar="referencia"
				cadena_a_comparar2="caja"
				cadena_a_comparar3="nomarticulo"
			case ucase(LitTicketTienda):
				strorder=strorder & "c.tienda,"
				strgroup=strgroup & "c.tienda,tda.descripcion"
				cadena_a_comparar="referencia"
				cadena_a_comparar2="tienda"
				cadena_a_comparar3="nomarticulo"
		end select
		if mid(strorder,len(strorder),1)="," then
			strorder=mid(strorder,1,len(strorder)-1)
		end if

		if strgroup="" then
			strgroup=""
		else
			strgroup=" group by upper(d.referencia),ti.divisa," & strgroup & ",a.referencia,a.nombre,a.familia,f.nombre,a.familia_padre,a.categoria "
		end if

		if strorder=" order by " then strorder=""

		if encabezado=1 then
			%><hr/><%
		end if
        strGroupIva =""
        strGroupIva = strgroup & ",d.iva,d.nticket"

		MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='139'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='139'", DSNIlion)
		%><input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'>                                
		<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'><%                                     


		strdrop ="if exists (select * from sysobjects where id = object_id('[" & session("usuario") & "]') and sysstat " & _
		" & 0xf = 3) drop table [" & session("usuario") & "]"
		rstAux.open strdrop,Session("backendlistados"),adUseClient,adLockReadOnly
		if rstAux.state<>0 then rstAux.close

		strselect3="CREATE TABLE [" & session("usuario") & "] (num int identity(1,1)"
		strselect3=strselect3 & ",referencia varchar(30),total_cant real,total_imp money, total_imp_iva money,divisa varchar(15)"

		select case mostrarcolum
			case ucase(LitTicketTpv):
		        strselect3=strselect3 & ",tpv varchar(8),nombretpv varchar(50)"
			case ucase(LitOperador):
		        strselect3=strselect3 & ",usuario varchar(20),nompersonal varchar(50)"
			case ucase(LitCaja):
		        strselect3=strselect3 & ",caja varchar(10),nombrecaja varchar(50)"
			case ucase(LitTicketTienda):
		        strselect3=strselect3 & ",tienda varchar(10),nombretienda varchar(50)"
        end select

		strselect3=strselect3 & ",nomarticulo varchar(100),familia varchar(10),nomfamilia varchar(50),fampadre varchar(10),art_cod_categoria varchar(10)"

		strselect3=strselect3 & ")"
		rstAux.open strselect3,Session("backendlistados"),adUseClient,adLockReadOnly
		if rstAux.State<>0 then rstAux.close
		GrantUser session("usuario"), Session("backendlistados")
		strselect3="insert into [" & session("usuario") & "](referencia,total_cant,total_imp, total_imp_iva,divisa " 

		select case mostrarcolum
			case ucase(LitTicketTpv):
		        strselect3=strselect3 & ",tpv,nombretpv"
			case ucase(LitOperador):
		        strselect3=strselect3 & ",usuario,nompersonal"
			case ucase(LitCaja):
		        strselect3=strselect3 & ",caja,nombrecaja"
			case ucase(LitTicketTienda):
		        strselect3=strselect3 & ",tienda,nombretienda"
        end select

		strselect3=strselect3 & ",nomarticulo,familia,nomfamilia,fampadre,art_cod_categoria "
		strselect3=strselect3 & ") "
		strselect3=strselect3 & strselect & strfrom & strwhere & strgroup & strorder
        
        set conn = Server.CreateObject("ADODB.Connection")
        conn.ConnectionTimeout = 300
        conn.CommandTimeout = 300
		conn.open session("backendlistados")
		set rst = conn.execute(strselect3)

        'DBG 30/06/2016 LLamo esta función para crear la tabla temporal para calcular el desglose IVA con los mismos where/group by que la tabla temporal de resumen de tickets TPV 
        CreaTablaTemporalIva

        if rst.State<>0 then rst.close
        conn.close
        set conn = nothing

		'calculamos ahora las referencias que hay
		cadena3="select count(distinct upper(d.referencia)) as regtotales from [" & session("usuario") & "] as d "
		rst.cursorlocation=3
		rst.open cadena3, Session("backendlistados")
		if not rst.eof then
			regtotales=rst("regtotales")
		else
			regtotales=0
		end if
		rst.close

		rst.cursorlocation=3
		rst.Open "select * from [" & session("usuario") & "] order by num", Session("backendlistados")
		if rst.eof then
			%><input type="hidden" name="NumRegs" value="0">
			<script>
			    window.alert("<%=LitMsgNoDocumentos%>");
			    parent.botones.document.location = "listado_tickets_bt.asp?mode=param";
				document.location="listado_tickets.asp?mode=param";
			</script><%
		else

			%><input type="hidden" name="NumRegs" value="<%=EncodeForHtml(regtotales)%>" /><%

			'Calculos de páginas--------------------------
			lote=limpiaCadena(Request.QueryString("lote"))
			if lote="" then
				lote=1
			end if
			sentido=limpiaCadena(Request.QueryString("sentido"))

			lotes=regtotales/MAXPAGINA
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
			
			%>                                                                                       
            <input type="hidden" name="lotesPag" value="<%=EncodeForHtml(lotes)%>" />
            <%

			rst.PageSize=MAXPAGINA
			rst.AbsolutePage=lote
			'-----------------------------------------

			NavPaginas lote,lotes,campo,criterio,texto,1

			%><table width="100%" style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%

			VinculosPagina(MostrarArticulos)=1:VinculosPagina(MostrarPersonal)=1
			VinculosPagina(MostrarTiendas)=1:VinculosPagina(MostrarArticulos)=1
			CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina

			mb=d_lookup("codigo","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",Session("backendlistados"))
            '' DBG 04/07/2016 Extraigo el factor de cambio para la mb( Moneda Base de la aplicación) para usarla en la tabla de Desglose Iva posteriormente
            mbFactorCambioDestino = d_lookup("FACTCAMBIO","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",Session("backendlistados"))
			dec_mb=d_lookup("ndecimales","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",Session("backendlistados"))
			abrev_mb=d_lookup("abreviatura","divisas","moneda_base<>0 and codigo like '" & session("ncliente") & "%'",Session("backendlistados"))
			precio_ticket=d_lookup("valor_ticket","configuracion","nempresa='" & session("ncliente") & "'",Session("backendlistados"))
			
			
''ricardo 23/4/2008 se cambia el from y el where
strfrom=" from [" & session("usuario") & "] as tmp "
strwhere=" where tmp.referencia like '" & session("ncliente") & "%' "

			dim lista_columnas()
			dim contador_columnas

			if not rst.eof then
				cadena=strfrom & strwhere
				select case mostrarcolum
					case ucase(LitTicketTpv):
						cadena="select distinct tmp.tpv as campo,tmp.nombretpv as descripcion " & cadena & " order by tmp.tpv"
						tdocumento="ticket"
					case ucase(LitOperador):
						cadena="select distinct tmp.usuario as campo,tmp.nompersonal as descripcion " & cadena & " order by tmp.usuario"
						tdocumento="operador"
					case ucase(LitCaja):
						cadena="select distinct tmp.caja as campo,tmp.nombrecaja as descripcion " & cadena & " order by tmp.caja"
						tdocumento="caja"
					case ucase(LitTicketTienda):
						cadena="select distinct tmp.tienda as campo,tmp.nombretienda as descripcion " & cadena & " order by tmp.tienda"
						tdocumento="tienda"
				end select

				rstcursor.cursorlocation=3
				rstcursor.open cadena, Session("backendlistados")
				if not rstcursor.eof then
					redim preserve lista_columnas(rstcursor.recordcount+1,4)
				end if
				contador_columnas=1
				while not rstcursor.eof
					lista_columnas(contador_columnas,1)=rstcursor("campo")
					lista_columnas(contador_columnas,2)=rstcursor("descripcion")
					lista_columnas(contador_columnas,3)=""
					lista_columnas(contador_columnas,4)=""
					rstcursor.movenext
					contador_columnas=contador_columnas+1
				wend
				rstcursor.close
				if contador_columnas>1 then
					contador_columnas=contador_columnas-1
				end if
                'DIBUJA LA CABECERA DEL LISTADO (CONTIENE LAS COLUMNAS (TPV,OPERADOR...))    
				if contador_columnas>=1 then
					CabeceraListado lista_columnas,contador_columnas,tdocumento,mostrarcolum,mostrardesc,mostrarinfo
				end if
			end if

            '-----------------------------------------------------------------------------------------------------

            ultima_familia=rst("familia") 
            ultima_fampadre=rst("fampadre") 
            ultima_categoria=rst("art_cod_categoria") 

			valor_old=ucase(rst(cadena_a_comparar))
			nombre_old=rst(cadena_a_comparar3)
			divisa_ultima=rst("divisa")

			sumaimporte=0
			sumacantidad=0

			if agruparpor=ucase(LitSubFamilia) or agruparpor=ucase(LitFamilia) or agruparpor=ucase(LitCategoria) then 
				DrawFila ""
				CloseFila
				DrawFila ""
				CloseFila
                
                if agruparpor=ucase(LitCategoria) and rst("art_cod_categoria")&"">"" then
                    nomcat=d_lookup("nombre","categorias","codigo='" & rst("art_cod_categoria")  & "'",Session("backendlistados"))
                else
                    if agruparpor=ucase(LitFamilia) and rst("fampadre")&"">""  then
                        nomfampad=d_lookup("nombre","familias_padre","codigo='" & rst("fampadre") & "'",Session("backendlistados"))
                        nomcat=d_lookup("nombre","categorias","codigo='" & rst("art_cod_categoria")  & "'",Session("backendlistados"))
                    else
                        if agruparpor=ucase(LitSubFamilia) and rst("familia")&"">""  then
                             nomfam=rst("nomfamilia") & ""   
                             nomfampad=d_lookup("nombre","familias_padre","codigo='" & rst("fampadre") & "'",Session("backendlistados"))
                             nomcat=d_lookup("nombre","categorias","codigo='" & rst("art_cod_categoria") & "'",Session("backendlistados"))
                        else
                            DrawCelda2 "dato style='font-size:10pt;padding-left:10px;padding-top:5px' colspan='" & tamano_colspan+2 & "' ", "left bottom",true, "Sin categoría/familia/subfamilia" 
                        end if
                    end if
                end if

                if nomcat & "" > "" then
                    DrawFila color_fondo
						tamano_colspan=3 + contador_columnas
						DrawCelda2 "dato style='font-size:10pt;padding-left:10px;padding-top:5px' colspan='" & tamano_colspan & "' ", "left",true, LitCategoria & ": " & trimCodEmpresa(rst("art_cod_categoria")) & " - " & nomcat
					CloseFila
                    nomcat=""
                end if

                if nomfampad & "" > "" then
                    DrawFila color_fondo
						tamano_colspan=3 + contador_columnas
						DrawCelda2 "dato style='font-size:9pt;padding-left:15px' colspan='" & tamano_colspan & "' ", "left",true, LitFamilia & ": " & trimCodEmpresa(rst("fampadre")) & " - " & nomfampad
					CloseFila
                    nomfampad=""
                end if

                if nomfam & "" > "" then
                    DrawFila color_fondo
						tamano_colspan=3 + contador_columnas
						 DrawCelda2 "dato style='font-size:8pt;padding-left:20px' colspan='" & tamano_colspan & "' ", "left",true, LitSubFamilia & " : " & trimCodEmpresa(rst("familia")) & " - " & nomfam
					CloseFila
                    nomfam=""
                end if
			end if
			
			fila=1
		    while not rst.EOF and fila<=MAXPAGINA
				CheckCadena rst(cadena_a_comparar)

				'Seleccionar el color de la fila.
				if ((fila+1) mod 2)=0 then
					color=color_blau
				else
					color=color_terra
				end if
				
				if valor_old<>rst(cadena_a_comparar) then
			        if detalle="conDet" then
					    DrawFila color
						    DrawCelda2 "dato width='15%'", "left", false,Escribirhref(valor_old,"articulo",nombre_old)
						    suma=0
						    importe=0
						    dto=0
						    for i=1 to contador_columnas

							    cantidad_a_poner=""
							    if lista_columnas(i,3)="0" then
								    cantidad_a_poner=""
							    else
								    if lista_columnas(i,3) & "">"" then
									    cantidad_a_poner=formatnumber(lista_columnas(i,3),dec_cant,-1,0,-1)
									    suma=suma + cantidad_a_poner
								    end if
							    end if


							    importe_a_poner=""
							    if lista_columnas(i,4)="0" then
								    importe_a_poner=""
							    else
								    if lista_columnas(i,4) & "">"" then
									    importe_a_poner=formatnumber(lista_columnas(i,4),dec_mb,-1,0,-1)
									    importe=importe + importe_a_poner
								    end if
							    end if

							    if mostrarinfo=ucase(LitCantidad) then
								    DrawCelda2 "dato align='right' ", "left", false,cantidad_a_poner
							    else
								    texto_a_poner=importe_a_poner
								    if comprobar_perso(dtofam,ultima_familia) then
									    if importe_a_poner & "">"" and cantidad_a_poner & "">"" then
										    dto_a_poner=(cantidad_a_poner*precio_ticket)-importe_a_poner
										    dto_a_poner=formatnumber(dto_a_poner,decpor,-1,0,-1)
										    texto_a_poner=texto_a_poner & "<br/>" & LitDescuento & " : " & dto_a_poner
									    end if
								    end if
								    DrawCelda2 "dato align='right' ", "left", false,texto_a_poner
							    end if
						    next
						    if suma & "">"" then
							    cantidad_a_poner=formatnumber(suma,dec_cant,-1,0,-1)
						    end if
						    DrawCelda2 "dato align='right' width='10%' ", "left", false,cantidad_a_poner
						    if importe & "">"" then
							    importe_a_poner=formatnumber(importe,dec_mb,-1,0,-1)
						    end if
						    texto_a_poner=importe_a_poner
						    if comprobar_perso(dtofam,ultima_familia) then
							    if importe_a_poner & "">"" and cantidad_a_poner & "">"" then
								    dto_a_poner=(cantidad_a_poner*precio_ticket)-importe_a_poner
								    dto_a_poner=formatnumber(dto_a_poner,decpor,-1,0,-1)
								    texto_a_poner=texto_a_poner & "<br/>" & LitTotalDto & " : " & dto_a_poner
							    end if
						    end if
						    DrawCelda2 "dato align='right' width='10%' ", "left", false,texto_a_poner
						    suma=0
						    importe=0
					    CloseFila
				    end if
					for i=1 to contador_columnas
						lista_columnas(i,3)=""
						lista_columnas(i,4)=""
					next
					fila=fila+1
				end if
				for i=1 to contador_columnas
					if lista_columnas(i,1) & ""=(rst(cadena_a_comparar2)) & "" then
						'se ponen para sumar por si hay divisas distintas con el mismo articulo
						if lista_columnas(i,3) & "">"" then
							lista_columnas(i,3)=lista_columnas(i,3) + rst("total_cant")
						else
							lista_columnas(i,3)=rst("total_cant")
						end if
						'se ponen para sumar por si hay divisas distintas con el mismo articulo
						if lista_columnas(i,4) & "">"" then
							lista_columnas(i,4)=lista_columnas(i,4) + cambiodivisa(rst("total_imp"),rst("divisa"),mb)
						else
							lista_columnas(i,4)=cambiodivisa(rst("total_imp"),rst("divisa"),mb)
						end if
					end if
				next


                '------------------------------------------------------------------------------------------------------------------------------------

				if agruparpor=ucase(LitSubFamilia) or agruparpor=ucase(LitFamilia) or agruparpor=ucase(LitCategoria) then 
                    ' PINTA TOTALES EN MEDIO
                    cadena=strfrom & strwhere
                    if agruparpor=ucase(LitCategoria) and (ultima_categoria & "")<>(rst("art_cod_categoria") & "") then
                        if ultima_categoria&""="" then 
                            DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"","",""
                        else
                            DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"",ultima_categoria,"categoria"
                        end if
                    else
                         if agruparpor=ucase(LitFamilia) and (ultima_fampadre & "")<>(rst("fampadre") & "") then
                            if ultima_fampadre&""="" and ultima_categoria&""="" then  
                                DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"","",""
                            else 
                                DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,ultima_fampadre,"","fampadre"
                                if  (rst("art_cod_categoria") & "")<>(ultima_categoria & "") then
                                    DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"",ultima_categoria,"categoria"
                                end if
                            end if
                        else
                            if agruparpor=ucase(LitSubFamilia) and (ultima_familia & "")<>(rst("familia") & "") then
                                if ultima_familia&""="" and ultima_fampadre&""="" and ultima_categoria&""="" then  
                                    DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"","",""
                                else                           
                                    DibujaLineaSubTotales lista_columnas,contador_columnas,strfrom & strwhere,ultima_familia,mostrarcolum,mostrarinfo,"","","familia"
                                    if  rst("fampadre") & "" <> ultima_fampadre & "" then
                                        DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,ultima_fampadre,"","fampadre"
                                    end if
                                    if  rst("art_cod_categoria") & "" <> ultima_categoria & "" then
                                        DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"",ultima_categoria,"categoria"
                                    end if
                                end if
                                
                            end if
                        end if
                    end if

					DrawFila ""
					CloseFila
					DrawFila ""
					CloseFila

                    ' PINTA LINEAS CABECERAS
                    if agruparpor=ucase(LitCategoria) and (ultima_categoria & "")<>(rst("art_cod_categoria") & "") then 
                        nomcat=d_lookup("nombre","categorias","codigo='" & rst("art_cod_categoria") & "'",Session("backendlistados"))
                    else
                        if agruparpor=ucase(LitFamilia) and (ultima_fampadre & "")<>(rst("fampadre") & "") then 
                            nomfampad=d_lookup("nombre","familias_padre","codigo='" & rst("fampadre") & "'",Session("backendlistados"))
                            if  (ultima_categoria & "")<>(rst("art_cod_categoria") & "") then
                                nomcat=d_lookup("nombre","categorias","codigo='" & rst("art_cod_categoria") & "'",Session("backendlistados"))
                            end if
                        else
                            if agruparpor=ucase(LitSubFamilia) and (ultima_familia & "")<>(rst("familia") & "") then
                                nomfam=rst("nomfamilia") & ""   
                                if  (ultima_fampadre & "")<>(rst("fampadre") & "") then
                                    nomfampad=d_lookup("nombre","familias_padre","codigo='" & rst("fampadre") & "'",Session("backendlistados"))
                                end if
                                if  (ultima_categoria & "")<>(rst("art_cod_categoria") & "") then
                                    nomcat=d_lookup("nombre","categorias","codigo='" & rst("art_cod_categoria") & "'",Session("backendlistados"))
                                end if
                            end if
                        end if
                    end if 

                    if nomcat & "" > "" then
                        DrawFila color_fondo
						    tamano_colspan=3 + contador_columnas
                            DrawCelda2 "dato style='font-size:10pt;padding-left:10px;padding-top:5px' colspan='" & tamano_colspan & "' ", "left",true, LitCategoria & ": " & trimCodEmpresa(rst("art_cod_categoria")) & " - " & nomcat
					    CloseFila
                        nomcat=""
                        perultima_categoria=ultima_categoria
                    end if

                    if nomfampad & "" > "" then
                        DrawFila color_fondo
						    tamano_colspan=3 + contador_columnas
                            DrawCelda2 "dato style='font-size:9pt;padding-left:15px' colspan='" & tamano_colspan & "' ", "left",true, LitFamilia & ": " & trimCodEmpresa(rst("fampadre")) & " - " & nomfampad
					    CloseFila
                        nomfampad=""
				        perultima_fampadre=ultima_fampadre
                    end if

                    if nomfam & "" > "" then
                        DrawFila color_fondo
						    tamano_colspan=3 + contador_columnas
                            DrawCelda2 "dato style='font-size:8pt;padding-left:20px' colspan='" & tamano_colspan & "' ", "left",true, LitSubFamilia & " : " & trimCodEmpresa(rst("familia")) & " - " & nomfam
					    CloseFila
                        nomfam=""
                        perultima_familia=ultima_familia
                    end if

				end if

				valor_old=ucase(rst(cadena_a_comparar))
				nombre_old=rst(cadena_a_comparar3)
				divisa_ultima=rst("divisa") 

				ultima_familia=rst("familia")
				ultima_fampadre=rst("fampadre")
                ultima_categoria=rst("art_cod_categoria")
                 
				rst.movenext
			wend
					if ((fila+1) mod 2)=0 then
						color=color_blau
					else
						color=color_terra
					end if
				    if detalle="conDet" then
					    DrawFila color
						    DrawCelda2 "dato width='15%'", "left", false,Escribirhref(valor_old,"articulo",nombre_old)
						    suma=0
						    importe=0
						    dto=0
						    dtot=0
						    for i=1 to contador_columnas

								    cantidad_a_poner=""
								    if lista_columnas(i,3)="0" then
									    cantidad_a_poner=""
								    else
									    if lista_columnas(i,3) & "">"" then
										    cantidad_a_poner=formatnumber(lista_columnas(i,3),dec_cant,-1,0,-1)
									    end if
								    end if

								    importe_a_poner=""
								    if lista_columnas(i,4)="0" then
									    importe_a_poner=""
								    else
									    if lista_columnas(i,4) & "">"" then
										    importe_a_poner=formatnumber(lista_columnas(i,4),dec_mb,-1,0,-1)
									    end if
								    end if
								    if importe_a_poner & "">"" then
									    importe_a_poner=cambiodivisa(importe_a_poner,divisa_ultima,mb)
								    end if

							    if mostrarinfo=ucase(LitCantidad) then
								    DrawCelda2 "dato align='right' ", "left", false,cantidad_a_poner
							    else
								    texto_a_poner=importe_a_poner
								    if comprobar_perso(dtofam,ultima_familia) then
									    if importe_a_poner & "">"" and cantidad_a_poner & "">"" then
										    dto_a_poner=(cantidad_a_poner*precio_ticket)-importe_a_poner
										    dto_a_poner=formatnumber(dto_a_poner,decpor,-1,0,-1)
										    texto_a_poner=texto_a_poner & "<br/>" & LitDescuento & " : " & dto_a_poner
									    end if
								    end if
								    DrawCelda2 "dato align='right' ", "left", false,texto_a_poner
							    end if
							    if lista_columnas(i,3) & "">"" then
								    suma=suma + lista_columnas(i,3)
							    end if
							    if lista_columnas(i,4) & "">"" then
								    importe=importe + cambiodivisa(lista_columnas(i,4),divisa_ultima,mb)
							    end if
						    next
						    if suma & "">"" then
							    cantidad_a_poner=formatnumber(suma,dec_cant,-1,0,-1)
						    end if
						    DrawCelda2 "dato align='right' width='10%' ", "left", false,cantidad_a_poner
						    if importe & "">"" then
							    importe_a_poner=formatnumber(importe,dec_mb,-1,0,-1)
						    end if
						    texto_a_poner=importe_a_poner
						    if comprobar_perso(dtofam,ultima_familia) then
							    if importe_a_poner & "">"" and cantidad_a_poner & "">"" then
								    dto_a_poner=(cantidad_a_poner*precio_ticket)-importe_a_poner
								    dto_a_poner=formatnumber(dto_a_poner,decpor,-1,0,-1)
								    texto_a_poner=texto_a_poner & "<br/>" & LitTotalDto & " : " & dto_a_poner
							    end if
						    end if
						    DrawCelda2 "dato align='right' width='10%' ", "left", false,texto_a_poner
						    suma=0
						    importe=0
						    dto=0
					    CloseFila
					end if
					for i=1 to contador_columnas
						lista_columnas(i,3)=""
						lista_columnas(i,4)=""
					next
					fila=fila+1

                if agruparpor=ucase(LitSubFamilia) or agruparpor=ucase(LitFamilia) or agruparpor=ucase(LitCategoria) then 
                    ' PINTA TOTALES AL FINAL
                    cadena=strfrom & strwhere
                    if agruparpor=ucase(LitCategoria) and (ultima_categoria & "")<>(perultima_categoria & "") then
                        if ultima_categoria&""="" then 
                            DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"","",""
                        else
                            DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"",ultima_categoria,"categoria"
                        end if
                    else
                         if agruparpor=ucase(LitFamilia) and (ultima_fampadre & "")<>(perultima_fampadre & "") then
                            if ultima_fampadre&""="" and ultima_categoria&""="" then  
                                DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"","",""
                            else 
                                DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,ultima_fampadre,"","fampadre"
                                if  (perultima_categoria & "")<>(ultima_categoria & "") then
                                    DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"",ultima_categoria,"categoria"
                                end if
                            end if
                        else
                            if agruparpor=ucase(LitSubFamilia) and (ultima_familia & "")<>(perultima_familia & "") then
                                if ultima_familia&""="" and ultima_fampadre&""="" and ultima_categoria&""="" then  
                                    DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"","",""
                                else                    
                                    DibujaLineaSubTotales lista_columnas,contador_columnas,strfrom & strwhere,ultima_familia,mostrarcolum,mostrarinfo,"","","familia"
                                    if  perultima_fampadre & "" <> ultima_fampadre & "" then
                                        DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,ultima_fampadre,"","fampadre"
                                    end if
                                    if  perultima_categoria & "" <> ultima_categoria & "" then
                                        DibujaLineaSubTotales lista_columnas,contador_columnas,cadena,"",mostrarcolum,mostrarinfo,"",ultima_categoria,"categoria"
                                    end if
                                end if
                            end if
                        end if
                    end if
				end if
			if cint(lote)=cint(lotes) then
				cambiar_literal=0
				if agruparpor & "">"" then
					cambiar_literal=1
				end if
				cadena=strfrom & strwhere
				DibujaLineaTotales lista_columnas,contador_columnas,cadena,cambiar_literal,mostrarcolum,mostrarinfo
                ''DBG 01/07/2016 Aqui llama a mostrar el desglose de IVA segun el valor del checkbox de desglose IVA
                if (TmpOpcDesgloseiva = "on") then
                    CalculaIva
                end if
			end if

			%></table><%

			NavPaginas lote,lotes,campo,criterio,texto,2

			rst.close
		end if
	end if

   %>

	<input type="hidden" name="nRegsImp" value="<%=fila-1%>" />

   </form>
	<%
	if mode="param" then
		%><iframe id='frComprLimitColum' name="fr_ComprLimitColum" style='display:none' src='listado_tickets_comp.asp?mode=param' width='500' height='200' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
	end if

end if%>
<%
set rst=nothing
set rst2=nothing
set rstAux=nothing
set rstcursor=nothing
set rstDesgloseIva = nothing
set rstDesgloseIvaRecorreBucle = nothing
%>
<%
if Not(connRound IS Nothing) then
    connRound.close
    set connRound = Nothing
end if%>
</body>
</html>