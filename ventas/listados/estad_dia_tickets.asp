<%@ Language=VBScript %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
    <title><%=LitTitulo%></title>
    <meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
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
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file= "../../CatFamSubResponsive.inc"-->
<!--#include file= "../../styles/formularios.css.inc"-->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" --> 

<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<body onload="inicio();" class="BODY_ASP" >

<%'MPC 16/11/2007 CAMBIO DSN PARA LISTADOS

sub CabeceraListado()%>
    <table border="0" cellspacing="3" cellpadding="1">
        <%DrawFila color_fondo
            DrawCelda "ENCABEZADOC","10","",0,LitFecha
            DrawCelda "ENCABEZADOC","10","",0,LitTickets
            DrawCelda "ENCABEZADOC","10","",0,LitVentaMedia
            DrawCelda "ENCABEZADOC","10","",0,LitVentaTotal
            DrawCelda "ENCABEZADOC","10","",0,LitArticulosDistintos
            DrawCelda "ENCABEZADOC","10","",0,LitTotalArticulos
            '**RGU 19/1/2007
            if vben&""="0" then
            else
                DrawCelda "ENCABEZADOC","10","",0,LitBeneficioCoste
                DrawCelda "ENCABEZADOC","10","",0,LitBeneficioVenta
            end if
            '**RGU**
    CloseFila
end sub

%>
 <form name="estad_dia_tickets" method="post">
 <%
 'leemos los parámetros de la página
  mode = enc.EncodeForJavascript(Request.QueryString("mode"))

  if request.querystring("Dfecha")>"" then
	TmpDfecha=limpiaCadena(request.querystring("Dfecha"))
  else
	TmpDfecha=limpiaCadena(request.Form("Dfecha"))
  end if
    if TmpDfecha &""="" then
        TmpDfecha=limpiaCadena(request.querystring("s"))
    end if

  if request.querystring("Hfecha")>"" then
	TmpHfecha=limpiaCadena(request.querystring("Hfecha"))
  else
	TmpHfecha=limpiaCadena(request.Form("Hfecha"))
  end if
    if TmpHfecha &""="" then
        TmpHfecha=limpiaCadena(request.querystring("t"))
    end if

  if request.querystring("almacen")&"">"" then
		almacen=limpiaCadena(request.querystring("almacen"))
	else
		almacen=limpiaCadena(request.form("almacen"))
  end if

  if request.querystring("iva")&"">"" then
		iva=limpiaCadena(request.querystring("iva"))
	else
		iva=limpiaCadena(request.Form("iva"))
  end if

  if request.querystring("ndoc")&"">"" then
  		tpv=limpiaCadena(request.querystring("ndoc"))
  else
  		tpv=limpiaCadena(request.Form("tpv"))
  end if

  'FLM:20090617:parametro de temporadas
  temporadas=limpiaCadena(request.Form("temporadas"))
  'FLM:20090617:parametro de numero de temporadas
  numTemp=limpiaCadena(request.Form("numTemp"))
  
  %>
<script type="text/javascript">
<%
    set rstTemporadas = Server.CreateObject("ADODB.Recordset")
    rstTemporadas.cursorlocation=3
    rstTemporadas.open "select t.codigo, t.descripcion, t.f_min, t.f_max from temporadas t with(nolock) where t.codigo like '" & session("ncliente") & "%' ",Session("backendlistados")

    if not rstTemporadas.Eof then
        response.write "var temporadas = new Array(['" & rstTemporadas("codigo") & "','" & rstTemporadas("descripcion") & "','" & rstTemporadas("f_min") & "','" & rstTemporadas("f_max") & "']"
        rstTemporadas.moveNext
        while not rstTemporadas.eof
            response.write ",['" & rstTemporadas("codigo") & "','" & rstTemporadas("descripcion") & "','" & rstTemporadas("f_min") & "','" & rstTemporadas("f_max") & "']"
            rstTemporadas.moveNext
        wend
        response.write ");"
    end if
    rstTemporadas.close
    set rstTemporadas= NOTHING
%>

    //FLM:20090617:Carga las temporadas según el intervalo de fechas indicados.
    function cargaTemporadas(){
        var i,count=0;
        var tempCargadas="<%=enc.EncodeForJavascript(temporadas)%>,";                                    
        document.estad_dia_tickets.temporadas.options.length =null;
        finicio=document.estad_dia_tickets.Dfecha.value;
        ffin=document.estad_dia_tickets.Hfecha.value;
        for(i=0;i<temporadas.length;i++){    
            if((DiferenciaTiempo(temporadas[i][2],finicio,"dias")>=0 && DiferenciaTiempo(ffin,temporadas[i][2],"dias")>=0) || (DiferenciaTiempo(temporadas[i][3],finicio,"dias")>=0 && DiferenciaTiempo(ffin,temporadas[i][3],"dias")>=0) || (DiferenciaTiempo(temporadas[i][2],finicio,"dias")<=0 && DiferenciaTiempo(ffin,temporadas[i][3],"dias")<=0) ){
                //Si ha sido seleccionada antes, la cargamos.
                sel=(tempCargadas.indexOf(temporadas[i][0]+",")!=-1?"true":"")
                document.estad_dia_tickets.temporadas.options[document.estad_dia_tickets.temporadas.options.length]=new Option(temporadas[i][1],temporadas[i][0],"",sel);
            
            }
        }
        //Para saber el numero de elementos que hay  cargados en el select lo guardo en un input.
        document.estad_dia_tickets.numTemp.value=document.estad_dia_tickets.temporadas.options.length;
     
        //Añado elemento vacío.
        document.estad_dia_tickets.temporadas.options[document.estad_dia_tickets.temporadas.options.length]=new Option("","","","");
    }

    //FLM:20090618:funcion que pinta las temporadas
    function pintaTemp(){
        var out="",count=0;
        var tempCargadas=document.estad_dia_tickets.temporadas.value+",";
        for(i=0;i<temporadas.length;i++){
            if(tempCargadas.indexOf(temporadas[i][0]+",")!=-1){
                out+=temporadas[i][1]+",";
                count++;
            }   
            
        }
        //FLM:20090814:se quiere mostrar siempre todas las temporadas.
        return out.substring(0,(out.length)-1);
    }

    function inicio(){
        <%if enc.EncodeForHtmlAttribute(null_s(request.querystring("mode")))="param" then %>                                               
            cargaTemporadas();
        <%end if %>
    }
</script>
    <input type="hidden" name="tpv" value="<%=enc.EncodeForHtmlAttribute(null_s(tpv))%>" />
  <%

    if tpv&"">"" then
        if TmpDfecha & ""="" then
            TmpDfecha=day(now)&"/"&month(now)&"/"&year(now)
        end if
        TmpHfecha=TmpDfecha
        iva="on"
    end if

  dim vben, alm
  obtenerParametros "estad_dia_tickets"

  if tpv&"">"" then
  	vben="0"
  end if

  PintarCabecera "estad_dia_tickets.asp"
  WaitBoxOculto LitEsperePorFavor

  if mode="browse" then 'si venimos en mode browse es porqué tenemos varias páginas y el usuario le ha dado a siguiente, luego estamos
    mode="imp"          'en mode imp
  end if

  if mode="param" then 'pintamos la pantalla para elegir los parámetros para obtener el listado
   set rstSelect = Server.CreateObject("ADODB.Recordset")  
     
        DrawDiv "1","",""
        DrawLabel "","",LitDesdeFecha
        DrawInput "", "", "Dfecha' onchange='cargaTemporadas();", iif(TmpDfecha>"",TmpDfecha,"01/01/" & year(date)), ""
        DrawCalendar "Dfecha"
        CloseDiv
        
        DrawDiv "1","",""
        DrawLabel "","",LitHastaFecha    
        DrawInput "", "", "Hfecha' onchange='cargaTemporadas();", iif(TmpHfecha>"",TmpHfecha,day(date) & "/" & month(date) & "/" & year(date)), ""
        DrawCalendar "Hfecha"
        CloseDiv
	 
		if alm&"">"" then
			StrIn=StrIn&" and codigo in "& replace(replace(replace(alm,",","','"),"(","('"),")","')")
			rstSelect.open "select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%' "&StrIn&"",Session("backendlistados"),adOpenKeyset,adLockOptimistic
			DrawSelectCeldaAll "","","",0,LitAlmacen,"almacen",rstSelect,almacen,"codigo","descripcion","",""
			rstSelect.close
		else
			rstSelect.open "select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%' "&StrIn&"",Session("backendlistados"),adOpenKeyset,adLockOptimistic
			DrawSelectCelda "","","",0,LitAlmacen,"almacen",rstSelect,almacen,"codigo","descripcion","",""
			rstSelect.close
		end if
	   
       EligeCelda "check","add","","","",0,LitIva,"iva",0, ""
	 
	 'FLM:20090617:añado temporadas.
       DrawDiv "1","",""
       DrawLabel "","",LitTemporada%><select class="width60" name="temporadas" multiple=multiple size="5"></select> 
	        <input type="hidden" name="numTemp" value="0" />
	 <%
       CloseDiv
	 
    elseif mode="imp" then 'pintamos el listado
   'primero pintamos los datos con los que hemos sacado el listado%>                                                          
   <input type="hidden" name="Dfecha" value="<%=enc.EncodeForHtmlAttribute(null_s(TmpDfecha))%>"/>
   <input type="hidden" name="Hfecha" value="<%=enc.EncodeForHtmlAttribute(null_s(TmpHfecha))%>"/>
   <input type="hidden" name="almacen" value="<%=enc.EncodeForHtmlAttribute(null_s(almacen))%>"/>
   <input type="hidden" name="iva" value="<%=enc.EncodeForHtmlAttribute(null_s(iva))%>"/>
   <input type="hidden" name="temporadas" value="<%=enc.EncodeForHtmlAttribute(null_s(temporadas))%>"/>
   <input type="hidden" name="numTemp" value="<%=enc.EncodeForHtmlAttribute(null_s(numTemp))%>" />

   <table width='100%' cellspacing="1" cellpadding="1">
   <tr>
   <td class=CELDARIGHT bgcolor="">
		<%fdesde=TmpDfecha                                                                                   
		fhasta=TmpHfecha
		fhasta=day(fhasta) & "/" & month(fhasta) & "/" & year(fhasta)
		if fdesde>"" then
            if fhasta>"" then
                %><%=LitPeriodoFechas%> : <%=enc.EncodeForHtmlAttribute(null_s(fdesde))%> - <%=enc.EncodeForHtmlAttribute(null_s(fhasta))%><%
            else
                %><%=LitPeriodoFechas%> : <%=LitDesde%>&nbsp;<%=enc.EncodeForHtmlAttribute(null_s(fdesde))%><%
            end if
        else
            if fhasta>"" then
                %><%=LitPeriodoFechas%> : <%=LitHasta%>&nbsp;<%=enc.EncodeForHtmlAttribute(null_s(fhasta))%><%
            end if
        end if%>
   </td>
   </tr>
   </table>
   <table width="100%"><tr><td width="85%" align="left" valign="top">
   <hr/>
<%
   'mostramos los parámetros de la búsqueda
   if almacen & "" > "" then
    'obtenemos el nombre del almacen
	descAlmacen=d_lookup("descripcion","almacenes","codigo='" & almacen & "' and codigo like '" & session("ncliente") & "%'",Session("backendlistados"))
    %><tr><td>
    <font class="cab"><b><%=LitAlmacen%> : </b></font><font class="cab"><%=enc.EncodeForHtmlAttribute(null_s(ucase(descAlmacen)))%></font><br/>
    </td></tr>
    <%end if
    'FLM:20090618:ponemos las temporadas
    if temporadas&""<>"" then%>
    <tr>
        <td class="CELDA7"><font class="cab"><b><%=LitTemporada%>:&nbsp;</b></font><script type="text/javascript">document.write(pintaTemp());</script></td>        
    </tr>
   <%end if
   'el iva lo ponemos siempre%>
   <tr><td>
   <%if iva="on" then%>
    <font class="cab"><b><%=LitConIva%></b></font><br/>
   <%else%>
    <font class="cab"><b><%=LitSinIva%></b></font><br/>
   <%end if%>
   <hr/>
    </td></tr>
    
	</table>
    <%

'   CabeceraListado
   'preparamos los parámetros para pasárselas al procedimiento
   dfecha = fdesde & " 00:00:00"
   hfecha = fhasta & " 23:59:59"

   if iva="on" then
    piva = 1
   else
    piva = 0
   end if
   almacen = almacen & ""

   empresa = session("ncliente") & "%"

   if tpv&"">"" then
	set rstAux = Server.CreateObject("ADODB.Recordset")
	strselect="select ti.almacen from tiendas ti with(nolock), cajas c with(nolock), tpv t with(nolock) "
	strselect=strselect&" where ti.codigo like '"&empresa&"' and c.codigo like '"&empresa&"' and t.tpv like '"&empresa&"' "
	strselect=strselect&" and t.tpv='"&session("ncliente")&tpv&"' and c.codigo=t.caja and c.tienda= ti.codigo "
    rstAux.cursorlocation=3
	rstAux.open strselect , Session("backendlistados")
	if not rstAux.eof then
		almacen=rstAux("almacen")
	else
		%><script language="javascript" type="text/javascript">
		      window.alert("<%=LitMsgTpvErr%>");
		      window.top.close();
		</script><%
	end if
	rstAux.close
    set rstAux=NOTHING

   end if

    'FLM:20090617:temporadas
    if temporadas&""<>"" then
        temporadas="'"&replace(temporadas,", ","','")&"'"        
    end if

   'lanzamos el procedimiento para obtener el listado
   set rst = Server.CreateObject("ADODB.Recordset")
   set conn = Server.CreateObject("ADODB.Connection")
   set command =  Server.CreateObject("ADODB.Command")

   conn.open Session("backendlistados")

   command.ActiveConnection =conn
   command.CommandTimeout = 0
   command.CommandText="listadoEstadisticasVentaPorTickets"
   command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
   command.Parameters.Append command.CreateParameter("@fechaInicio",adVarchar,adParamInput,50,dfecha)
   command.Parameters.Append command.CreateParameter("@fechaFin",adVarChar,adParamInput,50,hfecha)
   command.Parameters.Append command.CreateParameter("@almacen",adVarChar,adParamInput,50,almacen) 'mayúsculas o minúsculas!!!!
   command.Parameters.Append command.CreateParameter("@iva",adInteger,adParamInput,50,piva)
   command.Parameters.Append command.CreateParameter("@empresa", adVarChar,adParamInput,50,empresa)
   command.Parameters.Append command.CreateParameter("@temporadas", adVarChar,adParamInput,8000,temporadas)
   command.Parameters.Append command.CreateParameter("@usuario", adVarChar,adParamInput,30,session("usuario"))
   conn.cursorlocation=3
   set rst=command.Execute
   
   'preparamos la paginación
   MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='413'", DSNIlion)
   MAXPDF=d_lookup("maxpdf", "limites_listados", "item='413'", DSNIlion)%>
   <input type='hidden' name='maxpdf' value='<%=enc.EncodeForHtmlAttribute(null_s(MAXPDF))%>'/>                                                
   <input type='hidden' name='maxpagina' value='<%=enc.EncodeForHtmlAttribute(null_s(MAXPAGINA))%>'/>                                               
   <%if rst.eof then
%><input type="hidden" name="NumRegs" value="0">
	<script language="javascript" type="text/javascript">
	    window.alert("<%=LitMsgNoDocumentos%>");

	    <%'**RGU 19/1/2007**
	    if tpv&"">"" then%>
			window.top.close();
	    <%end if%>
        parent.window.frames["botones"].document.location="estad_dia_tickets_bt.asp?mode=param";
	    document.location="estad_dia_tickets.asp?mode=param";
	  
	</script><%

	 response.end
   else
    %><input type="hidden" name="NumRegs" value="<%=rst.RecordCount%>"/><%
   end if 'fin eof

   'Calculos de páginas--------------------------
   lote=limpiaCadena(Request.QueryString("lote"))
   if lote="" then
	 lote=1
   end if
   sentido=limpiaCadena(Request.QueryString("sentido"))

   lotes=rst.recordcount/MAXPAGINA

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
    rst.AbsolutePage=lote

    NavPaginas lote,lotes,campo,criterio,texto,1
    CabeceraListado
    fila = 0
        
    if not rst.eof and cint(lote)=cint(lotes) then
        'recogemos los totales
        total_tickets=rst("total_tickets")
        total_venta_media=rst("total_venta_media")
        total_venta_total=rst("total_venta_total")
        total_articulos_distintos=rst("total_articulos_distintos")
        total_total_articulos=rst("total_total_articulos")
        total_beneficio_coste=rst("total_beneficio_coste")
        total_beneficio_venta=rst("total_beneficio_venta")
    end if

    'obtenemos el número de decimales con los que vamos a pintar los números
    n_decimales=d_lookup("ndecimales","divisas","moneda_base=1 and codigo like '" & session("ncliente") & "%'",Session("backendlistados"))
    if n_decimales = "" then
	    n_decimales = 0
  	end if

    while not rst.eof and fila<MAXPAGINA
        if ((fila+1) mod 2)=0 then
            color=color_blau
        else
            color=color_terra
        end if

        DrawFila color
        DrawCelda "TDSINBORDECELDA7","","",0,rst("fecha")
        if rst("max_tickets")=rst("num_tickets") then 'vamos a pintar un máximo (color azul)
	        DrawCelda "CELDARIGHT ","","",0,"<font color='blue'><b>" & formatnumber(null_z(rst("num_tickets")),0,-1,0,-1) & "</b></font>"
        elseif rst("min_tickets")=rst("num_tickets") then 'vamos a pintar un mínimo (rojo)
	        DrawCelda "CELDARIGHT","","",0,"<font color='red'><b>" & formatnumber(null_z(rst("num_tickets")),0,-1,0,-1) & "</b></font>"
        else 'no vamos a pintar nada especial
	        DrawCelda "CELDARIGHT","","",0,formatnumber(null_z(rst("num_tickets")),0,-1,0,-1)
        end if

        if rst("max_venta_media")=rst("venta_media") then
	        DrawCelda "CELDARIGHT","","",0,"<font color='blue'><b>" & formatnumber(null_z(rst("venta_media")),n_decimales,-1,0,-1) & "</b></font>"
        elseif rst("min_venta_media")=rst("venta_media") then
	        DrawCelda "CELDARIGHT","","",0,"<font color='red'><b>" & formatnumber(null_z(rst("venta_media")),n_decimales,-1,0,-1) & "</b></font>"
        else
	        DrawCelda "CELDARIGHT","","",0,formatnumber(null_z(rst("venta_media")),n_decimales,-1,0,-1)
        end if

        if rst("max_venta_total")=rst("venta_total") then
	        DrawCelda "CELDARIGHT","","",0,"<font color='blue'><b>" & formatnumber(null_z(rst("venta_total")),n_decimales,-1,0,-1) & "</b></font>"
        elseif rst("min_venta_total")=rst("venta_total") then
	        DrawCelda "CELDARIGHT","","",0, "<font color='red'><b>" & formatnumber(null_z(rst("venta_total")),n_decimales,-1,0,-1) & "</b></font>"
        else
	        DrawCelda "CELDARIGHT","","",0,formatnumber(null_z(rst("venta_total")),n_decimales,-1,0,-1)
        end if

        if rst("max_articulos_distintos")=rst("articulos_distintos") then
  	        DrawCelda "CELDARIGHT","","",0,"<font color='blue'><b>" & formatnumber(null_z(rst("articulos_distintos")),0,-1,0,-1)  & "</b></font>"
        elseif rst("min_articulos_distintos")=rst("articulos_distintos") then
	        DrawCelda "CELDARIGHT","","",0,"<font color='red'><b>" & formatnumber(null_z(rst("articulos_distintos")),0,-1,0,-1) & "</b></font>"
        else
	        DrawCelda "CELDARIGHT","","",0,formatnumber(null_z(rst("articulos_distintos")),0,-1,0,-1)
        end if

        if rst("max_total_articulos")=rst("articulos_totales") then
	        DrawCelda "CELDARIGHT","","",0,"<font color='blue'><b>" & formatnumber(null_z(rst("articulos_totales")),0,-1,0,-1) & "</b></font>"
        elseif rst("min_total_articulos")=rst("articulos_totales") then
	        DrawCelda "CELDARIGHT","","",0,"<font color='red'><b>" & formatnumber(null_z(rst("articulos_totales")),0,-1,0,-1) & "</b></font>"
        else
	        DrawCelda "CELDARIGHT","","",0,formatnumber(null_z(rst("articulos_totales")),0,-1,0,-1)
        end if

        if vben&""="0" then
        else
	        if rst("max_beneficio_coste")=rst("beneficio_coste") then
		        DrawCelda "CELDARIGHT","","",0,"<font color='blue'><b>" & formatnumber(null_z(rst("beneficio_coste")),n_decimales,-1,0,-1) & "</b></font>"
	        elseif rst("min_beneficio_coste")=rst("beneficio_coste") then
		        DrawCelda "CELDARIGHT","","",0,"<font color='red'><b>" & formatnumber(null_z(rst("beneficio_coste")),n_decimales,-1,0,-1) & "</b></font>"
	        else
		        DrawCelda "CELDARIGHT","","",0,formatnumber(null_z(rst("beneficio_coste")),n_decimales,-1,0,-1)
	        end if

	        if rst("max_beneficio_venta")=rst("beneficio_venta") then
		        DrawCelda "CELDARIGHT","","",0, "<font color='blue'><b>" &  formatnumber(null_z(rst("beneficio_venta")),n_decimales,-1,0,-1) & "</b></font>"
	        elseif rst("min_beneficio_venta")=rst("beneficio_venta") then
		        DrawCelda "CELDARIGHT","","",0,"<font color='red'><b>" & formatnumber(null_z(rst("beneficio_venta")),n_decimales,-1,0,-1) & "</b></font>"
	        else
		        DrawCelda "CELDARIGHT","","",0,formatnumber(null_z(rst("beneficio_venta")),n_decimales,-1,0,-1)
	        end if
        end if
        CloseFila

        fila = fila + 1
        rst.movenext
    wend

    if cint(lote)=cint(lotes) then 'si estamos en la última hoja pintamos los totales
        DrawFila color_fondo
        DrawCelda "TDSINBORDECELDA7","","",0,"<b>Totales</b>"
        DrawCelda "CELDARIGHT","","",0,"<b>" & formatnumber(null_z(total_tickets),0,-1,0,-1)  & "</b>"
        DrawCelda "CELDARIGHT","","",0,"<b>" & formatnumber(null_z(total_venta_media),n_decimales,-1,0,-1)  & "</b>"
        DrawCelda "CELDARIGHT","","",0,"<b>" & formatnumber(null_z(total_venta_total),n_decimales,-1,0,-1)  & "</b>"
        DrawCelda "CELDARIGHT","","",0,"<b>" & formatnumber(null_z(total_articulos_distintos),0,-1,0,-1)  & "</b>"
        DrawCelda "CELDARIGHT","","",0,"<b>" & formatnumber(null_z(total_total_articulos),0,-1,0,-1)  & "</b>"
        if vben&""="0" then
        else
	        DrawCelda "CELDARIGHT","","",0,"<b>" & formatnumber(null_z(total_beneficio_coste),n_decimales,-1,0,-1) & "</b>"
	        DrawCelda "CELDARIGHT","","",0,"<b>" & formatnumber(null_z(total_beneficio_venta),n_decimales,-1,0,-1) & "</b>"
        end if
	    CloseFila
    end if
    NavPaginas lote,lotes,campo,criterio,texto,2
    %>
   </table>
   <%
    rst.close
    set rst=nothing
    set command=nothing
    set conn=nothing
end if 'fin mode imp
%>
</form>
<%end if
set rstSelect=nothing
 'fin acceso página%>
</body>
</html>