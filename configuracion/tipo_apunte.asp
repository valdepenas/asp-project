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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
<TITLE><%=LitTituloTAp%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    function Editar(p_codigo, p_npagina, p_criterio, p_campo, p_texto) {
        document.location = "tipo_apunte.asp?mode=edit&p_codigo=" + p_codigo
            + "&npagina=" + p_npagina
            + "&campo=" + p_campo
            + "&texto=" + p_texto
            + "&criterio=" + p_criterio;
        parent.botones.document.location = "tipo_apunte_bt.asp?mode=edit";
    }
</script>
<body onload="self.status='';" bgcolor=<%=color_blau%>>
<%
'*************************************************************************************************************'
' CODIGO PRINCIPAL DE LA PAGINA  *****************************************************************************'
'*************************************************************************************************************'

if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<form name="tipo_apunte" method="post" action="tipo_apunte.asp">
    <%PintarCabecera "tipo_apunte.asp"
    set rst = server.CreateObject("ADODB.Recordset")

	'Leer parámetros de la página'
	mode=EncodeForHtml(request("mode"))

	p_i_descripcion=limpiaCadena(request.form ("i_descripcion"))
	p_i_cuenta=limpiaCadena(request.form("i_cuenta"))
	p_e_descripcion=limpiaCadena(request.form("e_descripcion"))
	p_h_codigo=limpiaCadena(Request.Form("h_codigo"))
	checkCadena p_h_codigo
	p_e_cuenta=limpiaCadena(request.form("e_cuenta"))
	p_c_codigo=limpiaCadena(request("codigo"))
	checkCadena p_c_codigo
	p_criterio=limpiaCadena(request("criterio"))
	p_campo=limpiaCadena(request("campo"))
	p_texto=limpiaCadena(request("texto"))
	p_npagina=limpiaCadena(request("npagina"))
	p_pagina=limpiaCadena(request("pagina"))
	p_p_codigo=limpiaCadena(request("p_codigo"))
	checkCadena p_p_codigo

   'insertamos si nos llegan los valores'
      if p_i_descripcion>"" then
	      p_codigo=""
      	p_descripcion=p_i_descripcion
		p_cuenta=p_i_cuenta
      	rst.Open "select * from tipo_apuntes where codigo like '" & session("ncliente") & "%' order by convert(int,substring(codigo,6,9)) desc",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	      if rst.EOF then
			rst.AddNew
			rst("codigo")=session("ncliente") + "01"
	      	rst("descripcion")="POR DOCUMENTO"
			rst("p_cuenta")=NULL
      		rst.Update
			rst.close
      	else
			p_codigo=cint(null_z(trimCodEmpresa(rst("codigo"))))+1
	  		p_codigo=session("ncliente") & completar(cstr(p_codigo),3,"0")
			rst.AddNew
			rst("codigo")=p_codigo
			rst("descripcion")=p_descripcion
			rst("cuenta")=nulear(p_cuenta)
			rst.Update
			rst.close
		end if
	end if

  'actualizamos valores
  if p_e_descripcion>"" then
  	  p_codigoAnt = p_h_codigo
      p_codigo=p_h_codigo
      p_descripcion=p_e_descripcion
	  p_cuenta=p_e_cuenta
	  if p_codigo<>p_codigoAnt then
	  	rst.Open "select * from tipo_apuntes with(nolock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	  	if not rst.EOF then
			rst.close
			'ya existe el nuevo codigo que se quiere asignar a este medio de pago %>
			<SCRIPT language="JavaScript">
					window.alert("<%=LitMsgCodigoExiste%>")
					document.location="tipo_apunte.asp"
			</script><%
		else
			rst.close
			on error resume next
     		rst.Open "delete from tipo_apuntes with(rowlock) where codigo='" + p_codigoAnt + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	 		if err.number = -2147217900 then
				'existen documentos con el codigo anterior del medio de pago%>
	 			<SCRIPT language="JavaScript">
					window.alert("<%=LitMsgModifTipoApunte%>")
					document.location="tipo_apunte.asp"
				</script><%
			else
			 	rst.Open "select * from tipo_apuntes where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
				rst.AddNew
         		rst("codigo")  = p_codigo
         		rst("descripcion")   = p_descripcion
				rst("cuenta")   = nulear(p_cuenta)
         		rst.Update
				rst.close
			end if
		end if
	  else ' los codigos son iguales
	  	rst.Open "select * from tipo_apuntes with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		rst("codigo")  = p_codigo
        rst("descripcion")   = p_descripcion
		rst("cuenta")   = nulear(p_cuenta)
        rst.Update
		rst.close
	  end if
  end if

  'eliminamos valores
  if mode="delete" and p_c_codigo>"" then
  	 if p_c_codigo<>"01" then
		 on error resume next
    	 p_codigo=p_c_codigo
	     rst.Open "delete from tipo_apuntes with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		 if err.number = -2147217900 then %>
		 	<SCRIPT language="JavaScript">
						window.alert("<%=LitMsgBorrarTipoApunte%>")
						document.location="tipo_apunte.asp"
			</script><%
		 end if
	 else%>
	 	<SCRIPT language="JavaScript">
			window.alert("<%=LitMsgBorrarMedioMetalico%>")
			document.location="tipo_apunte.asp"
		</script><%
	 end if
  end if

   'if mode="search" then

   if p_texto>"" then
      if p_campo="codigo" then
        c_where=" where " + p_campo + " like'" + session("ncliente")
      else
        c_where=" where " + p_campo + " like'"
      end if
   else
      c_where=""
   end if

   if c_where>"" then
      select case p_criterio
         case "contiene"
            c_where=c_where + "%" + p_texto + "%'"
         case "termina"
            c_where=c_where + "%" + p_texto + "'"
         case "empieza"
            c_where=c_where + p_texto + "%'"
         case "igual"
         	if p_campo="codigo" then
         	  c_where=" where " + p_campo + "='" + session("ncliente") + p_texto + "'"
         	else
              c_where=" where " + p_campo + "='" + p_texto + "'"
            end if
         end select
      end if

   Alarma "tipo_apunte.asp" %>
   <hr>
  <%
    c_select="select * from tipo_apuntes with(nolock)"

        if c_where>"" then
           c_select=c_select+c_where+" and codigo like '"+session("ncliente")+"%'"
        else
           c_select=c_select+" where codigo like '"+session("ncliente")+"%'"
        end if


        if p_npagina="" then
           p_npagina=1
        end if

        select case p_pagina
           case "siguiente"
              p_npagina=p_npagina+1
           case "anterior"
              p_npagina=p_npagina-1
        end select
%>
  <input type="hidden" name="h_npagina" value="<%=EncodeForHtml(null_s(cstr(p_npagina)))%>">
<%
		rst.Open c_select,session("dsn_cliente"),adUseClient, adLockReadOnly
        if not rst.EOF then
           rst.PageSize=NumReg
           rst.AbsolutePage=p_npagina
        end if

  if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
		 <a class=CABECERA href="tipo_apunte.asp?pagina=anterior&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
		 <IMG SRC="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></a>
  	<%end if

    texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	<font class=CELDA> <%=EncodeForHtml(texto)%> </font> <%

     if clng(p_npagina)<rst.PageCount then %>
		<a class=CABECERA href="tipo_apunte.asp?pagina=siguiente&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
		<IMG SRC="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></a>
  	<%end if

	%><font class=CELDA>&nbsp;&nbsp; Ir a Pag. 
        <input class="CELDA" type="text" name="SaltoPagina1" size="2">&nbsp;&nbsp;
        <a class="CELDAREF inlineBlock" href="javascript:IrAPagina(1,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina');">Ir</a>
	  </font><%
  end if%>
  <table class="width100 xs-table-responsive bCollapse" BORDER="0" CELLSPACING="1" CELLPADDING="1">
      <%Drawfila color_fondo
            DrawCeldaDet "'ENCABEZADOL underOrange width5'", "", "", 1, "<b>" & LitCodigo & "</b>"
            DrawCeldaDet "'ENCABEZADOL underOrange width20'", "", "", 1, "<b>" & LitDescripcion & "</b>"
            DrawCeldaDet "'ENCABEZADOL underOrange width20'", "", "", 1, "<b>" & LitCuenta & "</b>"
        par=false
        i=1

        while not rst.EOF and i<=NumReg
           if par then
              Drawfila color_terra
              par=false
           else
              Drawfila color_blau
              par=true
           end if
           if mode="edit" and p_p_codigo=rst("codigo") then
		   	    %><input type="Hidden" name="h_codigo" value="<%=EncodeForHtml(null_s(rst("codigo")))%>"><%
                DrawCeldaDet "'CELDAL7 width5'", "", "", 0, EncodeForHtml(trimCodEmpresa(rst("codigo")))
                %><td class="CELDAL7 width20"><%DrawInput "width60","","e_descripcion",EncodeForHtml(null_s(rst("descripcion"))),"maxlength='50'"%></td><%
                %><td class="CELDAL7 width20"><%DrawInput "width20","","e_cuenta",EncodeForHtml(null_s(rst("cuenta"))),"maxlength='50'"%></td><%
           else
			  h_ref="javascript:Editar('" & EncodeForHtml(rst("codigo")) & "'," & _
			                           EncodeForHtml(p_npagina) & ",'" & _
									   EncodeForHtml(p_criterio) & "','" & _
									   EncodeForHtml(p_campo) & "','" & _
									   EncodeForHtml(p_texto) & "');"
			  if rst("codigo")<>session("ncliente")+"01" then
              		'DrawCeldaHref "CELDAREF","left",false,trimCodEmpresa(rst("codigo")),h_ref
                    %><td class="CELDAL7 width5"><%
                        DrawHref "CELDAREF","",EncodeForHtml(trimCodEmpresa(rst("codigo"))),h_ref
                    %></td><%
			  else
			  		DrawCeldaDet "'CELDAL7 width5'", "", "", 0, "<b>" & EncodeForHtml(trimCodEmpresa(rst("codigo"))) & "</b>"
			  end if
              DrawCeldaDet "'CELDAL7 width20'", "", "", 0, EncodeForHtml(null_s(rst("descripcion")))
			  DrawCeldaDet "'CELDAL7 width20'", "", "", 0, EncodeForHtml(null_s(rst("cuenta")))
           end if

           i = i + 1
           rst.MoveNext
        wend
        'rst.Close %>
   </table>
  <%
  if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
	  	 <a class=CABECERA href="tipo_apunte.asp?pagina=anterior&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
	     <IMG SRC="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> ALT="<%=LitAnterior%>"></a>
  	 <%end if
     texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	 <font class=CELDA> <%=EncodeForHtml(texto)%> </font> <%
     if clng(p_npagina)<rst.PageCount then %>
	 	<a class=CABECERA href="tipo_apunte.asp?pagina=siguiente&npagina=<%=enc.EncodeForJavascript(null_s(cstr(p_npagina)))%>&campo=<%=enc.EncodeForJavascript(p_campo)%>&criterio=<%=enc.EncodeForJavascript(p_criterio)%>&texto=<%=enc.EncodeForJavascript(p_texto)%>">
  		<IMG SRC="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> ALT="<%=LitSiguiente%>"></a>
  	 <%end if

	 %><font class="CELDA">&nbsp;&nbsp; Ir a Pag. <input class="CELDA" type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;
         <a class="CELDAREF inlineBlock" href="javascript:IrAPagina(2,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina')">Ir</a>
    </font><%

	 rst.Close

  end if%> <br>

   <%if mode<>"edit" then %>
   <hr>
   <table class="width100 md-table-responsive" BORDER="0" CELLSPACING="1" CELLPADDING="1">
       <%DrawceldaDet "'ENCABEZADOL underOrange width50'", "", "", 1,"<b>" & LitNBregistro & "</b>"%>
   </table>
	<table class="width100 underOrange xs-table-responsive bCollapse" BORDER="0" CELLSPACING="1" CELLPADDING="1">
        <tr class="underOrange">
            <td class="CELDA underOrange width5"></td>
            <%
                DrawCeldaDet "'ENCABEZADOL underOrange width20'","", "", 1,"<b>" & LitDescripcion & "</b>"
                DrawCeldaDet "'ENCABEZADOL underOrange width20'","", "", 1,"<b>" & LitCuenta & "</b>"
            %>
        </tr>
            <tr>
                <td class="CELDA underOrange width5"></td>
                <td class="CELDA underOrange width20">
                    <% DrawInput "width60", "", "i_descripcion", "", "maxlength='50'" %>
                </td>
                <td class="CELDA underOrange width20" style="text-align:left">
                    <%  DrawInput "width20", "", "i_cuenta", "", "maxlength='20'"  %>
                </td>
            </tr>
      </table>
      <hr>
   <%end if%>
   </form>
<%else
	MsgError LitSinSesion%>
	<br><a href="../" target="_top">Iniciar sesión</a>
<%end if%>
</BODY>
</HTML>