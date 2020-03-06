<%@ Language=VBScript %>

<% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<TITLE><%=LitTituloFP%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<script Language="JavaScript">
function Editar(p_codigo, p_npagina, p_campo, p_criterio, p_texto,viene) {
   	
	if(viene=="asistente"){
	    document.formas_pago.action="formas_pago.asp?mode=edit&p_codigo=" + p_codigo
											 +"&npagina="+ p_npagina
                                             +"&campo="  + p_campo
                                             +"&texto="  + p_texto
                                             +"&criterio=" + p_criterio
                                             +"&viene=" +viene;
	    document.formas_pago.submit();
	    parent.botones.document.location="formas_pago_bt.asp?mode=edit&viene=asistente";	     
	}
	else{
	    document.formas_pago.action="formas_pago.asp?mode=edit&p_codigo=" + p_codigo
											 +"&npagina="+ p_npagina
                                             +"&campo="  + p_campo
                                             +"&texto="  + p_texto
                                             +"&criterio=" + p_criterio;                                          
	    document.formas_pago.submit();
	    parent.botones.document.location="formas_pago_bt.asp?mode=edit";   
	}
}

</script>
<body bgcolor=<%=color_blau%>>
<%
'************************************************************************************************************'
' CODIGO PRINCIPAL DE LA PAGINA  ****************************************************************************'
'************************************************************************************************************'
if accesoPagina(session.sessionid,session("usuario"))=1 then%>
    <form name="formas_pago" method="post" action="formas_pago.asp"><%
	PintarCabecera "formas_pago.asp"
    set rst = server.CreateObject("ADODB.Recordset")

	'Leer parámetros de la página'
	mode=request("mode")
	
	'AMP 28/07/2010 : Añadimos parametro viene para adaptar las formas de pago  al asistente de puesta en marcha.
	viene=Request.QueryString("viene")
  
	p_i_codigo=limpiaCadena(Request.Form("i_codigo"))
	p_i_descripcion=limpiaCadena(request.form ("i_descripcion"))
	p_i_diasff=limpiaCadena(request.form("i_diasff"))
	p_i_dias=limpiaCadena(request.form("i_dias"))
	p_i_ncuotas=limpiaCadena(request.form("i_ncuotas"))
	p_e_codigo=limpiaCadena(request.form("e_codigo"))
	p_e_descripcion=limpiaCadena(request.form("e_descripcion"))
	p_h_codigo=limpiaCadena(Request.Form("h_codigo"))
	p_e_diasff=limpiaCadena(request.form("e_diasff"))
	p_e_dias=limpiaCadena(request.form("e_dias"))
	p_e_ncuotas=limpiaCadena(request.form("e_ncuotas"))
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
      if p_i_codigo>"" and p_i_descripcion>"" then
		p_codigo  = left(p_i_codigo,5)
		p_descripcion   = p_i_descripcion
		p_diasff = null_z((p_i_diasff))
		p_dias    = null_z(p_i_dias)
		p_ncuotas = null_z(p_i_ncuotas)
      rst.Open "select * from formas_pago where codigo='" & session("ncliente") & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
      if rst.EOF then
         rst.AddNew
         rst("codigo")  = session("ncliente") & p_codigo
         rst("descripcion")   = p_descripcion
		 'cag
		 rst("diasff") = p_diasff
		 rst("dias")    = p_dias
		 rst("ncuotas") = p_ncuotas
         rst.Update
      else %>
         <script>
            window.alert("<%=LitMsgCodigoExiste%>");
            history.back();
         </script>
  <%  end if
      rst.Close
   end if

  'actualizamos valores
  if p_e_codigo>"" or p_e_descripcion>"" then
	  p_codigoAnt = p_h_codigo
      p_codigo  = p_e_codigo
      p_descripcion   = p_e_descripcion
	  p_diasff = null_z(p_e_diasff)
	  p_dias    = null_z(p_e_dias)
	  p_ncuotas = null_z(p_e_ncuotas)
	  if p_codigo<>p_codigoAnt then
	  	rst.Open "select * from formas_pago with(nolock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	  	if not rst.EOF then
			rst.close
			'ya existe el nuevo codigo que se quiere asignar a esta forma de pago %>
			<SCRIPT language="JavaScript">
					window.alert("<%=LitMsgCodigoExiste%>")
					document.location="formas_pago.asp"
			</script><%
		else
			rst.close
			on error resume next
     		rst.Open "delete from formas_pago with(rowlock) where codigo='" + p_codigoAnt + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	 		if err.number = -2147217900 then
				'existen documentos con el codigo anterior de la forma de pago%>
	 			<SCRIPT language="JavaScript">
					window.alert("<%=LitMsgModifFormaPago%>")
					document.location="formas_pago.asp"
				</script><%
			else
			 	rst.Open "select * from formas_pago where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
				rst.AddNew
         		rst("codigo")  = p_codigo
         		rst("descripcion")   = p_descripcion
				rst("diasff") = p_diasff
		 		rst("dias")    = p_dias
		 		rst("ncuotas") = p_ncuotas
         		rst.Update
				rst.close
			end if
		end if
	  else ' los codigos son iguales
	  	rst.Open "select * from formas_pago with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		if not rst.eof then
			rst("codigo")  = p_codigo
			rst("descripcion")   = p_descripcion
			rst("diasff") = p_diasff
			rst("dias")    = p_dias
			rst("ncuotas") = p_ncuotas
			rst.Update
			rst.close
		else
			rst.close
			mode=""
			%><script language="javascript">
				window.alert("<%=LitMsgDatosNoExiste%>");
			</script><%
		end if
	  end if
  end if

    'eliminamos valores
    if mode="delete" and p_c_codigo>"" then
  	    on error resume next
        p_codigo=p_c_codigo
        rst.Open "delete from formas_pago with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	    if err.number = -2147217900 then%>
	 	    <script language="JavaScript">
				window.alert("<%=LitMsgBorrarFormaPago%>")
				document.location="formas_pago.asp"
		    </script>
		<%end if
    end if

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

    Alarma "formas_pago.asp" %>
	<hr>
    <%c_select="select * from formas_pago with(nolock)"

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
    end select%>
    <input type="hidden" name="h_npagina" value="<%=enc.EncodeForHtmlAttribute(cstr(p_npagina) & "")%>">
    <%rst.Open c_select,session("dsn_cliente"),adUseClient, adLockReadOnly

    if not rst.EOF then
        rst.PageSize=NumReg
        rst.AbsolutePage=p_npagina
    end if

    if mode<>"edit" and rst.RecordCount>NumReg then
        if clng(p_npagina) >1 then %>
		    <a class="CABECERA" href="formas_pago.asp?pagina=anterior&npagina=<%=enc.EncodeForHtmlAttribute(cstr(p_npagina) & "")%>&campo=<%=p_campo%>&criterio=<%=p_criterio%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto & "")%>">
		    <IMG SRC="<%=themeIlion %><%=ImgAnterior%>" align='top' ALT="<%=LitAnterior%>"></a>
  	    <%end if
        texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	    <font class="CELDA"> <%=texto%> </font>
  	    <%if clng(p_npagina)<rst.PageCount then %>
		    <a class="CABECERA" href="formas_pago.asp?pagina=siguiente&npagina=<%=enc.EncodeForHtmlAttribute(cstr(p_npagina) & "")%>&campo=<%=p_campo%>&criterio=<%=p_criterio%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto & "")%>">
		    <IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" align='top' ALT="<%=LitSiguiente%>"></a>
  	    <%end if%>
            <font class="CELDA">&nbsp;&nbsp; Ir a Pag. <input class="CELDA" type="text" name="SaltoPagina1" size="2">&nbsp;&nbsp;<a class="CELDAREF" href="javascript:IrAPagina(1,'<%=p_campo%>','<%=p_criterio%>','<%=enc.EncodeForHtmlAttribute(p_texto & "")%>',<%=rst.PageCount%>,'npagina');">Ir</a></font>
        <%end if%>

        <table class="width100 md-table-responsive bCollapse" BORDER="0" CELLSPACING="1" CELLPADDING="1">
            <%Drawfila color_fondo
                DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCodigo & "</b>"
                DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & LitDescripcion & "</b>"
	            DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & Litdiasff & "</b>"
                DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & LitDias & "</b>"
                DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & LitNcuotas & "</b>"
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
               if mode="edit" and p_p_codigo=rst("codigo") then%>
		   	        <input type="hidden" name="h_codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("codigo")))%>">
		   	        <%DrawceldaDet "'CELDAL7 width5'", "left","", false, trimCodEmpresa(rst("codigo"))%>
		            <input type="hidden" name="e_codigo" value="<%=enc.EncodeForHtmlAttribute(null_s(rst("codigo")))%>">
		            <td class="CELDAL7 width20">
                        <%DrawInput "width60","","e_descripcion",enc.EncodeForHtmlAttribute(null_s(rst("descripcion"))),"maxlength='50'"
                    %></td><td class="CELDAL7 width20"><%
		            ''MPC 04/11/2010 Se comenta dicho por JCI
                    DrawInput "width20","","e_diasff",enc.EncodeForHtmlAttribute(null_s(rst("diasff"))),"size=3"
                     %></td><td class="CELDAL7 width20"><%
                    ''FIN MPC 04/11/2010
                    'DrawInputCelda "CELDA","","","3",0,"","e_dias",rst("dias")
                    DrawInput "width20","","e_dias",enc.EncodeForHtmlAttribute(null_s(rst("dias"))),"size=3"
                    %></td><td class="CELDAL7 width20"><%
                    'DrawInputCelda "CELDA","","","3",0,"","e_ncuotas",rst("ncuotas")
                    DrawInput "width20","","e_ncuotas",rst("ncuotas"),"size=3"
                    %></td><%
               else
                    h_ref="javascript:Editar('" + rst("codigo") + "'," & _
			                               enc.EncodeForJavascript(p_npagina) & ",'" & _
									       enc.EncodeForJavascript(p_campo) & "','" & _
									       enc.EncodeForJavascript(p_criterio) & "','" & _
									       enc.EncodeForJavascript(p_texto) & "','" & enc.EncodeForJavascript(viene) & "');"
                    'DrawCeldaHref "CELDAREF","left",false,trimCodEmpresa(rst("codigo")),h_ref
                     %><td class="CELDAL7 width5">                
                        <%DrawHref "CELDAREF","",trimCodEmpresa(rst("codigo")),h_ref%></td><%  
                    DrawCeldaDet "'CELDAL7 width20'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))
			        DrawCeldaDet "'CELDAL7 width20'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("diasff")))
			        DrawCeldaDet "'CELDAL7 width20'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("dias")))
			        DrawCeldaDet "'CELDAL7 width20'","", "left", false, enc.EncodeForHtmlAttribute(null_s(rst("ncuotas")))
               end if
               i = i + 1
               rst.MoveNext
            wend%>
        </table>

        <%if mode<>"edit" and rst.RecordCount>NumReg then
            if clng(p_npagina) >1 then %>
	  	        <a class="CABECERA" href="formas_pago.asp?pagina=anterior&npagina=<%=enc.EncodeForHtmlAttribute(cstr(p_npagina) & "")%>&campo=<%=p_campo%>&criterio=<%=p_criterio%>&texto=<%=enc.EncodeForHtmlAttribute(cstr(p_texto & ""))%>">
			    <IMG SRC="<%=themeIlion %><%=ImgAnterior%>" align='top' ALT="<%=LitAnterior%>"></a>
  	        <%end if
            texto=LitPagina + " " + enc.EncodeForHtmlAttribute(cstr(p_npagina) & "")+ " "+ LitDe + " " + enc.EncodeForHtmlAttribute(cstr(rst.PageCount) & "")%>
  	        <font class="CELDA"> <%=texto%> </font>
  	        <%if clng(p_npagina)<rst.PageCount then %>
	 	        <a class="CABECERA" href="formas_pago.asp?pagina=siguiente&npagina=<%=enc.EncodeForHtmlAttribute(cstr(p_npagina))%>&campo=<%=enc.EncodeForHtmlAttribute(p_campo)%>&criterio=<%=enc.EncodeForHtmlAttribute(p_criterio)%>&texto=<%=enc.EncodeForHtmlAttribute(p_texto)%>">
		        <IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" align='top' ALT="<%=LitSiguiente%>"></a>
  	        <%end if%>

  	        <font class="CELDA">&nbsp;&nbsp; Ir a Pag. <input class="CELDA" type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;<a class="CELDAREF" href="javascript:IrAPagina(2,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=enc.EncodeForJavascript(rst.PageCount)%>,'npagina');">Ir</a></font>
  	        <%rst.Close
        end if%>

        <br>

        <%if mode<>"edit" then %>
        <hr>
            <table class="width100 md-table-responsive" width=100% BORDER="0" CELLSPACING="1" CELLPADDING="1">
                <%DrawceldaDet "'ENCABEZADOL underOrange width50'", "", "left", true,"<b>" & LitNBregistro & "</b>"%>
            </table>
	        <table class="width100 underOrange md-table-responsive bCollapse">
                <tr class="underOrange">
                <%DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitCodigo & "</b>"
                  DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & LitDescripcion & "</b>"
	              DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & Litdiasff & "</b>"
                  DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & LitDias & "</b>"
                  DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<b>" & LitNcuotas & "</b>"%>
                </tr>
                <tr><%
                ''FIN MPC 04/11/2010%>
                 <td class="CELDA underOrange width5"><%
                     DrawInput "width100","","i_codigo","","size=2"%></td>
                 <td class="CELDA underOrange width20"><%
                     DrawInput "width60","","i_descripcion","","size=50"%></td>
                 <td class="CELDA underOrange width20"><%
                     DrawInput "width20","","i_diasff","","size=3"%></td>
                 <td class="CELDA underOrange width20"><%
                     DrawInput "width20","","i_dias","","size=3"%></td>
                 <td class="CELDA underOrange width20" style="text-align:left;"><%
                     DrawInput "width20","","i_ncuotas","","size=3"%></td>
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