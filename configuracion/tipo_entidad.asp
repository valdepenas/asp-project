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
<!--#include file="../mensajes.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="agrupaciones.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<TITLE><%=LitTituloTEnt%></TITLE>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
</HEAD>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
function Editar(p_codigo, p_npagina, p_campo, p_criterio, p_texto)
{
	document.location="tipo_entidad.asp?mode=edit&p_codigo=" + p_codigo + "&npagina=" + p_npagina + "&campo=" + p_campo + "&texto=" + p_texto + "&criterio=" + p_criterio;
	parent.botones.document.location="tipo_entidad_bt.asp?mode=edit";
}

function VerGrupos()
{
    if (document.getElementById("tipo").value == "<%=LitGrupoComercial%>")
    {
        document.getElementById("cabgrupo").style.display = "";
        document.getElementById("grupo").style.display = "";
        document.getElementById("cabdias").style.display = "";
        document.getElementById("dias").style.display = "";
    }
    else
    {
        document.getElementById("cabgrupo").style.display = "none";
        document.getElementById("grupo").style.display = "none";
        document.getElementById("cabdias").style.display = "none";
        document.getElementById("dias").style.display = "none";
    }
}
</script>
<body bgcolor=<%=color_blau%>>
<%
'***********************************************************************************************************
' CODIGO PRINCIPAL DE LA PAGINA  ***************************************************************************
'***********************************************************************************************************
function verifica_codigo(f_codigo, fnew, ftipo_old, ftipo_new, edit)
    verifica_codigo = true
    if edit = 1 then
        msg1 = LitNoModificar
    else
        msg1 = LitNoBorrar
    end if
    'Verificamos que el codigo a editar o borrar no están ningún cliente, proveedor o contacto comercial
    if f_codigo>"" and (f_codigo<>fnew or ftipo_old<>ftipo_new) then
	    if d_lookup("ncliente","clientes","ncliente like '" & session("ncliente") & "%' and tipo_cliente='" & f_codigo & "'",session("dsn_cliente"))<>"" then
            verifica_codigo = false
            %><script>
                alert("<%=msg1%> <%=LitCliente%>");
            </script><%
	    else
	        if d_lookup("nproveedor","proveedores","nproveedor like '" & session("ncliente") & "%' and tipo_proveedor='" & f_codigo & "'",session("dsn_cliente"))<>"" then
		         verifica_codigo =  false
		         %><script>
		            alert("<%=msg1%> <%=LitProveedor%>");
		         </script><%
		    else
		         if d_lookup("codigo","contactoscomercial","codigo like '" & session("ncliente") & "%' and tipo_contacto='" & f_codigo & "'",session("dsn_cliente"))<>"" then
		                verifica_codigo =  false
		                %><script>
		                    alert("<%=msg1%> <%=LitContacto%>");
		                </script><%
		         end if
		    end if
	    end if
    end if
end function

if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<form name="tipo_entidad" method="post" action="tipo_entidad.asp">
    <%PintarCabecera "tipo_entidad.asp"
    set rst = server.CreateObject("ADODB.Recordset")
    set rstAux = server.CreateObject("ADODB.Recordset")

	'Leer parámetros de la página'
	mode=EncodeForHtml(request("mode"))

	p_i_codigo=limpiaCadena(Request.Form("i_codigo"))
	p_i_descripcion=limpiaCadena(request.form ("i_descripcion"))
	p_i_tipo=limpiaCadena(request.form("i_tipo"))
	p_i_dias=limpiaCadena(request.form("i_diasperm"))
	p_i_grupo=limpiaCadena(request.form("i_cambioagrupo"))
	p_e_codigo=limpiaCadena(request.form("e_codigo"))
	p_e_descripcion=limpiaCadena(request.form("e_descripcion"))
	p_e_dias=limpiaCadena(request.form("e_diasperm"))
	p_e_grupo=limpiaCadena(request.form("e_cambioagrupo"))
	p_hcodigo=limpiaCadena(Request.Form("hcodigo"))
	checkCadena p_hcodigo
	p_e_tipo=limpiaCadena(request.form("e_tipo"))
	p_htipo=limpiaCadena(request.form("htipo"))
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
        p_codigo=p_i_codigo
        p_descripcion=p_i_descripcion
        rst.Open "select * from tipos_entidades where codigo='" + session("ncliente") + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        if rst.EOF then
            rst.AddNew
            rst("codigo")=session("ncliente") & p_codigo
            rst("descripcion")=p_descripcion
            rst("tipo")=p_i_tipo
            rst("diaspermanencia") = null_z(p_i_dias)
            rst("cambioagrupo") = nulear(p_i_grupo)
            rst.Update
        else%>
            <script>
                alert("<%=LitMsgCodigoExiste%>");
            </script>
        <%end if
        rst.Close
    end if

    'actualizamos valores
    if p_e_codigo>"" or p_e_descripcion>"" then
        codigo_or = p_hcodigo
        p_codigo=p_e_codigo
        p_tipo_new = p_e_tipo
        p_tipo_old = p_htipo
        p_dias = p_e_dias
        p_grupo = p_e_grupo
        if verifica_codigo(codigo_or, p_codigo, p_tipo_new, p_tipo_old, 1) then
            p_descripcion=p_e_descripcion
            rst.Open "select * from tipos_entidades with(rowlock) where codigo='" + codigo_or + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
            if not rst.EOF then
                rst("codigo")=p_codigo
                rst("descripcion")=p_descripcion
                rst("tipo") = p_tipo_new
                rst("diaspermanencia") = null_z(p_dias)
                rst("cambioagrupo") = nulear(p_grupo)
                rst.Update
            else%>
                <script>
                    alert("<%=LitMsgCodigoExiste%>");
                </script>
            <%end if
            rst.Close
        end if
    end if

    'eliminamos valores
    if mode="delete" and p_c_codigo>"" then
        p_codigo=p_c_codigo
        'miramos si se puede borrar esta entidad o esta en algun articulo
        rst.open "select * from articulos with(nolock) where tipo_articulo='" & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        if not rst.eof then
            rst.close%>
            <script language="javascript">
                alert("<%=LitNoBorrTipEntporArt%>");
            </script>
        <%else
            rst.close
            if verifica_codigo(p_codigo, "", "old", "new", 0) then
                rst.Open "select * from tipos_entidades with(rowlock) where codigo='" + p_codigo + "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
                if not rst.eof then
                    rst.Delete
                else%>
                    <script>
                        alert("<%=LitMsgCodigoNoExiste%>");
                    </script>
                <%end if
                rst.Close
            end if
        end if
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
    Alarma "tipo_entidad.asp" %>
    <hr>
    <%c_select="select * from tipos_entidades"

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
    <input type="hidden" name="h_npagina" value="<%=EncodeForHtml(null_s(cstr(p_npagina)))%>">
    <%rst.Open c_select,session("dsn_cliente"),adUseClient, adLockReadOnly

    if not rst.EOF then
        rst.PageSize=NumReg
        rst.AbsolutePage=p_npagina
    end if

    if mode<>"edit" and rst.RecordCount>NumReg then
        if clng(p_npagina) >1 then%>
            <a class="CABECERA" href="tipo_entidad.asp?pagina=anterior&npagina=<%=EncodeForHtml(null_s(cstr(p_npagina)))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%=EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
            <IMG SRC="<%=themeIlion %><%=ImgAnterior%>" align="top" ALT="<%=LitAnterior%>"></a>
        <%end if

        texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
        <font class=CELDA> <%=texto%> </font>
        
        <%if clng(p_npagina)<rst.PageCount then%>
            <a class="CABECERA" href="tipo_entidad.asp?pagina=siguiente&npagina=<%=EncodeForHtml(null_s(cstr(p_npagina)))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%=EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
            <IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" align="top" ALT="<%=LitSiguiente%>"></a>
        <%end if%>
        <font class="CELDA">&nbsp;&nbsp; Ir a Pag. <input class="CELDA" type="text" name="SaltoPagina1" size="2">&nbsp;&nbsp;<a class="CELDAREF" href="javascript:IrAPagina(1,'<%=EncodeForHtml(p_campo)%>','<%=EncodeForHtml(p_criterio)%>','<%=EncodeForHtml(p_texto)%>',<%=rst.PageCount%>,'npagina');">Ir</a></font>
    <%end if%>

    <table class="width100 md-table-responsive bCollapse" border="0" cellspacing="1" cellpadding="1">
        <tr><%
        DrawceldaDet "'ENCABEZADOL width10'", "", "left", true,"<b>" & LitCodigo & "</b>"
        DrawceldaDet "'ENCABEZADOL width20'", "", "left", true,"<b>" & LitDescripcion & "</b>"
		DrawceldaDet "'ENCABEZADOL width20'", "", "left", true,"<b>" & LitTipoDe & "</b>"
		paso=false
		rstAux.cursorlocation=3
		rstAux.open "select top 1 codigo from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='" & LitGrupoComercial & "'",session("dsn_cliente")
	    if not rstAux.eof then
		DrawceldaDet "'ENCABEZADOL width20'", "", "left", true,"<b>" & LitDiasPerm & "</b>"
		DrawceldaDet "'ENCABEZADOL width20'", "", "left", true,"<b>" & LitCambioAGrupo & "</b>"
		paso=true
		end if
		rstAux.close
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
			    <input type="hidden" name="hcodigo" value="<%=EncodeForHtml(null_s(rst("codigo")))%>">
			    <input type="hidden" name="htipo" value="<%=EncodeForHtml(null_s(rst("tipo")))%>">
			    <input type="hidden" name="e_codigo" value="<%=EncodeForHtml(null_s(rst("codigo")))%>">
			    <%DrawCeldaDet "'CELDAL7 width10'", "left","", false, EncodeForHtml(trimCodEmpresa(rst("codigo")))%>
			    <td class="CELDAL7 width20"><input class="width70" type='text' name='e_descripcion' value='<%=EncodeForHtml(null_s(rst("descripcion")))%>' maxlength=50 size=50></td>
		        <td class="CELDAL7 width20">
				    <select class="width60" name='e_tipo'>
			  		    <option <%=iif(rst("tipo")="CLIENTE","selected","")%> value="CLIENTE"><%=LitTCliente%></option>
			  		    <option <%=iif(rst("tipo")="PROVEEDOR","selected","")%> value="PROVEEDOR"><%=LitTProveedor%></option>
					    <option <%=iif(rst("tipo")="CONTACTO COMERCIAL","selected","")%> value="CONTACTO COMERCIAL"><%=LitTContacto%></option>
					    <option <%=iif(rst("tipo")="PERSONAL","selected","")%> value="PERSONAL"><%=LitTPersonal%></option>
					    <option <%=iif(rst("tipo")="CENTRO","selected","")%> value="CENTRO"><%=LitTCentro%></option>
					    <option <%=iif(rst("tipo")="ARTICULO","selected","")%> value="ARTICULO"><%=LitTArticulo%></option>
					    <option <%=iif(rst("tipo")="GRUPO CONT.COMERCIAL","selected","")%> value="GRUPO CONT.COMERCIAL"><%=LitGrupoComercial%></option>
                        <option <%=iif(rst("tipo")="COMPANYGROUP","selected","")%> value="COMPANYGROUP"><%=LitcompanyGroup%></option>
                        <option <%=iif(rst("tipo")="REASON_REGULAR","selected","")%> value="REASON_REGULAR"><%=UCase(LitReasonRegularization)%></option>
                        <%if ModuloContratado(session("ncliente"),ModQSR) <> 0 then%>
                        <option <%=iif(rst("tipo")="ELABORATION","selected","")%> value="ELABORATION"><%=UCase(LitElaboration)%></option>
                        <%end if%>
				    </select>
			    </td>
			    <%if paso then
			        if rst("tipo") = LitGrupoComercial then%>
			            <td class="CELDAL7 width20"><input class="width70" type='text' name='e_diasperm' value='<%=EncodeForHtml(null_s(rst("diaspermanencia")))%>' maxlength="3" size="3"></td>
			            <%rstAux.cursorlocation=3
				        rstAux.open "select codigo, descripcion from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='" & LitGrupoComercial & "' order by codigo",session("dsn_cliente")
				        'DrawSelectCelda "CELDA",iif(mode<>"browse","200",""),"",0,"","e_cambioagrupo",rstAux,rst("cambioagrupo"),"codigo","descripcion","",""
                        %><td class="CELDAL7 width20">
                            <%DrawSelect "width60","","e_cambioagrupo",rstAux,EncodeForHtml(null_s(rst("cambioagrupo"))),"codigo","descripcion","",""%>
                          </td><%
				        rstAux.close
			        else
			            DrawCeldaDet "'CELDAL7 width20'","", "left", false, EncodeForHtml(null_s(rst("diaspermanencia")))
			            DrawCeldaDet "'CELDAL7 width20'","", "left", false, EncodeForHtml(d_lookup("descripcion", "tipos_entidades", "codigo like '" & session("ncliente") & "%' and codigo='" & null_s(rst("cambioagrupo")) & "'", session("dsn_cliente")))
				    end if
			    end if%>
			<%else
			    h_ref="javascript:Editar('" & rst("codigo") & "'," & _
			                               p_npagina & ",'" & _
							       p_campo & "','" & _
							       p_criterio & "','" & _
							       replace(p_texto, " ", "%20") & "');"
			    'DrawCeldaHref "CELDAREF","left",false,trimCodEmpresa(rst("codigo")),h_ref                               
                 %><td class="CELDAL7 width10"><%DrawHref "CELDAREF","",trimCodEmpresa(rst("codigo")),h_ref%></td><%
			    DrawCeldaDet "'CELDAL7 width20'","", "left", false, EncodeForHtml(null_s(rst("descripcion")))
			    'DrawCelda2 "CELDA", "left", false, rst("tipo")%>
			    <td class="CELDAL7 width20"><%=iif(rst("tipo")&""="COMPANYGROUP",LitcompanyGroup ,iif(rst("tipo")&""="ELABORATION",UCase(LitElaboration),iif(rst("tipo")&""="REASON_REGULAR",UCase(LitReasonRegularization), EncodeForHtml(rst("tipo")) )))%></td>
			    <%if paso then
			    DrawCeldaDet "'CELDAL7 width20'","", "left", false, EncodeForHtml(null_s(rst("diaspermanencia")))
			    DrawCeldaDet "'CELDAL7 width20'","", "left", false, EncodeForHtml(d_lookup("descripcion", "tipos_entidades", "codigo like '" & session("ncliente") & "%' and codigo='" & rst("cambioagrupo") & "'", session("dsn_cliente")))
			    end if
			end if
            i = i + 1
            rst.MoveNext
        wend
        'rst.Close%>
    </table>

    <%if mode<>"edit" and rst.RecordCount>NumReg then
        if clng(p_npagina) >1 then %>
            <a class="CABECERA" href="tipo_entidad.asp?pagina=anterior&npagina=<%=EncodeForHtml(null_s(cstr(p_npagina)))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%=EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
            <IMG SRC="<%=themeIlion %><%=ImgAnterior%>" align="top" ALT="<%=LitAnterior%>"></a>
        <%end if
        texto=LitPagina + " " + cstr(p_npagina)+ " "+ LitDe + " " + cstr(rst.PageCount)%>
  	    <font class="CELDA"> <%=EncodeForHtml(texto)%> </font> <%
        if clng(p_npagina)<rst.PageCount then %>
            <a class="CABECERA" href="tipo_entidad.asp?pagina=siguiente&npagina=<%=EncodeForHtml(null_s(cstr(p_npagina)))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
            <IMG SRC="<%=themeIlion %><%=ImgSiguiente%>" align="top" ALT="<%=LitSiguiente%>"></a>
        <%end if%>
	    <font class="CELDA">&nbsp;&nbsp; <%=LitPagIrA%> <input class="CELDA" type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;<a class="CELDAREF" href="javascript:IrAPagina(2,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina')"><%=LitIr%></a></font><%
	    rst.Close
    end if%>
    <br>
    <%if mode<>"edit" then %>
        <hr>
         <table class="width100 md-table-responsive" BORDER="0" CELLSPACING="1" CELLPADDING="1">
                <%DrawceldaDet "'ENCABEZADOL underOrange width50'", "", "left", true,"<span><b>" & LitNBregistro & "</b></span>"%>
        </table>
	    <table class="width100 underOrange md-table-responsive bCollapse">
            <tr class="underOrange">
            <%
                DrawceldaDet "'ENCABEZADOL underOrange width10'","", "left", true,"<span><b>" & LitCodigo & "</b></span>"
                DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<span><b>" & LitDescripcion & "</b></span>"
		        DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<span><b>" & LitTipoDe & "</b></span>"
		        DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<span style='display:none;' id='cabdias'><b> " & LitDiasPerm & "</b></span>"
		        DrawceldaDet "'ENCABEZADOL underOrange width20'","", "left", true,"<span style='display:none;' id='cabgrupo'><b>" & LitGrupoComercial & "</b></span>"
            %>
            </tr>
            <tr>
                <td class="CELDA underOrange width10"><input class="width70" type='text' name='i_codigo' value='' maxlength="5" size="5"></td>
                <td class="CELDA underOrange width20"><input class="width70" type='text' name='i_descripcion' value='' maxlength="50" size="50"></td>
                <td class="CELDA underOrange width20">
                    <select class="width60" id="tipo" name="i_tipo" onchange="javascript:VerGrupos();">
                        <option value="CLIENTE"><%=LitTCliente%></option>
			  		    <option value="PROVEEDOR"><%=LitTProveedor%></option>
					    <option value="CONTACTO COMERCIAL"><%=LitTContacto%></option>
					    <option value="PERSONAL"><%=LitTPersonal%></option>
					    <option value="CENTRO"><%=LitTCentro%></option>
					    <option value="ARTICULO"><%=LitTArticulo%></option>
					    <option value="GRUPO CONT.COMERCIAL"><%=LitGrupoComercial%></option>
                        <option value="COMPANYGROUP"><%=LitcompanyGroup%></option>
                        <option value="REASON_REGULAR"><%=UCase(LitReasonRegularization)%></option>
                        <%if ModuloContratado(session("ncliente"),ModQSR) <> 0 then%>
                        <option value="ELABORATION"><%=UCase(LitElaboration)%></option>
                        <%end if%>
		            </select>
		        </td>
		        <td class="CELDA underOrange width20" ><input class="width70" id="dias" type='text' style="display:none;" name='i_diasperm' value='0' maxlength="3" size="3"></td>
			    <%rstAux.cursorlocation=3
				rstAux.open "select codigo, descripcion from tipos_entidades with(nolock) where codigo like '" & session("ncliente") & "%' and tipo='" & LitGrupoComercial & "' order by codigo",session("dsn_cliente")%>
				<td class="CELDA underOrange width20" style="text-align:left">
				<select id="grupo" style="display:none" class="width60" name="i_cambioagrupo">
				<%while not rstAux.eof%>
				    <option value="<%=EncodeForHtml(null_s(rstAux("codigo")))%>"><%=EncodeForHtml(null_s(rstAux("descripcion")))%></option>
				    <%rstAux.movenext
				wend%>
				<option selected value=""></option>
				</select>
				</td>
				<%rstAux.close%>
           </tr>
	    </table>
        <hr>
        <%set rst = nothing
        set rstAux = nothing
    end if%>
    </form>
<%else
	'MsgError LitSinSesion%>
	<br><a href="../" target="_top">Iniciar sesi�n</a>
<%end if%>
</body>
</html>