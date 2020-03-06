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
'JMAN 13-06-03: Migración a monobase'
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
<!--#include file="../modulos.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<TITLE><%=LitTituloCaj%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<%si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)%>
<%si_tiene_modulo_petroleos= ModuloContratado(session("ncliente"),ModOrCU)%>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
function Insertar() {
	if (document.cajas.i_codigo.value==""){
		window.alert("<%=LitNoCodigo%>");
		return;
	}

	if(document.cajas.i_descripcion.value==""){
		window.alert("<%=LitNoDescr%>");
		return;
	}

	//Recargar el submarco de detalles
	fr_Tabla.document.cajas_det.action="cajas_det.asp?mode=save" +
		"&i_codigo=" + document.cajas.i_codigo.value +
		"&i_descripcion=" + document.cajas.i_descripcion.value +
		"&i_serie=" + document.cajas.i_serie.value +
		"&i_cuenta=" + document.cajas.i_cuenta.value +
		<%if si_tiene_modulo_tiendas<>0 then%>
	    	"&i_tapunte=" + document.cajas.i_tapunte.value +
	    <%end if%>
        "&i_tienda=" + document.cajas.i_tienda.value +
        "&i_seriefraord=" + document.cajas.i_seriefraord.value +
        <%if si_tiene_modulo_petroleos<>0 then%>
            "&i_seriefrarect=" + document.cajas.i_seriefrarect.value +
            "&i_seriefracont=" + document.cajas.i_seriefracont.value;
        <%else%>
            "&i_seriefrarect=" + document.cajas.i_seriefrarect.value;
        <%end if%>

	fr_Tabla.document.cajas_det.submit();
	
	
}
function clearForm(){
    //Limpiar los campos del formulario
    document.cajas.i_descripcion.value="";
	document.cajas.i_codigo.value="";
	document.cajas.i_serie.value="";
	document.cajas.i_cuenta.value="";
	<%if si_tiene_modulo_tiendas<>0 then%>
		document.cajas.i_tapunte.value="";
	<%end if%>
	document.cajas.i_tienda.value="";
	//Colocar el foco en el campo de cantidad.
	document.cajas.i_codigo.focus();
    document.cajas.i_seriefraord.value="";
    document.cajas.i_seriefrarect.value = "";
    <%if si_tiene_modulo_petroleos<>0 then%>
        document.cajas.i_seriefracont.value="";
	<%end if%>
}
function Mas(sentido,lote, texto) {
	document.getElementById("barras").style.display="none";
	fr_Tabla.document.cajas_det.action="cajas_det.asp?mode=ver&sentido=" + sentido + "&lote=" + lote + "&texto=" + texto;
	fr_Tabla.document.cajas_det.submit();
}

function Resize()
{
    var alto = 0;
    if (parent.document.body.offsetHeight) alto = parent.document.body.offsetHeight;
    else alto = parent.self.innerHeight;

	if (alto > 250)
    {
        if (alto - 220 > 250) document.getElementById("frtabla").style.height = alto - 220;
        else document.getElementById("frtabla").style.height = 250;
    }
    else document.getElementById("frtabla").style.height = 250;
}
</script>
<body onload="self.status=''" bgcolor="<%=color_blau%>" onresize="javascript:Resize();">
<%
'*************************************************************************************************************
' CODIGO PRINCIPAL DE LA PAGINA  *****************************************************************************
'*************************************************************************************************************

if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<form name="cajas" method="post" action="cajas.asp">

    <%PintarCabecera "cajas.asp"
    set rst = server.CreateObject("ADODB.Recordset")
    set rstAux = server.CreateObject("ADODB.Recordset")

	Alarma "cajas.asp"%>
    
    <%if request("mode")<>"edit" then
       
        
        %>
	    <!--<table BORDER="0" CELLSPACING="1" CELLPADDING="1">-->
        <table class="width90 underOrange md-table-responsive bCollapse">
            <%Drawfila color_terra
                    %><td class="ENCABEZADOL underOrange width5" ><b><%=LitCodigo%></b></td>
				    <td class="ENCABEZADOL underOrange width15" ><b><%=LitDescripcion%></b></td>
				    <td class="ENCABEZADOL underOrange width10" ><b><%=LitSerieFraSim%></b></td>
				    <td class="ENCABEZADOL underOrange width10" ><b><%=LitCuenta%></b></td>
                <%if si_tiene_modulo_tiendas<>0 then%>
					        <td class="ENCABEZADOL underOrange  width10" ><b><%=LitTapunte%></b></td>
                <%end if %>
					        <td class="ENCABEZADOL underOrange width10" ><b><%=LitTienda%></b></td>
					        <td class="ENCABEZADOL underOrange width10" ><b><%=LitSerieFraOrd%></b></td>
					        <td class="ENCABEZADOL underOrange width10" ><b><%=LitSerieFraRect%></b></td>
                <%if si_tiene_modulo_petroleos<>0 then%>
					        <td class="ENCABEZADOL underOrange width10" ><b><%=LitSerieFraCont%></b></td>
                <%end if%>
                            <td class="ENCABEZADOR underOrange width5" ><b><%=LitSaldo%></b></td>
					        <td class="ENCABEZADOL underOrange width5" >&nbsp</td><%
              CloseFila%>
            <tr>
                <td class='CELDAL7 underOrange width5'>
                    <input class="CELDAL7 width100" name="i_codigo">
                </td>
				<td class='CELDAL7 underOrange width15'>					
                    <textarea class="CELDAL7 width100" name="i_descripcion"></textarea>
				</td>
                <td class='CELDAL7 underOrange width10'>
				<%
                set conn=  server.CreateObject("ADODB.Connection")
                set cmd=  server.CreateObject("ADODB.Command")
                conn.open session("dsn_cliente")
                cmd.ActiveConnection=conn
                conn.cursorlocation=3
                cmd.CommandText="getAllSeriesByTypeDocument"
	            cmd.CommandType = adCmdStoredProc 
                cmd.Parameters.Append cmd.CreateParameter("@ncompany", adVarChar, , 5, session("ncliente"))
                cmd.Parameters.Append cmd.CreateParameter("@type_document", adVarChar, , 50, "TICKET")
                set rstSerie=cmd.execute
				'rstAux.open "select nserie,nombre from series with(nolock) where tipo_documento='TICKET' and nserie like '" & session("ncliente") & "%'",session("dsn_cliente")
				'DrawSelectCelda "CELDA","110","",0,"","i_serie",rstSerie,"","nserie","nombre","",""
                DrawSelect "'CELDAL7 width100'", "","i_serie",rstSerie,"","nserie","nombre","",""
				'rstAux.close%>
                </td>
				<td class='CELDAL7 underOrange width10'>
					<!--<textarea class=CELDAL7 name="i_cuenta" style="width: 100px;"></textarea>-->
                    <textarea class='CELDAL7 width100' name="i_cuenta"></textarea>
				</td>
                <td class='CELDAL7 underOrange width10'>
                    <%if si_tiene_modulo_tiendas<>0 then
					rstAux.cursorlocation=3
					rstAux.open "select codigo,descripcion from tipo_apuntes with(nolock) where codigo<>'" & session("ncliente") & "01' and codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente")
					'DrawSelectCelda "CELDA","110","",0,"","i_tapunte",rstAux,"","codigo","descripcion","",""
                     DrawSelect "'CELDAL7 width100'", "","i_tapunte",rstAux,"","codigo","descripcion","",""
					rstAux.close
				    end if%>

                </td>
                <td class='CELDAL7 underOrange width10'>
                    <%rstAux.cursorlocation=3
				rstAux.open "select codigo,descripcion from tiendas with(nolock) where codigo like '" & session("ncliente") & "%'",session("dsn_cliente")
				'DrawSelectCelda "CELDA","110","",0,"","i_tienda",rstAux,"","codigo","descripcion","",""
                DrawSelect "'CELDAL7 width100'", "","i_tienda",rstAux,"","codigo","descripcion","",""
				rstAux.close%>
                </td>
                <td class='CELDAL7 underOrange width10'>
                    <% if not rstSerie.bof then
                    rstSerie.movefirst
                end if
                'DrawSelectCelda "CELDA","110","",0,"","i_seriefraord",rstSerie,"","nserie","nombre","",""
                 DrawSelect "'CELDAL7 width100'", "","i_seriefraord",rstSerie,"","nserie","nombre","",""%>
                </td>
                <td class='CELDAL7 underOrange width10'>
                    <%if not rstSerie.bof then
                    rstSerie.movefirst
                end if
                'DrawSelectCelda "CELDA","110","",0,"","i_seriefrarect",rstSerie,"","nserie","nombre","",""
                 DrawSelect "'CELDAL7 width100'", "","i_seriefrarect",rstSerie,"","nserie","nombre","",""      
                        %>
                </td>
                <td class='CELDAL7 underOrange width10'>
                    <% if si_tiene_modulo_petroleos<>0 then
                    if not rstSerie.bof then
                        rstSerie.movefirst
                    end if
                    'DrawSelectCelda "CELDA","110","",0,"","i_seriefracont",rstSerie,"","nserie","nombre","",""
                     DrawSelect "'CELDAL7 width100'", "","i_seriefracont",rstSerie,"","nserie","nombre","",""
                end if
                conn.close
                set rstSerie=nothing
                set cmd=nothing
                set conn=nothing%>
                </td>
                <td class='CELDAL7 underOrange width5'>
                    &nbsp; 
                </td>
                <td class='CELDAR7 underOrange width5' >
				   <a href="javascript:Insertar();" class="ic-accept NoMTop" ><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgNuevo%> alt="<%=LitNuevo1%>"></a>
				</td>
            </tr>
			<%'closefila%>
        </table>
	    <script language="javascript">
	        document.cajas.i_codigo.focus();
	    </script>
   <%end if
   if request("mode")<>"edit" then%>
      <iframe id="frtabla" name="fr_Tabla" src='cajas_det.asp?mode=browse' class="width90 iframe-data md-table-responsive" height='250' frameborder="yes" noresize="noresize"></iframe>
        <script language="javascript">
            Resize();
	    </script>
   <%end if%>
   <table width="750">
        <%DrawFila ""%>
			<td class=CELDA7 width="250">
				<SPAN ID="barras" STYLE="display:none">
				</SPAN>
			</td>
		<%CloseFila%>
   </table>
<%' En el fichero BORRADO a partir de aqui'%>
</form>
<%else
	MsgError LitSinSesion
	%><br><a href="../" target="_top">Iniciar sesión</a><%
end if

set rst=nothing
set rstAux=nothing%>
</BODY>
</HTML>