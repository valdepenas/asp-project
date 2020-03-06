<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
</HEAD>
<script Language="JavaScript">
function Guardar(param) {
   ok=1;
   parent.pantalla.document.almacenar.refrescar.value="NO";
   switch(param){
      case 1:
        
		 if (isNaN(parent.pantalla.document.almacenar.e_stock.value)) {
			window.alert("<%=LitMsgStockNumerico%>");
			ok=0;
		 }
		 else{
		    if (isNaN(parent.pantalla.document.almacenar.e_smin.value)) {
			   window.alert("<%=LitMsgStockMinNumerico%>");
			   ok=0;
		    }
			else{
			   if (isNaN(parent.pantalla.document.almacenar.e_reposicion.value)) {
			      window.alert("<%=LitMsgReposicionNumerico%>");
			      ok=0;
		       }
			   else{
			      if (isNaN(parent.pantalla.document.almacenar.e_precibir.value)) {
			         window.alert("<%=LitMsgPedRecibirNumerico%>");
			         ok=0;
		          }
				  else{
				     if (isNaN(parent.pantalla.document.almacenar.e_pservir.value)) {
			            window.alert("<%=LitMsgPedServirNumerico%>");
			            ok=0;
		             }
					 else{
					    if (isNaN(parent.pantalla.document.almacenar.e_pmin.value)) {
			               window.alert("<%=LitMsgPedMinimoNumerico%>");
			               ok=0;
		                }
					 }
				  }
			   }
			}
		 }
		 if (ok==1){
   		    if (parent.pantalla.document.almacenar.e_predet.checked==true){
		       if (window.confirm("<%=LitMsgAlmacenPredet%>")==true){
			      parent.pantalla.document.almacenar.h_predet.value="SI";
			   }
		    }
		 }
   		    
		 break;
      case 2:
         if (parent.pantalla.document.almacenar.i_almacen.value=="")  {
            window.alert ("<%=LitMsgAlmacenNoNulo%>");
            ok=0;
         }  
		 
		 if (isNaN(parent.pantalla.document.almacenar.i_stock.value)) {
			window.alert("<%=LitMsgStockNumerico%>");
			ok=0;
		 }
		 else{
		    if (isNaN(parent.pantalla.document.almacenar.i_smin.value)) {
			   window.alert("<%=LitMsgStockMinNumerico%>");
			   ok=0;
		    }
			else{
			   if (isNaN(parent.pantalla.document.almacenar.i_reposicion.value)) {
			      window.alert("<%=LitMsgCantReposicionNumerico%>");
			      ok=0;
		       }
			   else{
			      if (isNaN(parent.pantalla.document.almacenar.i_precibir.value)) {
			         window.alert("<%=LitMsgPedRecibirNumerico%>");
			         ok=0;
		          }
				  else{
				     if (isNaN(parent.pantalla.document.almacenar.i_pservir.value)) {
			            window.alert("<%=LitMsgPedServirNumerico%>");
			            ok=0;
		             }
					 else{
					    if (isNaN(parent.pantalla.document.almacenar.i_pmin.value)) {
			               window.alert("<%=LitMsgPedMinimoNumerico%>");
			               ok=0;
		                }
					 }
				  }
			   }
			}
		 }
		 if (ok==1){
   		    if (parent.pantalla.document.almacenar.i_predet.checked==true){
		       if (window.confirm("<%=LitMsgAlmacenPredet%>")==true){
			      parent.pantalla.document.almacenar.h_predet.value="SI";
			   }
		    }
		 }
		 break;
   }
   if (ok==1) {
      	parent.pantalla.document.almacenar.submit();
	  	document.opciones.submit();
   }
   
}

function Buscar() {
  	    parent.pantalla.document.almacenar.refrescar.value="NO";
		parent.pantalla.document.location="almacenar.asp?mode=search&campo=" + document.opciones.campos.value + 
		"&criterio=" + document.opciones.criterio.value + "&texto=" + document.opciones.texto.value + "&sentido=next&npagina=" + parent.pantalla.document.almacenar.h_npagina.value +
		"&articulo=" + parent.pantalla.document.almacenar.h_articulo.value;
		/* parent.pantalla.document.tipo_pago.submit(); */
		document.location="almacenar_bt.asp";
	
}
function Eliminar(param) {
   parent.pantalla.document.almacenar.refrescar.value="NO";
   if (window.confirm("<%=LitMsgEliminarAlmacenConfirm%>")==true) {

      switch (param){
         case 1:
	        parent.pantalla.document.location="almacenar.asp?mode=delete&almacen=" + parent.pantalla.document.almacenar.h_almacen.value + "&articulo=" + parent.pantalla.document.almacenar.h_articulo.value + "&npagina=" + parent.pantalla.document.almacenar.h_npagina.value;
            break;
		 
         case 2:
		 	if (parent.pantalla.document.almacenar.i_referencia.value=="")  {
               window.alert ("<%=LitMsgReferenciaNoNulo%>");
            }
            else {     
               parent.pantalla.document.location="almacenar.asp?mode=delete&almacen=" + parent.pantalla.document.almacenar.h_almacen.value + "&articulo=" + parent.pantalla.document.almacenar.h_articulo.value + "&npagina=" + parent.pantalla.document.almacenar.h_npagina.value;
            }   
    		break;
      }   
      document.location="almacenar_bt.asp";
   }
	
}

function Cancelar() {
  	    parent.pantalla.document.almacenar.refrescar.value="NO";
		parent.pantalla.document.location="almacenar.asp?npagina="+parent.pantalla.document.almacenar.h_npagina.value + "&articulo=" + parent.pantalla.document.almacenar.h_articulo.value;
		        
		document.location="almacenar_bt.asp";
	
}

</script>

<body leftmargin="<%=LitLeftPosBT%>" topmargin="<%=LitTopPosBT%>" bgcolor="#000000">
<!--#include file="almacenar.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../tablas.inc" -->

<form name="opciones" method="post">

 <%if request("mode")="edit" then
 		param=1
  else
		param=2
 end if%>
					
	<table width=100% BORDER="0" CELLSPACING="1" CELLPADDING="1">
		<tr>
			<td CLASS=CELDABOT>
				<A CLASS=CELDAREF href="javascript:Guardar(<%=param%>);"><IMG HEIGHT=16 WIDTH=16 HSPACE=2 ALIGN="top" SRC="../images/save.gif" BORDER=0 ALT=<%=LitGuardar%>><%=LitGuardar%></A>
			</td>
			<%if request("mode")="edit" then%>
			<td CLASS=CELDABOT>
			    <A CLASS=CELDAREF href="javascript:Eliminar(<%=param%>);"><IMG HEIGHT=16 WIDTH=16 HSPACE=2 ALIGN="top" SRC="../images/del.gif" BORDER=0 ALT="<%=LitBorrar%>"><%=LitBorrar%></A>
			</td>
			<%end if%>
			<td CLASS=CELDABOT>
            	<A CLASS=CELDAREF href="javascript:Cancelar();"><IMG HEIGHT=16 WIDTH=16 HSPACE=2 ALIGN="top" SRC="../images/cncl.gif" BORDER=0 ALT=<%=LitCancelar%>><%=LitCancelar%></A>
			</td>
			
			<td CLASS=CELDABOT>
				<SELECT class=INPUT name="campos">
					<OPTION selected value="descripcion"><%=LitAlmacen%></OPTION>
				</SELECT>
			</td>
			<td CLASS=CELDABOT>
				<SELECT class=INPUT name="criterio">
					<OPTION value="contiene"><%=LitContiene%></OPTION>
					<OPTION value="empieza"><%=LitComienza%></OPTION>
					<OPTION value="termina"><%=LitTermina%></OPTION>
					<OPTION value="igual"><%=LitIgual%></OPTION>
				</SELECT>
			</td>
			<td CLASS=CELDABOT>
				<INPUT class=INPUT type="text" name="texto" size=15 maxLength=15 value="">
			</td>
			<td CLASS=CELDABOT>
			   <A CLASS=CELDAREF href="javascript:Buscar();"><IMG HEIGHT=16 WIDTH=16 HSPACE=2 ALIGN="top" SRC="../images/search.gif" BORDER=0 ALT="<%=LitBuscar%>"><%=LitBuscar%></A>
			</td>
		</tr>
	</table>
	
</form>
</form>
</BODY>
</HTML>
