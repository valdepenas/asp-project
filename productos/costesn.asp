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
<%'   JCI 17/06/2003 : MIGRACION A MONOBASE%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../catFamSubResponsive.inc" -->
<!--#include file="costes.inc" -->
<!--#include file="../common/campospersoResponsive.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/tabs.js.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<title><%=LitTitulo%></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
</head>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('Parametros', 'fade=1');
    animatedcollapse.addDiv('Detalle', 'fade=1');
    animatedcollapse.ontoggle = function (jQuery, divobj, state) { }
    animatedcollapse.init();
/*
if (window.document.addEventListener) {
    window.document.addEventListener("keydown", callkeydownhandler, false);
} else {
    window.document.attachEvent("onkeydown", callkeydownhandler);
}
function callkeydownhandler(evnt) {
    ev = (evnt) ? evnt : event;
    CampoRefPulsado('marcoCostes','costes',ev);
}
*/
function comprobar_enter()
{
    if (document.costes.RefFij.value!=""){
        CampoRefPulsado('marcoCostes','costes','');
    }
    else{
        if (document.costes.RefPro.value!=""){
            Cargar_B();
        }
    }
}
function CampoRefPulsado(marco,formulario,e){
    //var keycode = e.keyCode;
    //if (keycode==13){
    //    continuar=0;
    //    if (document.costes.RefFij.value!=""){
    //        continuar=1;
    //    }
    //    if (continuar==1){
            //ricardo 21-8-2007 se comenta estas tres lineas para que no se marque liberar pvp
            //document.costes.updpvp.checked=true;
            //document.costes.updpvpValor.value=1;
            //marcoCostes.document.costes_datos.updPvp.value=1;
            
			marcoCostes.document.getElementById("waitBoxOculto").style.visibility="visible";
			document.getElementById("barras").style.display="none";
			document.costes.target=marcoCostes.name;
			//document.costes.action="costes_datosn.asp?mode=ver";
			//document.costes.action="costes_datosn.asp?mode=creartab&camb=1&hNRegs="+marcoCostes.document.costes_datos.hNRegs.value;
			//document.costes.submit();
			cadena="&RefFij="+document.costes.RefFij.value;
			cadena=cadena+"&ndocumento=";//+document.costes.ndocumento.value;
			cadena=cadena+"&familia=";//+document.costes.familia.value;
			cadena=cadena+"&familia_padre=";//+document.costes.familia_padre.value;
			cadena=cadena+"&categoria=";//+document.costes.categoria.value;
			cadena=cadena+"&updpvpValor="+document.costes.updpvpValor.value;
			cadena=cadena+"&updcosteValor="+document.costes.updcosteValor.value;
			cadena=cadena+"&tipoarticulo=";//+document.costes.tipoarticulo.value;
			cadena=cadena+"&almacen=";//+document.costes.almacen.value;
			cadena=cadena+"&nombreart=";//+document.costes.nombreart.value;
			cadena=cadena+"&referencia=";//+document.costes.referencia.value;
			cadena=cadena+"&hfecha=";//+document.costes.hfecha.value;
			cadena=cadena+"&dfecha=";//+document.costes.dfecha.value;
			cadena=cadena+"&nproveedor=";//+document.costes.nproveedor.value;
			cadena=cadena+"&RefPro=";
			cadena=cadena+"&ordenar=" + document.costes.ordenar.value;
			cadena=cadena+"&de_que_campo_vengo=RefFij";
			marcoCostes.document.costes_datos.action="costes_datosn.asp?mode=creartab2&camb=1"+cadena;
			marcoCostes.document.costes_datos.submit();
			
            //Mostrar('RefFij');
    //    }
    //}
}

function PonerHtml(sentido,lote){
	cadena="";
	cadena="<table width='100%' BORDER='0' cellspacing='1' cellpadding='1'>";
	cadena = cadena + "<tr><td class='MAS'>";
	if (sentido=="next") lote=parseInt(marcoCostes.document.costes.lote.value) + 1;
	if (sentido=="prev") lote=parseInt(marcoCostes.document.costes.lote.value) - 1;
	if (sentido=="nulo") lote=parseInt(marcoCostes.document.costes.lote.value);
	lotes=parseInt(marcoCostes.document.costes.lotes.value);
	varias=false
	if (lote>1){
  		cadena=cadena + "<a class='CELDAREF' href=\"javascript:Mas('prev'," + lote + ");\">";
  		cadena=cadena + "<img height='16' width='16' hspace='2' align='top' src='../images/prev.gif' border='0' alt='<%=LitAnterior%>' title='<%=LitAnterior%>'/></a>";
  		varias=true
	}
	texto="<%=LitPagina%>" + " " + lote+ " " + "<%=LitDe%>" + " " + lotes;
	cadena=cadena + "<font class='CELDA'>" + texto + "</font>";

	if (lote<lotes){
		cadena=cadena + "<a class='CELDAREF' href=\"javascript:Mas('next'," + lote + ");\">";
		cadena=cadena + "<img height='16' width='16' hspace='2' align='top' SRC='../images/next.gif' border='0' alt='<%=LitSiguiente%>' title='<%=LitSiguiente%>'/></a>";
		varias=true
	}

	cadena=cadena + "</td></tr>";
	cadena=cadena + "</table>";
	barras.innerHTML=cadena;
}

function Mas(sentido,lote) {
	document.getElementById("barras").style.display="none";
	marcoCostes.document.costes_datos.action="costes_datosn.asp?mode=ver&sentido=" + sentido + "&lote=" + lote;
	marcoCostes.document.costes_datos.submit();
}

function Mas1(sentido,lote) {
    document.getElementById("barras2").style.display="none";
	marcoResult.document.costes_result.action="costes_resultn.asp?mode=ver&sentido=" + sentido + "&lote=" + lote;
	marcoResult.document.costes_result.submit();
}

function Mostrar(de_que_campo_vengo) {
  if (marcoCostes.document.getElementById("waitBoxOculto").style.visibility != "visible"){
	if (document.costes.nproveedor.value==""&&document.costes.RefPro.value>"")
	{
		window.alert("<%=LitNoProveedor%>");
		return;
	}

	marcoCostes.document.getElementById("waitBoxOculto").style.visibility="visible";
	document.getElementById("barras").style.display="none";
	document.costes.target=marcoCostes.name;
	if (de_que_campo_vengo=="RefFij") document.costes.action="costes_datosn.asp?mode=creartab2&de_que_campo_vengo=" + de_que_campo_vengo;
	else document.costes.action="costes_datosn.asp?mode=creartab&de_que_campo_vengo=" + de_que_campo_vengo;

	document.costes.submit();
    MostrarTABModificar();
  }
}

// jgc - Activa la tab=0 (primera)
function MostrarTABModificar() {
  var $tabs = jQuery('#tabs').tabs(); 
   $tabs.tabs('select', 0); 
} 

// cag
function bloquear(valor){
	/*alert("en bloquear"+document.costes.updcosteValor.value); */ /* siempre vale el valor de configuracion */

	/* para el marco interior pero esto provoca recarga */
    /*document.costes.all("barras").style.display="none";
	document.costes.target=document.frames("frcostes").name;
	document.costes.action="costes_datosn.asp?mode=ver";
	document.costes.submit();*/

	/* Para las opciones actualizacion masiva bajo el marco interior */

	checkDcha = document.costes.updpvp.checked;
	checkIzqda = document.costes.updcoste.checked;

	elementsMarco=marcoCostes.document.costes_datos.hNRegs.value; /* Para ver si actuan los 2 modificadores checkbox de actualizacion */
	p=1;
	activo=0;

	while (p<elementsMarco-1) {
	   if (eval("marcoCostes.document.costes_datos.check" + p + ".checked"))      activo++
	   p++;
	}
	if (activo==0) {

	     if (valor==1) {
			 if  (document.costes.updcosteValor.value==1) {
			  /*RGU 10/1/2007*/
			  document.costes.costegral.value='';
			  document.costes.costeIncDec.value='';
			  /*RGU*/
			  document.costes.costegral.disabled=true;
			  document.costes.costeIncDec.disabled=true;
			 }
			 else {
			  document.costes.costegral.disabled=false;
			  document.costes.costeIncDec.disabled=false;
			 }
		 }
		 if (valor==3) {
			 valorActpvp=eval("marcoCostes.document.costes_datos.updPvp.value")
			 
			  if  (document.costes.updpvpValor.value==1) {
				  /*RGU 10/1/2007*/
				  document.costes.precIncDec.value='';
				  document.costes.precIvaIncDec.value='';
				  document.costes.precgral.value='';
		  		  document.costes.precivagral.value='';
		  		  /*RGU*/
				  document.costes.precgral.disabled=true;
				  document.costes.precivagral.disabled=true;
				  document.costes.precIncDec.disabled=true;
				  document.costes.precIvaIncDec.disabled=true;
			  }
			  else {
			  document.costes.precgral.disabled=false;
			  document.costes.precivagral.disabled=false;
			  document.costes.precIncDec.disabled=false;
			  document.costes.precIvaIncDec.disabled=false;
			 }
		}

		/* Bloqueando inputs del marco interior */
		elementos=marcoCostes.document.costes_datos.hNRegs.value;

		if (valor==1) {
		  if  (document.costes.updcosteValor.value==1) {
			for (i=1;i<=elementos-1;i++) eval("marcoCostes.document.costes_datos.cactual" + i + ".disabled=true");
			document.costes.updcosteValor.value=0;
		  }
		  else {
			for (i=1;i<=elementos-1;i++) eval("marcoCostes.document.costes_datos.cactual" + i + ".disabled=false");
			document.costes.updcosteValor.value=1;
		  }
		 /* if (elementos>0){
			  if (marcoCostes.document.costes_datos.cactual1.disabled==true){
					marcoCostes.document.costes_datos.recargo1.focus();
					marcoCostes.document.costes_datos.recargo1.select();
			  }else{
					marcoCostes.document.costes_datos.cactual1.focus();
					marcoCostes.document.costes_datos.cactual1.select();
			  }
		   }*/
		}
		if (valor==3) {
		
		  if  (document.costes.updpvpValor.value==1) {
			for (i=1;i<=elementos-1;i++) {
					eval("marcoCostes.document.costes_datos.pvp" + i + ".disabled=true");
					eval("marcoCostes.document.costes_datos.pvpiva" + i + ".disabled=true");
			}
			document.costes.updpvpValor.value=0;
		  }
		  else {
			for (i=1;i<=elementos-1;i++) {
				eval("marcoCostes.document.costes_datos.pvp" + i + ".disabled=false");
				eval("marcoCostes.document.costes_datos.pvpiva" + i + ".disabled=false");
			}
			document.costes.updpvpValor.value=1;
		  }
		}
	    marcoCostes.ponDatosCheck();
	}// fin activo=0
	else {
	   window.alert("<%=LitNoPuedeCambiarChecks%>");
	  if ((valor==3) && (checkDcha==true))  document.costes.updpvp.checked=false;
  	  if ((valor==3) && (checkDcha==false))  document.costes.updpvp.checked=true;
  	  if ((valor==1) && (checkIzqda==true))  document.costes.updcoste.checked=false;
  	  if ((valor==1) && (checkIzqda==false))  document.costes.updcoste.checked=true;
	}
}
//fin cag

function seleccionar(marco,formulario,check) {
	nregistros=eval(marco + ".document." + formulario + ".hNRegs.value-1");
	if (eval("document.costes." + check + ".checked")){
		for (i=1;i<=nregistros;i++) {
			nombre="check" + i;
			eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
		}
	}
	else {
		for (i=1;i<=nregistros;i++) {
			nombre="check" + i;
			eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
		}
	}
}

function AplicarCoste(modo) {
	if (modo==1){
		if (!isNaN(document.costes.costegral.value.replace(",",".")) && document.costes.costegral.value!="") {
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
					eval("marcoCostes.document.costes_datos.cactual" + i + ".value=parseFloat(document.costes.costegral.value.replace(',','.'))");
					marcoCostes.calculapvp(i,0,1);
			}
		}
	}
	else{
		if (!isNaN(document.costes.costeIncDec.value.replace(",",".")) && document.costes.costeIncDec.value!="")
		{
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
					cactual= eval("marcoCostes.document.costes_datos.cactual" + i + ".value");
					cactual=parseFloat(cactual.replace(",","."));

					var incdec=parseFloat(document.costes.costeIncDec.value.replace(",","."));
					cactual = parseFloat(cactual) + ((parseFloat(cactual)*parseFloat(incdec))/100);
					cactual=cactual.toFixed(<%=dec_prec%>);

					eval("marcoCostes.document.costes_datos.cactual" + i + ".value=" + parseFloat(cactual.replace(",",".")));
					marcoCostes.calculapvp(i,0,1);
			}
		}
	}
}

function AplicarRecargo(modo) {
	if (modo==1){
		if (!isNaN(document.costes.recgral.value.replace(",",".")) && document.costes.recgral.value!="") {
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
					eval("marcoCostes.document.costes_datos.recargo" + i + ".value=parseFloat(document.costes.recgral.value.replace(',','.'))");
					marcoCostes.calculapvp(i,0);
			}
		}
	}
	else{
		if (!isNaN(document.costes.recIncDec.value.replace(",",".")) && document.costes.recIncDec.value!="")
		{
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
				/*if (eval("marcoCostes.document.costes_datos.check" + i + ".checked")){*/
					recargo=eval("marcoCostes.document.costes_datos.recargo" + i + ".value");
					recargo=parseFloat(recargo.replace(",","."));
					var incdec=parseFloat(document.costes.recIncDec.value.replace(",","."));
					recargo= parseFloat(recargo) + parseFloat(incdec);
					recargo=recargo.toFixed(2);
					//recargo=recargo.replace(".","");

					eval("marcoCostes.document.costes_datos.recargo" + i + ".value=" + parseFloat(recargo.replace(",",".")));
					marcoCostes.calculapvp(i,0);
				//}
			}
		}
	}
}

function AplicarPrec(modo) {
	if (modo==1){
		if (!isNaN(document.costes.precgral.value.replace(",",".")) && document.costes.precgral.value!="") {
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
				//if (eval("marcoCostes.document.costes_datos.check" + i + ".checked")){
					eval("marcoCostes.document.costes_datos.pvp" + i + ".value=parseFloat(document.costes.precgral.value.replace(',','.'))");
					marcoCostes.calcularecargo(i);
				//}
			}
		}
	}
	else{
		if (!isNaN(document.costes.precIncDec.value.replace(",",".")) && document.costes.precIncDec.value!="")
		{
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
						pvp=eval("marcoCostes.document.costes_datos.pvp" + i + ".value");
						pvp=parseFloat(pvp.replace(",","."));
						var incdec=parseFloat(document.costes.precIncDec.value.replace(",","."));
						pvp= parseFloat(pvp) + ((parseFloat(pvp)*parseFloat(incdec))/100);
						pvp=pvp.toFixed(<%=DEC_PREC%>);

						eval("marcoCostes.document.costes_datos.pvp" + i + ".value=" + parseFloat(pvp.replace(",",".")));
						marcoCostes.calcularecargo(i);
			}
		}

	}
}

function AplicarPrecIva(modo) {
	if (modo==1){
		if (!isNaN(document.costes.precivagral.value.replace(",",".")) && document.costes.precivagral.value!="") {
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
					eval("marcoCostes.document.costes_datos.pvpiva" + i + ".value=parseFloat(document.costes.precivagral.value.replace(',','.'))");
					marcoCostes.calculapvpsiniva(i);
			}
		}
	}
	else{
		if (!isNaN(document.costes.precIvaIncDec.value.replace(",",".")) && document.costes.precIvaIncDec.value!="")
		{
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
					pvp=eval("marcoCostes.document.costes_datos.pvpiva" + i + ".value");
					pvp=parseFloat(pvp.replace(",","."));

					var incdec=parseFloat(document.costes.precIvaIncDec.value.replace(",","."));
					pvp= parseFloat(pvp) + ((parseFloat(pvp)*parseFloat(incdec))/100);
					pvp=pvp.toFixed(<%=DEC_PREC%>);

					eval("marcoCostes.document.costes_datos.pvpiva" + i + ".value=" + parseFloat(pvp.replace(",",".")));
					marcoCostes.calculapvpsiniva(i);
			}
		}
	}
}

function AplicarMargen(modo) {
	if (modo==1){
		if (!isNaN(document.costes.margengral.value.replace(",",".")) && document.costes.margengral.value!="") {
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
					valor_a_poner=parseFloat(document.costes.margengral.value.replace(',','.'));
					eval("marcoCostes.document.costes_datos.margen" + i + ".value=valor_a_poner");
					marcoCostes.calculapvp(i,1);
			}
		}
	}
	else{
		if (!isNaN(document.costes.margenIncDec.value.replace(",",".")) && document.costes.margenIncDec.value!="")
		{
			elementos=marcoCostes.document.costes_datos.hNRegs.value;
			for (i=1;i<=elementos-1;i++) {
					margen=eval("marcoCostes.document.costes_datos.margen" + i + ".value");
					margen=parseFloat(margen.replace(",","."));

					var incdec=parseFloat(document.costes.margenIncDec.value.replace(",","."));
					margen= parseFloat(margen) + parseFloat(incdec);
					margen=margen.toFixed(2);

					eval("marcoCostes.document.costes_datos.margen" + i + ".value=" + parseFloat(margen.replace(",",".")));
					marcoCostes.calculapvp(i,1);
			}
		}
	}
}

function TraerProveedor(mode) {
//cag
//	document.location.href="costes.asp?nproveedor=" + document.costes.nproveedor.value + "&mode=" + mode + "&referencia=" + document.costes.referencia.value + "&nombreart=" + document.costes.nombreart.value + "&familia=" + document.costes.familia.value + "&almacen=" + document.costes.almacen.value + "&ordenar=" + document.costes.ordenar.value + "&dfecha=" + document.costes.dfecha.value + "&hfecha=" + document.costes.hfecha.value;
	document.location.href="costesn.asp?nproveedor=" + document.costes.nproveedor.value + "&mode=" + mode + "&referencia=" + document.costes.referencia.value + "&nombreart=" + document.costes.nombreart.value + "&familia=" + document.costes.familia.value + "&almacen=" + document.costes.almacen.value + "&ordenar=" + document.costes.ordenar.value + "&dfecha=" + document.costes.dfecha.value + "&hfecha=" + document.costes.hfecha.value+ "&tipoarticulo=" + document.costes.tipoarticulo.value;

//fincag
}
//----------------------------------------
//Funcion para guardar registro en la pantalla de resultados
//----------------------------------------
function Seleccionar2() {
   if ((document.costes.control.value==0)&&  (marcoCostes.document.getElementById("waitBoxOculto").style.visibility != "visible")){
   		document.costes.control.value=1;
		elementos=marcoCostes.document.costes_datos.hNRegs.value;
		if (elementos=="") elementos=0;
		error="NO";
		msg="";
		tiene=0;
		for (i=1;i<=elementos-1;i++) {

			if (eval("marcoCostes.document.costes_datos.check" + i + ".checked")){
				if (eval("isNaN(marcoCostes.document.costes_datos.cactual" + i + ".value.replace(',','.')) && marcoCostes.document.costes_datos.cactual" + i + ".value!=''")) {
					if (tiene==1) {msg=msg + " <%=LitY%> ";}
					msg=msg + "<%=LitMsgCosteLinea%> " + i + " <%=LitIncorr%>";
					error="SI";
					tiene=1;
				}

				if (eval("isNaN(marcoCostes.document.costes_datos.recargo" + i + ".value.replace(',','.')) && marcoCostes.document.costes_datos.recargo" + i + ".value!=''")) {
					if (tiene==1) {msg=msg + " <%=LitY%> ";}
					msg=msg + "<%=LitMsgRecargoLinea%> " + i + " <%=LitIncorr%>";
					error="SI";
					tiene=1;
				}

				if (eval("isNaN(marcoCostes.document.costes_datos.pvp" + i + ".value.replace(',','.')) && marcoCostes.document.costes_datos.pvp" + i + ".value!=''")) {
					if (tiene==1) {msg=msg + " <%=LitY%> ";}
					msg=msg + "<%=LitMsgPVPLinea%> " + i + " <%=LitIncorr%>";
					error="SI";
					tiene=1;
				}
			}
		}
		if (error=="SI") window.alert(msg);
		else {
			if (elementos>0) {
				marcoCostes.document.costes_datos.action="costes_datosn.asp?mode=save";
				marcoCostes.document.costes_datos.submit();
			}
		}
	}
}

function DeSeleccionar2(){
	if ((document.costes.control.value==0)&&  (marcoCostes.document.getElementById("waitBoxOculto").style.visibility != "visible"))
	{
		document.costes.control.value=1;
		marcoResult.document.costes_result.action="costes_resultn.asp?mode=deselect";
		marcoResult.document.costes_result.submit();
		marcoCostes.document.costes_datos.action="costes_datosn.asp?mode=edit";
		marcoCostes.document.costes_datos.submit();
	}
}
function GuardarValores()
{
    if(document.costes.updcoste.checked==true) m_coste="1";
    else m_coste="0";
    
    if(document.costes.updpvp.checked==true) m_pvp="1";
    else m_pvp="0";
    
    //comprobamos si estamos en modo multiempresa
    if (parseInt(nn)>1){
        //multiempresa
        pagina="costesn_multiempresa.asp?fv="+document.costes.fv.value+"&updcoste="+m_coste+"&updpvp="+m_pvp;
        Ven=AbrirVentanaRef(pagina,"P",<%=AltoVentana%>,<%=AnchoVentana%>);
    }
    else{
        //monoempresa
        if ((document.costes.control.value==0)&&  (marcoCostes.document.getElementById("waitBoxOculto").style.visibility != "visible"))
        {
            var hoy= new Date();
            if (marcoResult.document.costes_result.nact.value>0)
            {
                if ( checkdate(document.costes.fv) )
                {
  	                if (document.costes.fv.value != "")
                        {
	                    if ( DiferenciaTiempo (document.costes.fv.value, hoy.getDate()+"/"+(hoy.getMonth()+1)+"/"+hoy.getFullYear(), "dias") >0 )
                        {
		                    if(window.confirm("<%=LitMsgCrearAct%>")==true)
                            {
			                    document.costes.control.value=1;
                    			document.location="costesnResultado.asp?mode=save&fv="+document.costes.fv.value;
			                    parent.botones.document.location="costes_bt.asp";
		                    }
	                    }
	                    else window.alert("<%=LitMsgFechaMayor%>");
	                }
	                else
	                {
		                if(window.confirm("<%=LitMsgCrear%>")==true)
		                {
			                cadena= window.prompt(" <%=LitMsgObservaciones%> ","");
				            if (window.confirm(" <%=LitMsgCrear2%> "))
			                {
				                document.location="costesnResultado.asp?mode=save&obs="+cadena+"&fv="+document.costes.fv.value;
				                parent.botones.document.location="costes_bt.asp?mode=impresion";
			                }
		                }
	                }
                }
                else window.alert("<%=LitMsgErrFecha%>");
            }
            else window.alert("<%=LitMsgNoreg%>");
        }
    }
}
/*
function tier1Menu(objMenu,objImage)
{
	if (objMenu.style.display == "none") {
		objMenu.style.display = "";
		objImage.src = "../Images/<%=ImgCarpetaAbierta%>";


		switch (objMenu.id) {
			case "Modificar":
				document.getElementById("ACTPENDIENTE").style.display="none";
				document.getElementById("REVACT").style.display="none";
				document.getElementById("img2").src="../Images/<%=ImgCarpetaCerrada%>";
				document.getElementById("img3").src="../Images/<%=ImgCarpetaCerrada%>";
				break;
			case "ACTPENDIENTE":
				document.getElementById("Modificar").style.display="none";
				document.getElementById("REVACT").style.display="none";
				document.getElementById("img1").src="../Images/<%=ImgCarpetaCerrada%>";
				document.getElementById("img3").src="../Images/<%=ImgCarpetaCerrada%>";
				marcoActualizaciones.document.CostesnActualizaciones.action="CostesnActualizaciones.asp?mode=ver"
				marcoActualizaciones.document.CostesnActualizaciones.submit();
				break;
			case "REVACT":
				document.getElementById("ACTPENDIENTE").style.display="none";
				document.getElementById("Modificar").style.display="none";
				document.getElementById("img2").src="../Images/<%=ImgCarpetaCerrada%>";
				document.getElementById("img1").src="../Images/<%=ImgCarpetaCerrada%>";
				marcoRevActualizaciones.document.CostesnRevActualizaciones.action="CostesnActualizaciones.asp?mode=revisar"
				marcoRevActualizaciones.document.CostesnRevActualizaciones.submit();
				break;
		}
	}
	else {
		objMenu.style.display = "none";
		objImage.src = "../images/<%=ImgCarpetaCerrada%>";
	}

}
*/
function borrarActPre(fecha)
{
  window.alert(fecha);
}


function BorrarActPre2 ()
{
	if (marcoRevActualizaciones.document.CostesnRevActualizaciones.nregistros.value!=0 && marcoRevActualizaciones.document.CostesnRevActualizaciones.nregistros.value!="")
	{
		marcoRevActualizaciones.document.CostesnRevActualizaciones.action="CostesnActualizaciones.asp?mode=deleterev";
		marcoRevActualizaciones.document.CostesnRevActualizaciones.submit();
	}
}

function ActPre()
{
	if (marcoRevActualizaciones.document.CostesnRevActualizaciones.nregistros.value!=0 && marcoRevActualizaciones.document.CostesnRevActualizaciones.nregistros.value!="")
	{
		elementos=marcoRevActualizaciones.document.CostesnRevActualizaciones.nregistros.value;
		if (elementos=="") elementos=0;
		error="NO";
		msg="";
		tiene=0;
		for (i=1;i<=elementos-1;i++) {

			if (eval("marcoRevActualizaciones.document.CostesnRevActualizaciones.check" + i + ".checked")){

				if (eval("isNaN(marcoRevActualizaciones.document.CostesnRevActualizaciones.importe" + i + ".value.replace(',','.')) && marcoRevActualizaciones.document.CostesnRevActualizaciones.importe" + i + ".value!=''")) {
					if (tiene==1) {msg=msg + " <%=LitY%> ";}
					msg=msg + "<%=LitMsgCosteLinea%> " + i + " <%=LitIncorr%>";
					error="SI";
					tiene=1;
				}

				if (eval("isNaN(marcoRevActualizaciones.document.CostesnRevActualizaciones.recargo" + i + ".value.replace(',','.')) && marcoRevActualizaciones.document.CostesnRevActualizaciones.recargo" + i + ".value!=''")) {
					if (tiene==1) {msg=msg + " <%=LitY%> ";}
					msg=msg + "<%=LitMsgRecargoLinea%> " + i + " <%=LitIncorr%>";
					error="SI";
					tiene=1;
				}

				if (eval("isNaN(marcoRevActualizaciones.document.CostesnRevActualizaciones.pvp" + i + ".value.replace(',','.')) && marcoRevActualizaciones.document.CostesnRevActualizaciones.pvp" + i + ".value!=''")) {
					if (tiene==1) {msg=msg + " <%=LitY%> ";}
					msg=msg + "<%=LitMsgPVPLinea%> " + i + " <%=LitIncorr%>";
					error="SI";
					tiene=1;
				}
			}
		}
		if (error=="SI") window.alert(msg);
		else {
			marcoRevActualizaciones.document.CostesnRevActualizaciones.action="CostesnActualizaciones.asp?mode=save";
			marcoRevActualizaciones.document.CostesnRevActualizaciones.submit();
		}
	}
}

function Cargar_B() {
	//if (window.event.keyCode=="13" && document.costes.RefPro.value!="")
    if (document.costes.RefPro.value!="")
	{
		if (marcoCostes.document.getElementById("waitBoxOculto").style.visibility != "visible")
		{

		  if (document.costes.nproveedor.value=="") window.alert("<%=LitNoProveedor%>");
		  else{
			marcoCostes.document.getElementById("waitBoxOculto").style.visibility="visible";
			document.getElementById("barras").style.display="none";
			document.costes.target=marcoCostes.name;
			//document.costes.action="costes_datosn.asp?mode=ver";
			//document.costes.action="costes_datosn.asp?mode=creartab&camb=1&hNRegs="+marcoCostes.document.costes_datos.hNRegs.value;
			//document.costes.submit();
			cadena="&RefPro="+document.costes.RefPro.value;
			cadena=cadena+"&ndocumento="+document.costes.ndocumento.value;
			cadena=cadena+"&familia="+document.costes.familia.value;
			cadena=cadena+"&familia_padre="+document.costes.familia_padre.value;
			cadena=cadena+"&categoria="+document.costes.categoria.value;
			cadena=cadena+"&updpvpValor="+document.costes.updpvpValor.value;
			cadena=cadena+"&updcosteValor="+document.costes.updcosteValor.value;
			cadena=cadena+"&tipoarticulo="+document.costes.tipoarticulo.value;
			cadena=cadena+"&almacen="+document.costes.almacen.value;
			cadena=cadena+"&nombreart="+document.costes.nombreart.value;
			cadena=cadena+"&referencia="+document.costes.referencia.value;
			cadena=cadena+"&hfecha="+document.costes.hfecha.value;
			cadena=cadena+"&dfecha="+document.costes.dfecha.value;
			cadena=cadena+"&nproveedor="+document.costes.nproveedor.value;
			cadena=cadena+"&RefFij=";
			cadena=cadena+"&ordenar=" + document.costes.ordenar.value;
			cadena=cadena+"&de_que_campo_vengo=RefPro";
			marcoCostes.document.costes_datos.action="costes_datosn.asp?mode=creartab&camb=1"+cadena;
			marcoCostes.document.costes_datos.submit();

			//seleccionar('marcoCostes','costes','check');
		  }
		}
	}
}
</script>
<%mode=Request.QueryString("mode")%>
<body onload="self.status='';" bgcolor="<%=iif(mode="save_multi" or mode="save","",color_blau)%>">
<script language="javascript" type="text/javascript">
    var Ven = -1; //apunta a la ventana que se abre
    var nn = 0
    var cadEmp = ""
</script>
 <% 
 
m_coste=limpiaCadena(request.querystring("updcoste"))
m_pvp=limpiaCadena(request.querystring("updpvp"))

mode=Request.QueryString("mode")
if request.querystring("emp") >"" then
	emp = limpiaCadena(request.querystring("emp"))
	observaMM = limpiaCadena(request.querystring("obsv"))
	
	'vamos a crear la cadena con los nombres de las empresas
	set rstCadEmp = Server.CreateObject("ADODB.Recordset")
	rstCadEmp.cursorlocation=3
	
	'obtenemos la cadena de conexion
	ncliente=session("ncliente")
    dsnCliente=session("dsn_cliente")
    sesion_usuario=session("usuario")
    ''ricardo 13/10/2004 se cambiara el usuario del dsncliente por el de DSNImport
    initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")
    donde=inStr(1,DSNImport,"Initial Catalog=",1)
    donde_fin=InStr(donde,DSNImport,";",1)
    if donde_fin=0 then
        donde_fin=len(DSNImport)
    end if
    cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))

    rstCadEmp.Open "exec ObtenerCadenaEmpresas '" & "," & emp & "'", cadena_dsn_final,adUseClient, adLockReadOnly
    
    if not rstCadEmp.eof then
        nn=rstCadEmp.RecordCount
        cont=1
        while not rstCadEmp.EOF
            cadEmp=cadEmp+rstCadEmp("rsocial")
            if cont<nn then
                cadEmp=cadEmp + ","
            end if
            cont=cont+1
            rstCadEmp.MoveNext
        wend
    end if
    rstCadEmp.close
    set rstCadEmp = nothing
end if

'mmg 15/04/2008: Calculamos si nos encontramos en modo multiempresa o los cambios son sólo para la empresa actual
set rstMultEmp = Server.CreateObject("ADODB.Recordset")
cad="exec ObtenerDependenciaEmpresas '" & session("ncliente") & "','" & session("usuario") & "'"
rstMultEmp.cursorlocation=3
rstMultEmp.Open cad, session("dsn_cliente"),adUseClient, adLockReadOnly

if not rstMultEmp.eof then	
	%>
	<script language="javascript" type="text/javascript">
	    nn = '<%=rstMultEmp.RecordCount%>';
	</script>
	<%
	rstMultEmp.close
else 
    'la empresa es hija de otra
	rstMultEmp.close
end if
set rstMultEmp = nothing

'----------------------------------------------------------------------------
'Funciones
'----------------------------------------------------------------------------
'*************************************************************************************************************
'Botones de navegación para las búsquedas.
sub SpanNextPrev(lote,lotes,pos)
%>
<table width='100%' border='0' cellspacing="1" cellpadding="1">
	<tr><td class='MAS'><%
	   lote=cint(lote)
	   lotes=cint(lotes)
	    varias=false
		if lote>1 then
			%><a class='CELDAREF' href="javascript:Mas('prev',<%=EncodeForHtml(lote)%>);">
			<img src="../images/<%=ImgAnterior%>" <%=ParamImgAnterior%> alt="<%=LitAnterior%>" title="<%=LitAnterior%>"/></a><%
			varias=true
		end if
		texto=LitPagina + " " + cstr(lote)+ " "+ LitDe + " " + cstr(lotes)
		%><font class='CELDA'> <%=EncodeForHtml(texto)%> </font> <%


		if lote<lotes then
			%><a class='CELDAREF' href="javascript:Mas('next',<%=EncodeForHtml(lote)%>);">
			<img src="../images/<%=ImgSiguiente%>" <%=ParamImgSiguiente%> alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"/></a><%
			varias=true
		end if

	%></td></tr>
</table><%
end sub

'*************************************************************************************************************
sub SpanStock(updcoste,updpvp)
	%>
    <div class="overflowXauto">
    <table class="width90 lg-table-responsive bCollapse">
        <tr>
            <td class="ENCABEZADOC underOrange width5" style="text-align: center;" ><input type="checkbox" name="check1" onclick="seleccionar('marcoCostes','costes_datos','check1');"/></td>
			<td class="ENCABEZADOL underOrange width5" ><%=LitRef%></td>
			<td class="ENCABEZADOL underOrange width10" ><%=LitNombre%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitCosteMedio%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitCosteAnterior%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitCosteActual%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitRecargo%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitMargen%></td>
			<td class="ENCABEZADOC underOrange width10" ><%=LitPvp%></td>
			<td class="ENCABEZADOC underOrange width10" ><%=LitPvpIva%></td>
			<td class="ENCABEZADOC underOrange width5" ><%=LitPuc%></td>
            <td class="ENCABEZADOC underOrange width5" style="text-align: center;">&darr;&uarr;</td>
            <td class="ENCABEZADOC underOrange width5"></td>
        </tr>
		<%
	%></table>
	<iframe class='width90 iframe-input lg-table-responsive' name="marcoCostes" id='frcostes' src='costes_datosn.asp?mode=vacio' style="max-height:750px; height:300px" noresize="noresize"></iframe>
	</div>
        <table><%
		DrawFila ""
			%><td class='CELDA7' width="200">
				<span id="barras" style="display:none">
				</span>
			</td><%
		CloseFila
	%></table><%
			DrawDiv "4","",""
            DrawLabel "","",LitCosteGral%><input class='CELDAR7' name="costegral" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarCoste('1'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliCostGral%>" style="text-align:center" title="<%=LitApliCostGral%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitRecargoGral%><input class='CELDAR7' name="recgral" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarRecargo('1'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliRecGral%>" style="text-align:center" title="<%=LitApliRecGral%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitMargenGral%><input class='CELDAR7' name="margengral" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarMargen('1'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliMargGral%>" style="text-align:center" title="<%=LitApliMargGral%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitPvpGral%><input class='CELDAR7' name="precgral" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarPrec('1'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliPvpGral%>" style="text-align:center" title="<%=LitApliPvpGral%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitPvpIvaGral%><input class='CELDAR7' name="precivagral" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarPrecIva('1'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliPvpGral%>" style="text-align:center" title="<%=LitApliPvpGral%>"/></a><%CloseDiv
		    DrawDiv "4","",""
            DrawLabel "","",LitCosteIncDec%><input class='CELDAR7' name="costeIncDec" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarCoste('2'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliCostIncDec%>" title="<%=LitApliCostIncDec%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitRecargoIncDec%><input class='CELDAR7' name="recIncDec" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarRecargo('2'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliRecIncDec%>" title="<%=LitApliRecIncDec%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitMargenIncDec%><input class='CELDAR7' name="margenIncDec" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarMargen('2'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliMargIncDec%>" title="<%=LitApliMargIncDec%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitPVPIncDec%><input class='CELDAR7' name="precIncDec" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarPrec('2'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliPvpIncDec%>" title="<%=LitApliPvpIncDec%>"/></a><%CloseDiv
            DrawDiv "4","",""
            DrawLabel "","",LitPVPIvaIncDec%><input class='CELDAR7' name="precIvaIncDec" value="" size="5"/><a class='ic-accept noMTop' href="javascript:if(AplicarPrecIva('2'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitApliPvpIncDec%>" title="<%=LitApliPvpIncDec%>"/></a><%CloseDiv
            %><%
		DrawFila ""
				%><table width='100%'></table>
    <td width='350'></td><td width='100'><a href="javascript:Seleccionar2();"><img src="../images/<%=ImgAbajo%>" <%=ParamImgAbajo%> alt='' title=''/></a></td>
				<td width='100'><a href="javascript:DeSeleccionar2();"><img src="../images/<%=ImgArriba%>" <%=ParamImgArriba%> alt='' title=''/></a></td><%
		CloseFila
	%>
    <div class="overflowXauto">
	<table class="width90 lg-table-responsive bCollapse">
        <tr>
            <td class="ENCABEZADOC underOrange width5" style="text-align: center;" ><input type="checkbox" name="check2" onclick="seleccionar('marcoResult','costes_result','check2');"/></td>
			<td class="ENCABEZADOL underOrange width5" ><%=LitRef%></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitNombre%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitCosteMedio%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitCosteAnterior%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitCosteActual%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitRecargo%></td>
			<td class="ENCABEZADOR underOrange width5" ><%=LitMargen%></td>
			<td class="ENCABEZADOC underOrange width10" ><%=LitPvp%></td>
			<td class="ENCABEZADOC underOrange width10" ><%=LitPvpIva%></td>
			<td class="ENCABEZADOC underOrange width5" ><%=LitPuc%></td>
            <td class="ENCABEZADOC underOrange width5" style="text-align: center;" >&darr;&uarr;</td>
            <td class="ENCABEZADOC underOrange width5"></td>
		</tr>	
            </table>
		<iframe class='width90 iframe-input lg-table-responsive' name="marcoResult" id='frResultCostes' src='costes_resultn.asp?mode=vacio' style="max-height:750px; height:150px" noresize="noresize"></iframe>
	</div>
        <table style="width:90%"><%
		DrawFila ""
			%><td class='CELDA7' width="120">
				<span id="barras2" style="display:">
				</span>

			</td>
			<td class='CELDAR7' width="780"><a href="javascript:GuardarValores();"><img src="../images/<%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitActCostes%>" title="<%=LitActCostes%>"/></a>
			</td><%
		CloseFila %>
    </table>
<%end sub

sub SpanRevAct ()
	%>
    <div class="overflowXauto">
    <table class="width90 lg-table-responsive bCollapse" ><%
		'Drawfila ""
			%><!--<td class='CELDA7' heigth="15px">&nbsp;&nbsp;</td>--><%
		'CloseFila
			%><td class="ENCABEZADOL underOrange width10" ><input type='checkbox' name='seltodos' onclick="Seleccionar();" /></td><%
			DrawCeldaDet "'ENCABEZADOL underOrange width10'","width100","",0,LitRef
			DrawCeldaDet "'ENCABEZADOL underOrange width10'","width100","",0,LitNombre
			DrawCeldaDet "'ENCABEZADOR underOrange width10'","width100","",0,LitCoste
			DrawCeldaDet "'ENCABEZADOR underOrange width10'","width100","",0,LitRecargo
			DrawCeldaDet "'ENCABEZADOR underOrange width10'","width100","",0,LitMargen
			DrawCeldaDet "'ENCABEZADOR underOrange width10'","width100","",0,LitPvp
	%></table>
		<iframe class='width90 iframe-input lg-table-responsive' name="marcoRevActualizaciones" id='frrevactualizaciones' src='CostesnActualizaciones.asp?mode=revisar' style="max-height:750px; height:150px" noresize="noresize"></iframe>
    </div>
    <table style="width:90%"><%
		DrawFila ""
			%><td class='CELDA7' width="600">
			</td>
			<td class="CELDAR7">
			<a class='CELDAREF' href="javascript:ActPre();">
				<img src="../images/<%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitAct%>" title="<%=LitAct%>"/></a>
			</td>
			<td class="CELDAR7">
			<a class='ic-delete noMTop noMBottom' href="javascript:BorrarActPre2();">
				<img src="<%=themeIlion %><%=ImgEliminarDet%>" <%=ParamImgEliminar%> alt="<%=LitEl%>" title="<%=LitEl%>"/></a></td>
			<%
		CloseFila
	  %></table><%
end sub

sub SpanActPendiente()
	%>
    <div class="overflowXauto">
        <table class="width90 lg-table-responsive bCollapse"><%
		'Drawfila ""
			%><!--<td class='CELDA7' heigth="15px">&nbsp;&nbsp;</td>--><%
		'CloseFila
		'Drawfila color_terra
				DrawCeldaDet "'ENCABEZADOL underOrange width25'","width100","",0,LitFV
				DrawCeldaDet "'ENCABEZADOR underOrange width25'","width100","",0,LitArt
				DrawCeldaDet "'ENCABEZADOC underOrange width25'","width100","",0,LitRev
				DrawCeldaDet "'ENCABEZADOC underOrange width25'","width100","",0,LitEl
			'Closefila
	%></table>
		<iframe class='width90 iframe-input lg-table-responsive' name="marcoActualizaciones" id='fractualizaciones' src='CostesnActualizaciones.asp?mode=ver' style="max-height:750px; height:150px" noresize="noresize"></iframe>
    </div>
<%end sub

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0

  %>

<form name="costes" method="post">
	<%  PintarCabecera "costes.asp"
WaitBoxOculto LitEsperePorFavor
	Alarma "costes.asp"

'comprobar que el usuario está dado de alta en personal
 dni=d_lookup("dni","personal","login='" & session("usuario") & "' and dni like '"+Session("ncliente")+"%'",session("dsn_cliente"))

 if dni&""="" then
	waitbox LitMsgUsuarioPersonalNoExiste
 else
 
		'Leer parámetros de la página
		if request.querystring("lote") >"" then
		   lote = limpiaCadena(request.querystring("lote"))
		elseif request.form("lote")>"" then
		   lote = limpiaCadena(request.form("lote"))
		else
		   lote = 1
		end if
		if request.querystring("nproveedor") >"" then
		   nproveedor = limpiaCadena(request.querystring("nproveedor"))
		else
		   nproveedor = limpiaCadena(request.form("nproveedor"))
		end if
		if request.querystring("Hfecha")>"" then
			TmpHfecha=limpiaCadena(request.querystring("Hfecha"))
		else
			TmpHfecha=limpiaCadena(request.Form("Hfecha"))
		end if
		if request.querystring("Dfecha")>"" then
			TmpDfecha=limpiaCadena(request.querystring("Dfecha"))
		else
			TmpDfecha=limpiaCadena(request.Form("Dfecha"))
		end if
		if request.querystring("referencia") >"" then
		   referencia = limpiaCadena(request.querystring("referencia"))
		else
		   referencia = limpiaCadena(request.form("referencia"))
		end if
		if request.querystring("nombreart") >"" then
		   nombreart = limpiaCadena(request.querystring("nombreart"))
		else
		   nombreart = limpiaCadena(request.form("nombreart"))
		end if
		if request.querystring("almacen") >"" then
		   almacen = limpiaCadena(request.querystring("almacen"))
		else
		   almacen = limpiaCadena(request.form("almacen"))
		end if
		'cag
		if request.form("tipoarticulo")>"" then
			tipoarticulo = limpiaCadena(request.form("tipoarticulo"))
		else
			tipoarticulo = limpiaCadena(request.querystring("tipoarticulo"))
		end if
		if request.form("updcoste")>"" then
			updcoste = limpiaCadena(request.form("updcoste"))
		else
			updcoste = limpiaCadena(request.querystring("updcoste"))
		end if

		if request.form("updpvp")>"" then
			updpvp = limpiaCadena(request.form("updpvp"))
		else
			updpvp = limpiaCadena(request.querystring("updpvp"))
		end if

	    actCoste=d_lookup("updatecoste","configuracion","nempresa='"&session("ncliente")&"'",session("dsn_cliente"))
	    actPvp=d_lookup("updatepvp","configuracion","nempresa='"&session("ncliente")&"'",session("dsn_cliente"))

		'fin cag
		if request.querystring("familia") >"" then
		   familia = limpiaCadena(request.querystring("familia"))
		else
		   familia = limpiaCadena(request.form("familia"))
		end if
		if request.querystring("familia_padre") >"" then
		   familia_padre = limpiaCadena(request.querystring("familia_padre"))
		else
		   familia_padre = limpiaCadena(request.form("familia_padre"))
		end if
		if request.querystring("categoria") >"" then
		   categoria = limpiaCadena(request.querystring("categoria"))
		else
		   categoria = limpiaCadena(request.form("categoria"))
		end if
		if request.querystring("ordenar") >"" then
		   ordenar = limpiaCadena(request.querystring("ordenar"))
		else
		   ordenar = limpiaCadena(request.form("ordenar"))
		end if
		if request.querystring("ndocumento") >"" then
		   ndocumento = limpiaCadena(request.querystring("ndocumento"))
		else
		   ndocumento = limpiaCadena(request.form("ndocumento"))
		end if
		nombre=limpiaCadena(request.form("nombre"))

		if request.querystring("obs") >"" then
		   obs = limpiaCadena(request.querystring("obs"))
		end if
		if request.querystring("fv") >"" then
			fv = limpiaCadena(request.querystring("fv"))
		end if%>
	<table width='100%'>
   	<tr>
		<% if mode="ver" then %>
	  		<td class="ENCABEZADOL" align="right">
				<%pagina="costes_imp.asp?referencia=" & referencia & "&nombre=" & nombre & _
					"&familia=" & familia & "&almacen=" & almacen & "&ordenar=" + ordenar%>
          		<a class='CELDAREFB' href="javascript:AbrirVentana('<%=EncodeForHtml(pagina)%>','I',<%=AltoVentana%>,<%=AnchoVentana%>)" onmouseover="self.status='<%=LitFormatoImpre%>'; return true;" onmouseout="self.status=''; return true;"><font class=CELDA>Formato de Impresión</font></a>
	  		</td>
	  <% else %>
	  		<td>&nbsp;</td>
	  <% end if %>
   	</tr>
    </table>
<% if mode="ver" then %>
	<hr/><%
end if

  set rstAux = Server.CreateObject("ADODB.Recordset")
  set rst = Server.CreateObject("ADODB.Recordset")
    if mode="save_multi" then
            'obtenemos la cadena de conexión
            ncliente=session("ncliente")
            dsnCliente=session("dsn_cliente")
            sesion_usuario=session("usuario")
            ''ricardo 13/10/2004 se cambiara el usuario del dsncliente por el de DSNImport
            initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")
            donde=inStr(1,DSNImport,"Initial Catalog=",1)
            donde_fin=InStr(donde,DSNImport,";",1)
            if donde_fin=0 then
                donde_fin=len(DSNImport)
            end if
            cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
            
            ''obtenemos los decimales de la divisa de la empresa
            ndecimales=d_lookup("ndecimales","divisas","codigo like '" & ncliente & "%' and moneda_base<>0",session("dsn_cliente"))
    
            'obtenemos la fecha de vencimiento
            fv=limpiaCadena(request.querystring("fv"))
			
            rst.cursorlocation=3
            
		    cc="EXEC GuardarCostesMultiEmpresa @fecha='"&date()&"', @Nomtabla='"&session("usuario")&"', @ListaCliente='"&","&emp&"', @observaciones='"&observaMM&"', @fvenc='"&fv&"' , @ebs=0, @m_coste='"&m_coste&"', @m_pvp='"&m_pvp&"'"

		    rst.open cc,cadena_dsn_final,adUseClient,adLockReadOnly
		    if fv&"">"" then
			%><script language="javascript" type="text/javascript">
			      document.costes.action = "costesn.asp?mode=param&submode=fins";
			      document.costes.submit();
			</script><%
		    else%>
					<table border='0' cellspacing="1" cellpadding="1">
					    <tr>
					        <%set rstResp = Server.CreateObject("ADODB.Recordset")
					        rstResp.cursorlocation=3
					        rstResp.open "select nombre from personal with(nolock) where dni = '"+session("ncliente")+session("usuario")+"'",cadena_dsn_final,adUseClient,adLockReadOnly
					        
					        DrawFila color_blau
							    DrawCelda "ENCABEZADOL","","",0,LitResponsable & ":"
							    DrawCelda "CELDALEFT","","",0,EncodeForHtml(rstResp("nombre"))
							CloseFila
							DrawFila color_blau
							    DrawCelda "ENCABEZADOL","","",0,LitObservaciones & ":"
							    DrawCelda "CELDALEFT","","",0,EncodeForHtml(observaMM)
							CloseFila
							DrawFila color_blau
							    DrawCelda "ENCABEZADOL","","",0,LitEmpresas & ":"
							    DrawCelda "CELDALEFT","","",0,EncodeForHtml(cadEmp)
							CloseFila
							rstResp.Close()
                            set rstResp = nothing%>
					   </tr>
					</table>
					<br/>
					<table width='100%' border='0' cellspacing="1" cellpadding="1">
					    <%DrawFila color_fondo
							DrawCelda "ENCABEZADOL","","",0,LitItem
							DrawCelda "ENCABEZADOL","","",0,LitRSocial
							DrawCelda "ENCABEZADOL","","",0,LitRef
							DrawCelda "ENCABEZADOL","","",0,LitNombre
							DrawCelda "ENCABEZADOR","","",0,LitCosteAnterior
							DrawCelda "ENCABEZADOR","","",0,LitRecargoAnt
							DrawCelda "ENCABEZADOR","","",0,LitMargenAnt
							DrawCelda "ENCABEZADOR","","",0,LitPrecioCostAnt
							DrawCelda "ENCABEZADOR","","",0,LitPvpAnt
							DrawCelda "ENCABEZADOR","","",0,LitCosteNuevo
							DrawCelda "ENCABEZADOR","","",0,LitRecargo
							DrawCelda "ENCABEZADOR","","",0,LitMargen
							DrawCelda "ENCABEZADOR","","",0,LitPrecioCost
							DrawCelda "ENCABEZADOR","","",0,LitPvp
						CloseFila
					while not rst.EOF
						DrawFila color_blau
							DrawCelda "CELDALEFT","","",0,  EncodeForHtml(rst("item"))
							DrawCelda "CELDALEFT","","",0,  EncodeForHtml(rst("r_social"))
							DrawCelda "CELDALEFT","","",0,  EncodeForHtml(trimCodEmpresa(rst("ref")))
							DrawCelda "CELDALEFT","","",0,  EncodeForHtml(rst("nombreArt"))''d_lookup("nombre", "articulos", "referencia='"& rst("ref") &"'", session("dsn_cliente"))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("importeant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("recargoant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("margenant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("precioant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("pvpant"),ndecimales,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("importe"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("recargo"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("margen"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("precio"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0, EncodeForHtml(formatnumber(rst("pvp"),ndecimales,-1,0,-1))
						CloseFila
						rst.MoveNext
					wend
			end if
    end if

 	if mode="save" then
 	    'obtenemos la cadena de conexión
        ncliente=session("ncliente")
        dsnCliente=session("dsn_cliente")
        sesion_usuario=session("usuario")
        ''ricardo 13/10/2004 se cambiara el usuario del dsncliente por el de DSNImport
        initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")
        donde=inStr(1,DSNImport,"Initial Catalog=",1)
        donde_fin=InStr(donde,DSNImport,";",1)
        if donde_fin=0 then
            donde_fin=len(DSNImport)
        end if
        cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
        
        ''obtenemos los decimales de la divisa de la empresa
        ndecimales=d_lookup("ndecimales","divisas","codigo like '" & ncliente & "%' and moneda_base<>0",session("dsn_cliente"))
        
	    cc="EXEC GuardarCostes @fecha='"&date()&"', @Nomtabla='"&session("usuario")&"', @Ncliente='"&session("ncliente")&"', @observaciones='"&obs&"', @fvenc='"&fv&"' , @ebs=0"
	    rst.cursorlocation=3
	    rst.open cc,cadena_dsn_final,adUseClient,adLockReadOnly

		if fv&"">"" then
		    %><script language="javascript" type="text/javascript">
		          document.costes.action = "costesn.asp?mode=param&submode=fins";
		          document.costes.submit();
			</script><%
		else
		    if not rst.eof then
			    auditar_ins_bor session("usuario"),rst("ncambio"),"","","","","cambiarPrecios"%>
					<table width='100%' border='<%=borde%>' cellspacing="0" cellpadding="0">
						<tr bgcolor='<%=color_blau%>'><td>&nbsp;</td><td>&nbsp;</td></tr>
				            <%DrawFila color_blau%>
								<td class="CABECERA" width="50%" align="left"><%=LitCB%> :<%=EncodeForHtml(trimCodEmpresa(rst("ncambio")))%></td>
								<td class="CELDA" width="50%" align="center"><%=LitFecha%> : <%=EncodeForHtml(rst("fecha"))%></td>
						</tr>
						<tr bgcolor='<%=color_blau%>'>
							<td class="CELDA" colspan="2"><%= LitResponsable%> : <%= EncodeForHtml(d_lookup("nombre", "personal", "dni='"& rst("responsable") &"'", session("dsn_cliente"))) %></td>
						</tr>
						<tr bgcolor='<%=color_blau%>'>
							<td class="CELDA" colspan="2"><%= LitObservaciones%> : <%= EncodeForHtml(rst("observaciones")) %></td>
						</tr>
					</table>
					<br/>
					<table width='100%' border='0' cellspacing="1" cellpadding="1">
					    <%DrawFila color_fondo
							DrawCelda "ENCABEZADOL","","",0,LitItem
							DrawCelda "ENCABEZADOL","","",0,LitRef
							DrawCelda "ENCABEZADOL","","",0,LitNombre
							DrawCelda "ENCABEZADOR","","",0,LitCosteAnterior
							DrawCelda "ENCABEZADOR","","",0,LitRecargoAnt
							DrawCelda "ENCABEZADOR","","",0,LitMargenAnt
							DrawCelda "ENCABEZADOR","","",0,LitPrecioCostAnt
							DrawCelda "ENCABEZADOR","","",0,LitPvpAnt
							DrawCelda "ENCABEZADOR","","",0,LitCosteNuevo
							DrawCelda "ENCABEZADOR","","",0,LitRecargo
							DrawCelda "ENCABEZADOR","","",0,LitMargen
							DrawCelda "ENCABEZADOR","","",0,LitPrecioCost
							DrawCelda "ENCABEZADOR","","",0,LitPvp
						CloseFila
					while not rst.EOF
						DrawFila color_blau
							DrawCelda "CELDALEFT","","",0, EncodeForHtml(rst("item"))
							DrawCelda "CELDALEFT","","",0, EncodeForHtml(trimCodEmpresa(rst("referencia")))
							DrawCelda "CELDALEFT","","",0, EncodeForHtml(rst("nombreArt"))''d_lookup("nombre", "articulos", "referencia='"& rst("referencia") &"'", session("dsn_cliente"))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("importeant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("recargoant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("margenant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("precioant"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("pvpant"),ndecimales,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("importe"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("recargo"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("margen"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("precio"),DEC_PREC,-1,0,-1))
							DrawCelda "CELDARIGHT","","",0,EncodeForHtml(formatnumber(rst("pvp"),ndecimales,-1,0,-1))
						CloseFila
						rst.MoveNext
					wend
		    end if
		end if
		rst.close
   end if  'fin mode save %>

            <%
              '*********************************************************************************************
              'Se muestran parametros de seleccion
              '*********************************************************************************************
             if mode="param" or mode= "param1" then%>
                <div id="CollapseSection">
                  <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['Parametros', 'Detalle']); hideNoCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" title="" /></a> 
                  <a id="collapse_all_button"    href="javascript:animatedcollapse.hide(['Parametros', 'Detalle']); hideCollapse();" ><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" title="" /></a>
                </div>

              <div   class="Section" id="S_Parametros">
                <a href="#" rel="toggle[Parametros]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                    <div class="SectionHeader">
                        <%=LitAcordeon1 %>
                        <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
                    </div>   
                </a>
             <div class="SectionPanel" style="display: " id="Parametros">      
             <%
			    'DrawCelda2 "CELDAL7", "left", false, LitDesdeFechaCompra & ": "
			    'DrawInputCelda "CELDAL7","","",12,0,"","dfecha",TmpDfecha
                EligeCelda "input","add","left","","",0,LitDesdeFechaCompra,"dfecha",12,EncodeForHtml(TmpDfecha)
                DrawCalendar "dfecha"
			    'DrawCelda2 "CELDAL7", "left", false, LitHastaFechaCompra & ": "
			    'DrawInputCelda "CELDAL7","","",12,0,"","hfecha",TmpHfecha
                EligeCelda "input","add","left","","",0,LitHastaFechaCompra,"hfecha",12,EncodeForHtml(TmpHfecha)
                DrawCalendar "hfecha"
			    if nproveedor >"" then nproveedor=Completar(nproveedor,5,"0")
			    'DrawCelda2 "CELDAL7", "left", false, LitProveedor & " : "
                DrawDiv "1","",""
                DrawLabel "","",LitProveedor
                %><input class="CELDA" type="hidden" name="nproveedor" value="<%=EncodeForHtml(nproveedor)%>"/>
                 <input class="CELDA" type="hidden" name="nombre" value=""/>
                 <iframe class="width60 iframe-menu"  id='frProveedor' src='../compras/docproveedor_responsiveDet.asp?viene=costes&anterior=nombreampliado&nproveedor=<%=EncodeForHtml(p_nproveedor)%>' frameborder="no" scrolling="no" noresize="noresize"></iframe><%
                CloseDiv
                'DrawCelda2 "CELDAL7", "left", false, LitDocumentoEntrada & ": "
			    'DrawInputCelda "CELDAL7","","",17,0,"","ndocumento",""
		        EligeCelda "input","add","left","","",0,LitDocumentoEntrada ,"ndocumento",17,""
        	    'DrawCelda2 "CELDAL7", "left", false, LitConref + ": "
   	      	    'DrawInputCelda "CELDAL7","","",25,0,"","referencia",referencia
                EligeCelda "input","add","left","","",0,LitConref,"referencia",25,EncodeForHtml(referencia)
          	    'DrawCelda2 "CELDAL7", "left", false, LitConNombre + ": "
	      	    'DrawInputCelda "CELDAL7","","",25,0,"","nombreart",nombreart
                EligeCelda "input","add","left","","",0,LitConNombre,"nombreart",25,EncodeForHtml(nombreart)
			    'DrawCelda2 "CELDAL7", "left", false, LitAlmacen + ": "
		   	    rstAux.open " select codigo, descripcion from almacenes with(nolock) where codigo like '" & session("ncliente") & "%'", session("dsn_cliente"),adOpenKeyset,adLockOptimistic
	       	    'DrawSelectCelda "CELDAL7","170","",0,"","almacen",rstAux,almacen,"codigo","descripcion","",""
                DrawSelectCelda "CELDA","170","",0,LitAlmacen,"almacen",rstAux,almacen,"codigo","descripcion","",""
		   	    rstAux.close
			    'fin cag
                'DrawCelda2 "CELDAL7 valign='top'", "left", false, LitTipoArticulo & ": "
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
                DrawDiv "1","",""
                DrawLabel "","",LitTipoArticulo%><select multiple="multiple" size="6" class="width60" name="tipoarticulo">
			 	    <%while not rstArtType.eof%>
		   		        <option value="<%=EncodeForHtml(rstArtType("codigo"))%>"><%=EncodeForHtml(rstArtType("descripcion"))%></option>
				        <%rstArtType.movenext%>
				    <%wend
                    if tipoarticulo="" then %>
			            <option selected="selected" value=""> </option>
				    <%else%>
				        <option selected="selected" value="<%=EncodeForHtml(tipoarticulo)%>"> <%=EncodeForHtml(trimCodEmpresa(tipoarticulo))%></option>
			            <option value=""> </option>
				    <%end if%></select><%
                CloseDiv
                rstArtType.close
                conn.close
                set rstArtType = nothing
                set command = nothing
                set conn = nothing
                    DrawFila color_blau
			            dim ConfigDespleg (3,13)

			            i=0
			            ConfigDespleg(i,0)="categoria"
			            ConfigDespleg(i,1)=""
			            ConfigDespleg(i,2)="6"
			            ConfigDespleg(i,3)="select codigo, nombre from categorias where codigo like '" & session("ncliente") & "%' order by nombre"
			            ConfigDespleg(i,4)=1
			            ConfigDespleg(i,5)="width60"
			            ConfigDespleg(i,6)="MULTIPLE"
			            ConfigDespleg(i,7)="codigo"
			            ConfigDespleg(i,8)="nombre"
			            ConfigDespleg(i,9)=LitCategoria
			            ConfigDespleg(i,10)=categoria
			            ConfigDespleg(i,11)=""
			            ConfigDespleg(i,12)=""

			            i=1
			            ConfigDespleg(i,0)="familia_padre"
			            ConfigDespleg(i,1)=""
			            ConfigDespleg(i,2)="6"
			            ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre where codigo like '" & session("ncliente") & "%' order by nombre"
			            ConfigDespleg(i,4)=1
			            ConfigDespleg(i,5)="width60"
			            ConfigDespleg(i,6)="MULTIPLE"
			            ConfigDespleg(i,7)="codigo"
			            ConfigDespleg(i,8)="nombre"
			            ConfigDespleg(i,9)=LitFamilia
			            ConfigDespleg(i,10)=familia_padre
			            ConfigDespleg(i,11)=""
			            ConfigDespleg(i,12)=""

			            i=2
			            ConfigDespleg(i,0)="familia"
			            ConfigDespleg(i,1)=""
			            ConfigDespleg(i,2)="6"
			            ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias where codigo like '" & session("ncliente") & "%' order by nombre"
			            ConfigDespleg(i,4)=1
			            ConfigDespleg(i,5)="width60"
			            ConfigDespleg(i,6)="MULTIPLE"
			            ConfigDespleg(i,7)="codigo"
			            ConfigDespleg(i,8)="nombre"
			            ConfigDespleg(i,9)=LitSubFamilia
			            ConfigDespleg(i,10)=familia
			            ConfigDespleg(i,11)=""
			            ConfigDespleg(i,12)=""

			            DibujaDesplegables ConfigDespleg,session("dsn_cliente")

	                CloseFila
	            %><%
	                   'DrawCelda2 "CELDAL7 ", "left", false, LitOrdenar + ": "
                    DrawDiv "1","",""
                    DrawLabel "","",LitOrdenar
                    %><select class='width60' name="ordenar">
		              <option value="REFERENCIA_ASC"><%=LitReferenciaAsc%></option>
		              <option selected="selected" value="REFERENCIA_DESC"><%=LitReferenciaDesc%></option>
				      <%if ordenar="NOMBRE" then%>
		   			      <option selected="selected" value="NOMBRE"><%=LitNombreMay%></option>
				      <%else%>
		   			      <option value="NOMBRE"><%=LitNombreMay%></option>
				      <%end if
                    %></select><%
                    CloseDiv 
			        'DrawCelda2 "CELDAL7", "right", false, LitRegistros + ": "
                    DrawDiv "1","",""
                    DrawLabel "","",LitRegistros
                    %><input class="CELDAL7" size="5" type="Text" name="registrospag" value="100"/><%
                    DrawLabel "","","(Max. 200)"
                    CloseDiv
                    DrawDiv "1","",""
                    DrawLabel "","",LitRefPro1
                    %><input class="CELDAL7" type="text" name="RefPro" value="" size="20" runat="javascript:comprobar_enter();"/>
                    <label><%=LitRefPro2%></label><%
                    CloseDiv
		            %>
			        <%'if actCoste=True then %>
			        <%if updcoste1<>"" then
			            vCoste=updcoste1
			        else
			            if actCoste=True then
				            vCoste=1
				        else
				            vCoste=0
				        end if
			        end if

        			 if actPvp=True then
		        	    vPvp=1
			        else
			            vPvp=0
 		            end if

			        if vCoste=1 then
                        DrawDiv "1","",""
                        DrawLabel "","",LitActCoste
                        %><input class="CELDAL7" type="checkbox" name="updcoste" checked="checked" onclick="javascript:bloquear(1)"/>
		 	            <input type="hidden" name="updcoste1" value="1"/><%
                        CloseDiv
                    else
                        DrawDiv "1","",""
                        DrawLabel "","",LitActCoste
                        %><input class="CELDAL7" type="checkbox" name="updcoste" onclick="javascript:bloquear(1)"/>
		 	            <input type="hidden" name="updcoste1" value="0"/><%
                        CloseDiv
                    end if
                    %><input type="hidden" name="updcosteValor" value="<%=EncodeForHtml(vCoste)%>"/><%
                    DrawDiv "1","",""
                    DrawLabel "","",LitActPVP
			        if actPvp=True then 
                        %><input class="CELDAL7" type="checkbox" name="updpvp" checked="checked" onclick="javascript:bloquear(3);"/><%
                    else
                        %><input class="CELDAL7" type="checkbox" name="updpvp" onclick="javascript:bloquear(3);"/><%
                    end if
                    CloseDiv
                    %><input type="hidden" name="updpvpValor" value="<%=EncodeForHtml(vPvp)%>"/><%
                    'DrawCelda2 "CELDAL7 style='width:145px'", "left", false, LitFV &": "
		            'DrawInputCelda "CELDAL7 style='width:100px'","","",15,0,"","fv",""
                    EligeCelda "input","add","left","","",0,LitFV,"fv",15,""
                    DrawCalendar "fv"
                    DrawDiv "1","",""
                    DrawLabel "","",LitCostesRefPro1
                    %><input class="CELDA7" type="text" name="RefFij" value="" size="20" runat="javascript:comprobar_enter();" />
                    <label><%=LITREFPRO2%></label><%
                    CloseDiv
                    DrawDiv "1","",""
                    DrawLabel "","",LitCargarArticulos
                    %><a class='ic-accept noMTop' href="javascript:if (Mostrar(''));">
                        <img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>" title="<%=LitCargarArticulos%>"/>
                      </a><%
                    CloseDiv

		            ' BarraNavegacion mode

		            submode=limpiacadena(request.querystring("submode"))

		               '*********************************************************************************************
		               ' Se muestran los datos de la consulta
		               '*********************************************************************************************
			            %>

        </div>   
    </div>       

    <div class="Section" id="S_Detalle">
        <a href="#" rel="toggle[Detalle]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader">
                <%=LitAcordeon2 %>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
            </div>   
        </a>
        <div class="SectionPanel" style="display: " id="Detalle">      
        
            <div id="tabs" style="display:none">
                <ul>
                    <li><a href="#tabs-1"><%=LitControlCostes%></a></li>
                    <li><a href="#tabs-2"><%=LitActPendientes%></a></li>               
                    <li><a href="#tabs-3"><%=LitRevisarAct%></a></li>               
                </ul>                
                <div id="tabs-1" >
                    <span id="Modificar"><%
				        SpanStock updcoste,updpvp%>
			        </span>
                </div>
                    
                <div id="tabs-2" >
			        <span id="ACTPENDIENTE"><%
				        SpanActPendiente
			        %></span>
                </div>

                <div id="tabs-3" >
			        <span id="REVACT"><%
				        SpanRevAct
			        %></span>
                </div>
            </div>
        </div>
            <!--<a href="#" rel="toggle[AddDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>"><img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> /></a>-->
    </div>       

                   
            <%if submode="fins" then%>
                <script language="javascript" type="text/javascript">
				      document.getElementById("ACTPENDIENTE").style.display = "";
				      document.getElementById("Modificar").style.display = "";
				      document.getElementById("REVACT").style.display = "";

				      //document.getElementById("img2").src = "../Images/<%=ImgCarpetaAbierta%>";
				      //document.getElementById("img1").src = "../Images/<%=ImgCarpetaCerrada%>";
				      //document.getElementById("img3").src = "../Images/<%=ImgCarpetaCerrada%>";
				      setTabsSelected(1);
				      //marcoActualizaciones.document.CostesnActualizaciones.action="CostesnActualizaciones.asp?mode=ver"
				      //marcoActualizaciones.document.CostesnActualizaciones.submit();
				</script>
            <%end if
         end if%>
   <input type="hidden" name="control" value="0"/>
   <%end if 'fin de comprobar el dni del usuario
   %>
</form>
<%end if
set rstAux = nothing
set rst = nothing%>
</body>
</html>