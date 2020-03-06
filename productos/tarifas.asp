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
<%'' JCI 18/06/2003 : MIGRACION A MONOBASE
'RGU 13/10/2006: Añadir campo pvp+iva en el span de precios
%>
<%response.buffer=true%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../XSSProtection.inc" -->
<!--#include file="../calculos.inc" -->

<%
    ' si en ncompany no se ha podido recuperar, se direcciona
	ncompany=session("ncliente") & ""
	if ncompany="" then
			response.write("<script language='JavaScript'>")
			response.write("window.parent.document.location='/"&carpetaproduccion&"/desactiva.asp?mode=10';")
			response.write("</script>")
			response.end

	end if
%>
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->

<!--#include file="tarifas.inc" -->

<!--#include file="../tablasResponsive.inc" -->

<!--#include file="../catFamSubResponsive.inc" -->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/calendar.inc" -->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../js/tabs.js.inc" -->

<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/tabs.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<!--#include file="../styles/DetailTable.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>

<script language="javascript" type="text/javascript">

    function ReloadTab(tab) {
        switch (tab) {
            case 3:
                document.getElementById('fractualizaciones').contentDocument.location.reload(true);
                break;
            case 4:
                document.getElementById('frrevactualizaciones').contentDocument.location.reload(true);
                break;
            default:
                break;
        }
    }

    animatedcollapse.addDiv('CABECERA', 'fade=1')
    animatedcollapse.addDiv('PRECIOS', 'fade=1')

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()

    var tecla_pulsada;
    /*
    if (window.document.addEventListener) {
        window.document.addEventListener("keydown", callkeydownhandler, false);
    } else {
        window.document.attachEvent("onkeydown", callkeydownhandler);
    }
    
    function callkeydownhandler(evnt) {
        ev = (evnt) ? evnt : event;
        tecla_pulsada=ev.keyCode;
    }
    */
function IrARangos(){
    document.location='rangos.asp?mode=add';
    parent.botones.document.location='rangos_bt.asp?mode=add';
}
function IrATemporadas(){
    document.location='temporadas.asp?mode=add';
    parent.botones.document.location='temporadas_bt.asp?mode=add';
}

    function MasTAct(sentido,lote) {
        if (divObj = marcoRevActualizaciones.document.getElementById("waitBoxOculto"))
            marcoRevActualizaciones.document.getElementById("waitBoxOculto").style.visibility = "visible";
        marcoRevActualizaciones.document.TarifasRevActualizaciones.action="TarifasActualizaciones.asp?mode=revisar&sentido=" + sentido + "&lote=" + lote;
        marcoRevActualizaciones.document.TarifasRevActualizaciones.submit();
    }

    function CampoRefPulsado(mode,marco,formulario,queordenar,comoordenar){
        if (tecla_pulsada==13){
            continuar=0;
            if (mode=="ALTA"){
                if (document.tarifas.RefPro.value!="") continuar=1;
            }
            if (mode=="BAJA"){
                if (document.tarifas.bRefPro.value!="") continuar=1;
            }
            if (continuar==1) Insertar(mode,'1','insertar',queordenar,comoordenar);
        }
    }

    function OrdenarDatos(mode,marco,formulario,campo){
        campo=campo.toUpperCase();
        eval("queordenar=" + marco + ".document." + formulario + ".queordenar.value.toUpperCase()");
        eval("comoordenar=" + marco + ".document." + formulario + ".comoordenar.value.toUpperCase()");
        if (campo!=queordenar || comoordenar==""){
            queordenar=campo;
            comoordenar="ASC";
        }
        else{
            if (campo==queordenar && comoordenar=="ASC") comoordenar="DESC";
            else comoordenar="ASC";
        }
        queimagen1="";
        queimagen2="";
        queimagen3="";
        comoimagen="";
        if (comoordenar=="ASC") comoimagen="&darr;";
        if (comoordenar=="DESC") comoimagen="&uarr;";
        if(queordenar=="A.REFERENCIA"){
            queimagen1=comoimagen;
            queimagen2="&harr;";
            queimagen3="&harr;";
        }
        if(queordenar=="A.NOMBRE"){
            queimagen2=comoimagen;
            queimagen1="&harr;";
            queimagen3="&harr;";
        }
        if(queordenar=="F.NOMBRE"){
            queimagen3=comoimagen;
            queimagen2="D";
            queimagen1="D";
        }
        if (mode=="ALTA"){
            document.getElementById("OD1A").innerHTML=queimagen1;
            document.getElementById("OD2A").innerHTML=queimagen2;
            document.getElementById("OD3A").innerHTML=queimagen3;
        }
        else{
            document.getElementById("OD1B").innerHTML=queimagen1;
            document.getElementById("OD2B").innerHTML=queimagen2;
            document.getElementById("OD3B").innerHTML=queimagen3;
        }
        Insertar(mode,'1','first',queordenar,comoordenar);
    }

    function tier1Menu(objMenu,objImage) {
        if (objMenu.style.display == "none") {
            objMenu.style.display = "";
            objImage.src = "<%=themeIlion%><%=ImgCarpetaAbierta%>";
            switch (objMenu.id) {
                case "Cabecera":
                    Anadir.style.display = "none";
                    BorrModif.style.display = "none";
                    ActPend.style.display = "none";
                    ModActPend.style.display = "none";
                    document.getElementById("img2").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img5").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById('modo').innerHTML='  <%=LitModCabecera%>  ';
                    break;
                case "Anadir":
                    Cabecera.style.display = "none";
                    BorrModif.style.display = "none";
                    ActPend.style.display = "none";
                    ModActPend.style.display = "none";
                    document.getElementById("img1").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img5").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById('modo').innerHTML='  <%=LitModAnadir%>  ';
                    break;
                case "BorrModif":
                    Cabecera.style.display = "none";
                    Anadir.style.display = "none";
                    ActPend.style.display = "none";
                    ModActPend.style.display = "none";
                    document.getElementById("img1").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img2").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img5").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById('modo').innerHTML='  <%=LitModBorrModif%>  ';
                    break;
                case "ActPend":
                    Cabecera.style.display = "none";
                    Anadir.style.display = "none";
                    BorrModif.style.display = "none";
                    ModActPend.style.display = "none";
                    document.getElementById("img1").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img2").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img5").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById('modo').innerHTML='  <%=LitModoActPend%>  ';
                    marcoActualizaciones.document.TarifasActualizaciones.action="TarifasActualizaciones.asp?mode=ver";
                    marcoActualizaciones.document.TarifasActualizaciones.submit();
                    break;
                case "ModActPend":
                    Cabecera.style.display = "none";
                    Anadir.style.display = "none";
                    ActPend.style.display = "none";
                    BorrModif.style.display = "none";
                    document.getElementById("img1").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img2").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img3").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById("img4").src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
                    document.getElementById('modo').innerHTML='  <%=LitModModActPend%>  ';
                    break;
            }
        }
        else{
            objMenu.style.display = "none";
            objImage.src = "<%=themeIlion%><%=ImgCarpetaCerrada%>";
        }
    }

    function Insertar(mode,pag,sentido,queordenar,comoordenar)
    {
        switch (mode)
        {
            case "ALTA":
                mod="save";
                if(sentido=="first"){
                    sentido="&submode=first";
                    mod="first";

                    if (document.tarifas.rango.value != "" || document.tarifas.temporada.value !="")
                        document.tarifas.condbase.value="1";
                    else document.tarifas.condbase.value="0";
                }
			
                if(sentido=="insertar")
                {
                    sentido="&submode=insertar";
                    mod="insertar";
                }
                pagina="ArticulosDeTarifa.asp?mode="+mod+"&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.tarifas.htarifa.value + "&ref=" + document.tarifas.refcontiene.value +"&familia=" + document.tarifas.familia.value + "&categoria=" + document.tarifas.categoria.value + "&familia_padre=" + document.tarifas.familia_padre.value +"&tipoarticulo=" + document.tarifas.tipoarticulo.value + "&desc=" + document.tarifas.descontiene.value + "&nproveedor=" + document.tarifas.proveedor.value + "&RefPro=" + document.tarifas.RefPro.value + "&queordenar=" + queordenar + "&comoordenar=" + comoordenar + "&apliTodo="  + document.tarifas.apliTot.value + "&temporada=" + document.tarifas.temporada.value + "&rango=" + document.tarifas.rango.value;

                if (mod=="first")
                {
                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility="visible";
                    document.getElementById("frArticulosAdd").src=pagina
                }
                else
                {
                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility="visible";
                    marcoArticulosAdd.document.ArticulosDeTarifa.action=pagina
                    marcoArticulosAdd.document.ArticulosDeTarifa.submit();
                }
                document.tarifas.check.checked=true;
                break;
            case "BAJA":
                mod="save";
                if(sentido=="first")
                {
                    sentido="&submode=first";
                    mod="first";
                    if (document.tarifas.brango.value != "" || document.tarifas.btemporada.value !="")
                        document.tarifas.bcondbase.value="1";
                    else document.tarifas.bcondbase.value="0";
                }
			
                if(sentido=="insertar")
                {
                    sentido="&submode=insertar";
                    mod="insertar";
                }

                pagina="ArticulosDeTarifa2.asp?mode="+mod+"&npagina=" + pag + "&pagina=" + sentido + "&tarifa=" + document.tarifas.htarifa.value + "&ref=" + document.tarifas.brefcontiene.value +"&familia=" +  document.tarifas.familia1.value + "&categoria=" + document.tarifas.categoria1.value + "&familia_padre=" + document.tarifas.familia_padre1.value +"&tipoarticulo=" + document.tarifas.btipoarticulo.value + "&desc=" + document.tarifas.bdescontiene.value + "&nproveedor=" + document.tarifas.bproveedor.value + "&RefPro=" + document.tarifas.bRefPro.value + "&queordenar=" + queordenar + "&comoordenar=" + comoordenar + "&temporada=" + document.tarifas.btemporada.value + "&rango=" + document.tarifas.brango.value;
                if (mod=="first")
                {
                    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
                    document.getElementById("frArticulosBorrar").src=pagina
                }
                else
                {
                    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
                    marcoArticulosBorrar.document.ArticulosDeTarifa2.action=pagina
                    marcoArticulosBorrar.document.ArticulosDeTarifa2.submit();
                }

                document.tarifas.checkb.checked=true;
                break;
        }
    }

    function Editar(p_codigo, p_npagina, p_campo, p_criterio, p_texto)
    {
        document.location = "tarifas.asp?mode=edit&p_codigo=" + p_codigo + "&npagina=" + p_npagina + "&campo=" + p_campo + "&texto=" + p_texto + "&criterio=" + p_criterio;
        parent.botones.document.location = "tarifas_bt.asp?mode=edit";
    }

    //ricardo 29-10-2007 se añade el incremento del precio mas iva
    function AplicarPvpIVA()
    {
        var msg;
        msg="<%=LITMSGNOAPLITODOS%>";
        document.tarifas.apliTot.value="off";     
        if (document.tarifas.chkAplicar.checked){
            document.tarifas.apliTot.value="on"; 
            msg="<%=LITMSGAPLITODOS%>";
        }
	
        if (!isNaN(document.tarifas.precioIva.value.replace(",",".")) && document.tarifas.precioIva.value!="" && parseFloat(document.tarifas.precioIva.value)>=0)
        {
            if (window.confirm(msg)==true)
            {
                elementos=marcoArticulosAdd.document.ArticulosDeTarifa.length;
                ref=1;
                for (i=0;i<=elementos-1;i++) {
                    switch(marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(0,1)){
                        case "p":
                            precio2=parseFloat(document.tarifas.precioIva.value.replace(",","."));
                            precio2=precio2.toFixed(<%=DEC_PREC%>);
                            eval("iva=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.hiva" + ref + ".value)");
                            pvpsiniva=precio2/(1+(iva/100))
                            ndecimales=parseInt(marcoArticulosAdd.document.ArticulosDeTarifa.elements["costdiv" + ref].value.replace(",","."));
                            pvpsiniva=pvpsiniva.toFixed(ndecimales);
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.p" + ref + ".value=pvpsiniva");
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.vppiva" + ref + ".value=document.tarifas.precioIva.value");
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.ff" + ref + ".value=marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value");
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.d" + ref + ".value=''");
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.cc" + ref + ".value=''");
                            marcoArticulosAdd.document.ArticulosDeTarifa.elements["check"+ref].checked=true;

                            //RGU 17/1/2007
                            eval("marg=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.op" + ref + ".value.replace(',','.'))");
                            eval("cost=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.coste" + ref + ".value.replace(',','.'))");
                            if (marg !=0)
                            {
                                incpvp=((pvpsiniva-marg)*100)/marg;
                                incpvp=incpvp.toFixed(2);
                                marcoArticulosAdd.document.getElementById("pormargen"+ref).innerHTML=incpvp.toString();
                            }
                            if (cost!=0)
                            {
                                inccost=((pvpsiniva-cost)*100)/cost;
                                inccost=inccost.toFixed(2);
                                marcoArticulosAdd.document.getElementById("porcoste"+ref).innerHTML=inccost.toString();
                            }
                            //RGU

                            ref=ref+1
                            break;
                        case "d":
                            marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value="";
                            break;				
                    }
                } // fin for
                if (document.tarifas.chkAplicar.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots&esdto=0&precio="+pvpsiniva;
            }// fin if confirm
        } // fin if valor   
    }

    function AplicarPvpIVA2()
    {
        var msg;
        msg="<%=LITMSGNOAPLITODOS%>";         
        if (document.tarifas.chkAplicar2.checked) msg="<%=LITMSGAPLITODOS%>";
        if (!isNaN(document.tarifas.precioIvab.value.replace(",",".")) && document.tarifas.precioIvab.value!="" && parseFloat(document.tarifas.precioIvab.value)>=0) {
            elementos=marcoArticulosBorrar.document.ArticulosDeTarifa2.length;
            ref=1;
            if (window.confirm(msg)==true)
            {	
                for (i=0;i<=elementos-1;i++) {
                    switch(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].name.substr(0,1)){
                        case "p":
                            precio2=parseFloat(document.tarifas.precioIvab.value.replace(",","."));
                            precio2=precio2.toFixed(<%=DEC_PREC%>);
                            eval("iva=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.hiva" + ref + ".value)");
                            pvpsiniva=precio2/(1+(iva/100))
                            ndecimales=parseInt(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["costdiv" + ref].value.replace(",","."));
                            pvpsiniva=pvpsiniva.toFixed(ndecimales);
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.p" + ref + ".value=pvpsiniva");
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.vppiva" + ref + ".value=document.tarifas.precioIvab.value");
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.ff" + ref + ".value=marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value");
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.d" + ref + ".value=''");
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.cc" + ref + ".value=''");
                            marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["check"+ref].checked=true;

                            //RGU 17/1/2007
                            eval("marg=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.op" + ref + ".value.replace(',','.'))");
                            eval("cost=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.coste" + ref + ".value.replace(',','.'))");
                            if (marg !=0)
                            {
                                incpvp= ((pvpsiniva-marg)*100)/marg;
                                incpvp=incpvp.toFixed(2);
                                marcoArticulosBorrar.document.getElementById("pormargen"+ref).innerHTML=incpvp.toString();
                            }
                            if (cost!=0)
                            {
                                inccost=((pvpsiniva-cost)*100)/cost;
                                inccost=inccost.toFixed(2);
                                marcoArticulosBorrar.document.getElementById("porcoste"+ref).innerHTML=inccost.toString();
                            }
                            //RGU

                            ref=ref+1
                            break;
                        case "d":
                            marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value="";
                            break;
                    }
                } // fin for
                if (document.tarifas.chkAplicar2.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots2&esdto=0&precio="+pvpsiniva;
            } // fin if confirm
        } // fin if val
    }

    function AplicarIncDecCoste()
    {
        var msg;
        msg="<%=LITMSGNOAPLITODOS%>";
        document.tarifas.apliTot.value="off";     
        if (document.tarifas.chkAplicar.checked){
            document.tarifas.apliTot.value="on"; 
            msg="<%=LITMSGAPLITODOS%>";
        }

        incdec=parseFloat(document.tarifas.incdeccoste.value.replace(",","."));
        if (!isNaN(incdec) && incdec!="") 
        {
            if (parseFloat(incdec)<-100) incdec=-100;
            elementos=marcoArticulosAdd.document.ArticulosDeTarifa.length;
            maxelementos=parseInt("<%=MaxArticulos%>");
            elem_comp=0;
            if (window.confirm(msg)==true)
            {	
                for (i=0;i<=elementos-1;i++) {
                    if (marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(0,5)=="check") elem_comp++;
                    if (elem_comp>=1 && elem_comp<=maxelementos){
                        switch(marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(0,1)){
                            case "d":
                                eval("marcoArticulosAdd.document.ArticulosDeTarifa.d" + ref + ".value=''");
                                break;
                            case "p":
                                ref=marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(1);
                                precio=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.elements["coste" + ref].value.replace(",","."));
                                total=parseFloat(precio)+((parseFloat(precio)*incdec)/100);
                                eval("marcoArticulosAdd.document.ArticulosDeTarifa.cc" + ref + ".value=incdec.toString()");
                                ndecimales=parseInt(marcoArticulosAdd.document.ArticulosDeTarifa.elements["costdiv" + ref].value.replace(",","."));
                                eval("iva=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.hiva" + ref + ".value)");
                                pvpiva=parseFloat(total) + ((parseFloat(total)*iva)/100);
                                pvpiva=pvpiva.toFixed(ndecimales);
                                eval("marcoArticulosAdd.document.ArticulosDeTarifa.vppiva" + ref + ".value=pvpiva");
                                total=parseFloat(total).toFixed(<%=dec_prec%>);
                                eval("marcoArticulosAdd.document.ArticulosDeTarifa.ff" + ref + ".value=total.toString()");
                                eval("marcoArticulosAdd.document.ArticulosDeTarifa.p" + ref + ".value=''");
                                marcoArticulosAdd.document.ArticulosDeTarifa.elements["check"+ref].checked=true;

                                //RGU 17/1/2007
                                eval("pvpO=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.op" + ref + ".value.replace(',','.'))");
                                if (pvpO != 0)
                                {
                                    incpvp= ((total-pvpO)*100)/pvpO;
                                    incpvp=incpvp.toFixed(2);
                                    marcoArticulosAdd.document.getElementById("pormargen"+ref).innerHTML=incpvp.toString();
                                }
                                incdec=parseFloat(incdec).toFixed(2);
                                marcoArticulosAdd.document.getElementById("porcoste"+ref ).innerHTML=incdec.toString();//.replace(".",",");
                                //RGU
                                break;
                        }
                    }
                }//fin for
                if (document.tarifas.chkAplicar.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots&esdto=2&dto="+incdec;
            }//fin confirm.	    
        }//fin if val dato
    }

    function AplicarIncDecCoste2()
    {
        var msg;
        msg="<%=LITMSGNOAPLITODOS%>";       
        if (document.tarifas.chkAplicar2.checked) msg="<%=LITMSGAPLITODOS%>";
        incdec=parseFloat(document.tarifas.incdeccosteb.value.replace(",","."));
        if (!isNaN(incdec) && incdec!="")
        {
            if (parseFloat(incdec)<-100) incdec=-100;
            elementos=marcoArticulosBorrar.document.ArticulosDeTarifa2.length;
            maxelementos=parseInt("<%=MaxArticulos%>");
            elem_comp=0;

            if (window.confirm(msg)==true)
            {
                for (i=0;i<=elementos-1;i++) {
                    if (marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].name.substr(0,5)=="check") elem_comp++;
                    if (elem_comp>=1 && elem_comp<=maxelementos){
                        switch(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].name.substr(0,1)){
                            case "d":
                                eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.d" + ref + ".value=''");
                                break;
                            case "p":
                                ref=marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].name.substr(1);
                                precio=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["coste" + ref].value.replace(",","."));
                                total=parseFloat(precio)+((parseFloat(precio)*incdec)/100);
                                eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.cc" + ref + ".value=incdec.toString()");
                                ndecimales=parseInt(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["costdiv" + ref].value.replace(",","."));
                                eval("iva=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.hiva" + ref + ".value)");
                                pvpiva=parseFloat(total) + ((parseFloat(total)*iva)/100);
                                pvpiva=pvpiva.toFixed(ndecimales);
                                eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.vppiva" + ref + ".value=pvpiva");
                                total=parseFloat(total).toFixed(<%=dec_prec%>);
                                eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.ff" + ref + ".value=total.toString()");
                                //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value="";
                                eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.p" + ref + ".value=''");
                                marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["check"+ref].checked=true;

                                //RGU 17/1/2007
                                eval("pvpO=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.op" + ref + ".value.replace(',','.'))");
                                if (pvpO != 0)
                                {
                                    incpvp= ((total-pvpO)*100)/pvpO;
                                    incpvp=incpvp.toFixed(2);
                                    marcoArticulosBorrar.document.getElementById("pormargen"+ref).innerHTML=incpvp.toString();
                                }
                                incdec=parseFloat(incdec).toFixed(2);
                                marcoArticulosBorrar.document.getElementById("porcoste"+ref ).innerHTML=incdec.toString();//.replace(".",",");
                                //RGU
                                break;
                        }
                    } // fin for
                    if (document.tarifas.chkAplicar2.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots2&esdto=2&dto="+incdec;
                } // fin confirm
            } // fin if val
        }
    }

    function AplicarPrecios()
    {
        var msg
        msg="<%=LITMSGNOAPLITODOS%>";
        document.tarifas.apliTot.value="off";     
        if (document.tarifas.chkAplicar.checked){
            document.tarifas.apliTot.value="on"; 
            msg="<%=LITMSGAPLITODOS%>";
        }
		
        if (!isNaN(document.tarifas.pgeneral.value.replace(",",".")) && document.tarifas.pgeneral.value!="" && parseFloat(document.tarifas.pgeneral.value)>=0) {
            elementos=marcoArticulosAdd.document.ArticulosDeTarifa.length;
            ref=1;
            if (window.confirm(msg)==true)
            {
                for (i=0;i<=elementos-1;i++)
                {
                    switch(marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(0,1)){
                        case "p":
                            precio2=parseFloat(document.tarifas.pgeneral.value.replace(",","."));
                            precio2=precio2.toFixed(<%=DEC_PREC%>);
                            //iva=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+2].value);
                            eval("iva=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.hiva" + ref + ".value)");
                            pvpiva=parseFloat(precio2)+((parseFloat(precio2)*iva)/100)
                            ndecimales=parseInt(marcoArticulosAdd.document.ArticulosDeTarifa.elements["costdiv" + ref].value.replace(",","."));

                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value=document.tarifas.pgeneral.value.replace(".",",");
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.p" + ref + ".value=document.tarifas.pgeneral.value");
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+1].value=Redondear (pvpiva,ndecimales).replace(/[.]/g,"")
                            pvpiva=pvpiva.toFixed(ndecimales);
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.vppiva" + ref + ".value=pvpiva");
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+6].value=marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+7].value=marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.ff" + ref + ".value=marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value");
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+4].value="";
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.d" + ref + ".value=''");
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+5].value="";
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.cc" + ref + ".value=''");
                            marcoArticulosAdd.document.ArticulosDeTarifa.elements["check"+ref].checked=true;

                            //RGU 17/1/2007
                            eval("marg=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.op" + ref + ".value.replace(',','.'))");
                            eval("cost=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.coste" + ref + ".value.replace(',','.'))");
                            if (marg !=0)
                            {
                                incpvp= ((precio2-marg)*100)/marg;
                                incpvp=incpvp.toFixed(2);
                                marcoArticulosAdd.document.getElementById("pormargen"+ref).innerHTML=incpvp.toString();
                            }
                            if (cost!=0)
                            {
                                inccost=((precio2-cost)*100)/cost;
                                inccost=inccost.toFixed(2);
                                marcoArticulosAdd.document.getElementById("porcoste"+ref).innerHTML=inccost.toString();
                            }
                            //RGU

                            ref=ref+1
                            break;
                        case "d":
                            marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value="";
                            //eval("marcoArticulosAdd.document.ArticulosDeTarifa.d" + ref + ".value=''");
                            break;
                    }
                }// fin for
                if (document.tarifas.chkAplicar.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots&esdto=0&precio="+precio2;
            }// fin if confirm 	   
        } // fin for
    } 

    function AplicarPrecios2()
    {
        var msg
        msg="<%=LITMSGNOAPLITODOS%>";     
        if (document.tarifas.chkAplicar2.checked) msg="<%=LITMSGAPLITODOS%>";

        if (!isNaN(document.tarifas.pgeneralb.value.replace(",",".")) && document.tarifas.pgeneralb.value!="" && parseFloat(document.tarifas.pgeneralb.value)>=0)
        {
            elementos=marcoArticulosBorrar.document.ArticulosDeTarifa2.length;
            ref=1
            if (window.confirm(msg)==true)
            {
                for (i=0;i<=elementos-1;i++)
                {
                    switch(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].name.substr(0,1)){
                        case "p":

                            precio2=parseFloat(document.tarifas.pgeneralb.value.replace(",","."));
                            precio2=precio2.toFixed(<%=DEC_PREC%>);
                            //iva=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i-4].value);
                            eval("iva=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.hiva" + ref + ".value)");
                            pvpiva=parseFloat(precio2)+((parseFloat(precio2)*iva)/100)

                            ndecimales=parseInt(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["costdiv" + ref].value.replace(",","."));
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+1].value=Redondear (pvpiva,ndecimales).replace(/[.]/g,"")
                            pvpiva=pvpiva.toFixed(ndecimales);
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.vppiva" + ref + ".value=pvpiva");

                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value=document.tarifas.pgeneralb.value.replace(".",",");
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.p" + ref + ".value=document.tarifas.pgeneralb.value");
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+5].value=marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+6].value=marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.ff" + ref + ".value=marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value");
    					
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+3].value="";
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.d" + ref + ".value=''");
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+4].value="";
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.cc" + ref + ".value=''");
                            marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["check"+ref].checked=true;

                            //RGU 17/1/2007
                            eval("marg=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.op" + ref + ".value.replace(',','.'))");
                            eval("cost=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.coste" + ref + ".value.replace(',','.'))");
                            if (marg !=0)
                            {
                                incpvp= ((precio2-marg)*100)/marg;
                                incpvp=incpvp.toFixed(2);
                                marcoArticulosBorrar.document.getElementById("pormargen"+ref).innerHTML=incpvp.toString();
                            }
                            if (cost!=0)
                            {
                                inccost=((precio2-cost)*100)/cost;
                                inccost=inccost.toFixed(2);
                                marcoArticulosBorrar.document.getElementById("porcoste"+ref).innerHTML=inccost.toString();
                            }
                            //RGU

                            ref=ref+1
                            break;
                        case "d":
                            marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value="";
                            //eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.d" + ref + ".value=''");
                            break;
                    }
                }//fin for
                if (document.tarifas.chkAplicar2.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots2&esdto=0&precio="+precio2;
            }// fin confirm
        }//fin val datos
    }

    function AplicarDto()
    {
        var msg
        msg="<%=LITMSGNOAPLITODOS%>";
        document.tarifas.apliTot.value="off";     
        if (document.tarifas.chkAplicar.checked){
            document.tarifas.apliTot.value="on"; 
            msg="<%=LITMSGAPLITODOS%>";
        }

        dto=parseFloat(document.tarifas.dgeneral.value.replace(",","."));
        if (!isNaN(dto) && dto!="")
        {
            if (parseFloat(dto)<-100) dto=-100;
            elementos=marcoArticulosAdd.document.ArticulosDeTarifa.length;
            ref=1;
            if (window.confirm(msg)==true)
            {		
                for (i=0;i<=elementos-1;i++) {
                    switch(marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(0,1)){
                        case "d":
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value=dto.toString().replace(".",",");
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.d" + ref + ".value=dto.toString()");
    					
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+1].value="";
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.cc" + ref + ".value=''");
    					
                            fvalor=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.elements["op"+ref].value.replace(",","."));
                            //fvalor=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.elements["op"+ref].value.replace(",","."));
                            fdto=parseFloat(dto);
                            fvalor=parseFloat(fvalor) + ((parseFloat(fvalor)*fdto)/100);
                            ndecimales=parseInt(marcoArticulosAdd.document.ArticulosDeTarifa.elements["costdiv" + ref].value.replace(",","."));

                            //iva=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.elements[i-2].value);
                            eval("iva=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.hiva" + ref + ".value)");
                            pvpiva=parseFloat(fvalor) + ((parseFloat(fvalor)*iva)/100);
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i-3].value=Redondear (pvpiva,ndecimales).replace(/[.]/g,"")
                            pvpiva=pvpiva.toFixed(ndecimales);
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.vppiva" + ref + ".value=pvpiva");
                            fvalor=parseFloat(fvalor).toFixed(ndecimales);
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i+3].value=fvalor.toString().replace(".",",");
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.ff" + ref + ".value=fvalor.toString()");
                            marcoArticulosAdd.document.ArticulosDeTarifa.elements["check"+ref].checked=true;

                            //RGU 17/1/2007
                            eval("impO=parseFloat(marcoArticulosAdd.document.ArticulosDeTarifa.coste" + ref + ".value.replace(',','.'))");
                            if (impO != 0 )
                            {
                                inccost= ((fvalor-impO)*100)/impO;
                                inccost=inccost.toFixed(2);
                                marcoArticulosAdd.document.getElementById("porcoste"+ref ).innerHTML=inccost.toString();
                            }
                            dto=parseFloat(dto).toFixed(2);
                            marcoArticulosAdd.document.getElementById("pormargen"+ref).innerHTML=dto.toString();//.replace(".",",");
                            //RGU

                            ref=ref+1;
                            break;
                        case "p":
                            /*ref=marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(1);
                            precio=marcoArticulosAdd.document.ArticulosDeTarifa.elements["hp" + ref].value;
                            total=precio-((precio*dto)/100);*/
                            //marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].value="";
                            eval("marcoArticulosAdd.document.ArticulosDeTarifa.p" + ref + ".value=''");
                            break;
                    }			
                } // fin for
                if (document.tarifas.chkAplicar.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots&esdto=1&dto="+dto;
            }// fin if confirm
        } // fin if val
    }

    function AplicarDto2()
    {
        var msg
        msg="<%=LITMSGNOAPLITODOS%>";
        if (document.tarifas.chkAplicar2.checked) msg="<%=LITMSGAPLITODOS%>";

        dto=parseFloat(document.tarifas.dgeneralb.value.replace(",","."));
        if (!isNaN(dto) && dto!="")
        {
            if (parseFloat(dto)<-100) dto=-100;
            elementos=marcoArticulosBorrar.document.ArticulosDeTarifa2.length;
            ref=1
            if (window.confirm(msg)==true)
            {
                for (i=0;i<=elementos-1;i++)
                {
                    switch(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].name.substr(0,1))
                    {
                        case "d":
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value=dto.toString().replace(".",",");
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.d" + ref + ".value=dto.toString()");
    					
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+1].value="";
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.cc" + ref + ".value=''");
    					
                            fvalor=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["op"+ref].value.replace(",","."));
                            //fvalor=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+2].value.replace(",","."));
                            fdto=parseFloat(dto);
                            fvalor=parseFloat(fvalor) + ((parseFloat(fvalor)*fdto)/100);
                            ndecimales=parseInt(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["costdiv" + ref].value.replace(",","."));

                            //iva=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i-7].value);
                            eval("iva=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.hiva" + ref + ".value)");
                            pvpiva=parseFloat(fvalor) + ((parseFloat(fvalor)*iva)/100);
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i-2].value=Redondear (pvpiva,ndecimales).replace(/[.]/g,"")
                            pvpiva=pvpiva.toFixed(ndecimales);
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.vppiva" + ref + ".value=pvpiva");
                            fvalor=parseFloat(fvalor).toFixed(ndecimales);
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i+3].value=fvalor.toString().replace(".",",");
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.ff" + ref + ".value=fvalor.toString()");
                            marcoArticulosBorrar.document.ArticulosDeTarifa2.elements["check"+ref].checked=true;

                            //RGU 17/1/2007
                            eval("impO=parseFloat(marcoArticulosBorrar.document.ArticulosDeTarifa2.coste" + ref + ".value.replace(',','.'))");
                            if (impO != 0 )
                            {
                                inccost= ((fvalor-impO)*100)/impO;
                                inccost=inccost.toFixed(2);
                                marcoArticulosBorrar.document.getElementById("porcoste"+ref ).innerHTML=inccost.toString();
                            }
                            dto=parseFloat(dto).toFixed(2);
                            marcoArticulosBorrar.document.getElementById("pormargen"+ref).innerHTML=dto.toString();
                            //RGU

                            ref=ref+1
                            break;
                        case "p":
                            /*ref=marcoArticulosAdd.document.ArticulosDeTarifa.elements[i].name.substr(1);
                            precio=marcoArticulosAdd.document.ArticulosDeTarifa.elements["hp" + ref].value;
                            total=precio-((precio*dto)/100);*/
                            //marcoArticulosBorrar.document.ArticulosDeTarifa2.elements[i].value="";
                            eval("marcoArticulosBorrar.document.ArticulosDeTarifa2.p" + ref + ".value=''");
                            break;
                    }
                } // fin for
                if (document.tarifas.chkAplicar2.checked) parent.botones.document.location="tarifas_bt.asp?mode=aplicaTots2&esdto=1&dto="+dto;
            }// fin confirm
        }
    }

    function GetDiffDate(date, hour) {

        if(hour == '') hour = "00:00"

        var t = new Date()
        var myDate = date.split("/")
        var myHour = hour.split(":")
        
        var d1 = new Date(myDate[2],myDate[1]-1,myDate[0],myHour[0],myHour[1])
       
        return d1 - t
    }
function isTextSelected(input) 
{
    if (typeof input.selectionStart == "number") {
        return input.selectionStart == 0 &&
        input.selectionEnd == input.value.length;
    } else if (typeof document.selection != "undefined") {
        input.focus();
        return document.selection.createRange().text == input.value;
    }
}

    function GuardarArticulos(mode)
    {
        if (mode=="ALTA")
        { 
            var isValid = /^([0-1][0-9]|2[0-3]):([0-5][0-9])$/.test(document.tarifas.hv.value);
            var hVigor = document.tarifas.hv.value;
            var hoy= new Date();
            //if (marcoArticulosAdd.document.ArticulosDeTarifa.nact.value>0)
            //{
            if ( checkdate(document.tarifas.fv) )
            {
                if (document.tarifas.fv.value != "") {
                    if (document.tarifas.condbase.value == "1") window.alert("<%=LitMsgErrEntVig%>");
                    else {
                        if (GetDiffDate(document.tarifas.fv.value, hVigor) > 0) {
                            if (document.tarifas.hv.value == "") {
                                if (window.confirm("<%=LitMsgCrearAct%>") == true) {
                                    //DIFERIDO:-Código para generar un registro de actualizaciones de precios
                                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            
                                    marcoArticulosAdd.document.ArticulosDeTarifa.action = "ArticulosDeTarifa.asp?mode=save&submode=all&fv=" + document.tarifas.fv.value + "&hv=" + hVigor + "&tarifa=" + document.tarifas.htarifa.value + "&temporada=" + marcoArticulosAdd.document.ArticulosDeTarifa.temporada.value + "&rango=" + marcoArticulosAdd.document.ArticulosDeTarifa.rango.value;
                            
                                    marcoArticulosAdd.document.ArticulosDeTarifa.submit();


                                }
                            } 
                            else {
                                if (isValid) {
                                    if (window.confirm("<%=LitMsgCrearAct%>") == true) {
                                    //DIFERIDO:-Código para generar un registro de actualizaciones de precios
                                    marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            
                                    marcoArticulosAdd.document.ArticulosDeTarifa.action = "ArticulosDeTarifa.asp?mode=save&submode=all&fv=" + document.tarifas.fv.value + "&hv=" + hVigor + "&tarifa=" + document.tarifas.htarifa.value + "&temporada=" + marcoArticulosAdd.document.ArticulosDeTarifa.temporada.value + "&rango=" + marcoArticulosAdd.document.ArticulosDeTarifa.rango.value;
                            
                                    marcoArticulosAdd.document.ArticulosDeTarifa.submit();



                                    }
                                } else {
                                    window.alert("<%=LITMSGERRHORA%>");
                                }
                            }
                        } else {
                            window.alert("<%=LitMsgFechaMayor%>");
                        }
                    }
                }
                else {
                    if (window.confirm("<%=LitMsgCrear%>") == true) {
                        cadena = window.prompt(" <%=LitMsgObservaciones%> ", "");
                        if (window.confirm(" <%=LitMsgCrear2%> ")) {
                            //Crear u n documento de actualizacion de precios
                            marcoArticulosAdd.document.getElementById("waitBoxOculto").style.visibility = "visible";
                            paginaAdd = "ArticulosDeTarifa.asp?mode=save&submode=all&tarifa=" + document.tarifas.htarifa.value + "&observaciones=" + cadena;
                            if (document.tarifas.chkAplicar.checked) paginaAdd = paginaAdd + "&genTodPag=1";
                            else paginaAdd = paginaAdd + "&genTodPag=0";
                        <% 'if si_tiene_modulo_tiendas<>0 then' %>
                                paginaAdd=paginaAdd + "&temporada=" + marcoArticulosAdd.document.ArticulosDeTarifa.temporada.value + "&rango=" + marcoArticulosAdd.document.ArticulosDeTarifa.rango.value
                                    <% 'end if' %>
                            ;
                            marcoArticulosAdd.document.ArticulosDeTarifa.action = paginaAdd;
                            marcoArticulosAdd.document.ArticulosDeTarifa.submit();

                        }
                    }
                }
            }
            else window.alert("<%=LitMsgErrFecha%>");
        }
        else
        {
            var isValid = /^([0-1][0-9]|2[0-3]):([0-5][0-9])$/.test(document.tarifas.bhv.value);
            var bhVigor = document.tarifas.bhv.value;
            var hoy= new Date();
            //if (marcoArticulosAdd.document.ArticulosDeTarifa.nact.value>0) 
            //{
            if ( checkdate(document.tarifas.bfv) )
            {
                if (document.tarifas.bfv.value != "")
                {
                    if (document.tarifas.bcondbase.value=="1") window.alert("<%=LitMsgErrEntVig%>");
                    else
                    {
                        if ( GetDiffDate(document.tarifas.bfv.value, bhVigor) >0 )
                        {
                            if (document.tarifas.bhv.value=="" ) {
                                if(window.confirm("<%=LitMsgCrearAct%>")==true)
                                {
                                    //DIFERIDO:-Código para generar un registro de actualizaciones de precios 
                                    marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
                                    marcoArticulosBorrar.document.ArticulosDeTarifa2.action="ArticulosDeTarifa2.asp?mode=save&submode=all&fv="+document.tarifas.bfv.value+"&bhv="+bhVigor+"&tarifa=" + document.tarifas.htarifa.value
                                    <% 'if si_tiene_modulo_tiendas<>0 then' %>
                                         + "&temporada=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.temporada.value + "&rango=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.rango.value
                                    <%'end if'%>
                                        ;

                                    marcoArticulosBorrar.document.ArticulosDeTarifa2.submit();

                                }
                            } else {
                                if (isValid) {
                                    if(window.confirm("<%=LitMsgCrearAct%>")==true)
                                    {
                                        //DIFERIDO:-Código para generar un registro de actualizaciones de precios 
                                        marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
                                        marcoArticulosBorrar.document.ArticulosDeTarifa2.action="ArticulosDeTarifa2.asp?mode=save&submode=all&fv="+document.tarifas.bfv.value+"&bhv="+bhVigor+"&tarifa=" + document.tarifas.htarifa.value
                                        <%'if si_tiene_modulo_tiendas<>0 then'%>
                                             + "&temporada=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.temporada.value + "&rango=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.rango.value
                                        <%'end if'%>
                                            ;
                                        marcoArticulosBorrar.document.ArticulosDeTarifa2.submit();

                                    }
                                } else {
                                window.alert("<%=LITMSGERRHORA%>");
                                }
                            }
                        }
                        else window.alert("<%=LitMsgFechaMayor%>")
                    }
                }
                else
                {
                    if(window.confirm("<%=LitMsgCrear%>")==true)
                    {
                        cadena= window.prompt(" <%=LitMsgObservaciones%> ","");
                        if (window.confirm(" <%=LitMsgCrear2%> "))
                        {
                            //Crear u n documento de actualizacion de precios
                            marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
                            paginaAdd="ArticulosDeTarifa2.asp?mode=save&submode=all&tarifa=" + document.tarifas.htarifa.value+"&observaciones="+cadena;
                            if (document.tarifas.chkAplicar2.checked) paginaAdd=paginaAdd + "&genTodPag=1";
                            else paginaAdd=paginaAdd + "&genTodPag=0";
                            <%'if si_tiene_modulo_tiendas<>0 then'%>
                                paginaAdd=paginaAdd + "&temporada=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.temporada.value + "&rango=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.rango.value;
                            <%'end if'%>
                            marcoArticulosBorrar.document.ArticulosDeTarifa2.action=paginaAdd;
                            marcoArticulosBorrar.document.ArticulosDeTarifa2.submit();

                        }
                    }
                }
            }
            else window.alert("<%=LitMsgErrFecha%>");
        }
    }

    function BorrarArticulos() {
        if (confirm("<%=LitMsgEliminarRefTarifaConfirm%>")) {
            marcoArticulosBorrar.document.getElementById("waitBoxOculto").style.visibility="visible";
            marcoArticulosBorrar.document.ArticulosDeTarifa2.action="ArticulosDeTarifa2.asp?mode=delete&npagina=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.hnpagina.value + "&tarifa=" + document.tarifas.htarifa.value
            <%'if si_tiene_modulo_tiendas<>0 then'%>
                + "&temporada=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.temporada.value + "&rango=" + marcoArticulosBorrar.document.ArticulosDeTarifa2.rango.value
            <%'end if'%>
                ;
            marcoArticulosBorrar.document.ArticulosDeTarifa2.submit();
        }
    }

    function seleccionar(marco,formulario,check) {
        nregistros=eval(marco + ".document." + formulario + ".hNRegs.value-1");
        if (marco=="marcoRevActualizaciones"){
            if (eval("document.tarifas.seltodos.checked"))
            {
                for (i=1;i<=nregistros;i++)
                {
                    nombre="check" + i;
                    eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
                }
            }
            else{
                for (i=1;i<=nregistros;i++)
                {
                    nombre="check" + i;
                    eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
                }
            }
        }
        else
        {
            if (eval("document.tarifas." + check+ ".checked"))
            {
                for (i=1;i<=nregistros;i++)
                {
                    nombre="check" + i;
                    eval(marco + ".document." + formulario + ".elements[nombre].checked=true;");
                }
            }
            else
            {
                for (i=1;i<=nregistros;i++)
                {
                    nombre="check" + i;
                    eval(marco + ".document." + formulario + ".elements[nombre].checked=false;");
                }
            }
        }

    }

    function BorrarActPre2 ()
    {
        if (marcoRevActualizaciones.document.TarifasRevActualizaciones.nregistros.value!=0 && marcoRevActualizaciones.document.TarifasRevActualizaciones.nregistros.value!="")
        {
            marcoRevActualizaciones.document.TarifasRevActualizaciones.action="TarifasActualizaciones.asp?mode=deleterev";
            marcoRevActualizaciones.document.TarifasRevActualizaciones.submit();

        }
    }

    function ActPre()
    {
        if (marcoRevActualizaciones.document.TarifasRevActualizaciones.nregistros.value!=0 && marcoRevActualizaciones.document.TarifasRevActualizaciones.nregistros.value!="")
        {
            elementos=marcoRevActualizaciones.document.TarifasRevActualizaciones.nregistros.value;
            if (elementos=="") elementos=0;
            error="NO";
            msg="";
            tiene=0;
            for (i=1;i<=elementos-1;i++)
            {
                if (eval("marcoRevActualizaciones.document.TarifasRevActualizaciones.check" + i + ".checked"))
                    eval("isNaN(marcoRevActualizaciones.document.TarifasRevActualizaciones.ff" + i + ".value.replace(',','.')) && marcoRevActualizaciones.document.TarifsRevActualizaciones.ff" + i + ".value!=''");
            }
            if (error=="SI") window.alert(msg);
            else {
                marcoRevActualizaciones.document.TarifasRevActualizaciones.action="TarifasActualizaciones.asp?mode=save";
                marcoRevActualizaciones.document.TarifasRevActualizaciones.submit();
            }
        }
    }
</script>

<body class="BODY_ASP">
<%
'******************************************************************************************************************
'                                             FUNCIONES ASP
'******************************************************************************************************************
sub BarraNavegacion()
    %>
        <script language="javascript" type="text/javascript">
            jQuery("#S_CABECERA").hide();
            jQuery("#PRECIOS").show();
        </script>
    <%

end sub

'**********************************************************************************************************
sub SpanAltasArticulos(tar)
	dis=""
    strselect= "select referencia from articulos_tarifa with (nolock) where tarifa = ?"
    TieneArticulos = DLookupP1(strselect, tar&"", adVarchar,10, session("dsn_cliente"))
	if TieneArticulos="" then dis="disabled"
	'Línea para establecer los parámetros de relleno de iframe
            EligeCelda "input-detail",mode,"left","","",0,LitRefContiene,"refcontiene",20,""
            EligeCelda "input-detail",mode,"left","","",0,LitDesContiene,"descontiene",20,""

            strselect ="select * from tipos_entidades with(nolock) where codigo like ?+'%' and tipo='ARTICULO' order by descripcion"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 = Server.CreateObject("ADODB.Command")
            conn2.Open = session("dsn_cliente")
            conn2.CursorLocation = 3
            command2.ActiveConnection = conn2
            command2.CommandTimeout = 60
            command2.CommandText = strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@sesionNCliente",adVarChar,adParamInput,10,session("ncliente")&"")
            set rstAux= command2.Execute

            DrawDiv "1-detail","",""
            DrawLabel "","",LitTipoArt%><select style="display:; width:175px;" class='CELDAL7' name="tipoarticulo">
			    <%if tipoarticulo="" then %>
			    	<option selected value=""> </option>
				<%else%>
				    <option selected value="<%=EncodeForHtml(tipoarticulo)%>"> <%=EncodeForHtml(trimCodEmpresa(tipoarticulo))%></option>
			    	<option value=""> </option>
				<%end if%>

			 	<%while not rstAux.eof%>
		   			<option value="<%=EncodeForHtml(rstAux("codigo"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
					<%rstAux.movenext%>
				<%wend%>
		   	</select><%CloseDiv
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing
			dim ConfigDespleg (3,13)
			i=0
			ConfigDespleg(i,0)="categoria"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitCategoria
			ConfigDespleg(i,10)=EncodeForHtml(categoria)
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=1
			ConfigDespleg(i,0)="familia_padre"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitFamilia
			ConfigDespleg(i,10)=EncodeForHtml(familia_padre)
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=2
			ConfigDespleg(i,0)="familia"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitSubFamilia
			ConfigDespleg(i,10)=EncodeForHtml(familia)
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			DibujaDesplegablesDetail ConfigDespleg,session("dsn_cliente")
			'DrawCelda2 "CELDA7 style='width:125px'","left",false,LitTemporada & " : "
            strselect = "select codigo,descripcion from temporadas with(nolock) where codigo<>?+'BASE' and codigo like ?+'%' order by descripcion"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open = session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@sesionNCliente",adVarChar,adParamInput,15,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,15,session("ncliente")&"")
            set rstAux= command2.Execute
            DrawSelectCeldaDetail "CELDA7 " & dis,"175","",0,LitTemporada,"temporada",rstAux,"","codigo","descripcion","",""
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing

            strselect = "select codigo,descripcion from rangos with(nolock) where codigo<>?+'BASE' and codigo like ?+'%' order by descripcion"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open = session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@sesionNCliente",adVarChar,adParamInput,15,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,15,session("ncliente")&"")
            set rstAux= command2.Execute
            DrawSelectCeldaDetail "CELDA7 " & dis,"175","",0,LitRango,"rango",rstAux,"","codigo","descripcion","",""
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing

            strselect = "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like ?+'%' order by razon_social"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open = session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,10,session("ncliente")&"")
            set rstAux= command2.Execute
            DrawSelectCeldaDetail "CELDA7","175","",0,LitProveedor,"proveedor",rstAux,"","nproveedor","razon_social","",""
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing
            DrawDiv "col-lg-2 col-md-2 col-sm-2 col-xs-3 col-xxs-6","",""
            DrawLabel "","",LitFechaVigor
            DrawInput "'width70 fvigor'","left","fv","","size='20' "
            DrawCalendar "fv"
            CloseDiv
            DrawDiv "col-lg-2 col-md-2 col-sm-2 col-xs-3 col-xxs-6","",""
            DrawLabel "","",LITHORAVIGOR
            DrawInput "","left","hv","","id='hv'"
            CloseDiv%><script type="text/javascript">
            $("#hv").keypress(function (e) {
                if (isTextSelected(this)) {
                        return;
                    }
                    if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) 
                    {
                        return false;
                    }
                    else
                    {
                        var reg = /[0-9]/;
                        if (this.value.length == 2 && reg.test(this.value))
                            this.value = this.value + ":";
                        if (this.value.length > 4)
                        return false;
                    }
            });
            $(".fvigor").keypress(function (e) {
                if (isTextSelected(this)) {
                    return;
                }
                if(e.which !== 8) {
                    if (this.value.length > 9)
                        return false;
                    var numChars = $(this).val().length;
                    if(numChars === 2 || numChars === 5) {
                        var thisVal = $(this).val();
                        thisVal += '/';
                        $(this).val(thisVal);
                    }
                }
            });</script><%
            DrawDiv "1-detail","",""
            DrawLabel "","",LitTarRefPro1%><input class="CELDA7" type="text" name="RefPro" value="" size="20" id="KeyIntro" runat="javascript:CampoRefPulsado('ALTA','marcoArticulosAdd','ArticulosDeTarifa','A.referencia','asc');"><%DrawLabel "","",LitTarRefPro2
            CloseDiv
            DrawDiv "1-detail","",""
            DrawLabel "","",LitCargarArticulos%><a class="ic-accept noMTop" href="javascript:if(Insertar('ALTA','1','first','A.referencia','asc'));"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>" title="<%=LitCargarArticulos%>"></a><%CloseDiv
	%><div class="">
	<table class="width100 md-table-responsive bCollapse"><%
	escribe1="&darr;"
	escribe2="&harr;"
	escribe3="&harr;"
    colorflecha="blue"
			%><td class="ENCABEZADOL underOrange width5"><input type="Checkbox" name="check" onclick="seleccionar('marcoArticulosAdd','ArticulosDeTarifa','check');" ></td>
			<td class="ENCABEZADOL underOrange width10">
			    <%=LitReferencia%>
			    <!--<a class="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>;" id="OD1A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDeTarifa','A.referencia');" title="<%=LitOrdTarRef & " " & LitOrdSentidoD%>"><%=escribe1%></a>-->
			</td>
			<td class="ENCABEZADOL underOrange width10">
			    <%=LitDescripcion%>
			    <!--<a class="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>;" id="OD2A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDeTarifa','A.nombre');" title="<%=LitOrdTarDesc%>" ><%=escribe2%></a>-->
			</td>
			<td class="ENCABEZADOL underOrange width10">
			    <%=LitSubFamilia%>
			    <!--<a class="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>;" id="OD3A" href="javascript:OrdenarDatos('ALTA','marcoArticulosAdd','ArticulosDeTarifa','F.nombre');" title="<%=LitOrdTarSubf%>" ><%=escribe3%></a>-->
			</td>
			<td class="ENCABEZADOL underOrange width5"><%=LitCoste%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPOrigen%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPTarifa%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPvpIva%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPorcentajePvp%></td>
			<% if si_tiene_modulo_credito then  %>  
			<td class="ENCABEZADOL underOrange width5"><%=LitIncDecLitroPvp%></td>
			<% end if %>                            
			<td class="ENCABEZADOL underOrange width5"><%=LitPorcentajeCoste%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPVPFinal%></td>
	</table>
	<iframe class='width100 iframe-data md-table-responsive' height="260px" name="marcoArticulosAdd" id='frArticulosAdd' src='ArticulosDeTarifa.asp' noresize="noresize"></iframe>
    </div>
			<tr>
                <td class="CELDA7" style="width: 160px;">
				    <div align="left" valign="center" id="Nregs" style="width: 160px; font-weight: bold;"></div>
			    </td>
            </tr>
            <tr>
			<%
            DrawDiv "col-xxs-2 verticalAlignBottom","",""
            DrawLabel "","",LitPvpIvaGralTar%><input class='CELDAR7' name="precioIva" value="" size="5"><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarPvpIVA());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarPVPIVA%>" title="<%=LitAplicarPVPIVA%>"></a><%CloseDiv
            DrawDiv "col-xxs-3 verticalAlignBottom","",""
            DrawLabel "","",LitPrecGenIncDecCoste%><input class='CELDAR7' name="incdeccoste" value="" size="5"><font class="CELDA"> % </font><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarIncDecCoste());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarIncDecCoste%>" title="<%=LitAplicarIncDecCoste%>"></a><%CloseDiv
            DrawDiv "col-xxs-2 verticalAlignBottom","",""
            DrawLabel "","",LitPrecioGral%><input class='CELDAR7' name="pgeneral" value="" size="10"><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarPrecios());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarPrecio%>" title="<%=LitAplicarPrecio%>"></a><%CloseDiv
            DrawDiv "col-xxs-2 verticalAlignBottom","",""
            DrawLabel "","",LitDescuentoGral%><input class='CELDAR7' name="dgeneral" value="" size="5"><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarDto());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarDTO%>" title="<%=LitAplicarDTO%>"></a><%CloseDiv
            %></tr>
			    <div id="IcoIns" class="col-xxs-1" style="visibility: hidden;">
			    <a class='ic-save' href="javascript:if(GuardarArticulos('ALTA'));"><img src="<%=themeIlion%><%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitGuardarArt%>" title="<%=LitGuardarArt%>"></a>
			</div>
            <table style="width: 100%;"></table>
		    <td class="CELDA7" colspan="2" align="left"><div id="IcoIns2" class="col-xxs-12" style="visibility: hidden;"><table border="0" cellpadding="0" cellspacing="0"><tr><%
		        'DrawCelda2 "CELDAB7", "left", false,LITAPLITODOS + ": "
		        'DrawCheckCelda "CELDA","","",0,"","chkAplicar",""
                EligeCelda "check",mode,"","","",0,LITAPLITODOS ,"chkAplicar",0,""%>
		    </tr></table></div></td>
	<script>document.tarifas.chkAplicar.checked=false;</script><%
end sub

'**********************************************************************************************************
sub SpanBajasArticulos()
	'Línea para establecer los parámetros de relleno de iframe
			'Drawcelda2 "CELDA7 style='width:125px'", "left", false, LitRefContiene & " : "
			'DrawInputCelda "CELDA7 style='width:175px'","","",20,0,"","brefcontiene",""
            EligeCelda "input-detail",mode,"left","","",0,LitRefContiene,"brefcontiene",20,""
			'Drawcelda2 "CELDA7 style='width:90px'", "left", false, LitDesContiene & " : "
			'DrawInputCelda "CELDA7 style='width:150px'","","",20,0,"","bdescontiene",""
            EligeCelda "input-detail",mode,"left","","",0,LitDesContiene,"bdescontiene",20,""
            strselect ="select * from tipos_entidades with(nolock) where codigo like ?+'%' and tipo='ARTICULO' order by descripcion"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open = session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@sesionNCliente",adVarChar,adParamInput,10,session("ncliente")&"")
            set rstAux= command2.Execute
            DrawDiv "1-detail","",""
            DrawLabel "","",LitTipoArt%><select style="display:;width:175px;" class='CELDAL7' name="btipoarticulo">
			    <%if tipoarticulo="" then %>
			    	<option selected value=""> </option>
				<%else%>
				    <option selected value="<%=EncodeForHtml(tipoarticulo)%>"> <%=EncodeForHtml(trimCodEmpresa(tipoarticulo))%></option>
			    	<option value=""> </option>
				<%end if%>

			 	<%while not rstAux.eof%>
		   			<option value="<%=EncodeForHtml(rstAux("codigo"))%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
					<%rstAux.movenext%>
				<%wend%></select><%
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing
		    CloseDiv
			dim ConfigDespleg (3,13)

			i=0
			ConfigDespleg(i,0)="categoria1"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitCategoria
			ConfigDespleg(i,10)=EncodeForHtml(categoria)
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=1
			ConfigDespleg(i,0)="familia_padre1"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitFamilia
			ConfigDespleg(i,10)=EncodeForHtml(familia_padre)
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			i=2
			ConfigDespleg(i,0)="familia1"
			ConfigDespleg(i,1)="175"
			ConfigDespleg(i,2)="6"
			ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="ENCABEZADOL"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitSubFamilia
			ConfigDespleg(i,10)=EncodeForHtml(familia)
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			DibujaDesplegablesDetail ConfigDespleg,session("dsn_cliente")
                    
            strselect = "select codigo,descripcion from temporadas with(nolock) where codigo<>?+'BASE' and codigo like ?+'%' order by descripcion"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open = session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@sesionNCliente",adVarChar,adParamInput,15,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,15,session("ncliente")&"")
            set rstAux= command2.Execute
            DrawSelectCeldaDetail "CELDA7","175","",0,LitTemporada,"btemporada",rstAux,"","codigo","descripcion","",""
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing

            strselect = "select codigo,descripcion from rangos with(nolock) where codigo<>?+'BASE' and codigo like ?+'%' order by descripcion"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open = session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@sesionNCliente",adVarChar,adParamInput,15,session("ncliente")&"")
            command2.Parameters.Append command2.CreateParameter("@codigo",adVarChar,adParamInput,15,session("ncliente")&"")
            set rstAux= command2.Execute
            DrawSelectCeldaDetail "CELDA7","175","",0,LitRango ,"brango",rstAux,"","codigo","descripcion","",""
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing
            strselect = "select nproveedor,razon_social from proveedores with(nolock) where nproveedor like ?+'%' order by razon_social"
            set command2 = nothing
            set conn2 = Server.CreateObject("ADODB.Connection")
            set command2 =  Server.CreateObject("ADODB.Command")
            conn2.open = session("dsn_cliente")
            conn2.cursorlocation=3
            command2.ActiveConnection =conn2
            command2.CommandTimeout = 60
            command2.CommandText=strselect
            command2.CommandType = adCmdText
            command2.Parameters.Append command2.CreateParameter("@nproveedor",adVarChar,adParamInput,5,session("ncliente")&"")
            set rstAux= command2.Execute
            DrawSelectCeldaDetail "CELDA7","175","",0,LitProveedor,"bproveedor",rstAux,"","nproveedor","razon_social","",""
            conn2.close
            set conn2    =  nothing
            set command2 =  nothing
            set rstAux  =  nothing
            DrawDiv "col-lg-2 col-md-2 col-sm-2 col-xs-3 col-xxs-6","",""
            DrawLabel "","",LitFechaVigor
            DrawInput "'width70 fvigor2'","left","bfv","","size='20'"
            DrawCalendar "bfv"
            CloseDiv
            DrawDiv "col-lg-2 col-md-2 col-sm-2 col-xs-3 col-xxs-6","",""
            DrawLabel "","",LITHORAVIGOR
            DrawInput "","left","bhv","","id='bhv'"
            CloseDiv%><script type="text/javascript">
            $("#bhv").keypress(function (e) {
                if (isTextSelected(this)) {
                        return;
                    }
                    if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) 
                    {
                        return false;
                    }
                    else
                    {
                        var reg = /[0-9]/;
                        if (this.value.length == 2 && reg.test(this.value))
                            this.value = this.value + ":";
                        if (this.value.length > 4)
                        return false;
                    }
            });
            $(".fvigor2").keypress(function (e) {
                if (isTextSelected(this)) {
                    return;
                }
                if(e.which !== 8) {
                    if (this.value.length > 9)
                        return false;
                    var numChars = $(this).val().length;
                    if(numChars === 2 || numChars === 5) {
                        var thisVal = $(this).val();
                        thisVal += '/';
                        $(this).val(thisVal);
                    }
                }
            });</script><%
	      	DrawDiv "1-detail","",""
            DrawLabel "","",LitTarRefPro1
            %><input class="CELDA7" type="text" name="bRefPro" value="" size="20" id="KeyIntro2" runat="javascript:CampoRefPulsado('BAJA','marcoArticulosBorrar','ArticulosDeTarifa2','A.referencia','asc');"><%
            DrawLabel "","",LitTarRefPro2
            CloseDiv

            DrawDiv "1-detail","",""
            DrawLabel "","",LitCargarArticulos
            %><a class="ic-accept noMTop" href="javascript:if(Insertar('BAJA','1','first','A.referencia','asc'));">
                <img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitCargarArticulos%>" title="<%=LitCargarArticulos%>">
              </a><%
            CloseDiv
            %>
    <div class="">
	<table class="width100 md-table-responsive bCollapse">
	<%escribe1="&darr;"
	escribe2="&harr;"
	escribe3="&harr;"
    colorflecha="blue"
			%><td class="ENCABEZADOL underOrange width5"><input type="Checkbox" name="checkb" onclick="seleccionar('marcoArticulosBorrar','ArticulosDeTarifa2','checkb');"></td>
			<td class="ENCABEZADOL underOrange width10">
			    <%=LitReferencia%>
			    <a class="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>;" id="OD1B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDeTarifa2','A.referencia');" title="<%=LitOrdTarRef & " " & LitOrdSentidoD%>"><%=escribe1%></a>
			</td>
			<td class="ENCABEZADOL underOrange width10">
			    <%=LitDescripcion%>
			    <a class="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>;" id="OD2B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDeTarifa2','A.nombre');" title="<%=LitOrdTarDesc%>" ><%=escribe2%></a>
			</td>
			<td class="ENCABEZADOL underOrange width10">
			    <%=LitSubFamilia%>
			    <a class="CELDAREF7" style="font-size: larger; color: <%=colorflecha%>;" id="OD3B" href="javascript:OrdenarDatos('BAJA','marcoArticulosBorrar','ArticulosDeTarifa2','F.nombre');" title="<%=LitOrdTarSubf%>" ><%=escribe3%></a>
			</td>
			<td class="ENCABEZADOL underOrange width5"><%=LitCoste%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPOrigen%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPTarifa%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPvpIva%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPorcentajePvp%></td>
			<% if si_tiene_modulo_credito then %>   
			<td class="ENCABEZADOL underOrange width5"><%=LitIncDecLitroPvp%></td>
			<% end if %>                            
			<td class="ENCABEZADOL underOrange width5"><%=LitPorcentajeCoste%></td>
			<td class="ENCABEZADOL underOrange width5"><%=LitPVPFinal%></td>
	</table>
	<iframe class='width100 iframe-data md-table-responsive' height="250px" name="marcoArticulosBorrar" id='frArticulosBorrar' src='ArticulosDeTarifa2.asp' noresize="noresize"></iframe>
    </div>
			<tr>
                <td class="CELDA7" style="width: 160px;">
				    <div align="left" id="Nregsb" style="width: 160px; font-weight: bold;"></div>
			    </td>
            </tr>
            <tr><%
                DrawDiv "col-xxs-2 verticalAlignBottom","",""
                DrawLabel "","",LitPvpIvaGralTar%><input class='CELDAR7' name="precioIvab" value="" size="5"><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarPvpIVA2());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarPVPIVA%>" title="<%=LitAplicarPVPIVA%>"></a><%CloseDiv
                DrawDiv "col-xxs-3 verticalAlignBottom","",""
                DrawLabel "","",LitPrecGenIncDecCoste%><input class='CELDAR7' name="incdeccosteb" value="" size="5"><font class="CELDA"> % </font><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarIncDecCoste2());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarIncDecCoste%>" title="<%=LitAplicarIncDecCoste%>"></a><%CloseDiv
                DrawDiv "col-xxs-2 verticalAlignBottom","",""
                DrawLabel "","",LitPrecioGral%><input class='CELDAR7' name="pgeneralb" value="" size="10"><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarPrecios2());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarPrecio%>" title="<%=LitAplicarPrecio%>"></a><%CloseDiv
                DrawDiv "col-xxs-2 verticalAlignBottom","",""
                DrawLabel "","",LitDescuentoGral%><input class='CELDAR7' name="dgeneralb" value="" size="5"><a class="ic-accept noMTop inlineBlock floatNone" href="javascript:if(AplicarDto2());"><img align="center" src="<%=themeIlion %><%=ImgRefresh%>" <%=ParamImgAplicar%> alt="<%=LitAplicarDTO%>" title="<%=LitAplicarDTO%>"></a><%CloseDiv
                %></tr>
				<div class="col-xxs-1 verticalAlignBottom" id="IcoBorrModif" style="visibility: hidden;"><a class='ic-save verticalAlignBottom' href="javascript:if(GuardarArticulos('BAJA'));"><img src="<%=themeIlion%><%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitGuardarArt%>" title="<%=LitGuardarArt%>"></a>&nbsp;
				<a class='ic-delete verticalAlignBottom noMBottom' href="javascript:if(BorrarArticulos());"><img src="<%=themeIlion %><%=ImgEliminarDet%>" <%=ParamImgEliminar%> alt="<%=LitEliminarArt%>" title="<%=LitEliminarArt%>"></a></div>
            <table style="width: 100%;"></table>
		    <td class="CELDA7" colspan="2" align="left"><div class="col-xxs-6" id="IcoBorrModif2" style="visibility: hidden;"><table border="0" cellpadding="0" cellspacing="0"><tr><%
 
                EligeCelda "check",mode,"left","","",0,LITAPLITODOS,"chkAplicar2",0,""%>
		</tr></table></div></td>
	<script>document.tarifas.chkAplicar2.checked=false;</script><%
end sub

sub SpanRevAct (p_codigo)
    if si_tiene_modulo_credito then
        tamano_tabla_tarifas=800
    else
        tamano_tabla_tarifas=730
    end if
	%>
    <div class="">
        <table class="width100 md-table-responsive bCollapse" >
			<td class="ENCABEZADOL underOrange width5" ><input type='checkbox' name='seltodos' onclick="seleccionar('marcoRevActualizaciones','TarifasRevActualizaciones','check');" ></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitReferencia%>
			<td class="ENCABEZADOL underOrange width10"><%=LitDescripcion%></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitCoste%></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitPOrigen%></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitPTarifa%></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitPvpIva%></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitPorcentajePvp%></td>
			<% if si_tiene_modulo_credito then %>
			    <td class="ENCABEZADOL underOrange width10"><%=LitIncDecLitroPvp%></td>
			<% end if %>
			<td class="ENCABEZADOL underOrange width10"><%=LitPorcentajeCoste%></td>
			<td class="ENCABEZADOL underOrange width10"><%=LitPVPFinal%></td><%
	%></table>
		<iframe class='width100 iframe-data md-table-responsive' height="250px" name="marcoRevActualizaciones" id='frrevactualizaciones' src='TarifasActualizaciones.asp?mode=revisar&tarifa=<%=EncodeForHtml(p_codigo)%>' noresize="noresize"></iframe>
    </div>
<table style="width:90%"><%
			%><td class='CELDAL7' width="90%">
				<SPAN ID="barrasTar" style="display:">
				</SPAN>
			</td>
			<td class="CELDAR7">
			<a class='CELDAREF' href="javascript:ActPre();">
				<img src="<%=themeIlion%><%=ImgDiskette%>" <%=ParamImgDiskette%> alt="<%=LitAct%>" title="<%=LitAct%>"></a>
			<a style="text-align:right" class='ic-delete noMTop noMBottom' href="javascript:BorrarActPre2();">
				<img src="<%=themeIlion %><%=ImgEliminarDet%>" <%=ParamImgEliminar%> alt="<%=LitEl%>" title="<%=LitEl%>"></a></td>
			<%
	  %></table><%
end sub

sub SpanActPendiente(p_codigo)
	%>
    <div class="">
        <table class="width100 md-table-responsive bCollapse" ><%
				DrawCeldaDet "'ENCABEZADOL underOrange width25'","","",0,LitFV
				'DrawCeldaDet "'ENCABEZADOL underOrange width20'","","",0,LITHORAVIGOR
				DrawCeldaDet "'ENCABEZADOL underOrange width25'","","",0,LitArt
				DrawCeldaDet "'ENCABEZADOL underOrange width25'","","",0,LitRev
				DrawCeldaDet "'ENCABEZADOL underOrange width25'","","",0,LitEl
	%></table>
		<iframe class='width100 iframe-data md-table-responsive' height="250px" name="marcoActualizaciones" id='fractualizaciones' src='TarifasActualizaciones.asp?mode=ver&tarifa=<%=enc.EncodeForJavascript(p_codigo)%>' noresize="noresize"></iframe>
    </div>
<%end sub

'**************************************************************************************************
'                                   Código principal de la página
'**************************************************************************************************

   %><form name="tarifas" method="post" action="tarifas.asp"><%

	si_tiene_modulo_tiendas=ModuloContratado(session("ncliente"),ModTiendas)
	'JMMMM - 28/01/2010 --> Se añade modulo línea de crédito
	si_tiene_modulo_credito=ModuloContratado(session("ncliente"),ModLineaCredito)

	set rst = server.CreateObject("ADODB.Recordset")
	set rstAux = server.CreateObject("ADODB.Recordset")
	set rstAux2 = server.CreateObject("ADODB.Recordset")
	set rstAux3 = server.CreateObject("ADODB.Recordset")

	mode=EncodeForHtml(request.querystring("mode"))
	mode2=EncodeForHtml(request.querystring("mode2"))
    mode_accesos_tienda=mode
    if mode_accesos_tienda & ""="" then
        mode_accesos_tienda="add"
    end if
    %><input type="hidden" name="mode_accesos_tienda" value="<%=EncodeForHtml(mode_accesos_tienda)%>"/><%

	codigoI=left(limpiaCadena(Request.Form("i_codigo")),5)
	descripcionI=limpiaCadena(request.form("i_descripcion"))
	observacionesI = nulear(limpiaCadena(request.form("i_observaciones")))

	codigoE=limpiaCadena(Request.Form("e_codigo"))
	CheckCadena codigoE
	descripcionE=limpiaCadena(request.form("e_descripcion"))
	observacionesE = nulear(limpiaCadena(request.form("e_observaciones")))

	condbase=enc.EncodeForJavascript(request.form("condbase"))
	bcondbase=enc.EncodeForJavascript(request.form("bcondbase"))

    dniSELECT= "select dni from personal with (nolock) where login = ? and dni like ?+'%'"
    dni = DLookupP2(dniSELECT, session("usuario")&"", adVarchar, 50, session("ncliente")&"", adVarchar, 20, session("dsn_cliente"))%>
    <SPAN ID="ComprobarPerCom" style="display:none">
        <%waitbox LitMsgUsuarioPersonalNoExiste%>
    </SPAN>
	<%if dni & "">"" then
	else%>
	    <script language="javascript" type="text/javascript">
	        parent.botones.document.location = "tarifas_bt.asp?mode=XXXX";
	        //ComprobarPerCom.style.display = "";
	        jQuery("#ComprobarPerCom").show();
	    </script>
	    <%mode="NO ESTA DADO DE ALTA EN PERSONAL"
	    submode=mode
	    viene=mode
	    ncliente=mode
        Response.End
	end if

	if condbase&""=""then
		condbase="0"
	end if
	if bcondbase&""=""then
		bcondbase="0"
	end if
					
	%><input type="hidden" name="condbase" value="<%=EncodeForHtml(condbase)%>">
	<input type="hidden" name="bcondbase" value="<%=EncodeForHtml(bcondbase)%>">
	<input type="hidden" name="apliTot" value="off"><%

	if mode="delete" then
		p_codigo=limpiaCadena(request("codigo"))
''ricardo 28-1-2008 solamente se concatenara el nempresa cuando venga del mode=add
        'mmg:evita que casque el CheckCadena y te expulse del sistema
		if mode2="add" then
			p_codigo=session("ncliente")& p_codigo
		end if
	else
		p_codigo=limpiaCadena(request.form("codigo"))
		if p_codigo="" then
			p_codigo=limpiaCadena(request.querystring("codigo"))
		end if
		if p_codigo="" then p_codigo=limpiaCadena(request("p_codigo"))
		'DGM Añadimos la opción de que la pantalla proceda de central.asp, 
		'por lo que el codigo viene encapsualdo en el paramtero ndoc
		if request("ndoc")&"" > "" then
		    p_codigo = limpiaCadena(request("ndoc"))
		    pagina = "1"
		end if
	end if
	CheckCadena p_codigo

	'insertamos si nos llegan los valores
	if codigoI>"" and descripcionI>"" then
        rst.Open "select * from tarifas where codigo='" & session("ncliente")&codigoI & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		if rst.EOF then

            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas alta inicio " & session("ncliente")&codigoI & """," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","alta","","","precios"

			rst.AddNew
			rst("codigo")  = session("ncliente")&codigoI
			rst("descripcion")   = descripcionI
			rst("observaciones") = observacionesI
			rst.Update

            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas alta fin " & session("ncliente")&codigoI & """," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","alta","","","precios"

		else%>
			<script>
				window.alert("<%=LitMsgCodigoExiste%>");
				history.back();
			</script>
		<%end if
        rst.Close
	end if

	'actualizamos valores
	if codigoE>"" and descripcionE>"" and mode<>"delete" then
        rst.Open "select * from tarifas with(rowlock) where codigo='" & codigoE & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
		if not rst.EOF then
            
            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas modificar inicio " & codigoE & """," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","modificar","","","precios"

			'rst("codigo")  = codigoE
			rst("descripcion")   = descripcionE
			rst("observaciones") = observacionesE & ""
			rst.Update

            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas modificar fin " & codigoE & """," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","modificar","","","precios"

		else %>
			<script>
				window.alert("<%=LitMsgCodigoNoExiste%>");
				history.back();
			</script>
		<%end if
        rst.Close
	end if

	'eliminamos valores
	if mode="delete" and p_codigo>"" then
		'miramos a ver si esta puesta en algun documento
		no_borrar=0
        rst.cursorlocation=3
		rst.open "select tarifa from facturas_cli with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")
		
        if not rst.eof then
			no_borrar=1
		end if

        rst.close
		rst.cursorlocation=3
		rst.open "select tarifa from pedidos_cli with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")
		
        if not rst.eof then
			no_borrar=1
		end if

        rst.close
		rst.cursorlocation=3
		rst.open "select tarifa from albaranes_cli with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")

		if not rst.eof then
			no_borrar=1
		end if

        rst.close
		rst.cursorlocation=3
		rst.open "select tarifa from clientes with(nolock) where tarifa='" & p_codigo & "'",session("dsn_cliente")

		if not rst.eof then
			no_borrar=1
		end if
        
        rst.close

		if no_borrar=0 then

            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas baja inicio " & p_codigo & """," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

			%><div ID="splashScreenTarifas" style="position:absolute;z-index:5;top:30%;left:35%;">
				<table bgcolor="<%=color_negro_det%>" BORDER=1 BORDERCOLOR="<%=color_negro_det%>" cellpadding=0 cellspacing=0 HEIGHT=200 WIDTH=300>
					<tr>
						<td style="text-align:center" WIDTH="100%" HEIGHT="100%" bgcolor="<%=color_gris%>" >
							<BR><BR> &nbsp; &nbsp;
							<FONT FACE="Helvetica,Verdana,Arial" SIZE=3 COLOR="<%=color_azul_oscuro%>"><%=LitMsgBorrandoTarifa%></FONT>
							&nbsp; &nbsp; <BR>
							<img src="<%=ImgProcesandoEspere%>" <%=ParamImgProcesandoEspere%>><BR/><BR/>
						</td>
					</tr>
				</table>
			</div><%
			response.flush
            rst.open "select referencia,rango,temporada from precios with(nolock) where tarifa='" & p_codigo & "' order by referencia",session("dsn_cliente"),adUseClient, adLockReadOnly

			RefAnt=""

			for i=1 to rst.recordcount
				if rst("referencia")<>RefAnt then
                    sel1="select distinct rango from precios with(nolock) where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "' and rango<>'BASE'"
					rstAux.cursorlocation=3
					rstAux.open sel1,session("dsn_cliente")
					for j=1 to rstAux.recordcount
						sel2="select * from precios with(nolock) where referencia='" & rst("referencia") & "' and rango='" & rstAux("rango") & "' and tarifa<>'" & p_codigo & "'"
						rstAux2.cursorlocation=3
						rstAux2.open sel2,session("dsn_cliente")
						if rstAux2.eof then

                            registro = "{""fecha"":""" & now & """," &_
                            """usuario"":""" & session("usuario") & """," &_
                            """accion"":""precios tarifas baja inicio articulos_rango where referencia='" & rst("referencia") & "' and rango='" & rstAux("rango") & "'""," &_
                            """num"":1}"
                            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

							'SE puede borrar el registro de ARTICULOS_RANGO
							rstAux3.open "delete from articulos_rango with(rowlock) where referencia='" & rst("referencia") & "' and rango='" & rstAux("rango") & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
                            
                            registro = "{""fecha"":""" & now & """," &_
                            """usuario"":""" & session("usuario") & """," &_
                            """accion"":""precios tarifas baja fin articulos_rango where referencia='" & rst("referencia") & "' and rango='" & rstAux("rango") & "'""," &_
                            """num"":1}"
                            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

                        end if
						rstAux2.close
						rstAux.movenext
					next

					rstAux.close
					sel1="select distinct temporada from precios with(nolock) where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "' and temporada<>'BASE'"
					rstAux.cursorlocation=3
					rstAux.open sel1,session("dsn_cliente")
					for j=1 to rstAux.recordcount
						sel2="select * from precios with(nolock) where referencia='" & rst("referencia") & "' and temporada='" & rstAux("temporada") & "' and tarifa<>'" & p_codigo & "'"
						rstAux2.cursorlocation=3
						rstAux2.open sel2,session("dsn_cliente")
						if rstAux2.eof then
                						    
                            registro = "{""fecha"":""" & now & """," &_
                            """usuario"":""" & session("usuario") & """," &_
                            """accion"":""precios tarifas baja inicio articulos_temporada where referencia='" & rst("referencia") & "' and temporada='" & rstAux("temporada") & "'""," &_
                            """num"":1}"
                            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

							'SE puede borrar el registro de ARTICULOS_TEMPORADA
							rstAux3.open "delete from articulos_temporada with(rowlock) where referencia='" & rst("referencia") & "' and temporada='" & rstAux("temporada") & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
						    
                            registro = "{""fecha"":""" & now & """," &_
                            """usuario"":""" & session("usuario") & """," &_
                            """accion"":""precios tarifas baja fin articulos_temporada where referencia='" & rst("referencia") & "' and temporada='" & rstAux("temporada") & "'""," &_
                            """num"":1}"
                            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"
                        
                        end if
						rstAux2.close
						rstAux.movenext
					next
					rstAux.close

                    registro = "{""fecha"":""" & now & """," &_
                    """usuario"":""" & session("usuario") & """," &_
                    """accion"":""precios tarifas baja inicio articulos_tarifa where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "'""," &_
                    """num"":1}"
                    auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

					rstAux.open "delete from articulos_tarifa with(rowlock) where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
                    
                    registro = "{""fecha"":""" & now & """," &_
                    """usuario"":""" & session("usuario") & """," &_
                    """accion"":""precios tarifas baja fin articulos_tarifa where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "'""," &_
                    """num"":1}"
                    auditar_ins_bor session("usuario"),registro,"","baja","","","precios"
                
                end if
				RefAnt=rst("referencia")

                registro = "{""fecha"":""" & now & """," &_
                """usuario"":""" & session("usuario") & """," &_
                """accion"":""precios tarifas baja inicio precios where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "' and rango='" & rst("rango") & "' and temporada='" & rst("temporada") & "'""," &_
                """num"":1}"
                auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

				rstAux.open "delete from precios with(rowlock) where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "' and rango='" & rst("rango") & "' and temporada='" & rst("temporada") & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
				
                registro = "{""fecha"":""" & now & """," &_
                """usuario"":""" & session("usuario") & """," &_
                """accion"":""precios tarifas baja fin precios where referencia='" & rst("referencia") & "' and tarifa='" & p_codigo & "' and rango='" & rst("rango") & "' and temporada='" & rst("temporada") & "'""," &_
                """num"":1}"
                auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

                rst.movenext
		    next
			rst.close

            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas baja inicio tarifas where codigo='" & p_codigo & "'""," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

			'AHORA SE BORRA LA TARIFA
		    rst.Open "delete from tarifas with(rowlock) where codigo='" & p_codigo & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
            
            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas baja fin tarifas where codigo='" & p_codigo & "'""," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

            registro = "{""fecha"":""" & now & """," &_
            """usuario"":""" & session("usuario") & """," &_
            """accion"":""precios tarifas baja fin " & p_codigo & """," &_
            """num"":1}"
            auditar_ins_bor session("usuario"),registro,"","baja","","","precios"

			%><script language="javascript" type="text/javascript">
			    document.getElementById("splashScreenTarifas").style.visibility = "hidden";
			    document.getElementById("splashScreen").style.visibility = "hidden";
			</script><%
		else
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitTarifaNoBorrarAsocDoc%>");
			</script><%
		end if
	end if

      p_criterio=limpiaCadena(request("criterio"))
      p_campo=limpiaCadena(request("campo"))
      p_texto = limpiaCadena(request.QueryString("texto"))  
      p_npagina=limpiaCadena(request("npagina"))

      if p_texto>"" then
	  	 if p_campo="codigo" then p_campo="substring(codigo,6,10)"
         c_where=" where " & p_campo & " "
      else
         c_where=""
      end if

      if c_where>"" then
         select case p_criterio
            case "contiene"
               c_where=c_where+ "like '%" & p_texto & "%'"
            case "termina"
               c_where=c_where+ "like '%" & p_texto & "'"
            case "empieza"
               c_where=c_where+ "like '" & p_texto & "%'"
            case "igual"
              c_where=c_where + "='" & p_texto & "'"
         end select
		 c_where=c_where & " and codigo like '" & session("ncliente") & "%' and TarifaCliente is null"
	  else
	  	 c_where=" where codigo like '" & session("ncliente") & "%' and TarifaCliente is null"
      end if

   PintarCabecera "tarifas.asp"%>

    <div class="btn-radio-button-wrapper">
		<input type="radio" class="btn-radio-button" checked="checked" name="btnTarifas" id="btnTarifas"/>
			<label for="btnTarifas"><%=LitTitulo%></label>
		<input type="radio" class="btn-radio-button" name="btnTemporadas" id="btnTemporadas" onclick="IrATemporadas();" />
			<label for="btnTemporadas"><%=LitTemporadas%></label>
        <input type="radio" class="btn-radio-button" name="btnRango" id="btnRango" onclick="IrARangos();" />
			<label for="btnRango"><%=LitRangos%></label>
    </div>
<% Alarma "tarifas.asp" %>
    <hr/>
    <%c_select="select * from tarifas with(nolock)"

    if c_where>"" then
       c_select=c_select & c_where
    end if

    if p_npagina="" then
       p_npagina=1
    end if

    select case request("pagina")
       case "siguiente"
          p_npagina=p_npagina+1
       case "anterior"
          p_npagina=p_npagina-1
    end select%>
    <input type="hidden" name="h_npagina" value="<%=EncodeForHtml(cstr(p_npagina))%>">
	    <%
        rst.cursorlocation=3
        rst.Open c_select,session("dsn_cliente")

        if not rst.EOF then
           rst.PageSize=NumReg
           rst.AbsolutePage=p_npagina
        end if
    DrawDiv "col-xxs-12","",""
  if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
		 <a class=CABECERA href="tarifas.asp?pagina=anterior&npagina=<%=EncodeForHtml(cstr(p_npagina))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%=EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
  		 <img src="<%=themeIlion%><%=ImgAnterior%>" align="top" alt="<%=LitAnterior%>" title="<%=LitAnterior%>"></a>
  	<%end if

    texto=LitPagina & " " & cstr(p_npagina) & " " & LitDe & " " & cstr(rst.PageCount)%>
  	    <font class='CELDA'> <%=EncodeForHtml(texto)%> </font> <%

         if clng(p_npagina)<rst.PageCount then %>
		    <a class=CABECERA href="tarifas.asp?pagina=siguiente&npagina=<%=EncodeForHtml(cstr(p_npagina))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%=EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
  		    <img src="<%=themeIlion%><%=ImgSiguiente%>" align="top" alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"></a>
  	    <%end if

	    %><font class='CELDA'>&nbsp;&nbsp; <%=LitPagIrA%> <input class='CELDA' type="text" name="SaltoPagina1" size="2">&nbsp;&nbsp;<a class='CELDAREF inlineBlock' href="javascript:IrAPagina(1,'<%=enc.EncodeForJavascript(p_campo)%>','<%=enc.EncodeForJavascript(p_criterio)%>','<%=enc.EncodeForJavascript(p_texto)%>',<%=rst.PageCount%>,'npagina');"><%=LitIr%></a></font><%
    
  end if
  CloseDiv
  if mode<>"edit" then%>

   <table id="SearchResult" width="100%" border='0' cellspacing="1" cellpadding="1">
       <tr class="underOrange"><%
           DrawCeldaDet "CELDAL7","left","",0,LitCodigo
           DrawCeldaDet "CELDAL7","left","",0,LitDescripcion
           DrawCeldaDet "CELDAL7","left","",0,LitObservaciones%>
       </tr><%

        par=false
        i=1

        while not rst.EOF and i<=NumReg
           if mode="edit" and p_codigo=rst("codigo") then

           elseif mode<>"edit" then
                h_ref="javascript:Editar('" & enc.EncodeForJavascript(rst("codigo")) & "'," & _
			                              enc.EncodeForJavascript(p_npagina) & ",'" & _
				  					      enc.EncodeForJavascript(p_campo) & "','" & _
									      enc.EncodeForJavascript(p_criterio) & "','" & _
									      enc.EncodeForJavascript(p_texto) & "');"
				if ucase(rst("codigo"))<>session("ncliente") & "BASE" then
					if par then
						par=false
					else
	            	  	par=true
					end if
                    %><tr><td class="CELDA7"><%
                    DrawHref "CELDAREF","",EncodeForHtml(trimCodEmpresa(rst("codigo"))),h_ref
                    %></td><%
					DrawCeldaDet "CELDAL7", "left", "",0, EncodeForHtml(rst("descripcion"))
					DrawCeldaDet "CELDAL7", "left", "",0, pintar_saltos_nuevo(EncodeForHtml(rst("observaciones") & ""))
                    %></tr><%
				end if
           end if

           i = i + 1
           rst.MoveNext
        wend%>
   </table>

    <%
    end if
    DrawDiv "col-xxs-12","",""
    if mode<>"edit" and rst.RecordCount>NumReg then
      if clng(p_npagina) >1 then %>
		 <a class=CABECERA href="tarifas.asp?pagina=anterior&npagina=<%=EncodeForHtml(cstr(p_npagina))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%=EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
  		 <img src="<%=themeIlion%><%=ImgAnterior%>" align="top" alt="<%=LitAnterior%>" title="<%=LitAnterior%>"></a>
  	<%end if

    texto=LitPagina & " " & cstr(p_npagina) & " " & LitDe & " " & cstr(rst.PageCount)%>
  	    <font class='CELDA'> <%=EncodeForHtml(texto)%> </font> <%

         if clng(p_npagina)<rst.PageCount then %>
		    <a class=CABECERA href="tarifas.asp?pagina=siguiente&npagina=<%=EncodeForHtml(cstr(p_npagina))%>&campo=<%=EncodeForHtml(p_campo)%>&criterio=<%=EncodeForHtml(p_criterio)%>&texto=<%=EncodeForHtml(p_texto)%>">
  		    <img src="<%=themeIlion%><%=ImgSiguiente%>" align="top" alt="<%=LitSiguiente%>" title="<%=LitSiguiente%>"></a>
  	    <%end if%>
	    <font class='CELDA'>&nbsp;&nbsp; <%=LitPagIrA%> <input class='CELDA' type="text" name="SaltoPagina2" size="2">&nbsp;&nbsp;<a class='CELDAREF inlineBlock' href="javascript:IrAPagina(2,'<%=EncodeForHtml(p_campo)%>','<%=EncodeForHtml(p_criterio)%>','<%=EncodeForHtml(p_texto)%>',<%=rst.PageCount%>,'npagina');"><%=LitIr%></a></font><%
    
	rst.close
  end if
  CloseDiv%>

   <%if mode<>"edit" then %>
   <hr/>
   <table width="100%" border='0' cellspacing="1" cellpadding="1">
   <%
       Drawh6 "col-xxs-12","",LitNBregistro
   %></table>
		<table id="Test" width="100%" border='0' cellspacing="1" cellpadding="1"><%
            DrawDiv "col-lg-2 col-md-2 col-sm-2 col-xs-6 col-xxs-12","",""
            DrawLabel "txtMandatory","",LitCodigo
            DrawInput "CELDA","","i_codigo","","size='5'"
            CloseDiv
            DrawDiv "col-lg-4 col-md-4 col-sm-4 col-xs-6 col-xxs-12","",""
            DrawLabel "txtMandatory","",LitDescripcion
            DrawInput "CELDA","","i_descripcion","","size='57'"
            CloseDiv
            EligeCelda "text-cabecera-detail2","add","CELDA","","",2,LitObservaciones,"i_observaciones",2,""

      %></table>
      <hr/>
   <%end if

   '***************************************************************************
   'Zona de código para la gestión de artículos de la tarifa
   '***************************************************************************

   if mode="edit" then
   		%><%BarraNavegacion%>



        <div class="Section" id="S_CABECERA">
        <a href="#" rel="toggle[CABECERA]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader"><%=LITCABECERA%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
            </div>
        </a>
        <div class="SectionPanel" id="CABECERA" style="display:none;">

		<input type="hidden" name="htarifa" value="<%=EncodeForHtml(p_codigo)%>">

		   <table width="100%" border='0' cellspacing="1" cellpadding="1">
			  <%
				par=false
				i=1
				while not rst.EOF and i<=NumReg
				   if mode="edit" and p_codigo=rst("codigo") then
                        DrawDiv "col-lg-2 col-md-2 col-sm-2 col-xs-6 col-xxs-12","",""
                        DrawLabel "txtMandatory","",LitCodigo
                        DrawSpan "CELDA","",EncodeForHtml(trimCodEmpresa(rst("codigo"))),""
                        CloseDiv
						%><input type="hidden" name="e_codigo" value="<%=EncodeForHtml(rst("codigo"))%>"><%
                        DrawDiv "col-lg-4 col-md-4 col-sm-4 col-xs-6 col-xxs-12","",""
                        DrawLabel "txtMandatory","",LitDescripcion
                        DrawInput "CELDA","","e_descripcion",EncodeForHtml(rst("descripcion")),"size='57'"
                        CloseDiv
                        EligeCelda "text-cabecera-detail2",mode,"","","",0,LitObservaciones,"e_observaciones",2,EncodeForHtml(rst("observaciones"))
					  'CloseFila
				   end if

				   i = i + 1
				   rst.MoveNext
				wend
				'rst.Close %>
		   </table>
        </div>
        </div>
        <div class="Section" id="S_PRECIOS">
        <a href="#" rel="toggle[PRECIOS]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader"><%=LITPRECIOS%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" title="" <%=ParamImgCollapse %> />
            </div>
        </a>
        
        <div class="SectionPanel" id="PRECIOS">
                <div id="tabs" style="display:run-in" class="ui-tabs ui-widget ui-widget-content ui-corner-all">
                <ul class="ui-tabs-nav ui-helper-reset ui-helper-clearfix ui-widget-header ui-corner-all">
                    <li class="ui-state-default ui-corner-top"><a id="tabs1link" href="#tabs1"><%=LitAnadirArticulos%></a></li>
                    <li class="ui-state-default ui-corner-top"><a id="tabs2link" href="#tabs2"><%=LitBorrModifArticulos%></a></li>
                    <li class="ui-state-default ui-corner-top"><a id="tabs3link" onclick="ReloadTab(3);return false;" href="#tabs3"><%=LitActPend%></a></li>
                    <li class="ui-state-default ui-corner-top"><a id="tabs4link" onclick="ReloadTab(4);return false;" href="#tabs4"><%=LitModActPend%></a></li>
                </ul>
                <div id="tabs1" >
			        <%SpanAltasArticulos p_codigo%>
		        </div>
                <div id="tabs2" >
			        <%SpanBajasArticulos%>
		        </div>
                <div id="tabs3" >
			        <%SpanActPendiente (p_codigo)%>
		        </div>
                <div id="tabs4" >
			        <%SpanRevAct (p_codigo)%>
		        </div>
        </div>
        </div>


	<!--</td></tr></table>-->
	<%end if%>
   </form>
<%
	set rst = nothing
	set rstAux = nothing
	set rstAux2 = nothing
	set rstAux3 = nothing
end if%>
</body>
</html>