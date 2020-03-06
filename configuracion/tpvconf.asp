<%@ Language=VBScript %>
<%
''JA 20-03-2003: Gesti�n del campo COTA.
''Toni Climent 14-01-2009: Gestion del campo HORECAS
''TCD 19-01-2009 Gesti�n del campo PAGOSTCATICKETS
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html LANG="<%=session("lenguaje")%>">

<% dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")%>  

<head>
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="tpvconf.inc" -->
<title><%=LitTitulo%></title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0"/>
<meta HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>"/>
<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript" src="/lib/js/shortKey.js"></script>
<!--#include file="../ilion.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/font-face.css.inc" -->
<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../CatFamSubResponsive.inc"-->
<!--#include file="tpvconf_linkextra.inc" -->
<!--#include file="../js/dropdown.js.inc" -->
<!--#include file="../styles/formularios.css.inc" --> 
<!--#include file="../styles/dropdown.css.inc" -->
<script type="text/javascript" language="javascript">


    animatedcollapse.addDiv('DatosDC', 'fade=1')
    animatedcollapse.addDiv('DatosDG', 'fade=1')

    animatedcollapse.ontoggle = function (jQuery, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()

</script>
</head>
<script language="JavaScript">
/*
function tier1Menu(objMenu,objImage)
{
    if (objMenu.style.display == "none")
    {
        document.getElementById("DatosDG").style.display = "none";
        document.getElementById("img1").src = "../images/CarpetaCerrada.gif";
        document.getElementById("DatosDC").style.display = "none";
        document.getElementById("img2").src = "../images/CarpetaCerrada.gif";
        objMenu.style.display = "";
        objImage.src = "../Images/CarpetaAbierta.gif";
    }
    else
    {
        objMenu.style.display = "none";
	    objImage.src = "../images/CarpetaCerrada.gif";
    }
    Redimensionar();
}
*/
function Insertar() {
	if (document.tpvconf.i_tpv.value==""){
		window.alert("<%=LitNoTpv%>");
		return;
	}

	if (isNaN(document.tpvconf.i_tpv.value)){

		window.alert("<%=LitNoNumTpv%>");
		return;
	}

	if (isNaN(document.tpvconf.i_cota.value.replace(",","."))){

		window.alert("<%=LitNoNumCota%>");
		return;
	}

	if(document.tpvconf.i_caja.value==""){
			window.alert("<%=LitMsgCajaNoVal%>");
			return;
	}

	if (document.tpvconf.i_descripcion.value==""){
		window.alert("<%=LitNoDescr%>");
		return;
	}

    if (comp_car_ext(document.tpvconf.i_descripcion.value,0)==1){
		window.alert("<%=LitMsgTipoADesCarNoVal%>");
		return;
   }

   if (document.tpvconf.i_descripcion.value.length>=51){
   		window.alert("<%=LitMsgDesLongitudNoVal%>");
   		return;
   }

	//Recargar el submarco de detalles
	fr_Tabla.document.tpvconf_det.action="tpvconf_det.asp?mode=save&i_tpv= " + document.tpvconf.i_tpv.value + "&i_descripcion="+ document.tpvconf.i_descripcion.value +
	           "&i_caja=" + document.tpvconf.i_caja.value + "&i_cota=" + document.tpvconf.i_cota.value.replace(".",",") +
	           "&i_cajon=" + document.tpvconf.i_cajon.value + "&i_visor=" + document.tpvconf.i_visor.value + "&i_localTrace=" + document.tpvconf.i_localTrace.value +
               "&i_cloudTrace=" + document.tpvconf.i_cloudTrace.value;
	fr_Tabla.document.tpvconf_det.submit();
	//Limpiar los campos del formulario
	document.tpvconf.i_descripcion.value="";
	document.tpvconf.i_tpv.value="";
	document.tpvconf.i_caja.value="";
	document.tpvconf.i_cota.value = "1";
	document.tpvconf.i_localTrace.value = "";
	document.tpvconf.i_cloudTrace.value = "";
	//Colocar el foco en el campo de cantidad.
	document.tpvconf.i_tpv.focus();
}

function Mas(sentido,lote, texto) {
	document.getElementById("barras").style.display="none";
	fr_Tabla.document.tpvconf_det.action="tpvconf_det.asp?mode=ver&sentido=" + sentido + "&lote=" + lote + "&texto=" + texto;
	fr_Tabla.document.tpvconf_det.submit();
}
/*
if(window.document.addEventListener)
{
    window.document.addEventListener("keydown", callkeydownhandler, false);
}
else
{
    window.document.attachEvent("onkeydown", callkeydownhandler);
}

var ev = null;

function callkeydownhandler(evnt)
{
    ev = (evnt) ? evnt : event;
    keyPressed(ev);
}

//Comprueba si la tecla pulsada es CTRL+S. Si es as� guarda el registro.
function keyPressed(e)
{
    var keycode = e.keyCode;
	if (keycode==<%=TeclaGuardar%>) //CTRL+S
		Insertar();
}
*/
function Redimensionar()
{
    var alto = 0;
    if (parent.document.body.offsetHeight) alto = parent.document.body.offsetHeight;
    else alto = parent.self.innerHeight;
    if (document.getElementById("DatosDC")!=null){
	    if (document.getElementById("DatosDC").style.display=="")
	    {
            //window.alert("el alto es-" + alto + "-");
	        if (alto > 175)
            {
                if (alto - 375 > 175) document.getElementById("frtabla").style.height = alto - 375;
                else document.getElementById("frtabla").style.height = 175;
            }
            else document.getElementById("frtabla").style.height = 175;
        }
    }
}
</script>
<body bgcolor="<%=color_blau%>" onresize="javascript:Redimensionar();">

<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<%'Crea la tabla que contiene la barra de grupos de datos (Generales,Comerciales,etc)
sub BarraNavegacion(modo)    
        %>
        <script language="javascript" type="text/javascript">
            jQuery("#DatosDG").show();
            jQuery("#DatosDC").show();
        </script>
        <%
 end sub
 
            

'***********************************************************************************************************
' CODIGO PRINCIPAL DE LA PAGINA  ***************************************************************************
'***********************************************************************************************************

if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<form name="tpvconf" method="post" action="tpvconf.asp">
    <% set rst = server.CreateObject("ADODB.Recordset")
    set rstAux = server.CreateObject("ADODB.Recordset")

	PintarCabecera "tpvconf.asp"
	Alarma "tpvconf.asp"

	mode=request("mode")
    %><input type="hidden" name="mode_accesos_tienda" value="<%=enc.EncodeForHtmlAttribute(null_s(mode))%>" /><%
	viene = Request.QueryString("viene")

    if mode="select1" then mode="browse"
	
	if mode="save" then
	    ''ricardo 14-10-2008 el campo imprimirtickettpv se cogera de la tabla empresas en lugar de la tabla configuracion
	    StrSelConfTpv="update empresas with(rowlock) set imprimirtickettpv=" & nz_b(request.form("e_impticket")) & ", formatoticket="&request.form("e_fortic")&","
		StrSelConfTpv=StrSelConfTpv & " formatofacord='"&request.form("e_facord")&"', formatofacsimpl='"&request.form("e_facsimpl")
        StrSelConfTpv=StrSelConfTpv & "' from empresas "
        StrSelConfTpv=StrSelConfTpv & " where cif like '" & session("ncliente") & "%'"
        'response.Write(StrSelConfTpv)
        'response.End()
        rst.open StrSelConfTpv,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        if rst.State<>0 then rst.Close
        
        ''ricardo 14-10-2008 el campo imprimirtickettpv se cogera de la tabla empresas en lugar de la tabla configuracion	    
        ''StrSelConfTpv="select imprimirtickettpv,anularlineatpvaut,idiomatpv,sonidotpv,iptpv,porttpv,tiemporeintentotpv "
        StrSelConfTpv="select anularlineatpvaut,idiomatpv,sonidotpv,iptpv,porttpv,tiemporeintentotpv "
        StrSelConfTpv=StrSelConfTpv & " ,tpvtactil "
        'Toni Climent 14-01-2009 Agregamos el campo horecas a la consulta existente
        StrSelConfTpv=StrSelConfTpv & " ,horecas "
        'TCD 19-01-2009 Agregamos el campo pagosctatickets a la consulta existente
        StrSelConfTpv=StrSelConfTpv & " ,pagosctatickets "
        StrSelConfTpv=StrSelConfTpv & " from configuracion with(rowlock) "
        StrSelConfTpv=StrSelConfTpv & " where nempresa='" & session("ncliente") & "'"
        rst.open StrSelConfTpv,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
        if not rst.EOF then
            ''ricardo 14-10-2008 el campo imprimirtickettpv se cogera de la tabla empresas en lugar de la tabla configuracion	    
            ''rst("imprimirtickettpv")=nz_b(request.form("e_impticket"))
            rst("anularlineatpvaut")=nz_b(request.form("e_anilin"))
            rst("idiomatpv")=nulear(request.form("e_idioma"))
            rst("sonidotpv")=nz_b(request.form("e_sonido"))
            ''rst("iptpv")=nz_b(request.form("e_iptpv"))
            ''rst("porttpv")=nz_b(request.form("e_porttpv"))
            rst("tiemporeintentotpv")=nz_b(request.form("xxxxxx"))
            rst("tpvtactil")=nz_b(request.form("e_tactil"))
            ''Toni Climent 14-01-2009 el campo horecas se obtiene de la tabla configuracion
            rst("horecas")=nz_b(request.form("e_horecas"))
             ''TCD 19-01-2009 el campo pagosctatickets se obtiene de la tabla configuracion
            rst("pagosctatickets")=nz_b(request.form("e_pagosctatickets"))
            rst.Update
        end if
        rst.Close
            
	    mode="browse"
	    
	    if viene="asistente" then%>
	        <script language="javascript">parent.botones.document.location="tpvconf_bt.asp?mode=browse&viene=asistente";</script>    
	    <%else %>
	        <script language="javascript">parent.botones.document.location="tpvconf_bt.asp?mode=browse";</script>
	    <%end if%>
	    
    <%end if%>
	<table class="CELDA6" HEIGHT=1 width="1"><tr><td class="CELDA6" HEIGHT=1 width="1"><td><tr></table>
	<%BarraOpciones%>
	<table class="CELDA6" HEIGHT=1 width="1"><tr><td class="CELDA6" HEIGHT=1 width="1"><td><tr></table>
	<%BarraNavegacion mode%>
	<!--<table class=TBORDE width="100%"><tr><td>-->
	<!--<table class="CELDA6" HEIGHT=1 width="1"><tr><td class="CELDA6" HEIGHT=1 width="1"><td><tr></table>-->
	<%if mode="edit" then
	    mostrarDG=""
	    mostrarDC="none"
	else
	    mostrarDG="none"
	    mostrarDC=""
	end if

    ''ricardo 1-8-2008 se a�ade campos generales del tpv	
    if mode<>"edit" then
	%>
        <div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['DatosDG','DatosDC']); hideNoCollapse();"><img Class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['DatosDG','DatosDC']);hideCollapse();" style="display:none"><img Class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
        </div>
    <%else %>
        <div id="CollapseSection"> 
            <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['DatosDG']); hideNoCollapse();"><img Class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
            <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['DatosDG']);hideCollapse();" style="display:none"><img Class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
        </div>
    <%end if %>

    <div class="Section" id="S_DatosDG">
        <a href="#" rel="toggle[DatosDG]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
            <div class="SectionHeader">
                <%=LitDatosGeneralesConfTPV%>
                <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
            </div>
        </a>
    <div class="CELDA" style="display:<%=mostrarDG%>" id="DatosDG">
	<!--<br />-->
    <!--<center>-->

	        <%''ricardo 14-10-2008 el campo imprimirtickettpv se cogera de la tabla empresas en lugar de la tabla configuracion
            StrSelConfTpv="select imprimirtickettpv, formatoticket, formatofacord, formatofacsimpl "
            StrSelConfTpv=StrSelConfTpv & " from empresas with(NOLOCK) "
            StrSelConfTpv=StrSelConfTpv & " where cif like '" & session("ncliente") & "%'"
            rst.CursorLocation=3
            rst.open StrSelConfTpv,session("dsn_cliente")
            if not rst.eof then
                valor_imprimirtickettpv=nz_b(rst("imprimirtickettpv"))
                valor_formatoticket=rst("formatoticket")
				valor_formatofacord=rst("formatofacord")
				valor_formatofacsimpl=rst("formatofacsimpl")
            else
                valor_imprimirtickettpv=nz_b(rst("imprimirtickettpv"))
                valor_formatoticket=rst("formatoticket")
				valor_formatofacord=rst("formatofacord")
				valor_formatofacsimpl=rst("formatofacsimpl")
            end if
            rst.Close
            	        
            ''ricardo 14-10-2008 el campo imprimirtickettpv se cogera de la tabla empresas en lugar de la tabla configuracion
            ''StrSelConfTpv="select imprimirtickettpv,anularlineatpvaut,idiomatpv,sonidotpv,iptpv,porttpv,tiemporeintentotpv "
            StrSelConfTpv="select anularlineatpvaut,idiomatpv,sonidotpv,iptpv,porttpv,tiemporeintentotpv "
            StrSelConfTpv=StrSelConfTpv & " ,tpvtactil "
            'Toni Climent 14-01-2009 Agregamos el campo horecas a la consulta existente
            StrSelConfTpv=StrSelConfTpv & " ,horecas "
            'TCD 19-01-2009 Agregamos el campo pagosctatickets a la consulta existente
            StrSelConfTpv=StrSelConfTpv & " ,pagosctatickets "
            StrSelConfTpv=StrSelConfTpv & " from configuracion with(NOLOCK) "
            StrSelConfTpv=StrSelConfTpv & " where nempresa='" & session("ncliente") & "'"
            rst.CursorLocation=3
            rst.open StrSelConfTpv,session("dsn_cliente")
            if not rst.eof then        	     
              
                if  mode="edit" then                
                   
	    	        EligeCelda "check","edit","","","",0,LitTPVTactil,"e_tactil",8,iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("tpvtactil")))
	    	        EligeCelda "check","edit","","","",0,LitTPVSoni,"e_sonido",0,iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("sonidotpv")))
                else
                    'ºCeldaResponsive "","","","",LitTPVTactil,Visualizar(iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("tpvtactil"))))            
              
                    EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVTactil,"", Visualizar(iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("tpvtactil"))))
                    'DrawCeldaResponsive "","","","",LitTPVSoni,Visualizar(iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("sonidotpv")))) 
                    EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVSoni,"", Visualizar(iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("sonidotpv")))) 
               
                end if
	    	
        		    if mode="edit" then
                    DrawDiv "1", "", ""
                    DrawLabel "", "", LitTPVIdio%><select class="width60" name="e_idioma">
					            <option <%=iif(rst("idiomatpv")="ESP","selected","")%> value="ESP"><%=LitIdiomasConfTPV_Esp%></option>
					            <option <%=iif(rst("idiomatpv")="POR","selected","")%> value="POR"><%=LitIdiomasConfTPV_Por%></option>
				            </select>			                    
			        <%CloseDiv
                      else
			            descidioma=""
			            if rst("idiomatpv")="ESP" then
			                descidioma=LitIdiomasConfTPV_Esp
			            end if
			            if rst("idiomatpv")="POR" then
			                descidioma=LitIdiomasConfTPV_Por
			            end if	    	                              
                         EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVIdio,"", descidioma
	    	        end if
	    	        ''ricardo 14-10-2008 el campo imprimirtickettpv se cogera de la tabla empresas en lugar de la tabla configuracion
	    	        ''EligeCelda "check", mode,"CELDA","0","",0,LitTPVImpTic,"e_impticket",0,iif(e_impticket>"",nz_b(e_impticket),nz_b(rst("imprimirtickettpv")))
                     if mode="edit" then
	    	            EligeCelda "check","edit","CELDA","0","",0,LitTPVImpTic,"e_impticket",0,iif(e_impticket>"",nz_b(e_impticket),valor_imprimirtickettpv)
                     else                      
                         EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVImpTic,"", Visualizar(iif(e_impticket>"",nz_b(e_impticket),valor_imprimirtickettpv))
                     end if
	    	 
                     if mode="edit" then
	    	            EligeCelda "check", "edit","","0","",0,LitTPVAnulLin,"e_anilin",0,iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("anularlineatpvaut")))
                     else                      
                         EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVAnulLin,"", Visualizar(iif(e_tactil>"",nz_b(e_tactil),nz_b(rst("anularlineatpvaut"))))
                     end if
	    	        ''Toni Climent 14-01-2008 Substituimos la celda vacia por el control para dar valor al campo HORECAS
                    if mode="edit" then
	    	            EligeCelda "check", mode,"CELDA","0","",0,LitTPVHorecas,"e_horecas",0,iif(e_horecas>"",nz_b(e_horecas),nz_b(rst("horecas")))
                    else                       
                        EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVHorecas,"", Visualizar(iif(e_horecas>"",nz_b(e_horecas),nz_b(rst("horecas"))))
                    end if	    	      
	    	    ''TCD 19-01-2008 A�adimos una celda vacia y la celda para dar valor al campo PAGOSCTATICKETS
	    	 
	    	        ''MPC 24/02/2009 Se a�ade el campo formato ticket para el TPV
	    	        if mode="edit" then
	    	            rstAux.Open "select descripcion, personalizacion from clientes_formatos_imp cfi with(nolock), formatos_imp fi with(nolock) where cfi.ncliente='"&session("ncliente")&"' and fi.nformato=cfi.nformato and fi.tippdoc like 'TICKET'",DsnIlion,adUseClient, adLockReadOnly%>
					
                             <%DrawDiv "1", "", ""
                               DrawLabel "", "", LitTPVForTic%><select CLASS="width60" name='e_fortic'>                                                                         
	    	                        <%while not rstAux.EOF
		                                if cstr(valor_formatoticket)=cstr(rstAux("personalizacion")) then
			                                response.write("<option selected value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion"))) & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
		                                else
			                                response.write("<option value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion"))) & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
		                                end if
		                                rstAux.Movenext                                                      
	                                wend%>
							    </select>
	    	            <%
	    	            rstAux.Close
                        CloseDiv
	    	        else	    	                                  
                        EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVForTic,"", enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion", "clientes_formatos_imp cfi with(nolock), formatos_imp fi", "cfi.ncliente='"&session("ncliente")&"' and fi.nformato=cfi.nformato and fi.tippdoc like 'TICKET' and personalizacion='"&valor_formatoticket&"'", DsnIlion)))
	    	        end if
	    	        ''FIN MPC
                     if mode="edit" then
                        EligeCelda "check", mode,"CELDA","0","",0,LitTPVPagosctatickets,"e_pagosctatickets",0,iif(e_pagosctatickets>"",nz_b(e_pagosctatickets),nz_b(rst("pagosctatickets")))
                    else                        
                         EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVPagosctatickets,"", Visualizar(iif(e_pagosctatickets>"",nz_b(e_pagosctatickets),nz_b(rst("pagosctatickets"))))
                    end if	    	        	  	    	        
				
				if mode="edit" then
					rstAux.Open "select descripcion, personalizacion from clientes_formatos_imp cfi with(nolock), formatos_imp fi with(nolock) where cfi.ncliente='"&session("ncliente")&"' and fi.nformato=cfi.nformato and fi.tippdoc like 'FACTURA A CLIENTE'",DsnIlion,adUseClient, adLockReadOnly%>
						
                             <%DrawDiv "1", "", ""
                               DrawLabel "", "", LitTPVForFacSimpl%><select class="width60" name='e_facsimpl'>
								    <%while not rstAux.EOF
									    if rstAux("personalizacion") <> "" AND valor_formatofacsimpl & "" <> "" then
										    if cstr(valor_formatofacsimpl)=cstr(rstAux("personalizacion")) then                                                                
											    response.write("<option selected value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion")))   & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
										    else
											    response.write("<option value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion")))   & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
										    end if
									    else
										    response.write("<option value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion")))   & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
									    end if
									    rstAux.Movenext
								    wend%></select><%
					rstAux.Close
                    CloseDiv
				else					                  
                    EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVForFacSimpl,"", enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion", "clientes_formatos_imp cfi with(nolock), formatos_imp fi", "cfi.ncliente='"&session("ncliente")&"' and fi.nformato=cfi.nformato and fi.tippdoc like 'FACTURA A CLIENTE' and personalizacion='"&valor_formatofacsimpl&"'", DsnIlion)))
				end if
				
				if mode="edit" then
					rstAux.Open "select descripcion, personalizacion from clientes_formatos_imp cfi with(nolock), formatos_imp fi with(nolock) where cfi.ncliente='"&session("ncliente")&"' and fi.nformato=cfi.nformato and fi.tippdoc like 'FACTURA A CLIENTE'",DsnIlion,adUseClient, adLockReadOnly%>
					
                          <%DrawDiv "1", "", ""
                            DrawLabel "", "", LitTPVForFacOrd%><select class="width60" name='e_facord'>
								<%while not rstAux.EOF
									if rstAux("personalizacion") <> "" AND valor_formatofacord & "" <> "" then
										if cstr(valor_formatofacord)=cstr(rstAux("personalizacion")) then
											    response.write("<option selected value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion")))   & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
										    else
											    response.write("<option value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion")))   & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
										    end if
									    else
										    response.write("<option value='" & enc.EncodeForHtmlAttribute(null_s(rstAux("personalizacion")))   & "'>" & enc.EncodeForHtmlAttribute(null_s(rstAux("descripcion"))) & "</option>")
									end if
									rstAux.Movenext
								wend%></select><%
					rstAux.Close
                    CloseDiv
				else					
                     EligeCeldaResponsive "text",mode,"CELDA","","",0,"",LitTPVForFacOrd,"", enc.EncodeForHtmlAttribute(null_s(d_lookup("descripcion", "clientes_formatos_imp cfi with(nolock), formatos_imp fi", "cfi.ncliente='"&session("ncliente")&"' and fi.nformato=cfi.nformato and fi.tippdoc like 'FACTURA A CLIENTE' and personalizacion='"&valor_formatofacord&"'", DsnIlion)))
				end if	
				
		    end if
		    rst.Close%>	
   
    </div>
    </div>
    <%if mode<>"edit" then%>
        <div class="Section" id="S_DatosDC">
            <a href="#" rel="toggle[DatosDC]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitTPVS%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" style="display:<%=mostrarDC%>" id="DatosDC">	  
                <div id="tablas" class="overflowXauto">                    
		        <!--<table width="<%=TamTabla%>" style="width:<%=TamTabla%>px;" border="0" cellspacing="1" cellpadding="1">-->                    
                    <table class="width90 lg-table-responsive bCollapse"><%
		            'Drawfila color_fondo
			            'Drawcelda2 "CELDA style='width:" & TamCelda1 & "px;'", "left", false, LitTpv
			            'Drawcelda2 "CELDA style='width:" & TamCelda2 & "px;'", "left", false, LitCaja
			            'Drawcelda2 "CELDA style='width:" & TamCelda3 & "px;'", "left", false, LitDescripcion
			            'Drawcelda2 "CELDA style='width:" & TamCelda6 & "px;'", "left", false, LitCajConfTPV
			            'Drawcelda2 "CELDA style='width:" & TamCelda7 & "px;'", "left", false, LitVisConfTPV
			            ''Drawcelda2 "CELDA style='width:" & TamCelda4 & "px;'", "left", false, LitCota
                        'Drawcelda2 "CELDA style='width:" & TamCelda4 & "px;'", "left", false, ""
			            'Drawcelda2 "CELDA style='width:" & TamCelda5 & "px;'", "center", false, LitEstado
                        'Drawcelda2 "CELDA style='width:" & TamCelda9 & "px;'", "left", false, LitLocalTrace
                        'Drawcelda2 "CELDA style='width:" & TamCelda10 & "px;'", "left", false, LitCloudTrace
			            'Drawcelda2 "CELDA style='width:" & TamCelda8 & "px;'", "left", false, LitPinPadConfTPV
			            'Drawcelda2 "CELDA style='width:" & TamCeldaBotones & "px;'", "center", false, ""
		            'closefila
                    %><tr>
                        <td class='underOrange ENCABEZADOL width5'><%=LitTpv%></td>
                        <td class='underOrange ENCABEZADOL width10'><%=LitCaja%></td>
                        <td class='underOrange ENCABEZADOL width10'><%=LitDescripcion%></td>
                        <td class='underOrange ENCABEZADOL width10'><%=LitCajConfTPV%></td>
                        <td class='underOrange ENCABEZADOL width10'><%=LitVisConfTPV%></td>
                        <td class='underOrange ENCABEZADOL width10'></td>
                        <td class='underOrange ENCABEZADOL width10'><%=LitEstado%></td>
                        <td class='underOrange ENCABEZADOL width10'><%=LitLocalTrace%></td>
                        <td class='underOrange ENCABEZADOL width10'><%=LitCloudTrace%></td>
                        <td class='underOrange ENCABEZADOL width5'><%=LitPinPadConfTPV %></td>
                        <td class='underOrange ENCABEZADOL width5'></td>

                      </tr>

		            <%'Drawfila color_blau
			            %>
                        <tr>
                            <!--<td class="ENCABEZADOL" width="<%=TamCelda1%>"><input class="CELDA" type="text" style="width: <%=TamCelda1%>px;" maxlength="3" name="i_tpv"/></td>-->
                             <td class='ENCABEZADOL underOrange width5'><input type="text" class="width100" maxlength="3" name="i_tpv"/></td>
			               <!-- <td class="ENCABEZADOL" width="<%=TamCelda2%>"><select class="CELDA" name="i_caja" style="width: <%=TamCelda2%>px;">-->
                            <td class="ENCABEZADOL underOrange width10"><select name="i_caja" class="width100"><%
			                rst.open "select a.codigo,a.descripcion from cajas a with(nolock), tiendas b with(nolock), almacenes c with(nolock), series d with(nolock) where a.tienda=b.codigo and b.almacen=c.codigo and d.nserie=a.serie and a.codigo like '" & session("ncliente") & "%' order by a.descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			                while not rst.eof
				                %><option value="<%=enc.EncodeForHtmlAttribute(null_s(rst("codigo")))%>"><%=enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))%></option><%
				                rst.movenext                                                                                      
			                wend                                                                            
			                rst.close
			                %></select></td>
			                <td class="ENCABEZADOL underOrange width10"><input type="text" class="width100" maxlength="50" name="i_descripcion"/></td>
			
			                <!--<td class="ENCABEZADOL" width="<%=TamCelda6%>"><select class="CELDA" name="i_cajon" style="width: <%=TamCelda6%>px;">-->
                            <td class="ENCABEZADOL underOrange width10"><select name="i_cajon" class="width100"><%
			                rst.open "select null as codigo,'' as descripcion union all select c.codigo,c.nombre as descripcion from cajmonedero as c with(nolock) where c.codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			                while not rst.eof
				                %><option value="<%=enc.EncodeForHtmlAttribute(null_s(rst("codigo")))%>"><%=enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))%></option><%
				                rst.movenext
			                wend
			                rst.close
			                %></select></td>
			                <!--<td class="ENCABEZADOL" width="<%=TamCelda7%>"><select class="CELDA" name="i_visor" style="width: <%=TamCelda7%>px;">-->
                            <td class="ENCABEZADOL underOrange width10"><select class="width100" name="i_visor"><%
			                rst.open "select null as codigo,'' as descripcion union all select v.codigo,v.nombre as descripcion from visores as v with(nolock) where v.codigo like '" & session("ncliente") & "%' order by descripcion",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
			                while not rst.eof
				                %><option value="<%=enc.EncodeForHtmlAttribute(null_s(rst("codigo")))%>"><%=enc.EncodeForHtmlAttribute(null_s(rst("descripcion")))%></option><%
				                rst.movenext
			                wend
			                rst.close
			                %></select></td>

			                <!--<td class="ENCABEZADOL" width="<%=TamCelda4%>"><input class="CELDA" type="text" style="width: <%=TamCelda4%>px; display: none" maxlength="10" name="i_cota" value="1"/></td>-->
                            <td class="ENCABEZADOL underOrange width10"><input class="width100" type="text" style="display: none" maxlength="10" name="i_cota" value="1"/></td>
			                <!--<td class="ENCABEZADOL" width="<%=TamCelda5%>">&nbsp;</td>-->
                            <td class="ENCABEZADOL underOrange width10"></td>
                            <!--<td class="ENCABEZADOL" width="<%=TamCelda9%>"><select class="CELDA" name="i_localTrace" style="width: <%=TamCelda9%>px;">-->
                            <td class="ENCABEZADOL underOrange width10"><select class="width100" name="i_localTrace">
                                <option value=""></option>
				                <option value="ERROR">ERROR</option>
                                <option value="WARNING">WARNING</option>
                                <option value="DEBUG">DEBUG</option>
                                <option value="ALLINFO">ALLINFO</option>
			                </select></td>

                             <!--<td class="ENCABEZADOL" width="<%=TamCelda10%>"><select class="CELDA" name="i_cloudTrace" style="width: <%=TamCelda10%>px;">-->
                            <td class="ENCABEZADOL underOrange width10"><select class="width100" name="i_cloudTrace">
                                <option value=""></option>
				                <option value="ERROR">ERROR</option>
                                <option value="WARNING">WARNING</option>
                                <option value="DEBUG">DEBUG</option>
                                <option value="ALLINFO">ALLINFO</option>    
			                </select></td>
			                <!--<td class="ENCABEZADOL" width="<%=TamCelda8%>">&nbsp;</td>-->
                            <td class="ENCABEZADOL underOrange width5"></td>
			                <!--<td width="<%=TamCeldaBotones%>"><a id="iddetinssave" href="javascript:Insertar();" ><img src="../images/<%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a>-->
                            <td class="ENCABEZADOL underOrange width5"><a class='ic-accept noMTop' id="iddetinssave" href="javascript:Insertar();"><img src="<%=themeIlion %><%=ImgNuevo%>" <%=ParamImgNuevo%> alt="<%=LitNuevo%>" title="<%=LitNuevo%>"/></a></td>
                        </tr>
			        <%'closefila
		        %></table>                
               <!-- <iframe id="frtabla" name="fr_Tabla" src='tpvconf_det.asp?mode=browse' width="<%=TamTabla+TamCeldaScroll%>" height='250' frameborder="yes" noresize="noresize"></iframe>-->
                <iframe id="frtabla" name="fr_Tabla" src='tpvconf_det.asp?mode=browse' class="width90 iframe-data lg-table-responsive"  height='250' frameborder="yes" noresize="noresize"></iframe>
                <script type="text/javascript" language="javascript">Redimensionar();</script>
                </div>
                <table class="width100 bCollapse">
                    <%'DrawFila ""%>
                        <!--<td class=CELDA7 width="250">-->
				            <span id="barras" STYLE="display:none">
				            </span>
			            <!--</td>-->
		            <%'CloseFila%>
                </table>
               
        </div>
        </div>
    <%end if%>
    <% if mode<>"edit" then%>
    		<script type="text/javascript" language="javascript">
			document.tpvconf.i_tpv.focus();
		</script>
    <%end if%>
    
</form>
<%set rst = nothing
set rstAux = nothing
end if%>
</body>
</html>