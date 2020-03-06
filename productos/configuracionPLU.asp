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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html LANG="<%=session("lenguaje")%>">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>


<!--#include file="../ilion.inc" -->
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="configuracionPLU.inc" -->
<!--#include file="../tablasResponsive.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" --> 

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<!--<script language="javascript" src="../CMS/Gestor/SelectColor.js"></script>-->

<!--#include file="../js/generic.js.inc"-->
<!--#include file="../js/animatedCollapse.js.inc" -->
<!--#include file="../styles/Section.css.inc" -->
<!--#include file="../styles/font-face.css.inc" -->
<!--#include file="../styles/generalData.css.inc" -->
<!--#include file="../js/tabs.js.inc" -->
<!--#include file="../styles/Tabs.css.inc" -->
<!--#include file="../styles/ExtraLink.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

</head>
<script language="javascript" type="text/javascript">
    animatedcollapse.addDiv('DatosDGRP', 'fade=1')
    animatedcollapse.addDiv('DatosDPLU', 'fade=1')

    animatedcollapse.ontoggle = function ($, divobj, state) { //fires each time a DIV is expanded/contracted
        //$: Access to jQuery
        //divobj: DOM reference to DIV being expanded/ collapsed. Use "divobj.id" to get its ID
        //state: "block" or "none", depending on state
    }

    animatedcollapse.init()

function MasGrp(sentido,lote,campo,criterio,texto) {
	if (fr_GRP_det.document.getElementById("waitBoxOculto")=="[object]")
		fr_GRP_det.document.getElementById("waitBoxOculto").style.visibility="visible";
	if (texto=="undefined") texto="";
	fr_GRP_det.document.configuracionPLU_grp_det.action="configuracionPLU_grp_det.asp?mode=browse&sentido=" + sentido + "&lote=" + lote + "&campo=" + campo + "&criterio=" + criterio + "&texto=" + texto;
	fr_GRP_det.document.configuracionPLU_grp_det.submit();
}

function reloadIframe(id) {
    var iframe = document.getElementById(id);
    iframe.src = iframe.src;
}
    
function AnadirGrp(){
	if(document.getElementById("i_descripcion").value==""){
		window.alert("<%=LitDescGrpNul%>");
		return;
	}

	if(document.getElementById("i_descripcion").value.length>=31){
		window.alert("<%=LitDescLongGrpNul%>");
		return;
	}

	if(document.getElementById("i_nplu").value==""){
		window.alert("<%=LitNPluGrpNul%>");
		return;
	}

	if(isNaN(document.getElementById("i_nplu").value.replace(",","."))){
		window.alert("<%=LitNPluGrpNum%>");
		return;
	}

	if(parseInt(document.getElementById("i_nplu").value.replace(",","."))>255){
		window.alert("<%=LitNPluLongGrpNul%>");
		return;
	}

	if(document.getElementById("i_descripBotTPV").value==""){
		window.alert("<%=LitDescTpvGrpNul%>");
		return;
	}

	if(document.getElementById("i_descripBotTPV").value.length>=11){
		window.alert("<%=LitDescTpvLongGrpNul%>");
		return;
	}

	if(document.getElementById("i_color").value==""){
		window.alert("<%=LitColGrpNul%>");
		return;
	}

	//Recargar el submarco de detalles
	colorI=document.getElementById("i_color").value.substring(1,document.getElementById("i_color").value.length);
	
	datos_pagina="configuracionPLU_grp_det.asp?mode=save&i_grp=&i_descripcion="+ document.getElementById("i_descripcion").value +
	"&i_nplu=" + document.getElementById("i_nplu").value + "&i_descripBotTPV=" + document.getElementById("i_descripBotTPV").value.replace(".",",") +
	"&i_color=" + colorI;
	document.getElementById("i_descripcion").value="";
	document.getElementById("i_nplu").value="";
	document.getElementById("i_descripBotTPV").value="";
	document.getElementById("i_color").value="";
	
	fr_GRP_det.document.configuracionPLU_grp_det.action=datos_pagina;
	fr_GRP_det.document.configuracionPLU_grp_det.submit();
    document.getElementById("i_descripcion").focus();

    reloadIframe('frPLU');
}

//function CrearCapaColoresGRP()
//{
//    var _o='';
//    _o+='<span id="capaColores" style="display:none">';
//    //_o+='<span id="capaColores" >'
//    _o+='<table border="1" bgcolor="#888888" cellpadding="0" cellspacing="0">';
//    _o+='<!-- RED --><tr>';
//    var _a$=new Array("#ffeeee","#ffdddd","#ffcccc","#ffbbbb","#ffaaaa","#ff9999","#ff8888","#ff7777","#ff6666","#ff5555","#ff4444","#ff3333","#ff2222","#ff1111","#ff0000","#ee0000","#dd0000","#cc0000","#bb0000","#aa0000","#990000","#880000","#770000","#660000","#550000","#440000","#330000","#220000","#110000");
//    _o+=_fGRP(_a$);
//    _o+='</tr><!-- GREEN --><tr>';
//    var _b$=new Array("#eeffee","#ddffdd","#ccffcc","#bbffbb","#aaffaa","#99ff99","#88ff88","#77ff77","#66ff66","55ff55","#44ff44","#33ff33","#22ff22","#11ff11","#00ff00","#00ee00","#00dd00","#00cc00","#00bb00","#00aa00","#009900","#008800","#007700","#006600","#005500","#004400","#003300","#002200","#001100");
//    _o+=_fGRP(_b$);
//    _o+='</tr><!-- BLUE --><tr>';
//    var _c$=new Array("#eeeeff","#ddddff","#ccccff","#bbbbff","#aaaaff","#9999ff","#8888ff","#7777ff","#6666ff","#5555ff","#4444ff","#3333ff","#2222ff","#1111ff","#0000ff","#0000ee","#0000dd","#0000cc","#0000bb","#0000aa","#000099","#000088","#000077","#000066","#000055","#000044","#000033","#000022","#000011");
//    _o+=_fGRP(_c$);
//    _o+='</tr><!-- YELLOW --><tr>';
//    var _d$=new Array("#ffffee","#ffffdd","#ffffcc","#ffffbb","#ffffaa","#ffff99","#ffff88","#ffff77","#ffff66","#ffff55","#ffff44","#ffff33","#ffff22","#ffff11","#ffff00","#eeee00","#dddd00","#cccc00","#bbbb00","#aaaa00","#999900","#888800","#777700","#666600","#555500","#444400","#333300","#222200","#111100");
//    _o+=_fGRP(_d$);
//    _o+='</tr><!-- PURPLE --><tr>';
//    var _e$=new Array("#ffeeff","#ffddff","#ffccff","#ffbbff","#ffaaff","#ff99ff","#ff88ff","#ff77ff","#ff66ff","#ff55ff","#ff44ff","#ff33ff","#ff22ff","#ff11ff","#ff00ff","#ee00ee","#dd00dd","#cc00cc","#bb00bb","#aa00aa","#990099","#880088","#770077","#660066","#550055","#440044","#330033","#220022","#110011");
//    _o+=_fGRP(_e$);
//    _o+='</tr><!-- ORANGE --><tr>';
//    var _f$=new Array("#ffdddd","#ffeeaa","#ffee99","#ffdd88","#ffcc77","#ffcc66","#ffbb66","#ffaa55","#ffaa44","#ff9944","#ff8833","#ff8822","#ff7722","#ff6622","#ff6611","#ee5522","#ee5511","#dd4400","#cc3300","#bb2200","#aa2200","#992200","#882200","#772200","#662200","#552200","#442200","#332200","#222200");
//    _o+=_fGRP(_f$);
//    _o+='</tr><!-- CYAN --><tr>';
//    var _g$=new Array("#eeffff","#ddffff","#ccffff","#bbffff","#aaffff","#99ffff","#88ffff","#77ffff","#66ffff","#55ffff","#44ffff","#33ffff","#22ffff","#11ffff","#00ffff","#00eeee","#00dddd","#00cccc","#00bbbb","#00aaaa","#009999","#008888","#007777","#006666","#005555","#004444","#003333","#002222","#001111");
//    _o+=_fGRP(_g$);
//    _o+='</tr><!-- GRAY --><tr>';
//    var _h$=new Array("#ffffff","#eeeeee","#dddddd","#d0d0d0","#cccccc","#c0c0c0","#bbbbbb","#b0b0b0","#aaaaaa","#a0a0a0","#999999","#909090","#888888","#808080","#777777","#707070","#666666","#606060","#555555","#505050","#444444","#404040","#333333","#303030","#222222","#202020","#111111","#101010","#000000");
//    _o+=_fGRP(_h$);
//    _o+='</tr>';
//    _o+='<tr><td bgcolor="#ffffff" colspan="29" align="center"><input type="button" value="Cerrar" onclick="document.getElementById(\'capaColores\').style.display=\'none\';"></td></tr>';
//    _o+='</table></span>';
//    document.write(_o);
//}


//}
//function NuevaCapaColorGrp(NombreCampo)
//{
//    // Recoje el nombre el campo para luego asignarle el color
//    CampoColorSelect=NombreCampo;

//    // Mostrar la capa de los colores
//    document.getElementById("capaColores").style.display='inline';
//}

//function doItGrp(color)
//{
//    // Oculta la capa
//    document.getElementById("capaColores").style.display='none';
//    // Asigna el color al campo que ha solicitado el color
//    eval('document.getElementById("' + CampoColorSelect + '").value=\'' + color + '\'');
//    actualizaCampoColoresGRP();
//    /*
//    var fireOnThis = document.getElementById(CampoColorSelect);
//    if(document.createEvent)
//    {
//        var evObj = document.createEvent('HTMLEvents');
//        evObj.initEvent('change', true, false );
//        fireOnThis.dispatchEvent(evObj);
//    }
//    else if( document.createEventObject )
//    {
//        eval('document.getElementById("' + CampoColorSelect + '").fireEvent("onchange")');
//    }
//    */
//}

//function _fGRP(_a)
//{
//    var _o='';
//    for(var i=0;i<_a.length;i++)
//    {
//        _o += '<td style="font-size:2pt;font-family:Tahoma;width:14px;height:14px;background-color:' + _a[i] + '">';
//        _o += '<a href="#" onclick="doItGrp(\'' + _a[i] + '\');">';
//        _o += '<img src="" style="opacity:0;filter:alpha(opacity=60);width:14px;height:14px;" border="0" alt="' + _a[i] + '" title="' + _a[i] + '"/></a></td>';
//    }
//    return(_o);
//}

////////////////////////////////////////////////////////////
function updateColpick()
{
    $('#Box').colpick({
            colorScheme: 'dark',
            layout: 'hex',
            onSubmit: function (hsb, hex, rgb, el)
            {
                $(el).css('background-color', '#' + hex);
                //var colorJS = "#" + hex;
                $(el).attr("value", '#' + hex);
                $('#i_color').val('#' + hex.toUpperCase());
                $('#previous_color').val('#' + hex.toUpperCase());
                $(el).colpickHide();
            }
    });
}

function actualizaCampoColoresGRP()
{
    var previousColor = $("#previous_color").val();
    var isOk = /(^#[0-9A-F]{6}$)|(^#[0-9A-F]{3}$)/i.test($("#i_color").val());
    if (isOk) {
        document.getElementById("Box").style.backgroundColor = $("#i_color").val();
        $('#Box').attr("value", $("#i_color").val());
        $("#previous_color").val($("#i_color").val());
    } else {
        $("#i_color").val(previousColor);
        alert("<%=LitColorIncorr%>");
    }
}

function VistaPrevia()
{
	pagina="configuracionPLU_grp.asp?mode=browse";
	AbrirVentana(pagina,'P',195,500);
}

function Redimensionar()
{
    var alto = 0;
    if (parent.document.body.offsetHeight) alto = parent.document.body.offsetHeight;
    else alto = parent.self.innerHeight;
	if (document.getElementById("DatosDGRP").style.display=="")
	{
	    if (alto > 150)
        {
            if (alto - 310 > 150) document.getElementById("frGRP_det").style.height = alto - 310;
            else document.getElementById("frGRP_det").style.height = 150;
        }
        else document.getElementById("frGRP_det").style.height = 150;
    }
    else
    {
        if (document.getElementById("DatosDPLU").style.display=="")
	    {
	        if (alto > 150)
            {
                if (alto - 250 > 150) document.getElementById("frPLU").style.height = alto - 250;
                else document.getElementById("frPLU").style.height = 150;
            }
            else document.getElementById("frPLU").style.height = 150;
        }
    }
}
/*
function tier1Menu(objMenu,objImage)
{
    if (objMenu.style.display == "none")
    {
        document.getElementById("DatosDGRP").style.display="none";
        document.getElementById("img1").src="../Images/CarpetaCerrada.gif";
        document.getElementById("DatosDPLU").style.display="none";
        document.getElementById("img2").src="../Images/CarpetaCerrada.gif";
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
</script>
<body bgcolor="<%=color_blau%>" onresize="javascript:Redimensionar();">
<%'******************************************************************************
sub BarraNavegacion(modo)
                %>
                <script language="javascript" type="text/javascript">
                    $("#S_DatosDGRP").show();
                    $("#S_DatosDPLU").hide();
                </script>
                <%

end sub

'******************************************************************************
'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0

if accesoPagina(session.sessionid,session("usuario"))=1 then
    indice=limpiaCadena(Request.Querystring("indice"))
    if indice & ""="" then indice=1
    
    parametros=limpiaCadena(Request.Querystring("parametros"))
    if parametros & ""="" then parametros="none"
        
    narticulo=limpiaCadena(Request.Querystring("narticulo"))
    nomarticulo =limpiaCadena(Request.Querystring("nomarticulo"))
    
    nnumero=limpiaCadena(Request.Querystring("nnumero"))
    ndescripcion=limpiaCadena(Request.Querystring("ndescripcion"))
    ncolor=limpiaCadena(Request.Querystring("ncolor"))
        
    'Leer parámetros de la página
	mode=EncodeForHtml(request("mode"))

    if mode="select1" then mode="browse"
    
    PintarCabecera "configuracionPLU.asp"
    Alarma "configuracionPLU.asp"%>
    <style>
        .color-box{
            position: inherit;
            top: 0;
            left: 50%;
            width: 36px;
            height: 36px;
            /*margin-left: 15px;
            padding-top: 15px;*/
        }
        #Box{
            /*position: inherit;
            top: 0;
            left: 0;*/
            width: 36px;
            height: 36px;
            background: url(/Lib/estilos/hubble/images/select2.png) center/126% auto;
        }
    </style>
    <script type="text/javascript" src="/Lib/jQuery/jquery-1.11.1.min.js"></script>
    <script type="text/javascript" src="/Lib/jQuery/jquery-ui-1.10.2.custom.min.js"></script>
    <script type="text/javascript" src="/Lib/jQuery/colpickConfigPLU.js"></script>
    <link rel="stylesheet" type="text/css" href="/Lib/estilos/<%=Session("folder")%>/colpick.css"/>
    <table class="CELDA6" height="1" width="1"><tr><td class="CELDA6" height=1 width="1"></td></tr></table>
    
    <div class="headers-wrapper">
        <%DrawDiv "3", "", ""
            %><label class="width100"><%=LitEnlaceImg1%><a class="CELDAREF" href="../../controles/<%=LitNomFich%>"><%=LitEnlaceImg2%></a><%=LitEnlaceImg3%></label><%
        CloseDiv%>
    </div>
        
    <table style="width:100%"></table>
	<table class="CELDA6" height="1" width="1"><tr><td class="CELDA6" height=1 width="1"></td></tr></table>
	<%BarraNavegacion mode%>
	<!--<table class=TBORDE width="100%"><tr><td>
	<table class="CELDA6" height=1 width="1"><tr><td class="CELDA6" height=1 width="1"><td><tr></table>-->
	<%mostrarDGRP=""
	mostrarDPLU="none"
	TamTabla=980
    TamCeldaBotones=40
    TamCeldaScroll=20%>
            <div id="CollapseSection"> 
                <a id="nocollapse_all_button"  href="javascript:animatedcollapse.show(['DatosDGRP','DatosDPLU']); hideNoCollapse();"><img class="CollapseButton" src="<%=ImgCollapseAll%>" <%=ParamImgCollapseAll %> alt="" /></a> 
                <a id="collapse_all_button"  href="javascript:animatedcollapse.hide(['DatosDGRP','DatosDPLU']);hideCollapse();" style="display:none"><img class="CollapseButton" src="<%=ImgNoCollapseAll%>" <%=ParamImgNoCollapseAll %> alt="" /></a>
            </div>
        
        <div class="Section" id="S_DatosDGRP" style="width:100%">
        
            <a href="#" rel="toggle[DatosDGRP]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitCarTitlGrp%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                </div>
            </a>
        <div class="SectionPanel" id="DatosDGRP" style="display:<%=mostrarDGRP%>;width:100%;">
            <div class="overflowXauto">
	     <br />
         <!--<center>-->
        <table id="filaParametros2" class="width90 md-table-responsive bCollapse">
			    <tr>
                    <%DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitAgrCod & "</b>"
                    DrawceldaDet "'ENCABEZADOL underOrange width25'","", "left", true,"<b>" & LitDescripcion & "</b>"
	                DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitNumPlu & "</b>"
                    DrawceldaDet "'ENCABEZADOL underOrange width15'","", "left", true,"<b>" & LitDesBotTpv & "</b>"
                    DrawceldaDet "'ENCABEZADOL underOrange width5'","", "left", true,"<b>" & LitColor & "</b>"
                    DrawCeldaDet "'ENCABEZADOL underOrange width10'", "", "", 0, ""
                    DrawCeldaDet "'ENCABEZADOL underOrange width10'", "", "", 0, ""%>
                </tr>
                <tr>
                    <td class="CELDAL7 underOrange width5">
                    </td>
                    <td class="CELDAL7 underOrange width25">
                        <input id="i_descripcion" class="width80" type="text" name="i_descripcion" maxlength="30" size="40" value=""/>
                    </td>
                    <td class="CELDAL7 underOrange width5">
                        <input class="width80" type="text" id="i_nplu" name="i_nplu" maxlength="3" size="4" value=""/>
                    </td>
                    <td class="CELDAL7 underOrange width15">
                       <input class="width80" type="text" id="i_descripBotTPV" name="i_descripBotTPV" maxlength="10" size="15" value=""/>
                    </td>
                    <td class="CELDAL7 underOrange width5">
	                    <input class="CELDA" type="text" size="8" name="i_color" id="i_color" value="" onchange="actualizaCampoColoresGRP();"/>
                        <input type="hidden" id="previous_color" value=""/>
	                </td>
	                <td class="CELDAL7 underOrange width10">
                        <div class="col-sm-1 col-xxs-2 overflowI color-box" style="width: 45px; height:45px;">
                               <div id="Box" style="background-color:#cccccc;"></div> 
                            </div>
	                    <!--<input class="width40" type="text" size="1" style="border:1px solid black" readonly="readonly" name="color_colorGrp" id="color_colorGrp"/>
	                    <img border="0" style="cursor:pointer" onclick="NuevaCapaColorGrp('i_color')" alt="<%=LitSelColor%>" title="<%=LitSelColor%>" src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%>/>-->
                    </td>
                    <td class="CELDAL7 underOrange width10">
	                    <a class='ic-accept noMTop' href="javascript:AnadirGrp();" onblur="javascript:document.getElementById('i_descripcion').focus();"><img src="<%=themeIlion %><%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitNuevGrp%>" title="<%=LitNuevGrp%>"/></a>
                    </td>
                </tr>
            </table>
			<!--<table id="filaParametros2" style="display:" width="650">
			    <tr>
                    <td class="CELDA" width="40"><%=LitAgrCod%></td>
                    <td class="CELDA" width="250"><%=LitDescripcion%></td>
                    <td class="CELDA" width="60"><%=LitNumPlu%></td>
                    <td class="CELDA" width="135"><%=LitDesBotTpv%></td>
                    <td class="CELDA" width="55"><%=LitColor%></td>
                    <td class="CELDA" width="90">&nbsp;</td>
                    <td class="CELDA" width="20">&nbsp;</td>
                </tr>
                <tr>
                    <td class="CELDA">
                    </td>
                    <td class="CELDA">
                        <input id="i_descripcion" class="CELDA" type="text" name="i_descripcion" maxlength="30" size="40" value=""/>
                    </td>
                    <td><input class="CELDA" type="text" id="i_nplu" name="i_nplu" maxlength="3" size="4" value=""/></td>
                    <td><input class="CELDA" type="text" id="i_descripBotTPV" name="i_descripBotTPV" maxlength="10" size="15" value=""/></td>
                    <td width="65" class="CELDA">
	                    <input class="CELDA" type="text" size="8" onchange="actualizaCampoColoresGRP()" name="i_color" id="i_color" value=""/>
	                </td>
	                <td width="50" valign=middle>
	                    <input class="CELDA" type="text" size="1" style="border:1px solid black" readonly="readonly" name="color_colorGrp" id="color_colorGrp"/>
	                    <img border="0" style="cursor:pointer" onclick="NuevaCapaColorGrp('i_color')" alt="<%=LitSelColor%>" title="<%=LitSelColor%>" src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%>/>
                    </td>
                    <td class="CELDA">
	                    <a href="javascript:AnadirGrp();" onblur="javascript:document.getElementById('i_descripcion').focus();"><img src="../images/<%=ImgAplicar%>" <%=ParamImgAplicar%> alt="<%=LitNuevGrp%>" title="<%=LitNuevGrp%>"/></a>
                    </td>
                </tr>
                <tr>
                    <td colspan="11" align="center" valign="top"  >
                        <script language="javascript">CrearCapaColoresGRP();</script>
                    </td>
                </tr>
            </table>-->

			<iframe id="frGRP_det" name="fr_GRP_det" src='configuracionPLU_grp_det.asp?mode=browse' class="width90 iframe-data md-table-responsive" frameborder="yes" noresize="noresize" height="265"></iframe>
            <table width="750">
                <%DrawFila ""%>
                    <td class=CELDA7 width="200">
			            <span ID="barras" STYLE="display:initial"></span>
		            </td>
                    <td class="CELDA7">
                        <a class="CELDAREFB"  href="javascript:VistaPrevia()" onmouseover="self.status='<%=LitVistaPrevia%>'; return true;" onmouseout="self.status=''; return true;"><%=LitVistaPrevia%></A>
                    </td>
	            <%CloseFila%>
            </table>
	    <!--</center>-->
        </div>
            </div>
        </div>
        <div class="Section" id="S_DatosDPLU" style="width:100%">
            <a href="#" rel="toggle[DatosDPLU]" data-openimage="<%=ImgNoCollapse %>" data-closedimage="<%=ImgCollapse %>">
                <div class="SectionHeader">
                    <%=LitCarTitlPlu%>
                    <img class="btn_folder" src="<%=ImgNoCollapse %>" alt="" <%=ParamImgCollapse %> />
                </div>
            </a>
            <div class="SectionPanel" id="DatosDPLU" style="display:<%=mostrarDPLU%>;width:100%;">
	         <br />
             <!--<center>-->
         
			    <iframe class="width90 iframe-data md-table-responsive" id="frPLU" name="fr_PLU" src='configuracionPLU_art.asp?mode=browse' height='350' frameborder="yes" noresize="noresize"></iframe>
	        <!--</center>-->
            </div>
        </div>
    <script type="text/javascript">
        window.onload = function() {
            updateColpick();
        }
        Redimensionar();
        document.getElementById("i_descripcion").focus();

	</script>
    <!--</td></tr></table>-->
<%end if 'if accesoPagina%>
</body>
</html>