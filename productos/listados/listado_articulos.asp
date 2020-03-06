<%@ Language=VBScript %><% 
dim  enc
set enc = Server.CreateObject("Owasp_Esapi.Encoder")
function EncodeForHtml(data)
	if data & "" <>"" then
	  EncodeForHtml = enc.EncodeForHtmlAttribute(data)
	else
	  EncodeForHtml = data
	end if
end function
%>  

<% Server.ScriptTimeout = 1200 %>
<%
''ricardo 16-11-2007 se cambia la dsn desde dsncliente a backendlistados
%>
<%

''ricardo 2/4/2003 añadir como opcion para ver el coste medio
' VGR 	 30/4/2003 añadir campos opcionales y filtros.
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LitTitulo2%></title>

<meta http-equiv="Content-Type" content="text/html; charset=<%=session("caracteres")%>"/>

<link rel="stylesheet" href="../../pantalla.css" media="SCREEN"/>
<link rel="stylesheet" href="../../impresora.css" media="PRINT"/>
</head>

<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../XSSProtection.inc" -->
<!--#include file="../../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../modulos.inc" -->
<!--#include file="../articulos.inc" -->
<!--#include file="listado_articulos.inc" -->
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../../CatFamSubResponsive.inc" -->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->
<!--#include file="../../styles/formularios.css.inc" -->

<%
	si_tiene_modulo_terminales=ModuloContratado(session("ncliente"),ModTerminales)
	si_tiene_modulo_importaciones=ModuloContratado(session("ncliente"),ModImportaciones)
	'cag
	si_tiene_modulo_ebesa=ModuloContratado(session("ncliente"),ModEBESA)
	'fin cag
	''ricardo 5-12-2008 se mostraran los campos PLU y GRUPO
	si_tiene_modulo_02=ModuloContratado(session("ncliente"),"02")
	si_tiene_modulo_33=ModuloContratado(session("ncliente"),"33")
    ''AMF:25/9/2013:E-Commerce
    si_tiene_modulo_ecommerce = ModuloContratado(session("ncliente"),ModEComerce)
    si_tiene_modulo_OrCU=ModuloContratado(session("ncliente"),ModOrCU)
    'i(EJM 20/02/07)
    activarCampos=""
    if si_tiene_modulo_ebesa<>0 then
        activarCampos="checked"
    end if
    'fin(EJM 20/02/07)
%>

<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
    function cambiar_smi() {
        while (document.listado_articulos.stockmayoroigual.value.search(" ") != -1) {
            document.listado_articulos.stockmayoroigual.value = document.listado_articulos.stockmayoroigual.value.replace(" ", "");
        }
        //se podra dejar el stockmayoroigual a vacio
        if (isNaN(document.listado_articulos.stockmayoroigual.value.replace(",", "."))) {
            window.alert("<%=LITMSGSTOCKNUMERICO%>");
            document.listado_articulos.stockmayoroigual.value = "0";
            return;
        }
        document.listado_articulos.stockmayoroigual.value = document.listado_articulos.stockmayoroigual.value.replace(".", ",");
    }
    function cambiar_neg() {
        if (document.listado_articulos.stocksolonegativo.checked==true) {
            document.listado_articulos.stockmayoroigual.value="";
            document.listado_articulos.stockmayoroigual.disabled=true;
        }
        else {
            document.listado_articulos.stockmayoroigual.value="0";
            document.listado_articulos.stockmayoroigual.disabled=false;
        }
    }

    //Desencadena la búsqueda del artículo cuya referencia se indica
    function TraerProveedor(mode) {
        document.listado_articulos.action="listado_articulos.asp?nproveedor=" + document.listado_articulos.nproveedor.value + "&mode=" + mode + "&prov=" + document.listado_articulos.nproveedor.value;
        document.listado_articulos.method="post";
        document.listado_articulos.submit();
    }

    function sel_todos(modo) {
        <%if si_tiene_modulo_importaciones = 0 then%>
            parent.pantalla.document.listado_articulos.elements["ver_almacen"].checked = modo;
        <%end if%>
        parent.pantalla.document.listado_articulos.elements["ver_familia"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_dto"].checked=modo;
        //cag
        parent.pantalla.document.listado_articulos.elements["ver_margen"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_recargo"].checked=modo;
        //fin cag
        parent.pantalla.document.listado_articulos.elements["ver_iva"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_pvp"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_divisa"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_codbarras"].checked=modo;
        //cag

        //MAP - 02/01/2013 - Add Img1, Img2, Img3
        parent.pantalla.document.listado_articulos.elements["ver_Img1"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_Img2"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_Img3"].checked=modo;

        parent.pantalla.document.listado_articulos.elements["ver_coste"].checked=modo;
        si_tiene_modulo_ebesa="<%=si_tiene_modulo_ebesa%>";
        if (si_tiene_modulo_ebesa!="0") {
            parent.pantalla.document.listado_articulos.elements["ver_clave"].checked=modo;
            parent.pantalla.document.listado_articulos.elements["ver_modifs"].checked=modo;
        }
        //fin cag

        //i(EJM 19/02/07)
        if (si_tiene_modulo_ebesa!="0") {
            parent.pantalla.document.listado_articulos.elements["ver_Lemargen"].checked=modo;
        }
        parent.pantalla.document.listado_articulos.elements["ver_CodSub"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_Embalaje"].checked=modo;
        //fin(EJM 19/02/07)

        parent.pantalla.document.listado_articulos.elements["ver_stock"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_smin"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_smax"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_reposicion"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_precibir"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_pservir"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_pmin"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_coste_medio"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_pvpiva"].checked=modo;
        <%if si_tiene_modulo_terminales<>0 then%>
            parent.pantalla.document.listado_articulos.elements["ver_codTerminal"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_nomTerminal"].checked=modo;
        <%end if%>
        parent.pantalla.document.listado_articulos.elements["ver_desAmpliada"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_tipoArticulo"].checked=modo;
        num_campos_perso=parent.pantalla.document.listado_articulos.elements["num_puestos"].value;
        for (ki=1;ki<=num_campos_perso;ki++) {
            if (eval("parent.pantalla.document.listado_articulos.elements['si_campo" + ki + "'].value=='1'")) {
                eval("parent.pantalla.document.listado_articulos.elements['ver_campo" + ki + "'].checked=modo");
            }
        }
        <%if si_tiene_modulo_terminales<>0 then%>
            <%if si_tiene_modulo_02<>0 or si_tiene_modulo_33<>0 then%>
                parent.pantalla.document.listado_articulos.elements["ver_PLU"].checked=modo;
        parent.pantalla.document.listado_articulos.elements["ver_GRPPLU"].checked=modo;
        <%end if%>
    <%end if%>
    parent.pantalla.document.listado_articulos.elements["ver_pnf"].checked = modo;

        parent.pantalla.document.listado_articulos.elements["ver_medida"].checked = modo;
        parent.pantalla.document.listado_articulos.elements["ver_peso"].checked = modo;
        parent.pantalla.document.listado_articulos.elements["ver_medidaventa"].checked = modo;
    }


    /*FUNCTION ver_imagen ---- to show img1, img2, img3 in listing results (function poner_imagen doesn't work)*/
    function ver_imagen(mode, idFoto)
    {
        if (mode!="select1")
        {
            eval("opcimagen=document.listado_articulos.ver_Img"+idFoto+".value;")
            if (opcimagen=="on")
            {
                if (document.listado_articulos.NumRegs!=undefined)
                {
                    numregs=document.listado_articulos.NumRegs.value;
                    if (numregs>0)
                    {
                        for(i=0;i<numregs;i++)
                        {

                            valorI="";
                            if (numregs>1)
                            {
                                valorI="["+i+"]";
                            }
                
                            if(eval("document.listado_articulos.mostrar_foto"+idFoto+"_"+i)!=null) {
                                eval("mfoto=document.listado_articulos.mostrar_foto"+idFoto+"_"+i+".value;");
                            }

                            if (mfoto==1)
                            {
                                eval("document.all('capa_foto"+idFoto+"_"+i+"').style.display='';");
                                //cogemos las propiedades actuales de la foto
                                eval("w=document.listado_articulos.foto_articulo"+idFoto+"_"+i+".width;");
                                eval("h=document.listado_articulos.foto_articulo"+idFoto+"_"+i+".height;");

                                //cogemos las propiedades actuales y las ponemos como antiguas de la foto
                                eval("document.listado_articulos.foto_art"+idFoto+"_w_"+i+".value=w;");
                                eval("document.listado_articulos.foto_art"+idFoto+"_h_"+i+".value=h;");

                                if (w>100)
                                {
                                    ratio=w/100;
                                    w = w/ratio;
                                    h=h/ratio;
                                }

                                if (h>50)
                                {
                                    ratio=h/50;
                                    h = h/ratio;
                                    w=w/ratio;
                                }
                                eval("document.listado_articulos.foto_articulo"+idFoto+"_"+i+".width=w;");
                                eval("document.listado_articulos.foto_articulo"+idFoto+"_"+i+".height=h;");
                                eval("document.all('capa_foto"+idFoto+"_"+i+"').style.display='';");
                            }
                        }
                    }
                }
            }
        }
    }

    function ver_imagenAll(mode)
    {
        /* verdadero=1;
         formulario="listado_articulos";
     
         numregs=document.listado_articulos.NumRegs.value;
         if (numregs>0)
         {
             //IMG1
             eval("opcimagen1=document.listado_articulos.ver_Img1.value;")
             if (opcimagen1=="on")
             {
                 poner_imagen(verdadero,formulario,"capa_fotoA","foto_artA_w","foto_artA_h","mostrar_fotoA","foto_articuloA","NumRegs",formulario,100,50);
             }
             
             //IMG2
             eval("opcimagen2=document.listado_articulos.ver_Img2.value;")
             if (opcimagen2=="on")
             {
                 poner_imagen(verdadero,formulario,"capa_fotoB","foto_artB_w","foto_artB_h","mostrar_fotoB","foto_articuloB","NumRegs",formulario,100,50);
             }
     
     
             //IMG3
             eval("opcimagen3=document.listado_articulos.ver_Img3.value;")
             if (opcimagen3=="on")
             {
     (verdadero,formulario,"capa_fotoC","foto_artC_w","foto_artC_h","mostrar_fotoC","foto_articuloC","NumRegs",formulario,100,50);
             }
         }*/

        ver_imagen(mode,"1");
        ver_imagen(mode,"2");
        ver_imagen(mode,"3");
    }
</script>
<%function calcula_precio()
 dim temporada
 dim tarifa
 dim rango

    rstAux3.cursorlocation=3
  rstAux3.open "select * from articulos_temporada with(NOLOCK) where referencia='" & rst("referencia") & "'", session("backendlistados")
  if not rstAux3.eof then
    rstAux4.cursorlocation=3
	rstAux4.open "select * from temporadas with(NOLOCK) where codigo='" & rstAux3("temporada") & "' and CONVERT(char(12), GETDATE(), 3) >= f_min AND CONVERT(char(12), GETDATE(), 3) <= f_max", session("backendlistados")
	if not rstAux4.eof then
		temporada=rstAux4("codigo")
	else
		temporada=session("ncliente") & "BASE"
	end if
	rstAux4.close
  else
	temporada=session("ncliente") & "BASE"
  end if
  rstAux3.close

'  rstAux3.open "select * from articulos_tarifa where referencia='" & rst("referencia") & "'", session("backendlistados"),adUseClient, adLockReadOnly
'  if not rstAux3.eof then
'	rstAux4.open "select * from tarifas where codigo='" & rstAux3("tarifa") & "'", session("backendlistados"),adUseClient, adLockReadOnly
'	if not rstAux4.eof then
'		tarifa=rstAux4("codigo")
'	else
'		tarifa=session("ncliente") & "BASE"
'	end if
'	rstAux4.close
'  else
'	tarifa="BASE"
'  end if
'  rstAux3.close

'como estamos sacando un listado por articulo, que no por cliente
'y la tarifa solo se aplica a un cliente, entonces la tarifa sera BASE
tarifa=session("ncliente") & "BASE"
    rstAux3.cursorlocation=3
  rstAux3.open "select * from articulos_rango with(NOLOCK) where referencia='" & rst("referencia") & "'", session("backendlistados")
  if not rstAux3.eof then
  rstAux4.cursorlocation=3
	rstAux4.open "select * from rangos with(NOLOCK) where codigo='" & rstAux3("rango") & "' and ((minimo=1 AND maximo=1) or (minimo is null and maximo is null))", session("backendlistados")
	if not rstAux4.eof then
		rango=rstAux4("codigo")
	else
		rango=session("ncliente") & "BASE"
	end if
	rstAux4.close
  else
	rango=session("ncliente") & "BASE"
  end if
  rstAux3.close

  if temporada=session("ncliente") & "BASE" and tarifa=session("ncliente") & "BASE" and rango=session("ncliente") & "BASE" then
	calcula_precio=rst("pvp")
  else
    rstAux3.cursorlocation=3
	  rstAux3.open "select * from precios with(NOLOCK) where referencia='" & rst("referencia") & "' and rango='" & rango & "' and temporada='" & temporada & "' and tarifa='" & tarifa & "'", session("backendlistados")
	  if not rstAux3.eof then
			calcula_precio=rstAux3("PVPDTO")
			if calcula_precio<0 then calcula_precio=rst("pvp")-((rst("pvp")*(-calcula_precio))/100)
	  else
		calcula_precio=rst("pvp")
	  end if
	  rstAux3.close
  end if
end function

%><body class="BODY_ASP" onload="ver_imagenAll('<%=enc.EncodeForJavascript(mode)%>');self.status='';"><%

'*****************************************************************************
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************
'*****************************************************************************
const borde=0
%>
<form name="listado_articulos" method="post">
	<%  PintarCabecera "listado_articulos.asp"

		WaitBoxOculto LitEsperePorFavor

       
		'Leer parámetros de la página
  		mode=enc.EncodeForJavascript(Request.QueryString("mode"))
  		
		if mode& ""="" then mode=enc.EncodeForJavascript(Request.form("mode"))
		if trim(mode)="browse" then mode="ver"

    	apaisado=iif(limpiaCadena(request("apaisado"))>"","SI","")

		submode	= enc.EncodeForJavascript(Request.QueryString("submode"))
		if submode="" then
			submode	= enc.EncodeForJavascript(Request.form("submode"))
		end if
		if submode&""="" then submode="primero"
		%><input type="hidden" name="submode" value="<%=EncodeForHtml(submode)%>"/><%

		if enc.EncodeForJavascript(request.querystring("referencia")) >"" then
		   referencia = limpiaCadena(request.querystring("referencia"))
		else
		   referencia = limpiaCadena(request.form("referencia"))
		end if
		if enc.EncodeForJavascript(request.form("nombre"))>"" then
			nombre = limpiaCadena(request.form("nombre"))
		else
			nombre=limpiaCadena(request.querystring("nombre"))
		end if

		if enc.EncodeForJavascript(request.form("familia"))>"" then
			familia = limpiaCadena(request.form("familia"))
		else
			familia = limpiaCadena(request.querystring("familia"))
		end if

		if enc.EncodeForJavascript(request.form("familia_padre"))>"" then
			familia_padre = limpiaCadena(request.form("familia_padre"))
		else
			familia_padre = limpiaCadena(request.querystring("familia_padre"))
		end if

		if enc.EncodeForJavascript(request.form("categoria"))>"" then
			categoria = limpiaCadena(request.form("categoria"))
		else
			categoria = limpiaCadena(request.querystring("categoria"))
		end if

        if enc.EncodeForJavascript(request.form("tarifa"))>"" then
			tarifa = limpiaCadena(request.form("tarifa"))
		else
			tarifa = limpiaCadena(request.querystring("tarifa"))
		end if
        
		if enc.EncodeForJavascript(request.form("tipoarticulo"))>"" then
			tipoarticulo = limpiaCadena(request.form("tipoarticulo"))
		else
			tipoarticulo = limpiaCadena(request.querystring("tipoarticulo"))
		end if

        if enc.EncodeForJavascript(request.Form("pnf")) > "" then
            pnf = limpiaCadena(request.Form("pnf"))
        else
            pnf = limpiaCadena(request.QueryString("pnf"))
        end if
        
		if enc.EncodeForJavascript(request.form("ordenar"))>"" then
			ordenar = limpiaCadena(request.form("ordenar"))
		else
			ordenar = limpiaCadena(request.querystring("ordenar"))
		end if

		if enc.EncodeForJavascript(request.form("almacen"))>"" then
			almacen = limpiaCadena(request.form("almacen"))
		else
			almacen = limpiaCadena(request.querystring("almacen"))
		end if

		if enc.EncodeForJavascript(request.form("consigna"))>"" then
			consigna = limpiaCadena(request.form("consigna"))
		else
			consigna = limpiaCadena(request.querystring("consigna"))
		end if

		if enc.EncodeForJavascript(request.form("stockmayoroigual"))>"" then
			stockmayoroigual = limpiaCadena(request.form("stockmayoroigual"))
		else
			stockmayoroigual = limpiaCadena(request.querystring("stockmayoroigual"))
		end if
		
		''ricardo 23-3-2009 se gestiona que salgan solamente articulos con stock negativo
		if enc.EncodeForJavascript(request.form("stocksolonegativo"))>"" then
			stocksolonegativo = limpiaCadena(request.form("stocksolonegativo"))
		else
			stocksolonegativo = limpiaCadena(request.querystring("stocksolonegativo"))
		end if

		if enc.EncodeForJavascript(request.form("articulosTer"))>"" then
			articulosTer = limpiaCadena(request.form("articulosTer"))
		else
			articulosTer = limpiaCadena(request.querystring("articulosTer"))
		end if

		if enc.EncodeForJavascript(request.form("articulosBaja"))>"" then
			articulosBaja = limpiaCadena(request.form("articulosBaja"))
		else
			articulosBaja = limpiaCadena(request.querystring("articulosBaja"))
		end if

        if enc.EncodeForJavascript(request.form("onlyWebStore"))>"" then
			onlyWebStore = limpiaCadena(request.form("onlyWebStore"))
		else
			onlyWebStore = limpiaCadena(request.querystring("onlyWebStore"))
		end if

		if enc.EncodeForJavascript(request.form("ver_almacen"))>"" then
			ver_almacen = limpiaCadena(request.form("ver_almacen"))
		else
			ver_almacen = limpiaCadena(request.querystring("ver_almacen"))
		end if

		if enc.EncodeForJavascript(request.form("ver_familia"))>"" then
			ver_familia = limpiaCadena(request.form("ver_familia"))
		else
			ver_familia = limpiaCadena(request.querystring("ver_familia"))
		end if

		if enc.EncodeForJavascript(request.form("ver_dto"))>"" then
			ver_dto = limpiaCadena(request.form("ver_dto"))
		else
			ver_dto = limpiaCadena(request.querystring("ver_dto"))
		end if

		if enc.EncodeForJavascript(request.form("ver_margen"))>"" then
			ver_margen = limpiaCadena(request.form("ver_margen"))
		else
			ver_margen = limpiaCadena(request.querystring("ver_margen"))
		end if

		if enc.EncodeForJavascript(request.form("ver_coste"))>"" then
			ver_coste = limpiaCadena(request.form("ver_coste"))
		else
			ver_coste = limpiaCadena(request.querystring("ver_coste"))
		end if

		if enc.EncodeForJavascript(request.form("ver_recargo"))>"" then
			ver_recargo = limpiaCadena(request.form("ver_recargo"))
		else
			ver_recargo = limpiaCadena(request.querystring("ver_recargo"))
		end if
'cag clave y modificadores
		if enc.EncodeForJavascript(request.form("ver_clave"))>"" then
			ver_clave = limpiaCadena(request.form("ver_clave"))
		else
			ver_clave = limpiaCadena(request.querystring("ver_clave"))
		end if

		if enc.EncodeForJavascript(request.form("ver_modifs"))>"" then
			ver_modifs = limpiaCadena(request.form("ver_modifs"))
		else
			ver_modifs = limpiaCadena(request.querystring("ver_modifs"))
		end if
'fin cag

        'i(EJM 19/02/07)
		if enc.EncodeForJavascript(request.form("ver_Lemargen"))>"" then
			ver_Lemargen = limpiaCadena(request.form("ver_Lemargen"))
		else
			ver_Lemargen = limpiaCadena(request.querystring("ver_Lemargen"))
		end if

		if enc.EncodeForJavascript(request.form("ver_CodSub"))>"" then
			ver_CodSub = limpiaCadena(request.form("ver_CodSub"))
		else
			ver_CodSub = limpiaCadena(request.querystring("ver_CodSub"))
		end if

		if enc.EncodeForJavascript(request.form("ver_Embalaje"))>"" then
			ver_Embalaje = limpiaCadena(request.form("ver_Embalaje"))
		else
			ver_Embalaje = limpiaCadena(request.querystring("ver_Embalaje"))
		end if
        'fin(EJM 19/02/07)


		if enc.EncodeForJavascript(request.form("ver_iva"))>"" then
			ver_iva = limpiaCadena(request.form("ver_iva"))
		else
			ver_iva = limpiaCadena(request.querystring("ver_iva"))
		end if

		if enc.EncodeForJavascript(request.form("ver_pvp"))>"" then
			ver_pvp = limpiaCadena(request.form("ver_pvp"))
		else
			ver_pvp = limpiaCadena(request.querystring("ver_pvp"))
		end if

		if enc.EncodeForJavascript(request.form("ver_pvd"))>"" then
			ver_pvd = limpiaCadena(request.form("ver_pvd"))
		else
			ver_pvd = limpiaCadena(request.querystring("ver_pvd"))
		end if

		if enc.EncodeForJavascript(request.form("ver_divisa"))>"" then
			ver_divisa = limpiaCadena(request.form("ver_divisa"))
		else
			ver_divisa = limpiaCadena(request.querystring("ver_divisa"))
		end if

		if enc.EncodeForJavascript(request.form("ver_proveedor"))>"" then
			ver_proveedor = limpiaCadena(request.form("ver_proveedor"))
		else
			ver_proveedor = limpiaCadena(request.querystring("ver_proveedor"))
		end if

		if enc.EncodeForJavascript(request.form("ver_codbarras"))>"" then
			ver_codbarras = limpiaCadena(request.form("ver_codbarras"))
		else
			ver_codbarras = limpiaCadena(request.querystring("ver_codbarras"))
		end if

		if enc.EncodeForJavascript(request.form("ver_stock"))>"" then
			ver_stock = limpiaCadena(request.form("ver_stock"))
		else
			ver_stock = limpiaCadena(request.querystring("ver_stock"))
		end if

		if enc.EncodeForJavascript(request.form("ver_smin"))>"" then
			ver_smin = limpiaCadena(request.form("ver_smin"))
		else
			ver_smin = limpiaCadena(request.querystring("ver_smin"))
		end if

        if enc.EncodeForJavascript(request.form("ver_smax"))>"" then
			ver_smax = limpiaCadena(request.form("ver_smax"))
		else
			ver_smax = limpiaCadena(request.querystring("ver_smax"))
		end if



		if enc.EncodeForJavascript(request.form("ver_reposicion"))>"" then
			ver_reposicion = limpiaCadena(request.form("ver_reposicion"))
		else
			ver_reposicion = limpiaCadena(request.querystring("ver_reposicion"))
		end if

		if enc.EncodeForJavascript(request.form("ver_precibir"))>"" then
			ver_precibir = limpiaCadena(request.form("ver_precibir"))
		else
			ver_precibir = limpiaCadena(request.querystring("ver_precibir"))
		end if

		if enc.EncodeForJavascript(request.form("ver_pservir"))>"" then
			ver_pservir = limpiaCadena(request.form("ver_pservir"))
		else
			ver_pservir = limpiaCadena(request.querystring("ver_pservir"))
		end if

		if enc.EncodeForJavascript(request.form("ver_pmin"))>"" then
			ver_pmin = limpiaCadena(request.form("ver_pmin"))
		else
			ver_pmin = limpiaCadena(request.querystring("ver_pmin"))
		end if

		if enc.EncodeForJavascript(request.form("ver_coste_medio"))>"" then
			ver_coste_medio = limpiaCadena(request.form("ver_coste_medio"))
		else
			ver_coste_medio = limpiaCadena(request.querystring("ver_coste_medio"))
		end if

		if enc.EncodeForJavascript(request.form("ver_pvpiva"))>"" then
			ver_pvpiva = limpiaCadena(request.form("ver_pvpiva"))
		else
			ver_pvpiva = limpiaCadena(request.querystring("ver_pvpiva"))
		end if

		if enc.EncodeForJavascript(request.form("ver_codTerminal"))>"" then
			ver_codTerminal = limpiaCadena(request.form("ver_codTerminal"))
		else
			ver_codTerminal = limpiaCadena(request.querystring("ver_codTerminal"))
		end if

		if enc.EncodeForJavascript(request.form("ver_nomTerminal"))>"" then
			ver_nomTerminal = limpiaCadena(request.form("ver_nomTerminal"))
		else
			ver_nomTerminal = limpiaCadena(request.querystring("ver_nomTerminal"))
		end if

        'MAP 21/12/2012 - Recover img1, img2,img3 check from form

        if enc.EncodeForJavascript(request.Form("ver_Img1"))>"" then
            ver_Img1 = limpiaCadena(request.Form("ver_Img1"))
        else
            ver_Img1 = limpiaCadena(request.QueryString("ver_Img1"))
        end if
        if enc.EncodeForJavascript(request.Form("ver_Img2"))>"" then
            ver_Img2 = limpiaCadena(request.Form("ver_Img2"))
        else
            ver_Img2 = limpiaCadena(request.QueryString("ver_Img2"))
        end if
        if enc.EncodeForJavascript(request.Form("ver_Img3"))>"" then
           ver_Img3 = limpiaCadena(request.Form("ver_Img3"))
        else
            ver_Img3 = limpiaCadena(request.QueryString("ver_Img3"))
        end if

		if enc.EncodeForJavascript(request.form("ver_desAmpliada"))>"" then
			ver_desAmpliada = limpiaCadena(request.form("ver_desAmpliada"))
		else
			ver_desAmpliada = limpiaCadena(request.querystring("ver_desAmpliada"))
		end if

		if enc.EncodeForJavascript(request.form("ver_tipoArticulo"))>"" then
			ver_tipoArticulo = limpiaCadena(request.form("ver_tipoArticulo"))
		else
			ver_tipoArticulo = limpiaCadena(request.querystring("ver_tipoArticulo"))
		end if

        if enc.EncodeForJavascript(request.Form("pnf"))>"" then
            pnf = limpiaCadena(request.Form("pnf"))
        else
            pnf = limpiaCadena(request.QueryString("pnf"))
        end if

        if enc.EncodeForJavascript(request.Form("ver_pnf"))>"" then
            ver_pnf = limpiaCadena(request.Form("ver_pnf"))
        else
            ver_pnf = limpiaCadena(request.QueryString("ver_pnf"))
        end if

		if enc.EncodeForJavascript(request.form("desde_fb"))>"" then
			desdeFechaBaja = limpiaCadena(request.form("desde_fb"))
		else
			desdeFechaBaja = limpiaCadena(request.querystring("desde_fb"))
		end if

		if enc.EncodeForJavascript(request.form("hasta_fb"))>"" then
			hastaFechaBaja = limpiaCadena(request.form("hasta_fb"))
		else
			hastaFechaBaja = limpiaCadena(request.querystring("hasta_fb"))
		end if

        'MAP 20/12/2012 - Recover create date from form
        if enc.EncodeForJavascript(request.form("desde_fc"))>"" then
			desdeFechaCreacion = limpiaCadena(request.form("desde_fc"))
		else
			desdeFechaCreacion = limpiaCadena(request.querystring("desde_fc"))
		end if

		if enc.EncodeForJavascript(request.form("hasta_fc"))>"" then
			hastaFechaCreacion = limpiaCadena(request.form("hasta_fc"))
		else
			hastaFechaCreacion = limpiaCadena(request.querystring("hasta_fc"))
		end if

		if enc.EncodeForJavascript(request.form("campo1"))>"" then
			campo1 = limpiaCadena(request.form("campo1"))
		else
			campo1 = limpiaCadena(request.querystring("campo1"))
		end if

		if enc.EncodeForJavascript(request.form("campo2"))>"" then
			campo2 = limpiaCadena(request.form("campo2"))
		else
			campo2 = limpiaCadena(request.querystring("campo2"))
		end if

		if enc.EncodeForJavascript(request.form("campo3"))>"" then
			campo3 = limpiaCadena(request.form("campo3"))
		else
			campo3 = limpiaCadena(request.querystring("campo3"))
		end if

		if enc.EncodeForJavascript(request.form("campo4"))>"" then
			campo4 = limpiaCadena(request.form("campo4"))
		else
			campo4 = limpiaCadena(request.querystring("campo4"))
		end if

		if enc.EncodeForJavascript(request.form("campo5"))>"" then
			campo5 = limpiaCadena(request.form("campo5"))
		else
			campo5 = limpiaCadena(request.querystring("campo5"))
		end if

		if enc.EncodeForJavascript(request.form("campo6"))>"" then
			campo6 = limpiaCadena(request.form("campo6"))
		else
			campo6 = limpiaCadena(request.querystring("campo6"))
		end if

		if enc.EncodeForJavascript(request.form("campo7"))>"" then
			campo7 = limpiaCadena(request.form("campo7"))
		else
			campo7 = limpiaCadena(request.querystring("campo7"))
		end if

		if enc.EncodeForJavascript(request.form("campo8"))>"" then
			campo8 = limpiaCadena(request.form("campo8"))
		else
			campo8 = limpiaCadena(request.querystring("campo8"))
		end if

		if enc.EncodeForJavascript(request.form("campo9"))>"" then
			campo9 = limpiaCadena(request.form("campo9"))
		else
			campo9 = limpiaCadena(request.querystring("campo9"))
		end if

		if enc.EncodeForJavascript(request.form("campo10"))>"" then
			campo10 = limpiaCadena(request.form("campo10"))
		else
			campo10 = limpiaCadena(request.querystring("campo10"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo1"))>"" then
			ver_campo1 = limpiaCadena(request.form("ver_campo1"))
		else
			ver_campo1 = limpiaCadena(request.querystring("ver_campo1"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo2"))>"" then
			ver_campo2 = limpiaCadena(request.form("ver_campo2"))
		else
			ver_campo2 = limpiaCadena(request.querystring("ver_campo2"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo3"))>"" then
			ver_campo3 = limpiaCadena(request.form("ver_campo3"))
		else
			ver_campo3 = limpiaCadena(request.querystring("ver_campo3"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo4"))>"" then
			ver_campo4 = limpiaCadena(request.form("ver_campo4"))
		else
			ver_campo4 = limpiaCadena(request.querystring("ver_campo4"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo5"))>"" then
			ver_campo5 = limpiaCadena(request.form("ver_campo5"))
		else
			ver_campo5 = limpiaCadena(request.querystring("ver_campo5"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo6"))>"" then
			ver_campo6 = limpiaCadena(request.form("ver_campo6"))
		else
			ver_campo6 = limpiaCadena(request.querystring("ver_campo6"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo7"))>"" then
			ver_campo7 = limpiaCadena(request.form("ver_campo7"))
		else
			ver_campo7 = limpiaCadena(request.querystring("ver_campo7"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo8"))>"" then
			ver_campo8 = limpiaCadena(request.form("ver_campo8"))
		else
			ver_campo8 = limpiaCadena(request.querystring("ver_campo8"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo9"))>"" then
			ver_campo9 = limpiaCadena(request.form("ver_campo9"))
		else
			ver_campo9 = limpiaCadena(request.querystring("ver_campo9"))
		end if

		if enc.EncodeForJavascript(request.form("ver_campo10"))>"" then
			ver_campo10 = limpiaCadena(request.form("ver_campo10"))
		else
			ver_campo10 = limpiaCadena(request.querystring("ver_campo10"))
		end if
		
		''ricardo 5-12-2008 se mostraran los campos PLU y GRUPO
		if enc.EncodeForJavascript(request.form("ver_PLU"))>"" then
			ver_PLU = limpiaCadena(request.form("ver_PLU"))
		else
			ver_PLU = limpiaCadena(request.querystring("ver_PLU"))
		end if
		if enc.EncodeForJavascript(request.form("ver_GRPPLU"))>"" then
			ver_GRPPLU = limpiaCadena(request.form("ver_GRPPLU"))
		else
			ver_GRPPLU = limpiaCadena(request.querystring("ver_GRPPLU"))
		end if
		if enc.EncodeForJavascript(request.form("nproveedor"))>"" then
			nproveedor = limpiaCadena(request.form("nproveedor"))
		else
			nproveedor = limpiaCadena(request.querystring("nproveedor"))
		end if

		numregs	= limpiaCadena(Request.QueryString("numregs"))
		if numregs="" then
			numregs	= limpiaCadena(Request.form("numregs"))
		end if
		if numregs&""="" then numregs=0
		''MPC 17/09/2012 Add news fields
        if enc.EncodeForJavascript(request.Form("ver_medida"))>"" then
            ver_medida = limpiaCadena(request.Form("ver_medida"))
        else
            ver_medida = limpiaCadena(request.QueryString("ver_medida"))
        end if

        if enc.EncodeForJavascript(request.Form("ver_peso"))>"" then
            ver_peso = limpiaCadena(request.Form("ver_peso"))
        else
            ver_peso = limpiaCadena(request.QueryString("ver_peso"))
        end if

        if enc.EncodeForJavascript(request.Form("ver_medidaventa"))>"" then
            ver_medidaventa = limpiaCadena(request.Form("ver_medidaventa"))
        else
            ver_medidaventa = limpiaCadena(request.QueryString("ver_medidaventa"))
        end if
        ''END MPC

        'DBS 20140128 
        parametrosBD=obtener_param_obj("A71",session("usuario"),session("ncliente"),"mode")
        lista_obt_obj = Split(parametrosBD, "&")
            'response.Write "<br> parametrosBD "&parametrosBD
        Rates=""
        if isArray(lista_obt_obj) then
            
            tamanyo=Ubound(lista_obt_obj)+1 'tamanyo del vector
            i=0
            while i < tamanyo   
                if Mid(lista_obt_obj(i),1,5)&"" = "RATES" then
                    lista_obt_obj_Aux=Split(lista_obt_obj(i),"=")
                    Rates=lista_obt_obj_Aux(1)
                    obtener="1"                    
                end if                        
                i=i+1
            wend
            if obtener&""<>"1" then
                %><input type="hidden" name="lista_obt_objok" id="lista_obt_objok" value=""/><%
            else
                %><input type="hidden" name="lista_obt_objok" id="Hidden1" value="1"/><%
            end if
        end if
        'response.Write "<br> rates "&Rates
        'response.end
    %><input type="hidden" name="Rates" id="Rates" value="<%=EncodeForHtml(Rates)%>"/><%
  set rstAux = Server.CreateObject("ADODB.Recordset")
  set rstAux2 = Server.CreateObject("ADODB.Recordset")
  set rstAux3 = Server.CreateObject("ADODB.Recordset")
  set rstAux4 = Server.CreateObject("ADODB.Recordset")
  set rstAux5 = Server.CreateObject("ADODB.Recordset")
  set rstAux6 = Server.CreateObject("ADODB.Recordset")
  set rst = Server.CreateObject("ADODB.Recordset")
  set rstPrecio = Server.CreateObject("ADODB.Recordset")
  set rstModifs = Server.CreateObject("ADODB.Recordset")

    'i(EJM 20/02/07) Activar campos para el módulo ebesa
    if si_tiene_modulo_ebesa<> 0 and mode="add" then
        ver_clave="True"
        ver_modifs="True"
        ver_desAmpliada="True"
        ver_tipoArticulo="True"
        ver_Lemargen="True"
    end if
    'fin(EJM 20/02/07) Activar campos para el módulo ebesa
    if nproveedor>"" then
		nproveedor = limpiaCadena(nproveedor)
		rst.cursorlocation=3
		rst.open "select razon_social from proveedores with (NOLOCK) where nproveedor='" + session("ncliente") & Completar(nproveedor,5,"0") + "'" ,session("backendlistados")
		if rst.eof then
			nproveedor=""
			nombre=""%>
			<script type="text/javascript" language="javascript">
			    window.alert("<%=LitMsgProveedorNoExiste%>");
			</script>
        <%else
		    tnombre = rst("razon_social")
		end if
		rst.close
  	end if

	if mode="ver" then%>
		<table width='100%'>
   			<tr>
				<td width="30%" align="left">
					<font class="CELDAB7"><b></b></font>
					<font class="CELDA7">&nbsp;<%=EncodeForHtml("(" & LitEmitido & " " & day(date) & "/" & month(date) & "/" & year(date) & ")")%></font>
				</td>
			</tr>
		</table>
		<hr/>
	<%else%>
		<br/>
	<%end if%>
	<table>

    <%''MPC 25/04/2014 Obtengo los decimales del precio para que en caso de pintarlos salga con esos mismos decimales.
    set connDP=server.CreateObject("ADODB.Connection")
	set commandDP=server.CreateObject("ADODB.Command")
	connDP.open session("dsn_cliente")
	connDP.cursorlocation=3
	commandDP.activeConnection=connDP
	commandDP.CommandType = adCmdText
    commandDP.CommandText= "select dec_precios from configuracion with(nolock) where nempresa = ?"
    commandDP.Parameters.Append commandDP.CreateParameter("@nempresa",adVarChar,adParamInput,5, session("ncliente"))

    set rst = commandDP.execute
    if not rst.eof then
        DEC_PREC = rst("dec_precios")
    end if
	connDP.close
    set commandDP=nothing
    set connDP=nothing
    ''FIN MPC 25/04/2014
    ''ricardo 24-3-2004 si existen campos personalizables con titulo no nulo si saldra la pestaña de campos personalizables
	dim tipo_campo_perso
	dim titulo_campo_perso
	si_campo_personalizables=0
	num_campos_articulos=10
	rstAux.cursorlocation=3
	rstAux.open "select max(convert(int,substring(ncampo,6,len(ncampo)))) as contador from camposperso with(nolock) where tabla='ARTICULOS' and ncampo like '" & session("ncliente") & "%' and isnull(titulo,'')<>'' ",session("backendlistados")
	if not rstAux.eof and not isnull(rstAux("contador")) then
		num_campos_articulos=rstAux("contador")
		if num_campos_articulos<>0 then
		    si_campo_personalizables=1
		end if
	else
		num_campos_articulos=10
	end if
	rstAux.close
	
	%><input type="hidden" name="si_campo_personalizables" value="<%=EncodeForHtml(si_campo_personalizables)%>"/><%
	
	if num_campos_articulos & ""="" then
		num_campos_articulos=10
	end if
	
	redim tipo_campo_perso(num_campos_articulos+2)
	redim titulo_campo_perso(num_campos_articulos+2)

	if si_campo_personalizables=1 then
        for ki=1 to num_campos_articulos
            nom_campo="campo" & replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
            cadena_campo=replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
		    rstAux.CursorLocation=3
		    rstAux.open "select titulo,tipo from camposperso with(nolock) where ncampo='" & session("ncliente") & cadena_campo & "' and tabla='ARTICULOS' order by SECCIONCP,ncampo,titulo",session("backendlistados")
		    if not rstAux.EOF then
		        tipo_campo_perso(ki)=rstAux("tipo")
		        titulo_campo_perso(ki)=rstAux("titulo")
		    else
		        tipo_campo_perso(ki)=""
		        titulo_campo_perso(ki)=""
		    end if
		    rstAux.Close
        next
    else
        for ki=1 to num_campos_articulos
            tipo_campo_perso(ki)=""
            titulo_campo_perso(ki)=""
        next
    end if
			
		if familia>"" then
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false,"<b>" & LitSubFamilia + ": </b>"
				DrawCelda2 "CELDA", "left", false, NombresEntidades(familia,"familias","codigo","nombre",session("backendlistados"))
			CloseFila
		elseif familia_padre<>"" then
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false,"<b>" & LitFamilia + ": </b>"
				DrawCelda2 "CELDA", "left", false, NombresEntidades(familia_padre,"familias_padre","codigo","nombre",session("backendlistados"))
			CloseFila
		elseif categoria<>"" then
			DrawFila color_blau
				DrawCelda2 "CELDA", "left", false,"<b>" & LitCategoria + ": </b>"
				DrawCelda2 "CELDA", "left", false, NombresEntidades(categoria,"categorias","codigo","nombre",session("backendlistados"))
			CloseFila
		end if
		if tipoarticulo>"" then
		    DrawFila color_blau
			    DrawCelda2 "CELDA  ", "left", false, "<b>"&LitTipoArticulo & "</b>:"
			    DrawCelda2 "CELDA", "left", false, NombresEntidades(tipoarticulo,"tipos_entidades","codigo","descripcion",session("backendlistados"))
		    CloseFila
		end if

        if pnf > "" then
            DrawFila color_blau
                DrawCelda2 "CELDA", "left", true, LITPNF & ": "
                DrawCelda2 "CELDA", "left", false, pnf
            CloseFila
        end if

		if almacen>"" then 
			strselect ="select * from almacenes with(NOLOCK) where codigo='" & almacen & "'"
			rst.cursorlocation=3
			rst.open strselect, session("backendlistados")
			if not rst.eof then
				if si_tiene_modulo_importaciones=0 then
					DrawFila color_blau
      			      DrawCelda2 "CELDA", "left", false,"<b>" & LitAlmacen + ": </b>"
					DrawCelda2 "CELDA", "left", false, EncodeForHtml(rst("descripcion"))
					CloseFila
				end if
			end if
			rst.close
		end if
		if nproveedor>"" then
			strselect ="select razon_social from proveedores with(NOLOCK) where nproveedor='" & session("ncliente")+nproveedor & "'"
			rst.cursorlocation=3
			rst.open strselect, session("backendlistados")
			if not rst.eof then
				DrawFila color_blau
  			        DrawCelda2 "CELDA", "left", false,"<b>" & LitProveedor + ": </b>"
			    	DrawCelda2 "CELDA", "left", false, EncodeForHtml(nproveedor &" - "&rst("razon_social"))
				CloseFila
			end if
			rst.close
		end if		
		if referencia>"" then
			DrawFila color_blau
      	      DrawCelda2 "CELDA", "left", false,"<b>" & LitConref + ": </b>"
			DrawCelda2 "CELDA", "left", false, EncodeForHtml(referencia)
			CloseFila
		end if
		if nombre>"" then
			DrawFila color_blau
			DrawCelda2 "CELDA", "left", false,"<b>" & LitConNombre + ": </b>"
			DrawCelda2 "CELDA", "left", false, EncodeForHtml(nombre)
			CloseFila
		end if
		if stockmayoroigual>"" and mode<>"add" then
			DrawFila color_blau
			    DrawCelda2 "CELDA", "left", false,"<b>" & LitStockMayorOIgual + ": </b>"
			    DrawCelda2 "CELDA", "left", false, EncodeForHtml(stockmayoroigual)
			CloseFila
		end if
		''ricardo 23-3-2009 se gestiona que salgan solamente articulos con stock negativo
		if stocksolonegativo>"" and mode<>"add" then
			DrawFila color_blau
			    DrawCelda2 "CELDA colspan='2'", "left", false,"<b>" & LitArtSoloNeg & "</b>"
			CloseFila
		end if
		if articulosTer>"" and mode<>"add" then
			DrawFila color_blau
			if session("ncliente")=Empresa_BIERZO then
				DrawCelda2 "CELDA colspan='2'", "left", false,"<b>" & LitMostrarSoloArtTerBIERZO & "</b>"
			else
				DrawCelda2 "CELDA colspan='2'", "left", false,"<b>" & LitMostrarSoloArtTer & "</b>"
			end if
			CloseFila
		end if
		if desdeFechaBaja>"" and mode<>"add" then
			DrawFila color_blau
			DrawCelda2 "CELDA", "left", false,"<b>" & LitDesdeFechaBaja + ": </b>"
			DrawCelda2 "CELDA", "left", false, EncodeForHtml(desdeFechaBaja)
			CloseFila
		end if
		if hastaFechaBaja>"" and mode<>"add" then
			DrawFila color_blau
			DrawCelda2 "CELDA", "left", false,"<b>" & LitHastaFechaBaja + ": </b>"
			DrawCelda2 "CELDA", "left", false, EncodeForHtml(hastaFechaBaja)
			CloseFila
		end if
		if articulosBaja>"" and mode<>"add" then
			DrawFila color_blau
			    DrawCelda2 "CELDA colspan=2", "left", false,"<b>" & LitNoMostrarArticulosBaja & "</b>"
			CloseFila
		end if
        if onlyWebStore>"" and mode<>"add" then
			DrawFila color_blau
			    DrawCelda2 "CELDA colspan=2", "left", false,"<b>" & LITONLYECOMMERCEPRODUCTS & "</b>"
			CloseFila
		end if

        if desdeFechaCreacion>"" and mode<>"add" then
			DrawFila color_blau
			DrawCelda2 "CELDA", "left", false,"<b>" & LitDesdeFechaCreacion + ": </b>"
			DrawCelda2 "CELDA", "left", false, EncodeForHtml(desdeFechaCreacion)
			CloseFila
		end if
		if hastaFechaCreacion>"" and mode<>"add" then
			DrawFila color_blau
			DrawCelda2 "CELDA", "left", false,"<b>" & LitHastaFechaCreacion + ": </b>"
			DrawCelda2 "CELDA", "left", false, EncodeForHtml(hastaFechaCreacion)
			CloseFila
		end if

        if consigna>"" and mode<>"add" then
			DrawFila color_blau
			    DrawCelda2 "CELDA", "left", false,"<b>" & LitProductoConsigna + ": </b>"
                if consigna=0 then
			        DrawCelda2 "CELDA", "left", false, LitNo
                else
                    DrawCelda2 "CELDA", "left", false, LitYes
                end if
			CloseFila
		end if

		
        for ki=1 to num_campos_articulos
            nom_campo="campo" & cstr(ki)
            num_campo=replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
		    if request.form("ver_" & nom_campo)>"" then
			    ver_campoN = limpiaCadena(request.form("ver_" & nom_campo))
		    else
			    ver_campoN = limpiaCadena(request.querystring("ver_" & nom_campo))
		    end if
		    if request.form(nom_campo)>"" then
			    valor_campoN = limpiaCadena(request.form(nom_campo))
		    else
			    valor_campoN = limpiaCadena(request.querystring(nom_campo))
		    end if



            if valor_campoN & "">"" then
			        DrawFila color_blau
			            tipoCPN=cstr(tipo_campo_perso(ki))
			            tituloCPN=cstr(titulo_campo_perso(ki))

				        if cstr(tipoCPN)="2" then
					        valor_contiene=""
					        if ucase(valor_campoN)="ON" or cstr(null_s(valor_campoN))="1" then
						        valor_campoN="Sí"
					        else
						        valor_campoN="No"
					        end if
				        elseif cstr(tipoCPN)="3" then
					        valor_contiene=LitCampPersoParCont
					        if valor_campoN & "">"" then
						        valor_campoN=d_lookup("valor","campospersolista","ndetlista=" & valor_campoN & " and ncampo='" & session("ncliente") & num_campo & "' and tabla='ARTICULOS'",session("backendlistados"))
					        else
						        valor_campoN=""
					        end if
				        else
					        valor_contiene=LitCampPersoParCont
				        end if
				        DrawCelda2 "CELDA", "left", false,"<b>" & tituloCPN & valor_contiene & ": </b>"
				        DrawCelda2 "CELDA", "left", false,EncodeForHtml(valor_campoN)
			        CloseFila
                ''end if
            end if
        next%>
    </table>

	<%Alarma "listado_articulo.asp"

  '*********************************************************************************************
  'Se muestran parametros de seleccion
  '*********************************************************************************************

  if mode="add" then
        
        EligeCelda "input", "add", "", "", "", 0, LitConref, "referencia", "", EncodeForHtml(referencia)
        
        EligeCelda "input", "add", "", "", "", 0, LitConNombre, "nombre", "", EncodeForHtml(nombre)
        
        EligeCelda "input", "add", "", "", "", 0, LitDesdeFechaCreacion, "desde_fc", "", EncodeForHtml(desdeFechaCreacion)
        DrawCalendar "desde_fc"
        
        EligeCelda "input", "add", "", "", "", 0, LitHastaFechaCreacion, "hasta_fc", "", EncodeForHtml(hastaFechaCreacion)
        DrawCalendar "hasta_fc"
        
        EligeCelda "input", "add", "", "", "", 0, LitDesdeFechaBaja, "desde_fb", "", EncodeForHtml(desdeFechaBaja)
        DrawCalendar "desde_fb"

        EligeCelda "input", "add", "", "", "", 0, LitHastaFechaBaja, "hasta_fb", "", EncodeForHtml(hastaFechaBaja)
        DrawCalendar "hasta_fb"

        DrawDiv "1", "", ""
            DrawLabel "", "", LitOrdenar%><select class='width60' name="ordenar">
				<option selected="selected" value="REFERENCIA"><%=LITREFLISTART%></option>
				<option value="NOMBRE"><%=LITNOMLISTART%></option>
				<option value="CATEGORIA"><%=LITCATLISTART%></option>
				<option value="FAMILIA"><%=LITFAMLISTART%></option>
				<option value="SUBFAMILIA"><%=LITSUBFAMLISTART%></option>
				<%if si_tiene_modulo_terminales<>0 then%>
					<option value="CODTERMINAL"><%=LITCODTLISTART%></option>
					<option value="NOMTERMINAL"><%=LITNOMTLISTART%></option>
				<%end if%>
			</select>
		<%
        CloseDiv
        DrawDiv "1", "", ""
        DrawLabel "", "", LitProveedor%><input class='width15' type="text" name="nproveedor" value="<%=EncodeForHtml(iif(nproveedor>"",Completar(nproveedor,5,"0"),""))%>" size=5 onchange="TraerProveedor('<%=enc.EncodeForJavascript(mode)%>');"/><a class='CELDAREFB' href="javascript:AbrirVentana('../../compras/proveedores_busqueda.asp?ndoc=articulos_pro&titulo=<%=LITSELPROVLISTART%>&mode=search&viene=articulos_pro','P',<%=AltoVentana%>,<%=AnchoVentana%>)"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input disabled="disabled" class='width40' type="text" name="razon_social" value="<%=EncodeForHtml(tnombre)%>" size="40" /><%
	    CloseDiv


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

            'DrawSelectCelda "CELDA multiple size='8'","165","",0,LitTipoArticulo,"tipoarticulo",rstArtType,tipo_articulo,"codigo","descripcion","",""
            DrawSelectMultipleCelda "CELDA multiple size='8'","165","",0,LitTipoArticulo,"tipoarticulo",rstArtType,tipo_articulo,"codigo","descripcion","",""

            rstArtType.close
            conn.close
            set rstArtType = nothing
            set command = nothing
            set conn = nothing
            
            EligeCelda "input", "add", "", "", "", 0, LITPNF, "pnf", "", pnf

			dim ConfigDespleg (3,13)

				i=0
				ConfigDespleg(i,0)="categoria"
				ConfigDespleg(i,1)=""
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre from categorias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="width60"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitCategoria & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				ConfigDespleg(i,10)=categoria
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				i=1
				ConfigDespleg(i,0)="familia_padre"
				ConfigDespleg(i,1)=""
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
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
				ConfigDespleg(i,2)="8"
				ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(nolock) where codigo like '" & session("ncliente") & "%' order by nombre"
				ConfigDespleg(i,4)=1
				ConfigDespleg(i,5)="width60"
				ConfigDespleg(i,6)="MULTIPLE"
				ConfigDespleg(i,7)="codigo"
				ConfigDespleg(i,8)="nombre"
				ConfigDespleg(i,9)=LitSubFamilia2
				ConfigDespleg(i,10)=familia
				ConfigDespleg(i,11)=""
				ConfigDespleg(i,12)=""

				DibujaDesplegables ConfigDespleg,session("backendlistados")
            
            rstAux.cursorlocation=3
            DrawDiv "1", "", ""
             DrawLabel "", "", LitTarifa
			 if rates&"">"" then
                    rstAux.open " select codigo, descripcion from tarifas with(nolock) where codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' and codigo in "&Replace(Replace(Replace(rates,",","','"),"(","('"),")","')")&" order by descripcion", session("backendlistados")
                    %><select class="width60" multiple="multiple" size="8" name="tarifa" id="tarifa"><%
                        valor=true
                        while not rstAux.eof
                            if valor=true then
                                %><option selected="selected" value="<%=replace(EncodeForHtml(rstAux("codigo")),",","#coma#")%>"><%=EncodeForHtml(rstAux("descripcion")) %></option><%
                                valor=false
                            else
                                %><option value="<%=replace(EncodeForHtml(rstAux("codigo")),",","#coma#")%>"><%=EncodeForHtml(rstAux("descripcion")) %></option><%
                            end if
                            rstAux.movenext
                        wend
                    %></select>
                    <% 
                else
				    rstAux.open " select codigo, descripcion from tarifas with(nolock) where codigo<>'" & session("ncliente") & "BASE' and codigo like '" & session("ncliente") & "%' order by descripcion", session("backendlistados")
                    %><select class="width60"  multiple="multiple" size="8" name="tarifa"><%                    
                            while not rstAux.eof%>
                                <option value="<%=replace(EncodeForHtml(rstAux("codigo")),",","#coma#")%>"><%=EncodeForHtml(rstAux("descripcion"))%></option>
                                <%rstAux.movenext
                            wend%>
                            <option value="" selected="selected"></option>
                        </select>
                    <% 
                end if
			rstAux.close
	    CloseDiv
		if si_tiene_modulo_importaciones<>0 then
			%><input type="hidden" name="almacen" value="<%=EncodeForHtml(almacen)%>"><%
		else
                rstAux.cursorlocation=3
                rstAux.open " select codigo, descripcion from almacenes with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion", session("backendlistados")
                DrawSelectMultipleCelda "","","",0,LitAlmacen,"almacen",rstAux,almacen,"codigo","descripcion","",""
                rstAux.close
		        
		end if
        if si_tiene_modulo_OrCU<>0 then
            DrawDiv "1", "", ""
                DrawLabel "", "", LitProductoConsigna%><select class="width60" name="consigna">
                     <option value="0"><%=LitNo %></option>
                     <option value="1"><%=LitYes %></option>
                     <option selected="selected" value=""></option>
                 </select>
             <%
             CloseDiv
        end if
       
			if stockmayoroigual & ""="" then stockmayoroigual="0"
           	DrawInputCeldaActionDiv "", "", "", "3", 0, LitStocKMayorOIgual, "stockmayoroigual", stockmayoroigual, "onchange", "cambiar_smi()", false
            ''ricardo 23-3-2009 se gestiona que salgan solamente articulos con stock negativo
            DrawDiv "1", "", ""
                 DrawLabel "", "", LitArtSoloNeg%><input  type="checkbox" name="stocksolonegativo" <%=iif(cstr(stocksolonegativo)=-1 or cstr(stocksolonegativo)="True","checked","")%> value="ON" onclick="cambiar_neg()">
			    
		    <%
            CloseDiv
        displayTerminalesFilter = " style='display:none;'"
        if si_tiene_modulo_terminales<>0 then
            displayTerminalesFilter = ""
        end if

			if session("ncliente")=Empresa_BIERZO then
                EligeCelda "check", "add", "", "", "", 0, LitMostrarSoloArtTerBIERZO, "articulosTer", "", cstr(articulosTer)
			else
                EligeCelda "check", "add", displayTerminalesFilter, "", "", 0, LitMostrarSoloArtTer, "articulosTer", "", cstr(articulosTer)
			end if
		
		    EligeCelda "check", "add", "", "", "", 0, LitNoMostrarArticulosBaja, "articulosBaja", "", cstr(articulosBaja)

        displayECommerceFilter = " style='display:none'"
        if si_tiene_modulo_ecommerce<>0 then
            displayECommerceFilter = ""
        end if

        EligeCelda "check", "add", displayECommerceFilter, "", "", 0, LITONLYECOMMERCEPRODUCTS, "onlyWebStore", "", cstr(onlyWebStore)

	if si_campo_personalizables=1 then
		%><hr/>
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitCampPersoListArt%></h6>
        <%
	end if
	rst.cursorlocation=3
	rst.open "select * from camposperso with(NOLOCK) where tabla='ARTICULOS' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("backendlistados")
	if not rst.eof then
		num_campos_existen=rst.recordcount
				num_campo=1
				num_campo2=1
				num_puestos=0
				while not rst.eof
					'if num_campo2>1 and ((num_campo2-1) mod 2)=0 then
					if (num_puestos mod 2)=0 then
					end if
					if rst("titulo") & "">"" then

                           if ((num_puestos-1) mod 2)=0 then
						        end if
						        num_puestos=num_puestos+1
						        %><input type="hidden" name="<%=EncodeForHtml("si_campo" & num_campo)%>" value="1"><%
						        'EligeCelda "input", "add", "", "", "", 0, rst("titulo"), "campo" & num_campo, "", valor_campo_perso
						        valor_campo_perso=""
						        if rst("tipo")=1 then
							        if isNumeric(rst("tamany")) then
								        tamany=rst("tamany")
							        else
								        tamany=1
							        end if
                                    EligeCelda "input", "add", "", "", "", 0, rst("titulo"), "campo" & num_campo, "", EncodeForHtml(valor_campo_perso)
						        elseif rst("tipo")=2 then
							        EligeCelda "check", "add", "", "", "", 0, rst("titulo"), "campo" & num_campo, "", iif(valor_campo_perso="on",-1,0)
						        elseif rst("tipo")=3 then
							        num_campo_str=cstr(num_campo)
							        if len(num_campo_str)=1 then
								        num_campo_str="0" & num_campo_str
							        end if
							        strSelListVal="select ndetlista,valor from campospersolista with(NOLOCK) where tabla='ARTICULOS' and ncampo='" & session("ncliente") & num_campo_str & "' and valor is not null and valor<>'' order by valor,ndetlista"
							        rstAux.cursorlocation=3
							        rstAux.open strSelListVal,session("backendlistados"),adOpenKeyset, adLockOptimistic
							        DrawDiv "1", "", ""
                                        DrawLabel "", "", EncodeForHtml(rst("titulo"))%><select class="width60" name="campo<%=EncodeForHtml(num_campo)%>" >
									        <%encontrado=0
									        while not rstAux.eof
										        if valor_campo_perso & "">"" and isnumeric(valor_campo_perso) then
											        valor_campo_perso_aux=cint(valor_campo_perso)
										        else
											        valor_campo_perso_aux=0
										        end if
										        if valor_campo_perso_aux=cint(rstAux("ndetlista")) then
											        texto_selected="selected"
											        if encontrado=0 then encontrado=1
										        else
											        texto_selected=""
										        end if%>
										        <option value="<%=EncodeForHtml(rstAux("ndetlista"))%>"  <%=texto_selected%> ><%=EncodeForHtml(rstAux("valor"))%></option>
										        <%rstAux.movenext
									        wend%>
									        <option <%=iif(encontrado=1,"","selected")%> value=""></option>
								        </select>
							       <%CloseDiv
							        rstAux.close
						        elseif rst("tipo")=4 then
							        if isNumeric(rst("tamany")) then
								        tamany=rst("tamany")
							        else
								        tamany=1
							        end if
							        EligeCelda "input", "add", "", "", "", 0, EncodeForHtml(rst("titulo")), "campo" & num_campo, "", EncodeForHtml(valor_campo_perso)
						        elseif rst("tipo")=5 then
							        if isNumeric(rst("tamany")) then
								        tamany=rst("tamany")
							        else
								        tamany=1
							        end if
                                    EligeCelda "input", "add", "", "", "", 0, EncodeForHtml(rst("titulo")), "campo" & num_campo, "", EncodeForHtml(valor_campo_perso)
						        end if
					else
						%><input type="hidden" name="<%=EncodeForHtml("si_campo" & num_campo)%>" value="0"/>
						<input type="hidden" name="campo<%=EncodeForHtml(num_campo)%>" value=""/><%
					end if
					%><input type="hidden" name="tipo_campo<%=EncodeForHtml(num_campo)%>" value="<%=EncodeForHtml(rst("tipo"))%>"/>
					<input type="hidden" name="titulo_campo<%=EncodeForHtml(num_campo)%>" value="<%=EncodeForHtml(rst("titulo"))%>"/><%
					rst.movenext
					num_campo=num_campo+1
					if not rst.eof then
						if rst("titulo") & "">"" then
							num_campo2=num_campo2+1
						end if
					end if
				wend
		%>
		<input type="hidden" name="num_puestos" value="<%=EncodeForHtml(num_puestos)%>"/>
		<input type="hidden" name="num_campos" value="<%=EncodeForHtml(num_campos_existen)%>"/><%
	else
		%><input type="hidden" name="num_puestos" value="0"/>
		<input type="hidden" name="num_campos" value="0"/><%
	end if
	rst.close
	%><hr/>
        <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitCamposOpcionales%></h6>
    <%
        
				if si_tiene_modulo_importaciones=0 then
                    EligeCelda "check", "add", "", "", "", 0, LitAlmacen, "ver_almacen", "", cstr(ver_almacen)
				end if
                    EligeCelda "check", "add", "", "", "", 0, LitSubFamilia, "ver_familia", "", cstr(ver_familia)
		            EligeCelda "check", "add", "", "", "", 0, Litdto, "ver_dto", "", cstr(ver_dto) 		
		            EligeCelda "check", "add", "", "", "", 0, LitMargen, "ver_margen", "", cstr(ver_margen)		
                    EligeCelda "check", "add", "", "", "", 0, LitRecargo, "ver_recargo", "", cstr(ver_recargo)
				if si_tiene_modulo_ebesa<> 0 then
		            EligeCelda "check", "add", "", "", "", 0, LitClave, "ver_clave", "", cstr(ver_clave)		
				end if
		            EligeCelda "check", "add", "", "", "", 0, LitIVA, "ver_iva", "", cstr(ver_iva)	
		            EligeCelda "check", "add", "", "", "", 0, LitPvp, "ver_pvp", "", cstr(ver_pvp)		
		            EligeCelda "check", "add", "", "", "", 0, LitDivisa, "ver_divisa", "", cstr(ver_divisa)
                    EligeCelda "check", "add", "", "", "", 0, LitCodBarras, "ver_codbarras", "", cstr(ver_codbarras)
		            EligeCelda "check", "add", "", "", "", 0, LitCoste, "ver_coste", "", cstr(ver_coste)    
        
				if si_tiene_modulo_ebesa<> 0 then
		            EligeCelda "check", "add", "", "", "", 0, LitModifs, "ver_modifs", "", cstr(ver_modifs)			
				end if
		        
                EligeCelda "check", "add", "", "", "", 0, LitStock, "ver_stock", "", cstr(ver_stock)
		        EligeCelda "check", "add", "", "", "", 0, LitSmin, "ver_smin", "", cstr(ver_smin)	
                EligeCelda "check", "add", "", "", "", 0, LitSmax, "ver_smax", "", cstr(ver_smax)        
		        EligeCelda "check", "add", "", "", "", 0, LitReposicion, "ver_reposicion", "", cstr(ver_resposicion)
                EligeCelda "check", "add", "", "", "", 0, LitPrecibir, "ver_precibir", "", cstr(ver_precibir)        
                EligeCelda "check", "add", "", "", "", 0, LitPservir, "ver_pservir", "", cstr(ver_pservir)
                EligeCelda "check", "add", "", "", "", 0, LitP_min, "ver_pmin", "", cstr(ver_pmin)        
                EligeCelda "check", "add", "", "", "", 0, LitCosteMedioListArt, "ver_coste_medio", "", cstr(ver_coste_medio)    
		        EligeCelda "check", "add", "", "", "", 0, LitPvpIva, "ver_pvpiva", "", cstr(ver_pvpiva)
                EligeCelda "check", "add", "", "", "", 0, LitWeight, "ver_peso", "", cstr(ver_peso)
		        EligeCelda "check", "add", "", "", "", 0, LitDesAmpliada, "ver_desAmpliada", "", cstr(ver_desAmpliada)		
                EligeCelda "check", "add", "", "", "", 0, LitTipoArticulo, "ver_tipoArticulo", "", cstr(ver_tipoArticulo)
                EligeCelda "check", "add", "", "", "", 0, LitCodSub, "ver_CodSub", "", cstr(ver_CodSub)        
                EligeCelda "check", "add", "", "", "", 0, LitEmbalaje, "ver_Embalaje", "", cstr(ver_Embalaje)        
                EligeCelda "check", "add", "", "", "", 0, LitMedidaVenta, "ver_medidaventa", "", cstr(ver_medidaventa)
                EligeCelda "check", "add", "", "", "", 0, LitImg1, "ver_Img1", "", cstr(ver_Img1)
                EligeCelda "check", "add", "", "", "", 0, LitImg2, "ver_Img2", "", cstr(ver_Img2)
                EligeCelda "check", "add", "", "", "", 0, LitImg3, "ver_Img3", "", cstr(ver_Img3)
                
        
			if si_tiene_modulo_terminales<>0 then
		            EligeCelda "check", "add", "", "", "", 0, LitCodTerminal, "ver_codTerminal", "", cstr(ver_codTerminal)
		            EligeCelda "check", "add", "", "", "", 0, LitNomTerminal, "ver_nomTerminal", "", cstr(ver_nomTerminal)
				    if si_tiene_modulo_02<>0 or si_tiene_modulo_33<>0 then
		                EligeCelda "check", "add", "", "", "", 0, LitListArtPlu, "ver_PLU", "", cstr(ver_Plu)
					end if
				    if si_tiene_modulo_02<>0 or si_tiene_modulo_33<>0 then
		                EligeCelda "check", "add", "", "", "", 0, LitListArtGrpPlu, "ver_GRPPLU", "", cstr(ver_GRPPLU)
					end if    
    		end if

            rst.cursorlocation=3
            'CAMPOS PERSONALIZABLES EN CAMPOS OPCIONALES DEL LISTADO
			rst.open "select * from camposperso with(NOLOCK) where tabla='ARTICULOS' and ncampo like '" & session("ncliente") & "%' order by ncampo,titulo",session("backendlistados")'',adOpenKeyset, adLockOptimistic
			if not rst.eof then
					num_campo=1
					num_campo2=1
					while not rst.eof
						if num_campo2>1 and ((num_campo2-1) mod 4)=0 then
						end if
						if rst("titulo") & "">"" then
                            EligeCelda "check", "add", "", "", "", 0, rst("titulo"), "ver_campo" & num_campo, "", ""
						else
		                    'EligeCelda "check", "add", "display:none", "", "", 0, "", "", "", ""
						end if
						rst.movenext
						num_campo=num_campo+1
						if not rst.eof then
							if rst("titulo") & "">"" then
								num_campo2=num_campo2+1
							end if
						end if
					wend
			end if
			rst.close
                EligeCelda "check", "add", "", "", "", 0, LITPNF, "ver_pnf", "", cstr(ver_pnf)
            'i(EJM 19/02/07) Incluir columna Letra Margen
				if si_tiene_modulo_ebesa<> 0 then
					EligeCelda "check", "add", "", "", "", 0, LitLeMargen, "ver_Lemargen", "", cstr(ver_Lemargen)
                else
                    ''MPC 17/09/2012 Add new field MEDIDA
				    EligeCelda "check", "add", "", "", "", 0, LitMedida, "ver_medida", "", cstr(ver_medida)
				end if
				'fin(EJM 19/02/07) Incluir columna Letra Margen
        'i(EJM 23/02/07) Apaisado
		%><hr/>
            <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%=LitParResVenGen%></h6>        
        <%
            DrawDiv "1", "", ""
            DrawLabel "", "", LitApaisado%><input type="checkbox" name="apaisado" <%=iif(apaisado="SI" or apaisado="on" or apaisado="true" or apaisado="1","checked","")%>/><%
			CloseDiv
        'fin(EJM 23/02/07) Apaisado
       %>
      <hr/><table><%
		DrawFila color_blau%>
				<td class="CELDABOT" onclick="javascript:sel_todos(true);">
					<%PintarBotonBT LITBOTSELTODO,ImgSelecc_todos,ParamImgSelecc_todos,""%>
				</td>
				<td>&nbsp;</td>
				<td class="CELDABOT" onclick="javascript:sel_todos(false);">
					<%PintarBotonBT LITBOTDSELTODO,ImgDeselecc_todos,ParamImgDeselecc_todos,""%>
				</td>
		<%CloseFila
	%></table><%
   end if

   '*********************************************************************************************
   ' Se muestran los datos de la consulta
   '*********************************************************************************************

	if mode="ver" or mode="edit" then

''ricardo 25-5-2006 comienzo de la select
''usuario,ndocumento,npersona,accion,referencia,nserie,tipo


'Save filters in temporal table 
'***********************************
    strdrop ="if exists (select * from sysobjects where id = object_id('" & session("usuario") & "_temporal') and sysstat " & _
		" & 0xf = 3) drop table [" & session("usuario") & "_temporal]"
		rstAux.open strdrop,session("backendListados"),adUseClient,adLockReadOnly
		if rstAux.state<>0 then rstAux.close

		strselect="CREATE TABLE [" & session("usuario") & "_temporal] (referencia varchar(30),nproveedor varchar(100),nombre varchar(100),familia varchar(1000),familia_padre varchar(1000),categoria varchar(1000),tarifa varchar(1000),tipoarticulo varchar(1000),pnf  varchar(100),almacen varchar(100),consigna tinyint, stockmayoroigual real,stocksolonegativo tinyint,articulosTer  int,articulosBaja int,desde_fb varchar(20),hasta_fb varchar(20),ordenar  varchar(50),h_num_campos_articulos int,desde_fc varchar(20),hasta_fc varchar(20),onlyWebStore tinyint"

        'CAMPOS PERSONALIZADOS
        for x=1 to num_campos_articulos
		   nombrecampo= "campo" & x
		   nombreVerCampo= "ver_campo" & x
		   strselect=strselect & "," & nombrecampo & " varchar(50)"
   		   strselect=strselect & "," & nombreVerCampo & " tinyint"
		next
		strselect=strselect & ",numregs int,num_campos_perso int"
'		
		for x=1 to num_campos_articulos
		   nombrecampo= "tipo_campo_perso" & x
		   strselect=strselect & "," & nombrecampo & " varchar(10)"
		next

        'CAMPOS OPCIONALES DE LISTADO

          strselect=strselect & "," & "ver_almacen tinyint,ver_familia tinyint,ver_coste tinyint,ver_clave tinyint,claveV tinyint,ver_modifs tinyint,ver_Lemargen tinyint,ver_CodSub tinyint,ver_Embalaje tinyint,apaisado tinyint,ver_dto tinyint,ver_recargo tinyint,ver_margen tinyint,ver_pvp tinyint,ver_iva tinyint,ver_divisa tinyint,ver_codbarras tinyint,ver_stock tinyint,ver_smin tinyint, ver_smax tinyint,ver_reposicion tinyint,ver_precibir tinyint,ver_pservir tinyint,ver_pmin tinyint,ver_coste_medio tinyint,ver_pvpiva tinyint,ver_codTerminal tinyint,ver_nomTerminal tinyint,ver_desAmpliada tinyint,ver_tipoArticulo tinyint,ver_pnf tinyint,ver_medida tinyint,ver_peso tinyint,ver_medidaventa tinyint,ver_Img1 tinyint,ver_Img2 tinyint,ver_Img3 tinyint,ver_PLU tinyint,ver_GRPPLU tinyint"

		strselect=strselect & ")"

		rstAux.open strselect,session("backendListados"),adUseClient,adLockReadOnly
		GrantUser session("usuario") & "_temporal", session("backendListados")

		strselect="insert into [" & session("usuario") & "_temporal] (referencia,nproveedor,nombre ,familia ,familia_padre ,categoria ,tarifa ,tipoarticulo ,pnf  ,almacen , consigna, stockmayoroigual ,stocksolonegativo ,articulosTer  ,articulosBaja ,desde_fb ,hasta_fb ,ordenar  ,h_num_campos_articulos ,desde_fc ,hasta_fc, onlyWebStore"

        'CAMPOS PERSONALIZADOS
		for x=1 to num_campos_articulos
		   nombrecampo= "campo"&x
		   nombreVerCampo= "ver_campo" & x
		   strselect=strselect & "," & nombrecampo
   		   strselect=strselect & "," & nombreVerCampo
		next
		strselect=strselect & ",numregs,num_campos_perso"
		for x=1 to num_campos_articulos
		   nombrecampo= "tipo_campo_perso" & x
		   strselect=strselect & "," & nombrecampo
		next

        'se añaden campos opcionales
		strselect=strselect & ", ver_almacen, ver_familia, ver_coste, ver_clave, claveV, ver_modifs, ver_Lemargen, ver_CodSub, ver_Embalaje, apaisado, ver_dto, ver_recargo, ver_margen, ver_pvp, ver_iva, ver_divisa, ver_codbarras, ver_stock, ver_smin, ver_smax, ver_reposicion, ver_precibir, ver_pservir, ver_pmin, ver_coste_medio, ver_pvpiva, ver_codTerminal, ver_nomTerminal, ver_desAmpliada, ver_tipoArticulo, ver_pnf, ver_medida, ver_peso, ver_medidaventa, ver_Img1, ver_Img2, ver_Img3, ver_PLU, ver_GRPPLU"
		

		strselect=strselect & ")"
		strselect=strselect & " values "
		strselect=strselect & "("
        strselect=strselect & "'" & iif(referencia>"",referencia,"") &"','"& iif(nproveedor>"",nproveedor,"") &"','"& iif(nombre>"",nombre,"") &"','"& iif(familia>"",familia,"") &"','"& iif(familia_padre>"",familia_padre,"") &"','"& iif(categoria>"",categoria,"") &"','"& iif(tarifa>"",tarifa,"") &"','"& iif(tipoarticulo>"",tipoarticulo,"") &"','"& iif(pnf>"",pnf,"") &"','"& iif(almacen>"",almacen,"") &"','" & iif(consigna>"",consigna,"") &"','" & iif(stockmayoroigual>"",replace(stockmayoroigual,",","."),"") &"','"& iif(stocksolonegativo="ON",1,0) &"','"& iif(articulosTer="on",1,0) &"','" & iif(articulosBaja="on",1,0)&"','"& iif(desdeFechaBaja>"",desdeFechaBaja,"")&"','"& iif(hastaFechaBaja>"",hastaFechaBaja,"")&"','"& iif(ordenar>"",ordenar,"")&"','"& iif(h_num_campos_articulos>"",h_num_campos_articulos,0)&"','"& iif(desdeFechaCreacion>"",desdeFechaCreacion,"")&"','"& iif(hastaFechaCreacion>"",hastaFechaCreacion,"")&"','" & iif(onlyWebStore="on",1,0)&"'"

'CAMPOS PERSONALIZADOS

        for ki=1 to num_campos_articulos
            nom_campo="campo" & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_campo))>"" then
			    valor_campoN = limpiaCadena(request.form(nom_campo))
                 if valor_campoN="on" then
                    strselect=strselect & ",1"
                end if 
		    else
			   valor_campoN = limpiaCadena(request.querystring(nom_campo))
               strselect=strselect & ",'0'"
		    end if
		    nom_campo_ver="ver_campo" & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_campo_ver))>"" then
			    valor_ver_campoN = limpiaCadena(request.form(nom_campo_ver))
                if valor_ver_campoN="on" then
                    strselect=strselect & ",1"
                end if
		    else
			    valor_ver_campoN = limpiaCadena(request.querystring(nom_campo_ver))
                strselect=strselect & ",0"
		    end if
      next

      strselect=strselect & "," & numregs  & "," & num_campos_articulos
		for x=1 to num_campos_articulos
		   nombrecampoTipo= "tipo_campo_perso"  & cstr(x)
            if enc.EncodeForJavascript(request.form(nom_campo_ver))>"" then
			    valor_tipo_campoN = limpiaCadena(request.form(nombrecampoTipo))
                strselect=strselect & ",'"& valor_ver_campoN &"'"
		    else
			    valor_tipo_campoN = limpiaCadena(request.querystring(nombrecampoTipo))
                strselect=strselect & ",'"& valor_tipo_campoN &"'"
		    end if
		next


		'se añaden campos opcionales
		strselect=strselect & ",'" & iif(ver_almacen>"",1,0)& "','" & iif(ver_familia>"",1,0)& "','" & iif(ver_coste>"",1,0)& "','" & iif(ver_clave>"",1,0)& "','" & iif(claveV>"",1,0)& "','" & iif(ver_modifs>"",1,0)& "','" & iif(ver_Lemargen>"",1,0)& "','" & iif(ver_CodSub>"",1,0)& "','" & iif(ver_Embalaje>"",1,0)& "','" & iif(apaisado>"",1,0)& "','" & iif(ver_dto>"",1,0)& "','" & iif(ver_recargo>"",1,0)& "','" & iif(ver_margen>"",1,0)& "','" & iif(ver_pvp>"",1,0)& "','" & iif(ver_iva>"",1,0)& "','" & iif(ver_divisa>"",1,0)& "','" & iif(ver_codbarras>"",1,0)& "','" & iif(ver_stock>"",1,0)& "','" & iif(ver_smin>"",1,0)& "','" & iif(ver_smax>"",1,0)& "','" & iif(ver_reposicion>"",1,0)& "','" & iif(ver_precibir>"",1,0)& "','"& iif(ver_pservir>"",1,0)& "','"& iif(ver_pmin>"",1,0)& "','"& iif(ver_coste_medio>"",1,0)& "','"& iif(ver_pvpiva>"",1,0)& "','"& iif(ver_codTerminal>"",1,0)& "','"& iif(ver_nomTerminal>"",1,0)& "','"& iif(ver_desAmpliada>"",1,0)& "','"& iif(ver_tipoArticulo>"",1,0)& "','"& iif(ver_pnf>"",1,0)& "','"& iif(ver_medida>"",1,0)& "','"& iif(ver_peso>"",1,0)& "','"& iif(ver_medidaventa>"",1,0)& "','"& iif(ver_Img1>"",1,0)& "','"& iif(ver_Img2>"",1,0)& "','"& iif(ver_Img3>"",1,0)& "','"& iif(ver_PLU>"",1,0)& "','"& iif(ver_GRPPLU>"",1,0)

		strselect=strselect & "')"


        rstAux.cursorlocation=3
		rstAux.open strselect,session("backendListados")



'Audit listing
'****************************
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"inicio_listado_articulos"

		sentido=limpiaCadena(Request.QueryString("sentido"))


'HIDDEN FIELDS
'***************************
		%><input type="hidden" name="referencia" value="<%=EncodeForHtml(referencia)%>"/>
		<input type="hidden" name="nproveedor" value="<%=EncodeForHtml(nproveedor)%>"/>
 		<input type="hidden" name="nombre" value="<%=EncodeForHtml(nombre)%>"/>
		<input type="hidden" name="familia" value="<%=EncodeForHtml(familia)%>"/>
		<input type="hidden" name="familia_padre" value="<%=EncodeForHtml(familia_padre)%>"/>
		<input type="hidden" name="categoria" value="<%=EncodeForHtml(categoria)%>"/>
        <input type="hidden" name="tarifa" value="<%=EncodeForHtml(tarifa)%>"/>
		<input type="hidden" name="tipoarticulo" value="<%=EncodeForHtml(tipoarticulo)%>"/>
        <input type="hidden" name="pnf" value="<%=EncodeForHtml(pnf )%>" />
		<input type="hidden" name="almacen" value="<%=EncodeForHtml(almacen)%>"/>
    	<input type="hidden" name="consigna" value="<%=EncodeForHtml(consigna)%>"/>
		<input type="hidden" name="stockmayoroigual" value="<%=EncodeForHtml(stockmayoroigual)%>"/>
		<input type="hidden" name="stocksolonegativo" value="<%=EncodeForHtml(stocksolonegativo)%>"/>
		<input type="hidden" name="articulosTer" value="<%=EncodeForHtml(articulosTer)%>"/>
		<input type="hidden" name="articulosBaja" value="<%=EncodeForHtml(articulosBaja)%>"/>
		<input type="hidden" name="desde_fb" value="<%=EncodeForHtml(desdeFechaBaja)%>"/>
		<input type="hidden" name="hasta_fb" value="<%=EncodeForHtml(hastaFechaBaja)%>"/>
		<input type="hidden" name="ordenar" value="<%=EncodeForHtml(ordenar)%>"/>
		<input type="hidden" name="h_num_campos_articulos" value="<%=EncodeForHtml(num_campos_articulos)%>"/>
        <input type="hidden" name="desde_fc" value="<%=EncodeForHtml(desdeFechaCreacion )%>" />
        <input type="hidden" name="hasta_fc" value="<%=EncodeForHtml(hastaFechaCreacion)%>"/>
        <input type="hidden" name="onlyWebStore" value="<%=EncodeForHtml(onlyWebStore)%>"/>
        
		<%for ki=1 to num_campos_articulos
            nom_campo="campo" & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_campo))>"" then
			    valor_campoN = limpiaCadena(request.form(nom_campo))
		    else
			    valor_campoN = limpiaCadena(request.querystring(nom_campo))
		    end if
		    nom_campo_ver="ver_campo" & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_campo_ver))>"" then
			    valor_ver_campoN = limpiaCadena(request.form(nom_campo_ver))
		    else
			    valor_ver_campoN = limpiaCadena(request.querystring(nom_campo_ver))
		    end if%>
            <input type="hidden" name="<%=nom_campo%>" value="<%=EncodeForHtml(valor_campoN)%>"/>
            <input type="hidden" name="<%=nom_campo_ver%>" value="<%=EncodeForHtml(valor_ver_campoN)%>"/>
        <%next%>

		<input type="hidden" name="ver_almacen" value="<%=EncodeForHtml(ver_almacen)%>"/>
		<input type="hidden" name="ver_familia" value="<%=EncodeForHtml(ver_familia)%>"/>
		<!-- cag -->
		<input type="hidden" name="ver_coste" value="<%=EncodeForHtml(ver_coste)%>"/>
		<input type="hidden" name="ver_clave" value="<%=EncodeForHtml(ver_clave)%>"/>
		<input type="hidden" name="claveV" value="<%=EncodeForHtml(claveV)%>"/>
		<input type="hidden" name="ver_modifs" value="<%=EncodeForHtml(ver_modifs)%>"/>
		<!-- fin cag -->

        <!--i(EJM 19/02/07)-->
        <input type="hidden" name="ver_Lemargen" value="<%=EncodeForHtml(ver_Lemargen)%>"/>
        <input type="hidden" name="ver_CodSub" value="<%=EncodeForHtml(ver_CodSub)%>"/>
        <input type="hidden" name="ver_Embalaje" value="<%=EncodeForHtml(ver_Embalaje)%>"/>
		<input type="hidden" name="apaisado" value="<%=EncodeForHtml(apaisado)%>"/>
        <!--fin(EJM 19/02/07)-->

		<input type="hidden" name="ver_dto" value="<%=EncodeForHtml(ver_dto)%>"/>
		<!-- cag -->
		<input type="hidden" name="ver_recargo" value="<%=EncodeForHtml(ver_recargo)%>"/>
		<input type="hidden" name="ver_margen" value="<%=EncodeForHtml(ver_margen)%>"/>
		<input type="hidden" name="ver_pvp" value="<%=EncodeForHtml(ver_pvp)%>"/>
		<!-- fin cag -->
		<input type="hidden" name="ver_iva" value="<%=EncodeForHtml(ver_iva)%>"/>
		<input type="hidden" name="ver_divisa" value="<%=EncodeForHtml(ver_divisa)%>"/>
		<input type="hidden" name="ver_codbarras" value="<%=EncodeForHtml(ver_codbarras)%>"/>

		<input type="hidden" name="ver_stock" value="<%=EncodeForHtml(ver_stock)%>"/>
		<input type="hidden" name="ver_smin" value="<%=EncodeForHtml(ver_smin)%>"/>
        <input type="hidden" name="ver_smax" value="<%=EncodeForHtml(ver_smax)%>"/>
		<input type="hidden" name="ver_reposicion" value="<%=EncodeForHtml(ver_reposicion)%>"/>
		<input type="hidden" name="ver_precibir" value="<%=EncodeForHtml(ver_precibir)%>"/>
		<input type="hidden" name="ver_pservir" value="<%=EncodeForHtml(ver_pservir)%>"/>
		<input type="hidden" name="ver_pmin" value="<%=EncodeForHtml(ver_pmin)%>"/>
		<input type="hidden" name="ver_coste_medio" value="<%=EncodeForHtml(ver_coste_medio)%>"/>
		<input type="hidden" name="ver_pvpiva" value="<%=EncodeForHtml(ver_pvpiva)%>"/>
		<input type="hidden" name="ver_codTerminal" value="<%=EncodeForHtml(ver_codTerminal)%>"/>
		<input type="hidden" name="ver_nomTerminal" value="<%=EncodeForHtml(ver_nomTerminal)%>"/>
		<input type="hidden" name="ver_desAmpliada" value="<%=EncodeForHtml(ver_desAmpliada)%>"/>
		<input type="hidden" name="ver_tipoArticulo" value="<%=EncodeForHtml(ver_tipoArticulo)%>"/>
        <input type="hidden" name="ver_pnf" value="<%=EncodeForHtml(ver_pnf)%>" />

        <input type="hidden" name="ver_medida" value="<%=EncodeForHtml(ver_medida)%>" />
        <input type="hidden" name="ver_peso" value="<%=EncodeForHtml(ver_peso)%>" />
        <input type="hidden" name="ver_medidaventa" value="<%=EncodeForHtml(ver_medidaventa)%>" />

         <input type="hidden" name="ver_Img1" value="<%=EncodeForHtml(ver_Img1)%>" />
         <input type="hidden" name="ver_Img2" value="<%=EncodeForHtml(ver_Img2)%>" />
         <input type="hidden" name="ver_Img3" value="<%=EncodeForHtml(ver_Img3)%>" />
		
		<input type="hidden" name="ver_PLU" value="<%=EncodeForHtml(ver_PLU)%>"/>
		<input type="hidden" name="ver_GRPPLU" value="<%=EncodeForHtml(ver_GRPPLU)%>"/>
		<%MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='105'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='105'", DSNIlion)%>
		<input type='hidden' name='maxpdf' value='<%=EncodeForHtml(MAXPDF)%>'/>
		<input type='hidden' name='maxpagina' value='<%=EncodeForHtml(MAXPAGINA)%>'/>
		<%
	'creamos tabla temporal para la primera vez
	numregs=0

	if submode>"" then
		Strselect=""
		'strselect=strselect & "select case WHEN es_padre=1 THEN '<b>' + substring(a.referencia,6,30) + '</b>' ELSE (CASE WHEN ref_padre is not null THEN '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + substring(a.referencia,6,30) ELSE substring(a.referencia,6,30) END) END as RefPantalla"
		'strselect=strselect & ",CASE WHEN es_padre=1 THEN '<b>' + a.nombre + '</b>' ELSE (CASE WHEN ref_padre is not null THEN '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + a.nombre ELSE a.nombre END) END as NomPantalla"
		'strselect=strselect & ",a.referencia,a.nombre "
		'strselect=strselect & ",case when a.cantcompras=0 then 0 else (a.impcompras/a.cantcompras) end as coste_medio,te.descripcion"

		StrCamposOpcionales=""
		'if ver_familia="on" then
		'	StrCamposOpcionales=StrCamposOpcionales & ",isnull(a.familia,'ZZZZZ') as familia,f.nombre as nomfamilia"
		'else
		'	StrCamposOpcionales=StrCamposOpcionales & ",NULL,NULL"
		'end if

		if ver_codbarras="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.cod_barras"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		if ver_dto="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.descuento as dto"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		'cag
		if ver_margen="on"  or ver_clave="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.margen"
			claveV="1"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
			claveV="0"
		end if
        if ver_clave="on" then
			claveV="1"
		else
			claveV="0"
		end if
		'fin cag
		'cag coste y recargo
		if ver_coste="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.importe"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if

		if ver_recargo="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.recargo"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		if ver_modifs="on" then
			modifs="1"
		else
			modifs="0"
		end if

				'''''''''''''''''''''''''''''''''''''''''''''''''fin cag
		if ver_iva="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.iva"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		if ver_pvp="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.pvp"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		if ver_pvpiva="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",(a.pvp+(a.pvp*a.iva)/100) as pvpiva"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		'if ver_divisa="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.divisa"
		'else
		'	StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		'end if
		if ver_codTerminal="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",isnull(ter.codterminal,'') as codterminal"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		if ver_nomTerminal="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",isnull(ter.nomterminal,'') as nomterminal"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		if ver_desAmpliada="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",com.nombreadd"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if
		if ver_tipoarticulo="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",te.descripcion as tipo_articulo"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if

        if ver_pnf = "on" then
            StrCamposOpcionales = StrCamposOpcionales & ",a.pnf"
        else
            StrCamposOpcionales = StrCamposOpcionales & ",NULL"
        end if

        for ki=1 to num_campos_articulos
            nom_campo="campo" & replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
            nom_ver_campo="ver_campo" & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_ver_campo))>"" then
			    valor_ver_campoN = limpiaCadena(request.form(nom_ver_campo))
		    else
			    valor_ver_campoN = limpiaCadena(request.querystring(nom_ver_campo))
		    end if
		    if ucase(valor_ver_campoN)="ON" then
		        StrCamposOpcionales=StrCamposOpcionales & ",isnull(a." & nom_campo & ",'') as " & nom_campo
		    else
		        StrCamposOpcionales=StrCamposOpcionales & ",''"
		    end if
		 next

        'i(EJM 19/02/07) Incluir nuevos campos en el select
		if ver_Lemargen="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",isnull(a.medida,'') as Lemargen"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",''"
		end if

		if ver_CodSub="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",isnull(a.familia,'') as codSub"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",''"
		end if

		if ver_Embalaje="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.ue as embalaje"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",null"
		end if
        'fin(EJM 19/02/07) Incluir nuevos campos en el select

		'if ver_coste_medio="on" then
		'	StrCamposOpcionales=StrCamposOpcionales & ",a.coste_medio"
		'end if
		
		''ricardo 5-12-2008 se mostraran los campos PLU y GRUPO
		if ver_PLU="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.plunum"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",null as plunum"
		end if
		if ver_GRPPLU="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",a.grupo,grpa.descripcion as nomgrupo"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",null as grupo,null as nomgrupo"
		end if

        if ver_medida="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",isnull(a.medida,'') as medida"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if

        if ver_peso="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",isnull(a.weight,'') as weight"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if

        if ver_medidaventa="on" then
			StrCamposOpcionales=StrCamposOpcionales & ",isnull(a.medidaventa,'') as medidaventa"
		else
			StrCamposOpcionales=StrCamposOpcionales & ",NULL"
		end if

      

        ''ricardo 23-3-2009 se gestiona que salgan solamente articulos con stock negativo
		if stocksolonegativo="" and (ver_almacen="on" or ver_stock="on" or ver_smin="on" or ver_smax="on" or ver_reposicion="on" or ver_precibir="on" or ver_pservir="on" or ver_pmin="on") then
			poner_almacenar=1
		elseif stocksolonegativo>"" and (ver_almacen="on" or ver_stock="on" or ver_smin="on" or ver_smax="on" or ver_reposicion="on" or ver_precibir="on" or ver_pservir="on" or ver_pmin="on") then
			poner_almacenar=4
		elseif stockmayoroigual>"" then
			'strwhere = strwhere + " a.referencia in (select distinct articulo from almacenar where articulo=a.referencia and stock>=" & stockmayoroigual
			'if almacen > "" then
			'	strwhere = strwhere + " and almacen='" & almacen & "'"
			'end if
			'strwhere = strwhere & ") and"
			poner_almacenar=2
		elseif stocksolonegativo>"" then
		    poner_almacenar=3
		else
			poner_almacenar=0
		end if

		StrFrom=""
		'strFrom=StrFrom & " from articulos as a"
		'StrFrom=StrFrom & " left outer join familias as f on f.codigo=a.familia"
		'StrFrom=StrFrom & " left outer join tipos_entidades as te on a.tipo_articulo=te.codigo"
		'StrFrom=StrFrom & " left outer join articuloster as ter on ter.referencia=a.referencia"
		'StrFrom=StrFrom & " left outer join articuloscom as com on com.referencia=a.referencia"

		strwhere = " where a.referencia like '" & session("ncliente") & "%' and"

		if referencia > "" then
			strwhere = strwhere + " substring(a.referencia,6,30) like '%" + referencia + "%' and"
		end if
		if nproveedor > "" then
			strwhere = strwhere + " a.referencia in( select articulo from proveer with(nolock) where nproveedor = '" + session("ncliente")+nproveedor + "') and"
		end if		
		if nombre > "" then
			strwhere = strwhere + " a.nombre like '%" + nombre + "%' and"
		end if

		'FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA JCI 20/05/2005 ****************************************
		if familia<>"" then
			if instr(familia,",")>0 then
				strwhere = strwhere + " a.familia in ('" & replace(replace(familia," ",""),",","','") & "') and"
			else
				strwhere = strwhere + " a.familia ='" + familia + "' and"
			end if
		elseif familia_padre<>"" then
			if instr(familia_padre,",")>0 then
				strwhere = strwhere + " a.familia_padre in ('" & replace(replace(familia_padre," ",""),",","','") & "') and"
			else
				strwhere = strwhere + " a.familia_padre ='" + familia_padre + "' and"
			end if
		elseif categoria<>"" then
			if instr(categoria,",")>0 then
				strwhere = strwhere + " a.categoria in ('" & replace(replace(categoria," ",""),",","','") & "') and"
			else
				strwhere = strwhere + " a.categoria ='" + categoria + "' and"
			end if
		end if
		'FIN FILTRADO DE CATEGORIA - FAMILIA - SUBFAMILIA *************************************************

		if tipoarticulo > "" then
            if instr(tipoarticulo,",")>0 then
				strwhere = strwhere + " a.tipo_articulo in ('" & replace(replace(tipoarticulo," ",""),",","','") & "') and"
			else
				strwhere = strwhere + " a.tipo_articulo ='" + tipoarticulo + "' and"
			end if
		end if
		if almacen > "" then
		    ''ricardo 23-3-200 se modifica la consulta para que sea mas rapida
			''strwhere = strwhere + " a.referencia IN (SELECT articulo AS referencia FROM ALMACENAR WHERE almacen = '" + almacen + "') and"
			if poner_almacenar=0 then
			    poner_almacenar=2
		    end if
		end if

        if consigna > "" then
            strwhere = strwhere + " a.consigna = " + consigna +" and "
        end if

        if pnf > "" then
            strwhere = strwhere + " a.pnf like '%" + pnf +"%' and "
        end if

		if articulosTer="on" then
			strwhere = strwhere + " a.loadter='1' and"
		end if

		if articulosBaja="on" then
			strwhere = strwhere + " a.fbaja is null and"
		end if

        if onlyWebStore="on" then
            strwhere = strwhere + " a.IMPR_CATALOGO = 1 and"
        end if

		if desdeFechaBaja>"" then
			strwhere = strwhere + " a.fbaja>='" + desdeFechaBaja +"' and"
		end if
		if hastaFechaBaja>"" then
			strwhere = strwhere + " a.fbaja<='" + hastaFechaBaja +"' and"
		end if

        'MAP 20/12/2012 - Add create date as search filter
        if desdeFechaCreacion>"" then
			strwhere = strwhere + " ((a.fechacreacion>='" + desdeFechaCreacion +" 00:00:00.000' and a.ref_padre is null) or (a.ref_padre is not null and (select a2.fechacreacion from articulos a2 with(nolock) where a2.referencia = a.ref_padre)>='" + desdeFechaCreacion +" 00:00:00.000')) and"
            'strwhere = strwhere + " a.fechacreacion>='" + desdeFechaCreacion +" 00:00:00.000' and"
		end if
		if hastaFechaCreacion>"" then
			strwhere = strwhere + " ((a.fechacreacion<='" + hastaFechaCreacion +" 23:59:59.999' and a.ref_padre is null) or (a.ref_padre is not null and (select a2.fechacreacion from articulos a2 with(nolock) where a2.referencia = a.ref_padre)<='" + hastaFechaCreacion +" 23:59:59.999')) and"
            'strwhere = strwhere + " a.fechacreacion<='" + hastaFechaCreacion +" 23:59:59.999' and"
		end if


        for ki=1 to num_campos_articulos
            nom_campo="campo" & cstr(ki)
            num_campo=replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
            nom_v_campo="campo" & replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_campo))>"" then
			    valor_campoN = limpiaCadena(request.form(nom_campo))
		    else
			    valor_campoN = limpiaCadena(request.querystring(nom_campo))
		    end if
		    if valor_campoN & "">"" then
			    if cstr(tipo_campo_perso(ki))="2" then
				    if ucase(valor_campoN)="ON" or cstr(null_s(valor_campoN))="1" then
					    valor_campoN_where="='1'"
				    else
					    valor_campoN_where="='0'"
				    end if
		        elseif cstr(tipoCPN)="3" then
			        ''if valor_campoN & "">"" then
				    ''    valor_campoN=d_lookup("valor","campospersolista","ndetlista=" & valor_campoN & " and ncampo='" & session("ncliente") & num_campo & "' and tabla='ARTICULOS'",session("backendlistados"))
				    ''    valor_campoN_where=" like '%" & valor_campoN & "%'"
			        ''else
				    ''    valor_campoN=""
			        ''end if
			        valor_campoN_where="='" & valor_campoN & "'"
			    else
				    valor_campoN_where=" like '%" & valor_campoN & "%'"
			    end if
			    strwhere = strwhere + " a." & nom_v_campo & " " & valor_campoN_where & " and"
            end if
        next

		if strwhere="where" then
			strwhere=""
		else
         	strwhere=mid(strwhere,1,len(strwhere)-4) 'Quitamos el último AND
		end if

		strOrder=" order by "
	      if trim(ordenar) = "REFERENCIA" then
   			strOrder = strOrder + " a.referencia,c.nombre,fp.nombre,f.nombre"
		elseif trim(ordenar) = "NOMBRE"	 then
			strOrder = strOrder + " a.nombre,c.nombre,fp.nombre,f.nombre"
		'elseif trim(ordenar) = "FAMILIA" and ver_familia="on" then
			'   strOrder = strOrder + " a.familia"
			'MODIFICADO POR JESUS EL 6/3/02
			'MOTIVO : ¿QUE PASA SI ORDENO POR FAMILIA Y NO LA MUESTRO? --> NO ORDENA POR NADA
			'         ¿QUE PASA CON LOS ARTICULOS QUE NO ESTAN EN NINGUNA FAMILIA SI ORDENO POR FAMILIA? --> NO ORDENA BIEN
			'         SE TIENE QUE PODER ORDENAR POR FAMILIA AUNQUE NO LA MUESTRE
		elseif trim(ordenar) = "SUBFAMILIA" then
			strOrder = strOrder + " f.nombre,fp.nombre,c.nombre,a.referencia"
		elseif trim(ordenar) = "FAMILIA" then
			strOrder = strOrder + " fp.nombre,c.nombre,f.nombre,a.referencia"
		elseif trim(ordenar) = "CATEGORIA" then
			strOrder = strOrder + " c.nombre,fp.nombre,f.nombre,a.referencia"
		elseif trim(ordenar) = "CODTERMINAL" then
			strOrder = strOrder + " right('0000000000000'+ter.codterminal,13),a.referencia,c.nombre,fp.nombre,f.nombre"
		elseif trim(ordenar) = "NOMTERMINAL" then
			strOrder = strOrder + " ter.nomterminal,a.referencia,c.nombre,fp.nombre,f.nombre"
		end if

        strTarifas=""
        if tarifa<>"" then
			if instr(tarifa,",")>0 then
				strTarifas = strTarifas + " ('" & replace(replace(tarifa," ",""),",","'),('") & "')"
			else
				strTarifas = strTarifas + " ('" + tarifa + "')"
			end if
		end if
        strTarifas = replace(strTarifas, "#coma#", ",")
		'cag
		'seleccion="Exec listadoArticulos @poner_almacenar=" & iif(poner_almacenar>"",poner_almacenar,"''") & ",@almacen='" & almacen & "',@stockmayoroigual=" & iif(stockmayoroigual>"",stockmayoroigual,"''") & ",@StrCampOpc='" & replace(StrCamposOpcionales,"'","''") & "',@StrWhere='" & replace(strwhere,"'","''") & "',@StrOrder='" & replace(strOrder,"'","''") & "',@sesion_ncliente='" & session("ncliente") & "',@sesion_usuario='" & session("usuario") & "'"
		 'seleccion="Exec listadoArticulos @poner_almacenar=" & iif(poner_almacenar>"",poner_almacenar,"''") & ",@almacen='" & almacen & "',@stockmayoroigual=" & iif(stockmayoroigual>"",stockmayoroigual,"''") & ",@StrCampOpc='" & replace(StrCamposOpcionales,"'","''") & "',@StrWhere='" & replace(strwhere,"'","''") & "',@StrOrder='" & replace(strOrder,"'","''") & "',@sesion_ncliente='" & session("ncliente") & "',@sesion_usuario='" & session("usuario") & "'"
		 'cadena="Exec listadoArticulos @poner_almacenar=" & iif(poner_almacenar>"",poner_almacenar,"''") & ",@almacen='" & almacen & "',@stockmayoroigual='" & iif(stockmayoroigual>"",stockmayoroigual,"") & "',@StrCampOpc='" & replace(StrCamposOpcionales,"'","''") & "',@StrWhere='" & replace(strwhere,"'","''") & "',@StrOrder='" & replace(strOrder,"'","''") &"',@opcClave='" & claveV &"',@opcModifs='" & modifs & "',@sesion_ncliente='" & session("ncliente") & "'"
		'response.write(cadena)
		'response.end


          'MAP 20/12/2012 - Show Img1, img2, img3 in search results if fields are checked (select in procedure)
        ver_imagenes=""

        if ver_Img1="on" then
			ver_Imagenes=ver_Imagenes+",@img1=1"
        end if
	
        if ver_Img2="on" then
			ver_Imagenes=ver_Imagenes+",@img2=1"
		end if
        if ver_Img3="on" then
			ver_Imagenes=ver_Imagenes+",@img3=1"
		end if

		seleccion="Exec listadoArticulos @poner_almacenar=" & iif(poner_almacenar>"",poner_almacenar,"''") & ",@almacen='" & almacen & "',@stockmayoroigual=" & iif(stockmayoroigual>"",replace(stockmayoroigual,",","."),"NULL") & ",@StrCampOpc='" & replace(StrCamposOpcionales,"'","''") & "',@StrWhere='" & replace(strwhere,"'","''") & "',@StrOrder='" & replace(strOrder,"'","''") &"',@opcClave='" & claveV &"',@opcModifs='" & modifs & "', @strTarifas='" & replace(strTarifas,"'","''") & "',@sesion_ncliente='" & session("ncliente") & "',@sesion_usuario='" & session("usuario") & "'"


         if ver_Imagenes<>"" then

            seleccion=seleccion+ver_Imagenes

         end if

		set conListaArt=Server.CreateObject("ADODB.Connection")
		conListaArt.ConnectionTimeout = 180
		conListaArt.CommandTimeout = 180
		conListaArt.open session("backendlistados")
		conListaArt.execute(seleccion)
		conListaArt.close
		set conListaArt= nothing

		numregs=0
		rst.cursorlocation=3
		rst.open "select count(*) as contador from [" & session("usuario") & "]",session("backendlistados")
		if not rst.eof then
			numregs=null_z(rst("contador"))
		end if
		rst.close
	end if

	%><input type="hidden" name="NumRegs" value="<%=EncodeForHtml(numregs)%>"/><%

	lote=limpiaCadena(Request.QueryString("lote"))

	if lote="" then
		lote=1
	end if

	sentido=limpiaCadena(Request.QueryString("sentido"))
	lotes=(numregs/MAXPAGINA)
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

	set conListaArt= Server.CreateObject("ADODB.Connection")
	conListaArt.open session("backendlistados")
	set rst = conListaArt.execute("EXEC sp_Paginar @Nomtabla='[" & session("usuario") & "]', @Nregs=" & MAXPAGINA & ", @Npag= " & lote)

	if numregs=0 then
		%><input type="hidden" name="nRegsImp" value="0"/>
		<script type="text/javascript" language="javascript">
		    window.alert("<%=LitMsgDatosNoExiste%>");
		    parent.document.location = "../../central.asp?pag1=productos/listados/listado_articulos.asp&pag2=productos/listados/listado_articulos_bt.asp&mode=add";
		</script><%
	else
		 'rst.PageSize=MAXPAGINA
		 'rst.AbsolutePage=lote

		%><hr/><%

		NavPaginas lote,lotes,campo,criterio,texto,1

		 %><br/>
	     <table width='100%' style="border-collapse: collapse;" cellspacing="1" cellpadding="1"><%
    	     'Fila de encabezado
		      DrawFila color_fondo
            ver=""
            'i(EJM 19/02/07)
			if ver_CodSub="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitCodSub & "</b>"
			 ver=ver & "k"
			end if
			'fin(EJM 19/02/07)

		         DrawCelda "DATO","","",0,"<b>" &LitRef
                 
            'i(EJM 16/02/07) Nuevo campo
			if ver_Embalaje="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitEmbalaje & "</b>"
			 ver=ver & "k"
			end if
            'fin(EJM 16/02/07) Nuevo campo

			if ver_familia="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitCategoria & "</b>"
			 DrawCelda "DATO","","",0,"<b>" & LitFamilia & "</b>"
			 DrawCelda "DATO","","",0,"<b>" & LitSubFamilia & "</b>"
			 ver=ver & "333"
			end if
			if ver_codbarras="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitCodBarras & "</b>"
			 ver=ver & "4"
			end if

            DrawCelda "DATO","","",0,"<b>" & LitNombre & "</b>"
			ver=ver &"12"

			if ver_tipoarticulo="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitTipoArticulo & "</b>"
			 ver=ver & "l"
			end if

            IF ver_pnf = "on" then
                DrawCelda "DATO","","",0,"<b>" & LITPNF & "</b>"
                ver=ver& "p"
            end if

			if ver_desAmpliada="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitDesAmpliada & "</b>"
			 ver=ver & "k"
			end if

			'cag clave y modificadores
			if ver_modifs="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitModifs & "</b>"
			 ver=ver & "k"
			end if

			if ver_clave="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitClave & "</b>"
			 ver=ver & "k"
			end if

			'fin cag

			'cag coste y recargo
			if ver_coste="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitCoste & "</b>"
			 ver=ver & "0"
			end if
			'fin cag
			if ver_dto="on" then
 			 DrawCelda "DATO","","",0,"<b>" & Litdto & "</b>"
			 ver=ver & "5"
			end if
			'cag
			if ver_recargo="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitRecargo & "</b>"
			 ver=ver & "y"
			end if
 			if ver_margen="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitMargen & "</b>"
			 ver=ver & "1"
			end if
'			if ver_clave="on" then
' 			 DrawCelda "DATO","","",0,"<b>" & LitClave & "</b>"
'			 ver=ver & "1"
'			end if

			if ver_pvp="on" then
			 'DrawCelda "DATO","","",0,"<b>" & LitPvp & "</b>"
			 Drawcelda2 "DATO", "", true, LitPvp
			 ver=ver & "z"
			end if
			'fin cag

			if ver_iva="on" then
				DrawCelda "DATO","","",0,"<b>" & LitIva & "</b>"
				ver=ver & "6"
			end if
			'cag
			'if ver_pvp="on" then
			' 'DrawCelda "DATO","","",0,"<b>" & LitPvp & "</b>"
			' Drawcelda2 "DATO", "right", true, LitPvp
			' ver=ver & "7"
			'end if
			'fin cag
			if ver_pvpiva="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitPvpIva & "</b>"
			 ver=ver & "h"
			end if
			if ver_divisa="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitDivisa & "</b>"
			 ver=ver & "8"
			end if

			if ver_codTerminal="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitCodTerminal & "</b>"
			 ver=ver & "i"
			end if
			if ver_nomTerminal="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitNomTerminal & "</b>"
			 ver=ver & "j"
			end if

			'if ver_tipoarticulo="on" then
			' DrawCelda "DATO","","",0,"<b>" & LitTipoArticulo & "</b>"
			' ver=ver & "l"
			'end if

			'if ver_desAmpliada="on" then
			' DrawCelda "DATO","","",0,"<b>" & LitDesAmpliada & "</b>"
			' ver=ver & "k"
			'end if

            'i(EJM 19/02/07)
			if ver_Lemargen="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitLeMargen & "</b>"
			 ver=ver & "j"
			end if
			'fin(EJM 19/02/07)
			
			''ricardo 5-12-2008 se mostraran los campos PLU y GRUPO
			if ver_PLU="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitListArtPlu & "</b>"
			 ver=ver & "("
			end if
			if ver_GRPPLU="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitListArtGrpPlu & "</b>"
			 ver=ver & ")"
			end if			

            if ver_medida="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitMedida & "</b>"
			 ver=ver & ")"
			end if
            if ver_peso="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitWeight & "</b>"
			 ver=ver & ")"
			end if
            if ver_medidaventa="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitMedidaVenta & "</b>"
			 ver=ver & ")"
			end if


            'MAP 21/12/2012 - img1, img2, img3 in results
              if ver_Img1="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitImg1 & "</b>"
			 ver=ver & ")"
			end if
               if ver_Img2="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitImg2 & "</b>"
			 ver=ver & ")"
			end if
               if ver_Img3="on" then
 			 DrawCelda "DATO","","",0,"<b>" & LitImg3 & "</b>"
		     ver=ver & ")"
			end if


        for ki=1 to num_campos_articulos
            nom_campo="campo" & replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
            nom_ver_campo="ver_campo" & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_ver_campo))>"" then
			    valor_ver_campoN = limpiaCadena(request.form(nom_ver_campo))
		    else
			    valor_ver_campoN = limpiaCadena(request.querystring(nom_ver_campo))
		    end if
			if ucase(valor_ver_campoN)="ON" then
		        tituloCPN=cstr(titulo_campo_perso(ki))
				 DrawCelda "DATO","","",0,"<b>" & tituloCPN & "</b>"
				 ver=ver & "m"
			end if
		next

			if ver_almacen="on" then
			 if si_tiene_modulo_importaciones=0 then
			 	DrawCelda "DATO","","",0,"<b>" & LitAlmacen & "</b>"
				ver = ver & "9"
			 end if
			end if

			if ver_stock="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitStock & "</b>"
			 ver=ver & "a"
			end if

			if ver_smin="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitSMin & "</b>"
			 ver=ver & "b"
			end if

             if ver_smax="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitSMax & "</b>"
			 ver=ver & "h"
			end if

			if ver_reposicion="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitReposicion & "</b>"
			 ver=ver & "c"
			end if
			if ver_precibir="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitPRecibir & "</b>"
			 ver=ver & "d"
			end if
			if ver_pservir="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitPServir & "</b>"
			 ver=ver & "e"
			end if
			if ver_pmin="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitP_min & "</b>"
			 ver=ver & "f"
			end if
			if ver_coste_medio="on" then
			 DrawCelda "DATO","","",0,"<b>" & LitCosteMedioListArt & "</b>"
			 ver=ver & "g"
			end if

            if tarifa <> "" then
                tarifasSel = split(tarifa,", ",-1,1)
                for i = 0 to ubound(tarifasSel)
                    tarifasSel(i) = replace(tarifasSel(i), "#coma#", ",")
                    DrawCelda "DATO","","",0,"<b>" & EncodeForHtml(d_lookup("descripcion","tarifas","codigo='" & tarifasSel(i) & "'",session("backendlistados"))) & "</b>"
                next
            end if
             
		 CloseFila

		VinculosPagina(MostrarArticulos)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina
		fila=1
		while not rst.EOF and fila<=MAXPAGINA
			CheckCadena rst("referencia")
			'Seleccionar el color de la fila.
			if ((fila+1) mod 2)=0 then
				color=color_blau
				con_negrita=false
			else
				 color=color_blau
				 con_negrita=false
			end if

			DrawFila color
			if mode="edit"  then

			else
                'i(EJM 19/02/07)
				if ver_CodSub="on" then
					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(trimCodEmpresa(rst("codsub")))
				end if
                'fin(EJM 19/02/07)

				DrawCelda "DATO","","",0,Hiperv(OBJArticulos,rst("referencia"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("RefPantalla"),LitVerArticulo)
                
                'i(EJM 19/02/07) Nuevo campo
				if ver_Embalaje="on" then
					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(rst("embalaje"))
				end if
                'fin(EJM 19/02/07) Nuevo campo

				if ver_familia="on" then
					if rst("categoria")="ZZZZZ" then
						DrawCelda "DATO","","",0,"&nbsp;"
					else
						DrawCelda "DATO","","",0,EncodeForHtml(rst("nomcategoria"))
					end if
					if rst("familia_padre")="ZZZZZ" then
						DrawCelda "DATO","","",0,"&nbsp;"
					else
						DrawCelda "DATO","","",0,EncodeForHtml(rst("nomfamilia_padre"))
					end if
					if rst("familia")="ZZZZZ" then
						DrawCelda "DATO","","",0,"&nbsp;"
					else
						DrawCelda "DATO","","",0,EncodeForHtml(rst("nomfamilia"))
					end if
				end if

				if ver_codbarras="on" then
					DrawCelda "DATO","","",0,EncodeForHtml(rst("cod_barras"))
				end if

				DrawCelda "DATO","","",0,EncodeForHtml(rst("NomPantalla"))

                'ejm pruebas
				'if ver_codTerminal="on" then
				'	DrawCelda "DATO","","",0,rst("codterminal")
				'end if

				'if ver_nomTerminal="on" then
				'	DrawCelda "DATO","","",0,rst("nomterminal")
				'end if
                'fin ejm pruebas

				if ver_tipoArticulo="on" then
					DrawCelda "DATO","","",0,EncodeForHtml(rst("tipo_articulo"))
				end if
                
                if ver_pnf = "on" then
                    DrawCelda "DATO","","",0,EncodeForHtml(rst("pnf"))
                end if

				if ver_desAmpliada="on" then
					DrawCelda "DATO","","",0,EncodeForHtml(rst("nombreadd"))
				end if

				if ver_modifs="on" then
					'DrawCelda "DATO ALIGN=LEFT","","",0,null_z(rst("mod6"))&null_z(rst("mod1"))&null_z(rst("mod2"))&null_z(rst("mod3"))&null_z(rst("mod4"))&null_z(rst("mod5"))
					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(null_z(rst("mod1"))&null_z(rst("mod2"))&null_z(rst("mod3"))&null_z(rst("mod4"))&null_z(rst("mod5"))&null_z(rst("mod6")))
				end if  ' del ver_modifs

				'cag
				if ver_clave="on" then
				    ' Inverso del margen
				    claveAux=""

					clvAux=formatnumber(null_z(abs(rst("margen"))),2,-1,0,-1)
					clvAuxAbs= abs(clvAux)
'					si usamos abs perdemos los ceros decimales

					clvAuxC=CStr(clvAux)
					i=len(clvAuxC)
					while i >0
						caracter=mid(clvAuxC,i,1)
						if caracter<>"," and caracter<>"." then
						claveAux= claveAux & caracter
						end if
						i=i-1
					wend
					invPreu=""
					preu=formatnumber(null_z(abs(rst("precioMayor"))),2,-1,0,-1)
					preuAbs=abs(preu)
					preuC=CStr(preu)
					i=len(preuC)
					while i >0
						caracter=mid(preuC,i,1)
						if caracter<>"," and caracter<>"." then
						invPreu= invPreu & caracter
						end if
						i=i-1
					wend
					'digs = CInt(clvAux+preu) mod 100
					digs = (clvAuxAbs+preuAbs) mod 100
					digs= formatnumber(digs,0,-1,0,-1)
					if digs<10 then
					   digitos="0"+digs
					else
					   digitos=digs
					end if
					'Calculo digitos intermedios
					''digs = CInt(clvAux+preu) mod 100
					'digs= formatnumber(digs,0,-1,0,-1)

					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(invPreu&digitos&claveAux)
					'DrawCelda "DATO ALIGN=RIGHT","","",0,invPreu&claveAux
				end if   'end del ver_clave

                'fin cag

				'cag coste y recargo
				if ver_coste="on" then
					DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("coste")),DEC_PREC,-1,0,-1))
				end if

				if ver_dto="on" then
					DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("dto")),rst("ndecimales"),-1,0,-1))
				end if
				'cag coste y recargo
				if ver_recargo="on" then
					'DrawCelda "DATO","","",0,rst("recargo")
					DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("recargo")),rst("ndecimales"),-1,0,-1))
				end if
				if ver_margen="on" then
					DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("margen")),rst("ndecimales"),-1,0,-1))
				end if

				'fin cag
				if ver_pvp="on" then
  			        DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(calcula_precio()), DEC_PREC,-1,0,-1))
				end if
				if ver_iva="on" then
					'DrawCelda "DATO ","","",0,rst("iva")
					DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(rst("iva"))
				end if
				'cag
				'if ver_pvp="on" then
  			    '       DrawCelda "DATO ALIGN=RIGHT","","",0,formatnumber(null_z(calcula_precio()),rst("ndecimales"),-1,0,-1)
				'end if
				'fin cag
				if ver_pvpiva="on" then
					DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("pvpiva")),rst("ndecimales"),-1,0,-1))
				end if

				if ver_divisa="on" then
 				        DrawCelda "DATO","","",0,EncodeForHtml(rst("abreviatura"))
				end if

				if ver_codTerminal="on" then
					DrawCelda "DATO","","",0,EncodeForHtml(rst("codterminal"))
				end if

				if ver_nomTerminal="on" then
					DrawCelda "DATO","","",0,EncodeForHtml(rst("nomterminal"))
				end if

                'i(EJM 19/02/07) Posición anterior de los campos
				'if ver_tipoArticulo="on" then
				'	DrawCelda "DATO","","",0,rst("tipo_articulo")
				'end if

				'if ver_desAmpliada="on" then
				'	DrawCelda "DATO","","",0,rst("nombreadd")
				'end if
                'fin(EJM 19/02/07) Posición anterior de los campos

                'i(EJM 19/02/07)
				if ver_Lemargen="on" then
					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(rst("lemargen"))&"&nbsp;"
				end if
                'fin(EJM 19/02/07)

			    ''ricardo 5-12-2008 se mostraran los campos PLU y GRUPO
			    if ver_PLU="on" then
 			        DrawCelda "DATO","","",0,EncodeForHtml(rst("plunum"))
			    end if
			    if ver_GRPPLU="on" then
 			        DrawCelda "DATO","","",0,EncodeForHtml(iif(rst("grupo")>"",trimCodEmpresa(rst("grupo")) & " - " & rst("nomgrupo"),""))
			    end if
                if ver_medida="on" then
					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(rst("medida"))
				end if
                if ver_peso="on" then
					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(rst("weight"))
				end if
                if ver_medidaventa="on" then
					DrawCelda "DATO ALIGN=LEFT","","",0,EncodeForHtml(d_lookup("descripcion", "medidas", "codigo='"&rst("medidaventa")&"'", session("dsn_cliente")))
				end if

               'MAP 21/12/2012 - Img1,Img2,Img3
                if ver_Img1="on" then

                    if rst("tipo_foto") & ""="" then
                        response.write("<td class='TDBORDECELDA7'>")
						response.write("<input type='hidden' name='mostrar_foto1_"& fila-1 &"' value='0'/>")
						response.write("</td>")

                       else
                           num_foto=1
                           response.write("<td class='TDBORDECELDA7' style='width:100;height:50' align='center' valign='center'><div id='capa_foto1_"& fila-1 & "' style='display:none;width:100; height:50;'>")
					       response.write("<img src='../smuestra.asp?ref=" & EncodeForHtml(rst("referencia")) & "&num_foto="&num_foto&"&empresa=" & session("ncliente") & "'  border='0' alt='' title='' name='foto_articulo1_"& fila-1 &"'/>")
					       response.write("</div><input type='hidden' name='foto_art1_w_"& fila-1 &"' value=''/>")
					       response.write("<input type='hidden' name='foto_art1_h_"& fila-1 &"' value=''/>")
					       response.write("<input type='hidden' name='mostrar_foto1_"& fila-1 &"' value='1'/>")
					       response.write("</td>")

                    end if
				end if
                if ver_Img2="on" then

                      if rst("tipo_foto2") & ""="" then
                        response.write("<td class='TDBORDECELDA7'>")
						response.write("<input type='hidden' name='mostrar_foto2_"& fila-1 &"' value='0'/>")
						response.write("</td>")

                       else
                        num_foto=2
					    response.write("<td class='TDBORDECELDA7' style='width:100;height:50' align='center' valign='center'><div id='capa_foto2_"& fila-1 &"' style='display:none;width:100; height:50;'>")
					    response.write("<img src='../smuestra.asp?ref=" & EncodeForHtml(rst("referencia")) & "&num_foto="&num_foto&"&empresa=" & session("ncliente") & "'  border='0' alt='' title='' name='foto_articulo2_"& fila-1 &"'/>")
					    response.write("</div><input type='hidden' name='foto_art2_w_"& fila-1 &"' value=''/>")
					    response.write("<input type='hidden' name='foto_art2_h_"& fila-1 &"' value=''/>")
					    response.write("<input type='hidden' name='mostrar_foto2_"& fila-1 &"' value='1'/>")
					    response.write("</td>")

                    end if
				end if
                if ver_Img3="on" then
                    if rst("tipo_foto3") & ""="" then
                        response.write("<td class='TDBORDECELDA7'>")
						response.write("<input type='hidden' name='mostrar_foto3_"& fila-1 &"' value='0'/>")
						response.write("</td>")

                       else
                        num_foto=3
					    response.write("<td class='TDBORDECELDA7' style='width:100;height:50' align='center' valign='center'><div id='capa_foto3_"& fila-1 &"' style='display:none;width:100; height:50;'>")
					    response.write("<img src='../smuestra.asp?ref=" & EncodeForHtml(rst("referencia")) & "&num_foto="&num_foto&"&empresa=" & session("ncliente") & "'  border='0' alt='' title='' name='foto_articulo3_"& fila-1 &"'/>")
					    response.write("</div><input type='hidden' name='foto_art3_w_"& fila-1 &"' value=''/>")
					    response.write("<input type='hidden' name='foto_art3_h_"& fila-1 &"' value=''/>")
					    response.write("<input type='hidden' name='mostrar_foto3_"& fila-1 &"' value='1'/>")
					    response.write("</td>")

                    end if
				end if


        for ki=1 to num_campos_articulos
            nom_campo="campo" & replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
            nom_v_campo=replace(space(2-len(cstr(ki)))," ","0") & cstr(ki)
            nom_ver_campo="ver_campo" & cstr(ki)
		    if enc.EncodeForJavascript(request.form(nom_ver_campo))>"" then
			    valor_ver_campoN = limpiaCadena(request.form(nom_ver_campo))
		    else
			    valor_ver_campoN = limpiaCadena(request.querystring(nom_ver_campo))
		    end if
			if ucase(valor_ver_campoN)="ON" then
			    tipoCPN=cstr(tipo_campo_perso(ki))
					if cstr(tipoCPN)="2" then
						if ucase(rst(nom_campo))="ON" or cstr(null_s(rst(nom_campo)))="1" then
							valor_campoN="Sí"
						else
							valor_campoN="No"
						end if
					elseif cstr(tipoCPN)="3" then
						if rst(nom_campo) & "">"" then
							valor_campoN=d_lookup("valor","campospersolista","ndetlista=" & rst(nom_campo) & " and ncampo='" & session("ncliente") & nom_v_campo & "' and tabla='ARTICULOS'",session("backendlistados"))
						else
							valor_campoN=""
						end if
					else
						valor_campoN=rst(nom_campo)
					end if
					DrawCelda "DATO","","",0,iif(valor_campoN>"",EncodeForHtml(valor_campoN),"&nbsp;")
			end if
		next
             
				salto=0
                
                '---------------------------
				referencia_old=rst("referencia")
				referencia_new=rst("referencia")
				he_entrado_alm=0
                flag = 0
                flagaux = 0
				while referencia_old=referencia_new and not rst.eof
					he_entrado_alm=1
                    
				    if ver_almacen="on" then
                        flagaux = 1
					    if salto=1 then
						    lk=instr(1,ver,"9",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
		      	  	    DrawCelda "DATO","","",0,EncodeForHtml(rst("nomalmacen"))
                       
				    end if
				    if ver_stock="on" then
					    flagaux = 1
                        if salto=1 then
						    lk=instr(1,ver,"a",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
		        		    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("stock")),dec_cant,-1,0,-1))&"&nbsp;"
				    end if
				    if ver_smin="on" then
                        flagaux = 1
                	    if salto=1 then
 						    lk=instr(1,ver,"b",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
					    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("stock_minimo")),dec_cant,-1,0,-1))
				    end if
                    if ver_smax="on" then
                        flagaux = 1
   					    if salto=1 then
						    lk=instr(1,ver,"h",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
					    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("stock_maximo")),dec_cant,-1,0,-1))
				    end if
				    if ver_reposicion="on" then
                        flagaux = 1
					    if salto=1 then
						    lk=instr(1,ver,"c",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
					    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("reposicion")),dec_cant,-1,0,-1))
				    end if
				    if ver_precibir="on" then
                        flagaux = 1
					    if salto=1 then
						    lk=instr(1,ver,"d",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
					    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("p_recibir")),dec_cant,-1,0,-1))
				    end if
				    if ver_pservir="on" then
                        flagaux = 1
					    if salto=1 then
						    lk=instr(1,ver,"e",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
					    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("p_servir")),dec_cant,-1,0,-1))
				    end if
				    if ver_pmin="on" then
                        flagaux = 1
					    if salto=1 then
						    lk=instr(1,ver,"f",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
					    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(null_z(rst("p_min")),dec_cant,-1,0,-1))
				    end if
				    if ver_coste_medio="on" then
                        flagaux = 1
					    if salto=1 then
						    lk=instr(1,ver,"g",1)
						    for ki=1 to (lk-1)
							    DrawCelda "DATO","","",0,""
						    next
					    end if
					    salto=0
					    DrawCelda "DATO ALIGN=RIGHT","","",0,EncodeForHtml(formatnumber(rst("coste_medio"),rst("ndecimales"),-1,0,-1))
				    end if
                    if tarifa <> "" then
                        tarifasSel = split(tarifa,", ",-1,1)
                        for i = 0 to ubound(tarifasSel)
                            tarifasSel(i) = replace(tarifasSel(i), "#coma#", ",")
                            pvp = d_lookup("pvp","[" & session("usuario") & "_temp2]","referencia='" & rst("referencia") & "' and tarifa='" & tarifasSel(i) & "'",session("backendlistados"))
                            DrawCelda "DATO ALIGN=RIGHT","","",0,"" & EncodeForHtml(pvp)
                        next
                    end if
                    if flag = 1 then
                        if flagaux = 0 then
                            DrawCelda "DATO","","",0,""
                            DrawCelda "DATO","","",0,""
                        end if
                        
                    end if
                    flag = 1
					CloseFila
					salto=1
					rst.movenext
					if not rst.eof then
						referencia_new=rst("referencia")
					end if
				wend
			end if

			fila=fila+1
			if he_entrado_alm=0 then
				if not rst.eof then
					rst.MoveNext
				end if
			end if
		wend

		rst.close
		set rst=nothing
		conListaArt.close
		set conListaArt=nothing

        rstAux6.cursorlocation=3
		rstAux6.open "select max(id) as contador from [" & session("usuario") & "]",session("backendListados")
		if not rstAux6.eof then
			totalreg=rstAux6("contador")
		else
			totalreg=0
		end if
		
		%><input type="hidden" name="fila_puestos" value="<%=fila-1%>"/>
		<!--</tbody>-->
		<%if lote >= lotes then
			DrawFila ""
			CloseFila
			DrawFila ""
			DrawCelda "DATO ALIGN=RIGHT","","",0,"<b>" & "Total " & "</b>"
			DrawCeldaSpan "DATO ALIGN=RIGHT","","",0,"<b>" & EncodeForHtml(totalreg) & "</b>",1
			CloseFila
		end if
        rstAux6.close
        %>


	     </table>
		<br/>
		<input type="hidden" name="nRegsImp" value="<%=fila-1%>"/><%
              NavPaginas lote,lotes,campo,criterio,texto,2
	end if


''ricardo 25-5-2006 comienzo de la select
''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"fin_listado_articulos"
   end if%>
<iframe name="marcoExportar" style="display:none" src="listado_articulos_exportar.asp?mode=ver" frameborder="0" width="500" height="200"></iframe>
</form>
<%set rstAux=nothing
  set rstAux2=nothing
  set rstAux3=nothing
  set rstAux4=nothing
  set rstAux5=nothing
  set rstAux6=nothing
  set rst=nothing
  set rstPrecio=nothing
  set rstModifs=nothing
end if%>
</body>
</html>
