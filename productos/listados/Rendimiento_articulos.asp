<%@ Language=VBScript %>
<%
''ricardo 16-11-2007 se cambia la dsn desde dsncliente a backendlistados
%>
<%

''ricardo 7/3/2003
''se pone  'or SUM(DF.cantidad)=0' a la condicion where df.coste=0 then 100
''ya que como strcostes=df.coste*sum(df.cantidad)
''y se dividia por strcostes
''podia llegar a ocurrir que sum(df.cantidad) sea cero,
''sobre todo cuando se suman cantidades positivas y negativas

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
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
function EncodeForJS(data)
	if data & "" <>"" then
	  EncodeForJS = enc.EncodeForJavascript(data)
	else
	  EncodeForJS = data
	end if
end function
function pintar_saltos_nuevo(texto)
	texto=Replace(texto,"&#10;","")
	texto=Replace(texto,"&#13;","<br>")
	pintar_saltos_nuevo=texto
end function
%>

<html lang="<%=session("lenguaje")%>">
<head>
<title>Ilion SaaS</title>
<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
<!--#include file="../../constantes.inc" -->
<!--#include file="../../cache.inc" -->
<!--#include file="../../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then%>
<!--#include file="../../ilion.inc" -->
<!--#include file="../../mensajes.inc" -->
<!--#include file="../../adovbs.inc" -->
<!--#include file="../../varios.inc" -->
<!--#include file="../../ico.inc" -->
<!--#include file="../../tablasResponsive.inc" -->
<!--#include file="../../modulos.inc" -->
<!--#include file="../../catFamSubResponsive.inc" -->
<!--#include file="../../styles/formularios.css.inc" -->
<!--#include file="rendimiento_articulos.inc" -->
<!--#include file="../../js/generic.js.inc"-->
<!--#include file="../../js/calendar.inc" -->
<link rel="stylesheet" href="../../pantalla.css" media="SCREEN"/>
<link rel="stylesheet" href="../../impresora.css" media="PRINT"/>
</head>
<%si_tiene_modulo_proyectos=ModuloContratado(session("ncliente"),ModProyectos)
si_tiene_modulo_comercial=ModuloContratado(session("ncliente"),ModComercial)
si_tiene_modulo_mantenimiento=ModuloContratado(session("ncliente"),ModMantenimiento)
tiene_modulo_tiendas= ModuloContratado(session("ncliente"),ModTiendas)
has_module_advanced_management=ModuloContratado(session("ncliente"),ModCcostes_Gestion)
has_module_postventa=ModuloContratado(session("ncliente"),ModPostVenta)%>
<script language="javascript" type="text/javascript" src="../../jfunciones.js"></script>
<script language="javascript" type="text/javascript">
function cambiarNoDetalle(que_agrupacion){
    if (que_agrupacion=="3"){
        if (document.rendimiento_articulos.h_rendimiento.value=="articulos"){
            document.getElementById("ordarticulos").disabled=false;
            que_agrupacion=1;
        }
        else{
            document.getElementById("orddocumentos").disabled=false;
            que_agrupacion=2;
        }
    }

    if (que_agrupacion=="1"){
        que_he_elegido=document.rendimiento_articulos.agruparart.options[document.rendimiento_articulos.agruparart.selectedIndex].value;
        if (que_he_elegido=="CLIENTE" || que_he_elegido=="COMERCIAL" || que_he_elegido=="RESPONSABLE" || que_he_elegido=="PROYECTO"){
            document.getElementById("id_nodetallado").style.display="";
            document.getElementById("id_nodetallado2").style.display="none";
            if (document.rendimiento_articulos.nodetallado.checked==true){
                document.getElementById("ordarticulos").disabled=true;
            }
        }
        else{
            document.getElementById("id_nodetallado").style.display="none";
            document.getElementById("id_nodetallado2").style.display="";
            document.getElementById("ordarticulos").disabled=false;
        }
    }
    if (que_agrupacion=="2"){
        que_he_elegido=document.rendimiento_articulos.agrupardoc.options[document.rendimiento_articulos.agrupardoc.selectedIndex].value;
        if (que_he_elegido=="CLIENTE" || que_he_elegido=="COMERCIAL" || que_he_elegido=="RESPONSABLE" || que_he_elegido=="PROYECTO"){
            document.getElementById("id_nodetallado").style.display="";
            document.getElementById("id_nodetallado2").style.display="none";
            if (document.rendimiento_articulos.nodetallado.checked==true){
                document.getElementById("orddocumentos").disabled=true;
            }
        }
        else{
            document.getElementById("id_nodetallado").style.display="none";
            document.getElementById("id_nodetallado2").style.display="";
            document.getElementById("orddocumentos").disabled=false;
        }
    }
}

function tier2Menu(modo) {
    if (modo=="articulos"){
        document.getElementById("agrarticulos").style.display="";
        document.getElementById("bajaarticulos").style.display="";
        document.getElementById("agrdocumentos").style.display="none";
        document.getElementById("agrdocumentos2").style.display="none";
        <%if si_tiene_modulo_mantenimiento<>0 then%>
			document.all("agrdocumentos3").style.display="none";
        <%end if%>
		ordarticulos.style.display="";
        orddocumentos.style.display="none";
        filacostes.style.display="";
        document.rendimiento_articulos.h_rendimiento.value="articulos";

        cambiarNoDetalle('3');
		
    }
    else {
        document.getElementById("agrarticulos").style.display="none";
        document.getElementById("bajaarticulos").style.display="none";
        document.getElementById("agrdocumentos").style.display="";
        document.getElementById("agrdocumentos2").style.display="";
        <%if si_tiene_modulo_mantenimiento<>0 then%>
			document.getElementById("agrdocumentos3").style.display="";
        <%end if%>
		ordarticulos.style.display="none";
        orddocumentos.style.display="";
        filacostes.style.display="none";
        document.rendimiento_articulos.h_rendimiento.value="documentos";
        cambiarNoDetalle('3');
    }
}

//Desencadena la búsqueda del cliente cuya referencia se indica
function TraerCliente(mode) {
    document.rendimiento_articulos.action="rendimiento_articulos.asp" +
	"?ncliente=" + document.rendimiento_articulos.ncliente.value +
	"&mode=" + mode;

    document.rendimiento_articulos.submit();
}

//Desencadena la búsqueda del proveedor cuya referencia se indica
function TraerProveedor(mode) {
	document.rendimiento_articulos.action="rendimiento_articulos.asp" +
	"?nproveedor=" + document.rendimiento_articulos.nproveedor.value +
	"&mode=" + mode;

	document.rendimiento_articulos.submit();
}

//Desencadena la búsqueda del centro cuya referencia se indica
function GetCenter(mode, type)
{
    if (document.rendimiento_articulos.ncliente.value != "")
    {
        if(type == 0)
        {
            document.rendimiento_articulos.action="rendimiento_articulos.asp" +
	        "?ncenter=" + document.rendimiento_articulos.ncenter.value +
	        "&mode=" + mode;

            document.rendimiento_articulos.submit();
        }
        if (type == 1)
        {
            AbrirVentana('../../mantenimiento/centros_buscar.asp?ndoc=rendimiento_articulos&titulo=<%=LitSelCentro%>&mode=search&viene=rendimiento_articulos&viene2=rendimiento_articulos&ncliente=' + document.rendimiento_articulos.ncliente.value,'P',<%=altoventana%>,<%=anchoventana%>);
        }
    }
    else
    {
        document.rendimiento_articulos.ncenter.value = "";
        alert("<%=LitMsgClienteNoNulo%>");
    }
}
</script>
<body class="BODY_ASP">
<iframe name="frameExportar" style='display:none;' src="rendimiento_articulos_pdf.asp?mode=ver" frameborder='0' width='500' height='200'></iframe>
<%
'********************************************************************************************************
sub MuestraListadoCat()
	regcatdiv=0
	rstAux.cursorlocation=3
	rstAux.open "select distinct nomcategoria,divisa from [" & session("usuario") & "] group by nomcategoria,divisa",session("backendListados")
    if not rstAux.eof then
	    regcatdiv=rstAux.recordcount
    end if
	rstAux.close
	NUMREGISTROS=regcatdiv

	lote=limpiaCadena(Request.QueryString("lote"))

	if lote="" then
		lote=1
	end if
	sentido=limpiaCadena(Request.QueryString("sentido"))
	lotes=(NUMREGISTROS/MAXPAGINA)
	if lotes>(NUMREGISTROS/MAXPAGINA) then
		lotes=(NUMREGISTROS/MAXPAGINA)+1
	else
		lotes=(NUMREGISTROS/MAXPAGINA)
	end if

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

	rst.PageSize=MAXPAGINA
	rst.AbsolutePage=lote
	'-----------------------------------------'

	NavPaginas lote,lotes,campo,criterio,texto,1%>
	<hr/>
	<table width='100%' border='0' style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
		<%'Fila de encabezado'
		DrawFila color_fondo
			DrawCelda "TDBORDECELDAB7","","",0,LitCategoria
			DrawCelda "TDBORDECELDAB7","","",0,"Num. " & LitFamilias
			DrawCelda "TDBORDECELDAB7","","",0,"Num. " & LitSubfamilias
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitUnidades
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCostes
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBeneficio
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBCompras
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenUnidades
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenBeneficio
			Cambio=true
			CambioFamPadre=true
		CloseFila

		fila=1
		DivisaAnt=""
		AbrAnt=""
		DecAnt=""
		Col1Anterior=""
		datosCollAnterior=""
		AgrSecundariaAnt=""
		Tunidades=0
		Tventas=0
		Tbenef=0

        rstAux.CursorLocation=3
		rstAux.open "select nomcategoria,divisa,sum(unidades) as TotalUnidades,sum(ventas) as Totalventas,sum(beneficio) as Totalbeneficio  from [" & session("usuario") & "] group by nomcategoria,divisa",session("backendListados")
		while not rstAux.eof
			Tunidades=Tunidades + rstAux("TotalUnidades")
			Tventas=Tventas + CambioDivisa(rstAux("totalventas"),rstAux("divisa"),MB)
			Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),rstAux("divisa"),MB)
			rstAux.movenext
		wend
		rstAux.close

		while not rst.EOF and fila<=MAXPAGINA
			AgrSecundaria=rst("familia_padre")
			nomCol1="categoria"
			Col1Actual=rst("categoria")
			datosCol1=rst("nomcategoria")
			DivisaActual=rst("divisa")

			if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
				DrawFila color_blau
					rstAux.cursorlocation=3
					rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    Reg=rstAux.recordcount
                    end if
					rstAux.close
					rstAux.CursorLocation=3
					rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    RegFamPadre=rstAux.recordcount
                    end if
					rstAux.close
					rstAux.CursorLocation=3
					rstAux.open "select (sum(beneficio)/" & replace(Tbenef,",",".") & ")*100 as porcenbeneficioCategoriaTotal,(sum(unidades)/" & replace(Tunidades,",",".") & ")*100 as porcenUnidadesCategoriaTotal,(sum(ventas)/" & replace(Tventas,",",".") & ")*100 as porcenVentasCategoriaTotal,count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    DrawCelda "TDBORDECELDA7 align='LEFT'","","",0 ,EncodeForHtml(datosCollAnterior)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(RegFamPadre)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
    					DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenVentasCategoriaTotal"),2,-1,0,-1) & "%")
	    				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenUnidadesCategoriaTotal"),2,-1,0,-1) & "%")
		    			DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenBeneficioCategoriaTotal"),2,-1,0,-1) & "%")
                    else
                        DrawCelda "TDBORDECELDA7 align='LEFT'","","",0 ,EncodeForHtml(datosCollAnterior)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(RegFamPadre)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
    					DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
	    				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
		    			DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
                    end if
					rstAux.close
					fila=fila+1
					Cambio=true
				CloseFila
			end if

			AgrSecundariaAnt=rst("familia_padre")
			DivisaAnt=rst("divisa")
			AbrAnt=rst("abreviatura")
			DecAnt=rst("ndecimales")
			Col1Anterior=rst("categoria")
			if verCodCFS>"" and trimCodEmpresa(rst("categoria"))<>"XxYxZ" then
				datosCollAnterior=trimCodEmpresa(rst("categoria")) & " - " & rst("nomcategoria")
			else
				datosCollAnterior= rst("nomcategoria")
			end if

			rst.movenext
			Cambio=false
			CambioFamPadre=false
		wend
		if rst.eof then
			DrawFila color_blau
				rstAux.cursorlocation=3
				rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
				Reg=rstAux.recordcount
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
				RegFamPadre=rstAux.recordcount
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select (sum(beneficio)/" & replace(Tbenef,",",".") & ")*100 as porcenbeneficioCategoriaTotal,(sum(unidades)/" & replace(Tunidades,",",".") & ")*100 as porcenUnidadesCategoriaTotal,(sum(ventas)/" & replace(Tventas,",",".") & ")*100 as porcenVentasCategoriaTotal,count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
				DrawCelda "TDBORDECELDA7 align='LEFT'","","",0, EncodeForHtml(datosCollAnterior)
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(RegFamPadre)
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(Reg)
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")

				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenVentasCategoriaTotal"),2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenUnidadesCategoriaTotal"),2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenBeneficioCategoriaTotal"),2,-1,0,-1) & "%")
				rstAux.close
			CloseFila
			'FILA DEL GRAN TOTAL
			DrawFila color_fondo
				rst.movefirst
				DivisaAnt=""
				Col1Anterior=""
				Reg=0
				RegFamPadre=0
				Tunidades=0
				Tcostes=0
				Tventas=0
				Tbenef=0
				while not rst.EOF
					nomCol1="categoria"
					Col1Actual=rst("categoria")

					DivisaActual=rst("divisa")

					if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
						rstAux.cursorlocation=3
						rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
						Reg=Reg + rstAux.recordcount
						rstAux.close
						rstAux.CursorLocation=3
						rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
						RegFamPadre=RegFamPadre + rstAux.recordcount
						rstAux.close
						rstAux.CursorLocation=3
						rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
						Tunidades=Tunidades + rstAux("TotalUnidades")
						Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
						Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
						Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
						rstAux.close
					end if
					DivisaAnt=rst("divisa")
					Col1Anterior=Col1Actual
					rst.movenext
				wend

				rstAux.cursorlocation=3
				rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Reg=Reg + rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    RegFamPadre=RegFamPadre + rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select distinct categoria from [" & session("usuario") & "]",session("backendListados")
                if not rstAux.eof then
				    RegCategoria=rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Tunidades=Tunidades + rstAux("TotalUnidades")
				    Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
				    Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
				    Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
                end if
				rstAux.close

				if Tcostes>0 then BenCompras=(Tbenef*100)/Tcostes
				if Tventas>0 then BenVentas=(Tbenef*100)/Tventas

				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCategorias & " : " & EncodeForHtml(RegCategoria)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(RegFamPadre)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tunidades,DEC_CANT,-1,0,-1))
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tcostes,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tventas,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tbenef,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenCompras,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenVentas,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
			CloseFila
		end if
		NavPaginas lote,lotes,campo,criterio,texto,2
		rst.close%>
	</table>
<%end sub

'********************************************************************************************************
sub MuestraListadoFam()
	lote=limpiaCadena(Request.QueryString("lote"))

	if lote="" then
		lote=1
	end if
	sentido=limpiaCadena(Request.QueryString("sentido"))
	lotes=(NUMREGISTROS/MAXPAGINA)
	if lotes>(NUMREGISTROS/MAXPAGINA) then
		lotes=(NUMREGISTROS/MAXPAGINA)+1
	else
		lotes=(NUMREGISTROS/MAXPAGINA)
	end if

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

	rst.PageSize=MAXPAGINA
	rst.AbsolutePage=lote
	'-----------------------------------------'

	NavPaginas lote,lotes,campo,criterio,texto,1%>
	<hr/>
	<table width='100%' border='0' style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
		<%'Fila de encabezado'
		DrawFila color_fondo
			DrawCelda "TDBORDECELDAB7","","",0,LitCategoria
			DrawCelda "TDBORDECELDAB7","","",0,LitFamilia
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitUnidades
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCostes
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBeneficio
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBCompras
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenUnidades
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenBeneficio
			Cambio=true
		CloseFila

		fila=1
		DivisaAnt=""
		AbrAnt=""
		DecAnt=""
		Col1Anterior=""

		while not rst.EOF and fila<=MAXPAGINA
			nomCol1="categoria"
			Col1Actual=rst("categoria")
			if verCodCFS>"" and trimCodEmpresa(rst("categoria"))<>"XxYxZ" then
				datosCol1=trimCodEmpresa(rst("categoria")) & " - " & rst("nomCategoria")
			else
				datosCol1=rst("nomCategoria")
			end if
			DivisaActual=rst("divisa")

			if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
				DrawFila color_fondo
					rstAux.cursorlocation=3
					rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    Reg=rstAux.recordcount
                    end if
					rstAux.close
					rstAux.CursorLocation=3
					rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(Reg) ,2
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    else
					    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(Reg) ,2
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    end if
					rstAux.close
					Cambio=true
				CloseFila
				DrawFila color_blau
					DrawCeldaSpan "CELDA","","",0,"&nbsp;",9
				CloseFila
				DrawFila color_fondo
					DrawCelda "TDBORDECELDAB7","","",0,LitCategoria
					DrawCelda "TDBORDECELDAB7","","",0,LitFamilia
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitUnidades
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCostes
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitVentas
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBeneficio
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBCompras
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBVentas
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenVentas
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenUnidades
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenBeneficio
				CloseFila
			end if
			DrawFila color_blau
				rstAux.cursorlocation=3
				rstAux.open "select case when sum(unidades)=0 then 0 else (" & replace(rst("unidades"),",",".") & "/sum(unidades))*100 end as porcenUnidadesTotal,case when sum(ventas)=0 then 0 else (" & replace(rst("ventas"),",",".") & "/sum(ventas))*100 end as porcenVentasTotal,case when sum(beneficio)=0 then 0 else (" & replace(rst("beneficio"),",",".") & "/sum(beneficio))*100 end as porcenBeneficioTotal  from [" & session("usuario") & "] where divisa='" & rst("divisa") & "' and categoria='" & replace(rst("categoria"),"'","''") & "'",session("backendListados")
				DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(iif(Cambio,left(datosCol1,28),"&nbsp;"))
				if verCodCFS>"" and trimCodEmpresa(rst("familia_padre")) <> "XxYxZ" then
					DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(trimCodEmpresa(rst("familia_padre")) & " - " & left(rst("nomFamiliaPadre"),20))
				else
					DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(left(rst("nomFamiliaPadre"),20))
				end if
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("unidades"),DEC_CANT,-1,0,-1))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("costes"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("ventas"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("beneficio"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("bencompras"),2,-1,0,-1) & "%")
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("benventas"),2,-1,0,-1) & "%")
                if not rstAux.eof then
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenVentasTotal"),2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenUnidadesTotal"),2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("porcenBeneficioTotal"),2,-1,0,-1) & "%")
                else
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
                end if
				rstAux.close
			CloseFila
			DivisaAnt=rst("divisa")
			AbrAnt=rst("abreviatura")
			DecAnt=rst("ndecimales")
			Col1Anterior=rst("categoria")

			rst.movenext
			fila=fila+1
			Cambio=false
		wend
		if rst.eof then
			DrawFila color_fondo
				rstAux.cursorlocation=3
				rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Reg=rstAux.recordcount
                else
                    Reg=0
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(Reg),2
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                else
				    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(Reg),2
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                end if
				rstAux.close
			CloseFila
			DrawFila color_blau
					DrawCeldaSpan "CELDA","","",0,"&nbsp;",10
			CLoseFila
			'FILA DEL GRAN TOTAL
			DrawFila color_fondo
				rst.movefirst
				DivisaAnt=""
				Col1Anterior=""
				Reg=0
				Tunidades=0
				Tcostes=0
				Tventas=0
				Tbenef=0
				while not rst.EOF
					nomCol1="categoria"
					Col1Actual=rst("categoria")

					DivisaActual=rst("divisa")

					if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
						rstAux.cursorlocation=3
						rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                        if not rstAux.eof then
						    Reg=Reg + rstAux.recordcount
                        end if
						rstAux.close
						rstAux.CursorLocation=3
						rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                        if not rstAux.eof then
						    Tunidades=Tunidades + rstAux("TotalUnidades")
						    Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
						    Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
						    Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
                        end if
						rstAux.close
					end if
					DivisaAnt=rst("divisa")
					Col1Anterior=Col1Actual
					rst.movenext
				wend

				rstAux.cursorlocation=3
				rstAux.open "select distinct categoria from [" & session("usuario") & "]",session("backendListados")
                if not rstAux.eof then
				    Regcategoria=rstAux.recordcount
                else
                    Regcategoria=0
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Reg=Reg + rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Tunidades=Tunidades + rstAux("TotalUnidades")
				    Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
				    Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
				    Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
                end if
				rstAux.close

				if Tcostes>0 then BenCompras=(Tbenef*100)/Tcostes
				if Tventas>0 then BenVentas=(Tbenef*100)/Tventas

				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCategorias & " : " & EncodeForHtml(RegCategoria)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(Reg)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tunidades,DeMB,-1,0,-1))
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tcostes,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tventas,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tbenef,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenCompras,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenVentas,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
			CloseFila
		end if
		NavPaginas lote,lotes,campo,criterio,texto,2
		rst.close%>
	</table>
<%end sub

'*************************************************************************************************************'
sub MuestraListadoSubf()
	lote=limpiaCadena(Request.QueryString("lote"))

	if lote="" then
		lote=1
	end if
	sentido=limpiaCadena(Request.QueryString("sentido"))
	lotes=(NUMREGISTROS/MAXPAGINA)
	if lotes>(NUMREGISTROS/MAXPAGINA) then
		lotes=(NUMREGISTROS/MAXPAGINA)+1
	else
		lotes=(NUMREGISTROS/MAXPAGINA)
	end if

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

	rst.PageSize=MAXPAGINA
	rst.AbsolutePage=lote
	'-----------------------------------------'

	NavPaginas lote,lotes,campo,criterio,texto,1%>
	<hr/>
	<table width='100%' border='0' style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
		<%'Fila de encabezado'
		DrawFila color_fondo
			DrawCelda "TDBORDECELDAB7","","",0,LitCategoria
			DrawCelda "TDBORDECELDAB7","","",0,LitFamilia
			DrawCelda "TDBORDECELDAB7","","",0,LitSubfamilia
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitUnidades
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCostes
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBeneficio
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBCompras
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenUnidades
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenBeneficio
			Cambio=true
			CambioFamPadre=true
		CloseFila

		fila=1
		DivisaAnt=""
		AbrAnt=""
		DecAnt=""
		Col1Anterior=""
		datosCollAnterior=""
		AgrSecundariaAnt=""
		Tunidades=0
		Tventas=0
		Tbenef=0

		while not rst.EOF and fila<=MAXPAGINA
			AgrSecundaria=rst("familia_padre")
			nomCol1="categoria"
			Col1Actual=rst("categoria")
			if verCodCFS>"" and trimCodEmpresa(rst("categoria"))<>"XxYxZ" then
				datosCol1=trimCodEmpresa(rst("categoria")) & " - " & rst("nomcategoria")
			else
				datosCol1=rst("nomcategoria")
			end if
			DivisaActual=rst("divisa")

			if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
				DrawFila color_terra
					rstAux.cursorlocation=3
					DrawCelda "tdbordeCELDA7 bgcolor=" & color_blau,"","",0,""
					rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and familia_padre='" & replace(AgrSecundariaAnt,"'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    Reg=rstAux.recordcount
                    else
                        Reg=0
                    end if
					rstAux.close
					rstAux.CursorLocation=3
					rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and familia_padre='" & replace(AgrSecundariaAnt,"'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,""
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    else
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,""
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    end if
					rstAux.close
					CambioFamPadre=true
				CloseFila
				DrawFila color_fondo
					rstAux.cursorlocation=3
					rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    Reg=rstAux.recordcount
                    else
                        Reg=0
                    end if
					rstAux.close
					rstAux.CursorLocation=3
					rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    RegFamPadre=rstAux.recordcount
                    else
                        RegFamPadre=0
                    end if
					rstAux.close
					rstAux.CursorLocation=3
					rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    DrawCelda "TDBORDECELDAB7 align='LEFT'","","",0,""
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(RegFamPadre)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    else
					    DrawCelda "TDBORDECELDAB7 align='LEFT'","","",0,""
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(RegFamPadre)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    end if
					rstAux.close
					Cambio=true
				CloseFila
				DrawFila color_blau
					DrawCeldaSpan "CELDA","","",0,"&nbsp;",9
				CloseFila
				DrawFila color_fondo
					DrawCelda "TDBORDECELDAB7","","",0,LitCategoria
					DrawCelda "TDBORDECELDAB7","","",0,LitFamilia
					DrawCelda "TDBORDECELDAB7","","",0,LitSubfamilia
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitUnidades
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCostes
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitVentas
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBeneficio
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBCompras
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBVentas
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenVentas
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenUnidades
					DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenBeneficio
				CloseFila
			end if
			if AgrSecundaria<>AgrSecundariaAnt and AgrSecundariaAnt<>"" and cambio=false then
				DrawFila color_terra
					rstAux.cursorlocation=3
					DrawCelda "tdbordeCELDA7 bgcolor=" & color_blau,"","",0,""
					rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and familia_padre='" & replace(AgrSecundariaAnt,"'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    Reg=rstAux.recordcount
                    else
                        Reg=0
                    end if
					rstAux.close
					rstAux.CursorLocation=3
					rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and familia_padre='" & replace(AgrSecundariaAnt,"'","''") & "'",session("backendListados")
                    if not rstAux.eof then
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,""
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    else
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,""
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                    end if
					rstAux.close
					CambioFamPadre=true
				CloseFila
			end if
			DrawFila color_blau
				DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(iif(Cambio,left(datosCol1,23),"&nbsp;"))
				if verCodCFS>"" and trimCodEmpresa(rst("familia_padre")) <> "XxYxZ" then
					DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(iif(CambioFamPadre,trimCodEmpresa(rst("familia_padre")) & " - " & left(rst("nomFamiliaPadre"),15),"&nbsp;"))
				else
					DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(iif(CambioFamPadre,left(rst("nomFamiliaPadre"),15),"&nbsp;"))
				end if
				if verCodCFS>"" and trimCodEmpresa(rst("familia")) <> "XxYxZ" then
					DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(trimCodEmpresa(rst("familia")) & " - " & left(rst("nomFamilia"),15))
				else
					DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(left(rst("nomFamilia"),15))
				end if
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("unidades"),DEC_CANT,-1,0,-1))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("costes"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("ventas"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("beneficio"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("bencompras"),2,-1,0,-1) & "%")
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("benventas"),2,-1,0,-1) & "%")
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("porcenVentasTotal"),2,-1,0,-1) & "%")
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("porcenUnidadesTotal"),2,-1,0,-1) & "%")
				DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("porcenBeneficioTotal"),2,-1,0,-1) & "%")
			CloseFila

			AgrSecundariaAnt=rst("familia_padre")
			DivisaAnt=rst("divisa")
			AbrAnt=rst("abreviatura")
			DecAnt=rst("ndecimales")
			Col1Anterior=rst("categoria")
			datosCollAnterior=rst("nomcategoria")

			rst.movenext
			fila=fila+1
			Cambio=false
			CambioFamPadre=false
		wend
		if rst.eof then
			DrawFila color_terra
				rstAux.cursorlocation=3
				DrawCelda "tdbordeCELDA7 bgcolor=" & color_blau,"","",0,""
				rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and familia_padre='" & replace(AgrSecundariaAnt,"'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Reg=rstAux.recordcount
                else
                    Reg=0
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and familia_padre='" & replace(AgrSecundariaAnt,"'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,""
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                else
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,""
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubfamilias & " : " & EncodeForHtml(Reg)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                end if
				rstAux.close
			CloseFila
			DrawFila color_fondo
				rstAux.cursorlocation=3
				rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Reg=rstAux.recordcount
                else
                    Reg=0
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    RegFamPadre=rstAux.recordcount
                else
                    RegFamPadre=0
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    DrawCelda "TDBORDECELDAB7 align='LEFT'","","",0,""
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,Litfamilias & " : " & EncodeForHtml(RegFamPadre)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubFamilias & " : " & EncodeForHtml(Reg)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                else
                    DrawCelda "TDBORDECELDAB7 align='LEFT'","","",0,""
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,Litfamilias & " : " & EncodeForHtml(RegFamPadre)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubFamilias & " : " & EncodeForHtml(Reg)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DEC_CANT,-1,0,-1))
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,DecAnt,-1,0,-1) & AbrAnt)
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(0,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
                end if
				rstAux.close
			CloseFila
			DrawFila color_blau
					DrawCeldaSpan "CELDA","","",0,"&nbsp;",10
			CLoseFila
			'FILA DEL GRAN TOTAL
			DrawFila color_fondo
				rst.movefirst
				DivisaAnt=""
				Col1Anterior=""
				Reg=0
				RegFamPadre=0
				Tunidades=0
				Tcostes=0
				Tventas=0
				Tbenef=0
				while not rst.EOF
					nomCol1="categoria"
					Col1Actual=rst("categoria")

					DivisaActual=rst("divisa")

					if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
						rstAux.cursorlocation=3
						rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                        if not rstAux.eof then
						    Reg=Reg + rstAux.recordcount
                        end if
						rstAux.close
						rstAux.CursorLocation=3
						rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                        if not rstAux.eof then
						    RegFamPadre=RegFamPadre + rstAux.recordcount
                        end if
						rstAux.close
						rstAux.CursorLocation=3
						rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                        if not rstAux.eof then
						    Tunidades=Tunidades + rstAux("TotalUnidades")
						    Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
						    Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
						    Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
                        end if
						rstAux.close
					end if
					DivisaAnt=rst("divisa")
					Col1Anterior=Col1Actual
					rst.movenext
				wend

				rstAux.cursorlocation=3
				rstAux.open "select distinct familia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Reg=Reg + rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select distinct familia_padre from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    RegFamPadre=RegFamPadre + rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select distinct categoria from [" & session("usuario") & "]",session("backendListados")
                if not rstAux.eof then
				    RegCategoria=rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
				rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                if not rstAux.eof then
				    Tunidades=Tunidades + rstAux("TotalUnidades")
				    Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
				    Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
				    Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
                end if
				rstAux.close

				if Tcostes>0 then BenCompras=(Tbenef*100)/Tcostes
				if Tventas>0 then BenVentas=(Tbenef*100)/Tventas

				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCategorias & " : " & EncodeForHtml(RegCategoria)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitFamilias & " : " & EncodeForHtml(RegFamPadre)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitSubFamilias & " : " & EncodeForHtml(Reg)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tunidades,DeMB,-1,0,-1))
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tcostes,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tventas,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tbenef,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenCompras,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenVentas,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
			CloseFila
		end if
		NavPaginas lote,lotes,campo,criterio,texto,2
		rst.close%>
	</table>
<%end sub

'***************************************************************'

sub MuestraListado()
	lote=limpiaCadena(Request.QueryString("lote"))

	if lote="" then
		lote=1
	end if
	sentido=limpiaCadena(Request.QueryString("sentido"))

	lotes=(NUMREGISTROS/MAXPAGINA)
	if lotes>(NUMREGISTROS/MAXPAGINA) then
		lotes=(NUMREGISTROS/MAXPAGINA)+1
	else
		lotes=(NUMREGISTROS/MAXPAGINA)
	end if

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

	rst.PageSize=MAXPAGINA
	rst.AbsolutePage=lote
	'-----------------------------------------'

	NavPaginas lote,lotes,campo,criterio,texto,1%>
	<hr/>
	<table width='100%' border='0' style="border-collapse: collapse;" cellspacing="1" cellpadding="1">
		<%'Fila de encabezado'
		DrawFila color_fondo
			if ucase(agruparart)=ucase("CLIENTE") then
				LitCol1=LitCliente
			elseif ucase(agruparart)=ucase("COMERCIAL") then
				if si_tiene_modulo_comercial<>0 then
					LitCol1=LitComercialModCom
				else
					LitCol1=LitComercial
				end if
			elseif ucase(agruparart)="SUBFAMILIA" or ucase(agruparart)="ARTICULO" then
				LitCol1=LitSubFamilia
			elseif ucase(agruparart)=ucase("PROYECTO") then
				LitCol1=LitProyecto
			end if
            if ucase(agruparart)<>"ARTICULO" then
			    DrawCelda "TDBORDECELDAB7","","",0,LitCol1
            end if
            if ucase(agruparart)=ucase("CLIENTE") then
                if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                    DrawCelda "TDBORDECELDAB7","","",0,LitCentro
                end if
            end if
			if ucase(nodetallado)<>"ON" then
			    DrawCelda "TDBORDECELDAB7","","",0,LitReferencia
			    DrawCelda "TDBORDECELDAB7","","",0,LitDescripcion
			    if ucase(agruparart)<>"SUBFAMILIA" and ucase(agruparart)<>"ARTICULO" then
				    DrawCelda "TDBORDECELDAB7","","",0,LitSubFamilia
			    end if
			end if
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitUnidades
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCostes
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitVentas
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBeneficio
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBCompras
			DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBVentas
			if ucase(nodetallado)<>"ON" then
			    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenVentas
			    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenUnidades
			    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenBeneficio
			end if
			Cambio=true
		CloseFila

		fila=1
		DivisaAnt=""
		AbrAnt=""
		DecAnt=""
		Col1Anterior=""

		while not rst.EOF and fila<=MAXPAGINA
			if ucase(agruparart)=ucase("CLIENTE") then
				nomCol1="ncliente"
				Col1Actual=rst("ncliente")
				datosCol1=Hiperv(OBJClientes,rst("ncliente"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("nomCliente"),LitVerCliente)
                dataCenter=Hiperv(OBJCentros,rst("ncentro"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),rst("nomCentro"),LitVerCentro)
			elseif ucase(agruparart)=ucase("COMERCIAL") or ucase(agruparart)=ucase("RESPONSABLE") then
				nomCol1="comercial"
				Col1Actual=rst("comercial")
				datosCol1=rst("nomComercial")
			elseif ucase(agruparart)="SUBFAMILIA" then ''or ucase(agruparart)="ARTICULO" then
				nomCol1="familia"
				Col1Actual=rst("familia")
				if verCodCFS>"" and trimCodEmpresa(rst("familia"))<>"XxYxZ" then
					datosCol1=trimCodEmpresa(rst("familia"))& " - "& rst("nomFamilia")
				else
					datosCol1= rst("nomFamilia")
				end if
			elseif ucase(agruparart)=ucase("FAMILIA") then
				nomCol1="familia_padre"
				Col1Actual=rst("familia_padre")
				datosCol1=rst("nomFamiliaPadre")
			elseif ucase(agruparart)=ucase("PROYECTO") then
				nomCol1="isnull(cod_proyecto,'')"
				Col1Actual=rst("cod_proyecto")
				datosCol1=rst("nomProyecto")
            elseif ucase(agruparart)="ARTICULO" then
                Col1Actual =session("ncliente")&"XxYxZ"
			end if
			DivisaActual=rst("divisa")
			if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
			    if ucase(nodetallado)="ON" then
		            DrawFila color_blau
		        else
			        DrawFila color_fondo
			    end if
					rstAux.cursorlocation=3
                    if ucase(agruparart)<>"ARTICULO" then
                        strSel="select distinct referencia,ncliente,nomCliente,comercial,nomComercial,cod_proyecto,nomProyecto,ncentro,nomCentro from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'"
                    else
                        strSel="select distinct referencia,ncliente,nomCliente,comercial,nomComercial,cod_proyecto,nomProyecto,ncentro,nomCentro from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'"
                    end if
					rstAux.open strSel,session("backendListados")
					if not rstAux.eof then
					    Reg=rstAux.recordcount
					    valor_ncliente=rstAux("ncliente")
					    valor_nomCliente=rstAux("nomCliente")
					    valor_comercial=rstAux("comercial")
					    valor_nomComercial=rstAux("nomComercial")
					    valor_cod_proyecto=rstAux("cod_proyecto")
					    valor_nomProyecto=rstAux("nomProyecto")
                        valor_ncentro=rstAux("ncentro")
					    valor_nomCentro=rstAux("nomCentro")
					else
					    Reg=0
					    valor_ncliente=""
					    valor_nomCliente=""
					    valor_comercial=""
					    valor_nomComercial=""
					    valor_cod_proyecto=""
					    valor_nomProyecto=""
                        valor_ncentro=""
					    valor_nomCentro=""
					end if
					rstAux.close
					rstAux.CursorLocation=3
                    if ucase(agruparart)<>"ARTICULO" then
					    rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                    else
                        rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'",session("backendListados")
                    end if
				    if ucase(nodetallado)="ON" then
				        ''DrawCelda "TDBORDECELDAB7 align='LEFT'","","",0,valor_NomCliente
				        if ucase(agruparart)=ucase("CLIENTE") then
				            valor_datosCol1=Hiperv(OBJClientes,valor_ncliente,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),valor_NomCliente,LitVerCliente)
                            valor_dataCenter=Hiperv(OBJCentros,valor_ncentro,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),valor_nomCentro,LitVerCentro)
				        elseif ucase(agruparart)=ucase("COMERCIAL") or ucase(agruparart)=ucase("RESPONSABLE") then
				            valor_datosCol1=valor_nomComercial
				        elseif ucase(agruparart)=ucase("PROYECTO") then
				            valor_datosCol1=valor_nomProyecto
				        end if
				        DrawCelda "tdbordeCELDA7","","",0,valor_datosCol1
                        if ucase(agruparart)=ucase("CLIENTE") then
                            if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                                DrawCelda "tdbordeCELDA7","","",0,dataCenter
                            end if
                        end if
				        ''DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitRegistros & " : " & Reg & ".&nbsp;&nbsp;&nbsp;" & LitTotales & " : ",2
				        DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
				        claseTD="tdbordeCELDA7"
				    else
                        if ucase(agruparart)<>"ARTICULO" then
                            colspan = 3
                        else
                            colspan = 2
                        end if
                        if ucase(agruparart)=ucase("CLIENTE") then
                            if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                                colspan = 4
                            end if
                        end if
				        claseTD="TDBORDECELDAB7"
				        DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitRegistros & " : " & EncodeForHtml(Reg) & ".&nbsp;&nbsp;&nbsp;" & LitTotales & " : ",colspan
				        DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1)),iif(agruparart<>"SUBFAMILIA" and agruparart<>"ARTICULO",2,1)
                    end if
					DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
					DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
					DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
					DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
					DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
					if ucase(nodetallado)<>"ON" then
					    DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					    DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
					end if
					rstAux.close
					Cambio=true
				CloseFila
				if ucase(nodetallado)<>"ON" then
				    DrawFila color_blau
					    DrawCeldaSpan "CELDA","","",0,"&nbsp;",9
				    CloseFila
				    DrawFila color_fondo
                        if ucase(agruparart)<>"ARTICULO" then
					        DrawCelda "TDBORDECELDAB7","","",0,LitCol1
                        end if
                        if ucase(agruparart)=ucase("CLIENTE") then
                            if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                                DrawCelda "TDBORDECELDAB7","","",0,LitCentro
                            end if
                        end if
				        DrawCelda "TDBORDECELDAB7","","",0,LitReferencia
				        DrawCelda "TDBORDECELDAB7","","",0,LitDescripcion
				        if ucase(agruparart)<>ucase("subfamilia") and ucase(agruparart)<>ucase("articulo") then
					        DrawCelda "TDBORDECELDAB7","","",0,LitSubFamilia
				        end if
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitUnidades
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitCostes
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitVentas
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBeneficio
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBCompras
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitBVentas
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenVentas
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenUnidades
					    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,LitPorcenBeneficio
				    CloseFila
                end if
			end if

			if ucase(nodetallado)<>"ON" then
			    DrawFila color_blau
                    if ucase(agruparart)<>"ARTICULO" then
				        DrawCelda "tdbordeCELDA7","","",0,iif(Cambio,datosCol1,"&nbsp;")
                    end if
                    if ucase(agruparart)=ucase("CLIENTE") then
                        if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                            DrawCelda "tdbordeCELDA7","","",0,dataCenter
                        end if
                    end if
				    DrawCelda "tdbordeCELDA7","","",0,Hiperv(OBJArticulos,rst("referencia"),"","",Permisos,Enlaces,session("usuario"),session("ncliente"),trimCodEmpresa(rst("referencia")),LitVerArticulo)
				    DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(rst("descripcion"))
				    if ucase(agruparart)<>ucase("subfamilia") and ucase(agruparart)<>ucase("articulo") then
					    if verCodCFS>"" and trimCodEmpresa(rst("familia"))<>"XxYxZ" then
						    DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(trimCodEmpresa(rst("familia"))&" - "&rst("nomfamilia"))
					    else
						    DrawCelda "tdbordeCELDA7","","",0,EncodeForHtml(rst("nomfamilia"))
					    end if
				    end if
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("unidades"),DEC_CANT,-1,0,-1))
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("costes"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("ventas"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("beneficio"),rst("ndecimales"),-1,0,-1) & " " & rst("abreviatura"))
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("bencompras"),2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("benventas"),2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("porcenVentasTotal"),2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("porcenUnidadesTotal"),2,-1,0,-1) & "%")
				    DrawCelda "tdbordeCELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rst("porcenBeneficioTotal"),2,-1,0,-1) & "%")
			    CloseFila
            end if
			DivisaAnt=rst("divisa")
			AbrAnt=rst("abreviatura")
			DecAnt=rst("ndecimales")
			if ucase(agruparart)=ucase("CLIENTE") then
				Col1Anterior=rst("ncliente")
			elseif ucase(agruparart)=ucase("COMERCIAL") or ucase(agruparart)=ucase("RESPONSABLE") then
				Col1Anterior=rst("comercial")
			elseif ucase(agruparart)=ucase("SUBFAMILIA") or ucase(agruparart)=ucase("ARTICULO") then
				Col1Anterior=rst("familia")
			elseif ucase(agruparart)=ucase("PROYECTO") then
				Col1Anterior=rst("cod_proyecto")
			end if

			rst.movenext
			fila=fila+1
			Cambio=false
		wend
		if rst.eof then
		    if ucase(nodetallado)="ON" then
		        DrawFila color_blau
		    else
			    DrawFila color_fondo
			end if
				rstAux.cursorlocation=3
                if ucase(agruparart)<>"ARTICULO" then
                    strSel="select distinct referencia,ncliente,nomCliente,comercial,nomComercial,cod_proyecto,nomProyecto,ncentro,nomCentro from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'"
                else
                    strSel="select distinct referencia,ncliente,nomCliente,comercial,nomComercial,cod_proyecto,nomProyecto,ncentro,nomCentro from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'"
                end if
				rstAux.open strSel,session("backendListados")
				if not rstAux.eof then
				    Reg=rstAux.recordcount
				    valor_ncliente=rstAux("ncliente")
				    valor_nomCliente=rstAux("nomCliente")
				    valor_comercial=rstAux("comercial")
				    valor_nomComercial=rstAux("nomComercial")
				    valor_cod_proyecto=rstAux("cod_proyecto")
				    valor_nomProyecto=rstAux("nomProyecto")
                    valor_ncentro=rstAux("ncentro")
					valor_nomCentro=rstAux("nomCentro")
				else
				    Reg=0
				    valor_ncliente=""
				    valor_nomCliente=""
				    valor_comercial=""
				    valor_nomComercial=""
				    valor_cod_proyecto=""
				    valor_nomProyecto=""
                    valor_ncentro=""
					valor_nomCentro=""
				end if
				rstAux.close
				rstAux.CursorLocation=3
                if ucase(agruparart)<>"ARTICULO" then
				    rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                else
                    rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio,case when sum(costes)>0 then ((sum(beneficio)*100)/sum(costes)) else 0 end as BenCompras,case when sum(ventas)>0 then ((sum(beneficio)*100)/sum(ventas)) else 0 end as BenVentas from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'",session("backendListados")
                end if
				if ucase(nodetallado)="ON" then
			        if ucase(agruparart)=ucase("CLIENTE") then
			            valor_datosCol1=Hiperv(OBJClientes,valor_ncliente,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),valor_NomCliente,LitVerCliente)
                        valor_dataCenter=Hiperv(OBJCentros,valor_ncentro,"","",Permisos,Enlaces,session("usuario"),session("ncliente"),valor_nomCentro,LitVerCentro)
			        elseif ucase(agruparart)=ucase("COMERCIAL") or ucase(agruparart)=ucase("RESPONSABLE") then
			            valor_datosCol1=valor_nomComercial
			        elseif ucase(agruparart)=ucase("PROYECTO") then
			            valor_datosCol1=valor_nomProyecto
			        end if
			        DrawCelda "tdbordeCELDA7","","",0,valor_datosCol1
                    if ucase(agruparart)=ucase("CLIENTE") then
                        if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                            DrawCelda "tdbordeCELDA7","","",0,valor_dataCenter
                        end if
                    end if
				    ''DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitRegistros & " : " & Reg & ".&nbsp;&nbsp;&nbsp;" & LitTotales & " : ",2
				    DrawCelda "TDBORDECELDA7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1))
				    claseTD="TDBORDECELDA7"
				else
                    if ucase(agruparart)<>"ARTICULO" then
                        colspan = 3
                    else
                        colspan = 2
                    end if
                    if ucase(agruparart)=ucase("CLIENTE") then
                        if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                            colspan = 4
                        end if
                    end if
				    claseTD="TDBORDECELDAB7"
				    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitRegistros & " : " & EncodeForHtml(Reg) & ".&nbsp;&nbsp;&nbsp;" & LitTotales & " : ",colspan
				    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalUnidades"),DEC_CANT,-1,0,-1)),iif(agruparart<>"SUBFAMILIA" and agruparart<>"ARTICULO",2,1)
                end if
				
				DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("TotalCostes"),DecAnt,-1,0,-1) & AbrAnt)
				DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalventas"),DecAnt,-1,0,-1) & AbrAnt)
				DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("Totalbeneficio"),DecAnt,-1,0,-1) & AbrAnt)
				DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenCompras"),2,-1,0,-1) & "%")
				DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(rstAux("BenVentas"),2,-1,0,-1) & "%")
				if ucase(nodetallado)<>"ON" then
				    DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda claseTD & " align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				end if
				rstAux.close
			CloseFila
			DrawFila color_blau
				DrawCeldaSpan "CELDA","","",0,"&nbsp;",10
			CLoseFila
			'FILA DEL GRAN TOTAL
			DrawFila color_fondo
				rst.movefirst
				DivisaAnt=""
				Col1Anterior=""
				Reg=0
				Tunidades=0
				Tcostes=0
				Tventas=0
				Tbenef=0
				while not rst.EOF
					if ucase(agruparart)=ucase("CLIENTE") then
						nomCol1="ncliente"
						Col1Actual=rst("ncliente")
					elseif ucase(agruparart)=ucase("COMERCIAL") or ucase(agruparart)=ucase("RESPONSABLE") then
						nomCol1="comercial"
						Col1Actual=rst("comercial")
					elseif ucase(agruparart)=ucase("SUBFAMILIA") then ''or ucase(agruparart)=ucase("ARTICULO") then
						nomCol1="familia"
						Col1Actual=rst("familia")
					elseif ucase(agruparart)=ucase("PROYECTO") then
						nomCol1="isnull(cod_proyecto,'')"
						Col1Actual=rst("cod_proyecto")
					end if
					DivisaActual=rst("divisa")

					if (DivisaActual<>DivisaAnt and DivisaAnt<>"") or (ucase(Col1Actual)<>ucase(Col1Anterior) and Col1Anterior<>"") then
						rstAux.cursorlocation=3
                        if ucase(agruparart)<>"ARTICULO" then
                            strSel="select distinct referencia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'"
                        else
                            strSel="select distinct referencia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'"
                        end if

						rstAux.open strSel,session("backendListados")
						Reg=Reg + rstAux.recordcount
						rstAux.close
						rstAux.CursorLocation=3
                        if ucase(agruparart)<>"ARTICULO" then
						    rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                        else
                            rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'",session("backendListados")
                        end if
                        if not rstAux.eof then
						    Tunidades=Tunidades + rstAux("TotalUnidades")
						    Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
						    Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
						    Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
                        end if
						rstAux.close
					end if
					DivisaAnt=rst("divisa")
					Col1Anterior=Col1Actual
					rst.movenext
				wend

				rstAux.cursorlocation=3
                if ucase(agruparart)<>"ARTICULO" then
                    strSel="select distinct referencia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'"
                else
                    strSel="select distinct referencia from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'"
                end if
				rstAux.open strSel,session("backendListados")
                if not rstAux.eof then
				    Reg=Reg + rstAux.recordcount
                end if
				rstAux.close
				rstAux.CursorLocation=3
                if ucase(agruparart)<>"ARTICULO" then
				    rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "' and " & nomCol1 & "='" & replace(Col1Anterior&"","'","''") & "'",session("backendListados")
                else
                    rstAux.open "select count(*) as regs,sum(unidades) as TotalUnidades,sum(costes) as totalcostes, sum(ventas) as totalventas, sum(beneficio) as totalbeneficio from [" & session("usuario") & "] where divisa='" & DivisaAnt & "'",session("backendListados")
                end if
                if not rstAux.eof then
				    Tunidades=Tunidades + rstAux("TotalUnidades")
				    Tcostes=Tcostes + CambioDivisa(rstAux("totalcostes"),DivisaAnt,MB)
				    Tventas=Tventas + CambioDivisa(rstAux("totalventas"),DivisaAnt,MB)
				    Tbenef=Tbenef + CambioDivisa(rstAux("totalbeneficio"),DivisaAnt,MB)
                end if
				rstAux.close

				if Tcostes>0 then BenCompras=(Tbenef*100)/Tcostes
				if Tventas>0 then BenVentas=(Tbenef*100)/Tventas
                if ucase(nodetallado)="ON" then
                    if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                        if ucase(agruparart)=ucase("CLIENTE") then
                            colspan =2
                        else
                            colspan =1
                        end if
                    else
                        colspan =1
                    end if
                    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitTotales & " : ", colspan
                    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tunidades,DeMB,-1,0,-1)),1
                else
                    if ucase(agruparart)<>"ARTICULO" then
                        colspan = 3
                    else
                        colspan = 2
                    end if
                    if ucase(agruparart)=ucase("CLIENTE") then
                        if has_module_advanced_management <> 0 or has_module_postventa <> 0 then
                            colspan = 4
                        end if
                    end if
				    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,LitRegistros & " : " & EncodeForHtml(Reg) & ".&nbsp;&nbsp;&nbsp;" & LitTotales & " : ",colspan
				    DrawCeldaSpan "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tunidades,DeMB,-1,0,-1)),iif(agruparart<>"SUBFAMILIA" and agruparart<>"ARTICULO",2,1)
                end if
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tcostes,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tventas,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(Tbenef,DeMB,-1,0,-1) & AbMB)
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenCompras,2,-1,0,-1) & "%")
				DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(BenVentas,2,-1,0,-1) & "%")
				if ucase(nodetallado)<>"ON" then
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				    DrawCelda "TDBORDECELDAB7 align='RIGHT'","","",0,EncodeForHtml(formatnumber(100,2,-1,0,-1) & "%")
				end if
			CloseFila
		end if
		NavPaginas lote,lotes,campo,criterio,texto,2
		rst.close%>
	</table>
<%end sub

'*********************************************************************************************************
sub MuestraParamSelec()
	%><font class='CELDA7'><b><%=LitDesdeFecha%> :&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(fdesde)%></font><br/>
	<font class='CELDA7'><b><%=LitHastaFecha%> :&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(fhasta)%></font><br/><%
	if nserie & "" > "" then
		%><font class='CELDA7'><b><%=LitSerie%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(trimCodEmpresa(nserie))%>&nbsp;<%=EncodeForHtml(d_lookup("nombre","series","nserie='" & nserie & "'",session("backendListados")))%></font><br/><%
	end if

	if verCodCFS & "" > "" then
		%><font class='CELDA7'><b><%=LitVerCods_CFS2%>&nbsp;</b></font><br/><%
	end if

	if UCase(ordenar)="CODIGOCFS" then
		%><font class='CELDA7'><b><%=LitOrdenar%>:&nbsp;</b></font><font class='CELDA7'><%=LitCodCFS%></font><br/><%
	end if

	if ncliente & "" > "" then
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open session("backendListados")
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="select rsocial from clientes with(nolock) where ncliente=?"
        command.CommandType = adCmdText 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@ncliente", adVarChar, adParamInput, 10, ncliente)

        set rstCustomer = command.execute

        if not rstCustomer.eof then
            nomcliente = rstCustomer("rsocial")
        end if

        rstCustomer.close
        conn.close
        set rstCustomer = nothing
        set command = nothing
        set conn = nothing
		%><font class='CELDA7'><b><%=LitCliente%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(trimCodEmpresa(ncliente))%>&nbsp;<%=EncodeForHtml(nomcliente)%></font><br/><%
	end if
    if ncenter & "" > "" then
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open session("backendListados")
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="select rsocial from centros with(nolock) where ncentro=?"
        command.CommandType = adCmdText 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@ncentro", adVarChar, adParamInput, 10, ncenter)

        set rstCenter = command.execute

        if not rstCenter.eof then
            nameCenter = rstCenter("rsocial")
        end if

        rstCenter.close
        conn.close
        set rstCenter = nothing
        set command = nothing
        set conn = nothing
		%><font class='CELDA7'><b><%=LitCentro%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(trimCodEmpresa(ncenter))%>&nbsp;<%=EncodeForHtml(nameCenter)%></font><br/><%
	end if
	if tipoCliente & "" > "" then
		%><font class='CELDA7'><b><%=LitTipoCliente%>:&nbsp;</b></font><font class='CELDA7'>&nbsp;<%=EncodeForHtml(d_lookup("descripcion","tipos_entidades","codigo='" & tipoCliente & "'",session("backendListados")))%></font><br/><%
	end if
	if cod_proyecto & "" > "" then
		%><font class='CELDA7'><b><%=LitProyecto%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(d_lookup("nombre","proyectos","codigo='" & cod_proyecto & "'",session("backendListados")))%></font><br/><%
	end if

	if familia<>"" then
		desc_familia=NombresEntidades(familia,"familias","codigo","nombre",session("backendListados"))
		%><font class='CELDA7'><b><%=LitSubFamilia%> :&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(desc_familia)%></font><br/><%
	elseif familia_padre<>"" then
		desc_familia_padre=NombresEntidades(familia_padre,"familias_padre","codigo","nombre",session("backendListados"))
		%><font class='CELDA7'><b><%=LitFamilia%> :&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(desc_familia_padre)%></font><br/><%
	elseif categoria<>"" then
		desc_categoria=NombresEntidades(categoria,"categorias","codigo","nombre",session("backendListados"))
		%><font class='CELDA7'><b><%=LitCategoria%> :&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(desc_categoria)%></font><br/><%
	end if

	if referencia & "" > "" then
		%><font class='CELDA7'><b><%=LitConRef%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(referencia)%></font><br/><%
	end if
	if nombreart & "" > "" then
		%><font class='CELDA7'><b><%=LitConNombre%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(nombreart)%></font><br/><%
	end if
	if nproveedor & "" > "" then
        set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open session("backendListados")
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="select razon_social from proveedores with(nolock) where nproveedor=?"
        command.CommandType = adCmdText 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@nproveedor", adVarChar, adParamInput, 10, nproveedor)

        set rstSupplier = command.execute

        if not rstSupplier.eof then
            nomproveedor = rstSupplier("razon_social")
        end if

        rstSupplier.close
        conn.close
        set rstSupplier = nothing
        set command = nothing
        set conn = nothing
		%><font class='CELDA7'><b><%=LitProveedor%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(trimCodEmpresa(nproveedor))%>&nbsp;<%=EncodeForHtml(nomproveedor)%></font><br/><%
	end if
	if comercial & "" > "" then
		%><font class='CELDA7'><b><%
		if si_tiene_modulo_comercial<>0 then
			response.write(LitComercialModCom)
		else
			response.write(LitComercial)
		end if
		%>:&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(d_lookup("nombre","personal","dni='" & comercial & "'",session("backendListados")))%></font><br/><%
	end if

	''ricardo 7-11-2007 se añade el tipo de articulo para poder filtrar
	if tipoarticulo & "">"" then
        desc_tipoarticulo=NombresEntidades(tipoarticulo,"tipos_entidades","codigo","descripcion",session("backendListados"))
		%><font class='CELDA7'><b><%=LitTipArtRendArt%> :&nbsp;</b></font><font class='CELDA7'><%=EncodeForHtml(desc_tipoarticulo)%></font><br/><%
	end if
	
	if actividad & "" > "" then
		%><font class='CELDA7'><b><%=LITTIPACTIV%>:&nbsp;</b></font><font class='CELDA7'>&nbsp;<%=EncodeForHtml(d_lookup("descripcion","tipo_actividad","codigo='" & actividad & "'",session("backendListados")))%></font><br/><%
	end if
	if nodetallado & "" > "" then
	    ''if cstr(nodetallado)="1" then
	    if ucase(nodetallado)="ON" then
	        %><font class='CELDA7'><b><%=LITLISTDETALLADO%></b></font><br/><%
	    end if
		
	end if
end sub

'*****************************************************************************'
'********************** CODIGO PRINCIPAL DE LA PÁGINA ************************'
'*****************************************************************************'

	const borde=0

	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion%>

<form name="rendimiento_articulos" method="post">
    <%PintarCabecera "rendimiento_articulos.asp"
	WaitBoxOculto LitEsperePorFavor%>
    <input type="hidden" name="tiene_t" value="<%=tiene_modulo_tiendas%>"/><%
	'Leer parámetros de la página'
	mode		= enc.EncodeForJavascript(Request.QueryString("mode"))
	if mode="browse" then mode="imp"
	ncliente	= limpiaCadena(Request.QueryString("ncliente"))
	if ncliente ="" then
		ncliente	= limpiaCadena(Request.form("ncliente"))
	end if
	if ncliente > "" then
		ncliente = completar(trimCodEmpresa(ncliente),5,"0")
	end if

    ncenter	= limpiaCadena(Request.QueryString("ncenter"))
	if ncenter ="" then
		ncenter	= limpiaCadena(Request.form("ncenter"))
	end if
	if ncenter > "" then
		ncenter = completar(trimCodEmpresa(ncenter),5,"0")
	end if
	actividad	= limpiaCadena(Request.QueryString("actividad"))
	if actividad ="" then
		actividad	= limpiaCadena(Request.form("actividad"))
	end if
	checkCadena actividad
	
	nodetallado	= limpiaCadena(Request.QueryString("nodetallado"))
	if nodetallado ="" then
		nodetallado	= limpiaCadena(Request.form("nodetallado"))
	end if

	nproveedor	= limpiaCadena(Request.QueryString("nproveedor"))
	if nproveedor="" then
		nproveedor	= limpiaCadena(Request.form("nproveedor"))
	end if
	if nproveedor> "" then
		nproveedor= completar(trimCodEmpresa(nproveedor),5,"0")
	end if

	fdesde = limpiaCadena(Request.QueryString("fdesde"))
	if fdesde ="" then
		fdesde = limpiaCadena(Request.form("fdesde"))
	end if

	fhasta = limpiaCadena(Request.QueryString("fhasta"))
	if fhasta ="" then
		fhasta = limpiaCadena(Request.form("fhasta"))
	end if

	nserie=limpiaCadena(request.querystring("nserie"))
	if nserie ="" then
		nserie = limpiaCadena(Request.form("nserie"))
	end if
	CheckCadena nserie

	tipoCliente=limpiaCadena(request.querystring("tipoCliente"))
	if tipoCliente ="" then
		tipoCliente = limpiaCadena(Request.form("tipoCliente"))
	end if
	CheckCadena tipoCliente

	agrcliente=limpiaCadena(request.querystring("agrcliente"))
	if agrcliente ="" then
		agrcliente = limpiaCadena(Request.form("agrcliente"))
	end if

	agrproveedor=limpiaCadena(request.querystring("agrproveedor"))
	if agrproveedor="" then
		agrproveedor= limpiaCadena(Request.form("agrproveedor"))
	end if

	cod_proyecto	= limpiaCadena(Request.QueryString("cod_proyecto"))
	if cod_proyecto="" then
		cod_proyecto = limpiaCadena(Request.form("cod_proyecto"))
	end if
	CheckCadena cod_proyecto

	agrproyecto=limpiaCadena(request.querystring("agrproyecto"))
	if agrproyecto ="" then
		agrproyecto = limpiaCadena(Request.form("agrproyecto"))
	end if

	familia=limpiaCadena(request.querystring("familia"))
	if familia ="" then
		familia = limpiaCadena(Request.form("familia"))
	end if
	CheckCadena familia

	familia_padre=limpiaCadena(request.querystring("familia_padre"))
	if familia_padre ="" then
		familia_padre = limpiaCadena(Request.form("familia_padre"))
	end if

	categoria=limpiaCadena(request.querystring("categoria"))
	if categoria ="" then
		categoria = limpiaCadena(Request.form("categoria"))
	end if

	agrfamilia=limpiaCadena(request.querystring("agrfamilia"))
	if agrfamilia ="" then
		agrfamilia = limpiaCadena(Request.form("agrfamilia"))
	end if

	comercial=limpiaCadena(request.querystring("comercial"))
	if comercial="" then
		comercial= limpiaCadena(Request.form("comercial"))
	end if
	CheckCadena comercial

	agrcomercial=limpiaCadena(request.querystring("agrcomercial"))
	if agrcomercial="" then
		agrcomercial= limpiaCadena(Request.form("agrcomercial"))
	end if

	referencia=limpiaCadena(request.querystring("referencia"))
	if referencia ="" then
		referencia = limpiaCadena(Request.form("referencia"))
	end if

	nombreart=limpiaCadena(request.querystring("nombreart"))
	if nombreart ="" then
		nombreart = limpiaCadena(Request.form("nombreart"))
	end if

	coste=limpiaCadena(request.querystring("coste"))
	if coste ="" then
		coste = limpiaCadena(Request.form("coste"))
	end if

	ordenar=limpiaCadena(request.querystring("ordenar"))
	if ordenar ="" then
		ordenar = limpiaCadena(Request.form("ordenar"))
	end if

	ordenarDoc=limpiaCadena(request.querystring("ordenarDoc"))
	if ordenarDoc ="" then
		ordenarDoc = limpiaCadena(Request.form("ordenarDoc"))
	end if

	if request.QueryString("verCodCFS")>"" then
	verCodCFS=limpiaCadena(request.QueryString("verCodCFS"))
	end if
	if verCodCFS="" then
		if request.form("verCodCFS")>"" then
			verCodCFS = limpiaCadena(request.form("verCodCFS"))
		end if
	end if

	if request.QueryString("artbaja")>"" then
		artbaja=limpiaCadena(request.QueryString("artbaja"))
	end if
	if artbaja="" then
		if request.form("artbaja")>"" then
			artbaja = limpiaCadena(request.form("artbaja"))
		end if
	end if

	if request.QueryString("pedsinf")>"" then
		pedsinf=limpiaCadena(request.QueryString("pedsinf"))
	end if
	if pedsinf="" then
		if request.form("pedsinf")>"" then
			pedsinf = limpiaCadena(request.form("pedsinf"))
		end if
	end if

	if request.QueryString("ordsinf")>"" then
		ordsinf=limpiaCadena(request.QueryString("ordsinf"))
	end if
	if ordsinf="" then
		if request.form("ordsinf")>"" then
			ordsinf = limpiaCadena(request.form("ordsinf"))
		end if
	end if

	if request.querystring("calcost")>"" then
		calcost=limpiaCadena(request.querystring("calcost"))
	else
		calcost=limpiaCadena(request.form("calcost"))
	end if

	fila_ant	= cInt(null_z(limpiaCadena(Request.QueryString("fila_ant"))))
	if fila_ant ="" then
		fila_ant	= cInt(null_z(limpiaCadena(Request.form("fila_ant"))))
	end if

	DivisaAnt	= limpiaCadena(Request.QueryString("DivisaAnt"))
	if DivisaAnt ="" then
		DivisaAnt	= limpiaCadena(Request.form("DivisaAnt"))
	end if

	rendimiento=limpiaCadena(request.form("rendimiento"))

	if rendimiento>"" then
	else
		if request.form("h_rendimiento")>"" then
			rendimiento=limpiaCadena(request.form("h_rendimiento"))
		else
			rendimiento=limpiaCadena(request.querystring("h_rendimiento"))
		end if
	end if
	if rendimiento="" then rendimiento="articulos"

	if request.form("agruparart")>"" then
		agruparart=limpiaCadena(request.form("agruparart"))
	else
		agruparart=limpiaCadena(request.querystring("agruparart"))
	end if

	if request.form("agrupardoc")>"" then
		agrupardoc=limpiaCadena(request.form("agrupardoc"))
	else
		agrupardoc=limpiaCadena(request.querystring("agrupardoc"))
	end if

''ricardo 7-11-2007 se añade el tipo de articulo para poder filtrar
	if request.form("tipoarticulo")>"" then
		tipoarticulo=limpiaCadena(request.form("tipoarticulo"))
	else
		tipoarticulo=limpiaCadena(request.querystring("tipoarticulo"))
	end if

	if ncliente>"" then ncliente=session("ncliente") & ncliente
	if nproveedor>"" then nproveedor=session("ncliente") & nproveedor
    if ncenter>"" then ncenter=session("ncliente") & ncenter

	strwhere=""

	if mode="imp" then%>
		<table width='100%' cellspacing="1" cellpadding="1">
   			<tr>
				<td width="30%" align="left">
				</td>
				<td class=CELDARIGHT bgcolor="<%=color_blau%>">
					<%if fdesde>"" then
						if fhasta>"" then
						%><%=LitPeriodoFechas%> : <%=EncodeForHtml(fdesde)%> - <%=EncodeForHtml(fhasta)%><%
						else
							%><%=LitPeriodoFechas%> : <%=LitDesde%>&nbsp;<%=EncodeForHtml(fdesde)%><%
						end if
					else
						if fhasta>"" then
							%><%=LitPeriodoFechas%> : <%=LitHasta%>&nbsp;<%=EncodeForHtml(fhasta)%><%
						else
						end if
					end if%>
				</td>
	   		</tr>
		</table>
		<hr/>
    <%end if
	Alarma "rendimiento_articulos.asp"

	set rstSelect = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstAux2 = Server.CreateObject("ADODB.Recordset")
	set rst = Server.CreateObject("ADODB.Recordset")
	set conn=Server.CreateObject("ADODB.Connection")

	if mode="select1"then
		if ncliente & "">"" then
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("backendListados")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="select rsocial from clientes with(nolock) where ncliente=?"
            command.CommandType = adCmdText 'Procedimiento Almacenado
            command.Parameters.Append command.CreateParameter("@ncliente", adVarChar, adParamInput, 10, ncliente)

            set rstCustomer = command.execute

            if not rstCustomer.eof then
                nomcliente = rstCustomer("rsocial")
            end if

            rstCustomer.close
            conn.close
            set rstCustomer = nothing
            set command = nothing
            set conn = nothing
			if nomcliente & ""="" then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgClienteNoExiste%>");
				</script>
            <%end if
		end if
		if nproveedor & "">"" then
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("backendListados")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="select razon_social from proveedores with(nolock) where nproveedor=?"
            command.CommandType = adCmdText 'Procedimiento Almacenado
            command.Parameters.Append command.CreateParameter("@nproveedor", adVarChar, adParamInput, 10, nproveedor)

            set rstSupplier = command.execute

            if not rstSupplier.eof then
                nomproveedor = rstSupplier("razon_social")
            end if

            rstSupplier.close
            conn.close
            set rstSupplier = nothing
            set command = nothing
            set conn = nothing
			if nomproveedor & ""="" then
				%><script language="javascript" type="text/javascript">
				      window.alert("<%=LitMsgProveedorNoExiste%>");
				</script>
            <%end if
		end if
        
        if ncenter & "">"" then
            set conn = Server.CreateObject("ADODB.Connection")
            set command =  Server.CreateObject("ADODB.Command")
            conn.open session("backendListados")
            command.ActiveConnection =conn
            command.CommandTimeout = 0
            command.CommandText="select rsocial from centros with(nolock) where ncentro=? and ncliente=?"
            command.CommandType = adCmdText 'Procedimiento Almacenado
            command.Parameters.Append command.CreateParameter("@ncentro", adVarChar, adParamInput, 10, ncenter)
            command.Parameters.Append command.CreateParameter("@ncliente", adVarChar, adParamInput, 10, ncliente)

            set rstCenter = command.execute

            if not rstCenter.eof then
                nameCenter = rstCenter("rsocial")
            end if

            rstCenter.close
            conn.close
            set rstCenter = nothing
            set command = nothing
            set conn = nothing
			if nameCenter & ""="" then%>
                <script language="javascript" type="text/javascript">
                    window.alert("<%=LitMsgCentroNoEx%>");
				</script>
            <%end if
		end if%>

		
        <%
         DrawDiv "1","",""				
		 DrawLabel "","",LitPorArticulos%><input type="radio" name="rendimiento" value="articulos" <%=iif(rendimiento="articulos","checked","")%> onclick="tier2Menu('articulos')"/>
         <% CloseDiv
         DrawDiv "1","",""				
		 DrawLabel "","",LitPorDocumentos%><input type="radio" name="rendimiento" value="documentos" <%=iif(rendimiento="documentos","checked","")%> onclick="tier2Menu('documentos')"/><%
         CloseDiv%><input type='hidden' name='h_rendimiento' value='<%=EncodeForHtml(rendimiento)%>'/><br/>       
        <h6 class="col-lg-12 col-md-12 col-sm-12 col-xs-12" ><%=LitDocVentas%></h6>        
	    
		
        <%
            EligeCelda "input","add","left","","",0,LitDesdeFecha,"fdesde",0,EncodeForHtml(iif(fdesde>"",fdesde,"01/01/" & year(date)))
            DrawCalendar "fdesde"
			
            EligeCelda "input","add","left","","",0,LitHastaFecha,"fhasta",0,EncodeForHtml(iif(fhasta>"",fhasta,day(date) & "/" & month(date) & "/" & year(date)))
            DrawCalendar "fhasta"
		
            DrawDiv "1", "", ""
            DrawLabel "", "", LitIncluirFacturas
            rstSelect.CursorLocation=3
				rstSelect.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as nombre from series with(NOLOCK) where tipo_documento ='FACTURA A CLIENTE' and nserie like '" & session("ncliente") & "%' order by nombre ",session("backendListados")%><select name="seriesFacturas" align='left' class="width60" multiple="multiple" id="seriesFactura">
					
                    <%do while not rstSelect.Eof%>
						<option <%=iif(rstSelect("nserie")>"","selected","")%> value="<%=EncodeForHtml(rstSelect("nserie"))%>"> <%=EncodeForHtml(rstSelect("nombre"))%></option>
                        <%rstSelect.moveNext
					loop%>
                        <option value=''></option>
					</select>
                <%rstSelect.close%>				
                <%
             CloseDiv 
                   
             DrawDiv "1", "", ""
             DrawLabel "", "", LitAlbaranesPF       
             rstSelect.CursorLocation=3
		     rstSelect.open "select nserie, case when datalength(right(nserie,len(nserie)-5)+' '+nombre)<=21 then right(nserie,len(nserie)-5)+'-'+nombre else left(right(nserie,len(nserie)-5)+'-'+nombre,20)+'...' end as nombre from series with(NOLOCK) where tipo_documento ='ALBARAN DE SALIDA' and nserie like '" & session("ncliente") & "%' order by nombre ",session("backendListados")%><select name="seriesAlbaranes" align='left' class="width60" multiple="multiple" id="seriesAlbaranes">
					
                    <%do while not rstSelect.Eof%><option value="<%=EncodeForHtml(rstSelect("nserie"))%>"> <%=EncodeForHtml(rstSelect("nombre"))%></option>						
                        <%rstSelect.moveNext
					loop%><option value=''></option>                        
					</select>
                <%rstSelect.close
            CloseDiv
            if tiene_modulo_tiendas <> 0 then
               
            DrawDiv "1", "", ""
            DrawLabel "", "", LitTicketsPF    
                
                '**RGU 13/9/2007: Deben salir todas las series de tickets y no solo las que tengan pendientes de facturar
                rstSelect.CursorLocation=3				
				rstSelect.open "select S.nserie as nserie, S.nombre as nombre from series S with(nolock)  where S.tipo_documento='TICKET' and S.nserie like '" & session("ncliente") & "%' order by nombre", session("backendListados")%><select class="width60" name="seriesTPF" align='left'  multiple="multiple" id="selecttickets">
                    <%do while not rstSelect.Eof%>
						<option value="<%=EncodeForHtml(rstSelect("nserie"))%>"> <%=EncodeForHtml(rstSelect("nombre"))%></option>
                        <%rstSelect.moveNext
					loop%><option value=''></option>
					</select>
                <%rstSelect.close%>				
                <%
            CloseDiv
			end if%><br/>
		    <hr />		
        <%
			dim ConfigDespleg (3,13)

			i=0
			ConfigDespleg(i,0)="categoria"
			ConfigDespleg(i,1)=""
			ConfigDespleg(i,2)="5"
			ConfigDespleg(i,3)="select codigo, nombre from categorias with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre"
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
			ConfigDespleg(i,2)="5"
			ConfigDespleg(i,3)="select codigo, nombre,categoria from familias_padre with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre"
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
			ConfigDespleg(i,2)="5"
			ConfigDespleg(i,3)="select codigo, nombre,categoria,padre from familias with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by nombre"
			ConfigDespleg(i,4)=1
			ConfigDespleg(i,5)="width60"
			ConfigDespleg(i,6)="MULTIPLE"
			ConfigDespleg(i,7)="codigo"
			ConfigDespleg(i,8)="nombre"
			ConfigDespleg(i,9)=LitSubFamilia
			ConfigDespleg(i,10)=familia
			ConfigDespleg(i,11)=""
			ConfigDespleg(i,12)=""

			DibujaDesplegables ConfigDespleg,session("backendListados")
			
        DrawDiv "1", "", ""
        DrawLabel "", "", LitVerCods_CFS
				if verCodCFS="on" then
					%><input type='checkbox' name='verCodCFS' checked="checked" /><%
				else
					%><input type='checkbox' name='verCodCFS'/><%
				end if			
        CloseDiv
			
        DrawDiv "1", "", ""
        DrawLabel "", "", LitCodigo%><input class='width15' type="text" name="ncliente" value="<%=EncodeForHtml(trimcodempresa(ncliente))%>" onchange="TraerCliente('<%=enc.EncodeForJavascript(null_s(mode))%>');"/><a class='CELDAREFB' href="javascript:AbrirVentana('../../ventas/clientes_buscar.asp?ndoc=rendimiento_articulos&titulo=<%=LitSelCliente%>&mode=search&viene=rendimiento_articulos','P',<%=altoventana%>,<%=anchoventana%>);" onmouseover="self.status='<%=LitVerCliente%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' type="text" name="nombre" disabled="disabled" value="<%=EncodeForHtml(nomcliente)%>"  />
                <%if nomcliente & "" = "" then%>
                    <script language="javascript">document.rendimiento_articulos.ncliente.value = "";</script>
                <%end if
        CloseDiv

        DrawDiv "1", "", ""
        DrawLabel "", "", LITTIPACTIV%><select class='width60' name="actividad">
			        <%rstSelect.open "select codigo,descripcion from tipo_actividad with(NOLOCK) where codigo like '" & session("ncliente") & "%' order by descripcion",Session("backendlistados")
			        while not rstSelect.eof%>
			            <option value="<%=EncodeForHtml(rstSelect("codigo"))%>"><%=EncodeForHtml(rstSelect("descripcion"))%></option>
			            <%rstSelect.movenext
			        wend
			        rstSelect.close%>
			        <option selected="selected" value=''></option>
			    </select>
			<%			
        CloseDiv
        if has_module_advanced_management <> 0 or has_module_postventa <> 0 then        
        DrawDiv "1", "", ""
        DrawLabel "", "", LitCentro%><input class='width15' type="text" name="ncenter" value="<%=EncodeForHtml(trimCodEmpresa(ncenter))%>" size='10' onchange="GetCenter('<%=mode%>', 0);"/><a class='CELDAREFB' href="javascript:GetCenter('', 1);" onmouseover="self.status='<%=LitVerCentro%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' type="text" name="nameCenter" disabled="disabled" value="<%=EncodeForHtml(nameCenter)%>" />
			
            <%if nameCenter & "" = "" then%>
                <script language="javascript">document.rendimiento_articulos.ncenter.value = "";</script>
            <%end if
		
        CloseDiv
        end if
		
        DrawDiv "1", "", ""
        DrawLabel "", "", LitProveedor%><input class='width15' type="text" name="nproveedor" value="<%=EncodeForHtml(trimCodEmpresa(nproveedor))%>"  onchange="TraerProveedor('<%=enc.EncodeForJavascript(null_s(mode))%>');"/><a class='CELDAREFB' href="javascript:AbrirVentana('../../compras/proveedores_busqueda.asp?ndoc=rendimiento_articulos&titulo=<%=LitSelProveedor%>&mode=search&viene=rendimiento_articulos','P',<%=altoventana%>,<%=anchoventana%>);" onmouseover="self.status='<%=LitVerProveedor%>'; return true;" onmouseout="self.status=''; return true;"><img src="<%=ImgBuscarDinamic%>" <%=ParamImgBuscar%> alt="<%=LitBuscar%>" title="<%=LitBuscar%>"/></a><input class='width40' type="text" name="nomproveedor" disabled="disabled" value="<%=EncodeForHtml(nomproveedor)%>" /><%    
        CloseDiv
		if si_tiene_modulo_proyectos<>0 then%><div CLASS="col-lg-4 col-md-2 col-sm-3 col-xs-6"><label><%=LitProyecto%></label><input  type="hidden" name="cod_proyecto" value="<%=EncodeForHtml(cod_proyecto)%>"/><iframe id='frProyecto' class="width60 iframe-menu" src='../../mantenimiento/docproyectos_responsive.asp?viene=rendimiento_articulos&mode=<%=EncodeForHtml(mode)%>&cod_proyecto=<%=EncodeForHtml(cod_proyecto)%>'  frameborder=0 scrolling="no"></iframe></div><%			
		end if
		
			rstAux.CursorLocation=3
			rstAux.open "select codigo, descripcion from tipos_entidades with(NOLOCK) where codigo like '" & session("ncliente") & "%' and tipo='CLIENTE' order by descripcion",session("backendListados")
	       	DrawSelectCelda "","180","",0,LitTipoCliente,"tipoCliente",rstAux,tipoCliente,"codigo","descripcion","",""
		   	rstAux.close
						
			rstAux.CursorLocation=3
			rstAux.open "select dni, nombre from personal with(NOLOCK),comerciales with(NOLOCK) where comerciales.fbaja is null and dni=comercial and dni like '" & session("ncliente") & "%' order by nombre",session("backendListados")
			if si_tiene_modulo_comercial<>0 then				
                DrawSelectCelda "","","",0,LitComercialModCom,"comercial",rstAux,comercial,"dni","nombre","",""
			else				
                DrawSelectCelda "","","",0,LitComercial,"comercial",rstAux,comercial,"dni","nombre","",""
			end if                    
			rstAux.close			
            EligeCelda "input","add","left","","",0,LitConref,"referencia",0,EncodeForHtml(referencia)			
            EligeCelda "input","add","left","","",0,LitConNombre,"nombreart",0,EncodeForHtml(nombreart)
		stilo=""
		if rendimiento="articulos" then
	    	stilo="none"
		end if
		DrawDiv "1", "display:"&stilo&"", "filacostes" 
			%><input type="hidden" name="coste" value="DOCUMENTOS"/><%
		CloseDiv
		
        DrawDiv "1", "", ""
        DrawLabel "", "", LitOrdenar%><span id="ordarticulos" style="display:<%=iif(rendimiento="articulos","","none")%>"><select  class="width60" name="ordenar"><%if ordenar="REFERENCIA" then%>
							<option value="NOMBRE"><%=Ucase(LitNombre)%></option>
				   			<option selected="selected" value="REFERENCIA"><%=Ucase(LitReferencia)%></option>
							<option value="MAYOR BENEFICIO"><%=Ucase(LitMayorBeneficio)%></option>
						<%elseif ordenar="MAYOR BENEFICIO" then%>
							<option value="NOMBRE"><%=Ucase(LitNombre)%></option>
							<option value="REFERENCIA"><%=Ucase(LitReferencia)%></option>
		   					<option selected="selected" value="MAYOR BENEFICIO"><%=Ucase(LitMayorBeneficio)%></option>
						<%else%>
							<option selected="selected" value="NOMBRE"><%=Ucase(LitNombre)%></option>
							<option value="REFERENCIA"><%=Ucase(LitReferencia)%></option>
		   					<option value="MAYOR BENEFICIO"><%=Ucase(LitMayorBeneficio)%></option>
						<%end if %><option value="<%=Ucase(LitValueCodCFS)%>"><%=Ucase(LitCodCFS)%></option></select></span><%
                        %><span id="orddocumentos" style="display:<%=iif(rendimiento="articulos","none","")%>"><select class='width60' name="ordenarDOC">
						<%if ordenarDOC="CLIENTE" then%>
							<option value="FECHA"><%=Ucase(LitFecha)%></option>
			   				<option selected="selected" value="CLIENTE"><%=Ucase(LitCliente)%></option>
							<option value="MAYOR BENEFICIO"><%=Ucase(LitMayorBeneficio)%></option>
						<%elseif ordenarDOC="MAYOR BENEFICIO" then%>
							<option value="FECHA"><%=Ucase(LitFecha)%></option>
							<option value="CLIENTE"><%=Ucase(LitCliente)%></option>
			   				<option selected="selected" value="MAYOR BENEFICIO"><%=Ucase(LitMayorBeneficio)%></option>
						<%else%>
							<option selected="selected" value="FECHA"><%=Ucase(LitFecha)%></option>
							<option value="CLIENTE"><%=Ucase(LitCliente)%></option>
			   				<option value="MAYOR BENEFICIO"><%=Ucase(LitMayorBeneficio)%></option>
						<%end if %></select></span>			
			<%CloseDiv
            
            DrawDiv "1", "", ""
            DrawLabel "", "", LitAgrupar%><span id="agrarticulos" style="display:<%=iif(rendimiento="articulos" or rendimiento="","","none")%>"><select class='width60' name="agruparart" onchange="javascript:cambiarNoDetalle('1');"><option <%=iif(agruparart="CLIENTE" or agrupar="","selected","")%> value="CLIENTE"><%=Ucase(LitCliente)%></option>
						<%if si_tiene_modulo_comercial<>0 then%>
							<option <%=iif(agruparart="COMERCIAL","selected","")%> value="COMERCIAL"><%=Ucase(LitComercialModCom)%></option>
						<%else%>
							<option <%=iif(agruparart="RESPONSABLE","selected","")%> value="RESPONSABLE"><%=Ucase(LitComercial)%></option>
						<%end if%>
						<option <%=iif(agruparart="CATEGORÍA","selected","")%> value="CATEGORÍA"><%=Ucase(LitCategoria)%></option>
						<option <%=iif(agruparart="FAMILIA","selected","")%> value="FAMILIA"><%=Ucase(LitFamilia)%></option>
						<option <%=iif(agruparart="SUBFAMILIA","selected","")%> value="SUBFAMILIA"><%=Ucase(LitSubFamilia)%></option>
						<option <%=iif(agruparart="ARTICULO","selected","")%> value="ARTICULO"><%=Ucase(LitArticulo)%></option>
						
						<%if si_tiene_modulo_proyectos<>0 then%>
							<option <%=iif(agruparart="PROYECTO","selected","")%> value="PROYECTO"><%=Ucase(LitProyecto)%></option>
						<%end if%></select></span><%
                        %><span id="agrdocumentos" style="display:<%=iif(rendimiento="articulos","none","")%>"><select  class='width60' name="agrupardoc" onchange="javascript:cambiarNoDetalle('2');"><option <%=iif(agrupardoc="CLIENTE" or agrupar="","selected","")%> value="CLIENTE"><%=Ucase(LitCliente)%></option>
						<%if si_tiene_modulo_comercial<>0 then%>
							<option <%=iif(agrupardoc="COMERCIAL","selected","")%> value="COMERCIAL"><%=Ucase(LitComercialModCom)%></option>
						<%else%>
							<option <%=iif(agrupardoc="RESPONSABLE","selected","")%> value="RESPONSABLE"><%=Ucase(LitComercial)%></option>
						<%end if%>
						
						<%if si_tiene_modulo_proyectos<>0 then%>
							<option <%=iif(agrupardoc="PROYECTO","selected","")%> value="PROYECTO"><%=Ucase(LitProyecto)%></option>
						<%end if%></select></span>
			<%CloseDiv
					
            DrawDiv "1","","id_nodetallado"
            DrawLabel "", "", LITLISTDETALLADO%><input type='checkbox' name='nodetallado' value="on" onclick="cambiarNoDetalle('3');"/><span id="id_nodetallado2" " align="right" style='width:180px;border:0px solid black;display:none;'></span>			   
			<%CloseDiv          
       
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
            DrawSelectMultipleCelda "width60","","",0,LitTipArtRendArt,"tipoarticulo",rstArtType,EncodeForHtml(tipoarticulo),"codigo","descripcion","",""
            rstArtType.close
            conn.close
            set rstArtType = nothing
            set command = nothing
            set conn = nothing
		
        DrawDiv "1", "", "bajaarticulos"
        DrawLabel "", "", LitNoIncArtBaja			
					if artbaja="on" then
						%><input type='checkbox' name='artbaja' checked="checked" /><%
					else
						%><input type='checkbox' name='artbaja'/><%
					end if
	    CloseDiv
		
        DrawDiv "1", "", "agrdocumentos2"
        DrawLabel "", "", LitIncPedSinFacturar
			
					if pedsinf="on" then
						%><input type='checkbox' name='pedsinf' checked="checked" /><%
					else
						%><input type='checkbox' name='pedsinf'/><%
					end if				
        CloseDiv
		if si_tiene_modulo_mantenimiento<>0 then
			
        DrawDiv "1", "", "agrdocumentos3"
        DrawLabel "", "", LitIncOrdenesSinFacturar
				
						if ordsinf="on" then
							%><input type='checkbox' name='ordsinf' checked="checked" /><%
						else
							%><input type='checkbox' name='ordsinf'/><%
						end if					
        CloseDiv
		end if
	
        DrawDiv "1", "", ""
        DrawLabel "", "", LitCalCostEnFuncDe%><select name="calcost" class="width60"><option <%=iif(calcost="Coste de la venta" or calcost="","selected","")%> value="Coste de la venta"><%=LitCalCost1%></option><option <%=iif(calcost="Coste medio","selected","")%> value="Coste medio"><%=LitCalCost2%></option></select>			
		<%		
        CloseDiv
		elseif mode="imp" then
''ricardo 25-5-2006 comienzo de la select
''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"inicio_rend_articulos"

		seriestpf=	   cstr(Request.Form("seriesTPF"))

		lista_series_tpf ="('"
		lista_series_tpf2 ="('"
		''if c_tickets="on" and seriestpf > "" then
		if seriestpf & "" > "" then
			if instr(seriestpf,",")>0 then
				lista_series_tpf = lista_series_tpf & replace(replace(seriestpf," ",""),",","','")
				lista_series_tpf2 = lista_series_tpf2 & replace(replace(seriestpf," ",""),",","','")

			else
				lista_series_tpf = lista_series_tpf & seriestpf
				lista_series_tpf2 = lista_series_tpf2 & seriestpf
			end if
		end if
		if right(lista_series_tpf,3) <> "','" then
			lista_series_tpf= lista_series_tpf & "','"
		end if
		if right(lista_series_tpf2,3) <> "','" then
			lista_series_tpf2= lista_series_tpf2 & "','"
		end if
		lista_series_tpf = lista_series_tpf +"WvWvW')"
		lista_series_tpf2 = lista_series_tpf2 +"WvWvW')"


'ahora para series de facturas
		seriesfact=	   cstr(Request.Form("seriesFacturas"))
		lista_series_fact ="('"
		''if c_facturas="on" and seriesfact > "" then
		if seriesfact & "">"" then
			lista_series_fact = lista_series_fact & replace(replace(seriesfact," ",""),",","','")
		end if
		if right(lista_series_fact,3) <> "','" then
			lista_series_fact= lista_series_fact & "','"
		end if
		lista_series_fact = lista_series_fact & "WvWvW')"
		
'ahora para series de albaranes
		seriesalb=	   cstr(Request.Form("seriesAlbaranes"))
		lista_series_alb ="('"
		if seriesalb & "" > "" then
			lista_series_alb = lista_series_alb & replace(replace(seriesalb," ",""),",","','")
		end if
		if right(lista_series_alb,3) <> "','" then
			lista_series_alb= lista_series_alb & "','"
		end if
		lista_series_alb = lista_series_alb & "WvWvW')"

		MAXPAGINA=d_lookup("maxpagina", "limites_listados", "item='152'", DSNIlion)
		MAXPDF=d_lookup("maxpdf", "limites_listados", "item='152'", DSNIlion)
        rstAux.cursorlocation=3
		rstAux.open "select * from divisas with(NOLOCK) where moneda_base<>0 and codigo like '" & session("ncliente") & "%'",session("backendListados")
        if not rstAux.eof then
		    MB=rstAux("codigo")
		    AbMB=rstAux("abreviatura")
		    DeMB=rstAux("ndecimales")
		    FcMB=rstAux("factcambio")
        end if
		rstAux.close%>

		<input type='hidden' name='maxpdf'          value='<%=EncodeForHtml(MAXPDF)%>'/>
		<input type='hidden' name='maxpagina'       value='<%=EncodeForHtml(MAXPAGINA)%>'/>
		<input type="hidden" name="fdesde"          value="<%=EncodeForHtml(fdesde)%>"/>
		<input type="hidden" name="fhasta"          value="<%=EncodeForHtml(fhasta)%>"/>
		<input type="hidden" name="nserie"          value="<%=EncodeForHtml(nserie)%>"/>
		<input type="hidden" name="ncliente"        value="<%=EncodeForHtml(ncliente)%>"/>
		<input type="hidden" name="tipoCliente"     value="<%=EncodeForHtml(tipoCliente)%>"/>
		<input type="hidden" name="nproveedor"      value="<%=EncodeForHtml(nproveedor)%>"/>
		<input type="hidden" name="cod_proyecto"    value="<%=EncodeForHtml(cod_proyecto)%>"/>
		<input type="hidden" name="familia"         value="<%=EncodeForHtml(familia)%>"/>
		<input type="hidden" name="familia_padre"   value="<%=EncodeForHtml(familia_padre)%>"/>
		<input type="hidden" name="categoria"       value="<%=EncodeForHtml(categoria)%>"/>
		<input type="hidden" name="comercial"       value="<%=EncodeForHtml(comercial)%>"/>
		<input type="hidden" name="referencia"      value="<%=EncodeForHtml(referencia)%>"/>
		<input type="hidden" name="nombreart"       value="<%=EncodeForHtml(nombreart)%>"/>
		<input type="hidden" name="coste"           value="<%=EncodeForHtml(coste)%>"/>
		<input type="hidden" name="ordenar"         value="<%=EncodeForHtml(ordenar)%>"/>
		<input type="hidden" name="ordenardoc"      value="<%=EncodeForHtml(ordenardoc)%>"/>
		<input type="hidden" name="verCodCFS"       value="<%=EncodeForHtml(verCodCFS)%>"/>
		<input type="hidden" name="artbaja"         value="<%=EncodeForHtml(artbaja)%>"/>
		<input type="hidden" name="calcost"         value="<%=EncodeForHtml(calcost)%>"/>
		<input type="hidden" name="pedsinf"         value="<%=EncodeForHtml(pedsinf)%>"/>
		<input type="hidden" name="mode"            value="<%=EncodeForHtml(mode)%>"/>
		<input type="hidden" name="agruparart"      value="<%=EncodeForHtml(agruparart)%>"/>
		<input type="hidden" name="agrupardoc"      value="<%=EncodeForHtml(agrupardoc)%>"/>
		<input type="hidden" name="tipoarticulo"    value="<%=EncodeForHtml(tipoarticulo)%>"/>
		<input type="hidden" name="actividad"       value="<%=EncodeForHtml(actividad)%>"/>
		<input type="hidden" name="nodetallado"     value="<%=EncodeForHtml(nodetallado)%>"/>
        <input type="hidden" name="ncenter"         value="<%=EncodeForHtml(ncenter)%>"/>

        <%VinculosPagina(MostrarClientes)=1:VinculosPagina(MostrarArticulos)=1
		VinculosPagina(MostrarProveedores)=1:VinculosPagina(MostrarCentros)=1
		CargarRestricciones session("usuario"),session("ncliente"),Permisos,Enlaces,VinculosPagina%>
		<font class='CELDA7'><b><%=LitPorArticulos%></b></font><br/>

        <%set conn = Server.CreateObject("ADODB.Connection")
        set command =  Server.CreateObject("ADODB.Command")
        conn.open session("backendListados")
        command.ActiveConnection =conn
        command.CommandTimeout = 0
        command.CommandText="ListadoRendimientoArt"
        command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
        command.Parameters.Append command.CreateParameter("@fdesde", adVarChar, adParamInput, 25, fdesde)
        command.Parameters.Append command.CreateParameter("@fhasta", adVarChar, adParamInput, 25, fhasta & " 23:59:59")
        command.Parameters.Append command.CreateParameter("@serie", adVarChar, adParamInput, -1, lista_series_fact)
        command.Parameters.Append command.CreateParameter("@ncliente", adChar, adParamInput, 10, ncliente)
        command.Parameters.Append command.CreateParameter("@nproveedor", adChar, adParamInput, 10, nproveedor)
        command.Parameters.Append command.CreateParameter("@cod_proyecto", adVarChar, adParamInput, 15, cod_proyecto)
        command.Parameters.Append command.CreateParameter("@tipoCliente", adVarChar, adParamInput, 10, tipoCliente)
        command.Parameters.Append command.CreateParameter("@familia", adVarChar, adParamInput, -1, familia)
        command.Parameters.Append command.CreateParameter("@familia_padre", adVarChar, adParamInput, -1, familia_padre)
        command.Parameters.Append command.CreateParameter("@categoria", adVarChar, adParamInput, -1, categoria)
        command.Parameters.Append command.CreateParameter("@comercial", adVarChar, adParamInput, 20, comercial)
        command.Parameters.Append command.CreateParameter("@conReferencia", adVarChar, adParamInput, 30, referencia)
        command.Parameters.Append command.CreateParameter("@conNombre", adVarChar, adParamInput, 100, nombreart)
        command.Parameters.Append command.CreateParameter("@ordenarPorArt", adVarChar, adParamInput, 20, ordenar)
        command.Parameters.Append command.CreateParameter("@agruparArt", adVarChar, adParamInput, 20, iif(agruparart="SUBFAMILIA","FAMILIA",iif(agruparart="FAMILIA","FAMILIA_PADRE",agruparart)))
        command.Parameters.Append command.CreateParameter("@articulosBaja", adBoolean, adParamInput, , iif(artbaja>"",1,0))
        command.Parameters.Append command.CreateParameter("@calculoCoste", adVarChar, adParamInput, 20, calcost)
        command.Parameters.Append command.CreateParameter("@usuario", adVarChar, adParamInput, 50, session("usuario"))
        command.Parameters.Append command.CreateParameter("@sesion_ncliente", adVarChar, adParamInput, 5, session("ncliente"))
        command.Parameters.Append command.CreateParameter("@lista_series_tpf", adVarChar, adParamInput, -1, lista_series_tpf2)
        ''ricardo 7-11-2007 se añade el tipo de articulo para poder filtrar
        command.Parameters.Append command.CreateParameter("@tipo_articulo", adVarChar, adParamInput, -1, tipoarticulo)
        command.Parameters.Append command.CreateParameter("@lista_series_apf", adVarChar, adParamInput, -1, lista_series_alb)
        command.Parameters.Append command.CreateParameter("@actividad", adVarChar, adParamInput, 20, actividad)
        command.Parameters.Append command.CreateParameter("@nodetallado", adVarChar, adParamInput, 20, iif(ucase(nodetallado)="ON",1,0))
        ''MPC 14/05/2013 Se añade el centro para filtrar
        command.Parameters.Append command.CreateParameter("@ncenter", adVarChar, adParamInput, 10, ncenter)

		lote=limpiaCadena(Request.QueryString("lote"))
		if lote="" then
            set rstAux = command.execute
            conn.close
            set command = nothing
            set conn = nothing
		end if

        strordercompleto="order by "
		strorder="order by "
		if agruparart="CLIENTE" then
			strorder=strorder & " nomCliente,ncliente"
            strordercompleto=strordercompleto & " nomCliente,ncliente"
		elseif agruparart="COMERCIAL" or agruparart="RESPONSABLE" then
			strorder=strorder & " nomComercial,comercial"
            strordercompleto=strordercompleto & " nomComercial,comercial"
		elseif agruparart="SUBFAMILIA" and Ucase(ordenar)<>Ucase(LitValueCodCFS) then
			strorder=strorder & " nomCategoria,categoria"
            strordercompleto=strordercompleto & " nomCliente,ncliente"
		elseif agruparart="SUBFAMILIA" and Ucase(ordenar)=Ucase(LitValueCodCFS) then
			strorder=strorder & " categoria,divisa,familia_padre,familia"
            strordercompleto=strordercompleto & " nomCliente,ncliente,divisa"
		elseif agruparart="FAMILIA"  and Ucase(ordenar)<>Ucase(LitValueCodCFS) then
			strorder=strorder & " nomCategoria,categoria"
            strordercompleto=strordercompleto & " nomCliente,ncliente,divisa"
		elseif agruparart="FAMILIA"  and Ucase(ordenar)=Ucase(LitValueCodCFS) then
			strorder=strorder & " categoria,divisa,familia_padre"
            strordercompleto=strordercompleto & " nomCliente,ncliente,divisa"
		elseif agruparart="CATEGORÍA" and Ucase(ordenar)<>Ucase(LitValueCodCFS) then
			strorder=strorder & " nomCategoria,categoria"
            strordercompleto=strordercompleto & " nomCliente,ncliente,divisa"
		elseif agruparart="CATEGORÍA" and Ucase(ordenar)=Ucase(LitValueCodCFS) then
			strorder=strorder & " categoria,divisa"
            strordercompleto=strordercompleto & " nomCliente,ncliente,divisa"
		elseif agruparart="PROYECTO" then
			strorder=strorder & " nomProyecto,cod_proyecto"
            strordercompleto=strordercompleto & " nomProyecto,cod_proyecto"
		elseif agruparart="ARTICULO" and Ucase(ordenar)<>Ucase(LitValueCodCFS) and Ucase(ordenar)<>Ucase(LitReferencia) then
			strorder=strorder & " descripcion,referencia"
            strordercompleto=strordercompleto & " nomCliente,ncliente,divisa"
		elseif agruparart="ARTICULO" and Ucase(ordenar)=Ucase(LitValueCodCFS) then
			strorder=strorder & " descripcion,divisa"
            strordercompleto=strordercompleto & " nomCliente,ncliente,divisa"
		end if

		if Ucase(ordenar)<>Ucase(LitValueCodCFS) and Ucase(ordenar)<>Ucase(LitReferencia) then
		    strorder=strorder & ",divisa"
		end if

		if ordenar="NOMBRE" then
			if agruparart="CATEGORÍA" or agruparart="SUBFAMILIA" or agruparart="FAMILIA" then
				strorder=strorder & ",nomFamiliaPadre"
                strordercompleto=strordercompleto & ""
			else
                if agruparart<>"ARTICULO" then
				    strorder=strorder & ",descripcion"
                end if
                strordercompleto=strordercompleto & ",descripcion"
                
			end if
		elseif ordenar="REFERENCIA" then
			if agruparart<>"FAMILIA" AND agruparart<>"CATEGORÍA" AND agruparart<>"SUBFAMILIA" then
				strorder=strorder & "referencia"
                strordercompleto=strordercompleto & ",referencia"
			end if
		elseif ordenar="MAYOR BENEFICIO" then
            if agruparart="ARTICULO" then
                strorder= " order by beneficio desc"
            else
    			strorder=strorder & ",beneficio desc"
            end if
            strordercompleto=strordercompleto & "beneficio desc"
		elseif Ucase(ordenar)=Ucase(LitValueCodCFS) then
			if agruparart=ucase(LitComercial) or agruparart=ucase(LitComercialModCom) or agruparart="CLIENTE" _
			or agruparart="PROYECTO" then
				strorder=strorder & ",divisa,familia"
                strordercompleto=strordercompleto & ",divisa"
			end if
		end if
		rst.cursorlocation=3
		if agruparart="FAMILIA" then
			strSQL="select distinct categoria,nomcategoria,nomfamiliapadre,familia_padre,divisa,ndecimales,abreviatura,totalcostes as costes,totalventas as ventas,totalunidades as unidades,totalbeneficio as beneficio, "
			strSQL=strSQL & "case when totalcostes>0 then ((Totalbeneficio*100)/Totalcostes) else 0 end as BenCompras,case when TotalVentas>0 then ((TotalBeneficio*100)/TotalVentas) else 0 end as BenVentas "
			strSQL=strSQL & "from [" & session("usuario") & "] " & strorder
			rst.open strSQL,session("backendListados")
		else
            if ucase(nodetallado)<>"ON" then
			    rst.open "select * from [" & session("usuario") & "] " & strorder,session("backendListados")
            else
                rst.open "select * from [" & session("usuario") & "] " & strordercompleto,session("backendListados")
            end if
		end if
	    NUMREGISTROS=rst.Recordcount
		%><input type="hidden" name="NumRegsTotal" value="<%=EncodeForHtml(NUMREGISTROS)%>"/><%

		if rst.EOF then
			rst.Close
			%><script language="javascript" type="text/javascript">
			      window.alert("<%=LitMsgDatosNoExiste%>");
			      parent.botones.document.location = "rendimiento_articulos_bt.asp?mode=select1";
			      document.rendimiento_articulos.action = "rendimiento_articulos.asp?mode=select1";
			      document.rendimiento_articulos.submit();
			</script><%
		else
			MuestraParamSelec()
			if agruparart="FAMILIA" then
				MuestraListadoFam()
			elseif agruparart="CATEGORÍA" then
				MuestraListadoCat()
			elseif agruparart="SUBFAMILIA" then
				MuestraListadoSubf()
			else
				MuestraListado()
			end if
		end if

''ricardo 25-5-2006 comienzo de la select
''usuario,ndocumento,npersona,accion,referencia,nserie,tipo
auditar_ins_bor session("usuario"),"","","auditoria_listados",request.form,request.querystring,"fin_rend_articulos"
	end if%>
</form>
    <%set rstSelect = Nothing
	set rstAux = Nothing
	set rstAux2 = Nothing
	set rst = Nothing
	set conn=Nothing
    connRound.close
    set connRound = Nothing
end if%>
</body>
</html>