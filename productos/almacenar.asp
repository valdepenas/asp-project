<%@ Language=VBScript %>
<%' JCI 17/06/2003 : MIGRACION A MONOBASE'%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="<%=session("lenguaje")%>">
<head>
<title><%=LITTITULOALMACART%></title>

<meta http-equiv="Content-Type" content="text/html"; charset="<%=session("caracteres")%>"/>
</head>

<!--#include file="../constantes.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../calculos.inc" -->
<%if accesoPagina(session.sessionid,session("usuario"))=1 then %>

<!--#include file="../ilion.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->

<!--#include file="../compras/compras.inc" -->
<!--#include file="articulos.inc" -->

<!--#include file="../tablas.inc" -->
<!--#include file="../styles/formularios.css.inc" -->

<script language="javascript" type="text/javascript" src="../jfunciones.js"></script>
<script language="javascript" type="text/javascript">

function Valido(obj,linea,campo) {
	if (isNaN(obj.value) || obj.value=="") {
		window.alert ("<%=LITCAMPALMACART%>" + campo + "<%=LITDEBSNUMALMACART%>");
		obj.value="0";
	}
}
</script>

<body class="BODY_ASP MARGEN NoMargin" style="overflow-x:hidden;">

<%
auditoriaMM=""
sub ActualizaCosteMedio(TmpReferencia)
	' Actualizamos el coste medio en caso de que el artículo no aparezca en ningún documento de compras'
	connAlmacenar.open session("dsn_cliente")
	strselect="EXEC ActCosteMed @referencia='" & TmpReferencia & "'"
	set rst = connAlmacenar.execute(strselect)
	connAlmacenar.close
end sub

sub ActualizaStockPadreAlm(stold,st,rpad,alm)
	' Actualizamos el stock en caso de que el artículo sea hijo de tallas y colores'
	if not isnumeric(stold) then
		stold=0
	end if
	if not isnumeric(st) then
		st=0
	end if
	stock_nuevo=st-stold
	if not isnumeric(stock_nuevo) then
		stock_nuevo=0
	end if
	strselect="update almacenar set stock=stock + " & miround(stock_nuevo,DEC_CANT)
	strselect=strselect & " where articulo='" & rpad & "' and almacen='" & alm & "'"
	rstDet.open strselect,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
	if rstDet.state<>0 then rstDet.close
end sub

'****************************************************************************************************************'
'*************************************  CODIGO PRINCIPAL DE LA PAGINA  ******************************************'
'****************************************************************************************************************'
	set connRound = Server.CreateObject("ADODB.Connection")
	connRound.open dsnilion
%><form name="Almacenar" method="post"><%
	set rst = Server.CreateObject("ADODB.Recordset")
	set rstDet = Server.CreateObject("ADODB.Recordset")
	set rstAux = Server.CreateObject("ADODB.Recordset")
	set rstMM = Server.CreateObject("ADODB.Recordset")
	set connAlmacenar = Server.CreateObject("ADODB.Connection")

	WaitBoxOculto LitEsperePorFavor

	mode = Request.QueryString("mode")
	p_referencia=limpiaCadena(request.querystring("referencia"))
	CheckCadena p_referencia
	Nregs=limpiaCadena(request.form("hNregs"))
	predet=limpiaCadena(request.form("predet"))
	ps=limpiaCadena(request.querystring("ps"))
	''ricardo 11-12-2007 el ps sera siempre No para todos los campos excepto pendientes servir, recibir y minimo
	''deshabilitar=iif(ps="NO","readonly","")
	deshabilitar="readonly"
	deshabilitarpsp=iif(ps="NO","readonly","")
	
	'mmg >> obtenemos el nombre del usuario
    if mode="save" then
        strselect="select nombre from indice where entrada='"+session("usuario")+"'"
        rstMM.open strselect,DSNIlion,adOpenKeyset, adLockOptimistic
    
        nombreMM=rstMM("nombre")
        rstMM.Close()
    end if
	
	if mode="addalmacen" then
		almacen=limpiaCadena(request.querystring("almacen"))
		stock=limpiaCadena(request.querystring("stock"))
		smin=limpiaCadena(request.querystring("smin"))
        smax=limpiaCadena(request.querystring("smax"))
		reposicion=limpiaCadena(request.querystring("reposicion"))
		precibir=limpiaCadena(request.querystring("precibir"))
		pservir=limpiaCadena(request.querystring("pservir"))
		pmin=limpiaCadena(request.querystring("pmin"))
		ubicacion=limpiaCadena(request.querystring("ubicacion"))
		talla=limpiaCadena(request.querystring("talla"))
		color=limpiaCadena(request.querystring("color"))
	end if

	if mode>"" then
		if mode="save" then

			''comprobacion de si no se quiere como predeterminado, de si existe algun predeterminado
			hay_algun_predet=0
			for item=1 to Nregs-1
				
				if request.form("check" & item)>"" then
					''almacen=limpiaCadena(request.form("alm" & item))
					if cint(predet)=item then
						defecto=1
					else
						defecto=0
					end if
					if defecto=1 then
						hay_algun_predet=1
					end if
				end if
			next
			if hay_algun_predet=0 then
				Nregs=0
				%><script language="javascript" type="text/javascript">
					window.alert("<%=LITNOEXITPREDALMACART%>");
				</script><%
			end if

			lista=""
			for item=1 to Nregs-1
				if request.form("check" & item)>"" then
					almacen=limpiaCadena(request.form("alm" & item))
					stock=limpiaCadena(request.form("stock" & item))
					if not isnumeric(stock) then stock=0
					smin=limpiaCadena(request.form("smin" & item))
					if not isnumeric(smin) then smin=0
                    smax=limpiaCadena(request.form("smax" & item))
					if not isnumeric(smax) then smax=0
					reposicion=limpiaCadena(request.form("rep" & item))
					if not isnumeric(reposicion) then reposicion=0
					precibir=limpiaCadena(request.form("pr" & item))
					if not isnumeric(precibir) then precibir=0
					pservir=limpiaCadena(request.form("ps" & item))
					if not isnumeric(pservir) then pservir=0
					pmin=limpiaCadena(request.form("pm" & item))
					if not isnumeric(pmin) then pmin=0

					ubicacion=limpiaCadena(request.form("ubicacion" & item))
					if cint(predet)=item then
						defecto=1
					else
						defecto=0
					end if
					'response.write("PREDET : " & predet & " ITEM : " & item & " DEFECTO : " & defecto & "<BR>")
					if request.form("ta" & item)>"" then
						es_padre=1
						if request.form("co" & item)>"" then
							reftmp=p_referencia & "/" & limpiaCadena(request.form("ta" & item)) & "/" & limpiaCadena(request.form("co" & item))
						else
							reftmp=p_referencia & "/" & limpiaCadena(request.form("ta" & item))
						end if
					else
						if request.form("co" & item)>"" then
							es_padre=1
							reftmp=p_referencia & "/" & limpiaCadena(request.form("co" & item))
						else
							es_padre=0
							reftmp=p_referencia
						end if
					end if
					''ricardo 1/4/2004 comprobacion de que si el articulo es un equipo, no debe superar a los equipos en estado almacen
					es_equipo=d_lookup("ctrl_nserie","articulos","referencia='" & reftmp & "'",session("dsn_cliente"))
					if nz_b(es_equipo)=-1 then
						equipo_stock_continuar=1
						cuanto_stock_equipos=d_count("nserie","equipos","referencia='" & reftmp & "' and estado='ALMACEN' and fbaja is null and almacen='" & almacen & "'",session("dsn_cliente"))
						if cuanto_stock_equipos<>cLng(stock) then
							equipo_stock_continuar=0
							descripcion=d_lookup("descripcion","almacenes","codigo='" & almacen & "'",session("dsn_cliente"))
							lista=lista & descripcion & ","
						end if
					else
						equipo_stock_continuar=1
					end if
					if equipo_stock_continuar=1 then
						if es_padre=0 then
							''ricardo 10-11-2004 averiguamos el stock anterior solamente si es hijo de tallas y colores
							'stock_old=d_lookup("stock","almacenar","articulo='" & reftmp & "' and almacen='" & almacen & "'",session("dsn_cliente"))
							
							'mmg>>obtenemos los valores para la posterior auditoria
							strselect="select stock,stock_minimo,stock_maximo,reposicion,p_min from almacenar with (nolock) where articulo='" & reftmp & "' and almacen='" & almacen & "'"
							rstMM.open strselect,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
							
							'valores antiguos
							stock_old=rstMM("stock")
							stock_minimo_old=reemplazar(rstMM("stock_minimo"),",",".")
                            stock_maximo_old=reemplazar(rstMM("stock_maximo"),",",".")
							reposicion_old=reemplazar(rstMM("reposicion"),",",".")
							p_min_old=reemplazar(rstMM("p_min"),",",".")
							
							rstMM.Close()
							
							'valores nuevos
							smin=reemplazar(smin,",",".")
                            smax=reemplazar(smax,",",".")
							reposicion=reemplazar(reposicion,",",".")
							pmin=reemplazar(pmin,",",".")
							
							
							''ricardo 11-12-2007 solamente se podran cambiar el pendiente servir,recibir y minimo
							strselect="update almacenar set "
							''strselect=strselect & " stock=" & reemplazar(stock,",",".")
							''strselect=strselect & " , "
							strselect=strselect & " stock_minimo=" & miround(smin,DEC_CANT)
                            strselect=strselect & " ,stock_maximo=" & miround(smax,DEC_CANT)
							strselect=strselect & " ,reposicion=" & miround(reposicion,DEC_CANT)
							''strselect=strselect & " ,p_recibir=" & reemplazar(precibir,",",".")
							''strselect=strselect & " ,p_servir=" & reemplazar(pservir,",",".")
							strselect=strselect & " ,p_min=" & miround(pmin,DEC_CANT)
''response.Write("el predet es-" & predet & "-" & defecto & "-<br/>")
							strselect=strselect & " ,predet=" & defecto
							if ubicacion & "">"" then
								strselect=strselect & ",ubicacion='" & ubicacion & "'"
							else
								strselect=strselect & ",ubicacion=NULL "
							end if
							strselect=strselect & " where articulo='" & reftmp & "' and almacen='" & almacen & "'"
''response.Write(" el strselect es- " & strselect & "-<br/>")
''response.end
							rstDet.open strselect,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
						else
						    'mmg>>obtenemos los valores para la posterior auditoria
							strselect="select stock,stock_minimo,stock_maximo,reposicion,p_min from almacenar with (nolock) where articulo='" & reftmp & "' and almacen='" & almacen & "'"
							rstMM.open strselect,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
							
							'valores antiguos
							stock_old=rstMM("stock")
							stock_minimo_old=reemplazar(rstMM("stock_minimo"),",",".")
                            stock_maximo_old=reemplazar(rstMM("stock_maximo"),",",".")
							reposicion_old=reemplazar(rstMM("reposicion"),",",".")
							p_min_old=reemplazar(rstMM("p_min"),",",".")
							
							rstMM.Close()
							
							'valores nuevos
							smin=reemplazar(smin,",",".")
                            smax=reemplazar(smax,",",".")
							reposicion=reemplazar(reposicion,",",".")
							pmin=reemplazar(pmin,",",".")
						    
						    ''ricardo 11-12-2007 solamente se podran cambiar el pendiente servir,recibir y minimo
							strselect="update almacenar set "
							''strselect=strselect & " stock=" & reemplazar(stock,",",".")
							''strselect=strselect & " , "
							strselect=strselect & " stock_minimo=" & miround(smin,DEC_CANT)
                            strselect=strselect & " ,stock_maximo=" & miround(smax,DEC_CANT)
							strselect=strselect & " ,reposicion=" & miround(reposicion,DEC_CANT)
							''strselect=strselect & " ,p_recibir=" & reemplazar(precibir,",",".")
							''strselect=strselect & " ,p_servir=" & reemplazar(pservir,",",".")
							strselect=strselect & " ,p_min=" & miround(pmin,DEC_CANT)
							if ubicacion & "">"" then
								strselect=strselect & ",ubicacion='" & ubicacion & "'"
							else
								strselect=strselect & ",ubicacion=NULL "
							end if
							strselect=strselect & " where articulo='" & reftmp & "' and almacen='" & almacen & "'"
							rstDet.open strselect,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
						end if
                        
                        'mmg>>auditamos el cambio de stock
						if stock_minimo_old<>smin or stock_maximo_old<>smax or reposicion_old<>reposicion or p_min_old<>pmin then
						    strselect="insert into auditoria(nempresa,login,usuario,IP,fecha,accion,descripcion) values ('"+session("ncliente")+"','"+session("usuario")+"','"+nombreMM+"','"+Request.ServerVariables("REMOTE_ADDR")+"',getdate(),'MODIFICACION ALMACENAR','"
						    ccc=reftmp+";"+almacen+";"+cstr(stock_minimo_old)+";"+cstr(smin)+";"+cstr(stock_maximo_old)+";"+cstr(smax)+";"+cstr(reposicion_old)+";"+cstr(reposicion)+";"+cstr(p_min_old)+";" + cstr(pmin)
						    strselect=strselect+ccc+"')"
						    auditoriaMM=auditoriaMM+"<br/>"+ccc
						    rstMM.open strselect,session("dsn_cliente"),adOpenKeyset, adLockOptimistic
						end if
                        
                        ActualizaCosteMedio reftmp
						rref_padre=d_lookup("ref_padre","articulos","referencia='" & p_referencia & "'",session("dsn_cliente")) & ""
						if es_padre=1 or rref_padre>"" then
							'Si el padre no existe en el almacén donde se va a insertar el hijo, se da de alta
							if es_padre=1 then
								REFE=p_referencia
							else
								REFE=rref_padre
							end if
							rstDet.open "select * from almacenar where articulo='" & REFE & "' and almacen='" & almacen & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
							if rstDet.eof then
								rstDet.Addnew
								rstDet("articulo")=REFE
								rstDet("almacen")=almacen
								rstDet("stock")=0
								rstDet("stock_minimo")=0
                                rstDet("stock_maximo")=0
								rstDet("reposicion")=0
								rstDet("p_recibir")=0
								rstDet("p_servir")=0
								rstDet("p_min")=0
								rstDet("ubicacion")=NULL
								rstDet("reserva")=0
								rstDet("pteabonocanje")=0
								rstDet("sat")=0
								rstDet.update
							end if
							rstDet.close
							'Se actualizan los stocks del padre
							ActualizaStocks "","PADRE",REFE,almacen,"","",session("dsn_cliente")
							ActualizaCosteMedio REFE
							''ricardo 10-11-2004 si el articulo es un hijo de tallas y colores se llamara al procedimiento
							''de actualizar el stock al padre , dicho por JAR
							ActualizaStockPadreAlm stock_old,stock,rref_padre,almacen
						end if
					end if
				end if
			next
			''ricardo 1/4/2004
			if lista & "">"" then
				lista=mid(lista,1,len(lista)-1)
				%><script language="javascript" type="text/javascript">
					window.alert("<%=(LitStockAlmSupStockEquip & "\n" & lista)%>");
				</script><%
			end if
			mode="select"
		end if
		if mode="delete" then

			mode="select"
	DropTable session("usuario"), session("dsn_cliente")
strcrear="create table [" & session("usuario") & "](almacen varchar(10),se_puede_borrar int)"
rstDet.open strcrear,session("dsn_cliente")
if rstDet.state<>0 then rstDet.close
			for item=1 to Nregs-1
				if request.form("check" & item)>"" then
					almacen=limpiaCadena(request.form("check" & item))
seleccion="insert into [" & session("usuario") & "](almacen,se_puede_borrar) values ('" & almacen & "',1)"
rstDet.open seleccion,session("dsn_cliente")
if rstDet.state<>0 then rstDet.close
				end if
			next

''response.write("la referencia es-" & p_referencia & "-<br/>")
''response.end

''ricardo 17-3-2006 se cambia el modo de borrado para pasarlo a restricciones
''					if request.form("ta" & item)>"" then
''						es_padre=1
''						if request.form("co" & item)>"" then
''							reftmp=p_referencia & "/" & limpiaCadena(request.form("ta" & item)) & "/" & limpiaCadena(request.form("co" & item))
''						else
''							reftmp=p_referencia & "/" & limpiaCadena(request.form("ta" & item))
''						end if
''					else
''						if request.form("co" & item)>"" then
''							es_padre=1
''							reftmp=p_referencia & "/" & limpiaCadena(request.form("co" & item))
''						else
''							es_padre=0
''							reftmp=p_referencia
''						end if
''					end if
''
''
''					seleccion="select referencia from detalles_ped_cli where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_ped_pro where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_alb_cli where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_alb_pro where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_fac_cli where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_fac_pro where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_tickets where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''
''					seleccion=seleccion & " UNION select referencia from detalles_pre_cli where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_orden where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select referencia from detalles_orden_fab where (referencia='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select ref from detalles_movimientos where (ref='" & reftmp & "' and almorigen='" & almacen & "')"
''					seleccion=seleccion & " UNION select art_padre from escandallo where (art_padre='" & reftmp & "' and almacen='" & almacen & "')"
''					seleccion=seleccion & " UNION select art_hijo from escandallo where (art_hijo='" & reftmp & "' and almacen='" & almacen & "')"
''			''ricardo 1/4/2004 comprobacion de que si el articulo es un equipo, y tiene estado almacen no se pueda borrar
''					seleccion=seleccion & " UNION select referencia from equipos where estado='ALMACEN' and referencia='" & reftmp & "' and almacen='" & almacen & "' and fbaja is null"
''		'response.write seleccion
''a este modo

''ricardo 17-3-2006 a partir de esta fecha se borrara con procedimiento
	no_se_puede_borrar=0
	set conn = Server.CreateObject("ADODB.Connection")
	set command =  Server.CreateObject("ADODB.Command")
		
	conn.open session("dsn_cliente")
	command.ActiveConnection =conn
	command.CommandTimeout = 0
	command.CommandText="BorrarAlmacenar"
	command.CommandType = adCmdStoredProc 'Procedimiento Almacenado
	command.Parameters.Append command.CreateParameter("@almacen",adVarChar,adParamInput,10,"tabla")
	command.Parameters.Append command.CreateParameter("@articulo",adVarChar,adParamInput,30,p_referencia)
	command.Parameters.Append command.CreateParameter("@usuario",adVarChar,adParamInput,50,session("usuario"))
	command.Parameters.Append command.CreateParameter("@result",adInteger,adParamOutput)
	on error resume next
	command.Execute,,adExecuteNoRecords
	if err.number<>0 then
		no_se_puede_borrar=1
	else
		resultado=command.Parameters("@result").Value
		if resultado=1 then
			no_se_puede_borrar=1
		end if
	end if
	on error goto 0
	conn.close
	set command=nothing
	set conn=nothing

					''if not rstDet.eof then
					if no_se_puede_borrar=1 then

						''NombreALmacen=d_lookup("descripcion","almacenes","codigo='" & almacen & "'",session("dsn_cliente"))%>
						<script>
		    					//window.alert ("<%=trimCodEmpresa(reftmp)%> - <%=NombreAlmacen%>.\n<%=LitMsgRefenVentas2%>");
							window.alert ("<%=LITMSGREFENVENTAS2ALM1ALMACART%> <%=trimCodEmpresa(p_referencia)%> <%=LITMSGREFENVENTAS2ALM2ALMACART%>");
		   				</script><%
					else

''ricardo 17-3-2006 como solamente se puede borrar si el stock esta a cero, no habria que actualizar el stock del padre , ni el coste medio
''ya que ya estara bien.
''						''ricardo 10-11-2004 averiguamos el stock anterior solamente si es hijo de tallas y colores
''						stock_old=d_lookup("stock","almacenar","articulo='" & p_referencia & "' and almacen='" & almacen & "'",session("dsn_cliente"))
''''''''''''''
''
''         					rstDet.open "delete from escandallo where art_padre ='" & p_referencia & "' and almacen ='" & almacen &  "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
''
''						rstDet.open "delete from almacenar where articulo='" & reftmp & "' and almacen='" & almacen & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
''ActualizaCosteMedio reftmp
''						rstDet.open "select predet from almacenar where articulo='" & reftmp & "'",session("dsn_cliente"),adOpenKeyset, adLockOptimistic
''						HayPre=false
''						if not rstDet.eof then
''							while not rstDet.eof
''								if nz_b(rstDet("predet"))<>"0" then HayPre=true
''								rstDet.movenext
''							wend
''							if not HayPre then
''								rstDet.movefirst
''								rstDet("predet")=1
''								rstDet.update
''							end if
''							rstDet.close
''						else
''							rstDet.close
''						end if
''
''ricardo 17-3-2006 como solamente se puede borrar si el stock esta a cero, no habria que actualizar el stock del padre , ni el coste medio
''ya que ya estara bien.
''						rref_padre=d_lookup("ref_padre","articulos","referencia='" & p_referencia & "'",session("dsn_cliente")) & ""
''						if es_padre=1 or rref_padre>"" then
''							if es_padre=1 then
''								REFE=p_referencia
''							else
''								REFE=rref_padre
''							end if
''							'Se actualizan los stocks del padre
''							ActualizaStocks "","PADRE",REFE,almacen,"","",session("dsn_cliente")
''							'Si el único registro que queda en el almacén es el padre, se elimina porque el padre no se puede vender directamente
''
''							''ricardo 16-11-2004 se cambia el select por otro, ya que cuando se borraba un hijo se borraba el padre, aunque existiera otro hijo con el mismo almacen
''							''strselect="select predet from almacenar where articulo='" & REFE & "' and almacen='" & almacen & "'"
''							strselect="select * from almacenar where articulo in (select referencia from articulos where ref_padre='" & REFE & "') and almacen='" & almacen & "'"
''							rstDet.cursorlocation=3
''							rstDet.open strselect,session("dsn_cliente")
''							if rstdet.recordcount=0 then
''								rstdet.close
''								strselect="delete from almacenar where articulo='" & REFE & "' and almacen='" & almacen & "'"
''								rstDet.cursorlocation=3
''								rstDet.open strselect,session("dsn_cliente")
''								if rstDet.state<>0 then rstDet.close
''							else
''								rstdet.close
''							end if
''
''							ActualizaCosteMedio REFE
''							''ricardo 10-11-2004 si el articulo es un hijo de tallas y colores se llamara al procedimiento
''							''de actualizar el stock al padre , dicho por JAR
''							ActualizaStockPadreAlm stock_old,0,REFE,almacen
''						end if
					end if
''				end if
''			next
		end if
		if mode="addalmacen" then

			if talla>"" then
				predet="false"
				es_padre=1
				if color>"" then
					strwhere="where articulo='" & p_referencia & "/" & trimCodEmpresa(talla) & "/" & trimCodEmpresa(color) & "' and almacen='" & almacen & "'"
					reftmp=p_referencia & "/" & trimCodEmpresa(talla) & "/" & trimCodEmpresa(color)
				else
					strwhere="where articulo='" & p_referencia & "/" & trimCodEmpresa(talla) & "' and almacen='" & almacen & "'"
					reftmp=p_referencia & "/" & trimCodEmpresa(talla)
				end if
			else
				if color>"" then
					predet="false"
					es_padre=1
					strwhere="where articulo='" & p_referencia & "/" & trimCodEmpresa(color) & "' and almacen='" & almacen & "'"
					reftmp=p_referencia & "/" & trimCodEmpresa(color)
				else
					es_padre=0
					strwhere="where articulo='" & p_referencia & "' and almacen='" & almacen & "'"
					strwhere2="where articulo='" & p_referencia & "'"
					reftmp=p_referencia
				end if
			end if

			''ricardo 1/4/2004 comprobacion de que si el articulo es un equipo, no debe superar a los equipos en estado almacen
			lista=""
			equipo_stock_continuar=1
			es_equipo=d_lookup("ctrl_nserie","articulos","referencia='" & reftmp & "'",session("dsn_cliente"))
			if nz_b(es_equipo)=-1 then
				equipo_stock_continuar=1
				cuanto_stock_equipos=d_count("nserie","equipos","referencia='" & reftmp & "' and estado='ALMACEN' and fbaja is null and almacen='" & almacen & "'",session("dsn_cliente"))
				if cuanto_stock_equipos<>cLng(stock) then
					equipo_stock_continuar=0
					descripcion=d_lookup("descripcion","almacenes","codigo='" & almacen & "'",session("dsn_cliente"))
					lista=lista & descripcion & ","
				end if
			else
				equipo_stock_continuar=1
			end if

			mode="select"
			if equipo_stock_continuar=1 then
				rstDet.open "select * from almacenar " & strwhere2,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if rstDet.EOF then
					predet="true"
				end if
				rstDet.close
				rstDet.open "select * from almacenar " & strwhere,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
				if rstDet.EOF then
					if predet="true" then
						rstAux.open "update almacenar set predet=0 where articulo='" & p_referencia & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					end if
					'Si el hijo no tiene almacen predeterminado se le asigna este nuevo
					if es_padre=1 then
						ErrorTC=ReferenciaConTC(reftmp,p_referencia,almacen,"",session("dsn_cliente"))
						rstAux.open "select articulo,almacen from almacenar where articulo='" & reftmp & "' and predet<>0",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						if rstAux.EOF then predet="true"
						rstAux.close
					end if
					rstAux.open "select * from almacenar where articulo='" & reftmp & "' and almacen='"&almacen&"'",session("dsn_cliente")
					if rstAux.eof then
						rstDet.Addnew
						rstDet("articulo")=reftmp
						rstDet("almacen")=almacen
					else
						rstDet.close
						rstDet.open "select * from almacenar " & strwhere,session("dsn_cliente"),adOpenKeyset,adLockOptimistic
					end if
					rstAux.Close
					if ubicacion>"" then
						rstDet("ubicacion")=ubicacion
					end if
					''rstDet("stock")=stock
					rstDet("stock_minimo")=miround(smin,DEC_CANT)
                    rstDet("stock_maximo")=miround(smax,DEC_CANT)
					rstDet("reposicion")=miround(reposicion,DEC_CANT)
					''rstDet("p_recibir")=precibir
					''rstDet("p_servir")=pservir
					rstDet("p_min")=miround(pmin,DEC_CANT)
					rstDet("predet")=nz_b(predet)
					rstDet.update
					rstDet.close
                    ActualizaCosteMedio reftmp
					'response.write("ES PADRE : " & es_padre & " " & p_referencia & "<BR>")
					rref_padre=d_lookup("ref_padre","articulos","referencia='" & p_referencia & "'",session("dsn_cliente")) & ""
					if es_padre=1 or rref_padre>"" then
						'Si el padre no existe en el almacén donde se va a insertar el hijo, se da de alta
						if es_padre=1 then
							REFE=p_referencia
						else
							REFE=rref_padre
						end if
						rstDet.open "select * from almacenar where articulo='" & REFE & "' and almacen='" & almacen & "'",session("dsn_cliente"),adOpenKeyset,adLockOptimistic
						if rstDet.eof then
							rstDet.Addnew
							rstDet("articulo")=REFE
							rstDet("almacen")=almacen
							rstDet("stock")=0
							rstDet("stock_minimo")=0
                            rstDet("stock_maximo")=0
							rstDet("reposicion")=0
							rstDet("p_recibir")=0
							rstDet("p_servir")=0
							rstDet("p_min")=0
							rstDet.update
						end if
						rstDet.close
						'Se actualizan los stocks del padre
						ActualizaStocks "","PADRE",REFE,almacen,"","",session("dsn_cliente")
						ActualizaCosteMedio REFE
						''ricardo 10-11-2004 si el articulo es un hijo de tallas y colores se llamara al procedimiento
						''de actualizar el stock al padre , dicho por JAR
						ActualizaStockPadreAlm 0,stock,rref_padre,almacen
					end if

					' Actualizamos el coste medio en caso de que el artículo no aparezca en ningún documento de compras'
					''connAlmacenar.open session("dsn_cliente")
					''strselect="EXEC costeMedioInicial @referencia='" & p_referencia & "'"
					''set rst = connAlmacenar.execute(strselect)
					''connAlmacenar.close
				else
					rstDet.close
					%><script language="javascript" type="text/javascript">
						  alert("<%=LitMsgArtAlmacenExiste%>");
					</script><%
				end if
			else
				if lista & "">"" then
						lista=mid(lista,1,len(lista)-1)
					%><script language="javascript" type="text/javascript">
						window.alert("<%=(LitStockAlmSupStockEquip & "\n" & lista)%>");
					</script><%
				end if
			end if
		end if

		if mode="select" then '------------------------------------------------------------------------------
			if p_referencia&""="" then p_referencia=limpiaCadena(request.querystring("referencia"))
			if es_padre&""="" or es_padre=0 then es_padre=nz_b(limpiaCadena(request.querystring("espadre")))

            es_padre2=cstr(limpiaCadena(request.querystring("espadre"))&"")
            if es_padre2 & ""="" then
                es_padre2=cstr(es_padre&"")
            end if
            ''response.write("el es_padre es-" & es_padre & "-" & request.querystring("espadre") & "-" & es_padre2 & "-<br>")
			%><table class="width100 iframe-tab-nospace" style="" cellspacing="1" cellpadding="1"><%
			'DGB: 11/04/2016: no mostrar Almacen de tipo Tanque
            if es_padre<>0 then
				'rstDet.open "select AL.*,A.descripcion As Nalmacen,talla,color from almacenar AL,almacenes A,articulos AR where referencia=articulo and articulo in (select referencia from articulos where ref_padre='" & p_referencia & "') and AL.almacen=A.codigo order by A.descripcion,talla,color", session("dsn_cliente"),adOpenKeyset, adLockOptimistic
				'Temporalmente se muestra siempre el almacen del articulo padre. La linea anterior comentada es la que mustra padre e hijos
				rstDet.open "select AL.*,A.descripcion As Nalmacen from almacenar AL,almacenes A where articulo='" & p_referencia & "' and AL.almacen=A.codigo and a.tienda is null order by A.descripcion", session("dsn_cliente"),adOpenKeyset, adLockOptimistic
			else
				rstDet.open "select AL.*,A.descripcion As Nalmacen from almacenar AL,almacenes A where articulo='" & p_referencia & "' and AL.almacen=A.codigo and a.tienda is null order by A.descripcion", session("dsn_cliente"),adOpenKeyset, adLockOptimistic
			end if

			no_ir_al_principio=0
			rst.cursorlocation=3
			rst.open "select codigo,descripcion from ubicaciones where codigo like '" & session("ncliente") & "%'",session("dsn_cliente")
			if rst.eof then
				no_ir_al_principio=1
			end if
			linea=1
			while not rstDet.eof
				if no_ir_al_principio=0 then
					rst.movefirst
				end if
				'if es_padre<>0 then LINEA COMENTADA PARA NO MOSTRAR LOS ALMACENES DE HIJOS
				if false then%>
					<tr bgcolor="<%=color_blanc%>" onmouseover="parent.document.getElementById('RefHijo').innerText='REF : <%=rstDet("articulo")%>'" onmouseout="parent.document.getElementById('RefHijo').innerText=''">
						<td class='CELDA7' width="180"><%=rstDet("Nalmacen")%></td>
						<td class='CELDA7' width="90" style='width:90px'>
							<%'d_lookup("descripcion","ubicaciones","codigo='" & rstDet("ubicacion") & "'",session("dsn_cliente"))%>
							<select name='ubicacion<%=linea%>' class='CELDAL7' style='width:90px' <%=deshabilitar%>>
								<%
								seleccionado=""
								he_puesto=0
								while not rst.eof
									cad_ubi=trimCodEmpresa(rst("codigo")) & " - " & rst("descripcion")
									if rst("codigo")=rstDet("ubicacion") then
										seleccionado="selected"
										he_puesto=1
									end if
									%><option value='<%=rst("codigo")%>' <%=seleccionado%>><%=cad_ubi%></option><%
									rst.movenext
									seleccionado=""
								wend
								if he_puesto=0 then
									%><option value='' selected="selected"></option><%
								else
									%><option value=''></option><%
								end if
								%>
							</select>
						</td>
						<td class='CELDA7' width="50"><input type="hidden" name="ta<%=linea%>" value="<%=rstDet("talla")%>" <%=deshabilitar%>/><%=d_lookup("descripcion","tallas","codigo='" & rstDet("talla") & "'",session("dsn_cliente"))%></td>
						<td class='CELDA7' width="50"><input type="hidden" name="co<%=linea%>" value="<%=rstDet("color")%>" <%=deshabilitar%>/><%=d_lookup("descripcion","colores","codigo='" & rstDet("color") & "'",session("dsn_cliente"))%></td>
						<td width="49" align="center"><input class="INPUTALM" type="text" name="stock<%=linea%>" value="<%=rstDet("stock")%>" <%=deshabilitar%> onchange="Valido(this,'<%=linea%>','stock');"/></td>
						<td width="49" align="center"><input class="INPUTALM" type="text" name="smin<%=linea%>" value="<%=rstDet("stock_minimo")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','stock mínimo');"/></td>
                        <td width="49" align="center"><input class="INPUTALM" type="text" name="smax<%=linea%>" value="<%=rstDet("stock_maximo")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','stock máximo');"/></td>
						<td width="49" align="center"><input class="INPUTALM" type="text" name="rep<%=linea%>" value="<%=rstDet("reposicion")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','reposición');"/></td>
						<td width="49" align="center"><input class="INPUTALM" type="text" name="pr<%=linea%>" value="<%=rstDet("p_recibir")%>" <%=deshabilitar%> onchange="Valido(this,'<%=linea%>','pendiente de recibir');"/></td>
						<td width="49" align="center"><input class="INPUTALM" type="text" name="ps<%=linea%>" value="<%=rstDet("p_servir")%>" <%=deshabilitar%> onchange="Valido(this,'<%=linea%>','pendiente de sevir');"/></td>
						<td width="49" align="center"><input class="INPUTALM" type="text" name="pm<%=linea%>" value="<%=rstDet("p_min")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','pedido mínimo');"/></td>
						<td width="25"><input type="hidden" name="alm<%=linea%>" value="<%=rstDet("almacen")%>"/><input class="CHECK" type="Checkbox" name="check<%=linea%>" value="<%=rstDet("almacen")%>" checked="checked"/></td><%
					CloseFila
				else
					DrawFila color_blanc%>
						<td class='CELDA7 width10'><%=rstDet("Nalmacen")%></td>
                        <%
                        if cstr(es_padre2)<>"0" then
                            ancho_ubi="205px"
                        else
                            ancho_ubi="160px"
                        end if
                        %>
						<td class='CELDA7 width10'>
							<select name='ubicacion<%=linea%>' class='width100'>
							<!-- fin cag -->
								<%
								seleccionado=""
								he_puesto=0
								while not rst.eof
									cad_ubi=trimCodEmpresa(rst("codigo")) & " - " & rst("descripcion")
									if rst("codigo")=rstDet("ubicacion") then
										seleccionado="selected"
										he_puesto=1
									end if
									%><option value="<%=rst("codigo")%>" <%=seleccionado%>><%=cad_ubi%></option><%
									rst.movenext
									seleccionado=""
								wend
								if he_puesto=0 then
									%><option value="" selected="selected"></option><%
								else
									%><option value=""></option><%
								end if
								%>
							</select>
						</td>
						<td class="CELDAC7 width15"></td>
						<td class="CELDAC7 width5" align="center"><input class="width100" style="text-align:center;" type="text" name="stock<%=linea%>" value="<%=rstDet("stock")%>" <%=deshabilitar%> onchange="Valido(this,'<%=linea%>','stock');"/></td>
						<td class="CELDAC7 width5" align="center"><input class="width100" style="text-align:center;" type="text" name="smin<%=linea%>" value="<%=rstDet("stock_minimo")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','stock mínimo');"/></td>
                        <td class="CELDAC7 width5" align="center"><input class="width100" style="text-align:center;" type="text" name="smax<%=linea%>" value="<%=rstDet("stock_maximo")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','stock máximo');"/></td>
						<td class="CELDAC7 width5" align="center"><input class="width100" style="text-align:center;" type="text" name="rep<%=linea%>" value="<%=rstDet("reposicion")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','reposición');"/></td>
						<td class="CELDAC7 width5" align="center"><input class="width100" style="text-align:center;" type="text" name="pr<%=linea%>" value="<%=rstDet("p_recibir")%>" <%=deshabilitar%> onchange="Valido(this,'<%=linea%>','pendiente de recibir');"/></td>
                        <%
                        if cstr(es_padre2)<>"0" then
                            ancho_ps="45px"
                        else
                            ancho_ps="50px"
                        end if
                        %>
						<td class="CELDAC7 width5" align="center"><input class="width100" style="text-align:center;" type="text" name="ps<%=linea%>" value="<%=rstDet("p_servir")%>" <%=deshabilitar%> onchange="Valido(this,'<%=linea%>','pendiente de sevir');"/></td>
                        <%
                        if cstr(es_padre2)<>"0" then
                            ancho_pm="50px"
                        else
                            ancho_pm="55px"
                        end if
                        %>
						<td class="CELDAC7 width5" align="center"><input class="width100" style="text-align:center;" type="text" name="pm<%=linea%>" value="<%=rstDet("p_min")%>" <%=deshabilitarpsp%> onchange="Valido(this,'<%=linea%>','pedido mínimo');"/></td>
						<%
                        if cstr(es_padre2)<>"0" then
                            ancho_radio="15px"
                        else
                            ancho_radio="50px"
                        end if
                        if nz_b(rstDet("predet"))<>0 then %>
							<!-- cag -->
							<!-- <td width="49" align="center"><input class="CHECK" type="radio" name="predet" value="<%=linea%>" checked <%=deshabilitar%>></td> -->
							<td class="CELDAC7 width5" style="text-align:center;"><input class="" type="radio" name="predet" value="<%=linea%>" checked="checked" /></td>
							<!-- fin cag -->
						<%else%>
						    <!-- cag -->
							<!-- <td width="49" align="center"><input class="CHECK" type="radio" name="predet" value="<%=linea%>" <%=deshabilitar%>></td> -->
  						    <td class="CELDAC7 width5" style="text-align:center;"><input class="" type="radio" name="predet" value="<%=linea%>" /></td>
							<!-- fin cag -->
						<%end if%>
						<!-- cag -->
						<td class="CELDAC7 width5" style="text-align:center;">
                            <input type="hidden" name="alm<%=linea%>" value="<%=rstDet("almacen")%>"/>
                            <input class="" type="Checkbox" name="check<%=linea%>" value="<%=rstDet("almacen")%>" checked="checked" />
                        </td><%
					CloseFila
				end if
				linea=linea+1
				rstDet.movenext
			wend
			rstDet.close
			rst.close
			%></table>
			<input type="hidden" name="hNregs" value="<%=linea%>"/><%
		end if
	end if
	
	'comprobamos si hay que tratar alarma
	if auditoriaMM&"" <> "" then
	    Set rstTratarAlarmas = Server.CreateObject("ADODB.Connection")
		rstTratarAlarmas.open dsnilion
		      
		mensaje="artículo;almacén;stock mínimo anterior;stock mínimo;stock máximo anterior;stock máximo;stock reposición anterior;stock reposición;stock pedido anterior;stock pedido" 
		mensaje=mensaje+"<br/>"+auditoriaMM               
	    strselect = "exec tratar_alarmas @tipo_alarma='017', @empresa='" & session("ncliente") & "', @usuario='" & session("usuario") & "', @des_tecnica='" & mensaje & "', @CarpetaProduccion='" & CarpetaProduccion & "'"
	    'response.Write(strselect)
		rstTratarAlarmas.execute strselect
	end if
%></form><%
connRound.close
set connRound = Nothing
set rst = Nothing
set rstDet = Nothing
set rstAux = Nothing
set rstMM = Nothing
set connAlmacenar = Nothing
Set rstTratarAlarmas = Nothing

end if%>
</body>
</html>
