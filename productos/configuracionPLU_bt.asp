<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<TITLE><%=LitTitulo%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../varios_bt.inc" -->
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
</HEAD>
<script Language="JavaScript" src="../jfunciones.js"></script>
<body leftmargin="<%=LitLeftPosBT%>" topmargin="<%=LitTopPosBT%>" bgcolor="<%=color_fondo_bt%>">
<%mode=Request.QueryString("mode")
ImprimirPie_bt%>
</BODY>
</HTML>
