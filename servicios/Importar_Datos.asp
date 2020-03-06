<%@ Language=VBScript %>
<%Server.ScriptTimeout = 2400%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<% 
    dim  enc
    set enc = Server.CreateObject("Owasp_Esapi.Encoder")
%>
<HTML LANG="<%=session("lenguaje")%>">
<HEAD>
<TITLE><%=LitTitulo%></TITLE>
<META HTTP-EQUIV="Content-Type" Content="text/html"; charset="<%=session("caracteres")%>">
<!--#include file="../ilion.inc" -->
<!--#include file="../constantes.inc" -->
<!--#include file="../mensajes.inc" -->
<!--#include file="Importar_Datos.inc" -->
<!--#include file="../cache.inc" -->
<!--#include file="../tablas.inc" -->
<!--#include file="../adovbs.inc" -->
<!--#include file="../calculos.inc" -->
<!--#include file="../varios.inc" -->
<!--#include file="../ico.inc" -->
<!--#include file="../modulos.inc" -->
<!--#include file="../common/GetUploadsPathByLocalIP.inc" -->
<!--#include file="../styles/listTable.css.inc" -->
<!--#include file="../styles/formularios.css.inc" -->
<script language="JavaScript" src="../jfunciones.js"></script>
<script language="JavaScript" type="text/javascript">
    function Repinta() {
        if (document.importar_datos.proceso[0].checked) var proceso = "C";
        else {
            if (document.importar_datos.proceso[1].checked) var proceso = "P";
            else if (document.importar_datos.proceso[2].checked) var proceso = "A";
            else var proceso = "F";
		}

		var idfile02 = document.getElementById("idfile02");
		var idfile03 = document.getElementById("idfile03");
		var idfile04 = document.getElementById("idfile04");

		var idlitfile01 = document.getElementById("idlitfile01");
		var idlitfile02 = document.getElementById("idlitfile02");
		var idlitfile03 = document.getElementById("idlitfile03");
		var idlitfile04 = document.getElementById("idlitfile04");

		var idlitfile01p = document.getElementById("idlitfile01p");
		var idlitfile02p = document.getElementById("idlitfile02p");
		var idlitfile03p = document.getElementById("idlitfile03p");
		var idlitfile04p = document.getElementById("idlitfile04p");


        if (document.importar_datos.proceso[3] != null && document.importar_datos.proceso[3].checked == false && document.importar_datos.formato[0].disabled == true) {
            document.importar_datos.formato[0].disabled = false;

            idlitfile02.style.display = '';
            idlitfile03.style.display = '';
            idlitfile04.style.display = '';
            idlitfile02p.style.display = '';
            idlitfile03p.style.display = '';
            idlitfile04p.style.display = '';

            document.importar_datos.fichero02.style.display = 'none';
            document.importar_datos.examinar2.style.display = '';
            document.importar_datos.input_file2.style.display = '';

            document.importar_datos.fichero03.style.display = 'none';
            document.importar_datos.examinar3.style.display = '';
            document.importar_datos.input_file3.style.display = '';

            document.importar_datos.fichero04.style.display = 'none';
            document.importar_datos.examinar4.style.display = '';
            document.importar_datos.input_file4.style.display = '';
        }
        switch (proceso) {
			case "C":
                if (document.importar_datos.formato[0].checked) { //Clientes dbf
                    idlitfile01.firstChild.innerHTML = '<%=LitClientesDBF%> : ';
                    idlitfile02.firstChild.innerHTML = '<%=LitfpagoDBF%> : ';

                    idlitfile03.firstChild.style.visibility = 'visible';
                    document.importar_datos.examinar3.style.display = "";
                    document.importar_datos.input_file3.style.display = "";
                    idlitfile03.firstChild.innerHTML = '<%=LitagentesDBF%> : ';

                    idlitfile04.firstChild.style.visibility = 'visible';
                    idlitfile04.style.display = '';

                    idfile03.firstChild.style.visibility = 'visible';
                    idfile04.firstChild.style.visibility = 'visible';

                    document.getElementById("id_fichero04").style.display = 'none';
                    document.importar_datos.examinar4.style.display = '';
                    document.importar_datos.input_file4.style.display = '';

                    idlitfile04p.style.display = '';

                    idlitfile01p.style.display = 'none';
                    idlitfile02p.style.display = 'none';
                    idlitfile03p.style.display = 'none';
                    idlitfile04p.style.display = 'none';

                    idlitfile01p.firstChild.style.visibility = 'hidden';
                    idlitfile02p.firstChild.style.visibility = 'hidden';
                    idlitfile03p.firstChild.style.visibility = 'hidden';
					idlitfile04p.firstChild.style.visibility = 'hidden';

				} else { //Clientes csv
                    idlitfile01.firstChild.innerHTML = '<%=LitClientesCSV%> : ';
                    idlitfile02.firstChild.innerHTML = 'fpago.csv : ';
                    idlitfile03.firstChild.style.visibility = 'visible';
                    document.importar_datos.examinar3.style.display = "";
                    document.importar_datos.input_file3.style.display = "";

                    idlitfile03.firstChild.innerHTML = 'comercial.csv : ';
                    idlitfile04.firstChild.style.visibility = 'visible';
                    idfile04.firstChild.style.visibility = '';

                    document.getElementById("id_fichero04").style.display = 'none';
                    document.importar_datos.examinar4.style.display = 'none';
                    document.importar_datos.input_file4.style.display = 'none';

                    idlitfile01p.style.display = '';
                    idlitfile01p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/clientes.xls'><%=LitPlantilla%></a>";
                    idlitfile02p.style.display = '';
                    idlitfile02p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/fpago.xls'><%=LitPlantilla%></a>";
                    idlitfile03p.style.display = '';
                    idlitfile03p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/comercial.xls'><%=LitPlantilla%></a>";
                    idlitfile04p.style.display = '';
                    idlitfile04p.firstChild.innerHTML = "";

					idlitfile01p.firstChild.style.visibility = 'visible';                   
                    idlitfile02p.firstChild.style.visibility = 'visible';
                    idlitfile03p.firstChild.style.visibility = 'visible';
					idlitfile04p.firstChild.style.visibility = 'visible';

                }
                break;
			case "P":
                if (document.importar_datos.formato[0].checked) { //Proveedores dbf
                    idlitfile01.firstChild.innerHTML = '<%=LitProveedoDBF%> : ';
                    idlitfile02.firstChild.innerHTML = '<%=LitfpagoDBF%> : ';
                    idlitfile03.firstChild.style.visibility = 'hidden';
                    idlitfile03.firstChild.innerHTML = '<%=LitprovincDBF%> : ';
                    idlitfile04.firstChild.style.visibility = 'visible';
                    idlitfile04.style.display = '';

                    idfile03.firstChild.style.visibility = 'hidden';
                    document.importar_datos.examinar3.style.display = "none";
                    document.importar_datos.input_file3.style.display = "none";

                    idfile04.firstChild.style.visibility = 'visible';
                    document.getElementById("id_fichero04").style.display = 'none';
                    document.importar_datos.examinar4.style.display = '';
                    document.importar_datos.input_file4.style.display = '';
                    idlitfile04p.style.display = '';


                    idlitfile01p.style.display = 'none';
                    idlitfile02p.style.display = 'none';
					idlitfile03p.style.display = 'none';
					idlitfile04p.style.display = 'none';

                    idlitfile01p.firstChild.style.visibility = 'hidden';
                    idlitfile02p.firstChild.style.visibility = 'hidden';
					idlitfile03p.firstChild.style.visibility = 'hidden';

                    //idlitfile04p.style.display = 'none';
				} else { //Proveedores csv
                    idlitfile01.firstChild.innerHTML = '<%=LitProveedoCSV%> : ';
                    idlitfile02.firstChild.innerHTML = '<%=LitfpagoCSV%> : ';
                    idlitfile03.firstChild.style.visibility = 'hidden';
                    idlitfile03.firstChild.innerHTML = '';
                    idlitfile04.firstChild.style.visibility = 'hidden';

                    idfile03.firstChild.style.visibility = 'hidden';
                    document.importar_datos.examinar3.style.display = "none";
                    document.importar_datos.input_file3.style.display = "none";

                    idfile04.firstChild.style.visibility = 'hidden';
                    document.getElementById("id_fichero04").style.display = 'none';
                    document.importar_datos.examinar4.style.display = 'none';
                    document.importar_datos.input_file4.style.display = 'none';

                    idlitfile01p.style.display = '';
                    idlitfile01p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/proveedores.xls'><%=LitPlantilla%></a>";
                    idlitfile02p.style.display = '';
                    idlitfile02p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/fpago.xls'><%=LitPlantilla%></a>";
                    idlitfile03p.style.display = '';
                    idlitfile03p.firstChild.innerHTML = ""
                    idlitfile04p.style.display = '';
                    idlitfile04p.firstChild.innerHTML = "";

                    idlitfile01p.firstChild.style.visibility = 'visible';
                    idlitfile02p.firstChild.style.visibility = 'visible';
					idlitfile03p.firstChild.style.visibility = 'visible';
                }
                break;

			case "A":
                if (document.importar_datos.formato[0].checked) { //Articulos dbf

                    idlitfile01.firstChild.innerHTML = '<%=LitarticuloDBF%> : ';
                    idlitfile02.firstChild.innerHTML = '<%=LitfamiliasDBF%> : ';
                    idlitfile03.firstChild.style.visibility = 'visible';
                    document.importar_datos.examinar3.style.display = "";
                    document.importar_datos.input_file3.style.display = "";
                    idlitfile03.firstChild.innerHTML = '<%=LitstocksDBF%> : ';
					idlitfile04.style.display = 'none';

                    idfile03.firstChild.style.visibility = 'visible';
                    idlitfile01p.style.display = 'none';
                    idlitfile02p.style.display = 'none';
                    idlitfile03p.style.display = 'none';
					idlitfile04p.style.display = 'none';

                    idfile04.firstChild.style.visibility = '';
                    document.getElementById("id_fichero04").style.display = 'none';
                    document.importar_datos.examinar4.style.display = 'none';
                    document.importar_datos.input_file4.style.display = 'none';

                    idlitfile01p.firstChild.style.visibility = "hidden";
                    idlitfile02p.firstChild.style.visibility = "hidden";
					idlitfile03p.firstChild.style.visibility = "hidden";

                } else { //Articulos csv
					//console.log("A2");
                    idlitfile01.firstChild.innerHTML = '<%=LitarticuloCSV%> : ';
                    idlitfile02.firstChild.innerHTML = '<%=LitfamiliasCSV%> : ';
                    idlitfile03.firstChild.innerHTML = '<%=LitStocksCSV%> : ';
                    idlitfile03.firstChild.style.visibility = 'visible';


                    idlitfile01p.style.display = '';
                    idlitfile01p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/articulos.xls'><%=LitPlantilla%></a>";
                    idlitfile02p.style.display = '';
                    idlitfile02p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/familias.xls'><%=LitPlantilla%></a>";
                    idlitfile03p.style.display = '';
                    idlitfile03p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/stocks.xls'><%=LitPlantilla%></a>";

                    idlitfile01p.firstChild.style.visibility = "visible";
                    idlitfile02p.firstChild.style.visibility = "visible";
                    idlitfile03p.firstChild.style.visibility = "visible";


                    idfile04.firstChild.style.visibility = '';
                    document.getElementById("id_fichero04").style.display = 'none';
                    document.importar_datos.examinar4.style.display = 'none';
                    document.importar_datos.input_file4.style.display = 'none';
                    idlitfile04p.style.display = 'none';
					idlitfile04.style.display = 'none';

                }
                break;

			case "F":
				//console.log("F1");
                //if (document.importar_datos.formato[0].checked) { // FIDELIZACIÓN CSV
                document.importar_datos.formato[1].checked = true;
                document.importar_datos.formato[0].disabled = true;
                idlitfile01.firstChild.innerHTML = '<%=LitTarjetasCSV%>: ';

                idfile02.firstChild.style.visibility = '';
                document.importar_datos.fichero02.style.display = 'none';
                document.importar_datos.examinar2.style.display = 'none';
                document.importar_datos.input_file2.style.display = 'none';

                idfile03.firstChild.style.visibility = '';
                document.importar_datos.fichero03.style.display = 'none';
                document.importar_datos.examinar3.style.display = 'none';
                document.importar_datos.input_file3.style.display = 'none';

                idfile04.firstChild.style.visibility = '';
                document.getElementById("id_fichero04").style.display = 'none';
                document.importar_datos.examinar4.style.display = 'none';
                document.importar_datos.input_file4.style.display = 'none';

                idlitfile01p.style.display = '';
                idlitfile01p.firstChild.innerHTML = "<a class='CELDAREF7' href='../../Documentos/plantillas_importacion_datos/Tarjetas.xls'><%=LitPlantilla%></a>";
                idlitfile01p.firstChild.style.visibility = "visible";
                idlitfile02p.style.display = 'none';
                idlitfile02.style.display = 'none';
                idlitfile03p.style.display = 'none';
                idlitfile03.style.display = 'none';
                idlitfile04p.style.display = 'none';
				idlitfile04.style.display = 'none';

                break;
        }

    }

    function onSelectFileClick(numFile) {

        if (numFile == "1") {
            document.importar_datos.file1.value=document.importar_datos.fichero01.value;
        }
        else if (numFile == "2") {
            document.importar_datos.file2.value=document.importar_datos.fichero02.value;
        }
        else if (numFile == "3") {
            document.importar_datos.file3.value=document.importar_datos.fichero03.value;
        }
        else if (numFile == "4") {
            document.importar_datos.file4.value=document.importar_datos.fichero04.value;
        }
        else if (numFile == "5") {
            document.importar_datos.file5.value=document.importar_datos.fichero05.value;
        }
    }

    function onSelectFileChange(numFile) {

        if (numFile == "1") {
            document.importar_datos.file1.value=document.importar_datos.fichero01.value;
        }
        else if (numFile == "2") {
            document.importar_datos.file2.value=document.importar_datos.fichero02.value;
        }
        else if (numFile == "3") {
            document.importar_datos.file3.value=document.importar_datos.fichero03.value;
        }
        else if (numFile == "4") {
            document.importar_datos.file4.value=document.importar_datos.fichero04.value;
        }
        else if (numFile == "5") {
            document.importar_datos.file5.value=document.importar_datos.fichero05.value;
        }       
    }

</script>
</HEAD>
<BODY bgcolor=<%=color_blau%> class="BODY_ASP">
<%Sub BuildUploadRequest(RequestBin)
        'extensions allowed
    uploadsFileExtensionAllowed=Array("gif","jpg","jpeg","png","bmp","doc","docx","zip","rar","gz","tgz","tz","pdf","txt","xls","xlsx","xlsb","ppt","pps","pptx","log","sda","sdb","sdc","sdd","sdg","sdm","sds","sdv","sdw","ttf","txt","avi","mkv","mov","tiff","odt","fodt","ods","fods","odp","fodp","odb","odg","fodg","odf","rtf","pages","keynote","csv","xlsx","xlx","pdf")
   
  'Get the boundary
  PosBeg = 1
  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
  if PosEnd = 0 then
    Response.Write "<b>" & LitFormENCTYPE & "</b><br>"
    Response.End
  end if
  boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
  boundaryPos = InstrB(1,RequestBin,boundary)
  'Get all data inside the boundaries
  Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
    'Members variable of objects are put in a dictionary object
    Dim UploadControl
    Set UploadControl = CreateObject("Scripting.Dictionary")
    'Get an object name
    Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
    Pos = InstrB(Pos,RequestBin,getByteString("name="))
    PosBeg = Pos+6
    PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
    Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
    PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
    PosBound = InstrB(PosEnd,RequestBin,boundary)
    'Test if object is of file type
    If  PosFile<>0 AND (PosFile<PosBound) Then
      'Get Filename, content-type and content of file
      PosBeg = PosFile + 10
      PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
      FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      FileName = Mid(FileName,InStrRev(FileName,"\")+1)

       ' Parse File Ext
            If Not InStrRev(FileName, ".") = 0 Then
                FileExt = Mid(FileName, InStrRev(FileName, ".") + 1)
                FileExt = UCase(FileExt)
            End If

            uploadFileExtIsAllowed=0
     
            if isArray(uploadsFileExtensionAllowed) then 
    
               For i = 0 to ubound(uploadsFileExtensionAllowed)            
                    If StrComp(FileExt, uploadsFileExtensionAllowed(i),vbTextCompare) = 0 Then
                      uploadFileExtIsAllowed=1
                      Exit For
                    End If
               Next
            end if
            if uploadFileExtIsAllowed=1 then  
                  'Add filename to dictionary object
                  UploadControl.Add "FileName", FileName
                  Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
                  PosBeg = Pos+14
                  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
                  'Add content-type to dictionary object
                  ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
                  UploadControl.Add "ContentType",ContentType
                  'Get content of object
                  PosBeg = PosEnd+4
                  PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
                  Value = FileName
                  ValueBeg = PosBeg-1
                  ValueLen = PosEnd-Posbeg
            else
                response.Write "<script>alert('File not found!');</script>"
                response.End
               
            end if
    Else
      'Get content of object
      Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
      PosBeg = Pos+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      ValueBeg = 0
      ValueEnd = 0
    End If
    'Add content to dictionary object
    UploadControl.Add "Value" , Value
    UploadControl.Add "ValueBeg" , ValueBeg
    UploadControl.Add "ValueLen" , ValueLen
    'Add dictionary object to main dictionary
    UploadRequest.Add name, UploadControl
    'Loop to next object
    BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
  Loop
End Sub

'String to byte string conversion
Function getByteString(StringStr)
  For i = 1 to Len(StringStr)
 	  char = Mid(StringStr,i,1)
	  getByteString = getByteString & chrB(AscB(char))
  Next
End Function

'Byte string to string conversion
Function getString(StringBin)
  getString =""
  For intCount = 1 to LenB(StringBin)
	  getString = getString & chr(AscB(MidB(StringBin,intCount,1)))
  Next
End Function

Function UploadFormRequest(name)
  on error resume next
  if UploadRequest.Item(name) then
    UploadFormRequest = UploadRequest.Item(name).Item("Value")
  end if
End Function

sub PintaComprobacion(cursor,proc,form)
	Dim cont%>
	<table style="border-collapse:collapse"><%
		DrawFila color_fondo
			DrawCelda2 "TDBORDECELDA8", "left", true,LitFichero & " : " & filesinpath
		CloseFila%>
	</table>
	<br>
	<table style="border-collapse:collapse"><%
		DrawFila color_fondo
			DrawCelda2Span "TDBORDECELDA8", "left", true,LitRegimportar,22
		CloseFila
			if proc="ARTICULOS" then
				DrawFila color_terra
					DrawCelda2 "TDBORDECELDA7", "left", true,LitReferencia
					DrawCelda2 "TDBORDECELDA7", "left", true,LitNombre
					DrawCelda2 "TDBORDECELDA7", "left", true,LitFamilia
					DrawCelda2 "TDBORDECELDA7", "right", true,LitImporte
					DrawCelda2 "TDBORDECELDA7", "right", true,LitPVP
					DrawCelda2 "TDBORDECELDA7", "left", true,LitBeneficio
					DrawCelda2 "TDBORDECELDA7", "right", true,LitDescuento
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCodBarras
					DrawCelda2 "TDBORDECELDA7", "left", true,LitNproveedor
					DrawCelda2 "TDBORDECELDA7", "left", true,LitRefProv
					DrawCelda2 "TDBORDECELDA7", "left", true,LitIva
					DrawCelda2 "TDBORDECELDA7", "left", true,litControlarStock
					DrawCelda2 "TDBORDECELDA7", "left", true,litStock
					DrawCelda2 "TDBORDECELDA7", "left", true,LitStockMin
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDivisa
					DrawCelda2 "TDBORDECELDA7", "left", true,litAlmacen
				closefila
				cont=1
				while not cursor.eof and cont<=10
					DatosPartidos=split(cursor(2),";")
					if form="CSV" then
						DrawFila color_blau
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(0)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(1)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(2)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(4)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(5)
							DrawCelda2 "TDBORDECELDA7", "right", false,"&nbsp;"
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(6)
							DrawCelda2 "TDBORDECELDA7", "center", false,DatosPartidos(10)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(8)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(9)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(7)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(11)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(13)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(14)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(3)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(12)
						CloseFila
					else
						DrawFila color_blau
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(0)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(1)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(2)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(4)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(5)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(6)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(7)
							DrawCelda2 "TDBORDECELDA7", "center", false,DatosPartidos(11)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(9)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(10)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(8)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(12)
							DrawCelda2 "TDBORDECELDA7", "left", false,"&nbsp;"
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(13)
							DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(3)
							DrawCelda2 "TDBORDECELDA7", "left", false,"&nbsp;"
						CloseFila
					end if
					cursor.movenext
					cont=cont+1
				wend
			elseif proc="CLIENTES" then
				DrawFila color_terra
					DrawCelda2 "TDBORDECELDA7", "left", true,LitNcliente
					DrawCelda2 "TDBORDECELDA7", "left", true,LitRsocial
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCif
					DrawCelda2 "TDBORDECELDA7", "left", true,Littelefono1
					DrawCelda2 "TDBORDECELDA7", "left", true,LitTelefono2
					DrawCelda2 "TDBORDECELDA7", "left", true,LitFax
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDomicilio
					DrawCelda2 "TDBORDECELDA7", "left", true,LitPoblacion
					DrawCelda2 "TDBORDECELDA7", "left", true,LitProvincia
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCP
					DrawCelda2 "TDBORDECELDA7", "left", true,LitPais
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCContable
					DrawCelda2 "TDBORDECELDA7", "left", true,LitFormaPago
					DrawCelda2 "TDBORDECELDA7", "left", true,LitBanco
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDirBanco
					DrawCelda2 "TDBORDECELDA7", "left", true,LitPobBanco
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDivisa
					DrawCelda2 "TDBORDECELDA7", "left", true,LiEmail
					DrawCelda2 "TDBORDECELDA7", "left", true,LitContacto
					DrawCelda2 "TDBORDECELDA7", "left", true,LitObservaciones
					DrawCelda2 "TDBORDECELDA7", "left", true,LitComercial
				closefila
				cont=1
				while not cursor.eof and cont<=10
					DatosPartidos=split(cursor(2),";")
					if form="CSV" then
						DrawFila color_blau
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(0)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(1)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(3)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(10)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(11)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(12)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(5)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(6)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(7)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(8)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(9)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(18)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(16)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(19)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(20)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(21)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(15)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(14)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(13)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(22)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(17)
						CloseFila
					else
						DrawFila color_blau
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(0)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(1)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(8)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(5)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(6)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(7)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(3)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(4)
							DrawCelda2 "TDBORDECELDA7", "left", false,"&nbsp;"
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(19)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(17)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(15)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(14)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(10)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(11)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(12)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(18)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(16)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(9)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(13)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(20)
						CloseFila
					end if
					cursor.movenext
					cont=cont+1
				wend
			elseif proc="PROVEEDORES" then
				DrawFila color_terra
					DrawCelda2 "TDBORDECELDA7", "left", true,LitNproveedor
					DrawCelda2 "TDBORDECELDA7", "left", true,LitRsocial
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCif
					DrawCelda2 "TDBORDECELDA7", "left", true,Littelefono1
					DrawCelda2 "TDBORDECELDA7", "left", true,LitTelefono2
					DrawCelda2 "TDBORDECELDA7", "left", true,LitFax
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDomicilio
					DrawCelda2 "TDBORDECELDA7", "left", true,LitPoblacion
					DrawCelda2 "TDBORDECELDA7", "left", true,LitProvincia
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCP
					DrawCelda2 "TDBORDECELDA7", "left", true,LitPais
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCContable
					DrawCelda2 "TDBORDECELDA7", "left", true,LitFormaPago
					DrawCelda2 "TDBORDECELDA7", "left", true,LitBanco
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDirBanco
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDivisa
					DrawCelda2 "TDBORDECELDA7", "left", true,LiEmail
				closefila
				cont=1
				while not cursor.eof and cont<=10
					DatosPartidos=split(cursor(2),";")
					if form="CSV" then
	           		DrawFila color_blau
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(0)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(1)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(10)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(7)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(8)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(9)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(2)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(3)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(4)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(5)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(6)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(13)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(12)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(14)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(15)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(16)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(17)
						CloseFila
					else
						DrawFila color_blau
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(0)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(1)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(7)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(4)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(5)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(6)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(2)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(3)
							DrawCelda2 "TDBORDECELDA7", "left", false,"&nbsp;"
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(14)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(12)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(11)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(10)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(8)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(9)
							DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(13)
							DrawCelda2 "TDBORDECELDA7", "left", false,"&nbsp;"
						CloseFila
					end if
					cursor.movenext
					cont=cont+1
				wend
            elseif proc="TARJETAS" then
				DrawFila color_terra

					DrawCelda2 "TDBORDECELDA7", "left", true,Litcardnumber
					DrawCelda2 "TDBORDECELDA7", "left", true,Litexpiratedate
					DrawCelda2 "TDBORDECELDA7", "right", true,Litcardtype
					DrawCelda2 "TDBORDECELDA7", "right", true,Litbalance
					DrawCelda2 "TDBORDECELDA7", "left", true,LitNombre
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCif
					DrawCelda2 "TDBORDECELDA7", "left", true,LitDomicilio
					DrawCelda2 "TDBORDECELDA7", "left", true,LitPoblacion
					DrawCelda2 "TDBORDECELDA7", "left", true,LitCP
					DrawCelda2 "TDBORDECELDA7", "left", true,LitProvincia
					DrawCelda2 "TDBORDECELDA7", "left", true,LitPais
					DrawCelda2 "TDBORDECELDA7", "left", true,LitTelefono1
					DrawCelda2 "TDBORDECELDA7", "left", true,Littelefono2
					DrawCelda2 "TDBORDECELDA7", "left", true,LiEmail
					DrawCelda2 "TDBORDECELDA7", "left", true,Litborndate
					DrawCelda2 "TDBORDECELDA7", "right", true,Litpromonotiftype
				closefila
				cont=1
				while not cursor.eof and cont<=10
					DatosPartidos=split(cursor(2),";")
					DrawFila color_blau
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(0)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(1)
						DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(2)
						DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(3)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(4)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(5)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(6)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(7)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(8)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(9)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(10)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(11)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(12)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(13)
						DrawCelda2 "TDBORDECELDA7", "left", false,DatosPartidos(14)
						DrawCelda2 "TDBORDECELDA7", "right", false,DatosPartidos(15)
					CloseFila
                    cursor.movenext
					cont=cont+1
               wend
			end if
		%></table><%
end sub

sub PintaResultados(mens,cursor)

	if cursor.state=0 then
		%><script language="JavaScript" type="text/javascript">
		      alert("<%=mens%>");
		      document.location = "importar_datos.asp?mode=select&ncliente=<%=enc.EncodeForJavascript(CodEmpresa)%>&borra=SI&s=<%=enc.EncodeForJavascript(suministros)%>&tdocumento=<%=enc.EncodeForJavascript(tituloDoc)%>";
		      parent.botones.document.location = "importar_datos_bt.asp?mode=select";
		</script><%
	else
		%><table style="border-collapse:collapse"><%
			DrawFila color_fondo
				DrawCelda2 "TDBORDECELDA8", "left", true,LitFichero & " : " & filesinpath
			CloseFila
		%></table>
		<br>
		<table style="border-collapse:collapse"><%
			DrawFila color_fondo
				DrawCelda2Span "TDBORDECELDA8", "left", true,ucase(mens),11
			CloseFila
			DrawFila color_terra
				DrawCelda2 "TDBORDECELDA7", "left", true,LitFecha
				DrawCelda2 "TDBORDECELDA7", "left", true,LitSurtidor
				DrawCelda2 "TDBORDECELDA7", "right", true,LitLitros
				DrawCelda2 "TDBORDECELDA7", "right", true,LitPrecio
				DrawCelda2 "TDBORDECELDA7", "right", true,LitImporte
				DrawCelda2 "TDBORDECELDA7", "center", true,LitCliente
				DrawCelda2 "TDBORDECELDA7", "left", true,LitTarjeta
				DrawCelda2 "TDBORDECELDA7", "left", true,LitDato1
				DrawCelda2 "TDBORDECELDA7", "left", true,LitDato2
				DrawCelda2 "TDBORDECELDA7", "left", true,LitTurno
				DrawCelda2 "TDBORDECELDA7", "left", true,LitFactura
			closefila

			while not cursor.eof
				DrawFila color_blau
					DrawCelda2 "TDBORDECELDA7", "left", false,cursor("fecha")
					DrawCelda2 "TDBORDECELDA7", "left", false,cursor("nsurtidor")
					DrawCelda2 "TDBORDECELDA7", "right", false,cursor("cantidad")
					DrawCelda2 "TDBORDECELDA7", "right", false,cursor("plitro")
					DrawCelda2 "TDBORDECELDA7", "right", false,cursor("importe")
					DrawCelda2 "TDBORDECELDA7", "center", false,cursor("ncliente")
					DrawCelda2 "TDBORDECELDA7", "left", false,cursor("ntarjeta")
					DrawCelda2 "TDBORDECELDA7", "left", false,cursor("datoadd1")
					DrawCelda2 "TDBORDECELDA7", "left", false,cursor("datoadd2")
					DrawCelda2 "TDBORDECELDA7", "left", false,cursor("turno")
					DrawCelda2 "TDBORDECELDA7", "left", false,cursor("nfactura")
				CloseFila
				cursor.movenext
			wend
		%></table><%
		cursor.close
		%><script language="JavaScript" type="text/javascript">
		      alert("<%=Mensaje%>");
		      parent.botones.document.location = "importar_datos_bt.asp?mode=informe";
		</script><%
	end if
end sub

function GetLitErr(err, msg)
    select case err
        case 0:
            GetLitErr=msg&" "&LitCardsOk
        case 1:
            GetLitErr=LitErrFile
            if msg&"">"" then GetLitErr=GetLitErr&": "&msg
        case 2:
            GetLitErr=LitErrAsociation
        case 3:
            GetLitErr=LitErrBorndate
        case 4:
            GetLitErr=LitErrBalance
        case 5:
            GetLitErr=LitErrCardType
        case 6:
            GetLitErr=LitErrCardRepeated
        case 7:
            GetLitErr=LitErrCardLen
        case 8:
            GetLitErr=LitErrCardNoNum
        case 9:
            GetLitErr=LitErrNotChanel
        case 10:
            GetLitErr=LitErrCardExists
        case 11:
            GetLitErr=LitErrZip
        case 12:
            GetLitErr=LitErrFileFormat
            if msg&"">"" then GetLitErr=GetLitErr&": "&msg
        case 13:
            GetLitErr=LitErrExpirationdate
        
    end select
end function

'**********************************************************************************************************
' CODIGO PRINCIPAL
'**********************************************************************************************************

if accesoPagina(session.sessionid,session("usuario"))=1 then

mode= request.querystring("mode")
if request.querystring("s") & "" <> "" then
	suministros=LimpiaCadena(request.querystring("s"))
end if
if request.querystring("f") & "" <> "" then
	formato=LimpiaCadena(request.querystring("f"))
end if
if request.querystring("p") & "" <> "" then
	proceso=LimpiaCadena(request.querystring("p"))
end if
if request.querystring("ncliente") & "" <> "" then
	ncliente=LimpiaCadena(request.querystring("ncliente"))
else
	ncliente=""
end if

CodEmpresa=session("ncliente")

%><form name="importar_datos" method="post" ENCTYPE="multipart/form-data" class="col-lg-8 col-md-10 col-sm-12 col-xxs-12 overflowxauto">
	<input type="hidden" name="codempresa" value="<%=enc.EncodeForHtmlAttribute(null_s(CodEmpresa))%>">
    <%tituloDoc=Request.QueryString("tdocumento")&"" %>
    <input type="hidden" name="tituloDoc" value="<%=enc.EncodeForHtmlAttribute(null_s(tituloDoc))%>"><%
    
		session_ncliente=request.querystring("ncliente") & ""

        if session_ncliente = "00000" then      
        %><script language="JavaScript" type="text/javascript">
	            alert("<%=LitMsgImpSisGestion%>");
        </script><%
        response.end
        end if

	if session("ncliente")&""="00000" then
        PaintHeaderPopUp "ges_clientes.asp", tituloDoc
    else
	    PintarCabecera "Importar_Datos.asp"
    end if

    si_tiene_modulo_Fidelizacion=ModuloContratado(session("ncliente"),ModFidelizacion)
    if si_tiene_modulo_Fidelizacion=0 then si_tiene_modulo_Fidelizacion=ModuloContratado(session("ncliente"),ModFidelizacionPremium)

	%><br><br><%

	if mode="select" then 
		if request.querystring("borra")="SI" then
			'borrar carpeta remota
			newfolderpath = GetPathImport(request.ServerVariables("LOCAL_ADDR"), carpetaproduccion) 
            newfolderpath = newfolderpath  & CodEmpresa
			set filesys=CreateObject("Scripting.FileSystemObject")
			'Se borra la carpeta del servidor
			If filesys.FolderExists(newfolderpath) Then
				on error resume next
				filesys.DeleteFolder newfolderpath
				on error goto 0
			End If
			set filesys=nothing
		end if%>
		<input type="Hidden" name="s" value="<%=enc.EncodeForHtmlAttribute(null_s(suministros))%>">
		<table style="border-collapse:collapse">
			<%DrawFila color_fondo
				%><th colspan="2"><%= LitProceso %></th>
					<th colspan="2"><%= LitFormato %></th>
					<td></td>
					<th style="width:65%;"><%= LiFicheros %></th><%
			CloseFila
			if suministros<>"1" then
				DrawFila color_blau%>
					<td CLASS=TDBORDECELDAC8><input type="radio" name="proceso" value="clientes" checked onclick="Repinta();"></td><%
					DrawCelda2 "TDBORDECELDA8", "left", false,LitClientes%>
					<td CLASS=TDBORDECELDAC8><input type="radio" name="formato" value="dbf" checked onclick="Repinta();"></td><%
					DrawCelda2 "TDBORDECELDA8", "left", false,LitFacturaPlus%>
					<td style="background-color:white;"></td>
					<td CLASS=TDBORDECELDA8 rowspan="5" style="background-color:white;padding: 0;vertical-align:top;">
						<table>
							<tr style="display:none;"></tr>
							<tr id="trfile01">                          
                                <td CLASS=CELDA id="idlitfile01"><span><%=LitClientesDBF%> : </span></td>
                                <td id="idfile01"><INPUT style="display: none;" id="ponfich1" class="CELDA7" TYPE="file" size="40" NAME="fichero01"
                                    onclick="onSelectFileClick('1');" onchange="onSelectFileChange('1');">
                                    <input size="60" class="CELDA7" type="button" name="examinar1" value="<%=LitSelectFile %>" 
                                        onclick="javascript: document.getElementById('ponfich1').click();"/>
                                    <input style="margin-right: 5px;" id="input_file1" readonly="READONLY" type="text" name="file1" maxlength="255" size="15" value=""/>
                                </td>
                                <td class=CELDAREDBOLD style="width:5%;">[*]</td>
                                <td id="idlitfile01p" style="text-align:center;display:none;" class="CELDA"><span></span></td>
							</tr>   
							<tr id="trfile02">
								<td CLASS=CELDA id="idlitfile02"><span><%=LitfpagoDBF%> : </span></td>
                                <td id="idfile02"><INPUT style="display: none;" id="ponfich2" class="CELDA7" TYPE="file" size="40" NAME="fichero02"
                                    onclick="onSelectFileClick('2');" onchange="onSelectFileChange('2');">
                                    <input size="60" class="CELDA7" type="button" name="examinar2" value="<%=LitSelectFile %>" 
                                        onclick="javascript: document.getElementById('ponfich2').click();"/>
                                    <input style="margin-right: 5px;" id="input_file2" readonly="READONLY" type="text" name="file2" maxlength="255" size="15" value=""/>
                                </td>
                                <td class=CELDAREDBOLD></td>
                                <td id="idlitfile02p" style="text-align:center;display:none;" class="CELDA"><span></span></td>
							</tr>
							<tr id="trfile03">
								<td CLASS=CELDA id="idlitfile03"><span><%=LitagentesDBF%> : </span></td>
                                <td id="idfile03"><INPUT style="display: none;" id="ponfich3" class="CELDA7" TYPE="file" size="40" NAME="fichero03"
                                    onclick="onSelectFileClick('3');" onchange="onSelectFileChange('3');">
                                    <input size="60" class="CELDA7" type="button" name="examinar3" value="<%=LitSelectFile %>" 
                                        onclick="javascript: document.getElementById('ponfich3').click();"/>
                                    <input style="margin-right: 5px;" id="input_file3" readonly="READONLY" type="text" name="file3" maxlength="255" size="15" value=""/>
                                </td>
                                <td class=CELDAREDBOLD></td>
                                <td id="idlitfile03p" style="text-align:center;display:none;" class="CELDA"><span></span></td>
							</tr>   
							<tr id="trfile04">
								<td CLASS=CELDA id="idlitfile04"><span><%=LitprovincDBF%> : </span></td>
                                <td id="idfile04"><INPUT style="display: none;" id="id_fichero04" class="CELDA7" TYPE="file" size="40" NAME="fichero04"
                                    onclick="onSelectFileClick('4');" onchange="onSelectFileChange('4');">
                                    <input size="60" class="CELDA7" type="button" name="examinar4" value="<%=LitSelectFile %>" 
                                        onclick="javascript: document.getElementById('id_fichero04').click();"/>
                                    <input style="margin-right: 5px;" id="input_file4" readonly="READONLY" type="text" name="file4" maxlength="255" size="15" value=""/>
                                </td>
                                <td class=CELDAREDBOLD></td>
                                <td id="idlitfile04p" style="text-align:center;display:none;" class="CELDA"><span></span></td>
							</tr>
                        <%if este_se_oculta=1 then%>
							<tr id="trfile05">
								<td CLASS=CELDA id="idlitfile05"> : </td>
                                <td id="idfile05"><INPUT style="display: none;" id="ponfich5" class="CELDA7" TYPE="file" size="40" NAME="fichero05"
                                    onclick="onSelectFileClick('5');" onchange="onSelectFileChange('5');">
                                    <input size="60" class="CELDA7" type="button" name="examinar5" value="<%=LitSelectFile %>" 
                                        onclick="javascript: document.getElementById('ponfich5').click();"/><span></span>
                                    <input style="margin-right: 5px;" id="input_file5" readonly="READONLY" type="text" name="file5" maxlength="255" size="15" value=""/>
                                </td>
                                <td class=CELDAREDBOLD></td>
							</tr>
                        <%end if %>
						</table>
					</td><%
				CloseFila
				DrawFila color_blau%>
					<td CLASS=TDBORDECELDAC8><input type="radio" name="proceso" value="proveedores" onclick="Repinta();"></td><%
					DrawCelda2 "TDBORDECELDA8", "left", false,LitProveedores%>
					<td CLASS=TDBORDECELDAC8><input type="radio" name="formato" value="csv" onclick="Repinta();"></td><%
					DrawCelda2 "TDBORDECELDA8", "left", false,LitCSV
				CloseFila
				DrawFila color_blau%>
					<td CLASS=TDBORDECELDAC8><input type="radio" name="proceso" value="articulos" onclick="Repinta();"></td><%
					DrawCelda2 "TDBORDECELDA8", "left", false,LitArticulos%>
					<td CLASS=TDBORDECELDA8></td>
					<td CLASS=TDBORDECELDA8></td><%
				CloseFila
                if si_tiene_modulo_Fidelizacion then
                    DrawFila color_blau%>
					    <td CLASS=TDBORDECELDAC8><input type="radio" name="proceso" value="tarjetas" onclick="Repinta();"></td><%
					    DrawCelda2 "TDBORDECELDA8", "left", false,LitFidelizacion%>
					    <td CLASS=TDBORDECELDA8></td>
					    <td CLASS=TDBORDECELDA8></td><%
				    CloseFila
                end if
				DrawFila color_blau
				    ''MPC 22/06/2010 Comentado a petición de Jesús%>
					<td CLASS=TDBORDECELDAC8 style="height:26px"><!--<input type="radio" name="proceso" value="segCom" onclick="Repinta();">--></td><%
					'DrawCelda2 "TDBORDECELDA8", "left", false,"Seguimiento comercial"
					DrawCelda2 "TDBORDECELDA8", "left", false,"&nbsp;"
					''FIN MPC 22/06/2010%>
					<td CLASS=TDBORDECELDA8></td>
					<td CLASS=TDBORDECELDA8></td><%
				CloseFila
                %>
                <tr>
                <td colspan="4" class="CELDA" style="background: #FDFDFD">
                <a class="CELDAREF7" href="javascript:AbrirVentana('../../Documentos/plantillas_importacion_datos/Manual GESTOR - Importación de Datos.pdf','P','550','950');"><%=LitVerManual%></a>
                </td>
                </tr>
                <%
			else
				DrawFila color_blau%>
					<td CLASS=TDBORDECELDAC8><input type="radio" name="proceso" value="suministros" checked onclick=""></td><%
					DrawCelda2 "TDBORDECELDA8", "left", false,LitSuministros%>
					<td CLASS=TDBORDECELDAC8><input type="radio" name="formato" value="acc" checked onclick=""></td><%
					DrawCelda2 "TDBORDECELDA8", "left", false,LitAccom100%>
					<td CLASS=TDBORDECELDA8>
						<table>
							<%DrawFila Color_Blau%>
								<td CLASS=CELDA id="idlitfile01"></td><td id="idfile01"><INPUT class="CELDA7" TYPE="file" size="40" NAME="fichero01"></td><td class=CELDAREDBOLD>[*]</td>
							<%CloseFila%>
						</table>
					</td><%
				CloseFila
			end if%>
		</table><br>
		<div class=CELDAREDBOLD><%=LitObligatorio%></div>
		<br>
		<%waitboxoculto LitMsgEnviando
	elseif mode="procesa" then
		'Creamos una carpeta para el cliente para evitar errores de simultaneidad entre clientes
		dim filesys, newfolder, newfolderpath
		newfolderpath = GetPathImport(request.ServerVariables("LOCAL_ADDR"), carpetaproduccion)
        newfolderpath = newfolderpath & CodEmpresa		

		set filesys=CreateObject("Scripting.FileSystemObject")
		'Para evitar restos de algún posible error anterior, primero se borra
		If filesys.FolderExists(newfolderpath) Then
			on error resume next
   			filesys.DeleteFolder newfolderpath
			on error goto 0
		End If
		set filesys=nothing

		'Ahora es cuando se crea
		set filesys=CreateObject("Scripting.FileSystemObject")
		If Not filesys.FolderExists(newfolderpath) Then
		   Set newfolder = filesys.CreateFolder(newfolderpath)
		End If
		set filesys=nothing

		RequestBin = Request.BinaryRead(Request.TotalBytes)
  		Dim UploadRequest
  		Set UploadRequest = CreateObject("Scripting.Dictionary")
  		BuildUploadRequest RequestBin
  		if suministros & "" = "" then suministros=UploadRequest.Item("s").Item("Value")
		%><input type="Hidden" name="s" value="<%=enc.EncodeForHtmlAttribute(null_s(suministros))%>"><%

		Claves = UploadRequest.Keys
        
  		for indice = 0 to UploadRequest.Count - 1
    		Clave = Claves(indice)
    		'Guardar los ficheros

		    if UploadRequest.Item(Clave).Item("FileName") <> "" then
				GP_value = UploadRequest.Item(Clave).Item("Value")
		      	GP_valueBeg = UploadRequest.Item(Clave).Item("ValueBeg")
		      	GP_valueLen = UploadRequest.Item(Clave).Item("ValueLen")

				if GP_valueLen = 0 then
		    		%><script language="JavaScript" type="text/javascript">
		    		      alert("<%=LitMsgErrUploading%>")
					</script><%
		   			response.End
				end if

				'Se crean instancias de Streams
				Dim GP_strm1, GP_strm2
				Set GP_strm1 = Server.CreateObject("ADODB.Stream")
				Set GP_strm2 = Server.CreateObject("ADODB.Stream")

				'Se abre el stream
				GP_strm1.Open
				GP_strm1.Type = 1 'Binario
				GP_strm2.Open
				GP_strm2.Type = 1 'Binario

				GP_strm1.Write RequestBin
				GP_strm1.Position = GP_ValueBeg
				GP_strm1.CopyTo GP_strm2,GP_ValueLen

				'Crear y escribir el fichero
				on error resume next
				GP_strm2.SaveToFile Trim(newfolderpath)& "\" & UploadRequest.Item(Clave).Item("FileName"),2

				if err then
                    set conn = nothing
                    set conn = Server.CreateObject("ADODB.Connection")
                    conn.open DSNIlion
			        conn.CommandTimeout = 0
                    strAudit="EXEC almacenar_incidencia @fecha='', @usuario='" & session("usuario") & "', @empresa='" & CodEmpresa & "', @ip_servidor='" & Request.ServerVariables("REMOTE_ADDR") & "', @nombre_servidor='" & Request.ServerVariables("REMOTE_ADDR") & "', @num_error='" & err.number & "', @desc_error='" & err.Description & " File:" &  Replace(Trim(newfolderpath),"\","/") & "/" & UploadRequest.Item(Clave).Item("FileName") & "', @fichero='ilionp/services/importar_datos.asp', @linea='0', @fuente_fichero='" & err.Description & "', @tipo_explorador='', @parametros_get='', @parametros_post=''"
			        set rst = conn.execute(strAudit)
                    rst.Close
			        set conn = nothing
					%><script language="JavaScript" type="text/javascript">
					      alert("<%=LitMsgErrUploading%>");
					      document.location = "importar_datos.asp?mode=select&ncliente=<%=enc.EncodeForJavascript(CodEmpresa)%>&borra=SI&s=<%=enc.EncodeForJavascript(suministros)%>&tdocumento=<%=enc.EncodeForJavascript(tituloDoc)%>";
					      parent.botones.document.location = "importar_datos_bt.asp?mode=select";
					</script><%
					response.End
				end if
			end if
		next

		formato=ucase(UploadRequest.Item("formato").Item("Value"))
		proceso=ucase(UploadRequest.Item("proceso").Item("Value"))

		waitboxoculto LitMsgVerificando
		response.write("<script language='JavaScript' type='text/javascript'>document.getElementById('waitBoxOculto').style.visibility='visible';</script>")
		response.flush

		File01Saved = "NO":File02Saved = "NO":File03Saved = "NO":File04Saved = "NO"
		if UploadRequest.Item("fichero01").Item("FileName")<>"" then
			File01Saved="SI"
			filesinpath=UploadRequest.Item("fichero01").Item("FileName")
		end if
		if UploadRequest.Item("fichero02").Item("FileName")<>"" then File02Saved="SI"
		if UploadRequest.Item("fichero03").Item("FileName")<>"" then File03Saved="SI"
		if UploadRequest.Item("fichero04").Item("FileName")<>"" then File04Saved="SI"

		'Cerramos la captura de errores abierta antes de GP_strm2.SaveToFile....
		on error goto 0

		select case proceso
			case "CLIENTES" :
                strSelectNombreUsuario = "select nombre from indice with(nolock) where entrada = ?"
                NombreUsuario = DLookupP1(strSelectNombreUsuario,session("usuario")&"",adVarchar,50,DSNIlion&"")
                'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
				
                if formato="DBF" then
					mensaje=""
					'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
					strselect="EXEC ImportarClientesFraPlus @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
					strcheck="EXEC FormatoClientesFraPlus @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
				else
					mensaje=""
					'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
					strselect="EXEC ImportarClientesCSV @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
					strcheck="EXEC FormatoClientesCSV @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
				end if
			case "PROVEEDORES" :
                strSelectNombreUsuario = "select nombre from indice with(nolock) where entrada = ?"
                NombreUsuario = DLookupP1(strSelectNombreUsuario,session("usuario")&"",adVarchar,50,DSNIlion&"")
				if formato="DBF" then
					mensaje=""
					'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
					strselect="EXEC ImportarProveedoresFraPlus @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
					strcheck="EXEC FormatoProveedoresFraPlus @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
				else
					mensaje=""
					'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
					strselect="EXEC ImportarProveedoresCSV @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
					strcheck="EXEC FormatoProveedoresCSV @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
				end if
			case "ARTICULOS" :
                strSelectNombreUsuario = "select nombre from indice with(nolock) where entrada = ?"
                NombreUsuario = DLookupP1(strSelectNombreUsuario,session("usuario")&"",adVarchar,50,DSNIlion&"")
				if formato="DBF" then
					mensaje=""
					'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
					strselect="EXEC ImportarArticulosFraPlus @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
					strcheck="EXEC FormatoArticulosFraPlus @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
				else
					mensaje=""
					'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
					strselect="EXEC ImportarArticulosCSV @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
					strcheck="EXEC FormatoArticulosCSV @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "'"
				end if
			case "SUMINISTROS" :
                strSelectNombreUsuario = "select nombre from indice with(nolock) where entrada = ?"
                NombreUsuario = DLookupP1(strSelectNombreUsuario,session("usuario")&"",adVarchar,50,DSNIlion&"")

				if formato="ACC" then
					mensaje=""
					'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)
					strselect="EXEC ImportarSuministrosACCOM100 @ncliente='" & CodEmpresa & "', @login='" & session("usuario") & "', @usuario='" & NombreUsuario & "', @carpetaproduccion='" & carpetaproduccion & "', @ip='" & Request.ServerVariables("REMOTE_ADDR") & "', @f1='" & File01Saved & "', @f2='" & File02Saved & "', @f3='" & File03Saved & "', @f4='" & File04Saved & "', @pnomf1='" & filesinpath & "'"
				end if
            case "TARJETAS" :
				mensaje=""
                strSelectNombreUsuario = "select nombre from indice with(nolock) where entrada = ?"
                NombreUsuario = DLookupP1(strSelectNombreUsuario,session("usuario")&"",adVarchar,50,DSNIlion&"")
				'NombreUsuario=d_lookup("nombre","indice","entrada='" & session("usuario") & "'",DSNIlion)

				strselect="EXEC ImportCardsCSV @companyid='" & CodEmpresa & "', @filenm='"& UploadRequest.Item("fichero01").Item("FileName")&"', @userid='" & session("usuario") & "', @productionfolder='" & carpetaproduccion & "' "
				strcheck=strselect&" , @modecheck=1"
				
		end select
		set rst = Server.CreateObject("ADODB.RECORDSET")
	   	crear ="CREATE TABLE [dbo].[" & session("usuario") & "] (ID int IDENTITY,MENSAJE varchar(500))"
		strdrop ="if exists (select * from sysobjects where id = object_id('[" & session("usuario") & "]') and sysstat & 0xf = 3) drop table [" & session("usuario") & "]"

        if session("ncliente")<>SISTEMA_GESTION then
            dsnCliente = session("dsn_cliente")
        else
            strSelectDsnCliente = "select dsn from clientes with(nolock) where ncliente = ?"
            dsnCliente = DLookupP1(strSelectDsnCliente,CodEmpresa,adVarchar,5,DSNIlion&"")     
            'dsnCliente=d_lookup("dsn","clientes","ncliente='" & CodEmpresa & "'",DSNIlion)
        end if    
	    initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")
	    donde=inStr(1,DSNImport,"Initial Catalog=",1)
	    donde_fin=InStr(donde,DSNImport,";",1)
        if donde_fin=0 then
	        donde_fin=len(DSNImport)
	    end if
	    cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
	    dsnCliente=cadena_dsn_final
        dsnTmpTable=cadena_dsn_final

        if formato ="ACC" then
            rst.open strdrop,dsnCliente,adUseClient,adLockReadOnly
	        rst.open crear,dsnCliente,adUseClient,adLockReadOnly
	        'GrantUser session("usuario"), dsnCliente
        else
            ''ricardo 14-4-2010 se cambia el DSN
	        ''rst.open strdrop,DSNImport,adUseClient,adLockReadOnly
	        ''rst.open crear,DSNImport,adUseClient,adLockReadOnly
	        ''GrantUser session("usuario"), DSNImport
	        rst.open strdrop,dsnCliente,adUseClient,adLockReadOnly
	        rst.open crear,dsnCliente,adUseClient,adLockReadOnly
	        'GrantUser session("usuario"), dsnCliente
        end if
		set rst = nothing

		set conn = Server.CreateObject("ADODB.Connection")
		if formato="CSV" then
		    ''ricardo 14-4-2010 se cambia el DSN
			''conn.open DSNImport
            if proceso="TARJETAS" then
                initial_catalogC=encontrar_datos_dsn(dsnIlion,"Initial Catalog=")
	            donde=inStr(1,DSNImport,"Initial Catalog=",1)
	            donde_fin=InStr(donde,DSNImport,";",1)
                if donde_fin=0 then
	                donde_fin=len(DSNImport)
	            end if
	            cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
	            dsnCliente=cadena_dsn_final
                conn.open dsnCliente
            else
                conn.open dsnCliente
            end if
			
			conn.CommandTimeout = 0
			''on error resume next

			set rst = conn.execute(strcheck)
            ''if err.number<>0 then
            ''    response.write("el cadena_dsn_final es-" & cadena_dsn_final & "-<br>")
            ''    response.write("el strcheck es-" & strcheck & "-<br>")
            ''    response.end
            ''end if
            ''on error goto 0

			if cint(rst(0))=0 then 'FORMATO CORRECTO
				NumRegistros=rst(1) & " "
				PintaComprobacion rst,proceso,formato
				rst.close
				'on error goto 0
				set conn = nothing
				set rst=nothing
				%><input type="Hidden" name="qry" value="<%=enc.EncodeForHtmlAttribute(null_s(strselect))%>">
				<script language="JavaScript" type="text/javascript">
				    document.getElementById("waitBoxOculto").style.visibility = "hidden";
				    if (confirm("<%=LitAdvertencia1%><%=NumRegistros%> <%=lcase(proceso)%><%=LitAdvertencia2%>")) {
				        document.getElementById("waitBoxOcultoTexto").innerHTML = "<%=LitMsgImportando%>";
				        document.getElementById("waitBoxOculto").style.visibility = "visible";
				        document.importar_datos.action = "importar_datos.asp?mode=CON&ncliente=<%=enc.EncodeForJavascript(null_s(CodEmpresa))%>&p=<%=enc.EncodeForJavascript(null_s(proceso))%>&f=<%=enc.EncodeForJavascript(null_s(formato))%>";
				        document.importar_datos.submit();
				    } else {
				        document.location = "importar_datos.asp?mode=select&ncliente=<%=enc.EncodeForJavascript(CodEmpresa)%>&borra=SI&s=<%=enc.EncodeForJavascript(suministros)%>&tdocumento=<%=enc.EncodeForJavascript(tituloDoc)%>";
				        parent.botones.document.location = "importar_datos_bt.asp?mode=select";
				    }
				</script><%
			else 'FORMATO INCORRECTO. EL MENSAJE DE ERROR SE GRABO EN LA TABLA TEMPORAL DEL USUARIO
				rst.close
				'on error goto 0
				set conn = nothing
				set rst=nothing

			end if
		elseif formato="ACC" then
			on error resume next
			conn.open dsnCliente
			conn.CommandTimeout = 0
			set rst = conn.execute(strselect)
			set conn = nothing
			if err.number<>0 then
				Mensaje=litErrImportACC & err.description
			end if
			on error goto 0
		else 'DBF
		    ''ricardo 14-4-2010 se cambia el DSN
			''conn.open DSNImport 'session("dsn_cliente")
            if proceso="TARJETAS" then
                conn.open DSNILION
            else
			    conn.open dsnCliente
            end if
			'response.write(DSNImport)
			conn.CommandTimeout = 0
			on error resume next
			set rst = conn.execute("exec master..sp_dropserver 'IMPORT'")
			on error goto 0
			on error resume next
			set rst = conn.execute(strcheck)
			if err.number=0 then 'FORMATO CORRECTO
				on error goto 0
				NumRegistros=rst(1) & " "
				PintaComprobacion rst,proceso,formato
				rst.close
				set conn = nothing
				set rst=nothing
				%><input type="Hidden" name="qry" value="<%=enc.EncodeForHtmlAttribute(null_s(strselect))%>">
				<script language="JavaScript" type="text/javascript">
				    document.getElementById("waitBoxOculto").style.visibility = "hidden";
				    if (confirm("<%=LitAdvertencia1%><%=NumRegistros%> <%=lcase(proceso)%><%=LitAdvertencia2%>")) {
				        document.getElementById("waitBoxOcultoTexto").innerHTML = "<%=LitMsgImportando%>";
				        document.getElementById("waitBoxOculto").style.visibility = "visible";
				        document.importar_datos.action = "importar_datos.asp?mode=CON&ncliente=<%=enc.EncodeForJavascript(CodEmpresa)%>&p=<%=enc.EncodeForJavascript(proceso)%>&f=<%=enc.EncodeForJavascript(formato)%>&tdocumento=<%=enc.EncodeForJavascript(tituloDoc)%>";
				        document.importar_datos.submit();
				    } else {
				        document.location = "importar_datos.asp?mode=select&ncliente=<%=enc.EncodeForJavascript(CodEmpresa)%>&borra=SI&s=<%=enc.EncodeForJavascript(suministros)%>&tdocumento=<%=enc.EncodeForJavascript(tituloDoc)%>";
				        parent.botones.document.location = "importar_datos_bt.asp?mode=select";
				    }
				</script><%
			else
				'Formato incorrecto
				rst.close
			end if
			on error goto 0
			set conn = nothing
			set rst=nothing
		end if
		set rstMsg = Server.CreateObject("ADODB.RECORDSET")
		'rstMsg.open "select mensaje from [" & session("usuario") & "] order by id desc",session("dsn_cliente")
		if formato ="ACC" then
		    rstMsg.open "select mensaje from [dbo].[" & session("usuario") & "] order by id desc",dsnCliente
		else
	       rstMsg.open "select mensaje from [" & session("usuario") & "] order by id desc",dsnTmpTable
          
		end if
		if not rstMsg.eof then
            if proceso="TARJETAS" then
                tmpMSg=split(rstMsg("mensaje"),";")
                Mensaje=Mensaje&GetLitErr(tmpMSg(0),tmpMsg(1))

            else
			    Mensaje=rstMsg("mensaje")
            end if
			if formato="ACC" then
				PintaResultados Mensaje,rst
				set rst=nothing
			end if
		else
			if formato="DBF" then
				Mensaje=LitErrImportDBF
				set conn = Server.CreateObject("ADODB.Connection")
	            ''ricardo 14-4-2010 se cambia el DSN
			    ''conn.open DSNImport
			    conn.open dsnCliente
				on error resume next
				set rst = conn.execute("exec master..sp_dropserver 'IMPORT'")
				set conn = nothing
				set rst = nothing
				on error goto 0
			end if
		end if
		rstMsg.close
		set rstMsg=Nothing
           
		if formato<>"ACC" then
			%><script language="JavaScript" type="text/javascript">
			      document.all("waitBoxOculto").style.visibility = "hidden";
			      alert("<%=Mensaje%>");
			      document.location = "importar_datos.asp?mode=select&ncliente=<%=enc.EncodeForJavascript(CodEmpresa)%>&borra=SI&s=<%=enc.EncodeForJavascript(suministros)%>&tdocumento=<%=enc.EncodeForJavascript(tituloDoc)%>";
			      parent.botones.document.location = "importar_datos_bt.asp?mode=select";
			</script><%
		end if
	elseif mode="CON" then 'SE CONFIRMA LA IMPORTACION

		RequestBin = Request.BinaryRead(Request.TotalBytes)		
  		Set UploadRequest = CreateObject("Scripting.Dictionary")
  		BuildUploadRequest RequestBin
		strselect=UploadRequest.Item("qry").Item("Value")

        if session("ncliente")<>SISTEMA_GESTION then
                dsnCliente = session("dsn_cliente")
        else
                strSelectDsnCliente = "select dsn from clientes with(nolock) where ncliente = ?"
                dsnCliente = DLookupP1(strSelectDsnCliente,CodEmpresa,adVarchar,5,DSNIlion&"")     
                'dsnCliente=d_lookup("dsn","clientes","ncliente='" & CodEmpresa & "'",DSNIlion)
        end if    
	    initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")
	    donde=inStr(1,DSNImport,"Initial Catalog=",1)
	    donde_fin=InStr(donde,DSNImport,";",1)
        if donde_fin=0 then
	        donde_fin=len(DSNImport)
	    end if
	    cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
	    dsnCliente=cadena_dsn_final		

		set conn = Server.CreateObject("ADODB.Connection")
		if formato="CSV" then
           if request.QueryString("p")="TARJETAS" then
                initial_catalogC=encontrar_datos_dsn(dsnIlion,"Initial Catalog=")
            else
                initial_catalogC=encontrar_datos_dsn(dsnCliente,"Initial Catalog=")
            end if
	        donde=inStr(1,DSNImport,"Initial Catalog=",1)
	        donde_fin=InStr(donde,DSNImport,";",1)
            if donde_fin=0 then
	            donde_fin=len(DSNImport)
	        end if
	        cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
	        dsnCliente=cadena_dsn_final

            if request.QueryString("p")="TARJETAS" then
		        conn.open dsnCliente''DSNILION
            else
                conn.open dsnCliente
            end if
			conn.CommandTimeout = 0
			set rst = conn.execute(strselect)
		elseif formato="DBF" then
            ''ricardo 14-4-2010 se cambia el DSN
		    ''conn.open DSNImport
		    conn.open dsnCliente
			conn.CommandTimeout = 0
			on error resume next
			set rst = conn.execute(strselect)
			on error goto 0
		end if

		set conn = nothing
		set rst=nothing

		'Se recoge el mensaje resultante
		set rstMsg = Server.CreateObject("ADODB.RECORDSET")
		'rstMsg.open "select mensaje from [" & session("usuario") & "] order by id desc",session("dsn_cliente")
        ''ricardo 14-4-2010 se cambia el DSN
		''rstMsg.open "select mensaje from [" & session("usuario") & "] order by id desc",DSNImport
            initial_catalogC=encontrar_datos_dsn(session("dsn_cliente"),"Initial Catalog=")
	        donde=inStr(1,DSNImport,"Initial Catalog=",1)
	        donde_fin=InStr(donde,DSNImport,";",1)
            if donde_fin=0 then
	            donde_fin=len(DSNImport)
	        end if
	        cadena_dsn_final=mid(DSNImport,1,donde-1) & "Initial Catalog=" & initial_catalogC & mid(DSNImport,donde_fin,len(DSNImport))
	        dsnCliente=cadena_dsn_final
		rstMsg.open "select mensaje from [" & session("usuario") & "] order by id desc",dsnCliente
		if not rstMsg.eof then
            if request.QueryString("p")="TARJETAS" then
                tmpMSg=split(rstMsg("mensaje"),";")
                if tmpMSg(0)="0" then
                    Mensaje=GetLitErr(tmpMSg(0),tmpMsg(1))
                else
                    while not rstMsg.eof
                        tmpMSg=split(rstMsg("mensaje"),";")
                        Mensaje=Mensaje&GetLitErr(tmpMSg(0),"")&": "&tmpMSg(1)&"\n"
                        rstMsg.movenext
                    wend
                end if
                'response.write("topo")
                'response.write (tmpMSg(0)&"-"&tmpMSg(1))
                'response.Write(Mensaje)
                'response.End
            else
			    Mensaje=rstMsg("mensaje")
            end if
		else
			if formato="DBF" then
				Mensaje=LitErrImportDBF
				set conn = Server.CreateObject("ADODB.Connection")
	            ''ricardo 14-4-2010 se cambia el DSN
			    ''conn.open DSNImport
			    conn.open dsnCliente
				on error resume next
				set rst = conn.execute("exec master..sp_dropserver 'IMPORT'")
				set conn = nothing
				set rst = nothing
				on error goto 0
			end if
		end if
		rstMsg.close
		strdrop ="if exists (select * from sysobjects where id = object_id('[dbo].[" & session("usuario") & "]') and sysstat & 0xf = 3) drop table [dbo].[" & session("usuario") & "]"
		'rstMsg.open strdrop,session("dsn_cliente"),adUseClient,adLockReadOnly
		if formato ="ACC" then
		    rstMsg.open strdrop,dsnCliente,adUseClient,adLockReadOnly
		else
            ''ricardo 14-4-2010 se cambia el DSN
    		''rstMsg.open strdrop,DSNImport,adUseClient,adLockReadOnly
    		rstMsg.open strdrop,dsnCliente,adUseClient,adLockReadOnly
        end if
		set rstMsg=nothing

		%><script language="JavaScript" type="text/javascript">
		      //alert("hola1");
		      alert("<%=Mensaje%>");
		      //alert("adios1");
		      document.location = "importar_datos.asp?mode=select&ncliente=<%=enc.EncodeForJavascript(CodEmpresa)%>&borra=SI&s=<%=enc.EncodeForJavascript(suministros)%>&tdocumento=<%=enc.EncodeForJavascript(tituloDoc)%>";
		      parent.botones.document.location = "importar_datos_bt.asp?mode=select";
		</script><%
	end if%>
</form>
<%end if%>
</body>
</HTML>