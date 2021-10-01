<%
  Response.ContentType = "text/json"
  Response.ContentType = "application/json"
  Response.expires = 0
  Response.expiresabsolute = Now() - 1
%>

<%

	'NOTE - YOU MUST HAVE VBSCRIPT v5.0 INSTALLED ON YOUR WEB SERVER
	'	   FOR THIS LIBRARY TO FUNCTION CORRECTLY. YOU CAN OBTAIN IT
	'	   FREE FROM MICROSOFT WHEN YOU INSTALL INTERNET EXPLORER 5.0
	'	   OR LATER.

	Dim adTypeBinary, adTypeText, adModeReadWrite, adSaveCreateOverWrite
	adTypeBinary  = 1
	adTypeText = 2
	adModeReadWrite = 3
    adSaveCreateOverWrite = 2

	Dim x_FOLDER, x_FILENAME, x_TIPO

	URL_FILENET_UPLOAD = "http://service-test"

	x_FOLDER = Request.QueryString("FOLDER")
	x_FILENAME = Request.QueryString("FILENAME")
	x_TIPO = Request.QueryString("TIPO")

	If x_TIPO = "" Then
		x_TIPO = "image/png"
	End If

	response_filenet = publishFileNET(x_FOLDER, x_FILENAME)

	Response.Write "{ filename : '"  & x_FILENAME & "', success : '" & response_filenet & "'}"


	

    '******************************************
    '*** Function StringToBinary		    ***
    '*** - Transforma texto en binario 	    ***
    '******************************************
	Function StringToBinary(input)
		dim stream
		set stream = Server.CreateObject("ADODB.Stream")
		stream.Charset = "UTF-8"
		stream.Type = adTypeText 
		stream.Mode = adModeReadWrite 
		stream.Open
		stream.WriteText input
		stream.Position = 0
		stream.Type = adTypeBinary 
		StringToBinary = stream.Read
		stream.Close
		set stream = Nothing
	End Function

	Function stringToByte(toConv)

		Dim i, tempChar

		 For i = 1 to Len(toConv)
			tempChar = Mid(toConv, i, 1)
			stringToByte = stringToByte & chrB(AscB(tempChar))
		 Next
		 
	End Function
	
	Function BinaryToString(Binary)
	  'Antonin Foller, http://www.motobit.com
	  'Optimized version of a simple BinaryToString algorithm.
	  
	  Dim cl1, cl2, cl3, pl1, pl2, pl3
	  Dim L
	  cl1 = 1
	  cl2 = 1
	  cl3 = 1
	  L = LenB(Binary)
	  
	  Do While cl1<=L
		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1
		If cl3>300 Then
		  pl2 = pl2 & pl3
		  pl3 = ""
		  cl3 = 1
		  cl2 = cl2 + 1
		  If cl2>200 Then
			pl1 = pl1 & pl2
			pl2 = ""
			cl2 = 1
		  End If
		End If
	  Loop
	  BinaryToString = pl1 & pl2 & pl3
	End Function

    '******************************************
    '*** Function ReadBinaryFile		    ***
    '*** - Lee un archivo en binario 	    ***
    '******************************************
	Function ReadBinaryFile(fullFilePath) 
		Trace("upload_call.asp ReadBinaryFile " & fullFilePath)
		dim stream
		set stream = Server.CreateObject("ADODB.Stream")
		stream.Type = 1
		stream.Open()
		stream.LoadFromFile(fullFilePath)
		ReadBinaryFile = stream.Read()
		stream.Close
		set stream = nothing
	end function 

	private Sub writeBytes(file, bytes)
	  Dim binaryStream
	  Set binaryStream = CreateObject("ADODB.Stream")
	  binaryStream.Type = adTypeBinary
	  'Open the stream and write binary data
	  binaryStream.Open
	  binaryStream.Write bytes
	  'Save binary data to disk
	  binaryStream.SaveToFile file, adSaveCreateOverWrite
	End Sub

	private function encodeBase64(bytes)
	  dim DM, EL
	  Set DM = CreateObject("Microsoft.XMLDOM")
	  ' Create temporary node with Base64 data type
	  Set EL = DM.createElement("tmp")
	  EL.DataType = "bin.base64"
	  ' Set bytes, get encoded String
	  EL.NodeTypedValue = bytes
	  encodeBase64 = EL.Text
	end function
	  
	private function decodeBase64(base64)
	  dim DM, EL
	  Set DM = CreateObject("Microsoft.XMLDOM")
	  ' Create temporary node with Base64 data type
	  Set EL = DM.createElement("tmp")
	  EL.DataType = "bin.base64"
	  ' Set encoded String, get bytes
	  EL.Text = base64
	  decodeBase64 = EL.NodeTypedValue
	end function

    '******************************************
    '*** Function publishFileNET		    ***
    '******************************************
	Function publishFileNET(folderRoot, path_filename) 

		Dim BINARYPOST, binaryFile
		Dim boundary
		boundary = "--------------010201080703010401070605"

		binaryFile = ReadBinaryFile(folderRoot + path_filename)

		Set objXMLhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
		On Error Resume Next
			objXMLhttp.open "POST", URL_FILENET_UPLOAD, false
			objXMLhttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
			objXMLhttp.setRequestHeader "Authorization", "7343643678463748"
			objXMLhttp.setRequestHeader "User-Agent", "SOP/Custom"

			' Boundary inicial
			BINARYPOST = (vbCrLf & "--" & boundary & vbCrLf)
			' 01 - className
			BINARYPOST = BINARYPOST & ("" & _ 
			"Content-Disposition: form-data; name=""className""" & vbCrLf & vbCrLf & "Adquiriente" & vbCrLf)

			' 02 - properties
			BINARYPOST = BINARYPOST & ("--" & boundary & vbCrLf)
			BINARYPOST = BINARYPOST & ("" & _ 
			"Content-Disposition: form-data; name=""properties""" & vbCrLf & vbCrLf & "{data:'value'}" & vbCrLf)

			' 03 - file
			BINARYPOST = BINARYPOST & ("--" & boundary & vbCrLf)
			BINARYPOST = BINARYPOST & ("" & _ 
			"Content-Disposition: form-data; name=""file""; filename=""" & Replace(Replace(folderRoot, "\", "/"), "//", "/") + path_filename & """" & _
			 vbCrLf & "Content-Type: " & x_TIPO & vbCrLf & vbCrLf)

			objXMLhttp.send stringToByte(BINARYPOST)	' 1- send HTTP

			objXMLhttp.send binaryFile					' 2- send HTTP

			objXMLhttp.send stringToByte(vbCrLf & "--" & boundary & "--"  & vbCrLf)	' 3- send HTTP

			strResponse = objXMLhttp.responseText
			Set objXMLhttp = Nothing

			' Se retorna la URL para actualizar en la base de datos de imágenes
			publishFileNET = strResponse

		If Err.Number then
			' Custom code
			publishFileNET = ""
		End If
		On Error Goto 0

	End Function

%>