<%
	Option Explicit
	'NOTE - YOU MUST HAVE VBSCRIPT v5.0 INSTALLED ON YOUR WEB SERVER
	'	   FOR THIS LIBRARY TO FUNCTION CORRECTLY. YOU CAN OBTAIN IT
	'	   FREE FROM MICROSOFT WHEN YOU INSTALL INTERNET EXPLORER 5.0
	'	   OR LATER.

	Dim adTypeBinary, adTypeText, adModeReadWrite, adSaveCreateOverWrite
	Dim response_filenet

	adTypeBinary  = 1
	adTypeText = 2
	adModeReadWrite = 3
    adSaveCreateOverWrite = 2
	
	dim fs: set fs=Server.CreateObject("Scripting.FileSystemObject")

	Dim x_FOLDER: x_Folder = Request.QueryString("folder")
	Dim x_FILENAME: x_FILENAME = Request.QueryString("filename")
	Dim x_TIPO: x_TIPO = Request.QueryString("tipo")
	Dim URL_FILENET_UPLOAD: URL_FILENET_UPLOAD = "http://receiver.local"
	
	If x_FILENAME = "" Or x_FOLDER = "" Then
		Call Err.Raise(vbObjectError + 10, "Sender Web App", "Filename & folder variables are required.")
	End If

	If x_TIPO = "" Then
		x_TIPO = "image/png"
	End If

	response_filenet = publishFileNET(x_FOLDER, x_FILENAME)

	Response.Write "{ filename : '"  & x_FILENAME & "', success : '" & response_filenet & "'}"

    '******************************************
    '*** Function ReadBinaryFile		    ***
    '*** - Lee un archivo en binario 	    ***
    '******************************************
	Function ReadBinaryFile(FileName)
	  Const adTypeBinary = 1
  
	  'Create Stream object
	  Dim BinaryStream
	  Set BinaryStream = CreateObject("ADODB.Stream")

	  'Specify stream type - we want To get binary data.
	  BinaryStream.Type = adTypeBinary
  
	  'Open the stream
	  BinaryStream.Open
	
	  'Load the file data from disk To stream object
	  BinaryStream.LoadFromFile FileName
  
	  'Open the stream And get binary data from the object
	  ReadBinaryFile = BinaryStream.Read
	End Function

	Function StringToBinary(toConvert)
		Dim objStream, data

		Set objStream = Server.CreateObject("ADODB.Stream")

		objStream.Charset = "ISO-8859-1"
		objStream.Type = adTypeText '2
		objStream.Mode = adModeReadWrite '3
		objStream.Open
		objStream.WriteText toConvert

		objStream.Position = 0
		objStream.Type = adTypeBinary '1
		StringToBinary = objStream.Read

		objStream.Close
		Set objStream = Nothing
	End Function
    '******************************************
    '*** Function publishFileNET		    ***
    '******************************************
	Function publishFileNET(folderRoot, path_filename) 
		Dim sboundary: sBoundary = "--------------010201080703010401070605"
		
		Dim filePath: filePath = Server.MapPath(folderRoot) & "\" & path_filename
		
		If Not fs.FileExists(filePath) Then 
			Call Err.Raise(vbObjectError + 10, "Sender Web App", "File does not exist.")
		Else
			Dim binFile: binFile = ReadBinaryFile(filePath)
			Dim objHttp: Set objHttp = Server.Createobject("MSXml2.ServerXmlHttp.6.0")
			Dim objStream: Set objStream = Server.CreateObject("ADODB.Stream")
			Dim strRequestStart, strRequestEnd, strResponse
			Dim binPost

			strRequestStart = "--" & sBoundary& vbCrlf &_
			"Content-Disposition: form-data; name=""className""" & vbCrLf & vbCrLf & "Adquiriente" & vbCrLf & _
			vbCrlf &_
			"--" & sBoundary & vbCrlf &_
			"Content-Disposition: form-data; name=""file""; filename=""" & filePath & """" & vbCrlf &_
			"Content-Type: " & x_TIPO & vbCrlf &_
			vbCrlf

			strRequestEnd = vbCrLf & "--" & sBoundary & "--"

			objStream.Type = adTypeBinary '1
			objStream.Mode = adModeReadWrite '3
			objStream.Open
			objStream.Write StringToBinary(strRequestStart)
			objStream.Write binFile
			objStream.Write StringToBinary(strRequestEnd)
			objStream.Position = 0

			binPost = objStream.Read
			
			'Response.Write binPost

			objStream.Close	
			Set objStream = Nothing

			objHttp.open "POST", URL_FILENET_UPLOAD, false
			objHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & sBoundary
			objHttp.setRequestHeader "Authorization", "7343643678463748"
			objHttp.setRequestHeader "User-Agent", "SOP/Custom"

			objHttp.Send binPost

			strResponse = objHttp.responseText

			Set objHttp = Nothing

			'formData = "--" & boundary & vbCrLf & _
			'	"Content-Disposition: form-data; name=""file""; filename=""" & filePath & """" & vbCrLf & _
			'	"Content-Type: " & x_TIPO & vbCrLf & _
			'	vbCrLf & _
			'	binaryFile & vbCrLf & _
			'	"--" & boundary & "--"
			'	'"Content-Transfer-Encoding: base64" & vbCrLf & _

			' Boundary inicial
			'BINARYPOST = (vbCrLf & "--" & boundary & vbCrLf)
			' 01 - className
			'BINARYPOST = BINARYPOST & ("" & _ 
			'"Content-Disposition: form-data; name=""className""" & vbCrLf & vbCrLf & "Adquiriente" & vbCrLf)

			' 02 - properties
			'BINARYPOST = BINARYPOST & ("--" & boundary & vbCrLf)
			'BINARYPOST = BINARYPOST & ("" & _ 
			'"Content-Disposition: form-data; name=""properties""" & vbCrLf & vbCrLf & "{data:'value'}" & vbCrLf)

			' 03 - file
			'BINARYPOST = BINARYPOST & ("--" & boundary & vbCrLf)
			'BINARYPOST = BINARYPOST & ("" & _ 
			'"Content-Disposition: form-data; name=""file""; filename=""" & Replace(Replace(folderRoot, "\", "/"), "//", "/") + path_filename & """" & _
			'vbCrLf & "Content-Type: " & x_TIPO & vbCrLf & vbCrLf)

			'objXMLhttp.send stringToByte(BINARYPOST)	' 1- send HTTP

			'objXMLhttp.send binaryFile					' 2- send HTTP

			'objXMLhttp.send stringToByte(vbCrLf & "--" & boundary & "--"  & vbCrLf)	' 3- send HTTP

			' Se retorna la URL para actualizar en la base de datos de imágenes
			publishFileNET = strResponse

			If Err.Number then
				' Custom code
				publishFileNET = ""
			End If
			On Error Goto 0
		End If
	End Function
%>