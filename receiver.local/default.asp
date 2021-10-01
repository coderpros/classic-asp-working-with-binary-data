<% Option Explicit %>
<!--#include file="includes/clsUpload.asp"-->
<%
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        Dim o: Set o = New clsUpload
        
        If o.Count > 0 Then        
            ' Save file to uploads directory.
            o("file").SaveAs(Server.MapPath("./Uploads") & "\" & o("file").FileName)

            Response.Write o("file").FileName
        End If
    End If
%>