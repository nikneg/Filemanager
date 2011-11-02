<!-- #include file="filemanager.class.asp" -->
<%
'ASP Connector for simogeo's Filemanager (http://github.com/simogeo/Filemanager/archives/master)
'It's implemented to work with ckeditor only. Must adapt to work with other editors.
'Developed by Matheus Fraguas at Soft Seven Internet (http://www.seven.com.br)
'The filemanager class requires these components:
'Scripting.FileSystemObject used to acess the filesystem
'ADODB.Stream used to serve a file to the browser (download)
'Dundas.Upload.2 used when uploading a file to the server
'GflAx.GflAx used to get an image's dimensions
'Last update 2011-01-20

'************************** Nicola Negrelli Update 2011-10-31 **********************************
'Feauture: No external plugins (Dundas.Upload.2, GflAx.GflAx) to get image dimensions and 
'          upload function thanks to Lewis Moten script
'          * Sometimes can't read right dimensions (-1x-1) specially if you use Photoshop like 
'          * image editing software
'          * Try open and resave image with other image editing software like ImageViewer	
'          JSON class to build rigth sintax json objects (http://code.google.com/p/aspjson/)
'Configuration: 1)download simogeo filemanager zip file (es.simogeo-Filemanager-5255a33.zip)
'               2)unzip file
'               3)rename wrapper folder (es.'simogeo-Filemanager-5255a33') to 'filemanager'
'               4)copy this folder into the root of your web site 
' 		          * If you copy this folder into other subfolder of your website
'                 * add the path to 'userPath' and 'fileIconsPath' variables in filemanager.config.asp file
'                 * (es. wwwroot/subfolder/filemanager -> /subfolder/filemanager/userfiles/ and 
'                 *  /subfolder/filemanager/images/fileicons/)
'               5)in filemanager.config.js set lang variable to 'asp'
'               6)enjoy	
'***********************************************************************************************


Dim mode, FileManager, Json

Json = ""

Response.ContentType = "application/json"
Response.Charset = "ISO-8859-1"


'If Len(Session("codSite")) = 0 Then  '******** Here your ahthorization script ***********
'	showErrorMessage("Your session has expired. Please login again.")
'End If

Set FileManager = New cFileManager

If Len(Request.Querystring("mode")) > 0 Then
	mode = Request.Querystring("mode")
	Select Case(lCase(mode))
		Case "getinfo":
			path = getPath("path") 
			Json = FileManager.GetInfo(path)
		Case "getfolder":
			path = getPath("path") 
			Json = FileManager.GetFolder(path)
		Case "rename":
			oldName = getPath("old") 
			newName = Trim(Request("new")) 		
			Json =  FileManager.Rename(oldName, newName)
		Case "delete":
			path = getPath("path") 
			Json =  FileManager.Delete(path)
		Case "addfolder":
			path = getPath("path") 
			name = Trim(Request("name")) 
			Json =  FileManager.AddFolder(path, name)
		Case "download":
			path = getPath("path") 
			FileManager.Download(path)
		Case Else:
			Json = FileManager.ErrorMessage("Mode Error")
	End Select
Else
	If Left(lCase(Request.ServerVariables("HTTP_CONTENT_TYPE")),20) = "multipart/form-data;" Then	
		Response.ContentType = "text/html"
		Response.Write "<textarea>"
		Json = FileManager.Add()
		Response.Write Json 
		Response.Write "</textarea>"	
		Response.End
	End If
End If
				
Response.Write Json 

Set FileManager = Nothing

Function getPath(name)
	Dim path, Json
	path = Request(name)
	path = Trim(path)
	path = Replace(path,"\","/")
	If inStr(path,"../") > 0 Or Left(path,Len(userPath)) <> userPath Then
		Json = FileManager.ErrorMessage("Invalid path")
		Response.Write Json 
		Response.End
	End If
	getPath = path
End Function
%>