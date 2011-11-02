<!-- #include file="filemanager.config.asp" -->
<!-- #include file="class/JSON_2.0.4.asp" -->
<!-- #include file="class/clsUpload.asp"-->
<!-- #include file="class/clsImage.asp"-->
<%
Class cFileManager

	Private userPath
	Private fs
	Private objStream
	Private objImage
	Private objJson 
	Private objUpload

	Private Sub Class_Initialize()
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		Set objStream = Server.CreateObject("ADODB.Stream")
		Set objJson = jsObject()
		userPath = ""
		If enableImageHandle Then
			Set objImage = New clsImage
		End If
	End Sub
	
	Private Sub Class_Terminate()
		Set fs = Nothing
		Set objStream = Nothing
		Set objJson = Nothing		
		If enableImageHandle Then
			Set objImage = Nothing
		End If
	End Sub

	Private Function isFolder(path)
		isFolder = (Right(path,1) = "/")
	End Function
	
	Private Sub returnError(message)
		objJson("Code") = -1
		objJson("Error") = message
	End Sub

	Private Function getImageProp(path, byRef width, byRef height)
		getImageProp = False
		If enableImageHandle Then
			objImage.Read(Server.MapPath(path))
			width = objImage.Width
			height = objImage.Height
			getImageProp = True
		End If
	End Function
	
	Private Function isImageExt(ext)
		For x = 0 To uBound(imgExtensions)
			If ext = imgExtensions(x) Then
				isImageExt = True
				Exit Function
			End If
		Next
		isImageExt = False
	End Function
	
	Private Function getPreviews(ext, path)
		Dim preview
		If isImageExt(ext) Then
			preview = userPath & path
		Else
			Select Case ext
				Case "txt", "rtf": preview = "txt"
				Case "zip": preview = "zip"
				Case "doc", "docx": preview = "doc"
				Case "xls", "xlsx", "docx": preview = "xls"
				Case "pdf": preview = "pdf"
				Case "swf": preview = "swf"
				Case "htm", "html": preview = "htm"
				Case "wav", "mp3", "wma", "mid": preview = "other_music"
				Case "avi", "mpg", "mpeg", "wmv", "mp4", "mov", "swf": preview = "other_movie"
				Case Else: preview = "default"
			End Select
			preview = fileIconsPath & preview & ".png"
		End If
		getPreviews = preview
	End Function	

	Private Sub getFileInfoByFolder(path, file)
		Dim fileExt, width, height
		fileExt = lCase(Split(file,".")(uBound(Split(file,"."))))
		Set objJson(path) =  jsObject() 		
		objJson(path)("Path") =  path 
		objJson(path)("Filename") = file.Name
		objJson(path)("File Type") = fileExt
		objJson(path)("Preview") = getPreviews(fileExt, path)
			Set objJson(path)("Properties")	= jsObject()
			objJson(path)("Properties")("Date Created") = file.DateCreated
			objJson(path)("Properties")("Date Modified") = file.DateLastModified
			If isImageExt(fileExt) And enableImageHandle Then
				If getImageProp(path, width, height) Then
					objJson(path)("Properties")("Height") = height
					objJson(path)("Properties")("Width") =  width 
				End If
			End If
			objJson(path)("Properties")("Size") = file.Size
		objJson(path)("Error") = ""
		objJson(path)("Code") = 0
	End Sub
		
	Private Sub	getFileInfoByInfo(path, file)
		Dim fileExt, width, height
		fileExt = lCase(Split(file,".")(uBound(Split(file,"."))))
		objJson("Path") =  path 
		objJson("Filename") = file.Name
		objJson("File Type") = fileExt
		objJson("Preview") = getPreviews(fileExt, path)
			Set objJson("Properties")	= jsObject()
			objJson("Properties")("Date Created") = file.DateCreated
			objJson("Properties")("Date Modified") = file.DateLastModified
			If isImageExt(fileExt) And enableImageHandle Then
				If getImageProp(path, width, height) Then
					objJson("Properties")("Height") = height
					objJson("Properties")("Width") =  width 
				End If
			End If
			objJson("Properties")("Size") = file.Size
		objJson("Error") = ""
		objJson("Code") = 0
	End Sub
	
	Private Sub getFolderInfoByFolder(path, folder)
		Set objJson(path) =  jsObject() 
		objJson(path)("Path") =  path 
		objJson(path)("Filename") = folder.Name
		objJson(path)("File Type") = "dir"
		objJson(path)("Preview") = fileIconsPath & "_Close.png"
			Set objJson(path)("Properties")	= jsObject()
			objJson(path)("Properties")("Date Created") = folder.DateCreated
			objJson(path)("Properties")("Date Modified")= folder.DateLastModified
			objJson(path)("Properties")("Size") = folder.Size
		objJson(path)("Error") = ""
		objJson(path)("Code") = 0
	End Sub
		
	Private Sub	getFolderInfoByInfo(path, folder)
		objJson("Path") =  path 
		objJson("Filename") = folder.Name
		objJson("File Type") = "dir"
		objJson("Preview") = fileIconsPath & "_Close.png"
			Set objJson("Properties")	= jsObject()
			objJson("Properties")("Date Created") = folder.DateCreated
			objJson("Properties")("Date Modified")= folder.DateLastModified
			objJson("Properties")("Size") = folder.Size
		objJson("Error") = ""
		objJson("Code") = 0				
	End Sub

	Public Function GetInfo(path)
		Dim file
		On Error Resume Next
		If isFolder(path) Then
			Set file = fs.GetFolder(Server.MapPath(userPath + path))
		Else
			Set file = fs.GetFile(Server.MapPath(userPath + path))
		End If
		If Err.Number <> 0 Then
			Call returnError ("Can't open folder or path")
		Else
			If isFolder(path) Then
				Call getFolderInfoByInfo(path, file)
			Else
				Call getFileInfoByInfo(path, file)
			End If
		End If
		On Error Goto 0
		GetInfo = objJson.Flush
		Set file = Nothing
	End Function

	Public Function GetFolder(path)
		Dim folder
		On Error Resume Next	
		Set folder = fs.GetFolder(Server.MapPath(userPath + path))
		If Err.Number <> 0 Then
			Call returnError ("Can't open folder")
		Else
			For Each item in folder.subfolders
				Call getFolderInfoByFolder(path & item.name & "/" , item)			
			Next
			For Each item in folder.files
				Call getFileInfoByFolder(path & item.name, item)			
			Next		
		End If
		On Error Goto 0
		GetFolder = objJson.Flush
	End Function

	Public Function AddFolder(path, name)
		Dim newPath
		If inStr(name,"/") > 0 Or inStr(name,"\") > 0 Then
			Call returnError("Invalid name.")
		Else	
			newPath = Server.MapPath(userPath & path & name)
			If fs.FolderExists(newPath) Then
				Call returnError("Folder already exists")
			Else
				On Error Resume Next
				fs.CreateFolder(newPath)
				If Err.Number <> 0 Then
					Call returnError("Can't create folder")
				Else
					Call AddFolderJSONobj(path, name)
				End If
				On Error Goto 0
			End if
		End If	
		AddFolder = objJson.Flush
	End Function
	
	Private Sub AddFolderJSONobj(path, name)
		objJson("Parent") = path
		objJson("Name") = name
		objJson("Error") = ""
		objJson("Code") = 0
	End Sub

	Public Function Rename(oldName, newName)
		Dim item, arrPath, originalName
		If inStr(newName,"/") > 0 Or inStr(newName,"\") > 0 Then
			Call returnError("Invalid name.")
		Else	
			arrPath = Split(oldName,"/")
			On Error Resume Next
			If isFolder(oldName) Then
				Set item = fs.GetFolder(Server.MapPath(userPath + oldName))
				ReDim Preserve arrPath(uBound(arrPath)-2)
			Else
				Set item = fs.GetFile(Server.MapPath(userPath + oldName))
				ReDim Preserve arrPath(uBound(arrPath)-1)
			End If
			originalName = item.name
			item.Move(item.ParentFolder.Path & "\" & newName)
			If Err.Number <> 0 Then
				Call returnError("Can't rename folder or file")
			Else
				Call RenameJSONobj(oldName, originalName, arrPath, newName)
			End If
			On Error Goto 0
		End If
		Rename = objJson.Flush
	End Function

	Private Sub RenameJSONobj(oldName, originalName, arrPath, newName)
		objJson("Error") = ""
		objJson("Code") = 0
		objJson("Old Path") = oldName
		objJson("Old Name") = originalName
		objJson("New Path") = Join(arrPath,"/") & "/" & newName & "/"
		objJson("New Name") = newName
	End Sub

	Public Function Add()
		Dim path, fileName
		Set objUpload = new clsUpload
		mode = objUpload.Fields("mode").Value
		Select Case(lCase(mode))
			Case "add":
				fileName = objUpload.Fields("newfile").FileName
				path = objUpload.Fields("currentpath").Value
				On Error Resume Next
				objUpload.Fields("newfile").SaveAs(Server.MapPath(userPath & path & fileName))
				If Err.Number <> 0 Then
					Call returnError ("Can't save file")
				Else	
					Call addJSONobj(path, fileName)
				End If
				On Error Goto 0
			Case Else:
				Call returnError("Mode Error")
		End Select
		Add = objJson.Flush 
		Set objUpload = Nothing
	End Function
	
	Private Sub addJSONobj(path, fileName)
		objJson("Path") = path
		objJson("Name") = fileName
		objJson("Error") = ""
		objJson("Code") = "0"
	End Sub			

	Public Function Delete(path)
		Dim item
		On Error Resume Next
		If isFolder(path) Then
			Set item = fs.GetFolder(Server.MapPath(userPath + path))
		Else
			Set item = fs.GetFile(Server.MapPath(userPath + path))
		End If
		item.Delete(True)
		If Err.Number <> 0 Then
			Call returnError("Can't remove folder or file")
		Else
			Call DeleteJSONobj(path)
		End If
		On Error Goto 0
		Delete = objJson.Flush
	End Function
	
	Private Sub  DeleteJSONobj(path)
		objJson("Error") = ""
		objJson("Code") = 0
		objJson("Path") = path
	End Sub
	
	Public Sub Download(path)
		Dim item
    	Server.ScriptTimeout = 30000
		Response.Clear
		Response.ContentType = "application/x-download"
		Set item = fs.GetFile(Server.MapPath(userPath & path))
		Response.AddHeader "Content-Disposition", "attachment; filename=" & item.Name
'		Response.AddHeader "Content-Length", item.Size
		Set item = Nothing
		Response.Buffer = False 
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		objStream.Type = 1
		objStream.LoadFromFile(Server.MapPath(userPath + path))
		Response.BinaryWrite(objStream.Read)
		objStream.Close
		Response.End
	End Sub
	
	Public Function ErrorMessage(message)
		Call returnError(message)
		ErrorMessage = objJson.Flush
	End Function
	
End Class
%>