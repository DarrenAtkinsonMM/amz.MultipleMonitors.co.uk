<%
Dim HaveImgUplResizeObjs
Dim HaveImgGifSupport, HaveImgPngSupport, HaveImgCropSupport
Dim pcv_UploadObj
Dim	pcv_ResizeObj
Dim CUmsg
CUmsg=""

HaveImgUplResizeObjs=0
HaveImgGifSupport=0
HaveImgPngSupport=0
HaveImgCropSupport=0
pcv_UploadObj=0
pcv_ResizeObj=0

Function IsObjInstalled(strClassString)
    Dim tmpResult,uploadpath,imagepath
	On Error Resume Next
	' initialize default values
	tmpResult = False
	Err = 0
	' testing code
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then tmpResult = True
	' cleanup
	Set xTestObj = Nothing
	Err = 0
	'Installed, check active
	IF tmpResult = True THEN
		if Ucase(strClassString)="PERSITS.UPLOAD" then
			Set Upload = Server.CreateObject("Persits.Upload")
			
			If PageUpload<>"1" then
			
			Upload.IgnoreNoPost=True
			
			If PPD="1" then
				Upload.SaveVirtual "\"&scPcFolder&"\includes\uploadresize\"
			else
				if checkSubFolder="1" then
					Upload.SaveVirtual "..\..\includes\uploadresize\"
				else
					Upload.SaveVirtual "..\includes\uploadresize\"
				end if
			end if
			
			End if
		
			If err.number <> 0 Then
				CUmsg = CUmsg & "<li>The component: ASPUpload (Persits.Upload) was installed but cannot be used. <br><b>Error:</b> " & err.description & "</li>"
				tmpResult=false
				err.number=0
				err.description=0
			End if
			
			Set Upload=nothing
		end if
		
		if PPD="1" then
			uploadpath=Server.Mappath("\"&scPcFolder&"\includes\uploadresize\")
		else
			if checkSubFolder="1" then
				uploadpath=Server.Mappath("..\..\includes\uploadresize\")
			else
				uploadpath=Server.Mappath("..\includes\uploadresize\")
			end if
		end if
		uploadpath = uploadpath & "\"
		
		IF Ucase(strClassString)=Ucase("SoftArtisans.FileUp") THEN
			Set Upload = Server.CreateObject("SoftArtisans.FileUp")
	
			Upload.Path = uploadpath
			
			If err.number <> 0 Then
				CUmsg = CUmsg & "<li>The component: FileUp (SoftArtisans.FileUp) was installed but cannot be used. <br><b>Error:</b> " & err.description & "</li>"
				tmpResult=false
				err.number=0
				err.description=0
			End if
			
			Set Upload=nothing
		
		END IF
		
		IF Ucase(strClassString)=Ucase("aspSmartUpload.SmartUpload") THEN
			Set SmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

			If err.number <> 0 Then
				CUmsg = CUmsg & "<li>The component: ASPSmartUpload (aspSmartUpload.SmartUpload) was installed but cannot be used. <br><b>Error:</b> " & err.description & "</li>"
				tmpResult=false
				err.number=0
				err.description=0
			End if
			
			Set SmartUpload=nothing
		
		END IF
		
		if PPD="1" then
			imagepath=Server.Mappath ("/" & scAdminFolderName & "/images/")
		else
			if checkSubFolder="1" then
				imagepath=Server.Mappath ("../images/")
			else
				imagepath=Server.Mappath ("images/")
			end if
		end if
		imagepath = imagepath & "/"
		
		IF Ucase(strClassString)=Ucase("Persits.Jpeg") THEN
			FileName=imagepath & "pcIconStart.jpg"
			Set Jpeg = Server.CreateObject("Persits.Jpeg")
			Jpeg.Open FileName
			
			If err.number <> 0 Then
				CUmsg = CUmsg & "<li>The component: ASPJpeg (Persits.Jpeg) was installed but cannot be used.<br><b>Error:</b> " & err.description & "</li>"
				tmpResult=false
				err.number=0
				err.description=0
			End if
			
			Set Jpeg=nothing
		
		END IF
		
		IF Ucase(strClassString)=Ucase("AspImage.Image") THEN
			FileName=imagepath & "pcIconStart.jpg"
			Set Image = Server.CreateObject("AspImage.Image")
			Image.LoadImage(FileName)
			
			If err.number <> 0 Then
				CUmsg = CUmsg & "<li>The component: ASPJpeg (Persits.Jpeg) was installed but cannot be used. <br><b>Error:</b> " & err.description & "</li>"
				tmpResult=false
				err.number=0
				err.description=0
			End if
			
			Set Image=nothing
		
		END IF

	END IF
	IsObjInstalled=tmpResult
End Function

Function GetObjectVersion(strClassString)
	On Error Resume Next
	GetObjectVersion = "0"
	Err = 0
	
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	
	GetObjectVersion = xTestObj.Version
	Set xTestObj = Nothing
	Err = 0
End Function

Function GetImageDimensions(FileName, outWidth, outHeight)
	Select Case pcv_ResizeObj
	Case 1:
		Set Jpeg = Server.CreateObject("Persits.Jpeg")
		Jpeg.Open FileName
		
		outWidth = jpeg.OriginalWidth
		outHeight = jpeg.OriginalHeight
	
		Jpeg.Close
	Case 2:
		Set Image = Server.CreateObject("AspImage.Image")
		Image.LoadImage(FileName)
	
		outWidth = Image.MaxX
		outHeight = Image.MaxY
		Set Image = Nothing
	End Select
End Function

Function ValidateImageType(FileName, ImageType)
	validateErr = ""
	
	'// First validate if it's allowed
	If Not IsUploadAllowed(FileName) Then
		validateErr = "You have attempted to upload an image that has a filename or image type that is not allowed. Please check your image to ensure it has the correct file name and type and try again."
		ValidateImageType = validateErr
		Exit Function
	End If
	
	'// Then make sure the image type is supported
	If UCase(ImageType) <> "JPG" And UCase(ImageType) <> "PNG" And UCase(ImageType) <> "GIF" Then
		validateErr = "The type of the uploaded image is unknown or not supported by the <strong>" & GetResizeComponentName() & "</strong> component. Please use a supported image format for uploading images.<br/><br/>"
		validateErr = validateErr & "Supported image formats: "
		validateErr = validateErr & "<strong>"
		validateErr = validateErr & "JPG"
		If HaveImgPngSupport Then validateErr = validateErr & ", PNG"
		If HaveImgGifSupport Then validateErr = validateErr & ", GIF"
		validateErr = validateErr & "</strong>"
		
		ValidateImageType = validateErr
		Exit Function
	End If
	
	If UCase(ImageType) = "PNG" And (HaveImgPngSupport = 0 And pcv_UploadObj > 0) Then
		validateErr = "The installed version of the <strong>" & GetResizeComponentName() & "</strong> component does not support PNG files. Please use a different format for slideshow images."
	End If
	
	If UCase(ImageType) = "GIF" And (HaveImgGifSupport = 0  And pcv_UploadObj > 0) Then
		validateErr = "The installed version of the <strong>" & GetResizeComponentName() & "</strong> component does not support GIF files. Please use a different format for slideshow images."
	End If
	
	ValidateImageType = validateErr
End Function
	
Function GetUploadComponentName()
	Select Case pcv_UploadObj
	Case 1: GetUploadComponentName = "FileUp (SoftArtisans.FileUp)"
	Case 2:	GetUploadComponentName = "ASPUpload (Persits.Upload)"
	Case 3:	GetUploadComponentName = "ASPSmartUpload (aspSmartUpload.SmartUpload)"
	Case Else: GetUploadComponentName = "Base Uploader"
	End Select
End Function

Function GetResizeComponentName()
	Select Case pcv_ResizeObj
	Case 1: GetResizeComponentName = "ASPJpeg (Persits.Jpeg)"
	Case 2:	GetResizeComponentName = "AspImage (AspImage.Image)"
	End Select
End Function

if IsObjInstalled("SoftArtisans.FileUp") then
	HaveImgUplResizeObjs=1
	pcv_UploadObj=1
else
	if IsObjInstalled("Persits.Upload") then
		HaveImgUplResizeObjs=1
		pcv_UploadObj=2
	else
		if IsObjInstalled("aspSmartUpload.SmartUpload") then
			HaveImgUplResizeObjs=1
			pcv_UploadObj=3
		else
			HaveImgUplResizeObjs=0
			pcv_UploadObj=4
		end if
	end if
end if

if HaveImgUplResizeObjs=1 OR pcv_UploadObj=4 then
	If IsObjInstalled("Persits.Jpeg") then
		objVersion = GetObjectVersion("Persits.Jpeg")
		HaveImgUplResizeObjs=2
		pcv_ResizeObj=1
	else
		If IsObjInstalled("AspImage.Image") then
			objVersion = GetObjectVersion("AspImage.Image")
		
			HaveImgUplResizeObjs=2
			pcv_ResizeObj=2
		end if
	end if
end if

if HaveImgUplResizeObjs=2 then
	objVersionInt = Replace(Replace(objVersion, ".", ""), "(64-bit)", "")
	If Not IsNumeric(objVersionInt) Then
		objVersionInt = 0
	Else
		objVersionInt = CInt(objVersionInt)
	End If
	
	'Persits.Jpeg
	if pcv_ResizeObj=1 then
		
		'cropping support added in version 1.1+
		if objVersionInt >= 1100 then
			HaveImgCropSupport = 1
		end if
		
		'GIF support added in version 2.0+
		if objVersionInt >= 2000 then
			HaveImgGifSupport = 1
		end if
		
		'PNG support added in version 2.1+
		if objVersionInt >= 2100 then
			HaveImgPngSupport = 1
		end if
	'AspImage.Image
	elseif pcv_ResizeObj=2 then
		'cropping and PNG support added in 2.x+
		if objVersionInt >= 200 then
			HaveImgCropSupport = 1
			HaveImgGifSupport = 0
			HaveImgPngSupport = 1
		end if
	end if
end if

if HaveImgUplResizeObjs>=1 then
	HaveImgUplResizeObjs=1
else
	HaveImgUplResizeObjs=0
end if
%>