<%@ LANGUAGE="VBSCRIPT" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../../includes/ppdstatus.inc"-->
<!--#include file="../../includes/productcartFolder.asp"-->
<%checkSubFolder="1"%>
<%PageUpload=1%>
<!--#include file="checkImgUplResizeObjs.asp"-->
<!--#include file="../../includes/pcSanitizeUpload.asp"-->
<!--#include file="clsUpload.asp"-->

<!DOCTYPE html>
<html>
<head>
 	<title>Upload Images</title>
  <link href="../css/pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-image: none;">
<% 'Check if objects exists
If HaveImgUplResizeObjs=0 AND pcv_UploadObj<>4 then %>
    <table class="pcCPcontent">
        <tr>
            <td>
            <div class="pcCPmessage">We are unable to find compatible Upload and/or Image Resize server components. Please consult the User Guide for detailed system requirements.</div>
            </td>
        </tr>
        <tr>
            <td align="center"><input type="button" class="btn btn-default"  name="Close" value=" Close window " onClick="javascript:window.close();"></td>
        </tr>
    </table>
	<% Response.End 'kill page
End If 
'if NO objects, kill the page %>

<%
Dim catalogpath, uploadpath, thumbfilename, generalfilename, detailfilename, thumbnailsize, generalsize, detailsize, sharpen, countfiles
Dim randomnum, FileName, BigBeforeWidth, BigBeforeHeight, BigAfterWidth, BigAfterHeight, imgcomp
Dim Image
Dim resizexy
Dim DidntResize

DidntResize=0

Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function

if PPD="1" then
	catalogpath=Server.Mappath ("\"&scPcFolder&"\pc\catalog\")
else
	catalogpath=Server.Mappath ("..\..\pc\catalog\")
end if
catalogpath = catalogpath & "\"

if PPD="1" then
	uploadpath=Server.Mappath ("\"&scPcFolder&"\includes\uploadresize\")
else
	uploadpath=Server.Mappath ("..\..\includes\uploadresize\")
end if
uploadpath = uploadpath & "\"

If (pcv_ResizeObj<>1) AND (pcv_ResizeObj<>2) then
	uploadpath=catalogpath
End if

'on error resume next

Function SetSessionVars(thumbnailsize, generalsize, detailsize)
	Session("prdThumbSize") = thumbnailsize
	Session("prdGeneralSize") = generalsize
	Session("prdDetailSize") = detailsize
End Function

Function UseSAFileUp()

	'--- Instantiate the FileUp object
	Set Upload = Server.CreateObject("SoftArtisans.FileUp")
	
	Upload.Path = uploadpath
	
	thumbnailsize = Upload.Form("thumbnailsize")
	generalsize = Upload.Form("generalsize")
	detailsize = Upload.Form("detailsize")
	sharpen = Upload.Form("sharpen")	
	resizexy = Upload.Form("resizexy")
	
	call SetSessionVars(thumbnailsize, generalsize, detailsize)
	
	If Upload.Form("file1").UserFilename = "" Then 	%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload any images.</b><br><br>
									<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        </table>
		<% Response.End
	End If
	%>	
	
	<%
		
		FileName = Upload.Form("file1").UserFilename
		ImageType = Right(Replace(UCase(FileName), ".JPEG", ".JPG"), 3)
		
		validateErrMsg = ValidateImageType(FileName, ImageType)
		
		If Len(validateErrMsg) > 0 then %>
			<table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
			<tr> 
				<td> 
					<table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
						<tr> 
							<td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
						</tr>
						<tr> 
							<td height="10"></td>
						</tr>
						<tr> 
							<td align="center">
								<font face="Arial, Helvetica, sans-serif" size="2">
									<%= validateErrMsg %>
									<br><br>
									<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
								</font>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			</table>
		<% Upload.Delete
		Set Upload = nothing
		Response.End
	End If

	Upload.Form("FILE1").Save

	FileName = Mid(Upload.Form("file1").UserFilename, InstrRev(Upload.Form("file1").UserFilename, "\") + 1)
	FileName = lcase(Filename)

	If (pcv_ResizeObj=1) then
		call UseASPJpeg(FileName,uploadpath & FileName)
	else
		If (pcv_ResizeObj=2) then
			call UseAspImage(FileName,uploadpath & FileName)
		Else
			DidntResize=1
		End if
	end if

	Upload.Delete
	Set Upload = nothing

End Function

Function UseASPUpload()

	Set Upload = Server.CreateObject("Persits.Upload")
	
	If (pcv_ResizeObj<>1) AND (pcv_ResizeObj<>2) then
		If PPD="1" then
			Upload.SaveVirtual "\"&scPcFolder&"\pc\catalog\"
		else
			Upload.SaveVirtual "..\..\pc\catalog\"
		end if
	Else
		If PPD="1" then
			Upload.SaveVirtual "\"&scPcFolder&"\includes\uploadresize\"
		else
			Upload.SaveVirtual "..\..\includes\uploadresize\"
		end if
	End if
	
	thumbnailsize = Upload.Form("thumbnailsize")
	generalsize = Upload.Form("generalsize")
	detailsize = Upload.Form("detailsize")
	sharpen = Upload.Form("sharpen")
	resizexy = Upload.Form("resizexy")
	
	call SetSessionVars(thumbnailsize, generalsize, detailsize)
	
	countfiles = 0
	For Each File in Upload.Files
		countfiles = countfiles + 1
	Next
	
	'Count files in upload.  If none exist, exit script
	If countfiles = 0 Then%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload any images.</b><br><br>
									<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        </table>
        <% Response.End
	End If
	
	'Run the resizer 
	For Each File in Upload.Files
		validateErrMsg = ValidateImageType(File.FileName, File.ImageType)
		
		If Len(validateErrMsg) > 0 then %>
			<table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
			<tr> 
				<td> 
				<table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
					<tr> 
					<td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
					</tr>
					<tr> 
					<td height="10"></td>
					</tr>
					<tr> 
					<td align="center">
						<font face="Arial, Helvetica, sans-serif" size="2">
							<%= validateErrMsg %>
							<br/><br/>
							<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
						</font>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			</table>
			<%	
			'Delete source file & end script
			File.Delete
			Response.End
		End if
	
		FileName = lcase(File.FileName)
	
		If (pcv_ResizeObj=1) then
			call UseASPJpeg(FileName,File.Path)
		else
			If (pcv_ResizeObj=2) then
				call UseAspImage(FileName,File.Path)
			Else
				DidntResize=1
			End if
		end if
	
		'Delete source file
		File.Delete
	
	Next

End Function

Function UseASPSmartUpload()

	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	
	mySmartUpload.Upload
	
	If PPD="1" then
		intCount = mySmartUpload.Save(uploadpath)
	else
		intCount = mySmartUpload.Save(uploadpath)
	end if
	
	thumbnailsize = mySmartUpload.Form("thumbnailsize")
	generalsize = mySmartUpload.Form("generalsize")
	detailsize = mySmartUpload.Form("detailsize")
	sharpen = mySmartUpload.Form("sharpen")
	resizexy = mySmartUpload.Form("resizexy")
	
	call SetSessionVars(thumbnailsize, generalsize, detailsize)
	
	'Count files in mySmartUpload.  If none exist, exit script
	If intCount = 0 Then%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload any images.</b><br><br>
									<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        </table>
        <% Response.End
	End If
	
	'Run the resizer 
	For Each File in mySmartUpload.Files		
		
		FileName = File.FileName
		ImageType = Right(Replace(UCase(FileName), ".JPEG", ".JPG"), 3)
		
		validateErrMsg = ValidateImageType(FileName, ImageType)
	
		If Len(validateErrMsg) > 0 Then %>
			<table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
			<tr> 
				<td> 
					<table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
						<tr> 
							<td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
						</tr>
						<tr> 
							<td height="10"></td>
						</tr>
						<tr> 
							<td align="center">
								<font face="Arial, Helvetica, sans-serif" size="2">
									<%= validateErrMsg %>
									<br><br>
									<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
								</font>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			</table>
			<% 'Delete source file & end script
			Set fso=Server.CreateObject("Scripting.FileSystemObject")
			Set afi = fso.GetFile(uploadpath & FileName)
			afi.Delete
			Set afi=nothing
			Response.End
		End If
	
		If (pcv_ResizeObj=1) then
			call UseASPJpeg(FileName,uploadpath & FileName)
		else
			If (pcv_ResizeObj=2) then
				call UseAspImage(FileName,uploadpath & FileName)
			Else
				DidntResize=1
			End if
		end if
		
		'Delete source file
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set afi = fso.GetFile(uploadpath & FileName)
		afi.Delete
		Set afi=nothing
	
	Next

End Function

Sub ResizeX(intXSize)
	Dim intYSize
	intYSize = round((intXSize / Image.MaxX) * Image.MaxY)
	err.number=0
	Image.ResizeR intXSize, intYSize
	if err.number<>0 then
		Image.Resize intXSize, intYSize
	end if
End sub
	
Sub ResizeY(intYSize)
	Dim intXSize
	intXSize = round((intYSize / Image.MaxY) * Image.MaxX)
	err.number=0
	Image.ResizeR intXSize, intYSize
	if err.number<>0 then
		Image.Resize intXSize, intYSize
	end if
End sub

Sub UseAspImage(FileName,SourceFile)
	'Generate random number to append to filename
	randomnum = RandomNumber(2353)
	
	'Generate new thumbnail image filename
	If right(FileName, 4) = ".jpg" Then
		thumbfilename = replace(FileName,".jpg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		thumbfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		thumbfilename = replace(FileName,".jpe","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		thumbfilename = replace(FileName,".gif","") & "_" & randomnum & "_thumb.gif"
	ElseIf right(FileName, 4) = ".png" Then
		thumbfilename = replace(FileName,".gif","") & "_" & randomnum & "_thumb.png"
	End If
	thumbfilename = replace(thumbfilename,"%20","")
	thumbfilename = replace(thumbfilename," ","")
	
	'Generate new general image filename
	If right(FileName, 4) = ".jpg" Then
		generalfilename = replace(FileName,".jpg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		generalfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		generalfilename = replace(FileName,".jpe","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		generalfilename = replace(FileName,".gif","") & "_" & randomnum & "_general.gif"	
	ElseIf right(FileName, 4) = ".png" Then
		generalfilename = replace(FileName,".png","") & "_" & randomnum & "_general.png"	
	End If
	generalfilename = replace(generalfilename,"%20","")
	generalfilename = replace(generalfilename," ","")
	
	'Generate new detail image filename
	If right(FileName, 4) = ".jpg" Then
		detailfilename = replace(FileName,".jpg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		detailfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		detailfilename = replace(FileName,".jpe","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		detailfilename = replace(FileName,".gif","") & "_" & randomnum & "_detail.gif"
	ElseIf right(FileName, 4) = ".png" Then
		detailfilename = replace(FileName,".png","") & "_" & randomnum & "_detail.png"
	End If
	detailfilename = replace(detailfilename,"%20","")
	detailfilename = replace(detailfilename," ","")

	'---- SAVE THUMBNAIL IMAGE ----
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(SourceFile)

	BigBeforeWidth = Image.MaxX
	BigBeforeHeight = Image.MaxY
	
	If resizexy = "Width" Then
		jpg_width = cint(thumbnailsize)
		BigAfterWidth = jpg_width
		BigAfterHeight = round ((jpg_width / Image.MaxX) * Image.MaxY)
		If cint(BigAfterWidth) >= cint(BigBeforeWidth) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeX(jpg_width)
		End If
	Else
		jpg_height = cint(thumbnailsize)
		BigAfterHeight = jpg_height
		BigAfterWidth = round((jpg_height / Image.MaxY) * Image.MaxX)
		If cint(BigAfterHeight) >= cint(BigBeforeHeight) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeY(jpg_height)
		End If	
	End If
	
	Image.FileName = catalogpath & thumbfilename
	Image.SaveImage
	Set Image = nothing

	'---- SAVE GENERAL IMAGE ----
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(SourceFile)


	BigBeforeWidth = Image.MaxX
	BigBeforeHeight = Image.MaxY
	
	
	If resizexy = "Width" Then
		jpg_width = cint(generalsize)
		BigAfterWidth = jpg_width
		BigAfterHeight = round ((jpg_width / Image.MaxX) * Image.MaxY)
		If cint(BigAfterWidth) >= cint(BigBeforeWidth) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeX(jpg_width)
		End If
	Else
		jpg_height = cint(generalsize)
		BigAfterHeight = jpg_height
		BigAfterWidth = round((jpg_height / Image.MaxY) * Image.MaxX)
		If cint(BigAfterHeight) >= cint(BigBeforeHeight) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeY(jpg_height)
		End If	
	End If
	
	Image.FileName = catalogpath & generalfilename
	Image.SaveImage
	Set Image = nothing

	'---- SAVE DETAIL IMAGE ----
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(SourceFile)

	BigBeforeWidth = Image.MaxX
	BigBeforeHeight = Image.MaxY
	
	If resizexy = "Width" Then
		jpg_width = cint(detailsize)
		BigAfterWidth = jpg_width
		BigAfterHeight = round ((jpg_width / Image.MaxX) * Image.MaxY)
		If cint(BigAfterWidth) >= cint(BigBeforeWidth) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeX(jpg_width)
		End If
	Else
		jpg_height = cint(detailsize)
		BigAfterHeight = jpg_height
		BigAfterWidth = round((jpg_height / Image.MaxY) * Image.MaxX)
		If cint(BigAfterHeight) >= cint(BigBeforeHeight) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeY(jpg_height)
		End If	
	End If
	
	
	Image.FileName = catalogpath & detailfilename
	Image.SaveImage
	Set Image = nothing
End Sub

Sub UseASPJpeg(FileName,SourceFile)

	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	
	'Generate random number to append to filename
	randomnum = RandomNumber(2353)
	
	'Generate new thumbnail image filename
	If right(FileName, 4) = ".jpg" Then
		thumbfilename = replace(FileName,".jpg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		thumbfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		thumbfilename = replace(FileName,".jpe","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		thumbfilename = replace(FileName,".gif","") & "_" & randomnum & "_thumb.gif"
	ElseIf right(FileName, 4) = ".png" Then
		thumbfilename = replace(FileName,".png","") & "_" & randomnum & "_thumb.png"
	End If
	thumbfilename = replace(thumbfilename,"%20","")
	thumbfilename = replace(thumbfilename," ","")
	
	'Generate new general image filename
	If right(FileName, 4) = ".jpg" Then
		generalfilename = replace(FileName,".jpg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		generalfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		generalfilename = replace(FileName,".jpe","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		generalfilename = replace(FileName,".gif","") & "_" & randomnum & "_general.gif"	
	ElseIf right(FileName, 4) = ".png" Then
		generalfilename = replace(FileName,".png","") & "_" & randomnum & "_general.png"	
	End If
	generalfilename = replace(generalfilename,"%20","")
	generalfilename = replace(generalfilename," ","")
	
	'Generate new detail image filename
	If right(FileName, 4) = ".jpg" Then
		detailfilename = replace(FileName,".jpg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		detailfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		detailfilename = replace(FileName,".jpe","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		detailfilename = replace(FileName,".gif","") & "_" & randomnum & "_detail.gif"
	ElseIf right(FileName, 4) = ".png" Then
		detailfilename = replace(FileName,".png","") & "_" & randomnum & "_detail.png"
	End If
	detailfilename = replace(detailfilename,"%20","")
	detailfilename = replace(detailfilename," ","")
	
	'---- SAVE THUMBNAIL IMAGE ----
	Jpeg.Open SourceFile
		
	BigBeforeWidth = jpeg.OriginalWidth
	BigBeforeHeight = jpeg.OriginalHeight

	If resizexy = "Width" Then
		If cint(thumbnailsize) >= cint(BigBeforeWidth) Then
		Else
			BigAfterWidth = thumbnailsize
			BigAfterHeight = round((BigAfterWidth / jpeg.Width) * jpeg.Height)
			
			Jpeg.Width = BigAfterWidth
			jpeg.Height = BigAfterHeight
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85
			
		End If
	Else
		If cint(thumbnailsize) >= cint(BigBeforeHeight) Then
		Else
			BigAfterHeight = thumbnailsize
			BigAfterWidth = round((BigAfterHeight / jpeg.Height) * jpeg.Width)
	
			Jpeg.Height = BigAfterHeight
			jpeg.Width = BigAfterWidth
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85			
		End If	
	End If

	
	Jpeg.Save (catalogpath) & thumbfilename
	Jpeg.Close
	
	'---- SAVE GENERAL IMAGE ----
	Jpeg.Open SourceFile
	
	BigBeforeWidth = jpeg.OriginalWidth
	BigBeforeHeight = jpeg.OriginalHeight
	
	If resizexy = "Width" Then
		If cint(generalsize) >= cint(BigBeforeWidth) Then
		Else
			BigAfterWidth = generalsize
			BigAfterHeight = round((BigAfterWidth / jpeg.Width) * jpeg.Height)
			
			Jpeg.Width = BigAfterWidth
			jpeg.Height = BigAfterHeight
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85
			
		End If
	Else
		If cint(generalsize) >= cint(BigBeforeHeight) Then
		Else
			BigAfterHeight = generalsize
			BigAfterWidth = round((BigAfterHeight / jpeg.Height) * jpeg.Width)
	
			Jpeg.Height = BigAfterHeight
			jpeg.Width = BigAfterWidth
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85			
		End If	
	End If	
	
	
	Jpeg.Save (catalogpath) & generalfilename
	Jpeg.Close
	
	'---- SAVE DETAIL IMAGE ----
	Jpeg.Open SourceFile
	
	BigBeforeWidth = jpeg.OriginalWidth
	BigBeforeHeight = jpeg.OriginalHeight

	If resizexy = "Width" Then
		If cint(detailsize) >= cint(BigBeforeWidth) Then
		Else
			BigAfterWidth = detailsize
			BigAfterHeight = round((BigAfterWidth / jpeg.Width) * jpeg.Height)
			
			Jpeg.Width = BigAfterWidth
			jpeg.Height = BigAfterHeight
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85
			
		End If
	Else
		If cint(detailsize) >= cint(BigBeforeHeight) Then
		Else
			BigAfterHeight = detailsize
			BigAfterWidth = round((BigAfterHeight / jpeg.Height) * jpeg.Width)
	
			Jpeg.Height = BigAfterHeight
			jpeg.Width = BigAfterWidth
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85			
		End If	
	End If	
	
	Jpeg.Save (catalogpath) & detailfilename
	Jpeg.Close

End Sub

Function UseBasicUpload()
	'on error resume next
	pc_CodePage = Session.CodePage
	Session.CodePage = 1252
	Dim Upload : Set Upload = New clsUpload

	uploadErrorMsg = ""
	
	'// Catch any errors uploading
	If err.description & "" <> "" Then
		If InStr(err.description, "007") Then
			uploadErrorMsg = ""
			uploadErrorMsg = uploadErrorMsg  & "An Error occurred while attempting to upload your images: " & err.description & "<br><br>"
			uploadErrorMsg = uploadErrorMsg  & "This is probably due to a default configuration option in IIS, which only allows the entity size in a POST request to be 200,000 bytes (~200KB).<br><br>"
			uploadErrorMsg = uploadErrorMsg  & "You have to change this setting if you want to upload multiple images files or large image files through the ProductCart application (otherwise you can just FTP the file to the same location: e.g. the &quot;pc/catalog&quot; folder).<br><br>"
			uploadErrorMsg = uploadErrorMsg  & "To change this setting:<br><br>"
			uploadErrorMsg = uploadErrorMsg  & " - Open IIS Manager<br>"
			uploadErrorMsg = uploadErrorMsg  & " - Navigate the tree to your application<br>"
			uploadErrorMsg = uploadErrorMsg  & " - Double click the &quot;ASP&quot; icon in the main panel<br>"
			uploadErrorMsg = uploadErrorMsg  & " - Expand the &quot;Limits&quot; category<br>"
			uploadErrorMsg = uploadErrorMsg  & " - Modify the &quot;Maximum Requesting Entity Body Limit&quot; to a larger value."
		End If
		err.number=0
		err.description=""
	End If
	
	If Len(uploadErrorMsg) > 0 Then %>
		<table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
		<tr> 
			<td> 
				<table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
					<tr> 
						<td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
					</tr>
					<tr> 
						<td height="10"></td>
					</tr>
					<tr> 
						<td align="center">
							<font face="Arial, Helvetica, sans-serif" size="2">
								<%= uploadErrorMsg %>
								<br><br>
								<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
							</font>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
		</table>
		<%
		Response.End
	End If
	
	uploadedFiles=0

	If Len(uploadErrorMsg) < 1 Then
		'// Process Files
		For i = 0 To Upload.Files.Count - 1
			Set File = Upload.Files.Item(i)

			FileName = File.FileName
			ImageType = Right(Replace(UCase(FileName), ".JPEG", ".JPG"), 3)

			File.Save(uploadpath)

			uploadErrorMsg = ValidateImageType(FileName, ImageType)
	
			If uploadErrorMsg = "" Then
				FileName = LCase(FileName)
				uploadedFiles = uploadedFiles + 1
			Else
				DeleteUploadFile(FileName)
			End If
		Next
		
	If uploadedFiles = 0 Then%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload any images.</b><br><br>
									<a href="javascript:void(0);" onClick="history.go(-1);">Click Here to go Back</a>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        </table>
        <% Response.End
	End If

		'// Process Form
		For i = 0 To Upload.Form.Count - 1
			if Ucase(Upload.Form.Key(i))=Ucase("thumbnailsize") then
				thumbnailsize = Upload.Form.Item(i)
			end if
			if Ucase(Upload.Form.Key(i))=Ucase("generalsize") then
				generalsize = Upload.Form.Item(i)
			end if
			if Ucase(Upload.Form.Key(i))=Ucase("detailsize") then
				detailsize = Upload.Form.Item(i)
			end if
			if Ucase(Upload.Form.Key(i))=Ucase("sharpen") then
				sharpen = Upload.Form.Item(i)
			end if
			if Ucase(Upload.Form.Key(i))=Ucase("resizexy") then
				resizexy = Upload.Form.Item(i)
			end if
		Next
		call SetSessionVars(thumbnailsize, generalsize, detailsize)
	End If

	Set Upload = Nothing
	Session.CodePage = pc_CodePage
	
	If (pcv_ResizeObj=1) then
		call UseASPJpeg(FileName,uploadpath & FileName)
	else
		If (pcv_ResizeObj=2) then
			call UseAspImage(FileName,uploadpath & FileName)
		Else
			DidntResize=1
		End if
	end if
	
End Function

SELECT CASE pcv_UploadObj
	Case 1: UseSAFileUp()
	Case 2: UseASPUpload()
	Case 3: UseASPSmartUpload()
	Case Else: UseBasicUpload()
END SELECT%>

<script type=text/javascript>
	function fillparentform() {
		try{
			parent.opener.document.hForm.smallImageUrl.value = "<%= thumbfilename %>"
			parent.opener.document.hForm.imageUrl.value = "<%= generalfilename %>"
			parent.opener.document.hForm.largeImageUrl.value = "<%= detailfilename %>"
		}
		catch(err){}
	}
	
	fillparentform();
	
	imagename='';
	function enlrge(imgnme) {
		lrgewin=window.open("about:blank","","height=200,width=200")
		imagename=imgnme;
		setTimeout('update()',500)
	}
	
	function update() {
	doc=lrgewin.document;
	doc.open('text/html');
	doc.write('<HTML><HEAD><TITLE>Enlarged Image<\/TITLE><\/HEAD><BODY bgcolor="white" onLoad="if  (self.resizeTo)self.resizeTo((document.images[0].width+10),(document.images[0].height+80))" topmargin="4" leftmargin="0" rightmargin="0" bottommargin="0"><table border="0" cellspacing="0" cellpadding="0"><tr><td>');
	doc.write('<IMG SRC="' + imagename + '"><\/td><\/tr><tr><td><form name="viewn"><input type="image" src="../../pc/images/close.gif" align="right" value="Close Window" onClick="self.close()"><\/td><\/tr><\/table>');
	doc.write('<\/form><\/BODY><\/HTML>');
	doc.close();
	}
</script>
  <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
    <tr> 
      <td> 
        <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
          <tr> 
            <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
          </tr>
          <tr> 
            <td height="10"></td>
          </tr>
          <tr> 
            <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <b>Image Upload <%If DidntResize=0 then%>& Resizing<%End if%> Completed Successfully!</b><br><br>
			  <%If DidntResize=0 then%>
			  The filenames for the 3 images have been sent to the product window.<br><br><br>
			  <b>Thumbnail Image:</b><br><a href="javascript:enlrge('../../pc/catalog/<%= thumbfilename %>')"><%= thumbfilename %></a><br><br>
			  <b>General Image:</b><br><a href="javascript:enlrge('../../pc/catalog/<%= generalfilename %>')"><%= generalfilename %></a><br><br>
			  <b>Detail Image:</b><br><a href="javascript:enlrge('../../pc/catalog/<%= detailfilename %>')"><%= detailfilename %></a>
			  <%End If%>
			  <br><br><br>
			  <a href="#" onClick="self.close()"><img src="../../pc/images/close.gif" alt="Close Window" border="0"></a>
			  </font></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
</body>
</html>