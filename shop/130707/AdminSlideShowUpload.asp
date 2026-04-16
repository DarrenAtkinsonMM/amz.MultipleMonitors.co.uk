<%@ LANGUAGE="VBSCRIPT" %>

<% pageTitle="Upload Slideshow Image" %>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcSanitizeUpload.asp"-->
<!--#include file="uploadresize/clsUpload.asp"-->
<%PageUpload=1%>
<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
<%

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START - Admin Slideshow Methods
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Function RandomNumber(intHighestNumber)
		Randomize
		RandomNumber = Int(Rnd * intHighestNumber) + 1
	End Function
	
	Function DeleteUploadFile(uploadFileName)
		Set FS = Server.CreateObject("Scripting.FileSystemObject")
		FS.DeleteFile(uploadpath & uploadFileName)
		Set FS = Nothing
	End Function

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
		End If

		If Len(uploadErrorMsg) < 1 Then
			'// Process Files
			For i = 0 To Upload.Files.Count - 1
				Set File = Upload.Files.Item(i)
	
				FileName = File.FileName
				ImageType = Right(Replace(UCase(FileName), ".JPEG", ".JPG"), 3)
	
				File.Save(uploadpath)
	
				uploadErrorMsg = ValidateImageType(FileName, ImageType)
		
				If uploadErrorMsg = "" Then
					uploadSlideName = LCase(FileName)
					uploadedFiles = uploadedFiles + 1
				Else
					DeleteUploadFile(FileName)
				End If
			Next

			'// Process Form
			Set formArr = Server.CreateObject("Scripting.Dictionary")
			For i = 0 To Upload.Form.Count - 1
				formArr.Add Upload.Form.Key(i), Upload.Form.Item(i)
			Next
		End If

		Set Upload = Nothing
		Session.CodePage = pc_CodePage
	End Function

	Function UseSAFileUp()
	Dim i
		Set Upload = Server.CreateObject("SoftArtisans.FileUp")
	
		Upload.Path = uploadpath
		
		For i=1 to 100
		If IsObject(Upload.Form("imgupload" & i)) Then
			If Not Upload.Form("imgupload" & i).IsEmpty Then
				Upload.Form("imgupload" & i).Save
				
				FileName = Upload.Form("imgupload" & i).UserFilename
				ImageType = Right(Replace(UCase(FileName), ".JPEG", ".JPG"), 3)
			End If
		End If
		
		uploadErrorMsg = ""
		If FileName <> "" Then
			uploadErrorMsg = ValidateImageType(FileName, ImageType)
			
			If uploadErrorMsg = "" Then
				uploadSlideName = LCase(FileName)
				uploadedFiles = uploadedFiles + 1
			Else
				Upload.Delete
			End If
		End If
		Next
		
		Set formArr = Server.CreateObject("Scripting.Dictionary")
		For Each formItem In Upload.Form
			formArr.Add formItem, Upload.Form(formItem)
		Next		
		
		Set Upload = Nothing
	End Function
	
	Function UseAspUpload()
		Set Upload = Server.CreateObject("Persits.Upload")
		
		on error resume next
		
		If PPD="1" then
			Upload.SaveVirtual "\"&scPcFolder&"\includes\uploadresize\"
		else
			Upload.SaveVirtual "..\includes\uploadresize\"
		end if
		
		If err.number <> 0 Then
			uploadErrorMsg = err.description
			err.number=0
			err.description=0
		Else
		
			uploadErrorMsg = ""
			For Each File in Upload.Files
				FileName = Mid(File.Path,InStrRev(File.Path,"\")+1,len(File.Path))
				ImageType = File.ImageType
				
				uploadErrorMsg = ValidateImageType(FileName, ImageType)
				
				If uploadErrorMsg = "" Then
					uploadSlideName = LCase(FileName)
					uploadedFiles = uploadedFiles + 1
				Else
					File.Delete
				End If
			Next
			
			Set formArr = Server.CreateObject("Scripting.Dictionary")
			For Each formItem In Upload.Form
				formArr.Add formItem.Name, formItem.Value
			Next
		End If
		
		Set Upload = Nothing
	End Function
	
	Function UseASPSmartUpload()
	Dim rs, query, idSlide
		Set SmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
		
		SmartUpload.Upload
		uploadedFiles = SmartUpload.Save(uploadpath)
		
		uploadErrorMsg = ""
		For Each File In SmartUpload.Files
			FileName = File.FileName
			ImageType = Right(Replace(UCase(FileName), ".JPEG", ".JPG"), 3)
			
			uploadErrorMsg = ValidateImageType(FileName, ImageType)
			
			If uploadErrorMsg = "" Then
				uploadSlideName = LCase(FileName)
				uploadedFiles = uploadedFiles + 1
			Else
				DeleteUploadFile(FileName)
			End If
		Next
		
		Set formArr = Server.CreateObject("Scripting.Dictionary")
		formArr.Add "useDefault", SmartUpload.Form("useDefault")
		formArr.Add "cropslide", SmartUpload.Form("cropslide")
		formArr.Add "slideWidth", SmartUpload.Form("slideWidth")
		formArr.Add "slideHeight", SmartUpload.Form("slideHeight")
		formArr.Add "addimage", SmartUpload.Form("addimage")
		formArr.Add "effect", SmartUpload.Form("effect")
		formArr.Add "animSpeed", SmartUpload.Form("animSpeed")
		formArr.Add "pauseTime", SmartUpload.Form("pauseTime")
		formArr.Add "submit", SmartUpload.Form("submit")
		formArr.Add "slideName", SmartUpload.Form("slideName")
		formArr.Add "slideCaption", SmartUpload.Form("slideCaption")
		formArr.Add "slideUrl", SmartUpload.Form("slideUrl")
		formArr.Add "slideAlt", SmartUpload.Form("slideAlt")
		formArr.Add "slideX", SmartUpload.Form("slideX")
		formArr.Add "slideY", SmartUpload.Form("slideY")
		formArr.Add "reductionFactor", SmartUpload.Form("reductionFactor")
		
		query="SELECT idSlide FROM pcSlideShow WHERE idSetting = " & settingId
		If settingId = 1 Then
			query=query&" OR idSetting IS NULL"
		End If
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
			
		if not rs.eof then
			rs.MoveFirst
			do while not rs.eof
				idSlide = rs("idSlide")
				formArr.Add "slideCaption_" & idSlide, SmartUpload.Form("slideCaption_" & idSlide)
				formArr.Add "slideUrl_" & idSlide, SmartUpload.Form("slideUrl_" & idSlide)
				formArr.Add "slideAlt_" & idSlide, SmartUpload.Form("slideAlt_" & idSlide)
				formArr.Add "slideOrder_" & idSlide, SmartUpload.Form("slideOrder_" & idSlide)
						
				rs.MoveNext
			loop
		end if
		set rs=nothing

	End Function
			
	Function CropResizeAspJpeg(FileName, SourceFile, cropX, cropY, cropWidth, cropHeight, finalWidth, finalHeight)
		Set Jpeg = Server.CreateObject("Persits.Jpeg")
		Jpeg.Open SourceFile
		
		Jpeg.Crop cropX, cropY, cropX + cropWidth, cropY + cropHeight
		Jpeg.Width = finalWidth
		Jpeg.Height = finalHeight
		
		'// Set the quality
		Jpeg.Quality = 100
		
		'// Save the image
		Jpeg.Save catalogpath & FileName
		Jpeg.Close
		
		Set Jpeg = Nothing
	End Function
	
	Function CropResizeAspImage(FileName, SourceFile, cropX, cropY, cropWidth, cropHeight, finalWidth, finalHeight)
		Set Image = Server.CreateObject("AspImage.Image")
		Image.LoadImage(SourceFile)
		
		'on error resume next
		
		Image.CropImage cropX, cropY, cropX + cropWidth, cropY + cropHeight
		
		Image.ResizeR finalWidth, finalHeight
		
		'// Set the quality
		Image.JPEGQuality = 100
		
		'// Save the image
		Image.FileName = catalogpath & FileName
		Image.SaveImage
		
		Set Image = Nothing
	End Function
	
	Function GetSlideFileName(uploadFileName)
		randomnum = RandomNumber(5000)
		
		'get filename
		If right(uploadFileName, 4) = ".jpg" Then
			slideFileName = replace(uploadFileName,".jpg","") & "_" & randomnum & ".jpg"
		ElseIf right(uploadFileName, 5) = ".jpeg" Then
			slideFileName = replace(uploadFileName,".jpeg","") & "_" & randomnum & ".jpg"
		ElseIf right(uploadFileName, 4) = ".jpe" Then
			slideFileName = replace(uploadFileName,".jpe","") & "_" & randomnum & ".jpg"
		ElseIf right(uploadFileName, 4) = ".gif" Then
			slideFileName = replace(uploadFileName,".gif","") & "_" & randomnum & ".gif"
		ElseIf right(uploadFileName, 4) = ".png" Then
			slideFileName = replace(uploadFileName,".png","") & "_" & randomnum & ".png"
		End If
		slideFileName = replace(slideFileName,"%20","")
		slideFileName = replace(slideFileName," ","")
			
		GetSlideFileName = slideFileName
	End Function
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END - Admin Slideshow Methods
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>

<%
	pcBaseName = "AdminSlideShow.asp"
	pcPageName = Request.ServerVariables("SCRIPT_NAME")

	uploadFieldName = "imgupload"

	'// Get PC catalog directory path
	if PPD="1" then
		catalogpath=Server.Mappath ("\"&scPcFolder&"\pc\catalog\")
	else
		catalogpath=Server.Mappath ("..\pc\catalog\")
	end if
	catalogpath = catalogpath & "\"
	
	'// Get upload directory path
	if PPD="1" then
		uploadpath=Server.Mappath ("\"&scPcFolder&"\includes\uploadresize\")
	else
		uploadpath=Server.Mappath ("..\includes\uploadresize\")
	end if
	uploadpath = uploadpath & "\"
		
	Dim formArr, uploadedFiles, uploadSlideName, uploadSlideWidth, uploadSlideHeight, uploadErrorMsg
	
	uploadedFiles = 0
	uploadSizeError = false
	uploadSizeErrorStr = ""
  
	action = LCase(getUserInput(Request.QueryString("action"), 0))
	settingId = getUserInput(Request.QueryString("setting"), 0)
	slideId = getUserInput(Request.QueryString("slide"), 0)

	'// Default to save action
	If action & "" = "" Then
		action = "save"
	End If

	If settingId & "" = "" Then
		settingId = 1
	End If

	If settingId = 1 Then pcBaseName = pcBaseName & "#desktop"
	If settingId = 2 Then pcBaseName = pcBaseName & "#mobile"
	
	'// Load slideshow settings
	query="SELECT * FROM pcSlideShowSettings WHERE idSetting = " & settingId & ";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	pcIntSlideWidth=rstemp("slideWidth")
	pcIntSlideHeight=rstemp("slideHeight")
	pcStrEffect=rstemp("effect")
	pcIntAnimSpeed=rstemp("animSpeed")
	pcIntPauseTime=rstemp("pauseTime")
	set rstemp = nothing
	If err.number <> 0 Then
		call LogErrorToDatabase()
		call closeDb()
		response.redirect "techErr.asp?err=" & pcStrCustRefID
	End If
	
	'// Get slide information
	If slideId <> "" Then
		query="SELECT slideImage,slideCaption,slideUrl,slideAlt,slideNewWindow FROM pcSlideShow WHERE idSlide = " & slideId
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)
		If err.number <> 0 Then
			call LogErrorToDatabase()
			call closeDb()
			response.redirect "techErr.asp?err=" & pcStrCustRefID
		End If
		If Not rstemp.eof Then
			slideName = rstemp("slideImage")
			slideCaption = rstemp("slideCaption")
			slideUrl = rstemp("slideUrl")
			slideAlt = rstemp("slideAlt")
			slideNewWindow = rstemp("slideNewWindow")
		End If
		set rstemp=nothing
		
		slideIdStr = "&slide=" & slideId
	End If
	
	Select Case pcv_UploadObj
		Case 1: UseSAFileUp()
		Case 2: UseAspUpload()
		Case 3: UseASPSmartUpload()
    Case Else: UseBasicUpload()
	End Select
	
	If Len(uploadErrorMsg) > 0 Then		
		Session("message") = uploadErrorMsg
		Response.Redirect "msgb.asp?back=1"
	Else
		msg = ""
		If action = "upload" Then
			If uploadedFiles > 0 Then
				If (pcv_resizeObj > 0) AND (getUserInput(formArr("cropslide"), 0)="1") Then
					GetImageDimensions uploadpath & uploadSlideName, uploadSlideWidth, uploadSlideHeight
				Else
					action = "saveslide"
				End If
			Else
				msg = "You did not upload a slide image!"
			End If
		End If
	
		If Len(msg) < 1 Then
			Select Case action
			Case "addexist"
				slideFileName = getUserInput(formArr("addimage"), 0)
				query = "INSERT INTO pcSlideShow (idSetting, slideImage, slideCaption, slideUrl, slideAlt, slideDateUploaded, slideNewWindow) VALUES (" & settingId & ", '" & slideFileName & "', N'" & slideCaption & "', '" & slideUrl & "', N'" & slideAlt & "', '" & UtcNow() & "', 0);"
				conntemp.execute(query)
				
				Set FS = Nothing
				
				if err.number <> 0 then
					call LogErrorToDatabase()
					call closeDb()
					response.redirect "techErr.asp?err=" & pcStrCustRefID
				end if
				
				call closeDb()
				
				Session("pcCPCheckCode") = 1
				Session("pcCPCheckText") = "Added slideshow image successfully!"
		
				response.redirect pcBaseName
			Case "save"
				useDefault = getUserInput(formArr("useDefault"), 0)
		
				If useDefault & "" = "" Then
					useDefault = 0
				End If
		
				If useDefault = 1 Then
					query = "UPDATE pcSlideShowSettings SET useDefault = " & useDefault & " WHERE idSetting = " & settingId & ";"
				Else
					cropslide = getUserInput(formArr("cropslide"), 0)
					slideWidth = getUserInput(formArr("slideWidth"), 0)
                    If len(slideWidth)=0 Then
                        slideWidth = 0
                    End IF
					slideHeight = getUserInput(formArr("slideHeight"), 0)
                    If len(slideHeight)=0 Then
                        slideHeight = 0
                    End IF
					effect = getUserInput(formArr("effect"), 0)
					animSpeed = getUserInput(formArr("animSpeed"), 0)
					pauseTime = getUserInput(formArr("pauseTime"), 0)
				
					'update settings
					query = "UPDATE pcSlideShowSettings SET useDefault = " & useDefault & ", slideWidth=" & slideWidth & ", slideHeight=" & slideHeight & ", effect='" & effect & "', animSpeed=" & animSpeed & ", pauseTime=" & pauseTime & " WHERE idSetting = " & settingId & ";"
				End If
				conntemp.execute(query)
			
				'save form
				query="SELECT idSlide FROM pcSlideShow WHERE idSetting = " & settingId
				If settingId = 1 Then
				query=query&" OR idSetting IS NULL"
				End If
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
			
				if not rs.eof then
					rs.MoveFirst
					do while not rs.eof
						idSlide = rs("idSlide")
						
						slideCaption = pcf_ReplaceCharacters(formArr("slideCaption_" & idSlide))
						slideCaption = pcf_ReplaceQuotes(slideCaption)
						slideUrl = getUserInput(formArr("slideUrl_" & idSlide), 500)
						slideAlt = pcf_ReplaceCharacters(formArr("slideAlt_" & idSlide))
						slideAlt = pcf_ReplaceQuotes(slideAlt)
						slideOrder = getUserInput(formArr("slideOrder_" & idSlide), 0)
						slideNewWindow = getUserInput(formArr("slideNewWindow_" & idSlide), 0)
						slideStart = getUserInput(formArr("slideStart_" & idSlide), 0)
						slideEnd = getUserInput(formArr("slideEnd_" & idSlide), 0)
						if slideNewWindow = "" then
							slideNewWindow = 0
						end if
						if slideEnd = "1900-01-01" Then
							slideEnd = "2020-12-31"
						end if
						if len(slideEnd) = 0 Then
							slideEnd = "2020-12-31"
						end if
						if slideStart = "1900-01-01" Then
							slideStart = "2018-05-31"
						end if
						if len(slideStart) = 0 Then
							slideStart = "2018-05-31"
						end if
								
						query="UPDATE pcSlideShow SET idSetting = '" & settingId & "', slideOrder = '" & slideOrder & "', slideCaption = '" & slideCaption & "', slideUrl = '" & slideUrl & "', slideAlt = '" & slideAlt & "', slideDateUploaded = '" & UtcNow() & "', slideNewWindow = " & slideNewWindow & ", slideStart = '" & slideStart & "', slideEnd = '" & slideEnd & "' WHERE idSlide = " & idSlide & ";"
						conntemp.execute(query)
						If err.number <> 0 Then
							call LogErrorToDatabase()
							rs.Close
							set rs=nothing
							call closedb()
							Response.Redirect "techError.asp?err=" & pcStrCustRefID
						End If
						
						rs.MoveNext
					loop
				end if
				
				Session("pcCPCheckCode") = 1
				Session("pcCPCheckText") = "Successfully updated slideshow settings!"
		
				call closeDb()
				response.redirect pcBaseName
			Case "upload"
				action = "saveslide"
			Case "saveslide"
				'// We need the FS object here
				Set FS = Server.CreateObject("Scripting.FileSystemObject")
		
				If (pcv_ResizeObj > 0) AND (getUserInput(formArr("cropslide"), 0)="1") Then
					uploadSlideName = getUserInput(formArr("slideName"), 0)
				
					if formArr("submit") = "Cancel" then
						If FS.FileExists(uploadpath & uploadSlideName) Then FS.DeleteFile(uploadpath & uploadSlideName)
					
						call closeDb()
						response.redirect pcBaseName
					ElseIf formArr("submit") = "Back" then
						call closeDb()
						response.redirect pcBaseName
					end if
					
					slideFileName = GetSlideFileName(uploadSlideName)
				
					slideCaption = getUserInput(formArr("slideCaption"), 0)
					slideUrl = getUserInput(formArr("slideUrl"), 500)
					slideAlt = getUserInput(formArr("slideAlt"), 255)
					slideX = getUserInput(formArr("slideX"), 0)
					slideY = getUserInput(formArr("slideY"), 0)
					cropslide = getUserInput(formArr("cropslide"), 0)
					slideWidth = getUserInput(formArr("slideWidth"), 0)
					slideHeight = getUserInput(formArr("slideHeight"), 0)
					rf = getUserInput(formArr("reductionFactor"), 0)
		
					If slideX & "" = "" Then 
						slideX = CDbl(0)
					Else
						slideX = CDbl(slideX)
					End If
		
					If slideY & "" = "" Then 
						slideY = CDbl(0)
					Else
						slideY = CDbl(slideY)
					End If
		
					If slideWidth & "" = "" Then 
						slideWidth = CDbl(0)
					Else
						slideWidth = CDbl(slideWidth)
					End If
		
					If slideHeight & "" = "" Then 
						slideHeight = CDbl(0)
					Else
						slideHeight = CDbl(slideHeight)
					End If
		
					If rf & "" = "" Then 
						rf = CDbl(1)
					Else
						rf = CDbl(rf)
					End If
					
					cropX = slideX * r
					cropY = slideY * rf
					cropWidth = slideWidth * rf
					cropHeight = slideHeight * rf
					if cropslide="1" then
						Select Case pcv_ResizeObj
							Case 1: 
							CropResizeAspJpeg slideFileName, uploadpath & uploadSlideName, cropX, cropY, cropWidth, cropHeight, pcIntSlideWidth, pcIntSlideHeight
							Case 2: CropResizeAspImage slideFileName, uploadpath & uploadSlideName, cropX, cropY, cropWidth, cropHeight, pcIntSlideWidth, pcIntSlideHeight
						End Select
					end if
				Else
					slideFileName = GetSlideFileName(uploadSlideName)
		
					FS.CopyFile uploadpath & uploadSlideName, catalogpath & slideFileName
				End If
						
				'// Delete upload image
				If FS.FileExists(uploadpath & uploadSlideName) Then FS.DeleteFile(uploadpath & uploadSlideName)
				'// Check for previous uploaded image
				If FS.FileExists(catalogpath & slideName) Then FS.DeleteFile(catalogpath & slideName)
				
				If slideId <> "" Then
					query = "UPDATE pcSlideShow SET idSetting = " & settingId & ", slideImage = '" & slideFileName & "', slideCaption = '" & slideCaption & "', slideUrl = '" & slideUrl & "', slideAlt = '" & slideAlt & "' WHERE idSlide = " & slideId & ";"
				Else
					query = "INSERT INTO pcSlideShow (idSetting, slideImage, slideCaption, slideUrl, slideAlt, slideDateUploaded) VALUES (" & settingId & ", '" & slideFileName & "', N'" & slideCaption & "', '" & slideUrl & "', N'" & slideAlt & "', '" & UtcNow() & "');"
				End If
				conntemp.execute(query)
				
				Set FS = Nothing
				
				if err.number <> 0 then
					call LogErrorToDatabase()
					call closeDb()
					response.redirect "techErr.asp?err=" & pcStrCustRefID
				end if
				
				call closeDb()
				
				Session("pcCPCheckCode") = 1
				Session("pcCPCheckText") = "Added slideshow image successfully!"
		
				response.redirect pcBaseName
			End Select
		End If
	End If
	
%>

<!--#include file="AdminHeader.asp"-->

<link rel="stylesheet" type="text/css" href="../includes/jquery/jcrop/jquery.jcrop.min.css" />
<script type="text/javascript" src="../includes/jquery/jcrop/jquery.jcrop.min.js"></script>

<script type=text/javascript>
	$pc(window).on('load', function() {
		var minSlideWidth = <%= pcIntSlideWidth %>;
		var minSlideHeight = <%= pcIntSlideHeight %>;
		var targetAspectRatio = minSlideWidth / minSlideHeight;
		var originalSlideWidth = <%= uploadSlideWidth %>;
		var originalSlideHeight = <%= uploadSlideHeight %>;
		var originalAspectRatio = originalSlideWidth / originalSlideHeight;
		
		var reductionFactor = originalSlideWidth/$pc("#cropTarget").width();
		$pc("#reductionFactor").val(reductionFactor);
		
		var defaultX, defaultY, defaultWidth, defaultHeight;
		if (targetAspectRatio > originalAspectRatio) {
			defaultWidth = $pc("#cropTarget").width();
			defaultHeight = defaultWidth / targetAspectRatio;
			
			if (originalSlideHeight > originalSlideWidth) {
				defaultX = 0; 
				defaultY = 0;
			} else {
				//get midpoints
				var m1 = $pc("#cropTarget").height() / 2;
				var m2 = defaultHeight / 2;
				
				//center vertically
				defaultX = 0;
				defaultY = m1 - m2;
			}
		} else {
			defaultHeight = $pc("#cropTarget").height();
			defaultWidth = defaultHeight * targetAspectRatio;
			
			var m1 = $pc("#cropTarget").width() / 2;
			var m2 = defaultWidth / 2;
			
			defaultY = 0;
			if (Math.round(m1) == Math.round(m2)) {
				defaultX = 0;
			} else {
				defaultX = m1 + m2;
			}
		}
		
		$pc("#cropTarget").Jcrop({
			onChange: function(c) {				
				if (c.w * reductionFactor < minSlideWidth) {
					$pc("#slideDimError").html("The size of the selected area is smaller than the minimum slide dimensions (<strong>" + minSlideWidth + "</strong> x <strong>" + minSlideHeight + "</strong>). The image will be scaled up and will lose quality!");
					$pc("#slideDimError").slideDown();
				} else {
					$pc("#slideDimError").slideUp();
				}
				
				// Set hidden inputs (accurate)
				$pc("#slideX").val(c.x);
				$pc("#slideY").val(c.y);
				$pc("#slideWidth").val(c.w);
				$pc("#slideHeight").val(c.h);
				
				// Set display (rounded off)
				$pc(".slideXDisplay").html(Math.round(c.x * reductionFactor) + "px");
				$pc(".slideYDisplay").html(Math.round(c.y * reductionFactor) + "px");
				$pc(".slideWidthDisplay").html(Math.round(c.w * reductionFactor) + "px");
				$pc(".slideHeightDisplay").html(Math.round(c.h * reductionFactor) + "px");
			},
			aspectRatio: targetAspectRatio,
			setSelect: [defaultX, defaultY, defaultWidth, defaultHeight],
		}, function() {
			// this.getBounds()
		});
	});
</script>

<form method="post" enctype="multipart/form-data" action="<%= pcPageName %>?setting=<%= settingId %>&action=<%= action %><%= slideIdStr %>" name="slideForm" class="pcForms">
	<table class="pcCPcontent">
    <tr>
      <td colspan="2">
        <!--#include file="pcv4_showMessage.asp"-->
      </td>
    </tr>
    <% if uploadedFiles > 0 then %>
    	<tr>
      	<td style="width: 200px">Slide Caption (optional): </td>
        <td><input type="text" name="slideCaption" size="60" value="<%= slideCaption %>" /></td>
      </tr>
    	<tr>
      	<td style="width: 200px">Slide Link (optional): </td>
        <td><input type="text" name="slideUrl" size="60" value="<%= slideUrl %>" /></td>
      </tr>
    	<tr>
      	<td style="width: 200px">Slide Alt Text (optional): </td>
        <td><input type="text" name="slideAlt" size="40" value="<%= slideAlt %>" /></td>
      </tr>
      <tr>
      	<td colspan="2"><hr></td>
      </tr>
    	<tr>
      	<td colspan="2">
        	<p>
        		Click-and-drag on the image below to select the area of the image you would like to include for this slide. 
          	The aspect ratio is automatically locked to the slide dimensions you have selected. 
						<strong>NOTE:</strong> If you select an area smaller than the minimum slide image dimensions, the final image will be scaled up and will have a decrease in quality!
          </p>

        </td>
      </tr>
      <tr>
      	<td colspan="2">
        	<div class="pcCPslideCropArea">
        		<img id="cropTarget" src="../includes/uploadresize/<%= uploadSlideName %>" style="width: 100%;" />
							
						<h4>Selected Area*</h4>
						<table class="pcCPslideCropDim">
							<tr>
								<td>X Offset: <span class="slideXDisplay"></span></td>
								<td>Width: <span class="slideWidthDisplay"></span></td>
							</tr>
							<tr>
								<td>Y Offset: <span class="slideYDisplay"></span></td>
								<td>Height: <span class="slideHeightDisplay"></span></td>
							</tr>
						</table>
							
						<span class="pcSmallText">*The image will be resized to exactly <strong><%= pcIntSlideWidth %></strong> x <strong><%= pcIntSlideHeight %></strong> pixels regardless of the size of the selected area.</span>
							
          </div>
        	<input type="hidden" name="slideName" value="<%= uploadSlideName %>" />
					<input type="hidden" name="slideX" id="slideX" />
					<input type="hidden" name="slideY" id="slideY" />
					<input type="hidden" name="slideWidth" id="slideWidth" />
					<input type="hidden" name="slideHeight" id="slideHeight" />
					<input type="hidden" name="cropslide" id="cropslide" value="<%=getUserInput(formArr("cropslide"), 0)%>" />
          <input type="hidden" name="reductionFactor" id="reductionFactor" />
        </td>
      </tr>
      <tr>
      	<td colspan="2">  
        	<div id="slideDimError" class="pcCPmessage" style="display: none"></div>     
        </td>
      </tr>
      <tr> 
        <td class="pcCPspacer"></td>
      </tr>
      <tr> 
        <td colspan="2" align="center"> 
          <input name="submit" type="submit" class="btn btn-primary" value="Save Slide">
          &nbsp;
          <input type="submit" name="submit" value="Cancel">
        </td>
      </tr>

    <% else %>
    	<tr>
      	<td colspan="2">
          <a href="<%= pcBaseName %>" class="btn btn-default">Back</a>
        </td>
      </tr>
    <% end if %>
  </table>
</form>
<%Function UtcNow()
UtcNow = serverdate.toUTCString()
UtcNow = Replace(Right(UtcNow, Len(UtcNow) - Instr(UtcNow, ",")), "UTC", "")
End Function
%>
<script language="JScript" runat="server">
var serverdate=new Date();
</script>