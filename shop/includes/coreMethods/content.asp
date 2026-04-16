<%
'Get Image Paths
Public Function pcf_getImagePath(tmpStrPath, strFilename)
    
    Dim pcv_strRelativePath, pcv_strAbsolutePath
	
	strPath = tmpStrPath
	If strPath<>"" Then
		If Right(strPath,1)="/" Then
			strPath=Left(strPath,len(strPath)-1)
		end if
	End If
	
	Select Case strPath
		Case "../pc/catalog":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/catalog")
		Case "../UPSLicense":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("UPSLicense")
		Case "catalog":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/catalog")
		Case "images":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "../pc/images":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/images")
		Case "https://www.paypal.com/en_US/i/btn":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "https://www.paypalobjects.com/webstatic/en_US/btn":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "https://www.internetsecure.com/images":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "//assets.pinterest.com/images":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "images/sample":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/images/sample")
		Case "theme/v47/images":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/theme/v47/images")
		Case "":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case Else: 
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
	End Select

    If scCDN_IsEnabled = "1" Then
        strPath = pcv_strAbsolutePath
    Else
        strPath = pcv_strRelativePath
    End If

	If strPath<>"" then
        If Right(strPath,1)<>"/" then
            strPath = strPath & "/"
        End If
	End If

	pcf_getImagePath = strPath & strFileName
End Function


Public Function pcf_getAbsolutePath(strPath)
    pcf_getAbsolutePath = "//" & scCDN_Domain & "/" & scPcFolder & "/" & strPath
End Function


Public Function pcf_getJSPath(strPath, strFilename)
    
    Dim pcv_strRelativePath, pcv_strAbsolutePath
	
	if strPath<>"" then
		if Right(strPath,1)="/" then
			strPath=Left(strPath,len(strPath)-1)
		end if
	end if
	
	Select Case strPath
		Case "../includes":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes")
		Case "../includes/javascripts":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes/javascripts")
		Case "../includes/jquery/opentip":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes/jquery/opentip")
		Case "https://www.2checkout.com/static/checkout/javascript":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "https://developer.payeezy.com/v1":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "https://www.jellyfish.com/javascripts":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "https://checkout.google.com/files/digital":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "//assets.pinterest.com/js":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case "service/app":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/service/app")
		Case "../includes/jquery/smoothmenu":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes/jquery/smoothmenu")
		Case "../includes/mojozoom":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes/mojozoom")
		Case "js":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/js")
		Case "":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
		Case Else: 
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
	End Select
    
    If scCDN_IsEnabled = "1" Then
        strPath = pcv_strAbsolutePath
    Else
        strPath = pcv_strRelativePath
    End If
	
	If strPath<>"" then
        If Right(strPath,1)<>"/" then
            strPath=strPath & "/"
        End If
	End If
	
	strPath=strPath & strFileName
	pcf_getJSPath = strPath
End Function


Public Function pcf_getCSSPath(strPath,strFilename)
    
    Dim pcv_strRelativePath, pcv_strAbsolutePath
	
	If strPath<>"" Then
		If Right(strPath,1)="/" Then
			strPath=Left(strPath,len(strPath)-1)
		End If
	End If
	
	Select Case strPath
		Case "css":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("pc/css")
		Case "../includes/jquery/opentip":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes/jquery/opentip")
		Case "../includes/mojozoom":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes/mojozoom")
		Case "../includes/jquery/nivo-slider":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("includes/jquery/nivo-slider")
		Case "theme/v47/css":
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = pcf_getAbsolutePath("theme/v47/css")
		Case Else: 
			pcv_strRelativePath = strPath
            pcv_strAbsolutePath = strPath
	End Select
    
    If scCDN_IsEnabled = "1" Then
        strPath = pcv_strAbsolutePath
    Else
        strPath = pcv_strRelativePath
    End If
	
	If strPath<>"" then
        If Right(strPath,1)<>"/" then
            strPath=strPath & "/"
        End If
	End If
	
	strPath=strPath & strFileName
	pcf_getCSSPath=strPath
End Function
%>
