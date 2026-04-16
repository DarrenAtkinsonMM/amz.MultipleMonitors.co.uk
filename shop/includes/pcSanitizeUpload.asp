<%
Session.CodePage = 1252

Dim pcv_boolIsValidImage
Dim BlackList
BlackList= array(";", ":", ">", "<", "/" ,"\", "..", "?", "%", "$", "#", "&", ".asp", ".aspx", ".php", ".cgi", ".pl")


'// Sanitize Uploaded File
Public Function IsUploadAllowed(strFileName)
    on error resume next

    '// Validate Against Blacklist
    Dim tmpFileName
	tmpFileName=strFileName
	tmpFileName=Right(tmpFileName,Len(tmpFileName)-InstrRev(tmpFileName,"\"))
	IsUploadAllowed = True    
	TempStr = trim(tmpFileName)

	For i=lbound(BlackList) To ubound(BlackList)
 		If (instr(1,TempStr,BlackList(i),vbTextCompare)<>0) Then			
            IsUploadAllowed = False
			Exit Function
 		End If
 	Next

    '// Validate Extention
    extfile = ""
    If len(strFileName)>0 Then
        extfile=Right(ucase(strFileName),4)

        If Not (pcf_IsImagesOnly) Then
            '// *.txt, *.htm, *.html, *.gif, *.jpg, *.png, *.pdf, *.doc, *.zip, *.csv, *.xls
            If Not ((extfile=".TXT") Or (extfile=".HTM") Or (extfile=".GIF") Or (extfile=".JPG") Or (extfile=".PNG") Or (extfile=".PDF") Or (extfile=".DOC") Or (extfile=".ZIP") Or (extfile=".CSV") Or (extfile=".XLS")) Then
                IsUploadAllowed = False
                Exit Function         
            End If
        Else
            '// *.gif, *.jpg, *.png
            If Not ((extfile=".GIF") Or (extfile=".JPG") Or (extfile=".PNG")) Then
                IsUploadAllowed = False
                Exit Function         
            End If
        End If

        '// Extended Image Validation
        Dim pcv_strControlId 
        
        If ((extfile=".GIF") or (extfile=".JPG") or (extfile=".PNG")) Then

            pcv_strSection = Request.ServerVariables("SCRIPT_NAME")
			
            '// If "imageupl_popup.asp"
            If (InStr(lcase(pcv_strSection),"imageupl_popup.asp")>0) Then                    
                If Not (UploadRequest Is Nothing) Then

                    Select Case strFileName    
                        Case UploadRequest.Item("one").Item("FileName") : pcv_strControlId = "one"        
                        Case UploadRequest.Item("two").Item("FileName") : pcv_strControlId = "two"        
                        Case UploadRequest.Item("three").Item("FileName") : pcv_strControlId = "three"      
                        Case UploadRequest.Item("four").Item("FileName") : pcv_strControlId = "four"       
                        Case UploadRequest.Item("five").Item("FileName") : pcv_strControlId = "five"       
                        Case UploadRequest.Item("six").Item("FileName") : pcv_strControlId = "six"       
                    End Select 

                    Call IsValidImageType(UploadRequest.Item(pcv_strControlId).Item("Value"))

                    If len(pcv_boolIsValidImage)>0 Then
                        If pcv_boolIsValidImage = False Then
                            IsUploadAllowed = False
                            Exit Function
                        End If
                    End If

                End If
            End IF

        End If

    Else
        IsUploadAllowed = False
        Exit Function  
    End If

End Function


Public Function pcf_IsImagesOnly()

    pcv_boolImagesOnly = True
    
    pcv_strSection = Request.ServerVariables("SCRIPT_NAME")

    If (InStr(lcase(pcv_strSection),"userfileupl_popup.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"taxupl_popup.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"app-step1a.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"upload.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"catstep1a.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"cwstep1a.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"iistep2.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"ship-step1a.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"step1a.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"taxupl_popup.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"fileuploadb.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"fileuploadc.asp")>0) _
        Or (InStr(lcase(pcv_strSection),"adminfileupl_popup.asp")>0) Then
        pcv_boolImagesOnly = False
    End If
    
    pcf_IsImagesOnly = pcv_boolImagesOnly
    
End Function


Public Sub IsValidImageType(data)
    If LenB(data) > 0 Then
        Set objFileData = New FileData
        objFileData.Contents = data
        If objFileData.ImageWidth <= 0 Then
			pcv_boolIsValidImage = False
        Else
            pcv_boolIsValidImage = True
        End If
        Set objFileData = Nothing
    End If 
End Sub

Private Function AsciiToBinary(strAscii)
    Dim i, char, result
    result = ""
    For i=1 to Len(strAscii)
        char = Mid(strAscii, i, 1)
        result = result & chrB(AscB(char))
    Next
    AsciiToBinary = result
End Function

Private Function BinaryToAscii(strBinary)
    Dim i, result
    result = ""
    For i=1 to LenB(strBinary)
        result = result & chr(AscB(MidB(strBinary, i, 1))) 
    Next
    BinaryToAscii = result
End Function

Class FileData
    Private m_fileName
    Private m_contentType
    Private m_BinaryContents
    Private m_AsciiContents
    Private m_imageWidth
    Private m_imageHeight
    Private m_checkImage

    Public Property Get FileName
        FileName = m_fileName
    End Property

    Public Property Get ContentType
        ContentType = m_contentType
    End Property

    Public Property Get ImageWidth
        If m_checkImage=False Then Call CheckImageDimensions
        ImageWidth = m_imageWidth
    End Property

    Public Property Get ImageHeight
        If m_checkImage=False Then Call CheckImageDimensions
        ImageHeight = m_imageHeight
    End Property

    Public Property Let FileName(strName)
        Dim arrTemp
        arrTemp = Split(strName, "\")
        m_fileName = arrTemp(UBound(arrTemp))
    End Property

    Public Property Let CheckImage(blnCheck)
        m_checkImage = blnCheck
    End Property

    Public Property Let ContentType(strType)
        m_contentType = strType
    End Property

    Public Property Let Contents(strData)
        m_BinaryContents = strData
        m_AsciiContents = RSBinaryToString(m_BinaryContents)
    End Property

    Public Property Get Size
        Size = LenB(m_BinaryContents)
    End Property

    Private Sub CheckImageDimensions
        Dim width, height, colors
        Dim strType
        If gfxSpex(m_AsciiContents, width, height, colors, strType) = true then
            m_imageWidth = width
            m_imageHeight = height
        End If
        m_checkImage = True
    End Sub

    Private Sub Class_Initialize
        m_imageWidth = -1
        m_imageHeight = -1
        m_checkImage = False
    End Sub

    Private Function GetExtension(strPath)
        Dim arrTemp
        arrTemp = Split(strPath, ".")
        GetExtension = ""
        If UBound(arrTemp)>0 Then
            GetExtension = arrTemp(UBound(arrTemp))
        End If
    End Function

    Private Function RSBinaryToString(xBinary)
        'Antonin Foller, http://www.motobit.com
        'RSBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
        'to a string (BSTR) using ADO recordset

        Dim Binary
        'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
        If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary

        Dim RS, LBinary
        Const adLongVarChar = 201
        Set RS = CreateObject("ADODB.Recordset")
        LBinary = LenB(Binary)

        If LBinary>0 Then
            RS.Fields.Append "mBinary", adLongVarChar, LBinary
            RS.Open
            RS.AddNew
            RS("mBinary").AppendChunk Binary 
            RS.Update
            RSBinaryToString = RS("mBinary")
        Else  
            RSBinaryToString = ""
        End If
    End Function

    Function MultiByteToBinary(MultiByte)
        '© 2000 Antonin Foller, http://www.motobit.com
        ' MultiByteToBinary converts multibyte string To real binary data (VT_UI1 | VT_ARRAY)
        ' Using recordset
        Dim RS, LMultiByte, Binary
        Const adLongVarBinary = 205
        Set RS = CreateObject("ADODB.Recordset")
        LMultiByte = LenB(MultiByte)
        If LMultiByte>0 Then
            RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
            RS.Open
            RS.AddNew
            RS("mBinary").AppendChunk MultiByte & ChrB(0)
            RS.Update
            Binary = RS("mBinary").GetChunk(LMultiByte)
        End If
        MultiByteToBinary = Binary
    End Function

    Private Function BinaryToAscii(strBinary)
        Dim i, result
        result = ""
        For i=1 to LenB(strBinary)
            result = result & chr(AscB(MidB(strBinary, i, 1))) 
        Next
        BinaryToAscii = result
    End Function

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  This routine will attempt to identify any filespec passed  :::
    ':::  as a graphic file (regardless of the extension). This will :::
    ':::  work with BMP, GIF, JPG and PNG files.                     :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::          Based on ideas presented by David Crowell          :::
    ':::                   (credit where due)                        :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    '::: blah blah     Copyright *c* MM,  Mike Shaffer     blah blah :::
    '::: bh blah      ALL RIGHTS RESERVED WORLDWIDE      blah blah :::
    '::: blah blah  Permission is granted to use this code blah blah :::
    '::: blah blah   in your projects, as long as this     blah blah :::
    '::: blah blah      copyright notice is included       blah blah :::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    '::: blah blah blah blah blah blah blah blah blah blah blah blah :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  This function gets a specified number of bytes from any    :::
    ':::  file, starting at the offset (base 1)                      :::
    ':::                                                             :::
    ':::  Passed:                                                    :::
    ':::       flnm        => Filespec of file to read               :::
    ':::       offset      => Offset at which to start reading       :::
    ':::       bytes       => How many bytes to read                 :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Function GetBytes(flnm, offset, bytes)
        Dim startPos
        If offset=0 Then
            startPos = 1
        Else  
            startPos = offset
        End If
        if bytes = -1 then        ' Get All!
            GetBytes = flnm
        else
            GetBytes = Mid(flnm, startPos, bytes)
        end if
    End Function

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  Functions to convert two bytes to a numeric value (long)   :::
    ':::  (both little-endian and big-endian)                        :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Private Function lngConvert(strTemp)
        lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
    end function

    Private Function lngConvert2(strTemp)
        lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
    end function

    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                             :::
    ':::  This function does most of the real work. It will attempt  :::
    ':::  to read any file, regardless of the extension, and will    :::
    ':::  identify if it is a graphical image.                       :::
    ':::                                                             :::
    ':::  Passed:                                                    :::
    ':::       flnm        => Filespec of file to read               :::
    ':::       width       => width of image                         :::
    ':::       height      => height of image                        :::
    ':::       depth       => color depth (in number of colors)      :::
    ':::       strImageType=> type of image (e.g. GIF, BMP, etc.)    :::
    ':::                                                             :::
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    function gfxSpex(flnm, width, height, depth, strImageType)
        dim strPNG 
        dim strGIF
        dim strBMP
        dim strType
        dim strBuff
        dim lngSize
        dim flgFound
        dim strTarget
        dim lngPos
        dim ExitLoop
        dim lngMarkerSize

        strType = ""
        strImageType = "(unknown)"

        gfxSpex = False

        strPNG = chr(137) & chr(80) & chr(78)
        strGIF = "GIF"
        strBMP = chr(66) & chr(77)

        strType = GetBytes(flnm, 0, 3)

        if strType = strGIF then                ' is GIF
            strImageType = "GIF"
            Width = lngConvert(GetBytes(flnm, 7, 2))
            Height = lngConvert(GetBytes(flnm, 9, 2))
            Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
            gfxSpex = True
        elseif left(strType, 2) = strBMP then        ' is BMP
            strImageType = "BMP"
            Width = lngConvert(GetBytes(flnm, 19, 2))
            Height = lngConvert(GetBytes(flnm, 23, 2))
            Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
            gfxSpex = True
        elseif strType = strPNG then            ' Is PNG
            strImageType = "PNG"
            Width = lngConvert2(GetBytes(flnm, 19, 2))
            Height = lngConvert2(GetBytes(flnm, 23, 2))
            Depth = getBytes(flnm, 25, 2)
            select case asc(right(Depth,1))
                case 0
                    Depth = 2 ^ (asc(left(Depth, 1)))
                    gfxSpex = True
                case 2
                    Depth = 2 ^ (asc(left(Depth, 1)) * 3)
                    gfxSpex = True
                case 3
                    Depth = 2 ^ (asc(left(Depth, 1)))  '8
                    gfxSpex = True
                case 4
                    Depth = 2 ^ (asc(left(Depth, 1)) * 2)
                    gfxSpex = True
                case 6
                    Depth = 2 ^ (asc(left(Depth, 1)) * 4)
                    gfxSpex = True
                case else
                    Depth = -1
            end select
        else
            strBuff = GetBytes(flnm, 0, -1)        ' Get all bytes from file
            lngSize = len(strBuff)
            flgFound = 0

            strTarget = chr(255) & chr(216) & chr(255)
            flgFound = instr(strBuff, strTarget)

            if flgFound = 0 then
                exit function
            end if

            strImageType = "JPG"
            lngPos = flgFound + 2
            ExitLoop = false

            do while ExitLoop = False and lngPos < lngSize
                do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
                    lngPos = lngPos + 1
                loop

                if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
                    lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
                    lngPos = lngPos + lngMarkerSize  + 1
                else
                    ExitLoop = True
                end if
            loop

            if ExitLoop = False then
                Width = -1
                Height = -1
                Depth = -1
            else
                Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
                Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
                Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
                gfxSpex = True
            end if
        end if
    End Function
End Class
%>