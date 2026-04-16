<%
Function getByteString(StringStr)
	For i = 1 to Len(StringStr)
		char = Mid(StringStr,i,1)
		getByteString = getByteString & chrB(AscB(char))
	Next
End Function

Function getString(StringBin)
	getString = ""
	For intCount = 1 to LenB(StringBin)
		getString = getString & chr(AscB(MidB(StringBin, intCount, 1))) 
	Next
End Function

Sub BuildUploadRequest(RequestBin)
    PosBeg = 1
    PosEnd = InstrB(PosBeg, RequestBin, getByteString(chr(13)))
    boundary = MidB(RequestBin, PosBeg, PosEnd-PosBeg)
    boundaryPos = InstrB(1, RequestBin, boundary)
    
    Do until (boundaryPos = InstrB(RequestBin, boundary & getByteString("--")))

        Dim UploadControl
        Set UploadControl = Server.CreateObject("Scripting.Dictionary")

        Pos = InstrB(BoundaryPos, RequestBin, getByteString("Content-Disposition"))
        Pos = InstrB(Pos, RequestBin, getByteString("name="))
        PosBeg = Pos+6
        PosEnd = InstrB(PosBeg, RequestBin, getByteString(chr(34)))
        Name = getString(MidB(RequestBin, PosBeg, PosEnd-PosBeg))
        PosFile = InstrB(BoundaryPos, RequestBin, getByteString("filename="))
        PosBound = InstrB(PosEnd, RequestBin, boundary)

        If  PosFile <> 0 AND (PosFile < PosBound) Then
    
            PosBeg = PosFile + 10
            PosEnd = InstrB(PosBeg, RequestBin, getByteString(chr(34)))
            FileName = getString(MidB(RequestBin, PosBeg, PosEnd-PosBeg))
    
            UploadControl.Add "FileName", FileName
            Pos = InstrB(PosEnd, RequestBin, getByteString("Content-Type:"))
            PosBeg = Pos+14
            PosEnd = InstrB(PosBeg, RequestBin, getByteString(chr(13)))
    
            ContentType = getString(MidB(RequestBin, PosBeg, PosEnd-PosBeg))
            UploadControl.Add "ContentType", ContentType
    
            PosBeg = PosEnd+4
            PosEnd = InstrB(PosBeg, RequestBin, boundary)-2
            Value = MidB(RequestBin, PosBeg, PosEnd-PosBeg)
        Else
    
            Pos = InstrB(Pos, RequestBin, getByteString(chr(13)))
            PosBeg = Pos+4
            PosEnd = InstrB(PosBeg, RequestBin, boundary)-2
            Value = getString(MidB(RequestBin, PosBeg, PosEnd-PosBeg))
        End If

    	UploadControl.Add "Value", Value

    	UploadRequest.Add name, UploadControl

        BoundaryPos = InstrB(BoundaryPos+LenB(boundary), RequestBin, boundary)
    Loop

End Sub
%>