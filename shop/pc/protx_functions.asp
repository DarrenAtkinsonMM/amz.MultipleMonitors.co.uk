<%
' ***********************************************************************
' ** Version:			1.0	- 18/01/2002											**
' ** Author:			Mat Peck													**
' ** Function:		Contains simple procedures to encode, encrypt, 	 	**
' **					decode, decrypt and split the information POSTed 	**
' ** 					to and from VSP Form.									**
' **																				**
' ** Revision History:															**
' **	Version	Author			Date and notes								**
' **		1.0		Mat Peck		18/01/2002 - First release					**
' **		1.1		Mat Peck		07/03/2002 - Base64 routines patched		**
' ***********************************************************************


' ** Set variables to indentify the vendor **

	'VendorName="testvendor"
	'Password="testvendor"
	
' ** Your server's IP address or dns name and web app directory.  Full qualified **
' ** Examples : MyServer="https://www.newco.com/ASPFormKit/", MyServer="192.168.0.1/ASPFormKit", MyServer="http://localhost/ASPFormKit/" **

	MyServer="../../pc/"	
	
	eoln = chr(13) & chr(10)

' ** Set up the Base64 arrays

	const BASE_64_MAP_INIT ="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	dim nl
   	dim Base64EncMap(63)
   	dim Base64DecMap(127)
	initcodecs



' ** The initCodecs() routine initialises the Base64 arrays

   PUBLIC SUB initCodecs()
          nl = "<P>" & chr(13) & chr(10)
          dim max, idx
          max = len(BASE_64_MAP_INIT)
          for idx = 0 to max - 1
               Base64EncMap(idx) = mid(BASE_64_MAP_INIT, idx + 1, 1)
          next
          for idx = 0 to max - 1
               Base64DecMap(ASC(Base64EncMap(idx))) = idx
          next
   END SUB

' ** Base 64 Encoding function **

   PUBLIC FUNCTION base64Encode(plain)

			call initCodecs

          if len(plain) = 0 then
               base64Encode = ""
               exit function
          end if

          dim ret, ndx, by3, first, second, third
          by3 = (len(plain) \ 3) * 3
          ndx = 1
          do while ndx <= by3
               first  = asc(mid(plain, ndx+0, 1))
               second = asc(mid(plain, ndx+1, 1))
               third  = asc(mid(plain, ndx+2, 1))
               ret = ret & Base64EncMap(  (first \ 4) AND 63 )
               ret = ret & Base64EncMap( ((first * 16) AND 48) + ((second \ 16) AND 15 ) )
               ret = ret & Base64EncMap( ((second * 4) AND 60) + ((third \ 64) AND 3 ) )
               ret = ret & Base64EncMap( third AND 63)
               ndx = ndx + 3
          loop
          ' check for stragglers
          if by3 < len(plain) then
               first  = asc(mid(plain, ndx+0, 1))
               ret = ret & Base64EncMap(  (first \ 4) AND 63 )
               if (len(plain) MOD 3 ) = 2 then
                    second = asc(mid(plain, ndx+1, 1))
                    ret = ret & Base64EncMap( ((first * 16) AND 48) +((second \16) AND 15 ) )
                    ret = ret & Base64EncMap( ((second * 4) AND 60) )
               else
                    ret = ret & Base64EncMap( (first * 16) AND 48)
                    ret = ret & "="
               end if
               ret = ret & "="
          end if

          base64Encode = ret
     END FUNCTION

' ** Base 64 decoding function **

     PUBLIC FUNCTION base64Decode(scrambled)

          if len(scrambled) = 0 then
               base64Decode = ""
               exit function
          end if

          ' ignore padding
          dim realLen
          realLen = len(scrambled)
          do while mid(scrambled, realLen, 1) = "="
               realLen = realLen - 1
          loop
          do while instr(scrambled," ")<>0
          		scrambled=left(scrambled,instr(scrambled," ")-1) & "+" & mid(scrambled,instr(scrambled," ")+1)
          loop
          dim ret, ndx, by4, first, second, third, fourth
          ret = ""
          by4 = (realLen \ 4) * 4
          ndx = 1
          do while ndx <= by4
               first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
               second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
               third  = Base64DecMap(asc(mid(scrambled, ndx+2, 1)))
               fourth = Base64DecMap(asc(mid(scrambled, ndx+3, 1)))
               ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3))
               ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
               ret = ret & chr( ((third * 64) AND 255) +  (fourth AND 63) )
               ndx = ndx + 4
          loop
          ' check for stragglers, will be 2 or 3 characters
          if ndx < realLen then
               first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
               second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
               ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3))
               if realLen MOD 4 = 3 then
                    third = Base64DecMap(asc(mid(scrambled,ndx+2,1)))
                    ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
               end if
          end if

          base64Decode = ret
     END FUNCTION


' ** The SimpleXor encryption algorithm. **
' ** NOTE:		This is a placeholder really.  Future releases of VSP Form will use AES or TwoFish.  Proper encryption **
' ** 			This simple function and the Base64 will deter script kiddies and prevent the "View Source" type tampering **
' **			It won't stop a half decent hacker though, but the most they could do is change the amount field to something **
' **			else, so provided the vendor checks the reports and compares amounts, there is no harm done.  It's still **
' **			more secure than the other PSPs who don't both encrypting their forms at all **

Public Function SimpleXor(InString,Key)
    Dim myIN, myKEY, myC, myPub
    Dim Keylist()
    
    myIN = InString
    myKEY = Key
    
    redim KeyList(Len(myKEY))
    
    i = 1
    do while i<=Len(myKEY)
        KeyList(i) = Asc(Mid(myKEY, i, 1))
        i = i + 1
    loop       
    
    j = 1
    i = 1
    do while i<=Len(myIn)
        myC = myC & Chr(Asc(Mid(myIN, i, 1)) Xor KeyList(j))
        i = i + 1
        If j = Len(myKEY) Then j = 0
        j = j + 1
    loop
 
    SimpleXor = myC
End Function


' ** The getToken function. **
' ** NOTE:		A function of convenience that extracts the value from the "name=value&name2=value2..." VSP reply string **
' **			Works even if one of the values is a URL containing the & or = signs.  **

public function getToken(thisString,thisToken)

	' Can't just rely on & characters because these may be provided in the URLs.
	Dim Tokens
	Dim subString
	Tokens = Array("Status","StatusDetail","VendorTxCode","VPSTxID","TxAuthNo","AVSCV2","Amount")
	
	if instr(thisString,thisToken+"=")=0 then
	
		'  If the token isn't present, empty the output.  We can error later
		getToken=""
		
	else
		
		' Right get the rest of the string
		subString=mid(thisString,instr(thisString,thisToken)+len(thisToken)+1)
		
		' Now strip off all remaining tokens if they are present.
		
		i=0
		do while i<7
		
			'Find the next token and lop it off
			if Tokens(i)<>thisToken then
			
				if instr(subString,"&"+Tokens(i))<>0 then 
					substring=left(substring,instr(subString,"&"+Tokens(i))-1)
				end if
						
			end if
			
			i = i +1
		
		loop	
		
		getToken=subString
	
	end if

  
end function


' ## PJWW - AES code

' Rijndael.asp
' Copyright 2001 Phil Fresle 
' phil@frez.co.uk 
' http://www.frez.co.uk
' Implementation of the AES Rijndael Block Cipher. Inspired by Mike Scott's
' implementation in C. Permission for free direct or derivative use is granted
' subject to compliance with any conditions that the originators of the
' algorithm place on its exploitation.
' 3-Apr-2001: Functions added to the bottom for encrypting/decrypting large
' arrays of data. The entire length of the array is inserted as the first four
' bytes onto the front of the first block of the resultant byte array before
' encryption.
' 19-Apr-2001: Thanks to Paolo Migliaccio for finding a bug with 256 bit 
' key. Problem was in the gkey function. Now properly matches NIST values.

' =====================================
' Modified by Mat Peck at Sage Pay to run with 128-bit blocks (AES) with CBC and PKCS#5 padding.
' =====================================

Private m_lOnBits(30)
Private m_l2Power(30)
Private m_bytOnBits(7)
Private m_byt2Power(7)

Private m_InCo(3)

Private m_fbsub(255)
Private m_rbsub(255)
Private m_ptab(255)
Private m_ltab(255)
Private m_ftable(255)
Private m_rtable(255)
Private m_rco(29)

Private m_Nk
Private m_Nb
Private m_Nr
Private m_fi(23)
Private m_ri(23)
Private m_fkey(119)
Private m_rkey(119)

m_InCo(0) = &HB
m_InCo(1) = &HD
m_InCo(2) = &H9
m_InCo(3) = &HE
    
m_bytOnBits(0) = 1
m_bytOnBits(1) = 3
m_bytOnBits(2) = 7
m_bytOnBits(3) = 15
m_bytOnBits(4) = 31
m_bytOnBits(5) = 63
m_bytOnBits(6) = 127
m_bytOnBits(7) = 255
    
m_byt2Power(0) = 1
m_byt2Power(1) = 2
m_byt2Power(2) = 4
m_byt2Power(3) = 8
m_byt2Power(4) = 16
m_byt2Power(5) = 32
m_byt2Power(6) = 64
m_byt2Power(7) = 128
    
m_lOnBits(0) = 1
m_lOnBits(1) = 3
m_lOnBits(2) = 7
m_lOnBits(3) = 15
m_lOnBits(4) = 31
m_lOnBits(5) = 63
m_lOnBits(6) = 127
m_lOnBits(7) = 255
m_lOnBits(8) = 511
m_lOnBits(9) = 1023
m_lOnBits(10) = 2047
m_lOnBits(11) = 4095
m_lOnBits(12) = 8191
m_lOnBits(13) = 16383
m_lOnBits(14) = 32767
m_lOnBits(15) = 65535
m_lOnBits(16) = 131071
m_lOnBits(17) = 262143
m_lOnBits(18) = 524287
m_lOnBits(19) = 1048575
m_lOnBits(20) = 2097151
m_lOnBits(21) = 4194303
m_lOnBits(22) = 8388607
m_lOnBits(23) = 16777215
m_lOnBits(24) = 33554431
m_lOnBits(25) = 67108863
m_lOnBits(26) = 134217727
m_lOnBits(27) = 268435455
m_lOnBits(28) = 536870911
m_lOnBits(29) = 1073741823
m_lOnBits(30) = 2147483647
    
m_l2Power(0) = 1
m_l2Power(1) = 2
m_l2Power(2) = 4
m_l2Power(3) = 8
m_l2Power(4) = 16
m_l2Power(5) = 32
m_l2Power(6) = 64
m_l2Power(7) = 128
m_l2Power(8) = 256
m_l2Power(9) = 512
m_l2Power(10) = 1024
m_l2Power(11) = 2048
m_l2Power(12) = 4096
m_l2Power(13) = 8192
m_l2Power(14) = 16384
m_l2Power(15) = 32768
m_l2Power(16) = 65536
m_l2Power(17) = 131072
m_l2Power(18) = 262144
m_l2Power(19) = 524288
m_l2Power(20) = 1048576
m_l2Power(21) = 2097152
m_l2Power(22) = 4194304
m_l2Power(23) = 8388608
m_l2Power(24) = 16777216
m_l2Power(25) = 33554432
m_l2Power(26) = 67108864
m_l2Power(27) = 134217728
m_l2Power(28) = 268435456
m_l2Power(29) = 536870912
m_l2Power(30) = 1073741824

Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function LShiftByte(bytValue, bytShiftBits)
    If bytShiftBits = 0 Then
        LShiftByte = bytValue
        Exit Function
    ElseIf bytShiftBits = 7 Then
        If bytValue And 1 Then
            LShiftByte = &H80
        Else
            LShiftByte = 0
        End If
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    
    LShiftByte = ((bytValue And m_bytOnBits(7 - bytShiftBits)) * m_byt2Power(bytShiftBits))
End Function

Private Function RShiftByte(bytValue, bytShiftBits)
    If bytShiftBits = 0 Then
        RShiftByte = bytValue
        Exit Function
    ElseIf bytShiftBits = 7 Then
        If bytValue And &H80 Then
            RShiftByte = 1
        Else
            RShiftByte = 0
        End If
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    
    RShiftByte = bytValue \ m_byt2Power(bytShiftBits)
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function RotateLeftByte(bytValue, bytShiftBits)
    RotateLeftByte = LShiftByte(bytValue, bytShiftBits) Or RShiftByte(bytValue, (8 - bytShiftBits))
End Function

Private Function Pack(b())
    Dim lCount
    Dim lTemp
    
    For lCount = 0 To 3
        lTemp = b(lCount)
        Pack = Pack Or LShift(lTemp, (lCount * 8))
    Next
End Function

Private Function PackFrom(b(), k)
    Dim lCount
    Dim lTemp
    
    For lCount = 0 To 3
        lTemp = b(lCount + k)
        PackFrom = PackFrom Or LShift(lTemp, (lCount * 8))
    Next
End Function

Private Sub Unpack(a, b())
    b(0) = a And m_lOnBits(7)
    b(1) = RShift(a, 8) And m_lOnBits(7)
    b(2) = RShift(a, 16) And m_lOnBits(7)
    b(3) = RShift(a, 24) And m_lOnBits(7)
End Sub

Private Sub UnpackFrom(a, b(), k)
    b(0 + k) = a And m_lOnBits(7)
    b(1 + k) = RShift(a, 8) And m_lOnBits(7)
    b(2 + k) = RShift(a, 16) And m_lOnBits(7)
    b(3 + k) = RShift(a, 24) And m_lOnBits(7)
End Sub

Private Function xtime(a)
    Dim b
    
    If (a And &H80) Then
        b = &H1B
    Else
        b = 0
    End If
    
    xtime = LShiftByte(a, 1)
    xtime = xtime Xor b
End Function

Private Function bmul(x, y)
    If x <> 0 And y <> 0 Then
        bmul = m_ptab((CLng(m_ltab(x)) + CLng(m_ltab(y))) Mod 255)
    Else
        bmul = 0
    End If
End Function

Private Function SubByte(a)
    Dim b(3)
    
    Unpack a, b
    b(0) = m_fbsub(b(0))
    b(1) = m_fbsub(b(1))
    b(2) = m_fbsub(b(2))
    b(3) = m_fbsub(b(3))
    
    SubByte = Pack(b)
End Function

Private Function product(x, y)
    Dim xb(3)
    Dim yb(3)
    
    Unpack x, xb
    Unpack y, yb
    product = bmul(xb(0), yb(0)) Xor bmul(xb(1), yb(1)) Xor bmul(xb(2), yb(2)) Xor bmul(xb(3), yb(3))
End Function

Private Function InvMixCol(x)
    Dim y
    Dim m
    Dim b(3)
    
    m = Pack(m_InCo)
    b(3) = product(m, x)
    m = RotateLeft(m, 24)
    b(2) = product(m, x)
    m = RotateLeft(m, 24)
    b(1) = product(m, x)
    m = RotateLeft(m, 24)
    b(0) = product(m, x)
    y = Pack(b)
    
    InvMixCol = y
End Function

Private Function ByteSub(x)
    Dim y
    Dim z
    
    z = x
    y = m_ptab(255 - m_ltab(z))
    z = y
    z = RotateLeftByte(z, 1)
    y = y Xor z
    z = RotateLeftByte(z, 1)
    y = y Xor z
    z = RotateLeftByte(z, 1)
    y = y Xor z
    z = RotateLeftByte(z, 1)
    y = y Xor z
    y = y Xor &H63
    
    ByteSub = y
End Function

Public Sub gentables()
    Dim i
    Dim y
    Dim b(3)
    Dim ib
    
    m_ltab(0) = 0
    m_ptab(0) = 1
    m_ltab(1) = 0
    m_ptab(1) = 3
    m_ltab(3) = 1
    
    For i = 2 To 255
        m_ptab(i) = m_ptab(i - 1) Xor xtime(m_ptab(i - 1))
        m_ltab(m_ptab(i)) = i
    Next
    
    m_fbsub(0) = &H63
    m_rbsub(&H63) = 0
    
    For i = 1 To 255
        ib = i
        y = ByteSub(ib)
        m_fbsub(i) = y
        m_rbsub(y) = i
    Next
    
    y = 1
    For i = 0 To 29
        m_rco(i) = y
        y = xtime(y)
    Next
    
    For i = 0 To 255
        y = m_fbsub(i)
        b(3) = y Xor xtime(y)
        b(2) = y
        b(1) = y
        b(0) = xtime(y)
        m_ftable(i) = Pack(b)
        
        y = m_rbsub(i)
        b(3) = bmul(m_InCo(0), y)
        b(2) = bmul(m_InCo(1), y)
        b(1) = bmul(m_InCo(2), y)
        b(0) = bmul(m_InCo(3), y)
        m_rtable(i) = Pack(b)
    Next
End Sub

Public Sub gkey(nb, nk, key())                
    Dim i
    Dim j
    Dim k
    Dim m
    Dim N
    Dim C1
    Dim C2
    Dim C3
    Dim CipherKey(7)
    
    m_Nb = nb
    m_Nk = nk
    
    If m_Nb >= m_Nk Then
        m_Nr = 6 + m_Nb
    Else
        m_Nr = 6 + m_Nk
    End If
    
    C1 = 1
    If m_Nb < 8 Then
        C2 = 2
        C3 = 3
    Else
        C2 = 3
        C3 = 4
    End If
    
    For j = 0 To nb - 1
        m = j * 3
        
        m_fi(m) = (j + C1) Mod nb
        m_fi(m + 1) = (j + C2) Mod nb
        m_fi(m + 2) = (j + C3) Mod nb
        m_ri(m) = (nb + j - C1) Mod nb
        m_ri(m + 1) = (nb + j - C2) Mod nb
        m_ri(m + 2) = (nb + j - C3) Mod nb
    Next
    
    N = m_Nb * (m_Nr + 1)
    
    For i = 0 To m_Nk - 1
        j = i * 4
        CipherKey(i) = PackFrom(key, j)
    Next
    
    For i = 0 To m_Nk - 1
        m_fkey(i) = CipherKey(i)
    Next
    
    j = m_Nk
    k = 0
    Do While j < N
        m_fkey(j) = m_fkey(j - m_Nk) Xor _
            SubByte(RotateLeft(m_fkey(j - 1), 24)) Xor m_rco(k)
        If m_Nk <= 6 Then
            i = 1
            Do While i < m_Nk And (i + j) < N
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
                    m_fkey(i + j - 1)
                i = i + 1
            Loop
        Else
            i = 1
            Do While i < 4 And (i + j) < N
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
                    m_fkey(i + j - 1)
                i = i + 1
            Loop
            If j + 4 < N Then
                m_fkey(j + 4) = m_fkey(j + 4 - m_Nk) Xor _
                    SubByte(m_fkey(j + 3))
            End If
            i = 5
            Do While i < m_Nk And (i + j) < N
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor _
                    m_fkey(i + j - 1)
                i = i + 1
            Loop
        End If
        
        j = j + m_Nk
        k = k + 1
    Loop
    
    For j = 0 To m_Nb - 1
        m_rkey(j + N - nb) = m_fkey(j)
    Next
    
    i = m_Nb
    Do While i < N - m_Nb
        k = N - m_Nb - i
        For j = 0 To m_Nb - 1
            m_rkey(k + j) = InvMixCol(m_fkey(i + j))
        Next
        i = i + m_Nb
    Loop
    
    j = N - m_Nb
    Do While j < N
        m_rkey(j - N + m_Nb) = m_fkey(j)
        j = j + 1
    Loop
End Sub

Public Sub encrypt(buff())
    Dim i
    Dim j
    Dim k
    Dim m
    Dim a(7)
    Dim b(7)
    Dim x
    Dim y
    Dim t
    
    For i = 0 To m_Nb - 1
        j = i * 4
        
        a(i) = PackFrom(buff, j)
        a(i) = a(i) Xor m_fkey(i)
    Next
    
    k = m_Nb
    x = a
    y = b
    
    For i = 1 To m_Nr - 1
        For j = 0 To m_Nb - 1
            m = j * 3
            y(j) = m_fkey(k) Xor m_ftable(x(j) And m_lOnBits(7)) Xor _
                RotateLeft(m_ftable(RShift(x(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
                RotateLeft(m_ftable(RShift(x(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
                RotateLeft(m_ftable(RShift(x(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
            k = k + 1
        Next
        t = x
        x = y
        y = t
    Next
    
    For j = 0 To m_Nb - 1
        m = j * 3
        y(j) = m_fkey(k) Xor m_fbsub(x(j) And m_lOnBits(7)) Xor _
            RotateLeft(m_fbsub(RShift(x(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
            RotateLeft(m_fbsub(RShift(x(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
            RotateLeft(m_fbsub(RShift(x(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
        k = k + 1
    Next
    
    For i = 0 To m_Nb - 1
        j = i * 4
        UnpackFrom y(i), buff, j
        x(i) = 0
        y(i) = 0
    Next
End Sub

Public Sub decrypt(buff())
    Dim i
    Dim j
    Dim k
    Dim m
    Dim a(7)
    Dim b(7)
    Dim x
    Dim y
    Dim t
    
    For i = 0 To m_Nb - 1
        j = i * 4
        a(i) = PackFrom(buff, j)
        a(i) = a(i) Xor m_rkey(i)
    Next
    
    k = m_Nb
    x = a
    y = b
    
    For i = 1 To m_Nr - 1
        For j = 0 To m_Nb - 1
            m = j * 3
            y(j) = m_rkey(k) Xor m_rtable(x(j) And m_lOnBits(7)) Xor _
                RotateLeft(m_rtable(RShift(x(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
                RotateLeft(m_rtable(RShift(x(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
                RotateLeft(m_rtable(RShift(x(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
            k = k + 1
        Next
        t = x
        x = y
        y = t
    Next
    
    For j = 0 To m_Nb - 1
        m = j * 3
        
        y(j) = m_rkey(k) Xor m_rbsub(x(j) And m_lOnBits(7)) Xor _
            RotateLeft(m_rbsub(RShift(x(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
            RotateLeft(m_rbsub(RShift(x(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
            RotateLeft(m_rbsub(RShift(x(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
        k = k + 1
    Next
    
    For i = 0 To m_Nb - 1
        j = i * 4
        
        UnpackFrom y(i), buff, j
        x(i) = 0
        y(i) = 0
    Next
End Sub

Private Function IsInitialized(vArray)
    On Error Resume Next
    
    IsInitialized = IsNumeric(UBound(vArray))
End Function

Private Sub CopyBytesASP(bytDest, lDestStart, bytSource(), lSourceStart, lLength)
    Dim lCount
    
    lCount = 0
    Do
        bytDest(lDestStart + lCount) = bytSource(lSourceStart + lCount)
        lCount = lCount + 1
    Loop Until lCount = lLength
End Sub

public Sub XORBlock(bytData1,bytData2)
	Dim lCount

	for lCount=0 to 15
		bytData1(lCount) = (bytData1(lCount) XOR bytData2(lCount))
	next
end Sub

Public Function EncryptData(bytMessage, bytPassword)
    Dim bytKey(15)
    Dim bytIn()
    Dim bytOut()
	Dim bytLast(15)
    Dim bytTemp(15)
    Dim lCount
    Dim lLength
    Dim lEncodedLength
    Dim bytLen(3)
    Dim lPosition
    
    If Not IsInitialized(bytMessage) Then
        Exit Function
    End If
    If Not IsInitialized(bytPassword) Then
        Exit Function
    End If
    
    For lCount = 0 To UBound(bytPassword)
        bytKey(lCount) = bytPassword(lCount)
        If lCount = 15 Then
            Exit For
        End If
    Next
    
    gentables
    gkey 4, 4, bytKey
    
    lLength = UBound(bytMessage) + 1
    lEncodedLength = lLength
    
    If lEncodedLength Mod 16 <> 0 Then
        lEncodedLength = lEncodedLength + 16 - (lEncodedLength Mod 16)
    End If
	
    ReDim bytIn(lEncodedLength - 1)
    ReDim bytOut(lEncodedLength - 1)
    
	CopyBytesASP bytLast,0,bytPassword,0,16
	
	Unpack lLength, bytIn
    CopyBytesASP bytIn, 0, bytMessage, 0, lLength

    For lCount = 0 To lEncodedLength - 1 Step 16
        CopyBytesASP bytTemp, 0, bytIn, lCount, 16
		XORBlock bytTemp,bytLast 
        Encrypt bytTemp
        CopyBytesASP bytOut, lCount, bytTemp, 0, 16
        CopyBytesASP bytLast,0,bytTemp, 0, 16
    Next
    
    EncryptData = bytOut
End Function

Public Function DecryptData(bytIn, bytPassword)
    Dim bytMessage()
    Dim bytKey(15)
    Dim bytOut()
    Dim bytTemp(15)
 	Dim bytLast(15)
    Dim lCount
    Dim lLength
    Dim lEncodedLength
    Dim bytLen(3)
    Dim lPosition
    
    If Not IsInitialized(bytIn) Then
        Exit Function
    End If
    If Not IsInitialized(bytPassword) Then
        Exit Function
    End If
    
    lEncodedLength = UBound(bytIn) + 1
    
    If lEncodedLength Mod 16 <> 0 Then
        Exit Function
    End If
    
    For lCount = 0 To UBound(bytPassword)
        bytKey(lCount) = bytPassword(lCount)
        If lCount = 15 Then
            Exit For
        End If
    Next
    
    gentables
    gkey 4, 4, bytKey

	CopyBytesASP bytLast,0,bytPassword,0,16

    ReDim bytOut(lEncodedLength - 1)
    
    For lCount = 0 To lEncodedLength - 1 Step 16
        CopyBytesASP bytTemp, 0, bytIn, lCount, 16
        Decrypt bytTemp
		XORBlock bytTemp,bytLast
        CopyBytesASP bytLast, 0, bytIn, lCount, 16
        CopyBytesASP bytOut, lCount, bytTemp, 0, 16
    Next

    lLength = ubound(bytOut)
   
    ReDim bytMessage(lLength)
    CopyBytesASP bytMessage, 0, bytOut, 0, lLength+1
    
    DecryptData = bytMessage
End Function


Function AESEncrypt(sPlain, sPassword)
    Dim bytIn()
    Dim bytOut
    Dim bytPassword()
    Dim lCount
    Dim lLength
	Dim sTemp
	Dim lPadLength
	
    lLength = Len(sPlain)
	lPadLength=16-(lLength mod 16)

	for lCount=1 to lPadLength
		sPlain=sPlain & Chr(lPadLength)
	next

    lLength = Len(sPlain)
		
    ReDim bytIn(lLength-1)
    For lCount = 1 To lLength
        bytIn(lCount-1) = CByte(AscB(Mid(sPlain,lCount,1)))
    Next
	
    lLength = Len(sPassword)
    ReDim bytPassword(lLength-1)
    For lCount = 1 To lLength
        bytPassword(lCount-1) = CByte(AscB(Mid(sPassword,lCount,1)))
    Next

    bytOut = EncryptData(bytIn, bytPassword)

    sTemp = ""
    For lCount = 0 To UBound(bytOut)
        sTemp = sTemp & Right("0" & Hex(bytOut(lCount)), 2)
    Next

    AESEncrypt = sTemp
End Function

Function AESDecrypt(sCypher, sPassword)
    Dim bytIn()
    Dim bytOut
    Dim bytPassword()
    Dim lCount
    Dim lLength
	Dim sTemp
	
    lLength = Len(sCypher)
    ReDim bytIn(lLength/2-1)
    For lCount = 0 To lLength/2-1
        bytIn(lCount) = CByte("&H" & Mid(sCypher,lCount*2+1,2))
    Next
	
    lLength = Len(sPassword)
    ReDim bytPassword(lLength-1)
    For lCount = 1 To lLength
        bytPassword(lCount-1) = CByte(AscB(Mid(sPassword,lCount,1)))
    Next

    bytOut = DecryptData(bytIn, bytPassword)

    lLength = UBound(bytOut) + 1 - bytOut(UBound(bytOut))
	sTemp = ""
    For lCount = 0 To lLength - 1
        sTemp = sTemp & Chr(bytOut(lCount))
    Next

    AESDecrypt = sTemp
End Function


'** Wrapper function do encrypt an encode based on strEncryptionType setting **
public function EncryptAndEncode(strIn, strEncryptionPassword)	
	'** AES encryption, CBC blocking with PKCS5 padding then HEX encoding - DEFAULT **
	EncryptAndEncode = "@" & AESEncrypt(strIn, strEncryptionPassword)
end function

'** Wrapper function do decode then decrypt based on header of the encrypted field **
public function DecodeAndDecrypt(strIn, strEncryptionPassword)
	DecodeAndDecrypt = AESDecrypt(mid(strIn,2), strEncryptionPassword)	
end function

%>