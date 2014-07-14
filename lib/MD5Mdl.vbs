' 以下のページ参考.
' http://www.geocities.co.jp/SiliconValley/4334/unibon/asp/bitshift2.html
' http://vbscript.boris-toll.at/index.php?search=Calculate%20MD5%20Hash.vbs
' このプログラムコードは RFC 1321 (The MD5 Message-Digest Algorithm) の参照インプリメンテーションを
' unibon が VBScript および VB(Visual Basic) に移植したものです。
' 権利については RFC 1321 中の記述が優先されます。
' This program code is "derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm".
Option Explicit

Public Function MDString(ByVal stringx)
    Dim state(3)
    Dim count(1)
    Dim buffer(63)
    
    Dim digest(15)
    Dim lenx
    lenx = Len(stringx)

    Call MD5Init(state, count, buffer)
    Call MD5Update(state, count, buffer, ba(stringx), lenx)
    Call MD5Final(digest, state, count, buffer)

    Dim s
    s = MDPrint(digest)
    MDString = s
End Function

Public Function MDFileHash(ByVal strFile)

	Dim strMD5 : strMD5 = ""
	Dim ofso : Set ofso = CreateObject("Scripting.FileSystemObject")

	If ofso.FileExists(strFile) then

		strMD5 = BinaryToString(ReadTextFile(strFile, ""))

		MDFileHash = MDString(strMD5)

	Else

		MDFileHash = strFile & VbCrLf & "Error: File not found"

	End if

End Function
' --------------------------------------
Function ReadTextFile(ByVal FileName, ByVal CharSet)

	Const adTypeText = 2
	Dim BinaryStream : Set BinaryStream = CreateObject("ADODB.Stream")

	BinaryStream.Type = adTypeText

	If Len(CharSet) > 0 Then

		BinaryStream.CharSet = CharSet

	End If

	BinaryStream.Open
	BinaryStream.LoadFromFile FileName

	ReadTextFile = BinaryStream.ReadText

End Function

' -----------------------------
Function BinaryToString(ByRef Binary)

Dim cl1, cl2, cl3, pl1, pl2, pl3
Dim L
	cl1 = 1
	cl2 = 1
	cl3 = 1
	L = LenB(Binary)

	Do While cl1<=L

		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1

		If cl3>300 Then
			pl2 = pl2 & pl3
			pl3 = ""
			cl3 = 1
			cl2 = cl2 + 1

			If cl2>200 Then

				pl1 = pl1 & pl2
				pl2 = ""
				cl2 = 1

			End If
		End If
	Loop

	BinaryToString = pl1 & pl2 & pl3

End Function

Private Function sl(ByVal x, ByVal n) ' 左シフト
    If n = 0 Then
        sl = x
    Else
        Dim k
        k = CLng(2 ^ (32 - n - 1))
        Dim d
        d = x And (k - 1)
        Dim c
        c = d * CLng(2 ^ n)
        If x And k Then
            c = c Or &H80000000
        End If
        sl = c
    End If
End Function

Private Function sr(ByVal x, ByVal n) ' 右シフト(算術(>>)ではなく論理(>>>)シフトに相当)
    If n = 0 Then
        sr = x
    Else
        Dim y
        y = x And &H7FFFFFFF
        Dim z
        If n = 32 - 1 Then
            z = 0
        Else
            z = y \ CLng(2 ^ n)
        End If
        If y <> x Then
            z = z Or CLng(2 ^ (32 - n - 1))
        End If
        sr = z
    End If
End Function

Private Function add(ByVal a, ByVal b) ' オーバフローを無視して 32 ビットの加算をおこなう。
    If a >= 0 And b <= 0 Then
        add = a + b
    ElseIf a <= 0 And b >= 0 Then
        add = a + b
    Else
        Dim x
        x = a And &H3FFFFFFF
        Dim y
        y = b And &H3FFFFFFF
        Dim z
        z = x + y
        Dim f
        f = 0
        If z And &H40000000 Then
            f = f + 1
        End If
        z = z And &H3FFFFFFF
        If a And &H40000000 Then
            f = f + 1
        End If
        If a And &H80000000 Then
            f = f + 2
        End If
        If b And &H40000000 Then
            f = f + 1
        End If
        If b And &H80000000 Then
            f = f + 2
        End If
        If f And 1 Then
            z = z Or &H40000000
        End If
        If f And 2 Then
            z = z Or &H80000000
        End If
        add = z
    End If
End Function

Private Function addCur(ByVal a, ByVal b) ' オーバフローを無視して 32 ビットの加算をおこなう。
    Dim c
    c = CCur(a) + CCur(b)
    If c > &H7FFFFFFF Then
        c = c - CCur(2 ^ 32)
    ElseIf c < &H80000000 Then
        c = c + CCur(2 ^ 32)
    End If
    addCur = CLng(c)
End Function

Private Function ba(ByVal s) ' 文字列を文字の配列に変換する。
    Dim r
    If Len(s) = 0 Then ' 要素数が 0 個の配列のみ特別扱いする。
        r = Array()
    Else
        ReDim a(Len(s) - 1)
        Dim i
        For i = 0 To Len(s) - 1
            a(i) = Asc(Mid(s, i + 1, 1))
        Next
        r = a
    End If
    ba = r
End Function

Private Function FX(ByVal x, ByVal y, ByVal z)
    FX = (x And y) Or ((Not x) And z)
End Function

Private Function GX(ByVal x, ByVal y, ByVal z)
    GX = (x And z) Or (y And (Not z))
End Function

Private Function HX(ByVal x, ByVal y, ByVal z)
    HX = x Xor y Xor z
End Function

Private Function IX(ByVal x, ByVal y, ByVal z)
    IX = y Xor (x Or (Not z))
End Function

Private Function ROTATE_LEFT(ByVal x, ByVal n)
    ROTATE_LEFT = sl(x, n) Or sr(x, 32 - n)
End Function

Private Sub FF(ByRef a, ByVal b, ByVal c, ByVal d, ByVal x, ByVal s, ByVal ac)
    a = add(add(add(a, FX(b, c, d)), x), ac)
    a = ROTATE_LEFT(a, s)
    a = add(a, b)
End Sub

Private Sub GG(ByRef a, ByVal b, ByVal c, ByVal d, ByVal x, ByVal s, ByVal ac)
    a = add(add(add(a, GX(b, c, d)), x), ac)
    a = ROTATE_LEFT(a, s)
    a = add(a, b)
End Sub

Private Sub HH(ByRef a, ByVal b, ByVal c, ByVal d, ByVal x, ByVal s, ByVal ac)
    a = add(add(add(a, HX(b, c, d)), x), ac)
    a = ROTATE_LEFT(a, s)
    a = add(a, b)
End Sub

Private Sub II(ByRef a, ByVal b, ByVal c, ByVal d, ByVal x, ByVal s, ByVal ac)
    a = add(add(add(a, IX(b, c, d)), x), ac)
    a = ROTATE_LEFT(a, s)
    a = add(a, b)
End Sub

Private Sub MD5Init(ByRef state, ByRef count, ByRef buffer)
    count(0) = 0
    count(1) = 0
    
    state(0) = &H67452301
    state(1) = &HEFCDAB89
    state(2) = &H98BADCFE
    state(3) = &H10325476
End Sub

Private Sub MD5Update(ByRef state, ByRef count, ByRef buffer, ByRef inputx, ByVal inputLen)
    Dim i
    Dim index
    Dim partLen

    index = sr(count(0), 3) And &H3F
    
    count(0) = add(count(0), sl(inputLen, 3))
    If count(0) < sl(inputLen, 3) Then
        count(1) = add(count(1), 1)
    End If

    count(1) = add(count(1), sr(inputLen, 29))

    partLen = 64 - index

    If inputLen >= partLen Then
        Call MD5_memcpy(buffer, index, inputx, 0, partLen)
        Call MD5Transform(state, buffer, 0)

        For i = partLen To inputLen - 63 - 1 Step 64
            Call MD5Transform(state, inputx, i)
        Next
        index = 0
    Else
        i = 0
    End If

    Call MD5_memcpy(buffer, index, inputx, i, inputLen - i)
End Sub

Private Sub MD5Final(ByRef digest, ByRef state, ByRef count, ByRef buffer)
    Dim bits(7)
    Dim index
    Dim padLen

    Call Encode(bits, count, 8)

    index = sr(count(0), 3) And &H3F
    If index < 56 Then
        padLen = 56 - index
    Else
        padLen = 120 - index
    End If
    
    Dim PADDING
    PADDING = Array( _
        &H80, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
    )
    Call MD5Update(state, count, buffer, PADDING, padLen)

    Call MD5Update(state, count, buffer, bits, 8)

    Call Encode(digest, state, 16)
    
    Dim i
    For i = 0 To UBound(state)
        state(i) = 0
    Next
    For i = 0 To UBound(count)
        count(i) = 0
    Next
    For i = 0 To UBound(buffer)
        buffer(i) = 0
    Next
End Sub

Private Sub MD5Transform(ByRef state, ByRef block, ByVal offset)
    Dim a
    a = state(0)
    Dim b
    b = state(1)
    Dim c
    c = state(2)
    Dim d
    d = state(3)
    
    Dim x(15)
    Call Decode(x, block, offset, 64)

    ' Round 1
    Call FF(a, b, c, d, x( 0),  7, &HD76AA478) '  1 S11
    Call FF(d, a, b, c, x( 1), 12, &HE8C7B756) '  2 S12
    Call FF(c, d, a, b, x( 2), 17, &H242070DB) '  3 S13
    Call FF(b, c, d, a, x( 3), 22, &HC1BDCEEE) '  4 S14
    Call FF(a, b, c, d, x( 4),  7, &HF57C0FAF) '  5 S11
    Call FF(d, a, b, c, x( 5), 12, &H4787C62A) '  6 S12
    Call FF(c, d, a, b, x( 6), 17, &HA8304613) '  7 S13
    Call FF(b, c, d, a, x( 7), 22, &HFD469501) '  8 S14
    Call FF(a, b, c, d, x( 8),  7, &H698098D8) '  9 S11
    Call FF(d, a, b, c, x( 9), 12, &H8B44F7AF) ' 10 S12
    Call FF(c, d, a, b, x(10), 17, &HFFFF5BB1) ' 11 S13
    Call FF(b, c, d, a, x(11), 22, &H895CD7BE) ' 12 S14
    Call FF(a, b, c, d, x(12),  7, &H6B901122) ' 13 S11
    Call FF(d, a, b, c, x(13), 12, &HFD987193) ' 14 S12
    Call FF(c, d, a, b, x(14), 17, &HA679438E) ' 15 S13
    Call FF(b, c, d, a, x(15), 22, &H49B40821) ' 16 S14

    ' Round 2
    Call GG(a, b, c, d, x( 1),  5, &HF61E2562) ' 17 S21
    Call GG(d, a, b, c, x( 6),  9, &HC040B340) ' 18 S22
    Call GG(c, d, a, b, x(11), 14, &H265E5A51) ' 19 S23
    Call GG(b, c, d, a, x( 0), 20, &HE9B6C7AA) ' 20 S24
    Call GG(a, b, c, d, x( 5),  5, &HD62F105D) ' 21 S21
    Call GG(d, a, b, c, x(10),  9,  &H2441453) ' 22 S22
    Call GG(c, d, a, b, x(15), 14, &HD8A1E681) ' 23 S23
    Call GG(b, c, d, a, x( 4), 20, &HE7D3FBC8) ' 24 S24
    Call GG(a, b, c, d, x( 9),  5, &H21E1CDE6) ' 25 S21
    Call GG(d, a, b, c, x(14),  9, &HC33707D6) ' 26 S22
    Call GG(c, d, a, b, x( 3), 14, &HF4D50D87) ' 27 S23
    Call GG(b, c, d, a, x( 8), 20, &H455A14ED) ' 28 S24
    Call GG(a, b, c, d, x(13),  5, &HA9E3E905) ' 29 S21
    Call GG(d, a, b, c, x( 2),  9, &HFCEFA3F8) ' 30 S22
    Call GG(c, d, a, b, x( 7), 14, &H676F02D9) ' 31 S23
    Call GG(b, c, d, a, x(12), 20, &H8D2A4C8A) ' 32 S24

    ' Round 3
    Call HH(a, b, c, d, x( 5),  4, &HFFFA3942) ' 33 S31
    Call HH(d, a, b, c, x( 8), 11, &H8771F681) ' 34 S32
    Call HH(c, d, a, b, x(11), 16, &H6D9D6122) ' 35 S33
    Call HH(b, c, d, a, x(14), 23, &HFDE5380C) ' 36 S34
    Call HH(a, b, c, d, x( 1),  4, &HA4BEEA44) ' 37 S31
    Call HH(d, a, b, c, x( 4), 11, &H4BDECFA9) ' 38 S32
    Call HH(c, d, a, b, x( 7), 16, &HF6BB4B60) ' 39 S33
    Call HH(b, c, d, a, x(10), 23, &HBEBFBC70) ' 40 S34
    Call HH(a, b, c, d, x(13),  4, &H289B7EC6) ' 41 S31
    Call HH(d, a, b, c, x( 0), 11, &HEAA127FA) ' 42 S32
    Call HH(c, d, a, b, x( 3), 16, &HD4EF3085) ' 43 S33
    Call HH(b, c, d, a, x( 6), 23,  &H4881D05) ' 44 S34
    Call HH(a, b, c, d, x( 9),  4, &HD9D4D039) ' 45 S31
    Call HH(d, a, b, c, x(12), 11, &HE6DB99E5) ' 46 S32
    Call HH(c, d, a, b, x(15), 16, &H1FA27CF8) ' 47 S33
    Call HH(b, c, d, a, x( 2), 23, &HC4AC5665) ' 48 S34

    ' Round 4
    Call II(a, b, c, d, x( 0),  6, &HF4292244) ' 49 S41
    Call II(d, a, b, c, x( 7), 10, &H432AFF97) ' 50 S42
    Call II(c, d, a, b, x(14), 15, &HAB9423A7) ' 51 S43
    Call II(b, c, d, a, x( 5), 21, &HFC93A039) ' 52 S44
    Call II(a, b, c, d, x(12),  6, &H655B59C3) ' 53 S41
    Call II(d, a, b, c, x( 3), 10, &H8F0CCC92) ' 54 S42
    Call II(c, d, a, b, x(10), 15, &HFFEFF47D) ' 55 S43
    Call II(b, c, d, a, x( 1), 21, &H85845DD1) ' 56 S44
    Call II(a, b, c, d, x( 8),  6, &H6FA87E4F) ' 57 S41
    Call II(d, a, b, c, x(15), 10, &HFE2CE6E0) ' 58 S42
    Call II(c, d, a, b, x( 6), 15, &HA3014314) ' 59 S43
    Call II(b, c, d, a, x(13), 21, &H4E0811A1) ' 60 S44
    Call II(a, b, c, d, x( 4),  6, &HF7537E82) ' 61 S41
    Call II(d, a, b, c, x(11), 10, &HBD3AF235) ' 62 S42
    Call II(c, d, a, b, x( 2), 15, &H2AD7D2BB) ' 63 S43
    Call II(b, c, d, a, x( 9), 21, &HEB86D391) ' 64 S44

    state(0) = add(state(0), a)
    state(1) = add(state(1), b)
    state(2) = add(state(2), c)
    state(3) = add(state(3), d)

    Dim i
    For i = 0 To UBound(x)
        x(i) = 0
    Next
End Sub

Private Sub Encode(ByRef output, ByRef inputx, ByVal lenx)
    Dim i
    i = 0
    Dim j
    j = 0
    Do While j < lenx
        output(j) = inputx(i) And &HFF
        output(j + 1) = sr(inputx(i), 8) And &HFF
        output(j + 2) = sr(inputx(i), 16) And &HFF
        output(j + 3) = sr(inputx(i), 24) And &HFF
        i = i + 1
        j = j + 4
    Loop
End Sub

Private Sub Decode(ByRef output, ByRef inputx, ByVal inputxOffset, ByVal lenx)
    Dim i
    i = 0
    Dim j
    j = 0
    Do While j < lenx
        Dim k
        k = j + inputxOffset
        output(i) = inputx(k) Or sl(inputx(k + 1), 8) Or sl(inputx(k + 2), 16) Or sl(inputx(k + 3), 24)
        i = i + 1
        j = j + 4
    Loop
End Sub

Private Sub MD5_memcpy(ByRef output, ByVal outputOffset, ByRef inputx, ByVal inputxOffset, ByVal lenx)
    Dim i
    For i = 0 To lenx - 1
        output(i + outputOffset) = inputx(i + inputxOffset)
    Next
End Sub

Private Function MDPrint(ByRef digest)
    Dim s
    s = ""
    Dim i
    For i = 0 To 16 - 1
        s = s & Right("00" & LCase(Hex(digest(i))), 2)
    Next
    MDPrint = s
End Function
