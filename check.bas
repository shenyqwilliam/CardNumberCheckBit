'大陆18位身份证号校验
Public Function IdCard_CN(ByVal cardNo$) As Boolean
    Dim lastBit$, checkSum&
    checkSum = 0
    
    If Len(cardNo) <> 18 Then Exit Function
    
    On Error GoTo err
    
    For i = 1 To 17
        checkSum = (checkSum + CInt(Mid(cardNo, i, 1))) * 2
    Next i
    
    lastBit = Right(cardNo, 1)
    If lastBit = "X" Or lastBit = "x" Then checkSum = checkSum + 10 Else checkSum = checkSum + CInt(lastBit)
    
    If checkSum Mod 11 = 1 Then IdCard_CN = True: Exit Function
    
err:    Exit Function
End Function

'香港身份证号校验
Public Function IdCard_HK(ByVal cardNo$) As Boolean
    Dim bits%(7), leftBracket%, rightBracket%, lastBit$, checkSum%
    checkSum = 0
    
    If Len(cardNo) <> 10 Then Exit Function
    
    leftBracket = Asc(Mid(cardNo, 8))
    rightBracket = Asc(Mid(cardNo, 10))
    If (leftBracket <> 40 And leftBracket <> -23640) Then GoTo err
    If (rightBracket <> -23639 And rightBracket <> 41) Then GoTo err
    
    On Error GoTo err
    
    bits(0) = Asc(UCase(Left(cardNo, 1))): If bits(0) >= 65 And bits(0) <= 90 Then bits(0) = bits(0) - 64 Else GoTo err
    For i = 1 To 6
        bits(i) = CInt(Mid(cardNo, i + 1, 1))
    Next i
    lastBit = Mid(cardNo, 9, 1)
    If lastBit = "A" Or lastBit = "a" Then bits(7) = 10 Else bits(7) = CInt(lastBit)
    
    For i = 0 To 7
        checkSum = checkSum + bits(i) * (8 - i)
    Next i
    
    If checkSum Mod 11 = 0 Then IdCard_HK = True: Exit Function
    
err:    Exit Function
End Function

'澳门身份证号校验
Public Function IdCard_MC(ByVal cardNo$) As Boolean
    Dim bits%(7), leftBracket%, rightBracket%
    
    If Len(cardNo) <> 10 Then Exit Function
    
    leftBracket = Asc(Mid(cardNo, 8))
    rightBracket = Asc(Mid(cardNo, 10))
    If (leftBracket <> 40 And leftBracket <> -23640) Then GoTo err
    If (rightBracket <> -23639 And rightBracket <> 41) Then GoTo err
    
    On Error GoTo err
    
    For i = 0 To 6
        bits(i) = CInt(Mid(cardNo, i + 1, 1))
    Next i
    bits(7) = CInt(Mid(cardNo, 9, 1))
    
    IdCard_MC = Luhn(bits())
    
err:        Exit Function
End Function

'台湾身份证号校验
Public Function IdCard_TW(ByVal cardNo$) As Boolean
    Dim firstBit%, checkSum%
    checkSum = 0
    
    If Len(cardNo) <> 10 Then Exit Function
    
    On Error GoTo err
    
    firstBit = Asc(UCase(Left(cardNo, 1)))
    Select Case firstBit
        Case Is < 65, Is > 90
            Exit Function
        Case Is < 73
            firstBit = firstBit - 55
        Case 73
            firstBit = 34
        Case Is < 79
            firstBit = firstBit - 56
        Case 79
            firstBit = 35
        Case Else
            firstBit = firstBit - 57
    End Select
    checkSum = firstBit \ 10 + (firstBit Mod 10) * 9
    
    For i = 2 To 9
        checkSum = checkSum + CInt(Mid(cardNo, i, 1)) * (10 - i)
    Next i
    checkSum = checkSum + CInt(Right(cardNo, 1))
    
    If checkSum Mod 10 = 0 Then IdCard_TW = True
    
err:            Exit Function
End Function


'银行卡号校验
Public Function BankCard(ByVal cardNo$) As Boolean
    Dim bits%()
    
    If Not IsNumeric(cardNo) Then Exit Function
    
    ReDim Preserve bits(Len(cardNo) - 1)
    For i = 0 To UBound(bits())
        bits(i) = CInt(Mid(cardNo, i + 1, 1))
    Next i
    
    BankCard = Luhn(bits())
End Function

'Luhn算法
Public Function Luhn(ByRef bits%()) As Boolean
    Dim checkSum%: checkSum = 0
    
    For i = UBound(bits()) To 0 Step -2
        checkSum = checkSum + bits(i)
        If i <> 0 Then
            If bits(i - 1) > 4 Then
                checkSum = checkSum + (bits(i - 1) * 2) - 9
            Else
                checkSum = checkSum + bits(i - 1) * 2
            End If
        End If
    Next i
    
    If checkSum Mod 10 = 0 Then Luhn = True
End Function

