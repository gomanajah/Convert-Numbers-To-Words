Module Module1
    Function ConvertNumberToWords3(Number As Double, MainCurrency As String, SubCurrency As String)
        Dim MyArry1(0 To 9) As String
        Dim MyArry2(0 To 9) As String
        Dim MyArry3(0 To 9) As String
        Dim Myno As String
        Dim GetNo As String
        Dim RdNo As String
        Dim My100 As String
        Dim My10 As String
        Dim My1 As String
        Dim My11 As String
        Dim My12 As String
        Dim GetTxt As String
        Dim Mybillion As String
        Dim MyMillion As String
        Dim MyThou As String
        Dim MyHun As String
        Dim MyFraction As String
        Dim MyAnd As String
        Dim I As Integer
        Dim ReMark As String
        If Number > 999999999999.999 Then Exit Function
        If Number < 0 Then
            Number = Number * -1
            ReMark = "سالب "
        End If
        If Number = 0 Then
            ConvertNumberToWords3 = "صفر"
            Exit Function
        End If
        MyAnd = " و"
        MyArry1(0) = ""
        MyArry1(1) = "مائة"
        MyArry1(2) = "مائتان"
        MyArry1(3) = "ثلاثمائة"
        MyArry1(4) = "أربعمائة"
        MyArry1(5) = "خمسمائة"
        MyArry1(6) = "ستمائة"
        MyArry1(7) = "سبعمائة"
        MyArry1(8) = "ثمانمائة"
        MyArry1(9) = "تسعمائة"
        MyArry2(0) = ""
        MyArry2(1) = " عشر"
        MyArry2(2) = "عشرون"
        MyArry2(3) = "ثلاثون"
        MyArry2(4) = "أربعون"
        MyArry2(5) = "خمسون"
        MyArry2(6) = "ستون"
        MyArry2(7) = "سبعون"
        MyArry2(8) = "ثمانون"
        MyArry2(9) = "تسعون"
        MyArry3(0) = ""
        MyArry3(1) = "واحد"
        MyArry3(2) = "اثنان"
        MyArry3(3) = "ثلاثة"
        MyArry3(4) = "أربعة"
        MyArry3(5) = "خمسة"
        MyArry3(6) = "ستة"
        MyArry3(7) = "سبعة"
        MyArry3(8) = "ثمانية"
        MyArry3(9) = "تسعة"
        GetNo = Math.Round(Number, 3)
        GetNo = Format(Number, "000000000000.000")
        I = 0
        Do While I < 16
            If I < 12 Then
                Myno = Mid$(GetNo, I + 1, 3)
            Else
                Myno = Mid$(GetNo, I + 2, 3) + "0"
            End If
            If (Mid$(Myno, 1, 3)) > 0 Then
                RdNo = Mid$(Myno, 1, 1)
                My100 = MyArry1(RdNo)
                RdNo = Mid$(Myno, 3, 1)
                My1 = MyArry3(RdNo)
                RdNo = Mid$(Myno, 2, 1)
                My10 = MyArry2(RdNo)
                If Mid$(Myno, 2, 2) = 11 Then My11 = "إحدى عشر"
                If Mid$(Myno, 2, 2) = 12 Then My12 = "إثنى عشر"
                If Mid$(Myno, 2, 2) = 10 Then My10 = "عشرة"
                If ((Mid$(Myno, 1, 1)) > 0) And ((Mid$(Myno, 2, 2)) > 0) Then My100 = My100 + MyAnd
                If ((Mid$(Myno, 3, 1)) > 0) And ((Mid$(Myno, 2, 1)) > 1) Then My1 = My1 + MyAnd
                GetTxt = My100 + My1 + My10
                If ((Mid$(Myno, 3, 1)) = 1) And ((Mid$(Myno, 2, 1)) = 1) Then
                    GetTxt = My100 + My11
                    If ((Mid$(Myno, 1, 1)) = 0) Then GetTxt = My11
                End If
                If ((Mid$(Myno, 3, 1)) = 2) And ((Mid$(Myno, 2, 1)) = 1) Then
                    GetTxt = My100 + My12
                    If ((Mid$(Myno, 1, 1)) = 0) Then GetTxt = My12
                End If
                If (I = 0) And (GetTxt <> "") Then
                    If ((Mid$(Myno, 1, 3)) > 10) Then
                        Mybillion = GetTxt + " مليار"
                    Else
                        Mybillion = GetTxt + " مليارات"
                        If ((Mid$(Myno, 1, 3)) = 2) Then Mybillion = " مليار"
                        If ((Mid$(Myno, 1, 3)) = 2) Then Mybillion = " ملياران"
                    End If
                End If
                If (I = 3) And (GetTxt <> "") Then
                    If ((Mid$(Myno, 1, 3)) > 10) Then
                        MyMillion = GetTxt + " مليون"
                    Else
                        MyMillion = GetTxt + " ملايين"
                        If ((Mid$(Myno, 1, 3)) = 1) Then MyMillion = " مليون"
                        If ((Mid$(Myno, 1, 3)) = 2) Then MyMillion = " مليونان"
                    End If
                End If
                If (I = 6) And (GetTxt <> "") Then
                    If ((Mid$(Myno, 1, 3)) > 10) Then
                        MyThou = GetTxt + " ألف"
                    Else
                        MyThou = GetTxt + " آلاف"
                        If ((Mid$(Myno, 3, 1)) = 1) Then MyThou = " ألف"
                        If ((Mid$(Myno, 3, 1)) = 2) Then MyThou = " ألفان"
                    End If
                End If
                If (I = 9) And (GetTxt <> "") Then MyHun = GetTxt
                If (I = 12) And (GetTxt <> "") Then MyFraction = GetTxt
            End If
            I = I + 3
        Loop
        If (Mybillion <> "") Then
            If (MyMillion <> "") Or (MyThou <> "") Or (MyHun <> "") Then Mybillion = Mybillion + MyAnd
        End If
        If (MyMillion <> "") Then
            If (MyThou <> "") Or (MyHun <> "") Then MyMillion = MyMillion + MyAnd
        End If
        If (MyThou <> "") Then
            If (MyHun <> "") Then MyThou = MyThou + MyAnd
        End If
        If MyFraction <> "" Then
            If (Mybillion <> "") Or (MyMillion <> "") Or (MyThou <> "") Or (MyHun <> "") Then
                ConvertNumberToWords3 = ReMark + Mybillion + MyMillion + MyThou + MyHun + " " + MainCurrency + MyAnd + MyFraction + " " + SubCurrency
            Else
                ConvertNumberToWords3 = ReMark + MyFraction + " " + SubCurrency
            End If
        Else
            ConvertNumberToWords3 = ReMark + Mybillion + MyMillion + MyThou + MyHun + " " + MainCurrency
        End If
    End Function
    Function ConvertNumberToWords2(Number As Double, MainCurrency As String, SubCurrency As String)
        Dim Array1(0 To 9) As String
        Dim Array2(0 To 9) As String
        Dim Array3(0 To 9) As String
        Dim MyNumber As String
        Dim GetNumber As String
        Dim ReadNumber As String
        Dim My100 As String
        Dim My10 As String
        Dim My1 As String
        Dim My11 As String
        Dim My12 As String
        Dim GetText As String
        Dim Billion As String
        Dim Million As String
        Dim Thousand As String
        Dim Hundred As String
        Dim Fraction As String
        Dim MyAnd As String
        Dim I As Integer
        Dim ReMark As String
        If Number > 999999999999.99 Then Exit Function
        If Number < 0 Then
            Number = Number * -1
            ReMark = "سالب "
        End If
        If Number = 0 Then
            ConvertNumberToWords2 = "صفر"
            Exit Function
        End If
        MyAnd = " و"
        Array1(0) = ""
        Array1(1) = "مائة"
        Array1(2) = "مائتان"
        Array1(3) = "ثلاثمائة"
        Array1(4) = "أربعمائة"
        Array1(5) = "خمسمائة"
        Array1(6) = "ستمائة"
        Array1(7) = "سبعمائة"
        Array1(8) = "ثمانمائة"
        Array1(9) = "تسعمائة"
        Array2(0) = ""
        Array2(1) = " عشر"
        Array2(2) = "عشرون"
        Array2(3) = "ثلاثون"
        Array2(4) = "أربعون"
        Array2(5) = "خمسون"
        Array2(6) = "ستون"
        Array2(7) = "سبعون"
        Array2(8) = "ثمانون"
        Array2(9) = "تسعون"
        Array3(0) = ""
        Array3(1) = "واحد"
        Array3(2) = "اثنان"
        Array3(3) = "ثلاثة"
        Array3(4) = "أربعة"
        Array3(5) = "خمسة"
        Array3(6) = "ستة"
        Array3(7) = "سبعة"
        Array3(8) = "ثمانية"
        Array3(9) = "تسعة"
        GetNumber = Format(Number, "000000000000.00")
        I = 0
        Do While I < 15
            If I < 12 Then
                MyNumber = Mid$(GetNumber, I + 1, 3)
            Else
                MyNumber = "0" + Mid$(GetNumber, I + 2, 2)
            End If
            If (Mid$(MyNumber, 1, 3)) > 0 Then
                ReadNumber = Mid$(MyNumber, 1, 1)
                My100 = Array1(ReadNumber)
                ReadNumber = Mid$(MyNumber, 3, 1)
                My1 = Array3(ReadNumber)
                ReadNumber = Mid$(MyNumber, 2, 1)
                My10 = Array2(ReadNumber)
                If Mid$(MyNumber, 2, 2) = 11 Then My11 = "إحدى عشر"
                If Mid$(MyNumber, 2, 2) = 12 Then My12 = "إثنى عشر"
                If Mid$(MyNumber, 2, 2) = 10 Then My10 = "عشرة"
                If ((Mid$(MyNumber, 1, 1)) > 0) And ((Mid$(MyNumber, 2, 2)) > 0) Then My100 = My100 + MyAnd
                If ((Mid$(MyNumber, 3, 1)) > 0) And ((Mid$(MyNumber, 2, 1)) > 1) Then My1 = My1 + MyAnd
                GetText = My100 + My1 + My10
                If ((Mid$(MyNumber, 3, 1)) = 1) And ((Mid$(MyNumber, 2, 1)) = 1) Then
                    GetText = My100 + My11
                    If ((Mid$(MyNumber, 1, 1)) = 0) Then GetText = My11
                End If
                If ((Mid$(MyNumber, 3, 1)) = 2) And ((Mid$(MyNumber, 2, 1)) = 1) Then
                    GetText = My100 + My12
                    If ((Mid$(MyNumber, 1, 1)) = 0) Then GetText = My12
                End If
                If (I = 0) And (GetText <> "") Then
                    If ((Mid$(MyNumber, 1, 3)) > 10) Then
                        Billion = GetText + " مليار"
                    Else
                        Billion = GetText + " مليارات"
                        If ((Mid$(MyNumber, 1, 3)) = 2) Then Billion = " مليار"
                        If ((Mid$(MyNumber, 1, 3)) = 2) Then Billion = " ملياران"
                    End If
                End If
                If (I = 3) And (GetText <> "") Then
                    If ((Mid$(MyNumber, 1, 3)) > 10) Then
                        Million = GetText + " مليون"
                    Else
                        Million = GetText + " ملايين"
                        If ((Mid$(MyNumber, 1, 3)) = 1) Then Million = " مليون"
                        If ((Mid$(MyNumber, 1, 3)) = 2) Then Million = " مليونان"
                    End If
                End If
                If (I = 6) And (GetText <> "") Then
                    If ((Mid$(MyNumber, 1, 3)) > 10) Then
                        Thousand = GetText + " ألف"
                    Else
                        Thousand = GetText + " ألاف"
                        If ((Mid$(MyNumber, 3, 1)) = 1) Then Thousand = " ألف"
                        If ((Mid$(MyNumber, 3, 1)) = 2) Then Thousand = " ألفان"
                    End If
                End If
                If (I = 9) And (GetText <> "") Then Hundred = GetText
                If (I = 12) And (GetText <> "") Then Fraction = GetText
            End If
            I = I + 3
        Loop
        If (Billion <> "") Then
            If (Million <> "") Or (Thousand <> "") Or (Hundred <> "") Then Billion = Billion + MyAnd
        End If
        If (Million <> "") Then
            If (Thousand <> "") Or (Hundred <> "") Then Million = Million + MyAnd
        End If
        If (Thousand <> "") Then
            If (Hundred <> "") Then Thousand = Thousand + MyAnd
        End If
        If Fraction <> "" Then
            If (Billion <> "") Or (Million <> "") Or (Thousand <> "") Or (Hundred <> "") Then
                ConvertNumberToWords2 = ReMark + Billion + Million + Thousand + Hundred + " " + MainCurrency + MyAnd + Fraction + " " + SubCurrency
            Else
                ConvertNumberToWords2 = ReMark + Fraction + " " + SubCurrency
            End If
        Else
            ConvertNumberToWords2 = ReMark + Billion + Million + Thousand + Hundred + " " + MainCurrency
        End If
    End Function
    Sub JustNumbers(ByVal toolbox, ByVal e)
        If Char.IsControl(e.KeyChar) = False Then
            If Char.IsDigit(e.KeyChar) Or e.KeyChar = "." Then
                If toolbox.Text.Contains(".") Then
                    If toolbox.Text.Split(".")(1).Length < 3 Then
                        If Char.IsDigit(e.KeyChar) = False Then e.Handled = True
                    Else
                        e.Handled = True
                    End If
                End If
            Else
                e.Handled = True
            End If
        Else
            toolbox.text = ""
        End If
    End Sub
    Sub JustNumbers1(ByVal toolbox, ByVal e)
        If Char.IsControl(e.KeyChar) = False Then
            If Char.IsDigit(e.KeyChar) Or e.KeyChar = "." Then
                If toolbox.Text.Contains(".") Then
                    If toolbox.Text.Split(".")(1).Length < 2 Then
                        If Char.IsDigit(e.KeyChar) = False Then e.Handled = True
                    Else
                        e.Handled = True
                    End If
                End If
            Else
                e.Handled = True
            End If
        End If
    End Sub
End Module
