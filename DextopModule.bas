Attribute VB_Name = "DextopModule"
Option Explicit

Public db As Database
Public sCustomerAccountParentID As String
Public sSupplierAccountParentID As String
Public sGeneralCustomerAccountID As String
Public sGeneralSupplierAccountID As String
Public sCurrentUserCode As String
Public sCurrentUsername As String
Public sCurrentFinancialCode As String
Public dCurrentFinancialFromDate As Date
Public dCurrentFinancialToDate As Date
Public sCashAccount As String
Public sAsset As String
Public sLiabilities As String
Public sIncome         As String
Public sExpense   As String
Public sDirectIncome As String
Public sIndirectIncome         As String
Public sDirectExpense As String
Public sIndirectExpense As String
Public sPurchaseAccount As String
Public sSaleAccount As String
Public sStaffAccountParentID As String
Public sStockInHand         As String
Public sSaleDiscounts         As String
Public sPurchaseDiscounts As String
Public sProfitnLossAccount         As String
Public sSaleTax As String
Public sPurchaseReturnAccount As String
Public sSaleReturnAccount As String
Public sSaleReturnDiscounts         As String
Public sPurchaseReturnDiscounts As String
Public sReceivableAccount As String
Public sPayableAccount As String
Public sSalesForm8 As String
Public sSalesForm8B As String
Public sBankGroupCode As String

 
Public Sub initialisePublicVariables()
    Set db = OpenDatabase(App.Path & "\Storage.mdb", False, False, "MS Access;PWD=12345abcde")
End Sub

Public Function isLoadingFirstTime() As Boolean
    Dim bFirst As Boolean, sRead As String
    bFirst = False
    
    Open App.Path & "\Load.inf" For Input As #1
    Line Input #1, sRead
    Close #1
    If (sRead = "Load=First") Then
        Open App.Path & "\Load.inf" For Output As #1
        Print #1, "Not First"
        Close #1
        bFirst = True
    End If
    isLoadingFirstTime = bFirst
End Function

Public Sub arrangeFoldersAndFiles(bArrange As Boolean)
On Error GoTo Out
    If (bArrange = False) Then
        Exit Sub
    End If
    'Creates the Reports Folder in Application Path
    MkDir App.Path & "\Reports"
Out:
End Sub

Public Function getUniversaloFor(sPurchaseRate As String) As String
Dim tempPurchase As String, sCode As String
        tempPurchase = sPurchaseRate
        sCode = ""
        While Len(tempPurchase) > 0
            sCode = sCode & getCode(Left(tempPurchase, 1))
            tempPurchase = Right(tempPurchase, Len(tempPurchase) - 1)
        Wend
        getUniversaloFor = sCode
End Function

Private Function getCode(sChar As String) As String
    Select Case sChar
        Case "0"
            getCode = "O"
            Exit Function
        Case "1"
            getCode = "U"
            Exit Function
        Case "2"
            getCode = "N"
            Exit Function
        Case "3"
            getCode = "I"
            Exit Function
        Case "4"
            getCode = "V"
            Exit Function
        Case "5"
            getCode = "E"
            Exit Function
        Case "6"
            getCode = "R"
            Exit Function
        Case "7"
            getCode = "S"
            Exit Function
        Case "8"
            getCode = "A"
            Exit Function
        Case "9"
            getCode = "L"
            Exit Function
        Case "."
            getCode = "P"
            Exit Function
    End Select
End Function

Public Function getCentralAlignmentStartingPos(lPrintWidth As Long, sWord As String) As Long
Dim dPos As Long, lWordLen As Long
    lWordLen = Len(sWord)
    dPos = (lPrintWidth / 2) - (lWordLen / 2)
    getCentralAlignmentStartingPos = dPos
End Function


Public Function getFinancialCode(dDate As Date) As Long
    getFinancialCode = IIf(dDate >= DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), Year(dDate), Year(dDate) - 1)
End Function

Public Function getFinancialStartDate(dDate As Date) As Date
    getFinancialStartDate = IIf(dDate >= DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & (Year(dDate) - 1)))
End Function

Public Function getFinancialEndDate(dDate As Date) As Date
    getFinancialEndDate = IIf(dDate >= DateValue(Day(dCurrentFinancialFromDate) & "," & Month(dCurrentFinancialFromDate) & "," & Year(dDate)), DateValue(Day(dCurrentFinancialFromDate - 1) & "," & Month(dCurrentFinancialFromDate - 1) & "," & Year(dDate) + 1), DateValue(Day(dCurrentFinancialFromDate - 1) & "," & Month(dCurrentFinancialFromDate - 1) & "," & Year(dDate)))
End Function

Public Sub setDefaultParentIDAndAccountCode()
Dim rs As Recordset
    
    Set rs = db.OpenRecordset("Select * From AccountRegister Where (AccountRegister.FixedAccount<>'')")
    
    While rs.EOF = False
        sCustomerAccountParentID = IIf(rs!FixedAccount = "Sundry Debtors", rs!Code, sCustomerAccountParentID)
        sSupplierAccountParentID = IIf(rs!FixedAccount = "Sundry Creditors", rs!Code, sSupplierAccountParentID)
        sGeneralCustomerAccountID = IIf(rs!FixedAccount = "General Customer", rs!Code, sGeneralCustomerAccountID)
        sGeneralSupplierAccountID = IIf(rs!FixedAccount = "General Supplier", rs!Code, sGeneralSupplierAccountID)
        sCashAccount = IIf(rs!FixedAccount = "Cash", rs!Code, sCashAccount)
        sAsset = IIf(rs!FixedAccount = "Asset", rs!Code, sAsset)
        sLiabilities = IIf(rs!FixedAccount = "Liabilities", rs!Code, sLiabilities)
        sIncome = IIf(rs!FixedAccount = "Income", rs!Code, sIncome)
        sExpense = IIf(rs!FixedAccount = "Expense", rs!Code, sExpense)
        sDirectIncome = IIf(rs!FixedAccount = "Direct Income", rs!Code, sDirectIncome)
        sIndirectIncome = IIf(rs!FixedAccount = "Indirect Income", rs!Code, sIndirectIncome)
        sDirectExpense = IIf(rs!FixedAccount = "Direct Expense", rs!Code, sDirectExpense)
        sIndirectExpense = IIf(rs!FixedAccount = "Indirect Expense", rs!Code, sIndirectExpense)
        sPurchaseAccount = IIf(rs!FixedAccount = "Purchase Account", rs!Code, sPurchaseAccount)
        sSaleAccount = IIf(rs!FixedAccount = "Sale Account", rs!Code, sSaleAccount)
        sStaffAccountParentID = IIf(rs!FixedAccount = "Staff Accounts", rs!Code, sStaffAccountParentID)
        sStockInHand = IIf(rs!FixedAccount = "Stock In Hand", rs!Code, sStockInHand)
        sSaleDiscounts = IIf(rs!FixedAccount = "Sale Discounts", rs!Code, sSaleDiscounts)
        sPurchaseDiscounts = IIf(rs!FixedAccount = "Purchase Discounts", rs!Code, sPurchaseDiscounts)
        sProfitnLossAccount = IIf(rs!FixedAccount = "Profit & Loss Account", rs!Code, sProfitnLossAccount)
        sPurchaseReturnAccount = IIf(rs!FixedAccount = "Purchase Return Account", rs!Code, sPurchaseReturnAccount)
        sSaleReturnAccount = IIf(rs!FixedAccount = "Sale Return Account", rs!Code, sSaleReturnAccount)
        sSaleReturnDiscounts = IIf(rs!FixedAccount = "Sale Return Discounts", rs!Code, sSaleReturnDiscounts)
        sPurchaseReturnDiscounts = IIf(rs!FixedAccount = "Purchase Return Discounts", rs!Code, sPurchaseReturnDiscounts)
        sPayableAccount = IIf(rs!FixedAccount = "Payable", rs!Code, sPayableAccount)
        sReceivableAccount = IIf(rs!FixedAccount = "Receivable", rs!Code, sReceivableAccount)
        sSalesForm8 = IIf(rs!FixedAccount = "Sales Form 8", rs!Code, sSalesForm8)
        sSalesForm8B = IIf(rs!FixedAccount = "Sales Form 8B", rs!Code, sSalesForm8B)
        sBankGroupCode = IIf(rs!FixedAccount = "Bank Accounts", rs!Code, sBankGroupCode)
        rs.MoveNext
    Wend
    rs.Close
End Sub

Public Function isBillDateInFinancialYear(dTransactionDate As Date) As Boolean
Dim rs As Recordset
Dim isNotCorrect As Boolean

    If dTransactionDate < dCurrentFinancialFromDate Or dTransactionDate > dCurrentFinancialToDate Then
        isNotCorrect = False
    Else
        isNotCorrect = True
    End If
    isBillDateInFinancialYear = isNotCorrect
End Function

Public Function getGCodeOfAccount(sAccountCode As String) As String
Dim rs As Recordset
Dim sCode As String
    Set rs = db.OpenRecordset("Select AccountRegister.GroupCode From AccountRegister Where(AccountRegister.Code='" & sAccountCode & "' )")
    If rs.RecordCount > 0 Then
        sCode = "" & rs!GroupCode
    Else
        sCode = ""
    End If
    getGCodeOfAccount = sCode
End Function


Public Function getNewAccountcode() As String
Dim rs As Recordset, sAccountCode As String
    Set rs = db.OpenRecordset("Select Max(val(AccountRegister.Code))As ACode From AccountRegister")
    If rs.RecordCount > 0 Then
        sAccountCode = Val("" & rs!ACode) + 1
    Else
        sAccountCode = "1"
    
    End If
    rs.Close
    
    getNewAccountcode = sAccountCode
End Function

Public Function printGrid(gData As MSFlexGrid, sHeader() As String, sReportOn As String) As Boolean
'On Error GoTo GoOut
Dim x As Long, y As Long, r As Long, c As Long
Dim colCount As Long, colLength() As Long, lTotallength As Long, tempColLength() As Long
Dim dPercValOfWidth As Double, sPrintData As String
    
    colCount = gData.Cols
    ReDim colLength(colCount) As Long
    
    'SETTING MINIMUM LENGTH
    c = 0
    While c < colCount
        colLength(c) = Len(sHeader(c))
        c = c + 1
    Wend
    
    r = 0
    While r < gData.Rows
        c = 0
        While c < colCount
            If (Printer.TextWidth(gData.TextMatrix(r, c)) > colLength(c)) Then
                colLength(c) = Printer.TextWidth(gData.TextMatrix(r, c))
            End If
            c = c + 1
        Wend
        r = r + 1
    Wend
    
    'GETS TOTAL LENGTH
    lTotallength = 0
    c = 0
    While c < colCount
        lTotallength = lTotallength + colLength(c)
        c = c + 1
    Wend
    tempColLength = colLength
    
    'CALCULATES WIDTH FOR EACH COL
    dPercValOfWidth = (9600 / lTotallength) * 100 '11000-500=9600
    c = 0
    While c < colCount
        tempColLength(c) = ((tempColLength(c) * dPercValOfWidth) / 100)
        If c = 0 Then
            colLength(c) = 500 + ((colLength(c) * dPercValOfWidth) / 100)
        ElseIf c > 0 Then
            colLength(c) = 100 + (colLength(c - 1) + ((colLength(c) * dPercValOfWidth) / 100))
        End If
        c = c + 1
    Wend
    
    header sHeader, tempColLength, colLength, colCount, sReportOn 'GETS HEADER PART
    
    x = 500
    y = 2700
    
    r = 0
    While r < gData.Rows
        If y >= 15600 Then
            Printer.EndDoc
            'NEXT  PAGE
            header sHeader, tempColLength, colLength, colCount, sReportOn 'GETS HEADER PART
            y = 2700
        End If
        c = 0
        Printer.CurrentX = 500
        Printer.CurrentY = y
        While c < colCount
            sPrintData = gData.TextMatrix(r, c)
            While Printer.TextWidth(sPrintData) >= (tempColLength(c))
                sPrintData = Left(sPrintData, Len(sPrintData) - 1)
            Wend
            Printer.Print sPrintData
            Printer.CurrentX = colLength(c)
            Printer.CurrentY = y
            c = c + 1
        Wend
        y = y + 300
        
        r = r + 1
    Wend
    
    Printer.EndDoc
    
    printGrid = True
    Exit Function
GoOut:
    printGrid = False
End Function

Public Function getCurrentBalanceOf(sAccountCode As String) As Double
Dim rs As Recordset
Dim dCurrentBalance As Double
    Set rs = db.OpenRecordset("Select (Sum(AccountTransaction.Debit)- Sum(AccountTransaction.Credit)) As Balance From AccountTransaction Where (AccountTransaction.AccountCode = '" & sAccountCode & "')")
    If rs.RecordCount > 0 Then
        dCurrentBalance = Val("" & rs!Balance)
    Else
        dCurrentBalance = 0
    End If
    getCurrentBalanceOf = dCurrentBalance
End Function

Private Sub header(sHeader() As String, colLength() As Long, colStart() As Long, colCount As Long, sReportOn As String)
'SEETING NEW PAGE
Dim i, j, x, y As Double, c As Long
Dim sPrintData As String

    Printer.ScaleMode = 1
    Printer.FontName = "Arial"
    
    Printer.FontBold = True
    Printer.FontUnderline = False
    Printer.FontSize = 20
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("DYNAMIC")) / 2)
    Printer.CurrentY = 400
    Printer.Print "DEXTOP"
    y = 800
    
    Printer.FontSize = 12
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("DIGITAL SPOT")) / 2)
    Printer.CurrentY = y
    Printer.Print "SOFTWARE INNOVATIONS"
    
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Mayyeri Arcade, Thazhepalam, Tirur-1")) / 2)
    y = y + 300
    Printer.CurrentY = y
    Printer.Print "TIRUR"
    
    y = y + 200
    Printer.FontSize = 10
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Tel: 0494 2426886, Mob: 9072111195, Fax: 0494 3012234")) / 2)
    Printer.CurrentY = y
    Printer.Print "9633723993"
    
    y = y + 200
    Printer.FontSize = 10
    Printer.FontItalic = True
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("E-mail: info@dynamicdigitalspot.com, dynamiclaserprint@gmail.com")) / 2)
    Printer.CurrentY = y
    Printer.Print "E-mail: DEXTOPSOFWARE@gmail.com"
    Printer.FontItalic = False
    ''''''ONE MORE WEBSITE SHOULD BE SHOWN HERE
    y = y + 400
    Printer.FontUnderline = True
    Printer.FontSize = 12
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth(sReportOn)) / 2)
    Printer.CurrentY = y
    Printer.Print sReportOn
    Printer.FontUnderline = False
    Printer.FontSize = 10
    
    x = 500
    y = y + 500
    'y=2000
    
    'HEADING FOR THE ROWS
    c = 0
    Printer.FontBold = True
    Printer.CurrentX = x
    Printer.CurrentY = y
    While c < colCount
        sPrintData = sHeader(c)
        While Printer.TextWidth(sPrintData) >= (colLength(c))
            sPrintData = Left(sPrintData, Len(sPrintData) - 1)
        Wend
        Printer.Print sPrintData
        Printer.CurrentX = colStart(c)
        Printer.CurrentY = y
        c = c + 1
    Wend
    Printer.FontBold = False
End Sub

Public Function NumberToWords(number As Double) As String  'maximum number is Hundred Crore
Dim snumber As String, numberpart As Double, decimalpart As Integer, length As Long
Dim snumberarray() As String, i As Integer, temp As String, x As Variant, snumberpart As String
    
    If Len(FormatNumber(number, 0, , , vbFalse)) > 10 Then 'checks the limit
        
        Exit Function
    End If
    
    
    'splitting to  to parts (integer and decimal)
    snumber = FormatNumber(number, 2, , , vbFalse)
    length = InStr(1, snumber, ".")
    If length = 0 Then
        
        numberpart = Val(snumber)
        decimalpart = 0
        snumberpart = snumber
    Else
        
        numberpart = Val(VBA.Strings.Left(snumber, length - 1))
        decimalpart = Val(Trim(VBA.Strings.Right(snumber, Len(snumber) - length)))
        'for removing paise part and rounding to number part
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If decimalpart >= 50 Then
        '
        '    numberpart = numberpart + 1
        '    decimalpart = 0
        'Else
        '
        '    decimalpart = 0
        'End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        snumberpart = numberpart
    End If
    
    'working with number part
    'split  number part
    
    If Len(snumberpart) <= 3 Then 'declaring number array string
        
        ReDim snumberarray(0)
    ElseIf Len(snumberpart) <= 5 Then
        
        ReDim snumberarray(1)
    ElseIf Len(snumberpart) <= 7 Then
        
        ReDim snumberarray(2)
    Else
        
        ReDim snumberarray(3)
    End If
    
    i = 1
    temp = ""
    snumberpart = StrReverse(snumberpart)
    While i <= Len(snumberpart)
    
        If i <= 3 Then 'dividing to groups
        
            temp = temp & VBA.Mid(snumberpart, i, 1)
            If i = Len(snumberpart) Or i = 3 Then
                
                snumberarray(0) = StrReverse(temp)
                temp = ""
            End If
        ElseIf i <= 5 Then
            
            temp = temp & VBA.Mid(snumberpart, i, 1)
            If i = Len(snumberpart) Or i = 5 Then
                
                snumberarray(1) = StrReverse(temp)
                temp = ""
            End If
        ElseIf i <= 7 Then
            
            temp = temp & VBA.Mid(snumberpart, i, 1)
            If i = Len(snumberpart) Or i = 7 Then
                
                snumberarray(2) = StrReverse(temp)
                temp = ""
            End If
        Else
            
            temp = temp & VBA.Mid(snumberpart, i, 1)
            If i = Len(snumberpart) Then
                
                snumberarray(3) = StrReverse(temp)
                temp = ""
            End If
        End If
        
        
        i = i + 1
    Wend
    'changing group numbers to words
    temp = ""
    i = 1
    For Each x In snumberarray
        
        If i = 1 Then
            
            temp = WordOf(Val(x))
        ElseIf i = 2 Then
            
            temp = WordOf(Val(x)) & " Thousand " & temp
        ElseIf i = 3 Then
            
            temp = WordOf(Val(x)) & " Lakh " & temp
        ElseIf i = 4 Then
            
            temp = WordOf(Val(x)) & " Crore " & temp
        End If
        i = i + 1
    Next
    'setting rupees
    If number = 1 Then

        temp = " Rupee " & temp
    ElseIf number = 0 Then
        
        temp = ""
    ElseIf number > 1 Then
        
        temp = " Rupees " & temp
    End If
    
    'working with decimal part
    If decimalpart <> 0 Then
        
        temp = temp & " And" & WordOf(decimalpart) & " Paise"
    End If
    'returinig the value
    NumberToWords = temp
   
End Function

Private Function WordOf(number As Integer) As String
Dim i As Integer, snumber As String, x As Integer, word As String, firstnumber As Integer
    
    i = 1
    snumber = Str(number)
    word = ""
    While i <= Len(snumber)
        
        x = number Mod 10
        number = number \ 10
        Select Case i
            
            Case 1:
                Select Case x

                    Case 1: word = " One" & word
                    Case 2: word = " Two" & word
                    Case 3: word = " Three" & word
                    Case 4: word = " Four" & word
                    Case 5: word = " Five" & word
                    Case 6: word = " Six" & word
                    Case 7: word = " Seven" & word
                    Case 8: word = " Eight" & word
                    Case 9: word = " Nine" & word
                End Select
                firstnumber = x
            Case 2:
                Select Case x
                    
                    Case 1:
                    
                            If firstnumber = 0 Then
                                
                                word = " Ten"
                            ElseIf firstnumber = 1 Then
                                
                                word = " Eleven"
                            ElseIf firstnumber = 2 Then
                                
                                word = " Twelve"
                            ElseIf firstnumber = 3 Then
                                
                                word = " Thirteen"
                            ElseIf firstnumber = 4 Then
                                
                                word = " Forteen"
                            ElseIf firstnumber = 5 Then
                                
                                word = " Fifteen"
                            ElseIf firstnumber = 6 Then
                                
                                word = " Sixteen"
                            ElseIf firstnumber = 7 Then
                                
                                word = " Seventeen"
                            ElseIf firstnumber = 8 Then
                                
                                word = " Eighteen"
                            ElseIf firstnumber = 9 Then
                                
                                word = " Nineteen"
                            End If
                    Case 2: word = " Twenty" & word
                    Case 3: word = " Thirty" & word
                    Case 4: word = " Forty" & word
                    Case 5: word = " Fifty" & word
                    Case 6: word = " Sixty" & word
                    Case 7: word = " Seventy" & word
                    Case 8: word = " Eighty" & word
                    Case 9: word = " Ninty" & word
                End Select
            Case 3:
                Select Case x
                    
                    Case 1: word = "One Hundred" & word
                    Case 2: word = "Two Hundred" & word
                    Case 3: word = "Three Hundred" & word
                    Case 4: word = "Four Hundred" & word
                    Case 5: word = "Five Hundred" & word
                    Case 6: word = "Six Hundred" & word
                    Case 7: word = "Seven Hundred" & word
                    Case 8: word = "Eight Hundred" & word
                    Case 9: word = "Nine Hundred" & word
                End Select
        End Select
        i = i + 1
    Wend
    WordOf = word
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                   BARCODE CODING START                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub saveBarCode(sCode As String)
Dim rs As Recordset
    Set rs = db.OpenRecordset("Select * From Barcode")
    If rs.RecordCount > 0 Then
        rs.Edit
        rs!Code = sCode
        rs.Update
    Else
        rs.AddNew
        rs!Code = "000000"
        rs.Update
    End If
    rs.Close
End Sub

Public Function getNewBarCode() As String
Dim sBarCode As String, sEachCode(6) As String, i As Long, iPropagate As Single
Dim rs As Recordset
    Set rs = db.OpenRecordset("Select Barcode.Code From Barcode")
    If rs.RecordCount > 0 Then
        sBarCode = "" & rs!Code
    Else
        sBarCode = "000000"
        getNewBarCode = sBarCode
        Exit Function
    End If
    rs.Close
    
    sEachCode(5) = Mid(sBarCode, 1, 1)
    sEachCode(4) = Mid(sBarCode, 2, 1)
    sEachCode(3) = Mid(sBarCode, 3, 1)
    sEachCode(2) = Mid(sBarCode, 4, 1)
    sEachCode(1) = Mid(sBarCode, 5, 1)
    sEachCode(0) = Mid(sBarCode, 6, 1)
    
    i = 0
    iPropagate = 0
    While i <= 5
        If i = 0 Or iPropagate = 1 Then
            If sEachCode(i) = "Z" Then
                sEachCode(i) = incrementCode(sEachCode(i))
                iPropagate = 1
            Else
                sEachCode(i) = incrementCode(sEachCode(i))
                iPropagate = 0
            End If
        End If
        i = i + 1
    Wend
    sBarCode = sEachCode(5) & sEachCode(4) & sEachCode(3) & sEachCode(2) & sEachCode(1) & sEachCode(0)
    
    getNewBarCode = sBarCode
End Function

Private Function incrementCode(sCode As String) As String
        Select Case sCode
        Case "0"
            sCode = "1"
            GoTo GoOut
        Case "1"
            sCode = "2"
            GoTo GoOut
        Case "2"
            sCode = "3"
            GoTo GoOut
        Case "3"
            sCode = "4"
            GoTo GoOut
        Case "4"
            sCode = "5"
            GoTo GoOut
        Case "5"
            sCode = "6"
            GoTo GoOut
        Case "6"
            sCode = "7"
            GoTo GoOut
        Case "7"
            sCode = "8"
            GoTo GoOut
        Case "8"
            sCode = "9"
            GoTo GoOut
        Case "9"
            sCode = "A"
            GoTo GoOut
        Case "A"
            sCode = "B"
            GoTo GoOut
        Case "B"
            sCode = "C"
            GoTo GoOut
        Case "C"
            sCode = "D"
            GoTo GoOut
        Case "D"
            sCode = "E"
            GoTo GoOut
        Case "E"
            sCode = "F"
            GoTo GoOut
        Case "F"
            sCode = "G"
            GoTo GoOut
        Case "G"
            sCode = "H"
            GoTo GoOut
        Case "H"
            sCode = "I"
            GoTo GoOut
        Case "I"
            sCode = "J"
            GoTo GoOut
        Case "J"
            sCode = "K"
            GoTo GoOut
        Case "K"
            sCode = "L"
            GoTo GoOut
        Case "L"
            sCode = "M"
            GoTo GoOut
        Case "M"
            sCode = "N"
            GoTo GoOut
        Case "N"
            sCode = "O"
            GoTo GoOut
        Case "O"
            sCode = "P"
            GoTo GoOut
        Case "P"
            sCode = "Q"
            GoTo GoOut
        Case "Q"
            sCode = "R"
            GoTo GoOut
        Case "R"
            sCode = "S"
            GoTo GoOut
        Case "S"
            sCode = "T"
            GoTo GoOut
        Case "T"
            sCode = "U"
            GoTo GoOut
        Case "U"
            sCode = "V"
            GoTo GoOut
        Case "V"
            sCode = "W"
            GoTo GoOut
        Case "W"
            sCode = "X"
            GoTo GoOut
        Case "X"
            sCode = "Y"
            GoTo GoOut
        Case "Y"
            sCode = "Z"
            GoTo GoOut
        Case "Z"
            sCode = "0"
            GoTo GoOut
    End Select
GoOut:
    incrementCode = sCode
End Function

Public Function getBarCodeFor( _
    sItemCode As String, _
    sItemType As String, _
    sSupplierCode As String, _
    dUnitPurchaseRate As Double, _
    dUnitMRP As Double, _
    dUnitWholeSaleRate As Double _
                            ) As String

Dim rs As Recordset, sBarCode As String

    Set rs = db.OpenRecordset("Select BarcodeRegister.* From BarcodeRegister Where (BarcodeRegister.ItemCode = '" & sItemCode & "' ) And (BarcodeRegister.ItemType = '" & sItemType & "' ) And (BarcodeRegister.SupplierCode = '" & sSupplierCode & "' ) And (BarcodeRegister.UnitPurchaseRate = " & dUnitPurchaseRate & " ) And (BarcodeRegister.UnitMRP = " & dUnitMRP & " ) And (BarcodeRegister.UnitWholeSaleRate = " & dUnitWholeSaleRate & " )")
    If rs.RecordCount > 0 Then
        sBarCode = rs!BarCode
    Else
        sBarCode = getNewBarCode
        
        rs.AddNew
        rs!BarCode = sBarCode
        rs!ItemCode = sItemCode
        rs!ItemType = sItemType
        rs!SupplierCode = sSupplierCode
        rs!UnitPurchaseRate = dUnitPurchaseRate
        rs!UnitMRP = dUnitMRP
        rs!UnitWholeSaleRate = dUnitWholeSaleRate
        rs.Update
        
        saveBarCode sBarCode
    End If
    getBarCodeFor = sBarCode
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    BARCODE CODING END                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                UNWANTED CODE FROM OLD PROJECTS                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'Private Sub createNewCustomer()
'Dim rs As Recordset, sStatus As String, ssAccountCode As String, ssCustomerCode As String
'
'    ssAccountCode = getNewAccountcode()
'    ssCustomerCode = getNewCustomerCode()
'
'    Set rs = db.OpenRecordset("Select * From CustomerMaster ")
'    rs.AddNew
'    rs!AccountCode = ssAccountCode
'    rs!CustomerCode = ssCustomerCode
'    rs!CustomerName = Trim(CoCustomer.Text)
'    rs!Address1 = Trim(TAddress.Text)
'    rs!Address2 = ""
'    rs!Address3 = ""
'    rs!TinNo = ""
'    rs!Phone = ""
'    rs!Narration = "Auto Created Customer"
'    rs!Status = True
'    rs.Update
'    rs.Close
'
'    'CREATING ACCOUNT FOR THE CUSTOMER IN ACCOUNT MASTER
'    createAccount ssAccountCode
'
'End Sub
'
'Private Sub createAccount(sAccountCode As String)
'Dim rs As Recordset
'    Set rs = db.OpenRecordset("Select AccountMaster.* From AccountMaster")
'
'    rs.AddNew
'    rs!Code = sAccountCode
'    rs!Type = "BAccount"
'    rs!GroupCode = 1
'    rs!AccountName = Trim(CoCustomer.Text)
'    rs!Address1 = Trim(TAddress.Text)
'    rs!Address2 = ""
'    rs!Address3 = ""
'    rs!Phone = ""
'    rs!Narration = ""
'    rs!Status = True
'    rs.Update
'    rs.Close
'End Sub

'
'Private Function getNewCustomerCode() As String
'Dim rs As Recordset, sCCode As String
'
'    Set rs = db.OpenRecordset("Select Max(Val(CustomerCode)) As CCode From CustomerMaster")
'    If rs.RecordCount > 0 Then
'        sCCode = Val("" & rs!CCode) + 1
'    Else
'        sCCode = 1
'    End If
'    rs.Close
'
'    getNewCustomerCode = sCCode
'End Function
'



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                UNWANTED CODE FROM OLD PROJECTS                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
