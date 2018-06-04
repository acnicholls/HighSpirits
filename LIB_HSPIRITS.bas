Attribute VB_Name = "LIB_HSPIRITS"
Option Explicit

Public gintCounter As Integer

Public gstrMessage As String
Public gstrTitle As String
Public gintStyle As Integer
Public gintAnswer As Integer

Public gstrCurrentUserID As String
Public gstrCurrentUserLastName As String
Public gstrCurrentUserFirstName As String
Public gstrCurrentUserRights As String

Public Const gstrSuperPassword As String = "40proof"

Public Sub PrintInvoice(ByVal strInvoice As String)

    Dim intFileHandleTransactionDate As Integer
    Dim intFileHandleTransaction As Integer
    Dim intFileHandleProduct As Integer
    Dim strFileName As String
    Dim blnFound As Boolean
    Dim strCategory As String
    Dim strDataToPrint As String * 13
    Dim strTransactionID As String
    Dim strTransactionType As String
    Dim strTransactionDescription As String
    Dim strTransactionDate As String
    Dim strTransactionUPC As String
    Dim sngTransactionBefore As Single
    Dim sngTransactionValue As Single
    Dim sngTransactionAfter As Single
    Dim sngTransactionUnitPrice As Single
    Dim strProductUPC As String
    Dim strProductName As String
    Dim intProductCountry As Integer
    Dim intProductType As Integer
    Dim intProductVolume As Integer
    Dim sngProductPrice As Single
    Dim sngInvoiceTotal As Single
    Dim intNumberOfLines As Integer
    
    'Set Printer options
    With Printer
        .PrintQuality = vbPRPQDraft
        .FontName = "courier new"
        .FontBold = True
        .Orientation = vbPRORLandscape
    End With
    
    'First lecture to extract Invoice Date
    intFileHandleTransactionDate = FreeFile
    blnFound = False
    Open App.Path & "\Data\Transaction.dat" For Input As #intFileHandleTransactionDate
    Do While Not blnFound
        Input #intFileHandleTransactionDate, strTransactionID, strTransactionType, strTransactionDate, strTransactionUPC, _
                                                                    sngTransactionBefore, sngTransactionValue, sngTransactionAfter, sngTransactionUnitPrice
        If strTransactionType = UCase(strInvoice) Then
            blnFound = True
            Exit Do
        End If
    Loop
    Close #intFileHandleTransactionDate
    
    If Not blnFound Then
        gstrMessage = strInvoice & " is not an existing invoice reference number."
        gstrTitle = "Invalid request ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        Exit Sub
    End If
    
    'Print Invoice Header
    Printer.FontSize = 40
    Printer.Print
    Printer.Print " Invoice"
    Printer.Print
    Printer.FontSize = 16
    Printer.Print
    Printer.Print "  High Spirits Beverage Outlets"
    Printer.Print "  Invoice # " & UCase(strInvoice)
    Printer.Print "  Date: " & Format(strTransactionDate, "mm/dd/yyyy")
    Printer.Print
    Printer.Print
    Printer.FontSize = 10
    Printer.Print "    Prod. UPC", "Description", "Volume", "       Units", "   Unit Price", "    Sub Total"
    Printer.Print "    _______________________________________________________________________________________________"
    Printer.Print
    Printer.Print
    Printer.FontSize = 10
    
    'Print Invoice Detail
    intNumberOfLines = 0
    intFileHandleTransaction = FreeFile
    Open App.Path & "\Data\Transaction.dat" For Input As #intFileHandleTransaction
    Do While Not EOF(intFileHandleTransaction)
        Input #intFileHandleTransaction, strTransactionID, strTransactionType, strTransactionDate, strTransactionUPC, _
                                                             sngTransactionBefore, sngTransactionValue, sngTransactionAfter, sngTransactionUnitPrice
        If strTransactionType = UCase(strInvoice) Then
            blnFound = False
            intFileHandleProduct = FreeFile
            intNumberOfLines = intNumberOfLines + 1
            Select Case Mid(strTransactionUPC, 1, 1)
                Case "B"
                    strFileName = App.Path & "\Data\Beer.dat"
                    strCategory = "Beer"
                Case "W"
                    strFileName = App.Path & "\Data\Wine.dat"
                    strCategory = "Wine"
                Case "L"
                    strFileName = App.Path & "\Data\Liquor.dat"
                    strCategory = "Liquor"
                Case "A"
                    strFileName = App.Path & "\Data\Accessories.dat"
                    strCategory = "Accessories"
                Case Else
            End Select

            Open strFileName For Input As #intFileHandleProduct
            Do While Not blnFound
                Input #intFileHandleProduct, strProductUPC, strProductName, intProductCountry, intProductType, _
                                                               intProductVolume, sngProductPrice
                If strTransactionUPC = strProductUPC Then
                    blnFound = True
                    sngInvoiceTotal = sngInvoiceTotal + sngTransactionValue * sngProductPrice
                    Printer.Print Space(4) & strProductUPC,
                    Printer.Print Mid(strProductName, 1, 13),
                    Printer.Print ReturnVolume(intProductVolume),
                    RSet strDataToPrint = sngTransactionValue
                    Printer.Print strDataToPrint,
                    RSet strDataToPrint = Format(sngProductPrice, "currency")
                    Printer.Print strDataToPrint,
                    RSet strDataToPrint = Format(sngTransactionValue * sngProductPrice, "currency")
                    Printer.Print strDataToPrint
                End If
            Loop
            Close #intFileHandleProduct
            If intNumberOfLines = 10 Then
                Exit Do
            End If
        End If
    Loop
    Close
    Printer.FontSize = 16
    Printer.Print
    Printer.Print
    Printer.Print "  Total: " & Format(sngInvoiceTotal, "currency")
    Printer.EndDoc

End Sub

Public Sub ReportCompletedMessage()

    gstrMessage = "You report is now avalaible at the printer."
    gstrTitle = "Report completed ..."
    gintStyle = vbOKOnly
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)

End Sub

Public Sub ReportsHeading(ByVal strTitle As String, ByVal intOrientation As Integer)

    Dim strLongTitle As String
    Dim intMyTab As Integer
    
    strLongTitle = "***** " & strTitle & " REPORT" & " *****"
    
    With Printer
        .PrintQuality = vbPRPQDraft
        .FontName = "courier new"
        .FontBold = True
        .FontSize = 14
        Select Case intOrientation
            Case 1
                .Orientation = vbPRORPortrait
                intMyTab = 69
            Case 2
                .Orientation = vbPRORLandscape
                intMyTab = 95
            Case Else
                .Orientation = vbPRORPortrait
                intMyTab = 69
        End Select
    End With
    
    If strTitle = "USERS" Then
        Printer.Print
        Printer.Print
        Printer.FontSize = 40
        Printer.Print " CONFIDENTIAL"
        Printer.FontSize = 14
    End If
    Printer.Print
    Printer.Print
    Printer.Print Space((intMyTab - Len("High Spirit Beverage Outlets")) / 2); "High Spirit Beverage Outlets"
    Printer.Print Space((intMyTab - Len(strLongTitle)) / 2); strLongTitle
    Printer.Print
    Printer.FontSize = 12
    Printer.Print Space(3) & "DATE:"; Tab(intMyTab + 6); "TIME:"
    Printer.Print Space(3) & Format(Date, "mm/dd/yyyy"); Tab(intMyTab + 3); Format(Time, "hh:mm:ss")
    Printer.Print
    Printer.Print

End Sub

Public Sub TransactionsReport()

    Dim intFileHandle As Integer
    Dim strDataToPrint As String * 11
    Dim strLongDataToPrint As String * 13
    Dim strTransactionID As String
    Dim strTransactionType As String
    Dim strTransactionDescription As String
    Dim strTransactionDate As String
    Dim strTransactionUPC As String
    Dim sngTransactionBefore As Single
    Dim sngTransactionValue As Single
    Dim sngTransactionAfter As Single
    Dim sngTransactionUnitPrice As Single
    
    Printer.FontSize = 10
    Printer.Print Space(2) & String(11, "_"),
    For gintCounter = 1 To 7
        Printer.Print String(11, "_"),
    Next gintCounter
    Printer.Print String(13, "_")
    Printer.Print
    
    RSet strDataToPrint = "Trans. ID"
    Printer.Print Space(2) & strDataToPrint,
    RSet strDataToPrint = "Type"
    Printer.Print strDataToPrint,
    RSet strDataToPrint = "Date"
    Printer.Print strDataToPrint,
    RSet strDataToPrint = "Product UPC"
    Printer.Print strDataToPrint,
    RSet strDataToPrint = "Before"
    Printer.Print strDataToPrint,
    RSet strDataToPrint = "Transaction"
    Printer.Print strDataToPrint,
    RSet strDataToPrint = "After"
    Printer.Print strDataToPrint,
    RSet strDataToPrint = "Unit Price"
    Printer.Print strDataToPrint,
    RSet strLongDataToPrint = "Trans Total"
    Printer.Print strLongDataToPrint
    Printer.Print Space(2) & String(11, "_"),
    
    For gintCounter = 1 To 7
        Printer.Print String(11, "_"),
    Next gintCounter
    Printer.Print String(13, "_")
    Printer.Print
    
    intFileHandle = FreeFile
    Open App.Path & "\Data\Transaction.dat" For Input As #intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strTransactionID, strTransactionType, strTransactionDate, strTransactionUPC, _
                                          sngTransactionBefore, sngTransactionValue, sngTransactionAfter, sngTransactionUnitPrice
        
        Select Case strTransactionType
            Case "R"
                strTransactionDescription = "Receiving"
            Case "S"
                strTransactionDescription = "Shipping"
                sngTransactionValue = sngTransactionValue * -1
            Case Else
                strTransactionDescription = strTransactionType
                sngTransactionValue = sngTransactionValue * -1
        End Select
        
        RSet strDataToPrint = Space(2) & strTransactionID
        Printer.Print strDataToPrint,
        RSet strDataToPrint = strTransactionDescription
        Printer.Print strDataToPrint,
        RSet strDataToPrint = Format(strTransactionDate, "mm/dd/yyyy")
        Printer.Print strDataToPrint,
        RSet strDataToPrint = strTransactionUPC
        Printer.Print strDataToPrint,
        RSet strDataToPrint = sngTransactionBefore
        Printer.Print strDataToPrint,
        RSet strDataToPrint = sngTransactionValue
        Printer.Print strDataToPrint,
        RSet strDataToPrint = sngTransactionAfter
        Printer.Print strDataToPrint,
        RSet strDataToPrint = Format(sngTransactionUnitPrice, "currency")
        Printer.Print strDataToPrint,
        RSet strLongDataToPrint = Format(sngTransactionUnitPrice * sngTransactionValue, "currency")
        Printer.Print strLongDataToPrint
    Loop
    
    Close
    
    Printer.FontSize = 12
    Printer.Print
    Printer.Print
    Printer.Print "     End of Report"
    Printer.EndDoc
    
End Sub

Public Sub UsersReport()

    Dim intFileHandle As Integer
    Dim strUserID As String
    Dim strUserLastName As String
    Dim strUserFirstName As String
    Dim strUserRights As String
    Dim strUserPassword As String
    
    Printer.FontSize = 10
    Printer.Print ,
    For gintCounter = 1 To 4
        Printer.Print String(13, "_"),
    Next gintCounter
    Printer.Print String(13, "_")
    Printer.Print
    
    Printer.Print , "User ID", "Last Name", "First Name", "Rights", "Password"
    
    Printer.Print ,
    For gintCounter = 1 To 4
        Printer.Print String(13, "_"),
    Next gintCounter
    Printer.Print String(13, "_")
    Printer.Print
    
    intFileHandle = FreeFile
    Open App.Path & "\Data\Users.dat" For Input As #intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strUserID, strUserLastName, strUserFirstName, strUserRights, strUserPassword
        Printer.Print ,
        Printer.Print strUserID, strUserLastName, strUserFirstName, strUserRights, strUserPassword
    Loop
    
    Close
    
    Printer.FontSize = 12
    Printer.Print
    Printer.Print
    Printer.Print ,
    Printer.Print "Rights Codes:"
    Printer.Print
    Printer.Print ,
    Printer.Print "(A)dministrator"
    Printer.Print ,
    Printer.Print "(B)illing"
    Printer.Print ,
    Printer.Print "(I)nformation Reports"
    Printer.Print ,
    Printer.Print "(R)eceiving"
    Printer.Print ,
    Printer.Print "(S)hipping"
    Printer.Print
    Printer.Print ,
    Printer.Print "End of Report"
    Printer.EndDoc
    
End Sub

Public Sub InventoryReport(ByVal strTypeOfReport As String)

    Dim intFileHandleCategory As Integer
    Dim intFileHandleStock As Integer
    Dim blnFound As Boolean
    Dim strDataToPrint As String * 13
    Dim strCategoryUPC As String
    Dim strStockUPC As String
    Dim strDescription As String
    Dim sngUnits As Single
    Dim sngPrice As Single
    Dim sngUPCTotal As Single
    Dim sngInventoryTotal As Single
    Dim intCountry As Integer
    Dim intType As Integer
    Dim intVolume As Integer
    
    sngUPCTotal = 0
    sngInventoryTotal = 0
    
    Printer.FontSize = 10
    Printer.Print ,
    For gintCounter = 1 To 4
        Printer.Print String(13, "_"),
    Next gintCounter
    Printer.Print String(13, "_")
    Printer.Print
    
    Printer.Print , " Product UPC", " Description", "    Units", "    Price", " Stock Value"
    
    Printer.Print ,
    For gintCounter = 1 To 4
        Printer.Print String(13, "_"),
    Next gintCounter
    Printer.Print String(13, "_")
    Printer.Print
    
    If strTypeOfReport = "Beer" Or strTypeOfReport = "ALL" Then
        intFileHandleCategory = FreeFile
        Open App.Path & "\Data\Beer.dat" For Input As #intFileHandleCategory
        Do While Not EOF(intFileHandleCategory)
            Input #intFileHandleCategory, strCategoryUPC, strDescription, intCountry, intType, intVolume, sngPrice
            blnFound = False
            intFileHandleStock = FreeFile
            Open App.Path & "\Data\Stock.dat" For Input As #intFileHandleStock
            Do While Not blnFound
                Input #intFileHandleStock, strStockUPC, sngUnits
                If strStockUPC = strCategoryUPC Then
                    blnFound = True
                    Exit Do
                End If
            Loop
            Close #intFileHandleStock
            
            sngUPCTotal = sngUnits * sngPrice
            sngInventoryTotal = sngInventoryTotal + sngUPCTotal
            
            Printer.Print ,
            Printer.Print strCategoryUPC,
            Printer.Print Mid(strDescription, 1, 13),
            RSet strDataToPrint = sngUnits
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngPrice, "currency")
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngUPCTotal, "currency")
            Printer.Print strDataToPrint
        Loop
        Close
    End If
    
    If strTypeOfReport = "Wine" Or strTypeOfReport = "ALL" Then
        intFileHandleCategory = FreeFile
        Open App.Path & "\Data\Wine.dat" For Input As #intFileHandleCategory
        Do While Not EOF(intFileHandleCategory)
            Input #intFileHandleCategory, strCategoryUPC, strDescription, intCountry, intType, intVolume, sngPrice
            blnFound = False
            intFileHandleStock = FreeFile
            Open App.Path & "\Data\Stock.dat" For Input As #intFileHandleStock
            Do While Not blnFound
                Input #intFileHandleStock, strStockUPC, sngUnits
                If strStockUPC = strCategoryUPC Then
                    blnFound = True
                    Exit Do
                End If
            Loop
            Close #intFileHandleStock
            
            sngUPCTotal = sngUnits * sngPrice
            sngInventoryTotal = sngInventoryTotal + sngUPCTotal
            
            Printer.Print ,
            Printer.Print strCategoryUPC,
            Printer.Print Mid(strDescription, 1, 13),
            RSet strDataToPrint = sngUnits
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngPrice, "currency")
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngUPCTotal, "currency")
            Printer.Print strDataToPrint
        Loop
        Close
    End If
    
    If strTypeOfReport = "Liquor" Or strTypeOfReport = "ALL" Then
        intFileHandleCategory = FreeFile
        Open App.Path & "\Data\Liquor.dat" For Input As #intFileHandleCategory
        Do While Not EOF(intFileHandleCategory)
            Input #intFileHandleCategory, strCategoryUPC, strDescription, intCountry, intType, intVolume, sngPrice
            blnFound = False
            intFileHandleStock = FreeFile
            Open App.Path & "\Data\Stock.dat" For Input As #intFileHandleStock
            Do While Not blnFound
                Input #intFileHandleStock, strStockUPC, sngUnits
                If strStockUPC = strCategoryUPC Then
                    blnFound = True
                    Exit Do
                End If
            Loop
            Close #intFileHandleStock
            
            sngUPCTotal = sngUnits * sngPrice
            sngInventoryTotal = sngInventoryTotal + sngUPCTotal
            
            Printer.Print ,
            Printer.Print strCategoryUPC,
            Printer.Print Mid(strDescription, 1, 13),
            RSet strDataToPrint = sngUnits
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngPrice, "currency")
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngUPCTotal, "currency")
            Printer.Print strDataToPrint
        Loop
        Close
    End If
    
    If strTypeOfReport = "Accessories" Or strTypeOfReport = "ALL" Then
        intFileHandleCategory = FreeFile
        Open App.Path & "\Data\Accessories.dat" For Input As #intFileHandleCategory
        Do While Not EOF(intFileHandleCategory)
            Input #intFileHandleCategory, strCategoryUPC, strDescription, intCountry, intType, intVolume, sngPrice
            blnFound = False
            intFileHandleStock = FreeFile
            Open App.Path & "\Data\Stock.dat" For Input As #intFileHandleStock
            Do While Not blnFound
                Input #intFileHandleStock, strStockUPC, sngUnits
                If strStockUPC = strCategoryUPC Then
                    blnFound = True
                    Exit Do
                End If
            Loop
            Close #intFileHandleStock
            
            sngUPCTotal = sngUnits * sngPrice
            sngInventoryTotal = sngInventoryTotal + sngUPCTotal
            
            Printer.Print ,
            Printer.Print strCategoryUPC,
            Printer.Print Mid(strDescription, 1, 13),
            RSet strDataToPrint = sngUnits
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngPrice, "currency")
            Printer.Print strDataToPrint,
            RSet strDataToPrint = Format(sngUPCTotal, "currency")
            Printer.Print strDataToPrint
        Loop
        Close
    End If
    
    Printer.FontSize = 12
    Printer.Print
    Printer.Print
    Printer.Print ,
    Printer.Print "Total Inventory Value:  " & Format(sngInventoryTotal, "currency")
    Printer.Print ,
    Printer.Print "End of Report"
    Printer.EndDoc
    
End Sub

Public Function ReturnVolume(ByVal intVolume As Integer) As String

    Select Case intVolume
        Case 0
            ReturnVolume = "355 ml"
        Case 1
            ReturnVolume = "375 ml"
        Case 2
            ReturnVolume = "500 ml"
        Case 3
            ReturnVolume = "750 ml"
        Case 4
            ReturnVolume = "1.00 l"
        Case 5
            ReturnVolume = "1.18 l"
        Case 6
            ReturnVolume = "1.50 l"
        Case 7
            ReturnVolume = "3.00 l"
    End Select

End Function

Public Function ReturnCountry(ByVal intCountry As Integer) As String

    Select Case intCountry
        Case 0
            ReturnCountry = "Australia"
        Case 1
            ReturnCountry = "Belgium"
        Case 2
            ReturnCountry = "Canada"
        Case 3
            ReturnCountry = "France"
        Case 4
            ReturnCountry = "Italy"
        Case 5
            ReturnCountry = "Spain"
        Case 6
            ReturnCountry = "USA"
    End Select

End Function

Public Function ReturnType(ByVal strCategory As String, ByVal intType As Integer) As String

    Select Case strCategory
        Case "Beer"
            Select Case intType
                Case 0
                    ReturnType = "Alcohol free"
                Case 1
                    ReturnType = "Ale"
                Case 2
                    ReturnType = "Dark"
                Case 3
                    ReturnType = "Lager"
                Case 4
                    ReturnType = "Micro brewed"
            End Select
        Case "Wine"
            Select Case intType
                Case 0
                    ReturnType = "Red"
                Case 1
                    ReturnType = "Rose"
                Case 2
                    ReturnType = "Sparkling"
                Case 3
                    ReturnType = "White"
            End Select
        Case "Liquor"
            Select Case intType
                Case 0
                    ReturnType = "Aperitif Wine"
                Case 1
                    ReturnType = "Brandy"
                Case 2
                    ReturnType = "Cognac"
                Case 3
                    ReturnType = "Gin"
                Case 4
                    ReturnType = "Grappa"
                Case 5
                    ReturnType = "Ouzo"
                Case 6
                    ReturnType = "Rhum"
                Case 7
                    ReturnType = "Scotch"
                Case 8
                    ReturnType = "Vodka"
                Case 9
                    ReturnType = "Whiskey"
            End Select
        Case "Accessories"
            Select Case intType
                Case 0
                    ReturnType = "Books"
                Case 1
                    ReturnType = "Bottle openers"
                Case 2
                    ReturnType = "Gift Packs"
                Case 3
                    ReturnType = "Glasses"
                Case 4
                    ReturnType = "Wine Racks"
            End Select
        End Select

End Function

Public Function GetNextUPC(ByVal strCategory As String) As String

    Dim intFileHandle As Integer
    Dim strNextBeerUPC As String
    Dim strNextWineUPC As String
    Dim strNextLiquorUPC As String
    Dim strNextAccessoriesUPC As String
    
     intFileHandle = FreeFile
     Open App.Path & "\Data\UPCGenerator.dat" For Input As #intFileHandle
     Input #intFileHandle, strNextBeerUPC, strNextWineUPC, strNextLiquorUPC, strNextAccessoriesUPC
     Close #intFileHandle
     
     Select Case strCategory
        Case "Beer"
            GetNextUPC = strNextBeerUPC
        Case "Wine"
            GetNextUPC = strNextWineUPC
        Case "Liquor"
            GetNextUPC = strNextLiquorUPC
        Case "Accessories"
            GetNextUPC = strNextAccessoriesUPC
    End Select
    
End Function

Public Function GetNextTransactionID() As String

    Dim intFileHandle As Integer
    Dim strNextTransactionID As String
    
     intFileHandle = FreeFile
     Open App.Path & "\Data\TransactionGenerator.dat" For Input As #intFileHandle
     Input #intFileHandle, strNextTransactionID
     Close #intFileHandle
     
     GetNextTransactionID = strNextTransactionID

End Function

Public Function GetNextInvoiceID() As String
''''''''''''''''''''''''''''This sub get the next invoice ID for use
''''''''''''''''''''''''''''tblGeneratorforID.fldNextInvoiceID


    Dim intFileHandle As Integer
    Dim strNextInvoiceID As String
    
     intFileHandle = FreeFile
     Open App.Path & "\Data\InvoiceGenerator.dat" For Input As #intFileHandle
''''''''''''''''''''''''''''select fldnextInvoiceID FROM tblGeneratorforID
     Input #intFileHandle, strNextInvoiceID
     Close #intFileHandle
     
     GetNextInvoiceID = strNextInvoiceID

End Function

Public Function GetNextUserID() As String
'''''''''''''''''''''''''''''This sub get teh next user ID for use
'''''''''''''''''''''''''''''tblGeneratorforID.fldNextUserID

    Dim intFileHandle As Integer
    Dim strNextUserID As String
    
     intFileHandle = FreeFile
     Open App.Path & "\Data\IDGenerator.dat" For Input As #intFileHandle
''''''''''''''''''''''''''''''''''''select fldnextuserID from tblgeneratorforID
     Input #intFileHandle, strNextUserID
     Close #intFileHandle
     
     GetNextUserID = strNextUserID

End Function

Public Function GenerateNextUserID(ByVal strCurrentID As String) As Boolean
''''''''''''''''''''''''''''''''This sub generates a new UserID and stores it in a database field
''''''''''''''''''''''tblGeneratorforID.fldNextUserID


    Dim intFileHandle As Integer
        
     intFileHandle = FreeFile
     Open App.Path & "\Data\IDGenerator.dat" For Output As #intFileHandle
'''''''''''''''''''update database field
     Write #intFileHandle, "HSU" & Format(Val(Mid(strCurrentID, 4, 3)) + 1, "000")
     Close #intFileHandle
     
     GenerateNextUserID = True

End Function

Public Function GenerateNextTransactionID(ByVal strCurrentTransactionID As String) As Boolean
'''''''''''''''''''''''''This sub generates a new transaction code and stores it ina a database field
'''''''''''''''''''''tblGeneratorforID.fldNextTransactionID
    Dim intFileHandle As Integer
        
     intFileHandle = FreeFile
     Open App.Path & "\Data\TransactionGenerator.dat" For Output As #intFileHandle
    Write #intFileHandle, "T" & Format(Val(Mid(strCurrentTransactionID, 2, 5)) + 1, "00000")
     Close #intFileHandle
'
     GenerateNextTransactionID = True

End Function

Public Function GenerateNextInvoiceID(ByVal strCurrentInvoiceID As String) As Boolean
''''''''''''''''''''''''''''This sub generates a new invoice code and stores it in a datbase field
''''''''''''''''''''''''''''tblGeneratorforID.fldNextInvoice
    Dim intFileHandle As Integer
'

'''''''''''''''''''''
     intFileHandle = FreeFile
     Open App.Path & "\Data\InvoiceGenerator.dat" For Output As #intFileHandle

'''''''''''''''''''''''update the field
    Write #intFileHandle, "Z" & Format(Val(Mid(strCurrentInvoiceID, 2, 5)) + 1, "00000")
     Close #intFileHandle
     
     GenerateNextInvoiceID = True

End Function

Public Function GenerateNextUPC(ByVal strCurrentUPC As String) As Boolean
'''''''''''''''''''''''''This sub generates a new UPC code every time a new procudt is entered
'''''''''''''''''''''''''  The newly generated upc code is stored in a field
''''''''''''''''''''''''''''tblGeneratorforID
'''''''''''''''''''''''''''''''''''''''''''''''''fldNextAccessoriesID
'''''''''''''''''''''''''''''''''''''''''''''''''fldNextBeerID
'''''''''''''''''''''''''''''''''''''''''''''''''fldNEXTLiqourID
'''''''''''''''''''''''''''''''''''''''''''''''''fldNextWineID


'    Dim intFileHandle As Integer
    Dim strNextBeerUPC As String
    Dim strNextWineUPC As String
    Dim strNextLiquorUPC As String
    Dim strNextAccessoriesUPC As String
    Dim strNewUPC As String
        
    intFileHandle = FreeFile
     Open App.Path & "\Data\UPCGenerator.dat" For Input As #intFileHandle
     Input #intFileHandle, strNextBeerUPC, strNextWineUPC, strNextLiquorUPC, strNextAccessoriesUPC
     Close #intFileHandle
     intFileHandle = FreeFile
     Open App.Path & "\Data\UPCGenerator.dat" For Output As #intFileHandle
     Select Case strCurrentUPC
        Case "Beer"
            strNewUPC = "B" & Format(Val(Mid(strNextBeerUPC, 4, 3)) + 1, "0000")
            
 '''''''''''        update the field witha  new upc code
            Write #intFileHandle, strNewUPC; strNextWineUPC; strNextLiquorUPC; strNextAccessoriesUPC
        Case "Wine"
        ''''''''''''''''''''''update this fiels with a new upc code
            strNewUPC = "W" & Format(Val(Mid(strNextWineUPC, 4, 3)) + 1, "0000")
            Write #intFileHandle, strNextBeerUPC; strNewUPC; strNextLiquorUPC; strNextAccessoriesUPC
        Case "Liquor"
        ''''''''''''''''''''''''update this field with a new upc code
            strNewUPC = "L" & Format(Val(Mid(strNextLiquorUPC, 4, 3)) + 1, "0000")
            Write #intFileHandle, strNextBeerUPC; strNextWineUPC; strNewUPC; strNextAccessoriesUPC
        Case "Accessories"
        '''''''''''''''''''''update this field witha new upc
            strNewUPC = "A" & Format(Val(Mid(strNextAccessoriesUPC, 4, 3)) + 1, "0000")
            Write #intFileHandle, strNextBeerUPC; strNextWineUPC; strNextLiquorUPC; strNewUPC
    End Select
     Close #intFileHandle
     ''''''''''''''''''''''''set strCurrentUPC to true
     strCurrentUPC = True

End Function

