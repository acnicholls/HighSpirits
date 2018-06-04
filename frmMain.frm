VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "High Spirits Inventory System"
   ClientHeight    =   6930
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuNewProducts 
      Caption         =   "&New Products"
      Begin VB.Menu mnuNewBeer 
         Caption         =   "&Beer"
      End
      Begin VB.Menu mnuNewWine 
         Caption         =   "&Wine"
      End
      Begin VB.Menu mnuNewLiquor 
         Caption         =   "&Liquor"
      End
      Begin VB.Menu mnuNewAccessories 
         Caption         =   "&Accessories"
      End
   End
   Begin VB.Menu mnuModifyProducts 
      Caption         =   "Modify Products"
      Begin VB.Menu mnuModifyBeer 
         Caption         =   "&Beer"
      End
      Begin VB.Menu mnuModifyWine 
         Caption         =   "&Wine"
      End
      Begin VB.Menu mnuModifyLiquor 
         Caption         =   "&Liquor"
      End
      Begin VB.Menu mnuModifyAccessories 
         Caption         =   "&Accessories"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuTransactionReceiving 
         Caption         =   "&Receiving"
      End
      Begin VB.Menu mnuTransactionShipping 
         Caption         =   "&Shipping"
      End
   End
   Begin VB.Menu mnuInvoices 
      Caption         =   "&Invoices"
      Begin VB.Menu mnuInvoiceNew 
         Caption         =   "&New Invoice"
      End
      Begin VB.Menu mnuInvoicePrint 
         Caption         =   "&Print Invoice"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsTransactions 
         Caption         =   "&Transactions"
      End
      Begin VB.Menu mnuReportsInventory 
         Caption         =   "&Inventory"
      End
      Begin VB.Menu mnuReportsUsers 
         Caption         =   "&Users"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "&Users"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    If InStr(gstrCurrentUserRights, "A") <> 0 Then
        mnuUsers.Visible = True
    End If

    If InStr(gstrCurrentUserRights, "B") = 0 Then
        mnuInvoiceNew.Enabled = False
    End If

    If InStr(gstrCurrentUserRights, "I") = 0 Then
        mnuReports.Enabled = False
    End If

    If InStr(gstrCurrentUserRights, "R") = 0 Then
        mnuTransactionReceiving.Enabled = False
    End If

    If InStr(gstrCurrentUserRights, "S") = 0 Then
        Me.mnuTransactionShipping.Enabled = False
    End If

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal
    Set frmAbout = Nothing

End Sub

Private Sub mnuHelpHelp_Click()

    gstrMessage = "This option is not implemented yet!"
    gstrTitle = "Warning message ..."
    gintStyle = vbOKOnly + vbExclamation
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)

End Sub

Private Sub mnuInvoiceNew_Click()

    frmInvoice.Show vbModal
    Set frmInvoice = Nothing

End Sub

Private Sub mnuInvoicePrint_Click()

    Dim strText As String
    Dim strNextInvoiceNumber As String
    Dim strInvoiceToPrint As String
    
    strText = "Please, enter the Invoice # you want to print:"
    strInvoiceToPrint = InputBox(strText)
    
    If StrPtr(strInvoiceToPrint) = 0 Then
        Exit Sub
    End If
    
    If Len(strInvoiceToPrint) <> 6 Then
        gstrMessage = strInvoiceToPrint & " is not an existing invoice reference number."
        gstrTitle = "Invalid request ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        Exit Sub
    End If
    
    If (UCase(Mid(strInvoiceToPrint, 1, 1)) <> "Z") Or (Not IsNumeric(Mid(strInvoiceToPrint, 2, 5))) Then
        gstrMessage = strInvoiceToPrint & " is not an existing invoice reference number."
        gstrTitle = "Invalid request ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        Exit Sub
    End If
    
    strNextInvoiceNumber = GetNextInvoiceID
    If Val(Mid(strInvoiceToPrint, 2, 5)) > (Val(Mid(strNextInvoiceNumber, 2, 5)) - 1) Then
        gstrMessage = strInvoiceToPrint & " is not an existing invoice reference number."
        gstrTitle = "Invalid request ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        Exit Sub
    End If
    
    PrintInvoice (strInvoiceToPrint)
    
    gstrMessage = "You Invoice is now avalaible at the printer."
    gstrTitle = "Printing completed ..."
    gintStyle = vbOKOnly
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)

End Sub

Private Sub mnuModifyAccessories_Click()

    With frmProductSelector
        .optAccessories.Value = True
        .Show vbModal
    End With
    Set frmProductSelector = Nothing

End Sub

Private Sub mnuModifyBeer_Click()

    With frmProductSelector
        .optBeer.Value = True
        .Show vbModal
    End With
    Set frmProductSelector = Nothing

End Sub

Private Sub mnuModifyLiquor_Click()

    With frmProductSelector
        .optLiquor.Value = True
        .Show vbModal
    End With
    Set frmProductSelector = Nothing

End Sub

Private Sub mnuModifyWine_Click()

    With frmProductSelector
        .optWine.Value = True
        .Show vbModal
    End With
    Set frmProductSelector = Nothing

End Sub

Private Sub mnuNewAccessories_Click()

    With frmNewProducts
        .Tag = "Accessories"
        .Caption = "Add New Accessories to Database ..."
        .lblUPCData.Caption = GetNextUPC("Accessories")
        .cboCountry.Enabled = False
        .cboVolume.Enabled = False
        .Show vbModal
    End With
    Set frmNewProducts = Nothing

End Sub

Private Sub mnuNewBeer_Click()

    With frmNewProducts
        .Tag = "Beer"
        .Caption = "Add New Beer to Database ..."
        .lblUPCData.Caption = GetNextUPC("Beer")
        .Show vbModal
    End With
    Set frmNewProducts = Nothing

End Sub

Private Sub mnuNewLiquor_Click()

    With frmNewProducts
        .Tag = "Liquor"
        .Caption = "Add New Liquor to Database ..."
        .lblUPCData.Caption = GetNextUPC("Liquor")
        .Show vbModal
    End With
    Set frmNewProducts = Nothing

End Sub

Private Sub mnuNewWine_Click()

    With frmNewProducts
        .Tag = "Wine"
        .Caption = "Add New Wine to Database ..."
        .lblUPCData.Caption = GetNextUPC("Wine")
        .Show vbModal
    End With
    Set frmNewProducts = Nothing

End Sub

Private Sub mnuReportsInventory_Click()

    frmInventoryReportSelector.Show vbModal
    Set frmInventoryReportSelector = Nothing

End Sub

Private Sub mnuReportsTransactions_Click()

    ReportsHeading "TRANSACTIONS", 2
    TransactionsReport
    ReportCompletedMessage

End Sub

Private Sub mnuReportsUsers_Click()

    ReportsHeading "USERS", 1
    UsersReport
    ReportCompletedMessage

End Sub

Private Sub mnuTransactionReceiving_Click()

    frmReceiving.Show vbModal
    Set frmReceiving = Nothing

End Sub

Private Sub mnuTransactionShipping_Click()

    frmShipping.Show vbModal
    Set frmShipping = Nothing

End Sub

Private Sub mnuUsers_Click()

    frmAdministrator.Show vbModal
    Set frmAdministrator = Nothing

End Sub
