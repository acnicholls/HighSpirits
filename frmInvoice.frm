VERSION 5.00
Begin VB.Form frmInvoice 
   Caption         =   "High spirits Billing System "
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Invoice"
      Height          =   615
      Left            =   3960
      Picture         =   "frmInvoice.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraRemoveItem 
      Caption         =   "Remove Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   4560
      TabIndex        =   12
      Top             =   3840
      Width           =   2535
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtLine 
         Height          =   285
         Left            =   480
         MaxLength       =   2
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblRemoveLine 
         Caption         =   "Line:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraAddItem 
      Caption         =   "Add Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   480
      TabIndex        =   8
      Top             =   3840
      Width           =   3735
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtUnits 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtUPC 
         Height          =   285
         Left            =   360
         MaxLength       =   5
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblAddUnits 
         Caption         =   "Units:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblAddUPC 
         Caption         =   "Product UPC:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ListBox lstLine 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2010
      ItemData        =   "frmInvoice.frx":0102
      Left            =   240
      List            =   "frmInvoice.frx":0104
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ListBox lstSubTotal 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2010
      Left            =   7320
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ListBox lstUnitPrice 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2010
      Left            =   6120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ListBox lstUnits 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2010
      Left            =   5520
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.ListBox lstDescription 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   2010
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.ListBox lstProduct 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2010
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblTaxes 
      Alignment       =   1  'Right Justify
      Caption         =   "(Taxes included)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   5880
      TabIndex        =   23
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblInvoiceNumber 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblDate 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblSubTotal 
      Caption         =   "Sub total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7320
      TabIndex        =   20
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblUnitPrice 
      Caption         =   "Unit price:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblUnits 
      Caption         =   "Units:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblProduct 
      Caption         =   "Product:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   840
      TabIndex        =   16
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblLine 
      Caption         =   "Line:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintFileHandle As Integer
Private mstrFileName As String
Private mstrUPC As String
Private mstrName As String
Private mintCountry As Integer
Private mintType As Integer
Private mintVolume As Integer
Private msngPrice As Single
Private msngQuantity As Single

'This function is used to calaulate the total of the invoice shown in lblTotal label
Private Sub CalculateTotal()

    Dim intCounter As Integer
    Dim sngTotal As Single
    
    sngTotal = 0
    For intCounter = 0 To lstSubTotal.ListCount - 1
        sngTotal = sngTotal + lstSubTotal.List(intCounter)
    Next intCounter
    
    lblTotal.Caption = "TOTAL: " & Format(sngTotal, "currency")
    lblTaxes.Visible = True

End Sub

Private Sub cmdAdd_Click()

    Dim intUnits As Integer
    Dim blnFound As Boolean
    Dim strCategory As String
    Dim objControl As Control
    
    'Exit the Sub if invalid number of units
    intUnits = Val(txtUnits.Text)
    If intUnits < 1 Then
        gstrMessage = "You must specify a valid Number of Units."
        gstrTitle = "Invalid Number of Units ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtUnits.SetFocus
        txtUnits.SelStart = 0
        txtUnits.SelLength = Len(txtUnits.Text)
        Exit Sub
    End If
    
    'Exit the Sub if txtUPC does not correspond to a valid UPC number
    If Len(txtUPC.Text) < 5 Then
        gstrMessage = "You must specify a valid Product Number."
        gstrTitle = "Invalid Product Number ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtUPC.SetFocus
        txtUPC.SelStart = 0
        txtUPC.SelLength = Len(txtUPC.Text)
        Exit Sub
    End If
    
    'The right Category file name is
    Select Case UCase(Mid(txtUPC.Text, 1, 1))
        Case "B"
            strCategory = "Beer"
            mstrFileName = App.Path & "\Data\Beer.dat"
        Case "W"
            strCategory = "Wine"
            mstrFileName = App.Path & "\Data\Wine.dat"
        Case "L"
            strCategory = "Liquor"
            mstrFileName = App.Path & "\Data\Liquor.dat"
        Case "A"
            strCategory = "Accessories"
            mstrFileName = App.Path & "\Data\Accessories.dat"
        Case Else
            gstrMessage = "You must specify a valid Product Number."
            gstrTitle = "Invalid Product Number ..."
            gintStyle = vbOKOnly + vbExclamation
            gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
            txtUPC.SetFocus
            txtUPC.SelStart = 0
            txtUPC.SelLength = Len(txtUPC.Text)
            Exit Sub
    End Select
    
    mintFileHandle = FreeFile
    Open mstrFileName For Input As #mintFileHandle
    Do While (Not blnFound) And (Not EOF(mintFileHandle))
        Input #mintFileHandle, mstrUPC, mstrName, mintCountry, mintType, mintVolume, msngPrice
        If mstrUPC = UCase(txtUPC.Text) Then
            blnFound = True
        End If
    Loop
    Close
    
    If Not blnFound Then
        gstrMessage = "You must specify a valid Product Number."
        gstrTitle = "Invalid Product Number ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtUPC.SetFocus
        txtUPC.SelStart = 0
        txtUPC.SelLength = Len(txtUPC.Text)
        Exit Sub
    End If
    
    lstLine.AddItem lstLine.ListCount + 1
    lstProduct.AddItem mstrUPC
    If UCase(Mid(txtUPC.Text, 1, 1)) = "A" Then
        lstDescription.AddItem ReturnType(strCategory, mintType) & ", " & mstrName
    Else
        lstDescription.AddItem ReturnType(strCategory, mintType) & ", " & mstrName & ", " & ReturnVolume(mintVolume)
    End If
    lstUnits.AddItem intUnits
    lstUnitPrice.AddItem Format(msngPrice, "0.00")
    lstSubTotal.AddItem Format(msngPrice * intUnits, "0.00")
    
    CalculateTotal
    txtUnits.Text = ""
    txtUPC.Text = ""
    txtUPC.SetFocus
    If lstLine.ListCount > 9 Then
        cmdAdd.Enabled = False
    End If
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdPrint_Click()

    cmdPrint.Visible = False
    Height = 4095
    PrintForm
    Height = 5550
    cmdPrint.Visible = True

End Sub

Private Sub cmdRemove_Click()

    Dim intIndex As Integer
    Dim objControl As Control
    
    intIndex = Val(txtLine.Text)
    If intIndex < 1 Or intIndex > lstLine.ListCount Then
        gstrMessage = "You must specify a valid Line Number."
        gstrTitle = "Invalid Line ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtLine.SetFocus
        txtLine.SelStart = 0
        txtLine.SelLength = Len(txtLine.Text)
        Exit Sub
    End If
    
    gstrMessage = "Information on Line #" & intIndex
    gstrMessage = gstrMessage & " will be permanently deleted.  Do you wish to proceed?"
    gstrTitle = "Modify Invoice ..."
    gintStyle = vbYesNo + vbExclamation
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
    If gintAnswer = vbNo Then
        Exit Sub
    End If
    
    intIndex = Val(txtLine.Text)
    For Each objControl In Me
        If TypeOf objControl Is ListBox Then
            objControl.RemoveItem (intIndex - 1)
        End If
    Next objControl
    
    For gintCounter = 0 To lstLine.ListCount - 1
        lstLine.List(gintCounter) = gintCounter + 1
    Next gintCounter
    
    CalculateTotal
    txtLine.Text = ""
    txtUPC.SetFocus
    If lstLine.ListCount < 10 Then
        cmdAdd.Enabled = True
    End If

End Sub

Private Sub cmdSave_Click()

    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    Dim strCurrentTransactionID As String
    Dim strUPCToUpdate As String
    Dim sngQuantityToUpdate As Single
    Dim objControl As Control
    Dim blnRetValue As Boolean
    
    If lstLine.ListCount = 0 Then
        gstrMessage = "There is no information to save!"
        gstrTitle = "Empty invoice ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        Exit Sub
    End If
    
    'update the inventory
    For gintCounter = lstLine.ListCount - 1 To 0 Step -1
        strUPCToUpdate = lstProduct.List(gintCounter)
        intFileHandleInput = FreeFile
        Open App.Path & "\Data\Stock.dat" For Input As #intFileHandleInput
        intFileHandleOutput = FreeFile
        Open App.Path & "\Data\TempStock.dat" For Output As #intFileHandleOutput
        Do While Not EOF(intFileHandleInput)
            Input #intFileHandleInput, mstrUPC, msngQuantity
            If mstrUPC <> strUPCToUpdate Then
                Write #intFileHandleOutput, mstrUPC; msngQuantity
            Else
                sngQuantityToUpdate = msngQuantity
            End If
        Loop
        If sngQuantityToUpdate >= Val(lstUnits.List(gintCounter)) Then
            Write #intFileHandleOutput, strUPCToUpdate; sngQuantityToUpdate - Val(lstUnits.List(gintCounter))
            'write the transaction to file
            strCurrentTransactionID = GetNextTransactionID()
            mintFileHandle = FreeFile
            Open App.Path & "\Data\Transaction.dat" For Append As #mintFileHandle
            Write #mintFileHandle, strCurrentTransactionID; Right(lblInvoiceNumber.Caption, 6); Format(Date, "mm/dd/yyyy"); _
                                                  lstProduct.List(gintCounter); sngQuantityToUpdate; Val(lstUnits.List(gintCounter)); _
                                                  sngQuantityToUpdate - Val(lstUnits.List(gintCounter)); Val(lstUnitPrice.List(gintCounter) * 100) / 100
        Else
            Write #intFileHandleOutput, strUPCToUpdate; sngQuantityToUpdate
            gstrMessage = "Line " & gintCounter + 1 & " of the Invoice was cancelled as there is not enought units in stock."
            gstrTitle = "Not enough units avalaible ..."
            gintStyle = vbOKOnly + vbExclamation
            gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
            For Each objControl In Me
                If TypeOf objControl Is ListBox Then
                    objControl.RemoveItem (gintCounter)
                End If
            Next objControl
        End If
        Close
        GenerateNextTransactionID (strCurrentTransactionID)
        Kill App.Path & "\Data\Stock.dat"
        Name App.Path & "\Data\TempStock.dat" As App.Path & "\Data\Stock.dat"
    Next gintCounter
    
    cmdSave.Enabled = False
    blnRetValue = GenerateNextInvoiceID(Right(lblInvoiceNumber.Caption, 6))
    cmdPrint.Visible = True

End Sub

Private Sub Form_Load()

    lblDate.Caption = UCase(Format(Date, "dddd, mmmm d yyyy"))
    lblInvoiceNumber = "Invoice # " & GetNextInvoiceID
    
End Sub

Private Sub lstLine_Click()

    txtLine.Text = lstLine.List(lstLine.ListIndex)
    cmdRemove.SetFocus

End Sub

Private Sub txtLine_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 8
        Case 13
        Case 48 To 57
        Case Else
            KeyAscii = 0
            Beep
            MsgBox "Please, enter numerical data only.", vbOKOnly + vbExclamation, "Invalid Data"
            txtLine.SetFocus
            txtLine.SelStart = 0
            txtLine.SelLength = Len(txtLine.Text)
    End Select
    
End Sub

Private Sub txtUnits_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 8
        Case 13
        Case 48 To 57
        Case Else
            KeyAscii = 0
            Beep
            MsgBox "Please, enter numerical data only.", vbOKOnly + vbExclamation, "Invalid Data"
            txtUnits.SetFocus
            txtUnits.SelStart = 0
            txtUnits.SelLength = Len(txtUnits.Text)
    End Select

End Sub
