VERSION 5.00
Begin VB.Form frmReceiving 
   Caption         =   "Hight Spirits Receiving"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInventory 
      Caption         =   "Inventory:"
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
      Height          =   1575
      Left            =   2400
      TabIndex        =   12
      Top             =   1680
      Width           =   3735
      Begin VB.TextBox txtReceived 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   2520
         X2              =   3360
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblUpdatedData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2640
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblActualData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblUpdated 
         Caption         =   "Updated:"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblReceived 
         Caption         =   "Received:"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblActual 
         Caption         =   "Actual:"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search By:"
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
      Height          =   1335
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton optInvisible 
         BackColor       =   &H000000FF&
         Caption         =   "THIS RADIO BUTTON IS NOT VISIBLE AT RUNTIME"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.OptionButton optUPC 
         Caption         =   "UPC"
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
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "Description"
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.ComboBox cboSelection 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame fraCategory 
      Caption         =   "Category:"
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
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton optLiquor 
         Caption         =   "Liquor"
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
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton optWine 
         Caption         =   "Wine"
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
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optBeer 
         Caption         =   "Beer"
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
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optAccessories 
         Caption         =   "Accessories"
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "A UserID will be automaticly genereted"
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmReceiving"
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

Private Sub cboSelection_Click()
''''''''''''''''''''''''''''''''''this sub validates the current selection, making sure that the user
''''''''''''''''''''''''''''''''''doesn't mess around with it too much
    Dim strUPCToUpdate As String
    Dim blnFound As Boolean
    ''''''''''''''''''''''''''''''''''''if the user hasn't updated the current recordset ask if they want to
    If cmdUpdate.Enabled = True Then
        gstrMessage = "Do you wish to complete the current transaction before selecting another product?"
        gstrTitle = "Finalize the current transaction  ..."
        gintStyle = vbYesNo + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        If gintAnswer = vbYes Then
            txtReceived.SetFocus
            Exit Sub
        End If
    End If
    ''''''''''''''''''''''''''''''if they want to then do it
    '''**************************Connect to recordset and retrieve current stock value
    blnFound = False
    strUPCToUpdate = Mid(Tag, 1, 1) & Format(cboSelection.ItemData(cboSelection.ListIndex), "0000")
    mintFileHandle = FreeFile
    Open App.Path & "\Data\Stock.dat" For Input As #mintFileHandle
    Do While Not blnFound
        Input #mintFileHandle, mstrUPC, msngQuantity
        If mstrUPC = strUPCToUpdate Then
            lblActualData.Caption = msngQuantity
            blnFound = True
        End If
    Loop
    Close #mintFileHandle
  '''''''''''''''''''''''''''''''''''''''''''
    cmdUpdate.Enabled = True
    lblUpdated.Caption = ""
    lblUpdatedData.Caption = ""
    txtReceived.Text = "0"
    txtReceived.SetFocus

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdUpdate_Click()

    Dim strUPCToUpdate As String
    Dim strCurrentTransactionID As String
    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    Dim blnFound As Boolean
    
    If cboSelection.ListIndex = -1 Then
        gstrMessage = "Please select a product first!"
        gstrTitle = "Product missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        Exit Sub
    End If
    
    'update the inventory
    strUPCToUpdate = Mid(Tag, 1, 1) & Format(cboSelection.ItemData(cboSelection.ListIndex), "0000")
    intFileHandleInput = FreeFile
    Open App.Path & "\Data\Stock.dat" For Input As #intFileHandleInput
    intFileHandleOutput = FreeFile
    Open App.Path & "\Data\TempStock.dat" For Output As #intFileHandleOutput
    Do While Not EOF(intFileHandleInput)
        Input #intFileHandleInput, mstrUPC, msngQuantity
        If mstrUPC <> strUPCToUpdate Then
            Write #intFileHandleOutput, mstrUPC, msngQuantity
        End If
    Loop
    msngQuantity = Val(txtReceived.Text) + Val(lblActualData.Caption)
    lblUpdatedData.Caption = msngQuantity
    Write #intFileHandleOutput, strUPCToUpdate; msngQuantity
    Close
    Kill App.Path & "\Data\Stock.dat"
    Name App.Path & "\Data\TempStock.dat" As App.Path & "\Data\Stock.dat"
    
    'get the price of item
    mintFileHandle = FreeFile
    Open mstrFileName For Input As #mintFileHandle
    Do While Not blnFound
        Input #mintFileHandle, mstrUPC, mstrName, mintCountry, mintType, mintVolume, msngPrice
        If mstrUPC = strUPCToUpdate Then
            blnFound = True
        End If
    Loop
    Close
    
    'write the transaction to file
    strCurrentTransactionID = GetNextTransactionID()
    mintFileHandle = FreeFile
    Open App.Path & "\Data\Transaction.dat" For Append As #mintFileHandle
    Write #mintFileHandle, strCurrentTransactionID; "R"; Format(Date, "mm/dd/yyyy"); strUPCToUpdate _
                                          ; Val(lblActualData.Caption); Val(txtReceived.Text); msngQuantity; msngPrice
    Close
    GenerateNextTransactionID (strCurrentTransactionID)
    cmdUpdate.Enabled = False
    gstrMessage = "Transaction completed!"
    gstrTitle = "Update status ..."
    gintStyle = vbOKOnly + vbExclamation
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)

End Sub

Private Sub Form_Load()

    optBeer_Click

End Sub

Private Sub optAccessories_Click()

    lblActualData.Caption = ""
    txtReceived.Text = "0"
    lblUpdatedData.Caption = ""
    cmdUpdate.Enabled = False
    If optAccessories.Value = True Then
        mstrFileName = App.Path & "\Data\" & optAccessories.Caption & ".dat"
        Tag = optAccessories.Caption
    End If
    optInvisible.Value = True
    cboSelection.Clear

End Sub

Private Sub optBeer_Click()

    lblActualData.Caption = ""
    txtReceived.Text = "0"
    lblUpdatedData.Caption = ""
    cmdUpdate.Enabled = False
    If optBeer.Value = True Then
        mstrFileName = App.Path & "\Data\" & optBeer.Caption & ".dat"
        Tag = optBeer.Caption
    End If
    optInvisible.Value = True
    cboSelection.Clear
    
End Sub

Private Sub optDescription_Click()

    cboSelection.Clear
    mintFileHandle = FreeFile
    Open mstrFileName For Input As #mintFileHandle
    Do While Not EOF(mintFileHandle)
        Input #mintFileHandle, mstrUPC, mstrName, mintCountry, mintType, mintVolume, msngPrice
        If Tag = "Accessories" Then
            cboSelection.AddItem ReturnType(Tag, mintType) & ", " & mstrName
        Else
            cboSelection.AddItem ReturnType(Tag, mintType) & ", " & mstrName & ", " & ReturnVolume(mintVolume)
        End If
        cboSelection.ItemData(cboSelection.NewIndex) = Val(Mid(mstrUPC, 2, 4))
    Loop
    Close #mintFileHandle
    
End Sub

Private Sub optLiquor_Click()

    lblActualData.Caption = ""
    txtReceived.Text = "0"
    lblUpdatedData.Caption = ""
    cmdUpdate.Enabled = False
    If optLiquor.Value = True Then
        mstrFileName = App.Path & "\Data\" & optLiquor.Caption & ".dat"
        Tag = optLiquor.Caption
    End If
    optInvisible.Value = True
    cboSelection.Clear

End Sub

Private Sub optUPC_Click()

    cboSelection.Clear
    mintFileHandle = FreeFile
    Open mstrFileName For Input As #mintFileHandle
    Do While Not EOF(mintFileHandle)
        Input #mintFileHandle, mstrUPC, mstrName, mintCountry, mintType, mintVolume, msngPrice
        cboSelection.AddItem mstrUPC
        cboSelection.ItemData(cboSelection.NewIndex) = Val(Mid(mstrUPC, 2, 4))
    Loop
    Close #mintFileHandle

End Sub

Private Sub optWine_Click()

    lblActualData.Caption = ""
    txtReceived.Text = "0"
    lblUpdatedData.Caption = ""
    cmdUpdate.Enabled = False
    If optWine.Value = True Then
        mstrFileName = App.Path & "\Data\" & optWine.Caption & ".dat"
        Tag = optWine.Caption
    End If
    optInvisible.Value = True
    cboSelection.Clear

End Sub

Private Sub txtReceived_GotFocus()

    txtReceived.SelStart = 0
    txtReceived.SelLength = Len(txtReceived.Text)

End Sub

Private Sub txtReceived_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 8
        Case 13
            cmdUpdate_Click
        Case 27
            cmdCancel_Click
        Case 48 To 57
        Case Else
            Beep
            MsgBox "Only numerical digits accepted", vbOKOnly, "Incorrect data type"
            KeyAscii = 0
    End Select

End Sub
