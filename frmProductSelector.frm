VERSION 5.00
Begin VB.Form frmProductSelector 
   Caption         =   "Modify Product ..."
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
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
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5655
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
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   1695
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
         TabIndex        =   10
         Top             =   360
         Width           =   735
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
         Left            =   1320
         TabIndex        =   9
         Top             =   420
         Width           =   1095
      End
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
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
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
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   5655
      Begin VB.OptionButton optInvisible 
         BackColor       =   &H000000FF&
         Caption         =   "THIS RADIO BUTTON IS NOT VISIBLE AT RUNTIME"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.ComboBox cboSelection 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   3495
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
         TabIndex        =   3
         Top             =   360
         Width           =   1455
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
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "A UserID will be automaticly genereted"
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "frmProductSelector"
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

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdExecute_Click()

    Dim blnFound As Boolean
    Dim strUPCToModify As String
    
    If cboSelection.ListIndex = -1 Then
        gstrMessage = "Please select a product first!"
        gstrTitle = "Product missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        Exit Sub
    End If
    
    'get tinformation on selected item
    strUPCToModify = Mid(Tag, 1, 1) & Format(cboSelection.ItemData(cboSelection.ListIndex), "0000")
    mintFileHandle = FreeFile
    Open mstrFileName For Input As #mintFileHandle
    Do While Not blnFound
        Input #mintFileHandle, mstrUPC, mstrName, mintCountry, mintType, mintVolume, msngPrice
        If mstrUPC = strUPCToModify Then
            blnFound = True
        End If
    Loop
    Close

    With frmModifyProducts
        .Tag = Tag
        .Caption = "Updating Information to Database (Category: " & Tag & ")"
        .lblUPCData.Caption = mstrUPC
        .txtName.Text = mstrName
        .cboCountry.ListIndex = mintCountry
        .cboType.ListIndex = mintType
        .cboVolume.ListIndex = mintVolume
        .txtPrice.Text = Format(msngPrice, "0.00")
        If Tag = "Accessories" Then
            .cboCountry.Enabled = False
            .cboVolume.Enabled = False
        End If
        .Show vbModal
    End With
    Set frmModifyProducts = Nothing

End Sub

Private Sub optAccessories_Click()

    If optAccessories.Value = True Then
        mstrFileName = App.Path & "\Data\" & optAccessories.Caption & ".dat"
        Tag = optAccessories.Caption
        optInvisible.Value = True
        cboSelection.Clear
    End If

End Sub

Private Sub optBeer_Click()

    If optBeer.Value = True Then
        mstrFileName = App.Path & "\Data\" & optBeer.Caption & ".dat"
        Tag = optBeer.Caption
        optInvisible.Value = True
        cboSelection.Clear
    End If

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

    If optLiquor.Value = True Then
        mstrFileName = App.Path & "\Data\" & optLiquor.Caption & ".dat"
        Tag = optLiquor.Caption
        optInvisible.Value = True
        cboSelection.Clear
    End If

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

    If optWine.Value = True Then
        mstrFileName = App.Path & "\Data\" & optWine.Caption & ".dat"
        Tag = optWine.Caption
        optInvisible.Value = True
        cboSelection.Clear
    End If

End Sub
