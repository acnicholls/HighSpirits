VERSION 5.00
Begin VB.Form frmNewProducts 
   Caption         =   "Add New Product"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCountry 
      Height          =   315
      ItemData        =   "frmNewProducts.frx":0000
      Left            =   1440
      List            =   "frmNewProducts.frx":0019
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboVolume 
      Height          =   315
      ItemData        =   "frmNewProducts.frx":0054
      Left            =   1440
      List            =   "frmNewProducts.frx":0070
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblCountry 
      Caption         =   "Country:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblPrice 
      Caption         =   "Price:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblType 
      Caption         =   "Type:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
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
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblUPC 
      Caption         =   "UPC:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblUPCData 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmNewProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
'''''''''''''''''''''''''''''''''''''''''''''''  This sub adds a new product to the database
'''''''''''''''''''''''''''''''''''''''''''''''  it first takes the screen values and converts to variables, then appends this record to the db

    Dim intFileHandle As Integer
    Dim strUPC As String
    Dim strName As String
    Dim intCountry As Integer
    Dim intType As Integer
    Dim intVolume As Integer
    Dim sngPrice As Single
    Dim blnResult As Boolean
''''''''''''''''''''''''''''''''''''''''make sure there is a name
    If Len(txtName.Text) = 0 Then
        gstrMessage = "Please enter a Description in order to add this new product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtName.SetFocus
        Exit Sub
    End If
'''''''''''''''''''''''''''''''''''''make sure there is a country
    If cboCountry.ListIndex = -1 And cboCountry.Enabled = True Then
        gstrMessage = "Please select a Country in order to add this new product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        cboCountry.SetFocus
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''' make sure there is a type
    If cboType.ListIndex = -1 Then
        gstrMessage = "Please select a Type in order to add this new product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        cboType.SetFocus
        Exit Sub
    End If
    '''''''''''''''''''''''''''make sure there is a volume, if required
    If cboVolume.ListIndex = -1 And cboVolume.Enabled = True Then
        gstrMessage = "Please select a Volume in order to add this new product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        cboVolume.SetFocus
        Exit Sub
    End If
    ''''''''''''''''''''''''''make sure there is a price
    If Val(txtPrice.Text) <= 0 Then
        gstrMessage = "Please enter a valid Price in order to add this new product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtPrice.SetFocus
        Exit Sub
    End If
    '''''''''''''''''''''''''''''converts screen values to variables
    strUPC = lblUPCData.Caption
    strName = txtName.Text
    intCountry = cboCountry.ListIndex
    intType = cboType.ListIndex
    intVolume = cboVolume.ListIndex
    sngPrice = Val(txtPrice.Text)
'
    intFileHandle = FreeFile
    Open App.Path & "\Data\" & Tag & ".dat" For Append As #intFileHandle
    Write #intFileHandle, strUPC; strName; intCountry; intType; intVolume; sngPrice
    Close #intFileHandle
'''''''''''''''''''''''''''''''''''''''Connect here and append to recordset

'''''''''''''''''''''''''''''''''''''this generates a new UPC for the next new item
    blnResult = GenerateNextUPC(Tag)
    '''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''connect here and update stock recordsset
    
    intFileHandle = FreeFile
    Open App.Path & "\Data\Stock.dat" For Append As #intFileHandle
    Write #intFileHandle, strUPC; 0
    Close #intFileHandle
    ''''''''''''''''''''''''''''''''''''''''''''''''''''ask user if another must be added
    gstrMessage = "Do you want to add another " & Tag & "?"
    gstrTitle = "Add New " & Tag & " ..."
    gintStyle = vbYesNo + vbQuestion
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
    If gintAnswer = vbNo Then
        Unload Me
    End If
    '''''''''''''''''''''''''''''''''''''''''''''if yes prepare the form
    lblUPCData.Caption = GetNextUPC(Tag)
    txtName.Text = ""
    cboCountry.ListIndex = -1
    cboType.ListIndex = -1
    cboVolume.ListIndex = -1
    txtPrice.Text = ""

End Sub

Private Sub cmdCancel_Click()
'''''''''''''''''''''''''''''''''''return to main screen
    Unload Me

End Sub

Private Sub Form_Activate()
''''''''''''''''''''''''''''''''when the form is needed cbotype must be filled
    Select Case Tag
        Case "Beer"
            cboType.AddItem "Alcohol free"
            cboType.AddItem "Ale"
            cboType.AddItem "Dark"
            cboType.AddItem "Lager"
            cboType.AddItem "Micro brewed"
        Case "Wine"
            cboType.AddItem "Red"
            cboType.AddItem "Rose"
            cboType.AddItem "Sparkling"
            cboType.AddItem "White"
        Case "Liquor"
            cboType.AddItem "Aperitif Wine"
            cboType.AddItem "Brandy"
            cboType.AddItem "Cognac"
            cboType.AddItem "Gin"
            cboType.AddItem "Grappa"
            cboType.AddItem "Ouzo"
            cboType.AddItem "Rhum"
            cboType.AddItem "Scotch"
            cboType.AddItem "Vodka"
            cboType.AddItem "Whiskey"
        Case "Accessories"
            cboType.AddItem "Books"
            cboType.AddItem "Bottle openers"
            cboType.AddItem "Gift Packs"
            cboType.AddItem "Glasses"
            cboType.AddItem "Wine Racks"
    End Select

End Sub
