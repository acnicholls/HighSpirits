VERSION 5.00
Begin VB.Form frmModifyProducts 
   Caption         =   "Modify Product"
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
      ItemData        =   "frmModifyProducts.frx":0000
      Left            =   1440
      List            =   "frmModifyProducts.frx":0019
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
      ItemData        =   "frmModifyProducts.frx":0054
      Left            =   1440
      List            =   "frmModifyProducts.frx":0070
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
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
Attribute VB_Name = "frmModifyProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
'''''''''''''''''''''''''''''''''''''''''''''''This sub takes the current screen values and assigns them to variable
''''''''''''''''''''''''''''''''''''''''''''''''then updates the current record with the values assigned

    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    Dim strUPC As String
    Dim strName As String
    Dim intCountry As Integer
    Dim intType As Integer
    Dim intVolume As Integer
    Dim sngPrice As Single
    Dim blnResult As Boolean
    ''''''''''''''''''''''''''make sure there is a name
    If Len(txtName.Text) = 0 Then
        gstrMessage = "Please enter a Description in order to Modify this product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtName.SetFocus
        Exit Sub
    End If
    ''''''''''''''''''''''''''''make sure there is a country
    If cboCountry.ListIndex = -1 And cboCountry.Enabled = True Then
        gstrMessage = "Please select a Country in order to Modify this product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        cboCountry.SetFocus
        Exit Sub
    End If
    '''''''''''''''''''''''make sure there is a type
    If cboType.ListIndex = -1 Then
        gstrMessage = "Please select a Type in order to Modify this product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        cboType.SetFocus
        Exit Sub
    End If
    '''''''''''''''''''''''''make sure there is a volume
    If cboVolume.ListIndex = -1 And cboVolume.Enabled = True Then
        gstrMessage = "Please select a Volume in order to Modify this product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        cboVolume.SetFocus
        Exit Sub
    End If
    ''''''''''''''''''''''''''''make sure there is a price
    If Val(txtPrice.Text) <= 0 Then
        gstrMessage = "Please enter a valid Price in order to Modify this product!"
        gstrTitle = "Information missing ..."
        gintStyle = vbOKOnly + vbExclamation
        gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
        txtPrice.SetFocus
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''confirm user wants to modify
    strUPC = lblUPCData.Caption
    gstrMessage = "Information on Product #" & strUPC
    gstrMessage = gstrMessage & " will be permanently modify.  Do you wish to proceed?"
    gstrTitle = "Updating Product ..."
    gintStyle = vbYesNo + vbExclamation
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
    If gintAnswer = vbNo Then
        Exit Sub
    End If
    
    intFileHandleInput = FreeFile
    Open App.Path & "\Data\" & Tag & ".dat" For Input As #intFileHandleInput
    intFileHandleOutput = FreeFile
    Open App.Path & "\Data\Temp" & Tag & ".dat" For Output As #intFileHandleOutput
    Do While Not EOF(intFileHandleInput)
        Input #intFileHandleInput, strUPC, strName, intCountry, intType, intVolume, sngPrice
        If strUPC <> lblUPCData.Caption Then
            Write #intFileHandleOutput, strUPC; strName; intCountry; intType; intVolume; sngPrice
        End If
    Loop
''''''''''''''''''''''''''''''''''''  New info is entered assign variables the values of the info
    
    strUPC = lblUPCData.Caption
    strName = txtName.Text
    intCountry = cboCountry.ListIndex
    intType = cboType.ListIndex
    intVolume = cboVolume.ListIndex
    sngPrice = Val(txtPrice.Text * 100) / 100
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''Here I must connect to the recordset and update the current record
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
    Write #intFileHandleOutput, strUPC; strName; intCountry; intType; intVolume; sngPrice
    Close
    Kill App.Path & "\Data\" & Tag & ".dat"
    Name App.Path & "\Data\Temp" & Tag & ".dat" As App.Path & "\Data\" & Tag & ".dat"
 ''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''let the user know the values were updated
 
    gstrMessage = "Update on Product #" & strUPC
    gstrMessage = gstrMessage & " is completed."
    gstrTitle = "Updating Product ..."
    gintStyle = vbOKOnly
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
''''''''''''''''''''back to product selector
    Unload Me

End Sub

Private Sub cmdCancel_Click()
''''''''''''''back to product selector
    Unload Me

End Sub


Private Sub Form_Activate()
'''''''''''''''''''''''''''''''''''''''''''''When the form is called put these values in the selection box
'''''''''''''''''''''''''''''''''''''''''''''This could be modified to use the database, but this works so why bother!!

    Dim intItemSelected As Integer
    
    intItemSelected = cboType.ListIndex
    cboType.Clear
    
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
    
    cboType.ListIndex = intItemSelected

End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''this sub fills cboType will temp values to allow for the fill
    cboType.AddItem "Temp Value #1"
    cboType.AddItem "Temp Value #2"
    cboType.AddItem "Temp Value #3"
    cboType.AddItem "Temp Value #4"
    cboType.AddItem "Temp Value #5"
    cboType.AddItem "Temp Value #6"
    cboType.AddItem "Temp Value #7"
    cboType.AddItem "Temp Value #8"
    cboType.AddItem "Temp Value #9"
    cboType.AddItem "Temp Value #10"

End Sub
