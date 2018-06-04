VERSION 5.00
Begin VB.Form frmUserAdministration 
   Caption         =   "User Administration ..."
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   ScaleHeight     =   3855
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame fraRights 
      Caption         =   "&User Rights:"
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
      Height          =   2415
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   2175
      Begin VB.CheckBox chkRights 
         Caption         =   "(S)hipping"
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
         Index           =   5
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkRights 
         Caption         =   "(R)eceiving"
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
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkRights 
         Caption         =   "(I)nformation Reports"
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
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkRights 
         Caption         =   "(B)illing"
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
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   750
         Width           =   1455
      End
      Begin VB.CheckBox chkRights 
         Caption         =   "(A)dministrator"
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
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   1680
      MaxLength       =   13
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtLastName 
      Height          =   375
      Left            =   1680
      MaxLength       =   13
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblIDData 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblID 
      Caption         =   "USER ID:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblFirstName 
      Caption         =   "&First Name:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblLastName 
      Caption         =   "&Last Name:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmUserAdministration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintFileHandle As Integer
Private mintCounter As Integer
Private mstrID As String
Private mstrLastName As String
Private mstrFirstName As String
Private mstrRights As String
Private mstrPassword As String

Private Sub cmdAdd_Click()

    Dim blnResult As Boolean
    
    mstrID = lblIDData.Caption
    mstrLastName = txtLastName.Text
    mstrFirstName = txtFirstName.Text
    mstrRights = ""
    For mintCounter = 1 To 5
        If chkRights(mintCounter).Value = vbChecked Then
            mstrRights = mstrRights & Mid(chkRights(mintCounter).Caption, 2, 1)
        End If
    Next mintCounter
    mstrPassword = txtPassword.Text
        
     mintFileHandle = FreeFile
     Open App.Path & "\Data\Users.dat" For Append As #mintFileHandle
     Write #mintFileHandle, mstrID; mstrLastName; mstrFirstName; mstrRights; mstrPassword
     Close #mintFileHandle
     
     blnResult = GenerateNextUserID(lblIDData.Caption)
     
     Unload Me

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    
    gstrMessage = "Information on User #" & lblIDData.Caption
    gstrMessage = gstrMessage & " will be permanently deleted.  Do you wish to proceed?"
    gstrTitle = "Delete User ..."
    gintStyle = vbYesNo + vbExclamation
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
    If gintAnswer = vbNo Then
        Exit Sub
    End If
    
    intFileHandleInput = FreeFile
    Open App.Path & "\Data\" & "Users.dat" For Input As #intFileHandleInput
    intFileHandleOutput = FreeFile
    Open App.Path & "\Data\" & "TempUsers.dat" For Output As #intFileHandleOutput
    Do While Not EOF(intFileHandleInput)
        Input #intFileHandleInput, mstrID, mstrLastName, mstrFirstName, mstrRights, mstrPassword
        If mstrID <> lblIDData.Caption Then
            Write #intFileHandleOutput, mstrID; mstrLastName; mstrFirstName; mstrRights; mstrPassword
        End If
    Loop
    Close #intFileHandleInput
    Close #intFileHandleOutput
    Kill App.Path & "\Data\" & "Users.dat"
    Name App.Path & "\Data\" & "TempUsers.dat" As App.Path & "\Data\" & "Users.dat"
    Unload Me

End Sub

Private Sub cmdModify_Click()

    Dim intFileHandleInput As Integer
    Dim intFileHandleOutput As Integer
    
    gstrMessage = "Information on User #" & lblIDData.Caption
    gstrMessage = gstrMessage & " will be permanently modify.  Do you wish to proceed?"
    gstrTitle = "Modify User ..."
    gintStyle = vbYesNo + vbExclamation
    gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
    If gintAnswer = vbNo Then
        Exit Sub
    End If
    
    intFileHandleInput = FreeFile
    Open App.Path & "\Data\" & "Users.dat" For Input As #intFileHandleInput
    intFileHandleOutput = FreeFile
    Open App.Path & "\Data\" & "TempUsers.dat" For Output As #intFileHandleOutput
    Do While Not EOF(intFileHandleInput)
        Input #intFileHandleInput, mstrID, mstrLastName, mstrFirstName, mstrRights, mstrPassword
        If mstrID <> lblIDData.Caption Then
            Write #intFileHandleOutput, mstrID; mstrLastName; mstrFirstName; mstrRights; mstrPassword
        End If
    Loop
    mstrID = lblIDData.Caption
    mstrLastName = txtLastName.Text
    mstrFirstName = txtFirstName.Text
    mstrRights = ""
    For mintCounter = 1 To 5
        If chkRights(mintCounter).Value = vbChecked Then
            mstrRights = mstrRights & Mid(chkRights(mintCounter).Caption, 2, 1)
        End If
    Next mintCounter
    mstrPassword = txtPassword.Text
    Write #intFileHandleOutput, mstrID; mstrLastName; mstrFirstName; mstrRights; mstrPassword
    Close
    Kill App.Path & "\Data\" & "Users.dat"
    Name App.Path & "\Data\" & "TempUsers.dat" As App.Path & "\Data\" & "Users.dat"
    Unload Me

End Sub

Private Sub Form_Activate()

    Select Case Tag
        Case "Add"
            cmdAdd.Enabled = True
        Case "Modify"
            cmdModify.Enabled = True
        Case "Delete"
            cmdDelete.Enabled = True
        Case Else
    End Select

End Sub
