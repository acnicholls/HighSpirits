VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Login ...."
   ClientHeight    =   3255
   ClientLeft      =   2205
   ClientTop       =   2745
   ClientWidth     =   5265
   ForeColor       =   &H00000000&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserId 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      ForeColor       =   &H00800080&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox picKeys 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "frmLogin.frx":27A2
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblUserId 
      Caption         =   "&User ID:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password :"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblSecurityMessage 
      Caption         =   "Accessing this system requires that you enter your user ID as well as your valid password in the apropriate fields"
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
      Height          =   855
      Left            =   960
      TabIndex        =   6
      Top             =   195
      Width           =   4215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCounter As Integer
Private mstrUser As String

Private Sub cmdCancel_Click()

    Unload Me
    Set frmLogin = Nothing

End Sub

Private Sub cmdOK_Click()

    Dim intFileHandle As Integer
    Dim strID As String
    Dim strLastName As String
    Dim strFirstName As String
    Dim strRights As String
    Dim strPassword As String
    
    If txtUserId.Text = "" Then
        MsgBox " You must enter a User Id to proceed", vbOKOnly, "Missing field..."
        txtUserId.SetFocus
        Exit Sub
    End If
    
    If txtPassword.Text = "" Then
        MsgBox " You must enter a Password to proceed", vbOKOnly, "Missing field..."
        txtPassword.SetFocus
        Exit Sub
    End If
    
    If mintCounter < 3 Then
        If StrComp(LCase(txtUserId.Text), LCase("Administrator")) = 0 Then
            If StrComp(LCase(txtPassword.Text), LCase(gstrSuperPassword)) = 0 Then
               mstrUser = "admin"
            End If
        Else
          intFileHandle = FreeFile
            Open App.Path & "\Data\Users.dat" For Input As #intFileHandle
            Do While Not EOF(intFileHandle)
                Input #intFileHandle, strID, strLastName, strFirstName, strRights, strPassword
                    If StrComp(LCase(txtUserId.Text), LCase(strID)) = 0 Then
                        If StrComp(LCase(txtPassword.Text), LCase(strPassword)) = 0 Then
                            mstrUser = "regular"
                          Exit Do
                        End If
                    End If
            Loop
            Close #intFileHandle
        End If
        
        Select Case mstrUser
            Case "admin"
                gintStyle = vbOKOnly + vbExclamation
                gstrTitle = "Successful Login to High Spirits..."
                gstrMessage = "Welcome  " & vbNewLine & vbNewLine
                gstrMessage = gstrMessage & " Access to the system has been granted with administrator rights"
                MsgBox gstrMessage, gintStyle, gstrTitle
                Unload Me
                Set frmLogin = Nothing
                frmAdministrator.Show vbModal
                Set frmAdministrator = Nothing
            Case "regular"
                gintStyle = vbOKOnly + vbExclamation
                gstrTitle = "Successful Login to High Spirits..."
                gstrMessage = "Welcome  " & strFirstName & " " & strLastName & vbNewLine & vbNewLine
                gstrMessage = gstrMessage & " You have been granted access to the system with the following rights: " & " " & strRights
                MsgBox gstrMessage, gintStyle, gstrTitle
                Unload Me
                Set frmLogin = Nothing
                gstrCurrentUserID = strID
                gstrCurrentUserLastName = strLastName
                gstrCurrentUserFirstName = strFirstName
                gstrCurrentUserRights = strRights
                frmMain.Show vbModal
                Set frmMain = Nothing
            Case Else
                txtUserId.Text = ""
                txtPassword.Text = ""
                txtUserId.SetFocus
                mintCounter = mintCounter + 1
        End Select
    Else
        gintStyle = vbOKOnly + vbCritical
        gstrTitle = "High Spirits...Account Locked"
        gstrMessage = "You are restricted to three (3) attempts to login to the system" & vbNewLine
        gstrMessage = gstrMessage & " Please verify your login information with the system administrator!"
        MsgBox gstrMessage, gintStyle, gstrTitle
        Unload Me
    End If

End Sub

Private Sub Form_Load()

   mintCounter = 1
   mstrUser = ""

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdOK_Click
    End If

End Sub


Private Sub txtUserId_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtPassword.SetFocus
    End If
End Sub
