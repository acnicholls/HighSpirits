VERSION 5.00
Begin VB.Form frmAdministrator 
   Caption         =   "Managing Users (Administrator only)"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   6
      Top             =   1440
      Width           =   4335
      Begin VB.OptionButton optID 
         Caption         =   "ID"
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
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optName 
         Caption         =   "Name"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboSelection 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame fraOperation 
      Caption         =   "Operation:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.OptionButton optDelete 
         Caption         =   "Delete User"
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
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optModify 
         Caption         =   "Modify User"
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
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Add User"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "A UserID will be automaticly genereted"
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
      Height          =   615
      Left            =   1320
      TabIndex        =   10
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "frmAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intCounter As Integer
Private intFileHandle As Integer
Private strID As String
Private strLastName As String
Private strFirstName As String
Private strRights As String
Private strPassword As String

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdExecute_Click()

    Dim blnFound As Boolean
    
    'Check if a user is selected
    If optAdd.Value = False Then
        If cboSelection.ListIndex = -1 Then
            gstrMessage = "Please select a User first!"
            gstrTitle = "User missing ..."
            gintStyle = vbOKOnly + vbExclamation
            gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
            Exit Sub
        End If
    End If
    
    'Load frmUserAdministration with the correct option
    blnFound = False
    If optAdd.Value = True Then
        With frmUserAdministration
            .Tag = "Add"
            'Generate the next avalaible user ID
            .lblIDData.Caption = GetNextUserID()
        End With
    Else
        intFileHandle = FreeFile
        Open App.Path & "\Data\Users.dat" For Input As #intFileHandle
        'Search the file until the user is found
        Do While Not blnFound
            Input #intFileHandle, strID, strLastName, strFirstName, strRights, strPassword
            If optName.Value = True Then
                If strLastName & ", " & strFirstName = cboSelection.Text Then
                    'When found,  retreive information
                    With frmUserAdministration
                        .lblIDData.Caption = strID
                        .txtLastName.Text = strLastName
                        .txtFirstName.Text = strFirstName
                        .txtPassword.Text = strPassword
                        For intCounter = 1 To 5
                            'Retreive the rights
                            If InStr(strRights, Mid(.chkRights(intCounter).Caption, 2, 1)) > 0 Then
                                .chkRights(intCounter).Value = vbChecked
                            End If
                        Next intCounter
                    End With
                    blnFound = True
                End If
            Else
                If strID = cboSelection.Text Then
                    With frmUserAdministration
                        .lblIDData.Caption = strID
                        .txtLastName.Text = strLastName
                        .txtFirstName.Text = strFirstName
                        .txtPassword.Text = strPassword
                        For intCounter = 1 To 5
                            'Retreive the rights
                            If InStr(strRights, Mid(.chkRights(intCounter).Caption, 2, 1)) > 0 Then
                                .chkRights(intCounter).Value = vbChecked
                            End If
                        Next intCounter
                    End With
                    blnFound = True
                End If
            End If
        Loop
        'Close the file
        Close #intFileHandle
        'Assign the Tag property of the form to the type of operation
        If optModify.Value = True Then
            frmUserAdministration.Tag = "Modify"
        Else
            frmUserAdministration.Tag = "Delete"
        End If
    End If
    
    frmUserAdministration.Show vbModal

End Sub

Private Sub optAdd_Click()

    'If option is Add then fraSearch should be invisible
    If optAdd.Value = True Then
        fraSearch.Visible = False
    Else
        fraSearch.Visible = True
    End If

End Sub

Private Sub optDelete_Click()

    'If option is Add then fraSearch should be visible
    If optDelete.Value = False Then
        fraSearch.Visible = True
    Else
        fraSearch.Visible = False
    End If


End Sub

Private Sub optID_Click()

    'Fill cboSelection with information
    cboSelection.Clear
    intFileHandle = FreeFile
    Open App.Path & "\Data\Users.dat" For Input As #intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strID, strLastName, strFirstName, strRights, strPassword
        cboSelection.AddItem strID
    Loop
    Close #intFileHandle

End Sub

Private Sub optModify_Click()

    'If option is Add then fraSearch should be visible
    If optModify.Value = True Then
        fraSearch.Visible = True
    Else
        fraSearch.Visible = False
    End If

End Sub

Private Sub optName_Click()

    'Fill cboSelection with information
    cboSelection.Clear
    intFileHandle = FreeFile
    Open App.Path & "\Data\Users.dat" For Input As #intFileHandle
    Do While Not EOF(intFileHandle)
        Input #intFileHandle, strID, strLastName, strFirstName, strRights, strPassword
        cboSelection.AddItem strLastName & ", " & strFirstName
    Loop
    Close #intFileHandle

End Sub
