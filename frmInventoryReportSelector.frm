VERSION 5.00
Begin VB.Form frmInventoryReportSelector 
   Caption         =   "Inventory Report Seloector"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   4200
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
      Height          =   2535
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton optCategory 
         Caption         =   "ALL"
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
         Index           =   5
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton optCategory 
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
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton optCategory 
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
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optCategory 
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
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optCategory 
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
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "Please, select the category of items you want to be included in the report."
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
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmInventoryReportSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    'Call to ReportsHeading depending om the form Tag property value
    'ReportsHeading function as 2 parameters: Report title and printing "Orientation"
    ' 1 is for Portrait
    ' 2 is for Landscape
    Select Case Tag
        Case ""
            gstrMessage = "Please select a category first!"
            gstrTitle = "Report warning ..."
            gintStyle = vbOKOnly + vbExclamation
            gintAnswer = MsgBox(gstrMessage, gintStyle, gstrTitle)
            Exit Sub
        Case "ALL"
            ReportsHeading "TOTAL INVENTORY", 1
        Case Else
            ReportsHeading "INVENTORY (" & Tag & " only)", 1
    End Select
    
    InventoryReport (Tag)
    ReportCompletedMessage
    Unload Me

End Sub

Private Sub optCategory_Click(Index As Integer)

    'Tag property equals the name of the option selected
    Tag = optCategory(Index).Caption

End Sub
