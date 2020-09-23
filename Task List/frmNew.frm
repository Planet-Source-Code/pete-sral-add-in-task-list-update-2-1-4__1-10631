VERSION 5.00
Begin VB.Form frmNew 
   Caption         =   "New Task"
   ClientHeight    =   1815
   ClientLeft      =   5355
   ClientTop       =   2655
   ClientWidth     =   4215
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4215
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3060
      TabIndex        =   5
      Top             =   1440
      Width           =   1065
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   1905
      TabIndex        =   4
      Top             =   1440
      Width           =   1065
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   975
      Width           =   4035
   End
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   4035
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H8000000D&
      Caption         =   "<< VERSION >>"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1560
      Width           =   1710
   End
   Begin VB.Label lblDate 
      BackColor       =   &H80000017&
      Caption         =   "<< DATE >>"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1365
      Width           =   1710
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   735
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Task Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   1650
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TaskDesc As String
Public TaskPriority As String


Private Sub Command1_Click()

End Sub

Private Sub cmdAdd_Click()
    
   TaskDesc = Trim$(txtDesc.Text)
    TaskPriority = cboPriority.Text
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
   TaskDesc = ""
    TaskPriority = ""
    Unload Me
End Sub

Private Sub Form_Load()

'note - string sort
' 10 will come after 1
' add a 0 if you need to go past 9
' ie 01, 02, 03...
    lblDate.Caption = Format(Now, "mm/dd/yyyy")
    lblVersion.Caption = GetCurrentVersion()
    With cboPriority
        .AddItem "0 - None"
        .AddItem "1 - Low"
        .AddItem "2 - Medium/Low"
        .AddItem "3 - Medium"
        .AddItem "4 - Medium/High"
        .AddItem "5 - High"
        .AddItem "6 - Very High"
        .AddItem "7 - Highest"
    End With
    
    'last one
    cboPriority.ListIndex = cboPriority.ListCount - 1
End Sub


Private Sub Label3_Click()

End Sub
