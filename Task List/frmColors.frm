VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Grid Colors"
   ClientHeight    =   2550
   ClientLeft      =   4440
   ClientTop       =   1695
   ClientWidth     =   6165
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdColors 
      Left            =   2835
      Top             =   1035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   360
      Left            =   4335
      TabIndex        =   4
      Top             =   2055
      Width           =   1665
   End
   Begin VB.CommandButton cmdColor2 
      Caption         =   "Set Color 2"
      Height          =   360
      Left            =   2235
      TabIndex        =   3
      Top             =   2055
      Width           =   1665
   End
   Begin VB.CommandButton cmdColor1 
      Caption         =   "Set Color 1"
      Height          =   360
      Left            =   90
      TabIndex        =   2
      Top             =   2055
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin MSComctlLib.ListView lstTasks 
      Height          =   1965
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   3466
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Task"
         Object.Width           =   5365
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Added"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Priority"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Completed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Version"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub cmdColor1_Click()
    
Dim sColor As String

    cdColors.ShowColor
    sColor = cdColors.Color
    If Len(sColor) <> 0 Then
        If WriteToIni("GridColor", "Color1", sColor, gstrTaskFile) Then
            SetListViewColor lstTasks, Picture1
        End If
    End If
    lstTasks.Refresh
End Sub

Private Sub cmdColor2_Click()
Dim sColor As String

    cdColors.ShowColor
    sColor = cdColors.Color
    If Len(sColor) <> 0 Then
        If WriteToIni("GridColor", "Color2", sColor, gstrTaskFile) Then
            SetListViewColor lstTasks, Picture1
        End If
    End If
    lstTasks.Refresh
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

Dim itmx As ListItem

Set itmx = lstTasks.ListItems.Add(1)
        itmx.Text = "Color Line 1"
        
        
    Set itmx = lstTasks.ListItems.Add(2)
        itmx.Text = "Color Line 2"
    
   Set itmx = Nothing
   
   SetListViewColor lstTasks, Picture1
End Sub

Private Sub Form_Paint()
     SetListViewColor lstTasks, Picture1
End Sub

Private Sub lstTasks_ItemClick(ByVal Item As MSComctlLib.ListItem)
 If Item.Index = 1 Then
        cmdColor1_Click
    ElseIf Item.Index = 2 Then
        cmdColor2_Click
    End If
End Sub
