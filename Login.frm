VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   3480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Login.frx":0000
      Top             =   120
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\Programs\Login\Login 1\Login.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Login"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label MemID 
      DataField       =   "MemID"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Pass 
      DataField       =   "Pass"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Member ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Login.Data1.Recordset.FindFirst "memID = '" & Login.Text1.Text & "'"
    If Login.Pass.Caption = Login.Text2.Text Then
        MsgBox "Login Successful!"
        Login.MemID.Caption = ""
        Login.Pass.Caption = ""
        Login.Text1.Text = ""
        Login.Text2.Text = ""
        Exit Sub
    End If
    MsgBox "Login Unsuccessful!"
    Login.Text1.Text = ""
    Login.Text2.Text = ""
End Sub

Private Sub Command2_Click()
    Login.Data1.Recordset.AddNew
    Login.Data1.Recordset.Fields("memID") = "" & Login.Text1.Text & ""
    Login.Data1.Recordset.Fields("pass") = "" & Login.Text2.Text & ""
    Login.Data1.Recordset.Update
    Login.MemID.Caption = ""
    Login.Pass.Caption = ""
    Login.Text1.Text = ""
    Login.Text2.Text = ""
End Sub

Private Sub Command4_Click()
    Login.Command5.Visible = True
    Login.Command4.Visible = False
    Login.Width = 3465
End Sub

Private Sub Command5_Click()
    Login.Command4.Visible = True
    Login.Command5.Visible = False
    Login.Width = 5985
End Sub
