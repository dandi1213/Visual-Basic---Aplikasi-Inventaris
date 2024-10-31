VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   Picture         =   "Form1.1.frx":0000
   ScaleHeight     =   6645
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H006EE1C9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6720
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H006EE1C9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   8040
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   8040
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "user" And Text2.Text = "user" Then
    Unload Form1
    Form2.Show
ElseIf Text1.Text = "admin" And Text2.Text = "admin" Then
    Unload Form1
    Form4.Show
Else
    MsgBox "Wrong username or password", vbCritical, "Info"
End If
End Sub

Private Sub Command2_Click()
Unload Form1
End Sub

Private Sub Image1_Click()
Unload Form1
End Sub

Private Sub Image2_Click()
If Text1.Text = "user" And Text2.Text = "user" Then
    Unload Form1
    Form2.Show
ElseIf Text1.Text = "admin" And Text2.Text = "admin" Then
    Unload Form1
    Form4.Show
Else
    MsgBox "Wrong username or password", vbCritical, "Info"
End If
End Sub

