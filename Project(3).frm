VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8430
   LinkTopic       =   "Form3"
   ScaleHeight     =   6060
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Kembali"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Row"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove"
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   4080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Project(3).frx":0000
      OLEDBString     =   $"Project(3).frx":00B4
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TInventaris"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Faktur Terima Barang"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Faktur Keluar Barang"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Barang Pesanan"
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Inventaris"
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Barang"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Project(3).frx":0168
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "TBarang"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "TPemasok"
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "TInventaris"
Adodc1.Refresh
End Sub

Private Sub Command5_Click()
Dim hapus
hapus = MsgBox("Yakin ingin menghapus data?", vbQuestion + vbYesNo, "menghapus data")
If hapus = vbYes Then
    Adodc1.Recordset.Delete
End If
End Sub

Private Sub Command6_Click()
Unload Form3
Form1.Show
End Sub

Private Sub Command7_Click()
Adodc1.RecordSource = "TInvoice"
Adodc1.Refresh
End Sub

Private Sub Command8_Click()
Adodc1.RecordSource = "TFakturTerimaBarang"
Adodc1.Refresh
End Sub

Private Sub Command9_Click()
Unload Form3
Form4.Show
End Sub
