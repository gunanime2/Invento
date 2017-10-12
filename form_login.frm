VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form form_login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2805
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5010
   Icon            =   "form_login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1657.287
   ScaleMode       =   0  'User
   ScaleWidth      =   4704.119
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adodc_login 
      Height          =   330
      Left            =   2400
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adodc_login"
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
   Begin VB.TextBox text_username 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   2010
      TabIndex        =   1
      Top             =   1200
      Width           =   2325
   End
   Begin VB.CommandButton command_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1215
      TabIndex        =   4
      Top             =   2085
      Width           =   1140
   End
   Begin VB.CommandButton command_cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2820
      TabIndex        =   5
      Top             =   2085
      Width           =   1140
   End
   Begin VB.TextBox text_password 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2010
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1590
      Width           =   2325
   End
   Begin VB.Shape shape_dot 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   825
      TabIndex        =   0
      Top             =   1215
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   825
      TabIndex        =   2
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label label_header_app_name 
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape shape_header_background 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -240
      Top             =   -240
      Width           =   5295
   End
End
Attribute VB_Name = "form_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub command_cancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Call form_main.app_terminate
    Exit Sub
End Sub

Private Sub command_ok_Click()
    'Check if the username and password exists in the same record entry in the database.
    With adodc_login
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
        .RecordSource = "select * from accounts where account_username = '" & text_username.Text & "' and account_password = '" & text_password.Text & "'"
        .Refresh
        If .Recordset.RecordCount > 0 Then
            form_main.command_show_reports_frame.Enabled = True
            form_main.command_show_settings_frame.Enabled = True
            form_main.command_show_stocks_frame.Enabled = True
            form_main.command_logout.Enabled = True
            form_main.label_username.Caption = text_username.Text
            form_main.label_account_id.Caption = .Recordset!account_id
            
            text_password.Text = ""
            text_username.Text = ""
            
            user_logged_in = True
            LoginSucceeded = True
            form_main.command_show_stocks_frame.SetFocus
            Unload Me
        Else
            MsgBox "Invalid Username or Password, try again!", vbExclamation, "Login"
            text_password.SetFocus
        End If
    End With
End Sub

Private Sub Form_Load()
    If user_logged_in = True Then
        form_main.Show
    Else
        form_main.Show
        form_main.command_show_reports_frame.Enabled = False
        form_main.command_show_settings_frame.Enabled = False
        form_main.command_show_stocks_frame.Enabled = False
        form_main.command_logout.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call form_main.app_terminate
End Sub
