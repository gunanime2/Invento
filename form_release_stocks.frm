VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form form_release_stocks 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Release Stocks Form"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7485
   Icon            =   "form_release_stocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton command_release_stock_submit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      Picture         =   "form_release_stocks.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox text_release_stock_reference_number 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox text_release_stock_quantity 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton option_pulled_out 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pulled-out"
      Height          =   375
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton option_sold 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sold"
      Height          =   375
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   960
      Value           =   -1  'True
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "form_release_stocks.frx":0884
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "item_name"
         Caption         =   "Item"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "item_srp"
         Caption         =   "SRP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """PHP""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "item_total_quantity"
         Caption         =   "Stocks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "item_date_time_created"
         Caption         =   "item_date_time_created"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "item_date_time_updated"
         Caption         =   "item_date_time_updated"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "item_category"
         Caption         =   "item_category"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "item_rsp"
         Caption         =   "item_rsp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "item_srp"
         Caption         =   "item_srp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "item_total_quantity"
         Caption         =   "item_total_quantity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "account_id"
         Caption         =   "account_id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "account_name"
         Caption         =   "account_name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   4080.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodc_release_stocks 
      Height          =   375
      Left            =   4200
      Top             =   6240
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Caption         =   "adodc_release_stocks"
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
   Begin VB.TextBox text_item_name 
      BackColor       =   &H00E0E0E0&
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reference No.:"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quantity:"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   5880
      Y1              =   960
      Y2              =   1320
   End
   Begin VB.Shape shape_dot 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Shape shape_header_background 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "form_release_stocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command_release_stock_submit_Click()
    If Val(text_release_stock_quantity.Text) < 1 Or text_release_stock_reference_number.Text = " " Then
        MsgBox "Please check your inputs.", vbExclamation, "System Message "
        Exit Sub
    End If
    'Add this release stock transaction into the items_transactions table.
    With adodc_release_stocks
        'Check if recordset is empty. If empty exit sub.
        If .ConnectionString = "" Or .RecordSource = "" Then
            MsgBox "No item selected.", vbExclamation, "System Message "
            Exit Sub
        End If
        
        'Declare and Get the needed variables.
        Dim item_id As Integer
        Dim transaction_quantity As Integer
        Dim transaction_reference_number As Single
        Dim transaction_sequence_number As Single
        Dim total_quantity_old As Integer
        Dim total_quantity_new As Integer
        Dim item_rsp As Double
        Dim item_srp As Double
        Dim transaction_id As Integer
        
        item_rsp = .Recordset("item_rsp")
        item_srp = .Recordset("item_srp")
        item_id = .Recordset("item_id")
        transaction_quantity = Val(text_release_stock_quantity.Text)
        transaction_reference_number = Val(text_release_stock_reference_number.Text)
        total_quantity_old = .Recordset("item_total_quantity")
        
        'Check if the current quantity is enough for the transaction quantity.
        If (transaction_quantity > total_quantity_old) Then
            MsgBox "Not enough stocks!", vbExclamation, "System Message"
            text_release_stock_quantity.SetFocus
            Exit Sub
        End If
        
        'Get the new total quantity of the current item.
        total_quantity_new = total_quantity_old - transaction_quantity
        
        'Get a transaction_sequence_number.
        .RecordSource = "select * from sequence_number order by sequence_number_id desc"
        .Refresh
        .Recordset.MoveFirst
        transaction_sequence_number = .Recordset("sequence_number")
        transaction_sequence_number = transaction_sequence_number + 1
        'Done getting sequence_number.
        
        'Get the release transaction type.
        Dim release_type As Integer
        If option_sold.Value = True Then
            release_type = 1 'Value 1 means sold release type transaction.
        Else
            release_type = 2 'Value 2 means pulled-out release type transaction.
        End If
        'Done getting release transaction type.
        
        'Start adding the transaction to the items_transactions database
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
        .RecordSource = "select * from items_transactions"
        .Refresh
        .Recordset.AddNew
        .Recordset!transaction_type = release_type
        .Recordset!transaction_type_reason = release_type
        .Recordset!transaction_date_time = Now
        .Recordset!transaction_quantity = transaction_quantity
        .Recordset!transaction_sequence_number = transaction_sequence_number
        .Recordset!transaction_reference_number = transaction_reference_number
        .Recordset!account_id = Val(form_main.label_account_id.Caption)
        .Recordset!item_id = item_id
        .Recordset!item_srp = item_srp
        .Recordset!item_rsp = item_rsp
        .Recordset!item_total_quantity_old = total_quantity_old
        .Recordset!item_total_quantity_new = total_quantity_new
        .Recordset.Update
        'Done adding the transaction to the items_transactions database.
        
        'Get transaction id of the current transaction.
        .RecordSource = "select * from items_transactions where transaction_sequence_number = " & transaction_sequence_number & ""
        .Refresh
        .Recordset.MoveFirst
        transaction_id = .Recordset!transaction_id
        'Done getting transaction_id.
        
        'Start adding values to the items_master_history table.
        .RecordSource = "select * from items_master_history"
        .Refresh
        
        .Recordset.AddNew
        .Recordset!history_date_created = Now
        .Recordset!transaction_id = transaction_id
        .Recordset!transaction_sequence_number = transaction_sequence_number
        .Recordset!item_id = item_id
        .Recordset!item_rsp = item_rsp
        .Recordset!item_srp = item_srp
        .Recordset!item_total_quantity = total_quantity_old
        .Recordset.Update
        'Done adding to items_master_history
        
        'Start computing new items_master values
        .RecordSource = "select * from items_master where item_id = " & item_id & ""
        .Refresh
        .Recordset!item_total_quantity = total_quantity_new
        .Recordset!item_date_time_updated = Now
        .Recordset.Update
        .Refresh
        'Done computing and updating items_master record.
        
        form_main.adodc_stocks.Refresh
        form_main.adodc_transactions.Refresh
        
        MsgBox "Stock release submitted.", vbInformation, "System Message"
        text_release_stock_quantity.Text = ""
        text_release_stock_reference_number.Text = ""
        text_item_name.Text = ""
        text_item_name.SetFocus
    End With
End Sub

Private Sub text_item_name_Change()
    With adodc_release_stocks
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
        .RecordSource = "select * from items_master where item_name like '%" & text_item_name.Text & "%' order by item_date_time_updated desc"
        .Refresh
        
    End With
End Sub

Private Sub text_release_stock_reference_number_LostFocus()
    If text_release_stock_reference_number.Text = "" Then
        text_release_stock_reference_number.Text = "00000"
    End If
End Sub
