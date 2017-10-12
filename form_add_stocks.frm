VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form form_add_stocks 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Stocks Form"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   Icon            =   "form_add_stocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text_reference_number 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2880
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc adodc_datacombo_item_names 
      Height          =   375
      Left            =   1320
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from query_items_master_item_name"
      Caption         =   "adodc_datacombo_item_names"
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
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      DataField       =   "item_category"
      DataSource      =   "adodc_add_stocks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2880
      TabIndex        =   14
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      DataField       =   "item_srp"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """PHP""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adodc_add_stocks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5160
      TabIndex        =   12
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      DataField       =   "item_total_quantity"
      DataSource      =   "adodc_add_stocks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2880
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton command_add_stock_submit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      Picture         =   "form_add_stocks.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox text_stocks_to_add 
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo datacombo_item_name 
      Bindings        =   "form_add_stocks.frx":0884
      Height          =   480
      Left            =   2880
      TabIndex        =   0
      Top             =   4200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "item_name"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      DataField       =   "item_rsp"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """PHP""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adodc_add_stocks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2880
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      DataField       =   "item_name"
      DataSource      =   "adodc_add_stocks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   3855
   End
   Begin MSAdodcLib.Adodc adodc_add_stocks 
      Height          =   375
      Left            =   4200
      Top             =   6960
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from items_master"
      Caption         =   "adodc_add_stocks"
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
   Begin VB.Label label_reference_number 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reference No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   4920
      Width           =   2175
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SRP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stocks On Hand:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stocks to Add:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   9
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RSP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Filter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
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
      TabIndex        =   16
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape shape_header_background 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   -360
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "form_add_stocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub command_add_stock_submit_Click()
    'Check if the entered value on the stocks to add box is valid.
    Dim stocks_to_add As Integer
    stocks_to_add = Val(text_stocks_to_add.Text)
    If stocks_to_add < 1 Then
        'If the input was not valid. (ex. The input was not a number.)
        MsgBox "Please check your inputs properly.", vbExclamation, "System Message"
        text_stocks_to_add.SetFocus
    Else
        'Else if the input was valid. Continue with the submit.
        'Ask the user if he/she wants to continue with the transaction.
        If MsgBox("Do you want to proceed with this transaction?", vbYesNo, "System Message") = vbYes Then
            'If the user wants to continue with the transaction.
            
            'Declare variables.
            Dim item_id As Integer
            Dim transaction_id As Integer
            Dim quantity_to_add As Integer
            Dim current_quantity As Integer
            Dim sequence_number As Single
            Dim item_srp As Currency
            Dim item_rsp As Currency
            
            With adodc_add_stocks
                item_id = .Recordset("item_id")
                quantity_to_add = stocks_to_add
                current_quantity = .Recordset("item_total_quantity")
                item_rsp = .Recordset("item_rsp")
                item_srp = .Recordset("item_srp")
                
                'Check if the current item chosen in the datacombo exists in the items_master table.
                .RecordSource = "select * from items_master where item_name = '" & datacombo_item_name.Text & "'"
                .Refresh
                If .Recordset.RecordCount < 1 Then
                    'If the record does not exist.
                    'Exit the sub and alert the user.
                    MsgBox "This item name does not exist!", vbExclamation, "System Message"
                    Exit Sub
                End If
                
                'Get a new sequence number.
                .RecordSource = "select * from sequence_number order by sequence_number_id desc"
                .Refresh
                .Recordset.MoveFirst
                sequence_number = .Recordset!sequence_number + 1
                
                'Save new sequence number.
                .Recordset.AddNew
                .Recordset!sequence_number = sequence_number
                .Recordset.Update
                'Done saving new sequence number.
                
                'Start adding the transaction into the items_transactions table.
                .RecordSource = "select * from items_transactions"
                .Refresh
                
                .Recordset.AddNew
                .Recordset!transaction_type = 0 '0 means add stock
                .Recordset!transaction_type_reason = 0 '0 means none or no other reason
                .Recordset!transaction_date_time = Now
                .Recordset!transaction_reference_number = Val(text_reference_number.Text)
                .Recordset!transaction_sequence_number = sequence_number
                .Recordset!item_id = item_id
                .Recordset!item_rsp = item_rsp
                .Recordset!item_srp = item_srp
                .Recordset!item_total_quantity_old = current_quantity
                .Recordset!item_total_quantity_new = current_quantity + quantity_to_add
                .Recordset!transaction_quantity = quantity_to_add
                .Recordset!account_id = Val(form_main.label_account_id.Caption)
                .Recordset.Update
                'Done adding transaction.
                
                'Start adding the current items_master values of the current item into the items_master history table.
                'Get transaction id of the current transaction.
                .RecordSource = "select * from items_transactions where transaction_sequence_number = " & sequence_number & ""
                .Refresh
                .Recordset.MoveFirst
                transaction_id = .Recordset!transaction_id
                'Done getting transaction_id.
                
                .RecordSource = "select * from items_master_history"
                .Refresh
                
                .Recordset.AddNew
                .Recordset!history_date_created = Now
                .Recordset!transaction_id = transaction_id
                .Recordset!transaction_sequence_number = sequence_number
                .Recordset!item_id = item_id
                .Recordset!item_rsp = item_rsp
                .Recordset!item_srp = item_srp
                .Recordset!item_total_quantity = current_quantity
                .Recordset.Update
                'Done adding to items_master_history
                
                'Start computing new items_master values
                .RecordSource = "select * from items_master where item_id = " & item_id & ""
                .Refresh
                .Recordset!item_total_quantity = current_quantity + quantity_to_add
                .Recordset!item_date_time_updated = Now
                .Recordset.Update
                .Refresh
                'Done computing and updating items_master record.
                
                .RecordSource = "select * from items_master order by item_date_time_updated desc"
                .Refresh
                
                'Alert the user that the transaction was finished.
                MsgBox "Transaction submitted.", vbInformation, "System Message"
                
                'Return the focus into the filter box.
                text_stocks_to_add.Text = ""
                datacombo_item_name.SetFocus
                
                form_main.adodc_stocks.Refresh
                form_main.adodc_transactions.Refresh
            End With
            
        End If
    End If
End Sub

Private Sub datacombo_item_name_Change()
    adodc_add_stocks.RecordSource = "SELECT * FROM items_master WHERE item_name like '%" & datacombo_item_name.Text & "'"
    adodc_add_stocks.Refresh
End Sub

Private Sub text_stocks_to_add_LostFocus()
    Dim stocks_to_add As Integer
    stocks_to_add = Val(text_stocks_to_add.Text)
    If stocks_to_add < 1 Then
        'Check if the input was a number.
        'If the input was not a number.
        text_stocks_to_add.BackColor = &H8080FF
        MsgBox "Wrong Input!", vbCritical, "System Message"
    End If
End Sub
