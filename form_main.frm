VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form form_main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Invento | Simple Inventory"
   ClientHeight    =   9795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13890
   Icon            =   "form_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   13890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame_stocks 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stocks"
      Height          =   8235
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   12495
      Begin MSAdodcLib.Adodc adodc_stocks 
         Height          =   375
         Left            =   5280
         Top             =   7560
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "select * from items_master order by item_date_time_updated desc"
         Caption         =   "adodc_stocks"
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Controls"
         Height          =   3015
         Left            =   8040
         TabIndex        =   15
         Top             =   5040
         Width           =   4335
         Begin VB.CommandButton Command6 
            Height          =   975
            Left            =   2880
            TabIndex        =   41
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Height          =   975
            Left            =   240
            TabIndex        =   40
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            Height          =   975
            Left            =   1560
            TabIndex        =   39
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton command_register_stocks 
            Caption         =   "Register Stocks"
            Height          =   975
            Left            =   2880
            Picture         =   "form_main.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton command_release_stocks 
            Caption         =   "Release Stocks"
            Height          =   975
            Left            =   1560
            Picture         =   "form_main.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton command_add_stocks 
            Caption         =   "Add Stocks"
            Height          =   975
            Left            =   240
            Picture         =   "form_main.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Preview"
         Enabled         =   0   'False
         Height          =   4815
         Left            =   8040
         TabIndex        =   14
         Top             =   120
         Width           =   4335
         Begin VB.TextBox Text10 
            DataField       =   "account_name"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   35
            Top             =   4320
            Width           =   2655
         End
         Begin VB.TextBox Text9 
            DataField       =   "item_total_quantity"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   34
            Top             =   3864
            Width           =   2655
         End
         Begin VB.TextBox Text8 
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
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Top             =   3411
            Width           =   2655
         End
         Begin VB.TextBox Text7 
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
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   32
            Top             =   2958
            Width           =   2655
         End
         Begin VB.TextBox Text6 
            DataField       =   "item_category"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Top             =   2505
            Width           =   2655
         End
         Begin VB.TextBox Text5 
            DataField       =   "item_date_time_updated"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   30
            Top             =   2052
            Width           =   2655
         End
         Begin VB.TextBox Text4 
            DataField       =   "item_date_time_created"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   29
            Top             =   1599
            Width           =   2655
         End
         Begin VB.TextBox Text3 
            DataField       =   "item_description"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   28
            Top             =   1146
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            DataField       =   "item_name"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Top             =   693
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            DataField       =   "item_id"
            DataSource      =   "adodc_stocks"
            Height          =   285
            Left            =   1440
            TabIndex        =   26
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Updater:"
            Height          =   220
            Left            =   120
            TabIndex        =   25
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Quantity:"
            Height          =   220
            Left            =   120
            TabIndex        =   24
            Top             =   3864
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "SRP:"
            Height          =   220
            Left            =   120
            TabIndex        =   23
            Top             =   3411
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "RSP:"
            Height          =   220
            Left            =   120
            TabIndex        =   22
            Top             =   2958
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Category:"
            Height          =   220
            Left            =   120
            TabIndex        =   21
            Top             =   2505
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Updated:"
            Height          =   220
            Left            =   120
            TabIndex        =   20
            Top             =   2052
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Created:"
            Height          =   220
            Left            =   120
            TabIndex        =   19
            Top             =   1599
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Description:"
            Height          =   220
            Left            =   120
            TabIndex        =   18
            Top             =   1146
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name:"
            Height          =   220
            Left            =   120
            TabIndex        =   17
            Top             =   693
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "ID:"
            Height          =   220
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid datagrid_stocks 
         Bindings        =   "form_main.frx":1108
         Height          =   7815
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   13785
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "item_id"
            Caption         =   "ID"
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
            DataField       =   "item_name"
            Caption         =   "NAME"
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
         BeginProperty Column02 
            DataField       =   "item_total_quantity"
            Caption         =   "STOCKS"
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
            DataField       =   "item_rsp"
            Caption         =   "RSP"
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "item_category"
            Caption         =   "CATEGORY"
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
               Type            =   1
               Format          =   """PHP""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
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
            Caption         =   "STOCKS"
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
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1679.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2039.811
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
   End
   Begin VB.Frame frame_settings 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Settings"
      Height          =   8655
      Left            =   1800
      TabIndex        =   12
      Top             =   960
      Width           =   12015
      Begin MSAdodcLib.Adodc adodc_settings 
         Height          =   375
         Left            =   8400
         Top             =   8040
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         RecordSource    =   "select * from settings where setting_name like '%enterprise%'"
         Caption         =   "adodc_settings"
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
      Begin VB.Frame frame_design_customizer 
         Caption         =   "Design Customizer"
         Height          =   4455
         Left            =   5640
         TabIndex        =   65
         Top             =   4080
         Width           =   6255
         Begin VB.Frame Frame6 
            Caption         =   "Color Theme"
            Height          =   735
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   6015
            Begin VB.CommandButton command_save_color_settings 
               Caption         =   "Save"
               Height          =   375
               Left            =   4560
               TabIndex        =   69
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton command_header_text_color 
               Caption         =   "Header Text Color"
               Height          =   375
               Left            =   2400
               TabIndex        =   68
               Top             =   240
               Width           =   2055
            End
            Begin MSComDlg.CommonDialog common_dialog_color_picker 
               Left            =   5400
               Top             =   120
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton command_header_background_color 
               Caption         =   "Header Background Color"
               Height          =   375
               Left            =   120
               TabIndex        =   67
               Top             =   240
               Width           =   2175
            End
         End
      End
      Begin VB.Frame frame_settings_guide 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reminders"
         Height          =   3975
         Left            =   5640
         TabIndex        =   61
         Top             =   120
         Width           =   6255
         Begin VB.Label label_settings_guide 
            Height          =   3615
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.CommandButton command_setting_update 
         Caption         =   "Update"
         Height          =   375
         Left            =   4320
         TabIndex        =   58
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox text_setting_value 
         DataField       =   "setting_value"
         DataSource      =   "adodc_settings"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   3720
         Width           =   4095
      End
      Begin MSDataGridLib.DataGrid datagrid_settings 
         Bindings        =   "form_main.frx":1123
         Height          =   3375
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5953
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "setting_name"
            Caption         =   "setting_name"
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
            DataField       =   "setting_value"
            Caption         =   "setting_value"
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
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2835.213
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton command_close 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "form_main.frx":1140
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame frame_reports 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transactions"
      Height          =   8415
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   12255
      Begin VB.CommandButton command_search 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10440
         Picture         =   "form_main.frx":1582
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc adodc_transactions 
         Height          =   375
         Left            =   8640
         Top             =   7440
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
         RecordSource    =   $"form_main.frx":19C4
         Caption         =   "adodc_transactions"
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
      Begin VB.CommandButton command_print 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   10440
         Picture         =   "form_main.frx":1A4C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filter"
         Height          =   2055
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   10095
         Begin MSAdodcLib.Adodc adodc_filter 
            Height          =   375
            Left            =   3360
            Top             =   1440
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "adodc_filter"
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
         Begin MSAdodcLib.Adodc adodc_items_master 
            Height          =   375
            Left            =   6600
            Top             =   1440
            Visible         =   0   'False
            Width           =   3135
            _ExtentX        =   5530
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
            RecordSource    =   "select * from items_master order by item_date_time_updated desc"
            Caption         =   "adodc_items_master"
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
         Begin MSDataGridLib.DataGrid datagrid_transactions_filter 
            Bindings        =   "form_main.frx":1E8E
            Height          =   1215
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   2143
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "item_name"
               Caption         =   "Item Name"
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
               DataField       =   "item_date_time_updated"
               Caption         =   "Date Updated"
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
            BeginProperty Column02 
               DataField       =   "item_category"
               Caption         =   "Category"
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
               DataField       =   "item_rsp"
               Caption         =   "RSP"
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
            BeginProperty Column04 
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
            BeginProperty Column05 
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
            BeginProperty Column06 
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
            BeginProperty Column07 
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
                  ColumnWidth     =   2700.284
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2009.764
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1604.976
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin VB.OptionButton option_by_category 
            BackColor       =   &H00FFFFFF&
            Caption         =   "By Category"
            Height          =   255
            Left            =   1200
            TabIndex        =   55
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton option_by_item 
            BackColor       =   &H00FFFFFF&
            Caption         =   "By Item"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin MSDataListLib.DataCombo datacombo_filter 
            Bindings        =   "form_main.frx":1EAF
            Height          =   420
            Left            =   2640
            TabIndex        =   5
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   741
            _Version        =   393216
            ListField       =   "item_name"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date Range"
         Height          =   855
         Left            =   5640
         TabIndex        =   48
         Top             =   120
         Width           =   6375
         Begin MSComCtl2.DTPicker date_picker_to 
            Height          =   375
            Left            =   3840
            TabIndex        =   52
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            Format          =   104529921
            CurrentDate     =   42820
         End
         Begin MSComCtl2.DTPicker date_picker_from 
            Height          =   375
            Left            =   960
            TabIndex        =   49
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            Format          =   104529921
            CurrentDate     =   42820
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "To:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   51
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "From:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox checkbox_pulledout_stocks 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pulled-out Stocks"
         Height          =   375
         Left            =   2280
         TabIndex        =   47
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox checkbox_added_stocks 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Added Stocks"
         Height          =   375
         Left            =   4080
         TabIndex        =   46
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox checkbox_sold_stocks 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sold Stocks"
         Height          =   375
         Left            =   720
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid datagrid_transactions 
         Bindings        =   "form_main.frx":1ECA
         Height          =   5055
         Left            =   240
         TabIndex        =   44
         Top             =   3120
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   8916
         _Version        =   393216
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "transaction_type_name"
            Caption         =   "Type"
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
            DataField       =   "transaction_quantity"
            Caption         =   "Qty"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "item_category"
            Caption         =   "Category"
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
         BeginProperty Column05 
            DataField       =   "transaction_srp_amount"
            Caption         =   "Total SRP"
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
         BeginProperty Column06 
            DataField       =   "transaction_sequence_number"
            Caption         =   "Seq. No."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """PHP""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "transaction_reference_number"
            Caption         =   "Ref. No."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """PHP""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "item_total_quantity_new"
            Caption         =   "Avl. Stocks"
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
            DataField       =   "transaction_date_time"
            Caption         =   "Trans. Date"
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
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2160
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton command_logout 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "form_main.frx":1EEB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame frame_header_account 
      BackColor       =   &H00004080&
      Caption         =   "Account"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10440
      TabIndex        =   42
      Top             =   240
      Width           =   3375
      Begin VB.Label label_account_id 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   3000
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label label_username 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton command_show_settings_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "form_main.frx":232D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton command_show_reports_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Trans"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "form_main.frx":276F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton command_show_stocks_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "form_main.frx":2BB1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label label_version 
      BackStyle       =   0  'Transparent
      Caption         =   "label_version"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1920
      TabIndex        =   64
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label label_enterprise_description 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enterprise description."
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3000
      TabIndex        =   60
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label label_enterprise_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enterprise Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   59
      Top             =   240
      Width           =   6135
   End
   Begin VB.Shape shape_dot 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   360
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
      TabIndex        =   9
      Top             =   360
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   9015
      Left            =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Shape shape_header_background 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   120
      Width           =   14175
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enterprise_name, enterprise_description, enterprise_email, enterprise_address, enterprise_contact_number, enterprise_logo As String
Dim settings_header_background_color, settings_header_text_color As Long
Dim settings_guide As String

Private Sub command_add_stocks_Click()
    'Unload form_add_stocks
    form_add_stocks.Show
End Sub

Private Sub command_close_Click()
    Unload form_login
    Unload form_add_stocks
    Unload form_register_stocks
    Unload form_release_stocks
    Unload form_main
    Exit Sub
End Sub

Private Sub command_header_background_color_Click()
    ' Set Cancel to True
    common_dialog_color_picker.CancelError = True
    On Error GoTo ErrHandler
    'Set the Flags property
    common_dialog_color_picker.Flags = cdlCCRGBInit
    ' Display the Color Dialog box
    common_dialog_color_picker.ShowColor
    ' Set the form's background color to selected color
    settings_header_background_color = common_dialog_color_picker.Color
    Call apply_color_settings
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
End Sub

Private Sub command_header_text_color_Click()
    ' Set Cancel to True
    common_dialog_color_picker.CancelError = True
    On Error GoTo ErrHandler
    'Set the Flags property
    common_dialog_color_picker.Flags = cdlCCRGBInit
    ' Display the Color Dialog box
    common_dialog_color_picker.ShowColor
    ' Set the form's background color to selected color
    settings_header_text_color = common_dialog_color_picker.Color
    Call apply_color_settings
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
End Sub

Private Sub command_logout_Click()
    label_username.Caption = "None"
    command_show_stocks_frame.Enabled = False
    command_show_settings_frame.Enabled = False
    command_show_reports_frame.Enabled = False
    command_logout.Enabled = False
    user_logged_in = False
    Call form_main_default
    Call load_settings
    Call apply_color_settings
    form_login.Show
End Sub

Private Sub command_print_Click()
    Dim report_criteria_string As String
    
    report_criteria_string = "This report contains: "
    If checkbox_sold_stocks.Value = 1 Then
        report_criteria_string = report_criteria_string & "(Sold stocks)"
    End If
    If checkbox_added_stocks.Value = 1 Then
        report_criteria_string = report_criteria_string & "(Added stocks)"
    End If
    If checkbox_pulledout_stocks.Value = 1 Then
        report_criteria_string = report_criteria_string & "(Pulled Out stocks)"
    End If
    
    If checkbox_pulledout_stocks.Value = 0 And checkbox_added_stocks.Value = 0 And checkbox_sold_stocks.Value = 0 Then
        report_criteria_string = report_criteria_string & " All transactions from "
    Else
        report_criteria_string = report_criteria_string & " transactions from "
    End If
    
    Call load_settings
    
    Set datareport_items_transactions.DataSource = adodc_transactions
    report_criteria_string = report_criteria_string & " " & date_picker_from & " to " & date_picker_to & "."
    datareport_items_transactions.Sections("Section4").Controls("label_date_range").Caption = report_criteria_string
    
    Set datareport_items_transactions.Sections("Section4").Controls("Image1").Picture = LoadPicture(enterprise_logo)
    datareport_items_transactions.Sections("Section4").Controls("label_enterprise_name").Caption = enterprise_name
    datareport_items_transactions.Sections("Section4").Controls("label_description").Caption = enterprise_description
    datareport_items_transactions.Sections("Section4").Controls("label_email_contact_number_address").Caption = enterprise_email & " | " & enterprise_contact_number & " | " & enterprise_address
    
    datareport_items_transactions.Show
End Sub

Private Sub command_register_stocks_Click()
    'Unload form_register_stocks
    form_register_stocks.Show
End Sub

Private Sub command_release_stocks_Click()
    'Unload form_release_stocks
    form_release_stocks.Show
End Sub

Private Sub command_save_color_settings_Click()
    If MsgBox("Save this color settings?", vbYesNo, "System Message") = vbYes Then
        With adodc_settings
            'save the header background settings to the database
            .RecordSource = "select * from settings where setting_name = 'settings_header_background_color'"
            .Refresh
            .Recordset!setting_value = settings_header_background_color
            .Recordset.Update
            
            'save the header text color settings
            .RecordSource = "select * from settings where setting_name = 'settings_header_text_color'"
            .Refresh
            .Recordset!setting_value = settings_header_text_color
            .Recordset.Update
            
            'set the adodc_settings recordset back to the default
            .RecordSource = "select * from settings where setting_name like '%enterprise%'"
            .Refresh
            
            Call load_settings
            MsgBox "Color settings saved!", vbInformation, "System Message"
        End With
    End If
End Sub

Private Sub command_search_Click()
    'Build the query.
    Dim built_query As String
    Dim check_box_checked As Boolean
    
    check_box_checked = False
    built_query = " select * from query_items_transactions where "
    
    'Start building query from check boxes.
    'If the sold stocks checkbox is  checked.
    If checkbox_sold_stocks.Value = 1 Then
        built_query = built_query & " (transaction_type = 1 "
        check_box_checked = True
    End If
    'If the pulledout stocks checkbox is  checked.
    If checkbox_pulledout_stocks.Value = 1 Then
        If checkbox_sold_stocks.Value = 1 Then
            built_query = built_query & " or transaction_type = 2 "
        Else
            built_query = built_query & " (transaction_type = 2 "
            check_box_checked = True
        End If
    End If
    'If the added stocks checkbox is checked.
    If checkbox_added_stocks.Value = 1 Then
        If checkbox_sold_stocks.Value = 1 Or checkbox_pulledout_stocks.Value = 1 Then
            built_query = built_query & " or transaction_type = 0 "
        Else
            built_query = built_query & " (transaction_type = 0 "
            check_box_checked = True
        End If
    End If
    'Done building query from check boxes.
    
    'Start building query from date ranges.
    If check_box_checked = True Then
        built_query = built_query & " ) "
        built_query = built_query & " and ( transaction_date_time between #" & date_picker_from.Value & "# and #" & date_picker_to.Value & " 23:59:59# ) "
    Else
        built_query = built_query & " ( transaction_date_time between #" & date_picker_from.Value & "# and #" & date_picker_to.Value & " 23:59:59# ) "
    End If
    'Done building query from date ranges.
    
    'Build query from filter
    If datacombo_filter.Text <> "" Then
        With adodc_items_master
            If option_by_item.Value = True Then
                datacombo_filter.Text = .Recordset("item_name")
                built_query = built_query & " and (item_name like '%" & datacombo_filter.Text & "') "
            ElseIf option_by_category = True Then
                datacombo_filter.Text = .Recordset("item_category")
                built_query = built_query & " and (item_category like '%" & datacombo_filter.Text & "') "
            End If
        End With
    End If
    'Done building query from filter.
    'Finished building the query.
    
    With adodc_transactions
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=invento_db.mdb;Persist Security Info=False"
        .RecordSource = built_query & " order by transaction_date_time desc "
        'MsgBox built_query
        .Refresh
    End With
End Sub

Private Sub command_setting_update_Click()
    With adodc_settings
        .Recordset!setting_value = text_setting_value.Text
        .Recordset.Update
        .Refresh
        Call load_settings
        
        MsgBox "Settings updated successful!", vbInformation, "System Message"
    End With
End Sub

Private Sub command_show_reports_frame_Click()
    'check if the account is an admin account
    If label_account_id = 1 Then
        Call form_main_default
        frame_reports.Visible = True
        datacombo_filter.SetFocus
    Else
        MsgBox "Access denied!", vbExclamation, "System Message"
    End If
End Sub

Private Sub command_show_settings_frame_Click()
    'check if the account is an admin account
    If label_account_id = 1 Then
        Call form_main_default
        frame_settings.Visible = True
    Else
        MsgBox "Access denied!", vbExclamation, "System Message"
    End If
End Sub

Private Sub command_show_stocks_frame_Click()
    Call form_main_default
    frame_stocks.Visible = True
End Sub





Private Sub datacombo_filter_Change()
    If option_by_item.Value = True Then
        With adodc_items_master
            .RecordSource = "select * from items_master where item_name like '%" & datacombo_filter.Text & "'"
            .Refresh
        End With
    ElseIf option_by_category.Value = True Then
        With adodc_items_master
            .RecordSource = "select * from items_master where item_category like '%" & datacombo_filter.Text & "'"
            .Refresh
        End With
    End If
End Sub




Private Sub datagrid_settings_SelChange(Cancel As Integer)
    text_setting_value.Text = ""
    Set text_setting_value.DataSource = adodc_settings
    text_setting_value.DataField = "setting_value"
End Sub

Private Sub Form_Load()
    form_main.Caption = form_main.Caption & " beta v" & App.Major & "." & App.Minor & "." & App.Revision
    label_version.Caption = "beta v" & App.Major & "." & App.Minor & "." & App.Revision
    Call form_main_default
    Call load_settings
    
    'Uncomment these lines if you want to add time limits to app usage.
    'If Date > "4/10/2017" Then
        'MsgBox "Error: Usage exhausted.", vbCritical, "System Message"
        'form_login.Hide
    'End If
End Sub

Sub load_settings()
    With adodc_settings
        'Get enterprise_name.
        .RecordSource = "select setting_value from settings where setting_name = 'enterprise_name'"
        .Refresh
        enterprise_name = .Recordset("setting_value")
        
        'Get enterprise_description.
        .RecordSource = "select setting_value from settings where setting_name = 'enterprise_description'"
        .Refresh
        enterprise_description = .Recordset("setting_value")
        
        'Get enterprise_email.
        .RecordSource = "select setting_value from settings where setting_name = 'enterprise_email'"
        .Refresh
        enterprise_email = .Recordset("setting_value")
        
        'Get enterprise_address.
        .RecordSource = "select setting_value from settings where setting_name = 'enterprise_address'"
        .Refresh
        enterprise_address = .Recordset("setting_value")
        
        'Get enterprise_contact_number.
        .RecordSource = "select setting_value from settings where setting_name = 'enterprise_contact_number'"
        .Refresh
        enterprise_contact_number = .Recordset("setting_value")
        
        'Get enterprise_logo.
        .RecordSource = "select setting_value from settings where setting_name = 'enterprise_logo'"
        .Refresh
        enterprise_logo = .Recordset("setting_value")
        
        'Get settings_guide.
        .RecordSource = "select setting_value from settings where setting_name = 'settings_guide'"
        .Refresh
        settings_guide = .Recordset("setting_value")
        
        'Get settings_header_background_color.
        .RecordSource = "select setting_value from settings where setting_name = 'settings_header_background_color'"
        .Refresh
        settings_header_background_color = Val(.Recordset("setting_value"))
        
        'Get settings_header_text_color.
        .RecordSource = "select * from settings where setting_name='settings_header_text_color'"
        .Refresh
        settings_header_text_color = Val(.Recordset("setting_value"))
        
        'Load settings that can be edited by the users during runtime.
        .RecordSource = "select * from settings where setting_name like '%enterprise%'"
        .Refresh
    End With
    
    label_enterprise_name.Caption = enterprise_name
    label_enterprise_description = enterprise_description
    label_settings_guide.Caption = settings_guide
    
    Call apply_color_settings
End Sub

Sub apply_color_settings()
    'Load and apply color settings.
    label_header_app_name.ForeColor = settings_header_text_color
    shape_dot.FillColor = settings_header_text_color
    shape_header_background.FillColor = settings_header_background_color
    label_enterprise_name.ForeColor = settings_header_text_color
    label_enterprise_description.ForeColor = settings_header_text_color
    frame_header_account.BackColor = settings_header_background_color
    frame_header_account.ForeColor = settings_header_text_color
    label_username.ForeColor = settings_header_text_color
    label_version.ForeColor = settings_header_text_color
    
    With form_add_stocks
        .shape_dot.FillColor = settings_header_text_color
        .shape_header_background.FillColor = settings_header_background_color
        .label_header_app_name.ForeColor = settings_header_text_color
    End With
    
    With form_login
        .shape_dot.FillColor = settings_header_text_color
        .shape_header_background.FillColor = settings_header_background_color
        .label_header_app_name.ForeColor = settings_header_text_color
    End With
    
    With form_register_stocks
        .shape_dot.FillColor = settings_header_text_color
        .shape_header_background.FillColor = settings_header_background_color
        .label_header_app_name.ForeColor = settings_header_text_color
    End With
    
    With form_release_stocks
        .shape_dot.FillColor = settings_header_text_color
        .shape_header_background.FillColor = settings_header_background_color
        .label_header_app_name.ForeColor = settings_header_text_color
    End With
End Sub

Sub app_terminate()
    Unload form_login
    Unload form_add_stocks
    Unload form_register_stocks
    Unload form_release_stocks
    Unload form_main
    Unload datareport_items_transactions
End Sub

Sub form_main_default()
    frame_stocks.Visible = False
    frame_settings.Visible = False
    frame_reports.Visible = False
    form_add_stocks.Hide
    form_register_stocks.Hide
    form_release_stocks.Hide
    
    date_picker_from.Value = Date
    date_picker_to.Value = Date
End Sub

Private Sub Form_Terminate()
    Call app_terminate
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call app_terminate
End Sub

Private Sub option_by_category_Click()
    With adodc_filter
        .RecordSource = "select * from query_categories"
        .Refresh
        datacombo_filter.ListField = "item_category"
    End With
End Sub

Private Sub option_by_item_Click()
    With adodc_filter
        .RecordSource = "select * from items_master"
        .Refresh
        datacombo_filter.ListField = "item_name"
    End With
End Sub

Private Sub text_setting_value_GotFocus()
    With adodc_settings
        On Error Resume Next
        Set text_setting_value.DataSource = Null
        text_setting_value.DataField = ""
        text_setting_value.Text = .Recordset("setting_value")
    End With
End Sub
