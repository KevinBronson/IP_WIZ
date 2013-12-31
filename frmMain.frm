VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{53337153-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "ipinfo50.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Wizard"
   ClientHeight    =   7215
   ClientLeft      =   150
   ClientTop       =   510
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   90
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   975
      ScaleWidth      =   3045
      TabIndex        =   20
      Top             =   6120
      Width           =   3075
   End
   Begin MSComctlLib.ImageList imgGeneralTest 
      Left            =   9990
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2524
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2976
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":321A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":366C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4362
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5058
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":773A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8430
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8882
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":955C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A236
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF10
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BBEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C8C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D59E
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E278
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EF52
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FC2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10906
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":122BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14948
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1563E
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   270
      TabIndex        =   12
      Top             =   540
      Width           =   10185
      Begin VB.Frame Frame6 
         Height          =   825
         Left            =   180
         TabIndex        =   22
         Top             =   180
         Width           =   9825
         Begin VB.CommandButton cmdGenTestError 
            Caption         =   "Show Error Log"
            Height          =   375
            Left            =   8190
            TabIndex        =   30
            Top             =   270
            Width           =   1455
         End
         Begin VB.CommandButton cmdGenTestShowLog 
            Caption         =   "Show Details"
            Height          =   375
            Left            =   6660
            TabIndex        =   29
            Top             =   270
            Width           =   1455
         End
         Begin VB.CheckBox chkGenTest_InternetConn 
            Caption         =   "Check Internet Connection"
            Height          =   285
            Left            =   4320
            TabIndex        =   28
            Top             =   270
            Width           =   2265
         End
         Begin VB.CommandButton cmdGeneralTest 
            Caption         =   "Start"
            Height          =   375
            Left            =   180
            TabIndex        =   23
            Top             =   270
            Width           =   1185
         End
         Begin IP_WIZ.Light lightGenTest_InternetConn 
            Height          =   105
            Left            =   1530
            TabIndex        =   24
            Top             =   540
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
         End
         Begin IP_WIZ.Light lightGenTest_InProgress 
            Height          =   105
            Left            =   1530
            TabIndex        =   25
            Top             =   270
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
         End
         Begin VB.Label lblGenTest_InternetConn 
            Caption         =   "Internet Connection"
            Height          =   285
            Left            =   1890
            TabIndex        =   27
            Top             =   450
            Width           =   2895
         End
         Begin VB.Label lblGenTest_InProgress 
            Caption         =   "Test in Progress"
            Height          =   195
            Left            =   1890
            TabIndex        =   26
            Top             =   180
            Width           =   2715
         End
      End
      Begin IP_WIZ.IP_Config_Tree IP_Config_Tree1 
         Height          =   2355
         Left            =   180
         TabIndex        =   19
         Top             =   2790
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   4154
      End
      Begin IP_WIZ.VisualPingControl vpcGenTest 
         Height          =   1635
         Left            =   90
         TabIndex        =   14
         Top             =   1080
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   2884
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   270
      TabIndex        =   1
      Top             =   540
      Width           =   10185
      Begin IP_WIZ.IP_Config_Tree vpcIP_INFO_IP_Config_Tree 
         Height          =   4335
         Left            =   180
         TabIndex        =   18
         Top             =   810
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   7646
      End
      Begin VB.CommandButton cmdIP_Info_Collapse 
         Caption         =   "Collapse All"
         Height          =   375
         Left            =   8010
         TabIndex        =   17
         Top             =   270
         Width           =   1995
      End
      Begin VB.CommandButton cmdIP_Info_Expand 
         Caption         =   "Expand All"
         Height          =   375
         Left            =   5940
         TabIndex        =   16
         Top             =   270
         Width           =   1995
      End
      Begin VB.CommandButton cmdIP_INFO_Refresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3870
         TabIndex        =   15
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Just double click an IP Address to Ping it."
         Height          =   195
         Left            =   270
         TabIndex        =   31
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   270
      TabIndex        =   2
      Top             =   540
      Width           =   10185
      Begin VB.ComboBox txtTraceRouteRemoteHost 
         Height          =   330
         ItemData        =   "frmMain.frx":15EE2
         Left            =   2430
         List            =   "frmMain.frx":15EE4
         TabIndex        =   40
         Top             =   270
         Width           =   4605
      End
      Begin VB.CheckBox chkTraceRouteResolve 
         Caption         =   "Resolve Host Names"
         Height          =   285
         Left            =   8280
         TabIndex        =   34
         Top             =   180
         Width           =   1815
      End
      Begin VB.CommandButton cmdTraceRouteStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   7200
         TabIndex        =   33
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox chkTraceRouteCompare 
         Caption         =   "Compare Times"
         Height          =   285
         Left            =   8280
         TabIndex        =   35
         Top             =   450
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         Height          =   4515
         Left            =   90
         TabIndex        =   36
         Top             =   720
         Width           =   10005
         Begin VB.CommandButton cmdTraceRouteCollapseAll 
            Caption         =   "Collapse All"
            Height          =   375
            Left            =   7920
            TabIndex        =   39
            Top             =   270
            Width           =   1905
         End
         Begin VB.CommandButton cmdTraceRouteExpandall 
            Caption         =   "Expand All"
            Height          =   375
            Left            =   5940
            TabIndex        =   38
            Top             =   270
            Width           =   1905
         End
         Begin IP_WIZ.TraceRouteTree TraceRouteTree1 
            Height          =   3615
            Left            =   180
            TabIndex        =   37
            Top             =   720
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   6376
         End
         Begin IP_WIZ.Light lightTraceRoute_Test 
            Height          =   105
            Left            =   180
            TabIndex        =   47
            Top             =   330
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
         End
         Begin VB.Label lblfasfdsaf 
            Caption         =   "Test In Progress"
            Height          =   195
            Left            =   540
            TabIndex        =   48
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Label lblTraceRoute 
         Caption         =   "Computer Name or IP Address:"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   300
         Width           =   2265
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   270
      TabIndex        =   3
      Top             =   540
      Width           =   10185
      Begin VB.ComboBox txtLookup_Host 
         Height          =   330
         Left            =   2970
         TabIndex        =   46
         Top             =   270
         Width           =   4875
      End
      Begin IP_WIZ.Light lightLookup_TestInProgress 
         Height          =   105
         Left            =   180
         TabIndex        =   44
         Top             =   690
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   185
      End
      Begin VB.TextBox txtLookup_Results 
         Height          =   4155
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   43
         Top             =   990
         Width           =   9825
      End
      Begin VB.CommandButton cmdLookup_Start 
         Caption         =   "Start"
         Height          =   375
         Left            =   8010
         TabIndex        =   42
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Test in Progress"
         Height          =   285
         Left            =   540
         TabIndex        =   45
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Host Name or IP Address to Resolve:"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   270
         Width           =   3435
      End
      Begin IPINFOLibCtl.IPInfo objLookup_IPInfo 
         Left            =   2430
         Top             =   540
         PendingRequests =   0
         ServiceName     =   ""
         ServicePort     =   0
         ServiceProtocol =   ""
         WinsockLoaded   =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   270
      TabIndex        =   4
      Top             =   540
      Width           =   10185
      Begin VB.TextBox txtPING_Details 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3060
         Visible         =   0   'False
         Width           =   9825
      End
      Begin VB.CommandButton cmdPING_ShowHide 
         Caption         =   "Show Details >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Frame Frame1_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Width           =   9825
         Begin VB.CheckBox chkPING_QuickTest 
            Caption         =   "Quick Test"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7470
            TabIndex        =   11
            Top             =   270
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton cmdPING_Start 
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   8730
            TabIndex        =   7
            Top             =   180
            Width           =   1005
         End
         Begin VB.TextBox txtHost 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2430
            TabIndex        =   6
            Top             =   270
            Width           =   4785
         End
         Begin VB.Label lblHost 
            Caption         =   "Computer name or IP Address:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   8
            Top             =   270
            Width           =   2265
         End
      End
      Begin IP_WIZ.VisualPingControl vpcComputerCheck 
         Height          =   1635
         Left            =   90
         TabIndex        =   13
         Top             =   900
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   2884
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5955
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   10504
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Computer Check"
            Key             =   "PING"
            Object.ToolTipText     =   "Check to see if a computer is present using PING"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Trace to Computer"
            Key             =   "TraceRoute"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "IP Info for this Computer"
            Key             =   "IP_INFO"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lookup Domain Information"
            Key             =   "WHOIS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General Network Test"
            Key             =   "GeneralTest"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFooter 
      Height          =   1185
      Left            =   3240
      TabIndex        =   21
      Top             =   6120
      Width           =   7575
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu Contents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu Spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu LegalStatement 
         Caption         =   "&Legal Statement..."
      End
      Begin VB.Menu VersionInfo 
         Caption         =   "&Version Information..."
      End
      Begin VB.Menu Spacer4 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'Global Objects
Dim WithEvents objGenTestLog As ItemList
Attribute objGenTestLog.VB_VarHelpID = -1
Dim WithEvents objGenTestErrorLog As ItemList
Attribute objGenTestErrorLog.VB_VarHelpID = -1

'Flags
Dim blnGenTest_StopRequested As Boolean


'==================================================================================
'                                   MENU BAR
'==================================================================================

'About_Click
Private Sub About_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    objCompanyDisplay.About
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "About_Click")
End Sub

'Close_Click
Private Sub Close_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    End
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Close_Click")
End Sub

'Contents_Click
Private Sub Contents_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim objShell As New Shell32.Shell
    objShell.Open (App.Path & "\Help\help.chm")
    Set objShell = Nothing
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Contents_Click")
End Sub

'LegalStatement_Click
Private Sub LegalStatement_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    frmLegalStatement.Show
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "LegalStatement_Click")
End Sub

'Options_Click
Private Sub Options_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Options_Click")
End Sub

'Picture1_Click
Private Sub Picture1_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Call About_Click
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Picture1_Click")
End Sub

'VersionInfo_Click
Private Sub VersionInfo_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    MsgBox App.Title & " Version: " & App.Major & "." & App.Minor & "." & App.Revision
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "VersionInfo_Click")
End Sub


'==================================================================================
'                                  Lookup
'==================================================================================

'cmdLookup_Start_Click
Private Sub cmdLookup_Start_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Len(txtLookup_Host) > 0 Then
        txtLookup_Host = Trim(txtLookup_Host)
        If IS_IP_Address(txtLookup_Host) Then
            If IS_Valid_IP_Address(txtLookup_Host) Then
                'Start Search
                cmdLookup_Start.Enabled = False
                txtLookup_Results = ""
                lightLookup_TestInProgress.TurnLightON
                objLookup_IPInfo.ResolveAddress txtLookup_Host
            Else
                MsgBox "Invalid IP Address."
            End If
        Else
            'Start Search
            cmdLookup_Start.Enabled = False
            txtLookup_Results = ""
            lightLookup_TestInProgress.TurnLightON
            objLookup_IPInfo.ResolveName txtLookup_Host
        End If
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdLookup_Start_Click")
End Sub

'objLookup_IPInfo_Error
Private Sub objLookup_IPInfo_Error(ErrorCode As Integer, Description As String)
    
End Sub

'objLookup_IPInfo_RequestComplete
Private Sub objLookup_IPInfo_RequestComplete(RequestId As Long, StatusCode As Integer, Description As String)
    Dim objList As New ItemList
    Dim a As Long
    Dim arrTemp() As String
    
    objList.ADD "Request Completion: " & Description & vbCrLf & vbCrLf
    objList.ADD "Host Name: " & objLookup_IPInfo.HostName & vbCrLf
    objList.ADD "Host Address: " & objLookup_IPInfo.HostAddress & vbCrLf
    
    If Len(objLookup_IPInfo.HostAliases) > 0 Then
        objList.ADD vbCrLf
        objList.ADD "Host Aliases: "
        arrTemp = Split(Trim(objLookup_IPInfo.HostAliases), " ")
        If UBound(arrTemp) = 0 Then
            objList.ADD arrTemp(a) & vbCrLf
        Else
            objList.ADD vbCrLf
            For a = 0 To UBound(arrTemp)
                objList.ADD vbTab & arrTemp(a) & vbCrLf
            Next
        End If
    End If
    
    If Len(objLookup_IPInfo.OtherAddresses) > 0 Then
        objList.ADD vbCrLf
        objList.ADD "Other Addresses: "
        arrTemp = Split(Trim(objLookup_IPInfo.OtherAddresses), " ")
        If UBound(arrTemp) = 0 Then
            objList.ADD arrTemp(a) & vbCrLf
        Else
            objList.ADD vbCrLf
            For a = 0 To UBound(arrTemp)
                objList.ADD vbTab & arrTemp(a) & vbCrLf
            Next
        End If
    End If
    
    txtLookup_Results = txtLookup_Results & objList.CompleteString
    
    cmdLookup_Start.Enabled = True
    lightLookup_TestInProgress.TurnLightOFF
    Set objList = Nothing
End Sub


'==================================================================================
'                                    IP INFO
'==================================================================================

'cmdIP_INFO_Refresh_Click
Private Sub cmdIP_INFO_Refresh_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpcIP_INFO_IP_Config_Tree.Refresh
    vpcIP_INFO_IP_Config_Tree.SetFocus
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdIP_INFO_Refresh_Click")
End Sub

'cmdIP_Info_Collapse_Click
Private Sub cmdIP_Info_Collapse_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpcIP_INFO_IP_Config_Tree.CollapseAll
    vpcIP_INFO_IP_Config_Tree.SetFocus
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdIP_Info_Collapse_Click")
End Sub

'cmdIP_Info_Expand_Click
Private Sub cmdIP_Info_Expand_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    vpcIP_INFO_IP_Config_Tree.ExpandAll
    vpcIP_INFO_IP_Config_Tree.SetFocus
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdIP_Info_Expand_Click")
End Sub

'vpcIP_INFO_IP_Config_Tree_Error
Private Sub vpcIP_INFO_IP_Config_Tree_Error(ErrorMessage As String)
    MsgBox ErrorMessage
End Sub

'IP_Config_Tree1_Error
Private Sub IP_Config_Tree1_Error(ErrorMessage As String)
    MsgBox ErrorMessage
End Sub


'==================================================================================
'                           SINGLE COMPUTER PING
'==================================================================================

'cmdPING_ShowHide_Click
Private Sub cmdPING_ShowHide_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    txtPING_Details.Visible = Not (txtPING_Details.Visible)
    If txtPING_Details.Visible Then
        cmdPING_ShowHide.Caption = "Hide Details <<"
    Else
        cmdPING_ShowHide.Caption = "Show Details >>"
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdPING_ShowHide_Click")
    End If
End Sub

'cmdPING_Start_Click
Private Sub cmdPING_Start_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If cmdPING_Start.Caption = "Start" Then
        If txtHost = "" Then
            MsgBox "Please enter a host name or IP Address."
        Else
            vpcComputerCheck.ComputerNameOrIP = txtHost
            vpcComputerCheck.QuickTest = chkPING_QuickTest
            vpcComputerCheck.StartTest
            cmdPING_Start.Caption = "Stop"
        End If
    Else
        cmdPING_Start.Caption = "Start"
        vpcComputerCheck.StopTest
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdPING_Start_Click")
End Sub

'vpcComputerCheck_DetailsUpdate
Private Sub vpcComputerCheck_DetailsUpdate()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If txtPING_Details.Visible Then txtPING_Details = vpcComputerCheck.Details
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "vpcComputerCheck_DetailsUpdate")
End Sub

'vpcComputerCheck_Error
Private Sub vpcComputerCheck_Error(ErrorMessage As String)
    'MsgBox ErrorMessage
End Sub

'vpcComputerCheck_TestComplete
Private Sub vpcComputerCheck_TestComplete()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    cmdPING_Start.Caption = "Start"
    Beep
    Rest 200
    Beep
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "vpcComputerCheck_TestComplete")
End Sub


'==================================================================================
'                                     FORM
'==================================================================================

'Form_Load
Private Sub Form_Load()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Set Picture1.Picture = objCompanyDisplay.LogoImage
    About.Caption = "About " & objCompanyDisplay.CompanyName & "..."
    Call SizeAndMoveAllFrames
    Call HideAllFrames
    Frame1.Visible = True
    vpcComputerCheck.Caption = "Test Results"
    lightGenTest_InProgress.BaseColor = MODE_YELLOW
    lightTraceRoute_Test.BaseColor = MODE_YELLOW
    lightLookup_TestInProgress.BaseColor = MODE_YELLOW
    lblFooter.Caption = objCompanyDisplay.Description & vbCrLf & objCompanyDisplay.ContactInfo
    txtTraceRouteRemoteHost.List(0) = "www.yahoo.com"
    txtTraceRouteRemoteHost.List(1) = "www.google.com"
    txtTraceRouteRemoteHost.List(2) = "www.microsoft.com"
    txtTraceRouteRemoteHost.List(3) = "www.ibm.com"
    txtTraceRouteRemoteHost.List(4) = "www.ebay.com"
    txtTraceRouteRemoteHost.List(5) = "www.amazon.com"
    
    txtLookup_Host.List(0) = "www.yahoo.com"
    txtLookup_Host.List(1) = "www.google.com"
    txtLookup_Host.List(2) = "www.microsoft.com"
    txtLookup_Host.List(3) = "www.ibm.com"
    txtLookup_Host.List(4) = "www.ebay.com"
    txtLookup_Host.List(5) = "www.amazon.com"
    txtLookup_Host.List(6) = "127.0.0.1"
    txtLookup_Host.List(7) = "207.46.249.222"
    txtLookup_Host.List(8) = "66.218.71.83"
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Form_Load")
End Sub

'Form_Unload
Private Sub Form_Unload(Cancel As Integer)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Unload frmDetails
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "Form_Unload")
End Sub


'==================================================================================
'                                     TABS
'==================================================================================

'TabStrip1_Click
Private Sub TabStrip1_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Call HideAllFrames
    Select Case TabStrip1.SelectedItem.Key
        Case "PING"
            Frame1.Visible = True
        Case "TraceRoute"
            Frame2.Visible = True
        Case "IP_INFO"
            Frame3.Visible = True
        Case "WHOIS"
            Frame4.Visible = True
        Case "GeneralTest"
            Frame5.Visible = True
    End Select
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "TabStrip1_Click")
End Sub

'HideAllFrames
Private Sub HideAllFrames()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "HideAllFrames")
End Sub

'SizeAndMoveAllFrames
Private Sub SizeAndMoveAllFrames()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Frame2.Height = Frame1.Height
    Frame2.Width = Frame1.Width
    Frame2.Top = Frame1.Top
    Frame3.Height = Frame1.Height
    Frame3.Width = Frame1.Width
    Frame3.Top = Frame1.Top
    Frame4.Height = Frame1.Height
    Frame4.Width = Frame1.Width
    Frame4.Top = Frame1.Top
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "SizeAndMoveAllFrames")
End Sub


'==================================================================================
'                                TRACE ROUTE
'==================================================================================

'cmdTraceRouteStart_Click
Private Sub cmdTraceRouteStart_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If txtTraceRouteRemoteHost = "" Then
        MsgBox "Please enter a computer to trace."
    Else
        cmdTraceRouteStart.Enabled = False
        lightTraceRoute_Test.TurnLightON
        TraceRouteTree1.SetFocus
        'Start the Trace
        TraceRouteTree1.Start_Trace txtTraceRouteRemoteHost, _
                                    chkTraceRouteResolve.Value, _
                                    chkTraceRouteCompare.Value
                                    
        lightTraceRoute_Test.TurnLightOFF
        cmdTraceRouteStart.Enabled = True
        Beep
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdTraceRouteStart_Click")
End Sub

'cmdTraceRouteExpandall_Click
Private Sub cmdTraceRouteExpandall_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    TraceRouteTree1.ExpandAll
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdTraceRouteExpandall_Click")
End Sub

'cmdTraceRouteCollapseAll_Click
Private Sub cmdTraceRouteCollapseAll_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    TraceRouteTree1.CollapseAll
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdTraceRouteCollapseAll_Click")
End Sub

'chkTraceRouteResolve_Click
Private Sub chkTraceRouteResolve_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If chkTraceRouteResolve.Value Then
        MsgBox "Warning: Resolving host names may take a minute and may crash the program."
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "chkTraceRouteResolve_Click")
End Sub

'TraceRouteTree1_Error
Private Sub TraceRouteTree1_Error(ErrorMessage As String)
    MsgBox ErrorMessage
End Sub


'==================================================================================
'                            GENERAL NETWORK TEST
'==================================================================================

'cmdGeneralTest_Click
Private Sub cmdGeneralTest_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If cmdGeneralTest.Caption = "Start" Then
        If Not (vpcGenTest.TestInProgress) Then
            cmdGeneralTest.Caption = "Stop"
            Call RunGeneralTest
        End If
    Else
        cmdGeneralTest.Caption = "Start"
        vpcGenTest.StopTest
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdGeneralTest_Click")
End Sub

'cmdGenTestShowLog_Click
Private Sub cmdGenTestShowLog_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not objGenTestLog Is Nothing Then
        frmLogDisplay.Caption = "Test Details Log"
        frmLogDisplay.Show
        Set frmLogDisplay.objLog = objGenTestLog
        frmLogDisplay.ShowDetails
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdGenTestShowLog_Click")
End Sub

'cmdGenTestError_Click
Private Sub cmdGenTestError_Click()
    If Not objGenTestErrorLog Is Nothing Then
        frmLogDisplay.Caption = "Error Log"
        frmLogDisplay.Show
        Set frmLogDisplay.objLog = objGenTestErrorLog
        frmLogDisplay.ShowDetails
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdGenTestError_Click")
End Sub

'vpcGenTest_Error
Private Sub vpcGenTest_Error(ErrorMessage As String)
    MsgBox ErrorMessage
End Sub

'vpcGenTest_TestComplete
Private Sub vpcGenTest_TestComplete()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If blnGenTest_StopRequested Then
        blnGenTest_StopRequested = False
        cmdGeneralTest.Caption = "Start"
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "vpcGenTest_TestComplete")
End Sub
'chkGenTest_InternetConn_Click
Private Sub chkGenTest_InternetConn_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If chkGenTest_InternetConn Then
        MsgBox "Warning: Testing for an Internet connection may take several minutes. During Internet connection testing, this program may appear frozen."
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "chkGenTest_InternetConn_Click")
End Sub

'RunGeneralTest
Private Sub RunGeneralTest()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    '==========================================
    '       GENERAL OBJECT HIERARCHY
    '==========================================
    'IP_Config >> Connections >> IP Stacks
    '                         >> DNS Servers
    '==========================================
    '==========================================
    
    Dim lngConnection As Long
    Dim lngIP_Stack As Long
    Dim lngDNS_Server As Long
    Dim a As Long
    
    If Not objGenTestLog Is Nothing Then Set objGenTestLog = Nothing
    Set objGenTestLog = New ItemList
    
    If Not objGenTestErrorLog Is Nothing Then Set objGenTestErrorLog = Nothing
    Set objGenTestErrorLog = New ItemList
    
    IP_Config_Tree1.Refresh
    
    lblGenTest_InternetConn.Caption = "Internet Connection"
    lightGenTest_InternetConn.BaseColor = MODE_GREEN
    lightGenTest_InternetConn.TurnLightOFF
    
    lightGenTest_InProgress.TurnLightON
    
    'Make sure Log Display is closed or it will hang-on to the old object
    If frmLogDisplay.Visible Then
        Unload frmLogDisplay
        If frmLogDisplay.Caption = "Error Log" Then
            Call cmdGenTestError_Click
        Else
            Call cmdGenTestShowLog_Click
        End If
    End If
    
    'Test out anything with an IP Address
    With IP_Config_Tree1.IP_Configuartion
        objGenTestLog.ADD "=====================================================" & vbCrLf
        objGenTestLog.ADD "===          IP WIZ GENERAL NETWORK TEST          ===" & vbCrLf
        objGenTestLog.ADD "=====================================================" & vbCrLf
        objGenTestLog.ADD vbCrLf
        objGenTestLog.ADD "DATE: " & Now & vbCrLf
        objGenTestLog.ADD "Host: " & .HostName
        If .Primary_DNS_Suffix <> "" Then objGenTestLog.ADD "." & .Primary_DNS_Suffix
        objGenTestLog.ADD vbCrLf
        objGenTestLog.ADD "IP Routing Enabled: " & .IP_RoutingEnabled & vbCrLf
        objGenTestLog.ADD "WINS Proxy Enabled: " & .WINS_Proxy_Enabled & vbCrLf
        objGenTestLog.ADD vbCrLf & vbCrLf
    End With
    
    'Test self first
    Call GenTest_TestIP("Self (127.0.0.1)", "Self Test", "127.0.0.1", False)
    If cmdGeneralTest.Caption = "Start" Then GoTo TestHasBeenStopped
    
    'Connections
    For lngConnection = 1 To IP_Config_Tree1.IP_Configuartion.Connection.Count
        With IP_Config_Tree1.IP_Configuartion.Connection(lngConnection)
            objGenTestLog.ADD "Connection #" & CStr(lngConnection) & vbCrLf
            objGenTestLog.ADD vbTab & .Name & vbCrLf
            objGenTestLog.ADD vbTab & "Description: " & .Description & vbCrLf
            objGenTestLog.ADD vbTab & "Physical_Address: " & .Physical_Address & vbCrLf
            'Default Gateway
            If .Default_Gateway <> "" Then
                objGenTestLog.ADD vbCrLf & vbCrLf
                Call GenTest_TestIP("Default Gateway" & _
                                    " (" & .Default_Gateway & ")", _
                                    "Default Gateway Test", .Default_Gateway, False)
            Else
                objGenTestErrorLog.ADD "WARNING: No Default Gateway for Connection #" & lngConnection & vbCrLf
            End If
            'Primary WINS Server
            If .Primary_WINS_Server <> "" Then
                objGenTestLog.ADD vbCrLf & vbCrLf
                Call GenTest_TestIP("Primary WINS Server" & _
                                    " (" & .Primary_WINS_Server & ")", _
                                    "Primary WINS Server Test", .Primary_WINS_Server, False)
            End If
            'Secondary WINS Server
            If .Secondary_WINS_Server <> "" Then
                objGenTestLog.ADD vbCrLf & vbCrLf
                Call GenTest_TestIP("Secondary WINS Server" & _
                                    " (" & .Secondary_WINS_Server & ")", _
                                    "Secondary WINS Server Test", .Secondary_WINS_Server, False)
            End If
            'IP Stacks
            For lngIP_Stack = 1 To .IP_Stack.Count
                With .IP_Stack(lngIP_Stack)
                    Call GenTest_TestIP("IP Stack #" & CStr(lngIP_Stack) & _
                                        " (" & .IP_Address & ")", _
                                        "Self Test", .IP_Address, False)
                End With
                If cmdGeneralTest.Caption = "Start" Then GoTo TestHasBeenStopped
            Next
            
            'DNS Servers
            For lngDNS_Server = 1 To .DNS_Server.Count
                With .DNS_Server(lngDNS_Server)
                    Call GenTest_TestIP("DNS Server #" & CStr(lngDNS_Server) & _
                                        " (" & .IP_Address & ")", _
                                        "DNS Server Test", .IP_Address, False)
                End With
                If cmdGeneralTest.Caption = "Start" Then GoTo TestHasBeenStopped
            Next
        End With
    Next
    
    If chkGenTest_InternetConn Then
        'Check for Internet Connection
        Dim arrInternetSites(1 To 5) As String
        arrInternetSites(1) = "www.yahoo.com"
        arrInternetSites(2) = "www.microsoft.com"
        arrInternetSites(3) = "www.ibm.com"
        arrInternetSites(4) = "www.google.com"
        arrInternetSites(5) = "www.amazon.com"
        
        objGenTestLog.ADD vbCrLf & "Searching for Internet Connection..." & vbCrLf & vbCrLf
        
        For a = 1 To UBound(arrInternetSites)
            Call GenTest_TestIP(arrInternetSites(a), "Internet Connection Test", _
                                arrInternetSites(a), True)
            
            If vpcGenTest.ComputerFound Then
                'Internet Connection Found
                objGenTestLog.ADD vbCrLf & "Internet Connection is Present." & vbCrLf & vbCrLf
                lblGenTest_InternetConn.Caption = "Internet Connection Found."
                lightGenTest_InternetConn.BaseColor = MODE_GREEN
                lightGenTest_InternetConn.TurnLightON
                Exit For
            End If
        Next
        If a = UBound(arrInternetSites) + 1 Then
            'No Internet Connection Found
            objGenTestLog.ADD vbCrLf & "No Internet Connection Found." & vbCrLf & vbCrLf
            lblGenTest_InternetConn.Caption = "No Internet Connection Found."
            lightGenTest_InternetConn.BaseColor = MODE_RED
            lightGenTest_InternetConn.TurnLightON
        End If
    End If
    
TestHasBeenStopped:

    objGenTestLog.ADD vbCrLf & vbCrLf
    vpcGenTest.Reset
    vpcGenTest.Caption = ""
    cmdGeneralTest.Caption = "Start"
    lightGenTest_InProgress.TurnLightOFF
    Beep
    Rest 400
    MsgBox "General test is complete. Please refer to details and error logs for test results."
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "RunGeneralTest")
End Sub

'GenTest_TestIP
Private Sub GenTest_TestIP(strName As String, strTestType As String, _
                           strIP_Address As String, blnQuickCheck As Boolean)
    
    'Trap Hard Errors
    On Error GoTo HadHardError
                           
    If cmdGeneralTest.Caption = "Start" Then Exit Sub
    Call GenTest_HighlightNode(strIP_Address)
    vpcGenTest.Reset
    objGenTestLog.ADD vbTab & "Testing " & strName & "..." & vbCrLf & vbCrLf
    vpcGenTest.ComputerNameOrIP = strIP_Address
    vpcGenTest.Caption = strTestType & ": Pinging " & strIP_Address & "..."
    
    If blnQuickCheck Then
        vpcGenTest.QuickCheck (strIP_Address)
    Else
        vpcGenTest.StartTest
    End If
    'Wait for test to complete
    Do While vpcGenTest.TestInProgress
        Rest 500
        If cmdGeneralTest.Caption = "Start" Then Exit Do
    Loop
    objGenTestLog.ADD "******************************" & vbCrLf
    objGenTestLog.ADD "START: Results of Testing : " & strIP_Address & vbCrLf & vbCrLf
    objGenTestLog.ADD vpcGenTest.Details & vbCrLf
    objGenTestLog.ADD "END: Results of Testing: " & strIP_Address & vbCrLf
    objGenTestLog.ADD "******************************" & vbCrLf & vbCrLf
    'Check Results
    If vpcGenTest.ComputerFound Then
        If vpcGenTest.Percentage_SuccessfulTx < 66 Then
            If vpcGenTest.Percentage_SuccessfulTx < 33 Then
                objGenTestErrorLog.ADD "ERROR: "
            Else
                objGenTestErrorLog.ADD "WARNING: "
            End If
            objGenTestErrorLog.ADD "During " & strTestType & vbCrLf
            objGenTestErrorLog.ADD vbTab & "Losing packets to host " & strName & vbCrLf
        End If
        If vpcGenTest.Percentage_CommSpeed < 66 Then
            If vpcGenTest.Percentage_CommSpeed < 33 Then
                objGenTestErrorLog.ADD "ERROR: "
            Else
                objGenTestErrorLog.ADD "WARNING: "
            End If
            objGenTestErrorLog.ADD "During " & strTestType & vbCrLf
            objGenTestErrorLog.ADD vbTab & "Slow transmission rate to host " & strName & vbCrLf
        End If
    Else
        objGenTestErrorLog.ADD "ERROR: Host not found: " & strName & vbCrLf
    End If
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "GenTest_TestIP")
End Sub

'GenTest_HighlightNode
Private Sub GenTest_HighlightNode(strIP_Address As String)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim objNode As MSComctlLib.Node
    If strIP_Address = "127.0.0.1" Then
        For Each objNode In IP_Config_Tree1.IP_Configuartion_TreeView.Nodes
            If InStr(objNode.Text, "Host Name:") Then
                objNode.Selected = True
                If Not (frmLogDisplay.Visible) Then IP_Config_Tree1.SetFocus
                Exit For
            End If
        Next
    End If
    For Each objNode In IP_Config_Tree1.IP_Configuartion_TreeView.Nodes
        If InStr(objNode.Text, strIP_Address) Then
            objNode.Selected = True
            If Not (frmLogDisplay.Visible) Then IP_Config_Tree1.SetFocus
            Exit For
        End If
    Next
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "GenTest_HighlightNode")
End Sub




















