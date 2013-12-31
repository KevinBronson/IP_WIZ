VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{53337483-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "tracrt50.ocx"
Begin VB.UserControl TraceRouteTree 
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   ScaleHeight     =   5280
   ScaleWidth      =   7905
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   5580
      Top             =   90
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
            Picture         =   "TraceRouteTree.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":1E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":2290
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":26E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":2B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":2F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":33D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":382A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":3C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":40CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":4520
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":4972
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":4DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":5216
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":5668
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":5ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":5F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":635E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":67B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":6C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":78DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":85B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":9290
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":9F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":AC44
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":B91E
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":C5F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":D2D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":DFAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":EC86
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":F960
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":1063A
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":11314
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":11FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":12CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":1311A
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":1356C
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":139BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraceRouteTree.ctx":13E10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4630
      _Version        =   393217
      Indentation     =   141
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgIcons"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TRACEROUTELibCtl.TraceRoute TraceRoute1 
      Left            =   4770
      Top             =   270
      HopLimit        =   64
      HopTimeout      =   10
      QOSFlags        =   0
      ResolveNames    =   0   'False
      Timeout         =   60
      WinsockLoaded   =   -1  'True
   End
End
Attribute VB_Name = "TraceRouteTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'Internal Global Objects
Dim objIP_Config As IP_Config

'Public Events
Public Event Error(ErrorMessage As String)

'Internal Variables and Flags
Dim blnResolveNames As Boolean
Dim blnCompareTimes As Boolean
Dim strLastError As String


'==================================================================================
'                               PUBLIC PROPERTIES
'==================================================================================

'LastError
Public Property Get LastError() As String
    LastError = strLastError
End Property

'IP_Configuartion
Public Property Get IP_Configuartion() As IP_Config
    Set IP_Configuartion = objIP_Config
End Property

'objIP_Config_TreeView
Public Property Get TraceRout_TreeView() As TreeView
    Set TraceRout_TreeView = TreeView1
End Property


'==================================================================================
'                               METHODS
'==================================================================================

'Start_Trace
Public Sub Start_Trace(strRemoteHost As String, _
                        ResolveNames As Boolean, _
                        CompareTimes As Boolean)
    
    'Trap Hard Errors
    On Error GoTo HadHardError
                        
    blnResolveNames = ResolveNames
    blnCompareTimes = CompareTimes
    TreeView1.Nodes.Clear
    TraceRoute1.ResolveNames = blnResolveNames
    TraceRoute1.TraceTo strRemoteHost
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Start_Trace")
    RaiseEvent Error(strLastError)
End Sub

'Stop_Trace
Public Sub Stop_Trace()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Stop_Trace")
    RaiseEvent Error(strLastError)
End Sub

'Reset
Public Sub Reset()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Reset")
    RaiseEvent Error(strLastError)
End Sub

'ExpandAll
Public Sub ExpandAll()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    'Expand All
    Dim objNode As Node
    For Each objNode In TreeView1.Nodes
        objNode.Expanded = True
    Next
    If TreeView1.Nodes.Count > 0 Then TreeView1.Nodes(1).EnsureVisible
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "ExpandAll")
    RaiseEvent Error(strLastError)
End Sub

'CollapseAll
Public Sub CollapseAll()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim objNode As Node
    For Each objNode In TreeView1.Nodes
        objNode.Expanded = False
    Next
    If TreeView1.Nodes.Count > 0 Then TreeView1.Nodes(1).EnsureVisible
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "CollapseAll")
    RaiseEvent Error(strLastError)
End Sub


'==================================================================================
'                               PRIVATE FUNCTIONS
'==================================================================================

'AlignNumbersRight
Private Function AlignNumbersRight(lngInput As Long) As String
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim strNum As String
    If lngInput < 1000 Then
        strNum = Spaces(4 + (3 - Len(CStr(lngInput)))) & CStr(lngInput)
    Else
       strNum = Spaces(7 - Len(Format(CStr(lngInput), "#,###"))) & Format(CStr(lngInput), "#,###")
    End If
    AlignNumbersRight = strNum
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "AlignNumbersRight")
    RaiseEvent Error(strLastError)
End Function

'Spaces
Private Function Spaces(lngNumber As Long) As String
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim a As Long
    For a = 1 To lngNumber
        Spaces = Spaces & " "
    Next
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Spaces")
    RaiseEvent Error(strLastError)
End Function

'HaveNode
Private Function HaveNode(intHopNumber As Integer) As Boolean
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim objNode As MSComctlLib.Node
    For Each objNode In TreeView1.Nodes
        If objNode.Key = "HopNumber" & CStr(intHopNumber) Then
            HaveNode = True
            Exit For
        End If
    Next
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "HaveNode")
    RaiseEvent Error(strLastError)
End Function

'DefaultGateways
Private Function DefaultGateways() As String
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim a As Long
    For a = 1 To objIP_Config.Connection.Count
        DefaultGateways = DefaultGateways & "," & _
                          objIP_Config.Connection(a).Default_Gateway
    Next
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "DefaultGateways")
    RaiseEvent Error(strLastError)
End Function


'==================================================================================
'                                    EVENTS
'==================================================================================

'TraceRoute1_Error
Private Sub TraceRoute1_Error(ErrorCode As Integer, Description As String)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    strLastError = "Error message from Trace Route Object:" & vbCrLf & _
                    "Error Code: " & CStr(ErrorCode) & vbCrLf & Description
    RaiseEvent Error(strLastError)
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "TraceRoute1_Error")
    RaiseEvent Error(strLastError)
End Sub

'TraceRoute1_Hop
Private Sub TraceRoute1_Hop(HopNumber As Integer, HostAddress As String, Duration As Long)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim strNodeName_Hop As String
    Dim lngHopImage As Long
    Dim strText As String
    strNodeName_Hop = "HopNumber" & CStr(HopNumber)
    If TraceRoute1.RemoteHost = HostAddress Then
        lngHopImage = 13
    ElseIf HostAddress = "" Then
        lngHopImage = 4
    ElseIf InStr(DefaultGateways, HostAddress) Then
        lngHopImage = 36
    Else
        lngHopImage = 34
    End If
    If HopNumber < 10 Then
        strText = "Hop  " & HopNumber
    Else
        strText = "Hop " & HopNumber
    End If
    If Not (HaveNode(HopNumber)) Then
        If blnCompareTimes Then
            strText = strText & " [" & AlignNumbersRight(Duration) & " Milliseconds]"
        Else
            If Not (blnResolveNames) Then strText = strText & " [" & HostAddress & "]"
        End If
        TreeView1.Nodes.ADD , , strNodeName_Hop, strText, lngHopImage
        With TreeView1.Nodes.ADD(strNodeName_Hop, tvwChild, , "IP Address: " & HostAddress, 7)
            .Tag = "±±PINGABLE±±" & Trim(HostAddress)
        End With
    End If
    TreeView1.Nodes.ADD strNodeName_Hop, tvwChild, , "Duration = " & Duration & " ms", 7
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "TraceRoute1_Hop")
    RaiseEvent Error(strLastError)
End Sub

'TraceRoute1_HopResolved
Private Sub TraceRoute1_HopResolved(HopNumber As Integer, _
                                    StatusCode As Integer, HopHostName As String)
    
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Dim strNodeName_Hop As String
    strNodeName_Hop = "HopNumber" & CStr(HopNumber)
    TreeView1.Nodes.ADD strNodeName_Hop, tvwChild, , "Host Name: " & HopHostName, 7
    TreeView1.Nodes(strNodeName_Hop).Text = _
        TreeView1.Nodes(strNodeName_Hop).Text & _
        " " & HopHostName
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "TraceRoute1_HopResolved")
    RaiseEvent Error(strLastError)
End Sub

'TreeView1_DblClick
Private Sub TreeView1_DblClick()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not TreeView1.SelectedItem Is Nothing Then
        If TreeView1.SelectedItem.Tag <> "" Then
            If InStr(TreeView1.SelectedItem.Tag, "±±PINGABLE±±") Then
                frmPING.Show
                frmPING.PING Replace(TreeView1.SelectedItem.Tag, "±±PINGABLE±±", "")
            Else
                MsgBox TreeView1.SelectedItem.Tag
            End If
        End If
    End If
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "TreeView1_DblClick")
    RaiseEvent Error(strLastError)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Set objIP_Config = New IP_Config
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UserControl_Initialize")
    RaiseEvent Error(strLastError)
End Sub

'UserControl_Resize
Private Sub UserControl_Resize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    TreeView1.Width = UserControl.Width
    TreeView1.Height = UserControl.Height
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UserControl_Resize")
    RaiseEvent Error(strLastError)
End Sub























