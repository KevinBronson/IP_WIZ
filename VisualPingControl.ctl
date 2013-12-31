VERSION 5.00
Object = "{53337443-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "ping50.ocx"
Begin VB.UserControl VisualPingControl 
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   ScaleHeight     =   2100
   ScaleWidth      =   9960
   Begin VB.Timer Timer1 
      Left            =   540
      Top             =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9825
      Begin VB.Frame FrameSuccessfulTrans 
         Caption         =   "Successful Transmissions"
         Height          =   1095
         Left            =   6570
         TabIndex        =   13
         Top             =   270
         Width           =   3165
         Begin IP_WIZ.StatusIndicator indPacketLoss 
            Height          =   735
            Left            =   90
            TabIndex        =   14
            Top             =   270
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   1296
         End
         Begin VB.Label lblPacketPercentage 
            Caption         =   "0%"
            Height          =   195
            Left            =   270
            TabIndex        =   16
            Top             =   810
            Width           =   2625
         End
         Begin VB.Label lblPacketLossMessage 
            Height          =   285
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   2625
         End
         Begin VB.Line Line5 
            X1              =   270
            X2              =   3060
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line6 
            X1              =   270
            X2              =   270
            Y1              =   270
            Y2              =   720
         End
      End
      Begin VB.Frame FrameCommSpeed 
         Caption         =   "Communication Speed"
         Height          =   1095
         Left            =   3330
         TabIndex        =   9
         Top             =   270
         Width           =   3165
         Begin IP_WIZ.StatusIndicator indSpeed 
            Height          =   735
            Left            =   90
            TabIndex        =   10
            Top             =   270
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   1296
         End
         Begin VB.Label lblSpeedMessage 
            Height          =   285
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   2625
         End
         Begin VB.Label lblSpeedPercentage 
            Caption         =   "0%"
            Height          =   195
            Left            =   270
            TabIndex        =   11
            Top             =   810
            Width           =   2715
         End
         Begin VB.Line Line7 
            X1              =   270
            X2              =   3060
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line8 
            X1              =   270
            X2              =   270
            Y1              =   270
            Y2              =   720
         End
      End
      Begin VB.Frame FrameStatus 
         Caption         =   "Status"
         Height          =   1095
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   3165
         Begin VB.CommandButton cmdDetails 
            Caption         =   "Details ..."
            Height          =   300
            Left            =   2070
            TabIndex        =   2
            Top             =   720
            Width           =   1005
         End
         Begin IP_WIZ.Light LightOK 
            Height          =   105
            Left            =   90
            TabIndex        =   3
            Top             =   315
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
         End
         Begin IP_WIZ.Light LightTxRx 
            Height          =   105
            Left            =   90
            TabIndex        =   4
            Top             =   585
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
         End
         Begin IP_WIZ.Light LightTestInProgress 
            Height          =   105
            Left            =   90
            TabIndex        =   5
            Top             =   855
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
         End
         Begin VB.Label lblTestInProgress 
            Caption         =   "Test in Progress"
            Height          =   195
            Left            =   450
            TabIndex        =   8
            Top             =   810
            Width           =   2535
         End
         Begin VB.Label lblTxRx 
            Caption         =   "Sending\Recieving"
            Height          =   195
            Left            =   450
            TabIndex        =   7
            Top             =   540
            Width           =   2535
         End
         Begin VB.Label lblOK 
            Caption         =   "Computer Found"
            Height          =   195
            Left            =   450
            TabIndex        =   6
            Top             =   270
            Width           =   2535
         End
      End
   End
   Begin PINGLibCtl.Ping objPing 
      Left            =   90
      Top             =   1620
      PacketSize      =   64
      QOSFlags        =   0
      Timeout         =   10
      WinsockLoaded   =   -1  'True
   End
End
Attribute VB_Name = "VisualPingControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'Event Declarations
Public Event TxRx() 'Transmitting or Recieving
Public Event TestComplete() 'Test has been completed
Public Event Error(ErrorMessage As String) 'Error has occurred
Public Event DetailsButtonClicked() 'Details button has been clicked
Public Event DetailsUpdate() 'When the details string (strDetails) has been updated

'Internal Flags
Dim blnQuickTest As Boolean 'Quick Test - Does 5 pings if TRUE, otherwise does 20 pings
Dim lngPingCount As Long 'Number of Pings left to complete
Dim lngTotalPingCount As Long 'Total Number of Times to PING Remote Host
Dim blnStopRequested As Boolean 'True when the stop method is called

'Internal Variables
Dim lngPercentage_CommSpeed As Long
Dim lngPercentage_SuccessfulTx As Long
Dim strDetails As String
Dim objDetailsList As ItemList
Dim lngRecieved As Long
Dim arrResponseTimes() As Long
Dim strLastError As String
Dim dblAvgSpeed As Double
Dim dblMaxSpeed As Double
Dim dblMinSpeed As Double


'==================================================================================
'                               PUBLIC PROPERTIES
'==================================================================================

Public ComputerNameOrIP As String

'LastError
Public Property Get LastError() As String
    LastError = strLastError
End Property

'ComputerFound
Public Property Get ComputerFound() As Boolean
    ComputerFound = (lngRecieved > 0)
End Property

'Details
Public Property Get Details() As String
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not objDetailsList Is Nothing Then
        Details = objDetailsList.CompleteString
    Else
        Details = ""
    End If
    
    Exit Property
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Details")
    RaiseEvent Error(strLastError)
End Property

'Percentage_SuccessfulTx
Public Property Get Percentage_SuccessfulTx() As Long
    Percentage_SuccessfulTx = lngPercentage_SuccessfulTx
End Property

'Percentage_CommSpeed
Public Property Get Percentage_CommSpeed() As Long
    Percentage_CommSpeed = lngPercentage_CommSpeed
End Property

'TestInProgress
Public Property Get TestInProgress() As Boolean
    TestInProgress = LightTestInProgress.LightIsON
End Property

'QuickTest
Public Property Get QuickTest() As Boolean
    QuickTest = blnQuickTest
End Property

Public Property Let QuickTest(blnInput As Boolean)
    blnQuickTest = blnInput
End Property

'Caption
Public Property Get Caption() As String
    Caption = Frame1.Caption
End Property

Public Property Let Caption(strInput As String)
    Frame1.Caption = strInput
End Property

'Width
Public Property Get Width() As Long
    Width = UserControl.Width
End Property

Public Property Let Width(lngInput As Long)
    UserControl.Width = lngInput
End Property

'Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(lngInput As Boolean)
    UserControl.Enabled = lngInput
End Property

'Height
Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Let Height(lngInput As Long)
    UserControl.Height = lngInput
End Property

'DetailsButtonVisible
Public Property Get DetailsButtonVisible() As Boolean
    DetailsButtonVisible = cmdDetails.Visible
End Property

Public Property Let DetailsButtonVisible(lngInput As Boolean)
     cmdDetails.Visible = lngInput
End Property


'===============================================================================
'                                   METHODS
'===============================================================================

'QuickCheck
Public Function QuickCheck(strHost As String) As Boolean
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    ComputerNameOrIP = strHost
    Call StartTest(1)
    Do While TestInProgress
        Rest 500
    Loop
    QuickCheck = ComputerFound
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "QuickCheck")
    RaiseEvent Error(strLastError)
End Function

'StartTest
Public Sub StartTest(Optional lngTotalNumberOfTimesToPing As Long)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not (TestInProgress) Then
        Call Reset 'Clear Out All Variables
        'Set flags and start timer
        blnStopRequested = False
        LightTestInProgress.TurnLightON
        Timer1.Interval = 200 'Fire the timer soon to get it going
        Timer1.Enabled = True
        'Figure out the Test Mode
        If blnQuickTest Then
            lngTotalPingCount = 5
        Else
            lngTotalPingCount = 20
        End If
        If Not (lngTotalNumberOfTimesToPing = 0) Then lngTotalPingCount = lngTotalNumberOfTimesToPing
        lngPingCount = lngTotalPingCount
        ReDim arrResponseTimes(1 To lngTotalPingCount)
        'Get PING Object Ready
        objPing.Timeout = 2
        objPing.TimeToLive = 128
        'Get Details List ready
        If Not objDetailsList Is Nothing Then Set objDetailsList = Nothing
        Set objDetailsList = New ItemList
        'Record what is happening in Details List
        objDetailsList.ADD "Pinging "
        objDetailsList.ADD ComputerNameOrIP
        objDetailsList.ADD " with "
        objDetailsList.ADD objPing.PacketSize
        objDetailsList.ADD " bytes of data:"
        objDetailsList.ADD vbCrLf & vbCrLf
    Else
        'Ignore Request
    End If
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "StartTest")
    RaiseEvent Error(strLastError)
End Sub

'StopTest
Public Sub StopTest()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If TestInProgress Then blnStopRequested = True
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "StopTest")
    RaiseEvent Error(strLastError)
End Sub

'Reset
Public Sub Reset()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not (TestInProgress) Then
        lngRecieved = 0
        If Not objDetailsList Is Nothing Then Set objDetailsList = Nothing
        LightOK.BaseColor = MODE_GREEN
        LightOK.TurnLightOFF
        lblOK.Caption = "Computer Found"
        LightTxRx.TurnLightOFF
        LightTestInProgress.TurnLightOFF
        
        lblSpeedMessage = ""
        lblSpeedPercentage = "0%"
        indSpeed.Percentage = 0
        
        lblPacketLossMessage = ""
        lblPacketPercentage = "0%"
        indPacketLoss.Percentage = 0
        
        RaiseEvent DetailsUpdate
    End If
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Reset")
    RaiseEvent Error(strLastError)
End Sub


'===============================================================================
'                                   EVENTS
'===============================================================================

'UserControl_Initialize
Private Sub UserControl_Initialize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    LightTestInProgress.BaseColor = MODE_YELLOW
    DetailsButtonVisible = False
    blnQuickTest = True
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UserControl_Initialize")
    RaiseEvent Error(strLastError)
End Sub

'cmdDetails_Click
Private Sub cmdDetails_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    RaiseEvent DetailsButtonClicked
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "cmdDetails_Click")
    RaiseEvent Error(strLastError)
End Sub

'objPing_Error
Private Sub objPing_Error(ErrorCode As Integer, Description As String)
    strLastError = "Error message from Ping Object:" & vbCrLf & _
                    "Error Code: " & CStr(ErrorCode) & vbCrLf & Description
    RaiseEvent Error(strLastError)
End Sub


'===============================================================================
'                                   TIMER
'===============================================================================

'Timer1_Timer
Private Sub Timer1_Timer()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    On Error Resume Next
    Timer1.Enabled = False 'To stop reentrancy
    If blnStopRequested Then
        'STOP has been requested
        blnStopRequested = False
        lngPingCount = 0
        LightTestInProgress.TurnLightOFF
        objDetailsList.ADD vbCrLf
        objDetailsList.ADD "==================================" & vbCrLf
        objDetailsList.ADD "    Testing Has Been Stopped" & vbCrLf
        objDetailsList.ADD "==================================" & vbCrLf
        objDetailsList.ADD vbCrLf
        RaiseEvent DetailsUpdate
        RaiseEvent TestComplete
    Else
        lngPingCount = lngPingCount - 1
        Timer1.Interval = 1000 'Sets milliseconds between pings
        'Check to make sure ready to ping first
        If ComputerNameOrIP <> "" Then
            ComputerNameOrIP = Trim(ComputerNameOrIP)
            LightTxRx.Blink 2, 300, False 'Blink light to show currently pinging
            '*** Start PING Test ***
            objPing.RemoteHost = ComputerNameOrIP
            'See if an IP Address was given or found
            If objPing.RemoteHost = "" Then
                If IS_IP_Address(ComputerNameOrIP) Then
                    If IS_Valid_IP_Address(ComputerNameOrIP) Then
                        'This shouldn't happen
                    Else
                        RaiseEvent Error("Invalid IP Address")
                    End If
                Else
                    RaiseEvent Error("Unable to Resolve Computer Name")
                End If
            Else
                'See if anything was recieved
                If objPing.ResponseSource <> "" Then lngRecieved = lngRecieved + 1
                'Show if the computer has been found yet
                If ComputerFound Then
                    LightOK.BaseColor = MODE_GREEN
                    LightOK.TurnLightON
                    lblOK.Caption = "Computer Found"
                Else
                    LightOK.BaseColor = MODE_RED
                    LightOK.TurnLightON
                    lblOK.Caption = "Computer NOT Found"
                End If
                'Stash Response Time
                arrResponseTimes(lngPingCount + 1) = objPing.ResponseTime
            End If
        Else
            RaiseEvent Error("No Computer Name or IP given.")
        End If
        'Show and Tell - Display test results and update details
        UpdateDetailsList
        
        'See if test is complete
        If lngPingCount > 0 Then
            'Test is still going
            Dim X As Long
            X = Timer1.Interval
            Timer1.Enabled = True
        Else
            'Test is done
            Rest 1000 'Wait for TxRx Light to stop blinking
            Call CalculateStats
            'Add Final Stats to Detail List
            'Recieved Stats
            objDetailsList.ADD vbCrLf
            objDetailsList.ADD "Ping statistics for "
            objDetailsList.ADD objPing.RemoteHost
            objDetailsList.ADD ":"
            objDetailsList.ADD vbCrLf
            objDetailsList.ADD "   Packets: Sent =  "
            objDetailsList.ADD lngTotalPingCount
            objDetailsList.ADD ", Received = "
            objDetailsList.ADD lngRecieved
            objDetailsList.ADD ", Lost = "
            objDetailsList.ADD lngTotalPingCount - lngRecieved
            objDetailsList.ADD " ("
            objDetailsList.ADD 100 - lngPercentage_SuccessfulTx
            objDetailsList.ADD "% loss)"
            objDetailsList.ADD vbCrLf
            'Speed Stats
            objDetailsList.ADD "Approximate round trip times in milli-seconds:"
            objDetailsList.ADD vbCrLf
            objDetailsList.ADD "   Minimum = "
            objDetailsList.ADD dblMinSpeed
            objDetailsList.ADD "ms, Maximum = "
            objDetailsList.ADD dblMaxSpeed
            objDetailsList.ADD "ms, Average = "
            objDetailsList.ADD Int(dblAvgSpeed)
            objDetailsList.ADD "ms"
            objDetailsList.ADD vbCrLf
            'Reset Variables
            Timer1.Enabled = False
            LightTestInProgress.TurnLightOFF
            Call UpdateDisplays
            RaiseEvent DetailsUpdate
            RaiseEvent TestComplete
        End If
    End If
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Timer1_Timer")
    RaiseEvent Error(strLastError)
End Sub


'===============================================================================
'                         INTERNAL DISPLAY FUNCTIONS
'===============================================================================

'CalculateStats
Private Function CalculateStats()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    '-------------------------
    'Successful Transmissions
    lngPercentage_SuccessfulTx = (lngRecieved / lngTotalPingCount) * 100
    
    '-------------------------
    'Response Time Stats
    Dim a As Long
    'Reset Variables
    dblMaxSpeed = 0
    dblMinSpeed = 0
    dblAvgSpeed = 0
    'MAX Response Time
    For a = 1 To UBound(arrResponseTimes)
        If dblMaxSpeed < arrResponseTimes(a) Then dblMaxSpeed = arrResponseTimes(a)
    Next
    dblMinSpeed = arrResponseTimes(1)
    'MIN Response Time
    For a = 1 To UBound(arrResponseTimes)
        If dblMinSpeed > arrResponseTimes(a) Then dblMinSpeed = arrResponseTimes(a)
    Next
    'Average Response Time
    For a = 1 To UBound(arrResponseTimes)
        dblAvgSpeed = dblAvgSpeed + arrResponseTimes(a)
    Next
    dblAvgSpeed = dblAvgSpeed / UBound(arrResponseTimes)
    'Create 'ball-park' indicator of how fast the connection is
    lngPercentage_CommSpeed = 100 - (dblAvgSpeed / 10)
    If (lngRecieved < 1) Then lngPercentage_CommSpeed = 0
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "CalculateStats")
    RaiseEvent Error(strLastError)
End Function

'UpdateDisplays
Private Function UpdateDisplays()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    'Speed Display
    indSpeed.Percentage = lngPercentage_CommSpeed
    lblSpeedPercentage.Caption = CStr(lngPercentage_CommSpeed) & "%"
    Select Case lngPercentage_CommSpeed
        Case Is < 33
            lblSpeedMessage = "Slow"
            If lngPercentage_CommSpeed = 0 Then indSpeed.ShowAllRed
        Case 33 To 66
            lblSpeedMessage = "So-So"
        Case Is > 66
            lblSpeedMessage = "Good"
    End Select
    'Successful Transmissions
    indPacketLoss.Percentage = lngPercentage_SuccessfulTx
    lblPacketPercentage.Caption = CStr(lngPercentage_SuccessfulTx) & "%"
    Select Case lngPercentage_SuccessfulTx
        Case Is < 33
            lblPacketLossMessage = "Very Bad"
            If lngPercentage_SuccessfulTx = 0 Then indPacketLoss.ShowAllRed
        Case 33 To 66
            lblPacketLossMessage = "Not too good"
        Case Is > 66
            lblPacketLossMessage = "OK"
            If lngPercentage_SuccessfulTx > 90 Then lblPacketLossMessage = "Good"
    End Select
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UpdateDisplays")
    RaiseEvent Error(strLastError)
End Function

'UpdateDetailsList
Private Function UpdateDetailsList()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Err.Description <> "" Then
        If Err.Description = "Timeout." Then
            objDetailsList.ADD "      Request timed out" & vbCrLf
        Else
            objDetailsList.ADD "      ERROR: " & Err.Description & vbCrLf
        End If
        Err.Clear
    Else
        If objPing.RemoteHost = "" Then
            objDetailsList.ADD "      Invalid IP Address or Computer Name Could Not be Resolved"
            objDetailsList.ADD vbCrLf
        Else
            If objPing.ResponseSource <> "" Then
                objDetailsList.ADD "      Reply from "
                objDetailsList.ADD objPing.RemoteHost
                objDetailsList.ADD ": bytes="
                objDetailsList.ADD objPing.PacketSize
                objDetailsList.ADD " time"
                If objPing.ResponseTime = 0 Then
                    objDetailsList.ADD "<10"
                Else
                    objDetailsList.ADD "=" & objPing.ResponseTime
                End If
                objDetailsList.ADD "ms TTL="
                objDetailsList.ADD objPing.TimeToLive
                objDetailsList.ADD vbCrLf
            Else
                objDetailsList.ADD "      NO Reply from " & objPing.RemoteHost
                objDetailsList.ADD vbCrLf
            End If
        End If
    End If
    RaiseEvent DetailsUpdate
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UpdateDetailsList")
    RaiseEvent Error(strLastError)
End Function


















