VERSION 5.00
Begin VB.UserControl Light 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   ScaleHeight     =   780
   ScaleWidth      =   855
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   270
      Top             =   0
   End
   Begin VB.PictureBox picLight 
      BackColor       =   &H00FFFF00&
      Height          =   195
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "Light"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'Public Enums
Public Enum LIGHT_COLOR
    DARK_GREEN = &H4000&
    DARK_YELLOW = &H4040&
    DARK_RED = &H40&
    DARK_BLUE = &H400000
    LIGHT_GREEN = &HFF00&
    LIGHT_YELLOW = &HFFFF&
    LIGHT_RED = &HFF&
    LIGHT_BLUE = &HFFFF00
End Enum

Public Enum LIGHT_MODE
    MODE_GREEN = 1
    MODE_YELLOW = 2
    MODE_RED = 3
    MODE_BLUE = 4
End Enum

'Public Events
Public Event Blink()
Public Event Error(ErrorMessage As String)

'Interal Variables and Flags
Dim lngBaseColor As LIGHT_MODE
Dim lngBlinkCount As Long
Dim blnBeepDuringBlink As Long
Dim lngBlinkON_Duration As Long
Dim blnLightON_WhenStarted As Boolean
Dim strLastError As String


'==================================================================================
'                               PUBLIC PROPERTIES
'==================================================================================

'LastError
Public Property Get LastError() As String
    LastError = strLastError
End Property

'LightIsON
Public Property Get LightIsON() As Boolean
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    LightIsON = (picLight.BackColor = LIGHT_COLOR.LIGHT_BLUE Or _
    picLight.BackColor = LIGHT_COLOR.LIGHT_GREEN Or _
    picLight.BackColor = LIGHT_COLOR.LIGHT_RED Or _
    picLight.BackColor = LIGHT_COLOR.LIGHT_YELLOW)
    
    Exit Property
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "LightIsON")
    RaiseEvent Error(strLastError)
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
    picLight.Enabled = lngInput
End Property

'Height
Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Let Height(lngInput As Long)
    UserControl.Height = lngInput
End Property

'BaseColor
Public Property Get BaseColor() As LIGHT_MODE
    BaseColor = lngBaseColor
End Property

Public Property Let BaseColor(lngInput As LIGHT_MODE)
    Call SetBaseColor(lngInput)
End Property

'BlinkON_Duration
Public Property Get BlinkON_Duration() As Long
    BlinkON_Duration = lngBlinkON_Duration
End Property

Public Property Let BlinkON_Duration(lngInput As Long)
    lngBlinkON_Duration = lngInput
End Property


'===============================================================================
'                                   METHODS
'===============================================================================

'Blink
Public Sub Blink(lngTimesToBlink As Long, _
                 lngMillisecondsBetweenBlinks As Long, _
                 blnBeep As Boolean)
                 
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not Timer1.Enabled Then
        blnBeepDuringBlink = blnBeep
        lngBlinkCount = lngTimesToBlink
        Timer1.Interval = lngMillisecondsBetweenBlinks
        Timer1.Enabled = True
        blnLightON_WhenStarted = LightIsON
    End If
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Blink")
    RaiseEvent Error(strLastError)
End Sub

'TurnLightON
Public Sub TurnLightON()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If Not LightIsON Then ToggleLightONorOFF
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "TurnLightON")
    RaiseEvent Error(strLastError)
End Sub

'TurnLightOFF
Public Sub TurnLightOFF()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If LightIsON Then ToggleLightONorOFF
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "TurnLightOFF")
    RaiseEvent Error(strLastError)
End Sub


'==================================================================================
'                               PRIVATE FUNCTIONS
'==================================================================================

'SetBaseColor
Private Sub SetBaseColor(lngColor As LIGHT_MODE)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    lngBaseColor = lngColor
    Select Case lngColor
        Case LIGHT_MODE.MODE_BLUE
            If LightIsON Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_BLUE
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_BLUE
            End If
        Case LIGHT_MODE.MODE_GREEN
            If LightIsON Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_GREEN
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_GREEN
            End If
        Case LIGHT_MODE.MODE_RED
            If LightIsON Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_RED
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_RED
            End If
        Case LIGHT_MODE.MODE_YELLOW
            If LightIsON Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_YELLOW
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_YELLOW
            End If
    End Select
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "SetBaseColor")
    RaiseEvent Error(strLastError)
End Sub

'ToggleLightONorOFF
Private Sub ToggleLightONorOFF()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Select Case BaseColor
        Case LIGHT_MODE.MODE_BLUE
            If picLight.BackColor = LIGHT_COLOR.DARK_BLUE Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_BLUE
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_BLUE
            End If
        Case LIGHT_MODE.MODE_GREEN
            If picLight.BackColor = LIGHT_COLOR.DARK_GREEN Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_GREEN
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_GREEN
            End If
        Case LIGHT_MODE.MODE_RED
            If picLight.BackColor = LIGHT_COLOR.DARK_RED Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_RED
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_RED
            End If
        Case LIGHT_MODE.MODE_YELLOW
            If picLight.BackColor = LIGHT_COLOR.DARK_YELLOW Then
                picLight.BackColor = LIGHT_COLOR.LIGHT_YELLOW
            Else
                picLight.BackColor = LIGHT_COLOR.DARK_YELLOW
            End If
    End Select
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "ToggleLightONorOFF")
    RaiseEvent Error(strLastError)
End Sub


'==================================================================================
'                                    EVENTS
'==================================================================================

'Timer1_Timer
Private Sub Timer1_Timer()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If lngBlinkCount > 0 Then
        Call TurnLightON
        If blnBeepDuringBlink Then Beep
        UserControl.Refresh
        Sleep lngBlinkON_Duration
        Call TurnLightOFF
        RaiseEvent Blink
        If lngBlinkCount = 1 Then
            If blnLightON_WhenStarted Then
                TurnLightON
            Else
                TurnLightOFF
            End If
        End If
    Else
        Timer1.Interval = 0
        Timer1.Enabled = False
    End If
    lngBlinkCount = lngBlinkCount - 1
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Timer1_Timer")
    RaiseEvent Error(strLastError)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    'Set Defaults
    BaseColor = MODE_GREEN
    TurnLightOFF
    lngBlinkON_Duration = 400
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UserControl_Initialize")
    RaiseEvent Error(strLastError)
End Sub

'UserControl_Resize
Private Sub UserControl_Resize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    picLight.Width = UserControl.Width
    picLight.Height = UserControl.Height
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UserControl_Resize")
    RaiseEvent Error(strLastError)
End Sub




