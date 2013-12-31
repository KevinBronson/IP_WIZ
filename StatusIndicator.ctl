VERSION 5.00
Begin VB.UserControl StatusIndicator 
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ScaleHeight     =   1530
   ScaleWidth      =   480
   Begin VB.PictureBox picTop 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox picBottom 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   360
      Width           =   195
   End
End
Attribute VB_Name = "StatusIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'Public Enums
Public Enum POSITION_MODE_VALUES
    MODE_HORIZONTAL = 1
    MODE_VERTICAL = 2
End Enum

'Public Events
Public Event Error(ErrorMessage As String)

'Interal Variables and Flags
Dim lngPercentage As Long
Dim strLastError As String


'==================================================================================
'                               PUBLIC PROPERTIES
'==================================================================================

'PositionMode
Public PositionMode As POSITION_MODE_VALUES

'LastError
Public Property Get LastError() As String
    LastError = strLastError
End Property

'Height
Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Let Height(lngInput As Long)
    UserControl.Height = lngInput
End Property

'Width
Public Property Get Width() As Long
    Width = UserControl.Width
End Property

Public Property Let Width(lngInput As Long)
    UserControl.Width = lngInput
End Property

'Percentage
Public Property Get Percentage() As Long
    Percentage = lngPercentage
End Property

Public Property Let Percentage(lngInput As Long)
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    If lngInput > 100 Then lngInput = 100
    If lngInput < 0 Then lngInput = 0
    lngPercentage = lngInput
    'Get the resize event to fire
    Dim test As Long
    test = UserControl.Width
    test = test + 50
    UserControl.Width = test
    test = test - 50
    UserControl.Width = test
    
    Exit Property
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "Percentage")
    RaiseEvent Error(strLastError)
End Property


'==================================================================================
'                                   METHODS
'==================================================================================

'ShowAllRed
Public Sub ShowAllRed()
    picTop.BackColor = vbRed
End Sub



'==================================================================================
'                               PRIVATE FUNCTIONS
'==================================================================================

'PercentToPixels
Private Function PercentToPixels(lngPercentInput As Long, lngMaxPixels) As Long
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    PercentToPixels = lngPercentInput * (lngMaxPixels / 100)
    If lngPercentInput = 100 Then PercentToPixels = lngMaxPixels
    If lngPercentInput = 0 Then PercentToPixels = 0
    
    Exit Function
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "PercentToPixels")
    RaiseEvent Error(strLastError)
End Function

'SetColor
Private Sub SetColor()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    picTop.BackColor = vbWhite
    Select Case lngPercentage
        Case Is < 34 'Red
            picBottom.BackColor = vbRed
        Case 34 To 66 'Yellow
            picBottom.BackColor = vbYellow
        Case Is > 66 'Green
            picBottom.BackColor = vbGreen
    End Select
    picTop.Visible = Not (lngPercentage = 100)
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "SetColor")
    RaiseEvent Error(strLastError)
End Sub


'==================================================================================
'                                   EVENTS
'==================================================================================

'UserControl_Resize
Private Sub UserControl_Resize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    'MsgBox "UserControl_Resize"
    'Position all the top left corners to the same spot
    picBottom.Top = UserControl.ScaleTop
    picTop.Top = UserControl.ScaleTop
    picBottom.Left = UserControl.ScaleLeft
    picTop.Left = UserControl.ScaleLeft
    
    Select Case PositionMode
        Case POSITION_MODE_VALUES.MODE_HORIZONTAL
            'HORIZONTAL
            picBottom.ZOrder
            'Bottom
            picBottom.Width = PercentToPixels(lngPercentage, UserControl.Width)
            picBottom.Height = UserControl.Height
            'Top
            picTop.Width = UserControl.Width
            picTop.Height = UserControl.Height
            
        Case POSITION_MODE_VALUES.MODE_VERTICAL
            'VERTICAL
            picTop.ZOrder
            'Bottom
            picBottom.Width = UserControl.Width
            picBottom.Height = UserControl.Height
            'Top
            picTop.Width = UserControl.Width
            picTop.Height = UserControl.Height - PercentToPixels(lngPercentage, UserControl.Height)
    End Select
    Call SetColor
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UserControl_Resize")
    RaiseEvent Error(strLastError)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    PositionMode = MODE_VERTICAL
    
    Exit Sub
HadHardError:
    strLastError = FormatErrorMessage(Err, UserControl.Name, "UserControl_Initialize")
    RaiseEvent Error(strLastError)
End Sub
















