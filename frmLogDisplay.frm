VERSION 5.00
Begin VB.Form frmLogDisplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Display"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDetails 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5775
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   465
      Left            =   1170
      TabIndex        =   1
      Top             =   2880
      Width           =   2445
   End
End
Attribute VB_Name = "frmLogDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Public WithEvents objLog As ItemList
Attribute objLog.VB_VarHelpID = -1

'cmdClose_Click
Private Sub cmdClose_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Unload Me
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "cmdClose_Click")
End Sub

'ShowDetails
Public Sub ShowDetails()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    txtDetails = objLog.CompleteString
    txtDetails.SelStart = Len(txtDetails) 'Scrolls down to end of text
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "ShowDetails")
End Sub

'objLog_Change
Private Sub objLog_Change()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    Call ShowDetails
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "objLog_Change")
End Sub


