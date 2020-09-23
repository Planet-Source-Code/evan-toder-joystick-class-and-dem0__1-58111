VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start joystick"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   495
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find valid joystick"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Index           =   5
      Left            =   1575
      TabIndex        =   7
      Top             =   1665
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Index           =   4
      Left            =   1575
      TabIndex        =   6
      Top             =   1350
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Index           =   3
      Left            =   1575
      TabIndex        =   5
      Top             =   1044
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Index           =   2
      Left            =   1575
      TabIndex        =   4
      Top             =   741
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Index           =   1
      Left            =   1575
      TabIndex        =   3
      Top             =   438
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   2
      Top             =   135
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents cJoy As cJoy
Attribute cJoy.VB_VarHelpID = -1

Private Sub cJoy_JoyError(sErrMessage As String)
'
 Debug.Print sErrMessage
End Sub

'-- event raised from cJoy which is called every 25 milliseconds
'-- from timer_module and report all the joystick info we need
Private Sub cJoy_JoyInfo(BtnPressed As Long, leftStickX As Long, _
                         leftStickY As Long, rightStickX As Long, _
                         rightStickY As Long, DpadPos As Long)
'
 Label1(0) = "BUTTON PRESSED: " & BtnPressed
 Label1(1) = "LEFT JOYSTICK X POS: " & leftStickX
 Label1(2) = "LEFT JOYSTICK y POS: " & leftStickY
 Label1(3) = "RIGHT JOYSTICK X POS: " & rightStickX
 Label1(4) = "RIGHT JOYSTICK y POS: " & rightStickY
 Label1(5) = "D-PAD POSITION: " & DpadPos
End Sub

Private Sub Command1_Click(Index As Integer)
  Dim sMod As String
  
  If Index = 0 Then '-- find a valid joystick
     '-- test joystick 1 and report
     sMod = IIf(cJoy.IsJoystick1_Valid = True, " ", "NOT ")
     MsgBox "Joystick #1 is " & sMod & "a valid joystick"
      '-- test joystick 2 and report
     sMod = IIf(cJoy.IsJoystick2_Valid = True, " ", "NOT ")
     MsgBox "Joystick #2 is " & sMod & "a valid joystick"
     
  ElseIf Index = 1 Then
     '-- on MY computer joystick #2 is the valid 1
     '-- on your system it might be different so make
     '-- sure you enter the right number
     cJoy.Start_JoyMonitor JOYSTICKID1
     
  End If
End Sub

Private Sub Form_Load()
  Set cJoy = New cJoy
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set cJoy = Nothing
End Sub
 
