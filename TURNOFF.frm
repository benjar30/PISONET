VERSION 5.00
Begin VB.Form monitoroff 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "5"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   2280
   End
End
Attribute VB_Name = "monitoroff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' =========== Paste the lines below into a standard module ===============

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer
Private Declare Function MapVirtualKey Lib "user32" _
    Alias "MapVirtualKeyA" _
    (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Declare Function SendInput Lib "user32" _
    (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long
Private Type KeyboardInput       '   typedef struct tagINPUT {
   dwType As Long                '     DWORD type;
   wVK As Integer                '     union {MOUSEINPUT mi;
   wScan As Integer              '            KEYBDINPUT ki;
   dwFlags As Long               '            HARDWAREINPUT hi;
   dwTime As Long                '     };
   dwExtraInfo As Long           '   }INPUT, *PINPUT;
   dwPadding As Currency         '
End Type
'SendInput constants
Private Const INPUT_KEYBOARD As Long = 1
Private Const KEYEVENTF_KEYUP As Long = 2
Private Const VK_CAPITAL = &H14
Const SC_MONITORPOWER = &HF170&
Const MON_OFF = 2&
Const WM_SYSCOMMAND = &H112
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Private Sub procTurnOff()

  SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MON_OFF

End Sub


Public Function CapsLock() As Boolean
   ' Determine whether CAPSLOCK key is toggled on.
   CapsLock = CBool(GetKeyState(VK_CAPITAL) And 1)
End Function

Public Sub SetCapsLockState(bEnabled As Boolean)
    'CapsLock is already in desired state. Nothing to do.
    If CapsLock = bEnabled Then Exit Sub

    PressCapsLock
End Sub

Private Sub PressCapsLock()
    GenerateKeyboardEvent VK_CAPITAL, 0
    GenerateKeyboardEvent VK_CAPITAL, KEYEVENTF_KEYUP
End Sub

Private Sub GenerateKeyboardEvent(VirtualKey As Long, Flags As Long)
    Dim kevent As KeyboardInput

    With kevent
        .dwType = INPUT_KEYBOARD
        .wScan = MapVirtualKey(VirtualKey, 0)
        .wVK = VirtualKey
        .dwTime = 0
        .dwFlags = Flags
    End With
    SendInput 1, kevent, Len(kevent)
End Sub

Public Sub onnow()
SetCapsLockState True
SetCapsLockState False
End Sub


'====================End=================================================
Private Sub Timer1_Timer()
Call procTurnOff
End Sub

