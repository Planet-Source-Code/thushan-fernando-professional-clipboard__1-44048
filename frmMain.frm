VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Real-Time Clipboard Status"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfTest 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
      Height          =   330
      Left            =   5280
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   330
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "&Cut"
      Height          =   330
      Left            =   5280
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Clip"
      Height          =   330
      Left            =   5280
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   330
      Left            =   5280
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://hothtml3beta.wsoftware.biz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   2640
      MouseIcon       =   "frmMain.frx":007B
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Tag             =   "http://hothtml3beta.wsoftware.biz"
      Top             =   2640
      Width           =   2520
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit http://www.wSoftware.biz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmMain.frx":01CD
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Tag             =   "http://www.wSoftware.biz"
      Top             =   2640
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Paste:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lblClipboard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%lblClipboard%"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/////////////////////////////////////////////////////////////////////
'// WebSoftware SourceCode Archive: Professional Clipboard          //
'/////////////////////////////////////////////////////////////////////
'// Title: Professional Clipboard                                   //
'// Date: 16/03/2003                                                //
'// Author: Thushan Fernando [ thushan@wsoftware.biz ]              //
'// Purpose:                                                        //
'//         This gives your applications the ability to capture the //
'//         WM_DRAWCLIPBOARD message with the help of vbA's awesome //
'//         subclassing control. It would give it a more upper-class//
'//         look to any application that relies on the clipboard!   //
'/////////////////////////////////////////////////////////////////////
'// NOTE: This same/similar code is being used currently in our     //
'//       HotHTML 3 Professional application. I am releasing this   //
'//       code hoping that it helps you become comfortable with the //
'//       more advanced side of VB and to use the API/Subclassing in//
'//       a real life situation.                                    //
'/////////////////////////////////////////////////////////////////////
'// All we ask in return is a link to our website or a mention in   //
'// the credits!                                                    //
'/////////////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////////////
Implements ISubclass
'// This example uses the excellent subclasser by Steve McMahon and //
'// is available on his website: http://www.vbaccelerator.com       //
'/////////////////////////////////////////////////////////////////////
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const WM_DRAWCLIPBOARD = &H308

Private Sub cmdClear_Click()
Clipboard.Clear
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCopy_Click()
rtfCopy rtfTest
End Sub

Private Sub cmdCut_Click()
rtfCut rtfTest
End Sub

Private Sub cmdPaste_Click()
rtfTest.SelText = rtfPaste(rtfTest)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  DetachMessage Me, Me.hWnd, WM_DRAWCLIPBOARD
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set frmMain = Nothing
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
'//
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
   ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim bEnabled As Boolean
   Select Case iMsg
    Case WM_DRAWCLIPBOARD
        If LenB(Clipboard.GetText) = 0 Then bEnabled = False Else bEnabled = True
        lblClipboard.Caption = bEnabled
        cmdPaste.Enabled = bEnabled
   End Select
End Function

Private Sub Form_Load()
AttachMessage Me, Me.hWnd, WM_DRAWCLIPBOARD
SetClipboardViewer Me.hWnd
rtfTest.Text = "This example shows you how to capture the 'WM_DRAWCLIPBOARD' message via Subclassing to determine when the status of the Clipboard has changed." & vbCrLf & vbCrLf & "Try some pasting/copying etc inside this sample then 'Clear' the clipboard so that its empty, now go outside the application(Eg on a webpage) and copy some text, you'll notice the Paste command automatically enabled. Also try clearing the clipboard via another application and see it be disabled!"
End Sub


Private Sub lblLink_Click(Index As Integer)
Call ShellExecute(0&, vbNullString, lblLink(Index).Tag, vbNullString, vbNullString, 1)
End Sub

