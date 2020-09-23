Attribute VB_Name = "rtfFunctions"
Option Explicit
'/////////////////////////////////////////////////////////////////////
'// WebSoftware SourceCode Archive: Microsoft Visual Basic Code     //
'//             Visit http://www.wSoftware.biz                      //
'/////////////////////////////////////////////////////////////////////
'// Title: rtfFunctions                                             //
'// Date: 02/07/2000                                                //
'// Last Modified: 16/03/2003                                       //
'// Author: Thushan Fernando [ thushan@wsoftware.biz ]              //
'// Purpose:                                                        //
'//         Useful Richtextbox functions wrapped into a handy       //
'//         module which can be placed into any project that has a  //
'//         richtextbox control. This was originally written for my //
'//         HotHTML application.                                    //
'/////////////////////////////////////////////////////////////////////
'// NOTE: This is a scaled down/cut down version of this bas file,  //
'//       used for the example code only. The full bas file will be //
'//       updated and uploaded when time is available.              //
'/////////////////////////////////////////////////////////////////////

Private Const WM_COPY = &H301
Private Const WM_CUT = &H300

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub rtfCut(rtfControl As RichTextBox)
SendMessage rtfControl.hWnd, WM_CUT, 0&, 0&
End Sub
Public Sub rtfCopy(rtfControl As RichTextBox)
SendMessage rtfControl.hWnd, WM_COPY, 0&, 0&
End Sub
Public Function rtfPaste(rtfControl As RichTextBox) As String
    Dim strClipboard As String, lngStart As Long
    lngStart = rtfControl.SelStart
    strClipboard = Clipboard.GetText(vbCFText)
    rtfControl.SelText = strClipboard
    rtfControl.SetFocus
End Function
