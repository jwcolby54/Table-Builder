Attribute VB_Name = "basTools"
Option Compare Database
Option Explicit

#Const DebugPrint = True
#Const VBIDE = True

Public Sub assDebugPrint(ByVal strMsg As String, Optional boolPrint As Boolean = True)
#If DebugPrint Then
    If boolPrint = True Then Debug.Print strMsg
#End If
End Sub

'
'Returns a long integer random number between lngUpperBound and lngLowerBound
'
Function Random(lngUpperBound As Long, lngLowerBound As Long) As Long
    Random = Int((lngUpperBound - lngLowerBound + 1) * Rnd + lngLowerBound)
End Function

Public Sub RefreshNavigationPane()
    ' Collapse the Navigation Pane
    DoCmd.SelectObject acTable, , True
    DoCmd.RunCommand acCmdWindowHide

    ' Optional pause to ensure the pane collapses before reopening
    DoEvents

    ' Reopen the Navigation Pane
    DoCmd.RunCommand acCmdWindowUnhide
End Sub
'
'In order to use the VBIDE you must reference it.
'
'In the VBA editor, go to Tools > References.
'Check Microsoft Visual Basic for Applications Extensibility 5.3.
'
'SendKeys can be a bit finicky. Make sure the VBE is visible,
'and the Immediate Window is open when you run this.
'
'Sub ClearImmediateWindow()
'#If VBIDE Then
'    Dim VBEEditor As VBIDE.VBE
'    Set VBEEditor = Application.VBE
'    '
'    'Open the immediate window programmatically
'    'if it is closed
'    '
'    Application.VBE.CommandBars.FindControl(ID:=578).Execute ' Opens Immediate Window
'    '
'    'Set the focus into the immediate window
'    VBEEditor.Windows("Immediate").SetFocus
'
'    '
'    'Select everything in the window and delete it
'    SendKeys "^a"
'    SendKeys "{DEL}"
'#End If
'End Sub
Sub ClearImmediateWindow()
#If VBIDE Then
#If VBA7 Then
    Dim vbeWindow As VBIDE.Window
    On Error Resume Next
    Set vbeWindow = Application.VBE.Windows("Immediate")
    If Not vbeWindow Is Nothing Then
        vbeWindow.SetFocus
        SendKeys "^a", True
        SendKeys "{DEL}", True
    Else
        MsgBox "Immediate Window not available.", vbExclamation
    End If
    On Error GoTo 0
#End If
#End If
End Sub


