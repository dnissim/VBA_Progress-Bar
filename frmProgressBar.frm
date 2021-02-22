VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBar 
   Caption         =   "Updating Spreadsheet"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   OleObjectBlob   =   "frmProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Author:  David Nissim

Private Sub UserForm_Activate()

Me.Caption = "Updating " & ThisWorkbook.Name

End Sub

Public Sub Recaption( _
           Optional strUpperText As String = "*No Update*", _
           Optional strLowerText As String = "*No Update*")

'Changes userform captions without updating progressbar
    
'Don't change captions if nothing was provided.  If a nullstring is provided the caption will go to nullstring
If strUpperText <> "*No Update*" Then txtUpper.Caption = strUpperText
If strLowerText <> "*No Update*" Then txtLower.Caption = strLowerText
Repaint
           
End Sub

Public Sub Progress(ByVal currentItem As Integer, _
                    ByVal totalItems As Integer, _
           Optional strUpperText As String = "*No Update*", _
           Optional strLowerText As String = "*No Update*")
           
' Updates the progress bar length and % Complete caption
Dim percentComplete As Single

percentComplete = currentItem / totalItems * 100
percentComplete = Format(percentComplete, "0")

'Update progress bar visuals
txtBar.Caption = percentComplete & "% Complete"

Recaption strUpperText, strLowerText

Bar.Width = percentComplete * 2
Repaint

DoEvents
End Sub

Public Sub SetBarColor(Optional barColor As Long = &HB917&)
'Change the bar color.  Accepts inputs from the RGB function
'Default is green

frmProgressBar.Bar.BackColor = barColor
frmProgressBar.Repaint
End Sub
