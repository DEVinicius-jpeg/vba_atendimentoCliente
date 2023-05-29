VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents toNumero As MSForms.TextBox
Attribute toNumero.VB_VarHelpID = -1
Public WithEvents toData As MSForms.TextBox
Attribute toData.VB_VarHelpID = -1
Public WithEvents toTelefone As MSForms.TextBox
Attribute toTelefone.VB_VarHelpID = -1

Private Sub toData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    With toData
        .MaxLength = 10
        Select Case KeyAscii
           Case 8
           Case 13: SendKeys "{TAB}"
           Case 48 To 57
              If .SelStart = 2 Then .SelText = "/"
              If .SelStart = 5 Then .SelText = "/"
           Case Else: KeyAscii = 0
           
        End Select
        
    End With

End Sub

Private Sub toTelefone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    With toTelefone
        .MaxLength = 14
        Select Case KeyAscii
           Case 8
           Case 13: SendKeys "{TAB}"
           Case 32
           Case 40 To 41
           Case 48 To 57
              If .SelStart = 0 Then .SelText = "("
              If .SelStart = 3 Then .SelText = ")"
              If .SelStart = 4 Then .SelText = " "
           Case Else: KeyAscii = 0
           
        End Select
        
    End With

End Sub

Private Sub toNumero_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    With toNumero
        Select Case KeyAscii
            Case 48 To 57
            Case Else: KeyAscii = 0
            
        End Select
        
    End With
    
End Sub
