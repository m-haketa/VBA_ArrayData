Attribute VB_Name = "Module1"
Option Explicit

Sub “®‚­()
  Dim AD As ArrayData
  Set AD = New ArrayData
  
  Range("A1") = 1
  Range("B1") = 2
  Range("C1") = 3
  
  AD.Data = Range("A1:C1").Value
 
  Range("A2:C2").Value = AD.Data
End Sub

Sub “®‚©‚È‚¢()
  Dim AD As ArrayData
  Set AD = New ArrayData
  
  Range("A1") = 1
  Range("B1") = 2
  Range("C1") = 3
  
  AD = Range("A1:C1").Value

'‰º‹Ls‚ª‚¤‚Ü‚­“®‚©‚È‚¢
  Range("A3:C3").Value = AD
End Sub
