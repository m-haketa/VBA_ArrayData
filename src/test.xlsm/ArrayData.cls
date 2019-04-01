VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Data As Variant

Public Property Get Data() As Variant
Attribute Data.VB_UserMemId = 0
'=Default Property
  Data = m_Data
End Property

Public Property Let Data(Value As Variant)
  m_Data = Value
End Property





