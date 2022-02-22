VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker()
For i = 0 To 271
Cells(i + 2, 8).Value = Cells(262 * i + 2, 1).Value
Next i
End Sub

