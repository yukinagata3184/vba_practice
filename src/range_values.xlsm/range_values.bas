Attribute VB_Name = "range_values"
Option Explicit

Sub range_values()
    Range("A1:D3").ClearContents
    'practice Range("Cell num:Cell num")
    Range("A1").Value = "profit"
    Range("A2:D3").Value = 3184
    'practice ClearContents
    Range("D1:D3").ClearContents
End Sub
