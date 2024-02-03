Attribute VB_Name = "range_value_clear"
Sub range_value_clear()
    'Substitutions value to a cell
    Range("A1").Value = 3184
    Range("A2").Value = "Yuki"
    'Copy a value to another cell
    Range("B1").Value = Range("A1").Value
    Range("B2").Value = Range("A2").Value
    'Clear the value of the cell
    Range("B2").ClearContents
End Sub

