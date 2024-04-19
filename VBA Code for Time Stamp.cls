VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    Dim DateColumn As Range
    Dim ModifiedRow As Long

    ' Define the target range to track changes (columns A through N)
    Set KeyCells = Me.Range("A1:N1000") ' Adjust the range as needed

    ' Define the column where you want to record the date (column O in this example)
    Set DateColumn = Me.Range("O1:O1000") ' Adjust the range as needed

    ' If edited cell is within the target range
    If Not Application.Intersect(KeyCells, Target) Is Nothing Then
        ' Get the row of the modified cell
        ModifiedRow = Target.Row

        ' Place the current date in the corresponding cell of the date column
        DateColumn.Cells(ModifiedRow, 1).Value = Date
    End If
End Sub
