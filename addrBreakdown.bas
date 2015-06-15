Attribute VB_Name = "addrBreakdown"
Sub addrBreakdown()

Dim PICK_WB, PICK_WS, A_CELL, CITYCELL, STATECELL, ZIPCODECELL, COUNTRYCELL As String

'initialize vars

PICK_WS = "Arkansas Firms"
A_CELL = "A2"
CITYCELL = ""
STATECELL = "AR"
ZIPCODECELL = ""
COUNTRYCELL = "USA"

'initialize worksheet
ActiveSheet.Range(A_CELL).Select

'loop through page till empty cell in A column
Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))

    'select address cell
    ActiveCell.Offset(1, 1).Select
    'copy address data
    CITYCELL = ActiveCell.Value
    STATECELL = ActiveCell.Value
    ZIPCODECELL = ActiveCell.Value
    
    'copy country
    ActiveCell.Offset(1, 0).Select
    COUNTRYCELL = ActiveCell.Value
    
    'Move to new address location
    ActiveCell.Offset(-2, 1).Select
    CITYCELL = Trim(CITYCELL)
    CITYCELL = Left(CITYCELL, Len(CITYCELL) - 10)
    ActiveCell.Value = CITYCELL
    
    'Move to new state location
    ActiveCell.Offset(0, 1).Select
    STATECELL = Trim(STATECELL)
    STATECELL = Left(STATECELL, Len(STATECELL) - 6)
    STATECELL = Right(STATECELL, 2)
    ActiveCell.Value = STATECELL
    
    'Move to new zip location
    ActiveCell.Offset(0, 1).Select
    ZIPCODECELL = Trim(ZIPCODECELL)
    ZIPCODECELL = Right(ZIPCODECELL, 5)
    ActiveCell.Value = ZIPCODECELL
    
    'Move to new country location, set to Merica
    ActiveCell.Offset(0, 1).Select
    COUNTRYCELL = Trim(COUNTRYCELL)
    ActiveCell.Value = COUNTRYCELL
    
    'move to delete rows
    ActiveCell.Offset(1, -5).Select
    
    Rows(ActiveCell.Row).Delete Shift:=xlUp
    Rows(ActiveCell.Row).Delete Shift:=xlUp
    
    'move to next entry
    'ActiveCell.Offset(2, 0).Select
    
Loop
End Sub

