Attribute VB_Name = "Modu1_PT"

Function JourneyHead(Stops As Range, Headway As Integer, Comp As Double)
    Dim comp_r As Double
    Dim first_c As Integer
    Dim head As Single
    Dim Line As Integer
    Dim Row As Integer
    
    If Stops.Row = 2 Then 'first row
        JourneyHead = Stops.Cells(1, Headway)
        Exit Function
    End If
    
    head = Stops.Item(1, Headway)
    first_c = Stops.Column
    Line = Stops.Item(1, 1)
    comp_r = Stops.Cells(0, Comp)

    If head = 1 Then 'head == 1
        JourneyHead = head
        Exit Function
    ElseIf head = comp_r Then 'head == row before
        JourneyHead = head
        Exit Function
    ElseIf Stops.Cells(2, 1) <> Line Then 'last row of line
        JourneyHead = head
        Exit Function
    End If
         
    Row = 0
    For Each Item In Stops
    
        If Item.Column <> first_c Then
            GoTo NextIteration
        End If
        
        Row = Row + 1
        
        If Stops.Cells(Row + 1, 1) <> Line Then 'reaching last row of Line
            JourneyHead = head
            Exit Function
        ElseIf Item.Columns(Headway) = 1 Then 'reaching an new block with headway == 1
            JourneyHead = Item.Columns(Headway)
            Exit Function
        End If
        

NextIteration:
    Next Item
    
    JourneyHead = head
    
End Function

