Attribute VB_Name = "Module3"
Sub LimpiarTodo()
    Dim i As Integer
    Dim fila As Integer
    
    ' Limpiar tabla de procesos activos
    For fila = 8 To 13
        Cells(fila, 10).Value = ""
        Cells(fila, 11).Value = ""
        Cells(fila, 12).Value = ""
    Next fila
    
    ' Limpiar tabla de procesos en espera
    For fila = 15 To 20
        Cells(fila, 10).Value = ""
        Cells(fila, 11).Value = ""
        Cells(fila, 12).Value = ""
    Next fila
    
    ' Limpiar tabla de páginas ocupadas
    For i = 8 To 15
        Cells(i, 14).Value = ""
        Cells(i, 15).Value = ""
        Cells(i, 16).Value = ""
    Next i
    
    ' Recalcular las fórmulas en P17 y L5
    Range("P17").Calculate
    Range("L5").Calculate
    
    MsgBox "Todas las tablas han sido limpiadas."
End Sub

