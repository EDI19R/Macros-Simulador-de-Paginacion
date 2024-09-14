Attribute VB_Name = "Module1"
Sub IniciarProceso()
    Dim proceso As String
    Dim tamaño As Integer
    Dim i As Integer
    Dim espacioLibre As Integer
    Dim fila As Integer
    Dim procesoEnEspera As Boolean
    
    ' Obtener el nombre del proceso y el tamaño desde las celdas
    proceso = Range("D9").Value
    tamaño = Range("C11").Value
    
    ' Verificar espacio libre en la memoria principal
    espacioLibre = 0
    For i = 8 To 15 ' Asumiendo que tienes 8 marcos de memoria (N8:P15)
        If Cells(i, 14).Value = "" Then
            espacioLibre = espacioLibre + 1
        End If
    Next i
    
    ' Iniciar proceso si hay suficiente espacio en la memoria principal
    If espacioLibre >= tamaño Then
        espacioLibre = 0
        For i = 8 To 15
            If Cells(i, 14).Value = "" Then
                Cells(i, 14).Value = "#"
                Cells(i, 15).Value = "#"
                Cells(i, 16).Value = "#"
                espacioLibre = espacioLibre + 1
                If espacioLibre = tamaño Then Exit For
            End If
        Next i
        
        ' Actualizar tabla de procesos activos
        For fila = 8 To 13
            If Cells(fila, 10).Value = "" Then
                Cells(fila, 10).Value = proceso & tamaño
                Cells(fila, 11).Value = tamaño
                Cells(fila, 12).Value = "En ejecución"
                Exit For
            End If
        Next fila
        
        ' Recalcular las fórmulas en P17 y L5
        Range("P17").Calculate
        Range("L5").Calculate
        
        MsgBox "Proceso " & proceso & " iniciado."
    Else
        ' Verificar espacio en la tabla de procesos en espera
        procesoEnEspera = False
        For fila = 15 To 20
            If Cells(fila, 10).Value = "" Then
                Cells(fila, 10).Value = proceso & tamaño
                Cells(fila, 11).Value = tamaño
                Cells(fila, 12).Value = "En espera"
                procesoEnEspera = True
                Exit For
            End If
        Next fila
        
        If procesoEnEspera Then
            MsgBox "Proceso " & proceso & " en espera."
        Else
            MsgBox "No hay suficiente espacio para el proceso " & proceso & "."
        End If
    End If
End Sub

