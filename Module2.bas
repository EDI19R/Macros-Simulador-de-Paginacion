Attribute VB_Name = "Module2"
Sub TerminarProceso()
    Dim proceso As String
    Dim tama�o As Integer
    Dim i As Integer
    Dim fila As Integer
    
    ' Obtener el nombre del proceso desde la celda
    proceso = Range("D9").Value
    
    ' Encontrar el tama�o del proceso en la tabla de procesos activos
    For fila = 8 To 13
        If Cells(fila, 10).Value Like proceso & "*" Then
            tama�o = Cells(fila, 11).Value
            Exit For
        End If
    Next fila
    
    ' Terminar proceso y liberar p�ginas ocupadas
    Dim paginasLiberadas As Integer
    paginasLiberadas = 0
    For i = 8 To 15 ' Asumiendo que tienes 8 marcos de memoria (N8:P15)
        If Cells(i, 14).Value = "#" And Cells(i, 15).Value = "#" And Cells(i, 16).Value = "#" Then
            Cells(i, 14).Value = ""
            Cells(i, 15).Value = ""
            Cells(i, 16).Value = ""
            paginasLiberadas = paginasLiberadas + 1
            If paginasLiberadas = tama�o Then Exit For
        End If
    Next i
    
    ' Actualizar tabla de procesos activos
    For fila = 8 To 13
        If Cells(fila, 10).Value Like proceso & "*" Then
            Cells(fila, 10).Value = ""
            Cells(fila, 11).Value = ""
            Cells(fila, 12).Value = ""
            Exit For
        End If
    Next fila
    
    ' Recalcular las f�rmulas en P17 y L5
    Range("P17").Calculate
    Range("L5").Calculate
    
    MsgBox "Proceso " & proceso & " terminado y p�ginas liberadas."
    
    ' Verificar si hay procesos en espera que puedan ser movidos a la memoria principal
    For fila = 15 To 20
        If Cells(fila, 10).Value <> "" Then
            proceso = Left(Cells(fila, 10).Value, Len(Cells(fila, 10).Value) - Len(CStr(Cells(fila, 11).Value)))
            tama�o = Cells(fila, 11).Value
            
            ' Verificar espacio libre en la memoria principal
            espacioLibre = 0
            For i = 8 To 15
                If Cells(i, 14).Value = "" Then
                    espacioLibre = espacioLibre + 1
                End If
            Next i
            
            If espacioLibre >= tama�o Then
                espacioLibre = 0
                For i = 8 To 15
                    If Cells(i, 14).Value = "" Then
                        Cells(i, 14).Value = "#"
                        Cells(i, 15).Value = "#"
                        Cells(i, 16).Value = "#"
                        espacioLibre = espacioLibre + 1
                        If espacioLibre = tama�o Then Exit For
                    End If
                Next i
                
                ' Actualizar tabla de procesos activos
                For filaActiva = 8 To 13
                    If Cells(filaActiva, 10).Value = "" Then
                        Cells(filaActiva, 10).Value = proceso & tama�o
                        Cells(filaActiva, 11).Value = tama�o
                        Cells(filaActiva, 12).Value = "En ejecuci�n"
                        Exit For
                    End If
                Next filaActiva
                
                ' Limpiar proceso de la tabla de procesos en espera
                Cells(fila, 10).Value = ""
                Cells(fila, 11).Value = ""
                Cells(fila, 12).Value = ""
                
                ' Recalcular las f�rmulas en P17 y L5
                Range("P17").Calculate
                Range("L5").Calculate
                
                MsgBox "Proceso " & proceso & " movido de espera a ejecuci�n."
                Exit For
            End If
        End If
    Next fila
End Sub

