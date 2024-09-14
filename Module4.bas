Attribute VB_Name = "Module4"
Sub MostrarIntegrantes()
    Dim integrantes As String
    
    ' Lista de integrantes del proyecto
    integrantes = "Integrantes del Proyecto:" & vbCrLf & _
                  "1. Donmiguel Rosales Edgar" & vbCrLf & _
                  "2. Granados Hernández Oswaldo" & vbCrLf & _
                  "3. Ramírez de la Crus Daniel Isaí" & vbCrLf & _
                  "4. Santiago Ramírez Ricardo"
    
    ' Mostrar cuadro de mensaje
    MsgBox integrantes, vbInformation, "Integrantes del Proyecto"
End Sub


