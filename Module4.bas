Attribute VB_Name = "Module4"
Sub MostrarIntegrantes()
    Dim integrantes As String
    
    ' Lista de integrantes del proyecto
    integrantes = "Integrantes del Proyecto:" & vbCrLf & _
                  "1. Donmiguel Rosales Edgar" & vbCrLf & _
                  "2. Granados Hern�ndez Oswaldo" & vbCrLf & _
                  "3. Ram�rez de la Crus Daniel Isa�" & vbCrLf & _
                  "4. Santiago Ram�rez Ricardo"
    
    ' Mostrar cuadro de mensaje
    MsgBox integrantes, vbInformation, "Integrantes del Proyecto"
End Sub


