VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRestHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_cBaseUrl As String
Private m_iTimeoutForGetMethod As Integer

Public Property Let cBaseUrl(Value As String)
    m_cBaseUrl = Value
End Property

Public Property Get cBaseUrl() As String
    cBaseUrl = m_cBaseUrl
End Property

Public Property Let iTimeoutForGetMethod(Value As Integer)
    m_iTimeoutForGetMethod = Value
End Property

Public Property Get iTimeoutForGetMethod() As Integer
    iTimeoutForGetMethod = m_iTimeoutForGetMethod
End Property
'
'Private Function Execute(cResource As String, _
'                         cMethod As String, _
'                         Optional vBody As Variant = Nothing, _
'                         Optional oQueries As Dictionary = Nothing, _
'                         Optional oPaths As Dictionary = Nothing, _
'                         Optional oHeaders As Dictionary = Nothing) As Dictionary
Private Function Execute(cResource As String, _
                         cMethod As String, _
                         Optional vBody As Variant = Nothing, _
                         Optional oQueries As Dictionary = Nothing, _
                         Optional oPaths As Dictionary = Nothing, _
                         Optional oHeaders As Dictionary = Nothing) As Object
    On Error GoTo ErrorHandler

    Dim cKey As Variant
    Dim cResourceAux As String
    cResourceAux = cResource

    ' Agrega los valores del path.
    If Not oPaths Is Nothing Then
        For Each cKey In oPaths.keys
            Dim cKeyAux As String
            cKeyAux = "{" & cKey & "}"
            cResourceAux = Replace(cResourceAux, cKeyAux, oPaths.Item(cKey))

        Next

    End If

    ' Agrega los parametros de GetVerbString
    Dim cGetVerbString As String
    cGetVerbString = ToQueryString(oQueries)

    Dim cUrl As String
    If cGetVerbString = "" Then
        cUrl = m_cBaseUrl & "/" & cResourceAux

    Else
        cUrl = m_cBaseUrl & "/" & cResourceAux & "?" & cGetVerbString

    End If

    ' Crea la instancia del WinHttp.WinHttpRequest.5.1
    ' Link: https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106(v=vs.85).aspx
    ' Referencia para el proyecto c:\Windows\System32\winhttpcom.dll

    ' Uso dinamico.
    ' Uso dinamico.
     Dim oWinHttp As Object
     Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
     ' Uso tipado.
     'Dim oWinHttp As New  WinHttpRequest ' DA ERROR!!!
        
    
    If cMethod = "GET" _
       And m_iTimeoutForGetMethod > 0 Then
        oWinHttp.SetTimeouts m_iTimeoutForGetMethod, m_iTimeoutForGetMethod, m_iTimeoutForGetMethod, m_iTimeoutForGetMethod

    End If

    ' Open Method, Url, Async
    oWinHttp.Open cMethod, cUrl, False

    If oHeaders Is Nothing Then
        oWinHttp.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        oWinHttp.SetRequestHeader "Accept", "application/json"

    Else
        ' Agrega todas las cabeceras a la petición.
        For Each cKey In oHeaders.keys
            oWinHttp.SetRequestHeader cKey, oHeaders.Item(cKey)

        Next

    End If

    ' Si se envio el body se agrega, requiere que ya este serializado o sea serializable.
    If Not IsObject(vBody) Then
        If vBody <> "" Then
            oWinHttp.Send vBody

        Else
            oWinHttp.Send

        End If
    Else
        oWinHttp.Send

    End If

    If oWinHttp.Status = 204 Then
        Set Execute = Nothing

    Else
        ' Si el servidor devolvio distinto de 20X y el contenido es del tipo JSON lo parsea.
        Set Execute = GetResponseObject(oWinHttp)

    End If

    Exit Function

ErrorHandler:
    ' Set Execute = Nothing
    'Err.Raise Err.number, _
     '          "RestHelper::Execute(" & CStr(Erl) & ")->" & Err.source, _
     '          Err.description, _
     '          Err.helpFile, _
     '          Err.helpContext

    Dim Source As String
    Source = Err.Source

    Dim Description As String
    Description = Trim(Err.Description)

    Dim HelpFile As String
    HelpFile = Trim(Err.HelpFile)

    Dim HelpContext As String
    HelpContext = Trim(Err.HelpContext)

    Dim Number As String
    Number = Err.Number

    Set Execute = TryGetResponseObject(oWinHttp)
    If Execute Is Nothing Then
        Set Execute = New Dictionary
        Execute.Add "Content", CStr(Number) & " " & Description
        Execute.Add "Headers", Null
        Execute.Add "Status", 0
        Execute.Add "StatusText", ""
        
    End If

    Execute.Add "number", Number
    Execute.Add "description", Description
    Execute.Add "source", Source
    Execute.Add "helpFile", HelpFile
    Execute.Add "helpContext", HelpContext
    
    'Execute.Add "VB6ErrorUniqueId", GetGUID()

End Function

Private Function TryGetResponseObject(oWinHttp As Object) As Dictionary
    On Error Resume Next
    If Not oWinHttp Is Nothing Then
        Set TryGetResponseObject = GetResponseObject(oWinHttp)

    End If
End Function

Private Function ParseHeaders(cHeaders As String) As Dictionary
    On Error GoTo ErrorHandler

    Dim pos As Integer
    Dim idx As Integer
    Dim lines() As String
    Dim Line As String
    Dim cKey As String
    Dim cValue As String

    Set ParseHeaders = New Dictionary

    If Trim(cHeaders) <> "" Then
        lines = Split(cHeaders, vbCrLf)

        For idx = LBound(lines) To UBound(lines) - 1
            Line = Trim(lines(idx))
            If Line <> "" Then
                pos = InStr(Line, ":")
                cKey = Mid(Line, 1, pos - 1)
                cValue = Mid(Line, pos + 1)
                If Not ParseHeaders.Exists(Trim(cKey)) Then
                    ParseHeaders.Add Trim(cKey), Trim(cValue)
                
                End If

            End If

        Next

    End If

    Exit Function

ErrorHandler:
    Set ParseHeaders = Nothing
    Err.Raise Err.Number, "RestHelper::ParseHeaders(" & CStr(Erl) & ")->" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'Private Function GetResponseObject(oWinHttp As Object) As Dictionary
Private Function GetResponseObject(oWinHttp As Object) As Object
    On Error GoTo ErrorHandler

    Dim cHeaders As String
    cHeaders = oWinHttp.GetAllResponseHeaders()

    Dim oHeaders As Dictionary
    Set oHeaders = ParseHeaders(cHeaders)

    If oHeaders.Exists("Content-Type") Then
        Dim cContentType As String
        cContentType = oHeaders.Item("Content-Type")
        If cContentType Like "*application/json*" Then
            Set GetResponseObject = basJSON.parse(oWinHttp.responseText)

        Else
            Set GetResponseObject = GetResponseObjectDefault(oWinHttp, oHeaders)

        End If

    Else
        Set GetResponseObject = GetResponseObjectDefault(oWinHttp, oHeaders)

    End If

    Exit Function

ErrorHandler:
    Set GetResponseObject = Nothing
    Err.Raise Err.Number, _
              "RestHelper::GetResponseObject(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Private Function GetResponseObjectDefault(oWinHttp As Object, _
                                          oHeaders As Dictionary) As Dictionary
    On Error GoTo ErrorHandler

    Set GetResponseObjectDefault = New Dictionary
    GetResponseObjectDefault.Add "Content", oWinHttp.responseText
    GetResponseObjectDefault.Add "Headers", oHeaders
    GetResponseObjectDefault.Add "Status", oWinHttp.Status
    GetResponseObjectDefault.Add "StatusText", oWinHttp.StatusText

    Exit Function

ErrorHandler:
    Set GetResponseObjectDefault = Nothing
    Err.Raise Err.Number, _
              "RestHelper::GetResponseObjectDefault(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Private Function ExecuteWithBody(cResource As String, _
                                 cMethod As String, _
                                 Optional vBody As Variant = Nothing, _
                                 Optional oQueries As Dictionary = Nothing, _
                                 Optional oPaths As Dictionary = Nothing, _
                                 Optional oHeaders As Dictionary = Nothing) As Object
    On Error GoTo ErrorHandler

    Dim lvBody As Variant
    Dim cTypeName As String

    ' Convierte lo que se envia por Body a string.
    Select Case VarType(vBody)

        Case vbEmpty

        Case vbNull

        Case vbInteger
            lvBody = CStr(vBody)

        Case vbLong
            lvBody = CStr(vBody)

        Case vbSingle
            lvBody = CStr(vBody)

        Case vbDouble
            lvBody = CStr(vBody)

        Case vbCurrency
            lvBody = CStr(vBody)

        Case vbDate
            lvBody = CStr(vBody)

        Case vbString
            lvBody = vBody

        Case vbObject
            cTypeName = TypeName(vBody)
            If cTypeName = "Dictionary" Or cTypeName = "Collection" Then
                lvBody = basJSON.ToString(vBody)

            Else
                Err.Raise vbObjectError + 1, Err.Source, "Tipo de dato no soportado. No se puede serializar el tipo " & cTypeName & " a formato JSON, Por favor utilice Dictionary o Collection"

            End If

        Case vbError

        Case vbBoolean
            lvBody = CStr(vBody)

        Case vbVariant
            lvBody = CStr(vBody)

        Case vbDataObject

        Case vbDecimal
            lvBody = CStr(vBody)

        Case vbByte
            lvBody = CStr(vBody)

        Case vbUserDefinedType

        Case vbArray

    End Select

    Set ExecuteWithBody = Execute(cResource, cMethod, lvBody, oQueries, oPaths, oHeaders)

    Exit Function

ErrorHandler:
    Set ExecuteWithBody = Nothing
    Err.Raise Err.Number, _
              "RestHelper::ExecuteWithBody(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function PostVerb(cResource As String, _
                         Optional vBody As Variant = Nothing, _
                         Optional oQueries As Dictionary = Nothing, _
                         Optional oPaths As Dictionary = Nothing, _
                         Optional oHeaders As Dictionary = Nothing) As Object
    On Error GoTo ErrorHandler
    Set PostVerb = ExecuteWithBody(cResource, "POST", vBody, oQueries, oPaths, oHeaders)

    Exit Function

ErrorHandler:
    Set PostVerb = Nothing
    Err.Raise Err.Number, _
              "RestHelper::PostVerb(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function PutVerb(cResource As String, _
                           Optional vBody As Variant = Nothing, _
                           Optional oQueries As Dictionary = Nothing, _
                           Optional oPaths As Dictionary = Nothing, _
                           Optional oHeaders As Dictionary = Nothing) As Object
    On Error GoTo ErrorHandler
    Set PutVerb = ExecuteWithBody(cResource, "PUT", vBody, oQueries, oPaths, oHeaders)

    Exit Function

ErrorHandler:
    Set PutVerb = Nothing
    Err.Raise Err.Number, _
              "RestHelper::PutVerb(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function GetVerb(cResource As String, _
                        Optional vBody As Variant = Nothing, _
                        Optional oQueries As Dictionary = Nothing, _
                        Optional oPaths As Dictionary = Nothing, _
                        Optional oHeaders As Dictionary = Nothing) As Object
    On Error GoTo ErrorHandler
    Set GetVerb = Execute(cResource, "GET", Nothing, oQueries, oPaths, oHeaders)

    Exit Function

ErrorHandler:
    Set GetVerb = Nothing
    Err.Raise Err.Number, _
              "RestHelper::GetVerb(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function DeleteVerb(cResource As String, _
                           Optional oQueries As Dictionary = Nothing, _
                           Optional oPaths As Dictionary = Nothing, _
                           Optional oHeaders As Dictionary = Nothing) As Object
    On Error GoTo ErrorHandler
    Set DeleteVerb = Execute(cResource, "DeleteVerb", Nothing, oQueries, oPaths, oHeaders)

    Exit Function

ErrorHandler:
    Set DeleteVerb = Nothing
    Err.Raise Err.Number, _
              "RestHelper::DeleteVerb(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function


