VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRestHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_data As New Dictionary

Public Function ContentType(cValue As String) As clsRestHeaders
    On Error GoTo ErrorHandler

    AddKey "Content-Type", cValue

    Set ContentType = Me

    Exit Function

ErrorHandler:
    Set ContentType = Nothing
    Err.Raise Err.Number, _
              "Headers::ContentType(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function Accept(cValue As String) As clsRestHeaders
    On Error GoTo ErrorHandler

    AddKey "Accept", cValue

    Set Accept = Me

    Exit Function

ErrorHandler:
    Set Accept = Nothing
    Err.Raise Err.Number, _
              "Headers::Accept(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function AcceptJson() As clsRestHeaders
    On Error GoTo ErrorHandler

    Set AcceptJson = Accept("application/json")

    Exit Function

ErrorHandler:
    Set AcceptJson = Nothing
    Err.Raise Err.Number, _
              "Headers::AcceptJson(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function ContentTypeJson() As clsRestHeaders
    On Error GoTo ErrorHandler

    Set ContentTypeJson = ContentType("application/json")

    Exit Function

ErrorHandler:
    Set ContentTypeJson = Nothing
    Err.Raise Err.Number, _
              "Headers::ContentTypeJson(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function Custom(cKey As String, cValue As String) As clsRestHeaders
    On Error GoTo ErrorHandler
    AddKey cKey, cValue
    Set Custom = Me

    Exit Function

ErrorHandler:
    Set Custom = Nothing
    Err.Raise Err.Number, _
              "Headers::Custom(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function Clear() As clsRestHeaders
    On Error GoTo ErrorHandler
    m_data.RemoveAll
    Set Clear = Me

    Exit Function

ErrorHandler:
    Set Clear = Nothing
    Err.Raise Err.Number, _
              "Headers::Clear(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Private Sub AddKey(cKey As String, cValue As String, Optional bReplace = True)
    On Error GoTo ErrorHandler
    If m_data.Exists(cKey) And bReplace Then
        m_data.Remove (cKey)

    End If

    m_data.Add cKey, cValue

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, _
              "Headers::AddKey(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Sub

Public Function ToDictionary() As Dictionary
    On Error GoTo ErrorHandler
    Dim vKey As Variant
    Set ToDictionary = New Dictionary
    For Each vKey In m_data.keys
        ToDictionary.Add vKey, m_data.Item(vKey)

    Next

    Exit Function

ErrorHandler:
    Set ToDictionary = Nothing
    Err.Raise Err.Number, _
              "Headers::ToDictionary(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function ToJson() As String
    On Error GoTo ErrorHandler
    Dim dict As Dictionary
    Set dict = ToDictionary()
    ToJson = basJSON.ToString(dict)

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, _
              "Headers::ToJson(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function



