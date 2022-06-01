VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRestPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_data As New Dictionary

Public Function Add(cKey As String, cValue As String, Optional bReplace = True) As clsRestPath
    On Error GoTo ErrorHandler
    If m_data.Exists(cKey) And bReplace Then
        m_data.Remove (cKey)

    End If

    m_data.Add cKey, cValue

    Set Add = Me

    Exit Function

ErrorHandler:
    Set Add = Nothing
    Err.Raise Err.Number, _
              "Path::Add(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

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
              "Path::ToDictionary(" & CStr(Erl) & ")->" & Err.Source, _
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
              "Path::ToJson(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function




