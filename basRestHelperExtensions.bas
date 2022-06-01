Attribute VB_Name = "basRestHelperExtensions"
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MAX_PATH                   As Long = 260
Private Const ERROR_SUCCESS              As Long = 0

'Treat entire URL param as one URL segment
Private Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
Private Const URL_ESCAPE_PERCENT         As Long = &H1000
Private Const URL_UNESCAPE_INPLACE       As Long = &H100000

'escape #'s in paths
Private Const URL_INTERNAL_PATH          As Long = &H800000
Private Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
Private Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
Private Const URL_DONT_SIMPLIFY          As Long = &H8000000

'Converts unsafe characters,
'such as spaces, into their
'corresponding escape sequences.
Private Declare Function UrlEscape Lib "shlwapi" _
   Alias "UrlEscapeA" _
  (ByVal pszURL As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long
   
Private m_SafeChar(0 To 255) As Boolean

Private Declare Sub CopyToMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Enum ResponseType
    StatusMessage
    UnHandled
    Expected

End Enum

Public Function GetRestHelper(cBaseUrl As String, _
                              Optional iTimeoutForGetMethod As Integer = 30000) As clsRestHelper
    Set GetRestHelper = New clsRestHelper
    GetRestHelper.cBaseUrl = cBaseUrl
    GetRestHelper.iTimeoutForGetMethod = iTimeoutForGetMethod

End Function

Public Function GetHeaders() As clsRestHeaders
    Set GetHeaders = New clsRestHeaders

End Function

Public Function GetQueryString() As clsRestQueryString
    Set GetQueryString = New clsRestQueryString

End Function

Public Function GetForm() As clsRestQueryString
    Set GetForm = New clsRestQueryString

End Function

Public Function GetPath() As clsRestPath
    Set GetPath = New clsRestPath

End Function

Public Function GetResponseType(Value As Object) As ResponseType
    On Error GoTo ErrorHandler

    GetResponseType = UnHandled
    If IsObject(Value) Then
        If Not Value Is Nothing Then
            If IsStatusMessage(Value) Then
                GetResponseType = StatusMessage

            ElseIf IsUnHandled(Value) Then
                GetResponseType = UnHandled

            Else
                GetResponseType = Expected

            End If

        End If

    End If

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, _
              "RestHelperExtensions::GetResponseType(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

Public Function ToQueryString(Value As Dictionary, Optional Encode As Boolean = True) As String
    On Error GoTo ErrorHandler

    If Not Value Is Nothing Then
        Dim SB As New clsStringBuilder
        Dim vKey As Variant
        For Each vKey In Value.keys
            SB.Append CStr(vKey)

            If Not IsEmpty(Value.Item(vKey)) Then
                SB.Append "="
                
                If Encode Then
                    SB.Append URLEncode(Value.Item(vKey))
            
                Else
                    SB.Append Value.Item(vKey)
        
                End If

            End If

            SB.Append "&"

        Next
        
        Dim ret As String
        
        ret = SB.ToString
        ret = Left(ret, Len(ret) - 1)
        ToQueryString = ret
        
    End If

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, _
              "RestHelperExtensions::ToQueryString(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

'Private Function IsStatusMessage(Value As Dictionary) As Boolean
Private Function IsStatusMessage(Value As Object) As Boolean
    On Error GoTo ErrorHandler

    If TypeOf Value Is Dictionary Then
        IsStatusMessage = Value.Exists("errorCode") _
                      And Value.Exists("errorUniqueId")
                      
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, _
              "RestHelperExtensions::IsStatusMessage(" & CStr(Erl) & ")->" & Err.Source, _
              Err.Description, _
              Err.HelpFile, _
              Err.HelpContext

End Function

'Private Function IsUnHandled(Value As Dictionary) As Boolean
Private Function IsUnHandled(Value As Object) As Boolean
    On Error GoTo ErrorHandler
    IsUnHandled = Value.Exists("Content") _
                  And Value.Exists("Headers") _
                  And Value.Exists("Status")
    Exit Function

ErrorHandler:
'    Err.Raise Err.Number, _
'              "RestHelperExtensions::IsUnHandled(" & CStr(Erl) & ")->" & Err.Source, _
'              Err.Description, _
'              Err.HelpFile, _
'              Err.HelpContext

End Function

'Public Function IsUnHandledStatus200(oResponse As Dictionary) As Boolean
Public Function IsUnHandledStatus200(oResponse As Object) As Boolean
On Error GoTo ErrorHandler
    IsUnHandledStatus200 = False
    
    If oResponse.Exists("Status") Then
        If oResponse.Item("Status") = 200 Then
            IsUnHandledStatus200 = True
        End If
    End If
    
 Exit Function

ErrorHandler:
End Function

Public Function ExistsValidationErrorsItem(oError As Dictionary, sItem As String) As Boolean
    ExistsValidationErrorsItem = False
                
    If oError.Exists("validationErrors") Then
        If oError.Item("validationErrors").Exists(sItem) Then
            ExistsValidationErrorsItem = True
        End If
    End If
End Function

' Return a URL safe encoding of txt.
Private Function URLEncode(ByVal Value As String) As String

' https://www.ryadel.com/en/urlencode-utf8-visual-basic-6-vb6/

    Dim Index1 As Long
    Dim Index2 As Long
    Dim Result As String
    Dim Chars() As Byte
    Dim Char As String
    Dim Byte1 As Byte
    Dim Byte2 As Byte
    Dim UTF16 As Long
    Dim CharCode As Long
    Dim Space As String

Space = "+"

For Index1 = 1 To Len(Value)
  CopyToMemory Byte1, ByVal StrPtr(Value) + ((Index1 - 1) * 2), 1
  CopyToMemory Byte2, ByVal StrPtr(Value) + ((Index1 - 1) * 2) + 1, 1

  UTF16 = Byte2
  UTF16 = UTF16 * 256 + Byte1
  Chars = GetUTF8FromUTF16(UTF16)
  For Index2 = LBound(Chars) To UBound(Chars)
     Char = Chr(Chars(Index2))
     
     
'     If Char Like "[0-9A-Za-z]" Then
'        Result = Result & Char
'     Else
'        Result = Result & "%" & Hex(Asc(Char))
'     End If
CharCode = Chars(Index2)
 If 97 <= CharCode And CharCode <= 122 _
    Or 64 <= CharCode And CharCode <= 90 _
    Or 48 <= CharCode And CharCode <= 57 _
    Or 44 = CharCode _
    Or 45 = CharCode _
    Or 46 = CharCode _
    Or 95 = CharCode _
    Or 126 = CharCode Then
      Result = Result & Char
    ElseIf 32 = CharCode Then
      Result = Result & Space
    Else
      ' Result = Result & "&#" & CharCode & ";"
      Result = Result & "%" & Right("0" & Hex(CharCode), 2)

    End If
  Next
Next

URLEncode = Result

' Dim i, CharCode, Char, Space
'  Dim StringLen
'
'  StringLen = Len(value)
'  ReDim Result(StringLen)
'
'  Space = "+"
'  'Space = "%20"
'
'  For i = 1 To StringLen
'    Char = Mid(value, i, 1)
'    CharCode = AscW(Char)
'    If 97 <= CharCode And CharCode <= 122 _
'    Or 64 <= CharCode And CharCode <= 90 _
'    Or 48 <= CharCode And CharCode <= 57 _
'    Or 45 = CharCode _
'    Or 46 = CharCode _
'    Or 95 = CharCode _
'    Or 126 = CharCode Then
'      Result(i) = Char
'    ElseIf 32 = CharCode Then
'      Result(i) = Space
'    Else
'      ' result(i) = "&#" & CharCode & ";"
'      Result(i) = "%" & Right("0" & Hex(CharCode), 2)
'
'    End If
'  Next
'  URLEncode = Join(Result, "")


'Dim i As Integer
'Dim ch As String
'Dim ch_asc As Integer
'Dim result As String
'
'    SetSafeChars
'
'    result = ""
'    For i = 1 To Len(value)
'        ' Translate the next character.
'        ch = Mid$(value, i, 1)
'        ch_asc = Asc(ch)
'        If ch_asc = vbKeySpace Then
'            ' Use a plus.
'            result = result & "+"
'        ElseIf m_SafeChar(ch_asc) Then
'            ' Use the character.
'            result = result & ch
'        Else
'            ' Convert the character to hex.
'            result = result & "%" & Right$("0" & _
'                Hex$(ch_asc), 2)
'        End If
'    Next i
'
'    URLEncode = result

' Dim buff As String
'   Dim dwSize As Long
'   Dim dwFlags As Long
'
'   If Len(value) > 0 Then
'
'      buff = Space$(MAX_PATH)
'      dwSize = Len(buff)
'      dwFlags = URL_DONT_SIMPLIFY
'
'      If UrlEscape(value, _
'                   buff, _
'                   dwSize, _
'                   dwFlags) = ERROR_SUCCESS Then
'
'         URLEncode = Left$(buff, dwSize)
'
'      End If  'UrlEscape
'   End If  'Len(sUrl)
End Function


Private Function GetUTF8FromUTF16(ByVal UTF16 As Long) As Byte()
 
   Dim Result() As Byte
   If UTF16 < &H80 Then
      ReDim Result(0 To 0)
      Result(0) = UTF16
   
   ElseIf UTF16 < &H800 Then
      ReDim Result(0 To 1)
      Result(1) = &H80 + (UTF16 And &H3F)
      UTF16 = UTF16 \ &H40
      Result(0) = &HC0 + (UTF16 And &H1F)
   
   Else
      ReDim Result(0 To 2)
      Result(2) = &H80 + (UTF16 And &H3F)
      UTF16 = UTF16 \ &H40
      Result(1) = &H80 + (UTF16 And &H3F)
      UTF16 = UTF16 \ &H40
      Result(0) = &HE0 + (UTF16 And &HF)
   
   End If
   
   GetUTF8FromUTF16 = Result
   
End Function


' Set m_SafeChar(i) = True for characters that
' do not need protection.
Private Sub SetSafeChars()
Static done_before As Boolean
Dim i As Integer

    If done_before Then Exit Sub
    done_before = True

    For i = 0 To 47
        m_SafeChar(i) = False
    Next i
    For i = 48 To 57
        m_SafeChar(i) = True
    Next i
    For i = 58 To 64
        m_SafeChar(i) = False
    Next i
    For i = 65 To 90
        m_SafeChar(i) = True
    Next i
    For i = 91 To 96
        m_SafeChar(i) = False
    Next i
    For i = 97 To 122
        m_SafeChar(i) = True
    Next i
    For i = 123 To 255
        m_SafeChar(i) = False
    Next i
End Sub

'
'Public Function GetGUID() As String
''(c) 2000 Gus Molina
'
'    Dim udtGUID As Guid
'
'    If (CoCreateGuid(udtGUID) = 0) Then
'
'        GetGUID = _
'        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
'        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
'        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
'        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
'        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
'        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
'        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
'        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
'        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
'        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
'        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
'    Else
'        GetGUID = "?"
'    End If
'
'End Function

