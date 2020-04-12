Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "Kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function ReadSections(ByVal Filename As String) As String()
  Dim szBuf As String, Length As Integer
  Dim SectionArr() As String, m As Integer
  szBuf = String$(255, 0)
  Length = GetPrivateProfileSectionNames(szBuf, 255, Filename)
  szBuf = Left$(szBuf, Length)
  SectionArr = Split(szBuf, vbNullChar)
  ReadSections = SectionArr
End Function
Public Sub ReadKeys(ByVal Section As String, ByVal Filename As String, ByRef xArray() As String)
  Dim Result&, Buffer$
  Dim l%, p%, z%
    Buffer = Space(32767)
    Result = GetPrivateProfileSection(Section, Buffer, Len(Buffer), Filename)
    Buffer = Left$(Buffer, Result)
    If Buffer <> "" Then
      l = 1
      ReDim xArray(0)
      Do While l < Result
        p = InStr(l, Buffer, Chr$(0))
        If p = 0 Then Exit Do
        xArray(z) = Mid$(Buffer, l, p - l)
        z = z + 1
        ReDim Preserve xArray(0 To z)
        l = p + 1
      Loop
    End If
End Sub
Public Function INIRead(ByVal Section As String, ByVal Key As String, ByVal file As String, Optional ByVal default As String) As String
Dim lngResult As Long
Dim strResult As String
    If Len(Trim(default)) = 0 Then default = vbNullString
    strResult = Space(255)
    lngResult = GetPrivateProfileString(Section, Key, default, strResult, 255, file)
    INIRead = Trim(strResult)
End Function
Public Sub IniWrite(ByVal Section As String, ByVal Key As String, ByVal Value As String, ByVal file As String)
Dim lngResult As String
    lngResult = WritePrivateProfileString(Section, Key, Value, file)
End Sub

