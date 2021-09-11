Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function apiBeginUpdateResourceA Lib "kernel32.dll" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function apiCopyMemoryByteLong Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef pDst As Byte, ByVal pSrc As Long, ByVal ByteLen As Long) As Long
Private Declare Function apiEndUpdateResourceA Lib "kernel32.dll" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Private Declare Function apiFindResourceALong Lib "kernel32.dll" Alias "FindResourceA" (ByVal hLib As Long, ByVal lpName As String, ByVal lpType As Long) As Long
Private Declare Function apiFreeResource Lib "kernel32.dll" Alias "FreeResource" (ByVal hResData As Long) As Long
Private Declare Function apiFreeLibrary Lib "kernel32.dll" Alias "FreeLibrary" (ByVal hLib As Long) As Long
Private Declare Function apiLoadLibraryA Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal strFilePath As String) As Long
Private Declare Function apiLoadResource Lib "kernel32.dll" Alias "LoadResource" (ByVal hLib As Long, ByVal hRes As Long) As Long
Private Declare Function apiLockResource Lib "kernel32.dll" Alias "LockResource" (ByVal hRes As Long) As Long
Private Declare Function apilstrlenA Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function apiSizeofResource Lib "kernel32.dll" Alias "SizeofResource" (ByVal hModule As Long, ByVal hResInfo As Long) As Long
Private Declare Function apiUpdateResourceA Lib "kernel32.dll" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Sub Main()
   On Error Resume Next
   Dim sn As String
   Dim ser As String
   ser = Command
   ser = Replace(ser, Chr(34), "") 'strip any quotations etc.
   ser = Replace(ser, "-", "")
   ser = Replace(ser, " ", "")
   If IsNumeric(ser) = False Then  'make sure serial number is being passed
      Exit Sub
   End If
   If Len(ser) <> 20 Then
      ser = GetRandomNumber & ser & GetRandomNumber
   End If
   If Len(ser) <> 20 Then  'make sure 20 digit serial number is being passed"Pass a 20 digit numeric serial number as an argument, ie:" & vbCrLf & "12345-123-1234567-12345"
      Exit Sub
   End If
   Dim hLibrary As Long
   Dim pth As String
   pth = App.Path & "\VB6.exe"
   hLibrary = apiLoadLibraryA(pth)
   If hLibrary <> 0 Then
      Dim nams As String
      Dim lDataSize As Long
      Dim hresource As Long
      hresource = apiFindResourceALong(hLibrary, "#196", 106)
      If hresource <> 0 Then
         Dim lpData As Long
         Dim hData As Long
         hData = apiLoadResource(hLibrary, hresource)
         If hData <> 0 Then
            lpData = apiLockResource(hData)    'pointer to data
            If lpData <> 0 Then
               lDataSize = apiSizeofResource(hLibrary, hresource)
               If lDataSize <> 0 Then
                  Dim buff() As Byte
                  ReDim buff(0 To (lDataSize - 1)) As Byte
                  apiCopyMemoryByteLong buff(0), lpData, lDataSize
                  nams = VBA.StrConv(buff, vbUnicode) 'get owner's name and company name (if any already specified)
                  sn = VBA.Right(nams, 21) 'extract 20 digit serial number + Chr(0)
                  nams = VBA.Left(nams, apilstrlenA(lpData))
               End If
            End If
            apiFreeResource hData
         End If
      End If
      apiFreeLibrary hLibrary
      If nams <> "" Then
         '         Dim xs As String
         '         Dim si As String
         '         Dim i As Long
         '         For i = 1 To Len(s)
         '            xs = HexFromString(Mid(s, i, 1))
         '            si = si & xs & " "
         '         Next
         '         MsgBox nams
         '         MsgBox sn
         '         MsgBox si 'scrambled hex form to get username and companyname TODO
         Dim hUpdate As Long
         hUpdate = apiBeginUpdateResourceA(pth, 0)
         If hUpdate <> 0 Then
            Dim ret As Long
            Dim sNewValue As String
            sNewValue = nams & Chr(0) & ser
            ret = apiUpdateResourceA(hUpdate, 106, 196, 1033, sNewValue, lDataSize)
            If ret <> 0 Then
               ret = apiEndUpdateResourceA(hUpdate, 0)
               If ret = 0 Then
                  MsgBox "Never finished/finalized updating resource"
               End If
            Else
               MsgBox "Never tried to actually update resource"
            End If
         Else
            MsgBox "Never begun to update resource"
         End If
      End If
   Else
      MsgBox "Failed to load VB6.exe library with error code " & Err.LastDllError
   End If
End Sub
Public Function HexFromString(ByVal txt As String) As String
   On Error Resume Next
   Dim i As Long
   For i = 1 To VBA.Len(txt)
      HexFromString = HexFromString & VBA.Hex$(VBA.Asc(VBA.Mid(txt, i, 1))) & VBA.Space$(1)
   Next
   HexFromString = VBA.Mid$(HexFromString, 1, VBA.Len(HexFromString) - 1)
End Function
Private Function GetRandomNumber() As String
   Dim i As Long
   Dim r As Long
   Randomize
   i = Int((90000 * Rnd) + 1)
   r = i + 9999
   GetRandomNumber = CStr(r)
End Function

