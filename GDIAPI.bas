Attribute VB_Name = "GDIAPI"
Option Explicit
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
   ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
   ByVal cxWidth As Long, ByVal cyWidth As Long, _
   ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
   ByVal diFlags As Long) As Long
Public Const DI_NORMAL = &H3

Public Function AppPath() As String
   AppPath = App.path
   'Root ends in "\" so check first
   If Right$(AppPath, 1) <> "\" Then
      AppPath = AppPath & "\"
   End If
End Function

Public Sub DrawShell32Icon(ByVal hDC As Long, _
   ByVal Idx As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   Optional ByVal bSmall As Boolean = False)
   
   Dim hIcon      As Long
   Dim IconSize   As Long
   
   If bSmall Then 'Get small(16x16)
      ExtractIconEx "shell32.dll", Idx, ByVal 0&, hIcon, 1
   Else           'Get large(32x32)
      ExtractIconEx "shell32.dll", Idx, hIcon, ByVal 0&, 1
   End If
   If hIcon Then
      IconSize = IIf(bSmall, 16, 32)
         DrawIconEx hDC, x, y, _
            hIcon, _
            IconSize, _
            IconSize, _
            0, 0, DI_NORMAL
      'Cleanup
      DestroyIcon hIcon
   End If
End Sub

