VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBCSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'----------------------------
'BASeCamp System Tray Icon Class
'encapsulates a single Icon in the system tray.


Private Const WM_USER As Long = &H400
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128      'shell 5+
   dwState As Long            'shell 5+
   dwStateMask As Long        'shell 5+
   szInfo As String * 256     'shell 5+
   uTimeoutAndVersion As Long 'shell 5+
   szInfoTitle As String * 64 'shell 5+
   dwInfoFlags As Long        'shell 5+
   guidItem As GUID           'shell 6+
End Type
Public Enum NotifyIconFlags
    NIF_ICON = &H2
    NIF_INFO = &H10
    NIF_MESSAGE = &H1
    NIF_STATE = &H8
    NIF_TIP = &H4
End Enum
Private Const NOTIFYICON_VERSION = &H3




Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'icon flags

   
Private Const WM_USER = &H400


'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long




Public Enum NotifyIconIconFlags
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_GUID = &H5
    NIIF_ICON_MASK = &HF
    NIIF_NOSOUND = &H10
End Enum
Public Enum NotifyIconMessage
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIM_SETFOCUS = &H3
    NIM_SETVERSION = &H4
    NIM_VERSION = &H5
End Enum
Private Const NINF_KEY = &H1
Public Enum NotifyIconNotifications
    NIN_GETVERSION = (WM_USER)
    NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
    NIN_SELECT = (WM_USER + &H2)
    
    NIN_BALLOONSHOW = (WM_USER + 2)
    NIN_BALLOONHIDE = (WM_USER + 3)
    NIN_BALLOONTIMEOUT = (WM_USER + 4)
    NIN_BALLOONUSERCLICK = (WM_USER + 5)
End Enum
Public Enum NotifyIconStates
    NIS_HIDDEN = &H1
    NIS_SHAREDICON = &H2
End Enum




Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, ByRef lpData As NOTIFYICONDATA) As Long






Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long
   Private NID As NOTIFYICONDATA
   
   Private mhWndUse As Long
   Private mIconID As Long
   Private mhIcon As Long
   Private mszTip As String
   
   Public Property Let Icon(ByVal Vdata As Long)
    mhIcon = Vdata
   End Property
   Public Property Get Icon() As Long
    Icon = mhIcon
   End Property
   Public Property Let Tip(ByVal Vdata As String)
    mszTip As String
   End Property
   Public Property Get Tip() As String
    Tip = mszTip
   End Property
   
'Private member variables.

   
   
Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select
      
End Sub


Private Function IsShellVersion(ByVal version As Integer) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2  '*** see below for explanation
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function


Private Sub Class_Initialize()
    If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
End Sub
