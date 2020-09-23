VERSION 5.00
Begin VB.UserControl ucProgressBar 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ScaleHeight     =   990
   ScaleWidth      =   3000
   ToolboxBitmap   =   "ucProgressBar.ctx":0000
End
Attribute VB_Name = "ucProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucProgressBar - A Selfsubclassed Theme Aware ProgressBar Control which Provides Dynamic Properties
'
'   Product Name:
'       ucProgressBar.ctl
'
'   Compatability:
'       Widnows: 9x, ME, NT, 2K, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Paul Caton - Self-Subclassser)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       (Mario Flores - Cool XP ProgressBar 2.0)
'           http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56151&lngWId=1
'       (Randy Birch - IsWinXP)
'           http://vbnet.mvps.org/code/system/getversionex.htm
'
'   Legal Copyright & Trademarks:
'       Copyright © 2006-2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2006-2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Advance Research Systems shall not be liable for
'       any incidental or consequential damages suffered by any use of this software.
'       This software is owned by Paul R. Territo, Ph.D and is sold for use as a
'       license in accordance with the terms of the License Agreement in the
'       accompanying the documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       pwterrito@insightbb.com
'
'-  Modification(s) History:
'
'       25May06 - Initial Usercontrol Build (Modified from Mario Flores Cool XP ProgressBar 2.0)
'               - Added IsWinXP Method to handle non XP OS
'               - Added Classic ScrollBar Style
'               - Added Theme Support to allow Auto Selection
'
'   Build Date & Time: 6/25/2006 10:12:36 PM
Const Major As Long = 2
Const Minor As Long = 1
Const Revision As Long = 32
Const DateTime As String = "6/25/2006 10:12:36 PM "
'
'   Force Declarations
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal fnStyle As Integer, ByVal COLORREF As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As String, ByVal dwMaxNameChars As Integer, ByVal pszColorBuff As String, ByVal cchMaxColorChars As Integer, ByVal pszSizeBuff As String, ByVal cchMaxSizeChars As Integer) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: string providing arbitrary additional information
End Type

Private Const VER_PLATFORM_WIN32_NT = 2

Public Enum upbThemeEnum
    [upbAuto] = &H0
    [upbClassic] = &H1
    [upbBlue] = &H2
    [upbHomeStead] = &H3
    [upbMetallic] = &H4
End Enum

'=====================================================
'TEXT FORMAT CONST
Const DT_SINGLELINE   As Long = &H20
Const DT_CALCRECT     As Long = &H400
'=====================================================

'=====================================================
'BORDER FIELD CONST
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'=====================================================

'=====================================================
'THE POINTAPI STRUCTURE
Private Type POINTAPI
    X As Long                       ' The POINTAPI structure defines the x- and y-coordinates of a point.
    Y As Long
End Type
'=====================================================

'=====================================================
'THE RECT STRUCTURE
Private Type RECT
    Left      As Long     'The RECT structure defines the coordinates of the upper-left and lower-right corners of a rectangle
    Top       As Long
    Right     As Long
    Bottom    As Long
End Type
'=====================================================

'=====================================================
'THE BRUSHSTYLE ENUM
Public Enum BrushStyle
    HS_HORIZONTAL = 0
    HS_VERTICAL = 1
    HS_FDIAGONAL = 2
    HS_BDIAGONAL = 3
    HS_CROSS = 4
    HS_DIAGCROSS = 5
    HS_SOLID = 6
End Enum
'=====================================================

'=====================================================
'THE COOL XP PROGRESSBAR 2.0 STYLES
Public Enum cScrolling
    ccScrollingStandard = 0
    ccScrollingSmooth = 1
    ccScrollingSearch = 2
    ccScrollingOfficeXP = 3
    ccScrollingPastel = 4
    ccScrollingJavT = 5
    ccScrollingMediaPlayer = 6
    ccScrollingCustomBrush = 7
    ccScrollingPicture = 8
    ccScrollingMetallic = 9
    ccScrollingClassic = 10
End Enum
'=====================================================

'=====================================================
'THE ORIENTATION ENUM
Public Enum cOrientation
    ccOrientationHorizontal = 0
    ccOrientationVertical = 1
End Enum
'=====================================================

'----------------------------------------------------
Private m_Color       As OLE_COLOR
Private m_hDC         As Long
Private m_hWnd        As Long        'PROPERTIES VARIABLES
Private m_Max         As Long
Private m_Min         As Long
Private m_Value       As Long
Private m_ShowText    As Boolean
Private m_Scrolling   As cScrolling
Private m_Orientation As cOrientation
Private m_Brush       As BrushStyle
Private m_Picture     As StdPicture
Private m_Theme       As upbThemeEnum
'----------------------------------------------------

'----------------------------------------------------
Private m_MemDC    As Boolean
Private m_ThDC     As Long
Private m_hBmp     As Long
Private m_hBmpOld  As Long
Private iFnt       As IFont
Private m_fnt      As IFont          'VARIABLES USED IN PROCESS
Private hFntOld    As Long
Private m_lWidth   As Long
Private m_lHeight  As Long
Private fPercent   As Double
Private tR         As RECT
Private TBR        As RECT
Private TSR        As RECT
Private AT         As RECT
Private lSegmentWidth   As Long
Private lSegmentSpacing As Long

'==================================================================================================
' ucSubclass - A template UserControl for control authors that require self-subclassing without ANY
'              external dependencies. IDE safe.
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........
'==================================================================================================
'Subclasser declarations

Public Event MouseEnter()
Public Event MouseLeave()
Public Event Status(ByVal sStatus As String)

Private Const WM_EXITSIZEMOVE           As Long = &H232
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOVING                 As Long = &H216
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_SIZING                 As Long = &H214
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_THEMECHANGED           As Long = &H31A
Private Const WM_USER                   As Long = &H400

Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean
Private bSubClass                    As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                   'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                    'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                   'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                   'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                   'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                  'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                  'Table A (after) entry count patch offset

Private Type tSubData                                                               'Subclass data type
    hwnd                               As Long                                      'Handle of the window being subclassed
    nAddrSub                           As Long                                      'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                      'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                      'Msg after table entry count
    nMsgCntB                           As Long                                      'Msg before table entry count
    aMsgTblA()                         As Long                                      'Msg after table array
    aMsgTblB()                         As Long                                      'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                    'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    'Parameters:
        'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
        'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
        'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
        'hWnd     - The window handle
        'uMsg     - The message number
        'wParam   - Message related data
        'lParam   - Message related data
    'Notes:
        'If you really know what you're doing, it's possible to change the values of the
        'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
        'values get passed to the default handler.. and optionaly, the 'after' callback
    Static bMoving As Boolean
   
    Select Case uMsg
        Case WM_SYSCOLORCHANGE
            Refresh
        
        Case WM_THEMECHANGED
            Refresh
            
    End Select
    
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

    'Add a message to the table of those that will invoke a callback. You should Subclass_Subclass first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
        'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
        'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
        'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    'Parameters:
    'lng_hWnd  - The handle of the window to be subclassed
    'Returns;
    'The sc_aSubData() index
    Const CODE_LEN              As Long = 204                                       'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"                       'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                                'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"                        'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                                'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                                  'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                                  'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                                        'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                                        'Address of the previous WndProc
    Const PATCH_03              As Long = 78                                        'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                                       'Address of the previous WndProc
    Const PATCH_07              As Long = 121                                       'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                                       'Address of the owner object
    Static aBuf(1 To CODE_LEN)  As Byte                                             'Static code buffer byte array
    Static pCWP                 As Long                                             'Address of the CallWindowsProc
    Static pEbMode              As Long                                             'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                             'Address of the SetWindowsLong function
    Dim i                       As Long                                             'Loop index
    Dim j                       As Long                                             'Loop index
    Dim nSubIdx                 As Long                                             'Subclass data index
    Dim sHex                    As String                                           'Hex code string
    
    'If it's the first time through here..
    If aBuf(1) = 0 Then
        
        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
            "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
            "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
            "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                  'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                        'Next pair of hex characters
        
        'Get API function addresses
        If Subclass_InIDE Then                                                      'If we're running in the VB IDE
            aBuf(16) = &H90                                                         'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                         'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                 'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                     'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                             'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                        'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                       'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                        'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                     'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                    'Create a new sc_aSubData element
        End If
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lng_hWnd                                                            'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                               'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                  'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                      'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                             'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                   'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                             'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                   'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                             'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long
    
    i = UBound(sc_aSubData())                                                       'Get the upper bound of the subclass data array
    Do While i >= 0                                                                 'Iterate through each element
        With sc_aSubData(i)
            If .hwnd <> 0 Then                                                      'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hwnd)                                           'Subclass_Stop
            End If
        End With
        i = i - 1                                                                   'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    'Parameters:
    'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                         'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                      'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                      'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                  'Release the machine code memory
        .hwnd = 0                                                                   'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                               'Clear the before table
        .nMsgCntA = 0                                                               'Clear the after table
        Erase .aMsgTblB                                                             'Erase the before table
        Erase .aMsgTblA                                                             'Erase the after table
    End With
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
    
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for sc_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long                                                             'Message table entry index
    Dim nOff1   As Long                                                             'Machine code buffer offset 1
    Dim nOff2   As Long                                                             'Machine code buffer offset 2
    
    If uMsg = ALL_MESSAGES Then                                                     'If all messages
        nMsgCnt = ALL_MESSAGES                                                      'Indicates that all messages will callback
    Else                                                                            'Else a specific message number
        Do While nEntry < nMsgCnt                                                   'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            
            If aMsgTbl(nEntry) = 0 Then                                             'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                              'Re-use this entry
                Exit Sub                                                            'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                      'The msg is already in the table!
                Exit Sub                                                            'Bail
            End If
        Loop                                                                        'Next entry
        nMsgCnt = nMsgCnt + 1                                                       'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                     'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                              'If before
        nOff1 = PATCH_04                                                            'Offset to the Before table
        nOff2 = PATCH_05                                                            'Offset to the Before table entry count
    Else                                                                            'Else after
        nOff1 = PATCH_08                                                            'Offset to the After table
        nOff2 = PATCH_09                                                            'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                            'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                           'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                          'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for sc_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                                                     'If deleting all messages
        nMsgCnt = 0                                                                 'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                          'If before
            nEntry = PATCH_05                                                       'Patch the before table message count location
        Else                                                                        'Else after
            nEntry = PATCH_09                                                       'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                            'Patch the table message count to zero
    Else                                                                            'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                   'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                          'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                 'Mark the table slot as available
                Exit Do                                                             'Bail
            End If
        Loop                                                                        'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably sc_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                              'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hwnd = lng_hWnd Then                                                'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                    'If we're searching not adding
                    Exit Function                                                   'Found
                End If
            ElseIf .hwnd = 0 Then                                                   'If this an element marked for reuse.
                If bAdd Then                                                        'If we're adding
                    Exit Function                                                   'Re-use it
                End If
            End If
        End With
    zIdx = zIdx - 1                                                                 'Decrement the index
    Loop
    
    If Not bAdd Then
        Debug.Assert False                                                          'hWnd not found, programmer error
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
'======================================================================================================
'   End SubClass Sections
'======================================================================================================

Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
    '======================================================================
    'BLENDS 2 COLORS WITH A PREDEFINED ALPHA VALUE
    Dim lCFrom As Long
    Dim lCTo As Long
    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
    
    lCFrom = GetLngColor(oColorFrom)
    lCTo = GetLngColor(oColorTo)
    
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    
    BlendColor = RGB( _
    ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
    ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
    ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
    )
    
End Function

Public Property Let BrushStyle(ByVal Style As BrushStyle)
    m_Brush = Style
    PropertyChanged "BrushStyle"
End Property

Private Sub CalcBarSize()
    '==========================================================
    '/---Calculate Division Bars & Percent Values
    '==========================================================
    lSegmentWidth = IIf(m_Scrolling = 0, 6, 0) '/-- Windows Default
    lSegmentSpacing = 2                        '/-- Windows Default
    tR.Left = tR.Left + 3
    LSet TBR = tR
    fPercent = m_Value / 98
    If fPercent < 0# Then fPercent = 0#
    If m_Orientation = 0 Then
        '=======================================================================================
        '                                 Calc Horizontal ProgressBar
        '---------------------------------------------------------------------------------------
        TBR.Right = tR.Left + (tR.Right - tR.Left) * fPercent
        TBR.Right = TBR.Right - ((TBR.Right - TBR.Left) Mod (lSegmentWidth + lSegmentSpacing))
        If TBR.Right < tR.Left Then
            TBR.Right = tR.Left
        End If
    Else
        '=======================================================================================
        '                                 Calc Vertical ProgressBar
        '---------------------------------------------------------------------------------------
        fPercent = 1# - fPercent
        TBR.Top = tR.Top + (tR.Bottom - tR.Top) * fPercent
        TBR.Top = TBR.Top - ((TBR.Top - TBR.Bottom) Mod (lSegmentWidth + lSegmentSpacing))
        If TBR.Top > tR.Bottom Then TBR.Top = tR.Bottom
    End If
End Sub

Private Sub CalculateAlphaTextRect(ByVal ThisText As String)
    '======================================================================
    'ALPHA TEXT RECT FUNCTION
    
    '//--Calculates the Bounding Rects Of the Text using DT_CALCRECT
    DrawText m_hDC, ThisText, Len(ThisText), AT, DT_CALCRECT
    AT.Left = (tR.Right / 2) - ((AT.Right - AT.Left) / 2)
    AT.Top = (tR.Bottom / 2) - ((AT.Bottom - AT.Top) / 2)
End Sub

Public Property Get Color() As OLE_COLOR
    Color = m_Color
End Property

Public Property Let Color(ByVal lColor As OLE_COLOR)
    m_Color = GetLngColor(lColor)
    DrawProgressBar
End Property

Private Sub DrawAlphaText(ByVal ThisText As String)
    '======================================================================
    'ALPHA TEXT FUNCTION
    
    Set iFnt = Font                             '//--New Font
    hFntOld = SelectObject(m_hDC, iFnt.hFont)   '//--Use the New Font
    SetBkMode m_hDC, 1                          '//--Transparent Text
    '//-- This is When the Text is Drawn
    '//--Gives the Media Player Text Look (Changes Color When Progress is over the Text)
    If (tR.Right * (m_Value / 100)) >= AT.Left Then
        SetTextColor m_hDC, GetLngColor(IIf(m_Scrolling = ccScrollingMediaPlayer, ShiftColorXP(m_Color, 80), vbWhite))
        AT.Left = (tR.Right / 2) - ((AT.Right - AT.Left) / 2)
        AT.Right = (tR.Right * (m_Value / 100))
        DrawText m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE
        SelectObject m_hDC, hFntOld
    End If
    
End Sub

Private Sub DrawCustomBrushProgressbar()
    '==========================================================
    '/---CUSTOM BRUSH XP STYLE
    '==========================================================
    Dim hBrush As Long
    
    DrawEdge m_hDC, tR, 9, BF_RECT
    With TBR
        .Left = 2
        .Top = 2
        .Bottom = tR.Bottom - 2
        .Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
    End With
    
    hBrush = CreateHatchBrush(m_Brush, GetLngColor(Color))
    SetBkColor m_hDC, ShiftColorXP(m_Color, 140)
    FillRect m_hDC, TBR, hBrush
    DeleteObject hBrush
End Sub

Private Sub DrawDivisions()
    '==========================================================
    '/---Draw Division Bars
    '==========================================================
    Dim i As Long
    Dim hBR As Long
    
    hBR = CreateSolidBrush(vbWhite)
    LSet TSR = tR
    If m_Orientation = 0 Then
        '=======================================================================================
        '                                 Draw Horizontal ProgressBar
        '---------------------------------------------------------------------------------------
        For i = TBR.Left + lSegmentWidth To TBR.Right Step lSegmentWidth + lSegmentSpacing
            TSR.Left = i + 1
            TSR.Right = i + 1 + lSegmentSpacing
            FillRect m_hDC, TSR, hBR
        Next i
        '---------------------------------------------------------------------------------------
    Else
        '=======================================================================================
        '                                  Draw Vertical ProgressBar
        '---------------------------------------------------------------------------------------
        For i = TBR.Bottom To TBR.Top + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
            TSR.Top = i - 2
            TSR.Bottom = i - 2 + lSegmentSpacing
            FillRect m_hDC, TSR, hBR
        Next i
        '---------------------------------------------------------------------------------------
    End If
    DeleteObject hBR
End Sub

Private Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)
    '======================================================================
    'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
    Dim hBrush As Long
    
    hBrush = CreateSolidBrush(GetLngColor(Color))
    FillRect MyHdc, hRect, hBrush
    DeleteObject hBrush
End Sub

Public Sub DrawGradient(lEndColor As Long, lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hDC As Long, Optional bH As Boolean)
    '======================================================================
    'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
    On Error Resume Next
    
    ''Draw a Vertical Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    lEndColor = GetLngColor(lEndColor)
    lStartColor = GetLngColor(lStartColor)
    
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    sR = (sR - eR) / IIf(bH, X2, Y2)
    sG = (sG - eG) / IIf(bH, X2, Y2)
    sB = (sB - eB) / IIf(bH, X2, Y2)
    For ni = 0 To IIf(bH, X2, Y2)
        If bH Then
            DrawLine X + ni, Y, X + ni, Y2, hDC, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Else
            DrawLine X, Y + ni, X2, Y + ni, hDC, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        End If
    Next ni
End Sub
'======================================================================
Private Sub DrawJavTProgressbar()
    '==========================================================
    '/---JAVT XP STYLE
    '==========================================================
    DrawRectangle tR, ShiftColorXP(m_Color, 10), m_hDC
    TBR.Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
    DrawGradient m_Color, ShiftColorXP(m_Color, 100), 2, 2, tR.Right - 2, tR.Bottom - 5, m_hDC ', True
    DrawGradient ShiftColorXP(m_Color, 250), m_Color, 3, 3, TBR.Right, tR.Bottom - 6, m_hDC  ', True
    DrawLine TBR.Right, 2, TBR.Right, tR.Bottom - 2, m_hDC, ShiftColorXP(m_Color, 25)
    
End Sub

Public Sub DrawLine( _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal cHdc As Long, _
    ByVal Color As Long)
    '======================================================================
    'DRAWS A LINE WITH A DEFINED COLOR
    
    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim Outline As Long
    Dim Pos     As POINTAPI
    
    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    
    MoveToEx cHdc, X, Y, Pos
    LineTo cHdc, Width, Height
    
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1
    
End Sub

Private Sub DrawClassicProgressbar()
    '==========================================================
    '/---CLASSIC STYLE
    '==========================================================
    DrawRectangle tR, &HFFFFFF, m_hDC
    InflateRect tR, -1, -1
    DrawRectangle tR, TranslateColor(UserControl.Parent.BackColor), m_hDC
    DrawLine 0, 0, tR.Left + (tR.Right - tR.Left), 0, m_hDC, &H99A8AC
    DrawLine 0, 0, 0, tR.Bottom, m_hDC, &H99A8AC
    DrawLine 1, 1, tR.Left + (tR.Right - tR.Left), 1, m_hDC, &H707070
    DrawLine 1, 1, 1, tR.Bottom, m_hDC, &H707070
    With TBR
        .Left = 2
        .Top = 2
        .Bottom = tR.Bottom - 1
        .Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 100)
    End With
    DrawFillRectangle TBR, &H800000, m_hDC
End Sub

Private Sub DrawMediaProgressbar()
    '==========================================================
    '/---MEDIA PROGRESS XP STYLE
    '==========================================================
    DrawRectangle tR, BlendColor(m_Color, &H0, 200), m_hDC
    DrawGradient &H0&, ShiftColorXP(GetLngColor(BlendColor(m_Color, &H0, 100)), 10), 2, 2, tR.Left + (tR.Right - tR.Left - 5) * (m_Value / 100), tR.Bottom - 2, m_hDC, True
End Sub

Private Sub DrawMetalProgressbar()
    '==========================================================
    '/---METALLIC XP STYLE
    '==========================================================
    TBR.Right = tR.Left + (tR.Right - tR.Left - 4) * (m_Value / 100)
    DrawGradient vbWhite, &HC0C0C0, 2, 2, tR.Right - 3, (tR.Bottom - 3) / 2, m_hDC
    DrawGradient BlendColor(&HC0C0C0, &H0, 255), &HC0C0C0, 2, (tR.Bottom - 3) / 2, tR.Right - 3, (tR.Bottom - 3) / 2, m_hDC
    DrawGradient ShiftColorXP(m_Color, 150), BlendColor(m_Color, &H0, 180), 2, 2, TBR.Right, (tR.Bottom - 3) / 2, m_hDC
    DrawGradient BlendColor(m_Color, &H0, 190), m_Color, 2, (tR.Bottom - 3) / 2, TBR.Right, (tR.Bottom - 3) / 2, m_hDC
    tR.Left = tR.Left + 3
    pDrawBorder
End Sub

Private Sub DrawOfficeXPProgressbar()
    '==========================================================
    '/---OFFICE XP STYLE
    '==========================================================
    DrawRectangle tR, ShiftColorXP(m_Color, 100), m_hDC
    With TBR
        .Left = 1
        .Top = 1
        .Bottom = tR.Bottom - 1
        .Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 100)
    End With
    DrawFillRectangle TBR, ShiftColorXP(m_Color, 180), m_hDC
End Sub

Private Sub DrawPastelProgressbar()
    '==========================================================
    '/---PASTEL XP STYLE
    '==========================================================
    DrawEdge m_hDC, tR, 6, BF_RECT
    DrawGradient ShiftColorXP(m_Color, 140), ShiftColorXP(m_Color, 200), 2, 2, tR.Left + (tR.Right - tR.Left - 4) * (m_Value / 100), tR.Bottom - 3, m_hDC, True
End Sub

Private Sub DrawPictureProgressbar()
    '==========================================================
    '/---PICTURE STYLE
    '==========================================================
    Dim Brush      As Long
    Dim origBrush  As Long
    DrawEdge m_hDC, tR, 2, BF_RECT                       '//--- Draw ProgressBar Border
    If Nothing Is m_Picture Then Exit Sub                '//--- In Case No Picture is Choosen
    Brush = CreatePatternBrush(m_Picture.handle)         '//-- Use Pattern Picture Draw
    origBrush = SelectObject(m_hDC, Brush)
    TBR.Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
    PatBlt m_hDC, 2, 2, TBR.Right, tR.Bottom - 4, vbPatCopy
    SelectObject m_hDC, origBrush
    DeleteObject Brush
    
End Sub

Public Sub DrawProgressBar()
    '==========================================================
    '/---Draw ALL ProgressXP Bar  !!!!PUBLIC CALL!!!
    '==========================================================
    If m_Value > 100 Then m_Value = 100
    GetClientRect m_hWnd, tR               '//--- Reference = Control Client Area
    DrawFillRectangle tR, IIf(m_Scrolling = ccScrollingMediaPlayer, &H0, vbWhite), m_hDC '//--- Draw BackGround
    '//-- Draw ProgressBar Style
    '==========================================================
    '/---Draw METALLIC XP STYLE
    '==========================================================
    If m_Scrolling = ccScrollingMetallic Then
        DrawMetalProgressbar
        '==========================================================
        '/---Draw OFFICE XP STYLE
        '==========================================================
    ElseIf m_Scrolling = ccScrollingOfficeXP Then
        DrawOfficeXPProgressbar
        '==========================================================
        '/---Draw PASTEL XP STYLE
        '==========================================================
    ElseIf m_Scrolling = ccScrollingPastel Then
        DrawPastelProgressbar
        '==========================================================
        '/---Draw JAVT XP STYLE
        '==========================================================
    ElseIf m_Scrolling = ccScrollingJavT Then
        DrawJavTProgressbar
        '==========================================================
        '/---Draw MEDIA PLAYER XP STYLE
        '==========================================================
    ElseIf m_Scrolling = ccScrollingMediaPlayer Then
        DrawMediaProgressbar
        '==========================================================
        '/---Draw CUSTOM BRUSH XP WASH COLOR STYLE
        '==========================================================
    ElseIf m_Scrolling = ccScrollingCustomBrush Then
        DrawCustomBrushProgressbar
        '==========================================================
        '/---Draw PICTURE STYLE
        '==========================================================
    ElseIf m_Scrolling = ccScrollingPicture Then
        DrawPictureProgressbar
    ElseIf m_Scrolling = ccScrollingClassic Then
        DrawClassicProgressbar
    Else
        '==========================================================
        '/---Draw WINDOWS XP STYLE
        '==========================================================
        CalcBarSize                            '//--- Calculate Progress and Percent Values
        PBarDraw                               '//--- Draw Scolling Bar (Inside Bar)
        If m_Scrolling = 0 Then DrawDivisions  '//--- Draw SegmentSpacing (This Will Generate the Blocks Effect)
        pDrawBorder                            '//--- Draw The XP Look Border
    End If
    '==========================================================
    DrawTexto                                  '//--- Draw The Percent Text
    '==========================================================
    '/---Use the AntiFlicker DC
    '==========================================================
    If m_MemDC Then
        With UserControl
            pDraw .hDC, 0, 0, .ScaleWidth, .ScaleHeight, .ScaleLeft, .ScaleTop
        End With
    End If
End Sub

Private Sub DrawRectangle(ByRef BRect As RECT, ByVal Color As Long, ByVal hDC As Long)
    '======================================================================
    'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
    Dim hBrush As Long
    
    hBrush = CreateSolidBrush(Color)
    FrameRect hDC, BRect, hBrush
    DeleteObject hBrush
End Sub

Private Function DrawTexto()
    '======================================================================
    'DRAWS THE PERCENT TEXT ON PROGRESS BAR
    Dim ThisText As String
    Dim isAlpha  As Boolean
    
    If (m_Scrolling = ccScrollingMediaPlayer Or m_Scrolling = ccScrollingMetallic) Then isAlpha = True
    If m_Scrolling = ccScrollingSearch Then
        ThisText = "Searching.."
    Else
        ThisText = Round(m_Value) & " %"
    End If
    If (m_ShowText) Then
        Set iFnt = Font                             '//--New Font
        hFntOld = SelectObject(m_hDC, iFnt.hFont)   '//--Use the New Font
        SetBkMode m_hDC, 1                          '//--Transparent Text
        '//--Use the Alpha Text Color Look if Progress is MediaPlayer Style, else Normal (Gray)
        SetTextColor m_hDC, GetLngColor(IIf(m_Scrolling = ccScrollingMediaPlayer, &HC0C0C0, vbBlack))
        CalculateAlphaTextRect ThisText             '//--Calculate The Text Rectangle
        '//-- If ProgressBar is already over the Text don't draw the old text, yust draw the Alpha Text
        'It saves some memory
        If ((tR.Right * (m_Value / 100)) <= AT.Right) Or Not isAlpha Then
            DrawText m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE
        End If
        SelectObject m_hDC, hFntOld  'Delete the Used Font
        '//--Use the Alpha Text Look if Progress is AlPhA Style
        If isAlpha Then DrawAlphaText ThisText
    End If
End Function

Public Property Get Font() As IFont
    Set Font = m_fnt
End Property

Public Property Let Font(ByRef fnt As IFont)
    Set m_fnt = fnt
End Property

Public Property Set Font(ByRef fnt As IFont)
    Set m_fnt = fnt    'Defined By System but can change by user choice.(ADD Property!!)
End Property

Private Function GetLngColor(Color As Long) As Long
    '======================================================================
    'CONVERTION FUNCTION
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

Private Function GetThemeInfo() As String
    Dim lResult As Long
    Dim sFileName As String
    Dim sColor As String
    Dim lPos As Long
    
    If IsWinXP Then
        '   Allocate Space
        sFileName = Space(255)
        sColor = Space(255)
        '   Read the data
        If GetCurrentThemeName(sFileName, 255, sColor, 255, vbNullString, 0) <> &H0 Then
            GetThemeInfo = "UxTheme_Error"
            Exit Function
        End If
        '   Find our trailing null terminator
        lPos = InStrRev(sColor, vbNullChar)
        '   Parse it....
        sColor = Mid(sColor, 1, lPos)
        '   Now replace the nulls....
        sColor = Replace(sColor, vbNullChar, "")
        If Trim$(sColor) = vbNullString Then sColor = "None"
        GetThemeInfo = sColor
    Else
        sColor = "None"
    End If
End Function

Private Sub GetThemeProgressBar()
    Dim AutoTheme As String
    
    Select Case m_Theme
        Case [upbAuto]
            AutoTheme = GetThemeInfo
            Select Case AutoTheme
                Case "None"
                    GoTo Classic
                Case "NormalColor"
                    GoTo Blue
                Case "HomeStead"
                    GoTo HomeStead
                Case "Metallic"
                    GoTo Metallic
            End Select
        Case [upbClassic]
Classic:
            Scrolling = ccScrollingClassic
            DrawProgressBar
        Case [upbBlue]
Blue:
            Color = &HC56A31
            Scrolling = ccScrollingStandard
            DrawProgressBar
        Case [upbHomeStead]
HomeStead:
            Color = &H69A18B
            Scrolling = ccScrollingStandard
            DrawProgressBar
        Case [upbMetallic]
Metallic:
            Color = &HC0C0C0
            Scrolling = ccScrollingStandard
            DrawProgressBar
    End Select

End Sub

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Let hDC(ByVal cHdc As Long)
    '=============================================
    'AntiFlick...Cleaner HDC
    m_hDC = ThDC(UserControl.ScaleWidth, UserControl.ScaleHeight)
    
    If m_hDC = 0 Then
        m_hDC = UserControl.hDC   'On Fail...Do it Normally
    Else
        m_MemDC = True
    End If
    '=============================================
End Property

Public Property Get hwnd() As Long
    hwnd = m_hWnd
End Property

Public Property Let hwnd(ByVal chwnd As Long)
    m_hWnd = chwnd
End Property

Public Property Get Image() As StdPicture
    If Nothing Is m_Picture Then Exit Property
    Set Image = m_Picture
End Property

Public Property Set Image(ByVal handle As StdPicture)
    Set m_Picture = handle
    PropertyChanged "Image"
    DrawProgressBar
End Property

Public Function IsWinXP() As Boolean
    'returns True if running Windows XP
    Dim OSV As OSVERSIONINFO

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        IsWinXP = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
            (OSV.dwVerMajor = 5 And OSV.dwVerMinor = 1) And _
            (OSV.dwBuildNumber >= 2600)
    End If
End Function

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal cMaX As Long)
    m_Max = cMaX
    PropertyChanged "Max"
End Property

Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal cMin As Long)
    m_Min = cMin
    PropertyChanged "Min"
End Property

Public Property Get Orientation() As cOrientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal cOrientation As cOrientation)
    m_Orientation = cOrientation
    PropertyChanged "Orientation"
    DrawProgressBar
End Property

Private Sub PBarDraw()
    '==========================================================
    '/---Draw The ProgressXP Bar ;)
    '==========================================================
    Dim TempRect As RECT
    Dim iTemp    As Long
    
    If m_Orientation = 0 Then
        If TBR.Right <= 14 Then TBR.Right = 12
        TempRect.Left = 4
        TempRect.Right = IIf(TBR.Right + 4 > tR.Right, TBR.Right - 4, TBR.Right)
        TempRect.Top = 8
        TempRect.Bottom = tR.Bottom - 8
        '=======================================================================================
        '                                 Draw Horizontal ProgressBar
        '---------------------------------------------------------------------------------------
        If m_Scrolling = ccScrollingSearch Then
            GoSub HorizontalSearch
        Else
            DrawGradient ShiftColorXP(m_Color, 150), m_Color, 4, 3, TempRect.Right, 6, m_hDC
            DrawFillRectangle TempRect, m_Color, m_hDC
            DrawGradient m_Color, ShiftColorXP(m_Color, 150), 4, TempRect.Bottom - 2, TempRect.Right, 6, m_hDC
        End If
    Else
        TempRect.Left = 9
        TempRect.Right = tR.Right - 8
        TempRect.Top = TBR.Top
        TempRect.Bottom = tR.Bottom
        '=======================================================================================
        '                                 Draw Vertical ProgressBar
        '---------------------------------------------------------------------------------------
        If m_Scrolling = ccScrollingSearch Then
            GoSub VerticalSearch
        Else
            DrawGradient ShiftColorXP(m_Color, 150), m_Color, 4, TBR.Top, 4, tR.Bottom, m_hDC, True
            DrawFillRectangle TempRect, m_Color, m_hDC
            DrawGradient m_Color, ShiftColorXP(m_Color, 150), tR.Right - 8, TBR.Top, 4, tR.Bottom, m_hDC, True
        End If
        '--------------------   <-------- Gradient Color From (- to +)
        '||||||||||||||||||||   <-------- Fill Color
        '--------------------   <-------- Gradient Color From (+ to -)
    End If
    Exit Sub
    
HorizontalSearch:
    For iTemp = 0 To 2
        With TempRect
            .Left = TBR.Right + ((lSegmentSpacing + 10) * (iTemp)) - (45 * ((100 - m_Value) / 100))
            .Right = .Left + 10
            .Top = 8
            .Bottom = tR.Bottom - 8
            DrawGradient ShiftColorXP(m_Color, 220 - (40 * iTemp)), ShiftColorXP(m_Color, 200 - (40 * iTemp)), .Left, 3, 9, tR.Bottom - 2, m_hDC, True
        End With
    Next iTemp
    Return
    
VerticalSearch:
    For iTemp = 0 To 2
        With TempRect
            .Left = 8
            .Right = tR.Right - 8
            .Top = TBR.Top + ((lSegmentSpacing + 10) * iTemp)
            .Bottom = .Top + 10
            DrawGradient ShiftColorXP(m_Color, 220 - (40 * iTemp)), ShiftColorXP(m_Color, 200 - (40 * iTemp)), tR.Right - 2, .Top, 2, 9, m_hDC
        End With
    Next iTemp
    Return
    
End Sub

Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
    '======================================================================
    'CREATES THE TEMP DC
    Dim lhDCC As Long
    pDestroy
    lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If Not (lhDCC = 0) Then
        m_ThDC = CreateCompatibleDC(lhDCC)
        If Not (m_ThDC = 0) Then
            m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
            If Not (m_hBmp = 0) Then
                m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
                If Not (m_hBmpOld = 0) Then
                    m_lWidth = Width
                    m_lHeight = Height
                    DeleteDC lhDCC
                    Exit Sub
                End If
            End If
        End If
        DeleteDC lhDCC
        pDestroy
    End If
End Sub

Private Sub pDestroy()
    '======================================================================
    'DESTROYS THE TEMP DC
    If Not m_hBmpOld = 0 Then
        SelectObject m_ThDC, m_hBmpOld
        m_hBmpOld = 0
    End If
    If Not m_hBmp = 0 Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    If Not m_ThDC = 0 Then
        DeleteDC m_ThDC
        m_ThDC = 0
    End If
    m_lWidth = 0
    m_lHeight = 0
End Sub

Public Sub pDraw( _
    ByVal hDC As Long, _
    Optional ByVal XSrc As Long = 0, Optional ByVal YSrc As Long = 0, _
    Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
    Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
    )
    '======================================================================
    'DRAWS THE TEMP DC
    If WidthSrc <= 0 Then WidthSrc = m_lWidth
    If HeightSrc <= 0 Then HeightSrc = m_lHeight
    BitBlt hDC, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, XSrc, YSrc, vbSrcCopy
    
End Sub

Private Sub pDrawBorder()
    '==========================================================
    '/---Draw The ProgressXP Bar Border  ;)
    '==========================================================
    Dim RTemp As RECT
    
    tR.Left = tR.Left - 3
    
    Let RTemp = tR
    
    DrawLine 2, 1, tR.Right - 2, 1, m_hDC, &HBEBEBE
    DrawLine 2, tR.Bottom - 2, tR.Right - 2, tR.Bottom - 2, m_hDC, &HEFEFEF
    DrawLine 1, 2, 1, tR.Bottom - 2, m_hDC, &HBEBEBE
    DrawLine 2, 2, 2, tR.Bottom - 2, m_hDC, &HEFEFEF
    DrawLine 2, 2, tR.Right - 2, 2, m_hDC, &HEFEFEF
    DrawLine tR.Right - 2, 2, tR.Right - 2, tR.Bottom - 2, m_hDC, &HEFEFEF
    
    DrawRectangle tR, GetLngColor(&H686868), m_hDC
    
    Call SetPixelV(m_hDC, 0, 0, GetLngColor(vbWhite))
    Call SetPixelV(m_hDC, 0, 1, GetLngColor(&HA6ABAC))
    Call SetPixelV(m_hDC, 0, 2, GetLngColor(&H7D7E7F))
    Call SetPixelV(m_hDC, 1, 0, GetLngColor(&HA7ABAC)) '//TOP RIGHT CORNER
    Call SetPixelV(m_hDC, 1, 1, GetLngColor(&H777777))
    Call SetPixelV(m_hDC, 2, 0, GetLngColor(&H7D7E7F))
    Call SetPixelV(m_hDC, 2, 2, GetLngColor(&HBEBEBE))
    
    Call SetPixelV(m_hDC, 0, tR.Bottom - 1, GetLngColor(vbWhite))
    Call SetPixelV(m_hDC, 1, tR.Bottom - 1, GetLngColor(&HA6ABAC))
    Call SetPixelV(m_hDC, 2, tR.Bottom - 1, GetLngColor(&H7D7E7F))
    Call SetPixelV(m_hDC, 0, tR.Bottom - 3, GetLngColor(&H7D7E7F)) '//BOTTOM RIGHT CORNER
    Call SetPixelV(m_hDC, 0, tR.Bottom - 2, GetLngColor(&HA7ABAC))
    Call SetPixelV(m_hDC, 1, tR.Bottom - 2, GetLngColor(&H777777))
    
    Call SetPixelV(m_hDC, tR.Right - 1, 0, GetLngColor(vbWhite))
    Call SetPixelV(m_hDC, tR.Right - 1, 1, GetLngColor(&HBEBEBE))
    Call SetPixelV(m_hDC, tR.Right - 1, 2, GetLngColor(&H7D7E7F)) '//TOP LEFT CORNER
    Call SetPixelV(m_hDC, tR.Right - 2, 2, GetLngColor(&HBEBEBE))
    Call SetPixelV(m_hDC, tR.Right - 2, 1, GetLngColor(&H686868))
    
    Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 1, GetLngColor(vbWhite))
    Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 2, GetLngColor(&HBEBEBE))
    Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 3, GetLngColor(&H7D7E7F))
    Call SetPixelV(m_hDC, tR.Right - 2, tR.Bottom - 2, GetLngColor(&H777777)) '//TOP RIGHT CORNER
    Call SetPixelV(m_hDC, tR.Right - 2, tR.Bottom - 1, GetLngColor(&HBEBEBE))
    Call SetPixelV(m_hDC, tR.Right - 3, tR.Bottom - 1, GetLngColor(&H7D7E7F))
    
End Sub

Public Property Get Scrolling() As cScrolling
    Scrolling = m_Scrolling
End Property

Public Property Let Scrolling(ByVal lScrolling As cScrolling)
    m_Scrolling = lScrolling
    PropertyChanged "Scrolling"
    DrawProgressBar
End Property

Private Function ShiftColorXP(ByVal MyColor As Long, ByVal Base As Long) As Long
'======================================================================
'BLENDS AN SPECIFIED COLOR TO GET XP COLOR LOOK
    
    Dim R As Long, G As Long, b As Long, Delta As Long
    
    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    b = ((MyColor \ &H10000) Mod &H100)
    
    Delta = &HFF - Base
    
    b = Base + b * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF
    
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If b > 255 Then b = 255
    
    ShiftColorXP = R + 256& * G + 65536 * b
    
End Function

Public Property Get ShowText() As Boolean
    ShowText = m_ShowText
End Property

Public Property Let ShowText(ByVal bShowText As Boolean)
    m_ShowText = bShowText
    PropertyChanged "ShowText"
    DrawProgressBar
End Property

Private Function ThDC(Width As Long, Height As Long) As Long
    '======================================================================
    'CHECKS-CREATES CORRECT DIMENSIONS OF THE TEMP DC
    If m_ThDC = 0 Then
        If (Width > 0) And (Height > 0) Then
            pCreate Width, Height
        End If
    Else
        If Width > m_lWidth Or Height > m_lHeight Then
            pCreate Width, Height
        End If
    End If
    ThDC = m_ThDC
End Function

Public Property Get Theme() As upbThemeEnum
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As upbThemeEnum)
    m_Theme = New_Theme
    Call GetThemeProgressBar
    PropertyChanged "Theme"
End Property

Private Function TranslateColor(ByVal lColor As Long) As Long
    'System color code to long rgb
    On Error GoTo TranslateColor_Error
    
    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If
    Exit Function
    
TranslateColor_Error:
End Function

Private Sub UserControl_Initialize()
    Dim fnt As New StdFont
    
    Set Font = fnt
    
    With UserControl
        .BackColor = vbWhite
        .ScaleMode = vbPixels
    End With
    '----------------------------------------------------------
    'Default Values
    hDC = UserControl.hDC
    hwnd = UserControl.hwnd
    m_Max = 100
    m_Min = 0
    m_Value = 0
    m_Orientation = ccOrientationHorizontal
    m_Scrolling = ccScrollingStandard
    m_Color = GetLngColor(vbHighlight)
    DrawProgressBar
    '----------------------------------------------------------
End Sub

Private Sub UserControl_Paint()
    DrawProgressBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set Font = .ReadProperty("Font", Font)
        m_Brush = .ReadProperty("BrushStyle", 4)
        Color = .ReadProperty("Color", vbHighlight)
        Set m_Picture = .ReadProperty("Image", Nothing)
        Max = .ReadProperty("Max", 100)
        Min = .ReadProperty("Min", 0)
        Orientation = .ReadProperty("Orientation", ccOrientationHorizontal)
        Scrolling = .ReadProperty("Scrolling", ccScrollingStandard)
        ShowText = .ReadProperty("ShowText", False)
        m_Theme = .ReadProperty("Theme", [upbAuto])
        Value = .ReadProperty("Value", 0)
    End With
    If (Ambient.UserMode) Then                                                      'If we're not in design mode
        'Add the messages that we're interested in
        With UserControl
            '   Start Subclassing using our Handle
            Call Subclass_Start(.hwnd)
            Call Subclass_AddMsg(.hwnd, WM_SYSCOLORCHANGE)
            Call Subclass_AddMsg(.hwnd, WM_THEMECHANGED)
        End With
        bSubClass = True
    End If
End Sub

Private Sub UserControl_Resize()
    hDC = UserControl.hDC
    DrawProgressBar
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Catch
    pDestroy 'Destroy Temp DC
    If bSubClass Then
        'Stop all subclassing
        Call Subclass_StopAll
        '   Set our Flag that were done....
        bSubClass = False
    End If
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Font", Font)
        Call .WriteProperty("BrushStyle", m_Brush, 4)
        Call .WriteProperty("Color", m_Color, vbHighlight)
        Call .WriteProperty("Image", m_Picture, Nothing)
        Call .WriteProperty("Max", m_Max, 100)
        Call .WriteProperty("Min", m_Min, 0)
        Call .WriteProperty("Orientation", m_Orientation, ccOrientationHorizontal)
        Call .WriteProperty("Scrolling", m_Scrolling, ccScrollingStandard)
        Call .WriteProperty("ShowText", m_ShowText, False)
        Call .WriteProperty("Theme", m_Theme, [upbAuto])
        Call .WriteProperty("Value", m_Value, 0)
    End With
End Sub

Public Property Get Value() As Long
    Value = ((m_Value / 100) * m_Max) / IIf(m_Min > 0, m_Min, 1)
End Property

Public Property Let Value(ByVal cValue As Long)
    If m_Max = 0 Then Exit Property
    m_Value = ((cValue * 100) / m_Max) + m_Min
    PropertyChanged "Value"
    DrawProgressBar
End Property

Public Property Get Version(Optional ByVal bDateTime As Boolean = False) As String
    If bDateTime Then
        Version = Major & "." & Minor & "." & Revision & "(" & DateTime & ")"
    Else
        Version = Major & "." & Minor & "." & Revision
    End If
End Property

