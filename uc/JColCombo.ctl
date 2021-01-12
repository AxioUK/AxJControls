VERSION 5.00
Begin VB.UserControl axJColCombo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   Picture         =   "JColCombo.ctx":0000
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   173
   ToolboxBitmap   =   "JColCombo.ctx":3B8E
   Begin AxJControls.axJList PicList 
      Height          =   480
      Left            =   1320
      TabIndex        =   1
      Top             =   900
      Width           =   480
      _extentx        =   847
      _extenty        =   847
      headerh         =   24
      linecolor       =   15790320
      gridstyle       =   3
      striped         =   -1  'True
      stripedcolor    =   16645629
      selcolor        =   -2147483635
      itemh           =   0
      bordercolor     =   11709605
      header          =   0   'False
      forecolor       =   0
      visiblerows     =   10
      dropwidth       =   0
      font            =   "JColCombo.ctx":3EA0
      backcolor       =   16777215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   1080
   End
   Begin VB.TextBox Edit 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F4FA&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   375
   End
End
Attribute VB_Name = "axJColCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

' Determines if the control's parent form/window is an MDI child window
'Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32.dll" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'Private Const GWL_EXSTYLE    As Long = -20
Private Const WS_EX_MDICHILD As Long = &H40&

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As tTrackMouseEvent) As Long ' Win98 or later
Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As tTrackMouseEvent) As Long ' Win95 w/ IE 3.0
'Private Declare Function GetCapture Lib "user32.dll" () As Long
'Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
'Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long

'/Subclassing [?]
Private Const TME_LEAVE             As Long = &H2
Private Const WM_ACTIVATE           As Long = &H6
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_NCACTIVATE         As Long = &H86
Private Const WM_GETMINMAXINFO      As Long = &H24
Private Const WM_WINDOWPOSCHANGED   As Long = &H47
Private Const WM_WINDOWPOSCHANGING  As Long = &H46
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_SIZE               As Long = &H5
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_SETFOCUS           As Long = &H7
Private Const WM_KILLFOCUS          As Long = &H8
Private Const WM_MOVE               As Long = &H3
Private Const WM_TIMER              As Long = &H113
Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_MOUSEHOVER         As Long = &H2A1
Private Const WM_SYSCOMMAND         As Long = &H112

'?Edit SubClass
Private Const WM_CHAR                   As Long = &H102
Private Const WM_GETTEXT                As Long = &HD
Private Const WM_GETTEXTLENGTH          As Long = &HE
Private Const WM_SETTEXT                As Long = &HC
Private Const WM_CLEAR                  As Long = &H303
Private Const WM_CUT                    As Long = &H300
Private Const WM_PASTE                  As Long = &H302
Private Const WM_UNDO                   As Long = &H304

Private Const EM_GETSEL                 As Long = &HB0
Private Const EM_SETSEL                 As Long = &HB1
    
Private Type tTrackMouseEvent
    cbSize      As Long
    dwFlags     As Long
    hwndTrack   As Long
    dwHoverTime As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_TOOLWINDOW  As Long = &H80&
Private Const SWP_SHOWWINDOW As Long = &H40

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
'Private Declare Function GetFocus Lib "user32" () As Long

'Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'/Transparent Areas
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const RGN_OR As Long = 2

'/Render Strech
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'\Blend Color
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Type UcsRgbQuad
    r                       As Byte
    G                       As Byte
    B                       As Byte
    A                       As Byte
End Type

'/ImageList
'Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
'Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
'Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal flags As Long) As Long
'Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
'Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
'Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
'Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
'Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long

Private Const ILD_BLEND50       As Long = &H4
Private Const ILD_BLEND25       As Long = &H2
Private Const ILD_TRANSPARENT   As Long = &H1
Private Const CLR_NONE          As Long = &HFFFFFFFF
Private Const CLR_DEFAULT       As Long = &HFF000000

'/Text Ansi/Unicode
'/Draw Text
'Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_CENTER             As Long = &H1
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORD_ELLIPSIS      As Long = &H40000
Private Const DT_LEFT               As Long = &H0
Private Const DT_SINGLELINE         As Long = &H20
Private Const DT_BOTTOM             As Long = &H8
Private Const DT_FLAG               As Long = DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long

' para saber si el puntero se encuentra dentro de un rectángulo ( para las opciones del menú )
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'/Selecttion
'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'/Hand Cursor
Private Type GUID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7)    As Byte
End Type

Private Type PicBmp
    Size        As Long
    type        As Long
    hBmp        As Long
    hPal        As Long
    Reserved    As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (pDicDesc As PicBmp, RefIId As GUID, ByVal fPictureOwnsHandle As Long, lPic As IPicture) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Const IDC_HAND          As Long = 32649

Enum JComboShapes
    JcsRoundedRectangle
    JcsRectangle
    JcsCutLeft
    JcsCutRight
    JcsCutTop
    JcbsCutBottom
End Enum

Public Enum JComboListStyle
    JclsCombo
    JclsList
End Enum

Private Type tComboItem
    Text        As String
    Image       As Integer
    Key         As String
    Data        As Long
    Tag         As String
    ForeColor   As Long
    FontBold    As Boolean
End Type

'?Events
Event Click()
Event DblClick()
Event ItemClick(ByVal Item As Long)
Event SelectedIndexChanged(ByVal Item As Long)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event Scrolling()

'?Propertys
Private m_stdSkin       As StdPicture
Private m_lstStyle      As JComboListStyle
Private m_ForeColor     As OLE_COLOR
Private m_LForeColor    As OLE_COLOR
Private m_BackColor     As OLE_COLOR
Private m_LBackColor    As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_SelColor      As OLE_COLOR
Private m_Shape         As JComboShapes
Private m_sFocus        As Boolean
Private m_sdBack        As Boolean
Private m_HandCur       As Boolean
Private m_aComplete     As Boolean
Private m_dImage        As Integer
Private m_hintText      As String
Private m_ColumnInBox   As Integer
Private m_Header        As Boolean

'?Run
Private cSubClass       As c_SubClass
'Private tItem()         As tComboItem
'Private tItemR()        As RECT
'Private tTextR()        As RECT
'Private tImgR()         As RECT
Private tIndex          As Integer
Private tHIndex         As Integer
Private c_hSkin         As Long
'Private c_hIml          As Long
'Private c_ImgX          As Long
'Private c_ImgY          As Long
Private m_leState       As Integer
Private m_TR            As RECT
Private m_RctDrop       As RECT
Private sText           As String

Private m_ItemH         As Integer
Private m_TextH         As Integer
Private m_VisibleItems  As Integer
Private bcResizeFlag    As Boolean
Private m_lImage        As Integer
Private m_fIndex        As Integer
Private m_bHint         As Boolean

Private m_HasFocus        As Boolean
Private m_PhWnd           As Long
Private m_bIsTracking     As Boolean
Private m_bTrackHandler32 As Boolean
Private m_bSuppMouseTrack As Boolean



Public Sub AddColumn(ByVal Text As String, Optional ByVal Width As Long = 100, Optional ByVal Alignment As AlignmentConstants)
   
   PicList.AddColumn Text, Width, Alignment

End Sub

Public Function AddItem(ByVal Text As String, Optional ByVal IconIndex As Long = -1, Optional ByVal ItemData As Long, Optional ByVal ItemTag As String = "") As Long
    
    AddItem = PicList.AddItem(Text, IconIndex, ItemData, ItemTag)

End Function

Public Sub RemoveItem(ByVal Index As Long)

    PicList.RemoveItem Index
    
End Sub

Public Function pGetSystemHandCursor() As Picture
Dim Pic             As PicBmp
Dim IPic            As IPicture
Dim IID_IDispatch   As GUID
Dim hCur            As Long
        
        hCur = LoadCursor(ByVal 0&, IDC_HAND)
        With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        With Pic
            .Size = Len(Pic)
            .type = vbPicTypeIcon
            .hBmp = hCur
            .hPal = 0
        End With
        Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
        Set pGetSystemHandCursor = IPic
        DestroyCursor hCur
End Function

Public Sub ShowDropDown(Optional ByVal Visible As Boolean)
    If PicList.Visible = Visible Then Visible = Not Visible
    pShowList Visible
End Sub

'/Determina si la Funcion es Soportada por la Libreria
Private Function IsFunctionSupported(sFunction As String, sModule As String) As Boolean
Dim hModule As Long

    ' GetModuleHandle?
    hModule = GetModuleHandleA(sModule)
    If (hModule = 0) Then
        hModule = LoadLibrary(sModule)
    End If
    
    If (hModule) Then
        If (GetProcAddress(hModule, sFunction)) Then
            IsFunctionSupported = True
        End If
        FreeLibrary hModule
    End If
End Function


Private Function pBlendColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore         As UcsRgbQuad
    Dim clrBack         As UcsRgbQuad
 
    OleTranslateColor clrFirst, 0, VarPtr(clrFore)
    OleTranslateColor clrSecond, 0, VarPtr(clrBack)
    With clrFore
        .r = (.r * lAlpha + clrBack.r * (255 - lAlpha)) / 255
        .G = (.G * lAlpha + clrBack.G * (255 - lAlpha)) / 255
        .B = (.B * lAlpha + clrBack.B * (255 - lAlpha)) / 255
    End With
    CopyMemory VarPtr(pBlendColor), VarPtr(clrFore), 4
End Function

Private Sub pCreateRegions(Optional EllipseW As Long = 3, Optional EllipseH As Long = 3)

    With UserControl
        Dim hRgn  As Long
        Dim hRgn2 As Long
        
        If m_Shape = 1 Then
            hRgn = CreateRoundRectRgn(0, 0, .ScaleWidth + 1, .ScaleHeight + 1, 0, 0)
        Else
            hRgn = CreateRoundRectRgn(0, 0, .ScaleWidth + 1, .ScaleHeight + 1, EllipseW, EllipseH)
        End If
        
        Select Case m_Shape
            Case 0 'RoundedRectangle
            Case 1 'Rectangle
            Case 2 'CutLeft
                hRgn2 = CreateRectRgn(0, 0, .ScaleWidth / 2, .ScaleHeight + 1)
                CombineRgn hRgn, hRgn, hRgn2, RGN_OR
            Case 3 'CutRight
                hRgn2 = CreateRectRgn(.ScaleWidth / 2, 0, .ScaleWidth + 1, .Height + 1)
                CombineRgn hRgn, hRgn, hRgn2, RGN_OR
            Case 4 'CutTop
                hRgn2 = CreateRectRgn(0, 0, .ScaleWidth + 1, .ScaleHeight / 2)
                CombineRgn hRgn, hRgn, hRgn2, RGN_OR
            Case 5 'CutBottom
                hRgn2 = CreateRectRgn(0, .ScaleHeight / 2, .ScaleWidth + 1, .Height + 1)
                CombineRgn hRgn, hRgn, hRgn2, RGN_OR
        End Select
        
        DeleteObject hRgn2
        SetWindowRgn .hwnd, hRgn, True
        DeleteObject hRgn
    End With
    
End Sub

Private Sub pDrawControl(ByVal eState As Integer, Optional Force As Boolean)
On Error Resume Next
Dim j As Integer
Dim lPx As Integer
Dim lL As Integer
Dim lT  As Integer
Dim hColor As Long
Dim TR      As RECT

    If eState = m_leState And Not Force Then Exit Sub
    If m_stdSkin Is Nothing Then Set m_stdSkin = UserControl.Picture
    If Not c_hSkin Then pSelectHSkin m_stdSkin.Handle
    
    If eState = 0 And m_HasFocus Then eState = 3
    If Not UserControl.Enabled Then eState = 4
    
    With UserControl
       .Cls
       lPx = eState * 15
        .BackColor = m_BackColor
        If Edit.BackColor <> .BackColor Then Edit.BackColor = .BackColor
        Select Case m_lstStyle
            Case 0 '?Combo
            
                    lL = (.ScaleWidth - 16)
                    'If eState = 2 Or eState = 3 Then .ForeColor = GetPixel(c_hSkin, 15, 0) Else .ForeColor = GetPixel(c_hSkin, lPx, 0)
                    
                    .ForeColor = GetPixel(c_hSkin, lPx, 0)
                    RoundRect .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0
                    If eState = 3 Then
                        .ForeColor = pBlendColor(GetPixel(c_hSkin, lPx + 1, 1), .BackColor, 50)
                        RoundRect .hdc, 1, 1, .ScaleWidth - 1, .ScaleHeight - 1, 2, 2
                    End If
                    If m_sdBack Then
                        pRenderStretch .hdc, lL, 0, 16, .ScaleHeight, c_hSkin, lPx, 0, 15, 23, 3
                    ElseIf eState = 1 Or eState = 2 Then
                        pRenderStretch .hdc, lL, 0, 16, .ScaleHeight, c_hSkin, lPx, 0, 15, 23, 3
                    End If
                    
'                    If c_hIml Then
'                        If tIndex <> -1 Then
'                            If m_lImage <> tItem(tIndex).Image Then m_lImage = tItem(tIndex).Image
'                        End If
'                        ImageList_Draw c_hIml, m_lImage, .hdc, 3, ((.ScaleHeight - c_ImgY) \ 2), 0
'                    End If
                    
                    If Enabled Then
                        .ForeColor = m_ForeColor
                    ElseIf m_bHint Then
                        .ForeColor = &H808080
                    Else
                        .ForeColor = pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50)
                    End If
                    
            Case 1 '?List
                    CopyRect TR, m_TR
                    
                    If tIndex <> -1 Then
                        .ForeColor = m_ForeColor
                    Else
                        If Trim(m_hintText) <> "" Then
                            sText = m_hintText
                            .ForeColor = pBlendColor(&H808080, GetPixel(c_hSkin, lPx, 0), 50)
                        End If
                    End If
                    If Not .Enabled Then .ForeColor = pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50)
                    pRenderStretch .hdc, 0, 0, .ScaleWidth, .ScaleHeight, c_hSkin, lPx, 0, 15, 23, 3
                    
                     If eState = 2 Then OffsetRect TR, 0, 1
                    DrawText .hdc, sText, Len(sText), TR, DT_FLAG
                    ' If c_hIml And tItem(tIndex).Image <> -1 Then ImageList_Draw c_hIml, tItem(tIndex).Image, .hdc, 3, ((.ScaleHeight - c_ImgY) / 2) + IIf(eState = 2, 1, 0), 0
        End Select
        
        'Drop Arrow
                lL = (.ScaleWidth - 16) + ((15 - 4) / 2)
                lT = (.ScaleHeight - 3) / 2
                If m_sFocus And m_lstStyle = 1 Then lL = lL - 1
                If eState = 2 Then lT = lT + 1
        
                hColor = pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50)
                
                SetPixel .hdc, lL, lT, hColor
                SetPixel .hdc, lL + 1, lT, hColor
                SetPixel .hdc, lL + 2, lT, hColor
                SetPixel .hdc, lL + 3, lT, hColor
                SetPixel .hdc, lL + 4, lT, hColor
                
                SetPixel .hdc, lL + 1, lT + 1, hColor
                SetPixel .hdc, lL + 2, lT + 1, hColor
                SetPixel .hdc, lL + 3, lT + 1, hColor
                
                SetPixel .hdc, lL + 2, lT + 2, hColor
                If m_lstStyle = 1 Then UserControl.Line (.ScaleWidth - 16, 7 + IIf(eState = 2, 1, 0))-(.ScaleWidth - 16, .ScaleHeight - 7 + IIf(eState = 2, 1, 0)), GetPixel(c_hSkin, lPx, 0), B
        
        Select Case m_Shape
          Case 0 'RoundedRectangle
                SetPixelV .hdc, 1, 1, GetPixel(c_hSkin, lPx, 0)
                SetPixelV .hdc, 1, .ScaleHeight - 2, GetPixel(c_hSkin, lPx, 22)
                SetPixelV .hdc, .ScaleWidth - 2, 1, GetPixel(c_hSkin, lPx + 14, 0)
                SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, GetPixel(c_hSkin, lPx + 14, 22)
            Case 1 'Rectangle
            Case 2 'CutLeft
                SetPixelV .hdc, .ScaleWidth - 2, 1, GetPixel(c_hSkin, lPx + 14, 0)
                SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, GetPixel(c_hSkin, lPx + 14, 22)
            Case 3 'CutRight
                SetPixelV .hdc, 1, 1, GetPixel(c_hSkin, lPx, 0)
                SetPixelV .hdc, 1, .ScaleHeight - 2, GetPixel(c_hSkin, lPx, 22)
            Case 4 'CutTop
                SetPixelV .hdc, 1, .ScaleHeight - 2, GetPixel(c_hSkin, lPx, 22)
                SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, GetPixel(c_hSkin, lPx + 14, 22)
            Case 5 'CutBottom
                SetPixelV .hdc, 1, 1, GetPixel(c_hSkin, lPx, 0)
                SetPixelV .hdc, .ScaleWidth - 2, 1, GetPixel(c_hSkin, lPx + 14, 0)
        End Select
        
        '?FocusRect
        If m_sFocus And m_HasFocus And m_lstStyle = 1 Then
            For j = 4 To .ScaleWidth - (4) Step 2
                SetPixelV .hdc, j, 3, pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50) 'm_tBtnColors(10)
                SetPixelV .hdc, j, .ScaleHeight - 4, pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50) 'm_tBtnColors(10)
            Next
            For j = 4 To .ScaleHeight - 4 Step 2
                SetPixelV .hdc, 3, j, pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50) 'm_tBtnColors(10)
                SetPixelV .hdc, .ScaleWidth - (4), j, pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50)  'm_tBtnColors(10)
            Next
        End If
        
        If eState = 3 And m_HasFocus Then eState = 0
        m_leState = eState
        If Edit.ForeColor <> .ForeColor Then Edit.ForeColor = .ForeColor
        
        PicList.Header = m_Header
    End With
End Sub

Private Function pMouseOnHandle(hwnd As Long) As Boolean
    Dim PT As POINTAPI
    GetCursorPos PT
    pMouseOnHandle = (WindowFromPoint(PT.X, PT.Y) = hwnd)
End Function

Private Function pRenderStretch(ByVal destDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal DestW As Long, ByVal DestH As Long, ByVal SrcDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Size As Long, Optional MaskColor As Long = -1)
Dim Sx2 As Long
Sx2 = Size * 2

    If MaskColor <> -1 Then
        Dim mDC         As Long
        Dim mX          As Long
        Dim mY          As Long
        Dim DC          As Long
        Dim hBmp        As Long
        Dim hOldBmp     As Long
     
        mDC = destDC: DC = GetDC(0)
        destDC = CreateCompatibleDC(0)
        hBmp = CreateCompatibleBitmap(DC, DestW, DestH)
        hOldBmp = SelectObject(destDC, hBmp) ' save the original BMP for later reselection
        mX = destX: mY = destY
        destX = 0: destY = 0
    End If
 
        SetStretchBltMode destDC, vbPaletteModeNone
         
        BitBlt destDC, destX, destY, Size, Size, SrcDC, X, Y, vbSrcCopy  'TOP_LEFT
        StretchBlt destDC, destX + Size, destY, DestW - Sx2, Size, SrcDC, X + Size, Y, Width - Sx2, Size, vbSrcCopy 'TOP_CENTER
        BitBlt destDC, destX + DestW - Size, destY, Size, Size, SrcDC, X + Width - Size, Y, vbSrcCopy 'TOP_RIGHT
        StretchBlt destDC, destX, destY + Size, Size, DestH - Sx2, SrcDC, X, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_LEFT
        StretchBlt destDC, destX + Size, destY + Size, DestW - Sx2, DestH - Sx2, SrcDC, X + Size, Y + Size, Width - Sx2, Height - Sx2, vbSrcCopy 'MID_CENTER
        StretchBlt destDC, destX + DestW - Size, destY + Size, Size, DestH - Sx2, SrcDC, X + Width - Size, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_RIGHT
        BitBlt destDC, destX, destY + DestH - Size, Size, Size, SrcDC, X, Y + Height - Size, vbSrcCopy 'BOTTOM_LEFT
        StretchBlt destDC, destX + Size, destY + DestH - Size, DestW - Sx2, Size, SrcDC, X + Size, Y + Height - Size, Width - Sx2, Size, vbSrcCopy   'BOTTOM_CENTER
        BitBlt destDC, destX + DestW - Size, destY + DestH - Size, Size, Size, SrcDC, X + Width - Size, Y + Height - Size, vbSrcCopy 'BOTTOM_RIGHT

    If MaskColor <> -1 Then
        GdiTransparentBlt mDC, mX, mY, DestW, DestH, destDC, 0, 0, DestW, DestH, MaskColor
        SelectObject destDC, hOldBmp
        DeleteObject hBmp
        ReleaseDC 0&, DC
        DeleteDC destDC
    End If
End Function

Private Sub pSelectHSkin(Optional lHandle As Long = 0)
Dim j As Integer
    If c_hSkin Then Call DeleteDC(c_hSkin)
    c_hSkin = CreateCompatibleDC(0)
    Call SelectObject(c_hSkin, lHandle)
End Sub

Private Sub pShowList(ByVal Visible As Boolean)
Dim lW As Long
Dim Rct As RECT
Dim PT As POINTAPI
Dim lstTop  As Integer


    If Visible Then
        GetWindowRect UserControl.hwnd, Rct
        SetParent PicList.hwnd, 0
        If Rct.Bottom + PicList.ScaleHeight > Screen.Height / Screen.TwipsPerPixelY Then
          lstTop = Rct.Top - (PicList.ScaleHeight + 1)
        Else
          lstTop = Rct.Bottom + 1
        End If
        SetWindowPos PicList.hwnd, 0, Rct.Left, lstTop, UserControl.ScaleWidth, PicList.ItemHeight * PicList.VisibleRows, SWP_SHOWWINDOW
        PicList.Visible = True
        PicList.Refresh
    Else
        SetParent PicList.hwnd, UserControl.hwnd
        PicList.Visible = False
        If Timer1.Enabled Then Timer1.Enabled = False
        'tHIndex = -1
    End If
End Sub

Private Sub pUseHint(ByVal Value As Boolean)
    If Trim(m_hintText) = "" Then Exit Sub
    If Value Then
        m_bHint = True
        Edit.Text = m_hintText
        Edit.ForeColor = vbRed '&H808080
    Else
        Edit.Text = ""
        Edit.ForeColor = m_ForeColor
        m_bHint = False
    End If
End Sub


'/Start tracking of mouse leave event
Private Sub TrackMouseTracking(hwnd As Long)
Dim tEventTrack As tTrackMouseEvent
    
    With tEventTrack
        .cbSize = Len(tEventTrack)
        .dwFlags = TME_LEAVE
        .hwndTrack = hwnd
    End With
    If (m_bTrackHandler32) Then
        TrackMouseEvent tEventTrack
    Else
        TrackMouseEvent2 tEventTrack
    End If
End Sub

'?TextField
Private Sub Edit_Change()
    If Not m_bHint Then
        RaiseEvent Change
    'Else
        'Debug.Print "Uhint"
    End If
End Sub

Private Sub Edit_Click()
   If PicList.Visible Then pShowList False
End Sub

Private Sub Edit_GotFocus()
    If m_bHint Then
        pUseHint False
    Else
        Edit.SelStart = 0: Edit.SelLength = Len(Edit)
    End If
End Sub

Private Sub PicList_ItemClick(Item As Long)
Dim TextValue As String

TextValue = PicList.ItemText(Item, m_ColumnInBox)

If m_lstStyle = 0 Then
  Edit.Text = TextValue
Else
  sText = TextValue
  pDrawControl 2
End If

RaiseEvent ItemClick(Item)
End Sub

Private Sub Timer1_Timer()
  If Not pMouseOnHandle(PicList.hwnd) Then
        Timer1.Enabled = False
        tHIndex = -1
       
    End If
    DoEvents
End Sub

Private Sub UserControl_DblClick()
Dim PT As POINTAPI
    
    RaiseEvent DblClick

    GetCursorPos PT
    If WindowFromPoint(PT.X, PT.Y) = UserControl.hwnd Then
        ScreenToClient UserControl.hwnd, PT
        If PtInRect(m_RctDrop, PT.X, PT.Y) Then
            pDrawControl 2
            pShowList Not PicList.Visible
        End If
    End If
    
End Sub

'?UserControl
Private Sub UserControl_EnterFocus()
    m_HasFocus = True
    pDrawControl m_leState, True
End Sub

Private Sub UserControl_ExitFocus()
    pDrawControl m_leState, True
End Sub
Private Sub UserControl_Initialize()
    Set cSubClass = New c_SubClass
    tIndex = -1
    tHIndex = -1
    m_lImage = -1
    m_fIndex = -1
End Sub


Private Sub UserControl_InitProperties()
    Set m_stdSkin = UserControl.Picture
    m_BackColor = vbWhite
    m_LBackColor = vbWhite
    m_BorderColor = &HAEAEAE
    m_SelColor = &HFF6600
    m_ColumnInBox = 0
    
    UserControl.Enabled() = True
    PicList.Init UserControl.hwnd
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    'Debug.Print Chr(KeyAscii)
    'Debug.Print GetEditText
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PtInRect(m_RctDrop, X, Y) Then
            pDrawControl 2
            'pShowList Not PicList.Visible
            If PicList.Visible Then
                pShowList False
            Else
                pShowList True
            End If
    Else
        If PicList.Visible Then pShowList False
    End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
   If PtInRect(m_RctDrop, X, Y) Then
        If (Not m_bIsTracking) Then
                m_bIsTracking = True
                TrackMouseTracking UserControl.hwnd
            End If
            If Button = 1 Then pDrawControl 2 Else pDrawControl 1
            
           ' If UserControl.MousePointer <> 0 Then UserControl.MousePointer = 0
           If m_HandCur And UserControl.MousePointer <> vbCustom Then UserControl.MousePointer = vbCustom
    Else
        pDrawControl 0
        
        If UserControl.MousePointer <> vbNormal Then UserControl.MousePointer = vbNormal
    
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PtInRect(m_RctDrop, X, Y) Then
            If Button = 1 Then pDrawControl 1
    Else
        If Button = 1 And Not pMouseOnHandle(UserControl.hwnd) Then
            pDrawControl 0
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set m_stdSkin = .ReadProperty("SkinPicture", UserControl.Picture)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        Set PicList.Font = .ReadProperty("lFont", PicList.Font)
        
        m_lstStyle = .ReadProperty("lstStyle", 0)
        m_ForeColor = .ReadProperty("ForeColor", 0)
        m_LForeColor = .ReadProperty("ForeColorList", 0)
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        m_LBackColor = .ReadProperty("BackColorList", vbWhite)
        m_BorderColor = .ReadProperty("BorderColor", &HAEAEAE)
        m_SelColor = .ReadProperty("SelectColor", &HFF6600)
        m_Shape = .ReadProperty("Shape", 0)
        m_sdBack = .ReadProperty("sdBack", True)
        m_sFocus = .ReadProperty("sFocus", True)
        Edit.Text = .ReadProperty("Text", "")
        m_HandCur = .ReadProperty("HandCursor", False)
        m_aComplete = .ReadProperty("aComplete", False)
        m_hintText = .ReadProperty("Hint", "")
        m_ColumnInBox = .ReadProperty("ColumnInBox", 0)
        m_Header = .ReadProperty("Header", False)
        
        UserControl.Enabled() = .ReadProperty("Enabled", True)
    End With
    
        '?Subclass
        If Ambient.UserMode Then
                m_bTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
                m_bSuppMouseTrack = m_bTrackHandler32
                If Not m_bSuppMouseTrack Then m_bSuppMouseTrack = IsFunctionSupported("_TrackMouseEvent", "Comctl32")
                
                m_PhWnd = UserControl.Parent.hwnd
                
                With cSubClass
                        If .Subclass(UserControl.hwnd, , , Me) Then
                            If m_bSuppMouseTrack Then .AddMsg UserControl.hwnd, WM_MOUSELEAVE, MSG_AFTER
                            .AddMsg UserControl.hwnd, WM_KILLFOCUS, MSG_AFTER
                            .AddMsg UserControl.hwnd, WM_MOUSEWHEEL, MSG_AFTER
                            
                       End If
                       If .Subclass(m_PhWnd, , , Me) Then
                         .AddMsg m_PhWnd, WM_SIZE, MSG_AFTER
                         .AddMsg m_PhWnd, WM_MOVE, MSG_AFTER
                         .AddMsg m_PhWnd, WM_WINDOWPOSCHANGING, MSG_AFTER
                         .AddMsg m_PhWnd, 516, MSG_BEFORE ' MouseDown
                         .AddMsg m_PhWnd, 513, MSG_BEFORE ' MouseUp
                         .AddMsg m_PhWnd, 164, MSG_BEFORE ' Menu
                         .AddMsg m_PhWnd, WM_SYSCOMMAND, MSG_BEFORE
                        End If
                        If .Subclass(Edit.hwnd, , , Me) Then
                            .AddMsg Edit.hwnd, WM_KILLFOCUS, MSG_AFTER
                            .AddMsg Edit.hwnd, WM_MOUSEWHEEL, MSG_AFTER
                            
'                            '?AutoComplete
'                            .AddMsg Edit.hwnd, WM_CHAR, MSG_BEFORE_AFTER
'                            .AddMsg Edit.hwnd, WM_CLEAR, MSG_AFTER
'                            .AddMsg Edit.hwnd, WM_CUT, MSG_AFTER
'                            .AddMsg Edit.hwnd, WM_PASTE, MSG_AFTER
'                            .AddMsg Edit.hwnd, WM_UNDO, MSG_AFTER
                        End If
                End With
                
                SetWindowLongA PicList.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
        End If
    
    Set Edit.Font = UserControl.Font
    If Trim(Edit.Text) = "" And Trim(m_hintText) <> "" Then
       pUseHint True
    End If
      
    If m_HandCur Then
      UserControl.MouseIcon = pGetSystemHandCursor
      UserControl.MousePointer = vbCustom
    End If
    
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
Dim TextLeft As Integer
Dim TextTop As Integer
Dim TextW   As Integer
Dim CtrlH   As Integer

    If bcResizeFlag Then Exit Sub
        bcResizeFlag = True
    With UserControl
        m_TextH = .TextHeight("ÁjqWJ")
    
'        If c_hIml Then
'            CtrlH = c_ImgY + 4
'            If m_TextH > c_ImgY Then CtrlH = m_TextH + 4
'        Else
          CtrlH = m_TextH + 4
'        End If
        
        If CtrlH < 15 Then CtrlH = 15
        If .ScaleHeight < CtrlH Then UserControl.Height = CtrlH * 15
    
        Edit.Height = m_TextH + 2
        Select Case m_lstStyle
                Case 0
'                 If c_hIml Then
'                      Edit.Move c_ImgX + 5, (.ScaleHeight / 2) - (m_TextH / 2), .ScaleWidth - (c_ImgX + 6 + 16), m_TextH
'
'                 Else
                      Edit.Move 4, (.ScaleHeight - m_TextH) / 2, .ScaleWidth - (21), m_TextH
          
'                 End If
                 SetRect m_RctDrop, .ScaleWidth - 16, 0, ScaleWidth, .ScaleHeight
                        
                Case 1
                
                 SetRect m_RctDrop, 0, 0, ScaleWidth, .ScaleHeight
                 'SetRect sTextRect, Edit.Left, Edit.Top, .ScaleWidth - (c_ImgX + 22), m_TextH
                 
        End Select
            
            TextLeft = 5 ' TextLeft = IIf(c_hIml > 0, c_ImgX + 7, 5)
            TextTop = (.ScaleHeight - m_TextH) \ 2
            TextW = .ScaleWidth - (21)  'TextW = .ScaleWidth - (18 + IIf(c_hIml, c_ImgX + 6, 3))
            SetRect m_TR, TextLeft, TextTop, TextLeft + TextW, TextTop + m_TextH
            
            PicList.Width = UserControl.ScaleWidth
            
    End With
    pCreateRegions
    pDrawControl m_leState, True
    bcResizeFlag = False
End Sub

Private Sub UserControl_Terminate()
  cSubClass.UnSubclass UserControl.hwnd
  cSubClass.UnSubclass Edit.hwnd
  cSubClass.UnSubclass m_PhWnd

  Set cSubClass = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "SkinPicture", m_stdSkin, UserControl.Picture
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "lFont", PicList.Font
        
        .WriteProperty "lstStyle", m_lstStyle
        .WriteProperty "ForeColor", m_ForeColor
        .WriteProperty "ForeColorList", m_LForeColor
        .WriteProperty "BackColor", m_BackColor
        .WriteProperty "BackColorList", m_LBackColor
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "SelectColor", m_SelColor
        .WriteProperty "Shape", m_Shape
        .WriteProperty "sdBack", m_sdBack
        .WriteProperty "sFocus", m_sFocus
        .WriteProperty "Text", Edit.Text
        .WriteProperty "HandCursor", m_HandCur
        .WriteProperty "aComplete", m_aComplete
        .WriteProperty "Hint", m_hintText
        .WriteProperty "ColumnInBox", m_ColumnInBox, 0
        .WriteProperty "Header", m_Header, False
        
        .WriteProperty "Enabled", UserControl.Enabled, True
    End With
End Sub

Property Get AutoComplete() As Boolean: AutoComplete = m_aComplete: End Property
Property Let AutoComplete(NewProp As Boolean)
    m_aComplete = NewProp
    PropertyChanged "aComplete"
End Property

Property Get BackColor() As OLE_COLOR: BackColor = m_BackColor: End Property
Property Let BackColor(NewColor As OLE_COLOR)
        m_BackColor = NewColor
        PropertyChanged "BackColor"
        pDrawControl m_leState, True
End Property

Property Get BackColorList() As OLE_COLOR: BackColorList = m_LBackColor: End Property
Property Let BackColorList(NewColor As OLE_COLOR)
        m_LBackColor = NewColor
        PropertyChanged "BackColorList"
        
End Property

Property Get BorderColorList() As OLE_COLOR: BorderColorList = m_BorderColor: End Property
Property Let BorderColorList(ByVal New_Color As OLE_COLOR)
    m_BorderColor = New_Color
    Call PropertyChanged("BorderColor")
    
End Property

Property Get Enabled() As Boolean: Enabled = UserControl.Enabled: End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call PropertyChanged("Enabled")
    pDrawControl 0, True
    If New_Enabled = False And PicList.Visible Then pShowList False
End Property

Property Get FocusRect() As Boolean: FocusRect = m_sFocus: End Property
Property Let FocusRect(newValue As Boolean)
    m_sFocus = newValue
    PropertyChanged "sFocus"
    If m_HasFocus Then pDrawControl m_leState, True
End Property

Property Get FontList() As StdFont: Set FontList = PicList.Font: End Property
Property Set FontList(NewFont As StdFont)
    Set PicList.Font = NewFont
    PropertyChanged "lFont"
    
End Property

Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Property Set Font(NewFont As StdFont)
    Set UserControl.Font = NewFont
    Set Edit.Font = UserControl.Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

Property Get ForeColor() As OLE_COLOR: ForeColor = m_ForeColor: End Property
Property Let ForeColor(NewColor As OLE_COLOR)
    m_ForeColor = NewColor
    PropertyChanged "ForeColor"
    pDrawControl m_leState, True
    
End Property

Property Get ForeColorList() As OLE_COLOR: ForeColorList = m_LForeColor: End Property
Property Let ForeColorList(NewColor As OLE_COLOR)
    m_LForeColor = NewColor
    PropertyChanged "ForeColorList"
    
End Property

Property Get HandCursor() As Boolean: HandCursor = m_HandCur: End Property
Property Let HandCursor(NewProp As Boolean)
    m_HandCur = NewProp
    PropertyChanged "HandCursor"
    If Ambient.UserMode Then
        If NewProp Then
            UserControl.MouseIcon = pGetSystemHandCursor
            UserControl.MousePointer = vbCustom
        Else
            UserControl.MousePointer = vbNormal
        End If
    End If
End Property

Public Property Get Header() As Boolean
  Header = m_Header
End Property

Public Property Let Header(ByVal NewHeader As Boolean)
  m_Header = NewHeader
  PropertyChanged "Header"
End Property

Property Get HintText() As String: HintText = m_hintText: End Property
Property Let HintText(NewProp As String)
    m_hintText = NewProp
    PropertyChanged "Hint"
    If (m_bHint Or Trim(m_hintText) <> "" Or Edit = "") And m_lstStyle <> 1 Then pUseHint True
    If m_lstStyle = 1 And PicList.SelectedIndex = -1 Then pDrawControl m_leState, True
End Property

Property Get ItemText(ByVal Item As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    ItemText = PicList.ItemText(Item, Column)
End Property
Property Let ItemText(ByVal Item As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
    PicList.ItemText(Item, Column) = Value
End Property

Property Get ColumnCount() As Long
    ColumnCount = PicList.ColumnCount
End Property

Public Property Get ColumnInBox() As Long
    ColumnInBox = m_ColumnInBox
End Property

Public Property Let ColumnInBox(ByVal NewColumnInBox As Long)
  m_ColumnInBox = NewColumnInBox
  PropertyChanged "ColumnInBox"
End Property

Public Sub ColWidthAutoSize(Optional ByVal lCol As Long = -1)
  PicList.ColWidthAutoSize lCol
End Sub

Property Get ItemCount() As Long
    ItemCount = PicList.ItemCount
End Property

Property Get ItemHeight() As Long
  ItemHeight = PicList.ItemHeight
End Property

Property Let ItemHeight(ByVal Value As Long)
  PicList.ItemHeight = Value
  PropertyChanged "ItemH"
End Property

Property Get VisibleRows() As Long: VisibleRows = PicList.VisibleRows: End Property
Property Let VisibleRows(ByVal Value As Long)
    PicList.VisibleRows = Value
    PropertyChanged "VisibleRows"
End Property

Property Get ListStyle() As JComboListStyle
    ListStyle = m_lstStyle
End Property
Property Let ListStyle(NewStyle As JComboListStyle)
    m_lstStyle = NewStyle
    PropertyChanged "lstStyle"
    UserControl_Resize
    
    Edit.Visible = m_lstStyle = 0
    
End Property

Property Get SelectionColor() As OLE_COLOR: SelectionColor = m_SelColor: End Property
Property Let SelectionColor(ByVal New_Color As OLE_COLOR)
    m_SelColor = New_Color
    Call PropertyChanged("SelectColor")
    
End Property
Public Property Get ShapeStyle() As JComboShapes: ShapeStyle = m_Shape: End Property
Public Property Let ShapeStyle(Value As JComboShapes)
    m_Shape = Value
    PropertyChanged "Shape"
    'm_bCalculateRects = True: DrawButton m_tBtnSetting.State, True
    UserControl_Resize
End Property
Property Get ShowDropBack() As Boolean: ShowDropBack = m_sdBack: End Property
Property Let ShowDropBack(NewProp As Boolean)
    m_sdBack = NewProp
    PropertyChanged "sdBack"
    pDrawControl m_leState, True
End Property
Property Get SkinPicture() As StdPicture: Set SkinPicture = m_stdSkin: End Property
Property Set SkinPicture(NewSkin As StdPicture)
    Set m_stdSkin = NewSkin
    If m_stdSkin Is Nothing Then Set m_stdSkin = UserControl.Picture
    PropertyChanged "SkinPicture"
    pSelectHSkin m_stdSkin.Handle
    pDrawControl m_leState, True
End Property

Property Get Text() As String
  Text = Edit.Text
End Property

Property Let Text(newText As String)
  Edit.Text = newText
  PropertyChanged "Text"
End Property

' Ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hwnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)
       
       
    Select Case uMsg
        Case WM_MOUSELEAVE
            m_bIsTracking = False
            If m_HandCur And UserControl.MousePointer <> vbNormal Then UserControl.MousePointer = vbNormal
            pDrawControl 0, True
            
        Case WM_SIZE, WM_MOVE, WM_WINDOWPOSCHANGING, WM_KILLFOCUS, WM_SYSCOMMAND
            If PicList.Visible Then pShowList False
            
        Case WM_MOUSEWHEEL
        
                    If PicList.Visible Then
                        
                    Else
                        
                    End If
                    
        Case 516, 164, 269, 513
             If PicList.Visible Then pShowList False
             
        Case WM_CHAR
                
        Case Else
    End Select
End Sub

