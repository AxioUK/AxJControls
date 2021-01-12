VERSION 5.00
Begin VB.UserControl axJCombo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
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
   Picture         =   "axJCombo2.ctx":0000
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ToolboxBitmap   =   "axJCombo2.ctx":2CD2
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   1080
   End
   Begin VB.PictureBox PicList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1200
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   615
      Begin VB.VScrollBar Bar 
         Height          =   450
         Left            =   360
         Max             =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox Edit 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "axJCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''-------------------------
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

' Determines if the control's parent form/window is an MDI child window
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const WS_EX_MDICHILD As Long = &H40&

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As tTrackMouseEvent) As Long ' Win98 or later
Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As tTrackMouseEvent) As Long ' Win95 w/ IE 3.0
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

''/Subclassing [?]
Private Const TME_LEAVE     As Long = &H2

Private Const EM_GETSEL     As Long = &HB0
Private Const EM_SETSEL     As Long = &HB1
    
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

'Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'/Transparent Areas
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
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

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
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
Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal flags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long

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

'/Selecttion
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

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

Enum JComboShape
    JcsRoundedRectangle
    JcsRectangle
    JcsCutLeft
    JcsCutRight
    JcsCutTop
    JcbsCutBottom
End Enum

Public Enum JCListStyle
    JclsCombo
    JclsList
    'JclsSimple
End Enum

Private Type tCol
    Text        As String
End Type

Private Type tComboItem
    Col(3)      As tCol   'AxioUK
    Text        As String
    Image       As Integer
    Key         As String
    Data        As Long
    Tag         As String
    ForeColor   As Long
    FontBold    As Boolean
End Type

Public Enum iCols
    [Column1] = 0
    [Column2] = 1
    [Column3] = 2
    [Column4] = 3
End Enum

'?Events
Event Click()
Event DblClick()
Event ItemClick(Item As Integer, strCol1 As String, strCol2 As String, strCol3 As String, strCol4 As String)
Event ListIndexChanged(ByVal Item As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event Scrolling()

Private cSubClass     As c_SubClass

'?Propertys
Private m_stdSkin     As StdPicture
Private m_lstStyle    As JCListStyle
Private m_Text        As String
Private mc_Fore       As OLE_COLOR
Private mc_lFore      As OLE_COLOR
Private mc_Back       As OLE_COLOR
Private mc_lBack      As OLE_COLOR
Private mc_Border     As OLE_COLOR
Private mc_Selection  As OLE_COLOR
Private m_Shape       As JComboShape
Private m_sFocus      As Boolean
Private m_sdBack      As Boolean
Private m_HandCur     As Boolean
Private m_aComplete   As Boolean
Private m_dImage      As Integer
Private m_hintText    As String

'?Run
Private tItem()         As tComboItem
Private tItemR()        As RECT
Private tTextR()        As RECT
Private tImgR()         As RECT
Private tIndex          As Integer
Private tHIndex         As Integer
Private m_RctDrop       As RECT
Private c_hSkin         As Long
Private c_hIml          As Long
Private c_ImgX          As Long
Private c_ImgY          As Long
Private m_leState       As Integer
Private m_TR            As RECT
Private m_ItemH         As Integer
Private m_TextH         As Integer
Private m_VisibleItems  As Integer
Private bcResizeFlag    As Boolean
Private m_lImage        As Integer
Private m_fIndex        As Integer
Private m_bHint         As Boolean

Private m_HasFocus          As Boolean
Private m_PhWnd             As Long
Private m_bIsTracking       As Boolean
Private m_bTrackHandler32   As Boolean
Private m_bSuppMouseTrack    As Boolean

'--AxioUK-------------------------------
Private m_ColumnInList As iCols
Private isCleared As Boolean


Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
Dim sStr As String
sStr = "axJColCombo v2" & vbCrLf
sStr = sStr & "Original: jCombo UC by J.Elihu" & vbCrLf
sStr = sStr & "Modded by AxioUK"

    MsgBox sStr
End Sub

Private Sub UserControl_InitProperties()

    Set m_stdSkin = UserControl.Picture
    mc_Back = vbWhite
    mc_lBack = vbWhite
    mc_Border = &HAEAEAE
    mc_Selection = &HFF6600
    m_ColumnInList = 0
    
    UserControl.Enabled() = True
End Sub
Private Sub UserControl_Initialize()
    Set cSubClass = New c_SubClass
    tIndex = -1
    tHIndex = -1
    m_lImage = -1
    m_fIndex = -1
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    'Debug.Print Chr(KeyAscii)
    'Debug.Print GetEditText
End Sub

Private Sub UserControl_Terminate()
    cSubClass.Terminate
    Set cSubClass = Nothing
End Sub

Private Sub Bar_Scroll(): Bar_Change: End Sub
Private Sub Bar_Change()
    If PicList.Visible Then pDrawList
End Sub

'?Lista Desplegable
Private Sub PicList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '
End Sub

Private Sub PicList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lIndex As Integer
        lIndex = (Y - 1) \ m_ItemH
        If Y < 2 Or X < 2 Or X > PicList.ScaleWidth - 3 Then lIndex = -1
        If lIndex <> tHIndex Then
            tHIndex = lIndex
            pDrawList
        End If
        Timer1.Enabled = tHIndex <> -1
End Sub

Private Sub PicList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lIndex As Integer

    lIndex = (Y - 1) \ m_ItemH
    If pMouseOnHandle(PicList.hwnd) Then
        If Not (Bar + lIndex) > (ItemCount - 1) Then
            If tIndex <> Bar + lIndex Then
                tIndex = Bar + lIndex
                'pUpdateSelection
                RaiseEvent ListIndexChanged(tIndex)
            End If
        End If
    End If
    
    pUpdateSelection
    RaiseEvent ItemClick(tIndex, tItem(tIndex).Col(0).Text, tItem(tIndex).Col(1).Text, tItem(tIndex).Col(2).Text, tItem(tIndex).Col(3).Text)
    pShowList False
End Sub

Private Sub PicList_Resize()
On Error Resume Next
Dim TextH   As Integer

    With PicList
        Bar.Move .ScaleWidth - 17, 1, 16, .ScaleHeight - 2
        
            m_TextH = .TextHeight("ÁjqWJ")
            If c_hIml Then
                m_ItemH = c_ImgY + 4
                If TextH > c_ImgY Then m_ItemH = m_TextH + 4
            Else
               m_ItemH = m_TextH + 4
            End If
        
            If m_ItemH < 15 Then m_ItemH = 15
            'If .ScaleHeight < m_ItemH Then UserControl.Height = CtrlH * 15
        
        pUpdateList
    End With
End Sub

'?TextField
Private Sub Edit_Change()
    If Not m_bHint Then
        RaiseEvent Change
    'Else
        'Debug.Print "Uhint"
    End If
End Sub
Private Sub Edit_GotFocus()
    If m_bHint Then
        pUseHint False
    Else
        Edit.SelStart = 0: Edit.SelLength = Len(Edit)
    End If
End Sub
Private Sub Edit_Click()
   If PicList.Visible Then pShowList False
End Sub

'?UserControl
Private Sub UserControl_EnterFocus()
    m_HasFocus = True
    pDrawControl m_leState, True
End Sub
Private Sub UserControl_ExitFocus()
On Error Resume Next

    m_HasFocus = False
    If m_lstStyle = 1 Then GoTo zDraw
        
        If tIndex <> -1 Then
            If Edit <> tItem(tIndex).Text Then                  'Verificamos si el Texto es diferente del item
                If m_fIndex <> -1 And m_aComplete Then  'Si Busqueda Anterior es Diferente
                    ListIndex = m_fIndex
                ElseIf m_fIndex <> -1 Then
                    If tItem(m_fIndex).Text = Edit Then
                            ListIndex = m_fIndex
                    End If
                Else
                    Bar = 0: tIndex = -1
                    RaiseEvent ListIndexChanged(tIndex)
                    pDrawControl m_leState, True
                End If
            End If
        Else
                If m_fIndex <> -1 And m_aComplete Then
                    ListIndex = m_fIndex
                ElseIf m_fIndex <> -1 Then
                    If tItem(m_fIndex).Text = Edit Then
                            ListIndex = m_fIndex
                    End If
                End If
    End If
    If tIndex = -1 And Edit = "" And Trim(m_hintText) <> "" Then pUseHint True
zDraw:
    pDrawControl m_leState, True
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lT As Integer

        Select Case KeyCode
            Case 13                                       '{Enter}
                'RaiseEvent Click
                If m_lstStyle = 1 And PicList.Visible Then
                    If tHIndex <> -1 Then tIndex = Bar + tHIndex
                    pUpdateSelection
                    If PicList.Visible Then pShowList False
                End If
            Case 38                                       '{Up arrow}
                KeyCode = 0
                If tIndex > 0 Then ListIndex = ListIndex - 1
                RaiseEvent ItemClick(tIndex, tItem(tIndex).Col(0).Text, tItem(tIndex).Col(1).Text, tItem(tIndex).Col(2).Text, tItem(tIndex).Col(3).Text)

            Case 40                                       '{Down arrow}
                KeyCode = 0
                If tIndex < ItemCount - 1 Then ListIndex = ListIndex + 1
                RaiseEvent ItemClick(tIndex, tItem(tIndex).Col(0).Text, tItem(tIndex).Col(1).Text, tItem(tIndex).Col(2).Text, tItem(tIndex).Col(3).Text)

            Case 33                                       '{PageUp}
               ' If (m_ListIndex > m_VisibleRows) Then
                   ' ListIndex = ListIndex - (m_VisibleRows - 1)
                'Else                                      'NOT (M_LISTINDEX...
                    'ListIndex = 0
                'End If
            Case 34                                       '{PageDown}
               ' If (m_ListIndex < m_nItems - m_VisibleRows - 1) Then
                    'ListIndex = ListIndex + (m_VisibleRows - 1)
                'Else                                      'NOT (M_LISTINDEX...
                   ' ListIndex = m_nItems - 1
                'End If
            Case 36                                       '{Start}
                KeyCode = 0
                ListIndex = 0
            Case 35                                       '{End}
                'ListIndex = m_nItems - 1
                KeyCode = 0
                ListIndex = ItemCount - 1
            Case 32   '{Space} Select/Unselect
                    If m_lstStyle = JclsList Then
                        If tIndex = -1 Then ListIndex = 0
                    End If
            Case 27
                    If PicList.Visible Then pShowList False
            Case Else
                Dim NewIndex As Integer
                'Debug.Print KeyCode
                If Chr(KeyCode) = "" Then Exit Sub
                  If m_lstStyle = JclsList Then
                        NewIndex = pFindText(Chr(KeyCode), tIndex + 1, True)
                        If NewIndex <> -1 Then ListIndex = NewIndex
                 'Else
                        'Debug.Print Edit.Text
                  End If
        End Select
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
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PtInRect(m_RctDrop, X, Y) Then
            pDrawControl 2
            pShowList Not PicList.Visible
            Debug.Print "MouseDown_1"
    Else
        If PicList.Visible Then pShowList False
        Debug.Print "MouseDown_2"
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
If PtInRect(m_RctDrop, X, Y) Then
    If (Not m_bIsTracking) Then
             m_bIsTracking = True
             TrackMouseTracking UserControl.hwnd
    End If
    
    If Button = 1 Then pDrawControl 2 Else pDrawControl 1
         
    If m_HandCur And UserControl.MousePointer <> vbCustom Then UserControl.MousePointer = vbCustom
    Else: pDrawControl 0
    
    If UserControl.MousePointer <> vbNormal Then UserControl.MousePointer = vbNormal
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PtInRect(m_RctDrop, X, Y) Then
            If Button = 1 Then pDrawControl 1
            Debug.Print "MouseUp_1"
    Else
        If Button = 1 And Not pMouseOnHandle(UserControl.hwnd) Then
            pDrawControl 0
            Debug.Print "MouseUp_2"
        End If
    End If
End Sub

Private Sub UserControl_Resize()
Dim TextLeft As Integer
Dim TextTop As Integer
Dim TextW   As Integer
Dim TextH     As Integer
Dim CtrlH       As Integer

    If bcResizeFlag Then: Exit Sub
        bcResizeFlag = True
    With UserControl
        TextH = .TextHeight("ÁjqWJ")
    
        If c_hIml Then
            CtrlH = c_ImgY + 4
            If TextH > c_ImgY Then CtrlH = TextH + 4
        Else: CtrlH = TextH + 4: End If
        
        If CtrlH < 15 Then CtrlH = 15
        If .ScaleHeight < CtrlH Then UserControl.Height = CtrlH * 15
    
        Edit.Height = TextH + 2
        Select Case m_lstStyle
                Case 0:
                        If c_hIml Then
                            '--- AxioUK -Edit this Value----------------------------------------------------V
                            Edit.Move c_ImgX + 6, (.ScaleHeight - m_TextH) / 2, .ScaleWidth - (c_ImgX + 6 + 30), m_TextH
                        Else
                            '--- AxioUK -Edit this Value-------------------------------V
                            Edit.Move 4, (.ScaleHeight - m_TextH) / 2, .ScaleWidth - (35), m_TextH
                        End If
                        SetRect m_RctDrop, .ScaleWidth - 16, 0, ScaleWidth, .ScaleHeight
                        Edit.Visible = True
                Case 1
                        Edit.Visible = False
                        SetRect m_RctDrop, 0, 0, ScaleWidth, .ScaleHeight
            End Select
            
            TextLeft = IIf(c_hIml > 0, c_ImgX + 7, 5)
            TextTop = (.ScaleHeight - TextH) \ 2
            TextW = .ScaleWidth - (18 + IIf(c_hIml, c_ImgX + 6, 3)) '<----- AxioUK
            SetRect m_TR, TextLeft, TextTop, TextLeft + TextW, TextTop + TextH
            
            PicList.Width = UserControl.ScaleWidth
    End With
    pCreateRegions
    pDrawControl m_leState, True
    bcResizeFlag = False
End Sub

Public Sub AddItem(ByVal Col1Text As String, Optional ByVal Col2Text As String = "", _
                   Optional ByVal Col3Text As String = "", Optional ByVal Col4Text As String = "", _
                   Optional ByVal ItemImage As Integer = -1, Optional ItemKey As String = "", _
                   Optional ItemData As Long = 0, Optional ItemTag As String = "")
On Error Resume Next
Dim lJ      As Integer

    lJ = ItemCount
    ReDim Preserve tItem(lJ)
    
    With tItem(lJ)
        .Col(0).Text = Col1Text
        .Col(1).Text = Col2Text
        .Col(2).Text = Col3Text
        .Col(3).Text = Col4Text
        .Text = .Col(m_ColumnInList).Text
        .Image = ItemImage
        .Key = Trim(ItemKey)
        .Tag = ItemTag
        .Data = ItemData
        .ForeColor = -1
        
    End With
    pUpdateList
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
Dim vIndex  As Integer
Dim j     As Integer
Dim bc  As Boolean
    
    If ItemCount = 0 Or Index > ItemCount - 1 Or Index < 0 Then Exit Sub
    
    If ItemCount > 1 Then
            For j = Index To UBound(tItem) - 1
                LSet tItem(j) = tItem(j + 1)
            Next
            ReDim Preserve tItem(UBound(tItem) - 1)
    Else
        Erase tItem
    End If
    
    If Index = tIndex Then
        tIndex = -1
        bc = True
    ElseIf tIndex > Index Then
        tIndex = tIndex - 1
        bc = True
    End If
    pUpdateList
    If m_lstStyle <> JclsList Then pUpdateSelection False Else pUpdateSelection True
    If bc Then RaiseEvent ListIndexChanged(tIndex)
End Sub
Public Sub ShowDropDown(Optional ByVal Visible As Boolean)
    If PicList.Visible = Visible Then Visible = Not Visible
    pShowList Visible
End Sub

Public Sub CreateImageList(Optional Width As Integer = 16, Optional Height As Integer = 16, Optional hBitmap As Long, Optional MaskColor As Long = &HFFFFFFFF)
    
    c_hIml = ImageList_Create(Width, Height, &H20, 1, 1)
    If c_hIml And hBitmap Then
        If (MaskColor <> &HFFFFFFFF) Then
            ImageList_AddMasked c_hIml, hBitmap, MaskColor
        Else
            ImageList_Add c_hIml, hBitmap, 0
        End If
    End If
    Me.hImageList = c_hIml
End Sub


Private Sub pShowList(Optional ByVal Visible As Boolean)
Dim lW As Long
Dim Rct As RECT
Dim PT As POINTAPI
Dim lstTop  As Integer


    If Visible Then
       ' lH = UserControl.ScaleHeight + 1
        GetWindowRect UserControl.hwnd, Rct
        pDrawList
        SetParent PicList.hwnd, 0
        If Rct.Bottom + PicList.ScaleHeight > Screen.Height / Screen.TwipsPerPixelY Then lstTop = Rct.Top - (PicList.ScaleHeight + 1) Else lstTop = Rct.Bottom + 1
        SetWindowPos PicList.hwnd, 0, Rct.Left, lstTop, PicList.ScaleWidth, PicList.ScaleHeight, SWP_SHOWWINDOW
        'SetCapture PicList.hWnd
    Else
        SetParent PicList.hwnd, UserControl.hwnd
        'If (GetCapture = PicList.hWnd) Then ReleaseCapture
        PicList.Visible = False
        If Timer1.Enabled Then Timer1.Enabled = False
        tHIndex = -1
    End If
End Sub
Private Sub pUpdateSelection(Optional UpdateEditText As Boolean = True)
On Error Resume Next
    If tIndex = -1 Then m_Text = "" Else m_Text = tItem(tIndex).Text
    If UpdateEditText Then Edit.Text = m_Text
    pDrawControl m_leState, True
    If m_lstStyle = 0 Or m_lstStyle = 2 Then Edit.SelStart = 0: Edit.SelLength = Len(Edit)
End Sub

'-AxioUK-Update---------------------------
Private Sub pUpdateList()
Dim j       As Integer
Dim tLeft   As Integer
Dim tTop    As Integer
Dim tWidth  As Integer
Dim tHeight As Integer
Dim iWidth  As Integer
Dim iTop    As Integer

    If Not Ambient.UserMode Then Exit Sub
    With PicList
        m_VisibleItems = 5
        If ItemCount < m_VisibleItems Then m_VisibleItems = ItemCount
        If m_VisibleItems = 0 Then m_VisibleItems = 1
        .Height = (m_VisibleItems * m_ItemH) + 4
        
        j = ItemCount - m_VisibleItems
        Bar.Max = IIf(j > 0, j, 0)
        If Bar.Max > 0 Then
            Bar.Visible = True
            Bar.LargeChange = m_VisibleItems
        Else: Bar.Visible = False: End If
        
        tLeft = IIf(c_hIml, c_ImgX + 6, 6)
        iWidth = .ScaleWidth - (IIf(Bar.Max, 18, 2))
        tWidth = iWidth - 2
        
        If Not ItemCount = 0 Then
            ReDim tItemR(m_VisibleItems - 1): ReDim tTextR(m_VisibleItems - 1): ReDim tImgR(m_VisibleItems - 1)
        Else
            ReDim tItemR(0): ReDim tTextR(0): ReDim tImgR(0)
        End If
        
        For j = 0 To m_VisibleItems - 1
            iTop = (j * m_ItemH) + 2: tTop = iTop + (m_ItemH - m_TextH) \ 2
            SetRect tItemR(j), 2, iTop, iWidth, iTop + m_ItemH
            SetRect tTextR(j), tLeft, tTop, tWidth, tTop + m_TextH
            SetRect tImgR(j), 3, iTop + (m_ItemH - c_ImgY) / 2, 0, 0
        Next
        
    End With
    If PicList.Visible Then pDrawList
    
End Sub

Private Sub ShowColumnInList(ByVal iColumn As iCols)
Dim i As Long

    If Not ItemCount = 0 Then
        For i = 0 To UBound(tItem)
            tItem(i).Text = tItem(i).Col(iColumn).Text
        Next i
    End If
    
    Edit.Text = ""
    m_lImage = -1
    pUpdateList
    If PicList.Visible Then pDrawList
    
    pUpdateSelection
    
End Sub

Public Sub Clear()
    
    Erase tItem
    Edit.Text = ""
    m_lImage = -1
    pDrawControl 0, True
    isCleared = True
    pUpdateList
    If PicList.Visible Then pDrawList
    
End Sub

Private Sub pDrawControl(ByVal eState As Integer, Optional Force As Boolean)
On Error Resume Next
Dim j As Integer
Dim lPx As Integer
Dim lL As Integer
Dim lT  As Integer
Dim hColor As Long
Dim TR      As RECT
Dim sText   As String

    If eState = m_leState And Not Force Then Exit Sub
    If m_stdSkin Is Nothing Then Set m_stdSkin = UserControl.Picture
    If Not c_hSkin Then pSelectHSkin m_stdSkin.Handle
    
    If eState = 0 And m_HasFocus Then eState = 3
    If Not UserControl.Enabled Then eState = 4
    
    With UserControl
       .Cls
       lPx = eState * 15
        .BackColor = mc_Back
        If Edit.BackColor <> .BackColor Then Edit.BackColor = .BackColor
        Select Case m_lstStyle
            Case 0 '?Combo
            
                    lL = (.ScaleWidth - 28)
                                        
                    .ForeColor = GetPixel(c_hSkin, lPx, 0)
                    RoundRect .hdc, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0
                    If eState = 3 Then
                        .ForeColor = pBlendColor(GetPixel(c_hSkin, lPx + 1, 1), .BackColor, 50)
                        RoundRect .hdc, 1, 1, .ScaleWidth - 1, .ScaleHeight - 1, 0, 0
                    End If
                    If m_sdBack Then
                        pRenderStretch .hdc, lL, 0, 28, .ScaleHeight, c_hSkin, lPx, 0, 15, 23, 3 '15, 23, 3
                    ElseIf eState = 1 Or eState = 2 Then
                        pRenderStretch .hdc, lL, 0, 28, .ScaleHeight, c_hSkin, lPx, 0, 15, 23, 3 '15, 23, 3
                    End If
                    
                    If c_hIml Then
                        If tIndex <> -1 Then
                            If m_lImage <> tItem(tIndex).Image Then m_lImage = tItem(tIndex).Image
                        End If
                        ImageList_Draw c_hIml, m_lImage, .hdc, 3, ((.ScaleHeight - c_ImgY) \ 2), 0
                    End If
                    
                    If Enabled Then
                        .ForeColor = mc_Fore
                    ElseIf m_bHint Then
                        .ForeColor = &H808080
                    Else
                        .ForeColor = pBlendColor(vbBlack, GetPixel(c_hSkin, lPx, 0), 50)
                    End If
                    
            Case 1 '?List
                    CopyRect TR, m_TR
                    If tIndex <> -1 Then
                        .ForeColor = mc_Fore
                        sText = tItem(tIndex).Text
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
                     If c_hIml And tItem(tIndex).Image <> -1 Then ImageList_Draw c_hIml, tItem(tIndex).Image, .hdc, 3, ((.ScaleHeight - c_ImgY) / 2) + IIf(eState = 2, 1, 0), 0
            Case 2
                'Nothing...
        End Select
        
        If m_lstStyle <> 2 Then '?Drop Arrow

                lL = (.ScaleWidth - 26) + ((25 - 4) / 2) 'lL = (.ScaleWidth - 16) + ((15 - 4) / 2)
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
                If m_lstStyle = 1 Then UserControl.Line (.ScaleWidth - 26, 7 + IIf(eState = 2, 1, 0))-(.ScaleWidth - 26, .ScaleHeight - 7 + IIf(eState = 2, 1, 0)), GetPixel(c_hSkin, lPx, 0), B
        End If
        
        Select Case m_Shape
          Case 0 'RoundedRectangle
                SetPixelV .hdc, 1, 1, GetPixel(c_hSkin, lPx, 0)
                SetPixelV .hdc, 1, .ScaleHeight - 2, GetPixel(c_hSkin, lPx, 22)
                SetPixelV .hdc, .ScaleWidth - 2, 1, GetPixel(c_hSkin, lPx + 24, 0)
                SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, GetPixel(c_hSkin, lPx + 24, 22)
            Case 1 'Rectangle
            Case 2 'CutLeft
                SetPixelV .hdc, .ScaleWidth - 2, 1, GetPixel(c_hSkin, lPx + 24, 0)
                SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, GetPixel(c_hSkin, lPx + 24, 22)
            Case 3 'CutRight
                SetPixelV .hdc, 1, 1, GetPixel(c_hSkin, lPx, 0)
                SetPixelV .hdc, 1, .ScaleHeight - 2, GetPixel(c_hSkin, lPx, 22)
            Case 4 'CutTop
                SetPixelV .hdc, 1, .ScaleHeight - 2, GetPixel(c_hSkin, lPx, 22)
                SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, GetPixel(c_hSkin, lPx + 24, 22)
            Case 5 'CutBottom
                SetPixelV .hdc, 1, 1, GetPixel(c_hSkin, lPx, 0)
                SetPixelV .hdc, .ScaleWidth - 2, 1, GetPixel(c_hSkin, lPx + 24, 0)
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

    End With
End Sub
'-AxioUK-Update---------------------------
Private Sub pDrawList(Optional ItemHot As Integer = -1)
Dim j As Integer

Dim TR   As RECT
    With PicList
        .Cls
        .BackColor = mc_lBack
        PicList.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), mc_Border, B
        If ItemCount = 0 Then Exit Sub
        For j = 0 To m_VisibleItems - 1
           If j + Bar > ItemCount - 1 Then Exit For
            
                If (Bar + j) = tIndex Then '?Selected
                    FillRect .hdc, tItemR(j), CreateSolidBrush(pBlendColor(mc_Selection, .BackColor, 50))
                    .ForeColor = pBlendColor(mc_Selection, .BackColor, 120)
                    RoundRect .hdc, tItemR(j).Left, tItemR(j).Top, tItemR(j).Right, tItemR(j).Bottom, 3, 3
               ElseIf j = tHIndex Then '?Higlig
                    FillRect .hdc, tItemR(j), CreateSolidBrush(pBlendColor(mc_Selection, .BackColor, 40))
                    .ForeColor = pBlendColor(mc_Selection, .BackColor, 80)
                    RoundRect .hdc, tItemR(j).Left, tItemR(j).Top, tItemR(j).Right, tItemR(j).Bottom, 3, 3
                End If
            
                .ForeColor = mc_lFore
                DrawText .hdc, tItem(j + Bar).Text, Len(tItem(j + Bar).Text), tTextR(j), DT_FLAG
                
            If c_hIml And tItem(Bar + j).Image > -1 Then
                ImageList_Draw c_hIml, tItem(Bar + j).Image, .hdc, tImgR(j).Left, tImgR(j).Top, 0
            End If
        Next
    End With
End Sub
Private Function pFindText(ByVal Text As String, Optional ByVal iStart As Integer = -1, Optional IgnoreCase As Boolean, Optional CompleteString As Boolean) As Integer
On Error Resume Next
Dim j      As Integer
Dim iText    As String
Dim iRet      As Integer
Dim tLn       As Integer
            
            If ItemCount = 0 Then pFindText = -1: Exit Function
            'Debug.Print "B: " & Text
            If iStart > ItemCount - 1 Then iStart = 0
            If IgnoreCase Then Text = UCase(Text)
            
            iRet = -1
            tLn = Len(Text)
            
            For j = iStart To ItemCount - 1
            
                iText = IIf(CompleteString, tItem(j).Text, Left(tItem(j).Text, tLn))
                If IgnoreCase Then iText = UCase(iText)
                
                If iText <> "" Then
                        If Text = iText Then: iRet = j: Exit For
                End If
            Next
            If iRet = -1 And iStart > 0 Then
                For j = 0 To iStart
                     iText = IIf(CompleteString, tItem(j).Text, Left(tItem(j).Text, tLn))
                    If IgnoreCase Then iText = UCase(iText)
                
                    If iText <> "" Then
                            If Text = iText Then: iRet = j: Exit For
                    End If
                Next
            End If
            pFindText = iRet
End Function
Private Sub Timer1_Timer()
  If Not pMouseOnHandle(PicList.hwnd) Then
        Timer1.Enabled = False
        tHIndex = -1
       pDrawList
    End If
    DoEvents
End Sub

'?Autocomplete
Private Sub pAutocomplete(Optional ByVal DeleteKey As Boolean)
Dim NewIndex As Integer
Dim sText       As String
Dim iStart      As Integer
Dim sList          As Boolean

            If m_lstStyle = 1 Then Exit Sub
            
            If DeleteKey Then
                If tIndex <> -1 Then
                    tIndex = -1
                    RaiseEvent ListIndexChanged(tIndex)
                End If
                tHIndex = -1
                If PicList.Visible Then pDrawList
                Exit Sub
            End If
            
            NewIndex = pFindText(Edit.Text, , True)
            If NewIndex <> tIndex Then
                If NewIndex < Bar And NewIndex > -1 Then
                            Bar = NewIndex
                ElseIf (NewIndex > Bar + m_VisibleItems - 1) Then
                            Bar = NewIndex - m_VisibleItems + 1
                End If
            End If
            
           ' If m_aComplete Then m_fIndex = NewIndex
            sList = NewIndex <> -1
            m_fIndex = NewIndex
            
            If Not m_aComplete Then Exit Sub
            
             '?AutoComplete
             'm_fIndex = NewIndex
                If NewIndex <> -1 And Not DeleteKey Then
                    sText = tItem(NewIndex).Text
                    If Len(sText) = Len(Edit) Then Exit Sub
                    sText = Right(sText, Len(sText) - Len(Edit))
                    
                    iStart = Len(Edit)
                    Edit.Text = Edit.Text & sText
                    Edit.SelStart = iStart: Edit.SelLength = Len(Edit)
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
        Edit.ForeColor = mc_Fore
        m_bHint = False
    End If
End Sub

Property Get ItemCount() As Integer
On Error GoTo Err
    If isCleared Then
        ItemCount = 0
    Else
        ItemCount = UBound(tItem) + 1
    End If
    
    isCleared = False
    Exit Property
Err:
isCleared = False
ItemCount = 0
End Property

Property Get hImageList() As Long: hImageList = c_hIml: End Property
Property Let hImageList(ByVal hwnd As Long)
    c_hIml = hwnd
    If c_hIml Then ImageList_GetIconSize c_hIml, c_ImgX, c_ImgY Else c_ImgX = 0: c_ImgY = 0
    UserControl_Resize
    PicList_Resize
End Property
Property Get SkinPicture() As StdPicture: Set SkinPicture = m_stdSkin: End Property
Property Set SkinPicture(NewSkin As StdPicture)
    Set m_stdSkin = NewSkin
    If m_stdSkin Is Nothing Then Set m_stdSkin = UserControl.Picture
    PropertyChanged "stdSkin"
    pSelectHSkin m_stdSkin.Handle
    pDrawControl m_leState, True
End Property

Property Get ListStyle() As JCListStyle
    ListStyle = m_lstStyle
End Property
Property Let ListStyle(NewStyle As JCListStyle)
    m_lstStyle = NewStyle
    PropertyChanged "lstStyle"
    UserControl_Resize
    pUpdateSelection
    Edit.Visible = m_lstStyle = 0 Or m_lstStyle = 2
End Property

Property Get ListIndex() As Integer
  ListIndex = tIndex
End Property

Property Let ListIndex(NewIndex As Integer)
Dim lT As Integer
    If NewIndex > ItemCount - 1 Then Exit Property
    tIndex = NewIndex
    pUpdateSelection
    
    If tIndex < Bar And tIndex > -1 Then
        Bar = tIndex
    ElseIf (tIndex > Bar + m_VisibleItems - 1) Then
        Bar = tIndex - m_VisibleItems + 1
    End If
    RaiseEvent ListIndexChanged(tIndex)
    If PicList.Visible Then pDrawList
End Property

Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Property Set Font(NewFont As StdFont)
    Set UserControl.Font = NewFont
    Set Edit.Font = UserControl.Font
    UserControl_Resize
    PropertyChanged "Font"
End Property
Property Get FontList() As StdFont: Set FontList = PicList.Font: End Property
Property Set FontList(NewFont As StdFont)
    Set PicList.Font = NewFont
    PropertyChanged "lFont"
    pUpdateList
End Property
Property Get ForeColor() As OLE_COLOR: ForeColor = mc_Fore: End Property
Property Let ForeColor(NewColor As OLE_COLOR)
    mc_Fore = NewColor
    PropertyChanged "Fore"
    pDrawControl m_leState, True
    If PicList.Visible Then pDrawList
End Property
Property Get ForeColorList() As OLE_COLOR: ForeColorList = mc_lFore: End Property
Property Let ForeColorList(NewColor As OLE_COLOR)
    mc_lFore = NewColor
    PropertyChanged "lFore"
    If PicList.Visible Then pDrawList
End Property
Property Get Enabled() As Boolean: Enabled = UserControl.Enabled: End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call PropertyChanged("Enabled")
    pDrawControl 0, True
    If New_Enabled = False And PicList.Visible Then pShowList False
End Property
Property Get BackColor() As OLE_COLOR: BackColor = mc_Back: End Property
Property Let BackColor(NewColor As OLE_COLOR)
        mc_Back = NewColor
        PropertyChanged "cBack"
        pDrawControl m_leState, True
End Property
Property Get BackColorList() As OLE_COLOR: BackColorList = mc_lBack: End Property
Property Let BackColorList(NewColor As OLE_COLOR)
        mc_lBack = NewColor
        PropertyChanged "cLBack"
        If PicList.Visible Then pDrawList
End Property
Property Get BorderColorList() As OLE_COLOR: BorderColorList = mc_Border: End Property
Property Let BorderColorList(ByVal New_Color As OLE_COLOR)
    mc_Border = New_Color
    Call PropertyChanged("cBorder")
    If PicList.Visible Then pDrawList
End Property
Property Get SelectionColor() As OLE_COLOR: SelectionColor = mc_Selection: End Property
Property Let SelectionColor(ByVal New_Color As OLE_COLOR)
    mc_Selection = New_Color
    Call PropertyChanged("cSelection")
    If PicList.Visible Then pDrawList
End Property
Public Property Get ShapeStyle() As JComboShape: ShapeStyle = m_Shape: End Property
Public Property Let ShapeStyle(Value As JComboShape)
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
Property Get FocusRect() As Boolean: FocusRect = m_sFocus: End Property
Property Let FocusRect(newValue As Boolean)
    m_sFocus = newValue
    PropertyChanged "sFocus"
    If m_HasFocus Then pDrawControl m_leState, True
End Property

Property Get Text(Optional iCol As Integer) As String
    If m_lstStyle = 1 Then
        Text = m_Text
    Else
        'If Not m_bHint Then Text = Edit.Text
        If Not m_bHint Then
          If IsMissing(iCol) Then
              Text = Edit.Text
          Else
              Text = tItem(tIndex).Col(iCol).Text
          End If
        End If
    End If
    'tItem(tIndex).Col(iCol).Text
End Property

Property Let Text(Optional iCol As Integer, NewProp As String)
Dim j As Integer

    If m_lstStyle = JclsList Then
        j = pFindText(NewProp)
        If j <> -1 Then ListIndex = j
    Else
            If m_bHint Then pUseHint False
            
          If IsMissing(iCol) Then
              Edit.Text = NewProp
          Else
              tItem(tIndex).Col(iCol).Text = NewProp
          End If
    End If
    '
    m_Text = NewProp
    PropertyChanged "Text"
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
Property Get AutoComplete() As Boolean: AutoComplete = m_aComplete: End Property
Property Let AutoComplete(NewProp As Boolean)
    m_aComplete = NewProp
    PropertyChanged "aComplete"
End Property
Property Get HintText() As String: HintText = m_hintText: End Property
Property Let HintText(NewProp As String)
    m_hintText = NewProp
    PropertyChanged "Hint"
    If (m_bHint Or Trim(m_hintText) <> "" Or Edit = "") And m_lstStyle <> 1 Then pUseHint True
    If m_lstStyle = 1 And ListIndex = -1 Then pDrawControl m_leState, True
End Property
'-AxioUK----------------------------------------------
Public Property Get ColumnInList() As iCols
    ColumnInList = m_ColumnInList
End Property

Public Property Let ColumnInList(ByVal NewColumnInList As iCols)
    m_ColumnInList = NewColumnInList
    PropertyChanged "ColumnInList"
    Call ShowColumnInList(m_ColumnInList)
End Property
'-----------------------------------------------------
Property Get ItemData(Index As Integer) As Long
On Error GoTo Err
        ItemData = tItem(Index).Data
    Exit Property
Err:
ItemData = 0
End Property
Property Let ItemData(Index As Integer, Value As Long)
On Error GoTo Err
         tItem(Index).Data = Value
Err:
End Property
Property Get ItemText(ByVal Index As Integer, ByVal iCol As Integer) As String
On Error Resume Next
    If Index > ItemCount - 1 Or Index < 0 Or ItemCount = 0 Then Exit Property
    'ItemText = tItem(Index).Text
    ItemText = tItem(Index).Col(iCol).Text
End Property
Property Let ItemText(ByVal Index As Integer, ByVal iCol As Integer, ByVal NewProp As String)
    On Error Resume Next
    If Index > ItemCount - 1 Or Index < 0 Or ItemCount = 0 Then Exit Property
    'tItem(Index).Text = NewProp
    tItem(Index).Col(iCol).Text = NewProp
    If tIndex = Index Then pUpdateSelection
    If PicList.Visible Then pDrawList
End Property
Property Let ItemTag(ByVal Index As Integer, Value As String)
    If Index > ItemCount - 1 Or Index < 0 Or ItemCount = 0 Then Exit Property
    tItem(Index).Tag = Value
End Property
Property Get ItemTag(ByVal Index As Integer) As String
    If Index > ItemCount - 1 Or Index < 0 Or ItemCount = 0 Then Exit Property
    ItemTag = tItem(Index).Tag
End Property
Property Get ItemImage(ByVal Index As Integer) As Integer
On Error Resume Next
    If Index > ItemCount - 1 Or Index < 0 Or ItemCount = 0 Then GoTo Err
    ItemImage = tItem(Index).Image
    Exit Property
Err:
    ItemImage = -1
End Property
Property Let ItemImage(ByVal Index As Integer, ByVal NewProp As Integer)
On Error Resume Next
    If Index > ItemCount - 1 Or Index < 0 Or ItemCount = 0 Then Exit Property
     tItem(Index).Image = NewProp
     If tIndex = Index Then pUpdateSelection
     If PicList.Visible Then pDrawList
End Property
Property Get ItemKey(ByVal Index As Integer) As String
On Error Resume Next
    If Index > ItemCount - 1 Or Index < 0 Or ItemCount = 0 Then Exit Property
    ItemKey = tItem(Index).Key
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set m_stdSkin = .ReadProperty("stdSkin", UserControl.Picture)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        Set PicList.Font = .ReadProperty("lFont", PicList.Font)
        
        m_lstStyle = .ReadProperty("lstStyle", 0)
        mc_Fore = .ReadProperty("Fore", 0)
        mc_lFore = .ReadProperty("lFore", 0)
        mc_Back = .ReadProperty("cBack", vbWhite)
        mc_lBack = .ReadProperty("cLBack", vbWhite)
        mc_Border = .ReadProperty("cBorder", &HAEAEAE)
        mc_Selection = .ReadProperty("cSelection", &HFF6600)
        m_Shape = .ReadProperty("Shape", 0)
        m_sdBack = .ReadProperty("sdBack", True)
        m_sFocus = .ReadProperty("sFocus", True)
        m_Text = .ReadProperty("Text", "")
        m_HandCur = .ReadProperty("HandCursor", False)
        m_aComplete = .ReadProperty("aComplete", False)
        m_hintText = .ReadProperty("Hint", "")
        m_ColumnInList = .ReadProperty("ColumnInList", 0)   'AxioUK

        UserControl.Enabled() = .ReadProperty("Enabled", True)
    End With
    
        '?Subclass
  With cSubClass
        If Ambient.UserMode Then
                m_bTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
                m_bSuppMouseTrack = m_bTrackHandler32
                If Not m_bSuppMouseTrack Then m_bSuppMouseTrack = IsFunctionSupported("_TrackMouseEvent", "Comctl32")
                
                m_PhWnd = UserControl.Parent.hwnd
                
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
                     '?AutoComplete
                     .AddMsg Edit.hwnd, WM_CHAR, MSG_BEFORE_AFTER
                     .AddMsg Edit.hwnd, WM_CLEAR, MSG_AFTER
                     .AddMsg Edit.hwnd, WM_CUT, MSG_AFTER
                     .AddMsg Edit.hwnd, WM_PASTE, MSG_AFTER
                     .AddMsg Edit.hwnd, WM_UNDO, MSG_AFTER
                 End If
                
                SetWindowLongA PicList.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
        End If
  End With

    Set Edit.Font = UserControl.Font
    If Trim(m_Text) <> "" And m_lstStyle = 1 Then
        Edit = m_Text
    ElseIf Trim(m_hintText) <> "" Then
       pUseHint True
    End If
      
      If m_HandCur Then
        UserControl.MouseIcon = pGetSystemHandCursor
        UserControl.MousePointer = vbCustom
      End If
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "stdSkin", m_stdSkin, UserControl.Picture
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "lFont", PicList.Font
        
        .WriteProperty "lstStyle", m_lstStyle
        .WriteProperty "Fore", mc_Fore
        .WriteProperty "lFore", mc_lFore
        .WriteProperty "cBack", mc_Back
        .WriteProperty "cLBack", mc_lBack
        .WriteProperty "cBorder", mc_Border
        .WriteProperty "cSelection", mc_Selection
        .WriteProperty "Shape", m_Shape
        .WriteProperty "sdBack", m_sdBack
        .WriteProperty "sFocus", m_sFocus
        .WriteProperty "Text", m_Text
        .WriteProperty "HandCursor", m_HandCur
        .WriteProperty "aComplete", m_aComplete
        .WriteProperty "Hint", m_hintText
        .WriteProperty "ColumnInList", m_ColumnInList, 0    'AxioUK
        
        .WriteProperty "Enabled", UserControl.Enabled, True
    End With
End Sub
Private Sub pSelectHSkin(Optional lHandle As Long = 0)
Dim j As Integer
    If c_hSkin Then Call DeleteDC(c_hSkin)
    c_hSkin = CreateCompatibleDC(0)
    Call SelectObject(c_hSkin, lHandle)
End Sub
Private Function pMouseOnHandle(hwnd As Long) As Boolean
    Dim PT As POINTAPI
    GetCursorPos PT
    pMouseOnHandle = (WindowFromPoint(PT.X, PT.Y) = hwnd)
End Function

Private Sub pCreateRegions(Optional EllipseW As Long = 5, Optional EllipseH As Long = 5)

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
Private Function GetHitTest() As Integer
Dim PT          As POINTAPI
Dim lW          As Integer
    
    GetCursorPos PT
    If WindowFromPoint(PT.X, PT.Y) = PicList.hwnd Then
        ScreenToClient PicList.hwnd, PT
        lW = PicList.ScaleWidth - 3
        
        If Not PT.X > lW And Not PT.X < 2 Then
            GetHitTest = (PT.Y - 1) \ m_ItemH
        Else
            GetHitTest = -1
        End If
        If GetHitTest > ItemCount - 1 Then GetHitTest = -1
    Else
        GetHitTest = -1
    End If
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

' Ordinal #2
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
                        If Bar.Visible Then
                            If wParam < 0 Then
                                If Bar < Bar.Max Then Bar = Bar + 1
                            Else
                                If Bar > 0 Then Bar = Bar - 1
                            End If
                        End If
                    Else
                        If wParam < 0 Then
                            If tIndex < ItemCount - 1 And ItemCount > 0 Then
                                ListIndex = ListIndex + 1
                            End If
                        Else
                            If tIndex > 0 And ItemCount > 0 Then
                                ListIndex = ListIndex - 1
                            End If
                        End If
                    End If
                    
        Case 516, 513, 164, 269
             If PicList.Visible Then pShowList False
        Case WM_CHAR
                If bBefore Then Exit Sub
                If ItemCount = 0 Then Exit Sub
                If hwnd = Edit.hwnd Then
                    ' If wParam > 0 And wParam < 27 And wParam <> 8 And wParam <> 13 Then Exit Sub
                    
                    If wParam = 8 Or wParam = 24 Then 'Borrar / Cortar Texto
                        pAutocomplete True
                        Exit Sub
                    End If
                    If wParam = 13 Then                         '/Enter
                              If m_fIndex <> -1 Then             'Si se Encontro en Busqueda Anterior
                                    ListIndex = m_fIndex
                               Else
                                        If tIndex <> -1 Then                                'Si el control tiene texto diferente
                                            If Edit <> tItem(tIndex).Text Then      'que el item Seleccionado
                                                Bar = 0: tIndex = -1
                                                pDrawControl m_leState, True
                                                RaiseEvent ListIndexChanged(tIndex)
                                            End If
                                        Else 'Buscar nuevo Texto ingresado en los Items
                                            If m_fIndex = -1 Then
                                                Dim nIndex As Integer
                                                nIndex = pFindText(Edit, , True, True)
                                                If nIndex <> -1 Then ListIndex = nIndex
                                            End If
                                        End If
                            End If
                            If PicList.Visible Then pShowList False
                            Exit Sub
                    End If
                     If wParam > 0 And wParam < 27 Then Exit Sub
                    pAutocomplete
                   
                End If
        Case Else
    End Select
End Sub

