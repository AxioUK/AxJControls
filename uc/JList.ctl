VERSION 5.00
Begin VB.UserControl axJList 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "JList.ctx":0000
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ToolboxBitmap   =   "JList.ctx":1042
End
Attribute VB_Name = "axJList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type
Private Type RECT
    L   As Long
    T   As Long
    r   As Long
    B   As Long
End Type

Private Const HWND_TOPMOST    As Long = -1
Private Const HWND_NOTOPMOST  As Long = -2
Private Const SWP_NOSIZE      As Long = &H1
Private Const SWP_NOMOVE      As Long = &H2
Private Const SWP_NOACTIVATE  As Long = &H10
Private Const SWP_SHOWWINDOW  As Long = &H40
Private Const SWP_HIDEWINDOW  As Long = &H80

Private Const C_NULL_RESULT   As Long = -1

'/Window
'Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
'Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
'Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
'Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function GetFocus Lib "user32" () As Long
'Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetCapture Lib "user32.dll" () As Long
'Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
'Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long

'/WindowMessages
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'/Theme
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As Any) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
'Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

'?Border
Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'/Draw
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function OleTranslateColor2 Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
'Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
'Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
'Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Private Type tHeader
    Text    As String
    Image   As Long
    Width   As Long
    Aling   As Integer
    IAlign  As Integer
End Type
Private Type tSubItem
    Text    As String
    Icon    As Long
End Type
Private Type tItem
    Item()  As tSubItem
    Data    As Long
    Tag     As String
End Type

Event ItemClick(Item As Long)

Private cSubClass As c_SubClass
Private WithEvents cScroll As c_ScrollBars
Attribute cScroll.VB_VarHelpID = -1

Private m_ItemH         As Long
Private m_HeaderH       As Long
Private m_GridLineColor As Long
Private m_GridStyle     As Integer
Private m_Striped       As Boolean
Private m_Header        As Boolean
Private m_DrawEmpty     As Boolean
Private m_StripedColor  As Long
Private m_ForeColor     As OLE_COLOR
Private m_SelColor      As OLE_COLOR
Private m_ForeSel       As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_VisibleRows   As Long
Private m_DropW         As Long
Private m_ShowHeader    As Boolean
Private m_ColRet        As Long

Private pmTrack(3)      As Long
Private m_PhWnd         As Long
Private m_hWnd          As Long
Private m_cols()        As tHeader
Private m_items()       As tItem
Private m_Iml           As Long

Private m_SelRow        As Long
Private m_bTrack        As Boolean
Private m_img           As POINTAPI
Private m_GridW         As Long
Private m_RowH          As Long
Private t_Row           As Long

Private e_Scale         As Long
Private lnScale         As Long
Private m_Visible       As Boolean



Private Sub UserControl_InitProperties()
    m_HeaderH = 22
    m_ItemH = 17
    m_GridLineColor = &HF0F0F0
    m_GridStyle = 3
    m_Striped = True
    m_StripedColor = &HFDFDFD
    m_SelColor = vbHighlight  '&HDDAC84
    m_BorderColor = &H908782  '&HB2ACA5
    m_Header = True
    m_VisibleRows = 8
End Sub

Private Sub UserControl_Initialize()
    Set cSubClass = New c_SubClass
    Set cScroll = New c_ScrollBars
    t_Row = -1: m_SelRow = -1
    e_Scale = GetWindowsDPI
    Select Case e_Scale
        Case 1, 2: lnScale = 1
        Case 3, 4: lnScale = 2
        Case 5: lnScale = 4
    End Select
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub UserControl_Click()
On Error Resume Next
    If t_Row <> -1 Then
        If Not IsCompleteVisibleRow(t_Row) Then SetVisibleItem t_Row
            SendText m_items(t_Row).Item(m_ColRet).Text
            RaiseEvent ItemClick(t_Row)
            HideList

'        If IsCompleteVisibleRow(t_Row) Then
'            SendText m_items(t_Row).Item(m_ColRet).Text
'            RaiseEvent ItemClick(t_Row)
'            HideList
'        Else
'            SetVisibleItem t_Row
'        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If t_Row <> -1 Then
        'If Not IsCompleteVisibleRow(t_Row) Then SetVisibleItem t_Row
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lRow    As Long

    If Not m_bTrack Then
        TrackMouseEvent pmTrack(0)
        'RaiseEvent MouseEnter
        m_bTrack = True
    End If

    lRow = GetRowFromY(Y)
    If X > m_GridW Then lRow = -1
    If lRow <> t_Row Then
        t_Row = lRow
        DrawGrid
    End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

    If Ambient.UserMode Then
        'UserControl.BorderStyle = 1
    Else
        UserControl.BorderStyle = 0
        UserControl.Width = 32 * Screen.TwipsPerPixelX
        UserControl.Height = 32 * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub UserControl_Terminate()
    Erase m_items
    Erase m_cols
    cSubClass.UnSubclass UserControl.hwnd
    cSubClass.UnSubclass m_hWnd
    
    Set cSubClass = Nothing
    Set cScroll = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "HeaderH", m_HeaderH
        .WriteProperty "LineColor", m_GridLineColor
        .WriteProperty "GridStyle", m_GridStyle
        .WriteProperty "Striped", m_Striped
        .WriteProperty "StripedColor", m_StripedColor
        .WriteProperty "SelColor", m_SelColor
        .WriteProperty "ItemH", m_ItemH, 17
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "Header", m_Header
        .WriteProperty "ForeColor", m_ForeColor
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "VisibleRows", m_VisibleRows
        .WriteProperty "DropWidth", m_DropW
        .WriteProperty "ColumnInBox", m_ColRet, 0
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim H As Double
    With PropBag
        m_HeaderH = .ReadProperty("HeaderH", 24)
        m_GridLineColor = .ReadProperty("LineColor", &HF0F0F0)
        m_GridStyle = .ReadProperty("GridStyle", 3)
        m_Striped = .ReadProperty("Striped", True)
        m_StripedColor = .ReadProperty("StripedColor", &HFDFDFD)
        m_SelColor = .ReadProperty("SelColor", vbHighlight)
        m_ItemH = .ReadProperty("ItemH", 17)
        m_BorderColor = .ReadProperty("BorderColor", &HB2ACA5)
        m_Header = .ReadProperty("Header", True)
        m_ForeColor = .ReadProperty("ForeColor", 0)
        m_VisibleRows = .ReadProperty("VisibleRows", 8)
        m_DropW = .ReadProperty("DropWidth", 0)
        Set UserControl.Font() = .ReadProperty("Font", UserControl.Font)
        UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
        m_ColRet = .ReadProperty("ColumnInBox", 0)
    End With


    If Ambient.UserMode Then
            
        With cScroll
            .Create UserControl.hwnd
            .SmallChange(0) = 20 '48
            .SmallChange(1) = 16
        End With
        
        With cSubClass
        
            If .Subclass(UserControl.hwnd, , , Me) Then
                .AddMsg UserControl.hwnd, WM_WINDOWPOSCHANGING, MSG_AFTER
                .AddMsg UserControl.hwnd, WM_WINDOWPOSCHANGED, MSG_AFTER
                .AddMsg UserControl.hwnd, WM_GETMINMAXINFO, MSG_AFTER
                .AddMsg UserControl.hwnd, WM_LBUTTONDOWN, MSG_AFTER
                .AddMsg UserControl.hwnd, WM_SIZE, MSG_AFTER
                
            End If
          
'            If .Subclass(m_PhWnd, , , Me) Then
'                .AddMsg UserControl.Parent.hwnd, WM_KILLFOCUS, MSG_AFTER
'                .AddMsg UserControl.Parent.hwnd, WM_SETFOCUS, MSG_AFTER
'                .AddMsg UserControl.Parent.hwnd, WM_MOUSEWHEEL, MSG_AFTER
'                .AddMsg UserControl.Parent.hwnd, WM_MOUSELEAVE, MSG_AFTER
'            End If
            
        End With
        
        pmTrack(0) = 16&
        pmTrack(1) = &H2
        pmTrack(2) = UserControl.hwnd
        
        SetWindowLongA UserControl.hwnd, -20, GetWindowLong(UserControl.hwnd, -20) Or &H80&
        SetClassLong UserControl.hwnd, -26, GetClassLong(UserControl.hwnd, -26) Or &H20000  'CS_DROPSHADOW
        
        m_HeaderH = m_HeaderH * e_Scale
        H = UserControl.TextHeight("Ájq")
        If m_HeaderH < (H + (6 * e_Scale)) Then HeaderHeight = H + (6 * e_Scale)
        
        Extender.Visible = False
        UserControl.BorderStyle = 1
        UserControl.Picture = Nothing
        UpdateValues
    End If
    
End Sub

'/Scroll
Private Sub cScroll_Change(eBar As EFSScrollBarConstants)
    DrawGrid
End Sub

Private Sub cScroll_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
    DrawGrid
End Sub

'/Public
Public Sub Init(hwnd As Long)
Dim uRct As RECT

    If m_hWnd Then cSubClass.UnSubclass m_hWnd
    If m_DropW = 0 Then
        GetWindowRect hwnd, uRct
        m_DropW = uRct.r - uRct.L
    End If
    
    m_hWnd = hwnd
    If cSubClass.Subclass(m_hWnd, , , Me) Then
        Call cSubClass.AddMsg(m_hWnd, WM_KILLFOCUS, MSG_BEFORE)
    End If
    
End Sub
Public Sub AddColumn(ByVal Text As String, Optional ByVal Width As Long = 100, Optional ByVal Alignment As AlignmentConstants)
Dim L       As Long
Dim i       As Long

    Width = Width * e_Scale
    L = ColumnCount
    
    ReDim Preserve m_cols(L)
    With m_cols(L)
        .Text = Text
        .Width = Width
        .Aling = Alignment
    End With
    m_GridW = m_GridW + Width
End Sub

Public Function AddItem(ByVal Text As String, Optional ByVal IconIndex As Long = -1, Optional ByVal ItemData As Long, Optional ByVal ItemTag As String = "") As Long
On Local Error Resume Next
Dim L   As Long
Dim i   As Long

    L = ItemCount
    ReDim Preserve m_items(L)
    
    With m_items(L)
        ReDim .Item(ColumnCount - 1)
        .Item(0).Text = Text
        .Item(0).Icon = IconIndex
        .Data = ItemData
        .Tag = ItemTag
        
        For i = 1 To ColumnCount - 1: .Item(i).Icon = -1: Next

    End With
    AddItem = L
    UpdateScrollV
End Function

Public Sub RemoveItem(ByVal Index As Long)
On Local Error Resume Next
Dim j As Integer

    If ItemCount = 0 Or Index > ItemCount - 1 Or ItemCount < 0 Or Index < 0 Then Exit Sub
    
    If ItemCount > 1 Then
         For j = Index To UBound(m_items) - 1
            LSet m_items(j) = m_items(j + 1)
         Next
        ReDim Preserve m_items(UBound(m_items) - 1)
    Else
        Erase m_items
    End If
    
    UpdateScrollV
    If m_SelRow <> -1 Then
        If m_SelRow = Index Then m_SelRow = -1
        If m_SelRow > Index Then m_SelRow = m_SelRow - 1
    End If
    DrawGrid
End Sub

Public Sub ClearItems()
    Erase m_items
    m_SelRow = -1
    UpdateScrollV
End Sub

Public Sub ColWidthAutoSize(Optional ByVal lCol As Long = C_NULL_RESULT)
Dim lngC As Long
  
   If lCol = C_NULL_RESULT Then
      For lngC = 0 To UBound(m_cols)
         Call ColSizing(lngC)
      Next lngC
   Else
      Call ColSizing(lCol)
   End If
      
End Sub

Private Sub ColSizing(ByVal lCol As Long)
Dim lngLW As Long, lngCW As Long, lRow As Long
Dim strTemp As String

For lRow = 0 To UBound(m_items) - 1
  strTemp = m_items(lRow).Item(lCol).Text
  lngLW = UserControl.TextWidth(strTemp)
  If lngCW < lngLW Then lngCW = lngLW
Next lRow

  m_cols(lCol).Width = lngCW + Screen.TwipsPerPixelX
End Sub

Public Sub ShowList()
Dim lW  As Long
Dim lH  As Long
Dim lT  As Long
Dim Rct As RECT
Dim PT  As POINTAPI

    GetWindowRect m_hWnd, Rct
    lW = IIf(m_DropW, m_DropW, Rct.r - Rct.L)
    lH = ((m_VisibleRows * m_RowH) + lHeaderH)

    SetParent UserControl.hwnd, 0
    lT = IIf(Rct.B + UserControl.ScaleHeight > Screen.Height / Screen.TwipsPerPixelY, Rct.T - (UserControl.ScaleHeight + 1), Rct.B + 1)
    
    ''Call SetWindowLong(UserControl.hwnd, -20, GetWindowLong(UserControl.hwnd, -20) Or &H80)
    SetWindowPos UserControl.hwnd, HWND_TOPMOST, Rct.L, lT, lW, lH, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Call SetWindowLong(UserControl.hwnd, -8, Parent.hwnd)
    
    UpdateScrollV
    Call DrawGrid
    
    SetFocusEx UserControl.hwnd
    SetFocusEx m_hWnd
    'SetCapture UserControl.hwnd
            
End Sub

Public Sub HideList()
    'SetParent UserControl.hwnd, m_PhWnd
    'SetWindowPos UserControl.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_HIDEWINDOW
    ShowWindow UserControl.hwnd, 0
    m_Visible = False
End Sub

Property Get BackColor() As OLE_COLOR: BackColor = UserControl.BackColor: End Property
Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor = Value
    PropertyChanged "BackColor"
End Property

Property Get ColumnCount() As Long
On Local Error Resume Next
    ColumnCount = UBound(m_cols) + 1
End Property
Property Get ItemCount() As Long
On Local Error Resume Next
    ItemCount = UBound(m_items) + 1
End Property

Property Get ItemText(ByVal Item As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    ItemText = m_items(Item).Item(Column).Text
End Property
Property Let ItemText(ByVal Item As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
    If m_items(Item).Item(Column).Text = Value Then Exit Property
    m_items(Item).Item(Column).Text = Value
End Property

Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Property Set Font(ByVal Value As StdFont)
    Set UserControl.Font() = Value
    PropertyChanged "Font"
End Property
Property Get Header() As Boolean: Header = m_Header: End Property
Property Let Header(Value As Boolean)
    m_Header = Value
    PropertyChanged "Header"
End Property
Property Get HeaderHeight() As Long: HeaderHeight = m_HeaderH: End Property
Property Let HeaderHeight(ByVal Value As Long)
    m_HeaderH = Value
    UpdateValues
    PropertyChanged "HeaderH"
End Property
Property Get ItemHeight() As Long: ItemHeight = m_ItemH: End Property
Property Let ItemHeight(ByVal Value As Long)
    m_ItemH = Value
    UpdateValues
    PropertyChanged "ItemH"
End Property
Property Get GridLineColor() As OLE_COLOR: GridLineColor = m_GridLineColor: End Property
Property Let GridLineColor(ByVal Value As OLE_COLOR)
    m_GridLineColor = Value
    PropertyChanged "LineColor"
    DrawGrid
End Property
Property Get GridLineStyle() As ScrollBarConstants: GridLineStyle = m_GridStyle: End Property
Property Let GridLineStyle(ByVal Value As ScrollBarConstants)
    m_GridStyle = Value
    UpdateValues
    PropertyChanged "GridStyle"
End Property
Property Get StripedGrid() As Boolean: StripedGrid = m_Striped: End Property
Property Let StripedGrid(ByVal Value As Boolean)
    m_Striped = Value
    PropertyChanged "Striped"
End Property
Property Get StripBackColor() As OLE_COLOR: StripBackColor = m_StripedColor: End Property
Property Let StripBackColor(ByVal Value As OLE_COLOR)
    m_StripedColor = Value
    PropertyChanged "StripedColor"
End Property
Property Get SelectionColor() As OLE_COLOR: SelectionColor = m_SelColor: End Property
Property Let SelectionColor(ByVal Value As OLE_COLOR)
    m_SelColor = Value
    PropertyChanged "SelColor"
End Property

Public Property Get SelectedIndex() As Long
  SelectedIndex = t_Row
End Property

Property Get BorderColor() As OLE_COLOR: BorderColor = m_BorderColor: End Property
Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    PropertyChanged "BorderColor"
End Property
Property Get ForeColor() As OLE_COLOR: ForeColor = m_ForeColor: End Property
Property Let ForeColor(ByVal Value As OLE_COLOR)
    m_ForeColor = Value
    PropertyChanged "ForeColor"
End Property

Property Get VisibleRows() As Long: VisibleRows = m_VisibleRows: End Property
Property Let VisibleRows(ByVal Value As Long)
    m_VisibleRows = Value
    PropertyChanged "VisibleRows"
End Property

Property Get DropWidth() As Long: DropWidth = m_DropW: End Property
Property Let DropWidth(ByVal Value As Long)
    m_DropW = Value
    PropertyChanged "DropWidth"
End Property

Public Property Get ColumnInBox() As Long
    ColumnInBox = m_ColRet
End Property

Public Property Let ColumnInBox(ByVal NewColumnInBox As Long)
On Error Resume Next
If UBound(m_cols) > NewColumnInBox Then
    m_ColRet = NewColumnInBox
Else
    m_ColRet = UBound(m_cols)
End If
    PropertyChanged "ColumnInBox"
End Property

Private Property Get lHeaderH() As Long
    lHeaderH = IIf(m_Header = True, m_HeaderH, 0)
End Property

Private Property Get lGridH() As Long
    lGridH = (m_VisibleRows * m_RowH) + lHeaderH + (5 * e_Scale)
    'lGridH = UserControl.ScaleHeight - lHeaderH
End Property

Private Function GetScroll(eBar As EFSScrollBarConstants) As Long
    GetScroll = IIf(cScroll.Visible(eBar), cScroll.Value(eBar), 0)
End Function

Private Function SendText(Text As String)
    SendMessage m_hWnd, &HC, 0&, Text
    SendMessage m_hWnd, &HB1, Len(Text), Len(Text)
End Function

Public Property Get hdc() As Long
  hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long
  hwnd = UserControl.hwnd
End Property

Public Property Get ScaleHeight() As Long
  ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ScaleWidth() As Long
  ScaleWidth = UserControl.ScaleWidth
End Property

Private Sub UpdateValues()
Dim TH As Integer
Dim Px As Long
    
    Px = 6 * e_Scale
    TH = UserControl.TextHeight("ÀJ")
    
    If TH + Px > m_ItemH Then m_ItemH = TH + Px
    If m_img.Y + Px > m_ItemH Then m_ItemH = m_img.Y + Px
    m_RowH = m_ItemH
    If m_GridStyle = 1 Or m_GridStyle = 3 Then m_RowH = m_RowH + (1 * lnScale)
    UpdateScrollV
End Sub

Private Sub UpdateScrollV()
On Local Error Resume Next
Dim lHeight     As Long
Dim lProportion As Long
Dim ly          As Long
Dim bFlag       As Boolean

    ly = lGridH
    lHeight = ((ItemCount * m_ItemH) + (2 * e_Scale)) - ly
    
    If (lHeight > 0) Then
      lProportion = lHeight \ (ly + 1)
      cScroll.LargeChange(1) = lHeight \ lProportion
      cScroll.Max(1) = IIf(m_Header, lHeight + (m_ItemH * 3), lHeight + (m_ItemH * 2))
      cScroll.Visible(1) = True
    Else
      cScroll.Visible(1) = False
    End If
    
End Sub

Private Function GetRowFromY(ByVal Y As Long) As Long
    If m_Header And Y <= lHeaderH Then
        GetRowFromY = -1
        Exit Function
    End If
    Y = Y + GetScroll(1) - lHeaderH
    GetRowFromY = Y \ m_RowH
    If GetRowFromY >= ItemCount Then GetRowFromY = -1
End Function

Private Function IsVisibleRow(ByVal eRow As Long) As Boolean
'On Error Resume Next
Dim Y As Long
    If cScroll.Visible(1) = False Then
      IsVisibleRow = True
    Else
      Y = (eRow * m_RowH) - GetScroll(1)
      IsVisibleRow = (Y + m_ItemH > 0) And Y <= lGridH
    End If
    
    Debug.Print "IsVisibleRow=" & IsVisibleRow
End Function

Private Function IsCompleteVisibleRow(eRow As Long) As Boolean
'On Local Error Resume Next
Dim Y       As Long
Dim bRow    As Boolean
    Y = (eRow * m_RowH) - GetScroll(1)
    bRow = (Y >= 0) And (Y + m_ItemH >= lGridH)
    IsCompleteVisibleRow = bRow
    Debug.Print "IsCompleteVisibleRow=" & bRow
End Function

Private Sub SetVisibleItem(eRow As Long)
On Error GoTo zErr
Dim lx  As Integer
Dim ly  As Integer

    If eRow = -1 Then Exit Sub
    ly = eRow * m_RowH

    '?Vertical
    If (ly + m_RowH) - lGridH > GetScroll(1) Then
        cScroll.Value(1) = ((ly + m_RowH)) - lGridH
    ElseIf ly < GetScroll(1) Then
        cScroll.Value(1) = ly
    End If
zErr:
    DrawGrid
End Sub

Public Sub Refresh()
DrawGrid
End Sub

Private Sub DrawGrid()
On Local Error Resume Next
Dim lCol    As Long
Dim lRow    As Long
Dim ly      As Long
Dim lx      As Long
Dim lColW   As Long
Dim dvc     As Long
Dim iRct    As RECT
Dim tRct    As RECT
Dim lPx     As Long
Dim lPx2    As Long

    UserControl.Cls
    
    lCol = 0
    lRow = 0

    ly = -GetScroll(1)
    dvc = UserControl.hdc

    ly = ly + lHeaderH
   ' Debug.Print "ly=" & ly

    Do While lRow <= ItemCount - 1 And ly < UserControl.ScaleHeight
        
        If ly + m_RowH > 0 Then '?Visible
            
            SetRect iRct, 0, ly, UserControl.ScaleWidth, ly + m_ItemH
            If m_Striped And lRow Mod 2 Then _
                DrawBack dvc, SysColor(m_StripedColor), iRct
            
            '\ Seleccion
            lPx2 = m_GridW - lnScale
            If lPx2 > UserControl.ScaleWidth + (8 * lnScale) Then lPx2 = UserControl.ScaleWidth + (8 * lnScale)
            If lRow = m_SelRow Then
                DrawSelection dvc, 0, ly, lPx2, m_ItemH, 1
            ElseIf lRow = t_Row Then
                DrawSelection dvc, 0, ly, lPx2, m_ItemH, 1
            End If
                
            lPx2 = lnScale \ 2
            '?GridLines 0N,1H,2V,3B -> Horizontal
            If m_GridStyle = 1 Or m_GridStyle = 3 Then _
                DrawLine dvc, lx, ly + m_ItemH + lPx2, UserControl.ScaleWidth, ly + m_ItemH + lPx2, m_GridLineColor
            
             Do While lCol < ColumnCount And lx < UserControl.ScaleWidth
             
                lColW = m_cols(lCol).Width
                If m_GridStyle = 2 Or m_GridStyle = 3 Then lColW = lColW - lnScale
                
                    SetRect iRct, lx, ly, lx + lColW, ly + m_ItemH
                   
                    '?GridLines 0N,1H,2V,3B - > Vertical
                     If m_GridStyle = 2 Or m_GridStyle = 3 Then _
                       DrawLine dvc, lx + lColW + lPx2, ly, lx + lColW + lPx2, ly + m_ItemH, m_GridLineColor
                       
                    If Trim(m_items(lRow).Item(lCol).Text) <> vbNullString Then
                        
                        SetRect tRct, lx + 4 + lPx, ly, lx + lColW - 3, ly + m_ItemH
                        If tRct.r < tRct.L Then tRct.r = tRct.L
                        If tRct.r > tRct.L Then _
                        DrawText dvc, m_items(lRow).Item(lCol).Text, Len(m_items(lRow).Item(lCol).Text), tRct, GetTextFlag(lCol)
                    
                    End If
eDrawNext:
                lx = lx + m_cols(lCol).Width
                lCol = lCol + 1
                
             Loop
            '?Reset to Scroll Position
            lCol = 0
            lx = 0
        End If
        
        ly = ly + m_RowH
        lRow = lRow + 1
    Loop
    Call DrawHeader
End Sub

Private Function GetTextFlag(Col As Long) As Long
    GetTextFlag = &H4 Or &H20 Or &H40000 '-> VCenter Or SingleLine Or WordElipsis
    Select Case m_cols(Col).Aling
        Case 1: GetTextFlag = GetTextFlag Or &H2
        Case 2: GetTextFlag = GetTextFlag Or &H1
    End Select
End Function

Private Function DrawHeader()
Dim uTheme  As Long
Dim uRct    As RECT
Dim Col    As Long
Dim lx      As Long
Dim lW      As Long

    uTheme = OpenThemeData(UserControl.hwnd, StrPtr("Header"))
    If uTheme = 0 Then Exit Function
    
    SetRect uRct, 0, 0, UserControl.ScaleWidth, lHeaderH
    Call DrawThemeBackground(uTheme, UserControl.hdc, 0, 0&, uRct, ByVal 0&)
    
    Do While Col < ColumnCount And lx < UserControl.ScaleWidth
    
        lW = m_cols(Col).Width
        SetRect uRct, lx, 0, lx + lW, lHeaderH
        
        Call DrawThemeBackground(uTheme, UserControl.hdc, 1, 1, uRct, ByVal 0&)
        
        uRct.r = uRct.r - (10 * 1)
        OffsetRect uRct, 5 * 1, 0
        
        DrawText UserControl.hdc, m_cols(Col).Text, Len(m_cols(Col).Text), uRct, GetTextFlag(Col)
        
        lx = lx + m_cols(Col).Width
        Col = Col + 1
        
    Loop
    
End Function
Private Function SysColor(oColor As Long) As Long
    OleTranslateColor2 oColor, 0, SysColor
End Function

Private Sub DrawSelection(lpDC As Long, X As Long, Y As Long, W As Long, H As Long, lIndex As Long)
Dim hBmp    As Long
Dim DC      As Long
Dim hDCMem  As Long
Dim hPen    As Long
Dim Alpha1  As Long
Dim lColor  As Long
Dim lH      As Long
Dim Px      As Long
Dim out     As Long
Dim i       As Long
Dim DivValue    As Double


    Select Case lIndex
        Case 0: lColor = pvAlphaBlend(m_SelColor, vbWhite, 110)
        Case 1: lColor = pvAlphaBlend(m_SelColor, vbWhite, 190)
        Case 2: lColor = m_SelColor
    End Select

    Px = lnScale \ 2
    out = Px \ 2
    lH = H - (2 * lnScale)

    DC = GetDC(0)
    hDCMem = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, 1, lH)
    Call SelectObject(hDCMem, hBmp)
    
    Alpha1 = pvAlphaBlend(lColor, vbWhite, 45)
    For i = 0 To lH
        DivValue = ((i * 100) / lH)
        SetPixelV hDCMem, 0, i, pvAlphaBlend(lColor, Alpha1, DivValue)
    Next
    
    StretchBlt lpDC, X + lnScale, Y + lnScale, W - (lnScale * 2), lH, hDCMem, 0, 0, 1, lH, vbSrcCopy
    
    hPen = CreatePen(0, lnScale, lColor)
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X + Px, Y + Px, X + W - out, Y + H - out, 3 * lnScale, 3 * lnScale
    DeleteObject hPen
    
    hPen = CreatePen(0, lnScale, pvAlphaBlend(lColor, vbWhite, 18))
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X + Px + lnScale, Y + Px + lnScale, X + W - (lnScale + out), Y + H - (lnScale + out), 3 * lnScale, 3 * lnScale
    
    DeleteObject hPen
    DeleteObject hBmp
    DeleteDC DC
    DeleteDC hDCMem
    
End Sub

Private Function pvAlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
Dim clrFore(3)      As Byte
Dim clrBack(3)      As Byte

    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
    
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
    CopyMemory pvAlphaBlend, clrFore(0), 4
End Function

Private Sub DrawBorder(lpDC As Long, Color As Long, X As Long, Y As Long, W As Long, H As Long)
Dim hPen As Long
    hPen = CreatePen(0, lnScale, Color)
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, X, Y, X + W, Y + H, 0, 0
    DeleteObject hPen
End Sub

Private Sub DrawBack(lpDC As Long, Color As Long, Rct As RECT)
Dim hBrush  As Long

    hBrush = CreateSolidBrush(Color)
    Call FillRect(lpDC, Rct, hBrush)
    Call DeleteObject(hBrush)
    
End Sub

Private Sub DrawLine(lpDC As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color As Long)
Dim PT      As POINTAPI
Dim hPen    As Long
Dim hPenOld As Long

    hPen = CreatePen(0, lnScale, Color)
    hPenOld = SelectObject(lpDC, hPen)
    Call MoveToEx(lpDC, X, Y, PT)
    Call LineTo(lpDC, X2, Y2)
    Call SelectObject(lpDC, hPenOld)
    Call DeleteObject(hPen)
    
End Sub

Private Function GetWindowsDPI() As Double
Dim hdc As Long
Dim lPx  As Double

Const LOGPIXELSX As Long = 88

    hdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc
    
    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
    
End Function

Private Sub WndProc(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hwnd As Long, _
       ByVal uMsg As ssc_eMsg, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)

    
'    Select Case hwnd
'
'        Case UserControl.hwnd
            Select Case uMsg
                Case WM_WINDOWPOSCHANGING
                    HideList
                Case WM_MOUSELEAVE
                    If t_Row <> -1 Then
                        t_Row = -1
                        DrawGrid
                    End If
                    m_bTrack = False
                    'RaiseEvent MouseExit
                Case WM_KILLFOCUS
                    HideList
                    'Debug.Print "bye"
                Case WM_MOUSEWHEEL
                    'Debug.Print "Wheel"
                Case WM_NCPAINT
                    If UserControl.BorderStyle = 0 Then Exit Sub
                    Dim Rct As RECT
                    Dim DC As Long
                    Dim ix As Long
                    Dim zs As Long
                        
                    DC = GetWindowDC(hwnd)
                    GetWindowRect hwnd, Rct
                            
                    Rct.r = Rct.r - Rct.L
                    Rct.B = Rct.B - Rct.T
                    Rct.L = IIf(lnScale > 1, 1, 0)
                    Rct.T = IIf(lnScale > 1, 1, 0)
                    
                    ix = GetSystemMetrics(6)
                    zs = ix * lnScale
                    ExcludeClipRect DC, zs + lnScale, zs + lnScale, Rct.r - (zs + lnScale), Rct.B - (zs + lnScale)
                    
                    Dim hPen        As Long
                    Dim OldPen      As Long

                    hPen = CreatePen(0, lnScale, m_BorderColor)
                    OldPen = SelectObject(DC, hPen)
                    Rectangle DC, Rct.L, Rct.T, Rct.r, Rct.B
                    Call SelectObject(DC, OldPen)
                    DeleteObject hPen
                    ReleaseDC hwnd, DC

            End Select
            'WM_CONTEXTMENU, WM_LBUTTONUP, WM_NCACTIVATE, WM_ACTIVATE, WM_COMMAND788792
'        Case m_hWnd
'            Select Case uMsg
'                Case WM_KILLFOCUS
'                     HideList
'                     'Debug.Print "m_hWnd WM_KILLFOCUS"
'            End Select
'        Case m_PhWnd
'            Select Case uMsg
'                Case WM_SIZE, WM_MOVE, WM_WINDOWPOSCHANGING, WM_LBUTTONDOWN, WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONUP, WM_MENUCOMMAND, 164, WM_SYSCOMMAND
'                    HideList
'                    'Debug.Print "m_PhWnd WM_KILLFOCUS"
'            End Select
'
'    End Select
End Sub



