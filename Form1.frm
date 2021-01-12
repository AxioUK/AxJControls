VERSION 5.00
Object = "{D53004A4-B6ED-4733-B6AE-C4684DB2DB7D}#9.0#0"; "AxJControls.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "axColCombo"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9720
   FillColor       =   &H00C07000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin AxJControls.axJList axJList1 
      Height          =   480
      Left            =   2595
      TabIndex        =   18
      Top             =   4395
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      HeaderH         =   22
      LineColor       =   15790320
      GridStyle       =   3
      Striped         =   -1  'True
      StripedColor    =   16645629
      SelColor        =   -2147483635
      BorderColor     =   9471874
      Header          =   -1  'True
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      VisibleRows     =   8
      DropWidth       =   0
   End
   Begin AxJControls.axJColCombo axJColCombo1 
      Height          =   390
      Left            =   3510
      TabIndex        =   17
      Top             =   2790
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty lFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      lstStyle        =   0
      ForeColor       =   0
      ForeColorList   =   0
      BackColor       =   16777215
      BackColorList   =   16777215
      BorderColor     =   11447982
      SelectColor     =   16737792
      Shape           =   0
      sdBack          =   0   'False
      sFocus          =   0   'False
      Text            =   ""
      HandCursor      =   0   'False
      aComplete       =   0   'False
      Hint            =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   3015
      TabIndex        =   7
      Top             =   3345
      Width           =   345
   End
   Begin VB.OptionButton Option1 
      Caption         =   "JList"
      Height          =   240
      Index           =   1
      Left            =   2790
      TabIndex        =   4
      Top             =   510
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Caption         =   "JCombo"
      Height          =   240
      Index           =   0
      Left            =   2790
      TabIndex        =   3
      Top             =   240
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Header"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   165
      Width           =   1260
   End
   Begin VB.Frame frame1 
      Caption         =   "Frame1"
      Height          =   2280
      Left            =   6270
      TabIndex        =   10
      Top             =   2790
      Width           =   3255
      Begin VB.TextBox txt4 
         Height          =   330
         Left            =   255
         TabIndex        =   22
         Top             =   1845
         Width           =   2820
      End
      Begin VB.TextBox txt3 
         Height          =   330
         Left            =   255
         TabIndex        =   21
         Top             =   1500
         Width           =   2820
      End
      Begin VB.TextBox txt2 
         Height          =   330
         Left            =   255
         TabIndex        =   20
         Top             =   1155
         Width           =   2820
      End
      Begin VB.TextBox txt1 
         Height          =   330
         Left            =   255
         TabIndex        =   19
         Top             =   810
         Width           =   2820
      End
      Begin AxJControls.axJCombo axJCombo1 
         Height          =   405
         Left            =   210
         TabIndex        =   9
         Top             =   315
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty lFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         lstStyle        =   0
         Fore            =   0
         lFore           =   0
         cBack           =   16777215
         cLBack          =   16777215
         cBorder         =   11447982
         cSelection      =   16737792
         Shape           =   0
         sdBack          =   0   'False
         sFocus          =   0   'False
         Text            =   ""
         HandCursor      =   0   'False
         aComplete       =   0   'False
         Hint            =   ""
      End
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1785
      TabIndex        =   2
      Text            =   "0"
      Top             =   495
      Width           =   420
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   180
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   3825
      Width           =   2820
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   180
      TabIndex        =   5
      Text            =   "Text1"
      ToolTipText     =   "Click para mostrar List, ESC para ocultar..."
      Top             =   2865
      Width           =   2970
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3345
      Width           =   2820
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   1365
      Left            =   6330
      TabIndex        =   16
      Top             =   1575
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":00DE
      Height          =   780
      Left            =   3390
      TabIndex        =   15
      Top             =   1590
      Width           =   2610
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ListBox multicolumna oculto, se puede activar al tomar foco el control Text anclado o por un boton al pasarle el sub ShowList."
      Height          =   975
      Left            =   285
      TabIndex        =   14
      Top             =   1590
      Width           =   2730
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "axJCombo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   7020
      TabIndex        =   13
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "axColCombo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3945
      TabIndex        =   12
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "axJList"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1095
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ColumnInBox"
      Height          =   195
      Left            =   660
      TabIndex        =   1
      Top             =   555
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsVisible As Boolean


Private Sub axJCombo1_ItemClick(Item As Integer, strCol1 As String, strCol2 As String, strCol3 As String, strCol4 As String)
txt1.Text = strCol1
txt2.Text = strCol2
txt3.Text = strCol3
txt4.Text = strCol4
End Sub

Private Sub Check1_Click()
axJColCombo1.Header = Check1.Value
axJList1.Header = Check1.Value
End Sub

Private Sub chkLK_Click()
'axJColCombo1.ComboStyle = chkLK.Value
End Sub

Private Sub Command2_Click()
axJList1.Init Text2.hWnd
If IsVisible = False Then
  axJList1.ShowList
  IsVisible = True
Else
  axJList1.HideList
  IsVisible = False
End If
End Sub

Private Sub Form_Load()
'gbAllowSubclassing = True
'SubclassToSeeMessages Me.hWnd

IsVisible = False

Dim i As Long

    With axJColCombo1
        .AddColumn "User Name"
        .AddColumn "Last Name"
        .AddColumn "Age"

        For i = 1 To 20
            .AddItem "My user " & i, 0
            .ItemText(.ItemCount - 1, 1) = "My last_name " & i
            .ItemText(.ItemCount - 1, 2) = 25 + i
        Next
        .ColWidthAutoSize
        .ColumnInBox = CInt(Text4.Text)
    End With
    
    With axJList1
        .AddColumn "User Name"
        .AddColumn "Last Name"
        .AddColumn "Age"

        For i = 1 To 20
            .AddItem "My user " & i, 0
            .ItemText(.ItemCount - 1, 1) = "My last_name " & i
            .ItemText(.ItemCount - 1, 2) = 25 + i
        Next
        .Init Text1.hWnd
        .Header = Check1.Value
        .HeaderHeight = 17
        .ItemHeight = 17
        .ColWidthAutoSize
    End With
    
    With axJCombo1
       ' .CreateImageList 20, 20, Image1, vbBlack
        For i = 0 To 20
            .AddItem "Item" & i & "_Col0", "Item" & i & "_Col1", "Item" & i & "_Col2", "Item" & i & "_Col3", 1
        Next
    End With

End Sub

Private Sub axJColCombo1_ItemClick(ByVal Item As Long)
    Text1 = axJColCombo1.ItemText(Item, 0)
    Text2 = axJColCombo1.ItemText(Item, 1)
    Text3 = axJColCombo1.ItemText(Item, 2)

    Text1.SelStart = 0
    Text1.SelLength = Len(Text2)

End Sub

Private Sub axJList1_ItemClick(Item As Long)
Text1.Text = axJList1.ItemText(Item, 0)
Text2.Text = axJList1.ItemText(Item, 1)
Text3.Text = axJList1.ItemText(Item, 2)
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
  Case 0
    axJColCombo1.ListStyle = JclsCombo
    axJCombo1.ListStyle = JclsCombo
  Case 1
    axJColCombo1.ListStyle = JclsList
    axJCombo1.ListStyle = JclsList
End Select
End Sub

Private Sub Text1_GotFocus()
axJList1.Init Text1.hWnd
axJList1.ShowList
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then axJList1.HideList

End Sub

Private Sub Text4_Change()
On Error Resume Next
axJColCombo1.ColumnInBox = CInt(Text4.Text)
axJList1.ColumnInBox = CInt(Text4.Text)
axJCombo1.ColumnInList = CInt(Text4.Text)
End Sub
