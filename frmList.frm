VERSION 5.00
Begin VB.Form ListBoxForm 
   Caption         =   "ListBox Example"
   ClientHeight    =   3030
   ClientLeft      =   2355
   ClientTop       =   3645
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   3930
   Begin VB.CheckBox chkMulti 
      Caption         =   " Mutli-select (not finished yet)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   2760
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3225
      TabIndex        =   3
      Top             =   2370
      Width           =   570
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Text            =   "add"
      Top             =   2370
      Width           =   2985
   End
   Begin VB.VScrollBar lstScroll 
      Height          =   2130
      Left            =   3510
      Max             =   0
      TabIndex        =   0
      Top             =   165
      Width           =   255
   End
   Begin VB.PictureBox picList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   180
      ScaleHeight     =   2055
      ScaleWidth      =   3285
      TabIndex        =   1
      Top             =   165
      Width           =   3345
      Begin VB.PictureBox picListBuffer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -30
         ScaleHeight     =   2055
         ScaleWidth      =   3885
         TabIndex        =   4
         Top             =   -30
         Visible         =   0   'False
         Width           =   3945
         Begin VB.Timer tmrScroll 
            Enabled         =   0   'False
            Interval        =   3000
            Left            =   645
            Top             =   375
         End
      End
   End
End
Attribute VB_Name = "ListBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private topIndex    As Long
Private listCount   As Long
Private List()      As String
Private Selected()  As Boolean
Private MultiSelect As Boolean
Private SelectedInd As Integer

Private lngHLBGclr  As Long
Private lngHLForeClr    As Long
Private lngForeClr  As Long
Private lngBorderClr    As Long
Sub TimeOut(Duration)
    StartTime = Timer
    Do While Timer - StartTime < Duration
        X = DoEvents()
    Loop
End Sub


Private Sub AddItem(strItem As String)
    listCount = listCount + 1
    ReDim Preserve List(1 To listCount) As String
    List(listCount) = strItem
    ReDim Preserve Selected(1 To listCount) As Boolean
    
    lstScroll.Min = 1
    lstScroll.Max = listCount
    '* Update..
    PrintList
End Sub

Private Sub PrintList()
    DoEvents
    picListBuffer.Cls
    
    Dim intItemsToPrint As Integer, itemHeight As Integer
    Dim MaxItem As Integer, i As Integer, curY As Integer
    curY = 0
    
    itemHeight = picListBuffer.TextHeight("Ab¯_")
    intItemsToPrint = picListBuffer.ScaleHeight / itemHeight + 1
    
    MaxItem = topIndex + intItemsToPrint
    If MaxItem > listCount Then MaxItem = listCount
    
    For i = topIndex To MaxItem
        If Selected(i) Then
            picListBuffer.Line (0, curY)-(picList.ScaleWidth - 20, curY + itemHeight), lngHLBGclr, BF
            picListBuffer.CurrentY = curY
            picListBuffer.ForeColor = lngHLForeClr
        Else
            picListBuffer.ForeColor = lngForeClr
        End If
        
        If SelectedInd = i Then
            picListBuffer.Line (0, curY)-(picList.ScaleWidth - 20, curY + itemHeight), lngBorderClr, B
            picListBuffer.CurrentY = curY
        End If
        
        curY = curY + itemHeight
        picListBuffer.CurrentX = 10
        picListBuffer.Print List(i)
    Next i
    
    picList.Picture = picListBuffer.Image
End Sub

Private Sub ResetSelected()
    ReDim Selected(1 To listCount) As Boolean
End Sub

Sub UpdateScroll()
    lstScroll.Min = 1
    lstScroll.Max = listCount
    lstScroll.Value = topIndex
End Sub

Private Sub chkMulti_Click()
    If chkMulti.Value Then
        MultiSelect = True
    Else
        MultiSelect = False
    End If
    
End Sub

Private Sub Command1_Click()
    AddItem Text1
    PrintList
End Sub


Private Sub Command2_Click()
    For i = 1 To 20
        AddItem Text1 & i
    Next i
End Sub


Private Sub Form_Load()
    topIndex = 1
    lngForeClr = &H80000008
    lngHLBGclr = &H8000000D
    lngHLForeClr = &H8000000E
    lngBorderClr = &HFFC0C0
End Sub

Private Sub lstScroll_Change()
    topIndex = lstScroll.Value
    
    '* Update...
    PrintList
End Sub


Private Sub lstScroll_Scroll()
    topIndex = lstScroll.Value
    
    '* Update...
    PrintList

End Sub


Private Sub picList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then    'down
        If SelectedInd = listCount Then Exit Sub
        SelectedInd = SelectedInd + 1
        
        If MultiSelect = False Then
            ResetSelected
        End If
        Selected(SelectedInd) = True
        PrintList
    ElseIf KeyCode = 38 Then    'up
        If SelectedInd = 1 Then Exit Sub
        SelectedInd = SelectedInd - 1
        
        If MultiSelect = False Then
            ResetSelected
        End If
        Selected(SelectedInd) = True
        PrintList
    End If
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ItemClicked As Integer, itemHeight As Integer
    itemHeight = picListBuffer.TextHeight("AbWw_¯")
    
    ItemClicked = Y \ itemHeight + topIndex
    
    If ItemClicked > listCount Or ItemClicked < 1 Then Exit Sub
    
    If MultiSelect = False Then
        ResetSelected
        Selected(ItemClicked) = True
    Else
        Selected(ItemClicked) = Not Selected(ItemClicked)
    End If
    SelectedInd = ItemClicked
    
    PrintList
End Sub


Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim ItemClicked As Integer, itemHeight As Integer
        itemHeight = picListBuffer.TextHeight("AbWw_¯")
        
        ItemClicked = Y \ itemHeight + topIndex
        
        If ItemClicked > listCount Or ItemClicked < 1 Then Exit Sub
        
        If MultiSelect = False Then
            ResetSelected
            Selected(ItemClicked) = True
        Else
            Selected(ItemClicked) = True
        End If
        SelectedInd = ItemClicked
        
        If Y < 0 Then
            'scroll up
            If topIndex = 1 Then Exit Sub
            topIndex = topIndex - 1
            UpdateScroll
            tmrScroll.Tag = "up"
            If tmrScroll.Enabled = False Then tmrScroll.Enabled = True
        ElseIf Y > picList.ScaleHeight Then
            'scroll down
            If topIndex = listCount Then Exit Sub
            topIndex = topIndex + 1
            UpdateScroll
            tmrScroll.Tag = "down"
            If tmrScroll.Enabled = False Then tmrScroll.Enabled = True
        Else
            tmrScroll.Enabled = False
        End If
        
        UpdateScroll
        DoEvents
        PrintList
    End If
End Sub


Private Sub tmrScroll_Timer()
    UpdateScroll
    If tmrScroll.Tag = "up" Then
        If topIndex > 1 Then topIndex = topIndex - 1
    ElseIf tmrScroll.Tag = "down" Then
        If topIndex < listCount Then topIndex = topIndex + 1
    End If
End Sub


