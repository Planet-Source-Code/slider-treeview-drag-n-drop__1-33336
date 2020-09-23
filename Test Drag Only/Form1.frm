VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fTestDrag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test Drag and Drop Function Only"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView oTree 
      Height          =   5475
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9657
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilDialog 
      Left            =   2730
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2236
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fTestDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================================================
' Debugging... Saves adding the debug statements to the form events
'
#Const DEBUGMODE = 0                    '## 0=No debug
                                        '   1=debug
#Const MOUSEEVENTS = 0                  '## 0=No mouse events
                                        '   1=Mouse Up & Mouse Down
                                        '   2=All Mouse events
#If DEBUGMODE = 1 Then
    Private dbgFormName  As String
#End If

'===========================================================================
' Private: Variables and Declarations
'
Private Enum eCodeScrollView     '## Scroll Treeview
    [Home] = 0
    [Page Up] = 1
    [Up] = 2
    [Down] = 3
    [Page Down] = 4
    [End] = 5
    [Left] = 6
    [Page Left] = 7
    [Line Left] = 8
    [Line Right] = 9
    [Page Right] = 10
    [Right] = 11
End Enum

Private moDragNode        As MSComctlLib.Node
Private moInDrag          As MSComctlLib.Node
Private mbDragEnabled     As Boolean
Private mbStartDrag       As Boolean
Private mbInDrag          As Boolean
Private mlNodeHeight      As Long
Private mlDragExpandTime  As Long
Private mlDragScrollTime  As Long
Private mlAutoScroll      As Long     '## Distance in which auto-scrolling happens
Private mszDrag           As Size           '## X and Y distance cursor moves before dragging begins, in pixels
Private mptBtnDown        As POINTAPI

Private WithEvents moDragExpand As XTimer
Attribute moDragExpand.VB_VarHelpID = -1
Private WithEvents moDragScroll As XTimer
Attribute moDragScroll.VB_VarHelpID = -1

'===========================================================================
' Private: APIs
'
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As SB_Type, lpsi As SCROLLINFO) As Boolean
Private Declare Function PtInRect Lib "user32" (lprc As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Enum RectFlags
    rfLeft = &H1
    rfTop = &H2
    rfRight = &H4
    rfBottom = &H8
End Enum

Private Enum ScrollDirectionFlags
    sdLeft = &H1
    sdUp = &H2
    sdRight = &H4
    sdDown = &H8
End Enum

Private Enum SB_Type
    SB_HORZ = 0
    SB_VERT = 1
    SB_CTL = 2
    SB_BOTH = 3
End Enum

Private Enum SIF_Mask
    SIF_RANGE = &H1
    SIF_PAGE = &H2
    SIF_POS = &H4
    SIF_DISABLENOSCROLL = &H8
    SIF_TRACKPOS = &H10
    SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum

Private Type SCROLLINFO
    cbSize    As Long
    fMask     As SIF_Mask
    nMin      As Long
    nMax      As Long
    nPage     As Long
    nPos      As Long
    nTrackPos As Long
End Type

Private Type Size
  cx As Long
  cy As Long
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI   ' pt
    x As Long
    y As Long
End Type

Private Const TV_FIRST            As Long = &H1100

Private Const TVM_DELETEITEM      As Long = (TV_FIRST + 1)
Private Const TVM_GETITEMRECT     As Long = (TV_FIRST + 4)
Private Const TVM_SETIMAGELIST    As Long = (TV_FIRST + 9)
Private Const TVM_GETNEXTITEM     As Long = (TV_FIRST + 10)
Private Const TVM_SELECTITEM      As Long = (TV_FIRST + 11)
Private Const TVM_GETITEM         As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM         As Long = (TV_FIRST + 13)
Private Const TVM_HITTEST         As Long = (TV_FIRST + 17)
Private Const TVM_CREATEDRAGIMAGE As Long = (TV_FIRST + 18)

'## TVM_GETNEXTITEM wParam values
Public Enum TVGN_Flags
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
'#If (WIN32_IE >= &H400) Then   ' >= Comctl32.dll v4.71
    TVGN_LASTVISIBLE = &HA
'#End If
End Enum

Private Const GWL_STYLE         As Long = (-16)

Private Const SM_CXDRAG         As Long = &H44
Private Const SM_CYDRAG         As Long = &H45

'---------------------------------------------------------------------------

' Scroll Bar Commands
Private Const SB_LINEUP         As Long = 0
Private Const SB_LINELEFT       As Long = 0
Private Const SB_LINEDOWN       As Long = 1
Private Const SB_LINERIGHT      As Long = 1
Private Const SB_PAGEUP         As Long = 2
Private Const SB_PAGELEFT       As Long = 2
Private Const SB_PAGEDOWN       As Long = 3
Private Const SB_PAGERIGHT      As Long = 3
Private Const SB_THUMBPOSITION  As Long = 4
Private Const SB_THUMBTRACK     As Long = 5
Private Const SB_TOP            As Long = 6
Private Const SB_LEFT           As Long = 6
Private Const SB_BOTTOM         As Long = 7
Private Const SB_RIGHT          As Long = 7
Private Const SB_ENDSCROLL      As Long = 8

Private Const WM_HSCROLL        As Long = &H114
Private Const WM_VSCROLL        As Long = &H115
Private Const WS_HSCROLL        As Long = &H100000
Private Const WS_VSCROLL        As Long = &H200000

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hWnd As Long, _
                                         ByVal wMsg As Long, _
                                         ByVal wParam As Long, _
                                               lParam As Any) As Long

Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long

'===========================================================================
' Form Events
'
Private Sub Form_Load()
    #If DEBUGMODE = 1 Then
        dbgFormName = Me.Name
    #End If
    Debug.Print vbCrLf + "=========================================" + _
                vbCrLf + " Started : " + Time$ + _
                vbCrLf + "-----------------------------------------"
    pInitTree
    Set moDragExpand = New XTimer
    Set moDragScroll = New XTimer

    '***********************************
    '** PROTOTYPING PURPOSES ONLY
    Dim bState As Boolean
    Dim RC As RECT

    With oTree
        bState = .Scroll
        .Scroll = False
        mlNodeHeight = .Height \ .GetVisibleCount
        .Scroll = bState
    
        mlDragExpandTime = 1000
        mlDragScrollTime = 200
        mbDragEnabled = True
    
        RC.Left = SendMessageAny(.hWnd, TVM_GETNEXTITEM, ByVal TVGN_ROOT, ByVal 0&)
        If SendMessageAny(.hWnd, TVM_GETITEMRECT, ByVal 1, RC) Then
            mlAutoScroll = (RC.Bottom - RC.Top) * 2
        Else
            mlAutoScroll = 32
        End If
    End With
    '***********************************

    mszDrag.cx = GetSystemMetrics(SM_CXDRAG)
    mszDrag.cy = GetSystemMetrics(SM_CYDRAG)

    moDragExpand.Interval = mlDragExpandTime
    moDragScroll.Interval = mlDragScrollTime

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Debug.Print vbCrLf + "-----------------------------------------" + _
                vbCrLf + " Finished : " + Time$ + _
                vbCrLf + "========================================="
    Set moDragExpand = Nothing
    Set moDragScroll = Nothing
End Sub

Private Sub moDragScroll_Tick()
'    #If DEBUGMODE = 1 Then
'        Debug.Print dbgFormName; "::moDragScroll -> Tick!"
        Debug.Print "Drag Scroll -> Tick!"
'    #End If

    Dim pt            As POINTAPI
    Dim rcClient      As RECT
    Dim dwRectFlags   As RectFlags
    Dim dwScrollFlags As ScrollDirectionFlags

    If mbInDrag = False Then
        moDragScroll.Enabled = False
        Exit Sub
    End If
    '
    '## Get the cursor postion in TreeView client coords
    '
    With oTree
        GetCursorPos pt
        ScreenToClient .hWnd, pt
        GetClientRect .hWnd, rcClient
    End With
    '
    '## If the cursor is within an auto scroll region in the TreeView's client area...
    '
    dwRectFlags = PtInRectRegion(rcClient, mlAutoScroll, pt)
    If dwRectFlags Then
        '
        '## Determine which direction the TreeView can be scrolled...
        '
        dwScrollFlags = IsWindowScrollable(oTree.hWnd)
        '
        '## If the cursor is within the respective drag region specified by the
        '   mlAutoScroll distance, and if the TreeView can be scrolled
        '   in that direction, send the TreeView that respective scroll message.
        '
        Select Case True
            Case (dwRectFlags And rfLeft) And (dwScrollFlags And sdLeft)
                'Debug.Print "Left"
                ScrollView [Line Left]
            Case (dwRectFlags And rfRight) And (dwScrollFlags And sdRight)
                'Debug.Print "Right"
                ScrollView [Line Right]
            Case (dwRectFlags And rfTop) And (dwScrollFlags And sdUp)
                'Debug.Print "Up"
                ScrollView [Up]
            Case (dwRectFlags And rfBottom) And (dwScrollFlags And sdDown)
                'Debug.Print "Down"
                ScrollView [Down]
            Case Else
                moDragScroll.Enabled = False
        End Select
    End If

End Sub

Private Sub moDragExpand_Tick()
'    #If DEBUGMODE = 1 Then
'        Debug.Print dbgFormName; "::moDragExpand -> Tick!"
        Debug.Print "Drag Expand -> Tick!"
'    #End If
    With oTree
        Select Case True
            Case (.DropHighlight Is Nothing), (moDragNode Is Nothing)
                '## Avoid possible error - should not be here! But it does happen.
                '!! Zhu, Exit Sub can't go here as we need to disable the timer
                '   first. More effecient code is to do nothing here.
            Case (.DropHighlight.Children > 0) And (.DropHighlight.Expanded = False)
                .DropHighlight.Expanded = True
        End Select
    End With
    moDragExpand.Enabled = False
End Sub

'===========================================================================
' otree Events
'
Private Sub otree_DragDrop(Source As Control, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgFormName; "::DragDrop -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            .DropHighlight = .HitTest(x, y)
            If Not (.DropHighlight Is Nothing) Then                 '## Did we drop a node?
                If moDragNode <> .DropHighlight Then                '## Yes. Did we drag the node onto itself?
                    'Debug.Print "Node " + moDragNode + " dropped on otree:" + .DropHighlight.Text
                    'RaiseEvent Dropped(moDragNode, .DropHighlight)  '## Notify programmer & Reset
                    Dropped moDragNode, .DropHighlight
                End If
            End If
            '## Reset
            Set .DropHighlight = Nothing
            Set moDragNode = Nothing
            mbInDrag = False
            mbStartDrag = False
            '.Drag vbEndDrag                    '!! Moved to otree_MouseUp
        End With
    End If

End Sub

Private Sub otree_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgFormName; "::DragOver -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            Set .DropHighlight = .HitTest(x, y)
            If .DropHighlight Is Nothing Then
                .DragIcon = LoadPicture(App.Path + "\no_m.CUR")
            Else
                .DragIcon = moDragNode.CreateDragImage
            End If
        End With
        pDoDrag
    End If

End Sub

Private Sub otree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgFormName; "::MouseDown -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    With oTree
        If mbDragEnabled Then                               '## Is drag'n'drop allowed?
            GetCursorPos mptBtnDown
            If Button = vbLeftButton Then
                Set moDragNode = .HitTest(x, y)             '## Capture the node to be dragged
            End If
        End If
    End With

End Sub

Private Sub otree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 2 Then
            Debug.Print dbgFormName; "::MouseMove -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    Dim pt As POINTAPI

    On Error GoTo ErrorHandler                              '@@ v01.00.03

    If mbDragEnabled Then                                   '## Is drag'n'drop allowed?
        If Button = vbLeftButton Then                       '## Yes. Signal a Drag operation.
            With oTree
                If Not (.HitTest(x, y) Is Nothing) Then     '## Do we have a node selected?
                    If mbStartDrag = True Then
                        mbInDrag = True                         '## Yes. Set the flag to true.
                        '.DragIcon = moDragNode.CreateDragImage '!! Moved to otree_DragOver
                        .Drag vbBeginDrag                       '## Signal VB to start drag operation.
                    Else
                        If Not (moDragNode Is Nothing) Then
                            'RaiseEvent StartDrag(moDragNode)    '## Notify programmer starting drag operation
                            GetCursorPos pt
                            If (Abs(pt.x - mptBtnDown.x) >= mszDrag.cx) Or (Abs(pt.y - mptBtnDown.y) >= mszDrag.cy) Then
                                StartDrag moDragNode
                                'Debug.Print "Start Drag with otree:" + moDragNode.Text
                                mbStartDrag = True
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If
    Exit Sub

ErrorHandler:                                               '@@ v01.00.03
    mbInDrag = False                                        '@@

End Sub

Private Sub otree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgFormName; "::MouseUp -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If
    mbStartDrag = False
    mbInDrag = False
    
    '********************************************
    '** NOTE: ADDED TO SEE IF FIXES PROBLEM!   **
    '********************************************
    oTree.Drag vbEndDrag
    '********************************************
    moDragExpand.Enabled = False
    moDragScroll.Enabled = False
End Sub

Private Sub otree_NodeClick(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgFormName; "::NodeClick"
    #End If
End Sub

'===========================================================================
' Drag Events
'
Private Sub StartDrag(DragNode As MSComctlLib.Node)
    Debug.Print "Start Dragging "; DragNode.Text
End Sub

Private Sub Dragging(DragNode As MSComctlLib.Node, DestNode As MSComctlLib.Node)
    Select Case True
        Case (moInDrag Is Nothing), moInDrag.Index <> DestNode.Index
            Set moInDrag = DestNode
            Debug.Print "Dragging"; DragNode.Text; " over "; DestNode.Text
    End Select
End Sub

Private Sub Dropped(DragNode As MSComctlLib.Node, DestNode As MSComctlLib.Node)
    Debug.Print "Dropped "; DragNode.Text; " on "; DestNode.Text
    If Not NodeMove(DestNode, DragNode) Then
        '
        '## Problems with moving the node. Most likely a root node was dragged!
        '
        MsgBox "Unable to move the selected node.", _
               vbApplicationModal + vbExclamation + vbOKOnly, _
               App.Title
    Else
        Debug.Print "-----------------------------------------"
    End If
End Sub

'===========================================================================
' Internal Functions
'

Private Function PtInRectRegion(RC As RECT, cxyRegion As Long, pt As POINTAPI) As RectFlags
    '
    '## Returns a set of bit flags indicating whether the specified point resides in
    '   the specified size region with the perimeter of the specified rect. cxyRegion
    '   defines the rectangular region within rc, and must be a positive value

    Dim dwFlags As RectFlags

    If PtInRect(RC, pt.x, pt.y) Then
        dwFlags = (rfLeft And (pt.x <= (RC.Left + cxyRegion)))
        dwFlags = dwFlags Or (rfRight And (pt.x >= (RC.Right - cxyRegion)))
        dwFlags = dwFlags Or (rfTop And (pt.y <= (RC.Top + cxyRegion)))
        dwFlags = dwFlags Or (rfBottom And (pt.y >= (RC.Bottom - cxyRegion)))
    End If

    PtInRectRegion = dwFlags

End Function

Private Function IsWindowScrollable(hWnd As Long) As ScrollDirectionFlags
    '
    '## Returns a set of bit flags indicating whether the specified
    '   window can be scrolled in any given direction.

    Dim si            As SCROLLINFO
    Dim dwScrollFlags As ScrollDirectionFlags

    si.cbSize = Len(si)
    si.fMask = SIF_ALL
    '
    '## Get the horizontal scrollbar's info (GetScrollInfo returns
    '   TRUE after a scrollbar has been added to a window,
    '   even if the respective style bit is not set...)
    '
    If (GetWindowLong(hWnd, GWL_STYLE) And WS_HSCROLL) Then
        If GetScrollInfo(hWnd, SB_HORZ, si) Then
            dwScrollFlags = (sdLeft And (si.nPos > 0))
            dwScrollFlags = dwScrollFlags Or (sdRight And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
        End If
    End If
    '
    '## Get the vertical scrollbar's info.
    '
    If (GetWindowLong(hWnd, GWL_STYLE) And WS_VSCROLL) Then
        If GetScrollInfo(hWnd, SB_VERT, si) Then
            dwScrollFlags = dwScrollFlags Or (sdUp And (si.nPos > 0))
            dwScrollFlags = dwScrollFlags Or (sdDown And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
        End If
    End If

    IsWindowScrollable = dwScrollFlags

End Function

Private Sub pDoDrag()

    Dim pt         As POINTAPI
    Dim rcClient   As RECT
    Static lOldNdx As Long

    With oTree
        If mbStartDrag = True Then
            If mbInDrag = True Then
                '
                '## If the cursor is still over same item as it was on the previous call,
                '   the cursor is over button, label, or icon of a collapsed parent item,
                '   start the auto expand timer, disable the timer otherwise.
                If Not (.DropHighlight Is Nothing) Then
                    If lOldNdx <> .DropHighlight.Index Then
                        If (.DropHighlight.Children > 0) And (.DropHighlight.Expanded = False) Then
                            moDragExpand.Enabled = True
                        Else
                            moDragExpand.Enabled = False
                        End If
                    End If
                    lOldNdx = .DropHighlight.Index
                End If
                '
                '## If the window is scrollable, and the cursor is within that auto scroll
                '   distance, start the auto scroll timer, disable the timer otherwise.
                '
                GetCursorPos pt
                ScreenToClient .hWnd, pt
                GetClientRect .hWnd, rcClient
                If (IsWindowScrollable(.hWnd) And PtInRectRegion(rcClient, mlAutoScroll, pt)) Then
                    moDragScroll.Enabled = True
                Else
                    moDragScroll.Enabled = False
                End If
                If Not (.DropHighlight Is Nothing) Then
                    '## We're over a node
                    'Debug.Print "Node " + moDragNode.Text + " dragging over otree:" + .DropHighlight
                    'RaiseEvent Dragging(moDragNode, .DropHighlight)
                    Dragging moDragNode, .DropHighlight
                End If
            End If
        End If
    End With

End Sub

Private Sub pInitTree()

    With oTree
        .Style = tvwTreelinesPlusMinusPictureText
        .LineStyle = tvwRootLines
        .Indentation = 10
        .ImageList = ilDialog
        '.FullRowSelect = True
        .HideSelection = False
        .HotTracking = True
        With .Nodes
            Dim lLoop As Long
            For lLoop = 0 To 25
                .Add , , Chr$(65 + lLoop), "Node " + Chr$(65 + lLoop), 1, 2
            Next
'            .Add , , "B", "Node B", 1, 2
'            .Add , , "C", "Node C", 1, 2
'            .Add , , "D", "Node D", 1, 2
'            .Add , , "E", "Node E", 1, 2

            Dim oNode As MSComctlLib.Node
            Set oNode = .Add(, , "X1", "Node Item 1", 1, 2)
            oNode.Expanded = True
            For lLoop = 2 To 20
                Set oNode = .Add(oTree.Nodes("X" + CStr(lLoop - 1)), _
                                 tvwChild, _
                                 "X" + CStr(lLoop), _
                                 "Node Item " + CStr(lLoop), 1, 2)
                oNode.Expanded = True
            Next
        End With
    End With

End Sub

Private Function NodeMove(ParentNode As MSComctlLib.Node, _
                          ChildNode As MSComctlLib.Node, _
           Optional ByVal bSelect As Boolean = True) As Boolean

    Dim lNDX   As Long
    Dim lCount As Long
    Dim lLoop  As Long
    Dim bRoot  As Boolean

    With ChildNode
        If ParentNode = ChildNode Then
            '## Same node - therefore no point
            Exit Function
        End If
        If IsParentNode(ParentNode, ChildNode) Then '## Are we moving a parent node?
            If IsRootNode(ChildNode) Then           '## Yes. Is it a root node?
                Exit Function                       '## Yes. Can't move a root node.
            End If
            '## move the children before moving the designated node
            lCount = .Children
            For lLoop = 1 To lCount
                lNDX = .Child.Index
                Set oTree.Nodes(lNDX).Parent = .Parent
            Next
        End If
        '## Force the ParentNode to be expanded before the move
        ParentNode.Expanded = True                  '@@ v01.00.03
        '## Give the child a new parent
        Set .Parent = ParentNode
        If bSelect Then
            .EnsureVisible
            .Selected = bSelect
        End If
    End With
    NodeMove = True

End Function

Private Function IsParentNode(ChildNode As MSComctlLib.Node, _
                              ParentNode As MSComctlLib.Node) As Boolean
    '## Checks if one node is the parent of another.
    '   This is a recursive routine that steps down through
    '   the branches of the parent node.

    Dim lNDX As Long

    If ParentNode.Children Then             '## Does the parent node have children?
        lNDX = ParentNode.Child.Index       '## Yes, remember the first child
        Do                                  '## Step through all child nodes
            If lNDX = ChildNode.Index Then  '## is ChildNode the test node?
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If IsParentNode(ChildNode, oTree.Nodes(lNDX)) Then  '## Step down through the branches
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If lNDX <> ParentNode.Child.LastSibling.Index Then  '## Have we tested the last child node?
                lNDX = oTree.Nodes(lNDX).Next.Index         '## No. Point to the next child node
            Else
                Exit Do                                         '## Yes.
            End If
        Loop
    End If

End Function

Private Function IsRootNode(Node As MSComctlLib.Node) As Boolean
    '## Check is selected node is a root node.
    With Node
        IsRootNode = (.FullPath = .Text)
    End With
End Function

Private Sub ScrollView(ByVal Dir As eCodeScrollView)
    '
    '## Scrolls the treview using code
    '
    Dim lHwnd As Long
    Dim lPos  As Long
    Dim lBar1 As Long
    Dim lBar2 As Long
    Dim lDir  As Long

    lHwnd = oTree.hWnd
    Select Case Dir
        Case [Home]:       SendMessageAny lHwnd, WM_VSCROLL, SB_TOP, vbNull
        Case [Page Up]:    SendMessageAny lHwnd, WM_VSCROLL, SB_PAGEUP, vbNull
        Case [Up]:         SendMessageAny lHwnd, WM_VSCROLL, SB_LINEUP, vbNull
        Case [Down]:       SendMessageAny lHwnd, WM_VSCROLL, SB_LINEDOWN, vbNull
        Case [Page Down]:  SendMessageAny lHwnd, WM_VSCROLL, SB_PAGEDOWN, vbNull
        Case [End]:        SendMessageAny lHwnd, WM_VSCROLL, SB_BOTTOM, vbNull
        Case [Left]:       SendMessageAny lHwnd, WM_HSCROLL, SB_LEFT, vbNull
        Case [Page Left]:  SendMessageAny lHwnd, WM_HSCROLL, SB_PAGELEFT, vbNull
        Case [Line Left]:  SendMessageAny lHwnd, WM_HSCROLL, SB_LINELEFT, vbNull
        Case [Line Right]: SendMessageAny lHwnd, WM_HSCROLL, SB_LINERIGHT, vbNull
        Case [Page Right]: SendMessageAny lHwnd, WM_HSCROLL, SB_PAGERIGHT, vbNull
        Case [Right] ':    SendMessageAny lHwnd, WM_HSCROLL, SB_RIGHT, vbNull
            '
            '## For some reason, the treeview doesn't respond to the above commented
            '   out message. Therefore a work-around is required.
            '
            '## To stop flickering, the control is hidden temporarily.
            oTree.Visible = False
            ' ## Loop until we've scrolled to the far right side
            Do
                lPos = GetScrollPos(lHwnd, 0&)
                SendMessageAny lHwnd, WM_HSCROLL, SB_PAGERIGHT, vbNull
            Loop Until (lPos = GetScrollPos(lHwnd, 0&))
            '## Now show the control
            oTree.Visible = True
    End Select

End Sub
