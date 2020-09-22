VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
' HTML Data Tree Class v1.00
' Copyright ©2002 by Adam Ruggles, All Rights Reserved.
'*********************************************************************
' If you like this code, please vote for it at Planet-Source-Code.com
' Thank you
'*********************************************************************
' You are free to use this code within your own applications, but you
' are expressly forbidden from selling or otherwise distributing the
' source code without prior written consent.
'*********************************************************************
' This Class Uses the cAppendString for quicker string concatenation,
' Please see that code for any Credits and Copyrights it may have
'*********************************************************************

Public Enum NodeType
  ntText = 1
  ntTag = 2
  ntEndTag = 3
End Enum

Private Type TreeNode
  ID As String
  ParentID As String
  TagName As String
  Text As String
  Child() As Long
  Pos As Long
  NType As NodeType
  Linked As Boolean
End Type

Private AppendStr As cAppendString
Private LinkedList() As TreeNode
Private RootList() As Long
Private SavePos As Long
Private SaveID As Long
Private bInA As Boolean
Private bUsedCrLf As Boolean
'These three Variables are for the MakeLink Function
'There where declared here so they would not be
'Reproduced in the recursion
Private LinkChild() As String
Private MatchTag As String
Private LinkCnt As Long
Option Explicit
Event TreeError(Pos As Long, Text As String, Tag As String, ID As Long)

Public Sub InitializeHTML(HTML As String)
  'This makes the data tree for use in all of the functions
  'This needs to be called first before using any of the
  'other procedures or functions
  ReDim LinkedList(0)
  'Produces a list of tag, end tags, and text from the HTML string
  Parse HTML
  'Links the list into a Multiple Trees
  'and puts the Roots into the RootLists
  DoLink
End Sub

Public Sub ProduceList(ByRef inListView As ListView, Optional HTML As String = "", _
  Optional CallInit As Boolean = False)
'*** Produces a simple list from the HTML string and puts
'*** it into a ListView Control
  Dim lCnt As Long
  Dim lUBound As Long
  inListView.ListItems.Clear
  If CallInit = True Then InitializeHTML HTML
  lUBound = UBound(LinkedList)
  For lCnt = 1 To lUBound
    With LinkedList(lCnt)
      inListView.ListItems.Add , , .Text
    End With
  Next lCnt
End Sub
'****************************************************************
'*** The next two functions add the HTML tree into a TreeView ***
'****************************************************************
Public Sub ProduceTree(ByRef inTreeView As TreeView, Optional HTML As String = "", _
  Optional CallInit As Boolean = False)
'*** Produces a list from the HTML string then links
'*** the list into a tree structure then puts it into the TreeView
  On Error Resume Next
  Dim lCnt, lCnt2 As Long
  Dim lUBound As Long
  inTreeView.Nodes.Clear
  If CallInit = True Then InitializeHTML HTML
  For lCnt = 1 To UBound(RootList)
    With LinkedList(RootList(lCnt))
      inTreeView.Nodes.Add , , .ID, "<" & .TagName & ">"
      If .Child(0) <> -1 Then
        'A recusive procedure that walks the tree
        'and adds the nodes to the TreeView
        AddRecTreeView RootList(lCnt), .ID, inTreeView
      End If
    End With
  Next lCnt
End Sub

Private Sub AddRecTreeView(CurIndex As Long, CurID As String, ByRef inTreeView As TreeView)
  'Recursivly Walks the tree adding each element of the
  'Tree Data structure into the TreeView Control
  Dim lCnt As Long
  For lCnt = UBound(LinkedList(CurIndex).Child) To 0 Step -1
    With LinkedList(LinkedList(CurIndex).Child(lCnt))
      If .NType = ntTag Then
        inTreeView.Nodes.Add CurID, tvwChild, .ID, "<" & .TagName & ">"
        If .Child(0) <> -1 Then
          AddRecTreeView LinkedList(CurIndex).Child(lCnt), .ID, inTreeView
        End If
      ElseIf .NType = ntEndTag Then
        inTreeView.Nodes.Add CurID, tvwChild, .ID, "</" & .TagName & ">"
      End If
    End With
  Next lCnt
End Sub

'****************************************************************
'*** The next two functions re-format the HTML String         ***
'****************************************************************
Public Function FormatHTML(Optional HTML As String = "", Optional CallInit As Boolean = False) As String
'*** Takes an HTML string and formats to its
'*** easier to read
  Dim lCnt As Long
  Set AppendStr = New cAppendString
  If CallInit = True Then InitializeHTML HTML
  For lCnt = 1 To UBound(RootList)
    With LinkedList(RootList(lCnt))
      'I used the Append Class so it would put the
      'new HTML string faster
      AppendStr.Append .Text & vbCrLf
      bUsedCrLf = True
      If .Child(0) <> -1 Then
        'Call the RecFormat function to walk the tree
        'And add each piece to the new HTML
        RecFormat RootList(lCnt), 1
      End If
    End With
  Next lCnt
  'Return the full new HTML string
  FormatHTML = AppendStr.Value
  'Clear the AppendStr class
  AppendStr.Clear
End Function

Private Sub RecFormat(CurIndex As Long, Lvl As Long)
  'Recursivly walks the tree putting together the
  'New HTML string
  Dim lCnt As Long
  For lCnt = UBound(LinkedList(CurIndex).Child) To 0 Step -1
    With LinkedList(LinkedList(CurIndex).Child(lCnt))
      If .NType = ntEndTag Then
        'if its an end Tag we want the LVL to be the same
        'as the begin tag
        If bInA And .TagName = "A" Then
          AppendStr.Append .Text
          bUsedCrLf = False
          bInA = False
        ElseIf bInA Then
          AppendStr.Append .Text
          bUsedCrLf = False
        Else
          If bUsedCrLf = True Then
            AppendStr.Append String((Lvl - 1) * 2, Chr(32)) & .Text & vbCrLf
          Else
            AppendStr.Append .Text & vbCrLf
          End If
          bUsedCrLf = True
        End If
      Else
        If .TagName = "A" Then bInA = True
        If bInA And .TagName = "A" Then
          If bUsedCrLf = True Then
            AppendStr.Append String(Lvl * 2, Chr(32)) & .Text
          Else
            AppendStr.Append .Text
          End If
          bUsedCrLf = False
        ElseIf bInA Then
          AppendStr.Append .Text
          bUsedCrLf = False
        Else
          If bUsedCrLf = True Then
            AppendStr.Append String(Lvl * 2, Chr(32)) & .Text & vbCrLf
          Else
            If .NType = ntTag Then
              AppendStr.Append vbCrLf & String(Lvl * 2, Chr(32)) & .Text & vbCrLf
            Else
              AppendStr.Append .Text & vbCrLf
            End If
          End If
          bUsedCrLf = True
        End If
      End If
      If .Child(0) <> -1 Then
        RecFormat LinkedList(CurIndex).Child(lCnt), Lvl + 1
      End If
    End With
  Next lCnt
End Sub
'****************************************************************
'*** The next three functions output data from the data tree  ***
'****************************************************************
Public Function ReturnPos(ListIndex As Long) As Long
  ReturnPos = LinkedList(ListIndex).Pos
End Function
Public Function ReturnLen(ListIndex As Long) As Long
  ReturnLen = Len(LinkedList(ListIndex).Text)
End Function
Public Function ReturnType(ListIndex As Long) As NodeType
  ReturnType = LinkedList(ListIndex).NType
End Function


Private Sub Parse(HTML As String)
  '*** This produces a List but doesn't link the list into a tree
  Dim lBegin, lEnd, lUBound, lLenOfStr, lCnt As Long
  Dim sTemp, sStrip() As String
  lLenOfStr = Len(HTML)
  lBegin = 1
  lEnd = 1
  'Go through each "<" & ">" to parse the file into pieces
  Do
    lBegin = InStr(lEnd, HTML, "<", vbBinaryCompare)
    If lBegin = 0 Then Exit Do
    '*** Everything here is not a Tag (ie. Not "<" * ">")
    If lBegin <> lEnd Then
      sTemp = Replace(Trim(Mid$(HTML, lEnd + 1, lBegin - lEnd - 1)), _
        vbTab, "", , , vbBinaryCompare)
      
      If Replace(sTemp, vbCrLf, "", , , vbBinaryCompare) <> "" And _
         Replace(sTemp, vbLf, "", , , vbBinaryCompare) <> "" Then
        'This breaks up each piece of the text by crlf
        sStrip = Split(sTemp, vbLf, , vbBinaryCompare)
        For lCnt = 0 To UBound(sStrip)
          If Replace(sStrip(lCnt), vbCr, "") <> "" Then
            lUBound = UBound(LinkedList) + 1
            ReDim Preserve LinkedList(lUBound)
            With LinkedList(lUBound)
              ReDim .Child(0)
              .Child(0) = -1
              .Linked = False
              .Pos = 0
              .Text = Trim(Replace(sStrip(lCnt), vbCr, ""))
              .NType = ntText
              .ID = lUBound & "ID"
             End With
          End If
        Next lCnt
      End If
    End If
    lEnd = InStr(lBegin, HTML, ">", vbBinaryCompare)
    If lEnd = 0 Then Exit Do
    '*** Everything here is a Tag (ie. "<" * ">")
    If lBegin <> lEnd Then
      sTemp = Trim(Mid$(HTML, lBegin, lEnd - lBegin + 1))
      lUBound = UBound(LinkedList) + 1
      ReDim Preserve LinkedList(lUBound)
      With LinkedList(lUBound)
        ReDim .Child(0)
        .ID = lUBound & "ID"
        .Child(0) = -1
        .Pos = lBegin
        .Text = sTemp
        .Linked = False
        'Now we want to categorize each tag
        If Mid$(sTemp, 2, 1) = "/" Then
          .NType = ntEndTag
          .TagName = UCase(Mid$(sTemp, 3, Len(sTemp) - 3))
        ElseIf Mid$(sTemp, 2, 1) = "!" Then
          .TagName = UCase(Mid$(sTemp, 2, Len(sTemp) - 2))
          .NType = ntTag
        Else
          .NType = ntTag
          .TagName = Mid$(sTemp, 2, Len(sTemp) - 2)
          sStrip = Split(.TagName, " ", , vbBinaryCompare)
          .TagName = UCase(sStrip(0))
        End If
      End With
      'Check For a Script and Handle it differently
      If (LinkedList(lUBound).TagName = "SCRIPT" Or _
         LinkedList(lUBound).TagName = "STYLE") _
         And LinkedList(lUBound).NType = ntTag Then
        'I used lCnt as a temp variable in the following lines
        lCnt = lEnd
        If LinkedList(lUBound).TagName = "STYLE" Then
          lEnd = InStr(lEnd + 1, HTML, "</style>", vbTextCompare)
        Else
          lEnd = InStr(lEnd + 1, HTML, "</script>", vbTextCompare)
        End If
        sTemp = Trim(Mid$(HTML, lCnt + 1, lEnd - lCnt - 1))
        'I assigned the script the same as any text except it is not
        'parsed into smaller text fields on the tree.  This is for
        'the format function. We don't want to mess with the script
        'formatting
        sStrip = Split(sTemp, vbLf, , vbBinaryCompare)
        For lCnt = 0 To UBound(sStrip)
          If Replace(sStrip(lCnt), vbCr, "") <> "" Then
            lUBound = UBound(LinkedList) + 1
            ReDim Preserve LinkedList(lUBound)
            With LinkedList(lUBound)
              ReDim .Child(0)
              .Child(0) = -1
              .Linked = False
              .Pos = 0
              .Text = Trim(Replace(sStrip(lCnt), vbCr, ""))
              .NType = ntText
              .ID = "ID" & lUBound
             End With
          End If
        Next lCnt
      End If
    End If
  Loop Until lEnd > lLenOfStr Or lEnd = 0
End Sub

Private Sub DoLink()
  'This link fills the Child Property
  'so the tree after this function is created
  ReDim RootList(0)
  Dim lCnt As Long
  Dim lCnt2 As Long
  For lCnt = 1 To UBound(LinkedList)
    With LinkedList(lCnt)
      'First we will search for an end tag
      If .NType = ntEndTag Then
        MatchTag = .TagName
        ReDim LinkChild(0)
        LinkChild(0) = lCnt
        SavePos = .Pos
        SaveID = lCnt
        'Now we recusivly move up the tree to find
        'The parent
        .ParentID = MakeLink(lCnt - 1)
        If .ParentID <> "" Then .Linked = True
      End If
    End With
  Next lCnt
  'Now we gather all of the nodes with out parents
  'and put a pointer to them in the RootList array
  For lCnt = 1 To UBound(LinkedList)
    With LinkedList(lCnt)
      If .Linked = False Then
        lCnt2 = UBound(RootList) + 1
        ReDim Preserve RootList(lCnt2)
        RootList(lCnt2) = lCnt
      End If
    End With
  Next lCnt
End Sub

Private Function MakeLink(CurIndex As Long) As String
  'This function makes the links
  If CurIndex = 0 Then
    'If we've hit the top of the tree then we couldn't
    'find the begin tag for an end tag so we generate
    'an error
    RaiseEvent TreeError(SavePos, "End tag without matching begin tag", MatchTag, SaveID)
    MakeLink = ""
    Exit Function
  End If
  With LinkedList(CurIndex)
    'We are looking for the top tag
    If .TagName = MatchTag And .Child(0) = -1 And .NType = ntTag Then
      'If we have found the top tag then we link it's Children
      ReDim .Child(UBound(LinkChild))
      'This can be done faster but it works
      'We are copying one array to another
      For LinkCnt = 0 To UBound(LinkChild)
        .Child(LinkCnt) = LinkChild(LinkCnt)
      Next LinkCnt
      'Now return the top tag ID
      MakeLink = .ID
    Else
      If .Linked = False Then
        'we havn't found the top tag and we have no parent yet (unlinked)
        LinkCnt = UBound(LinkChild) + 1
        'Now add this tag to the child tag array
        ReDim Preserve LinkChild(LinkCnt)
        LinkChild(LinkCnt) = CurIndex
        'Call MakeLink again and find the parent
        .ParentID = MakeLink(CurIndex - 1)
        'Return the top tag ID
        MakeLink = .ParentID
        If .ParentID <> "" Then .Linked = True
      Else
        'If we have already been linked we still are looking
        'for the top tag so we call MakeLink again
        MakeLink = MakeLink(CurIndex - 1)
      End If
    End If
  End With
End Function

Private Sub Class_Initialize()
  ReDim LinkedList(0)
  ReDim RootList(0)
  ReDim LinkChild(0)
  MatchTag = ""
  LinkCnt = 0
End Sub
