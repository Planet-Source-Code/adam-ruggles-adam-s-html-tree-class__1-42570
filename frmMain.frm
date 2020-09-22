VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInit 
      Caption         =   "Initialize"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox 
      Height          =   5415
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9551
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      FileName        =   "C:\Documents and Settings\Adam\Desktop\test.txt"
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "Format"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdPopTree 
      Caption         =   "Populate Tree"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdPopList 
      Caption         =   "Populate List"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin MSComctlLib.ListView lview 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Error Report"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.TreeView tview 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7011
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   178
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private WithEvents MyTree As clsTree
Attribute MyTree.VB_VarHelpID = -1
Option Explicit


Private Sub cmdFormat_Click()
  lview.ListItems.Clear
  RichTextBox.Text = MyTree.FormatHTML()
  'Now I need to rebuild the Tree so when you click on the tags in the tree
  'It goes to the right place in the richtextbox
  MyTree.InitializeHTML RichTextBox.Text
  cmdPopTree_Click
End Sub

Private Sub cmdInit_Click()
  'Clears the listview so we can put errors in there if any are generated
  lview.ListItems.Clear
  'Rebuild the data tree
  MyTree.InitializeHTML RichTextBox.Text
End Sub

Private Sub cmdPopList_Click()
  'This sub just produces a raw listing of what was parsed from the HTML
  'I reused the Error ListView for this.
  MyTree.ProduceList lview
End Sub


Private Sub cmdPopTree_Click()
  'This displays the info from the current data tree into the treeview
  'If you've made a change to the text after initializing make sure
  'You rebuild the tree by running the InitializeHTML
  Dim lCnt As Long
  LockWindowUpdate tview.hWnd
  frmMain.MousePointer = vbHourglass
  MyTree.ProduceTree tview
  For lCnt = 1 To tview.Nodes.Count
    tview.Nodes(lCnt).Expanded = True
  Next lCnt
  tview.SelectedItem = tview.Nodes(1)
  frmMain.MousePointer = vbNormal
  LockWindowUpdate False
End Sub

Private Sub Form_Load()
  RichTextBox.RightMargin = 99999
  Set MyTree = New clsTree
  MyTree.InitializeHTML RichTextBox.Text
  cmdPopTree_Click
End Sub

Private Sub lview_ItemClick(ByVal Item As MSComctlLib.ListItem)
  On Error Resume Next
  'Shows you how to navigate to places in the RichTextbox and tree from the tree data
  tview.SelectedItem = tview.Nodes(lview.SelectedItem.Key)
  tview.SelectedItem.EnsureVisible
  RichTextBox.SelStart = MyTree.ReturnPos(Val(lview.SelectedItem.Key)) - 1
  RichTextBox.SelLength = MyTree.ReturnLen(Val(lview.SelectedItem.Key))
  RichTextBox.SetFocus
End Sub

Private Sub MyTree_TreeError(Pos As Long, Text As String, Tag As String, ID As Long)
  'Get a listing of the Errors produced in the Linking Process
  lview.ListItems.Add , ID & "ID", Text & " (tag=" & Tag & ")"
End Sub

Private Sub tview_NodeClick(ByVal Node As MSComctlLib.Node)
  'Shows how to go to a place in the richtextbox by clicking on a node in the treeview
  RichTextBox.SelStart = MyTree.ReturnPos(Val(Node.Key)) - 1
  RichTextBox.SelLength = MyTree.ReturnLen(Val(Node.Key))
  RichTextBox.SetFocus
End Sub
