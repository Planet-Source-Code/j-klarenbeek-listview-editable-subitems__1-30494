VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "ListView Edit SubItem Example"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tbListViewEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin ComctlLib.ListView lvVSS 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu lvVSSMenu 
      Caption         =   "VSSMENU"
      Visible         =   0   'False
      Begin VB.Menu lvVSSMenuEdit 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu lvVSSMenuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu lvVSSMenuAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu lvVSSMenuRemove 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*****************************************************************
'*
'* Form1.frm
'*
'* Version 0.01 on 07 Janauri 2002 by the author J. Klarenbeek
'*
'* This form demonstrates howto use the accompanied file:
'* ListView32.bas
'*
'* You may and shall not remove any comments or remarks in this
'* file if not explicitly allowed by the author. You are allowed
'* to insert your own comments and or change the actual runtime
'* code as pleased.
'*
'*****************************************************************


Dim tHt As LVHITTESTINFO

Private Sub Form_Load()

    Randomize
    
    'Initialize listview
    tHt.lItem = -1
    
    ' set lvVSS to set nodes for project.
    Call ListView_FullRowSelect(lvVSS)
    Call ListView_GridLines(lvVSS)
    lvVSS.ColumnHeaders.Add , , "VSS Node"
    lvVSS.ColumnHeaders.Add , , "Local Directory"
        
End Sub

Private Sub lvVSS_DblClick()
    
    Call lvVSSMenuEdit_Click
    
End Sub

Private Sub lvVSS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
    Call ListView_AfterEdit(lvVSS, tHt, tbListViewEdit)
      
    tHt = ListView_HitTest(lvVSS, x, y)
        
    If Button <> 2 Then Exit Sub
    
    If tHt.lItem = -1 Then
        lvVSSMenuRemove.Enabled = False
        lvVSSMenuEdit.Enabled = False
    Else
        lvVSSMenuRemove.Enabled = True
        lvVSSMenuEdit.Enabled = True
        lvVSS.ListItems(tHt.lItem + 1).Selected = True
    End If
    
    PopupMenu lvVSSMenu
    
End Sub

Private Sub lvVSSMenuAdd_Click()

    Dim lvVSSItem As ListItem
    Set lvVSSItem = lvVSS.ListItems.Add(, , "New " & Rnd())
    lvVSSItem.SubItems(1) = "Dir " & Rnd()
    lvVSSItem.Selected = True
    
End Sub

Private Sub lvVSSMenuEdit_Click()

        Call ListView_ScaleEdit(lvVSS, tHt, tbListViewEdit)
        
        Call ListView_BeforeEdit(lvVSS, tHt, tbListViewEdit)

End Sub

Private Sub lvVSSMenuRemove_Click()
    
    Dim lvVSSItem As ListItem

    Set lvVSSItem = ListView_ReturnSelected(lvVSS)
    
    If lvVSSItem Is Nothing Then Exit Sub
    
    lvVSS.ListItems.Remove lvVSSItem.Index
    
End Sub

Private Sub tbListViewEdit_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case vbKeyEscape
        tbListViewEdit.Visible = False
    Case vbKeyReturn
        Call ListView_AfterEdit(lvVSS, tHt, tbListViewEdit)
    End Select
    
End Sub

Private Sub tbListViewEdit_LostFocus()

    Dim bNextItem As Boolean
    bNextItem = False
    
    If tbListViewEdit.Visible = True Then
        bNextItem = True
    End If
    
    Call ListView_AfterEdit(lvVSS, tHt, tbListViewEdit)
        
    If bNextItem = True Then
            lvVSS.ListItems(tHt.lItem + 1).Selected = True
            Call lvVSSMenuEdit_Click
    End If
    
End Sub
