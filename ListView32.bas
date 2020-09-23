Attribute VB_Name = "ListView32"
Option Explicit

'*****************************************************************
'*
'* ListView32.bas
'*
'* Version 0.01 on 07 Janauri 2002 by the author J. Klarenbeek
'* Usee
'* e-mail: johan@easysoftwaresolutions.com
'*
'* You may and shall not remove any comments or remarks in this
'* file if not explicitly allowed by the author. You are allowed
'* to insert your own comments and or change the actual runtime
'* code as pleased.
'*
'* If you don't want to use the flex grid control to edit subitems
'* of a listview, this is for you!! When you include this module
'* in your project, add 1 listview and 1 textbox to your form.
'* Only a few lines of code enables you to create a 2 dimensional
'* editable listview.
'*
'* No subclassing and no other components or references required.
'*
'* The original code can be found on www.planet-source-code.com,
'* and is accompanied by 2 other files; 1 example form and 1
'* project file compressed in a .zip file.
'*
'*****************************************************************

' listview structures
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type LVHITTESTINFO
    pt As POINTAPI
    lFlags As Long
    lItem As Long
    lSubItem As Long
End Type

' send message to listview
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' styles for listview
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVM_FIRST = &H1000
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H37
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H36
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)

' hittest constants
Public Const LVHT_NOWHERE = &H1
Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8
Public Const LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

' edit subitem constants
Public Const LVIR_BOUNDS = 0
Public Const LVIR_ICON = 1
Public Const LVIR_LABEL = 2
Public Const LVIR_SELECTBOUNDS = 3

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Sub ListView_FullRowSelect(ByRef ListView As ListView)

    Dim lStyle As Long
    lStyle = SendMessage(ListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lStyle = lStyle Or LVS_EX_FULLROWSELECT
    Call SendMessage(ListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)

End Sub

Public Sub ListView_GridLines(ByRef ListView As ListView)
    
    Dim lStyle As Long
    lStyle = SendMessage(ListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lStyle = lStyle Or LVS_EX_GRIDLINES
    Call SendMessage(ListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)

End Sub

Public Function ListView_ReturnSelected(ListView As ListView) As ListItem

    Set ListView_ReturnSelected = Nothing
    
    Dim lvCount As Integer
    lvCount = ListView.ListItems.Count
    
    If lvCount = 0 Then Exit Function
        
    Dim lvIndex As Integer
    For lvIndex = 1 To lvCount
        If ListView.ListItems(lvIndex).Selected = True Then
            Set ListView_ReturnSelected = ListView.ListItems(lvIndex)
            Exit Function
        End If
    Next lvIndex
        
End Function

Public Function ListView_HitTest(ListView As ListView, x As Single, y As Single) As LVHITTESTINFO
    
    Dim lRet As Long
    Dim lX As Long
    Dim lY As Long

    'x and y are in twips; convert them to pixels for the API call
    lX = x / Screen.TwipsPerPixelX
    lY = y / Screen.TwipsPerPixelY

    Dim tHitTest As LVHITTESTINFO
    
    With tHitTest
        .lFlags = 0
        .lItem = 0
        .lSubItem = 0
        .pt.x = lX
        .pt.y = lY
    End With
    
    'return the filled Structure to the routine
    lRet = SendMessage(ListView.hwnd, LVM_SUBITEMHITTEST, 0, tHitTest)
    
    ListView_HitTest = tHitTest
    
End Function

Public Sub ListView_ScaleEdit(ListView As ListView, tHitTest As LVHITTESTINFO, TextBox As TextBox)

    If tHitTest.lItem = -1 Then
        TextBox.Visible = False
        Exit Sub
    End If
        
    Dim XPixels As Integer
    Dim YPixels As Integer
    
    XPixels = Screen.TwipsPerPixelX
    YPixels = Screen.TwipsPerPixelY
    
    Dim tRec As RECT
    tRec.Top = tHitTest.lSubItem
    tRec.Left = LVIR_LABEL
    tRec.Bottom = 0
    tRec.Right = 0

    Dim lRet As Long
    lRet = SendMessage(ListView.hwnd, LVM_GETSUBITEMRECT, tHitTest.lItem, tRec)

    Dim lvRect As RECT
    lRet = GetClientRect(ListView.hwnd, lvRect)
    lvRect.Bottom = lvRect.Bottom * YPixels
    lvRect.Right = lvRect.Right * XPixels
    lvRect.Top = Round((ListView.Width - lvRect.Right) / 2)
    lvRect.Left = Round((ListView.Height - lvRect.Bottom) / 2)
    
    TextBox.Top = ListView.Top + lvRect.Top + tRec.Top * YPixels
    TextBox.Left = ListView.Left + lvRect.Left + tRec.Left * XPixels
    TextBox.Width = (tRec.Right - tRec.Left) * XPixels
    TextBox.Height = (tRec.Bottom - tRec.Top) * YPixels

End Sub

Public Sub ListView_BeforeEdit(ListView As ListView, tHitTest As LVHITTESTINFO, TextBox As TextBox)

    If tHitTest.lItem = -1 Then
        Exit Sub
    End If

    If tHitTest.lSubItem = 0 Then
        TextBox.Text = ListView.ListItems(tHitTest.lItem + 1).Text
    Else
        TextBox.Text = ListView.ListItems(tHitTest.lItem + 1).SubItems(tHitTest.lSubItem)
    End If
    
    TextBox.Visible = True
    TextBox.SetFocus
    If Len(TextBox.Text) > 0 Then
        TextBox.SelStart = 0
        TextBox.SelLength = Len(TextBox.Text)
    End If
    
    
End Sub

Public Sub ListView_AfterEdit(ListView As ListView, tHitTest As LVHITTESTINFO, TextBox As TextBox)

    Dim bEditMode As Boolean
    bEditMode = False
    
    If tHitTest.lItem > -1 Then
        If TextBox.Visible = True Then
            bEditMode = True
        End If
    End If
       
    TextBox.Visible = False
    
    If bEditMode = True Then
    
        If tHitTest.lSubItem = 0 Then
            ListView.ListItems(tHitTest.lItem + 1).Text = TextBox.Text
        Else
            ListView.ListItems(tHitTest.lItem + 1).SubItems(tHitTest.lSubItem) = TextBox.Text
        End If
        
        tHitTest.lSubItem = (tHitTest.lSubItem + 1) Mod ListView.ColumnHeaders.Count
        If tHitTest.lSubItem = 0 Then
            tHitTest.lItem = (tHitTest.lItem + 1) Mod ListView.ListItems.Count
        End If
    End If
        
End Sub

