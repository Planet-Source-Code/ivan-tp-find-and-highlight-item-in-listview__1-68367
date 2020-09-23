VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "List View"
   ClientHeight    =   5100
   ClientLeft      =   6240
   ClientTop       =   1935
   ClientWidth     =   6585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6585
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "Fill the [Name] to search"
      Top             =   4200
      Width           =   5655
   End
   Begin VB.TextBox txtNIK 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "Fill the [NIK] to search"
      Top             =   3720
      Width           =   5655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NIK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   6068
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************** Finding Item and Highlight it *******'
'Listview properties : Hide Selection dan multiselect ( Uncheck)

'Based On : D K Richmond Media Library Sample
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60524&lngWId=1
'This my first submission when i see the DK Richmond submission,
'give me idea to develop this code.
'please do not laugh as I am a beginner in this :D


Private Sub Form_Load()
    
    Dim i As Integer
    'Add Item to ListView
    ListView1.ListItems.Add , , "001"
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "David K Richmond"
    ListView1.ListItems.Add , , "002"
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Heriberto MS"
    ListView1.ListItems.Add , , "003"
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Dream VB/Ben Jones"
    ListView1.ListItems.Add , , "004"
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Richard Mewett"
    ListView1.ListItems.Add , , "005"
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Juan Carlos SR"
    ListView1.ListItems.Add , , "006"
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Option Explicit"
    ListView1.ListItems.Add , , "007"
    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Lavolpe"
    For i = 1 To 30
        ListView1.ListItems.Add , , "001" & i
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "PSC" & i
    Next i
    
End Sub



Private Sub txtName_Change()
    
    Dim ListVwItem As MSComctlLib.ListItem 'common control Library
    Dim Value As String
    Dim length As Integer
    Dim i As Integer
    
    Value = txtName.Text
    
    For Each ListVwItem In ListView1.ListItems
        length = Len(ListVwItem.SubItems(1))
        For i = 1 To length
            If LCase(Mid(ListVwItem.SubItems(1), 1, i)) = LCase(Value) Then 'find one per one character
                ListVwItem.Selected = True
                ListVwItem.EnsureVisible
            End If
        Next i
    Next
    
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    Dim ListVwItem As MSComctlLib.ListItem
    Dim Value As String
    
    If KeyAscii = 13 Then ' enter key
            
    
        Value = txtName.Text
    
        For Each ListVwItem In ListView1.ListItems
                If LCase(ListVwItem.SubItems(1)) = LCase(Value) Then 'find complete words
                    ListVwItem.Selected = True
                    ListVwItem.EnsureVisible
                    Exit For
                End If
                
        Next
        ListView1.SetFocus
    End If
End Sub

Private Sub txtNIK_Change()
    
    Dim ListVwItem As MSComctlLib.ListItem
    Dim Value As String
    
    Value = txtNIK.Text
    For Each ListVwItem In ListView1.ListItems
        If ListVwItem.Text = Value Then 'find complete words
            ListVwItem.Selected = True
            ListVwItem.EnsureVisible
            Exit For
        End If
    Next
    ListView1.SetFocus
End Sub

