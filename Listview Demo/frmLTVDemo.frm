VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLTVDemo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "grumtek - Listview Demo"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetText 
      Caption         =   "Get &Text Of Selected"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdRemoveChecked 
      Caption         =   "R&emove All Checked"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdUncheckAll 
      Caption         =   "&Uncheck All"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "C&heck All"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Index"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdIndex1 
      Caption         =   "+ &1 To Index"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdIndex 
      Caption         =   "Find Listview &Index"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Listview"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDemoIt 
      Caption         =   "&Demo It!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   -2147483647
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Demo View 1"
         Object.Width           =   3476
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Demo View 2"
         Object.Width           =   3476
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Demo View 3"
         Object.Width           =   3476
      EndProperty
   End
End
Attribute VB_Name = "frmLTVDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'I created this demo to help out everyone who wanted to learn how to
'use a listview. A listview can be very VERY useful when you need
'columns to seperate things. There is no real limit to Listview. What
'I have provided is just the basics that you will use the most on a listview
'but not what it's limited to. You can add icons and even custom sort
'with your own coding. But, just have fun with it. I know it's a LOT different
'from a regular listbox as things like the Listindex and such. But, once you get
'the hang of it I am sure you will like it and use it a lot more than a regular
'listbox...

'All Coding Provided grumtek
'Please, VOTE!

Private Sub cmdCheckAll_Click()
    If ListView1.ListItems.Count = 0 Then: Exit Sub
    Dim intX As Integer
    
    'This will check ALL of the rows
    intX = 1
    ListView1.ListItems(1).Selected = True
    For intX = intX To ListView1.ListItems.Count
        ListView1.ListItems(intX).Selected = True
        ListView1.SelectedItem.Checked = True
    Next
End Sub

Private Sub cmdClear_Click()
    'Resets the intire listview
    ListView1.ListItems.Clear
End Sub

Private Sub cmdDemoIt_Click()
    Dim ListItem As ListItem
    
    'To add to Listview
    Set ListItem = ListView1.ListItems.Add(, , "DEMO 1")
    ListItem.SubItems(1) = "DEMO 2"
    ListItem.SubItems(2) = "DEMO 3"
    
    'Just to set the Listview back to focus..
    'And set the listindex at the bottom..
    'You do NOT have to do this..
    ListView1.ListItems(ListView1.ListItems.Count).Selected = True
    ListView1.SetFocus
End Sub

Private Sub cmdGetText_Click()
    If ListView1.ListItems.Count = 0 Then: Exit Sub
    
    'This will grab every columns text of the selected row.
    MsgBox "Column 1: " & ListView1.SelectedItem.Text & vbCrLf & _
        "Column 2: " & ListView1.SelectedItem.SubItems(1) & vbCrLf & _
        "Column 3: " & ListView1.SelectedItem.SubItems(2), vbInformation + vbOKOnly, "All Text"
End Sub

Private Sub cmdIndex_Click()
    On Error Resume Next
    
    'This is just so you can find the listview index because..
    'It is a bit different from a regular listbox..
    MsgBox ListView1.SelectedItem.Index
End Sub

Private Sub cmdIndex1_Click()
    If ListView1.ListItems.Count = 0 Then: Exit Sub
    'This will show you how to move around in the Listview..
    'Your basic listindex for a Listview..
    Dim intX As Integer
    intX = ListView1.SelectedItem.Index
    If intX = ListView1.ListItems.Count Then
        'Index starts at 1 but the very next line is going..
        'To add + 1 to intX so it doesn't matter..
        intX = 0
    End If
    intX = intX + 1
    ListView1.ListItems(intX).Selected = True
    ListView1.SetFocus
    'Please, NOTE that a listview's listindex starts at 1..
    'And NOT Zero like a listbox!"
End Sub

Private Sub cmdRemove_Click()
    If ListView1.ListItems.Count = 0 Then: Exit Sub
    'This will actually remove the selected item from the Listview.
    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    'Just to set the focus back.
    ListView1.SetFocus
End Sub

Private Sub cmdRemoveChecked_Click()
    On Error Resume Next
    If ListView1.ListItems.Count = 0 Then: Exit Sub
    Dim intX As Integer
    Dim blnStop As Boolean
    blnStop = False
    intX = 1
    
    'This will remove all the columns that are checked from listview
    Do Until blnStop = True
        ListView1.ListItems(intX).Selected = True
        If ListView1.SelectedItem.Checked = True Then
            ListView1.ListItems.Remove (intX)
            intX = intX - 1
        End If
        intX = intX + 1
        If intX = ListView1.ListItems.Count + 1 Then
            Exit Do
        End If
    Loop
End Sub

Private Sub cmdUncheckAll_Click()
    If ListView1.ListItems.Count = 0 Then: Exit Sub
    Dim intX As Integer
    
    'This will Uncheck ALL of the rows
    intX = 1
    ListView1.ListItems(1).Selected = True
    For intX = intX To ListView1.ListItems.Count
        ListView1.ListItems(intX).Selected = True
        ListView1.SelectedItem.Checked = False
    Next
End Sub

Private Sub Form_Load()
    'The following options can always be preset in the Listview properties...
    'I suggest using the settings in properties instead of here but...
    'I'm putting it here for you.
    'I also set the settings in the listview properties as well..
    
    'Shows the column headers.
    ListView1.View = lvwReport
    
    'Prevents people for editing the listview while running.
    ListView1.LabelEdit = lvwManual
    
    'Lets you see which one is selected always.
    ListView1.HideSelection = False
    
    'This is kind of a fun option that will allow the user to drag and drop
    'Columns in different orders. It works really good is pretty cool so
    'But I am only turning it on for you.
    ListView1.AllowColumnReorder = True
    
    'Puts checkboxes beside each row..
    ListView1.Checkboxes = True
    
    'Lets you select the intire row instead of just each column.
    ListView1.FullRowSelect = True
    
    'Finally, This shows the nice looking grid lines on the Listview
    ListView1.GridLines = True
    
    'The following will show you how to ADD columns to the listview
    'with code. But, I already have them set in the properties so
    'I am just going to comment them out..
    
'    On Error Resume Next
'    Dim ListColumn As ListItem
'    Set ListColumn = ListView1.ColumnHeaders.Add(, , "Demo View 1", 1970.64)
'    Set ListColumn = ListView1.ColumnHeaders.Add(, , "Demo View 2", 1970.64)
'    Set ListColumn = ListView1.ColumnHeaders.Add(, , "Demo View 3", 1970.64)
End Sub
