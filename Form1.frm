VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ListView Saving Example"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Load from Registry"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save to Registry"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Checked"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   300
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load from INI file"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save to INI file"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data 3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Tag"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Key"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Data 3"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Data 2"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Data 1"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Saving/Loading a listview Example
'Shows how to add items to a listview and then save those items to either
'an INI file or to Registry. Once saved you can then load the data
'back into the listview.

'Registry Saving - The advantage with saving to registry is speed.
'If you need to be able to retrieve the data straight away then this
'is the way to go. The disadvantage is you need to have a registry
'module in your project.

'INI file saving - The advantage of this method is less code. If you only
'need to save the settings at the unload event then this method is fine.
'Read/Write to an INI file is slower than Registry Read/Write.
'So much so that I've put a delay on enabling the 'Load from INI'
'Command button or the data may not load.

'Personally I prefer to use registry because invariably there's
'other reasons for saving to registry than just listview data in my apps.

'API for INI Read/Write
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim Ret As String
Dim Retlen As String


Public Sub WriteINI(Filename As String, Section As String, Key As String, Text As String)
WritePrivateProfileString Section, Key, Text, Filename
End Sub

Public Function ReadINI(Filename As String, Section As String, Key As String)
Ret = Space$(255)
Retlen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), Filename)
Ret = Left$(Ret, Retlen)
ReadINI = Ret
End Function

Private Sub Command1_Click()
If ListView1.ListItems.Count = 0 Then Exit Sub 'If there's no data bail out now
If FileExists(App.Path + "\LV1.ini") Then Kill App.Path + "\LV1.ini" 'Start a fresh INI file
'First store the number of items we're saving - this makes it easier
'to read the data when we load it later
WriteINI App.Path + "\LV1.ini", "General", "ListItemCount", Str(ListView1.ListItems.Count)
'Go through the listview item by item and save all the data
For x = 1 To ListView1.ListItems.Count
    WriteINI App.Path + "\LV1.ini", "ListItems", Str(x), ListView1.ListItems(x).Text
    WriteINI App.Path + "\LV1.ini", "SubItems1", Str(x), ListView1.ListItems(x).SubItems(1)
    WriteINI App.Path + "\LV1.ini", "SubItems2", Str(x), ListView1.ListItems(x).SubItems(2)
    WriteINI App.Path + "\LV1.ini", "Key", Str(x), ListView1.ListItems(x).Key
    WriteINI App.Path + "\LV1.ini", "Tag", Str(x), ListView1.ListItems(x).Tag
    WriteINI App.Path + "\LV1.ini", "Checked", Str(x), Str(ListView1.ListItems(x).Checked)
Next x
'Need to pause here - just for this demo -  because the INI takes a little
'time to appear in the folder its being written to. If you were saving
'data on the unload event or if you weren't going to load it straight
'away this would not be neccessary.
Sleep 5000
'Clear the listview ready to demonstrate the loading of data. Once again
'in a real app you wouldn't do this.
ListView1.ListItems.Clear
'OK we've probably waited long enough - you can press 'Load from INI file' now
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Dim temp As String, temp1 As String, temp2 As String, temp3 As String, temp4 As String, temp5 As String, temp6 As String
Dim ic As ListItem
'How many items are we loading ?
temp = ReadINI(App.Path + "\LV1.ini", "General", "ListItemCount")
If temp = "" Then Exit Sub 'If we're not loading any then bail out
'Go through all the INI data and read the values into the Listview
For x = 1 To Val(temp)
    temp1 = ReadINI(App.Path + "\LV1.ini", "ListItems", Str(x))
    temp2 = ReadINI(App.Path + "\LV1.ini", "SubItems1", Str(x))
    temp3 = ReadINI(App.Path + "\LV1.ini", "SubItems2", Str(x))
    temp4 = ReadINI(App.Path + "\LV1.ini", "Key", Str(x))
    temp5 = ReadINI(App.Path + "\LV1.ini", "Tag", Str(x))
    temp6 = ReadINI(App.Path + "\LV1.ini", "Checked", Str(x))
    Set ic = ListView1.ListItems.Add(, temp4, temp1)
    ic.Tag = temp5
    ic.SubItems(1) = temp2
    ic.SubItems(2) = temp3
    If temp6 = "True" Then
        ic.Checked = True
    Else
        ic.Checked = False
    End If
Next x

End Sub

Private Sub Command3_Click()
'This is just for the demo - the user would not normally have
'such control over input
Dim ic As ListItem
If Text1 = "" Then
    MsgBox "You need to enter some text for the listitem Data 1"
    Text1.SetFocus
    Exit Sub
End If
If Text4 = "" Then
    MsgBox "You need to enter some text for the Key"
    Text1.SetFocus
    Exit Sub
End If
For x = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(x).Key = Text4 Then
        'The Key has to be different
        MsgBox "Key is not unique. Please enter a different Key"
        Exit Sub
    End If
    If ListView1.ListItems(x).Tag = Text5 Then
        'The Tag should be different in most cases
        'but it's not vital like the key
        MsgBox "This Tag is in use.Try another."
        Exit Sub
    End If
Next x
'load the data into the listview
Set ic = ListView1.ListItems.Add(, Text4, Text1)
ic.Tag = Text5
ic.SubItems(1) = Text2
ic.SubItems(2) = Text3
ic.Checked = Check1
'Empty the textboxes ready for the next entry
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Check1.Value = 0
Text1.SetFocus

End Sub

Private Sub Command4_Click()
'Delete the old key just to keep things tidy in registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo"
'First store the number of items we're saving - this makes it easier
'to read the data when we load it later
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "ListItemCount", Str(ListView1.ListItems.Count)
'Go through the listview item by item and save all the data
For x = 1 To ListView1.ListItems.Count
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "ListItems" + Str(x), ListView1.ListItems(x).Text
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "SubItems1" + Str(x), ListView1.ListItems(x).SubItems(1)
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "SubItems2" + Str(x), ListView1.ListItems(x).SubItems(2)
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "Key" + Str(x), ListView1.ListItems(x).Key
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "Tag" + Str(x), ListView1.ListItems(x).Tag
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "Checked" + Str(x), Str(ListView1.ListItems(x).Checked)
Next x
'Clear the listview ready to demonstrate the loading of data. Once again
'in a real app you wouldn't do this.
ListView1.ListItems.Clear

End Sub

Private Sub Command5_Click()
Dim temp As String, temp1 As String, temp2 As String, temp3 As String, temp4 As String, temp5 As String, temp6 As String
Dim ic As ListItem
'How many items are we loading ?
temp = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "ListItemCount")
If temp = "" Then Exit Sub 'If we're not loading any then bail out
'Go through all the data and read the values into the Listview
For x = 1 To Val(temp)
    temp1 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "ListItems" + Str(x))
    temp2 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "SubItems1" + Str(x))
    temp3 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "SubItems2" + Str(x))
    temp4 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "Key" + Str(x))
    temp5 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "Tag" + Str(x))
    temp6 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\LVDemo", "Checked" + Str(x))
    Set ic = ListView1.ListItems.Add(, temp4, temp1)
    ic.Tag = temp5
    ic.SubItems(1) = temp2
    ic.SubItems(2) = temp3
    If temp6 = "True" Then
        ic.Checked = True
    Else
        ic.Checked = False
    End If
Next x

End Sub

Private Sub Form_Load()
Me.Show
Text1.SetFocus

End Sub
Function FileExists(ByVal Filename As String) As Integer
'Just used here to varify the presence of the INI file
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(Filename)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
                End
            End If
    End Select
End Function
