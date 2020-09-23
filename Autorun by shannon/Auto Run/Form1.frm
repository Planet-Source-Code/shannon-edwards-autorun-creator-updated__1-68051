VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shannons Autorun Creator"
   ClientHeight    =   2010
   ClientLeft      =   2250
   ClientTop       =   3105
   ClientWidth     =   5760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5760
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Text            =   "Text3"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   1335
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2160
      Width           =   3135
   End
   Begin VB.PictureBox Picture6 
      Height          =   550
      Left            =   3600
      ScaleHeight     =   495
      ScaleWidth      =   2055
      TabIndex        =   15
      Top             =   840
      Width           =   2115
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   550
      Left            =   50
      ScaleHeight     =   495
      ScaleWidth      =   5595
      TabIndex        =   13
      Top             =   1440
      Width           =   5650
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Icon/Program to  new folder "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   5580
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   550
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   11
      Top             =   840
      Width           =   1750
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save  Autorun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   550
      Left            =   50
      ScaleHeight     =   495
      ScaleWidth      =   1695
      TabIndex        =   9
      Top             =   840
      Width           =   1750
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate Autorun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   350
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   2190
      TabIndex        =   7
      Top             =   480
      Width           =   2250
      Begin VB.CommandButton cmdBrowseICO 
         Caption         =   "Browse for ICON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   350
      Left            =   3480
      ScaleHeight     =   285
      ScaleWidth      =   2190
      TabIndex        =   5
      Top             =   120
      Width           =   2250
      Begin VB.CommandButton cmdBrowseEXE 
         Caption         =   "Browse for Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtICO 
      Height          =   285
      Left            =   50
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtEXE 
      Height          =   285
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblIcon 
      Caption         =   "icon="
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblopen 
      Caption         =   "open="
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblAuto 
      Caption         =   "[autorun]"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UPDATED
Private Sub cmdBrowseEXE_Click()
    CommonDialog1.Filter = "Program Files  |*.exe"    'Looks for only files that is EXE's
    CommonDialog1.DialogTitle = "Locate a Program to autorun"    'Changes the CommonDialog title
    CommonDialog1.ShowSave    'Shows ya the Save button
    txtEXE.Text = CommonDialog1.FileName    'shows txtEXE Full Path to the file
    txtOutput.Text = lblAuto.Caption & vbCrLf + lblopen.Caption + CommonDialog1.FileTitle & vbCrLf + lblIcon.Caption + CommonDialog1.FileTitle
    Text2.Text = CommonDialog1.FileTitle
End Sub

Private Sub cmdBrowseICO_Click()
    CommonDialog1.Filter = "Icon Files  |*.ico"    'Shows only Icon Files
    CommonDialog1.DialogTitle = "Locate an icon to use for autorun"    'Changes the CommonDialog title
    CommonDialog1.ShowSave    'again shows the Save button
    txtICO.Text = CommonDialog1.FileName    'Displays the full path to the Icon file, in the textbox
    Text3.Text = CommonDialog1.FileTitle
End Sub

Private Sub cmdCopy_Click()
    If txtICO.Text = "Using Default program Icon" Then
        FileCopy txtEXE.Text, Text1.Text + Text2.Text
        FileCopy CommonDialog1.FileName, Text1.Text & CommonDialog1.FileTitle
    Else
        FileCopy txtEXE.Text, Text1.Text & Text2.Text
        FileCopy txtICO.Text, Text1.Text & Text3.Text
        FileCopy CommonDialog1.FileName, Text1.Text & CommonDialog1.FileTitle
    End If
'Just like before, It displays the message in the message box.. and Ask's you if you would like to open the location to the new folder created, which now has all the files, such as Program,INF file.. or Program,Icon, and inf file
    If MsgBox("Completed!, You can find your file(s) in C:\Autrun..if you would like to go to the location please click on yes, otherwise click no  ", vbYesNo + vbQuestion, "Shannons Autorun") = vbYes Then
        Shell ("C:\Program Files\Internet Explorer\Iexplore.exe C:\Autorun Files\"), vbNormalFocus
    Else    ' if vbNo was clicked then it perfomes the actions listed below
        Exit Sub    'Basically, Exit the sub with no errors, or ne other actions
    End If
    Form1.Caption = "Shannons Autorun Creator"
End Sub

Private Sub cmdExit_Click()
    End    'Quits the program
End Sub

Private Sub cmdGenerate_Click()
    If txtICO.Text = "" Then
        txtOutput.Text = lblAuto.Caption & vbCrLf + lblopen.Caption + CommonDialog1.FileTitle & vbCrLf + lblIcon.Caption + CommonDialog1.FileTitle + ",1"
        txtICO.Text = "Using Default program Icon"
    Else
        txtOutput.Text = lblAuto.Caption & vbCrLf + lblopen.Caption + Text2.Text & vbCrLf + lblIcon.Caption + CommonDialog1.FileTitle
    End If
'Displays the [auto run]
'open= the program file
'icon= the icon file
    MsgBox "Generated complete.. Please continue to next step", vbOKOnly + vbInformation, "Generated Complete"
'again displays the massage above in the Messagebox
    cmdSave.Enabled = True
    cmdCopy.Enabled = True
End Sub

Private Sub cmdSave_Click()
    Dim TheFile, X As Integer    '---> NOT coded by me
    CommonDialog1.CancelError = True
    CommonDialog1.FileName = "Autorun"
    CommonDialog1.Filter = "Autorun |*.inf"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.ShowSave
'****************************************************
'*The code Below, I did NOT make, I found it off    *
'*Planet-Source-Code.. And i forgot who it was from *
'*so i'm truly sorry using your code and not knowing*
'*who you are.. Sorry!                              *
'****************************************************
    TheFile = CommonDialog1.FileName  'using the filename and path from the program or icon
    If Len(Dir$(TheFile)) <> 0 Then    'Checks to see if there is a file with the same name
        X = MsgBox("This file already exists: " + TheFile + ", do you want replace it?", vbYesNo + vbCritical, "Error")    'Displays the message box
        If X = vbNo Then Exit Sub    'Exit the sub withot errors or any other action once clicked on no
    End If
    Open TheFile For Output As #1    'gets the operation ready???.. Again Guessing
    Print #1, txtOutput.Text    'Copies all the information in the textbox
    Close #1    'closes the file (I THINK) I'm just guessing
    Form1.Caption = "Shannons Autorun Creator -Saved complete"
End Sub

Private Sub Form_Load()
    On Error Resume Next
    MkDir "C:\Autorun Files"    'This creats a new directory or folder
    Text1.Text = "C:\Autorun Files\"    'This little guy displays the
'path where the new folder was created, which we have already learned from the above
    cmdSave.Enabled = False
    cmdCopy.Enabled = False
End Sub
