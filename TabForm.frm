VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "MSWEBDVD.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TabForm 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000000&
      Height          =   6630
      Index           =   0
      Left            =   90
      ScaleHeight     =   6570
      ScaleWidth      =   9840
      TabIndex        =   1
      Top             =   390
      Width           =   9900
      Begin VB.CommandButton Command11 
         Caption         =   "Play"
         Height          =   285
         Left            =   4770
         TabIndex        =   40
         Top             =   5820
         Width           =   795
      End
      Begin VB.Frame Frame3 
         Caption         =   "Menus"
         Height          =   1845
         Left            =   8550
         TabIndex        =   36
         Top             =   60
         Width           =   1215
         Begin VB.CommandButton Command3 
            Caption         =   "Jump To"
            Height          =   285
            Left            =   60
            TabIndex        =   38
            Top             =   1470
            Width           =   1065
         End
         Begin VB.ListBox List1 
            Height          =   1230
            ItemData        =   "TabForm.frx":0000
            Left            =   60
            List            =   "TabForm.frx":0016
            TabIndex        =   37
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Play"
         Height          =   285
         Index           =   0
         Left            =   8790
         TabIndex        =   33
         Top             =   6150
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stop"
         Height          =   285
         Index           =   1
         Left            =   7770
         TabIndex        =   32
         Top             =   6150
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pause"
         Height          =   285
         Index           =   2
         Left            =   6750
         TabIndex        =   31
         Top             =   6150
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Last Chapter"
         Height          =   285
         Index           =   5
         Left            =   4170
         TabIndex        =   30
         Top             =   6150
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next Chapter"
         Height          =   285
         Index           =   6
         Left            =   5460
         TabIndex        =   29
         Top             =   6150
         Width           =   1245
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Time Done Properly not like others i have seen"
         ForeColor       =   &H80000007&
         Height          =   1275
         Left            =   5730
         TabIndex        =   24
         Top             =   3390
         Width           =   4035
         Begin VB.Label lblTimeTrackerValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time 00:00:00:00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   240
            Left            =   690
            TabIndex        =   28
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label lbltotaltime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   240
            Left            =   180
            TabIndex        =   27
            Top             =   480
            Width           =   975
         End
         Begin VB.Label LblChapter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chapter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   240
            Left            =   420
            TabIndex        =   26
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Lbllang 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Language"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   240
            Left            =   210
            TabIndex        =   25
            Top             =   960
            Width           =   915
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "TabForm.frx":004A
         Top             =   3090
         Width           =   3795
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Zoom In"
         Height          =   285
         Left            =   6180
         TabIndex        =   22
         Top             =   1200
         Width           =   945
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save Book Mark"
         Height          =   285
         Left            =   6210
         TabIndex        =   21
         Top             =   60
         Width           =   1875
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Restore Book Mark"
         Height          =   285
         Left            =   6210
         TabIndex        =   20
         Top             =   360
         Width           =   1875
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Zoom Out"
         Height          =   285
         Left            =   6180
         TabIndex        =   19
         Top             =   900
         Width           =   945
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fast Forward and Rewind"
         Height          =   1395
         Left            =   6450
         TabIndex        =   12
         Top             =   4710
         Width           =   3315
         Begin VB.ListBox List2 
            Height          =   1035
            ItemData        =   "TabForm.frx":0050
            Left            =   240
            List            =   "TabForm.frx":0063
            TabIndex        =   16
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Rewind"
            Height          =   315
            Index           =   4
            Left            =   690
            TabIndex        =   15
            Top             =   960
            Width           =   1245
         End
         Begin VB.CommandButton Command1 
            Caption         =   "FastForward"
            Height          =   315
            Index           =   3
            Left            =   1980
            TabIndex        =   14
            Top             =   960
            Width           =   1245
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Step"
            Height          =   285
            Index           =   7
            Left            =   690
            TabIndex        =   13
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "0.5 = Slow Motion"
            Height          =   225
            Index           =   0
            Left            =   720
            TabIndex        =   18
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label Label2 
            Caption         =   "Gos back in rewind"
            Height          =   225
            Index           =   1
            Left            =   1680
            TabIndex        =   17
            Top             =   630
            Width           =   1515
         End
      End
      Begin VB.ListBox List3 
         Height          =   2205
         Left            =   3960
         TabIndex        =   11
         Top             =   3120
         Width           =   765
      End
      Begin VB.CommandButton Command7 
         Caption         =   "List Chapters"
         Height          =   465
         Left            =   3960
         TabIndex        =   10
         Top             =   5340
         Width           =   795
      End
      Begin VB.ListBox List4 
         Height          =   2205
         Left            =   4770
         TabIndex        =   9
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton Command8 
         Caption         =   "List Titles"
         Height          =   465
         Left            =   4770
         TabIndex        =   8
         Top             =   5340
         Width           =   795
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Play"
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Top             =   5820
         Width           =   795
      End
      Begin MSWEBDVDLibCtl.MSWebDVD DVD 
         Height          =   2745
         Left            =   90
         TabIndex        =   34
         Top             =   60
         Width           =   6045
         _cx             =   10663
         _cy             =   4842
         DisableAutoMouseProcessing=   0   'False
         BackColor       =   1048592
         EnableResetOnStop=   0   'False
         ColorKey        =   1048592
         WindowlessActivation=   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Vote for me PLEASE, Click Here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6240
         TabIndex        =   43
         Top             =   2070
         Width           =   2505
      End
      Begin VB.Label Label5 
         Caption         =   "Email Me, Click Here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6270
         TabIndex        =   42
         Top             =   1680
         Width           =   1905
      End
      Begin VB.Label Label4 
         Caption         =   $"TabForm.frx":007C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   885
         Left            =   6240
         TabIndex        =   41
         Top             =   2460
         Width           =   3525
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   300
         X2              =   300
         Y1              =   2610
         Y2              =   2910
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   270
         X2              =   270
         Y1              =   2610
         Y2              =   2910
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "V"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   2850
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Event Codes triggerd from the DVD"
         Height          =   225
         Left            =   780
         TabIndex        =   35
         Top             =   2850
         Width           =   2565
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3195
      Index           =   5
      Left            =   6270
      ScaleHeight     =   3135
      ScaleWidth      =   3390
      TabIndex        =   6
      Top             =   3720
      Width           =   3450
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3750
      Index           =   4
      Left            =   5670
      ScaleHeight     =   3690
      ScaleWidth      =   3990
      TabIndex        =   5
      Top             =   3180
      Width           =   4050
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   4290
      Index           =   3
      Left            =   5010
      ScaleHeight     =   4230
      ScaleWidth      =   4650
      TabIndex        =   4
      Top             =   2640
      Width           =   4710
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   4770
      Index           =   2
      Left            =   4440
      ScaleHeight     =   4710
      ScaleWidth      =   5220
      TabIndex        =   3
      Top             =   2160
      Width           =   5280
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   5190
      Index           =   1
      Left            =   3870
      ScaleHeight     =   5130
      ScaleWidth      =   5790
      TabIndex        =   2
      Top             =   1710
      Width           =   5850
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7125
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   12568
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 1"
            Key             =   "picture1p"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 2"
            Key             =   "picture2p"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 3"
            Key             =   "picture3p"
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 4"
            Key             =   "picture4p"
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 5"
            Key             =   "picture5p"
            Object.Tag             =   "5"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TabForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Picture1p()
End Sub
Private Sub Picture2p()
End Sub
Private Sub Picture3p()
End Sub
Private Sub Picture4p()
End Sub
Private Sub Picture5p()
End Sub

Private Sub Command10_Click()
MsgBox "This function will not work on my machine if someone can make it work let me know how"
Exit Sub
'DVD.PlayChapterInTitle(List4.ListIndex, List3.ListIndex)
End Sub

Private Sub Command11_Click()
MsgBox "This function will not work on my machine if someone can make it work let me know how"
Exit Sub

If List4.Text = "" Then
MsgBox "Select a title to jump to first", vbOKOnly, "How to use"
Exit Sub
End If
If DVD.PlayState = dvdState_Running Then
If DVD.CurrentDomain = 4 Then
If DVD.UOPValid(0 * 0) Then ' Check MSWEBDVD if its ok to Playtitle
DVD.PlayTitle (List4.ListIndex)
End If
Exit Sub
End If
End If
CallDomain
End Sub

Private Sub Form_Load()
TabStrip1.Tabs("picture1p").Selected = True ' select tab one
List2.Selected(1) = True ' sets ff & rewind speed
DVD.Render (0)
Call Command8_Click ' Populate List4 with the titles avalible
Call Command7_Click ' Populate List3 with the chapters avalible
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo ErrLinec ' If an error occurs trap it
Select Case Index
Case 0
    If DVD.CurrentDomain >= 1 Then
            DVD.Play 'plays the current DVD title
        Exit Sub
    End If

Case 1
'First check the domain is valid for the oparation then ask
'MSWEBDVD if its ok to preform
    If DVD.CurrentDomain >= 1 Then
        If DVD.UOPValid(0 * 4) = True Then
                DVD.Stop 'Stops the current DVD title
            Exit Sub
        End If
    End If
    
Case 2
    If DVD.CurrentDomain = 4 Then
        If DVD.UOPValid(0 * 40000) = True Then 'Double check its ok to pause
              DVD.Pause 'pauses playback at the current location
        Exit Sub
        End If
    End If
    
Case 3
    If DVD.CurrentDomain = 4 Then
        If DVD.UOPValid(0 * 80) Then ' Check MSWEBDVD if its ok to PlayForwards
            Select Case List2.ListIndex
                Case 0
                    DVD.PlayForwards 0.5 ' slow motion
                Case 1
                    DVD.PlayForwards 1 ' Standard play speed good for reverse
                Case 2
                    DVD.PlayForwards 2 ' Twice the play speed
                Case 3
                    DVD.PlayForwards 4
                Case 4
                    DVD.PlayForwards 8
                End Select
            Exit Sub
        End If
    End If
Case 4
    If DVD.CurrentDomain = 4 Then
        If DVD.UOPValid(0 * 100) Then ' Check MSWEBDVD if its ok to PlayBackwards
            Select Case List2.ListIndex
                Case 0
                    DVD.PlayBackwards 0.5
                Case 1
                    DVD.PlayBackwards 1
                Case 2
                    DVD.PlayBackwards 2
                Case 3
                    DVD.PlayBackwards 4
                Case 4
                    DVD.PlayBackwards 8
            End Select
            Exit Sub
        End If
    End If

Case 5
  
        If DVD.UOPValid(0 * 20) Then ' Check MSWEBDVD if its ok to PlayPrevChapter or ReplayChapter
            DVD.PlayPrevChapter
                Exit Sub
        End If


Case 6
 
        If DVD.UOPValid(0 * 40) Then ' Check MSWEBDVD if its ok to PlayNextChapter
                DVD.PlayNextChapter
            Exit Sub
        End If


Case 7
    If DVD.CurrentDomain = 4 Then
        Select Case List2.ListIndex
            Case 0
                'can not step 0.5
            Case 1
                DVD.Step 1
            Case 2
                DVD.Step 2
            Case 3
                DVD.Step 4
            Case 4
                DVD.Step 8
            End Select
        Exit Sub
    End If
End Select

CallDomain ' If non come true tell the user why
ErrLinec:
Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
End Sub

Private Sub Command2_Click()
If Val(Command6.Tag) = 0 Then ' prevent the message coming up every time
MsgBox "Click and hold the left mouse button on the video to move around", vbOKOnly, "Goto options to turn message's off"
Command6.Tag = 1 ' Zoom in button
Command2.Tag = 1 ' Zoom Out button
End If

DVD.Zoom 360, 270, 2
End Sub

Private Sub Command3_Click()

On Error GoTo ErrLinec

If DVD.CurrentDomain = 1 Then GoTo SKIP ' if the disc is stopped tell the user
If DVD.CurrentDomain = 5 Then
    GoTo SKIP
            Else
                Select Case List1.ListIndex
                    Case 0: Call DVD.ShowMenu(3) 'Root
                    Case 1: Call DVD.ShowMenu(2) 'Title
                    Case 2: Call DVD.ShowMenu(5) 'Audio
                    Case 3: Call DVD.ShowMenu(6) 'Angle
                    Case 4: Call DVD.ShowMenu(7) 'Chapter
                    Case 5: Call DVD.ShowMenu(4) 'Subpicture
                End Select
            Exit Sub
End If
SKIP:
CallDomain 'ensure DVD Navigator is in a valid domain
Exit Sub
ErrLinec:
Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear

End Sub

Function CallDomain() 'Call this function to ensure the DVD Navigator is in a valid domain
' for the method you are about to call. For example, before calling PlayTitle,
' check the CurrentDomain property to make sure that the DVD Navigator is not in the Stop or First Play domain.

Select Case DVD.CurrentDomain
Case 1: MsgBox "This operation is not permitted at this time, please wait ", vbOKOnly, "First play"
Case 2: MsgBox "This operation is not permitted from the Video Manager Menu", vbOKOnly, "Video Manager Menu"
Case 3: MsgBox "This operation is not permitted from the Video Title Set Menu", vbOKOnly, "Video Title Menu"
Case 4: MsgBox "This operation is not permitted while the disc is playing, click stop first ", vbOKOnly, "Disc Playing"
Case 5: MsgBox "This operation is not permitted while the disc is stopped, click play first ", vbOKOnly, "Disc Stopped"
End Select

End Function


Private Sub Command4_Click()

On Error GoTo SKIP
DVD.DeleteBookmark ' To prevent an error get rid of any old book marks
DoEvents
DVD.SaveBookmark
MsgBox "To return to this book mark use the file menu Load Book Mark", vbOKOnly, "Book mark saved"
Exit Sub
SKIP:
MsgBox "An error occured durring save, please try while the DVD is running"
End Sub

Private Sub Command5_Click()

            If DVD.CurrentDomain = 1 Then GoTo SKIP
            If DVD.CurrentDomain = 2 Then GoTo SKIP
            If DVD.CurrentDomain = 3 Then GoTo SKIP
            If DVD.CurrentDomain = 4 Then GoTo SKIP
                    CallDomain 'ensure DVD Navigator is in a valid domain
            Exit Sub

SKIP:
On Error GoTo SKIPBOOK
DVD.RestoreBookmark
Exit Sub
SKIPBOOK:
MsgBox "Either there is no book mark stored or the book mark stored is not for this title", vbOKOnly, "Can not restore"

End Sub

Private Sub Command6_Click()
If Val(Command6.Tag) = 0 Then ' prevent the message coming up every time
MsgBox "Click and hold the left mouse button on the video screen to move around", vbOKOnly, "How to use"
Command6.Tag = 1 ' Zoom in button
Command2.Tag = 1 ' Zoom Out button
End If
DVD.Zoom 360, 270, 0.5
End Sub

Private Sub Command7_Click()
For i = 0 To DVD.GetNumberOfChapters(1) - 1
List3.AddItem i
Next i
End Sub

Private Sub Command8_Click()
For i = 0 To DVD.TitlesAvailable - 1
List4.AddItem i
Next i
End Sub

Private Sub Command9_Click()
If List3.Text = "" Then
MsgBox "Select a chapter to jump to first", vbOKOnly, "How to use"
Exit Sub
End If

If DVD.CurrentDomain = 4 Then
If DVD.UOPValid(0 * 40) Then ' Check MSWEBDVD if its ok to PlayNextChapter
DVD.PlayChapter (List3.ListIndex)
End If
Exit Sub
End If
CallDomain
End Sub

Private Sub DVD_DVDNotify(ByVal lEventCode As Long, ByVal lParam1 As Variant, ByVal lParam2 As Variant)
DVD.NotifyParentalLevelChange (True) 'application is notified it encounters video segments with a rating more restrictive than the overall rating for the disc
Dim id
Dim paramList
Dim codeConstant
Dim strlEventCode
Dim strlParam1
Dim szDVDlEventCode(70) ' Only needs to go to 25 but for testing I have set it higher just in case of an odscure EC
Dim eventOffset
DVD.ShowCursor = True

            On Local Error GoTo ErrLine
If 282 = lEventCode Then '282 is the event code for the time event, pass in param1 to get you the current time-convert to hh:mm:ss:ff format with DVDTimeCode2BSTR API
    If lblTimeTrackerValue.Caption <> CStr(DVD.DVDTimeCode2bstr(lParam1)) Then _
        lblTimeTrackerValue.Caption = "Time " & CStr(DVD.DVDTimeCode2bstr(lParam1))
        Lbllang.Caption = "Language " & DVD.GetLangFromLangID(0)
        LblChapter.Caption = "Chapter " & DVD.CurrentChapter & " off " & DVD.GetNumberOfChapters(1)
        lbltotaltime.Caption = "TotalTime " & DVD.TotalTitleTime

    End If

paramList = id
codeConstant = id

strlEventCode = szDVDlEventCode(eventOffset)
codeConstant = "DVDNotify lEventCode = " + CStr(lEventCode) + " = " + Str(lEventCode)
paramList = "lParam1 = " + CStr(lParam1) + "  lParam2 = " + CStr(lParam2)

eventOffset = lEventCode - 257 'Convet from a c++ readable value to a vb readable value

'If someone has a KARAOKE dvd try to figure out some of the missing EC's
'If you do figure more of these EC's out, please email me them

szDVDlEventCode(0) = " (EC_DVD_DOMAIN_CHANGE)"               '257 Indicates the DVD player's new domain.
szDVDlEventCode(1) = " (EC_DVD_TITLE_CHANGE)"                '258 Occurs when the current title number changes.
szDVDlEventCode(2) = " (EC_DVD_CHAPTER_START)"               '259 Signals that the DVD player has started playback of a new program in the TT_DOM domain.
szDVDlEventCode(3) = " (EC_DVD_AUDIO_STREAM_CHANGE)"         '260 Signals that the current user audio stream number has changed for the main title.
szDVDlEventCode(4) = " (EC_DVD_SUBPICTURE_STREAM_CHANGE)"    '261 Signals that the current user subpicture stream number has changed for the main title.
szDVDlEventCode(5) = " (EC_DVD_ANGLE_CHANGE)"                '262 Signals that either the number of available angles has changed, or the current user angle number has changed.
szDVDlEventCode(6) = " (EC_DVD_BUTTON_CHANGE)"               '263 Signals that either the number of available buttons has changed, or the currently selected button number has changed.
szDVDlEventCode(7) = " (EC_DVD_VALID_UOPS_CHANGE)"           '264 Signals that the available set of IDvdControl interface methods has changed.
szDVDlEventCode(8) = " (EC_DVD_STILL_ON)"                    '265 Signals the beginning of any still (PGC, Cell, or VOBU).
szDVDlEventCode(9) = " (EC_DVD_STILL_OFF)"                   '266 Signals the end of any still (PGC, Cell, or VOBU).
szDVDlEventCode(10) = " (EC_DVD_CURRENT_TIME)"               '267 Signals the beginning of each video object unit (VOBU), which occurs every 0.4 to 1.0 seconds.
szDVDlEventCode(11) = " (EC_DVD_ERROR)"                      '268 Signals a DVD error condition.
szDVDlEventCode(12) = " (EC_DVD_WARNING)"                    '269 Signals a DVD warning condition.
szDVDlEventCode(13) = " (EC_DVD_CHAPTER_AUTOSTOP)"           '270 Indicates that playback has stopped as the result of a call to the ChapterPlayAutoStop method.
szDVDlEventCode(14) = " (EC_DVD_NO_FP_PGC)"                  '271 Signals that the DVD does not have a First Play Program Chain (FP_PGC) and that the DVD Navigator will not automatically load any program chain (PGC) or start playback.
szDVDlEventCode(15) = " (EC_DVD_PLAYBACK_RATE_CHANGE)"       '272 Signals that a rate change in the playback has been initiated.
szDVDlEventCode(16) = " (???????????????????????)"           '????????????????????????????????????????
szDVDlEventCode(17) = " (EC_DVD_PLAYBACK_STOPPED)"           '274 Indicates that playback has been stopped. The DVD Navigator has completed playback of the program chain (PGC) and did not find any other branching instruction for subsequent playback.
szDVDlEventCode(18) = " (EC_DVD_ANGLES_AVAILABLE)"           '275 Occurs when an angle block is being played and angle changes can be performed.
szDVDlEventCode(19) = " (EC_DVD_ANGLE_CHANGE)"
szDVDlEventCode(20) = " (???????????????????????)"           '????????????????????????????????????????
szDVDlEventCode(21) = " (???????????????????????)"           '????????????????????????????????????????
szDVDlEventCode(22) = " (???????????????????????)"           '????????????????????????????????????????
szDVDlEventCode(23) = " (???????????????????????)"           '????????????????????????????????????????
szDVDlEventCode(24) = " (???????????????????????)"           '????????????????????????????????????????
szDVDlEventCode(25) = " (EC_DVD_CURRENT_HMSF_TIME)"          '282 is the event code for the time event

a = TabForm.Text1.Text

TabForm.Text1.Text = szDVDlEventCode(eventOffset) & "  Event Code : " & eventOffset & vbNewLine & a
a = ""
'Below the ones that got away, if you can work these out please mail me

'EC_DVD_ANGLES_AVAILABLE
'EC_DVD_PLAYPERIOD_AUTOSTOP
'EC_DVD_PARENTAL_LEVEL_CHANGE
'EC_DVD_KARAOKE_MODE
'EC_DVD_DISC_INSERTED
'EC_DVD_DISC_EJECTED
'EC_DVD_CMD_END
'EC_DVD_CHAPTER_AUTOSTOP
'EC_DVD_BUTTON_AUTO_ACTIVATED
'EC_DVD_WARNING

            Exit Sub
ErrLine:
            Err.Clear
            Exit Sub
            
End Sub


Private Sub Label4_Click()
On Error GoTo er
 Call Run("http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&B1=Quick+Search&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=jamie+pocock&optSort=Alphabetical")
 Exit Sub
er: MsgBox "An error has occured while trying to connect to the internet" & vbNewLine & "Please ensure that your internet connection is functioning correctly and try again."

End Sub

Private Sub Label5_Click()
On Error GoTo er
 Call Run("mailto:MSWEBDVD@micracom.com")
 Exit Sub
er: MsgBox "An error has occured while trying to connect to the internet" & vbNewLine & "Please ensure that your internet connection is functioning correctly and try again."

End Sub

Private Sub Label6_Click()
On Error GoTo er
 Call Run("http://www.planetsourcecode.com/xq/ASP/txtCodeId.30717/lngWId.1/qx/vb/scripts/ShowCode.htm")
 Exit Sub
er: MsgBox "An error has occured while trying to connect to the internet" & vbNewLine & "Please ensure that your internet connection is functioning correctly and try again."

End Sub

Private Sub TabStrip1_Click()
On Error GoTo 10
Picture1(TabStrip1.SelectedItem.Index - 1).Move TabStrip1.ClientLeft _
, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight

Picture1(TabStrip1.SelectedItem.Index - 1).ZOrder
Select Case TabStrip1.SelectedItem.Index

    Case 1
        Picture1p
    Case 2
        Picture2p
    Case 3
        Picture3p
    Case 4
        Picture4p
    Case 5
        Picture5p
End Select
10: Exit Sub
End Sub
