VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TrayIcon v2.1"
   ClientHeight    =   5925
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdShowBalloon 
      Caption         =   "Balloon"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopyUrl 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   4920
      Width           =   735
   End
   Begin VB.Frame fraBalloon 
      Caption         =   "Balloons"
      Height          =   5655
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox chkShown 
         Caption         =   "Is Shown"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CheckBox chkClick 
         Caption         =   "Is Clicked"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   4440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkTimeout 
         Caption         =   "Timed Out"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   5160
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "Warning"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   3360
         Width           =   1575
      End
      Begin VB.OptionButton optType 
         Caption         =   "Information"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   3000
         Width           =   1575
      End
      Begin VB.OptionButton optType 
         Caption         =   "Error"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   2640
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtContent 
         Height          =   960
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtTitle 
         Height          =   360
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Select the balloon style:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Notify me when the shown balloon is:"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Balloon Content:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Frame fraIcon 
      Caption         =   "Change Icon"
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
      Begin VB.Timer tmrAnimate 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2880
         Top             =   360
      End
      Begin VB.CommandButton cmdAnimate 
         Caption         =   "Animate!"
         Height          =   375
         Left            =   2160
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.PictureBox img 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   240
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox img 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":0884
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Lit Dynamite"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Locked Padlock"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "With a nifty trick you can animate the tray"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tooltip"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtTooltip 
         Height          =   360
         Left            =   240
         MaxLength       =   127
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblTooltipSize 
         Alignment       =   2  'Center
         Caption         =   "0 / 127 Allowed Characters"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   3060
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Icon"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=64701"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Thanks for taking the time to download my application! Please comment if you used this code or even if you find a bug:"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuBalloon 
         Caption         =   "Show Balloon"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' The key factor here is WithEvents because
' the class will return events when the user
' clicks onto the task tray ect....
'
Private WithEvents Tray As clsTray
Attribute Tray.VB_VarHelpID = -1
Private SaveState As Integer
Private IconAdded As Boolean

Private Sub cmdAdd_Click()
    '
    ' Add an icon into the task
    ' tray so the user can control it
    '
    If Not Tray.AddIcon Then
        ' Couldnt add the icon into the task tray, uh oh
        MsgBox "Unable to add icon into the task tray!", vbCritical, "Error"
        Exit Sub
    End If

    ' Added sucessfully
    IconAdded = True
    cmdAdd.Enabled = False
    cmdRemove.Enabled = True
    cmdShowBalloon.Enabled = True
End Sub

Private Sub cmdAnimate_Click()
    '
    ' Start and stop an animation
    '
    If cmdAnimate.Caption = "Animate!" Then
        ' Lets begin the animation
        tmrAnimate.Enabled = True
        cmdAnimate.Caption = "Stop"
    Else
        ' We should stop the animation
        cmdAnimate.Caption = "Animate!"
        tmrAnimate.Enabled = False
    End If
End Sub

Private Sub cmdCopyUrl_Click()
    ' And copy the url this was downloaded from so u can comment!
    Clipboard.Clear
    Clipboard.SetText "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=64701"
End Sub

Private Sub cmdRemove_Click()
    '
    ' Remove the icon from tray
    '
    If Not Tray.RemoveIcon Then
        ' Couldnt remove the icon from the task tray
        MsgBox "Unable to remove icon from the task tray, possibly caused by no icon there in the first place.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Removed sucessfully
    IconAdded = False
    cmdRemove.Enabled = False
    cmdAdd.Enabled = True
    cmdShowBalloon.Enabled = False
End Sub

Private Sub cmdShowBalloon_Click()
    Dim Types As EBalloonIconTypes
    
    ' Try and get the type of balloon to display
    If optType(0) Then
        Types = NIIF_ERROR
    ElseIf optType(1) Then
        Types = NIIF_INFO
    ElseIf optType(2) Then
        Types = NIIF_WARNING
    End If
    
    ' By placing a 1 in the timeout variable means that the system
    ' will automatically raise it to the minimum value allowed (usually
    ' the default minimum is 10 seconds)
    Tray.ShowBalloonTip txtContent.Text, txtTitle.Text, Types, 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    '
    ' Setup how we want the
    ' task tray to work and display
    '
    Set Tray = New clsTray
    
    If Not Tray.Initialize(Me) Then
        ' Cant initialize the tray class so END the application immediately
        MsgBox "Fatal Error! Cannot initialize the tray class!", vbCritical, "Fatal Error"
        End
    End If
    
    ' Setup the default values here
    txtTooltip.Text = "Im a little tooltip!"
    txtTitle.Text = "Example Balloon Title"
    txtContent.Text = "Tip! The example only shows you three out of the possible 5 types of balloons, jump down to the source to find out the others!"
    
    ' Initialize settings here
    Tray.AutoRefresh = True
    Tray.Icon = img(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Tray.Visible Then
        ' Tell the user the proper procedure of this classes termination
        MsgBox "You should remove the icon before you exit, but because the tray class is built cleverly all you need to do is destroy the class in order to remove the icon and anything to do with it. Ending the application now will automatically terminate the class!", vbInformation, "Did you know?"
    End If
End Sub

Private Sub mnuBalloon_Click()
    '
    ' Attempt to show a balloon
    '
    If Not Tray.ShowBalloonTip("Welcome to TrayIcon v2.1, I hope you like this example!", "TrayIcon v2.1", NIIF_INFO, 15000) Then
        ' Couldnt display the balloon tip, not had this before so i cant definitely
        ' say what the exact causes of it may be just yet
        MsgBox "Unable to present a balloon tip, possibly your operating system does not support them.", vbCritical, "Error"
    End If
    '
    ' Just because it returned true
    ' doenst mean windows has shown
    ' the balloon to the user. It merely
    ' means windows has got our message
    '
End Sub

Private Sub mnuExit_Click()
    '
    ' Called from the menu to exit
    ' so simply exit then, the class
    ' will handle deleting the icon here
    '
    Unload Me
End Sub

Private Sub optIcon_Click(Index As Integer)
    Tray.Icon = img(Index).Picture.Handle
End Sub

Private Sub tmrAnimate_Timer()
    Static Icon As Boolean
    
    If Icon = False Then
        ' Animation frame 1
        Icon = True
        Call optIcon_Click(0)
    Else
        ' Animation frame 2
        Icon = False
        Call optIcon_Click(1)
    End If
    
End Sub

Private Sub Tray_BalloonClicked()
    ' Alert the user if they wanted notification when the balloon is clicked
    If chkClick Then MsgBox "Your balloon has been clicked!", vbInformation, "Clicked"
End Sub

Private Sub Tray_BalloonShow()
    ' Alert the user if they wanted notification when the balloon is shown
    If chkShown Then MsgBox "Your balloon has been shown! (NOTE: This does not mean it is visible, merely that windows has got the message!)", vbInformation, "Balloon Shown"
End Sub

Private Sub Tray_BalloonTimeout()
    ' Alert the user if they wanted notification when the balloon is timed out
    If chkTimeout Then MsgBox "Your balloon has timed out! (NOTE: This can also mean the user closed the balloon using the X)", vbInformation, "Timeout"
End Sub

Private Sub Tray_MouseDown(Button As Integer)
    '
    ' If they right click on the task tray then
    ' we will simply show them a popup menu
    '
    If Button = 1 Then
        ' And popup the menu
        PopupMenu mnuPopup, vbPopupMenuRightButton, , , mnuShow
    End If
End Sub

Private Sub txtTooltip_Change()
    ' Print the size of the message here so they can see it
    lblTooltipSize.Caption = Len(txtTooltip.Text) & " / 127 Allowed Characters"
    
    ' Add it to the task tray now
    Tray.Tooltip = txtTooltip.Text
End Sub
