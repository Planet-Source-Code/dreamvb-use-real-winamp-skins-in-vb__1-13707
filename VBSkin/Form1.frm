VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test Numbers"
      Height          =   315
      Left            =   120
      TabIndex        =   44
      Top             =   3510
      Width           =   1305
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   1980
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0002
      TabIndex        =   43
      Top             =   1980
      Width           =   4125
   End
   Begin VB.PictureBox playList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   4125
      TabIndex        =   42
      Top             =   1740
      Width           =   4125
   End
   Begin VB.PictureBox sk8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1860
      ScaleHeight     =   315
      ScaleWidth      =   765
      TabIndex        =   39
      Top             =   6450
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox sk7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4215
      ScaleHeight     =   615
      ScaleWidth      =   825
      TabIndex        =   34
      Top             =   6720
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox sk6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   765
      ScaleHeight     =   285
      ScaleWidth      =   705
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5505
      Top             =   2535
   End
   Begin VB.PictureBox sk5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3450
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   19
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox sk4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   690
      ScaleHeight     =   225
      ScaleWidth      =   960
      TabIndex        =   18
      Top             =   7215
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox sk3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   3525
      ScaleHeight     =   180
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   7245
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox sk2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1890
      ScaleHeight     =   360
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   7095
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Sk1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   2760
      ScaleHeight     =   390
      ScaleWidth      =   630
      TabIndex        =   1
      Top             =   7020
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox titlebar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   0
      Width           =   4125
      Begin VB.PictureBox mMin 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   3645
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   15
         Top             =   60
         Width           =   135
      End
      Begin VB.PictureBox mExit 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   3960
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   14
         Top             =   60
         Width           =   135
      End
      Begin VB.PictureBox mMinsize 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   3810
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   13
         Top             =   60
         Width           =   135
      End
      Begin VB.PictureBox mAbout 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   90
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   12
         Top             =   60
         Width           =   135
      End
      Begin VB.Label minButtons 
         Height          =   135
         Index           =   5
         Left            =   3240
         TabIndex        =   32
         Top             =   45
         Width           =   135
      End
      Begin VB.Label minButtons 
         Height          =   135
         Index           =   4
         Left            =   3090
         TabIndex        =   31
         Top             =   45
         Width           =   135
      End
      Begin VB.Label minButtons 
         Height          =   135
         Index           =   3
         Left            =   2940
         TabIndex        =   30
         Top             =   45
         Width           =   135
      End
      Begin VB.Label minButtons 
         Height          =   135
         Index           =   2
         Left            =   2790
         TabIndex        =   29
         Top             =   45
         Width           =   135
      End
      Begin VB.Label minButtons 
         Height          =   135
         Index           =   1
         Left            =   2640
         TabIndex        =   28
         Top             =   45
         Width           =   135
      End
      Begin VB.Label minButtons 
         Height          =   135
         Index           =   0
         Left            =   2490
         TabIndex        =   27
         Top             =   45
         Width           =   135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   120
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   45
         Width           =   105
      End
   End
   Begin VB.PictureBox Base 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   0
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   2
      Top             =   0
      Width           =   4125
      Begin VB.PictureBox mVolume 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1575
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   68
         TabIndex        =   38
         Top             =   840
         Width           =   1020
         Begin VB.PictureBox mVolPos 
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   150
            Left            =   75
            ScaleHeight     =   10
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   12
            TabIndex        =   40
            Top             =   30
            Width           =   180
         End
      End
      Begin VB.PictureBox mpl 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   3630
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   37
         Top             =   870
         Width           =   345
      End
      Begin VB.PictureBox mEq 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   3300
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   36
         Top             =   870
         Width           =   330
      End
      Begin VB.PictureBox mbut 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   228
         Left            =   3150
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   35
         Top             =   1320
         Width           =   420
      End
      Begin VB.PictureBox shuff 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   228
         Left            =   2475
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   33
         Top             =   1320
         Width           =   690
      End
      Begin VB.PictureBox num4 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   735
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   26
         Top             =   390
         Width           =   135
      End
      Begin VB.PictureBox num3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   915
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   25
         Top             =   390
         Width           =   135
      End
      Begin VB.PictureBox num2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1155
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   24
         Top             =   390
         Width           =   135
      End
      Begin VB.PictureBox num1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1335
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   22
         Top             =   390
         Width           =   135
      End
      Begin VB.PictureBox mStereo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   3585
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   21
         Top             =   600
         Width           =   435
      End
      Begin VB.PictureBox mMono 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   3180
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   20
         Top             =   600
         Width           =   405
      End
      Begin VB.PictureBox ProgBar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   225
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   248
         TabIndex        =   16
         Top             =   1065
         Width           =   3720
         Begin VB.PictureBox PosBar 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   120
            Left            =   15
            ScaleHeight     =   8
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   28
            TabIndex        =   17
            Top             =   15
            Width           =   420
         End
      End
      Begin VB.PictureBox mcd 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2055
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   1320
         Width           =   315
      End
      Begin VB.PictureBox mNext 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1605
         ScaleHeight     =   270
         ScaleWidth      =   345
         TabIndex        =   10
         Top             =   1305
         Width           =   339
      End
      Begin VB.PictureBox mstop 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1245
         ScaleHeight     =   270
         ScaleWidth      =   330
         TabIndex        =   9
         Top             =   1305
         Width           =   330
      End
      Begin VB.PictureBox mpause 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   900
         ScaleHeight     =   270
         ScaleWidth      =   330
         TabIndex        =   8
         Top             =   1305
         Width           =   330
      End
      Begin VB.PictureBox mPlay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   540
         ScaleHeight     =   270
         ScaleWidth      =   330
         TabIndex        =   7
         Top             =   1305
         Width           =   330
      End
      Begin VB.PictureBox mBack 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   225
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   6
         Top             =   1305
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1665
         TabIndex        =   41
         Top             =   330
         Width           =   645
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Hi this is a simple was of how to inport real winamp skins into your applactions
' Well anyway about a month ago I heard some requests about how to add winamp skins to a program
' or were they might be able to find a ocx file. so I set about doing my own since there are loads of
' Skin exampls in vb useing bitblt api function well anyway this is what i have done so far
' you can load almost any winamp skin into this. I will in about a month make this into a ocx
' and just do a bit of cleaning up the code. well hope you like it

' Name Ben Jones
' Website http://www.dreamvb.s5.com
' Email Dreamvb@yahoo.com
 
 ' p.s you make any inprovment please let me know or send me what you have done.
 ' Thanks
 



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'1755

Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1

Const TitleBar_Height = 14.5
Const TitleBar_Width = 275
Const Buttons_Height = 18
Const Buttons_Width = 136

Const Base_Height = 116
Const Base_Width = 275

Public DragFlag, SlideFlag, PlVisFlag, SlideFlag1
Public IX, IY, TX, TY, FX, FY
Public x1, y1, x2, y2, x3, y3

Dim mTimer As Integer
Dim n1, n2, n3, n4 As Integer
Dim UpDown As Boolean
Dim mPlayLst As Boolean

Sub ShowPlayList()
    Select Case mPlayLst
        Case False
            Form1.Height = 1755
            mPlayLst = True
        Case True
            Form1.Height = 4755
            mPlayLst = False
        End Select
        
End Sub



Sub MinUpDown()
    Select Case UpDown
        Case False
            BitBlt titlebar.hDC, -32, 0, 307, 14.5, Sk1.hDC, -5, 14.6 * 2, SRCCOPY
            titlebar.Refresh
            Base.Visible = False
            Form1.Height = 240
            UpDown = True
        Case True
            BitBlt titlebar.hDC, -32, 0, 307, 14.5, Sk1.hDC, -5, 0, SRCCOPY
            titlebar.Refresh
            Base.Visible = True
            Form1.Height = 4755
            UpDown = False
        End Select
        
End Sub
Sub IsMonoOrSteroeo(isMono As Boolean)
    If isMono Then
        BitBlt mMono.hDC, 0, 0, 29, 12, sk5.hDC, 29, 0, SRCCOPY ' Mono Off
        BitBlt mStereo.hDC, 0, 0, 29, 12, sk5.hDC, 0, 12, SRCCOPY ' Stereo Off
        mMono.Refresh
        mStereo.Refresh
     Else
        BitBlt mMono.hDC, 0, 0, 29, 12, sk5.hDC, 29, 12, SRCCOPY ' Mono Off
        BitBlt mStereo.hDC, 0, 0, 29, 12, sk5.hDC, 0, 0, SRCCOPY ' Stereo Off
        mMono.Refresh
        mStereo.Refresh
        
    End If
    
End Sub





Private Sub Command1_Click()
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Load()
    n1 = 0
    n2 = 0
    n3 = 0
    n4 = 0

      Sk1.Picture = LoadPicture(App.Path & "\Skins\TITLEBAR.BMP") ' Loads the title bar bitmap
      sk2.Picture = LoadPicture(App.Path & "\Skins\MAIN.BMP")     ' Loads the Main amp bitmap
      sk3.Picture = LoadPicture(App.Path & "\Skins\cbuttons.bmp") ' Loads the buttons bitmap
      sk4.Picture = LoadPicture(App.Path & "\Skins\POSBAR.BMP")   ' Loads the posbar bitmap
      sk5.Picture = LoadPicture(App.Path & "\Skins\MONOSTER.BMP") ' Loads the monoster bitmap
      sk6.Picture = LoadPicture(App.Path & "\Skins\NUMBERS.BMP")
      sk7.Picture = LoadPicture(App.Path & "\Skins\SHUFREP.BMP")
      sk8.Picture = LoadPicture(App.Path & "\Skins\Volume.bmp")
      
      
      
      
      
      BitBlt mBack.hDC, 0, 0, 22, 22, sk3.hDC, 0, 0, SRCCOPY ' Back button
      
      BitBlt mPlay.hDC, 0, 0, 22, 22, sk3.hDC, 22, 0, SRCCOPY ' Play Button
      
      BitBlt mpause.hDC, -1, 0, 22.6, 22.6, sk3.hDC, 44, 0, SRCCOPY ' Pause Button
      
      BitBlt mstop.hDC, -2, 0, 24, 22, sk3.hDC, 66, 0, SRCCOPY ' Stop Button
      
      BitBlt mNext.hDC, -2, 0, 27.6, 22.6, sk3.hDC, 88, 0, SRCCOPY ' Next Button
      
      BitBlt mcd.hDC, -1, 0, 22, 16.2, sk3.hDC, 114, 0, SRCCOPY ' Eject CD Button
      
      
      BitBlt mAbout.hDC, 0, 0, 9, 9, Sk1.hDC, 0, 0, SRCCOPY 'About button
      
      BitBlt mMinsize.hDC, 0, 0, 9, 9, Sk1.hDC, 0, 18, SRCCOPY 'Minsize Button
      
      BitBlt mExit.hDC, 0, 0, 9, 9, Sk1.hDC, 18, 0, SRCCOPY 'Exit Button
      
      BitBlt mMin.hDC, 0, 0, 9, 9, Sk1.hDC, 9, 0, SRCCOPY 'Minsize Button2
      
      BitBlt shuff.hDC, 0, 0, 46, 15.2, sk7.hDC, 29, 0, SRCCOPY ' Shufrep Button
      
      BitBlt mbut.hDC, 0, 0, 28, 15.2, sk7.hDC, 0, 0, SRCCOPY
      
      BitBlt mEq.hDC, -1, 0, 23, 10, sk7.hDC, 0, 62, SRCCOPY
      
      BitBlt mpl.hDC, 0, 0, 23, 10, sk7.hDC, 23, 62, SRCCOPY
      
      BitBlt mVolume.hDC, 0, 0, 68, 14.8, sk8.hDC, 0, 0, SRCCOPY
      
      BitBlt mVolPos.hDC, 0, 0, 12, 11, sk8.hDC, 17, 423, SRCCOPY
         
      BitBlt titlebar.hDC, -32, 0, 307, 16, Sk1.hDC, -5, 0, SRCCOPY ' Titlebar
      BitBlt playList.hDC, -32, 0, 307, 16, Sk1.hDC, -5, 0, SRCCOPY
      
      BitBlt Base.hDC, 0, -1.2, Base_Width, Base_Height, sk2.hDC, 0, 0, SRCCOPY
      BitBlt ProgBar.hDC, 0, 0, 248, 10, sk4.hDC, 0, 0, SRCCOPY
      BitBlt PosBar.hDC, 0, 0, 28, 8, sk4.hDC, 249, 0, SRCCOPY
      
      BitBlt mMono.hDC, 0, 0, 29, 12, sk5.hDC, 29, 12, SRCCOPY ' Mono Off
      BitBlt mStereo.hDC, 0, 0, 29, 12, sk5.hDC, 0, 12, SRCCOPY ' Stereo Off
      
      ' Number Skins
      
      BitBlt num1.hDC, 0, 0, 10, 13, sk6.hDC, 0, 0, SRCCOPY
      BitBlt num2.hDC, 0, 0, 10, 13, sk6.hDC, 0, 0, SRCCOPY
      BitBlt num3.hDC, 0, 0, 10, 13, sk6.hDC, 0, 0, SRCCOPY
      BitBlt num4.hDC, 0, 0, 10, 13, sk6.hDC, 0, 0, SRCCOPY
      
      ' Text Display Skin

      
      
      
      
      
      'BitBlt Base.hDC, -6, 88, 136, 17, sk3.hDC, -22, 0, SRCCOPY
   
      
      
      
      ' Base buttons Play,Pause,Stop etc
      
      mBack.Refresh
      mPlay.Refresh
      mpause.Refresh
      mstop.Refresh
      mNext.Refresh
      mcd.Refresh
      shuff.Refresh
      mbut.Refresh
      mEq.Refresh
      mpl.Refresh
      mVolume.Refresh
      mVolPos.Refresh
          
      
      ' Toolbar Buttons
      
      mAbout.Refresh
      mMinsize.Refresh
      mExit.Refresh
      mMin.Refresh
      
      
      
      Base.Refresh
      ProgBar.Refresh
      playList.Refresh
      
      PosBar.Refresh
      mMono.Refresh
      mStereo.Refresh
      
      
      ' Numbers
      
      num1.Refresh
      num2.Refresh
      num3.Refresh
      num4.Refresh
      

      
    For k = 0 To 5
        minButtons(k).BackStyle = 0
    Next
        

      
      
End Sub

Private Sub mAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mAbout.hDC, 0, 0, 9, 9, Sk1.hDC, 0, 9, SRCCOPY 'About button
    mAbout.Refresh
    
End Sub

Private Sub mAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mAbout.hDC, 0, 0, 9, 9, Sk1.hDC, 0, 0, SRCCOPY 'About button
    mAbout.Refresh
    
End Sub

Private Sub mBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mBack.hDC, 0, 0, 22, 22, sk3.hDC, 0, 18, SRCCOPY
    mBack.Refresh
    Base.Refresh
    
End Sub

Private Sub mBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        BitBlt mBack.hDC, 0, 0, 22, 22, sk3.hDC, 0, 0, SRCCOPY
        mBack.Refresh
        Base.Refresh
        
End Sub

Private Sub mbut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mbut.hDC, 0, 0, 28, 15.2, sk7.hDC, 0, 15.2, SRCCOPY
    mbut.Refresh
    
End Sub

Private Sub mbut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mbut.hDC, 0, 0, 28, 15.2, sk7.hDC, 0, 0, SRCCOPY
    mbut.Refresh
    
End Sub

Private Sub mcd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mcd.hDC, -1, 0, 22, 22.6, sk3.hDC, 114, 16, SRCCOPY ' Eject CD Button
    mcd.Refresh
    l = True
     
End Sub

Private Sub mcd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mcd.hDC, -1, 0, 22, 22.6, sk3.hDC, 114, 0, SRCCOPY ' Eject CD Button
     mcd.Refresh
     
End Sub

Private Sub mEq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mEq.hDC, -1, 0, 23, 10, sk7.hDC, 46, 62, SRCCOPY ' Down
    mEq.Refresh
    
End Sub

Private Sub mEq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mEq.hDC, -1, 0, 23, 10, sk7.hDC, 0, 62, SRCCOPY
    mEq.Refresh
    
End Sub

Private Sub mExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mExit.hDC, 0, 0, 9, 9, Sk1.hDC, 18, 9, SRCCOPY 'Exit Button
    mExit.Refresh
    
End Sub

Private Sub mExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mExit.hDC, 0, 0, 9, 9, Sk1.hDC, 18, 0, SRCCOPY 'Exit Button
    mExit.Refresh
    End
    
End Sub

Private Sub minButtons_Click(Index As Integer)
    If UpDown Then
        Select Case Index
            Case 0
            ' Back Button
            Case 1
            ' Play Button
            Case 2
            ' Pause Button
            Case 3
            ' Stop Button
            Case 4
            ' Forward Button
            Case 5
            ' CD Button
        End Select
    End If
        
End Sub

Private Sub minButtons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        ReleaseCapture
        SendMessage Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 1
    End If
    
End Sub

Private Sub mMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mMin.hDC, 0, 0, 9, 9, Sk1.hDC, 9, 18, SRCCOPY 'Minsize Button2
    mMin.Refresh
    
End Sub

Private Sub mMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mMin.hDC, 0, 0, 9, 9, Sk1.hDC, 9, 0, SRCCOPY 'Minsize Button2
    mMin.Refresh
    
End Sub

Private Sub mMinsize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mMinsize.hDC, 0, 0, 9, 9, Sk1.hDC, 9, 18, SRCCOPY 'Minsize Button
    mMinsize.Refresh
    MinUpDown
    
End Sub

Private Sub mMinsize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mMinsize.hDC, 0, 0, 9, 9, Sk1.hDC, 0, 18, SRCCOPY 'Minsize Button
    mMinsize.Refresh
    
End Sub

Private Sub mNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mNext.hDC, -2, 0, 25.2, 22, sk3.hDC, 88, 18, SRCCOPY ' Next Button
    mNext.Refresh
    
End Sub

Private Sub mNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mNext.hDC, -2, 0, 25.6, 22, sk3.hDC, 88, 0, SRCCOPY ' Next Button
    mNext.Refresh
    
End Sub

Private Sub mpause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mpause.hDC, -1, 0, 22.6, 22.6, sk3.hDC, 44, 18, SRCCOPY ' Pause Button
    mpause.Refresh
    
End Sub

Private Sub mpause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mpause.hDC, -1, 0, 22.6, 22.6, sk3.hDC, 44, 0, SRCCOPY ' Pause Button
    mpause.Refresh
    
End Sub


    


Private Sub mpl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mpl.hDC, 0, 0, 23, 10, sk7.hDC, 69, 62, SRCCOPY
    mpl.Refresh
    
End Sub

Private Sub mpl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mpl.hDC, 0, 0, 23, 10, sk7.hDC, 23, 62, SRCCOPY
    mpl.Refresh
    ShowPlayList
    
End Sub

Private Sub mPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mPlay.hDC, 0, 0, 22, 22, sk3.hDC, 22, 18, SRCCOPY
    mPlay.Refresh
    Base.Refresh
    
End Sub

Private Sub mPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mPlay.hDC, 0, 0, 22, 22, sk3.hDC, 22, 0, SRCCOPY ' Play Button
    mPlay.Refresh
    
End Sub

Private Sub mstop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mstop.hDC, -2, 0, 24, 22, sk3.hDC, 66, 18, SRCCOPY ' Stop Button
    mstop.Refresh
    
End Sub

Private Sub mstop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mstop.hDC, -2, 0, 24, 22, sk3.hDC, 66, 0, SRCCOPY ' Stop Button
    mstop.Refresh
    
End Sub

Private Sub mVolPos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mVolPos.hDC, 0, 0, 12, 11, sk8.hDC, 1.2, 423, SRCCOPY
    X = 0
    Y = 0
    
    If SlideFlag1 = False Then
        x1 = X: x3 = mVolPos.Left
        x2 = Screen.TwipsPerPixelX
        SlideFlag1 = True
        
    End If
    
    
    
    mVolPos.Refresh
    
End Sub

Private Sub mVolPos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim A As Integer
    
    If SlideFlag1 = True Then
        pos = x3 + (X - x1) / x2
        If pos < 3 Then pos = 3
        If pos > 55 Then pos = 55
        x3 = pos: mVolPos.Left = pos
        A = Format(pos, "#") / 2.1
        Label2 = "Volume: " & A * 4 - 4 & "%"
        BitBlt mVolume.hDC, 0, 0, 68, 14.8, sk8.hDC, 0, A * 15, SRCCOPY
        mVolume.Refresh
    
    End If
    
End Sub

Private Sub mVolPos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt mVolPos.hDC, 0, 0, 12, 11, sk8.hDC, 17, 423, SRCCOPY
    mVolPos.Refresh
    SlideFlag1 = False
    
End Sub

Private Sub PosBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PosBar.hDC, 0, 0, 28, 8, sk4.hDC, 280, 0, SRCCOPY
    PosBar.Refresh
    X = 0
    Y = 0
    If SlideFlag = False Then
        IX = X: FX = PosBar.Left
        TX = Screen.TwipsPerPixelX
        SlideFlag = True
    End If
    
    
End Sub

Private Sub PosBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If SlideFlag = True Then
        pos = FX + (X - IX) / TX
        If pos < 1 Then pos = 1
        If pos > 222 Then pos = 222
        FX = pos: PosBar.Left = pos
    End If
    
End Sub

Private Sub PosBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt PosBar.hDC, 0, 0, 28, 8, sk4.hDC, 249, 0, SRCCOPY
    PosBar.Refresh
    SlideFlag = False
    
End Sub

Private Sub shuff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt shuff.hDC, 0, 0, 46, 15, sk7.hDC, 29, 15.2, SRCCOPY ' Shufrep Button
    shuff.Refresh
    
End Sub

Private Sub shuff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt shuff.hDC, 0, 0, 46, 15.2, sk7.hDC, 29, 0, SRCCOPY ' Shufrep Button
    shuff.Refresh
    
End Sub

Private Sub Timer1_Timer()
    mTimer = mTimer + 1
    If mTimer = 10 Then
        mTimer = 0
        n2 = n2 + 1
    ElseIf n2 = 10 Then
        n2 = 0
        n3 = n3 + 1
    ElseIf n3 = 10 Then
        n3 = 0
        n4 = n4 + 1
    ElseIf n4 = 10 Then
        n4 = 0
        
    End If
    
    BitBlt num1.hDC, 0, 0, 10, 13, sk6.hDC, mTimer * 9, 0, SRCCOPY
    BitBlt num2.hDC, 0, 0, 10, 13, sk6.hDC, n2 * 9, 0, SRCCOPY
    BitBlt num3.hDC, 0, 0, 10, 13, sk6.hDC, n3 * 9, 0, SRCCOPY
    BitBlt num4.hDC, 0, 0, 10, 13, sk6.hDC, n4 * 9, 0, SRCCOPY
    
    num1.Refresh
    num2.Refresh
    num3.Refresh
    num4.Refresh

End Sub



Private Sub titlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        ReleaseCapture
        SendMessage Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 1
    End If


End Sub
