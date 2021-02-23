VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Gabut 
   BorderStyle     =   0  'None
   ClientHeight    =   9435
   ClientLeft      =   405
   ClientTop       =   1845
   ClientWidth     =   18480
   Icon            =   "Ollie.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   18480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8400
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8280
      Top             =   2160
   End
   Begin VB.Image Image1 
      Height          =   4950
      Left            =   8160
      Picture         =   "Ollie.frx":0CCA
      Top             =   -120
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   12015
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   21193
      _cy             =   1508
   End
End
Attribute VB_Name = "Gabut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000

Public HP As String


Private Sub Form_Activate()
    Me.BackColor = vbBlue
    Trans 1
End Sub

Private Sub Trans(Level As Integer)
    Dim Msg As Long
    Msg = GetWindowLong(Me.hwnd, G)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, G, Msg
    SetLayeredWindowAttributes Me.hwnd, vbBlue, Level, LWA_COLORKEY
End Sub

Private Sub Form_Load()
WindowsMediaPlayer1.settings.volume = 100
WindowsMediaPlayer1.URL = (App.Path & "\" & "Ollie_Laughter.mp3")
Image1.Picture = LoadPicture(App.Path & "\" & "ollie_kanan.gif")

HP = 1

Image1.Height = 6135
Image1.Left = 120
Image1.Top = 1000

End Sub

Private Sub Image1_Click()
HP = HP - 1
If HP = 0 Then
Unload Me
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lokasi As Integer
Randomize
lokasi = Int((5 * Rnd) + 1)

Select Case lokasi
Case 1
    Image1.Top = 1000
    Image1.Left = 120
    Timer1.Enabled = True
    Timer2.Enabled = False
    Image1.Picture = LoadPicture(App.Path & "\" & "ollie_kanan.gif")
Case 2
    Image1.Top = 4000
    Image1.Left = 120
    Timer1.Enabled = True
    Timer2.Enabled = False
    Image1.Picture = LoadPicture(App.Path & "\" & "ollie_kanan.gif")
Case 3
    Image1.Top = 1000
    Image1.Left = 16200
    Timer2.Enabled = True
    Timer1.Enabled = False
    Image1.Picture = LoadPicture(App.Path & "\" & "ollie_kiri.gif")
Case 4
    Image1.Top = 4000
    Image1.Left = 16200
    Timer2.Enabled = True
    Timer1.Enabled = False
    Image1.Picture = LoadPicture(App.Path & "\" & "ollie_kiri.gif")
Case 5
    Image1.Top = 1000
    Image1.Left = 8840
Case Else
    Image1.Top = 4000
    Image1.Left = 8840
End Select



End Sub

Private Sub Timer1_Timer()
Image1.Left = Image1.Left + 80
If Image1.Left = 16200 Then
Timer2.Enabled = True
Timer1.Enabled = False
Image1.Picture = LoadPicture(App.Path & "\" & "ollie_kiri.gif")
End If
End Sub

Private Sub Timer2_Timer()
Image1.Left = Image1.Left - 80
If Image1.Left = 120 Then
Timer1.Enabled = True
Timer2.Enabled = False
Image1.Picture = LoadPicture(App.Path & "\" & "ollie_kanan.gif")
End If
End Sub

