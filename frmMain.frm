VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "GhostYourself - By Jesse Seidel"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Status of GhostYourself"
      Top             =   4740
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Welcome to GhostYourself!"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   5235
      ScaleHeight     =   3885
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   0
      Width           =   735
      Begin MSComDlg.CommonDialog cd1 
         Left            =   0
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ComctlLib.Slider barAmount 
         Height          =   2235
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "Change transparency level"
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3942
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   51
         Max             =   255
         SelStart        =   180
         TickStyle       =   2
         TickFrequency   =   51
         Value           =   180
      End
      Begin VB.Label Label1 
         Caption         =   "180"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   5970
      TabIndex        =   3
      Top             =   3885
      Width           =   5970
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Start webcam"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "Stop webcam"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save As"
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         ToolTipText     =   "Save image"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Set BG"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         ToolTipText     =   "Set the background image to current image"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Timer tmrMain 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   3360
         Top             =   240
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   4440
         X2              =   4440
         Y1              =   240
         Y2              =   730
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   240
         Y2              =   730
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   4440
         X2              =   4440
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   2760
         X2              =   2760
         Y1              =   240
         Y2              =   720
      End
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   120
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   323
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4845
      Begin VB.Image Image1 
         Height          =   15
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   120
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   323
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4845
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   3600
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GhostYourself by Jesse Seidel
'Please do not steal my code without including me somewhere in your program
'I will get pissed off and you never know what happens then :@

Option Explicit

'This API is the key to the translucency
Private Declare Function AlphaBlend Lib "msimg32" ( _
ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
Destination As Any, Source As Any, ByVal Length As Long)

'For changing transparency levels
Private Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type

Private Sub barAmount_Scroll()
    Dim tProperties As typeBlendProperties
    Dim lngBlend As Long
    picDestination.Cls 'Clear the destination picturebox
    tProperties.tBlendAmount = 255 - barAmount 'Change transparency level
    CopyMemory lngBlend, tProperties, 4
    AlphaBlend picDestination.hDC, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, _
    picSource.hDC, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, lngBlend 'Fade image
    picDestination.Refresh 'Refresh to display
    Label1.Caption = barAmount.Value 'Change the label to value of transparency
End Sub

Private Sub Command1_Click()
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 320, 240, Me.hWnd, 0) 'Get hWnd for webcam so we can use it
DoEvents: SendMessage mCapHwnd, CONNECT, 0, 0 'Capture from webcam
tmrMain.Enabled = True 'Enable timer to refresh webcam images
Command2.Enabled = True 'Make stop button enabled
Command1.Enabled = False 'Make start button disabled
Command4.Enabled = True 'Make Set BG button enabled
sb1.SimpleText = "Webcam started..." 'Change statusbar caption
End Sub

Private Sub Command2_Click()
tmrMain.Enabled = False 'Disable refreshing of webcam images
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0 'Stop capturing of images from webcam
Command1.Enabled = True 'Make start enabled
Command2.Enabled = False 'Make stop disabled
Command4.Enabled = False 'Make Set BG button disabled
sb1.SimpleText = "Webcam stopped..." 'Change statusbar caption
End Sub

Private Sub Command3_Click()
Image1.Picture = picDestination.Image 'Capture current image for saving
cd1.Filter = "Bitmap (*.BMP)|*.BMP|Jpeg (*.JPEG)|*.JPG|Gif (*.GIF)|*.GIF" 'Supported file-types
cd1.ShowSave 'Show save dialog
SavePicture Image1, cd1.FileName 'Write picture to hard-drive
sb1.SimpleText = "Image saved to " & cd1.FileName 'Change statusbar caption
End Sub

Private Sub Command4_Click()
picSource.Picture = picDestination.Image 'Set background image
sb1.SimpleText = "Background set..." 'Change statusbar caption
End Sub

Private Sub Form_Load()
barAmount_Scroll 'This is just for my own purposes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0 'Stop webcam capturing
End Sub

Private Sub tmrMain_Timer()
On Error Resume Next
SendMessage mCapHwnd, GET_FRAME, 0, 0 'Capture frame from webcam
SendMessage mCapHwnd, COPY, 0, 0 'Copy frame
picDestination.Picture = Clipboard.GetData 'Paste captured frame from clipboard
Clipboard.Clear 'Clear clipboard
barAmount_Scroll 'Change alpha-blending and such
picSource.Height = picDestination.Height 'Make sure both the source and destination pictures are the same height/width
picSource.Width = picDestination.Width 'Make sure both the source and destination pictures are the same height/width
End Sub
