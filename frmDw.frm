VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Begin VB.Form frmDw 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "YouTubeTool-Beta Release"
   ClientHeight    =   11610
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   16035
   Icon            =   "frmDw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11610
   ScaleWidth      =   16035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPR 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "YouTube Preview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   5160
      TabIndex        =   3
      Top             =   6720
      Width           =   6735
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF 
         Height          =   3615
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   6495
         _cx             =   11456
         _cy             =   6376
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   "-1"
         Loop            =   "-1"
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   "-1"
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "ShowAll"
         DeviceFont      =   "0"
         EmbedMovie      =   "0"
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   "1"
         Profile         =   "0"
         ProfileAddress  =   ""
         ProfilePort     =   "0"
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click and wait for video to Load..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   6495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6480
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   2295
      End
   End
   Begin VB.Frame frmDW 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   8160
      TabIndex        =   1
      Top             =   2280
      Width           =   7095
      Begin VB.CommandButton cmdPR 
         BackColor       =   &H00808080&
         Caption         =   "Preview"
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   615
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1320
         Width           =   3735
      End
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   2055
         TabIndex        =   21
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtWatchUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "http://www.youtube.com/watch?v=SnooUEuyn_M"
         Top             =   360
         Width           =   6255
      End
      Begin VB.CommandButton cmdG 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Get Direct Link"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6120
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton getFlv 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Get FLV"
         Enabled         =   0   'False
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
         Left            =   5760
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton getHDH 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Get HD Low"
         Enabled         =   0   'False
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
         Left            =   4440
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton getHDL 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Get HD High"
         Enabled         =   0   'False
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
         Left            =   1440
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton getMp4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Get MP4"
         Enabled         =   0   'False
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
         MouseIcon       =   "frmDw.frx":87D9
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3120
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar pbDownload 
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   2760
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblDur 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblCat 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label lblRat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rating"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6120
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblRating 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   6120
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblPublished 
         BackStyle       =   0  'Transparent
         Caption         =   "Published"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblTitle1 
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter YouTube video URL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label cmdCLose 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.ListBox lstID 
      Height          =   255
      ItemData        =   "frmDw.frx":8AE3
      Left            =   13680
      List            =   "frmDw.frx":8AE5
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   14280
      Top             =   1440
   End
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   13560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   11880
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   14400
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   11160
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   7935
      Left            =   0
      Picture         =   "frmDw.frx":8AE7
      ScaleHeight     =   7875
      ScaleWidth      =   7995
      TabIndex        =   6
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton cmdDW 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Downloader"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4680
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1BD76&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   600
         TabIndex        =   13
         Top             =   840
         Width           =   5055
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00808080&
         Caption         =   "Search"
         Height          =   615
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox lstTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1BD76&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2370
         ItemData        =   "frmDw.frx":CCD3
         Left            =   600
         List            =   "frmDw.frx":CCD5
         TabIndex        =   11
         ToolTipText     =   "Double Click to Download"
         Top             =   1680
         Width           =   6135
      End
      Begin MSComctlLib.ProgressBar pgSt 
         Height          =   210
         Left            =   3240
         TabIndex        =   33
         ToolTipText     =   "Common Status"
         Top             =   4465
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Max             =   12
         Scrolling       =   1
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D6B05A&
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "By Rajendra Khope"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   4440
         Width           =   6975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Search Query"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00ECA646&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         MouseIcon       =   "frmDw.frx":CCD7
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Don't worry abt this extra Inet..."
      Height          =   495
      Left            =   11160
      TabIndex        =   41
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      MouseIcon       =   "frmDw.frx":CFE1
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
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
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "frmDw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim urlMP4 As String
Dim urlFLV_VHigh As String
Dim urlFLV_High As String
Dim urlFLV_Low As String

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub cmdDW_Click()
    frmDW.Left = 240
    frmDW.Top = 500
    'frmSR.Left = 8000
End Sub
Private Sub cmdPR_Click()
 Call InitCommonControls
        frmPR.Left = 300
        frmPR.Top = 450
        LoadFLVPlayer URLDecode(urlFLV_High)
End Sub
Private Sub Command2_Click()
'Not in Use ... for further use...testing
   Dim J As Long
    
    File1.Pattern = "*.jpg"
    File1.Path = App.Path & "\Cache\"

    ImageList1.ListImages.Clear
    ImageList1.ImageHeight = 50
    ImageList1.ImageWidth = 50
    For J = 0 To 10
        'ImageList1.ListImages.Add , "img" & J, LoadPicture(File1.Path & "\" & File1.List(J)) '
        ImageList1.ListImages.Add , "img" & J, LoadPicture(File1.Path & "\default.jpg")
    Next
    ListView1.View = lvwIcon
    ListView1.Icons = ImageList1
    For J = 0 To 10
        ListView1.ListItems.Add , , File1.List(J), "img" & J
    Next
End Sub
Private Sub cmdCLose_Click()
    frmDW.Left = 8000
End Sub
Private Sub cmdG_Click()
Dim StrRsponse As String
    cmdG.Enabled = False
    picTitle.Picture = LoadPicture(App.Path & "\loading.jpg")
    StrRsponse = LoadPage(txtWatchUrl.Text)
    
    GetURLs StrRsponse
    GetAllInfo StrRsponse
    
    GetTitleImage Inet2, vInfo.vThumbURL, vInfo.vID & ".jpg"
    picTitle.Picture = LoadPicture(App.Path & "\Cache\" & vInfo.vID & ".jpg")
        
    lblTitle1.Caption = vInfo.vTitle
    txtDesc.Text = vInfo.vDesciption
    lblCat.Caption = vInfo.vCategory
    lblAuthor.Caption = vInfo.vAuthor
    lblDur.Caption = vInfo.vDuration & " Sec."
    lblPublished.Caption = vInfo.vPub
    lblRating.Caption = vInfo.vRatings
End Sub
Function LoadPage(strURL As String)
On Error GoTo ErroControl
    Dim strGarb As String
    Dim TempBuff As String
    Dim ContentLength As String
    
    strGarb = Inet1.OpenURL(strURL)
    
    Do While Inet1.StillExecuting
        DoEvents
    Loop
    
    'ContentLength = Inet1.GetHeader
    'MsgBox ContentLength
    strGarb = strGarb & Inet1.GetChunk(1024, icString)
    Do
        DoEvents
        TempBuff = Inet1.GetChunk(1024, icString)
        If Len(TempBuff) = 0 Then Exit Do
        strGarb = strGarb & TempBuff
        lblTitle.Caption = "Reading Data..."
    Loop
    
    lblTitle.Caption = "Complete!"
    pgSt.Value = 12
    LoadPage = strGarb
    
    Exit Function
ErroControl:
    If InStr(1, Err.Description, "12002") Then
        lblTitle.Caption = "Slow Connection"
    End If
    MsgBox "Something Went Wrong" & vbCrLf & Err.Description
End Function
Function GetURLs(strHypertext As String)
'Debug.Print ""
    Dim str3 As String
    Dim strArr() As String
    Dim I, intr, urlCode As Integer
    'Dim urlMP4, urlFLV_VHigh, urlFLV_High, urlFLV_Low As String
    
    str3 = UrlParser(strHypertext)
    str3 = URLDecode(str3)
    strArr = Split(str3, ",")

    For I = 0 To UBound(strArr)
        intr = InStr(strArr(I), "|")
        urlCode = Mid(strArr(I), 1, intr - 1)
        'MsgBox urlCode
        
        Select Case urlCode
            Case "22"
                urlMP4 = Mid(strArr(I), intr + 1, Len(strArr(I)))
                getMp4.Enabled = True
            Case "34"
                urlFLV_VHigh = Mid(strArr(I), intr + 1, Len(strArr(I)))
                getHDH.Enabled = True
            Case "35"
                urlFLV_High = Mid(strArr(I), intr + 1, Len(strArr(I)))
                getHDL.Enabled = True
            Case "5"
                urlFLV_Low = Mid(strArr(I), intr + 1, Len(strArr(I)))
                getFlv.Enabled = True
        End Select
        'Text2.Text = Text2.Text & urlOK & vbCrLf & vbCrLf
    Next
End Function
Function URLDecode(str)
Dim st, sR As String
Dim I As Integer

        str = Replace(str, "+", " ")
        For I = 1 To Len(str)
            st = Mid(str, I, 1)
            If st = "%" Then
                If I + 2 < Len(str) Then
                    sR = sR & _
                        Chr(CLng("&H" & Mid(str, I + 1, 2)))
                    I = I + 2
                End If
            Else
                sR = sR & st
            End If
        Next
        URLDecode = sR
    End Function
Private Sub cmdSearch_Click()
If txtSearch.Text <> "" Then
    Dim StrRsponse As String
        lstTitle.Clear
        lstTitle.AddItem "Loading.."
    cmdSearch.Enabled = False
    StrRsponse = LoadPage("http://www.youtube.com/results?search_query=" & txtSearch.Text)
    
    SearchEngine StrRsponse, lstTitle, lstID
    cmdSearch.Enabled = True
Else
    MsgBox "Enter Search Query"
End If
End Sub
Private Sub Form_Initialize()
    Label9.Caption = App.ProductName & " " & "v" & App.Major & "." & App.Minor & "." & App.Revision
    Call InitCommonControls
    SetBG Picture1, Me
End Sub
Private Sub Form_Load()
    Me.Width = 7170
    Me.Height = 4845
End Sub
Function LoadFLVPlayer(flvLink)
    SWF.Movie = App.Path + "\flvplayer.swf"

        Call SWF.SetVariable("file", flvLink)
        Call SWF.Play
        SWF.ToolTipText = "Preview of " & vInfo.vTitle
        SWF.WMode = "Fullscreen"
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub getHDH_Click()
    Dim a As Integer
    a = MsgBox("Do you want Download or Play this video?" & vbCrLf & "Click yes to Download, No to Play", vbYesNo, "Download/Play")
    If a = 6 Then
        DWFile URLDecode(urlFLV_VHigh), vInfo.vTitle & "_Full_HD.flv", pbDownload, Label2, Inet1
    Else
         Call InitCommonControls
        frmPR.Left = 240
        frmPR.Top = 120
        LoadFLVPlayer URLDecode(urlFLV_VHigh)
    End If
End Sub
Private Sub getHDL_Click()
    Dim a As Integer
    a = MsgBox("Do you want Download or Play this video?" & vbCrLf & "Click yes to Download, No to Play", vbYesNo, "Download/Play")
    If a = 6 Then
        DWFile URLDecode(urlFLV_High), vInfo.vTitle & "_HD.flv", pbDownload, Label2, Inet1
    Else
         Call InitCommonControls
        frmPR.Left = 240
        frmPR.Top = 120
        LoadFLVPlayer URLDecode(urlFLV_High)
    End If
End Sub
Private Sub getMp4_Click()
    DWFile URLDecode(urlMP4), vInfo.vTitle & ".mp4", pbDownload, Label2, Inet1
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
    pgSt.Value = State
Select Case State
    Case icResolvingHost                                        ' 1
         lblTitle.Caption = "Looking up IP address"
    Case icHostResolved                                         ' 2
        lblTitle.Caption = "IP address found"
    Case icConnecting                                           ' 3
        lblTitle.Caption = "Connecting Host"
    Case icConnected                                            ' 4
        lblTitle.Caption = "Connected."
    Case icRequesting                                           ' 5
        lblTitle.Caption = "Making Request..."
    Case icRequestSent                                          ' 6
        lblTitle.Caption = "Request sent."
    Case icReceivingResponse                                    ' 7
        lblTitle.Caption = "Receiving Response.."
    Case icResponseReceived                                     ' 8
        lblTitle.Caption = "Response received."
    Case icDisconnecting                                        ' 9
        lblTitle.Caption = "Disconnecting..."
    Case icDisconnected                                          ' 10
        lblTitle.Caption = "Disconnected."
    Case icError ' 11
        lblTitle.Caption = "Error " & Inet1.ResponseCode & " " & Inet1.ResponseInfo
        
        Exit Sub
    Case icResponseCompleted                                      ' 12
End Select
End Sub
Private Sub getFlv_Click()
    Dim a As Integer
    a = MsgBox("Do you want Download or Play this video?" & vbCrLf & "Click yes to Download, No to Play", vbYesNo, "Download/Play")
    If a = 6 Then
        DWFile URLDecode(urlFLV_Low), vInfo.vTitle & ".flv", pbDownload, Label2, Inet1
    Else
         Call InitCommonControls
        frmPR.Left = 240
        frmPR.Top = 120
        LoadFLVPlayer URLDecode(urlFLV_Low)
    End If
End Sub
Private Sub Label6_Click()
    frmPR.Left = 8000
End Sub
Private Sub Label7_Click()
    Inet1.Cancel
    End
End Sub
Private Sub Label9_Click()
    frmPR.Left = 8000
End Sub
Private Sub lstTitle_DblClick()
On erro GoTo errorHandler
    If txtWatchUrl.Text <> "" Then
        txtWatchUrl.Text = "http://www.youtube.com/watch?v=" & lstID.List(lstTitle.ListIndex)
        cmdG_Click
        cmdDW_Click
    Else

    End If
    Exit Sub
errorHandler:
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub txtSearch_Change()
    cmdSearch.Enabled = True
End Sub
Private Sub txtWatchUrl_Change()
    cmdG.Enabled = True
End Sub
Private Sub txtWatchUrl_Click()
    txtWatchUrl.SelStart = 0
    txtWatchUrl.SelLength = Len(txtWatchUrl.Text)
   ' txtWatchUrl.Text = Clipboard.GetText
End Sub
