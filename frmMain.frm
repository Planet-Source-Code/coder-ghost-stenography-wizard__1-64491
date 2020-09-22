VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stenography Wizard - Final Stand Productions"
   ClientHeight    =   5115
   ClientLeft      =   1095
   ClientTop       =   1410
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
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
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBars 
      Interval        =   100
      Left            =   720
      Top             =   4680
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2005 - 2006, Final Stand Productions"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label lbCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Stenography Wizard"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.CommandButton btCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton btBack 
      Caption         =   "<< &Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton btFinish 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton btNext 
      Caption         =   "&Next >>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   3605
      Index           =   6
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   7935
      TabIndex        =   88
      Top             =   960
      Width           =   7935
      Begin VB.CommandButton btSave6 
         Caption         =   "&Save Data"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   95
         Top             =   3120
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar Bar6 
         Height          =   375
         Left            =   240
         TabIndex        =   90
         Top             =   2640
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton btStop6 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   89
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Decoding Progress - "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Top             =   2400
         Width           =   7455
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0442
         Height          =   975
         Left            =   120
         TabIndex        =   93
         Top             =   360
         Width           =   7575
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Decode a stenographic image - Step 3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label lbStatus6 
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting..."
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   3120
         Width           =   5895
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   3605
      Index           =   2
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   7935
      TabIndex        =   15
      Top             =   960
      Width           =   7935
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   71
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox Picture6 
         Height          =   255
         Left            =   7320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   69
         Top             =   1440
         Width           =   255
         Begin VB.PictureBox picTmp2 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   255
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   70
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.TextBox txtSource2 
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1920
         Width           =   7575
      End
      Begin VB.CommandButton btBrowse2 
         Caption         =   "Browse ..."
         Height          =   375
         Left            =   6480
         TabIndex        =   61
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton btLoad2 
         Caption         =   "&Load Image"
         Height          =   375
         Left            =   5040
         TabIndex        =   60
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Size: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Storage:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label lbSize2 
         BackStyle       =   0  'Transparent
         Caption         =   "0 x 0"
         Height          =   255
         Left            =   2280
         TabIndex        =   66
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lbMax2 
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":05C9
         Height          =   975
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Select an image to decode:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Decode a stenographic image -  Step 1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   3605
      Index           =   1
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   7935
      TabIndex        =   8
      Top             =   960
      Width           =   7935
      Begin VB.CommandButton btLoad1 
         Caption         =   "&Load Image"
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton btBrowse1 
         Caption         =   "&B&rowse ..."
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtSource1 
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1920
         Width           =   7695
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   7560
         ScaleHeight     =   495
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   3240
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   255
         Left            =   7560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   21
         Top             =   3240
         Width           =   255
         Begin VB.PictureBox picTmp1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   1095
            Left            =   0
            ScaleHeight     =   69
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   69
            TabIndex        =   22
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label lbRec1 
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lbMax1 
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lbSize1 
         BackStyle       =   0  'Transparent
         Caption         =   "0 x 0"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Recommended Storage:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Storage:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Size: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Please select an image now:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   7335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":06C4
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Encode a stenographic image - Step 1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   3605
      Index           =   0
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   7935
      TabIndex        =   7
      Top             =   960
      Width           =   7935
      Begin VB.OptionButton optMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Decode an attachment from a stenographic image. "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   2280
         Width           =   7095
      End
      Begin VB.OptionButton optMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create a stenographic image from a source image and file."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   7095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0838
         Height          =   735
         Left            =   720
         TabIndex        =   13
         Top             =   2520
         Width           =   6855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":08F1
         Height          =   735
         Left            =   720
         TabIndex        =   12
         Top             =   1320
         Width           =   6855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":09CB
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   3605
      Index           =   3
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   7935
      TabIndex        =   31
      Top             =   960
      Width           =   7935
      Begin VB.CommandButton btLoad3 
         Caption         =   "&Prepare File"
         Height          =   375
         Left            =   5040
         TabIndex        =   37
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton btBrowse3 
         Caption         =   "Browse ..."
         Height          =   375
         Left            =   6480
         TabIndex        =   35
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtSource3 
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1920
         Width           =   7575
      End
      Begin VB.Label lbPixel3 
         BackStyle       =   0  'Transparent
         Caption         =   "0 pixels"
         Height          =   255
         Left            =   6480
         TabIndex        =   45
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Per Character (Affects Quality):"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   44
         Top             =   3240
         Width           =   4815
      End
      Begin VB.Label lbSize3 
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         Height          =   255
         Left            =   2160
         TabIndex        =   43
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label lbMax3 
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lbRec3 
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         Height          =   255
         Left            =   2160
         TabIndex        =   41
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum File Size:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Recommened File Size: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected File Size: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a file to attach:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Encode a stenographic image - Step 2"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0AE2
         Height          =   975
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   3605
      Index           =   4
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   7935
      TabIndex        =   46
      Top             =   960
      Width           =   7935
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   59
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         Height          =   255
         Left            =   7560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   57
         Top             =   120
         Width           =   255
         Begin VB.PictureBox picTmp4 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   23
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   23
            TabIndex        =   58
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton btStop4 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   56
         Top             =   3120
         Width           =   1455
      End
      Begin VB.PictureBox picQuality 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         ScaleHeight     =   825
         ScaleWidth      =   7425
         TabIndex        =   47
         Top             =   1320
         Width           =   7455
         Begin VB.Label lbCaption4 
            BackStyle       =   0  'Transparent
            Caption         =   "High Quality (10+)"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   1440
            TabIndex        =   50
            Top             =   0
            Width           =   3615
         End
         Begin VB.Label lbDesc4 
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   7215
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Output Quality: "
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   0
            Width           =   2055
         End
      End
      Begin MSComctlLib.ProgressBar Bar4 
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   2640
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbStatus4 
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting..."
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   3120
         Width           =   5895
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Encode a stenographic image - Step 3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0C40
         Height          =   855
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   7575
      End
      Begin VB.Label lbProgress4 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Generation Progress - "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   2400
         Width           =   7455
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   3605
      Index           =   5
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   7935
      TabIndex        =   72
      Top             =   960
      Width           =   7935
      Begin VB.CommandButton btLoad5 
         Caption         =   "&Load Image"
         Height          =   375
         Left            =   5040
         TabIndex        =   78
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton btBrowse5 
         Caption         =   "Browse ..."
         Height          =   375
         Left            =   6480
         TabIndex        =   77
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtSource5 
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1920
         Width           =   7575
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   73
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox Picture9 
         Height          =   255
         Left            =   7320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   74
         Top             =   1440
         Width           =   255
         Begin VB.PictureBox picTmp5 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   255
            Left            =   0
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   75
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label lbMatch5 
         BackStyle       =   0  'Transparent
         Caption         =   "Images appear to be related."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   2760
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lbWarn5 
         BackStyle       =   0  'Transparent
         Caption         =   "Images do not appear to be related."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Decode a stenographic image - Step 2"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the image key:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0D55
         Height          =   975
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label lbPixels5 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Pixels"
         Height          =   255
         Left            =   2280
         TabIndex        =   82
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lbSize5 
         BackStyle       =   0  'Transparent
         Caption         =   "0 x 0"
         Height          =   255
         Left            =   2280
         TabIndex        =   81
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Per Character:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Size: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   3000
         Width           =   3015
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   528
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   528
      Y1              =   305
      Y2              =   305
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Declerations
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Global 'Emergency Break'
Dim Halt As Boolean

'Mode 0 = Encode, Mode 1 = Decode
Dim Mode As Integer

'Phyiscal Maximum Size in Bytes
Dim sMaxSize As Double

Dim sFileSize As Double

'Number of pixels to use per character
Dim sPixels As Integer

'Currently selected page
Dim Page As Integer

'The current progress bar value
Dim aPos As Long 'Current
Dim mPos As Long 'Max

Dim Data As String

Public Sub Show_Page(NewPage As Integer)
  Dim A As Integer
  
  Page = NewPage
  For A = picPage.LBound To picPage.UBound
    If A = NewPage Then
       picPage(A).Visible = True
    Else
       picPage(A).Visible = False
    End If
  Next A
End Sub

Private Sub btBack_Click()
  Select Case Page
    Case 1:
     Call Show_Page(0)
     btNext.Enabled = True
     btBack.Enabled = False
    Case 2:
     Call Show_Page(0)
     btNext.Enabled = True
     btBack.Enabled = False
    Case 3:
     Call Show_Page(1)
     btNext.Enabled = True
    Case 4:
     Call Show_Page(3)
     btFinish.Enabled = False
     btNext.Enabled = True
     btBack.Enabled = True
    Case 5:
     Call Show_Page(2)
     btNext.Enabled = True
     btBack.Enabled = True
    Case 6:
     Call Show_Page(5)
     btFinish.Enabled = False
     btNext.Enabled = True
     btBack.Enabled = True
  End Select
End Sub

Private Sub btBrowse1_Click()
  On Error GoTo Browse1Skip
  CD.Filter = "All Files (*.*)|*.*"
  CD.DialogTitle = "Select a source image"
  CD.ShowOpen
  txtSource1.Text = CD.FileName
  Call btLoad1_Click
Browse1Skip:
  On Error GoTo 0
End Sub

Private Sub btBrowse2_Click()
  On Error GoTo Browse2Skip
  CD.Filter = "All Files (*.*)|*.*"
  CD.DialogTitle = "Select an encoded image"
  CD.ShowOpen
  txtSource2.Text = CD.FileName
  Call btLoad2_Click
Browse2Skip:
  On Error GoTo 0
End Sub

Private Sub btBrowse3_Click()
  On Error GoTo Browse3Skip
  CD.Filter = "All Files (*.*)|*.*"
  CD.DialogTitle = "Select a source file"
  CD.ShowOpen
  txtSource3.Text = CD.FileName
  Call btLoad3_Click
Browse3Skip:
  On Error GoTo 0
End Sub

Private Sub btBrowse5_Click()
  On Error GoTo Browse5Skip
  CD.Filter = "All Files (*.*)|*.*"
  CD.DialogTitle = "Select the key image"
  CD.ShowOpen
  txtSource5.Text = CD.FileName
  Call btLoad5_Click
Browse5Skip:
  On Error GoTo 0
End Sub

Private Sub btCancel_Click()
  Dim Rtn As Integer
  Rtn = MsgBox("Are you sure you want to cancel?", vbYesNo, "Wizard")
  If Rtn = vbYes Then
     Unload Me
     End
  End If
End Sub

Private Sub btFinish_Click()
  Dim Rtn As Integer

  If btFinish.Caption = "&Exit" Then
     Unload Me
     End
     Exit Sub
  End If
  
  Halt = False
  btStop4.Enabled = True
  btStop6.Enabled = True
  btSave6.Enabled = False
  btFinish.Enabled = False
  
  If Page = 4 Then
     'Begin Encoding
     Call Encode
     
     'Save The File
     If Halt = False Then
        On Error GoTo Skip4
Retry4:
        CD.FileName = BaseName(txtSource1.Text) & "_code.bmp"
        CD.Filter = "Bitmap Files (*.bmp)|*.bmp"
        CD.DialogTitle = "Save the Encoded Image"
        CD.ShowSave
   
        Rtn = vbYes
        If Exists(CD.FileName) Then
           'Do we overwrite?
           Rtn = MsgBox("This file already exists! Overwrite it?", vbYesNo + vbExclamation, "Wizard")
        End If
        
        If Rtn = vbYes Then
           Call SavePicture(picTmp4.Image, CD.FileName)
           Call MsgBox("Image saved to " & CD.FileName, vbInformation + vbOKOnly, "Wizard")
           Unload Me
           End
        End If
Skip4:
        On Error GoTo 0
     End If
  ElseIf Page = 6 Then
     'Begin Decoding
     Call Decode
     
     'Save The File
     If Halt = False Then
        aPos = mPos
        
        lbStatus6.Caption = "Data decoded and ready for saving."
        lbStatus6.Refresh
        
        btSave6.Enabled = True
     End If
  End If
  
  If Halt Then
     aPos = 0
     lbStatus4.Caption = "Process interrupted."
     lbStatus6.Caption = "Process interrupted."
  End If
  
  btStop4.Enabled = False
  btStop6.Enabled = False
  btFinish.Enabled = True
End Sub

Private Sub btLoad1_Click()
  Dim Max As Single, Rec As Single
  
  btNext.Enabled = False
  lbSize1.Caption = "0 x 0"
  lbMax1.Caption = "0 KB"
  lbRec1.Caption = "0 KB"
  
  On Error GoTo Fail1:
  picTmp1.Picture = LoadPicture("")
  picTmp1.Picture = LoadPicture(txtSource1.Text)
  
  If Trim(txtSource1.Text) = "" Then GoTo Fail1
  
  sMaxSize = ((picTmp1.ScaleWidth * picTmp1.ScaleHeight) - 2)
  
  Max = Round(((picTmp1.ScaleWidth * picTmp1.ScaleHeight) - 2) / 1024, 2)
  Rec = Round((((picTmp1.ScaleWidth * picTmp1.ScaleHeight) - 2) / 20) / 1024, 2)
  lbSize1.Caption = picTmp1.ScaleWidth & " x " & picTmp1.ScaleHeight
  lbMax1.Caption = Max & " KB"
  lbRec1.Caption = Rec & " KB"
  
  lbMax3.Caption = lbMax1.Caption
  lbRec3.Caption = lbRec1.Caption
  
  btNext.Enabled = True
Fail1:
  On Error GoTo 0
End Sub

Private Sub btLoad2_Click()
  Dim Max As Single
  
  btNext.Enabled = False
  lbSize2.Caption = "0 x 0"
  lbMax2.Caption = "0 KB"
  
  On Error GoTo Fail2:
  picTmp2.Picture = LoadPicture("")
  picTmp2.Picture = LoadPicture(txtSource2.Text)
  
  If Trim(txtSource2.Text) = "" Then GoTo Fail2
  
  sMaxSize = ((picTmp2.ScaleWidth * picTmp2.ScaleHeight) - 2)
  
  Max = Round(((picTmp2.ScaleWidth * picTmp2.ScaleHeight) - 2) / 1024, 2)
  lbSize2.Caption = picTmp2.ScaleWidth & " x " & picTmp2.ScaleHeight
  lbMax2.Caption = Max & " KB"
  
  'lbMax3.Caption = lbMax1.Caption
  'lbRec3.Caption = lbRec1.Caption
  
  btNext.Enabled = True
Fail2:
  On Error GoTo 0
End Sub

Private Sub btLoad3_Click()
  Dim Size As Single
    
  btNext.Enabled = False
  lbSize3.Caption = "0 KB"
  lbSize3.ForeColor = vbBlack
  
  On Error GoTo Fail3:
  If Trim(txtSource3.Text) = "" Then GoTo Fail3
  
  Open txtSource3.Text For Binary As #1
  Close #1
  
  Size = FileLen(txtSource3.Text)
  sFileSize = Size
  lbSize3.Caption = Round(Size / 1024, 2) & " KB"
  
  If (Size < 1024) Then lbSize3.Caption = Size & " B"
  If ((Size / 1024) > 1024) Then lbSize3.Caption = Round((Size / 1024) / 1024, 2) & " MB"
  
  sPixels = sMaxSize / Size
  
  If (sPixels * Size) > sMaxSize Then sPixels = sPixels - 1
  
  If sPixels > 20 Then sPixels = 20
  If sPixels < 1 Then sPixels = 0
  
  lbPixel3.Caption = sPixels & " pixels"
  
  If sPixels >= 10 Then
     lbCaption4.Caption = "High Quality (10-15 Pixels)"
     If sPixels >= 15 Then lbCaption4.Caption = "Excellant Quality (15-20 Pixels)"
     lbCaption4.ForeColor = &HC000&
     
     lbDesc4.Caption = "With the current file sizes, you will get " & sPixels & " pixels per character, resulting in a fairly nice output.  To improve further, use smaller attachments and larger source images."
  Else
     lbCaption4.Caption = "Pathetic Quality (1-5 Pixels)"
     If sPixels >= 5 Then lbCaption4.Caption = "Low Quality (5-10 Pixels)"
     lbCaption4.ForeColor = vbRed
     
     lbDesc4.Caption = "Due to the respective file sizes, you will only get " & sPixels & " pixels per character, resulting in a shoddy output.  Try including a smaller file or use a larger source image for better results."
  End If
  
  If Size < sMaxSize And Size > 0 Then
     btNext.Enabled = True
  Else
     lbSize3.ForeColor = vbRed
  End If
Fail3:
  On Error GoTo 0
End Sub

Private Sub btLoad5_Click()
  btNext.Enabled = False
  lbSize5.Caption = "0 x 0"
  lbPixels5.Caption = "0 Pixels"
  lbMatch5.Visible = False
  lbWarn5.Visible = True
  
  On Error GoTo Fail5:
  picTmp5.Picture = LoadPicture("")
  picTmp5.Picture = LoadPicture(txtSource5.Text)
  
  If Trim(txtSource5.Text) = "" Then GoTo Fail5
  
  lbSize5.Caption = picTmp5.ScaleWidth & " x " & picTmp5.ScaleHeight
  
  sPixels = Compare5()
  
  If picTmp5.ScaleHeight = picTmp2.ScaleHeight And picTmp5.ScaleWidth = picTmp2.ScaleWidth And sPixels > 0 Then
     'Images match!
     lbPixels5.Caption = sPixels & " Pixels"
     lbWarn5.Visible = False
     lbMatch5.Visible = True
     btNext.Enabled = True
  End If

Fail5:
  On Error GoTo 0
End Sub

Public Function Compare5() As Integer
  Dim PntA As Long, PntB As Long
  Dim aR As Long, aG As Long, aB As Long, aStr As String
  Dim bR As Long, bG As Long, bB As Long, bStr As String
  
  Compare5 = 0
   
  PntA = picTmp5.Point(0, 0)  'Key Image
  PntB = picTmp2.Point(0, 0)  'Encoded Image
   
  If PntA <= 0 Then aR = 0: aG = 0: aB = 0
  If PntB <= 0 Then bR = 0: bG = 0: bB = 0
   
  If PntA > 0 Then
     aStr = FixHex(PntA)
     aR = CLng("&H" + (Mid(aStr, 5, 2)))
     aG = CLng("&H" + (Mid(aStr, 3, 2)))
     aB = CLng("&H" + (Mid(aStr, 1, 2)))
  End If
   
  If PntB > 0 Then
     bStr = FixHex(PntB)
     bR = CLng("&H" + (Mid(bStr, 5, 2)))
     bG = CLng("&H" + (Mid(bStr, 3, 2)))
     bB = CLng("&H" + (Mid(bStr, 1, 2)))
  End If
   
  Compare5 = Abs(bB - aB)
   
  If Compare5 < 1 Or Compare5 > 20 Then Compare5 = 0
  
  If aG <> bG Then Compare5 = 0
  If aR <> bR Then Compare5 = 0
End Function

Private Sub btNext_Click()
  Select Case Page
    Case 0: 'Main Page
     If Mode = 0 Then Call Show_Page(1)
     If Mode = 1 Then Call Show_Page(2)
     btNext.Enabled = False
     btBack.Enabled = True
    Case 1:
     Call Show_Page(3)
     btNext.Enabled = False
     btBack.Enabled = True
    Case 2:
     Call Show_Page(5)
     btNext.Enabled = False
     btBack.Enabled = True
    Case 3:
     Call Show_Page(4)
     btFinish.Enabled = True
     btNext.Enabled = False
     btBack.Enabled = True
     Call btFinish_Click
    Case 5:
     Call Show_Page(6)
     btFinish.Enabled = True
     btNext.Enabled = False
     btBack.Enabled = True
     Call btFinish_Click
  End Select
End Sub

Private Sub btSave6_Click()
   Dim FF As Integer, Rtn As Integer

   On Error GoTo Skip6
   CD.FileName = "data.bin"
   CD.Filter = "All Files (*.*)|*.*"
   CD.DialogTitle = "Save the extracted data"
   CD.ShowSave
       
   Rtn = vbYes
   
   If Exists(CD.FileName) Then
      'Do we overwrite?
      Rtn = MsgBox("This file already exists! Overwrite it?", vbYesNo + vbExclamation, "Wizard")
   End If
   
   If Rtn = vbYes Then
      FF = FreeFile
      Open CD.FileName For Binary Access Write As #FF
      Put #FF, , Data
      Close #FF
  
      'Call MsgBox("Data saved to " & CD.FileName, vbInformation + vbOKOnly, "Wizard")
    
      btFinish.Caption = "&Exit"
   
      'Unload Me
      'End
   End If
Skip6:
End Sub

Private Sub btStop4_Click()
  Halt = True
End Sub

Private Sub btStop6_Click()
  Halt = True
End Sub

Private Sub Form_Load()
  Call Show_Page(0)
  Halt = False
End Sub

Private Sub optMode_Click(Index As Integer)
  Mode = Index
  If Page = 0 Then btNext.Enabled = True
End Sub

Public Sub Encode()
   Dim Slots(1 To 60) As Integer
   Dim tX As Integer, tY As Integer
   Dim x As Integer, y As Integer
   Dim A As Long, B As Long
   Dim Pnt As Long, sCnt As Integer, Pref As Integer
   Dim Dat As Integer, Tmp As String, Perc As Integer
   Dim tVal As Integer
   
   Dim FF As Integer
   
   Dim tR As Long, tG As Long, tB As Long, tStr As String
   
   'aPrev = -1
   'aTime = -1
   
   lbStatus4.Caption = "Loading source image into work environment..."
   lbStatus4.Refresh
   
   picTmp4.Picture = LoadPicture(txtSource1.Text)
   picTmp4.Refresh
   x = 1
   y = 0
   
   lbStatus4.Caption = "Attaching data pixel to output image..."
   lbStatus4.Refresh
   
   '--------------------------> Key Bit <---------------------------
   Pnt = GetPixel(picTmp1.hdc, 0, 0)
   If Pnt <= 0 Then
      tR = 0: tG = 0: tB = 0
   Else
      tStr = FixHex(Pnt)
      tR = CLng("&H" + (Mid(tStr, 5, 2)))
      tG = CLng("&H" + (Mid(tStr, 3, 2)))
      tB = CLng("&H" + (Mid(tStr, 1, 2)))
   End If
   If tB + sPixels > 255 Then
      tB = tB - sPixels
   Else
      tB = tB + sPixels
   End If
   Call SetPixel(picTmp4.hdc, 0, 0, RGB(tR, tG, tB))
   'picTmp4.PSet (0, 0), (RGB(tB, tG, tR))
   
   'Begin Encoding!
   Tmp = sPixels * sFileSize
   Perc = Int((Tmp / (picTmp1.ScaleWidth * picTmp1.ScaleHeight) * 100))
   mPos = sFileSize + 1
   Bar4.Value = 0
   
   lbStatus4.Caption = "Open source file for input..."
   lbStatus4.Refresh
   
   'Open the source file for reading
   FF = FreeFile
   Open txtSource3.Text For Binary As FF
            
   lbStatus4.Caption = "Encoding source image with data from source file..."
   lbStatus4.Refresh
   
   For A = 1 To sFileSize + 1
      aPos = A
    
      If A = sFileSize + 1 Then
        Dat = 256 'Data terminator - 'Virtual' character
      Else
        'Read a character from the source file
        Dat = Asc(Input(1, FF))
      End If
           
      'Grab Pixels into Memory
      tX = x: tY = y: sCnt = 1
      For B = 1 To sPixels
        Pnt = GetPixel(picTmp1.hdc, tX, tY)
        'Pnt = picKey.Point(tX, tY)
        If Pnt <= 0 Then
           tR = 0: tG = 0: tB = 0
        Else
           tStr = FixHex(Pnt)
           Do
            If Len(tStr) >= 6 Then Exit Do
            tStr = "0" & tStr
           Loop
           tR = CLng("&H" + (Mid(tStr, 5, 2)))
           tG = CLng("&H" + (Mid(tStr, 3, 2)))
           tB = CLng("&H" + (Mid(tStr, 1, 2)))
         End If
        Slots(sCnt + 0) = tR
        Slots(sCnt + 1) = tG
        Slots(sCnt + 2) = tB
        sCnt = sCnt + 3
        tX = tX + 1
        If tX >= picTmp1.ScaleWidth Then tX = 0: tY = tY + 1
      Next B
      
      If Halt Then Exit Sub
      
      'Encode
      Do
         Pref = Int(Dat / (sCnt - 1))
         If Dat Mod (sCnt - 1) <> 0 Then Pref = Pref + 1
         
         For B = 1 To (sCnt - 1)
          tVal = Dat
          If tVal > Pref Then tVal = Pref
          If Slots(B) + tVal > 255 Then
             If Slots(B) - tVal > 0 Then
                'Use -
                Slots(B) = Slots(B) - tVal
                Dat = Dat - tVal
             End If
          Else
             'Use +
             Slots(B) = Slots(B) + tVal
             Dat = Dat - tVal
          End If
           
         Next B
         
      Loop Until Dat <= 0
      
      If Halt Then Exit Sub
      
      'Replace Pixels
      sCnt = 1
      For B = 1 To sPixels
        Call SetPixel(picTmp4.hdc, x, y, RGB(Slots(sCnt + 0), Slots(sCnt + 1), Slots(sCnt + 2)))
        'picEnd.PSet (x, y), (RGB(Slots(SCnt + 2), Slots(SCnt + 1), Slots(SCnt + 0)))
        x = x + 1
        If x >= picTmp1.ScaleWidth Then x = 0: y = y + 1
        sCnt = sCnt + 3
      Next B
      
      If Halt Then Exit Sub
      
      DoEvents
   Next A
   
   lbStatus4.Caption = "Closing source file..."
   lbStatus4.Refresh
   
   Close FF
   
   lbStatus4.Caption = "Finished encoding! You may now save your stenographic image."
   lbStatus4.Refresh
   
   picTmp4.Refresh
End Sub


Public Function FixHex(ByVal Num As Long) As String
   FixHex = Hex(Num)
   Do
     If Len(FixHex) >= 6 Then Exit Do
     FixHex = "0" & FixHex
   Loop
End Function

Private Sub tmrBars_Timer()
   If mPos <= 0 Then mPos = 1
   If Bar4.Max <> mPos Then Bar4.Max = mPos
   If Bar6.Max <> mPos Then Bar6.Max = mPos
   
   Bar4.Value = aPos
   Bar6.Value = aPos
End Sub

Public Function BaseName(ByVal Msg As String) As String
   Dim tPos As Integer
   
   BaseName = App.Path & "\MyImage"
   
   tPos = InStr(1, Msg, ".")
   If tPos > 1 Then
      BaseName = Mid(Msg, 1, tPos - 1)
   End If
End Function

Public Sub Decode()
   Dim x As Integer, y As Integer
   Dim PntA As Long, PntB As Long
   Dim Dat As Integer, Cnt As Integer
   Dim Sum As Integer, Diff As Integer
   Dim Bar As Long, BCnt As Long, tStr As String
   Dim aR As Long, aG As Long, aB As Long
   Dim bR As Long, bG As Long, bB As Long
   Dim tData As String
   
   Data = ""
   Cnt = 0
   Sum = 0
   aPos = 0
   
   'ePixel is already set, we'll use that for our offset
   '-----
   ' Goal: Read (ePixel) values into memory. Compare them to the source. Sum them up.
   '  Repeat until you run out of pixels.
    
   lbStatus6.Caption = "Preparing to decode data from stenographic image..."
   lbStatus6.Refresh
   
   Bar = ((picTmp5.ScaleWidth * picTmp5.ScaleHeight) / sPixels) / 100
   mPos = 100
   Bar6.Value = 0
   BCnt = 0
     
   lbStatus6.Caption = "Decoding data from stenographic image..."
   lbStatus6.Refresh
   
   For y = 0 To picTmp5.ScaleHeight - 1
    For x = 0 To picTmp5.ScaleWidth - 1
     If y = 0 And x = 0 Then x = 1 'Skip the key character
     
     'Read the two target pixels
     PntA = GetPixel(picTmp5.hdc, x, y) 'Key
     PntB = GetPixel(picTmp2.hdc, x, y) 'Encoded
     
     'Convert to RGB codes.
     If PntA <= 0 Then
        aR = 0: aG = 0: aB = 0
     Else
        tStr = FixHex(PntA)
        aR = CLng("&H" + (Mid(tStr, 5, 2)))
        aG = CLng("&H" + (Mid(tStr, 3, 2)))
        aB = CLng("&H" + (Mid(tStr, 1, 2)))
     End If
     
     If PntB <= 0 Then
        bR = 0: bG = 0: bB = 0
     Else
        tStr = FixHex(PntB)
        bR = CLng("&H" + (Mid(tStr, 5, 2)))
        bG = CLng("&H" + (Mid(tStr, 3, 2)))
        bB = CLng("&H" + (Mid(tStr, 1, 2)))
     End If
     
     'Create a difference value
     Diff = 0
     Diff = Diff + Abs(bR - aR)
     Diff = Diff + Abs(bG - aG)
     Diff = Diff + Abs(bB - aB)
     Sum = Sum + Diff
     
     'Count the number of pixels we have read
     Cnt = Cnt + 1
     
     If Cnt = sPixels Then
        'Update the progress
        BCnt = BCnt + 1
        If BCnt = Bar Then
          BCnt = 0
          If aPos < mPos Then aPos = aPos + 1
        End If
        
        'All characters collected, build the character.
        ' 256 is a virtual character, a file terminator
        If Sum > 256 Then
           'ERROR
           Call MsgBox("Fatal Decoder Error! Aborting Decode." & Chr(10) & "(Pixel " & x & ", " & y & ")", vbCritical + vbOKOnly, "Wizard")
           Halt = True
           Exit Sub
        ElseIf Sum = 256 Then
           'End of file
           aPos = mPos
           Data = Data & tData
           Exit Sub
        Else
           tData = tData & Chr(Sum)
           If Len(tData) > 1000 Then
              Data = Data & tData
              tData = ""
           End If
        End If
        Sum = 0
        Cnt = 0
        
        'Allow events every time we complete a character.
        DoEvents
        
        'Halt Event
        If Halt Then Exit Sub
     End If
    Next x
   Next y
   
   Data = Data & tData
   
   aPos = mPos
End Sub

Public Function Exists(FName As String) As Boolean
   Dim FF As Integer
   
   Exists = False
   
   On Error GoTo DoesNotExist
   
   FF = FreeFile
   Open FName For Input As #FF
   Close #FF
   
   Exists = True
DoesNotExist:
On Error GoTo 0
End Function
