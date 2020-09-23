VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Advanced MP3 Info Editor"
   ClientHeight    =   9390
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList Buttons 
      Left            =   10380
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0077
            Key             =   "addi"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00F5
            Key             =   "del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01E6
            Key             =   "deli"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02D7
            Key             =   "next"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03C6
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04B6
            Key             =   "previ"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10320
      TabIndex        =   1
      Top             =   2220
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10470
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   7
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0688
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10710
      Top             =   3390
   End
   Begin VB.Frame Frame1 
      Caption         =   "MP3 ID Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   60
      TabIndex        =   2
      Top             =   4950
      Width           =   9735
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   9255
         Begin VB.TextBox txtTracksTotal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5640
            TabIndex        =   16
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtLyrics 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   5400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1440
            Width           =   3855
         End
         Begin VB.TextBox txtYear 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6960
            TabIndex        =   18
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtComments 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   1440
            Width           =   3615
         End
         Begin VB.TextBox txtTrackNumber 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   14
            Top             =   1080
            Width           =   495
         End
         Begin VB.ComboBox cmbGenre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":075A
            Left            =   600
            List            =   "frmMain.frx":091D
            TabIndex        =   12
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtAlbum 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   10
            Top             =   720
            Width           =   8655
         End
         Begin VB.TextBox txtArtist 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   8
            Top             =   360
            Width           =   8655
         End
         Begin VB.TextBox txtTitle 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   6
            Top             =   0
            Width           =   8655
         End
         Begin VB.Label Label39 
            Caption         =   "of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5400
            TabIndex        =   15
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "Lyrics:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   21
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Year:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   17
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Comments:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Track:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4260
            TabIndex        =   13
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Genre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Album:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Artist:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Title:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         Begin VB.VScrollBar VScroll1 
            Height          =   2655
            LargeChange     =   5
            Left            =   9000
            Max             =   29
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   9825
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   9015
            Begin VB.TextBox txtInterpretedBy 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   32
               Top             =   1080
               Width           =   7095
            End
            Begin VB.TextBox txtPublisher 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   50
               Top             =   4320
               Width           =   7095
            End
            Begin VB.TextBox txtDiscsTotal 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6000
               TabIndex        =   86
               Top             =   9480
               Width           =   495
            End
            Begin VB.TextBox txtDiscNumber 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5160
               TabIndex        =   84
               Top             =   9480
               Width           =   495
            End
            Begin VB.TextBox txtInternetRadioStationURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   72
               Top             =   7920
               Width           =   7095
            End
            Begin VB.TextBox txtAudioSourceURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   70
               Top             =   7560
               Width           =   7095
            End
            Begin VB.TextBox txtArtistURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   67
               Top             =   7200
               Width           =   5850
            End
            Begin VB.TextBox txtCopyrightInfo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   63
               Top             =   6480
               Width           =   7095
            End
            Begin VB.TextBox txtCommercialInfo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   60
               Top             =   6120
               Width           =   5850
            End
            Begin VB.TextBox txtISRC 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   56
               Top             =   5400
               Width           =   7095
            End
            Begin VB.TextBox txtInternetRadioStationOwner 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   54
               Top             =   5040
               Width           =   7095
            End
            Begin VB.TextBox txtInternetRadioStationName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   52
               Top             =   4680
               Width           =   7095
            End
            Begin VB.TextBox txtFileOwner 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   48
               Top             =   3960
               Width           =   7095
            End
            Begin VB.TextBox txtConductor 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   30
               Top             =   720
               Width           =   7095
            End
            Begin VB.TextBox txtBand 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   28
               Top             =   360
               Width           =   7095
            End
            Begin VB.TextBox txtPublisherURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   76
               Top             =   8640
               Width           =   7095
            End
            Begin VB.TextBox txtComposer 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   26
               Top             =   0
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalArtist 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   36
               Top             =   1800
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalAlbum 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   38
               Top             =   2160
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalFileName 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   40
               Top             =   2520
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalLyricist 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   42
               Top             =   2880
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalReleaseYear 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   44
               Top             =   3240
               Width           =   7095
            End
            Begin VB.TextBox txtCopyright 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   46
               Top             =   3600
               Width           =   7095
            End
            Begin VB.TextBox txtLanguages 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   58
               Top             =   5760
               Width           =   7095
            End
            Begin VB.TextBox txtLyricist 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   34
               Top             =   1440
               Width           =   7095
            End
            Begin VB.TextBox txtBPM 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   480
               TabIndex        =   80
               Top             =   9480
               Width           =   1095
            End
            Begin VB.ComboBox cmbKey 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmMain.frx":0EF2
               Left            =   2880
               List            =   "frmMain.frx":0F62
               TabIndex        =   82
               Top             =   9480
               Width           =   1455
            End
            Begin VB.TextBox txtAudioURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   65
               Top             =   6840
               Width           =   7095
            End
            Begin VB.TextBox txtPaymentURL 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   74
               Top             =   8280
               Width           =   7095
            End
            Begin VB.TextBox txtEncodedBy 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   78
               Top             =   9000
               Width           =   7095
            End
            Begin VB.Label countArtistURL 
               Alignment       =   1  'Right Justify
               Caption         =   "0/0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   7680
               TabIndex        =   68
               Top             =   7260
               Width           =   570
            End
            Begin VB.Label countCommercialInfo 
               Alignment       =   1  'Right Justify
               Caption         =   "0/0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   7680
               TabIndex        =   61
               Top             =   6180
               Width           =   570
            End
            Begin VB.Image delArtistURL 
               Height          =   285
               Left            =   8685
               Picture         =   "frmMain.frx":111A
               ToolTipText     =   "Delete Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image nextArtistURL 
               Height          =   285
               Left            =   8475
               Picture         =   "frmMain.frx":11FB
               ToolTipText     =   "Next Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image prevArtistURL 
               Height          =   285
               Left            =   8265
               Picture         =   "frmMain.frx":12DA
               ToolTipText     =   "Previous Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image delCommercialInfo 
               Height          =   285
               Left            =   8685
               Picture         =   "frmMain.frx":13BA
               ToolTipText     =   "Delete Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Image nextCommercialInfo 
               Height          =   285
               Left            =   8475
               Picture         =   "frmMain.frx":149B
               ToolTipText     =   "Next Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Image prevCommercialInfo 
               Height          =   285
               Left            =   8265
               Picture         =   "frmMain.frx":157A
               ToolTipText     =   "Previous Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Label Label40 
               Caption         =   "Interpreted by:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   31
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label Label38 
               Caption         =   "Publisher:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   49
               Top             =   4320
               Width           =   1695
            End
            Begin VB.Label Label37 
               Caption         =   "of"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5760
               TabIndex        =   85
               Top             =   9480
               Width           =   255
            End
            Begin VB.Label Label36 
               Caption         =   "Disc:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4680
               TabIndex        =   83
               Top             =   9480
               Width           =   375
            End
            Begin VB.Label Label35 
               Caption         =   "Net Radio Station URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   71
               Top             =   7920
               Width           =   1695
            End
            Begin VB.Label Label34 
               Caption         =   "Audio Source URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   69
               Top             =   7560
               Width           =   1695
            End
            Begin VB.Label Label33 
               Caption         =   "Artist URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   66
               Top             =   7200
               Width           =   1695
            End
            Begin VB.Label Label32 
               Caption         =   "Copyright Info URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   62
               Top             =   6480
               Width           =   1695
            End
            Begin VB.Label Label31 
               Caption         =   "Commercial Info URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   59
               Top             =   6120
               Width           =   1695
            End
            Begin VB.Label Label30 
               Caption         =   "ISRC:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   55
               Top             =   5400
               Width           =   1695
            End
            Begin VB.Label Label29 
               Caption         =   "Net Radio Stn. Owner:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   53
               Top             =   5040
               Width           =   1695
            End
            Begin VB.Label Label28 
               Caption         =   "Net Radio Stn. Name:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   4680
               Width           =   1695
            End
            Begin VB.Label Label27 
               Caption         =   "File Owner/Licensee:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   47
               Top             =   3960
               Width           =   1695
            End
            Begin VB.Label Label26 
               Caption         =   "Conductor:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   29
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label25 
               Caption         =   "Band/Orchestra:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   27
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label24 
               Caption         =   "Publisher URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   75
               Top             =   8640
               Width           =   1695
            End
            Begin VB.Label Label9 
               Caption         =   "Composer:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Width           =   1695
            End
            Begin VB.Label Label10 
               Caption         =   "Original Artist:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   35
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label11 
               Caption         =   "Original Album:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   37
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label12 
               Caption         =   "Original Filename:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   2520
               Width           =   1695
            End
            Begin VB.Label Label13 
               Caption         =   "Original Lyricist:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   41
               Top             =   2880
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "Original Release Year:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   43
               Top             =   3240
               Width           =   1695
            End
            Begin VB.Label Label15 
               Caption         =   "Copyright:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   45
               Top             =   3600
               Width           =   1695
            End
            Begin VB.Label Label16 
               Caption         =   "Languages:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   57
               Top             =   5760
               Width           =   1695
            End
            Begin VB.Label Label17 
               Caption         =   "Lyricist:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   33
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label19 
               Caption         =   "BPM:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   79
               Top             =   9480
               Width           =   375
            End
            Begin VB.Label Label20 
               Caption         =   "Initial Key:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1920
               TabIndex        =   81
               Top             =   9480
               Width           =   855
            End
            Begin VB.Label Label21 
               Caption         =   "Audio URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   64
               Top             =   6840
               Width           =   1695
            End
            Begin VB.Label Label22 
               Caption         =   "Payment URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   73
               Top             =   8280
               Width           =   1695
            End
            Begin VB.Label Label23 
               Caption         =   "Encoded by:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   77
               Top             =   9000
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   2
         Left            =   240
         TabIndex        =   92
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         Begin VB.ComboBox cmbImageType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":165A
            Left            =   3960
            List            =   "frmMain.frx":166A
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   1920
            Width           =   1275
         End
         Begin VB.PictureBox picArt 
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   3960
            ScaleHeight     =   1755
            ScaleWidth      =   1755
            TabIndex        =   93
            Top             =   0
            Width           =   1815
            Begin VB.Image imgArt 
               Height          =   1755
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1755
            End
            Begin VB.Label lblBrowse 
               Alignment       =   2  'Center
               BackColor       =   &H8000000C&
               BackStyle       =   0  'Transparent
               Caption         =   "Click here to browse..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   495
               Left            =   0
               TabIndex        =   94
               Top             =   720
               Width           =   1815
            End
         End
         Begin VB.ComboBox cmbPictureType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":1683
            Left            =   3960
            List            =   "frmMain.frx":16C6
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   2280
            Width           =   2520
         End
         Begin VB.Label countArt 
            Alignment       =   1  'Right Justify
            Caption         =   "0/0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   5265
            TabIndex        =   99
            Top             =   1980
            Width           =   570
         End
         Begin VB.Image delArt 
            Height          =   285
            Left            =   6270
            Picture         =   "frmMain.frx":17F5
            Top             =   1920
            Width           =   210
         End
         Begin VB.Image nextArt 
            Height          =   285
            Left            =   6060
            Picture         =   "frmMain.frx":18D6
            Top             =   1920
            Width           =   210
         End
         Begin VB.Image prevArt 
            Height          =   285
            Left            =   5850
            Picture         =   "frmMain.frx":19B5
            Top             =   1920
            Width           =   210
         End
         Begin VB.Label Label43 
            Caption         =   "Picture type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   97
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label41 
            Caption         =   "Image type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   95
            Top             =   1920
            Width           =   975
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Search Info via &Microsoft"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   89
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Search Info in iTunes Store"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete MP3 Tags"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   91
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Update MP3 Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   90
         Top             =   3600
         Width           =   1455
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5741
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Basic"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Advanced"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Album Art"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10530
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.TreeView tvwLibrary 
      Height          =   5175
      Left            =   11010
      TabIndex        =   100
      Top             =   4140
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   9128
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList imgLibrary 
      Left            =   10380
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   10290
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A95
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2841
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3253
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5437
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6305
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":669F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":776D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D07
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80A1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   106
      Top             =   9075
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14579
            Picture         =   "frmMain.frx":843B
            Text            =   "Neo Player: Media Library"
            TextSave        =   "Neo Player: Media Library"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "frmMain.frx":89D5
            Text            =   "Records: "
            TextSave        =   "Records: "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicClientArea 
      Height          =   4965
      Left            =   60
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   645
      TabIndex        =   102
      Top             =   30
      Width           =   9735
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   120
         Width           =   7785
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Add Folder ....."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   103
         Top             =   90
         Width           =   1395
      End
      Begin MSComctlLib.TreeView TreeFiles 
         Height          =   4245
         Left            =   105
         TabIndex        =   104
         Top             =   480
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   7488
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgIconos"
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4245
         Left            =   2850
         TabIndex        =   105
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7488
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   16384
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Artist"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Album"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Genre"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Track No."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Tracks Total"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Year"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Duration"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Bit Rate"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Comments"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   101
      Top             =   4770
      Width           =   7275
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAddMedia 
         Caption         =   "Add Media to Library"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuNewLibrary 
         Caption         =   "New Llibrary"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuArt 
      Caption         =   "ArtMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuArtItem 
         Caption         =   "&Copy"
         Index           =   0
      End
      Begin VB.Menu mnuArtItem 
         Caption         =   "&Paste"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'DefLng A-Z

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Const SW_SHOW As Long = 5
Private Const CF_BITMAP As Long = 2
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4

Private Const S_OTHER As String = "Other"

Private Const FILTER_BMP As String = "*.bmp;*.dib"
Private Const FILTER_GIF As String = "*.gif"
Private Const FILTER_JPEG As String = "*.jpeg;*.jpg;*.jpe;*.jfif;*.jfi;*.jif"
Private Const FILTER_PNG As String = "*.png"
Private Const FILTER_SUPPORTED As String = FILTER_BMP & ";" & FILTER_GIF & ";" & FILTER_JPEG & ";" & FILTER_PNG

Private Const MNU_COPY As Long = 0
Private Const MNU_PASTE As Long = 1

Private Const PASTE_TXT_1 As String = "&Paste"
Private Const PASTE_TXT_2 As String = PASTE_TXT_1 & " (this will change the current image)"

Dim myWindowState As Integer
Dim bInitialized As Boolean


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
Dim AutoSize As Boolean
Dim strSQL As String
Dim sConnectionString As String
Dim SQL As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim cnnMusic As ADODB.Connection
Dim cmd   As ADODB.Command


Public Function BindToSQL(strSQL As String, Optional iSQL As ADODB.Connection, Optional iRs As ADODB.Recordset)
   On Error GoTo err_handler:
   'On Error Resume Next
    If iSQL Is Nothing Then
        Set iSQL = SQL
        If iSQL.state > 0 Then iSQL.Close
        Dim connstring As String
        connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Library\music.mdb;Persist Security Info=False"
        iSQL.Open connstring
    End If
    If strSQL = "" Then Exit Function
    
    If iRs Is Nothing Then Set iRs = rs
    If iRs.state > 0 Then iRs.Close
    iRs.Open strSQL, iSQL, adOpenKeyset, adLockOptimistic
    
    Dim sF As Field
    With ListView1
    .ListItems.Clear
    .ColumnHeaders.Clear
    Dim i As Integer
    i = 0
    For Each sF In iRs.Fields
        i = i + 1
        Select Case i
           Case 1: .ColumnHeaders.Add , , sF.name, 1.9 * ListView1.Width / 9
           Case 2: .ColumnHeaders.Add , , sF.name, 1.2 * ListView1.Width / 9
           Case 5: .ColumnHeaders.Add , , sF.name, 0.5 * ListView1.Width / 9
           Case 6: .ColumnHeaders.Add , , sF.name, 0.6 * ListView1.Width / 9
           Case 7: .ColumnHeaders.Add , , sF.name, 0.8 * ListView1.Width / 9
           Case 9: .ColumnHeaders.Add , , sF.name, 2 * ListView1.Width / 9
           Case Else: .ColumnHeaders.Add , , sF.name, ListView1.Width / 9
        End Select
    Next
 
    Dim x As Integer, Item
    Do While Not iRs.EOF
        For x = 0 To iRs.Fields.Count - 1
            If x = 0 Then
              If iRs(x).Value <> "" Then Set Item = ListView1.ListItems.Add(, , iRs(x).Value)
            Else
              If iRs(x).Value <> "" Then Item.SubItems(x) = iRs(x).Value
            End If
        Next
    iRs.MoveNext
    Loop
   
    End With

Exit Function
err_handler:
    MsgBox "Error binding ListView to SQL" & vbNewLine & vbNewLine & "Error code:" & Err.Number & vbNewLine & "Error desc:" & Err.Description, vbCritical
End Function

Private Sub cmdSearch_Click()
    Dim Folder As String
    Dim sExistingFolder As String
    
    If Right$(Text1, 1) = "\" Then
        sExistingFolder = Text1
    Else
        sExistingFolder = Text1 & "\"
    End If
    
    Folder = BrowseForFolder(hWnd, "Select a folder:", sExistingFolder)
    If Folder <> "" Then
        Text1 = Folder
        LoadFileEntries Folder
        LoadLibrary
    End If
End Sub

Private Sub Command5_Click()
'ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & tAppConfig.AppConfig & "Library\music.mdb;Persist Security Info=False"
LoadLibrary
BindToSQL "SELECT * FROM music"

End Sub

Private Sub cmbImageType_Change()
    cmbImageType_Click
End Sub

Private Sub cmbImageType_Click()
    Dim MIMEType As String
    Dim PNGIndex As Long
    If cmbImageType.Enabled Then
        If cmbImageType.ListCount = 4 Then
            Select Case cmbImageType.ListIndex
                Case 0: MIMEType = ImageBMP
                Case 1: MIMEType = ImageGIF
                Case 2: MIMEType = ImageJPEG
                Case 3: MIMEType = ImagePNG
            End Select
            PNGIndex = 3
        Else
            Select Case cmbImageType.ListIndex
                Case 0: MIMEType = ImageJPEGOld
                Case 1: MIMEType = ImagePNGOld
            End Select
            PNGIndex = 1
        End If
        SetItem cAPICIType, indAPIC, MIMEType
        If cmbPictureType.ListIndex = 1 And cmbImageType.ListIndex <> PNGIndex Then
            cmbPictureType.ListIndex = 2
            SetItem cAPICType, indAPIC, cmbImageType.ListIndex
        End If
    End If
End Sub

Private Sub cmbPictureType_Change()
    cmbPictureType_Click
End Sub

Private Sub cmbPictureType_Click()
    If cmbPictureType.Enabled Then
        If cmbPictureType.ListIndex = 1 Then
            If cmbImageType.ListIndex <> (1 + 2 * (cmbImageType.ListCount \ 4)) Or HimetricToPixelsX(imgArt.Picture.Width) <> 32 Or HimetricToPixelsY(imgArt.Picture.Height) <> 32 Then
                cmbPictureType.ListIndex = 2
            End If
        End If
        SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
    End If
End Sub

Private Sub Command2_Click()
     Dim ID3 As New clsID3
    Dim i As Long
    
    With ID3
        .FileName = ListView1.SelectedItem.SubItems(8)
        .Title = txtTitle
        .Artist = txtArtist
        .Album = txtAlbum
        .Genre = cmbGenre.Text
        .GenreID = .ToGenreID(.Genre)
        .TrackNumber = txtTrackNumber
        .TracksTotal = txtTracksTotal
        .Year = txtYear
        .Comments = txtComments
        .Lyrics = txtLyrics
        .Composer = txtComposer
        .Band = txtBand
        .Conductor = txtConductor
        .InterpretedBy = txtInterpretedBy
        .Lyricist = txtLyricist
        .OriginalArtist = txtOriginalArtist
        .OriginalAlbum = txtOriginalAlbum
        .OriginalFileName = txtOriginalFileName
        .OriginalLyricist = txtOriginalLyricist
        .OriginalReleaseYear = txtOriginalReleaseYear
        .Copyright = txtCopyright
        .FileOwner = txtFileOwner
        .Publisher = txtPublisher
        .InternetRadioStationName = txtInternetRadioStationName
        .InternetRadioStationOwner = txtInternetRadioStationOwner
        .ISRC = txtISRC
        .Languages = txtLanguages
        .CommercialInfo.Clear
        For i = 1 To cWCOM.Count
            .CommercialInfo.Add cWCOM(i)
        Next
        .CopyrightInfo = txtCopyrightInfo
        .AudioURL = txtAudioURL
        .ArtistURL.Clear
        For i = 1 To cWOAR.Count
            .ArtistURL.Add cWOAR(i)
        Next
        .AudioSourceURL = txtAudioSourceURL
        .InternetRadioURL = txtInternetRadioStationURL
        .PaymentURL = txtPaymentURL
        .PublisherURL = txtPublisherURL
        .EncodedBy = txtEncodedBy
        .BeatsPerMinute = txtBPM
        .InitialKey = cmbKey
        .DiscNumber = txtDiscNumber
        .DiscsTotal = txtDiscsTotal
        For i = 1 To cAPICData.Count
            MakeNecessaryChanges i
        Next
        Set .AttachedPictures = APICData
        
        If MousePointer = vbDefault Then
            MousePointer = vbHourglass
            DoEvents
        End If
        
        .UpdateID3Tags
        
        If MousePointer = vbHourglass Then _
           MousePointer = vbDefault
        
        ListView1_ItemClick ListView1.SelectedItem
    End With
End Sub

Private Sub Command3_Click()
    Dim ID3 As New clsID3
    
    With ID3
        .FileName = Text1 & "\" & ListView1.SelectedItem.Text
        
        If MousePointer = vbDefault Then
            MousePointer = vbHourglass
            DoEvents
        End If
        
        .DeleteID3Tags
        
        If MousePointer = vbHourglass Then _
           MousePointer = vbDefault
        
        ListView1_ItemClick ListView1.SelectedItem
    End With
End Sub

Private Function Ceiling(ByVal num As Double) As Double
    Dim d As Double
    d = num
    If num <> Fix(num) Then d = d + 1
    Ceiling = d
End Function

Private Sub Command4_Click()
    On Error Resume Next
    
    bConnect = True
    If Dir$(GetSpecialFolderLocation(CSIDL_SYSTEM) & "\vbzlib1.dll") = "" Then
        bRet = False
        frmGZip.Show 1
        If bRet Then GoTo NormProc
    Else
NormProc:
        bRet = False
        bItunes = True
        frmConnecting.Show 1
        If bRet Then
            bRet = False
            frmSelection.Show 1
            If bRet Then
                bRet = False
                bConnect = False
                frmConnecting.Show 1
            ElseIf bLaunchArtDL Then
                frmAlbumArtDL.Show 1
            End If
        End If
    End If
End Sub



Private Sub Command6_Click()
    On Error Resume Next
    
    bConnect = True
    bRet = False
    bItunes = False
    frmConnecting.Show 1
    If bRet Then
        bRet = False
        frmSelection.Show 1
        If bRet Then
            bRet = False
            bConnect = False
            frmConnecting.Show 1
        ElseIf bLaunchArtDL Then
            frmAlbumArtDL.Show 1
        End If
    End If
End Sub

Private Sub delArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(delArt, x, Y) Then
        DelProc picArt, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End If
End Sub

Private Sub delArtistURL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(delArtistURL, x, Y) Then
        DelProc txtArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    End If
End Sub

Private Sub delCommercialInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(delCommercialInfo, x, Y) Then
        DelProc txtCommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    End If
End Sub

Private Sub imgArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If WithinBounds(imgArt, x, Y) Then
        Select Case Button
            Case vbLeftButton: ImageBrowse
            Case vbRightButton: If ValidateMenu Then PopupMenu mnuArt
        End Select
    End If
End Sub

Private Sub lblBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If WithinBounds(picArt, x + lblBrowse.Left, Y + lblBrowse.Top) Then
        Select Case Button
            Case vbLeftButton: ImageBrowse
            Case vbRightButton: If ValidateMenu Then PopupMenu mnuArt
        End Select
    End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        ListView1_DblClick
    End If
End Sub

Private Sub mnuAddmedia_Click()
cmdSearch_Click
End Sub

Private Sub mnuArtItem_Click(Index As Integer)
    On Error Resume Next
    Dim hMem As Long
    Dim mPic As StdPicture
    Dim GPC As GDIPlusCandy
    Dim st As String
    Select Case Index
        Case MNU_COPY
            If OpenClipboard(0) Then
                EmptyClipboard
                hMem = CopyImage(imgArt.Picture.Handle, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                SetClipboardData CF_BITMAP, hMem
                DeleteObject hMem
                CloseClipboard
            End If
        Case MNU_PASTE
            Set mPic = Clipboard.GetData(CF_BITMAP)
            If Not mPic Is Nothing And mPic.Handle <> 0 Then
                cmbImageType.ListIndex = 1 + 2 * (cmbImageType.ListCount \ 4)
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                Set GPC = New GDIPlusCandy
                st = GPC.ImageToData(mPic, ImagePNG)
                ArtAddProc ImagePNG, GPC.DataToImage(st), st
                Set GPC = Nothing
            End If
    End Select
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub


Private Sub mnuNewLibrary_Click()
Dim k
If k = 7 Then Exit Sub
k = MsgBox("Are you sure you want to Erase all data from library?", vbYesNo, "NeoPlayer Media Library")
'k=6=Yes,  k=7=no
If k = 6 Then
On Error GoTo hell
Dim rs As New ADODB.Recordset
rs.Open "SELECT FILE FROM MUSIC", cnnMusic, adOpenDynamic, adLockOptimistic

Do Until rs.EOF
   rs.Delete
   rs.MoveNext
Loop
rs.Update
rs.Close
Set rs = Nothing
End If
hell:
LoadLibrary
ListView1.ListItems.Clear

End Sub

Private Sub mnuRefresh_Click()
    LoadLibrary
End Sub

Private Sub nextArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(nextArt, x, Y) Then
        NextProc picArt, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End If
End Sub

Private Sub nextArtistURL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(nextArtistURL, x, Y) Then
        NextProc txtArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    End If
End Sub

Private Sub nextCommercialInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(nextCommercialInfo, x, Y) Then
        NextProc txtCommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    End If
End Sub

Private Sub picArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If WithinBounds(picArt, x, Y) Then
        Select Case Button
            Case vbLeftButton: ImageBrowse
            Case vbRightButton: If ValidateMenu Then PopupMenu mnuArt
        End Select
    End If
End Sub

Private Sub prevArt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(prevArt, x, Y) Then
        PrevProc picArt, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End If
End Sub

Private Sub prevArtistURL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(prevArtistURL, x, Y) Then
        PrevProc txtArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    End If
End Sub

Private Sub prevCommercialInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton And WithinBounds(prevCommercialInfo, x, Y) Then
        PrevProc txtCommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    End If
End Sub

Private Sub TabStrip1_Click()
    On Error Resume Next
    Dim i As Long
    For i = 1 To TabStrip1.Tabs.Count
        If Frame2(i - 1).Visible <> TabStrip1.Tabs(i).Selected Then
            Frame2(i - 1).Visible = TabStrip1.Tabs(i).Selected
        End If
    Next
    ShowOrHideNecessaryFields
End Sub



Private Sub TreeFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.MousePointer = 1
TreeFiles.MousePointer = 1
End Sub

Private Sub TreeFiles_NodeClick(ByVal Node As MSComctlLib.Node)
Dim Color As Single
Dim sAlbum As String
Dim sArtist As String
Dim sGenre As String
Dim sSQl As String
Dim stipo As String
Dim aEle() As String
Dim scampos As String
Dim sWhere As String
On Error Resume Next

'Stipo FOR SELECTED  KEY IN TREEVIEW
  stipo = Left(TreeFiles.SelectedItem.Key, 5)
    
  If stipo <> "kPlaE" And stipo <> "kRecI" And stipo <> "kRecP" And stipo <> "kTopH" And stipo <> "FL FO" And stipo <> "CDMCA" And stipo <> "kAll" And stipo <> "A  AL" And stipo <> "AA AR" And stipo <> "AA AL" And stipo <> "FL CA" And stipo <> "GAAGE" Then Exit Sub
    
  scampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
    
  If stipo = "kAll" Then
    sWhere = "WHERE ONCD=FALSE"
    sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK ON ALBUMS
  If stipo = "A  AL" Then
     sAlbum = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND ALBUM='" & sAlbum & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK ON ARTIST
  If stipo = "AA AR" Then
     sArtist = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
    ' sArtist = aEle(1)
  
  '//CLICK ON ARTIST - ALBUM
  If stipo = "AA AL" Then
     aEle = Split(TreeFiles.SelectedItem.Key, "|", , vbTextCompare)
     sAlbum = aEle(2)
     sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "' AND ALBUM='" & sAlbum & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '// CLICK ON FILE LOCATION
  If stipo = "FL CA" Or stipo = "FL FO" Then
     If TreeFiles.SelectedItem.Children = 0 Then
        sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
        sGenre = Left(sGenre, Len(sGenre) - 1)
        sWhere = "WHERE ONCD=FALSE AND FILEPATH='" & sGenre & "'"
        sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
     
     End If
  End If
  
  '// CD MEDIA
  If stipo = "CDMCA" Then
     If TreeFiles.SelectedItem.Children = 0 Then
        sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
        sGenre = Left(sGenre, Len(sGenre) - 1)
        sWhere = "WHERE ONCD=TRUE AND FILEPATH='" & sGenre & "'"
        sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
     Else
        If Len(TreeFiles.SelectedItem.Key) = 8 Then
           sGenre = Right(TreeFiles.SelectedItem.Key, 3)
           sWhere = "WHERE ONCD=TRUE AND DRIVE='" & sGenre & "'"
           sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
        End If
     End If
  End If

  
   '//CLICK ON GENRES
  If stipo = "GAAGE" Then
     sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND GENRE='" & sGenre & "'"
     sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK ON TOP HITS
  If stipo = "kTopH" Then
    sWhere = "WHERE ONCD=FALSE AND PLAYCOUNT>0 ORDER BY PLAYCOUNT DESC "
      'sSQL = "SELECT TOP 20 PLAYCOUNT," & sCampos & " FROM MUSIC " & sWhere
    sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
   End If
  
     
  '//CLICK ON RECENTLY PLAYED
  If stipo = "kRecP" Then
      sWhere = "WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL " & "ORDER BY PLAYEDLAST DESC"
      'sSQL = "SELECT TOP 10 PLAYEDLAST, " & sCampos & " FROM MUSIC " & sWhere
      sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  '*****iMPORTANT*****SPACE in sql SATAEMENET IS VERY IMP. "SELECT " IS CORRECT BUT "SELECT"IS WRONG
  '*****"SELECT TOP 20 LASTUPDATE, " IS CORRECT "SELECT TOP 20 LASTUPDATE," WRONG
    
  '//CLICK ON RECENTLY ADDED
  If stipo = "kRecI" Then
      sWhere = "WHERE ONCD=FALSE AND LASTUPDATE IS NOT NULL " & "ORDER BY LASTUPDATE DESC"
      'sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
      ' sSQL = "SELECT TOP 20 LASTUPDATE, " & sCampos & " FROM MUSIC " & sWhere 'CORRECT STATEMENT TO GET TOP 20 ENTRIES
      scampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,LASTUPDATE,FILE"
      sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK EN PLAYLISTS
  If stipo = "kPlaE" Then
      sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
      'Cargar_PlayListTracks sGenre
      Exit Sub
  End If

 BindToSQL sSQl
     
 rs.Close
 If stipo = "kTopH" Then sWhere = "WHERE ONCD=FALSE AND PLAYCOUNT>0 "
 If stipo = "kRecP" Then sWhere = "WHERE ONCD=FALSE AND PLAYEDLAST IS NOT NULL "
 
 'On Error Resume Next
 rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO FROM MUSIC " & sWhere, cnnMusic, adOpenForwardOnly, adLockReadOnly
   
    '// UPDATE STATUS BAR
 Dim lKilobytes As Long, lSeconds As Long
 lKilobytes = CLng(rs!Total / 1024)
 lSeconds = CLng(rs!TIEMPO)
 rs.Close
 Call UpdateStatusBar(lKilobytes, lSeconds)
End Sub

Private Sub txtArtistURL_Change()
    TextProc txtArtistURL, S_AURL, countArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
End Sub

Private Sub txtCommercialInfo_Change()
    TextProc txtCommercialInfo, S_CURL, countCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
End Sub

Private Sub VScroll1_Change()
    Dim FTop As Single: FTop = -CSng(VScroll1.Value) * 360
    Dim FTopMax As Single: FTopMax = -Frame3.Height + Frame2(1).Height
    
    If (VScroll1.Value = VScroll1.Max And FTop > FTopMax) Or FTop < FTopMax Then
        If Frame3.Top <> FTopMax Then Frame3.Top = FTopMax
    Else
        If Frame3.Top <> FTop Then Frame3.Top = FTop
    End If
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    Dim hBitmap As Long
   Dim i As Integer

  
    Dim t As Long
    Dim strT As String
    
    bInitialized = False
    
    XSize = GetSetting(Caption, "Window", "XSize", Width)
    If Err Then
        XSize = Width
        Err.Clear
    End If
    
    YSize = GetSetting(Caption, "Window", "YSize", Height)
    If Err Then
        YSize = Height
        Err.Clear
    End If
    
    XPos = GetSetting(Caption, "Window", "XPos", Left)
    If Err Then
        XPos = Left
        Err.Clear
    End If
    
    YPos = GetSetting(Caption, "Window", "YPos", Top)
    If Err Then
        YPos = Top
        Err.Clear
    End If
    
    myWindowState = GetSetting(Caption, "Window", "State", vbNormal)
    If Err Then
        myWindowState = vbNormal
        Err.Clear
    End If
    
    If XSize < 568 * Screen.TwipsPerPixelX Then XSize = 568 * Screen.TwipsPerPixelX
    If XSize > Screen.Width Then XSize = Screen.Width
    
    If YSize < 445 * Screen.TwipsPerPixelY Then YSize = 445 * Screen.TwipsPerPixelY
    If YSize > Screen.Height Then YSize = Screen.Height
    
    If XPos < 0 Then XPos = 0
    If XPos > Screen.Width - Width Then XPos = Screen.Width - Width
    
    If YPos < 0 Then YPos = 0
    If YPos > Screen.Height - Height Then YPos = Screen.Height - Height
    
    If myWindowState <> vbNormal And myWindowState <> vbMaximized Then _
        myWindowState = vbNormal
    
    If Width <> XSize Then _
        Width = XSize
    If Height <> YSize Then _
        Height = YSize
    
    If Left <> XPos Then _
        Left = XPos
    If Top <> YPos Then _
        Top = YPos
    
    If WindowState <> myWindowState Then _
        WindowState = myWindowState
    
    bInitialized = True
    
    With ListView1
        .ColumnHeaderIcons = ImageList1
        
        .SortKey = GetSetting(Caption, "Columns", "SortKey", 0)
        If Err Then
            Err.Clear
            .SortKey = 0
        End If
        
        .SortOrder = GetSetting(Caption, "Columns", "SortOrder", lvwAscending)
        If Err Then
            Err.Clear
            .SortOrder = lvwAscending
        End If
        Resort = False
        ShowListViewColumnHeaderSortIcon ListView1
        

    End With
    
   
    strT = GetSetting(Caption, "MP3s", "Directory", GetSpecialFolderLocation(CSIDL_PERSONAL))
    If Right$(strT, 1) = "\" Then strT = Left$(strT, Len(strT) - 1)
    
    If Dir$(strT & "\") = "" Then
        strT = GetSpecialFolderLocation(CSIDL_PERSONAL)
    End If
    
    Text1 = strT
    Timer1.Enabled = True
    
    
  Set cnnMusic = New ADODB.Connection

  With cnnMusic
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source") = App.Path & "\Library\music.mdb"
    '.Properties("Jet OLEDB:Database Password") = "Licenciao159"
    .CursorLocation = adUseClient
    .Open
  End With
   Set cmd = New ADODB.Command
   
  cmd.ActiveConnection = cnnMusic
  'Treefiles.SelectedItem.Index
  'TreeFiles_Click
  LoadLibrary
 Dim scampos As String
Dim sWhere As String
Dim sSQl As String
ListView1.ListItems.Clear

scampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
sWhere = "WHERE ONCD=FALSE"
sSQl = "SELECT " & scampos & " FROM MUSIC " & sWhere
BindToSQL sSQl

  gHW = hWnd
  Hook
    
End Sub

Private Sub ListView1_DblClick()
  Dim sFile As String
  'If ListView1.ListItems.Count < 1 Then Exit Sub
  If SelectedIndex(ListView1) <> -1 Then
        sFile = ListView1.SelectedItem.SubItems(8)
        ShellExecute 0&, "open", sFile, vbNullString, vbNullString, SW_SHOW
        UpdatePlaycount (sFile)
  End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
   Dim ID3 As New clsID3
    Dim HourPart As String
    Dim tempStr As String
    Dim sGenreID As String
    Dim bResort As Boolean
    Dim sItem As ListItem
    Dim bRefresh As Boolean
    Dim idx As Long
    
    bResort = False
    bRefresh = False
    Set sItem = ListView1.SelectedItem
   
    With ID3
        .FileName = sItem.SubItems(8)
        ID3Revision = .ID3RevisionV2
        ShowOrHideNecessaryFields
        txtTitle = .Title
        txtArtist = .Artist
        txtAlbum = .Album
        cmbGenre = FormatGenre(ID3, .GenreID, .Genre)
        txtTrackNumber = .TrackNumber
        txtTracksTotal = .TracksTotal
        txtYear = .Year
        txtComments = .Comments
        txtLyrics = .Lyrics
        txtComposer = .Composer
        txtBand = .Band
        txtConductor = .Conductor
        txtInterpretedBy = .InterpretedBy
        txtLyricist = .Lyricist
        txtOriginalArtist = .OriginalArtist
        txtOriginalAlbum = .OriginalAlbum
        txtOriginalFileName = .OriginalFileName
        txtOriginalLyricist = .OriginalLyricist
        txtOriginalReleaseYear = .OriginalReleaseYear
        txtCopyright = .Copyright
        txtFileOwner = .FileOwner
        txtPublisher = .Publisher
        txtInternetRadioStationName = .InternetRadioStationName
        txtInternetRadioStationOwner = .InternetRadioStationOwner
        txtISRC = .ISRC
        txtLanguages = .Languages
        LoadMultiData txtArtistURL, .ArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
        txtCopyrightInfo = .CopyrightInfo
        txtAudioURL = .AudioURL
        LoadMultiData txtCommercialInfo, .CommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
        txtAudioSourceURL = .AudioSourceURL
        txtInternetRadioStationURL = .InternetRadioURL
        txtPaymentURL = .PaymentURL
        txtPublisherURL = .PublisherURL
        txtEncodedBy = .EncodedBy
        txtBPM = .BeatsPerMinute
        cmbKey = .InitialKey
        txtDiscNumber = .DiscNumber
        txtDiscsTotal = .DiscsTotal
        LoadMultiData picArt, .AttachedPictures, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End With
    
   
    If bRefresh Then EnsureSelVisible ListView1, True

StatusBar1.Panels(1).Text = "RECORD:[" + str(ListView1.SelectedItem.Index) + "] " + ListView1.SelectedItem
End Sub

Private Sub ShowOrHideNecessaryFields()
    Dim bShow As Boolean
    Dim lAdd As Long
    
    If Frame2(1).Visible And Frame3.Visible Then
        If ID3Revision > 2 Then
            bShow = True
            lAdd = 360
        Else
            bShow = False
            lAdd = -360
        End If
        
        ' Hide the text fields not supported by ID3v2.0 and ID3v2.2
        
        If txtFileOwner.Visible <> bShow Then
            Label27.Visible = bShow
            txtFileOwner.Visible = bShow
            VMove lAdd, Label38, txtPublisher, _
                        Label28, txtInternetRadioStationName, _
                        Label29, txtInternetRadioStationOwner, _
                        Label30, txtISRC, _
                        Label16, txtLanguages, _
                        Label31, txtCommercialInfo, _
                            countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, _
                        Label32, txtCopyrightInfo, _
                        Label21, txtAudioURL, _
                        Label33, txtArtistURL, _
                            countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, _
                        Label34, txtAudioSourceURL, _
                        Label35, txtInternetRadioStationURL, _
                        Label22, txtPaymentURL, _
                        Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        If txtInternetRadioStationName.Visible <> bShow Then
            Label28.Visible = bShow
            txtInternetRadioStationName.Visible = bShow
            VMove lAdd, Label29, txtInternetRadioStationOwner, _
                        Label30, txtISRC, _
                        Label16, txtLanguages, _
                        Label31, txtCommercialInfo, _
                            countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, _
                        Label32, txtCopyrightInfo, _
                        Label21, txtAudioURL, _
                        Label33, txtArtistURL, _
                            countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, _
                        Label34, txtAudioSourceURL, _
                        Label35, txtInternetRadioStationURL, _
                        Label22, txtPaymentURL, _
                        Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        If txtInternetRadioStationOwner.Visible <> bShow Then
            Label29.Visible = bShow
            txtInternetRadioStationOwner.Visible = bShow
            VMove lAdd, Label30, txtISRC, _
                        Label16, txtLanguages, _
                        Label31, txtCommercialInfo, _
                            countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, _
                        Label32, txtCopyrightInfo, _
                        Label21, txtAudioURL, _
                        Label33, txtArtistURL, _
                            countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, _
                        Label34, txtAudioSourceURL, _
                        Label35, txtInternetRadioStationURL, _
                        Label22, txtPaymentURL, _
                        Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        If txtPaymentURL.Visible <> bShow Then
            Label22.Visible = bShow
            txtPaymentURL.Visible = bShow
            VMove lAdd, Label24, txtPublisherURL, _
                        Label23, txtEncodedBy, _
                        Label19, txtBPM, _
                        Label20, cmbKey, _
                        Label36, txtDiscNumber, _
                        Label37, txtDiscsTotal
            Frame3.Height = Frame3.Height + lAdd
        End If
        
        AdjustVScrollProps
    End If
End Sub


Private Function FormatGenre(ByVal ID3Class As clsID3, ByVal GenreID As GenreConstants, ByVal Genre As String) As String
    If (GenreID = OtherGenre Or GenreID = Unknown) And Genre <> "" Then
        FormatGenre = Genre
    Else
        FormatGenre = ID3Class.GenreName(GenreID)
    End If
End Function

Private Function FormatTime(ByVal TimeVal As Double, Optional ByVal StoreTime As Boolean = False) As String
    On Error Resume Next
    
    Dim tv As Double
    Dim hr As Double
    Dim min As Double
    Dim sec As Double
    Dim ts As String
    
    tv = TimeVal
    If tv <= 0 Then
        If StoreTime Then dDuration = 0
    Else
        If StoreTime Then dDuration = tv
        
        tv = Fix(tv)
        min = Fix(tv / 60)
        sec = tv - 60 * min
        hr = Fix(min / 60)
        min = min - 60 * hr
        
        ts = ":" & Format$(sec, "00")
        If hr > 0 Then
            ts = CStr(hr) & ":" & Format$(min, "00") & ts
        Else
            ts = CStr(min) & ts
        End If
        
        FormatTime = ts
    End If
End Function

Private Function FormatBitRate(ByVal BitRate As Double, ByVal Encoding As EncodingEnum, Optional ByVal StoreBitRate As Boolean = False) As String
    On Error Resume Next
    
    Dim br As Double
    br = BitRate
    If br <= 0 Then
        If StoreBitRate Then dBitRate = 0
    Else
        If StoreBitRate Then dBitRate = br
        FormatBitRate = CStr(Fix(br / 1000)) & " kbps " & IIf(Encoding = CBR, "CBR", "VBR")
    End If
End Function

Private Sub RemoveAPICItem(ByVal Index As Long)
    Dim i As Long
    
    cAPICIType.Remove Index
    cAPICType.Remove Index
    cAPICData.Remove Index
    APICData.Remove cAPIC0(Index)
    cAPIC0.Remove Index
    
    For i = Index To cAPIC0.Count
        SetItem cAPIC0, i, cAPIC0(i) - 1
    Next
End Sub

Private Sub AddAPICItem(ByVal MIMEType As String, ByVal PictureType As PictureType, ByVal Data As String)
    Dim APD As APicDecoder
    
    If Data = "" Then
        cAPICIType.Add ""
        cAPICType.Add ""
        cAPICData.Add ""
    Else
        cAPICIType.Add MIMEType
        cAPICType.Add PictureType
        cAPICData.Add Data
    End If
    APICData.Add ""
    
    If Data <> "" Then
        Set APD = New APicDecoder
        APD.InsertImageData APICData, APICData.Count, MIMEType, PictureType, Data, ID3Revision
        Set APD = Nothing
    End If
    
    cAPIC0.Add APICData.Count
End Sub

Private Sub MakeNecessaryChanges(ByVal Index As Long)
    Dim APD As APicDecoder
    Dim GPC As GDIPlusCandy
    
    Dim MIMEType As String
    Dim PictureType As PictureType
    Dim Pic As StdPicture
    Dim PicData As String
    
    Set APD = New APicDecoder
    APD.DecodeImage APICData, cAPIC0(Index), MIMEType, PictureType, Pic, ID3Revision
    If MIMEType = cAPICIType(Index) Then
        PicData = cAPICData(Index)
    Else
        Set GPC = New GDIPlusCandy
        PicData = GPC.ImageToData(Pic, cAPICIType(Index))
        Set GPC = Nothing
    End If
    APD.InsertImageData APICData, cAPIC0(Index), cAPICIType(Index), cAPICType(Index), PicData, ID3Revision
    Set APD = Nothing
End Sub

Private Function FilterEntry(ByVal Description As String, ByVal Filter As String) As String
    FilterEntry = Description & "|" & Filter & "|"
End Function

Private Sub ImageBrowse()
    Dim fn As String
    Dim f As Integer
    Dim st As String
    Dim sMIMEType As String
    Dim tMIMEType As String
    Dim sExt As String
    Dim GPC As GDIPlusCandy
    Dim sPic As StdPicture
    Dim i As Long
    Dim idx As Long
    Dim bConvertImage As Boolean
    
    If ListView1.ListItems.Count > 0 Then
        fn = ShowOpenDialog(hWnd, FilterEntry("All Supported Formats", FILTER_SUPPORTED) & FilterEntry("Windows Bitmap", FILTER_BMP) & FilterEntry("Graphics Interchange Format", FILTER_GIF) & FilterEntry("JPEG File Interchange Format", FILTER_JPEG) & FilterEntry("Portable Network Graphics", FILTER_PNG), "Select Image")
        If fn <> "" Then
            i = InStrRev(fn, ".")
            If i > 0 Then
                sExt = Mid$(LCase$(fn), i + 1)
                Select Case sExt
                    Case "bmp", "dib"
                        sMIMEType = ImageTypeFromIndex(0, ID3Revision)
                        idx = 1
                        If ID3Revision > 2 Then idx = 0
                        bConvertImage = (ID3Revision <= 2)
                    Case "gif"
                        sMIMEType = ImageTypeFromIndex(1, ID3Revision)
                        idx = 1
                        bConvertImage = (ID3Revision <= 2)
                    Case "jpeg", "jpg", "jpe", "jfif", "jfi", "jif"
                        sMIMEType = ImageTypeFromIndex(2, ID3Revision)
                        idx = 0
                        If ID3Revision > 2 Then idx = 2
                    Case "png"
                        sMIMEType = ImageTypeFromIndex(3, ID3Revision)
                        idx = 1
                        If ID3Revision > 2 Then idx = 3
                    Case Else
                        sMIMEType = ""
                        idx = -1
                End Select
                If idx <> -1 Then cmbImageType.ListIndex = idx
            Else
                sMIMEType = ""
            End If
            
            f = FreeFile
            Open fn For Binary Access Read Shared As #f
                st = Space$(LOF(f))
                Get #f, , st
            Close #f
            
            Set GPC = New GDIPlusCandy
            Set sPic = GPC.DataToImage(st)
            Set GPC = Nothing
            
            If Not sPic Is Nothing Then
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                
                ' As ID3v2.0 and ID3v2.2 allow only JPEG and PNG images, do the necessary conversion for BMP and GIF images
                If bConvertImage Then
                    Set GPC = New GDIPlusCandy
                    st = GPC.ImageToData(sPic, ImagePNG)
                    Set sPic = GPC.DataToImage(st) ' Show the converted image
                    Set GPC = Nothing
                End If
                
                tMIMEType = DetermineImageType(st, ID3Revision)
                If sMIMEType <> tMIMEType And tMIMEType <> ImageUnsupported Then
                    sMIMEType = tMIMEType
                    cmbImageType.ListIndex = GetIndex(sMIMEType, ID3Revision)
                End If
                ArtAddProc sMIMEType, sPic, st
            End If
        End If
    End If
End Sub

Private Sub TextProc(Ctl As Object, ByVal Description As String, CountControl As Label, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    If Index = Total Then
        If Ctl = "" Then
            If Index > 0 Then
                FrameBlank = True
                Col.Remove Index
                Index = Index - 1
                Total = Total - 1
                Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                NextControl.ToolTipText = ""
                If Index = 0 Then
                    Set DelControl.Picture = Buttons.ListImages(I_DELI).Picture
                    DelControl.ToolTipText = ""
                End If
            End If
        Else
            If FrameBlank Then
                FrameBlank = False
                Col.Add Ctl.Text
                Index = Index + 1
                Total = Total + 1
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            Else
                If Index > 0 Then SetItem Col, Index, Ctl.Text
            End If
        End If
    Else
        If Index > 0 Then SetItem Col, Index, Ctl.Text
    End If
    CountControl = CStr(Index) & "/" & CStr(Total)
End Sub

Private Sub ArtAddProc(ByVal MIMEType As String, ByVal Pic As StdPicture, ByVal Data As String)
    picArt.ToolTipText = S_APICTT
    imgArt.ToolTipText = S_APICTT
    imgArt.Visible = True
    Set imgArt.Picture = Nothing
    StretchImage Pic
    Set imgArt.Picture = Pic
    SetBG True
    lblBrowse.Visible = False
    
    If indAPIC = totAPIC Then
        If bAPICBlank Then
            bAPICBlank = False
            AddAPICItem MIMEType, cmbPictureType.ListIndex, Data
            indAPIC = indAPIC + 1
            totAPIC = totAPIC + 1
            Set nextArt.Picture = Buttons.ListImages(I_ADD).Picture
            nextArt.ToolTipText = S_ADD & S_APIC
            Set delArt.Picture = Buttons.ListImages(I_DEL).Picture
            delArt.ToolTipText = S_DEL & S_APIC
        Else
            If indAPIC > 0 Then
                SetItem cAPICData, indAPIC, Data
                SetItem cAPICIType, indAPIC, MIMEType
                SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
            End If
        End If
    Else
        If indAPIC > 0 Then
            SetItem cAPICData, indAPIC, Data
            SetItem cAPICIType, indAPIC, MIMEType
            SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
        End If
    End If
    countArt = CStr(indAPIC) & "/" & CStr(totAPIC)
End Sub

Private Sub PrevProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean, Optional ByVal DeleteMode As Boolean = False)
    Dim bPic As Boolean, GDP As GDIPlusCandy, Pic As StdPicture
    bPic = (TypeName(Ctl) = "PictureBox")
    If (Index > 0 And DeleteMode) Or (Index > 1 And Not DeleteMode) Or (Index > 0 And Not DeleteMode And FrameBlank) Then
        If Index = Total Then
            Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
            DelControl.ToolTipText = S_DEL & Description
            If FrameBlank Then
IsBlank:
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
                FrameBlank = False
            Else
                Index = Index - 1
                If Index = 0 Then
                    GoTo IsBlank
                Else
                    Set NextControl.Picture = Buttons.ListImages(I_NEXT).Picture
                    NextControl.ToolTipText = S_NEXT & Description
                End If
            End If
        Else
            Index = Index - 1
        End If
        If Index <= 1 Then
            Set PrevControl.Picture = Buttons.ListImages(I_PREVI).Picture
            PrevControl.ToolTipText = ""
        End If
        If Index = 0 Then
            If bPic Then
                SetBG False
                Set imgArt.Picture = Nothing
                StretchImage imgArt.Picture
                imgArt.Visible = False
                cmbImageType.Enabled = False
                cmbPictureType.Enabled = False
                cmbImageType.ListIndex = 2 * (cmbImageType.ListIndex \ 4)
                cmbPictureType.ListIndex = 0
                lblBrowse.Visible = True
                Ctl.ToolTipText = ""
                imgArt.ToolTipText = ""
            Else
                Ctl = ""
            End If
            FrameBlank = True
        Else
            If bPic Then
                lblBrowse.Visible = False
                imgArt.Visible = True
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                Set GDP = New GDIPlusCandy
                Set Pic = GDP.DataToImage(Col(Index))
                Set GDP = Nothing
                Set imgArt.Picture = Nothing
                StretchImage Pic
                Set imgArt.Picture = Pic
                SetBG True
                cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                cmbPictureType.ListIndex = cAPICType(Index)
                Ctl.ToolTipText = S_APICTT
                imgArt.ToolTipText = S_APICTT
            Else
                Ctl = Col(Index)
            End If
        End If
        CountControl = CStr(Index) & "/" & CStr(Total)
    End If
End Sub

Private Sub NextProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    Dim bPic As Boolean, bBlank As Boolean, GDP As GDIPlusCandy, Pic As StdPicture, lIType As Long, vType As Variant
    Dim bRefresh As Boolean
    Dim blFrameBlank As Boolean
    bPic = (TypeName(Ctl) = "PictureBox")
    If Index < Total Then
        bRefresh = True
        Index = Index + 1
        If Index = Total Then
            Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
            NextControl.ToolTipText = S_ADD & Description
        End If
    Else
        If bPic Then
            If imgArt.Visible Then
                bRefresh = True
                blFrameBlank = True
                bBlank = True
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            End If
        Else
            If Ctl <> "" Then
                bRefresh = True
                blFrameBlank = True
                Index = Index + 1
                Col.Add ""
                Total = Col.Count
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            End If
        End If
    End If
    If bRefresh Then
        Set PrevControl.Picture = Buttons.ListImages(I_PREV).Picture
        PrevControl.ToolTipText = S_PREV & Description
        If bPic Then
            If bBlank Then
                cmbImageType.Enabled = False
                cmbPictureType.Enabled = False
                imgArt.Visible = False
                SetBG False
                Set imgArt.Picture = Nothing
                StretchImage imgArt.Picture
                cmbImageType.ListIndex = 2 * (cmbImageType.ListCount \ 4)
                cmbPictureType.ListIndex = 0
                lblBrowse.Visible = True
                Ctl.ToolTipText = ""
                imgArt.ToolTipText = ""
                Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                NextControl.ToolTipText = ""
            Else
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                imgArt.Visible = True
                Set GDP = New GDIPlusCandy
                Set Pic = GDP.DataToImage(Col(Index))
                Set GDP = Nothing
                Set imgArt.Picture = Nothing
                StretchImage Pic
                Set imgArt.Picture = Pic
                SetBG True
                cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                cmbPictureType.ListIndex = cAPICType(Index)
                lblBrowse.Visible = False
                Ctl.ToolTipText = S_APICTT
                imgArt.ToolTipText = S_APICTT
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
            End If
        Else
            Ctl = Col(Index)
        End If
        If blFrameBlank Then FrameBlank = True
        CountControl = CStr(Index) & "/" & CStr(Total)
    End If
End Sub

Private Sub DelProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    Dim bPic As Boolean, GDP As GDIPlusCandy, Pic As StdPicture
    bPic = (TypeName(Ctl) = "PictureBox")
    If Total > 0 Then
        If FrameBlank Then
            PrevProc Ctl, Description, CountControl, PrevControl, NextControl, DelControl, Col, Index, Total, FrameBlank, True
        Else
            If bPic Then
                RemoveAPICItem Index
            Else
                Col.Remove Index
            End If
            Total = Col.Count
            If Index > Total Then Index = Total
            If Index = 0 Then
                If bPic Then
                    cmbImageType.Enabled = False
                    cmbPictureType.Enabled = False
                    SetBG False
                    Set imgArt.Picture = Nothing
                    StretchImage imgArt.Picture
                    cmbImageType.ListIndex = 2 * (cmbImageType.ListIndex \ 4)
                    cmbPictureType.ListIndex = 0
                    imgArt.Visible = False
                    lblBrowse.Visible = True
                    Ctl.ToolTipText = ""
                    imgArt.ToolTipText = ""
                Else
                    Ctl = ""
                End If
                FrameBlank = True
            Else
                If bPic Then
                    cmbImageType.Enabled = True
                    cmbPictureType.Enabled = True
                    imgArt.Visible = True
                    Set GDP = New GDIPlusCandy
                    Set Pic = GDP.DataToImage(Col(Index))
                    Set GDP = Nothing
                    Set imgArt.Picture = Nothing
                    StretchImage Pic
                    Set imgArt.Picture = Pic
                    SetBG True
                    cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                    cmbPictureType.ListIndex = cAPICType(Index)
                    lblBrowse.Visible = False
                    Ctl.ToolTipText = S_APICTT
                    imgArt.ToolTipText = S_APICTT
                Else
                    Ctl = Col(Index)
                End If
            End If
            If Index = Total Then
                If Index = 0 Then
                    Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                    NextControl.ToolTipText = ""
                Else
                    Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                    NextControl.ToolTipText = S_ADD & Description
                End If
            End If
            If Index <= 1 Then
                Set PrevControl.Picture = Buttons.ListImages(I_PREVI).Picture
                PrevControl.ToolTipText = ""
                If Total = 0 Then
                    Set DelControl.Picture = Buttons.ListImages(I_DELI).Picture
                    DelControl.ToolTipText = ""
                End If
            End If
            CountControl = CStr(Index) & "/" & CStr(Total)
        End If
    End If
End Sub

Private Function WithinBounds(ByVal Obj As Object, ByVal x As Single, ByVal Y As Single) As Boolean
    Dim oWidth As Single, oHeight As Single
    If TypeName(Obj) = "PictureBox" Then
        oWidth = Obj.ScaleWidth
        oHeight = Obj.ScaleHeight
    Else
        oWidth = Obj.Width
        oHeight = Obj.Height
    End If
    WithinBounds = (x >= 0 And x <= oWidth And Y >= 0 And Y <= oHeight)
End Function

Private Sub ChangeFields(ByVal bEnabled As Boolean)

    Dim BG As Long
    BG = IIf(bEnabled, vbWindowBackground, vbButtonFace)

    If txtTitle.Enabled <> bEnabled Then
        txtTitle.Enabled = bEnabled
        txtTitle.BackColor = BG
    End If

    If txtArtist.Enabled <> bEnabled Then
        txtArtist.Enabled = bEnabled
        txtArtist.BackColor = BG
    End If

    If txtAlbum.Enabled <> bEnabled Then
        txtAlbum.Enabled = bEnabled
        txtAlbum.BackColor = BG
    End If

    If cmbGenre.Enabled <> bEnabled Then
        cmbGenre.Enabled = bEnabled
        cmbGenre.BackColor = BG
    End If

    If txtTrackNumber.Enabled <> bEnabled Then
        txtTrackNumber.Enabled = bEnabled
        txtTrackNumber.BackColor = BG
    End If

    If txtTracksTotal.Enabled <> bEnabled Then
        txtTracksTotal.Enabled = bEnabled
        txtTracksTotal.BackColor = BG
    End If

    If txtYear.Enabled <> bEnabled Then
        txtYear.Enabled = bEnabled
        txtYear.BackColor = BG
    End If

    If txtComments.Enabled <> bEnabled Then
        txtComments.Enabled = bEnabled
        txtComments.BackColor = BG
    End If

    If txtLyrics.Enabled <> bEnabled Then
        txtLyrics.Enabled = bEnabled
        txtLyrics.BackColor = BG
    End If

    If txtComposer.Enabled <> bEnabled Then
        txtComposer.Enabled = bEnabled
        txtComposer.BackColor = BG
    End If

    If txtBand.Enabled <> bEnabled Then
        txtBand.Enabled = bEnabled
        txtBand.BackColor = BG
    End If

    If txtConductor.Enabled <> bEnabled Then
        txtConductor.Enabled = bEnabled
        txtConductor.BackColor = BG
    End If

    If txtInterpretedBy.Enabled <> bEnabled Then
        txtInterpretedBy.Enabled = bEnabled
        txtInterpretedBy.BackColor = BG
    End If

    If txtLyricist.Enabled <> bEnabled Then
        txtLyricist.Enabled = bEnabled
        txtLyricist.BackColor = BG
    End If

    If txtOriginalArtist.Enabled <> bEnabled Then
        txtOriginalArtist.Enabled = bEnabled
        txtOriginalArtist.BackColor = BG
    End If

    If txtOriginalAlbum.Enabled <> bEnabled Then
        txtOriginalAlbum.Enabled = bEnabled
        txtOriginalAlbum.BackColor = BG
    End If

    If txtOriginalFileName.Enabled <> bEnabled Then
        txtOriginalFileName.Enabled = bEnabled
        txtOriginalFileName.BackColor = BG
    End If

    If txtOriginalLyricist.Enabled <> bEnabled Then
        txtOriginalLyricist.Enabled = bEnabled
        txtOriginalLyricist.BackColor = BG
    End If

    If txtOriginalReleaseYear.Enabled <> bEnabled Then
        txtOriginalReleaseYear.Enabled = bEnabled
        txtOriginalReleaseYear.BackColor = BG
    End If

    If txtCopyright.Enabled <> bEnabled Then
        txtCopyright.Enabled = bEnabled
        txtCopyright.BackColor = BG
    End If

    If txtFileOwner.Enabled <> bEnabled Then
        txtFileOwner.Enabled = bEnabled
        txtFileOwner.BackColor = BG
    End If

    If txtPublisher.Enabled <> bEnabled Then
        txtPublisher.Enabled = bEnabled
        txtPublisher.BackColor = BG
    End If

    If txtInternetRadioStationName.Enabled <> bEnabled Then
        txtInternetRadioStationName.Enabled = bEnabled
        txtInternetRadioStationName.BackColor = BG
    End If

    If txtInternetRadioStationOwner.Enabled <> bEnabled Then
        txtInternetRadioStationOwner.Enabled = bEnabled
        txtInternetRadioStationOwner.BackColor = BG
    End If

    If txtISRC.Enabled <> bEnabled Then
        txtISRC.Enabled = bEnabled
        txtISRC.BackColor = BG
    End If

    If txtLanguages.Enabled <> bEnabled Then
        txtLanguages.Enabled = bEnabled
        txtLanguages.BackColor = BG
    End If

    If txtCommercialInfo.Enabled <> bEnabled Then
        txtCommercialInfo.Enabled = bEnabled
        txtCommercialInfo.BackColor = BG
    End If

    If txtCopyrightInfo.Enabled <> bEnabled Then
        txtCopyrightInfo.Enabled = bEnabled
        txtCopyrightInfo.BackColor = BG
    End If

    If txtAudioURL.Enabled <> bEnabled Then
        txtAudioURL.Enabled = bEnabled
        txtAudioURL.BackColor = BG
    End If

    If txtArtistURL.Enabled <> bEnabled Then
        txtArtistURL.Enabled = bEnabled
        txtArtistURL.BackColor = BG
    End If

    If txtAudioSourceURL.Enabled <> bEnabled Then
        txtAudioSourceURL.Enabled = bEnabled
        txtAudioSourceURL.BackColor = BG
    End If

    If txtInternetRadioStationURL.Enabled <> bEnabled Then
        txtInternetRadioStationURL.Enabled = bEnabled
        txtInternetRadioStationURL.BackColor = BG
    End If

    If txtPaymentURL.Enabled <> bEnabled Then
        txtPaymentURL.Enabled = bEnabled
        txtPaymentURL.BackColor = BG
    End If

    If txtPublisherURL.Enabled <> bEnabled Then
        txtPublisherURL.Enabled = bEnabled
        txtPublisherURL.BackColor = BG
    End If

    If txtEncodedBy.Enabled <> bEnabled Then
        txtEncodedBy.Enabled = bEnabled
        txtEncodedBy.BackColor = BG
    End If

    If txtBPM.Enabled <> bEnabled Then
        txtBPM.Enabled = bEnabled
        txtBPM.BackColor = BG
    End If

    If cmbKey.Enabled <> bEnabled Then
        cmbKey.Enabled = bEnabled
        cmbKey.BackColor = BG
    End If

    If txtDiscNumber.Enabled <> bEnabled Then
        txtDiscNumber.Enabled = bEnabled
        txtDiscNumber.BackColor = BG
    End If

    If txtDiscsTotal.Enabled <> bEnabled Then
        txtDiscsTotal.Enabled = bEnabled
        txtDiscsTotal.BackColor = BG
    End If

End Sub

Private Sub VMove(ByVal By As Long, ParamArray Objs() As Variant)
    Dim i As Long
    For i = LBound(Objs) To UBound(Objs)
        Objs(i).Top = Objs(i).Top + By
    Next
End Sub

Private Sub AdjustVScrollProps()
    Dim VMax As Integer
    VMax = Ceiling((Frame3.Height - Frame2(1).Height) / 360) - 1
    If VScroll1.Max <> VMax Then
        VScroll1.Max = VMax
        If VScroll1.Max = 0 Then
            VScroll1.Visible = False
        Else
            VScroll1.Visible = True
        End If
    End If
    If VScroll1.Max = 0 Then
        Frame3.Width = Frame2(1).Width
    Else
        Frame3.Width = Frame2(1).Width - 255
    End If
    VScroll1_Change
End Sub


Private Sub LoadFileEntries(ByVal Path As String)
  On Error Resume Next
  Dim rst As New ADODB.Recordset
  Dim s As String
  Dim lSeconds As Long
  Dim sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String

    Dim ID3 As New clsID3
    Dim sPath As String
    Dim d As String
    Dim HourPart As String
    Dim BlankWCOM As New MultiFrameData
    Dim BlankWOAR As New MultiFrameData
    Dim BlankAPIC As New MultiFrameData
    
    sPath = Path
    If Right$(Path, 1) <> "\" Then sPath = sPath & "\"
    
    d = Dir$(sPath)
    
    ID3Revision = 3
    ShowOrHideNecessaryFields
    txtTitle = ""
    txtArtist = ""
    txtAlbum = ""
    cmbGenre = ""
    txtTrackNumber = ""
    txtTracksTotal = ""
    txtYear = ""
    txtComments = ""
    txtLyrics = ""
    txtComposer = ""
    txtBand = ""
    txtConductor = ""
    txtInterpretedBy = ""
    txtLyricist = ""
    txtOriginalArtist = ""
    txtOriginalAlbum = ""
    txtOriginalFileName = ""
    txtOriginalLyricist = ""
    txtOriginalReleaseYear = ""
    txtCopyright = ""
    txtFileOwner = ""
    txtPublisher = ""
    txtInternetRadioStationName = ""
    txtInternetRadioStationOwner = ""
    txtISRC = ""
    txtLanguages = ""
    LoadMultiData txtCommercialInfo, BlankWCOM, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
    txtCopyrightInfo = ""
    txtAudioURL = ""
    LoadMultiData txtArtistURL, BlankWOAR, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
    txtAudioSourceURL = ""
    txtInternetRadioStationURL = ""
    txtPaymentURL = ""
    txtPublisherURL = ""
    txtEncodedBy = ""
    txtBPM = ""
    cmbKey = ""
    txtDiscNumber = ""
    txtDiscsTotal = ""
    LoadMultiData picArt, BlankAPIC, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    lblBrowse.Visible = False
    
    ChangeFields False
    If Command1.Enabled Then Command1.Enabled = False
    If Command2.Enabled Then Command2.Enabled = False
    If Command3.Enabled Then Command3.Enabled = False
    If Command4.Enabled Then Command4.Enabled = False
    If Command5.Enabled Then Command5.Enabled = False
    If Command6.Enabled Then Command6.Enabled = False
    
    Do Until d = ""
        If d <> "." And d <> ".." Then
            If LCase$(Right$(d, 4)) = ".mp3" Then
                With ListView1
                    If MousePointer = vbDefault Then
                        MousePointer = vbHourglass
                        DoEvents
                    End If
                   ' .ListItems.Add Text:=d
                     ID3.FileName = sPath & d
                     rst.Open "SELECT * FROM Music", cnnMusic, adOpenDynamic, adLockOptimistic
                     rst.AddNew
                     sTitle = ID3.Title
                     If sTitle = "" Then
                       sTitle = GetFileTitle(sPath & d)
                       sTitle = Left(sTitle, Len(sTitle) - 4)
                     End If
                     rst!File = sPath & d
                     rst!Title = sTitle
                     rst!Artist = ID3.Artist
                     If ID3.Artist = "" Then rst!Artist = "Unknown"
                     rst!Album = ID3.Album
                     If ID3.Album = "" Then rst!Album = "Unknown"
                     rst!Year = ID3.Year
                     rst!Genre = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
                     rst!Comments = ID3.Comments
                     rst!length = FormatTime(ID3.length)
                     rst!BYTES = FileLen(sPath & d)
                     rst!Seconds = ID3.length
                     rst!PlayCount = 0
                     rst!FilePath = Left(sPath & d, InStrRev(sPath & d, "\") - 1)
                     rst!Drive = Left(sPath & d, 3)
                     rst.Update
                     rst.Close
                     
                     Dim Item
                     Set Item = ListView1.ListItems.Add(, , sTitle)
                     Item.SubItems(1) = ID3.Artist
                     Item.SubItems(2) = ID3.Album
                     If ID3.Album = "" Then Item.SubItems(2) = "Unknown"
                     If ID3.Artist = "" Then Item.SubItems(1) = "Unknown"
                     Item.SubItems(3) = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
                     Item.SubItems(4) = ID3.Year
                     Item.SubItems(5) = FormatTime(ID3.length)
                     Item.SubItems(6) = ""
                     Item.SubItems(7) = ""
                     Item.SubItems(8) = sPath & d
                End With
            End If
        End If
        d = Dir$
    Loop
    
    Resort = True
    'SortLvwOnLong ListView1, ListView1.SortKey + 1
    Resort = False
    
    
    If MousePointer = vbHourglass Then _
       MousePointer = vbDefault
    
    If Not Command1.Enabled Then Command1.Enabled = True
    If Not Command5.Enabled Then Command5.Enabled = True
    
    If ListView1.ListItems.Count > 0 Then
        ChangeFields True
        If Not Command2.Enabled Then Command2.Enabled = True
        If Not Command3.Enabled Then Command3.Enabled = True
        If Not Command4.Enabled Then Command4.Enabled = True
        If Not Command6.Enabled Then Command6.Enabled = True
        ListView1.ListItems(1).Selected = True
        ListView1_ItemClick ListView1.ListItems(1)
    Else
        ChangeFields False
        If Command2.Enabled Then Command2.Enabled = False
        If Command3.Enabled Then Command3.Enabled = False
        If Command4.Enabled Then Command4.Enabled = False
        If Command6.Enabled Then Command6.Enabled = False
    End If
End Sub


Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        If bInitialized Then myWindowState = WindowState
        'Text1.Width = ScaleWidth - 1560
        'Command1.Left = ScaleWidth - 1335
        'Command5.Left = ScaleWidth - 840
        PicClientArea.Width = ScaleWidth - 225
        PicClientArea.Height = 4 * ScaleHeight \ 5 - 2536
        Text1.Width = PicClientArea.ScaleWidth - 18 - CmdSearch.Width
        CmdSearch.Left = PicClientArea.ScaleWidth - CmdSearch.Width - 6
        ListView1.Width = PicClientArea.ScaleWidth - ListView1.Left - 6
        ListView1.Height = PicClientArea.ScaleHeight - 35 ' ScaleHeight - 800 '4 * ScaleHeight \ 5 - 2961
        TreeFiles.Height = ListView1.Height
        Frame1.Width = ScaleWidth - 225
        Frame1.Top = 4 * ScaleHeight \ 5 - 2526
        Frame1.Height = ScaleHeight \ 5 + 2271
        Command2.Left = Frame1.Width - 3135
        Command2.Top = Frame1.Height - 495
        Command3.Left = Frame1.Width - 1575
        Command3.Top = Frame1.Height - 495
        Command4.Top = Frame1.Height - 495
        Command6.Top = Frame1.Height - 495
        TabStrip1.Width = Frame1.Width - 240
        TabStrip1.Height = Frame1.Height - 840
        Frame2(0).Width = Frame1.Width - 480
        Frame2(0).Height = Frame1.Height - 1440
        Frame2(1).Width = Frame1.Width - 480
        Frame2(1).Height = Frame1.Height - 1440
        Frame2(2).Width = Frame1.Width - 480
        Frame2(2).Height = Frame1.Height - 1440
        VScroll1.Height = Frame2(1).Height
        VScroll1.Left = Frame2(1).Width - 255
        AdjustVScrollProps
        txtTitle.Width = Frame2(0).Width - 600
        txtArtist.Width = Frame2(0).Width - 600
        txtAlbum.Width = Frame2(0).Width - 600
        txtComments.Width = Frame2(0).Width \ 2 - 1012
        txtComments.Height = Frame2(0).Height - 1440
        Label8.Left = Frame2(0).Width \ 2 + 173
        txtLyrics.Left = Frame2(0).Width \ 2 + 773
        txtLyrics.Height = Frame2(0).Height - 1440
        txtLyrics.Width = Frame2(0).Width \ 2 - 772
        txtComposer.Width = Frame3.Width - 1920
        txtBand.Width = Frame3.Width - 1920
        txtConductor.Width = Frame3.Width - 1920
        txtInterpretedBy.Width = Frame3.Width - 1920
        txtLyricist.Width = Frame3.Width - 1920
        txtOriginalArtist.Width = Frame3.Width - 1920
        txtOriginalAlbum.Width = Frame3.Width - 1920
        txtOriginalFileName.Width = Frame3.Width - 1920
        txtOriginalLyricist.Width = Frame3.Width - 1920
        txtOriginalReleaseYear.Width = Frame3.Width - 1920
        txtCopyright.Width = Frame3.Width - 1920
        txtFileOwner.Width = Frame3.Width - 1920
        txtPublisher.Width = Frame3.Width - 1920
        txtInternetRadioStationName.Width = Frame3.Width - 1920
        txtInternetRadioStationOwner.Width = Frame3.Width - 1920
        txtISRC.Width = Frame3.Width - 1920
        txtLanguages.Width = Frame3.Width - 1920
        txtCommercialInfo.Width = Frame3.Width - 3165
        countCommercialInfo.Left = Frame3.Width - 1335
        prevCommercialInfo.Left = Frame3.Width - 750
        nextCommercialInfo.Left = Frame3.Width - 540
        delCommercialInfo.Left = Frame3.Width - 330
        txtCopyrightInfo.Width = Frame3.Width - 1920
        txtAudioURL.Width = Frame3.Width - 1920
        txtArtistURL.Width = Frame3.Width - 3165
        countArtistURL.Left = Frame3.Width - 1335
        prevArtistURL.Left = Frame3.Width - 750
        nextArtistURL.Left = Frame3.Width - 540
        delArtistURL.Left = Frame3.Width - 330
        txtAudioSourceURL.Width = Frame3.Width - 1920
        txtInternetRadioStationURL.Width = Frame3.Width - 1920
        txtPaymentURL.Width = Frame3.Width - 1920
        txtPublisherURL.Width = Frame3.Width - 1920
        txtEncodedBy.Width = Frame3.Width - 1920
        picArt.Left = Frame2(2).Width \ 2 - (Frame2(2).Height - 840) \ 2
        picArt.Width = Frame2(2).Height - 840 ' Most album art is square
        picArt.Height = Frame2(2).Height - 840
        StretchImage imgArt.Picture
        lblBrowse.Width = picArt.ScaleWidth
        lblBrowse.Top = picArt.Height \ 2 - 247
        Label41.Left = Frame2(2).Width \ 2 - 1800
        Label41.Top = Frame2(2).Height - 735
        cmbImageType.Left = Frame2(2).Width \ 2 - 720
        cmbImageType.Top = Frame2(2).Height - 735
        countArt.Left = Frame2(2).Width \ 2 + 585
        countArt.Top = Frame2(2).Height - 675
        prevArt.Left = Frame2(2).Width \ 2 + 1170
        prevArt.Top = Frame2(2).Height - 735
        nextArt.Left = Frame2(2).Width \ 2 + 1380
        nextArt.Top = Frame2(2).Height - 735
        delArt.Left = Frame2(2).Width \ 2 + 1590
        delArt.Top = Frame2(2).Height - 735
        Label43.Left = Frame2(2).Width \ 2 - 1800
        Label43.Top = Frame2(2).Height - 375
        cmbPictureType.Left = Frame2(2).Width \ 2 - 720
        cmbPictureType.Top = Frame2(2).Height - 375
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    SaveSetting Caption, "Window", "XSize", XSize
    SaveSetting Caption, "Window", "YSize", YSize
    SaveSetting Caption, "Window", "XPos", XPos
    SaveSetting Caption, "Window", "YPos", YPos
    SaveSetting Caption, "Window", "State", myWindowState
    
    With ListView1
        SaveSetting Caption, "Columns", "SortKey", .SortKey
        SaveSetting Caption, "Columns", "SortOrder", .SortOrder
        
        For i = 1 To .ColumnHeaders.Count
            SaveSetting Caption, "ColumnPos", Format$(.ColumnHeaders(i).Position, "00"), i
            SaveSetting Caption, "Columns", Format$(i, "00"), .ColumnHeaders(i).Width
        Next
    End With
    
    SaveSetting Caption, "MP3s", "Directory", Text1.Text
    Unhook
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Long
   Dim idx As Long
   On Error Resume Next
   SortLvwOnLong Me.ListView1, ColumnHeader.Index
   ShowListViewColumnHeaderSortIcon ListView1
   EnsureSelVisible ListView1
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim FileName As String
Dim Extension As String
Dim bfilefound As Boolean
Dim tt, icnt, iPos As Integer
If Effect <> 7 Then Exit Sub
For icnt = 1 To Data.Files.Count
  'If Dir(Data.Files(icnt)) <> "" Then
            'This function will add the index of this file added to the listview in order to
            'create a sequence of playback in normal sequential mode or shuffle mode
            iPos = InStrRev(Data.Files(icnt), ".")
            Extension = Mid(Data.Files(icnt), iPos + 1, Len(Data.Files(icnt)) - iPos)
        If UCase(Extension) = "MP3" Or UCase(Extension) = "MP2" Or UCase(Extension) = "MP1" Or UCase(Extension) = "WAV" Then
             iPos = InStrRev(Data.Files(icnt), "\")
             FileName = (Mid(Data.Files(icnt), iPos + 1, Len(Data.Files(icnt)) - iPos - 4))
             Call AddTrack(Data.Files(icnt))
        Else
             LoadFileEntries (Data.Files(icnt))
        End If
    
  'End If
1:
Next icnt
LoadLibrary
End Sub

Private Sub PicClientArea_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Me.MousePointer = vbDefault
 
 If Button = vbLeftButton And PicClientArea.MousePointer = 9 Then
  If x > 100 + TreeFiles.Left And x < TreeFiles.Left + PicClientArea.ScaleWidth - 100 Then
   TreeFiles.Width = x - TreeFiles.Left - 2
  ' ListView1.Left = X + 1
 '  ListView1.Width = PicClientArea.ScaleWidth - ListView1.Left - 6
   ListView1.Move x + 1, ListView1.Top, PicClientArea.ScaleWidth - ListView1.Left - 6, ListView1.Height
   'Form_Resize
  End If
End If
 If x > TreeFiles.Left + TreeFiles.Width And x < TreeFiles.Left + TreeFiles.Width + 6 Then
   PicClientArea.MousePointer = 9
 Else
   PicClientArea.MousePointer = vbDefault
 End If

End Sub
Private Sub PicClientArea_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If PicClientArea.MousePointer = 9 Then Form_Resize
End Sub
Public Sub LoadLibrary()
 Dim sClave As String
 Dim sLastNode As String
 Dim sLAlbum As String
 Dim sLArtist As String
 Dim sNode As String
 Dim sAlbum As String
 Dim sArtist As String
 Dim sGenre As String
 Dim rsFiles As New ADODB.Recordset
 
 'On Error GoTo HELL
  TreeFiles.Nodes.Clear
  TreeFiles.Nodes.Add , , "kAll", "Local Audio", 3
  TreeFiles.Nodes.Add , , "kMediaLibrary", "Media Library", 5
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kAlbum", "Album", 6
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kArtist\Album", "Artist\Album", 7
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kCDMedia", "CD Media", 8
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kPath", "File Location", 9
  TreeFiles.Nodes.Add "kPath", tvwChild, "kFullPath", "Full Path", 9
  TreeFiles.Nodes.Add "kPath", tvwChild, "kFolder", "by Folder", 2
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kGenre", "Genre", 13
  TreeFiles.Nodes.Add , , "kPlayList", "Play List", 14
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kTopHits", "Top Hits", 12
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kRecP", "Recently Played", 11
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kRecI", "Recently Imported", 10
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kPlaL", "Play Lists", 14
  
  
 cmd.CommandText = "SELECT DISTINCT GENRE,ARTIST FROM MUSIC WHERE ONCD =FALSE ORDER BY GENRE"
 
 Set rsFiles = cmd.Execute
  
 Dim rsArtist As New ADODB.Recordset
 Dim rsAlbum As New ADODB.Recordset

''  // GENRE
    Do Until rsFiles.EOF
       sGenre = CStr(Trim(rsFiles!Genre))
       If sGenre = "" Then sGenre = "Desconocido"
       
       If "GAAGE" & LCase(sGenre) <> sLastNode Then
           sLastNode = "GAAGE" & LCase(sGenre)
           TreeFiles.Nodes.Add "kGenre", tvwChild, sLastNode, sGenre, 13
       End If
       rsFiles.MoveNext
    Loop
  
  
 cmd.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE ORDER BY ALBUM"
 Set rsFiles = cmd.Execute
 sArtist = ""
 sAlbum = ""
  
''  // ALBUM
    Do Until rsFiles.EOF
       sAlbum = CStr(Trim(rsFiles!Album))
       If sAlbum = "" Then sAlbum = "Unknown"
       
       If "A  AL" & LCase(sAlbum) <> sLastNode Then
           sLastNode = "A  AL" & LCase(sAlbum)
           TreeFiles.Nodes.Add "kAlbum", tvwChild, sLastNode, sAlbum, 6
       End If
       rsFiles.MoveNext
    Loop

'  // ARTIST - ALBUMS

 
 cmd.CommandText = "SELECT DISTINCT ARTIST FROM MUSIC WHERE ONCD=FALSE ORDER BY ARTIST"
 Set rsFiles = cmd.Execute
   sLastNode = ""
   sLAlbum = ""
     Do Until rsFiles.EOF
       sArtist = CStr(Trim(rsFiles!Artist))
       If sArtist = "" Then sArtist = "Desconocido"
       If "AA AR" & LCase(sArtist) <> sLastNode Then
           sLastNode = "AA AR" & LCase(sArtist)
           TreeFiles.Nodes.Add "kArtist\Album", tvwChild, sLastNode, sArtist, 7
           cmd.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE AND ARTIST='" & rsFiles!Artist & "'"
           Set rsAlbum = cmd.Execute
           If rsAlbum.RecordCount > 1 Then
           Do Until rsAlbum.EOF
                sAlbum = CStr(Trim(rsAlbum!Album))
                If sAlbum = "" Then sAlbum = "Desconocido"
                If "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum) <> sLAlbum Then
                sLAlbum = "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum)
                TreeFiles.Nodes.Add sLastNode, tvwChild, sLAlbum, sAlbum, 6
                End If
                rsAlbum.MoveNext
           
           Loop
           End If
           rsAlbum.Close
       
       End If
       rsFiles.MoveNext
    Loop
    
 '// FILE LOCATION
 Dim sKey As String, s As String
 Dim sPath() As String

 On Error Resume Next

 cmd.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=FALSE"
 Set rsFiles = cmd.Execute
 sLastNode = ""
 sArtist = ""
 sAlbum = ""

 '// add albums folders
 Do Until rsFiles.EOF
    s = rsFiles!FilePath
        
    sPath = Split(s, "\", , vbTextCompare)
    TreeFiles.Nodes.Add "kFolder", tvwChild, "FL FO" & CStr(s & "\"), sPath(UBound(sPath)), 2
        
    If sLastNode <> sPath(0) Then
       TreeFiles.Nodes.Add "kFullPath", tvwChild, "FL CA" & CStr(sPath(0) & "\"), sPath(0), 1
       sLastNode = sPath(0)
    End If
    
    sKey = "FL CA" & sPath(0) & "\"
    
    For i = 1 To UBound(sPath)
       'If TreeFiles.Nodes(sKey).Children = 0 Then
          TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
        'End If
      If i = UBound(sPath) Then
        sKey = sKey & sPath(i)
      Else
        sKey = sKey & sPath(i) & "\"
      End If
    Next i
    rsFiles.MoveNext
    sKey = ""
 Loop
 
 
 '// CD MEDIA
 
 cmd.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=TRUE"
 Set rsFiles = cmd.Execute
 sLastNode = ""
 sArtist = ""
 sAlbum = ""

 '// add albums folders
 Do Until rsFiles.EOF
    s = rsFiles!FilePath
        
    sPath = Split(s, "\", , vbTextCompare)
    
    If sLastNode <> sPath(0) Then
       TreeFiles.Nodes.Add "kCDMedia", tvwChild, "CDMCA" & CStr(sPath(0) & "\"), sPath(0), 1
       sLastNode = sPath(0)
    End If
    
    sKey = "CDMCA" & sPath(0) & "\"
    
    For i = 1 To UBound(sPath)
       'If TreeFiles.Nodes(sKey).Children = 0 Then
          TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
        'End If
      If i = UBound(sPath) Then
        sKey = sKey & sPath(i)
      Else
        sKey = sKey & sPath(i) & "\"
      End If
    Next i
    rsFiles.MoveNext
    sKey = ""
 Loop
 

TreeFiles.Nodes("kMediaLibrary").Expanded = True
TreeFiles.Nodes("kPlayList").Expanded = True

rsFiles.Close
rsArtist.Close
rsAlbum.Close
Set rsFiles = Nothing
Set rsArtist = Nothing
Set rsAlbum = Nothing

Exit Sub
hell:
MsgBox Err.Description

End Sub

Public Sub UpdateStatusBar(lKilobytes As Long, lSeconds As Long)
    StatusBar1.Panels(2).Text = "RECORDS:[ " & ListView1.ListItems.Count & " ]   -   "
    StatusBar1.Panels(1).Text = "RECORD:[" + str(ListView1.SelectedItem.Index) + "] " + ListView1.SelectedItem

    Dim DD As Long, HH As Long, MM As Long, ss As Long, sTempTime As String
    
    sTempTime = "SIZE: [ " & Format(lKilobytes, "000,000") & " KB. ]   -   "
    
    StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text + sTempTime
    sTempTime = ""
    DD = lSeconds \ 86400     ' Days
    lSeconds = Abs(lSeconds - (DD * 86400))
    HH = lSeconds \ 3600      ' Hours
    MM = lSeconds \ 60 Mod 60 ' Minutes
    ss = lSeconds Mod 60      ' Seconds
    sTempTime = "TIME:[ "
    If DD > 0 Then sTempTime = sTempTime & DD & " days. "
    If HH > 0 Then sTempTime = sTempTime & HH & " Hr. "
    StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text + sTempTime & MM & " Min. " & Format$(ss, "00") & " Sec. ]"

End Sub


Public Sub UpdatePlaycount(sFile As String, Optional bcheckExistinLIBRARY As Boolean)
Dim rsAct As New ADODB.Recordset
Dim iContar As Integer
On Error GoTo hell
Dim s As String

s = Replace(sFile, "'", "''", , , vbTextCompare)
rsAct.Open "SELECT PLAYCOUNT,PLAYEDLAST FROM MUSIC WHERE FILE='" & s & "'", cnnMusic, adOpenDynamic, adLockPessimistic
iContar = rsAct!PlayCount + 1
cmd.CommandText = "UPDATE MUSIC SET PLAYCOUNT=" & iContar & ",PLAYEDLAST='" & Now() & "' WHERE FILE='" & s & "'"
cmd.Execute
rsAct.Close
Set rsAct = Nothing
Exit Sub
hell:
MsgBox Err.Description
End Sub

Private Function ValidateMenu() As Boolean
    On Error Resume Next
    
    Dim bVal1 As Boolean
    Dim tPic As StdPicture
    Dim bVal2 As Boolean
    Dim bBW As Boolean
    
    If ListView1.ListItems.Count > 0 Then
        bVal1 = imgArt.Visible
        Set tPic = Clipboard.GetData(CF_BITMAP)
        bVal2 = (Not tPic Is Nothing And tPic.Handle <> 0)
    End If
    
    ' Apparently, order seems to matter when it comes to hiding certain items
    bBW = (Not bVal1 And bVal2)
    If bBW Then GoTo PasteItem
CopyItem:
    If mnuArtItem(MNU_COPY).Visible <> (Not bVal2 Or bVal1) Then mnuArtItem(MNU_COPY).Visible = (Not bVal2 Or bVal1)
    If bBW Then GoTo ContinueProc
PasteItem:
    If mnuArtItem(MNU_PASTE).Visible <> (Not bVal1 Or bVal2) Then mnuArtItem(MNU_PASTE).Visible = (Not bVal1 Or bVal2)
    If bBW Then GoTo CopyItem
    
ContinueProc:
    If bVal1 And bVal2 Then
        If mnuArtItem(MNU_PASTE).Caption <> PASTE_TXT_2 Then mnuArtItem(MNU_PASTE).Caption = PASTE_TXT_2
    Else
        If mnuArtItem(MNU_PASTE).Caption <> PASTE_TXT_1 Then mnuArtItem(MNU_PASTE).Caption = PASTE_TXT_1
    End If
    
    ValidateMenu = bVal1 Or bVal2
End Function

Public Sub AddTrack(sFile As String)
 On Error Resume Next
  Dim rst As New ADODB.Recordset
  Dim s As String
  Dim lSeconds As Long
  Dim sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String

  Dim ID3 As New clsID3
  Dim sPath As String
  Dim d As String
  ID3.FileName = sFile
  sTitle = ID3.Title
  If sTitle = "" Then sTitle = GetFileTitle(sFile): sTitle = Left(sTitle, Len(sTitle) - 4)
                     rst.Open "SELECT * FROM Music", cnnMusic, adOpenDynamic, adLockOptimistic
                     rst.AddNew
                     rst!File = sFile
                     rst!Title = sTitle
                     rst!Artist = ID3.Artist
                     If ID3.Artist = "" Then rst!Artist = "Unknown"
                     rst!Album = ID3.Album
                     If rst!Album = "" Then rst!Album = "Unknown"
                     rst!Year = ID3.Year
                     rst!Genre = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
                     rst!Comments = ID3.Comments
                     rst!length = FormatTime(ID3.length)
                     rst!BYTES = FileLen(sFile)
                     rst!Seconds = ID3.length
                     rst!PlayCount = 0
                     rst!FilePath = Left(sFile, InStrRev(sFile, "\") - 1)
                     rst!Drive = Left(sFile, 3)
                     rst.Update
                     rst.Close
                  
                     Dim Item
                     Set Item = ListView1.ListItems.Add(, , sTitle)
                     Item.SubItems(1) = ID3.Artist
                     If ID3.Artist = "" Then Item.SubItems(1) = "Unknown"
                     Item.SubItems(2) = ID3.Album
                     If ID3.Album = "" Then Item.SubItems(2) = "Unknown"
                     Item.SubItems(3) = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
                     Item.SubItems(4) = ID3.Year
                     Item.SubItems(5) = FormatTime(ID3.length)
                     Item.SubItems(6) = ""
                     Item.SubItems(7) = ""
                     Item.SubItems(8) = sFile
End Sub

Public Function GetFileTitle(ByVal sFileName As String) As String
    'Returns FileTitle Without Path
    Dim lPos As Long
    lPos = InStrRev(sFileName, "\")

    If lPos > 0 Then
        If lPos < Len(sFileName) Then
            GetFileTitle = Mid$(sFileName, lPos + 1)
        Else
            GetFileTitle = ""
        End If
    Else
        GetFileTitle = Left(sFileName, Len(sFileName) - 4)
    End If
    
End Function
