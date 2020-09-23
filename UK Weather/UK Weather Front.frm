VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "UK Weather Forcast"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   115
      Top             =   6945
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Weather Information"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   114
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weather Forcast"
      Height          =   6255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   9135
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   8520
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   22
         Top             =   4680
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   21
         Top             =   4560
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   8640
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   7320
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
         Begin SHDocVwCtl.WebBrowser WeatherBrowse 
            Height          =   1020
            Index           =   4
            Left            =   1560
            TabIndex        =   18
            Top             =   1320
            Width           =   885
            ExtentX         =   1561
            ExtentY         =   1799
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   5520
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
         Begin SHDocVwCtl.WebBrowser WeatherBrowse 
            Height          =   1140
            Index           =   3
            Left            =   1560
            TabIndex        =   16
            Top             =   1320
            Width           =   1245
            ExtentX         =   2196
            ExtentY         =   2011
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3720
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
         Begin SHDocVwCtl.WebBrowser WeatherBrowse 
            Height          =   1020
            Index           =   2
            Left            =   1560
            TabIndex        =   14
            Top             =   1320
            Width           =   885
            ExtentX         =   1561
            ExtentY         =   1799
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   1920
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
         Begin SHDocVwCtl.WebBrowser WeatherBrowse 
            Height          =   1020
            Index           =   1
            Left            =   1560
            TabIndex        =   12
            Top             =   1320
            Width           =   1005
            ExtentX         =   1773
            ExtentY         =   1799
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.PictureBox WeatherPic 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
         Begin SHDocVwCtl.WebBrowser WeatherBrowse 
            Height          =   1740
            Index           =   0
            Left            =   1560
            TabIndex        =   10
            Top             =   1320
            Width           =   1845
            ExtentX         =   3254
            ExtentY         =   3069
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   5
         Left            =   240
         ScaleHeight     =   210
         ScaleWidth      =   8535
         TabIndex        =   8
         Top             =   570
         Width           =   8535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Five Day Forecast "
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   8415
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   8535
         TabIndex        =   6
         Top             =   4920
         Width           =   8535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Current Nearest Observations "
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   4680
         Width           =   8415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   110
         Top             =   720
         Width           =   8655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   109
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   108
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   107
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   106
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   105
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   104
         Top             =   5880
         Width           =   8895
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   103
         Top             =   5640
         Width           =   8895
      End
      Begin VB.Label Label15 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7080
         TabIndex        =   102
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label14 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   101
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label13 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   100
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Label12 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   99
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label Label11 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   98
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Visibility:"
         Height          =   255
         Index           =   39
         Left            =   6480
         TabIndex        =   97
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Pressure (mB):"
         Height          =   255
         Index           =   38
         Left            =   3480
         TabIndex        =   96
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Wind Speed (mph):"
         Height          =   255
         Index           =   37
         Left            =   3480
         TabIndex        =   95
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Relative Humidity:"
         Height          =   255
         Index           =   36
         Left            =   480
         TabIndex        =   94
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Temperature:"
         Height          =   255
         Index           =   35
         Left            =   480
         TabIndex        =   93
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label10 
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   92
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label10 
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   91
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label10 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   90
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label10 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   89
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label10 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   88
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label9 
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   87
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label9 
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   86
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label9 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   85
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label9 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   84
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label9 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   83
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label8 
         Height          =   255
         Index           =   4
         Left            =   8280
         TabIndex        =   82
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label8 
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   81
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label8 
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   80
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label8 
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   79
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label8 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   78
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label7 
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   77
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label7 
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   76
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label7 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   75
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label7 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   74
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label7 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   73
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label6 
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   72
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label6 
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   71
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label6 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   70
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label6 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   69
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label6 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   68
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Set:"
         Height          =   255
         Index           =   34
         Left            =   7320
         TabIndex        =   67
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Set:"
         Height          =   255
         Index           =   33
         Left            =   5520
         TabIndex        =   66
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Set:"
         Height          =   255
         Index           =   32
         Left            =   3720
         TabIndex        =   65
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Set:"
         Height          =   255
         Index           =   31
         Left            =   1920
         TabIndex        =   64
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Set:"
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   63
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Rise:"
         Height          =   255
         Index           =   29
         Left            =   7320
         TabIndex        =   62
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Rise:"
         Height          =   255
         Index           =   28
         Left            =   5520
         TabIndex        =   61
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Rise:"
         Height          =   255
         Index           =   27
         Left            =   3720
         TabIndex        =   60
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Rise:"
         Height          =   255
         Index           =   26
         Left            =   1920
         TabIndex        =   59
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Rise:"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   58
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Wind Speed:"
         Height          =   255
         Index           =   24
         Left            =   7320
         TabIndex        =   57
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Wind Speed:"
         Height          =   255
         Index           =   23
         Left            =   5520
         TabIndex        =   56
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Wind Speed:"
         Height          =   255
         Index           =   22
         Left            =   3720
         TabIndex        =   55
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Wind Speed:"
         Height          =   255
         Index           =   21
         Left            =   1920
         TabIndex        =   54
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Wind Speed:"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Index:"
         Height          =   255
         Index           =   19
         Left            =   7320
         TabIndex        =   52
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Index:"
         Height          =   255
         Index           =   18
         Left            =   5520
         TabIndex        =   51
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Index:"
         Height          =   255
         Index           =   17
         Left            =   3720
         TabIndex        =   50
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Index:"
         Height          =   255
         Index           =   16
         Left            =   1920
         TabIndex        =   49
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sun Index:"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   48
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Air Polution:"
         Height          =   255
         Index           =   14
         Left            =   7320
         TabIndex        =   47
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Air Polution:"
         Height          =   255
         Index           =   13
         Left            =   5520
         TabIndex        =   46
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Air Polution:"
         Height          =   255
         Index           =   12
         Left            =   3720
         TabIndex        =   45
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Air Polution:"
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   44
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Air Polution:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   43
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   42
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   41
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   40
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   39
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   38
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   37
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Min Temp:"
         Height          =   255
         Index           =   9
         Left            =   7320
         TabIndex        =   36
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Min Temp:"
         Height          =   255
         Index           =   8
         Left            =   5520
         TabIndex        =   35
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Min Temp:"
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   34
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Max Temp:"
         Height          =   255
         Index           =   6
         Left            =   7320
         TabIndex        =   33
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Max Temp:"
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   32
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Max Temp:"
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   31
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   30
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   29
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Min Temp:"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   28
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Max Temp:"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   27
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   26
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Min Temp:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label4 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   24
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Max Temp:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Aelect Local Area"
      Height          =   6255
      Left            =   9360
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Frame Frame3 
         Caption         =   "Current Local Area Selection"
         Height          =   735
         Left            =   120
         TabIndex        =   111
         Top             =   5400
         Width           =   4335
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   360
            Width           =   3855
         End
      End
      Begin MSComctlLib.TreeView WeatherView 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8281
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label18 
         Caption         =   "Select your area for upto 5 days local weather information."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
   End
   Begin RichTextLib.RichTextBox WeatherText 
      Height          =   375
      Left            =   4080
      TabIndex        =   113
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"UK Weather Front.frx":0000
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   8640
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub populate_tree()
' This Code Poulates the Tree view
Dim States(100)
WeatherView.Nodes.Add , , "main", "Europe" 'Create Main Parent
WeatherView.Nodes.Add "main", tvwChild, "UK", "UK" 'Create Main Parent
States(0) = State(0)
cnt = 0
LstCnt = LstStationCnt
If Country(a) = "UK" Then
For a = 0 To LstCnt
    NewOne = True
    For B = 0 To cnt
    If State(a) = States(B) Then NewOne = False
    Next B
    If NewOne = True Then cnt = cnt + 1: States(cnt) = State(a)
Next a
    For a = 0 To cnt
    WeatherView.Nodes.Add "UK", tvwChild, States(a), States(a)
    Next a
    For a = 0 To LstStationCnt
    Id = Location(a) + Code(a)
    WeatherView.Nodes.Add State(a), tvwChild, Id, Area(a)
    Next a
End If
WeatherView.Nodes.Item(1).Expanded = True
End Sub
Private Sub Command1_Click()
' Get Weather Information and display
If CheckConnection() = False Then GoTo MissInternetCheck ' Check for Internet connection
Form1.MousePointer = 11
StatusBar1.Panels(1).Text = "Retrieving Weather information from server. Please Wait..."
DownloadFile Label19.Caption, App.Path + "\weather.htm"
Form1.WeatherText.LoadFile App.Path + "\weather.htm", 1
Call Get_Days
For a = 0 To 4
Label2(a).Caption = Days(a): Label2(a).ForeColor = &HFF&
Label4(a).Caption = inf2(a): Label4(a).ForeColor = &HFF&
Label5(a).Caption = inf3(a): Label5(a).ForeColor = &HFF&
Label6(a).Caption = inf4(a): Label6(a).ForeColor = &HFF&
Label7(a).Caption = inf5(a): Label7(a).ForeColor = &HFF&
Label8(a).Caption = inf6(a): Label8(a).ForeColor = &HFF&
Label9(a).Caption = inf7(a): Label9(a).ForeColor = &HFF&
Label10(a).Caption = inf8(a): Label10(a).ForeColor = &HFF&
WeatherBrowse(a).Navigate (App.Path + "\weather\" + inf1(a) + ".gif"): WeatherBrowse(a).ToolTipText = inf1(a)
Next a
Label11.Caption = inf9
Label12.Caption = inf10
Label13.Caption = inf11
Label14.Caption = inf12
Label15.Caption = inf13
For a = 1 To Len(inf14)
If Mid$(inf14, a, 4) = "<br>" Then Label16(0).Caption = Mid$(inf14, 1, a - 1): Label16(1).Caption = Mid$(inf14, a + 4, (Len(inf14) - (a + 4)))
Next a
StatusBar1.Panels(1).Text = ""
Label1.Caption = Label17.Caption
Form1.MousePointer = 0
MissInternetCheck:
Command1.Enabled = False
End Sub
Private Sub Form_Load()
' Get Weather Locations from wlf file
cnt = 0
Open App.Path + "\locations.wlf" For Input As 1
LoopAgain:
If EOF(1) = True Then GoTo finloop
Input #1, Country(cnt), PartCount(cnt), State(cnt), Area(cnt), Code(cnt), Location(cnt)
cnt = cnt + 1
GoTo LoopAgain
finloop:
LstStationCnt = cnt - 1
Close 1
WeatherPic.ScaleMode = 3: WeatherPic.ScaleWidth = 105: WeatherPic.ScaleHeight = 89
Picture1.ScaleMode = 3: Picture1.ScaleWidth = 105: Picture1.ScaleHeight = 89
Picture3.ScaleMode = 3: Picture3.ScaleWidth = 105: Picture3.ScaleHeight = 89
Picture4.ScaleMode = 3: Picture4.ScaleWidth = 105: Picture4.ScaleHeight = 89
Picture5.ScaleMode = 3: Picture5.ScaleWidth = 105: Picture5.ScaleHeight = 89
For a = 0 To 4
WeatherBrowse(a).Left = 0
WeatherBrowse(a).Top = 0
WeatherBrowse(a).Width = 155
WeatherBrowse(a).Height = 140
WeatherBrowse(a).Navigate (App.Path + "\weather\blank.gif"): WeatherBrowse(a).ToolTipText = "Select An Area then click the Get Information Button"
Next a
Call populate_tree
End Sub
Private Sub WeatherView_DblClick()
Label17.Caption = WeatherView.SelectedItem.Text
Label19.Caption = WeatherView.SelectedItem.Key
Command1.Enabled = True
End Sub
Public Sub Get_Days()
'scan through the html file that was downloaded and recover the weather information
MousePointer = 4
lines = 0
startpos = 10500
stoppos = 0
For a = 9000 To (Len(Form1.WeatherText.Text) - 14000)
If Mid$(Form1.WeatherText.Text, a, 3) = "<td" Then startpos = a + 3
If Mid$(Form1.WeatherText.Text, a, 5) = "</td>" Then stoppos = a - 1
If stoppos > startpos Then textline(lines) = Mid$(Form1.WeatherText.Text, startpos, stoppos - startpos): lines = lines + 1: stoppos = 0
Next a
lastline = lines - 1
'find day 1
linecout = 0
foundline = 0
findfirst:
For a = 1 To Len(textline(linecout))
If foundline = 0 Then
    If Mid$(textline(linecout), a, 6) = "Monday" Or Mid$(textline(linecout), a, 7) = "Tuesday" Or Mid$(textline(linecout), a, 9) = "Wednesday" Or Mid$(textline(linecout), a, 8) = "Thursday" Or Mid$(textline(linecout), a, 6) = "Friday" Or Mid$(textline(linecout), a, 8) = "Saturday" Or Mid$(textline(linecout), a, 6) = "Sunday" Then foundline = linecout
End If
Next a
If foundline = 0 Then linecout = linecout + 1: GoTo findfirst
'get days
daynum = 0
For a = foundline To foundline + 4
For lenth = 1 To Len(textline(a))
If Mid$(textline(a), lenth, 6) = "Monday" Then Days(daynum) = "Monday": daynum = daynum + 1
If Mid$(textline(a), lenth, 7) = "Tuesday" Then Days(daynum) = "Tuesday": daynum = daynum + 1
If Mid$(textline(a), lenth, 9) = "Wednesday" Then Days(daynum) = "Wednesday": daynum = daynum + 1
If Mid$(textline(a), lenth, 8) = "Thursday" Then Days(daynum) = "Thursday": daynum = daynum + 1
If Mid$(textline(a), lenth, 6) = "Friday" Then Days(daynum) = "Friday": daynum = daynum + 1
If Mid$(textline(a), lenth, 8) = "Saturday" Then Days(daynum) = "Saturday": daynum = daynum + 1
If Mid$(textline(a), lenth, 6) = "Sunday" Then Days(daynum) = "Sunday": daynum = daynum + 1
Next lenth
Next a
'get inf1 (weather)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 6 To foundline + 10
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 4) = "alt=" Then startpos = lenth + 5
If Mid$(textline(a), lenth, 6) = "border" Then stoppos = lenth - 2
If stoppos > startpos Then inf1(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf2 (Max Temp)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 12 To foundline + 16
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 4) = "Max:" Then startpos = lenth + 5
If Mid$(textline(a), lenth, 4) = "<br>" Then stoppos = lenth
If stoppos > startpos Then inf2(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf3 (Min Temp)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 12 To foundline + 16
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 4) = "Min:" Then startpos = lenth + 5
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf3(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf4 (Air Polution)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 18 To foundline + 22
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf4(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf5 (Sun Index)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 24 To foundline + 28
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf5(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf6 (Wind Speed)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 30 To foundline + 34
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf6(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf7 (Sun Rise)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 36 To foundline + 40
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf7(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf8 (Sun Set)
daynum = 0
startpos = 0
stoppos = 0
For a = foundline + 42 To foundline + 46
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf8(daynum) = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
Next a
'get inf9 (Current Temperature)
daynum = 0
startpos = 0
stoppos = 0
a = foundline + 51
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf9 = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
'get inf10 (Humidity)
daynum = 0
startpos = 0
stoppos = 0
a = foundline + 53
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf10 = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
'get inf11 (Wind Speed)
daynum = 0
startpos = 0
stoppos = 0
a = foundline + 55
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf11 = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
'get inf12 (Pressure)
daynum = 0
startpos = 0
stoppos = 0
a = foundline + 57
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf12 = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
'get inf13 (Visibility)
daynum = 0
startpos = 0
stoppos = 0
a = foundline + 59
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf13 = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
'get inf14 (Station Data)
daynum = 0
startpos = 0
stoppos = 0
a = foundline + 60
For lenth = 10 To Len(textline(a))
If Mid$(textline(a), lenth, 5) = "serif" Then startpos = lenth + 7
If Mid$(textline(a), lenth, 4) = "</fo" Then stoppos = lenth
If stoppos > startpos Then inf14 = Mid$(textline(a), startpos, stoppos - startpos): stoppos = 0: daynum = daynum + 1
Next lenth
MousePointer = 0
End Sub

