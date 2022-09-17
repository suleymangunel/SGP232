VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SGP232"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm COM 
      Left            =   6960
      Top             =   7140
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   0   'False
      BaudRate        =   57600
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Toolkit"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ayarlar"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Devre Þemasý"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Hakkýnda"
      TabPicture(3)   =   "Form1.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   5955
         Left            =   -74890
         Picture         =   "Form1.frx":037A
         ScaleHeight     =   5895
         ScaleWidth      =   9870
         TabIndex        =   69
         Top             =   540
         Width           =   9930
      End
      Begin VB.Frame Frame9 
         Height          =   6075
         Left            =   120
         TabIndex        =   63
         Top             =   420
         Width           =   9915
         Begin VB.TextBox Text1 
            Height          =   1275
            Left            =   2640
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   70
            Text            =   "Form1.frx":BDD34
            Top             =   4560
            Width           =   4755
         End
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            Height          =   540
            Left            =   9240
            Picture         =   "Form1.frx":BDE4F
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   64
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "net: http://www.geocities.com/interrupt21"
            Height          =   195
            Left            =   3240
            TabIndex        =   77
            Top             =   4200
            Width           =   3615
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "e-mail: interrupt21@yahoo.com"
            Height          =   195
            Left            =   3660
            TabIndex        =   76
            Top             =   3900
            Width           =   2640
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Shareware - 2001"
            Height          =   195
            Left            =   4260
            TabIndex        =   68
            Top             =   3420
            Width           =   1515
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Süleyman GÜNEL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4080
            TabIndex        =   67
            Top             =   3000
            Width           =   1860
         End
         Begin VB.Label Label11 
            Caption         =   "Süleyman GÜNEL's EEPROM Programmer                 (RS232 Powered)"
            Height          =   435
            Left            =   3300
            TabIndex        =   66
            Top             =   2400
            Width           =   3540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "SGP232"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   72
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            Left            =   2640
            TabIndex        =   65
            Top             =   660
            Width           =   4875
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6075
         Left            =   -74880
         TabIndex        =   7
         Top             =   420
         Width           =   9915
         Begin VB.Frame Frame2 
            Caption         =   "I/O Türü (Ýletiþim Arabirimi)"
            Height          =   735
            Left            =   3480
            TabIndex        =   81
            Top             =   180
            Width           =   6315
            Begin VB.OptionButton optWin 
               Caption         =   "Seri Port'a Windows API'leri ile eriþim"
               Height          =   375
               Left            =   2700
               TabIndex        =   83
               Top             =   240
               Value           =   -1  'True
               Width           =   3495
            End
            Begin VB.OptionButton optDirect 
               Caption         =   "Seri Port'a direkt eriþim"
               Height          =   375
               Left            =   180
               TabIndex        =   82
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame Menu2 
            Caption         =   "Servis"
            Height          =   2415
            Left            =   120
            TabIndex        =   29
            Top             =   3540
            Width           =   3255
            Begin VB.CheckBox Check7 
               Caption         =   "Chip Takýlý"
               Height          =   975
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   300
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.Frame Frame8 
               Height          =   1095
               Left            =   1680
               TabIndex        =   33
               Top             =   180
               Width           =   1455
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Power"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   36
                  Top             =   180
                  Width           =   540
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "SDA"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   35
                  Top             =   780
                  Width           =   390
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "SCL"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   34
                  Top             =   480
                  Width           =   360
               End
               Begin VB.Image SclLed 
                  Height          =   225
                  Left            =   1080
                  Picture         =   "Form1.frx":BE159
                  Top             =   480
                  Width           =   225
               End
               Begin VB.Image SdaLed 
                  Height          =   225
                  Left            =   1080
                  Picture         =   "Form1.frx":BE253
                  Top             =   780
                  Width           =   225
               End
               Begin VB.Image PwrLed 
                  Height          =   225
                  Left            =   1080
                  Picture         =   "Form1.frx":BE34D
                  Top             =   180
                  Width           =   225
               End
            End
            Begin VB.CheckBox Check3 
               Caption         =   "SCL"
               Height          =   915
               Left            =   1140
               Picture         =   "Form1.frx":BE447
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   1380
               Width           =   975
            End
            Begin VB.CheckBox Check2 
               Caption         =   "SDA"
               Height          =   915
               Left            =   2160
               Picture         =   "Form1.frx":BE889
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   1380
               Width           =   975
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Power"
               Height          =   915
               Left            =   120
               Picture         =   "Form1.frx":BECCB
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   1380
               Width           =   975
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Dosya Ýþlemleri"
            Height          =   4995
            Left            =   3480
            TabIndex        =   22
            Top             =   960
            Width           =   6315
            Begin VB.Frame Frame10 
               Height          =   435
               Left            =   3420
               TabIndex        =   51
               Top             =   180
               Width           =   2715
               Begin VB.CheckBox Check6 
                  Caption         =   "S"
                  Height          =   195
                  Left            =   1860
                  TabIndex        =   54
                  Top             =   180
                  Width           =   435
               End
               Begin VB.CheckBox Check5 
                  Caption         =   "H"
                  Height          =   195
                  Left            =   1140
                  TabIndex        =   53
                  Top             =   180
                  Width           =   435
               End
               Begin VB.CheckBox Check4 
                  Caption         =   "R"
                  Height          =   195
                  Left            =   420
                  TabIndex        =   52
                  Top             =   180
                  Value           =   1  'Checked
                  Width           =   495
               End
            End
            Begin VB.TextBox FileAt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5160
               TabIndex        =   50
               Top             =   3300
               Width           =   1035
            End
            Begin VB.TextBox FileLn 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3420
               TabIndex        =   49
               Top             =   3300
               Width           =   1695
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Sil"
               Height          =   1095
               Left            =   4320
               Picture         =   "Form1.frx":BF10D
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   3780
               Width           =   1935
            End
            Begin VB.TextBox FileNm 
               Height          =   315
               Left            =   180
               TabIndex        =   28
               Top             =   3300
               Width           =   3195
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Buffer'a Yükle"
               Height          =   1095
               Left            =   2220
               Picture         =   "Form1.frx":BF9D7
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   3780
               Width           =   1995
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Buffer Kaydet"
               Height          =   1095
               Left            =   180
               Picture         =   "Form1.frx":C0819
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   3780
               Width           =   1935
            End
            Begin VB.FileListBox File1 
               Height          =   2430
               Left            =   3420
               System          =   -1  'True
               TabIndex        =   25
               Top             =   720
               Width           =   2775
            End
            Begin VB.DirListBox Dir1 
               Height          =   2340
               Left            =   120
               TabIndex        =   24
               Top             =   720
               Width           =   3255
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   120
               TabIndex        =   23
               Top             =   300
               Width           =   3255
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Genel ayarlar"
            Height          =   3315
            Left            =   120
            TabIndex        =   8
            Top             =   180
            Width           =   3255
            Begin VB.VScrollBar ETV 
               Height          =   375
               Left            =   1380
               TabIndex        =   79
               Top             =   2850
               Width           =   195
            End
            Begin VB.TextBox ET 
               Height          =   315
               Left            =   660
               TabIndex        =   78
               Top             =   2880
               Width           =   615
            End
            Begin VB.VScrollBar TTV 
               Height          =   375
               LargeChange     =   5
               Left            =   1380
               TabIndex        =   61
               Top             =   2400
               Width           =   195
            End
            Begin VB.VScrollBar PTV 
               Height          =   375
               LargeChange     =   5
               Left            =   1380
               TabIndex        =   60
               Top             =   1980
               Width           =   195
            End
            Begin VB.VScrollBar DTV 
               Height          =   375
               LargeChange     =   5
               Left            =   1380
               TabIndex        =   59
               Top             =   1560
               Width           =   195
            End
            Begin VB.VScrollBar WTV 
               Height          =   375
               LargeChange     =   5
               Left            =   1380
               TabIndex        =   58
               Top             =   1140
               Width           =   195
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   660
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   360
               Width           =   930
            End
            Begin VB.VScrollBar RTV 
               Height          =   375
               LargeChange     =   5
               Left            =   1380
               TabIndex        =   56
               Top             =   740
               Width           =   195
            End
            Begin VB.TextBox TT 
               Height          =   315
               Left            =   660
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   2460
               Width           =   615
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Kaydet"
               Height          =   915
               Left            =   1740
               Picture         =   "Form1.frx":C165B
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   2280
               Width           =   1395
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Geçerli"
               Height          =   915
               Left            =   1740
               Picture         =   "Form1.frx":C1A9D
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   1280
               Width           =   1395
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Optimum"
               Height          =   915
               Left            =   1740
               Picture         =   "Form1.frx":C1EDF
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   280
               Width           =   1395
            End
            Begin VB.TextBox PT 
               Height          =   315
               Left            =   660
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   2040
               Width           =   615
            End
            Begin VB.TextBox DT 
               Height          =   315
               Left            =   660
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   1620
               Width           =   615
            End
            Begin VB.TextBox WT 
               Height          =   315
               Left            =   660
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox RT 
               Height          =   315
               Left            =   660
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   770
               Width           =   615
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "ET"
               Height          =   195
               Left            =   180
               TabIndex        =   80
               Top             =   2940
               Width           =   255
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "TT"
               Height          =   195
               Left            =   180
               TabIndex        =   18
               Top             =   2520
               Width           =   255
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "PT"
               Height          =   195
               Left            =   180
               TabIndex        =   17
               Top             =   2100
               Width           =   255
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "DT"
               Height          =   195
               Left            =   180
               TabIndex        =   16
               Top             =   1680
               Width           =   270
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "WT"
               Height          =   195
               Left            =   180
               TabIndex        =   15
               Top             =   1275
               Width           =   315
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "RT"
               Height          =   195
               Left            =   180
               TabIndex        =   14
               Top             =   840
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "COM"
               Height          =   195
               Left            =   180
               TabIndex        =   13
               Top             =   420
               Width           =   420
            End
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6075
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   9915
         Begin ComctlLib.ProgressBar StatBar 
            Height          =   315
            Left            =   4740
            TabIndex        =   84
            Top             =   5640
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Frame Menu1 
            Caption         =   "MENU"
            Height          =   5355
            Left            =   8460
            TabIndex        =   42
            Top             =   180
            Width           =   1335
            Begin VB.TextBox ExChangeDec 
               Height          =   315
               Left            =   720
               MaxLength       =   3
               TabIndex        =   75
               Top             =   3780
               Width           =   485
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Deðiþtir"
               Height          =   1035
               Left            =   120
               Picture         =   "Form1.frx":C2321
               Style           =   1  'Graphical
               TabIndex        =   74
               Top             =   4200
               Width           =   1095
            End
            Begin VB.TextBox ExChangeHex 
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   73
               Top             =   3780
               Width           =   375
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               ItemData        =   "Form1.frx":C262B
               Left            =   120
               List            =   "Form1.frx":C262D
               Style           =   2  'Dropdown List
               TabIndex        =   62
               Top             =   300
               Width           =   1095
            End
            Begin VB.CommandButton Command14 
               Caption         =   "Doldur"
               Height          =   1035
               Left            =   120
               Picture         =   "Form1.frx":C262F
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Text>Buffer"
               Height          =   1035
               Left            =   120
               Picture         =   "Form1.frx":C2939
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   1860
               Width           =   1095
            End
         End
         Begin TabDlg.SSTab SSTAB 
            Height          =   5295
            Left            =   1500
            TabIndex        =   39
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   9340
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Buffer 0"
            TabPicture(0)   =   "Form1.frx":C2C43
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "MemWindow(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Buffer 1"
            TabPicture(1)   =   "Form1.frx":C2C5F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "MemWindow(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Buffer 2"
            TabPicture(2)   =   "Form1.frx":C2C7B
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MemWindow(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Text 0"
            TabPicture(3)   =   "Form1.frx":C2C97
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "MemWindowText(0)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Text 1"
            TabPicture(4)   =   "Form1.frx":C2CB3
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "MemWindowText(1)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Text 2"
            TabPicture(5)   =   "Form1.frx":C2CCF
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "MemWindowText(2)"
            Tab(5).ControlCount=   1
            Begin VB.TextBox MemWindowText 
               Height          =   4755
               Index           =   2
               Left            =   -74880
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   48
               Top             =   420
               Width           =   6615
            End
            Begin VB.TextBox MemWindowText 
               Height          =   4755
               Index           =   1
               Left            =   -74880
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   47
               Top             =   420
               Width           =   6615
            End
            Begin VB.TextBox MemWindow 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   4755
               Index           =   2
               Left            =   -74880
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   46
               Top             =   420
               Width           =   6615
            End
            Begin VB.TextBox MemWindow 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   4755
               Index           =   1
               Left            =   -74880
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   45
               Top             =   420
               Width           =   6615
            End
            Begin VB.TextBox MemWindowText 
               Height          =   4755
               Index           =   0
               Left            =   -74880
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   41
               Top             =   420
               Width           =   6615
            End
            Begin VB.TextBox MemWindow 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   4755
               Index           =   0
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   420
               Width           =   6615
            End
         End
         Begin VB.Frame Menu0 
            Caption         =   "MENU"
            Height          =   5355
            Left            =   120
            TabIndex        =   2
            Top             =   180
            Width           =   1275
            Begin VB.CommandButton Command11 
               Caption         =   "Blnk Test"
               Height          =   915
               Left            =   120
               Picture         =   "Form1.frx":C2CEB
               Style           =   1  'Graphical
               TabIndex        =   71
               Top             =   4320
               Width           =   1035
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Test"
               Height          =   915
               Left            =   120
               Picture         =   "Form1.frx":C2FF5
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   3300
               Width           =   1035
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Sil"
               Height          =   915
               Left            =   120
               Picture         =   "Form1.frx":C32FF
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   2280
               Width           =   1035
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Yaz"
               Height          =   915
               Left            =   120
               Picture         =   "Form1.frx":C3609
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   1260
               Width           =   1035
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Oku"
               Height          =   915
               Left            =   120
               Picture         =   "Form1.frx":C3913
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   240
               Width           =   1035
            End
         End
         Begin VB.Label Info 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   5640
            Width           =   4575
         End
      End
   End
   Begin VB.Image LedOff 
      Height          =   225
      Left            =   420
      Picture         =   "Form1.frx":C3C1D
      Top             =   7260
      Width           =   225
   End
   Begin VB.Image LedOn 
      Height          =   225
      Left            =   120
      Picture         =   "Form1.frx":C3D17
      Top             =   7260
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IOType As Boolean
Dim PortAddr, PortAdress(4) As Integer
Dim PortVal As Long
Dim bSize As Byte
Dim PortRet As Boolean
Dim ErrWait, ErrW As Integer
Dim ConfigFile As String
Dim BRK As Boolean
Dim i_byte, Prt As Byte
Dim Memory(3, 2048), WriteBuffer(16) As Byte
Dim ErrCount, SysSpeed, speed, speedR, speedW, speedD, speedP, speedT As Single
Dim PwrStt, SclStt, SdaStt As Boolean
Sub SDA_high()
 If IOType = True Then
  PortRet = GetPortVal(PortAddr + 4, PortVal, 1)
  PortRet = SetPortVal(PortAddr + 4, PortVal Or 1, 1)
 Else
  COM.DTREnable = True
  COM.Break = True
 End If
End Sub
Sub SDA_low()
 If IOType = True Then
  PortRet = GetPortVal(PortAddr + 4, PortVal, 1)
  PortRet = SetPortVal(PortAddr + 4, PortVal And 254, 1)
 Else
  COM.DTREnable = False
  COM.Break = True
 End If
End Sub
Sub SCL_high()
 If IOType = True Then
  PortRet = GetPortVal(PortAddr + 4, PortVal, 1)
  PortRet = SetPortVal(PortAddr + 4, PortVal Or 2, 1)
 Else
  COM.RTSEnable = True
  COM.Break = True
 End If
 bekle
End Sub
Sub SCL_low()
 If IOType = True Then
  PortRet = GetPortVal(PortAddr + 4, PortVal, 1)
  PortRet = SetPortVal(PortAddr + 4, PortVal And 253, 1)
 Else
  COM.RTSEnable = False
  COM.Break = True
 End If
 bekle
End Sub
Function InBit() As Byte
 If IOType = True Then
  PortRet = GetPortVal(PortAddr + 6, PortVal, 1)
  InBit = ((PortVal And 16) / 16) And 1
 Else
  If COM.CTSHolding Then InBit = 1 Else InBit = 0
 End If
End Function
Sub Startp()
 SDA_high
 SCL_high
 SDA_low
 SCL_low
End Sub
Sub Stopp()
 SDA_low
 SCL_high
 SDA_high
 SCL_low
End Sub
Function Ack() As Boolean
 SDA_high
 SCL_high
 Ack = Abs(InBit() - 1)
 SCL_low
End Function
Function NoAck()
 NoAck = Not (Ack)
End Function
Sub OutByte(veri As Byte)
Dim n, c As Byte
For n = 7 To 0 Step -1
 c = veri And (2 ^ n)
 If c = 0 Then
  SDA_low
 Else
  SDA_high
 End If
 SCL_high
 SCL_low
Next
End Sub
Function InByte() As Byte
Dim veri, n, b, c As Byte
veri = 0
For n = 7 To 0 Step -1
 SCL_high
 veri = veri Or ((2 ^ n) * InBit())
 SCL_low
Next
InByte = veri
End Function
Sub bekle()
 If speed <> 0 Then delay (speed)
End Sub
Function delay(ms As Single)
 Dim mj, t As Single
 While (mj < ms * SysSpeed)
  mj = mj + 1
  t = Timer
 Wend
End Function

Function Zaman()
Dim c, t As Single
MsgBox "Lütfen program ekraný görünene kadar herhangi bir iþlem yapmayýn, varsa aktif uygulamalarý kapatýn.", vbOKOnly + vbExclamation, "Uyarý!"
t = Int(Timer)
While (t = Int(Timer)): Wend
t = Int(Timer)
t = Timer + 1
While (Timer < t)
 c = c + 1
Wend
SysSpeed = c / 1000
End Function
Sub PowerOff()
 SDA_low
 SCL_low
 COM.Break = False
 If COM.PortOpen = True Then COM.PortOpen = False
 PwrStt = False
End Sub
Sub PowerOn()
 If COM.PortOpen = False Then COM.PortOpen = True
 COM.Break = True
 delay (3)
End Sub
Function WriteBytePage(Adress As Integer) As Byte
 Dim m As Integer
 BRK = False
 delay (speedD)
 Startp
 OutByte ((&HA0 Or ((Adress And &H700) / 128)))
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then WriteBytePage = 1: Exit Function
 Wend
 OutByte (Adress And 255)
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then WriteBytePage = 1: Exit Function
 Wend
 For m = 0 To 15
  OutByte (WriteBuffer(m))
  ErrCount = 0
  While (Not (Ack))
   ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
   If BRK Then WriteBytePage = 1: Exit Function
  Wend
 Next m
 Stopp
 WriteBytePage = 0
End Function
Function WriteByte(Adress As Integer, Data As Byte) As Byte
 BRK = False
 Startp
 OutByte ((&HA0 Or ((Adress And &H700) / 128)))
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then WriteByte = 1: Exit Function
 Wend
 OutByte (Adress And 255)
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then WriteByte = 1: Exit Function
 Wend
 OutByte (Data And 255)
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then WriteByte = 1: Exit Function
 Wend
 Stopp
 WriteByte = 0
End Function
Function ReadRandomByte(Adress As Integer) As Byte
 BRK = False
 Startp
 OutByte ((&HA0 Or ((Adress And &H700) / 128)))
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then ReadRandomByte = 1: Exit Function
 Wend
 OutByte (Adress And 255)
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then ReadRandomByte = 1: Exit Function
 Wend
 Startp
 OutByte ((&HA1 Or ((Adress And &H700) / 128)))
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then ReadRandomByte = 1: Exit Function
 Wend
 i_byte = InByte
 ErrCount = 0
 While (Not (NoAck))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then ReadRandomByte = 1: Exit Function
 Wend
 Stopp
 ReadRandomByte = i_byte
End Function
Function ReadCurrentByte() As Byte
 BRK = False
 Startp
 OutByte (&HA1)
 ErrCount = 0
 While (Not (Ack))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then ReadCurrentByte = 1: Exit Function
 Wend
 i_byte = InByte
 ErrCount = 0
 While (Not (NoAck))
  ErrCount = ErrCount + 1: If ErrCount >= ErrWait * SysSpeed Then BRK = True
  If BRK Then ReadCurrentByte = 1: Exit Function
 Wend
 Stopp
 ReadCurrentByte = i_byte
End Function

Private Sub Check1_Click()
If Check1.Value = 1 Then
 PowerOn
 PwrLed = LedOn
Else
 If Check7.Value = 1 Then
  Check2.Value = 0
  Check3.Value = 0
 End If
 PwrLed = LedOff
 PowerOff
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
 If Check7.Value = 1 Then Check1.Value = 1
 SDA_high
 SdaLed = LedOn
Else
 SDA_low
 SdaLed = LedOff
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
 If Check7.Value = 1 Then Check1.Value = 1
 SCL_high
 SclLed = LedOn
Else
 SCL_low
 SclLed = LedOff
End If
End Sub

Private Sub Check4_Click()
File1.ReadOnly = Check4.Value
End Sub

Private Sub Check5_Click()
File1.Hidden = Check5.Value
End Sub

Private Sub Check6_Click()
File1.System = Check6.Value
End Sub

Private Sub Command1_Click()
 Dim ndx As Byte
 Dim n As Integer
   
 ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
 speed = speedR
 Info = "EEPROM Okunuyor.."
 If PwrStt = False Then PowerOn: delay (1500)
 ReadRandomByte (&H7FF)
 StatBar.Min = 0: StatBar.Max = 2047
 For n = 0 To 2047
  If BRK Then PowerOff: MsgBox "Hata": Exit Sub
  Memory(ndx, n) = ReadCurrentByte
 StatBar = n
 Next n
 Info = "": StatBar = 0: PowerOff
 MemDisplay (ndx)
End Sub

Private Sub Command10_Click()
Dim konum, dosya, ozellik As String
Dim cevap, att As Integer
cevap = vbYes
If Trim(FileNm) <> "" Then
 konum = Dir1.Path: If Right(konum, 1) <> "\" Then konum = konum + "\"
 dosya = konum + FileNm
 If Dir(dosya, vbHidden + vbReadOnly + vbSystem) <> "" Then
  cevap = MsgBox(dosya + " Dosyasýný silmek istediðinizden eminmisiniz ?", vbYesNo, "Silem Onayý")
  If cevap = vbYes Then
   att = GetAttr(dosya)
   If (att And vbReadOnly) Then ozellik = ozellik + "Salt okunur" + vbCrLf
   If (att And vbHidden) Then ozellik = ozellik + "Gizli" + vbCrLf
   If (att And vbSystem) Then ozellik = ozellik + "Sistem" + vbCrLf
   If ozellik <> "" Then
    cevap = MsgBox("Dosya özellikleri:" + vbCrLf + vbCrLf + ozellik + vbCrLf + "Bu dosyayý silmek istediðinizden eminmisiniz?", vbYesNo + vbExclamation, "Uyarý !")
    If cevap = vbYes Then SetAttr dosya, vbNormal
   End If
  End If
  If cevap = vbYes Then
   Kill (dosya)
   File1.Refresh
   FileNm = ""
   FileLn = ""
   Dir1.Refresh: File1.Refresh
  End If
 End If
End If
End Sub

Private Sub Command11_Click()
 Dim k, ndx As Byte
 Dim n As Integer
   
 speed = speedR
 Info = "EEPROM Test Ediliyor (FFh)..."
 If PwrStt = False Then PowerOn: delay (1500)
 ReadRandomByte (&H7FF)
 If BRK Then PowerOff: MsgBox "Hata": Exit Sub
 StatBar.Min = 0: StatBar.Max = 2047
 For n = 0 To 2047
  If ReadCurrentByte <> &HFF Then
  If BRK Then PowerOff: MsgBox "Hata": Exit Sub
  Info = "": StatBar = 0: PowerOff
  MsgBox "EEPROM Boþ deðil !" + vbCrLf + vbCrLf + "Adres:" + Hex(n) + "h", vbExclamation, "EEPROM Blank Test Sonucu"
  Exit Sub
 End If
 StatBar = n
 Next n
 Info = "": StatBar = 0: PowerOff
 MsgBox "EEPROM Boþ.", vbInformation, "EEPROM Blank Test Sonucu"
End Sub

Private Sub Command12_Click()
Dim St, Hs, c, j1 As String
Dim Ss, Ssc As Integer
Dim Rm, ndx As Byte
Dim a1, a2, a3 As Integer
 
ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
St = MemWindow(ndx)
c = MemWindow(ndx).SelText: Ss = MemWindow(ndx).SelStart
If Len(Trim(c)) <> 2 Then Exit Sub
If Left(c, 1) = " " Then Ss = Ss + 1
c = Trim(c)
If St <> "" Then
 j1 = MemWindowText(ndx)
 Ssc = Val(Trim(ExChangeDec)): ExChangeDec = Ssc
 If Ssc < 0 Or Ssc > 255 Then Ssc = 0: ExChangeDec = 0
 a1 = Int(Ss / 54): a2 = Int((Ss - a1 * 54) / 3): a3 = (a1 * 16 + a2) - 1
 If a2 < 1 Then Exit Sub
 Rm = Val(Trim(ExChangeDec)): Hs = Trim(Hex(Rm))
 If Len(Hs) > 1 Then
  Mid$(St, Ss + 1, 1) = Left(Hs, 1)
 Else
  Mid$(St, Ss + 1, 1) = "0"
 End If
 Mid$(St, Ss + 2, 1) = Right(Hs, 1)
 MemWindow(ndx) = St: MemWindow(ndx).SelStart = Ss
 Memory(ndx, a3) = Rm
 If j1 <> "" And Rm <> 0 Then
  Mid(j1, a3 + 1, 1) = Chr(Rm)
  MemWindowText(ndx) = j1
 End If
End If
End Sub

Private Sub Command13_Click()
Dim ndx As Byte
Dim lmt, n As Integer

lmt = 2047
ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
For n = 0 To 2048: Memory(ndx, n) = 0: Next
If Len(MemWindowText(ndx)) = 0 Then GoTo devam
If Len(MemWindowText(ndx)) > 2048 Then lmt = 2047
If Len(MemWindowText(ndx)) < 2048 Then lmt = Len(MemWindowText(ndx)) - 1
For n = 0 To lmt
 Memory(ndx, n) = Asc(Mid(MemWindowText(ndx), n + 1, 1))
Next
devam:
MemDisplay (ndx)
End Sub

Private Sub Command14_Click()
Dim ndx, n As Integer
Dim fnum, i As Byte

If Combo2.ListIndex = 0 Then fnum = 0
If Combo2.ListIndex = 1 Then fnum = -1
If Combo2.ListIndex = 2 Then fnum = &HFF
ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
StatBar.Min = 0: StatBar.Max = 2047
For n = 0 To 2047
 If fnum = -1 Then
  Memory(ndx, n) = i
 Else
  Memory(ndx, n) = fnum
 End If
 StatBar = n
 If i + 1 = 256 Then i = 0 Else i = i + 1
Next
MemDisplay (ndx)
End Sub

Private Sub Command2_Click()
 Dim n, j, cevap, ndx As Integer

 ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
 If PwrStt = False Then PowerOn: delay (1500)
 speed = speedW
 cevap = MsgBox("EEPROM'daki veriler deðiþecektir !" + vbCrLf + "Ýþleme devam etmek istiyormusunuz ?", vbYesNo + vbExclamation, "EEPROM Yazma Onayý")
 If cevap = vbNo Then PowerOff: Exit Sub
 Info = "EEPROM Yazýlýyor..."
 StatBar.Min = 0: StatBar.Max = 127
 StatBar = 0
 For n = 0 To 127
  For j = 0 To 15
   WriteBuffer(j) = Memory(ndx, n * 16 + j)
  Next j
  If n = 127 Then speed = speedP
  If WriteBytePage(n * 16) Then
   PowerOff
   Info = "EEPROM Yazýlamadý !"
   Exit Sub
  End If
  StatBar = n
 Next n
 Info = "": StatBar = 0: PowerOff
 MsgBox "EEPROM Yazýldý.", vbInformation
End Sub

Private Sub Command3_Click()
 Dim n, cevap, ndx As Integer

 ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
 If PwrStt = False Then PowerOn: delay (1500)
 speed = speedW
 cevap = MsgBox("EEPROM'daki veriler silinecektir !" + vbCrLf + "Ýþleme devam etmek istiyormusunuz ?", vbYesNo + vbExclamation, "EEPROM Silme Onayý")
 If cevap = vbNo Then PowerOff: Exit Sub
 Info = "EEPROM Siliniyor..."
 StatBar.Min = 0: StatBar.Max = 127
 StatBar = 0
 For n = 0 To 2047: Memory(ndx, n) = &HFF: Next
 For n = 0 To 15: WriteBuffer(n) = &HFF: Next
 For n = 0 To 127
  If n = 127 Then speed = speedP
  If WriteBytePage(n * 16) Then
   PowerOff
   Info = "EEPROM Silinemedi !"
   Exit Sub
  End If
  StatBar = n
 Next n
 Info = "": StatBar = 0: PowerOff
 MemDisplay (ndx)
 MsgBox "EEPROM Silindi.", vbInformation
End Sub

Private Sub Command4_Click()
 Dim c, ec, em, cevap, cevap2, ErrFnd, ErrCrr, n As Integer
 Dim ndx, k, cop As Byte
 Dim acr As Boolean
 
 acr = False
 ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
 If PwrStt = False Then PowerOn: delay (1500)
 speed = speedR
 Info = "EEPROM Test Ediliyor..."
 cevap = MsgBox("Hatalar otomatik olarak düzeltinsinmi ?", vbYesNoCancel + vbQuestion, "Hata Düzeltme Onayý")
 If cevap = vbCancel Then PowerOff: Exit Sub
 If cevap = vbYes Then acr = True Else acr = False
 ReadRandomByte (&H7FF)
 If BRK Then PowerOff: MsgBox "Hata": Exit Sub
 StatBar.Min = 0: StatBar.Max = 2047
 For n = 0 To 2047
  If BRK Then PowerOff: MsgBox "Hata": Exit Sub
  k = ReadCurrentByte
  If k <> Memory(ndx, n) Then
   ec = 0
   Do
    If Memory(ndx, n) <> ReadRandomByte((n)) Then ec = ec + 1
    If BRK Then PowerOff: MsgBox "Hata": Exit Sub
   Loop While (ec <> 3 And ec <> 0)
   If ec = 3 Then
    em = 0: ErrFnd = ErrFnd + 1
    If acr = False Then
     cevap2 = MsgBox("Hata düzeltilsinmi ?" + Chr(13) + "EEPROM - > Adres:" + Hex(n) + " , Veri:" + Hex(k) + Chr(13) + "Bellek -> Adres:" + Hex(n) + " , Veri:" + Hex(Memory(ndx, n)), vbYesNoCancel + vbExclamation, "Farklý Veriye Rastlandý")
     If cevap2 = vbCancel Then Info = "": StatBar = 0: PowerOff: Exit Sub
    End If
    If acr = True Or cevap2 = vbYes Then
     Do
      speed = speedT: cop = WriteByte((n), (Memory(ndx, n))): speed = speedR
      If BRK Then PowerOff: MsgBox "Hata": Exit Sub
      ec = 0: em = em + 1
      Do
       If Memory(ndx, n) <> ReadRandomByte((n)) Then ec = ec + 1
       If BRK Then PowerOff: MsgBox "Hata": Exit Sub
      Loop While (ec <> 3 And ec <> 0)
     Loop While (ec <> 0 And em < 3)
     If em = 4 And ec = 3 Then
      MsgBox "EEPROM Hatasý !" + Chr(13) + "Adres:" + Hex(n) + " (Hex) , " + Str(n) + " (Dec)", vbCritical
     End If
     ErrCrr = ErrCrr + 1
    End If
   End If
  End If
  StatBar = n
 Next n
 Info = "": StatBar = 0: PowerOff
 cevap = MsgBox("Hatalý Birim Sayýsý......:" + Str(ErrFnd) + vbCrLf + "Düzeltilen Birim Sayýsý:" + Str(ErrCrr), vbOKOnly + vbInformation, "EEPROM Test Sonucu")
End Sub

Private Sub Command5_Click()
RT = 0: WT = 0: DT = 6.5: PT = 2.5: TT = 6.5: ET = 10: optDirect = True
End Sub

Private Sub Command6_Click()
LoadConfig
RT = speedR: RTV = RT * 10
WT = speedW: WTV = WT * 10
DT = speedD: DTV = DT * 10
PT = speedP: PTV = PT * 10
TT = speedT: TTV = TT * 10
ET = ErrW: ETV = ET: ErrWait = ErrW * SysSpeed
optDirect = IOType: optWin = Not (IOType)
PortAddr = PortAdress(Combo1.ListIndex)
MsgBox "Geçerli ayarlar yüklendi."
End Sub

Private Sub Command7_Click()
Dim xT As Single
Dim FF, i As Integer
FF = FreeFile
If Dir(ConfigFile) <> "" Then
 If FileLen(ConfigFile) <> 898 Then
  SetAttr ConfigFile, vbNormal
  Kill (ConfigFile)
 End If
End If
If Combo1.ListCount <> 0 Then
 Prt = Val(Right(Combo1.Text, 1))
 If COM.PortOpen = True Then COM.PortOpen = False
 COM.CommPort = Prt
End If
speedR = RT: speedW = WT: speedD = DT
speedP = PT: speedT = TT: ErrW = ET: ErrWait = ErrW * SysSpeed
IOType = optDirect
Open ConfigFile For Random As #FF
 Put #FF, 1, Prt
 Put #FF, 2, IOType
 Put #FF, 3, speedR
 Put #FF, 4, speedW
 Put #FF, 5, speedD
 Put #FF, 6, speedP
 Put #FF, 7, speedT
 Put #FF, 8, ErrW
Close #FF
File1.Refresh
PortAddr = PortAdress(Combo1.ListIndex)
MsgBox "Ayarlar kayýt edildi."
End Sub

Private Sub Command8_Click()
On Error GoTo hata
Dim konum, dosya, ozellik, ds As String
Dim FF, i, ndx, cevap, att As Integer
Dim veri As String
FileNm = Trim(FileNm)
If FileNm <> "" Then
 cevap = vbYes
 ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
 FileNm = Trim(FileNm)
 konum = Dir1.Path: If Right(konum, 1) <> "\" Then konum = konum + "\"
 dosya = konum + FileNm
 If Dir(dosya, vbNormal + vbHidden + vbReadOnly + vbSystem) <> "" Then
  cevap = MsgBox("Bu isim altýnda bir dosya var !" + vbCrLf + "Üzerine yazýlsýnmý ?" + vbCrLf + vbCrLf + dosya, vbYesNo, "Uyarý !")
  ozellik = ""
  If cevap = vbYes Then
   att = GetAttr(dosya)
   If (att And vbReadOnly) Then ozellik = ozellik + "Salt okunur" + vbCrLf
   If (att And vbHidden) Then ozellik = ozellik + "Gizli" + vbCrLf
   If (att And vbSystem) Then ozellik = ozellik + "Sistem" + vbCrLf
   If ozellik <> "" Then
    cevap = MsgBox("Dosya özellikleri:" + vbCrLf + vbCrLf + ozellik + vbCrLf + "Bu dosyanýn üzerine yazýlsýnmý ?", vbYesNo + vbExclamation, "Uyarý !")
    If cevap = vbYes Then SetAttr dosya, vbNormal
   End If
  End If
 End If
 If cevap = vbYes Then
  If Dir(dosya, vbNormal) <> "" Then Kill (dosya)
  FF = FreeFile: veri = " "
  Open dosya For Binary As #FF
  For i = 0 To 2047
   veri = Chr(Memory(ndx, i))
   Put #FF, i + 1, veri
  Next i
  Close #FF
  Dir1.Refresh: File1.Refresh
 End If
End If
Exit Sub
hata:
MsgBox "Hata oldu !"
Err.Clear
Exit Sub
End Sub

Private Sub Command9_Click()
Dim dosya, konum As String
Dim boy, ndx, n, FF As Integer
Dim veri As String
ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
FileNm = Trim(FileNm)
konum = Dir1.Path: If Right(konum, 1) <> "\" Then konum = konum + "\"
dosya = konum + FileNm
If FileNm <> "" And Dir(dosya, vbNormal + vbHidden + vbReadOnly + vbSystem) <> "" Then
 If FileLen(dosya) <> 0 Then
  If FileLen(dosya) < 2048 Then boy = FileLen(dosya) Else boy = 2048
  FF = FreeFile: veri = " "
  Open dosya For Binary As #FF
   For n = 0 To boy - 1
    Get #FF, n + 1, veri
    Memory(ndx, n) = Asc(veri)
   Next n
  Close #FF
  MemDisplay (ndx)
  MsgBox "Yüklenecek dosya: " & UCase(FileNm) & vbCrLf & vbCrLf & "Dosya yüklendi."
 End If
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
FileNm = ""
End Sub

Private Sub Drive1_Change()
On Error GoTo hata
 Dir1.Path = Drive1.Drive
 FileNm = ""
 Exit Sub
hata:
If Err.Number = 68 Then Drive1.Drive = "C:\"
Err.Clear
Resume Next
End Sub

Private Sub ETV_Change()
ET = ETV
End Sub

Private Sub File1_Click()
Dim konum, dosya As String
Dim att As Integer
FileAt = ""
konum = Dir1.Path: If Right(konum, 1) <> "\" Then konum = konum + "\"
FileNm = File1.filename
dosya = konum + FileNm
att = GetAttr(dosya)
If (att And vbReadOnly) Then FileAt = FileAt + "+R" Else FileAt = FileAt + "-R"
If (att And vbHidden) Then FileAt = FileAt + "+H" Else FileAt = FileAt + "-H"
If (att And vbSystem) Then FileAt = FileAt + "+S" Else FileAt = FileAt + "-S"
FileLn = FileLen(dosya)
End Sub

Private Sub File1_DblClick()
Command9_Click
End Sub

Private Sub Form_Load()
Dim cc As Integer
speed = 0: speedR = 0: speedW = 0: speedD = 6.5: speedP = 2.5: speedT = 6.5
ConfigFile = "SGP232.CFG"
FindPorts
LoadConfig
If Combo1.ListCount <> 0 Then
 COM.CommPort = Prt
Else
 Menu0.Enabled = False
 Menu1.Enabled = False
 Menu2.Enabled = False
 MsgBox "Kart ile iletiþim kurulacak bir seri port bulunamadý !", vbExclamation, "Kritik Hata !"
End If
RT = speedR: RTV = RT * 10
WT = speedW: WTV = WT * 10
DT = speedD: DTV = DT * 10
PT = speedP: PTV = PT * 10
TT = speedT: TTV = TT * 10
ET = ErrW: ETV = ErrW
Combo2.AddItem "0,0,0,0"
Combo2.AddItem "0,1,2,3"
Combo2.AddItem "FFh,FFh"
Combo2.ListIndex = 0
Zaman
ErrWait = ErrW * SysSpeed
StatBar = 0
PortFind
cc = Val(Right(Combo1.Text, 1)) - 1
PortAddr = PortAdress(cc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
PowerOff
If COM.PortOpen = True Then COM.PortOpen = False
End
End Sub
Sub LoadConfig()
Dim FF, n As Integer
Dim Bulundu As Boolean
FF = FreeFile
If Dir(ConfigFile, vbNormal + vbReadOnly + vbHidden + vbSystem) <> "" Then
 If FileLen(ConfigFile) <> 898 Then
  SetAttr ConfigFile, vbNormal
  Kill (ConfigFile)
  MsgBox "Lütfen programla ilgili ayarlarý giriniz."
  If Combo1.ListCount <> 0 Then
   Combo1.ListIndex = 0
   Prt = Val(Right(Combo1.Text, 1))
  End If
  SSTab1.Tab = 1
  Exit Sub
 End If
 Open ConfigFile For Random As #FF
  Get #FF, 1, Prt
  Get #FF, 2, IOType
  Get #FF, 3, speedR
  Get #FF, 4, speedW
  Get #FF, 5, speedD
  Get #FF, 6, speedP
  Get #FF, 7, speedT
  Get #FF, 8, ErrW
 Close #FF
 optDirect = IOType
 Bulundu = False
 If Combo1.ListCount <> 0 Then
  For n = 0 To Combo1.ListCount - 1
   Combo1.ListIndex = n
   If Val(Right(Combo1.Text, 1)) = Prt Then
    Bulundu = True
    Prt = Val(Right(Combo1.Text, 1))
    Exit For
   End If
  Next
  If Bulundu = False Then
   Combo1.ListIndex = 0
   Prt = Val(Right(Combo1.Text, 1))
  End If
 End If
Else
 If Combo1.ListCount <> 0 Then
  Combo1.ListIndex = 0
  Prt = Val(Right(Combo1.Text, 1))
 End If
End If
End Sub
Sub FindPorts()
Dim Tamam As Boolean
Dim n As Byte
On Error GoTo hata
If COM.PortOpen = True Then COM.PortOpen = False
For n = 0 To 8
 Tamam = True
 COM.CommPort = n
 COM.PortOpen = True
 If COM.PortOpen = True Then
  COM.PortOpen = False
 End If
 If Tamam = True Then Combo1.AddItem "COM" & n
Next
Exit Sub
hata:
Err.Clear
Tamam = False
Resume Next
End Sub

Private Sub MemWindow_DblClick(Index As Integer)
Dim r, ndx As Byte
Dim c As String

ndx = SSTAB.Tab: If ndx > 2 Then ndx = ndx - 3
If MemWindow(ndx) <> "" Then
 c = Trim(MemWindow(ndx).SelText)
 If Right(c, 1) <> ":" Then
  ExChangeHex = c
  r = Val(Left(c, 1)) * 16
  If Left(c, 1) = "A" Then r = 10 * 16
  If Left(c, 1) = "B" Then r = 11 * 16
  If Left(c, 1) = "C" Then r = 12 * 16
  If Left(c, 1) = "D" Then r = 13 * 16
  If Left(c, 1) = "E" Then r = 14 * 16
  If Left(c, 1) = "F" Then r = 15 * 16
  r = r + Val(Right(c, 1))
  If Right(c, 1) = "A" Then r = r + 10
  If Right(c, 1) = "B" Then r = r + 11
  If Right(c, 1) = "C" Then r = r + 12
  If Right(c, 1) = "D" Then r = r + 13
  If Right(c, 1) = "E" Then r = r + 14
  If Right(c, 1) = "F" Then r = r + 15
  ExChangeDec = r
 End If
End If
End Sub

Private Sub RTV_Change()
RT = RTV * 0.1
End Sub

Private Sub WTV_Change()
WT = WTV * 0.1
End Sub
Private Sub DTV_Change()
DT = DTV * 0.1
End Sub
Private Sub PTV_Change()
PT = PTV * 0.1
End Sub
Private Sub TTV_Change()
TT = TTV * 0.1
End Sub
Sub MemDisplay(ndx As Integer)
Dim j1, j2 As String
Dim n, s, m As Integer
Dim mm As Byte
 
j1 = "": j2 = "000: "
StatBar.Min = 0: StatBar.Max = 2047
For n = 0 To 2047
 StatBar = n
 mm = Memory(ndx, n)
 If mm <> 0 Then j1 = j1 + Chr(mm) Else j1 = j1 + " "
 If mm <= &HF Then
  j2 = j2 + "0" + Hex(mm)
 Else
  j2 = j2 + Hex(mm)
 End If
 s = s + 1
 If s = 16 Then
  m = n + 1
  If m <> 2048 Then
   j2 = j2 + Chr(13) + Chr(10)
   If m <= &HF Then j2 = j2 + "00" + Hex(m) + ": "
   If m > &HF And m < &HFF Then j2 = j2 + "0" + Hex(m) + ": "
   If m > &HFF Then j2 = j2 + Hex(m) + ": "
   s = 0
  End If
 Else
  j2 = j2 + " "
 End If
Next n
MemWindow(ndx) = j2
MemWindowText(ndx) = j1
StatBar = 0
End Sub
Sub PortFind()
Dim BiosAdr, GeciciPort, GeciciPort2 As Long
Dim durum As Boolean
Dim i, m, GeciciPort3 As Integer

For i = 0 To 6 Step 2
 BiosAdr = &H400 + i
 durum = GetPhysLong(BiosAdr, GeciciPort)
 GeciciPort2 = GeciciPort And 65535: GeciciPort = GeciciPort2
 If GeciciPort > 32767 Then GeciciPort3 = -32768 + (GeciciPort - 32768) Else GeciciPort3 = GeciciPort
 PortAdress(m) = GeciciPort3
 m = m + 1
Next
End Sub

